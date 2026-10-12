"""Authorized terminal operations and the small, optional management panel."""
import os
import sqlite3
from datetime import date
from pathlib import Path

import streamlit as st

from admin_auth import active_account, login, logout, require_permission
from hikvision import HikvisionClient, HikvisionError, response_message
from registration_sites import load_sites
from roster_store import RosterStore, normalized


def secrets_dict():
    try:
        return st.secrets.to_dict()
    except Exception:
        return {}


def roster_store():
    config = secrets_dict()
    path = config.get('ROSTER_DB_PATH') or os.environ.get('ROSTER_DB_PATH')
    return RosterStore(path or Path(__file__).parent / 'data' / 'employee_roster.sqlite3')


def client_for(site_name, sites):
    secret_values = secrets_dict()
    configured = load_sites(secret_values, os.environ)[site_name]
    fallback = sites[site_name]
    # Reuse the report app's existing connection when no per-site override is set.
    password = configured.password or fallback['password']
    return HikvisionClient(configured.url, configured.username, password,
                           configured.timeout, configured.verify_tls)


def verify_identity(client, employee):
    live = client.get_user(employee['employeeNo'])
    if not live or normalized(live.get('name', '')) != normalized(employee.get('name', '')):
        raise HikvisionError('This employee changed on the device. Refresh the user list and select again.')
    return live


def add_employee(state, client, store, record, site, department):
    actor = require_permission(state, 'add')
    client.create_user(record)
    store.save(actor, 'add', record['employee_no'], record['name'], destination=site, department=department)


def delete_employee(state, client, store, employee, site):
    actor = require_permission(state, 'delete')
    verify_identity(client, employee)
    client.delete_user(employee['employeeNo'])
    store.save(actor, 'delete', employee['employeeNo'], employee['name'], source=site)


def prepare_transfer(state, source_client, target_client, store, employee, source, target, department):
    actor = require_permission(state, 'transfer')
    if source == target or source_client.base_url == target_client.base_url:
        raise HikvisionError('Choose a different destination location and terminal.')
    live = verify_identity(source_client, employee)
    # Preserve access rights and validity, but never copy passwords, cards or biometrics.
    if target_client.get_user(live['employeeNo']):
        raise HikvisionError('That employee ID already exists at the destination. No records were changed.')
    allowed = ('employeeNo', 'name', 'userType', 'gender', 'Valid', 'doorRight', 'RightPlan')
    user = {key: live[key] for key in allowed if key in live}
    result = target_client._request('PUT', '/ISAPI/AccessControl/UserInfo/Record?format=json', json={'UserInfo': user})
    if not result.ok:
        raise HikvisionError('Destination employee was not created: ' + response_message(result))
    verify_identity(target_client, employee)
    store.save(actor, 'prepare transfer', employee['employeeNo'], employee['name'], destination=target, department=department)
    return {'employee': dict(employee), 'source': source, 'target': target, 'department': department}


def finish_transfer(state, source_client, target_client, store, pending):
    actor = require_permission(state, 'transfer')
    employee = pending['employee']
    if source_client.base_url == target_client.base_url:
        raise HikvisionError('Source and destination resolve to the same terminal. Check site configuration.')
    verify_identity(target_client, employee)
    verify_identity(source_client, employee)
    source_client.delete_user(employee['employeeNo'])
    store.save(actor, 'transfer', employee['employeeNo'], employee['name'],
               source=pending['source'], destination=pending['target'], department=pending['department'])


def render_login():
    account = active_account(st.session_state)
    if not account:
        with st.popover('Login', use_container_width=True):
            with st.form('login_form'):
                username = st.text_input('Username', key='login_username')
                password = st.text_input('Password', type='password', key='login_password')
                submitted = st.form_submit_button('Log in', use_container_width=True)
            if submitted:
                ok, message = login(st.session_state, username, password,
                                    secrets_dict().get('admin_accounts', {}))
                if ok:
                    st.session_state.pop('login_password', None)
                    st.session_state['management_open'] = True
                    st.rerun()
                st.error(message)
    else:
        with st.popover(account, use_container_width=True):
            st.caption('Employee management')
            if st.button('Manage employees', use_container_width=True):
                st.session_state['management_open'] = not st.session_state.get('management_open', False)
                st.rerun()
            if st.button('Log out', use_container_width=True):
                logout(st.session_state)
                st.rerun()


def department_choice(site, sites, key):
    departments = [name for name, dept in sites[site]['employees'].items() if not dept.get('visitor_only')]
    return st.selectbox('Department *', departments, index=None, placeholder='Where should this person work?', key=key)


def selected_employee(site, sites, key):
    if st.button('Load / refresh employees', key=key + '_refresh'):
        require_permission(st.session_state, 'list')
        st.session_state[key + '_users'] = client_for(site, sites).search_users()
    users = st.session_state.get(key + '_users', [])
    if not users:
        st.info('Load employees from the selected terminal first.')
        return None
    choice = st.selectbox('Employee', range(len(users)), index=None,
        format_func=lambda i: f"{users[i].get('employeeNo', '')} · {users[i].get('name', '')}", key=key + '_choice')
    return users[choice] if choice is not None else None


def enroll_employee(employee, site, sites, key):
    st.caption(f"Biometrics: {employee['name']} · {employee['employeeNo']} · {site}")
    photo = st.file_uploader('Face photograph (JPEG)', type=['jpg', 'jpeg'], key=key + '_photo')
    if st.button('Save face at this location', disabled=photo is None, key=key + '_face'):
        require_permission(st.session_state, 'enroll')
        client = client_for(site, sites)
        verify_identity(client, employee)
        result = client.upload_face(employee['employeeNo'], photo.getvalue(), photo.name)
        if not result.ok:
            raise HikvisionError('Face was not saved: ' + response_message(result))
        st.success('Face saved at ' + site)
    st.caption('For fingerprint capture, the employee must stand beside this location’s terminal.')
    if st.button('Capture and save fingerprint', key=key + '_finger'):
        require_permission(st.session_state, 'enroll')
        client = client_for(site, sites)
        verify_identity(client, employee)
        captured = client.capture_fingerprint(1)
        if not captured.data:
            raise HikvisionError(captured.message or 'No fingerprint returned by the terminal.')
        result = client.apply_fingerprint(employee['employeeNo'], captured.data, 1)
        if not result.ok:
            raise HikvisionError('Fingerprint was not saved: ' + response_message(result))
        st.success('Fingerprint saved at ' + site)


def render_management(sites, store):
    if not active_account(st.session_state) or not st.session_state.get('management_open'):
        return
    with st.expander('Employee management', expanded=True):
        st.caption('Add, delete or transfer employees on the selected Hikvision terminal and in the report roster.')
        try:
            _render_management(sites, store)
        except (HikvisionError, PermissionError) as exc:
            st.error(str(exc))
        except (OSError, sqlite3.Error):
            st.error('Could not save the roster. The terminal may already have changed. Refresh the device list and check the data directory before continuing.')


def _render_management(sites, store):
    add_tab, transfer_tab, delete_tab, history_tab = st.tabs(['Add employee', 'Transfer employee', 'Delete employee', 'History'])
    names = list(sites)
    with add_tab:
        site = st.selectbox('Where should this person be placed? *', names, index=None, key='manage_add_site')
        if site:
            department = department_choice(site, sites, 'manage_add_department_' + site)
            with st.form('manage_add_form_' + site):
                employee_id = st.text_input('Employee ID *', max_chars=32)
                full_name = st.text_input('Full name *', max_chars=128)
                gender = st.selectbox('Gender', ['unspecified', 'male', 'female'])
                c1, c2 = st.columns(2)
                start = c1.date_input('Valid from', date.today())
                end = c2.date_input('Valid until', date(2036, 12, 31))
                plan = st.number_input('Access plan template', min_value=1, max_value=255, value=1)
                confirm = st.checkbox('I confirm this person should be added to this location and department.')
                submitted = st.form_submit_button('Add employee')
            if submitted:
                if not department or not full_name.strip() or not employee_id.strip().isalnum() or end < start or not confirm:
                    st.error('Enter a valid ID and name, choose a department, check the dates and confirm the placement.')
                else:
                    record = {'employee_no': employee_id.strip(), 'name': ' '.join(full_name.split()),
                              'gender': gender, 'valid_from': start.isoformat(), 'valid_until': end.isoformat(), 'access_plan': int(plan)}
                    add_employee(st.session_state, client_for(site, sites), store, record, site, department)
                    st.session_state['management_enroll'] = {'site': site, 'employee': {'employeeNo': record['employee_no'], 'name': record['name']}}
                    st.success('Employee added to ' + site + ' / ' + department)
            saved = st.session_state.get('management_enroll')
            if saved and saved['site'] == site:
                enroll_employee(saved['employee'], site, sites, 'manage_add_enroll_' + site)
    with transfer_tab:
        pending = st.session_state.get('management_transfer')
        if pending:
            st.info(f"{pending['employee']['name']}: destination record saved at {pending['target']}. Source remains active at {pending['source']} until you finish.")
            st.warning('Face, fingerprint and card credentials are not copied automatically. Enroll at the destination or arrange enrollment before removing the source record.')
            enroll_employee(pending['employee'], pending['target'], sites, 'manage_transfer_enroll')
            confirmed = st.checkbox('The destination is ready, or enrollment has been arranged. Remove this employee and their credentials from the source.', key='manage_transfer_confirm')
            if st.button('Finish transfer', disabled=not confirmed):
                finish_transfer(st.session_state, client_for(pending['source'], sites), client_for(pending['target'], sites), store, pending)
                st.session_state.pop('management_transfer', None)
                st.session_state['management_message'] = 'Transfer completed. Employee removed from the source and placed at the destination.'
                st.rerun()
            if st.button('Leave both locations active for now'):
                st.session_state.pop('management_transfer', None)
                st.session_state['management_message'] = 'Both records remain active. Use Delete employee at the source when ready.'
                st.rerun()
        else:
            source = st.selectbox('Current location *', names, index=None, key='manage_transfer_source')
            if source:
                employee = selected_employee(source, sites, 'manage_transfer_' + source)
                target = st.selectbox('Where should this person be transferred? *', [n for n in names if n != source], index=None, key='manage_transfer_target_' + source)
                department = department_choice(target, sites, 'manage_transfer_department_' + target) if target else None
                confirmed = st.checkbox('I confirm the selected employee and destination.', key='manage_transfer_start_confirm_' + source)
                if st.button('Prepare destination', disabled=not (employee and target and department and confirmed)):
                    st.session_state['management_transfer'] = prepare_transfer(st.session_state, client_for(source, sites), client_for(target, sites), store, employee, source, target, department)
                    st.rerun()
    with delete_tab:
        site = st.selectbox('Delete from which location? *', names, index=None, key='manage_delete_site')
        if site:
            employee = selected_employee(site, sites, 'manage_delete_' + site)
            if employee:
                st.warning(f"Delete {employee['name']} (ID {employee['employeeNo']}) from {site}? This removes their terminal access, linked credentials and report roster placement. Other locations are unchanged.")
                typed = st.text_input('Type the employee ID to confirm deletion', key='manage_delete_confirm_' + site)
                if st.button('Delete employee', disabled=typed.strip() != str(employee['employeeNo'])):
                    delete_employee(st.session_state, client_for(site, sites), store, employee, site)
                    st.session_state.pop('manage_delete_' + site + '_users', None)
                    st.session_state['management_message'] = 'Employee deleted from ' + site
                    st.rerun()
    with history_tab:
        st.dataframe(store.history(), use_container_width=True, hide_index=True)
