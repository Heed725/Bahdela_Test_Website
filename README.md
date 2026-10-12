# Bahdela Test Website

The existing seven-site attendance report website, with an optional **Login** button at the top right. Reports, colours, departments, exports and the existing mobile-manager reporting remain available as before.

## Employee management

The requested **Kanza**, **HR** and **admin** accounts all have permission to add, delete, transfer and enroll employees. Passwords are verified on the server against PBKDF2 hashes. Login expires after 30 minutes of inactivity. Log out on shared computers. The management panel is hidden until login.

- **Add employee:** choose “Where should this person be placed?”, then a department. Enter the employee ID, full name, gender, validity dates and access plan. Confirm and click **Add employee**. An existing terminal ID is rejected. Add a face photograph or capture a fingerprint afterward if needed.
- **Transfer employee:** choose the current location and **Load / refresh employees**. Select the employee by ID and name, choose their destination location and department, then **Prepare destination**. The destination record is created and checked first; the source stays active. Enroll the person at the destination, then confirm and **Finish transfer**. Only this final step removes the source record and its credentials. If interrupted, both records remain active; load the source in **Delete employee** to complete removal once the destination is ready.
- **Delete employee:** choose the location, load employees and select the exact person. Type their employee ID and click **Delete employee**. This deletes their terminal access and linked credentials at that location and removes the roster placement there. Other locations and historical scan events are unchanged.
- **History:** see the last 100 roster operations, with the operator, employee ID and locations.

Face images, fingerprint templates and cards are **not copied during transfer**. A JPEG may be uploaded at the destination and fingerprint slot 1 may be captured there, matching the registration app. The employee must stand beside the destination terminal for fingerprint capture. Biometric files are sent to the terminal, not written to the roster database. Choose “Leave both locations active for now” if enrollment must happen later.

Device failures stop the workflow and are displayed. A rejected destination or duplicate ID never triggers source deletion. Removal is confirmed by checking the terminal again. The app only sends a single explicitly selected employee ID to the delete endpoint; it never requests deletion of all users.

## Locations and connections

Locations match [Bahdela Hikvision Registration](https://github.com/Heed725/Bahdela_Hikvision_Registration_Streamlit): Buguruni, Puma Upanga, Puma Ocean Road, Puma Survey, India, Kibaha and Livingstone. The registration repository itself is unchanged.

Management uses the same default terminal addresses as that app and can use the same `[sites."Location"]` Streamlit Secrets configuration: `url`, `username`, `password`, `timeout`, `verify_tls`. Environment overrides use names such as `SITE_PUMA_UPANGA_PASSWORD`. When no device password override is configured, the report website's existing connection credentials are reused. Logged-in management does not need the public registration enrollment PIN.

## Saved roster and hosting

Roster additions, removals, transfers and history are saved in SQLite at `data/employee_roster.sqlite3`. They survive page refreshes and process restarts **when this directory is retained**. They are applied after the original roster initialization on every run, so the Python roster lists do not need manual editing.

For durable hosting, set `ROSTER_DB_PATH` in Streamlit Secrets or the environment to a SQLite file on a **persistent server volume**, for example `/var/lib/bahdela/employee_roster.sqlite3`, and back up that file. The process must have write permission. On multiple replicas, use one shared supported persistent database deployment rather than independent local files.

**Streamlit Community Cloud's local disk is not durable across replacement/redeployment.** The default local database is therefore insufficient for permanent cloud roster storage. Deploy on a server with a persistent volume before relying on saved roster changes. Device records themselves are stored on the Hikvision terminals. This update does not provision hosting storage or change your deployment settings.

Optional `[admin_accounts]` secrets may override each account's password hash. Generate a hash without putting the password into a command argument:

```bash
python -c 'import getpass,secrets; from admin_auth import password_hash; print(password_hash(getpass.getpass("New password: "),secrets.token_hex(16)))'
```

Place the returned value under the account name in `[admin_accounts]`. Do not commit real secrets, biometric files or the database.

## Run and verify

```bash
pip install -r requirements.txt
streamlit run app.py
python -m unittest discover -s tests -v
```

Tests use mocked terminals and temporary databases; they never delete or register real people. Live terminal behavior depends on the model, firmware and network connectivity and must be checked with a supervised test employee. The new client follows Hikvision's UserInfo Search/Record and single-person UserInfoDetail Delete APIs.
