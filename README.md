# Bahdela Test Website

The existing seven-site attendance report website, with an optional **Login** button at the top right. Reports, colours, departments, exports and the existing mobile-manager reporting remain available as before.

## Employee management

The requested **Kanza**, **HR** and **admin** accounts all have permission to add, delete, transfer and enroll employees. Passwords are verified on the server against PBKDF2 hashes. Login expires after 30 minutes of inactivity. Log out on shared computers. The management panel is hidden until login.

- **Add employee:** choose “Where should this person be placed?”, then a department and a job title. Jobs are scoped to that department; choose **Add a new job title…** if needed. Enter the employee ID, full name, gender, validity dates and access plan. Confirm and click **Add employee**. An existing terminal ID is rejected. Add a face photograph or capture a fingerprint afterward if needed.
- **Transfer employee:** choose the current location and **Load / refresh report employees**. Select the employee by ID and name, choose their destination location, department and job, then **Prepare destination**. The app exports their enrolled fingerprints, face photograph and cards, creates the destination employee, copies the credentials and reads them back to verify the copy. Review the copied counts, confirm and click **Finish transfer**. This rechecks the source and destination credentials, then removes the source employee and updates both report rosters. A partial copy keeps the source active and offers **Retry credential copy**. **Pause transfer (keep both active)** retains the pending workflow while the current browser session remains available; reopen **Manage employees** to continue.
- **Delete employee:** choose the location, click **Load / refresh report employees**, then tick the people to remove in the checkbox list. Check the selected count and click **Delete selected employees**. Multiple employees can be removed in one action; unchecked employees are untouched. Each person is verified and deleted individually. Successful deletions and failures are shown separately, and failed employees remain in the list for review or retry. This deletes their terminal access and linked credentials at that location and removes the roster placement there. Other locations and historical scan events are unchanged.
- **History:** see the last 100 roster operations, with the operator, employee ID and locations.

Delete and transfer selectors show only employees assigned to the chosen site’s **report roster**, including absent staff, who match a terminal record. Unlisted terminal accounts and visiting staff assigned elsewhere are excluded. Known employee IDs and full names must match; no fuzzy first-name matching is used. The roster is checked again before deletion or transfer.

To create a department, select a site under **Add employee**, open **Add a new department at [site]**, enter its name, default job title, shift start/end and report colour, then click **Create department**. Select the new department and job before adding the employee. It is saved for that site only, appears as a normal report section and can also be selected as a transfer destination. Assigned job titles appear in the report’s existing Role column; original staff roles are retained.

Transfers automatically copy all exported card numbers (including leading zeros), fingerprint slots and fingerprint types, plus the enrolled JPEG face photograph. A card already owned by another destination employee is rejected; the app never reassigns it. Card settings and fingerprint templates are compared after copying. Face content is checked with tolerance for JPEG recompression, and the destination image is checked again before source removal. Photographs, card numbers and templates stay in memory during requests; the pending session stores only identity, locations, counts and verification digests. The roster database contains placement/history, not credential data. No physical fingerprint capture is required for an automatic transfer.

Automatic copying requires the devices' ISAPI export/import permissions and compatible firmware/fingerprint formats. This implementation targets standalone station terminals with destination fingerprint reader **1** and one face photograph per employee in face library **1**. It supports JSON and JSON/JPEG multipart face export, including device-hosted image URLs routed through the configured station address. If a terminal reports enrolled credentials but cannot export them, returns only a face model without a photograph, rejects an import or cannot confirm the saved data, the transfer stops with the source intact. It never silently skips enrolled credentials. If credentials change after preparation, use **Retry credential copy** before finishing. Test recognition on the destination terminal before confirming the final removal.

Pending transfers are kept in the current Streamlit session, not persisted across logout, a new browser session or deployment. If that session is lost after destination creation, both terminal records remain. The app rejects an already-existing destination ID rather than assuming ownership: an administrator must review the destination record and reconcile the interrupted transfer before starting another one. Historical attendance scan events are not migrated or deleted.

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

Tests use mocked terminals and temporary databases; they never delete or register real people. Coverage includes multi-card/multi-fingerprint/face transfer, multipart and URL face export, partial failure and retry, card conflicts, count mismatches, changed credentials and blocked source deletion. Live terminal behavior depends on the model, firmware and network connectivity and must be checked with a supervised test employee. The client follows [Hikvision's documented ISAPI APIs](https://open.hikvision.com/hardware/v2/08%E5%8D%8F%E8%AE%AE%E9%80%8F%E4%BC%A0/%E6%98%8E%E7%9C%B8%E9%97%A8%E7%A6%81.html) for UserInfo Search/Record/Delete, CardInfo Search/Record, FingerPrintUpload, FingerPrint/SetUp and FDLib face search/upload.
