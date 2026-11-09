# Workout Log

Workout Log is a mobile-first static web app. It can run entirely in one
browser, or use a Google Sheet through the Apps Script web app in
[`backend/Code.gs`](backend/Code.gs).

## Local/browser-only mode

Open [`index.html`](index.html) directly, or serve this directory with any
static web server. In browser-only mode, data is stored in that browser's
local storage. Use the History tab to export CSV backups.

For a local smoke test:

```sh
python3 -m http.server 8765
```

Then open <http://127.0.0.1:8765/>.

## Google Sheets backend setup

The backend is an Apps Script project bound to the Google Sheet that stores
the data. The backend creates and maintains these tabs:

- `Profiles`
- `Exercises`
- `Workouts`
- `Lifts`
- `Body_Weight`
- `Config`

### 1. Create the Sheet and Apps Script project

1. Create a Google Sheet for the workout database.
2. Open **Extensions → Apps Script**.
3. Replace the default script with the contents of
   [`backend/Code.gs`](backend/Code.gs).
4. If importing the project with clasp, also use
   [`backend/appsscript.json`](backend/appsscript.json).

### 2. Initialize the database

1. In the Apps Script editor, select `setup` as the function.
2. Run it once.
3. Grant the requested permissions.
4. Confirm that the six tabs above exist.
5. Open the `Config` tab and copy the generated `join_code`. Share this code
   only with people allowed to create profiles.

Run `setup` again only when needed. It is designed to be safe to rerun: it
repairs headers and preserves existing rows, configuration, and the signing
secret.

### 3. Deploy the Apps Script web app

1. Select **Deploy → New deployment**.
2. Choose **Web app**.
3. Set **Execute as** to the account that owns the Sheet.
4. Set **Who has access** to the audience that should use the app. The
   frontend calls the web app from a separate static origin, so the endpoint
   must be reachable by users who are allowed to sign in.
5. Deploy and copy the URL ending in `/exec`.

Do not use the development `/dev` URL in the frontend. The `/dev` URL is
restricted to editors of the Apps Script project and is not the production
endpoint.

### 4. Connect the static frontend

In [`index.html`](index.html), set `API_URL` to the deployed `/exec` URL:

```js
const API_URL = 'https://script.google.com/macros/s/DEPLOYMENT_ID/exec';
```

Do not commit a URL containing credentials or other secrets. The deployment
URL itself is not a password; profile passcodes and the Apps Script signing
secret remain server-side.

The frontend deliberately sends POST requests without a JSON content type.
This keeps the request CORS-simple for Apps Script. Do not change it to
`Content-Type: application/json` unless the endpoint's cross-origin behavior
is re-tested.

### 5. Verify the deployment

Open the deployed static frontend and verify this sequence:

1. The sign-in screen appears instead of the local demo.
2. Create a profile using the `join_code`.
3. Build and save a workout.
4. Log at least one set and body-weight value.
5. Refresh the page and confirm the values remain.
6. Open the Sheet and confirm rows were added to `Profiles`, `Workouts`,
   `Lifts`, and `Body_Weight`.
7. Sign out, create or sign in to a second profile, and confirm the first
   profile's data is not visible.
8. Temporarily disconnect the network, make a change, reconnect, and confirm
   the sync status eventually reports that it was saved.

If a deployment is updated, use **Deploy → Manage deployments → Edit** and
create a new version. Keep the same deployment URL so the static frontend does
not need another configuration change.

## Data and security notes

- The Sheet contains profile salts and passcode hashes, not plain-text
  passcodes.
- The signing secret is stored in Apps Script Script Properties, not in the
  Sheet or frontend.
- The backend validates dates, numbers, names, notes, workout sizes, and
  exercise metadata.
- Values beginning with `=`, `+`, `-`, or `@` are protected from being
  interpreted as formulas when written to the Sheet.
- Keep the Sheet and Apps Script project access restricted to trusted
  administrators.
- Make regular copies or exports of the Sheet. The app's sync design avoids
  duplicate retries, but it does not replace Google Sheet backups.

## Validation

There is currently no package manager or automated test suite. The useful
smoke checks are:

```sh
python3 - <<'PY'
from pathlib import Path
import re

html = Path("index.html").read_text()
scripts = re.findall(r"<script(?:\s[^>]*)?>(.*?)</script>", html, re.S | re.I)
Path("/tmp/workout-log-inline.js").write_text("\n".join(scripts))
PY
node --check /tmp/workout-log-inline.js
cp backend/Code.gs /tmp/workout-log-Code.js
node --check /tmp/workout-log-Code.js
```

The Apps Script backend still needs one real deployment test because
`SpreadsheetApp`, `LockService`, `CacheService`, and `PropertiesService` are
only available inside Apps Script.
