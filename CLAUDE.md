# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A mobile-first workout log with no build step, no package manager and no test suite. Two source files do everything:

- `index.html`: the whole frontend (HTML, CSS and one inline `<script>`), served statically by GitHub Pages.
- `backend/Code.gs`: a Google Apps Script web app bound to a Google Sheet, which serves as the database. `backend/appsscript.json` is its manifest (V8 runtime).

`README.md` has the deployment steps, the intensity-technique (sub-set) feature and its research citations, and the Sheet schema. Read it before changing either of those.

## Commands

Run locally (with `API_URL` empty, the app runs in browser-only mode against localStorage):

```sh
python3 -m http.server 8765   # then open http://127.0.0.1:8765/
```

Syntax check, the only automated validation:

```sh
python3 - <<'PY'
from pathlib import Path
import re
html = Path("index.html").read_text()
scripts = re.findall(r"<script(?:\s[^>]*)?>(.*?)</script>", html, re.S | re.I)
Path("/tmp/workout-log-inline.js").write_text("\n".join(scripts))
PY
node --check /tmp/workout-log-inline.js
cp backend/Code.gs /tmp/workout-log-Code.js && node --check /tmp/workout-log-Code.js
```

`SpreadsheetApp`, `LockService`, `CacheService` and `PropertiesService` exist only inside Apps Script, so a backend change needs a real deployment test (the checklist is in README §5).

## Architecture

### Two modes, chosen by `API_URL`

`const API_URL` near the top of the script in `index.html` sets `REMOTE`. When it is empty, data lives in localStorage under `workoutlog.v5`, and first run seeds demo workouts and example history (`seed()`/`addExamples()`, marked `S.demo`; the example rows are never persisted). When it is set, the user signs in to a profile, and a per-profile copy is cached under `workoutlog.v6.*` as an offline copy. The committed `API_URL` is the live production `/exec` deployment.

### Every data change is an "op"

Frontend state is the global `S` (`lifts`, `bw`, `workouts`, `exercises`, `outbox`). The Train tab draft is `D`, and the Build tab state is `B`. All edits go through `act(...ops)`, which applies them with `applyOp`, persists them atomically with `commit` (rolled back on failure), pushes them to `S.outbox` in remote mode, and calls `sync()`.

Op types (`t`): `lifts`, `bw`, `workouts`, `exercise`, `delExercise`, `delDay`. **They are implemented twice**: `applyOp` in `index.html` and the `apply` table inside `sync()` in `Code.gs`. A new op or a semantic change has to be made in both places.

`sync()` sends at most 100 queued ops in a single POST. The server applies them and returns the profile's full data set. The client replaces `S` with that data and replays any ops queued during the request. Lift and body-weight rows are deduplicated by exact `JSON.stringify` match on both sides, so retries are idempotent. An op that fails validation comes back in `rejected` and is dropped. It does not block the rest of the queue.

### Data shapes are positional arrays

Lift rows are arrays: `LIFT_COLS = [date, workout, exercise, set, reps, weight, notes]`. A sub-set row appends `TECH_COLS = [sub, technique, rest_s, rom]` (11 values in total). A straight set stays 7 long, and the dedup check relies on this. In the Sheet, each row carries a leading `profile` column (`TABS` in `Code.gs` defines the headers). Changing a column order or length affects the dedup check, the CSV import/export (`toCsv`/`importCsv`), and `cleanLift`/`liftFromSheet` on the server. A workout's planned sub-sets are stored as JSON in the `techniques` column of `Workouts`.

### Backend specifics

- `doPost` dispatches on `req.action`, one of `signup`, `login` or `sync`. Auth uses a signed token (`tokenFor`/`whoIs`) whose secret is kept in Script Properties. Profiles store a salt and a passcode hash.
- `setup()` is safe to rerun. It creates or repairs headers and keeps existing rows, so new columns are added by extending `TABS` and running `setup`.
- All input passes through the `clean*`/`label` validators, which enforce the limits in `MAX`. Strings that start with `= + - @` are escaped (`FORMULA_START`) to prevent formula injection. The frontend applies the same check when it exports CSV.
- `MUSCLES`, `EQUIP`, `MECHANICS`, `TECHNIQUES` and `SUB` are duplicated in both files and must stay in sync. `FINE` in `Code.gs` must match the keys of `XWALK` in `index.html`: these are the exact muscles an added exercise stores (`ex.fine = [primary, secondary]`, the `Primary_muscles`/`Secondary_muscles` columns). Its muscle groups (`primary`/`secondary`) are derived from them by the form.

### Constraints

- POST requests to Apps Script are deliberately sent **without** `Content-Type: application/json`. That keeps them CORS-simple, since Apps Script cannot answer a preflight request. Do not add the header.
- Deploy the backend before the frontend: paste `Code.gs`, run `setup`, then use Manage deployments → Edit → New version to keep the same `/exec` URL. After that, publish `index.html`. An older backend rejects newer op shapes.
- `BANK` is the built-in exercise list, stored as compact tuples that index into `MUSCLES`/`EQUIP`. Custom exercises from the Sheet are merged over it by `reindex()`.
- The frontend views are tabs (`vTrain`, `vBuild`, `vHistory`, `vTrends`), each with a `draw*()` function. `refresh()` redraws them all but skips the Train tab while one of its inputs has focus.
- User-facing copy is written in plain, direct sentences. Training guide numbers (`SESSION_REF = 11`, `WEEK_REF = 10`) and their sources appear in the app under Trends → "Where the numbers come from".
