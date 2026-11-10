# Workout Log

Workout Log is a mobile-first static web app. It can run entirely in one
browser, or use a Google Sheet through the Apps Script web app in
[`backend/Code.gs`](backend/Code.gs).

## Intensity techniques (sub-sets)

Any set can carry sub-sets that extend it past the point where the straight
set ended:

| Technique | What you log | How it is done |
| --- | --- | --- |
| Partial reps | reps, share of the range (¼, ½, ¾) | At failure, keep going through part of the range, in the stretched part of the lift. Same weight. |
| Myo reps | rest in seconds, weight, reps | Rest about 20 seconds, then a mini-set of a few reps; repeat. Weight left blank means the same weight. |
| Rest-pause | extra reps | Rack the weight for a few breaths, then add the reps you can. Same weight. |
| Drop set | weight, reps | Cut the weight by about a fifth and go again straight away. |

- **Chain:** tap **+** at the end of a set row again to add another sub-set
  after the last one (drop, then drop, then partials).
- **Nest:** tap **+** on a sub-set row to put a sub-set under it, one step
  further in. It takes its weight from the row it sits under, so partials
  nested under a drop are logged at the dropped weight. Up to 3 levels, 9 per
  level.
- **Build tab:** "Add a technique to the last set" plans sub-sets for an
  exercise. They appear on its last set when you train that workout.
- **Train tab:** add or remove sub-sets on any set on the day. A planned
  sub-set left blank is skipped and not saved.

### How they count

- Sub-sets are tallied on their own, in grey, with the same weighting as sets
  (1 for a primary muscle, ½ for a secondary one). They are **not** added to
  the set totals or compared with the 10-per-week and 11-per-session guide
  lines, because no study gives a conversion from a sub-set to a straight set.
- Their weight × reps does count toward **Load**. Partial reps are scaled by
  their share of the range.
- Bests, targets, the grey "last session" numbers and the **Strength** index
  use straight sets only.

### What the research says

- **Drop sets:** about the same muscle growth as straight sets, in a half to a
  third of the training time. Sødal et al. 2023, meta-analysis of 5 studies,
  *Sports Med Open* 9:66, <https://doi.org/10.1186/s40798-023-00620-5>.
- **Rest-pause:** matched straight sets for growth when total volume was
  equal, in 28 trained men over 8 weeks (Enes et al. 2021, *Appl Physiol Nutr
  Metab*, <https://pubmed.ncbi.nlm.nih.gov/34260860/>). An earlier 6-week
  study of 18 trained people found more thigh growth with rest-pause and no
  difference elsewhere (Prestes et al. 2019, *J Strength Cond Res* 33:S113,
  <https://pubmed.ncbi.nlm.nih.gov/28617715/>).
- **Myo reps:** a rest-pause variant (one set near failure, then mini-sets on
  short rests) from coaching practice. No trial of myo reps by name is cited
  here; the rest-pause studies above are the closest evidence.
- **Partial reps past failure:** calves grew 9.6% against 6.7% in 10 weeks
  when stretched-position partials were added after full-rep failure. One
  within-person study of 23 untrained men (Larsen et al. 2025, *Front
  Psychol* 16:1494323, <https://doi.org/10.3389/fpsyg.2025.1494323>).
- **Which part of the range:** full range is a sound default, and partials do
  best at long muscle lengths; the differences are small (Wolf et al. 2023,
  *Int J Strength Cond* 3(1), <https://doi.org/10.47206/ijsc.v3i1.182>).
- **Overview:** a systematic review of advanced techniques found too little
  evidence to give volume guidelines for them (Krzysztofik et al. 2019, *Int J
  Environ Res Public Health* 16:4897,
  <https://doi.org/10.3390/ijerph16244897>).

The study figures above come from the papers' abstracts. Sources for the rest
of the app's numbers are listed in the app under **Trends → Where the numbers
come from**.

### How they are stored

A sub-set is one more row in `Lifts` with the same `set` number as the set it
belongs to, and four extra columns filled in:

| Column | Meaning |
| --- | --- |
| `sub` | Place under the set: `2` is the second sub-set, `2.1` the first one nested under it. |
| `technique` | `partial`, `myo`, `rest-pause` or `drop`. Blank on a straight set. |
| `rest_s` | Myo-rep rest in seconds, if entered. |
| `rom` | Share of the range for partial reps: 0.25, 0.5 or 0.75. |

`weight_lb` always holds the weight actually used, including on rows that
took it from the row above. In pandas, `df[df.technique.isna()]` gives the
straight sets. A workout's planned sub-sets are kept as JSON in the
`techniques` column of `Workouts`. CSV files exported before these columns
existed still import.

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

### Updating a deployment for intensity techniques

Do this before publishing the new `index.html`. An older backend still saves
straight sets, but it rejects sub-set rows and drops the planned sub-sets of
saved workouts.

1. Paste the new [`backend/Code.gs`](backend/Code.gs) into the Apps Script
   editor.
2. Run `setup` once. It adds the `sub`, `technique`, `rest_s` and `rom`
   headers to `Lifts` and `techniques` to `Workouts`. Existing rows are kept.
3. **Deploy → Manage deployments → Edit**, choose **New version**, deploy.
4. Publish the new `index.html`.

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
