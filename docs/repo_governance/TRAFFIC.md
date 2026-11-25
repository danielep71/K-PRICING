# 📈 Repository traffic history

GitHub keeps repository traffic for 14 days. The optional **Traffic history**
workflow ([traffic-history.yml](../../.github/workflows/traffic-history.yml))
collects it once a day, with the repository's stars, forks, watchers and open
issues, and merges it into CSV files on the data-only `traffic-history`
branch, so the history outlives that window. It also writes badge files there
and opens an issue when traffic stands out (*Traffic alert* below). It is
repository analytics, not CI: it builds and tests nothing.

## 🔌 Enabling

The workflow does nothing until the repository variable `TRAFFIC_HISTORY` is
`enabled`; until then its scheduled runs are skipped. Enable it in this order:

1. **Create the branch** from a clone of the default branch. The commands
   work alike in PowerShell and `sh`, and need Git 2.42 or later:

   ```sh
   git worktree add --orphan -b traffic-history ../traffic-history
   python tools/traffic_history.py init --data ../traffic-history/data
   git -C ../traffic-history add data
   git -C ../traffic-history commit -m "chore: start the traffic history"
   git -C ../traffic-history push origin traffic-history
   git worktree remove ../traffic-history
   ```

   `git worktree add --orphan` starts the branch in a separate folder beside
   the clone, with no files and no history in common with the default branch;
   the clone stays on its branch, and the last command removes the folder.
   The branch is never merged and never used as a task base. The setup
   baseline's **Protect traffic history** ruleset blocks its deletion and
   force pushes.
2. **Settings → Environments → New environment** `analytics`. Under
   *Deployment branches and tags*, select *Selected branches and tags* and
   allow the default branch only. Do this before adding the App's credentials:
   a run that names a missing environment creates it without that restriction.
3. **Create the App:** under **Settings → Developer settings → GitHub Apps**
   of the account that owns the repository, create a GitHub App for traffic
   collection only. Disable its webhook, grant only the repository permission
   **Administration: Read-only** (GitHub adds *Metadata: Read-only*) and allow
   installation only on this account. The traffic API needs that permission,
   and the workflow token cannot have it. If you already have such an App for
   another repository, reuse it and skip this step.
4. **Install the App** on this repository: on a new installation select only
   this repository; on an existing one add this repository to it. Each run's
   token is limited to the repository it runs in.
5. In the `analytics` environment, add the App's **Client ID** as the
   environment variable **TRAFFIC_APP_CLIENT_ID**. Generate a private key for
   this repository (an App can hold several, so one leaked key can be revoked
   without affecting other repositories) and add the whole `.pem` content as
   the environment secret **TRAFFIC_APP_PRIVATE_KEY**, then delete the
   downloaded file. Never put the key in a file, issue, comment or chat.
6. **Issues → Milestones**: keep an open milestone with a due date after
   the collection day. The *Traffic alert* job files each alert issue under
   one, and fails when an alert is due and none qualifies; a run with no alert
   due needs no milestone.
7. **Settings → Secrets and variables → Actions → Variables**: add the
   repository variable `TRAFFIC_HISTORY` with the value `enabled`.
8. Run **Actions → Traffic history → Run workflow** on the default branch and
   check that it passes and that `traffic-history` gains a commit.

To pause collection, delete the variable or change its value; the history
stays on the branch.

## ⏱️ Schedule

The workflow is scheduled daily at **06:17 UTC** and can be started manually
on the default branch; a manual run on any other branch is skipped. GitHub may
start scheduled runs late; each collection records the time it was actually
taken.

## 🔑 Credentials

Each run asks the App for an installation token limited to this repository
and to *Administration: read*. The token expires after an hour and is revoked
when the job ends, so there is nothing to renew. Only the token step receives
the private key, and only the fetch step, which runs `curl` and no repository
code, receives the token. The merge script runs without either; the push uses
the workflow token with `contents: write`.

### Rotating the private key

Rotate the key if it may have been exposed, or if anyone but me changes the
workflow:

1. Generate a new private key for this repository in the App's settings.
2. Replace `TRAFFIC_APP_PRIVATE_KEY` in **Settings → Environments →
   `analytics`**. A repository secret with the same name does not take
   effect: the environment secret takes precedence.
3. Run the workflow manually on the default branch and check that **Fetch
   traffic from the GitHub API** passes and `traffic-history` gains a commit.
4. Only then delete the old key in the App's settings, checking its
   fingerprint first.

Collections missed while the App cannot issue tokens are lost once they fall
outside the 14-day window.

## 🗄️ Stored data

The `traffic-history` branch holds these files under `data/`:

| File | Key | Contents |
| --- | --- | --- |
| `daily.csv` | `date` | Views and unique views, clones and unique clones per UTC day. For each pair, `*_collected_at` gives the collection that last changed it and `*_observed_at` the newest collection that reported it. |
| `snapshots.csv` | `snapshot_date` | One row per collection day: the first and last day the API reported, the 14-day totals, the repository's `stars`, `forks`, `watchers` and `open_issues` at collection time, `collected_at`, the collection that last changed that day's snapshot, referrers or paths, and `observed_at`, the newest collection of that day. |
| `referrers.csv` | `snapshot_date`, `rank` | The top 10 referrers over the 14 days before each collection. |
| `paths.csv` | `snapshot_date`, `rank` | The top 10 repository paths over the same window. |
| `badges/*.json` | — | Badge files, rewritten from the stored data after each merge (*Badges* below). |

Referrers and paths are kept per collection day, because the API reports them
only as 14-day totals. `open_issues` is GitHub's count, which includes open
pull requests.

### Badges

`data/badges/` holds three files in the shields.io *endpoint* format:

| File | Label | Counts |
| --- | --- | --- |
| `clones.json` | `clones` | Every clone in `daily.csv`, summed since tracking began. |
| `visitors.json` | `visitors (14d)` | Unique visitors over the API's 14-day window, from the latest snapshot. Daily unique counts cannot be summed into a unique count. |
| `visits.json` | `visits` | Daily unique visitors summed since tracking began: a visitor counts once per day they visit. |

On a public repository, a README can show one with
`https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/OWNER/REPOSITORY/traffic-history/data/badges/clones.json`.
Large numbers are shortened (`1.2k`, `11k`). A badge file changes only when its
value does. Clones include those made by workflow checkouts, so CI activity
raises the count; for the same reason they never raise a
[traffic alert](#-traffic-alert).

The README shows the three badges for this repository. **Repository integrity**
fails while a README badge reads another repository's `traffic-history` branch:
point the badges at this repository, or remove them if you do not enable
collection.

### Merge rules

`tools/traffic_history.py` merges each collection idempotently: merging the
same collection again changes nothing.

- **Overlapping windows:** for every date the collection reports, its values
  replace the stored ones, so a day that was incomplete when first collected
  is corrected later. Views and clones are merged separately, each with its
  own collection and observation times.
- **Older history:** dates the collection does not report keep their values.
  Nothing is deleted.
- **No double counting:** values are replaced, never added. Unique counts are
  never summed across days or windows; a 14-day unique count exists only in
  `snapshots.csv`.
- **Same-day runs:** a collection day's snapshot, referrers and paths are one
  unit, replaced only if any of the three changed. A newer collection with the
  same data changes no value or collection time; it records only its
  observation time, so each successful run adds a commit to the branch.
- **Order:** the newest collection that reported a value wins, even when it
  left the value unchanged or reported a lower count. An older collection
  merged after a newer one, for example when a saved collection is replayed,
  never replaces what the newer one reported; it can only fill a date or
  metric that nothing newer reported.
- **Earlier layouts:** a history file written before the observation times
  existed is read with each observation time equal to its collection time and
  rewritten with the current header on the next merge.

### Zero or unavailable

The API reports each day of its window, days without activity included, as
`0`. So:

- **`0`**: the day was collected and had no activity;
- **no row**: no collection covered the day, because it is before tracking
  began or because collections stopped for more than 14 days;
- **a blank metric**: one of the two series was not reported for that day.

The step summary of each run lists the dates between the first and the last
stored day that have no row.

### Visibility

The branch has the repository's visibility: on a public repository anyone can
read its CSV files. The data is aggregate (counts, referrers and paths) and
identifies no visitor. The counts include my own visits and every clone made
by CI checkouts.

GitHub shows top referrers and popular paths only to repository
administrators. I publish them on purpose, public repositories included,
because they are aggregate and show which sites and pages draw visitors. A
repository that prefers not to publish them can stop collecting them: in the
workflow's *Fetch traffic from the GitHub API* step, remove
`traffic/popular/referrers` and `traffic/popular/paths` from the request loop
and write an empty list instead, for example `echo '[]' > "$RAW/referrers.json"`
and the same for `paths.json`. The merge accepts empty lists and records no
rankings for new collection days, so the *New referrer* alert can no longer
fire. The opt-out stops future collection only: rows from earlier days stay in
`referrers.csv` and `paths.csv`, because the merge never removes other dates.
To clear them from the current files, commit both files with only their header
lines to `traffic-history`; earlier commits keep them in the branch history,
which the **Protect traffic history** ruleset preserves.

## 🚨 Traffic alert

After each collection the **Traffic alert** job compares the latest snapshot
with the one before (`traffic_history.py alert`) and opens one issue, titled
`Traffic alert: <snapshot date>`, when any of these holds:

- **Views spike:** the 14-day view total at least doubles and grows by at
  least 100 views;
- **Stars or forks increase;**
- **New referrer:** a referrer never stored before brings at least 5 views.

Clones never raise an alert. GitHub counts every workflow checkout as a clone,
so a day of pull requests raises the clone count far more than outside
interest does, and a clone alert would fire on the repository's own CI. Clones
are still recorded in `daily.csv` and `snapshots.csv`, shown in the alert
issue's table and summed in the clones badge.

At most one issue is opened per snapshot date: the job skips a date whose
issue exists, open or closed. The issue is assigned to the issue-form account,
labelled `P3` and `repository`, filed under the open milestone with the
earliest due date after the collection day, and carries a dependency section,
a start date (the snapshot date) and a target end date (seven days after the snapshot, or the day before the
milestone's due date if that is earlier), so it passes **Issue triage**. The
job first decides from the history alone whether an alert is due
(`traffic_history.py alert --snapshot-output`), and looks up the milestone
only then. It compares the due dates with the snapshot's UTC collection date,
not with the date the job runs. When an alert is due, the job needs at least
one open milestone due after the collection day.

The job uses the workflow token with `issues: write` and `contents: read`. It
has no environment, so it never receives the App's key or token, and it runs
only after a successful collection. It reads the `traffic-history` commit that
its collection pushed, or merged onto when nothing changed, not the branch's
current tip, so a re-run of the job reads the same data even after the branch
has moved.

## 🩺 Failure diagnosis and recovery

Each run fails before anything is written if a response is missing or
invalid, and leaves the history unchanged. Read the failing step:

| Step and message | Cause | Action |
| --- | --- | --- |
| *Check the traffic App configuration*: `TRAFFIC_APP_CLIENT_ID or TRAFFIC_APP_PRIVATE_KEY is not set` | Variable or secret missing, or the run is not in the `analytics` environment. | Follow *Enabling*, steps 2–5. |
| *Create the traffic App token* fails | Wrong client ID or key, the key was deleted, or the App is not installed on this repository. | Check the App's installation, then rotate the key as above. |
| *Fetch*: HTTP 401 | The installation token was rejected, for example because the App was uninstalled during the run. | Check the App's installation and run the workflow again. |
| *Fetch*: HTTP 403 or 404 | The App lacks *Administration: read*, or is not installed on this repository. | Grant the permission (the account must accept the change) or install the App, then run again. |
| *Fetch*: HTTP 429 or 5xx after retries | Rate limit or GitHub outage. | Run it again later; the next run collects the same 14 days. |
| *Merge*: `ERROR: ...; the history was not changed` | The API returned an unexpected shape, or a history file no longer has the expected layout. | Inspect the response or the named file; a file edited by hand must keep its header and columns. |
| *Commit and push*: `Could not push the history` | `traffic-history` kept moving, or a rule rejected the push. | Check the branch's rules and recent commits, then run again. |
| *Check out traffic-history* fails | The branch does not exist or was deleted. | Create it as in *Enabling*, or restore it from the last known commit; never recreate it empty over existing history. |
| *Traffic alert*: `No open milestone due after the collection date` | An alert is due, and every open milestone is due on or before the collection day or has no due date, or none is open. | Open the next milestone with a later due date, or move the current one's due date, then run the workflow again. |
| *Traffic alert*: `gh issue create` fails | A label, the assignee or the milestone no longer exists. | Restore it, then run the workflow again; the history is already saved. |

A failed run can always be repeated: the merge is idempotent, and each
collection covers the last 14 days. To repair the branch, change it with a new
commit made with `python tools/traffic_history.py merge` or by hand, keeping
every header. `python tools/traffic_history.py init --data data` recreates
missing files with their headers and never overwrites an existing one.
