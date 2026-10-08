# SECC Trade Fair Calendar (auto-updating)

Pulls every event from https://secc.com.vn/events into a calendar that Outlook
subscribes to. GitHub runs the update for free every 3 months (1 Jan, 1 Apr,
1 Jul, 1 Oct at 08:00 Vietnam time). Outlook then picks up the changes by itself.

## How it works

1. GitHub runs `secc_calendar.py` on schedule.
2. The script reads the SECC event list and merges it into `data/events.json`.
   - New fairs are added.
   - Fairs with changed dates or halls are updated (same event, no duplicate).
   - Past fairs are kept, so your calendar keeps the history.
   - An upcoming fair that disappears from the SECC site is NOT deleted. It is
     marked "Tentative" with a note to check if it was moved or cancelled.
3. It writes `docs/secc_trade_fairs.ics` (the calendar),
   `docs/events.json` (for Power Automate) and
   `docs/secc_trade_fairs.xlsx` (same list as a spreadsheet).
4. GitHub Pages publishes the `.ics` at a fixed web address. Outlook is
   subscribed to that address and refreshes on its own.

Each event has: name, description (from SECC), dates, hours, halls, venue
address, event page link, brochure link and post show report link.

## One-time setup (about 15 minutes)

### Step 1 - Create the repo on GitHub
1. Go to https://github.com/new
2. Repository name: `secc-trade-fair-calendar`
3. Choose **Public** (needed for the free calendar web address - the data is
   public SECC event info anyway).
4. Click **Create repository**.

### Step 2 - Upload the files
1. On the new repo page click **uploading an existing file**.
2. Unzip `secc-trade-fair-calendar.zip`. Drag in everything inside the folder:
   `secc_calendar.py`, `requirements.txt`, `README.md`, `.gitignore`, and the
   `data` and `docs` folders.
3. Click **Commit changes**.

### Step 3 - Add the schedule file
(The `.github` folder is hidden on Mac, so it is easiest to create it by hand.)
1. In the repo click **Add file > Create new file**.
2. File name: `.github/workflows/update-calendar.yml`
3. Open `update-calendar.yml` from the zip (folder `.github/workflows`) in
   TextEdit, copy everything, paste it into GitHub.
4. Click **Commit changes**.

### Step 4 - Turn on GitHub Pages
1. Repo **Settings > Pages**.
2. Source: **Deploy from a branch**. Branch: **main**, folder: **/docs**. Save.
3. After 1-2 minutes your calendar address is:
   `https://YOUR-GITHUB-USERNAME.github.io/secc-trade-fair-calendar/secc_trade_fairs.ics`
   Open it in a browser - it should download a file. That means it works.

### Step 5 - Test the live update once
1. Repo **Actions** tab. If asked, click **I understand my workflows, go ahead
   and enable them**.
2. Click **Update SECC calendar > Run workflow > Run workflow**.
3. After about 1 minute you get a green tick. A red cross means it failed -
   click it to see why. (GitHub also emails you if a scheduled run fails.)

### Step 6 - Get the fairs into Outlook (choose ONE)

**Option A - Power Automate (recommended): fairs go into your MAIN calendar
as normal events you can edit.** Follow `power-automate/POWER_AUTOMATE_SETUP.md`
and skip the rest of this step.

**Option B - Subscribe: fairs show as a separate read-only calendar.**
Use Outlook on the web (it then syncs to your Mac, phone and Windows):
1. Go to https://outlook.office.com/calendar (or outlook.live.com for a
   personal account).
2. **Add calendar > Subscribe from web**.
3. Paste the address from Step 4. Name it `SECC Trade Fairs`. Click **Import**.

The fairs appear as a separate calendar you can switch on and off. Do NOT also
import the `.ics` file by hand, or you will see every fair twice.

## Running it on your own computer (optional)

```bash
cd ~/Downloads/secc-trade-fair-calendar
pip3 install -r requirements.txt
python3 secc_calendar.py
```

The updated files are saved in the `docs` folder.

## Manual update any time
**Actions > Update SECC calendar > Run workflow.**

## Why it also runs in other months
GitHub switches off schedules in repos that have no activity for 60 days. In
the months between the 3-month updates, the job only saves a one-line note
(`data/keepalive.txt`) to keep the schedule alive. It does not touch the
calendar.

## Troubleshooting
- **Red cross "permission denied" on push:** Settings > Actions > General >
  Workflow permissions > **Read and write permissions** > Save. Run again.
- **"0 events found":** SECC changed their website layout or blocked the
  request. Nothing is overwritten. Send the error to Claude with a fresh saved
  copy of the page.
- **Outlook not showing new events yet:** Outlook refreshes subscribed
  calendars on its own timing (usually within a day).
