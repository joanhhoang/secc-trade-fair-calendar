# Power Automate flow - SECC fairs into your main Outlook calendar

**What it does:** every 3 months the flow downloads the SECC event list that
GitHub publishes. Then, for each fair:
- **New fair:** creates the event in your main Outlook calendar and saves its
  Outlook ID in the Excel tracker.
- **Changed fair** (dates, halls or status changed): updates the same Outlook
  event, so you don't get a duplicate.
- **Same as last time:** skips it.

**Before you start:** do GitHub setup steps 1-5 in README.md first, and skip
step 6 (Outlook subscribe). If you do both, every fair shows twice.
Check that this address opens in your browser and shows text:
`https://YOUR-GITHUB-USERNAME.github.io/secc-trade-fair-calendar/events.json`

---

## Part A - Put the tracker in OneDrive (2 minutes)

1. Open OneDrive (https://onedrive.live.com or your Microsoft 365 OneDrive).
2. Upload `SECC Trade Fairs Tracker.xlsx` to the main OneDrive folder.
3. Don't rename the table inside it. It's called `SECC` and the flow looks for it.

---

## Part B - Build the flow (about 20 minutes)

Go to https://make.powerautomate.com > **Create > Scheduled cloud flow**.

**IMPORTANT:** rename each step exactly as written below (click the step title,
or "..." > Rename). The expressions refer to these names. If a name is
different, the expression won't work.

### Trigger - Recurrence
- Flow name: `SECC trade fairs to Outlook`
- Starting: **1 Jan 2027, 12:00 PM**
- Repeat every: **3 Month**
- (Click the trigger after creating it.) Time zone: **(UTC+07:00) Bangkok, Hanoi, Jakarta**

GitHub updates the list at 08:00 on the same day, so the flow runs 4 hours later.

### Step 1 - HTTP  -> rename to `Get events`
- Method: `GET`
- URI: `https://YOUR-GITHUB-USERNAME.github.io/secc-trade-fair-calendar/events.json`

### Step 2 - Parse JSON (Data Operation)  -> rename to `Parse`
- Content: click in the box, then choose **Body** from `Get events`
- Schema: open `parse_json_schema.json` in TextEdit, copy everything, paste

### Step 3 - Apply to each  -> rename to `For each fair`
- Select an output: click **fx** (expression) and paste:
  ```
  body('Parse')?['events']
  ```
- Leave concurrency **off** (the default), because Excel doesn't like many
  writes at the same time.

All the steps below go **inside** `For each fair`.

### Step 3.1 - Excel Online (Business): List rows present in a table  -> rename to `Find row`
- Location: **OneDrive for Business**
- Document Library: **OneDrive**
- File: `/SECC Trade Fairs Tracker.xlsx`
- Table: `SECC`
- Show advanced options > **Filter Query**: type `Key eq '` then add this
  expression with **fx**, then type `'` at the end:
  ```
  items('For_each_fair')?['key']
  ```
  The result should look like: `Key eq '[items(...)]'`

### Step 3.2 - Condition  -> rename to `Is it new`
- Left box (fx): `length(outputs('Find_row')?['body/value'])`
- Middle: **is equal to**
- Right box: `0`

### If yes (new fair)

**3.2a - Office 365 Outlook: Create event (V4)**  -> rename to `Create event`

| Field | What to put (use fx for expressions) |
|---|---|
| Calendar id | `Calendar` |
| Subject | `items('For_each_fair')?['name']` |
| Start time | `items('For_each_fair')?['start_time']` |
| End time | `items('For_each_fair')?['end_time']` |
| Time zone | `(UTC+07:00) Bangkok, Hanoi, Jakarta` |
| Body | `items('For_each_fair')?['body_html']` |
| Location | `items('For_each_fair')?['location']` |
| Is all day event | `Yes` |
| Show as | `Free` |
| Is reminder on | `No` (or Yes, if you want reminders) |

(The "advanced options" link shows the fields that are hidden at first.)

**3.2b - Excel Online (Business): Add a row into a table**  -> rename to `Add row`
- Same Location / Library / File / Table as `Find row`
- Key: `items('For_each_fair')?['key']`
- Name: `items('For_each_fair')?['name']`
- Start: `items('For_each_fair')?['start']`
- End: `items('For_each_fair')?['end']`
- Status: `items('For_each_fair')?['status']`
- Version: `items('For_each_fair')?['version']`
- Outlook_ID: `body('Create_event')?['id']`
- Last_Updated: `utcNow()`

### If no (fair already in your calendar)

**3.2c - Condition**  -> rename to `Has it changed`
- Left (fx): `items('For_each_fair')?['version']`
- Middle: **is not equal to**
- Right (fx): `int(first(outputs('Find_row')?['body/value'])?['Version'])`

**If yes** inside `Has it changed`:

**Office 365 Outlook: Update event (V4)**  -> rename to `Update event`
- Calendar id: `Calendar`
- Id (fx): `first(outputs('Find_row')?['body/value'])?['Outlook_ID']`
- Fill in Subject, Start time, End time, Time zone, Body, Location, Is all day
  event, Show as exactly like `Create event`.

**Excel Online (Business): Update a row**  -> rename to `Update row`
- Same Location / Library / File / Table
- Key Column: `Key`
- Key Value: `items('For_each_fair')?['key']`
- Start, End, Status, Version, Last_Updated: same as `Add row`
  (leave Outlook_ID empty so it is not changed)

**If no** inside `Has it changed`: leave empty (nothing to do).

Click **Save**.

---

## Part C - First run (now)

1. In the flow, click **Test > Manually > Test**.
2. Wait about 2-4 minutes (29 fairs).
3. Check:
   - Outlook: the fairs from 14 Oct 2026 onwards are in your main calendar.
   - Excel tracker: 29 rows, each with an Outlook_ID.
4. Run **Test** again. Nothing new should be created, because every fair is
   skipped as unchanged. This proves there are no duplicates.

---

## Good to know
- **Don't delete rows in the tracker.** If you do, the flow forgets that fair
  and creates it again next time. If you delete an event in Outlook, also
  delete its row, and the flow will create it fresh.
- **Fairs that vanish from the SECC site** stay in your calendar. The status
  changes to "Not listed since ..." and the event text starts with a note to
  check whether it was moved or cancelled.
- **Your own edits:** you can add categories or notes in Outlook. They are
  only overwritten if SECC changes that fair (then the flow updates it).
- **If the flow fails**, Power Automate emails you. The usual cause is the
  GitHub address being wrong or GitHub not having run yet.

## Common errors
| Error | Fix |
|---|---|
| `The template language expression ... cannot be evaluated` | A step name doesn't match. Rename the step exactly as in this guide. |
| `Find row` fails with a filter error | Check that the Filter Query has the single quotes: `Key eq '...'` |
| `Update event` fails "item not found" | The event was deleted in Outlook. Delete its row in the tracker, then run again. |
| HTTP step 404 | GitHub Pages is not on yet, or the username in the address is wrong. |
