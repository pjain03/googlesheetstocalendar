✨ *Vibe Coded with Antigravity* ✨

# Google Sheets to Google Calendar Birthday Sync

A purely stateless, lock-free Google Apps Script designed to flawlessly map a list of names and dates natively from a Google Sheet directly into a dedicated Google Calendar as recurring annual all-day events. 

## 🏗 Architecture & Design Decisions

This script opts for absolute data correctness and fault tolerance over raw execution speed by utilizing a **Stateless Calendar Wipe** architecture combined with a custom **Properties Polling Debouncer**. 

### 1. The Core Tradeoff: Stateless vs Stateful Syncing
Google Apps Script interacting with the Calendar API faces the "Batch Throttle" problem—creating 50 events using `CalendarApp` takes about ~30 seconds of synchronous loop executions. 
There are generally two ways to build a sync tool:

* **Stateful Syncing (The Fast Way):** 
  When a row is created, the script creates an event and physically stores the generated `Event ID` in a hidden column back on the Spreadsheet. When an edit occurs, the script instantly grabs that exact ID, deletes only that one event, and builds only one new event. This strategy executes entirely in milliseconds. However, it requires cluttering the spreadsheet with internal tracking IDs, creates ghost duplicates if the script crashes, and irreparably breaks if the calendar is ever edited manually by a human or the user clears the IDs column accidentally.
* **Stateless Wiping (Our Approach):** 
  The script assumes nothing and stores nothing. When an edit trigger fires, this script queries the target calendar for the next *100 years*, identifies every single event currently on it, and aggressively deletes them all. Then, it sweeps the entire spreadsheet natively from top to bottom and recreates every event from scratch. 
  * *Tradeoff*: It takes ~30 seconds to rebuild a large calendar.
  * *Benefit*: It is profoundly fault-tolerant. The spreadsheet is the absolute single source of truth. If the calendar is vandalized, dragged, or duplicated manually by mistake, the user simply makes any edit on the spreadsheet, and the calendar magically heals itself perfectly into a 1:1 mirror.

### 2. The Custom Polling Debouncer
Because the script relies on a heavy Wipe-and-Rebuild architecture, a "stampeding herd" problem arises: if a user bulk-pastes 10 names into the sheet at once, Google Sheets fires 10 simultaneous scripts which all try to wipe the calendar simultaneously, causing race conditions and duplicate ghost events.
Standard Google `LockService` solves this via locking the server context, but forces users to wait up to 30 seconds for the queue to resolve. 

Instead, this script uses a **Custom Polling Debouncer**:
1. When an edit triggers the script, it generates a precise Timestamp and saves it to a global `Property`.
2. The script `sleeps` for precisely 2 seconds.
3. It wakes up and checks if the global `Property` still perfectly matches the timestamp it wrote. 
4. If a newer edit occurred while it was sleeping, the property will have changed. The script instantly and silently terminates itself, knowing it is no longer the final edit in the batch.
5. If the property exactly matches, the script officially breaks the loop, assuming the uninterrupted 2-second window implies the user has finished pasting all data, and executes the calendar rebuild precisely *once*.
6. *Final Safety Check*: As the script officially begins the rebuild, it applies a brief 2-second `LockService` lock. In the highly improbable event that two scripts perfectly tie the timestamp check, the lock catches them and queues them cleanly.

### 3. Deleting Events and Series Handling
Google Calendar treats repeating events internally as a "Series". If you use the API to get events from Jan 1st - Dec 31st and request it to delete every single one you find, you will encounter the same "Series" multiple times (e.g., Alice's Birthday on Jan 1st is the same internal series object as Alice's Birthday next year, caught by the 100-year fetch window). Attempting to `deleteEventSeries()` on the second occurrence will throw an API error because the first occurrence already wiped the entire object.

Our script handles this frictionlessly by storing a transient tracking `Set` (`deletedSeriesIds`) exclusively during runtime. As it loops through the 100-year fetch scraping events:
- It checks if the event belongs to a series.
- It grabs the `Series ID`.
- It checks the runtime dictionary. If the ID hasn't been seen, it wipes the series and adds the ID to the dictionary.
- All future occurrences natively skip the API hit, ensuring smooth, silent, optimized deletion wipes.

## 🚀 Setup & Installation

1. Create a Google Sheet. Ensure the first tab is named `Birthdays`.
2. Add the headers `Name` to `A1` and `Date` (DD/MM formatting) to `B1`.
3. Open `Extensions -> Apps Script`.
4. Replace the generic code with the contents of `Code.gs`.
5. Enter a target `CALENDAR_ID` at the top of the file. (It is *highly* recommended to create a dedicated secondary Google Calendar for this, as the script's stateless wipe will aggressively delete everything on the target calendar).
6. Run `setupTrigger()` from the dropdown menu to grant permissions and create the `onInstallableEdit` hook.
7. Start typing!
