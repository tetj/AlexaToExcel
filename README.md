# Alexa To Excel

Exports data from your Alexa account to an XLSX file. Currently exports **reminders** from all your Echo devices, polling on a schedule so your history keeps growing.

You could probably automate this using [Claude Cowork](https://support.claude.com/en/articles/13854387-schedule-recurring-tasks-in-cowork) but here is a free standalone version you can run on your own machine, no subscription required.

---
# Why use this ?

Alexa does not keep the full history, you can only see the last few reminders but I often need the full list to see what I did and when.

# How to use

First, see section [How to get your cookie](#how-to-get-your-cookie) below
```
# Run once for a US account
AlexaToExcel.exe --country us --poll-interval 0
or AlexaToExcel.exe -c us -p 0

# Poll every 30 minutes for a UK account
AlexaToExcel.exe --country uk --poll-interval 30
or AlexaToExcel.exe -c uk -p 30
```

# How to run on startup

1. Download AlexaToExcel.exe from the [releases](https://github.com/tetj/AlexaToExcel/releases/)
2. Right-click on AlexaToExcel.exe and select Copy.
3. Press Windows Key + R to open the Run dialog.
4. Type shell:startup and click OK.
5. In the folder that opens, right-click and select Paste.

---

## Command-line Arguments

Arguments override the corresponding values in `config.json` for that run only — they are never written back to the file.

| Argument | Short | Description |
|---|---|---|
| `--country <code>` | `-c` | Set the Amazon country. Overrides `BaseUrl` in config.json. |
| `--poll-interval <mins>` | `-p` | Poll interval in minutes. Use `0` to run once and exit. Overrides `PollIntervalMinutes` in config.json. |
| `--debug` | | Print raw HTTP request/response details. |
| `--help` | | Show usage summary and exit. |

**Supported country codes:**

| Code | Store |
|---|---|
| `us` | amazon.com (United States) |
| `ca` | amazon.ca (Canada) |
| `uk` or `gb` | amazon.co.uk (United Kingdom) |
| `de` | amazon.de (Germany) |
| `fr` | amazon.fr (France) |
| `es` | amazon.es (Spain) |
| `it` | amazon.it (Italy) |
| `au` | amazon.com.au (Australia) |
| `jp` | amazon.co.jp (Japan) |
| `in` | amazon.in (India) |
| `mx` | amazon.com.mx (Mexico) |
| `br` | amazon.com.br (Brazil) |

---

## How to get your cookie

The most common cause of 401 errors is copying the cookie from the wrong place. Follow these steps **exactly**:

### Step 1 — Open the right URL in Chrome

Open Chrome and navigate to:
```
https://alexa.amazon.ca/api/devices-v2/device?raw=false
```
(Use `alexa.amazon.com` if you're on a US account — always `alexa.` not `www.`)

You should see a JSON response listing your Echo devices. If you see a login page, sign in first then try again.

### Step 2 — Open DevTools Network tab

Press **F12** → click the **Network** tab → press **F5** to reload.

### Step 3 — Find the request

In the Network tab, click the request named **`device`** (it will be the first one after reload).

### Step 4 — Copy the cookie header

- Click the **Headers** sub-tab → scroll to **Request Headers**
- Find **`cookie:`** (lowercase, not `Cookie:` or `Set-Cookie:`)
- Click the value to select it, then press **Ctrl+A** to select all and **Ctrl+C** to copy

> ⚠️ Do NOT use the browser console (`document.cookie`) — it cannot read HttpOnly cookies, which are required for authentication. Copy only the value from the Network tab.

### Step 5 — Paste into config.json

Open `config.json` and paste the value as the `CookieString`:

```json
{
  "BaseUrl": "https://www.amazon.ca",
  "CookieString": "session-id=xxx; csrf=1234567890; ubid-acbca=xxx; ...",
  "OutputPath": "alexa_reminders.xlsx",
  "PollIntervalMinutes": 60
}
```

### Step 6 — Verify csrf is present

Your cookie string **must contain `csrf=`** (a number, like `csrf=1465446206`).  
If it's missing, do this first:
1. Go to `https://alexa.amazon.ca/spa/index.html`
2. Wait for the page to fully load (this sets the csrf cookie)
3. Then repeat from Step 1

---

## Troubleshooting 401 errors

Run with `--debug` to see exactly what's happening:

```
AlexaToExcel.exe --debug
```

This prints:
- The exact URL being called
- The HTTP status code returned
- The response body (which often contains a helpful error message)
- Which cookies were found (by name only, not values)
- The csrf value being sent

**Common causes:**

| Symptom | Fix |
|---|---|
| `csrf` shown as `(NOT FOUND)` | Cookie was copied from `www.amazon.ca`, not `alexa.amazon.ca`. Redo Step 1. |
| `csrf` value looks wrong | Make sure you copied the *value* of the `cookie:` header, not a single cookie. |
| 401 after working previously | Session expired. Re-copy cookie from browser. |

---

## Configuration

| Field | Description |
|---|---|
| `BaseUrl` | `https://www.amazon.ca` for Canada, `https://www.amazon.com` for US, etc. |
| `CookieString` | Full `cookie:` request header value from Chrome DevTools |
| `OutputPath` | Where to save the XLSX file |
| `PollIntervalMinutes` | How often to poll. Set to `0` to run once and exit. |

---

## Potential future exports

The same cookie authentication gives access to other Alexa data. Planned candidates:

- **Shopping list** — items added via voice or the Alexa app (` /api/todos?type=SHOPPING_ITEM `)
- **To-do list** — tasks created on your Echo devices (` /api/todos?type=TASK `)
- **Voice history** — a log of everything said to Alexa, with timestamps (` /api/activities `)

# Exporting Google Sheet as CSV from HTML

a) To export a Google Sheet as HTML, you can replace the **/edit** portion of the spreadsheet URL with specific parameters such as :

1. Preview Mode: Replace /edit with /preview.
2. HTML View: Replace /edit with /htmlview. 
3. Direct HTML Export: Replace /edit with /export?format=html. 

b) I tested with **/preview** and I was able to download it as HTML using **right-click -> Save As** ...

c) Get the latest release here : https://github.com/tetj/AlexaToExcel/releases/

d) Then using a command prompt :

```
AlexaToExcel --html-to-csv "file.html" output.csv
```
