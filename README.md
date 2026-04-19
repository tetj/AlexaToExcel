# Alexa To Excel

Exports data from your Alexa account to an XLSX file. Currently exports **reminders** from all your Echo devices, polling on a schedule so your history keeps growing.

You could probably automate this using [Claude Cowork](https://support.claude.com/en/articles/13854387-schedule-recurring-tasks-in-cowork) but here is a free standalone version you can run on your own machine, no subscription required.

---
# Why use this ?

Alexa does not keep the full history, you can only see the last few reminders but I often need the full list to see what I did and when.

# How to use

On first run, a login window will open automatically — just sign in to your Amazon account and the app takes care of the rest.
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

You don't need to copy anything manually. The app handles this automatically:

1. On first run, if no cookie is configured, a built-in browser window will open.
2. Sign in to your Amazon account as normal.
3. Once the app detects your session is ready (it looks for the required `csrf` and `session-id` cookies), the window closes automatically and your credentials are saved to `config.json`.

From that point on, the app runs silently in the background using the saved cookie. If your session expires, the login window will appear again automatically.

> ⚠️ If the login window never closes after signing in, it means the `csrf` cookie was not set. Try navigating manually to `https://alexa.amazon.<your-country>/spa/index.html` inside the login window and wait for it to fully load — this triggers the csrf cookie to be set.

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
| Login window never closes after signing in | `csrf` cookie not set. Navigate to `alexa.amazon.<country>/spa/index.html` inside the window and wait for it to fully load. |
| 401 after working previously | Session expired. Delete `CookieString` from `config.json` and restart — the login window will appear again. |
| `csrf` shown as `(NOT FOUND)` in debug output | The saved cookie is stale. Clear `CookieString` in `config.json` and re-login. |

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
