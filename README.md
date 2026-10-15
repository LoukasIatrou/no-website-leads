# no-website-leads

Finds businesses on Google Maps that have **no website** listed in an area or country (default: Cyprus) and exports them to an Excel file. The result is a lead list for anyone selling web design.

It is one Python file (`leads.py`) that uses the Google **Places API (New)** Text Search endpoint.

---

## Contents

1. [How it works](#1-how-it-works)
2. [Cost: read this first](#2-cost-read-this-first)
3. [Set up Google Cloud](#3-set-up-google-cloud-one-time)
4. [Install](#4-install)
5. [Verify the install](#5-verify-the-install)
6. [Run a small test](#6-run-a-small-test-do-this-before-cyprus)
7. [Run the real thing](#7-run-the-real-thing)
8. [Options](#8-options)
9. [The Excel output](#9-the-excel-output)
10. [Troubleshooting](#10-troubleshooting)
11. [Limitations](#11-limitations)
12. [Terms of service](#12-terms-of-service)
13. [Outreach (optional)](#13-outreach-optional-outreachpy)

---

## 1. How it works

1. **Resolve the area.** `--area "Nicosia"` is sent as a text search. The top result's map viewport (a rectangle: south, west, north, east) becomes the search area. `--bbox` skips this step and lets you give the rectangle yourself.
2. **Search each category.** For each category (restaurant, plumber, hair salon, ...) the script runs a Text Search restricted to the rectangle. Results come 20 per page, up to 3 pages.
3. **Split when the cap is hit.** Google returns at most 60 results per search. If a search returns the full 60, there are probably more businesses than it could show. The script splits that rectangle into 4 quarters and searches each one again, repeating down to roughly 50 m wide. This is how it gets past the 60-result cap.
4. **Deduplicate.** Places are keyed by Google's place ID, so a business found by several categories or in overlapping tiles appears once.
5. **Filter.** It keeps places with no `websiteUri` whose status is not closed. A place with no status is treated as open.
6. **Export.** It writes the result to `.xlsx`. If the run fails partway (API error, Ctrl+C), whatever was found so far is still saved.

## 2. Cost: read this first

Requesting website and phone fields puts every call in Google's **Text Search Enterprise** pricing tier. At the time of writing that is about **$35 per 1,000 calls**, with a limited monthly free allowance. Check the current figures on Google's [Places API pricing page](https://developers.google.com/maps/billing-and-pricing/pricing).

Number of calls is roughly `categories x tiles x pages`. A whole-country run with the 40 default categories can reach **thousands of calls**, because dense areas keep splitting. Do this before any large run:

- Set a **budget alert** and a **daily quota cap** (see step 3.6).
- Start with a **small area and 2-3 categories** (section 6) and look at the printed call count.
- Scale up from that number.

The script prints the running call count after every category.

## 3. Set up Google Cloud (one time)

1. Go to <https://console.cloud.google.com/> and sign in.
2. Create a project: top bar, project picker, **New Project**. Name it anything, for example `no-website-leads`.
3. Enable billing for the project: **Billing** in the left menu, then link a billing account. The Places API will not work without one.
4. Enable the API: **APIs & Services**, **Library**, search **Places API (New)**, click **Enable**. Make sure it says "(New)". The older "Places API" is a different product and will not work here.
5. Create a key: **APIs & Services**, **Credentials**, **Create credentials**, **API key**. Copy it.
6. Restrict and cap it:
   - Click the key, then under **API restrictions** choose **Restrict key** and select only **Places API (New)**. Save.
   - **APIs & Services**, **Places API (New)**, **Quotas & System Limits**: lower the daily requests limit to a number you can afford.
   - **Billing**, **Budgets & alerts**: create a budget with email alerts.

Treat the key like a password. Do not commit it or paste it into screenshots.

**If a key leaks** (you pasted it somewhere public, or you see charges you don't recognise): go to **APIs & Services**, **Credentials**, click the key, then **Regenerate key** (or **Delete** it and create a new one). The old key stops working right away. Then check **Billing**, **Reports** for unexpected usage. The same goes for the Gmail app password in section 13: revoke it at <https://myaccount.google.com/apppasswords> and make a new one.

## 4. Install

No Git or programming experience is needed. Follow these steps in order.

**4.1 Install Python (once).** Download Python 3.8 or newer from <https://www.python.org/downloads/>. On Windows, tick **"Add python.exe to PATH"** on the first installer screen, then click Install. To check it worked, open a terminal (step 4.3) and run `python --version`. It should print a version number. On macOS/Linux you may need `python3` instead of `python` everywhere in this guide.

**4.2 Get the code.** Pick one:

- **No Git (easiest):** open <https://github.com/LoukasIatrou/no-website-leads>, click the green **Code** button, then **Download ZIP**. Right-click the ZIP and extract it. You now have a folder called `no-website-leads-main`.
- **With Git:** `git clone https://github.com/LoukasIatrou/no-website-leads.git`

**4.3 Open a terminal inside that folder.**

- **Windows:** open the folder in File Explorer, click the address bar, type `cmd`, press Enter.
- **macOS:** right-click the folder, choose **New Terminal at Folder** (enable it under System Settings, Keyboard, Keyboard Shortcuts, Services if you don't see it).
- **Linux:** right-click inside the folder, choose **Open in Terminal**.

Every command in this guide is typed into that terminal window, then Enter.

**4.4 Install the one library the script needs.**

```
pip install openpyxl
```

`openpyxl` writes the Excel file. Nothing else is needed, because HTTP uses the standard library.

**4.5 Set your API key.** It is read from the `GOOGLE_MAPS_API_KEY` environment variable (a setting that lives in the terminal window you typed it in).

| Shell | Command |
|---|---|
| Windows PowerShell | `$env:GOOGLE_MAPS_API_KEY = "your-key"` |
| Windows cmd | `set GOOGLE_MAPS_API_KEY=your-key` |
| macOS / Linux / Git Bash | `export GOOGLE_MAPS_API_KEY=your-key` |

The variable lasts only for that terminal window, so set it again in each new window. Do not put the real key in any file you share or upload to GitHub.

## 5. Verify the install

```
python leads.py --selftest
```

Expected output: `selftest ok`. This test uses fake data and makes **no API calls and costs nothing**. It checks the tile-splitting logic and the no-website filter.

## 6. Run a small test (do this before Cyprus)

```
python leads.py --area Nicosia --categories "plumber,hair salon" --out test.xlsx
```

Expected output looks like:

```
area: Nicosia
plumber: 41 places so far, 9 API calls
hair salon: 118 places so far, 31 API calls
37 of 118 places have no website -> test.xlsx (31 API calls)
```

(The numbers are examples.) Then:

1. Open `test.xlsx`.
2. Pick 3 rows and click their **Maps URL**. Confirm the listing really has no website button.
3. Run the same command again. The row count should be about the same, which confirms deduplication works. Counts can differ slightly because Google's results shift.
4. Note the API call count and multiply by the price in section 2 to estimate a bigger run.

## 7. Run the real thing

Default is the whole of Cyprus, all 40 categories:

```
python leads.py
```

Or a single city, or your own categories:

```
python leads.py --area Limassol
python leads.py --area Larnaca --categories "restaurant,cafe,barber shop"
python leads.py --area Cyprus --categories "plumber,electrician" --out cyprus-trades.xlsx
```

You can stop it with Ctrl+C at any time. It saves what it has found so far.

Note that "Cyprus" resolves to a rectangle that also covers the north of the island and some sea. Empty tiles cost a call each, but they do not split further.

## 8. Options

| Option | Default | Meaning |
|---|---|---|
| `--area NAME` | `Cyprus` | Area or country name, for example `Nicosia`, `Paphos`, `Greece` |
| `--bbox S,W,N,E` | none | Exact rectangle in decimal degrees, for example `35.15,33.33,35.20,33.40`. Overrides `--area` |
| `--categories "a,b,c"` | 40 built-in types | Comma-separated search terms. Any text Google Maps search understands works |
| `--out FILE` | `leads.xlsx` | Output file. Existing files are overwritten |
| `--selftest` | off | Run the offline self-check and exit |

To get a `--bbox`, you can draw a rectangle on <https://bboxfinder.com> and copy the coordinates. Reorder them as south, west, north, east.

The built-in category list is the `CATEGORIES` string near the top of `leads.py`. Edit it to change the default.

## 9. The Excel output

One sheet, header row frozen, filters on every column.

| Column | Content |
|---|---|
| Name | Business name |
| Category | Google's primary type, for example `hair_care` |
| Address | Formatted address |
| Phone | National-format phone number, blank if none listed |
| Rating | Google rating |
| Reviews | Number of reviews |
| Maps URL | Link to the listing |
| Lat / Lng | Coordinates |
| Place ID | Google's unique ID |

Tip: sort or filter by Reviews to find established businesses. Filter out blank Phone cells if you plan to call them.

## 10. Troubleshooting

| Message | Cause and fix |
|---|---|
| `'python' is not recognized` (Windows) or `command not found` | Python is not installed or not on PATH. Reinstall it and tick "Add python.exe to PATH" (section 4.1), or try `python3` / `py`. |
| `need GOOGLE_MAPS_API_KEY` | The environment variable is not set in this terminal. See section 4.5. |
| `need GMAIL_USER, GMAIL_APP_PASSWORD` | Set both variables in this terminal window (section 13.2). |
| `imaplib.error ... AUTHENTICATIONFAILED` | Use the 16-character app password, not your normal Gmail password, and make sure 2-Step Verification is on. |
| `API error 403 ... API_KEY_INVALID` or `PERMISSION_DENIED` | Wrong key, or the key is restricted to a different API. Enable **Places API (New)** and allow it in the key restrictions. |
| `API error 403 ... billing` | Billing is not enabled on the project. See step 3.3. |
| `API error 429` | Quota exceeded: your daily cap, or Google's rate limit. Wait, or raise the quota. Results so far are saved. |
| `could not find area: X` | Google found no place with that name. Try a more specific name such as `Nicosia, Cyprus`, or use `--bbox`. |
| `ModuleNotFoundError: openpyxl` | Run `pip install openpyxl`. |
| Excel file is empty (header only) | No operational places without websites were found. Try other categories or a bigger area. |
| Permission error saving the file | `leads.xlsx` is open in Excel. Close it and run again, or use a different `--out`. |

## 11. Limitations

- **Coverage is best-effort.** Google's text search is ranked and capped. The tile-splitting recovers most businesses, but it cannot guarantee every listing. It also only finds what your categories name.
- **"No website" means no website on the Google listing.** Some of these businesses have a Facebook page, an Instagram page, or a site Google doesn't know about. Check before pitching.
- **Opening status is whatever Google reports.** Some listings are stale.
- **No retry or resume.** An API error stops the run; the partial file is saved, but a new run starts from scratch and pays again.
- **Rate limits.** Calls are sequential, so rate limiting is unlikely for normal use.

## 12. Terms of service

The Google Maps Platform terms limit how Places data may be stored and reused. This tool is intended for your own business outreach research. Do not resell or republish the exported data, and read Google's current terms before relying on it commercially. Contacting businesses is also subject to the rules where they operate, such as GDPR and ePrivacy in the EU, which covers Cyprus.

## 13. Outreach (optional): `outreach.py`

Takes the Excel file from `leads.py`, drops businesses that have a website anywhere, looks for an email address, and saves an intro as a **Gmail draft**. It never sends anything: you open Gmail, read each draft, and press Send yourself.

> **Status: experimental.** The web-search step scrapes Bing, and Bing currently returns unrelated results to scripted requests, so most rows will end up as `no email`. The search function (`search()` in `outreach.py`) needs a better backend before this is useful. Everything else (Gmail drafts, the skip-already-done logic) is in place. Do not send drafts without checking them.

**Before you start:** finish sections 3 to 7 so you have a `leads.xlsx` (or `test.xlsx`) file in the folder.

**13.1 Make a Gmail app password.** This is a separate 16-character password that lets the script create drafts without your real password.

1. Go to <https://myaccount.google.com/security> and turn on **2-Step Verification** if it is off.
2. Open <https://myaccount.google.com/apppasswords>, name it `leads`, click **Create**, and copy the 16 characters.

**13.2 Set the two variables** in the terminal (see section 4.3), in the same window you will run the script in:

| Shell | Commands |
|---|---|
| Windows PowerShell | `$env:GMAIL_USER = "you@gmail.com"` then `$env:GMAIL_APP_PASSWORD = "abcdefghijklmnop"` |
| Windows cmd | `set GMAIL_USER=you@gmail.com` then `set GMAIL_APP_PASSWORD=abcdefghijklmnop` |
| macOS / Linux / Git Bash | `export GMAIL_USER=you@gmail.com` then `export GMAIL_APP_PASSWORD=abcdefghijklmnop` |

**13.3 Write your email.** Open `outreach.py` in Notepad (or any text editor) and edit `SUBJECT` and `TEMPLATE` near the top. `{name}` is replaced by each business name. Replace the sample text and the "Louks" signature with your own, and keep the opt-out line.

**13.4 Check the script works offline** (no cost):

```
python outreach.py --selftest
```

It should print `selftest ok`.

**13.5 Try 2 leads first:**

```
python outreach.py test.xlsx --limit 2
```

Then open Gmail, go to **Drafts**, and read them. Check the recipient and name are right. Run the same command again: it must not create the same drafts twice.

**13.6 Run more.** `--limit 20` is the default and a sensible batch size. Re-running continues where it stopped. Press Ctrl+C at any time; progress is saved.

**Notes**

- Adds `Email` and `Status` columns (`site <url>`, `no email`, or `drafted <date>`). Rows with a Status are skipped on re-runs, so nothing is searched or drafted twice.
- Edit `SUBJECT` / `TEMPLATE` at the top of `outreach.py` first.
- A "site" is any search result whose domain contains the business name and is not social media or a directory. Greek-only names will not match; check by hand.
- Keep it B2B, keep the opt-out line, and honour replies (ePrivacy/GDPR).
- `python outreach.py --selftest` runs offline.
