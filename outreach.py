"""Leads xlsx -> Gmail drafts: drop businesses that have a site anywhere, find an email, draft an intro.

Usage: GMAIL_USER=... GMAIL_APP_PASSWORD=... python outreach.py leads.xlsx [--limit 20]
       python outreach.py --selftest
Adds Email + Status columns to the xlsx; rows with a Status are skipped on re-runs.
"""
import argparse, base64, datetime, html, imaplib, os, re, time, urllib.parse, urllib.request
from email.message import EmailMessage
from openpyxl import load_workbook

SEARCH = "https://www.bing.com/search?q="
SOCIAL = {"facebook", "instagram", "tiktok", "linkedin", "twitter", "x", "youtube", "tripadvisor", "yelp",
          "google", "foursquare", "wolt", "foody", "bing", "yellowpages", "11888", "cylex", "wikipedia"}
JUNK = ("example.", "sentry", "wixpress", ".png", ".jpg", ".gif", ".webp", "schema.org")
EMAIL = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
SUBJECT = "Quick idea for {name}"
TEMPLATE = """Hi {name} team,

I noticed {name} doesn't have a website yet. I build simple, fast sites for local businesses
like yours, so customers searching on Google can find you, see your work and contact you.

Happy to send a free mockup if you're interested. Just reply to this email.

Best,
Louks

(Not interested? Reply "no" and I won't contact you again.)
"""


def get(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Accept-Language": "en"})
    return urllib.request.urlopen(req, timeout=15).read().decode("utf-8", "ignore")


def search(q):
    """Top Bing results as [{url, title, description}]. Free HTML scrape; no API key."""
    out = []
    for blk in re.findall(r'<li class="b_algo".*?</li>', get(SEARCH + urllib.parse.quote(q)), flags=re.S):
        a = re.search(r'<h2[^>]*><a[^>]*href="([^"]+)"[^>]*>(.*?)</a>', blk, flags=re.S)
        if not a:
            continue
        url = html.unescape(a[1])
        if u := re.search(r"[?&]u=a1([\w-]+)", url):  # Bing wraps links as base64url(real url)
            url = base64.urlsafe_b64decode(u[1] + "=" * (-len(u[1]) % 4)).decode("utf-8", "ignore")
        text = lambda h: html.unescape(re.sub(r"<[^>]+>", " ", h))
        p = re.search(r"<p[^>]*>(.*?)</p>", blk, flags=re.S)
        out.append({"url": url, "title": text(a[2]), "description": text(p[1]) if p else ""})
    return out[:5]  # ponytail: Bing may captcha/layout-change a big run; swap in another search backend here


def has_site(results, name):
    """URL of a result that looks like the business's own site, else None."""
    words = re.findall(r"[a-z0-9]{4,}", name.lower())
    if not words:
        return None
    key = max(words, key=len)  # ponytail: longest-word match misses Greek names / unrelated domains; loosen if needed
    for r in results:
        host = urllib.parse.urlparse(r["url"]).netloc.lower()
        if not SOCIAL & set(host.split(".")) and key in host.replace("-", ""):  # label match: "bing" must not hit "plumbing"
            return r["url"]


def find_email(results):
    """First plausible email in the result snippets, then in the (non-login) result pages."""
    def scan(text):
        return next((m for m in EMAIL.findall(text) if not any(j in m.lower() for j in JUNK)), None)
    for r in results:
        if hit := scan(r.get("title", "") + " " + r.get("description", "")):
            return hit
    for r in results:
        if any(s in r["url"] for s in ("facebook", "instagram")):
            continue
        try:
            if hit := scan(get(r["url"])):
                return hit
        except Exception:  # dead/blocked page: just try the next one
            pass


def draft(imap, to, name):
    msg = EmailMessage()
    msg["To"], msg["Subject"] = to, SUBJECT.format(name=name)
    msg.set_content(TEMPLATE.format(name=name))
    imap.append('"[Gmail]/Drafts"', "\\Draft", imaplib.Time2Internaldate(time.time()), msg.as_bytes())


def run(path, limit):
    user, pw = os.environ.get("GMAIL_USER"), os.environ.get("GMAIL_APP_PASSWORD")
    if not (user and pw):
        raise SystemExit("need GMAIL_USER, GMAIL_APP_PASSWORD")
    wb = load_workbook(path)
    ws = wb.active
    cols = [c.value for c in ws[1]]
    for h in ("Email", "Status"):
        if h not in cols:
            ws.cell(1, len(cols) + 1, h)
            cols.append(h)
    name_i, addr_i, email_i, status_i = (cols.index(h) + 1 for h in ("Name", "Address", "Email", "Status"))
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(user, pw)
    n = site = noemail = drafted = 0
    try:
        for row in range(2, ws.max_row + 1):
            if drafted >= limit:
                break
            if ws.cell(row, status_i).value:
                continue
            name = ws.cell(row, name_i).value or ""
            n += 1
            results = search(f'"{name}" {ws.cell(row, addr_i).value or ""}')
            time.sleep(2)  # be polite; scraping a search page has no quota but does get blocked
            if url := has_site(results, name):
                ws.cell(row, status_i, f"site {url}")
                site += 1
            elif to := find_email(results):
                draft(imap, to, name)
                ws.cell(row, email_i, to)
                ws.cell(row, status_i, f"drafted {datetime.date.today()}")
                drafted += 1
            else:
                ws.cell(row, status_i, "no email")
                noemail += 1
    finally:  # an API error or Ctrl+C midway still saves what was paid for
        wb.save(path)
        print(f"{n} searched: {site} have a site, {noemail} no email, {drafted} drafted")


def selftest():
    hits = [{"url": "https://www.facebook.com/nikos-plumbing"}, {"url": "https://nikosplumbing.com.cy/contact"}]
    assert has_site(hits, "Nikos Plumbing") == hits[1]["url"]
    assert has_site(hits[:1], "Nikos Plumbing") is None  # social page is not a site
    assert has_site([{"url": "https://acme.com"}], "Nikos Plumbing") is None  # unrelated domain
    assert has_site(hits, "Bo") is None  # no usable name word
    snip = [{"url": "https://x.cy", "title": "Nikos", "description": "mail logo@wixpress.com or info@nikos.com.cy"}]
    assert find_email(snip) == "info@nikos.com.cy"
    print("selftest ok")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", nargs="?")
    ap.add_argument("--limit", type=int, default=20, help="max drafts per run")
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args()
    if a.selftest:
        selftest()
    elif a.xlsx:
        run(a.xlsx, a.limit)
    else:
        ap.error("need an xlsx path")
