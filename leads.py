"""Find Google Maps businesses with no website in a bounding box -> .xlsx.

Usage: GOOGLE_MAPS_API_KEY=... python leads.py [--area Nicosia] [--categories "a,b"] [--out leads.xlsx]
       (--area defaults to Cyprus; --bbox S,W,N,E overrides it)
       python leads.py --selftest
"""
import argparse, json, os, urllib.error, urllib.request
from openpyxl import Workbook

URL = "https://places.googleapis.com/v1/places:searchText"
FIELDS = ("places.id,places.displayName,places.formattedAddress,places.nationalPhoneNumber,"
          "places.websiteUri,places.googleMapsUri,places.businessStatus,places.types,"
          "places.rating,places.userRatingCount,places.location,nextPageToken")
CATEGORIES = ("restaurant,cafe,bar,bakery,hair salon,barber shop,beauty salon,nail salon,spa,gym,"
              "plumber,electrician,locksmith,carpenter,painter,roofing contractor,cleaning service,"
              "auto repair,car wash,dentist,doctor,physiotherapist,veterinarian,florist,clothing store,"
              "shoe store,jewelry store,gift shop,grocery store,butcher,hardware store,real estate agency,"
              "accountant,lawyer,photographer,tattoo shop,laundry,tailor,taxi service,driving school")
PAGE_CAP = 60   # Places Text Search returns at most 3 pages x 20 per query
MIN_SPAN = 0.0005  # ~50 m; stop splitting below this
calls = 0


def search(query, tile, key):
    """All pages for one query inside one tile (S, W, N, E)."""
    global calls
    s, w, n, e = tile
    body = {"textQuery": query, "pageSize": 20, "locationRestriction": {"rectangle": {
        "low": {"latitude": s, "longitude": w}, "high": {"latitude": n, "longitude": e}}}}
    out = []
    while True:
        calls += 1
        r = post(body, FIELDS, key)
        out += r.get("places", [])
        if not r.get("nextPageToken"):
            return out
        body["pageToken"] = r["nextPageToken"]


def post(body, mask, key):
    req = urllib.request.Request(URL, json.dumps(body).encode(), {
        "Content-Type": "application/json", "X-Goog-Api-Key": key, "X-Goog-FieldMask": mask})
    try:
        return json.load(urllib.request.urlopen(req))
    except urllib.error.HTTPError as e:  # surface Google's error text (bad key, quota, billing)
        raise SystemExit(f"API error {e.code}: {e.read().decode()}")


def area_tile(name, key):
    """Resolve an area/country name to (S, W, N, E) via the top Text Search hit's viewport."""
    global calls
    calls += 1
    hits = post({"textQuery": name, "pageSize": 1}, "places.displayName,places.viewport", key).get("places")
    if not hits:
        raise SystemExit(f"could not find area: {name}")
    v = hits[0]["viewport"]
    print(f"area: {hits[0]['displayName']['text']}")
    return v["low"]["latitude"], v["low"]["longitude"], v["high"]["latitude"], v["high"]["longitude"]


def crawl(query, tile, fn, found):
    """Search tile; if the 60-result cap was hit, split into quadrants and recurse."""
    places = fn(query, tile)
    for p in places:
        found.setdefault(p["id"], p)
    s, w, n, e = tile
    if len(places) >= PAGE_CAP and min(n - s, e - w) > MIN_SPAN * 2:
        ms, mw = (s + n) / 2, (w + e) / 2
        for t in ((s, w, ms, mw), (s, mw, ms, e), (ms, w, n, mw), (ms, mw, n, e)):
            crawl(query, t, fn, found)


def no_website(found):
    return [p for p in found.values()
            if not p.get("websiteUri") and p.get("businessStatus", "OPERATIONAL") == "OPERATIONAL"]


def export(places, path):
    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Category", "Address", "Phone", "Rating", "Reviews", "Maps URL", "Lat", "Lng", "Place ID"])
    for p in places:
        loc = p.get("location", {})
        ws.append([p.get("displayName", {}).get("text"), (p.get("types") or [""])[0], p.get("formattedAddress"),
                   p.get("nationalPhoneNumber"), p.get("rating"), p.get("userRatingCount"),
                   p.get("googleMapsUri"), loc.get("latitude"), loc.get("longitude"), p["id"]])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def selftest():
    # 200 places on a line; every 4th has a website, every 10th is closed.
    allp = [{"id": str(i), "location": {"latitude": i * 0.001, "longitude": 0},
             **({"websiteUri": "x"} if i % 4 == 0 else {}),
             "businessStatus": "CLOSED_PERMANENTLY" if i % 10 == 0 else "OPERATIONAL"} for i in range(200)]

    def fake(_q, t):
        s, w, n, e = t
        hit = [p for p in allp if s <= p["location"]["latitude"] <= n and w <= p["location"]["longitude"] <= e]
        return hit[:PAGE_CAP]

    found = {}
    crawl("x", (0, -1, 1, 1), fake, found)
    assert len(found) == 200, len(found)  # splitting recovered everything past the 60 cap
    want = {p["id"] for p in allp if "websiteUri" not in p and p["businessStatus"] == "OPERATIONAL"}
    assert {p["id"] for p in no_website(found)} == want
    print("selftest ok")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--area", default="Cyprus", help="area or country name, e.g. Nicosia (default: Cyprus)")
    ap.add_argument("--bbox", help="south,west,north,east; overrides --area")
    ap.add_argument("--categories", default=CATEGORIES)
    ap.add_argument("--out", default="leads.xlsx")
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args()
    if a.selftest:
        selftest()
        raise SystemExit
    key = os.environ.get("GOOGLE_MAPS_API_KEY")
    if not key:
        ap.error("need GOOGLE_MAPS_API_KEY")
    tile = tuple(float(x) for x in a.bbox.split(",")) if a.bbox else area_tile(a.area, key)
    found = {}
    try:
        for cat in a.categories.split(","):
            crawl(cat.strip(), tile, lambda q, t: search(q, t, key), found)
            print(f"{cat.strip()}: {len(found)} places so far, {calls} API calls")
    finally:  # an API error or Ctrl+C midway still saves what was paid for
        leads = no_website(found)
        export(leads, a.out)
        print(f"{len(leads)} of {len(found)} places have no website -> {a.out} ({calls} API calls)")
