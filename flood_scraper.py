import os
import sys
import time
import csv
import glob
import random
from datetime import datetime, timezone, timedelta
import requests
import pandas as pd

# ========================= CONFIG =========================
API_KEY = os.getenv("TWITTERAPI_IO_KEY", "YOUR_API_KEY_HERE")

DEFAULT_CREDITS = 27980
CREDITS_AVAILABLE = int(os.getenv("TWITTERAPI_CREDITS", DEFAULT_CREDITS))
CREDITS_PER_TWEET = 15

# Test cap: 0 = full run
TEST_LIMIT_TWEETS = int(os.getenv("TEST_LIMIT_TWEETS", "100"))

# Free tier pacing (1 req / ~5s). Adjust via env if needed.
MIN_INTERVAL_SEC = float(os.getenv("MIN_INTERVAL_SEC", "6.2"))
BACKOFF_MIN = 3.0
BACKOFF_MAX = 60.0

# Stop if we see this many consecutive pages with +0 additions
STOP_AFTER_ZERO_PAGES = int(os.getenv("STOP_AFTER_ZERO_PAGES", "10"))

# Only fetch tweets strictly **after** the newest one already saved (default ON)
ONLY_AFTER_LAST_SAVED = os.getenv("ONLY_AFTER_LAST_SAVED", "1") == "1"

SAVE_CHECKPOINT_EVERY = 2000
CHECKPOINT_CSV = "Pakistan_Floods_2025_checkpoint.csv"
SEEN_IDS_FILE = "seen_ids.txt"
STATE_CURSOR_FILE = "cursor_state.txt"
OUTPUT_XLSX = "Pakistan_Floods_2025_Tweets.xlsx"

BASE_URL = "https://api.twitterapi.io/twitter/tweet/advanced_search"
HEADERS = {"x-api-key": API_KEY}

# === Filters ===
# Default ON to be maximally strict
STRICT_PAK_FILTER = os.getenv("STRICT_PAK_FILTER", "1") == "1"

PAK_POSITIVE = [
    "pakistan","punjab pakistan","punjab, pakistan","sindh","balochistan",
    "khyber pakhtunkhwa","kpk","azad kashmir","gilgit-baltistan",
    "karachi","lahore","quetta","peshawar","islamabad",
    "multan","faisalabad","rawalpindi","hyderabad pakistan","hyderabad, pakistan"
]

# Countries/regions to exclude if mentioned in text or author location
OTHER_COUNTRY_NEG = [
    "india","indian","punjab, india","indian punjab","amritsar","ludhiana","chandigarh","jalandhar",
    "bathinda","patiala","mohali","haryana","himachal pradesh","delhi","new delhi","bharat","hindustan",
    "bangladesh","dhaka","chittagong","nepal","kathmandu","sri lanka","colombo",
    "afghanistan","kabul","iran","tehran","china","xinjiang","uae","dubai",
    "saudi","riyadh","qatar","oman","turkey"
]

# Pakistan-only hashtags (list for post-filter and a joined string for query)
HASHTAGS_PAK_LIST = [
    "#MonsoonPakistan", "#MonsoonFloods", "#FloodsPakistan", "#FloodInPakistan","#FloodingPakistan",
    "#PakistanFloods", "#FloodsInPakistan", "#PakistanRains", "#PakistanRain",
    "#SindhFloods", "#BalochistanFloods", "#KPFloods", "#KPKFloods", "#PunjabFloods",
    "#GBFloods", "#AJKFloods",
    "#KarachiRain", "#LahoreRain", "#QuettaRain", "#PeshawarRain", "#IslamabadRain",
    "#HyderabadRain", "#MultanRain", "#FaisalabadRain", "#RawalpindiRain",
    "#سیلاب", "#پاکستان_سیلاب", "#بارش", "#طوفانی_بارش"
]
HASHTAGS_PAK_QUERY = " OR ".join(HASHTAGS_PAK_LIST)

# Extra negatives for other countries (used inside the query too)
NEGATIVE_REGION_QUERY = (
    '"Punjab, India" OR "Indian Punjab" OR Amritsar OR Ludhiana OR Chandigarh OR Jalandhar OR '
    'Bathinda OR Patiala OR Mohali OR Haryana OR "Himachal Pradesh" OR Delhi OR "New Delhi" OR '
    'Gujarat OR Maharashtra OR "Uttar Pradesh" OR Rajasthan OR Bharat OR Hindustan OR India OR '
    'Bangladesh OR Dhaka OR Chittagong OR Nepal OR Kathmandu OR "Sri Lanka" OR Colombo OR '
    'Afghanistan OR Kabul OR Iran OR Tehran OR China OR Xinjiang OR UAE OR Dubai OR '
    'Saudi OR Riyadh OR Qatar OR Oman OR Turkey'
)

def looks_pakistan_only(tweet: dict) -> bool:
    """
    Post-filter:
    - If geotag PK -> accept.
    - If geotag non-PK -> reject.
    - Else require PK-positive token or PK hashtag, and reject if any other-country term appears
      in text or author location.
    """
    txt = (tweet.get("text") or tweet.get("full_text") or "").lower()
    author = tweet.get("author") or tweet.get("user") or {}
    author_loc = (author.get("location") or "").lower()
    place = tweet.get("place") or {}
    cc = (place.get("country_code") or place.get("countryCode") or "").upper()

    if cc == "PK":
        return True
    if cc and cc != "PK":
        return False

    # reject if any non-PK country/region signals show up
    for neg in OTHER_COUNTRY_NEG:
        if neg in txt or neg in author_loc:
            return False

    # require at least one PK-positive token if no geotag
    for pos in PAK_POSITIVE:
        if pos in txt or pos in author_loc:
            return True

    # OR allow a PK-specific hashtag
    txt_no_hash = txt  # already lowercased
    for h in HASHTAGS_PAK_LIST:
        if h.lower() in txt_no_hash:
            return True

    return False

# ---- Time window base ----
def fmt_utc(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%d_%H:%M:%S_UTC")

SINCE_UTC_BASE = "2025-01-01_00:00:00_UTC"
UNTIL_UTC_DEFAULT = fmt_utc(datetime.now(timezone.utc))

# Optional explicit overrides to time-slice (set via env)
SINCE_UTC_OVERRIDE = os.getenv("SINCE_UTC_OVERRIDE", "").strip()
UNTIL_UTC_OVERRIDE = os.getenv("UNTIL_UTC_OVERRIDE", "").strip()

def build_query(since_utc: str, until_utc: str) -> str:
    """
    Pakistan-strict query:
    - Require PK geotag OR explicit PK provinces/cities OR PK-specific hashtags.
    - Disallow geotags for India + neighbors.
    - Apply extra textual negatives for those regions.
    """
    pk_anchor = (
        '(place_country:PK OR '
        '"Punjab Pakistan" OR "Punjab, Pakistan" OR Sindh OR Balochistan OR '
        '"Khyber Pakhtunkhwa" OR KPK OR "Azad Kashmir" OR "Gilgit-Baltistan" OR '
        'Karachi OR Lahore OR Quetta OR Peshawar OR Islamabad OR '
        'Multan OR Faisalabad OR Rawalpindi OR "Hyderabad Pakistan" OR "Hyderabad, Pakistan" OR '
        f'({HASHTAGS_PAK_QUERY}))'
    )

    flood_terms = '(flood OR floods OR flooding OR "flash flood" OR "river flood" OR monsoon OR "heavy rain")'

    return (
        f'({flood_terms}) '
        f'AND {pk_anchor} '
        '-place_country:IN -place_country:BD -place_country:NP -place_country:LK '
        '-place_country:AF -place_country:IR -place_country:CN -place_country:AE '
        '-place_country:SA -place_country:QA -place_country:OM -place_country:TR '
        f'-({NEGATIVE_REGION_QUERY}) '
        '(lang:en OR lang:ur) '
        f"since:{since_utc} until:{until_utc}"
    )

# Sort mode via env: QUERY_TYPE=Top or Latest
QUERY_TYPE = os.getenv("QUERY_TYPE", "Latest")

# Columns (includes Verified + Location)
COLUMNS = [
    "Tweet Content",
    "URL",
    "Date Posted",
    "Number of likes",
    "Account",
    "Followers",
    "Following",
    "Retweets",
    "Verified",
    "Location",
]

# ========================= HELPERS =========================
class RateLimitError(Exception):
    def __init__(self, retry_after_sec: int, body: str = ""):
        super().__init__(f"429 rate limit. Retry after ~{retry_after_sec}s. Body: {body[:200]}")
        self.retry_after_sec = retry_after_sec

class CreditsError(Exception):
    def __init__(self, body: str = ""):
        super().__init__(f"402 credits exhausted. Body: {body[:200]}")

def sleep_min_interval(last_ts: float) -> float:
    now = time.time()
    wait = MIN_INTERVAL_SEC - (now - last_ts)
    if wait > 0:
        time.sleep(wait + random.uniform(0.05, 0.35))
    return time.time()

def merge_entities(t: dict) -> dict:
    urls = []
    ent = t.get("entities") or {}
    if isinstance(ent, dict):
        urls += ent.get("urls") or []
    ext = t.get("extended_tweet") or {}
    if isinstance(ext, dict):
        ee = ext.get("entities") or {}
        if isinstance(ee, dict):
            urls += ee.get("urls") or []
    nt = t.get("note_tweet") or t.get("noteTweet") or {}
    if isinstance(nt, dict):
        ne = nt.get("entities") or {}
        if isinstance(ne, dict):
            urls += ne.get("urls") or []
    return {"urls": urls}

def extract_text(t: dict) -> str:
    nt = t.get("note_tweet") or t.get("noteTweet") or {}
    if isinstance(nt, dict):
        if nt.get("text"):
            return nt["text"]
        note = nt.get("note")
        if isinstance(note, dict) and note.get("text"):
            return note["text"]
    long_txt = t.get("longText") or t.get("long_text")
    if isinstance(long_txt, dict) and long_txt.get("content"):
        return long_txt["content"]
    if isinstance(long_txt, str) and long_txt:
        return long_txt
    ext = t.get("extended_tweet") or {}
    if isinstance(ext, dict) and ext.get("full_text"):
        return ext["full_text"]
    if t.get("full_text"):
        return t["full_text"]
    if t.get("text"):
        return t["text"]
    return ""

def replace_tco_with_expanded(text: str, entities: dict) -> str:
    if not text:
        return text
    urls = []
    if isinstance(entities, dict):
        urls = entities.get("urls") or []
    for u in urls:
        short = u.get("url")
        expanded = (
            u.get("expanded_url") or u.get("expandedUrl") or
            u.get("unwound_url") or u.get("display_url")
        )
        if short and expanded:
            try:
                text = text.replace(short, expanded)
            except Exception:
                pass
    return text

def extract_location(t: dict) -> str:
    place = t.get("place") or {}
    if isinstance(place, dict):
        full = place.get("full_name") or place.get("fullName")
        if full:
            return str(full)
        name = place.get("name")
        country = place.get("country")
        cc = place.get("country_code") or place.get("countryCode")
        if name and country:
            return f"{name}, {country}"
        if name and cc:
            return f"{name}, {cc}"
    geo = t.get("geo") or {}
    if isinstance(geo, dict):
        coords_obj = geo.get("coordinates") or {}
        if isinstance(coords_obj, dict):
            coords = coords_obj.get("coordinates") or coords_obj.get("coord")
            if isinstance(coords, (list, tuple)) and len(coords) == 2:
                # lon, lat -> "lat,lon"
                return f"{coords[1]},{coords[0]}"
    coords_obj = t.get("coordinates") or {}
    if isinstance(coords_obj, dict):
        coords = coords_obj.get("coordinates")
        if isinstance(coords, (list, tuple)) and len(coords) == 2:
            return f"{coords[1]},{coords[0]}"
    return ""

def fetch_page(session: requests.Session, cursor: str, query: str):
    params = {"query": query, "queryType": QUERY_TYPE}
    if cursor:
        params["cursor"] = cursor
    r = session.get(BASE_URL, headers=HEADERS, params=params, timeout=60)
    if r.status_code == 429:
        retry_after = r.headers.get("Retry-After")
        raise RateLimitError(int(retry_after) if retry_after and retry_after.isdigit() else 5, r.text)
    if r.status_code == 402:
        raise CreditsError(r.text)
    if r.status_code != 200:
        raise RuntimeError(f"HTTP {r.status_code}: {r.text[:500]}")
    data = r.json()
    tweets = data.get("tweets", []) or []
    next_cursor = data.get("next_cursor") or data.get("nextCursor") or ""
    has_next = bool((data.get("has_next_page") or data.get("hasNextPage")) and next_cursor)
    return tweets, next_cursor, has_next

def extract_author(t: dict):
    a = t.get("author") or t.get("user") or {}
    username = a.get("userName") or a.get("username") or a.get("screenName") or a.get("screen_name") or ""
    followers = a.get("followers") or a.get("followers_count") or 0
    following = a.get("following") or a.get("friends_count") or 0
    verified_raw = (
        a.get("verified") or a.get("isVerified") or a.get("blueVerified") or
        a.get("is_blue_verified") or a.get("isBlueVerified")
    )
    vtype = (a.get("verifiedType") or "").lower()
    verified = bool(verified_raw) or vtype in {"blue", "business", "government", "official"}
    return username, int(followers or 0), int(following or 0), verified

def ensure_url_from_id_and_user(t: dict, username: str) -> str:
    tid = str(t.get("id", "")).strip()
    if not tid:
        return ""
    return f"https://x.com/{username}/status/{tid}" if username else f"https://x.com/i/web/status/{tid}"

def row_from_tweet(t: dict):
    username, followers, following, verified = extract_author(t)
    text = extract_text(t)
    entities = merge_entities(t)
    text = replace_tco_with_expanded(text, entities)
    created = t.get("createdAt") or t.get("created_at") or t.get("created_at_time") or ""
    likes = t.get("likeCount") or t.get("favorite_count") or 0
    rts = t.get("retweetCount") or t.get("retweet_count") or 0
    url = t.get("url") or ensure_url_from_id_and_user(t, username)
    location = extract_location(t)
    return {
        "Tweet Content": text,
        "URL": url,
        "Date Posted": created,
        "Number of likes": int(likes or 0),
        "Account": f"@{username}" if username else "",
        "Followers": int(followers or 0),
        "Following": int(following or 0),
        "Retweets": int(rts or 0),
        "Verified": "Yes" if verified else "No",
        "Location": location,
    }

# --------- FILE-LOCK-RESILIENT WRITES ----------
def _safe_write_csv(rows, path, mode="a", retries=8):
    delay = 1.0
    for i in range(retries):
        try:
            write_header = (mode == "w") or (not os.path.exists(path)) or (os.path.getsize(path) == 0)
            with open(path, mode, newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=COLUMNS)
                if write_header:
                    w.writeheader()
                w.writerows(rows)
            return path
        except PermissionError as e:
            print(f"[FILE-LOCK] {e}. Retry in {delay:.1f}s (attempt {i+1}/{retries})…")
            time.sleep(min(delay, 10.0))
            delay = min(delay * 1.8, 8.0)
    base, ext = os.path.splitext(path)
    alt = f"{base}.staging.{int(time.time())}{ext}"
    print(f"[FILE-LOCK] Still locked. Writing to sidecar: {alt}")
    write_header = (not os.path.exists(alt)) or (os.path.getsize(alt) == 0)
    with open(alt, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        if write_header:
            w.writeheader()
        w.writerows(rows)
    return alt

def save_checkpoint(rows, path=CHECKPOINT_CSV, mode="a"):
    _safe_write_csv(rows, path, mode=mode)

def load_seen_ids(path=SEEN_IDS_FILE):
    s = set()
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    s.add(line)
    return s

def persist_seen_ids(seen, path=SEEN_IDS_FILE):
    try:
        with open(path, "w", encoding="utf-8") as f:
            for tid in seen:
                f.write(str(tid) + "\n")
    except PermissionError as e:
        print(f"[FILE-LOCK] Can't update {path} now: {e}")

def load_cursor(path=STATE_CURSOR_FILE):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read().strip()
    return ""

def save_cursor(cursor: str, path=STATE_CURSOR_FILE):
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(cursor or "")
    except PermissionError as e:
        print(f"[FILE-LOCK] Can't update {path}: {e}")

def dedupe_and_export(csv_path=CHECKPOINT_CSV, xlsx_path=OUTPUT_XLSX):
    parts = []
    if os.path.exists(csv_path):
        parts.append(csv_path)
    parts += sorted(glob.glob(csv_path.replace(".csv", ".staging*.csv")))
    if not parts:
        print("No CSV parts found; skipping export.")
        return
    frames = []
    for p in parts:
        try:
            frames.append(pd.read_csv(p))
        except Exception as e:
            print(f"[WARN] Could not read {p}: {e}")
    if not frames:
        print("No readable CSV content; skipping export.")
        return
    df = pd.concat(frames, ignore_index=True)
    if "URL" in df.columns:
        df = df.drop_duplicates(subset=["URL"], keep="first")
    keys = [c for c in ["Tweet Content", "Date Posted"] if c in df.columns]
    if keys:
        df = df.drop_duplicates(subset=keys, keep="first")
    df = df.reindex(columns=COLUMNS)
    tmp_csv = csv_path + ".consolidated.tmp"
    try:
        df.to_csv(tmp_csv, index=False, encoding="utf-8")
        os.replace(tmp_csv, csv_path)
        for p in parts:
            if p != csv_path and ".staging." in p:
                try:
                    os.remove(p)
                except Exception:
                    pass
        print(f"Consolidated CSV → {csv_path}")
    except PermissionError as e:
        print(f"[FILE-LOCK] Can't update {csv_path} now. Consolidated copy left at {tmp_csv}. "
              f"Close Excel and rerun export later.")
    try:
        df.to_excel(xlsx_path, index=False)
        print(f"Excel written → {xlsx_path}")
    except PermissionError as e:
        print(f"[FILE-LOCK] Can't write {xlsx_path}. Close Excel and rerun to export.")

# ========================= MAIN =========================
def main():
    global MIN_INTERVAL_SEC

    if not API_KEY or API_KEY.strip() == "":
        print("ERROR: Missing TWITTERAPI_IO_KEY.", file=sys.stderr)
        sys.exit(1)

    # --------- ADDITIVE TARGET ----------
    target_by_credits = max(0, CREDITS_AVAILABLE // CREDITS_PER_TWEET)

    existing = 0
    if os.path.exists(CHECKPOINT_CSV):
        try:
            for chunk in pd.read_csv(CHECKPOINT_CSV, chunksize=50000):
                existing += len(chunk)
        except Exception as e:
            print(f"Warning scanning checkpoint: {e}")

    add_this_run = target_by_credits
    if TEST_LIMIT_TWEETS > 0:
        add_this_run = min(add_this_run, TEST_LIMIT_TWEETS)

    run_target = existing + add_this_run
    print(f"Existing: {existing} | Add this run: {add_this_run} | Target total: {run_target}")
    print(f"Credits: {CREDITS_AVAILABLE} | Cost/tweet: {CREDITS_PER_TWEET} | Cap by credits (this run): {target_by_credits}")
    print(f"QUERY_TYPE = {QUERY_TYPE}")
    print(f"Baseline interval ≈ {MIN_INTERVAL_SEC:.1f}s per request (free-tier friendly)")

    # --------- DYNAMIC since/until ----------
    effective_since_dt = datetime.strptime(SINCE_UTC_BASE, "%Y-%m-%d_%H:%M:%S_UTC").replace(tzinfo=timezone.utc)
    effective_until = UNTIL_UTC_DEFAULT

    if SINCE_UTC_OVERRIDE and UNTIL_UTC_OVERRIDE:
        try:
            effective_since_dt = datetime.strptime(SINCE_UTC_OVERRIDE, "%Y-%m-%d_%H:%M:%S_UTC").replace(tzinfo=timezone.utc)
            effective_until = UNTIL_UTC_OVERRIDE
        except Exception as e:
            print(f"[WARN] Bad override format: {e}. Using defaults.")

    if ONLY_AFTER_LAST_SAVED and os.path.exists(CHECKPOINT_CSV):
        try:
            newest = None
            for chunk in pd.read_csv(CHECKPOINT_CSV, usecols=["Date Posted"], chunksize=200000):
                s = pd.to_datetime(chunk["Date Posted"], errors="coerce", utc=True)
                mx = s.max()
                if pd.notna(mx):
                    newest = mx if newest is None else max(newest, mx)
            if newest is not None:
                candidate = newest.to_pydatetime() + timedelta(seconds=60)  # skip the last saved
                if candidate > effective_since_dt:
                    effective_since_dt = candidate
        except Exception as e:
            print(f"[WARN] Could not tighten since: {e}")

    since_utc = fmt_utc(effective_since_dt)
    query = build_query(since_utc, effective_until)
    print(f"Using since: {since_utc}  until: {effective_until}")

    session = requests.Session()
    last_ts = 0.0
    backoff = BACKOFF_MIN
    page = 0
    zero_pages = 0

    total_rows = existing
    seen_ids = load_seen_ids()
    cursor = load_cursor()
    buffer_rows = []

    try:
        while total_rows < run_target:
            last_ts = sleep_min_interval(last_ts)
            try:
                tweets, next_cursor, has_next = fetch_page(session, cursor, query)
                page += 1
                backoff = BACKOFF_MIN
            except RateLimitError as rl:
                hint = max(rl.retry_after_sec, 5)
                MIN_INTERVAL_SEC = max(MIN_INTERVAL_SEC, hint + 1.0)
                print(f"[429] Slowing to ~{MIN_INTERVAL_SEC:.1f}s/call. Sleeping {hint}s…")
                time.sleep(hint + random.uniform(0.2, 0.8))
                continue
            except CreditsError as ce:
                print(f"[402] Credits exhausted: {ce}. Saving progress and exiting loop.")
                if buffer_rows:
                    mode = "a" if os.path.exists(CHECKPOINT_CSV) else "w"
                    save_checkpoint(buffer_rows, CHECKPOINT_CSV, mode=mode)
                    persist_seen_ids(seen_ids)
                    save_cursor(cursor)
                    print(f"Final checkpoint | +{len(buffer_rows)} rows.")
                    buffer_rows = []
                break
            except Exception as e:
                print(f"[WARN] {e}. Backoff {backoff:.1f}s…")
                time.sleep(backoff)
                backoff = min(backoff * 2, BACKOFF_MAX)
                continue

            if not tweets:
                print("No tweets returned; stopping.")
                break

            added = 0
            for t in tweets:
                tid = str(t.get("id", "")).strip()
                if not tid:
                    url = t.get("url") or ""
                    if "/status/" in url:
                        tid = url.rsplit("/status/", 1)[-1].split("?", 1)[0]
                if not tid or tid in seen_ids:
                    continue

                # PK post-filter (default ON)
                if STRICT_PAK_FILTER and not looks_pakistan_only(t):
                    continue

                seen_ids.add(tid)
                buffer_rows.append(row_from_tweet(t))
                added += 1
                if total_rows + added >= run_target:
                    break

            total_rows += added
            zero_pages = zero_pages + 1 if added == 0 else 0
            print(f"Page {page}: +{added} | total {total_rows}/{run_target} | empty-streak={zero_pages} | next={'yes' if has_next else 'no'}")

            if len(buffer_rows) >= SAVE_CHECKPOINT_EVERY or total_rows >= run_target:
                mode = "a" if os.path.exists(CHECKPOINT_CSV) else "w"
                save_checkpoint(buffer_rows, CHECKPOINT_CSV, mode=mode)
                persist_seen_ids(seen_ids)
                save_cursor(next_cursor)
                print(f"Checkpoint saved (batch={len(buffer_rows)}).")
                buffer_rows = []

            # Early exit if we're clearly paging through all duplicates
            if zero_pages >= STOP_AFTER_ZERO_PAGES:
                save_cursor(next_cursor)
                print(f"No new tweets for {zero_pages} consecutive pages → stopping to save credits.")
                break

            if total_rows >= run_target:
                print("Hit run target for this run. Stopping.")
                break

            if has_next:
                cursor = next_cursor
                save_cursor(cursor)
            else:
                print("No more pages. Stopping.")
                break

        if buffer_rows:
            mode = "a" if os.path.exists(CHECKPOINT_CSV) else "w"
            save_checkpoint(buffer_rows, CHECKPOINT_CSV, mode=mode)
            persist_seen_ids(seen_ids)
            save_cursor(cursor)
            print(f"Final checkpoint | +{len(buffer_rows)} rows.")

        print("Exporting Excel…")
        dedupe_and_export(CHECKPOINT_CSV, OUTPUT_XLSX)
        print(f"Done → {OUTPUT_XLSX}")

    except KeyboardInterrupt:
        if buffer_rows:
            mode = "a" if os.path.exists(CHECKPOINT_CSV) else "w"
            save_checkpoint(buffer_rows, CHECKPOINT_CSV, mode=mode)
            persist_seen_ids(seen_ids)
            save_cursor(cursor)
            print(f"\nInterrupted. Saved partial batch ({len(buffer_rows)} rows).")
        raise
    finally:
        session.close()

if __name__ == "__main__":
    main()
