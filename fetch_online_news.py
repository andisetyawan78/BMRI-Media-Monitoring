"""
fetch_online_news.py — Ambil berita Bank Mandiri & BMRI dari berbagai sumber online.

Sumber yang dicari:
  - Google News RSS        (portal berita, blog, forum)
  - Portal berita langsung (Kompas, Detik, Bisnis, Kontan, CNBC Indonesia, Tribun)
  - YouTube Data API       (butuh YOUTUBE_API_KEY)
  - Reddit                 (gratis, JSON API publik)
  - Twitter/X              (butuh X_BEARER_TOKEN — opsional, berbayar)

Cara pakai:
  pip install feedparser requests beautifulsoup4
  python fetch_online_news.py

Env vars opsional:
  YOUTUBE_API_KEY   — untuk data YouTube
  X_BEARER_TOKEN    — untuk Twitter/X (butuh akun developer)
  ANTHROPIC_API_KEY — untuk analisis sentimen (sama seperti main.py)
"""

import os, sys, json, datetime, time, re
import urllib.parse, urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from main import analyze_sentiment

try:
    import feedparser
except ImportError:
    print("Install dulu: pip install feedparser")
    sys.exit(1)

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("Install dulu: pip install requests beautifulsoup4")
    sys.exit(1)

# ─── Config ──────────────────────────────────────────────────────────────────

KEYWORDS       = [
    "Bank Mandiri", "BMRI",
    "Mandiri bank", "PT Bank Mandiri",
]
DOCS_DIR       = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs")
ARTICLES_JSON  = os.path.join(DOCS_DIR, "articles.json")
YOUTUBE_KEY    = os.environ.get("YOUTUBE_API_KEY", "")
X_BEARER       = os.environ.get("X_BEARER_TOKEN", "")
DELAY          = 2        # detik jeda antar request
MAX_AGE_DAYS   = 365      # ambil data maksimal 1 tahun ke belakang

CUTOFF = (datetime.datetime.utcnow() - datetime.timedelta(days=MAX_AGE_DAYS))

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

# ─── Helper ──────────────────────────────────────────────────────────────────

def load_existing():
    if not os.path.exists(ARTICLES_JSON):
        return [], set()
    try:
        with open(ARTICLES_JSON, encoding="utf-8") as f:
            data = json.load(f)
        seen = set(a.get("title","")[:80].lower() for a in data)
        return data, seen
    except Exception:
        return [], set()

def save_articles(all_articles):
    os.makedirs(DOCS_DIR, exist_ok=True)
    cutoff_str = CUTOFF.strftime("%Y-%m-%d")
    all_articles = [a for a in all_articles if a.get("run_date","") >= cutoff_str]
    all_articles.sort(key=lambda x: x.get("run_date",""), reverse=True)
    with open(ARTICLES_JSON, "w", encoding="utf-8") as f:
        json.dump(all_articles, f, ensure_ascii=False, separators=(",",":"))
    dates = sorted(set(a.get("run_date","") for a in all_articles))
    print(f"  ✅ Tersimpan: {len(all_articles)} artikel, {len(dates)} hari "
          f"({dates[0] if dates else '-'} s/d {dates[-1] if dates else '-'})")

def parse_date(date_str):
    """Parse berbagai format tanggal → YYYY-MM-DD atau None."""
    if not date_str:
        return None
    # feedparser struct_time
    if hasattr(date_str, "tm_year"):
        try:
            return datetime.date(*date_str[:3]).strftime("%Y-%m-%d")
        except Exception:
            return None
    date_str = str(date_str).strip()
    for fmt in [
        "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%SZ",
        "%a, %d %b %Y %H:%M:%S %z", "%a, %d %b %Y %H:%M:%S GMT",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
        "%d %b %Y", "%B %d, %Y",
    ]:
        try:
            d = datetime.datetime.strptime(date_str[:len(fmt)+5].strip(), fmt)
            return d.strftime("%Y-%m-%d")
        except Exception:
            pass
    # regex fallback
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", date_str)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return None

def is_recent(date_str):
    if not date_str:
        return False
    try:
        d = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        return d >= CUTOFF
    except Exception:
        return False

def clean_text(text):
    if not text:
        return ""
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:300]

def make_article(title, source, link, pub_date, media_type="News", reason=""):
    return {
        "run_date":   pub_date,
        "title":      clean_text(title),
        "source":     source,
        "sentiment":  "",   # diisi nanti oleh analyze_sentiment
        "score":      0,
        "reason":     reason,
        "media_type": media_type,
        "link":       link or "",
    }

# ─── 1. Google News RSS ───────────────────────────────────────────────────────

def fetch_google_news_rss(keyword, before=None, after=None):
    """
    Ambil artikel dari Google News RSS untuk keyword tertentu.
    Google News RSS bisa difilter dengan rentang tanggal via parameter `when`.
    """
    items = []
    # Google News RSS (ceid=ID:id → bahasa Indonesia)
    # Tambah before/after jika ada (format: YYYY-MM-DD)
    q = urllib.parse.quote(keyword)
    url = f"https://news.google.com/rss/search?q={q}&hl=id&gl=ID&ceid=ID:id"
    if after:
        # Google News RSS tidak support date filter langsung di URL,
        # tapi bisa pakai operator: keyword after:YYYY-MM-DD
        q2 = urllib.parse.quote(f"{keyword} after:{after}")
        url = f"https://news.google.com/rss/search?q={q2}&hl=id&gl=ID&ceid=ID:id"

    try:
        feed = feedparser.parse(url)
        for entry in feed.entries:
            pub = parse_date(entry.get("published_parsed") or entry.get("published",""))
            if not pub or not is_recent(pub):
                continue
            title = entry.get("title","")
            link  = entry.get("link","")
            src   = entry.get("source",{}).get("title","") or _domain(link)
            items.append(make_article(title, src, link, pub, "News"))
    except Exception as e:
        print(f"    ⚠ Google News RSS '{keyword}': {e}")
    return items

def _domain(url):
    try:
        return urllib.parse.urlparse(url).netloc.replace("www.","")
    except Exception:
        return ""

# ─── 2. Portal Berita Langsung (RSS) ─────────────────────────────────────────

PORTAL_RSS = [
    # (nama, rss_url, media_type)
    ("Kompas.com",       "https://rss.kompas.com/money/read/xml/indeks/1/",       "News"),
    ("Detik Finance",    "https://finance.detik.com/rss",                          "News"),
    ("Bisnis Indonesia", "https://feeds.bisnis.com/bisnis/rss/finansial/keuangan", "News"),
    ("Kontan.co.id",     "https://www.kontan.co.id/rss/news.rss",                  "News"),
    ("CNBC Indonesia",   "https://www.cnbcindonesia.com/rss",                      "News"),
    ("Tribun Bisnis",    "https://www.tribunnews.com/rss/bisnis",                  "News"),
    ("IDX Channel",      "https://www.idxchannel.com/feed",                        "News"),
    ("Investor Daily",   "https://investor.id/feed",                               "News"),
    ("Tempo Money",      "https://rss.tempo.co/bisnis",                            "News"),
    ("Republika",        "https://rss.republika.co.id/rss/ekonomi/keuangan",       "News"),
]

def fetch_portal_rss(keywords):
    items = []
    kw_lower = [k.lower() for k in keywords]
    for name, url, mtype in PORTAL_RSS:
        try:
            feed = feedparser.parse(url)
            count = 0
            for entry in feed.entries:
                title = entry.get("title","")
                summary = entry.get("summary","")
                combined = (title + " " + summary).lower()
                if not any(kw in combined for kw in kw_lower):
                    continue
                pub = parse_date(entry.get("published_parsed") or entry.get("published",""))
                if not pub or not is_recent(pub):
                    continue
                link = entry.get("link","")
                items.append(make_article(title, name, link, pub, mtype))
                count += 1
            if count:
                print(f"    {name}: {count} artikel")
        except Exception as e:
            print(f"    ⚠ {name}: {e}")
        time.sleep(0.5)
    return items

# ─── 3. Blog & Forum via Google News RSS ─────────────────────────────────────

BLOG_FORUM_KEYWORDS = [
    "Bank Mandiri review",
    "BMRI saham forum",
    "Mandiri tabungan pengalaman",
    "Bank Mandiri Kaskus",
    "BMRI investasi blog",
    "Bank Mandiri kasus",
    "Bank Mandiri OJK",
    "Bank Mandiri tersangka",
    "Bank Mandiri gugat",
    "Bank Mandiri fraud",
    "Bank Mandiri korupsi",
    "Toba Surimi Bank Mandiri",
    "cek palsu Bank Mandiri",
    "kredit Bank Mandiri bermasalah",
    "Bank Mandiri dilaporkan",
    "BMRI kasus hukum",
]

def fetch_blogs_forums(keywords):
    """Cari artikel blog & forum via Google News RSS dengan keyword spesifik."""
    items = []
    # Cari di Google News (termasuk blog & forum yang terindeks)
    for kw in keywords:
        q = urllib.parse.quote(kw)
        url = f"https://news.google.com/rss/search?q={q}&hl=id&gl=ID&ceid=ID:id"
        try:
            feed = feedparser.parse(url)
            for entry in feed.entries:
                pub = parse_date(entry.get("published_parsed") or entry.get("published",""))
                if not pub or not is_recent(pub):
                    continue
                title = entry.get("title","")
                link  = entry.get("link","")
                src   = entry.get("source",{}).get("title","") or _domain(link)
                # Tentukan apakah blog/forum berdasarkan domain
                domain = _domain(link).lower()
                if any(x in domain for x in ["kaskus","reddit","detikforum",
                                              "blogspot","wordpress","medium",
                                              "kompasiana","seputarforex"]):
                    mtype = "Blog"
                else:
                    mtype = "News"
                items.append(make_article(title, src, link, pub, mtype))
        except Exception as e:
            print(f"    ⚠ Blog/Forum '{kw}': {e}")
        time.sleep(1)
    return items

# ─── 4. Reddit ────────────────────────────────────────────────────────────────

REDDIT_SUBREDDITS = ["indonesia", "investasiid", "finansial", "saham"]

def fetch_reddit(keywords):
    items = []
    kw_lower = [k.lower() for k in keywords]
    for sub in REDDIT_SUBREDDITS:
        for kw in keywords:
            url = (f"https://www.reddit.com/r/{sub}/search.json"
                   f"?q={urllib.parse.quote(kw)}&sort=new&t=year&limit=100")
            try:
                resp = requests.get(url, headers=HEADERS, timeout=10)
                if resp.status_code != 200:
                    continue
                data = resp.json()
                posts = data.get("data",{}).get("children",[])
                count = 0
                for post in posts:
                    p = post.get("data",{})
                    title = p.get("title","")
                    if not any(kw in title.lower() for kw in kw_lower):
                        continue
                    ts = p.get("created_utc", 0)
                    if not ts:
                        continue
                    pub = datetime.datetime.utcfromtimestamp(ts).strftime("%Y-%m-%d")
                    if not is_recent(pub):
                        continue
                    link = f"https://reddit.com{p.get('permalink','')}"
                    items.append(make_article(title, f"Reddit r/{sub}", link, pub, "Blog"))
                    count += 1
                if count:
                    print(f"    Reddit r/{sub} '{kw}': {count} post")
            except Exception as e:
                print(f"    ⚠ Reddit r/{sub}: {e}")
            time.sleep(1)
    return items

# ─── 5. YouTube ───────────────────────────────────────────────────────────────

def fetch_youtube(keywords):
    if not YOUTUBE_KEY:
        print("    ⚠ YOUTUBE_API_KEY tidak di-set — skip YouTube")
        print("      Set dengan: $env:YOUTUBE_API_KEY = 'AIza...'")
        return []
    items = []
    one_year_ago = (datetime.datetime.utcnow() - datetime.timedelta(days=365))
    published_after = one_year_ago.strftime("%Y-%m-%dT%H:%M:%SZ")

    for kw in keywords:
        url = (
            "https://www.googleapis.com/youtube/v3/search"
            f"?part=snippet&q={urllib.parse.quote(kw)}"
            f"&type=video&order=date&maxResults=50"
            f"&publishedAfter={published_after}"
            f"&relevanceLanguage=id"
            f"&key={YOUTUBE_KEY}"
        )
        try:
            resp = requests.get(url, timeout=15)
            if resp.status_code != 200:
                print(f"    ⚠ YouTube API error: {resp.status_code}")
                continue
            data = resp.json()
            count = 0
            for item in data.get("items", []):
                snippet = item.get("snippet", {})
                title   = snippet.get("title", "")
                channel = snippet.get("channelTitle", "")
                vid_id  = item.get("id", {}).get("videoId", "")
                pub_raw = snippet.get("publishedAt", "")
                pub     = parse_date(pub_raw)
                if not pub or not is_recent(pub):
                    continue
                link = f"https://www.youtube.com/watch?v={vid_id}"
                items.append(make_article(title, f"YouTube: {channel}", link, pub, "YouTube"))
                count += 1
            if count:
                print(f"    YouTube '{kw}': {count} video")
        except Exception as e:
            print(f"    ⚠ YouTube '{kw}': {e}")
        time.sleep(1)

    # Juga cari channel resmi Bank Mandiri
    channel_url = (
        "https://www.googleapis.com/youtube/v3/search"
        "?part=snippet&channelId=UCLjNRwnRBBq3w5lz8T7Cp5Q"
        "&type=video&order=date&maxResults=50"
        f"&publishedAfter={published_after}"
        f"&key={YOUTUBE_KEY}"
    )
    try:
        resp = requests.get(channel_url, timeout=15)
        if resp.status_code == 200:
            data = resp.json()
            count = 0
            for item in data.get("items", []):
                snippet = item.get("snippet", {})
                title   = snippet.get("title", "")
                vid_id  = item.get("id", {}).get("videoId", "")
                pub     = parse_date(snippet.get("publishedAt",""))
                if not pub or not is_recent(pub):
                    continue
                link = f"https://www.youtube.com/watch?v={vid_id}"
                items.append(make_article(title, "YouTube: Bank Mandiri", link, pub, "YouTube"))
                count += 1
            if count:
                print(f"    YouTube (channel resmi): {count} video")
    except Exception as e:
        print(f"    ⚠ YouTube channel resmi: {e}")

    return items

# ─── 6. Twitter/X ────────────────────────────────────────────────────────────

def fetch_twitter(keywords):
    if not X_BEARER:
        print("    ⚠ X_BEARER_TOKEN tidak di-set — skip Twitter/X")
        print("      Butuh akun developer Twitter: https://developer.x.com")
        print("      Set dengan: $env:X_BEARER_TOKEN = 'AAAAAA...'")
        return []
    items = []
    headers_x = {**HEADERS, "Authorization": f"Bearer {X_BEARER}"}
    one_year_ago = (datetime.datetime.utcnow() - datetime.timedelta(days=7))
    # Twitter Basic API hanya 7 hari ke belakang; Academic perlu apply khusus
    start_time = one_year_ago.strftime("%Y-%m-%dT%H:%M:%SZ")

    for kw in keywords:
        url = (
            "https://api.twitter.com/2/tweets/search/recent"
            f"?query={urllib.parse.quote(kw + ' lang:id -is:retweet')}"
            "&tweet.fields=created_at,author_id,text"
            "&max_results=100"
            f"&start_time={start_time}"
        )
        try:
            resp = requests.get(url, headers=headers_x, timeout=15)
            if resp.status_code == 429:
                print("    ⚠ Twitter rate limit — tunggu 15 menit")
                time.sleep(900)
                continue
            if resp.status_code != 200:
                print(f"    ⚠ Twitter API {resp.status_code}: {resp.text[:200]}")
                continue
            data = resp.json()
            tweets = data.get("data", [])
            count = 0
            for tw in tweets:
                pub = parse_date(tw.get("created_at",""))
                if not pub or not is_recent(pub):
                    continue
                title = tw.get("text","")[:200]
                tid   = tw.get("id","")
                link  = f"https://twitter.com/i/web/status/{tid}"
                items.append(make_article(title, "Twitter/X", link, pub, "Twitter"))
                count += 1
            if count:
                print(f"    Twitter/X '{kw}': {count} tweet")
        except Exception as e:
            print(f"    ⚠ Twitter/X '{kw}': {e}")
        time.sleep(2)
    return items

# ─── 7. Instagram & TikTok & Facebook ────────────────────────────────────────

def fetch_instagram_tiktok_facebook():
    """
    Instagram, TikTok, Facebook tidak memiliki API publik gratis untuk search.
    Alternatif:
    - Instagram: butuh Instagram Graph API + akun bisnis + review Meta
    - TikTok: TikTok Research API (apply dulu, hanya untuk researcher)
    - Facebook: CrowdTangle (discontinued) atau Meta Content Library (terbatas)

    Saat ini: ambil dari akun resmi Bank Mandiri via RSS jika tersedia,
    atau monitor hashtag via agregator pihak ketiga.
    """
    items = []
    # Akun resmi Bank Mandiri di YouTube (sebagai proxy social media)
    # Facebook page Bank Mandiri tidak ada RSS publik
    # TikTok @bankmandiri tidak ada API publik
    print("    ℹ️  Instagram/TikTok/Facebook: tidak ada API publik gratis")
    print("       Untuk social media, gunakan tool seperti:")
    print("       - Talkwalker (sudah terhubung via Gmail ✅)")
    print("       - Brand24, Sprout Social, atau Mention.com")
    return items

# ─── 8. Google News per bulan (backfill historis) ─────────────────────────────

MONTH_KEYWORDS = [
    # Keyword utama
    "Bank Mandiri", "BMRI",
    # Saham & keuangan
    "saham BMRI", "kredit Mandiri", "Mandiri BMRI",
    "dividen BMRI", "laba Bank Mandiri", "NPL Bank Mandiri",
    # Produk & layanan
    "Livin Mandiri", "Kopra Mandiri", "KPR Mandiri",
    "Mandiri Syariah", "BSI Bank Mandiri",
    # Kasus & hukum — penting agar tidak terlewat
    "Bank Mandiri OJK", "Bank Mandiri kasus",
    "Bank Mandiri dilaporkan", "Bank Mandiri gugat",
    "Toba Surimi Bank Mandiri", "cek palsu Bank Mandiri",
    "Bank Mandiri kredit bermasalah",
    # Direksi & manajemen
    "Darmawan Junaidi Bank Mandiri",
    "Direktur Bank Mandiri",
]

def fetch_google_news_monthly(keywords):
    """
    Cari berita per bulan untuk 12 bulan terakhir.
    Google News RSS mendukung operator after: dan before: di query.
    """
    items = []
    today = datetime.date.today()
    for months_ago in range(0, 13):
        # Hitung rentang bulan
        first_of_month = (today.replace(day=1) -
                          datetime.timedelta(days=months_ago * 28)).replace(day=1)
        # Last day of that month
        if first_of_month.month == 12:
            last_of_month = first_of_month.replace(year=first_of_month.year+1, month=1, day=1) - datetime.timedelta(days=1)
        else:
            last_of_month = first_of_month.replace(month=first_of_month.month+1, day=1) - datetime.timedelta(days=1)

        if first_of_month < CUTOFF.date():
            break

        after_str  = first_of_month.strftime("%Y-%m-%d")
        before_str = last_of_month.strftime("%Y-%m-%d")
        month_label = first_of_month.strftime("%B %Y")

        month_items = []
        for kw in keywords:
            q = urllib.parse.quote(f'"{kw}" after:{after_str} before:{before_str}')
            url = f"https://news.google.com/rss/search?q={q}&hl=id&gl=ID&ceid=ID:id"
            try:
                feed = feedparser.parse(url)
                for entry in feed.entries:
                    pub = parse_date(entry.get("published_parsed") or entry.get("published",""))
                    if not pub or not is_recent(pub):
                        continue
                    title = entry.get("title","")
                    link  = entry.get("link","")
                    src   = entry.get("source",{}).get("title","") or _domain(link)
                    month_items.append(make_article(title, src, link, pub, "News"))
            except Exception as e:
                print(f"    ⚠ Google News '{kw}' {month_label}: {e}")
            time.sleep(1)

        # Deduplikasi dalam bulan ini
        seen_titles = set()
        unique = []
        for a in month_items:
            key = a["title"][:60].lower()
            if key not in seen_titles and key:
                seen_titles.add(key)
                unique.append(a)

        if unique:
            print(f"    {month_label}: {len(unique)} artikel")
        items.extend(unique)
        time.sleep(1)

    return items

# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "="*60)
    print("  FETCH ONLINE NEWS — Bank Mandiri & BMRI")
    print("="*60)
    print(f"  Periode: {CUTOFF.strftime('%d %b %Y')} s/d hari ini")
    print(f"  Keywords: {KEYWORDS}\n")

    all_articles, seen_titles = load_existing()
    print(f"[0] Data existing: {len(all_articles)} artikel dari {len(set(a.get('run_date','') for a in all_articles))} hari\n")

    new_items = []

    # ── Google News per bulan (backfill historis) ──
    print("[1] Google News RSS — per bulan (12 bulan terakhir)...")
    new_items += fetch_google_news_monthly(MONTH_KEYWORDS)
    print()

    # ── Portal berita via RSS ──
    print("[2] Portal berita Indonesia via RSS...")
    new_items += fetch_portal_rss(KEYWORDS)
    print()

    # ── Blog & Forum via Google News ──
    print("[3] Blog & Forum via Google News...")
    new_items += fetch_blogs_forums(BLOG_FORUM_KEYWORDS)
    print()

    # ── Reddit ──
    print("[4] Reddit (r/indonesia, r/investasiid, r/saham)...")
    new_items += fetch_reddit(KEYWORDS)
    print()

    # ── YouTube ──
    print("[5] YouTube...")
    new_items += fetch_youtube(KEYWORDS)
    print()

    # ── Twitter/X ──
    print("[6] Twitter/X...")
    new_items += fetch_twitter(KEYWORDS)
    print()

    # ── Instagram / TikTok / Facebook ──
    print("[7] Instagram / TikTok / Facebook...")
    fetch_instagram_tiktok_facebook()
    print()

    # ── Deduplikasi global ──
    print("[8] Deduplikasi & filter baru...")
    added = 0
    for a in new_items:
        key = a.get("title","")[:80].lower()
        if key and key not in seen_titles and a.get("run_date") and is_recent(a.get("run_date","")):
            seen_titles.add(key)
            all_articles.append(a)
            added += 1
    print(f"    Artikel baru (belum ada): {added}")
    print(f"    Total sebelum sentimen : {len(all_articles)}")

    if added == 0:
        print("\n  Tidak ada artikel baru yang perlu diproses.")
        return

    # ── Analisis sentimen untuk artikel baru ──
    print(f"\n[9] Analisis sentimen {added} artikel baru...")
    to_analyze = [a for a in all_articles if not a.get("sentiment")]
    batch_size = 20
    for i in range(0, len(to_analyze), batch_size):
        batch = to_analyze[i:i+batch_size]
        try:
            analyzed = analyze_sentiment(batch)
            for orig, result in zip(batch, analyzed):
                orig["sentiment"] = result.get("sentiment", "netral")
                orig["score"]     = result.get("score", 5)
                orig["reason"]    = result.get("reason", "")
            print(f"    Batch {i//batch_size+1}/{(len(to_analyze)+batch_size-1)//batch_size} selesai")
        except Exception as e:
            print(f"    ⚠ Sentimen batch {i//batch_size+1}: {e}")
        time.sleep(2)

    # ── Simpan ──
    print("\n[10] Menyimpan...")
    save_articles(all_articles)

    # ── Ringkasan ──
    dates = sorted(set(a.get("run_date","") for a in all_articles if a.get("run_date","")))
    print(f"\n{'='*60}")
    print(f"  ✅ SELESAI!")
    print(f"  Artikel baru  : {added}")
    print(f"  Total artikel : {len(all_articles)}")
    print(f"  Total hari    : {len(dates)}")
    if dates:
        print(f"  Rentang       : {dates[0]} s/d {dates[-1]}")

    missing_days = 365 - len(dates)
    if missing_days > 0:
        print(f"\n  ℹ️  Masih ada ~{missing_days} hari tanpa data")
        print(f"     (hari libur/weekend biasanya tidak ada berita — ini normal)")
    print("="*60)

    print("\nLangkah selanjutnya:")
    print("  Upload docs/articles.json ke GitHub → dashboard otomatis update")


if __name__ == "__main__":
    main()
