"""
Bank Mandiri Media Monitoring
==============================
Otomatis setiap hari jam 07:00 WIB via GitHub Actions.
Alur: Gmail (Talkwalker alerts) → Parse → Analisis AI → Grafik → Telegram + Email

Diperlukan environment variables (simpan sebagai GitHub Secrets):
  GMAIL_CREDENTIALS_JSON   : isi file credentials.json Google OAuth2 (base64-encoded)
  GMAIL_TOKEN_JSON         : isi file token.json Google OAuth2 (base64-encoded)
  ANTHROPIC_API_KEY        : API key Anthropic Claude
  TELEGRAM_BOT_TOKEN       : token bot Telegram
  TELEGRAM_CHAT_ID         : chat_id tujuan Telegram
  EMAIL_SENDER             : alamat Gmail pengirim (misal: sivbmri@gmail.com)
  EMAIL_APP_PASSWORD       : Gmail App Password (bukan password biasa)
  EMAIL_RECIPIENT          : sivbmri@gmail.com
"""

import os
import base64
import json
import re
import smtplib
import tempfile
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from html.parser import HTMLParser

import requests
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import anthropic

# ─────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────
TALKWALKER_SENDER = "alerts@talkwalker.com"
MONITOR_KEYWORDS  = ["Bank Mandiri", "BMRI"]
EMAIL_RECIPIENT   = os.environ.get("EMAIL_RECIPIENT", "sivbmri@gmail.com")
EMAIL_SENDER      = os.environ.get("EMAIL_SENDER", "sivbmri@gmail.com")
EMAIL_APP_PASS    = os.environ.get("EMAIL_APP_PASSWORD", "")
TG_TOKEN          = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TG_CHAT_ID        = os.environ.get("TELEGRAM_CHAT_ID", "")
ANTHROPIC_KEY     = os.environ.get("ANTHROPIC_API_KEY", "")
GMAIL_SCOPES      = ["https://www.googleapis.com/auth/gmail.readonly"]


# ─────────────────────────────────────────────
# 1. AUTENTIKASI GMAIL
# ─────────────────────────────────────────────
def get_gmail_service():
    """Buat Gmail API service dari credentials yang tersimpan di env vars."""
    creds = None

    # Decode credentials dari env (base64)
    creds_b64   = os.environ.get("GMAIL_CREDENTIALS_JSON", "")
    token_b64   = os.environ.get("GMAIL_TOKEN_JSON", "")

    def safe_b64decode(s):
        """Decode base64 dengan auto-fix padding dan strip whitespace."""
        s = s.strip().replace("\n", "").replace("\r", "").replace(" ", "")
        # Tambah padding jika kurang
        missing = len(s) % 4
        if missing:
            s += "=" * (4 - missing)
        return base64.b64decode(s)

    with tempfile.TemporaryDirectory() as tmpdir:
        creds_path  = os.path.join(tmpdir, "credentials.json")
        token_path  = os.path.join(tmpdir, "token.json")

        if creds_b64:
            with open(creds_path, "w") as f:
                f.write(safe_b64decode(creds_b64).decode())

        if token_b64:
            with open(token_path, "w") as f:
                f.write(safe_b64decode(token_b64).decode())
            creds = Credentials.from_authorized_user_file(token_path, GMAIL_SCOPES)

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())

        service = build("gmail", "v1", credentials=creds)
    return service


# ─────────────────────────────────────────────
# 2. AMBIL EMAIL TALKWALKER KEMARIN
# ─────────────────────────────────────────────
def get_yesterday_range():
    """Return (after_str, before_str) dalam format YYYY/MM/DD untuk kemarin."""
    today     = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)
    return yesterday.strftime("%Y/%m/%d"), today.strftime("%Y/%m/%d")


def fetch_talkwalker_emails(service):
    """Ambil semua email Talkwalker dari kemarin, return list HTML bodies."""
    after, before = get_yesterday_range()
    query = f"from:{TALKWALKER_SENDER} after:{after} before:{before}"
    print(f"[Gmail] Query: {query}")

    result   = service.users().messages().list(userId="me", q=query, maxResults=50).execute()
    messages = result.get("messages", [])
    print(f"[Gmail] Ditemukan {len(messages)} pesan")

    html_bodies = []
    for msg_meta in messages:
        msg = service.users().messages().get(
            userId="me", id=msg_meta["id"], format="full"
        ).execute()
        html = _extract_html(msg)
        if html:
            html_bodies.append(html)
    return html_bodies


def _extract_html(msg):
    """Ekstrak HTML body dari pesan Gmail."""
    payload = msg.get("payload", {})
    return _walk_parts(payload)


def _walk_parts(part):
    mime = part.get("mimeType", "")
    body = part.get("body", {})
    data = body.get("data", "")

    if mime == "text/html" and data:
        return base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")

    for sub in part.get("parts", []):
        result = _walk_parts(sub)
        if result:
            return result
    return None


# ─────────────────────────────────────────────
# 3. PARSE HTML → DAFTAR ARTIKEL
# ─────────────────────────────────────────────
class _HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.fed = []
        self._skip = False

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self._skip = True

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self._skip = False

    def handle_data(self, d):
        if not self._skip:
            self.fed.append(d)

    def get_data(self):
        return "\n".join(self.fed)


def html_to_text(html):
    s = _HTMLStripper()
    s.feed(html)
    text = s.get_data()
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r" {2,}", " ", text)
    return text.strip()


def parse_articles(html_bodies):
    """
    Parse artikel dari HTML email Talkwalker.
    Strategi: cari baris tanggal/sumber (format Talkwalker), lalu scan mundur
    untuk menemukan judul dan snippet.
    Tidak menggunakan filter keyword karena Talkwalker sudah memfilter untuk Bank Mandiri.
    Return list of dict: {title, snippet, date, source, media_type}
    """
    articles = []

    # Format Talkwalker: "15/06/26 05:34 | Indonesia | katadata.co.id"
    # atau sumber di baris berikutnya: "15/06/26 05:34 | Indonesia |"
    date_pattern = re.compile(
        r"^(\d{2}/\d{2}/\d{2})\s+\d{2}:\d{2}\s*\|\s*([\w][\w\s]*)\|\s*(.*)$"
    )
    # Fallback: format tanpa jam "DD/MM/YY | Country | source"
    date_pattern_nohour = re.compile(
        r"^(\d{2}/\d{2}/\d{2,4})\s*\|\s*([\w][\w\s]*)\|\s*(.*)$"
    )
    # Baris noise yang harus diabaikan saat scan mundur
    NOISE = {
        "Tweet", "Show less", "Tell a Friend", "liking", "following",
        "News", "Blogs", "Twitter", "new results", "Delete Alert",
        "Manage Alerts", "Create Alert", "Unsubscribe", "View in browser",
    }
    NOISE_RE = re.compile(r"^\d+\s+new results?$", re.IGNORECASE)

    for html in html_bodies:
        text = html_to_text(html)
        # Collapse tab/spasi ganda tapi jaga newline
        text = re.sub(r"[ \t]{2,}", " ", text)

        all_lines = text.splitlines()
        lines = [l.strip() for l in all_lines]  # semua baris (termasuk kosong)

        # Deteksi media_type dari 40 baris pertama yang tidak kosong
        media_type = "News"
        non_empty_head = [l for l in lines[:60] if l][:40]
        for line in non_empty_head:
            ll = line.lower()
            if "blogs" in ll:
                media_type = "Blog"
                break
            elif "twitter" in ll:
                media_type = "Twitter"
                break

        # Indeks baris tidak-kosong untuk navigasi
        nonempty_lines = [(i, l) for i, l in enumerate(lines) if l]

        for idx, (line_idx, line) in enumerate(nonempty_lines):
            # Coba cocokkan pola tanggal
            m = date_pattern.match(line) or date_pattern_nohour.match(line)
            if not m:
                continue

            date = m.group(1)
            # group(3) = sumber jika ada di baris yang sama
            inline_source = m.group(3).strip() if m.lastindex >= 3 else ""

            # Jika sumber kosong, cek baris non-kosong berikutnya
            source = inline_source
            if not source and idx + 1 < len(nonempty_lines):
                next_line = nonempty_lines[idx + 1][1]
                # Validasi: terlihat seperti domain (ada titik, tidak terlalu panjang)
                if re.match(r"^[\w\-]+\.[\w\.\-]{2,}$", next_line) and len(next_line) < 60:
                    source = next_line

            if not source:
                continue

            # Scan mundur untuk menemukan judul dan snippet
            title   = ""
            snippet = ""
            candidates = []  # (baris non-kosong sebelum date line)

            for back_idx in range(idx - 1, max(idx - 20, -1), -1):
                prev_line = nonempty_lines[back_idx][1]

                # Hentikan jika ketemu baris tanggal lain (artikel sebelumnya)
                if date_pattern.match(prev_line) or date_pattern_nohour.match(prev_line):
                    break
                # Hentikan jika ketemu noise header
                if prev_line in NOISE or NOISE_RE.match(prev_line):
                    break
                # Lewati baris sangat pendek (1-3 karakter) — biasanya artefak HTML
                if len(prev_line) <= 3:
                    continue

                candidates.insert(0, prev_line)
                if len(candidates) >= 5:
                    break

            # Bersihkan candidates dari noise
            candidates = [
                c for c in candidates
                if c not in NOISE and not NOISE_RE.match(c)
            ]

            if not candidates:
                continue

            # Gabungkan kandidat pendek berturut-turut (kata terpotong akibat HTML tag)
            # Contoh: ["Bank", "Mandiri", "Gelar Program..."] → "Bank Mandiri Gelar Program..."
            title = candidates[0]
            rest  = candidates[1:]
            while len(title.split()) <= 2 and rest:
                title = title + " " + rest[0]
                rest  = rest[1:]
            snippet = " ".join(rest[:2])[:300]

            # Filter judul yang tidak masuk akal
            if len(title) < 8:
                continue
            skip_titles = {"tell a friend", "new results", "delete alert",
                           "manage alerts", "create alert", "unsubscribe"}
            if any(s in title.lower() for s in skip_titles):
                continue

            articles.append({
                "title"      : title,
                "snippet"    : snippet,
                "date"       : date,
                "source"     : source.strip(),
                "media_type" : media_type,
            })

    # Deduplikasi berdasarkan judul
    seen   = set()
    unique = []
    for a in articles:
        key = a["title"][:60].lower()
        if key not in seen:
            seen.add(key)
            unique.append(a)

    print(f"[Parser] Total artikel unik: {len(unique)}")
    return unique


# ─────────────────────────────────────────────
# 4. ANALISIS SENTIMEN DENGAN CLAUDE AI
# ─────────────────────────────────────────────
def _analyze_batch(client, batch):
    """Analisis sentimen untuk satu batch artikel (maks 30)."""
    articles_text = ""
    for i, a in enumerate(batch, 1):
        articles_text += (
            f"{i}. JUDUL: {a['title']}\n"
            f"   SNIPPET: {a['snippet'][:150]}\n"
            f"   SUMBER: {a['source']}\n\n"
        )

    prompt = f"""Kamu adalah analis media monitoring profesional untuk Bank Mandiri (BMRI) Indonesia.

Berikut daftar artikel/berita:

{articles_text}

Tugasmu:
1. Tentukan sentimen setiap artikel: "positif" atau "negatif" terhadap citra/kinerja Bank Mandiri.
2. Berikan skor 1-10 (10 = paling berdampak tinggi dalam kategorinya).
3. Berikan alasan singkat (maks 15 kata).
4. Tandai "irrelevant" jika artikel tidak terkait langsung Bank Mandiri BUMN (spam, iklan, dll).

Output HANYA JSON array ini, tanpa teks lain:
[
  {{"id": 1, "sentiment": "positif", "score": 9, "reason": "Alasan singkat"}},
  {{"id": 2, "sentiment": "negatif", "score": 7, "reason": "Alasan singkat"}},
  {{"id": 3, "sentiment": "irrelevant"}}
]"""

    try:
        response = client.messages.create(
            model="claude-opus-4-8",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = response.content[0].text.strip()

        # Ekstrak JSON array dari response
        json_match = re.search(r"\[.*\]", raw, re.DOTALL)
        if not json_match:
            print(f"[AI] Tidak ada JSON dalam response, pakai default")
            raise ValueError("No JSON found")

        scores     = json.loads(json_match.group())
        score_map  = {item["id"]: item for item in scores}

        result = []
        for i, a in enumerate(batch, 1):
            info = score_map.get(i, {})
            if info.get("sentiment") == "irrelevant":
                continue
            a["sentiment"] = info.get("sentiment", "netral")
            a["score"]     = info.get("score", 5)
            a["reason"]    = info.get("reason", "")
            result.append(a)
        return result

    except Exception as e:
        print(f"[AI] Error batch: {e} — pakai sentimen default")
        for a in batch:
            a.setdefault("sentiment", "netral")
            a.setdefault("score", 5)
            a.setdefault("reason", "")
        return batch


def analyze_sentiment(articles):
    """
    Kirim daftar artikel ke Claude untuk scoring sentimen (diproses per batch).
    Return list dengan field tambahan: sentiment ('positif'/'negatif'), score (1-10), reason
    """
    if not articles:
        return []

    client      = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
    BATCH_SIZE  = 30
    all_results = []
    total_batch = (len(articles) + BATCH_SIZE - 1) // BATCH_SIZE

    print(f"[AI] {len(articles)} artikel → {total_batch} batch @{BATCH_SIZE}")

    for b in range(total_batch):
        start = b * BATCH_SIZE
        batch = articles[start:start + BATCH_SIZE]
        print(f"[AI] Batch {b+1}/{total_batch} ({len(batch)} artikel)...")
        all_results.extend(_analyze_batch(client, batch))

    print(f"[AI] Selesai: {len(all_results)} artikel relevan")
    return all_results


# ─────────────────────────────────────────────
# 5. BUAT GRAFIK SENTIMEN
# ─────────────────────────────────────────────
def generate_chart(articles, output_path):
    """Buat grafik batang horizontal sentimen + skor, simpan ke output_path."""
    negatif = sorted(
        [a for a in articles if a.get("sentiment") == "negatif"],
        key=lambda x: x["score"], reverse=True
    )
    positif = sorted(
        [a for a in articles if a.get("sentiment") == "positif"],
        key=lambda x: x["score"], reverse=True
    )

    # Batasi tampilan ke top 10 masing-masing
    negatif = negatif[:10]
    positif = positif[:10]

    items  = negatif + positif
    labels = []
    scores = []
    colors = []

    for a in negatif:
        title = a["title"][:50] + "…" if len(a["title"]) > 50 else a["title"]
        labels.append(f"[{a['source'][:15]}] {title}")
        scores.append(-a["score"])   # negatif ke kiri
        colors.append("#E24B4A")

    for a in positif:
        title = a["title"][:50] + "…" if len(a["title"]) > 50 else a["title"]
        labels.append(f"[{a['source'][:15]}] {title}")
        scores.append(a["score"])    # positif ke kanan
        colors.append("#639922")

    n      = len(labels)
    height = max(6, n * 0.45 + 2)

    fig, ax = plt.subplots(figsize=(14, height))
    fig.patch.set_facecolor("#FAFAFA")
    ax.set_facecolor("#FAFAFA")

    bars = ax.barh(range(n), scores, color=colors, height=0.65, zorder=3)

    # Label skor di ujung bar
    for i, (bar, score) in enumerate(zip(bars, scores)):
        val   = abs(score)
        xpos  = bar.get_width() + (0.15 if score >= 0 else -0.15)
        ha    = "left" if score >= 0 else "right"
        color = "#3B6D11" if score >= 0 else "#A32D2D"
        ax.text(xpos, i, str(val), va="center", ha=ha, fontsize=8,
                color=color, fontweight="bold")

    ax.set_yticks(range(n))
    ax.set_yticklabels(labels, fontsize=7.5)
    ax.axvline(0, color="#888", linewidth=0.8, zorder=2)
    ax.set_xlim(-11, 11)
    ax.set_xticks([-10, -8, -6, -4, -2, 0, 2, 4, 6, 8, 10])
    ax.set_xticklabels(
        ["10","8","6","4","2","0","2","4","6","8","10"], fontsize=8
    )
    ax.grid(axis="x", color="#DDDDDD", linewidth=0.5, zorder=1)
    ax.spines[["top","right","left","bottom"]].set_visible(False)
    ax.tick_params(axis="y", length=0)

    yesterday = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%d %B %Y")
    ax.set_title(
        f"Bank Mandiri (BMRI) — Analisis Sentimen Media\n{yesterday}",
        fontsize=12, fontweight="bold", pad=14, color="#222222"
    )

    neg_patch = mpatches.Patch(color="#E24B4A", label=f"Negatif ({len(negatif)})")
    pos_patch = mpatches.Patch(color="#639922", label=f"Positif ({len(positif)})")
    ax.legend(handles=[neg_patch, pos_patch], loc="lower right",
              fontsize=9, framealpha=0.7)

    ax.text(0.02, -0.04,
            "Skor 1–10 | Sumber: Talkwalker Alerts | Analisis: Claude AI",
            transform=ax.transAxes, fontsize=7, color="#888888")

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.close()
    print(f"[Chart] Grafik disimpan: {output_path}")


# ─────────────────────────────────────────────
# 6. FORMAT KONTEN TELEGRAM & EMAIL
# ─────────────────────────────────────────────
def format_telegram_negative(articles, date_str):
    negatif = sorted(
        [a for a in articles if a.get("sentiment") == "negatif"],
        key=lambda x: x["score"], reverse=True
    )
    if not negatif:
        return "✅ Tidak ada berita negatif hari ini."

    lines = [
        f"🔴 *MEDIA MONITORING BANK MANDIRI — NEGATIF*",
        f"📅 {date_str}",
        f"",
        f"Total: *{len(negatif)} berita negatif*",
        f"─────────────────────",
    ]
    for i, a in enumerate(negatif, 1):
        lines += [
            f"",
            f"*{i}. Skor: {a['score']}/10*",
            f"📰 {a['title'][:120]}",
            f"🔍 _{a['reason']}_",
            f"📡 `{a['source']}`  |  {a['date']}",
        ]
    lines += ["", "─────────────────────",
              "_Analisis otomatis oleh Claude AI_"]
    return "\n".join(lines)


def format_telegram_positive(articles, date_str):
    positif = sorted(
        [a for a in articles if a.get("sentiment") == "positif"],
        key=lambda x: x["score"], reverse=True
    )
    if not positif:
        return "ℹ️ Tidak ada berita positif signifikan hari ini."

    lines = [
        f"🟢 *MEDIA MONITORING BANK MANDIRI — POSITIF*",
        f"📅 {date_str}",
        f"",
        f"Total: *{len(positif)} berita positif*",
        f"─────────────────────",
    ]
    for i, a in enumerate(positif, 1):
        lines += [
            f"",
            f"*{i}. Skor: {a['score']}/10*",
            f"📰 {a['title'][:120]}",
            f"💡 _{a['reason']}_",
            f"📡 `{a['source']}`  |  {a['date']}",
        ]
    lines += ["", "─────────────────────",
              "_Analisis otomatis oleh Claude AI_"]
    return "\n".join(lines)


def format_summary(articles, date_str):
    positif = [a for a in articles if a.get("sentiment") == "positif"]
    negatif = [a for a in articles if a.get("sentiment") == "negatif"]
    total   = len(articles)
    pct_pos = round(len(positif) / total * 100) if total else 0

    avg_pos = round(sum(a["score"] for a in positif) / len(positif), 1) if positif else 0
    avg_neg = round(sum(a["score"] for a in negatif) / len(negatif), 1) if negatif else 0
    top_neg = negatif[0] if negatif else None
    top_pos = positif[0] if positif else None

    sentiment_icon = "🟢" if pct_pos >= 60 else ("🟡" if pct_pos >= 40 else "🔴")

    lines = [
        f"📊 *RINGKASAN MEDIA MONITORING BANK MANDIRI*",
        f"📅 {date_str}",
        f"",
        f"{sentiment_icon} Sentimen Keseluruhan: *{'POSITIF' if pct_pos >= 60 else 'NEGATIF' if pct_pos < 40 else 'NETRAL'}* ({pct_pos}%)",
        f"",
        f"📈 Total Berita Relevan : {total}",
        f"🟢 Positif  : {len(positif)} berita  |  Avg Skor: {avg_pos}",
        f"🔴 Negatif  : {len(negatif)} berita  |  Avg Skor: {avg_neg}",
    ]

    if top_neg:
        lines += [
            f"",
            f"⚠️ *Berita Negatif Utama:*",
            f"_{top_neg['title'][:100]}_",
            f"Skor {top_neg['score']}/10 — {top_neg['source']}",
        ]
    if top_pos:
        lines += [
            f"",
            f"✨ *Berita Positif Utama:*",
            f"_{top_pos['title'][:100]}_",
            f"Skor {top_pos['score']}/10 — {top_pos['source']}",
        ]

    lines += ["", "_📬 Laporan dikirim ke Email & Telegram_",
              "_Analisis otomatis oleh Claude AI_"]
    return "\n".join(lines)


# ─────────────────────────────────────────────
# 7. KIRIM KE TELEGRAM
# ─────────────────────────────────────────────
def send_telegram_text(text):
    url  = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    data = {
        "chat_id"    : TG_CHAT_ID,
        "text"       : text,
        "parse_mode" : "Markdown",
    }
    r = requests.post(url, data=data, timeout=30)
    if not r.ok:
        print(f"[Telegram] Error text: {r.text}")
    else:
        print("[Telegram] Teks terkirim.")


def send_telegram_photo(image_path, caption=""):
    url  = f"https://api.telegram.org/bot{TG_TOKEN}/sendPhoto"
    with open(image_path, "rb") as f:
        r = requests.post(
            url,
            data={"chat_id": TG_CHAT_ID, "caption": caption, "parse_mode": "Markdown"},
            files={"photo": f},
            timeout=60,
        )
    if not r.ok:
        print(f"[Telegram] Error photo: {r.text}")
    else:
        print("[Telegram] Grafik terkirim.")


# ─────────────────────────────────────────────
# 8. KIRIM EMAIL (SMTP GMAIL)
# ─────────────────────────────────────────────
def build_html_email(articles, date_str, chart_path):
    """Buat HTML email lengkap dengan tabel berita dan grafik embedded."""
    positif = sorted(
        [a for a in articles if a.get("sentiment") == "positif"],
        key=lambda x: x["score"], reverse=True
    )
    negatif = sorted(
        [a for a in articles if a.get("sentiment") == "negatif"],
        key=lambda x: x["score"], reverse=True
    )
    total   = len(articles)
    pct_pos = round(len(positif) / total * 100) if total else 0

    def rows(items, color, badge_color):
        html = ""
        for i, a in enumerate(items, 1):
            bar_width = int(a["score"] * 10)
            html += f"""
            <tr style="border-bottom:1px solid #f0f0f0">
              <td style="padding:10px 8px;text-align:center;font-weight:700;color:{color};font-size:15px">{a['score']}</td>
              <td style="padding:10px 8px">
                <div style="font-weight:600;color:#1a1a1a;font-size:13px;line-height:1.4">{a['title']}</div>
                <div style="color:#666;font-size:11px;margin-top:4px">{a.get('reason','')}</div>
                <div style="background:#eee;border-radius:3px;height:4px;margin-top:6px;overflow:hidden">
                  <div style="background:{color};width:{bar_width}%;height:4px;border-radius:3px"></div>
                </div>
              </td>
              <td style="padding:10px 8px;font-size:11px;color:#888;white-space:nowrap">
                <span style="background:{badge_color};color:{color};padding:2px 7px;border-radius:10px;font-size:10px;font-weight:600">{a.get('media_type','')}</span><br>
                <span style="margin-top:4px;display:block">{a['source']}</span>
                <span>{a.get('date','')}</span>
              </td>
            </tr>"""
        return html

    neg_rows = rows(negatif, "#C0392B", "#FDECEA")
    pos_rows = rows(positif, "#27AE60", "#E8F5E9")

    sentiment_label = "POSITIF ✅" if pct_pos >= 60 else ("NEGATIF ⚠️" if pct_pos < 40 else "NETRAL ⚡")
    sentiment_color = "#27AE60" if pct_pos >= 60 else ("#C0392B" if pct_pos < 40 else "#F39C12")

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Media Monitoring Bank Mandiri — {date_str}</title></head>
<body style="margin:0;padding:0;background:#F5F6FA;font-family:Arial,Helvetica,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#F5F6FA;padding:20px 0">
<tr><td align="center">
<table width="680" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.08)">

  <!-- Header -->
  <tr><td style="background:linear-gradient(135deg,#003087,#0055B3);padding:28px 32px">
    <div style="color:#fff;font-size:22px;font-weight:700">📊 Media Monitoring Bank Mandiri</div>
    <div style="color:#B0C8FF;font-size:13px;margin-top:6px">{date_str} &nbsp;|&nbsp; Sumber: Talkwalker Alerts &nbsp;|&nbsp; Analisis: Claude AI</div>
  </td></tr>

  <!-- Summary Cards -->
  <tr><td style="padding:24px 32px 16px">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td width="25%" style="padding:4px">
          <div style="background:#F0F4FF;border-radius:10px;padding:14px 12px;text-align:center">
            <div style="font-size:28px;font-weight:700;color:#003087">{total}</div>
            <div style="font-size:11px;color:#666;margin-top:4px">Total Berita</div>
          </div>
        </td>
        <td width="25%" style="padding:4px">
          <div style="background:#FEF0F0;border-radius:10px;padding:14px 12px;text-align:center">
            <div style="font-size:28px;font-weight:700;color:#C0392B">{len(negatif)}</div>
            <div style="font-size:11px;color:#666;margin-top:4px">Berita Negatif</div>
          </div>
        </td>
        <td width="25%" style="padding:4px">
          <div style="background:#F0FBF4;border-radius:10px;padding:14px 12px;text-align:center">
            <div style="font-size:28px;font-weight:700;color:#27AE60">{len(positif)}</div>
            <div style="font-size:11px;color:#666;margin-top:4px">Berita Positif</div>
          </div>
        </td>
        <td width="25%" style="padding:4px">
          <div style="background:#FFF8E6;border-radius:10px;padding:14px 12px;text-align:center">
            <div style="font-size:22px;font-weight:700;color:{sentiment_color}">{pct_pos}%</div>
            <div style="font-size:11px;color:{sentiment_color};margin-top:4px;font-weight:600">{sentiment_label}</div>
          </div>
        </td>
      </tr>
    </table>
  </td></tr>

  <!-- Chart -->
  <tr><td style="padding:8px 32px 8px">
    <div style="font-size:14px;font-weight:700;color:#333;margin-bottom:10px">📈 Grafik Sentimen</div>
    <img src="cid:sentiment_chart" width="100%" style="border-radius:8px;border:1px solid #eee" alt="Grafik Sentimen">
  </td></tr>

  <!-- Negative News -->
  <tr><td style="padding:20px 32px 8px">
    <div style="font-size:14px;font-weight:700;color:#C0392B;margin-bottom:10px">🔴 Berita Negatif — Sorted by Score</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #F5E6E6;border-radius:8px;overflow:hidden">
      <tr style="background:#FEF0F0">
        <th style="padding:8px;text-align:center;font-size:11px;color:#888;width:50px">SKOR</th>
        <th style="padding:8px;text-align:left;font-size:11px;color:#888">JUDUL & ANALISIS</th>
        <th style="padding:8px;text-align:left;font-size:11px;color:#888;width:130px">SUMBER</th>
      </tr>
      {neg_rows if neg_rows else '<tr><td colspan="3" style="padding:14px;text-align:center;color:#888;font-size:12px">Tidak ada berita negatif hari ini ✅</td></tr>'}
    </table>
  </td></tr>

  <!-- Positive News -->
  <tr><td style="padding:20px 32px 8px">
    <div style="font-size:14px;font-weight:700;color:#27AE60;margin-bottom:10px">🟢 Berita Positif — Sorted by Score</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #E6F5EA;border-radius:8px;overflow:hidden">
      <tr style="background:#F0FBF4">
        <th style="padding:8px;text-align:center;font-size:11px;color:#888;width:50px">SKOR</th>
        <th style="padding:8px;text-align:left;font-size:11px;color:#888">JUDUL & ANALISIS</th>
        <th style="padding:8px;text-align:left;font-size:11px;color:#888;width:130px">SUMBER</th>
      </tr>
      {pos_rows if pos_rows else '<tr><td colspan="3" style="padding:14px;text-align:center;color:#888;font-size:12px">Tidak ada berita positif hari ini.</td></tr>'}
    </table>
  </td></tr>

  <!-- Footer -->
  <tr><td style="padding:20px 32px;background:#F8F9FB;border-top:1px solid #eee;text-align:center">
    <div style="font-size:11px;color:#999">
      Laporan ini dibuat secara otomatis oleh sistem Media Monitoring Bank Mandiri<br>
      Powered by <strong>Claude AI (Anthropic)</strong> &amp; Talkwalker Alerts<br>
      Dikirim setiap hari pukul 07.00 WIB
    </div>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>"""
    return html


def send_email(articles, date_str, chart_path):
    """Kirim email HTML dengan grafik embedded ke EMAIL_RECIPIENT."""
    msg = MIMEMultipart("related")
    positif = len([a for a in articles if a.get("sentiment") == "positif"])
    negatif = len([a for a in articles if a.get("sentiment") == "negatif"])

    msg["Subject"] = f"📊 Media Monitoring BMRI — {date_str} | {positif} Positif / {negatif} Negatif"
    msg["From"]    = EMAIL_SENDER
    msg["To"]      = EMAIL_RECIPIENT

    html_body = build_html_email(articles, date_str, chart_path)
    alternative = MIMEMultipart("alternative")
    alternative.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(alternative)

    # Embed grafik sebagai inline image
    with open(chart_path, "rb") as f:
        img = MIMEImage(f.read(), _subtype="png")
    img.add_header("Content-ID", "<sentiment_chart>")
    img.add_header("Content-Disposition", "inline", filename="sentiment_chart.png")
    msg.attach(img)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
            smtp.login(EMAIL_SENDER, EMAIL_APP_PASS)
            smtp.send_message(msg)
        print(f"[Email] Terkirim ke {EMAIL_RECIPIENT}")
    except Exception as e:
        print(f"[Email] Error: {e}")
        raise


# ─────────────────────────────────────────────
# 9. MAIN
# ─────────────────────────────────────────────
def main():
    yesterday = (datetime.date.today() - datetime.timedelta(days=1))
    date_str  = yesterday.strftime("%d %B %Y")
    print(f"\n{'='*55}")
    print(f"  BANK MANDIRI MEDIA MONITORING — {date_str}")
    print(f"{'='*55}\n")

    # 1. Ambil email Gmail
    print("[1/6] Mengambil email Talkwalker dari Gmail...")
    service     = get_gmail_service()
    html_bodies = fetch_talkwalker_emails(service)

    if not html_bodies:
        print("  ⚠ Tidak ada email Talkwalker ditemukan untuk kemarin.")
        send_telegram_text(
            f"ℹ️ *Media Monitoring Bank Mandiri — {date_str}*\n\n"
            "Tidak ada email Talkwalker yang ditemukan untuk periode ini."
        )
        return

    # 2. Parse artikel
    print("[2/6] Mem-parsing artikel...")
    articles = parse_articles(html_bodies)

    if not articles:
        print("  ⚠ Tidak ada artikel relevan ditemukan.")
        return

    # 3. Analisis sentimen AI
    print("[3/6] Menganalisis sentimen dengan Claude AI...")
    articles = analyze_sentiment(articles)

    # 4. Buat grafik
    print("[4/6] Membuat grafik sentimen...")
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        chart_path = tmp.name
    generate_chart(articles, chart_path)

    # 5. Kirim Telegram
    print("[5/6] Mengirim ke Telegram...")
    send_telegram_text(format_telegram_negative(articles, date_str))
    send_telegram_text(format_telegram_positive(articles, date_str))
    send_telegram_photo(
        chart_path,
        caption=f"📊 *Grafik Sentimen Bank Mandiri — {date_str}*\n_{len(articles)} berita dianalisis_"
    )
    send_telegram_text(format_summary(articles, date_str))

    # 6. Kirim Email
    print("[6/6] Mengirim email...")
    send_email(articles, date_str, chart_path)

    # Cleanup
    os.unlink(chart_path)

    print(f"\n✅ Selesai! {len(articles)} artikel dianalisis dan dikirim.")


if __name__ == "__main__":
    main()
