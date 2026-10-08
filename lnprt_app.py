# ============================================================
# LNPRT - Scraper Berita Lembaga Non-Profit Rumah Tangga
# Output: Tanggal (dd/mm/yyyy), Judul, Sumber, Wilayah, Kategori, URL
# ============================================================

import streamlit as st
import time
import re
import os
import base64
import pandas as pd
import datetime as dt
from datetime import datetime
from typing import List, Dict, Set, Tuple, Any, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
import threading
import random
import requests
import feedparser
import urllib.parse
from googlenewsdecoder import gnewsdecoder
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

# ── Konfigurasi halaman ──────────────────────────────────────────────────────
st.set_page_config(page_title="Scraper Berita LNPRT", layout="wide", page_icon="📰")

# ── Path referensi ───────────────────────────────────────────────────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
KATA_KUNCI_PATH = os.path.join(_HERE, "Kata Kunci.xlsx")
KATEGORI_PATH   = os.path.join(_HERE, "Kategori.xlsx")
WILAYAH_PATH    = os.path.join(_HERE, "Daftar Wilayah.xlsx")
LOGO_PATH       = os.path.join(_HERE, "Logo.png")

UMUM        = "Umum"
DELAY_REQ   = 3  # Detik delay antar request untuk hindari blokir
MAX_WORKERS = 5  # Worker paralel

try:
    NEWS_API_KEY = st.secrets["NEWS_API_KEY"]
except Exception:
    NEWS_API_KEY = "4cf8032e0a0d4107a68443615aefd46a"

# ============================================================
# 1. LOAD DATA REFERENSI
# ============================================================

@st.cache_data(show_spinner=False)
def load_kata_kunci() -> Dict[str, List[str]]:
    """Baris 0 = nama kategori, baris 1+ = keyword."""
    if not os.path.exists(KATA_KUNCI_PATH): return {}
    try:
        df = pd.read_excel(KATA_KUNCI_PATH, sheet_name='keyword', header=None)
    except Exception:
        df = pd.read_excel(KATA_KUNCI_PATH, header=None)
    result: Dict[str, List[str]] = {}
    for col in range(df.shape[1]):
        cat = str(df.iloc[0, col]).strip()
        kws = [str(v).strip() for v in df.iloc[1:, col] if pd.notna(v) and str(v).strip()]
        if cat and kws:
            result[cat] = kws
    return result

@st.cache_data(show_spinner=False)
def load_kategori_dict() -> Dict[str, List[str]]:
    """Membaca keyword kategori dari Kategori.xlsx untuk proses penandaan."""
    if not os.path.exists(KATEGORI_PATH): return {}
    df = pd.read_excel(KATEGORI_PATH, header=None)
    result: Dict[str, List[str]] = {}
    for col in range(df.shape[1]):
        cat = str(df.iloc[0, col]).strip()
        kws = [str(v).strip() for v in df.iloc[1:, col] if pd.notna(v) and str(v).strip()]
        if cat and kws:
            result[cat] = kws
    return result

@st.cache_data(show_spinner=False)
def load_persepsi() -> Tuple[List[str], List[str]]:
    """Membaca keyword persepsi dari sheet 'persepsi'."""
    if not os.path.exists(KATA_KUNCI_PATH): return [], []
    try:
        try:
            df = pd.read_excel(KATA_KUNCI_PATH, sheet_name='persepsi', header=None)
        except Exception:
            df = pd.read_excel(KATA_KUNCI_PATH, sheet_name='persepi', header=None)
    except Exception:
        return [], []
        
    pos_kws = []
    neg_kws = []
    for col in range(df.shape[1]):
        header = str(df.iloc[0, col]).strip()
        kws = [str(v).strip() for v in df.iloc[1:, col] if pd.notna(v) and str(v).strip()]
        if header in ['1', '1.0']:
            pos_kws.extend(kws)
        elif header in ['-1', '-1.0']:
            neg_kws.extend(kws)
    return pos_kws, neg_kws

@st.cache_data(show_spinner=False)
def load_wilayah():
    if not os.path.exists(WILAYAH_PATH):
        return [], [], [], {}
    df = pd.read_excel(WILAYAH_PATH, header=0)
    df.columns = [c.strip() for c in df.columns]

    kab_set: Dict[str, Tuple[str, str]] = {}
    prov_seen: Dict[str, str] = {}
    prov_ui: Dict[str, str] = {}
    kab_ui: Dict[str, List[Tuple[str, str]]] = {}
    seen_kab_kode: Set[str] = set()

    for _, row in df.iterrows():
        kode_prov  = str(int(row["KODE PROV"])).zfill(2)
        nama_prov  = str(row["NAMA PROV"]).strip()
        kode_kab   = str(int(row["KODE KAB"])).zfill(4)
        nama_kab   = str(row["NAMA KAB"]).strip()

        clean_kab = re.sub(r"^(KABUPATEN|KOTA)\s+", "", nama_kab.upper()).strip()
        for name in {nama_kab.upper(), clean_kab}:
            if len(name) >= 3:
                kab_set[name] = (kode_kab, kode_prov)
        prov_seen.setdefault(kode_prov, nama_prov.upper())

        prov_ui.setdefault(kode_prov, nama_prov.title())
        if kode_kab not in seen_kab_kode:
            seen_kab_kode.add(kode_kab)
            kab_ui.setdefault(kode_prov, []).append((nama_kab.title(), kode_kab))

    kab_items  = sorted([(n, v[0], v[1]) for n, v in kab_set.items()],
                        key=lambda x: len(x[0]), reverse=True)
    prov_items = sorted([(name, kode) for kode, name in prov_seen.items()],
                        key=lambda x: len(x[0]), reverse=True)
    sorted_provs = sorted(prov_ui.items(), key=lambda x: x[1])
    for kp in kab_ui:
        kab_ui[kp] = sorted(kab_ui[kp], key=lambda x: x[0])

    return kab_items, prov_items, sorted_provs, kab_ui

# ============================================================
# 2. FUNGSI UTILITAS
# ============================================================

@st.cache_data(show_spinner=False)
def load_exclusion_list() -> List[str]:
    exclusion_path = os.path.join(_HERE, "Exclusion-list.txt")
    if not os.path.exists(exclusion_path):
        return []
    with open(exclusion_path, "r", encoding="utf-8") as f:
        lines = f.read().splitlines()
    return [l.strip() for l in lines if l.strip()]

def is_excluded(url: str, exclusion_list: List[str]) -> bool:
    if not url: return False
    url_lower = url.lower()
    for exc in exclusion_list:
        if exc.lower() in url_lower:
            return True
    return False

def parse_relative_time_to_date(text: str) -> str:
    now = datetime.now()
    text_lower = text.lower()
    
    m_min = re.search(r'(\d+)\s*(?:m\b|mnt\b|menit\b)', text_lower)
    if m_min and 'minggu' not in text_lower:
        return (now - dt.timedelta(minutes=int(m_min.group(1)))).strftime("%d/%m/%Y")
        
    m_hour = re.search(r'(\d+)\s*(?:j\b|jam\b|h\b|hour\b)', text_lower)
    if m_hour and 'hari' not in text_lower:
        return (now - dt.timedelta(hours=int(m_hour.group(1)))).strftime("%d/%m/%Y")
        
    m_day = re.search(r'(\d+)\s*(?:hari\b|d\b|day\b)', text_lower)
    if m_day:
        return (now - dt.timedelta(days=int(m_day.group(1)))).strftime("%d/%m/%Y")
        
    m_week = re.search(r'(\d+)\s*(?:mgg\b|minggu\b|w\b|week\b)', text_lower)
    if m_week:
        return (now - dt.timedelta(weeks=int(m_week.group(1)))).strftime("%d/%m/%Y")
        
    m_month = re.search(r'(\d+)\s*(?:bln\b|bulan\b|mo\b|month\b)\s?', text_lower)
    if m_month:
        return (now - dt.timedelta(days=int(m_month.group(1))*30)).strftime("%d/%m/%Y")
        
    m_year = re.search(r'(\d+)\s*(?:thn\b|tahun\b|yr\b|year\b)\s?', text_lower)
    if m_year:
        return (now - dt.timedelta(days=int(m_year.group(1))*365)).strftime("%d/%m/%Y")
        
    return text

def clean_source_and_date(src: str, pub: str) -> Tuple[str, str]:
    if not src: src = "-"
    if not pub: pub = ""

    _TIME_PAT = (
        r'\d+\s*'
        r'(?:jam|hour|mnt|menit|hari|day|mgg|minggu|week|bln|bulan|months?|mo|tahun|thn|years?|yr'
        r'|[jhdwm]\b)'
        r'(?:\s+yang\s+lalu|\s+ago)?'
    )

    m_end = re.search(rf'({_TIME_PAT})$', src, flags=re.IGNORECASE)
    if m_end:
        time_str = m_end.group(1)
        src = src[:m_end.start()].strip()
        if not pub or pub == "-":
            pub = time_str

    m_start = re.match(rf'^({_TIME_PAT})\s*', src, flags=re.IGNORECASE)
    if m_start:
        time_str = m_start.group(1)
        src = src[m_start.end():].strip()
        if not pub or pub == "-":
            pub = time_str

    m_msn = re.search(r'\s*on\s+MSN\w*', src, flags=re.IGNORECASE)
    if m_msn:
        src = src[:m_msn.start()].strip()

    src = re.sub(r'\s*\d+(?:[.,]\d+)?\s*(?:juta|ribu|rb|k|m)?\s*$', '', src, flags=re.IGNORECASE).strip()

    if re.match(r'^\d+$', src.strip()):
        src = "-"

    if 0 < len(src) < 3:
        src = "-"

    if not src: src = "-"
    return src, pub


def source_from_url(url: str) -> str:
    if not url: return "-"
    try:
        from urllib.parse import urlparse
        host = urlparse(url).hostname or ""
        host = re.sub(r'^(www|m|mobile|amp)\.', '', host)
        _DOMAIN_MAP = {
            "kompas.com": "Kompas", "detik.com": "Detik", "tribunnews.com": "Tribun News",
            "cnnindonesia.com": "CNN Indonesia", "tempo.co": "Tempo", "liputan6.com": "Liputan6",
            "okezone.com": "Okezone", "republika.co.id": "Republika", "jpnn.com": "JPNN",
            "medcom.id": "Medcom", "beritasatu.com": "BeritaSatu", "antaranews.com": "Antara News",
            "sindonews.com": "Sindo News", "merdeka.com": "Merdeka", "suara.com": "Suara",
            "bisnis.com": "Bisnis Indonesia", "msn.com": "MSN", "viva.co.id": "VIVA",
            "cnbcindonesia.com": "CNBC Indonesia", "inews.id": "iNews", "tvonenews.com": "tvOneNews",
            "kumparan.com": "Kumparan", "pikiran-rakyat.com": "Pikiran Rakyat", "jawapos.com": "Jawa Pos"
        }
        for domain, name in _DOMAIN_MAP.items():
            if host.endswith(domain):
                return name
        parts = host.rsplit(".")
        if len(parts) >= 3 and parts[-2].lower() in ["co", "or", "go", "ac", "sch", "my", "biz", "web", "desa"]:
            return parts[-3].capitalize()
        return parts[-2].capitalize() if len(parts) >= 2 else host
    except Exception:
        return "-"

def format_tanggal(published: str) -> str:
    if not published or published == "-":
        return "-"
        
    rel_date = parse_relative_time_to_date(published)
    if re.match(r'\d{2}/\d{2}/\d{4}', rel_date):
        return rel_date
        
    for fmt in ("%a, %d %b %Y %H:%M:%S %Z", "%a, %d %b %Y %H:%M:%S %z", "%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%d"):
        try:
            return datetime.strptime(published.strip(), fmt).strftime("%d/%m/%Y")
        except Exception:
            pass
            
    try:
        return datetime.fromisoformat(published.strip()).strftime("%d/%m/%Y")
    except Exception:
        pass
        
    try:
        from dateutil import parser
        bulan = {
            'januari': 'Jan', 'februari': 'Feb', 'maret': 'Mar', 'april': 'Apr',
            'mei': 'May', 'juni': 'Jun', 'juli': 'Jul', 'agustus': 'Aug',
            'september': 'Sep', 'oktober': 'Oct', 'november': 'Nov', 'desember': 'Dec'
        }
        date_lower = published.strip().lower()
        for id_month, en_month in bulan.items():
            date_lower = re.sub(r'\b' + id_month + r'\b', en_month.lower(), date_lower)
            
        parsed_date = parser.parse(date_lower, fuzzy=True)
        return parsed_date.strftime("%d/%m/%Y")
    except Exception:
        pass
        
    return published


def detect_wilayah(text: str, kab_items, prov_items) -> str:
    if not text: return ""
    text_up = text.upper()
    found_kab: Set[str] = set()
    covered_prov: Set[str] = set()

    for (name, kode_kab, kode_prov) in kab_items:
        pattern = r"(?<![A-Z])" + re.escape(name) + r"(?![A-Z])"
        if re.search(pattern, text_up):
            found_kab.add(kode_kab)
            covered_prov.add(kode_prov)

    found_prov: Set[str] = set()
    for (name, kode_prov) in prov_items:
        if kode_prov in covered_prov:
            continue
        pattern = r"(?<![A-Z])" + re.escape(name) + r"(?![A-Z])"
        if re.search(pattern, text_up):
            found_prov.add(kode_prov)

    all_codes = sorted(found_kab | found_prov)
    return ", ".join(all_codes)


def detect_persepsi(title: str, text: str, pos_kws: List[str], neg_kws: List[str], check_text: bool) -> int:
    def count_kws(content: str, kws: List[str]) -> int:
        if not content: return 0
        count = 0
        content_up = content.upper()
        for kw in kws:
            pattern = r"(?<![A-Z])" + re.escape(kw.upper()) + r"(?![A-Z])"
            count += len(re.findall(pattern, content_up))
        return count

    pos_count = count_kws(title, pos_kws)
    neg_count = count_kws(title, neg_kws)
    
    if pos_count == 0 and neg_count == 0 and check_text and text:
        pos_count = count_kws(text, pos_kws)
        neg_count = count_kws(text, neg_kws)
        
    if pos_count > neg_count:
        return 1
    elif neg_count > pos_count:
        return -1
    else:
        return 0


def detect_kategori(text: str, kategori_dict: Dict[str, List[str]], initial_cats: Set[str]) -> str:
    text_up = text.upper()
    matched: Set[str] = set(initial_cats)
    
    matched.discard("Umum")
    matched.discard("Custom Keyword")
    
    for cat, kws in kategori_dict.items():
        if cat in matched:
            continue
        for kw in kws:
            pattern = r"(?<![A-Z])" + re.escape(kw.upper()) + r"(?![A-Z])"
            if re.search(pattern, text_up):
                matched.add(cat)
                break
                
    if matched:
        return ", ".join(sorted(matched))
    return ""


def fetch_article_data(url: str) -> Dict[str, str]:
    result = {"text": "", "date": ""}
    try:
        import requests
        from bs4 import BeautifulSoup
        import json
        import trafilatura
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            html = response.text
            result["text"] = trafilatura.extract(html) or ""
            
            soup = BeautifulSoup(html, 'html.parser')
            for script in soup.find_all('script', {'type': 'application/ld+json'}):
                try:
                    data = json.loads(script.string)
                    if isinstance(data, dict) and data.get('datePublished'):
                        result["date"] = data.get('datePublished')
                        break
                except Exception:
                    pass
    except Exception:
        pass
    return result

# ============================================================
# 3. EXPORT & DISPLAY
# ============================================================

_persepsi_renderer = JsCode("""
function(params) {
    let val = params.value;
    if (val == 1 || val == '1') {
        return '<span style="color:#4CAF50; font-weight:bold; font-size:16px; display:block; text-align:center;">▲</span>';
    } else if (val == -1 || val == '-1') {
        return '<span style="color:#F44336; font-weight:bold; font-size:16px; display:block; text-align:center;">▼</span>';
    } else {
        return '<span style="color:#9E9E9E; font-weight:bold; font-size:16px; display:block; text-align:center;">=</span>';
    }
}
""")

_link_btn = JsCode("""
function(params) {
    return '<a href="' + params.value + '" target="_blank" style="color:#2196F3; font-weight:bold; text-decoration:none;">🔗 Buka</a>';
}
""")

def to_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Berita LNPRT")
    return buf.getvalue()


def show_aggrid(df: pd.DataFrame):
    df_excel = df.reset_index(drop=True)

    df_display = df_excel.copy()
    df_display.insert(0, "Buka", df_display["URL"])

    gb = GridOptionsBuilder.from_dataframe(df_display)
    gb.configure_pagination(paginationPageSize=15)
    gb.configure_side_bar()
    gb.configure_default_column(editable=False, groupable=True, resizable=True)
    gb.configure_grid_options(enableRangeSelection=True, enableCellTextSelection=True)
    gb.configure_column("Persepsi", cellRenderer=_persepsi_renderer, width=90, type=["numericColumn"])
    gb.configure_column("Buka", cellRenderer=_link_btn, width=90, pinned="left", suppressSizeToFit=True)
    gridOptions = gb.build()

    c1, c2 = st.columns([8, 2])
    with c1:
        st.markdown("<h3 style='margin:0;font-size:24px;'>Hasil Scraping</h3>", unsafe_allow_html=True)
    with c2:
        fname = f"berita_lnprt_{dt.date.today().strftime('%Y%m%d')}.xlsx"
        st.download_button("⬇️ Download Excel", data=to_excel(df_excel),
                           file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

    AgGrid(df_display, gridOptions=gridOptions, theme="light",
           fit_columns_on_grid_load=False, height=500,
           suppressRowClickSelection=True, allow_unsafe_jscode=True)

# ============================================================
# 4. CACHED SEARCH & DECODE
# ============================================================

@st.cache_data(ttl=3600, show_spinner=False)
def cached_bing_search(keyword: str, start_date: dt.date, end_date: dt.date, wilayah_term: str = "") -> Tuple[List[Dict], List[str]]:
    from bs4 import BeautifulSoup
    import urllib.parse

    all_entries: List[Dict] = []
    errors: List[str] = []
    seen_urls: Set[str] = set()

    q_base = f'{keyword} {wilayah_term}'.strip() if wilayah_term else keyword
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'id-ID,id;q=0.9,en;q=0.8',
    }

    try:
        encoded_query = urllib.parse.quote(q_base)
        url = f'https://www.bing.com/news/search?q={encoded_query}&setlang=id-ID'

        for attempt in range(3):
            try:
                response = requests.get(url, headers=headers, timeout=15)
                if response.status_code == 200:
                    soup = BeautifulSoup(response.text, 'html.parser')
                    articles = soup.select('div.news-card, div.newsitem')

                    for article in articles:
                        try:
                            title_elem = article.select_one('a.title, h3 a, .title a')
                            if not title_elem: continue

                            title = title_elem.text.strip()
                            link = title_elem.get('href', '')

                            if not link or link in seen_urls: continue

                            src_elem = article.select_one('.source, .provider, cite')
                            src_raw = src_elem.text.strip() if src_elem else ""

                            src_from_url = source_from_url(link)
                            src = src_from_url if src_from_url != "-" else clean_source_and_date(src_raw, "")[0]

                            published = ""
                            date_elem = article.select_one('time, .date, .timestamp')
                            if date_elem:
                                published = date_elem.get('datetime', '') or date_elem.text.strip()

                            if not published and src_raw:
                                _, published = clean_source_and_date(src_raw, "")

                            seen_urls.add(link)
                            all_entries.append({
                                "title": title, "published": published, "link": link, "source": src
                            })
                        except Exception:
                            continue
                    break
            except Exception as exc:
                if attempt == 2:
                    errors.append(f"Bing News: {type(exc).__name__}: {exc}")
                time.sleep(2 ** attempt)

        time.sleep(DELAY_REQ)
    except Exception as exc:
        errors.append(f"Search error: {type(exc).__name__}: {exc}")

    return all_entries, errors


def _strip_title_source_suffix(title: str, source_name: str) -> str:
    if not title or not source_name or source_name == "-":
        return title
    suffix = f" - {source_name}"
    if title.lower().endswith(suffix.lower()):
        return title[: -len(suffix)].strip()
    return title


_GNEWS_UA_LIST = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
]

def _fetch_gnews_rss_day(q_base: str, day: dt.date, _retry_count: int = 3) -> Tuple[List[Dict], Optional[str]]:
    next_day = day + dt.timedelta(days=1)
    query = f'{q_base} after:{day.strftime("%Y-%m-%d")} before:{next_day.strftime("%Y-%m-%d")}'
    url = f"https://news.google.com/rss/search?q={urllib.parse.quote(query)}&hl=id&gl=ID&ceid=ID:id"

    entries: List[Dict] = []
    for attempt in range(_retry_count):
        headers = {'User-Agent': _GNEWS_UA_LIST[attempt % len(_GNEWS_UA_LIST)]}
        try:
            resp = requests.get(url, headers=headers, timeout=20)
            if resp.status_code == 200:
                feed = feedparser.parse(resp.content)
                for e in feed.entries:
                    title     = getattr(e, "title", "-") or "-"
                    published = getattr(e, "published", "") or ""
                    link      = getattr(e, "link", "") or ""

                    src = "-"
                    try:
                        src = e.source.title
                    except Exception:
                        pass

                    title = _strip_title_source_suffix(title, src)
                    src, published = clean_source_and_date(src, published)

                    if link:
                        entries.append({"title": title, "published": published, "link": link, "source": src})
                return entries, None
        except Exception:
            pass
        time.sleep(2 ** attempt)

    return entries, f"{day}: GNews RSS Failed"


@st.cache_data(ttl=3600, show_spinner=False)
def cached_google_search(keyword: str, start_date: dt.date, end_date: dt.date, wilayah_term: str = "") -> Tuple[List[Dict], List[str]]:
    all_entries: List[Dict] = []
    errors: List[str] = []
    q_base = f'{keyword} {wilayah_term}'.strip() if wilayah_term else keyword

    current = start_date
    while current <= end_date:
        entries, err = _fetch_gnews_rss_day(q_base, current)
        all_entries.extend(entries)
        if err: errors.append(err)
        current += dt.timedelta(days=1)
        time.sleep(1.0)

    return all_entries, errors


@st.cache_data(ttl=24*3600, show_spinner=False)
def decode_url_once(link: str) -> str:
    import concurrent.futures as _cf
    def _do_decode():
        r = gnewsdecoder(link)
        return r["decoded_url"] if r.get("status") else link
    try:
        with _cf.ThreadPoolExecutor(max_workers=1) as _ex:
            fut = _ex.submit(_do_decode)
            return fut.result(timeout=10)
    except Exception:
        return link

@st.cache_data(ttl=3600, show_spinner=False)
def cached_fetch_article_data(url: str) -> Dict[str, str]:
    return fetch_article_data(url)

def _call_google_search_locked(keyword: str, start_date: dt.date, end_date: dt.date, wilayah_term: str, lock: threading.Lock) -> Tuple[List[Dict], List[str]]:
    with lock:
        return cached_google_search(keyword, start_date, end_date, wilayah_term)

# ============================================================
# 5. MAIN SCRAPER FUNCTION
# ============================================================

def jalankan_scraper(kata_kunci, kategori_dict, custom_kw_list, kab_items, prov_items, pos_kws, neg_kws, selected_cats, start_date, end_date, wilayah_term="", per_kw_limit=0, decode_url=True, fetch_artikel=True, max_ws=3, max_wd=5, max_wf=3, via_selected="Semua"):
    sources = ["google", "bing"] if via_selected == "Semua" else (["google"] if via_selected == "Google News" else ["bing"])

    tasks: List[Tuple[str, str, str]] = []
    if "Custom Keyword" in selected_cats:
        for kw in custom_kw_list:
            for src in sources: tasks.append((kw, "Custom Keyword", src))
    else:
        for cat in selected_cats:
            for kw in kata_kunci.get(cat, []):
                for src in sources: tasks.append((kw, cat, src))

    if not tasks:
        st.warning("Tidak ada kata kunci yang dipilih.")
        return

    progress = st.progress(0.0)
    status   = st.empty()
    status.info(f"🔄 Mempersiapkan {len(tasks)} pencarian...")

    exclusion_list = load_exclusion_list()
    _google_lock = threading.Lock()
    done = 0
    by_link: Dict[str, Dict] = {}
    all_search_errors: List[str] = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        future_map = {}
        for kw, cat, src in tasks:
            if src == "google":
                fut = ex.submit(_call_google_search_locked, kw, start_date, end_date, wilayah_term, _google_lock)
            else:
                fut = ex.submit(cached_bing_search, kw, start_date, end_date, wilayah_term)
            future_map[fut] = (kw, cat, src)

        for fut in as_completed(future_map):
            kw, cat, src = future_map[fut]
            try:
                entries, errs = fut.result()
                entries = entries or []
                all_search_errors.extend([f"[{kw}][{src}] {e}" for e in errs])
            except Exception as exc:
                entries = []
                all_search_errors.append(f"[{kw}][{src}] Failed: {exc}")

            if per_kw_limit > 0: entries = entries[:per_kw_limit]

            for e in entries:
                raw_link = e.get("link", "") or ""
                if is_excluded(raw_link, exclusion_list): continue
                
                pub_date_str = format_tanggal(e.get("published", ""))
                link = raw_link.split("?")[0]
                if not link: continue
                
                if link not in by_link:
                    by_link[link] = {
                        "title":     e.get("title", "-"),
                        "published": pub_date_str,
                        "source":    e.get("source", "-"),
                        "cats":      set([cat]),
                        "keywords":  set([kw]),
                    }
                else:
                    by_link[link]["cats"].add(cat)
                    by_link[link]["keywords"].add(kw)

            done += 1
            progress.progress(done / max(1, len(tasks)))

    if not by_link:
        progress.empty(); status.empty()
        st.warning("Tidak ada artikel ditemukan.")
        st.session_state.scraped_data = pd.DataFrame(columns=["Tanggal", "Judul", "Sumber", "Wilayah", "Kategori", "Persepsi", "Keywords", "Hashtag", "URL"])
        return

    gnews_links = list(by_link.keys())
    decoded_map = {ln: decode_url_once(ln) if decode_url else ln for ln in gnews_links}
    
    records = []
    for gnews_link, obj in by_link.items():
        real_url = decoded_map.get(gnews_link, gnews_link)
        art_data = cached_fetch_article_data(real_url) if fetch_artikel else {"text": "", "date": ""}
        full_text = obj["title"] + " " + art_data["text"]

        wilayah  = detect_wilayah(full_text, kab_items, prov_items)
        kategori = detect_kategori(full_text, kategori_dict, obj["cats"])
        persepsi = detect_persepsi(obj["title"], art_data["text"], pos_kws, neg_kws, fetch_artikel)

        published = obj["published"]
        if (not published or published == "-") and art_data.get("date"):
            published = format_tanggal(art_data["date"])

        records.append({
            "Tanggal":  published,
            "Judul":    obj["title"],
            "Sumber":   obj["source"],
            "Wilayah":  wilayah,
            "Kategori": kategori,
            "Persepsi": persepsi,
            "Keywords": ", ".join(sorted(obj["keywords"])),
            "Hashtag":  ", ".join(sorted(set(re.findall(r'#\w+', full_text)))),
            "URL":      real_url,
        })

    df = pd.DataFrame(records)
    st.session_state.scraped_data = df
    progress.empty(); status.empty()
    st.success(f"✅ Selesai! {len(df)} artikel terproses.")

# ============================================================
# 6. STREAMLIT UI
# ============================================================

with st.spinner("Memuat data referensi..."):
    kata_kunci = load_kata_kunci()
    kategori_dict = load_kategori_dict()
    pos_kws, neg_kws = load_persepsi()
    kab_items, prov_items, sorted_provs, kab_by_prov = load_wilayah()

semua_kategori = list(kata_kunci.keys())

with st.container(border=True):
    col_kat, col_prov, col_kab = st.columns(3)
    with col_kat:
        selected_cat = st.selectbox("Pilih kategori", ["Semua", "Custom Keyword"] + semua_kategori)
    with col_prov:
        selected_prov_name = st.selectbox("Pilih provinsi", ["Semua"] + [name for _, name in sorted_provs])
    with col_kab:
        if selected_prov_name == "Semua":
            kab_display_list = ["—"]
            kab_disabled = True
        else:
            selected_kode_prov = next((k for k, n in sorted_provs if n == selected_prov_name), None)
            kab_list = kab_by_prov.get(selected_kode_prov, [])
            kab_display_list = ["Semua"] + [n for n, _ in kab_list]
            kab_disabled = False
        selected_kab_name = st.selectbox("Pilih kab/kota", kab_display_list, disabled=kab_disabled)

    if selected_cat == "Custom Keyword":
        custom_kw_str = st.text_input("📝 Masukkan custom keyword (pisahkan dengan koma)")
        custom_kw_list = [k.strip() for k in custom_kw_str.split(",") if k.strip()]
    else:
        custom_kw_list = []

    selected_cats = semua_kategori if selected_cat == "Semua" else ([selected_cat] if selected_cat != "Custom Keyword" else ["Custom Keyword"])

    col_tgl, col_src, col_limit = st.columns(3)
    with col_tgl:
        today = dt.date.today()
        periode = st.date_input("Periode Tanggal", value=(today - dt.timedelta(days=30), today))
        start_date, end_date = periode if isinstance(periode, tuple) and len(periode) == 2 else (today - dt.timedelta(days=30), today)
    with col_src:
        _via_selected = st.selectbox("Sumber Berita", ["Semua", "Google News", "Bing News"])
    with col_limit:
        per_kw_limit = st.selectbox("Limit", [0, 5, 10, 25, 50, 100], format_func=lambda x: "Tidak Terbatas" if x == 0 else str(x))

    col_opt1, col_opt2 = st.columns(2)
    with col_opt1: decode_url_toggle = st.checkbox("🔓 Decode URL asli", value=True)
    with col_opt2: fetch_teks_toggle = st.checkbox("📄 Fetch teks artikel", value=False)

wilayah_term = "" if selected_prov_name == "Semua" else (selected_prov_name if selected_kab_name in ("Semua", "—", None) else selected_kab_name)

scrape_button = st.button("🔍 Mulai Scraping", use_container_width=True)

if "scraped_data" not in st.session_state:
    st.session_state.scraped_data = pd.DataFrame(columns=["Tanggal", "Judul", "Sumber", "Wilayah", "Kategori", "Persepsi", "Keywords", "Hashtag", "URL"])

if scrape_button:
    jalankan_scraper(
        kata_kunci, kategori_dict, custom_kw_list, kab_items, prov_items,
        pos_kws, neg_kws, selected_cats, start_date, end_date,
        wilayah_term, per_kw_limit, decode_url_toggle, fetch_teks_toggle,
        via_selected=_via_selected
    )

if not st.session_state.scraped_data.empty:
    show_aggrid(st.session_state.scraped_data)
else:
    st.info("Belum ada data. Klik **Mulai Scraping** untuk memulai.")
