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

UMUM       = "Umum"
DELAY_REQ  = 3  # Detik delay antar request untuk hindari blokir
MAX_WORKERS = 5  # Worker paralel (Streamlit Cloud: 2 vCPUs)

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
    """
    Return:
      kab_items   : list[(name_upper, kode_kab, kode_prov)]  – for regex detection
      prov_items  : list[(name_upper, kode_prov)]             – for regex detection
      sorted_provs: list[(kode_prov, display_name)]           – sorted alpha, for UI
      kab_by_prov : dict[kode_prov -> list[(display_name, kode_kab)]] – for UI
    """
    df = pd.read_excel(WILAYAH_PATH, header=0)
    df.columns = [c.strip() for c in df.columns]

    # --- detection structures ---
    kab_set: Dict[str, Tuple[str, str]] = {}   # name_upper -> (kode_kab, kode_prov)
    prov_seen: Dict[str, str] = {}             # kode_prov  -> name_upper

    # --- UI structures ---
    prov_ui: Dict[str, str] = {}               # kode_prov -> display name (Title Case)
    kab_ui: Dict[str, List[Tuple[str, str]]] = {}  # kode_prov -> [(display_name, kode_kab)]
    seen_kab_kode: Set[str] = set()

    for _, row in df.iterrows():
        kode_prov  = str(int(row["KODE PROV"])).zfill(2)
        nama_prov  = str(row["NAMA PROV"]).strip()
        kode_kab   = str(int(row["KODE KAB"])).zfill(4)
        nama_kab   = str(row["NAMA KAB"]).strip()

        # detection
        clean_kab = re.sub(r"^(KABUPATEN|KOTA)\s+", "", nama_kab.upper()).strip()
        for name in {nama_kab.upper(), clean_kab}:
            if len(name) >= 3:
                kab_set[name] = (kode_kab, kode_prov)
        prov_seen.setdefault(kode_prov, nama_prov.upper())

        # UI
        prov_ui.setdefault(kode_prov, nama_prov.title())
        if kode_kab not in seen_kab_kode:
            seen_kab_kode.add(kode_kab)
            kab_ui.setdefault(kode_prov, []).append((nama_kab.title(), kode_kab))

    kab_items  = sorted([(n, v[0], v[1]) for n, v in kab_set.items()],
                        key=lambda x: len(x[0]), reverse=True)
    prov_items = sorted([(name, kode) for kode, name in prov_seen.items()],
                        key=lambda x: len(x[0]), reverse=True)
    sorted_provs = sorted(prov_ui.items(), key=lambda x: x[1])  # [(kode_prov, display_name)]
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
    
    # Gunakan \b (word boundary) setelah setiap pola satuan waktu
    # agar "j" tidak cocok dengan "Jul", "m" tidak cocok dengan "Mar", dll.
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
        r'|[jhdwm]\b)'  # Huruf tunggal HARUS di word boundary agar tidak cocok dengan 'juta', 'meter', dll
        r'(?:\s+yang\s+lalu|\s+ago)?'
    )

    # 1. Strip waktu di AKHIR string: "Tempo.co1h", "TribunNews2thn"
    m_end = re.search(rf'({_TIME_PAT})$', src, flags=re.IGNORECASE)
    if m_end:
        time_str = m_end.group(1)
        src = src[:m_end.start()].strip()
        if not pub or pub == "-":
            pub = time_str

    # 2. Strip waktu di AWAL string: "9hon MSNOpinion" → "on MSNOpinion"
    m_start = re.match(rf'^({_TIME_PAT})\s*', src, flags=re.IGNORECASE)
    if m_start:
        time_str = m_start.group(1)
        src = src[m_start.end():].strip()
        if not pub or pub == "-":
            pub = time_str

    # 3. Hapus "on MSN" + kata kategori opsional setelahnya: "on MSNOpinion", "on MSN"
    m_msn = re.search(r'\s*on\s+MSN\w*', src, flags=re.IGNORECASE)
    if m_msn:
        src = src[:m_msn.start()].strip()

    # 4. Strip angka view/read count yang menempel: "Media Indonesia27 juta", "56 juta", "3471"
    src = re.sub(r'\s*\d+(?:[.,]\d+)?\s*(?:juta|ribu|rb|k|m)?\s*$', '', src, flags=re.IGNORECASE).strip()

    # 5. Tolak jika sumber hanya angka murni
    if re.match(r'^\d+$', src.strip()):
        src = "-"

    # 6. Tolak sumber terlalu pendek (< 3 karakter) — kemungkinan artefak HTML seperti "Co"
    if 0 < len(src) < 3:
        src = "-"

    if not src: src = "-"
    return src, pub


def source_from_url(url: str) -> str:
    """Tebak nama sumber dari domain URL sebagai fallback."""
    if not url:
        return "-"
    try:
        from urllib.parse import urlparse
        host = urlparse(url).hostname or ""
        # Hapus www. dan subdomain umum
        host = re.sub(r'^(www|m|mobile|amp)\.', '', host)
        # Mapping domain terkenal ke nama yang lebih bersih
        _DOMAIN_MAP = {
            "kompas.com":      "Kompas",
            "detik.com":       "Detik",
            "tribunnews.com":  "Tribun News",
            "cnnindonesia.com":"CNN Indonesia",
            "tempo.co":        "Tempo",
            "liputan6.com":    "Liputan6",
            "okezone.com":     "Okezone",
            "republika.co.id": "Republika",
            "jpnn.com":        "JPNN",
            "medcom.id":       "Medcom",
            "beritasatu.com":  "BeritaSatu",
            "antara.news":     "Antara News",
            "antaranews.com":  "Antara News",
            "sindonews.com":   "Sindo News",
            "merdeka.com":     "Merdeka",
            "suara.com":       "Suara",
            "bisnis.com":      "Bisnis Indonesia",
            "msn.com":         "MSN",
            "viva.co.id":      "VIVA",
            "cnbcindonesia.com": "CNBC Indonesia",
            "inews.id":        "iNews",
            "tvonenews.com":   "tvOneNews",
            "kumparan.com":    "Kumparan",
            "pikiran-rakyat.com": "Pikiran Rakyat",
            "jawapos.com":     "Jawa Pos"
        }
        # Cek exact dan partial match
        for domain, name in _DOMAIN_MAP.items():
            if host.endswith(domain):
                return name
        # Fallback: ambil domain utama dan kapitalisasi
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
            
    # Fallback to isoformat
    try:
        return datetime.fromisoformat(published.strip()).strftime("%d/%m/%Y")
    except Exception:
        pass
        
    # Terjemahkan bulan Indonesia dan gunakan dateutil.parser untuk parse teks bebas
    try:
        from dateutil import parser
        bulan = {
            'januari': 'Jan', 'februari': 'Feb', 'maret': 'Mar', 'april': 'Apr',
            'mei': 'May', 'juni': 'Jun', 'juli': 'Jul', 'agustus': 'Aug',
            'september': 'Sep', 'oktober': 'Oct', 'november': 'Nov', 'desember': 'Dec',
            'jan': 'Jan', 'feb': 'Feb', 'mar': 'Mar', 'apr': 'Apr',
            'agu': 'Aug', 'sep': 'Sep', 'okt': 'Oct', 'nov': 'Nov', 'des': 'Dec'
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
    if not text:
        return ""
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


def detect_kategori(text: str, kategori_dict: Dict[str, List[str]],
                    initial_cats: Set[str]) -> str:
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
        # Tambahkan timeout agar tidak hang
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            html = response.text
            
            # 1. Ekstrak teks menggunakan trafilatura
            result["text"] = trafilatura.extract(html) or ""
            
            # 2. Ekstrak tanggal sebagai fallback menggunakan BeautifulSoup
            soup = BeautifulSoup(html, 'html.parser')
            
            # Coba JSON-LD dulu
            for script in soup.find_all('script', {'type': 'application/ld+json'}):
                try:
                    data = json.loads(script.string)
                    if isinstance(data, dict):
                        date_pub = data.get('datePublished')
                        if date_pub:
                            result["date"] = date_pub
                            break
                    elif isinstance(data, list):
                        for item in data:
                            if isinstance(item, dict) and item.get('datePublished'):
                                result["date"] = item.get('datePublished')
                                break
                except Exception:
                    pass
                if result["date"]: break
            
            # Coba meta tag jika JSON-LD gagal
            if not result["date"]:
                meta_tags = [
                    ('meta', {'property': 'article:published_time'}, 'content'),
                    ('meta', {'name': 'pubdate'}, 'content'),
                    ('time', {}, 'datetime')
                ]
                for tag, attrs, prop in meta_tags:
                    for el in soup.find_all(tag, attrs):
                        val = el.get(prop)
                        if val:
                            result["date"] = val
                            break
                    if result["date"]: break
    except Exception:
        pass
    return result

# ============================================================
# 3. EXPORT & DISPLAY
# ============================================================

_persepsi_renderer = JsCode("""
class PersepsiRenderer {
    init(params) {
        this.eGui = document.createElement('span');
        let val = params.value;
        if (val == 1 || val == '1') {
            this.eGui.innerHTML = '▲';
            this.eGui.style.color = '#4CAF50';
        } else if (val == -1 || val == '-1') {
            this.eGui.innerHTML = '▼';
            this.eGui.style.color = '#F44336';
        } else {
            this.eGui.innerHTML = '=';
            this.eGui.style.color = '#9E9E9E';
        }
        this.eGui.style.fontWeight = 'bold';
        this.eGui.style.fontSize = '16px';
        this.eGui.style.textAlign = 'center';
        this.eGui.style.display = 'block';
    }
    getGui() { return this.eGui; }
}
""")

_link_btn = JsCode("""
class LinkRenderer {
    init(params) {
        this.eGui = document.createElement('a');
        this.eGui.innerHTML = '🔗 Buka';
        this.eGui.setAttribute('href', params.value);
        this.eGui.setAttribute('target', '_blank');
        this.eGui.style.cssText = 'color:#2196F3;font-weight:bold;text-decoration:none;';
    }
    getGui() { return this.eGui; }
}
""")

def to_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Berita LNPRT")
    return buf.getvalue()


def show_aggrid(df: pd.DataFrame):
    # Excel uses original df (no Buka column)
    df_excel = df.reset_index(drop=True)

    # Display df adds Buka button column
    df_display = df_excel.copy()
    df_display.insert(0, "Buka", df_display["URL"])

    gb = GridOptionsBuilder.from_dataframe(df_display)
    gb.configure_pagination(paginationPageSize=15)
    gb.configure_side_bar()
    gb.configure_default_column(editable=False, groupable=True, resizable=True)
    gb.configure_grid_options(enableRangeSelection=True, enableCellTextSelection=True)
    gb.configure_column("Persepsi", cellRenderer=_persepsi_renderer, width=90, type=["numericColumn"])
    gb.configure_column("Buka", cellRenderer=_link_btn, width=90,
                        pinned="left", suppressSizeToFit=True)
    gridOptions = gb.build()

    c1, c2 = st.columns([8, 2])
    with c1:
        st.markdown(
            "<h3 style='margin:0;font-size:24px;'>Hasil Scraping</h3>",
            unsafe_allow_html=True
        )
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
def cached_bing_search(keyword: str,
                       start_date: dt.date,
                       end_date: dt.date,
                       wilayah_term: str = "") -> List[Dict]:
    from bs4 import BeautifulSoup
    import urllib.parse

    all_entries: List[Dict] = []
    errors: List[str] = []
    seen_urls: Set[str] = set()

    # Build query - natural search without restrictive quotes
    if wilayah_term:
        q_base = f'{keyword} {wilayah_term}'.strip()
    else:
        q_base = keyword

    # Headers to mimic real browser
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'id-ID,id;q=0.9,en;q=0.8',
    }

    # Use Bing News - more reliable than Google News
    try:
        encoded_query = urllib.parse.quote(q_base)
        url = f'https://www.bing.com/news/search?q={encoded_query}&setlang=id-ID'

        for attempt in range(3):
            try:
                response = requests.get(url, headers=headers, timeout=15)
                if response.status_code == 200:
                    soup = BeautifulSoup(response.text, 'html.parser')

                    # Bing News article selectors
                    articles = soup.select('div.news-card, div.newsitem')

                    for article in articles:
                        try:
                            # Extract title and link
                            title_elem = article.select_one('a.title, h3 a, .title a')
                            if not title_elem:
                                continue

                            title = title_elem.text.strip()
                            link = title_elem.get('href', '')

                            if not link or link in seen_urls:
                                continue

                            # Selalu baca teks elemen sumber Bing — untuk ekstrak tanggal relatif (misal: Detik5j)
                            src_elem = article.select_one('.source, .provider, cite')
                            src_raw = src_elem.text.strip() if src_elem else ""

                            # Sumber: utamakan dari URL jika domain ada di prelist
                            src_from_url = source_from_url(link)
                            if src_from_url != "-":
                                src = src_from_url
                            else:
                                # Domain tidak dikenal — ambil dan bersihkan dari HTML Bing
                                src, _ = clean_source_and_date(src_raw, "")
                                if src == "-":
                                    src = src_from_url

                            # Tanggal: cek atribut datetime dulu, lalu teks elemen date, lalu teks elemen sumber
                            published = ""
                            date_elem = article.select_one('time, .date, .timestamp')
                            if date_elem:
                                published = (date_elem.get('datetime', '') or
                                             date_elem.get('data-content', '') or
                                             date_elem.text.strip())

                            # Fallback: waktu relatif dari teks elemen sumber (misal "Detik5j" → pub="5j")
                            if not published and src_raw:
                                _, published = clean_source_and_date(src_raw, "")

                            seen_urls.add(link)
                            all_entries.append({
                                "title": title,
                                "published": published,
                                "link": link,
                                "source": src
                            })
                        except Exception:
                            continue

                    break  # Success, exit retry loop

                else:
                    errors.append(f"Bing News HTTP {response.status_code}")

            except Exception as exc:
                if attempt == 2:  # Last attempt
                    errors.append(f"Bing News: {type(exc).__name__}: {exc}")
                time.sleep(2 ** attempt)

        time.sleep(DELAY_REQ)

    except Exception as exc:
        errors.append(f"Search error: {type(exc).__name__}: {exc}")

    return all_entries, errors


def _strip_title_source_suffix(title: str, source_name: str) -> str:
    """
    RSS Google News selalu memformat title sebagai 'Judul Artikel - Nama Media'.
    Karena nama media sudah ada terpisah di kolom source, suffix ini dibuang
    dari judul agar tidak dobel/redundan.
    """
    if not title or not source_name or source_name == "-":
        return title
    suffix = f" - {source_name}"
    if title.lower().endswith(suffix.lower()):
        return title[: -len(suffix)].strip()
    return title


_GNEWS_RSS_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept-Language': 'id-ID,id;q=0.9,en;q=0.8',
}

def _fetch_gnews_rss_day(q_base: str, day: dt.date) -> Tuple[List[Dict], Optional[str]]:
    """Ambil 1 hari RSS Google News untuk 1 query. Return (entries, error_message)."""
    next_day = day + dt.timedelta(days=1)
    query = f'{q_base} after:{day.strftime("%Y-%m-%d")} before:{next_day.strftime("%Y-%m-%d")}'
    url = ("https://news.google.com/rss/search?"
           f"q={urllib.parse.quote(query)}&hl=id&gl=ID&ceid=ID:id")

    entries: List[Dict] = []
    last_exc = None
    for attempt in range(3):
        try:
            resp = requests.get(url, headers=_GNEWS_RSS_HEADERS, timeout=15)
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
                        try:
                            s = getattr(e, "source", None)
                            if isinstance(s, dict):
                                src = s.get("title", "-") or "-"
                        except Exception:
                            pass

                    # Bersihkan suffix " - Nama Media" dari judul SEBELUM src
                    # dimodifikasi oleh clean_source_and_date (agar suffix
                    # yang dicocokkan persis sama dengan yang ditulis Google).
                    title = _strip_title_source_suffix(title, src)

                    src, published = clean_source_and_date(src, published)

                    if link:
                        entries.append({"title": title, "published": published,
                                        "link": link, "source": src})
                return entries, None
            else:
                last_exc = Exception(f"HTTP {resp.status_code}")
        except Exception as exc:
            last_exc = exc
        time.sleep(2 ** attempt)

    return entries, f"{day}: {type(last_exc).__name__}: {last_exc}" if last_exc else None


@st.cache_data(ttl=3600, show_spinner=False)
def cached_google_search(keyword: str,
                          start_date: dt.date,
                          end_date: dt.date,
                          wilayah_term: str = "") -> Tuple[List[Dict], List[str]]:
    """
    Google News search menggunakan RSS (news.google.com/rss/search).
    Karena RSS Google News hanya menampilkan maksimal ±100 artikel per request,
    rentang tanggal dipecah menjadi 1 request RSS per hari per keyword.
    """
    all_entries: List[Dict] = []
    errors: List[str] = []
    q_base = f'{keyword} {wilayah_term}'.strip() if wilayah_term else keyword

    days: List[dt.date] = []
    current = start_date
    while current <= end_date:
        days.append(current)
        current += dt.timedelta(days=1)

    # Request harian dijalankan paralel ringan (per keyword) agar tidak terlalu lambat,
    # namun tetap diberi jeda kecil untuk menghindari blokir.
    with ThreadPoolExecutor(max_workers=min(5, max(1, len(days)))) as ex:
        future_map = {ex.submit(_fetch_gnews_rss_day, q_base, d): d for d in days}
        for fut in as_completed(future_map):
            try:
                entries, err = fut.result()
                all_entries.extend(entries)
                if err:
                    errors.append(err)
            except Exception as exc:
                d = future_map[fut]
                errors.append(f"{d}: {type(exc).__name__}: {exc}")
            time.sleep(DELAY_REQ / 5)  # jeda kecil antar hasil agar tidak terlalu agresif

    time.sleep(DELAY_REQ)
    return all_entries, errors


# (Fungsi DDG Dihapus)


@st.cache_data(ttl=24*3600, show_spinner=False)
def decode_url_once(link: str) -> str:
    """Decode Google News link dengan batas waktu 10 detik."""
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

# ============================================================
# 5. MAIN SCRAPER FUNCTION
# ============================================================

def jalankan_scraper(
    kata_kunci: Dict[str, List[str]],
    kategori_dict: Dict[str, List[str]],
    custom_kw_list: List[str],
    kab_items, prov_items,
    pos_kws: List[str],
    neg_kws: List[str],
    selected_cats: List[str],
    start_date: dt.date,
    end_date: dt.date,
    wilayah_term: str = "",
    per_kw_limit: int = 0,
    decode_url: bool = True,
    fetch_artikel: bool = True,
    max_ws: int = 3,
    max_wd: int = 5,
    max_wf: int = 3,
    via_selected: str = "Semua",
):
    # Determine which sources to use
    if via_selected == "Semua":
        sources = ["google", "bing"]
    elif via_selected == "Google News":
        sources = ["google"]
    elif via_selected == "Bing News":
        sources = ["bing"]
    else:
        sources = ["bing"]

    # Build tasks: (keyword, category, source)
    tasks: List[Tuple[str, str, str]] = []
    
    if "Custom Keyword" in selected_cats:
        for kw in custom_kw_list:
            for src in sources:
                tasks.append((kw, "Custom Keyword", src))
    else:
        for cat in selected_cats:
            for kw in kata_kunci.get(cat, []):
                for src in sources:
                    tasks.append((kw, cat, src))

    if not tasks:
        st.warning("Tidak ada kata kunci yang dipilih.")
        return

    progress = st.progress(0.0)
    status   = st.empty()
    status.info(f"🔄 Mempersiapkan {len(tasks)} pencarian ({len(sources)} sumber, {max_ws} worker)...")

    exclusion_list = load_exclusion_list()

    # ── Step 1: Search (paralel) ─────────────────────────────
    done = 0
    by_link: Dict[str, Dict] = {}
    all_search_errors: List[str] = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        future_map = {}
        for kw, cat, src in tasks:
            if src == "google":
                fut = ex.submit(cached_google_search, kw, start_date, end_date, wilayah_term)
            elif src == "bing":
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

            # Apply per-keyword limit AFTER cache lookup
            if per_kw_limit > 0:
                entries = entries[:per_kw_limit]

            # Dedup by normalized URL (strip query params)
            for e in entries:
                raw_link = e.get("link", "") or ""
                if is_excluded(raw_link, exclusion_list):
                    continue
                
                # Filter date range
                pub_date_str = format_tanggal(e.get("published", ""))
                if re.match(r'\d{2}/\d{2}/\d{4}', pub_date_str):
                    try:
                        parsed_d = datetime.strptime(pub_date_str, "%d/%m/%Y").date()
                        # Allow 1 day buffer for timezones
                        if parsed_d < start_date - dt.timedelta(days=1) or parsed_d > end_date + dt.timedelta(days=1):
                            continue
                    except Exception:
                        pass
                
                link = raw_link.split("?")[0]
                if not link:
                    continue
                if link not in by_link:
                    by_link[link] = {
                        "title":     e.get("title", "-"),
                        "published": pub_date_str,
                        "source":    e.get("source", "-"),
                        "cats":      set([cat]),
                        "keywords":  set([kw]),  # Track keywords that found this article
                    }
                else:
                    by_link[link]["cats"].add(cat)
                    by_link[link]["keywords"].add(kw)  # Add keyword to existing entry

            done += 1
            progress.progress(done / max(1, len(tasks)))
            status.write(f"🔎 Pencarian: {done}/{len(tasks)} selesai | "
                         f"{len(by_link)} artikel unik")

    if not by_link:
        progress.empty(); status.empty()
        if all_search_errors:
            with st.expander("⚠️ Error detail (klik untuk lihat)"):
                for err in all_search_errors[:20]:
                    st.code(err)
        st.warning("Tidak ada artikel ditemukan.")
        st.session_state.scraped_data = pd.DataFrame(
            columns=["Tanggal", "Judul", "Sumber", "Wilayah", "Kategori", "Persepsi", "Keywords", "Hashtag", "URL"])
        return

    status.write(f"🔗 Total artikel unik: {len(by_link)}")

    # ── Step 2: Decode URL (paralel, opsional) ────────────────────────────
    gnews_links = list(by_link.keys())
    decoded_map: Dict[str, str] = {}

    if decode_url:
        done = 0
        progress.progress(0.0)
        with ThreadPoolExecutor(max_workers=max_wd) as ex:
            future_map2 = {ex.submit(decode_url_once, ln): ln for ln in gnews_links}
            for fut in as_completed(future_map2, timeout=None):
                ln = future_map2[fut]
                try:
                    decoded_map[ln] = fut.result(timeout=12)  # Max 12 detik per URL
                except Exception:
                    decoded_map[ln] = ln  # Fallback ke link asli jika timeout/error
                done += 1
                progress.progress(done / max(1, len(gnews_links)))
                status.write(f"🔓 Decode URL: {done}/{len(gnews_links)}...")
    else:
        decoded_map = {ln: ln for ln in gnews_links}

    # ── Step 3: Fetch teks artikel (paralel, opsional) ────────────────────
    data_map: Dict[str, Dict[str, str]] = {}
    real_urls = [decoded_map[ln] for ln in gnews_links]

    if fetch_artikel:
        done = 0
        progress.progress(0.0)
        with ThreadPoolExecutor(max_workers=max_wf) as ex:
            future_map3 = {ex.submit(cached_fetch_article_data, url): url for url in real_urls}
            for fut in as_completed(future_map3):
                url = future_map3[fut]
                try:
                    data_map[url] = fut.result()
                except Exception:
                    data_map[url] = {"text": "", "date": ""}
                done += 1
                progress.progress(done / max(1, len(real_urls)))
                status.write(f"📄 Fetch data artikel: {done}/{len(real_urls)}...")
    else:
        data_map = {url: {"text": "", "date": ""} for url in real_urls}

    # ── Step 4: Build records ─────────────────────────────────────────────
    records = []
    for gnews_link, obj in by_link.items():
        real_url = decoded_map.get(gnews_link, gnews_link)
        art_data = data_map.get(real_url, {"text": "", "date": ""})
        art_text = art_data["text"]
        full_text = obj["title"] + " " + art_text

        wilayah  = detect_wilayah(full_text, kab_items, prov_items)
        kategori = detect_kategori(full_text, kategori_dict, obj["cats"])
        persepsi = detect_persepsi(obj["title"], art_text, pos_kws, neg_kws, fetch_artikel)

        # Fallback date dari scrape artikel jika kosong atau "-"
        published = obj["published"]
        if (not published or published == "-") and art_data.get("date"):
            published = format_tanggal(art_data["date"])

        # Keywords: join all keywords that found this article
        keywords_str = ", ".join(sorted(obj["keywords"]))

        # Hashtags: extract from full text
        hashtags = re.findall(r'#\w+', full_text)
        hashtags_str = ", ".join(sorted(set(hashtags))) if hashtags else ""

        records.append({
            "Tanggal":  published,
            "Judul":    obj["title"],
            "Sumber":   obj["source"],
            "Wilayah":  wilayah,
            "Kategori": kategori,
            "Persepsi": persepsi,
            "Keywords": keywords_str,
            "Hashtag":  hashtags_str,
            "URL":      real_url,
        })

    df = pd.DataFrame(records)
    st.session_state.scraped_data = df
    progress.empty(); status.empty()
    st.success(f"✅ Selesai! {len(df)} artikel terproses.")

# ============================================================
# 6. STREAMLIT UI
# ============================================================

# ── CSS ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Biarkan header bawaan Streamlit tetap ada agar menu Settings (titik tiga) bisa diklik,
       tapi buat transparan agar custom header kita di belakangnya terlihat. */
    header[data-testid="stHeader"] { 
        background: transparent !important; 
        z-index: 999999 !important;
    }
    
    div[data-testid="stMarkdownContainer"] p { margin-bottom: 4px !important; }
    div[role="radiogroup"] { margin-top: -12px !important; }
    
    /* Beri jarak lebih atas pada konten agar tidak tertutup sticky header */
    .block-container { padding-top: 100px !important; }
    
    div[data-baseweb="input"], div[data-baseweb="datepicker"],
    div[data-baseweb="select"] > div {
        height: 50px !important; min-height: 38px !important;
        border-radius: 6px !important; 
        padding: 4px 10px !important;
        display: flex; align-items: center;
        font-size: 14px !important; line-height: 1.4 !important;
    }
    
    div.stButton > button {
        background: #2196F3 !important; color: #FFF !important;
        border-radius: 6px !important; border: none; padding: 8px 18px !important;
    }
    div.stButton > button:hover { background: #1565C0 !important; }
    div.stDownloadButton > button {
        background-color: #2196F3; color: white; font-weight: bold;
        border-radius: 8px; padding: 0.5em 1em;
    }
    div.stDownloadButton > button:hover { background-color: #1565C0; }

    /* Sembunyikan toolbar Streamlit (hamburger, github, dll) di mobile */
    @media (max-width: 768px) {
        [data-testid="stToolbar"],
        [data-testid="stToolbarActions"],
        #MainMenu {
            display: none !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# ── Logo & Judul di Top Navigation Pane ──────────────────────────────────
encoded_logo = ""
if os.path.exists(LOGO_PATH):
    with open(LOGO_PATH, "rb") as f:
        encoded_logo = base64.b64encode(f.read()).decode()

logo_html = f'<img src="data:image/png;base64,{encoded_logo}" class="header-logo">' if encoded_logo else ""

st.markdown(f"""
<style>
    .header-logo {{
        height: 60px;
        position: absolute;
        left: 20px;
    }}
    .header-title-container {{
        text-align: center;
        padding: 0 10px;
    }}
    .header-title {{
        font-size: 30px;
        font-weight: bold;
        line-height: 1.1;
    }}
    .header-subtitle {{
        font-size: 14px;
        color: gray;
    }}
    
    @media (max-width: 768px) {{
        .header-logo {{
            display: none !important;
        }}
        .header-title {{
            font-size: 22px !important;
        }}
        .header-subtitle {{
            font-size: 11px !important;
        }}
    }}
</style>
<div style="
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 80px;
    background-color: var(--background-color, #FFF);
    color: var(--text-color, #000);
    z-index: 99999;
    border-bottom: 3px solid #e7dfdd;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
">
    {logo_html}
    <div class="header-title-container">
        <div class="header-title">BERITA LNPRT</div>
        <div class="header-subtitle">Scraper Berita Lembaga Non-Profit yang Melayani Rumah Tangga</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Load referensi ────────────────────────────────────────────────────────
with st.spinner("Memuat data referensi..."):
    kata_kunci = load_kata_kunci()
    kategori_dict = load_kategori_dict()
    pos_kws, neg_kws = load_persepsi()
    kab_items, prov_items, sorted_provs, kab_by_prov = load_wilayah()

semua_kategori = list(kata_kunci.keys())

# ── Filter Section ────────────────────────────────────────────────────────
with st.container(border=True):
    # Baris 1: Kategori, Provinsi, Kab/Kota
    col_kat, col_prov, col_kab = st.columns(3)
    
    with col_kat:
        st.markdown("**Kategori**")
        cat_options = ["Semua", "Custom Keyword"] + semua_kategori
        selected_cat = st.selectbox(
            "Pilih kategori", cat_options,
            label_visibility="collapsed",
            key="selected_cat"
        )
        
    with col_prov:
        st.markdown("**Provinsi**")
        prov_display_list = ["Semua"] + [name for _, name in sorted_provs]
        selected_prov_name = st.selectbox(
            "Pilih provinsi", prov_display_list,
            label_visibility="collapsed",
            key="selected_prov"
        )
        
    with col_kab:
        st.markdown("**Kabupaten/Kota**")
        if selected_prov_name == "Semua":
            kab_display_list = ["—"]
            kab_disabled = True
        else:
            selected_kode_prov = next(
                (k for k, n in sorted_provs if n == selected_prov_name), None
            )
            kab_list = kab_by_prov.get(selected_kode_prov, [])
            kab_display_list = ["Semua"] + [n for n, _ in kab_list]
            kab_disabled = False

        selected_kab_name = st.selectbox(
            "Pilih kab/kota", kab_display_list,
            label_visibility="collapsed",
            key="selected_kab",
            disabled=kab_disabled
        )

    st.markdown("<br/>", unsafe_allow_html=True)
    
    if selected_cat == "Custom Keyword":
        custom_kw_str = st.text_input("📝 Masukkan custom keyword (pisahkan dengan koma)", placeholder="contoh: bansor sumatera, banjir sumatera")
        st.caption("💡 *Pastikan Anda menekan **Enter** di keyboard setelah mengetik agar keyword tersimpan sebelum klik Mulai Scraping.*")
        custom_kw_list = [k.strip() for k in custom_kw_str.split(",") if k.strip()]
    else:
        custom_kw_list = []
        
    if selected_cat == "Semua":
        selected_cats = semua_kategori
    elif selected_cat == "Custom Keyword":
        selected_cats = ["Custom Keyword"]
    else:
        selected_cats = [selected_cat]
        
    # Baris 2: Tanggal, Sumber, Limit
    col_tgl, col_src, col_limit = st.columns(3)
    
    with col_tgl:
        st.markdown("**Periode Tanggal**")
        today         = dt.date.today()
        default_start = today - dt.timedelta(days=30)
        periode = st.date_input(
            "Periode Tanggal",
            label_visibility="collapsed",
            key="Tanggal",
            value=(default_start, today),
            format="DD/MM/YYYY"
        )
        if isinstance(periode, tuple) and len(periode) == 2:
            start_date, end_date = periode
        else:
            st.error("⚠️ Harap pilih rentang tanggal.")
            start_date, end_date = default_start, today

    with col_src:
        st.markdown("**Sumber Berita**")
        _via_options = ["Semua", "Google News", "Bing News"]
        _via_selected = st.selectbox(
            "Sumber Berita",
            _via_options,
            index=0,
            label_visibility="collapsed",
            key="via_select"
        )

    with col_limit:
        st.markdown("**Limit artikel per keyword**")
        _limit_map = {"Tidak Terbatas": 0, "5": 5, "10": 10, "25": 25, "50": 50, "100": 100}
        _limit_label = st.selectbox(
            "Limit", list(_limit_map.keys()),
            label_visibility="collapsed", key="limit_select"
        )
        per_kw_limit = _limit_map[_limit_label]

    st.markdown("<br/>", unsafe_allow_html=True)
    
    # Baris 3: Opsi Checkbox
    st.markdown("**Opsi Tambahan**")
    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
        decode_url_toggle  = st.checkbox("🔓 Decode URL asli", value=True)
    with col_opt2:
        fetch_teks_toggle  = st.checkbox("📄 Fetch teks artikel *(akurat, lebih lambat)*", value=False)

# ── Wilayah term untuk keyword pencarian ─────────────────────────────────
if selected_prov_name == "Semua":
    wilayah_term = ""                          # tidak ada filter wilayah
elif selected_kab_name in ("Semua", "—", None):
    wilayah_term = selected_prov_name          # hanya nama provinsi
else:
    wilayah_term = selected_kab_name           # nama kab/kota saja

# ── Info pencarian ────────────────────────────────────────────────────────
total_kw = sum(len(kata_kunci.get(c, [])) for c in selected_cats)
wilayah_info = f"**{wilayah_term}**" if wilayah_term else "Seluruh Indonesia"
st.caption(
    f"📌 Kategori: **{selected_cat}** | {total_kw} kata kunci | "
    f"Wilayah: {wilayah_info} | Sumber: **{_via_selected}**"
)

# ── Tombol scraping ───────────────────────────────────────────────────────
st.markdown("")
_, col_btn, _ = st.columns([4, 3, 4])
with col_btn:
    scrape_button = st.button("🔍 Mulai Scraping", use_container_width=True)

# ── Init session state ────────────────────────────────────────────────────
if "scraped_data" not in st.session_state:
    st.session_state.scraped_data = pd.DataFrame(
        columns=["Tanggal", "Judul", "Sumber", "Wilayah", "Kategori", "Persepsi", "Keywords", "Hashtag", "URL"])

# ── Jalankan scraper ──────────────────────────────────────────────────────
if scrape_button:
    if not selected_cats or (selected_cat == "Custom Keyword" and not custom_kw_list):
        st.warning("Pilih minimal satu kategori atau masukkan custom keyword.")
    else:
        jalankan_scraper(
            kata_kunci=kata_kunci,
            kategori_dict=kategori_dict,
            custom_kw_list=custom_kw_list,
            kab_items=kab_items,
            prov_items=prov_items,
            pos_kws=pos_kws,
            neg_kws=neg_kws,
            selected_cats=selected_cats,
            start_date=start_date,
            end_date=end_date,
            wilayah_term=wilayah_term,
            per_kw_limit=per_kw_limit,
            decode_url=decode_url_toggle,
            fetch_artikel=fetch_teks_toggle,
            max_ws=5,   # Ditingkatkan dari 3 ke 5 (Pencarian lebih cepat)
            max_wd=15,  # Ditingkatkan dari 5 ke 15 (Dekode link Google super cepat)
            max_wf=8,   # Ditingkatkan dari 3 ke 8 (Download konten artikel lebih paralel)
            via_selected=_via_selected,
        )

# ── Tampilkan hasil ───────────────────────────────────────────────────────
if not st.session_state.scraped_data.empty:
    show_aggrid(st.session_state.scraped_data)
else:
    st.info("Belum ada data. Pilih kategori dan periode, lalu klik **Mulai Scraping**.")
