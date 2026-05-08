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
from pygooglenews import GoogleNews
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

gn         = GoogleNews(lang="id", country="ID")
DATE_DELTA = dt.timedelta(days=30)
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
    
    m_min = re.search(r'(\d+)\s*(?:m|mnt|menit)', text_lower)
    if m_min and 'minggu' not in text_lower:
        return (now - dt.timedelta(minutes=int(m_min.group(1)))).strftime("%d/%m/%Y")
        
    m_hour = re.search(r'(\d+)\s*(?:j|jam|h|hour)', text_lower)
    if m_hour and 'hari' not in text_lower:
        return (now - dt.timedelta(hours=int(m_hour.group(1)))).strftime("%d/%m/%Y")
        
    m_day = re.search(r'(\d+)\s*(?:hari|d|day)', text_lower)
    if m_day:
        return (now - dt.timedelta(days=int(m_day.group(1)))).strftime("%d/%m/%Y")
        
    m_week = re.search(r'(\d+)\s*(?:mgg|minggu|w|week)', text_lower)
    if m_week:
        return (now - dt.timedelta(weeks=int(m_week.group(1)))).strftime("%d/%m/%Y")
        
    m_month = re.search(r'(\d+)\s*(?:bln|bulan|mo|month)s?', text_lower)
    if m_month:
        return (now - dt.timedelta(days=int(m_month.group(1))*30)).strftime("%d/%m/%Y")
        
    m_year = re.search(r'(\d+)\s*(?:thn|tahun|yr|year)s?', text_lower)
    if m_year:
        return (now - dt.timedelta(days=int(m_year.group(1))*365)).strftime("%d/%m/%Y")
        
    return text

def clean_source_and_date(src: str, pub: str) -> Tuple[str, str]:
    if not src: src = "-"
    if not pub: pub = ""
    
    # Pisahkan angka di akhir string source. Misal: Tempo.co1h, TribunNews2thn
    m = re.search(r'(\d+\s*(?:j|jam|h|hour|m|mnt|menit|hari|d|day|mgg|minggu|w|week|bln|bulan|mo|months?|thn|tahun|yr|years?)(?:\s+yang\s+lalu|\s+ago)?)$', src, flags=re.IGNORECASE)
    if m:
        time_str = m.group(1)
        src = src[:m.start()].strip()
        if not pub or pub == "-":
            pub = time_str
            
    # Hapus suffix "on MSN"
    m_msn = re.search(r'\s+on\s+MSN$', src, flags=re.IGNORECASE)
    if m_msn:
        src = src[:m_msn.start()].strip()
            
    if not src: src = "-"
    return src, pub

def format_tanggal(published: str) -> str:
    if not published or published == "-":
        return "-"
        
    rel_date = parse_relative_time_to_date(published)
    if re.match(r'\d{2}/\d{2}/\d{4}', rel_date):
        return rel_date
        
    for fmt in ("%a, %d %b %Y %H:%M:%S %Z", "%a, %d %b %Y %H:%M:%S %z", "%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%d"):
        try:
            return datetime.strptime(published.strip(), fmt).strftime("%d/%m/%Y")
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


def fetch_article_text(url: str) -> str:
    try:
        import trafilatura
        html = trafilatura.fetch_url(url)
        if html:
            return trafilatura.extract(html) or ""
    except Exception:
        pass
    return ""

# ============================================================
# 3. EXPORT & DISPLAY
# ============================================================

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

                            # Get source
                            src_elem = article.select_one('.source, .provider, cite')
                            src = src_elem.text.strip() if src_elem else "-"

                            # Get published date
                            date_elem = article.select_one('.date, .timestamp, time')
                            published = date_elem.text.strip() if date_elem else ""

                            src, published = clean_source_and_date(src, published)

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


@st.cache_data(ttl=3600, show_spinner=False)
def cached_google_search(keyword: str,
                          start_date: dt.date,
                          end_date: dt.date,
                          wilayah_term: str = "") -> List[Dict]:
    """Google News search menggunakan pygooglenews"""
    all_entries: List[Dict] = []
    errors: List[str] = []
    current = start_date
    q_base = f'{keyword} {wilayah_term}'.strip() if wilayah_term else keyword

    while current < end_date:
        batch_end = min(current + DATE_DELTA, end_date)
        last_exc = None
        for attempt in range(3):
            try:
                hasil = gn.search(
                    q_base,
                    from_=current.strftime("%Y-%m-%d"),
                    to_=batch_end.strftime("%Y-%m-%d")
                )
                for e in hasil.get("entries", []):
                    title     = getattr(e, "title",     None) or e.get("title",     "-") or "-"
                    published = getattr(e, "published", None) or e.get("published", "") or ""
                    link      = getattr(e, "link",      None) or e.get("link",      "") or ""
                    src = "-"
                    try:
                        src = e.source.title
                    except Exception:
                        try:
                            s = e.get("source", None)
                            if isinstance(s, dict):
                                src = s.get("title", "-") or "-"
                        except Exception:
                            pass
                            
                    src, published = clean_source_and_date(src, published)
                    
                    if link:
                        all_entries.append({"title": title, "published": published,
                                            "link": link, "source": src})
                last_exc = None
                break
            except Exception as exc:
                last_exc = exc
                time.sleep(2 ** attempt)
        if last_exc:
            errors.append(f"{current}→{batch_end}: {type(last_exc).__name__}: {last_exc}")
        current = batch_end
        time.sleep(0.5)
    return all_entries, errors


# (Fungsi DDG Dihapus)


@st.cache_data(ttl=24*3600, show_spinner=False)
def decode_url_once(link: str) -> str:
    try:
        r = gnewsdecoder(link)
        return r["decoded_url"] if r.get("status") else link
    except Exception:
        return link


@st.cache_data(ttl=3600, show_spinner=False)
def cached_fetch_text(url: str) -> str:
    return fetch_article_text(url)

# ============================================================
# 5. MAIN SCRAPER FUNCTION
# ============================================================

def jalankan_scraper(
    kata_kunci: Dict[str, List[str]],
    kategori_dict: Dict[str, List[str]],
    custom_kw_list: List[str],
    kab_items, prov_items,
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
            columns=["Tanggal", "Judul", "Sumber", "Wilayah", "Kategori", "Keywords", "Hashtag", "URL"])
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
            for fut in as_completed(future_map2):
                ln = future_map2[fut]
                try:
                    decoded_map[ln] = fut.result()
                except Exception:
                    decoded_map[ln] = ln
                done += 1
                progress.progress(done / max(1, len(gnews_links)))
                status.write(f"🔓 Decode URL: {done}/{len(gnews_links)}...")
    else:
        decoded_map = {ln: ln for ln in gnews_links}

    # ── Step 3: Fetch teks artikel (paralel, opsional) ────────────────────
    text_map: Dict[str, str] = {}
    real_urls = [decoded_map[ln] for ln in gnews_links]

    if fetch_artikel:
        done = 0
        progress.progress(0.0)
        with ThreadPoolExecutor(max_workers=max_wf) as ex:
            future_map3 = {ex.submit(cached_fetch_text, url): url for url in real_urls}
            for fut in as_completed(future_map3):
                url = future_map3[fut]
                try:
                    text_map[url] = fut.result()
                except Exception:
                    text_map[url] = ""
                done += 1
                progress.progress(done / max(1, len(real_urls)))
                status.write(f"📄 Fetch teks artikel: {done}/{len(real_urls)}...")
    else:
        text_map = {url: "" for url in real_urls}

    # ── Step 4: Build records ─────────────────────────────────────────────
    records = []
    for gnews_link, obj in by_link.items():
        real_url = decoded_map.get(gnews_link, gnews_link)
        art_text = text_map.get(real_url, "")
        full_text = obj["title"] + " " + art_text

        wilayah  = detect_wilayah(full_text, kab_items, prov_items)
        kategori = detect_kategori(full_text, kategori_dict, obj["cats"])

        # Keywords: join all keywords that found this article
        keywords_str = ", ".join(sorted(obj["keywords"]))

        # Hashtags: extract from full text
        hashtags = re.findall(r'#\w+', full_text)
        hashtags_str = ", ".join(sorted(set(hashtags))) if hashtags else ""

        records.append({
            "Tanggal":  obj["published"],
            "Judul":    obj["title"],
            "Sumber":   obj["source"],
            "Wilayah":  wilayah,
            "Kategori": kategori,
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
</style>
""", unsafe_allow_html=True)

# ── Logo & Judul di Top Navigation Pane ──────────────────────────────────
encoded_logo = ""
if os.path.exists(LOGO_PATH):
    with open(LOGO_PATH, "rb") as f:
        encoded_logo = base64.b64encode(f.read()).decode()

logo_html = f'<img src="data:image/png;base64,{encoded_logo}" style="height: 60px; position: absolute; left: 20px;">' if encoded_logo else ""

st.markdown(f"""
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
    <div style="text-align: center;">
        <div style="font-size: 30px; font-weight: bold; line-height: 1.1;">BERITA LNPRT</div>
        <div style="font-size: 14px; color: gray;">Scraper Berita Lembaga Non-Profit yang Melayani Rumah Tangga</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Load referensi ────────────────────────────────────────────────────────
with st.spinner("Memuat data referensi..."):
    kata_kunci = load_kata_kunci()
    kategori_dict = load_kategori_dict()
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
        columns=["Tanggal", "Judul", "Sumber", "Wilayah", "Kategori", "Keywords", "Hashtag", "URL"])

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
