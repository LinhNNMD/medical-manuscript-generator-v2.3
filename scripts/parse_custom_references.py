"""
parse_custom_references.py — Phân tích tài liệu PDF/DOCX tự chọn, cào trích dẫn con và gộp cơ sở dữ liệu.
"""
import io
import os
import re
import sys
import json
import time
import logging
import argparse
import xml.etree.ElementTree as ET
import requests
import random
import urllib.parse
from datetime import datetime

# Global state to track last request timestamp per domain
LAST_REQUEST_TIME = {}

def safe_get(url, **kwargs):
    """
    Wrapper for requests.get with rate limiting, randomized jitter,
    and HTTP 429 error handling with exponential backoff.
    """
    parsed_url = urllib.parse.urlparse(url)
    domain = parsed_url.netloc
    
    # Base delay of 1.0 second per domain
    min_delay = 1.0
    
    # Apply rate limiting logic
    last_time = LAST_REQUEST_TIME.get(domain, 0)
    now = time.time()
    elapsed = now - last_time
    
    # Add randomized jitter to standard delay (0.1 to 0.5 seconds)
    jitter = random.uniform(0.1, 0.5)
    required_delay = min_delay + jitter
    
    if elapsed < required_delay:
        sleep_time = required_delay - elapsed
        logger.debug(f"[Rate Limiter] Sleeping {sleep_time:.2f}s for domain {domain}")
        time.sleep(sleep_time)
        
    max_retries = 5
    base_backoff = 2.0  # seconds
    
    for attempt in range(max_retries + 1):
        try:
            # Update last request time before request starts
            LAST_REQUEST_TIME[domain] = time.time()
            
            response = requests.get(url, **kwargs)
            
            # Handle HTTP 429 (Too Many Requests)
            if response.status_code == 429:
                if attempt == max_retries:
                    logger.error(f"[HTTP 429] Max retries reached for URL: {url}")
                    return response
                
                # Check for Retry-After header
                retry_after = response.headers.get("Retry-After")
                wait_seconds = 0
                if retry_after:
                    try:
                        wait_seconds = float(retry_after)
                        logger.warning(f"[HTTP 429] Received Retry-After: {retry_after}s for {url}")
                    except ValueError:
                        pass
                
                if wait_seconds <= 0:
                    # Exponential backoff + randomized jitter
                    wait_seconds = (base_backoff * (2 ** attempt)) + random.uniform(0.5, 2.0)
                
                logger.warning(f"[HTTP 429] Rate limited. Attempt {attempt + 1}/{max_retries + 1}. Waiting {wait_seconds:.2f}s to retry {url}")
                time.sleep(wait_seconds)
                continue
                
            return response
            
        except requests.exceptions.RequestException as e:
            if attempt == max_retries:
                logger.error(f"[Request Error] Max retries reached after exception: {e} for URL: {url}")
                raise e
            
            wait_seconds = (base_backoff * (2 ** attempt)) + random.uniform(0.5, 2.0)
            logger.warning(f"[Request Error] {e}. Attempt {attempt + 1}/{max_retries + 1}. Retrying in {wait_seconds:.2f}s...")
            time.sleep(wait_seconds)

# Tránh lỗi ModuleNotFoundError khi chạy thư viện ngoài
try:
    import pypdf
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf"])
    import pypdf

try:
    import docx
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
    import docx

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("parse_custom")

# Configs từ fetch_pubmed.py
PUBMED_EMAIL = "nguyenngoclinh.vmmu@gmail.com"
NCBI_API_KEY = "b8a896c1834a8062a17d6390e0cfaf04ac08"

# ── Hướng dẫn Regex ──────────────────────────────────────────────────────────
# Dò quét DOI (được bắt đầu bằng 10. và theo sau bởi ký tự hợp lệ)
DOI_REGEX = re.compile(r'\b(10\.\d{4,9}/[-._;()/:A-Z0-9]+)\b', re.IGNORECASE)
# Dò quét PMID (dạng PMID: 123456 hoặc PMID 123456)
PMID_REGEX = re.compile(r'\bPMID\s*[:\s]\s*(\d{5,10})\b', re.IGNORECASE)


def normalize_filename(filename):
    """Chuẩn hóa tên file thành ID không dấu tiếng Việt và ký tự đặc biệt."""
    # Bỏ phần mở rộng
    name, _ = os.path.splitext(filename)
    # Loại bỏ dấu và ký tự lạ
    name = re.sub(r'[^\w\s-]', '', name)
    name = re.sub(r'[\s-]+', '_', name)
    return name.lower()


def extract_text_from_pdf(pdf_path):
    """Trích xuất văn bản từ tệp PDF."""
    try:
        reader = pypdf.PdfReader(pdf_path)
        text_parts = []
        for i, page in enumerate(reader.pages):
            page_text = page.extract_text()
            if page_text:
                text_parts.append(page_text)
        
        # Lấy metadata cơ bản
        metadata = reader.metadata or {}
        title = metadata.get('/Title', '')
        
        return "\n".join(text_parts), title
    except Exception as e:
        logger.error(f"Lỗi đọc PDF {pdf_path}: {e}")
        return "", ""


def extract_text_from_docx(docx_path):
    """Trích xuất văn bản từ tệp DOCX."""
    try:
        doc = docx.Document(docx_path)
        paragraphs = [p.text for p in doc.paragraphs if p.text]
        # Lấy tiêu đề từ đoạn đầu tiên nếu có
        title = paragraphs[0] if paragraphs else ""
        return "\n".join(paragraphs), title
    except Exception as e:
        logger.error(f"Lỗi đọc DOCX {docx_path}: {e}")
        return "", ""


def parse_ris_file(file_path, folder_path="custom_references"):
    """Phân tích cú pháp tệp .ris và trích xuất siêu dữ liệu CSL JSON."""
    items = []
    current_item = {}
    authors = []
    
    try:
        with open(file_path, "r", encoding="utf-8-sig") as f:
            lines = f.readlines()
    except Exception as e:
        logger.error(f"Không thể đọc file RIS {file_path}: {e}")
        return items

    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        match = re.match(r'^([A-Z0-9]{2})\s*-\s*(.*)$', line)
        if match:
            key = match.group(1).upper()
            val = match.group(2).strip()
            
            if key == "TY":
                current_item = {
                    "type": "article-journal"
                }
                authors = []
            elif key == "ER":
                if current_item:
                    if authors:
                        current_item["author"] = authors
                    
                    doi = current_item.get("DOI")
                    pmid = current_item.get("PMID")
                    
                    if pmid:
                        csl_id = f"PMID_{pmid}"
                    elif doi:
                        csl_id = f"DOI_{doi.replace('/', '_').replace('.', '_')}"
                    else:
                        title_slug = re.sub(r'[^\w]', '_', current_item.get("title", ""))[:30].lower()
                        csl_id = f"custom_{title_slug}"
                        
                    current_item["id"] = csl_id
                    
                    # Thử tải toàn văn
                    citation_key = csl_id.replace("PMID_", "")
                    full_text_path = None
                    if pmid:
                        full_text_path = fetch_full_text_from_pmc(pmid, citation_key, folder_path)
                    if not full_text_path and doi:
                        full_text_path = fetch_full_text_from_unpaywall(doi, citation_key, folder_path)
                        
                    if full_text_path:
                        current_item["full_text_file"] = full_text_path
                        
                    items.append(current_item)
                    current_item = {}
                    authors = []
            else:
                if not current_item:
                    continue
                    
                if key in ["TI", "T1"]:
                    current_item["title"] = val
                elif key in ["AU", "A1"]:
                    parts = val.split(",", 1)
                    if len(parts) == 2:
                        authors.append({"family": parts[0].strip(), "given": parts[1].strip()})
                    else:
                        authors.append({"family": val})
                elif key in ["JO", "JF", "T2"]:
                    current_item["container-title"] = val
                elif key in ["PY", "Y1"]:
                    year_match = re.match(r'^(\d{4})', val)
                    if year_match:
                        current_item["issued"] = {"date-parts": [[int(year_match.group(1))]]}
                elif key == "VL":
                    current_item["volume"] = val
                elif key == "IS":
                    current_item["issue"] = val
                elif key == "SP":
                    current_item["page"] = val
                elif key == "EP":
                    if "page" in current_item:
                        current_item["page"] = f"{current_item['page']}-{val}"
                    else:
                        current_item["page"] = val
                elif key == "DO":
                    current_item["DOI"] = val
                elif key == "AN":
                    pmid_match = re.search(r'\b(?:PMID\s*[:\s]\s*)?(\d{5,10})\b', val, re.IGNORECASE)
                    if pmid_match:
                        current_item["PMID"] = pmid_match.group(1)
                elif key in ["AB", "N2"]:
                    current_item["abstract"] = val
                    
    return items


def fetch_metadata_by_doi(doi):
    """Truy vấn siêu dữ liệu của một DOI sử dụng OpenAlex API và CrossRef fallback."""
    doi = doi.strip().replace("https://doi.org/", "")
    logger.info(f"Đang tìm kiếm thông tin DOI: {doi}")
    
    # 1. Thử với OpenAlex
    url = f"https://api.openalex.org/works/https://doi.org/{doi}"
    try:
        response = safe_get(url, timeout=15)
        if response.status_code == 200:
            item = response.json()
            authors = []
            for a in item.get("authorships", []):
                name = a.get("author", {}).get("display_name", "")
                if name:
                    parts = name.split(" ")
                    if len(parts) > 1:
                        authors.append({"family": parts[-1], "given": " ".join(parts[:-1])})
                    else:
                        authors.append({"family": name})
            
            year = item.get("publication_year")
            pmid = item.get("ids", {}).get("pmid", "").replace("https://pubmed.ncbi.nlm.nih.gov/", "")
            
            csl_item = {
                "id": f"PMID_{pmid}" if pmid else f"DOI_{doi.replace('/', '_').replace('.', '_')}",
                "type": "article-journal",
                "title": item.get("display_name", ""),
                "author": authors,
                "issued": {"date-parts": [[int(year)]]} if year else {},
                "DOI": doi,
                "container-title": item.get("host_venue", {}).get("display_name", "") or item.get("locations", [{}])[0].get("source", {}).get("display_name", ""),
                "source": "OpenAlex"
            }
            if pmid:
                csl_item["PMID"] = pmid
            return csl_item
    except Exception as e:
        logger.debug(f"OpenAlex lookup failed for DOI {doi}: {e}")

    # 2. CrossRef Fallback
    url = f"https://api.crossref.org/works/{doi}"
    try:
        response = safe_get(url, headers={"User-Agent": f"mailto:{PUBMED_EMAIL}"}, timeout=15)
        if response.status_code == 200:
            data = response.json().get("message", {})
            authors = []
            for a in data.get("author", []):
                authors.append({
                    "family": a.get("family", ""),
                    "given": a.get("given", "")
                })
            
            year = None
            date_parts = data.get("published-print", {}).get("date-parts", []) or data.get("published-online", {}).get("date-parts", [])
            if date_parts:
                year = date_parts[0][0]
                
            csl_item = {
                "id": f"DOI_{doi.replace('/', '_').replace('.', '_')}",
                "type": "article-journal",
                "title": data.get("title", [""])[0],
                "author": authors,
                "issued": {"date-parts": [[int(year)]]} if year else {},
                "DOI": doi,
                "container-title": data.get("container-title", [""])[0],
                "source": "CrossRef"
            }
            return csl_item
    except Exception as e:
        logger.warning(f"CrossRef lookup failed for DOI {doi}: {e}")

    return None


def fetch_pubmed_by_ids(pmids):
    """Truy vấn hàng loạt PMID từ PubMed E-utilities API."""
    if not pmids:
        return []
    
    url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
    params = {
        "db": "pubmed",
        "id": ",".join(pmids),
        "retmode": "xml",
        "tool": "medical_manuscript_generator",
        "email": PUBMED_EMAIL,
    }
    if NCBI_API_KEY:
        params["api_key"] = NCBI_API_KEY

    results = []
    try:
        response = safe_get(url, params=params, timeout=20)
        response.raise_for_status()
        root = ET.fromstring(response.content)
        
        for article in root.findall(".//PubmedArticle"):
            pmid = article.findtext(".//PMID")
            title = article.findtext(".//ArticleTitle") or ""
            pub_year = article.findtext(".//PubDate/Year") or article.findtext(".//ArticleDate/Year") or "2026"
            journal = article.findtext(".//Journal/Title") or article.findtext(".//MedlineTA") or ""
            volume = article.findtext(".//JournalIssue/Volume") or ""
            issue = article.findtext(".//JournalIssue/Issue") or ""
            pages = article.findtext(".//Pagination/MedlinePgn") or ""
            
            authors = []
            for author in article.findall(".//AuthorList/Author"):
                last_name = author.findtext("LastName")
                fore_name = author.findtext("ForeName") or author.findtext("Initials")
                if last_name:
                    author_dict = {"family": last_name}
                    if fore_name:
                        author_dict["given"] = fore_name
                    authors.append(author_dict)
            
            doi = ""
            for aid in article.findall(".//ArticleIdList/ArticleId"):
                if aid.get("IdType") == "doi":
                    doi = aid.text
                    break

            abstract_elements = article.findall(".//AbstractText")
            abstract_text = " ".join([elem.text for elem in abstract_elements if elem.text])

            csl_item = {
                "id": f"PMID_{pmid}",
                "type": "article-journal",
                "title": title,
                "container-title": journal,
                "issued": {"date-parts": [[int(pub_year)]]},
                "source": "pubmed",
                "abstract": abstract_text or "Không có tóm tắt."
            }
            if authors:
                csl_item["author"] = authors
            if volume:
                csl_item["volume"] = volume
            if issue:
                csl_item["issue"] = issue
            if pages:
                csl_item["page"] = pages
            if doi:
                csl_item["DOI"] = doi
                
            results.append(csl_item)
    except Exception as e:
        logger.error(f"Lỗi fetch PMID list {pmids}: {e}")
        
    return results


def fetch_full_text_from_pmc(pmid, citation_key, folder_path="custom_references"):
    """Chuyển đổi PMID sang PMCID và tải toàn văn từ PMC bằng E-utilities elink."""
    conv_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/elink.fcgi"
    params = {
        "dbfrom": "pubmed",
        "db": "pmc",
        "id": pmid,
        "linkname": "pubmed_pmc",
        "retmode": "json",
        "tool": "medical_manuscript_generator",
        "email": PUBMED_EMAIL,
    }
    if NCBI_API_KEY:
        params["api_key"] = NCBI_API_KEY
        
    try:
        resp = safe_get(conv_url, params=params, timeout=15)
        if resp.status_code == 200:
            data = resp.json()
            linksets = data.get("linksets", [])
            if linksets:
                linksetdbs = linksets[0].get("linksetdbs", [])
                if linksetdbs:
                    links = linksetdbs[0].get("links", [])
                    if links:
                        pmcid_num = links[0]
                        pmcid = f"PMC{pmcid_num}"
                        logger.info(f"Tìm thấy PMCID: {pmcid} cho PMID: {pmid}")
                        fetch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
                        fetch_params = {
                            "db": "pmc",
                            "id": pmcid,
                            "retmode": "xml",
                            "tool": "medical_manuscript_generator",
                            "email": PUBMED_EMAIL,
                        }
                        if NCBI_API_KEY:
                            fetch_params["api_key"] = NCBI_API_KEY
                        
                        xml_resp = safe_get(fetch_url, params=fetch_params, timeout=20)
                        if xml_resp.status_code == 200:
                            root = ET.fromstring(xml_resp.content)
                            body_elem = root.find(".//body")
                            if body_elem is not None:
                                paragraphs = []
                                for p in body_elem.findall(".//p"):
                                    p_text = "".join(p.itertext()).strip()
                                    if p_text:
                                        paragraphs.append(p_text)
                                
                                full_text = "\n\n".join(paragraphs)
                                if full_text:
                                    full_text_dir = os.path.join(folder_path, "scraped_full_texts")
                                    os.makedirs(full_text_dir, exist_ok=True)
                                    full_text_path = os.path.join(full_text_dir, f"{citation_key}.txt")
                                    with open(full_text_path, "w", encoding="utf-8") as f:
                                        f.write(full_text)
                                    logger.info(f"Đã lưu toàn văn bài báo PMC {pmcid} tại {full_text_path}")
                                    return f"custom_references/scraped_full_texts/{citation_key}.txt"
    except Exception as e:
        logger.warning(f"Lỗi tải full-text qua PMC cho PMID {pmid}: {e}")
    return None


def fetch_full_text_from_unpaywall(doi, citation_key, folder_path="custom_references"):
    """Thử tải full-text từ Unpaywall API và trích xuất text từ PDF Open Access."""
    doi = doi.strip().replace("https://doi.org/", "")
    url = f"https://api.unpaywall.org/v2/{doi}?email={PUBMED_EMAIL}"
    try:
        response = safe_get(url, timeout=15)
        if response.status_code == 200:
            data = response.json()
            if data.get("is_oa"):
                best_loc = data.get("best_oa_location") or {}
                pdf_url = best_loc.get("url_for_pdf")
                if not pdf_url and data.get("oa_locations"):
                    for loc in data.get("oa_locations", []):
                        if loc.get("url_for_pdf"):
                            pdf_url = loc.get("url_for_pdf")
                            break
                
                if pdf_url:
                    logger.info(f"Tìm thấy PDF Open Access cho DOI {doi}: {pdf_url}")
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                        "Accept": "application/pdf,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
                        "Accept-Language": "en-US,en;q=0.5",
                        "Referer": "https://api.unpaywall.org/",
                    }
                    pdf_resp = safe_get(pdf_url, timeout=30, headers=headers)
                    if pdf_resp.status_code == 200:
                        temp_pdf_path = os.path.join(folder_path, f"temp_{citation_key}.pdf")
                        with open(temp_pdf_path, "wb") as f:
                            f.write(pdf_resp.content)
                        
                        full_text, _ = extract_text_from_pdf(temp_pdf_path)
                        
                        if os.path.exists(temp_pdf_path):
                            os.remove(temp_pdf_path)
                            
                        if full_text:
                            full_text_dir = os.path.join(folder_path, "scraped_full_texts")
                            os.makedirs(full_text_dir, exist_ok=True)
                            full_text_path = os.path.join(full_text_dir, f"{citation_key}.txt")
                            with open(full_text_path, "w", encoding="utf-8") as f:
                                f.write(full_text)
                            logger.info(f"Đã lưu toàn văn bài báo DOI {doi} tại {full_text_path}")
                            return f"custom_references/scraped_full_texts/{citation_key}.txt"
                    else:
                        logger.warning(f"Không thể tải PDF từ {pdf_url}. Status code: {pdf_resp.status_code}")
    except Exception as e:
        logger.warning(f"Lỗi tải full-text qua Unpaywall cho DOI {doi}: {e}")
    return None


def process_custom_folder(folder_path="custom_references", db_path="references_db.json"):
    """Quét và xử lý thư mục tài liệu tự chọn."""
    if not os.path.exists(folder_path):
        os.makedirs(folder_path, exist_ok=True)
        logger.info(f"Đã tạo thư mục trống {folder_path}. Chưa có tệp nào để quét.")
        return

    files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.pdf', '.docx', '.ris'))]
    if not files:
        logger.info(f"Thư mục '{folder_path}' trống hoặc không chứa file PDF/DOCX/RIS.")
        return

    logger.info(f"Tìm thấy {len(files)} tài liệu tự chọn cần xử lý.")
    
    # Đọc db hiện có
    existing_db = []
    if os.path.exists(db_path):
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                existing_db = json.load(f)
        except Exception:
            existing_db = []
            
    existing_dois = {item.get("DOI", "").lower().strip() for item in existing_db if item.get("DOI")}
    existing_pmids = {item.get("PMID", "").strip() for item in existing_db if item.get("PMID")}
    existing_ids = {item.get("id") for item in existing_db}

    # Báo cáo ngữ cảnh cuối cùng
    reports = []
    
    # Tập hợp các PMID/DOI con cần cào
    child_pmids = set()
    child_dois = set()
    
    new_csl_items = []

    for file in files:
        file_path = os.path.join(folder_path, file)
        logger.info(f"📄 Đang xử lý file: {file}")
        
        if file.lower().endswith('.ris'):
            # Xử lý tệp RIS
            parsed_items = parse_ris_file(file_path, folder_path)
            for item in parsed_items:
                item_id = item["id"]
                item_doi = item.get("DOI", "").lower().strip()
                item_pmid = item.get("PMID", "").strip()
                
                is_dup = item_id in existing_ids or (
                    item_doi and item_doi in existing_dois
                ) or (
                    item_pmid and item_pmid in existing_pmids
                )
                
                if not is_dup:
                    new_csl_items.append(item)
                    existing_db.append(item)
                    existing_ids.add(item_id)
                    if item_doi:
                        existing_dois.add(item_doi)
                    if item_pmid:
                        existing_pmids.add(item_pmid)
                        
                reports.append({
                    "filename": file,
                    "title": item.get("title", file),
                    "key": item_id.replace("PMID_", ""),
                    "is_standard": True,
                    "summary": item.get("abstract", "Không có tóm tắt.")[:800],
                    "children_found": 0,
                    "full_text_file": item.get("full_text_file")
                })
            continue

        # 1. Trích xuất text cho PDF/DOCX cục bộ
        if file.lower().endswith('.pdf'):
            text, extracted_title = extract_text_from_pdf(file_path)
        elif file.lower().endswith('.docx'):
            text, extracted_title = extract_text_from_docx(file_path)
        else:
            continue
            
        if not text:
            logger.warning(f"File {file} không có nội dung văn bản.")
            continue

        # 2. Tìm metadata của bản thân tài liệu PDF/DOCX
        doc_doi = None
        doc_pmid = None
        is_standard = False
        
        # Dò trong 500 ký tự đầu tiên
        header_text = text[:500]
        doi_matches = DOI_REGEX.findall(header_text)
        if doi_matches:
            doc_doi = doi_matches[0]
        pmid_matches = PMID_REGEX.findall(header_text)
        if pmid_matches:
            doc_pmid = pmid_matches[0]

        normalized_key = normalize_filename(file)
        custom_citation_key = f"custom_{normalized_key}"
        csl_id = f"PMID_{custom_citation_key}"
        
        csl_metadata = None
        
        # Nếu có DOI hoặc PMID, thử fetch chuẩn hóa
        if doc_pmid:
            pubmed_results = fetch_pubmed_by_ids([doc_pmid])
            if pubmed_results:
                csl_metadata = pubmed_results[0]
                csl_id = csl_metadata["id"]
                logger.info(f"  ✓ Đã chuẩn hóa thành công tài liệu theo PMID: {doc_pmid}")
        elif doc_doi:
            csl_metadata = fetch_metadata_by_doi(doc_doi)
            if csl_metadata:
                csl_id = csl_metadata["id"]
                logger.info(f"  ✓ Đã chuẩn hóa thành công tài liệu theo DOI: {doc_doi}")

        is_standard = csl_metadata is not None
        
        # Nếu không có metadata, tạo tài liệu nội bộ (chỉ để tham khảo, không trích dẫn)
        if not csl_metadata:
            title = extracted_title.strip() if extracted_title.strip() else file
            # Giới hạn title không quá dài
            if len(title) > 150:
                title = title[:150] + "..."
            
            csl_metadata = {
                "id": csl_id,
                "type": "article-journal",
                "title": title,
                "author": [{"family": "Tài liệu", "given": "Nội bộ"}],
                "issued": {"date-parts": [[datetime.now().year]]},
                "container-title": "Tài liệu nội bộ",
                "source": "custom_file"
            }
            logger.info(f"  ⚠ Không thể cào metadata cho '{file}'. File này sẽ chỉ dùng để tham khảo nội dung và KHÔNG tạo khóa trích dẫn.")

        # Ghi nhận text thô từ file cục bộ ra tệp .txt trong scraped_full_texts
        citation_key = csl_id.replace("PMID_", "")
        full_text_dir = os.path.join(folder_path, "scraped_full_texts")
        os.makedirs(full_text_dir, exist_ok=True)
        full_text_path = os.path.join(full_text_dir, f"{citation_key}.txt")
        try:
            with open(full_text_path, "w", encoding="utf-8") as f_txt:
                f_txt.write(text)
            csl_metadata["full_text_file"] = f"custom_references/scraped_full_texts/{citation_key}.txt"
        except Exception as e_txt:
            logger.warning(f"Không thể lưu text cục bộ cho {file}: {e_txt}")

        # Lưu thông tin để ghi log
        doc_summary = text[:800].replace("\n", " ").strip() + "..."
        reports.append({
            "filename": file,
            "title": csl_metadata.get("title", file),
            "key": csl_id.replace("PMID_", "") if is_standard else "",
            "is_standard": is_standard,
            "summary": doc_summary,
            "children_found": 0,
            "full_text_file": csl_metadata.get("full_text_file")
        })
        
        # Chỉ lưu vào cơ sở dữ liệu trích dẫn nếu đây là tài liệu chuẩn có metadata
        if is_standard:
            item_doi = doc_doi.lower().strip() if doc_doi else ""
            item_pmid = doc_pmid.strip() if doc_pmid else ""
            is_dup = csl_id in existing_ids or (
                item_doi and item_doi in existing_dois
            ) or (
                item_pmid and item_pmid in existing_pmids
            )
            if not is_dup:
                new_csl_items.append(csl_metadata)
                existing_db.append(csl_metadata)
                existing_ids.add(csl_id)
                if item_doi:
                    existing_dois.add(item_doi)
                if item_pmid:
                    existing_pmids.add(item_pmid)

        # 3. Dò quét tài liệu con (References of References)
        all_dois = DOI_REGEX.findall(text)
        all_pmids = PMID_REGEX.findall(text)
        
        # Bỏ đi chính nó nếu trùng
        found_dois = {d.lower().strip() for d in all_dois if d.lower().strip() != (doc_doi or "").lower().strip()}
        found_pmids = {p.strip() for p in all_pmids if p.strip() != (doc_pmid or "")}
        
        # Lọc những bài đã có trong db
        new_dois = found_dois - existing_dois
        new_pmids = found_pmids - existing_pmids
        
        reports[-1]["children_found"] = len(found_dois) + len(found_pmids)
        logger.info(f"  → Tìm thấy {reports[-1]['children_found']} tài liệu con trong bài (Độc nhất: {len(new_dois)} DOIs, {len(new_pmids)} PMIDs)")
        
        child_dois.update(new_dois)
        child_pmids.update(new_pmids)

    # 4. Bắt đầu cào tài liệu con và gộp
    logger.info(f"=== BẮT ĐẦU CÀO TÀI LIỆU CON: {len(child_pmids)} PMIDs, {len(child_dois)} DOIs ===")
    
    # Tải PMID con bằng cách gộp batch
    child_pmids_list = list(child_pmids)
    crawled_count = 0
    if child_pmids_list:
        logger.info(f"Đang cào {len(child_pmids_list)} PMID bài báo con từ PubMed...")
        # Chia batch 10 PMID mỗi lượt để tránh URL quá dài
        for i in range(0, len(child_pmids_list), 10):
            batch = child_pmids_list[i:i+10]
            fetched = fetch_pubmed_by_ids(batch)
            for item in fetched:
                item_id = item["id"]
                item_doi = item.get("DOI", "").lower().strip()
                item_pmid = item.get("PMID", "").strip()
                
                is_dup = item_id in existing_ids or (
                    item_doi and item_doi in existing_dois
                ) or (
                    item_pmid and item_pmid in existing_pmids
                )
                
                if not is_dup:
                    citation_key = item_id.replace("PMID_", "")
                    # Tải toàn văn PMC
                    full_text_path = fetch_full_text_from_pmc(citation_key, citation_key, folder_path)
                    if not full_text_path and item.get("DOI"):
                        full_text_path = fetch_full_text_from_unpaywall(item["DOI"], citation_key, folder_path)
                    if full_text_path:
                        item["full_text_file"] = full_text_path
                    new_csl_items.append(item)
                    existing_db.append(item)
                    existing_ids.add(item_id)
                    if item_doi:
                        existing_dois.add(item_doi)
                    if item_pmid:
                        existing_pmids.add(item_pmid)
                    crawled_count += 1
                    logger.info(f"  ✓ Đã cào thành công bài báo con PMID {citation_key}")
                
    # Tải DOI con
    for doi in child_dois:
        # Giới hạn cào tối đa tránh làm chậm
        if crawled_count >= 30:
            logger.info("Đạt giới hạn cào tối đa 30 tài liệu con trong một phiên làm việc.")
            break
        doi_clean = doi.lower().strip()
        if doi_clean in existing_dois:
            continue
        csl_item = fetch_metadata_by_doi(doi)
        if csl_item:
            item_id = csl_item["id"]
            item_doi = csl_item.get("DOI", "").lower().strip()
            item_pmid = csl_item.get("PMID", "").strip()
            
            is_dup = item_id in existing_ids or (
                item_doi and item_doi in existing_dois
            ) or (
                item_pmid and item_pmid in existing_pmids
            )
            
            if not is_dup:
                citation_key = item_id.replace("PMID_", "")
                # Thử PMC trước nếu có PMID trong metadata
                full_text_path = None
                pmid_in_meta = csl_item.get("PMID")
                if pmid_in_meta:
                    full_text_path = fetch_full_text_from_pmc(pmid_in_meta, citation_key, folder_path)
                if not full_text_path:
                    full_text_path = fetch_full_text_from_unpaywall(doi, citation_key, folder_path)
                if full_text_path:
                    csl_item["full_text_file"] = full_text_path
                new_csl_items.append(csl_item)
                existing_db.append(csl_item)
                existing_ids.add(item_id)
                if item_doi:
                    existing_dois.add(item_doi)
                if item_pmid:
                    existing_pmids.add(item_pmid)
                crawled_count += 1
                logger.info(f"  ✓ Đã cào thành công bài báo con DOI {doi}")

    with open(db_path, "w", encoding="utf-8") as f:
        json.dump(existing_db, f, ensure_ascii=False, indent=4)
        
    logger.info(f"Đã cập nhật database {db_path} (Tổng số tài liệu khả dụng: {len(existing_db)}).")

    # 6. Tạo tóm tắt ngữ cảnh nghiên cứu custom_references_log.md
    log_path = os.path.join(os.getcwd(), "custom_references_log.md")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("# NHẬT KÝ VÀ NGỮ CẢNH TÀI LIỆU THAM KHẢO TỰ CHỌN\n\n")
        f.write(f"Thời gian quét: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Số lượng file tự chọn được quét thành công: {len(reports)}\n")
        f.write(f"Số lượng tài liệu con được cào thêm thành công: {crawled_count}\n\n")
        
        f.write("## 1. Danh sách Tài liệu Tham khảo Tự chọn & Khóa trích dẫn\n\n")
        f.write("| Tên file | Tiêu đề tài liệu | Khóa trích dẫn (Citation Key) | Loại | Tài liệu con phát hiện |\n")
        f.write("| --- | --- | --- | --- | --- |\n")
        for r in reports:
            if r["is_standard"]:
                citation_str = f"`[@PMID:{r['key']}]`"
                type_str = "Chuẩn hóa (Có trích dẫn)"
            else:
                citation_str = "`Không trích dẫn (Chỉ tham khảo)`"
                type_str = "Nội bộ (Không trích dẫn)"
            f.write(f"| `{r['filename']}` | {r['title']} | {citation_str} | {type_str} | {r['children_found']} |\n")
        
        f.write("\n## 2. Chi tiết Nội dung & Tóm tắt các Tài liệu Tự chọn\n\n")
        for r in reports:
            f.write(f"### 📄 Tài liệu: {r['title']}\n")
            f.write(f"- **Tên tệp tin**: `{r['filename']}`\n")
            if r["is_standard"]:
                f.write(f"- **Cách trích dẫn trong văn bản**: sử dụng mỏ neo `[@PMID:{r['key']}]`\n")
            else:
                f.write(f"- **Cách tham chiếu trong văn bản**: Tài liệu nội bộ này KHÔNG có khóa trích dẫn trực tiếp. Hãy tham khảo và tích hợp sâu nội dung thực tế của nó vào phần mô tả tương ứng trong bản thảo mà không chèn mỏ neo trích dẫn nào.\n")
            if r.get("full_text_file"):
                f.write(f"- **Tập tin toàn văn (Full text)**: [{r['full_text_file']}](file:///{os.path.abspath(r['full_text_file']).replace('\\', '/')}) (Đã tải về và chuyển đổi sang text)\n")
            f.write(f"- **Tóm tắt nội dung thô / Đoạn mở đầu**:\n  > {r['summary']}\n\n")
            f.write("---\n\n")
            
        f.write("## 3. Danh sách Tài liệu con đã cào (Scraped Citations)\n\n")
        f.write("Các tài liệu dưới đây đã được cào siêu dữ liệu và đưa trực tiếp vào `references_db.json`. Bạn có thể sử dụng các khóa trích dẫn này trong bài viết:\n\n")
        # Thu thập tất cả key chuẩn của các tài liệu tự chọn
        standard_keys = {r["key"] for r in reports if r["is_standard"] and r["key"]}
        for item in new_csl_items:
            # Chỉ hiển thị các bài cào thêm
            item_id = item["id"]
            citation_key = item_id.replace("PMID_", "")
            if citation_key not in standard_keys:
                author_str = ", ".join([a.get("family", "") for a in item.get("author", [])[:3]])
                if len(item.get("author", [])) > 3:
                    author_str += " et al."
                year_parts = item.get("issued", {}).get("date-parts", [["2026"]])
                year = year_parts[0][0] if year_parts else "2026"
                full_text_line = ""
                if item.get("full_text_file"):
                    full_text_line = f" | [Toàn văn (Full text)](file:///{os.path.abspath(item['full_text_file']).replace('\\', '/')})"
                f.write(f"- **`[@PMID:{citation_key}]`**: {item.get('title')} ({author_str}, {year}){full_text_line}\n")

    logger.info(f"Đã xuất nhật ký ngữ cảnh nghiên cứu thành công tại {log_path}")
    print(json.dumps({
        "status": "success",
        "scanned_files": len(reports),
        "crawled_references": crawled_count,
        "total_references": len(existing_db),
        "message": "Phân tích tài liệu và cào tham khảo con thành công."
    }))


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8") if hasattr(sys, 'stdout') and hasattr(sys.stdout, 'buffer') else sys.stdout
    import io
    
    parser = argparse.ArgumentParser(description="Parse custom references.")
    parser.add_argument("--folder", type=str, default="custom_references", help="Path to custom references folder")
    parser.add_argument("--output-db", type=str, default="references_db.json", help="Path to references_db.json")
    args = parser.parse_args()
    
    process_custom_folder(args.folder, args.output_db)
