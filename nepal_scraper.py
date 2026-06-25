"""
nepal_scraper.py
================
Scraper for https://www.jobsnepal.com/jobs

Architecture mirrors the established five-file pattern but is self-contained
in one file for ease of deployment. Matches the Malawi scraper structure:
  - Paginated listing crawl (no sitemap — jobsnepal.com uses ?page=N)
  - Detail page parse with verified selectors from live page inspection
  - Deduplication via processed_jobs.csv tracker
  - Public-apply-only rule: no email/URL → flagged_no_apply.csv, never posted
  - Mistral paraphrase (title + description + company details)
  - WordPress posting via WP Job Manager REST API (/wp-json/wp/v2/job-listings)
  - Excel export to jobs_output.xlsx
  - Early-stop pagination: stops when all jobs on a page are already seen

Verified selectors (from live page fetch 2026-06-25):
  Listing:  h2 > a           → job title + relative URL
            ul > li           → company name (1st), location (2nd), category (3rd+)
            img[src]          → placeholder on listing, skip — use detail page
  Detail:   h1                → job title
            h2 > a[/employer] → company name + employer page URL
            meta[og:image]    → company logo (img.jobsnepal.com/big/HASH.ext)
            p (above Details) → company short blurb
            table (Overview)  → Apply Before, Position Type, City, Posted Date,
                                 Experience, Education, Openings, Offered Salary
            div.details-requirements or div after h2:contains("Details")
                              → job description body
            a[href^=mailto:]  → apply email (inside description)
            a[href^=http] in description → external apply URL
"""

import os
import re
import csv
import sys
import time
import json
import base64
import hashlib
import logging
from datetime import datetime, timedelta
from urllib.parse import urljoin, urlparse, urlsplit, urlunsplit, parse_qsl, urlencode

import requests
import urllib3
from bs4 import BeautifulSoup

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import openpyxl
    _XLSX_AVAILABLE = True
except ImportError:
    _XLSX_AVAILABLE = False

try:
    from sentence_transformers import SentenceTransformer, util as st_util
    _NLP_AVAILABLE = True
except ImportError:
    _NLP_AVAILABLE = False

# Suppress InsecureRequestWarning for internal WP API calls
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =============================================================================
#  CONFIG
# =============================================================================

BASE_URL        = "https://www.jobsnepal.com"
JOBS_LIST_URL   = "https://www.jobsnepal.com/jobs"   # ?page=N for pagination

MAX_JOBS        = int(os.environ.get("MAX_JOBS", "0"))
MAX_PAGES       = int(os.environ.get("MAX_PAGES", "0"))   # 0 = all pages
REQUEST_DELAY   = float(os.environ.get("REQUEST_DELAY", "2.0"))
REQUEST_TIMEOUT = int(os.environ.get("REQUEST_TIMEOUT", "25"))

RESOLVE_APPLY_URLS = os.environ.get("RESOLVE_APPLY_URLS", "1") != "0"
RESOLVE_DELAY      = float(os.environ.get("RESOLVE_DELAY", "0.5"))

OUTPUT_FILE        = "jobs_output.xlsx"
PROCESSED_IDS_FILE = "processed_jobs.csv"
FLAGGED_FILE       = "flagged_no_apply.csv"

_TRACKER_FIELDS = ["Job ID", "Job URL", "Job Title", "Company Name",
                   "Status", "Timestamp", "WP ID"]
_FLAGGED_FIELDS = ["Job ID", "Job URL", "Job Title", "Company Name",
                   "Reason", "Timestamp"]

# ── WordPress ─────────────────────────────────────────────────────────────────
WP_URL      = os.environ.get("WP_BASE_URL", "").rstrip("/")
WP_USER     = os.environ.get("WP_USERNAME", "")
WP_PASSWORD = os.environ.get("WP_APP_PASSWORD", "")
WP_BASE      = WP_URL
WP_API_BASE      = f"{WP_BASE}/wp-json/wp/v2"
WP_JOBS_URL      = f"{WP_API_BASE}/job-listings"   # WP Job Manager custom post type
WP_MEDIA_URL     = f"{WP_API_BASE}/media"
WP_TAX_TYPE_URL  = f"{WP_API_BASE}/job_listing_type"    # employment type taxonomy
WP_TAX_TAG_URL   = f"{WP_API_BASE}/job_listing_tag"     # tag taxonomy (if registered)
WP_TAX_CAT_URL   = f"{WP_API_BASE}/job_listing_category" # category taxonomy

# ── Mistral ───────────────────────────────────────────────────────────────────
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY", "")
MISTRAL_MODEL   = "mistral-small-latest"
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"
ENABLE_PARAPHRASE = True

JOB_TYPE_MAPPING = {
    "full time": "full-time", "full-time": "full-time",
    "part time": "part-time", "part-time": "part-time",
    "contract":  "contract",  "contractor": "contract",
    "temporary": "temporary", "internship": "internship",
    "freelance": "freelance", "volunteer":  "volunteer",
    "other":     "other",
}

# Hosts where href="https://gmail.com" etc. but the real email is the anchor text
_GENERIC_MAIL_HOSTS = {
    "gmail.com", "mail.google.com", "yahoo.com", "yahoo.co.uk", "yahoo.com.np",
    "hotmail.com", "outlook.com", "live.com", "protonmail.com",
    "mail.com", "ymail.com", "aol.com",
}

def _is_generic_mail_href(href):
    """Return True when href points to a webmail homepage, not an apply portal."""
    try:
        host = urlparse(href).netloc.lower().lstrip("www.")
        return host in _GENERIC_MAIL_HOSTS
    except Exception:
        return False

# =============================================================================
#  QUALIFICATION / EXPERIENCE / JOB-FIELD NORMALISATION
# =============================================================================

QUALIFICATION_TIERS = [
    ("PhD / Doctorate",
     ["phd","ph.d","doctorate","doctoral","doctor of philosophy"]),
    ("Master's Degree",
     ["master","msc","m.sc","ma ","m.a ","mba","m.b.a","meng","m.eng","mphil",
      "m.p.h","mph","postgraduate","post-graduate","post graduate"]),
    ("Bachelor's Degree",
     ["bachelor","bsc","b.sc","ba ","b.a ","beng","b.eng","bcom","b.com","bba",
      "llb","degree in","undergraduate degree","honours degree","hons"]),
    ("Higher National Diploma",
     ["hnd","hnc","higher national diploma","higher national certificate",
      "higher diploma","advanced diploma"]),
    ("Diploma",
     ["diploma","dip ","dip.","associate degree","foundation degree",
      "proficiency certificate"]),
    ("Professional Certification",
     ["acca","cpa","cfa","cima","pmp","prince2","cissp","aws certified",
      "comptia","cisco","ccna","ccnp","shrm","cipd","chartered",
      "certified public","certified financial","certified project",
      "professional certification","professional certificate"]),
    ("A-Levels / HSC",
     ["a-level","a level","hsc","higher school certificate","ib diploma",
      "international baccalaureate","gce advanced","10+2","class 12","slc passed"]),
    ("O-Levels / School Certificate",
     ["o-level","o level","igcse","gcse","school certificate",
      "sc ","slc","see","secondary education examination"]),
    ("No Formal Qualification Required",
     ["no qualification","no degree","no formal","school leaver",
      "entry level","no experience required","training provided","will train"]),
]

_NO_EXP_KW  = ["no experience","no prior experience","fresh graduate","freshers",
                "entry level","entry-level","training provided","will train",
                "no experience required"]
_LESS1_KW   = ["less than 1 year","under 1 year","6 months","less than a year",
                "some experience","minimal experience"]

# Compiled patterns for safe word-boundary matching
_NO_EXP_RE  = re.compile(
    r"\bno\s+(?:prior\s+)?experience\b|"
    r"\bfresh\s+graduate[s]?\b|\bfreshers?\b|"
    r"\bentry[\s-]level\b|"
    r"\btraining\s+provided\b|\bwill\s+train\b|"
    r"\b0\s+years?\b|"
    r"\bno\s+experience\s+required\b",
    re.I,
)
_LESS1_RE   = re.compile(
    r"\bless\s+than\s+(?:a\s+)?1\s+year\b|\bunder\s+1\s+year\b|"
    r"\b6\s+months?\b|\bsix\s+months?\b|"
    r"\bminimal\s+experience\b|\bsome\s+experience\b",
    re.I,
)

def normalise_qualification(raw: str) -> str:
    """Map raw education text to a standardised tier label."""
    if not raw:
        return ""
    low = raw.lower()
    for label, keywords in QUALIFICATION_TIERS:
        if any(kw in low for kw in keywords):
            return label
    # If it just says "please check details/vacancy" return empty
    if re.search(r"please\s+check", low):
        return ""
    return raw.strip()

def normalise_experience(raw: str) -> str:
    """
    Map raw experience text to a standard band:
      No Experience Required | Less than 1 Year | 1 - 2 Years |
      3 - 5 Years | 6 - 10 Years | 10+ Years
    """
    if not raw:
        return ""
    low = raw.lower()
    if re.search(r"please\s+check", low):
        return ""
    if _NO_EXP_RE.search(low):
        return "No Experience Required"
    if _LESS1_RE.search(low):
        return "Less than 1 Year"
    # Extract the first integer — handles "2+ years", "2–3 years", "at least 3 years"
    nums = re.findall(r"\d+", low)
    if nums:
        n        = int(nums[0])
        has_plus = bool(re.search(r"\d+\s*\+|\bmore\s+than\b|\babove\b|\bover\b", low))
        if n == 0:                    return "No Experience Required"
        if n <= 2:                    return "1 - 2 Years"
        if n <= 5:                    return "3 - 5 Years"
        if n < 10:                    return "6 - 10 Years"
        if n == 10 and not has_plus:  return "6 - 10 Years"
        return "10+ Years"
    return raw.strip()

# Job field mapping — maps category/keyword to a canonical field label
JOB_FIELD_MAP = [
    ("Information Technology",      ["information technology","it jobs","software","developer","programmer",
                                     "data analyst","cyber","network","database","devops","computer"]),
    ("Accounting & Finance",        ["accounting","finance","financial","audit","tax","bookkeeping",
                                     "treasurer","fiscal","accounts"]),
    ("Administration & Management", ["administration","management","admin","office manager",
                                     "executive assistant","operations","coordinator"]),
    ("Healthcare & Public Health",  ["health","medical","nurse","doctor","clinical","pharmacy",
                                     "public health","epidemiology","physician","midwife"]),
    ("Engineering",                 ["engineer","engineering","civil","mechanical","electrical",
                                     "structural","architecture","construction","surveyor"]),
    ("Education & Training",        ["teacher","education","training","lecturer","tutor","instructor",
                                     "academic","school","university","curriculum"]),
    ("NGO / Development",           ["ngo","ingo","development project","humanitarian","social work",
                                     "community development","advocacy","livelihood","gender"]),
    ("Research & Consultancy",      ["research","survey","evaluation","consultant","consultancy",
                                     "analyst","assessment","study","investigation"]),
    ("Sales & Marketing",           ["sales","marketing","business development","brand","digital marketing",
                                     "advertising","promotions","crm"]),
    ("Logistics & Supply Chain",    ["logistics","supply chain","procurement","warehouse","driver",
                                     "transport","fleet","delivery","customs"]),
    ("Legal",                       ["legal","lawyer","attorney","law","compliance","paralegal","judicial"]),
    ("Media & Communications",      ["media","journalism","communication","pr ","public relations",
                                     "photography","videography","content","reporter","editor"]),
    ("Agriculture & Environment",   ["agriculture","agronomy","livestock","veterinary","forestry",
                                     "environment","climate","irrigation","fisheries"]),
    ("Hospitality & Tourism",       ["hospitality","hotel","tourism","catering","restaurant",
                                     "travel","housekeeping","chef","cook"]),
    ("Tender / Expression of Interest", ["tender","expression of interest","eoi","bid","itb","rfp",
                                          "request for proposal","invitation to bid","reoi","tor "]),
]

def normalise_job_field(category: str, title: str = "", description: str = "") -> str:
    """Map category / title / description text to a canonical job field."""
    combined = f"{category} {title}".lower()
    for field_label, keywords in JOB_FIELD_MAP:
        if any(kw in combined for kw in keywords):
            return field_label
    # Try description as last resort
    if description:
        desc_low = description[:500].lower()
        for field_label, keywords in JOB_FIELD_MAP:
            if any(kw in desc_low for kw in keywords):
                return field_label
    return "Other"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept":         "text/html,application/xhtml+xml",
    "Accept-Language": "en-US,en;q=0.9",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)

_apply_url_cache: dict = {}

# jobsnepal.com placeholder image — never use this as a logo
JOBSNEPAL_PLACEHOLDER = "gray-back.jpg"

# Domains that are document/file hosts, NOT application portals.
# A link to drive.google.com is the vacancy PDF, not an apply button.
# workspace.google.com is Gmail's homepage — a resolved mailto: redirect.
# These are never valid apply_url values.
_DOCUMENT_HOST_RE = re.compile(
    r"^https?://(?:www\.)?"
    r"(?:drive\.google\.com"
    r"|docs\.google\.com"
    r"|workspace\.google\.com"
    r"|mail\.google\.com"
    r"|dropbox\.com"
    r"|onedrive\.live\.com"
    r"|1drv\.ms"
    r"|sharepoint\.com"
    r"|box\.com"
    r")/",
    re.I,
)

def _is_document_host(url):
    """Return True if url points to a file/document host rather than an apply portal."""
    return bool(_DOCUMENT_HOST_RE.match(url or ""))

# =============================================================================
#  LOGGING / COLOUR
# =============================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log_ = logging.getLogger(__name__)

_USE_COLOUR = sys.stdout.isatty()

def _c(code, text):
    return f"\033[{code}m{text}\033[0m" if _USE_COLOUR else text

C_HEADER  = lambda t: _c("1;36",  t)
C_LABEL   = lambda t: _c("1;33",  t)
C_VALUE   = lambda t: _c("97",    t)
C_DIM     = lambda t: _c("2",     t)
C_GREEN   = lambda t: _c("1;32",  t)
C_RED     = lambda t: _c("1;31",  t)
C_BLUE    = lambda t: _c("1;34",  t)
C_DIVIDER = lambda: _c("2", "─" * 80)

def log(msg):
    print(msg, flush=True)

TRACKING_PARAM_PREFIXES = ("utm_",)
TRACKING_PARAM_EXACT = {
    "fbclid", "gclid", "msclkid", "mc_cid", "mc_eid", "ref", "referrer",
}

EMAIL_RE        = re.compile(r"[A-Za-z0-9.+_-]+@[A-Za-z0-9-]+\.[A-Za-z0-9.-]+")
META_REFRESH_RE = re.compile(
    r'<meta[^>]+http-equiv=["\']refresh["\'][^>]*content=["\'][^"\'>]*url=([^"\'>]+)', re.I)
JS_REDIRECT_RE  = re.compile(
    r'(?:window\.)?location(?:\.href)?\s*(?:=\s*|\.replace\(\s*)["\']([^"\']+)["\']', re.I)

BOILERPLATE_PATTERNS = [
    re.compile(r"save to favou?rite.*", re.I | re.S),
    re.compile(r"refer to a friend.*", re.I | re.S),
    re.compile(r"similar jobs.*",       re.I | re.S),
    re.compile(r"jobs by categor.*",    re.I | re.S),
    re.compile(r"share\s*×.*",          re.I | re.S),
]

_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]

# Deadline extraction — jobsnepal.com uses "Apply Before: DD Mon, YYYY" in the
# Overview table, which is the most reliable source. Also scan description body
# for "Application Deadline:", "Deadline:", etc. as fallback.
_DEADLINE_LABEL_RE = re.compile(
    r'(?:the\s+)?'
    r'(?:apply\s+(?:before|by)'
    r'|application\s+deadline'
    r'|closing\s+date(?:\s+for\s+(?:receipt|receiving)\s+(?:of\s+)?applications)?'
    r'|closure\s+date'
    r'|deadline(?:\s+for\s+applications)?'
    r')'
    r'\s*(?:is\s+(?:(?:on\s+or\s+before|on)\s+)?|:\s*(?:(?:on\s+or\s+before|on)\s+)?|[:\-–]\s*)',
    re.I,
)
_DEADLINE_NOISE_RE = re.compile(
    r'^(?:the\s+)?(?:apply\s+(?:before|by)|application\s+deadline|closing\s+date|deadline)'
    r'(?:\s+for\s+(?:receipt|receiving)\s+(?:of\s+)?applications)?'
    r'\s*(?:is\s+(?:(?:on\s+or\s+before|on)\s+)?|:\s*)',
    re.I,
)
_DEADLINE_DATE_RE = re.compile(
    r'^((?:[A-Za-z]+,?\s+)?'
    r'(?:\d{1,2}(?:st|nd|rd|th)?\s+\w+,?\s+\d{4}'
    r'|\d{1,2}\s+\w+,?\s+\d{4}'
    r'|\d{1,2}/\w+/\d{4}'
    r'|\d{1,2}/\d{1,2}/\d{4}'
    r'|\d{4}-\d{2}-\d{2}'
    r'|\w+\s+\d{1,2},?\s+\d{4}'
    r')(?:[, ]+(?:at\s+)?[\d:]+\s*(?:am|pm|CAT|EAT|GMT|UTC|NPT)?)?)',
    re.I,
)

# =============================================================================
#  TEXT / URL HELPERS
# =============================================================================

def _fix_mojibake(text):
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)

def sanitize_text(text, is_url=False):
    if not isinstance(text, str):
        text = str(text) if (text is not None and str(text) not in ("nan","None","NaN")) else ""
    text = text.strip()
    if text in ("nan","None","NaN","","N/A","n/a","NA","na"):
        return ""
    text = _fix_mojibake(text)
    if is_url:
        return re.sub(r"[ \t\r\n\f\v]+", " ", text).strip()
    text = re.sub(r"#+\s*", "", text)
    text = re.sub(r"\*\*", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def clean_text(el):
    if el is None:
        return ""
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)).strip()

def clean_description(text):
    if not text:
        return text
    for pat in BOILERPLATE_PATTERNS:
        text = pat.sub("", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def strip_tracking_params(url):
    if not url:
        return url
    parts = urlsplit(url)
    if not parts.query:
        return url
    kept = [(k, v) for k, v in parse_qsl(parts.query, keep_blank_values=True)
            if not k.lower().startswith(TRACKING_PARAM_PREFIXES)
            and k.lower() not in TRACKING_PARAM_EXACT]
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(kept), parts.fragment))

def _same_site(url):
    try:
        return urlparse(url).netloc == urlparse(BASE_URL).netloc
    except Exception:
        return False

def extract_email(text):
    if not text:
        return ""
    m = EMAIL_RE.search(text)
    return m.group(0) if m else ""

def absolute_url(href):
    if not href:
        return ""
    return href if href.startswith("http") else urljoin(BASE_URL, href)

def extract_deadline_from_text(text):
    """Scan body text for a deadline sentence and return the date portion only."""
    if not text:
        return ""
    m = _DEADLINE_LABEL_RE.search(text)
    if not m:
        return ""
    remainder = text[m.end():].strip()
    remainder = _DEADLINE_NOISE_RE.sub("", remainder).strip()
    dm = _DEADLINE_DATE_RE.match(remainder)
    if dm:
        return dm.group(1).strip().rstrip(",")
    return re.split(r"[.\n]", remainder)[0].strip()[:80]

def resolve_relative_date(text):
    """Convert '3 days ago', '1 week ago' → YYYY-MM-DD. Returns text unchanged if not matched."""
    if not text:
        return text
    m = re.match(r"(\d+)\s+(day|week|month)s?\s+ago", text.strip(), re.I)
    if not m:
        return text
    n, unit = int(m.group(1)), m.group(2).lower()
    delta = {"day": timedelta(days=n), "week": timedelta(weeks=n),
             "month": timedelta(days=n * 30)}[unit]
    return (datetime.now() - delta).strftime("%Y-%m-%d")

# =============================================================================
#  HTTP HELPERS
# =============================================================================

def get_soup(url, retries=2):
    for attempt in range(retries + 1):
        try:
            resp = SESSION.get(url, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            resp.encoding = resp.encoding or "utf-8"
            return BeautifulSoup(resp.text, "lxml")
        except requests.RequestException as e:
            if attempt < retries:
                time.sleep(2 ** attempt)
                continue
            raise

def _find_html_redirect(html, current_url):
    m = META_REFRESH_RE.search(html) or JS_REDIRECT_RE.search(html)
    if not m:
        return ""
    return urljoin(current_url, m.group(1).strip().strip("'\""))

def resolve_apply_url(raw):
    if not raw:
        return ""
    if raw in _apply_url_cache:
        return _apply_url_cache[raw]
    resolved = ""
    try:
        resp = SESSION.get(raw, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        final = resp.url
        if final and not _same_site(final):
            resolved = final
        else:
            html_target = _find_html_redirect(resp.text, final)
            resolved = html_target if html_target else (final if final != raw else "")
    except requests.RequestException as e:
        log(f"    WARNING: could not resolve apply URL {raw}: {e}")
    resolved = strip_tracking_params(resolved)
    _apply_url_cache[raw] = resolved
    if RESOLVE_DELAY:
        time.sleep(RESOLVE_DELAY)
    return resolved

def resolve_application_contact(raw_apply_url, description, raw_anchor_text=""):
    """
    Return dict with apply_url and/or apply_email.

    Priority order:
      1. Email from raw_apply_url (when it's already a bare email address)
      2. Email from raw_anchor_text (visible text of the apply link)
      3. Email scanned from description body text
      4. External URL (only when no email found anywhere)

    Document/file hosts (Google Drive, Dropbox, etc.) are always rejected as URLs.
    """
    result = {"apply_url": "", "apply_email": "", "apply_raw": raw_apply_url}

    # Step 1: raw value is already a bare email (mailto: stripped by caller)
    if raw_apply_url and re.match(r"^[A-Za-z0-9.+_-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$",
                                   raw_apply_url):
        result["apply_email"] = raw_apply_url
        return result

    # Step 2 & 3: look for email in anchor text then description — email wins over URL
    email = extract_email(raw_anchor_text) or extract_email(description)
    if email:
        result["apply_email"] = email
        return result

    # Step 4: no email found — try resolving the URL
    if raw_apply_url and RESOLVE_APPLY_URLS:
        resolved = resolve_apply_url(raw_apply_url)
        if resolved and not _is_document_host(resolved):
            result["apply_url"] = resolved

    return result

# =============================================================================
#  MISTRAL PARAPHRASE
# =============================================================================

_sim_model = None

def _get_sim_model():
    global _sim_model
    if _sim_model is None and _NLP_AVAILABLE:
        try:
            _sim_model = SentenceTransformer("all-MiniLM-L6-v2", device="cpu")
        except Exception as e:
            log_.warning(f"SentenceTransformer init failed: {e}")
    return _sim_model

def similarity_score(a, b):
    model = _get_sim_model()
    if model:
        try:
            emb = model.encode([a, b], convert_to_tensor=True)
            return float(st_util.pytorch_cos_sim(emb[0], emb[1]))
        except Exception:
            pass
    def tokens(s):
        return set(re.sub(r"[^a-z0-9 ]", " ", s.lower()).split())
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / max(len(ta), len(tb))

def clean_output(text):
    text = _fix_mojibake(text)
    for pat in [r"\[/?INST\]", r"</?s>",
                r"(?i)(rewritten?|rephrased?|output|paraphrase[d]?)[:\s]+",
                r"\*\*", r"###", r"---"]:
        text = re.sub(pat, "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def mistral_generate(prompt, max_tokens=400, temperature=0.7):
    if not MISTRAL_API_KEY:
        return ""
    try:
        r = requests.post(
            MISTRAL_URL,
            headers={"Authorization": f"Bearer {MISTRAL_API_KEY}",
                     "Content-Type": "application/json"},
            json={"model": MISTRAL_MODEL,
                  "messages": [{"role": "user", "content": prompt}],
                  "max_tokens": max_tokens, "temperature": temperature},
            timeout=30,
        )
        r.raise_for_status()
        return r.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log_.error(f"Mistral API error: {e}")
        return ""

def paraphrase_title(title):
    if not ENABLE_PARAPHRASE or not MISTRAL_API_KEY:
        return title
    clean = sanitize_text(title)
    if not clean:
        return title
    # Skip Nepali script titles — Mistral may mangle them
    if re.search(r"[\u0900-\u097F]", clean):
        return clean
    best_result, best_sim = None, 0.0
    for attempt in range(4):
        temp = round(0.68 + attempt * 0.06, 2)
        prompt = (f"Rewrite this job title professionally using different words. "
                  f"Output ONLY the rewritten title, nothing else. "
                  f"Keep it between 4 and 12 words.\n\nJob title: {clean}")
        raw    = mistral_generate(prompt, max_tokens=50, temperature=temp)
        result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")
        wc     = len(result.split()) if result else 0
        sim    = similarity_score(clean, result) if result else 0.0
        is_dup = result.lower().strip() == clean.lower().strip()
        if bool(result) and 4 <= wc <= 14 and sim >= 0.55 and not is_dup and sim > best_sim:
            best_sim, best_result = sim, result
        time.sleep(1)
    return best_result if best_result else clean

def paraphrase_description(text):
    if not ENABLE_PARAPHRASE or not MISTRAL_API_KEY:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text
    paragraphs = [p.strip() for p in re.split(r"\n+", clean) if p.strip()] or [clean]
    rewritten = []
    for para in paragraphs:
        # Skip Nepali script paragraphs unchanged
        if re.search(r"[\u0900-\u097F]", para):
            rewritten.append(para)
            continue
        prompt = (f"Rewrite this job description paragraph professionally. "
                  f"Keep ALL facts, requirements, and responsibilities. "
                  f"Use different sentence structure and vocabulary. "
                  f"Output ONLY the rewritten paragraph — no labels, no explanation.\n\n"
                  f"Original:\n{para}")
        accepted, best_result, best_sim = None, None, 0.0
        for attempt in range(3):
            raw    = mistral_generate(prompt, max_tokens=500,
                                      temperature=round(0.65 + attempt * 0.08, 2))
            result = clean_output(raw).strip()
            rw     = len(result.split()) if result else 0
            sim    = similarity_score(para, result) if result and rw >= 5 else 0.0
            if bool(result) and rw >= 8 and sim >= 0.48:
                accepted = result
                break
            if result and sim > best_sim:
                best_sim, best_result = sim, result
            time.sleep(1)
        rewritten.append(accepted or (best_result if best_result and best_sim >= 0.40 else para))
    return "\n\n".join(rewritten)

def paraphrase_company(text):
    if not ENABLE_PARAPHRASE or not MISTRAL_API_KEY:
        return text
    clean = sanitize_text(text)
    if not clean or re.search(r"[\u0900-\u097F]", clean):
        return text
    prompt = (f"Rewrite this company description professionally. "
              f"Preserve all facts. Use different wording. "
              f"Output ONLY the rewritten description.\n\nOriginal:\n{clean}")
    raw    = mistral_generate(prompt, max_tokens=600, temperature=0.68)
    result = clean_output(raw)
    time.sleep(1)
    return result if result and len(result.split()) >= 10 else clean

# =============================================================================
#  TRACKER
# =============================================================================

def _init_tracker():
    for fpath, fields in [
        (PROCESSED_IDS_FILE, _TRACKER_FIELDS),
        (FLAGGED_FILE,        _FLAGGED_FIELDS),
    ]:
        if not os.path.exists(fpath):
            try:
                with open(fpath, "w", newline="", encoding="utf-8") as f:
                    csv.writer(f).writerow(fields)
            except Exception as e:
                log_.error(f"Could not create {fpath}: {e}")

def load_processed_ids():
    _init_tracker()
    ids, urls = set(), set()
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("Job ID"):  ids.add(row["Job ID"].strip())
                if row.get("Job URL"): urls.add(row["Job URL"].strip())
    except Exception as e:
        log_.error(f"Could not read tracker: {e}")
    return ids, urls

def _upsert_row(job_id, updates):
    _init_tracker()
    rows = []
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
    except Exception:
        pass
    found = False
    for row in rows:
        if row.get("Job ID", "").strip() == str(job_id):
            row.update(updates)
            row["Timestamp"] = datetime.now().isoformat()
            found = True
            break
    if not found:
        new_row = {k: "" for k in _TRACKER_FIELDS}
        new_row["Job ID"]    = str(job_id)
        new_row["Timestamp"] = datetime.now().isoformat()
        new_row.update(updates)
        rows.append(new_row)
    try:
        with open(PROCESSED_IDS_FILE, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=_TRACKER_FIELDS, extrasaction="ignore")
            w.writeheader()
            w.writerows(rows)
    except Exception as e:
        log_.error(f"Tracker write error: {e}")

def flag_job(job_id, job_url, title, company, reason="no_apply_contact"):
    log_.warning(f"FLAGGED [{reason}]: {title} | {job_url}")
    rows = []
    try:
        with open(FLAGGED_FILE, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
    except Exception:
        pass
    if any(r.get("Job ID", "").strip() == str(job_id) for r in rows):
        return
    rows.append({"Job ID": job_id, "Job URL": job_url, "Job Title": title,
                 "Company Name": company, "Reason": reason,
                 "Timestamp": datetime.now().isoformat()})
    try:
        with open(FLAGGED_FILE, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=_FLAGGED_FIELDS, extrasaction="ignore")
            w.writeheader()
            w.writerows(rows)
    except Exception as e:
        log_.error(f"Flagged file write error: {e}")

def make_job_id(job_url, title="", company=""):
    seed = job_url or f"{title}{company}"
    return hashlib.md5(seed.encode()).hexdigest()[:16]

def mark_scraped(job_id, job_url, title, company):
    _upsert_row(job_id, {"Job URL": job_url, "Job Title": title,
                         "Company Name": company, "Status": "scraped", "WP ID": ""})

def mark_paraphrased(job_id):
    _upsert_row(job_id, {"Status": "paraphrased"})

def mark_posted(job_id, wp_id):
    _upsert_row(job_id, {"Status": "posted", "WP ID": str(wp_id)})

def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})

# =============================================================================
#  STEP 1 — COLLECT JOB URLS (paginated listing)
# =============================================================================

def collect_job_urls(processed_urls):
    """
    Paginate https://www.jobsnepal.com/jobs?page=N and collect all job URLs.
    Implements early-stop: if an entire page consists only of already-seen URLs,
    stops pagination (since listing is newest-first).
    Returns list of (job_url, company_name, location, category) tuples so we
    avoid re-fetching that data from the detail page when not needed.
    """
    log(f"\n{'='*80}\nCOLLECTING JOB URLS FROM: {JOBS_LIST_URL}\n{'='*80}")
    all_jobs   = []
    page       = 1
    seen_slugs = set()

    while True:
        url  = f"{JOBS_LIST_URL}?page={page}" if page > 1 else JOBS_LIST_URL
        log(f"  Page {page}: {url}")
        try:
            soup = get_soup(url)
        except Exception as e:
            log(f"  ERROR fetching page {page}: {e}")
            break

        # Each job is in a structure: img-link + h2>a + ul>li... + Apply Now + badge
        # The h2 elements with job titles all have an <a> linking to the job slug
        job_cards = soup.select("h2 a[href]")
        if not job_cards:
            log(f"  No job cards found on page {page} — stopping.")
            break

        new_on_page = 0
        for a in job_cards:
            href = a.get("href", "")
            if not href or "/employer/" in href or "/category/" in href:
                continue
            job_url = absolute_url(href)
            if job_url in seen_slugs:
                continue
            seen_slugs.add(job_url)

            # Walk up to the card container to grab company/location/category from li elements
            card = a.find_parent(["div", "article", "li", "section"])
            if not card:
                card = a.parent

            title   = clean_text(a)
            lis     = card.select("ul li") if card else []
            company  = clean_text(lis[0]) if len(lis) > 0 else ""
            location = clean_text(lis[1]) if len(lis) > 1 else ""
            # Categories are remaining li items
            category = ", ".join(clean_text(li) for li in lis[2:] if clean_text(li))

            if job_url not in processed_urls:
                new_on_page += 1
            all_jobs.append({
                "job_url":  job_url,
                "title":    title,
                "company":  company,
                "location": location,
                "category": category,
            })

        log(f"    Found {len(job_cards)} cards, {new_on_page} new")

        # Early stop: entire page was already seen
        if new_on_page == 0:
            log(f"  All jobs on page {page} already seen — stopping pagination.")
            break

        # Check for next page link
        next_link = soup.select_one("a[href*='?page=']")
        # Find the specific "Next" link
        next_page_link = None
        for a in soup.select("a"):
            txt = clean_text(a)
            if txt.lower() in ("next", "next »", "next»", "»"):
                next_page_link = a
                break
            if a.get("href", "").endswith(f"page={page + 1}"):
                next_page_link = a
                break

        if not next_page_link:
            log(f"  No next page link found after page {page} — done.")
            break

        if MAX_PAGES and page >= MAX_PAGES:
            log(f"  MAX_PAGES ({MAX_PAGES}) reached — stopping.")
            break

        page += 1
        time.sleep(REQUEST_DELAY)

    log(f"  Total URLs collected: {len(all_jobs)}")
    return all_jobs

# =============================================================================
#  STEP 2 — PARSE INDIVIDUAL JOB DETAIL PAGE
# =============================================================================

def _parse_overview_table(soup):
    """
    Parse the Overview table on a jobsnepal.com detail page.
    Returns dict with keys: position_type, city, posted_date, apply_before,
    experience, education, openings, salary, category.

    Table structure (verified from live page):
      <table>
        <tr><td>Category</td><td>...</td></tr>
        <tr><td>Openings</td><td>...</td></tr>
        <tr><td>Position Type</td><td><a href="/job-type/full-time">Full Time</a></td></tr>
        <tr><td>Experience</td><td>...</td></tr>
        <tr><td>Education</td><td>...</td></tr>
        <tr><td>Posted Date</td><td>25 Jun, 2026</td></tr>
        <tr><td>Apply Before</td><td>30 Jun, 2026</td></tr>
        <tr><td>City</td><td><a href="/city/kathmandu">Kathmandu</a></td></tr>
      </table>
    """
    result = {k: "" for k in ["position_type", "city", "posted_date",
                               "apply_before", "experience", "education",
                               "openings", "salary", "category"]}

    # Find the overview table — it's typically the only table on the page
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        for row in rows:
            cells = row.find_all(["td", "th"])
            if len(cells) < 2:
                continue
            label = clean_text(cells[0]).lower().strip().rstrip(":")
            value = clean_text(cells[1])
            if not value:
                continue
            if "category" in label:
                result["category"] = value
            elif "openings" in label or "opening" in label:
                result["openings"] = value
            elif "position type" in label or "job type" in label:
                result["position_type"] = value
            elif "experience" in label:
                result["experience"] = value
            elif "education" in label:
                result["education"] = value
            elif "posted date" in label or "date posted" in label:
                result["posted_date"] = value
            elif "apply before" in label or "apply by" in label or "deadline" in label:
                result["apply_before"] = value
            elif "city" in label or "location" in label:
                result["city"] = value
            elif "salary" in label or "offered salary" in label:
                result["salary"] = value
    return result

def parse_job_page(hint):
    """
    Parse a single jobsnepal.com job detail page.

    Verified selectors (from live fetch of /program-officer-142969 and /supervisor-142973):

    Title:          h1 (the only h1 on the page)
    Company:        h2 > a[href*="/employer/"] text
    Company logo:   meta[property="og:image"] → img.jobsnepal.com/big/HASH.ext
                    (listing page img is always gray-back.jpg placeholder — skip it)
    Company blurb:  <p> element between company h2 and the "Details/requirements" h2
                    (visible above-fold blurb; some jobs have it, some don't)
    Posted date:    "Job posted on DD Mon, YYYY" text near h1, OR Overview table
    Deadline:       Overview table "Apply Before" row — most reliable
                    Fallback: "Apply before: DD Mon, YYYY" text near company h2
                    Fallback: extract_deadline_from_text(description)
    Job description: div containing all body content after the h2 "Details / requirements:"
                    Identify by finding the h2 with that text and taking next sibling div
    Apply contact:  Inside description — a[href^="mailto:"] for email,
                    or external a[href^="http"] that isn't a jobsnepal.com link
    Overview table: parsed by _parse_overview_table()
    """
    url   = hint["job_url"]
    soup  = get_soup(url)

    # ── Title ─────────────────────────────────────────────────────────────────
    title = ""
    h1 = soup.find("h1")
    if h1:
        title = clean_text(h1)
    if not title:
        og = soup.find("meta", property="og:title")
        if og:
            title = og.get("content", "").strip()
            title = re.sub(r"\s*[-–|]\s*.*$", "", title).strip()
    # Fall back to hint title from listing page
    if not title:
        title = hint.get("title", "")

    # ── Company ───────────────────────────────────────────────────────────────
    company = ""
    # Primary: h2 with an <a href="/employer/...">
    employer_link = soup.select_one("h2 a[href*='/employer/']")
    if employer_link:
        company = clean_text(employer_link)
    if not company:
        company = hint.get("company", "")

    # ── Company logo ──────────────────────────────────────────────────────────
    # jobsnepal.com uploads all company images to img.jobsnepal.com/big/
    # The og:image meta is the most reliable source (always the big version)
    logo = ""
    og_img = soup.find("meta", property="og:image")
    if og_img:
        src = og_img.get("content", "").strip()
        if src and JOBSNEPAL_PLACEHOLDER not in src and "img.jobsnepal.com" in src:
            logo = src
    # Fallback: look for img.jobsnepal.com/big/ anywhere on the page
    if not logo:
        for img in soup.find_all("img", src=True):
            src = img["src"]
            if "img.jobsnepal.com/big/" in src and JOBSNEPAL_PLACEHOLDER not in src:
                logo = src
                break

    # ── Company short blurb ───────────────────────────────────────────────────
    # On jobsnepal.com the blurb is a <p> between the company h2 and the
    # "Details / requirements:" h2. It's visible above the fold on the page.
    company_details = ""
    details_h2 = None
    for h2 in soup.find_all("h2"):
        if re.search(r"details.*requirements", clean_text(h2), re.I):
            details_h2 = h2
            break

    if details_h2:
        # Walk backwards from details_h2 collecting <p> siblings until we hit
        # the company h2 or the h1
        blurb_parts = []
        for sib in details_h2.find_all_previous():
            tag = sib.name
            if tag in ("h1", "h2"):
                break
            if tag == "p":
                txt = clean_text(sib)
                if txt and len(txt) > 20:
                    blurb_parts.insert(0, txt)
        company_details = " ".join(blurb_parts).strip()

    # ── Posted date ───────────────────────────────────────────────────────────
    # "Job posted on 25 Jun, 2026" appears in a <p> near the h1.
    # Must stop capture before "Apply before ..." which appears in the same element.
    posted_date = ""
    for tag in soup.find_all(["p", "span", "div"]):
        txt = clean_text(tag)
        m = re.search(
            r"[Jj]ob\s+posted\s+on\s+"
            r"(\d{1,2}\s+\w+,?\s+\d{4})",   # "25 Jun, 2026" — stop at comma+year
            txt,
        )
        if m:
            posted_date = m.group(1).strip().rstrip(",")
            break

    # ── Overview table ────────────────────────────────────────────────────────
    overview = _parse_overview_table(soup)

    # Fill posted_date from table if not already found
    if not posted_date and overview["posted_date"]:
        posted_date = overview["posted_date"]
    posted_date = resolve_relative_date(posted_date)

    # ── Deadline ─────────────────────────────────────────────────────────────
    # Stage 1: Overview table "Apply Before" — most reliable
    deadline = overview["apply_before"]

    # Stage 2: "Apply before: DD Mon, YYYY" text near company h2
    if not deadline:
        for tag in soup.find_all(["p", "span", "div"]):
            txt = clean_text(tag)
            m = re.search(r"[Aa]pply\s+before:?\s+([\w ,]+\d{4})", txt)
            if m:
                deadline = m.group(1).strip()
                break

    # ── Location ─────────────────────────────────────────────────────────────
    location = overview["city"] or hint.get("location", "")

    # ── Job type ─────────────────────────────────────────────────────────────
    job_type = overview["position_type"] or ""

    # ── Salary ───────────────────────────────────────────────────────────────
    salary = overview["salary"]

    # ── Category ─────────────────────────────────────────────────────────────
    category = overview["category"] or hint.get("category", "")

    # ── Job description ───────────────────────────────────────────────────────
    # The description body follows the h2 "Details / requirements:" heading.
    description = ""
    if details_h2:
        # Collect all sibling content after the h2 until we hit another h2
        # (e.g. "Overview" or "Similar Jobs")
        parts = []
        for sib in details_h2.find_next_siblings():
            if sib.name == "h2":
                break
            txt = clean_text(sib)
            if txt:
                parts.append(txt)
        description = clean_description("\n\n".join(parts))

    # Fallback: largest text block on page
    if not description:
        best, best_len = "", 0
        for div in soup.find_all("div"):
            t = clean_text(div)
            if len(t) > best_len and len(t) > 200:
                best, best_len = t, len(t)
        description = clean_description(best)

    # Stage 3 deadline: scan description body text
    if not deadline:
        deadline = extract_deadline_from_text(description)

    # ── Apply contact ─────────────────────────────────────────────────────────
    # Priority: email > URL.
    # jobsnepal.com uses three link patterns:
    #   1. <a href="mailto:email@org.np">
    #   2. <a href="https://gmail.com"><u>email@org.np</u></a>  ← most common trap
    #   3. <a href="https://company.org/apply">
    # We always prefer an email address. URL is only used when no email found anywhere.

    raw_apply      = ""    # will hold email string or URL string
    raw_anchor_txt = ""    # visible text of the apply anchor

    def _extract_from_anchor(a_tag):
        """Return (value, anchor_text) — value is an email string or href URL."""
        href = a_tag.get("href", "")
        text = clean_text(a_tag)
        if href.lower().startswith("mailto:"):
            email = re.sub(r"^mailto:", "", href, flags=re.I).split("?")[0].strip()
            return email, text
        if _is_generic_mail_href(href):
            m = EMAIL_RE.search(text)
            if m:
                return m.group(0), text
            return "", text
        return href, text

    candidate_url = ""   # best external URL found (used only if no email anywhere)

    search_containers = []
    if details_h2:
        sib = details_h2.find_next_sibling()
        if sib:
            search_containers.append(sib)
    # Always also check whole page as fallback
    search_containers.append(soup)

    for container in search_containers:
        if raw_apply:   # already found an email
            break
        for a in container.find_all("a", href=True):
            href = a.get("href", "")
            # Email sources: mailto: or generic webmail href
            if href.lower().startswith("mailto:") or _is_generic_mail_href(href):
                val, txt = _extract_from_anchor(a)
                if val:
                    raw_apply      = val
                    raw_anchor_txt = txt
                    break
            # URL source: external, not jobsnepal, not doc host — store but keep scanning
            if (not candidate_url
                    and href.startswith("http")
                    and "jobsnepal.com" not in href
                    and not _is_document_host(href)
                    and not _is_generic_mail_href(href)):
                candidate_url  = href
                raw_anchor_txt = clean_text(a)

    # If no email link found, fall back to the best URL candidate
    if not raw_apply and candidate_url:
        raw_apply = candidate_url

    log(f"    Resolving apply contact for '{title}' …")
    # resolve_application_contact will also scan description text for email
    # before accepting any URL, enforcing the email-first rule end-to-end
    application = resolve_application_contact(raw_apply, description, raw_anchor_txt)

    return {
        "title":           title,
        "job_url":         url,
        "job_type":        job_type,
        "location":        location,
        "category":        category,
        "salary":          salary,
        "experience":      overview["experience"],
        "education":       overview["education"],
        "openings":        overview["openings"],
        "posted_date":     posted_date,
        "deadline":        deadline,
        "description":     description,
        "apply_url":       application["apply_url"],
        "apply_email":     application["apply_email"],
        "apply_raw":       application["apply_raw"],
        "company_name":    company,
        "company_details": company_details,
        "company_logo":    logo,
    }

# =============================================================================
#  STEP 3 — DEDUPLICATE + PUBLIC-APPLY RULE + PARAPHRASE
# =============================================================================

def process_job(raw_job, processed_ids, processed_urls, seen_content):
    job_url = raw_job.get("job_url", "")
    title   = raw_job.get("title", "")
    company = raw_job.get("company_name", "")
    job_id  = make_job_id(job_url, title, company)

    # Dedup: already in tracker
    if job_id in processed_ids or job_url in processed_urls:
        log(C_DIM(f"  ⧳ Already processed — skipped: {job_url}"))
        return None

    # Dedup: same content seen this run
    fp = (title.lower().strip(), company.lower().strip(),
          raw_job.get("location", "").lower().strip())
    if fp in seen_content:
        log(C_DIM(f"  ⧳ Duplicate this run — skipped: {title}"))
        return None
    seen_content.add(fp)

    # Public-apply rule: no email or external URL → flag, never post
    apply_url   = raw_job.get("apply_url", "")
    apply_email = raw_job.get("apply_email", "")
    if not apply_url and not apply_email:
        log(C_RED(f"  ⚑ FLAGGED (no apply contact): {title}"))
        flag_job(job_id, job_url, title, company)
        mark_scraped(job_id, job_url, title, company)
        _upsert_row(job_id, {"Status": "flagged|no_apply_contact"})
        processed_ids.add(job_id)
        processed_urls.add(job_url)
        return None

    mark_scraped(job_id, job_url, title, company)
    processed_ids.add(job_id)
    processed_urls.add(job_url)

    description = raw_job.get("description", "")
    blurb       = raw_job.get("company_details", "")

    para_title = title
    para_desc  = description
    para_blurb = blurb

    if ENABLE_PARAPHRASE and MISTRAL_API_KEY:
        print(C_BLUE(f"\n  ✍️  Paraphrasing '{title}' …"))
        para_title = paraphrase_title(title)
        para_desc  = paraphrase_description(description)
        if blurb:
            para_blurb = paraphrase_company(blurb)
        mark_paraphrased(job_id)
    else:
        print(C_DIM("  ⚠️  Paraphrasing skipped"))

    application = apply_url or apply_email

    company_website = ""
    if apply_url:
        try:
            parts = urlsplit(apply_url)
            if parts.scheme and parts.netloc and "jobsnepal.com" not in parts.netloc:
                company_website = f"{parts.scheme}://{parts.netloc}"
        except Exception:
            pass

    apply_method = ("resolved_redirect" if apply_url
                    else "description_email" if apply_email
                    else "not_found")

    return {
        "jobTitle":          para_title,
        "jobDescription":    para_desc,
        "companyDetails":    para_blurb,
        "originalTitle":     title,
        "originalDesc":      description,
        "jobType":           raw_job.get("job_type", ""),
        "jobLocation":       raw_job.get("location", ""),
        "category":          raw_job.get("category", ""),
        "salary":            raw_job.get("salary", ""),
        "experience":        normalise_experience(raw_job.get("experience", "")),
        "education":         normalise_qualification(raw_job.get("education", "")),
        "openings":          raw_job.get("openings", ""),
        "datePosted":        raw_job.get("posted_date", ""),
        "deadline":          raw_job.get("deadline", ""),
        "application":       application,
        "companyName":       company,
        "companyLogo":       raw_job.get("company_logo", ""),
        "companyWebsite":    company_website,
        "jobUrl":            job_url,
        "jobField":          normalise_job_field(
                                 raw_job.get("category", ""),
                                 title,
                                 description,
                             ),
        "_jobId":            job_id,
        "_apply_method":     apply_method,
        "_apply_raw":        raw_job.get("apply_raw", ""),
    }

# =============================================================================
#  VERBOSE PRINTER
# =============================================================================

def print_job_verbose(index, job):
    desc = job.get("jobDescription", "")
    desc_preview = (desc[:400] + " [...]") if len(desc) > 400 else desc
    print()
    print(C_DIVIDER())
    print(C_HEADER(f"  JOB #{index}"))
    print(C_DIVIDER())
    print(f"  {C_LABEL('Title (original)')}    : {C_VALUE(job.get('originalTitle',''))}")
    print(f"  {C_LABEL('Title (paraphrased)')} : {C_GREEN(job.get('jobTitle',''))}")
    print(f"  {C_LABEL('Job Type')}             : {job.get('jobType','') or C_DIM('—')}")
    print(f"  {C_LABEL('Job Field')}            : {job.get('jobField','') or C_DIM('—')}")
    print(f"  {C_LABEL('Category')}             : {job.get('category','') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}             : {job.get('jobLocation','') or C_DIM('—')}")
    print(f"  {C_LABEL('Salary')}               : {job.get('salary','') or C_DIM('—')}")
    print(f"  {C_LABEL('Experience')}           : {job.get('experience','') or C_DIM('—')}")
    print(f"  {C_LABEL('Education')}            : {job.get('education','') or C_DIM('—')}")
    print(f"  {C_LABEL('Openings')}             : {job.get('openings','') or C_DIM('—')}")
    print(f"  {C_LABEL('Posted')}               : {job.get('datePosted','') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}             : {job.get('deadline','') or C_DIM('—')}")
    application = job.get("application", "")
    print(f"  {C_LABEL('Apply')}                : "
          f"{C_GREEN(application) if application else C_DIM('— not found —')}")
    print(f"  {C_LABEL('Apply Method')}         : {C_DIM(job.get('_apply_method',''))}")
    print()
    print(f"  {C_BLUE('── COMPANY ──────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}      : {C_VALUE(job.get('companyName','') or C_DIM('—'))}")
    print(f"  {C_LABEL('Website')}   : {job.get('companyWebsite','') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}      : {job.get('companyLogo','') or C_DIM('— none —')}")
    about = job.get("companyDetails", "")
    if about:
        preview = (about[:200] + " [...]") if len(about) > 200 else about
        print(f"  {C_LABEL('About')}     : {preview}")
    print()
    print(f"  {C_BLUE('── DESCRIPTION PREVIEW ─────────────────────────────')}")
    print(desc_preview if desc_preview else C_DIM("   — no description —"))
    print(f"  {C_LABEL('Job URL')}   : {job.get('jobUrl','')}")
    print(C_DIVIDER())

# =============================================================================
#  WORDPRESS POSTING
# =============================================================================

def _wp_auth_headers():
    token = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}

def _wp_get(path, params=None):
    return requests.get(path, params=params, headers=_wp_auth_headers(),
                        auth=(WP_USER, WP_PASSWORD), timeout=15, verify=False)

def _wp_post(path, json_data):
    return requests.post(path, json=json_data, headers=_wp_auth_headers(),
                         auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)

def _wp_json(r, context=""):
    if not r.content:
        log_.error(f"WP API empty response [{r.status_code}] {context}")
        return None
    try:
        return r.json()
    except Exception:
        log_.error(f"WP API non-JSON [{r.status_code}] {context}: {r.text[:200]}")
        return None

def probe_wp_api():
    """
    Auto-discover the correct WP REST API base URL and verify the
    job-listings endpoint is reachable (requires WP Job Manager plugin).
    """
    root_domain = re.sub(r"^(https?://)[^.]+\.", r"\1", WP_BASE)
    candidates = list(dict.fromkeys(filter(None, [
        WP_API_BASE,
        f"{root_domain.rstrip('/')}/wp-json/wp/v2",
        f"{WP_BASE.rstrip('/')}/?rest_route=/wp/v2",
    ])))
    for base in candidates:
        test_url = f"{base}/job-listings"
        try:
            r = requests.get(test_url, params={"per_page": 1},
                             headers=_wp_auth_headers(),
                             auth=(WP_USER, WP_PASSWORD),
                             timeout=10, verify=False)
            if r.status_code in (200, 401):
                log_.info(f"✅ WP job-listings endpoint reachable: {base} [{r.status_code}]")
                return base
            elif r.status_code == 404:
                log_.warning(
                    f"WP job-listings 404 at {base} — is WP Job Manager installed & activated?"
                )
        except Exception as e:
            log_.debug(f"WP probe error {base}: {e}")
    log_.error(
        "❌ WP Job Manager REST endpoint unreachable.\n"
        "   Check: WP Job Manager plugin installed, WP_BASE_URL correct, pretty permalinks on."
    )
    return ""

def get_or_create_term(taxonomy_url, name):
    if not name or not name.strip():
        return None
    slug = re.sub(r"-{2,}", "-", re.sub(r"[^a-z0-9-]", "-", name.lower().strip())).strip("-")
    try:
        r    = _wp_get(taxonomy_url, params={"slug": slug})
        data = _wp_json(r, f"GET term slug={slug}")
        if isinstance(data, list) and data:
            return data[0]["id"]
    except Exception:
        pass
    try:
        r    = _wp_post(taxonomy_url, {"name": name, "slug": slug})
        data = _wp_json(r, f"POST term name={name}")
        if data and isinstance(data, dict):
            return data.get("id")
    except Exception as e:
        log_.error(f"Term create error '{name}': {e}")
    return None

def build_jsonld(job):
    """Build a JSON-LD JobPosting schema block."""
    application = job.get("application", "")
    apply_email = ""
    apply_url   = ""
    if re.match(r"^[A-Za-z0-9.+_-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$", application):
        apply_email = application
    elif application.startswith("http"):
        apply_url = application

    schema = {
        "@context":    "https://schema.org",
        "@type":       "JobPosting",
        "title":       sanitize_text(job.get("jobTitle", "")),
        "description": sanitize_text(job.get("jobDescription", "")),
        "datePosted":  sanitize_text(job.get("datePosted", "")),
        "hiringOrganization": {
            "@type": "Organization",
            "name":  sanitize_text(job.get("companyName", "")),
        },
        "jobLocation": {
            "@type":   "Place",
            "address": {
                "@type":             "PostalAddress",
                "addressLocality":   sanitize_text(job.get("jobLocation", "")),
                "addressCountry":    "NP",
            },
        },
    }
    if job.get("deadline"):
        schema["validThrough"] = sanitize_text(job["deadline"])
    if job.get("jobType"):
        schema["employmentType"] = sanitize_text(job["jobType"]).upper().replace(" ", "_")
    if job.get("salary"):
        schema["baseSalary"] = {"@type": "MonetaryAmount", "currency": "NPR",
                                "value": sanitize_text(job["salary"])}
    if apply_url:
        schema["url"] = apply_url
    if apply_email:
        schema["applicationContact"] = {"@type": "ContactPoint", "email": apply_email}
    if job.get("companyWebsite"):
        schema["hiringOrganization"]["sameAs"] = sanitize_text(job["companyWebsite"], is_url=True)
    return f'\n\n<script type="application/ld+json">\n{json.dumps(schema, indent=2, ensure_ascii=False)}\n</script>\n'

def post_job_to_wordpress(job):
    if not WP_USER or not WP_PASSWORD:
        return None, None

    title       = sanitize_text(job.get("jobTitle", ""))
    description = sanitize_text(job.get("jobDescription", ""))
    if not title or not description:
        log_.warning("Skipping WP post — missing title or description")
        return None, None

    slug = re.sub(r"-{2,}", "-", re.sub(r"[^a-z0-9-]", "-", title.lower())).strip("-")[:80]

    # Duplicate check
    try:
        r     = _wp_get(WP_JOBS_URL, params={"slug": slug, "status": "publish"})
        posts = _wp_json(r, f"duplicate check slug={slug}")
        if isinstance(posts, list) and posts:
            log_.info(f"⏭ Already on WP: {title}")
            return posts[0]["id"], posts[0].get("link")
    except Exception:
        pass

    logo_url    = sanitize_text(job.get("companyLogo", ""),    is_url=True)
    location    = sanitize_text(job.get("jobLocation", ""))
    raw_type    = sanitize_text(job.get("jobType", ""))        or "Full Time"
    job_type_s  = JOB_TYPE_MAPPING.get(raw_type.lower().strip(), "full-time")
    company     = sanitize_text(job.get("companyName", ""))
    application = sanitize_text(job.get("application", ""),    is_url=True)
    deadline    = sanitize_text(job.get("deadline", ""))
    co_website  = sanitize_text(job.get("companyWebsite", ""), is_url=True)
    about       = sanitize_text(job.get("companyDetails", ""))
    date_posted = sanitize_text(job.get("datePosted", ""))
    salary      = sanitize_text(job.get("salary", ""))
    category    = sanitize_text(job.get("category", ""))
    qualif      = sanitize_text(job.get("education", ""))      # already normalised
    experience  = sanitize_text(job.get("experience", ""))     # already normalised
    job_field   = sanitize_text(job.get("jobField", ""))

    # Validate application value
    is_email_app = bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", application))
    is_url_app   = bool(re.match(r"^https?://[^\s]+$", application))
    if not (is_email_app or is_url_app):
        application = ""

    # Upload logo
    attachment_id = None
    if logo_url:
        try:
            img_r = SESSION.get(logo_url, timeout=15)
            if img_r.status_code == 200:
                ct  = img_r.headers.get("Content-Type", "image/jpeg")
                ext = ("png" if "png" in ct else "gif" if "gif" in ct
                       else "webp" if "webp" in ct else "jpg")
                fn  = re.sub(r"-{2,}", "-",
                             re.sub(r"[^a-z0-9]", "-", company.lower())).strip("-") + f"-logo.{ext}"
                up_h = {**_wp_auth_headers(),
                        "Content-Disposition": f'attachment; filename="{fn}"',
                        "Content-Type": ct}
                up_r = requests.post(WP_MEDIA_URL, headers=up_h, data=img_r.content,
                                     auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
                if up_r.status_code in (200, 201):
                    up_data = _wp_json(up_r, "media upload")
                    if up_data:
                        attachment_id = up_data.get("id")
        except Exception as e:
            log_.warning(f"Logo upload failed: {e}")

    # WP Job Manager taxonomy terms
    # job_listing_type  → employment type (Full Time, Contract, etc.)
    # job_listing_category → job category / field
    # job_listing_tag  → additional tags (company name, source, etc.)
    type_term_id = get_or_create_term(WP_TAX_TYPE_URL, job_type_s.replace("-", " ").title()) if job_type_s else None
    cat_term_id  = get_or_create_term(WP_TAX_CAT_URL,  job_field) if job_field else None
    tag_names    = list(filter(None, [
        category.split(",")[0].strip() if category else None,
        company or None,
        "JobsNepal",
    ]))
    tag_ids = [tid for name in tag_names
               for tid in [get_or_create_term(WP_TAX_TAG_URL, name)] if tid]

    expiry_comment = f"<!-- job-expiry: {deadline} -->" if deadline else ""
    content = description + "\n\n" + expiry_comment + build_jsonld(job)

    payload = {
        "title":          title,
        "content":        content,
        "slug":           slug,
        "status":         "publish",
        "featured_media": attachment_id or 0,
        "meta": {
            "_job_title":          title,
            "_job_location":       location,
            "_job_type":           job_type_s,
            "_job_description":    description,
            "_application":        application,
            "_company_url":        co_website,        # alias expected by some WP Job Manager versions
            "_job_expires":        deadline,
            "_company_name":       company,
            "_company_website":    co_website,
            "_company_logo":       str(attachment_id) if attachment_id else "",
            "_company_industry":   job_field,         # best proxy we have for industry
            "_company_address":    location,          # city-level address from Overview table
            "_company_founded":    "",                # not available on jobsnepal.com
            "_company_type":       "",                # not available on jobsnepal.com
            "_company_tagline":    "",                # not available on jobsnepal.com
            "_company_details":    about,
            "_job_qualifications": qualif,
            "_job_experiences":    experience,
            "_job_field":          job_field,
            "_job_salary":         salary,
            "_date_posted":        date_posted,
        },
    }
    if type_term_id:
        payload["job_listing_type"]     = [type_term_id]
    if cat_term_id:
        payload["job_listing_category"] = [cat_term_id]
    if tag_ids:
        payload["job_listing_tag"]      = tag_ids

    h = _wp_auth_headers()
    for attempt in range(3):
        try:
            r = requests.post(WP_JOBS_URL, json=payload, headers=h,
                              auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
            r.raise_for_status()
            post = _wp_json(r, f"create post '{title}'")
            if not post:
                raise ValueError("Empty/invalid JSON in post response")
            log_.info(f"✅ Posted: '{title}' → WP ID {post.get('id')}")
            return post.get("id"), post.get("link")
        except Exception as e:
            log_.error(f"WP post attempt {attempt+1} failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)
    return None, None
# =============================================================================
#  EXCEL EXPORT
# =============================================================================

EXCEL_HEADERS = [
    "Job Title", "Job Type", "Job Field", "Category", "Job Location", "Salary",
    "Experience", "Education", "Openings", "Date Posted", "Deadline",
    "Job Description", "Application", "Apply Method",
    "Company Name", "Company Logo", "Company Website", "Company Details",
    "Job URL",
]

def _save_excel(jobs):
    if not _XLSX_AVAILABLE:
        log_.warning("openpyxl not installed — skipping Excel export")
        return
    if not jobs:
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nepal Jobs"
    ws.append(EXCEL_HEADERS)
    for job in jobs:
        ws.append([
            job.get("jobTitle", ""),      job.get("jobType", ""),
            job.get("jobField", ""),      job.get("category", ""),
            job.get("jobLocation", ""),   job.get("salary", ""),
            job.get("experience", ""),    job.get("education", ""),
            job.get("openings", ""),      job.get("datePosted", ""),
            job.get("deadline", ""),      job.get("jobDescription", ""),
            job.get("application", ""),   job.get("_apply_method", ""),
            job.get("companyName", ""),   job.get("companyLogo", ""),
            job.get("companyWebsite",""), job.get("companyDetails", ""),
            job.get("jobUrl", ""),
        ])
    wb.save(OUTPUT_FILE)
    log_.info(f"Saved {len(jobs)} rows → {OUTPUT_FILE}")
    if not _XLSX_AVAILABLE:
        log_.warning("openpyxl not installed — skipping Excel export")
        return
    if not jobs:
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nepal Jobs"
    ws.append(EXCEL_HEADERS)
    for job in jobs:
        ws.append([
            job.get("jobTitle", ""),     job.get("jobType", ""),
            job.get("category", ""),     job.get("jobLocation", ""),
            job.get("salary", ""),       job.get("experience", ""),
            job.get("education", ""),    job.get("openings", ""),
            job.get("datePosted", ""),   job.get("deadline", ""),
            job.get("jobDescription",""),job.get("application", ""),
            job.get("_apply_method",""), job.get("companyName", ""),
            job.get("companyLogo", ""),  job.get("companyWebsite", ""),
            job.get("companyDetails",""),job.get("jobUrl", ""),
        ])
    wb.save(OUTPUT_FILE)
    log_.info(f"Saved {len(jobs)} rows → {OUTPUT_FILE}")

# =============================================================================
#  MAIN
# =============================================================================

def main():
    # Must be declared before any read or write of these module-level variables
    global WP_API_BASE, WP_JOBS_URL, WP_MEDIA_URL, WP_TAX_TYPE_URL, WP_TAX_TAG_URL, WP_TAX_CAT_URL
    start_time = datetime.now()
    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  JOBSNEPAL.COM SCRAPER + MISTRAL PARAPHRASE + WORDPRESS POSTING"))
    print(C_HEADER("=" * 80))
    print(f"  Source          : {JOBS_LIST_URL}")
    print(f"  Request delay   : {REQUEST_DELAY}s")
    print(f"  Max new jobs    : {'unlimited' if not MAX_JOBS else MAX_JOBS}")
    print(f"  Max pages       : {'unlimited' if not MAX_PAGES else MAX_PAGES}")
    print(f"  Resolve apply   : {'✅ enabled' if RESOLVE_APPLY_URLS else '❌ disabled'}")
    print(f"  Paraphrase      : {'✅ enabled' if (ENABLE_PARAPHRASE and MISTRAL_API_KEY) else '❌ disabled'}")
    print(f"  WordPress post  : {'✅ enabled' if (WP_USER and WP_PASSWORD) else '❌ disabled'}")
    print(f"  Excel export    : {'✅ enabled' if _XLSX_AVAILABLE else '❌ disabled (pip install openpyxl)'}")
    print(f"  WP API base     : {WP_API_BASE or '(not configured)'}")
    print(f"  Started         : {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 80))

    # Auto-discover WP REST API endpoint
    if WP_USER and WP_PASSWORD:
        discovered = probe_wp_api()
        if discovered and discovered != WP_API_BASE:
            log(C_BLUE(f"  ℹ️  WP API base corrected to: {discovered}"))
            WP_API_BASE      = discovered
            WP_JOBS_URL      = f"{WP_API_BASE}/job-listings"
            WP_MEDIA_URL     = f"{WP_API_BASE}/media"
            WP_TAX_TYPE_URL  = f"{WP_API_BASE}/job_listing_type"
            WP_TAX_TAG_URL   = f"{WP_API_BASE}/job_listing_tag"
            WP_TAX_CAT_URL   = f"{WP_API_BASE}/job_listing_category"
        elif not discovered:
            log(C_RED("  ⚠️  WP posting will be skipped — REST API unreachable"))

    _init_tracker()
    processed_ids, processed_urls = load_processed_ids()
    print(f"  Tracker loaded: {len(processed_ids)} previously processed job IDs\n")

    # Collect all URLs (with early-stop pagination)
    hints = collect_job_urls(processed_urls)
    if not hints:
        print(C_RED("  No URLs collected — exiting."))
        return

    jobs_out      = []
    seen_content  = set()
    posted_count  = 0
    flagged_count = 0
    errors        = 0

    for i, hint in enumerate(hints, start=1):
        job_url = hint["job_url"]

        # Skip if already processed (quick check before fetching detail page)
        job_id_quick = make_job_id(job_url)
        if job_id_quick in processed_ids or job_url in processed_urls:
            log(C_DIM(f"\n⧳ [{i}/{len(hints)}] Already processed — skipped: {job_url}"))
            continue

        log(f"\nScraping [{i}/{len(hints)}]: {job_url}")
        try:
            raw_job = parse_job_page(hint)
        except Exception as e:
            errors += 1
            log(C_RED(f"  ERROR scraping {job_url}: {e}"))
            time.sleep(REQUEST_DELAY)
            continue

        try:
            job = process_job(raw_job, processed_ids, processed_urls, seen_content)
        except Exception as e:
            errors += 1
            log(C_RED(f"  ERROR processing job: {e}"))
            time.sleep(REQUEST_DELAY)
            continue

        if job is None:
            if not raw_job.get("apply_url") and not raw_job.get("apply_email"):
                flagged_count += 1
            time.sleep(REQUEST_DELAY)
            continue

        jobs_out.append(job)
        print_job_verbose(len(jobs_out), job)

        print(C_BLUE("\n  📤 Posting to WordPress …"))
        wp_id, wp_url = post_job_to_wordpress(job)
        if wp_id:
            mark_posted(job["_jobId"], wp_id)
            posted_count += 1
            print(C_GREEN(f"  ✅ WP ID={wp_id}  🔗 {wp_url}"))
        else:
            mark_failed(job["_jobId"], "wp_post_failed_or_skipped")
            print(C_RED("  ❌ WordPress post failed / skipped"))

        if len(jobs_out) % 25 == 0:
            _save_excel(jobs_out)

        if MAX_JOBS and len(jobs_out) >= MAX_JOBS:
            log(f"\nMAX_JOBS limit ({MAX_JOBS}) reached — stopping.")
            break

        time.sleep(REQUEST_DELAY)

    _save_excel(jobs_out)

    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds() / 60.0
    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 80))
    print(f"  {C_LABEL('Total hints collected')} : {len(hints)}")
    print(f"  {C_LABEL('New jobs processed')}    : {C_GREEN(str(len(jobs_out)))}")
    print(f"  {C_LABEL('Posted to WordPress')}   : {C_GREEN(str(posted_count))}")
    print(f"  {C_LABEL('Flagged (no apply)')}    : {C_RED(str(flagged_count)) if flagged_count else '0'}")
    print(f"  {C_LABEL('Errors')}                : {C_RED(str(errors)) if errors else '0'}")
    print(f"  {C_LABEL('Duration')}              : ~{duration:.1f} min")
    print(f"  {C_LABEL('Output file')}           : {OUTPUT_FILE}")
    print(f"  {C_LABEL('Tracker file')}          : {PROCESSED_IDS_FILE}")
    print(f"  {C_LABEL('Flagged file')}          : {FLAGGED_FILE}")
    print(C_HEADER("=" * 80))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
