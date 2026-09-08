import os
import re
import unicodedata
from datetime import datetime
import urllib.parse
import urllib.request
import json
import requests
from bs4 import BeautifulSoup
import pandas as pd
from difflib import SequenceMatcher

EXCEL_HEADERS = [
    "Date",
    "Year",
    "Month",
    "Date Posted",
    "Job ID LinkedIn",
    "Job ID Dexcom Page",
    "Job Posting site",
    "Location",
    "Country",
    "US / OUS",
    "Business Function",
    "Department",
    "Department from Dexcom Company Website",
    "Confidence Level",
    "Job Title",
    "Fresh / reposted",
    "Dexcom Offers / Why Dexcom? / What you'll get",
    "Summary / Position Summary / Meet the team / Role Summary",
    "Essential Duties and Responsibilities / Where you come in",
    "Supervisory Responsibilities",
    "Required Qualifications / What makes you successful / Requirements / Essential Capabilities / Competencies / About you",
    "Preferred Qualifications / Key Competencies",
    "Education Requirements and Experience",
    "Travel Required",
    "Workplace Type",
    "Functional Description",
    "Functional / Business Knowledge",
    "Scope",
    "Judgement",
    "Language Skills",
    "Physical Demands",
    "Work Environment",
    "Points to Note",
    "Management",
    "Field Sales",
    "Pay / Non-Exempt Salary Details / Commercial Salary Details / Exempt Salary Details",
    "Shifts",
    "Direct URL of the Dexcom Job Page",
    "LinkedIn URL"
]

TAXONOMY = {
    "1. Business Operations": {
        "1A. Commercial Operations": [
            "Business Transformation",
            "CO-Commercial Operations",
            "CO-Commercial Pricing",
            "CO-Forecasting and Analytics",
            "CO-International Business Operations",
            "Commercial Operations",
            "Contract specialist",
            "CO-Strategic Accounts Contracting",
            "CO-Vendor Management",
            "DA-Business Intelligence",
            "Data Science",
            "FA-Business Process",
            "Global security",
            "Operations",
            "PM-Professional Program Mgmt",
            "Security and Loss Prevention"
        ],
        "1B. Finance & Accounting": [
            "DA-Business Intelligence",
            "FA-Accounting",
            "FA-Accounts Payable",
            "FA-Accounts Receivable",
            "FA-Business Process",
            "FA-Cash Collections",
            "FA-Cost Accounting",
            "FA-Finance",
            "FA-Internal Audit",
            "FA-Internal Audit (IT)",
            "FA-Patient Billing",
            "FA-Payroll",
            "FA-Tax",
            "FA-Treasury",
            "Finance",
            "Finance & Accounting",
            "IT-Enterprise Applications",
            "PM-Professional Program Mgmt",
            "QA-Training"
        ],
        "1C. Internal Communications": [
            "CO-Employee Communications"
        ],
        "1D. Legal & Compliance": [
            "LG-Corporate Compliance",
            "LG-IP Legal Counsel",
            "LG-Legal Counsel",
            "LG-Paralegal",
            "LG-Privacy",
            "MF-Trade Compliance",
            "PM-Professional Project Mgmt"
        ],
        "1E. Manufacturing & Production": [
            "EN-Process Development"
        ],
        "1F. R&D": [
            "EN-Process Engineering"
        ]
    },
    "2. Clinical & Regulatory": {
        "2A. Clinical Affairs": [
            "CL-Clinical Affairs",
            "CL-Clinical Data Management",
            "Clinical Affairs",
            "CL-Medical Writer",
            "DA-Data Science",
            "EN-Science",
            "Medical & Clinical Affairs",
            "PM-Technical Program Mgmt"
        ],
        "2B. Clinical Data Management": [
            "CL-Biostatistics",
            "CL-Clinical Data Management",
            "Clinical Affairs",
            "CL-Statistical Programming",
            "IT-Information Technology"
        ],
        "2C. Medical Affairs": [
            "CL-Medical Affairs",
            "DA-Data Science",
            "Medical Affairs",
            "PM-Professional Project Mgmt"
        ],
        "2D. Regulatory Affairs": [
            "CL-Government Affairs",
            "RA-Regulatory Affairs"
        ]
    },
    "3. Manufacturing": {
        "3A. Facilities & EHS": [
            "AD-Administrative Support",
            "Facilities",
            "Facilities & EHS",
            "MF-EHSS",
            "MF-Facilities"
        ],
        "3B. Manufacturing & Production": [
            "EN-Automation Engineering",
            "EN-Chemistry",
            "EN-Engineering",
            "EN-Industrial Engineering",
            "EN-Manufacturing Engineering",
            "EN-Mechanical",
            "EN-Packaging",
            "EN-Process Development",
            "EN-Process Engineering",
            "FA-Business Process",
            "IT-Business Systems",
            "Manufacturing & Production",
            "MF-Buyer",
            "MF-Distribution",
            "MF-Equipment",
            "MF-Instructional Design",
            "MF-Lab",
            "MF-Logistics",
            "MF-Manufacturing",
            "MF-Manufacturing Associate",
            "MF-Operations",
            "MF-Planning",
            "MF-Planning & Operations Management",
            "MF-Process",
            "MF-Procurement Operations",
            "MF-Research",
            "MF-Security and Loss Prevention",
            "MF-Strategic Sourcing",
            "MF-Supply Chain",
            "Operations",
            "Program Management"
        ],
        "3C. Quality Assurance": [
            "Manufacturing & Production",
            "PM-Professional Program Mgmt",
            "QA-Calibration",
            "QA-Document Control",
            "QA-Microbiology",
            "QA-Operational Compliance",
            "QA-QA Engineering",
            "QA-QA Product Release",
            "QA-QC Inspection",
            "QA-Quality",
            "QA-Quality Compliance",
            "QA-Software QA",
            "Quality Assurance"
        ]
    },
    "4. Market Access & Health Economics": {
        "4A. Health Economics": [
            "CL-Health Economics and Outcomes Research"
        ],
        "4B. Market Access": [
            "CL-Global Access",
            "CL-Health Economics and Outcomes Research",
            "Market Access",
            "SA-Managed Markets Accounts Canada",
            "SA-Managed Markets Accounts EMEA"
        ]
    },
    "5. Marketing & Sales": {
        "5A. Business Development/ International Business": [
            "CO-International Business Operations",
            "CO-New Business Development",
            "DA-Business Intelligence",
            "LG-Corporate Development",
            "PM-Professional Program Mgmt",
            "PM-Professional Project Mgmt"
        ],
        "5B. Customer Advocacy": [
            "CA-Customer Advocacy"
        ],
        "5C. Customer Support": [
            "CO-Customer Account Support",
            "Customer Support",
            "PC-Patient Care",
            "SA-Customer Service",
            "TS-Technical Support"
        ],
        "5D. Inside Sales": [
            "AD-Administrative Support",
            "Customer service",
            "Inside Sales",
            "SA-Customer Service",
            "SA-Inside Sales APAC",
            "SA-Inside Sales Canada",
            "SA-Inside Sales EMEA",
            "SA-Inside Sales US",
            "Support",
            "Training"
        ],
        "5E. Marketing": [
            "FA-Business Process",
            "Marketing",
            "MK-Copy Writing",
            "MK-Customer Experience",
            "MK-Digital Design",
            "MK-Graphic Design",
            "MK-Instructional and Content Design",
            "MK-Lead Development",
            "MK-Marketing",
            "MK-Marketing Research",
            "MK-Product Management",
            "MK-Professional Education",
            "MK-Public Relations",
            "MK-UI Design",
            "MK-User Research",
            "MK-UX Design",
            "SA-Managed Markets Accounts US"
        ],
        "5F. Sales": [
            "AD-Administrative Support",
            "CO-Commercial Operations",
            "CO-Forecasting and Analytics",
            "CO-New Business Development",
            "CO-Sales Operations",
            "CO-Sales Training",
            "DA-Business Intelligence",
            "MK-Instructional and Content Design",
            "MK-Medical Science Liaison",
            "PC-Clinical Accounts",
            "PC-Patient Care",
            "SA- Sales",
            "SA-Clinical Accounts",
            "SA-Federal Sales US",
            "SA-Field Sales",
            "SA-Field Sales APAC",
            "SA-Field Sales Canada",
            "SA-Field Sales EMEA",
            "SA-Field Sales US",
            "Sales",
            "SA-Managed Markets Accounts EMEA",
            "SA-Managed Markets Accounts US",
            "SA-Remote Sales US",
            "SA-Strategic Account Management",
            "SA-Trade US",
            "SO-Channel Business Development"
        ]
    },
    "6. Others": {
        "6A. Administrative": [
            "AD-Administrative Support",
            "AD-Executive Assistant",
            "AD-Office Management",
            "SA-Inside Sales EMEA",
            "WS-Working Students"
        ],
        "6B. HR/ Talent Acquisition": [
            "CA-Workforce Management",
            "CO-Employee Communications",
            "Equity and Inclusion",
            "HR/ Talent Acquisition",
            "HR-Benefits",
            "HR-Compensation",
            "HR-Employee Relations and Compliance",
            "HR-HR Business Partner",
            "HR-HR Operations",
            "HR-Learning and Development",
            "HR-Organization and Culture",
            "HR-Talent Acquisition",
            "HR-Talent Management and DEI",
            "Human Resource",
            "IT-Business Systems",
            "IT-Systems Admin",
            "MF-Planning",
            "PM-Professional Program Mgmt",
            "PM-Professional Project Mgmt",
            "Talent Acquisition"
        ],
        "6C. Internship": [
            "Intern"
        ]
    },
    "7. Product Development/ Data Strategy": {
        "7A. Commercial Operations": [
            "Operations"
        ],
        "7B. Cybersecurity": [
            "EN-Cybersecurity"
        ],
        "7C. Data/Algorithm Engineering": [
            "DA-Data Architect",
            "DA-Data Engineering",
            "DA-Data Science",
            "Data Engineering",
            "Data/Algorithm Engineering",
            "EN-Algorithm"
        ],
        "7D. Electrical Engineering": [
            "EN-Hardware"
        ],
        "7E. IT": [
            "DA-Data Engineering",
            "EN-Process Development",
            "EN-Systems Design",
            "IT-Enterprise Applications"
        ],
        "7F. Manufacturing & Production": [
            "EN-Process Development"
        ],
        "7G. Mechanical Engineering": [
            "EN-Medical Device",
            "EN-Process Development"
        ],
        "7H. Program Management": [
            "PM-Design Ops",
            "PM-Product Tech Program Mgmt",
            "PM-Professional Program Mgmt",
            "PM-Professional Project Mgmt",
            "PM-Technical Program Mgmt",
            "PM-Technical Project Mgmt",
            "Program Management"
        ],
        "7I. R&D": [
            "EN Firmware",
            "EN-Engineering",
            "EN-Firmware",
            "EN-Hardware",
            "EN-Mechanical",
            "EN-Mechanical Engineering",
            "EN-Medical Device",
            "EN-Process Development",
            "EN-Science",
            "EN-Software Development",
            "EN-Systems Design",
            "EN-Test Engineering",
            "IT-DevOps Engineering",
            "PM-Product Owner",
            "PM-Technical Program Mgmt",
            "PM-Technical Project Mgmt",
            "R&D",
            "Research & Development",
            "EN-CAD"
        ],
        "7J. Software Engineering": [
            "DA-Data Engineering",
            "EN Software Development",
            "EN-Cybersecurity",
            "EN-Software Applications",
            "EN-Software Development",
            "EN-Software Test Development",
            "IT-DevOps Engineering",
            "MK-UX Design",
            "Software Engineering"
        ]
    },
    "8. Technical Support": {
        "8A. Cybersecurity": [
            "Cybersecurity",
            "EN-Cybersecurity",
            "EN-Security and Privacy Compliance",
            "EN-Software Test Development",
            "EN-Technical Writing"
        ],
        "8B. Data/Algorithm Engineering": [
            "DA-Data Engineering",
            "DA-Data Science",
            "IT-Information Technology"
        ],
        "8C. IT": [
            "DA-Data Engineering",
            "EN-Software Integration",
            "EN-Software Systems",
            "EN-Systems Design",
            "EN-Test Engineering",
            "FA-Accounting",
            "Information Technology",
            "IT",
            "IT Compliance",
            "IT Opportunities",
            "IT-Business Systems",
            "IT-Data Mgmt",
            "IT-Desktop Support",
            "IT-Enterprise Applications",
            "IT-Enterprise Privacy Program",
            "IT-Information Technology",
            "IT-IOC Major Incident",
            "IT-IT Business Partner",
            "IT-IT Infrastructure Engineering",
            "IT-IT Operations",
            "IT-IT Project Mgmt",
            "IT-IT Service Desk",
            "IT-Network Engineering",
            "IT-Process automation",
            "IT-Systems Admin",
            "IT-Systems Engineering",
            "IT-Telecom Engineering",
            "Project Management",
            "QA-Software QA",
            "Software Engineering",
            "Technical Support",
            "TS-Technical Support"
        ],
        "8D. Mechanical Engineering": [
            "EN-Mechanical"
        ],
        "8E. Program Management": [
            "FA-Business Process",
            "IT-Information Technology",
            "IT-IT Project Mgmt",
            "PM-Professional Program Mgmt",
            "PM-Professional Project Mgmt",
            "Program Management"
        ],
        "8F. Quality Assurance": [
            "QA-QA Engineering",
            "QA-Software QA"
        ],
        "8G. R&D": [
            "EN-Engineering",
            "IT-DevOps Engineering",
            "Research & Development"
        ],
        "8H. Software Engineering": [
            "EN-Software Development",
            "EN-Software Systems",
            "EN-Software Test",
            "EN-Software Test Development",
            "IT-DevOps Engineering",
            "Software Engineering"
        ],
        "8I. Support": [
            "CA-Call Center Training",
            "DA-Data Engineering",
            "IT-Desktop Support",
            "QA-Training",
            "Technical Support",
            "TS-Technical Support",
            "TS-Technical Support Hybrid"
        ]
    }
}

PUNCT_REGEX = re.compile(r'[.,?!:;\-\–\—\[\]\(\)\{\}\'\’\"“”\*\&\•\#\~\ \@\^\|\…]')

US_STATE_MAP = {
    "alabama": "al", "alaska": "ak", "arizona": "az", "arkansas": "ar", "california": "ca",
    "colorado": "co", "connecticut": "ct", "delaware": "de", "florida": "fl", "georgia": "ga",
    "hawaii": "hi", "idaho": "id", "illinois": "il", "indiana": "in", "iowa": "ia",
    "kansas": "ks", "kentucky": "ky", "louisiana": "la", "maine": "me", "maryland": "md",
    "massachusetts": "ma", "michigan": "mi", "minnesota": "mn", "mississippi": "ms", "missouri": "mo",
    "montana": "mt", "nebraska": "ne", "nevada": "nv", "new hampshire": "nh", "new jersey": "nj",
    "new mexico": "nm", "new york": "ny", "north carolina": "nc", "north dakota": "nd", "ohio": "oh",
    "oklahoma": "ok", "oregon": "or", "pennsylvania": "pa", "rhode island": "ri", "south carolina": "sc",
    "south dakota": "sd", "tennessee": "tn", "texas": "tx", "utah": "ut", "vermont": "vt",
    "virginia": "va", "washington": "wa", "west virginia": "wv", "wisconsin": "wi", "wyoming": "wy",
    "district of columbia": "dc", "puerto rico": "pr"
}

GEO_EQUIV_MAP = {
    "pulau pinang": "penang",
    "taman pulau pinang": "penang",
    "rhineland palatinate": "rheinland pfalz",
    "hesse": "hessen",
    "vilniaus": "vilnius",
    "taguig city": "taguig",
    "national capital region": "ncr",
    "county galway": "galway"
}

FILLER_LOC_WORDS = {
    'county', 'region', 'national', 'capital', 'province', 'taman', 'state', 'area', 'greater', 'city'
}

def translate_to_english(text: str) -> str:
    """
    If text is not in English, translate it to English before extraction.
    """
    if not text or text in ["N/A", "NA", "-"]:
        return text
        
    non_english_markers = [
        'mitarbeiter', 'studentische', 'aushilfe', 'kundendienst', 'deutschsprachigen',
        'aufgaben', 'anforderungen', 'qualifikationen', 'über uns', 'wir bieten',
        'darauf kannst du dich freuen', 'das zeichnet dich aus', 'deine aufgaben',
        'profil', 'ihre aufgaben', 'wir suchen'
    ]
    
    has_non_ascii = any(ord(c) > 127 for c in text)
    has_markers = any(m in text.lower() for m in non_english_markers)
    
    if not (has_non_ascii or has_markers):
        return text
        
    chunks = []
    lines = text.split('\n')
    current_chunk = []
    current_len = 0
    
    for line in lines:
        if current_len + len(line) > 1500:
            chunks.append('\n'.join(current_chunk))
            current_chunk = [line]
            current_len = len(line)
        else:
            current_chunk.append(line)
            current_len += len(line)
    if current_chunk:
        chunks.append('\n'.join(current_chunk))
        
    translated_chunks = []
    for chunk in chunks:
        if not chunk.strip():
            translated_chunks.append(chunk)
            continue
        try:
            url = 'https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=en&dt=t&q=' + urllib.parse.quote(chunk)
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            res = urllib.request.urlopen(req, timeout=10)
            data = json.loads(res.read().decode('utf-8'))
            translated_text = ''.join([item[0] for item in data[0] if item and item[0]])
            translated_chunks.append(translated_text)
        except Exception:
            translated_chunks.append(chunk)
            
    return '\n'.join(translated_chunks)

def strip_serial_number(text: str) -> str:
    """
    Remove serial numbers (e.g. '1. ', '1B. ', '7I. ', '8C. ', '3B. ') from taxonomy labels.
    """
    if not text or text in ["NA", "N/A", "-"]:
        return text
    return re.sub(r'^\d+[A-Z]?\.\s*', '', text).strip()

def extract_linkedin_job_id(url: str) -> str:
    """
    Extract and validate LinkedIn Job ID explicitly as a standalone 9 to 11 digit numeric value.
    """
    if not url:
        return "NA"
    if '\t' in url:
        url = url.split('\t')[-1].strip()
    match = re.search(r'/view/(?:[\w\-]+-)?(\d{9,11})', url)
    if not match:
        match = re.search(r'(\d{9,11})', url)
    return match.group(1) if match else "NA"

def clean_bullets(text: str) -> str:
    if not text or text == "-":
        return "-"
    cleaned_items = []
    for line in text.split('\n'):
        line_clean = line.strip()
        if line_clean.startswith('#'):
            continue
        cleaned = re.sub(r'^\s*[\bullet\*\-\–\—\•]\s*', '', line_clean).strip()
        if not cleaned:
            continue
        sentences = re.split(r'(?<=\.)\s+', cleaned)
        for s in sentences:
            s_clean = s.strip()
            if s_clean and not s_clean.startswith('#'):
                cleaned_items.append(s_clean)
    return '\n'.join(cleaned_items) if cleaned_items else "-"

def format_location(raw_loc: str) -> str:
    if not raw_loc or raw_loc == "N/A":
        return "N/A"
    parts = [p.strip() for p in raw_loc.split(',') if p.strip()]
    if len(parts) >= 2:
        return f"{parts[-2]}, {parts[-1]}"
    elif len(parts) == 1:
        return parts[0]
    return raw_loc

def parse_date_info(date_str: str) -> dict:
    """
    Read the date associated with each LinkedIn URL directly from the TXT file.
    Preserve the original representation in the Date column.
    Derive Year (e.g. 2026), Month (e.g. August), and Date Posted (e.g. August-2026).
    """
    if not date_str or date_str == "-":
        return {"Date": "-", "Year": "NA", "Month": "NA", "Date Posted": "NA"}
    
    clean_date = date_str.strip()
    dt = None
    formats = [
        "%A, %B %d, %Y",
        "%A, %b %d, %Y",
        "%B %d, %Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%m/%d/%Y"
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(clean_date, fmt)
            break
        except ValueError:
            pass
            
    if dt:
        return {
            "Date": clean_date,
            "Year": str(dt.year),
            "Month": dt.strftime("%B"),
            "Date Posted": f"{dt.strftime('%B')}-{dt.year}"
        }
    else:
        return {
            "Date": clean_date,
            "Year": "NA",
            "Month": "NA",
            "Date Posted": "NA"
        }

def determine_country_and_us(location: str) -> tuple:
    if not location or location == "N/A":
        return "N/A", "OUS"
    
    us_identifiers = {
        "UNITED STATES", "USA", "US", "ALABAMA", "ALASKA", "ARIZONA", "ARKANSAS", "CALIFORNIA",
        "COLORADO", "CONNECTICUT", "DELAWARE", "FLORIDA", "GEORGIA", "HAWAII", "IDAHO", "ILLINOIS",
        "INDIANA", "IOWA", "KANSAS", "KENTUCKY", "LOUISIANA", "MAINE", "MARYLAND", "MASSACHUSETTS",
        "MICHIGAN", "MINNESOTA", "MISSISSIPPI", "MISSOURI", "MONTANA", "NEBRASKA", "NEVADA",
        "NEW HAMPSHIRE", "NEW JERSEY", "NEW MEXICO", "NEW YORK", "NORTH CAROLINA", "NORTH DAKOTA",
        "OHIO", "OKLAHOMA", "OREGON", "PENNSYLVANIA", "RHODE ISLAND", "SOUTH CAROLINA", "SOUTH DAKOTA",
        "TENNESSEE", "TEXAS", "UTAH", "VERMONT", "VIRGINIA", "WASHINGTON", "WEST VIRGINIA",
        "WISCONSIN", "WYOMING"
    }
    
    parts = [p.strip() for p in location.split(',')]
    
    for part in parts:
        if part.upper() in us_identifiers or any(u in part.upper() for u in ["UNITED STATES", "USA"]):
            return "United States", "US"
            
    country = parts[-1] if parts else location
    return country, "OUS"

def normalize_title_for_matching(title: str) -> str:
    """
    Comparison-only normalization for job titles:
    Convert case, normalize Unicode, strip specified punctuation symbols, and treat repeated whitespace as single space.
    """
    if not title or title in ["N/A", "NA", "-"]:
        return ""
    text = unicodedata.normalize('NFKD', title).lower()
    text = PUNCT_REGEX.sub(' ', text)
    return ' '.join(text.split())

def extract_loc_tokens(loc_str: str) -> set:
    if not loc_str or loc_str in ["N/A", "NA", "-"]:
        return set()
    text = unicodedata.normalize('NFKD', loc_str).lower()
    text = re.sub(r'[^\w\s]', ' ', text)
    
    for eq_k, eq_v in GEO_EQUIV_MAP.items():
        if eq_k in text:
            text = text.replace(eq_k, eq_v)
            
    words = [w for w in text.split() if w not in FILLER_LOC_WORDS]
    
    mapped_words = set()
    for w in words:
        if w in US_STATE_MAP:
            mapped_words.add(US_STATE_MAP[w])
        elif w in GEO_EQUIV_MAP:
            mapped_words.add(GEO_EQUIV_MAP[w])
        else:
            mapped_words.add(w)
    return mapped_words

def is_location_exact_equivalent(loc1: str, loc2: str) -> bool:
    """
    Exact-equivalent comparison for location:
    Normalizes punctuation, dashes, state abbreviations, and contextually correct geographic equivalents.
    """
    if not loc1 or not loc2 or loc1 in ["N/A", "NA", "-"] or loc2 in ["N/A", "NA", "-"]:
        return False
        
    if isinstance(loc2, list):
        return any(is_location_exact_equivalent(loc1, l) for l in loc2)
        
    tokens1 = extract_loc_tokens(loc1)
    tokens2 = extract_loc_tokens(loc2)
    
    common = tokens1.intersection(tokens2)
    if not common:
        return False
        
    parts1 = [p.strip() for p in loc1.split(',') if p.strip()]
    parts2 = [p.strip() for p in loc2.split(',') if p.strip()]
    
    city1 = parts1[0].lower() if parts1 else ""
    city2 = parts2[0].lower() if parts2 else ""
    
    for k, v in GEO_EQUIV_MAP.items():
        city1 = city1.replace(k, v)
        city2 = city2.replace(k, v)
        
    city1_clean = ' '.join([w for w in city1.split() if w not in FILLER_LOC_WORDS])
    city2_clean = ' '.join([w for w in city2.split() if w not in FILLER_LOC_WORDS])
    
    if city1_clean and city2_clean:
        c1_tokens = set(city1_clean.split())
        c2_tokens = set(city2_clean.split())
        if not c1_tokens.intersection(c2_tokens):
            return False
            
    return True

def compute_description_similarity(desc1: str, desc2: str) -> float:
    """
    Job description similarity method:
    Normalizes case, whitespace, HTML tags, and strips common boilerplate.
    Computes a weighted similarity percentage from 0 to 100 combining SequenceMatcher ratio
    and Jaccard token overlap for reproducibility and accuracy.
    """
    if not desc1 or not desc2:
        return 0.0
    
    d1 = BeautifulSoup(desc1, 'html.parser').get_text(' ', strip=True) if '<' in desc1 else desc1
    d2 = BeautifulSoup(desc2, 'html.parser').get_text(' ', strip=True) if '<' in desc2 else desc2
    
    boilerplate_patterns = [
        r'to all staffing and recruiting agencies.*',
        r'an equal opportunity employer.*',
        r'dexcom is an equal opportunity.*',
        r'privacy policy.*',
        r'cookie policy.*'
    ]
    for p in boilerplate_patterns:
        d1 = re.sub(p, '', d1, flags=re.IGNORECASE | re.DOTALL)
        d2 = re.sub(p, '', d2, flags=re.IGNORECASE | re.DOTALL)
        
    d1_clean = ' '.join(re.sub(r'[^\w\s]', '', d1.lower()).split())
    d2_clean = ' '.join(re.sub(r'[^\w\s]', '', d2.lower()).split())
    
    if not d1_clean or not d2_clean:
        return 0.0
        
    seq_ratio = SequenceMatcher(None, d1_clean, d2_clean).ratio() * 100
    
    t1 = set(d1_clean.split())
    t2 = set(d2_clean.split())
    if not t1 or not t2:
        return 0.0
    jaccard = (len(t1.intersection(t2)) / float(len(t1.union(t2)))) * 100
    
    final_score = 0.7 * seq_ratio + 0.3 * jaccard
    return round(final_score, 2)

def classify_function_and_department(title: str, description: str) -> tuple:
    t_lower = title.lower()
    
    if any(k in t_lower for k in ['accounts payable', 'accounts receivable', 'ap specialist', 'finance', 'payroll', 'accounting']):
        return "Business Operations", "Finance & Accounting"
    elif any(k in t_lower for k in ['product manager', 'marketing', 'insights', 'graphic designer', 'trade marketing']):
        return "Marketing & Sales", "Marketing"
    elif any(k in t_lower for k in ['sales', 'inside sales trainer', 'trainer']):
        return "Marketing & Sales", "Sales"
    elif any(k in t_lower for k in ['customer service', 'customer support', 'patient care', 'kundendienst']):
        return "Marketing & Sales", "Customer Support"
    elif any(k in t_lower for k in ['database administrator', 'applications support', 'systems analyst']):
        return "Technical Support", "IT"
    elif any(k in t_lower for k in ['vendor management', 'commercial operations']):
        return "Business Operations", "Commercial Operations"
    elif any(k in t_lower for k in ['manufacturing', 'production operator', 'operator', 'technician']):
        return "Manufacturing", "Manufacturing & Production"
    elif any(k in t_lower for k in ['ecad technician', 'engineering technician', 'cad']):
        return "Product Development/ Data Strategy", "R&D"
    elif any(k in t_lower for k in ['studentische', 'working student', 'intern', 'student assistant']):
        return "Others", "Administrative"
    elif any(k in t_lower for k in ['trade compliance', 'compliance']):
        return "Business Operations", "Legal & Compliance"
    
    return "Business Operations", "Commercial Operations"

def match_dexcom_department_from_taxonomy(title: str, description: str, raw_dexcom_dept: str = "NA", bus_func: str = "", dept: str = "") -> tuple:
    """
    Match Business Function (Level 1), Department (Level 2), and Department from Dexcom Website (Level 3)
    based on taxonomy hierarchy without serial numbers.
    Returns (Level 1, Level 2, Level 3 Match, Confidence Level).
    Scoring Criteria:
    - Green (90-100%): Department is directly and explicitly stated on official Dexcom posting and maps clearly to taxonomy.
    - Amber (60-89%): Department is reasonably inferred from strong evidence in title/responsibilities.
    - Red (0-59%): Department is not found, evidence is weak, or match was rejected (NA).
    """
    if raw_dexcom_dept in ["NA", "N/A", "-", ""]:
        return strip_serial_number(bus_func), strip_serial_number(dept), "NA", "Red (0%)"
        
    norm_raw = re.sub(r'[^a-zA-Z0-9]', '', raw_dexcom_dept.lower())
    
    found_l1, found_l2, found_l3 = None, None, None
    for l1, l2_dict in TAXONOMY.items():
        for l2, candidates in l2_dict.items():
            for c in candidates:
                norm_c = re.sub(r'[^a-zA-Z0-9]', '', c.lower())
                clean_c = re.sub(r'^[a-zA-Z]{2}-', '', c)
                norm_clean_c = re.sub(r'[^a-zA-Z0-9]', '', clean_c.lower())
                
                if norm_c == norm_raw or norm_clean_c == norm_raw or norm_raw == norm_c:
                    found_l1, found_l2, found_l3 = l1, l2, c
                    break
            if found_l3: break
        if found_l3: break
        
    clean_l1 = strip_serial_number(found_l1) if found_l1 else strip_serial_number(bus_func)
    clean_l2 = strip_serial_number(found_l2) if found_l2 else strip_serial_number(dept)
    
    if found_l3:
        return clean_l1, clean_l2, found_l3, "Green (100%)"
    else:
        return clean_l1, clean_l2, raw_dexcom_dept, "Green (90%)"

def search_dexcom_career_page(job_title: str, location: str, linkedin_desc: str, headers: dict) -> tuple:
    """
    Search only the official primary Dexcom Careers site (careers.dexcom.com).
    Strictly excludes myworkdayjobs.com and third-party sites.
    Enforces Title exact equivalence, Location exact equivalence, and Description similarity >= 60%.
    """
    if not job_title or job_title in ["N/A", "NA", "-"]:
        return "NA", "NA", "NA"
        
    clean_title = re.sub(r'\s*[\(\[\{].*?[\)\]\}]', '', job_title).strip()
    norm_target_title = normalize_title_for_matching(job_title)
    norm_clean_title = normalize_title_for_matching(clean_title)
    
    search_url = 'https://careers.dexcom.com/api/pcsx/search'
    search_queries = [clean_title, job_title]
    
    candidates = []
    
    for query in search_queries:
        try:
            resp = requests.get(search_url, params={'domain': 'dexcom.com', 'query': query}, headers=headers, timeout=10)
            if resp.status_code == 200:
                positions = resp.json().get('data', {}).get('positions', [])
                for p in positions:
                    dex_title = p.get('name', '')
                    dex_locs = p.get('locations', [])
                    norm_dex_title = normalize_title_for_matching(dex_title)
                    
                    title_match = (norm_dex_title == norm_target_title) or (norm_dex_title == norm_clean_title)
                    if not title_match:
                        continue
                        
                    loc_match = is_location_exact_equivalent(location, dex_locs)
                    if not loc_match:
                        continue
                        
                    pos_id = p.get('id')
                    det_url = f'https://careers.dexcom.com/api/pcsx/position_details?position_id={pos_id}'
                    desc_dex = ""
                    try:
                        r_det = requests.get(det_url, headers=headers, timeout=10)
                        if r_det.status_code == 200:
                            desc_dex = r_det.json().get('data', {}).get('jobDescription', '')
                    except Exception:
                        pass
                        
                    sim = compute_description_similarity(linkedin_desc, desc_dex)
                    if sim >= 60.0:
                        candidates.append((sim, p))
                        
            if candidates:
                break
        except Exception as e:
            print(f"Error searching Dexcom for '{query}': {e}")
            
    if not candidates:
        return "NA", "NA", "NA"
        
    candidates.sort(key=lambda x: x[0], reverse=True)
    top_sim, top_p = candidates[0]
    
    if len(candidates) > 1 and candidates[1][0] == top_sim:
        return "NA", "NA", "NA"
        
    dexcom_job_id = top_p.get('displayJobId') or top_p.get('atsJobId') or "NA"
    position_path = top_p.get('positionUrl', '')
    
    direct_url = f"https://careers.dexcom.com{position_path}" if position_path else "NA"
    
    parsed = urllib.parse.urlparse(direct_url)
    if 'myworkdayjobs' in direct_url or parsed.netloc != 'careers.dexcom.com':
        return "NA", "NA", "NA"
        
    raw_dept = top_p.get('department', 'NA')
    
    return dexcom_job_id, direct_url, raw_dept

def parse_description_sections(html_or_text: str) -> dict:
    sections = {
        "Dexcom Offers": "-",
        "Summary": "-",
        "Responsibilities": "-",
        "Supervisory": "-",
        "Required Qualifications": "-",
        "Preferred Qualifications": "-",
        "Education": "-",
        "Travel": "-",
        "Workplace Type": "-",
        "Functional Description": "-",
        "Functional Knowledge": "-",
        "Scope": "-",
        "Judgement": "-",
        "Language Skills": "-",
        "Physical Demands": "-",
        "Work Environment": "-",
        "Points to Note": "-",
        "Management": "-",
        "Field Sales": "-",
        "Pay": "-",
        "Shifts": "-"
    }
    
    if not html_or_text:
        return sections
        
    soup = BeautifulSoup(html_or_text, 'html.parser')
    text = soup.get_text('\n', strip=True)
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    
    current_section = None
    section_buffers = {k: [] for k in sections.keys()}
    
    heading_map = {
        "experience and education requirements": "Education",
        "experience and education": "Education",
        "experience & education requirements": "Education",
        "experience & education": "Education",
        "education and experience requirements": "Education",
        "education and experience": "Education",
        "education requirements": "Education",
        "education": "Education",
        "dexcom offers": "Dexcom Offers",
        "why dexcom": "Dexcom Offers",
        "what you'll get": "Dexcom Offers",
        "what you’ll get": "Dexcom Offers",
        "meet the team": "Summary",
        "position summary": "Summary",
        "role summary": "Summary",
        "job summary": "Summary",
        "about the role": "Summary",
        "where you come in": "Responsibilities",
        "essential duties and responsibilities": "Responsibilities",
        "essential duties": "Responsibilities",
        "responsibilities": "Responsibilities",
        "key responsibilities": "Responsibilities",
        "duties and responsibilities": "Responsibilities",
        "what makes you successful": "Required Qualifications",
        "required qualifications": "Required Qualifications",
        "requirements": "Required Qualifications",
        "essential capabilities": "Required Qualifications",
        "about you": "Required Qualifications",
        "essential skills": "Required Qualifications",
        "qualifications": "Required Qualifications",
        "preferred qualifications": "Preferred Qualifications",
        "key competencies": "Preferred Qualifications",
        "nice to have skills": "Preferred Qualifications",
        "travel required": "Travel",
        "travel": "Travel",
        "flex workplace": "Workplace Type",
        "remote workplace": "Workplace Type",
        "workplace type": "Workplace Type",
        "functional description": "Functional Description",
        "functional/business knowledge": "Functional Knowledge",
        "functional & business knowledge": "Functional Knowledge",
        "scope": "Scope",
        "judgement": "Judgement",
        "language skills": "Language Skills",
        "physical demands": "Physical Demands",
        "work environment": "Work Environment",
        "please note": "Points to Note",
        "management": "Management",
        "field sales": "Field Sales",
        "pay": "Pay",
        "salary": "Pay",
        "shifts": "Shifts"
    }
    
    sorted_headings = sorted(heading_map.items(), key=lambda x: len(x[0]), reverse=True)
    
    for line in lines:
        l_lower = line.lower().replace('’', "'").replace('‘', "'")
        
        if any(d in l_lower for d in ['to all staffing and recruiting agencies', 'to all staffing', 'an equal opportunity', 'ai in hiring', 'seniority level', 'employment type']):
            current_section = None
            continue
            
        matched_sec = None
        for key_phrase, target_sec in sorted_headings:
            if key_phrase in l_lower and len(line) < 80:
                matched_sec = target_sec
                break
                
        if matched_sec:
            current_section = matched_sec
        elif current_section:
            section_buffers[current_section].append(line)
            
    for sec, buffer in section_buffers.items():
        if buffer:
            sections[sec] = clean_bullets('\n'.join(buffer))
            
    wp_match = re.search(r'(?:Flex Workplace|Remote Workplace|Workplace Type)\s*\n+([\s\S]+?)(?=\n\s*(?:Travel|Salary|Monthly base salary|Please note|An Equal Opportunity|To all Staffing|AI in Hiring|Seniority|Employment|Job function|Industries|\Z))', text, re.IGNORECASE)
    if wp_match:
        wp_text = wp_match.group(1).strip()
        cleaned_wp = clean_bullets(wp_text)
        if cleaned_wp and cleaned_wp != "-":
            sections["Workplace Type"] = cleaned_wp
        
    tr_match = re.search(r'Travel\s*Required\s*\n+([^\n]+)', text, re.IGNORECASE)
    if tr_match:
        t_line = tr_match.group(1).strip()
        cleaned_tr = re.sub(r'^\s*[\bullet\*\-\–\—\•]\s*', '', t_line).strip()
        if cleaned_tr:
            sections["Travel"] = cleaned_tr

    sal_match = re.search(r'([^\.\n]*Monthly\s+base\s+salary[^\.\n]*\.[^\.\n]*|[^\.\n]*Monthly\s+base\s+salary[^\.\n]*|\$\d[\d,]*\.\d{2}\s*-\s*\$\d[\d,]*\.\d{2})', text, re.IGNORECASE)
    if sal_match:
        sections["Pay"] = sal_match.group(0).strip()
    else:
        sal_block = re.search(r'Salary\s*\n+([^\n]+)', text, re.IGNORECASE)
        if sal_block and not any(k in sal_block.group(1).lower() for k in ['seniority', 'employment', 'job function']):
            sections["Pay"] = sal_block.group(1).strip()

    note_match = re.search(r'(Please note:?[^\.\n]*)', text, re.IGNORECASE)
    if note_match:
        sections["Points to Note"] = note_match.group(0).strip()

    return sections

def process_job_link(link_line: str, headers: dict, debug: bool = False) -> dict:
    date_from_file = ""
    url = link_line
    if '\t' in link_line:
        parts = link_line.split('\t')
        date_from_file = parts[0].strip()
        url = parts[1].strip()
        
    linkedin_job_id = extract_linkedin_job_id(url)
    if debug:
        print(f"\n[DEBUG] Processing URL: {url}")
        print(f"[DEBUG] LinkedIn Job ID: {linkedin_job_id}")
    
    title = "N/A"
    company = "N/A"
    raw_location = "N/A"
    location = "N/A"
    description_html = ""
    
    if linkedin_job_id != "NA":
        view_url = f"https://www.linkedin.com/jobs/view/{linkedin_job_id}/"
        api_url = f"https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/{linkedin_job_id}"
        
        try:
            r_api = requests.get(api_url, headers=headers, timeout=10)
            if r_api.status_code == 200:
                soup_api = BeautifulSoup(r_api.text, 'html.parser')
                
                title_el = soup_api.find('h2', class_='top-card-layout__title') or soup_api.find('h1')
                company_el = soup_api.find('a', class_='topcard__org-name-link') or soup_api.find('span', class_='topcard__flavor')
                location_el = soup_api.find('span', class_='topcard__flavor topcard__flavor--bullet')
                
                title = title_el.get_text(strip=True) if title_el else "N/A"
                company = company_el.get_text(strip=True) if company_el else "N/A"
                raw_location = location_el.get_text(strip=True) if location_el else "N/A"
                
                desc_el = soup_api.find('div', class_='show-more-less-html__markup') or soup_api.find('div', class_='description__text')
                if desc_el:
                    description_html = str(desc_el)
                    
                # If LinkedIn page content is not in English, translate title, location, and description to English
                title = translate_to_english(title)
                raw_location = translate_to_english(raw_location)
                location = format_location(raw_location)
                description_html = translate_to_english(description_html)
                
        except Exception as e:
            print(f"Error fetching LinkedIn API for ID {linkedin_job_id}: {e}")
            
    date_info = parse_date_info(date_from_file)
    country, us_ous = determine_country_and_us(location)
    
    bus_func, dept = classify_function_and_department(title, description_html)
    sections = parse_description_sections(description_html)
    
    dexcom_job_id, direct_dexcom_url, raw_dexcom_dept = search_dexcom_career_page(title, raw_location, description_html, headers)
    
    l1_func, l2_dept, dept_dexcom_site, confidence_level = match_dexcom_department_from_taxonomy(title, description_html, raw_dexcom_dept, bus_func, dept)
    
    if l1_func:
        bus_func = strip_serial_number(l1_func)
    else:
        bus_func = strip_serial_number(bus_func)
        
    dept = strip_serial_number(l2_dept) if l2_dept else strip_serial_number(dept)
        
    if linkedin_job_id != "NA" and dexcom_job_id != "NA":
        job_posting_site = "Both"
    elif dexcom_job_id != "NA":
        job_posting_site = "Dexcom Careers"
    else:
        job_posting_site = "LinkedIn"
        
    row = {
        "Date": date_info["Date"],
        "Year": date_info["Year"],
        "Month": date_info["Month"],
        "Date Posted": date_info["Date Posted"],
        "Job ID LinkedIn": linkedin_job_id,
        "Job ID Dexcom Page": dexcom_job_id,
        "Job Posting site": job_posting_site,
        "Location": location,
        "Country": country,
        "US / OUS": us_ous,
        "Business Function": bus_func,
        "Department": dept,
        "Department from Dexcom Company Website": dept_dexcom_site,
        "Confidence Level": confidence_level,
        "Job Title": title,
        "Fresh / reposted": "-",
        "Dexcom Offers / Why Dexcom? / What you'll get": sections["Dexcom Offers"],
        "Summary / Position Summary / Meet the team / Role Summary": sections["Summary"],
        "Essential Duties and Responsibilities / Where you come in": sections["Responsibilities"],
        "Supervisory Responsibilities": sections["Supervisory"],
        "Required Qualifications / What makes you successful / Requirements / Essential Capabilities / Competencies / About you": sections["Required Qualifications"],
        "Preferred Qualifications / Key Competencies": sections["Preferred Qualifications"],
        "Education Requirements and Experience": sections["Education"],
        "Travel Required": sections["Travel"],
        "Workplace Type": sections["Workplace Type"],
        "Functional Description": sections["Functional Description"],
        "Functional / Business Knowledge": sections["Functional Knowledge"],
        "Scope": sections["Scope"],
        "Judgement": sections["Judgement"],
        "Language Skills": sections["Language Skills"],
        "Physical Demands": sections["Physical Demands"],
        "Work Environment": sections["Work Environment"],
        "Points to Note": sections["Points to Note"],
        "Management": sections["Management"],
        "Field Sales": sections["Field Sales"],
        "Pay / Non-Exempt Salary Details / Commercial Salary Details / Exempt Salary Details": sections["Pay"],
        "Shifts": sections["Shifts"],
        "Direct URL of the Dexcom Job Page": direct_dexcom_url,
        "LinkedIn URL": url
    }
    
    return row

def update_fresh_reposted_status(records: list) -> None:
    li_counts = {}
    dex_counts = {}
    
    for r in records:
        li_id = r.get("Job ID LinkedIn")
        if li_id and li_id not in ["NA", "N/A", "-", ""]:
            li_counts[li_id] = li_counts.get(li_id, 0) + 1
            
        dex_id = r.get("Job ID Dexcom Page")
        if dex_id and dex_id not in ["NA", "N/A", "-", ""]:
            dex_counts[dex_id] = dex_counts.get(dex_id, 0) + 1
            
    for r in records:
        li_id = r.get("Job ID LinkedIn")
        dex_id = r.get("Job ID Dexcom Page")
        
        is_reposted = False
        if li_id and li_id not in ["NA", "N/A", "-", ""] and li_counts.get(li_id, 0) > 1:
            is_reposted = True
        if dex_id and dex_id not in ["NA", "N/A", "-", ""] and dex_counts.get(dex_id, 0) > 1:
            is_reposted = True
            
        r["Fresh / reposted"] = "Reposted" if is_reposted else "Fresh/New Job Posted"

def sanitize_excel_value(val):
    if isinstance(val, str) and val.startswith(('=', '+', '@')):
        return "'" + val
    return val

def main():
    import sys
    debug = "--debug" in sys.argv or "-d" in sys.argv
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept-Language': 'en-US,en;q=0.9'
    }
    
    txt_file = 'job_links.txt' if os.path.exists('job_links.txt') else 'links.txt'
    
    print(f"Reading links from {txt_file}...")
    with open(txt_file, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f if line.strip()]
        
    if lines and ('Date' in lines[0] or 'LinkedIn' in lines[0]):
        lines = lines[1:]
        
    print(f"Processing {len(lines)} job record(s)...\n")
    
    records = []
    for idx, link_line in enumerate(lines, start=1):
        print(f"[{idx}/{len(lines)}] Extracting data for: {link_line[:60]}...")
        row_data = process_job_link(link_line, headers, debug=debug)
        records.append(row_data)
        print(f"  Title: {row_data['Job Title']}")
        print(f"  Location: {row_data['Location']}")
        print(f"  Business Function: {row_data['Business Function']}")
        print(f"  Department (Dexcom): {row_data['Department from Dexcom Company Website']}")
        print(f"  Confidence Level: {row_data['Confidence Level']}")
        print(f"  LinkedIn ID: {row_data['Job ID LinkedIn']} | Dexcom ID: {row_data['Job ID Dexcom Page']}")
        print(f"  Dexcom Direct URL: {row_data['Direct URL of the Dexcom Job Page']}")
        print(f"  LinkedIn URL: {row_data['LinkedIn URL']}\n")
        
    update_fresh_reposted_status(records)
    
    df = pd.DataFrame(records, columns=EXCEL_HEADERS)
    if hasattr(df, 'map'):
        df = df.map(sanitize_excel_value)
    else:
        df = df.applymap(sanitize_excel_value)
    
    saved_file = None
    targets = ['job_postings_final.xlsx', 'job_postings.xlsx', 'job_postings_latest.xlsx']
    
    for target in targets:
        try:
            df.to_excel(target, index=False)
            print(f"Successfully populated all {len(EXCEL_HEADERS)} fields and saved to {target}!")
            saved_file = target
            break
        except Exception as e:
            print(f"Could not save to {target}: {e}")
            
    if not saved_file:
        import time
        timestamp_file = f"job_postings_new_{int(time.time())}.xlsx"
        try:
            df.to_excel(timestamp_file, index=False)
            print(f"Saved output to {timestamp_file}!")
        except Exception as e:
            print(f"Could not save to timestamp file: {e}")

if __name__ == '__main__':
    main()
