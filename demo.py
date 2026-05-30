import re
import time
import json
import csv
import random
import requests
from bs4 import BeautifulSoup
from dataclasses import dataclass, asdict
from urllib.parse import urljoin
from typing import List, Optional

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "en-CA,en;q=0.9",
}

TIME_RE = re.compile(r"\b(1[0-2]|0?[1-9]):[0-5]\d\s?(AM|PM)\b", re.I)
DATE_RE = re.compile(r"\b(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)(?:day)?\b.*?\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b", re.I)

@dataclass
class ShowTime:
    source: str
    theatre: str
    address: str
    movie: str
    date: str
    time: str
    format: str
    show_url: str

class FetchError(Exception):
    pass

def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()

def session_with_retry():
    s = requests.Session()
    s.headers.update(HEADERS)
    return s

def fetch_html(url: str, timeout: int = 30, retries: int = 3, sleep_range=(0.8, 1.8)) -> str:
    last = None
    s = session_with_retry()
    for attempt in range(retries):
        try:
            r = s.get(url, timeout=timeout)
            r.raise_for_status()
            if "maintenance" in r.text.lower() or "sorry for the inconvenience" in r.text.lower():
                return r.text
            return r.text
        except Exception as e:
            last = e
            time.sleep(random.uniform(*sleep_range))
    raise FetchError(str(last))

def fetch_rendered_html(url: str):
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(
            user_agent=HEADERS["User-Agent"],
            locale="en-CA",
            viewport={"width": 1440, "height": 1200}
        )
        page.goto(url, wait_until="networkidle", timeout=60000)
        html = page.content()
        browser.close()
        return html

def likely_js_rendered(html: str) -> bool:
    t = html.lower()
    return ("__next" in t) or ("hydration" in t) or ("application/ld+json" not in t and "showtime" not in t and len(t) < 10000)

def extract_theatre_and_address(soup: BeautifulSoup):
    title = clean(soup.find("h1").get_text(" ", strip=True)) if soup.find("h1") else ""
    address = ""
    text = clean(soup.get_text(" ", strip=True))
    m = re.search(r"(\d+[^|]+(?:ON|BC|AB|MB|SK|QC|NS|NB|NL|PE)\s+[A-Z0-9 ]+)", text)
    if m:
        address = clean(m.group(1))
    return title, address

def parse_time_nodes(soup: BeautifulSoup):
    hits = []
    for node in soup.find_all(string=TIME_RE):
        txt = clean(str(node))
        m = TIME_RE.search(txt)
        if m:
            hits.append((m.group(0), node.parent))
    return hits

def parse_landmark_html(html: str, url: str) -> List[ShowTime]:
    soup = BeautifulSoup(html, "html.parser")
    theatre, address = extract_theatre_and_address(soup)
    if "sorry for the inconvenience" in soup.get_text(" ", strip=True).lower():
        return []

    results = []
    for t, parent in parse_time_nodes(soup):
        context = clean(parent.get_text(" ", strip=True))
        movie = ""
        fmt = ""
        prev = parent.find_previous(["h2", "h3", "h4", "strong", "div", "span"])
        if prev:
            ptxt = clean(prev.get_text(" ", strip=True))
            if len(ptxt) > 2:
                movie = ptxt
        for token in ["IMAX", "3D", "Laser Ultra", "Premiere", "VIP", "Recliner", "Reserved"]:
            if token.lower() in context.lower():
                fmt = token
                break
        results.append(ShowTime("landmark", theatre, address, movie, "", t, fmt, url))
    return dedupe(results)

def parse_cineplex_html(html: str, url: str) -> List[ShowTime]:
    soup = BeautifulSoup(html, "html.parser")
    theatre, address = extract_theatre_and_address(soup)

    results = []
    for node in soup.find_all(string=TIME_RE):
        m = TIME_RE.search(clean(str(node)))
        if not m:
            continue
        parent = node.parent
        context = clean(parent.get_text(" ", strip=True))
        movie = ""
        fmt = ""
        prev = parent.find_previous(["h2", "h3", "h4", "div", "span", "strong"])
        if prev:
            movie = clean(prev.get_text(" ", strip=True))
        for token in ["IMAX", "UltraAVX", "VIP", "3D", "D-BOX", "Laser", "Closed Caption", "Descriptive Video"]:
            if token.lower() in context.lower():
                fmt = token
                break
        results.append(ShowTime("cineplex", theatre, address, movie, "", m.group(0), fmt, url))
    return dedupe(results)

def dedupe(rows: List[ShowTime]) -> List[ShowTime]:
    seen = set()
    out = []
    for r in rows:
        key = (r.source, r.theatre, r.movie, r.date, r.time, r.format, r.show_url)
        if key not in seen:
            seen.add(key)
            out.append(r)
    return out

def scrape(source: str, url: str) -> List[ShowTime]:
    html = fetch_html(url)
    if likely_js_rendered(html):
        try:
            html = fetch_rendered_html(url)
        except Exception:
            pass
    if source == "landmark":
        return parse_landmark_html(html, url)
    if source == "cineplex":
        return parse_cineplex_html(html, url)
    raise ValueError("source must be 'landmark' or 'cineplex'")

urls = [
    ("landmark", "https://www.landmarkcinemas.com/showtimes/waterloo"),
    ("cineplex", "https://www.cineplex.com/?openTM=true")
]

all_rows = []
for source, url in urls:
    try:
        all_rows.extend(scrape(source, url))
    except Exception as e:
        all_rows.append(ShowTime(source, "", "", "", "", "", f"ERROR: {e}", url))

with open("showtimes.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=ShowTime.__annotations__.keys())
    w.writeheader()
    w.writerows(asdict(r) for r in all_rows)