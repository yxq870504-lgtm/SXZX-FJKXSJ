from __future__ import annotations

from pathlib import Path
from datetime import datetime
import json
import re
import sys
import urllib.parse
import urllib.request

PROJECT = Path(__file__).resolve().parent
CONFIG = PROJECT / 'official_calendar_sources.json'
OUT_DIR = PROJECT / 'official_calendar_cache'
OUT_FILE = OUT_DIR / 'candidates.json'

DATE_RE = re.compile(r'(?:(2027)年)?\s*(1|2)月\s*(\d{1,2})日')


def http_get(url: str, timeout: int = 15) -> str:
    req = urllib.request.Request(url, headers={
        'User-Agent': 'Mozilla/5.0 official-calendar-refresh/1.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    })
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = resp.read(900_000)
        ctype = resp.headers.get_content_charset() or 'utf-8'
        try:
            return data.decode(ctype, errors='ignore')
        except LookupError:
            return data.decode('utf-8', errors='ignore')


def strip_html(html: str) -> str:
    html = re.sub(r'<script[\s\S]*?</script>', ' ', html, flags=re.I)
    html = re.sub(r'<style[\s\S]*?</style>', ' ', html, flags=re.I)
    text = re.sub(r'<[^>]+>', ' ', html)
    text = re.sub(r'&nbsp;|&#160;', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text


def search_bing(query: str, limit: int = 8):
    url = 'https://www.bing.com/search?q=' + urllib.parse.quote(query)
    try:
        html = http_get(url)
    except Exception:
        return []
    links = []
    patterns = [
        r'<a\s+href="(https?://[^"]+)"',
        r'<h2>\s*<a\s+href="(https?://[^"]+)"',
    ]
    for pat in patterns:
        for m in re.finditer(pat, html, flags=re.I):
            href = m.group(1)
            href = href.replace('&amp;', '&')
            if 'bing.com' in href or 'microsoft.com' in href:
                continue
            if href not in links:
                links.append(href)
            if len(links) >= limit:
                return links
    return links


def parse_candidate(province: str, url: str, text: str):
    if '2027' not in text and '2026-2027' not in text and '2026—2027' not in text and '2026－2027' not in text:
        return None
    if not any(k in text for k in ['寒假', '开学', '校历', '放假', '教学日历', '学年安排']):
        return None

    windows = []
    for m in DATE_RE.finditer(text):
        start = max(0, m.start() - 80)
        end = min(len(text), m.end() + 80)
        ctx = text[start:end]
        month = int(m.group(2)); day = int(m.group(3))
        try:
            date = f'2027-{month:02d}-{day:02d}'
        except Exception:
            continue
        windows.append((date, ctx))

    holiday = None
    spring = None
    for date, ctx in windows:
        if date.startswith('2027-01') and any(k in ctx for k in ['寒假', '放假', '假期开始', '起放寒假']):
            holiday = holiday or date
        if date.startswith('2027-02') and any(k in ctx for k in ['开学', '上课', '报到', '注册']):
            spring = spring or date

    if not holiday and not spring:
        return None

    confidence = 0.65
    netloc = urllib.parse.urlparse(url).netloc
    if '.gov.cn' in url:
        confidence += 0.2
    if any(k in text[:5000] for k in ['教育', '教委', '教育局', '教育厅', '学校', '中学']):
        confidence += 0.05
    if any(k in netloc for k in ['edu', 'jy', 'jyt', 'jw']):
        confidence += 0.04
    if holiday and spring:
        confidence += 0.1
    confidence = min(confidence, 0.98)

    return {
        'province': province,
        'holiday_date': holiday,
        'spring_date': spring,
        'source_url': url,
        'source_title': text[:80],
        'confidence': round(confidence, 2),
        'found_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }


def main():
    cfg = json.loads(CONFIG.read_text(encoding='utf-8'))
    candidates = []
    for item in cfg['provinces']:
        province = item['province']
        domains = item.get('domains', [])
        short = item.get('short') or province
        city_keywords = item.get('cities', [])
        queries = []
        for domain in domains:
            queries.extend([
                f"site:{domain} {province} 2026-2027学年度 校历 寒假 开学",
                f"site:{domain} {short} 2027 寒假 开学时间 中小学",
                f"site:{domain} {short} 2026-2027学年 普通高中 义务教育 寒假",
            ])
        for city in city_keywords:
            queries.extend([
                f"{city} 2026-2027学年度 中小学 校历 寒假 开学 官方",
                f"site:gov.cn {city} 2027 寒假 开学 中小学",
                f"site:edu.cn {city} 2026-2027 校历 寒假 开学",
            ])
        queries.extend([
            f"{province} 2026-2027学年度 中小学 校历 寒假 开学 官方",
            f"{short} 2027 寒假 开学时间 普通高中 义务教育 官方",
            f"{short} 市教育局 2026-2027 校历 寒假 开学",
            f"{short} 学校 2026-2027 校历 寒假 开学",
        ])
        seen_urls = set()
        for q in queries[:14]:
            for url in search_bing(q, limit=8):
                if url in seen_urls:
                    continue
                seen_urls.add(url)
                netloc = urllib.parse.urlparse(url).netloc
                allowed = any(d in netloc for d in domains)
                official_like = any(x in netloc for x in ['.gov.cn', '.edu.cn', 'edu.', 'jy.', 'jyt.', 'jw.'])
                if not (allowed or official_like):
                    continue
                try:
                    html = http_get(url)
                    text = strip_html(html)
                    cand = parse_candidate(province, url, text)
                    if cand:
                        candidates.append(cand)
                except Exception:
                    continue
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    payload = {'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'candidates': candidates}
    OUT_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'wrote {OUT_FILE} candidates={len(candidates)}')


if __name__ == '__main__':
    main()
