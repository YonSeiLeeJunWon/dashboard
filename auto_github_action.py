import os
import re
import time
from datetime import datetime, timedelta
import pandas as pd
import requests
from bs4 import BeautifulSoup
from pathlib import Path
import json

# ============================================================
# 0. 경로 설정 (GitHub Actions 서버 전용 상대 경로)
# ============================================================
BASE_DIR = Path(__file__).resolve().parent
DATA_IN_DIR = BASE_DIR / "DATA" / "IN"
DATA_OUT_DIR = BASE_DIR / "DATA" / "OUT"

os.makedirs(DATA_IN_DIR, exist_ok=True)
os.makedirs(DATA_OUT_DIR, exist_ok=True)

NEWS_FILE_PATH = DATA_IN_DIR / "보험뉴스_Data.xlsx"
INDEX_HTML_PATH = BASE_DIR / "index.html"
OUT_DASHBOARD_PATH = DATA_OUT_DIR / "보험회사_종합분석_대시보드.html"

CATEGORIES = [
    "보험공통(IFRS17, K-ICS 등)",
    "생명보험(산업, 개별 생명보험회사)",
    "손해보험(산업, 개별 손해보험회사)",
    "금융위원회, 금융감독원, 보험연구원, 보험개발원",
    "M&A, 실손보험, 보험사기, 판매채널(GA)",
    "금융사고(책무구조도, 자금세탁방지), 부실채권(NPL), 보이스피싱",
    "AI(보험, 업스테이지)",
    "증시",
    "통계, 물가, 부동산, 기타"
]

CAT_PATTERNS = [
    (0, r"1\.\s*보험공통"), (1, r"2\.\s*생명보험"), (2, r"3\.\s*손해보험"),
    (3, r"4\.\s*금융위"),   (4, r"5\.\s*M&A"),      (5, r"6\.\s*금융사고"),
    (6, r"7\.\s*AI"),      (7, r"8\.\s*증시"),     (8, r"9\.\s*통계")
]

EXCLUDE_KEYWORDS = [
    "오늘의 날씨", "오늘의날씨", "한국은행 지표", "한국은행지표", 
    "광고문의", "이전글", "이전 글", "다음글", "다음 글", "목록으로", 
    "URL 복사", "좋아요", "댓글", "답글"
]

PRESS_NAMES = ["경북도민일보", "아주경제", "매일경제", "한국경제", "서울경제", "조선일보", "중앙일보", "동아일보", "연합뉴스", "뉴시스", "뉴스1", "이투데이", "파이낸셜뉴스"]
CLUB_ID = "31337941"
MENU_ID = "1"
BASE_URL = f"https://cafe.naver.com/ArticleList.nhn?search.clubid={CLUB_ID}&search.menuid={MENU_ID}&search.boardtype=L"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

def clean_sheet_name(name):
    return name.strip()[:31]

def extract_valid_date(text):
    if not text:
        return ""
    cleaned_text = re.sub(r"(20[2-9]\d)0\.", r"\1.", text)
    match = re.search(r"(20[2-9]\d)\.(\d{2})\.(\d{2})", cleaned_text)
    if match:
        year, month, day = match.groups()
        if 1 <= int(month) <= 12 and 1 <= int(day) <= 31:
            return f"{year}.{month}.{day}"
    return ""

# ============================================================
# 1. 크롤링 로직 (HTTP Request 기반)
# ============================================================
def parse_article_content(soup, fallback_date=""):
    article_data = {cat: [] for cat in CATEGORIES}
    title_elem = soup.select_one(".title_text, .title_area, .ArticleTitle .title, h3.title_text")
    article_date = extract_valid_date(title_elem.get_text(strip=True)) if title_elem else ""
            
    if not article_date:
        date_elem = soup.select_one(".date_text, .article_writer .date, .date, .ArticleWriter .date, .m-t-13 .date")
        if date_elem:
            article_date = extract_valid_date(date_elem.text)

    if not article_date:
        article_date = fallback_date

    main_content = soup.select_one(".article_viewer, .se-main-container, #tbody, .ContentRenderer, .ArticleContentBox") or soup
    current_category_idx = 0

    for elem in main_content.find_all(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'li', 'span', 'a']):
        elem_text = elem.get_text(strip=True)
        if elem_text and len(elem_text) < 80:
            for cat_idx, pattern in CAT_PATTERNS:
                if re.search(pattern, elem_text):
                    current_category_idx = cat_idx
                    break

        if elem.name == 'a' and elem.has_attr('href'):
            news_url = elem['href']
            if not news_url.startswith("http") or "cafe.naver.com" in news_url or any(ex_domain in news_url for ex_domain in ["hanmail.net", "kcisc.kr"]):
                continue

            raw_title = elem.get_text(strip=True)
            news_title = raw_title
            if len(raw_title) < 12 or any(press in raw_title for press in PRESS_NAMES):
                if elem.parent:
                    parent_text = elem.parent.get_text(strip=True)
                    if len(parent_text) > len(raw_title):
                        news_title = parent_text
            
            if any(ex_kw in news_title for ex_kw in EXCLUDE_KEYWORDS):
                continue

            if len(news_title) > 2:
                target_cat = CATEGORIES[current_category_idx]
                already_exists = any(item['기사 주소'] == news_url for c in CATEGORIES for item in article_data[c])
                if not already_exists:
                    article_data[target_cat].append({
                        "기사 제목": news_title,
                        "연월일": article_date,
                        "기사 주소": news_url
                    })

    return article_date, article_data

def crawl_naver_cafe_news_auto():
    now = datetime.now()
    start_date_str = (now - timedelta(days=3)).strftime("%Y.%m.%d")
    end_date_str = now.strftime("%Y.%m.%d")
    
    start_date = datetime.strptime(start_date_str, "%Y.%m.%d")
    end_date = datetime.strptime(end_date_str, "%Y.%m.%d")
    
    print(f"\n[1/3] 클라우드 서버에서 뉴스 수집 시작: {start_date_str} ~ {end_date_str}")
    
    session = requests.Session()
    session.headers.update(HEADERS)
    
    # Secrets 환경변수에 등록된 쿠키가 있다면 적용
    naver_cookie = os.getenv("NAVER_COOKIE")
    if naver_cookie:
        session.headers.update({"Cookie": naver_cookie})

    all_collected_data = {cat: [] for cat in CATEGORIES}
    stop_crawling = False
    page = 1

    while not stop_crawling:
        target_page_url = f"{BASE_URL}&search.page={page}"
        resp = session.get(target_page_url)
        soup = BeautifulSoup(resp.text, "html.parser")
        
        # 카페 메인 프레임 URL 추출
        iframe = soup.select_one("#cafe_main")
        if iframe and iframe.has_attr("src"):
            iframe_url = "https://cafe.naver.com" + iframe["src"]
            resp = session.get(iframe_url)
            soup = BeautifulSoup(resp.text, "html.parser")

        rows = soup.select("#main-area .article-board tr, #main-area .board-list tr, .article-board tbody tr, .article-board tr")
        if not rows:
            break

        article_target_list = []
        for row in rows:
            notice_check = row.select_one(".ico_notice, .td_notice, .notice, span.type, .ico_type")
            row_class = " ".join(row.get("class", []))
            if notice_check or "notice" in row_class or "recommend" in row_class:
                continue

            num_elem = row.select_one("td.td_article, .inner_number, .td_num")
            if num_elem and not num_elem.text.strip().isdigit():
                continue

            title_elem = row.select_one("a.article, a.a-title")
            date_elem = row.select_one("td.td_date, .td_date, .date, .time")
            if not title_elem:
                continue
            
            title_text = title_elem.text.strip()
            date_text = date_elem.text.strip() if date_elem else ""
            
            extracted_date_str = extract_valid_date(title_text)
            if not extracted_date_str:
                if ":" in date_text:
                    extracted_date_str = datetime.now().strftime("%Y.%m.%d")
                elif date_text:
                    extracted_date_str = extract_valid_date(date_text)

            if extracted_date_str:
                try:
                    row_date = datetime.strptime(extracted_date_str, "%Y.%m.%d")
                    if row_date > end_date:
                        continue
                    elif row_date < start_date:
                        stop_crawling = True
                        break
                except ValueError:
                    pass

            href = title_elem.get('href')
            if href:
                article_url = f"https://cafe.naver.com{href}" if href.startswith("/") else href
                article_target_list.append((article_url, extracted_date_str))

        for art_url, art_date_str in article_target_list:
            art_resp = session.get(art_url)
            sub_soup = BeautifulSoup(art_resp.text, "html.parser")
            
            sub_iframe = sub_soup.select_one("#cafe_main")
            if sub_iframe and sub_iframe.has_attr("src"):
                art_resp = session.get("https://cafe.naver.com" + sub_iframe["src"])
                sub_soup = BeautifulSoup(art_resp.text, "html.parser")

            parsed_date, art_data = parse_article_content(sub_soup, fallback_date=art_date_str)
            final_date = parsed_date if parsed_date else art_date_str

            for cat in CATEGORIES:
                if art_data[cat]:
                    all_collected_data[cat].extend(art_data[cat])

        if stop_crawling:
            break
        page += 1

    save_or_update_excel(all_collected_data)

def save_or_update_excel(new_data_dict):
    def write_to_excel(writer, data_dict, existing_file_path=None):
        for cat in CATEGORIES:
            sheet_title = clean_sheet_name(cat)
            new_df = pd.DataFrame(data_dict[cat])
            if existing_file_path and os.path.exists(existing_file_path):
                try:
                    existing_df = pd.read_excel(existing_file_path, sheet_name=sheet_title)
                    combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                except Exception:
                    combined_df = new_df
            else:
                combined_df = new_df

            if not combined_df.empty and "기사 주소" in combined_df.columns:
                combined_df.drop_duplicates(subset=["기사 주소"], keep="first", inplace=True)
                if "연월일" in combined_df.columns:
                    combined_df.sort_values(by="연월일", ascending=False, inplace=True)
            combined_df.to_excel(writer, sheet_name=sheet_title, index=False)

    if os.path.exists(NEWS_FILE_PATH):
        with pd.ExcelWriter(NEWS_FILE_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            write_to_excel(writer, new_data_dict, NEWS_FILE_PATH)
    else:
        with pd.ExcelWriter(NEWS_FILE_PATH, engine='openpyxl') as writer:
            write_to_excel(writer, new_data_dict)
    print("✓ 엑셀 업데이트 완료.")

# ============================================================
# 2. 대시보드 빌드 및 index.html 동기화
# ============================================================
def generate_dashboard_and_sync():
    print("\n[2/3] 대시보드 HTML 파일 생성 중...")
    
    # auto_dashboard_module.py 모듈 호출
    from auto_dashboard_module import build_dashboard
    output_html_path = build_dashboard(data_dir=DATA_IN_DIR)
    
    # 대시보드 HTML 생성물을 root의 index.html로 복사
    with open(output_html_path, "r", encoding="utf-8") as f_src:
        html_content = f_src.read()

    with open(INDEX_HTML_PATH, "w", encoding="utf-8") as f_dst:
        f_dst.write(html_content)

    print("✓ index.html 및 대시보드 HTML 파일 생성/동기화 완료!")

if __name__ == "__main__":
    crawl_naver_cafe_news_auto()
    generate_dashboard_and_sync()
