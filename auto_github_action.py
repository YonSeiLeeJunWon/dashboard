import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
import pandas as pd
import requests
from bs4 import BeautifulSoup
from pathlib import Path
import json

# ============================================================
# 0. 경로 및 시간 설정 (KST 한국 표준시 적용)
# ============================================================
BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

DATA_IN_DIR = BASE_DIR / "DATA" / "IN"
DATA_OUT_DIR = BASE_DIR / "DATA" / "OUT"

os.makedirs(DATA_IN_DIR, exist_ok=True)
os.makedirs(DATA_OUT_DIR, exist_ok=True)

NEWS_FILE_PATH = DATA_IN_DIR / "보험뉴스_Data.xlsx"
INDEX_HTML_PATH = BASE_DIR / "index.html"
OUT_DASHBOARD_PATH = DATA_OUT_DIR / "보험회사_종합분석_대시보드.html"

KST = timezone(timedelta(hours=9))

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

# 클라우드 IP 차단 우회용 네이버 브라우저 완벽 모사 헤더
HEADERS_MOBILE_FULL = {
    "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 17_4_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.3.1 Mobile/15E148 Safari/604.1",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept-Encoding": "gzip, deflate, br",
    "Origin": "https://m.cafe.naver.com",
    "Referer": f"https://m.cafe.naver.com/ca-fe/web/cafes/{CLUB_ID}/menus/{MENU_ID}",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin"
}

HEADERS_PC_FULL = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": f"https://cafe.naver.com/ArticleList.nhn?search.clubid={CLUB_ID}&search.menuid={MENU_ID}&search.boardtype=L",
    "Sec-Ch-Ua": '"Google Chrome";v="123", "Not:A-Brand";v="8", "Chromium";v="123"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "same-origin"
}

def clean_sheet_name(name):
    return name.strip()[:31]

def extract_valid_date(text):
    if not text:
        return ""
    cleaned_text = re.sub(r"(20[2-9]\d)0\.", r"\1.", text)
    match = re.search(r"(20[2-9]\d)\.(\d{1,2})\.(\d{1,2})", cleaned_text)
    if match:
        year, month, day = match.groups()
        m_int, d_int = int(month), int(day)
        if 1 <= m_int <= 12 and 1 <= d_int <= 31:
            return f"{year}.{m_int:02d}.{d_int:02d}"
    return ""

def get_latest_collected_date():
    if not os.path.exists(NEWS_FILE_PATH):
        return None
    
    latest_date = None
    try:
        xls = pd.ExcelFile(NEWS_FILE_PATH)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            if not df.empty and "연월일" in df.columns:
                valid_dates = df["연월일"].dropna().astype(str)
                for d_str in valid_dates:
                    try:
                        parsed_dt = datetime.strptime(d_str.strip(), "%Y.%m.%d")
                        if latest_date is None or parsed_dt > latest_date:
                            latest_date = parsed_dt
                    except ValueError:
                        continue
    except Exception as e:
        print(f"⚠️ 기존 엑셀 날짜 읽기 중 경고: {e}")
    
    return latest_date

def parse_article_content(soup, fallback_date=""):
    article_data = {cat: [] for cat in CATEGORIES}
    title_elem = soup.select_one(".title_text, .title_area, .ArticleTitle .title, h3.title_text, .tit_box")
    article_date = extract_valid_date(title_elem.get_text(strip=True)) if title_elem else ""
            
    if not article_date:
        date_elem = soup.select_one(".date_text, .article_writer .date, .date, .ArticleWriter .date, .m-t-13 .date, .date_time")
        if date_elem:
            article_date = extract_valid_date(date_elem.text)

    if not article_date:
        article_date = fallback_date

    main_content = soup.select_one(".article_viewer, .se-main-container, #tbody, .ContentRenderer, .ArticleContentBox, .post_area") or soup
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
                        "기사 주소": news_url,
                        "메모": ""
                    })

    return article_date, article_data

def crawl_naver_cafe_news_auto():
    now_kst = datetime.now(KST).replace(tzinfo=None)
    latest_dt = get_latest_collected_date()
    
    if latest_dt:
        start_date = latest_dt - timedelta(days=1)
        print(f"📌 기존 엑셀 최신 수집일 감지: {latest_dt.strftime('%Y.%m.%d')} -> 수집 시작일: {start_date.strftime('%Y.%m.%d')}")
    else:
        start_date = now_kst - timedelta(days=3)
        print(f"📌 기존 수집 데이터 없음 -> 기본 수집 시작일: {start_date.strftime('%Y.%m.%d')}")

    naver_cookie = os.getenv("NAVER_COOKIE")
    
    # [시도 1] 모바일 REST API 방식
    print("\n🔍 [1차 시도] 모바일 REST API 우회 수집 중...")
    session_m = requests.Session()
    session_m.headers.update(HEADERS_MOBILE_FULL)
    if naver_cookie:
        session_m.headers.update({"Cookie": naver_cookie})

    api_url = f"https://m.cafe.naver.com/api/cafe/{CLUB_ID}/menu/{MENU_ID}/articles?page=1"
    try:
        api_resp = session_m.get(api_url, timeout=10)
        if api_resp.status_code == 200:
            res_json = api_resp.json()
            article_list = res_json.get("message", {}).get("result", {}).get("articleList", [])
            if article_list:
                print(f"✅ [1차 성공] 모바일 API로 게시글 {len(article_list)}개가 발견되었습니다.")
                return crawl_via_mobile_api(session_m, start_date=start_date, end_date=now_kst)
    except Exception as e:
        print(f"⚠️ 1차 모바일 API 접속 실패: {e}")

    # [시도 2] 모바일 Ajax Endpoint 방식
    print("🔄 [2차 시도] 모바일 Ajax 목록 수집 중...")
    ajax_url = f"https://m.cafe.naver.com/ArticleListAjax.nhn?search.clubid={CLUB_ID}&search.menuid={MENU_ID}&search.page=1"
    try:
        ajax_resp = session_m.get(ajax_url, timeout=10)
        if ajax_resp.status_code == 200:
            ajax_soup = BeautifulSoup(ajax_resp.text, "html.parser")
            items = ajax_soup.select("li, .board_box, a.article")
            if items:
                print(f"✅ [2차 성공] Ajax 엔드포인트로 {len(items)}개 게시글을 확인했습니다.")
                return crawl_via_mobile_ajax(session_m, start_date=start_date, end_date=now_kst)
    except Exception as e:
        print(f"⚠️ 2차 Ajax 접속 실패: {e}")

    # [시도 3] PC Cafe Main Frame 파싱 방식
    print("🔄 [3차 시도] PC 버전 우회 파싱 수집 중...")
    session_pc = requests.Session()
    session_pc.headers.update(HEADERS_PC_FULL)
    if naver_cookie:
        session_pc.headers.update({"Cookie": naver_cookie})

    pc_base_url = f"https://cafe.naver.com/ArticleList.nhn?search.clubid={CLUB_ID}&search.menuid={MENU_ID}&search.boardtype=L"
    try:
        pc_resp = session_pc.get(pc_base_url, timeout=10)
        pc_soup = BeautifulSoup(pc_resp.text, "html.parser")
        
        iframe = pc_soup.select_one("#cafe_main")
        if iframe and iframe.has_attr("src"):
            iframe_url = "https://cafe.naver.com" + iframe["src"]
            pc_resp = session_pc.get(iframe_url, timeout=10)
            pc_soup = BeautifulSoup(pc_resp.text, "html.parser")

        rows = pc_soup.select("#main-area .article-board tr, #main-area .board-list tr, .article-board tbody tr, .article-board tr")
        if rows:
            print(f"✅ [3차 성공] PC 버전으로 게시글 {len(rows)}개가 확인되었습니다.")
            return crawl_via_pc_html(session_pc, pc_base_url, start_date=start_date, end_date=now_kst)
    except Exception as e:
        print(f"⚠️ 3차 PC 버전 접속 실패: {e}")

    print("\n=======================================================")
    print("❌ [조회 실패 알림] 네이버 카페 글이 조회가 되지 않습니다!")
    print("👉 이유: NAVER_COOKIE 세션 만료, 캡차 차단 또는 해외 IP 제약")
    print("=======================================================\n")
    return False

def crawl_via_mobile_api(session, start_date, end_date):
    all_collected_data = {cat: [] for cat in CATEGORIES}
    page = 1
    stop_crawling = False
    
    while not stop_crawling and page <= 10:
        api_url = f"https://m.cafe.naver.com/api/cafe/{CLUB_ID}/menu/{MENU_ID}/articles?page={page}"
        resp = session.get(api_url)
        if resp.status_code != 200:
            break
            
        res_json = resp.json()
        article_list = res_json.get("message", {}).get("result", {}).get("articleList", [])
        if not article_list:
            break
            
        for item in article_list:
            article_id = item.get("articleId")
            subject = item.get("subject", "")
            write_date_ts = item.get("writeDateTimestamp", 0) / 1000
            art_dt = datetime.fromtimestamp(write_date_ts)
            art_date_str = art_dt.strftime("%Y.%m.%d")
            
            if art_dt < start_date:
                stop_crawling = True
                break
                
            art_url = f"https://m.cafe.naver.com/ArticleRead.nhn?clubid={CLUB_ID}&articleid={article_id}"
            art_resp = session.get(art_url)
            sub_soup = BeautifulSoup(art_resp.text, "html.parser")
            
            parsed_date, art_data = parse_article_content(sub_soup, fallback_date=art_date_str)
            for cat in CATEGORIES:
                if art_data[cat]:
                    all_collected_data[cat].extend(art_data[cat])
                    
        page += 1
        
    save_or_update_excel(all_collected_data)
    return True

def crawl_via_mobile_ajax(session, start_date, end_date):
    all_collected_data = {cat: [] for cat in CATEGORIES}
    page = 1
    stop_crawling = False
    
    while not stop_crawling and page <= 10:
        ajax_url = f"https://m.cafe.naver.com/ArticleListAjax.nhn?search.clubid={CLUB_ID}&search.menuid={MENU_ID}&search.page={page}"
        resp = session.get(ajax_url)
        if resp.status_code != 200:
            break
            
        soup = BeautifulSoup(resp.text, "html.parser")
        links = soup.select("a.article, li a")
        if not links:
            break
            
        for a in links:
            href = a.get("href", "")
            if "articleid=" in href:
                art_url = f"https://m.cafe.naver.com{href}" if href.startswith("/") else href
                art_resp = session.get(art_url)
                sub_soup = BeautifulSoup(art_resp.text, "html.parser")
                parsed_date, art_data = parse_article_content(sub_soup, fallback_date=datetime.now().strftime("%Y.%m.%d"))
                for cat in CATEGORIES:
                    if art_data[cat]:
                        all_collected_data[cat].extend(art_data[cat])
        page += 1
        
    save_or_update_excel(all_collected_data)
    return True

def crawl_via_pc_html(session, base_url, start_date, end_date):
    now_kst = datetime.now(KST).replace(tzinfo=None)
    all_collected_data = {cat: [] for cat in CATEGORIES}
    stop_crawling = False
    page = 1

    while not stop_crawling and page <= 10:
        target_page_url = f"{base_url}&search.page={page}"
        resp = session.get(target_page_url)
        soup = BeautifulSoup(resp.text, "html.parser")
        
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

            title_elem = row.select_one("a.article, a.a-title")
            date_elem = row.select_one("td.td_date, .td_date, .date, .time")
            if not title_elem:
                continue
            
            title_text = title_elem.text.strip()
            date_text = date_elem.text.strip() if date_elem else ""
            
            extracted_date_str = extract_valid_date(title_text)
            if not extracted_date_str:
                if ":" in date_text:
                    extracted_date_str = now_kst.strftime("%Y.%m.%d")
                elif date_text:
                    extracted_date_str = extract_valid_date(date_text)

            if extracted_date_str:
                try:
                    row_date = datetime.strptime(extracted_date_str, "%Y.%m.%d")
                    if row_date < start_date:
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
            for cat in CATEGORIES:
                if art_data[cat]:
                    all_collected_data[cat].extend(art_data[cat])

        if stop_crawling:
            break
        page += 1

    save_or_update_excel(all_collected_data)
    return True

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
                combined_df = combined_df.drop_duplicates(subset=["기사 주소"], keep="first")
                if "연월일" in combined_df.columns:
                    combined_df = combined_df.sort_values(by="연월일", ascending=False)
            combined_df.to_excel(writer, sheet_name=sheet_title, index=False)

    if os.path.exists(NEWS_FILE_PATH):
        with pd.ExcelWriter(NEWS_FILE_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            write_to_excel(writer, new_data_dict, NEWS_FILE_PATH)
    else:
        with pd.ExcelWriter(NEWS_FILE_PATH, engine='openpyxl') as writer:
            write_to_excel(writer, new_data_dict)
    print("✓ 엑셀 업데이트 완료.")

def generate_dashboard_and_sync():
    print("\n[2/3] 대시보드 HTML 파일 생성 중...")
    current_dir = str(Path(__file__).resolve().parent)
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)
    
    from auto_dashboard_module import build_dashboard
    output_html_path = build_dashboard(data_dir=DATA_IN_DIR)
    print("✓ index.html 및 대시보드 HTML 파일 생성/동기화 완료!")

if __name__ == "__main__":
    success = crawl_naver_cafe_news_auto()
    if success:
        generate_dashboard_and_sync()
    else:
        print("⚠️ 카페 게시글 조회가 실패하여 대시보드 업데이트 작업을 중단했습니다.")
