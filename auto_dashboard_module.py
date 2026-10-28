import os
import sys
import json
import re
import pandas as pd
from pathlib import Path

def build_dashboard(data_dir=None):
    """
    DATA/IN/보험뉴스_Data.xlsx 파일의 카테고리별 시트를 읽어와
    index.html 내의 뉴스 데이터 변수를 최신 데이터로 동적 갱신 후 빌드하는 함수
    """
    if data_dir is None:
        base_dir = Path(__file__).resolve().parent
        data_dir = base_dir / "DATA" / "IN"
    else:
        data_dir = Path(data_dir)

    news_excel_path = data_dir / "보험뉴스_Data.xlsx"
    root_dir = data_dir.parent.parent
    index_html_path = root_dir / "index.html"
    out_dir = data_dir.parent / "OUT"
    os.makedirs(out_dir, exist_ok=True)
    
    output_html_path = out_dir / "보험회사_종합분석_대시보드.html"

    # 1. 엑셀 데이터 파싱 및 JSON 구조 변환
    news_data_dict = {}
    if news_excel_path.exists():
        try:
            xls = pd.ExcelFile(news_excel_path)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                # 데이터 정제
                df = df.fillna("")
                records = df.to_dict(orient="records")
                news_data_dict[sheet_name] = records
            print(f"✓ 엑셀 뉴스 데이터 로드 완료 ({len(news_data_dict)}개 카테고리)")
        except Exception as e:
            print(f"⚠️ 엑셀 뉴스 데이터 읽기 실패: {e}")

    # 2. index.html 읽기
    if not index_html_path.exists():
        raise FileNotFoundError(f"⚠️ {index_html_path} 파일이 존재하지 않습니다.")

    with open(index_html_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    # 3. HTML 내부의 NEWS_DATA / newsData JavaScript 데이터 영역 동적 치환
    if news_data_dict:
        json_news_data = json.dumps(news_data_dict, ensure_ascii=False, indent=2)
        
        # Pattern 1: const NEWS_DATA = { ... }; 또는 var newsData = { ... };
        pattern = r"(const\s+NEWS_DATA\s*=\s*|var\s+newsData\s*=\s*|let\s+newsData\s*=\s*)(\{[\s\S]*?\});"
        
        if re.search(pattern, html_content):
            html_content = re.sub(
                pattern,
                r"\1" + json_news_data.replace("\\", "\\\\") + ";",
                html_content,
                count=1
            )
            print("✓ HTML 내부 뉴스 데이터 변수(NEWS_DATA) 치환 완료")
        else:
            # 주입 스크립트 태그가 별도로 없을 경우 script 태그 상단에 삽입
            inject_script = f"\n<script>\nconst NEWS_DATA = {json_news_data};\n</script>\n"
            html_content = html_content.replace("</head>", f"{inject_script}</head>")
            print("✓ HTML <head> 내 뉴스 데이터(NEWS_DATA) 신규 주입 완료")

    # 4. DATA/OUT 폴더 및 root index.html 파일 동기화 저장
    with open(output_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    with open(index_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"✓ 대시보드 대시보드 빌드 및 동기화 완료: {output_html_path}")
    return output_html_path

if __name__ == "__main__":
    build_dashboard()
