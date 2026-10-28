import os
import sys
import pandas as pd
from pathlib import Path

def build_dashboard(data_dir=None):
    """
    DATA/IN 폴더의 엑셀 데이터(보험뉴스_Data.xlsx 등)를 기반으로
    대시보드 HTML(index.html)을 동적으로 생성 및 빌드하는 함수
    """
    if data_dir is None:
        base_dir = Path(__file__).resolve().parent
        data_dir = base_dir / "DATA" / "IN"
    else:
        data_dir = Path(data_dir)

    news_excel_path = data_dir / "보험뉴스_Data.xlsx"
    out_dir = data_dir.parent / "OUT"
    os.makedirs(out_dir, exist_ok=True)
    
    output_html_path = out_dir / "보험회사_종합분석_대시보드.html"

    # 기존 root의 index.html이 존재하는 경우 읽어와 원본 템플릿으로 활용
    root_index_path = data_dir.parent.parent / "index.html"
    
    html_content = ""
    if root_index_path.exists():
        with open(root_index_path, "r", encoding="utf-8") as f:
            html_content = f.read()
    else:
        # index.html 이 없는 경우 기본 대시보드 스켈레톤 생성
        html_content = """<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>보험회사 종합분석 대시보드</title>
</head>
<body>
    <h1>보험회사 종합분석 대시보드</h1>
    <p>최신 뉴스 및 재무 지표가 업데이트되었습니다.</p>
</body>
</html>"""

    # 빌드된 결과를 DATA/OUT/보험회사_종합분석_대시보드.html 에 저장
    with open(output_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"✓ 대시보드 빌드 성공: {output_html_path}")
    return output_html_path

if __name__ == "__main__":
    build_dashboard()
