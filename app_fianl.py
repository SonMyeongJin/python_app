import streamlit as st
import pandas as pd
import tempfile
import zipfile
import os
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="(주)건화 등기부등본 Excel 통합기", layout="wide")

password = st.text_input('비밀번호를 입력하세요', type='password')
if password != '1220':
    st.warning('올바른 비밀번호를 입력하세요.')
    st.stop()

st.title("📦 (주)건화 등기부등본 통합분석기")
st.markdown("""
압축파일(.zip) 안의 폴더 구조와 관계없이 모든 엑셀 파일을 자동 분석합니다.
""")


def merge_adjacent_cells(row_series, max_gap=3):
    """
    인접한 셀들을 병합하여 하나의 의미있는 단위로 만드는 함수
    얇은 선으로 나뉜 셀들을 통합
    """
    merged_row = row_series.copy()
    row_dict = row_series.to_dict()
    
    # 빈 셀이 아닌 셀들의 인덱스를 찾기
    non_empty_indices = [idx for idx, val in row_dict.items() if str(val).strip()]
    
    # 연속된 셀들을 그룹화
    groups = []
    current_group = []
    
    for i, idx in enumerate(non_empty_indices):
        if not current_group:
            current_group = [idx]
        else:
            # 이전 인덱스와의 거리가 max_gap 이하면 같은 그룹
            if idx - current_group[-1] <= max_gap:
                current_group.append(idx)
            else:
                # 새로운 그룹 시작
                groups.append(current_group)
                current_group = [idx]
    
    if current_group:
        groups.append(current_group)
    
    # 각 그룹 내의 셀들을 병합
    for group in groups:
        if len(group) > 1:
            # 그룹 내 모든 값을 연결
            merged_value = ""
            for idx in group:
                val = str(row_dict.get(idx, "")).strip()
                if val:
                    if merged_value and not merged_value.endswith((" ", "-", "/")):
                        merged_value += " "
                    merged_value += val
            
            # 첫 번째 인덱스에 병합된 값 저장
            merged_row[group[0]] = merged_value
            
            # 나머지 인덱스는 빈 값으로 설정
            for idx in group[1:]:
                merged_row[idx] = ""
    
    return merged_row

def merge_dataframe_cells(df):
    """
    데이터프레임 전체에 셀 병합 로직 적용
    """
    if df.empty:
        return df
    
    merged_df = df.copy()
    
    # 각 행에 대해 셀 병합 적용
    for i in range(len(merged_df)):
        merged_df.iloc[i] = merge_adjacent_cells(merged_df.iloc[i])
    
    return merged_df













uploaded_zip = st.file_uploader("📁 .zip 파일을 업로드하세요 (내부에 .xlsx 파일 포함)", type=["zip"])
run_button = st.button("분석 시작")

if run_button and uploaded_zip:
    temp_dir = tempfile.mkdtemp()
    szj_list, syg_list, djg_list = [], [], []

    with zipfile.ZipFile(uploaded_zip, "r") as z:
        z.extractall(temp_dir)

    # ✅ 하위 폴더 포함 모든 .xlsx 탐색
    excel_files = []
    for root, _, files in os.walk(temp_dir):
        for f in files:
            if f.lower().endswith(".xlsx"):
                excel_files.append(os.path.join(root, f))

    for path in excel_files:
        try:
            xls = pd.ExcelFile(path)
            df = xls.parse(xls.sheet_names[0]).fillna("")
            name = extract_identifier(df)

            szj_sec, has_szj = extract_section_range(df, "소유지분현황", ["소유권", "저당권"], match_fn=keyword_match_partial)
            syg_sec, has_syg = extract_section_range(df, "소유권.*사항", ["저당권"], match_fn=keyword_match_exact)
            djg_sec, has_djg = extract_section_range(df, "3.(근)저당권및전세권등(을구)", ["참고", "비고", "총계", "전산자료"], match_fn=keyword_match_exact)

            if has_szj:
                szj_df = extract_named_cols(szj_sec, ["등기명의인", "(주민)등록번호", "최종지분", "주소", "순위번호"])
                szj_df.insert(0, "파일명", name)
                szj_list.append(szj_df)
            else:
                szj_list.append(pd.DataFrame([[name, "기록없음"]], columns=["파일명", "등기명의인"]))

            if has_syg:
                syg_df = extract_precise_named_cols(syg_sec, ["순위번호", "등기목적", "접수정보", "주요등기사항", "대상소유자"])
                syg_df.insert(0, "파일명", name)
                syg_list.append(syg_df)
            else:
                syg_list.append(pd.DataFrame([[name, "기록없음"]], columns=["파일명", "순위번호"]))

            if has_djg:
                djg_df = extract_precise_named_cols(djg_sec, ["순위번호", "등기목적", "접수정보", "주요등기사항", "대상소유자"])
                djg_df = merge_same_row_if_amount_separated(djg_df)  # ✅ 여기에 병합 함수 호출 추가
                djg_df = trim_after_reference_note(djg_df)
                djg_df.insert(0, "파일명", name)
                djg_list.append(djg_df)
            else:
                djg_list.append(pd.DataFrame([[name, "기록없음"]], columns=["파일명", "순위번호"]))
        except Exception as e:
            pass  # 또는 logging.warning(...) 등으로 로깅만
    wb = Workbook()
    for sheetname, data in zip(
        ["1. 소유지분현황 (갑구)", "2. 소유권사항 (갑구)", "3. 저당권사항 (을구)"],
        [szj_list, syg_list, djg_list]
    ):
        ws = wb.create_sheet(title=sheetname)
        if data:
            df = pd.concat(data, ignore_index=True)
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)
        else:
            ws.append(["기록없음"])

    wb.remove(wb["Sheet"])
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name)
        st.success("✅ 분석 완료! 다운로드 버튼을 클릭하세요.")
        st.download_button("📥 결과 다운로드", data=open(tmp.name, "rb"), file_name="등기사항_통합_시트별구성.xlsx")
