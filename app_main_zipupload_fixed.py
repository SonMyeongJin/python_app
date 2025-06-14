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

uploaded_zip = st.file_uploader("📁 .zip 파일을 업로드하세요 (내부에 .xlsx 파일 포함)", type=["zip"])
run_button = st.button("분석 시작")

def trim_after_reference_note(df):
    for i, row in df.iterrows():
        row_text = "".join(str(cell) for cell in row)
        normalized = re.sub(r"\s+", "", row_text)
        if "참고사항" in normalized or "참고" in normalized or "비고" in normalized:
            return df.iloc[:i]
    return df

def extract_identifier(df):
    for i in range(len(df)):
        row = df.iloc[i]
        row_text = " ".join(str(cell) for cell in row)
        if "고유번호" in row_text:
            for j in range(i+1, min(i+10, len(df))):
                content = " ".join(str(cell) for cell in df.iloc[j])
                if content.strip().startswith(("[토지]", "[건물]")):
                    return content.strip()
            break
    return "알수없음"

def keyword_match_partial(cell, keyword):
    if pd.isnull(cell): return False
    return keyword.replace(" ", "") in str(cell).replace(" ", "")

def keyword_match_exact(cell, keyword):
    if pd.isnull(cell): return False
    return re.sub(r"\s+", "", str(cell)) == re.sub(r"\s+", "", keyword)

def merge_split_headers(header_row):
    """분리된 헤더를 병합하는 함수"""
    merged_row = header_row.copy()
    
    # 주요 키워드들의 분리 패턴 정의
    split_patterns = {
        "주소": ["주", "소"],
        "등기명의인": ["등기", "명의인"],
        "주민등록번호": ["주민", "등록번호"],
        "최종지분": ["최종", "지분"],
        "순위번호": ["순위", "번호"]
    }
    
    for target_keyword, split_parts in split_patterns.items():
        # 분리된 부분들을 찾아서 병합
        found_indices = []
        for part in split_parts:
            for idx, cell_value in header_row.items():
                cell_str = str(cell_value).strip()
                if cell_str == part:
                    found_indices.append(idx)
                    break
        
        # 연속된 인덱스인지 확인하고 병합
        if len(found_indices) == len(split_parts):
            # 인덱스가 연속적인지 또는 하나씩 떨어져 있는지 확인
            if all(found_indices[i+1] - found_indices[i] <= 2 for i in range(len(found_indices)-1)):
                # 첫 번째 위치에 병합된 값 저장
                merged_row[found_indices[0]] = target_keyword
                # 나머지 위치는 빈 값으로 설정
                for idx in found_indices[1:]:
                    merged_row[idx] = ""
    
    return merged_row

def enhanced_keyword_match(header_row, keyword, max_distance=2):
    """인접한 셀들을 고려한 키워드 매칭"""
    # 먼저 일반적인 매칭 시도
    for idx, cell in header_row.items():
        if keyword_match_partial(cell, keyword):
            return idx
    
    # 분리된 키워드 매칭 시도
    keyword_chars = list(keyword.replace(" ", ""))
    if len(keyword_chars) <= 1:
        return None
    
    for start_idx, cell in header_row.items():
        if str(cell).strip() == keyword_chars[0]:
            # 첫 글자가 매칭되면 다음 글자들을 인접 셀에서 찾기
            current_text = str(cell).strip()
            current_idx = start_idx
            
            for i in range(1, len(keyword_chars)):
                found_next = False
                # 최대 max_distance까지 떨어진 셀에서 다음 글자 찾기
                for offset in range(1, max_distance + 1):
                    next_idx = current_idx + offset
                    if next_idx in header_row:
                        next_cell = str(header_row[next_idx]).strip()
                        if next_cell == keyword_chars[i]:
                            current_text += next_cell
                            current_idx = next_idx
                            found_next = True
                            break
                
                if not found_next:
                    break
            
            # 전체 키워드가 매칭되었는지 확인
            if current_text == keyword.replace(" ", ""):
                return start_idx
    
    return None

def extract_section_range(df, start_kw, end_kw_list, match_fn):
    df = df.fillna("")
    df.columns = range(df.shape[1])
    start_idx, end_idx = None, len(df)
    for i, row in df.iterrows():
        if any(match_fn(cell, start_kw) for cell in row):
            start_idx = i + 1
            break
    if start_idx is None:
        return pd.DataFrame(), False
    for i in range(start_idx, len(df)):
        row = df.iloc[i]
        if any(any(match_fn(cell, end_kw) for cell in row) for end_kw in end_kw_list):
            end_idx = i
            break
    section = df.iloc[start_idx:end_idx].copy()
    is_empty = section.replace("", pd.NA).dropna(how="all").empty
    return section if not is_empty else pd.DataFrame([["기록없음"]]), not is_empty

# 소유지분현황(갑구)에서 필요한 열을 추출
def extract_named_cols(section, col_keywords):
    if section.empty:
        return pd.DataFrame([["기록없음"]])
    
    header_row = section.iloc[0]
    merged_header = merge_split_headers(header_row)
    
    col_map = {}
    for target in col_keywords:
        col_idx = enhanced_keyword_match(merged_header, target)
        if col_idx is not None:
            col_map[target] = col_idx
        else:
            for idx, val in merged_header.items():
                if keyword_match_partial(val, target):
                    col_map[target] = idx
                    break

    # --- 최종지분이 두 칸으로 나뉘어 있는 경우 합치기 ---
    # '최종지분'이 col_map에 없고, '최종'과 '지분'이 연속된 칸에 있으면 합치기
    # 셀 병합되어 한 셀에 '20분의10' 등으로 들어있는 경우는 그대로 사용됨
    if "최종지분" not in col_map:
        idx_최종 = None
        idx_지분 = None
        for idx, val in merged_header.items():
            if str(val).strip() == "최종":
                idx_최종 = idx
            if str(val).strip() == "지분":
                idx_지분 = idx
        if idx_최종 is not None and idx_지분 is not None and abs(idx_최종 - idx_지분) == 1:
            col_map["최종지분"] = (min(idx_최종, idx_지분), max(idx_최종, idx_지분))

    rows = []
    for i in range(1, len(section)):
        row = section.iloc[i]
        row_dict = {}
        for key in col_keywords:
            if key == "최종지분":
                # 1. 헤더가 분리된 경우(최종/지분)
                if isinstance(col_map.get("최종지분"), tuple):
                    idx1, idx2 = col_map["최종지분"]
                    val1 = str(row.get(idx1, "")).strip()
                    val2 = str(row.get(idx2, "")).strip()
                    if val1 and val2:
                        row_dict[key] = val1 + val2
                    else:
                        row_dict[key] = val1 or val2
                # 2. 헤더가 한 칸인데 데이터가 두 칸으로 나뉜 경우(엑셀 병합이 아닌 실제 데이터가 두 칸)
                elif isinstance(col_map.get("최종지분"), int):
                    idx = col_map["최종지분"]
                    val1 = str(row.get(idx, "")).strip()
                    # 다음 칸 값도 확인 (엑셀에서 병합이 안 된 경우)
                    val2 = ""
                    if (idx + 1) in row and not str(merged_header.get(idx + 1, "")).strip():
                        val2 = str(row.get(idx + 1, "")).strip()
                    if val1 and val2:
                        row_dict[key] = val1 + val2
                    else:
                        row_dict[key] = val1 or val2
                else:
                    row_dict[key] = ""
            elif key in col_map:
                row_dict[key] = row.get(col_map[key], "")
            else:
                row_dict[key] = ""
        rows.append(row_dict)
    return pd.DataFrame(rows)

def find_keyword_header(section, col_keywords, max_search_rows=15):
    section = section.fillna("").astype(str)
    for i in range(min(max_search_rows, len(section))):
        row = section.iloc[i]
        match_count = sum(any(keyword_match_exact(cell, kw) for cell in row) for kw in col_keywords)
        if match_count >= 3:
            return i, row
    return None, None

def find_col_index(header_row, keyword):
    for idx, val in header_row.items():
        if keyword_match_exact(val, keyword):
            return idx
    return None

# 소유권사항 (갑구)와 에서 필요한 열 추출
def extract_precise_named_cols(section, col_keywords):
    header_idx, header_row = find_keyword_header(section, col_keywords)
    if header_idx is None:
        header_row = section.iloc[0]
        start_row = 1
    else:
        start_row = header_idx + 1
    col_map = {key: find_col_index(header_row, key) for key in col_keywords if find_col_index(header_row, key) is not None}
    if not col_map:
        return pd.DataFrame([["기록없음"]])
    rows = []
    for i in range(start_row, len(section)):
        row = section.iloc[i]
        row_dict = {key: row[col_map[key]] if col_map[key] in row else "" for key in col_map}
        rows.append(row_dict)
    return pd.DataFrame(rows)
def merge_same_row_if_amount_separated(df):
    df = df.copy()
    for i in range(len(df) - 1):
        row = df.iloc[i]
        main = str(row["주요등기사항"])

        if "채권최고액" in main:
            # 현재 행과 다음 행 모두 병합 텍스트 구성
            combined_row = list(row.values) + list(df.iloc[i + 1].values)
            combined_text = " ".join(str(x) for x in combined_row if pd.notnull(x))

            # 금액 패턴 추출
            match = re.search(r"금[\d,]+원", combined_text)
            if match and match.group(0) not in main:
                df.at[i, "주요등기사항"] = main + " " + match.group(0)
    return df
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
