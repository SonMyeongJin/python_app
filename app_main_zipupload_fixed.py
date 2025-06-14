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
    """분리된 헤더를 병합하는 함수 - 개선된 버전"""
    # 먼저 인접 셀 병합 적용
    merged_row = merge_adjacent_cells(header_row)
    
    # 기존 특정 키워드 병합 로직도 유지
    split_patterns = {
        "주소": ["주", "소"],
        "등기명의인": ["등기", "명의인"],
        "주민등록번호": ["주민", "등록번호"],
        "최종지분": ["최종", "지분"],
        "순위번호": ["순위", "번호"],
        "등기목적": ["등기", "목적"],
        "접수정보": ["접수", "정보"],
        "주요등기사항": ["주요", "등기사항"],
        "대상소유자": ["대상", "소유자"]
    }
    
    for target_keyword, split_parts in split_patterns.items():
        found_indices = []
        for part in split_parts:
            for idx, cell_value in merged_row.items():
                cell_str = str(cell_value).strip()
                if cell_str == part:
                    found_indices.append(idx)
                    break
        
        if len(found_indices) == len(split_parts):
            if all(found_indices[i+1] - found_indices[i] <= 3 for i in range(len(found_indices)-1)):
                merged_row[found_indices[0]] = target_keyword
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
    
    # 전체 섹션에 셀 병합 적용
    section = merge_dataframe_cells(section)
    
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

    # 최종지분 처리 로직은 기존과 동일하게 유지
    if "최종지분" not in col_map:
        idx_최종 = None
        idx_지분 = None
        for idx, val in merged_header.items():
            if str(val).strip() == "최종":
                idx_최종 = idx
            if str(val).strip() == "지분":
                idx_지분 = idx
        if idx_최종 is not None and idx_지분 is not None and abs(idx_최종 - idx_지분) <= 3:
            col_map["최종지분"] = (min(idx_최종, idx_지분), max(idx_최종, idx_지분))

    rows = []
    for i in range(1, len(section)):
        row = section.iloc[i]
        row_dict = {}
        for key in col_keywords:
            if key == "최종지분":
                if isinstance(col_map.get("최종지분"), tuple):
                    idx1, idx2 = col_map["최종지분"]
                    val1 = str(row.get(idx1, "")).strip()
                    val2 = str(row.get(idx2, "")).strip()
                    if val1 and val2:
                        row_dict[key] = val1 + val2
                    else:
                        row_dict[key] = val1 or val2
                elif isinstance(col_map.get("최종지분"), int):
                    idx = col_map["최종지분"]
                    val1 = str(row.get(idx, "")).strip()
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
        
        # 등기명의인과 주민번호 분리 처리
        if "등기명의인" in row_dict and "(주민)등록번호" in col_keywords:
            owner_text = str(row_dict["등기명의인"]).strip()
            
            # 주민등록번호가 등기명의인 필드에 있는 경우
            jumin = extract_jumin_number(owner_text)
            if jumin:
                # 주민번호는 주민등록번호 필드에 넣고, 등기명의인에서는 제거
                row_dict["(주민)등록번호"] = jumin
                row_dict["등기명의인"] = owner_text.replace(jumin, "").strip()
        
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
    # 전체 섹션에 셀 병합 적용
    section = merge_dataframe_cells(section)
    
    header_idx, header_row = find_keyword_header(section, col_keywords)
    if header_idx is None:
        header_row = merge_split_headers(section.iloc[0])
        start_row = 1
    else:
        header_row = merge_split_headers(header_row)
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
def is_jumin_number(text):
    """
    주민등록번호 패턴을 확인하는 함수
    예: 123456-1234567 또는 123456-*******
    """
    if not isinstance(text, str):
        return False
    
    # 주민등록번호 패턴 (숫자6자리-숫자또는*)
    pattern = re.compile(r'\d{6}-[\d\*]+')
    return bool(re.search(pattern, text))

def extract_jumin_number(text):
    """
    문자열에서 주민등록번호 패턴을 추출
    """
    if not isinstance(text, str):
        return ""
    
    pattern = re.compile(r'\d{6}-[\d\*]+')
    match = re.search(pattern, text)
    return match.group(0) if match else ""

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
                
                # 데이터 후처리 - 등기명의인과 주민등록번호 정리
                for idx, row in szj_df.iterrows():
                    # 등기명의인에서 주민번호 패턴이 있으면 분리
                    if pd.notna(row["등기명의인"]):
                        jumin = extract_jumin_number(str(row["등기명의인"]))
                        if jumin:
                            szj_df.at[idx, "(주민)등록번호"] = jumin
                            szj_df.at[idx, "등기명의인"] = str(row["등기명의인"]).replace(jumin, "").strip()
                
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
