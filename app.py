import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="MSDS 성분 및 유해물질 자동 분석기", layout="wide")
st.title("🧪 MSDS 성분/함유량 추출 & 유해물질 자동 판별 도구")
st.write("MSDS PDF에서 **구성성분명, CAS No., 함유량(%)**을 자동으로 인식하여 텍스트화하고 유해물질 DB와 대조합니다.")

# CAS 번호 정규식
CAS_REGEX = r'\b[1-9]\d{1,6}-\d{2}-\d\b'

# 1. 사이드바: 유해물질 기준 DB 등록
st.sidebar.header("📁 1. 유해물질 기준 DB 등록")
db_file = st.sidebar.file_uploader("유해물질 목록 엑셀 (.xlsx, .xls)", type=["xlsx", "xls"])

target_dict = {}
if db_file:
    df_db = pd.read_excel(db_file)
    cas_col = [c for c in df_db.columns if 'CAS' in str(c).upper()]
    if cas_col:
        c_name = cas_col[0]
        for _, row in df_db.iterrows():
            cas_val = str(row[c_name]).strip()
            # 규제구분, 비고, 물질명 중 있는 값을 가져옴
            info = row.get('규제구분', row.get('물질명', '유해물질 등록됨'))
            target_dict[cas_val] = str(info)
        st.sidebar.success(f"기준 DB 등록 완료: 총 {len(target_dict)}종 로드됨")
    else:
        st.sidebar.error("엑셀 파일 내에 'CAS'가 포함된 열 이름이 필요합니다.")

# 2. MSDS 구성성분 추출 핵심 함수
def extract_composition_data(pdf_file):
    results = []
    
    with pdfplumber.open(pdf_file) as pdf:
        # 1단계: 표(Table) 형태가 인식되는지 우선 검사 (3번 항목 표 추출)
        table_found = False
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for tbl in tables:
                if not tbl:
                    continue
                # 표 내용 중 CAS No나 구성성분 관련 단어가 있는지 확인
                tbl_text = " ".join([str(c) for row in tbl for c in row if c])
                if re.search(r'CAS\s*(No|번호)?', tbl_text, re.I) and ('성분' in tbl_text or '함유' in tbl_text or 'Ingredient' in tbl_text or 'Composition' in tbl_text):
                    # 표 헤더 인덱스 파악
                    header = [str(col).replace('\n', ' ').strip() for col in tbl[0] if col is not None]
                    
                    cas_idx = -1
                    name_idx = -1
                    content_idx = -1
                    
                    for idx, col_name in enumerate(tbl[0]):
                        c_str = str(col_name).upper()
                        if 'CAS' in c_str:
                            cas_idx = idx
                        elif '함유' in c_str or 'WT' in c_str or '%' in c_str or 'CONTENT' in c_str:
                            content_idx = idx
                        elif '성분' in c_str or 'NAME' in c_str or '명칭' in c_str or 'INGREDIENT' in c_str:
                            if name_idx == -1:
                                name_idx = idx

                    # 표의 데이터 행 파싱
                    for row in tbl[1:]:
                        row_cells = [str(c).strip() if c else "" for c in row]
                        row_full_text = " ".join(row_cells)
                        cas_match = re.search(CAS_REGEX, row_full_text)
                        
                        if cas_match:
                            cas_no = cas_match.group(0)
                            
                            # 성분명 추출
                            c_name = row_cells[name_idx] if (name_idx != -1 and name_idx < len(row_cells)) else ""
                            c_name = c_name.replace('\n', ' ')
                            # 성분명이 비었으면 CAS 앞의 셀들 조합
                            if not c_name and cas_idx > 0:
                                c_name = " ".join([r for r in row_cells[:cas_idx] if r and r != cas_no])
                                
                            # 함유량 추출
                            content = row_cells[content_idx] if (content_idx != -1 and content_idx < len(row_cells)) else ""
                            content = content.replace('\n', ' ')
                            # 함유량 컬럼을 못 찾았을 경우 숫자 패턴 추출 (예: 12.5(10~15), 70%, 5-10 등)
                            if not content:
                                cnt_match = re.search(r'(\d+(?:\.\d+)?(?:\s*\(.*?\))?%?|\d+\s*~\s*\d+%)', row_full_text.replace(cas_no, ''))
                                content = cnt_match.group(0) if cnt_match else "-"
                                
                            results.append({
                                "구성성분명": c_name if c_name else "-",
                                "CAS No.": cas_no,
                                "함유량": content if content else "-"
                            })
                            table_found = True

        # 2단계: 표가 깨졌거나 텍스트로만 나열된 경우 (3번 섹션 텍스트 블록 기반 백업 추출)
        if not table_found:
            full_text = ""
            for page in pdf.pages:
                full_text += (page.extract_text() or "") + "\n"
            
            # '3. 구성성분'부터 '4. 응급조치' 사이의 본문만 슬라이싱
            sec3_match = re.search(r'(3\.\s*(?:구성성분|혼합물의\s*구성|성분명칭).*?)(?=4\.\s*(?:응급조치|응급처치)|$)', full_text, re.S)
            target_chunk = sec3_match.group(1) if sec3_match else full_text
            
            lines = target_chunk.split('\n')
            for line in lines:
                cas_match = re.search(CAS_REGEX, line)
                if cas_match:
                    cas_no = cas_match.group(0)
                    parts = line.split(cas_no)
                    name_part = parts[0].strip().strip('|').strip()
                    content_part = parts[1].strip().strip('|').strip() if len(parts) > 1 else "-"
                    
                    results.append({
                        "구성성분명": name_part if name_part else "-",
                        "CAS No.": cas_no,
                        "함유량": content_part if content_part else "-"
                    })

    # 중복 제거 (순서 유지)
    unique_results = []
    seen_cas = set()
    for item in results:
        if item["CAS No."] not in seen_cas:
            seen_cas.add(item["CAS No."])
            unique_results.append(item)
            
    return unique_results

# 3. 메인 화면 업로드 및 분석
st.subheader("📄 2. 검토할 MSDS PDF 파일 업로드")
uploaded_pdfs = st.file_uploader("MSDS PDF 파일을 드래그해서 올려주세요 (여러 개 가능)", type=["pdf"], accept_multiple_files=True)

if uploaded_pdfs:
    all_rows = []
    
    with st.spinner("MSDS 구성성분 및 함유량 분석 중..."):
        for pdf_file in uploaded_pdfs:
            extracted_items = extract_composition_data(pdf_file)
            
            if not extracted_items:
                all_rows.append({
                    "파일명": pdf_file.name,
                    "구성성분명": "미발견(스캔본 또는 양식 상이)",
                    "CAS No.": "-",
                    "함유량": "-",
                    "유해물질 해당 여부": "-",
                    "상세 규제 정보": "-"
                })
            else:
                for item in extracted_items:
                    cas = item["CAS No."]
                    is_hazard = cas in target_dict if target_dict else "DB 미등록"
                    hazard_info = target_dict.get(cas, "-") if target_dict else "-"
                    
                    all_rows.append({
                        "파일명": pdf_file.name,
                        "구성성분명": item["구성성분명"],
                        "CAS No.": cas,
                        "함유량": item["함유량"],
                        "유해물질 해당 여부": "⚠️ 해당" if (is_hazard is True) else ("정상 (미해당)" if is_hazard is False else "확인 필요"),
                        "상세 규제 정보": hazard_info
                    })

    df_results = pd.DataFrame(all_rows)
    
    # 4. 결과 출력
    st.subheader("📊 3. 추출 및 검별 결과 (텍스트/표)")
    st.dataframe(df_results, use_container_width=True)

    # 5. 엑셀 다운로드
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_results.to_excel(writer, index=False, sheet_name='MSDS_성분_검토결과')
    
    st.download_button(
        label="📥 결과 엑셀 파일 다운로드",
        data=buffer.getvalue(),
        file_name="MSDS_구성성분_검토결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
