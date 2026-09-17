import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="MSDS 성분 및 유해물질 자동 분석기", layout="wide")
st.title("🧪 MSDS 성분/함유량 추출 & 유해물질 자동 판별 도구")
st.write("MSDS PDF에서 **구성성분명, CAS No., 함유량(%)**을 자동 추출하고 유해물질 DB와 대조합니다.")

CAS_REGEX = r'\b[1-9]\d{1,6}-\d{2}-\d\b'

# 1. 사이드바: DB 등록
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
            info = row.get('규제구분', row.get('물질명', '유해물질 등록됨'))
            target_dict[cas_val] = str(info)
        st.sidebar.success(f"기준 DB 등록 완료: 총 {len(target_dict)}종 로드됨")
    else:
        st.sidebar.error("엑셀 파일 내에 'CAS'가 포함된 열 이름이 필요합니다.")

# 2. 강력해진 MSDS 구성성분 추출 함수
def extract_composition_data(pdf_file):
    results = []
    
    with pdfplumber.open(pdf_file) as pdf:
        full_text = ""
        table_found = False
        
        # 1) 표(Table) 형태 탐색
        for page in pdf.pages:
            t_text = page.extract_text() or ""
            full_text += t_text + "\n"
            
            tables = page.extract_tables()
            for tbl in tables:
                if not tbl or len(tbl) < 2:
                    continue
                
                tbl_flat = " ".join([str(c) for row in tbl for c in row if c])
                # CAS/CAS번호/성분/물질명 키워드가 표 안에 있는지 검사
                if re.search(r'CAS', tbl_flat, re.I) and ('성분' in tbl_flat or '함유' in tbl_flat or '물질명' in tbl_flat or '명칭' in tbl_flat):
                    
                    cas_col_idx = -1
                    name_col_idx = -1
                    content_col_idx = -1
                    
                    # 헤더 열 찾기
                    for idx, col in enumerate(tbl[0]):
                        col_clean = str(col).replace('\n', '').replace(' ', '').upper()
                        if 'CAS' in col_clean:
                            cas_col_idx = idx
                        elif any(k in col_clean for k in ['함유', 'WT', '%', 'CONTENT', '농도']):
                            content_col_idx = idx
                        elif any(k in col_clean for k in ['물질명', '성분', 'NAME', '화학명', 'INGREDIENT']):
                            if name_col_idx == -1:
                                name_col_idx = idx

                    # 데이터 행 추출
                    for row in tbl[1:]:
                        row_cells = [str(c).strip() if c is not None else "" for c in row]
                        row_line = " ".join(row_cells)
                        cas_match = re.search(CAS_REGEX, row_line)
                        
                        if cas_match:
                            cas_no = cas_match.group(0)
                            
                            # 성분명
                            m_name = ""
                            if name_col_idx != -1 and name_col_idx < len(row_cells):
                                m_name = row_cells[name_col_idx].replace('\n', ' ').strip()
                            if not m_name and cas_col_idx > 0:
                                m_name = row_cells[0].replace('\n', ' ').strip()
                                
                            # 함유량
                            cnt = ""
                            if content_col_idx != -1 and content_col_idx < len(row_cells):
                                cnt = row_cells[content_col_idx].replace('\n', ' ').strip()
                            if not cnt:
                                remainder = row_line.replace(cas_no, '')
                                c_match = re.search(r'(\d+(?:\.\d+)?(?:\s*\(.*?\))?%?|\d+\s*~\s*\d+%)', remainder)
                                cnt = c_match.group(0) if c_match else "-"
                                
                            results.append({
                                "구성성분명": m_name if m_name else "-",
                                "CAS No.": cas_no,
                                "함유량": cnt if cnt else "-"
                            })
                            table_found = True

        # 2) 표 인식이 안 되었을 때의 줄 단위 스마트 텍스트 추출 (한일시멘트 스타일 완벽 대응)
        if not table_found:
            # '3. 구성성분...' 섹션부터 '4. 응급조치/처치' 전까지의 본문만 발췌
            sec3_match = re.search(r'(3\.\s*구성\s*성분.*?)(?=4\.\s*응급|\Z)', full_text, re.S)
            sec_text = sec3_match.group(1) if sec3_match else full_text
            
            lines = sec_text.split('\n')
            for line in lines:
                cas_match = re.search(CAS_REGEX, line)
                if cas_match:
                    cas_no = cas_match.group(0)
                    
                    # 파이프(|)로 구분되어 들어온 경우 처리
                    if '|' in line:
                        parts = [p.strip() for p in line.split('|') if p.strip()]
                        m_name = parts[0] if parts else "-"
                        # 함유량은 보통 맨 마지막 파트에 위치
                        cnt = parts[-1] if len(parts) >= 3 and parts[-1] != cas_no else "-"
                    else:
                        parts = line.split(cas_no)
                        m_name = parts[0].strip()
                        cnt = parts[1].strip() if len(parts) > 1 else "-"
                    
                    results.append({
                        "구성성분명": m_name if m_name else "-",
                        "CAS No.": cas_no,
                        "함유량": cnt if cnt else "-"
                    })

    # 중복 제거 (CAS 기준)
    unique_results = []
    seen = set()
    for item in results:
        if item["CAS No."] not in seen:
            seen.add(item["CAS No."])
            unique_results.append(item)
            
    return unique_results

# 3. 메인 화면
st.subheader("📄 2. 검토할 MSDS PDF 파일 업로드")
uploaded_pdfs = st.file_uploader("MSDS PDF 파일을 선택하세요 (여러 개 동시 업로드 가능)", type=["pdf"], accept_multiple_files=True)

if uploaded_pdfs:
    all_rows = []
    
    with st.spinner("MSDS 구성성분 분석 중..."):
        for pdf_file in uploaded_pdfs:
            extracted_items = extract_composition_data(pdf_file)
            
            if not extracted_items:
                all_rows.append({
                    "파일명": pdf_file.name,
                    "구성성분명": "미발견 (스캔본 이미지이거나 성분 정보 없음)",
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
    
    st.subheader("📊 3. 분석 결과")
    st.dataframe(df_results, use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_results.to_excel(writer, index=False, sheet_name='MSDS검토결과')
    
    st.download_button(
        label="📥 결과 엑셀 파일 다운로드",
        data=buffer.getvalue(),
        file_name="MSDS_유해성분_검토결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
