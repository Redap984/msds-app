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

# 2. MSDS 구성성분 추출 함수
def extract_composition_data(pdf_file):
    results = []
    has_trade_secret = False
    
    with pdfplumber.open(pdf_file) as pdf:
        full_text = ""
        table_found = False
        
        # 전체 텍스트 수집 및 영업비밀 키워드 확인
        for page in pdf.pages:
            t_text = page.extract_text() or ""
            full_text += t_text + "\n"
            
            tables = page.extract_tables()
            for tbl in tables:
                if not tbl or len(tbl) < 2:
                    continue
                
                tbl_flat = " ".join([str(c) for row in tbl for c in row if c])
                if re.search(r'CAS', tbl_flat, re.I) and any(k in tbl_flat for k in ['성분', '함유', '물질명', '명칭']):
                    cas_col_idx = -1
                    name_col_idx = -1
                    content_col_idx = -1
                    
                    for idx, col in enumerate(tbl[0]):
                        col_clean = str(col).replace('\n', '').replace(' ', '').upper()
                        if 'CAS' in col_clean:
                            cas_col_idx = idx
                        elif any(k in col_clean for k in ['함유', 'WT', '%', 'CONTENT', '농도']):
                            content_col_idx = idx
                        elif any(k in col_clean for k in ['물질명', '성분', 'NAME', '화학명', 'INGREDIENT']):
                            if name_col_idx == -1:
                                name_col_idx = idx

                    for row in tbl[1:]:
                        row_cells = [str(c).strip() if c is not None else "" for c in row]
                        row_line = " ".join(row_cells)
                        
                        if '영업비밀' in row_line or 'Trade Secret' in row_line:
                            has_trade_secret = True
                            
                        cas_match = re.search(CAS_REGEX, row_line)
                        if cas_match:
                            cas_no = cas_match.group(0)
                            m_name = row_cells[name_col_idx].replace('\n', ' ').strip() if (name_col_idx != -1 and name_col_idx < len(row_cells)) else ""
                            if not m_name and cas_col_idx > 0:
                                m_name = row_cells[0].replace('\n', ' ').strip()
                                
                            cnt = row_cells[content_col_idx].replace('\n', ' ').strip() if (content_col_idx != -1 and content_col_idx < len(row_cells)) else ""
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

        if not table_found:
            sec3_match = re.search(r'(3\.\s*구성\s*성분.*?)(?=4\.\s*응급|\Z)', full_text, re.S)
            sec_text = sec3_match.group(1) if sec3_match else full_text
            
            if '영업비밀' in sec_text or 'Trade Secret' in sec_text:
                has_trade_secret = True
                
            lines = sec_text.split('\n')
            for line in lines:
                cas_match = re.search(CAS_REGEX, line)
                if cas_match:
                    cas_no = cas_match.group(0)
                    if '|' in line:
                        parts = [p.strip() for p in line.split('|') if p.strip()]
                        m_name = parts[0] if parts else "-"
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

    unique_results = []
    seen = set()
    for item in results:
        if item["CAS No."] not in seen:
            seen.add(item["CAS No."])
            unique_results.append(item)
            
    return unique_results, has_trade_secret

# 3. 메인 화면 업로드 및 분석
st.subheader("📄 2. 검토할 MSDS PDF 파일 업로드")
uploaded_pdfs = st.file_uploader("MSDS PDF 파일을 선택하세요 (여러 개 동시 업로드 가능)", type=["pdf"], accept_multiple_files=True)

if uploaded_pdfs:
    all_rows = []
    warning_count = 0
    
    with st.spinner("MSDS 구성성분 분석 중..."):
        for pdf_file in uploaded_pdfs:
            extracted_items, has_trade_secret = extract_composition_data(pdf_file)
            
            if not extracted_items:
                warning_count += 1
                status_text = "🚨 영업비밀 표기 의심 (성분 미기재)" if has_trade_secret else "🚨 성분표 미발견 (스캔본/양식확인 필요)"
                all_rows.append({
                    "파일명": pdf_file.name,
                    "구성성분명": "공급사 성분 미기재",
                    "CAS No.": "-",
                    "함유량": "-",
                    "판정결과": status_text,
                    "상세 규제 정보": "제조사/공급처에 최신 개정본 또는 성분확인서 징구 요망"
                })
            else:
                for item in extracted_items:
                    cas = item["CAS No."]
                    is_hazard = cas in target_dict if target_dict else "DB 미등록"
                    hazard_info = target_dict.get(cas, "-") if target_dict else "-"
                    
                    verdict = "⚠️ 유해물질 해당" if (is_hazard is True) else ("정상 (미해당)" if is_hazard is False else "확인 필요")
                    
                    all_rows.append({
                        "파일명": pdf_file.name,
                        "구성성분명": item["구성성분명"],
                        "CAS No.": cas,
                        "함유량": item["함유량"],
                        "판정결과": verdict,
                        "상세 규제 정보": hazard_info
                    })

    df_results = pd.DataFrame(all_rows)
    
    # 4. 상단 요약 배너
    st.subheader("📊 3. 분석 결과")
    if warning_count > 0:
        st.warning(f"⚠️ **주의**: 3번 성분표가 누락되었거나 영업비밀로 표기된 파일이 **{warning_count}건** 발견되었습니다. 해당 품목은 공급처에 성분확인서(비함유증명서)를 별도 요청하세요.")
    else:
        st.success("✅ 모든 파일의 3번 구성성분 및 CAS No. 표가 정상 인식되었습니다.")

    # 5. 테이블 조건부 서식(색상 하이라이트) 함수
    def highlight_status(row):
        val = str(row['판정결과'])
        if '🚨' in val:
            return ['background-color: #ffe6cc; color: #b35900; font-weight: bold;'] * len(row)  # 주황/노랑 (경고)
        elif '⚠️' in val:
            return ['background-color: #ffcccc; color: #cc0000; font-weight: bold;'] * len(row)  # 붉은색 (유해물질)
        elif '정상' in val:
            return ['background-color: #e6ffed; color: #155724;'] * len(row)  # 연초록 (정상)
        return [''] * len(row)

    styled_df = df_results.style.apply(highlight_status, axis=1)
    st.dataframe(styled_df, use_container_width=True)

    # 6. 엑셀 다운로드
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_results.to_excel(writer, index=False, sheet_name='MSDS검토결과')
    
    st.download_button(
        label="📥 결과 엑셀 파일 다운로드",
        data=buffer.getvalue(),
        file_name="MSDS_유해성분_검토결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
