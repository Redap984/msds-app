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

# 2. 강력한 MSDS 성분 추출 엔진
def extract_composition_data(pdf_file):
    results = []
    has_trade_secret = False
    
    with pdfplumber.open(pdf_file) as pdf:
        all_pages_text = []
        for p in pdf.pages:
            txt = p.extract_text(layout=False) or ""
            all_pages_text.append(txt)
            
        full_document = "\n---PAGE---\n".join(all_pages_text)

        # 1. 영업비밀 키워드 탐색
        if re.search(r'영업비밀|Trade\s*Secret', full_document, re.I):
            has_trade_secret = True

        # 2. '3. 구성성분...' 섹션만 정확히 잘라내기
        sec3_match = re.search(
            r'(?:3\s*[\.\,\)]\s*(?:구성\s*성분|혼합물의\s*구성|성분의\s*명칭|성분명칭).*?)(?=(?:4\s*[\.\,\)]\s*(?:응급|응급처치|응급조치))|\Z)', 
            full_document, 
            re.S | re.I
        )
        
        target_chunk = sec3_match.group(0) if sec3_match else full_document
        
        # 3. 표 형태든 줄글이든 각 줄별로 정밀 분석
        raw_lines = target_chunk.split('\n')
        
        for idx, line in enumerate(raw_lines):
            line_str = line.strip()
            cas_found = re.search(CAS_REGEX, line_str)
            
            if cas_found:
                cas_no = cas_found.group(0)
                
                # 라인 정리 (| 및 특수문자 제거)
                cleaned_line = line_str.replace('|', ' ')
                
                # CAS 번호 앞부분 -> 물질명 추정
                parts = cleaned_line.split(cas_no)
                before_cas = parts[0].strip()
                after_cas = parts[1].strip() if len(parts) > 1 else ""
                
                # 성분명 추출 (헤더 단어 제외 필터링)
                words_before = [w for w in before_cas.split() if w not in ['물질명', '이명', '관용명', '화학명', '성분', 'Name', 'CAS', 'CAS번호', 'CASNo']]
                m_name = " ".join(words_before) if words_before else "-"
                
                # 만약 같은 줄 앞쪽에 성분명이 없고 윗줄에 성분명이 적혀 있는 경우
                if (not m_name or m_name == "-") and idx > 0:
                    prev_line = raw_lines[idx-1].replace('|', ' ').strip()
                    prev_words = [w for w in prev_line.split() if not re.search(CAS_REGEX, w) and w not in ['물질명', '이명', '관용명', '구성성분']]
                    if prev_words:
                        m_name = " ".join(prev_words)

                # 함유량 추출 (% 또는 숫자 패턴)
                content = "-"
                # 1순위: CAS 뒷부분에서 탐색
                cnt_match = re.search(r'(\d+(?:\.\d+)?(?:\s*\(.*?\))?%?|\d+\s*~\s*\d+%)', after_cas)
                if cnt_match:
                    content = cnt_match.group(0).strip()
                else:
                    # 2순위: 전체 줄에서 CAS 제외 후 탐색
                    rem = cleaned_line.replace(cas_no, '').replace(m_name, '')
                    c_match2 = re.search(r'(\d+(?:\.\d+)?(?:\s*\(.*?\))?%?|\d+\s*~\s*\d+%)', rem)
                    if c_match2:
                        content = c_match2.group(0).strip()

                results.append({
                    "구성성분명": m_name if m_name else "-",
                    "CAS No.": cas_no,
                    "함유량": content if content else "-"
                })

    # 중복 제거 (등장 순서 보존)
    unique_results = []
    seen_cas = set()
    for item in results:
        if item["CAS No."] not in seen_cas:
            seen_cas.add(item["CAS No."])
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
    
    # 4. 상단 알림
    st.subheader("📊 3. 분석 결과")
    if warning_count > 0:
        st.warning(f"⚠️ **주의**: 3번 성분표가 누락되었거나 영업비밀로 표기된 파일이 **{warning_count}건** 있습니다.")
    else:
        st.success("✅ 모든 파일의 3번 구성성분 및 CAS No. 표가 정상 분석되었습니다.")

    # 5. 색상 하이라이트
    def highlight_status(row):
        val = str(row['판정결과'])
        if '🚨' in val:
            return ['background-color: #ffe6cc; color: #b35900; font-weight: bold;'] * len(row)
        elif '⚠️' in val:
            return ['background-color: #ffcccc; color: #cc0000; font-weight: bold;'] * len(row)
        elif '정상' in val:
            return ['background-color: #e6ffed; color: #155724;'] * len(row)
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
