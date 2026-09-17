import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="MSDS 성분 분석기", layout="wide")
st.title("🧪 MSDS 성분/함유량 추출 & 유해물질 판별 도구")
st.write("MSDS 문서의 **3번 단락(구성성분)과 4번 단락(응급조치)** 사이 텍스트를 분석하여 성분명, CAS 번호, 함유량을 추출합니다.")

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
            info = row.get('규제구분', row.get('물질명', '유해물질 등록됨'))
            target_dict[cas_val] = str(info)
        st.sidebar.success(f"기준 DB 등록 완료: 총 {len(target_dict)}종 로드됨")
    else:
        st.sidebar.error("엑셀에 'CAS'가 포함된 열 이름이 필요합니다.")

# 2. 단락(3번 ~ 4번 사이) 텍스트 분석 핵심 함수
def extract_from_section_3(pdf_file):
    results = []
    
    with pdfplumber.open(pdf_file) as pdf:
        # 전체 페이지의 텍스트를 하나의 본문으로 합치기
        full_text = "\n".join([(page.extract_text() or "") for page in pdf.pages])
        
    has_trade_secret = bool(re.search(r'영업비밀|Trade\s*Secret', full_text, re.I))

    # [핵심] '3. 구성성분...' 단락 시작부터 '4. 응급...' 단락 시작 직전까지만 추출
    # 3. / 3 / 3- 형태 등 모든 단락 번호 표기 대응
    sec3_pattern = r'(?:3\s*[\.\,\-\)]\s*구성\s*성분[\s\S]*?)(?=4\s*[\.\,\-\)]\s*응급|\Z)'
    sec3_match = re.search(sec3_pattern, full_text, re.I)
    
    # 3번 단락을 찾지 못했을 경우 백업: '구성성분의 명칭' 키워드부터 '응급' 키워드 전까지
    if not sec3_match:
        sec3_pattern_backup = r'(?:구성\s*성분[\s\S]*?)(?=응급\s*(?:조치|처치)|\Z)'
        sec3_match = re.search(sec3_pattern_backup, full_text, re.I)

    if not sec3_match:
        return [], has_trade_secret

    section_text = sec3_match.group(0)

    # 3번 단락 안의 줄(Line)들을 순회하면서 CAS 번호와 성분명, 함유량 수집
    lines = section_text.split('\n')
    for line in lines:
        cas_match = re.search(CAS_REGEX, line)
        if cas_match:
            cas_no = cas_match.group(0)
            
            # 특수문자 및 불필요한 공백 정리
            clean_line = line.replace('|', ' ').replace('\t', ' ')
            
            # CAS 번호를 기준으로 앞부분(물질명 추정)과 뒷부분(함유량 추정) 분리
            parts = clean_line.split(cas_no)
            before = parts[0].strip()
            after = parts[1].strip() if len(parts) > 1 else ""
            
            # 1) 구성성분명 정리 (테이블 헤더 단어 제외)
            name_words = [w for w in before.split() if w not in ['물질명', '이명', '관용명', '화학명', '성분', 'Name', 'CAS', 'CAS번호', 'CASNo']]
            comp_name = " ".join(name_words) if name_words else "-"
            
            # 2) 함유량 정리 (숫자, %, 괄호 범위 패턴 추출)
            cnt = "-"
            # CAS 뒷부분에서 먼저 검색 (예: 71, 12.5(10~15), 5~10% 등)
            cnt_match = re.search(r'(\d+(?:\.\d+)?(?:\s*\(.*?\))?%?|\d+\s*~\s*\d+%)', after)
            if cnt_match:
                cnt = cnt_match.group(0).strip()
            else:
                # 뒷부분에 없으면 앞부분의 끝자리 숫자 검색
                cnt_match2 = re.search(r'(\d+(?:\.\d+)?%?)$', before)
                if cnt_match2:
                    cnt = cnt_match2.group(0).strip()

            results.append({
                "구성성분명": comp_name,
                "CAS No.": cas_no,
                "함유량": cnt
            })

    # 중복 CAS 번호 제거
    unique_results = []
    seen = set()
    for item in results:
        if item["CAS No."] not in seen:
            seen.add(item["CAS No."])
            unique_results.append(item)

    return unique_results, has_trade_secret

# 3. 메인 화면: 파일 업로드 및 결과 도출
st.subheader("📄 2. 검토할 MSDS PDF 파일 업로드")
uploaded_pdfs = st.file_uploader("MSDS PDF 파일 업로드 (여러 개 동시 가능)", type=["pdf"], accept_multiple_files=True)

if uploaded_pdfs:
    all_rows = []
    warning_count = 0
    
    with st.spinner("MSDS 3번 단락 분석 중..."):
        for pdf_file in uploaded_pdfs:
            extracted_items, has_trade_secret = extract_from_section_3(pdf_file)
            
            if not extracted_items:
                warning_count += 1
                status_text = "🚨 영업비밀 표기 의심" if has_trade_secret else "🚨 3번 단락/성분 미발견"
                all_rows.append({
                    "파일명": pdf_file.name,
                    "구성성분명": "성분 미기재 또는 확인 필요",
                    "CAS No.": "-",
                    "함유량": "-",
                    "판정결과": status_text,
                    "상세 규제 정보": "공급처에 성분확인서 징구 요망"
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
    
    # 4. 상단 요약 알림
    st.subheader("📊 3. 분석 결과")
    if warning_count > 0:
        st.warning(f"⚠️ **주의**: 3번 단락에서 성분을 찾지 못했거나 영업비밀로 표기된 파일이 **{warning_count}건** 있습니다.")
    else:
        st.success("✅ 모든 파일의 3번 단락 구성성분 분석이 완료되었습니다.")

    # 5. 시각적 색상 하이라이트
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
        df_results.to_excel(writer, index=False, sheet_name='MSDS단락검토결과')
    
    st.download_button(
        label="📥 결과 엑셀 파일 다운로드",
        data=buffer.getvalue(),
        file_name="MSDS_단락분석_결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
