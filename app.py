import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="MSDS 성분 및 유해물질 자동 분석기", layout="wide")
st.title("🧪 MSDS 성분/함유량 추출 & 유해물질 자동 판별 도구")
st.write("MSDS PDF에서 **구성성분명, CAS No., 함유량(%)**을 자동 추출하고 유해물질 DB와 대조합니다.")

CAS_REGEX = r'\b[1-9]\d{1,6}-\d{2}-\d\b'

# 1. 유해물질 DB 로드
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

# 2. 강력한 성분 추출 함수
def extract_composition_data(pdf_file):
    results = []
    has_trade_secret = False
    
    with pdfplumber.open(pdf_file) as pdf:
        target_page_indices = []
        
        # 1단계: "3. 구성성분" 단어가 있는 페이지 번호 찾기
        for idx, page in enumerate(pdf.pages):
            p_text = page.extract_text() or ""
            # 영업비밀 여부 체크
            if "영업비밀" in p_text or "Trade Secret" in p_text:
                has_trade_secret = True
                
            # 3번 구성성분 관련 키워드가 있는지 확인
            clean_text = p_text.replace(" ", "")
            if re.search(r'3\.(?:구성성분|혼합물의구성|성분의명칭|성분명칭)', clean_text):
                target_page_indices.append(idx)

        # 3번 섹션이 있는 페이지의 텍스트와 표를 집중 분석
        pages_to_scan = [pdf.pages[i] for i in target_page_indices] if target_page_indices else pdf.pages
        
        for p in pages_to_scan:
            # 먼저 해당 페이지의 표(Table) 시도
            tables = p.extract_tables() or []
            for tbl in tables:
                for row in tbl:
                    if not row:
                        continue
                    row_str = " ".join([str(c) for c in row if c])
                    cas_match = re.search(CAS_REGEX, row_str)
                    if cas_match:
                        cas_no = cas_match.group(0)
                        cells = [str(c).strip().replace('\n', ' ') for c in row if c and str(c).strip()]
                        
                        # 성분명 & 함유량 파싱
                        name = "-"
                        cnt = "-"
                        for cell in cells:
                            if cell == cas_no:
                                continue
                            if re.search(r'^\d+(\.\d+)?(\s*\(.*?\))?%?$', cell) or re.search(r'^\d+\s*~\s*\d+%', cell):
                                cnt = cell
                            elif name == "-" and not any(h in cell for h in ['물질명', '이명', 'CAS', '번호', 'No']):
                                name = cell
                        
                        results.append({
                            "구성성분명": name,
                            "CAS No.": cas_no,
                            "함유량": cnt
                        })

            # 표로 안 뽑혔다면 페이지 일반 텍스트 라인 분석 (한일시멘트 대응)
            p_raw = p.extract_text() or ""
            # 3. 구성성분 부터 4. 응급조치 직전까지만 추출
            if "3." in p_raw:
                p_chunk = re.split(r'4\.\s*응급', p_raw)[0]
            else:
                p_chunk = p_raw
                
            for line in p_chunk.split('\n'):
                cas_match = re.search(CAS_REGEX, line)
                if cas_match:
                    cas_no = cas_match.group(0)
                    
                    # 이미 위에서 추출한 CAS면 스킵
                    if any(r["CAS No."] == cas_no for r in results):
                        continue
                        
                    # 텍스트 정제
                    cleaned = line.replace('|', ' ')
                    tokens = [t.strip() for t in cleaned.split() if t.strip()]
                    
                    # CAS 번호 앞쪽 토큰을 물질명으로 취급
                    cas_pos = -1
                    for ti, tok in enumerate(tokens):
                        if cas_no in tok:
                            cas_pos = ti
                            break
                            
                    name = " ".join([tok for tok in tokens[:cas_pos] if tok not in ['물질명', '이명', '관용명', '화학명', 'CAS번호', 'CASNo']]) if cas_pos > 0 else "-"
                    
                    # CAS 뒷쪽 토큰 중 숫자나 %를 함유량으로 취급
                    cnt = "-"
                    if cas_pos != -1 and cas_pos < len(tokens) - 1:
                        after_tokens = tokens[cas_pos + 1:]
                        for at in after_tokens:
                            if re.search(r'\d+', at):
                                cnt = at
                                break
                    
                    results.append({
                        "구성성분명": name if name else "-",
                        "CAS No.": cas_no,
                        "함유량": cnt if cnt else "-"
                    })

    # 중복 제거
    unique_results = []
    seen = set()
    for item in results:
        if item["CAS No."] not in seen:
            seen.add(item["CAS No."])
            unique_results.append(item)
            
    return unique_results, has_trade_secret

# 3. 파일 업로드 및 분석
st.subheader("📄 2. 검토할 MSDS PDF 파일 업로드")
uploaded_pdfs = st.file_uploader("MSDS PDF 파일을 선택하세요 (여러 개 가능)", type=["pdf"], accept_multiple_files=True)

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

    # 6. 다운로드
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_results.to_excel(writer, index=False, sheet_name='MSDS검토결과')
    
    st.download_button(
        label="📥 결과 엑셀 파일 다운로드",
        data=buffer.getvalue(),
        file_name="MSDS_유해성분_검토결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
