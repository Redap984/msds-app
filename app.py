import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="MSDS 성분/CAS 자동 분석기", layout="wide")
st.title("🧪 MSDS 성분 & CAS 번호 정밀 판별 도구")
st.write("외부 API 호출 없이, 국내외 제조사 MSDS 양식 전체를 대응하는 로컬 정밀 파서입니다.")

CAS_REGEX = r'\b[1-9]\d{1,6}-\d{2}-\d\b'

# 1. 법정 유해물질 기본 내장 DB
DEFAULT_HAZARDOUS_DB = {
    "67-56-1": "화관법(유독물질, 사고대비물질), 산안법(관리대상, 특별관리물질)", # 메탄올
    "7664-93-9": "화관법(유독물질, 사고대비물질), 산안법(특별관리물질)",        # 황산
    "7697-37-2": "화관법(유독물질, 사고대비물질), 산안법(관리대상유해물질)",      # 질산
    "7647-01-0": "화관법(유독물질, 사고대비물질), 산안법(관리대상유해물질)",      # 염산
    "7664-39-3": "화관법(유독물질, 사고대비물질), 산안법(특별관리물질)",        # 불산
    "50-00-0": "화관법(유독물질, 사고대비물질), 산안법(특별관리물질, 발암성1A)",  # 포름알데히드
    "71-43-2": "화관법(유독물질, 사고대비물질), 산안법(특별관리물질, 발암성1A)",    # 벤젠
    "108-88-3": "화관법(유독물질, 사고대비물질), 산안법(관리대상유해물질)",       # 톨루엔
    "1330-20-7": "화관법(유독물질), 산안법(관리대상유해물질)",                  # 자일렌
    "75-09-2": "화관법(유독물질), 산안법(특별관리물질)",                       # 디클로로메탄
    "79-01-6": "화관법(유독물질, 사고대비물질), 산안법(특별관리물질)",           # 트리클로로에틸렌
    "127-18-4": "화관법(유독물질, 사고대비물질), 산안법(특별관리물질)",          # 테트라클로로에틸렌
    "94-36-0": "화관법(유독물질), 위험물안전관리법(제5류 유기과산화물)",          # 과산화 벤조일
    "7439-96-5": "산안법(관리대상유해물질, 특수건강진단대상)",                  # 망간
    "13463-67-7": "산안법(관리대상유해물질, 발암성 2)",                       # 이산화티타늄
    "14808-60-7": "산안법(관리대상유해물질, 발암성 1A)",                      # 산화규소(결정체 석영)
    "1344-28-1": "산안법(관리대상유해물질, 노출기준설정물질)",                 # 산화 알루미늄
    "38668-48-3": "화관법(유독물질)",                                        # 1,1'-(p-톨릴이미노)디프로판-2-올
    "9016-87-9": "산안법(관리대상유해물질, 노출기준설정)",                     # MDI 중합체
    "115-10-6": "고압가스안전관리법(가연성가스)",                              # 다이메틸에테르
    "74-98-6": "고압가스안전관리법(가연성가스)",                               # 프로페인
    "75-28-5": "고압가스안전관리법(가연성가스)",                               # 아이소뷰태인
    "7439-89-6": "자율관리(철)",                                            # 철
}

# 2. 사이드바: 추가 DB 등록
st.sidebar.header("📁 사내 유해물질 기준 DB (선택)")
db_file = st.sidebar.file_uploader("사내 엑셀 DB (.xlsx, .xls)", type=["xlsx", "xls"])

target_dict = DEFAULT_HAZARDOUS_DB.copy()
if db_file:
    df_db = pd.read_excel(db_file)
    cas_col = [c for c in df_db.columns if 'CAS' in str(c).upper()]
    if cas_col:
        c_name = cas_col[0]
        for _, row in df_db.iterrows():
            cas_val = str(row[c_name]).strip()
            info = row.get('규제구분', row.get('물질명', '사내 유해물질 등록됨'))
            target_dict[cas_val] = f"[사내] {info}"
        st.sidebar.success(f"기준 DB 등록 완료: 총 {len(target_dict)}종 로드됨")

# 화학식(C14H10O4, SiO2 등) 여부 검사
def is_chemical_formula(val):
    v = val.strip()
    if re.search(r'[A-Za-z]', v) and any(c.isdigit() for c in v):
        if not re.search(r'^\d+[\~–\-]\d+%$', v):
            return True
    return False

# 3. 전수 케이스 대응 파서 함수
def extract_clean_components(pdf_file):
    results = []
    has_trade_secret = False
    
    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join([(p.extract_text() or "") for p in pdf.pages])

    if re.search(r'영업비밀|Trade\s*Secret', full_text, re.I):
        has_trade_secret = True

    # 식별번호/노이즈 사전 마스킹
    clean_text = re.sub(r'KE-\d+', ' ', full_text)
    clean_text = re.sub(r'EC\s*No\.?\s*[\d\-]+', ' ', clean_text, flags=re.I)
    clean_text = re.sub(r'CAS\s*Number\s*:\s*', '', clean_text, flags=re.I)

    # 3번 단락들 전수 발췌 (복수 키트 포함)
    sec3_regex = r'(?:3(?:\.|\s*항목|\s*[\,\-\)])\s*(?:구성\s*성분|혼합물의\s*구성|킷\s*내용)[\s\S]*?)(?=(?:4(?:\.|\s*항목|\s*[\,\-\)])\s*응급|4\s*항목|\Z)'
    sec3_chunks = [m.group(0) for m in re.finditer(sec3_regex, clean_text, re.I)]
    
    if not sec3_chunks:
        # 단락 번호가 뭉개진 경우의 백업
        sec3_fallback = re.search(r'(?:구성\s*성분[\s\S]*?)(?=응급\s*(?:조치|처치)|\Z)', clean_text, re.I)
        sec3_chunks = [sec3_fallback.group(0)] if sec3_fallback else [clean_text]

    for chunk in sec3_chunks:
        cas_matches = list(re.finditer(CAS_REGEX, chunk))
        if not cas_matches:
            continue

        for i, match in enumerate(cas_matches):
            cas_no = match.group(0)
            start_idx = match.start()
            end_idx = match.end()

            # --- A. 물질명 추출 ---
            prev_end = cas_matches[i-1].end() if i > 0 else 0
            before_str = chunk[prev_end:start_idx].replace('|', ' ').replace('/', ' ')
            words = before_str.split()
            
            clean_words = []
            for w in words:
                if any(h in w for h in ['3.', '3항목', '구성성분', '명칭', '함유량', '물질명', '이명', '화학물질명', '식별번호', '제형', '%', '혼합물']):
                    continue
                # 이전 물질의 함유량 수치 제외
                if re.match(r'^[<>]?\s*\d+(\.\d+)?(\s*[\~–\-]\s*\d+(\.\d+)?)?%?$', w):
                    continue
                clean_words.append(w)
            
            comp_name = " ".join(clean_words) if clean_words else "-"

            # --- B. 함유량 추출 ---
            next_start = cas_matches[i+1].start() if i+1 < len(cas_matches) else len(chunk)
            after_str = chunk[end_idx:next_start].replace('|', ' ')
            tokens = after_str.split()

            content = "-"
            for tok in tokens:
                t = tok.strip()
                if is_chemical_formula(t):
                    continue
                if t in ['/', '유해화학물질', '번호:-', '-']:
                    continue
                
                # 순수 백분율/범위 표현 포착 (예: 10-15%, 15~20, 12.5(10~15), 71 등)
                cnt_m = re.search(r'([<>]?\s*\d+(?:\.\d+)?\s*(?:[\~–\-]\s*\d+(?:\.\d+)?)?(?:\(.*?\))?\s*%?)', t)
                if cnt_m and any(c.isdigit() for c in t):
                    content = cnt_m.group(0).strip()
                    break

            results.append({
                "구성성분명": comp_name,
                "CAS No.": cas_no,
                "함유량": content
            })

    # 중복 CAS 제거
    unique_items = []
    seen = set()
    for item in results:
        if item["CAS No."] not in seen:
            seen.add(item["CAS No."])
            unique_items.append(item)

    return unique_items, has_trade_secret

# 4. 화면 업로드 및 분석
st.subheader("📄 검토할 MSDS PDF 업로드")
uploaded_pdfs = st.file_uploader("MSDS PDF 파일들을 드래그해서 올려주세요 (다중 가능)", type=["pdf"], accept_multiple_files=True)

if uploaded_pdfs:
    all_rows = []
    warning_count = 0

    with st.spinner("모든 제조사 포맷 규칙을 적용하여 분석 중입니다..."):
        for pdf_file in uploaded_pdfs:
            extracted_items, has_trade_secret = extract_clean_components(pdf_file)

            if not extracted_items:
                warning_count += 1
                status = "🚨 영업비밀 표기 의심" if has_trade_secret else "🚨 3번 성분 미검출"
                all_rows.append({
                    "파일명": pdf_file.name,
                    "구성성분명": "공급사 성분 미기재",
                    "CAS No.": "-",
                    "함유량": "-",
                    "판정결과": status,
                    "상세 규제 정보": "공급처 비함유확인서 확인 필요"
                })
            else:
                for item in extracted_items:
                    cas = item["CAS No."]
                    is_hazard = cas in target_dict
                    hazard_info = target_dict.get(cas, "-")

                    verdict = "⚠️ 유해물질 해당" if is_hazard else "정상 (미해당)"

                    all_rows.append({
                        "파일명": pdf_file.name,
                        "구성성분명": item["구성성분명"],
                        "CAS No.": cas,
                        "함유량": item["함유량"],
                        "판정결과": verdict,
                        "상세 규제 정보": hazard_info
                    })

    df_results = pd.DataFrame(all_rows)

    st.subheader("📊 분석 결과")
    if warning_count > 0:
        st.warning(f"⚠️ 확인이 필요한 파일(성분 미검출 또는 영업비밀)이 {warning_count}건 있습니다.")
    else:
        st.success("✅ 모든 파일의 3번 구성성분 및 CAS 번호가 정상 추출되었습니다.")

    def highlight_status(row):
        val = str(row['판정결과'])
        if '🚨' in val:
            return ['background-color: #ffe6cc; color: #b35900; font-weight: bold;'] * len(row)
        elif '⚠️' in val:
            return ['background-color: #ffcccc; color: #cc0000; font-weight: bold;'] * len(row)
        elif '정상' in val:
            return ['background-color: #e6ffed; color: #155724;'] * len(row)
        return [''] * len(row)

    st.dataframe(df_results.style.apply(highlight_status, axis=1), use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_results.to_excel(writer, index=False, sheet_name='MSDS검토결과')

    st.download_button(
        label="📥 결과 엑셀 다운로드",
        data=buffer.getvalue(),
        file_name="MSDS_분석결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
