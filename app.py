import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="MSDS 성분 분석기", layout="wide")
st.title("🧪 MSDS 성분/함유량 추출 & 유해물질 판별 도구")
st.write("MSDS 문서의 **3번 구성성분** 단락을 분석하여 물질명, CAS 번호, 함유량을 정확하게 추출합니다.")

CAS_REGEX = r'\b[1-9]\d{1,6}-\d{2}-\d\b'

# 1. 유해물질 DB 등록
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

# 2. 3번 단락 집중 파싱 함수
def extract_from_section_3(pdf_file):
    results = []
    
    with pdfplumber.open(pdf_file) as pdf:
        raw_text_list = []
        for page in pdf.pages:
            t = page.extract_text() or ""
            raw_text_list.append(t)
        full_text = "\n".join(raw_text_list)
        
    has_trade_secret = bool(re.search(r'영업비밀|Trade\s*Secret', full_text, re.I))

    # 3번 구성성분 단락 발췌
    sec3_match = re.search(r'(?:3\s*[\.\,\-\)]\s*구성\s*성분[\s\S]*?)(?=(?:4\s*[\.\,\-\)]\s*응급)|\Z)', full_text, re.I)
    if sec3_match:
        section_text = sec3_match.group(0)
    else:
        sec3_match_bk = re.search(r'(?:구성\s*성분[\s\S]*?)(?=(?:응급\s*(?:조치|처치))|\Z)', full_text, re.I)
        section_text = sec3_match_bk.group(0) if sec3_match_bk else full_text

    # KE 번호 등 다른 식별번호를 임시 마스킹 처리하여 함유량 오인 방지 (예: KE-21487 -> KE_NUM)
    clean_sec_text = re.sub(r'KE-\d+', 'KE_NUM', section_text)
    clean_sec_text = re.sub(r'EC\s*No\.?\s*[\d\-]+', 'EC_NUM', clean_sec_text, flags=re.I)

    cas_matches = list(re.finditer(CAS_REGEX, clean_sec_text))
    if not cas_matches:
        return [], has_trade_secret

    for i, match in enumerate(cas_matches):
        cas_no = match.group(0)
        start_idx = match.start()
        end_idx = match.end()
        
        # 1) 구성성분명 파싱
        prev_end = cas_matches[i-1].end() if i > 0 else 0
        before_text = clean_sec_text[prev_end:start_idx]
        
        before_text_clean = before_text.replace('|', ' ').replace('\n', ' ').replace('/', ' ')
        words = before_text_clean.split()
        
        filtered_words = []
        for w in words:
            if any(h in w for h in ['3.', '구성성분', '명칭', '함유량', '물질명', '이명', '관용명', '화학물질명', 'CAS', '번호', '식별번호', 'KE_NUM', 'EC_NUM', '%']):
                continue
            # 순수 숫자나 범위 표현(이전 항목의 함유량)은 이름에서 제외
            if re.match(r'^[<>]?\s*\d+(\.\d+)?(\s*[\~–\-]\s*\d+(\.\d+)?)?%?$', w):
                continue
            filtered_words.append(w)
            
        comp_name = " ".join(filtered_words) if filtered_words else "-"
        
        # 2) 함유량 파싱 (다음 CAS 번호 전까지의 텍스트 탐색)
        next_start = cas_matches[i+1].start() if i+1 < len(cas_matches) else len(clean_sec_text)
        after_text = clean_sec_text[end_idx:next_start]
        
        # 함유량 정밀 패턴: 15~20, 5-10, 71, 12.5(10~15), < 5, 1 ~ 5 % 등
        cnt_pattern = r'([<>]?\s*\d+(?:\.\d+)?\s*(?:[\~–\-]\s*\d+(?:\.\d+)?)?\s*(?:\(.*?\))?\s*%?)'
        
        content = "-"
        # 줄 단위로 쪼개어 KE_NUM을 건너뛰고 함유량 타겟팅
        for token in after_text.replace('|', ' ').split():
            if 'KE_NUM' in token or 'EC_NUM' in token or token in ['/', '식별번호']:
                continue
            cnt_match = re.search(cnt_pattern, token)
            if cnt_match and any(char.isdigit() for char in token):
                content = token.strip()
                break
                
        # 토큰 단위로 못 찾았을 경우 전체 문자열 패턴 매칭
        if content == "-":
            after_clean = re.sub(r'KE_NUM|EC_NUM|/', ' ', after_text)
            fallback_match = re.search(r'([<>]?\s*\d+(?:\.\d+)?\s*[\~–\-]\s*\d+(?:\.\d+)?%?|\b\d+(?:\.\d+)?%?\b)', after_clean)
            if fallback_match:
                content = fallback_match.group(0).strip()

        results.append({
            "구성성분명": comp_name,
            "CAS No.": cas_no,
            "함유량": content
        })

    # CAS 중복 제거
    unique_results = []
    seen = set()
    for item in results:
        if item["CAS No."] not in seen:
            seen.add(item["CAS No."])
            unique_results.append(item)

    return unique_results, has_trade_secret

# 3. 메인 화면
st.subheader("📄 2. 검토할 MSDS PDF 파일 업로드")
uploaded_pdfs = st.file_uploader("MSDS PDF 파일 업로드 (여러 개 가능)", type=["pdf"], accept_multiple_files=True)

if uploaded_pdfs:
    all_rows = []
    warning_count = 0
    
    with st.spinner("MSDS 3번 단락 정밀 분석 중..."):
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
    
    # 4. 결과 출력
    st.subheader("📊 3. 분석 결과")
    if warning_count > 0:
        st.warning(f"⚠️ **주의**: 3번 단락 성분을 추출하지 못했거나 영업비밀인 파일이 **{warning_count}건** 있습니다.")
    else:
        st.success("✅ 3번 구성성분 분석이 성공적으로 완료되었습니다!")

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

    # 5. 엑셀 다운로드
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_results.to_excel(writer, index=False, sheet_name='MSDS검토결과')
    
    st.download_button(
        label="📥 결과 엑셀 파일 다운로드",
        data=buffer.getvalue(),
        file_name="MSDS_구성성분_추출결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
