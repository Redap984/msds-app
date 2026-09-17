import streamlit as st
import pdfplumber
import pandas as pd
import json
import re
import io
import google.generativeai as genai

st.set_page_config(page_title="AI 기반 MSDS 성분 및 유해물질 분석기", layout="wide")
st.title("🧪 AI MSDS 성분 추출 & 법정 유해물질 자동 판별 도구")
st.write("문맥 이해 AI가 양식에 상관없이 **구성성분명, CAS No., 함유량(%)**을 완벽하게 추출하고 유해물질 여부를 판별합니다.")

# -------------------------------------------------------------
# 1. 대한민국 법정 규제 유해화학물질 기본 내장 DB (확장 가능)
# -------------------------------------------------------------
DEFAULT_HAZARDOUS_DB = {
    # 화관법 사고대비물질 / 유독물질 / 산안법 특별관리물질 대표 목록
    "67-56-1": "화관법(유독물질, 사고대비물질), 산안법(관리대상, 특별관리물질)", # 메탄올
    "7664-93-9": "화관법(유독물질, 사고대비물질), 산안법(특별관리물질)",        # 황산
    "7697-37-2": "화관법(유독물질, 사고대비물질), 산안법(관리대상유해물질)",      # 질산
    "7647-01-0": "화관법(유독물질, 사고대비물질), 산안법(관리대상유해물질)",      # 염화수소(염산)
    "7664-39-3": "화관법(유독물질, 사고대비물질), 산안법(특별관리물질)",        # 불화수소(불산)
    "50-00-0": "화관법(유독물질, 사고대비물질), 산안법(특별관리물질, 발암성1A)",  # 포름알데히드
    "71-43-2": "화관법(유독물질, 사고대비물질), 산안법(특별관리물질, 발암성1A)",    # 벤젠
    "108-88-3": "화관법(유독물질, 사고대비물질), 산안법(관리대상유해물질)",       # 톨루엔
    "1330-20-7": "화관법(유독물질), 산안법(관리대상유해물질)",                  # 자일렌
    "75-09-2": "화관법(유독물질), 산안법(특별관리물질)",                       # 디클로로메탄
    "79-01-6": "화관법(유독물질, 사고대비물질), 산안법(특별관리물질)",           # 트리클로로에틸렌(TCE)
    "127-18-4": "화관법(유독물질, 사고대비물질), 산안법(특별관리물질)",          # 테트라클로로에틸렌(PCE)
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
}

# -------------------------------------------------------------
# 2. 사이드바 설정 (API 키 & 추가 DB 업로드)
# -------------------------------------------------------------
st.sidebar.header("🔑 Gemini API 설정")
api_key = st.sidebar.text_input("Gemini API Key를 입력하세요", type="password", help="Google AI Studio에서 무료로 즉시 발급 가능합니다.")

st.sidebar.header("📁 사용자 추가 DB (선택사항)")
user_db_file = st.sidebar.file_uploader("사내 관리 엑셀 추가 등록 (.xlsx)", type=["xlsx", "xls"])

target_dict = DEFAULT_HAZARDOUS_DB.copy()
if user_db_file:
    df_user = pd.read_excel(user_db_file)
    cas_cols = [c for c in df_user.columns if 'CAS' in str(c).upper()]
    if cas_cols:
        c_name = cas_cols[0]
        for _, row in df_user.iterrows():
            cas_val = str(row[c_name]).strip()
            info = row.get('규제구분', row.get('물질명', '사내 유해물질 등록'))
            target_dict[cas_val] = f"[사내DB] {info}"
        st.sidebar.success(f"사내 DB 통합 완료 (총 {len(target_dict)}종)")

# -------------------------------------------------------------
# 3. AI 문맥 분석 엔진 (LLM 기반 추출)
# -------------------------------------------------------------
def analyze_msds_with_gemini(text_chunk, key):
    genai.configure(api_key=key)
    model = genai.GenerativeModel("gemini-1.5-flash")
    
    prompt = f"""
당신은 화학물질 안전관리 전문가입니다. 아래 제공된 MSDS 텍스트에서 '구성성분의 명칭 및 함유량' 정보를 정확히 추출하세요.

[추출 규칙]
1. 물질명(화학명/관용명), CAS 번호, 함유량(%)을 각각 분리하세요.
2. 제형/화학식(예: C14H10O4, SiO2 등)이나 고유식별번호(KE-xxxxx, EC No 등)는 함유량이 아니므로 절대 함유량에 넣지 마세요.
3. 함유량은 10-15%, 15~20, 71, 1~5 등 순수한 백분율 수치나 범위만 추출하세요.
4. 만약 영업비밀(Trade Secret)로 기재되어 CAS 번호가 없다면 CAS No에 '영업비밀'로 표기하세요.
5. 반드시 아래 JSON 배열 형식으로만 응답하세요. 다른 설명 문구는 일절 넣지 마세요.

JSON 응답 형식:
[
  {{"name": "물질명", "cas": "CAS번호", "content": "함유량"}},
  ...
]

[MSDS 텍스트]:
{text_chunk}
"""
    response = model.generate_content(prompt)
    clean_resp = re.sub(r'```json|```', '', response.text).strip()
    return json.loads(clean_resp)

# -------------------------------------------------------------
# 4. 파일 업로드 및 실행
# -------------------------------------------------------------
st.subheader("📄 검토할 MSDS PDF 업로드")
uploaded_pdfs = st.file_uploader("하나 이상의 MSDS PDF를 선택하세요", type=["pdf"], accept_multiple_files=True)

if uploaded_pdfs:
    if not api_key:
        st.error("👈 왼쪽 사이드바에 Gemini API Key를 먼저 입력해 주세요. (Google AI Studio에서 무료로 10초 만에 발급받으실 수 있습니다.)")
    else:
        all_rows = []
        warning_count = 0
        
        with st.spinner("AI가 문서 문맥을 읽고 성분 및 함유량을 정밀 분석 중입니다..."):
            for pdf_file in uploaded_pdfs:
                with pdfplumber.open(pdf_file) as pdf:
                    # 1~5페이지 위주 텍스트 수집 (구성성분 섹션 포함 구간)
                    pages_text = [p.extract_text() or "" for p in pdf.pages[:min(6, len(pdf.pages))]]
                    full_chunk = "\n".join(pages_text)
                    
                try:
                    parsed_items = analyze_msds_with_gemini(full_chunk, api_key)
                    
                    if not parsed_items:
                        warning_count += 1
                        all_rows.append({
                            "파일명": pdf_file.name,
                            "구성성분명": "성분 정보 미검출",
                            "CAS No.": "-",
                            "함유량": "-",
                            "판정결과": "🚨 확인 필요",
                            "상세 규제 정보": "공급처 성분 확인 필요"
                        })
                    else:
                        for item in parsed_items:
                            cas = str(item.get("cas", "-")).strip()
                            name = item.get("name", "-")
                            content = item.get("content", "-")
                            
                            is_hazard = cas in target_dict
                            hazard_info = target_dict.get(cas, "-")
                            
                            if "영업비밀" in cas:
                                warning_count += 1
                                verdict = "🚨 영업비밀 (성분 미기재)"
                                hazard_info = "공급처 비함유확인서(비유해성확인서) 징구 요망"
                            elif is_hazard:
                                verdict = "⚠️ 유해물질 해당"
                            else:
                                verdict = "정상 (미해당)"
                                
                            all_rows.append({
                                "파일명": pdf_file.name,
                                "구성성분명": name,
                                "CAS No.": cas,
                                "함유량": content,
                                "판정결과": verdict,
                                "상세 규제 정보": hazard_info
                            })
                except Exception as e:
                    warning_count += 1
                    all_rows.append({
                        "파일명": pdf_file.name,
                        "구성성분명": f"분석 오류: {str(e)[:30]}",
                        "CAS No.": "-",
                        "함유량": "-",
                        "판정결과": "🚨 오류 발생",
                        "상세 규제 정보": "파일 포맷 점검 요망"
                    })

        df_results = pd.DataFrame(all_rows)
        
        # 결과 표시
        st.subheader("📊 분석 결과")
        if warning_count > 0:
            st.warning(f"⚠️ 확인이 필요한 항목(유해물질, 영업비밀 등)이 {warning_count}건 감지되었습니다.")
        else:
            st.success("✅ 모든 성분이 정상 판별되었습니다.")

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
            label="📥 결과 엑셀 파일 다운로드",
            data=buffer.getvalue(),
            file_name="MSDS_AI검토결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
