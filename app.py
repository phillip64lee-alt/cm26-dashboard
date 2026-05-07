import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import json

# ==========================================
# 1. 설정 및 기본 UI 구성
import httplib2
import requests
import subprocess
import tempfile
import os
import re

st.set_page_config(layout="wide", page_title="통합 E-Commerce 대시보드")

st.title("📊 통합 E-Commerce 대시보드")
st.markdown("""
여러 구글 스프레드시트에 흩어진 **키워드 순위, 플랫폼별 통계, 매출 데이터**를 한 화면에서 분석합니다.
(💡 *현재 안전을 위해 원본 데이터 조작이 불가능한 **읽기 전용 모드**로 실행 중입니다.*)
""")

# ==========================================
# 1. Google Sheets API 연동 설정 (배포/로컬 호환)
# ==========================================
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

try:
    # Streamlit Secrets에서 먼저 찾기 시도 (배포용)
    has_secrets = False
    try:
        if "gcp_service_account" in st.secrets:
            has_secrets = True
    except Exception:
        pass # 로컬 환경 등 secrets.toml이 없는 경우 예외 무시

    if has_secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    else:
        # 로컬 환경 (로컬 파일 사용)
        creds = ServiceAccountCredentials.from_json_keyfile_name('C:/service_account_cm26.json', scope)
        
    client = gspread.authorize(creds)
except Exception as e:
    st.error(f"인증 오류가 발생했습니다. 서비스 계정 설정을 확인해주세요: {e}")
    st.stop()

# ==========================================
# 2. 문서 및 시트 설정
# ==========================================
SHEET_CONFIG = [
    {"name": "스스_주요키워드 순위", "doc_id": "1jVx2AkrDNLRWlHwDr5eeSZOygPvKe8sMxgd-55NA7Vo", "gid": "1885188752"},
    {"name": "씨엠_매출기록 (스스_매출자료)", "doc_id": "11z05hilB6I7kkiqj_qoUFFf_BXYLalr3lj4eMZPxnEA", "gid": "414495420"},
    {"name": "씨엠_매출기록 (스스_합계)", "doc_id": "11z05hilB6I7kkiqj_qoUFFf_BXYLalr3lj4eMZPxnEA", "gid": "530048172"},
    {"name": "씨엠_매출기록 (오집_합계)", "doc_id": "11z05hilB6I7kkiqj_qoUFFf_BXYLalr3lj4eMZPxnEA", "gid": "21311813"},
    {"name": "씨엠_매출기록 (오집_일별_주별_월별_매출)", "doc_id": "11z05hilB6I7kkiqj_qoUFFf_BXYLalr3lj4eMZPxnEA", "gid": "1564635243"},
    {"name": "스마트스토어 공구 진행 일정 공유", "doc_id": "199efQ27Ih28urfz4wp2GwCkKXcSPzRLP74ZPpPgLD40", "gid": "218218718"},
    {"name": "문서 3 (권한 연동 중..)", "doc_id": "1hNhCe86wsZOnIH3fwTZ3QZn4OU3grbZBSe8CKFwl6EU", "gid": "772653874"},
    {"name": "이정호팀장_WORK 오늘의집 매출데이터", "doc_id": "1V-pcrRmDYKMdmYeV5kVp9O23RyrUl1j0eDFqoqWYdJs", "gid": "0"},
]

@st.cache_data(ttl=60)
def load_sheet_data(doc_id, gid):
    try:
        ss = client.open_by_key(doc_id)
        ws = next((s for s in ss.worksheets() if str(s.id) == str(gid)), None)
        if ws:
            data = ws.get_all_values()
            if data:
                if creds.access_token_expired or not creds.access_token:
                    creds.refresh(httplib2.Http())
                
                url = f'https://sheets.googleapis.com/v4/spreadsheets/{doc_id}?includeGridData=false&fields=sheets(properties/sheetId,data(columnMetadata/hiddenByUser,rowMetadata/hiddenByUser,columnMetadata/hiddenByFilter,rowMetadata/hiddenByFilter))'
                headers = {'Authorization': 'Bearer ' + creds.access_token}
                res = requests.get(url, headers=headers).json()
                
                hidden_cols = set()
                hidden_rows = set()
                
                sheet_meta = next((s for s in res.get('sheets', []) if str(s.get('properties', {}).get('sheetId')) == str(gid)), None)
                if sheet_meta and 'data' in sheet_meta and sheet_meta['data']:
                    d = sheet_meta['data'][0]
                    for i, c in enumerate(d.get('columnMetadata', [])):
                        if c.get('hiddenByUser') or c.get('hiddenByFilter'):
                            hidden_cols.add(i)
                    for i, r in enumerate(d.get('rowMetadata', [])):
                        if r.get('hiddenByUser') or r.get('hiddenByFilter'):
                            hidden_rows.add(i)
                
                filtered_data = []
                for r_idx, row in enumerate(data):
                    if r_idx in hidden_rows:
                        continue
                    filtered_row = [val for c_idx, val in enumerate(row) if c_idx not in hidden_cols]
                    filtered_data.append(filtered_row)

                if not filtered_data:
                    return pd.DataFrame()

                raw_cols = filtered_data[0]
                cols = []
                seen = {}
                for c in raw_cols:
                    c_str = str(c).strip()
                    if c_str in seen:
                        seen[c_str] += 1
                        cols.append(f"{c_str}_{seen[c_str]}")
                    else:
                        seen[c_str] = 0
                        cols.append(c_str)
                df = pd.DataFrame(filtered_data[1:], columns=cols)
                
                df.replace(r'^\s*$', pd.NA, regex=True, inplace=True)
                df.dropna(axis=0, how='all', inplace=True)
                df.dropna(axis=1, how='all', inplace=True)
                df.fillna("", inplace=True)
                
                return df
        return pd.DataFrame()
    except Exception as e:
        if "PermissionError" in str(repr(e)):
            st.error("🔒 이 문서에 접근할 권한이 없습니다.")
        else:
            st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {e}")
        return pd.DataFrame()

# ==========================================
# 3. 사이드바 및 메인 화면
# ==========================================
st.sidebar.header("📁 데이터 소스 선택")
selected_sheet = st.sidebar.radio("조회할 문서를 선택하세요:", [sheet["name"] for sheet in SHEET_CONFIG])

st.sidebar.markdown("---")
if st.sidebar.button("🔄 데이터 새로고침"):
    st.cache_data.clear()
    st.rerun()

current_config = next((s for s in SHEET_CONFIG if s["name"] == selected_sheet), None)

if current_config:
    st.subheader(f"조회 중: {current_config['name']}")
    
    with st.spinner("데이터를 불러오는 중..."):
        df = load_sheet_data(current_config["doc_id"], current_config["gid"])
        
    if not df.empty:
        import shutil
        has_nlm = shutil.which("nlm") is not None
        
        if has_nlm:
            tabs = st.tabs(["📊 데이터 대시보드", "🧠 AI 인사이트 브리핑"])
            tab_data, tab_ai = tabs[0], tabs[1]
        else:
            tabs = st.tabs(["📊 데이터 대시보드"])
            tab_data = tabs[0]
            tab_ai = None
        
        with tab_data:
            if "오집_일별_주별_월별_매출" in selected_sheet and "날짜" in df.columns and "일 매출" in df.columns:
                st.markdown("### 📈 매출 추이 요약")
                try:
                    chart_df = df.copy()
                    chart_df = chart_df[chart_df["일 매출"] != ""]
                    chart_df["일 매출"] = chart_df["일 매출"].astype(str).str.replace(",", "").apply(pd.to_numeric, errors='coerce')
                    st.line_chart(chart_df.set_index("날짜")["일 매출"], use_container_width=True)
                except Exception as e:
                    st.warning("차트를 생성하는 중 오류가 발생했습니다. 데이터 형식을 확인해주세요.")
                    
            elif "스스_매출자료" in selected_sheet:
                st.markdown("### 📈 브랜드별 일간 매출 추이 (코드리빙 vs 코드26)")
                try:
                    chart_df = df.copy()
                    header_row = chart_df.iloc[0].values
                    date_col = chart_df.columns[list(header_row).index("날짜")] if "날짜" in list(header_row) else chart_df.columns[0]
                    sales_indices = [i for i, x in enumerate(header_row) if x == "일간 매출"]
                    
                    if len(sales_indices) >= 2:
                        cliving_col = chart_df.columns[sales_indices[0]]
                        c26_col = chart_df.columns[sales_indices[1]]
                        
                        chart_df = chart_df.iloc[1:]
                        chart_df = chart_df[chart_df[date_col].notna() & (chart_df[date_col] != "")]
                        
                        chart_df["코드리빙 매출"] = chart_df[cliving_col].astype(str).str.replace(r"[₩,]", "", regex=True).apply(pd.to_numeric, errors='coerce')
                        chart_df["코드26 매출"] = chart_df[c26_col].astype(str).str.replace(r"[₩,]", "", regex=True).apply(pd.to_numeric, errors='coerce')
                        
                        final_chart_df = chart_df.set_index(date_col)[["코드리빙 매출", "코드26 매출"]]
                        final_chart_df.index.name = "날짜"
                        st.line_chart(final_chart_df, use_container_width=True)
                except Exception as e:
                    st.warning(f"차트를 생성하는 중 오류가 발생했습니다: {e}")

            st.markdown("### 📝 전체 데이터표 (틀고정 적용)")
            if not df.empty and len(df.columns) > 0:
                display_df = df.set_index(df.columns[0])
                st.dataframe(display_df, use_container_width=True)
            else:
                st.dataframe(df, use_container_width=True)
            
        if tab_ai:
            with tab_ai:
                st.markdown("### 🤖 NotebookLM 기반 데이터 인사이트 분석")
                st.info("현재 조회 중인 데이터를 구글의 **NotebookLM AI**가 분석하여 주요 비즈니스 인사이트를 제공합니다.")
                
                user_query = st.text_area(
                    "💡 AI에게 물어볼 질문을 자유롭게 입력하세요:",
                    value="첨부된 데이터의 핵심적인 매출 또는 데이터 트렌드 특징 3가지를 분석하고, 실질적인 비즈니스 액션 아이디어를 제시해줘. 가독성 있게 마크다운으로 정리해줘.",
                    height=100
                )
                
                if st.button("✨ AI 인사이트 브리핑 생성하기", key="generate_insight"):
                    with st.spinner("NotebookLM AI 서버에 데이터를 전송하고 심층 분석을 수행 중입니다... (최대 1~2분 소요)"):
                        import subprocess
                        import tempfile
                        import os
                        import re
                        
                        temp_path = ""
                        try:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".csv", mode='w', encoding='utf-8') as f:
                                df.to_csv(f.name, index=False)
                                temp_path = f.name
                                
                            nb_name = f"Dashboard_{str(current_config['name']).replace(' ', '_')}"
                            create_res = subprocess.run(["nlm", "create", "notebook", nb_name], capture_output=True, text=True, encoding='utf-8')
                            
                            stdout_str = str(create_res.stdout) if create_res.stdout else ""
                            stderr_str = str(create_res.stderr) if create_res.stderr else ""
                            
                            if create_res.returncode != 0 or "Authentication expired" in stderr_str:
                                st.error("⚠️ NotebookLM 인증이 만료되었거나 연동에 실패했습니다.")
                                st.markdown("터미널 창을 열고 `nlm login` 명령어를 실행하여 구글 계정을 다시 연동해주세요.")
                            else:
                                uuid_match = re.search(r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}', stdout_str)
                                if not uuid_match:
                                    st.error("노트북 ID를 가져오지 못했습니다.")
                                else:
                                    nb_id = uuid_match.group(0)
                                    st.write("✅ NotebookLM 임시 분석 공간이 생성되었습니다. (데이터 업로드 중...)")
                                    
                                    src_res = subprocess.run(["nlm", "source", "add", nb_id, "--file", temp_path, "--wait"], capture_output=True, text=True, encoding='utf-8')
                                    if src_res.returncode != 0:
                                        st.error("데이터 업로드 중 오류가 발생했습니다.")
                                    else:
                                        st.write("✅ 데이터 업로드 및 인덱싱이 완료되었습니다. (분석 중...)")
                                        
                                        query = user_query
                                        query_res = subprocess.run(["nlm", "query", "notebook", nb_id, query], capture_output=True, text=True, encoding='utf-8')
                                        
                                        query_stdout = str(query_res.stdout) if query_res.stdout else ""
                                        
                                        if query_res.returncode != 0:
                                            st.error("AI 인사이트 도출 중 오류가 발생했습니다.")
                                        else:
                                            st.success("🎉 분석 완료!")
                                            st.markdown("---")
                                            
                                            # JSON 형태로 반환되는 응답에서 answer 부분만 추출
                                            try:
                                                import json
                                                parsed_res = json.loads(query_stdout)
                                                # nlm cli 응답 구조에 맞춰 파싱
                                                if "value" in parsed_res and "answer" in parsed_res["value"]:
                                                    final_answer = parsed_res["value"]["answer"]
                                                else:
                                                    final_answer = query_stdout
                                            except Exception:
                                                # JSON 파싱 실패 시 원본 출력
                                                final_answer = query_stdout
                                                
                                            st.markdown(final_answer)
                                            
                                            subprocess.run(["nlm", "delete", "notebook", nb_id, "--confirm"], capture_output=True, text=True, encoding='utf-8')
                        except Exception as e:
                            st.error(f"오류가 발생했습니다: {e}")
                        finally:
                            if os.path.exists(temp_path):
                                os.remove(temp_path)

    else:
        st.warning("데이터가 비어있거나 권한 문제로 불러올 수 없습니다.")

st.sidebar.markdown("---")
st.sidebar.info("💡 **알림**: 'NotebookLM AI 인사이트 브리핑' 기능이 통합되었습니다! 메인 화면의 탭을 확인해보세요. 향후 '데이터 에디팅(수정)' 기능도 추가될 예정입니다.")
