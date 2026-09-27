import glob
import os
import re
from datetime import datetime
import pandas as pd
import streamlit as st

# 페이지 설정 (라이트 모드 고정 및 깔끔한 폭 조절)
st.set_page_config(page_title="인하대병원 의료장비 예방점검 라벨", layout="centered")

# ==========================================
# 📂 데이터 파일 로드 함수
# ==========================================
def find_latest_status_file(prefix="의료기기 현황조회", extensions=(".xlsx", ".xlsb", ".xls")):
    files = []
    for ext in extensions:
        files.extend(glob.glob(f"{prefix}*{ext}"))
    if not files:
        return None
    def extract_date(filename):
        remainder = filename[len(prefix):]
        matches = re.findall(r'(\d+)', remainder)
        if matches:
            return matches[-1]
        return ""
    files.sort(key=extract_date, reverse=True)
    return files[0]

@st.cache_data
def load_latest_data():
    status_file = find_latest_status_file("의료기기 현황조회")
    repair_files = []
    for ext in (".xlsx", ".xlsb", ".xls"):
        repair_files.extend(glob.glob(f"수리접수 내역*{ext}"))
    if not status_file:
        raise FileNotFoundError("필요한 '의료기기 현황조회' 파일을 찾을 수 없습니다.")
        
    df_status = pd.read_excel(status_file)
    repair_dfs = []
    for r_file in repair_files:
        try:
            if r_file.endswith('.xlsb'):
                df_r = pd.read_excel(r_file, engine='pyxlsb')
            else:
                df_r = pd.read_excel(r_file)
            repair_dfs.append(df_r)
        except Exception:
            pass
            
    if repair_dfs:
        df_repair = pd.concat(repair_dfs, ignore_index=True)
        df_repair = df_repair.drop_duplicates()
    else:
        df_repair = pd.DataFrame()
        
    return df_status, df_repair, status_file

try:
    df_status, df_repair, latest_status_path = load_latest_data()
except Exception as e:
    st.error(f"데이터 파일을 불러오는 중 오류가 발생했습니다: {e}")
    st.stop()

# ==========================================
# 🔗 URL 쿼리 파라미터 처리 (?mgm=관리번호)
# ==========================================
query_params = st.query_params
mgm_query = query_params.get("mgm", "")
if isinstance(mgm_query, list):
    mgm_query = mgm_query[0]
mgm_query = mgm_query.strip().upper()

st.markdown("### 🏷️ 의료장비 예방점검 라벨 조회")
st.markdown("---")

if not mgm_query:
    st.warning("💡 관리번호가 지정되지 않았습니다.")
    st.info("URL 뒤에 `?mgm=관리번호`를 붙여서 접속해 주세요.\n\n**사용 예시:** `https://your-app-url.streamlit.app/?mgm=50A1100001`")
else:
    dept_col = '사용\n부서' if '사용\n부서' in df_status.columns else '사용부서'
    matched_status = df_status[df_status['관리번호'].astype(str).str.strip().str.upper() == mgm_query]
    
    if matched_status.empty:
        st.error(f"입력하신 관리번호 (**{mgm_query}**)에 해당하는 장비를 찾을 수 없습니다.")
    else:
        status_record = matched_status.iloc[0]
        equipment_name = status_record.get('장비명/구성품명', '-')
        raw_dept_val = status_record.get(dept_col, '-')
        dept_name = str(raw_dept_val).replace('\n', ' ')
        
        count_col = [c for c in status_record.index if '정도관리' in str(c) or '회/년' in str(c) or '점검횟수' in str(c)]
        inspection_count_val = status_record.get(count_col[0], 1) if count_col else 1
        
        risk_col = [c for c in status_record.index if '위험' in str(c) or '등급' in str(c)]
        risk_grade = status_record.get(risk_col[0], '-') if risk_col else '-'
        if pd.isna(risk_grade) or str(risk_grade).strip() == '':
            risk_grade = "정보 없음"
        
        prevent_col = [c for c in df_repair.columns if '예방' in c and '점검' in c]
        if prevent_col and not df_repair.empty:
            p_col = prevent_col[0]
            repair_matched = df_repair[
                (df_repair['관리번호'].astype(str).str.strip().str.upper() == mgm_query) &
                (df_repair[p_col].astype(str).str.strip().str.upper() == 'Y')
            ].copy()
        else:
            repair_matched = pd.DataFrame()
            
        last_inspection_date = "-"
        next_inspection_date = "-"
        repairer_name = status_record.get('의공담당', status_record.get('담당자', '-'))
        
        if not repair_matched.empty and '완료일자' in repair_matched.columns:
            repair_matched['완료일자_dt'] = pd.to_datetime(repair_matched['완료일자'], errors='coerce')
            valid_df = repair_matched.dropna(subset=['완료일자_dt'])
            if not valid_df.empty:
                latest_row = valid_df.loc[valid_df['완료일자_dt'].idxmax()]
                latest_dt = latest_row['완료일자_dt']
                last_inspection_date = latest_dt.strftime('%Y-%m-%d')
                
                if '담당자' in latest_row and pd.notna(latest_row['담당자']):
                    repairer_name = latest_row['담당자']
                    
                match_cnt = re.search(r'(\d+)', str(inspection_count_val))
                cnt = int(match_cnt.group(1)) if match_cnt else 1
                    
                if cnt > 0:
                    add_months = max(1, int(round(12 / cnt)))
                    next_dt = latest_dt + pd.DateOffset(months=add_months)
                    next_inspection_date = next_dt.strftime('%Y-%m-%d')
                else:
                    next_dt = latest_dt + pd.DateOffset(months=12)
                    next_inspection_date = next_dt.strftime('%Y-%m-%d')

        if pd.isna(repairer_name) or str(repairer_name).strip() == '':
            repairer_name = "-"

        date_color = "#111"
        alert_html = ""
        is_disposed_dept = str(raw_dept_val).strip() in ['88', '77']

        if is_disposed_dept:
            next_inspection_display = '<span style="color: #d9534f; font-weight: bold;">폐기장비로 해당없음</span>'
        else:
            if next_inspection_date != "-":
                try:
                    today = datetime.now().date()
                    next_dt_obj = datetime.strptime(next_inspection_date, "%Y-%m-%d").date()
                    diff_days = (next_dt_obj - today).days
                    
                    if diff_days < 0:
                        date_color = "#d9534f"
                        alert_html = '<span style="color: #d9534f; font-weight: bold; margin-left: 5px;">(예방점검 의뢰요망)</span>'
                    elif 0 <= diff_days <= 14:
                        date_color = "#0275d8"
                        alert_html = '<span style="color: #0275d8; font-weight: bold; margin-left: 5px;">(예방점검 임박)</span>'
                except Exception:
                    pass
            next_inspection_display = f'<span style="color: {date_color}; font-weight: bold;">{next_inspection_date}</span> {alert_html}'

        # 예방점검 라벨 카드 UI 출력
        st.markdown(
            f"""
            <div style="border: 3px solid #333; padding: 20px; border-radius: 10px; background-color: #ffffff; font-family: sans-serif; color: #111; max-width: 500px; margin: auto; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                <div style="display: flex; justify-content: space-between; font-weight: bold; font-size: 1.15em; margin-bottom: 8px;">
                    <span>🏢 {dept_name}</span>
                    <span>🆔 {mgm_query}</span>
                </div>
                <div style="font-size: 1.1em; font-weight: bold; margin-bottom: 15px; color: #222;">
                    📦 {equipment_name}
                </div>
                <hr style="border: 0.5px solid #ccc; margin: 10px 0;">
                <div style="font-size: 1.05em; margin: 8px 0;"><b>점 검 일 자 :</b> {last_inspection_date}</div>
                <div style="font-size: 1.05em; margin: 8px 0;"><b>차 기 점 검 :</b> {next_inspection_display}</div>
                <div style="font-size: 0.95em; margin-top: 12px; color: #555;"><b>점검주기/등급 :</b> 년 {inspection_count_val}회 / {risk_grade}</div>
                <div style="text-align: center; font-weight: bold; font-size: 1.1em; margin-top: 20px; color: #222; border-top: 1px dashed #ddd; padding-top: 10px;">
                    인하대병원 의용공학팀 &nbsp;|&nbsp; 정비자: {repairer_name}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        if last_inspection_date == "-":
            st.info("ℹ️ 해당 장비의 수리내역 중 예방점검('Y') 이력이 존재하지 않습니다.")

        # 요청하신 링크 버튼 추가 영역
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            """
            <div style="text-align: center;">
                <a href="https://buly.kr/6BzfJgY" target="_blank" style="
                    display: inline-block;
                    background-color: #f0f2f6;
                    color: #262730;
                    padding: 0.75rem 1.25rem;
                    border-radius: 0.5rem;
                    font-weight: bold;
                    text-decoration: none;
                    border: 1px solid #d6d9dc;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
                    font-size: 1rem;
                ">🔗 장비 상세내역 및 수리이력 확인(회원가입 필요)</a>
            </div>
            """,
            unsafe_allow_html=True
        )