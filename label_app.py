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
# 🔗 URL 쿼리 파라미터 및 검색창 처리
# ==========================================
query_params = st.query_params
mgm_query = query_params.get("mgm", "")
if isinstance(mgm_query, list):
    mgm_query = mgm_query[0]
mgm_query = mgm_query.strip().upper()

# 전체 관리번호 목록 추출
dept_col = '사용\n부서' if '사용\n부서' in df_status.columns else '사용부서'
all_mgm = df_status['관리번호'].dropna().astype(str).str.strip().str.upper().unique().tolist()
all_mgm.sort()

# 기본 관리번호 설정
if not mgm_query or mgm_query not in all_mgm:
    if all_mgm:
        mgm_query = all_mgm[0]

# 📱 [특정 앱 실행 버튼] 패키지명(com.example_qr_web_opener) 연동 인텐트 적용 (상단 배치)
st.markdown(
    """
    <div style="text-align: center; margin-bottom: 10px;">
        <a href="intent://#Intent;package=com.example_qr_web_opener;end;" style="
            display: inline-block;
            background-color: #ff4b4b;
            color: white;
            padding: 0.6rem 1rem;
            border-radius: 0.5rem;
            font-weight: bold;
            text-decoration: none;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            font-size: 0.95rem;
            width: 100%;
            text-align: center;
        ">📷 휴대폰 스캔 앱(QR/바코드) 열기</a>
    </div>
    """,
    unsafe_allow_html=True
)

# 🔍 검색창과 [조회] 버튼을 나란히 배치하기 위한 컬럼 분할
col1, col2 = st.columns([4, 1])

with col1:
    search_input = st.text_input("🔍 장비 검색 (관리번호 또는 장비명 입력):", value=mgm_query, label_visibility="collapsed")

with col2:
    st.markdown("<div style='margin-top: 2px;'></div>", unsafe_allow_html=True) # 줄맞춤용 여백
    search_btn = st.button("조회", use_container_width=True)

# 버튼이 눌렸거나 검색창 입력값이 변경된 경우 처리
if search_input:
    search_keyword = search_input.strip().upper()
    matched_status = df_status[
        (df_status['관리번호'].astype(str).str.strip().str.upper() == search_keyword) |
        (df_status['장비명/구성품명'].astype(str).str.upper().str.contains(search_keyword, na=False))
    ]
    if not matched_status.empty:
        mgm_query = str(matched_status.iloc[0]['관리번호']).strip().upper()
    else:
        mgm_query = search_keyword

st.markdown("---")

if not mgm_query or all_mgm.count(mgm_query) == 0:
    st.error(f"입력하신 검색어(**{mgm_query}**)에 해당하는 장비를 찾을 수 없습니다. 올바른 관리번호나 장비명을 입력해 주세요.")
else:
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
            <div style="border: 3px solid #333; padding: 18px; border-radius: 10px; background-color: #ffffff; font-family: sans-serif; color: #111; max-width: 500px; margin: auto; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                <div style="display: flex; justify-content: space-between; font-weight: bold; font-size: 1.1em; margin-bottom: 6px;">
                    <span>🏢 {dept_name}</span>
                    <span>🆔 {mgm_query}</span>
                </div>
                <div style="font-size: 1.05em; font-weight: bold; margin-bottom: 12px; color: #222;">
                    📦 {equipment_name}
                </div>
                <hr style="border: 0.5px solid #ccc; margin: 8px 0;">
                <div style="font-size: 1em; margin: 6px 0;"><b>점 검 일 자 :</b> {last_inspection_date}</div>
                <div style="font-size: 1em; margin: 6px 0;"><b>차 기 점 검 :</b> {next_inspection_display}</div>
                <div style="font-size: 0.9em; margin-top: 10px; color: #555;"><b>점검주기/등급 :</b> 년 {inspection_count_val}회 / {risk_grade}</div>
                <div style="text-align: center; font-weight: bold; font-size: 1em; margin-top: 15px; color: #222; border-top: 1px dashed #ddd; padding-top: 8px;">
                    인하대병원 의용공학팀 &nbsp;|&nbsp; 정비자: {repairer_name}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        if last_inspection_date == "-":
            st.info("ℹ️ 해당 장비의 수리내역 중 예방점검('Y') 이력이 존재하지 않습니다.")

        # 장비 상세내역 및 수리이력 확인 링크 버튼 영역
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            """
            <div style="text-align: center;">
                <a href="https://buly.kr/6BzfJgY" target="_blank" style="
                    display: inline-block;
                    background-color: #f0f2f6;
                    color: #262730;
                    padding: 0.6rem 1rem;
                    border-radius: 0.5rem;
                    font-weight: bold;
                    text-decoration: none;
                    border: 1px solid #d6d9dc;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
                    font-size: 0.95rem;
                ">🔗 장비 상세내역 및 수리이력 확인(회원가입 필요)</a>
            </div>
            """,
            unsafe_allow_html=True
        )