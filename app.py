import streamlit as st
import pandas as pd
import json
import os
import openpyxl
from io import BytesIO
from datetime import datetime
from PIL import Image as PILImage
import uuid

# --- 0. 파일 경로 설정 ---
DB_FILE = "patent_library_v2.csv"
TRASH_FILE = "patent_trash_v2.csv"
TECH_TREE_FILE = "tech_trees.json"
EVAL_CRITERIA_FILE = "eval_criteria.json"
UPDATES_FILE = "system_updates.json"

# --- 1. 유틸리티 함수 ---
def load_csv(fp):
    if os.path.exists(fp):
        try: return pd.read_csv(fp)
        except: pass
    return pd.DataFrame()

def save_csv(fp, df):
    df.to_csv(fp, index=False, encoding='utf-8-sig')

def load_json(fp):
    if os.path.exists(fp):
        try:
            with open(fp, 'r', encoding='utf-8') as f: return json.load(f)
        except: pass
    return []

def save_json(fp, data):
    with open(fp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def ensure_columns(df, is_trash=False):
    cols = ['출원번호', '출원인', 'LibraryName', 'is_pinned', 'Code', 'Lv1', 'Lv2', 'Lv3', 'Lv4', 'Lv5', 
            '중요도', '코멘트', '평가근거', '요약제목', '목적', '방법', 'Recipe', '연구시사점', '키워드', '원본링크', '이미지경로']
    if is_trash: cols.append('원래Library위치')
    for c in cols:
        if c not in df.columns:
            df[c] = False if c == 'is_pinned' else "-"
    return df

# --- 2. 초기 상태 설정 ---
st.set_page_config(page_title="특허 자동 센싱 시스템", layout="wide")

if 'library_data' not in st.session_state: st.session_state.library_data = ensure_columns(load_csv(DB_FILE))
if 'trash_data' not in st.session_state: st.session_state.trash_data = ensure_columns(load_csv(TRASH_FILE), True)
if 'updates_data' not in st.session_state: st.session_state.updates_data = load_json(UPDATES_FILE)

# --- 관리 메뉴용 상태 초기화 (list of dict) ---
if 'tech_trees' not in st.session_state: st.session_state.tech_trees = load_json(TECH_TREE_FILE)
if 'eval_criteria' not in st.session_state: st.session_state.eval_criteria = load_json(EVAL_CRITERIA_FILE)

# --- 3. 사이드바 메뉴 ---
menu = st.sidebar.radio(
    "메뉴 이동",
    ["🏠 홈", "(1) 기술 트리 관리", "(2) 중요도 기준 관리", "(3) AI 분석 작업 시작", 
     "(4) Library", "(5) 휴지통", "(6) Library 간 비교", "(7) 주요 특허 리포트", "(8) 업데이트 사항"]
)

# ==========================================
# 🏠 홈
# ==========================================
if menu == "🏠 홈":
    st.title("🏠 특허 자동 센싱 시스템")
    st.markdown("""
    ### 📌 시스템 가이드라인
    이 시스템은 반도체 특허 데이터를 관리하고 분석하기 위한 통합 환경입니다.
    * **(1)~(2) 관리**: 기술 트리와 중요도 기준을 설정합니다.
    * **(3) AI 분석**: 엑셀 업로드 시 대표 이미지와 청구항을 추출하여 다각도로 분석합니다.
    * **(4)~(5) Library/휴지통**: 분석된 특허를 보관하고, 삭제 및 순서 변경을 관리합니다.
    * **(6) 비교**: 서로 다른 Library 간의 분석 일치도를 정밀 비교합니다.
    * **(7) 리포트**: 핀(PIN) 고정 기능을 통해 중요 특허의 요약 리포트를 자동 생성합니다.
    """)

# ==========================================
# (1) 기술 트리 관리
# ==========================================
elif menu == "(1) 기술 트리 관리":
    st.title("📂 (1) 기술 트리 관리")
    
    with st.expander("➕ 새 기술 트리 추가", expanded=True):
        new_tt_title = st.text_input("기술 트리 제목", placeholder="예: 2026년 상반기 M3D 분류 기준")
        new_tt_content = st.text_area("상세 내용 (슬래시 '/' 구분)", placeholder="예: 소자/FeVNAND/M3D/중간절연막/M_IL")
        if st.button("저장하기", key="btn_save_tt"):
            if new_tt_title and new_tt_content:
                new_item = {"id": str(uuid.uuid4()), "title": new_tt_title, "content": new_tt_content}
                st.session_state.tech_trees.append(new_item)
                save_json(TECH_TREE_FILE, st.session_state.tech_trees)
                st.success("저장되었습니다.")
                st.rerun()
            else:
                st.warning("제목과 내용을 모두 입력해주세요.")
                
    st.markdown("### 📋 저장된 기술 트리 목록")
    if not st.session_state.tech_trees:
        st.info("저장된 기술 트리가 없습니다.")
    else:
        for idx, item in enumerate(st.session_state.tech_trees):
            with st.container(border=True):
                col1, col2, col3 = st.columns([0.6, 0.2, 0.2])
                with col1:
                    # 제목 수정 기능
                    edit_title = st.text_input(f"제목 수정", value=item['title'], key=f"tt_title_{item['id']}", label_visibility="collapsed")
                with col2:
                    if st.button("💾 제목 수정", key=f"tt_edit_{item['id']}"):
                        st.session_state.tech_trees[idx]['title'] = edit_title
                        save_json(TECH_TREE_FILE, st.session_state.tech_trees)
                        st.success("수정됨")
                with col3:
                    if st.button("🗑️ 삭제", key=f"tt_del_{item['id']}"):
                        st.session_state.tech_trees.pop(idx)
                        save_json(TECH_TREE_FILE, st.session_state.tech_trees)
                        st.rerun()
                st.text(item['content'])

# ==========================================
# (2) 중요도 기준 관리
# ==========================================
elif menu == "(2) 중요도 기준 관리":
    st.title("📊 (2) 중요도 평가 기준 관리")
    
    with st.expander("➕ 새 평가 기준 추가", expanded=True):
        new_ec_title = st.text_input("평가 기준 제목", placeholder="예: 2026년 특허 중요도 가이드")
        new_ec_content = st.text_area("상세 기준", placeholder="[S급]: 핵심 원천 기술\n[A급]: 양산 적용성 높음")
        if st.button("저장하기", key="btn_save_ec"):
            if new_ec_title and new_ec_content:
                new_item = {"id": str(uuid.uuid4()), "title": new_ec_title, "content": new_ec_content}
                st.session_state.eval_criteria.append(new_item)
                save_json(EVAL_CRITERIA_FILE, st.session_state.eval_criteria)
                st.success("저장되었습니다.")
                st.rerun()
            else:
                st.warning("제목과 내용을 모두 입력해주세요.")

    st.markdown("### 📋 저장된 평가 기준 목록")
    if not st.session_state.eval_criteria:
        st.info("저장된 평가 기준이 없습니다.")
    else:
        for idx, item in enumerate(st.session_state.eval_criteria):
            with st.container(border=True):
                col1, col2, col3 = st.columns([0.6, 0.2, 0.2])
                with col1:
                    edit_title = st.text_input(f"제목 수정", value=item['title'], key=f"ec_title_{item['id']}", label_visibility="collapsed")
                with col2:
                    if st.button("💾 제목 수정", key=f"ec_edit_{item['id']}"):
                        st.session_state.eval_criteria[idx]['title'] = edit_title
                        save_json(EVAL_CRITERIA_FILE, st.session_state.eval_criteria)
                        st.success("수정됨")
                with col3:
                    if st.button("🗑️ 삭제", key=f"ec_del_{item['id']}"):
                        st.session_state.eval_criteria.pop(idx)
                        save_json(EVAL_CRITERIA_FILE, st.session_state.eval_criteria)
                        st.rerun()
                st.text(item['content'])

# ==========================================
# (3) AI 분석 작업 시작
# ==========================================
elif menu == "(3) AI 분석 작업 시작":
    st.title("🚀 (3) AI 분석 작업 시작")
    
    uploaded_file = st.file_uploader("📂 특허 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])
    
    # 🔴 수정사항: 업로드 전에 아래 UI를 숨김
    if uploaded_file is not None:
        st.markdown("---")
        st.subheader("⚙️ 분석 옵션 설정")
        lib_name = st.text_input("저장할 Library 이름 지정", value="New_Patent_Set")
        
        # 적용할 트리 및 기준 선택 (저장된 목록 활용)
        tt_opts = [t['title'] for t in st.session_state.tech_trees] if st.session_state.tech_trees else ["(기본 트리)"]
        ec_opts = [e['title'] for e in st.session_state.eval_criteria] if st.session_state.eval_criteria else ["(기본 기준)"]
        
        col_tt, col_ec = st.columns(2)
        sel_tt = col_tt.selectbox("적용할 기술 트리", tt_opts)
        sel_ec = col_ec.selectbox("적용할 중요도 기준", ec_opts)
        
        if st.button("▶️ 분석 시작", type="primary"):
            wb = openpyxl.load_workbook(BytesIO(uploaded_file.read()), data_only=True)
            ws = wb.active
            
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file)
            
            claim_col = '전체청구항' if '전체청구항' in df.columns else ('대표청구항' if '대표청구항' in df.columns else None)
            
            if not claim_col:
                st.error("엑셀 파일에 '전체청구항' 또는 '대표청구항' 열이 없습니다.")
            else:
                img_dir = f"patent_images_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                os.makedirs(img_dir, exist_ok=True)
                
                image_map = {}
                if hasattr(ws, '_images'):
                    for img in ws._images:
                        if img.anchor._from.col == 1: # B열 
                            row_idx = img.anchor._from.row + 1 
                            try:
                                pil_img = PILImage.open(BytesIO(img._data()))
                                img_path = f"{img_dir}/row_{row_idx}.png"
                                pil_img.save(img_path)
                                image_map[row_idx] = img_path
                            except: pass

                st.success(f"✅ 분석 기준 열: **{claim_col}** (적용 트리: {sel_tt} / 기준: {sel_ec})")
                log_container = st.empty()
                new_rows = []
                
                for idx, row in df.iterrows():
                    real_row = idx + 2
                    app_num = str(row.get('출원번호', f"Dummy_{idx}"))
                    log_container.text(f"⏳ 실시간 로그: [행 {real_row}] {app_num} 분석 중... (기준: {claim_col})")
                    
                    ai_result = {
                        'Code': 'M_IL', 'Lv1': '소자', 'Lv2': 'FeVNAND', 'Lv3': 'M3D', 'Lv4': '중간절연막', 'Lv5': 'M_IL',
                        '중요도': 'A', '코멘트': 'AI 자동 요약 코멘트', '평가근거': '수직 구조 적용성 우수',
                        '요약제목': '초고층 3D 낸드 절연막 개선', '목적': '누설 전류 차단 및 집적도 향상',
                        '방법': 'High-K 물질과 산화막 교대 증착', 'Recipe': 'ALD 공정 180도, AL2O3 5nm', 
                        '연구시사점': 'FeFET 계면 특성 개선에 적용 가능성 높음', '키워드': 'M3D, ALD, High-K, 누설전류'
                    }
                    
                    record = {
                        '출원번호': app_num, '출원인': str(row.get('출원인', 'Unknown')), 'LibraryName': lib_name, 'is_pinned': False,
                        '원본링크': str(row.get('원본링크', '-')), '이미지경로': image_map.get(real_row, "-"),
                        **ai_result
                    }
                    new_rows.append(record)
                    
                new_df = pd.DataFrame(new_rows)
                st.session_state.library_data = pd.concat([st.session_state.library_data, new_df], ignore_index=True)
                save_csv(DB_FILE, st.session_state.library_data)
                log_container.text(f"🎉 총 {len(new_rows)}건 분석 완료 및 Library 저장 성공!")

# ==========================================
# (4) Library
# ==========================================
elif menu == "(4) Library":
    st.title("📚 (4) Library")
    
    # 🔴 수정사항: 데이터가 비어있을 때 명확히 안내
    if st.session_state.library_data.empty:
        st.warning("⚠️ 현재 Library에 저장된 데이터가 없습니다. '(3) AI 분석 작업 시작'에서 데이터를 먼저 분석해 주세요.")
    else:
        libs = st.session_state.library_data['LibraryName'].unique()
        if len(libs) == 0:
             st.warning("⚠️ 저장된 Library 그룹이 없습니다.")
        else:
            selected_lib = st.selectbox("조회할 Library 선택", libs)
            
            lib_df = st.session_state.library_data[st.session_state.library_data['LibraryName'] == selected_lib]
            
            if lib_df.empty:
                st.info(f"선택하신 '{selected_lib}'에 해당하는 특허가 없습니다.")
            else:
                st.dataframe(lib_df[['출원번호', '요약제목', '중요도', 'Code', 'is_pinned']])
                
                st.markdown("---")
                st.subheader("🛠️ 개별 특허 관리 (삭제/순서 변경)")
                
                all_indices = lib_df.index.tolist()
                if all_indices:
                    sel_idx = st.selectbox("관리할 특허 선택 (출원번호 기준)", all_indices, 
                                           format_func=lambda x: st.session_state.library_data.at[x, '출원번호'])
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        if st.button("🗑️ 삭제 (휴지통으로)"):
                            target = st.session_state.library_data.loc[[sel_idx]].copy()
                            target['원래Library위치'] = selected_lib
                            st.session_state.trash_data = pd.concat([st.session_state.trash_data, target], ignore_index=True)
                            st.session_state.library_data = st.session_state.library_data.drop(sel_idx).reset_index(drop=True)
                            save_csv(DB_FILE, st.session_state.library_data)
                            save_csv(TRASH_FILE, st.session_state.trash_data)
                            st.rerun()
                    
                    with col2:
                        if st.button("⬆️ 순서 위로") and sel_idx > 0:
                            b, a = st.session_state.library_data.iloc[sel_idx].copy(), st.session_state.library_data.iloc[sel_idx - 1].copy()
                            st.session_state.library_data.iloc[sel_idx], st.session_state.library_data.iloc[sel_idx - 1] = a, b
                            save_csv(DB_FILE, st.session_state.library_data)
                            st.rerun()
                    with col3:
                        if st.button("⬇️ 순서 아래로") and sel_idx < len(st.session_state.library_data) - 1:
                            b, a = st.session_state.library_data.iloc[sel_idx].copy(), st.session_state.library_data.iloc[sel_idx + 1].copy()
                            st.session_state.library_data.iloc[sel_idx], st.session_state.library_data.iloc[sel_idx + 1] = a, b
                            save_csv(DB_FILE, st.session_state.library_data)
                            st.rerun()

# ==========================================
# (5) 휴지통
# ==========================================
elif menu == "(5) 휴지통":
    st.title("🗑️ (5) 휴지통")
    if st.session_state.trash_data.empty:
        st.info("휴지통이 비어 있습니다.")
    else:
        st.write("삭제된 데이터입니다. (원래 Library 위치 포함)")
        cols = ['원래Library위치', '출원번호', '출원인', '요약제목', '중요도', 'Code']
        st.dataframe(st.session_state.trash_data[cols], use_container_width=True)

# ==========================================
# (6) Library 간 비교
# ==========================================
elif menu == "(6) Library 간 비교":
    st.title("⚖️ (6) Library 간 일치도 비교")
    
    if st.session_state.library_data.empty:
        st.warning("⚠️ 현재 Library에 저장된 데이터가 없습니다.")
    else:
        libs = st.session_state.library_data['LibraryName'].unique()
        if len(libs) >= 2:
            col1, col2 = st.columns(2)
            with col1: lib1 = st.selectbox("Library 1 선택", libs, index=0)
            with col2: lib2 = st.selectbox("Library 2 선택", libs, index=1)
            
            st.markdown("### ⚙️ 비교 옵션")
            c1, c2, c3, c4, c5 = st.columns(5)
            opt_sa = c1.checkbox("S/A급 동일 간주")
            opt_bx = c2.checkbox("B/X급 동일 간주")
            opt_code = c3.checkbox("Code만 비교")
            opt_imp = c4.checkbox("중요도만 비교")
            opt_ex_sam = c5.checkbox("자사(Samsung Electronics) 제외")
            
            if st.button("비교 실행"):
                df1 = st.session_state.library_data[st.session_state.library_data['LibraryName'] == lib1]
                df2 = st.session_state.library_data[st.session_state.library_data['LibraryName'] == lib2]
                
                if opt_ex_sam:
                    df1 = df1[~df1['출원인'].str.contains('Samsung Electronics', case=False, na=False)]
                    df2 = df2[~df2['출원인'].str.contains('Samsung Electronics', case=False, na=False)]
                    
                merged = pd.merge(df1, df2, on='출원번호', suffixes=('_1', '_2'))
                
                match_list, mismatch_list = [], []
                for _, row in merged.iterrows():
                    i1, i2 = str(row['중요도_1']).upper(), str(row['중요도_2']).upper()
                    if opt_sa:
                        i1 = 'A' if i1 == 'S' else i1
                        i2 = 'A' if i2 == 'S' else i2
                    if opt_bx:
                        i1 = 'X' if i1 == 'B' else i1
                        i2 = 'X' if i2 == 'B' else i2
                        
                    is_code_match = (row['Code_1'] == row['Code_2'])
                    is_imp_match = (i1 == i2)
                    
                    if opt_code and not opt_imp: is_match = is_code_match
                    elif opt_imp and not opt_code: is_match = is_imp_match
                    else: is_match = (is_code_match and is_imp_match) 
                    
                    res_dict = {
                        '출원번호': row['출원번호'],
                        'Code_1': row['Code_1'], 'Code_2': row['Code_2'],
                        '중요도_1': row['중요도_1'], '중요도_2': row['중요도_2'],
                        '코멘트_1': row['코멘트_1'], '코멘트_2': row['코멘트_2'],
                        '판단근거_1': row['평가근거_1'], '판단근거_2': row['평가근거_2']
                    }
                    
                    if is_match: match_list.append(res_dict)
                    else: mismatch_list.append(res_dict)
                
                st.success(f"**총 비교 건수:** {len(merged)}건 | **일치:** {len(match_list)}건 | **불일치:** {len(mismatch_list)}건")
                st.subheader("🟢 일치 건")
                st.dataframe(pd.DataFrame(match_list))
                st.subheader("🔴 불일치 건")
                st.dataframe(pd.DataFrame(mismatch_list))
        else:
            st.warning("비교할 Library가 2개 이상 필요합니다.")

# ==========================================
# (7) 주요 특허 리포트
# ==========================================
elif menu == "(7) 주요 특허 리포트":
    st.title("📄 (7) 주요 특허 요약 리포트")
    
    # 🔴 수정사항: 데이터가 비어있을 때 명확히 안내
    if st.session_state.library_data.empty:
         st.warning("⚠️ 현재 작성 가능한 리포트 데이터가 없습니다. 먼저 특허 분석을 진행해주세요.")
    else:
        libs = st.session_state.library_data['LibraryName'].unique()
        if len(libs) == 0:
             st.warning("⚠️ 저장된 Library 그룹이 없습니다.")
        else:
            c1, c2 = st.columns(2)
            sel_lib = c1.selectbox("Library 선택", libs, key="rep_lib")
            
            rep_df = st.session_state.library_data[st.session_state.library_data['LibraryName'] == sel_lib]
            
            if rep_df.empty:
                st.info(f"'{sel_lib}'에 분석된 특허가 없습니다.")
            else:
                imps = ['전체'] + list(rep_df['중요도'].unique())
                sel_imp = c2.selectbox("중요도 필터", imps, key="rep_imp")
                
                if sel_imp != '전체':
                    rep_df = rep_df[rep_df['중요도'] == sel_imp]
                    
                rep_df = rep_df.sort_values(by='is_pinned', ascending=False)
                
                if not rep_df.empty:
                    pat_options = {idx: f"{'📌' if row['is_pinned'] else ''} [{row['중요도']}] {row['출원번호']} / {row['출원인']} / {row['요약제목']}" 
                                   for idx, row in rep_df.iterrows()}
                    
                    sel_pat_idx = st.selectbox("특허 선택", list(pat_options.keys()), format_func=lambda x: pat_options[x])
                    p = rep_df.loc[sel_pat_idx]
                    
                    pin_label = "📌 핀 해제하기" if p['is_pinned'] else "🏳️ 핀 고정하기"
                    if st.button(pin_label):
                        st.session_state.library_data.at[sel_pat_idx, 'is_pinned'] = not p['is_pinned']
                        save_csv(DB_FILE, st.session_state.library_data)
                        st.rerun()

                    st.markdown("---")
                    col_text, col_img = st.columns([0.7, 0.3])
                    
                    with col_text:
                        st.markdown(f"### 🏷️ {p['요약제목']}")
                        st.markdown(f"**출원번호**: {p['출원번호']} | **중요도**: {p['중요도']} | **원본링크**: {p['원본링크']}")
                        st.markdown(f"**기술 트리**: {p['Lv1']} > {p['Lv2']} > {p['Lv3']} > {p['Lv4']} > {p['Lv5']} (Code: **{p['Code']}**)")
                        st.markdown(f"**키워드**: {p['키워드']}")
                        st.markdown("---")
                        st.markdown(f"**🎯 목적**: {p['목적']}")
                        st.markdown(f"**🛠️ 방법**: {p['방법']}")
                        st.markdown(f"**🧪 Recipe**: {p['Recipe']}")
                        st.markdown(f"**💡 연구 시사점**: {p['연구시사점']}")
                    
                    with col_img:
                        img_path = str(p['이미지경로'])
                        if img_path and os.path.exists(img_path):
                            st.image(img_path, caption="대표 이미지", use_container_width=True)
                        else:
                            st.info("저장된 대표 이미지가 없습니다.")
                else:
                    st.warning("선택하신 중요도 조건에 맞는 특허가 없습니다.")

# ==========================================
# (8) 업데이트 사항
# ==========================================
elif menu == "(8) 업데이트 사항":
    st.title("🔔 (8) 시스템 업데이트 사항")
    
    with st.expander("➕ 새 업데이트 기록하기", expanded=True):
        up_title = st.text_input("업데이트 제목")
        up_content = st.text_area("상세 내용")
        if st.button("저장하기"):
            new_update = {
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "title": up_title,
                "content": up_content
            }
            st.session_state.updates_data.insert(0, new_update)
            save_json(UPDATES_FILE, st.session_state.updates_data)
            st.success("업데이트가 저장되었습니다.")
            
    st.markdown("---")
    for up in st.session_state.updates_data:
        st.markdown(f"#### 🔹 {up['title']}")
        st.caption(f"🕒 {up['time']}")
        st.write(up['content'])
        st.markdown("---")