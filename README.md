            sel_e = st.selectbox("평가 기준 선택", e_keys, key="task_sel_eval")
            
            # 🔴 [UI 통합] 옵션을 하나로 합치고 기본값을 '덮어쓰기'로 설정
            analysis_mode = st.radio(
                "🎯 분석 모드 설정",
                ["덮어쓰기 (모든 행 새로 분석)", "건너뛰기 (이미 분석된 행 제외)"],
                index=0, # 기본값: 덮어쓰기
                horizontal=True
            )

            # 건너뛰기 선택 시 설명 표시
            if "건너뛰기" in analysis_mode:
                st.info("💡 **건너뛰기 모드:** 엑셀의 'Code' 또는 '중요도' 칸이 이미 채워져 있는 경우, AI 분석을 생략하고 기존 데이터를 유지합니다.")

            if st.button("🚀 분석 프로세스 시작", disabled=st.session_state.is_running, type="primary", use_container_width=True, key="task_start_btn"):
                st.session_state.is_running = True
                st.session_state.task_finished = False
                st.session_state.thread_flag = [False]
                st.session_state.analysis_logs = [] 
                st.session_state.processed_indices = set()
                st.session_state.current_library_rows = []
                st.session_state.start_time = time.time()
                st.session_state.current_wb = openpyxl.load_workbook(BytesIO(st.session_state.uploaded_wb_bytes))
                
                ws = st.session_state.current_wb.worksheets[2]
                df_target = st.session_state.uploaded_df.iloc[int(start_num)-6 : int(end_num)-5]
                ai_cols = ['Lv1', 'Lv2', 'Lv3', 'Lv4', 'Lv5', 'Code', '중요도', '코멘트', '평가근거']
                col_idx_map = {name: get_or_create_col_idx(ws, 5, name) for name in ai_cols}
                
                tech_table = parse_tech_tree_to_df(st.session_state['tech_trees'][sel_t]).to_csv(index=False)
                tech_tree_df = parse_tech_tree_to_df(st.session_state['tech_trees'][sel_t])
                eval_guide = st.session_state['eval_criteria'][sel_e]

                def run_background():
                    MAX_RETRIES = 3 
                    RETRY_DELAY = 2 
                    
                    try:
                        for df_idx, row in df_target.iterrows():
                            if st.session_state.thread_flag[0]:
                                break
                                
                            real_row = df_idx + 6
                            
                            # 🔴 [건너뛰기 판단 로직] 엑셀에서 기존 값들을 읽어옴
                            ext_code = str(get_val_robust(row, ['Code', '코드'])).strip()
                            ext_imp = str(get_val_robust(row, ['중요도', '등급', 'Importance'])).strip()
                            ext_comment = str(get_val_robust(row, ['코멘트', '요약', 'Comment'])).strip()
                            
                            # 건너뛰기 조건: 모드가 '건너뛰기'이고, 코드나 중요도 중 하나라도 채워져 있을 때
                            is_skip_target = False
                            if "건너뛰기" in analysis_mode:
                                if (ext_code not in ["-", "", "nan"]) or (ext_imp not in ["-", "", "nan"]):
                                    is_skip_target = True

                            if is_skip_target:
                                # 🔴 [요청 반영] 기존 중요도와 코멘트를 그대로 사용하고 로그에 표시
                                skip_data = {
                                    "Lv1": "-", "Lv2": "-", "Lv3": "-", "Lv4": "-", "Lv5": "-",
                                    "Code": ext_code,
                                    "중요도": ext_imp,
                                    "코멘트": ext_comment if ext_comment not in ["-", "nan"] else "(기존 코멘트 없음)",
                                    "평가근거": "이미 분석이 완료된 건이므로 분석을 건너뜁니다."
                                }
                                st.session_state.analysis_logs.append({"행": real_row, "상태": "⏭️ 생략", **skip_data})
                                
                                # Library 저장을 위해 base_data 구성 및 리스트 추가
                                skip_base = {
                                    '평가날짜': datetime.now().strftime('%Y-%m-%d'), '평가자': evaluator,
                                    '출원번호': str(get_val_robust(row, ['출원번호', '출원 번호'])).strip(),
                                    '대표청구항': str(get_val_robust(row, ['전체청구항', '전체 청구항', '대표청구항'])),
                                    '원본링크': ws.cell(row=real_row, column=1).hyperlink.target if ws.cell(row=real_row, column=1).hyperlink else ""
                                }
                                skip_base.update(skip_data)
                                st.session_state.current_library_rows.append(skip_base)
                                st.session_state.processed_indices.add(df_idx)
                                continue

                            # --- 여기서부터는 '덮어쓰기'이거나 '빈 칸'인 경우 실행되는 기존 AI 분석 로직 ---
                            
                            # 1. 청구항 추출
                            full_claim = get_val_robust(row, ['전체청구항', '전체 청구항', 'All Claims'])
                            rep_claim = get_val_robust(row, ['대표청구항', '대표 청구항', '청구항', '독립항'])
                            target_claim = full_claim if full_claim != "-" else rep_claim

                            # 2. base_data 준비
                            base_data = {
                                '평가날짜': datetime.now().strftime('%Y-%m-%d'), 
                                '평가자': evaluator,
                                '일련번호': get_val_robust(row, ['일련번호', 'No', 'No.']), 
                                '출원인': get_val_robust(row, ['출원인', '출원인대표명', '현재권리자', 'Applicant']),
                                '출원번호': get_val_robust(row, ['출원번호', '출원 번호']), 
                                '출원일': get_val_robust(row, ['출원일', '출원일자'], True),
                                '등록번호': get_val_robust(row, ['등록번호', '등록 번호']), 
                                '등록일': get_val_robust(row, ['등록일', '등록일자'], True),
                                '대표청구항': target_claim, 
                                '발명자평균등급': get_val_robust(row, ['발명자평균등급', '발명자 평균등급', '등급'])
                            }
                            
                            sys_p, usr_p = generate_ai_prompt(tech_table, eval_guide, base_data['대표청구항'])
                            
                            success = False
                            ai_data = {k: "-" for k in ai_cols} 
                            
                            for attempt in range(MAX_RETRIES):
                                try:
                                    time.sleep(1.0) 
                                    # res_raw = my_company_ai_api_call(sys_p, usr_p) 
                                    res_raw = '{"Lv1":"반도체","Lv2":"Memory","Lv3":"FeVNAND","Lv4":"-","Lv5":"-","Code":"MEM","중요도":"S","코멘트":"AI 분석 완료.","평가근거":"테스트"}' 
                                    
                                    ai_data = json.loads(re.search(r'\{.*\}', res_raw, re.DOTALL).group(0))
                                    null_words = ["none", "null", "", "nan", "n/a", "na", "-", "_", "."]
                                    for k in ai_cols: 
                                        v = str(ai_data.get(k, "-")).strip()
                                        ai_data[k] = "-" if v.lower() in null_words else v
                                        
                                    parsed_lvs = [str(ai_data.get(f'Lv{i}', '-')).strip() for i in range(1, 6)]
                                    match = tech_tree_df[
                                        (tech_tree_df['Lv1'].astype(str).str.strip() == parsed_lvs[0]) &
                                        (tech_tree_df['Lv2'].astype(str).str.strip() == parsed_lvs[1]) &
                                        (tech_tree_df['Lv3'].astype(str).str.strip() == parsed_lvs[2]) &
                                        (tech_tree_df['Lv4'].astype(str).str.strip() == parsed_lvs[3]) &
                                        (tech_tree_df['Lv5'].astype(str).str.strip() == parsed_lvs[4])
                                    ]
                                    if not match.empty: ai_data['Code'] = match.iloc[0]['Code']
                                    else: ai_data['Code'] = "-"
                                    success = True
                                    break 
                                except Exception as e:
                                    time.sleep(RETRY_DELAY * (attempt + 1))
                            
                            # 엑셀 셀 기록 및 세션 누적
                            if success:
                                for k in ai_cols: ws.cell(row=real_row, column=col_idx_map[k]).value = ai_data.get(k, "-")
                                st.session_state.analysis_logs.append({"행": real_row, "상태": "✅ 완료", **ai_data})
                                base_data.update(ai_data)
                            else:
                                for k in ai_cols: ws.cell(row=real_row, column=col_idx_map[k]).value = "-"
                                st.session_state.analysis_logs.append({"행": real_row, "상태": "❌ 실패", "내용": "분석 실패"})
                                base_data.update({k: "-" for k in ai_cols})
                                
                            base_data['원본링크'] = ws.cell(row=real_row, column=1).hyperlink.target if ws.cell(row=real_row, column=1).hyperlink else ""
                            st.session_state.current_library_rows.append(base_data)
                            st.session_state.processed_indices.add(df_idx)
                            
                    except Exception as e: 
                        st.session_state.analysis_logs.append({"행": "-", "상태": "❌ 오류", "내용": str(e)})
                    finally: 
                        st.session_state.task_finished = True
