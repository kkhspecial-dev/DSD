# DSD

            sel_e = st.selectbox("평가 기준 선택", e_keys, key="task_sel_eval")
            
            # 🔴 [수정] 스킵 기준을 엑셀 데이터 존재 여부로 명확히 표시합니다.
            skip_analyzed = st.checkbox("🔄 엑셀에 이미 분석된 항목(Code 또는 중요도 존재)은 건너뛰기", value=True)
            
            if st.button("🚀 분석 프로세스 시작", disabled=st.session_state.is_running, type="primary", use_container_width=True, key="task_start_btn"):




                        for df_idx, row in df_target.iterrows():
                            if st.session_state.thread_flag[0]:
                                break
                                
                            real_row = df_idx + 6
                            
                            # 1. 청구항 추출
                            full_claim = get_val_robust(row, ['전체청구항', '전체 청구항', 'All Claims'])
                            rep_claim = get_val_robust(row, ['대표청구항', '대표 청구항', '청구항', '독립항'])
                            target_claim = full_claim if full_claim != "-" else rep_claim
                            
                            # 🔴🔴🔴 [여기서부터 덮어씌우세요! 엑셀 값 기준 건너뛰기 로직] 🔴🔴🔴
                            # 엑셀 현재 행의 Code와 중요도 열을 읽어옵니다.
                            ext_code = str(get_val_robust(row, ['Code', '코드'])).strip()
                            ext_imp = str(get_val_robust(row, ['중요도', '등급', 'Importance'])).strip()
                            
                            is_already_analyzed = False
                            if skip_analyzed:
                                # 빈칸이나 "-"가 아니라 진짜 데이터가 하나라도 들어있다면 스킵 대상으로 판정합니다.
                                if (ext_code not in ["-", "", "nan"]) or (ext_imp not in ["-", "", "nan"]):
                                    is_already_analyzed = True

                            if is_already_analyzed:
                                # AI를 호출하지 않고, 기존 엑셀에 있던 값을 그대로 유지하며 화면 로그만 찍고 넘어갑니다.
                                st.session_state.analysis_logs.append({
                                    "행": real_row, "상태": "⏭️ 생략", 
                                    "Lv1": "-", "Lv2": "-", "Lv3": "-", "Lv4": "-", "Lv5": "-", 
                                    "Code": ext_code, "중요도": ext_imp, 
                                    "코멘트": "엑셀에 이미 분석 결과가 기재되어 있어 건너뜁니다.", "평가근거": "-"
                                })
                                st.session_state.processed_indices.add(df_idx)
                                continue  # 👈 [핵심] 이 줄을 만나면 아래 코드를 전부 무시하고 다음 특허 행으로 휙 넘어갑니다!
                            # 🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴🔴

                            # 2. base_data 준비 (이 아래부터는 기존 코드 그대로 유지)
                            base_data = {



