import requests
from bs4 import BeautifulSoup
import re


def format_to_google_patent_id(app_number):
    """일반 출원번호/등록번호를 Google Patents URL 포맷으로 변환합니다."""
    # 1. 특수문자 및 공백 제거
    clean_num = re.sub(r'[^a-zA-Z0-9]', '', str(app_number))
    
    # 2. 한국 특허(10으로 시작하는 13자리 이상 번호) 처리 로직
    # 이미 KR이 붙어있지 않다면 KR을 붙여줍니다.
    if clean_num.startswith("10") and not clean_num.upper().startswith("KR"):
        # 보통 공개특허는 A, 등록특허는 B1이 붙지만 구글 검색은 A만 붙여도 대부분 찾아줍니다.
        clean_num = f"KR{clean_num}A" 
        
    return clean_num

def get_google_patent_image(app_number):
    """Google Patents에서 특허 번호로 대표 이미지 URL을 스크래핑합니다."""
    patent_id = format_to_google_patent_id(app_number)
    url = f"https://patents.google.com/patent/{patent_id}/en"
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            # Google Patents는 보통 meta 태그에 대표 이미지를 박아둡니다.
            meta_img = soup.find('meta', attrs={'name': 'citation_image'})
            if meta_img and 'content' in meta_img.attrs:
                return meta_img['content']
            
            # meta 태그가 없다면 본문의 첫 번째 도면 이미지를 찾습니다.
            img_tag = soup.find('img', itemprop='image')
            if img_tag and 'src' in img_tag.attrs:
                return img_tag['src']
    except Exception as e:
        pass # 통신 에러 발생 시 None 반환
    
    return None # 이미지를 못 찾았을 경우










    # --- [메뉴 선택 로직 어딘가에 추가] ---
elif menu == "📑 주요 특허 리포트":
    st.title("📑 주요 특허 심층 리포트 (S/A 등급)")
    st.markdown("AI 센싱 결과 **S급** 및 **A급**으로 분류된 핵심 특허들의 요약 리포트를 확인합니다.")
    
    # 1. Library(DB) 선택
    # (예시: 저장된 Library 목록을 불러오는 로직. 실제 연구원님 환경에 맞게 수정하세요)
    # available_libraries = ["FeVNAND_Library.xlsx", "2D_Materials_Library.xlsx"]
    # selected_lib = st.selectbox("📂 분석할 Library(DB) 선택", available_libraries)
    
    # [임시] df는 선택된 Library에서 불러온 데이터프레임이라고 가정합니다.
    # df = pd.read_excel(selected_lib)
    
    if not df.empty:
        # 2. S급, A급 특허만 필터링
        # (실제 엑셀의 열 이름이 '중요도 등급'인지 '중요도'인지 확인 필요)
        target_grades = ['S', 'A']
        df_sa = df[df['중요도'].isin(target_grades)]
        
        if df_sa.empty:
            st.warning("선택한 Library에 S/A 등급 특허가 없습니다.")
        else:
            # 3. 특허 선택용 목록 만들기 (출원번호 - 제목 형태)
            df_sa['selectbox_label'] = df_sa['중요도'] + "등급 | " + df_sa['출원번호'].astype(str) + " - " + df_sa['특허제목']
            
            selected_label = st.selectbox("🔍 리포트를 확인할 특허를 선택하세요:", df_sa['selectbox_label'].tolist())
            
            st.markdown("---")
            
            # 선택된 특허의 데이터 추출
            patent_data = df_sa[df_sa['selectbox_label'] == selected_label].iloc[0]
            
            # 4. 리포트 레이아웃 구성 (좌측: 정보, 우측: 이미지)
            col_info, col_img = st.columns([6, 4])
            
            with col_info:
                st.subheader(f"[{patent_data.get('중요도', '-')}] {patent_data.get('특허제목', '제목 없음')}")
                st.caption(f"**출원번호:** {patent_data.get('출원번호', '-')} | **출원일자:** {patent_data.get('출원일자', '-')} | **출원인:** {patent_data.get('출원인', '-')}")
                
                # 검토 의견 및 평가 근거 (AI 센싱 결과)
                st.markdown("### 💡 AI 검토 의견 및 평가 근거")
                st.info(patent_data.get('평가근거', '평가 근거가 없습니다.'))
                st.markdown(f"**📝 코멘트:** {patent_data.get('코멘트', '-')}")
                
                # 대표 청구항 (데이터에 없을 경우를 대비해 get 사용)
                st.markdown("### 📜 대표 청구항")
                with st.expander("청구항 내용 보기", expanded=False):
                    st.write(patent_data.get('대표청구항', 'Library에 대표청구항 데이터가 존재하지 않습니다.'))

            with col_img:
                st.markdown("### 🖼️ 대표 도면")
                with st.spinner("Google Patents에서 대표 이미지를 불러오는 중..."):
                    app_num = patent_data.get('출원번호', '')
                    img_url = get_google_patent_image(app_num)
                    
                    if img_url:
                        # 이미지를 화면에 꽉 차게 출력
                        st.image(img_url, use_container_width=True, caption=f"Google Patents 발췌: {app_num}")
                        # 원본 링크로 갈 수 있는 버튼
                        google_link = f"https://patents.google.com/patent/{format_to_google_patent_id(app_num)}/en"
                        st.link_button("🌐 Google Patents 원본 보기", google_link)
                    else:
                        st.warning("Google Patents에서 이미지를 가져올 수 없거나 도면이 없는 특허입니다.")
                        # 구글 특허 검색 링크라도 제공
                        search_link = f"https://patents.google.com/search?q={app_num}"
                        st.link_button("🔍 Google Patents 수동 검색", search_link)