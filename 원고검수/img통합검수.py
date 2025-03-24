import streamlit as st
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from docx import Document
from datetime import datetime
import io
import base64

def check_spelling(text):
    """네이버 맞춤법 검사기 사용"""
    driver = None
    try:
        # Chrome 드라이버 초기화 (headless 모드)
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')  # 백그라운드 실행
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        driver = webdriver.Chrome(options=options)
        
        # 네이버 맞춤법 검사기 페이지 열기
        driver.get("https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=0&ie=utf8&query=맞춤법+검사기")
        
        # 텍스트 입력창 대기
        textarea = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="grammar_checker"]/div/div/div[2]/div[1]/div/div[1]/textarea'))
        )
        
        # 텍스트 입력
        textarea.clear()
        textarea.send_keys(text)
        
        # 검사하기 버튼 클릭
        check_button = driver.find_element(By.XPATH, '//*[@id="grammar_checker"]/div/div/div[2]/div[1]/div/div[2]/button')
        check_button.click()
        
        # 결과가 나올 때까지 대기
        time.sleep(5)
        
        # 결과 저장 - XPath 사용
        result = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="grammar_checker"]/div/div/div[2]/div[2]/div/div/div[2]/div'))
        ).text
        
        return result, None
        
    except Exception as e:
        return None, str(e)
    
    finally:
        if driver:
            driver.quit()

def create_word_document(original_text, corrected_text):
    """워드 문서 생성"""
    doc = Document()
    
    # 제목 추가
    doc.add_heading('맞춤법 검사 결과', 0)
    
    # 원본 텍스트 추가
    doc.add_heading('원본 텍스트:', level=1)
    doc.add_paragraph(original_text)
    
    # 교정된 텍스트 추가
    doc.add_heading('교정된 텍스트:', level=1)
    doc.add_paragraph(corrected_text)
    
    # 현재 시간 추가
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    doc.add_paragraph(f'\n검사 시간: {now}')
    
    # 메모리에 저장
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    
    return doc_io

def main():
    st.title('네이버 맞춤법 검사기')
    
    # 파일 업로드
    uploaded_file = st.file_uploader("텍스트 파일 선택 (.txt)", type=['txt'])
    
    if uploaded_file:
        text = uploaded_file.read().decode('utf-8')
        st.text_area("원본 텍스트", text, height=200)
        
        if st.button('맞춤법 검사 시작'):
            with st.spinner('맞춤법 검사 중...'):
                result, error = check_spelling(text)
                
                if result:
                    st.success('맞춤법 검사가 완료되었습니다!')
                    st.text_area("검사 결과", result, height=200)
                    
                    # 워드 파일 생성
                    doc_io = create_word_document(text, result)
                    
                    # 다운로드 버튼
                    st.download_button(
                        label="결과 다운로드 (DOCX)",
                        data=doc_io.getvalue(),
                        file_name=f'맞춤법검사결과_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx',
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    )
                else:
                    st.error(f'오류가 발생했습니다: {error}')

if __name__ == '__main__':
    main() 
