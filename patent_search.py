from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

def get_filtered_patents_data(query, company_name, num_results):
    # 크롬 드라이버 경로 설정
    service = Service(executable_path='./chromedriver')
    driver = webdriver.Chrome()
    
    
 # 구글 특허 검색 URL 구성
    url = f"https://patents.google.com/?q={query}&assignee={company_name}&num={num_results}"
    driver.get(url)
    
    patents = []
    
    while True:
        try:
            # 페이지 로딩 대기 (최대 30초 대기)
            WebDriverWait(driver, 30).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'article.search-result-item'))
            )
            
            # 특허 제목 및 링크 추출
            results = driver.find_elements(By.CSS_SELECTOR, 'article.search-result-item')
            
            found_query_in_titles = False
            
            for item in results:
                try:
                    # 'a' 태그가 포함된 요소를 찾기
                    link_element = item.find_element(By.CSS_SELECTOR, 'a')
                    title = link_element.find_element(By.CSS_SELECTOR, 'h3').text.strip()
                    link = link_element.get_attribute('href')
                    
                    if query.lower() in title.lower():  # 제목에 query가 포함되는지 확인
                        patents.append({'Title': title, 'Link': link})
                        found_query_in_titles = True
                    
                except Exception as e:
                    print(f"Error extracting data from item: {e}")

            # '다음 페이지' 버튼 클릭
            if found_query_in_titles:  # query가 포함된 제목이 있는 경우에만 다음 페이지로 이동
                try:
                    next_button = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'paper-icon-button[icon="chevron-right"]'))
                    )
                    next_button.click()
                    time.sleep(2)  # 페이지가 로드될 때까지 잠시 대기
                except Exception as e:
                    print(f"Error or no more pages: {e}")
                    break  # 다음 페이지가 없으면 반복 종료
            else:
                print("No more results containing the query.")
                break  # 제목에 query가 포함되지 않은 경우 반복 종료

        except Exception as e:
            print(f"Error during page load or element search: {e}")
            break

    driver.quit()
    
    return patents

def save_to_excel(data, filename):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)

if __name__ == "__main__":
    query = "다공성 실리콘"
    company_name = "삼성전자주식회사"  # 검색할 회사명으로 대체
    num_results = 100  # 가져올 최대 검색 결과 수
    
    patents_data = get_filtered_patents_data(query, company_name, num_results)

    print(patents_data)

    save_to_excel(patents_data, 'filtered_patents.xlsx')