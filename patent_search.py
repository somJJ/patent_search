import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

def get_filtered_patents_data(query, company_name, num_results, company_verification, mode):
    service = Service(executable_path='./chromedriver')
    driver = webdriver.Chrome()

    url = f"https://patents.google.com/?q={query}&assignee={company_name}&num={num_results}"
    driver.get(url)

    patents = []
    
    while True:
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'article.search-result-item'))
            )

            results = driver.find_elements(By.CSS_SELECTOR, 'article.search-result-item')

            for item in results:
                try:
                    span_elements = item.find_elements(By.CSS_SELECTOR, 'span#htmlContent')
                    
                    title = span_elements[0].text.strip() if len(span_elements) > 1 else "N/A"
                    company = span_elements[2].text.strip() if len(span_elements) > 0 else "N/A"
                    content = span_elements[3].text.strip() if len(span_elements) > 0 else "N/A"

                    # 링크 추출
                    # 특허 번호 추출해서 링크 주소 만드는 방식으로 함
                    patent_number_element = item.find_element(By.CSS_SELECTOR, 'span[data-proto="OPEN_PATENT_PDF"]')
                    patent_number = patent_number_element.text.strip()

                    # 전체 링크 생성
                    link = f"https://patents.google.com/patent/{patent_number}"

                    if mode == "titleonly":
                        if query.lower() in title.lower() and company_verification.lower() in company.lower():
                            patents.append({'Company': company, 'Title': title, 'Link': link})
                    else:
                        if company_verification.lower() in company.lower(): 
                            if query.lower() in title.lower() or query.lower() in content.lower():
                                patents.append({'Company': company, 'Title': title, 'Link': link, 'Content': content})

                except Exception as e:
                    print(f"Error extracting data from item: {e}")

            try:
                next_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'paper-icon-button[icon="chevron-right"]'))
                )
                next_button.click()
                time.sleep(2)
            except Exception as e:
                print(f"No more pages or error: {e}")
                break

        except Exception as e:
            print(f"Error during page load or element search: {e}")
            break

    driver.quit()
    
    return patents

def save_to_excel(data, filename):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)

def run_crawler():
    query = query_entry.get()
    company_name = company_entry.get()
    company_verification = verification_entry.get()
    num_results = 100
    mode = mode_var.get()

    if not query or not company_name or not company_verification or not num_results or not mode:
        messagebox.showwarning("Input Error", "모든 필드를 입력하세요.")
        return

    try:
        patents_data = get_filtered_patents_data(query, company_name, num_results, company_verification, mode)
        save_to_excel(patents_data, 'filtered_patents.xlsx')
        messagebox.showinfo("Success", "데이터 수집이 완료되었습니다. 'filtered_patents.xlsx' 파일이 생성되었습니다.")
    except Exception as e:
        messagebox.showerror("Error", f"오류 발생: {e}")

# GUI 생성
root = tk.Tk()
root.title("Patent Search")
root.geometry("640x200+300+300")

# Query 입력
tk.Label(root, text="Search Word:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
query_entry = tk.Entry(root, width=40)
query_entry.grid(row=0, column=1, padx=10, pady=5)

# Company Name 입력
tk.Label(root, text="Company Name:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
company_entry = tk.Entry(root, width=40)
company_entry.grid(row=1, column=1, padx=10, pady=5)

# Company Verification 입력
tk.Label(root, text="Company Verification:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
verification_entry = tk.Entry(root, width=40)
verification_entry.grid(row=2, column=1, padx=10, pady=5)


# Mode 선택
tk.Label(root, text="Mode:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
mode_var = tk.StringVar(value="titleonly")
mode_titleonly = tk.Radiobutton(root, text="Title Only", variable=mode_var, value="titleonly")
mode_includingContent = tk.Radiobutton(root, text="Including Content", variable=mode_var, value="includingContent")
mode_titleonly.grid(row=4, column=1, sticky=tk.W, padx=10)
mode_includingContent.grid(row=4, column=1, padx=100)

# 실행 버튼
run_button = tk.Button(root, text="START", command=run_crawler)
run_button.grid(row=5, column=1, pady=20)

# GUI 실행
root.mainloop()
