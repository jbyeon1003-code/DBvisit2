from flask import Flask, render_template, request, jsonify
import time
from multiprocessing import Pool
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC
import os
import datetime
from datetime import timedelta
import pymsteams

app = Flask(__name__)

TEAMS_URL = "https://asml.webhook.office.com/webhookb2/345970d0-d1e6-4432-b01a-b0018f62458d@af73baa8-f594-4eb2-a39d-93e96cad61fc/IncomingWebhook/0f1ede623d814ee6a9fe2175299cc5aa/263ad6e2-5556-4f9c-a7e1-bbd25a18ca4e/V2-w_v_LbU_NMXQxLYq0iv3sQj92pAcolXIygxRJ8zBiY1"

# 데이터 정의
APPLICANTS = {
    '한준석': {'name': '한준석\t19770103', 'SN': 'PF32SF80(JS)'},
    '최선국': {'name': '최선국\t19790503', 'SN': 'PF35R8Y6(SK)'},
    '조영택': {'name': '조영택\t19850120', 'SN': 'PF410726(YT)'},
    '신영재': {'name': '신영재\t19850405', 'SN': 'PF46LP83(YJ)'},
    '이영훈': {'name': '이영훈\t19880620', 'SN': 'PF2RW6F4(YH)'},
    '강정규': {'name': '강정규\t19900307', 'SN': 'PF3AWQP1(JG)'},
    '변정호': {'name': '변정호\t19900812', 'SN': 'PF46CER0(JH)'},
    '김시준': {'name': '김시준\t19920108', 'SN': 'PF46FEY5(SJ)'},
    '권이건': {'name': '권이건\t19941210', 'SN': 'PF3YV1PX(LK)'},
    '장덕수': {'name': '장덕수\t19960414', 'SN': '5CG4255D3R(DS)'},
    '백한빈': {'name': '백한빈\t19960831', 'SN': 'PF3YAA2F(HB)'},
    '이수한': {'name': '이수한\t19980309', 'SN': 'PF4SLTF9(SH)'},
    '박찬순': {'name': '박찬순\t19990407', 'SN': 'PF4SLTEF(CS)'}
}

CUSTOMERS = ['채명주', '이영휘', '서형석', '윤여철', '박종우', '안현진']

def run_selenium_task(data):
    date_str = data['date']
    applicant = data['applicant']
    customer = data['customer']
    serial_nums = data['serial_nums']
    
    options = Options()
    options.add_experimental_option('detach', True)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("lang=ko")
    options.add_argument('--disable-popup-blocking')
    options.page_load_strategy = 'eager'
    
    os.environ["SE_DRIVER_MIRROR_URL"] = "https://msedgedriver.microsoft.com"
    driver = webdriver.Edge(options=options)
    
    try:
        url = 'https://ims.dbhitek.com/'
        driver.set_window_size(1440, 900)
        driver.get(url)

        # 방문 신청 버튼
        wait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[onclick="popRegVisit();"]'))).click()

        # 필수 동의
        driver.find_element(By.CSS_SELECTOR, 'input[type="checkbox"][name="ckAgree1"][value="Y"]').click()
        
        # 위치 선택
        location_select = Select(driver.find_element(By.NAME, 'Location'))
        if customer == '안현진':
            location_select.select_by_visible_text('DB HiTek 부천캠퍼스')
        else:
            location_select.select_by_visible_text('DB HiTek 상우캠퍼스')

        # 날짜 입력
        date_input = driver.find_element(By.CSS_SELECTOR, 'input[name="VisitStartDate"]')
        date_input.clear()
        date_input.send_keys(date_str)
        date_input.send_keys(Keys.ENTER)

        # 담당자 검색
        driver.find_element(By.CSS_SELECTOR, 'input[name="ContactName"]').click()
        time.sleep(1)
        cust_input = driver.find_element(By.CSS_SELECTOR, 'input[name="Name"]')
        cust_input.send_keys(customer)
        cust_input.send_keys(Keys.ENTER)
        
        wait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#dvSearchPersonList > tbody > tr'))).click()

        # 세부 장소 및 목적
        place_select = Select(driver.find_element(By.NAME, 'PlaceCodeID'))
        if customer == '안현진':
            place_select.select_by_visible_text('부천캠퍼스 FAB동')
        else:
            place_select.select_by_visible_text('상우캠퍼스 어드민동')

        Select(driver.find_element(By.NAME, 'VisitPurposeCodeID')).select_by_visible_text('공사/수리/Setup')

        # 신청자 입력
        driver.find_element(By.CSS_SELECTOR, 'input[name="Name[]"]').send_keys(applicant['name'])

        # 휴대물품 등록
        button = driver.find_element(By.CSS_SELECTOR, 'button.btn-green[onclick="goCarryItem(this);"]')
        driver.execute_script("arguments[0].click();", button)
        time.sleep(1)

        Select(driver.find_element(By.NAME, 'ImportPurposeCodeID')).select_by_visible_text('기타')
        Select(driver.find_element(By.NAME, 'CarryItemCodeID')).select_by_visible_text('노트북 및 PC')

        driver.find_element(By.CSS_SELECTOR, 'input[name="ItemName"]').send_keys('노트북')
        driver.find_element(By.CSS_SELECTOR, 'input[name="ItemSN"]').send_keys(applicant['SN'])
        driver.find_element(By.CSS_SELECTOR, 'input[name="Quantity"]').send_keys('1')

        # 추가 노트북
        for idx, sn in enumerate(serial_nums):
            driver.find_element(By.CSS_SELECTOR, '#btn-add-carryitem').click()
            time.sleep(0.5)
            add_items = driver.find_elements(By.XPATH, "//*[@id='reg-form-wrap-carryitem']/ul/li[2]/div[2]/input")
            add_sns = driver.find_elements(By.XPATH, "//*[@id='reg-form-wrap-carryitem']/ul/li[3]/div[2]/input")
            add_nums = driver.find_elements(By.XPATH, "//*[@id='reg-form-wrap-carryitem']/ul/li[6]/div[2]/input")
            
            add_items[idx+1].send_keys('노트북')
            add_sns[idx+1].send_keys(sn)
            add_nums[idx+1].send_keys('1')

        driver.find_element(By.CLASS_NAME, "pop-btn-green").click()
        time.sleep(1)

        # 선택 동의 및 신청
        driver.find_element(By.CSS_SELECTOR, 'input[type="checkbox"][name="ckAgree2"][value="Y"]').click()
        driver.find_element(By.CSS_SELECTOR, 'button.btn-req[onclick="saveVisitForm()"]').click()
        
        Alert(driver).accept()
        
        # Teams 알림
        myTeamsMessage = pymsteams.connectorcard(TEAMS_URL)
        myTeamsMessage.text(f"{date_str} 신청 완료. 방문객: {applicant['name'][:3]}, 담당자: {customer}, 추가: {serial_nums}")
        myTeamsMessage.send()
        
        return True
    except Exception as e:
        print(f"Error: {e}")
        return False
    finally:
        driver.quit()

@app.route('/')
def index():
    # 다음 14일간의 날짜 리스트 생성
    dates = []
    now = datetime.datetime.now()
    weekday_map = ['(월)', '(화)', '(수)', '(목)', '(금)', '(토)', '(일)']
    for i in range(14):
        target = now + timedelta(days=i)
        dates.append({
            'val': target.strftime('%Y-%m-%d'),
            'label': target.strftime('%Y-%m-%d') + " " + weekday_map[target.weekday()]
        })
    return render_template('index.html', applicants=APPLICANTS, customers=CUSTOMERS, dates=dates)

@app.route('/apply', methods=['POST'])
def apply():
    form_data = request.json
    selected_dates = form_data.get('dates', [])
    applicant_key = form_data.get('applicant')
    customer = form_data.get('customer')
    colleagues = form_data.get('colleagues', [])
    
    applicant = APPLICANTS.get(applicant_key)
    serial_nums = [APPLICANTS[c]['SN'] for c in colleagues if c in APPLICANTS]
    
    tasks = []
    for d in selected_dates:
        tasks.append({
            'date': d,
            'applicant': applicant,
            'customer': customer,
            'serial_nums': serial_nums
        })
    
    # 멀티프로세싱 실행
    with Pool(processes=min(len(tasks), 4)) as pool:
        results = pool.map(run_selenium_task, tasks)
    
    return jsonify({'status': 'success', 'results': results})

if __name__ == '__main__':
    app.run(debug=True, port=5000)
