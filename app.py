from flask import Flask, render_template, request, jsonify, Response, stream_with_context
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
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

# 데이터 정의 (생략 - 기존과 동일)
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
    
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1440,900")
    chrome_options.binary_location = "/usr/bin/google-chrome"
    
    # 한국어 설정 강제 적용
    chrome_options.add_argument("--lang=ko_KR")
    chrome_options.add_experimental_option("prefs", {"intl.accept_languages": "ko,ko_KR"})
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.get('https://ims.dbhitek.com/')
        
        # 방문 신청 버튼 클릭
        wait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[onclick="popRegVisit();"]'))).click()
        
        # 필수 동의 (대기 시간 강화)
        wait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[name="ckAgree1"]'))).click()
        
        # 위치 선택 (텍스트가 아닌 '포함하는 글자'로 선택하도록 개선)
        location_el = wait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'Location')))
        location_select = Select(location_el)
        target_campus = '부천캠퍼스' if customer == '안현진' else '상우캠퍼스'
        
        found = False
        for option in location_select.options:
            if target_campus in option.text:
                location_select.select_by_visible_text(option.text)
                found = True
                break
        if not found:
            raise Exception(f"캠퍼스 목록에서 '{target_campus}'를 찾을 수 없습니다.")

        # 날짜 입력
        date_input = driver.find_element(By.CSS_SELECTOR, 'input[name="VisitStartDate"]')
        date_input.clear()
        date_input.send_keys(date_str)
        date_input.send_keys(Keys.ENTER)

        # 담당자 검색
        driver.find_element(By.CSS_SELECTOR, 'input[name="ContactName"]').click()
        time.sleep(2)
        cust_input = wait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="Name"]')))
        cust_input.send_keys(customer)
        cust_input.send_keys(Keys.ENTER)
        
        # 검색 결과 대기 및 클릭
        wait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#dvSearchPersonList > tbody > tr'))).click()
        time.sleep(1)

        # 세부 장소 선택 (개선된 방식 적용)
        place_el = wait(driver, 15).until(EC.presence_of_element_located((By.NAME, 'PlaceCodeID')))
        place_select = Select(place_el)
        target_place = '부천캠퍼스 FAB동' if customer == '안현진' else '상우캠퍼스 어드민동'
        
        found_place = False
        for opt in place_select.options:
            if target_place in opt.text:
                place_select.select_by_visible_text(opt.text)
                found_place = True
                break
        if not found_place:
            place_select.select_by_index(1) # 못 찾으면 첫 번째 항목이라도 선택

        Select(driver.find_element(By.NAME, 'VisitPurposeCodeID')).select_by_visible_text('공사/수리/Setup')
        driver.find_element(By.CSS_SELECTOR, 'input[name="Name[]"]').send_keys(applicant['name'])

        # 휴대물품
        button = driver.find_element(By.CSS_SELECTOR, 'button.btn-green[onclick="goCarryItem(this);"]')
        driver.execute_script("arguments[0].click();", button)
        time.sleep(2)
        
        Select(wait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'ImportPurposeCodeID')))).select_by_visible_text('기타')
        Select(driver.find_element(By.NAME, 'CarryItemCodeID')).select_by_visible_text('노트북 및 PC')
        driver.find_element(By.CSS_SELECTOR, 'input[name="ItemName"]').send_keys('노트북')
        driver.find_element(By.CSS_SELECTOR, 'input[name="ItemSN"]').send_keys(applicant['SN'])
        driver.find_element(By.CSS_SELECTOR, 'input[name="Quantity"]').send_keys('1')

        for idx, sn in enumerate(serial_nums):
            driver.find_element(By.CSS_SELECTOR, '#btn-add-carryitem').click()
            time.sleep(0.5)
            driver.find_elements(By.XPATH, "//*[@id='reg-form-wrap-carryitem']/ul/li[2]/div[2]/input")[idx+1].send_keys('노트북')
            driver.find_elements(By.XPATH, "//*[@id='reg-form-wrap-carryitem']/ul/li[3]/div[2]/input")[idx+1].send_keys(sn)
            driver.find_elements(By.XPATH, "//*[@id='reg-form-wrap-carryitem']/ul/li[6]/div[2]/input")[idx+1].send_keys('1')

        driver.find_element(By.CLASS_NAME, "pop-btn-green").click()
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR, 'input[type="checkbox"][name="ckAgree2"][value="Y"]').click()
        driver.find_element(By.CSS_SELECTOR, 'button.btn-req[onclick="saveVisitForm()"]').click()
        
        try:
            wait(driver, 5).until(EC.alert_is_present())
            driver.switch_to.alert.accept()
        except: pass
        
        # Teams 알림
        myTeamsMessage = pymsteams.connectorcard(TEAMS_URL)
        myTeamsMessage.text(f"[{date_str}] 신청 완료\n방문객: {applicant['name'][:3]}\n담당자: {customer}")
        myTeamsMessage.send()
        return True
    except Exception as e:
        print(f"FAIL: {date_str} - {str(e)}")
        return False
    finally:
        if 'driver' in locals(): driver.quit()

@app.route('/')
def index():
    dates = []
    now = datetime.datetime.now()
    weekday_map = ['(월)', '(화)', '(수)', '(목)', '(금)', '(토)', '(일)']
    for i in range(14):
        target = now + timedelta(days=i)
        dates.append({'val': target.strftime('%Y-%m-%d'), 'label': target.strftime('%Y-%m-%d') + " " + weekday_map[target.weekday()]})
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
    tasks = [{'date': d, 'applicant': applicant, 'customer': customer, 'serial_nums': serial_nums} for d in selected_dates]
    
    def generate():
        total = len(tasks)
        completed = 0
        yield json.dumps({'progress': 5, 'message': '엔진 가동 중...'}) + "\n"
        
        # Render 무료 플랜 안정성을 위해 순차적으로 처리 (1개씩)
        for task in tasks:
            success = run_selenium_task(task)
            completed += 1
            progress = int((completed / total) * 100)
            status = "✅ 성공" if success else "❌ 실패 (로그 확인)"
            yield json.dumps({'progress': progress, 'message': f'[{task["date"]}] {status} ({completed}/{total})'}) + "\n"

    return Response(stream_with_context(generate()), mimetype='application/json')

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
