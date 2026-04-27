from flask import Flask, render_template, request, jsonify, Response, stream_with_context
import time
import json
import os
import datetime
from datetime import timedelta
from playwright.sync_api import sync_playwright
import pymsteams

app = Flask(__name__)

# 데이터 정의 (기존과 동일)
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
TEAMS_URL = "https://asml.webhook.office.com/webhookb2/345970d0-d1e6-4432-b01a-b0018f62458d@af73baa8-f594-4eb2-a39d-93e96cad61fc/IncomingWebhook/0f1ede623d814ee6a9fe2175299cc5aa/263ad6e2-5556-4f9c-a7e1-bbd25a18ca4e/V2-w_v_LbU_NMXQxLYq0iv3sQj92pAcolXIygxRJ8zBiY1"

def run_visit_task(data):
    date_str = data['date']
    applicant = data['applicant']
    customer = data['customer']
    serial_nums = data['serial_nums']

    with sync_playwright() as p:
        # 초경량 모드로 브라우저 실행
        browser = p.chromium.launch(headless=True, args=[
            '--no-sandbox', 
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage', # 메모리 부족 방지 핵심 설정
            '--single-process'         # 메모리 사용량 최소화
        ])
        context = browser.new_context(locale='ko-KR')
        page = context.new_page()
        page.route("**/*.{png,jpg,jpeg,gif,svg,css,woff2}", lambda route: route.abort()) # CSS까지 차단하여 속도 극대화

        try:
            page.goto('https://ims.dbhitek.com/', timeout=60000)
            page.click('button[onclick="popRegVisit();"]')
            page.wait_for_selector('input[name="ckAgree1"]', timeout=30000)
            page.click('input[name="ckAgree1"]')
            
            target_campus = '부천캠퍼스' if customer == '안현진' else '상우캠퍼스'
            page.select_option('select[name="Location"]', label=f'DB HiTek {target_campus}')
            page.fill('input[name="VisitStartDate"]', date_str)
            page.keyboard.press('Enter')

            page.click('input[name="ContactName"]')
            page.wait_for_selector('input[name="Name"]')
            page.fill('input[name="Name"]', customer)
            page.keyboard.press('Enter')
            page.wait_for_selector('#dvSearchPersonList > tbody > tr', timeout=30000)
            page.click('#dvSearchPersonList > tbody > tr')

            page.wait_for_selector('select[name="PlaceCodeID"]')
            target_place = '부천캠퍼스 FAB동' if customer == '안현진' else '상우캠퍼스 어드민동'
            page.select_option('select[name="PlaceCodeID"]', label=target_place)
            page.select_option('select[name="VisitPurposeCodeID"]', label='공사/수리/Setup')
            page.fill('input[name="Name[]"]', applicant['name'])

            page.click('button.btn-green[onclick="goCarryItem(this);"]')
            page.wait_for_selector('select[name="ImportPurposeCodeID"]')
            page.select_option('select[name="ImportPurposeCodeID"]', label='기타')
            page.select_option('select[name="CarryItemCodeID"]', label='노트북 및 PC')
            page.fill('input[name="ItemName"]', '노트북')
            page.fill('input[name="ItemSN"]', applicant['SN'])
            page.fill('input[name="Quantity"]', '1')

            for sn in serial_nums:
                page.click('#btn-add-carryitem')
                page.locator('input[name="ItemName[]"]').last.fill('노트북')
                page.locator('input[name="ItemSN[]"]').last.fill(sn)
                page.locator('input[name="Quantity[]"]').last.fill('1')

            page.click('.pop-btn-green')
            page.click('input[name="ckAgree2"]')
            page.on("dialog", lambda dialog: dialog.accept())
            page.click('button.btn-req[onclick="saveVisitForm()"]')
            time.sleep(2)
            
            myTeamsMessage = pymsteams.connectorcard(TEAMS_URL)
            myTeamsMessage.text(f"[{date_str}] 신청 완료: {applicant['name'][:3]} -> {customer}")
            myTeamsMessage.send()
            return True
        except Exception as e:
            print(f"FAIL: {date_str} -> {str(e)}")
            return False
        finally:
            browser.close()

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
    try:
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
            yield json.dumps({'progress': 5, 'message': '엔진 시동 중...'}) + "\n"
            for i, task in enumerate(tasks):
                success = run_visit_task(task)
                progress = int(((i + 1) / total) * 100)
                status = "성공" if success else "실패"
                yield json.dumps({'progress': progress, 'message': f'[{task["date"]}] {status}'}) + "\n"

        return Response(stream_with_context(generate()), mimetype='application/json')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
