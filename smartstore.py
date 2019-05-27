import glob
import json
import os
import time

import openpyxl
import requests
import xlwt
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException

from secret import *

if __name__ == '__main__':
    options = webdriver.ChromeOptions()
    options.add_argument('user-data-dir={}'.format(chrome_profile_path))
    options.add_argument('profile-directory=Default')

    if chrome_headless:
        options.add_argument('disable-extensions')
        options.add_argument('headless')
        options.add_argument('disable-gpu')
        options.add_argument('no-sandbox')
        options.add_argument('window-size=1920x1080')

    options.add_experimental_option('prefs', {'download.default_directory': download_path})

    driver = webdriver.Chrome(executable_path=executable_path, chrome_options=options)

    driver.implicitly_wait(3)

    driver.get('https://nid.naver.com/nidlogin.login')

    driver.execute_script('document.getElementsByName("id")[0].value="{}"'.format(username))
    time.sleep(1)
    driver.execute_script('document.getElementsByName("pw")[0].value="{}"'.format(password))
    time.sleep(1)

    driver.find_element_by_xpath('//*[@id="frmNIDLogin"]/fieldset/input').click()
    print('로그인 완료')

    try:
        driver.find_element_by_class_name('btn_cancel').click()
        print('새로운 기기 등록 안함')
        time.sleep(2)
    except NoSuchElementException:
        pass

    driver.get('https://sell.smartstore.naver.com/#/naverpay/sale/delivery?summaryInfoType=DELIVERY_READY')
    print('스마트스토어 이동')
    time.sleep(3)

    driver.switch_to.frame(driver.find_element_by_id('__naverpay'))
    print('iframe 선택')
    time.sleep(3)

    driver.find_element_by_xpath('//select[@name="orderStatus"]/option[text()="발송대기"]').click()
    print('발송대기 선택')
    time.sleep(3)

    # 엑셀 다운로드
    driver.find_element_by_class_name('_excelDownloadBtn').click()
    print('엑셀 다운로드')
    time.sleep(5)

    # 엑셀 열기
    files = [f for f in glob.glob(os.path.join('{}/*.xlsx'.format(download_path))) if os.path.isfile(f)]

    if len(files) == 1 and files[0].endswith('.xlsx'):
        wb_order = openpyxl.load_workbook(files[0])
        ws_order = wb_order.active

        wb_batch = xlwt.Workbook()
        ws_batch = wb_batch.add_sheet('발송처리')
        ws_batch.write(0, 0, '상품주문번호')
        ws_batch.write(0, 1, '배송방법')
        ws_batch.write(0, 2, '택배사')
        ws_batch.write(0, 3, '송장번호')

        order_count = 0

        for row in ws_order.rows:
            if row[0].row > 2:
                order_count += 1

                # TODO: REST API 비동기 발송 처리

                # 구매자명, 판매자 상품코드, 수량, 수취인연락처1, 상품주문번호
                payload = {
                    'customer': row[0].value,
                    'product': int(row[2].value),
                    'quantity': int(row[5].value),
                    'phone': row[1].value,
                    'order': row[17].value
                }
                headers = {
                    'Content-Type': 'application/json',
                    'Cache-Control': 'no-cache',
                    'Authorization': 'Token {}'.format(api_token),
                }
                print(payload)
                requests.post(api_url, data=json.dumps(payload), headers=headers)

                # 상품주문번호, 배송방법(배송없음), 택배사(공란), 송장번호(공란)
                ws_batch.write(order_count, 0, row[17].value)
                ws_batch.write(order_count, 1, '직접전달')
                ws_batch.write(order_count, 2, '')
                ws_batch.write(order_count, 3, '')

        if order_count > 0:
            wb_batch.save(os.path.join(download_path, batch_excel))
            print('일괄발송 엑셀 파일 저장')

            parent = driver.current_window_handle
            child = None

            driver.find_element_by_xpath('//button[text()="엑셀 일괄발송"]').click()
            print('일괄발송 엑셀 팝업 띄우기')
            time.sleep(3)

            for handle in driver.window_handles:
                if handle != parent:
                    child = handle

            driver.switch_to.window(child)
            time.sleep(3)

            driver.find_element_by_xpath('//input[@name="uploadedFile"]') \
                .send_keys(os.path.join(download_path, batch_excel))
            print('일괄발송 엑셀 파일 선택')
            time.sleep(5)

            driver.find_element_by_xpath('//span[text()="일괄 발송처리"]').click()
            print('일괄발송 엑셀 파일 업로드')
            time.sleep(3)

            driver.find_element_by_xpath('//span[text()="닫기"]').click()
            print('일괄발송 엑셀 팝업 닫기')
            time.sleep(2)

            driver.switch_to.window(parent)
            time.sleep(3)

            os.remove(os.path.join(download_path, batch_excel))
            print('일괄발송 엑셀 로컬 삭제')

        wb_order.close()
        os.remove(files[0])
        print('엑셀 로컬 삭제')

    driver.quit()
