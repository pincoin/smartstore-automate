import glob
import os
import time

import openpyxl
from selenium import webdriver

from secret import *

if __name__ == '__main__':
    options = webdriver.ChromeOptions()
    options.add_experimental_option('prefs', {
        'download.default_directory': download_path
    })
    driver = webdriver.Chrome(executable_path=executable_path, chrome_options=options)

    driver.implicitly_wait(3)

    driver.get('https://nid.naver.com/nidlogin.login')

    driver.execute_script('document.getElementsByName("id")[0].value="' + username + '"')
    time.sleep(1)
    driver.execute_script('document.getElementsByName("pw")[0].value="' + password + '"')
    time.sleep(1)

    driver.find_element_by_xpath('//*[@id="frmNIDLogin"]/fieldset/input').click()
    print('로그인 완료')

    driver.find_element_by_class_name('btn_cancel').click()
    print('새로운 기기 등록 안함')
    time.sleep(1)

    ###
    # TODO: 5분 마다 반복 스레드 처리

    driver.get('https://sell.smartstore.naver.com/#/naverpay/sale/delivery?summaryInfoType=DELIVERY_READY')
    print('스마트스토어 이동')
    time.sleep(3)

    driver.switch_to.frame(driver.find_element_by_id('__naverpay'))
    print('iframe 선택')

    driver.find_element_by_xpath('//select[@name="orderStatus"]/option[text()="발송대기"]').click()
    print('발송대기 선택')

    # 엑셀 다운로드
    driver.find_element_by_class_name('_excelDownloadBtn').click()
    print('엑셀 다운로드')
    time.sleep(5)

    # 엑셀 열기
    files = [f for f in glob.glob(os.path.join('{}/*.xlsx'.format(download_path))) if os.path.isfile(f)]

    if len(files) == 1 and files[0].endswith('.xlsx'):
        wb_order = openpyxl.load_workbook(files[0])
        ws_order = wb_order.active

        wb_batch = openpyxl.Workbook()
        ws_batch = wb_batch.active
        ws_batch.title = '발송처리'
        ws_batch.cell(row=1, column=1, value='상품주문번호')
        ws_batch.cell(row=1, column=2, value='배송방법')
        ws_batch.cell(row=1, column=3, value='택배사')
        ws_batch.cell(row=1, column=4, value='송장번호')

        order_count = 1

        for row in ws_order.rows:
            if row[0].row > 2:
                order_count += 1

                # 구매자명, 판매자 상품코드, 수량, 수취인연락처1, 상품주문번호
                customer = row[0].value
                product = row[2].value
                quantity = int(row[5].value)
                phone = row[1].value
                order = row[17].value

                # TODO: REST API 비동기 발송 처리

                # 상품주문번호, 배송방법(배송없음), 택배사(공란), 송장번호(공란)
                ws_batch.cell(row=1 + order_count, column=1, value=order)
                ws_batch.cell(row=1 + order_count, column=2, value='배송없음')
                ws_batch.cell(row=1 + order_count, column=3, value='')
                ws_batch.cell(row=1 + order_count, column=4, value='')

        if order_count > 0:
            wb_batch.save(os.path.join(download_path, batch_excel))
            wb_batch.close()
            print('일괄발송 엑셀 파일 저장')

            # TODO: 일괄발송 엑셀 업로드: 발송상태 변경 처리
            parent = driver.current_window_handle
            child = None

            driver.find_element_by_xpath('//button[text()="엑셀 일괄발송"]').click()

            for handle in driver.window_handles:
                print(handle)
                if handle != parent:
                    child = handle

            driver.switch_to.window(child)

            # driver.find_element_by_id('uploadedFile').sendKeys(os.path.join(download_path, batch_excel))

            driver.switch_to.window(parent)

            os.remove(os.path.join(download_path, batch_excel))
            print('일괄발송 엑셀 로컬 삭제')

        wb_order.close()
        os.remove(files[0])
        print('엑셀 로컬 삭제')
