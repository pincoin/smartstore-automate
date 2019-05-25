import time

from selenium import webdriver

from .secret import *

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

    # 로그인
    driver.find_element_by_xpath('//*[@id="frmNIDLogin"]/fieldset/input').click()

    driver.find_element_by_class_name('btn_cancel').click()
    time.sleep(1)

    ###
    # 5분 마다 반복

    # 스마트스토어 이동
    driver.get('https://sell.smartstore.naver.com/#/naverpay/sale/delivery?summaryInfoType=DELIVERY_READY')
    time.sleep(3)

    # iframe 선택
    driver.switch_to.frame(driver.find_element_by_id('__naverpay'))

    # 발송대기 선택
    driver.find_element_by_xpath('//select[@name="orderStatus"]/option[text()="배송완료"]').click()

    # 엑셀 다운로드
    driver.find_element_by_class_name('_excelDownloadBtn').click()
    time.sleep(3)

    # 엑셀 파일 열기

    # 행 반복적으로 처리하기
    # 구매자명, 판매자 상품코드, 수량, 수취인연락처1, 상품주문번호

    # 발송 처리
    # API 호출

    # 엑셀 파일 저장 : 파일명 발송처리
    # 상품주문번호, 배송방법(배송없음), 택배사(공란), 송장번호(공란)

    # 엑셀 파일 업로드

    # 엑셀 파일 삭제
