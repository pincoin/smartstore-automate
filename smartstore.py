import argparse
import glob
import json
import logging
import os
import signal
import sys
import time

import openpyxl
import requests
import xlwt
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException

from secret import *


class SmartStore:
    def __init__(self, log_file=None):
        logging.basicConfig(level=logging.INFO, format='%(message)s')
        self.logger = logging.getLogger('SmartStore')
        self.log_file = log_file

        if log_file:
            self.log_handler = logging.FileHandler(self.log_file)
            self.logger.addHandler(self.log_handler)

        self.__stop = False

        signal.signal(signal.SIGINT, self.stop)
        signal.signal(signal.SIGTERM, self.stop)

        options = webdriver.ChromeOptions()
        options.add_argument('user-data-dir={}'.format(chrome_profile_path))
        options.add_argument('profile-directory=Default')

        if chrome_headless:
            options.add_argument('disable-extensions')
            options.add_argument('disable-popup-blocking')
            options.add_argument('headless')
            options.add_argument('disable-gpu')
            options.add_argument('no-sandbox')
            options.add_argument('window-size=1920x1080')

        options.add_experimental_option('prefs', {
            'download.default_directory': download_path,
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
        })

        self.driver = webdriver.Chrome(executable_path=executable_path, chrome_options=options)

        if chrome_headless:
            self.driver.command_executor._commands['send_command'] \
                = ('POST', '/session/$sessionId/chromium/send_command')
            self.driver.execute('send_command', {
                'cmd': 'Page.setDownloadBehavior',
                'params': {
                    'behavior': 'allow',
                    'downloadPath': download_path
                }
            })

        self.logger.info('Start crawling, PID {}'.format(os.getpid()))

        self.driver.implicitly_wait(5)

        self.driver.get('https://nid.naver.com/nidlogin.login')

        self.driver.execute_script('document.getElementsByName("id")[0].value="{}"'.format(username))
        time.sleep(1)
        self.driver.execute_script('document.getElementsByName("pw")[0].value="{}"'.format(password))
        time.sleep(1)

        self.driver.find_element_by_xpath('//*[@id="frmNIDLogin"]/fieldset/input').click()
        self.logger.info('로그인 완료')

        try:
            self.driver.find_element_by_class_name('btn_cancel').click()
            self.logger.info('새로운 기기 등록 안함')
            time.sleep(2)
        except NoSuchElementException:
            pass

    def main(self):
        while not self.__stop:
            self.driver.get('https://sell.smartstore.naver.com/#/naverpay/sale/delivery?summaryInfoType=DELIVERY_READY')
            self.logger.info('스마트스토어 이동')
            time.sleep(3)

            self.driver.switch_to.frame(self.driver.find_element_by_id('__naverpay'))
            self.logger.info('iframe 선택')
            time.sleep(3)

            self.driver.find_element_by_xpath('//select[@name="orderStatus"]/option[text()="발송대기"]').click()
            self.logger.info('발송대기 선택')
            time.sleep(3)

            # 엑셀 다운로드
            self.driver.find_element_by_class_name('_excelDownloadBtn').click()
            self.logger.info('엑셀 다운로드')
            time.sleep(5)

            # 엑셀 열기
            files = [f for f in glob.glob(os.path.join('{}/*.xlsx'.format(download_path))) if os.path.isfile(f)]
            self.logger.info(files)

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
                        self.logger.info(payload)
                        requests.post(api_url, data=json.dumps(payload), headers=headers)

                        # 상품주문번호, 배송방법(배송없음), 택배사(공란), 송장번호(공란)
                        ws_batch.write(order_count, 0, row[17].value)
                        ws_batch.write(order_count, 1, '직접전달')
                        ws_batch.write(order_count, 2, '')
                        ws_batch.write(order_count, 3, '')

                if order_count > 0:
                    wb_batch.save(os.path.join(download_path, batch_excel))
                    self.logger.info('일괄발송 엑셀 파일 저장')

                    parent = self.driver.current_window_handle
                    child = None

                    self.driver.find_element_by_xpath('//button[text()="엑셀 일괄발송"]').click()
                    self.logger.info('일괄발송 엑셀 팝업 띄우기')
                    time.sleep(3)

                    for handle in self.driver.window_handles:
                        if handle != parent:
                            child = handle

                    self.driver.switch_to.window(child)
                    time.sleep(3)

                    self.driver.find_element_by_xpath('//input[@name="uploadedFile"]') \
                        .send_keys(os.path.join(download_path, batch_excel))
                    self.logger.info('일괄발송 엑셀 파일 선택')
                    time.sleep(5)

                    self.driver.find_element_by_xpath('//span[text()="일괄 발송처리"]').click()
                    self.logger.info('일괄발송 엑셀 파일 업로드')
                    time.sleep(5)

                    self.driver.find_element_by_xpath('//span[text()="닫기"]').click()
                    self.logger.info('일괄발송 엑셀 팝업 닫기')
                    time.sleep(3)

                    self.driver.switch_to.window(parent)
                    time.sleep(3)

                    os.remove(os.path.join(download_path, batch_excel))
                    self.logger.info('일괄발송 엑셀 로컬 삭제')

                wb_order.close()
                os.remove(files[0])
                self.logger.info('엑셀 로컬 삭제')

            time.sleep(60)
        else:
            self.driver.quit()

    def stop(self, signum, frame):
        # SIGINT, SIGTERM 시그널 수신 종료 핸들러
        self.__stop = True
        self.logger.info('Receive Signal {}'.format(signum))
        self.logger.info('Stop crawling')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--pid', help='pid filename', required=True)
    parser.add_argument('--log', help='log filename', default=None)
    args = parser.parse_args()

    # 첫 번째 fork
    pid = os.fork()
    if pid > 0:
        # 부모 프로세스 그냥 종료
        exit(0)
    else:
        # 부모 환경과 분리
        os.chdir('/')
        os.setsid()
        os.umask(0)

        # 두 번째 fork
        pid = os.fork()
        if pid > 0:
            exit(0)
        else:
            sys.stdout.flush()
            sys.stderr.flush()

            si = open(os.devnull, 'r')
            so = open(os.devnull, 'a+')
            se = open(os.devnull, 'a+')

            os.dup2(si.fileno(), sys.stdin.fileno())
            os.dup2(so.fileno(), sys.stdout.fileno())
            os.dup2(se.fileno(), sys.stderr.fileno())

            with open(args.pid, 'w') as pid_file:
                pid_file.write(str(os.getpid()))

            store = SmartStore(args.log)
            store.main()
