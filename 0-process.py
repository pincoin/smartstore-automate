import argparse
import logging
import os
import signal
import time


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

    def main(self):
        i = 0
        self.logger.info('Start Singing, PID {}'.format(os.getpid()))
        while not self.__stop:
            self.logger.info(i)
            i += 1
            time.sleep(1)

    def stop(self, signum, frame):
        # SIGINT, SIGTERM 시그널 수신 종료 핸들러
        self.__stop = True
        self.logger.info('Receive Signal {}'.format(signum))
        self.logger.info('Stop Singing')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--log', help='log filename', default=None)
    args = parser.parse_args()

    store = SmartStore(args.log)
    store.main()
