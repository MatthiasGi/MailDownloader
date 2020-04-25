from config import Config
from mail import Mail
import time

import os

def main():
    config = Config()
    mail = Mail(config)

    while True:
        try:
            mail.check()
            time.sleep(10 * 60) # 10 minutes
        except KeyboardInterrupt:
            break

if __name__ == '__main__':
    main()
