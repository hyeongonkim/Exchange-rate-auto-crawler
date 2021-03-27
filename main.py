import pandas as pd
import time
from selenium import webdriver
from datetime import date
from os.path import expanduser
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def get_path(filename):
    home = expanduser('~')
    return '{}/Desktop/{}.xlsx'.format(home, filename)


def parse_rate(specific_date, code, amount):
    home = expanduser('~')
    driver_path = '{}/Desktop/chromedriver'.format(home)
    options = webdriver.ChromeOptions()
    # options.add_argument('headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument("disable-gpu")
    options.add_argument('no-sandbox')
    options.add_argument(
        "user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36")
    driver = webdriver.Chrome(driver_path, options=options)
    formatted_date = date.fromisoformat(specific_date)
    print('크롤링 시도')
    driver.get('https://www.exchange-rates.org/Rate/{}/USD/{}-{}-{}'.format(code, formatted_date.month, formatted_date.day, formatted_date.year))
    wait = WebDriverWait(driver, 600)
    element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ratesTable"]/tbody/tr[1]/td[3]')))
    rate = float(element.text[:-4])
    print('크롤링 성공 {}'.format(rate))
    driver.quit()
    return rate, round(float(amount) * rate, 2)


def write_data_at_sheet():
    exist_df = open_file()
    exist_df['환율'] = 0.0
    exist_df['환전금액'] = 0.0

    for row in range(exist_df.shape[0] - 1):
        code = exist_df.loc[row]['결제국가\n화폐\nBuyer Currency']
        if code == 'USD' or code == 'KRW':
            continue

        print(exist_df.loc[row])

        exist_df.loc[row, '환율'], exist_df.loc[row, '환전금액'] = parse_rate(
            str(exist_df.loc[row]['결제날짜\nTransaction Date'])[:10],
            exist_df.loc[row]['결제국가\n화폐\nBuyer Currency'],
            exist_df.loc[row]['결제금액\nAmount (Buyer Currency)']
        )

        print(exist_df.loc[row])
        time.sleep(0.5)

    writer = pd.ExcelWriter(get_path('to'), engine='xlsxwriter')
    exist_df.to_excel(writer, sheet_name='변환값')
    writer.close()


def open_file():
    return pd.read_excel(
        get_path('target'),
        sheet_name='결제매출내역_2019.03월-2020.12월',
        skiprows=1,
        usecols='B, P, Q',
        parse_dates=['결제날짜\nTransaction Date']
    )


if __name__ == '__main__':
    write_data_at_sheet()
