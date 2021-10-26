import time
import os
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait


def wait_for_element(browser, delay, xpath):
    loaded = WebDriverWait(browser, delay).until(EC.visibility_of_element_located((By.XPATH, xpath)))
    if loaded:
        print('Loaded!')
        return True
    else:
        return False


def main():
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': f'{os.getcwd()}\\output'}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(executable_path='chromedriver', options=chrome_options)
    dep_to_scrap = "'National Science Foundation'"
    workbook = xlsxwriter.Workbook('output\\write_data.xlsx')

    driver.get('https://itdashboard.gov/')

    dive_it = driver.find_element(By.XPATH, "//*[@id='node-23']/div/div/div/div/div/div/div/a")
    dive_it.click()
    print('Click!')

    if wait_for_element(driver, 10, '//*[@id="agency-tiles-container"]'):
        deps = driver.find_elements(By.XPATH, '//div[@id="agency-tiles-widget"]//span[@class="h4 w200"]')
        budgets = driver.find_elements(By.XPATH, '//div[@id="agency-tiles-widget"]//span[@class=" h1 w900"]')
        worksheet = workbook.add_worksheet(name="Departments")
        for row_num, data in enumerate(deps):
            worksheet.write(row_num, 0, data.text)
        for row_num, data in enumerate(budgets):
            worksheet.write(row_num, 1, data.text)

        driver.find_element(By.XPATH, f"//span[contains(text(), {dep_to_scrap})]").click()
        print('Move on')
    else:
        print('Error')

    if wait_for_element(driver, 10, '//*[@id="investments-table-object_length"]/label/select'):

        select = Select(driver.find_element(By.XPATH, '//*[@id="investments-table-object_length"]/label/select'))
        select.select_by_value('-1')

        time.sleep(10)

        table = driver.find_elements(By.XPATH, '//*[@id="investments-table-object"]//tr//td')
        agencies = workbook.add_worksheet("Agencies")
        row_num = 0
        col_num = 0
        for data in table:
            if data.text != '':
                if col_num == 7:
                    row_num += 1
                    col_num = 0
                agencies.write(row_num, col_num, data.text)
                col_num += 1
        workbook.close()

        urls = driver.find_elements(By.XPATH, '//*[@id="investments-table-object"]//tr//a')
        links = []
        for url in urls:
            links.append(url.get_attribute("href"))
        print(links)
        for link in links:
            driver.get(link)
            if wait_for_element(driver, 10, '//*[@id="business-case-pdf"]/a'):
                button = driver.find_element(By.XPATH, '//*[@id="business-case-pdf"]/a')
                driver.execute_script("arguments[0].click();", button)
                time.sleep(10)
    else:
        print('Error!')

    time.sleep(20)
    driver.close()


if __name__ == '__main__':
    main()
