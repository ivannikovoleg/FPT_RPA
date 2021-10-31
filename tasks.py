import time
import os
import xlsxwriter
import fitz
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait


def wait_for_downloads(path):
    print('Waiting for downloads', end='')
    while any([filename.endswith('.crdownload') for filename in
               os.listdir(path)]):
        time.sleep(2)
        print(".", end="")
    print('\nAll files downloaded!')


def parse_pdf_data(filename):
    doc = fitz.open(filename)
    investment_name = '1. Name of this Investment:'
    uii = '2. Unique Investment Identifier (UII):'
    page = doc.load_page(0)
    page_text = page.get_text("text")
    values = []
    for line in page_text.split('\n'):
        if uii in line:
            values.append(line.split(':')[1].strip())
        elif investment_name in line:
            values.append(line.split(':')[1].strip())
    return values


def main():
    print('Initialize.')

    with open('config.txt') as file:
        dep_to_scrap = file.readline().strip()
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': f'{os.getcwd()}/output'}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=chrome_options)

    print(f'Start!\nDepartment to search:{dep_to_scrap}')
    driver.get('https://itdashboard.gov/')
    dive_it = driver.find_element(By.XPATH, "//*[@id='node-23']/div/div/div/div/div/div/div/a")
    dive_it.click()
    print('Click on "Dive in"!')
    deps = WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located(
        (By.XPATH, '//div[@id="agency-tiles-widget"]//span[@class="h4 w200"]')))
    budgets = driver.find_elements(By.XPATH, '//div[@id="agency-tiles-widget"]//span[@class=" h1 w900"]')

    workbook = xlsxwriter.Workbook('output/write_data.xlsx')
    worksheet = workbook.add_worksheet(name="Departments")
    worksheet.set_column(0, 0, 45)
    for row_num, data in enumerate(deps):
        worksheet.write(row_num, 0, data.text)
    for row_num, data in enumerate(budgets):
        worksheet.write(row_num, 1, data.text)
    print('Departments budgets: Done!')

    driver.find_element(By.XPATH, f"//span[contains(text(), '{dep_to_scrap}')]").click()
    select = Select(WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, '//*[@id="investments-table-object_length"]/label/select'))))
    select.select_by_value('-1')
    print('Show all entries.')
    time.sleep(15)

    tmp_table = []
    for i in driver.find_elements(By.XPATH, '//*[@id="investments-table-object"]//tr'):
        if i.text != '':
            tmp_table.append(i.text)

    table = driver.find_elements(By.XPATH, '//*[@id="investments-table-object"]//tr//td')
    agencies = workbook.add_worksheet("Agencies")
    agencies.set_column(0, 0, 15)
    agencies.set_column(1, 1, 45)
    agencies.set_column(2, 2, 40)
    agencies.set_column(4, 4, 30)
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
    print('Excel file is done!')

    urls = driver.find_elements(By.XPATH, '//*[@id="investments-table-object"]//tr//a')
    links = []
    for url in urls:
        links.append(url.get_attribute("href"))
    for link in links:
        driver.get(link)
        button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located(
            (By.XPATH, '//*[@id="business-case-pdf"]/a')))
        driver.execute_script("arguments[0].click();", button)
        print('Downloading file.')
        time.sleep(10)

    wait_for_downloads('output')

    file_list = []
    for file in os.listdir('output'):
        if file.endswith('.pdf'):
            file_list.append(os.path.join('output', file))
    for file in file_list:
        v = parse_pdf_data(file)
        print(f'Checking entries {v[1]}: {v[0]}')
        for element in tmp_table:
            if v[0] and v[1] in element:
                print('UII and Name of Investment are same!')

    driver.close()
    print('Browser closed!')


if __name__ == '__main__':
    main()
