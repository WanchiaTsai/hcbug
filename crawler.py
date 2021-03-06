#!/usr/bin/python
#coding:utf-8
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import re
import pandas
import csv
import os
import sys
from datetime import datetime


def get_data_frame(file_loc):
    df = pandas.DataFrame()
    data = pandas.read_excel(file_loc)
    df = df.append(data)
    return df


def get_file_header(df):
    headers = []
    for lists in df.items():
        list_name = lists[0]
        if(isinstance(list_name, str) and list_name.find('Unnamed')==-1):
            headers.append(list_name)
    return headers


def get_company_list(df, headers):

    lists = df[headers[0]].values
    return lists


def get_company_key_word(company_name):
    keyword = company_name

    index = company_name.find("公司")
    if index != -1:
        keyword = company_name[:index+2]
    index = company_name.find("(股)")
    if index != -1:
        keyword = company_name[:index + 3]
        keyword = str(keyword).replace('(股)', "股份有限")
        keyword = str(keyword).replace('（股）', "股份有限")
    index = company_name.find("（股）")
    if index != -1:
        keyword = company_name[:index + 3]
        keyword = str(keyword).replace('（股）', "股份有限")
    return keyword


def do_search(keyword):
    driver = webdriver.Chrome()
    driver.get("http://findbiz.nat.gov.tw/fts/query/QueryList/queryList.do")

    elem = driver.find_element_by_id("qryCond")
    elem.clear()

    # elem.send_keys("65282782")
    elem.send_keys(keyword)
    # elem.send_keys("台灣端板鋼鐵企業股份有限公司高雄廠")
    # elem.send_keys("nonono")

    # elem.send_keys(Keys.RETURN)
    elem.send_keys(Keys.ENTER)
    return driver


def get_search_results(driver):
    results = driver.find_elements_by_class_name("companyName")
    return results


def check_search_results(search_results):
    if len(search_results) == 0:
        return False
        # exit(print('無結果：' + str(noResult), '大於一筆：' + str(moreThanOne)))


def get_company_page(search_results, driver):
    # first_result_name = search_results[0].text
    first_result_name=driver.find_element_by_css_selector('#vParagraph > div > div.panel-heading.companyName > a:nth-child(1)').text
    # print(first_result_name)
    link = driver.find_element_by_link_text(first_result_name)
    link.click()
    return driver


def analyze_company_table(driver):
    columns = []
    fields = []
    column = ''
    typicalColumns = ['統一編號', '公司狀況', '公司名稱', '資本總額(元)', '代表人姓名', '公司所在地', '登記機關', '核准設立日期', '最後核准變更日期', '所營事業資料']
    info = {}
    for row in driver.find_elements_by_css_selector('#tabCmpyContent > div > table > tbody > tr'):
        cells = row.find_elements_by_tag_name("td")
        for idx, cell in enumerate(cells):
            if idx == 0 and cell.get_attribute("class") != 'txt_td':
                break
            rawText = re.sub(' +', ' ', cell.text.strip())
            if len(rawText) == 0:
                break
            if idx == 0:
                column = rawText.replace(" ", "")
                columns.append(column)
            if idx == 1:
                if column != '所營事業資料':
                    field = rawText.split(' ')[0]
                    fields.append(field)
                else:
                    field = rawText
                    fields.append(field)
                info[column]=field
    return info


def write_data(file_name, data):
    f = open(file_name, "a+")
    w = csv.writer(f)
    w.writerows(data)
    f.close()


def modified_company_name(company_name):
    company_name = str(company_name)
    company_name = company_name.replace('\r', "")
    company_name = company_name.replace('\n', "")
    company_name = company_name.replace('"', "")
    if company_name == "nan":
        return False
    if len(company_name) < 3:
        return False
    return company_name


def open_new_csv(file_name):
    # http: // dewerzht - blog.logdown.com / posts / 701780 - python - log -with-csv
    f = open(file_name, "w+")
    f.close()
    return f


def row_count(file_name):
    # print(file_name)
    # exit(0)
    input_file = open(file_name, "r+")
    reader_file = csv.reader(input_file)
    value = len(list(reader_file))
    input_file.close()
    return value


def get_csv_file_name(file_loc):

    index=file_loc.rfind("/")
    file_dir = file_loc[:index]
    org_file_name_ext = file_loc[index+1:]

    index=org_file_name_ext.rfind(".")
    org_file_name = org_file_name_ext[:index]

    # exit()

    # 指定路徑
    pathProg = '/home/wanchia/PycharmProjects/hcbug'+'/'+file_dir
    os.chdir(pathProg)

    if os.getcwd() != pathProg:
        print("EEROR: the file path incorrect.")
        sys.exit()

    fileDT = datetime.now().strftime('%Y%m%d_%H%M%S')
    file_name = pathProg+'/'+org_file_name+".csv"
    return file_name


def main():
    # print(sys.version)  # parentheses necessary in python 3.
    # print(sys.version_info)
    # print(sys.version)
    # print(sys.path)
    # exit(0)

    print('Number of arguments:', len(sys.argv), 'arguments.')
    print('Argument List:', str(sys.argv))
    # my code here
    file_loc = sys.argv[1]
    # file_loc='0-工業區名單/0-南投/10.南崗工業區廠商名錄.xlsx'
    # file_loc = '0-工業區名單/0-台南/08安平工業區廠商名錄.xlsx'
    print(file_loc)

    df = get_data_frame(file_loc)

    headers = get_file_header(df)
    company_list = get_company_list(df, headers)
    # company_list =['慶宏技術開發有限公司']

    csv_file_name = get_csv_file_name(file_loc)
    finish_count=0
    if os.path.isfile(csv_file_name):
        csv_row = row_count(csv_file_name)
        finish_count=csv_row-1
    if finish_count == 0:
        open_new_csv(csv_file_name)
        csv_headers = [[
            '廠商名稱',
            '搜尋關鍵字',
            '結果數量',
            '公司狀況',
            '公司名稱',
            '統一編號',
            '代表人姓名',
            '公司所在地']]
        write_data(csv_file_name, iter(csv_headers))
    # else:
    #     exit(0)
    logs = []
    count = 0
    no_result_combo=0
    for idx, company_name in enumerate(iter(company_list)):
        # if(idx >= 3):
        #     break
        if no_result_combo >= 5:
            exit("Combo 5 of no result")
        if(len(logs) >= 10):
            write_data(csv_file_name, logs)
            print("WRITE~")
            logs =[]
        log = []
        modified_name = modified_company_name(company_name)
        if modified_name == False:
            continue

        count += 1
        print(file_loc,":",company_name)
        print("count:~", finish_count, "...", count, "/", len(company_list))

        if(count <= finish_count):
            continue
        log.append(modified_name)
        # print(modified_name)
        keyword = get_company_key_word(modified_name)
        log.append(str(keyword))
        if keyword == False:
            logs.append(log)
            continue

        driver = do_search(keyword)
        search_results = get_search_results(driver)
        count_result=len(search_results)
        log.append(count_result)
        if len(search_results) == 0:
            print("NO RESULT")
            logs.append(log)
            no_result_combo += 1
            driver.close()
            continue
        else:
            no_result_combo = 0
        driver = get_company_page(search_results, driver)

        info = analyze_company_table(driver)
        critical_columns = ['公司狀況', '公司名稱', '統一編號', '代表人姓名', '公司所在地']

        # print(info)
        driver.close()
        for column in critical_columns:
            log.append(info.get(column, ''))
        log.append(info)
        logs.append(log)
        # print(log)
    write_data(csv_file_name, logs)


if __name__ == "__main__":
    main()
