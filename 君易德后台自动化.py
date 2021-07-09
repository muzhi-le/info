import requests,time,re,openpyxl,pinyin
from selenium import webdriver


def getweb(name, url, abb, s):
    # print(name, url, abb)
    # driver.find_element_by_xpath('//*[@id="main"]/section/aside/ul/div[5]/li/div').click()
    # time.sleep(1)
    # driver.find_element_by_xpath('//*[@id="main"]/section/aside/ul/div[5]/li/ul/li[2]/ul/li').click()
    # time.sleep(1)
    driver.get('http://admin.caishuiyide.com/admin/article/articlesource/add/')
    driver.find_element_by_xpath('//*[@id="id_name"]').send_keys(name)
    time.sleep(0.5)
    driver.find_element_by_xpath('//*[@id="id_display_name"]').send_keys(name)
    time.sleep(0.5)
    driver.find_element_by_xpath('//*[@id="id_domain"]').send_keys(url)
    time.sleep(0.5)
    driver.find_element_by_xpath('//*[@id="id_code"]').send_keys(abb)
    time.sleep(0.5)
    driver.find_element_by_xpath('//*[@id="articlesource_form"]/div/fieldset/div[7]/div/div/span/span[1]/span').click()
    time.sleep(1)
    driver.find_element_by_xpath('/html/body/span/span/span[1]/input').send_keys(s)
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="select2-id_region-results"]/li[1]').click()
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="articlesource_form"]/div/div/input[3]').click()
    time.sleep(1)
    driver.get('http://admin.caishuiyide.com/admin/article/articlesourcedomain/add/')
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="articlesourcedomain_form"]/div/fieldset/div[1]/div/div/span/span[1]/span').click()
    time.sleep(2)
    driver.find_element_by_xpath('/html/body/span/span/span[1]/input').send_keys(name)
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="select2-id_source-results"]/li').click()
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="id_domain"]').send_keys(url)
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="articlesourcedomain_form"]/div/div/input[3]').click()
    time.sleep(1)

if __name__ == '__main__':
    admin_url = 'http://admin.caishuiyide.com/admin/#/admin/article/articlesource/'
    driver = webdriver.Chrome()
    # driver.maximize_window()
    driver.get(admin_url)
    driver.find_element_by_xpath('//*[@id="login-form"]/div[1]/div/input').send_keys('lile')
    driver.find_element_by_xpath('//*[@id="login-form"]/div[2]/div/input').send_keys('lile123456')
    driver.find_element_by_xpath('//*[@id="login-form"]/div[3]/button').click()
    i = int(input('请输入索引起始位置：'))
    j = int(input('请输入索引结束位置：'))
    s = input('请输入城市：')
    work = openpyxl.load_workbook('C:/Users/muzhi/Desktop/二期采集目录(已自动还原).xlsx')
    rk = work.get_sheet_by_name("目录")
    web_name = rk["D"][i:j]
    data_name = []
    abbreviation = []
    names = []
    for x in range(len(web_name)):
        x_name = web_name[x].value
        # names.append(x_name)
        xl_name = '市_'.join(x_name.split('市'))
        data_name.append(x_name)
        # print(x_name)
        p = []
        for pin in xl_name:
            py = pinyin.get(pin)[0]
            p.append(py)
        py_first = 'sxs_' + ''.join(p)  # 'nxhzzzq_' +
        abbreviation.append(py_first)
    web_url = rk["E"][i:j]
    data_url = []
    for y in range(len(web_url)):
        y_url = web_url[y].value.split('//')[-1].split('/')[0]
        data_url.append(y_url)

    for name, url, abb in zip(data_name, data_url, abbreviation):
        print(name, url, abb)
        getweb(name, url, abb, s)

    driver.close()
