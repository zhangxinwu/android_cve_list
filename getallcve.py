from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd

# 设置Edge浏览器驱动程序
edge_options = webdriver.EdgeOptions()
edge_options.use_chromium = True
edge_options.add_argument('--proxy-server=socks5://127.0.0.1:7890')
browser = webdriver.Edge(executable_path='msedgedriver.exe', options=edge_options)

all_table = {}
urls = []
# 访问页面
for year in range(2019, 2023):
    for month in range(1, 13):
        if year == 2019 and month < 2:
            continue
        if year == 2023 and month > 3:
            continue
        u = 'https://source.android.com/docs/security/bulletin/{}-{}-01?hl=zh-cn'.format(year, str(month).zfill(2))
        print(u)
        urls.append(u)
        browser.get(u)
        body = browser.find_element(By.CSS_SELECTOR, '#gc-wrapper > main > devsite-content > article > div.devsite-article-body.clearfix')
        els = body.find_elements(By.XPATH, '*')
        elsi = 0
        a = browser.find_elements(By.XPATH, '//div[@class="devsite-table-wrapper"]//table')
        print(len(a))
        lasth3 = 0
        llasth3 = 0
        for b in a:
            while elsi < len(els):
                el = els[elsi]
                if el.tag_name == 'h3' or el.tag_name == 'h2':
                    lasth3 = elsi
                elsi += 1
                if el.tag_name == 'div' and el.get_attribute('class') == 'devsite-table-wrapper' and b == el.find_element(By.TAG_NAME, 'table'):
                    break
            if llasth3 == lasth3:
                continue
            llasth3 = lasth3
            table_name = els[lasth3].text.strip().lower()
            if table_name not in all_table:
                all_table[table_name] = []
            tr = b.find_element(By.TAG_NAME, 'tbody').find_elements(By.TAG_NAME, 'tr')
            name = []
            data = []
            for i in tr:
                # table header
                for t in i.find_elements(By.TAG_NAME, 'th'):
                    name.append(t.text.replace('\n', ' '))

                # table body
                tb = []
                mp = {}
                es = i.find_elements(By.TAG_NAME, 'td')
                for i in range(len(es)):
                    t = es[i]
                    if len(name) == 0:
                        break
                    mp[name[i]] = t.text.replace('\n', ' ')
                    for ue in t.find_elements(By.TAG_NAME, 'a'):
                        url = ue.get_attribute('href')
                        if not url.endswith('#asterisk'):
                            if '修复链接' in mp:
                                mp['修复链接'] = mp['修复链接'] + '\n' + url
                            else:
                                mp['修复链接'] = url
                    # if len(t.find_elements(By.TAG_NAME, 'a')) > 0:
                    #     url = t.find_element(By.TAG_NAME, 'a').get_attribute('href')
                    #     if not url.endswith('#asterisk') or True:
                    #         mp['修复连接'] = url
                if len(mp) > 0:
                    mp['公告日期'] = '{}-{}'.format(year, str(month).zfill(2))
                    data.append(mp)
            all_table[table_name].extend(data)

        # 每次都保存一次
        outTables = {}
        writer = pd.ExcelWriter("output.xlsx", engine='xlsxwriter')
        for k in all_table:
            pd.DataFrame(all_table[k]).to_excel(writer, sheet_name=k, index=False)
        writer.save()  # 保存数据

# 关闭浏览器
browser.quit()