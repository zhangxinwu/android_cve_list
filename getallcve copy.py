import threading
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import queue
import concurrent.futures

urls = []

# 访问页面
def target_function(u, year, month, q, qb):
    # 设置Edge浏览器驱动程序
    edge_options = webdriver.EdgeOptions()
    edge_options.add_argument('--headless')
    edge_options.add_argument('--disable-gpu')
    # edge_options.add_argument('--window-size=80,80')
    edge_options.add_argument('--incognito')
    edge_options.use_chromium = False
    edge_options.add_argument('--proxy-server=http://127.0.0.1:7890')
    browser = webdriver.Edge(executable_path='msedgedriver.exe', options=edge_options)
    print(u)
    all_table = {}
    browser.get(u)
    body = browser.find_element(By.CSS_SELECTOR, '#gc-wrapper > main > devsite-content > article > div.devsite-article-body.clearfix')
    els = body.find_elements(By.XPATH, '*')
    elsi = 0
    a = browser.find_elements(By.XPATH, '//div[@class="devsite-table-wrapper"]//table')
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
            if len(es) > 0:
                if len(es) < len(name):
                    for ii in range(len(name)-len(es)):
                        if len(data) > 0:
                            mp[name[ii]] = data[-1][name[ii]]
                        else:
                            mp[name[ii]] = ""
                for j in range(len(es)):
                    ii = j + len(name)-len(es)
                        
                    t = es[ii]
                    if len(name) == 0:
                        break
                    if ii < 2 and t.text.strip() == "" and len(data) > 0:
                        mp[name[ii]] = data[-1][name[ii]]
                        if '修复链接' in data[-1]:
                            mp['修复链接'] = data[-1]['修复链接']
                    else:
                        mp[name[ii]] = t.text.replace('\n', ' ')
                    for ue in t.find_elements(By.TAG_NAME, 'a'):
                        url = ue.get_attribute('href')
                        if not url.endswith('#asterisk'):
                            if '修复链接' in mp:
                                mp['修复链接'] = mp['修复链接'] + '\n' + url
                            else:
                                mp['修复链接'] = url
                if len(mp) > 0:
                    mp['公告日期'] = '{}-{}'.format(year, str(month).zfill(2))
                    data.append(mp)
        all_table[table_name].extend(data)
    q.put(all_table)
    browser.quit()

# 创建线程池
max_thread_num = 15
pool = concurrent.futures.ThreadPoolExecutor(max_workers=max_thread_num)
result_queue = queue.Queue()
qb = queue.Queue()

# 设置Edge浏览器驱动程序
edge_options = webdriver.EdgeOptions()
# edge_options.add_argument('--headless')
# edge_options.add_argument('--disable-gpu')
# edge_options.add_argument('--window-size=80,80')
# edge_options.add_argument('--incognito')
# edge_options.use_chromium = False
edge_options.add_argument('--proxy-server=http://127.0.0.1:7890')
def run(q):
    q.put(webdriver.Edge(executable_path='msedgedriver.exe', options=edge_options))
# for _ in range(max_thread_num):
#     pool.submit(run, qb)

for year in range(2019, 2024):
    for month in range(1, 13):
        if year == 2019 and month < 2:
            continue
        if year == 2023 and month > 2:
            continue
        if not (year == 2020 and month == 9):
            continue
        u = 'https://source.android.com/docs/security/bulletin/{}-{}-01?hl=zh-cn'.format(year, str(month).zfill(2))
        urls.append(u)
        pool.submit(target_function, u, year, month, result_queue, qb)

pool.shutdown(wait=True)

all_tables = {}
while not result_queue.empty():
    r = result_queue.get()
    for k in r:
        if k not in all_tables:
            all_tables[k] = []
        all_tables[k].extend(r[k])

        

# 保存数据
outTables = {}
writer = pd.ExcelWriter("output_multi.xlsx", engine='openpyxl')
for k in all_tables:
    pd.DataFrame(all_tables[k]).to_excel(writer, sheet_name=k, index=False)
writer.save()  # 保存数据
