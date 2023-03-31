# 读取excel的每个sheet数据，按照header中的"CVE"进行匹配，将excel A中的"修复链接"数据替换给excel B中的"修复链接"数据
import pandas as pd

# 读取excel A中的数据
df_a = pd.read_excel('A.xlsx', sheet_name='漏洞')
# 读取excel B中的数据
df_b = pd.read_excel('B.xlsx', sheet_name=None)
mp = {}
# 读取excel B中的所有"CVE":"修复链接"数据
for i in df_b:
    sht = df_b[i]
    col = sht.columns.to_list()
    if 'CVE' in col and '修复链接' in col:
        ia = col.index('CVE')
        ib = col.index('修复链接')
        for l in sht.values:
            mp[l[ia]] = l[ib]
# 将mp中的数据替换给excel A中的数据
num_rows = df_a.shape[0]
for i in range(num_rows):
    cve = df_a.iloc[i, 2]
    if cve in mp:
        df_a.iloc[i,8] = mp[cve]
    else:
        print(cve)

# 将excel A的数据重新写入文件
df_a.to_excel('C.xlsx', sheet_name='漏洞')