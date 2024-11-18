import requests
from bs4 import BeautifulSoup
import pandas as pd
requests.packages.urllib3.disable_warnings()
import time
from retrying import retry
import random

#爬取各学校子页面参数
url = r"https://xkkm.sdzk.cn/web/xx.html"
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36'}

response = requests.get(url=url, headers=headers)
response.encoding = "utf-8"
soup = BeautifulSoup(response.text,"lxml")


tem = []
for i in soup.find_all("tr"):
    for x in i.find_all("td"):
        tem.append((x.text).strip())
tem = pd.DataFrame(tem)
# 此数据导出到excel只有一列，需要使用公式额外处理（excel公式：INDEX($A:$A,5*ROW(A1)-5+COLUMN(A1))）
tem.to_excel("山东省教育招生考试院学校参数爬取.xlsx", index = False)

#爬取各学校页面详细数据

#设置代理
proxy = '127.0.0.1:4780'
proxies = {
    'http': 'http://' + proxy,
    'https': 'http://' + proxy}

@retry(stop_max_attempt_number=3)
def rty_post(url,headers,payload):
    response = requests.post(url=url, data=payload, headers=headers, timeout=10, proxies=proxies, verify=False)
    return response


univ = []
num = []
lev = []
sub = []
rang = []
sub_re = []

url = "https://xkkm.sdzk.cn/xkkm/queryXxInfor"
headers = {'DNT': '1',
           'Sec-Fetch-User': '?1',
           'Upgrade-Insecure-Requests': '1',
           'Cookie': 'Secrue; PTGK-PT=36703209; Secrue',
           'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
           'Content-Type': 'application/x-www-form-urlencoded',
           'Accept': '*/*',
           'Host': 'xkkm.sdzk.cn',
           'Connection': 'keep-alive'}

for i in range(0,url_source.shape[0]):
    payload=f'dm={url_source.iloc[i,4]}&mc={url_source.iloc[i,8]}&yzm=ok&nf=2024'
    response = rty_post(url=url, headers=headers, payload=payload)
    time.sleep(random.randint(1,3))
    soup = BeautifulSoup(response.text,"lxml")
    for z in soup.find_all("tr"):
        for x in z.find_all("td"):
            if x["width"] == "5%":
                num.append(x.text.strip())
                univ.append(soup.find("title").text.replace("选考科目范围-",""))
            elif x["width"] == "10%":
                lev.append(x.text.strip())
            elif x["width"] == "25%":
                sub.append(x.text.strip())
            elif x["width"] == "30%"and x["style"] == "display:table-cell; vertical-align:middle;":
                rang.append(x.text.strip())
            elif x["style"] == "white-space: nowrap;":
                sub_re.append((str(x).replace("<br/>",",").replace(" ","").replace('\r','').replace('\n','').replace('</td>','').replace('<tdalign="left"style="white-space:nowrap;"width="30%"><!--用jstl的fn标签库对传过来的专业中的\'、\'进行替换成,-->','')))
    print(f"{url_source.iloc[i,2]}爬取完成")
    

tem = list(zip(univ,num,lev,sub,rang,sub_re))
df1 = pd.DataFrame(tem)
col = ["学校名称","序号","层次","专业类名称","选考科目范围","类中所含专业"]
df1.set_axis(col, axis=1,inplace=True)

df1.to_excel("2024年山东省普通高校招生专业（专业类）选考科目要求.xlsx", index=False)
