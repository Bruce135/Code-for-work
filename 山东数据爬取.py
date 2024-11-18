#!/usr/bin/env python
# coding: utf-8

# In[476]:


import requests
from bs4 import BeautifulSoup
import pandas as pd
requests.packages.urllib3.disable_warnings()
import time
from retrying import retry
import random


# In[467]:


url = r"https://xkkm.sdzk.cn/web/xx.html"


# In[109]:


headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36'}


# In[110]:


response = requests.get(url=url, headers=headers)


# In[111]:


response.encoding = "utf-8"


# In[112]:


soup = BeautifulSoup(response.text,"lxml")


# In[132]:


tem = []
for i in soup.find_all("tbody"):
    for x in i.find_all("input"):
        tem.append(x["value"])


# In[29]:


tem = []
for i in soup.find_all("tr"):
    for x in i.find_all("td"):
        tem.append((x.text).strip())


# In[134]:


tem = pd.DataFrame(tem)


# In[135]:


tem.to_excel("山东省教育招生考试院数据爬取更新.xlsx", index = False)


# In[31]:


tem1 = pd.DataFrame(tem)


# In[33]:


tem1.to_excel("山东省教育招生考试院数据爬取.xlsx", index = False)


# In[452]:


url_source = pd.read_excel(r"C:\Users\shrk-3121\SynologyDrive\Work\Docs\临时任务\2024年7月\山东省教育招生考试院数据爬取.xlsx")


# In[462]:


url_source.iloc[0,4]


# In[470]:


url_source.iloc[0,2]


# In[446]:


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
payload=f'dm={}&mc={}&yzm=ok&nf=2024'


# In[447]:


response = requests.post(url=url, headers=headers, data=payload)


# In[448]:


soup = BeautifulSoup(response.text,"lxml")


# In[449]:


univ = []
num = []
lev = []
sub = []
rang = []
sub_re = []
for i in soup.find_all("tr"):
    for x in i.find_all("td"):
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


# In[450]:


tem = list(zip(univ,num,lev,sub,rang,sub_re))


# In[451]:


tem


# In[443]:


df1 = pd.DataFrame(tem)
col = ["col1","col2","col3","col4","col5","col6"]
df1.set_axis(col, axis=1,inplace=True)


# In[445]:


df1.to_excel("北京大学测试.xlsx", index=False)


# In[488]:


# 完整代码（链接文件已经导入（url_source））

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


# In[497]:


@retry(stop_max_attempt_number=3)
def rty_post(url,headers,payload):
    response = requests.post(url=url, data=payload, headers=headers, timeout=10, proxies=proxies, verify=False)
    return response

def params(dm,mc):
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
    payload=f'dm={dm}&mc={mc}&yzm=ok&nf=2024'

    return url,headers,payload

def fetch_data(soup,name):
    for z in soup.find_all("tr"):
        for x in z.find_all("td"):
            if x["width"] == "5%":
                num.append(x.text.strip())
                univ.append(name)
            elif x["width"] == "10%":
                lev.append(x.text.strip())
            elif x["width"] == "25%":
                sub.append(x.text.strip())
            elif x["width"] == "30%"and x["style"] == "display:table-cell; vertical-align:middle;":
                rang.append(x.text.strip())
            elif x["style"] == "white-space: nowrap;":
                sub_re.append((str(x).replace("<br/>",",").replace(" ","").replace('\r','').replace('\n','').replace('</td>','').replace('<tdalign="left"style="white-space:nowrap;"width="30%"><!--用jstl的fn标签库对传过来的专业中的\'、\'进行替换成,-->','')))
    print(f"{name}爬取完成")

def main():
    for i in range(0,url_source.shape[0]):
        dm,mc = url_source.iloc[i,4], url_source.iloc[i,8]
        params(dm,mc)
        response = rty_post(url=url, headers=headers, payload=payload)
        time.sleep(random.randint(1,3))
        soup = BeautifulSoup(response.text,"lxml")
        name = soup.find("title").text.replace("选考科目范围-","")
        fetch_data(soup,name)
        
if __name__ == '__main__':
    univ = []
    num = []
    lev = []
    sub = []
    rang = []
    sub_re = []
    main()


# In[495]:


univ


# In[ ]:




