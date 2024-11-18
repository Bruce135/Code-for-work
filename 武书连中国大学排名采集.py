#!/usr/bin/env python
# coding: utf-8

# In[85]:


import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
import os
import glob
import json
requests.packages.urllib3.disable_warnings()


# In[110]:


url = r"https://mallapi.wurank.net/RankApi/SearchApi/GetMingBanDuliUniDataPageList/mingbanduliunidata"


# In[95]:


def rty_post(url,headers,body):
    response = requests.post(url=url, data=json.dumps(body), headers=headers, timeout=20, verify=False)
    return response


# In[111]:


headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
           'Accept':'application/json, text/javascript, */*; q=0.01',
           'Accept-Language':'zh-CN,zh;q=0.9,en;q=0.8,en-US;q=0.7,zh-MY;q=0.6,en-MY;q=0.5',
           'Connection': 'keep-alive',
           'Origin':'https://www.wurank.net',
           'Content-Type':'application/json'}

body = { "PageIndex": 1, 
        "PageSize": 400,
        "filter": "intYear=2024", 
        "sort": "intVictorOrder=0" }


# In[112]:


response = rty_post(url=url, headers=headers, body=body)


# In[107]:


Name = []
Rank = []
Province_name = []
province_order = []
Schtype = []
Classorder = []
Totalscore = []
Schstyle = []

for i in response.json()["data"]:
    Name.append(i['schchnname'])
    Rank.append(i['victororder'])
    Province_name.append(i['provincename'])
    province_order.append(i['provinceorder'])
    Schtype.append(i['schtype'])
    Classorder.append(i['classorder'])
    Totalscore.append(i['totalscore'])


# In[108]:


tem = list(zip(Name,Rank,Province_name,province_order,Schtype,Classorder,Totalscore))
df1 = pd.DataFrame(tem)
col = ["机构名称","总排名","省份名称","省份排名","学校类型","类型排名","总分"]
df1.set_axis(col, axis=1,inplace=True)


# In[109]:


with pd.ExcelWriter(r'C:\Users\shrk-3121\Desktop\武书连2024中国独立学院排名.xlsx',engine='xlsxwriter',engine_kwargs={'options':{'strings_to_urls': False}}) as writer:
        df1.to_excel(writer, index=False)


# In[ ]:




