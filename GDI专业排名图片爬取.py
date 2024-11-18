#!/usr/bin/env python
# coding: utf-8

# In[4]:


import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
import glob
import json
requests.packages.urllib3.disable_warnings()


# In[180]:


# 下载所有专业类ID
url_page = f"https://huacheng.gz-cmc.com/json/topic/2024/06/14/42889fe292b04d0caee191fc36f8eb8e/42889fe292b04d0caee191fc36f8eb8e.topicjson"
headers = {"User-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"}
response = requests.get(url=url_page,headers = headers)
response_json = response.json()
Name = []
url_id = []

for i in response_json["channels"]:
    Name.append(i["name"])
    url_id.append(i["id"])
    
tem = list(zip(Name,url_id))
df_id = pd.DataFrame(tem)
col = ["专业类名称","id"]
df_id.set_axis(col, axis=1,inplace=True)
df_id.to_excel("GDI查询ID.xlsx",index = False)


# In[ ]:


# 读取上一步怕爬好的专业类ID文件
df_id = pd.read_excel(r"C:\Users\shrk-3121\GDI查询ID.xlsx")


# In[182]:


# 根据爬取的专业类ID组合为新的URL，爬取专业类下包含各专业的json文件
for i in range(0,df_id.shape[0]):
    try:
        url_page = f"https://huacheng.gz-cmc.com/json/topic/2024/06/14/42889fe292b04d0caee191fc36f8eb8e/{df_id.iloc[i,1]}_1.topicjson"
        headers = {"User-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"}
        response = requests.get(url=url_page,headers = headers)
        time.sleep(random.randint(1,2))
        if response.status_code == 200:
            response_json = response.json()
            with open(f'C:\\Users\\shrk-3121\\Desktop\\GDI数据爬取\\{df_id.iloc[i,0]}.json','w',encoding='utf8') as f2:
                json.dump(response_json,f2,ensure_ascii=False,indent=4)
    except Exception as e:
        print(e)


# In[183]:


# 依次读取专业类的json文件，并提取出其中的专业页面URL

path1 = r"C:\Users\shrk-3121\Desktop\GDI数据爬取"
file_list = glob.glob(path1 + "/*.json")

Name =[]
url = []

for file in file_list:
    with open(file,'r',encoding='utf8') as f2:
        data = json.load(f2)
        data_sub = data["dataList"]
    for i in data["dataList"]:
        Name.append(i["data"]["title"])
        url.append(i["data"]["url"])

tem = list(zip(Name,url))
df_maj = pd.DataFrame(tem)
col = ["专业名称","链接"]
df_maj.set_axis(col, axis=1,inplace=True)
df_maj.to_excel("GDI查询链接.xlsx",index = False)


# In[212]:


# 读取下载好的GDI专业排名链接
df_maj = pd.read_excel(r"C:\Users\shrk-3121\GDI查询链接.xlsx")


# In[221]:


# 下载专业排名图片保存到本地
for i in range(0,df.shape[0]):
    url_page = f"{df.iloc[i,1]}"
    headers = {"User-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"}
    try:
        response = requests.get(url=url_page,headers = headers)  
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")
            for i in soup.find(class_="insert-img-container"):
                url = (i["src"])
                r = requests.get(url)
                with open(f'C:\\Users\\shrk-3121\\Desktop\\GDI数据爬取\\专业排名图片爬取\\{soup.find("title").text}.png','wb') as f:
                    f.write(r.content) 
        time.sleep(random.randint(1,3))
    except Exception as e:
        print(e)


# In[ ]:




