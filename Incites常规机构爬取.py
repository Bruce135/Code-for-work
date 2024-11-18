#!/usr/bin/env python
# coding: utf-8

# In[1]:


import requests,json,time,os
import pandas as pd
from retrying import retry
import csv
import glob
requests.packages.urllib3.disable_warnings()


# In[2]:


proxy = '127.0.0.1:7890'
proxies = {'http': 'http://' + proxy,'https': 'http://' + proxy}
# 定义爬取函数
@retry(stop_max_attempt_number=8)
def rty_post(url,headers,body,proxies):
    response = requests.post(url=url, data=json.dumps(body), headers=headers, proxies = proxies,timeout=20, verify=False)
    return response


# In[3]:


# 读取Q1期刊总表
path = r"C:\Users\shrk-3121\Desktop\爬虫相关文件夹\Incites数据爬取\GRAS1-2023 学科分类.xlsx"
# df = pd.read_excel(path)
# 根据学科代码透视出该学科下所有的ISSN号
# df = df.groupby('code')['ISSN'].apply(list).reset_index()


# In[4]:


df = pd.read_excel(path,sheet_name = "完整学科映射")


# In[5]:


df = df.groupby('学科代码')['WOS学科英文名'].apply(list).reset_index()


# In[12]:


for i in range(0,df.shape[0]):
    # 各学校论文数页面
    url_page = f"https://incites.clarivate.com/incites-app/explore/0/organization/data/table/page"
    # 指定学科下论文总数页面
    url_total = f"https://incites.clarivate.com/incites-app/explore/0/organization/data/table/total"
    
    headers = {
                "Host": "incites.clarivate.com",
                "Connection": "keep-alive",
                "Content-Length": "755",
                "sec-ch-ua": '"Not.A/Brand";v="8", "Chromium";v="114", "Microsoft Edge";v="114"',
                "Accept": "application/json, text/plain, */*",
                "Content-Type": "application/json",
                "Accept-Language": "zh",
                "sec-ch-ua-mobile": "?0",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.67",
                "sec-ch-ua-platform": '"Windows"',
                "Origin": "https://incites.clarivate.com",
                "Sec-Fetch-Site": "same-origin",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Dest": "empty",
                "Referer": "https://incites.clarivate.com/zh/",
                "Accept-Encoding": "gzip, deflate, br",
                # Cookie需要在incites登录后获取
                "Cookie": '_biz_uid=a7f36da978a242c9ce4a4d96bcc6a9fe; _vwo_uuid_v2=DE27F2972774062E6FEFDA7851983F4AE|f250727f864149b93a3fb3277f92bd38; _vwo_uuid=D3652980D183D6035F9225A698033EDF6; OptanonAlertBoxClosed=2024-04-10T05:21:50.881Z; _biz_flagsA=%7B%22Version%22%3A1%2C%22XDomain%22%3A%221%22%2C%22ViewThrough%22%3A%221%22%7D; ELOQUA=GUID=15831C3EFF2E4FC6AED1FC9D43733DB3; _gcl_au=1.1.1068835458.1723024182; _zitok=42d7d33c2cfc1d9f14781723024185; _ga_D5KRF08D0Q=GS1.2.1723465899.5.1.1723465971.0.0.0; _ga_3YQJ1BS28G=GS1.1.1725527443.12.1.1725527443.0.0.0; _gid=GA1.2.732950021.1725849087; OptanonConsent=isGpcEnabled=0&datestamp=Mon+Sep+09+2024+16%3A28%3A28+GMT%2B0800+(%E9%A6%99%E6%B8%AF%E6%A0%87%E5%87%86%E6%97%B6%E9%97%B4)&version=202406.1.0&isIABGlobal=false&hosts=&consentId=72e6b7c0-9dfb-4708-a0a8-a349cdeff343&interactionCount=3&landingPath=NotLandingPage&groups=C0001%3A1%2CC0003%3A1%2CC0004%3A1%2CC0002%3A1&AwaitingReconsent=false&geolocation=HK%3BHCW&browserGpcFlag=0&isAnonUser=1; _vis_opt_s=2%7C; _biz_nA=90; _biz_pendingA=%5B%5D; _uetvid=b55e1e304eaa11eda0836344de5d9eee; _ga=GA1.2.1050746225.1666083698; _vwo_ds=3%3At_0%2Ca_0%3A0%241723024181%3A54.45508778%3A%3A54_0%2C53_0%2C52_0%2C51_0%2C50_0%2C25_0%3A3_0%2C2_0%3A0; _ga_9R70GJ8HZF=GS1.1.1725944781.38.0.1725944781.60.0.1018933095; _sp_ses.2f26=*; USERNAME="kaikai.lin@shanghairanking.com"; STEAM_USER_ID="20896045"; truid="c8153180-792f-11ee-a502-6da1181491fa"; PSSID="H4-c8153180-792f-11ee-a502-6da1181491fa-569d77fc-86ab-4ba1-8b1a-0bcdba72eada"; IC2_SID="H4-c8153180-792f-11ee-a502-6da1181491fa-569d77fc-86ab-4ba1-8b1a-0bcdba72eada"; CUSTOMER_NAME="Shanghai Ranking Consultancy"; E_GROUP_NAME="IC2 Platform"; SUBSCRIPTION_GROUP_ID="817876"; SUBSCRIPTION_GROUP_NAME="Shanghai Ranking Consultancy Co Ltd_TRIAL07242637INC_1"; CUSTOMER_GROUP_ID="469453"; ROAMING_DISABLED="false"; ACCESS_METHOD="UNP"; firstName="Kaikai"; lastName="Lin"; userAuthType="TruidAuth"; userAuthIDType="c8153180-792f-11ee-a502-6da1181491fa"; SECCONTEXT=32904028ca9f532d29c2a45ba5748f2eb0a253f72efc80b1d0a1bc08107ca97712bfe6b890f72b3576190a867ffa4a75365d17077410bc6f554c14bf54c386d9204a8fe886d46cf62a45d305d2d9e16c70d729806f1ebe7bdefe4324f88677b07e399c322d4fde73308a398f9c1f5184d8f6909e09606aef2b8ea11cd23117a908bad710a92387d2d72b42fd940eccfb8219ec387567855000a6b2aa166ffd2381bcec466f3ce6abd9c4cae0ea63f18c4ddbf3895b09fbb8bdcce9c0281ebb0ce48732e1c7709e22b4dbbcc21cde8918f3d9790a8cefe735e48b26dd140e94f1a8c9d424afbebc68d10d7e8348c44f78f1e01a24a6f1e4f5e7c535ee675afd2f04647ad2fe55c0c439f811f96a01bff1db7e84c8737c1f5b51547d205196c169be7ea36b59ba58a82a36325111e7bc7882ff8f7880779c5fdb88ac0ae220901a; _ga_E0YH6TPFNF=GS1.2.1726107714.269.1.1726107919.0.0.0; _sp_id.2f26=23a17ced-878f-49e1-bbfe-f4f20e1b3b03.1666083697.430.1726107945.1726047689.4026bb3e-78ed-4532-b300-faf456a7e261'}
    
    body = {"filters":{"period":{"is":[2019,2023]},
        "orgtype":{"is":["Academic"]},
        "personIdTypeGroup":{"is":"authorRecord"},
        "geographicCollabType":{"is":"All"},
        "orgname":{"is":"Universite Paris Cite"},
        "articletype":{"is":"Article"},
        "sbjname":{"is": df.iloc[i,1]},
        "schema":{"is":"Web of Science"},
        "earlyAccess":{"is":[1]},
#         "issn":{"is":df.iloc[i,1]},
        "publisherType":{"is":"All"},
        "fundingAgencyType":{"is":"All"},
        "fundingDataSource":{"is":"All Sources"},
        "personIdType":{"is":"authorRecord"}},
        "pinned":[],
        "skip":0,
        "sortOrder":"desc",
        "take":400,
        "groupPinned":[],
        "sortBy":"orgName",
        "indicators":["key","seqNumber","hasProfile","esibimonthlyhotpapers","esibimonthlyhighlycitedpapers","orgName","wosDocuments","timesCited","norm","highlyCitedPapers"]}

    response = rty_post(url=url_total, headers=headers, body=body,proxies = proxies)
    total = response.json()["totalItems"]
    print(f'捕获到{df.iloc[i,0]},论文数：',end=" ")
    print(total)
    for j in range(0,total,400):
        body["skip"] = j
        response = rty_post(url=url_page, headers=headers, body=body,proxies = proxies)
        response_json = response.json()
        # 按照学科，分别保存学校论文数数据到json文件
        with open(f'C:\\Users\\shrk-3121\\Desktop\\巴黎西岱大学2018-2022Incites数据明细\\{df.iloc[i,0]}.json','w',encoding='utf8') as f2:
            json.dump(response_json,f2,ensure_ascii=False,indent=4)


# In[13]:


# 读取爬取好的json文件夹下所有链接
path1 = r"C:\Users\shrk-3121\Desktop\巴黎西岱大学2018-2022Incites数据明细"
file_list = glob.glob(path1 + "/*.json")


# In[16]:


Name = []
Pub = []
CNCI = []
IC = []
Code = []
Q1 = []

for file in file_list:
    with open(file,'r',encoding='utf8') as f2:
        data = json.load(f2)
        data_sub = data["items"]
    for x in data_sub:
        Name.append(x["orgName"])
        Pub.append(x["wosDocuments"]["value"])
        CNCI.append(x["norm"])
        IC.append(x["prcntIntCollab"])
        Q1.append(x["jifdocsq1"]["value"])
        Code.append(file.replace(f"C:\\Users\\shrk-3121\\Desktop\\巴黎西岱大学2018-2022Incites数据明细\\","").replace(".json",""))


tem = list(zip(Name,Loc,Pub,Q1,CNCI,IC,Code))
df1 = pd.DataFrame(tem)
col = ["机构名称","国家地区","Pub","Q1","CNCI","IC","学科代码"]
df1.set_axis(col, axis=1,inplace=True)


# In[14]:


# 保存文件到excel
tem = list(zip(Name,Num,Code))
df1 = pd.DataFrame(tem)
col = ["机构名称","论文数","学科代码"]
df1.set_axis(col, axis=1,inplace=True)
with pd.ExcelWriter(r'C:\Users\shrk-3121\Desktop\爬虫相关文件夹\Incites数据爬取\TOP数据爬取更新（Q1）\常规数据\数据合并-常规.xlsx',engine='xlsxwriter',engine_kwargs={'options':{'strings_to_urls': False}}) as writer:
        df1.to_excel(writer, index=False)


# In[ ]:




