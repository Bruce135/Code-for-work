import requests,json,time,os
import pandas as pd
from retrying import retry
import csv
import glob
import random
requests.packages.urllib3.disable_warnings()

# 设置代理，搭配本地Clash使用
proxy = '127.0.0.1:7890'
proxies = {'http': 'http://' + proxy,'https': 'http://' + proxy}

# 定义爬取函数
@retry(stop_max_attempt_number=8)
def rty_post(url,headers,body,proxies):
    response = requests.post(url=url, data=json.dumps(body), headers=headers, proxies = proxies,timeout=20, verify=False)
    return response

# 读取TOP期刊总表
path = r"C:\Users\shrk-3121\Desktop\爬虫相关文件夹\TOP论文数&明细采集\TOP期刊明细采集（8月）\AESTOP期刊完整列表.xlsx"
df = pd.read_excel(path,sheet_name = "AES")
#读取含地矿油ID文件
dky = pd.read_excel(r"C:\Users\shrk-3121\Desktop\爬虫相关文件夹\Incites数据爬取\地矿油数据\地矿油数据库ID.xlsx")

#依次取地矿油学校进行循环爬取明细数据
for name in range(0,dky.shape[0]):
     if not os.path.exists(f'C:\\Users\\shrk-3121\\Desktop\\爬虫相关文件夹\\Incites数据爬取\\地矿油TOP论文明细（0920）\\{dky.iloc[name,1]}'):
            os.makedirs(f'C:\\Users\\shrk-3121\\Desktop\\爬虫相关文件夹\\Incites数据爬取\\地矿油TOP论文明细（0920）\\{dky.iloc[name,1]}')
     for i in range(0,df.shape[0]):
        url_page = f"https://incites.clarivate.com/incites-app/explore/{dky.iloc[name,3]}/organization/data/table/page"
        headers = {"Host": "incites.clarivate.com",
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
                "Cookie": '_biz_uid=a7f36da978a242c9ce4a4d96bcc6a9fe; _vwo_uuid_v2=DE27F2972774062E6FEFDA7851983F4AE|f250727f864149b93a3fb3277f92bd38; _vwo_uuid=D3652980D183D6035F9225A698033EDF6; OptanonAlertBoxClosed=2024-04-10T05:21:50.881Z; _biz_flagsA=%7B%22Version%22%3A1%2C%22XDomain%22%3A%221%22%2C%22ViewThrough%22%3A%221%22%7D; ELOQUA=GUID=15831C3EFF2E4FC6AED1FC9D43733DB3; _gcl_au=1.1.1068835458.1723024182; _zitok=42d7d33c2cfc1d9f14781723024185; _ga_D5KRF08D0Q=GS1.2.1723465899.5.1.1723465971.0.0.0; _ga_3YQJ1BS28G=GS1.1.1725527443.12.1.1725527443.0.0.0; OptanonConsent=isGpcEnabled=0&datestamp=Mon+Sep+09+2024+16%3A28%3A28+GMT%2B0800+(%E9%A6%99%E6%B8%AF%E6%A0%87%E5%87%86%E6%97%B6%E9%97%B4)&version=202406.1.0&isIABGlobal=false&hosts=&consentId=72e6b7c0-9dfb-4708-a0a8-a349cdeff343&interactionCount=3&landingPath=NotLandingPage&groups=C0001%3A1%2CC0003%3A1%2CC0004%3A1%2CC0002%3A1&AwaitingReconsent=false&geolocation=HK%3BHCW&browserGpcFlag=0&isAnonUser=1; _vis_opt_s=2%7C; _biz_nA=90; _biz_pendingA=%5B%5D; _uetvid=b55e1e304eaa11eda0836344de5d9eee; _ga=GA1.2.1050746225.1666083698; _vwo_ds=3%3At_0%2Ca_0%3A0%241723024181%3A54.45508778%3A%3A54_0%2C53_0%2C52_0%2C51_0%2C50_0%2C25_0%3A3_0%2C2_0%3A0; _ga_9R70GJ8HZF=GS1.1.1725944781.38.0.1725944781.60.0.1018933095; _gid=GA1.2.1149996272.1726624814; SECCONTEXT=f53b741b1020db9adc8ff2581f48adf47f9abdd05d15f619b2d145974873684720bf8ef76e17de40ecae3b816dfe0c11c8b6122350963993d6456fa50a0eb8891c05b1d42db76e229cf5b1c2e0c73adde7f34e4d3bad24a7e498a86f3ce4f1c7052925ac25fedd8a114f724fe505a33bb50343438fa226ab67e028dd2daf928d00b127af901014a05feeb2f3db13f732aed56ee147f652b210c908eb2728ccbd65e6b0827d54c4bc244cc72f020dd25776fabb7a4b78b4a274c49cf576bc4247c26e0561045efb48e071cc612123589ad2dcc58321ef3d8abf0a7f125a01b1d1b94a676beca184a8e29d5516efda93fd7a88d994b92448d30db72057863719b7b5e951d1a2a8267a2c816a3675b7d35626bbd54e7698b3ef728d1cbb30477882113b71ab9cec18c986214ad656bcc3b7a78ee54f657d21a955fd4832ae5fb195; _sp_ses.2f26=*; USERNAME="kaikai.lin@shanghairanking.com"; STEAM_USER_ID="20896045"; truid="c8153180-792f-11ee-a502-6da1181491fa"; PSSID="H4-c8153180-792f-11ee-a502-6da1181491fa-a4ef27de-1bcc-42b8-a8bf-6245748157ad"; IC2_SID="H4-c8153180-792f-11ee-a502-6da1181491fa-a4ef27de-1bcc-42b8-a8bf-6245748157ad"; CUSTOMER_NAME="Shanghai Ranking Consultancy"; E_GROUP_NAME="IC2 Platform"; SUBSCRIPTION_GROUP_ID="817876"; SUBSCRIPTION_GROUP_NAME="Shanghai Ranking Consultancy Co Ltd_TRIAL07242637INC_1"; CUSTOMER_GROUP_ID="469453"; ROAMING_DISABLED="false"; ACCESS_METHOD="UNP"; firstName="Kaikai"; lastName="Lin"; userAuthType="TruidAuth"; userAuthIDType="c8153180-792f-11ee-a502-6da1181491fa"; _sp_id.2f26=23a17ced-878f-49e1-bbfe-f4f20e1b3b03.1666083697.446.1726798884.1726735717.f0c310cf-d365-4ec8-9323-873785304752; _ga_E0YH6TPFNF=GS1.2.1726798184.285.1.1726798884.0.0.0'}
        #获取论文总数参数
         body = {"filters":{"period":{"is":[2019,2023]},
            "orgtype":{"is":["Academic"]},
            "personIdTypeGroup":{"is":"authorRecord"},
            "geographicCollabType":{"is":"All"},
            "orgname":{"is":dky.iloc[name,2]},
            "articletype":{"is":df.iloc[i,4]},
            "schema":{"is":"Web of Science"},
            "earlyAccess":{"is":[1]},
            "issn":{"is":df.iloc[i,3]},
            "publisherType":{"is":"All"},
            "fundingAgencyType":{"is":"All"},
            "fundingDataSource":{"is":"All Sources"},
            "personIdType":{"is":"authorRecord"}},
            "pinned":[],
            "skip":0,
            "sortOrder":"desc",
            "take":10,
            "groupPinned":[],
            "sortBy":"orgName",
            "indicators":["key","seqNumber","hasProfile","esibimonthlyhotpapers","esibimonthlyhighlycitedpapers","orgName","wosDocuments","timesCited","norm","highlyCitedPapers"]}
        #获取论文明细参数
        body2 = {"filters":{"period":{"is":[2019,2023]},
            "orgtype":{"is":["Academic"]},
            "personIdTypeGroup":{"is":"authorRecord"},
            "geographicCollabType":{"is":"All"},
            "articletype":{"is":df.iloc[i,4]},
            "schema":{"is":"Web of Science"},
            "issn":{"is":df.iloc[i,3]},
            "earlyAccess":{"is":[1]},
            "publisherType":{"is":"All"},
            "fundingAgencyType":{"is":"All"},
            "fundingDataSource":{"is":"All Sources"},
            "personIdType":{"is":"authorRecord"}},
            "pinned":[]}
        #发送请求获取论文总数
        response = rty_post(url=url_page, headers=headers, body=body,proxies=proxies)
        response_json = response.json()
        #下载学校的论文明细并保存为Json格式
        for items in response_json["items"]:
            paper_number = items["wosDocuments"]["value"]
            paper_url = "https://incites.clarivate.com"+items["wosDocuments"]["data"]+f"&skip=0&take={paper_number}&sortBy=cites&sortOrder=desc"
            response = rty_post(url=paper_url, headers=headers, body=body2,proxies = proxies)
            paper_json = response.json()
            paper_json["orgName"] = items["orgName"]            
            with open(f'C:\\Users\\shrk-3121\\Desktop\\爬虫相关文件夹\\Incites数据爬取\\地矿油TOP论文明细（0920）\\{dky.iloc[name,1]}\\{df.iloc[i,0]}-{df.iloc[i,3]}.json','w',encoding='utf8') as f2:
                json.dump(paper_json,f2,ensure_ascii=False,indent=4)
                print(f'{dky.iloc[name,1]}-{df.iloc[i,0]}-{df.iloc[i,3]}下载完成')
#***************************************************************************************************************#
# 把下载好的Json明细文件汇总到一起
for i in range(0,dky.shape[0]):
    path1 = f"C:\\Users\\shrk-3121\\Desktop\\爬虫相关文件夹\\Incites数据爬取\\地矿油TOP数据明细爬取\\{dky.iloc[i,1]}"
    file_list = glob.glob( path1 + "/*.json")
    
    Year = []
    Title = []
    Url = []
    Source = []
    Authors = []
    Code = []
    Name = []

    for file in file_list:
        with open(file,'r',encoding='utf8') as f2:
            data = json.load(f2)
            data_sub = data["items"]
        for x in data_sub:
            Year.append(x["date"])
            Title.append(x["a"]["title"])
            Url.append(x["a"]["url"])
            Source.append(x["source"])
            Authors.append(x["authors"])
            Name.append(dky.iloc[i,1])
            Code.append(file.replace(f"C:\\Users\\shrk-3121\\Desktop\\爬虫相关文件夹\\Incites数据爬取\\地矿油TOP数据明细爬取\\{dky.iloc[i,1]}\\","").replace(".json",""))
    try:
        
        tem = list(zip(Name,Year,Title,Url,Source,Authors,Code))
        df1 = pd.DataFrame(tem)
        col = ["机构名称","年份","论文名称","链接","来源","作者","学科代码"]
        df1.set_axis(col, axis=1,inplace=True)
        with pd.ExcelWriter(f'C:\\Users\\shrk-3121\\Desktop\\爬虫相关文件夹\\Incites数据爬取\\地矿油TOP数据明细爬取\\{dky.iloc[i,1]}\\数据合并-{dky.iloc[i,1]}.xlsx',engine='xlsxwriter',engine_kwargs={'options':{'strings_to_urls': False}}) as writer:
                df1.to_excel(writer, index=False)
    except Exception as e:
        print(e)
        pass
        
