import requests
from urllib import parse
import json


sess  = requests.session()
headers = {
    'Host': 'newsale.chnl.zj.chinamobile.com',
    'Connection': 'keep-alive',
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36 MicroMessenger/7.0.9.501 NetType/WIFI MiniProgramEnv/Windows WindowsWechat',
    'content-type': 'application/json',
    'Referer': 'https://servicewechat.com/wxf26e1816ed35f0fd/83/page-frame.html',
    'Accept-Encoding': 'gzip, deflate, br',
}
sess.headers = headers


# url = "https://newsale.chnl.zj.chinamobile.com/newsale-web/wechat/validateUserInfo2"
# post_data = {"bill": "13486731435"}
# response0 = sess.post(url, data=post_data)

# url = "https://newsale.chnl.zj.chinamobile.com/newsale-web/wechat/qryMarketKindById"
# post_data = {"marketId":"600000661055"}
# response1 = sess.post(url, data=post_data, verify=False).text
# print(response1)



phoneNo = "13486731435"

link = "https://newsale.chnl.zj.chinamobile.com/newsale-web/wechat/queryStoreInfoByCode"
# phoneList = [
# "13676598904",
# "13456093683",
# "15058337939"
# ]
# for phone in phoneList:

formData = {"storeCode": "40072035"}

headers = {
    'Connection': 'keep-alive',
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36 MicroMessenger/7.0.9.501 NetType/WIFI MiniProgramEnv/Windows WindowsWechat',
    'content-type': 'application/json',
    'Referer': 'https://servicewechat.com/wxf26e1816ed35f0fd/83/page-frame.html',
    'Accept-Encoding': 'gzip, deflate, br'
}
print(sess.post(link, data=json.dumps(formData), headers=headers, verify=False).text)

    # 40088710 20088322
    # 40072035 20118569

