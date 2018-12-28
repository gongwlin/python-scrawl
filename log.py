import requests
import time,random,json
from openpyxl import Workbook
from requests.exceptions import ReadTimeout, ConnectTimeout, HTTPError, ConnectionError, RequestException


def getOnePage(i):
    url = 'https://mp.weixin.qq.com/wxopen/wxaalarm?action=search_jserr'
    headers = {
    # 'Accept': 'application/json,text/javascript,*/*;q=0.01',
    # 'Connection': 'keep-alive',
    # 'Content-Length': '169',
    'Host': 'mp.weixin.qq.com',
    'Origin': 'https://mp.weixin.qq.com',
    'User - Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36',
    'Referer': 'https://mp.weixin.qq.com/wxopen/wxaalarm?action=get_jserr&token=1923478722&lang=zh_CN',
    'Cookie': 'pgv_pvi=7409933312; ptui_loginuin=1124884264; pt2gguin=o1124884264; RK=NvAEArOEaP; ptcz=0569f6441daaba023d5f55c237ec98a0724f851158da46b117d93da323113a21; ua_id=iAgnSy8N2U2yELjyAAAAAM9MFRQjyN7fMMZcsaOmK8c=; mm_lang=zh_CN; noticeLoginFlag=1; pgv_pvid=3499651865; remember_acct=yhxiaocx%40126.com; pgv_si=s2094211072; uuid=45695a919a7a9cf5cb28d6492f85c63c; slave_user=gh_f5cd32cf3467; bizuin=3508335167; ticket=7d017a4f1eb3837c2cb598718e69eaa2ef0f3ce6; ticket_id=gh_f5cd32cf3467; cert=vNxSwaG2hOroUcG1Zez9I_jyGsEiiJLj; data_bizuin=3508335167; data_ticket=ZxbM0UoBgx8ugQ150BE4hB45P/PpJJxnQUFxgX5Td79XKPY57pflKYNYAu6NCDAt; slave_sid=Z1A0MUZuVjl4UV9aanh0bkloYmh0eHhoWm5pTURobVB5OE41S3MyeVlsSmNoeFE4U2NFbDZuODk1ZzZKOVlvdG1HR0hBazdLbGtqdFRDSUFpaE81bWlUb1NwU3Joek9ES2h1TElvbUROdndOdm80ZHpqNkRIMjJJR3NtUUpTeDJmVHpORHJCNDJZckxYQ2dS; xid=5cc377eeb7ac0feabf41d4e8e6c54b7d; openid2ticket_oP_Ic0f-wgtahqlIiwiVqDTk2Wxc=PqB87GQs5j5iOFw+fGiq8S9uemvHWG5qlAKmTGeQOyY=; pgv_info=ssid=s6698425776; wx_csrf_cookie=f6111c56b4a1b8fb2634a3dabf811135; ts_last=mp.weixin.qq.com/wxopen/frame; ts_uid=9456698730',
    'X-Requested-With': 'XMLHttpRequest'
    }
    data = {
        'token': '********',
        'lang': 'zh_CN',
        'f': 'json',
        'ajax': 1,
        'random': random.random(),
        'errmsg_keyword': '',
        'type': 1,
        'client_version': '',
        'start_time': 1532253600,
        'end_time': 1533204000,
        'start': i,
        'limit': 10
    }
    print(data['start'])
    try:
        res = requests.post(url, headers=headers, data=data,timeout=40)
        print(res.text)
        return json.loads(res.text)
    # 规定时间内未响应就抛出异常
    except ReadTimeout as e:
        print('请求超时')
    except ConnectionError as e:
        print('连接失败')
    except RequestException as e:
        print('请求失败')

    # req = requests.post(url, headers=headers, data=data)
    # print(req.text)
    # if req.status_code == '200':
    #     return json.loads(req.text)
    # else:
    #     print("Error")

#data = {"base_resp":{"err_msg":"ok","ret":0},"results":[{"app_version":"112","client_version":"6.6.0","digest":"e9261f0464b5e32b58bb90741fb2eebe8167ea82","errmsg":"Cannot convert undefined or null to object;at pages/index/index onHide function;at api getClipboardData success callback function","error_stack_buffid_list":[],"error_stack_md5_list":[],"time":1532448000,"total_error_cnt":1,"version_error_cnt":0},{"app_version":"114","client_version":"6.6.7","digest":"6e679cbb1f69cee3f4d6523c8621b7fe61317584","errmsg":"Cannot convert undefined or null to object;at App onHide function;at api getClipboardData success callback function","error_stack_buffid_list":[],"error_stack_md5_list":[],"time":1532448000,"total_error_cnt":0,"version_error_cnt":0},{"app_version":"114","client_version":"6.6.3","digest":"8fb49d3608784ab190bd524cf4637527ec7888c0","errmsg":"undefined is not an object (evaluating 't.dispatch((0,i.initCarts)()).then'); [Component] Event Handler Error @ pages/carts/index#bound value","error_stack_buffid_list":[],"error_stack_md5_list":[],"time":1532448000,"total_error_cnt":32,"version_error_cnt":0},{"app_version":"114","client_version":"6.5.9","digest":"8df145c9d883179f69843a5f78b7a82fa90c56f0","errmsg":"TypeError: undefined is not an object (evaluating 'getApp().window.screenHeight')","error_stack_buffid_list":[],"error_stack_md5_list":[],"time":1532448000,"total_error_cnt":7,"version_error_cnt":0},{"app_version":"114","client_version":"6.7.1","digest":"fc5709996f18378e8d4e91aac118f17326481d35","errmsg":"null is not an object (evaluating 'i.num'); [Component] Event Handler Error @ onlinePages/cartssub/index#(anonymous)","error_stack_buffid_list":[],"error_stack_md5_list":[],"time":1532448000,"total_error_cnt":0,"version_error_cnt":0},{"app_version":"113","client_version":"6.6.1","digest":"16f3eb7df29b0c6cc25d9708e7d3c273f5175d54","errmsg":"Cannot read property 'storeid' of undefined;at pages/carts/index page onPullDownRefresh function","error_stack_buffid_list":[4099108276551722],"error_stack_md5_list":["b79dec9adc7b984c38d2acc0650be427"],"time":1532448000,"total_error_cnt":6,"version_error_cnt":3},{"app_version":"113","client_version":"6.6.1","digest":"1a2ebf26d555b9b0869a413b2e409481e5ea28d1","errmsg":"h.default.seteExtraAnalyticsData is not a function;at \"onlinePages/orderdetail/index\" page lifeCycleMethod onHide function","error_stack_buffid_list":[4283860655357994],"error_stack_md5_list":["22b791a40e4f5437803c6fe48bdb07db"],"time":1532448000,"total_error_cnt":5,"version_error_cnt":2}],"total":1767}

#将Unix时间戳转为时间
def timestamp_datetime(value):
    # format = '%Y-%m-%d %H:%M:%S'
    format = '%Y-%m-%d'
    # value为传入的值为时间戳(整形)，如：1332888820
    value = time.localtime(value)
    ## 经过localtime转换后变成
    ## time.struct_time(tm_year=2012, tm_mon=3, tm_mday=28, tm_hour=6, tm_min=53, tm_sec=40, tm_wday=2, tm_yday=88, tm_isdst=0)
    # 最后再经过strftime函数转换为正常日期格式。
    dt = time.strftime(format, value)
    return dt


#保存到Excel表中
def saveToExcel(data):
    # data = {"base_resp":{"err_msg":"ok","ret":0},"results":[{"app_version":"112","client_version":"6.6.0","digest":"e9261f0464b5e32b58bb90741fb2eebe8167ea82","errmsg":"Cannot convert undefined or null to object;at pages/index/index onHide function;at api getClipboardData success callback function","error_stack_buffid_list":[],"error_stack_md5_list":[],"time":1532448000,"total_error_cnt":1,"version_error_cnt":0},{"app_version":"114","client_version":"6.6.7","digest":"6e679cbb1f69cee3f4d6523c8621b7fe61317584","errmsg":"Cannot convert undefined or null to object;at App onHide function;at api getClipboardData success callback function","error_stack_buffid_list":[],"error_stack_md5_list":[],"time":1532448000,"total_error_cnt":0,"version_error_cnt":0},{"app_version":"114","client_version":"6.6.3","digest":"8fb49d3608784ab190bd524cf4637527ec7888c0","errmsg":"undefined is not an object (evaluating 't.dispatch((0,i.initCarts)()).then'); [Component] Event Handler Error @ pages/carts/index#bound value","error_stack_buffid_list":[],"error_stack_md5_list":[],"time":1532448000,"total_error_cnt":32,"version_error_cnt":0},{"app_version":"114","client_version":"6.5.9","digest":"8df145c9d883179f69843a5f78b7a82fa90c56f0","errmsg":"TypeError: undefined is not an object (evaluating 'getApp().window.screenHeight')","error_stack_buffid_list":[],"error_stack_md5_list":[],"time":1532448000,"total_error_cnt":7,"version_error_cnt":0},{"app_version":"114","client_version":"6.7.1","digest":"fc5709996f18378e8d4e91aac118f17326481d35","errmsg":"null is not an object (evaluating 'i.num'); [Component] Event Handler Error @ onlinePages/cartssub/index#(anonymous)","error_stack_buffid_list":[],"error_stack_md5_list":[],"time":1532448000,"total_error_cnt":0,"version_error_cnt":0},{"app_version":"113","client_version":"6.6.1","digest":"16f3eb7df29b0c6cc25d9708e7d3c273f5175d54","errmsg":"Cannot read property 'storeid' of undefined;at pages/carts/index page onPullDownRefresh function","error_stack_buffid_list":[4099108276551722],"error_stack_md5_list":["b79dec9adc7b984c38d2acc0650be427"],"time":1532448000,"total_error_cnt":6,"version_error_cnt":3},{"app_version":"113","client_version":"6.6.1","digest":"1a2ebf26d555b9b0869a413b2e409481e5ea28d1","errmsg":"h.default.seteExtraAnalyticsData is not a function;at \"onlinePages/orderdetail/index\" page lifeCycleMethod onHide function","error_stack_buffid_list":[4283860655357994],"error_stack_md5_list":["22b791a40e4f5437803c6fe48bdb07db"],"time":1532448000,"total_error_cnt":5,"version_error_cnt":2}],"total":1767}
    # sheet.title = "Sheet1"
    global row
    if 'results' in data:
        #print("长度为:")
        #print(len(data['results']))
        for i in data['results']:
            print(i)
            sheet["A%d" % (row+1)].value = timestamp_datetime(i['time'])
            sheet["B%d" % (row+1)].value = i['client_version']
            sheet["C%d" % (row+1)].value = i['app_version']
            sheet["D%d" % (row+1)].value = i['version_error_cnt']
            sheet["E%d" % (row+1)].value = i['total_error_cnt']
            sheet["F%d" % (row+1)].value = i['errmsg']
            print('第%d行已完成'%row)
            row += 1


row = 0  # 全局变量，Excel行的下标
wb = Workbook()  #全局变量
sheet = wb.active

if __name__ == '__main__':
    num = 353
    for i in range(1,num+1):
        data = getOnePage(i)
        saveToExcel(data)
        print('第%d页已完成'%i)
        time.sleep(random.randint(2, 7))
    wb.save('yc.xlsx')


