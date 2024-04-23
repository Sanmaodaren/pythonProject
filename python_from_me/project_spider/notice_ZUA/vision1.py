import datetime
import re
import time as t

import requests
import schedule


def get_info(url):
    """从网页源码中提取活动名称和详情页网址"""
    response = requests.get(url)
    response.encoding = response.apparent_encoding
    html = response.text
    pattern = "<a href=\"..(.*?)\" target=\"_blank\" title"
    info = re.findall(pattern, html)  # 正则提取网址
    pattern2 = "<span class=\"time\">(.*?)</span>"
    info2 = re.findall(pattern2, html)  # 正则提取发布日期
    pattern3 = "a href=.*?target=.*? title=.*?>(.*?)</a>"
    info3 = re.findall(pattern3, html, re.S)  # 正则提取标题
    info_list = []
    x = len(info)
    for i in range(x):
        my_data = {
            "href": "https://sczx.zua.edu.cn" + info[i],
            "date": info2[i],
            "title": info3[i]
        }
        info_list.append(my_data)
    return info_list


def judge_part(mydata):
    """根据日期判断是否推送"""
    date_today = datetime.date.today()
    yesterday = date_today - datetime.timedelta(days=3)
    title_date = mydata["date"]
    time1 = datetime.datetime.strptime(str(yesterday), "%Y-%m-%d")
    time2 = datetime.datetime.strptime(str(title_date), "%Y-%m-%d")
    if time2 == time1:
        return 1
    else:
        return 0


def sendMessage(token, mydata):
    """推送模块，需要token和推送内容"""
    baseurl = "http://wx.xtuis.cn/"
    url = baseurl + token + ".send"

    data = {
        "text": mydata["title"], "desp": '内容:' + mydata["title"] + '<br>' \
                                                                     '网址:' + mydata["href"] + '<br>' \
                                                                                                '发布日期:' + mydata[
                                             "date"]}
    requests.post(url, data=data)


def execute(tokens):
    for my_data in my_datas:
        judge_part(my_data)
        num = judge_part(my_data)
        while num:
            for token in tokens:
                sendMessage(token, my_data)
                t.sleep(2)
            num = 0


tokens = ['48YMrNu8g32XNNtDIhseqQ7NU']
tokens.append('x3sOMNv9OxZHldsRrmLSlvRUO')  # 严圣淇
base_url = 'https://sczx.zua.edu.cn/index/jstz.htm'
my_datas = get_info(base_url)
schedule.every().day.at("07:24").do(execute)
try:
    while True:
        schedule.run_pending()
        t.sleep(1)
except:
    KeyboardInterrupt
