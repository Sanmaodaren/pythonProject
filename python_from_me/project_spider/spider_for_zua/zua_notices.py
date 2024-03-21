import requests
import re
import datetime

base_url = "https://www.zua.edu.cn/list.jsp?urltype=tree.TreeTempUrl&wbtreeid=1066"

def get_info(url):
    """从网页源码中提取活动名称和详情页网址"""
    response = requests.get(url)
    html = response.text
    pattern = "<li style.*?><a href=\"(.*?)\" target.*?title=\"(.*?)\" style"
    info = re.findall(pattern, html)  # 正则提取名称和网址
    pattern2 = "<span style=.*?55\">(.*?)</span></li>"
    info2 = re.findall(pattern2, html)  # 正则提取发布时间
    i = 0
    # 三类信息放在一个列表里
    info_list = []
    for x in info:
        x = list(x)
        x.append(info2[i])
        i += 1
        my_data = {
            "href": 'https://www.zua.edu.cn/' + x[0],
            "title": x[1],
            "date": x[2]
        }
        info_list.append(my_data)
    return info_list

def judge_part(mydata):
    date_today = datetime.date.today()
    title_date = mydata["date"]
    time1 = datetime.datetime.strptime(str(date_today), "%Y-%m-%d")
    time2 = datetime.datetime.strptime(str(title_date), "%Y-%m-%d")
    if time2 >= time1:
        return 1
    else:
        return 0

def sendMessage(token, mydata):
    baseurl = "http://wx.xtuis.cn/"
    url = baseurl + token + ".send"

    data = {
        "text": mydata["title"],"desp": '内容:' + mydata["title"] + '<br>'\
        '网址:' + mydata["href"] + '<br>'\
        '发布日期:' + mydata["date"]}
    requests.post(url, data=data)


token = '48YMrNu8g32XNNtDIhseqQ7NU'
for page in range(1,3):  # 构造不同页的url
    url = (f"https://www.zua.edu.cn/list.jsp?a126128t=73&a126128p={page}"
           f"&a126128c=10&urltype=tree.TreeTempUrl&wbtreeid=1066")
    my_datas = get_info(url)  # 提取页包含的十条信息，为一个列表
    for my_data in my_datas:
        judge_part(my_data)
        num = judge_part(my_data)
        while num:
            sendMessage(token, my_data)
            num = 0
