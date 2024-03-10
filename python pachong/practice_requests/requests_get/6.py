import requests

r = requests.get('http://jwglxt.zua.edu.cn/eams/loginExt.action', auth=('2205010118', '1221chujiDA'))
print(r.status_code)