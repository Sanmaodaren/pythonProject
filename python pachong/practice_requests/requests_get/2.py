import requests

date = {
    'name':'germey',
    'age':25
}
r = requests.get('https://httpbin.org/get', params=date)
print(r.text)