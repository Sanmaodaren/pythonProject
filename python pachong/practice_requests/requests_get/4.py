import requests

file = {'file': open('favicon.ico', 'rb')}
r = requests.post('https://www.httpbin.org/post', files=file)
print(r.text)