import re

content = 'adnjgfd4787853hjbfz'
content = re.sub('\d+','', content)
print(content)