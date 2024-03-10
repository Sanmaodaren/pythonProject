import requests

headers = {
    'Cookie':'_octo=GH1.1.659230845.1693305428; logged_in=yes; dotcom_user=Sanmaodaren; '
             'color_mode=%7B%22color_mode%22%3A%22auto%22%2C%22light_theme%22%3A%7B%22name%22%3A%22light%22%2C'
             '%22color_mode%22%3A%22light%22%7D%2C%22dark_theme%22%3A%7B%22name%22%3A%22dark%22%2C%22color_mode%22%3A'
             '%22dark%22%7D%7D; preferred_color_mode=light; tz=Asia%2FShanghai'
}

r = requests.get('https://github.com/', headers=headers)
print(r.text)