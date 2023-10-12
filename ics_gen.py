import requests
from datetime import datetime
from urllib import parse
import time
import requests
from bs4 import BeautifulSoup
from urllib.parse import quote
import re
from datetime import datetime

username = "" # 学号
password = "" # 密码

BASE_URL = 'http://app.buaa.edu.cn/'
LOGIN_URL = 'https://sso.buaa.edu.cn/login'

def get_login_token(session):
    r = session.get("https://sso.buaa.edu.cn/login?service=https%3A%2F%2Fapp.buaa.edu.cn%2Fa_buaa%2Fapi%2Fcas%2Findex%3Fredirect%3D%252Fsite%252Fcenter%252Fpersonal%26from%3Dwap%26login_from%3D&noAutoRedirect=1")
    soup = BeautifulSoup(r.content, 'html.parser')
    return soup.find('input', {'name': 'execution'})['value']

def login(session):
    formdata = {
        'username': username,
        'password': password,
        'execution': get_login_token(session),
        'type': 'username_password',
        '_eventId': 'submit',
        'submit': '登陆'
    }
    r = session.post(LOGIN_URL, data=formdata, allow_redirects=False)
    location = r.headers.get('Location')
    if location:
        return location
    else:
        exit('登陆失败，请自行排错')

def get_eai_sess() -> str:
    url = 'https://app.buaa.edu.cn/uc/wap/login'
    r = requests.get(url, allow_redirects=False)
    cookies = requests.utils.dict_from_cookiejar(r.cookies)
    return cookies["eai-sess"]

def verify_eai_sess():
    session = requests.session()
    url = login(session)
    eai_sess = get_eai_sess()

    hearders ={
        'cookie': eai_sess
    }
    r = requests.get(url,headers=hearders,allow_redirects=False)
    cookies = requests.utils.dict_from_cookiejar(r.cookies)
    date = r.headers.get('date')
    if date:
        return cookies["eai-sess"]
    else:
        exit('以上内容出错，请自行排错')

def merge_adjacent_classes(classes: list) -> list:
    i = 1
    while i < len(classes):
        if classes[i]["course_name"] == classes[i-1]["course_name"]:
            classes[i-1]["lessons"] += f'{classes[i]["lessons"]}'
            classes[i-1]["course_time"] = f'{classes[i-1]["course_time"].split("～")[0]}～{classes[i]["course_time"].split("～")[1]}'
            classes.pop(i)
        else:
            i += 1
    return classes

def get_class_by_week(year: str, term: str, week: str, eai_sess: str) -> list[dict]:
    class_url = "https://app.buaa.edu.cn/timetable/wap/default/get-datatmp"
    header = {
        "Cookie": f"eai-sess={eai_sess}",
    }
    data = {
        "year": year,
        "term": term,
        "week": week,
        "type": "1",
    }

    r = requests.post(class_url, data=data, headers=header)
    days = r.json()["d"]["weekdays"]
    classes = r.json()["d"]["classes"]
    classes = merge_adjacent_classes(sorted(classes, key=lambda klass: int(klass["weekday"]) * 100 + int(klass["lessons"][0:1])))

    for klass in classes:
        klass["date"] = days[int(klass["weekday"]) - 1].replace("-", "")
        klass["start"] = klass["course_time"].split("～")[0].replace(":", "")
        klass["end"] = klass["course_time"].split("～")[1].replace(":", "")
        klass["lessons"] = ", ".join([klass["lessons"][i:i+2] for i in range(0, len(klass["lessons"]), 2)])

    return classes

def generate_ics(title: str, classes: list) -> str:
    ics_payload = f"""BEGIN:VCALENDAR
VERSION:2.0
X-WR-CALNAME:{title}
CALSCALE:GREGORIAN
BEGIN:VTIMEZONE
TZID:Asia/Shanghai
TZURL:http://tzurl.org/zoneinfo-outlook/Asia/Shanghai
X-LIC-LOCATION:Asia/Shanghai
BEGIN:STANDARD
TZOFFSETFROM:+0800
TZOFFSETTO:+0800
TZNAME:CST
DTSTART:19700101T000000
END:STANDARD
END:VTIMEZONE"""

    for klass in classes:
        event_start = f'{klass["date"]}T{klass["start"].rjust(4, "0")}00'
        event_end = f'{klass["date"]}T{klass["end"].rjust(4, "0")}00'
        event_description = f"""编号：{klass["course_id"]}
名称：{klass["course_name"]}
教师：{klass["teacher"]}
学分：{klass["credit"]}
类型：{klass["course_type"]}
课时：{klass["course_hour"]}
上课时间：{klass["course_time"]}；第 {klass["lessons"]} 节""".replace("\n", "\\n")
        event_info = f"""
BEGIN:VEVENT
DESCRIPTION:{event_description}
DTSTART;TZID=Asia/Shanghai:{event_start}
DTEND;TZID=Asia/Shanghai:{event_end}
LOCATION:{klass["location"]}
SUMMARY:{klass["course_name"]}
BEGIN:VALARM
TRIGGER:-PT30M
REPEAT:1
DURATION:PT1M
END:VALARM
END:VEVENT"""
        ics_payload += event_info

    ics_payload += "\nEND:VCALENDAR"
    return ics_payload

if __name__ == "__main__":
    year = datetime.now().year
    month = datetime.now().month
    year_str = f"{year-1}-{year}" if month < 7 else f"{year}-{year+1}"
    term_str = "2" if month < 7 else "1"
    eai_sess = verify_eai_sess()

    weekly_classes = []
    for week in range(1, 20):
        week_str = str(week)
        weekly_classes.extend(get_class_by_week(year_str, term_str, week_str, eai_sess))

    with open("classes.ics", "w+", encoding="utf-8") as f:
        f.write(generate_ics(f"北航 {year_str} 第 {term_str} 学期课程表", weekly_classes))