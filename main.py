import requests, sys, json, time, winsound

from win32com.client import Dispatch

speaker = Dispatch("SAPI.SpVoice")

speaker.Speak("检测脚本启动")
code = sys.argv[1]
price_up = float(sys.argv[2])
price_down = float(sys.argv[3])
session = requests.session()


def get():
    global session, code, price_down, price_up
    url = "https://stock.xueqiu.com/v5/stock/realtime/quotec.json?symbol=" + code + "&_=1661133743660"
    header_dict = {
        "accept": "application/json, text/plain, */*",
        "accept-encoding": "gzip, deflate, br",
        "accept-language": "accept-language",
        "cookie": "device_id=23a0db6ec944ccdaa408aff3b0fd533c; xq_a_token=28ed0fb1c0734b3e85f9e93b8478033dbc11c856; xqat=28ed0fb1c0734b3e85f9e93b8478033dbc11c856; xq_r_token=bf8193ec3b71dee51579211fc4994d03f17c64ac; xq_id_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJ1aWQiOi0xLCJpc3MiOiJ1YyIsImV4cCI6MTY2MzExMzIyMSwiY3RtIjoxNjYxMTMwNTgzNTk4LCJjaWQiOiJkOWQwbjRBWnVwIn0.A3ckUMjlFXybVuX9lCgnXSEmOECqGarcb8I7eQpsJyWf4VZztEca5G9M3JTlCepdiwku-trbAQbQ3-mzL7tir0druBq86XQHZWYEsSz3Igh68kKrSXj_JaglNkGRffZNd0pezvAVU6WgMYgnmeKY6z2RyXEQChT3Pf6zbXqPGF6TRI0HLpPd8X0E8CJl_CnUYaweAEZBgBQLBQ2SEKslzdzvNHvDNvBmppIWOKKA5P1-Bcf_0LTEV7p3mlDSVKumljBvUpUs4IFzdsdOZLSoCU3YHuHS-ecnSphG5pyiXWCps_F3SbOPr3-zVBHdZsd-KCDz8qWgxXa3wNwWNpswNg; u=701661130626725; Hm_lvt_1db88642e346389874251b5a1eded6e3=1659880074,1661130629; Hm_lpvt_1db88642e346389874251b5a1eded6e3=1661130662",
        "origin": "https://xueqiu.com",
        "referer": "https://xueqiu.com/S/" + code,
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.5112.81 Safari/537.36 Edg/104.0.1293.54"
    }
    respond = session.get(url=url, headers=header_dict)
    return respond


def deal_info(respond):
    if respond.status_code == 200:
        back_dict = json.loads(respond.content)
        price_now = back_dict["data"][0]["current"]
        percent = back_dict["data"][0]["percent"]
        if percent < 0:
            print("\033[1;32m", str(price_now) + "  " + str(percent) + "%", "\033[0m")

            # winsound.MessageBeep(1111)
        else:
            print("\033[0;31;40m", str(price_now) + "  " + str(percent) + "%", "\033[0m")
        if price_now > price_up:
            speaker.Speak("股票上涨， 现价" + str(price_now))
        elif price_now < price_down:
            speaker.Speak("股票下跌， 现价" + str(price_now))


def main():
    while True:
        try:
            deal_info(get())
            time.sleep(3)
        except Exception as e:
            pass


main()