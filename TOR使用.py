# -*- coding:utf-8 -*-

import socket,re,os
import socks
import requests
from bs4 import BeautifulSoup
from stem import Signal
from stem.control import Controller
import threading
import queue as Queue
import openpyxl as excel

controller = Controller.from_port(port=9151)
controller.authenticate()

socks.set_default_proxy(socks.SOCKS5, "127.0.0.1", 9150)
socket.socket = socks.socksocket


def put_cai_wu_value(path,value_content):
    # 打开excel
    name=os.path.join(path)
    wb=excel.load_workbook(filename=name)
    # 获取并打开工作册
    sheets = wb.get_sheet_names()
    ws = wb.get_sheet_by_name(sheets[0])



    for i in range(len(value_content)):

        ws.cell(row = i+1, column = 1,value=value_content[i])

    wb.save(path)





# 此程序用于采集athome数据
# result用于存放物件详细信息
result=[]
city_name={'千代田区': '/chintai/tokyo/chiyoda-city/list/', '中央区': '/chintai/tokyo/chuo-city/list/',
           '港区': '/chintai/tokyo/minato-city/list/', '新宿区': '/chintai/tokyo/shinjuku-city/list/',
           '文京区': '/chintai/tokyo/bunkyo-city/list/', '台東区': '/chintai/tokyo/taito-city/list/',
           '墨田区': '/chintai/tokyo/sumida-city/list/', '江東区': '/chintai/tokyo/koto-city/list/',
           '品川区': '/chintai/tokyo/shinagawa-city/list/', '目黒区': '/chintai/tokyo/meguro-city/list/',
           '大田区': '/chintai/tokyo/ota-city/list/', '世田谷区': '/chintai/tokyo/setagaya-city/list/',
           '渋谷区': '/chintai/tokyo/shibuya-city/list/', '中野区': '/chintai/tokyo/nakano-city/list/',
           '杉並区': '/chintai/tokyo/suginami-city/list/', '豊島区': '/chintai/tokyo/toshima-city/list/',
           '北区': '/chintai/tokyo/kita-city/list/', '荒川区': '/chintai/tokyo/arakawa-city/list/',
           '板橋区': '/chintai/tokyo/itabashi-city/list/', '練馬区': '/chintai/tokyo/nerima-city/list/',
           '足立区': '/chintai/tokyo/adachi-city/list/', '葛飾区': '/chintai/tokyo/katsushika-city/list/',
           '江戸川区': '/chintai/tokyo/edogawa-city/list/'}
pages={'千代田区': '18', '中央区': '40', '港区': '75', '新宿区': '107', '文京区': '52', '渋谷区': '88', '台東区': '50', '墨田区': '49', '江東区': '73', '荒川区': '42','足立区': '89', '葛飾区': '70','江戸川区': '137', '品川区': '85', '目黒区': '76', '大田区': '157', '世田谷区': '255', '中野区': '109', '杉並区': '162', '練馬区': '134', '豊島区': '81', '北区': '62', '板橋区': '114'}

#urls用于存放所有陈列页面的链接
page_urls=[]
wu_jian_urls=[]
# 头设置
headers = {
    'Accept': '*/*',
    'Accept-Language': 'en-US,en;q=0.8',
    'Cache-Control': 'max-age=0',
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.116 Safari/537.36',
    'Connection': 'keep-alive',

}
path="https://www.athome.co.jp/chintai/tokyo/edogawa-city/list/"

# 输入一个预览页的链接，得到其中每个物件的信息
def get_link(path):

    controller.signal(Signal.NEWNYM)
    a=requests.get(path,headers=headers)
    html=a.text.encode(a.encoding).decode("utf-8")

    soup=BeautifulSoup(html, features="lxml")
    meta = soup.find_all("meta")
    try:
        if "name" in meta[0].attrs:
            if meta[0]["name"] == "ROBOTS":
                get_link(path)
        else:
            try:
                a=soup.find_all("a")
                for i in range(len(a)):
                    if "詳細を見る" in a[i].text:
                        a_link="https://www.athome.co.jp"+a[i].attrs['href']
                        print(a_link)
                        wu_jian_urls.append(a_link)

            except:
                get_link(path)
    except:
        print(path+"此页出错")

def get_values(path):
    link="https://www.athome.co.jp"+path

    controller.signal(Signal.NEWNYM)
    a = requests.get(link, headers=headers)
    html = a.text.encode(a.encoding).decode("utf-8")
    soup=BeautifulSoup(html, features="lxml")
    meta = soup.find_all("meta")
    if "name" in meta[0].attrs:
        if meta[0]["name"] == "ROBOTS":
            get_values(path)
    else:

        jiao_tong=soup.find(name="th",text="交通").find_next(name="td").get_text()
        if soup.find(name="th",text="その他交通"):
            others_jiao_tong=soup.find(name="th",text="その他交通").find_next(name="td").get_text()
        else:
            others_jiao_tong =""
        address=soup.find(name="th",text="所在地").find_next(name="td").get_text()
        kind=soup.find(name="th",text="物件種目").find_next(name="td").get_text()
        rent=soup.find(name="th",text="賃料").find_next(name="td").get_text()
        guan_li_fei = soup.find(name="th", text="管理費等").find_next(name="td").get_text()
        fu_jin_and_bao_zheng_jin=soup.find(name="th", text="敷金/保証金").find_next(name="td").get_text()
        li_jin=soup.find(name="th", text="礼金").find_next(name="td").get_text()
        wu_jin_ming=soup.find(name="th", text="建物名・部屋番号").find_next(name="td").get_text()
        fang_xing=soup.find(name="th", text="間取り").find_next(name="td").get_text()
        mian_jin=soup.find(name="th", text="専有面積").find_next(name="td").get_text()
        lou_ceng = soup.find(name="th", text="階建 / 階").find_next(name="td").get_text()
        jie_gou=soup.find(name="th", text="建物構造").find_next(name="td").get_text()
        jian_zao_ri=soup.find(name="th", text="築年月").find_next(name="td").get_text()
        out=[jiao_tong,others_jiao_tong,address,kind,rent,guan_li_fei,fu_jin_and_bao_zheng_jin,li_jin,wu_jin_ming,fang_xing,
             mian_jin,lou_ceng,jie_gou,jian_zao_ri]
        return out



'''
for k in city_name:
    link="https://www.athome.co.jp"+city_name[k]
    def get_page_length(link):


        controller.signal(Signal.NEWNYM)

        a = requests.get(link, headers=headers)
        html = a.text.encode(a.encoding).decode("utf-8")

        soup = BeautifulSoup(html, features="lxml")
        meta = soup.find_all("meta")
        if "name" in meta[0].attrs:
            if meta[0]["name"] == "ROBOTS":
                get_page_length(link)
        else:
            print(soup.find("a", onclick="javascript:pushGapCustomForPaging('last');"))
            last_page = soup.find("a", onclick="javascript:pushGapCustomForPaging('last');").attrs['href'].split("page")[1][:-1]
            return int(last_page)


    last_page=get_page_length(link)
    print(last_page)
    print(type(last_page))
    urls.append(link)
    for i in range(2,last_page+1):
        urls.append(link+"page"+str(i)+"/")
'''

for name in city_name:
    link="https://www.athome.co.jp"+city_name[name]
    page_urls.append(link)
    length=int(pages[name])
    for i in range(2,length+1):
        page_urls.append(link + "page" + str(i) + "/")

print(len(page_urls))


class myThread(threading.Thread):
    def __init__(self,name,q):
        threading.Thread.__init__(self)
        self.name=name
        self.q=q
    def run(self):
        print(self.name)
        while not self.q.empty():
            get_link(self.q.get())
workQueue=Queue.Queue(len(page_urls))
threads=[]
for i in range(100):
    thread=myThread("线程"+str(i),workQueue)
    thread.start()
    threads.append(thread)

for url in page_urls:
    workQueue.put(url)

for t in threads:
    t.join()
print(wu_jian_urls)
print(len(wu_jian_urls))
put_cai_wu_value("C:\\Users\\user\\Desktop\\网址.xlsx",wu_jian_urls)
