from linebot import (LineBotApi, WebhookHandler)
from linebot.exceptions import (InvalidSignatureError)
from linebot.models import *
from ChannelTry import *
import googlemaps
import time #python內建的時間模組後面設定間隔時間會用到
import pandas as pd
import re
import requests
import openpyxl
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import json
import urllib

gmaps = googlemaps.Client(key = "您自己的")
GOOGLE_API_KEY = '您自己的'

station = ["南港展覽館站", "南港軟體園區站", "東湖站", "葫洲站", "大湖公園站", "內湖站", "文德站", "港墘站",
 "西湖站", "劍南路", "大直站", "松山機場站", "中山國中站", "南京復興站", "忠孝復興站", "大安站", "科技大樓站", "六張犁站",
 "麟光站", "辛亥站", "萬芳醫院站", "萬芳社區站", "木柵站", "動物園站",
 "頂埔站", "永寧站", "土城站", "海山站", "亞東醫院站", "府中站", "板橋站", "新埔站", "江子翠站", "龍山寺站", "西門站",
 "台北車站", "善導寺站", "忠孝新生站", "忠孝復興站", "忠孝敦化站", "國父紀念館站", "市政府站", "永春站", "後山埤站", 
 "昆陽站", "南港站", "南港展覽館站","象山站", "台北101/世貿站", "信義安和站", "大安森林公園站", "東門站", "中正紀念堂站", "台大醫院站", "台北車站", "中山站",
 "雙連站", "民權西路站", "圓山站", "劍潭站", "士林站", "芝山站", "明德站", "石牌站", "唭哩岸站", "奇岩站", "北投站", 
 "新北投站", "復興崗站", "忠義站", "關渡站", "竹圍站", "紅樹林站", "淡水站",
 "新店站", "新店區公所站", "七張站", "小碧潭站", "大坪林站", "景美站", "萬隆站", "公館站", "台電大樓站", "古亭站", 
 "中正紀念堂站", "小南門站", "西門站", "北門站", "中山站", "松江南京站", "南京復興站", "台北小巨蛋站", "南京三民站", "松山站",
 "南勢角站", "景安站", "永安市場站", "頂溪站", "古亭站", "東門站", "忠孝新生站", "松江南京站", "行天宮站", "中山國小站", 
 "民權西路站", "大橋頭站", "台北橋站", "菜寮站", "三重站", "先嗇宮站", "頭前庄站", "新莊站", "輔大站", "丹鳳站", "迴龍站",
 "三重國小站", "三和國中站", "徐匯中學站", "三民高中站", "蘆洲站"]

def drinks(x):
        if '老賴茶棧' in x:i = '老賴茶棧'
        elif '可不可熟成紅茶' in x:i = '可不可熟成紅茶'
        elif '50嵐' in x:i = '50嵐'
        elif '清心福全' in x:i = '清心福全'
        elif '五桐號' in x:i = '五桐號'
        elif '迷客夏' in x:i = '迷客夏'
        elif '麻古' in x:i = '麻古'
        elif '龜記' in x:i = '龜記'
        elif 'Combuy' in x:i = 'Combuy'
        elif '茶湯會' in x:i = '茶湯會'
        elif 'CoCo' in x:i = 'CoCo'
        elif '大苑子' in x:i = '大苑子'
        elif '珍煮丹' in x:i = '珍煮丹'
        elif '樺達奶茶' in x:i = '樺達奶茶'
        elif '鶴茶樓' in x:i = '鶴茶樓'
        else:i = '其他'
        return i

def shoppic(x):
    if '老賴茶棧' in x:i = 'https://shijuecanyin.com/editpic/image/20200102/20200102094571187118.jpg'
    elif '可不可熟成紅茶' in x:i = 'https://www.kebuke.com/wp-content/uploads/2020/12/fb-banner.png'
    elif '50嵐' in x:i = 'https://static.iyp.tw/5910/products/photooriginal-480051-VxXeN.png'
    elif '清心福全' in x:i = 'https://payload.cargocollective.com/1/9/306969/6096935/02-02.jpg'
    elif '五桐號' in x:i = 'https://ap-south-1.linodeobjects.com/nidin-production/store/icons/b_1212_icon_20220810_095745_dd2c4.png'
    elif '迷客夏' in x:i = 'https://www.febigcity.com/bigcity/ImgUpload/brand/Untitled-1_1666344418.jpg'
    elif '麻古' in x:i = 'https://i.imgur.com/m371msX.png'
    elif '龜記' in x:i = 'https://i.imgur.com/E39J91p.png'
    elif 'Combuy' in x:i = 'https://i.imgur.com/XoFR4X5.png'
    elif '茶湯會' in x:i = 'https://shoplineimg.com/5fa3d2b246e9ed0029f5f6bf/603720ae5e5c3b001d6aa912/400x.jpg?'
    elif 'CoCo' in x:i = 'https://img.sj33.cn/uploads/202104/7-21041021351A14.jpg'
    elif '大苑子' in x:i = 'https://www.lohasglasses.com/photo/image/vvip/55/202077540.jpg'
    elif '珍煮丹' in x:i = 'https://refine.tw/wp-content/uploads/zhen-zhu-dan-logo-original-expand-1.jpg'
    elif '樺達奶茶' in x:i = 'https://www.skmpark.com/WebFiles/20220114/20220114_d535461b22914bb397488578bdf6a6c2.jpg'
    elif '鶴茶樓' in x:i = 'https://hechaloutea.com.tw/wp-content/uploads/2020/03/-e1583732113182.png'
    else:i = '其他'
    return i

def shopnum(x):
    if '老賴茶棧' in x:i = 'M12'
    elif '可不可熟成紅茶' in x:i = 'M01'
    elif '50嵐' in x:i = 'M02'
    elif '清心福全' in x:i = 'M03'
    elif '五桐號' in x:i = 'M04'
    elif '迷客夏' in x:i = 'M05'
    elif '麻古' in x:i = 'M06'
    elif '龜記' in x:i = 'M07'
    elif 'Combuy' in x:i = 'M08'
    elif '茶湯會' in x:i = 'M09'
    elif 'CoCo' in x:i = 'M10'
    elif '大苑子' in x:i = 'M11'
    elif '珍煮丹' in x:i = 'M13'
    elif '樺達奶茶' in x:i = 'M14'
    elif '鶴茶樓' in x:i = 'M15'
    else:i = '其他'
    return i

def search(x):
    ids = [] 
    stores_info = []
    results = []
    geocode_result = gmaps.geocode(station[x])
    loc = geocode_result[0]['geometry']['location']
    query_result = gmaps.places_nearby(keyword = "飲料店",location = loc, radius = 500)
    results.extend(query_result['results'])

    while query_result.get('next_page_token'):
        time.sleep(2)
        query_result = gmaps.places_nearby(page_token = query_result['next_page_token'])
        results.extend(query_result['results'])

    for place in results:
        ids.append(place['place_id'])
        
    # 去除重複id
    ids = list(set(ids))
    for id in ids:
        stores_info.append(gmaps.place(place_id = id, language = 'zh-TW')['result'])
        
    pids = []
    for place in gmaps.places_nearby(keyword = "飲料店", location = loc, radius = 500)['results']:
        pids.append(place['place_id'])  #只取出result裡面place_id的部分加到list裏頭
        
    for id in pids:
    #     print ("running")
        stores_info.append(gmaps.place(place_id=id, language='zh-TW')['result'])
        #每次間隔0.3sec
        time.sleep(0.3)
        
    output1 = pd.DataFrame.from_dict(stores_info)
    output2 = output1.rename(columns = {'name': '店名', 'user_ratings_total': '評論數', 'formatted_address': '地址', 'formatted_phone_number': '電話', 'website': '網站'})
    cn_storeinfo = output2[['店名', '地址', '電話', '網站', '評論數']]

    # output2.index = output2.index + 1
    output2 = output2[['店名', '地址', '電話']]
    output2['連鎖店店名'] = output2['店名'].apply(lambda x:drinks(x))
    output2['連鎖店照片'] = output2['店名'].apply(lambda x:shoppic(x))
    output2['菜單代碼'] = output2['店名'].apply(lambda x:shopnum(x))
    # output3 = output2[output2['店名']]

    output99 = output2[output2['連鎖店店名']!='其他']
    # output99.to_csv("final.csv")
    # dff = pd.read_csv('final.csv', index_col = 0)
    # column_names = ['店名', '地址', '電話', '連鎖店店名']
    output99 = output99.drop_duplicates()
    output99.index = range(len(output99))

    return output99

def get_latitude_longtitude(address):
    # decode url
    address = urllib.request.quote(address)
    url = "https://maps.googleapis.com/maps/api/geocode/json?address=" + address + "&key=" + GOOGLE_API_KEY
    
    while True:
        res = requests.get(url)
        js = json.loads(res.text)

        if js["status"] != "OVER_QUERY_LIMIT":
            time.sleep(1)
            break

    result = js["results"][0]["geometry"]["location"]
    return result

def getlat(result):
    lat = result["lat"]
    return lat
def getlng(result):
    lng = result["lng"]
    return lng
#     print(lat,lng)
    # return lat, lng
# ask = input("請輸入想知道經緯度的地址：")
# get_latitude_longtitude(ask)
