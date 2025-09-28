from flask import Flask, request, abort
from linebot import (LineBotApi, WebhookHandler)
from linebot.exceptions import (InvalidSignatureError)
from linebot.models import *
from MenuAndRandom import *
from StoreInfo import *

import random

app = Flask(__name__)

line_bot_api = LineBotApi('fxP76FN1ngpTlzEsbDQjAl+TbmtTKMwamySyV8Esz/hgivTzflPouA9wN6Dn0NdkMUiJJWtYHQmJ+3ahCbFuAT3N2BtO4VExSfvvAx2fEJvKuV25b6QCLIRQH7Yu+1XSGOWLkcBXIXFM7TZ1gfSF5AdB04t89/1O/w1cDnyilFU=')
handler = WebhookHandler('2c6f48674a3f81d049e67c92fe64ae79')

# 監聽所有來自 /callback 的 Post Request
@app.route("/callback", methods=['POST'])

def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']
    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body:" + body)
    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    usermessage=event.message.text
    if usermessage=='test': #測試連結是否成功
        message = TextSendMessage(text=usermessage)
        line_bot_api.reply_message(event.reply_token, message)

    elif usermessage=='關於我們-製作者名單':
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/WGNvW7p.png",
                preview_image_url = "https://i.imgur.com/WGNvW7p.png"),
            TextSendMessage(  #傳送文字
                    text = "以上為製作者的介紹！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == '查詢飲料店菜單':
        message = TextSendMessage(text=menu_text)
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M01': #可不可熟成紅茶的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/b24BNdR.jpg",
                preview_image_url = "https://i.imgur.com/b24BNdR.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為可不可熟成紅茶的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M02': #50嵐的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://twcoupon.com/images/menu/p_50lan.jpg",
                preview_image_url = "https://twcoupon.com/images/menu/p_50lan.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為50嵐的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M03': #清心福全的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/glpwkPj.jpg",
                preview_image_url = "https://i.imgur.com/glpwkPj.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為清心福全的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M04': #五桐號的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/25gxMpi.jpg",
                preview_image_url = "https://i.imgur.com/25gxMpi.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為五桐號的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M05': #迷客夏的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://1.bp.blogspot.com/-_QaH_2DfHYc/YK5K4yztWGI/AAAAAAAAaFk/c9QdimW_A1kxWpNcZTpHrkgZM56yveyywCLcBGAsYHQ/s2048/%25E3%2580%2590%25E8%25BF%25B7%25E5%25AE%25A2%25E5%25A4%258F%25E3%2580%25912021%25E8%258F%259C%25E5%2596%25AE%25E5%2583%25B9%25E7%259B%25AE%25E8%25A1%25A8.jpg",
                preview_image_url = "https://1.bp.blogspot.com/-_QaH_2DfHYc/YK5K4yztWGI/AAAAAAAAaFk/c9QdimW_A1kxWpNcZTpHrkgZM56yveyywCLcBGAsYHQ/s2048/%25E3%2580%2590%25E8%25BF%25B7%25E5%25AE%25A2%25E5%25A4%258F%25E3%2580%25912021%25E8%258F%259C%25E5%2596%25AE%25E5%2583%25B9%25E7%259B%25AE%25E8%25A1%25A8.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為迷客夏的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M06': #麻古的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/GjsteZG.jpg",
                preview_image_url = "https://i.imgur.com/GjsteZG.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為麻古的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M07': #龜記的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/fwMXnWH.jpg",
                preview_image_url = "https://i.imgur.com/fwMXnWH.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為龜記的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M08': #Comebuy的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/uTlujx6.jpg",
                preview_image_url = "https://i.imgur.com/uTlujx6.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為Comebuy的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M09': #茶湯會的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/vKO06io.jpg",
                preview_image_url = "https://i.imgur.com/vKO06io.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為茶湯會的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M10': #CoCo的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/T3wxr3g.jpg",
                preview_image_url = "https://i.imgur.com/T3wxr3g.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為CoCo的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M11': #大苑子的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/uOb8rct.jpg",
                preview_image_url = "https://i.imgur.com/uOb8rct.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為大苑子的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M12': #老賴茶棧的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/XlFg2AQ.jpg",
                preview_image_url = "https://i.imgur.com/XlFg2AQ.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為老賴茶棧的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M13': #珍煮丹的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/aSzyl1t.jpg",
                preview_image_url = "https://i.imgur.com/aSzyl1t.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為珍煮丹的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M14': #樺達奶茶的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/Xdqlixd.jpg",
                preview_image_url = "https://i.imgur.com/Xdqlixd.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為樺達奶茶的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    elif usermessage == 'M15': #鶴茶慺的菜單
        message = [
            ImageSendMessage(
                original_content_url = "https://i.imgur.com/kLm30XX.jpg",
                preview_image_url = "https://i.imgur.com/kLm30XX.jpg"),
            TextSendMessage(  #傳送文字
                    text = "以上為鶴茶慺的菜單！"
                )
            ]
        line_bot_api.reply_message(event.reply_token,message)
    
    elif usermessage == '查詢飲料店官網':
        message2 = imagemap_message()
        message1=TextSendMessage("點擊飲料店LOGO進入官網！")
        line_bot_api.reply_message(event.reply_token, message2)
        line_bot_api.reply_message(event.reply_token, message1)

    elif '查詢捷運站附近的飲料店' in usermessage:
        message3 = Carousel_Template()
        line_bot_api.reply_message(event.reply_token, message3)
    elif '文湖線' in usermessage:
        message = '以下是「文湖線」的車站：\n1南港展覽館站\n2南港軟體園區站\n3東湖站\n4葫洲站\n5大湖公園站\n6內湖站\n7文德站\n8港墘站\n9西湖站\n10劍南路\n11大直站\n12松山機場站\n13中山國中站\n14南京復興站\n15忠孝復興站\n16大安站\n17科技大樓站\n18六張犁站\n19麟光站\n20辛亥站\n21萬芳醫院站\n22萬芳社區站\n23木柵站\n24動物園站\n請輸入捷運站「數字」代碼來搜尋！'
        message1 = TextSendMessage(text=message)
        line_bot_api.reply_message(event.reply_token, message1)
    elif '板南線' in usermessage:
        message = '以下是「板南線」的車站：\n25頂埔站\n26永寧站\n27土城站\n28海山站\n29亞東醫院站\n30府中站\n31板橋站\n32新埔站\n33江子翠站\n34龍山寺站\n35西門站\n36台北車站\n37善導寺站\n38忠孝新生站\n39忠孝復興站\n40忠孝敦化站\n41國父紀念館站\n42市政府站\n43永春站\n44後山埤站\n45昆陽站\n46南港站\n47南港展覽館站\n請輸入捷運站「數字」代碼來搜尋！'
        message1 = TextSendMessage(text=message)
        line_bot_api.reply_message(event.reply_token, message1)
    elif '淡水信義線' in usermessage:
        message = '以下是「淡水信義線」的車站：\n48象山站\n49台北101/世貿站\n50信義安和站\n51大安森林公園站\n52東門站\n53中正紀念堂站\n54台大醫院站\n55台北車站\n56中山站\n57雙連站\n58民權西路站\n59圓山站\n60劍潭站\n61士林站\n62芝山站\n63明德站\n64石牌站\n65唭哩岸站\n66奇岩站\n67北投站\n68新北投站\n69復興崗站\n70忠義站\n71關渡站\n72竹圍站\n73紅樹林站\n74淡水站\n請輸入捷運站「數字」代碼來搜尋！'
        message1 = TextSendMessage(text=message)
        line_bot_api.reply_message(event.reply_token, message1)
    elif '松山新店線' in usermessage:
        message = '以下是「松山新店線」的車站：\n75新店站\n76新店區公所站\n77七張站\n78小碧潭站\n79大坪林站\n80景美站\n81萬隆站\n82公館站\n83台電大樓站\n84古亭站\n85中正紀念堂站\n86小南門站\n87西門站\n88北門站\n89中山站\n90松江南京站\n91南京復興站\n92台北小巨蛋站\n93南京三民站\n94松山站\n請輸入捷運站「數字」代碼來搜尋！'
        message1 = TextSendMessage(text=message)
        line_bot_api.reply_message(event.reply_token, message1)
    elif '中和新蘆線' in usermessage:
        message = '以下是「中和新蘆線」的車站：\n95南勢角站\n96景安站\n97永安市場站\n98頂溪站\n99古亭站\n100東門站\n101忠孝新生站\n102松江南京站\n103行天宮站\n104中山國小站\n105民權西路站\n106大橋頭站\n107台北橋站\n108菜寮站\n109三重站\n110先嗇宮站\n111頭前庄站\n112新莊站\n113輔大站\n114丹鳳站\n115迴龍站\n116三重國小站\n117三和國中站\n118徐匯中學站\n119三民高中站\n120蘆洲站\n請輸入捷運站「數字」代碼來搜尋！'
        message1 = TextSendMessage(text=message)
        line_bot_api.reply_message(event.reply_token, message1)

    elif usermessage == '隨機推薦飲料店':
        computer_store = random.choice(store_name)
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text=computer_store))

    elif usermessage =='隨機推薦飲料':
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text="飲料店名單如下：\n50嵐\n清心福全\n五桐號\n迷客夏\n麻古\n龜記\nComebuy\n茶湯會\nCoCo\n大苑子\n老賴茶棧\n珍煮丹\n樺達奶茶\n鶴茶慺\n請問要推薦哪間飲料店的飲料呢?"))
    elif usermessage == '麻古':
        computer_drinks = random.choice(drinks_麻古)
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text="為你推薦一杯: " + computer_drinks))
    elif usermessage == "50嵐":
        computer_drinks = random.choice(drinks_50嵐)
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text="為你推薦一杯: " + computer_drinks))
    elif usermessage == "CoCo":
            computer_drinks2 = random.choice(drinks_CoCo)
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="為你推薦一杯: " + computer_drinks2)
            )
    elif usermessage== "Comebuy":
        computer_drinks3 = random.choice(drinks_Comebuy)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks3)
        )
    elif usermessage== "五桐號":
        computer_drinks4 = random.choice(drinks_五桐號)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks4)
        )
    elif usermessage== "可不可熟成紅茶":
        computer_drinks5 = random.choice(drinks_可不可熟成紅茶)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks5)
        )
    elif usermessage== "大苑子":
        computer_drinks6 = random.choice(drinks_大苑子)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks6)
        )
    elif usermessage== "樺達奶茶":
        computer_drinks7 = random.choice(drinks_樺達奶茶)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks7)
        )
    elif usermessage== "清心福全":
        computer_drinks8 = random.choice(drinks_清心福全)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks8)
        )
    elif usermessage== "珍煮丹":
        computer_drinks9 = random.choice(drinks_珍煮丹)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks9)
        )
    elif usermessage== "老賴茶棧":
        computer_drinks10 = random.choice(drinks_老賴茶棧)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks10)
        )
    elif usermessage== "茶湯會":
        computer_drinks11 = random.choice(drinks_茶湯會)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks11)
        )
    elif usermessage== "麻古":
        computer_drinks12 = random.choice(drinks_麻古)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks12)
        )
    elif usermessage== "龜記":
        computer_drinks13 = random.choice(drinks_龜記)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks13)
        )
    elif usermessage== "迷客夏":
        computer_drinks14 = random.choice(drinks_迷客夏)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks14)
        )
    elif usermessage== "鶴茶樓":
        computer_drinks15 = random.choice(drinks_鶴茶樓)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="為你推薦一杯: " + computer_drinks15)
        )
    # elif usermessage 1:
    #     station_num = int(message1)
    #     message=store_info()
    #     line_bot_api.reply_message(event.reply_token, message)

    elif usermessage.isdigit()==True: #輸入其他內容
        station_num = int(usermessage) - 1
        timee=0
        a = search(station_num)
        Carousel_Columns = []
        for i in range(len(a)):
            if timee==10:
                break
            else:
                timee+=1
                column = CarouselColumn(
                    thumbnail_image_url=a.iloc[i]["連鎖店照片"],
                    title=a.iloc[i]["店名"],
                    text="電話："+a.iloc[i]["電話"]+"\n地址："+a.iloc[i]["地址"],
                    actions=[
                        MessageTemplateAction(
                            label="我想知道這店家在哪裡！",
                            text=a.iloc[i]["店名"]
                        ),MessageTemplateAction(
                            label="我想看MENU！",
                            text=a.iloc[i]["菜單代碼"]
                        ),
                    ]
                )
                Carousel_Columns.append(column)
        message4 = [TextSendMessage(text=station[station_num]+"附近的連鎖飲料店在這裡！"),
                    TemplateSendMessage(alt_text=station[station_num]+'附近的連鎖飲料店在這裡！',template=CarouselTemplate(columns = Carousel_Columns)),]
        line_bot_api.reply_message(event.reply_token,message4)
    else:
        try:
            result=get_latitude_longtitude(usermessage)
            lat=getlat(result)
            lng=getlng(result)
            message=LocationSendMessage(
                title= usermessage,
                address= usermessage,
                latitude= lat ,
                longitude= lng ,)
            line_bot_api.reply_message(event.reply_token,message)
        except:
            message=[
                TextSendMessage(text='不好意思！現在還沒有支援搜尋"'+usermessage+'"的功能噢！'),
                StickerSendMessage(package_id = '8522', sticker_id = '16581287'),]
            line_bot_api.reply_message(event.reply_token,message)
    


import os  #伺服器
if __name__ == "__main__":
    port = int(os.environ.get('PORT',5000))
    app.run(host='0.0.0.0',port=port)