#這些是LINE官方開放的套件組合透過import來套用這個檔案上
from linebot import (LineBotApi, WebhookHandler)
from linebot.exceptions import (InvalidSignatureError)
from linebot.models import *

menu_text='''以下是可搜尋菜單的飲料店：
M01.可不可熟成紅茶
M02.50嵐
M03.清心福全
M04.五桐號
M05.迷客夏
M06.麻古
M07.龜記
M08.Comebuy
M09.茶湯會
M10.CoCo
M11.大苑子
M12.老賴茶棧
M13.珍煮丹
M14.樺達奶茶
M15.鶴茶慺
請輸入飲料店代碼來搜尋！'''
store_name = ["50嵐","龜記","可不可熟成紅茶","老賴茶棧","樺達奶茶","CoCo","麻古","清心福全","鶴茶樓","五桐號","迷客夏","茶湯會","Comebuy","大苑子","珍煮丹"]
drinks_50嵐=["四季春加珍波椰", "冰淇淋紅茶", "烏龍瑪奇朵", "波霸奶茶", "8冰綠"]
drinks_龜記=["紅釉翡翠", "龜記濃乳茶", "紅烏鮮乳", "阿源楊桃紅", "三韻紅萱"]
drinks_茶湯會=["觀音拿鐵", "翡翠檸檬", "蔗香紅茶", "珍珠奶茶", "普洱小珍珠拿鐵"]
drinks_可不可熟成紅茶=["春蘋紅茶", "胭脂多多", "胭脂紅茶", "白玉歐蕾", "熟成紅茶"]
drinks_老賴茶棧=["老賴紅茶", "豆香紅茶", "招牌奶茶", "(珍珠)太后牛乳", "冬瓜檸檬"]
drinks_樺達奶茶=["樺達奶茶", "益壽奶茶", "美容奶茶", "紅龍奶茶", "烏梅普洱"]
drinks_CoCo=["百香雙響炮", "奶茶三兄弟", "珍珠奶茶", "蜜香檸凍紅茶", "紅果小姐"]
drinks_麻古=["楊枝甘露2.0", "芝芝(當季水果)系列", "金萱雙(三)Q", "梅子冰茶", "柳橙果粒茶"]
drinks_清心福全=["烏龍綠茶", "梅子綠茶", "冬瓜菁茶", "隱藏版", "蘆薈優多綠茶"]
drinks_鶴茶樓=["鶴頂紅茶", "綺夢那提+鶴頂紅茶凍", "藝伎紅茶", "鶴頂燕麥奶凍", "神濃氏那堤"]
drinks_五桐號=["杏仁凍五桐茶", "清香奶霜烏龍+小芋圓", "雪絨草莓奶酪", "綠茶凍五桐奶茶", "招牌五桐奶霜"]
drinks_迷客夏=["娜杯紅茶拿鐵", " 珍珠紅茶拿鐵", "青檸香茶", "冰萃柳丁", "手炒黑糖鮮奶"]
drinks_珍煮丹=["黑糖珍珠鮮奶", "泰泰鮮奶茶", "覓蜜芋圓鮮奶", "十份芋芋鮮奶", "姍姍紅茶拿鐵"]
drinks_大苑子=["台灣鮮搾柳橙綠", "莓好時光", "芭樂檸檬", "莓好花漾", "番茄梅"]
drinks_Comebuy=["絕代雙Q奶茶", "蘋果冰茶", "海神/ 海神奶茶", "芭樂檸檬綠", "百香搖果樂"]
store_name = ["50嵐","龜記","可不可熟成紅茶","老賴茶棧","樺達奶茶","CoCo","麻古","清心福全","鶴茶樓","五桐號","迷客夏","茶湯會","Comebuy","大苑子","珍煮丹"]

# ImagemapSendMessage(組圖訊息)
def imagemap_message():
    message2 = ImagemapSendMessage(
        base_url="https://i.imgur.com/KCXVHF7.jpg",
        alt_text='官網在這裡！',
        base_size=BaseSize(height=2600, width=1500),
        actions=[
            URIImagemapAction(
                #可不可熟成紅茶
                link_uri="https://www.kebuke.com/",
                area=ImagemapArea(
                    x=0, y=0, width=500, height=500
                )
            ),
            URIImagemapAction(
                #50嵐
                link_uri="http://50lan.com/web/news.asp",
                area=ImagemapArea(
                    x=0, y=500, width=500, height=500
                )
            ),
            URIImagemapAction(
                #清心福全
                link_uri="https://www.chingshin.tw/",
                area=ImagemapArea(
                    x=0, y=1000, width=500, height=500
                )
            ),
            URIImagemapAction(
                #五桐號
                link_uri="https://www.wootea.com/",
                area=ImagemapArea(
                    x=0, y=1500, width=500, height=500
                )
            ),
            URIImagemapAction(
                #迷客夏
                link_uri="http://www.milkshoptea.com/index.php",
                area=ImagemapArea(
                    x=0, y=2000, width=500, height=500
                )
            ),
            URIImagemapAction(
                #麻古
                link_uri="https://macutea.com.tw/",
                area=ImagemapArea(
                    x=500, y=0, width=500, height=500
                )
            ),
            URIImagemapAction(
                #龜記
                link_uri="https://guiji-group.com/",
                area=ImagemapArea(
                    x=500, y=500, width=500, height=500
                )
            ),
            URIImagemapAction(
                #Comebuy
                link_uri="https://www.comebuy2002.com.tw/",
                area=ImagemapArea(
                    x=500, y=1000, width=500, height=500
                )
            ),
            URIImagemapAction(
                #茶湯會
                link_uri="https://tw.tp-tea.com/",
                area=ImagemapArea(
                    x=500, y=1500, width=500, height=500
                )
            ),
            URIImagemapAction(
                #CoCo
                link_uri="https://www.coco-tea.com/",
                area=ImagemapArea(
                    x=500, y=2000, width=500, height=500
                )
            ),
            URIImagemapAction(
                #大苑子
                link_uri="https://www.dayungs.com/",
                area=ImagemapArea(
                    x=1000, y=0, width=500, height=500
                )
            ),
            URIImagemapAction(
                #老賴茶棧
                link_uri="https://www.liketeashop.com/",
                area=ImagemapArea(
                    x=1000, y=500, width=500, height=500
                )
            ),
            URIImagemapAction(
                #珍煮丹
                link_uri="https://www.truedan.com.tw/",
                area=ImagemapArea(
                    x=1000, y=1000, width=500, height=500
                )
            ),
            URIImagemapAction(
                #樺達奶茶
                link_uri="https://www.facebook.com/HWADAmilktea/?locale=zh_TW",
                area=ImagemapArea(
                    x=1000, y=1500, width=500, height=500
                )
            ),
            URIImagemapAction(
                #鶴茶慺
                link_uri="https://hechaloutea.com.tw/",
                area=ImagemapArea(
                    x=1000, y=2000, width=500, height=500
                )
            ),
           
        ]
    )
    return message2


def Carousel_Template():
    message3 = TemplateSendMessage(
        alt_text='查詢捷運站附近的飲料店',
        template=CarouselTemplate(
            columns=[
                CarouselColumn(
                    thumbnail_image_url='https://upload.wikimedia.org/wikipedia/commons/thumb/7/71/Taipei_Metro_Line_BR.svg/330px-Taipei_Metro_Line_BR.svg.png',
                    title='文湖線',
                    text='請點選下方連結查找站名',
                    actions=[
                        MessageTemplateAction(
                            label='文湖線站名',
                            text='文湖線'
                        ),
                    ]
                ),
                CarouselColumn(
                    thumbnail_image_url='https://upload.wikimedia.org/wikipedia/commons/thumb/2/21/Taipei_Metro_Line_BL.svg/1200px-Taipei_Metro_Line_BL.svg.png',
                    title='板南線',
                    text='請點選下方連結查找站名',
                    actions=[
                        MessageTemplateAction(
                            label='板南線站名',
                            text='板南線'
                        ),
                    ]
                ),
                CarouselColumn(
                    thumbnail_image_url='https://upload.wikimedia.org/wikipedia/commons/thumb/f/f3/Taipei_Metro_Line_R.svg/800px-Taipei_Metro_Line_R.svg.png',
                    title='淡水信義線',
                    text='請點選下方連結查找站名',
                    actions=[
                        MessageTemplateAction(
                            label='淡水信義線站名',
                            text='淡水信義線'
                        ),
                    ]
                ),
                CarouselColumn(
                    thumbnail_image_url='https://upload.wikimedia.org/wikipedia/commons/thumb/1/1f/Taipei_Metro_Line_G.svg/330px-Taipei_Metro_Line_G.svg.png',
                    title='松山新店線',
                    text='請點選下方連結查找站名',
                    actions=[
                        MessageTemplateAction(
                            label='松山新店線站名',
                            text='松山新店線'
                        ),
                    ]
                ),
                CarouselColumn(
                    thumbnail_image_url='https://upload.wikimedia.org/wikipedia/commons/thumb/e/eb/Taipei_Metro_Line_O.svg/1200px-Taipei_Metro_Line_O.svg.png',
                    title='中和新蘆線',
                    text='請點選下方連結查找站名',
                    actions=[
                        MessageTemplateAction(
                            label='中和新蘆線站名',
                            text='中和新蘆線'
                        ),
                    ]
                ),
            ]
        )
    )
    return message3
