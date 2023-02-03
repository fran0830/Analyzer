from win32wifi import Win32Wifi as ww
import threading,socket

#########メール送信処理に必要なモジュール群############

import smtplib,os,datetime

from datetime import timedelta

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
####################################################

 ###########リレー先情報の記入################
smtp_server = "smtp.gmail.com"
smtp_port = 587
smtp_user = ""
smtp_passwd = ""
############################################

my = ""       #メールアドレスを指定する。


#周辺のSSIDを取得する関数
def get_ssid(trial=3):
    ##Wlanが変更を登録するEvent
    th_ev = threading.Event()

    def callback(wnd, p):
        th_ev.set()

    interfaces = ww.getWirelessInterfaces()
    handle = ww.WlanOpenHandle()

    ##試行回数
    for i in range(trial):

        ssid_dict = {}

        for interface in interfaces:

            ##検索
            ws = ww.WlanScan(handle, interface.guid)

            #登録完了を最大10秒間待機する。
            cb = ww.WlanRegisterNotification(handle, callback)

            th_ev.wait(10)
            th_ev.clear()
            
            networks = ww.getWirelessNetworkBssList(interface)

            for network in networks:
                ssid = network.ssid.decode()
                bssid = network.bssid
                quality = network.link_quality

                ##ssidがなければリストを追加
                if ssid not in ssid_dict.keys():
                    ssid_dict[ssid] = []

                
                ##bssidと電波強度を追加
                ssid_dict[ssid].append((bssid, quality))

        
        #空でなければBreak
        if ssid_dict !={}:
            break

    ww.WlanCloseHandle(handle)

    return ssid_dict


#メール送信処理
def smail(sfal):
   
    
    #時刻取得
    dt_now = datetime.datetime.now()


    

    to_address = my

    from_address = smtp_user

    #配送先工場


    subject = "未確認SSIDを検知"        #件名

    
    hostname = socket.gethostname()

    #本文
    body = f'''
各位
標題の通り、指定WIFi以外のSSIDを検知しました。

以下、検知されたSSID、MACアドレス、電波強度になります。
※電波強度は数値が大きいほど電波強度が強いです。

検出したPC名:  {hostname}
'''

    #本文に結合する。
    for row in sfal:
        ape = f'SSID: {row["SSID"]},  MACアドレス: {row["MACaddr"]}, 電波強度:  {row["Strength"]}'
        re = '\n'.join([ape])
        body = body+ re + "\n"


    filepath = ""
    filename = os.path.basename(filepath)

    msg = MIMEMultipart()

    msg["Subject"] = subject
    msg["From"] = from_address
    msg["To"] = to_address

    msg.attach(MIMEText(body))

    

   
    s = smtplib.SMTP(smtp_server, smtp_port)
    s.starttls()
    s.login(smtp_user, smtp_passwd)
    sendToList = to_address.split(',')
    s.sendmail(from_address, sendToList, msg.as_string())
        
    s.quit()

    #ここより以下はメールの送信内容を画像化するための処理
    
    return 

def mmail(sfal):
    
    
    #時刻取得
    dt_now = datetime.datetime.now()


    to_address = my

    from_address = smtp_user

    #配送先工場


    subject = "未確認非表示WIfiを検知"        #件名

    
    hostname = socket.gethostname()

    #本文
    body = f'''
各位
標題の通り、指定WIFi以外のSSIDを検知しました。

以下、検知されたSSID、MACアドレス、電波強度になります。
※電波強度は数値が大きいほど電波強度が強いです。

検出したPC名:  {hostname}
'''

    #本文に結合する。
    for row in sfal:
        ape = f'SSID:  非表示の可能性,  MACアドレス: {row[0]}, 電波強度:  {row[1]}'
        re = '\n'.join([ape])
        body = body+ re + "\n"


    filepath = ""
    filename = os.path.basename(filepath)

    msg = MIMEMultipart()

    msg["Subject"] = subject
    msg["From"] = from_address
    msg["To"] = to_address

    msg.attach(MIMEText(body))

    

   
    s = smtplib.SMTP(smtp_server, smtp_port)
    s.starttls()
    s.login(smtp_user, smtp_passwd)
    sendToList = to_address.split(',')
    s.sendmail(from_address, sendToList, msg.as_string())
        
    s.quit()

    #ここより以下はメールの送信内容を画像化するための処理
    
    return 

#稼働確認用の処理
def qmail():

    
    #時刻取得
    dt_now = datetime.datetime.now()


    

     
    to_address = my     #宛先のメールアドレス

    from_address = smtp_user

    #配送先工場


    subject = "7の倍数日稼働確認チェック"        #件名

    
    hostname = socket.gethostname()

    #本文
    body = f'''
稼働確認チェックです。
該当プログラムは稼働中です。

該当PC:  {hostname}
プログラムの機能：
指定したmacアドレス以外のwifi以外のwifiなどが検知されると通報するシステム。
'''

    

    filepath = ""
    filename = os.path.basename(filepath)

    msg = MIMEMultipart()

    msg["Subject"] = subject
    msg["From"] = from_address
    msg["To"] = to_address

    msg.attach(MIMEText(body))

    

   
    s = smtplib.SMTP(smtp_server, smtp_port)
    s.starttls()
    s.login(smtp_user, smtp_passwd)
    sendToList = to_address.split(',')
    s.sendmail(from_address, sendToList, msg.as_string())
        
    s.quit()

    #ここより以下はメールの送信内容を画像化するための処理
    
    return 


if __name__ == "__main__":
    
    company = []
    cheat = []
    cheat_ssid = []
    send_list = []

    #付近のSSID等を取得する。
    ssidlist = get_ssid()
    company = ssidlist["SSID"]  #ここにSSIDを指定
    nonssid = ssidlist[""]      #非表示のSSID

    #取得したSSIDをdict型からlist型へ変換
    ssid_list = list(ssidlist.items())

    #SSIDがそもそも違うやつを取得
    for di in ssid_list:
        if di[0] == "":
            pass

        elif di[0] == "":
            pass
        
        else:
            cheat_ssid.append(di)


    #ここに検知除外するMACアドレスの前の三区切りを入力する。
    safe_list = []

    #会社のwifiに偽装していると思われるやつを洗い出す。
    for co in company:
        if safe_list[0] in co[0]:
            print("含まれる")
        elif safe_list[1] in co[0]:
            print("含まれる")

        else:
            cheat.append(co)
            print("含まれない。")

    #SSID名が取得できないもののwifiを洗い出す。
    for no in nonssid:
        if safe_list[0] in no[0]:
            print("含まれる")

        elif safe_list[1] in no[0]:
            print("含まれる")

        else:
            cheat.append(no)
            print("含まれない。")

    #どちらのリストかに不正なwifiが検知されると以下の処理が動く
    if len(cheat) != 0 or len(cheat_ssid) != 0:
        if len(cheat_ssid) != 0:
            for ch in cheat_ssid:
                SSID = ch[0]
                MACaddr = ch[1][0][0]
                Strength = ch[1][0][1]
                send_list.append({"SSID":SSID, "MACaddr":MACaddr, "Strength":Strength})
            
            #メール送信処理へ
            smail(send_list)

        if len(cheat) != 0:
            mmail(cheat)
            print("")

    #プログラムの稼働確認。
    neko = datetime.datetime.today()
    lp = neko.day
    
    if lp == 7 or lp == 14 or lp == 21 or lp == 28:
        if neko.hour == 8 or neko.hour == 12:
            qmail()

        

    print("")

