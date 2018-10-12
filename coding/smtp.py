#-*- encoding: utf-8 -*-
'''
函数说明：Send_email_text() 函数实现发送带有附件的邮件，可以群发，附件格式包括：xlsx,pdf,txt,jpg,mp3等
参数说明：
    1. subject：邮件主题
    2. content：邮件正文
    3. filepath：附件的地址, 输入格式为["","",...]
    4. receive_email：收件人地址, 输入格式为["","",...]
'''
from email.header import Header
from email import encoders
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import datetime
import time


def Send_email_text(subject, content, filepath, receive_email):
    # sender = "digitalplanet@163.com"
    sender = "haxidata@163.com"
    passwd = "kong368534"
    receivers = receive_email  # 收件人邮箱

    msgRoot = MIMEMultipart()

    msgRoot['From'] = sender
    msgRoot['To'] = receivers

    # if len(receivers) > 1:
    # 	# msgRoot['To'] = receivers[0]  # 群发邮件
    # 	msgRoot['To'] = receivers[0]
    # 	msgRoot['CC'] = receivers[1]
    # 	# msgRoot['Bcc'] = ','.join(receivers[2:])
    # else:
    # 	msgRoot['To'] = receivers[0]

    #     添加图片
    fp = open('../data/img/有你Offer.jpg', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgImage.add_header('Content-ID', '<image1>')
    msgRoot.attach(msgImage)

    msgRoot['Subject'] = subject
    part = MIMEText(content)
    msgRoot.attach(part)

    # html格式构造
    html = """
            <html> 
              <head></head> 
              <body> 
                  <p>
                    %s
                   <br><img src="cid:image1"></br> 
                  </p>
              </body> 
            </html> 
            """ % (content)
    htm = MIMEText(html, 'html', 'utf-8')
    msgRoot.attach(htm)

    ##添加附件部分
    for path in filepath:
        if ".jpg" in path or ".png" in path:
            # jpg类型附件
            jpg_name = path.split("/")[-1]
            part = MIMEApplication(open(path, 'rb').read())
            part.add_header('Content-Disposition', 'attachment', filename=('gbk', '',jpg_name))
            msgRoot.attach(part)

        if ".pdf" in path:
            # pdf类型附件
            pdf_name = path.split("/")[-1]
            part = MIMEApplication(open(path, 'rb').read())
            part.add_header('Content-Disposition', 'attachment', filename=pdf_name)
            msgRoot.attach(part)

        if ".xlsx" in path:
            # xlsx类型附件
            xlsx_name = path.split("/")[-1]
            part = MIMEApplication(open(path, 'rb').read())
            # part.add_header('Content-Disposition', 'attachment', filename=xlsx_name)
            part.add_header('Content-Disposition', 'attachment', filename=('gbk', '', xlsx_name))
            # encoders.encode_base64(part)
            msgRoot.attach(part)

        if ".txt" in path:
            # txt类型附件
            txt_name = path.split("/")[-1]
            part = MIMEApplication(open(path, 'rb').read())
            # part.add_header('Content-Disposition', 'attachment', filename=txt_name)
            part.add_header('Content-Disposition', 'attachment', filename=('gbk', '', txt_name))
            encoders.encode_base64(part)
            msgRoot.attach(part)


        if ".mp3" in path:
            # mp3类型附件
            mp3_name = path.split("/")[-1]
            part = MIMEApplication(open(path, 'rb').read())
            part.add_header('Content-Disposition', 'attachment', filename=mp3_name)
            msgRoot.attach(part)

    try:
        # s = smtplib.SMTP()
        # s.set_debuglevel(1)
        # print(msgRoot.as_string())
        s = smtplib.SMTP()
        s.connect("smtp.163.com")  # 这里我使用的是阿里云邮箱,也可以使用163邮箱：smtp.163.com
        s.login(sender, passwd)
        s.sendmail(sender, receivers, msgRoot.as_string())
        print("邮件发送成功")
    except smtplib.SMTPException as e:
        print(e)
        print("Error, 发送失败")
    finally:
        s.quit()


def email_text(subject, content,  filepath , receive_email):
    # sender = "digitalplanet@163.com"
    sender = "haxidata@163.com"
    passwd = "kong368534"
    receivers = receive_email  # 收件人邮箱

    msgRoot = MIMEMultipart()

    msgRoot['From'] = sender
    msgRoot['To'] = receivers

    #     添加图片
    fp = open('../data/img/有你Offer.jpg', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgImage.add_header('Content-ID', '<image1>')
    msgRoot.attach(msgImage)

    msgRoot['Subject'] = subject
    part = MIMEText(content)
    msgRoot.attach(part)

    # html格式构造
    html = """
            <html> 
              <head></head> 
              <body> 
                  <p>
                    %s
                   <br><img src="cid:image1"></br> 
                  </p>
              </body> 
            </html> 
            """ % (content)
    htm = MIMEText(html, 'html', 'utf-8')
    msgRoot.attach(htm)

    ##添加附件部分
    for path in filepath:
        if ".jpg" in path or ".png" in path:
            # jpg类型附件
            jpg_name = path.split("/")[-1]
            part = MIMEApplication(open(path, 'rb').read())
            part.add_header('Content-Disposition', 'attachment', filename=('gbk', '',jpg_name))
            msgRoot.attach(part)

        if ".pdf" in path:
            # pdf类型附件
            pdf_name = path.split("/")[-1]
            part = MIMEApplication(open(path, 'rb').read())
            part.add_header('Content-Disposition', 'attachment', filename=pdf_name)
            msgRoot.attach(part)

        # if ".xlsx" in path:
        #     # xlsx类型附件
        #     xlsx_name = path.split("/")[-1]
        #     part = MIMEApplication(open(path, 'rb').read())
        #     # part.add_header('Content-Disposition', 'attachment', filename=xlsx_name)
        #     part.add_header('Content-Disposition', 'attachment', filename=('gbk', '', xlsx_name))
        #     # encoders.encode_base64(part)
        #     msgRoot.attach(part)

        if ".txt" in path:
            # txt类型附件
            txt_name = path.split("/")[-1]
            part = MIMEApplication(open(path, 'rb').read())
            # part.add_header('Content-Disposition', 'attachment', filename=txt_name)
            part.add_header('Content-Disposition', 'attachment', filename=('gbk', '', txt_name))
            encoders.encode_base64(part)
            msgRoot.attach(part)


        if ".mp3" in path:
            # mp3类型附件
            mp3_name = path.split("/")[-1]
            part = MIMEApplication(open(path, 'rb').read())
            part.add_header('Content-Disposition', 'attachment', filename=mp3_name)
            msgRoot.attach(part)

    try:
        # s = smtplib.SMTP()
        # s.set_debuglevel(1)
        # print(msgRoot.as_string())
        s = smtplib.SMTP()
        s.connect("smtp.163.com")  # 这里我使用的是阿里云邮箱,也可以使用163邮箱：smtp.163.com
        s.login(sender, passwd)
        s.sendmail(sender, receivers, msgRoot.as_string())
        print(subject,"邮件发送成功")
    except smtplib.SMTPException as e:
        print(e)
        print("Error, 发送失败")
    finally:
        s.quit()

now =  time.strftime('%Y-%m-%d', time.localtime(time.time()))
subject = "2019届校招内推汇总表（更新至{0}.{1}）-有你Offer出品".format(now.split('-')[1].lstrip('0'), now.split('-')[2].lstrip('0'))
# subject = "有你Offer全体成员祝你中秋节快乐哦"
content = '小主，提前祝您国庆快乐哦！您的秋招内推汇总到了，请查收！感谢订阅有你Offer的秋招内推信息服务，有你Offer是面向全球求职者的一站式求职服务平台。我们拥有海量的实习信息、求职干货和就业指导，并提供1v1简历指导课程及1v1面试指导课程，助您斩获理想的offer~更多详情敬请关注【有你Offer】公众号~'
# content = '小主，有你Offer全体成员祝你中秋节快乐哦!'+'\n'+'愿每位小主都能早日斩获心仪Offer。'+'\n'+'很抱歉告知大家，今天的内推订阅服务因假期停更噢，但有你Offer的心永远与你同在~~~'
jpg_path = '../data/img/有你Offer.jpg'
# pdf_path = "c.pdf"
# txt_path = "d\\e\\f.txt"
xlsx_path = "../data/work/2019届校招内推汇总表（更新至{0}.{1}）-有你Offer出品.xlsx".format(now.split('-')[1].lstrip('0'), now.split('-')[2].lstrip('0'))
file_path = [xlsx_path]  #发送三个文件到两个邮箱
test_email = ['635516607@qq.com', 'lvzeqin@126.com']

nov1108 = [
    '744568921@qq.com',
    'haxidata@163.com',
    'fsltus@126.com',
    'jieeliao@qq.com',
    'licong6990@outlook.com',
    '380521340@qq.com',
    'sulunkang@163.com',
    '812299133@qq.com',
    '575546908@qq.com',
    'tajiaayan@126.com',
    'chlw25@163.com',
    '892893234@qq.com',
    'dongxinyuancrystal@163.com',
    'lenagulina@163.com',
    'xuy2017@mpacc.nai.edu.cn',
    '532817138@qq.com',
    '1723041930@qq.com',
    'wymm520@126.com',
    '441811919@qq.com',
    '334679466@qq.com',
    '2044750616@qq.com',
    'taomy0522@163.com',
    'yeq1702@126.com',
    '18810589317@163.com',
]

nov1110 = [
    '491660265@qq.com',
    '641213180@qq.com',
    '2523698440@qq.com',
    '1175618019@qq.com',
    'mistmichael1994@gmail.com',
    '13322457055@163.com',
    '1427843763@qq.com',
    '2545546799@qq.com',
    '1043460563@qq.com',
    '513305365@qq.com',
    '531292087@qq.com',
    '602589698@qq.com',
    '991891149@qq.com',
    'shibocheng33@foxmail.com',
    '545333288@qq.com',
    '447945592@qq.com',
    'sun125210206@hotmail.com',
    'minying.kong@outlook.com',
    '775642161@qq.com',
    '1262119059@qq.com',
    'qiaodan13@yeah.net',
    'andrewchia@foxmail.com',
]

nov1112 = [
    'yyxgio94@163.com',
    'cynthiafu9508@gmail.com',
    'liu.yaqi@husky.neu.edu',
    'jihoyi@foxmail.com',
    '1604858331@qq.com',
    'Adley_Ren@163.com',
    '745594319@qq.com',
    '729774384@qq.com',
    'vanorazyw@163.com',
    'zjhseu1995@163.com',
    '1214268412@qq.com',
    'vita.yuhua.zhang@outlook.com',
    '516601367@qq.com',
    'plus_wqn@163.com',
    'sqteen177@163.com',
    '690793055@qq.com',
    '464081842@qq.com',
    '939540558@qq.com',
    'litmary@163.com',
    '574205255@qq.com',
    'Lindsayemp@163.com',
]

nov1114 = [
    '897988295@qq.com',
    'rl_96118@163.com',
    '554631355@qq.com',
    'chengyusun18@163.com',
    '871558974@qq.com',
    '295988190@qq.com',
    'vita.yuhua.zhang@outlook.com',
    '1213767565@qq.com',
    'fantasticbaobao@163.com'
]

nov1116 = [
    'sherodjs@163.com'
]

nov1124 = [
    '473155897@qq.com'
]

nov1126 = [
    '741301256@qq.com',
    '3294674415@qq.com',
    '1029696036@qq.com',
]

nov1128 = [
    'wangyue0710@163.com',
]

vip = [
    'wangyue0710@163.com',
]

dec1209 = ['18721225119@163.com']

oct1010 = ['mia.bai.xxx@gmail.com']

dec1211 = ['zhangmengqi0331@163.com', 'swing.chen@outlook.com', '18811405506@163.com', 'sherxy16@163.com']

oct1013 = ['348419137@qq.com', '704327115@qq.com', '472327077@qq.com', '897979317@qq.com', '1070461019@qq.com']

dec1213 = ['767035966@qq.com', 'zhoujingyi1993@163.com', 'yueqizhang127@163.com', '1823480524@qq.com', '617273459@qq.com', 'cristal_jj@163.com', '17607133268@163.com', '403607837@qq.com']

dec1215 = ['Ldq527@163.com', '1161229739@qq.com', '1540326006@qq.com', '3293581247@qq.com']

oct1001 = ['617642097@qq.com']
dec1203 = ['362855216@qq.com']

oct1017 = ['2775907278@qq.com', '154434273@qq.com']

dec1217 = ['15863717312@163.com', 'Guodingsan@outlook.com']
dec1219 = ['1540621110@qq.com', 'ydc19940214@163.com', 'weiyaxin1997@126.com', '17805132858@163.com', '18817870461@163.com']
oct1019 = ['1311339200@qq.com', '505612217@qq.com', '1398113974@qq.com', '1451792508@qq.com', 'dayvivid@163.com']
sep0926 = ['huanghuang9503@outlook.com', '496197883@qq.com']
oct1020 = ['HouMengjuan123@163.com']
oct1023 = ['flowergogo@163.com']
oct1025 = ['bkqiuhuan@163.com']
oct1002 = ['frankachen@163.com']
oct1030 = ['xt17510@163.com']
# 全部的

for emails in [oct1030, oct1002, oct1025, test_email, oct1023, oct1020, dec1219, oct1019, dec1217, oct1017, dec1215, oct1013, dec1213, dec1211, oct1010, oct1001, nov1108, nov1110, nov1112, nov1114, nov1116, nov1124, nov1126, nov1128,dec1203,vip, dec1209]:
    for receive_email in emails:
        print(receive_email)
        # time.sleep(4)
        Send_email_text(subject, content, file_path, receive_email)

# 落下的
# for emails in [test_email]:
#     for receive_email in emails:
#         print(receive_email)
#         Send_email_text(subject, content, file_path, receive_email)

