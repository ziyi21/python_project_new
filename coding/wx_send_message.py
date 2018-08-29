#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @File  : wx_send_message.py
# @Author: ziyi
# @Date  : 2018/8/29
# @Desc  :


import itchat
import time
import datetime


def timerfun():

    while True:
        now = datetime.datetime.now()
        current_time = time.localtime(time.time())

        if ((current_time.tm_hour == 8) and (current_time.tm_min == 35)):
            print('send-case')
            case_path = '../data/chatrooms/internet_information.txt'
            case_contents = get_contents(case_path)
            if case_contents:
                print(case_contents)
                sent_information(case_contents)
                print(now, '互联网行业的实习信息发送成功')
                with open('../data/chatrooms/all_work.txt', 'a') as f:
                    f.write('\n')
                    f.write('\n')
                    f.write('\n')
                    f.write(str(now))
                    f.write('\n')
                    f.write(case_contents)
            time.sleep(60)  # 每次判断间隔1s，避免多次触发事件


def sent_information(send_contents):
    chatrooms = itchat.get_chatrooms(update=True)
    print('手机中的群聊数目为：', len(chatrooms))
    j = 0
    # 互联网类的群
    for i, one_name in enumerate(chatrooms):
        name = one_name['NickName']
        # print('开始判断')
        if '实习有你' in name or '哈希大数据' in name or '智新赋能' in name or 'Deep learning' in name or '苏州市人工智能' in name or '有你offer' \
                in name or '[2.' in name or '[1.' in name or '南大 大数据' in name or '数据魔术师' in name or '数据家交流群' in name \
                or '有你Offer' in name or 'Deep Learning' in name or '南京高校学联' in name or '数据分析与应用' in name \
                or '数据挖掘与案例分析' in name or '苏商学院' in name or '南京区块链联盟' in name or '高级研修班' in name:
            if '公安' not in name and '栖霞' not in name and '世界杯' not in name and '优惠' not in name and '哈希大数据' != \
                    name.replace(' ', ''):
                print(name)
                searchName = name
                iRoom = itchat.search_chatrooms(searchName)
                j = j + 1
                for room in iRoom:
                    if room['NickName'] == searchName:
                        userName = room['UserName']
                        print(j, i, room['NickName'])
                        # 实现发送消息的功能
                        itchat.send_msg(send_contents, userName)
                        print(userName, '发送成功')
                        time.sleep(2)


def get_contents(path):
    send_contents = None
    try:
        with open(path, encoding='utf16') as f:
            contents = f.readlines()
        send_contents = ''.join(content for content in contents)
    except Exception as e:
        print('不是utf16的编码格式',e)

    try:
        with open(path, encoding='utf8') as f:
            contents = f.readlines()
        send_contents = ''.join(content for content in contents)
    except Exception as e:
        print('不是utf8的编码格式',e)
    return send_contents


if __name__ == '__main__':
    # itchat.auto_login()
    itchat.auto_login(hotReload=True)  # 首次扫描登录后后续直接确认登录，无需再扫描
    path = '../data/chatrooms/management_information.txt'
    nowtime = datetime.datetime.now()
    print(nowtime)
    timerfun()