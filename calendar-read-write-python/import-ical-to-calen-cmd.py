#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Google Calendar iCal file import scrypt
# GoogleカレンダーにiCal形式ファイルをインポートするPythonスクリプト
#
# (C) 2014 INOUE Hirokazu
# GNU GPL Free Software (GPL Version 2)
#
# Version 0.1 (2014/03/03)
#

import sys
import os

# ユーザディレクトリにモジュールをインストールした場合、環境変数 PYTHONPATH を
# 設定するか、次のようにスクリプト中でインストールディレクトリを指定する。
# 【インストール時のコマンドライン ： python ./setup.py install --home=~ 】
sys.path.append(os.path.expanduser("~") + '/lib/python')
sys.path.append(os.path.expanduser("~") + '/lib/python/icalendar-3.6.1-py2.7.egg')

import time
import datetime
import ConfigParser
import base64
import atom
import icalendar
import gdata.calendar.client

#####
# iCalファイルよりEventを読み取り、リストに格納する
# 戻り値 : list_schedules
def read_ical_file(filename):

    # Eventリスト（戻り値）
    list_schedules = []

    # icalファイルをテキスト str_icsdata に格納する
    try:
        fh = open(filename, "r")
        str_icsdata = fh.read()
        fh.close()
    except:
        print "iCal file open error"
        return

    # icalファイル「テキスト」を解析し cal に取り込む
    cal = icalendar.Calendar.from_ical(str_icsdata)

    for e in cal.walk():
        if e.name == 'VEVENT' :
            # Eventを1つずつ、辞書形式dict_scheduleに一旦代入し、それをリストlist_schedulesに追加する
            dict_schedule = {"title":unicode(e.decoded("summary"),'utf8') if e.get("summary") else "",
                            "place":unicode(e.decoded("location"),'utf8') if e.get("location") else "",
                            "desc":unicode(e.decoded("description"),'utf8') if e.get("description") else "",
                            "start":e.decoded("dtstart"),
                            "end":e.decoded("dtend"),
                            "updated":e.decoded("dtstamp")
                            }
            list_schedules.append(dict_schedule)

    # 予定表Eventを格納したリストを返す
    return list_schedules

#####
# Googleサーバに接続し、カレンダー イベントを追加する
def insert_event_list_google_calendar(list_schedules, login_user, login_password):

    # Eventリストが空の場合は何もしない
    if list_schedules is None or len(list_schedules) <= 0:
        print "No Event in iCal file"
        return

    # Google カレンダーサービスに接続する
    try:
        calendar_service = gdata.calendar.client.CalendarClient()
        calendar_service.ssl = True
        calendar_service.ClientLogin(login_user, login_password, "test python script");
    except:
        print "Google logon authenticate error"
        return

    # Eventリストを1つずつ登録済みかチェックし、未登録の場合はEventの新規登録を行う
    try:
        for list_schedule_item in list_schedules:
            if check_event_google_calendar(calendar_service, list_schedule_item) == True:
                print "same(skip) : " + list_schedule_item["title"] + "(" + \
                        list_schedule_item["start"].strftime("%Y/%m/%d") + ")"
            else:
                print "new add    : " + list_schedule_item["title"] + "(" + \
                        list_schedule_item["start"].strftime("%Y/%m/%d") + ")"

                # EventをGoogleカレンダーに新規登録
                insert_event_google_calendar(calendar_service, list_schedule_item)
    except:
        print "Google calendar access error"
        return

#####
# GoogleカレンダーにEventを1件新規登録する
def insert_event_google_calendar(calendar_service, list_schedule_item):

    # Event1件のデータをCalendarEventEntryに設定する
    event = gdata.calendar.data.CalendarEventEntry()
    event.title = atom.data.Title(text=list_schedule_item["title"])
    event.where.append(gdata.calendar.data.CalendarWhere(value=list_schedule_item["place"]))
    event.content = atom.data.Content(text=list_schedule_item["desc"])
    if type(list_schedule_item["start"]) is  datetime.date:
        start_time = list_schedule_item["start"].strftime("%Y-%m-%d")
        end_time = list_schedule_item["end"].strftime("%Y-%m-%d")
    else:
        start_time = list_schedule_item["start"].strftime("%Y-%m-%dT%H:%M:%S")
        end_time = list_schedule_item["end"].strftime("%Y-%m-%dT%H:%M:%S")
    event.when.append(gdata.data.When(start=start_time, end=end_time))

    # Googleカレンダーに新規登録
    calendar_service.InsertEvent(event)


# GoogleカレンダーのEventに同一のものがあるかチェックする
def check_event_google_calendar(calendar_service, list_schedule_item):

# Debug : 指定したイベントの開始・終了日時を画面表示
    #print list_schedule_item["start"]
    #print list_schedule_item["end"]

    # 検索開始・終了日時の設定
    if type(list_schedule_item["start"]) is  datetime.date:
        # 終日の予定の場合
        str_date_start = list_schedule_item["start"].strftime("%Y-%m-%dT00:00:00")
        str_date_end = list_schedule_item["end"].strftime("%Y-%m-%dT00:00:01")
    else:
        # 時刻指定の予定の場合 （末尾にtimezoneを指定しないと、うまく検索されない）
        str_date_start = list_schedule_item["start"].strftime("%Y-%m-%dT%H:%M:%S.000+09:00")
        str_date_end = list_schedule_item["end"].strftime("%Y-%m-%dT%H:%M:%S.000+09:00")

# Debug : 指定したイベントの開始・終了日時を画面表示
    #print " " + str_date_start + "〜" + str_date_end

    # Googleカレンダーにクエリ発行
    query = gdata.calendar.client.CalendarEventQuery(start_min=str_date_start, start_max=str_date_end,
                max_results='50', orderby='starttime', sortorder='ascending')
    feed = calendar_service.GetCalendarEventFeed(q=query)

    # クエリ結果に指定されたEventが発見されたら Trueを返す
    for i,ev in enumerate(feed.entry):
        if list_schedule_item["title"] == ev.title.text:
            return True

    return False

#####
# 設定ファイルから読み込む
def config_file_read(user, password):
    parser = ConfigParser.SafeConfigParser()
    configfile = os.path.join(os.environ['HOME'], '.google-mail-python-progs')

    try:
        fp = open(configfile, 'r')
        parser.readfp(fp)
        fp.close()
        user = parser.get("DEFAULT", "example@gmail.com")
        password = base64.b64decode(parser.get("DEFAULT", "password"))
    except:
        # 設定ファイルが読み込めない場合、書き込みを行う（新規作成の時を意図）
        print >> sys.stderr, "config file error (not found or syntax error)"
        config_file_write(user, password)
        return user, password

    return user, password

#####
# 設定ファイルに書き込む
def config_file_write(user, password):
    parser = ConfigParser.SafeConfigParser()
    configfile = os.path.join(os.environ['HOME'], '.google-mail-python-progs')

    try:
        fp = open(configfile, 'w')
        parser.set("DEFAULT", "example@gmail.com", user)
        parser.set("DEFAULT", "password", base64.b64encode(password))
        parser.write(fp)
        fp.close()
    except IOError:
        print >> sys.stderr, "config write error"


print "Import ical file to Google Calendar"
user, password = config_file_read("example@gmail.com", "password")

# スクリプトの引数
argv = sys.argv
if len(argv) != 2:
    print "usage : " + argv[0] + " filename.ics"
    exit()

list_schedules = read_ical_file(argv[1])
insert_event_list_google_calendar(list_schedules, user, password)

