#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Google Calendar iCal file import scrypt
# GoogleカレンダーにiCal形式ファイルをインポートするPythonスクリプト
#
# (C) 2014 INOUE Hirokazu
# GNU GPL Free Software (GPL Version 2)
#
# Version 0.1 (2014/03/03)
# Version 0.2 (2014/03/12)
# Version 0.3 (2014/03/14)
#

import sys
import os
import pwd

# ユーザディレクトリにモジュールをインストールした場合、環境変数 PYTHONPATH を
# 設定するか、次のようにスクリプト中でインストールディレクトリを指定する。
# 【インストール時のコマンドライン ： python ./setup.py install --home=~ 】
sys.path.append(os.path.expanduser("~") + '/lib/python')
sys.path.append(os.path.expanduser("~") + '/lib/python/icalendar-3.6.1-py2.7.egg')
sys.path.append(os.path.expanduser("~") + '/lib/python/dateutils-0.6.6-py2.7.egg')
sys.path.append(os.path.expanduser("~") + '/lib/python/pytz-2013.9-py2.7.egg')
sys.path.append(os.path.expanduser("~") + '/lib/python/python_dateutil-2.2-py2.7.egg')
sys.path.append(os.path.expanduser("~") + '/lib/python/six-1.5.2-py2.7.egg')

import argparse
import time
import datetime
import ConfigParser
import base64
import atom
import icalendar
import gdata.calendar.client

#####
# グローバル変数
flag_verbose_stdout = False
flag_silent_stdout = False
flag_new_event_only = False

#####
# シェル内実行ではない場合、ユーザルートディレクトリの指定（Webプログラムから実行の場合等）
config_file_root = pwd.getpwuid(os.getuid())[5]
if config_file_root[-1:] != '/':
    config_file_root = config_file_root + '/'

#####
# iCalファイルよりEventを読み取り、リストに格納する
# 戻り値 : list_schedules
def read_ical_file(filename):

    # Eventリスト（戻り値）
    list_schedules = []

    # icalファイルをテキスト str_icsdata に格納する
    if 1:
#    try:
        fh = open(filename, "r")
        str_icsdata = fh.read()
        fh.close()
    #except:
        #if flag_silent_stdout == False:
            #print "iCal file open error"
        #return

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

    global flag_silent_stdout
    global flag_new_event_only

    # Eventリストが空の場合は何もしない
    if list_schedules is None or len(list_schedules) <= 0:
        if flag_silent_stdout == False:
            print "No Event in iCal file"
        return 1

    print "iCal file read success, contain data number = ",
    print len(list_schedules)

    # Google カレンダーサービスに接続する
    try:
        calendar_service = gdata.calendar.client.CalendarClient()
        calendar_service.ssl = True
        calendar_service.ClientLogin(login_user, login_password, "test python script");
    except:
        if flag_silent_stdout == False:
            print "Google logon authenticate error"
        return 1
    if flag_verbose_stdout is True:
        print "Google Calendar API, connection succeed"
        print "  gdata.calendar.client.api_version = " + calendar_service.api_version
        print "  gdata.calendar.client.contact_list = " + calendar_service.contact_list
        print "  gdata.calendar.client.server  = " + calendar_service.server
        print "  gdata.calendar.client.ssl  = ",
        print calendar_service.ssl
        print "  Login User = " + login_user

    # Eventリストを1つずつ登録済みかチェックし、未登録の場合はEventの新規登録を行う
    try:
        for list_schedule_item in list_schedules:
            # Googleカレンダーに同一日時でEventが登録済みかチェック
            if check_event_google_calendar(calendar_service, list_schedule_item) == True:
                if flag_silent_stdout == False:
                    print "same(skip) : " + list_schedule_item["title"].encode('utf8') + "(" + \
                            list_schedule_item["start"].strftime("%Y/%m/%d") + ")"
            # iCalのEventが、現在日時より過去のものでないかチェック
            elif flag_new_event_only == True and ((type(list_schedule_item["start"]) is datetime.date \
                and list_schedule_item["start"] < datetime.date.today()) or \
                (type(list_schedule_item["start"]) is datetime.datetime and \
                list_schedule_item["start"].replace(tzinfo=None) < datetime.datetime.today())):
                if flag_silent_stdout == False:
                    print "old(skip) : " + list_schedule_item["title"].encode('utf8') + "(" + \
                            list_schedule_item["start"].strftime("%Y/%m/%d") + ")"
            # Googleカレンダーに未登録で、現在より未来のEventであれば…
            else:
                if flag_silent_stdout == False:
                    print "new add    : " + list_schedule_item["title"].encode('utf8') + "(" + \
                            list_schedule_item["start"].strftime("%Y/%m/%d") + ")"
                    print "  place : " + list_schedule_item["place"].encode('utf8')
                    print "  time : ",
                    print list_schedule_item["start"],
                    print " 〜 ",
                    print list_schedule_item["end"]
                # EventをGoogleカレンダーに新規登録
                insert_event_google_calendar(calendar_service, list_schedule_item)
    except:
        if flag_silent_stdout == False:
            print "Google calendar access error"
        return 1

    return 0

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
    if flag_verbose_stdout is True:
        print "checking event data ..."
        print "  title : " + list_schedule_item["title"].encode('utf8')
        print "  start : ",
        print list_schedule_item["start"]
        print "  end   : ",
        print list_schedule_item["end"]
        print "  place : " + list_schedule_item["place"].encode('utf8')
        print "  desc  : " + list_schedule_item["desc"].encode('utf8')
        print "  update: ",
        print list_schedule_item["updated"]

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
    if 'HOME' in os.environ:
        configfile = os.path.join(os.environ['HOME'], '.google-mail-python-progs')
    else:
        configfile = config_file_root + '.google-mail-python-progs'

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
    if 'HOME' in os.environ:
        configfile = os.path.join(os.environ['HOME'], '.google-mail-python-progs')
    else:
        configfile = config_file_root + '.google-mail-python-progs'

    try:
        fp = open(configfile, 'w')
        parser.set("DEFAULT", "example@gmail.com", user)
        parser.set("DEFAULT", "password", base64.b64encode(password))
        parser.write(fp)
        fp.close()
    except IOError:
        print >> sys.stderr, "config write error"


#####
# Main : プログラム エントリーポイント

parser = argparse.ArgumentParser(description='Import iCal file to Google Calendar')
parser.add_argument('ical_filename', nargs='?', help='iCal file to import')
parser.add_argument('-u', help='user name of google calendar (example@gmail.com)', metavar='User')
parser.add_argument('-p', help='password', metavar='Password')
parser.add_argument('-n', help='add future schedule only', action="store_true")
parser.add_argument('-v', help='verbose output message', action="store_true")
parser.add_argument('-s', help='silent mode (without stdout message)', action="store_true")
parser.add_argument('-auth_nosave', help='do not save user/password data', action="store_true")

args = parser.parse_args()

if args.ical_filename is None:
    print "iCal file name is missing\n-h option for help"
    exit(1)
elif not os.path.isfile(args.ical_filename):
    print "iCal file ["+args.ical_filename+"] is not exist"
    exit(1)
ics_file = args.ical_filename

if (args.u is not None and args.p is None) or (args.u is None and args.p is not None):
    print "user or password is missing"
    exit(1)
if args.u is not None:
    user = args.u
    password = args.p
else:
    user, password = config_file_read("example@gmail.com", "password")

if args.n is True:
    flag_new_event_only = True

if args.s is True:
    flag_silent_stdout = True
if args.v is True:
    flag_verbose_stdout = True
if args.s is True and args.v is True:
    print "-v and -s option is incompatible"
    exit(1)

# 処理ファイル名の画面表示
if flag_silent_stdout == False:
    print "iCal file : [" + ics_file +"]"
    if flag_new_event_only == True:
        print "add future schedule only mode : enable"

list_schedules = read_ical_file(ics_file)
result_value = insert_event_list_google_calendar(list_schedules, user, password)

# 正常終了で、ユーザ名/パスワードが入力されていた場合は、設定ファイルを書き換える
if args.auth_nosave is False:
    if args.u is not None and result_value == 0:
        config_file_write(user, password)

# スクリプトの戻り値 0:正常, 1:エラー
sys.exit(result_value)
