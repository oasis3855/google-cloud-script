#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Googleカレンダーの予定（イベント）を読み出す
# test version 0.1 (2014/03/01)

# Google Data APIs : gdata-python-client が必要
# https://developers.google.com/gdata/docs/client-libraries

import sys
import gtk
import re
import datetime
import gdata.calendar.client
import os
import ConfigParser
import base64

def print_event(user, password):

    # デフォルトの検索開始・終了 日時
    datetime_start = datetime.date.today()
    datetime_end = datetime.date.today() + datetime.timedelta(31)
    # 最大検索数
    max_results = 5

    # 検索開始年月日と、ユーザ名・パスワードを入力するダイアログを表示する
    dlg = gtk.MessageDialog(type=gtk.MESSAGE_INFO, buttons=gtk.BUTTONS_OK_CANCEL, message_format='条件の設定を行います')
    dlg.set_title("GoogleカレンダーReader")
    dlg_content = dlg.get_content_area()
    hbox_calendar = gtk.HBox()

    vbox_calendar_start = gtk.VBox()
    vbox_calendar_start.add(gtk.Label("チェック開始日"))
    calendar_start = gtk.Calendar()
    calendar_start.select_month(datetime_start.month-1, datetime_start.year)
    calendar_start.select_day(datetime_start.day)
    hbox_monthday_start = gtk.HBox()
    hour_start = gtk.Adjustment(value=0, lower=0, upper=12, step_incr=1)
    spin_hour_start = gtk.SpinButton(hour_start)
    minute_start = gtk.Adjustment(value=0, lower=0, upper=59, step_incr=1)
    spin_minute_start = gtk.SpinButton(minute_start)
    hbox_monthday_start.add(spin_hour_start)
    hbox_monthday_start.add(gtk.Label("時"))
    hbox_monthday_start.add(spin_minute_start)
    hbox_monthday_start.add(gtk.Label("分"))
    vbox_calendar_start.add(calendar_start)
    vbox_calendar_start.add(hbox_monthday_start)

    vbox_calendar_end = gtk.VBox()
    vbox_calendar_end.add(gtk.Label("チェック終了日"))
    calendar_end = gtk.Calendar()
    calendar_end.select_month(datetime_end.month-1, datetime_end.year)
    calendar_end.select_day(datetime_end.day)
    hbox_monthday_end = gtk.HBox()
    hour_end = gtk.Adjustment(value=0, lower=0, upper=12, step_incr=1)
    spin_hour_end = gtk.SpinButton(hour_end)
    minute_end = gtk.Adjustment(value=0, lower=0, upper=59, step_incr=1)
    spin_minute_end = gtk.SpinButton(minute_end)
    hbox_monthday_end.add(spin_hour_end)
    hbox_monthday_end.add(gtk.Label("時"))
    hbox_monthday_end.add(spin_minute_end)
    hbox_monthday_end.add(gtk.Label("分"))
    vbox_calendar_end.add(calendar_end)
    vbox_calendar_end.add(hbox_monthday_end)

    hbox_calendar.add(vbox_calendar_start)
    hbox_calendar.add(vbox_calendar_end)
    dlg_content.add(hbox_calendar)

    hbox_results = gtk.HBox()
    hbox_results.add(gtk.Label("最大検索数"))
    results = gtk.Adjustment(value=5, lower=1, upper=20, step_incr=1)
    spin_results = gtk.SpinButton(results)
    hbox_results.add(spin_results)
    dlg_content.add(hbox_results)

    hbox_user = gtk.HBox()
    hbox_user.add(gtk.Label("Googleユーザ名"))
    text_entry_user = gtk.Entry()
    text_entry_user.set_text(user)
    hbox_user.add(text_entry_user)
    dlg_content.add(hbox_user)

    hbox_password = gtk.HBox()
    hbox_password.add(gtk.Label("パスワード"))
    text_entry_password = gtk.Entry()
    text_entry_password.set_width_chars(30)
    text_entry_password.set_visibility(False)
    text_entry_password.set_text(password)
    hbox_password.add(text_entry_password)
    dlg_content.add(hbox_password)

    dlg.show_all()
    ret = dlg.run()

    # ダイアログに設定された内容を読み取る
    if ret == gtk.RESPONSE_OK:
        # 検索開始日
        year, month, day = calendar_start.get_date()
        month = month + 1
        hour = hour_start.get_value()
        minute = minute_start.get_value()
        datetime_start = datetime.datetime(int(year), int(month), int(day), int(hour), int(minute), 0)
        # 検索終了日
        year, month, day = calendar_end.get_date()
        month = month + 1
        hour = hour_end.get_value()
        minute = minute_end.get_value()
        datetime_end = datetime.datetime(int(year), int(month), int(day), int(hour), int(minute), 0)
        if datetime_end < datetime_start:
            datetime_end = datetime_start + datetime.timedelta(31)
        # 最大検索数
        max_results = int(results.get_value())
        # ユーザ名とパスワード
        user = text_entry_user.get_text().decode('utf8')
        password = text_entry_password.get_text().decode('utf8')
    else:
        return
    dlg.destroy()

    # Googleサーバに接続し、カレンダー イベントを読み出す
    try:
        list_schedules = read_event_google_calendar(datetime_start, datetime_end, max_results, user, password)
    except:
        # 読み出し失敗のメッセージボックスを表示
        dlg = gtk.MessageDialog(type=gtk.MESSAGE_INFO, buttons=gtk.BUTTONS_OK, message_format='Googleカレンダーから読み出せません')
        dlg.set_title("GoogleカレンダーReader")
        dlg.show_all()
        ret = dlg.run()
        dlg.destroy()
        return

    # ユーザ名とパスワードを設定ファイルに書きこんで保存する
    config_file_write(user, password)

    # カレンダー イベントをダイアログに表示する
    dlg = gtk.MessageDialog(type=gtk.MESSAGE_WARNING, buttons=gtk.BUTTONS_OK, message_format='Googleカレンダー')
    str_message = ""
    for list_item in list_schedules:
        time_start_date = parse_datetime_str(list_item["start"])
        str_start_date = "%4d/%02d/%02d" % (time_start_date.year, time_start_date.month, time_start_date.day)
        str_message += (str_start_date + " : " + list_item["title"] + "\n")
    dlg.format_secondary_text(str_message)
    dlg.run()
    dlg.destroy()


# Googleサーバに接続し、カレンダー イベントを読み出す
def read_event_google_calendar(datetime_start, datetime_end, max_results, login_user, login_password):

    list_schedules = []

    str_date_start = '%4d-%02d-%02d' % (datetime_start.year, datetime_start.month, datetime_start.day)
    str_date_end = '%4d-%02d-%02d' % (datetime_end.year, datetime_end.month, datetime_end.day)

    calendar_service = gdata.calendar.client.CalendarClient()
    calendar_service.ssl = True
    calendar_service.ClientLogin(login_user, login_password, "test python script");

    query = gdata.calendar.client.CalendarEventQuery(start_min=str_date_start, start_max=str_date_end,
                max_results=max_results, orderby='starttime', sortorder='ascending')
    feed = calendar_service.GetCalendarEventFeed(q=query)

    for i,ev in enumerate(feed.entry):
        for e in ev.when:
            dict_schedule = {"title":ev.title.text, "start":e.start, "end":e.end, "updated":ev.updated.text}
            list_schedules.append(dict_schedule)
            break

    return list_schedules

def parse_datetime_str(str_datetime):
    # Googleカレンダーの日時文字列（"2013-01-01T12:59:59.000+09:00" or "2013-01-01"）を
    # struct_time 形式に変換する

    datetime_ret = datetime.datetime(1990,1,1)

    # 末尾のTZ（+09:00）があれば除去する
    str_datetime = str_datetime.split("+")[0]

    try:
        # 最低限の変換をまず試してみる
        list_datetime = re.split("[\-/T:\.]", str_datetime)
        if len(list_datetime) >= 3:
            datetime_ret = datetime.datetime(int(list_datetime[0]),int(list_datetime[1]),int(list_datetime[2]))

        if len(str_datetime) == 23:
            # "2013-01-01T12:59:59.000" == 23文字
            datetime_ret = datetime.datetime.strptime(str_datetime, '%Y-%m-%dT%H:%M:%S.000')
        elif len(str_datetime) == 10:
            # "2013-01-01" == 10文字
            datetime_ret = datetime.datetime.strptime(str_datetime, '%Y-%m-%d')
        else:
            raise
    except:
        pass    # do nothing

    return datetime_ret

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


print "Read from Google Calendar"
user, password = config_file_read("example@gmail.com", "password")
print_event(user, password)
