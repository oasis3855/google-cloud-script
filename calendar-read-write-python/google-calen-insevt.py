#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Googleカレンダーに予定（イベント）を書き込む
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
import atom

def insert_event(user, password):

    # デフォルトのイベント開始・終了 日時
    datetime_start = datetime.date.today()
    datetime_end = datetime.date.today()

    # 検索開始年月日と、ユーザ名・パスワードを入力するダイアログを表示する
    dlg = gtk.MessageDialog(type=gtk.MESSAGE_INFO, buttons=gtk.BUTTONS_OK_CANCEL, message_format='条件の設定を行います')
    dlg.set_title("GoogleカレンダーWriter")
    dlg_content = dlg.get_content_area()

    hbox_title = gtk.HBox()
    hbox_title.add(gtk.Label("予定の名称"))
    text_entry_title = gtk.Entry()
    text_entry_title.set_text("")
    hbox_title.add(text_entry_title)
    dlg_content.add(hbox_title)

    hbox_place = gtk.HBox()
    hbox_place.add(gtk.Label("場所"))
    text_entry_place = gtk.Entry()
    text_entry_place.set_text("")
    hbox_place.add(text_entry_place)
    dlg_content.add(hbox_place)

    hbox_content = gtk.HBox()
    hbox_content.add(gtk.Label("説明"))
    text_entry_content = gtk.Entry()
    text_entry_content.set_text("")
    hbox_content.add(text_entry_content)
    dlg_content.add(hbox_content)

    check_allday = gtk.CheckButton("終日の予定（時刻を無視する）")
    check_allday.set_active(True)
    dlg_content.add(check_allday)

    hbox_calendar = gtk.HBox()

    vbox_calendar_start = gtk.VBox()
    vbox_calendar_start.add(gtk.Label("予定開始日"))
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
    vbox_calendar_end.add(gtk.Label("予定終了日"))
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
        # 予定
        title = text_entry_title.get_text().decode('utf8')
        place = text_entry_place.get_text().decode('utf8')
        content = text_entry_content.get_text().decode('utf8')
        # 開始日
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
            datetime_end = datetime_start
        # 終日
        allday = check_allday.get_active()
        # ユーザ名とパスワード
        user = text_entry_user.get_text().decode('utf8')
        password = text_entry_password.get_text().decode('utf8')
    else:
        return
    dlg.destroy()

    # Googleサーバに接続し、カレンダー イベントを追加する
    try:
        insert_event_google_calendar(datetime_start, datetime_end, allday, title, place, content, user, password)
    except:
        # 書きこみ失敗のメッセージボックスを表示
        dlg = gtk.MessageDialog(type=gtk.MESSAGE_INFO, buttons=gtk.BUTTONS_OK, message_format='Googleカレンダーに書き込めません')
        dlg.set_title("GoogleカレンダーWriter")
        dlg.show_all()
        ret = dlg.run()
        dlg.destroy()
        return

    # ユーザ名とパスワードを設定ファイルに書きこんで保存する
    config_file_write(user, password)

    # 書きこみ成功のメッセージボックスを表示
    dlg = gtk.MessageDialog(type=gtk.MESSAGE_INFO, buttons=gtk.BUTTONS_OK, message_format='予定を書き込みました')
    dlg.set_title("GoogleカレンダーWriter")
    dlg.show_all()
    ret = dlg.run()
    dlg.destroy()

# Googleサーバに接続し、カレンダー イベントを追加する
def insert_event_google_calendar(datetime_start, datetime_end, allday, title, place, content, login_user, login_password):

    calendar_service = gdata.calendar.client.CalendarClient()
    calendar_service.ssl = True
    calendar_service.ClientLogin(login_user, login_password, "test python script");

    event = gdata.calendar.data.CalendarEventEntry()
    event.title = atom.data.Title(text=title)
    event.content = atom.data.Content(text=content)
    event.where.append(gdata.calendar.data.CalendarWhere(value=place))
    if allday == True:
        start_time = '%d-%02d-%02d' % (datetime_start.year, datetime_start.month, datetime_start.day)
        end_time = '%d-%02d-%02d' % (datetime_end.year, datetime_end.month, datetime_end.day)
    else:
        start_time = '%d-%02d-%02dT%02d:%02d:00.000' % (datetime_start.year, datetime_start.month,
            datetime_start.day, datetime_start.hour, datetime_start.minute)
        end_time = '%d-%02d-%02dT%02d:%02d:00.000' % (datetime_end.year, datetime_end.month,
            datetime_end.day, datetime_end.hour, datetime_end.minute)
    event.when.append(gdata.data.When(start=start_time, end=end_time))

    calendar_service.InsertEvent(event)

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


print "Insert Event to Google Calendar"
user, password = config_file_read("example@gmail.com", "password")
insert_event(user, password)
