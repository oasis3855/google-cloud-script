#!/usr/bin/env python
# -*- coding: utf-8 -*-

# PyDrive : auto delete old files
# 指定されたGoogleDriveフォルダ内の古いファイルを一括削除するスクリプト

# (C) INOUE Hirokazu
# GNU GPL Free Software (http://www.opensource.jp/gpl/gpl.ja.html)
#
# version 1.0 (2019/Feb/07)

import sys
import datetime
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# 指定フォルダのID
# GoogleドライブをWebブラウザで表示した時のURLよりID部分をコピーする
drive_folder_id = '5Ycn_hYk36Niemo-mRcZua3j1bQV2u8Df'
# ファイル名のパターン（ファイル名の先頭または末尾に一致）
filename_pattern = 'image_20'
# 24h * 2回/h * 20日 = 960回
max_files_limit = 960
# メッセージ出力の有効化
message_enable = 1

if len(sys.argv) == 2 and sys.argv[1] == '--nomessage':
    message_enable = 0
else :
    print("PyDrive auto delete old files\n  usage :\n  %s [--nomessage]\n" % sys.argv[0])

# スクリプト自身のディレクトリに、client_secrets.json と settings.yaml が保存されていること
import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

try :

    if message_enable == 1 :
        print('connect to Google Drive ...')

    gauth = GoogleAuth()
    gauth.CommandLineAuth()
    drive = GoogleDrive(gauth)

    if message_enable == 1 :
        print('caluculate number of files ...')

    # ファイル数を数えるため、ファイル一覧を得る
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false and mimeType='image/jpeg' and title contains '%s'" %
                    (drive_folder_id, filename_pattern)}).GetList()

    current_files_count = len(file_list)

    # ファイル数が規定に達していない時は、削除処理を行わない
    if current_files_count <= max_files_limit :
        if message_enable == 1 :
            print("current %d files, is under limit %d\nexit script" %
            (current_files_count, max_files_limit))
        sys.exit()

    if message_enable == 1 :
        print("current %d files, now exec task to delete %d files" %
        (current_files_count, current_files_count - max_files_limit))

    if message_enable == 1 :
        print('getting file list ...')

    # 削除のためのファイル一覧を得る
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false and mimeType='image/jpeg' and title contains '%s'" %
                    (drive_folder_id, filename_pattern),
                    'maxResults' : (current_files_count - max_files_limit),
                    'orderBy' : 'createdDate'}).GetList()

    # ファイルを削除する
    for file1 in file_list:
        if message_enable == 1 :
            print('delete title: %s' % file1['title'])
        file1.Delete()

except Exception, e:
    print "===== PyDrive fatal error ====="
    print e
    sys.exit()

if message_enable == 1 :
    print("done all")
