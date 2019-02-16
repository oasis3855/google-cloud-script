#!/usr/bin/env python
# -*- coding: utf-8 -*-


# PyDrive : first time run (setup credentials.json), only display file list
# 最初に実行する（credentials.json セットアップ用）, ファイルリストを表示するだけのスクリプト
#
# コマンドラインで入出力可能な状態で実行すること

# (C) INOUE Hirokazu
# GNU GPL Free Software (http://www.opensource.jp/gpl/gpl.ja.html)
#
# version 1.0 (2019/Feb/07)

import datetime
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# 指定フォルダのID
drive_folder_id = '1Ep_ageG5RMJem-q5LtZDIbLlZ5V2u8Df'
# ルートフォルダ（以下の全フォルダ対象）とする場合の設定
#drive_folder_id = 'root'

# 表示するファイルの数
query_files_limit = 20


# スクリプト自身のディレクトリに、client_secrets.json と settings.yaml が保存されていること
import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

try :

    # Google Drive APIに接続（初回実行時に、credentials.jsonが作成される）
    gauth = GoogleAuth()
    gauth.CommandLineAuth()
    drive = GoogleDrive(gauth)

    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % drive_folder_id,
                    'maxResults' : query_files_limit, 'orderBy' : 'createdDate'}).GetList()

    for file1 in file_list:
        #print('title: %s, id: %s' % (file1['title'], file1['id']))
        #print(file1.items())
        #print("\n\n")
        print(' ' + file1['title'] + ', ' +
            (file1['fileSize'] if 'fileSize' in file1.keys() else '---') + " Bytes, " +
            file1['modifiedDate'] + "\n")

    print("done\n")


except Exception, e:
    print "===== PyDrive fatal error ====="
    print e
    sys.exit()

