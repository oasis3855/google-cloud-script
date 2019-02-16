#!/usr/bin/env python
# -*- coding: utf-8 -*-

# PyDrive : web camera jpeg uploader for Google Drive
# Webカメラのjpeg画像をGoogle Driveにアップロードするスクリプト

# (C) INOUE Hirokazu
# GNU GPL Free Software (http://www.opensource.jp/gpl/gpl.ja.html)
#
# version 1.0 (2019/Feb/07)

import datetime
import commands
import sys
import os

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# 指定フォルダのID
# GoogleドライブをWebブラウザで表示した時のURLよりID部分をコピーする
drive_folder_id = '5Ycn_hYk36Niemo-mRcZua3j1bQV2u8Df'

# スクリプト自身のディレクトリに、client_secrets.json と settings.yaml が保存されていること
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Webカメラ画像の取得
time_now = datetime.datetime.now()
image_filename = format(time_now, "image_%Y%m%d-%H%M%S.jpg")
temp_filename = "/tmp/temp_webcamera.jpg"
if os.path.exists(temp_filename) :
    os.remove(temp_filename)
str_result = commands.getoutput("fswebcam -D 3 -S 50 -v -r 640x480 %s" % temp_filename)

if( os.path.exists(temp_filename) is not True) :
    sys.exit()

try :
    # PyDriveでGoogleドライブに接続
    gauth = GoogleAuth()
    gauth.CommandLineAuth()
    drive = GoogleDrive(gauth)

    # 画像ファイルのアップロード
    file = drive.CreateFile({'title' : image_filename,
                'parents' : [{'id' : drive_folder_id}],
                'mimeType' : 'image/jpeg'})
    file.SetContentFile(temp_filename)
    file.Upload()
except Exception, e:
    print "===== PyDrive fatal error ====="
    print e
    sys.exit()
finally:
    if os.path.exists(temp_filename) :
        os.remove(temp_filename)

