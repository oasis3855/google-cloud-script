#!/usr/bin/perl

# Google Calendar iCal file import scrypt (Web interface)
# GoogleカレンダーにiCal形式ファイルをインポートする Web インターフェース
#
# (C) 2014 INOUE Hirokazu
# GNU GPL Free Software (GPL Version 2)
#
# Version 0.1 (2014/03/14)
#
use strict;
use warnings;
use utf8;

use CGI;
use File::Basename;
use File::Copy;

my $flag_os = 'linux';  # linux/windows
my $flag_charcode = 'utf8';     # utf8/shiftjis
# IOの文字コードを規定
if($flag_charcode eq 'utf8'){
    binmode(STDIN, ":utf8");
    binmode(STDOUT, ":utf8");
    binmode(STDERR, ":utf8");
}
if($flag_charcode eq 'shiftjis'){
    binmode(STDIN, "encoding(sjis)");
    binmode(STDOUT, "encoding(sjis)");
    binmode(STDERR, "encoding(sjis)");
}
my $str_this_script = basename($0);             # このスクリプト自身のファイル名
my $str_filepath_ical_tmp = './data/temp.ics';   # アップロード時の一時ファイル名

{       # 利用変数がグローバル化しないように囲う
my $q = new CGI;

# HTML出力を開始する
sub_print_start_html(\$q);

# 処理内容に合わせた処理と、画面表示
if(defined($q->url_param('mode'))){
    if($q->url_param('mode') eq 'logoff'){
        print "<p>ログオフ 完了</p>\n";
    }
    else{
        print("<p class=\"error\">URLパラメータ（mode）が想定外です</p>\n");
    }
}
elsif(defined($q->param('uploadfile')) && length($q->param('uploadfile'))>0){
        sub_upload_ical(\$q);
}
else{
        sub_disp_home();
}

# HTML出力を閉じる（フッタ部分の表示）
sub_print_close_html(\$q);

}       # 利用変数がグローバル化しないように囲う（ここまで）

exit;

# htmlを開始する（HTML構文を開始して、ヘッダを表示する）
sub sub_print_start_html{

    my $q_ref = shift;      # CGIオブジェクト

    print($$q_ref->header(-type=>'text/html', -charset=>'utf-8'));
    print($$q_ref->start_html(-title=>"GoogleカレンダーにiCalファイルをインポートする",
                    -dtd=>['-//W3C//DTD XHTML 1.0 Transitional//EN','http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'],
                    -lang=>'ja-JP',
                    -style=>{'src'=>'style.css'}));

# ヘッダの表示
print << '_EOT';
<div style="height:100px; width:100%; padding:0px; margin:0px;">
<p><span style="margin:0px 20px; font-size:30px; font-weight:lighter;">iCal-Import</span><span style="margin:0px 0px; font-size:25px; font-weight:lighter; color:lightgray;">Google Calendar iCal import</span></p>
</div>
_EOT

    # 左ペイン（メニュー）の表示
    print("<div id=\"main_content_left\">\n".
            "<h2>System</h2>\n");
    {
        my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime(time);
        printf("<p>%04d/%02d/%02d %02d:%02d:%02d</p>\n", $year+1900, $mon+1, $mday, $hour, $min, $sec);
    }
    print("<h2>Menu</h2>\n".
            "<ul>\n".
            "<li><a href=\"".$str_this_script."\">Home</a></li>\n".
            "</ul>\n".
            "</div> <!-- id=\"main_content_left\" -->\n");

    # 右ペイン（主要表示部分）の表示
    print("<div id=\"main_content_right\">\n");

}

# htmlを閉じる（フッタ部分を表示して、HTML構文を閉じる）
sub sub_print_close_html{
        my $q_ref = shift;      # CGIオブジェクトへのリファレンス

print << '_EOT_FOOTER';
<p>&nbsp;</p>
</div>  <!-- id="main_content_right" -->
<p>&nbsp;</p>
<div class="clear"></div>
<div id="footer">
<p><a href="http://oasis.halfmoon.jp/software/">iCal-Import</a> version 0.1 &nbsp;&nbsp; GNU GPL free software</p>
</div>  <!-- id="footer" -->
_EOT_FOOTER

        print $$q_ref->end_html;
}

# ホーム画面（DB内のデータ数を表示）
sub sub_disp_home{
    print("<h1>iCal file upload (iCalファイルのアップロード)</h1>\n".
            "<p>iCalファイルをGoogleカレンダーにインポートします</p>\n".
            "<p>&nbsp;</p>\n".
            "<form method=\"post\" action=\"$str_this_script\" enctype=\"multipart/form-data\">\n".
            "<table>\n".
            "<tr><td>iCalファイル</td><td><input type=\"file\" name=\"uploadfile\" value=\"\" size=\"20\" /><br /></td></tr>\n".
            "<tr><td></td><td><input type=\"checkbox\" name=\"switch_n\" value=\"enable\" checked=\"checked\" />現在より未来の予定のみをインポートする</td></tr>\n".
            "<tr><td></td><td><input type=\"checkbox\" name=\"switch_s\" value=\"enable\" />メッセージを表示しない</td></tr>\n".
            "<tr><td></td><td><input type=\"checkbox\" name=\"switch_v\" value=\"enable\" />詳細なメッセージを表示する</td></tr>\n".
            "<tr><td>ユーザ名</td><td><input type=\"text\" name=\"user\" value=\"\" size=\"30\" />（前回の値を使う場合は空欄）</td></tr>\n".
            "<tr><td>パスワード</td><td><input type=\"password\" name=\"pass\" value=\"\" size=\"30\" />（前回の値を使う場合は空欄）</td></tr>\n".
            "<tr><td></td><td><input type=\"checkbox\" name=\"switch_auth_nosave\" value=\"enable\" />今回の認証データ（ユーザ名・パスワード）を保存しない</td></tr>\n".
            "</table>\n".
            "<input type=\"submit\" value=\"アップロード\" />\n".
            "</form>\n");
}

# エラー終了時に呼ばれるsub
# sub_error_exit('message');
# sub_error_exit('message', \$q);   # HTML構文を始める場合
sub sub_error_exit{
    my $str = shift;    # 出力する文字列
    my $q_ref = shift;  # CGIオブジェクトへのリファレンス：HTML構文を始める場合のみ

    # HTML構文を始める
    if(defined($q_ref)){
        sub_print_start_html($q_ref);
    }
    
    print("<p class=\"error\">".(defined($str)?$str:'error')."</p>\n");
    sub_print_close_html($q_ref);
    exit;
}


# CSVファイルをアップロード
sub sub_upload_ical{
    my $q_ref = shift;

    my $flag_n = 0;
    my $flag_s = 0;
    my $flag_v = 0;
    my $flag_auth_nosave = 0;
    my $str_user = "";
    my $str_password = "";

    if(defined($$q_ref->param('switch_n')) and $$q_ref->param('switch_n') eq "enable"){
        $flag_n = 1;
    }
    if(defined($$q_ref->param('switch_s')) and $$q_ref->param('switch_s') eq "enable"){
        $flag_s = 1;
    }
    if(defined($$q_ref->param('switch_v')) and $$q_ref->param('switch_v') eq "enable"){
        $flag_v = 1;
    }
    if(defined($$q_ref->param('switch_auth_nosave')) and $$q_ref->param('switch_auth_nosave') eq "enable"){
        $flag_auth_nosave = 1;
    }
    if(defined($$q_ref->param('user')) and $$q_ref->param('user') ne ""){
        $str_user = $$q_ref->param('user');
    }
    if(defined($$q_ref->param('pass')) and $$q_ref->param('pass') ne ""){
        $str_password = $$q_ref->param('pass');
    }

    print("<h1>Upload iCal datafile (iCalファイルのアップロード)</h1>\n".
        "<p>アップロードファイルを一時保存中 ...</p>\n");
    my $str_filename = $$q_ref->param('uploadfile');

    my $fh = $$q_ref->upload('uploadfile');
    my $str_temp_filepath = $$q_ref->tmpFileName($fh);

    print("<p>ファイルアップロード処理中 ...(".$str_temp_filepath.")</p>\n");

    File::Copy::move($str_temp_filepath, $str_filepath_ical_tmp) or sub_error_exit("Error : 一時ファイル ".$str_filepath_ical_tmp." の移動処理失敗");

    close($fh);

    unless( -f $str_filepath_ical_tmp ){ sub_error_exit("Error : 一時ファイル ".$str_filepath_ical_tmp." の存在が検知できない"); }


    print("<p>アップロード完了</p>\n");

    my $str_cmdline = "python import-ical-to-calen-cmd.py";
    if($flag_n){ $str_cmdline .= ' -n'; }
    if($flag_s){ $str_cmdline .= ' -s'; }
    if($flag_v){ $str_cmdline .= ' -v'; }
    if($flag_auth_nosave){ $str_cmdline .= ' -auth_nosave'; }
    if($str_user ne ""){ $str_cmdline .= (' -u '.$str_user); }
    if($str_password ne ""){ $str_cmdline .= (' -p '.$str_password); }
    $str_cmdline .= " data/temp.ics 2>&1";

    print("<pre>\n");
#    system("python import-ical-to-calen.py data/temp.ics 2>&1");
    my @result = `$str_cmdline`;
#    my @result = `python import-ical-to-calen-cmd.py -n data/temp.ics 2>&1`;
    print "program done\n";
    open(FH, ">test.txt");
    binmode(FH, ":utf8");
    binmode(STDOUT, ":bytes");

    foreach my $line (@result){
        print($line . "\n");
        print(FH $line);
    }
    close(FH);

    print("</pre>\n");
    
    unlink($str_filepath_ical_tmp);

    
}
