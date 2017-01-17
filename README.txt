ms_access_logger

1. 概要

Windows のログイン、ログアウト時にアクセスログを記録します
アクセスログのユーザは画面に表示されるテキストボックスに入力したユーザとなります
＃そのため、入力間違いとかは防げません。。。

ログは C:\Windows\access.log に記録されます
[date] [LOGIN|LOGOUT] [UserName]
みたいなフォーマットで記録されます


2. インストール

2.1 スクリプトの配置

logon.vbs     // Windows ログオン時に実行されるスクリプト
logging.vbs   // ログをファイルに書くための関数
logout.vbs    // Windows ログアウト時に実行されるスクリプト

上記３ファイルを C:\Windows 配下におきます。

2.2. スクリプトの登録

Windows は、ログインやログアウトのイベント発生時に実行するスクリプトを登録できます。
Windows のユーザログオンスクリプトを以下の手順に従って登録します
https://technet.microsoft.com/ja-jp/library/cc770908(v=ws.11).aspx

ローカルグループポリシーエディターの開き方
https://technet.microsoft.com/ja-jp/library/cc731745(v=ws.11).aspx


