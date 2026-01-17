@echo off
echo ☆☆☆☆☆☆釧路支店のコピーを開始します
robocopy "\\192.168.1.205\◆ 釧路 ◆" "D:\◆釧路支店◆" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆函館支店のコピーを開始します
robocopy "\\192.168.1.205\◆ 函館 ◆" "D:\◆函館支店◆" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆役員専用のコピーを開始します
robocopy "\\192.168.1.205\☆役員専用☆" "D:\☆本社☆\☆役員専用☆" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆レンタル事業部のコピーを開始します
robocopy "\\192.168.1.205\★ﾚﾝﾀﾙ事業本部" "D:\☆本社☆\★レンタル事業部" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆★規定のコピーを開始します
robocopy "\\192.168.1.205\★規定" "D:\☆本社☆\★規定" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆★給与･賞与のコピーを開始します
robocopy "\\192.168.1.205\★給与･賞与" "D:\☆本社☆\★給与･賞与" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆★掲示板のコピーを開始します
robocopy "\\192.168.1.205\★掲示板" "D:\☆本社☆\★掲示板" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆LANのコピーを開始します
robocopy "\\192.168.1.205\★本社lan" "D:\☆本社☆\★lan" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆scanのコピーを開始します
robocopy "\\192.168.1.205\★本社scan" "D:\☆本社☆\★scan" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆営業部のコピーを開始します
robocopy "\\192.168.1.205\★本社営業部" "D:\☆本社☆\★営業部" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆経理部のコピーを開始します
robocopy "\\192.168.1.205\★本社経理部" "D:\☆本社☆\★経理部" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆生産本部のコピーを開始します
robocopy "\\192.168.1.205\★本社生産本部" "D:\☆本社☆\★生産本部" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆営業月報のコピーを開始します
robocopy "\\192.168.1.205\営業月報" "D:\△共有△\営業月報" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆営業部共有のコピーを開始します
robocopy "\\192.168.1.205\営業部共有" "D:\△共有△\営業部共有" /mir /e /z /fft /xa:h /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" /xd ".trashbox-2026.01.13schedule_for_deletion_by_honsya_natori"
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆過去通達文書のコピーを開始します
robocopy "\\192.168.1.205\過去通達文書" "D:\△共有△\過去通達文書" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆概算損益のコピーを開始します
robocopy "\\192.168.1.205\概算損益\概算損益" "D:\△共有△\概算損益" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆管理本部共有のコピーを開始します
robocopy "\\192.168.1.205\管理本部共有" "D:\△共有△\管理本部共有" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆経理部共有のコピーを開始します
robocopy "\\192.168.1.205\経理部共有" "D:\△共有△\経理部共有" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆公開文書のコピーを開始します
robocopy "\\192.168.1.205\公開文書" "D:\△共有△\公開文書" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆工場月報のコピーを開始します
robocopy "\\192.168.1.205\工場月報" "D:\△共有△\工場月報" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆生産課共有のコピーを開始します
robocopy "\\192.168.1.205\生産課共有" "D:\△共有△\生産課共有" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆相談室共有のコピーを開始します
robocopy "\\192.168.1.205\相談室共有" "D:\△共有△\相談室共有" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆総務部共有のコピーを開始します
robocopy "\\192.168.1.205\総務部共有" "D:\△共有△\総務部共有" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆電子帳簿保存法のコピーを開始します
robocopy "\\192.168.1.205\電子帳簿保存法" "D:\△共有△\電子帳簿保存法" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆年計資料のコピーを開始します
robocopy "\\192.168.1.205\年計資料" "D:\△共有△\年計資料" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆売上速報のコピーを開始します
robocopy "\\192.168.1.205\売上速報" "D:\△共有△\売上速報" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆部門別損益のコピーを開始します
robocopy "\\192.168.1.205\部門別損益" "D:\△共有△\部門別損益" /mir /e /z /fft /ipg:75 /r:1 /w:2 /v /fp /np /ndl /tee /log+:"C:\robocopy_log\2026.01.17_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo ☆☆☆☆☆☆全ての工程が完了しました。
GOTO :EOF

:ERROR
echo ☆☆☆☆☆☆robocopy が中断または失敗しました。処理を終了します。
pause
exit /b 1