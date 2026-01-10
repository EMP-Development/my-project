@echo off
echo 営業月報のコピーを開始します
robocopy "\\192.168.1.205\営業月報" "D:\△共有△\営業月報" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\営業月報copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 営業部共有のコピーを開始します
robocopy "\\192.168.1.205\営業部共有" "D:\△共有△\営業部共有" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\営業部共有copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 過去通達文書のコピーを開始します
robocopy "\\192.168.1.205\過去通達文書" "D:\△共有△\過去通達文書" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\過去通達文書copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 概算損益のコピーを開始します
robocopy "\\192.168.1.205\概算損益" "D:\△共有△\概算損益" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\概算損益copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 管理本部共有のコピーを開始します
robocopy "\\192.168.1.205\管理本部共有" "D:\△共有△\管理本部共有" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\管理本部共有copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 経費統計のコピーを開始します
robocopy "\\192.168.1.205\経費統計" "D:\△共有△\経費統計" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\経費統計copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 経理部共有のコピーを開始します
robocopy "\\192.168.1.205\経理部共有" "D:\△共有△\経理部共有" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\経理部共有copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 公開文書のコピーを開始します
robocopy "\\192.168.1.205\公開文書" "D:\△共有△\公開文書" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\公開文書copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 工場月報のコピーを開始します
robocopy "\\192.168.1.205\工場月報" "D:\△共有△\工場月報" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\工場月報copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 釧路支店のコピーを開始します
robocopy "\\192.168.1.205\◆釧路◆" "D:\◆釧路支店◆" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\釧路支店copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 函館支店のコピーを開始します
robocopy "\\192.168.1.205\◆函館◆" "D:\◆函館支店◆" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\函館支店copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 役員専用のコピーを開始します
robocopy "\\192.168.1.205\☆役員専用☆" "D:\☆本社☆\☆役員専用☆" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\役員専用copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo レンタル事業部のコピーを開始します
robocopy "\\192.168.1.205\★ﾚﾝﾀﾙ事業本部" "D:\☆本社☆\★レンタル事業部" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\レンタル事業部copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo LANのコピーを開始します
robocopy "\\192.168.1.205\★本社lan" "D:\☆本社☆\★lan" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\lan_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo scanのコピーを開始します
robocopy "\\192.168.1.205\★本社scan" "D:\☆本社☆\★scan" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\scan_copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 営業部のコピーを開始します
robocopy "\\192.168.1.205\★本社営業部" "D:\☆本社☆\★営業部" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\本社営業部copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 経理部のコピーを開始します
robocopy "\\192.168.1.205\★本社経理部" "D:\☆本社☆\★経理部" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\本社経理部copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 生産本部のコピーを開始します
robocopy "\\192.168.1.205\★本社生産本部" "D:\☆本社☆\★生産本部" /e /z /fft /ipg:75 /r:1 /w:2 /np /ndl /tee /log+:"C:\robocopy_log\生産本部copylog.txt" 
IF ERRORLEVEL 8 GOTO :ERROR

echo 全ての工程が完了しました。
GOTO :EOF

:ERROR
echo robocopy が中断または失敗しました。処理を終了します。
pause
exit /b 1