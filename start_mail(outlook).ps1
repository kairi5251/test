###############################################################
#   平日判定処理
###############################################################

# 引数チェック
Param($DateStr = (Get-Date).ToString('yyyyMMdd'))
try {
    $CheckDate = [DateTime]::ParseExact($DateStr, 'yyyyMMdd', $null)
} catch {
    echo 'Invalid argument'
    exit 255
}
$CachePath = 'C:\Users\●●●\Documents\iMacros\Downloads\' # 内閣府提供の祝日ファイルをキャッシュするディレクトリ
$HolidayFile = Join-Path $CachePath holiday.csv   # 祝日登録ファイル名
$Limit = (Get-Date).AddMonths(-3)                 # ３ヶ月以上前は古い祝日登録ファイルとする

# 祝日登録ファイルが無い、もしくは祝日登録ファイルの更新日が古くなった場合、再取得する
if (! (Test-Path $HolidayFile) -or $Limit -gt (Get-ItemProperty $HolidayFile).LastWriteTime) {
    try {
        Invoke-WebRequest https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv -OutFile $HolidayFile
    } catch {
        echo $_.Exception.Message
        exit 250
    }
}

# バグ対応として完全一致となるよう日付に","を挿入（9/2は平日にもかかわらず9/21や9/22の祝日と判定されたため）
$CheckDate1 = (Get-Date $CheckDate).ToString('yyyy/M/d') + ","

# 祝日として登録されていれば 0 を返却して終了
if (Select-String -Quiet $CheckDate1 $HolidayFile) {
    exit 0
}

# 土日なら 0 を返却して終了
$DayOfWeek = (Get-Date $CheckDate).DayOfWeek
if ($DayOfWeek -in @('Saturday', 'Sunday')) {
    exit 0
}

# 年末年始（12月30日～1月3日）なら 0 を返却して終了
$MMDD = (Get-Date $CheckDate).ToString('MMdd')
if ($MMDD -ge '1230' -or $MMDD -le '0103') {
    exit 0
} 

# 上記いずれでもなければ平日として処理継続

################################################################
##   Outlook メール通知処理
################################################################

# Outlookプロセスの起動チェック
$TEST =Get-Process|Where-Object {$_.Name -match "OUTLOOK"}
if ($TEST -eq $null){
    $existsOutlook = $false
}else{
    $existsOutlook = $true
    }

# プロセスを起動または取得。PowerShellを管理者実行してると，普通に開いたオブジェクトを取ることが出来ない．
if ($existsOutlook) {
    $OutlookObj = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
} else {
    $OutlookObj = New-Object -ComObject Outlook.Application
    }

#件名
$result_sub = Get-Content -Encoding UTF8 C:\Users\●●●\Documents\iMacros\Downloads\sub_start_mail.txt
$result_sub += Get-Date -Format "yyyy/MM/dd"

#本文
$result = Get-Content -Encoding UTF8 C:\Users\●●●\Documents\iMacros\Downloads\body_start_mail.txt

##新規メールの作成
$mail=$OutlookObj.CreateItem(0)
$mail.Subject = $result_sub
$mail.Body = $result | out-string
$mail.To = "●宛先メールアドレス●"
$mail.Cc = "●Ccメールアドレス●"
$mail.save()
$mail.Send()
