'###### IPアドレスを固定にするスクリプト ######
Option Explicit

Dim WMI, OS, Value, Shell

do while WScript.Arguments.Count = 0 and WScript.Version >= 5.7
    '##### WScript5.7 または Vista 以上かをチェック
    Set WMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set OS = WMI.ExecQuery("SELECT *FROM Win32_OperatingSystem")
    For Each Value in OS
        if left(Value.Version, 3) < 6.0 then exit do
    Next

    '##### 管理者権限で実行
    Set Shell = CreateObject("Shell.Application")
    Shell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ uac", "", "runas"
    WScript.Quit
loop

Dim objShell, iName, i, WSHShell, networks, Item, wShell
Dim oWshShell, owShell, oExec
Set objShell = CreateObject("Shell.Application")


'##### メニュー選択
Dim d
Dim menu

Const Title = "メニュー"

Set d = CreateObject("htmlfile")
'##### ポップアップ用割り込み処理(メニューのポップアップを最前面に表示するための前段処理）
d.parentWindow.setTimeout GetRef("SetWindow"),100

'##### メニュー画面表示
menu = InputBox("メニューを選択してください。" & vbCrLf & vbCrLf & "1.IPアドレス確認" & vbCrLf & vbCrLf & "2.IPアドレス変更" & vbCrLf & vbLf & "3.DHCP有効化",Title,2)

'##### ハンドル取得
Sub SetWindow()
    Dim Application
    Dim hwnd

    Set Application = CreateObject("Excel.Application")

    hwnd=Application.ExecuteExcel4Macro("CALL(""user32"",""FindWindowA"",""JJC"",0,""" & Title & """)")

    '##### ハンドルが取得できたらSetWindowPosでメニュー画面を最前面に表示
    If hwnd <> 0 then
        Application.ExecuteExcel4Macro("CALL(""user32"",""SetWindowPos"",""JJJJJJJJ""," & hwnd & ",-1,0,0,0,0,3)")
    Else
        MsgBox "ハンドルが取得できませんでした。"
    End If

    Set Application = Nothing
End Sub

'##### キャンセルボタン押下
If IsEmpty(menu) Then
    MsgBox "キャンセルしました。"
    WScript.Quit
End If

'##### メニューの入力チェック
Dim loopflag
loopflag = 1

Do while loopflag = 1
    '##### 値の入力なし
    If Len(menu) = 0 Then
        MsgBox "値が入力されていません。"
        menu = InputBox("メニューを選択してください。" & vbCrLf & vbCrLf & "1.IPアドレス確認" & vbCrLf & vbCrLf & "2.IPアドレス変更" & vbCrLf & vbLf & "3.DHCP有効化",Title,2)
        If IsEmpty(menu) Then
            MsgBox "キャンセルしました。"
            WScript.Quit
        End If

    '##### 数値以外の入力
    Else If Not IsNumeric(menu) Then
        MsgBox "番号を入力してください。"
        menu = InputBox("メニューを選択してください。" & vbCrLf & vbCrLf & "1.IPアドレス確認" & vbCrLf & vbCrLf & "2.IPアドレス変更" & vbCrLf & vbLf & "3.DHCP有効化",Title,2)
        If IsEmpty(menu) Then
            MsgBox "キャンセルしました。"
            WScript.Quit
        End If

    '##### 数値（範囲外）の入力
    Else If CDbl(menu) > 5 Or CDbl(menu) < 1 Then
        MsgBox "表示された範囲内の番号を入力してください。"
        menu = InputBox("メニューを選択してください。" & vbCrLf & vbCrLf & "1.IPアドレス確認" & vbCrLf & vbCrLf & "2.IPアドレス変更" & vbCrLf & vbLf & "3.DHCP有効化",Title,2)
        If IsEmpty(menu) Then
                MsgBox "キャンセルしました。"
                WScript.Quit
        End If
    '##### 上記以外（正常値）
    Else
        loopflag = 0
        Exit Do
    End If
    End If
    End If
Loop

'##### ネットワークカード情報取得
Set networks = objShell.Namespace(&H31&)'NETWORK_CONNECTIONS

'##### NIC選択ダイアログに表示する As String

iName = ""

'##### ネットワークカード名の配列

Dim iNameAry()

ReDim iNameAry(networks.Items.Count)

i=0

For Each Item in networks.Items

'##### ネットワークアダプタ名を取得

iName = iName & vbCrLf & i & ": """ & Item.Name & """"

iNameAry(i) = """" & Item.Name & """"

i=i+1

Next

'##### NICが複数ある場合の変数定義
Dim nicSelect

'##### 現行IPアドレス、サブネットマスク取得
Dim qu,cl,swbe,service,OldIP,OldMsk
Set swbe = WScript.CreateObject("WbemScripting.SWbemLocator")
Set service = swbe.ConnectServer
Set qu = service.ExecQuery("Select * From Win32_NetworkAdapterConfiguration where IPEnabled='true'")
For Each cl In qu
    OldIP = cl.IPAddress(0)
    OldMsk = cl.IPSubnet(0)
Next

'##### メニュー別処理
Select Case menu
    '##### IPアドレス確認
    Case    1
        '##### IPアドレス表示
        Set WSHShell = WScript.CreateObject("WScript.Shell")
        Set oExec = WshShell.Exec("netsh interface ipv4 show config")
        WScript.echo oExec.StdOut.ReadAll

    '##### IPアドレス変更
    Case    2
        If networks.Items.Count > 1 Then
            '##### メニュー表示
            nicSelect = InputBox("設定するネットワークカードの数字を入力してOK押下選んでください。(0～" & networks.Items.Count-1 & "の値)" & vbCrLf & iName,Title,1)
            '##### キャンセルボタン押下
            If IsEmpty(nicSelect ) Then
                MsgBox "キャンセルしました。"
                WScript.Quit
            End If
            
            '##### 入力チェック
            loopflag = 1
            Do while loopflag = 1
                '##### 値の入力なし
                If Len(nicSelect) = 0 Then
                    MsgBox "値が入力されていません。"
                    nicSelect = InputBox("設定するネットワークカードの数字を入力してOK押下選んでください。(0～" & networks.Items.Count-1 & "の値)" & vbCrLf & iName,Title,1)
                    If IsEmpty(nicSelect) Then
                        MsgBox "キャンセルしました。"
                        WScript.Quit
                    End If
                '##### 数値以外の入力
                Else If Not IsNumeric(nicSelect) Then
                    MsgBox "番号を入力してください。"
                    nicSelect = InputBox("設定するネットワークカードの数字を入力してOK押下選んでください。(0～" & networks.Items.Count-1 & "の値)" & vbCrLf & iName,Title,1)
                    If IsEmpty(nicSelect) Then
                        MsgBox "キャンセルしました。"
                        WScript.Quit
                    End If
                '##### 数値（範囲外）の入力
                ElseIf CDbl(nicSelect) > (networks.Items.Count-1) Or nicSelect*1000 Mod 1000 <> 0 Or CDbl(nicSelect) < 0 Then
                    MsgBox "表示された範囲内の番号を入力してください。"
                    nicSelect = InputBox("設定するネットワークカードの数字を入力してOK押下選んでください。(0～" & networks.Items.Count-1 & "の値)" & vbCrLf & iName,Title,1)
                    If IsEmpty(nicSelect) Then
                        MsgBox "キャンセルしました。"
                        WScript.Quit
                    End If
                '##### 上記以外（正常値）
                Else
                    loopflag = 0
                    Exit Do
                End If
                End If
            Loop
        End If

        '##### ネットワーク設定入力
        Const Title1 = "IPアドレス"
        Const Title2 = "サブネットマスク"
        Const Title3 = "デフォルトゲートウェイ"
        Const Title4 = "DNSサーバー1"
        Const Title5 = "DNSサーバー2"
        
        Dim NewIP,NewMsk,NewGWY,DnsSv1,DnsSv2
        
        '##### IPアドレス設定
        loopflag = 1
        Do while loopflag = 1
            NewIP = InputBox("IPアドレスを入力してください。" & vbCrLf & "例：10.255.8.101",Title1,"10.255.8.101")
            '##### キャンセル処理
            If IsEmpty(NewIP ) Then
                MsgBox "キャンセルしました。"
                WScript.Quit
            End If
            '##### 値の入力なし
            If Len(NewIP) = 0 Then
                MsgBox "値が入力されていません。"
                NewIP = InputBox("IPアドレスを入力してください。" & vbCrLf & "例：10.255.8.101",Title1,"10.255.8.101")
                If IsEmpty(NewIP) Then
                    MsgBox "キャンセルしました。"
                    WScript.Quit
                End If
            '##### 上記以外（正常値）
            Else
                loopflag = 0
                Exit Do
            End If
        Loop
        
        '##### サブネットマスク設定
        loopflag = 1
        Do while loopflag = 1
            NewMsk = InputBox("サブネットマスクを入力してください。" & vbCrLf & "例：255.255.255.0",Title2,"255.255.255.0")
            '##### キャンセル処理
            If IsEmpty(NewMsk ) Then
                MsgBox "キャンセルしました。"
                WScript.Quit
            End If
            '##### 値の入力なし
            If Len(NewMsk) = 0 Then
                MsgBox "値が入力されていません。"
                NewMsk = InputBox("サブネットマスクを入力してください。" & vbCrLf & "例：255.255.255.0",Title2,"255.255.255.0")
                If IsEmpty(NewMsk) Then
                    MsgBox "キャンセルしました。"
                    WScript.Quit
                End If
            '##### 上記以外（正常値）
            Else
                loopflag = 0
                Exit Do
            End If
        Loop
        
        '##### デフォルトゲートウェイ設定
        loopflag = 1
        Do while loopflag = 1
            NewGWY = InputBox("デフォルトゲートウェイを入力してください。" & vbCrLf & "例：10.255.8.254",Title2,"10.255.8.254")
            '##### キャンセル処理
            If IsEmpty(NewGWY ) Then
                MsgBox "キャンセルしました。"
                WScript.Quit
            End If
            '##### 値の入力なし
            If Len(NewGWY) = 0 Then
                MsgBox "値が入力されていません。"
                NewGWY = InputBox("デフォルトゲートウェイを入力してください。" & vbCrLf & "例：10.255.8.254",Title2,"10.255.8.254")
                If IsEmpty(NewGWY) Then
                    MsgBox "キャンセルしました。"
                    WScript.Quit
                End If
            '##### 上記以外（正常値）
            Else
                loopflag = 0
                Exit Do
            End If
        Loop
        
        '##### DNS設定1
        DnsSv1 = InputBox("優先DNSサーバーを入力してください。" & vbCrLf & "例：8.8.8.8（設定なしでもO.K）",Title4)
            '##### キャンセル処理
            If IsEmpty(DnsSv1 ) Then
                MsgBox "キャンセルしました。"
                WScript.Quit
            End If
        
        '##### DNS設定2
        DnsSv2 = InputBox("代替DNSサーバーを入力してください。" & vbCrLf & "例：8.8.8.8（設定なしでもO.K）",Title5)
            '##### キャンセル処理
            If IsEmpty(DnsSv2 ) Then
                MsgBox "キャンセルしました。"
                WScript.Quit
            End If
        
        '##### 設定変更
        Set WSHShell = WScript.CreateObject("WScript.Shell")
        WSHShell.Run "cmd.exe /c netsh interface ip set address " & iNameAry(nicSelect) & " static " & NewIP & " " & NewMsk & " " & NewGWY
        WSHShell.Run "cmd.exe /c netsh interface ip set address """ & iNameAry(nicSelect) & """ static " & NewIP & " " & NewMsk & " " & NewGWY
        WSHShell.Run "cmd.exe /c netsh interface ip set dns " & iNameAry(nicSelect) & " static " & DnsSv1 , 0, True
        WSHShell.Run "cmd.exe /c netsh interface ip add dns " & iNameAry(nicSelect) & " addr=" & DnsSv2 ,0, True
        
        '##### IPアドレス表示
        Set d=CreateObject("htmlfile")
        d.parentWindow.setTimeout GetRef("proc"),100
        Sub proc
        Set wShell = CreateObject("WScript.Shell")
        wShell.PopUp "処理中です。しばらくお待ちください。",4
        WScript.Timeout=100

        End Sub

        WScript.Sleep 4000
        Set WSHShell = WScript.CreateObject("WScript.Shell")
        Set oExec = WshShell.Exec("netsh interface ipv4 show config")
        WScript.echo oExec.StdOut.ReadAll


    '##### DHCP有効化
    Case    3
        ' ##### InputBox表示
        If networks.Items.Count > 1 Then
            nicSelect = InputBox("設定するネットワークカードの数字を入力してOK押下選んでください。(0～" & networks.Items.Count-1 & "の値)" & vbCrLf & iName,Title,1)
            ' ##### キャンセルボタン押下
            If IsEmpty(nicSelect ) Then
                MsgBox "キャンセルしました。"
                WScript.Quit
            End If
            '##### 値の入力なし
            If Not IsNumeric(nicSelect) Then
                MsgBox "値が入力されていません。もう一度実行し直してください。"
                WScript.Quit
                '##### 誤った値の入力
                ElseIf CDbl(nicSelect) > (networks.Items.Count-1) Or nicSelect*1000 Mod 1000 <> 0 Or CDbl(nicSelect) < 0 Then
                    MsgBox "設定した値に誤りがあります。もう一度実行し直してください。"
                    WScript.Quit
                Else
             End If
        End If
        Set WSHShell = WScript.CreateObject("WScript.Shell")
        WSHShell.Run "cmd.exe /c netsh interface ip set address """ & iNameAry(nicSelect) &""" dhcp"
        WSHShell.Run "cmd.exe /c netsh interface ip set dns """ & iNameAry(nicSelect) &""" dhcp"
        WSHShell.Run "cmd.exe /c netsh interface ip set address " & iNameAry(nicSelect) & " dhcp"
        WSHShell.Run "cmd.exe /c netsh interface ip set dns " & iNameAry(nicSelect) & " dhcp"
        '##### IPアドレス表示
        Set d=CreateObject("htmlfile")
        d.parentWindow.setTimeout GetRef("proc"),100
        Sub proc
        Set wShell = CreateObject("WScript.Shell")
        wShell.PopUp "処理中です。しばらくお待ちください。",4
        WScript.Timeout=100

        End Sub

        WScript.Sleep 4000
        Set WSHShell = WScript.CreateObject("WScript.Shell")
        Set oExec = WshShell.Exec("netsh interface ipv4 show config")
        WScript.echo oExec.StdOut.ReadAll

    '##### 例外処理
    Case Else
        WScript.Echo "設定した値に誤りがあります。もう一度実行し直してください。"
        WScript.Quit
End Select
