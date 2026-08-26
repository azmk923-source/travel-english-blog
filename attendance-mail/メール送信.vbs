'==========================================================
' 13時間超過防止メール 一斉送信スクリプト
'  使い方: このファイルをダブルクリック（Outlook 使用）
'  データ : 同じフォルダの 勤務時間管理.xlsx
'==========================================================
Option Explicit

Const BOOK_NAME   = "勤務時間管理.xlsx"
Const SHEET_DATA  = "勤怠"
Const SHEET_CONF  = "設定"
Const SHEET_BODY  = "本文"

Const COL_NAME = 1
Const COL_MAIL = 2
Const COL_IN   = 3
Const COL_LIM  = 4
Const COL_SENT = 5

Dim fso, shellApp, bookPath
Set fso = CreateObject("Scripting.FileSystemObject")
bookPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), BOOK_NAME)

If Not fso.FileExists(bookPath) Then
  MsgBox BOOK_NAME & " が同じフォルダに見つかりません。", vbCritical, "勤怠メール"
  WScript.Quit 1
End If

'--- Excel を開く（画面非表示） ---
Dim xl, wb, ws, cf, bs
On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
  MsgBox "Excel を起動できませんでした。", vbCritical, "勤怠メール" : WScript.Quit 1
End If
On Error GoTo 0
xl.Visible = False
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Open(bookPath)
Set ws = wb.Worksheets(SHEET_DATA)
Set cf = wb.Worksheets(SHEET_CONF)
Set bs = wb.Worksheets(SHEET_BODY)

'--- 設定読み込み ---
Dim limitHours, startTime, amLimit, subjTpl, ccAll, clearIn
limitHours = ConfNum(cf, "上限時間", 13)
startTime  = ConfTime(cf, "送信可能開始時刻", TimeSerial(9, 0, 0))
amLimit    = ConfTime(cf, "午前打刻とみなす上限", TimeSerial(12, 0, 0))
subjTpl    = ConfStr(cf, "件名", "【勤務時間】13時間到達まで残り{remain}です")
clearIn    = (UCase(Trim(ConfStr(cf, "送信後に打刻をクリア", "ON"))) = "ON")
ccAll      = ConfStr(cf, "CC", "")

'--- 9時前は実行しない ---
If TimeValue(Now) < startTime Then
  CloseAll xl, wb
  MsgBox "送信可能時刻（" & FmtHM(startTime) & "）より前のため送信しません。" & vbCrLf & _
         "現在時刻: " & FmtHM(TimeValue(Now)), vbExclamation, "勤怠メール"
  WScript.Quit 0
End If

'--- 本文テンプレート ---
Dim bodyTpl, r, lastBody
bodyTpl = "" : lastBody = bs.Cells(bs.Rows.Count, 1).End(-4162).Row
For r = 3 To lastBody
  bodyTpl = bodyTpl & CStr(bs.Cells(r, 1).Value) & vbCrLf
Next

'--- Outlook ---
Dim ol
On Error Resume Next
Set ol = CreateObject("Outlook.Application")
If Err.Number <> 0 Then
  CloseAll xl, wb
  MsgBox "Outlook を起動できませんでした。", vbCritical, "勤怠メール" : WScript.Quit 1
End If
On Error GoTo 0

'--- 送信ループ ---
Dim lastRow, nm, ml, inT, limT, sentV, remainMin
Dim sentCnt, skipDup, skipPm, skipOver, errCnt, logTxt
sentCnt = 0 : skipDup = 0 : skipPm = 0 : skipOver = 0 : errCnt = 0 : logTxt = ""
lastRow = ws.Cells(ws.Rows.Count, COL_MAIL).End(-4162).Row

For r = 2 To lastRow
  nm = Trim(CStr(ws.Cells(r, COL_NAME).Value))
  ml = Trim(CStr(ws.Cells(r, COL_MAIL).Value))
  If nm <> "" And ml <> "" And Len(Trim(CStr(ws.Cells(r, COL_IN).Value))) > 0 Then
    inT = ToTime(ws.Cells(r, COL_IN).Value)
    sentV = ws.Cells(r, COL_SENT).Value

    If IsDate(sentV) Then
      If DateValue(CDate(sentV)) = Date Then
        skipDup = skipDup + 1
        logTxt = logTxt & "  ・" & nm & " : 本日送信済みのためスキップ" & vbCrLf
        inT = -1
      End If
    End If

    If inT <> -1 Then
      If inT >= amLimit Then
        skipPm = skipPm + 1
        logTxt = logTxt & "  ・" & nm & " : 午後出勤のため対象外" & vbCrLf
      Else
        limT = inT + CDbl(limitHours) / 24
        remainMin = CLng((limT - TimeValue(Now)) * 24 * 60)

        Dim subj, body, remainTxt
        If remainMin <= 0 Then
          remainTxt = "0時間0分（" & HM(Abs(remainMin)) & "超過）"
          skipOver = skipOver + 1
        Else
          remainTxt = HM(remainMin)
        End If

        subj = Repl(subjTpl, nm, inT, limT, remainTxt, limitHours)
        body = Repl(bodyTpl, nm, inT, limT, remainTxt, limitHours)

        On Error Resume Next
        Dim mail
        Set mail = ol.CreateItem(0)
        mail.To = ml
        If ccAll <> "" Then mail.CC = ccAll
        mail.Subject = subj
        mail.Body = body
        mail.Send
        If Err.Number <> 0 Then
          errCnt = errCnt + 1
          logTxt = logTxt & "  ×" & nm & " : 送信失敗 (" & Err.Description & ")" & vbCrLf
          Err.Clear
        Else
          sentCnt = sentCnt + 1
          ws.Cells(r, COL_SENT).Value = Date   '重複送信防止
          If clearIn Then ws.Cells(r, COL_IN).ClearContents  '翌日の誤送信防止
          logTxt = logTxt & "  ○" & nm & " : 残り " & remainTxt & vbCrLf
        End If
        On Error GoTo 0
        Set mail = Nothing
      End If
    End If
  End If
Next

wb.Save
CloseAll xl, wb

MsgBox "送信完了" & vbCrLf & vbCrLf & _
       "送信 " & sentCnt & " 件 / 送信済みスキップ " & skipDup & " 件 / 午後出勤 " & skipPm & " 件" & vbCrLf & _
       "うち13時間超過 " & skipOver & " 件 / 失敗 " & errCnt & " 件" & vbCrLf & vbCrLf & _
       logTxt, vbInformation, "勤怠メール"

'==================== 関数 ====================

' 分 → 「○時間○分」
Function HM(totalMin)
  Dim m : m = CLng(totalMin)
  HM = CStr(m \ 60) & "時間" & CStr(m Mod 60) & "分"
End Function

Function FmtHM(t)
  FmtHM = Right("0" & Hour(t), 2) & ":" & Right("0" & Minute(t), 2)
End Function

' セル値（時刻 or シリアル値）→ 1日を1とする小数
Function ToTime(v)
  If IsDate(v) Then
    ToTime = CDbl(TimeValue(CDate(v)))
  Else
    ToTime = CDbl(v) - Int(CDbl(v))
  End If
End Function

Function Repl(tpl, nm, inT, limT, remainTxt, lh)
  Dim s : s = tpl
  s = Replace(s, "{name}", nm)
  s = Replace(s, "{in}", FmtHM(CDate(inT)))
  s = Replace(s, "{limit}", FmtHM(CDate(limT)))
  s = Replace(s, "{remain}", remainTxt)
  s = Replace(s, "{hours}", CStr(lh))
  s = Replace(s, "{date}", Year(Date) & "/" & Right("0" & Month(Date), 2) & "/" & Right("0" & Day(Date), 2))
  Repl = s
End Function

Function ConfRow(cf, key)
  Dim i, last
  last = cf.Cells(cf.Rows.Count, 1).End(-4162).Row
  ConfRow = 0
  For i = 2 To last
    If Trim(CStr(cf.Cells(i, 1).Value)) = key Then ConfRow = i : Exit For
  Next
End Function

Function ConfStr(cf, key, dflt)
  Dim i, v
  i = ConfRow(cf, key)
  If i = 0 Then
    v = dflt
  Else
    v = CStr(cf.Cells(i, 2).Value)
    If Trim(v) = "" And key <> "CC" Then v = dflt
  End If
  ConfStr = v
End Function

Function ConfNum(cf, key, dflt)
  Dim i : i = ConfRow(cf, key)
  If i = 0 Then ConfNum = dflt Else ConfNum = CDbl(cf.Cells(i, 2).Value)
End Function

Function ConfTime(cf, key, dflt)
  Dim i, v : i = ConfRow(cf, key)
  If i = 0 Then ConfTime = dflt : Exit Function
  v = cf.Cells(i, 2).Value
  If v = "" Then ConfTime = dflt Else ConfTime = ToTime(v)
End Function

Sub CloseAll(xl, wb)
  On Error Resume Next
  wb.Close True
  xl.Quit
  Set wb = Nothing : Set xl = Nothing
End Sub
