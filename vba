Option Explicit

'==================================================
' シート名
'==================================================
Private Const SHEET_SCREEN As String = "QR読込画面"
Private Const SHEET_LOG As String = "使用記録ログ"
Private Const SHEET_MEDIA As String = "媒体マスタ"
Private Const SHEET_USER As String = "利用者マスタ"
Private Const SHEET_PERMIT As String = "媒体利用許可マスタ"
Private Const SHEET_ERROR As String = "異常ログ"

'==================================================
' テーブル名
'==================================================
Private Const TBL_LOG As String = "tblLog"
Private Const TBL_MEDIA As String = "tblMedia"
Private Const TBL_USER As String = "tblUser"
Private Const TBL_PERMIT As String = "tblPermit"
Private Const TBL_ERROR As String = "tblError"

'==================================================
' QR読込画面のセル
' 画像レイアウト版
'==================================================
Private Const CELL_STATUS As String = "B3"          '現在状態エリア
Private Const CELL_RESULT_ICON As String = "B7"     '結果アイコン
Private Const CELL_RESULT As String = "C7"          '処理結果
Private Const CELL_RESULT_DETAIL As String = "C9"   '処理結果詳細
Private Const CELL_INPUT As String = "C13"          'QR読込欄
Private Const CELL_MEDIA_NAME As String = "J25"     '読込済媒体
Private Const CELL_USER_NAME As String = "J26"      '読込済利用者
Private Const CELL_LAST_TIME As String = "J1"       '最終更新

'処理中データ保存用
Private Const CELL_PENDING_MEDIA_QR As String = "K3"
Private Const CELL_PENDING_MEDIA_NAME As String = "K4"

'貸出中一覧の表示位置
Private Const LOAN_LIST_TITLE_ROW As Long = 17
Private Const LOAN_LIST_HEADER_ROW As Long = 18
Private Const LOAN_LIST_FIRST_ROW As Long = 19
Private Const LOAN_LIST_MAX_ROW As Long = 34

'==================================================
' 初期表示
'==================================================
Public Sub QR_InitializeScreen()

    On Error GoTo ErrHandler

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    QR_BuildDashboardLayout

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SCREEN)

    With ws
        .Range(CELL_INPUT).MergeArea.ClearContents
        .Range(CELL_INPUT).NumberFormat = "@"

        .Range(CELL_PENDING_MEDIA_QR).ClearContents
        .Range(CELL_PENDING_MEDIA_NAME).ClearContents
        .Range(CELL_MEDIA_NAME).ClearContents
        .Range(CELL_USER_NAME).ClearContents

        .Range(CELL_STATUS).MergeArea.Value = "媒体QRを読み込んでください"
        .Range(CELL_RESULT_ICON).MergeArea.Value = "●"
        .Range(CELL_RESULT).MergeArea.Value = "待機中"
        .Range(CELL_RESULT_DETAIL).MergeArea.Value = "QRを読み込んでください"
        .Range(CELL_LAST_TIME).Value = "最終更新：" & Format(Now, "yyyy/mm/dd hh:mm:ss")
        .Range("J27").Value = "媒体QR待ち"
        .Range("J28").Value = "0"
    End With

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    QR_UpdateLoanList
    QR_ApplyScreenProtection
    QR_FocusInput

    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "初期表示でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation

End Sub

'==================================================
' QR読込画面レイアウト作成
'==================================================
Public Sub QR_BuildDashboardLayout()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SCREEN)

    Application.ScreenUpdating = False

    On Error Resume Next
    ws.Unprotect Password:="qr"
    On Error GoTo ErrHandler

    ws.Cells.UnMerge
    ws.Cells.Clear

    ws.Cells.Font.Name = "Meiryo UI"
    ws.Cells.Font.Size = 11

    '列幅
    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 13
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 20
    ws.Columns("G").ColumnWidth = 13
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 16
    ws.Columns("K:L").Hidden = True

    '行高さ
    ws.Rows("1").RowHeight = 30
    ws.Rows("3:5").RowHeight = 32
    ws.Rows("7:10").RowHeight = 30
    ws.Rows("13:14").RowHeight = 28
    ws.Rows("17:34").RowHeight = 22
    ws.Rows("36:40").RowHeight = 22

    'ヘッダー
    With ws.Range("A1:H1")
        .Merge
        .Value = "QR読込画面（媒体貸出管理システム）"
        .Interior.Color = RGB(0, 97, 45)
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 18
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("I1:J1")
        .Merge
        .Value = "最終更新："
        .Interior.Color = RGB(0, 97, 45)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    '現在状態エリア
    With ws.Range("B3:J5")
        .Merge
        .Value = "媒体QRを読み込んでください"
        .Interior.Color = RGB(255, 235, 238)
        .Font.Bold = True
        .Font.Size = 28
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(220, 80, 80)
    End With

    '処理結果アイコン
    With ws.Range("B7:B10")
        .Merge
        .Value = "●"
        .Interior.Color = RGB(226, 239, 218)
        .Font.Color = RGB(0, 97, 0)
        .Font.Bold = True
        .Font.Size = 42
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(112, 173, 71)
    End With

    '処理結果
    With ws.Range("C7:J8")
        .Merge
        .Value = "待機中"
        .Interior.Color = RGB(226, 239, 218)
        .Font.Color = RGB(0, 97, 0)
        .Font.Bold = True
        .Font.Size = 26
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(112, 173, 71)
    End With

    '処理結果詳細
    With ws.Range("C9:J10")
        .Merge
        .Value = "QRを読み込んでください"
        .Interior.Color = RGB(226, 239, 218)
        .Font.Color = RGB(0, 97, 0)
        .Font.Size = 16
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(112, 173, 71)
    End With

    'QR入力エリア
    With ws.Range("B13:B14")
        .Merge
        .Value = "QR 読込入力欄"
        .Interior.Color = RGB(0, 92, 175)
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("C13:H14")
        .Merge
        .Value = ""
        .Interior.Color = RGB(242, 248, 255)
        .Font.Color = RGB(0, 92, 175)
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 92, 175)
    End With

    With ws.Range("I13:J14")
        .Merge
        .Value = "取消（SYS:CANCEL）"
        .Interior.Color = RGB(242, 236, 220)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 170, 150)
    End With

    '貸出中一覧タイトル
    With ws.Range("A17:H17")
        .Merge
        .Value = "現在の貸出中一覧（格納時間が空欄のもの）"
        .Interior.Color = RGB(0, 97, 45)
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    '貸出中一覧ヘッダー
    ws.Range("A18").Value = "No"
    ws.Range("B18").Value = "登録番号"
    ws.Range("C18").Value = "媒体種別"
    ws.Range("D18").Value = "媒体QR"
    ws.Range("E18").Value = "使用者"
    ws.Range("F18").Value = "取出時間"
    ws.Range("G18").Value = "経過時間"
    ws.Range("H18").Value = "備考"

    With ws.Range("A18:H18")
        .Interior.Color = RGB(226, 239, 218)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 180, 180)
    End With

    With ws.Range("A19:H34")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(220, 220, 220)
    End With

    'クイック操作
    With ws.Range("I17:J17")
        .Merge
        .Value = "クイック操作（QR読取でOK）"
        .Interior.Color = RGB(0, 92, 175)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("I19:J20")
        .Merge
        .Value = "リセット（SYS:RESET）"
        .Interior.Color = RGB(242, 236, 220)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    With ws.Range("I22:J23")
        .Merge
        .Value = "取消（SYS:CANCEL）"
        .Interior.Color = RGB(242, 236, 220)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    'システム状態
    With ws.Range("I24:J24")
        .Merge
        .Value = "システム状態"
        .Interior.Color = RGB(112, 48, 160)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ws.Range("I25").Value = "読込済媒体"
    ws.Range("I26").Value = "読込済利用者"
    ws.Range("I27").Value = "待ち状態"
    ws.Range("I28").Value = "読込エラー回数"

    ws.Range("J25").Value = ""
    ws.Range("J26").Value = ""
    ws.Range("J27").Value = "媒体QR待ち"
    ws.Range("J28").Value = "0"

    With ws.Range("I25:J28")
        .Interior.Color = RGB(245, 245, 245)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 210)
    End With

    '使い方
    With ws.Range("A36:J40")
        .Interior.Color = RGB(255, 248, 225)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 180, 120)
    End With

    ws.Range("A36").Value = "使い方"
    ws.Range("A36").Font.Bold = True

    ws.Range("A37").Value = "1. 媒体QRを読み込む　→　未貸出なら利用者QR待ち、貸出中なら返納"
    ws.Range("A38").Value = "2. 利用者QRを読み込む　→　許可確認後、貸出記録"
    ws.Range("A39").Value = "3. 取消したい場合は SYS:CANCEL、リセットは SYS:RESET を読み込む"
    ws.Range("A40").Value = "4. 日を跨いで返納した場合、備考に期間を自動記録"

    ws.Range("A1:J40").VerticalAlignment = xlCenter

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "QR読込画面レイアウト作成でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation

End Sub

'==================================================
' QR読込欄にフォーカスを戻す
'==================================================
Public Sub QR_FocusInput()

    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SCREEN)

    ws.Activate
    ws.Range(CELL_INPUT).Select

    On Error GoTo 0

End Sub

'==================================================
' QR処理メイン
'==================================================
Public Sub QR_Process(ByVal rawQr As String)

    On Error GoTo ErrHandler

    Dim qr As String
    qr = NormalizeQR(rawQr)

    If qr = "" Then
        SetScreenMessage "媒体QRを読み込んでください", "QRが空です。", True
        LogError "読込", rawQr, "", "", "空QR", "QRの値が空でした。"
        CountReadError
        QR_FocusInput
        Exit Sub
    End If

    Select Case qr

        Case "SYS:CANCEL"
            ClearPending
            SetScreenMessage "媒体QRを読み込んでください", "処理を取消しました。", False

        Case "SYS:RESET"
            ClearPending
            SetScreenMessage "媒体QRを読み込んでください", "画面をリセットしました。", False

        Case Else

            If Left(qr, 6) = "MEDIA:" Then
                ProcessMediaQR qr

            ElseIf Left(qr, 5) = "USER:" Then
                ProcessUserQR qr

            Else
                SetScreenMessage "媒体QRを読み込んでください", "不明なQRです：" & rawQr, True
                LogError "読込", rawQr, "", "", "不明QR", "MEDIA、USER、SYSのいずれでもありません。"
                CountReadError
            End If

    End Select

    QR_UpdateLoanList
    QR_FocusInput
    Exit Sub

ErrHandler:
    SetScreenMessage "媒体QRを読み込んでください", "エラー：" & Err.Description, True
    LogError "システム", rawQr, "", "", "VBAエラー", Err.Description
    CountReadError
    QR_UpdateLoanList
    QR_FocusInput

End Sub

'==================================================
' 媒体QR処理
'==================================================
Private Sub ProcessMediaQR(ByVal mediaQR As String)

    Dim registerNo As String
    Dim mediaType As String
    Dim mediaState As String

    If Not FindMedia(mediaQR, registerNo, mediaType, mediaState) Then
        SetScreenMessage "媒体QRを読み込んでください", "未登録の媒体です：" & mediaQR, True
        LogError "読込", mediaQR, mediaQR, "", "未登録媒体", "媒体マスタに登録されていません。"
        CountReadError
        ClearPending
        Exit Sub
    End If

    If mediaState <> "使用可" Then
        SetScreenMessage "媒体QRを読み込んでください", "この媒体は使用できません：" & registerNo & " / 状態：" & mediaState, True
        LogError "貸出", mediaQR, mediaQR, "", "媒体使用不可", "媒体状態が使用可ではありません。状態：" & mediaState
        CountReadError
        ClearPending
        Exit Sub
    End If

    Dim activeRowIndex As Long
    Dim activeCount As Long

    FindActiveLogRows registerNo, activeRowIndex, activeCount

    If activeCount > 1 Then
        SetScreenMessage "管理者に連絡してください", "同じ媒体の貸出中記録が複数あります：" & registerNo, True
        LogError "返納", mediaQR, mediaQR, "", "貸出中重複", "同一媒体の格納時間空欄が複数あります。登録番号：" & registerNo
        CountReadError
        ClearPending
        Exit Sub
    End If

    If activeCount = 1 Then
        ReturnMedia mediaQR, registerNo, activeRowIndex
        ClearPending
    Else
        SetPendingMedia mediaQR, registerNo
        SetScreenMessage "利用者QRを読み込んでください", "媒体読込済：" & registerNo, False
    End If

End Sub

'==================================================
' 利用者QR処理
'==================================================
Private Sub ProcessUserQR(ByVal userQR As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SCREEN)

    Dim mediaQR As String
    Dim registerNo As String

    mediaQR = CStr(ws.Range(CELL_PENDING_MEDIA_QR).Value)
    registerNo = CStr(ws.Range(CELL_PENDING_MEDIA_NAME).Value)

    If mediaQR = "" Then
        SetScreenMessage "媒体QRを読み込んでください", "先に媒体QRを読み込んでください。", True
        LogError "読込", userQR, "", userQR, "読込順序エラー", "利用者QRが先に読み込まれました。"
        CountReadError
        Exit Sub
    End If

    Dim userName As String

    If Not FindUser(userQR, userName) Then
        SetScreenMessage "利用者QRを読み込んでください", "未登録の利用者です：" & userQR, True
        LogError "貸出", userQR, mediaQR, userQR, "未登録利用者", "利用者マスタに登録されていません。"
        CountReadError
        Exit Sub
    End If

    ws.Range(CELL_USER_NAME).Value = userName

    If Not IsPermitted(mediaQR, userQR) Then
        SetScreenMessage "媒体QRを読み込んでください", "貸出不可：" & userName & " は " & registerNo & " を使用できません。", True
        LogError "貸出", userQR, mediaQR, userQR, "権限なし", userName & " は " & registerNo & " の利用許可がありません。"
        CountReadError
        ClearPending
        Exit Sub
    End If

    Dim activeRowIndex As Long
    Dim activeCount As Long

    FindActiveLogRows registerNo, activeRowIndex, activeCount

    If activeCount > 0 Then
        SetScreenMessage "媒体QRを読み込んでください", "この媒体はすでに貸出中です：" & registerNo, True
        LogError "貸出", userQR, mediaQR, userQR, "二重貸出", "貸出中の媒体を再度貸出しようとしました。"
        CountReadError
        ClearPending
        Exit Sub
    End If

    AddLoanLog registerNo, userName

    SetScreenMessage "媒体QRを読み込んでください", "貸出完了：" & registerNo & " / " & userName, False
    ClearPending

End Sub

'==================================================
' 貸出記録を追加
'==================================================
Private Sub AddLoanLog(ByVal registerNo As String, ByVal userName As String)

    Dim lo As ListObject
    Set lo = GetTable(SHEET_LOG, TBL_LOG)

    Dim lr As ListRow
    Set lr = lo.ListRows.Add

    Dim nowTime As Date
    nowTime = Now

    With lr.Range
        .Cells(1, GetColIndex(lo, "年月日")).Value = Date
        .Cells(1, GetColIndex(lo, "登録番号")).Value = registerNo
        .Cells(1, GetColIndex(lo, "使用者")).Value = userName
        .Cells(1, GetColIndex(lo, "取出時間")).Value = nowTime
        .Cells(1, GetColIndex(lo, "格納時間")).Value = ""
        .Cells(1, GetColIndex(lo, "備考")).Value = ""
    End With

    FormatLogColumns

End Sub

'==================================================
' 返納処理
'==================================================
Private Sub ReturnMedia(ByVal mediaQR As String, ByVal registerNo As String, ByVal activeRowIndex As Long)

    Dim lo As ListObject
    Set lo = GetTable(SHEET_LOG, TBL_LOG)

    Dim cOut As Long
    Dim cIn As Long
    Dim cUser As Long
    Dim cNote As Long

    cOut = GetColIndex(lo, "取出時間")
    cIn = GetColIndex(lo, "格納時間")
    cUser = GetColIndex(lo, "使用者")
    cNote = GetColIndex(lo, "備考")

    Dim inTime As Date
    inTime = Now

    Dim outValue As Variant
    outValue = lo.DataBodyRange.Cells(activeRowIndex, cOut).Value

    lo.DataBodyRange.Cells(activeRowIndex, cIn).Value = inTime

    Dim noteText As String
    noteText = ""

    If IsDate(outValue) Then
        If DateValue(CDate(outValue)) <> DateValue(inTime) Then
            noteText = Format(CDate(outValue), "yyyy/mm/dd") & "〜" & Format(inTime, "yyyy/mm/dd")
        End If
    End If

    If noteText <> "" Then
        Dim oldNote As String
        oldNote = CStr(lo.DataBodyRange.Cells(activeRowIndex, cNote).Value)

        If oldNote = "" Then
            lo.DataBodyRange.Cells(activeRowIndex, cNote).Value = noteText
        ElseIf InStr(oldNote, noteText) = 0 Then
            lo.DataBodyRange.Cells(activeRowIndex, cNote).Value = oldNote & " / " & noteText
        End If
    End If

    FormatLogColumns

    Dim userName As String
    userName = CStr(lo.DataBodyRange.Cells(activeRowIndex, cUser).Value)

    SetScreenMessage "媒体QRを読み込んでください", "返納完了：" & registerNo & " / " & userName, False

End Sub

'==================================================
' 現在貸出中一覧をQR読込画面に表示
'==================================================
Public Sub QR_UpdateLoanList()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SCREEN)

    On Error Resume Next
    ws.Range("A" & LOAN_LIST_FIRST_ROW & ":H" & LOAN_LIST_MAX_ROW).UnMerge
    On Error GoTo ErrHandler

    ws.Range("A" & LOAN_LIST_FIRST_ROW & ":H" & LOAN_LIST_MAX_ROW).ClearContents

    With ws.Range("A" & LOAN_LIST_FIRST_ROW & ":H" & LOAN_LIST_MAX_ROW)
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(220, 220, 220)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Dim lo As ListObject
    Set lo = GetTable(SHEET_LOG, TBL_LOG)

    If lo.DataBodyRange Is Nothing Then
        ws.Range("A" & LOAN_LIST_FIRST_ROW & ":H" & LOAN_LIST_FIRST_ROW).Merge
        ws.Range("A" & LOAN_LIST_FIRST_ROW).Value = "貸出中なし"
        ws.Range("A" & LOAN_LIST_FIRST_ROW).HorizontalAlignment = xlCenter
        Exit Sub
    End If

    Dim cReg As Long
    Dim cUser As Long
    Dim cOut As Long
    Dim cIn As Long
    Dim cNote As Long

    cReg = GetColIndex(lo, "登録番号")
    cUser = GetColIndex(lo, "使用者")
    cOut = GetColIndex(lo, "取出時間")
    cIn = GetColIndex(lo, "格納時間")
    cNote = GetColIndex(lo, "備考")

    Dim i As Long
    Dim outRow As Long
    Dim no As Long

    outRow = LOAN_LIST_FIRST_ROW
    no = 1

    For i = 1 To lo.DataBodyRange.Rows.Count

        If Trim(CStr(lo.DataBodyRange.Cells(i, cIn).Value)) = "" Then

            If outRow > LOAN_LIST_MAX_ROW Then
                ws.Cells(outRow - 1, "H").Value = "ほかにも貸出中あり"
                Exit For
            End If

            Dim registerNo As String
            Dim mediaQR As String
            Dim mediaType As String

            registerNo = CStr(lo.DataBodyRange.Cells(i, cReg).Value)
            mediaQR = ""
            mediaType = ""

            Call FindMediaByRegisterNo(registerNo, mediaQR, mediaType)

            ws.Cells(outRow, "A").Value = no
            ws.Cells(outRow, "B").Value = registerNo
            ws.Cells(outRow, "C").Value = mediaType
            ws.Cells(outRow, "D").Value = mediaQR
            ws.Cells(outRow, "E").Value = lo.DataBodyRange.Cells(i, cUser).Value
            ws.Cells(outRow, "F").Value = lo.DataBodyRange.Cells(i, cOut).Value
            ws.Cells(outRow, "F").NumberFormat = "yyyy/mm/dd hh:mm:ss"

            If IsDate(lo.DataBodyRange.Cells(i, cOut).Value) Then
                ws.Cells(outRow, "G").Value = FormatElapsed(CDate(lo.DataBodyRange.Cells(i, cOut).Value), Now)
            Else
                ws.Cells(outRow, "G").Value = "不明"
            End If

            ws.Cells(outRow, "H").Value = lo.DataBodyRange.Cells(i, cNote).Value

            outRow = outRow + 1
            no = no + 1

        End If

    Next i

    If outRow = LOAN_LIST_FIRST_ROW Then
        ws.Range("A" & LOAN_LIST_FIRST_ROW & ":H" & LOAN_LIST_FIRST_ROW).Merge
        ws.Range("A" & LOAN_LIST_FIRST_ROW).Value = "貸出中なし"
        ws.Range("A" & LOAN_LIST_FIRST_ROW).HorizontalAlignment = xlCenter
    End If

    ws.Range("A" & LOAN_LIST_FIRST_ROW & ":H" & LOAN_LIST_MAX_ROW).HorizontalAlignment = xlCenter
    ws.Range("B" & LOAN_LIST_FIRST_ROW & ":H" & LOAN_LIST_MAX_ROW).HorizontalAlignment = xlLeft
    ws.Range("F" & LOAN_LIST_FIRST_ROW & ":G" & LOAN_LIST_MAX_ROW).HorizontalAlignment = xlCenter

    Exit Sub

ErrHandler:
    ws.Range("A" & LOAN_LIST_FIRST_ROW).Value = "貸出中一覧の更新でエラー：" & Err.Description

End Sub

'==================================================
' 媒体検索
'==================================================
Private Function FindMedia(ByVal mediaQR As String, ByRef registerNo As String, ByRef mediaType As String, ByRef mediaState As String) As Boolean

    Dim lo As ListObject
    Set lo = GetTable(SHEET_MEDIA, TBL_MEDIA)

    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim cQR As Long
    Dim cReg As Long
    Dim cType As Long
    Dim cState As Long

    cQR = GetColIndex(lo, "媒体QR")
    cReg = GetColIndex(lo, "登録番号")
    cType = GetColIndex(lo, "媒体種別")
    cState = GetColIndex(lo, "状態")

    Dim i As Long

    For i = 1 To lo.DataBodyRange.Rows.Count

        If NormalizeQR(lo.DataBodyRange.Cells(i, cQR).Value) = mediaQR Then
            registerNo = CStr(lo.DataBodyRange.Cells(i, cReg).Value)
            mediaType = CStr(lo.DataBodyRange.Cells(i, cType).Value)
            mediaState = CStr(lo.DataBodyRange.Cells(i, cState).Value)
            FindMedia = True
            Exit Function
        End If

    Next i

End Function

'==================================================
' 登録番号から媒体QR・媒体種別を取得
'==================================================
Private Function FindMediaByRegisterNo(ByVal registerNo As String, _
                                       ByRef mediaQR As String, _
                                       ByRef mediaType As String) As Boolean

    Dim lo As ListObject
    Set lo = GetTable(SHEET_MEDIA, TBL_MEDIA)

    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim cQR As Long
    Dim cReg As Long
    Dim cType As Long

    cQR = GetColIndex(lo, "媒体QR")
    cReg = GetColIndex(lo, "登録番号")
    cType = GetColIndex(lo, "媒体種別")

    Dim i As Long

    For i = 1 To lo.DataBodyRange.Rows.Count

        If CStr(lo.DataBodyRange.Cells(i, cReg).Value) = registerNo Then
            mediaQR = CStr(lo.DataBodyRange.Cells(i, cQR).Value)
            mediaType = CStr(lo.DataBodyRange.Cells(i, cType).Value)
            FindMediaByRegisterNo = True
            Exit Function
        End If

    Next i

End Function

'==================================================
' 利用者検索
'==================================================
Private Function FindUser(ByVal userQR As String, ByRef userName As String) As Boolean

    Dim lo As ListObject
    Set lo = GetTable(SHEET_USER, TBL_USER)

    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim cQR As Long
    Dim cName As Long

    cQR = GetColIndex(lo, "利用者QR")
    cName = GetColIndex(lo, "使用者")

    Dim i As Long

    For i = 1 To lo.DataBodyRange.Rows.Count

        If NormalizeQR(lo.DataBodyRange.Cells(i, cQR).Value) = userQR Then
            userName = CStr(lo.DataBodyRange.Cells(i, cName).Value)
            FindUser = True
            Exit Function
        End If

    Next i

End Function

'==================================================
' 利用許可確認
'==================================================
Private Function IsPermitted(ByVal mediaQR As String, ByVal userQR As String) As Boolean

    Dim lo As ListObject
    Set lo = GetTable(SHEET_PERMIT, TBL_PERMIT)

    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim cMedia As Long
    Dim cUser As Long

    cMedia = GetColIndex(lo, "媒体QR")
    cUser = GetColIndex(lo, "利用者QR")

    Dim i As Long

    For i = 1 To lo.DataBodyRange.Rows.Count

        If NormalizeQR(lo.DataBodyRange.Cells(i, cMedia).Value) = mediaQR _
           And NormalizeQR(lo.DataBodyRange.Cells(i, cUser).Value) = userQR Then

            IsPermitted = True
            Exit Function

        End If

    Next i

End Function

'==================================================
' 貸出中記録検索
'==================================================
Private Sub FindActiveLogRows(ByVal registerNo As String, ByRef activeRowIndex As Long, ByRef activeCount As Long)

    activeRowIndex = 0
    activeCount = 0

    Dim lo As ListObject
    Set lo = GetTable(SHEET_LOG, TBL_LOG)

    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim cReg As Long
    Dim cIn As Long

    cReg = GetColIndex(lo, "登録番号")
    cIn = GetColIndex(lo, "格納時間")

    Dim i As Long

    For i = lo.DataBodyRange.Rows.Count To 1 Step -1

        If CStr(lo.DataBodyRange.Cells(i, cReg).Value) = registerNo _
           And Trim(CStr(lo.DataBodyRange.Cells(i, cIn).Value)) = "" Then

            activeCount = activeCount + 1

            If activeRowIndex = 0 Then
                activeRowIndex = i
            End If

        End If

    Next i

End Sub

'==================================================
' 異常ログ記録
'==================================================
Private Sub LogError(ByVal processType As String, _
                     ByVal readQR As String, _
                     ByVal mediaQR As String, _
                     ByVal userQR As String, _
                     ByVal errorClass As String, _
                     ByVal errorMessage As String)

    On Error Resume Next

    Dim lo As ListObject
    Set lo = GetTable(SHEET_ERROR, TBL_ERROR)

    Dim lr As ListRow
    Set lr = lo.ListRows.Add

    With lr.Range
        .Cells(1, GetColIndex(lo, "発生日時")).Value = Now
        .Cells(1, GetColIndex(lo, "処理区分")).Value = processType
        .Cells(1, GetColIndex(lo, "読込QR")).Value = readQR
        .Cells(1, GetColIndex(lo, "媒体QR")).Value = mediaQR
        .Cells(1, GetColIndex(lo, "利用者QR")).Value = userQR
        .Cells(1, GetColIndex(lo, "エラー分類")).Value = errorClass
        .Cells(1, GetColIndex(lo, "エラー内容")).Value = errorMessage
        .Cells(1, GetColIndex(lo, "備考")).Value = ""
    End With

    lo.ListColumns("発生日時").DataBodyRange.NumberFormat = "yyyy/mm/dd hh:mm:ss"

    On Error GoTo 0

End Sub

'==================================================
' 読込エラー回数カウント
'==================================================
Private Sub CountReadError()

    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SCREEN)

    If IsNumeric(ws.Range("J28").Value) Then
        ws.Range("J28").Value = CLng(ws.Range("J28").Value) + 1
    Else
        ws.Range("J28").Value = 1
    End If

    On Error GoTo 0

End Sub

'==================================================
' 処理中媒体を保存
'==================================================
Private Sub SetPendingMedia(ByVal mediaQR As String, ByVal registerNo As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SCREEN)

    ws.Range(CELL_PENDING_MEDIA_QR).Value = mediaQR
    ws.Range(CELL_PENDING_MEDIA_NAME).Value = registerNo
    ws.Range(CELL_MEDIA_NAME).Value = registerNo
    ws.Range(CELL_USER_NAME).ClearContents

End Sub

'==================================================
' 処理中データをクリア
'==================================================
Private Sub ClearPending()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SCREEN)

    ws.Range(CELL_PENDING_MEDIA_QR).ClearContents
    ws.Range(CELL_PENDING_MEDIA_NAME).ClearContents
    ws.Range(CELL_MEDIA_NAME).ClearContents
    ws.Range(CELL_USER_NAME).ClearContents

End Sub

'==================================================
' 画面メッセージ表示
'==================================================
Private Sub SetScreenMessage(ByVal statusText As String, ByVal resultText As String, Optional ByVal isError As Boolean = False)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SCREEN)

    Dim bgColor As Long
    Dim fgColor As Long
    Dim borderColor As Long
    Dim iconText As String
    Dim waitState As String

    If isError Then
        bgColor = RGB(255, 235, 238)
        fgColor = RGB(156, 0, 6)
        borderColor = RGB(220, 80, 80)
        iconText = "×"
        waitState = "エラー"
    ElseIf InStr(resultText, "読込済") > 0 Or InStr(statusText, "利用者QR") > 0 Then
        bgColor = RGB(221, 235, 247)
        fgColor = RGB(0, 92, 175)
        borderColor = RGB(0, 92, 175)
        iconText = "▶"
        waitState = "利用者QR待ち"
    ElseIf InStr(resultText, "取消") > 0 Or InStr(resultText, "リセット") > 0 Then
        bgColor = RGB(255, 242, 204)
        fgColor = RGB(156, 87, 0)
        borderColor = RGB(230, 160, 60)
        iconText = "!"
        waitState = "媒体QR待ち"
    Else
        bgColor = RGB(226, 239, 218)
        fgColor = RGB(0, 97, 0)
        borderColor = RGB(112, 173, 71)
        iconText = "✓"
        waitState = "媒体QR待ち"
    End If

    With ws
        .Range(CELL_STATUS).MergeArea.Value = statusText

        .Range(CELL_RESULT_ICON).MergeArea.Value = iconText
        .Range(CELL_RESULT_ICON).MergeArea.Interior.Color = bgColor
        .Range(CELL_RESULT_ICON).MergeArea.Font.Color = fgColor
        .Range(CELL_RESULT_ICON).MergeArea.Borders.Color = borderColor

        .Range(CELL_RESULT).MergeArea.Value = resultText
        .Range(CELL_RESULT).MergeArea.Interior.Color = bgColor
        .Range(CELL_RESULT).MergeArea.Font.Color = fgColor
        .Range(CELL_RESULT).MergeArea.Borders.Color = borderColor

        .Range(CELL_RESULT_DETAIL).MergeArea.Value = Format(Now, "yyyy/mm/dd hh:mm:ss") & " に処理しました"
        .Range(CELL_RESULT_DETAIL).MergeArea.Interior.Color = bgColor
        .Range(CELL_RESULT_DETAIL).MergeArea.Font.Color = fgColor
        .Range(CELL_RESULT_DETAIL).MergeArea.Borders.Color = borderColor

        .Range(CELL_LAST_TIME).Value = "最終更新：" & Format(Now, "yyyy/mm/dd hh:mm:ss")
        .Range("J27").Value = waitState
    End With

End Sub

'==================================================
' QR正規化
' 例：
' 01        → MEDIA:01
' 1234      → USER:1234
' MEDIA:01  → MEDIA:01
' MEDIA01   → MEDIA:01
' USER:1234 → USER:1234
' USER1234  → USER:1234
'==================================================
Private Function NormalizeQR(ByVal value As Variant) As String

    Dim s As String
    s = Trim(CStr(value))

    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, "　", "")
    s = Replace(s, " ", "")

    If s = "" Then
        NormalizeQR = ""
        Exit Function
    End If

    s = UCase(s)

    If Left(s, 6) = "MEDIA:" Then
        NormalizeQR = s
        Exit Function
    End If

    If Left(s, 5) = "MEDIA" Then
        NormalizeQR = "MEDIA:" & Mid(s, 6)
        Exit Function
    End If

    If Left(s, 5) = "USER:" Then
        NormalizeQR = s
        Exit Function
    End If

    If Left(s, 4) = "USER" Then
        NormalizeQR = "USER:" & Mid(s, 5)
        Exit Function
    End If

    If Left(s, 4) = "SYS:" Then
        NormalizeQR = s
        Exit Function
    End If

    '媒体番号のみ：2桁
    If Len(s) = 2 And IsNumeric(s) Then
        NormalizeQR = "MEDIA:" & Format(CLng(s), "00")
        Exit Function
    End If

    '利用者番号のみ：4桁
    If Len(s) = 4 And IsNumeric(s) Then
        NormalizeQR = "USER:" & Format(CLng(s), "0000")
        Exit Function
    End If

    NormalizeQR = s

End Function

'==================================================
' QR読込画面の保護設定
' QR入力欄だけ入力可
'==================================================
Public Sub QR_ApplyScreenProtection()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SCREEN)

    Application.ScreenUpdating = False

    On Error Resume Next
    ws.Unprotect Password:="qr"
    On Error GoTo ErrHandler

    ws.Cells.Locked = True

    'QR読込欄だけ入力可
    ws.Range(CELL_INPUT).MergeArea.Locked = False
    ws.Range(CELL_INPUT).MergeArea.NumberFormat = "@"

    '内部処理用列を非表示
    ws.Columns("K:L").Hidden = True

    'QR入力欄以外は選択不可
    ws.EnableSelection = xlUnlockedCells

    ws.Protect Password:="qr", _
               UserInterfaceOnly:=True, _
               Contents:=True, _
               DrawingObjects:=False, _
               Scenarios:=True, _
               AllowFiltering:=True

    QR_FocusInput

SafeExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "QR読込画面の保護設定でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation

End Sub

'==================================================
' テーブル取得
'==================================================
Private Function GetTable(ByVal sheetName As String, ByVal tableName As String) As ListObject

    Set GetTable = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)

End Function

'==================================================
' 列番号取得
'==================================================
Private Function GetColIndex(ByVal lo As ListObject, ByVal colName As String) As Long

    On Error GoTo ErrHandler

    GetColIndex = lo.ListColumns(colName).Index
    Exit Function

ErrHandler:
    Err.Raise vbObjectError + 1000, , "テーブル「" & lo.Name & "」に列「" & colName & "」がありません。"

End Function

'==================================================
' 使用記録ログの日時表示形式
'==================================================
Private Sub FormatLogColumns()

    On Error Resume Next

    Dim lo As ListObject
    Set lo = GetTable(SHEET_LOG, TBL_LOG)

    lo.ListColumns("年月日").DataBodyRange.NumberFormat = "yyyy/mm/dd"
    lo.ListColumns("取出時間").DataBodyRange.NumberFormat = "yyyy/mm/dd hh:mm:ss"
    lo.ListColumns("格納時間").DataBodyRange.NumberFormat = "yyyy/mm/dd hh:mm:ss"

    On Error GoTo 0

End Sub

'==================================================
' 経過時間表示
'==================================================
Private Function FormatElapsed(ByVal startTime As Date, ByVal endTime As Date) As String

    Dim totalMinutes As Long
    totalMinutes = DateDiff("n", startTime, endTime)

    If totalMinutes < 0 Then totalMinutes = 0

    Dim d As Long
    Dim h As Long
    Dim m As Long

    d = totalMinutes \ 1440
    h = (totalMinutes Mod 1440) \ 60
    m = totalMinutes Mod 60

    If d > 0 Then
        FormatElapsed = d & "日 " & Format(h, "00") & ":" & Format(m, "00")
    Else
        FormatElapsed = h & "時間" & m & "分"
    End If

End Function