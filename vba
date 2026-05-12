Option Explicit

'==============================
' シート名
'==============================
Private Const SHEET_SCREEN As String = "QR読込画面"
Private Const SHEET_LOG As String = "使用記録ログ"
Private Const SHEET_MEDIA As String = "媒体マスタ"
Private Const SHEET_USER As String = "利用者マスタ"
Private Const SHEET_PERMIT As String = "媒体利用許可マスタ"
Private Const SHEET_ERROR As String = "異常ログ"

'==============================
' テーブル名
'==============================
Private Const TBL_LOG As String = "tblLog"
Private Const TBL_MEDIA As String = "tblMedia"
Private Const TBL_USER As String = "tblUser"
Private Const TBL_PERMIT As String = "tblPermit"
Private Const TBL_ERROR As String = "tblError"

'==============================
' QR読込画面のセル
'==============================
Private Const CELL_STATUS As String = "B3"
Private Const CELL_INPUT As String = "B5"
Private Const CELL_MEDIA_NAME As String = "B7"
Private Const CELL_USER_NAME As String = "B8"
Private Const CELL_RESULT As String = "B10"
Private Const CELL_LAST_TIME As String = "B12"

'処理中データ保存用
Private Const CELL_PENDING_MEDIA_QR As String = "H3"
Private Const CELL_PENDING_MEDIA_NAME As String = "H4"

'貸出中一覧の表示開始位置
Private Const LOAN_LIST_TITLE_ROW As Long = 14
Private Const LOAN_LIST_HEADER_ROW As Long = 15
Private Const LOAN_LIST_FIRST_ROW As Long = 16

'==================================================
' 初期表示
'==================================================
Public Sub QR_InitializeScreen()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SCREEN)

    With ws
        .Range("A1").Value = "可搬記憶媒体 QR貸出・返納システム"

        .Range("A3").Value = "現在状態"
        .Range("A5").Value = "QR読込欄"
        .Range("A7").Value = "読込済媒体"
        .Range("A8").Value = "読込済利用者"
        .Range("A10").Value = "処理結果"
        .Range("A12").Value = "最終処理時刻"

        .Range(CELL_INPUT).NumberFormat = "@"
        .Range(CELL_INPUT).ClearContents

        .Range(CELL_PENDING_MEDIA_QR).ClearContents
        .Range(CELL_PENDING_MEDIA_NAME).ClearContents
        .Range(CELL_MEDIA_NAME).ClearContents
        .Range(CELL_USER_NAME).ClearContents

        .Range(CELL_STATUS).Value = "媒体QRを読み込んでください"
        .Range(CELL_RESULT).Value = ""
        .Range(CELL_LAST_TIME).Value = ""

        .Columns("H").Hidden = True
    End With

    QR_UpdateLoanList
    QR_FocusInput

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

            If Left(qr, 6) = "MEDIA" Then
                ProcessMediaQR qr

            ElseIf Left(qr, 5) = "USER:" Then
                ProcessUserQR qr

            Else
                SetScreenMessage "媒体QRを読み込んでください", "不明なQRです：" & rawQr, True
                LogError "読込", rawQr, "", "", "不明QR", "MEDIA、USER、SYSのいずれでもありません。"
            End If

    End Select

    QR_UpdateLoanList
    QR_FocusInput
    Exit Sub

ErrHandler:
    SetScreenMessage "媒体QRを読み込んでください", "エラー：" & Err.Description, True
    LogError "システム", rawQr, "", "", "VBAエラー", Err.Description
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
        ClearPending
        Exit Sub
    End If

    If mediaState <> "使用可" Then
        SetScreenMessage "媒体QRを読み込んでください", "この媒体は使用できません：" & registerNo & " / 状態：" & mediaState, True
        LogError "貸出", mediaQR, mediaQR, "", "媒体使用不可", "媒体状態が使用可ではありません。状態：" & mediaState
        ClearPending
        Exit Sub
    End If

    Dim activeRowIndex As Long
    Dim activeCount As Long

    FindActiveLogRows registerNo, activeRowIndex, activeCount

    If activeCount > 1 Then
        SetScreenMessage "管理者に連絡してください", "同じ媒体の貸出中記録が複数あります：" & registerNo, True
        LogError "返納", mediaQR, mediaQR, "", "貸出中重複", "同一媒体の格納時間空欄が複数あります。登録番号：" & registerNo
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
        Exit Sub
    End If

    Dim userName As String

    If Not FindUser(userQR, userName) Then
        SetScreenMessage "利用者QRを読み込んでください", "未登録の利用者です：" & userQR, True
        LogError "貸出", userQR, mediaQR, userQR, "未登録利用者", "利用者マスタに登録されていません。"
        Exit Sub
    End If

    ws.Range(CELL_USER_NAME).Value = userName

    If Not IsPermitted(mediaQR, userQR) Then
        SetScreenMessage "媒体QRを読み込んでください", "貸出不可：" & userName & " は " & registerNo & " を使用できません。", True
        LogError "貸出", userQR, mediaQR, userQR, "権限なし", userName & " は " & registerNo & " の利用許可がありません。"
        ClearPending
        Exit Sub
    End If

    Dim activeRowIndex As Long
    Dim activeCount As Long

    FindActiveLogRows registerNo, activeRowIndex, activeCount

    If activeCount > 0 Then
        SetScreenMessage "媒体QRを読み込んでください", "この媒体はすでに貸出中です：" & registerNo, True
        LogError "貸出", userQR, mediaQR, userQR, "二重貸出", "貸出中の媒体を再度貸出しようとしました。"
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

    ws.Range("A" & LOAN_LIST_TITLE_ROW & ":D1000").ClearContents
    ws.Range("A" & LOAN_LIST_TITLE_ROW & ":D1000").Interior.Pattern = xlNone

    ws.Range("A" & LOAN_LIST_TITLE_ROW).Value = "現在貸出中一覧"
    ws.Range("A" & LOAN_LIST_HEADER_ROW).Value = "登録番号"
    ws.Range("B" & LOAN_LIST_HEADER_ROW).Value = "使用者"
    ws.Range("C" & LOAN_LIST_HEADER_ROW).Value = "取出時間"
    ws.Range("D" & LOAN_LIST_HEADER_ROW).Value = "経過時間"

    ws.Range("A" & LOAN_LIST_HEADER_ROW & ":D" & LOAN_LIST_HEADER_ROW).Font.Bold = True

    Dim lo As ListObject
    Set lo = GetTable(SHEET_LOG, TBL_LOG)

    If lo.DataBodyRange Is Nothing Then
        ws.Range("A" & LOAN_LIST_FIRST_ROW).Value = "貸出中なし"
        Exit Sub
    End If

    Dim cReg As Long
    Dim cUser As Long
    Dim cOut As Long
    Dim cIn As Long

    cReg = GetColIndex(lo, "登録番号")
    cUser = GetColIndex(lo, "使用者")
    cOut = GetColIndex(lo, "取出時間")
    cIn = GetColIndex(lo, "格納時間")

    Dim i As Long
    Dim outRow As Long
    outRow = LOAN_LIST_FIRST_ROW

    For i = 1 To lo.DataBodyRange.Rows.Count

        If Trim(CStr(lo.DataBodyRange.Cells(i, cIn).Value)) = "" Then

            ws.Cells(outRow, "A").Value = lo.DataBodyRange.Cells(i, cReg).Value
            ws.Cells(outRow, "B").Value = lo.DataBodyRange.Cells(i, cUser).Value
            ws.Cells(outRow, "C").Value = lo.DataBodyRange.Cells(i, cOut).Value
            ws.Cells(outRow, "C").NumberFormat = "yyyy/mm/dd hh:mm:ss"

            If IsDate(lo.DataBodyRange.Cells(i, cOut).Value) Then
                ws.Cells(outRow, "D").Value = FormatElapsed(CDate(lo.DataBodyRange.Cells(i, cOut).Value), Now)
            Else
                ws.Cells(outRow, "D").Value = "不明"
            End If

            outRow = outRow + 1

        End If

    Next i

    If outRow = LOAN_LIST_FIRST_ROW Then
        ws.Range("A" & LOAN_LIST_FIRST_ROW).Value = "貸出中なし"
    End If

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

    With ws
        .Range(CELL_STATUS).Value = statusText
        .Range(CELL_RESULT).Value = resultText
        .Range(CELL_LAST_TIME).Value = Now
        .Range(CELL_LAST_TIME).NumberFormat = "yyyy/mm/dd hh:mm:ss"

        If isError Then
            .Range(CELL_RESULT).Interior.Color = RGB(255, 199, 206)
            .Range(CELL_RESULT).Font.Color = RGB(156, 0, 6)
        Else
            .Range(CELL_RESULT).Interior.Color = RGB(198, 239, 206)
            .Range(CELL_RESULT).Font.Color = RGB(0, 97, 0)
        End If
    End With

End Sub

'==================================================
' QR正規化
' 例：
' 01       → MEDIA:01
' 1234     → USER:1234
' MEDIA:01 → MEDIA:01
' USER:1234 → USER:1234
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

    If Left(s, 6) = "MEDIA" Then
        If Left(s, 6) = "MEDIA:" Then
            NormalizeQR = s
        Else
            NormalizeQR = Replace(s, "MEDIA", "MEDIA:")
        End If
        Exit Function
    End If

    If Left(s, 5) = "USER:" Then
        NormalizeQR = s
        Exit Function
    End If

    If Left(s, 4) = "USER" Then
        NormalizeQR = Replace(s, "USER", "USER:")
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
