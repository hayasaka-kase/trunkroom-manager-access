VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MFZ0000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 郵便番号変換
'
'        PROGRAM_NAME    : 郵便番号変換
'        PROGRAM_ID      : MFZ0000
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2004/04/09
'        CERATER         : K.ISHIZAKA
'        Ver             : 0.0
'
'        UPDATE          :
'        UPDATER         :
'        Ver             :
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private bSkip               As Boolean
Private lngMaxCount         As Long

'==============================================================================*
'
'        MODULE_NAME      :検索
'        MODULE_ID        :MFZ0000_FindSearch
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MFZ0000_FindSearch(lngSTEP As Long) As Boolean
    Dim lngIndex            As Long
    Dim strSearch           As String
    On Error GoTo Err_Handle

    Screen.MousePointer = 11
    strSearch = Me.MFZ0000_TextTODON.VALUE
    With Me.MFZ0000_ListTODON
        If strSearch <> "" Then
            For lngIndex = IIf(.ListIndex < 0, IIf(lngSTEP > 0, 0, lngMaxCount - 1), .ListIndex) + lngSTEP _
                To IIf(lngSTEP < 0, 0, lngMaxCount - 1) _
                Step lngSTEP
                If InStr(.Column(0, lngIndex), strSearch) > 0 Then
                    .Selected(lngIndex) = True
                    Screen.MousePointer = 0
                    On Error Resume Next
                    If .Column(0, lngIndex) = strSearch Then
                        Me.MFZ0000_CmdChange.SetFocus
                    Else
                        Me.MFZ0000_TextTODON.SetFocus
                    End If
                    MFZ0000_FindSearch = True
                    Exit Function
                End If
            Next
        End If
        If .ListIndex > 0 Then
            .Selected(.ListIndex) = False
        End If
        '.Selected(IIF(lngStep < 0, 0, lngMaxCount - 1)) = True
        Screen.MousePointer = 0
        On Error Resume Next
        Me.MFZ0000_TextTODON.SetFocus
        MFZ0000_FindSearch = False
    End With
    Exit Function

Err_Handle:
    MsgBox Err.Description
    On Error Resume Next
    MFZ0000_FindSearch = False
End Function

'==============================================================================*
'
'        MODULE_NAME      :検索下方向
'        MODULE_ID        :FMZ0000_CmdDown_Click
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FMZ0000_CmdDown_Click()
    Call MFZ0000_FindSearch(1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :検索上方向
'        MODULE_ID        :FMZ0000_CmdUP_Click
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FMZ0000_CmdUP_Click()
    Call MFZ0000_FindSearch(-1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :住所表示
'        MODULE_ID        :FMZ0000_ListTODONRequery
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FMZ0000_ListTODONRequery(strWhere As String)
    Dim strSQL              As String
    Dim rsAddr              As Recordset
    On Error GoTo Err_Handle

    strSQL = " FROM ADDR_TABL WHERE " & strWhere

    Set rsAddr = CurrentDb.OpenRecordset("SELECT COUNT(*) " & strSQL, dbOpenDynaset, dbReadOnly)
    lngMaxCount = rsAddr.Fields(0).VALUE
    rsAddr.Close

    strSQL = "SELECT trim(ADDRT_TODON) + trim(ADDRT_SITSN) + trim(ADDRT_CHINN),ADDRT_NYUBB " & strSQL
    strSQL = strSQL & " ORDER BY ADDRT_ORDEB"
    Me.MFZ0000_ListTODON.RowSource = strSQL
    Me.MFZ0000_ListTODON.Requery

    Me.MFZ0000_TextNYUBB.VALUE = ""
    Me.MFZ0000_TextTODON.VALUE = ""
    Exit Sub

Err_Handle:
    MsgBox Err.Description
    On Error Resume Next
End Sub

'==============================================================================*
'
'        MODULE_NAME      :初期表示
'        MODULE_ID        :FMZ0000_Default
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FMZ0000_Default()
    Dim iMidL               As Integer
    Dim strSQL              As String
    Dim rsAddr              As Recordset
    On Error GoTo Err_Handle

    If MSZ0000_GET_NYUBB() <> "" Then
        strSQL = "SELECT ADDRT_TODON FROM ADDR_TABL " & _
                 "WHERE ADDRT_NYUBB = '" & MSZ0000_GET_NYUBB() & "'"

        Set rsAddr = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
        If Not rsAddr.EOF Then
            Me.MFZ0000_CmbTODON = rsAddr.Fields(0).VALUE
            rsAddr.Close
            Me.MFZ0000_Option.VALUE = 2
            Call MFZ0000_Option_Click
            Me.MFZ0000_TextNYUBB.VALUE = MSZ0000_GET_NYUBB()
            Me.MFZ0000_TextTODON.VALUE = MSZ0000_GET_ADDRN()
            Me.MFZ0000_CmdChange.SetFocus
            Exit Sub
        End If
        rsAddr.Close
    End If

    If MSZ0000_GET_ADDRN() <> "" Then
        iMidL = InStr(MSZ0000_GET_ADDRN(), "県")
        If iMidL > 0 Then
            strSQL = Left(MSZ0000_GET_ADDRN(), iMidL)
        Else
            strSQL = Left(MSZ0000_GET_ADDRN(), 3)
            Select Case strSQL
            Case "東京都", "京都府", "大阪府", "北海道"
            Case Else
                strSQL = ""
            End Select
        End If
        If strSQL <> "" Then
            strSQL = "SELECT ADDRT_TODON FROM ADDR_TABL " & _
                     "WHERE ADDRT_TODON = '" & strSQL & "' GROUP BY ADDRT_TODON"
        End If
    End If

    If strSQL <> "" Then
        Set rsAddr = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
        If Not rsAddr.EOF Then
            Me.MFZ0000_CmbTODON = rsAddr.Fields(0).VALUE
            rsAddr.Close
            Me.MFZ0000_Option.VALUE = 2
            Call MFZ0000_Option_Click
            Me.MFZ0000_TextNYUBB.VALUE = MSZ0000_GET_NYUBB()
            Me.MFZ0000_TextTODON.VALUE = MSZ0000_GET_ADDRN()
            Me.MFZ0000_CmdChange.SetFocus
            Exit Sub
        End If
        rsAddr.Close
    End If

    Me.MFZ0000_CmbTODON = "神奈川県"
    Me.MFZ0000_Option.VALUE = 1
    Call MFZ0000_Option_Click
    Me.MFZ0000_TextNYUBB.VALUE = MSZ0000_GET_NYUBB()
    Me.MFZ0000_TextTODON.VALUE = MSZ0000_GET_ADDRN()
    Exit Sub

Err_Handle:
    MsgBox Err.Description
    On Error Resume Next
End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォーム初期値
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    bSkip = True
    Call FMZ0000_Default
    bSkip = False
End Sub

'==============================================================================*
'
'        MODULE_NAME      :都道府県選択
'        MODULE_ID        :MFZ0000_CmbTODON_Change
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_CmbTODON_Change()
    On Error GoTo Err_Handle
    With Me.MFZ0000_CmbTODON
        If .ListIndex >= 0 Then
            Call FMZ0000_ListTODONRequery("ADDRT_TODON = '" & .ItemData(.ListIndex) & "'")
        End If
    End With
    Exit Sub

Err_Handle:
    MsgBox Err.Description
    On Error Resume Next
End Sub

'==============================================================================*
'
'        MODULE_NAME      :キャンセル
'        MODULE_ID        :MFZ0000_CmdCansel_Click
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_CmdCansel_Click()
    On Error GoTo Err_Handle
    doCmd.Close
    Exit Sub

Err_Handle:
    MsgBox Err.Description
    On Error Resume Next
End Sub

'==============================================================================*
'
'        MODULE_NAME      :変換
'        MODULE_ID        :MFZ0000_CmdChange_Click
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_CmdChange_Click()
    Dim lngIndex            As Long
    Dim intXOR              As Integer
    Dim strSearch           As String
    On Error GoTo Err_Handle

    bSkip = True
    If Me.MFZ0000_ListTODON.ListIndex = -1 Then
        strSearch = Trim(Nz(Me.MFZ0000_TextNYUBB.VALUE))
        If InStr(strSearch, "-") > 0 Then
            strSearch = Left(strSearch, 3) & Mid(strSearch, 5)
        End If
        If strSearch = "" Then
            strSearch = Trim(Nz(Me.MFZ0000_TextTODON.VALUE))
            intXOR = 1
        Else
            intXOR = 0
        End If
        If strSearch <> "" Then
            With Me.MFZ0000_ListTODON
                For lngIndex = 0 To lngMaxCount - 1
                    If InStr(.Column(1 Xor intXOR, lngIndex), strSearch) > 0 Then
                        .Selected(lngIndex) = True
                        strSearch = .Column(0, lngIndex)
                        If intXOR Then
                            Me.MFZ0000_TextNYUBB.VALUE = .Column(1, .ListIndex)
                        End If
                        Exit For
                    End If
                Next
                If lngIndex >= lngMaxCount Then
                    strSearch = ""
                End If
            End With
        End If
    Else
        strSearch = ""
        With Me.MFZ0000_ListTODON
            If .ListIndex >= 0 Then
                Me.MFZ0000_TextNYUBB.VALUE = .Column(1, .ListIndex)
                strSearch = .Column(0, .ListIndex)
            End If
        End With
    End If
    lngIndex = InStr(strSearch, "以下に掲載がない場合")
    If lngIndex > 0 Then
        strSearch = Left(strSearch, lngIndex - 1)
    Else
        lngIndex = InStr(strSearch, "（")
        If lngIndex > 0 Then
            strSearch = Left(strSearch, lngIndex - 1)
        End If
    End If
    If strSearch <> "" Then
        Me.MFZ0000_TextTODON.VALUE = strSearch
        If lngIndex > 0 Then
            With Me.MFZ0000_TextTODON
                .SetFocus
                .SelStart = LenB(.VALUE)
            End With
        Else
            Me.MFZ0000_CmdExit.SetFocus
        End If
    Else
        With Me.MFZ0000_TextTODON
            .SetFocus
            .SelStart = 0
            .SelLength = LenB(.VALUE)
        End With
    End If
    bSkip = False
    Exit Sub

Err_Handle:
    MsgBox Err.Description
    On Error Resume Next
End Sub

'==============================================================================*
'
'        MODULE_NAME      :設定
'        MODULE_ID        :MFZ0000_CmdExit_Click
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_CmdExit_Click()
    On Error GoTo Err_Handle
    Call MSZ0000_SET(Me.MFZ0000_TextNYUBB.VALUE, Me.MFZ0000_TextTODON.VALUE)
    Exit Sub

Err_Handle:
    MsgBox Err.Description
    On Error Resume Next
End Sub

'==============================================================================*
'
'        MODULE_NAME      :変換
'        MODULE_ID        :MFZ0000_ListTODON_DblClick
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_ListTODON_DblClick(Cancel As Integer)
    Call MFZ0000_CmdChange_Click
End Sub

'==============================================================================*
'
'        MODULE_NAME      :規定値ボタン設定
'        MODULE_ID        :MFZ0000_ListTODON_Enter
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_ListTODON_Enter()
    On Error Resume Next
    Me.MFZ0000_CmdChange.Default = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :規定値ボタン解除
'        MODULE_ID        :MFZ0000_ListTODON_Exit
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_ListTODON_Exit(Cancel As Integer)
    On Error Resume Next
    Me.MFZ0000_CmdChange.Default = False
End Sub

'==============================================================================*
'
'        MODULE_NAME      :近隣・都道府県選択
'        MODULE_ID        :MFZ0000_Option_Click
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_Option_Click()
    On Error GoTo Err_Handle
    Select Case Me.MFZ0000_Option.VALUE
    Case 1
        Me.MFZ0000_CmbTODON.Enabled = False
        Call FMZ0000_ListTODONRequery("ADDRT_ORDEB < '990000000'")
    Case 2
        Me.MFZ0000_CmbTODON.Enabled = True
        Call MFZ0000_CmbTODON_Change
    End Select
    Exit Sub

Err_Handle:
    MsgBox Err.Description
    On Error Resume Next
End Sub

'==============================================================================*
'
'        MODULE_NAME      :初期値設定
'        MODULE_ID        :MFZ0000_TextNYUBB_Change
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_TextNYUBB_Change()
    On Error GoTo Err_Handle
    If bSkip Then Exit Sub
    Me.MFZ0000_TextTODON.VALUE = ""
    With Me.MFZ0000_ListTODON
        If .ListIndex >= 0 Then
            .Selected(.ListIndex) = False
        End If
    End With
    Exit Sub

Err_Handle:
    MsgBox Err.Description
    On Error Resume Next
End Sub

'==============================================================================*
'
'        MODULE_NAME      :規定値ボタン設定
'        MODULE_ID        :MFZ0000_TextNYUBB_Enter
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_TextNYUBB_Enter()
    On Error Resume Next
    Me.MFZ0000_CmdChange.Default = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :規定値ボタン解除
'        MODULE_ID        :MFZ0000_TextNYUBB_Exit
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_TextNYUBB_Exit(Cancel As Integer)
    On Error Resume Next
    Me.MFZ0000_CmdChange.Default = False
End Sub

'==============================================================================*
'
'        MODULE_NAME      :初期値設定
'        MODULE_ID        :MFZ0000_TextTODON_Change
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_TextTODON_Change()
    On Error Resume Next
    If bSkip Then Exit Sub
    Me.MFZ0000_TextNYUBB.VALUE = ""
End Sub

'==============================================================================*
'
'        MODULE_NAME      :規定値ボタン設定
'        MODULE_ID        :MFZ0000_TextNYUBB_Enter
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_TextTODON_Enter()
    On Error Resume Next
    Me.FMZ0000_CmdDown.Default = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :規定値ボタン解除
'        MODULE_ID        :MFZ0000_TextNYUBB_Exit
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFZ0000_TextTODON_Exit(Cancel As Integer)
    On Error Resume Next
    Me.FMZ0000_CmdDown.Default = False
End Sub
'****************************  ended or program ********************************
