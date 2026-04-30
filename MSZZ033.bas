Attribute VB_Name = "MSZZ033"
'****************************  strat or program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬総合システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : 入力チェック
'       PROGRAM_ID      : MSZZ033
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2007/07/12
'       CERATER         : K.ISHIZAKA
'       Ver             : 0.0
'
'       UPDATE          : 2007/11/15
'       UPDATER         : S.SHIBAZAKI
'       Ver             : 0.1
'                       : 入力桁数チェックでＴＲＩＭしない
'
'       UPDATE          : 2008/04/12
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.2
'                       : MSZZ043の追加による関数[checkAnsiChar]の追加
'
'       UPDATE          : 2008/12/04
'       UPDATER         : S.SHIBAZAKI
'       Ver             : 0.3
'                       : 数値チェックで小数入力不可のメッセージ出力追加
'
'       UPDATE          : 2010/02/20
'       UPDATER         : S.SHIBAZAKI
'       Ver             : 0.4
'                       : 区分重複チェック関数のパラメータ追加
'
'==============================================================================*
Option Compare Database
Option Explicit

Private Const C_COLOR_ERROR     As Long = 16711935
Private Const C_COLOR_INFO      As Long = 33023
Private Const C_COLOR_LABEL     As Long = 8388608

'==============================================================================*
'
'       MODULE_NAME     : 必須入力チェック
'       MODULE_ID       : checkInput
'       CREATE_DATE     : 2007/07/12            K.ISHIZAKA
'       PARAM           : ctrl                  コントロール(I)
'                       : [maxLength]           最大バイト数：省略可(I)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function checkInput(ctrl As Control, Optional maxLength As Long = 256, Optional strTypeName As String = "") As Boolean
    Dim strValue            As String
    On Error GoTo ErrorHandler
    
    If (ctrl.Enabled = False) Or (ctrl.Locked) Then
        checkInput = True
        Exit Function
    End If
    If strTypeName = "" Then
        strTypeName = TypeName(ctrl)
    End If
    Select Case UCase(strTypeName)
    Case "COMBOBOX"
        If ctrl.ListIndex >= 0 Then
            checkInput = True
        Else
            checkInput = checkMsgBox(ctrl, "選択してください。")
        End If
    Case "TEXTBOX"
        strValue = Trim(Nz(ctrl, ""))
        If strValue <> "" Then
            checkInput = checkLength(ctrl, maxLength)
        Else
            checkInput = checkMsgBox(ctrl, "入力してください。")
        End If
    Case Else
        checkInput = checkMsgBox(ctrl, "データ型[" & TypeName(ctrl) & "]に適応してません。管理者に連絡してください。")
    End Select
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "checkInput" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 日付入力チェック
'       MODULE_ID       : checkIsDate
'       CREATE_DATE     : 2007/07/12            K.ISHIZAKA
'       PARAM           : ctrl                  コントロール(I)
'                       : maxLength             最大バイト数(I)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function checkIsDate(ctrl As Control, Optional maxLength As Long = 8) As Boolean
    Dim strValue            As String
    On Error GoTo ErrorHandler
    
    strValue = Trim(Nz(ctrl, ""))
    If strValue = "" Then
        checkIsDate = True
    Else
        If Len(strValue) <> maxLength Then
            checkIsDate = checkMsgBox(ctrl, Format(maxLength) & "桁の日付(" & Left("yyyymmdd", maxLength) & ")を入力してください。")
        Else
            If MSZZ012_M00(strValue) Then
                checkIsDate = True
            Else
                checkIsDate = checkMsgBox(ctrl, "日付(" & Left("yyyymmdd", maxLength) & ")を入力してください。")
            End If
        End If
    End If
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "checkIsDate" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 数値入力チェック
'       MODULE_ID       : checkIsNumeric
'       CREATE_DATE     : 2007/07/12            K.ISHIZAKA
'       PARAM           : ctrl                  コントロール(I)
'                       : lngPrecision          有効桁数(I)
'                       : [strOtherChar]        数値以外で許可する文字(I)
'                       : [lngScale]            小数点部桁数(I)省略時:0
'                       : [bMinus]              マイナス文字の許可(I)省略時:False
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function checkIsNumeric(ctrl As Control, ByVal lngPrecision As Long, Optional strOtherChar As String = "", Optional lngScale As Long = 0, Optional bMinus As Boolean = False) As Boolean
    Dim strValue            As String
    Dim strArrs()           As String
    Dim lngLens(1)           As Long
    Dim i                   As Long
    On Error GoTo ErrorHandler
    
    strValue = Trim(Nz(ctrl, ""))
    If strValue = "" Then
        checkIsNumeric = True
    Else
        lngLens(0) = lngPrecision - lngScale
        lngLens(1) = lngScale
        strValue = Replace(strValue, ",", "")
        If bMinus And (Left(strValue, 1) = "-") Then
            strArrs = Split(Mid(strValue, 2), ".")
        Else
            strArrs = Split(strValue, ".")
        End If
        If UBound(strArrs) > 1 Then
            checkIsNumeric = checkMsgBox(ctrl, "数字を入力してください。")
            Exit Function
        End If
        '↓INSERT 2008/12/04 SHIBAZAKI
        If UBound(strArrs) > 0 And lngScale = 0 Then
            checkIsNumeric = checkMsgBox(ctrl, "小数の入力は出来ません。")
            Exit Function
        End If
        '↑INSERT 2008/12/04 SHIBAZAKI
        For i = 0 To UBound(strArrs)
            If MSZZ0016_M00(Replace(strArrs(i), strOtherChar, ""), charTypeNumber) = False Then
                checkIsNumeric = checkMsgBox(ctrl, "数字を入力してください。")
                Exit Function
            End If
            If CmLenb(strArrs(i)) > lngLens(i) Then
                'DELETE 2008/12/04 SHIBAZAKI
                'checkIsNumeric = checkMsgBox(ctrl, IIf(lngScale > 0, Choose(i + 1, "整数部", "少数部"), "桁") & "が大きすぎます。" & vbCrLf & Format(lngLens(i)) & "桁までにしてください。")
                'INSERT 2008/12/04 SHIBAZAKI
                checkIsNumeric = checkMsgBox(ctrl, IIf(lngScale > 0, Choose(i + 1, "整数部", "小数部"), "桁") & "が大きすぎます。" & vbCrLf & Format(lngLens(i)) & "桁までにしてください。")
                Exit Function
            End If
        Next
        checkIsNumeric = True
    End If
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "checkIsNumeric" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 符号チェック
'       MODULE_ID       : checkSign
'       CREATE_DATE     : 2007/07/12            K.ISHIZAKA
'       PARAM           : ctrl                  コントロール(I)
'                       : intSignError          エラーの符号(I)
'                                               マイナスがエラーのとき(-1)
'                                               ゼロがエラーのとき(0)
'                                               プラスがエラーのとき(1)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function checkSign(ctrl As TextBox, ByVal intSignError As Integer) As Boolean
    On Error GoTo ErrorHandler
    
    If Sgn(Nz(ctrl.VALUE, 0)) = intSignError Then
        checkSign = checkMsgBox(ctrl, "０" & Choose(intSignError + 2, "より大きい", "以外の", "より小さい") & "数を入力してください。")
    Else
        checkSign = True
    End If
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "checkSign" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 半角カナ入力チェック
'       MODULE_ID       : checkIsKana
'       CREATE_DATE     : 2007/07/12            K.ISHIZAKA
'       PARAM           : ctrl                  コントロール(I)
'                       : maxLength             最大バイト数(I)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function checkIsKana(ctrl As TextBox, ByVal maxLength As Long) As Boolean
    On Error GoTo ErrorHandler

    If checkLength(ctrl, maxLength) = False Then
        checkIsKana = False
        Exit Function
    End If
    If MSZZ0016_M00(Replace(Nz(ctrl), " ", ""), charTypeKana) = False Then
        checkIsKana = checkMsgBox(ctrl, "半角カナで入力してください。")
        Exit Function
    End If

    checkIsKana = True
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "checkIsKana" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 区分重複チェック
'       MODULE_ID       : checkRepeat
'       CREATE_DATE     : 2007/07/12            K.ISHIZAKA
'       PARAM           : ctrl                  コントロール(I)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'UPDATE 2010/02/20 SHIBAZAKI
'Public Function checkRepeat(ctrl As ComboBox) As Boolean
Public Function checkRepeat(ctrl As ComboBox, Optional strAddCond As String = "") As Boolean
    Dim frm                 As Form
    Dim lngCnt              As Long
    Dim strCond             As String               'INSERT 2010/02/20 SHIBAZAKI
    On Error GoTo ErrorHandler
    
    Set frm = GetFormObject(ctrl)
    With ctrl
        If .OldValue = .VALUE Then
            lngCnt = 0
        Else
'            lngCnt = DCount(.ControlSource, frm.RecordSource, .ControlSource & "='" & .Value & "'")    'DELETE 2010/02/20 SHIBAZAKI
            '↓INSERT 2010/02/20 SHIBAZAKI
            strCond = .ControlSource & "='" & .VALUE & "'"
            If strAddCond <> "" Then
                strCond = strCond & " AND " & strAddCond
            End If
            lngCnt = DCount(.ControlSource, frm.RecordSource, strCond)
            '↑INSERT 2010/02/20 SHIBAZAKI
        End If
    End With
    If lngCnt > 0 Then
        checkRepeat = checkMsgBox(ctrl, "既に選択されている区分です。他を選択してください。")
    Else
        checkRepeat = True
    End If
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "checkRepeat" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 最大桁数入力チェック
'       MODULE_ID       : checkLength
'       CREATE_DATE     : 2007/07/12            K.ISHIZAKA
'       PARAM           : ctrl                  コントロール(I)
'                       : maxLength             最大バイト数(I)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function checkLength(ctrl As Control, ByVal maxLength As Long) As Boolean
    Dim strValue            As String
    On Error GoTo ErrorHandler
    
'    strValue = Trim(Nz(ctrl, ""))      'DELETE 2007/11/15 SHIBAZAKI
    strValue = Nz(ctrl, "")             'INSERT 2007/11/15 SHIBAZAKI
    If strValue = "" Then
        checkLength = True
    Else
        If CmLenb(strValue) <= maxLength Then
'            checkLength = True                                                 'DELETE 2008/04/12 K.ISHIZAKA
            checkLength = checkAnsiChar(ctrl)                                   'INSERT 2008/04/12 K.ISHIZAKA
        Else
            checkLength = checkMsgBox(ctrl, "文字列が長すぎます。" & vbCrLf & Format(maxLength) & "桁までにしてください。")
        End If
    End If
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "checkLength" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ANSI変換できない文字の入力チェック
'       MODULE_ID       : checkAnsiChar
'       CREATE_DATE     : 2008/04/12            K.ISHIZAKA
'       PARAM           : ctrl                  コントロール(I)
'                       : maxLength             最大バイト数(I)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function checkAnsiChar(ctrl As Control) As Boolean
    Dim i                   As Long
    On Error GoTo ErrorHandler
    
    i = MSZZ0043_M00(Nz(ctrl, ""))
    If i > 0 Then
        checkAnsiChar = checkMsgBox(ctrl, "利用できない文字が含まれています。" & Format(i) & "文字目")
    Else
        checkAnsiChar = True
    End If
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "checkAnsiChar" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : メッセージボックス
'       MODULE_ID       : checkMsgBox
'       CREATE_DATE     : 2007/07/12            K.ISHIZAKA
'       PARAM           : ctrl                  コントロール(I)
'                       : strMsg                メッセージ(I)
'                       : YesNo                 質問するか(I)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function checkMsgBox(ctrl As Control, ByVal strMsg As String, Optional YesNo As Boolean = False) As Boolean
    Dim frm                 As Form
    Dim lblCtrl             As label
    Dim bkColor             As Long
    Dim strCtrl             As String
    On Error GoTo ErrorHandler
    
    Set frm = GetFormObject(ctrl)
    On Error Resume Next
    ctrl.SetFocus
    Select Case Left(ctrl.NAME, 4)
    Case "txt_", "cmb_"
        Set lblCtrl = frm.Controls("lbl_" & Mid(ctrl.NAME, 5))
    Case Else
        Set lblCtrl = frm.Controls("lbl_" & ctrl.NAME)
    End Select
    On Error GoTo ErrorHandler
    bkColor = ctrl.BackColor
    If Not lblCtrl Is Nothing Then
        lblCtrl.BackColor = C_COLOR_INFO
    End If
    If ctrl.ControlSource = "" Then
        ctrl.BackColor = C_COLOR_ERROR
    End If
    If YesNo Then
        If MsgBox(strMsg, vbQuestion + vbYesNo, frm.Caption) = vbYes Then
            checkMsgBox = True
            Exit Function
        End If
    Else
        MsgBox strMsg, vbInformation + vbOKOnly, frm.Caption
    End If
    ctrl.BackColor = bkColor
    If Not lblCtrl Is Nothing Then
        lblCtrl.BackColor = C_COLOR_LABEL
    End If
    checkMsgBox = False
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "checkMsgBox" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 対象コントロールを反転強調してメッセージボックス
'       MODULE_ID       : BlnkMsgBox
'       CREATE_DATE     : 2007/07/12            K.ISHIZAKA
'       PARAM           : ctrl                  コントロール(I)
'                       : strMsg                メッセージ(I)
'                       : lblCtrls ...          反転ラベル(I)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function BlnkMsgBox(ctrl As Control, ByVal strMsg As String, ParamArray lblCtrls()) As Boolean
    Dim lblCtrl             As Variant
    On Error GoTo ErrorHandler
    
    For Each lblCtrl In lblCtrls
        lblCtrl.BackColor = C_COLOR_INFO
    Next
    BlnkMsgBox = checkMsgBox(ctrl, strMsg)
    For Each lblCtrl In lblCtrls
        lblCtrl.BackColor = C_COLOR_LABEL
    Next
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "BlnkMsgBox" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  ended or program ********************************
