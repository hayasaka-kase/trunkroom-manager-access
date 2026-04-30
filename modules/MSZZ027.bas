Attribute VB_Name = "MSZZ027"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : コンボボックス制御関数
'        PROGRAM_ID      : MSZZ027
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2007/02/22
'        CERATER         : K.ISHZIAKA
'        Ver             : 0.0
'
'        UPDATE          : 2007/02/26
'        UPDATER         : K.ISHZIAKA
'        Ver             : 0.1
'                        : イベント呼び出し中にComboBoxのリストが更新されると
'                          想定外の名称などが表示されてしまうので
'                          名称など全てを設定した後にイベントを呼び出す
'
'        UPDATE          : 2007/03/19
'        UPDATER         : K.ISHZIAKA
'        Ver             : 0.2
'                        : ComboBoxのリストのみ表示する列の場合
'                          名称などを設定するコントロールがないので
'                          空文字指定されたときの対応
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'
'       MODULE_NAME     : 区分などを選択後、名称などを表示
'       MODULE_ID       : ComboBox_AfterUpdate
'       CREATE_DATE     : 2007/02/22            K.ISHIZAKA
'       PARAM           : cbo                   区分などのコンボボックス(I)
'                       : CtrlNames ...         名称などのコントロール名(I)
'                                               リストの列と同じ順番にすること
'       選択した行の列に対して、指定したコントロールに値を設定し
'       そのコントロールに[AfterUpdate]イベントがあれば呼び出す
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub ComboBox_AfterUpdate(cbo As ComboBox, ParamArray CtrlNames())
    On Error GoTo ErrorHandler
    Dim i                   As Long
    Dim j                   As Long
    Dim frm                 As Form
    
    Set frm = GetFormObject(cbo)
    j = cbo.ListIndex
    For i = LBound(CtrlNames) To UBound(CtrlNames)
        If CtrlNames(i) <> "" Then                                              'INSERT 2007/03/19 K.ISHIZAKA
            With frm.Controls(CtrlNames(i))
                .VALUE = IIf(j >= 0, cbo.Column(i + 1, j), "")
'>>> DELETE START 2007/02/26 K.ISHIZAKA >>>
'            If .AfterUpdate = "[Event Procedure]" Then
'                On Error Resume Next
'                CallByName frm, .Name & "_AfterUpdate", VbMethod
'                If Err.Number <> 0 Then
'                    On Error GoTo ErrorHandler
'                    Call MSZZ024_M10(.Name & "_AfterUpdate", "イベントプロシージャ[" & .Name & "_AfterUpdate]をPublicで宣言してください。")
'                End If
'                On Error GoTo ErrorHandler
'            End If
'<<< DELETE END   2007/02/26 K.ISHIZAKA <<<
            End With
        End If                                                                  'INSERT 2007/03/19 K.ISHIZAKA
    Next
'>>> INSERT START 2007/02/26 K.ISHIZAKA >>>
    For i = LBound(CtrlNames) To UBound(CtrlNames)
        If CtrlNames(i) <> "" Then                                              'INSERT 2007/03/19 K.ISHIZAKA
            With frm.Controls(CtrlNames(i))
                If .AfterUpdate = "[Event Procedure]" Then
                    On Error Resume Next
                    CallByName frm, .NAME & "_AfterUpdate", VbMethod
                    If Err.Number <> 0 Then
                        On Error GoTo ErrorHandler
                        Call MSZZ024_M10(.NAME & "_AfterUpdate", "イベントプロシージャ[" & .NAME & "_AfterUpdate]をPublicで宣言してください。")
                    End If
                    On Error GoTo ErrorHandler
                End If
            End With
        End If                                                                  'INSERT 2007/03/19 K.ISHIZAKA
    Next
'<<< INSERT END   2007/02/26 K.ISHIZAKA <<<
Exit Sub
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "ComboBox_AfterUpdate" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : リストにない入力をしたとき
'       MODULE_ID       : ComboBox_NotInList
'       CREATE_DATE     : 2007/02/22            K.ISHIZAKA
'       PARAM           : cbo                   区分などのコンボボックス(I)
'                       : NewData               イベントプロシージャの同名引数(I/O)
'                       : Response              イベントプロシージャの同名引数(I/O)
'       コンボボックスのプロパティ「入力チェック」が「はい」になってること
'       メッセージボックスのタイトルをフォームのキャプションで表示し
'       前回の値に戻す
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub ComboBox_NotInList(cbo As ComboBox, NewData As String, Response As Integer)
    On Error GoTo ErrorHandler
    Dim frm                 As Form
    
    Set frm = GetFormObject(cbo)
    MsgBox "リストから選択してください", vbOKOnly + vbInformation, frm.Caption
    Response = acDataErrContinue
    cbo.Undo
Exit Sub
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "ComboBox_NotInList" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォーム取得
'       MODULE_ID       : GetFormObject
'       CREATE_DATE     : 2007/02/22            K.ISHIZAKA
'       PARAM           : ctrl                  コントロール(I)
'       RETURN          : コントロールのフォーム(Form)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetFormObject(ctrl As Object) As Form
    If Left(UCase(TypeName(ctrl)), 5) = "FORM_" Then
        Set GetFormObject = ctrl
    Else
        Set GetFormObject = GetFormObject(ctrl.Parent)
    End If
End Function

'****************************  ended or program ********************************
