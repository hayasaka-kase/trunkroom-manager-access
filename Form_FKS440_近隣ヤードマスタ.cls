VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS440_近隣ヤードマスタ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　（株）加瀬倉庫
'   システム名　　　：　倉庫管理システム
'**********************************************
'   プログラム名　　：　近隣ヤードマスタ
'   プログラムＩＤ　：
'   作　成　日　　　：  2005/10/24
'   作　成　者　　　：  イーグルソフト
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/03/27
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  緯度経度対応
'   Ver             ：  0.9
'**********************************************
'修正履歴
'   修　正　日　　　：　2006/06/24
'   修　正　者　　　：　tajima
'   修　正　内　容　：  修繕呼出対応
'   Ver             ：　1.0
'**********************************************
Option Compare Database
Option Explicit

Private pobjDB As Database

Private Const strmsg重複近隣        As String = "近隣ﾔｰﾄﾞｺｰﾄﾞが重複しています。"
Private Const strMsg_kaku           As String = "確認"
Private Const strmsg自ヤード        As String = "自ヤードコードは入力できません。"
Private Const strMSGデータなし      As String = "存在しないコードです。"
Private Const strMsg_err            As String = "エラー"

Private Const pcintMINUS            As Integer = 45      ' マイナス符号         'INSERT 2006/03/27 K.ISHIZAKA

'==============================================================================*
'
'        MODULE_NAME      :Form_Loadイベント
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    Set pobjDB = CurrentDb
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Closeイベント
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    pobjDB.Close
    Set pobjDB = Nothing
End Sub


'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDownイベント
'        MODULE_ID        :Form_KeyDown
'        CREATE_DATE      :2005/11/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   Call CmReprMod.OpenFVS90n(KeyCode, Me.cmb_近隣コード)  ' 2006/06/24 Add tajima

    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.Parent.cmd_終了.SetFocus

        Case vbKeyF2
            KeyCode = 0
            If Me.Parent.opg_処理選択.VALUE = 3 Then
                Me.Parent.cmd_削除.SetFocus
            Else
                Me.Parent.cmd_保存.SetFocus
            End If

' 2005/11/21 ADD T.SUZUKI Start
        Case vbKeyF5
            ' F5キー押下により価格表(上段)へフォーカス移動
            On Error Resume Next
            Me.Parent.txt_Top_Dummy.SetFocus
'            Call Parent.cmd_Bottom_Click
' 2005/11/21 ADD T.SUZUKI End

        Case vbKeyF3, vbKeyF6 To vbKeyF12            ' 2005/11/21 ADD T.SUZUKI
'        Case vbKeyF3, vbKeyF6 To vbKeyF12, vbKeyF5  ' 2005/11/21 DEL T.SUZUKI
            KeyCode = 0

        Case Else
            Exit Sub
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_近隣コード_AfterUpdateイベント
'        MODULE_ID        :cmb_近隣コード_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_近隣コード_AfterUpdate()

    Dim vCode   As Variant
    Dim strSQL  As String
    Dim objRs   As Recordset

    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_近隣コード, 0)

    '何も入力されていない時は、処理をしない
    If vCode = 0 Then
        Me.cmb_近隣コード = "000000"
        Me.txt_近隣名 = ""
        Me.txt_近隣住所 = ""
        Me.txt_NYAR_IDO = ""                                                    'INSERT 2006/03/27 K.ISHIZAKA
        Me.txt_NYAR_KEIDO = ""                                                  'INSERT 2006/03/27 K.ISHIZAKA
        Me.txt_NYAR_KIRO = ""                                                   'INSERT 2006/03/27 K.ISHIZAKA
        Exit Sub
    End If

    If Me.Parent.opg_処理選択 = 1 Or Me.Parent.opg_処理選択 = 2 Then
        'データの呼び出し
        strSQL = DSP_MODEUPDATE_KINRIN_YARD(vCode)
        Set objRs = pobjDB.OpenRecordset(strSQL, dbOpenForwardOnly)
        If Not (objRs.BOF And objRs.EOF) Then
           '画面に値をセット
'            Me.cmb_近隣コード = Format(Nz(objRs(0), ""), "000000")             'DELETE 2006/03/27 K.ISHIZAKA
'            Me.txt_近隣名 = Nz(objRs(1), "")                                   'DELETE 2006/03/27 K.ISHIZAKA
'            Me.txt_近隣住所 = Nz(objRs(2), "") & Nz(objRs(3), "")              'DELETE 2006/03/27 K.ISHIZAKA
'>>> INSERT START 2006/03/27 K.ISHIZAKA >>>
            Me.cmb_近隣コード = Format(Nz(objRs("YARD_CODE"), ""), "000000")
            Me.txt_近隣名 = Nz(objRs("YARD_NAME"), "")
            Me.txt_近隣住所 = Nz(objRs("YARD_ADDR_1"), "") & Nz(objRs("YARD_ADDR_2"), "") & Nz(objRs("YARD_ADDR_3"), "")
            Me.txt_NYAR_IDO = Nz(objRs("YARD_IDO"), "")
            Me.txt_NYAR_KEIDO = Nz(objRs("YARD_KEIDO"), "")
            If (Nz(Me.Parent.txt_YARD_ZAHYOKEI, "") <> "") And _
                (Nz(Me.Parent.txt_YARD_IDO, "") <> "") And _
                (Nz(Me.Parent.txt_YARD_KEIDO, "") <> "") And _
                (Nz(Me.txt_NYAR_IDO, "") <> "") And _
                (Nz(Me.txt_NYAR_KEIDO, "") <> "") Then
                Me.txt_NYAR_KIRO = Format(MNN020.KmFromDo( _
                    CInt(Me.Parent.txt_YARD_ZAHYOKEI), _
                    CDbl(Me.Parent.txt_YARD_IDO), _
                    CDbl(Me.Parent.txt_YARD_KEIDO), _
                    CDbl(Me.txt_NYAR_IDO), _
                    CDbl(Me.txt_NYAR_KEIDO)), "0.0000")
            Else
                Me.txt_NYAR_KIRO = ""
            End If
'<<< INSERT END   2006/03/27 K.ISHIZAKA <<<
        End If

        doCmd.Hourglass False
        doCmd.SetWarnings True
        Set objRs = Nothing
    End If
    Me.cmb_近隣コード = Format(Me.cmb_近隣コード, "000000")

    Parent.blnDataUpdate = True  ' 2005/11/08 ADD T.SUZUKI

    ' コンテナ番号を入力した際、レコードを確定する
    doCmd.Echo False
    Me.Requery
    doCmd.GoToRecord , , acLast
    doCmd.Echo True

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_近隣コード_BeforeUpdateイベント
'        MODULE_ID        :cmb_近隣コード_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_近隣コード_BeforeUpdate(Cancel As Integer)

    Dim strSQL      As String
    Dim vCode       As Variant

    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_近隣コード, 0)

    '何も入力されていない時は、処理をしない
    If vCode = 0 Then Exit Sub

    If Me.Parent.opg_処理選択 = 1 Or Me.Parent.opg_処理選択 = 2 Then
        '存在チェック
        If CHK_YARD_CODE_EXIST_Kinrin(Me.cmb_近隣コード) = False Then
            If vCode <> 0 Then
                Cancel = True
                SendKeys "{ESC}"
                Exit Sub
            End If
        End If

        '重複チェック(近隣コード)
        If vCode <> 0 And Format(vCode, "000000") <> Me.Parent.cmb_ヤードコード Then
            If CHK_KINRIN_CODE_EXIST(Format(Me.Parent.cmb_ヤードコード, "000000"), Format(vCode, "000000")) Then  ' 2005/11/08 ADD T.SUZUKI
'            If CHK_KINRIN_CODE_EXIST(Format(vCode, "000000")) Then                                              ' 2005/11/08 DEL T.SUZUKI
                MsgBox strmsg重複近隣, vbInformation + vbOKOnly, strMsg_kaku
                Cancel = True
                SendKeys "{ESC}"
                Exit Sub
            End If

        ElseIf vCode <> 0 And Format(vCode, "000000") = Me.Parent.cmb_ヤードコード Then
            ' メインヤードを選択した場合
            MsgBox strmsg自ヤード, vbInformation + vbOKOnly, strMsg_kaku
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        End If
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_近隣コード_KeyPressイベント
'        MODULE_ID        :cmb_近隣コード_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_近隣コード_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_近隣コード, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CHK_YARD_CODE_EXIST_Kinrin
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'*****************************************************
'存在チェック（近隣ヤードコード用）Add 01-13-2000 Araki
'ヤードコードがＤＢに既に存在するかチェックする。
'True=存在する、False=存在しない
'*****************************************************
Private Function CHK_YARD_CODE_EXIST_Kinrin(ByVal lcode As Long) As Boolean

    Dim stWHERE As String
    Dim vRet    As Variant

    On Error GoTo err_rtn

    CHK_YARD_CODE_EXIST_Kinrin = False

    stWHERE = ""
    stWHERE = stWHERE & " YARD_CODE = " & (lcode)

    vRet = Nz(DLookup("YARD_CODE", "dbo_YARD_MAST", stWHERE), "")

    If vRet = "" Then
        MsgBox strMSGデータなし, vbInformation, strMsg_err
        GoTo Exit_rtn
    End If

    CHK_YARD_CODE_EXIST_Kinrin = True

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CHK_KINRIN_CODE_EXIST
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'存在チェック
'近隣コードの重複チェック
'True=存在する、False=存在しない
'**********************************************
Private Function CHK_KINRIN_CODE_EXIST(ByVal strYcode As String, ByVal strNcode As String) As Boolean  ' 2005/11/08 ADD T.SUZUKI
'Private Function CHK_KINRIN_CODE_EXIST(ByVal strNCODE As String) As Boolean                           ' 2005/11/08 DEL T.SUZUKI

    Dim varRtn    As Variant
    Dim strWhere  As String

    On Error GoTo err_rtn

    CHK_KINRIN_CODE_EXIST = False

    strWhere = ""
    strWhere = strWhere & " NYAR_YCODE = " & strYcode   ' 2005/11/08 ADD T.SUZUKI
'    strWHERE = strWHERE & " NYAR_YCODE = " & strNCODE  ' 2005/11/08 DEL T.SUZUKI
    strWhere = strWhere & " AND NYAR_NCODE = " & strNcode

    varRtn = Nz(DLookup("NYAR_YCODE", "FKS440_W01", strWhere), "")

    ' 該当データが存在する場合エラーとする
    If varRtn <> "" Then
        CHK_KINRIN_CODE_EXIST = True
        GoTo Exit_rtn
    End If

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :DSP_MODEUPDATE_KINRIN_YARD
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'近隣コードで検索されたレコードを画面にセットする
'パラメータ：近隣コード
'戻り値：ＳＱＬ
'**********************************************
Private Function DSP_MODEUPDATE_KINRIN_YARD(ByVal vCode As Variant) As String

    Dim strSQL As String

    On Error GoTo err_rtn

    DSP_MODEUPDATE_KINRIN_YARD = ""

    If vCode = "" Then
        DSP_MODEUPDATE_KINRIN_YARD = ""
        GoTo Exit_rtn
    End If

    doCmd.Hourglass True
    doCmd.SetWarnings False

    'ﾔｰﾄﾞﾏｽﾀ
    strSQL = ""
    strSQL = "SELECT"
    strSQL = strSQL & " dbo_YARD_MAST.YARD_CODE,"       '0  ﾔｰﾄﾞｺｰﾄﾞ
    strSQL = strSQL & " dbo_YARD_MAST.YARD_NAME,"       '1  ﾔｰﾄﾞ名
    strSQL = strSQL & " dbo_YARD_MAST.YARD_ADDR_1,"     '2  住所１
    strSQL = strSQL & " dbo_YARD_MAST.YARD_ADDR_2,"     '3  住所２
    strSQL = strSQL & " dbo_YARD_MAST.YARD_ADDR_3,"     '4  住所３
    strSQL = strSQL & " dbo_YARD_MAST.YARD_IDO,"        '5  緯度                'INSERT 2006/03/27 K.ISHIZAKA
    strSQL = strSQL & " dbo_YARD_MAST.YARD_KEIDO "      '6  経度                'INSERT 2006/03/27 K.ISHIZAKA
    strSQL = strSQL & " FROM dbo_YARD_MAST"
    strSQL = strSQL & " WHERE dbo_YARD_MAST.YARD_CODE = " & (vCode)

    DSP_MODEUPDATE_KINRIN_YARD = strSQL

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      : txt_NYAR_KIRO_KeyPressイベント
'        MODULE_ID        : txt_NYAR_KIRO_KeyPress
'        CREATE_DATE      : 2006/03/27
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_NYAR_KIRO_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = pcintMINUS Then KeyAscii = 0: Exit Sub

    If Me.Parent.fncTxtInputCheck(txt_NYAR_KIRO, 999, 7) = False Then
        KeyAscii = 0: Exit Sub
    End If

    Call ETC_CUREN_RTN(txt_NYAR_KIRO, 3, 4, KeyAscii)
End Sub

'==============================================================================*
'
'        MODULE_NAME      : cmb_近隣コード_DblClickイベント
'        MODULE_ID        : cmb_近隣コード_DblClick
'        CREATE_DATE      : 2006/06/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_近隣コード_DblClick(Cancel As Integer)
   Call Form_KeyDown(CmReprMod.P_呼出_ヤード, 0)
End Sub

'****************************  ended or program ********************************
