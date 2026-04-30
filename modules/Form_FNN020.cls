VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FNN020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : 緯度経度変換
'        PROGRAM_ID      : FNN020
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2006/02/13
'        CERATER         : K.ISHZIAKA
'        Ver             : 0.0
'
'        UPDATE          :
'        UPDATER         :
'        Ver             :
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const adCmdText = 1
Private Const adOpenForwardOnly = 0
Private Const adLockReadOnly = 1

'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private IEApp               As Object           'インターネット・エクスプローラー
Private strInitAddr         As String           '初期表示住所

'==============================================================================*
'
'       MODULE_NAME     : キャンセルボタン
'       MODULE_ID       : cmdCancel_Click
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler
    doCmd.Close acForm, Me.NAME
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmdCancel_Click", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 設定ボタン
'       MODULE_ID       : cmdSet_Click
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdSet_Click()
    On Error GoTo ErrorHandler
    Call MNN020.SetData
    doCmd.Close acForm, Me.NAME
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmdSet_Click", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 北へボタン
'       MODULE_ID       : cmdN_Click
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdN_Click()
    Dim dIdo                As Double
    
    On Error Resume Next
    dIdo = CDbl(txtIDOPO.VALUE)
    If Err.Number = 0 Then
        txtIDOPO.VALUE = Format(dIdo + 0.00009)
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 南へボタン
'       MODULE_ID       : cmdS_Click
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdS_Click()
    Dim dIdo                As Double
    
    On Error Resume Next
    dIdo = CDbl(txtIDOPO.VALUE)
    If Err.Number = 0 Then
        txtIDOPO.VALUE = Format(dIdo - 0.00009)
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 東へボタン
'       MODULE_ID       : cmdE_Click
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdE_Click()
    Dim dKeido              As Double
    
    On Error Resume Next
    dKeido = CDbl(txtKEIPO.VALUE)
    If Err.Number = 0 Then
        txtKEIPO.VALUE = Format(dKeido + 0.00011)
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 西へボタン
'       MODULE_ID       : cmdW_Click
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdW_Click()
    Dim dKeido              As Double
    
    On Error Resume Next
    dKeido = CDbl(txtKEIPO.VALUE)
    If Err.Number = 0 Then
        txtKEIPO.VALUE = Format(dKeido - 0.00011)
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォーム開く
'       MODULE_ID       : Form_Open
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    Dim i                   As Integer
    On Error GoTo ErrorHandler
    
    Set IEApp = Nothing
    strInitAddr = MNN020.GetData()
    i = SetItem_CmbTODON()
    OptTODON.VALUE = i
    CmbTODON.Enabled = (i <= 0)
    Call CmbTODON_AfterUpdate
    strInitAddr = ""
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("Form_Open", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 位置の確認
'       MODULE_ID       : cmdPOSITION_Click
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdPOSITION_Click()
    Dim dIdo                As Double
    Dim dKeido              As Double
    
    On Error Resume Next
    dIdo = CDbl(txtIDOPO.VALUE)
    dKeido = CDbl(txtKEIPO.VALUE)
    If Err.Number = 0 Then
        On Error GoTo ErrorHandler
'        Call MNN020.WebPointMSN(IEApp, CDbl(txtIDOPO), CDbl(txtKEIPO))
        Call MNN020.WebPointGoogle(IEApp, CDbl(txtIDOPO), CDbl(txtKEIPO), Me.txtAddress)
    End If
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmdPOSITION_Click", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 都道府県選択
'       MODULE_ID       : OptTODON_Click
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub OptTODON_Click()
    Dim i                   As Integer
    On Error GoTo ErrorHandler
    
    i = OptTODON.VALUE - 1
    With CmbTODON
        If i < 0 Then
            .Enabled = True
        Else
            .Enabled = False
            .VALUE = .ItemData(i)
            Call CmbTODON_AfterUpdate
        End If
    End With
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("OptTODON_Click", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 都道府県コンボボックス選択
'       MODULE_ID       : CmbTODON_AfterUpdate
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CmbTODON_AfterUpdate()
    Dim bClr                As Boolean
    Dim i                   As Integer
    On Error GoTo ErrorHandler
    
    bClr = Not ((strInitAddr <> "") And (Nz(txtIDOPO, "") <> "") And (Nz(txtKEIPO, "") <> ""))
    Call Clear_lstSITSN(bClr)
    i = CmbTODON.ListIndex + 1
    If SetListItem(Sql_GISAM_SITSN(i), lstSITSN, i) Then
        Call lstSITSN_AfterUpdate
    End If
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("CmbTODON_AfterUpdate", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 市区町村名リストボックス選択
'       MODULE_ID       : lstSITSN_AfterUpdate
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub lstSITSN_AfterUpdate()
    Dim bClr                As Boolean
    On Error GoTo ErrorHandler
    
    bClr = Not ((strInitAddr <> "") And (Nz(txtIDOPO, "") <> "") And (Nz(txtKEIPO, "") <> ""))
    Call Clear_lstCHINN(bClr)
    If SetListItem(Sql_GISAM_CHINN(CmbTODON.ListIndex + 1, lstSITSN.VALUE), lstCHINN) Then
        Call lstCHINN_AfterUpdate
    End If
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("lstSITSN_AfterUpdate", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 大字・町丁目リストボックス選択
'       MODULE_ID       : lstCHINN_AfterUpdate
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub lstCHINN_AfterUpdate()
    Dim bClr                As Boolean
    On Error GoTo ErrorHandler
    
    bClr = Not ((strInitAddr <> "") And (Nz(txtIDOPO, "") <> "") And (Nz(txtKEIPO, "") <> ""))
    Call Clear_lstTIBAN(bClr)
    If SetListItem(Sql_GISAM_GAIKN(CmbTODON.ListIndex + 1, lstSITSN.VALUE, lstCHINN.VALUE), lstTIBAN) Then
        If bClr Then
            Call lstTIBAN_AfterUpdate
        End If
    End If
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("lstCHINN_AfterUpdate", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 街区符号・地番リストボックス選択
'       MODULE_ID       : lstTIBAN_AfterUpdate
'       CREATE_DATE     : 2006/02/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub lstTIBAN_AfterUpdate()
    Dim dbSQLServer         As Object       'ADODB.Connection
    Dim rsSqlServer         As Object       'ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Call Clear_lstKEIBI
    'SQL-Server接続
    Set dbSQLServer = MNN020.OpenMapDB()
    On Error GoTo ErrorHandler2
    'レコードセットのオープン
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    With rsSqlServer
        .Open Sql_GISAM_ZAHYI(CmbTODON.ListIndex + 1, lstSITSN.VALUE, lstCHINN.VALUE, lstTIBAN.VALUE), _
            dbSQLServer, adOpenForwardOnly, adLockReadOnly, adCmdText
        On Error GoTo ErrorHandler3
        If Not .EOF Then
            txtKEIBI.VALUE = .Fields("GISAM_ZAHYI")
            txtIDOPO.VALUE = .Fields("GISAM_IDOPO")
            txtKEIPO.VALUE = .Fields("GISAM_KEIPO")
        End If
        .Close
    End With
    dbSQLServer.Close
Exit Sub
    
ErrorHandler3:
    rsSqlServer.Close
ErrorHandler2:
    dbSQLServer.Close
ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("lstTIBAN_AfterUpdate", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 都道府県コンボのリスト設定
'       MODULE_ID       : SetItem_CmbTODON
'       CREATE_DATE     : 2006/02/14
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SetItem_CmbTODON() As Integer
    Dim dbSQLServer         As Object       'ADODB.Connection
    Dim rsSqlServer         As Object       'ADODB.Recordset
    Dim strRowsource        As String
    Dim strwork             As String
    Dim iDaihyo             As Integer
    Dim i                   As Integer
    On Error GoTo ErrorHandler
    
    iDaihyo = 14
    SetItem_CmbTODON = 14
    'SQL-Server接続
    Set dbSQLServer = MNN020.OpenMapDB()
    On Error GoTo ErrorHandler2
    'レコードセットのオープン
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    With rsSqlServer
        .Open "SELECT TODOM_TODON FROM TODO_MAST ORDER BY TODOM_TODOI", _
            dbSQLServer, adOpenForwardOnly, adLockReadOnly, adCmdText
        On Error GoTo ErrorHandler3
        i = 0
        While Not .EOF
            i = i + 1
            strwork = Trim(.Fields(0))
            If strwork = Left(strInitAddr, Len(strwork)) Then
                strInitAddr = Mid(strInitAddr, Len(strwork) + 1)
                iDaihyo = i
                If (11 <= i) And (i <= 14) Then
                    SetItem_CmbTODON = i
                Else
                    SetItem_CmbTODON = 0
                End If
            End If
            strRowsource = strRowsource & strwork & ";"
            .MoveNext
        Wend
        .Close
        On Error GoTo ErrorHandler2
    End With
    dbSQLServer.Close
    On Error GoTo ErrorHandler
    With CmbTODON
        .RowSource = strRowsource
        .VALUE = .ItemData(iDaihyo - 1)
    End With
Exit Function
    
ErrorHandler3:
    rsSqlServer.Close
ErrorHandler2:
    dbSQLServer.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "SetItem_CmbTODON" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : リストアイテム設定
'       MODULE_ID       : SetListItem
'       CREATE_DATE     : 2006/02/14
'       PARAM           : strSQL                ＳＱＬ文(I)
'                       : lstObject             リストボックス(I)
'       RETURN          : 代表を設定した(True)／代表設定なし(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SetListItem(ByVal strSQL As String, lstObject As ListBox, Optional iNo As Integer = -1) As Boolean
    Dim dbSQLServer         As Object       'ADODB.Connection
    Dim rsSqlServer         As Object       'ADODB.Recordset
    Dim strRowsource        As String
    Dim strwork             As String
    Dim strDaihyo           As String
    On Error GoTo ErrorHandler
    
    'SQL-Server接続
    Set dbSQLServer = MNN020.OpenMapDB()
    On Error GoTo ErrorHandler2
    'レコードセットのオープン
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    With rsSqlServer
        If iNo >= 0 Then
            .Open Sql_EXISTS_GISA_MAST(iNo), dbSQLServer, adOpenForwardOnly, adLockReadOnly, adCmdText
            On Error GoTo ErrorHandler3
            If Not .EOF Then
                iNo = -1
                txtSEIBY = Nz(.Fields(0))
            Else
                txtSEIBY = "----"
            End If
            .Close
            On Error GoTo ErrorHandler2
        End If
        If iNo < 0 Then
            .Open strSQL, dbSQLServer, adOpenForwardOnly, adLockReadOnly, adCmdText
            On Error GoTo ErrorHandler3
            While Not .EOF
                strwork = Trim(.Fields(0))
                If Trim(Nz(Me.txtAddress)) <> "" Then
                    If IsNumeric(strwork) And IsNumeric(strInitAddr) Then
                        If (Val(strwork) >= Val(strInitAddr)) And (strDaihyo = "") Then
                            'strInitAddr = ""
                            strDaihyo = strwork
                        End If
                    ElseIf strwork = Left(strInitAddr, Len(strwork)) Then
                        strDaihyo = strwork
                    End If
                Else
                    If Nz(.Fields(1), "") <> "" Then
                        strDaihyo = strwork
                    End If
                End If
                strRowsource = strRowsource & strwork & ";"
                .MoveNext
            Wend
            .Close
            On Error GoTo ErrorHandler2
        End If
    End With
    dbSQLServer.Close
    On Error GoTo ErrorHandler
    With lstObject
        .RowSource = strRowsource
        .VALUE = strDaihyo
    End With
    If strDaihyo <> "" Then
        strInitAddr = Mid(strInitAddr, Len(strDaihyo) + 1)
        If Mid(strInitAddr, 2, 2) = "丁目" Then
            strInitAddr = MNN020.NumbConv(Left(strInitAddr, 1))
        End If
    End If
    SetListItem = (strDaihyo <> "")
Exit Function
    
ErrorHandler3:
    rsSqlServer.Close
ErrorHandler2:
    dbSQLServer.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "SetListItem" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : マスタ存在チェックＳＱＬ
'       MODULE_ID       : Sql_EXISTS_GISA_MAST
'       CREATE_DATE     : 2006/02/14
'       PARAM           : iNo                   都道府県番号(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Sql_EXISTS_GISA_MAST(ByVal iNo As Integer) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler

    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "    TODOM_SEIBY "
    strSQL = strSQL & "FROM   TODO_MAST "
    strSQL = strSQL & "WHERE  TODOM_TODOI = " & Format(iNo) & " "
    strSQL = strSQL & "AND  EXISTS "
    strSQL = strSQL & "     ( "
    strSQL = strSQL & "     SELECT "
    strSQL = strSQL & "         * "
    strSQL = strSQL & "     FROM   SYSOBJECTS "
    strSQL = strSQL & "     WHERE  ID = OBJECT_ID(N'[GISA_MAST_" & Format(iNo, "00") & "]') "
    strSQL = strSQL & "     AND    OBJECTPROPERTY(ID, N'ISUSERTABLE') = 1"
    strSQL = strSQL & "     ) "
    
    Sql_EXISTS_GISA_MAST = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Sql_EXISTS_GISA_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 市区町村名ＳＱＬ
'       MODULE_ID       : Sql_GISAM_SITSN
'       CREATE_DATE     : 2006/02/14
'       PARAM           : iNo                   都道府県番号(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Sql_GISAM_SITSN(ByVal iNo As Integer) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler

    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "    GISAM_SITSN, "
    strSQL = strSQL & "    MAX(GISAM_TODOC) "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    [GISA_MAST_" & Format(iNo, "00") & "] "
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "    GISAM_SITSN "
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "    GISAM_SITSN "
    
    Sql_GISAM_SITSN = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Sql_GISAM_SITSN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 大字・町丁目ＳＱＬ
'       MODULE_ID       : Sql_GISAM_CHINN
'       CREATE_DATE     : 2006/02/14
'       PARAM           : iNo                   都道府県番号(I)
'                       : strSITSN              市区町村名(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Sql_GISAM_CHINN(ByVal iNo As Integer, ByVal strSITSN As String) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler

    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "    GISAM_CHINN, "
    strSQL = strSQL & "    MAX(GISAM_SITSC) "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    [GISA_MAST_" & Format(iNo, "00") & "] "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "    GISAM_SITSN = '" & strSITSN & "' "
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "    GISAM_CHINN "
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "    GISAM_CHINN "
    
    Sql_GISAM_CHINN = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Sql_GISAM_CHINN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 街区符号・地番ＳＱＬ
'       MODULE_ID       : Sql_GISAM_GAIKN
'       CREATE_DATE     : 2006/02/14
'       PARAM           : iNo                   都道府県番号(I)
'                       : strSITSN              市区町村名(I)
'                       : strCHINN              大字・町丁目(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Sql_GISAM_GAIKN(ByVal iNo As Integer, ByVal strSITSN As String, ByVal strCHINN As String) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler

    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "    GISAM_GAIKN, "
    strSQL = strSQL & "    GISAM_CHINC "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    [GISA_MAST_" & Format(iNo, "00") & "] "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "    GISAM_SITSN = '" & strSITSN & "' "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "    GISAM_CHINN = '" & strCHINN & "' "
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "    ISNULL(GISAM_TIBAI, 99999), "
    strSQL = strSQL & "    GISAM_GAIKN "
    
    Sql_GISAM_GAIKN = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Sql_GISAM_GAIKN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 座標系番号・緯度・経度ＳＱＬ
'       MODULE_ID       : Sql_GISAM_ZAHYI
'       CREATE_DATE     : 2006/02/14
'       PARAM           : iNo                   都道府県番号(I)
'                       : strSITSN              市区町村名(I)
'                       : strCHINN              大字・町丁目(I)
'                       : strTIBAN              街区符号・地番(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Sql_GISAM_ZAHYI(ByVal iNo As Integer, ByVal strSITSN As String, ByVal strCHINN As String, ByVal strTIBAN As String) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler

    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "    GISAM_ZAHYI,"
    strSQL = strSQL & "    GISAM_IDOPO,"
    strSQL = strSQL & "    GISAM_KEIPO "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    [GISA_MAST_" & Format(iNo, "00") & "] "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "    GISAM_SITSN = '" & strSITSN & "' "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "    GISAM_CHINN = '" & strCHINN & "' "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "    GISAM_GAIKN = '" & strTIBAN & "' "
    
    Sql_GISAM_ZAHYI = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Sql_GISAM_ZAHYI" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 市区町村名クリア
'       MODULE_ID       : Clear_lstSITSN
'       CREATE_DATE     : 2006/02/14
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Clear_lstSITSN(ByVal bClr As Boolean)
    On Error GoTo ErrorHandler
    
    lstSITSN.RowSource = ""
    Call Clear_lstCHINN(bClr)
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Clear_lstSITSN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 大字・町丁目クリア
'       MODULE_ID       : Clear_lstCHINN
'       CREATE_DATE     : 2006/02/14
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Clear_lstCHINN(ByVal bClr As Boolean)
    On Error GoTo ErrorHandler
    
    lstCHINN.RowSource = ""
    Call Clear_lstTIBAN(bClr)
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Clear_lstCHINN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 街区符号・地番クリア
'       MODULE_ID       : Clear_lstTIBAN
'       CREATE_DATE     : 2006/02/14
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Clear_lstTIBAN(ByVal bClr As Boolean)
    On Error GoTo ErrorHandler
    
    lstTIBAN.RowSource = ""
    If bClr Then
        Call Clear_lstKEIBI
    End If
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Clear_lstTIBAN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 座標系番号・緯度・経度クリア
'       MODULE_ID       : Clear_lstKEIBI
'       CREATE_DATE     : 2006/02/14
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Clear_lstKEIBI()
    On Error GoTo ErrorHandler
    
    txtKEIBI.VALUE = ""
    txtIDOPO.VALUE = ""
    txtKEIPO.VALUE = ""
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Clear_lstKEIBI" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'****************************  ended or program ********************************
