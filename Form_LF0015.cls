VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LF0015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 経営支援サブシステム
'
'        PROGRAM_NAME    :
'        PROGRAM_ID      : LF0015
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2001/05/11
'        CERATER         : N.MIURA
'
'        UPDATE          : 2003/12/09
'        UPDATER         : N.MIURA
'                        : バグ修正
'
'        UPDATE          : 2006/08/30
'        UPDATER         : N.MIURA
'        Ver             : 0.2
'                        : 出力条件変更
'
'        UPDATE          : 2011/07/21
'        UPDATER         : M.RYU
'        Ver             : 0.3
'                          ワークテーブルWK_MQLY010_TEMP変更⇒WK_LQ0052_TEMP
'                          フォームMFLY010変更⇒LF0052
'
'        UPDATE          : 2019/07/11
'        UPDATER         : M.HONDA
'        Ver             : 0.4
'                          段区分を追加 + SQL書き直し
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "LF0015"
'
Private USER_ID         As String
Private MODL_ID         As String
'
Private dbs             As Database
Private strSQL          As String
'
Private INPCHEK_STS     As Boolean
'
Private stDocName       As String
Private stLinkCriteria  As String
'
Private WK_BUMOC        As String
Private WK_BUMON        As String                       'INSERT 2011/07/21 M.RYU

Private Const WK_TEMP   As String = "WK_LQ0052_TEMP"    'INSERT 2011/07/21 M.RYU
'
'==============================================================================*
'
'        MODULE_NAME      :フォームのオープン
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Form_Open
    
    Call MSHS010_M00(PROG_ID, "1")
    
    Call LF0015_DBOPEN
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル44.BackColor = RGB(255, 0, 204)
    End Select
    USER_ID = LsGetUserName()

Exit_Form_Open:
    Exit Sub

Err_Form_Open:
    MsgBox Err.Description
    Resume Exit_Form_Open
End Sub
'==============================================================================*
'
'        MODULE_NAME      :実行
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0015_EXE1B_Click()
    On Error GoTo Err_LF0015_EXE1B_Click
    
    If (MsgBox("実行しますか？", 4) = 7) Then
        Exit Sub
    End If
    
   ' Call LF0015_DROPTB
    Call LF0015_SELECT
    
'    stDocName = "MFLY010"      'DELETE 2011/07/21 M.RYU
    stDocName = "LF0052"        'INSERT 2011/07/21 M.RYU
    doCmd.OpenForm stDocName, , , stLinkCriteria

Exit_LF0015_EXE1B_Click:
    Exit Sub

Err_LF0015_EXE1B_Click:
    MsgBox Err.Description
    Resume Exit_LF0015_EXE1B_Click
End Sub
'==============================================================================*
'
'        MODULE_NAME      :画面最大表示
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    On Error GoTo Err_Form_Activate
    
    doCmd.Maximize

Exit_Form_Activate:
    Exit Sub

Err_Form_Activate:
    MsgBox Err.Description
    Resume Exit_Form_Activate
End Sub
'==============================================================================*
'
'        MODULE_NAME      :閉じる
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0015_EXITB_Click()
    On Error GoTo Err_LF0015_EXITB_Click

    Call LF0015_DBCLOS
    
    Call MSHS010_M00(PROG_ID, "2")
    
    doCmd.Close

Exit_LF0015_EXITB_Click:
    Exit Sub

Err_LF0015_EXITB_Click:
    MsgBox Err.Description
    Resume Exit_LF0015_EXITB_Click
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルの削除
'        MODULE_ID        :LF0015_DROPTB
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0015_DROPTB()
    On Error GoTo Err_LF0015_DROPTB
    
    doCmd.Echo True, "ワークテーブルの削除"
    
'    strSQL = ""                                'DELETE 2011/07/21 M.RYU-------<S>
'    strSQL = strSQL & "DROP TABLE "
'    strSQL = strSQL & "WK_MQLY010_TEMP "
'    strSQL = strSQL & ";"                      'DELETE 2011/07/21 M.RYU-------<E>

    strSQL = " DELETE FROM " & WK_TEMP & " ; "  'INSERT 2011/07/21 M.RYU
    
    dbs.Execute strSQL, dbFailOnError
 
Exit_LF0015_DROPTB:
    Exit Sub

Err_LF0015_DROPTB:
    MsgBox Err.Description
    Resume Exit_LF0015_DROPTB
End Sub
'==============================================================================*
'
'        MODULE_NAME      :抽出
'        MODULE_ID        :LF0015_SELECT
'        CREATE_DATE      :2006/08/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0015_SELECT()
    On Error GoTo Err_LF0015_SELECT

    Dim strBumonCode As String
    Dim strDataSource As String
    Dim lngCnt As Integer
    Dim rsSqlServer         As Object       'ADODB.Recordset
    Dim objCon              As Object       'ADODB.Connection
    Dim iSetp               As Integer
    
    MODL_ID = "LF0015_SELECT"
   
    doCmd.Echo True, "契約データの抽出"
    
    '部門コード取得
    WK_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1")
    WK_BUMON = DLookup("BUMOM_BUMON", "BUMO_MAST", "BUMOM_BUMOC = '" & WK_BUMOC & "'")      'INSERT 2011/07/21 M.RYU
    
 '   doCmd.Hourglass True
    
   'SQL-Server接続
    Set objCon = ADODB_Connection(WK_BUMOC)
    
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT "
    strSQL = strSQL & " '" & WK_BUMOC & "'       AS BUMOC,  "
    strSQL = strSQL & " '" & WK_BUMON & "'       AS BUMON,  "
    strSQL = strSQL & "     RIGHT('000000' + CONVERT(VARCHAR(6), CARG_UCODE),6) AS KOKYC, "
    strSQL = strSQL & "     USER_NAME AS KOKYN, "
    strSQL = strSQL & "     USER_TANM AS KOTANTN, "
    strSQL = strSQL & "     'FALSE' AS FLGSEL "
    strSQL = strSQL & "FROM  "
    strSQL = strSQL & "     CARG_FILE "
    strSQL = strSQL & "     INNER Join "
    strSQL = strSQL & "     USER_MAST ON "
    strSQL = strSQL & "             CARG_FILE.CARG_UCODE = USER_CODE AND "
    strSQL = strSQL & "             CARG_FILE.CARG_UCODE NOT BETWEEN  '720000' AND '729999' "
    
    strSQL = strSQL & "         INNER JOIN CNTA_MAST ON "
    strSQL = strSQL & "             CNTA_CODE = CARG_YCODE AND "
    strSQL = strSQL & "             CNTA_NO = CARG_NO "
    strSQL = strSQL & "     WHERE "
    strSQL = strSQL & "CARG_YCODE        = " & LF0015_YARDC & " AND "
    
    strSQL = strSQL & " ISNULL(CARG_KYDATE,'2999/01/01')        > '" & DateSerial(Year(DATE), Month(DATE), 0) & "' AND "
    
    
    strSQL = strSQL & "USER_SKBN   BETWEEN " & LF0015_SKB1I & " AND " & LF0015_SKB2I & " AND "
    strSQL = strSQL & "CARG_AGRE <> 9  "
    If Me.chk_sita = -1 And Me.chk_ue = -1 Then
    ElseIf Me.chk_sita = -1 Then
        strSQL = strSQL & "    AND     CNTA_STEP = 1 "
    ElseIf Me.chk_ue = -1 Then
        strSQL = strSQL & "   AND      CNTA_STEP = 0 "
    End If
    
    
    lngCnt = ADODB_CreateWorkTable_CurrentDb(WK_TEMP, strSQL, objCon)

Exit_LF0015_SELECT:
    Exit Sub

Err_LF0015_SELECT:
    MsgBox Err.Description
    Resume Exit_LF0015_SELECT
End Sub
'==============================================================================*
'
'        MODULE_NAME      :抽出
'        MODULE_ID        :LF0015_SELECT
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20060830 N.MIURA
'Private Sub LF0015_SELECT()
'    On Error GoTo Err_LF0015_SELECT
'    MODL_ID = "LF0015_SELECT"
'    doCmd.Echo True, "契約データの抽出"
'    WK_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1")
'    strSQL = ""
'    strSQL = strSQL & "SELECT "
'    strSQL = strSQL & "'" & WK_BUMOC & "'          AS BUMOC, "
'    strSQL = strSQL & "Format(CARG_UCODE,'000000') AS KOKYC  "
'    strSQL = strSQL & "INTO "
'    strSQL = strSQL & "WK_MQLY010_TEMP "
'    strSQL = strSQL & "FROM "
'    strSQL = strSQL & "dbo_CARG_FILE "
'    strSQL = strSQL & "INNER JOIN dbo_USER_MAST ON dbo_CARG_FILE.CARG_UCODE = dbo_USER_MAST.USER_CODE "
'    strSQL = strSQL & "WHERE "
'    strSQL = strSQL & "CARG_YCODE        = " & LF0015_YARDC & " AND "
'    strSQL = strSQL & "USER_SKBN   BETWEEN " & LF0015_SKB1I & " AND " & LF0015_SKB2I & " AND "
'    strSQL = strSQL & "CARG_KYDATE IS NULL "
'    strSQL = strSQL & "GROUP BY "
'    strSQL = strSQL & "CARG_UCODE "
'    strSQL = strSQL & "ORDER BY "
'    strSQL = strSQL & "CARG_UCODE "
'    strSQL = strSQL & ";"
'    DBS.Execute strSQL, dbFailOnError
'Exit_LF0015_SELECT:
'    Exit Sub
'Err_LF0015_SELECT:
'    MsgBox Err.Description
'    Resume Exit_LF0015_SELECT
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＯＰＥＮ
'        MODULE_ID        :LF0015_DBOPEN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0015_DBOPEN()
    On Error GoTo Err_LF0015_DBOPEN
    
    Set dbs = CurrentDb()
    
Exit_LF0015_DBOPEN:
    Exit Sub

Err_LF0015_DBOPEN:
    MsgBox Err.Description
    Resume Exit_LF0015_DBOPEN
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＣＬＯＳＥ
'        MODULE_ID        :LF0015_DBCLOS
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0015_DBCLOS()
    On Error GoTo Err_LF0015_DBCLOS
    
    dbs.Close
    
    Set dbs = Nothing

Exit_LF0015_DBCLOS:
    Exit Sub

Err_LF0015_DBCLOS:
    MsgBox Err.Description
    Resume Exit_LF0015_DBCLOS
End Sub
'****************************  ended or program ********************************

