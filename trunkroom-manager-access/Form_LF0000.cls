VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LF0000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 分析帳票
'
'        PROGRAM_NAME    : 契約状況一覧表
'        PROGRAM_ID      : LF0000
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2004/03/02
'        CERATER         : N.MIURA
'        Ver             : 0.0
'
'        UPDATE          : 2004/06/26
'        UPDATER         : N.MIURA
'        Ver             : 0.1
'                        :
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "LF0000"
'
Private USER_ID         As String
Private MODL_ID         As String
'
Private dbs             As Database
Private KONT_DB         As Database
'
Private strSQL          As String
'
Private RST_CNTA        As Recordset
Private RST_Q000        As Recordset
'
Private SET_BUMOC       As String
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
    
    SET_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    
    Call LF0000_DBOPEN
    
    USER_ID = LsGetUserName()
    
    Call LF0000_CLEA0

Exit_Form_Open:
    Exit Sub

Err_Form_Open:
    MsgBox Err.Description
    Resume Exit_Form_Open
End Sub
'==============================================================================*
'
'        MODULE_NAME      :実行
'        MODULE_ID        :LF0000_EXE1B_Click
'        CREATE_DATE      :2004/06/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0000_EXE1B_Click()
    On Error GoTo Err_LF0000_EXE1B_Click

    'If LF0000_INCHK0 = True Then
    
       If (MsgBox("実行しますか？", vbYesNo) = vbYes) Then
        
           Call LF0000_DROPTB
           Call LF0000_SELECT
    
           Select Case LF0000_LISTG
                  Case 0
                       doCmd.OpenReport "LR0000_M00", acPreview, "", ""
                  Case 1
                       doCmd.OpenReport "LR0000_M10", acPreview, "", ""
           End Select
   
       End If

    'End If

Exit_LF0000_EXE1B_Click:
    Exit Sub

Err_LF0000_EXE1B_Click:
    MsgBox Err.Description
    Resume Exit_LF0000_EXE1B_Click
End Sub
'==============================================================================*
'
'        MODULE_NAME      :EXCEL
'        MODULE_ID        :LF0000_EXE2B_Click
'        CREATE_DATE      :2004/06/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0000_EXE2B_Click()
    On Error GoTo Err_LF0000_EXE2B_Click

    'If LF0000_INCHK0 = True Then
    
       If (MsgBox("実行しますか？", vbYesNo) = vbYes) Then
        
           Call LF0000_DROPTB
           Call LF0000_SELECT
    
           doCmd.OutputTo acQuery, "LQ0000_M00", "MicrosoftExcel(*.xls)", "", True, ""
    
       End If

    'End If

Exit_LF0000_EXE2B_Click:
    Exit Sub

Err_LF0000_EXE2B_Click:
    MsgBox Err.Description
    Resume Exit_LF0000_EXE2B_Click
End Sub
'==============================================================================*
'
'        MODULE_NAME      :実行
'        MODULE_ID        :LF0000_EXE1B_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 200040625 N.MIURA
'Private Sub LF0000_EXE1B_Click()
'    On Error GoTo Err_LF0000_EXE1B_Click
'    Select Case LF0000_LISTG
'           Case 0
'                doCmd.OpenReport "LR0000_M00", acPreview, "", ""
'           Case 1
'                doCmd.OpenReport "LR0000_M10", acPreview, "", ""
'    End Select
'Exit_LF0000_EXE1B_Click:
'    Exit Sub
'Err_LF0000_EXE1B_Click:
'    MsgBox err.Description
'    Resume Exit_LF0000_EXE1B_Click
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :EXCEL
'        MODULE_ID        :LF0000_EXE2B_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20040625 N.MIURA
'Private Sub LF0000_EXE2B_Click()
'    On Error GoTo Err_LF0000_EXE2B_Click
'    'doCmd.OutputTo acQuery, "LQ0000_M00", "MicrosoftExcel(*.xls)", "", True, ""
'    doCmd.OutputTo acOutputTable, "WK_LQ0000_TEMP", "MicrosoftExcel(*.xls)", "", True, ""
'Exit_LF0000_EXE2B_Click:
'    Exit Sub
'Err_LF0000_EXE2B_Click:
'    MsgBox err.Description
'    Resume Exit_LF0000_EXE2B_Click
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :閉じる
'        MODULE_ID        :LF0000_EXITB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0000_EXITB_Click()
    On Error GoTo Err_LF0000_EXITB_Click

    'Call LF0000_DBCLOS                                                         'DELETE 20040625 N.MIURA
    'Call MSHS010_M00(PROG_ID, "2")                                             'DELETE 20040625 N.MIURA
    
    doCmd.Close

Exit_LF0000_EXITB_Click:
    Exit Sub

Err_LF0000_EXITB_Click:
    MsgBox Err.Description
    Resume Exit_LF0000_EXITB_Click
End Sub
'==============================================================================*
'
'        MODULE_NAME      :フォームクローズ
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :2004/06/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    Call LF0000_DBCLOS
    Call MSHS010_M00(PROG_ID, "2")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :初期値
'        MODULE_ID        :LF0000_CLEA0
'        CREATE_DATE      :2004/06/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0000_CLEA0()
    
    With Me
    .LF0000_YAR1C = 0
    .LF0000_YAR2C = 999999
    .LF0000_KAIYG = 0
    .LF0000_LISTG = 0
    End With

End Sub
'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルの削除
'        MODULE_ID        :LF0000_DROPTB
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0000_DROPTB()
    On Error GoTo Err_LF0000_DROPTB
    
    strSQL = "DROP TABLE WK_LQ0000_TEMP"
    Call LF0000_SQLEXEC(strSQL)
    
    strSQL = "CREATE TABLE WK_LQ0000_TEMP "
    strSQL = strSQL & "("
    strSQL = strSQL & "LQ0000_YARDC Integer,  "
    strSQL = strSQL & "LQ0000_YARDN TEXT(36), "
    strSQL = strSQL & "LQ0000_KONTC Integer,  "
    strSQL = strSQL & "LQ0000_STEPI Integer,  "
    strSQL = strSQL & "LQ0000_SYUKI Integer,  "
    strSQL = strSQL & "LQ0000_KOKYC Integer,  "
    strSQL = strSQL & "LQ0000_KOKYN TEXT(36), "
    strSQL = strSQL & "LQ0000_KAISD TEXT(08), "
    strSQL = strSQL & "LQ0000_MANRD TEXT(08), "
    strSQL = strSQL & "LQ0000_KAIYD TEXT(08), "
    strSQL = strSQL & "LQ0000_RENTA Integer,  "
    strSQL = strSQL & "LQ0000_HOSYG Integer,  "
    strSQL = strSQL & "LQ0000_YNAME TEXT(36)  "
    strSQL = strSQL & ");"
    Call LF0000_SQLEXEC(strSQL)
    
Exit_LF0000_DROPTB:
    Exit Sub

Err_LF0000_DROPTB:
    MsgBox Err.Description
    Resume Exit_LF0000_DROPTB
End Sub
'==============================================================================*
'
'        MODULE_NAME      :抽出
'        MODULE_ID        :LF0000_SELECT
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0000_SELECT()
    On Error GoTo Err_LF0000_SELECT
    
    If Me.LF0000_KAIYG = 2 Then
       Call LF0000_SEL_CNTA("F", "A")
    Else
       Call LF0000_SEL_CNTA("F", "Z")
    End If
    
    Call LF0000_INSERT("O")

    Do Until RST_CNTA.EOF
       
       If Not RST_CNTA.Fields("CNTA_SYU") = 0 Then
          
          Select Case Me.LF0000_KAIYG
                 Case 0  '解約済明細全て除く
                      
                      If IsNull(RST_CNTA.Fields("CARG_KYDATE")) = True Then
                         
                         Call LF0000_INSERT("I")
                      
                      End If
                 
                 Case 1  '解約済明細を当月のむ含める
                      
                      'If IsNull(RST_CNTA.Fields("CARG_KYDATE")) = True Or _
                      '   Format(RST_CNTA.Fields("CARG_KYDATE"), "YYYYDD") >= Format(Now(), "YYYYMM") Then
                      If IsNull(RST_CNTA.Fields("CARG_KYDATE")) = True Or _
                         Format(RST_CNTA.Fields("CARG_KYDATE"), "YYYYMM") >= Format(Now(), "YYYYMM") Then
                         
                         Call LF0000_INSERT("I")
                      
                      End If
                 
                 Case 2  '全て含める
                      
                      Call LF0000_INSERT("I")
            
            End Select
       
       End If
       
       Call LF0000_SEL_CNTA("S", "Z")
    
    Loop
    
    RST_CNTA.Close
    RST_Q000.Close

Exit_LF0000_SELECT:
    Exit Sub

Err_LF0000_SELECT:
    MsgBox Err.Description
    Resume Exit_LF0000_SELECT
End Sub
'==============================================================================*
'
'        MODULE_NAME      :出力
'        MODULE_ID        :LF0000_INSERT
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0000_SEL_CNTA(strACCEI As String, strKBI As String) As Boolean
    
    If strACCEI = "F" Then
       If strKBI = "A" Then
          strSQL = ""
          strSQL = strSQL & "SELECT "
          strSQL = strSQL & "CNTA_CODE,   "
          strSQL = strSQL & "YARD_NAME,   "
          strSQL = strSQL & "CNTA_NO,     "
          strSQL = strSQL & "CNTA_STEP,   "
          strSQL = strSQL & "CNTA_SYU,    "
          strSQL = strSQL & "CNTA_YNAME,  "
          strSQL = strSQL & "CARG_UCODE,  "
          strSQL = strSQL & "USER_NAME,   "
          strSQL = strSQL & "CARG_STDATE, "
          strSQL = strSQL & "CARG_EDDATE, "
          strSQL = strSQL & "CARG_KYDATE, "
          strSQL = strSQL & "CARG_RENTKG, "
          strSQL = strSQL & "CARG_SECUKG  "
          strSQL = strSQL & "FROM "
          strSQL = strSQL & "(((CNTA_MAST "
          strSQL = strSQL & "INNER JOIN YARD_MAST ON    CNTA_MAST.CNTA_CODE  = YARD_MAST.YARD_CODE  )  "
          strSQL = strSQL & "LEFT  JOIN CARG_FILE ON  ( CNTA_MAST.CNTA_NO    = CARG_FILE.CARG_NO    )  "
          strSQL = strSQL & "                     AND ( CNTA_MAST.CNTA_CODE  = CARG_FILE.CARG_YCODE )) "
          strSQL = strSQL & "LEFT JOIN USER_MAST  ON    CARG_FILE.CARG_UCODE = USER_MAST.USER_CODE  )  "
          strSQL = strSQL & "WHERE "
          strSQL = strSQL & "CNTA_CODE Between " & Me.LF0000_YAR1C & " And " & Me.LF0000_YAR2C
          strSQL = strSQL & ";"
          Set RST_CNTA = KONT_DB.OpenRecordset(strSQL, dbOpenForwardOnly, dbSQLPassThrough + dbReadOnly)
       Else
          strSQL = ""
          strSQL = strSQL & "SELECT "
          strSQL = strSQL & "dbo_CNTA_MAST.CNTA_CODE, "
          strSQL = strSQL & "dbo_YARD_MAST.YARD_NAME,   "
          strSQL = strSQL & "dbo_CNTA_MAST.CNTA_NO,"
          strSQL = strSQL & "dbo_CNTA_MAST.CNTA_STEP,"
          strSQL = strSQL & "dbo_CNTA_MAST.CNTA_SYU, "
          strSQL = strSQL & "dbo_CNTA_MAST.CNTA_YNAME,  "
          strSQL = strSQL & "LQ0000_M20.CARG_UCODE, "
          strSQL = strSQL & "LQ0000_M20.USER_NAME,"
          strSQL = strSQL & "LQ0000_M20.CARG_STDATE,"
          strSQL = strSQL & "LQ0000_M20.CARG_EDDATE,"
          strSQL = strSQL & "LQ0000_M20.CARG_KYDATE,"
          strSQL = strSQL & "LQ0000_M20.CARG_RENTKG, "
          strSQL = strSQL & "LQ0000_M20.CARG_SECUKG "
          strSQL = strSQL & "FROM "
          strSQL = strSQL & "(dbo_CNTA_MAST "
          strSQL = strSQL & "LEFT JOIN LQ0000_M20     ON  ( dbo_CNTA_MAST.CNTA_CODE = LQ0000_M20.CARG_YCODE  ) "
          strSQL = strSQL & "                         AND ( dbo_CNTA_MAST.CNTA_NO   = LQ0000_M20.CARG_NO     )) "
          strSQL = strSQL & "INNER JOIN dbo_YARD_MAST ON    dbo_CNTA_MAST.CNTA_CODE = dbo_YARD_MAST.YARD_CODE   "
          strSQL = strSQL & "WHERE "
          strSQL = strSQL & "CNTA_CODE Between " & Me.LF0000_YAR1C & " And " & Me.LF0000_YAR2C
          strSQL = strSQL & ";"
          Set RST_CNTA = dbs.OpenRecordset(strSQL, dbOpenForwardOnly, dbReadOnly)
      End If
    Else
       RST_CNTA.MoveNext
    End If

End Function
'==============================================================================*
'
'        MODULE_NAME      :出力
'        MODULE_ID        :LF0000_INSERT
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0000_INSERT(strACCEI As String) As Boolean
    If strACCEI = "O" Then
       
       Set RST_Q000 = dbs.OpenRecordset("WK_LQ0000_TEMP", dbOpenDynaset, dbAppendOnly)
    
    Else
       With RST_Q000
       .AddNew
       .Fields("LQ0000_YARDC") = RST_CNTA.Fields("CNTA_CODE")
       .Fields("LQ0000_YARDN") = RST_CNTA.Fields("YARD_NAME")
       .Fields("LQ0000_KONTC") = RST_CNTA.Fields("CNTA_NO")
       .Fields("LQ0000_STEPI") = RST_CNTA.Fields("CNTA_STEP")
       .Fields("LQ0000_SYUKI") = RST_CNTA.Fields("CNTA_SYU")
       .Fields("LQ0000_KOKYC") = RST_CNTA.Fields("CARG_UCODE")
       .Fields("LQ0000_KOKYN") = RST_CNTA.Fields("USER_NAME")
       .Fields("LQ0000_KAISD") = Format(RST_CNTA.Fields("CARG_STDATE"), "YYYYMMDD")
       .Fields("LQ0000_MANRD") = Format(RST_CNTA.Fields("CARG_EDDATE"), "YYYYMMDD")
       .Fields("LQ0000_KAIYD") = Format(RST_CNTA.Fields("CARG_KYDATE"), "YYYYMMDD")
        
       If IsNull(RST_CNTA.Fields("CARG_RENTKG")) = True Then
          .Fields("LQ0000_RENTA") = 0
       Else
          .Fields("LQ0000_RENTA") = RST_CNTA.Fields("CARG_RENTKG")
       End If
        
       If IsNull(RST_CNTA.Fields("CARG_SECUKG")) = True Then
          .Fields("LQ0000_HOSYG") = 0
       Else
          .Fields("LQ0000_HOSYG") = RST_CNTA.Fields("CARG_SECUKG")
       End If
       .Fields("LQ0000_YNAME") = RST_CNTA.Fields("CNTA_YNAME")
       .UPDATE
       End With
    End If
End Function
'==============================================================================*
'
'        MODULE_NAME      :SQLの実行
'        MODULE_ID        :LF0000_SQLEXEC
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0000_SQLEXEC(strSQL As String) As Boolean
    On Error GoTo Err_LF0000_SQLEXEC

    dbs.Execute strSQL, dbFailOnError

    LF0000_SQLEXEC = True
    Exit Function

Err_LF0000_SQLEXEC:
    MsgBox Err.Description
    LF0000_SQLEXEC = False
End Function
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＯＰＥＮ
'        MODULE_ID        :LF0000_DBOPEN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0000_DBOPEN()
    On Error GoTo Err_LF0000_DBOPEN
    
    Dim strDataSource        As String

    Set dbs = CurrentDb()
 
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & SET_BUMOC & "'"), "")
    
    If strDataSource = "" Then
        MsgBox "SETU_TABLの設定が不正です。", vbExclamation, PROG_ID
        Set KONT_DB = Nothing
    Else
        Set KONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, True, MSZZ007_M00(SET_BUMOC))
    End If
    
Exit_LF0000_DBOPEN:
    Exit Sub

Err_LF0000_DBOPEN:
    MsgBox Err.Description
    Resume Exit_LF0000_DBOPEN
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＣＬＯＳＥ
'        MODULE_ID        :LF0000_DBCLOS
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0000_DBCLOS()
    On Error GoTo Err_LF0000_DBCLOS
    
    If Not dbs Is Nothing Then                                                  'INSERT 20040626 N.MIURA START
       dbs.Close
       Set dbs = Nothing
    End If                                                                      'INSERT 20040626 N.MIURA ENDED
    
    'dbs.CLOSE                                                                  'DELETE 20040626 N.MIURA
    'Set dbs = Nothing                                                          'DELETE 20040626 N.MIURA

Exit_LF0000_DBCLOS:
    Exit Sub

Err_LF0000_DBCLOS:
    MsgBox Err.Description
    Resume Exit_LF0000_DBCLOS
End Sub
'==============================================================================*
'
'        MODULE_NAME      :画面最大表示
'        MODULE_ID        :Form_Activate
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
'****************************  ended or program ********************************

