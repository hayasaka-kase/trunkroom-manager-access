VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LF0003"
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
'        PROGRAM_NAME    : コンテナ分析表（使用基準）
'        PROGRAM_ID      : LF0003
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2004/03/02
'        CERATER         : N.MIURA
'        Ver             : 0.0
'
'        UPDATE          : 2004/06/25
'        UPDATER         : N.MIURA
'        Ver             : 0.1
'                        :
'
'        UPDATE          : 2004/09/25
'        UPDATER         : N.MIURA
'        Ver             : 0.2
'                        :
'
'        UPDATE          : 2004/11/11
'        UPDATER         : N.MIURA
'        Ver             : 0.3
'                        : 不具合修正
'
'        UPDATE          : 2007/04/25
'        UPDATER         : イーグルソフト 上野
'        Ver             : 0.4
'                        : ﾃﾞｰﾀをﾘﾝｸﾃｰﾌﾞﾙからでなくSQLServerより取得するよう修正
'                        : 年月項目を削除
'                        : 総件数、使用数の取得条件変更
'
'        UPDATE          : 2007/09/03
'        UPDATER         : tajima
'        Ver             : 0.5
'                        : 年月指定を復活
'
'        UPDATE          : 2008/08/01
'        UPDATER         : N.MIURA
'        Ver             : 0.6
'                        : レポートレイアウト変更
'
'        UPDATE          : 2011/11/20
'        UPDATER         : M.HONDA
'        Ver             : 0.7
'                        : 予約数を表示するように修正
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "LF0003"
Private USER_ID         As String
Private MODL_ID         As String
'
'Private dbs             As Database                                            'DELETE 20040922 N.MIURA
Private dbMDB_DB          As Database                                           'INSERT 20040922 N.MIURA
Private strSQL          As String
'
Private INPCHEK_STS     As Boolean

'変更前処理で入力した日付を保存する
Private VarChangeDate As Variant    '2007/09/03 add tajima
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
    
    Call LF0003_DBOPEN
    
    USER_ID = LsGetUserName()

    Call LF0003_CLEA0

Exit_Form_Open:
    Exit Sub

Err_Form_Open:
    MsgBox Err.Description
    Resume Exit_Form_Open
End Sub
'==============================================================================*
'
'        MODULE_NAME      :実行
'        MODULE_ID        :LF0003_EXE1B_Click
'        CREATE_DATE      :2004/09/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0003_EXE1B_Click()
    On Error GoTo Err_LF0003_EXE1B_Click
    
    If LF0003_INCHK0 = True Then
       
       If (MsgBox("実行しますか？", vbYesNo) = vbYes) Then
     
           Call LF0003_M10
           
           If LF0003_M20() > 0 Then
              
              If Me.LF0003_LISTG = 0 Then
                 doCmd.OpenReport "LR0003_M00", acPreview, "", ""
              Else
                 doCmd.OpenReport "LR0003_M10", acPreview, "", ""
              End If
           
           Else
              MsgBox ("該当データがありません。")
           End If
        
        End If
        
        Call LF0003_CLEA3
    
    End If
  
Exit_LF0003_EXE1B_Click:
    Exit Sub

Err_LF0003_EXE1B_Click:
    MsgBox Err.Description
    Resume Exit_LF0003_EXE1B_Click
End Sub
'==============================================================================*
'
'        MODULE_NAME      :実行
'        MODULE_ID        :LF0003_EXE1B_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20040925 N.MIURA
'Private Sub LF0003_EXE1B_Click()
'    On Error GoTo Err_LF0003_EXE1B_Click
'    Select Case LF003_LISTG
'           Case 0
'                doCmd.OpenReport "LR0003_M00", acPreview, "", ""
'           Case 1
'                doCmd.OpenReport "LR0003_M10", acPreview, "", ""
'    End Select
'Exit_LF0003_EXE1B_Click:
'    Exit Sub
'Err_LF0003_EXE1B_Click:
'    MsgBox err.Description
'    Resume Exit_LF0003_EXE1B_Click
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :EXCEL
'        MODULE_ID        :LF0003_EXE2B_Click
'        CREATE_DATE      :2004/09/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0003_EXE2B_Click()
    On Error GoTo Err_LF0003_EXE2B_Click
     
    If LF0003_INCHK0 = True Then
       
       If (MsgBox("実行しますか？", vbYesNo) = vbYes) Then
           
           Call LF0003_M10
           
           If LF0003_M20() > 0 Then
              doCmd.OutputTo acQuery, "LQ0003_M00", "MicrosoftExcel(*.xls)", "", True, ""
           Else
              MsgBox ("該当データがありません。")
           End If
        
        End If
        
        Call LF0003_CLEA3
    
    End If
  
Exit_LF0003_EXE2B_Click:
    Exit Sub

Err_LF0003_EXE2B_Click:
    MsgBox Err.Description
    Resume Exit_LF0003_EXE2B_Click
End Sub
'==============================================================================*
'
'        MODULE_NAME      :EXCEL
'        MODULE_ID        :LF0003_EXE2B_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20040925 N.MIURA
'Private Sub LF0003_EXE2B_Click()
'    On Error GoTo Err_LF0003_EXE2B_Click
'   doCmd.OutputTo acQuery, "LQ0001_M00", "MicrosoftExcel(*.xls)", "", True, ""
'   'doCmd.OutputTo acOutputTable, "WK_MQAM010_TEMP", "MicrosoftExcel(*.xls)", "", True, ""
'Exit_LF0003_EXE2B_Click:
'    Exit Sub
'Err_LF0003_EXE2B_Click:
'    MsgBox err.Description
'    Resume Exit_LF0003_EXE2B_Click
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :フォームクローズ
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :2004/09/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    
    Call LF0003_DROP
    Call LF0003_DBCLOS
    Call MSHS010_M00(PROG_ID, "2")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :閉じる
'        MODULE_ID        :LF0003_EXITB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0003_EXITB_Click()
    On Error GoTo Err_LF0003_EXITB_Click

    'Call LF0003_DBCLOS                                                         'DELETE 20040922 N.MIURA
    'Call MSHS010_M00(PROG_ID, "2")                                             'DELETE 20040922 N.MIURA
    
    doCmd.Close

Exit_LF0003_EXITB_Click:
    Exit Sub

Err_LF0003_EXITB_Click:
    MsgBox Err.Description
    Resume Exit_LF0003_EXITB_Click
End Sub
'==============================================================================*
'
'        MODULE_NAME      :画面クリアー０
'        MODULE_ID        :LF0003_CLEA0
'        CREATE_DATE      :2004/09/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0003_CLEA0()
    
    With Me
' 2007/04/25 DEL H.UENO Start
'''    .LF0003_YYMMD = Format(Now(), "YYYYMM")
' 2007/04/25 DEL H.UENO Start
    .LF0003_YAR1C = LF0003_YAR1C.ItemData(0)
    .LF0003_YAR2C = LF0003_YAR2C.ItemData(LF0003_YAR2C.ListCount - 1)
    .LF0003_LISTG = 0
    
    '↓INSERT 2007/09/03 tajima
    '対象年月の初期値は前月
    .LF0003_YYMMD.VALUE = Format(DateSerial(Year(Now()), Month(Now()) - 1, 1), "yyyy/mm")
    '↑INSERT 2007/09/03
    
    End With

End Sub
'==============================================================================*
'
'        MODULE_NAME      :画面クリアー３
'        MODULE_ID        :LF0003_CLEA3
'        CREATE_DATE      :2004/09/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0003_CLEA3()

' 2007/04/25 UPD H.UENO Start
'''    Me.LF0003_YYMMD.SetFocus
    Me.LF0003_YAR1C.SetFocus
' 2007/04/25 UPD H.UENO End


End Sub
'==============================================================================*
'
'        MODULE_NAME      :入力チェック０
'        MODULE_ID        :LF0003_INCHK0
'        CREATE_DATE      :2004/09/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0003_INCHK0() As Boolean
    
    LF0003_INCHK0 = True
    
    '必須入力
    If IsNull(Me.LF0003_YAR2C) Then
       Me.LF0003_YAR2C.SetFocus
       LF0003_INCHK0 = False
    End If
    If IsNull(Me.LF0003_YAR1C) Then
       Me.LF0003_YAR1C.SetFocus
       LF0003_INCHK0 = False
    End If
' 2007/04/25 DEL H.UENO Start
'''    If IsNull(Me.LF0003_YYMMD) Then
'''       Me.LF0003_YYMMD.SetFocus
'''       LF0003_INCHK0 = False
'''    End If
' 2007/04/25 DEL H.UENO End
    If Not LF0003_INCHK0 Then
       MsgBox ("必須入力項目です。")
       Exit Function
    End If
    
    '桁数
    If Len(Me.LF0003_YAR2C) <> 6 Then
       Me.LF0003_YAR2C.SetFocus
       LF0003_INCHK0 = False
    End If
    If Len(Me.LF0003_YAR1C) <> 6 Then
       Me.LF0003_YAR1C.SetFocus
       LF0003_INCHK0 = False
    End If
' 2007/04/25 DEL H.UENO Start
'''    If Len(Me.LF0003_YYMMD) <> 6 Then
'''       Me.LF0003_YYMMD.SetFocus
'''       LF0003_INCHK0 = False
'''    End If
' 2007/04/25 DEL H.UENO End
    If Not LF0003_INCHK0 Then
       MsgBox ("入力桁数が不正です。")
       Exit Function
    End If
    
' 2007/04/25 DEL H.UENO Start
'''    If (Not IsDate(Left(Me.LF0003_YYMMD, 4) & "/" & Mid(Me.LF0003_YYMMD, 5, 2) & "/01")) Then
'''        Me.LF0003_YYMMD.SetFocus
'''       LF0003_INCHK0 = False
'''    End If
'''    If Not LF0003_INCHK0 Then
'''        MsgBox "日付の入力形式が正しくありません。"
'''       Exit Function
'''    End If
' 2007/04/25 DEL H.UENO End
    
    If Me.LF0003_YAR1C > Me.LF0003_YAR2C Then
       Me.LF0003_YAR1C.SetFocus
       LF0003_INCHK0 = False
    End If
    
    If Not LF0003_INCHK0 Then
       MsgBox ("範囲指定が不正です。")
       Exit Function
    End If

End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :LF0003_M10
'        CREATE_DATE      :2004/09/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0003_M10()
    On Error GoTo Err_LF0003_M10
    
    Call LF0003_DROP
    
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE "
    strSQL = strSQL & "WK_LQ0003_CARG "
    strSQL = strSQL & "("
    strSQL = strSQL & "LQ0003_YARDC TEXT(06)  NOT NULL,"
    strSQL = strSQL & "LQ0003_URIAQ Single, "
    strSQL = strSQL & "LQ0003_URIAA Single  "
    strSQL = strSQL & ");"
    LF0003_SQLEXEC (strSQL)
    
    '2013/11/20 M.HONDA INS
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE "
    strSQL = strSQL & "WK_LQ0003_YOUK "
    strSQL = strSQL & "("
    strSQL = strSQL & "LQ0003_YARDC TEXT(06)  NOT NULL,"
    strSQL = strSQL & "LQ0003_YOUKC Single "
    strSQL = strSQL & ");"
    LF0003_SQLEXEC (strSQL)
    '2013/11/20 M.HONDA INS
    
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE "
    strSQL = strSQL & "WK_LQ0003_CNTA "
    strSQL = strSQL & "("
    strSQL = strSQL & "LQ0003_YARDC TEXT(06)  NOT NULL,"
    strSQL = strSQL & "LQ0003_SOHOQ Single, "
' 2007/04/25 ADD H.UENO Start
    strSQL = strSQL & "LQ0003_YARDN TEXT(36) "
' 2007/04/25 ADD H.UENO End
    strSQL = strSQL & ");"
    LF0003_SQLEXEC (strSQL)
    
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE "
    strSQL = strSQL & "WK_LQ0003_JARG "
    strSQL = strSQL & "("
    strSQL = strSQL & "LQ0003_YARDC TEXT(06)  NOT NULL,"
    strSQL = strSQL & "LQ0003_BUKHC TEXT(06)  NOT NULL,"
    strSQL = strSQL & "LQ0003_SIIGA Single "
    strSQL = strSQL & ");"
    LF0003_SQLEXEC (strSQL)
    
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE "
    strSQL = strSQL & "WK_LQ0003_TEMP "
    strSQL = strSQL & "("
    strSQL = strSQL & "LQ0003_YARDC TEXT(06)  NOT NULL,"
    strSQL = strSQL & "LQ0003_YARDN TEXT(36),"
    strSQL = strSQL & "LQ0003_BUKHC TEXT(06),"
    strSQL = strSQL & "LQ0003_SIIGA Single, "
    strSQL = strSQL & "LQ0003_URIAA Single, "
    strSQL = strSQL & "LQ0003_KEKKA Single, "
    strSQL = strSQL & "LQ0003_SOHOQ Single, "
    strSQL = strSQL & "LQ0003_ASSIQ Single, "
    strSQL = strSQL & "LQ0003_DASSQ Single, "
    strSQL = strSQL & "LQ0003_YOUKC Single, "  ''2013/11/20 M.HONDA INS
    strSQL = strSQL & "LQ0003_SIYOR Single  "
    strSQL = strSQL & ");"

    LF0003_SQLEXEC (strSQL)

Exit_LF0003_M10:
    Exit Function

Err_LF0003_M10:
    MsgBox Err.Description
    Resume Exit_LF0003_M10
End Function
' 2007/04/25 ADD H.UENO Start
' ﾃﾞｰﾀをﾘﾝｸﾃｰﾌﾞﾙからでなくSQLServerより取得するよう修正
''''==============================================================================*
''''
''''        MODULE_NAME      :
''''        MODULE_ID        :LF0003_M20
''''        CREATE_DATE      :2004/09/25
''''
''''==============================================================================*
''''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'''Private Function LF0003_M20() As Long
'''    On Error GoTo Err_LF0003_M20
'''
'''    Dim strMsg          As String
'''    Dim RST_TRAN        As Recordset
'''    Dim RST_WORK        As Recordset
'''
'''    strMsg = "（　契約本数・金額集計＝＞"
'''    doCmd.Echo True, PROG_ID & ":" & strMsg
'''
'''    strSQL = ""
'''    strSQL = strSQL & "SELECT "
'''    strSQL = strSQL & "CARG_YCODE, "
'''    strSQL = strSQL & "Count(CARG_YCODE), "
'''    strSQL = strSQL & "Sum(CARG_RENTKG) "
'''    strSQL = strSQL & "FROM "
'''    strSQL = strSQL & "dbo_CARG_FILE "
'''    strSQL = strSQL & "WHERE "
'''    strSQL = strSQL & " FORMAT(CARG_YCODE,'000000') BETWEEN '" & Me.LF0003_YAR1C & "' AND '" & Me.LF0003_YAR2C & "' AND "
'''    strSQL = strSQL & " Format(CARG_STDATE,'yyyymm') <= '" & Me.LF0003_YYMMD & "' AND "
'''    strSQL = strSQL & "(Format(CARG_KYDATE,'yyyymm') >= '" & Me.LF0003_YYMMD & "' Or Format(CARG_KYDATE,'yyyymm') = '') "
'''    strSQL = strSQL & "GROUP BY "
'''    strSQL = strSQL & "CARG_YCODE "
'''    strSQL = strSQL & "; "
'''    Set RST_TRAN = dbMDB_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)
'''
'''    Set RST_WORK = dbMDB_DB.OpenRecordset("WK_LQ0003_CARG", dbOpenDynaset, dbAppendOnly)
'''
'''    Do Until RST_TRAN.EOF
'''       With RST_WORK
'''       .AddNew
'''       '.Fields("LQ0003_YARDC") = RST_TRAN.Fields(0)                            'DELETE 20041111 N.MIURA
'''       .Fields("LQ0003_YARDC") = Format(RST_TRAN.Fields(0), "000000")           'INSERT 20041111 N.MIURA
'''       .Fields("LQ0003_URIAQ") = RST_TRAN.Fields(1)
'''       .Fields("LQ0003_URIAA") = RST_TRAN.Fields(2)
'''       .UPDATE
'''       End With
'''       RST_TRAN.MoveNext
'''    Loop
'''
'''    RST_TRAN.Close
'''    RST_WORK.Close
'''
'''    strMsg = strMsg & "総本数＝＞"
'''    doCmd.Echo True, PROG_ID & ":" & strMsg
'''
'''    strSQL = ""
'''    strSQL = strSQL & "SELECT "
'''    strSQL = strSQL & "CNTA_CODE, "
'''    strSQL = strSQL & "Count(CNTA_NO) "
'''    strSQL = strSQL & "FROM "
'''    strSQL = strSQL & "dbo_CNTA_MAST "
'''    strSQL = strSQL & "WHERE "
'''    strSQL = strSQL & "FORMAT(CNTA_CODE,'000000') BETWEEN '" & Me.LF0003_YAR1C & "' AND '" & Me.LF0003_YAR2C & "' AND "
'''    strSQL = strSQL & "CNTA_SYU > 0 "
'''    strSQL = strSQL & "GROUP BY "
'''    strSQL = strSQL & "CNTA_CODE "
'''    strSQL = strSQL & "; "
'''    Set RST_TRAN = dbMDB_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)
'''
'''    Set RST_WORK = dbMDB_DB.OpenRecordset("WK_LQ0003_CNTA", dbOpenDynaset, dbAppendOnly)
'''
'''    Do Until RST_TRAN.EOF
'''       With RST_WORK
'''       .AddNew
'''       '.Fields("LQ0003_YARDC") = RST_TRAN.Fields(0)                            'DELETE 20041111 N.MIURA
'''       .Fields("LQ0003_YARDC") = Format(RST_TRAN.Fields(0), "000000")           'INSERT 20041111 N.MIURA
'''       .Fields("LQ0003_SOHOQ") = RST_TRAN.Fields(1)
'''       .UPDATE
'''       End With
'''       RST_TRAN.MoveNext
'''    Loop
'''
'''    RST_TRAN.Close
'''    RST_WORK.Close
'''
'''    strMsg = strMsg & "地代計算＝＞"
'''    doCmd.Echo True, PROG_ID & ":" & strMsg
'''
'''    '地代
'''    strSQL = ""
'''    strSQL = strSQL & "SELECT "
'''    strSQL = strSQL & "JARG_YCODE,"
'''    strSQL = strSQL & "FORMAT(First(JARG_BKNO),'000000'), "
'''    strSQL = strSQL & "Sum(JARG_RENTKG) "
'''    strSQL = strSQL & "FROM "
'''    strSQL = strSQL & "dbo_JARG_FILE "
'''    strSQL = strSQL & "WHERE "
'''    strSQL = strSQL & " FORMAT(JARG_YCODE,'000000') BETWEEN '" & Me.LF0003_YAR1C & "' AND '" & Me.LF0003_YAR2C & "' AND "
'''    strSQL = strSQL & " Format(JARG_STDATE,'yyyymm') <= '" & Me.LF0003_YYMMD & "' AND "
'''    strSQL = strSQL & "(Format(JARG_KYDATE,'yyyymm') >= '" & Me.LF0003_YYMMD & "' Or Format(JARG_KYDATE,'yyyymm') = '') "
'''    strSQL = strSQL & "GROUP BY "
'''    strSQL = strSQL & "JARG_YCODE "
'''    strSQL = strSQL & "; "
'''    Set RST_TRAN = dbMDB_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)
'''
'''    Set RST_WORK = dbMDB_DB.OpenRecordset("WK_LQ0003_JARG", dbOpenDynaset, dbAppendOnly)
'''
'''    Do Until RST_TRAN.EOF
'''       With RST_WORK
'''       .AddNew
'''       '.Fields("LQ0003_YARDC") = RST_TRAN.Fields(0)                            'DELETE 20041111 N.MIURA
'''       .Fields("LQ0003_YARDC") = Format(RST_TRAN.Fields(0), "000000")           'INSERT 20041111 N.MIURA
'''       .Fields("LQ0003_BUKHC") = RST_TRAN.Fields(1)
'''       .Fields("LQ0003_SIIGA") = RST_TRAN.Fields(2)
'''       .UPDATE
'''       End With
'''       RST_TRAN.MoveNext
'''    Loop
'''
'''    RST_TRAN.Close
'''    RST_WORK.Close
'''
'''    strMsg = strMsg & "編集処理）"
'''    doCmd.Echo True, PROG_ID & ":" & strMsg
'''
'''    strSQL = ""
'''    strSQL = strSQL & "SELECT "
'''    strSQL = strSQL & "WK_LQ0003_CNTA.LQ0003_YARDC, "
'''    strSQL = strSQL & "WK_LQ0003_JARG.LQ0003_BUKHC, "
'''    strSQL = strSQL & "WK_LQ0003_JARG.LQ0003_SIIGA, "
'''    strSQL = strSQL & "WK_LQ0003_CARG.LQ0003_URIAA, "
'''    strSQL = strSQL & "WK_LQ0003_CARG.LQ0003_URIAQ, "
'''    strSQL = strSQL & "WK_LQ0003_CNTA.LQ0003_SOHOQ  "
'''    strSQL = strSQL & "FROM ("
'''    strSQL = strSQL & "WK_LQ0003_CNTA "
'''    strSQL = strSQL & "LEFT JOIN WK_LQ0003_CARG ON WK_LQ0003_CNTA.LQ0003_YARDC = WK_LQ0003_CARG.LQ0003_YARDC ) "
'''    strSQL = strSQL & "LEFT JOIN WK_LQ0003_JARG ON WK_LQ0003_CNTA.LQ0003_YARDC = WK_LQ0003_JARG.LQ0003_YARDC   "
'''    strSQL = strSQL & "; "
'''    Set RST_TRAN = dbMDB_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)
'''
'''    Set RST_WORK = dbMDB_DB.OpenRecordset("WK_LQ0003_TEMP", dbOpenDynaset, dbAppendOnly)
'''
'''    Do Until RST_TRAN.EOF
'''       With RST_WORK
'''       .AddNew
'''       .Fields("LQ0003_YARDC") = RST_TRAN.Fields("LQ0003_YARDC")
'''       '.Fields("LQ0003_YARDN") = RST_TRAN.Fields(0)
'''       .Fields("LQ0003_BUKHC") = RST_TRAN.Fields("LQ0003_BUKHC")
'''       .Fields("LQ0003_SIIGA") = Nz(RST_TRAN.Fields("LQ0003_SIIGA"), 0)
'''       .Fields("LQ0003_URIAA") = Nz(RST_TRAN.Fields("LQ0003_URIAA"), 0)
'''
'''       If Nz(RST_TRAN.Fields("LQ0003_SIIGA"), 0) = 0 Then
'''          .Fields("LQ0003_KEKKA") = 0
'''       Else
'''          .Fields("LQ0003_KEKKA") = Int((Nz(RST_TRAN.Fields("LQ0003_URIAA"), 0) / Nz(RST_TRAN.Fields("LQ0003_SIIGA"), 0)) * 100)
'''       End If
'''
'''       .Fields("LQ0003_SOHOQ") = Nz(RST_TRAN.Fields("LQ0003_SOHOQ"), 0)
'''       .Fields("LQ0003_ASSIQ") = Nz(RST_TRAN.Fields("LQ0003_URIAQ"), 0)
'''       .Fields("LQ0003_DASSQ") = Nz(RST_TRAN.Fields("LQ0003_SOHOQ"), 0) - Nz(RST_TRAN.Fields("LQ0003_URIAQ"), 0)
'''
'''       If Nz(RST_TRAN.Fields("LQ0003_SOHOQ"), 0) = 0 Then
'''          .Fields("LQ0003_SIYOR") = 0
'''       Else
'''          .Fields("LQ0003_SIYOR") = Int((Nz(RST_TRAN.Fields("LQ0003_URIAQ"), 0) / Nz(RST_TRAN.Fields("LQ0003_SOHOQ"), 0)) * 100)
'''       End If
'''
'''       .UPDATE
'''       End With
'''       RST_TRAN.MoveNext
'''    Loop
'''
'''    RST_TRAN.Close
'''    RST_WORK.Close
'''
'''    '担当者名称変換
'''    If LF0003_YARDN() = False Then
'''        Exit Function
'''    End If
'''
'''    LF0003_M20 = 1
'''
'''Exit_LF0003_M20:
'''    Exit Function
'''
'''Err_LF0003_M20:
'''    MsgBox Err.Description
'''    Resume Exit_LF0003_M20
'''End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :LF0003_M20
'        CREATE_DATE      :2004/09/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0003_M20() As Long
    On Error GoTo Err_LF0003_M20
    
    Dim strMsg           As String
    Dim RST_TRAN         As Recordset
    Dim RST_WORK         As Recordset
    
    Dim strBumonCd       As String                   ' 部門CD
    Dim dbSQLServer      As Database                 ' ﾃﾞｰﾀﾍﾞｰｽｵﾌﾞｼﾞｪｸﾄ(KONT_DB)
    Dim lngSYard         As Long                     ' 指定ﾔｰﾄﾞｺｰﾄﾞ(開始)
    Dim lngEyard         As Long                     ' 指定ﾔｰﾄﾞｺｰﾄﾞ(終了)
    Dim strYMD           As String
    Dim strOpenRowSetSql As String
    
    '▼ADD 2007/09/03 tajima
    Dim strDateFrom      As String                   '条件範囲From
    Dim strDateTo        As String                   '条件範囲To
    strDateFrom = Me.LF0003_YYMMD & "/01"
    strDateTo = DateSerial(Year(strDateFrom), Month(strDateFrom) + 1, 0)
    '▲ADD
    
    '部門コードのセット
    strBumonCd = fncGetBumonCd
    'DB接続
    Call subConnectServer(dbSQLServer, strBumonCd)
    
    '指定条件セット
    lngSYard = CLng(Me.LF0003_YAR1C)
    lngEyard = CLng(Me.LF0003_YAR2C)
    strYMD = Format$(DATE, "YYYYMMDD")
    
    doCmd.Echo True, PROG_ID & ":" & strMsg
    
    strSQL = ""
    strSQL = strSQL & "SELECT " & Chr(13)
    strSQL = strSQL & "       CARG_YCODE, " & Chr(13)
    strSQL = strSQL & "       Count(CARG_YCODE) CARG_COUNT,  " & Chr(13)
    strSQL = strSQL & "       Sum(CARG_RENTKG)  CARG_SUM " & Chr(13)
    strSQL = strSQL & "FROM " & Chr(13)
    strSQL = strSQL & "       CARG_FILE " & Chr(13)
    strSQL = strSQL & "       INNER JOIN YARD_MAST ON YARD_CODE = CARG_YCODE " & Chr(13)
    strSQL = strSQL & "WHERE " & Chr(13)
    strSQL = strSQL & "       CARG_YCODE BETWEEN " & lngSYard & " AND " & lngEyard & Chr(13)
    '▼CHG 2007/09/03 tajima 契約状態ではなく指定範囲内に借りていたかどうか
    'strSQL = strSQL & "       CARG_AGRE <> 9 " & Chr(13)
    strSQL = strSQL & "   AND CARG_STDATE <= '" & strDateTo & "' "
    strSQL = strSQL & "   AND ISNULL(CARG_KYDATE,'9999/12/31') >= '" & strDateFrom & "' "
    '▲CHG
    
    strSQL = strSQL & "       AND ISNULL(YARD_END_DAY,'9999/12/31') > GETDATE() " & Chr(13)
    strSQL = strSQL & "GROUP BY " & Chr(13)
    strSQL = strSQL & "       CARG_YCODE "
    Set RST_TRAN = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    Set RST_WORK = dbMDB_DB.OpenRecordset("WK_LQ0003_CARG", dbOpenDynaset, dbAppendOnly)
    
    Do Until RST_TRAN.EOF
       With RST_WORK
       .AddNew
       .Fields("LQ0003_YARDC") = Format(RST_TRAN.Fields("CARG_YCODE"), "000000")
       .Fields("LQ0003_URIAQ") = RST_TRAN.Fields("CARG_COUNT")
       .Fields("LQ0003_URIAA") = RST_TRAN.Fields("CARG_SUM")
       .UPDATE
       End With
       RST_TRAN.MoveNext
    Loop
    
    RST_TRAN.Close
    RST_WORK.Close
    
    '2013/1120 M.HONDA INS
    strSQL = ""
    strSQL = strSQL & "select YOUKT_YCODE, "
    strSQL = strSQL & "       SUM(CASE YOUKT_YUKBN WHEN  10 then "
    strSQL = strSQL & "1 "
    strSQL = strSQL & "       Else "
    strSQL = strSQL & "0 "
    strSQL = strSQL & "       end ) uke "
    strSQL = strSQL & "from YOUK_TRAN "
    strSQL = strSQL & "     INNER JOIN YARD_MAST ON "
    strSQL = strSQL & "           YARD_CODE = YOUKT_YCODE AND "
    strSQL = strSQL & "           ISNULL(YARD_END_DAY,'9999/12/31') > GETDATE() "
    strSQL = strSQL & "where YOUKT_UKDATE <= '" & strDateTo & "'  "
    strSQL = strSQL & "group by YOUKT_YCODE "
    Set RST_TRAN = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    Set RST_WORK = dbMDB_DB.OpenRecordset("WK_LQ0003_YOUK", dbOpenDynaset, dbAppendOnly)
    
    Do Until RST_TRAN.EOF
       With RST_WORK
       .AddNew
       .Fields("LQ0003_YARDC") = Format(RST_TRAN.Fields("YOUKT_YCODE"), "000000")
       .Fields("LQ0003_YOUKC") = RST_TRAN.Fields("uke")
       .UPDATE
       End With
       RST_TRAN.MoveNext
    Loop
    
    RST_TRAN.Close
    RST_WORK.Close
    '2013/1120 M.HONDA INS
    
    strMsg = strMsg & "総本数＝＞"
    doCmd.Echo True, PROG_ID & ":" & strMsg
    
    strSQL = ""
    strSQL = strSQL & "SELECT " & Chr(13)
    strSQL = strSQL & "       CNTA_CODE, " & Chr(13)
    strSQL = strSQL & "       Count(CNTA_NO) CNTA_COUNT, " & Chr(13)
    strSQL = strSQL & "       Min(YARD_NAME) YARD_NAME " & Chr(13)
    strSQL = strSQL & "FROM " & Chr(13)
    strSQL = strSQL & "       CNTA_MAST " & Chr(13)
    strSQL = strSQL & "INNER  JOIN YARD_MAST " & Chr(13)
    strSQL = strSQL & "       ON CNTA_CODE = YARD_CODE  " & Chr(13)
    strSQL = strSQL & "WHERE  CNTA_CODE BETWEEN " & lngSYard & " AND " & lngEyard & " AND " & Chr(13)
    strSQL = strSQL & "       CNTA_USE <> 9 " & Chr(13)
    strSQL = strSQL & "       AND ISNULL(YARD_END_DAY,'9999/12/31') > GETDATE() " & Chr(13)
    strSQL = strSQL & "GROUP BY " & Chr(13)
    strSQL = strSQL & "       CNTA_CODE, YARD_NAME "
    Set RST_TRAN = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    Set RST_WORK = dbMDB_DB.OpenRecordset("WK_LQ0003_CNTA", dbOpenDynaset, dbAppendOnly)

    Do Until RST_TRAN.EOF
       With RST_WORK
       .AddNew
       .Fields("LQ0003_YARDC") = Format(RST_TRAN.Fields("CNTA_CODE"), "000000")
       .Fields("LQ0003_SOHOQ") = RST_TRAN.Fields("CNTA_COUNT")
       .Fields("LQ0003_YARDN") = RST_TRAN.Fields("YARD_NAME")
       .UPDATE
       End With
       RST_TRAN.MoveNext
    Loop
    
    RST_TRAN.Close
    RST_WORK.Close
    
    strMsg = strMsg & "地代計算＝＞"
    doCmd.Echo True, PROG_ID & ":" & strMsg
    
    '地代
    strSQL = ""
    strSQL = strSQL & "SELECT " & Chr(13)
    strSQL = strSQL & "       JARG_YCODE," & Chr(13)
    strSQL = strSQL & "       MIN(JARG_BKNO) JARG_BKNO, " & Chr(13)
    strSQL = strSQL & "       Sum(JARG_RENTKG) JARG_SUM " & Chr(13)
    strSQL = strSQL & "FROM " & Chr(13)
    strSQL = strSQL & "       JARG_FILE " & Chr(13)
    strSQL = strSQL & "       INNER JOIN YARD_MAST ON YARD_CODE = JARG_YCODE " & Chr(13)
    strSQL = strSQL & "WHERE " & Chr(13)
    strSQL = strSQL & "       JARG_YCODE BETWEEN " & lngSYard & " AND " & lngEyard & " AND " & Chr(13)
    strSQL = strSQL & "       JARG_STDATE <= CONVERT(DATETIME,'" & strYMD & "',112) AND " & Chr(13)
    strSQL = strSQL & "      (JARG_KYDATE >= CONVERT(DATETIME,'" & strYMD & "',112) Or ISNULL(JARG_KYDATE,' ') = ' ') " & Chr(13)
    strSQL = strSQL & "       AND ISNULL(YARD_END_DAY,'9999/12/31') > GETDATE() " & Chr(13)
    strSQL = strSQL & "GROUP BY " & Chr(13)
    strSQL = strSQL & "       JARG_YCODE "
    Set RST_TRAN = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    Set RST_WORK = dbMDB_DB.OpenRecordset("WK_LQ0003_JARG", dbOpenDynaset, dbAppendOnly)
    
    Do Until RST_TRAN.EOF
       With RST_WORK
       .AddNew
       .Fields("LQ0003_YARDC") = Format(RST_TRAN.Fields("JARG_YCODE"), "000000")
       .Fields("LQ0003_BUKHC") = Format(RST_TRAN.Fields("JARG_BKNO"), "000000")
       .Fields("LQ0003_SIIGA") = RST_TRAN.Fields("JARG_SUM")
       .UPDATE
       End With
       RST_TRAN.MoveNext
    Loop
    
    RST_TRAN.Close
    RST_WORK.Close

    strMsg = strMsg & "編集処理）"
    doCmd.Echo True, PROG_ID & ":" & strMsg
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "WK_LQ0003_CNTA.LQ0003_YARDC, "
    strSQL = strSQL & "WK_LQ0003_CNTA.LQ0003_YARDN, "
    strSQL = strSQL & "WK_LQ0003_JARG.LQ0003_BUKHC, "
    strSQL = strSQL & "WK_LQ0003_JARG.LQ0003_SIIGA, "
    strSQL = strSQL & "WK_LQ0003_CARG.LQ0003_URIAA, "
    strSQL = strSQL & "WK_LQ0003_CARG.LQ0003_URIAQ, "
    strSQL = strSQL & "WK_LQ0003_YOUK.LQ0003_YOUKC, "
    strSQL = strSQL & "WK_LQ0003_CNTA.LQ0003_SOHOQ  "
    
    strSQL = strSQL & "FROM  "
    
    strSQL = strSQL & "((WK_LQ0003_CNTA "
    strSQL = strSQL & "LEFT JOIN WK_LQ0003_CARG ON WK_LQ0003_CNTA.LQ0003_YARDC = WK_LQ0003_CARG.LQ0003_YARDC) "
    strSQL = strSQL & "LEFT JOIN WK_LQ0003_JARG ON WK_LQ0003_CNTA.LQ0003_YARDC = WK_LQ0003_JARG.LQ0003_YARDC) "
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & "WK_LQ0003_YOUK ON "
    strSQL = strSQL & "WK_LQ0003_CNTA.LQ0003_YARDC = WK_LQ0003_YOUK.LQ0003_YARDC "
    ''2013/11/20 M.HONDA DEL
    ''strSQL = strSQL & "WK_LQ0003_CNTA "
    ''strSQL = strSQL & "LEFT JOIN WK_LQ0003_CARG ON WK_LQ0003_CNTA.LQ0003_YARDC = WK_LQ0003_CARG.LQ0003_YARDC ) "
    ''strSQL = strSQL & "LEFT JOIN WK_LQ0003_JARG ON WK_LQ0003_CNTA.LQ0003_YARDC = WK_LQ0003_JARG.LQ0003_YARDC   "
    ''2013/11/20 M.HONDA DEL
    strSQL = strSQL & "; "
    Set RST_TRAN = dbMDB_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)
    
    Set RST_WORK = dbMDB_DB.OpenRecordset("WK_LQ0003_TEMP", dbOpenDynaset, dbAppendOnly)
    
    Do Until RST_TRAN.EOF
       With RST_WORK
       .AddNew
       .Fields("LQ0003_YARDC") = RST_TRAN.Fields("LQ0003_YARDC")
       .Fields("LQ0003_YARDN") = RST_TRAN.Fields("LQ0003_YARDN")
       .Fields("LQ0003_BUKHC") = RST_TRAN.Fields("LQ0003_BUKHC")
       .Fields("LQ0003_SIIGA") = Nz(RST_TRAN.Fields("LQ0003_SIIGA"), 0)
       .Fields("LQ0003_URIAA") = Nz(RST_TRAN.Fields("LQ0003_URIAA"), 0)
       .Fields("LQ0003_YOUKC") = Nz(RST_TRAN.Fields("LQ0003_YOUKC"), 0)     ''2013/11/20 M.HONDA INS
       
       If Nz(RST_TRAN.Fields("LQ0003_SIIGA"), 0) = 0 Then
          .Fields("LQ0003_KEKKA") = 0
       Else
       
          .Fields("LQ0003_KEKKA") = Int((Nz(RST_TRAN.Fields("LQ0003_URIAA"), 0) / Nz(RST_TRAN.Fields("LQ0003_SIIGA"), 0)) * 100)
       End If
       
       .Fields("LQ0003_SOHOQ") = Nz(RST_TRAN.Fields("LQ0003_SOHOQ"), 0)
       .Fields("LQ0003_ASSIQ") = Nz(RST_TRAN.Fields("LQ0003_URIAQ"), 0)
       .Fields("LQ0003_DASSQ") = Nz(RST_TRAN.Fields("LQ0003_SOHOQ"), 0) - Nz(RST_TRAN.Fields("LQ0003_URIAQ"), 0)
       
       If Nz(RST_TRAN.Fields("LQ0003_SOHOQ"), 0) = 0 Then
          .Fields("LQ0003_SIYOR") = 0
       Else
          .Fields("LQ0003_SIYOR") = Int((Nz(RST_TRAN.Fields("LQ0003_URIAQ"), 0) / Nz(RST_TRAN.Fields("LQ0003_SOHOQ"), 0)) * 100)
       End If
       
       .UPDATE
       End With
       RST_TRAN.MoveNext
    Loop
    
    RST_TRAN.Close
    RST_WORK.Close
        
    'DB切断
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
        
    LF0003_M20 = 1

Exit_LF0003_M20:
    Exit Function

Err_LF0003_M20:
    MsgBox Err.Description
    Resume Exit_LF0003_M20
End Function
' 2007/04/25 ADD H.UENO End

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
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :LF0003_DROP
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0003_DROP() As Boolean
    
    strSQL = "DROP TABLE WK_LQ0003_CNTA ;"
    LF0003_SQLEXEC (strSQL)
    
    '2013/11/20 M.HONDA INS
    strSQL = "DROP TABLE WK_LQ0003_YOUK ;"
    LF0003_SQLEXEC (strSQL)
    '2013/11/20 M.HONDA INS
    
    strSQL = "DROP TABLE WK_LQ0003_CARG ;"
    LF0003_SQLEXEC (strSQL)
    
    strSQL = "DROP TABLE WK_LQ0003_JARG ;"
    LF0003_SQLEXEC (strSQL)
    
    strSQL = "DROP TABLE WK_LQ0003_TEMP ;"
    LF0003_SQLEXEC (strSQL)

End Function
' 2007/04/25 DEL H.UENO Start
'==============================================================================*
'
'        MODULE_NAME      :ﾔｰﾄﾞ名称変換
'        MODULE_ID        :LF0003_YARDN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'''Private Function LF0003_YARDN() As Boolean
'''
'''    strSQL = ""
'''    strSQL = strSQL & "UPDATE "
'''    strSQL = strSQL & "WK_LQ0003_TEMP "
'''    strSQL = strSQL & ",dbo_YARD_MAST "
'''    strSQL = strSQL & "SET "
'''    strSQL = strSQL & "LQ0003_YARDN = YARD_NAME "
'''    strSQL = strSQL & "WHERE "
'''    strSQL = strSQL & "FORMAT(YARD_CODE,'000000') = LQ0003_YARDC "
'''
'''    LF0003_YARDN = LF0003_SQLEXEC(strSQL)
'''
'''End Function
' 2007/04/25 DEL H.UENO End

'==============================================================================*
'
'        MODULE_NAME      :SQLの実行
'        MODULE_ID        :LF0003_SQLEXEC
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0003_SQLEXEC(strSQL As String) As Boolean
    On Error GoTo Err_LF0003_SQLEXEC

    dbMDB_DB.Execute strSQL, dbFailOnError

    LF0003_SQLEXEC = True

Exit_LF0003_SQLEXEC:
    Exit Function

Err_LF0003_SQLEXEC:
    LF0003_SQLEXEC = False
    Resume Exit_LF0003_SQLEXEC
End Function
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＯＰＥＮ
'        MODULE_ID        :LF0003_DBOPEN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0003_DBOPEN()
    On Error GoTo Err_LF0003_DBOPEN
    
    Set dbMDB_DB = CurrentDb()                                                  'INSERT 20040922 N.MIURA
    'Set dbs = CurrentDb()                                                      'DELETE 20040922 N.MIURA
    
Exit_LF0003_DBOPEN:
    Exit Sub

Err_LF0003_DBOPEN:
    MsgBox Err.Description
    Resume Exit_LF0003_DBOPEN
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＣＬＯＳＥ
'        MODULE_ID        :LF0003_DBCLOS
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0003_DBCLOS()
    On Error GoTo Err_LF0003_DBCLOS
    
    If Not dbMDB_DB Is Nothing Then                                             'INSERT 20040922 N.MIURA START
       dbMDB_DB.Close
       Set dbMDB_DB = Nothing
    End If                                                                      'INSERT 20040922 N.MIURA ENDED
    
    'dbs.Close                                                                  'DELETE 20040922 N.MIURA
    'Set dbs = Nothing                                                          'DELETE 20040922 N.MIURA

Exit_LF0003_DBCLOS:
    Exit Sub

Err_LF0003_DBCLOS:
    MsgBox Err.Description
    Resume Exit_LF0003_DBCLOS
End Sub

' 2007/04/25 ADD H.UENO Start
'==============================================================================*
'
'        MODULE_NAME      :部門コード取得
'        MODULE_ID        :fncGetBumonCd
'        PARAMETER        :
'        RETURN           :部門コード
'        CREATE_DATE      :2007/04/25
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetBumonCd() As String

    Dim objDb       As Database
    Dim strSQL      As String
    Dim objRs       As Recordset
    Dim strBumonCd  As String
    Dim blnErrFlg   As Boolean

    On Error GoTo fncGetBumonCd_Err

    strBumonCd = ""

    Set objDb = CurrentDb

    ' 部門コード取得
    strSQL = "SELECT BUMOM_BUMOC "             ' 部門コード
    strSQL = strSQL & " FROM  dbo_CONT_MAST "  ' コントロールマスタ(コンテナDB)
    strSQL = strSQL & "      ,BUMO_MAST "      ' 部門マスタ(加瀬DB)
    strSQL = strSQL & " WHERE CONT_BUMOC = BUMOM_BUMOC "
    Set objRs = objDb.OpenRecordset(strSQL, dbReadOnly)

    With objRs
        If Not .EOF Then
            strBumonCd = Nz(.Fields("BUMOM_BUMOC"), "")
        End If
    End With

    fncGetBumonCd = strBumonCd

    GoTo fncGetBumonCd_Exit

fncGetBumonCd_Err:
    blnErrFlg = True

fncGetBumonCd_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If

    ' カレントDBクローズ
    objDb.Close
    Set objDb = Nothing

    If blnErrFlg = True Then
        ' エラー発生時
        Call Err.Raise(Err.Number, "fncGetBumonCd" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : KONTデータベース接続
'       MODULE_ID       : subConnectServer
'       CREATE_DATE     : 2007/04/25
'       PARAM           : dbSqlServer    -データベースオブジェクト
'                       : strBumonCd     -部門コード
'       RETURN          :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subConnectServer(ByRef dbSQLServer As Database, ByVal strBumonCd As String)

On Error GoTo ErrorHandler

    Dim strSqlserver    As String
    Dim strConnect      As String
    Dim blnError        As Boolean
    
    blnError = False

    'SQL-Server名
    strSqlserver = fncGetSqlServer(strBumonCd)

    '接続文字列取得
    strConnect = fncGetConnectString(strBumonCd)

    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "subConnectServer" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : SQLサーバー名取得
'       MODULE_ID       : fncGetBumonArr
'       CREATE_DATE     : 2007/04/25
'       PARAM           : strBumonCode          部門コード(省略可)
'       RETURN          : SQLサーバー名
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetSqlServer(Optional strBumonCode As String = "") As String

On Error GoTo ErrorHandler

    Dim strSqlserver    As String
    Dim strParam        As String
    Dim blnError        As Boolean
    
    blnError = False

    strParam = "ODBC_DATA_SOURCE_NAME"
    If strBumonCode <> "" Then
        strParam = strParam & "_" & strBumonCode
    End If

    strSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = '" & strParam & "'"))
    If strSqlserver = "" Then
        'テーブル[SETU_TABL]の設定不正
        Call MSZZ024_M10("DlookUp", "SQL-Server名の設定不正。")
    End If

    fncGetSqlServer = strSqlserver

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "fncGetSqlServer" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : 接続文字列取得
'       MODULE_ID       : fncGetConnectString
'       CREATE_DATE     : 2007/04/25
'       PARAM           : strBumonCode          部門コード(省略可)
'       RETURN          : 接続文字列
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetConnectString(Optional strBumonCode As String = "") As String

On Error GoTo ErrorHandler

    Dim strConnectString    As String
    Dim blnError        As Boolean
    
    blnError = False

    strConnectString = MSZZ007_M00(strBumonCode)
    If strConnectString = "" Then
        'テーブル[SETU_TABL]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "接続文字列の設定不正。")
    End If

    fncGetConnectString = strConnectString

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "fncGetConnectString" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function


' 2007/04/25 ADD H.UENO End

'****************************  ended or program ********************************
'▼▼▼▼▼▼▼▼▼▼▼▼▼▼ 2007/09/03 add tajima ▼▼▼▼▼▼▼▼▼▼▼▼▼▼
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :LF0003_YYMMD_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0003_YYMMD_BeforeUpdate(Cancel As Integer)

    VarChangeDate = CmDateChecker(Me.LF0003_YYMMD.VALUE & "/01")
    If IsNull(VarChangeDate) Then
        MsgBox "日付を正しく入力してください。", vbInformation, 1 'P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :LF0003_YYMMD_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0003_YYMMD_AfterUpdate()

    Me.LF0003_YYMMD.VALUE = Format(Me.LF0003_YYMMD.VALUE & "/01", "yyyy/mm")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :LF0003_YYMMD_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0003_YYMMD_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(LF0003_YYMMD, KeyAscii, 7)
End Sub

