VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS051"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : コンテナ管理サブシステム
'
'        PROGRAM_NAME    :
'        PROGRAM_ID      : FKS051
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2003/10/17
'        CERATER         : N.MIURA
'
'        UPDATE          :
'        UPDATER         :
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
Private Const PROG_ID = "FKS051"
'
Private USER_ID         As String
Private MODL_ID         As String
'
Private dbs             As Database
Private strSQL          As String
'
Private RST_CARG        As Recordset
Private RST_REQU        As Recordset
Private RST_CONT        As Recordset
'
Private WK_KJDATE       As String
Private WK_NKDATE       As String
Private WK_SDATE        As String
'
Private WK_REQNO        As Double
'
Private INPCHEK_STS     As Boolean
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
    
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル13.BackColor = RGB(255, 0, 204)
    End Select
    
    Call MSHS010_M00(PROG_ID, "1")
    
    Call FKS051_DBOPEN
    
    USER_ID = LsGetUserName()

Exit_Form_Open:
    Exit Sub

Err_Form_Open:
    MsgBox Err.Description
    Resume Exit_Form_Open
End Sub
'==============================================================================*
'
'        MODULE_NAME      :顧客カナ
'        MODULE_ID        :FKS051_KOKYF_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_KOKYF_AfterUpdate()
    On Error GoTo Err_FKS051_KOKYF_AfterUpdate
       
    doCmd.Requery "FKS051_KOKYC"
    
    Call FKS051_CLE0R
    Call FKS051_CLE1R
    Call FKS051_CLE2R
    
Exit_FKS051_KOKYF_AfterUpdate:
    Exit Sub

Err_FKS051_KOKYF_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_FKS051_KOKYF_AfterUpdate


End Sub
'==============================================================================*
'
'        MODULE_NAME      :顧客ｺｰﾄﾞ
'        MODULE_ID        :FKS051_KOKYC_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_KOKYC_AfterUpdate()
    On Error GoTo Err_FKS051_KOKYC_AfterUpdate

    Me.FKS051_KOKYF = DLookup("USER_KANA", "dbo_USER_MAST", "FORMAT(USER_CODE,'000000') = FKS051_KOKYC ")
    Me.FKS051_KOKYN = DLookup("USER_NAME", "dbo_USER_MAST", "FORMAT(USER_CODE,'000000') = FKS051_KOKYC ")
    
    Call FKS051_CLE1R
    Call FKS051_CLE2R
    
Exit_FKS051_KOKYC_AfterUpdate:
    Exit Sub

Err_FKS051_KOKYC_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_FKS051_KOKYC_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :契約番号
'        MODULE_ID        :FKS051_KEIYB_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_KEIYB_AfterUpdate()
    On Error GoTo Err_FKS051_KEIYB_AfterUpdate
    
    Call FKS051_CLE2R
    
Exit_FKS051_KEIYB_AfterUpdate:
    Exit Sub

Err_FKS051_KEIYB_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_FKS051_KEIYB_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :売上年月
'        MODULE_ID        :FKS051_URYMD_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_URYMD_AfterUpdate()
    On Error GoTo Err_FKS051_URYMD_AfterUpdate
    
    If IsNull(Me.FKS051_URYMD) Or _
       Len(Me.FKS051_URYMD) <> 6 Or _
      (Not IsDate(Left(Me.FKS051_URYMD, 4) & "/" & Mid(Me.FKS051_URYMD, 5, 2) & "/01")) Then
        Me.FKS051_KEIYB.SetFocus
        MsgBox "日付の入力形式が正しくありません。"
    Else
        WK_KJDATE = Mid(Me.FKS051_URYMD, 1, 4) & "/" & Mid(Me.FKS051_URYMD, 5, 2) & "/01"
    
    Call FKS051_SEL_CARG
    Set RST_CARG = dbs.OpenRecordset(strSQL, dbOpenForwardOnly)
    
    If RST_CARG.EOF Then
       MsgBox ("契約番号の入力が正しくありません。")
       Me.FKS051_KEIYB.SetFocus
    Else
       strSQL = ""
       strSQL = strSQL & "SELECT "
       strSQL = strSQL & "REQU_REQNO, "
       strSQL = strSQL & "REQU_YYYYMM, "
       strSQL = strSQL & "REQU_KJDATE, "
       strSQL = strSQL & "REQU_SDATE, "
       strSQL = strSQL & "REQU_NKDATE, "
       strSQL = strSQL & "REQU_KCODE, "
       strSQL = strSQL & "REQU_NKBN, "
       strSQL = strSQL & "REQU_YCODE, "
       strSQL = strSQL & "REQU_NO, "
       strSQL = strSQL & "REQU_TTANKA, "
       strSQL = strSQL & "REQU_TUBOSU, "
       strSQL = strSQL & "REQU_KINGAK, "
       strSQL = strSQL & "REQU_SYOZEI, "
       strSQL = strSQL & "REQU_SECUKG, "
       strSQL = strSQL & "REQU_TOTAL, "
       strSQL = strSQL & "REQU_TEKI, "
       strSQL = strSQL & "REQU_FLG, "
       strSQL = strSQL & "REQU_UPDATE, "
       strSQL = strSQL & "REQU_ACPTNO "
       strSQL = strSQL & "FROM "
       strSQL = strSQL & "dbo_REQU_FILE "
       strSQL = strSQL & "WHERE "
       strSQL = strSQL & "REQU_KCODE  = " & RST_CARG.Fields("CARG_UCODE").VALUE & " AND "
       strSQL = strSQL & "REQU_YCODE  = " & RST_CARG.Fields("CARG_YCODE").VALUE & " AND "
       strSQL = strSQL & "REQU_NO     = " & RST_CARG.Fields("CARG_NO").VALUE & "    AND "
       strSQL = strSQL & "REQU_KJDATE =     CDate(#" & WK_KJDATE & "#) " & "        AND "
       strSQL = strSQL & "REQU_ACPTNO = " & "'" & Me.FKS051_KEIYB & "'"
       Set RST_REQU = dbs.OpenRecordset(strSQL, dbOpenForwardOnly)
       If RST_REQU.EOF Then
          '新規登録
          Me.FKS051_REQNQ = "000000"
          Me.FKS051_SYO1C = Format(RST_CARG.Fields("CARG_YCODE").VALUE, "000000")
          Me.FKS051_SYO2C = Format(RST_CARG.Fields("CARG_NO").VALUE, "000000")
          Me.FKS051_NKDATE = Format(DateAdd("D", -1, WK_KJDATE), "YYYYMMDD")
          Me.FKS051_SDATE = Format(Now(), "YYYYMMDD")
          Me.FKS051_URIAQ = RST_CARG.Fields("CARG_TUBOSU").VALUE
          Me.FKS051_URIAU = RST_CARG.Fields("CARG_TTANKA").VALUE
          
          If Format(RST_CARG.Fields("CARG_STDATE").VALUE, "YYYYMM") = Me.FKS051_URYMD Then
             '初回請求
             Me.FKS051_URIAA = RST_CARG.Fields("CARG_FRSTKG").VALUE
             Me.FKS051_URIAZ = RST_CARG.Fields("CARG_FSYOZEI").VALUE
             Me.FKS051_HOSYA = RST_CARG.Fields("CARG_SECUKG").VALUE
             Me.FKS051_URIAG = RST_CARG.Fields("CARG_FRSTKG").VALUE + RST_CARG.Fields("CARG_FSYOZEI").VALUE
          Else
             '初回移行
             Me.FKS051_URIAA = RST_CARG.Fields("CARG_RENTKG").VALUE
             Me.FKS051_URIAZ = RST_CARG.Fields("CARG_SYOZEI").VALUE
             Me.FKS051_HOSYA = 0
             Me.FKS051_URIAG = RST_CARG.Fields("CARG_RENTKG").VALUE + RST_CARG.Fields("CARG_SYOZEI").VALUE
          End If
          Me.FKS051_TEKI = Null
          
          Me.FKS051_INSEB.Enabled = True
          Me.FKS051_UPDAB.Enabled = False
          Me.FKS051_DELEB.Enabled = False

       Else
          '修正
          Me.FKS051_REQNQ = Format(RST_REQU.Fields("REQU_REQNO").VALUE, "000000")
          Me.FKS051_SYO1C = Format(RST_REQU.Fields("REQU_YCODE").VALUE, "000000")
          Me.FKS051_SYO2C = Format(RST_REQU.Fields("REQU_NO").VALUE, "000000")
          Me.FKS051_NKDATE = Format(RST_REQU.Fields("REQU_NKDATE").VALUE, "YYYYMMDD")
          Me.FKS051_SDATE = Format(RST_REQU.Fields("REQU_SDATE").VALUE, "YYYYMMDD")
          Me.FKS051_URIAQ = RST_REQU.Fields("REQU_TUBOSU").VALUE
          Me.FKS051_URIAU = RST_REQU.Fields("REQU_TTANKA").VALUE
          Me.FKS051_URIAA = RST_REQU.Fields("REQU_KINGAK").VALUE
          Me.FKS051_URIAZ = RST_REQU.Fields("REQU_SYOZEI").VALUE
          Me.FKS051_HOSYA = RST_REQU.Fields("REQU_SECUKG").VALUE
          Me.FKS051_URIAG = RST_REQU.Fields("REQU_TOTAL").VALUE
          Me.FKS051_TEKI = RST_REQU.Fields("REQU_TEKI").VALUE
        
          Me.FKS051_INSEB.Enabled = False
          Me.FKS051_UPDAB.Enabled = True
          Me.FKS051_DELEB.Enabled = True
       End If
       RST_REQU.Close
    End If
    
    RST_CARG.Close
    Set RST_CARG = Nothing
    Set RST_REQU = Nothing
    
    Call FKS051_SYO1C_AfterUpdate
    Call FKS051_SYO2C_AfterUpdate
        
    End If

Exit_FKS051_URYMD_AfterUpdate:
    Exit Sub

Err_FKS051_URYMD_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_FKS051_URYMD_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :商品ｺｰﾄﾞ１
'        MODULE_ID        :FKS051_SYO1C_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_SYO1C_AfterUpdate()
    On Error GoTo Err_FKS051_SYO1C_AfterUpdate
    
    doCmd.Requery "FKS051_SYO2C"

Exit_FKS051_SYO1C_AfterUpdate:
    Exit Sub

Err_FKS051_SYO1C_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_FKS051_SYO1C_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :商品ｺｰﾄﾞ２
'        MODULE_ID        :FKS051_SYO2C_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_SYO2C_AfterUpdate()
    On Error GoTo Err_FKS051_SYO2C_AfterUpdate
    
    Me.FKS051_SYOHN = DLookup("YARD_NAME", "dbo_YARD_MAST", "FORMAT(YARD_CODE,'000000') = FKS051_SYO1C")

Exit_FKS051_SYO2C_AfterUpdate:
    Exit Sub

Err_FKS051_SYO2C_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_FKS051_SYO2C_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :登録
'        MODULE_ID        :FKS051_INSEB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_INSEB_Click()
    On Error GoTo Err_FKS051_INSEB_Click

    Call FKS051_INCHK(INPCHEK_STS)
    
    If INPCHEK_STS = True Then
       If (MsgBox("登録しますか？", vbYesNo) = vbYes) Then
    
          WK_NKDATE = Mid(Me.FKS051_NKDATE, 1, 4) & "/" & Mid(Me.FKS051_NKDATE, 5, 2) & "/" & Mid(Me.FKS051_NKDATE, 7, 2)
          WK_SDATE = Mid(Me.FKS051_SDATE, 1, 4) & "/" & Mid(Me.FKS051_SDATE, 5, 2) & "/" & Mid(Me.FKS051_SDATE, 7, 2)
           
          strSQL = ""
          strSQL = strSQL & "SELECT "
          strSQL = strSQL & "CONT_KEY, "
          strSQL = strSQL & "CONT_REQNO, "
          strSQL = strSQL & "CONT_UPDATE "
          strSQL = strSQL & "FROM "
          strSQL = strSQL & "dbo_CONT_MAST "
          strSQL = strSQL & "WHERE "
          strSQL = strSQL & "CONT_KEY = 1 "
          Set RST_CONT = dbs.OpenRecordset(strSQL, dbDenyRead)
          If RST_CONT.EOF Then
             MsgBox "コントロールマスタに登録がありません。", vbExclamation
          Else
             
             RST_CONT.Edit
             WK_REQNO = RST_CONT.Fields("CONT_REQNO").VALUE + 1
          
             strSQL = ""
             strSQL = strSQL & "SELECT "
             strSQL = strSQL & "REQU_REQNO, "
             strSQL = strSQL & "REQU_YYYYMM, "
             strSQL = strSQL & "REQU_KJDATE, "
             strSQL = strSQL & "REQU_SDATE, "
             strSQL = strSQL & "REQU_NKDATE, "
             strSQL = strSQL & "REQU_KCODE, "
             strSQL = strSQL & "REQU_NKBN, "
             strSQL = strSQL & "REQU_YCODE, "
             strSQL = strSQL & "REQU_NO, "
             strSQL = strSQL & "REQU_TTANKA, "
             strSQL = strSQL & "REQU_TUBOSU, "
             strSQL = strSQL & "REQU_KINGAK, "
             strSQL = strSQL & "REQU_SYOZEI, "
             strSQL = strSQL & "REQU_SECUKG, "
             strSQL = strSQL & "REQU_TOTAL, "
             strSQL = strSQL & "REQU_TEKI, "
             strSQL = strSQL & "REQU_FLG, "
             strSQL = strSQL & "REQU_UPDATE, "
             strSQL = strSQL & "REQU_ACPTNO "
             strSQL = strSQL & "FROM "
             strSQL = strSQL & "dbo_REQU_FILE "
             Set RST_REQU = dbs.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)
    
             With RST_REQU
              .AddNew
              .Fields("REQU_REQNO").VALUE = WK_REQNO
              .Fields("REQU_YYYYMM").VALUE = DATE
              .Fields("REQU_KJDATE").VALUE = CDate(WK_KJDATE)
              .Fields("REQU_SDATE").VALUE = CDate(WK_SDATE)
              .Fields("REQU_NKDATE").VALUE = CDate(WK_NKDATE)
              .Fields("REQU_KCODE").VALUE = Me.FKS051_KOKYC
              .Fields("REQU_NKBN").VALUE = 0
              .Fields("REQU_YCODE").VALUE = Me.FKS051_SYO1C
              .Fields("REQU_NO").VALUE = Me.FKS051_SYO2C
              .Fields("REQU_TTANKA").VALUE = Me.FKS051_URIAU
              .Fields("REQU_TUBOSU").VALUE = Me.FKS051_URIAQ
              .Fields("REQU_KINGAK").VALUE = Me.FKS051_URIAA
              .Fields("REQU_SYOZEI").VALUE = Me.FKS051_URIAZ
              .Fields("REQU_SECUKG").VALUE = Me.FKS051_HOSYA
              .Fields("REQU_TOTAL").VALUE = Me.FKS051_URIAG
              .Fields("REQU_TEKI").VALUE = Me.FKS051_TEKI
              .Fields("REQU_FLG").VALUE = "9"
              .Fields("REQU_UPDATE").VALUE = DATE
              .Fields("REQU_ACPTNO").VALUE = Me.FKS051_KEIYB
              .UPDATE
             End With

             RST_CONT.Fields("CONT_REQNO").VALUE = WK_REQNO
             RST_CONT.Fields("CONT_UPDATE").VALUE = Now()
             RST_CONT.UPDATE
          
             RST_REQU.Close
             RST_CONT.Close

             Set RST_REQU = Nothing
             Set RST_CONT = Nothing
          
             Call FKS051_CLE2R
          
             'Me.FKS051_INSEB.Enabled = False
             ''Me.FKS051_UPDAB.Enabled = False
             ''Me.FKS051_DELEB.Enabled = False
          
             MsgBox "登録完了", vbExclamation
          End If
       End If
       
       Call FKS051_CLE3R
       
    Else
    End If

Exit_FKS051_INSEB_Click:
    Exit Sub

Err_FKS051_INSEB_Click:
    MsgBox Err.Description
    Resume Exit_FKS051_INSEB_Click
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :修正
'        MODULE_ID        :FKS051_UPDAB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_UPDAB_Click()
    On Error GoTo Err_FKS051_UPDAB_Click

    Call FKS051_INCHK(INPCHEK_STS)
    
    If INPCHEK_STS = True Then
       If (MsgBox("修正しますか？", vbYesNo) = vbYes) Then
    
          WK_NKDATE = Mid(Me.FKS051_NKDATE, 1, 4) & "/" & Mid(Me.FKS051_NKDATE, 5, 2) & "/" & Mid(Me.FKS051_NKDATE, 7, 2)
          
          strSQL = ""
          strSQL = strSQL & "SELECT "
          'strSql = strSql & "REQU_REQNO, "
          'strSql = strSql & "REQU_YYYYMM, "
          'strSql = strSql & "REQU_KJDATE, "
          'strSql = strSql & "REQU_SDATE, "
          strSQL = strSQL & "REQU_NKDATE, "
          'strSql = strSql & "REQU_KCODE, "
          'strSql = strSql & "REQU_NKBN, "
          strSQL = strSQL & "REQU_YCODE, "
          strSQL = strSQL & "REQU_NO, "
          strSQL = strSQL & "REQU_TTANKA, "
          strSQL = strSQL & "REQU_TUBOSU, "
          strSQL = strSQL & "REQU_KINGAK, "
          strSQL = strSQL & "REQU_SYOZEI, "
          strSQL = strSQL & "REQU_SECUKG, "
          strSQL = strSQL & "REQU_TOTAL, "
          strSQL = strSQL & "REQU_TEKI, "
          'strSql = strSql & "REQU_FLG, "
          strSQL = strSQL & "REQU_UPDATE "
          'strSql = strSql & "REQU_ACPTNO "
          strSQL = strSQL & "FROM "
          strSQL = strSQL & "dbo_REQU_FILE "
          strSQL = strSQL & "WHERE "
          strSQL = strSQL & "REQU_KCODE  = " & Me.FKS051_KOKYC & "          AND "
          strSQL = strSQL & "REQU_YCODE  = " & Me.FKS051_SYO1C & "          AND "
          strSQL = strSQL & "REQU_NO     = " & Me.FKS051_SYO2C & "          AND "
          strSQL = strSQL & "REQU_KJDATE = CDate(#" & WK_KJDATE & "#) " & " AND "
          strSQL = strSQL & "REQU_ACPTNO = " & "'" & Me.FKS051_KEIYB & "'"
          Set RST_REQU = dbs.OpenRecordset(strSQL, dbDenyRead)
    
          With RST_REQU
              .Edit
              '.Fields("REQU_REQNO").Value = Me.FKS051_REQNQ
              '.Fields("REQU_YYYYMM").Value = DATE
              '.Fields("REQU_KJDATE").Value = CDate(WK_KJDATE)
              '.Fields("REQU_SDATE").Value = Me.FKS051_SDATE
              .Fields("REQU_NKDATE").VALUE = WK_NKDATE
              '.Fields("REQU_KCODE").Value = Me.FKS051_KOKYC
              '.Fields("REQU_NKBN").Value = 0
              .Fields("REQU_YCODE").VALUE = Me.FKS051_SYO1C
              .Fields("REQU_NO").VALUE = Me.FKS051_SYO2C
              .Fields("REQU_TTANKA").VALUE = Me.FKS051_URIAU
              .Fields("REQU_TUBOSU").VALUE = Me.FKS051_URIAQ
              .Fields("REQU_KINGAK").VALUE = Me.FKS051_URIAA
              .Fields("REQU_SYOZEI").VALUE = Me.FKS051_URIAZ
              .Fields("REQU_SECUKG").VALUE = Me.FKS051_HOSYA
              .Fields("REQU_TOTAL").VALUE = Me.FKS051_URIAG
              .Fields("REQU_TEKI").VALUE = Me.FKS051_TEKI
              '.Fields("REQU_FLG").Value = "9"
              .Fields("REQU_UPDATE").VALUE = DATE
              '.Fields("REQU_ACPTNO").Value = Me.FKS051_KEIYB
              .UPDATE
          End With

          RST_REQU.Close
          Set RST_REQU = Nothing
          
          Call FKS051_CLE2R
          
          ''Me.FKS051_INSEB.Enabled = False
          'Me.FKS051_UPDAB.Enabled = False
          'Me.FKS051_DELEB.Enabled = False
          
          MsgBox "修正完了", vbExclamation
       
       End If
       
       Call FKS051_CLE3R
       
    Else
    End If
    
Exit_FKS051_UPDAB_Click:
    Exit Sub

Err_FKS051_UPDAB_Click:
    MsgBox Err.Description
    Resume Exit_FKS051_UPDAB_Click
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :削除
'        MODULE_ID        :FKS051_DELEB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_DELEB_Click()
    On Error GoTo Err_FKS051_DELEB_Click

    Call FKS051_INCHK(INPCHEK_STS)
    
    If INPCHEK_STS = True Then
       If (MsgBox("削除しますか？", vbYesNo) = vbYes) Then
    
          strSQL = ""
          strSQL = strSQL & "SELECT "
          strSQL = strSQL & "REQU_REQNO, "
          strSQL = strSQL & "REQU_YYYYMM, "
          strSQL = strSQL & "REQU_KJDATE, "
          strSQL = strSQL & "REQU_SDATE, "
          strSQL = strSQL & "REQU_NKDATE, "
          strSQL = strSQL & "REQU_KCODE, "
          strSQL = strSQL & "REQU_NKBN, "
          strSQL = strSQL & "REQU_YCODE, "
          strSQL = strSQL & "REQU_NO, "
          strSQL = strSQL & "REQU_TTANKA, "
          strSQL = strSQL & "REQU_TUBOSU, "
          strSQL = strSQL & "REQU_KINGAK, "
          strSQL = strSQL & "REQU_SYOZEI, "
          strSQL = strSQL & "REQU_SECUKG, "
          strSQL = strSQL & "REQU_TOTAL, "
          strSQL = strSQL & "REQU_ACPTNO "
          strSQL = strSQL & "FROM "
          strSQL = strSQL & "dbo_REQU_FILE "
          strSQL = strSQL & "WHERE "
          strSQL = strSQL & "REQU_KCODE  = " & Me.FKS051_KOKYC & "          AND "
          strSQL = strSQL & "REQU_YCODE  = " & Me.FKS051_SYO1C & "          AND "
          strSQL = strSQL & "REQU_NO     = " & Me.FKS051_SYO2C & "          AND "
          strSQL = strSQL & "REQU_KJDATE = CDate(#" & WK_KJDATE & "#) " & " AND "
          strSQL = strSQL & "REQU_ACPTNO = " & "'" & Me.FKS051_KEIYB & "'"
          Set RST_REQU = dbs.OpenRecordset(strSQL, dbDenyRead)
    
          With RST_REQU
              .Edit
              .Delete
          End With

          RST_REQU.Close
          Set RST_REQU = Nothing
          
          Call FKS051_CLE2R
          
          MsgBox "削除完了", vbExclamation
       
       End If
       
       Call FKS051_CLE3R
       
       ''Me.FKS051_INSEB.Enabled = False
       'Me.FKS051_UPDAB.Enabled = False
       'Me.FKS051_DELEB.Enabled = False
       
    Else
    End If
    
Exit_FKS051_DELEB_Click:
    Exit Sub

Err_FKS051_DELEB_Click:
    MsgBox Err.Description
    Resume Exit_FKS051_DELEB_Click
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :FKS051_SEL_CARG
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_SEL_CARG()
    On Error GoTo Err_FKS051_SEL_CARG
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "CARG_YCODE, "
    strSQL = strSQL & "CARG_NO, "
    strSQL = strSQL & "CARG_UCODE, "
    strSQL = strSQL & "CARG_AGRE, "
    strSQL = strSQL & "CARG_FSDATE, "
    strSQL = strSQL & "CARG_STDATE, "
    strSQL = strSQL & "CARG_EDDATE, "
    strSQL = strSQL & "CARG_CYDATE, "
    strSQL = strSQL & "CARG_KYDATE, "
    strSQL = strSQL & "CARG_CONTNA, "
    strSQL = strSQL & "CARG_TUBOSU, "
    strSQL = strSQL & "CARG_TTANKA, "
    strSQL = strSQL & "CARG_RENTKG, "
    strSQL = strSQL & "CARG_SYOZEI, "
    strSQL = strSQL & "CARG_SECUKG, "
    strSQL = strSQL & "CARG_FRSTKG, "
    strSQL = strSQL & "CARG_FSYOZEI, "
    strSQL = strSQL & "CARG_BIKO, "
    strSQL = strSQL & "CARG_ACPTNO  "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "dbo_CARG_FILE "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CARG_UCODE  =       " & Me.FKS051_KOKYC & " AND "
    strSQL = strSQL & "CARG_ACPTNO = " & "'" & Me.FKS051_KEIYB & "'"
    
Exit_FKS051_SEL_CARG:
    Exit Sub

Err_FKS051_SEL_CARG:
    MsgBox Err.Description
    Resume Exit_FKS051_SEL_CARG

End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＯＰＥＮ
'        MODULE_ID        :FKS051_DBOPEN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_DBOPEN()
    On Error GoTo Err_FKS051_DBOPEN
    
    Set dbs = CurrentDb()
    
Exit_FKS051_DBOPEN:
    Exit Sub

Err_FKS051_DBOPEN:
    MsgBox Err.Description
    Resume Exit_FKS051_DBOPEN

End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＣＬＯＳＥ
'        MODULE_ID        :FKS051_DBCLOS
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_DBCLOS()
    On Error GoTo Err_FKS051_DBCLOS
    
    dbs.Close
    
    Set dbs = Nothing

Exit_FKS051_DBCLOS:
    Exit Sub

Err_FKS051_DBCLOS:
    MsgBox Err.Description
    Resume Exit_FKS051_DBCLOS

End Sub
'==============================================================================*
'
'        MODULE_NAME      :クリアー
'        MODULE_ID        :FKS051_CLE0R
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_CLE0R()
    On Error GoTo Err_FKS051_CLE0R
    
    'Me.FKS051_KOKYF = Null
    Me.FKS051_KOKYC = Null
    Me.FKS051_KOKYN = Null
       
Exit_FKS051_CLE0R:
    Exit Sub

Err_FKS051_CLE0R:
    MsgBox Err.Description
    Resume Exit_FKS051_CLE0R
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :クリアー１
'        MODULE_ID        :FKS051_CLE1R
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_CLE1R()
    On Error GoTo Err_FKS051_CLE1R
    
    doCmd.Requery "FKS051_FKS051S10"
    Me.FKS051_KEIYB = Null
    doCmd.Requery "FKS051_KEIYB"
       
Exit_FKS051_CLE1R:
    Exit Sub

Err_FKS051_CLE1R:
    MsgBox Err.Description
    Resume Exit_FKS051_CLE1R
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :クリアー２
'        MODULE_ID        :FKS051_CLE2R
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_CLE2R()
    On Error GoTo Err_FKS051_CLE2R
    
    Me.FKS051_URYMD = Null
    Me.FKS051_SYO1C = Null
    Me.FKS051_SYO2C = Null
    Me.FKS051_HOSYA = 0
    Me.FKS051_URIAQ = 0
    Me.FKS051_URIAU = 0
    Me.FKS051_URIAA = 0
    Me.FKS051_URIAZ = 0
    Me.FKS051_URIAG = 0
    Me.FKS051_NKDATE = Null
    Me.FKS051_SDATE = Null
    Me.FKS051_REQNQ = "000000"
    Me.FKS051_TEKI = Null
    
    Call FKS051_SYO1C_AfterUpdate
    Call FKS051_SYO2C_AfterUpdate
    
    doCmd.Requery "FKS051_FKS051S00"
        
Exit_FKS051_CLE2R:
    Exit Sub

Err_FKS051_CLE2R:
    MsgBox Err.Description
    Resume Exit_FKS051_CLE2R
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :クリアー３
'        MODULE_ID        :FKS051_CLE2R
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_CLE3R()
    On Error GoTo Err_FKS051_CLE3R
    
    Me.FKS051_URYMD.SetFocus
       
Exit_FKS051_CLE3R:
    Exit Sub

Err_FKS051_CLE3R:
    MsgBox Err.Description
    Resume Exit_FKS051_CLE3R
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :入力チェック
'        MODULE_ID        :FKS051_INCHK
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_INCHK(INPSTS As Boolean)
    On Error GoTo FKS051_INCHK_Err
    
    INPSTS = True
    
    If IsNull(Me.FKS051_REQNQ) Then
       Me.FKS051_REQNQ.SetFocus
       INPSTS = False
    End If
    If IsNull(Me.FKS051_URIAG) Then
       Me.FKS051_URIAG.SetFocus
       INPSTS = False
    End If
    If IsNull(Me.FKS051_HOSYA) Then
       Me.FKS051_HOSYA.SetFocus
       INPSTS = False
    End If
    If IsNull(Me.FKS051_URIAZ) Then
       Me.FKS051_URIAZ.SetFocus
       INPSTS = False
    End If
    If IsNull(Me.FKS051_URIAA) Then
       Me.FKS051_URIAA.SetFocus
       INPSTS = False
    End If
    If IsNull(Me.FKS051_URIAU) Then
       Me.FKS051_URIAU.SetFocus
       INPSTS = False
    End If
    If IsNull(Me.FKS051_URIAQ) Then
       Me.FKS051_URIAQ.SetFocus
       INPSTS = False
    End If
    If IsNull(Me.FKS051_NKDATE) Then
       Me.FKS051_NKDATE.SetFocus
       INPSTS = False
    End If
    If IsNull(Me.FKS051_SYO2C) Then
       Me.FKS051_SYO2C.SetFocus
       INPSTS = False
    End If
    If IsNull(Me.FKS051_SYO1C) Then
       Me.FKS051_SYO1C.SetFocus
       INPSTS = False
    End If
    If IsNull(Me.FKS051_SDATE) Then
       Me.FKS051_SDATE.SetFocus
       INPSTS = False
    End If
    If IsNull(Me.FKS051_URYMD) Then
       Me.FKS051_URYMD.SetFocus
       INPSTS = False
    End If
    If Not INPSTS Then
       MsgBox ("必須入力項目です。")
       Exit Sub
    End If
    
    'If LENB(Me.FKS051_TEKI) > 36 Then
    '   Me.FKS051_TEKI.SetFocus
    '   INPSTS = False
    'End If
    If Len(Me.FKS051_REQNQ) <> 6 Then
       Me.FKS051_REQNQ.SetFocus
       INPSTS = False
    End If
    'If IsNull(Me.FKS051_URIAG) Then
    '   Me.FKS051_URIAG.SetFocus
    '   INPSTS = False
    'End If
    'If IsNull(Me.FKS051_HOSYA) Then
    '   Me.FKS051_HOSYA.SetFocus
    '   INPSTS = False
    'End If
    'If IsNull(Me.FKS051_URIAZ) Then
    '   Me.FKS051_URIAZ.SetFocus
    '   INPSTS = False
    'End If
    'If IsNull(Me.FKS051_URIAA) Then
    '   Me.FKS051_URIAA.SetFocus
    '   INPSTS = False
    'End If
    'If IsNull(Me.FKS051_URIAU) Then
    '   Me.FKS051_URIAU.SetFocus
    '   INPSTS = False
    'End If
    'If IsNull(Me.FKS051_URIAQ) Then
    '   Me.FKS051_URIAQ.SetFocus
    '   INPSTS = False
    'End If
    If Len(Me.FKS051_NKDATE) <> 8 Then
       Me.FKS051_NKDATE.SetFocus
       INPSTS = False
    End If
    If Len(Me.FKS051_SYO2C) <> 6 Then
       Me.FKS051_SYO2C.SetFocus
       INPSTS = False
    End If
    If Len(Me.FKS051_SYO1C) <> 6 Then
       Me.FKS051_SYO1C.SetFocus
       INPSTS = False
    End If
    If Len(Me.FKS051_SDATE) <> 8 Then
       Me.FKS051_SDATE.SetFocus
       INPSTS = False
    End If
    If Len(Me.FKS051_URYMD) <> 6 Then
       Me.FKS051_URYMD.SetFocus
       INPSTS = False
    End If
    If Not INPSTS Then
       MsgBox ("入力桁数が不正です。")
       Exit Sub
    End If
    
    If (Not IsDate(Left(Me.FKS051_NKDATE, 4) & "/" & Mid(Me.FKS051_NKDATE, 5, 2) & "/" & Right(Me.FKS051_NKDATE, 2))) Then
        Me.FKS051_NKDATE.SetFocus
       INPSTS = False
    End If
    If (Not IsDate(Left(Me.FKS051_SDATE, 4) & "/" & Mid(Me.FKS051_SDATE, 5, 2) & "/" & Right(Me.FKS051_SDATE, 2))) Then
        Me.FKS051_SDATE.SetFocus
       INPSTS = False
    End If
    If (Not IsDate(Left(Me.FKS051_URYMD, 4) & "/" & Mid(Me.FKS051_URYMD, 5, 2) & "/01")) Then
        Me.FKS051_URYMD.SetFocus
       INPSTS = False
    End If
    If Not INPSTS Then
        MsgBox "日付の入力形式が正しくありません。"
        Exit Sub
    End If
    
FKS051_INCHK_Exit:
    Exit Sub

FKS051_INCHK_Err:
    MsgBox Error$
    Resume FKS051_INCHK_Exit
End Sub
'==============================================================================*
'
'        MODULE_NAME      :閉じる
'        MODULE_ID        :FKS051_EXITB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FKS051_EXITB_Click()
    On Error GoTo Err_FKS051_EXITB_Click
    'On Error Resume Next

    Call FKS051_DBCLOS
    
    Call MSHS010_M00(PROG_ID, "2")
    
    'doCmd.CLOSE
    doCmd.Close acForm, Me.FormName

Exit_FKS051_EXITB_Click:
    Exit Sub

Err_FKS051_EXITB_Click:
    MsgBox Err.Description
    Resume Exit_FKS051_EXITB_Click
    
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
