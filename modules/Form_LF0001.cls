VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LF0001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : コンテナ管理サブシステム
'        SUB_SYSTEM_NAME : 分析帳票
'
'        PROGRAM_NAME    : 契約・解約集計表
'        PROGRAM_ID      : LF0001
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2004/05/28
'        CERATER         : N.MIURA
'
'        UPDATE          : 2007/04/23
'        UPDATER         : H.UENO
'                        : 契約・解約本数に契約移動を含めない
'                          日付項目入力チェック追加
'
'        UPDATE          : 2010/08/24
'        UPDATER         : M.Ryu
'                        : 出力Excelに新規金額欄に”初回金額”を表示
'                          修正前：月額料金　⇒　修正後：初回金額　（三浦氏依頼）
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "LF0001"
'
Private USER_ID         As String
Private MODL_ID         As String
'
Private dbMDB           As Database
Private strSQL          As String
'
Private RST_YARD        As Recordset
'Private RST_NWCR        As Recordset
'Private RST_KACR        As Recordset
Private RST_Q001        As Recordset
'
Private INPCHEK_STS     As Boolean

' 2007/04/23 ADD H.UENO Start
Private Const P_DATE_LENGTH         As Integer = 10      '日付
Private pvarChangeDate              As Variant           '変更前処理で入力した日付を保存する
Private pstrBumonCd                 As String            '部門CD
' 2007/04/23 ADD H.UENO End

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
    
' 2007/04/23 ADD H.UENO Start
    Dim dtStart          As Date
    Dim dtEnd            As Date
' 2007/04/23 ADD H.UENO End
            
    Call MSHS010_M00(PROG_ID, "1")
    
    Call LF0001_DBOPEN
    
    USER_ID = LsGetUserName()
    
' 2007/04/23 ADD H.UENO Start
    '部門コードのセット
    pstrBumonCd = fncGetBumonCd
    '日付項目初期値セット
    dtEnd = DateAdd("d", -1, DateSerial(Year(Now), Month(Now), 1))
    dtStart = DateSerial(Year(dtEnd), Month(dtEnd), 1)
    Me.LF0001_DAT1D = Format$(dtStart, "YYYY/MM/DD")
    Me.LF0001_DAT2D = Format$(dtEnd, "YYYY/MM/DD")
    'ヤードコード初期値セット
    Me.LF0001_YAR1C = Format$(Me.LF0001_YAR1C.ItemData(0), "000000")
    Me.LF0001_YAR2C = Format$(Me.LF0001_YAR2C.ItemData(Me.LF0001_YAR2C.ListCount - 1), "000000")
' 2007/04/23 ADD H.UENO End

Exit_Form_Open:
    Exit Sub

Err_Form_Open:
    MsgBox Err.Description
    Resume Exit_Form_Open
End Sub

'==============================================================================*
'
'        MODULE_NAME      :実行
'        MODULE_ID        :LF0001_EXE1B_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_EXE1B_Click()
    On Error GoTo Err_LF0001_EXE1B_Click
    
' 2007/04/23 ADD H.UENO Start
    If fncInputCheck = False Then
        Exit Sub
    End If
' 2007/04/23 ADD H.UENO End
    
    If LF0001_CREATE = True Then
    
    'Select Case LF0001_LISTG
    '       Case 0
                doCmd.OpenReport "LR0001_M00", acPreview, "", ""
    '       Case 1
    '            doCmd.OpenReport "LR0001_M10", acPreview, "", ""
    'End Select
    
    End If

Exit_LF0001_EXE1B_Click:
    Exit Sub

Err_LF0001_EXE1B_Click:
    MsgBox Err.Description
    Resume Exit_LF0001_EXE1B_Click
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :EXCEL
'        MODULE_ID        :LF0001_EXE2B_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_EXE2B_Click()
    On Error GoTo Err_LF0001_EXE2B_Click
       
' 2007/04/23 ADD H.UENO Start
    If fncInputCheck = False Then
        Exit Sub
    End If
' 2007/04/23 ADD H.UENO End
        
    If LF0001_CREATE = True Then
       doCmd.OutputTo acQuery, "LQ0001_M00", "MicrosoftExcel(*.xls)", "", True, ""
       'doCmd.OutputTo acOutputTable, "WK_LQ0001_TEMP", "MicrosoftExcel(*.xls)", "", True, ""
    End If

Exit_LF0001_EXE2B_Click:
    Exit Sub

Err_LF0001_EXE2B_Click:
    MsgBox Err.Description
    Resume Exit_LF0001_EXE2B_Click
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :閉じる
'        MODULE_ID        :LF0001_EXITB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_EXITB_Click()
    On Error GoTo Err_LF0001_EXITB_Click

    Call LF0001_DBCLOS
    
    Call MSHS010_M00(PROG_ID, "2")
    
    doCmd.Close

Exit_LF0001_EXITB_Click:
    Exit Sub

Err_LF0001_EXITB_Click:
    MsgBox Err.Description
    Resume Exit_LF0001_EXITB_Click
    
End Sub
' 2007/04/23 ADD H.UENO Start
'==============================================================================*
'
'        MODULE_NAME      :LF0001_DAT1D_AfterUpdate
'        機能             :LF0001_DAT1D_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_DAT1D_AfterUpdate()

    On Error GoTo Exception

    LF0001_DAT1D = Format(pvarChangeDate, "yyyy/mm/dd")
    
    Exit Sub
    
Exception:
    Call MSZZ024_M00("LF0001_DAT1D_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :LF0001_DAT1D_BeforeUpdate
'        機能             :LF0001_DAT1D_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_DAT1D_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Exception

    Call fncDateCheck(LF0001_DAT1D, Cancel)

    Exit Sub
    
Exception:
    Call MSZZ024_M00("LF0001_DAT1D_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :LF0001_DAT1D_GotFocus
'        機能             :LF0001_DAT1D_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_DAT1D_GotFocus()

    On Error GoTo Exception

    LF0001_DAT1D = Format(LF0001_DAT1D, "yy/mm/dd")

    Exit Sub
    
Exception:
    Call MSZZ024_M00("LF0001_DAT1D_GotFocus", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :LF0001_DAT1D_KeyPress
'        機能             :LF0001_DAT1D_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_DAT1D_KeyPress(KeyAscii As Integer)

    On Error GoTo Exception
    
    Call CmKeyPress_For_Date(LF0001_DAT1D, KeyAscii, P_DATE_LENGTH)

    Exit Sub
    
Exception:
    Call MSZZ024_M00("LF0001_DAT1D_KeyPress", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :LF0001_DAT1D_LostFocus
'        機能             :LF0001_DAT1D_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_DAT1D_LostFocus()

    On Error GoTo Exception
    
    LF0001_DAT1D = Format(LF0001_DAT1D, "yyyy/mm/dd")

    Exit Sub
    
Exception:
    Call MSZZ024_M00("LF0001_DAT1D_LostFocus", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :LF0001_DAT2D_AfterUpdate
'        機能             :LF0001_DAT2D_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_DAT2D_AfterUpdate()

    On Error GoTo Exception
    
    LF0001_DAT2D = Format(pvarChangeDate, "yyyy/mm/dd")

    Exit Sub
    
Exception:
    Call MSZZ024_M00("LF0001_DAT2D_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :LF0001_DAT2D_BeforeUpdate
'        機能             :LF0001_DAT2D_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_DAT2D_BeforeUpdate(Cancel As Integer)

    On Error GoTo Exception
    
    Call fncDateCheck(LF0001_DAT2D, Cancel)

    Exit Sub
    
Exception:
    Call MSZZ024_M00("LF0001_DAT2D_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :LF0001_DAT2D_GotFocus
'        機能             :LF0001_DAT2D_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_DAT2D_GotFocus()

    On Error GoTo Exception
    
    LF0001_DAT2D = Format(LF0001_DAT2D, "yy/mm/dd")

    Exit Sub
    
Exception:
    Call MSZZ024_M00("LF0001_DAT2D_GotFocus", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :LF0001_DAT2D_KeyPress
'        機能             :LF0001_DAT2D_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_DAT2D_KeyPress(KeyAscii As Integer)

    On Error GoTo Exception
    
    Call CmKeyPress_For_Date(LF0001_DAT2D, KeyAscii, P_DATE_LENGTH)

    Exit Sub
    
Exception:
    Call MSZZ024_M00("LF0001_DAT2D_KeyPress", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :LF0001_DAT2D_LostFocus
'        機能             :LF0001_DAT2D_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_DAT2D_LostFocus()

    On Error GoTo Exception
    
    LF0001_DAT2D = Format(LF0001_DAT2D, "yyyy/mm/dd")

    Exit Sub
    
Exception:
    Call MSZZ024_M00("LF0001_DAT2D_LostFocus", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncDateCheck
'        機能             :日付入力チェック
'        IN               :対象テキスト
'                          更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDateCheck(txtTarget As Variant, ByRef Cancel As Integer) As Boolean

    On Error GoTo ErrorHandler

    Dim blnError As Boolean

    fncDateCheck = True
    blnError = False
    
    pvarChangeDate = CmDateChecker(txtTarget)
    If IsNull(pvarChangeDate) Then
        MsgBox "日付を正しく入力してください。", vbInformation, "エラー"
        Cancel = True
        SendKeys "{ESC}"
        fncDateCheck = False
    End If

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "fncDateCheck" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncInputCheck
'        機能             :未入力チェック
'        IN               :対象テキスト
'                          更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheck() As Boolean

    On Error GoTo ErrorHandler

    Dim blnError As Boolean

    fncInputCheck = False
    blnError = False
    
    If Nz(Me.LF0001_DAT1D.VALUE, "") = "" Then
        MsgBox "日付を入力して下さい。", vbInformation, "確認"
        Me.LF0001_DAT1D.SetFocus
        Exit Function
    End If
    If Nz(Me.LF0001_DAT2D.VALUE, "") = "" Then
        MsgBox "日付を入力して下さい。", vbInformation, "確認"
        Me.LF0001_DAT2D.SetFocus
        Exit Function
    End If
    If Nz(LF0001_YAR1C.VALUE, "") = "" Then
        MsgBox "ヤードコードを入力して下さい。", vbInformation, "確認"
        Me.LF0001_YAR1C.SetFocus
        Exit Function
    End If
    If Nz(LF0001_YAR2C.VALUE, "") = "" Then
        MsgBox "ヤードコードを入力して下さい。", vbInformation, "確認"
        Me.LF0001_YAR2C.SetFocus
        Exit Function
    End If

    fncInputCheck = True
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "fncDateCheck" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :部門コード取得
'        MODULE_ID        :fncGetBumonCd
'        PARAMETER        :
'        RETURN           :部門コード
'        CREATE_DATE      :2007/04/23
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
'       CREATE_DATE     :
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
'       CREATE_DATE     :
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
'       CREATE_DATE     :
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
' 2007/04/23 ADD H.UENO End

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :LF0001_CREATE
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0001_CREATE() As Boolean
    On Error GoTo Err_LF0001_CREATE
    
' 2007/04/23 ADD H.UENO Start
    Dim dbSQLServer As Database
    'DB接続
    Call subConnectServer(dbSQLServer, pstrBumonCd)
' 2007/04/23 ADD H.UENO End
    
    LF0001_CREATE = True
    
    strSQL = "DELETE FROM WK_LQ0001_TEMP ; "
    dbMDB.Execute strSQL, dbFailOnError

    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "YARD_CODE, "
    strSQL = strSQL & "YARD_NAME, "
    strSQL = strSQL & "YARD_YUBINO, "
    strSQL = strSQL & "YARD_ADDR_1, "
    strSQL = strSQL & "YARD_ADDR_2  "
    strSQL = strSQL & "FROM "
' 2007/04/23 UPDATE H.UENO Start
'''    strSQL = strSQL & "dbo_YARD_MAST "
    strSQL = strSQL & "YARD_MAST "
' 2007/04/23 UPDATE H.UENO End
    strSQL = strSQL & "WHERE "
' 2007/04/23 UPDATE H.UENO Start
'''    strSQL = strSQL & "YARD_CODE Between " & Me.LF0001_YAR1C & " AND " & Me.LF0001_YAR2C & " "
    If Me.LF0001_YAR1C >= Me.LF0001_YAR2C Then
        strSQL = strSQL & "YARD_CODE Between " & Me.LF0001_YAR2C & " AND " & Me.LF0001_YAR1C & " "
    Else
        strSQL = strSQL & "YARD_CODE Between " & Me.LF0001_YAR1C & " AND " & Me.LF0001_YAR2C & " "
    End If
' 2007/04/23 UPDATE H.UENO End
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "YARD_CODE "
    strSQL = strSQL & "; "
' 2007/04/23 UPDATE H.UENO Start
'''    Set RST_YARD = dbMDB.OpenRecordset(strSQL, dbReadOnly)
    Set RST_YARD = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
' 2007/04/23 UPDATE H.UENO End
   
    Set RST_Q001 = dbMDB.OpenRecordset("WK_LQ0001_TEMP", dbOpenDynaset, dbAppendOnly)
    
    Do Until RST_YARD.EOF
       
       With RST_Q001
       .AddNew
       .Fields("LQ0001_YARDC") = RST_YARD.Fields("YARD_CODE")
       .Fields("LQ0001_YARDN") = RST_YARD.Fields("YARD_NAME")
       .Fields("LQ0001_SYUBI") = Null
       .Fields("LQ0001_SYUBN") = Null
       .Fields("LQ0001_NEWKQ") = 0
       .Fields("LQ0001_NEWKA") = 0
       .Fields("LQ0001_KAIKQ") = 0
       .Fields("LQ0001_KAIKA") = 0
       .UPDATE
       End With
       
       RST_YARD.MoveNext
    Loop
    
    RST_Q001.Close
    RST_YARD.Close
        
' 2007/04/23 UPDATE H.UENO Start
'''    If LF0001_NEW_CNT = True Then
    If LF0001_NEW_CNT(dbSQLServer) = True Then
' 2007/04/23 UPDATE H.UENO End
          'GoTo Err_LF0001_CREATE
    End If
       
' 2007/04/23 UPDATE H.UENO Start
'''    If LF0001_KAI_CNT = True Then
    If LF0001_KAI_CNT(dbSQLServer) = True Then
' 2007/04/23 UPDATE H.UENO End
          'GoTo Err_LF0001_CREATE
    End If
    
    strSQL = "DELETE FROM WK_LQ0001_TEMP WHERE LQ0001_NEWKQ = 0 AND LQ0001_KAIKQ = 0; "
    dbMDB.Execute strSQL, dbFailOnError
    
' 2007/04/23 ADD H.UENO Start
    'DB切断
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
' 2007/04/23 ADD H.UENO End
    
Exit_LF0001_CREATE:
    Exit Function

Err_LF0001_CREATE:
    MsgBox Err.Description
    LF0001_CREATE = False
    Resume Exit_LF0001_CREATE
End Function

'==============================================================================*
'
'        MODULE_NAME      :新規契約本数取得
'        MODULE_ID        :LF0001_NEW_CNT
'        CREATE_DATE      :
'        PARAM            :dbSqlServer - KONTに接続したデータベースオブジェクト
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 2007/04/23 UPDATE H.UENO Start
'''Private Function LF0001_NEW_CNT() As Boolean
Private Function LF0001_NEW_CNT(dbSQLServer As Database) As Boolean
' 2007/04/23 UPDATE H.UENO End
    On Error GoTo Err_LF0001_NEW_CNT
    
    Dim strSQL              As String
    Dim rsKEIY_COUNT        As Recordset

    strSQL = ""
' 2007/04/23 UPDATE H.UENO Start
'''    strSQL = strSQL & "SELECT "
'''    strSQL = strSQL & "CARG_YCODE, "
'''    strSQL = strSQL & "COUNT(*)         AS COUNT, "
'''    strSQL = strSQL & "Sum(CARG_RENTKG) AS GOUKE  "
'''    strSQL = strSQL & "FROM "
'''    strSQL = strSQL & "dbo_CARG_FILE "
'''    strSQL = strSQL & "WHERE "
'''    strSQL = strSQL & "Format([CARG_STDATE],'yyyymmdd') Between '" & Me.LF0001_DAT1D & "' AND '" & Me.LF0001_DAT2D & "'"
    strSQL = "SELECT " & Chr(13)
    strSQL = strSQL & "CARG_YCODE, " & Chr(13)
    strSQL = strSQL & "COUNT(*)         CNT,  " & Chr(13)
    
    '---2010/08/24--ryu--update---<s>--'修正前：月額料金　⇒　修正後：初回金額　（三浦氏依頼）
    strSQL = strSQL & "SUM(CARG_FRSTKG) GOUKE " & Chr(13)
    'strSQL = strSQL & "SUM(CARG_RENTKG) GOUKE " & Chr(13)
    '---2010/08/24--ryu--update---<e>
    
    strSQL = strSQL & "FROM " & Chr(13)
    strSQL = strSQL & "CARG_FILE " & Chr(13)
    strSQL = strSQL & "WHERE " & Chr(13)
    If Me.LF0001_DAT1D >= Me.LF0001_DAT2D Then
        strSQL = strSQL & "CONVERT(varchar, CARG_STDATE, 111) Between '" & Me.LF0001_DAT2D & "' AND '" & Me.LF0001_DAT1D & "'" & Chr(13)
    Else
        strSQL = strSQL & "CONVERT(varchar, CARG_STDATE, 111) Between '" & Me.LF0001_DAT1D & "' AND '" & Me.LF0001_DAT2D & "'" & Chr(13)
    End If
    If Me.LF0001_YAR1C >= Me.LF0001_YAR2C Then
        strSQL = strSQL & "AND CARG_YCODE Between " & Me.LF0001_YAR2C & " AND " & Me.LF0001_YAR1C & " "
    Else
        strSQL = strSQL & "AND CARG_YCODE Between " & Me.LF0001_YAR1C & " AND " & Me.LF0001_YAR2C & " "
    End If
    strSQL = strSQL & "AND NOT EXISTS( " & Chr(13)
    strSQL = strSQL & "               SELECT * FROM YOUK_TRAN " & Chr(13)
    strSQL = strSQL & "               WHERE    YOUKT_UKNO = CARG_UKNO " & Chr(13)
    strSQL = strSQL & "               AND      ISNULL(YOUKT_MOTO_ACPTNO,'') <> '' ) " & Chr(13)
' 2007/04/23 UPDATE H.UENO End
    strSQL = strSQL & "GROUP BY CARG_YCODE ;"

' 2007/04/23 UPDATE H.UENO Start
'''    Set rsKEIY_COUNT = dbMDB.OpenRecordset(strSQL, dbReadOnly)
    Set rsKEIY_COUNT = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
' 2007/04/23 UPDATE H.UENO End

    Do Until rsKEIY_COUNT.EOF
    
       strSQL = "SELECT * FROM WK_LQ0001_TEMP WHERE LQ0001_YARDC = " & rsKEIY_COUNT.Fields("CARG_YCODE") & ";"
       Set RST_Q001 = dbMDB.OpenRecordset(strSQL, dbOpenDynaset)
    
' 2007/04/23 UPDATE H.UENO Start
        If RST_Q001.BOF = False Then
' 2007/04/23 UPDATE H.UENO End
           With RST_Q001
              .Edit
' 2007/04/23 UPDATE H.UENO Start
'''       .Fields("LQ0001_NEWKQ") = Nz(rsKEIY_COUNT.Fields("COUNT").Value, 0)
               .Fields("LQ0001_NEWKQ") = Nz(rsKEIY_COUNT.Fields("CNT").VALUE, 0)
' 2007/04/23 UPDATE H.UENO End
               .Fields("LQ0001_NEWKA") = Nz(rsKEIY_COUNT.Fields("GOUKE").VALUE, 0)
               .UPDATE
           End With
' 2007/04/23 UPDATE H.UENO Start
       End If
' 2007/04/23 UPDATE H.UENO End
       RST_Q001.Close
       rsKEIY_COUNT.MoveNext
    Loop
    
    rsKEIY_COUNT.Close

    LF0001_NEW_CNT = True
    Exit Function

Err_LF0001_NEW_CNT:
    MsgBox Err.Description
    LF0001_NEW_CNT = False
End Function
'==============================================================================*
'
'        MODULE_NAME      :解約本数取得
'        MODULE_ID        :LF0001_KAI_CNT
'        CREATE_DATE      :
'        PARAM            :dbSqlServer - KONTに接続したデータベースオブジェクト
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 2007/04/23 UPDATE H.UENO Start
'''Private Function LF0001_KAI_CNT() As Boolean
Private Function LF0001_KAI_CNT(dbSQLServer As Database) As Boolean
' 2007/04/23 UPDATE H.UENO End

    On Error GoTo Err_LF0001_KAI_CNT
    
    Dim strSQL              As String
    Dim rsKEIY_COUNT        As Recordset

    strSQL = ""
' 2007/04/23 UPDATE H.UENO Start
'''    strSQL = strSQL & "SELECT "
'''    strSQL = strSQL & "CARG_YCODE, "
'''    strSQL = strSQL & "COUNT(*)         AS COUNT, "
'''    strSQL = strSQL & "Sum(CARG_RENTKG) AS GOUKE  "
'''    strSQL = strSQL & "FROM "
'''    strSQL = strSQL & "dbo_CARG_FILE "
'''    strSQL = strSQL & "WHERE "
'''    strSQL = strSQL & "Format([CARG_KYDATE],'yyyymmdd') Between '" & Me.LF0001_DAT1D & "' AND '" & Me.LF0001_DAT2D & "'"
    strSQL = "SELECT " & Chr(13)
    strSQL = strSQL & "CARG_YCODE, " & Chr(13)
    strSQL = strSQL & "COUNT(*)         CNT,  " & Chr(13)
    strSQL = strSQL & "SUM(CARG_RENTKG) GOUKE " & Chr(13)
    strSQL = strSQL & "FROM " & Chr(13)
    strSQL = strSQL & "CARG_FILE " & Chr(13)
    strSQL = strSQL & "WHERE " & Chr(13)
    If Me.LF0001_DAT1D >= Me.LF0001_DAT2D Then
        strSQL = strSQL & "CONVERT(varchar, CARG_KYDATE, 111) Between '" & Me.LF0001_DAT2D & "' AND '" & Me.LF0001_DAT1D & "'" & Chr(13)
    Else
        strSQL = strSQL & "CONVERT(varchar, CARG_KYDATE, 111) Between '" & Me.LF0001_DAT1D & "' AND '" & Me.LF0001_DAT2D & "'" & Chr(13)
    End If
    If Me.LF0001_YAR1C >= Me.LF0001_YAR2C Then
        strSQL = strSQL & "AND CARG_YCODE Between " & Me.LF0001_YAR2C & " AND " & Me.LF0001_YAR1C & " "
    Else
        strSQL = strSQL & "AND CARG_YCODE Between " & Me.LF0001_YAR1C & " AND " & Me.LF0001_YAR2C & " "
    End If
    strSQL = strSQL & "AND NOT EXISTS( " & Chr(13)
    strSQL = strSQL & "               SELECT * FROM YOUK_TRAN " & Chr(13)
    strSQL = strSQL & "               WHERE    YOUKT_MOTO_ACPTNO = CARG_ACPTNO " & Chr(13)
    strSQL = strSQL & "               AND      ISNULL(YOUKT_MOVEKBN, 0) <> 0 ) " & Chr(13)
' 2007/04/23 UPDATE H.UENO End
    strSQL = strSQL & "GROUP BY CARG_YCODE ;"

' 2007/04/23 UPDATE H.UENO Start
'''    Set rsKEIY_COUNT = dbMDB.OpenRecordset(strSQL, dbReadOnly)
    Set rsKEIY_COUNT = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
' 2007/04/23 UPDATE H.UENO End
    
    Do Until rsKEIY_COUNT.EOF
    
       strSQL = "SELECT * FROM WK_LQ0001_TEMP WHERE LQ0001_YARDC = " & rsKEIY_COUNT.Fields("CARG_YCODE") & ";"
       Set RST_Q001 = dbMDB.OpenRecordset(strSQL, dbOpenDynaset)
    
' 2007/04/23 UPDATE H.UENO Start
       If RST_Q001.BOF = False Then
' 2007/04/23 UPDATE H.UENO End
           With RST_Q001
               .Edit
' 2007/04/23 UPDATE H.UENO Start
'''       .Fields("LQ0001_KAIKQ") = Nz(rsKEIY_COUNT.Fields("COUNT").Value, 0)
               .Fields("LQ0001_KAIKQ") = Nz(rsKEIY_COUNT.Fields("CNT").VALUE, 0)
' 2007/04/23 UPDATE H.UENO End
               .Fields("LQ0001_KAIKA") = Nz(rsKEIY_COUNT.Fields("GOUKE").VALUE, 0)
               .UPDATE
           End With
' 2007/04/23 UPDATE H.UENO Start
       End If
' 2007/04/23 UPDATE H.UENO End
       RST_Q001.Close
       rsKEIY_COUNT.MoveNext
    Loop
    
    rsKEIY_COUNT.Close

    LF0001_KAI_CNT = True
    Exit Function

Err_LF0001_KAI_CNT:
    MsgBox Err.Description
    LF0001_KAI_CNT = False
End Function
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＯＰＥＮ
'        MODULE_ID        :LF0001_DBOPEN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_DBOPEN()
    On Error GoTo Err_LF0001_DBOPEN
    
    Set dbMDB = CurrentDb()
    
Exit_LF0001_DBOPEN:
    Exit Sub

Err_LF0001_DBOPEN:
    MsgBox Err.Description
    Resume Exit_LF0001_DBOPEN

End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＣＬＯＳＥ
'        MODULE_ID        :LF0001_DBCLOS
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0001_DBCLOS()
    On Error GoTo Err_LF0001_DBCLOS
    
    dbMDB.Close
    
    Set dbMDB = Nothing

Exit_LF0001_DBCLOS:
    Exit Sub

Err_LF0001_DBCLOS:
    MsgBox Err.Description
    Resume Exit_LF0001_DBCLOS

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

