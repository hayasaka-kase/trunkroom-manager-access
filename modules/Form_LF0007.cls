VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LF0007"
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
'        PROGRAM_NAME    : 稼動明細一覧表
'        PROGRAM_ID      : LF0007
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2004/03/02
'        CERATER         : N.MIURA
'
'        UPDATE          : 2006/03/07
'        UPDATER         : N.MIURA
'                        :
'
'        UPDATE          : 2007/05/18 V0.2
'        UPDATER         : S.SHIBAZAKI
'                        : 分析帳票強化対応
'
'        UPDATE          : 2007/11/02 V0.3
'        UPDATER         : S.SHIBAZAKI
'                        : 年月の範囲条件を追加
'
'        UPDATE          : 2016/02/10 V0.4
'        UPDATER         : M.HONDA
'                        : 所有区分を表示するように対応
'
'        UPDATE          : 2019/07/09 V0.5
'        UPDATER         : M.HONDA
'                        : 坪数・単価・金額をKASE_DBより取得するように変更
'
'        UPDATE          : 2019/10/23 V0.6
'        UPDATER         : M.HONDA
'                        : 満室時の金額を表示する。
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "LF0007"

'↓2007/05/18 S.SHIBAZAKI INSERT
Private Const WORK_TABLE = "LF0007_W01"
Private Const P_STEP_上段 = "0"
Private Const P_STEP_下段 = "1"
Private Const P_USE_撤去 = "9"
Private Const P_REPR_修理中 = "1"
Private Const P_AGRE_完了 = "9"
Private Const P_USAGE_NAME_ID = "086"
'ADO Constant
Private Const adOpenForwardOnly = 0
Private Const adLockReadOnly = 1
Private Const adCmdText = 1
'↑2007/05/18 S.SHIBAZAKI INSERT

Private USER_ID         As String
Private MODL_ID         As String
'
Private dbs             As Database
Private strSQL          As String
'
Private INPCHEK_STS     As Boolean

'↓2007/05/18 S.SHIBAZKI INSERT
'==============================================================================*
'
'        MODULE_NAME      :cmbYardFrom_AfterUpdate
'        MODULE_ID        :cmbYardFrom_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbYardFrom_AfterUpdate()

    txtYardFrom = fncComboAfterUpdate(cmbYardFrom, 6)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbYardFrom_BeforeUpdate
'        MODULE_ID        :cmbYardFrom_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbYardFrom_BeforeUpdate(Cancel As Integer)

    Call fncChkComboInput(cmbYardFrom, 6, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbYardFrom_KeyPress
'        MODULE_ID        :cmbYardFrom_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbYardFrom_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbYardFrom, KeyAscii, 6)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbYardTo_AfterUpdate
'        MODULE_ID        :cmbYardTo_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbYardTo_AfterUpdate()

    txtYardTo = fncComboAfterUpdate(cmbYardTo, 6)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbYardTo_BeforeUpdate
'        MODULE_ID        :cmbYardTo_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbYardTo_BeforeUpdate(Cancel As Integer)

    Call fncChkComboInput(cmbYardTo, 6, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbYardTo_KeyPress
'        MODULE_ID        :cmbYardTo_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbYardTo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbYardTo, KeyAscii, 6)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncChkComboInput
'        機能             :コンボボックス入力チェック
'        IN               :対象コンボボックス
'        IN               :①対象コンボボックス
'                         :②入力最大桁数
'                         :③更新後処理のキャンセル
'        OUT              :0=正常
'                          1=入力規則に反している
'                          2=入力値がリストに存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkComboInput(cmbCheck As ComboBox, intMaxLength As Integer, ByRef Cancel As Integer) As Integer

    Dim strCode As String
    Dim intRet  As Integer
    Dim intCount    As Integer  '06.02.3 SHIBAZAKI 追加
    Dim strValue    As String   '06.02.3 SHIBAZAKI 追加
    Dim blnFound    As Boolean  '06.02.3 SHIBAZAKI 追加
    
    intRet = 0
    
    strCode = Nz(cmbCheck, "")
    
    '空白なら入力チェックしない
    If strCode = "" Then
        GoTo ChkComboInput_End
    End If
    
    '入力チェック
    Select Case CmNumCheck(strCode, intMaxLength, 0)
        Case 3
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation
            intRet = 1
            GoTo ChkComboInput_End
        Case 2
            MsgBox "コードを入力してください。", vbInformation
            SendKeys "{ESC}"
            intRet = 1
            GoTo ChkComboInput_End
    End Select
    
    '入力値がリストに存在しているかチェック
    If cmbCheck.ListCount <> 0 Then
        If cmbCheck.ListIndex = -1 Then
            blnFound = False
            '入力値を０サプレス
            strValue = Format(cmbCheck, Mid("0000000000", 1, intMaxLength))
            'コンボボックス内にコードが存在してる？
            For intCount = 0 To cmbCheck.ListCount - 1
                'コード値比較
                If cmbCheck.Column(0, intCount) = strValue Then
                    blnFound = True
                    Exit For
                End If
            Next
        Else
            blnFound = True
        End If
        
        If Not blnFound Then
            intRet = 2
            MsgBox "該当するヤードマスタがありません。", vbInformation
            SendKeys "{ESC}"
            GoTo ChkComboInput_End
        End If
    End If
    
ChkComboInput_End:
    If intRet <> 0 Then
        Cancel = True
    End If
    
    fncChkComboInput = intRet

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncComboAfterUpdate
'        機能             :コンボボックス更新後イベント時の処理
'                         :０埋めした値を対象コントロールに設定する
'        IN               :①対象コントロール
'                         :②入力最大桁数(数値項目の場合のみ意味を持つ)
'        OUT              :名称
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncComboAfterUpdate(objTarget As Object, intLength As Integer) As String

    fncComboAfterUpdate = ""
    
    'ゼロサプレス
    objTarget = Format(objTarget, Mid("0000000000", 1, intLength))
    
    '名称を返却する
    If Nz(objTarget) <> "" And _
       objTarget.ListCount <> 0 _
    Then
        fncComboAfterUpdate = Nz(objTarget.Column(1))
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :フォームのクローズ
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
On Error Resume Next

    Dim dbAccess            As Database
    
    'ACCESS接続
    Set dbAccess = CurrentDb
    'ワークテーブル削除
    If TableExist(dbAccess) Then
        Call dbAccess.TableDefs.Delete(WORK_TABLE)
    End If
    
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing

End Sub
'↑2007/05/18 S.SHIBAZKI INSERT

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
    
    Call LF0007_DBOPEN
     Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル44.BackColor = RGB(255, 0, 204)
    End Select
    USER_ID = LsGetUserName()
    
    '↓INSERT 2007/11/02 SHIBAZAKI
    '対象年月の初期値は前月
    txtYearMonthFrom = Format(DateSerial(Year(Now()), Month(Now()) - 1, 1), "yyyy/mm")
    txtYearMonthTo = txtYearMonthFrom
    '↑INSERT 2007/11/02 SHIBAZAKI

Exit_Form_Open:
    Exit Sub

Err_Form_Open:
    MsgBox Err.Description
    Resume Exit_Form_Open
End Sub
'==============================================================================*
'
'        MODULE_NAME      :実行
'        MODULE_ID        :LF0007_EXE1B_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0007_EXE1B_Click()
    On Error GoTo Err_LF0007_EXE1B_Click
    
    '2007/05/18 S.SHIBAZAKI INSERT
    If fncMakeReport() Then
        doCmd.OpenReport "LR0007_M00", acPreview, "", ""
    End If
    
Exit_LF0007_EXE1B_Click:
    Exit Sub

Err_LF0007_EXE1B_Click:
'    MsgBox Err.Description                 '2007/05/18 S.SHIBAZAKI DELETE
'    Resume Exit_LF0007_EXE1B_Click         '2007/05/18 S.SHIBAZAKI DELETE

    '2007/05/18 S.SHIBAZAKI INSERT
    Call MSZZ024_M00("LF0007_EXE1B_Click", True)
    Err.Clear

End Sub
'==============================================================================*
'
'        MODULE_NAME      :EXCEL
'        MODULE_ID        :LF0007_EXE2B_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0007_EXE2B_Click()
    On Error GoTo Err_LF0007_EXE2B_Click
    
    '2007/05/18 S.SHIBAZAKI INSERT
    If fncMakeReport() Then
        '2007/05/18 S.SHIBAZAKI UPDATE
        On Error Resume Next
        doCmd.OutputTo acOutputTable, WORK_TABLE, "MicrosoftExcel(*.xls)", "", True, ""
        On Error GoTo Err_LF0007_EXE2B_Click
    End If

'    doCmd.OutputTo acQuery, "LQ0007_M00", "MicrosoftExcel(*.xls)", "", True, ""
   'doCmd.OutputTo acOutputTable, "WK_MQAM010_TEMP", "MicrosoftExcel(*.xls)", "", True, ""

Exit_LF0007_EXE2B_Click:
    Exit Sub

Err_LF0007_EXE2B_Click:
'    MsgBox Err.Description             '2007/05/18 S.SHIBAZAKI DELETE
'    Resume Exit_LF0007_EXE2B_Click     '2007/05/18 S.SHIBAZAKI DELETE
    
    '2007/05/18 S.SHIBAZAKI INSERT
    Call MSZZ024_M00("LF0007_EXE2B_Click", True)
    Err.Clear
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :閉じる
'        MODULE_ID        :LF0007_EXITB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0007_EXITB_Click()
    On Error GoTo Err_LF0007_EXITB_Click

    Call LF0007_DBCLOS
    
    Call MSHS010_M00(PROG_ID, "2")
    
    doCmd.Close

Exit_LF0007_EXITB_Click:
    Exit Sub

Err_LF0007_EXITB_Click:
    MsgBox Err.Description
    Resume Exit_LF0007_EXITB_Click
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
'    On Error GoTo Err_Form_Activate                                            'DELETE 20060307 N.MIURA
    
    doCmd.Maximize

'Exit_Form_Activate:                                                            'DELETE 20060307 N.MIURA START
'    Exit Sub
'Err_Form_Activate:
'    MsgBox err.Description
'    Resume Exit_Form_Activate                                                  'DELETE 20060307 N.MIURA ENDED
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＯＰＥＮ
'        MODULE_ID        :LF0007_DBOPEN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0007_DBOPEN()
    On Error GoTo Err_LF0007_DBOPEN
    
    Set dbs = CurrentDb()
    
Exit_LF0007_DBOPEN:
    Exit Sub

Err_LF0007_DBOPEN:
    MsgBox Err.Description
    Resume Exit_LF0007_DBOPEN
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＣＬＯＳＥ
'        MODULE_ID        :LF0007_DBCLOS
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0007_DBCLOS()
    On Error GoTo Err_LF0007_DBCLOS
    
    If Not dbs Is Nothing Then                                                  'INSERT 20060307 N.MIURA START
       dbs.Close
       Set dbs = Nothing
    End If                                                                      'INSERT 20060307 N.MIURA ENDED
    
    'DBS.Close                                                                  'DELETE 20060307 N.MIURA
    'Set DBS = Nothing                                                          'DELETE 20060307 N.MIURA

Exit_LF0007_DBCLOS:
    Exit Sub

Err_LF0007_DBCLOS:
    MsgBox Err.Description
    Resume Exit_LF0007_DBCLOS
End Sub

'↓2007/05/18 S.SHIBAZAKI INSERT
'==============================================================================*
'
'        MODULE_NAME      :fncMakeReport
'        機能             :レポート用ワークテーブル作成
'        IN               :
'        OUT              :True=正常 False=異常
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeReport() As Boolean
On Error GoTo ErrorHandler
    
    Dim dbSQLServer         As Object       'ADODB.Connection
    Dim rsSqlServer         As Object       'ADODB.Recordset
    Dim dbAccess            As Database
    Dim rsAccess            As Recordset
    Dim strBumonCode        As String
    Dim blnError            As Boolean
    
    fncMakeReport = False
    blnError = False
    
    doCmd.Hourglass True
    doCmd.SetWarnings False
    
    '↓INSERT 2007/11/02 SHIBAZAKI
    '必須チェック
    If Not fncRequireCheck() Then
        GoTo ExitHandler
    End If
    '↑INSERT 2007/11/02 SHIBAZAKI
    
    '部門コード取得
    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("DLookup(""CONT_BUMOC"", ""dbo_CONT_MAST"", ""CONT_KEY=1"")", "テーブル[CONT_MAST]の設定不正。")
    End If
    
    'SQL-Server接続
    Set dbSQLServer = CreateObject("ADODB.Connection")
    dbSQLServer.CommandTimeout = 180
    Call dbSQLServer.Open(MSZZ007_M10(strBumonCode))
    
    'ACCESS接続
    Set dbAccess = CurrentDb
    
    Call ClearWork(dbAccess)
    Call fncMakeSourceRecordSet(dbSQLServer, rsSqlServer)
    
    'ワークテーブルレコード追加
    Set rsAccess = dbAccess.OpenRecordset(WORK_TABLE, dbOpenDynaset, dbAppendOnly)
    
    With rsSqlServer
        If .EOF Then
            MsgBox "該当データが存在しません。", vbInformation
            GoTo ExitHandler
        End If
        While Not .EOF
            Call subAddWork(rsSqlServer, rsAccess)
            .MoveNext
        Wend
    End With
    
    fncMakeReport = True
    GoTo ExitHandler
    
ErrorHandler:
    blnError = True
    
ExitHandler:
    If Not rsSqlServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    If Not rsAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing

    doCmd.Hourglass False
    doCmd.SetWarnings True

    If blnError Then
        Call Err.Raise(Err.Number, "fncMakeReport" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :ClearWork
'        機能             :ワークテーブルクリア
'        IN               :ACCESS接続オブジェクト
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ClearWork(dbAccess As Database) As Boolean
On Error GoTo ExitHandler

    Dim tdfNew      As TableDef
    
    ClearWork = False
    
    'ワークテーブル削除
    If TableExist(dbAccess) Then
        Call dbAccess.TableDefs.Delete(WORK_TABLE)
    End If

    'ワークテーブル作成
    Set tdfNew = dbAccess.CreateTableDef(WORK_TABLE)
    Call FieldAppend(tdfNew)
    Call dbAccess.TableDefs.Append(tdfNew)
    
    ClearWork = True

ExitHandler:
    If Not tdfNew Is Nothing Then Set tdfNew = Nothing
    
    If Not ClearWork Then
        Call Err.Raise(Err.Number, "ClearWork" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :TableExist
'        機能             :ACCESSテーブル存在チェック
'        IN               :ACCESS接続オブジェクト
'        OUT              :True=存在する False=存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function TableExist(dbAccess As Database) As Boolean

    Dim tdf As TableDef
    
    TableExist = False
    
    For Each tdf In dbAccess.TableDefs
        If tdf.NAME = WORK_TABLE Then
            TableExist = True
            Exit For
        End If
    Next tdf

End Function

'==============================================================================*
'
'        MODULE_NAME      :FieldAppend
'        機能             :ワークテーブル列作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FieldAppend(tdfNew As TableDef)

    Dim fldNew      As Field
    Dim intCount    As Integer

    With tdfNew
        
        'INSERT 2007/11/02 SHIBAZAKI
        Call .Fields.Append(.CreateField("対象年月", DataTypeEnum.dbText, 16))
 
        Call .Fields.Append(.CreateField("ﾔｰﾄﾞｺｰﾄﾞ", DataTypeEnum.dbText, 6))
        Call .Fields.Append(.CreateField("ﾔｰﾄﾞ名", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("種別", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("用途", DataTypeEnum.dbText, 20))
        Call .Fields.Append(.CreateField("サイズ", DataTypeEnum.dbDouble))
        
        Call .Fields.Append(.CreateField("開始日", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("終了日", DataTypeEnum.dbText, 36))
        
        Call .Fields.Append(.CreateField("上段設置数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("上段使用数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("上段空数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("上段修理数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("上段撤去数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("上段売上", DataTypeEnum.dbLong))

        Call .Fields.Append(.CreateField("下段設置数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("下段使用数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("下段空数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("下段修理数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("下段撤去数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("下段売上", DataTypeEnum.dbLong))

        Call .Fields.Append(.CreateField("売上計", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("坪数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("単価", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("地代", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("物件No", DataTypeEnum.dbText, 6))
        
        '2019/10/23 M.HONDA INS
        Call .Fields.Append(.CreateField("上段価格", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("下段価格", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("上段満室", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("下段満室", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("総売上", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("キャンペーン1", DataTypeEnum.dbText, 80))
        Call .Fields.Append(.CreateField("キャンペーン2", DataTypeEnum.dbText, 80))
        Call .Fields.Append(.CreateField("キャンペーン3", DataTypeEnum.dbText, 80))
        Call .Fields.Append(.CreateField("条件", DataTypeEnum.dbText, 60))
        '2019/10/23 M.HONDA INS
        

        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount

    End With
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncMakeSourceRecordSet
'        機能             :SQLサーバーから元ネタ取得
'        IN               :①SQLサーバー接続オブジェクト(ADODB)
'                         :②レコードセットオブジェクト(ADODB)
'        OUT              :True=正常 False=異常
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeSourceRecordSet(dbSQLServer As Object, ByRef rsSqlServer As Object) As Boolean
On Error GoTo ExitHandler
    Dim strSQL              As String
    
    fncMakeSourceRecordSet = False
    
    Set rsSqlServer = Nothing
    
    'サーバー上にテンポラリテーブルを作成
    dbSQLServer.Execute fncMakeInsertTempSql(), adCmdText
    
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    rsSqlServer.Open fncMakeGetDataSql(), dbSQLServer, adOpenForwardOnly, adLockReadOnly
    
    fncMakeSourceRecordSet = True
    
ExitHandler:

    If Not fncMakeSourceRecordSet Then
        Call Err.Raise(Err.Number, "fncMakeSourceRecordSet" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeInsertTempSql
'        機能             :SQLサーバーにテンポラリテーブルを作成するSQL
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeInsertTempSql() As String
    
    Dim strSQL              As String
    
    '↓INSERT 2007/11/02 SHIBAZAKI
    Dim strFromFirstDay     As String
    Dim strToLastDay        As String
    Dim strTemp             As String
    Dim strDbName           As String
    Dim SET_BUMOC           As String
    
    strDbName = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME'"), "")
    SET_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    
    'FROMの年月のイッピ
    strFromFirstDay = fncFirstDay(txtYearMonthFrom)
    If strFromFirstDay = "" Then
        strFromFirstDay = "1753/01/01"
    End If
    
    'TOの年月の最終日
    strTemp = fncFirstDay(txtYearMonthTo)
    If strTemp = "" Then
        strToLastDay = "9999/12/31"
    Else
        strToLastDay = DateSerial(Year(strTemp), Month(strTemp) + 1, 0)
    End If
    '↑INSERT 2007/11/02 SHIBAZAKI
    
    strSQL = ""
    strSQL = strSQL & " SELECT YARD_CODE, "
    strSQL = strSQL & "        YARD_NAME, "
    strSQL = strSQL & "        CNTA_USAGE, "
    strSQL = strSQL & "        CNTA_SIZE, "
    
    strSQL = strSQL & fncMakeInsTempSelSql(True)
    strSQL = strSQL & fncMakeInsTempSelSql(False)
    
    strSQL = strSQL & "        C86.NAME_NAME AS USAGE_NAME, "
    strSQL = strSQL & "        NAME_C210.NAME_NAME AS SYOUI, "

    strSQL = strSQL & "        YARD_BEGIN_DAY, "
    strSQL = strSQL & "        YARD_END_DAY, "

    '2019/10/23 M.HONDA INS
    strSQL = strSQL & "        MAX(CASE WHEN CNTA_STEP = '0' THEN ISNULL(PRIC_PRICE,0) ELSE 0 END) AS CNTA_UP_PRICE, "
    strSQL = strSQL & "        MAX(CASE WHEN CNTA_STEP = '1' THEN ISNULL(PRIC_PRICE,0) ELSE 0 END) AS CNTA_DW_PRICE, "
    strSQL = strSQL & "        0 AS UP_PRICE_SUM, "
    strSQL = strSQL & "        0 AS DW_PRICE_SUM, "
    strSQL = strSQL & "        0 AS PRICE_SUM, "
    strSQL = strSQL & "        YARD_SEV1N, "
    strSQL = strSQL & "        YARD_SEV2N, "
    strSQL = strSQL & "        YARD_SEV3N, "
    strSQL = strSQL & "        YARD_ENDEN "
    '2019/10/23 M.HONDA INS
    
    strSQL = strSQL & "   INTO #WK_TEMP  "
    strSQL = strSQL & "   FROM YARD_MAST "
    strSQL = strSQL & "        INNER JOIN CNTA_MAST "
    strSQL = strSQL & "           ON CNTA_CODE = YARD_CODE "
    
    '2019/10/23 M.HONDA INS
    strSQL = strSQL & "   Left Join "
    strSQL = strSQL & "   PRIC_TABL ON "
    strSQL = strSQL & "   PRIC_YCODE =  CNTA_CODE AND "
    strSQL = strSQL & "   PRIC_USAGE = CNTA_USAGE AND "
    strSQL = strSQL & "   PRIC_SIZE = CNTA_SIZE AND "
    strSQL = strSQL & "   PRIC_STEP = CNTA_STEP "
    '2019/10/23 M.HONDA INS
    
    '↓DELETE 2007/11/02 SHIBAZAKI
'    strSQL = strSQL & "        LEFT OUTER JOIN CARG_FILE "
'    strSQL = strSQL & "          ON CARG_YCODE = CNTA_CODE "
'    strSQL = strSQL & "         AND CARG_NO = CNTA_NO "
'    strSQL = strSQL & "         AND CARG_AGRE != '" & P_AGRE_完了 & "' "
    '↑DELETE 2007/11/02 SHIBAZAKI
    
    '↓INSERT 2007/11/02 SHIBAZAKI
    '年月範囲内で使用中のコンテナを重複無しで取得して外部ＪＯＩＮ
    strSQL = strSQL & "        LEFT OUTER JOIN "
    strSQL = strSQL & "        ( " & fncMakeCargJoinSql(strFromFirstDay, strToLastDay) & " ) JOIN_CARG "
    strSQL = strSQL & "          ON JOIN_CARG.CARG_YCODE = CNTA_CODE "
    strSQL = strSQL & "         AND JOIN_CARG.CARG_NO = CNTA_NO "
    '↑INSERT 2007/11/02 SHIBAZAKI
    
    strSQL = strSQL & "        LEFT OUTER JOIN NAME_MAST C86 "
    strSQL = strSQL & "          ON C86.NAME_ID = '" & P_USAGE_NAME_ID & "' "
    strSQL = strSQL & "         AND C86.NAME_CODE = CNTA_USAGE "
    
    
    '2016/02/10 M.HONDA INS
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & strDbName & ".dbo.TINH_MAST "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  TINHM_BUMOC = '" & SET_BUMOC & "' "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  TINHM_TINTC = YARD_CODE "
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & strDbName & ".dbo.BUKE_MAST "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  TINHM_BUKHC = BUKEM_BUKHC "
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C210 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C210.NAME_CODE = BUKEM_SYOUI "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C210.NAME_ID   = '210' "
    '2016/02/10 M.HONDA INS
    
    
    'strSQL = strSQL & "  WHERE ISNULL(YARD_END_DAY,'9999/12/31') > GETDATE() "                 'DELETE 2007/11/02 SHIBAZAKI
    'ヤード営業終了日がFROMのイッピ以降
    strSQL = strSQL & "  WHERE ISNULL(YARD_END_DAY,'9999/12/31') > '" & strFromFirstDay & "' "  'INSERT 2007/11/02 SHIBAZAKI
    strSQL = strSQL & fncMakeBetween("YARD_CODE", Nz(cmbYardFrom, ""), Nz(cmbYardTo, ""))
    
    strSQL = strSQL & "  GROUP BY YARD_CODE, "
    strSQL = strSQL & "        YARD_NAME, "
    strSQL = strSQL & "        CNTA_SIZE, "
    strSQL = strSQL & "        YARD_BEGIN_DAY, "
    strSQL = strSQL & "        YARD_END_DAY, "
    strSQL = strSQL & "        CNTA_USAGE, "
    strSQL = strSQL & "        NAME_C210.NAME_NAME, "
    strSQL = strSQL & "        C86.NAME_NAME, "
    '2019/10/23 M.HONDA INS
    strSQL = strSQL & "        YARD_SEV1N, "
    strSQL = strSQL & "        YARD_SEV2N, "
    strSQL = strSQL & "        YARD_SEV3N, "
    strSQL = strSQL & "        YARD_ENDEN; "
    '2019/10/23 M.HONDA INS
    
    
    '売上を検索してテンポラリテーブル更新
    strSQL = strSQL & " UPDATE #WK_TEMP "
    strSQL = strSQL & "    SET #WK_TEMP.CNTA_UP_URIAGE = URIAGE.CNTA_UP_URIAGE, "
    strSQL = strSQL & "        #WK_TEMP.CNTA_LOW_URIAGE = URIAGE.CNTA_LOW_URIAGE, "
    '2019/10/23 M.HONDA INS
    strSQL = strSQL & "        #WK_TEMP.UP_PRICE_SUM = #WK_TEMP.CNTA_UP_PRICE * #WK_TEMP.CNTA_UP_ALL_CNT , "
    strSQL = strSQL & "        #WK_TEMP.DW_PRICE_SUM = #WK_TEMP.CNTA_DW_PRICE * #WK_TEMP.CNTA_LOW_ALL_CNT, "
    strSQL = strSQL & "        #WK_TEMP.PRICE_SUM = (#WK_TEMP.CNTA_UP_PRICE * #WK_TEMP.CNTA_UP_ALL_CNT + #WK_TEMP.CNTA_DW_PRICE * #WK_TEMP.CNTA_LOW_ALL_CNT) "
    '2019/10/23 M.HONDA INS
    strSQL = strSQL & "   FROM #WK_TEMP "
    strSQL = strSQL & "        INNER JOIN "
    '2007/11/02 SHIBAZAKI
'    strSQL = strSQL & "        ( " & fncMakeInsTempUriageSql() & " ) URIAGE "
    strSQL = strSQL & "        ( " & fncMakeInsTempUriageSql(strFromFirstDay, strToLastDay) & " ) URIAGE "
    
    strSQL = strSQL & "        ON #WK_TEMP.YARD_CODE  = URIAGE.CNTA_CODE "
    strSQL = strSQL & "       AND #WK_TEMP.CNTA_SIZE  = URIAGE.CNTA_SIZE "
    strSQL = strSQL & "       AND #WK_TEMP.CNTA_USAGE = URIAGE.CNTA_USAGE; "

    fncMakeInsertTempSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeBetween
'        機能             :範囲条件作成
'        IN               :第一引数－対象テーブルのカラム名
'                         :第二引数－範囲条件値FROM
'                         :第三引数－範囲条件値TO
'        OUT              :条件文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeBetween(strColName As String, strfrom As String, strTo As String) As String

    Dim strTemp     As String
    
    strTemp = ""
    
    If strfrom <> "" And strTo <> "" Then
        '共に空白ではない場合、
        If strfrom = strTo Then
            'FROMとTOが同一の場合、一致条件
            strTemp = " AND " & strColName & " = '" & strfrom & "' "
        Else
            'FROMとTOが異なる場合、BETWEEN条件
            strTemp = " AND " & strColName & " BETWEEN '" & strfrom & "' AND '" & strTo & "' "
        End If
    ElseIf strfrom <> "" Then
        'FROMのみの場合、それ以上であることが条件
        strTemp = " AND " & strColName & " >= '" & strfrom & "' "
    ElseIf strTo <> "" Then
        'TOのみの場合、それ以下であることが条件
        strTemp = " AND " & strColName & " <= '" & strTo & "' "
    End If
    
    fncMakeBetween = strTemp

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeInsTempSelSql
'        機能             :テンポラリテーブル作成時のSELECT区SQL
'        IN               :True=上段 False=下段
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeInsTempSelSql(blnUp As Boolean) As String

    Dim strSQL              As String
    Dim strAlias            As String
    Dim strStep             As String
    
    If blnUp Then
        strStep = P_STEP_上段
        strAlias = "UP"
    Else
        strStep = P_STEP_下段
        strAlias = "LOW"
    End If
    
    '総数(撤去数除く)
    strSQL = strSQL & " SUM( "
    strSQL = strSQL & "   CASE WHEN CNTA_STEP = '" & strStep & "' "
    strSQL = strSQL & "         AND CNTA_USE != '" & P_USE_撤去 & "' "
    strSQL = strSQL & "        THEN 1 ELSE 0 END "
    strSQL = strSQL & " ) AS CNTA_" & strAlias & "_ALL_CNT, "
    
    '使用数
    strSQL = strSQL & " SUM( "
    strSQL = strSQL & "   CASE WHEN CNTA_STEP = '" & strStep & "' "
  '  strSQL = strSQL & "         AND CNTA_USE != '" & P_USE_撤去 & "' " dell 2007/11/02 SHIBAZAKI
    strSQL = strSQL & "         AND CARG_NO IS NOT NULL "
    strSQL = strSQL & "        THEN 1 ELSE 0 END "
    strSQL = strSQL & " ) AS CNTA_" & strAlias & "_USED_CNT, "
    
    '空きの内の修理数
    strSQL = strSQL & " SUM( "
    strSQL = strSQL & "   CASE WHEN CNTA_STEP = '" & strStep & "' "
    strSQL = strSQL & "         AND CNTA_USE != '" & P_USE_撤去 & "' "
    strSQL = strSQL & "         AND CNTA_REPR_KBN = '" & P_REPR_修理中 & "' "
    strSQL = strSQL & "         AND CARG_NO IS NULL "
    strSQL = strSQL & "        THEN 1 ELSE 0 END "
    strSQL = strSQL & " ) AS CNTA_" & strAlias & "_REPR_CNT, "
    
    '撤去数
    strSQL = strSQL & " SUM( "
    strSQL = strSQL & "   CASE WHEN CNTA_STEP = '" & strStep & "' "
    strSQL = strSQL & "         AND CNTA_USE = '" & P_USE_撤去 & "' "
    strSQL = strSQL & "        THEN 1 ELSE 0 END "
    strSQL = strSQL & " ) AS CNTA_" & strAlias & "_REMOVE, "
    
    '売上
    strSQL = strSQL & " 0 AS CNTA_" & strAlias & "_URIAGE, "

    fncMakeInsTempSelSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeInsTempSelSql
'        機能             :テンポラリテーブル作成時の売上検索SQL
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2007/11/02 SHIBAZAKI
'Private Function fncMakeInsTempUriageSql() As String
Private Function fncMakeInsTempUriageSql(strDateFrom As String, strDateTo As String) As String

    Dim strSQL              As String
'    Dim dtLastOfMonth       As Date                                'DELETE 2007/11/02 SHIBAZAKI

'    dtLastOfMonth = DateSerial(Year(Now()), Month(Now()) + 1, 0)   'DELETE 2007/11/02 SHIBAZAKI
    
    strSQL = strSQL & " SELECT CNTA_CODE, "
    strSQL = strSQL & "        CNTA_SIZE, "
    strSQL = strSQL & "        CNTA_USAGE, "
    strSQL = strSQL & "        SUM( "
    strSQL = strSQL & "          CASE WHEN CNTA_STEP = '" & P_STEP_上段 & "' "
    strSQL = strSQL & "               THEN REQU_KINGAK  ELSE 0 END "
    strSQL = strSQL & "        ) AS CNTA_UP_URIAGE, "
    strSQL = strSQL & "        SUM( "
    strSQL = strSQL & "          CASE WHEN CNTA_STEP = '" & P_STEP_下段 & "' "
    strSQL = strSQL & "               THEN REQU_KINGAK  ELSE 0 END "
    strSQL = strSQL & "        ) AS CNTA_LOW_URIAGE "
    
    strSQL = strSQL & "   FROM CNTA_MAST "
    strSQL = strSQL & "        INNER JOIN REQU_FILE "
    strSQL = strSQL & "           ON REQU_YCODE = CNTA_CODE "
    strSQL = strSQL & "          AND REQU_NO = CNTA_NO "
    
    strSQL = strSQL & "  WHERE EXISTS( SELECT * "
    strSQL = strSQL & "                  FROM #WK_TEMP "
    strSQL = strSQL & "                 WHERE YARD_CODE = CNTA_CODE "
    strSQL = strSQL & "        ) "
    
    '↓INSERT 2007/11/02 SHIBAZAKI
    '請求対象年月が指定条件の範囲内
    strSQL = strSQL & "    AND REQU_KJDATE BETWEEN '" & strDateFrom & "' "
    strSQL = strSQL & "                        AND '" & strDateTo & "' "
    '↑INSERT 2007/11/02 SHIBAZAKI
    '↓DELETE 2007/11/02 SHIBAZAKI
'    strSQL = strSQL & "    AND REQU_KJDATE BETWEEN '" & Format(Now(), "yyyy/mm/01") & "' "
'    strSQL = strSQL & "                        AND '" & Format(dtLastOfMonth, "yyyy/mm/dd") & "' "
    '↑DELETE 2007/11/02 SHIBAZAKI
    
    strSQL = strSQL & "  GROUP BY CNTA_CODE, CNTA_SIZE, CNTA_USAGE "

    fncMakeInsTempUriageSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeGetDataSql
'        機能             :テンポラリテーブルを検索するSQL
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSql() As String
    
    Dim strSQL              As String

    '↓INSERT 2007/11/02 SHIBAZAKI
    Dim strFromFirstDay     As String
    Dim strToLastDay        As String
    Dim strTemp             As String
    Dim strDataSource       As String
    'FROMの年月のイッピ
    strFromFirstDay = fncFirstDay(txtYearMonthFrom)
    If strFromFirstDay = "" Then
        strFromFirstDay = "1753/01/01"
    End If
    
    'TOの年月の最終日
    strTemp = fncFirstDay(txtYearMonthTo)
    If strTemp = "" Then
        strToLastDay = "9999/12/31"
    Else
        strToLastDay = DateSerial(Year(strTemp), Month(strTemp) + 1, 0)
    End If
    '↑INSERT 2007/11/02 SHIBAZAKI

    strSQL = ""

'    strSQL = strSQL & " SELECT TEMP.*, YARD_RENTGK as CHIDAI "
'    strSQL = strSQL & "        ( SELECT SUM(JARG_RENTKG) "
'    strSQL = strSQL & "            FROM JARG_FILE "
''    strSQL = strSQL & "           WHERE ISNULL(JARG_KYDATE,'9999/12/31') > GETDATE() "     'DELETE 2007/11/02 SHIBAZAKI
'    '↓INSERT 2007/11/02 SHIBAZAKI
'    '契約開始日がTOの最終日以前でかつ、解約日がFROMのイッピ以降
'    strSQL = strSQL & "           WHERE JARG_STDATE <= '" & strToLastDay & "' "
'    strSQL = strSQL & "             AND ISNULL(JARG_KYDATE,'9999/12/31') >= '" & strFromFirstDay & "' "
'    '↑INSERT 2007/11/02 SHIBAZAKI
'    strSQL = strSQL & "             AND JARG_YCODE = YARD_CODE "
    
'    strSQL = strSQL & "        ) AS CHIDAI "
'    strSQL = strSQL & "   FROM #WK_TEMP TEMP "
'    strSQL = strSQL & "        INNER JOIN YARD_MAST A ON "
'    strSQL = strSQL & "        A.YARD_CODE = TEMP.YARD_CODE "
'    strSQL = strSQL & "  ORDER BY TEMP.YARD_CODE, "
'    strSQL = strSQL & "           TEMP.CNTA_USAGE, "
'    strSQL = strSQL & "           TEMP.CNTA_SIZE; "
    
     strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    
    strSQL = strSQL & "   SELECT "
    strSQL = strSQL & "       TEMP.* "
    'strSql = strSql & "       ,HKEIT_KEIYQ, "
    'strSql = strSql & "       HKEIT_KEIYU, "
    'strSql = strSql & "       HKEIT_KEIYA "
   ' strSql = strSql & "       ,HKEIT_BUKHC "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       #WK_TEMP TEMP "
    'strSql = strSql & "       Left Join "
    'strSql = strSql & strDataSource & ".dbo.HKEI_TRAN ON "
    'strSql = strSql & "           HKEIT_SYO1C = CONVERT(varchar,YARD_CODE) AND "
    'strSql = strSql & "           HKEIT_KAIYD IS NULL "
    strSQL = strSQL & "   ORDER BY "
    strSQL = strSQL & "       TEMP.YARD_CODE, "
    strSQL = strSQL & "       TEMP.CNTA_USAGE, "
    strSQL = strSQL & "       TEMP.CNTA_SIZE; "

    fncMakeGetDataSql = strSQL
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :subAddWork
'        機能             :作業テーブル追加
'        IN               :①元ネタ
'                         :②ワークテーブルレコードセット
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subAddWork(rsSource As Object, rsDestination As Recordset)

    With rsSource
        rsDestination.AddNew
        
        '2007/11/02 S.SHIBAZAKI INSERT
        rsDestination.Fields("対象年月") = txtYearMonthFrom & "～" & txtYearMonthTo
        
        rsDestination.Fields("ﾔｰﾄﾞｺｰﾄﾞ") = Format(.Fields("YARD_CODE"), "000000")
        rsDestination.Fields("ﾔｰﾄﾞ名") = .Fields("YARD_NAME")
        rsDestination.Fields("用途") = Nz(.Fields("USAGE_NAME"), "")
        rsDestination.Fields("サイズ") = .Fields("CNTA_SIZE")
                
        rsDestination.Fields("種別") = .Fields("SYOUI")
        
        rsDestination.Fields("開始日") = .Fields("YARD_BEGIN_DAY")
        rsDestination.Fields("終了日") = .Fields("YARD_END_DAY")
                
        rsDestination.Fields("上段設置数") = .Fields("CNTA_UP_ALL_CNT")
        rsDestination.Fields("上段使用数") = .Fields("CNTA_UP_USED_CNT")
        rsDestination.Fields("上段空数") = .Fields("CNTA_UP_ALL_CNT") - .Fields("CNTA_UP_USED_CNT")
        rsDestination.Fields("上段修理数") = .Fields("CNTA_UP_REPR_CNT")
        rsDestination.Fields("上段撤去数") = .Fields("CNTA_UP_REMOVE")
        rsDestination.Fields("上段売上") = .Fields("CNTA_UP_URIAGE")
        
        rsDestination.Fields("下段設置数") = .Fields("CNTA_LOW_ALL_CNT")
        rsDestination.Fields("下段使用数") = .Fields("CNTA_LOW_USED_CNT")
        rsDestination.Fields("下段空数") = .Fields("CNTA_LOW_ALL_CNT") - .Fields("CNTA_LOW_USED_CNT")
        rsDestination.Fields("下段修理数") = .Fields("CNTA_LOW_REPR_CNT")
        rsDestination.Fields("下段撤去数") = .Fields("CNTA_LOW_REMOVE")
        rsDestination.Fields("下段売上") = .Fields("CNTA_LOW_URIAGE")
        
        rsDestination.Fields("売上計") = .Fields("CNTA_UP_URIAGE") + .Fields("CNTA_LOW_URIAGE")
    '      rsDestination.Fields("坪数") = Nz(.Fields("HKEIT_KEIYQ"), 0)
     '     rsDestination.Fields("単価") = Nz(.Fields("HKEIT_KEIYU"), 0)
      '    rsDestination.Fields("地代") = Nz(.Fields("HKEIT_KEIYA"), 0)
    '    rsDestination.Fields("物件No") = Nz(.Fields("HKEIT_BUKHC"), 0)
        
        rsDestination.Fields("上段価格") = Nz(.Fields("CNTA_UP_PRICE"), 0)
        rsDestination.Fields("下段価格") = Nz(.Fields("CNTA_DW_PRICE"), 0)
        rsDestination.Fields("上段満室") = Nz(.Fields("UP_PRICE_SUM"), 0)
        rsDestination.Fields("下段満室") = Nz(.Fields("DW_PRICE_SUM"), 0)
        rsDestination.Fields("総売上") = Nz(.Fields("PRICE_SUM"), 0)
        rsDestination.Fields("キャンペーン1") = Nz(.Fields("YARD_SEV1N"), 0)
        rsDestination.Fields("キャンペーン2") = Nz(.Fields("YARD_SEV2N"), 0)
        rsDestination.Fields("キャンペーン3") = Nz(.Fields("YARD_SEV3N"), 0)
        rsDestination.Fields("条件") = Nz(.Fields("YARD_ENDEN"), 0)
        
        rsDestination.UPDATE
    End With
    
End Sub
'↑2007/05/18 S.SHIBAZAKI INSERT

'↓2007/11/02 S.SHIBAZAKI INSERT
'==============================================================================*
'
'        MODULE_NAME      :txtYearMonthFrom_AfterUpdate
'        MODULE_ID        :txtYearMonthFrom_AfterUpdate
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtYearMonthFrom_AfterUpdate()

    txtYearMonthFrom = Format(CmDateChecker(fncFirstDay(txtYearMonthFrom)), "yyyy/mm")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtYearMonthFrom_BeforeUpdate
'        MODULE_ID        :txtYearMonthFrom_BeforeUpdate
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtYearMonthFrom_BeforeUpdate(Cancel As Integer)
        
    Call fncDateCheck(fncFirstDay(txtYearMonthFrom), Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtYearMonthFrom_GotFocus
'        MODULE_ID        :txtYearMonthFrom_GotFocus
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtYearMonthFrom_GotFocus()

    txtYearMonthFrom = Format(fncFirstDay(txtYearMonthFrom), "yy/mm")
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtYearMonthFrom_KeyPress
'        MODULE_ID        :txtYearMonthFrom_KeyPress
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtYearMonthFrom_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtYearMonthFrom, KeyAscii, 7)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtYearMonthFrom_LostFocus
'        MODULE_ID        :txtYearMonthFrom_LostFocus
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtYearMonthFrom_LostFocus()

    txtYearMonthFrom = Format(fncFirstDay(txtYearMonthFrom), "yyyy/mm")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtYearMonthTo_AfterUpdate
'        MODULE_ID        :txtYearMonthTo_AfterUpdate
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtYearMonthTo_AfterUpdate()

    txtYearMonthTo = Format(CmDateChecker(fncFirstDay(txtYearMonthTo)), "yyyy/mm")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtYearMonthTo_BeforeUpdate
'        MODULE_ID        :txtYearMonthTo_BeforeUpdate
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtYearMonthTo_BeforeUpdate(Cancel As Integer)
        
    Call fncDateCheck(fncFirstDay(txtYearMonthTo), Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtYearMonthTo_GotFocus
'        MODULE_ID        :txtYearMonthTo_GotFocus
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtYearMonthTo_GotFocus()

    txtYearMonthTo = Format(fncFirstDay(txtYearMonthTo), "yy/mm")
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtYearMonthTo_KeyPress
'        MODULE_ID        :txtYearMonthTo_KeyPress
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtYearMonthTo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtYearMonthTo, KeyAscii, 7)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtYearMonthTo_LostFocus
'        MODULE_ID        :txtYearMonthTo_LostFocus
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtYearMonthTo_LostFocus()

    txtYearMonthTo = Format(fncFirstDay(txtYearMonthTo), "yyyy/mm")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncFirstDay
'        機能             :月初日
'        IN               :対象年月のコントロール
'        OUT              :月初日
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncFirstDay(objTarget As Object) As String

    If Nz(objTarget) = "" Then
        fncFirstDay = ""
    Else
        fncFirstDay = objTarget & "/01"
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncDateCheck
'        機能             :日付入力チェック
'        IN               :①対象テキスト
'                          ②更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDateCheck(txtTarget As Variant, ByRef Cancel As Integer) As Boolean

    fncDateCheck = True
    
    If IsNull(CmDateChecker(txtTarget)) Then
        MsgBox "日付を正しく入力してください。", vbInformation, Me.NAME
        Cancel = True
        SendKeys "{ESC}"
        fncDateCheck = False
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncRequireCheck
'        機能             :必須入力チェック
'        IN               :
'        OUT              :True=必須入力が全て入力されている False=されていない
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncRequireCheck() As Boolean

    fncRequireCheck = False
    
    If Not fncRequireControl(txtYearMonthFrom, "対象年月FROM") Then Exit Function
    If Not fncRequireControl(txtYearMonthTo, "対象年月TO") Then Exit Function
    
    fncRequireCheck = True
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncRequireControl
'        機能             :必須チェック
'        IN               :①対象コントロール　②エラーメッセージに出す文言
'        OUT              :True=入力されている False=されていない
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncRequireControl(objTarget As Object, strMsgText As String) As Boolean

    fncRequireControl = True
    
    If objTarget.Enabled Then
        If Nz(objTarget.VALUE) = "" Then
            MsgBox strMsgText & "は必須入力です。", vbInformation, Me.NAME
            objTarget.SetFocus
            fncRequireControl = False
        End If
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeCargJoinSql
'        機能             :コンテナ契約トランを検索するSQL作成
'        IN               :①FROM年月のイッピ
'                         :②TO年月の最終日
'        OUT              :SQL文
'        CREATE_DATE      :2007/11/02 S.SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeCargJoinSql(strDateFrom As String, strDateTo As String) As String
    
    Dim strSQL              As String
    
    '年月範囲内で使用中のコンテナを重複無しで取得
    strSQL = strSQL & " SELECT CARG_YCODE, "
    strSQL = strSQL & "        CARG_NO "
    strSQL = strSQL & "   FROM CARG_FILE "
    
    '契約開始日がTOの最終日以前でかつ、解約日がFROMのイッピ以降
    strSQL = strSQL & "  WHERE CARG_STDATE <= '" & strDateTo & "' "
    strSQL = strSQL & "    AND ISNULL(CARG_KYDATE,'9999/12/31') >= '" & strDateFrom & "' "
    
    strSQL = strSQL & fncMakeBetween("CARG_YCODE", Nz(cmbYardFrom, ""), Nz(cmbYardTo, ""))
    
    strSQL = strSQL & "  GROUP BY CARG_YCODE, CARG_NO "

    fncMakeCargJoinSql = strSQL

End Function

'↑2007/11/02 S.SHIBAZAKI INSERT

'****************************  ended or program ********************************
