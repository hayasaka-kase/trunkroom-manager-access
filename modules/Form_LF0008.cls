VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LF0008"
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
'        PROGRAM_NAME    : 予約集計表
'        PROGRAM_ID      : LF0008
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2007/05/22 V0.0
'        CERATER         : S.SHIBAZAKI
'
'        UPDATE          : 2007/09/09 V0.1
'        UPDATER         : tajima
'                        : ネット予約の件数カウントの見方修正
'
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const WORK_TABLE = "LF0008_W01"
Private Const REPORT_NAME = "LR0008"

'発生区分
Private Const P_GENKBN_電話 = "0"
Private Const P_GENKBN_インターネット = "1"

'予約受付状態
Private Const P_YUKBN_自動紹介 = "0"
Private Const P_YUKBN_自社理由 = "1"
Private Const P_YUKBN_取置 = "2"
Private Const P_YUKBN_受付 = "10"
Private Const P_YUKBN_契約 = "20"
Private Const P_YUKBN_予約キャンセル = "51"
Private Const P_YUKBN_取置キャンセル = "52"
Private Const P_YUKBN_受付キャンセル = "53"

'移動区分
Private Const P_MOVEKBN_無し = "0"

'ADO Constant
Private Const adOpenForwardOnly = 0
Private Const adLockReadOnly = 1

'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'変更前処理で入力した日付を保存する
Private VarChangeDate As Variant

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
    Dim intCount    As Integer
    Dim strValue    As String
    Dim blnFound    As Boolean
    
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
            MsgBox "桁数オーバーしています。", vbInformation, Me.NAME
            intRet = 1
            GoTo ChkComboInput_End
        Case 2
            MsgBox "コードを入力してください。", vbInformation, Me.NAME
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
            MsgBox "該当するヤードマスタがありません。", vbInformation, Me.NAME
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
'        MODULE_NAME      :cmdClose_Click
'        MODULE_ID        :cmdClose_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdClose_Click()

    On Error GoTo ErrorHandler
    
    Call MSHS010_M00(Me.NAME, "2")
    
    doCmd.Close

    Exit Sub

ErrorHandler:
    MsgBox Err.Description
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdExcel_Click
'        MODULE_ID        :cmdExcel_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdExcel_Click()

    On Error GoTo ErrorHandler
    
    If fncMakeReport() Then
        On Error Resume Next
        doCmd.OutputTo acOutputTable, WORK_TABLE, "MicrosoftExcel(*.xls)", "", True, ""
        On Error GoTo ErrorHandler
    End If

    Exit Sub

ErrorHandler:
    Call MSZZ024_M00("cmdExcel_Click", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdPrint_Click
'        MODULE_ID        :cmdPrint_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdPrint_Click()
    
    On Error GoTo ErrorHandler
    
    If fncMakeReport() Then
        doCmd.OpenReport REPORT_NAME, acPreview, "", ""
    End If
    
    Exit Sub

ErrorHandler:
    Call MSZZ024_M00("cmdPrint_Click", True)
    Err.Clear

End Sub

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
    
    Dim dtTemp      As Date
    
    Call MSHS010_M00(Me.NAME, "1")
    
    dtTemp = Now()
    txtDateFrom = Format(DateSerial(Year(dtTemp), Month(dtTemp) - 1, 1), "yyyy/mm/dd")
    txtDateTo = Format(DateSerial(Year(dtTemp), Month(dtTemp), 0), "yyyy/mm/dd")
    
Exit_Form_Open:
    Exit Sub

Err_Form_Open:
    MsgBox Err.Description
    Resume Exit_Form_Open
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
    
    doCmd.Maximize

End Sub

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
    
    'ワークテーブルクリア
    Call ClearWork(dbAccess)
    
    'レコードセット作成
    Call fncMakeSourceRecordSet(dbSQLServer, rsSqlServer)
    
    'ワークテーブルレコード追加
    Set rsAccess = dbAccess.OpenRecordset(WORK_TABLE, dbOpenDynaset, dbAppendOnly)
    
    With rsSqlServer
        If .EOF Then
            MsgBox "該当データが存在しません。", vbInformation, Me.NAME
            GoTo ExitHandler
        End If
        
        While Not .EOF
            'ワークのレコードセットをAddNew
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
 
        Call .Fields.Append(.CreateField("対象日FROM", DataTypeEnum.dbText, 10))
        Call .Fields.Append(.CreateField("対象日TO", DataTypeEnum.dbText, 10))
        Call .Fields.Append(.CreateField("ヤードコード", DataTypeEnum.dbText, 6))
        Call .Fields.Append(.CreateField("ヤード名", DataTypeEnum.dbText, 36))
        
        Call .Fields.Append(.CreateField("自動紹介(電話)予約数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("自動紹介(電話)受付数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("自動紹介(電話)契約数", DataTypeEnum.dbLong))

        Call .Fields.Append(.CreateField("自動紹介(ｲﾝﾀｰﾈｯﾄ)予約数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("自動紹介(ｲﾝﾀｰﾈｯﾄ)受付数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("自動紹介(ｲﾝﾀｰﾈｯﾄ)契約数", DataTypeEnum.dbLong))
        
        Call .Fields.Append(.CreateField("手動紹介(電話)受付数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("手動紹介(電話)契約数", DataTypeEnum.dbLong))
        
        Call .Fields.Append(.CreateField("総受付数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("総契約数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("解約数", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("差引本数", DataTypeEnum.dbLong))
        
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
    
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    rsSqlServer.Open fncMakeSql(), dbSQLServer, adOpenForwardOnly, adLockReadOnly
    
    fncMakeSourceRecordSet = True
    
ExitHandler:

    If Not fncMakeSourceRecordSet Then
        Call Err.Raise(Err.Number, "fncMakeSourceRecordSet" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeSql
'        機能             :SQLサーバーからデータを取得するSQL
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeSql() As String
    
    Dim strSQL              As String
    Dim strCondYard         As String
    Dim strCondDate         As String
    Dim strFromTemp         As String
    Dim strToTemp           As String
    Dim str予約数日付範囲   As String   '2007/09/09 tajima add
        
    'ヤードコードの範囲条件
    strFromTemp = IIf(Nz(cmbYardFrom, "") = "", "000000", cmbYardFrom)
    strToTemp = IIf(Nz(cmbYardTo, "") = "", "999999", cmbYardTo)
    strCondYard = fncMakeBetween(strFromTemp, strToTemp)
    
    '日付の範囲条件
    strFromTemp = IIf(Nz(txtDateFrom, "") = "", "1900/01/01", txtDateFrom)
    strToTemp = IIf(Nz(txtDateTo, "") = "", "9999/12/31", txtDateTo)
    strCondDate = fncMakeBetween(Format(strFromTemp, "yyyy/mm/dd"), Format(strToTemp, "yyyy/mm/dd"))
    str予約数日付範囲 = fncMakeBetween(Format(strFromTemp, "yyyy/mm/dd") & " 00:00:00", Format(strToTemp, "yyyy/mm/dd") & " 23:59:59") '2007/09/09 tajima add
    
    strSQL = ""
    strSQL = strSQL & " SELECT YARD_CODE, "
    strSQL = strSQL & "        YARD_NAME ,  "
    strSQL = strSQL & "        TEMP_REQ.*, "
    strSQL = strSQL & "        TEMP_UKE.*, "
    strSQL = strSQL & "        TEMP_KEI.*, "
    strSQL = strSQL & "        TEMP_KAIYAKU.* "
    strSQL = strSQL & "   FROM YARD_MAST "
    
    '予約数
    strSQL = strSQL & "        LEFT OUTER JOIN "
    strSQL = strSQL & "            ( " & fncMakeReqSql(strCondYard, str予約数日付範囲) & " ) TEMP_REQ "  '2007/09/09 tajima add
    strSQL = strSQL & "          ON TEMP_REQ.YOUKT_REQ_YCODE = YARD_CODE "
    
    '受付数
    strSQL = strSQL & "        LEFT OUTER JOIN  "
    strSQL = strSQL & "            ( " & fncMakeUkeSql(strCondYard, strCondDate) & " ) TEMP_UKE "
    strSQL = strSQL & "          ON TEMP_UKE.YOUKT_YCODE = YARD_CODE "
    
    '契約数
    strSQL = strSQL & "        LEFT OUTER JOIN "
    strSQL = strSQL & "            (  " & fncMakeKeiSql(strCondYard, strCondDate) & " ) TEMP_KEI "
    strSQL = strSQL & "          ON TEMP_KEI.CARG_YCODE = YARD_CODE "
    
    '解約数
    strSQL = strSQL & "        LEFT OUTER JOIN "
    strSQL = strSQL & "            ( " & fncMakeKaiyakuSql(strCondYard, strCondDate) & " ) TEMP_KAIYAKU "
    strSQL = strSQL & "          ON TEMP_KAIYAKU.CARG_YCODE = YARD_CODE "
    
    strSQL = strSQL & "  WHERE ISNULL(YARD_END_DAY,'9999/12/31') > GETDATE() "
    strSQL = strSQL & "    AND YARD_CODE " & strCondYard

    fncMakeSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeReqSql
'        機能             :予約数検索SQL作成
'        IN               :第一引数－ﾔｰﾄﾞｺｰﾄﾞの条件
'                         :第二引数－日付の条件
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeReqSql(strCondYard As String, strCondDate As String) As String

    Dim strSQL              As String
    
    strSQL = strSQL & " SELECT YOUKT_REQ_YCODE, "
    
    '自動紹介（電話）予約数
    strSQL = strSQL & "        SUM( CASE WHEN YOUKT_GENKBN = " & P_GENKBN_電話
    strSQL = strSQL & "                   AND YOUKT_ORG_YUKBN = " & P_YUKBN_自動紹介  '2007/09/09 add tajima
    strSQL = strSQL & "                   AND YOUKT_YUKBN != " & P_YUKBN_予約キャンセル  '2007/09/09 add tajima
    strSQL = strSQL & "                  THEN 1 ELSE 0 "
    strSQL = strSQL & "             END ) AS AUTO_TEL_REQ, "
    
    '自動紹介（インターネット）予約数
    strSQL = strSQL & "        SUM( CASE WHEN YOUKT_GENKBN = " & P_GENKBN_インターネット
    strSQL = strSQL & "                   AND YOUKT_ORG_YUKBN IN(" & P_YUKBN_受付 & "," & P_YUKBN_予約キャンセル & ")" '2007/09/09 add tajima
    strSQL = strSQL & "                  THEN 1 ELSE 0 "
    strSQL = strSQL & "             END ) AS AUTO_INET_REQ "
    
    strSQL = strSQL & "   FROM YOUK_TRAN "
    
    strSQL = strSQL & "  WHERE YOUKT_REQ_YCODE " & strCondYard
    strSQL = strSQL & "    AND YOUKT_UKDATE " & strCondDate
    strSQL = strSQL & "    AND YOUKT_GENKBN IN(" & P_GENKBN_電話 & "," & P_GENKBN_インターネット & ") "
'▼2007/09/09 ↓fix tajima    strSQL = strSQL & "    AND YOUKT_ORG_YUKBN = " & P_YUKBN_自動紹介
    strSQL = strSQL & "    AND YOUKT_ORG_YUKBN IN( " & P_YUKBN_自動紹介 & "," & P_YUKBN_受付 & "," & P_YUKBN_予約キャンセル & ")"
'   del strSQL = strSQL & "    AND YOUKT_YUKBN != " & P_YUKBN_予約キャンセル
'▲2007/09/09 ↑
    strSQL = strSQL & "  GROUP BY YOUKT_REQ_YCODE "

    fncMakeReqSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeUkeSql
'        機能             :受付数数検索SQL作成
'        IN               :第一引数－ﾔｰﾄﾞｺｰﾄﾞの条件
'                         :第二引数－日付の条件
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeUkeSql(strCondYard As String, strCondDate As String) As String

    Dim strSQL              As String
    
    strSQL = strSQL & " SELECT YOUKT_YCODE, "
    
    '自動紹介（電話）受付数
    strSQL = strSQL & "        SUM( CASE WHEN YOUKT_GENKBN = " & P_GENKBN_電話
    strSQL = strSQL & "                   AND YOUKT_ORG_YUKBN = " & P_YUKBN_自動紹介
    strSQL = strSQL & "                  THEN 1 ELSE 0 "
    strSQL = strSQL & "             END ) AS AUTO_TEL_UKE, "
    
    '自動紹介（インターネット）受付数
    strSQL = strSQL & "        SUM( CASE WHEN YOUKT_GENKBN = " & P_GENKBN_インターネット
    strSQL = strSQL & "                   AND YOUKT_ORG_YUKBN = " & P_YUKBN_受付 '2007/09/09 fix tajima old:P_YUKBN_自動紹介
    strSQL = strSQL & "                   AND YOUKT_YUKBN != " & P_YUKBN_受付キャンセル '2007/09/09 add tajima
    strSQL = strSQL & "                  THEN 1 ELSE 0 "
    strSQL = strSQL & "             END ) AS AUTO_INET_UKE, "
    
    '手動紹介（電話）受付数
    strSQL = strSQL & "        SUM( CASE WHEN YOUKT_GENKBN = " & P_GENKBN_電話
    strSQL = strSQL & "                   AND YOUKT_ORG_YUKBN != " & P_YUKBN_自動紹介
    strSQL = strSQL & "                  THEN 1 ELSE 0 "
    strSQL = strSQL & "             END ) AS MANUAL_TEL_UKE "
    
    strSQL = strSQL & "   FROM YOUK_TRAN "
    
    strSQL = strSQL & "  WHERE YOUKT_YCODE " & strCondYard
    strSQL = strSQL & "    AND YOUKT_KUDATE " & strCondDate
    strSQL = strSQL & "    AND YOUKT_GENKBN IN(" & P_GENKBN_電話 & "," & P_GENKBN_インターネット & ") "
    strSQL = strSQL & "    AND YOUKT_ORG_YUKBN IN(" & P_YUKBN_自動紹介 & "," _
                                                    & P_YUKBN_自社理由 & "," _
                                                    & P_YUKBN_取置 & "," _
                                                    & P_YUKBN_受付 & ") "
    strSQL = strSQL & "    AND YOUKT_YUKBN IN(" & P_YUKBN_受付 & "," _
                                                & P_YUKBN_契約 & "," _
                                                & P_YUKBN_受付キャンセル & ") "
'▼2007/09/09 fix tajima ↓strSQL = strSQL & "    AND YOUKT_MOVEKBN = " & P_MOVEKBN_無し
    strSQL = strSQL & "    AND ISNULL(YOUKT_MOVEKBN," & P_MOVEKBN_無し & ") = " & P_MOVEKBN_無し
'▲2007/09/09 fix tajima
strSQL = strSQL & "  GROUP BY YOUKT_YCODE "

    fncMakeUkeSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeKeiSql
'        機能             :契約数数検索SQL作成
'        IN               :第一引数－ﾔｰﾄﾞｺｰﾄﾞの条件
'                         :第二引数－日付の条件
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeKeiSql(strCondYard As String, strCondDate As String) As String

    Dim strSQL              As String
    
    strSQL = strSQL & " SELECT CARG_YCODE, "
    
    '自動紹介（電話）契約数
    strSQL = strSQL & "        SUM( CASE WHEN YOUKT_GENKBN = " & P_GENKBN_電話
    strSQL = strSQL & "                   AND YOUKT_ORG_YUKBN = " & P_YUKBN_自動紹介
    strSQL = strSQL & "                  THEN 1 ELSE 0 "
    strSQL = strSQL & "             END ) AS AUTO_TEL_KEI, "
    
    '自動紹介（インターネット）契約数
    strSQL = strSQL & "        SUM( CASE WHEN YOUKT_GENKBN = " & P_GENKBN_インターネット
    strSQL = strSQL & "                   AND YOUKT_ORG_YUKBN = " & P_YUKBN_受付 '2007/09/09 fix tajima old:P_YUKBN_自動紹介
    strSQL = strSQL & "                  THEN 1 ELSE 0 "
    strSQL = strSQL & "             END ) AS AUTO_INET_KEI, "
    
    '手動紹介（電話）契約数
    strSQL = strSQL & "        SUM( CASE WHEN YOUKT_GENKBN = " & P_GENKBN_電話
    strSQL = strSQL & "                   AND YOUKT_ORG_YUKBN != " & P_YUKBN_自動紹介
    strSQL = strSQL & "                  THEN 1 ELSE 0 "
    strSQL = strSQL & "             END ) AS MANUAL_TEL_KEI "
    
    strSQL = strSQL & "   FROM CARG_FILE "
    strSQL = strSQL & "        INNER JOIN YOUK_TRAN "
    strSQL = strSQL & "           ON YOUKT_UKNO = CARG_UKNO "
    
    strSQL = strSQL & "  WHERE CARG_YCODE " & strCondYard
    strSQL = strSQL & "    AND CARG_STDATE " & strCondDate
    strSQL = strSQL & "    AND YOUKT_GENKBN IN(" & P_GENKBN_電話 & "," & P_GENKBN_インターネット & ") "
    strSQL = strSQL & "    AND YOUKT_ORG_YUKBN IN(" & P_YUKBN_自動紹介 & "," _
                                                    & P_YUKBN_自社理由 & "," _
                                                    & P_YUKBN_取置 & "," _
                                                    & P_YUKBN_受付 & ") "
'▼2007/09/09 fix tajima ↓strSQL = strSQL & "    AND YOUKT_MOVEKBN = " & P_MOVEKBN_無し
    strSQL = strSQL & "    AND ISNULL(YOUKT_MOVEKBN," & P_MOVEKBN_無し & ") = " & P_MOVEKBN_無し
'▲2007/09/09 fix tajima
    
    strSQL = strSQL & "   GROUP BY CARG_YCODE "

    fncMakeKeiSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeKaiyakuSql
'        機能             :解約数検索SQL作成
'        IN               :第一引数－ﾔｰﾄﾞｺｰﾄﾞの条件
'                         :第二引数－日付の条件
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeKaiyakuSql(strCondYard As String, strCondDate As String) As String

    Dim strSQL              As String
    
    strSQL = strSQL & " SELECT CARG_YCODE, "
    strSQL = strSQL & "        COUNT(*) AS KAIYAKU "
    strSQL = strSQL & "   FROM CARG_FILE "
    strSQL = strSQL & "  WHERE CARG_YCODE " & strCondYard
    strSQL = strSQL & "    AND CARG_KYDATE " & strCondDate
    strSQL = strSQL & "    AND NOT EXISTS "
    strSQL = strSQL & "        ( SELECT * FROM YOUK_TRAN "
    strSQL = strSQL & "           WHERE YOUKT_MOTO_ACPTNO = CARG_ACPTNO ) "
    strSQL = strSQL & "  GROUP BY CARG_YCODE "

    fncMakeKaiyakuSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeBetween
'        機能             :範囲条件作成
'        IN               :第一引数－範囲条件値FROM
'                         :第二引数－範囲条件値TO
'        OUT              :条件文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeBetween(strfrom As String, strTo As String) As String

    Dim strTemp     As String
    
    strTemp = ""
    
    If strfrom <> "" And strTo <> "" Then
        '共に空白ではない場合、
        If strfrom = strTo Then
            'FROMとTOが同一の場合、一致条件
            strTemp = " = '" & strfrom & "' "
        Else
            'FROMとTOが異なる場合、BETWEEN条件
            strTemp = " BETWEEN '" & strfrom & "' AND '" & strTo & "' "
        End If
    ElseIf strfrom <> "" Then
        'FROMのみの場合、それ以上であることが条件
        strTemp = " >= '" & strfrom & "' "
    ElseIf strTo <> "" Then
        'TOのみの場合、それ以下であることが条件
        strTemp = " <= '" & strTo & "' "
    End If
    
    fncMakeBetween = strTemp

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

    Dim strFromTemp     As String
    Dim strToTemp       As String
    
    If Nz(txtDateFrom, "") = "" Then
        strFromTemp = "1900/01/01"
    Else
        strFromTemp = txtDateFrom
    End If
    
    If Nz(txtDateTo, "") = "" Then
        strToTemp = "9999/12/31"
    Else
        strToTemp = txtDateTo
    End If

    With rsDestination
        .AddNew

        .Fields("対象日FROM") = Format(strFromTemp, "yyyy/mm/dd")
        .Fields("対象日TO") = Format(strToTemp, "yyyy/mm/dd")
        .Fields("ヤードコード") = Format(rsSource.Fields("YARD_CODE"), "000000")
        .Fields("ヤード名") = rsSource.Fields("YARD_NAME")
                
        .Fields("自動紹介(電話)予約数") = Nz(rsSource.Fields("AUTO_TEL_REQ"), 0)
        .Fields("自動紹介(電話)受付数") = Nz(rsSource.Fields("AUTO_TEL_UKE"), 0)
        .Fields("自動紹介(電話)契約数") = Nz(rsSource.Fields("AUTO_TEL_KEI"), 0)
        
        .Fields("自動紹介(ｲﾝﾀｰﾈｯﾄ)予約数") = Nz(rsSource.Fields("AUTO_INET_REQ"), 0)
        .Fields("自動紹介(ｲﾝﾀｰﾈｯﾄ)受付数") = Nz(rsSource.Fields("AUTO_INET_UKE"), 0)
        .Fields("自動紹介(ｲﾝﾀｰﾈｯﾄ)契約数") = Nz(rsSource.Fields("AUTO_INET_KEI"), 0)
                
        .Fields("手動紹介(電話)受付数") = Nz(rsSource.Fields("MANUAL_TEL_UKE"), 0)
        .Fields("手動紹介(電話)契約数") = Nz(rsSource.Fields("MANUAL_TEL_KEI"), 0)
                
        .Fields("総受付数") = .Fields("自動紹介(電話)受付数") _
                            + .Fields("自動紹介(ｲﾝﾀｰﾈｯﾄ)受付数") _
                            + .Fields("手動紹介(電話)受付数")
        
        .Fields("総契約数") = .Fields("自動紹介(電話)契約数") _
                            + .Fields("自動紹介(ｲﾝﾀｰﾈｯﾄ)契約数") _
                            + .Fields("手動紹介(電話)契約数")
        
        .Fields("解約数") = Nz(rsSource.Fields("KAIYAKU"), 0)
        
        .Fields("差引本数") = .Fields("総契約数") - .Fields("解約数")
        
        .UPDATE
    End With
    
End Sub
    
'==============================================================================*
'
'        MODULE_NAME      :txtDateFrom_AfterUpdate
'        MODULE_ID        :txtDateFrom_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtDateFrom_AfterUpdate()
    
    txtDateFrom = Format(VarChangeDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtDateFrom_BeforeUpdate
'        MODULE_ID        :txtDateFrom_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtDateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call fncDateCheck(txtDateFrom, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtDateFrom_GotFocus
'        MODULE_ID        :txtDateFrom_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtDateFrom_GotFocus()

    txtDateFrom = Format(txtDateFrom, "yy/mm/dd")
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtDateFrom_KeyPress
'        MODULE_ID        :txtDateFrom_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtDateFrom, KeyAscii, 10)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtDateFrom_LostFocus
'        MODULE_ID        :txtDateFrom_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtDateFrom_LostFocus()

    txtDateFrom = Format(txtDateFrom, "yyyy/mm/dd")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txtDateTo_AfterUpdate
'        MODULE_ID        :txtDateTo_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtDateTo_AfterUpdate()
    
    txtDateTo = Format(VarChangeDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtDateTo_BeforeUpdate
'        MODULE_ID        :txtDateTo_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtDateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call fncDateCheck(txtDateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtDateTo_GotFocus
'        MODULE_ID        :txtDateTo_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtDateTo_GotFocus()

    txtDateTo = Format(txtDateTo, "yy/mm/dd")
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtDateTo_KeyPress
'        MODULE_ID        :txtDateTo_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtDateTo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtDateTo, KeyAscii, 10)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtDateTo_LostFocus
'        MODULE_ID        :txtDateTo_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtDateTo_LostFocus()

    txtDateTo = Format(txtDateTo, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :DateCheck
'        機能             :日付入力チェック
'        IN               :対象テキスト
'                          更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDateCheck(txtTarget As Variant, ByRef Cancel As Integer) As Boolean

    fncDateCheck = True
    
    VarChangeDate = CmDateChecker(txtTarget)
    If IsNull(VarChangeDate) Then
        MsgBox "日付を正しく入力してください。", vbInformation, Me.NAME
        Cancel = True
        SendKeys "{ESC}"
        fncDateCheck = False
    End If

End Function

'****************************  ended or program ********************************

