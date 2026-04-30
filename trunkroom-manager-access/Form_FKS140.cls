VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　倉庫管理システム
'**********************************************
'   プログラム名　　：　契約一覧表
'   プログラムＩＤ　：
'   作　成　日　　　：
'   作　成　者　　　：　αシステム  Ishii Kaoru
'   Ver             ： 0.0
'**********************************************
'修正履歴
'   修　正　日　　　：2004/11/24
'   修　正　者　　　：N.MIURA
'   修　正　内　容　：フォーム名変更
'   Ver             ： 0.1
'**********************************************
'   修　正　日　　　：2007/06/25
'   修　正　者　　　：tajima
'   修　正　内　容　：契約リスト出力機能追加
'   Ver             ： 0.2
'**********************************************
'   修　正　日　　　：2007/07/11
'   修　正　者　　　：tajima
'   修　正　内　容　：顧客マスタから遷移時に印刷ボタンへセットフォーカス
'   Ver             ： 0.2a
'**********************************************
'   修　正　日　　　：2007/07/30
'   修　正　者　　　：tajima
'   修　正　内　容　：ヤード廃止移動届書機能追加
'   Ver             ： 0.3
'**********************************************
'   修　正　日　　　：2009/12/08
'   修　正　者　　　：M.RYU
'   修　正　内　容　：①レポートRKS230⇒ページ設定調整
'                   ：②モジュールPrint230⇒
'                       ワークテーブルを作成するとき、保証会社名称をDBから取得
'                   ：③レポートRKS481⇒保証会社名列を追加
'                   ：④ワークTBLのRKS481_W01⇒保証会社ｺｰﾄﾞや名称項目を追加
'                   ：⑤モジュールPrint481⇒
'                       ワークTBLを作成するとき、保証会社コードや名称を追加
'   Ver             ： 0.4
'**********************************************
'   修　正　日　　　：2010/02/03
'   修　正　者　　　：M.Ryu
'   修　正　内　容　：一時使用契約変更承諾書（自社移動理由）出力機能追加
'                     ①ﾌｫｰﾑ　　⇒　ボタンを追加
'                     ①ﾚﾎﾟｰﾄ 　⇒　RSK231作成
'                     ①ﾓｼﾞｭｰﾙ　⇒　Print231作成
'   Ver             ： 0.5
'**********************************************
'   修　正　日　　　：2011/03/02
'   修　正　者　　　：M.Ryu
'   修　正　内　容　：検索条件に保証会社ｺｰﾄﾞを追加
'   Ver             ： 0.7
'**********************************************
'   修　正　日　　　：2011/03/02
'   修　正　者　　　：M.Ryu
'   修　正　内　容　：検索条件に保証会社ｺｰﾄﾞを追加
'   Ver             ： 0.7
'**********************************************
'   修　正　日　　　：2012/04/12
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：書類送付先を追加
'   Ver             ： 0.8
'**********************************************
'   修　正　日　　　：2013/02/20
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：補償加入費一覧を追加
'   Ver             ： 0.9
'**********************************************
'   修　正　日　　　：2013/05/01
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：委託料加瀬負担一覧を追加
'   Ver             ： 1.0
'**********************************************
'   修　正　日　　　：2013/06/03
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：終了メッセージを表示させないように修正
'   Ver             ： 1.1
'**********************************************
'   修　正　日　　　：2015/01/08
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：配送サービス一覧を追加
'   Ver             ： 1.2
'**********************************************
'   修　正　日　　　：2019/09/07
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：予約リストを追加
'   Ver             ： 1.3
'**********************************************
'   修　正　日　　　：2035/12/16
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：契約リストに担当者カナを追加
'   Ver             ： 1.4
'**********************************************

Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FKS140"

Const strMsg_kaku = "確認"
Const strMsg_err = "エラー"

Dim strMsg              As String                           'メッセージ表示用
Dim strSQL              As String                           'ＳＱＬ文格納用
Dim intAns              As Integer                          '確認入力用
Private m_strTimeStp    As String                           '実行日時
Private m_strZTimeStp   As String                           '１日前の実行日時
'▼ '2007/06/25 add tajima
Private var変更日付 As Variant
Private Const P_印刷種別_契約一覧表 As Integer = 1
Private Const P_印刷種別_契約リスト As Integer = 2
'▼ '2007/07/30 add tajima
Private Const P_印刷種別_廃止移動届 As Integer = 3

Private Const P_印刷種別_一時使用契約変更承諾書 As Integer = 4  '----20100203--M.Ryu--add--
Private Const P_印刷種別_補償加入費一覧 As Integer = 5  '----20130220--M.HONDA--add--
Private Const P_印刷種別_委託料加瀬負担一覧 As Integer = 6  '----20130501--M.HONDA--add--
Private Const P_印刷種別_配送サービス一覧 As Integer = 7  '----20150108--M.HONDA--add--
Private Const P_印刷種別_予約リスト As Integer = 8  '2019/09/07 M.HONDA
Private Const C_TEMP_TABLE = "予約リスト"


'**********************************************

'画面入力より
'SQL の WHERE文作成
'**********************************************
Private Function Make_Where() As String

    Dim strCodefr   As String
    Dim strCODETO   As String
    Dim strUserfr   As String
    Dim strUserto   As String
    Dim strDayfr    As String
    Dim strDayto    As String
    Dim strSwap     As String
    Dim strWhere    As String

    On Error GoTo Err_Make_Where

    '*** 画面指定値取得 ***'
    strCodefr = Nz(Me.cmb_CODEFROM, "")
    strCODETO = Nz(Me.cmb_CODETO, "")
    strUserfr = Nz(Me.cmb_顧客FROM, "")
    strUserto = Nz(Me.cmb_顧客TO, "")
    strDayfr = Nz(Me.txt_日付FROM, "")
    strDayto = Nz(Me.txt_日付TO, "")
    strSwap = ""

    '*** ＩＤの大小チェック ***
    If Trim(strCodefr) <> "" And Trim(strCODETO) <> "" Then
        '*** ＣＯＤＥ入替 ***
        If strCodefr > strCODETO Then
            strSwap = strCodefr
            strCodefr = strCODETO
            strCODETO = strSwap
        End If
    End If

    strSwap = ""
    If Trim(strUserfr) <> "" And Trim(strUserto) <> "" Then
        '*** ＣＯＤＥ入替 ***
        If strUserfr > strUserto Then
            strSwap = strUserfr
            strUserfr = strUserto
            strUserto = strSwap
        End If
    End If

    '更新日のチェック
    If Me.txt_日付FROM <> "" And Me.txt_日付TO <> "" Then
        strDayfr = Format(Nz(Me.txt_日付FROM, ""), "YYYY/MM/DD")
        strDayto = Format(Nz(Me.txt_日付TO, ""), "YYYY/MM/DD")
        If strDayto <> "" And strDayfr > strDayto Then
            'strSwap = strDayto                      'Del 01-11-2000 Araki
            'Me.txt_日付ＦＲＯＭ = strSwap            'Del 01-11-2000 Araki
            'txt_日付ＴＯ = strDayfr                  'Del 01-11-2000 Araki
            
            '$$$ start $$$ 契約日選択範囲の大小逆転時のデータ抽出ミス修正のため追加 01-11-2000 Araki
            strSwap = strDayfr
            strDayfr = strDayto
            strDayto = strSwap
            Me.txt_日付FROM = strDayfr
            Me.txt_日付TO = strDayto
            '$$$ end $$$
        End If
    End If

    strWhere = ""
    '契約状態 2=契約中
    '解約データも対象とする。1999/12/14 UPDATE Ishii
'    strWhere = strWhere & " CARG_AGRE =2 "
'    If strCodefr <> "" Or strCodeto <> "" Then
'        strWhere = strWhere & " AND "
'    End If

    'ヤードコード
    'FROM →
    If Trim(strCodefr) <> "" And Trim(strCODETO) = "" Then
        strWhere = strWhere & " CARG_YCODE >= " & Trim(strCodefr) & ""
    '← TO
    ElseIf Trim(strCodefr) = "" And Trim(strCODETO) <> "" Then
        strWhere = strWhere & "CARG_YCODE <= " & strCODETO & ""
    'FROM ⇔ TO
    ElseIf Trim(strCodefr) <> "" And Trim(strCODETO) <> "" Then
        If Trim(strCodefr) <> Trim(strCODETO) Then
            strWhere = strWhere & " CARG_YCODE BETWEEN " & strCodefr & " "
            strWhere = strWhere & " AND " & strCODETO & " "
        Else
            strWhere = strWhere & " CARG_YCODE = " & Trim(strCodefr) & ""
        End If
    '指定なし（全件）
    ElseIf Trim(strCodefr) = "" And Trim(strCODETO) = "" Then

    End If

    If strWhere <> "" Then
        If strUserfr <> "" Or strUserto <> "" Then
            strWhere = strWhere & " AND "
        End If
    End If

    '顧客コード
    '← TO
    If Trim(strUserfr) = "" And Trim(strUserto) <> "" Then
        strWhere = strWhere & " CARG_UCODE <= " & strUserto & ""
    'FROM ⇔ TO
    ElseIf Trim(strUserfr) <> "" And Trim(strUserto) <> "" Then
        If Trim(strUserfr) <> Trim(strUserto) Then
            strWhere = strWhere & " CARG_UCODE BETWEEN " & strUserfr & " "
            strWhere = strWhere & " AND " & strUserto & " "
        Else
            strWhere = strWhere & " CARG_UCODE = " & Trim(strUserfr) & ""
        End If
    'FROM →
    ElseIf Trim(strUserfr) <> "" And Trim(strUserto) = "" Then
        strWhere = strWhere & "  CARG_UCODE >= " & Trim(strUserfr) & ""
    '指定なし（全件）
    ElseIf Trim(strUserfr) = "" And Trim(strUserto) = "" Then

    End If

    If strWhere <> "" Then
        If strDayfr <> "" Or strDayto <> "" Then
            strWhere = strWhere & " AND "
        End If
    End If

    '初回契約日
    If Trim(strDayfr) = "" And Trim(strDayto) <> "" Then
        strWhere = strWhere & " CARG_FSDATE <= CONVERT(datetime," & Chr$(34) & strDayto & Chr$(34) & ")"
    'FROM ⇔ TO
    ElseIf Trim(strDayfr) <> "" And Trim(strDayto) <> "" Then
        If Trim(strDayfr) <> Trim(strDayto) Then
            strWhere = strWhere & " CARG_FSDATE BETWEEN CONVERT(datetime," & Chr$(34) & strDayfr & Chr$(34) & ")"
            strWhere = strWhere & " AND CONVERT(datetime," & Chr$(34) & strDayto & Chr$(34) & ")"
        Else
            strWhere = strWhere & " CARG_FSDATE = CONVERT(datetime," & Chr$(34) & strDayto & Chr$(34) & ")"
        End If
    'FROM →
    ElseIf Trim(strDayfr) <> "" And Trim(strDayto) = "" Then
        strWhere = strWhere & "  CARG_FSDATE >= CONVERT(datetime," & Chr$(34) & strDayfr & Chr$(34) & ")"
    '指定なし（全件）
    ElseIf Trim(strDayfr) = "" And Trim(strDayto) = "" Then

    End If

    Make_Where = strWhere

Exit_Make_Where:
    Exit Function

Err_Make_Where:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Make_Where

End Function

'**********************************************
'SQLServer側にワークテーブル作成
'戻り値 ： strWhere,INPUT,String ,SQL条件文
'**********************************************
Private Function Make_Work() As Boolean

    Dim objDb       As Database
    Dim ObjQDF      As QueryDef
    Dim objRst      As Recordset
    Dim strConnect  As String               'Connect情報
    Dim strError    As String               'エラーメッセージ用
    Dim errObj      As Error
    Dim strWhere    As String

    On Error GoTo Err_Make_Work

    Make_Work = False

    'ワークの削除
    Call DEL_RTN

    strWhere = Make_Where

    Set objDb = CurrentDb

    strConnect = DLookup("CONT_SERVER_PATH", "dbo_CONT_MAST", "CONT_KEY=1")
    If IsNull(strConnect) Then
        MsgBox "ＳＱＬＳｅｒｖｅｒのＰＡＳＳが見つかりません。", vbInformation + vbOKOnly, "エラー"
        GoTo Exit_Make_Work
    End If

    ' 規定の QueryTimeout プロパティを変更します。
    Debug.Print "Defauit QueryTimeout of Database: " & objDb.QueryTimeout
    objDb.QueryTimeout = 1200                             '２０分（０：タイムアウト無し）
    Debug.Print "New     QueryTimeout of Database: " & objDb.QueryTimeout

    Set ObjQDF = objDb.CreateQueryDef("")
    ObjQDF.Connect = strConnect

    ObjQDF.ReturnsRecords = True

    ' 規定の ODBCTimeout プロパティを変更します。
    Debug.Print "Defauit ODBCTimeout of Database: " & ObjQDF.ODBCTimeout
    ObjQDF.ODBCTimeout = 1200                             '２０分（０：タイムアウト無し）
    Debug.Print "New     ODBCTimeout of Database: " & ObjQDF.ODBCTimeout

    strSQL = ""
    strSQL = "DECLARE @Day_Time DateTime "                            '実行日時

    '文字列の中にｼﾝｸﾞﾙｸｫｰﾃｰｼｮﾝが含まれる場合、エラーになってしまうので、
    'ﾀﾞﾌﾞﾙｸｫｰﾃｰｼｮﾝで囲み、文字として扱うための設定
    '必ず一番最後で、設定をONに戻すこと！
    '戻さないと、ﾃｰﾌﾞﾙを開いたとき、ODBC呼び出しは失敗しましたがでて、見れません！！！
    strSQL = strSQL & "SET QUOTED_IDENTIFIER OFF "
    strSQL = strSQL & "SELECT @Day_Time =convert(DateTime,'" & Trim(m_strTimeStp) & "') "

    'dbo_RKS140_WORKへ追加
    strSQL = strSQL & "INSERT INTO RKS140_WORK"
    strSQL = strSQL & "( RKS140_DATETIME,"
    strSQL = strSQL & "CARG_YCODE,"
    strSQL = strSQL & "CARG_NO,"
    strSQL = strSQL & "CARG_UCODE,"
    strSQL = strSQL & "CARG_AGRE,"
    strSQL = strSQL & "CARG_FSDATE,"
    strSQL = strSQL & "CARG_STDATE,"
    strSQL = strSQL & "CARG_EDDATE,"
    strSQL = strSQL & "CARG_CYDATE,"
    strSQL = strSQL & "CARG_KYDATE,"
    strSQL = strSQL & "CARG_CONTNA,"
    strSQL = strSQL & "CARG_DOCU1,"
    strSQL = strSQL & "CARG_DOCU2,"
    strSQL = strSQL & "CARG_TUBOSU,"
    strSQL = strSQL & "CARG_TTANKA,"
    strSQL = strSQL & "CARG_RENTKG,"
    strSQL = strSQL & "CARG_SYOZEI,"
    strSQL = strSQL & "CARG_SECUKG,"
    strSQL = strSQL & "CARG_BIKO,"
    strSQL = strSQL & "CARG_UPDATE)"
    strSQL = strSQL & " SELECT @Day_Time,"
    strSQL = strSQL & " CARG_YCODE,"
    strSQL = strSQL & " CARG_NO,"
    strSQL = strSQL & " CARG_UCODE,"
    strSQL = strSQL & " CARG_AGRE,"
    strSQL = strSQL & " CARG_FSDATE,"
    strSQL = strSQL & " CARG_STDATE,"
    strSQL = strSQL & " CARG_EDDATE,"
    strSQL = strSQL & " CARG_CYDATE,"
    strSQL = strSQL & " CARG_KYDATE,"
    strSQL = strSQL & " CARG_CONTNA,"
    strSQL = strSQL & " CARG_DOCU1,"
    strSQL = strSQL & " CARG_DOCU2,"
    strSQL = strSQL & " CARG_TUBOSU,"
    strSQL = strSQL & " CARG_TTANKA,"
    strSQL = strSQL & " CARG_RENTKG,"
    strSQL = strSQL & " CARG_SYOZEI,"
    strSQL = strSQL & " CARG_SECUKG,"
    strSQL = strSQL & " CARG_BIKO,"
    strSQL = strSQL & " CARG_UPDATE"
    strSQL = strSQL & " FROM CARG_FILE"

    If strWhere <> "" Then
        strSQL = strSQL & " WHERE " & strWhere
    End If

    '必ず必要！！！
    strSQL = strSQL & " SET QUOTED_IDENTIFIER ON "

    ObjQDF.sql = strSQL
    Set objRst = ObjQDF.OpenRecordset(dbOpenSnapshot)

    If objRst(0) <> 0 Then
        MsgBox "ワーク生成中にエラーが発生しました。", vbCritical + vbOKOnly, "エラー"
        GoTo Exit_Make_Work
    End If

    objDb.Close
    Set objDb = Nothing
    
    Make_Work = True

Exit_Make_Work:
    Exit Function

Err_Make_Work:
    Errors.Refresh
    strError = "ワーク生成中にエラーが発生しました。" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    For Each errObj In DBEngine.Errors
        If Format$(errObj.Number) <> 3325 Then
            strError = strError & Format$(errObj.Number)
            strError = strError & " : " & errObj.Description
            strError = strError & " (" & errObj.Source & ") . "
            strError = strError & Chr$(13) & Chr$(10)
        Else
            'レコードが返されなかったのエラーだから、正常
            Make_Work = True
            GoTo Exit_Make_Work
        End If
    Next
    MsgBox strError, vbCritical + vbOKOnly, strMsg_err
    GoTo Exit_Make_Work

End Function

'*********************************
'ワーク削除
'*********************************
Private Sub DEL_RTN()
    
    Dim objDb   As Database
    
    On Error Resume Next

    Set objDb = CurrentDb()

    '*** SQLServer ***
    strSQL = ""
    strSQL = "DELETE FROM dbo_RKS140_WORK"
    strSQL = strSQL & " WHERE RKS140_DATETIME = CDate(#" & m_strTimeStp & "#)"
    strSQL = strSQL & " OR RKS140_DATETIME <= CDate(#" & m_strZTimeStp & "#)"
    objDb.Execute strSQL

    '*** JET ***
    strSQL = ""
    strSQL = "DELETE FROM コンテナ契約ＷＫ"
    objDb.Execute strSQL, dbFailOnError

    objDb.Close
    Set objDb = Nothing

End Sub

'**********************************************
'SQLSERVERのテーブルよりJETのワークテーブルに取り込み
'戻り値 ： strWhere,INPUT,String ,SQL条件文
'**********************************************
Private Function Make_JetWork() As Boolean

    Dim lngKensu As Long
    Dim strSQL As String
    Dim objDb As Database
    Dim strMsg As String

    On Error GoTo Err_Make_JetWork

    Make_JetWork = False
    
    Set objDb = CurrentDb

    strSQL = ""
    strSQL = strSQL & " INSERT INTO コンテナ契約ＷＫ("
    strSQL = strSQL & "RKS140_DATETIME,"
    strSQL = strSQL & "CARG_YCODE,"
    strSQL = strSQL & "CARG_NO,"
    strSQL = strSQL & "CARG_UCODE,"
    strSQL = strSQL & "CARG_AGRE,"
    strSQL = strSQL & "CARG_FSDATE,"
    strSQL = strSQL & "CARG_STDATE,"
    strSQL = strSQL & "CARG_EDDATE,"
    strSQL = strSQL & "CARG_CYDATE,"
    strSQL = strSQL & "CARG_KYDATE,"
    strSQL = strSQL & "CARG_CONTNA,"
    strSQL = strSQL & "CARG_DOCU1,"
    strSQL = strSQL & "CARG_DOCU2,"
    strSQL = strSQL & "CARG_TUBOSU,"
    strSQL = strSQL & "CARG_TTANKA,"
    strSQL = strSQL & "CARG_RENTKG,"
    strSQL = strSQL & "CARG_SYOZEI,"
    strSQL = strSQL & "CARG_SECUKG,"
    strSQL = strSQL & "CARG_BIKO,"
    strSQL = strSQL & "CARG_UPDATE)"
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " RKS140_DATETIME,"
    strSQL = strSQL & " CARG_YCODE,"
    strSQL = strSQL & " CARG_NO,"
    strSQL = strSQL & " CARG_UCODE,"
    strSQL = strSQL & " CARG_AGRE,"
    strSQL = strSQL & " CARG_FSDATE,"
    strSQL = strSQL & " CARG_STDATE,"
    strSQL = strSQL & " CARG_EDDATE,"
    strSQL = strSQL & " CARG_CYDATE,"
    strSQL = strSQL & " CARG_KYDATE,"
    strSQL = strSQL & " CARG_CONTNA,"
    strSQL = strSQL & " CARG_DOCU1,"
    strSQL = strSQL & " CARG_DOCU2,"
    strSQL = strSQL & " CARG_TUBOSU,"
    strSQL = strSQL & " CARG_TTANKA,"
    strSQL = strSQL & " CARG_RENTKG,"
    strSQL = strSQL & " CARG_SYOZEI,"
    strSQL = strSQL & " CARG_SECUKG,"
    strSQL = strSQL & " CARG_BIKO,"
    strSQL = strSQL & " CARG_UPDATE"
    strSQL = strSQL & " FROM dbo_RKS140_WORK"
    strSQL = strSQL & " WHERE RKS140_DATETIME = CDate(#" & m_strTimeStp & "#)"

    objDb.Execute strSQL, dbFailOnError

    '*** データチェック ***
    'アクションクエリーの結果で判断
    lngKensu = objDb.RecordsAffected
    If lngKensu < 1 Then
        strMsg = "対象となるデータが存在しませんでした。"
        MsgBox strMsg, vbInformation + vbOKOnly, strMsg_kaku
        GoTo Exit_Make_JetWork
    End If

    objDb.Close
    Set objDb = Nothing

    Make_JetWork = True

Exit_Make_JetWork:
    Exit Function

Err_Make_JetWork:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Make_JetWork

End Function

'**********************************************
'cmb_CODEFROM_AfterUpdateイベント
'**********************************************
Private Sub cmb_CODEFROM_AfterUpdate()
On Error GoTo Err_cmb_CODEFROM_AfterUpdate
    
    Dim varCode As Variant
    Dim strCARG As String
    Dim strWhere As String

    varCode = Me.cmb_CODEFROM
    strCARG = ""

    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " dbo_YARD_MAST.YARD_CODE=" & varCode & ""
        strCARG = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strWhere), "")
    End If
    Me.cmb_CODEFROM = Format(Me.cmb_CODEFROM, "000000")
    Me.txt_CODEFROM = strCARG
    Me.cmb_CODETO.SetFocus

Exit_cmb_CODEFROM_AfterUpdate:
    Exit Sub

Err_cmb_CODEFROM_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_CODEFROM_AfterUpdate

End Sub

'**********************************************
'cmb_CODEFROM_KeyPressイベント
'**********************************************
Private Sub cmb_CODEFROM_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_CODEFROM, KeyAscii, 6)
End Sub

'**********************************************
'cmb_CODETO_AfterUpdateイベント
'**********************************************
Private Sub cmb_CODETO_AfterUpdate()
On Error GoTo Err_cmb_CODETO_AfterUpdate
    
    Dim varCode As Variant
    Dim varChCode As Variant
    Dim strCARG As String
    Dim strWhere As String
    
    varCode = Me.cmb_CODETO
    strCARG = ""
    
    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " YARD_CODE=" & varCode & ""
        strCARG = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strWhere), "")
    End If
    Me.cmb_CODETO = Format(Me.cmb_CODETO, "000000")
    Me.txt_CODETO = strCARG
    Me.cmb_顧客FROM.SetFocus

Exit_cmb_CODETO_AfterUpdate:
    Exit Sub

Err_cmb_CODETO_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_CODETO_AfterUpdate

End Sub

'**********************************************
'cmb_顧客FROM_KeyPressイベント
'**********************************************
Private Sub cmb_顧客FROM_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_顧客FROM, KeyAscii, 6)
End Sub

'**********************************************
'cmb_顧客TO_AfterUpdateイベント
'**********************************************
Private Sub cmb_顧客TO_AfterUpdate()
On Error GoTo Err_cmb_CODETO_AfterUpdate
    
    Dim varCode As Variant
    Dim varChCode As Variant
    Dim strCARG As String
    Dim strWhere As String
    
    varCode = Me.cmb_顧客TO
    strCARG = ""

    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " USER_CODE=" & varCode & ""
        strCARG = Nz(DLookup("USER_NAME", "dbo_USER_MAST", strWhere), "")
    End If
    Me.cmb_顧客TO = Format(Me.cmb_顧客TO, "000000")
    Me.txt_顧客TO = strCARG

Exit_cmb_CODETO_AfterUpdate:
    Exit Sub

Err_cmb_CODETO_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_CODETO_AfterUpdate

End Sub

'**********************************************
'cmb_顧客FROM_AfterUpdateイベント
'**********************************************
Private Sub cmb_顧客FROM_AfterUpdate()
On Error GoTo Err_cmb_CODETO_AfterUpdate

    Dim varCode As Variant
    Dim varChCode As Variant
    Dim strCARG As String
    Dim strWhere As String

    varCode = Me.cmb_顧客FROM
    strCARG = ""

    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " USER_CODE=" & varCode & ""
        strCARG = Nz(DLookup("USER_NAME", "dbo_USER_MAST", strWhere), "")
    End If
    Me.cmb_顧客FROM = Format(Me.cmb_顧客FROM, "000000")
    Me.Txt_顧客FROM = strCARG
    If Me.cmb_顧客TO.Visible = True Then Me.cmb_顧客TO.SetFocus '2007/07/30 tajima fix

Exit_cmb_CODETO_AfterUpdate:
    Exit Sub

Err_cmb_CODETO_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_CODETO_AfterUpdate

End Sub

'**********************************************
'cmb_CODETO_KeyPressイベント
'**********************************************
Private Sub cmb_CODETO_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_CODETO, KeyAscii, 6)
End Sub

'**********************************************
'cmb_顧客TO_KeyPressイベント
'**********************************************
Private Sub cmb_顧客TO_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_顧客TO, KeyAscii, 6)
End Sub


'**********************************************
'cmd_プレビュー_Clickイベント
'**********************************************
Private Sub cmd_プレビュー_Click()
    
    Dim strMsg As String
    Dim intAns As Integer
    
    On Error GoTo Err_cmd_プレビュー_Click

    '▼ 2007/06/25 add tajima
    If frm_印刷種別.VALUE = P_印刷種別_契約リスト Then
        Call sub契約リスト出力(Print481.P481_MODE_PREVIEW)
        Exit Sub
    End If
    '▲ 2007/06/25 add tajima

    '▼ 2007/07/30 add tajima
    If frm_印刷種別.VALUE = P_印刷種別_廃止移動届 Then
        Call sub廃止移動届出力(Print230.P230_MODE_PREVIEW)
        Exit Sub
    End If
    '▲ 2007/07/30 add tajima
    
    '----20100203----M.Ryu----add-------<s>
    If frm_印刷種別.VALUE = P_印刷種別_一時使用契約変更承諾書 Then
        Call sub一時使用契約変更承諾書(Print231.P231_MODE_PREVIEW)
        Exit Sub
    End If
    '----20100203----M.Ryu----add-------<e>
    
    '----20130501----M.HONDA----add-------<s>
    If frm_印刷種別.VALUE = P_印刷種別_委託料加瀬負担一覧 Then
        Call sub委託料加瀬負担一覧(Print231.P231_MODE_PREVIEW)
        Exit Sub
    End If
    '----20130501----M.HONDA----add-------<e>
    
    
    
    strMsg = "印刷内容を画面に出力します。" '& vbCr
    strMsg = strMsg & "よろしいですか？"
    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)

    If intAns = vbNo Then
        GoTo Exit_cmd_プレビュー_Click
    End If
    
    
    '*** ワーク作成 ***
    doCmd.Hourglass True
    Call DEL_RTN
    If Make_Work = False Then
        '*** エラー発生またはデータなし ***
        GoTo Exit_cmd_プレビュー_Click
    End If

    '*** Jetワーク作成 ***
    If Make_JetWork = False Then
        'エラー発生またはデータなし
        GoTo Exit_cmd_プレビュー_Click
    End If

    doCmd.OpenReport "RKS140_契約一覧表", acPreview

Exit_cmd_プレビュー_Click:
    doCmd.Hourglass False
    'Me.txt_日付FROM.SetFocus 2007/07/30 del tajima
    Exit Sub

Err_cmd_プレビュー_Click:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmd_プレビュー_Click

End Sub

'**********************************************
'cmd_印刷_Clickイベント
'**********************************************
Private Sub cmd_印刷_Click()
On Error GoTo Err_cmd_印刷_Click
    Dim strMsg As String
    Dim intAns As Integer

    strMsg = "印刷内容をプリンタに印字します。" & vbCr
    strMsg = strMsg & "よろしいですか？"
    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)

    If intAns = vbNo Then
        GoTo Exit_cmd_印刷_Click
    End If

    '*** ワーク作成 ***
    doCmd.Hourglass True
    Call DEL_RTN
    If Make_Work = False Then
        'エラー発生またはデータなし
        GoTo Exit_cmd_印刷_Click
    End If

    '*** Jetワーク作成 ***
    If Make_JetWork = False Then
        'エラー発生またはデータなし
        GoTo Exit_cmd_印刷_Click
    End If

    '印刷実行
    doCmd.OpenReport "RKS140_契約一覧表", acNormal

Exit_cmd_印刷_Click:
    doCmd.Hourglass False
    ' Me.txt_日付FROM.SetFocus 2007/07/30 del tajima
    Exit Sub

Err_cmd_印刷_Click:
    Select Case Err
        Case 2501:
            MsgBox "印刷を中止しました!!", vbInformation + vbOKOnly, strMsg_kaku
        Case 2201:
            MsgBox "プリンタドライバーの設定に誤りがあります。", vbInformation + vbOKOnly, strMsg_kaku
        Case Else:
            MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    End Select

    Resume Exit_cmd_印刷_Click
End Sub

'**********************************************
'cmd_終了_Clickイベント
'**********************************************
Private Sub cmd_終了_Click()
On Error Resume Next
    'doCmd.Close acForm, "FKS140_契約一覧表"   'DELETE 20041124 N.MIURA
    doCmd.Close acForm, "FKS140"               'INSERT 20041124 N.MIURA

End Sub

'**********************************************
'cmd_保証会社_AfterUpdate
'-----2011/03/02 M.RYU INSERT
'**********************************************
Private Sub cmd_保証会社_AfterUpdate()
On Error Resume Next
    Me.txt_保証会社 = Me.cmd_保証会社.Column(1)
End Sub

'**********************************************
'Form_Activateイベント
'**********************************************
Private Sub Form_Activate()
    doCmd.Maximize

End Sub
'**********************************************
'Form_Errorイベント
'**********************************************
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Select Case DataErr
        Case 2113
            Response = acDataErrContinue
            MsgBox "数字を入力してください。", 0 + 64, "確認"
            SendKeys "{ESC}"
            Exit Sub
        Case 2279
            Response = acDataErrContinue
            MsgBox "正しく入力してください。", 0 + 64, "確認"
            SendKeys "{ESC}"
            Exit Sub
        Case 3316
            Response = acDataErrContinue
            SendKeys "{ESC}{ESC}"
            Exit Sub
        Case Else
            MsgBox DataErr
    End Select

End Sub

'**********************************************
'Form_KeyDownイベント
'**********************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Form_KeyDown

    'SHIFT、CTRL、ALTのいずれかが押されている場合は何もしない
    If Shift > 0 Then
        Exit Sub
    End If

    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_終了.SetFocus
        Case vbKeyF2
            KeyCode = 0
            Me.cmd_プレビュー.SetFocus
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

Exit_Form_KeyDown:
    Exit Sub

Err_Form_KeyDown:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Form_KeyDown

End Sub

'**********************************************
'Form_Openイベント
'**********************************************
Private Sub Form_Open(Cancel As Integer)

    Dim wkDate As Date
    '実行日時
    wkDate = Now()
    m_strTimeStp = Format(wkDate, "yyyy/mm/dd hh:nn:ss")
    m_strZTimeStp = Format(DateAdd("d", -1, wkDate), "yyyy/mm/dd hh:nn:ss")
     Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_Title.BackColor = RGB(255, 0, 204)
    End Select
   
    '＊＊＊画面の最大化＊＊＊
    doCmd.Maximize

    '＊＊＊画面表示クリア＊＊＊
    Me.cmb_CODEFROM = ""
    Me.cmb_CODETO = ""
    Me.txt_CODEFROM = ""
    Me.txt_CODETO = ""
    Me.txt_日付FROM = ""
    Me.txt_日付TO = ""
    Me.cmb_顧客FROM = ""
    Me.cmb_顧客TO = ""
    Me.Txt_顧客FROM = ""
    Me.txt_顧客TO = ""
    
    '▼2007/06/25 add tajima
    Dim args() As String
    args = Split(Nz(Me.OpenArgs), ",")
    If UBound(args) >= 3 Then
        Me.cmb_顧客FROM.VALUE = args(0)
        Me.cmb_顧客TO.VALUE = args(0)
        Me.txt_顧客情報変更日From = args(1)
        Me.txt_顧客情報変更日To = args(1)
        Me.cmb_顧客情報変更理由 = args(2)
        Me.cmb_保証区分 = args(3)
        Me.frm_印刷種別.VALUE = P_印刷種別_契約リスト
        ' cmb_に対応する txt_の設定
        ' ※AfterUpdateは呼べないため
        '   フォーカスがないコントロールのプロパティ(多分ListIndex)は参照できんのだってさ
        Me.Txt_顧客FROM.VALUE = Me.cmb_顧客FROM.Column(1, Me.cmb_顧客FROM.ListIndex)
        Me.txt_顧客TO.VALUE = Me.cmb_顧客TO.Column(1, Me.cmb_顧客TO.ListIndex)
        Me.txt_顧客情報変更理由.VALUE = Me.cmb_顧客情報変更理由.Column(1, Me.cmb_顧客情報変更理由.ListIndex)
        Me.txt_保証区分.VALUE = Me.cmb_保証区分.Column(1, Me.cmb_保証区分.ListIndex)
        Call Me.cmd_プレビュー.SetFocus
    Else
        Me.frm_印刷種別.VALUE = P_印刷種別_契約一覧表
    End If
    
    Call frm_印刷種別_AfterUpdate
    
    
End Sub

'**********************************************
'Form_Unloadイベント
'**********************************************
Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Err_Form_Unload

    Dim strMsg As String
    Dim intAns As Integer

    '' 2013/06/03 M.HONDA START
'    strMsg = "終了します。" '& vbCr
'    strMsg = strMsg & "よろしいですか？"
'    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)
'
'    If intAns = vbNo Then
'        Cancel = True
'        ' Me.txt_日付FROM.SetFocus 2007/07/30 del tajima
'        GoTo Exit_Form_Unload
'    Else
'        Call DEL_RTN
'    End If
    '' 2013/06/03 M.HONDA END

Exit_Form_Unload:
    Exit Sub

Err_Form_Unload:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Form_Unload

End Sub

'**********************************************
'txt_日付FROM_AfterUpdateイベント
'**********************************************
Private Sub txt_日付FROM_AfterUpdate()
On Error GoTo Err_txt_日付FROM_AfterUpdate

    '*-Nullチェック-*
    If IsNull(Me.txt_日付FROM) Or Me.txt_日付FROM = "" Then
        Exit Sub
    End If
    
    '*-Check_Date関数を使用-*
    Me.txt_日付FROM = Check_Date(Me.txt_日付FROM)
    
    Exit Sub
    
Err_txt_日付FROM_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub

End Sub

'**********************************************
'txt_日付FROM_BeforeUpdateイベント
'**********************************************
Private Sub txt_日付FROM_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_txt_日付FROM_BeforeUpdate

    Dim vDATE As Variant
    vDATE = Nz(Me.txt_日付FROM, "")

    '*-Nullチェック-*
    If vDATE = "" Then
        Exit Sub
    End If
    
    '*-Check_Date関数を使用-*
    vDATE = Check_Date(vDATE)
    vDATE = CMCheck_Date(vDATE)
    
    '*-入力チェック-*
    Select Case CmBCheck_Date(vDATE)
        Case 0
      
        Case 2
        '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_txt_日付FROM_BeforeUpdate
        Case 3, 9
            MsgBox "日付を正しく入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_txt_日付FROM_BeforeUpdate
    End Select

Exit_txt_日付FROM_BeforeUpdate:
    
    Exit Sub
    
Err_txt_日付FROM_BeforeUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub

End Sub

'**********************************************
'txt_日付FROM_GotFocusイベント
'**********************************************
Private Sub txt_日付FROM_GotFocus()

    '*-YY/MM/DDの形に変換-*
    Me.txt_日付FROM = Format(Me.txt_日付FROM, "YYYY/MM/DD")

End Sub

'**********************************************
'txt_日付FROM_KeyPressイベント
'**********************************************
Private Sub txt_日付FROM_KeyPress(KeyAscii As Integer)

    Call KeyPress_For_Date(Me.txt_日付FROM, KeyAscii, 10)

End Sub

'**********************************************
'txt_日付FROM_LostFocusイベント
'**********************************************
Private Sub txt_日付FROM_LostFocus()

    '*-YY/MM/DDの形に変換-*
    Me.txt_日付FROM = Format(Me.txt_日付FROM, "yyyy/mm/dd")

End Sub

'**********************************************
'txt_日付TO_AfterUpdateイベント
'**********************************************
Private Sub txt_日付TO_AfterUpdate()
On Error GoTo Err_txt_日付TO_AfterUpdate

    '*-Nullチェック-*
    If IsNull(Me.txt_日付TO) Or Me.txt_日付TO = "" Then
        Exit Sub
    End If
    
    '*-Check_Date関数-*
    Me.txt_日付TO = Check_Date(Me.txt_日付TO)

    Exit Sub

Err_txt_日付TO_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub

End Sub

'**********************************************
'txt_日付TO_BeforeUpdateイベント
'**********************************************
Private Sub txt_日付TO_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_txt_日付TO_BeforeUpdate

    Dim vDATE As Variant
    vDATE = Nz(Me.txt_日付TO, "")

    '*-Nullチェック-*
    If vDATE = "" Then
        Exit Sub
    End If

    '*-Check_Date関数をしよう-*
    vDATE = Check_Date(vDATE)
    vDATE = CMCheck_Date(vDATE)

    '*-入力チェック-*
    Select Case CmBCheck_Date(vDATE)
        Case 0
      
        Case 2
        '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_txt_日付TO_BeforeUpdate
        Case 3, 9
            MsgBox "日付を正しく入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_txt_日付TO_BeforeUpdate
    End Select

Exit_txt_日付TO_BeforeUpdate:

    Exit Sub

Err_txt_日付TO_BeforeUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub
End Sub

'**********************************************
'txt_日付TO_GotFocusイベント
'**********************************************
Private Sub txt_日付TO_GotFocus()

    '**--YY/MM/DDの形に変換--**
    Me.txt_日付TO = Format(Me.txt_日付TO, "yyyy/mm/dd")

End Sub

'**********************************************
'txt_日付TO_LostFocusイベント
'**********************************************
Private Sub txt_日付TO_LostFocus()

    '**--YYYY/MM/DDの形に変換--**
    Me.txt_日付TO = Format(Me.txt_日付TO, "yyyy/mm/dd")

End Sub

'**********************************************
'txt_日付TO_KeyPressイベント
'**********************************************
Private Sub txt_日付TO_KeyPress(KeyAscii As Integer)
    
    Call KeyPress_For_Date(Me.txt_日付TO, KeyAscii, 10)

End Sub

'**********************************************
'txt_End_Enterイベント
'**********************************************
Private Sub txt_END_Enter()
    Me.cmd_プレビュー.SetFocus
End Sub

'↓共通関数に入れたい・・・。
'**********************************************
' @f
'
' 機　能　 :キーの取消 日付のみ対応
' 返り値　 :Keyacii
' 引　数　 :ojCtrl   - コントロール
' 　　　    Keyascii - Keyascii
' 　　　　  inLen    - 入力する最大文字長
'
' 機能説明 :引数で指定したコントロールに入力したAsciiCodeを有効か無効か判断し
' 　　　　  返り値をKeyasciiに格納する
'
'**********************************************
Private Sub KeyPress_For_Date(ojCtrl As Control, KeyAscii As Integer, inLen As Integer)

    With ojCtrl
        Select Case KeyAscii
            Case 8, 9, 10, 13                     ' 8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9"), Asc("/")   ' 0～9迄、/ を許可する
                ''' 入力された文字列が引数の最大文字長に達した場合、入力したキーを無効にする
                If Len(.Text) > (inLen - 1) And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With

End Sub
Private Function CMCheck_Date(ByVal vDATE As Variant)
'-----------------------------------------------------------'
'  Check_Dateでは'198/03といった形が変換されるための補助検査
'-----------------------------------------------------------'
  
    Dim sYear As String
      
    CMCheck_Date = Null
      
    If Nz(vDATE, "") = "" Then
        Exit Function
    End If
      
    If IsDate(vDATE) = False Then
        Exit Function
    End If
      
    sYear = Format(vDATE, "yyyy")
    
    If CmLenb(sYear) <> 4 Then
        Exit Function
    End If
      
    CMCheck_Date = vDATE

End Function
Private Function CmBCheck_Date(ByVal oText As Variant, Optional ByVal ilen As Variant) As Integer
'------------------------'
'  入金年月日のチェック
'------------------------'
  '引数
  'oText 対象項目
  'iLen  桁数(省略時は、"yyyy/mm/dd"として10バイト)
    
  '返り値は、
  '２：桁数が溢れている時
  '３：日付と判断できない時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時
On Error GoTo Err_CmBCheck_Date
  
    Dim sYear  As String
    Dim intLen As Integer
    
    CmBCheck_Date = 2
    
    If IsMissing(ilen) = False Then
        intLen = Nz(CmNumeric(ilen), 0)
    Else
        intLen = 10
    End If
    
    oText = Nz(oText, "")
    
    '桁数チェック
    If CmLenb(oText) > intLen Then
        GoTo Exit_CmBCheck_Date
    End If
    
    CmBCheck_Date = 3
    
    If Nz(oText, "") = "" Then
        GoTo Exit_CmBCheck_Date
    End If
    
    '論理チェック
    If IsDate(oText) = False Then
        GoTo Exit_CmBCheck_Date
    End If
    
    sYear = Format(oText, "yyyy")
    
    If CmLenb(sYear) <> 4 Then
        GoTo Exit_CmBCheck_Date
    End If
    
    CmBCheck_Date = 0

Exit_CmBCheck_Date:
  Exit Function
  
Err_CmBCheck_Date:
  MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
  Err.Clear
  CmBCheck_Date = 9
  GoTo Exit_CmBCheck_Date
  
End Function
'★☆★☆★☆★☆★☆★★☆☆ 2007/06/25 add tajima ★☆★☆★☆★☆★☆★☆★☆
'==============================================================================*
'
'        MODULE_NAME      :func保存処理
'        機能             :保存処理の実行
'        IN               :
'        OUT              :True=正常 False=異常
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客情報変更理由_BeforeUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客情報変更理由_BeforeUpdate(Cancel As Integer)
  
  If Nz(Me.cmb_顧客情報変更理由.VALUE, "") = "" Then Exit Sub

  If Me.cmb_顧客情報変更理由.ListIndex < 0 Then
       MsgBox "そのコードは登録されていません。", vbInformation, "確認"
       Cancel = True
       SendKeys "{ESC}"
  End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客情報変更理由_KeyPress
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客情報変更理由_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_顧客情報変更理由, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客情報変更理由_AfterUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客情報変更理由_AfterUpdate()
    
    With Me.cmb_顧客情報変更理由
        If Nz(.VALUE, "") = "" Then
            Me.txt_顧客情報変更理由.VALUE = Null
        ElseIf .ListIndex >= 0 Then
            .VALUE = Format(.Text, .Format)
            Me.txt_顧客情報変更理由.VALUE = .Column(1, .ListIndex)
        End If
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_保証区分_BeforeUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_保証区分_BeforeUpdate(Cancel As Integer)

  If Nz(Me.cmb_保証区分.VALUE, "") = "" Then Exit Sub

  If Me.cmb_保証区分.ListIndex < 0 Then
       MsgBox "そのコードは登録されていません。", vbInformation, "確認"
       Cancel = True
       SendKeys "{ESC}"
  End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_保証区分_KeyPress
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_保証区分_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_保証区分, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_保証区分_AfterUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_保証区分_AfterUpdate()
    
    With Me.cmb_保証区分
        If Nz(.VALUE, "") = "" Then
            Me.txt_保証区分.VALUE = Null
        ElseIf .ListIndex >= 0 Then
            .VALUE = Format(.Text, .Format)
            Me.txt_保証区分.VALUE = .Column(1, .ListIndex)
        End If
        
        Me.cmd_保証会社 = ""                    'INSERT 2011/03/02 M.RYU----<S>
        Me.txt_保証会社 = ""
        If Nz(.VALUE, "") = "01" Then
            Me.cmd_保証会社.Visible = True
            Me.txt_保証会社.Visible = True
        Else
            Me.cmd_保証会社.Visible = False
            Me.txt_保証会社.Visible = False
        End If                                  'INSERT 2011/03/02 M.RYU----<E>
    End With
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日From_AfterUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日From_AfterUpdate()
    Me.txt_顧客情報変更日From = Format(var変更日付, "yyyy/mm/dd")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日From_BeforeUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日From_BeforeUpdate(Cancel As Integer)
    Call isDateFormat(Me.txt_顧客情報変更日From, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日From_GotFocus
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日From_GotFocus()
    txt_顧客情報変更日From = Format(txt_顧客情報変更日From, "yy/mm/dd")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日From_KeyPress
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日From_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(txt_顧客情報変更日From, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日From_LostFocus
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日From_LostFocus()
    txt_顧客情報変更日From = Format(txt_顧客情報変更日From, "yyyy/mm/dd")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日To_AfterUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日To_AfterUpdate()
    Me.txt_顧客情報変更日To = Format(var変更日付, "yyyy/mm/dd")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日To_BeforeUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日To_BeforeUpdate(Cancel As Integer)
    Call isDateFormat(Me.txt_顧客情報変更日To, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日To_GotFocus
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日To_GotFocus()
    txt_顧客情報変更日To = Format(txt_顧客情報変更日To, "yy/mm/dd")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日To_KeyPress
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日To_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(txt_顧客情報変更日To, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日To_LostFocus
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日To_LostFocus()
    txt_顧客情報変更日To = Format(txt_顧客情報変更日To, "yyyy/mm/dd")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :isDateFormat
'        機能             :日付入力ﾃﾞｰﾀのチェック
'        IN               :
'        OUT              :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function isDateFormat(aTarget As Variant, ByRef Cancel As Integer)

    isDateFormat = False
    var変更日付 = CmDateChecker(aTarget)
    If IsNull(var変更日付) Then
        MsgBox "日付を正しく入力してください。", vbInformation, "エラー"
        Cancel = True
        SendKeys "{ESC}"
    Else
        isDateFormat = True
    End If

End Function
'==============================================================================*
'
'        MODULE_NAME      :sub契約リスト出力
'        機能             :
'        IN               :
'        OUT              :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Sub sub契約リスト出力(ByVal a印刷モード As Integer)

  On Error GoTo Exception

    ' 必須入力チェック
    If Nz(Me.cmb_顧客情報変更理由.VALUE, "") = "" Then
      If (Nz(Me.cmb_CODEFROM.VALUE, "") = "" Or Nz(Me.cmb_CODETO.VALUE, "") = "") And _
         (Nz(Me.cmb_顧客FROM.VALUE, "") = "" Or Nz(Me.cmb_顧客TO.VALUE, "") = "") Then
          MsgBox "ヤードコードまたは顧客コードは範囲指定で選択して下さい。" & vbCr & _
                 "※ヤードまたは顧客はなるべく絞って下さい。", vbInformation, "確認"
         
          Exit Sub
      End If
    End If
      
    '処理確認
    Dim str確認文言 As String
    Select Case a印刷モード
        Case Print481.P481_MODE_EXCEL
              str確認文言 = "Excel出力を実行しますか？"
        Case Print481.P481_MODE_PRINT
              str確認文言 = "印刷を実行しますか？"
        Case Print481.P481_MODE_PREVIEW
              str確認文言 = "印刷内容を画面に出力しますか？"
    End Select
    If (MsgBox(str確認文言, vbYesNo + vbQuestion, "確認") = vbNo) Then
        Exit Sub
    End If

    '処理実行
    If False = Print481.PrintUserContractList(a印刷モード, _
                                    Nz(Me.txt_日付FROM.VALUE, ""), _
                                    Nz(Me.txt_日付FROM.VALUE, ""), _
                                    Nz(Me.cmb_CODEFROM.VALUE, ""), _
                                    Nz(Me.cmb_CODETO.VALUE, ""), _
                                    Nz(Me.cmb_顧客FROM.VALUE, ""), _
                                    Nz(Me.cmb_顧客TO.VALUE, ""), _
                                    Nz(Me.cmb_顧客情報変更理由.VALUE, ""), _
                                    Nz(Me.cmb_顧客情報変更理由.VALUE, ""), _
                                    Nz(Me.txt_顧客情報変更日From.VALUE, ""), _
                                    Nz(Me.txt_顧客情報変更日To.VALUE, ""), _
                                    Nz(Me.cmb_保証区分.VALUE, ""), _
                                    Nz(Me.cmb_保証区分.VALUE, ""), _
                                    Nz(Me.cmd_保証会社.VALUE, "") _
                                  ) Then        'UPDATE 2011/03/02 M.RYU '保証会社ｺｰﾄﾞを追加
    
    
        MsgBox "指定した条件に該当するデータはありませんでした。", vbInformation, "確認"
    
    End If
                                    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "sub契約リスト出力" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :frm_印刷種別_AfterUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub frm_印刷種別_AfterUpdate()
    
    Select Case frm_印刷種別.VALUE
        
        Case P_印刷種別_契約一覧表
        
              lbl_CODEFROM.Visible = True
              lblCODETO.Visible = True
              lbl_顧客TO.Visible = True
              cmb_CODEFROM.Visible = True
              cmb_CODETO.Visible = True
              txt_CODEFROM.Visible = True
              txt_CODETO.Visible = True
              Txt_顧客FROM.Visible = True
              lbl_顧客FROM.Visible = True
              cmb_顧客FROM.Visible = True
              cmd_プレビュー.Visible = True
        
              Me.lbl_顧客情報変更内容.Visible = False
              cmb_顧客情報変更理由.Visible = False
              txt_顧客情報変更理由.Visible = False
              lbl_顧客情報変更日.Visible = False
              txt_顧客情報変更日From.Visible = False
              lbl_顧客情報変更日bt.Visible = False
              txt_顧客情報変更日To.Visible = False
              lbl_保証区分.Visible = False
              cmb_保証区分.Visible = False
              txt_保証区分.Visible = False
              lbl_顧客FROM.Caption = "顧客ｺｰﾄﾞ(FROM)" '2007/07/30 add tajima
              lbl_顧客TO.Visible = True               '2007/07/30 add tajima
              cmb_顧客TO.Visible = True               '2007/07/30 add tajima
              txt_顧客TO.Visible = True               '2007/07/30 add tajima
              lbl_日付.Visible = True                 '2007/07/30 add tajima
              txt_日付FROM.Visible = True             '2007/07/30 add tajima
              lbl_日付から.Visible = True             '2007/07/30 add tajima
              txt_日付TO.Visible = True               '2007/07/30 add tajima
              lbl_売上年月.Visible = False            '2013/02/21 add M.HONDA
              txt_売上年月.Visible = False            '2013/02/21 add M.HONDA
              
        
        Case P_印刷種別_契約リスト
        
              lbl_CODEFROM.Visible = True
              lblCODETO.Visible = True
              lbl_顧客TO.Visible = True
              cmb_CODEFROM.Visible = True
              cmb_CODETO.Visible = True
              txt_CODEFROM.Visible = True
              txt_CODETO.Visible = True
              Txt_顧客FROM.Visible = True
              lbl_顧客FROM.Visible = True
              cmb_顧客FROM.Visible = True
              cmd_プレビュー.Visible = True
        
              Me.lbl_顧客情報変更内容.Visible = True
              cmb_顧客情報変更理由.Visible = True
              txt_顧客情報変更理由.Visible = True
              lbl_顧客情報変更日.Visible = True
              txt_顧客情報変更日From.Visible = True
              lbl_顧客情報変更日bt.Visible = True
              txt_顧客情報変更日To.Visible = True
              lbl_保証区分.Visible = True
              cmb_保証区分.Visible = True
              txt_保証区分.Visible = True
              lbl_顧客FROM.Caption = "顧客ｺｰﾄﾞ(FROM)" '2007/07/30 add tajima
              lbl_顧客TO.Visible = True               '2007/07/30 add tajima
              cmb_顧客TO.Visible = True               '2007/07/30 add tajima
              txt_顧客TO.Visible = True               '2007/07/30 add tajima
              lbl_日付.Visible = True                 '2007/07/30 add tajima
              txt_日付FROM.Visible = True             '2007/07/30 add tajima
              lbl_日付から.Visible = True             '2007/07/30 add tajima
              txt_日付TO.Visible = True               '2007/07/30 add tajima
              lbl_売上年月.Visible = False            '2013/02/21 add M.HONDA
              txt_売上年月.Visible = False            '2013/02/21 add M.HONDA
              
    
        '▼ '2007/07/30 add tajima
        Case P_印刷種別_廃止移動届
        
              lbl_CODEFROM.Visible = True
              lblCODETO.Visible = True
              lbl_顧客TO.Visible = True
              cmb_CODEFROM.Visible = True
              cmb_CODETO.Visible = True
              txt_CODEFROM.Visible = True
              txt_CODETO.Visible = True
              Txt_顧客FROM.Visible = True
              lbl_顧客FROM.Visible = True
              cmb_顧客FROM.Visible = True
              cmd_プレビュー.Visible = True
        
              Me.lbl_顧客情報変更内容.Visible = False
              cmb_顧客情報変更理由.Visible = False
              txt_顧客情報変更理由.Visible = False
              lbl_顧客情報変更日.Visible = False
              txt_顧客情報変更日From.Visible = False
              lbl_顧客情報変更日bt.Visible = False
              txt_顧客情報変更日To.Visible = False
              lbl_保証区分.Visible = True
              cmb_保証区分.Visible = True
              txt_保証区分.Visible = True
              lbl_顧客FROM.Caption = "顧客ｺｰﾄﾞ"
              lbl_顧客TO.Visible = False
              cmb_顧客TO.Visible = False
              txt_顧客TO.Visible = False
              lbl_日付.Visible = False
              txt_日付FROM.Visible = False
              lbl_日付から.Visible = False
              txt_日付TO.Visible = False
              cmb_保証区分.VALUE = "04" 'リプラスを初期設定
              Me.txt_保証区分.VALUE = Me.cmb_保証区分.Column(1, Me.cmb_保証区分.ListIndex)
        '▲ '2007/07/30 add tajima
              lbl_売上年月.Visible = False            '2013/02/21 add M.HONDA
              txt_売上年月.Visible = False            '2013/02/21 add M.HONDA
        
        
        Case P_印刷種別_一時使用契約変更承諾書          '----20100203----M.Ryu----add----<s>
        
              lbl_CODEFROM.Visible = True
              lblCODETO.Visible = True
              lbl_顧客TO.Visible = True
              cmb_CODEFROM.Visible = True
              cmb_CODETO.Visible = True
              txt_CODEFROM.Visible = True
              txt_CODETO.Visible = True
              Txt_顧客FROM.Visible = True
              lbl_顧客FROM.Visible = True
              cmb_顧客FROM.Visible = True
        
              Me.lbl_顧客情報変更内容.Visible = False
              cmb_顧客情報変更理由.Visible = False
              txt_顧客情報変更理由.Visible = False
              lbl_顧客情報変更日.Visible = False
              txt_顧客情報変更日From.Visible = False
              lbl_顧客情報変更日bt.Visible = False
              txt_顧客情報変更日To.Visible = False
              lbl_保証区分.Visible = False
              cmb_保証区分.Visible = False
              txt_保証区分.Visible = False
              lbl_顧客FROM.Caption = "顧客ｺｰﾄﾞ(FROM)"
              lbl_顧客TO.Visible = True
              cmb_顧客TO.Visible = True
              txt_顧客TO.Visible = True
              lbl_日付.Visible = False
              txt_日付FROM.Visible = False
              lbl_日付から.Visible = False
              txt_日付TO.Visible = False                '----20100203----M.Ryu----add----<e>
              lbl_売上年月.Visible = False            '2013/02/21 add M.HONDA
              txt_売上年月.Visible = False            '2013/02/21 add M.HONDA
              cmd_プレビュー.Visible = True
              
              
        Case P_印刷種別_補償加入費一覧
              Me.lbl_顧客情報変更内容.Visible = False
              cmb_顧客情報変更理由.Visible = False
              txt_顧客情報変更理由.Visible = False
              lbl_顧客情報変更日.Visible = False
              txt_顧客情報変更日From.Visible = False
              lbl_顧客情報変更日bt.Visible = False
              txt_顧客情報変更日To.Visible = False
              lbl_保証区分.Visible = False
              cmb_保証区分.Visible = False
              txt_保証区分.Visible = False
              lbl_顧客TO.Visible = False
              cmb_顧客TO.Visible = False
              txt_顧客TO.Visible = False
              lbl_日付.Visible = False
              txt_日付FROM.Visible = False
              lbl_日付から.Visible = False
              txt_日付TO.Visible = False
              lbl_売上年月.Visible = True             '2013/02/21 add M.HONDA
              txt_売上年月.Visible = True             '2013/02/21 add M.HONDA
              
              lbl_CODEFROM.Visible = False
              lblCODETO.Visible = False
              lbl_顧客TO.Visible = False
              cmb_CODEFROM.Visible = False
              cmb_CODETO.Visible = False
              txt_CODEFROM.Visible = False
              txt_CODETO.Visible = False
              Txt_顧客FROM.Visible = False
              lbl_顧客FROM.Visible = False
              cmb_顧客FROM.Visible = False
              cmd_プレビュー.Visible = False
        
        Case P_印刷種別_委託料加瀬負担一覧
        
              lbl_CODEFROM.Visible = False
              lblCODETO.Visible = False
              lbl_顧客TO.Visible = False
              cmb_CODEFROM.Visible = False
              cmb_CODETO.Visible = False
              txt_CODEFROM.Visible = False
              txt_CODETO.Visible = False
              Txt_顧客FROM.Visible = False
              lbl_顧客FROM.Visible = False
              cmb_顧客FROM.Visible = False
              cmd_プレビュー.Visible = True
        
              Me.lbl_顧客情報変更内容.Visible = False
              cmb_顧客情報変更理由.Visible = False
              txt_顧客情報変更理由.Visible = False
              lbl_顧客情報変更日.Visible = False
              txt_顧客情報変更日From.Visible = False
              lbl_顧客情報変更日bt.Visible = False
              txt_顧客情報変更日To.Visible = False
              lbl_保証区分.Visible = False
              cmb_保証区分.Visible = False
              txt_保証区分.Visible = False
              lbl_顧客FROM.Caption = "顧客ｺｰﾄﾞ(FROM)"
              lbl_顧客TO.Visible = False
              cmb_顧客TO.Visible = False
              txt_顧客TO.Visible = False
              lbl_日付.Visible = True
              txt_日付FROM.Visible = True
              lbl_日付から.Visible = True
              txt_日付TO.Visible = True
              lbl_売上年月.Visible = False
              txt_売上年月.Visible = False
              lbl_売上年月.Caption = "売上年月"
        Case P_印刷種別_配送サービス一覧 '20150108 M.HONDA INS
        
              lbl_CODEFROM.Visible = True
              lblCODETO.Visible = True
              lbl_顧客TO.Visible = True
              cmb_CODEFROM.Visible = True
              cmb_CODETO.Visible = True
              txt_CODEFROM.Visible = True
              txt_CODETO.Visible = True
              Txt_顧客FROM.Visible = True
              lbl_顧客FROM.Visible = True
              cmb_顧客FROM.Visible = True
              cmd_プレビュー.Visible = True
        
              Me.lbl_顧客情報変更内容.Visible = False
              cmb_顧客情報変更理由.Visible = False
              txt_顧客情報変更理由.Visible = False
              lbl_顧客情報変更日.Visible = False
              txt_顧客情報変更日From.Visible = False
              lbl_顧客情報変更日bt.Visible = False
              txt_顧客情報変更日To.Visible = False
              lbl_保証区分.Visible = False
              cmb_保証区分.Visible = False
              txt_保証区分.Visible = False
              lbl_顧客FROM.Caption = "顧客ｺｰﾄﾞ(FROM)"
              lbl_顧客TO.Visible = True
              cmb_顧客TO.Visible = True
              txt_顧客TO.Visible = True
              lbl_日付.Visible = True
              txt_日付FROM.Visible = True
              lbl_日付から.Visible = True
              txt_日付TO.Visible = True
              lbl_売上年月.Visible = False
              txt_売上年月.Visible = False
        
         Case P_印刷種別_予約リスト  '2019/09/07 M.HONDA INS
        
              Me.lbl_顧客情報変更内容.Visible = False
              cmb_顧客情報変更理由.Visible = False
              txt_顧客情報変更理由.Visible = False
              lbl_顧客情報変更日.Visible = False
              txt_顧客情報変更日From.Visible = False
              lbl_顧客情報変更日bt.Visible = False
              txt_顧客情報変更日To.Visible = False
              lbl_保証区分.Visible = False
              cmb_保証区分.Visible = False
              txt_保証区分.Visible = False
              lbl_顧客TO.Visible = False
              cmb_顧客TO.Visible = False
              txt_顧客TO.Visible = False
              lbl_日付.Visible = False
              txt_日付FROM.Visible = False
              lbl_日付から.Visible = False
              txt_日付TO.Visible = False
              lbl_売上年月.Visible = True
              txt_売上年月.Visible = True
              
              lbl_CODEFROM.Visible = False
              lblCODETO.Visible = False
              lbl_顧客TO.Visible = False
              cmb_CODEFROM.Visible = False
              cmb_CODETO.Visible = False
              txt_CODEFROM.Visible = False
              txt_CODETO.Visible = False
              Txt_顧客FROM.Visible = False
              lbl_顧客FROM.Visible = False
              cmb_顧客FROM.Visible = False
              cmd_プレビュー.Visible = False
              lbl_売上年月.Caption = "予約年月"
    End Select
    
    cmd_保証会社.Visible = False            '2011/03/02 add m.ryu
    txt_保証会社.Visible = False            '2011/03/02 add m.ryu
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_Excel_Click
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()
    
    On Error GoTo Exception

    If frm_印刷種別.VALUE = P_印刷種別_契約一覧表 Then
        MsgBox "顧客契約一覧ではExcel出力はできません。", vbInformation, strMsg_kaku
        Exit Sub
    End If

    '▼ 2007/07/30 add tajima
    If frm_印刷種別.VALUE = P_印刷種別_契約リスト Then
        Call sub契約リスト出力(Print481.P481_MODE_EXCEL)
        Exit Sub
    End If
    
    If frm_印刷種別.VALUE = P_印刷種別_廃止移動届 Then
        Call sub廃止移動届出力(Print230.P230_MODE_EXCEL)
        Exit Sub
    End If
    '▲ 2007/07/30 add tajima

    '----20100203----M.Ryu----add-------<s>
    If frm_印刷種別.VALUE = P_印刷種別_一時使用契約変更承諾書 Then
        Call sub一時使用契約変更承諾書(Print231.P231_MODE_EXCEL)
        Exit Sub
    End If
    '----20100203----M.Ryu----add-------<e>

    '----20130220----M.HONDA----add-------<s>
    If frm_印刷種別.VALUE = P_印刷種別_補償加入費一覧 Then
        Call sub補償加入費一覧(Print232.P232_MODE_EXCEL)
        Exit Sub
    End If
    '----20130220----M.HONDA----add-------<e>

    '----20130501----M.HONDA----add-------<s>
    If frm_印刷種別.VALUE = P_印刷種別_委託料加瀬負担一覧 Then
        Call sub委託料加瀬負担一覧(Print232.P232_MODE_EXCEL)
        Exit Sub
    End If
    '----20130501----M.HONDA----add-------<e>
    
    If frm_印刷種別.VALUE = P_印刷種別_配送サービス一覧 Then
        Call 配送サービス一覧(Print232.P232_MODE_EXCEL)
        Exit Sub
    End If

    '2019/09/07 M.HONDA INS
    If frm_印刷種別.VALUE = P_印刷種別_予約リスト Then
        Call sub予約リスト(Print232.P232_MODE_EXCEL)
        Exit Sub
    End If


    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Excel_Click", True)
End Sub
'★☆★☆★☆★☆★☆★★☆☆ 2007/07/30 add tajima ★☆★☆★☆★☆★☆★☆★☆
'==============================================================================*
'
'        MODULE_NAME      :sub廃止移動届出力
'        機能             :
'        IN               :
'        OUT              :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Sub sub廃止移動届出力(ByVal a印刷モード As Integer)

  On Error GoTo Exception

    ' 必須入力チェック
    If (Nz(Me.cmb_CODEFROM.VALUE, "") = "" Or Nz(Me.cmb_CODETO.VALUE, "") = "") Then
        MsgBox "ヤードコードは範囲指定で選択して下さい。" & vbCr & _
               "※ヤードはなるべく絞って下さい。", vbInformation, "確認"
       
        Exit Sub
    End If
      
    '処理確認
    Dim str確認文言 As String
    Select Case a印刷モード
        Case Print230.P230_MODE_EXCEL
              str確認文言 = "Excel出力を実行しますか？"
        Case Print230.P230_MODE_PRINT
              str確認文言 = "印刷を実行しますか？"
        Case Print230.P230_MODE_PREVIEW
              str確認文言 = "印刷内容を画面に出力しますか？"
    End Select
    If (MsgBox(str確認文言, vbYesNo + vbQuestion, "確認") = vbNo) Then
        Exit Sub
    End If

    '処理実行
    If False = Print230.PrintUserMoveList(a印刷モード, _
                                    Nz(Me.cmb_CODEFROM.VALUE, ""), _
                                    Nz(Me.cmb_CODETO.VALUE, ""), _
                                    "", _
                                    Nz(Me.cmb_顧客FROM.VALUE, ""), _
                                    Nz(Me.cmb_保証区分.VALUE, "") _
                                  ) Then
    
    
        MsgBox "指定した条件に該当するデータはありませんでした。", vbInformation, "確認"
    
    End If
                                    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "sub廃止移動届出力" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'★☆★☆★☆★☆★☆★★☆☆ 2007/07/30 add tajima ★☆★☆★☆★☆★☆★☆★☆
'==============================================================================*
'
'        MODULE_NAME      :sub一時使用契約変更承諾書
'        機能             :一時使用契約変更承諾書をﾌﾟﾚﾋﾞｭｰ・Excel出力
'        IN               :印刷モード(Integer)
'        OUT              :
'        作成日付         :20100203--M.Ryu
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Sub sub一時使用契約変更承諾書(ByVal a印刷モード As Integer)

  On Error GoTo Exception

'    ' 必須入力チェック
'    If (Nz(Me.cmb_CODEFROM.VALUE, "") = "" Or Nz(Me.cmb_CODETO.VALUE, "") = "") Then
'        MsgBox "ヤードコードは範囲指定で選択して下さい。" & vbCr & _
'               "※ヤードはなるべく絞って下さい。", vbInformation, "確認"
'
'        Exit Sub
'    End If
      
    '処理確認
    Dim str確認文言 As String
    Select Case a印刷モード
        Case Print231.P231_MODE_EXCEL
              str確認文言 = "Excel出力を実行しますか？"
        Case Print231.P231_MODE_PREVIEW
              str確認文言 = "印刷内容を画面に出力しますか？"
    End Select
    If (MsgBox(str確認文言, vbYesNo + vbQuestion, "確認") = vbNo) Then
        Exit Sub
    End If

    '処理実行
    If Not Print231.PrintUserMoveList(a印刷モード, Nz(Me.cmb_CODEFROM.VALUE, ""), _
                Nz(Me.cmb_CODETO.VALUE, ""), Nz(Me.cmb_顧客FROM.VALUE, ""), _
                Nz(Me.cmb_顧客TO.VALUE, "")) Then
        MsgBox "指定した条件に該当するデータはありませんでした。", vbInformation, "確認"
    End If
                                    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "sub一時使用契約変更承諾書" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :sub補償加入費一覧
'        機能             :補償加入費一覧をﾌﾟﾚﾋﾞｭｰ・Excel出力
'        IN               :印刷モード(Integer)
'        OUT              :
'        作成日付         :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Sub sub補償加入費一覧(ByVal a印刷モード As Integer)

  On Error GoTo Exception

    ' 必須入力チェック
    If (Nz(Me.txt_売上年月.VALUE, "")) = "" Then
        MsgBox "売上年月を指定してください。", vbInformation, "確認"
        Exit Sub
    End If
      
    '処理確認
    Dim str確認文言 As String
    Select Case a印刷モード
        Case Print232.P232_MODE_EXCEL
              str確認文言 = "Excel出力を実行しますか？"
        Case Print232.P232_MODE_PREVIEW
              str確認文言 = "印刷内容を画面に出力しますか？"
    End Select
    If (MsgBox(str確認文言, vbYesNo + vbQuestion, "確認") = vbNo) Then
        Exit Sub
    End If

    '処理実行
    If Not Print232.PrintUserMoveList(a印刷モード, Nz(Me.txt_売上年月.VALUE, "")) Then
        MsgBox "指定した条件に該当するデータはありませんでした。", vbInformation, "確認"
    End If
                                    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "sub補償加入費一覧" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :sub予約リスト
'        機能             :予約リストをﾌﾟﾚﾋﾞｭｰ・Excel出力
'        IN               :印刷モード(Integer)
'        OUT              :
'        作成日付         :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Sub sub予約リスト(ByVal a印刷モード As Integer)

    On Error GoTo Exception
    Dim strMsg     As String
    Dim intAns     As Integer

 
    ' 必須入力チェック
    If (Nz(Me.txt_売上年月.VALUE, "")) = "" Then
        MsgBox "予約年月を指定してください。", vbInformation, "確認"
        Exit Sub
    End If
    
    strMsg = "印刷内容をEXCELに出力します。" '& vbCr
    strMsg = strMsg & "よろしいですか？"
    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)

    '*** ワーク作成 ***
    If Make_yoyaku = False Then
        '*** エラー発生またはデータなし ***
        Exit Sub
    End If

    doCmd.OutputTo acOutputTable, C_TEMP_TABLE, "MicrosoftExcel(*.xls)", "", True, ""

    Exit Sub


Exception:
    Call Err.Raise(Err.Number, "sub補償加入費一覧" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :sub委託料加瀬負担一覧
'        機能             :委託料加瀬負担一覧をﾌﾟﾚﾋﾞｭｰ・Excel出力
'        IN               :印刷モード(Integer)
'        OUT              :
'        作成日付         :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Sub sub委託料加瀬負担一覧(ByVal a印刷モード As Integer)

  On Error GoTo Exception

    ' 必須入力チェック
    If (Nz(Me.txt_日付FROM.VALUE, "")) = "" Or (Nz(Me.txt_日付TO.VALUE, "")) = "" Then
        MsgBox "契約日を指定してください。", vbInformation, "確認"
        Exit Sub
    End If
      
    '処理確認
    Dim str確認文言 As String
    Select Case a印刷モード
        Case Print232.P232_MODE_EXCEL
              str確認文言 = "Excel出力を実行しますか？"
        Case Print232.P232_MODE_PREVIEW
              str確認文言 = "印刷内容を画面に出力しますか？"
    End Select
    If (MsgBox(str確認文言, vbYesNo + vbQuestion, "確認") = vbNo) Then
        Exit Sub
    End If

    '処理実行
    If Not Print233.PrintUserMoveList(a印刷モード, Nz(Me.txt_日付FROM.VALUE, ""), Nz(Me.txt_日付TO.VALUE, "")) Then
        MsgBox "指定した条件に該当するデータはありませんでした。", vbInformation, "確認"
    End If
                                    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "sub委託料加瀬負担一覧" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :配送サービス一覧
'        機能             :
'        IN               :
'        OUT              :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Sub 配送サービス一覧(ByVal a印刷モード As Integer)

  On Error GoTo Exception

    ' 必須入力チェック
    If (Nz(Me.cmb_CODEFROM.VALUE, "") = "" Or Nz(Me.cmb_CODETO.VALUE, "") = "") And _
       (Nz(Me.cmb_顧客FROM.VALUE, "") = "" Or Nz(Me.cmb_顧客TO.VALUE, "") = "") Then
       MsgBox "ヤードコードまたは顧客コードは範囲指定で選択して下さい。" & vbCr & _
         "※ヤードまたは顧客はなるべく絞って下さい。", vbInformation, "確認"
          Exit Sub
     End If
      
    '処理確認
    Dim str確認文言 As String
    Select Case a印刷モード
        Case Print234.P234_MODE_EXCEL
              str確認文言 = "Excel出力を実行しますか？"
        Case Print234.P234_MODE_PREVIEW
              str確認文言 = "印刷内容を画面に出力しますか？"
    End Select
    If (MsgBox(str確認文言, vbYesNo + vbQuestion, "確認") = vbNo) Then
        Exit Sub
    End If

    '処理実行
    If False = Print234.PrintDlveList(a印刷モード, _
                                    Nz(Me.txt_日付FROM.VALUE, ""), _
                                    Nz(Me.txt_日付FROM.VALUE, ""), _
                                    Nz(Me.cmb_CODEFROM.VALUE, ""), _
                                    Nz(Me.cmb_CODETO.VALUE, ""), _
                                    Nz(Me.cmb_顧客FROM.VALUE, ""), _
                                    Nz(Me.cmb_顧客TO.VALUE, "")) Then
    
    
        MsgBox "指定した条件に該当するデータはありませんでした。", vbInformation, "確認"
    
    End If
                                    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "sub契約リスト出力" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Make_yoyaku() As Boolean
On Error GoTo Err_Make_Work
'*** ヤードマスタ一覧表作成 ***
    Dim strWhere  As String
    Dim lngKensu As Long
    Dim strSQL As String
    Dim dbMydb As Database
    Dim strMsg As String
    Dim intAns As Integer
    Dim varChData As Variant
    Dim strBumonCode As String
    Dim strDataSource As String
    Dim lngCnt As Integer
    Dim rsSqlServer         As Object       'ADODB.Recordset
    Dim strfrom As String
    Dim strTo As String
    
    
    '部門コード取得
    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    Make_yoyaku = False
    doCmd.Hourglass True

    strTo = DateSerial(Year(Format(Me.txt_売上年月 & "01", "####/##/##")), Month(Format(Me.txt_売上年月 & "01", "####/##/##")), 1)
    strfrom = DateSerial(Year(Format(Me.txt_売上年月 & "01", "####/##/##")), Month(Format(Me.txt_売上年月 & "01", "####/##/##")) + 1, 0)



   'SQL-Server接続
    Dim objCon         As Object       'ADODB.Connection
    Set objCon = ADODB_Connection(strBumonCode)
    
    strSQL = strSQL & " select CONVERT(VARCHAR(30), YOUKT_UKDATE, 111)  as 予約日, "
    strSQL = strSQL & "        SUM(case YOUKT_GENKBN when 1 then 1 else 0 end) as ネット, "
    strSQL = strSQL & "        SUM(case YOUKT_GENKBN when 0 then 1 else 0 end) as 電話 "
    strSQL = strSQL & " from   YOUK_TRAN "
    strSQL = strSQL & "        inner join NAME_MAST NAME_080 on "
    strSQL = strSQL & "          NAME_080.NAME_ID = '080' and "
    strSQL = strSQL & "          NAME_080.NAME_CODE = YOUKT_YUKBN "
    strSQL = strSQL & "        inner join NAME_MAST NAME_083 on "
    strSQL = strSQL & "          NAME_083.NAME_ID = '083' and "
    strSQL = strSQL & "          NAME_083.NAME_CODE = YOUKT_GENKBN "
    strSQL = strSQL & " where  YOUKT_UKDATE between '" & Format(strTo, "yyyymmdd") & "' and '" & Format(strfrom, "yyyymmdd") & "'"
    strSQL = strSQL & " and    YOUKT_MOVEKBN = '0' "
    strSQL = strSQL & " and  (YOUKT_MAIL  not like '%kasegroup%' or YOUKT_MAIL  is null ) "
    strSQL = strSQL & " AND    YOUKT_NAME  not like '%テスト%' "
    strSQL = strSQL & " group by CONVERT(VARCHAR(30), YOUKT_UKDATE, 111)  "
    strSQL = strSQL & " order by CONVERT(VARCHAR(30), YOUKT_UKDATE, 111)  "
    
    lngCnt = ADODB_CreateWorkTable_CurrentDb(C_TEMP_TABLE, strSQL, objCon)

 
    '*** データチェック ***
    'アクションクエリーの結果で判断
    If lngCnt < 0 Then
        strMsg = "対象となるデータが存在しませんでした。"
        MsgBox strMsg, vbInformation + vbOKOnly
        GoTo Exit_Make_Work
    End If
    
    Make_yoyaku = True

Exit_Make_Work:
    doCmd.Hourglass False
    Exit Function

Err_Make_Work:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Make_Work

End Function

'****************************  ended or program ********************************

