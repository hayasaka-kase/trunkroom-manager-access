Attribute VB_Name = "LS0060"
'****************************  strat of program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : コンテナ管理システム
'        SUB_SYSTEM_NAME : 帳票
'
'        PROGRAM_NAME    : コンテナ在庫一覧表
'        PROGRAM_ID      : LS0060
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2019/11/20
'        CERATER         : Y.WADA
'        Ver             : 0.0
'
'        UPDATE          : 2020/06/09
'        UPDATER         : Y.WADA
'        Ver             : 0.1
'                        : 仕入商品区分の判断方法変更
'                          タイトル／シート名変更
'                          　自社ISOコンテナ　 →　スタンダードコンテナ（ISO）
'                          　自社JISコンテナ　 →　建築確認コンテナ（JIS）
'                          　自社コンテナ以外　→　その他コンテナなど
'                          各シート共通
'                              ②出港数・・・800017：直送輸入済みコンテナ分を含む
'                              ④海外からの入庫数　※②と同様・営業ヤードへ直送分は除く・・・800017：直送輸入済みコンテナ分を含まない
'                              ⑦営業ヤードへの出庫数・・・800017：直送輸入済みコンテナ分を含まない
'                              ⑨梶山ヤードからの入庫数※⑦と同様・港から営業ヤードへ直送分含む・・・800017：直送輸入済みコンテナ分を含む
'                          自社コンテナ以外
'                              サイズ、ドア数を追加
'                              ※バイクボックスにサイズとドア数がありそれぞれ分かれているので追加してください。台車やｶﾞｰﾄﾞﾏﾝﾎﾞｯｸｽはサイズなどないので空白でかまいません。
'                          Excel出力のレスポンス改善（1分半→10秒）
'
'        UPDATE          : 2020/07/07
'        UPDATER         : Y.WADA
'        Ver             : 0.1
'                        : 項目追加：港保管、不足在庫
'
'        UPDATE          : 2022/09/30
'        UPDATER         : K.KINEBUCHI
'        Ver             : 0.3
'                        : 項目追加：港保管前月在庫、港保管入庫数、港保管出庫数
'                        : 　　　　　中古購入前月在庫、中古購入から自社発注、中古購入から梶山ヤードへ出港
'
'==============================================================================*
Option Compare Database
Option Explicit

Const CON_FORM_NAME As String = "LF0060"

Private Sub TEST_LS0060_xlPreview()
    Call LS0060_xlPreview("201910")
    
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 書類プレビュー
'       MODULE_ID       : xlPrintPreview
'       CREATE_DATE     : 2019/11/20            Y.WADA
'       PARAM           : strSETDATE_YM         設置日（年月）YYYYMM(I)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function LS0060_xlPreview(ByVal strSETDATE_YM As String) As Boolean
    Dim strPath             As String
    Dim strExcel            As String
    Dim strFile             As String
    Dim xlApp               As Object
    Dim xlBook              As Object
    Dim blnRet              As Boolean
    On Error GoTo ErrorHandler
    
    'ファイル名取得
    strFile = "在庫一覧表"
    strExcel = Select_INIT_FILE(strFile)
    
    If strExcel = "" Then
        Call MSZZ024_M10("OPEN", strFile & "INTIテーブルに設定がありません。[" & strFile & "]")
    End If
    
    If Dir(strExcel, vbNormal) = "" Then
        Call MSZZ024_M10("OPEN", strFile & "ファイルが見つかりませんでした。[" & strExcel & "]")
    End If
    
    '保存先の指定がある場合
    strPath = Select_INIT_FILE("出力先")
    If strPath <> "" Then
        If Right(strPath, 1) <> "\" Then
            strPath = strPath & "\"
        End If
        strFile = strPath & strSETDATE_YM & "_" & Dir(strExcel)
        '保存するファイル名が残っていたら削除する
        If Dir(strFile, vbNormal) <> "" Then
            On Error Resume Next
            Kill strFile
            '削除できないときは閉じてねメッセージ
            If Err.Number = 70 Then
                Err.Clear
                On Error GoTo ErrorHandler
                MsgBox "ファイルが開かれています。" & vbCrLf & _
                    strFile & vbCrLf & vbCrLf & _
                    "ファイルを閉じて再実行してください。", vbOKOnly + vbExclamation, "LS0060"
                LS0060_xlPreview = True
                Exit Function
            End If
            On Error GoTo ErrorHandler
        End If
    Else
        strFile = ""
    End If

    'エクセル起動
    Set xlApp = CreateObject("Excel.Application")
    On Error GoTo ErrorHandler1
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    Set xlBook = xlApp.Workbooks.Open(strExcel, 0, True) 'ReadOnly で開く
    On Error GoTo ErrorHandler2
    
    'データ取得
    blnRet = False
    blnRet = blnRet Or GetData_ZaikoIchiran(xlBook, strSETDATE_YM, 1)     '1...ISO
    blnRet = blnRet Or GetData_ZaikoIchiran(xlBook, strSETDATE_YM, 2)     '2...JIS
    blnRet = blnRet Or GetData_ZaikoIchiran(xlBook, strSETDATE_YM, 3)     '3...その他
    
    If blnRet Then
        If strFile <> "" Then
            '保存先の指定がある場合、保存する
            If Dir(strFile, vbNormal) <> "" Then
                Kill strFile
            End If
            xlApp.DisplayAlerts = False
'            'シート選択
'            xlBook.sheets(2).SELECT False
'            xlBook.sheets(1).SELECT False
            xlBook.SaveAs strFile
        End If
        '表示する
        xlApp.ScreenUpdating = True
        xlApp.Visible = True
        On Error GoTo ErrorHandler
        LS0060_xlPreview = True
    Else
        xlBook.Close False
        On Error GoTo ErrorHandler1
        xlApp.DisplayAlerts = False
        xlApp.Quit
        On Error GoTo ErrorHandler
        LS0060_xlPreview = False
    End If
Exit Function

ErrorHandler2:
    xlBook.Close False
ErrorHandler1:
    xlApp.DisplayAlerts = False
    xlApp.Quit
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "LS0060_xlPreview" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : INIT_FILE読込
'       MODULE_ID       : Select_INIT_FILE
'       CREATE_DATE     : 2019/11/20            Y.WADA
'       PARAM           : strINTIF_RECFB        ファイル種類or出力先(I)
'                       : [strBUMOC]            部門コード：省略可
'       RETURN          : パス(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Select_INIT_FILE(ByVal strINTIF_RECFB As String, Optional strBUMOC As String = "") As String
    Dim strINTIF_RECDB      As String
    Dim strWhere            As String
    On Error GoTo ErrorHandler
    
    strINTIF_RECDB = ""
'    strWhere = "INTIF_PROGB = '" & Me.NAME & "' AND INTIF_RECFB = '" & strINTIF_RECFB
    strWhere = "INTIF_PROGB = '" & CON_FORM_NAME & "' AND INTIF_RECFB = '" & strINTIF_RECFB

    If strBUMOC <> "" Then
        strINTIF_RECDB = Nz(DLookup("INTIF_RECDB", "INTI_FILE", strWhere & "-" & strBUMOC & "'"), "")
    End If
    If strINTIF_RECDB = "" Then
        strINTIF_RECDB = Nz(DLookup("INTIF_RECDB", "INTI_FILE", strWhere & "'"), "")
    End If
    Select_INIT_FILE = strINTIF_RECDB
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Select_INIT_FILE" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : データ取得(在庫一覧表)
'       MODULE_ID       : GetData_ZaikoIchiran
'       CREATE_DATE     : 2019/11/20            Y.WADA
'       PARAM           : xlBook                エクセルブックオブジェクト(I)
'                       : strSETDATE_YM         設置日（年月）YYYYMM(I)
'                       : strZaikoKbn           在庫種類(I)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetData_ZaikoIchiran(xlBook As Object, strSETDATE_YM As String, intZaikoKbn As Integer) As Boolean
    Dim varValue()          As Variant
    Dim strCONT_BUMOC       As String
    Dim objCon              As Object
    Dim objRst              As Object
    Dim xlSheet             As Object
    Dim lngFldCnt           As Long
    Dim lngRow              As Long
    Dim lngCol              As Long
    Dim i                   As Long
    Dim j                   As Long
    
    'INSERT 2020/06/09 Y.WADA Start
    
    Dim lngRow_start        As Long
    Dim lngRow_end          As Long
    Dim lngCol_start        As Long
    Dim lngCol_end          As Long
    'INSERT 2020/06/09 Y.WADA End
    
    On Error GoTo ErrorHandler

    '部門コード取得
    strCONT_BUMOC = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    Set objCon = ADODB_Connection(strCONT_BUMOC)
    On Error GoTo ErrorHandler1
    
    Set objRst = ADODB_Recordset(Select_ZaikoIchiran(strSETDATE_YM, intZaikoKbn), objCon)
    On Error GoTo ErrorHandler2
    
    Set xlSheet = xlBook.Sheets(intZaikoKbn)
    
    'ヘッダ情報
    xlSheet.Range("設置年月").VALUE = Left(strSETDATE_YM, 4) & "/" & Mid(strSETDATE_YM, 5, 2) & "/01"
    
'DELETE 2020/06/09 Y.WADA Start
'    '明細行
'    With xlSheet.Range("data")
'        lngRow = .row
'        lngCol = .Column
'    End With
'DELETE 2020/06/09 Y.WADA End

    With objRst
'        lngFldCnt = .Fields.count - 1  'DELETE 2020/06/09 Y.WADA
        If Not .EOF Then
            
            'INSERT 2020/06/09 Y.WADA Start
            'DBデータ配列初期化
            lngFldCnt = objRst.Fields.Count - 1
            lngRow_start = xlSheet.Range("data").row
            lngRow_end = xlSheet.Range("data_end").row
            lngCol_start = xlSheet.Range("data").Column
            lngCol_end = lngCol_start + lngFldCnt
            ReDim varValue(lngRow_start To lngRow_end, lngCol_start To lngCol_end)
            
            '明細行
            lngRow = lngRow_start
            lngCol = lngCol_start
            'INSERT 2020/06/09 Y.WADA End
            
            While Not .EOF()
'DELETE 2020/06/09 Y.WADA Start
'                ReDim varValue(1 To 1, lngCol To lngCol + lngFldCnt)
'                For i = 0 To lngFldCnt
'                    varValue(1, lngCol + i) = .Fields(i).VALUE
'                Next
'                With xlSheet
'                    '明細行追加
'                    .Rows(lngRow).insert
'                    .Rows(lngRow + 1).Copy .Rows(lngRow)
'
'                    .Range(.Cells(lngRow, lngCol), .Cells(lngRow, lngCol + lngFldCnt)).VALUE = varValue
'                End With
'DELETE 2020/06/09 Y.WADA End
                'INSERT 2020/06/09 Y.WADA Start
                If lngRow > lngRow_end Then
                    Call MSZZ024_M10("GetData", "テンプレートの最大行数を超えました。テンプレートに行追加して再実行してください。")
                End If
                
                For i = 0 To lngFldCnt
                   varValue(lngRow, lngCol + i) = objRst.Fields(i).VALUE
                Next
                'INSERT 2020/06/09 Y.WADA Start
                
                lngRow = lngRow + 1
                .MoveNext
            Wend

            'INSERT 2020/06/09 Y.WADA Start
            'DBデータ配列をセルにセット
            xlSheet.Range(xlSheet.Cells(lngRow_start, lngCol_start), xlSheet.Cells(lngRow - 1, lngCol_end)).VALUE = varValue
        
            If lngRow <= lngRow_end Then
                '余った行を削除
                xlSheet.Rows(lngRow & ":" & Format(lngRow_end)).Delete
            End If
            'INSERT 2020/06/09 Y.WADA End

'            xlBook.sheets(1).SELECT
            GetData_ZaikoIchiran = True
        Else
            GetData_ZaikoIchiran = False
        End If
        .Close
        On Error GoTo ErrorHandler1
    End With
    objCon.Close
    On Error GoTo ErrorHandler

'    xlSheet.Range("data").EntireRow.Hidden = True  'DELETE 2020/06/09 Y.WADA Start

Exit Function

ErrorHandler2:
    objRst.Close
ErrorHandler1:
    objCon.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "GetData_ZaikoIchiran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'       MODULE_NAME     : SELECT文（在庫一覧表）
'       MODULE_ID       : Select_ZaikoIchiran
'       CREATE_DATE     : 2019/11/20            Y.WADA
'       PARAM           : strSETDATE_YM         設置日（年月）YYYYMM(I)
'                       : strZaikoKbn           在庫種類(I)
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Select_ZaikoIchiran(ByVal strSETDATE_YM As String, intZaikoKbn As Integer) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = strSQL & vbCrLf & "DECLARE @SETDATE_YM CHAR(6) = '" & strSETDATE_YM & "';    --設置日（年月）"
    strSQL = strSQL & vbCrLf & "SELECT"
    strSQL = strSQL & vbCrLf & "    nm451.NAME_VALUE_FROM   AS [在庫種類]"
    strSQL = strSQL & vbCrLf & ",   ZAIKM_SISYI             AS [仕入商品区分]"
    strSQL = strSQL & vbCrLf & ",   nm451.NAME_NAME         AS [仕入商品名]"
    strSQL = strSQL & vbCrLf & ",   ZAIKM_SIZE              AS [サイズ]"
    strSQL = strSQL & vbCrLf & ",   nm271.NAME_KANA         AS [サイズ名]"
    strSQL = strSQL & vbCrLf & ",   ZAIKM_DOOR              AS [ドア数]"
    strSQL = strSQL & vbCrLf & ",   nm272.NAME_RYAK         AS [ドア数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKM_ZASTD             AS [在庫開始日]"
    strSQL = strSQL & vbCrLf & ",   ZAIKM_ZAEDD             AS [在庫終了日]"
    strSQL = strSQL & vbCrLf & ",   ZAIKM_PRICE             AS [単価]"
    strSQL = strSQL & vbCrLf & ",   ZAIKM_ORDER             AS [出力順]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZAISD             AS [在庫集計年月]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZKAIQ             AS [海外前月在庫数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZKAJQ             AS [梶山在庫前月在庫数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZEIGQ             AS [営業ヤード前月在庫数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_HACYQ             AS [発注数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_SYUKQ             AS [出港数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_KAINQ             AS [海外からの入庫数]"          'INSERT 2020/06/09 Y.WADA
    strSQL = strSQL & vbCrLf & ",   ZAIKS_TEKYQ             AS [営業ヤードからの撤去数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_BAIKQ             AS [売却数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ESYKQ             AS [営業ヤードへの出庫数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_KAJNQ             AS [梶山ヤードからの入庫数]"    'INSERT 2020/06/09 Y.WADA
    strSQL = strSQL & vbCrLf & ",   ZAIKS_MIHOQ             AS [港保管]"        'INSERT 2020/07/07 Y.WADA
    strSQL = strSQL & vbCrLf & ",   ZAIKS_FUSZQ             AS [不足在庫]"      'INSERT 2020/07/07 Y.WADA
'INSERT 2022/09/30 K.KINEBUCHI start
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZMIHQ             AS [港保管前月在庫数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_MHNYQ             AS [港保管入庫数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_MHSYQ             AS [港保管出庫数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZTASQ             AS [中古購入前月在庫数]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZTHAQ             AS [中古購入から自社へ発注]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZTSKQ             AS [中古購入から梶山ヤードへ出庫]"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZTSEQ             AS [中古購入から営業ヤードへ出庫]"
'INSERT 2022/09/30 K.KINEBUCHI end
    strSQL = strSQL & vbCrLf & "FROM    ZAIK_MAST"
    strSQL = strSQL & vbCrLf & "INNER JOIN NAME_MAST nm451"
    strSQL = strSQL & vbCrLf & "    ON  nm451.NAME_ID = '451'   --仕入商品区分"
    strSQL = strSQL & vbCrLf & "    AND nm451.NAME_CODE = ZAIKM_SISYI"
    strSQL = strSQL & vbCrLf & "LEFT JOIN NAME_MAST nm271"
    strSQL = strSQL & vbCrLf & "    ON  nm271.NAME_ID = '271'   --サイズ"
    strSQL = strSQL & vbCrLf & "    AND nm271.NAME_CODE = ZAIKM_SIZE"
    strSQL = strSQL & vbCrLf & "LEFT JOIN NAME_MAST nm272"
    strSQL = strSQL & vbCrLf & "    ON  nm272.NAME_ID = '272'   --ドア数"
    strSQL = strSQL & vbCrLf & "    AND nm272.NAME_CODE = ZAIKM_DOOR"
    strSQL = strSQL & vbCrLf & "LEFT JOIN   ZAIK_SUMM"
    strSQL = strSQL & vbCrLf & "    ON  ZAIKS_SISYI =   ZAIKM_SISYI"
    strSQL = strSQL & vbCrLf & "    AND ISNULL(ZAIKS_SIZE,0)    =   ISNULL(ZAIKM_SIZE, 0)"
    strSQL = strSQL & vbCrLf & "    AND ISNULL(ZAIKS_DOOR, 0)   =   ISNULL(ZAIKM_DOOR, 0)"
    strSQL = strSQL & vbCrLf & "    AND ZAIKS_ZAISD =   @SETDATE_YM"
    strSQL = strSQL & vbCrLf & "WHERE"
    strSQL = strSQL & vbCrLf & "    @SETDATE_YM BETWEEN LEFT(ZAIKM_ZASTD, 6) AND LEFT(ISNULL(ZAIKM_ZAEDD, '999912'), 6)"
'とりあえず台車までの表示にする。※更新は全て行っている
'INSERT 2022/09/30 K.KINEBUCHI start
    strSQL = strSQL & vbCrLf & "AND"
    strSQL = strSQL & vbCrLf & "    ZAIKM_SISYI NOT IN (5,6,99)"
'INSERT 2022/09/30 K.KINEBUCHI end
    strSQL = strSQL & vbCrLf & "AND nm451.NAME_VALUE_FROM = " & Format(intZaikoKbn)
    strSQL = strSQL & vbCrLf & "ORDER BY"
    strSQL = strSQL & vbCrLf & "    nm451.NAME_VALUE_FROM"
    strSQL = strSQL & vbCrLf & ",   ZAIKM_ORDER"
    strSQL = strSQL & vbCrLf & ",   ZAIKM_SISYI"
    strSQL = strSQL & vbCrLf & ",   ZAIKM_SIZE"
    strSQL = strSQL & vbCrLf & ",   ZAIKM_DOOR"
    strSQL = strSQL & vbCrLf & ";"
    
    Select_ZaikoIchiran = strSQL
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Select_ZaikoIchiran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  strat of program ********************************
