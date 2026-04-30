Attribute VB_Name = "Print600"
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　メンテナンス依頼書出力
'   プログラムＩＤ　：　Print600
'   作　成　日　　　：  2007/02/12
'   作　成　者　　　：  イーグルソフト 鈴木
'   Ver             ：  0.0
'   備考            ：
'
'   UPDATE          : 2011/09/22
'   UPDATER         : M.RYU
'   Ver             : 0.1
'   変更内容        : ワークテーブルRKS600_W01にヤード住所を追加、レポートで表示
'
'   UPDATE          : 2013/03/26
'   UPDATER         : M.HONDA
'   Ver             : 0.2
'   変更内容        : カード番号の後ろに部屋番号を追加
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'処理モード
Public Const P600_MODE_PREVIEW              As Integer = 1  '印刷プレビューを表示
Public Const P600_MODE_EXCEL                As Integer = 2  'Excelに出力
Public Const P600_MODE_PRINT                As Integer = 3  'プレビューを表示しないで印刷

'ワークテーブル名
Private Const P_WORK_TABLE                  As String = "RKS600_W01"

'レポート名
Private Const P_REPORT                      As String = "RKS600"

Private strATESAKI                          As String
Private strCONT_KAISYA                      As String
Private strCONT_SCD_NOTE                    As String
Private strCONT_SCD_FAX                     As String

'***************************************
' テストプロ
'***************************************
Sub a00Test_fncPrintMaintenanceRequest()

    If Not fncPrintSecurityCardRequest(P600_MODE_PREVIEW, "502", "2007/01/01", "2007/01/01") Then
'    If Not fncPrintSecurityCardRequest(P600_MODE_PREVIEW, "503", "2008/01/01", "2008/01/01") Then
'    If Not fncPrintSecurityCardRequest(P600_MODE_PREVIEW, "504", "2007/01/01", "2007/01/01") Then
        MsgBox "False"
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : セキュリティーカード依頼書出力
'       MODULE_ID       : fncPrintSecurityCardRequest
'       CREATE_DATE     : 2007/02/12
'                       :
'       PARAM           : intMode        - 1=印刷プレビュー 2=Excel出力 3=印刷（定数宣言あり）
'                       : strYardCode    - ヤードコード
'                       : strStopDayFrom - 停止日ＦＲＯＭ（省略可）
'                       : strStopDayTo   - 停止日ＴＯ（省略可）
'                       : strRoom        - 部屋番号（省略可）
'                       :
'       NOTE            :
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'                       : 不正終了時は例外を発生。
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncPrintSecurityCardRequest(intMode As Integer, _
                                            strYardCode As String, _
                                            Optional strStopDayFrom As String = "", _
                                            Optional strStopDayTo As String = "", _
                                            Optional strRoom As String = "") As Boolean

    Dim dbSQLServer     As Database
    Dim rsGetData       As Recordset
    Dim blnError        As Boolean

On Error GoTo ErrorHandler

    blnError = False

    fncPrintSecurityCardRequest = False

    'DB接続
    Call subConnectServer(dbSQLServer)

    'コントロールマスタデータ取得
    Call subGetControl_Data(dbSQLServer)

    'データ検索
    If Not fncGetData(dbSQLServer, rsGetData, strYardCode, strStopDayFrom, strStopDayTo, strRoom) Then
        '該当データ無し
        GoTo ExitRtn
    End If

    'ワークテーブル作成
    Call subMakeWork(rsGetData, intMode)

    'DB切断
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    '出力
    Select Case intMode
    Case P600_MODE_PREVIEW:
        'レポートプレビュー
        doCmd.OpenReport P_REPORT, acViewPreview
    Case P600_MODE_EXCEL:
        'EXCELファイル出力
        On Error Resume Next
        doCmd.OutputTo acOutputTable, P_WORK_TABLE, acFormatXLS, , True
        On Error GoTo ErrorHandler
    Case P600_MODE_PRINT:
        'レポート印刷
        On Error Resume Next
        doCmd.OpenReport P_REPORT
        On Error GoTo ErrorHandler
    End Select

    fncPrintSecurityCardRequest = True

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    If blnError Then
        Call Err.Raise(Err.Number, "fncPrintSecurityCardRequest" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :subClearWork
'        機能             :ワークテーブルクリア
'        IN               :dbAccess     - ACCESSデータベースオブジェクト(省略可)
'                         :strTableName - テーブル名(省略可)
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subClearWork(Optional dbAccess As Database = Null, _
                         Optional strTable As String = P_WORK_TABLE)

    Dim tdfNew      As TableDef
    Dim blnError    As Boolean
    Dim blnConnect  As Boolean

On Error GoTo ErrorHandler

    blnError = False

    'データベースに未接続ならばCurrentDbに接続する
    If dbAccess Is Nothing Then
        Set dbAccess = CurrentDb
        blnConnect = True
    Else
        blnConnect = False
    End If

    'ワークテーブル削除
    If fncTableExist(dbAccess, strTable) Then
        Call dbAccess.TableDefs.Delete(strTable)
    End If

    'ワークテーブル作成
    Set tdfNew = dbAccess.CreateTableDef(strTable)
    Call subFieldAppend(tdfNew)
    Call dbAccess.TableDefs.Append(tdfNew)

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not tdfNew Is Nothing Then Set tdfNew = Nothing
    If blnConnect And Not dbAccess Is Nothing Then
        dbAccess.Close
        Set dbAccess = Nothing
    End If

    If blnError Then
        Call Err.Raise(Err.Number, "subClearWork" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :コントロールマスタデータ取得
'        MODULE_ID        :subGetControl_Data
'        IN               :dbSqlServer      = DB接続
'        CREATE_DATE      :2007/02/12
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subGetControl_Data(dbSQLServer As Database)

    Dim strSQL  As String
    Dim objRs   As Recordset

    On Error GoTo ErrorHandler

    strSQL = "SELECT " & Chr(13)
    strSQL = strSQL & "CONT_SCD_KAISYA, " & Chr(13)      ' カード管理会社名
    strSQL = strSQL & "CONT_SCD_EIGYOSYO, " & Chr(13)    ' カード管理営業所名
    strSQL = strSQL & "CONT_SCD_KAITANTO, " & Chr(13)    ' カード管理会社担当名
    strSQL = strSQL & "CONT_KAISYA, " & Chr(13)          ' 会社名
    strSQL = strSQL & "CONT_SCD_NOTE, " & Chr(13)        ' カード管理依頼文言
    strSQL = strSQL & "CONT_SCD_FAX " & Chr(13)          ' カード管理会社FAX
    
    strSQL = strSQL & "FROM CONT_MAST " & Chr(13)
    Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)

    With objRs
        If Not .EOF Then
            strATESAKI = .Fields("CONT_SCD_KAISYA") & Space(2) & .Fields("CONT_SCD_EIGYOSYO") & Space(2) & .Fields("CONT_SCD_KAITANTO")
            strCONT_KAISYA = .Fields("CONT_KAISYA")
            strCONT_SCD_NOTE = .Fields("CONT_SCD_NOTE")
            strCONT_SCD_FAX = .Fields("CONT_SCD_FAX")

        Else
            strATESAKI = Null
            strCONT_KAISYA = Null
            strCONT_SCD_NOTE = Null
            strCONT_SCD_FAX = Null
        End If
    End With

subGetControl_Data_Exit:
    'DB切断
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing

    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subGetControl_Data" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    GoTo subGetControl_Data_Exit
End Sub

'==============================================================================*
'
'       MODULE_NAME     : データ検索
'       MODULE_ID       : fncGetData
'       CREATE_DATE     : 2007/02/12
'                       :
'       PARAM           : dbSqlServer    - KOMSに接続したデータベースオブジェクト
'                       : rsGetData      - 検索結果を格納するレコードセット
'                       : strYardCode    - ヤードコード
'                       : strStopDayFrom - 停止日範囲ＦＲＯＭ
'                       : strStopDayTo   - 停止日範囲ＴＯ
'                       : strRoom        - 部屋番号
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetData(dbSQLServer As Database, _
                            ByRef rsGetData As Recordset, _
                            strYardCode As String, _
                            strStopDayFrom As String, _
                            strStopDayTo As String, _
                            strRoom As String) As Boolean

    Dim strSQL      As String

On Error GoTo ErrorHandler

    fncGetData = False

    'SQL文作成
    strSQL = fncMakeGetDataSql(strYardCode, strStopDayFrom, strStopDayTo, strRoom)

    '検索
    Set rsGetData = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)

    'データが存在しない場合Falseを返却
    fncGetData = Not rsGetData.EOF

    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncGetData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : SQL文作成
'       MODULE_ID       : fncMakeGetDataSql
'       CREATE_DATE     : 2007/02/12
'                       :
'       PARAM           : strYardCode    - ヤードコード
'                       : strStopDayFrom - 停止日範囲ＦＲＯＭ
'                       : strStopDayTo   - 停止日範囲ＴＯ
'                       :
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSql(strYardCode As String, _
                                   strStopDayFrom As String, _
                                   strStopDayTo As String, _
                                   strRoom As String) As String

    Dim strSQL              As String

    strSQL = " SELECT SCRD_MAST.SCRDM_YCODE     AS SCRDM_YCODE,     " & Chr(13)     ' ヤードコード
    strSQL = strSQL & " YARD_MAST.YARD_NAME     AS YARD_NAME,       " & Chr(13)     ' ヤード名
    strSQL = strSQL & " SCRD_MAST.SCRDM_ROOM_NO AS SCRDM_ROOM_NO,   " & Chr(13)     ' 部屋番号
    strSQL = strSQL & " YARD_MAST.YARD_ADDR_1 ,                     " & Chr(13)     ' 住所１    'INSERT 2011/09/22 M.RYU
    strSQL = strSQL & " YARD_MAST.YARD_ADDR_2 ,                     " & Chr(13)     ' 住所２    'INSERT 2011/09/22 M.RYU
    strSQL = strSQL & " YARD_MAST.YARD_ADDR_3 ,                     " & Chr(13)     ' 住所３    'INSERT 2011/09/22 M.RYU
    strSQL = strSQL & " SCRD_MAST.SCRDM_CARD_NO AS SCRDM_CARD_NO    " & Chr(13)     ' カード番号
    strSQL = strSQL & " FROM SCRD_MAST, " & Chr(13)
    strSQL = strSQL & "      YARD_MAST " & Chr(13)
    strSQL = strSQL & " WHERE SCRD_MAST.SCRDM_YCODE = YARD_MAST.YARD_CODE " & Chr(13)
    strSQL = strSQL & "   AND SCRD_MAST.SCRDM_YCODE =  '" & strYardCode & "' " & Chr(13)

    '停止日の範囲条件
    strSQL = strSQL & fncMakeBetween("SCRD_MAST.SCRDM_STOP_DAY", strStopDayFrom, strStopDayTo)

    ' 部屋番号が指定された場合、その部屋と部屋未割当(-1)のものを対象とする
    If strRoom <> "" Then
        strSQL = strSQL & " AND ( SCRD_MAST.SCRDM_ROOM_NO = '" & strRoom & "' "
        strSQL = strSQL & "    OR SCRD_MAST.SCRDM_ROOM_NO = -1) " & Chr(13)
    End If

    'ソート句
    strSQL = strSQL & " ORDER BY SCRD_MAST.SCRDM_CARD_NO "

    fncMakeGetDataSql = strSQL

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
'        MODULE_NAME      :fncTableExist
'        機能             :ACCESSテーブル存在チェック
'        IN               :dbAccess     - ACCESSデータベースオブジェクト
'                         :strTableName - テーブル名
'        OUT              :True=存在する False=存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncTableExist(dbAccess As Database, strTableName As String) As Boolean
    
    Dim tdf As TableDef

    fncTableExist = False

    For Each tdf In dbAccess.TableDefs
        If tdf.NAME = strTableName Then
            fncTableExist = True
            Exit For
        End If
    Next tdf
End Function

'==============================================================================*
'
'        MODULE_NAME      :subFieldAppend
'        機能             :ワークテーブル列作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFieldAppend(tdfNew As TableDef)

    Dim fldNew      As Field
    Dim intCount    As Integer

    With tdfNew
 
        Call .Fields.Append(.CreateField("宛先", DataTypeEnum.dbText, 255))                 '宛先
        Call .Fields.Append(.CreateField("会社名", DataTypeEnum.dbText, 30))                '会社名
        Call .Fields.Append(.CreateField("ヤードコード", DataTypeEnum.dbText, 6))           'ヤードコード
        Call .Fields.Append(.CreateField("ヤード名", DataTypeEnum.dbText, 36))              'ヤード名
        Call .Fields.Append(.CreateField("部屋番号", DataTypeEnum.dbText, 100))             '部屋番号
        Call .Fields.Append(.CreateField("ヤード住所", DataTypeEnum.dbText, 100))            'ヤード住所     'INSERT 2011/09/22 M.RYU
        Call .Fields.Append(.CreateField("カード番号01", DataTypeEnum.dbText, 50))          'カード番号
        Call .Fields.Append(.CreateField("カード番号02", DataTypeEnum.dbText, 50))          'カード番号
        Call .Fields.Append(.CreateField("カード番号03", DataTypeEnum.dbText, 50))          'カード番号
        Call .Fields.Append(.CreateField("カード番号04", DataTypeEnum.dbText, 50))          'カード番号
        Call .Fields.Append(.CreateField("カード番号05", DataTypeEnum.dbText, 50))          'カード番号
        Call .Fields.Append(.CreateField("カード番号06", DataTypeEnum.dbText, 50))          'カード番号
        Call .Fields.Append(.CreateField("カード番号07", DataTypeEnum.dbText, 50))          'カード番号
        Call .Fields.Append(.CreateField("カード番号08", DataTypeEnum.dbText, 50))          'カード番号
        Call .Fields.Append(.CreateField("カード番号09", DataTypeEnum.dbText, 50))          'カード番号
        Call .Fields.Append(.CreateField("カード番号10", DataTypeEnum.dbText, 50))          'カード番号
        Call .Fields.Append(.CreateField("カード管理依頼文言", DataTypeEnum.dbText, 60))    'カード管理依頼文言
        Call .Fields.Append(.CreateField("カード管理会社FAX", DataTypeEnum.dbText, 15))     'カード管理会社FAX
        Call .Fields.Append(.CreateField("カード枚数", DataTypeEnum.dbInteger, 2))          'カード枚数
        

        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount

    End With

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subMakeWork
'        機能             :ワークテーブルデータ追加
'        IN               :rsSource    - 検索結果が格納されたレコードセット
'                         :intMode     - 1=印刷プレビュー 2=Excel出力 3=印刷（定数宣言あり）
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subMakeWork(rsSource As Recordset, intMode As Integer)

    Dim dbAccess        As Database
    Dim rsDestination   As Recordset
    Dim blnError        As Boolean

    Dim intPlace        As Integer

On Error GoTo ErrorHandler

    blnError = False

    Set dbAccess = CurrentDb

    'ワークテーブルクリア
    Call subClearWork(dbAccess, P_WORK_TABLE)

    'ワークテーブルのレコードセット
    Set rsDestination = dbAccess.OpenRecordset(P_WORK_TABLE, dbOpenDynaset, dbAppendOnly)

    intPlace = 0

    'データ追加
    With rsSource
        While Not rsSource.EOF

            If intPlace > 10 Then
                ' 明細数が10件を超えた場合
                rsDestination.UPDATE

                intPlace = 0
            End If

            If intPlace = 0 Then
                ' NEWレコード追加
                rsDestination.AddNew

                ' 先頭ポジションに移動
                intPlace = 1

                rsDestination.Fields("宛先") = strATESAKI                                              '宛先              ＝コントロールマスタ．カード管理会社名 & コントロールマスタ．カード管理営業所名 & コントロールマスタ．カード管理会社担当名
                rsDestination.Fields("会社名") = strCONT_KAISYA                                        '会社名            ＝コントロールマスタ．会社名
                rsDestination.Fields("ヤードコード") = Format(.Fields("SCRDM_YCODE"), "000000")        'ヤードコード      ＝セキュリティーカードマスタ．ヤードコード
                rsDestination.Fields("ヤード名") = Format(.Fields("SCRDM_YCODE"), "000000") & " : " & .Fields("YARD_NAME")                                'ヤード名          ＝ヤードマスタ．ヤードコード
                rsDestination.Fields("部屋番号") = .Fields("SCRDM_ROOM_NO")                            '部屋番号          ＝セキュリティーカードマスタ．部屋番号
                
                'INSERT 2011/09/22 M.RYU
                rsDestination.Fields("ヤード住所") = .Fields("YARD_ADDR_1") & "　" & _
                                        .Fields("YARD_ADDR_2") & "　" & .Fields("YARD_ADDR_3")         'ヤード住所        ＝ヤードマスタ．+住所１+住所２+住所３
                                        
                rsDestination.Fields("カード管理依頼文言") = strCONT_SCD_NOTE                          'カード管理依頼文言＝コントロールマスタ．カード管理依頼文言
                rsDestination.Fields("カード管理会社FAX") = strCONT_SCD_FAX                            'カード管理会社FAX ＝コントロールマスタ．カード管理会社FAX

            Else
            
            End If

            Select Case intPlace
                Case 1
                    rsDestination.Fields("カード番号01") = .Fields("SCRDM_CARD_NO") & " / " & .Fields("SCRDM_ROOM_NO")  'カード番号        ＝コントロールマスタ．ヤードコード

                Case 2
                    rsDestination.Fields("カード番号02") = .Fields("SCRDM_CARD_NO") & " / " & .Fields("SCRDM_ROOM_NO")  'カード番号        ＝コントロールマスタ．ヤードコード

                Case 3
                    rsDestination.Fields("カード番号03") = .Fields("SCRDM_CARD_NO") & " / " & .Fields("SCRDM_ROOM_NO")  'カード番号        ＝コントロールマスタ．ヤードコード

                Case 4
                    rsDestination.Fields("カード番号04") = .Fields("SCRDM_CARD_NO") & " / " & .Fields("SCRDM_ROOM_NO")  'カード番号        ＝コントロールマスタ．ヤードコード

                Case 5
                    rsDestination.Fields("カード番号05") = .Fields("SCRDM_CARD_NO") & " / " & .Fields("SCRDM_ROOM_NO")  'カード番号        ＝コントロールマスタ．ヤードコード

                Case 6
                    rsDestination.Fields("カード番号06") = .Fields("SCRDM_CARD_NO") & " / " & .Fields("SCRDM_ROOM_NO")  'カード番号        ＝コントロールマスタ．ヤードコード

                Case 7
                    rsDestination.Fields("カード番号07") = .Fields("SCRDM_CARD_NO") & " / " & .Fields("SCRDM_ROOM_NO")  'カード番号        ＝コントロールマスタ．ヤードコード

                Case 8
                    rsDestination.Fields("カード番号08") = .Fields("SCRDM_CARD_NO") & " / " & .Fields("SCRDM_ROOM_NO")  'カード番号        ＝コントロールマスタ．ヤードコード

                Case 9
                    rsDestination.Fields("カード番号09") = .Fields("SCRDM_CARD_NO") & " / " & .Fields("SCRDM_ROOM_NO")  'カード番号        ＝コントロールマスタ．ヤードコード

                Case 10
                    rsDestination.Fields("カード番号10") = .Fields("SCRDM_CARD_NO") & " / " & .Fields("SCRDM_ROOM_NO")  'カード番号        ＝コントロールマスタ．ヤードコード
            End Select

            rsDestination.Fields("カード枚数") = intPlace

            ' ポジション移動
            intPlace = intPlace + 1

            .MoveNext
        Wend

        ' 更新要の場合の処理
        rsDestination.UPDATE
    End With

    GoTo EndRtn

ErrorHandler:
    blnError = True

EndRtn:
    If Not rsDestination Is Nothing Then rsDestination.Close: Set rsDestination = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    
    If blnError Then
        Call Err.Raise(Err.Number, "subMakeWork" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : KOMSデータベース接続
'       MODULE_ID       : subConnectServer
'       CREATE_DATE     :
'       PARAM           : データベースオブジェクト
'       RETURN          :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subConnectServer(ByRef dbSQLServer As Database)

    Dim strSqlserver    As String
    Dim strConnect      As String
    Dim strBUMOC        As String
    
On Error GoTo ErrorHandler

    '部門コード
    strBUMOC = fncGetBumonCode()

    'SQL-Server名
    strSqlserver = fncGetSqlServer(strBUMOC)

    '接続文字列取得
    strConnect = fncGetConnectString(strBUMOC)

    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)

    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subConnectServer" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 部門コード取得
'       MODULE_ID       : fncGetBumonCode
'       CREATE_DATE     :
'       PARAM           :
'       RETURN          : 部門コード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetBumonCode() As String

    Dim strBumonCode        As String

On Error GoTo ErrorHandler

    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    fncGetBumonCode = strBumonCode

    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncGetBumonCode" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

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

    Dim strSqlserver    As String
    Dim strParam        As String

On Error GoTo ErrorHandler

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

    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncGetSqlServer" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
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

    Dim strConnectString    As String

On Error GoTo ErrorHandler

    strConnectString = MSZZ007_M00(strBumonCode)
    If strConnectString = "" Then
        'テーブル[SETU_TABL]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "接続文字列の設定不正。")
    End If

    fncGetConnectString = strConnectString

    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncGetConnectString" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
