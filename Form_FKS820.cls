VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS820"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ボックスマスタメンテ
'   プログラムＩＤ　：　FKS820
'   作　成　日　　　：  2008/06/30
'   作　成　者　　　：  N.IMAI
'   Ver             ：  0.0

'   更　新　日　　　：  2009/10/02
'   更　新　者　　　：  KIHARA
'   Ver             ：  0.1
'                   ：  ボックス明細を追加
'
'   更　新　日　　　：  2009/10/07
'   更　新　者　　　：  KIHARA
'   Ver             ：  0.2
'                   ：  ボックス明細のクエリにorderbyを追加
'
'   更　新　日　　　：  2013/06/03
'   更　新　者　　　：  M.HONDA
'   Ver             ：  0.3
'                   ：  終了メッセージを表示させないように修正

'**********************************************

Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_QUERY_NAME As String = "FKS820_Q01"         'クエリ名
Private Const P_QUERY_NAME2 As String = "FKS820_Q02"        'クエリ名(ボックス明細）         'ADD 2009/10/02 KIHARA
Private Const P_ROOM_WORK_TABLE As String = "WK_FKS820"     'ワークテーブル名
Private Const P_ROOM_WORK_TABLE2 As String = "WK_FKS820_01" 'ワークテーブル名(ボックス明細） 'ADD 2009/10/02 KIHARA

Private pobjAdoDbConnection         As Object               'SQLサーバー
Private pobjDB                      As Database             'ローカルDB
Private pobjTableDef                As DAO.TableDef         'ワークテーブル
Private pobjQueryDef                As DAO.QueryDef         'クエリ
Private pstrBumonCode               As String               '部門コード

'カラム情報
Private Type ColumnsDef
    HeadderName                 As String
    SIZE                        As Double
    Door                        As Integer
    TotalCount                  As Long
End Type

'==============================================================================*
'
'        MODULE_NAME      :フォームロード時処理
'        MODULE_ID        :Form_Load
'        NOTE             :
'        CREATE_DATE      :2008/06/30
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

On Error GoTo Exception

    pstrBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If pstrBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    Set pobjDB = CurrentDb
    
    Call subDropWorkTable
    Call subDropQuery
    
Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("Form_Load", True)
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォームアクティブ処理
'        MODULE_ID        :Form_Activate
'        NOTE             :
'        CREATE_DATE      :2008/06/30
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()

    doCmd.Maximize
    doCmd.Requery
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :終了ボタンクリック処理
'        MODULE_ID        :cmd_終了_Click
'        NOTE             :
'        CREATE_DATE      :2008/06/30
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()

On Error GoTo Exception
    '’2013/06/03 M.HONDA START
'    If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo, "確認") = vbYes Then
'        Call subDropWorkTable
'        Call subDropQuery
'        doCmd.Close acForm, Me.FormName
'    End If
    Call subDropWorkTable
    Call subDropQuery
    doCmd.Close acForm, Me.FormName

    '’2013/06/03 M.HONDA END

Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("cmd_終了_Click", True)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Close
'        MODULE_ID        :Form_Close
'        NOTE             :
'        CREATE_DATE      :2008/06/30
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()

On Error GoTo Exception
    
    'コネクションクローズ
    If Not pobjAdoDbConnection Is Nothing Then
       pobjAdoDbConnection.Close
       Set pobjAdoDbConnection = Nothing
    End If
    If Not pobjDB Is Nothing Then
       pobjDB.Close
       Set pobjDB = Nothing
    End If
    
    Call subDropWorkTable
    Call subDropQuery
    
Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("Form_Close", True)
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ヤードコード(From)入力後時処理
'        MODULE_ID        :cmb_ヤードコード_From_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/30
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤードコード_From_AfterUpdate()

On Error GoTo Exception

    If Me.cmb_ヤードコード_From.ListIndex = -1 Then
        Me.txt_ヤード名_From = ""
    Else
        Me.txt_ヤード名_From = Me.cmb_ヤードコード_From.Column(1, Me.cmb_ヤードコード_From.ListIndex)
    End If
    
    Me.cmb_ヤードコード_From = Format(Nz(Me.cmb_ヤードコード_From), "000000")

Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("cmb_ヤード_AfterUpdate", True)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ヤードコード(To)入力後時処理
'        MODULE_ID        :cmb_ヤードコード_From_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/30
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤードコード_To_AfterUpdate()

On Error GoTo Exception

    If Me.cmb_ヤードコード_To.ListIndex = -1 Then
        Me.txt_ヤード名_To = ""
    Else
        Me.txt_ヤード名_To = Me.cmb_ヤードコード_To.Column(1, Me.cmb_ヤードコード_To.ListIndex)
    End If
    
    Me.cmb_ヤードコード_To = Format(Nz(Me.cmb_ヤードコード_To), "000000")

Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("cmb_ヤード_AfterUpdate", True)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :EXCELボタンクリック処理
'        MODULE_ID        :cmd_EXCEL_Click
'        NOTE             :
'        CREATE_DATE      :2008/06/30
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()

    On Error GoTo Exception

    '入力チェック
    'ADD 2009/10/02 KIHARA Start
    If Nz(Me.cmb_Excel出力パターン.VALUE) = "" Then
        MsgBox "Excel出力パターンを選択して下さい。", vbInformation, "確認"
        Me.cmb_Excel出力パターン.SetFocus
        Exit Sub
    End If
    'ADD 2009/10/02 KIHARA End

    If Trim(Nz(Me.cmb_ヤードコード_To.VALUE)) <> "" _
    And (Nz(Me.cmb_ヤードコード_From.VALUE) > Nz(Me.cmb_ヤードコード_To.VALUE)) Then
        MsgBox "ヤードコードの指定に誤りがあります。", vbInformation, "確認"
        Me.cmb_ヤードコード_From.SetFocus
        Exit Sub
    End If
        
    If vbYes = MsgBox("印刷内容をEXCELに出力します。よろしいですか？", vbQuestion + vbYesNo, "確認") Then
        'DEL 2009/10/02 KIHARA Start
'        Call subOutputYardData
'        'Excel起動
'        doCmd.OutputTo acQuery, P_QUERY_NAME, "MicrosoftExcel(*.xls)", "", True, ""
        'DEL 2009/10/02 KIHARA End
        
        'ADD 2009/10/02 KIHARA Start
        Select Case Me.cmb_Excel出力パターン.VALUE
        Case 0 'ボックスサイズ別
            Call subOutputYardData
            'Excel起動
            doCmd.OutputTo acQuery, P_QUERY_NAME, "MicrosoftExcel(*.xls)", "", True, ""
        Case 1 'ボックス明細
            Call subOutputYardData2
            'Excel起動
            doCmd.OutputTo acQuery, P_QUERY_NAME2, "MicrosoftExcel(*.xls)", "", True, ""
        End Select
        'ADD 2009/10/02 KIHARA End
    End If

Exception:

    If Err <> 0 And Err <> 2501 Then
        Call MSZZ024_M00("cmd_EXCEL_Click", True)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル削除
'        MODULE_ID        :subDropWorkTable
'        NOTE             :
'        CREATE_DATE      :2008/06/30
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDropWorkTable()

On Error GoTo Exception

'    doCmd.DeleteObject acTable, P_ROOM_WORK_TABLE       'DEL 2009/10/02 KIHARA

    'ADD 2009/10/02 KIHARA Start
    Select Case Me.cmb_Excel出力パターン.VALUE
    Case 0 'ボックスサイズ別
        doCmd.DeleteObject acTable, P_ROOM_WORK_TABLE
    Case 1 'ボックス明細
        doCmd.DeleteObject acTable, P_ROOM_WORK_TABLE2
    End Select
    'ADD 2009/10/02 KIHARA End
Exception:
    Err.Clear
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :クエリ作成
'        MODULE_ID        :subCreateQuery
'        NOTE             :
'        CREATE_DATE      :2008/06/30
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCreateQuery()

On Error GoTo Exception

    Dim strSQL As String    'ADD 2009/10/02 KIHARA
    
'    Set pobjQueryDef = pobjDb.CreateQueryDef(P_QUERY_NAME)     'DEL 2009/10/02 KIHARA
'    pobjQueryDef.sql = "SELECT * FROM " & P_ROOM_WORK_TABLE    'DEL 2009/10/02 KIHARA

    'ADD 2009/10/02 KIHARA Start
    Select Case Me.cmb_Excel出力パターン.VALUE
    Case 0 'ボックスサイズ別
        Set pobjQueryDef = pobjDB.CreateQueryDef(P_QUERY_NAME)
        pobjQueryDef.sql = "SELECT * FROM " & P_ROOM_WORK_TABLE
    Case 1 'ボックス明細
        Set pobjQueryDef = pobjDB.CreateQueryDef(P_QUERY_NAME2)
        'count列以外をselectする
        strSQL = "SELECT ヤードコード, ヤード名,ボックス番号,タイプ,内装,サイズ"
        strSQL = strSQL & ",ドア,入庫日,設置日,階層,状況,割付部屋番号,備考"
        strSQL = strSQL & " FROM " & P_ROOM_WORK_TABLE2
        strSQL = strSQL & " ORDER BY ヤードコード,ボックス番号 "    'ADD 2009/10/07 KIHARA
        pobjQueryDef.sql = strSQL
    End Select
    'ADD 2009/10/02 KIHARA End
    
Exception:
    Err.Clear
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :クエリ削除
'        MODULE_ID        :subDropQuery
'        NOTE             :
'        CREATE_DATE      :2008/06/30
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDropQuery()

On Error GoTo Exception

'    doCmd.DeleteObject acQuery, P_QUERY_NAME       'DEL 2009/10/02 KIHARA

    'ADD 2009/10/02 KIHARA Start
    Select Case Me.cmb_Excel出力パターン.VALUE
    Case 0 'ボックスサイズ別
        doCmd.DeleteObject acQuery, P_QUERY_NAME
    Case 1 'ボックス明細
        doCmd.DeleteObject acQuery, P_QUERY_NAME2
    End Select
    'ADD 2009/10/02 KIHARA End

Exception:
    Err.Clear
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :EXCELに出力するデータをワークテーブルへ作成
'        MODULE_ID        :subOutputYardData
'        NOTE             :
'        CREATE_DATE      :2008/06/30
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subOutputYardData()

    Dim strSQL          As String
    Dim strWhere        As String
    Dim rsDataYard      As Object
    Dim rsDataCount     As Object
    Dim rsDataSize      As Object
    Dim rsDataDoor      As Object
    Dim rsDataWork      As Object
    Dim WorkColumns()   As ColumnsDef
    Dim intSizeCount    As Integer
    Dim intDoorCount    As Integer
    Dim intCount        As Integer
    Dim i               As Integer

On Error GoTo Exception

    '初期値設定
    intSizeCount = 1
    intDoorCount = 1
    intCount = 0
    
    'コネクション接続
    Set pobjAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)

    'クエリ作成
    Call subCreateQuery
    
    'ワークテーブルに必要なカラム（サイズ×ドア数）を取得
    'サイズ
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM NAME_MAST WHERE NAME_ID = '271' ORDER BY NAME_CODE"
    Set rsDataSize = MSZZ025.ADODB_Recordset(strSQL, pobjAdoDbConnection, adoReadWrite)
    'ドア数
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM NAME_MAST WHERE NAME_ID = '272' ORDER BY NAME_CODE"
    Set rsDataDoor = MSZZ025.ADODB_Recordset(strSQL, pobjAdoDbConnection, adoReadWrite)
    While Not rsDataSize.EOF
        intDoorCount = 1
        While Not rsDataDoor.EOF
            'カラム設定保存
            ReDim Preserve WorkColumns(intCount)
            WorkColumns(intCount).HeadderName = rsDataSize.Fields("NAME_NAME") & " " & rsDataDoor.Fields("NAME_NAME")
            WorkColumns(intCount).SIZE = rsDataSize.Fields("NAME_CODE")
            WorkColumns(intCount).Door = rsDataDoor.Fields("NAME_CODE")
            WorkColumns(intCount).TotalCount = 0
            intCount = intCount + 1
            intDoorCount = intDoorCount + 1
            rsDataDoor.MoveNext
        Wend
        intSizeCount = intSizeCount + 1
        rsDataSize.MoveNext
        '先頭に戻る
        rsDataDoor.MoveFirst
    Wend
    
    'ワークテーブル作成
    Call subDropWorkTable
    Set pobjTableDef = pobjDB.CreateTableDef(P_ROOM_WORK_TABLE)
    Call pobjTableDef.Fields.Append(pobjTableDef.CreateField("ヤードコード", dbLong, 6))
    pobjTableDef.Fields.Append pobjTableDef.CreateField("ヤード名", dbText, 36)
    For i = 0 To UBound(WorkColumns)
        pobjTableDef.Fields.Append pobjTableDef.CreateField(WorkColumns(i).HeadderName, dbInteger, 10)
    Next i
    pobjTableDef.Fields.Append pobjTableDef.CreateField("ヤード計", dbInteger, 10)
    pobjDB.TableDefs.Append pobjTableDef

    'ワークテーブルのカラム単位でボックスマスタの件数を取得
    strSQL = ""
    strSQL = strSQL & "SELECT YARD_MAST.YARD_CODE"
    strSQL = strSQL & "      ,MAX(YARD_MAST.YARD_NAME) AS YARD_NAME "
    strSQL = strSQL & "      ,CBOX_MAST.CBOXM_SIZE "
    strSQL = strSQL & "      ,CBOX_MAST.CBOXM_DOOR "
    strSQL = strSQL & "      ,CBOX_MAST.CBOXM_DOORPSTN"
    strSQL = strSQL & "      ,CBOX_MAST.CBOXM_DOORTYPE"
    strSQL = strSQL & "      ,COUNT(*) AS DATA_COUNT "
    strSQL = strSQL & " FROM YARD_MAST, CBOX_MAST, CBYS_TRAN "
    '検索条件
    If Trim(Nz(Me.cmb_ヤードコード_From.VALUE)) <> "" Then
        strSQL = strSQL & " WHERE YARD_MAST.YARD_CODE >= " & Me.cmb_ヤードコード_From.VALUE
        If Trim(Nz(Me.cmb_ヤードコード_To.VALUE)) <> "" Then
            strSQL = strSQL & " AND YARD_MAST.YARD_CODE <= " & Me.cmb_ヤードコード_To.VALUE
        End If
    Else
        If Trim(Nz(Me.cmb_ヤードコード_To.VALUE)) <> "" Then
            strSQL = strSQL & " WHERE YARD_MAST.YARD_CODE <= " & Me.cmb_ヤードコード_To.VALUE
        Else
            strSQL = strSQL & " WHERE YARD_MAST.YARD_CODE >= 0"
        End If
    End If
    If Trim(Nz(Me.cmb_タイプ.VALUE)) <> "" Then
        strSQL = strSQL & " AND CBOXM_TYPE = " & Me.cmb_タイプ.Column(1, Me.cmb_タイプ.ListIndex)
    End If
    If Trim(Nz(Me.cmb_ドア数.VALUE)) <> "" Then
        strSQL = strSQL & " AND CBOXM_DOOR = " & Me.cmb_ドア数.Column(1, Me.cmb_ドア数.ListIndex)
    End If
    If Trim(Nz(Me.cmb_内装.VALUE)) <> "" Then
        strSQL = strSQL & " AND CBOXM_INTYPE = " & Me.cmb_内装.Column(1, Me.cmb_内装.ListIndex)
    End If
    If Trim(Nz(Me.cmb_ドア位置.VALUE)) <> "" Then
        strSQL = strSQL & " AND CBOXM_DOORPSTN = " & Me.cmb_ドア位置.Column(1, Me.cmb_ドア位置.ListIndex)
    End If
    If Trim(Nz(Me.cmb_サイズ.VALUE)) <> "" Then
        strSQL = strSQL & " AND CBOXM_SIZE = " & Me.cmb_サイズ.Column(2, Me.cmb_サイズ.ListIndex)
    End If
    If Trim(Nz(Me.cmb_ドアタイプ.VALUE)) <> "" Then
        strSQL = strSQL & " AND CBOXM_DOORTYPE = " & Me.cmb_ドアタイプ.Column(1, Me.cmb_ドアタイプ.ListIndex)
    End If
    If Trim(Nz(Me.cmb_設置階層.VALUE)) <> "" Then
        strSQL = strSQL & " AND CBYST_SETFLOOR = " & Me.cmb_設置階層.Column(1, cmb_設置階層.ListIndex)
    End If
    strSQL = strSQL & " AND   CBYS_TRAN.CBYST_YCODE = YARD_MAST.YARD_CODE"
    strSQL = strSQL & " AND   CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & " GROUP BY YARD_MAST.YARD_CODE, CBOX_MAST.CBOXM_SIZE, CBOX_MAST.CBOXM_DOOR,CBOX_MAST.CBOXM_DOORPSTN,CBOX_MAST.CBOXM_DOORTYPE "
    strSQL = strSQL & " ORDER BY YARD_MAST.YARD_CODE, CBOX_MAST.CBOXM_SIZE, CBOX_MAST.CBOXM_DOOR,CBOX_MAST.CBOXM_DOORPSTN,CBOX_MAST.CBOXM_DOORTYPE "
    Set rsDataCount = MSZZ025.ADODB_Recordset(strSQL, pobjAdoDbConnection, adoReadWrite)
    While Not rsDataCount.EOF
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM " & P_ROOM_WORK_TABLE
        strSQL = strSQL & " WHERE ヤードコード = " & rsDataCount.Fields("YARD_CODE")
        Set rsDataWork = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
        If rsDataWork.EOF Then
            rsDataWork.AddNew
            '初期値設定
            For i = 2 To rsDataWork.Fields.Count - 1
                rsDataWork.Fields(i) = 0
            Next i
        Else
            rsDataWork.Edit
        End If
        rsDataWork.Fields("ヤードコード") = rsDataCount.Fields("YARD_CODE")
        rsDataWork.Fields("ヤード名") = rsDataCount.Fields("YARD_NAME")
        For i = 0 To UBound(WorkColumns)
            If WorkColumns(i).SIZE = rsDataCount.Fields("CBOXM_SIZE") _
            And WorkColumns(i).Door = rsDataCount.Fields("CBOXM_DOOR") Then
                '該当カラムに加算
                rsDataWork.Fields(i + 2) = rsDataWork.Fields(i + 2) + rsDataCount.Fields("DATA_COUNT")
                'ヤード計算出（ヨコ計）
                rsDataWork.Fields("ヤード計") = rsDataWork.Fields("ヤード計") + rsDataCount.Fields("DATA_COUNT")
                'カラム計算出（タテ計）
                WorkColumns(i).TotalCount = WorkColumns(i).TotalCount + rsDataCount.Fields("DATA_COUNT")
                Exit For
            End If
        Next i
        rsDataWork.UPDATE
        rsDataWork.Close
        Set rsDataWork = Nothing
        rsDataCount.MoveNext
    Wend

    '総合計行追加
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM " & P_ROOM_WORK_TABLE
    strSQL = strSQL & " WHERE ヤードコード = 0"
    Set rsDataWork = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
    rsDataWork.AddNew
    rsDataWork.Fields("ヤード名") = "総合計"
    rsDataWork.Fields("ヤード計") = 0
    For i = 0 To UBound(WorkColumns)
        rsDataWork.Fields(i + 2) = WorkColumns(i).TotalCount
        rsDataWork.Fields("ヤード計") = rsDataWork.Fields("ヤード計") + WorkColumns(i).TotalCount
    Next i
    rsDataWork.UPDATE
    rsDataWork.Close
    Set rsDataWork = Nothing

Exception:

    If Not rsDataYard Is Nothing Then
        rsDataYard.Close
        Set rsDataYard = Nothing
    End If
    If Not rsDataCount Is Nothing Then
        rsDataCount.Close
        Set rsDataCount = Nothing
    End If
    If Not rsDataSize Is Nothing Then
        rsDataSize.Close
        Set rsDataSize = Nothing
    End If
    If Not rsDataDoor Is Nothing Then
        rsDataDoor.Close
        Set rsDataDoor = Nothing
    End If
    If Not rsDataWork Is Nothing Then
        rsDataWork.Close
        Set rsDataWork = Nothing
    End If
    If Not pobjAdoDbConnection Is Nothing Then
        'コネクション解放
        pobjAdoDbConnection.Close
        Set pobjAdoDbConnection = Nothing
    End If
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "subOutputExcel" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    

End Sub

'==============================================================================*
'
'        MODULE_NAME      :EXCELに出力するデータをワークテーブルへ作成(ボックス明細）
'        MODULE_ID        :subOutputYardData2
'        NOTE             :
'        CREATE_DATE      :2009/10/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subOutputYardData2()

On Error GoTo Exception
   
    'コネクション接続
    Set pobjAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)

    'クエリ作成
    Call subCreateQuery

    'ワークテーブル削除
    Call subDropWorkTable
    
    'ワークテーブル作成
    Set pobjTableDef = pobjDB.CreateTableDef(P_ROOM_WORK_TABLE2)
    Call FieldAppend
    pobjDB.TableDefs.Append pobjTableDef
    
    'ワークテーブル更新
    Call getSelectBoxData(createBoxDataSQL)
    
Exception:
    If Not pobjAdoDbConnection Is Nothing Then
        'コネクション解放
        pobjAdoDbConnection.Close
        Set pobjAdoDbConnection = Nothing
    End If
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "subOutputYardData2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル列作成
'        MODULE_ID        :FieldAppend
'        NOTE             :
'        CREATE_DATE      :2009/10/02
'        UPDATE_DATE      :

'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FieldAppend()

    Dim intCount    As Integer

    With pobjTableDef
        Call .Fields.Append(.CreateField("ヤードコード", dbText, 36))
        Call .Fields.Append(.CreateField("ヤード名", dbText, 50))
        Call .Fields.Append(.CreateField("ボックス番号", dbText, 36))
        Call .Fields.Append(.CreateField("タイプ", dbText, 6))
        Call .Fields.Append(.CreateField("内装", dbText, 6))
        Call .Fields.Append(.CreateField("サイズ", dbText, 6))
        Call .Fields.Append(.CreateField("ドア", dbText, 50))
        Call .Fields.Append(.CreateField("入庫日", dbText, 10))
        Call .Fields.Append(.CreateField("設置日temp", dbText, 8))  '設置日(yyyymmdd）
        Call .Fields.Append(.CreateField("設置日", dbText, 10))     '設置日(yyyy/mm/dd)
        Call .Fields.Append(.CreateField("階層", dbText, 6))
        Call .Fields.Append(.CreateField("状況", dbText, 50))
        Call .Fields.Append(.CreateField("割付部屋番号", dbText, 50))
        Call .Fields.Append(.CreateField("備考", dbText, 50))
        Call .Fields.Append(.CreateField("count", dbLong))  'カウンタ（割付部屋番号数をカウントする）
        
        '空文字列の登録を許可する
        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount

    End With
    
End Sub
    

'==============================================================================*
'
'        MODULE_NAME      :一覧表時用ワークテーブル更新
'        MODULE_ID        :getSelectBoxData
'        NOTE             :
'        CREATE_DATE      :2009/10/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getSelectBoxData(strSQL As String)
    
    Dim rsData               As Object
    Dim rsBoxData            As Recordset
    Dim sSQL                 As String
    
    On Error GoTo Exception
    
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, pobjAdoDbConnection, adoReadWrite)

    While Not rsData.EOF
        'ワークテーブルの存在チェック
        sSQL = ""
        sSQL = sSQL & "SELECT * FROM " & P_ROOM_WORK_TABLE2
        sSQL = sSQL & " WHERE ボックス番号 = '" & rsData.Fields("BOXNO") & "'"
        sSQL = sSQL & " AND 設置日temp = '" & rsData.Fields("SETDATE") & "'"   '設置日temp(テキスト型）
        Set rsBoxData = pobjDB.OpenRecordset(sSQL, dbOpenDynaset)
        With rsBoxData
            If .EOF Then
                .AddNew
                .Fields("count") = 1
                .Fields("ヤードコード") = Format(Nz(rsData.Fields("YARD_CODE")), "000000")
                .Fields("ヤード名") = rsData.Fields("YARD_NAME")
                .Fields("ボックス番号") = rsData.Fields("BOXNO")
                .Fields("タイプ") = rsData.Fields("TYPE")
                .Fields("内装") = rsData.Fields("INTYPE")
                .Fields("サイズ") = rsData.Fields("SIZE")
                .Fields("ドア") = rsData.Fields("DOOR") & "," & rsData.Fields("DOORPSTN") & "," & rsData.Fields("DOORTYPE")
                .Fields("入庫日") = Left(rsData.Fields("STOCKDATE"), 4) & "/" & Mid(rsData.Fields("STOCKDATE"), 5, 2) & "/" & _
                                    Right(rsData.Fields("STOCKDATE"), 2)
                .Fields("状況") = "設置中"
                If Nz(rsData.Fields("CNO"), "") = "" Then
                    .Fields("割付部屋番号") = "（未割付）"
                Else
                    .Fields("割付部屋番号") = Format(Nz(rsData.Fields("CNO")), "000000")
                End If
                .Fields("設置日temp") = rsData.Fields("SETDATE")  '設置日temp(yyyymmdd)
                .Fields("設置日") = Left(rsData.Fields("SETDATE"), 4) & "/" & Mid(rsData.Fields("SETDATE"), 5, 2) & "/" & _
                                    Right(rsData.Fields("SETDATE"), 2)
                .Fields("階層") = rsData.Fields("SETFLOOR")
                .Fields("備考") = rsData.Fields("BIKON")
                .UPDATE
            Else
                Select Case .Fields("count")  'カウンタ
                Case Is < 5
                    '4つ目までの割付部屋番号は、カンマ区切りで列挙する
                    .Edit
                    .Fields("割付部屋番号") = .Fields("割付部屋番号") & "," & Format(Nz(rsData.Fields("CNO")), "000000")
                    .Fields("count") = .Fields("count") + 1
                    .UPDATE
                Case 5
                    '5つ目の割付部屋番号は、カンマ区切りで列挙し、その番号の後に「他」と入力する
                    .Edit
                    .Fields("割付部屋番号") = .Fields("割付部屋番号") & "他"
                    .Fields("count") = .Fields("count") + 1
                    .UPDATE
                End Select
            End If
            .Close
            Set rsBoxData = Nothing
            rsData.MoveNext
        End With
    Wend

    rsData.Close
    Set rsData = Nothing
    pobjAdoDbConnection.Close
    Set pobjAdoDbConnection = Nothing
   
    Exit Function

Exception:
    If Not rsData Is Nothing Then: rsData.Close: Set rsData = Nothing
    If Not rsBoxData Is Nothing Then: rsBoxData.Close: Set rsBoxData = Nothing
    If Not pobjAdoDbConnection Is Nothing Then: pobjAdoDbConnection.Close: Set pobjAdoDbConnection = Nothing
    Call Err.Raise(Err.Number, "getBoxSelectData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :ボックス明細SQL
'        MODULE_ID        :createBoxDataSQL
'        NOTE             :FKS800.createBoxDataSQL09を参考に作成
'        CREATE_DATE      :2009/10/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createBoxDataSQL() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      YARD_MAST.YARD_CODE as YARD_CODE "
    strSQL = strSQL & "     ,YARD_MAST.YARD_NAME as YARD_NAME "
    strSQL = strSQL & "     ,Right('000000'+ CBOX_MAST.CBOXM_BOXNO, 6) as BOXNO_6 "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BOXNO as BOXNO "
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
    strSQL = strSQL & "     ,w01.CBRAT_CNO as CNO "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_BIKON as BIKON " & vbCrLf
    strSQL = strSQL & " from CBOX_MAST " & vbCrLf
    strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' " & vbCrLf
    strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' " & vbCrLf
    strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' " & vbCrLf
    strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' " & vbCrLf
    strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' " & vbCrLf
    strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' " & vbCrLf
    strSQL = strSQL & "  LEFT OUTER JOIN (select CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_YCODE,CBRA_TRAN.CBRAT_CNO from CBRA_TRAN,CNTA_MAST where CBRAT_CNO = CNTA_NO and CBRAT_YCODE = CNTA_CODE) as w01 on w01.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO " & vbCrLf
    strSQL = strSQL & " ,CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR AND NAME_MAST7.NAME_ID = '090' " & vbCrLf
    strSQL = strSQL & " ,YARD_MAST " & vbCrLf
    strSQL = strSQL & " where CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO " & vbCrLf

    strSQL = strSQL & MakeWhere() & vbCrLf

    strSQL = strSQL & " order by YARD_CODE,BOXNO_6,SETDATE,CNO "
    
    createBoxDataSQL = strSQL
    
End Function


'==============================================================================*
'
'        MODULE_NAME      :検索条件取得
'        MODULE_ID        :MakeWhere
'        NOTE             :
'        CREATE_DATE      :2009/10/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeWhere()
        
    On Error GoTo Exception

    Dim strSQL      As String
    
    '検索条件
    If Trim(Nz(Me.cmb_ヤードコード_From.VALUE)) <> "" Then
        strSQL = strSQL & " and YARD_MAST.YARD_CODE >= '" & Me.cmb_ヤードコード_From.VALUE & "'"
        If Trim(Nz(Me.cmb_ヤードコード_To.VALUE)) <> "" Then
            strSQL = strSQL & " and YARD_MAST.YARD_CODE <= '" & Me.cmb_ヤードコード_To.VALUE & "'"
        End If
    Else
        If Trim(Nz(Me.cmb_ヤードコード_To.VALUE)) <> "" Then
            strSQL = strSQL & " and YARD_MAST.YARD_CODE <= '" & Me.cmb_ヤードコード_To.VALUE & "'"
        Else
            strSQL = strSQL & " and YARD_MAST.YARD_CODE >= '0'"
        End If
    End If
    If Trim(Nz(Me.cmb_タイプ.VALUE)) <> "" Then
        strSQL = strSQL & " and CBOXM_TYPE = " & Me.cmb_タイプ.Column(1, Me.cmb_タイプ.ListIndex)
    End If
    If Trim(Nz(Me.cmb_ドア数.VALUE)) <> "" Then
        strSQL = strSQL & " and CBOXM_DOOR = " & Me.cmb_ドア数.Column(1, Me.cmb_ドア数.ListIndex)
    End If
    If Trim(Nz(Me.cmb_内装.VALUE)) <> "" Then
        strSQL = strSQL & " and CBOXM_INTYPE = " & Me.cmb_内装.Column(1, Me.cmb_内装.ListIndex)
    End If
    If Trim(Nz(Me.cmb_ドア位置.VALUE)) <> "" Then
        strSQL = strSQL & " and CBOXM_DOORPSTN = " & Me.cmb_ドア位置.Column(1, Me.cmb_ドア位置.ListIndex)
    End If
    If Trim(Nz(Me.cmb_サイズ.VALUE)) <> "" Then
        strSQL = strSQL & " and CBOXM_SIZE = " & Me.cmb_サイズ.Column(2, Me.cmb_サイズ.ListIndex)
    End If
    If Trim(Nz(Me.cmb_ドアタイプ.VALUE)) <> "" Then
        strSQL = strSQL & " and CBOXM_DOORTYPE = " & Me.cmb_ドアタイプ.Column(1, Me.cmb_ドアタイプ.ListIndex)
    End If
    If Trim(Nz(Me.cmb_設置階層.VALUE)) <> "" Then
        strSQL = strSQL & " and CBYST_SETFLOOR = " & Me.cmb_設置階層.Column(1, cmb_設置階層.ListIndex)
    End If
    
    strSQL = strSQL & " and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = YARD_MAST.YARD_CODE "

    MakeWhere = strSQL
    
    Exit Function

Exception:
    Call Err.Raise(Err.Number, "MakeWhere" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'****************************  ended or program ********************************




