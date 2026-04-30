VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS830"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  start or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤードコンテナ一覧
'   プログラムＩＤ　：　FKS830
'   作　成　日　　　：  2014/02/03
'   作　成　者　　　：  EGL
'   Ver             ：  0.0
'
'**********************************************
'   プログラム名　　：　ヤードコンテナ一覧
'   プログラムＩＤ　：　FKS830
'   作　成　日　　　：  2014/05/02
'   作　成　者　　　：  EGL
'   Ver             ：  0.1
'                   ：  検索条件の所有区分より「その他」削除、それに伴う検索条件組立部分を変更
'                   ：  検索条件の購入区分に「不明」追加、それに伴う検索条件組立部分を変更
'                   ：  検索条件の設置用途に「看板」追加、それに伴う検索条件組立部分を変更
'                   ：  検索条件の設置用途の「その他」を「不明」に変更(フォームのみ)
'                   ：  検索条件のコンテナ分類の「その他」を「不明」に変更(フォームのみ)
'
'**********************************************
'   プログラム名　　：　ヤードコンテナ一覧
'   プログラムＩＤ　：　FKS830
'   作　成　日　　　：  2015/03/25
'   作　成　者　　　：  M.HODNA
'   Ver             ：  0.2
'                       表示用の変数の型をInteger→Longへ変更
'**********************************************
'   修　正　日　　　：  2018/09/07
'   修　正　者　　　：  tajima
'   Ver             ：  0.3
'                       システム制作課題90対応、15ft,19.8ft対応
'**********************************************

Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FKS830"

Private Const P_TITLE                   As String = "ヤードコンテナ一覧"
Private Const P_LIST_WORK_TABLE         As String = "WK_FKS830_S01_TEMP"
Private Const P_QUERY_NAME              As String = "FKS830_Q01"            'クエリ名

Private Const P_ERROR                   As String = "エラー"

Private pobjDB                          As Object
Private pobjQueryDef                    As DAO.QueryDef                     'クエリ

Private pstrBumonCode                   As String                           '部門コード

Public gstrUserName                     As String

Private pstrDateValue                   As String                           '日付の入力チェック用

'==============================================================================*
'
'        MODULE_NAME      :cmb_ヤード_FROM_AfterUpdate
'        MODULE_ID        :cmb_ヤード_FROM_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤード_FROM_AfterUpdate()

    Dim isException                     As Boolean

    isException = False

On Error GoTo Exception

    If Me.cmb_ヤード_FROM.ListIndex = -1 Then
        Me.lbl_ヤード名_FROM.Caption = ""
    Else
        Me.lbl_ヤード名_FROM.Caption = Me.cmb_ヤード_FROM.Column(1, Me.cmb_ヤード_FROM.ListIndex)
    End If

    Me.cmb_ヤード_FROM = Format(Nz(Me.cmb_ヤード_FROM), "000000")

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_ヤード_FROM_AfterUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ヤード_TO_AfterUpdate
'        MODULE_ID        :cmb_ヤード_TO_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤード_TO_AfterUpdate()

    Dim isException                     As Boolean

    isException = False

On Error GoTo Exception

    If Me.cmb_ヤード_TO.ListIndex = -1 Then
        Me.lbl_ヤード名_TO.Caption = ""
    Else
        Me.lbl_ヤード名_TO.Caption = Me.cmb_ヤード_TO.Column(1, Me.cmb_ヤード_TO.ListIndex)
    End If

    Me.cmb_ヤード_TO = Format(Nz(Me.cmb_ヤード_TO), "000000")

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_ヤード_FROM_AfterUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_EXCEL出力_Click
'        MODULE_ID        :cmd_EXCEL出力_Click
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_EXCEL出力_Click()

    '合計行(フッター)の初期化
    Form_FKS830_S01.txt_合計_設置数_他社.VALUE = ""
    Form_FKS830_S01.txt_合計_設置数_自社.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_5Ft.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_6Ft.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_8Ft.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_10Ft.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_12Ft.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_20Ft.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_その他.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_計.VALUE = ""
    Form_FKS830_S01.txt_合計_市販物置.VALUE = ""
    Form_FKS830_S01.txt_合計_ユニットハウス.VALUE = ""
    Form_FKS830_S01.txt_合計_荷物収納_その他.VALUE = ""
    Form_FKS830_S01.txt_合計_バイク収納.VALUE = ""
    Form_FKS830_S01.txt_合計_通路.VALUE = ""
    Form_FKS830_S01.txt_合計_外枠.VALUE = ""
    Form_FKS830_S01.txt_合計_自社資材等.VALUE = ""
    Form_FKS830_S01.txt_合計_看板.VALUE = ""                    'INSERT 2014/05/02 MIYAMOTO
    Form_FKS830_S01.txt_合計_設置用途_その他.VALUE = ""
'    Form_FKS830_S01.txt_合計_設置数_その他.VALUE = ""          'DELETE 2014/05/02 MIYAMOTO
    Form_FKS830_S01.txt_合計_合計.VALUE = ""

    '検索条件入力チェック
    If fncSearchInputCheck = False Then
        Exit Sub
    End If

    'EXCEL出力
    Call createExcelData

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_検索_Click
'        MODULE_ID        :cmd_検索_Click
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_検索_Click()

    '合計行(フッター)の初期化
    Form_FKS830_S01.txt_合計_設置数_他社.VALUE = ""
    Form_FKS830_S01.txt_合計_設置数_自社.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_5Ft.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_6Ft.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_8Ft.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_10Ft.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_12Ft.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_20Ft.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_その他.VALUE = ""
    Form_FKS830_S01.txt_合計_海上コンテナ_計.VALUE = ""
    Form_FKS830_S01.txt_合計_市販物置.VALUE = ""
    Form_FKS830_S01.txt_合計_ユニットハウス.VALUE = ""
    Form_FKS830_S01.txt_合計_荷物収納_その他.VALUE = ""
    Form_FKS830_S01.txt_合計_バイク収納.VALUE = ""
    Form_FKS830_S01.txt_合計_通路.VALUE = ""
    Form_FKS830_S01.txt_合計_外枠.VALUE = ""
    Form_FKS830_S01.txt_合計_自社資材等.VALUE = ""
    Form_FKS830_S01.txt_合計_看板.VALUE = ""                    'INSERT 2014/05/02 MIYAMOTO
    Form_FKS830_S01.txt_合計_設置用途_その他.VALUE = ""
'    Form_FKS830_S01.txt_合計_設置数_その他.VALUE = ""          'DELETE 2014/05/02 MIYAMOTO
    Form_FKS830_S01.txt_合計_合計.VALUE = ""

    '検索条件入力チェック
    If fncSearchInputCheck = False Then
        Exit Sub
    End If

    '検索
    Call createListData

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Click
'        MODULE_ID        :cmd_終了_Click
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()

    Dim isException                     As Boolean

    isException = False

On Error GoTo Exception

    doCmd.Close acForm, Me.FormName

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_終了_Click", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        MODULE_ID        :Form_Activate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
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
'        MODULE_NAME      :Form_Close
'        MODULE_ID        :Form_Close
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()

    Dim isException                     As Boolean

    isException = False

On Error GoTo Exception

    'ワークテーブル削除
    Call dropWorkTable

    'コネクションクローズ
    If Not pobjDB Is Nothing Then
       pobjDB.Close
       Set pobjDB = Nothing
    End If

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Close", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        MODULE_ID        :Form_Load
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    Dim isException                     As Boolean

    isException = False

On Error GoTo Exception

    pstrBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If pstrBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    Set pobjDB = CurrentDb

    'ワークテーブルあったら削除
    If fncTableExist(P_LIST_WORK_TABLE) Then
        Call dropWorkTable
    End If

    'ワークテーブル作成
    Call createWorkTable

    gstrUserName = MSZZ000.LsGetUserName

    '画面初期設定
    Me.txt_解約予定日.VALUE = Format(DATE, "YYYY/MM/DD")
    Me.chk_賃貸種別_レンタルコンテナ = -1
    Me.chk_賃貸種別_トランクルーム = -1
    Me.chk_賃貸種別_バイク = -1
    Me.chk_所有区分_自社 = -1
    Me.chk_所有区分_他社 = -1
'    Me.chk_所有区分_その他 = -1            'DELETE 2014/05/02 MIYAMOTO
    Me.chk_購入区分_新品 = -1
    Me.chk_購入区分_中古 = -1
'    Me.chk_所有区分_その他 = -1            'DELETE 2014/05/02 MIYAMOTO
    Me.chk_購入区分_不明 = -1               'INSERT 2014/05/02 MIYAMOTO
    Me.chk_設置用途_荷物収納 = -1
    Me.chk_設置用途_バイク収納 = -1
    Me.chk_設置用途_通路 = -1
    Me.chk_設置用途_外枠 = -1
    Me.chk_設置用途_自社資材等 = -1
    Me.chk_設置用途_看板 = -1               'INSERT 2014/05/02 MIYAMOTO
    Me.chk_設置用途_その他 = -1
    Me.chk_コンテナ分類_海上コンテナ = -1
    Me.chk_コンテナ分類_物置 = -1
    Me.chk_コンテナ分類_ユニットハウス = -1
    Me.chk_コンテナ分類_その他 = -1

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Load", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ACCESSテーブル存在チェック
'        MODULE_ID        :fncTableExist
'        NOTE             :IN strTableName = テーブル名
'        RETURN           :True=存在する False=存在しない
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncTableExist(strTableName As String) As Boolean

    Dim tdf                             As TableDef

    fncTableExist = False

    For Each tdf In pobjDB.TableDefs
        If tdf.NAME = strTableName Then
            fncTableExist = True
            Exit For
        End If
    Next tdf

End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル削除
'        MODULE_ID        :dropWorkTable
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub dropWorkTable()

On Error GoTo Exception

    Me.ヤードコンテナ一覧.Form.RecordSource = ""
    Me.ヤードコンテナ一覧.Form.OrderBy = ""
    Me.ヤードコンテナ一覧.Form.Requery
    pobjDB.Execute "DROP TABLE " & P_LIST_WORK_TABLE & " ;", dbFailOnError

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "dropWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル作成
'        MODULE_ID        :createWorkTable
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub createWorkTable()

    Dim strSQL                          As String

On Error GoTo Exception

    '一覧画面用ワークテーブル
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE " & P_LIST_WORK_TABLE & " ("
    strSQL = strSQL & " ヤード               VARCHAR(6) "
    strSQL = strSQL & ",名称                 VARCHAR(36) "
    strSQL = strSQL & ",設置数_他社          NUMERIC "
    strSQL = strSQL & ",設置数_自社          NUMERIC "
    strSQL = strSQL & ",海上コンテナ_5Ft     NUMERIC "
    strSQL = strSQL & ",海上コンテナ_6Ft     NUMERIC "
    strSQL = strSQL & ",海上コンテナ_8Ft     NUMERIC "
    strSQL = strSQL & ",海上コンテナ_10Ft    NUMERIC "
    strSQL = strSQL & ",海上コンテナ_12Ft    NUMERIC "
    strSQL = strSQL & ",海上コンテナ_20Ft    NUMERIC "
    strSQL = strSQL & ",海上コンテナ_その他  NUMERIC "
    strSQL = strSQL & ",海上コンテナ_計      NUMERIC "
    strSQL = strSQL & ",市販物置             NUMERIC "
    strSQL = strSQL & ",ユニットハウス       NUMERIC "
    strSQL = strSQL & ",荷物収納_その他      NUMERIC "
    strSQL = strSQL & ",バイク収納           NUMERIC "
    strSQL = strSQL & ",通路                 NUMERIC "
    strSQL = strSQL & ",外枠                 NUMERIC "
    strSQL = strSQL & ",自社資材等           NUMERIC "
    strSQL = strSQL & ",看板                 NUMERIC "              'INSERT 2014/05/02 MIYAMOTO
    strSQL = strSQL & ",設置用途_その他      NUMERIC "
'    strSQL = strSQL & ",設置数_その他        NUMERIC "             'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & ",合計                 NUMERIC "
    strSQL = strSQL & ",設置未割付数         NUMERIC "
    strSQL = strSQL & ",建築確認             VARCHAR(6) "
    strSQL = strSQL & ",用途                 VARCHAR(10) "
    strSQL = strSQL & ",営業開始日           VARCHAR(10) "
    strSQL = strSQL & ",営業終了日           VARCHAR(10)"
    strSQL = strSQL & ",解約予定日           VARCHAR(10)"
    strSQL = strSQL & ",借上物件             VARCHAR(6) "
    strSQL = strSQL & ",部屋_ボックス        NUMERIC "
    strSQL = strSQL & ",部屋_バイク          NUMERIC "
    strSQL = strSQL & ",部屋_屋外            NUMERIC "
    strSQL = strSQL & ",部屋_室内            NUMERIC "
    strSQL = strSQL & ",部屋_その他          NUMERIC "
    strSQL = strSQL & ",部屋未割付数         NUMERIC "
    strSQL = strSQL & ")"
    pobjDB.Execute strSQL, dbFailOnError

    'RecordSource設定
    'ソート順：ヤードの昇順
    Me.ヤードコンテナ一覧.Form.RecordSource = "SELECT * FROM " & P_LIST_WORK_TABLE & " ORDER BY ヤード ASC;"
    Me.ヤードコンテナ一覧.Form.OrderBy = "ヤード"

    'サブフォームのリフレッシュ
    Me.ヤードコンテナ一覧.Form.Requery

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "createWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルデータ削除
'        MODULE_ID        :deleteWorkTable
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub deleteWorkTable()

    Dim strSQL                          As String

On Error GoTo Exception

    Call dropWorkTable
    Call createWorkTable

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "deleteWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :EXCEL作成
'        MODULE_ID        :createExcelData
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub createExcelData()

On Error GoTo Exception

    'クエリ作成
    Call subCreateQuery

    'ワークテーブル削除
    Call deleteWorkTable

    'ヤードコンテナ一覧データ取得
    Call getSelectListData

    'Excel起動
    doCmd.OutputTo acQuery, P_QUERY_NAME, "MicrosoftExcel(*.xls)", "", True, ""

    Exit Sub

Exception:
    If Err <> 0 And Err <> 2501 Then
        Call Err.Raise(Err.Number, "createListData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :クエリ作成
'        MODULE_ID        :subCreateQuery
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCreateQuery()

On Error GoTo Exception

    Dim strSQL As String

    Set pobjQueryDef = pobjDB.CreateQueryDef(P_QUERY_NAME)

    strSQL = ""
    strSQL = strSQL & "SELECT * FROM " & P_LIST_WORK_TABLE & " ORDER BY ヤード ASC"
    pobjQueryDef.sql = strSQL

Exception:
    Err.Clear
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :一覧作成
'        MODULE_ID        :createListData
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub createListData()

On Error GoTo Exception

    'ワークテーブル削除
    Call deleteWorkTable

    'ヤードコンテナ一覧データ取得
    Call getSelectListData

    'サブフォームのリフレッシュ
    Me.ヤードコンテナ一覧.Form.Requery

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "createListData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :一覧表時用ワークテーブル更新
'        MODULE_ID        :getSelectListData
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getSelectListData()

    Dim objAdoDbConnection              As Object
    Dim rsData                          As Object       'ADODB.Recordset
    Dim rsBoxData                       As Object
    Dim rsContData                      As Object
    Dim rsListData                      As Recordset
    Dim strSQL                          As String
    Dim strSize                         As String
    Dim intKEIYAKU_CNT                  As Integer
    Dim strKaseSqlSvr                   As String

    Dim StrYard                         As String       'ヤード
    Dim lngYARD                         As Long         'ヤード
    Dim strTINTI(2)                     As String       '賃貸種別チェックボックス値格納用

    '2015/03/25 M.HONDA UPD
    '合計用変数
    Dim intTASYA_CNT                    As Long      '設置数_他社
    Dim intJISYA_CNT                    As Long      '設置数_自社
    Dim intCONT5FT_CNT                  As Long      '海上コンテナ5Ft
    Dim intCONT6FT_CNT                  As Long      '海上コンテナ6Ft
    Dim intCONT8FT_CNT                  As Long      '海上コンテナ8Ft
    Dim intCONT10FT_CNT                 As Long      '海上コンテナ10Ft
    Dim intCONT12FT_CNT                 As Long      '海上コンテナ12Ft
    Dim intCONT20FT_CNT                 As Long      '海上コンテナ20Ft
    Dim intCONTETC_CNT                  As Long      '海上コンテナその他
    Dim intCONTTTL_CNT                  As Long      '海上コンテナ計
    Dim intMONOOKI_CNT                  As Long      '市販物置
    Dim intUHOUSE_CNT                   As Long      'ユニットハウス
    Dim intNSYUNO_ETC_CNT               As Long      '荷物収納_その他
    Dim intBSYUNO_CNT                   As Long      'バイク収納
    Dim intTUURO_CNT                    As Long      '通路
    Dim intSOTOWAKU_CNT                 As Long      '外枠
    Dim intJSIZAI_CNT                   As Long      '自社資材等
    Dim intKANBAN_CNT                   As Long      '看板
    Dim intSETTI_YOUTO_ETC_CNT          As Long      '設置用途_その他
    Dim intTOTAL_CNT                    As Long      '合計
    Dim intSETTI_MI_CNT                 As Long      '設置未割付数
    Dim intROOM_BOX_CNT                 As Long      '部屋_ボックス
    Dim intROOM_BIKE_CNT                As Long      '部屋_バイク
    Dim intROOM_SOTO_CNT                As Long      '部屋_屋外
    Dim intROOM_UCHI_CNT                As Long      '部屋_室内
    Dim intROOM_ETC_CNT                 As Long      '部屋_その他
    Dim intROOM_MI_CNT                  As Long      '部屋未割付数
    Dim debugCNT                        As Long      'デバッグ用カウンター
    
'    Dim intTASYA_CNT                    As Integer      '設置数_他社
'    Dim intJISYA_CNT                    As Integer      '設置数_自社
'    Dim intCONT5FT_CNT                  As Integer      '海上コンテナ5Ft
'    Dim intCONT6FT_CNT                  As Integer      '海上コンテナ6Ft
'    Dim intCONT8FT_CNT                  As Integer      '海上コンテナ8Ft
'    Dim intCONT10FT_CNT                 As Integer      '海上コンテナ10Ft
'    Dim intCONT12FT_CNT                 As Integer      '海上コンテナ12Ft
'    Dim intCONT20FT_CNT                 As Integer      '海上コンテナ20Ft
'    Dim intCONTETC_CNT                  As Integer      '海上コンテナその他
'    Dim intCONTTTL_CNT                  As Integer      '海上コンテナ計
'    Dim intMONOOKI_CNT                  As Integer      '市販物置
'    Dim intUHOUSE_CNT                   As Integer      'ユニットハウス
'    Dim intNSYUNO_ETC_CNT               As Integer      '荷物収納_その他
'    Dim intBSYUNO_CNT                   As Integer      'バイク収納
'    Dim intTUURO_CNT                    As Integer      '通路
'    Dim intSOTOWAKU_CNT                 As Integer      '外枠
'    Dim intJSIZAI_CNT                   As Integer      '自社資材等
'    Dim intKANBAN_CNT                   As Integer      '看板                       'INSERT 2014/05/02 MIYAMOTO
'    Dim intSETTI_YOUTO_ETC_CNT          As Integer      '設置用途_その他
''    Dim intETC_CNT                      As Integer      '設置数_その他             'DELETE 2014/05/02 MIYAMOTO
'    Dim intTOTAL_CNT                    As Integer      '合計
'    Dim intSETTI_MI_CNT                 As Integer      '設置未割付数
'    Dim intROOM_BOX_CNT                 As Integer      '部屋_ボックス
'    Dim intROOM_BIKE_CNT                As Integer      '部屋_バイク
'    Dim intROOM_SOTO_CNT                As Integer      '部屋_屋外
'    Dim intROOM_UCHI_CNT                As Integer      '部屋_室内
'    Dim intROOM_ETC_CNT                 As Integer      '部屋_その他
'    Dim intROOM_MI_CNT                  As Integer      '部屋未割付数
'    Dim debugCNT                        As Integer      'デバッグ用カウンター

On Error GoTo Exception

    'ヤードコンテナ一覧データ取得
    strSQL = ""
    strSQL = strSQL & "SELECT      YARD_MAST.YARD_CODE AS YARD_CODE" & vbCrLf           'ヤード
    strSQL = strSQL & ",           YARD_MAST.YARD_NAME AS YARD_NAME" & vbCrLf           '名称
    strSQL = strSQL & ",           YARD_MAST.YARD_TINTI AS TINTI_VALUE" & vbCrLf        '用途(賃貸種別)コード
    strSQL = strSQL & ",           NAME_MAST1.NAME_RYAK AS TINTI" & vbCrLf              '用途(賃貸種別)略称
    strSQL = strSQL & ",           YARD_MAST.YARD_KENTIKU AS KENTIKU_VALUE" & vbCrLf    '建築確認コード
    strSQL = strSQL & ",           NAME_MAST2.NAME_RYAK AS KENTIKU" & vbCrLf            '建築確認略称
    strSQL = strSQL & ",           YARD_MAST.YARD_BEGIN_DAY AS BEGIN_DAY" & vbCrLf      '営業開始日
    strSQL = strSQL & ",           YARD_MAST.YARD_RENTEND_DAY AS RENTEND_DAY" & vbCrLf  '営業終了日
    strSQL = strSQL & ",           YARD_MAST.YARD_END_DAY AS END_DAY" & vbCrLf          '解約予定日
    strSQL = strSQL & ",           YARD_MAST.YARD_LEASE_NO AS LEASE_NO" & vbCrLf        '借上物件

'↓ INSERT 2014/05/02 MIYAMOTO
    '設置数
    strSQL = strSQL & ",           ISNULL(TASYA.TASYA_CNT, 0) AS TASYA_CNT" & vbCrLf                                '設置数_他社
    strSQL = strSQL & ",           ISNULL(JISYA.JISYA_CNT, 0) AS JISYA_CNT" & vbCrLf                                '設置数_自社
    strSQL = strSQL & ",           ISNULL(CONT5FT.CONT5FT_CNT, 0) AS CONT5FT_CNT" & vbCrLf                          '海上コンテナ5Ft
    strSQL = strSQL & ",           ISNULL(CONT6FT.CONT6FT_CNT, 0) AS CONT6FT_CNT" & vbCrLf                          '海上コンテナ6Ft
    strSQL = strSQL & ",           ISNULL(CONT8FT.CONT8FT_CNT, 0) AS CONT8FT_CNT" & vbCrLf                          '海上コンテナ8Ft
    strSQL = strSQL & ",           ISNULL(CONT10FT.CONT10FT_CNT, 0) AS CONT10FT_CNT" & vbCrLf                       '海上コンテナ10Ft
    strSQL = strSQL & ",           ISNULL(CONT12FT.CONT12FT_CNT, 0) AS CONT12FT_CNT" & vbCrLf                       '海上コンテナ12Ft
    strSQL = strSQL & ",           ISNULL(CONT20FT.CONT20FT_CNT, 0) AS CONT20FT_CNT" & vbCrLf                       '海上コンテナ20Ft
    strSQL = strSQL & ",           ISNULL(CONTETC.CONTETC_CNT, 0) AS CONTETC_CNT" & vbCrLf                          '海上コンテナその他
    strSQL = strSQL & ",           ISNULL(CONTTTL.CONTTTL_CNT, 0) AS CONTTTL_CNT" & vbCrLf                          '海上コンテナ計
    strSQL = strSQL & ",           ISNULL(MONOOKI.MONOOKI_CNT, 0) AS MONOOKI_CNT" & vbCrLf                          '市販物置
    strSQL = strSQL & ",           ISNULL(UHOUSE.UHOUSE_CNT, 0) AS UHOUSE_CNT" & vbCrLf                             'ユニットハウス
    strSQL = strSQL & ",           ISNULL(NSYUNO_ETC.NSYUNO_ETC_CNT, 0) AS NSYUNO_ETC_CNT" & vbCrLf                 '荷物収納_その他
    strSQL = strSQL & ",           ISNULL(BSYUNO.BSYUNO_CNT, 0) AS BSYUNO_CNT" & vbCrLf                             'バイク収納
    strSQL = strSQL & ",           ISNULL(TUURO.TUURO_CNT, 0) AS TUURO_CNT" & vbCrLf                                '通路
    strSQL = strSQL & ",           ISNULL(SOTOWAKU.SOTOWAKU_CNT, 0) AS SOTOWAKU_CNT" & vbCrLf                       '外枠
    strSQL = strSQL & ",           ISNULL(JSIZAI.JSIZAI_CNT, 0) AS JSIZAI_CNT" & vbCrLf                             '自社資材等
    strSQL = strSQL & ",           ISNULL(KANBAN.KANBAN_CNT, 0) AS KANBAN_CNT" & vbCrLf                             '看板
    strSQL = strSQL & ",           ISNULL(SETTI_YOUTO_ETC.SETTI_YOUTO_ETC_CNT, 0) AS SETTI_YOUTO_ETC_CNT" & vbCrLf  '荷物収納_その他
    strSQL = strSQL & ",           ISNULL(TOTAL.TOTAL_CNT, 0) AS TOTAL_CNT" & vbCrLf                                '合計

    '部屋数
    strSQL = strSQL & ",           ISNULL(BOX.BOX_CNT, 0) AS BOX_CNT" & vbCrLf                                      'ボックス
    strSQL = strSQL & ",           ISNULL(BIKE.BIKE_CNT, 0) AS BIKE_CNT" & vbCrLf                                   'バイク
    strSQL = strSQL & ",           ISNULL(OUTDOOR.OUTDOOR_CNT, 0) AS OUTDOOR_CNT" & vbCrLf                          '屋外
    strSQL = strSQL & ",           ISNULL(INDOOR.INDOOR_CNT, 0) AS INDOOR_CNT" & vbCrLf                             '室内
    strSQL = strSQL & ",           ISNULL(ETC.ETC_CNT, 0) AS ETC_CNT" & vbCrLf                                      '不明
'↑ INSERT 2014/05/02 MIYAMOTO
    
    strSQL = strSQL & "FROM        YARD_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON YARD_MAST.YARD_TINTI = NAME_MAST1.NAME_CODE AND NAME_MAST1.NAME_ID = '207'" & vbCrLf
    strSQL = strSQL & "                      LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON YARD_MAST.YARD_KENTIKU = NAME_MAST2.NAME_CODE AND NAME_MAST2.NAME_ID = '041'" & vbCrLf
    
'↓ INSERT 2014/05/02 MIYAMOTO
    '設置数取得条件
    strSQL = strSQL & createBoxData

    '部屋数取得条件
    strSQL = strSQL & createContData
'↑ INSERT 2014/05/02 MIYAMOTO

    ' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    'WHERE句作成
    strSQL = strSQL & "WHERE       1 = 1" & vbCrLf

    '<<ヤード_FROM>>
    If Nz(Me.cmb_ヤード_FROM.VALUE, "") <> "" Then
        strSQL = strSQL & "AND         YARD_MAST.YARD_CODE >= " & Me.cmb_ヤード_FROM.VALUE & vbCrLf
    End If

    '<<ヤード_TO>>
    If Nz(Me.cmb_ヤード_TO.VALUE, "") <> "" Then
        strSQL = strSQL & "AND         YARD_MAST.YARD_CODE <= " & Me.cmb_ヤード_TO.VALUE & vbCrLf
    End If

    '<<営業開始日_FROM>>
    If Nz(Me.txt_営業開始日_FROM.VALUE, "") <> "" Then
        strSQL = strSQL & "AND         CONVERT(VARCHAR, YARD_MAST.YARD_BEGIN_DAY, 111) >= '" & Me.txt_営業開始日_FROM.VALUE & "'" & vbCrLf
    End If

    '<<営業開始日_TO>>
    If Nz(Me.txt_営業開始日_TO.VALUE, "") <> "" Then
        strSQL = strSQL & "AND         CONVERT(VARCHAR, YARD_MAST.YARD_BEGIN_DAY, 111) <= '" & Me.txt_営業開始日_TO.VALUE & "'" & vbCrLf
    End If

    '<<営業終了日_FROM>>
    If Nz(Me.txt_営業終了日_FROM.VALUE, "") <> "" Then
        strSQL = strSQL & "AND         CONVERT(VARCHAR, YARD_MAST.YARD_RENTEND_DAY, 111) >= '" & Me.txt_営業終了日_FROM.VALUE & "'" & vbCrLf
    End If

    '<<営業終了日_TO>>
    If Nz(Me.txt_営業終了日_TO.VALUE, "") <> "" Then
        strSQL = strSQL & "AND         CONVERT(VARCHAR, YARD_MAST.YARD_RENTEND_DAY, 111) <= '" & Me.txt_営業終了日_TO.VALUE & "'" & vbCrLf
    End If

    '<<解約予定日>>
    If Nz(Me.txt_解約予定日.VALUE, "") <> "" Then
        strSQL = strSQL & "AND         (YARD_MAST.YARD_END_DAY Is Null OR CONVERT(VARCHAR, YARD_MAST.YARD_END_DAY, 111) > '" & Me.txt_解約予定日.VALUE & "')" & vbCrLf
    End If

    '<<賃貸種別>>
    Erase strTINTI
    If Nz(Me.chk_賃貸種別_レンタルコンテナ.VALUE, "") = -1 Then
        strTINTI(0) = "'0'"
    Else
        strTINTI(0) = "'-0'"
    End If
    If Nz(Me.chk_賃貸種別_トランクルーム.VALUE, "") = -1 Then
        strTINTI(1) = "'1'"
    Else
        strTINTI(1) = "'-1'"
    End If
    If Nz(Me.chk_賃貸種別_バイク.VALUE, "") = -1 Then
        strTINTI(2) = "'4'"
    Else
        strTINTI(2) = "'-4'"
    End If
    If Nz(strTINTI(0), "") <> "" Or Nz(strTINTI(1), "") <> "" Or Nz(strTINTI(2), "") <> "" Then
        strSQL = strSQL & "AND         CONVERT(VARCHAR, YARD_MAST.YARD_TINTI) IN (" & Replace(strTINTI(0) & strTINTI(1) & strTINTI(2), "''", "', '") & ")" & vbCrLf
    End If

    strSQL = strSQL & ";" & vbCrLf
    ' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)

    If rsData.EOF = True Then
        MsgBox "指定した検索条件に該当するデータはありません。", vbInformation, P_ERROR
    Else

        debugCNT = 0

        '合計用変数の初期化
        intTASYA_CNT = 0            '設置数_他社
        intJISYA_CNT = 0            '設置数_自社
        intCONT5FT_CNT = 0          '海上コンテナ5Ft
        intCONT6FT_CNT = 0          '海上コンテナ6Ft
        intCONT8FT_CNT = 0          '海上コンテナ8Ft
        intCONT10FT_CNT = 0         '海上コンテナ10Ft
        intCONT12FT_CNT = 0         '海上コンテナ12Ft
        intCONT20FT_CNT = 0         '海上コンテナ20Ft
        intCONTETC_CNT = 0          '海上コンテナその他
        intCONTTTL_CNT = 0          '海上コンテナ計
        intMONOOKI_CNT = 0          '市販物置
        intUHOUSE_CNT = 0           'ユニットハウス
        intNSYUNO_ETC_CNT = 0       '荷物収納_その他
        intBSYUNO_CNT = 0           'バイク収納
        intTUURO_CNT = 0            '通路
        intSOTOWAKU_CNT = 0         '外枠
        intJSIZAI_CNT = 0           '自社資材等
        intKANBAN_CNT = 0           '看板                           'INSERT 2014/05/02 MIYAMOTO
        intSETTI_YOUTO_ETC_CNT = 0  '設置用途_その他
'        intETC_CNT = 0              '設置数_その他                 'DELETE 2014/05/02 MIYAMOTO
        intTOTAL_CNT = 0            '合計
        intSETTI_MI_CNT = 0         '設置未割付数
        intROOM_BOX_CNT = 0         '部屋_ボックス
        intROOM_BIKE_CNT = 0        '部屋_バイク
        intROOM_SOTO_CNT = 0        '部屋_屋外
        intROOM_UCHI_CNT = 0        '部屋_室内
        intROOM_ETC_CNT = 0         '部屋_その他
        intROOM_MI_CNT = 0          '部屋未割付数

        While Not rsData.EOF

            debugCNT = debugCNT + 1
            StrYard = Format(rsData.Fields("YARD_CODE"), "000000")
            lngYARD = Trim(rsData.Fields("YARD_CODE"))

            strSQL = ""
            strSQL = strSQL & "SELECT * FROM " & P_LIST_WORK_TABLE
            Set rsListData = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)

            rsListData.AddNew

            'ヤード情報設定
            rsListData.Fields("ヤード") = StrYard
            rsListData.Fields("名称") = rsData.Fields("YARD_NAME")
            rsListData.Fields("建築確認") = rsData.Fields("KENTIKU")
            rsListData.Fields("用途") = rsData.Fields("TINTI")
            rsListData.Fields("営業開始日") = Format(rsData.Fields("BEGIN_DAY"), "yyyy/mm/dd")
            rsListData.Fields("営業終了日") = Format(rsData.Fields("RENTEND_DAY"), "yyyy/mm/dd")
            rsListData.Fields("解約予定日") = Format(rsData.Fields("END_DAY"), "yyyy/mm/dd")
            rsListData.Fields("借上物件") = rsData.Fields("LEASE_NO")

'↓ DELETE 2014/05/02 MIYAMOTO
'            '設置数取得
'            strSQL = createBoxData(strYard)
'            Set rsBoxData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)
'↑ DELETE 2014/05/02 MIYAMOTO

'↓ UPDATE 2014/05/02 MIYAMOTO
'            'コンテナ情報設定
'            rsListData.Fields("設置数_他社") = Format(rsBoxData.Fields("TASYA_CNT"), "##,##0")
'            rsListData.Fields("設置数_自社") = Format(rsBoxData.Fields("JISYA_CNT"), "##,##0")
'            rsListData.Fields("海上コンテナ_5Ft") = Format(rsBoxData.Fields("CONT5FT_CNT"), "##,##0")
'            rsListData.Fields("海上コンテナ_6Ft") = Format(rsBoxData.Fields("CONT6FT_CNT"), "##,##0")
'            rsListData.Fields("海上コンテナ_8Ft") = Format(rsBoxData.Fields("CONT8FT_CNT"), "##,##0")
'            rsListData.Fields("海上コンテナ_10Ft") = Format(rsBoxData.Fields("CONT10FT_CNT"), "##,##0")
'            rsListData.Fields("海上コンテナ_12Ft") = Format(rsBoxData.Fields("CONT12FT_CNT"), "##,##0")
'            rsListData.Fields("海上コンテナ_20Ft") = Format(rsBoxData.Fields("CONT20FT_CNT"), "##,##0")
'            rsListData.Fields("海上コンテナ_その他") = Format(rsBoxData.Fields("CONTETC_CNT"), "##,##0")
'            rsListData.Fields("海上コンテナ_計") = Format(rsBoxData.Fields("CONTTTL_CNT"), "##,##0")
'            rsListData.Fields("市販物置") = Format(rsBoxData.Fields("MONOOKI_CNT"), "##,##0")
'            rsListData.Fields("ユニットハウス") = Format(rsBoxData.Fields("UHOUSE_CNT"), "##,##0")
'            rsListData.Fields("荷物収納_その他") = Format(rsBoxData.Fields("NSYUNO_ETC_CNT"), "##,##0")
'            rsListData.Fields("バイク収納") = Format(rsBoxData.Fields("BSYUNO_CNT"), "##,##0")
'            rsListData.Fields("通路") = Format(rsBoxData.Fields("TUURO_CNT"), "##,##0")
'            rsListData.Fields("外枠") = Format(rsBoxData.Fields("SOTOWAKU_CNT"), "##,##0")
'            rsListData.Fields("自社資材等") = Format(rsBoxData.Fields("JSIZAI_CNT"), "##,##0")
'            rsListData.Fields("看板") = Format(rsBoxData.Fields("KANBAN_CNT"), "##,##0")                        'INSERT 2014/05/02 MIYAMOTO
'            rsListData.Fields("設置用途_その他") = Format(rsBoxData.Fields("SETTI_YOUTO_ETC_CNT"), "##,##0")
'            rsListData.Fields("設置数_その他") = Format(rsBoxData.Fields("ETC_CNT"), "##,##0")
'            rsListData.Fields("合計") = Format(rsBoxData.Fields("TOTAL_CNT"), "##,##0")
'            rsListData.Fields("設置未割付数") = Format(0, "##,##0")
'
'            '合計行の集計
'            intTASYA_CNT = intTASYA_CNT + rsBoxData.Fields("TASYA_CNT")                                 '設置数_他社
'            intJISYA_CNT = intJISYA_CNT + rsBoxData.Fields("JISYA_CNT")                                 '設置数_自社
'            intCONT5FT_CNT = intCONT5FT_CNT + rsBoxData.Fields("CONT5FT_CNT")                           '海上コンテナ5Ft
'            intCONT6FT_CNT = intCONT6FT_CNT + rsBoxData.Fields("CONT6FT_CNT")                           '海上コンテナ6Ft
'            intCONT8FT_CNT = intCONT8FT_CNT + rsBoxData.Fields("CONT8FT_CNT")                           '海上コンテナ8Ft
'            intCONT10FT_CNT = intCONT10FT_CNT + rsBoxData.Fields("CONT10FT_CNT")                        '海上コンテナ10Ft
'            intCONT12FT_CNT = intCONT12FT_CNT + rsBoxData.Fields("CONT12FT_CNT")                        '海上コンテナ12Ft
'            intCONT20FT_CNT = intCONT20FT_CNT + rsBoxData.Fields("CONT20FT_CNT")                        '海上コンテナ20Ft
'            intCONTETC_CNT = intCONTETC_CNT + rsBoxData.Fields("CONTETC_CNT")                           '海上コンテナその他
'            intCONTTTL_CNT = intCONTTTL_CNT + rsBoxData.Fields("CONTTTL_CNT")                           '海上コンテナ計
'            intMONOOKI_CNT = intMONOOKI_CNT + rsBoxData.Fields("MONOOKI_CNT")                           '市販物置
'            intUHOUSE_CNT = intUHOUSE_CNT + rsBoxData.Fields("UHOUSE_CNT")                              'ユニットハウス
'            intNSYUNO_ETC_CNT = intNSYUNO_ETC_CNT + rsBoxData.Fields("NSYUNO_ETC_CNT")                  '荷物収納_その他
'            intBSYUNO_CNT = intBSYUNO_CNT + rsBoxData.Fields("BSYUNO_CNT")                              'バイク収納
'            intTUURO_CNT = intTUURO_CNT + rsBoxData.Fields("TUURO_CNT")                                 '通路
'            intSOTOWAKU_CNT = intSOTOWAKU_CNT + rsBoxData.Fields("SOTOWAKU_CNT")                        '外枠
'            intJSIZAI_CNT = intJSIZAI_CNT + rsBoxData.Fields("JSIZAI_CNT")                              '自社資材等
'            intKANBAN_CNT = intKANBAN_CNT + rsBoxData.Fields("KANBAN_CNT")                              '看板               'INSERT 2014/05/02 MIYAMOTO
'            intSETTI_YOUTO_ETC_CNT = intSETTI_YOUTO_ETC_CNT + rsBoxData.Fields("SETTI_YOUTO_ETC_CNT")   '設置用途_その他
'            intETC_CNT = intETC_CNT + rsBoxData.Fields("ETC_CNT")                                       '設置数_その他
'            intTOTAL_CNT = intTOTAL_CNT + rsBoxData.Fields("TOTAL_CNT")                                 '合計
'            intSETTI_MI_CNT = intSETTI_MI_CNT + 0                                                       '設置未割付数
'
'            rsBoxData.Close
'            Set rsBoxData = Nothing

            'コンテナ情報設定
            rsListData.Fields("設置数_他社") = Format(rsData.Fields("TASYA_CNT"), "##,##0")
            rsListData.Fields("設置数_自社") = Format(rsData.Fields("JISYA_CNT"), "##,##0")
            rsListData.Fields("海上コンテナ_5Ft") = Format(rsData.Fields("CONT5FT_CNT"), "##,##0")
            rsListData.Fields("海上コンテナ_6Ft") = Format(rsData.Fields("CONT6FT_CNT"), "##,##0")
            rsListData.Fields("海上コンテナ_8Ft") = Format(rsData.Fields("CONT8FT_CNT"), "##,##0")
            rsListData.Fields("海上コンテナ_10Ft") = Format(rsData.Fields("CONT10FT_CNT"), "##,##0")
            rsListData.Fields("海上コンテナ_12Ft") = Format(rsData.Fields("CONT12FT_CNT"), "##,##0")
            rsListData.Fields("海上コンテナ_20Ft") = Format(rsData.Fields("CONT20FT_CNT"), "##,##0")
            rsListData.Fields("海上コンテナ_その他") = Format(rsData.Fields("CONTETC_CNT"), "##,##0")
            rsListData.Fields("海上コンテナ_計") = Format(rsData.Fields("CONTTTL_CNT"), "##,##0")
            rsListData.Fields("市販物置") = Format(rsData.Fields("MONOOKI_CNT"), "##,##0")
            rsListData.Fields("ユニットハウス") = Format(rsData.Fields("UHOUSE_CNT"), "##,##0")
            rsListData.Fields("荷物収納_その他") = Format(rsData.Fields("NSYUNO_ETC_CNT"), "##,##0")
            rsListData.Fields("バイク収納") = Format(rsData.Fields("BSYUNO_CNT"), "##,##0")
            rsListData.Fields("通路") = Format(rsData.Fields("TUURO_CNT"), "##,##0")
            rsListData.Fields("外枠") = Format(rsData.Fields("SOTOWAKU_CNT"), "##,##0")
            rsListData.Fields("自社資材等") = Format(rsData.Fields("JSIZAI_CNT"), "##,##0")
            rsListData.Fields("看板") = Format(rsData.Fields("KANBAN_CNT"), "##,##0")
            rsListData.Fields("設置用途_その他") = Format(rsData.Fields("SETTI_YOUTO_ETC_CNT"), "##,##0")
            rsListData.Fields("合計") = Format(rsData.Fields("TOTAL_CNT"), "##,##0")
            rsListData.Fields("設置未割付数") = Format(0, "##,##0")

            '合計行の集計
            intTASYA_CNT = intTASYA_CNT + rsData.Fields("TASYA_CNT")                                    '設置数_他社
            intJISYA_CNT = intJISYA_CNT + rsData.Fields("JISYA_CNT")                                    '設置数_自社
            intCONT5FT_CNT = intCONT5FT_CNT + rsData.Fields("CONT5FT_CNT")                              '海上コンテナ5Ft
            intCONT6FT_CNT = intCONT6FT_CNT + rsData.Fields("CONT6FT_CNT")                              '海上コンテナ6Ft
            intCONT8FT_CNT = intCONT8FT_CNT + rsData.Fields("CONT8FT_CNT")                              '海上コンテナ8Ft
            intCONT10FT_CNT = intCONT10FT_CNT + rsData.Fields("CONT10FT_CNT")                           '海上コンテナ10Ft
            intCONT12FT_CNT = intCONT12FT_CNT + rsData.Fields("CONT12FT_CNT")                           '海上コンテナ12Ft
            intCONT20FT_CNT = intCONT20FT_CNT + rsData.Fields("CONT20FT_CNT")                           '海上コンテナ20Ft
            intCONTETC_CNT = intCONTETC_CNT + rsData.Fields("CONTETC_CNT")                              '海上コンテナ不明
            intCONTTTL_CNT = intCONTTTL_CNT + rsData.Fields("CONTTTL_CNT")                              '海上コンテナ計
            intMONOOKI_CNT = intMONOOKI_CNT + rsData.Fields("MONOOKI_CNT")                              '市販物置
            intUHOUSE_CNT = intUHOUSE_CNT + rsData.Fields("UHOUSE_CNT")                                 'ユニットハウス
            intNSYUNO_ETC_CNT = intNSYUNO_ETC_CNT + rsData.Fields("NSYUNO_ETC_CNT")                     '荷物収納_不明
            intBSYUNO_CNT = intBSYUNO_CNT + rsData.Fields("BSYUNO_CNT")                                 'バイク収納
            intTUURO_CNT = intTUURO_CNT + rsData.Fields("TUURO_CNT")                                    '通路
            intSOTOWAKU_CNT = intSOTOWAKU_CNT + rsData.Fields("SOTOWAKU_CNT")                           '外枠
            intJSIZAI_CNT = intJSIZAI_CNT + rsData.Fields("JSIZAI_CNT")                                 '自社資材等
            intKANBAN_CNT = intKANBAN_CNT + rsData.Fields("KANBAN_CNT")                                 '看板
            intSETTI_YOUTO_ETC_CNT = intSETTI_YOUTO_ETC_CNT + rsData.Fields("SETTI_YOUTO_ETC_CNT")      '設置用途_不明
            intTOTAL_CNT = intTOTAL_CNT + rsData.Fields("TOTAL_CNT")                                    '合計
            intSETTI_MI_CNT = intSETTI_MI_CNT + 0                                                       '設置未割付数
'↑ UPDATE 2014/05/02 MIYAMOTO

'↓ DELETE 2014/05/02 MIYAMOTO
'            '部屋数取得
'            strSQL = createContData(lngYARD)
'            Set rsContData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)
'↑ DELETE 2014/05/02 MIYAMOTO
            
'↓ UPDATE 2014/05/02 MIYAMOTO
'            '部屋情報設定
'            rsListData.Fields("部屋_ボックス") = Format(rsContData.Fields("BOX_CNT"), "##,##0")
'            rsListData.Fields("部屋_バイク") = Format(rsContData.Fields("BIKE_CNT"), "##,##0")
'            rsListData.Fields("部屋_屋外") = Format(rsContData.Fields("OUTDOOR_CNT"), "##,##0")
'            rsListData.Fields("部屋_室内") = Format(rsContData.Fields("INDOOR_CNT"), "##,##0")
'            rsListData.Fields("部屋_その他") = Format(rsContData.Fields("ETC_CNT"), "##,##0")
'            rsListData.Fields("部屋未割付数") = Format(0, "##,##0")
'
'            '合計行の集計
'            intROOM_BOX_CNT = intROOM_BOX_CNT + rsContData.Fields("BOX_CNT")                            '部屋_ボックス
'            intROOM_BIKE_CNT = intROOM_BIKE_CNT + rsContData.Fields("BIKE_CNT")                         '部屋_バイク
'            intROOM_SOTO_CNT = intROOM_SOTO_CNT + rsContData.Fields("OUTDOOR_CNT")                      '部屋_屋外
'            intROOM_UCHI_CNT = intROOM_UCHI_CNT + rsContData.Fields("INDOOR_CNT")                       '部屋_室内
'            intROOM_ETC_CNT = intROOM_ETC_CNT + rsContData.Fields("ETC_CNT")                            '部屋_その他
'            intROOM_MI_CNT = intROOM_MI_CNT + 0                                                         '部屋未割付数
'
'            rsContData.Close
'            Set rsContData = Nothing

            '部屋情報設定
            rsListData.Fields("部屋_ボックス") = Format(rsData.Fields("BOX_CNT"), "##,##0")
            rsListData.Fields("部屋_バイク") = Format(rsData.Fields("BIKE_CNT"), "##,##0")
            rsListData.Fields("部屋_屋外") = Format(rsData.Fields("OUTDOOR_CNT"), "##,##0")
            rsListData.Fields("部屋_室内") = Format(rsData.Fields("INDOOR_CNT"), "##,##0")
            rsListData.Fields("部屋_その他") = Format(rsData.Fields("ETC_CNT"), "##,##0")
            rsListData.Fields("部屋未割付数") = Format(0, "##,##0")

            '合計行の集計
            intROOM_BOX_CNT = intROOM_BOX_CNT + rsData.Fields("BOX_CNT")                                '部屋_ボックス
            intROOM_BIKE_CNT = intROOM_BIKE_CNT + rsData.Fields("BIKE_CNT")                             '部屋_バイク
            intROOM_SOTO_CNT = intROOM_SOTO_CNT + rsData.Fields("OUTDOOR_CNT")                          '部屋_屋外
            intROOM_UCHI_CNT = intROOM_UCHI_CNT + rsData.Fields("INDOOR_CNT")                           '部屋_室内
            intROOM_ETC_CNT = intROOM_ETC_CNT + rsData.Fields("ETC_CNT")                                '部屋_不明
            intROOM_MI_CNT = intROOM_MI_CNT + 0                                                         '部屋未割付数
'↑ UPDATE 2014/05/02 MIYAMOTO

            rsListData.UPDATE
            rsListData.Close
            Set rsListData = Nothing

            rsData.MoveNext
        Wend
    End If

    rsData.Close
    Set rsData = Nothing
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

    '合計行のセット
    Form_FKS830_S01.txt_合計_設置数_他社.VALUE = Format(intTASYA_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_設置数_自社.VALUE = Format(intJISYA_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_海上コンテナ_5Ft.VALUE = Format(intCONT5FT_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_海上コンテナ_6Ft.VALUE = Format(intCONT6FT_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_海上コンテナ_8Ft.VALUE = Format(intCONT8FT_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_海上コンテナ_10Ft.VALUE = Format(intCONT10FT_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_海上コンテナ_12Ft.VALUE = Format(intCONT12FT_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_海上コンテナ_20Ft.VALUE = Format(intCONT20FT_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_海上コンテナ_その他.VALUE = Format(intCONTETC_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_海上コンテナ_計.VALUE = Format(intCONTTTL_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_市販物置.VALUE = Format(intMONOOKI_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_ユニットハウス.VALUE = Format(intUHOUSE_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_荷物収納_その他.VALUE = Format(intNSYUNO_ETC_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_バイク収納.VALUE = Format(intBSYUNO_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_通路.VALUE = Format(intTUURO_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_外枠.VALUE = Format(intSOTOWAKU_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_自社資材等.VALUE = Format(intJSIZAI_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_看板.VALUE = Format(intKANBAN_CNT, "##,##0")                       'INSERT 2014/05/02 MIYAMOTO
    Form_FKS830_S01.txt_合計_設置用途_その他.VALUE = Format(intSETTI_YOUTO_ETC_CNT, "##,##0")
'    Form_FKS830_S01.txt_合計_設置数_その他.VALUE = Format(intETC_CNT, "##,##0")                'DELETE 2014/05/02 MIYAMOTO
    Form_FKS830_S01.txt_合計_合計.VALUE = Format(intTOTAL_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_設置未割付数.VALUE = Format(intSETTI_MI_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_部屋_ボックス.VALUE = Format(intROOM_BOX_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_部屋_バイク.VALUE = Format(intROOM_BIKE_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_部屋_屋外.VALUE = Format(intROOM_SOTO_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_部屋_室内.VALUE = Format(intROOM_UCHI_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_部屋_その他.VALUE = Format(intROOM_ETC_CNT, "##,##0")
    Form_FKS830_S01.txt_合計_部屋未割付数.VALUE = Format(intROOM_MI_CNT, "##,##0")

    Form_FKS830_S01.Refresh

    Exit Function

Exception:
    MsgBox "エラー発生" & vbCrLf & Err.Description & vbCrLf & debugCNT & "件目：ヤード = " & StrYard

    If Not rsData Is Nothing Then: rsData.Close: Set rsData = Nothing
    If Not rsListData Is Nothing Then: rsListData.Close: Set rsListData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    Call Err.Raise(Err.Number, "getSelectListData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :部屋数取得用SQL作成
'        MODULE_ID        :createContData
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'↓ UPDATE 2014/05/02 MIYAMOTO
'Private Function createContData(lngYARD As Long) As String
Private Function createContData() As String
'↑ UPDATE 2014/05/02 MIYAMOTO

    Dim strSQL                          As String

    strSQL = ""
'↓ DELETE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "SELECT      BOX.BOX_CNT" & vbCrLf
'    strSQL = strSQL & ",           BIKE.BIKE_CNT" & vbCrLf
'    strSQL = strSQL & ",           OUTDOOR.OUTDOOR_CNT" & vbCrLf
'    strSQL = strSQL & ",           INDOOR.INDOOR_CNT" & vbCrLf
'    strSQL = strSQL & ",           ETC.ETC_CNT" & vbCrLf
'↑ DELETE 2014/05/02 MIYAMOTO
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "FROM        (" & vbCrLf
'    strSQL = strSQL & "                /* ボックス */" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS BOX_CNT" & vbCrLf
    'ボックス
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CNTA_MAST.CNTA_CODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS BOX_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CNTA_MAST INNER JOIN NAME_MAST ON CNTA_MAST.CNTA_USAGE = NAME_MAST.NAME_CODE AND NAME_MAST.NAME_ID = '086'" & vbCrLf
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "                WHERE       CNTA_MAST.CNTA_CODE = " & lngYARD & vbCrLf
'    strSQL = strSQL & "                AND         NAME_MAST.NAME_CODE = 0" & vbCrLf
'    strSQL = strSQL & "            ) BOX" & vbCrLf
    strSQL = strSQL & "                WHERE       NAME_MAST.NAME_CODE = 0" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CNTA_MAST.CNTA_CODE" & vbCrLf
    strSQL = strSQL & "            ) BOX ON BOX.CNTA_CODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                /* バイク */" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS BIKE_CNT" & vbCrLf
    'バイク
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CNTA_MAST.CNTA_CODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS BIKE_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CNTA_MAST INNER JOIN NAME_MAST ON CNTA_MAST.CNTA_USAGE = NAME_MAST.NAME_CODE AND NAME_MAST.NAME_ID = '086'" & vbCrLf
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "                WHERE       CNTA_MAST.CNTA_CODE = " & lngYARD & vbCrLf
'    strSQL = strSQL & "                AND         NAME_MAST.NAME_CODE IN (3, 30, 31, 32, 33, 39)" & vbCrLf
'    strSQL = strSQL & "            ) BIKE" & vbCrLf
    strSQL = strSQL & "                WHERE       NAME_MAST.NAME_CODE IN (3, 30, 31, 32, 33, 39)" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CNTA_MAST.CNTA_CODE" & vbCrLf
    strSQL = strSQL & "            ) BIKE ON BIKE.CNTA_CODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                /* 屋外 */" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS OUTDOOR_CNT" & vbCrLf
    '屋外
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CNTA_MAST.CNTA_CODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS OUTDOOR_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CNTA_MAST INNER JOIN NAME_MAST ON CNTA_MAST.CNTA_USAGE = NAME_MAST.NAME_CODE AND NAME_MAST.NAME_ID = '086'" & vbCrLf
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "                WHERE       CNTA_MAST.CNTA_CODE = " & lngYARD & vbCrLf
'    strSQL = strSQL & "                AND         NAME_MAST.NAME_CODE = 31" & vbCrLf
'    strSQL = strSQL & "            ) OUTDOOR" & vbCrLf
    strSQL = strSQL & "                WHERE       NAME_MAST.NAME_CODE = 31" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CNTA_MAST.CNTA_CODE" & vbCrLf
    strSQL = strSQL & "            ) OUTDOOR ON OUTDOOR.CNTA_CODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                /* 室内 */" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS INDOOR_CNT" & vbCrLf
    '室内
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CNTA_MAST.CNTA_CODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS INDOOR_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CNTA_MAST INNER JOIN NAME_MAST ON CNTA_MAST.CNTA_USAGE = NAME_MAST.NAME_CODE AND NAME_MAST.NAME_ID = '086'" & vbCrLf
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "                WHERE       CNTA_MAST.CNTA_CODE = " & lngYARD & vbCrLf
'    strSQL = strSQL & "                AND         NAME_MAST.NAME_CODE = 32" & vbCrLf
'    strSQL = strSQL & "            ) INDOOR" & vbCrLf
    strSQL = strSQL & "                WHERE       NAME_MAST.NAME_CODE = 32" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CNTA_MAST.CNTA_CODE" & vbCrLf
    strSQL = strSQL & "            ) INDOOR ON INDOOR.CNTA_CODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                /* その他 */" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS ETC_CNT" & vbCrLf
    '不明
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CNTA_MAST.CNTA_CODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS ETC_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CNTA_MAST INNER JOIN NAME_MAST ON CNTA_MAST.CNTA_USAGE = NAME_MAST.NAME_CODE AND NAME_MAST.NAME_ID = '086'" & vbCrLf
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "                WHERE       CNTA_MAST.CNTA_CODE = " & lngYARD & vbCrLf
'    strSQL = strSQL & "                AND         NAME_MAST.NAME_CODE NOT IN (0, 3, 30, 31, 32, 33, 39)" & vbCrLf
'    strSQL = strSQL & "            ) ETC" & vbCrLf
    strSQL = strSQL & "                WHERE       NAME_MAST.NAME_CODE NOT IN (0, 3, 30, 31, 32, 33, 39)" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CNTA_MAST.CNTA_CODE" & vbCrLf
    strSQL = strSQL & "            ) ETC ON ETC.CNTA_CODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

'    strSQL = strSQL & ";" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO

    createContData = strSQL

End Function
'==============================================================================*
'
'        MODULE_NAME      :設置数取得用SQL作成
'        MODULE_ID        :createBoxData
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'↓ UPDATE 2014/05/02 MIYAMOTO
'Private Function createBoxData(strYard As String) As String
Private Function createBoxData() As String
'↑ UPDATE 2014/05/02 MIYAMOTO

    Dim strSQL                          As String

'↓ UPDATE 2014/05/02 MIYAMOTO
'    Dim strSYOYU(2)                     As String       '所有チェックボックス値格納用
'    Dim strKONYU(1)                     As String       '購入区分チェックボックス値格納用
'    Dim strCTYTO(5)                     As String       '設置用途チェックボックス値格納用
    Dim strSYOYU(1)                     As String       '所有チェックボックス値格納用
    Dim strKONYU(2)                     As String       '購入区分チェックボックス値格納用
    Dim strCTYTO(6)                     As String       '設置用途チェックボックス値格納用
'↑ UPDATE 2014/05/02 MIYAMOTO
    Dim strCTBNR(3)                     As String       'コンテナ分類チェックボックス値格納用

    Dim strJokenSTOCKDATE               As String       '入庫日条件文
    Dim strJokenSYOYU                   As String       '所有区分条件文
    Dim strJokenKONYU                   As String       '購入区分条件文
    Dim strJokenCTYTO                   As String       '設置用途条件文
    Dim strJokenCTBNR                   As String       'コンテナ分類条件文

    '<<入庫日_FROM>>
    strJokenSTOCKDATE = ""
    If Nz(Me.txt_入庫日_FROM.VALUE, "") <> "" Then
        strJokenSTOCKDATE = strJokenSTOCKDATE & "                AND         CBOX_MAST.CBOXM_STOCKDATE >= '" & Replace(Me.txt_入庫日_FROM.VALUE, "/", "") & "'" & vbCrLf
    End If

    '<<入庫日_TO>>
    If Nz(Me.txt_入庫日_TO.VALUE, "") <> "" Then
        strJokenSTOCKDATE = strJokenSTOCKDATE & "                AND         CBOX_MAST.CBOXM_STOCKDATE <= '" & Replace(Me.txt_入庫日_TO.VALUE, "/", "") & "'" & vbCrLf
    End If

    '<<所有>>
    Erase strSYOYU
    strJokenSYOYU = ""
    If Nz(Me.chk_所有区分_自社.VALUE, "") = -1 Then
        strSYOYU(0) = "'1'"
    Else
        strSYOYU(0) = "'-1'"
    End If
    If Nz(Me.chk_所有区分_他社.VALUE, "") = -1 Then
        strSYOYU(1) = "'2'"
    Else
        strSYOYU(1) = "'-2'"
    End If
'↓ UPDATE 2014/05/02 MIYAMOTO
'    If Nz(Me.chk_所有区分_その他.VALUE, "") = -1 Then
'        strSYOYU(2) = "'99'"
'    Else
'        strSYOYU(2) = "'-99'"
'    End If
'    strJokenSYOYU = "                AND        (CBOX_MAST.CBOXM_KRKBN IN (" & Replace(strSYOYU(0) & strSYOYU(1), "''", "', '") & ")" & vbCrLf
'    If strSYOYU(2) = "'99'" Then
'        strJokenSYOYU = strJokenSYOYU & "                    OR      CBOX_MAST.CBOXM_KRKBN NOT IN (" & Replace(Replace(strSYOYU(0), "-", "") & Replace(strSYOYU(1), "-", ""), "''", "', '") & "))" & vbCrLf
'    Else
'        strJokenSYOYU = strJokenSYOYU & "                    OR      CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(2) & ")" & vbCrLf
'    End If
    strJokenSYOYU = "                AND        CBOX_MAST.CBOXM_KRKBN IN (" & Replace(strSYOYU(0) & strSYOYU(1), "''", "', '") & ")" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '<<購入区分>>
    Erase strKONYU
    strJokenKONYU = ""
    If Nz(Me.chk_購入区分_新品.VALUE, "") = -1 Then
        strKONYU(0) = "'1'"
    Else
        strKONYU(0) = "'-1'"
    End If
    If Nz(Me.chk_購入区分_中古.VALUE, "") = -1 Then
        strKONYU(1) = "'2'"
    Else
        strKONYU(1) = "'-2'"
    End If
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strJokenKONYU = "                AND         CBOX_MAST.CBOXM_KNKBN IN (" & Replace(strKONYU(0) & strKONYU(1), "''", "', '") & ")" & vbCrLf
    If Nz(Me.chk_購入区分_不明.VALUE, "") = -1 Then
        strKONYU(2) = "'99'"
    Else
        strKONYU(2) = "'-99'"
    End If
    strJokenKONYU = "                AND        (CBOX_MAST.CBOXM_KNKBN IN (" & Replace(strKONYU(0) & strKONYU(1), "''", "', '") & ")" & vbCrLf
    If strKONYU(2) = "'99'" Then
        strJokenKONYU = strJokenKONYU & "                    OR      CBOX_MAST.CBOXM_KNKBN NOT IN (" & Replace(Replace(strKONYU(0), "-", "") & Replace(strKONYU(1), "-", ""), "''", "', '") & "))" & vbCrLf
    Else
        strJokenKONYU = strJokenKONYU & "                    OR      CBOX_MAST.CBOXM_KNKBN = " & strKONYU(2) & ")" & vbCrLf
    End If
'↑ UPDATE 2014/05/02 MIYAMOTO

    '<<設置用途>>
    Erase strCTYTO
    strJokenCTYTO = ""
    If Nz(Me.chk_設置用途_荷物収納.VALUE, "") = -1 Then
        strCTYTO(0) = "'1'"
    Else
        strCTYTO(0) = "'-1'"
    End If
    If Nz(Me.chk_設置用途_バイク収納.VALUE, "") = -1 Then
        strCTYTO(1) = "'2'"
    Else
        strCTYTO(1) = "'-2'"
    End If
    If Nz(Me.chk_設置用途_通路.VALUE, "") = -1 Then
        strCTYTO(2) = "'50'"
    Else
        strCTYTO(2) = "'-50'"
    End If
    If Nz(Me.chk_設置用途_外枠.VALUE, "") = -1 Then
        strCTYTO(3) = "'60'"
    Else
        strCTYTO(3) = "'-60'"
    End If
    If Nz(Me.chk_設置用途_自社資材等.VALUE, "") = -1 Then
        strCTYTO(4) = "'80'"
    Else
        strCTYTO(4) = "'-80'"
    End If
'↓ INSERT 2014/05/02 MIYAMOTO
    If Nz(Me.chk_設置用途_看板.VALUE, "") = -1 Then
        strCTYTO(5) = "'90'"
    Else
        strCTYTO(5) = "'-90'"
    End If
'↑ INSERT 2014/05/02 MIYAMOTO
'↓ UPDATE 2014/05/02 MIYAMOTO
'    If Nz(Me.chk_設置用途_その他.VALUE, "") = -1 Then
'        strCTYTO(5) = "'99'"
'    Else
'        strCTYTO(5) = "'-99'"
'    End If
'    strJokenCTYTO = "                AND        (CBOX_MAST.CBOXM_CTYTO IN (" & Replace(strCTYTO(0) & strCTYTO(1) & strCTYTO(2) & strCTYTO(3) & strCTYTO(4), "''", "', '") & ")" & vbCrLf
'    If strCTYTO(5) = "'99'" Then
'        strJokenCTYTO = strJokenCTYTO & "                    OR      CBOX_MAST.CBOXM_CTYTO NOT IN (" & Replace(Replace(strCTYTO(0), "-", "") & Replace(strCTYTO(1), "-", "") & Replace(strCTYTO(2), "-", "") & Replace(strCTYTO(3), "-", "") & Replace(strCTYTO(4), "-", ""), "''", "', '") & "))" & vbCrLf
'    Else
'        strJokenCTYTO = strJokenCTYTO & "                    OR      CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(5) & ")" & vbCrLf
'    End If
    If Nz(Me.chk_設置用途_その他.VALUE, "") = -1 Then
        strCTYTO(6) = "'99'"
    Else
        strCTYTO(6) = "'-99'"
    End If
    strJokenCTYTO = "                AND        (CBOX_MAST.CBOXM_CTYTO IN (" & Replace(strCTYTO(0) & strCTYTO(1) & strCTYTO(2) & strCTYTO(3) & strCTYTO(4) & strCTYTO(5), "''", "', '") & ")" & vbCrLf
    If strCTYTO(6) = "'99'" Then
        strJokenCTYTO = strJokenCTYTO & "                    OR      CBOX_MAST.CBOXM_CTYTO NOT IN (" & Replace(Replace(strCTYTO(0), "-", "") & Replace(strCTYTO(1), "-", "") & Replace(strCTYTO(2), "-", "") & Replace(strCTYTO(3), "-", "") & Replace(strCTYTO(4), "-", "") & Replace(strCTYTO(5), "-", ""), "''", "', '") & "))" & vbCrLf
    Else
        strJokenCTYTO = strJokenCTYTO & "                    OR      CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(6) & ")" & vbCrLf
    End If
'↑ UPDATE 2014/05/02 MIYAMOTO

    '<<コンテナ分類>>
    Erase strCTBNR
    strJokenCTBNR = ""
    If Nz(Me.chk_コンテナ分類_海上コンテナ.VALUE, "") = -1 Then
        strCTBNR(0) = "'1'"
    Else
        strCTBNR(0) = "'-1'"
    End If
    If Nz(Me.chk_コンテナ分類_物置.VALUE, "") = -1 Then
        strCTBNR(1) = "'40'"
    Else
        strCTBNR(1) = "'-40'"
    End If
    If Nz(Me.chk_コンテナ分類_ユニットハウス.VALUE, "") = -1 Then
        strCTBNR(2) = "'50'"
    Else
        strCTBNR(2) = "'-50'"
    End If
    If Nz(Me.chk_コンテナ分類_その他.VALUE, "") = -1 Then
        strCTBNR(3) = "'99'"
    Else
        strCTBNR(3) = "'-99'"
    End If
    strJokenCTBNR = "                AND        ((CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(0) & " AND (CBOX_MAST.CBOXM_CTBNR IN (" & Replace(strCTBNR(0) & strCTBNR(1) & strCTBNR(2), "''", "', '") & ")"
    If strCTBNR(3) = "'99'" Then
        strJokenCTBNR = strJokenCTBNR & " OR CBOX_MAST.CBOXM_CTBNR NOT IN (" & Replace(Replace(strCTBNR(0), "-", "") & Replace(strCTBNR(1), "-", "") & Replace(strCTBNR(2), "-", ""), "''", "', '") & ")))" & vbCrLf
    Else
        strJokenCTBNR = strJokenCTBNR & " OR CBOX_MAST.CBOXM_CTBNR = " & strCTBNR(3) & "))" & vbCrLf
    End If
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strJokenCTBNR = strJokenCTBNR & "                    OR      (CBOX_MAST.CBOXM_CTYTO IN (" & Replace(strCTYTO(1) & strCTYTO(2) & strCTYTO(3) & strCTYTO(4), "''", "', '") & ")"
'    If strCTYTO(5) = "'99'" Then
'        strJokenCTBNR = strJokenCTBNR & " OR CBOX_MAST.CBOXM_CTYTO NOT IN (" & Replace(Replace(strCTYTO(0), "-", "") & Replace(strCTYTO(1), "-", "") & Replace(strCTYTO(2), "-", "") & Replace(strCTYTO(3), "-", "") & Replace(strCTYTO(4), "-", ""), "''", "', '") & ")))" & vbCrLf
'    Else
'        strJokenCTBNR = strJokenCTBNR & " OR CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(5) & "))" & vbCrLf
'    End If
    strJokenCTBNR = strJokenCTBNR & "                    OR      (CBOX_MAST.CBOXM_CTYTO IN (" & Replace(strCTYTO(1) & strCTYTO(2) & strCTYTO(3) & strCTYTO(4) & strCTYTO(5), "''", "', '") & ")"
    If strCTYTO(6) = "'99'" Then
        strJokenCTBNR = strJokenCTBNR & " OR CBOX_MAST.CBOXM_CTYTO NOT IN (" & Replace(Replace(strCTYTO(0), "-", "") & Replace(strCTYTO(1), "-", "") & Replace(strCTYTO(2), "-", "") & Replace(strCTYTO(3), "-", "") & Replace(strCTYTO(4), "-", "") & Replace(strCTYTO(5), "-", ""), "''", "', '") & ")))" & vbCrLf
    Else
        strJokenCTBNR = strJokenCTBNR & " OR CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(6) & "))" & vbCrLf
    End If
'↑ UPDATE 2014/05/02 MIYAMOTO

    strSQL = ""
'↓ DELETE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "SELECT      TASYA.TASYA_CNT" & vbCrLf                        '設置数_他社
'    strSQL = strSQL & ",           JISYA.JISYA_CNT" & vbCrLf                        '設置数_自社
'    strSQL = strSQL & ",           CONT5FT.CONT5FT_CNT" & vbCrLf                    '海上コンテナ5Ft
'    strSQL = strSQL & ",           CONT6FT.CONT6FT_CNT" & vbCrLf                    '海上コンテナ6Ft
'    strSQL = strSQL & ",           CONT8FT.CONT8FT_CNT" & vbCrLf                    '海上コンテナ8Ft
'    strSQL = strSQL & ",           CONT10FT.CONT10FT_CNT" & vbCrLf                  '海上コンテナ10Ft
'    strSQL = strSQL & ",           CONT12FT.CONT12FT_CNT" & vbCrLf                  '海上コンテナ12Ft
'    strSQL = strSQL & ",           CONT20FT.CONT20FT_CNT" & vbCrLf                  '海上コンテナ20Ft
'    strSQL = strSQL & ",           CONTETC.CONTETC_CNT" & vbCrLf                    '海上コンテナその他
'    strSQL = strSQL & ",           CONTTTL.CONTTTL_CNT" & vbCrLf                    '海上コンテナ計
'    strSQL = strSQL & ",           MONOOKI.MONOOKI_CNT" & vbCrLf                    '市販物置
'    strSQL = strSQL & ",           UHOUSE.UHOUSE_CNT" & vbCrLf                      'ユニットハウス
'    strSQL = strSQL & ",           NSYUNO_ETC.NSYUNO_ETC_CNT" & vbCrLf              '荷物収納_その他
'    strSQL = strSQL & ",           BSYUNO.BSYUNO_CNT" & vbCrLf                      'バイク収納
'    strSQL = strSQL & ",           TUURO.TUURO_CNT" & vbCrLf                        '通路
'    strSQL = strSQL & ",           SOTOWAKU.SOTOWAKU_CNT" & vbCrLf                  '外枠
'    strSQL = strSQL & ",           JSIZAI.JSIZAI_CNT" & vbCrLf                      '自社資材等
'    strSQL = strSQL & ",           KANBAN.KANBAN_CNT" & vbCrLf                      '看板                   'INSERT 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           TOTAL.TOTAL_CNT" & vbCrLf                        '合計
'    strSQL = strSQL & ",           SETTI_YOUTO_ETC.SETTI_YOUTO_ETC_CNT" & vbCrLf    '荷物収納_その他
'    strSQL = strSQL & ",           ETC.ETC_CNT" & vbCrLf                            '設置数_その他
'↑ DELETE 2014/05/02 MIYAMOTO

    '設置数_他社
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "FROM        (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS TASYA_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS TASYA_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(1) & vbCrLf     '所有区分=他社
    If Nz(strJokenSTOCKDATE, "") <> "" Then '入庫日
        strSQL = strSQL & strJokenSTOCKDATE
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    If Nz(strJokenCTYTO, "") <> "" Then
        strSQL = strSQL & strJokenCTYTO     '設置用途
    End If
    If Nz(strJokenCTBNR, "") <> "" Then
        strSQL = strSQL & strJokenCTBNR     'コンテナ分類
    End If
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) TASYA" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) TASYA ON TASYA.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '設置数_自社
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS JISYA_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS JISYA_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    If Nz(strJokenCTYTO, "") <> "" Then
        strSQL = strSQL & strJokenCTYTO     '設置用途
    End If
    If Nz(strJokenCTBNR, "") <> "" Then
        strSQL = strSQL & strJokenCTBNR     'コンテナ分類
    End If
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) JISYA" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) JISYA ON JISYA.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '海上コンテナ5Ft
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS CONT5FT_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS CONT5FT_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(0) & vbCrLf     '設置用途=荷物収納
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTBNR = " & strCTBNR(0) & vbCrLf     'コンテナ分類=海上コンテナ(設置用途=荷物収納)の時のみ)
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_SIZE = '5'" & vbCrLf                 'サイズ=5Ft(コンテナ分類=海上コンテナ)の時のみ)
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) CONT5FT" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) CONT5FT ON CONT5FT.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '海上コンテナ6Ft
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS CONT6FT_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS CONT6FT_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(0) & vbCrLf     '設置用途=荷物収納
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTBNR = " & strCTBNR(0) & vbCrLf     'コンテナ分類=海上コンテナ(設置用途=荷物収納)の時のみ)
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_SIZE = '6'" & vbCrLf                 'サイズ=6Ft(コンテナ分類=海上コンテナ)の時のみ)
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) CONT6FT" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) CONT6FT ON CONT6FT.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '海上コンテナ8Ft
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS CONT8FT_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS CONT8FT_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(0) & vbCrLf     '設置用途=荷物収納
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTBNR = " & strCTBNR(0) & vbCrLf     'コンテナ分類=海上コンテナ(設置用途=荷物収納)の時のみ)
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_SIZE = '8'" & vbCrLf                 'サイズ=8Ft(コンテナ分類=海上コンテナ)の時のみ)
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) CONT8FT" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) CONT8FT ON CONT8FT.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '海上コンテナ10Ft
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS CONT10FT_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS CONT10FT_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(0) & vbCrLf     '設置用途=荷物収納
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTBNR = " & strCTBNR(0) & vbCrLf     'コンテナ分類=海上コンテナ(設置用途=荷物収納)の時のみ)
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_SIZE = '10'" & vbCrLf                'サイズ=10Ft(コンテナ分類=海上コンテナ)の時のみ)
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) CONT10FT" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) CONT10FT ON CONT10FT.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '海上コンテナ12Ft
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS CONT12FT_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS CONT12FT_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(0) & vbCrLf     '設置用途=荷物収納
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTBNR = " & strCTBNR(0) & vbCrLf     'コンテナ分類=海上コンテナ(設置用途=荷物収納)の時のみ)
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_SIZE = '12'" & vbCrLf                'サイズ=12Ft(コンテナ分類=海上コンテナ)の時のみ)
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) CONT12FT" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) CONT12FT ON CONT12FT.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '海上コンテナ20Ft
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS CONT20FT_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS CONT20FT_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(0) & vbCrLf     '設置用途=荷物収納
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTBNR = " & strCTBNR(0) & vbCrLf     'コンテナ分類=海上コンテナ(設置用途=荷物収納)の時のみ)
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_SIZE = '20'" & vbCrLf                'サイズ=20Ft(コンテナ分類=海上コンテナ)の時のみ)
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) CONT20FT" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) CONT20FT ON CONT20FT.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '海上コンテナその他
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS CONTETC_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS CONTETC_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(0) & vbCrLf     '設置用途=荷物収納
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTBNR = " & strCTBNR(0) & vbCrLf     'コンテナ分類=海上コンテナ(設置用途=荷物収納)の時のみ)
'↓ CHG 2018/09/07 tajima
'   strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_SIZE = '99'" & vbCrLf                'サイズ=－(コンテナ分類=海上コンテナ)の時のみ)
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_SIZE not in('5','6','8','10','12','20')" & vbCrLf  'サイズ=メジャーサイズ以外(コンテナ分類=海上コンテナ)の時のみ)
'↑ CHG 2018/09/07 tajima

'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) CONTETC" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) CONTETC ON CONTETC.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '海上コンテナ計
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS CONTTTL_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS CONTTTL_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(0) & vbCrLf     '設置用途=荷物収納
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTBNR = " & strCTBNR(0) & vbCrLf     'コンテナ分類=海上コンテナ(設置用途=荷物収納)の時のみ)
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) CONTTTL" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) CONTTTL ON CONTTTL.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '市販物置
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS MONOOKI_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS MONOOKI_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(0) & vbCrLf     '設置用途=荷物収納
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTBNR = " & strCTBNR(1) & vbCrLf     'コンテナ分類=物置(設置用途=荷物収納)の時のみ)
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) MONOOKI" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) MONOOKI ON MONOOKI.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    'ユニットハウス
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS UHOUSE_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS UHOUSE_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(0) & vbCrLf     '設置用途=荷物収納
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTBNR = " & strCTBNR(2) & vbCrLf     'コンテナ分類=ユニットハウス(設置用途=荷物収納)の時のみ)
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) UHOUSE" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) UHOUSE ON UHOUSE.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '荷物収納_その他
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS NSYUNO_ETC_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS NSYUNO_ETC_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(0) & vbCrLf     '設置用途=荷物収納
    If strCTBNR(3) = "'99'" Then
        strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTBNR NOT IN (" & Replace(Replace(strCTBNR(0), "-", "") & Replace(strCTBNR(1), "-", "") & Replace(strCTBNR(2), "-", ""), "''", "', '") & ")" & vbCrLf
    Else
        strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTBNR = " & strCTBNR(3) & vbCrLf
    End If
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) NSYUNO_ETC" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) NSYUNO_ETC ON NSYUNO_ETC.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    'バイク収納
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS BSYUNO_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS BSYUNO_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(1) & vbCrLf     '設置用途=バイク収納
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) BSYUNO" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) BSYUNO ON BSYUNO.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '通路
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS TUURO_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS TUURO_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(2) & vbCrLf     '設置用途=通路
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) TUURO" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) TUURO ON TUURO.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '外枠
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS SOTOWAKU_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS SOTOWAKU_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(3) & vbCrLf     '設置用途=外枠
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) SOTOWAKU" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) SOTOWAKU ON SOTOWAKU.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

    '自社資材等
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS JSIZAI_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS JSIZAI_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(4) & vbCrLf     '設置用途=自社資材等
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) JSIZAI" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) JSIZAI ON JSIZAI.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

'↓ INSERT 2014/05/02 MIYAMOTO
    '看板
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS KANBAN_CNT" & vbCrLf
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(5) & vbCrLf     '設置用途=看板
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) KANBAN ON KANBAN.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ INSERT 2014/05/02 MIYAMOTO

    '設置用途_その他
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS SETTI_YOUTO_ETC_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS SETTI_YOUTO_ETC_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(0) & vbCrLf     '所有区分=自社
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
'↓ UPDATE 2014/05/02 MIYAMOTO
'    If strCTYTO(5) = "'99'" Then
'        strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO NOT IN (" & Replace(Replace(strCTYTO(0), "-", "") & Replace(strCTYTO(1), "-", "") & Replace(strCTYTO(2), "-", "") & Replace(strCTYTO(3), "-", "") & Replace(strCTYTO(4), "-", ""), "''", "', '") & ")" & vbCrLf
'    Else
'        strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(5) & vbCrLf
'    End If
'    strSQL = strSQL & "            ) SETTI_YOUTO_ETC" & vbCrLf
    If strCTYTO(6) = "'99'" Then
        strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO NOT IN (" & Replace(Replace(strCTYTO(0), "-", "") & Replace(strCTYTO(1), "-", "") & Replace(strCTYTO(2), "-", "") & Replace(strCTYTO(3), "-", "") & Replace(strCTYTO(4), "-", "") & Replace(strCTYTO(5), "-", ""), "''", "', '") & ")" & vbCrLf
    Else
        strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_CTYTO = " & strCTYTO(6) & vbCrLf
    End If
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) SETTI_YOUTO_ETC ON SETTI_YOUTO_ETC.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

'↓ DELETE 2014/05/02 MIYAMOTO
'    '設置数_その他
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS ETC_CNT" & vbCrLf
'    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
'    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf
'    If strSYOYU(2) = "'99'" Then
'        strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN NOT IN (" & Replace(Replace(strSYOYU(0), "-", "") & Replace(strSYOYU(1), "-", ""), "''", "', '") & ")" & vbCrLf
'    Else
'        strSQL = strSQL & "                AND         CBOX_MAST.CBOXM_KRKBN = " & strSYOYU(2) & vbCrLf
'    End If
'    If Nz(strJokenSTOCKDATE, "") <> "" Then
'        strSQL = strSQL & strJokenSTOCKDATE '入庫日
'    End If
'    If Nz(strJokenKONYU, "") <> "" Then
'        strSQL = strSQL & strJokenKONYU     '購入区分
'    End If
'    If Nz(strJokenCTYTO, "") <> "" Then
'        strSQL = strSQL & strJokenCTYTO     '設置用途
'    End If
'    If Nz(strJokenCTBNR, "") <> "" Then
'        strSQL = strSQL & strJokenCTBNR     'コンテナ分類
'    End If
'    strSQL = strSQL & "            ) ETC" & vbCrLf
'↑ DELETE 2014/05/02 MIYAMOTO

    '合計
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & ",           (" & vbCrLf
'    strSQL = strSQL & "                SELECT      COUNT(*) AS TOTAL_CNT" & vbCrLf
    strSQL = strSQL & "            LEFT OUTER JOIN" & vbCrLf
    strSQL = strSQL & "            (" & vbCrLf
    strSQL = strSQL & "                SELECT      CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "                ,           COUNT(*) AS TOTAL_CNT" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO
    strSQL = strSQL & "                FROM        CBYS_TRAN INNER JOIN CBOX_MAST ON CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO AND CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO" & vbCrLf
    strSQL = strSQL & "                WHERE       CBYS_TRAN.CBYST_SYUBETSU = '0'" & vbCrLf
'    strSQL = strSQL & "                AND         CBYS_TRAN.CBYST_YCODE = '" & strYard & "'" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO
    If Nz(strJokenSTOCKDATE, "") <> "" Then
        strSQL = strSQL & strJokenSTOCKDATE '入庫日
    End If
    If Nz(strJokenSYOYU, "") <> "" Then
        strSQL = strSQL & strJokenSYOYU     '所有区分
    End If
    If Nz(strJokenKONYU, "") <> "" Then
        strSQL = strSQL & strJokenKONYU     '購入区分
    End If
    If Nz(strJokenCTYTO, "") <> "" Then
        strSQL = strSQL & strJokenCTYTO     '設置用途
    End If
    If Nz(strJokenCTBNR, "") <> "" Then
        strSQL = strSQL & strJokenCTBNR     'コンテナ分類
    End If
'↓ UPDATE 2014/05/02 MIYAMOTO
'    strSQL = strSQL & "            ) TOTAL" & vbCrLf
    strSQL = strSQL & "                GROUP BY    CBYS_TRAN.CBYST_YCODE" & vbCrLf
    strSQL = strSQL & "            ) TOTAL ON TOTAL.CBYST_YCODE = YARD_MAST.YARD_CODE" & vbCrLf
'↑ UPDATE 2014/05/02 MIYAMOTO

'    strSQL = strSQL & ";" & vbCrLf     'DELETE 2014/05/02 MIYAMOTO

    createBoxData = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :検索条件入力チェック
'        MODULE_ID        :fncSearchInputCheck
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSearchInputCheck() As Boolean

On Error GoTo Exception

    fncSearchInputCheck = False

    If (Nz(Me.cmb_ヤード_FROM.VALUE, "") <> "" And Nz(Me.cmb_ヤード_TO.VALUE, "") <> "") Then
        If Me.cmb_ヤード_FROM.VALUE > Me.cmb_ヤード_TO.VALUE Then
            '大小チェックエラー
            MsgBox "ヤードの範囲指定が矛盾しています。" & vbCrLf & "From <= To となるように入力して下さい。", vbInformation, P_ERROR
            Me.cmb_ヤード_FROM.SetFocus
            Exit Function
        End If
    End If

    If (Nz(Me.txt_入庫日_FROM.VALUE, "") <> "" And Nz(Me.txt_入庫日_TO.VALUE, "") <> "") Then
        If Me.txt_入庫日_FROM.VALUE > Me.txt_入庫日_TO.VALUE Then
            '大小チェックエラー
            MsgBox "入庫日の範囲指定が矛盾しています。" & vbCrLf & "From <= To となるように入力して下さい。", vbInformation, P_ERROR
            Me.txt_入庫日_FROM.SetFocus
            Exit Function
        End If
    End If

    If (Nz(Me.txt_営業開始日_FROM.VALUE, "") <> "" And Nz(Me.txt_営業開始日_TO.VALUE, "") <> "") Then
        If Me.txt_営業開始日_FROM.VALUE > Me.txt_営業開始日_TO.VALUE Then
            '大小チェックエラー
            MsgBox "営業開始日の範囲指定が矛盾しています。" & vbCrLf & "From <= To となるように入力して下さい。", vbInformation, P_ERROR
            Me.txt_営業開始日_FROM.SetFocus
            Exit Function
        End If
    End If

    If (Nz(Me.txt_営業終了日_FROM.VALUE, "") <> "" And Nz(Me.txt_営業終了日_TO.VALUE, "") <> "") Then
        If Me.txt_営業終了日_FROM.VALUE > Me.txt_営業終了日_TO.VALUE Then
            '大小チェックエラー
            MsgBox "営業終了日の範囲指定が矛盾しています。" & vbCrLf & "From <= To となるように入力して下さい。", vbInformation, P_ERROR
            Me.txt_営業終了日_FROM.SetFocus
            Exit Function
        End If
    End If

    fncSearchInputCheck = True

    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncSearchInputCheck" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

Private Sub Form_Open(Cancel As Integer)
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_標題.BackColor = RGB(255, 0, 204)
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_営業開始日_FROM_AfterUpdate
'        MODULE_ID        :txt_営業開始日_FROM_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_営業開始日_FROM_AfterUpdate()

    Me.txt_営業開始日_FROM.VALUE = pstrDateValue

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_営業開始日_FROM_BeforeUpdate
'        MODULE_ID        :txt_営業開始日_FROM_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_営業開始日_FROM_BeforeUpdate(Cancel As Integer)

    If Nz(Me.txt_営業開始日_FROM.VALUE) <> "" Then
        pstrDateValue = Nz(CmDateChecker(Nz(Me.txt_営業開始日_FROM.VALUE)))
        If pstrDateValue = "" Then
            MsgBox "日付を入力して下さい。（YYYY/MM/DD）"
            Cancel = True
            SendKeys "{ESC}"
        End If
    Else
        pstrDateValue = ""
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_営業開始日_TO_AfterUpdate
'        MODULE_ID        :txt_営業開始日_TO_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_営業開始日_TO_AfterUpdate()

    Me.txt_営業開始日_TO.VALUE = pstrDateValue

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_営業開始日_TO_BeforeUpdate
'        MODULE_ID        :txt_営業開始日_TO_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_営業開始日_TO_BeforeUpdate(Cancel As Integer)

    If Nz(Me.txt_営業開始日_TO.VALUE) <> "" Then
        pstrDateValue = Nz(CmDateChecker(Nz(Me.txt_営業開始日_TO.VALUE)))
        If pstrDateValue = "" Then
            MsgBox "日付を入力して下さい。（YYYY/MM/DD）"
            Cancel = True
            SendKeys "{ESC}"
        End If
    Else
        pstrDateValue = ""
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_営業終了日_FROM_AfterUpdate
'        MODULE_ID        :txt_営業終了日_FROM_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_営業終了日_FROM_AfterUpdate()

    Me.txt_営業終了日_FROM.VALUE = pstrDateValue

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_営業終了日_FROM_BeforeUpdate
'        MODULE_ID        :txt_営業終了日_FROM_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_営業終了日_FROM_BeforeUpdate(Cancel As Integer)

    If Nz(Me.txt_営業終了日_FROM.VALUE) <> "" Then
        pstrDateValue = Nz(CmDateChecker(Nz(Me.txt_営業終了日_FROM.VALUE)))
        If pstrDateValue = "" Then
            MsgBox "日付を入力して下さい。（YYYY/MM/DD）"
            Cancel = True
            SendKeys "{ESC}"
        End If
    Else
        pstrDateValue = ""
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_営業終了日_TO_AfterUpdate
'        MODULE_ID        :txt_営業終了日_TO_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_営業終了日_TO_AfterUpdate()

    Me.txt_営業終了日_TO.VALUE = pstrDateValue

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_営業終了日_TO_BeforeUpdate
'        MODULE_ID        :txt_営業終了日_TO_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_営業終了日_TO_BeforeUpdate(Cancel As Integer)

    If Nz(Me.txt_営業終了日_TO.VALUE) <> "" Then
        pstrDateValue = Nz(CmDateChecker(Nz(Me.txt_営業終了日_TO.VALUE)))
        If pstrDateValue = "" Then
            MsgBox "日付を入力して下さい。（YYYY/MM/DD）"
            Cancel = True
            SendKeys "{ESC}"
        End If
    Else
        pstrDateValue = ""
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_解約予定日_AfterUpdate
'        MODULE_ID        :txt_解約予定日_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約予定日_AfterUpdate()

    Me.txt_解約予定日.VALUE = pstrDateValue

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_解約予定日_BeforeUpdate
'        MODULE_ID        :txt_解約予定日_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約予定日_BeforeUpdate(Cancel As Integer)

    If Nz(Me.txt_解約予定日.VALUE) <> "" Then
        pstrDateValue = Nz(CmDateChecker(Nz(Me.txt_解約予定日.VALUE)))
        If pstrDateValue = "" Then
            MsgBox "日付を入力して下さい。（YYYY/MM/DD）"
            Cancel = True
            SendKeys "{ESC}"
        End If
    Else
        pstrDateValue = ""
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_入庫日_FROM_AfterUpdate
'        MODULE_ID        :txt_入庫日_FROM_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_入庫日_FROM_AfterUpdate()

    Me.txt_入庫日_FROM.VALUE = pstrDateValue

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_入庫日_FROM_BeforeUpdate
'        MODULE_ID        :txt_入庫日_FROM_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_入庫日_FROM_BeforeUpdate(Cancel As Integer)

    If Nz(Me.txt_入庫日_FROM.VALUE) <> "" Then
        pstrDateValue = Nz(CmDateChecker(Nz(Me.txt_入庫日_FROM.VALUE)))
        If pstrDateValue = "" Then
            MsgBox "日付を入力して下さい。（YYYY/MM/DD）"
            Cancel = True
            SendKeys "{ESC}"
        End If
    Else
        pstrDateValue = ""
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_入庫日_TO_AfterUpdate
'        MODULE_ID        :txt_入庫日_TO_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_入庫日_TO_AfterUpdate()

    Me.txt_入庫日_TO.VALUE = pstrDateValue

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_入庫日_TO_BeforeUpdate
'        MODULE_ID        :txt_入庫日_TO_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_入庫日_TO_BeforeUpdate(Cancel As Integer)

    If Nz(Me.txt_入庫日_TO.VALUE) <> "" Then
        pstrDateValue = Nz(CmDateChecker(Nz(Me.txt_入庫日_TO.VALUE)))
        If pstrDateValue = "" Then
            MsgBox "日付を入力して下さい。（YYYY/MM/DD）"
            Cancel = True
            SendKeys "{ESC}"
        End If
    Else
        pstrDateValue = ""
    End If

End Sub
'****************************  ended or program ********************************
