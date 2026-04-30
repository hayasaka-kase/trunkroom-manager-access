VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS450"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　価格テンプレート設定ヤード一覧
'   プログラムＩＤ　：　FKS450
'   作　成　日　　　：  2017/11/01
'   作　成　者　　　：  イーグルソフト Y
'   Ver             ：  1.0
'**********************************************
'修正履歴
'   更　新　日　　　：  2017/12/01
'   更　新　者　　　：  イーグルソフト Y
'   修　正　内　容　：　キャンペーン一括設定対応
'   Ver             ：  1.1
'**********************************************
'修正履歴
'   更　新　日　　　：  2018/01/26
'   更　新　者　　　：  イーグルソフト
'   修　正　内　容　：　test_kase_db 残し漏れ
'   Ver             ：  1.2
'**********************************************
'修正履歴
'   更　新　日　　　：  2018/08/11
'   更　新　者　　　：  イーグルソフト Y
'   修　正　内　容　：　事務手数料対応
'   Ver             ：  1.3
'**********************************************
'修正履歴
'   更　新　日　　　：  2018/11/15
'   更　新　者　　　：  tajima@eagle-soft
'   修　正　内　容　：　適用満了月数の追加
'   Ver             ：  1.4
'**********************************************
'修正履歴
'   更　新　日　　　：  2019/02/25
'   更　新　者　　　：  M.HONDA
'   修　正　内　容　：　割引の設定条件を変更
'   Ver             ：  1.5
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'接続ＤＢオブジェクト
Private dbSQLServer     As Database
Private dbAccess        As Database

'接続ＤＢ名
Private strSqlserver    As String

'接続文字列
Private strConnect      As String

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FKS450"

Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

Private Const P_SUB_FORM            As String = "FKS450_S01"
Private Const P_WORK_TABLE          As String = "FKS450_W01"
Private Const P_WORK_EXCEL          As String = "FKS450_Q01"

Private P_USER_ID                   As String
Private P_BUMON_CODE                As String
Private P_DATE_VALUE                As String
Private P_WARIBIKI_CODE(0 To 5)     As String

Public p450_RET                     As Boolean
Public p450_WaribikiNo              As String
Public p450_Youto                   As String
Public p450_Mongon                  As String

Private tab_Kensaku_flg             As Boolean  '2018/08/11 検索パネルフラグ

'==============================================================================*
'
'        MODULE_NAME      :cmb_Kakaku_Ikatu_Setei_AfterUpdate
'        機能             :事務手数料一括設定コンボボックス更新
'        CREATE BY        : 2018/08/11
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Jimute_Ikatu_Setei_AfterUpdate()
    Dim strSQL          As String
On Error GoTo Exception
    If dbAccess Is Nothing Then: Set dbAccess = CurrentDb
    strSQL = "update " & P_WORK_TABLE & " set KAKAKU_KKTHM_TCODE = '" & Me.cmb_Jimute_Ikatu_Setei.Column(1, cmb_Jimute_Ikatu_Setei.ListIndex) & "'"
    dbAccess.Execute strSQL, dbFailOnError
    FKS450_S01.Requery
    Exit Sub
Exception:
    Call MSZZ024_M00("cmb_Kakaku_Ikatu_Setei_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Kakaku_Ikatu_Setei_AfterUpdate
'        機能             :価格一括設定コンボボックス更新
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Kakaku_Ikatu_Setei_AfterUpdate()
    Dim strSQL          As String
On Error GoTo Exception
    If dbAccess Is Nothing Then: Set dbAccess = CurrentDb
    strSQL = "update " & P_WORK_TABLE & " set KAKAKU_KKTHM_TCODE = '" & Me.cmb_Kakaku_Ikatu_Setei.Column(1, cmb_Kakaku_Ikatu_Setei.ListIndex) & "'"
    dbAccess.Execute strSQL, dbFailOnError
    FKS450_S01.Requery
    Exit Sub
Exception:
    Call MSZZ024_M00("cmb_Kakaku_Ikatu_Setei_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncMakeWork
'        機能             :ワークテーブル作成
'        IN               :True=該当データ無しメッセージを表示 False=表示しない
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncKadourituChk(objRstLst As Object) As Boolean
    Dim intKadoRitu As Integer
    fncKadourituChk = False
    '稼働率
    If objRstLst.Fields("KADOU_BUNBO") <> 0 Then
        intKadoRitu = Int((objRstLst.Fields("KADOU_BUNSI") / objRstLst.Fields("KADOU_BUNBO")) * 100)
    Else
        intKadoRitu = 0
    End If
    '無条件チェック
    If Nz(cmb_Kadoritu.Column(1, cmb_Kadoritu.ListIndex), "X") = "X" Then
        'OK
        fncKadourituChk = True
        Exit Function
    End If
    '稼働率チェック
    If intKadoRitu >= Val(Nz(cmb_Kadoritu.Column(1, cmb_Kadoritu.ListIndex), 0)) And intKadoRitu <= Val(Nz(cmb_Kadoritu.Column(2, cmb_Kadoritu.ListIndex), 0)) Then
        'OK
        fncKadourituChk = True
        Exit Function
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :cmb_Waribiki_Ikatu_Setei_AfterUpdate
'        機能             :割引一括設定コンボボックス更新
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Waribiki_Ikatu_Setei_AfterUpdate()
    Dim strSQL          As String
On Error GoTo Exception
    If dbAccess Is Nothing Then: Set dbAccess = CurrentDb
    strSQL = "update " & P_WORK_TABLE & " set WARIBIKI_KKTHM_TCODE = '" & Me.cmb_Waribiki_Ikatu_Setei.Column(1, cmb_Waribiki_Ikatu_Setei.ListIndex) & "'"
    dbAccess.Execute strSQL, dbFailOnError
    FKS450_S01.Requery
    Exit Sub
Exception:
    Call MSZZ024_M00("cmb_Waribiki_Ikatu_Setei_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_from_AfterUpdate
'        機能             :ヤードコードFROMの更新後
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_from_AfterUpdate()
On Error GoTo Exception
    Me.txt_Yard_From = ""
    ' ヤードコード取得
    If Nz(Me.cmb_Yard_from, "") = "" Then Exit Sub
    If IsNull(Me.cmb_Yard_from.Column(1)) Then
        
        MsgBox "該当するヤードが登録がされておりません。", vbInformation, P_ERROR
        Exit Sub
    End If
    Me.txt_Yard_From = Me.cmb_Yard_from.Column(1)
    Me.cmb_Yard_To.SetFocus
    Exit Sub
Exception:
    Call MSZZ024_M00("cmb_Yard_from_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_from_KeyPress
'        機能             :数字のみ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_from_KeyPress(KeyAscii As Integer)
On Error GoTo Exception
    Call CmKeyPress_For_Num(Me.cmb_Yard_from, KeyAscii, 6)
    Exit Sub
Exception:
    Call MSZZ024_M00("cmb_Yard_from_KeyPress", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_Leader_AfterUpdate
'        機能             :ヤードリーダーの更新後
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_Leader_AfterUpdate()
On Error GoTo Exception
    Me.txt_Yard_Leader = ""
    If Nz(Me.cmb_Yard_Leader, "") = "" Then Exit Sub
    If IsNull(Me.cmb_Yard_Leader.Column(1)) Then
        MsgBox "該当するヤードリーダーが登録がされておりません。", vbInformation, P_ERROR
        Exit Sub
    End If
    Me.txt_Yard_Leader = Me.cmb_Yard_Leader.Column(1)
    If tab_Kensaku.VALUE = 0 Then '2018/08/11 ins
        Me.txt_Keyword_Tiku_K.SetFocus
    Else
        Me.txt_Keyword_Tiku_J.SetFocus
    End If
    Exit Sub
Exception:
    Call MSZZ024_M00("cmb_Yard_Leader_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_to_AfterUpdate
'        機能             :ヤードコードTOの更新後
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_To_AfterUpdate()
On Error GoTo Exception
    Me.txt_Yard_To = ""
    If Nz(Me.cmb_Yard_To, "") = "" Then Exit Sub
    If IsNull(Me.cmb_Yard_To.Column(1)) Then
        MsgBox "該当するヤードが登録がされておりません。", vbInformation, P_ERROR
        Exit Sub
    End If
    Me.txt_Yard_To = Me.cmb_Yard_To.Column(1)
    Me.txt_Todoufuken.SetFocus
    Exit Sub
Exception:
    Call MSZZ024_M00("cmb_Yard_To_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_To_KeyPress
'        機能             :数字のみ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_To_KeyPress(KeyAscii As Integer)
On Error GoTo Exception
    Call CmKeyPress_For_Num(Me.cmb_Yard_To, KeyAscii, 6)
    Exit Sub
Exception:
    Call MSZZ024_M00("cmb_Yard_To_KeyPress", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Excel_Click
'        機能             :Excelボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()
On Error GoTo Exception
    '件数確認
    If Me!FKS450_S01.Form.Recordset.RecordCount = 0 Then
        MsgBox "該当するデータがありません。", vbInformation, P_ERROR
        Exit Sub
    End If

    '実行確認
    If (MsgBox("Excel出力を実行しますか？", vbYesNo + vbQuestion, P_KAKUNIN) = vbNo) Then
        Exit Sub
    End If

    'クエリ作成
    If Not fncCreatQryExcel() Then
        Exit Sub
    End If
    
    'Excel出力
On Error Resume Next
    doCmd.OutputTo acQuery, P_WORK_EXCEL, acFormatXLS, , True

    'クエリ削除
On Error GoTo 0: On Error GoTo Exception
    Call subDeleteQry(P_WORK_EXCEL)
    
    Exit Sub
    
Exception:
    Call MSZZ024_M00("cmd_EXCEL_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncCreatQryExcel
'        機能             :Excel出力用クエリ作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreatQryExcel() As Boolean

    Dim qdf         As QueryDef
    Dim isException As Boolean
    
    isException = False
    
    fncCreatQryExcel = False
    
On Error GoTo Exception
    
    'ローカルDB
    If dbAccess Is Nothing Then
        Set dbAccess = CurrentDb
    End If
    
    'クエリ削除
    Call subDeleteQry(P_WORK_EXCEL)
    
    'クエリ作成
    Set qdf = dbAccess.CreateQueryDef("FKS450_Q01", fncGetExcelSql())
    Set qdf = Nothing
    
    '正常終了
    fncCreatQryExcel = True
    GoTo Finally
    
Exception:
    isException = True
    
Finally:
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    If isException Then Call Err.Raise(Err.Number, "fncCreatQryExcel" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGetExcelSql
'        機能             :Excel出力用クエリSQL
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetExcelSql() As String
    Dim strSQL  As String
    strSQL = strSQL & " select " & vbCrLf
    strSQL = strSQL & " YARD_CODE                  as ヤードコード" & vbCrLf
    strSQL = strSQL & ",YARD_NAME                  as ヤード名" & vbCrLf
    strSQL = strSQL & ",YARD_ADDR                  as 所在地" & vbCrLf
    'トランク・オフィスのチェック 2018/08/11 ins
    Select Case P_BUMON_CODE
    Case "8", "1" ', "H"
        'トランク・オフィス
        strSQL = strSQL & ",TANTM_TANTN                as 稼働率" & vbCrLf
    Case Else
        'その他
        strSQL = strSQL & ",TANTM_TANTN                as ヤードリーダー" & vbCrLf
    End Select
    strSQL = strSQL & ",KAKAKU_GENZAI_TNAME        as 現在適用・価格テンプレート" & vbCrLf
    strSQL = strSQL & ",KAKAKU_GENZAI_STDTE        as 現在適用・価格テンプレート設定日" & vbCrLf
    strSQL = strSQL & ",KAKAKU_ZENKAI_TNAME        as 前回適用・価格テンプレート" & vbCrLf
    strSQL = strSQL & ",KAKAKU_ZENKAI_STDTE        as 前回適用・価格テンプレート設定日" & vbCrLf
    strSQL = strSQL & ",KAKAKU_ZENZENKAI_TNAME     as 前々回適用・価格テンプレート" & vbCrLf
    strSQL = strSQL & ",KAKAKU_ZENZENKAI_STDTE     as 前々回適用・価格テンプレート設定日" & vbCrLf
    strSQL = strSQL & ",WARIBIKI_GENZAI_TEXT       as 現在適用・割引" & vbCrLf
    strSQL = strSQL & ",WARIBIKI_GENZAI_FROM       as 現在適用・割引設定日（FROM）" & vbCrLf
    strSQL = strSQL & ",WARIBIKI_GENZAI_TO         as 現在適用・割引設定日（TO）" & vbCrLf
    strSQL = strSQL & ",WARIBIKI_ZENKAI_TEXT       as 前回適用・割引" & vbCrLf
    strSQL = strSQL & ",WARIBIKI_ZENKAI_FROM       as 前回適用・割引設定日（FROM）" & vbCrLf
    strSQL = strSQL & ",WARIBIKI_ZENKAI_TO         as 前回適用・割引設定日（TO）" & vbCrLf
    strSQL = strSQL & ",WARIBIKI_ZENZENKAI_TEXT    as 前々回適用・割引" & vbCrLf
    strSQL = strSQL & ",WARIBIKI_ZENZENKAI_FROM    as 前々回適用・割引設定日（FROM）" & vbCrLf
    strSQL = strSQL & ",WARIBIKI_ZENZENKAI_TO      as 前々回適用・割引設定日（TO）" & vbCrLf
    strSQL = strSQL & " from " & P_WORK_TABLE & vbCrLf
    strSQL = strSQL & " order by YARD_CODE" & vbCrLf
    fncGetExcelSql = strSQL
End Function

'==============================================================================*
'
'        MODULE_NAME      :subDeleteQry
'        機能             :クエリ削除
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDeleteQry(strQry As String)
On Error GoTo Exception
    If dbAccess Is Nothing Then
        Set dbAccess = CurrentDb
    End If
    On Error Resume Next
    dbAccess.QueryDefs.Delete strQry
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subDeleteQry" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Waribiki_1_Click
'        機能             :割引マスタ画面予備゛だし
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Waribiki_Click()
On Error GoTo Exception
    Call subWaribikiMstCall
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Waribiki_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subWaribikiMstCall
'        機能             :割引マスタ表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subWaribikiMstCall()
    Dim strParameter        As String
    Dim strTabNo            As String
    Dim strActiveControl    As String
On Error GoTo Exception
    'タブ番号
    strTabNo = Right(ActiveControl.NAME, 1)
    '(0)モード - (1)ヤードコード - (2)ヤード名 - (3)割引数
    strParameter = "割引選択" & "," & "000000" & "," & "1"
    '割引マスタオープン
    doCmd.OpenForm "FKS700", , , , , acDialog, strParameter
    'OK/CANCEL 判定
    If Form_FKS450.p450_RET = True Then
        strActiveControl = "txt_Kikan_From_" & strTabNo
        '→OK
        '画面更新
        Call subWaribikiTabUpdate(strTabNo)
        'カーソルを実施日付FROMへ
        Me.Controls(strActiveControl).SetFocus
    End If
Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("subWaribikiMstCall", True)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subWaribikiTabUpdate
'        機能             :タブエリア表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subWaribikiTabUpdate(strTabNo As String)
    Dim vDatefrom   As Date
    Dim vDateto     As Date
On Error GoTo Exception
    '日付計算
    vDatefrom = DateAdd("m", 1, CDate(Format$(Now, "yyyy/mm") & "/01"))
    vDateto = DateAdd("d", -1, DateAdd("m", 1, vDatefrom))
    '内容更新
    Me.Controls("txt_Waribiki_no1_" & strTabNo) = p450_WaribikiNo                                                   '割引番号1
    Me.Controls("txt_Waribiki_no2_" & strTabNo) = fncGetWaribikiNo2(p450_WaribikiNo)                                '割引番号2
    Me.Controls("txt_Kikan_From_" & strTabNo) = Format(vDatefrom, "YYYY/MM/DD")                                     '実施日付FROM
    Me.Controls("txt_Kikan_To_" & strTabNo) = Format(vDateto, "YYYY/MM/DD")                                         '実施日付TO
    Me.Controls("txt_Mongon_" & strTabNo) = p450_Mongon                                                             '提示文言
    Me.Controls("txt_Sabisu_Tekiyou_" & strTabNo) = fncGet割引適用利用月数(p450_WaribikiNo)                         '適用満了月数 2018/11/15
    'タイトル
    If Me.Controls("txt_Waribiki_no2_" & strTabNo) = "" And p450_Youto = "0" Then
        Me.Controls("lbl_Waribiki_Taitle_" & strTabNo).Caption = "この割引には同じ条件の" & Chr(34) & "ｺﾝﾃﾅ(ﾊﾞｲｸ)" & Chr(34) & "はありません御注意下さい"
        Me.Controls("lbl_Waribiki_Taitle_" & strTabNo).ForeColor = 255
    Else
        Me.Controls("lbl_Waribiki_Taitle_" & strTabNo).Caption = "(割引番号: " & p450_WaribikiNo & ")"
        Me.Controls("lbl_Waribiki_Taitle_" & strTabNo).ForeColor = 16777215
    End If
    '確定ボタン
    Call subTabLock(Me.Controls("cmd_Kakutei_" & strTabNo), True)
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subWaribikiTabUpdate" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncGetWaribikiNo2
'        機能             :割引番号2取得
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetWaribikiNo2(strWaribikiNo As String)
    Dim objCon          As Object
    Dim objWaribiki     As Object
    Dim isException As Boolean
    isException = False

On Error GoTo Exception

    fncGetWaribikiNo2 = ""

    'ADOコネクト
    Set objCon = ADODB_Connection(P_BUMON_CODE)

    'レコードセット作成
    Set objWaribiki = ADODB_Recordset(fncGetWaribikiNo2Sql(strWaribikiNo), objCon)
    
    'ワークテーブル書き出し
    With objWaribiki
        Do While Not .EOF
           fncGetWaribikiNo2 = Nz(.Fields("DCNTM_NO"), "")
           Exit Do
        Loop
    End With

    GoTo Finally
    
Exception:
    isException = True
    
Finally:
    If Not objWaribiki Is Nothing Then objWaribiki.Close: Set objWaribiki = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    If isException Then Call Err.Raise(Err.Number, "fncGetWaribikiNo2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGetWaribikiNo2Sql
'        機能             :割引番号2取得SQL
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetWaribikiNo2Sql(strWaribikiNo As String)
    Dim strSQL As String
    strSQL = strSQL & " select B.DCNTM_NO"
'2018/01/26    strSql = strSql & " FROM   TEST_KASE_DB.dbo.DCNT_MAST B,"
    strSQL = strSQL & " FROM " & strSqlserver & ".dbo.DCNT_MAST B," '2018/01/26 chg
    strSQL = strSQL & "     ("
    strSQL = strSQL & "         select A.*"
'2018/01/26    strSql = strSql & "         from  TEST_KASE_DB.dbo.DCNT_MAST A"
    strSQL = strSQL & "         from " & strSqlserver & ".dbo.DCNT_MAST A"  '2018/01/26 chg
    strSQL = strSQL & "         where A.DCNTM_NO        = '" & strWaribikiNo & "'"
    strSQL = strSQL & "         AND   A.DCNTM_USAGE     = 0"
    strSQL = strSQL & "         AND   DCNTM_BUMOC       = '" & P_BUMON_CODE & "'"
    strSQL = strSQL & "     ) A"
    strSQL = strSQL & " where"
    strSQL = strSQL & "     b.DCNTM_BUMOC = a.DCNTM_BUMOC"
    strSQL = strSQL & " AND B.DCNTM_KIND                = A.DCNTM_KIND"
    strSQL = strSQL & " AND B.DCNTM_TYPE                = A.DCNTM_TYPE"
    strSQL = strSQL & " AND B.DCNTM_VALUE               = A.DCNTM_VALUE"
    strSQL = strSQL & " AND B.DCNTM_PERIOD              = A.DCNTM_PERIOD"
    strSQL = strSQL & " AND ISNULL(B.DCNTM_SIZE_FROM,0) = ISNULL(A.DCNTM_SIZE_FROM,0)"
    strSQL = strSQL & " AND ISNULL(B.DCNTM_SIZE_TO,0)   = ISNULL(A.DCNTM_SIZE_TO,0)"
    strSQL = strSQL & " AND ISNULL(B.DCNTM_FLOOR,0)     = ISNULL(A.DCNTM_FLOOR,0)"
    strSQL = strSQL & " AND B.DCNTM_GENKBN              = A.DCNTM_GENKBN"
    strSQL = strSQL & " AND B.DCNTM_USE_PERIOD          = A.DCNTM_USE_PERIOD"
    strSQL = strSQL & " AND B.DCNTM_USAGE               = 33"
    fncGetWaribikiNo2Sql = strSQL
End Function

'==============================================================================*
'
'        MODULE_NAME      :cmd_Waribiki_1_Click
'        機能             :割引ボタン1
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Waribiki_1_Click()
On Error GoTo Exception
    Call subWaribikiMstCall
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Waribiki_1_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Waribiki_2_Click
'        機能             :割引ボタン2
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Waribiki_2_Click()
On Error GoTo Exception
    Call subWaribikiMstCall
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Waribiki_2_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Waribiki_3_Click
'        機能             :割引ボタン3
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Waribiki_3_Click()
On Error GoTo Exception
    Call subWaribikiMstCall
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Waribiki_3_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Waribiki_4_Click
'        機能             :割引ボタン4
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Waribiki_4_Click()
On Error GoTo Exception
    Call subWaribikiMstCall
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Waribiki_4_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Waribiki_5_Click
'        機能             :割引ボタン5
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Waribiki_5_Click()
On Error GoTo Exception
    Call subWaribikiMstCall
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Waribiki_5_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Waribiki_Clear_Click
'        機能             :割引クリア ボタン
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Waribiki_Clear_Click()
On Error GoTo Exception
    Call subWaribikiClear
    tab_Waribiki_1.SetFocus
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Waribiki_Clear_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Waribiki_Clear_Click
'        機能             :割引エリアクリア
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subWaribikiClear()
    Dim i As Integer
On Error GoTo Exception
    For i = 1 To 5
        '入力ロック
        Me.Controls("tab_Waribiki_" & CStr(i)).Visible = True
        Me.Controls("cmd_Kakutei_" & CStr(i)).Caption = "編集"
        Me.Controls("cmd_Kakutei_" & CStr(i)).Enabled = False
        Me.Controls("cmd_Waribiki_" & CStr(i)).Enabled = False
        Me.Controls("txt_Kikan_From_" & CStr(i)).Enabled = False
        Me.Controls("txt_Kikan_To_" & CStr(i)).Enabled = False
        Me.Controls("txt_Mongon_" & CStr(i)).Enabled = False
        Me.Controls("txt_Biko_" & CStr(i)).Enabled = False
        Me.Controls("txt_Sabisu_1_" & CStr(i)).Enabled = False
        Me.Controls("txt_Sabisu_2_" & CStr(i)).Enabled = False
        Me.Controls("txt_Sabisu_3_" & CStr(i)).Enabled = False
        Me.Controls("txt_Sabisu_Kikan_" & CStr(i)).Enabled = False
        Me.Controls("txt_Sabisu_Tekiyou_" & CStr(i)).Enabled = False    '2018/11/15 add
        '入力内容
        Me.Controls("txt_Kikan_From_" & CStr(i)) = ""
        Me.Controls("txt_Kikan_To_" & CStr(i)) = ""
        Me.Controls("txt_Mongon_" & CStr(i)) = ""
        Me.Controls("txt_Biko_" & CStr(i)) = ""
        Me.Controls("txt_Sabisu_1_" & CStr(i)) = ""
        Me.Controls("txt_Sabisu_2_" & CStr(i)) = ""
        Me.Controls("txt_Sabisu_3_" & CStr(i)) = ""
        Me.Controls("txt_Sabisu_Kikan_" & CStr(i)) = ""
        Me.Controls("txt_Sabisu_Tekiyou_" & CStr(i)) = ""               '2018/11/15 add
        Me.Controls("lbl_Waribiki_Taitle_" & CStr(i)).Caption = ""
        Me.Controls("lbl_Waribiki_Taitle_" & CStr(i)).ForeColor = 16777215
        '隠しエリア
        Me.Controls("txt_Waribiki_no1_" & CStr(i)) = ""
        Me.Controls("txt_Waribiki_no2_" & CStr(i)) = ""
        Me.Controls("txt_kakutei_" & CStr(i)) = ""
        '割引ボタンのみ有効
        Me.Controls("cmd_Waribiki_" & CStr(i)).Enabled = True
    Next
    '割引一括コンボ更新
    Call subSetCmbIkatuWaribiki
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subWaribikiClear" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncInputServiceCheck
'        機能             :割引エリア　入力チェック
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputServiceCheck(objBotton As Object)
    Dim strTabNo            As String
    Dim strNewDateFrom      As String
    Dim strNewDateTo        As String
    
On Error GoTo Exception
    
    fncInputServiceCheck = False
        
    'タブ番号
    strTabNo = Right(objBotton.NAME, 1)

    '入力ステータス初期化
    Me.Controls("txt_kakutei_" & strTabNo) = ""
    
    '確定の場合
    If objBotton.Caption = "確定" Then
            
        '実施期間FROM
        strNewDateFrom = Format(Me.Controls("txt_Kikan_From_" & strTabNo).VALUE, "yyyymmdd")
        
        '実施期間TO
        strNewDateTo = Format(Me.Controls("txt_Kikan_To_" & strTabNo).VALUE, "yyyymmdd")
        
        '割引番号
        If Nz(Me.Controls("txt_Waribiki_no1_" & strTabNo).VALUE, "") = "" Then
            Call MsgBox("割引が選択されていません。", vbExclamation + vbOKOnly, P_ERROR)
            Exit Function
        End If
        
        '実施期間from必須入力チェック
        If Nz(Me.Controls("txt_Kikan_From_" & strTabNo).VALUE, "") = "" Then
            Call MsgBox("実施期間(FROM)を入力して下さい。", vbExclamation + vbOKOnly, P_ERROR)
            Me.Controls("txt_Kikan_From_" & strTabNo).SetFocus
            Exit Function
        End If
    
        '実施期間from日付形式チェック
        If Not IsDate(Nz(Me.Controls("txt_Kikan_From_" & strTabNo).VALUE)) Or CmLenb(Nz(Me.Controls("txt_Kikan_From_" & strTabNo).VALUE)) < 8 Then
            Call MsgBox("実施期間(FROM)はYYYY/MM/DD形式です。", vbExclamation + vbOKOnly, P_ERROR)
            Me.Controls("txt_Kikan_From_" & strTabNo).SetFocus
            Exit Function
        End If
        
        '実施期間from日付のチェック(システム日付以降)
        If Format(CDate(Nz(Me.Controls("txt_Kikan_From_" & strTabNo).VALUE)), "yyyy/mm/dd") < Format(Now, "yyyy/mm/dd") Then
            Call MsgBox("実施期間(FROM)は、本日以降を入力して下さい。", vbExclamation + vbOKOnly, P_ERROR)
            Me.Controls("txt_Kikan_From_" & strTabNo).SetFocus
            Exit Function
        End If
    
        '実施期間to必須入力チェック
        If Nz(Me.Controls("txt_Kikan_To_" & strTabNo).VALUE, "") = "" Then
            Call MsgBox("実施期間(TO)を入力して下さい。", vbExclamation + vbOKOnly, P_ERROR)
            Me.Controls("txt_Kikan_To_" & strTabNo).SetFocus
            Exit Function
        End If
        
        '実施期間to日付形式チェック
        If Not IsDate(Nz(Me.Controls("txt_Kikan_To_" & strTabNo).VALUE)) Or CmLenb(Nz(Me.Controls("txt_Kikan_To_" & strTabNo).VALUE)) < 8 Then
            Call MsgBox("実施期間(TO)はYYYY/MM/DD形式です。", vbExclamation + vbOKOnly, P_ERROR)
            Me.Controls("txt_Kikan_To_" & strTabNo).SetFocus
            Exit Function
        End If
                
        '実施期間To日付のチェック(翌日以降)
        If Format(CDate(Nz(Me.Controls("txt_Kikan_To_" & strTabNo).VALUE)), "yyyy/mm/dd") < Format(DateAdd("d", 1, Now), "yyyy/mm/dd") Then
            Call MsgBox("実施期間(TO)は、翌日以降を入力して下さい。", vbExclamation + vbOKOnly, P_ERROR)
            Me.Controls("txt_Kikan_To_" & strTabNo).SetFocus
            Exit Function
        End If

        '実施期間from～to大小チェック
        If strNewDateFrom > strNewDateTo Then
            Call MsgBox("実施期間の指定が矛盾しています。" & vbCrLf & "From <= To となるように入力して下さい。", vbExclamation + vbOKOnly, P_ERROR)
            Me.Controls("txt_Kikan_To_" & strTabNo).SetFocus
            Exit Function
        End If
        
        '提示文言必須入力チェック
        If Nz(Me.Controls("txt_Mongon_" & strTabNo).VALUE, "") = "" Then
            Call MsgBox("提示文言を入力して下さい。", vbExclamation + vbOKOnly, P_ERROR)
            Me.Controls("txt_Mongon_" & strTabNo).SetFocus
            Exit Function
        End If

        'サービス１必須入力チェック
        If Nz(Me.Controls("txt_Sabisu_1_" & strTabNo).VALUE, "") = "" Then
            Call MsgBox("サービス１を入力して下さい。", vbExclamation + vbOKOnly, P_ERROR)
            Me.Controls("txt_Sabisu_1_" & strTabNo).SetFocus
            Exit Function
        End If

        '適用満了月数の数値チェック add 2018/11/15
        If Not IsNumeric(Nz(Me.Controls("txt_Sabisu_Tekiyou_" & strTabNo).VALUE)) Then
            Call MsgBox("適用満了月数は数値を入力して下さい。", vbExclamation + vbOKOnly, P_ERROR)
            Me.Controls("txt_Sabisu_Tekiyou_" & strTabNo).SetFocus
            Exit Function
        End If
        
        '入力ステータス
        Me.Controls("txt_kakutei_" & strTabNo) = "OK"

    End If
    fncInputServiceCheck = True
    Exit Function
Exception:
    Call Err.Raise(Err.Number, "fncInputServiceCheck" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :subTabLock
'        機能             :割引エリア　ロック制御
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subTabLock(objBotton As Object, blnOnOff As Boolean)
    Dim i As Integer
On Error GoTo Exception
    i = Val(Right(objBotton.NAME, 1))
    Me.Controls("cmd_Waribiki_" & CStr(i)).Enabled = True
    Me.Controls("cmd_Kakutei_" & CStr(i)).Enabled = True
    Me.Controls("txt_Kikan_From_" & CStr(i)).Enabled = blnOnOff
    Me.Controls("txt_Kikan_To_" & CStr(i)).Enabled = blnOnOff
    Me.Controls("txt_Mongon_" & CStr(i)).Enabled = blnOnOff
    Me.Controls("txt_Biko_" & CStr(i)).Enabled = blnOnOff
    Me.Controls("txt_Sabisu_1_" & CStr(i)).Enabled = blnOnOff
    Me.Controls("txt_Sabisu_2_" & CStr(i)).Enabled = blnOnOff
    Me.Controls("txt_Sabisu_3_" & CStr(i)).Enabled = blnOnOff
    Me.Controls("txt_Sabisu_Kikan_" & CStr(i)).Enabled = blnOnOff
    Me.Controls("txt_Sabisu_Tekiyou_" & CStr(i)).Enabled = blnOnOff   '2018/11/15 add
    If blnOnOff = False Then
        objBotton.Caption = "編集"
    Else
        objBotton.Caption = "確定"
    End If
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subTabLock" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subKakuteiBotton
'        機能             :割引エリア　表示制御
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKakuteiBotton(objBotton As Object)
On Error GoTo Exception
    '画面チェック
    If fncInputServiceCheck(ActiveControl) = False Then
        Exit Sub
    End If
    '画面ロック
    If objBotton.Caption = "編集" Then
        '編集→確定
        Call subTabLock(ActiveControl, True)
    Else
        '確定→編集
        Call subTabLock(ActiveControl, False)
    End If
    '割引一括コンボ更新
    Call subSetCmbIkatuWaribiki
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subKakuteiBotton" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetCmbIkatuWaribiki
'        機能             :割引:一括設定コンボボックスの内容をセット
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetCmbIkatuWaribiki()
    Dim i As Integer
On Error GoTo Exception
    With cmb_Waribiki_Ikatu_Setei
        'コンボボックスセット
        .VALUE = ""
        .RowSource = ""
        Erase P_WARIBIKI_CODE
        '表示データセット
        .AddItem ";"
        .AddItem "現在の割引を止める;0"
        '割引ｺｰﾄﾞ格納
        P_WARIBIKI_CODE(0) = "0"
        For i = 1 To 5
            '入力ステータスのチェック
            If Me.Controls("txt_kakutei_" & CStr(i)) = "OK" Then
                '正常入力されていればコンボに追加
                .AddItem StrConv(Me.Controls("tab_Waribiki_" & CStr(i)).Caption, vbWide) & "にする" & ";" & CStr(i)
                '割引ｺｰﾄﾞ格納
                P_WARIBIKI_CODE(i) = CStr(i)
            End If
        Next
        '明細部
        If Me!FKS450_S01.Form.Recordset.RecordCount > 0 Then
            With Me.FKS450_S01.Form!cmb_Waribiki_New_Setei
                .RowSource = ""
                For i = 0 To cmb_Waribiki_Ikatu_Setei.ListCount - 1
                    .AddItem Me.cmb_Waribiki_Ikatu_Setei.Column(0, i) & ";" & Me.cmb_Waribiki_Ikatu_Setei.Column(1, i)
                Next
            End With
            FKS450_S01.Requery
        End If
    End With
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subSetCmbIkatuWaribiki" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Kakutei_1_Click
'        機能             :割引エリア1　確定/編集
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Kakutei_1_Click()
On Error GoTo Exception
    Call subKakuteiBotton(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Kakutei_1_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Kakutei_2_Click
'        機能             :割引エリア2　確定/編集
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Kakutei_2_Click()
On Error GoTo Exception
    Call subKakuteiBotton(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Kakutei_2_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Kakutei_3_Click
'        機能             :割引エリア3　確定/編集
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Kakutei_3_Click()
On Error GoTo Exception
    Call subKakuteiBotton(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Kakutei_3_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Kakutei_4_Click
'        機能             :割引エリア4　確定/編集
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Kakutei_4_Click()
On Error GoTo Exception
    Call subKakuteiBotton(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Kakutei_4_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Kakutei_5_Click
'        機能             :割引エリア5　確定/編集
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Kakutei_5_Click()
On Error GoTo Exception
    Call subKakuteiBotton(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_Kakutei_5_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_検索_Click
'        機能             :検索ボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_検索_Click()
On Error GoTo Exception

    '一括コンボボックス更新(価格)
    If Not fncSetCmbIkatuKakaku() Then
        Exit Sub
    End If
    
    '一括コンボボックス更新(事務手数料) 2018/08/21 ins
    If Not fncSetCmbIkatuJimute() Then
        Exit Sub
    End If
    
    'ワークテーブルとデータ作成
    If Not fncMakeWork() Then
        Exit Sub
    End If
    
    '新設定リストボックス
    Call subSetCmb
    
    '明細域のラベル
    Call subMeisaiLbl
    
    '再描画
    FKS450_S01.Requery

    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_検索_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Click
'        機能             :終了ボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
On Error Resume Next
    doCmd.Close acForm, Me.FormName
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_保存_Click
'        機能             :保存
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_Click()
    Dim blnKoushinFlg As Boolean
On Error GoTo Exception
    '件数確認
    If Me!FKS450_S01.Form.Recordset.RecordCount = 0 Then
        MsgBox "該当するデータがありません。", vbInformation, P_ERROR
        Exit Sub
    End If

    '保存確認
    If MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbNo Then
        Exit Sub
    End If

    '価格・割引テーブル更新
    If fncDbWrite(blnKoushinFlg) = False Then
        Exit Sub
    End If
    
    '更新件数チェック 2018/08/11
    If blnKoushinFlg = False Then
        MsgBox "保存対象はありませんでした。" & vbCrLf & "同じテンプレートまたは割引で連続更新は出来ません。", vbExclamation, P_KAKUNIN
        Exit Sub
    Else
        MsgBox "保存しました。", vbInformation, P_KAKUNIN
    End If
    
    '再検索
    Call cmd_検索_Click
    
    Exit Sub
Exception:
    Call MSZZ024_M00("cmd_保存_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncChkWaribikiCode
'        機能             :割引ｺｰﾄﾞチェック
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkWaribikiCode(strWARIBIKI_KKTHM_TCODE As String) As Boolean
    Dim i As Integer
On Error GoTo Exception
    fncChkWaribikiCode = False
    For i = 0 To 5
        If strWARIBIKI_KKTHM_TCODE = P_WARIBIKI_CODE(i) Then
            fncChkWaribikiCode = True
            Exit Function
        End If
    Next
    Exit Function
Exception:
    Call Err.Raise(Err.Number, "fncChkWaribikiCode" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncDbWrite
'        機能             :保存ボタンのDB更新処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDbWrite(ByRef blnKoushinFlg As Boolean) As Boolean
    Dim objCon                  As Object
    Dim objRsWK450              As Recordset
    Dim strSQL                  As String
    Dim datNowDate              As Date
    Dim isException As Boolean
    
    isException = False
    fncDbWrite = False
    blnKoushinFlg = False
    
On Error GoTo Exception1
    
    Set objCon = ADODB_Connection(P_BUMON_CODE)
    'ローカルDB
    If dbAccess Is Nothing Then
        Set dbAccess = CurrentDb
    End If
    
    datNowDate = Now
    strSQL = "select * from " & P_WORK_TABLE
    Set objRsWK450 = dbAccess.OpenRecordset(strSQL, dbReadOnly)
    
    'トランザクション開始
    objCon.BeginTrans
    
On Error GoTo Exception2
    
    With objRsWK450
        Do Until objRsWK450.EOF
            '価格の変更判断
            If Nz(.Fields("KAKAKU_KKTHM_TCODE"), "") <> "" Then 'スペース以外は更新対象
                '変更判断
                If Nz(.Fields("KAKAKU_KKTHM_TCODE"), "") <> Nz(.Fields("KAKAKU_GENZAI_TCODE"), "") Then
                    '履歴作成
                    If fncRirekiInsert(objCon, .Fields("YARD_CODE"), datNowDate, .Fields("KAKAKU_KKTHM_TCODE")) = False Then
                        GoTo Exception2
                    End If
                    '価格テーブル更新
                    If tab_Kensaku.VALUE = 0 Then '2018/08/11 ins
                        '価格更新
                        If fncKakakuUpdate(objCon, .Fields("YARD_CODE"), datNowDate, .Fields("KAKAKU_KKTHM_TCODE")) = False Then
                            GoTo Exception2
                        End If
                    Else
                        '事務手数料更新　2018/08/11 ins
                        If fncJimuteUpdate(objCon, .Fields("YARD_CODE"), datNowDate, .Fields("KAKAKU_KKTHM_TCODE")) = False Then
                            GoTo Exception2
                        End If
                    End If
                    '更新OK 2018/08/11
                    blnKoushinFlg = True
                End If
            End If
            '割引の変更判断
            If Nz(.Fields("WARIBIKI_KKTHM_TCODE"), "") <> "" Then 'スペース以外は更新対象
                '割引コンボチェック判定(コンボにないものが選ばれている場合はNG)
                If fncChkWaribikiCode((Nz(.Fields("WARIBIKI_KKTHM_TCODE"), ""))) = True Then
                    '変更判断
                    If Nz(.Fields("WARIBIKI_KKTHM_TCODE"), "") <> Nz(.Fields("WARIBIKI_GENZAI_TEXT"), "") Then '変更があった場合は更新対象
                        '割引テーブル無効化更新
                        If fncWaribikiInitUpdate(objCon, .Fields("YARD_CODE"), datNowDate, .Fields("WARIBIKI_KKTHM_TCODE")) = False Then
                            GoTo Exception2
                        End If
                        '割引テーブル追加
                        If fncWaribikiInsert(objCon, .Fields("YARD_CODE"), datNowDate, .Fields("WARIBIKI_KKTHM_TCODE")) = False Then
                            GoTo Exception2
                        End If
                    End If
                    '更新OK 2018/08/11
                    blnKoushinFlg = True
                End If
            End If
            '次へ
            .MoveNext
        Loop
    End With
    
    'コミット
    objCon.CommitTrans
    
    fncDbWrite = True
    
    GoTo Finally
    
Exception2:
    'ロールバック
    objCon.RollbackTrans
    
Exception1:
    isException = True

Finally:
    If Not objRsWK450 Is Nothing Then objRsWK450.Close: Set objRsWK450 = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    If isException Then Call Err.Raise(Err.Number, "fncDbWrite" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncCheckMeisaiCmb
'        機能             :明細ワークテーブルの更新チェック
'        IN               :
'        OUT              :
'
'==============================================================================*
Private Function fncCheckMeisaiCmb() As Boolean
    Dim objRsWK450  As Recordset
    Dim strSQL      As String
    Dim isException As Boolean
'On Error GoTo Exception1
    'ローカルDB
    If dbAccess Is Nothing Then
        Set dbAccess = CurrentDb
    End If
    '設定した件数取得
    strSQL = "select count(*) "
    strSQL = strSQL & " from " & P_WORK_TABLE
    strSQL = strSQL & " where KAKAKU_KKTHM_TCODE   is not null"
    strSQL = strSQL & " or    WARIBIKI_KKTHM_TCODE is not null"
    Set objRsWK450 = dbAccess.OpenRecordset(strSQL, dbReadOnly)
    With objRsWK450
        If .Fields(0) = 0 Then
            fncCheckMeisaiCmb = True
        Else
            fncCheckMeisaiCmb = False
        End If
    End With
    GoTo Finally
Exception1:
    isException = True
    
Finally:
    If Not objRsWK450 Is Nothing Then objRsWK450.Close: Set objRsWK450 = Nothing
    If isException Then Call Err.Raise(Err.Number, "fncCheckPnlKakakuJimute" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncRirekiInsert
'        機能             :履歴テーブル追加
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncRirekiInsert(objCon As Object, strYcode As String, datNowDate As Date, strTcode As String)
    Dim objRst              As Object
    fncRirekiInsert = False
    
On Error GoTo Exception

    Set objRst = ADODB_Recordset("KTPS_TRAN", objCon, adoAppendOnly)
    
On Error GoTo Exception1
    With objRst
        .AddNew
        On Error GoTo Exception2
        .Fields("KTPST_YCODE") = Format(Val(strYcode), "000000")
        .Fields("KTPST_STDTE") = datNowDate
        .Fields("KTPST_TCODE") = strTcode
        .Fields("KTPST_TTYPE") = IIf(tab_Kensaku.VALUE = 0, "0", "63") '2018/08/11 ins
        .Fields("KTPST_INSED") = Format(datNowDate, "YYYYMMDD")
        .Fields("KTPST_INSEJ") = Format(datNowDate, "HHNNSS")
        .Fields("KTPST_INSPB") = PROG_ID
        .Fields("KTPST_INSUB") = P_USER_ID
        .UPDATE
        On Error GoTo Exception1
        .Close
        On Error GoTo Exception
    End With
    fncRirekiInsert = True
Exit Function

Exception2:
    objRst.CancelUpdate
Exception1:
    objRst.Close
Exception:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncRirekiInsert" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncKakakuUpdate
'        機能             :価格テーブル更新(価格)
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncKakakuUpdate(objCon As Object, strYcode As String, datNowDate As Date, strTcode As String)
    Dim strSQL  As String
    Dim lngCnt  As Long
    
    fncKakakuUpdate = False
    
On Error GoTo Exception

'2018/08/11 del
'    strSQL = strSQL & " UPDATE PRIC_TABL"
'    strSQL = strSQL & " set"
'    strSQL = strSQL & "      PRIC_TABL.PRIC_PRICE        = KKTP_MAST.KKTPM_PRICE"
'    strSQL = strSQL & "     ,PRIC_TABL.PRIC_FZAPPI_CODE  = KKTP_MAST.KKTPM_FZAPPI_1_CODE"
'    strSQL = strSQL & "     ,PRIC_TABL.PRIC_FZAPPI       = KKTP_MAST.KKTPM_FZAPPI_1_COST"
'    strSQL = strSQL & "     ,PRIC_TABL.PRIC_EZAPPI_CODE  = KKTP_MAST.KKTPM_EZAPPI_1_CODE"
'    strSQL = strSQL & "     ,PRIC_TABL.PRIC_EZAPPI       = KKTP_MAST.KKTPM_EZAPPI_1_COST"
'    strSQL = strSQL & "     ,PRIC_TABL.PRIC_UPDAD       = '" & Format(datNowDate, "YYYYMMDD") & "'"
'    strSQL = strSQL & "     ,PRIC_TABL.PRIC_UPDAT       = '" & Format(datNowDate, "HHNNSS") & "'"
'    strSQL = strSQL & "     ,PRIC_TABL.PRIC_UPDPB       = '" & PROG_ID & "'"
'    strSQL = strSQL & "     ,PRIC_TABL.PRIC_UPDUB       = '" & P_USER_ID & "'"
'    strSQL = strSQL & " from PRIC_TABL"
'    strSQL = strSQL & " inner join KKTP_MAST"
'    strSQL = strSQL & "     on"
'    strSQL = strSQL & "         KKTP_MAST.KKTPM_TCODE = '" & strTcode & "'"
'    strSQL = strSQL & "     and KKTP_MAST.KKTPM_USAGE = PRIC_TABL.PRIC_USAGE"
'    strSQL = strSQL & "     and KKTP_MAST.KKTPM_SIZE  = PRIC_TABL.PRIC_SIZE"
'    strSQL = strSQL & "     and KKTP_MAST.KKTPM_STEP  = PRIC_TABL.PRIC_STEP"
'    strSQL = strSQL & " where"
'    strSQL = strSQL & "      PRIC_TABL.PRIC_YCODE = " & strYCODE

    strSQL = strSQL & " update PRIC_TABL"
    strSQL = strSQL & " set"
    '価格
    strSQL = strSQL & "  PRIC_TABL.PRIC_PRICE        = KKTP_MAST.KKTPM_PRICE"
    '初回
    strSQL = strSQL & " ,PRIC_TABL.PRIC_FZAPPI_CODE  = KKTP_MAST.KKTPM_FZAPPI_1_CODE"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_FZAPPI       = KKTP_MAST.KKTPM_FZAPPI_1_COST"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_FZAPPI_CODE2 = KKTP_MAST.KKTPM_FZAPPI_2_CODE "
    strSQL = strSQL & " ,PRIC_TABL.PRIC_FZAPPI2      = KKTP_MAST.KKTPM_FZAPPI_2_COST "
    strSQL = strSQL & " ,PRIC_TABL.PRIC_FZAPPI_CODE3 = KKTP_MAST.KKTPM_FZAPPI_3_CODE "
    strSQL = strSQL & " ,PRIC_TABL.PRIC_FZAPPI3      = KKTP_MAST.KKTPM_FZAPPI_3_COST "
    '毎月
    strSQL = strSQL & " ,PRIC_TABL.PRIC_EZAPPI_CODE  = KKTP_MAST.KKTPM_EZAPPI_1_CODE"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_EZAPPI       = KKTP_MAST.KKTPM_EZAPPI_1_COST"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_EZAPPI_CODE2 = KKTP_MAST.KKTPM_EZAPPI_2_CODE "
    strSQL = strSQL & " ,PRIC_TABL.PRIC_EZAPPI2      = KKTP_MAST.KKTPM_EZAPPI_2_COST "
    strSQL = strSQL & " ,PRIC_TABL.PRIC_EZAPPI_CODE3 = KKTP_MAST.KKTPM_EZAPPI_3_CODE "
    strSQL = strSQL & " ,PRIC_TABL.PRIC_EZAPPI3      = KKTP_MAST.KKTPM_EZAPPI_3_COST "
    '共通
    strSQL = strSQL & " ,PRIC_TABL.PRIC_UPDAD       = '" & Format(datNowDate, "YYYYMMDD") & "'"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_UPDAT       = '" & Format(datNowDate, "HHNNSS") & "'"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_UPDPB       = '" & PROG_ID & "'"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_UPDUB       = '" & P_USER_ID & "'"
    strSQL = strSQL & " from PRIC_TABL"
    strSQL = strSQL & " inner join KKTP_MAST"
    strSQL = strSQL & "     on"
    strSQL = strSQL & "         KKTP_MAST.KKTPM_TCODE = '" & strTcode & "'"
    strSQL = strSQL & "     and KKTP_MAST.KKTPM_USAGE = PRIC_TABL.PRIC_USAGE"
    strSQL = strSQL & "     and KKTP_MAST.KKTPM_SIZE  = PRIC_TABL.PRIC_SIZE"
    strSQL = strSQL & "     and KKTP_MAST.KKTPM_STEP  = PRIC_TABL.PRIC_STEP"
    strSQL = strSQL & " where"
    strSQL = strSQL & "      PRIC_TABL.PRIC_YCODE = " & strYcode
    
    '実行
    lngCnt = ADODB_Execute(strSQL, objCon)
    
    fncKakakuUpdate = True

    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncKakakuUpdate" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncJimuteUpdate
'        機能             :価格テーブル更新(事務手数料)
'        CREATE DATE      :2018/08/11
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncJimuteUpdate(objCon As Object, strYcode As String, datNowDate As Date, strTcode As String)
    Dim strSQL  As String
    Dim lngCnt  As Long
    
    fncJimuteUpdate = False
    
On Error GoTo Exception

    '更新
    strSQL = strSQL & " update PRIC_TABL"
    strSQL = strSQL & " set"
    strSQL = strSQL & " PRIC_TABL.PRIC_FZAPPI  = "
    strSQL = strSQL & "         CASE WHEN IsNull(PRIC_TABL.PRIC_FZAPPI_CODE, '0') <> '63' "
    strSQL = strSQL & "         THEN "
    strSQL = strSQL & "             PRIC_TABL.PRIC_FZAPPI   "
    strSQL = strSQL & "         ELSE "
    strSQL = strSQL & "             CASE WHEN IsNull(KKTP_MAST.KKTPM_FZAPPI_1_COST, 0) = 999999999 "
    strSQL = strSQL & "             THEN"
    strSQL = strSQL & "                 PRIC_TABL.PRIC_PRICE"
    strSQL = strSQL & "             ELSE"
    strSQL = strSQL & "                 KKTP_MAST.KKTPM_FZAPPI_1_COST"
    strSQL = strSQL & "             END   "
    strSQL = strSQL & "         END"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_FZAPPI2  = "
    strSQL = strSQL & "         CASE WHEN IsNull(PRIC_TABL.PRIC_FZAPPI_CODE2, '0') <> '63' "
    strSQL = strSQL & "         THEN "
    strSQL = strSQL & "             PRIC_TABL.PRIC_FZAPPI2"
    strSQL = strSQL & "         ELSE "
    strSQL = strSQL & "             CASE WHEN IsNull(KKTP_MAST.KKTPM_FZAPPI_1_COST, 0) = 999999999 "
    strSQL = strSQL & "             THEN"
    strSQL = strSQL & "                 PRIC_TABL.PRIC_PRICE"
    strSQL = strSQL & "             ELSE"
    strSQL = strSQL & "                 KKTP_MAST.KKTPM_FZAPPI_1_COST"
    strSQL = strSQL & "             END    "
    strSQL = strSQL & "         END"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_FZAPPI3  = "
    strSQL = strSQL & "         CASE WHEN IsNull(PRIC_TABL.PRIC_FZAPPI_CODE3, '0') <> '63' "
    strSQL = strSQL & "         THEN "
    strSQL = strSQL & "             PRIC_TABL.PRIC_FZAPPI3"
    strSQL = strSQL & "         ELSE "
    strSQL = strSQL & "             CASE WHEN IsNull(KKTP_MAST.KKTPM_FZAPPI_1_COST, 0) = 999999999 "
    strSQL = strSQL & "             THEN"
    strSQL = strSQL & "                 PRIC_TABL.PRIC_PRICE"
    strSQL = strSQL & "             ELSE"
    strSQL = strSQL & "                 KKTP_MAST.KKTPM_FZAPPI_1_COST"
    strSQL = strSQL & "             END   "
    strSQL = strSQL & "         END"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_UPDAD       = '" & Format(datNowDate, "YYYYMMDD") & "'"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_UPDAT       = '" & Format(datNowDate, "HHNNSS") & "'"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_UPDPB       = '" & PROG_ID & "'"
    strSQL = strSQL & " ,PRIC_TABL.PRIC_UPDUB       = '" & P_USER_ID & "'"
    strSQL = strSQL & " from PRIC_TABL"
    strSQL = strSQL & " inner join KKTP_MAST"
    strSQL = strSQL & "     on"
    strSQL = strSQL & "         KKTP_MAST.KKTPM_TCODE = '" & strTcode & "'"
    strSQL = strSQL & " and IsNull(KKTP_MAST.KKTPM_USAGE,PRIC_TABL.PRIC_USAGE) = PRIC_TABL.PRIC_USAGE"
    strSQL = strSQL & " and IsNull(KKTP_MAST.KKTPM_SIZE, PRIC_TABL.PRIC_SIZE)  = PRIC_TABL.PRIC_SIZE"
    strSQL = strSQL & " and IsNull(KKTP_MAST.KKTPM_STEP, PRIC_TABL.PRIC_STEP)  = PRIC_TABL.PRIC_STEP"
    strSQL = strSQL & " where"
    strSQL = strSQL & "      PRIC_TABL.PRIC_YCODE = " & strYcode

    '実行
    lngCnt = ADODB_Execute(strSQL, objCon)

    fncJimuteUpdate = True

    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncJimuteUpdate" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncWaribikiUpdate
'        機能             :割引テーブル更新
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncWaribikiInitUpdate(objCon As Object, strYcode As String, datNowDate As Date, strTabNo As String)
    Dim strfrom     As String
    Dim strTo       As String
    Dim strSQL      As String
    Dim lngCnt      As Long
    
On Error GoTo Exception
    
    fncWaribikiInitUpdate = False
    strfrom = Format(Me.Controls("txt_Kikan_From_" & strTabNo), "YYYYMMDD")     '実施期間From
    strTo = Format(Me.Controls("txt_Kikan_To_" & strTabNo), "YYYYMMDD")         '実施期間To
    P_USER_ID = LsGetUserName()
    
    '2019/02/26 M.HONDA INS
    '未来の割引があれば強制的に無効
    strSQL = strSQL & " update RLDN_TRAN set" & vbCrLf
    strSQL = strSQL & "  RLDNT_ENABLE    = '0'" & vbCrLf
    strSQL = strSQL & " ,RLDNT_GNKBN     = '1'" & vbCrLf
    strSQL = strSQL & " ,RLDNT_CTLFLG    = '1'" & vbCrLf
    strSQL = strSQL & " ,RLDNT_UPDAD     = '" & Format(datNowDate, "YYYYMMDD") & "'" & vbCrLf
    strSQL = strSQL & " ,RLDNT_UPDAJ     = '" & Format(datNowDate, "HHNNSS") & "'" & vbCrLf
    strSQL = strSQL & " ,RLDNT_UPDPB     = '" & PROG_ID & "'" & vbCrLf
    strSQL = strSQL & " ,RLDNT_UPDUB     = '" & P_USER_ID & "'" & vbCrLf
    strSQL = strSQL & " where" & vbCrLf
    strSQL = strSQL & "     RLDNT_BUMOC  = '" & P_BUMON_CODE & "'" & vbCrLf
    strSQL = strSQL & " and RLDNT_YCODE  = '" & Format(Val(strYcode), "000000") & "'" & vbCrLf
    strSQL = strSQL & " and RLDNT_ENABLE = '1'" & vbCrLf
    If strTabNo <> "0" Then
        strfrom = Format(Me.Controls("txt_Kikan_From_" & strTabNo), "YYYYMMDD")     '実施期間From
        strSQL = strSQL & " AND RLDNT_FROM >= " & strfrom
    End If
    '実行
    lngCnt = ADODB_Execute(strSQL, objCon)
    
    strSQL = ""
    '現在行っている割引がある場合に終了日を開始日の前日に設定
    strSQL = strSQL & " update RLDN_TRAN set" & vbCrLf
    strSQL = strSQL & "        RLDNT_TO = '" & Replace(DateAdd("d", -1, CDate(Format(strfrom, "####/##/##"))), "/", "") & "'"
    strSQL = strSQL & " ,RLDNT_UPDAD     = '" & Format(datNowDate, "YYYYMMDD") & "'" & vbCrLf
    strSQL = strSQL & " ,RLDNT_UPDAJ     = '" & Format(datNowDate, "HHNNSS") & "'" & vbCrLf
    strSQL = strSQL & " ,RLDNT_UPDPB     = '" & PROG_ID & "'" & vbCrLf
    strSQL = strSQL & " ,RLDNT_UPDUB     = '" & P_USER_ID & "'" & vbCrLf
    strSQL = strSQL & " where" & vbCrLf
    strSQL = strSQL & "     RLDNT_BUMOC  = '" & P_BUMON_CODE & "'" & vbCrLf
    strSQL = strSQL & " and RLDNT_YCODE  = '" & Format(Val(strYcode), "000000") & "'" & vbCrLf
    strSQL = strSQL & " and RLDNT_ENABLE = '1'" & vbCrLf
    If strTabNo <> "0" Then
        strSQL = strSQL & " and ("
        strSQL = strSQL & "   RLDNT_FROM between " & strfrom & " And " & strTo & vbCrLf     '--選択した割引きの実施期間に跨がる開始日
        strSQL = strSQL & "   or"
        strSQL = strSQL & "   RLDNT_TO   between " & strfrom & " And " & strTo & vbCrLf     '--選択した割引きの実施期間に跨がる終了日
        strSQL = strSQL & "   or"
        strSQL = strSQL & "   " & strfrom & " between RLDNT_FROM and RLDNT_TO"              '--選択した割引き
        strSQL = strSQL & "   or"
        strSQL = strSQL & "   " & strTo & " between RLDNT_FROM and RLDNT_TO"                '--選択した割引き
        strSQL = strSQL & "   or"
        strSQL = strSQL & "  (RLDNT_FROM = '00000000' and RLDNT_TO = '99999999')"           '--無期限割引きは削除
        strSQL = strSQL & " )"
    End If
    
    '実行
    lngCnt = ADODB_Execute(strSQL, objCon)
    '2019/02/26 M.HONDA INS

    
    '現在の割引を止める
'    strSql = strSql & " update RLDN_TRAN set" & vbCrLf
'    strSql = strSql & "  RLDNT_ENABLE    = '0'" & vbCrLf
'    strSql = strSql & " ,RLDNT_GNKBN     = '1'" & vbCrLf
'    strSql = strSql & " ,RLDNT_CTLFLG    = '1'" & vbCrLf
'    strSql = strSql & " ,RLDNT_UPDAD     = '" & Format(datNowDate, "YYYYMMDD") & "'" & vbCrLf
'    strSql = strSql & " ,RLDNT_UPDAJ     = '" & Format(datNowDate, "HHNNSS") & "'" & vbCrLf
'    strSql = strSql & " ,RLDNT_UPDPB     = '" & PROG_ID & "'" & vbCrLf
'    strSql = strSql & " ,RLDNT_UPDUB     = '" & P_USER_ID & "'" & vbCrLf
'    strSql = strSql & " where" & vbCrLf
'    strSql = strSql & "     RLDNT_BUMOC  = '" & P_BUMON_CODE & "'" & vbCrLf
'    strSql = strSql & " and RLDNT_YCODE  = '" & Format(Val(strYCODE), "000000") & "'" & vbCrLf
'    strSql = strSql & " and RLDNT_ENABLE = '1'" & vbCrLf
'    If strTabNo <> "0" Then
'        strFrom = Format(Me.Controls("txt_Kikan_From_" & strTabNo), "YYYYMMDD")     '実施期間From
'        strTo = Format(Me.Controls("txt_Kikan_To_" & strTabNo), "YYYYMMDD")         '実施期間To
'        strSql = strSql & " and ("
'        strSql = strSql & "   RLDNT_FROM between " & strFrom & " And " & strTo & vbCrLf     '--選択した割引きの実施期間に跨がる開始日
'        strSql = strSql & "   or"
'        strSql = strSql & "   RLDNT_TO   between " & strFrom & " And " & strTo & vbCrLf     '--選択した割引きの実施期間に跨がる終了日
'        strSql = strSql & "   or"
'        strSql = strSql & "   " & strFrom & " between RLDNT_FROM and RLDNT_TO"              '--選択した割引き
'        strSql = strSql & "   or"
'        strSql = strSql & "   " & strTo & " between RLDNT_FROM and RLDNT_TO"                '--選択した割引き
'        strSql = strSql & "   or"
'        strSql = strSql & "  (RLDNT_FROM = '00000000' and RLDNT_TO = '99999999')"           '--無期限割引きは削除
'        strSql = strSql & " )"
'    End If
'    '実行
'    lngCNT = ADODB_Execute(strSql, objCon)
    
    fncWaribikiInitUpdate = True

    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncWaribikiInitUpdate" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncWaribikiUpdate
'        機能             :割引テーブル更新
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncWaribikiInsert(objCon As Object, strYcode As String, datNowDate As Date, strTabNo As String)
    
    Dim objRst          As Object
    Dim strSQL          As String
    Dim blnUpdateFlg    As Boolean
   
    fncWaribikiInsert = False
    
    '1～5のみ追加、それ以外はスキップ
    If strTabNo = "0" Then
        fncWaribikiInsert = True
        Exit Function
    End If

On Error GoTo Exception

    '割引番号1
    blnUpdateFlg = False
    strSQL = ""
    strSQL = strSQL & " select * from RLDN_TRAN "
    strSQL = strSQL & " where RLDNT_YCODE = '" & Format(Val(strYcode), "000000") & "'"
    strSQL = strSQL & " and   RLDNT_NO = '" & Me.Controls("txt_Waribiki_no1_" & strTabNo) & "'"
    strSQL = strSQL & " and   RLDNT_FROM = '" & Format(Me.Controls("txt_Kikan_From_" & strTabNo), "YYYYMMDD") & "'"
    Set objRst = ADODB_Recordset(strSQL, objCon, adoReadWrite)
    
On Error GoTo Exception1

    With objRst
        While Not .EOF
            blnUpdateFlg = True
            On Error GoTo Exception2
            '上書き
            .Fields("RLDNT_ENABLE") = "1"                                                           '4 .有効可否
            .Fields("RLDNT_FROM") = Format(Me.Controls("txt_Kikan_From_" & strTabNo), "YYYYMMDD")   '5 .実施期間From
            .Fields("RLDNT_TO") = Format(Me.Controls("txt_Kikan_To_" & strTabNo), "YYYYMMDD")       '6 .実施期間To
            .Fields("RLDNT_ORDER") = "1"                                                            '7 .適用優先順
            .Fields("RLDNT_TEXT") = Me.Controls("txt_Mongon_" & strTabNo)                           '8 .提示文言
            .Fields("RLDNT_NOTE") = Me.Controls("txt_Biko_" & strTabNo)                             '9 .備考
            .Fields("RLDNT_GNKBN") = "1"                                                            '10.発生区分
            .Fields("RLDNT_SEV1N") = Me.Controls("txt_Sabisu_1_" & strTabNo)                        '11.サービス内容１
            .Fields("RLDNT_SEV2N") = Me.Controls("txt_Sabisu_2_" & strTabNo)                        '12.サービス内容２
            .Fields("RLDNT_SEV3N") = Me.Controls("txt_Sabisu_3_" & strTabNo)                        '13.サービス内容３
            .Fields("RLDNT_ENDEN") = Me.Controls("txt_Sabisu_Kikan_" & strTabNo)                    '14.サービス期間
            .Fields("RLDNT_SEV_EXMONTH") = Me.Controls("txt_Sabisu_Tekiyou_" & strTabNo)            '適用満了月数 add 2018/11/15
            .Fields("RLDNT_CTLFLG") = "1"                                                           '15.制御フラグ
            .Fields("RLDNT_UPDAD") = Format(datNowDate, "YYYYMMDD")                                 '16.作成日付
            .Fields("RLDNT_UPDAJ") = Format(datNowDate, "HHNNSS")                                   '17.作成時刻
            .Fields("RLDNT_UPDPB") = PROG_ID                                                        '18.作成プログラムID
            .Fields("RLDNT_UPDUB") = P_USER_ID                                                      '19.作成ユーザーID
            .UPDATE
            .MoveNext
        Wend
        If blnUpdateFlg = False Then
            '追加
            .AddNew
            .Fields("RLDNT_BUMOC") = P_BUMON_CODE                                                   '1 .部門コード
            .Fields("RLDNT_YCODE") = Format(Val(strYcode), "000000")                                '2 .ヤードコード
            .Fields("RLDNT_NO") = Me.Controls("txt_Waribiki_no1_" & strTabNo)                       '3 .割引番号
            .Fields("RLDNT_ENABLE") = "1"                                                           '4 .有効可否
            .Fields("RLDNT_FROM") = Format(Me.Controls("txt_Kikan_From_" & strTabNo), "YYYYMMDD")   '5 .実施期間From
            .Fields("RLDNT_TO") = Format(Me.Controls("txt_Kikan_To_" & strTabNo), "YYYYMMDD")       '6 .実施期間To
            .Fields("RLDNT_ORDER") = "1"                                                            '7 .適用優先順
            .Fields("RLDNT_TEXT") = Me.Controls("txt_Mongon_" & strTabNo)                           '8 .提示文言
            .Fields("RLDNT_NOTE") = Me.Controls("txt_Biko_" & strTabNo)                             '9 .備考
            .Fields("RLDNT_GNKBN") = "1"                                                            '10.発生区分
            .Fields("RLDNT_SEV1N") = Me.Controls("txt_Sabisu_1_" & strTabNo)                        '11.サービス内容１
            .Fields("RLDNT_SEV2N") = Me.Controls("txt_Sabisu_2_" & strTabNo)                        '12.サービス内容２
            .Fields("RLDNT_SEV3N") = Me.Controls("txt_Sabisu_3_" & strTabNo)                        '13.サービス内容３
            .Fields("RLDNT_ENDEN") = Me.Controls("txt_Sabisu_Kikan_" & strTabNo)                    '14.サービス期間
            .Fields("RLDNT_SEV_EXMONTH") = Me.Controls("txt_Sabisu_Tekiyou_" & strTabNo)            '適用満了月数 add 2018/11/15
            .Fields("RLDNT_CTLFLG") = "1"                                                           '15.制御フラグ
            .Fields("RLDNT_INSED") = Format(datNowDate, "YYYYMMDD")                                 '16.作成日付
            .Fields("RLDNT_INSEJ") = Format(datNowDate, "HHNNSS")                                   '17.作成時刻
            .Fields("RLDNT_INSPB") = PROG_ID                                                        '18.作成プログラムID
            .Fields("RLDNT_INSUB") = P_USER_ID                                                      '19.作成ユーザーID
            .UPDATE
        End If
        .Close: Set objRst = Nothing
    End With

    '割引番号2判断
    If Me.Controls("txt_Waribiki_no2_" & strTabNo) = "" Then
        '無い場合は更新しない
        GoTo Exit_rtn
    End If

On Error GoTo Exception

    '割引番号2
    blnUpdateFlg = False
    strSQL = ""
    strSQL = strSQL & " select * from RLDN_TRAN "
    strSQL = strSQL & " where RLDNT_YCODE = '" & Format(Val(strYcode), "000000") & "'"
    strSQL = strSQL & " and   RLDNT_NO = '" & Me.Controls("txt_Waribiki_no2_" & strTabNo) & "'"
    strSQL = strSQL & " and   RLDNT_FROM = '" & Format(Me.Controls("txt_Kikan_From_" & strTabNo), "YYYYMMDD") & "'"
    Set objRst = ADODB_Recordset(strSQL, objCon, adoReadWrite)
    
On Error GoTo Exception1

    With objRst
        While Not .EOF
            blnUpdateFlg = True
            On Error GoTo Exception2
            '上書き
            .Fields("RLDNT_ENABLE") = "1"                                                           '4 .有効可否
            .Fields("RLDNT_FROM") = Format(Me.Controls("txt_Kikan_From_" & strTabNo), "YYYYMMDD")   '5 .実施期間From
            .Fields("RLDNT_TO") = Format(Me.Controls("txt_Kikan_To_" & strTabNo), "YYYYMMDD")       '6 .実施期間To
            .Fields("RLDNT_ORDER") = "1"                                                            '7 .適用優先順
            .Fields("RLDNT_TEXT") = Me.Controls("txt_Mongon_" & strTabNo)                           '8 .提示文言
            .Fields("RLDNT_NOTE") = Me.Controls("txt_Biko_" & strTabNo)                             '9 .備考
            .Fields("RLDNT_GNKBN") = "1"                                                            '10.発生区分
            .Fields("RLDNT_SEV1N") = Me.Controls("txt_Sabisu_1_" & strTabNo)                        '11.サービス内容１
            .Fields("RLDNT_SEV2N") = Me.Controls("txt_Sabisu_2_" & strTabNo)                        '12.サービス内容２
            .Fields("RLDNT_SEV3N") = Me.Controls("txt_Sabisu_3_" & strTabNo)                        '13.サービス内容３
            .Fields("RLDNT_ENDEN") = Me.Controls("txt_Sabisu_Kikan_" & strTabNo)                    '14.サービス期間
            .Fields("RLDNT_SEV_EXMONTH") = Me.Controls("txt_Sabisu_Tekiyou_" & strTabNo)            '適用満了月数 add 2018/11/15
            .Fields("RLDNT_CTLFLG") = "1"                                                           '15.制御フラグ
            .Fields("RLDNT_UPDAD") = Format(datNowDate, "YYYYMMDD")                                 '16.作成日付
            .Fields("RLDNT_UPDAJ") = Format(datNowDate, "HHNNSS")                                   '17.作成時刻
            .Fields("RLDNT_UPDPB") = PROG_ID                                                        '18.作成プログラムID
            .Fields("RLDNT_UPDUB") = P_USER_ID                                                      '19.作成ユーザーID
            .UPDATE
            .MoveNext
        Wend
        If blnUpdateFlg = False Then
            '追加
            .AddNew
            On Error GoTo Exception2
            .Fields("RLDNT_BUMOC") = P_BUMON_CODE                                                   '1 .部門コード
            .Fields("RLDNT_YCODE") = Format(Val(strYcode), "000000")                                '2 .ヤードコード
            .Fields("RLDNT_NO") = Me.Controls("txt_Waribiki_no2_" & strTabNo)                       '3 .割引番号
            .Fields("RLDNT_ENABLE") = "1"                                                           '4 .有効可否
            .Fields("RLDNT_FROM") = Format(Me.Controls("txt_Kikan_From_" & strTabNo), "YYYYMMDD")   '5 .実施期間From
            .Fields("RLDNT_TO") = Format(Me.Controls("txt_Kikan_To_" & strTabNo), "YYYYMMDD")       '6 .実施期間To
            .Fields("RLDNT_ORDER") = "2"                                                            '7 .適用優先順
            .Fields("RLDNT_TEXT") = Me.Controls("txt_Mongon_" & strTabNo)                           '8 .提示文言
            .Fields("RLDNT_NOTE") = Me.Controls("txt_Biko_" & strTabNo)                             '9 .備考
            .Fields("RLDNT_GNKBN") = "1"                                                            '10.発生区分
            .Fields("RLDNT_SEV1N") = Me.Controls("txt_Sabisu_1_" & strTabNo)                        '11.サービス内容１
            .Fields("RLDNT_SEV2N") = Me.Controls("txt_Sabisu_2_" & strTabNo)                        '12.サービス内容２
            .Fields("RLDNT_SEV3N") = Me.Controls("txt_Sabisu_3_" & strTabNo)                        '13.サービス内容３
            .Fields("RLDNT_ENDEN") = Me.Controls("txt_Sabisu_Kikan_" & strTabNo)                    '14.サービス期間
            .Fields("RLDNT_SEV_EXMONTH") = Me.Controls("txt_Sabisu_Tekiyou_" & strTabNo)            '適用満了月数 add 2018/11/15
            .Fields("RLDNT_CTLFLG") = "0"                                                           '15.制御フラグ
            .Fields("RLDNT_INSED") = Format(datNowDate, "YYYYMMDD")                                 '16.作成日付
            .Fields("RLDNT_INSEJ") = Format(datNowDate, "HHNNSS")                                   '17.作成時刻
            .Fields("RLDNT_INSPB") = PROG_ID                                                        '18.作成プログラムID
            .Fields("RLDNT_INSUB") = P_USER_ID                                                      '19.作成ユーザーID
            .UPDATE
        End If
        .Close: Set objRst = Nothing
    End With
  
Exit_rtn:
    fncWaribikiInsert = True
Exit Function

Exception2:
    Exit Function
Exception1:
    objRst.Close: Set objRst = Nothing
Exception:
    Call Err.Raise(Err.Number, "fncWaribikiInsert" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        機能             :Form_Activateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()

    doCmd.Maximize

End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォームクローズ
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    
    Call MSHS010_M00(PROG_ID, "2")
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDown
'        機能             :Form_KeyDownイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    'F1=「終了」 F2=「検索」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_終了.SetFocus
        Case vbKeyF2
            KeyCode = 0
            Me.cmd_検索.SetFocus
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        機能             :Form Loadイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    
    Dim blnRet          As Boolean
    Dim strTemp         As String
    Dim strParam()      As String
    Dim i               As Integer

On Error GoTo Exception

    'ユーザー
    P_USER_ID = LsGetUserName()

    'SQL-Server接続
    P_BUMON_CODE = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    
    'P_BUMON_CODE = "O" 'test
    
    If P_BUMON_CODE = "" Then
        'テーブル[CONT_MAST]の設定不正
        doCmd.Close acForm, Me.FormName
        Exit Sub
    End If
    'KASE_DBのDB名の取得を行う
    strSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    If strSqlserver = "" Then
        'テーブル[SETU_TABL]の設定不正
        doCmd.Close acForm, Me.FormName
        Exit Sub
    End If
    
    '接続文字列取得
    strConnect = MSZZ007_M00(P_BUMON_CODE)
    If strConnect = "" Then
        'テーブル[SETU_TABL]の設定不正
        doCmd.Close acForm, Me.FormName
        Exit Sub
    End If

    'ワーククリア
    If Not fncClearWork(P_WORK_TABLE) Then
        doCmd.Close acForm, Me.FormName
        Exit Sub
    End If
    
    'ヤードリーダー/稼働率の切り替え
    Call subChgYrdKado
    
    '稼働率コンボセット
    Call subKadoritu
    
    '割引エリア
    Call subWaribikiClear

    Exit Sub
    
Exception:
    Call MSZZ024_M00("Form_Load", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subKadou
'        機能             :
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKadoritu()
    Dim objCon              As Object
    Dim objRst              As Object
    Dim strSQL              As String
    Dim isException         As Boolean
    cmb_Kadoritu.RowSource = ""
    cmb_Kadoritu.AddItem ";X;0"
    Set objCon = ADODB_Connection(P_BUMON_CODE)
    strSQL = "select * from NAME_MAST where NAME_ID = '328' order by NAME_CODE"
    Set objRst = ADODB_Recordset(strSQL, objCon)
    With objRst
        While Not .EOF
            cmb_Kadoritu.AddItem Nz(.Fields("NAME_NAME")) & ";" & Nz(.Fields("NAME_VALUE_FROM")) & ";" & Nz(.Fields("NAME_VALUE_TO"))
            .MoveNext
        Wend
    End With
    GoTo Finally
    
Exception:
    isException = True
    
Finally:
    If Not objRst Is Nothing Then objRst.Close: Set objRst = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    If isException Then Call Err.Raise(Err.Number, "fncSetCmbIkatuKakaku" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subChgYrdKado
'        機能             :ヤードリーダー/稼働率のコンボボックスの切り替え
'        CREATE DATE      : 2018/08/11
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subChgYrdKado()
    Select Case P_BUMON_CODE
    Case "8", "1" ', "H"
        '■トランク/オフィス
        '稼働率
        lbl_Kadoritu.Visible = True
        cmb_Kadoritu.Visible = True
        'ヤードリーダー
        lbl_Yard_Leader.Visible = False
        cmb_Yard_Leader.Visible = False
        txt_Yard_Leader.Visible = False
    Case Else
        '■その他
        '稼働率
        lbl_Kadoritu.Visible = False
        cmb_Kadoritu.Visible = False
        'ヤードリーダー
        lbl_Yard_Leader.Visible = True
        cmb_Yard_Leader.Visible = True
        txt_Yard_Leader.Visible = True
    End Select
    '明細ラベル
    Call subMeisaiLbl
End Sub

'==============================================================================*
'
'        MODULE_NAME      :SubFormReload
'        機能             :一覧再表示
'        IN               :True=該当データ無しメッセージを表示 False=表示しない
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function subFormReload(Optional blnNotFoundMsg As Boolean = True) As Boolean

On Error GoTo Exception

    subFormReload = False
    
    '一括コンボボックス更新
    If Not fncSetCmbIkatuKakaku() Then
        Exit Function
    End If
    
    'ワークテーブルとデータ作成
    If Not fncMakeWork() Then
        Exit Function
    End If
    
    '再描画
    FKS450_S01.Requery
    
    subFormReload = True

    Exit Function
    
Exception:
    Call Err.Raise(Err.Number, "subFormReload" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncClearWork
'        機能             :ワークテーブルクリア
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncClearWork(strTable As String) As Boolean
    Dim tdfNew      As TableDef
    Dim blnConnect  As Boolean
    Dim isException As Boolean

On Error GoTo Exception
    
    isException = False
    fncClearWork = False
    blnConnect = False
    
    If dbAccess Is Nothing Then
        Set dbAccess = CurrentDb
        blnConnect = True
    End If
        
    FKS450_S01.SourceObject = ""
    
    'ワークテーブル削除
    If fncTableExist(strTable) Then
        Call dbAccess.TableDefs.Delete(strTable)
    End If

    'ワークテーブル作成
    Set tdfNew = dbAccess.CreateTableDef(strTable)
    Call subFieldAppend(tdfNew)
    Call dbAccess.TableDefs.Append(tdfNew)
    
    If Not tdfNew Is Nothing Then Set tdfNew = Nothing
    
    fncClearWork = True
    
    GoTo Finally
    
Exception:
    isException = True
    
Finally:
    If blnConnect Then
        If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    End If
    If Not tdfNew Is Nothing Then Set tdfNew = Nothing
    FKS450_S01.SourceObject = P_SUB_FORM
    If isException Then Call Err.Raise(Err.Number, "fncClearWork" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
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
    
    On Error GoTo Exception

    With tdfNew
        Call .Fields.Append(.CreateField("YARD_CODE", DataTypeEnum.dbText, 6))                      'ヤードコード
        Call .Fields.Append(.CreateField("YARD_NAME", DataTypeEnum.dbText, 36))                     'ヤード名
        Call .Fields.Append(.CreateField("YARD_ADDR", DataTypeEnum.dbText, 36 * 3))                 '所在地
        Call .Fields.Append(.CreateField("TANTM_TANTN", DataTypeEnum.dbMemo, 180))                   'ヤードリーダー
        
        '価格
        Call .Fields.Append(.CreateField("KAKAKU_KKTHM_TCODE", DataTypeEnum.dbText, 20))            '新テンプレートコード
        Call .Fields.Append(.CreateField("KAKAKU_GENZAI_STATUS", DataTypeEnum.dbText, 60))          '現在ステータス
        Call .Fields.Append(.CreateField("KAKAKU_GENZAI_TCODE", DataTypeEnum.dbText, 20))           '現在テンプレCODE
        Call .Fields.Append(.CreateField("KAKAKU_GENZAI_TNAME", DataTypeEnum.dbText, 20))           '現在テンプレNAME
        Call .Fields.Append(.CreateField("KAKAKU_GENZAI_STDTE", DataTypeEnum.dbText, 20))           '現在日付
        Call .Fields.Append(.CreateField("KAKAKU_ZENKAI_STATUS", DataTypeEnum.dbText, 60))          '前回ステータス
        Call .Fields.Append(.CreateField("KAKAKU_ZENKAI_TCODE", DataTypeEnum.dbText, 20))           '前回テンプレCODE
        Call .Fields.Append(.CreateField("KAKAKU_ZENKAI_TNAME", DataTypeEnum.dbText, 20))           '前回テンプレNAME
        Call .Fields.Append(.CreateField("KAKAKU_ZENKAI_STDTE", DataTypeEnum.dbText, 20))           '前回日付
        Call .Fields.Append(.CreateField("KAKAKU_ZENZENKAI_STATUS", DataTypeEnum.dbText, 60))       '前々回ステータス
        Call .Fields.Append(.CreateField("KAKAKU_ZENZENKAI_TCODE", DataTypeEnum.dbText, 20))        '前々回テンプレCODE
        Call .Fields.Append(.CreateField("KAKAKU_ZENZENKAI_TNAME", DataTypeEnum.dbText, 20))        '前々回テンプレNAME
        Call .Fields.Append(.CreateField("KAKAKU_ZENZENKAI_STDTE", DataTypeEnum.dbText, 20))        '前々回日付
        
        '割引
        Call .Fields.Append(.CreateField("WARIBIKI_KKTHM_TCODE", DataTypeEnum.dbText, 20))          '新割引テンプレートコード
        Call .Fields.Append(.CreateField("WARIBIKI_GENZAI_STATUS", DataTypeEnum.dbText, 60))        '現在割引ステータス
        Call .Fields.Append(.CreateField("WARIBIKI_GENZAI_TEXT", DataTypeEnum.dbText, 60))          '現在割引文言
        Call .Fields.Append(.CreateField("WARIBIKI_GENZAI_FROM", DataTypeEnum.dbText, 20))          '現在割引FROM
        Call .Fields.Append(.CreateField("WARIBIKI_GENZAI_TO", DataTypeEnum.dbText, 20))            '現在割引TO
        Call .Fields.Append(.CreateField("WARIBIKI_ZENKAI_STATUS", DataTypeEnum.dbText, 60))        '前回割引ステータス
        Call .Fields.Append(.CreateField("WARIBIKI_ZENKAI_TEXT", DataTypeEnum.dbText, 60))          '前回割引文言
        Call .Fields.Append(.CreateField("WARIBIKI_ZENKAI_FROM", DataTypeEnum.dbText, 20))          '前回割引FROM
        Call .Fields.Append(.CreateField("WARIBIKI_ZENKAI_TO", DataTypeEnum.dbText, 20))            '前回割引TO
        Call .Fields.Append(.CreateField("WARIBIKI_ZENZENKAI_STATUS", DataTypeEnum.dbText, 60))     '前々回割引ステータス
        Call .Fields.Append(.CreateField("WARIBIKI_ZENZENKAI_TEXT", DataTypeEnum.dbText, 60))       '前々回割引文言
        Call .Fields.Append(.CreateField("WARIBIKI_ZENZENKAI_FROM", DataTypeEnum.dbText, 20))       '前々回割引FROM
        Call .Fields.Append(.CreateField("WARIBIKI_ZENZENKAI_TO", DataTypeEnum.dbText, 20))         '前々回割引TO
        
        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount
    End With
    
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subFieldAppend" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncSetCmbIkatuKakaku
'        機能             :価格:一括設定コンボボックスの内容をセット
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSetCmbIkatuKakaku() As Boolean
    Dim objCon              As Object
    Dim objRst              As Object
    Dim strSQL              As String
    Dim strWhere            As String
    Dim isException         As Boolean
    
    isException = False
    fncSetCmbIkatuKakaku = False

On Error GoTo Exception
   
    Set objCon = ADODB_Connection(P_BUMON_CODE)
    
    'SQL
    strSQL = " select KKTHM_TCODE,KKTHM_TNAME" & vbCrLf
    strSQL = strSQL & " from   KKTH_MAST" & vbCrLf
    
    'キーワード地区
    If Nz(Me.txt_Keyword_Tiku_K, "") <> "" Then
        strWhere = " KKTHM_KEYAD like '%" & Replace(Nz(Me.txt_Keyword_Tiku_K, ""), "'", "") & "%'"
    End If
    
    'キーワード価格
    If Nz(Me.txt_Keyword_Kakaku_K, "") <> "" Then
         If strWhere <> "" Then
             strWhere = strWhere & " and"
         End If
         strWhere = strWhere & " KKTHM_KEYPR like '%" & Replace(Nz(Me.txt_Keyword_Kakaku_K, ""), "'", "") & "%'"
    End If
    
    If strWhere <> "" Then
        '2018/08/11 upd strWhere = " where KKTHM_TCODE = '0' or ( " & strWhere & ")" & vbCrLf
        strWhere = " where KKTHM_TCODE = '0' and KKTHM_TTYPE <> '63' or ( " & strWhere & ")" & vbCrLf
    Else
        '2018/08/11 add
        strWhere = " where KKTHM_TTYPE <> '63'" & vbCrLf
    End If
    
    strSQL = strSQL & strWhere
    
    'コンボボックスセット
    cmb_Kakaku_Ikatu_Setei.VALUE = ""
    cmb_Kakaku_Ikatu_Setei.RowSource = ""
    Set objRst = ADODB_Recordset(strSQL, objCon)
    cmb_Kakaku_Ikatu_Setei.AddItem ";"
    With objRst
        While Not .EOF
            cmb_Kakaku_Ikatu_Setei.AddItem .Fields("KKTHM_TNAME") & ";" & .Fields("KKTHM_TCODE")
            .MoveNext
        Wend
    End With
    
    fncSetCmbIkatuKakaku = True
    
    GoTo Finally
    
Exception:
    isException = True
    
Finally:
    If Not objRst Is Nothing Then objRst.Close: Set objRst = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    If isException Then Call Err.Raise(Err.Number, "fncSetCmbIkatuKakaku" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncSetCmbIkatuJimute
'        機能             :事務手数料:一括設定コンボボックスの内容をセット
'        CREATE_DATE      : 2018/08/11
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSetCmbIkatuJimute() As Boolean
    Dim objCon              As Object
    Dim objRst              As Object
    Dim strSQL              As String
    Dim strWhere            As String
    Dim isException         As Boolean
    
    isException = False
    fncSetCmbIkatuJimute = False

On Error GoTo Exception
   
    Set objCon = ADODB_Connection(P_BUMON_CODE)
    
    'SQL
    strSQL = " select KKTHM_TCODE,KKTHM_TNAME" & vbCrLf
    strSQL = strSQL & " from   KKTH_MAST" & vbCrLf

    'キーワード地区
    If Nz(Me.txt_Keyword_Tiku_J, "") <> "" Then
        strWhere = " KKTHM_KEYAD like '%" & Replace(Nz(Me.txt_Keyword_Tiku_J, ""), "'", "") & "%'"
    End If
    
    'キーワード価格
    If Nz(Me.txt_Keyword_Kakaku_J, "") <> "" Then
        If strWhere <> "" Then
            strWhere = strWhere & " and"
        End If
        strWhere = strWhere & " KKTHM_KEYPR like '%" & Replace(Nz(Me.txt_Keyword_Kakaku_J, ""), "'", "") & "%'"
    End If
    
    If strWhere <> "" Then
        '2018/08/11 upd strWhere = " where KKTHM_TCODE = '0' or ( " & strWhere & ")" & vbCrLf
        strWhere = " where KKTHM_TCODE = '0' and KKTHM_TTYPE = '63' or ( " & strWhere & ")" & vbCrLf
    Else
        '2018/08/11 ins
        strWhere = " where KKTHM_TTYPE = '63'" & vbCrLf
    End If
    
    strSQL = strSQL & strWhere
    
    'コンボボックスセット
    cmb_Jimute_Ikatu_Setei.VALUE = ""
    cmb_Jimute_Ikatu_Setei.RowSource = ""
    Set objRst = ADODB_Recordset(strSQL, objCon)
    cmb_Jimute_Ikatu_Setei.AddItem ";"
    With objRst
        While Not .EOF
            cmb_Jimute_Ikatu_Setei.AddItem .Fields("KKTHM_TNAME") & ";" & .Fields("KKTHM_TCODE")
            .MoveNext
        Wend
    End With
    
    fncSetCmbIkatuJimute = True
    
    GoTo Finally
    
Exception:
    isException = True
    
Finally:
    If Not objRst Is Nothing Then objRst.Close: Set objRst = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    If isException Then Call Err.Raise(Err.Number, "fncSetCmbIkatuJimute" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeWork
'        機能             :ワークテーブル作成
'        IN               :True=該当データ無しメッセージを表示 False=表示しない
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeWork() As Boolean
    
    Dim objCon              As Object
    Dim objRstLst           As Object
    Dim objRst450           As Object
    Dim objFld              As Object
    Dim intKadoRitu         As Integer
    Dim lngRecCnt           As Long
    Dim isException As Boolean

On Error GoTo Exception

    isException = False
    fncMakeWork = False
    
    doCmd.Hourglass True
        
    'ACCESS接続
    If dbAccess Is Nothing Then
        Set dbAccess = CurrentDb
    End If
        
    'ワークテーブルクリア
    If Not fncClearWork(P_WORK_TABLE) Then
        GoTo Exception:
    End If

    'ADOコネクト
    Set objCon = ADODB_Connection(P_BUMON_CODE)

    'ワークテーブルレコード追加
    Set objRst450 = dbAccess.OpenRecordset(P_WORK_TABLE, dbOpenDynaset, dbAppendOnly)
    
    'レコードセット作成
    Set objRstLst = ADODB_Recordset(fncGetLstSql(), objCon)
    
    'ワークテーブル書き出し
    With objRstLst
        While Not .EOF
            'トランク・オフィスのチェック 2018/08/11 ins
            Select Case P_BUMON_CODE
            Case "8", "1" ', "H"
                '稼働率チェック
                If fncKadourituChk(objRstLst) = False Then
                    'NG
                    GoTo skip:
                End If
            End Select
            'レコード追加
            lngRecCnt = lngRecCnt + 1
            Call subAddWork(objCon, objRstLst, objRst450)
skip:
            .MoveNext
        Wend
    End With
    
    '取得件数確認
    If lngRecCnt = 0 Then
        MsgBox "該当データが存在しません。", vbInformation, P_KAKUNIN
        GoTo Finally
    End If

    fncMakeWork = True
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If Not objRst450 Is Nothing Then objRst450.Close: Set objRst450 = Nothing
    If Not objRstLst Is Nothing Then objRstLst.Close: Set objRstLst = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    doCmd.Hourglass False
    If isException Then Call Err.Raise(Err.Number, "fncMakeWork" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGetListSql
'        機能             :一覧表示のSQL
'        IN               :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetLstSql() As String
    Dim strSQL              As String
    Dim strWhere            As String
    
On Error GoTo Exception
    
    'SQL
    strSQL = strSQL & " select" & vbCrLf
    strSQL = strSQL & "  YDM.YARD_CODE" & vbCrLf
    strSQL = strSQL & " ,YDM.YARD_NAME" & vbCrLf
    strSQL = strSQL & " ,YDM.YARD_ADDR_1" & vbCrLf
    strSQL = strSQL & " ,YDM.YARD_ADDR_2" & vbCrLf
    
    strSQL = strSQL & " ,ISNULL(YARD_SEV1N,'') AS YARD_SEV1N  " & vbCrLf
    strSQL = strSQL & " ,ISNULL(YARD_SEV2N,'') AS YARD_SEV2N " & vbCrLf
    strSQL = strSQL & " ,ISNULL(YARD_SEV3N,'') AS YARD_SEV3N " & vbCrLf
    strSQL = strSQL & " ,ISNULL(YARD_ENDEN,'') AS YARD_ENDEN " & vbCrLf
    
    
    strSQL = strSQL & " ,TNM.TANTM_TANTN" & vbCrLf
        
    'トランク・オフィスは稼働率の分子、分母を取得 2018/08/11 ins
    Select Case P_BUMON_CODE
    Case "8", "1" ', "H"
        '稼働率(分子)
        strSQL = strSQL & "  ,(     select "
        strSQL = strSQL & "             Count(CARG_YCODE) CARG_COUNT"
        strSQL = strSQL & "         from"
        strSQL = strSQL & "             CARG_FILE"
        strSQL = strSQL & "             inner JOIN YARD_MAST ON YARD_CODE = CARG_YCODE "
        strSQL = strSQL & "         where"
        strSQL = strSQL & "             CARG_YCODE = YDM.YARD_CODE"
        
'BUG 2018/09/08 YAMA        strSQL = strSQL & "             and CARG_AGRE <> 9 "
        strSQL = strSQL & "             and CARG_STDATE <= '" & Format(DateSerial(Year(Now), Month(Now), 0), "yyyy/mm/dd") & "'"
'BUG 2018/09/08 YAMA        strSQL = strSQL & "             and ISNULL(CARG_KYDATE,'9999/12/31') >= '" & Format(Now, "yyyy/mm") & "/01'"
        strSQL = strSQL & "             and ISNULL(CARG_KYDATE,'9999/12/31') >= '" & Format(DateAdd("m", -1, Now), "yyyy/mm") & "/01'"
        
        strSQL = strSQL & "             and ISNULL(YARD_END_DAY,'9999/12/31') > GETDATE() "
        strSQL = strSQL & "    ) as KADOU_BUNSI"
        strSQL = strSQL & "    "
        '稼働率(分母)
        strSQL = strSQL & "  ,(     select "
        strSQL = strSQL & "             Count(CNTA_NO) CNTA_COUNT"
        strSQL = strSQL & "         from "
        strSQL = strSQL & "             CNTA_MAST "
        strSQL = strSQL & "             inner  JOIN YARD_MAST "
        strSQL = strSQL & "             ON CNTA_CODE = YARD_CODE  "
        strSQL = strSQL & "         where  CNTA_CODE = YDM.YARD_CODE"
        strSQL = strSQL & "             and CNTA_USE <> 9 "
        strSQL = strSQL & "             and ISNULL(YARD_END_DAY,'9999/12/31') > GETDATE() "
        strSQL = strSQL & "    ) as KADOU_BUNBO" '分母
    End Select
    
    'テンプレート
    strSQL = strSQL & " ,(   select count(*)" & vbCrLf
    strSQL = strSQL & "      from KTPS_TRAN" & vbCrLf
    strSQL = strSQL & "      where KTPST_YCODE = RIGHT('000000' + convert(varchar,YDM.YARD_CODE), 6)" & vbCrLf
    strSQL = strSQL & "  ) as KAKAKU_RIREKI_COUNT" & vbCrLf
    
    '割引(現在)
    strSQL = strSQL & " ,(   select count(*)" & vbCrLf
    strSQL = strSQL & "      from" & vbCrLf
    strSQL = strSQL & "          RLDN_TRAN" & vbCrLf
    strSQL = strSQL & "      where" & vbCrLf
    strSQL = strSQL & "         ((RLDNT_ENABLE = '1' and CONVERT(varchar(8),GETDATE(), 112) between RLDNT_FROM and RLDNT_TO)" & vbCrLf
    strSQL = strSQL & "      or  (RLDNT_ENABLE = '1' and CONVERT(varchar(8),GETDATE(), 112) < RLDNT_FROM))" & vbCrLf
    strSQL = strSQL & "      and RLDNT_BUMOC = '" & P_BUMON_CODE & "'" & vbCrLf
    strSQL = strSQL & "      and RLDNT_YCODE = RIGHT('000000' + convert(varchar,YDM.YARD_CODE), 6)" & vbCrLf
    strSQL = strSQL & "  ) as WARIBIKI_GENZAI_CNT" & vbCrLf
    
    '割引(前回)
    strSQL = strSQL & "  ,(  select count(*)" & vbCrLf
    strSQL = strSQL & "      from" & vbCrLf
    strSQL = strSQL & "          RLDN_TRAN" & vbCrLf
    strSQL = strSQL & "      where" & vbCrLf
    strSQL = strSQL & "      (RLDNT_ENABLE = '0' or  CONVERT(varchar(8),GETDATE(), 112) > RLDNT_TO)"
    strSQL = strSQL & "      and RLDNT_BUMOC = '" & P_BUMON_CODE & "'" & vbCrLf
    strSQL = strSQL & "      and RLDNT_YCODE = RIGHT('000000' + convert(varchar,YDM.YARD_CODE), 6)" & vbCrLf
    strSQL = strSQL & "  ) as WARIBIKI_ZENKAI_CNT" & vbCrLf
    
    strSQL = strSQL & " from YARD_MAST as YDM" & vbCrLf
    strSQL = strSQL & " left join " & strSqlserver & ".dbo.TANT_MAST as TNM" & vbCrLf
    strSQL = strSQL & " on (TNM.TANTM_TANTC = YARD_TANTO and TNM.TANTM_BUMOC = '" & P_BUMON_CODE & "')" & vbCrLf
    strWhere = "where ISNULL(CONVERT(varchar(8),YARD_RENTEND_DAY, 112), '99991231') > CONVERT(varchar(8),GETDATE(), 112) " & vbCrLf
    
    'ヤード
    If Nz(Me.cmb_Yard_from, "") <> "" And Nz(Me.cmb_Yard_To, "") <> "" Then
        If Me.cmb_Yard_from < Me.cmb_Yard_To Then
            strWhere = strWhere & " and YDM.YARD_CODE between " & Val(Me.cmb_Yard_from) & " and " & Val(Me.cmb_Yard_To) & vbCrLf
        Else
            strWhere = strWhere & " and YDM.YARD_CODE between " & Val(Me.cmb_Yard_To) & " and " & Val(Me.cmb_Yard_from) & vbCrLf
        End If
    ElseIf Nz(Me.cmb_Yard_from, "") <> "" Then
        strWhere = strWhere & " and YDM.YARD_CODE >= " & Val(Me.cmb_Yard_from) & vbCrLf
    ElseIf Nz(Me.cmb_Yard_To, "") <> "" Then
        strWhere = strWhere & " and YDM.YARD_CODE <= " & Val(Me.cmb_Yard_To) & vbCrLf
    End If
    
    '地区
    If Nz(Me.txt_Todoufuken, "") <> "" Then
        strWhere = strWhere & " and (YARD_ADDR_1 like '%" & Replace(Me.txt_Todoufuken, "'", "") & "%' or YARD_ADDR_2 like '%" & Replace(Me.txt_Todoufuken, "'", "") & "%')" & vbCrLf
    End If
    
    'ヤードリーダー
    If Nz(Me.cmb_Yard_Leader, "") <> "" Then
        strWhere = strWhere & " and YARD_TANTO = '" & Me.cmb_Yard_Leader & "'" & vbCrLf
    End If
    
    '条件
    strSQL = strSQL & strWhere
    
    '並び順
    strSQL = strSQL & " order by YDM.YARD_CODE" & vbCrLf
    
    '戻り
    fncGetLstSql = strSQL
    
    Exit Function
Exception:
    Call Err.Raise(Err.Number, "fncGetLstSql" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :subSetCmb
'        機能             :一括設定コンボボックス
'        IN               :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetCmb()
    Dim i As Integer
On Error GoTo Exception
    If tab_Kensaku.VALUE = 0 Then '2018/08/11 ins
        '価格
        With Me.FKS450_S01.Form!cmb_Kakaku_New_Setei
            .RowSource = ""
            For i = 0 To cmb_Kakaku_Ikatu_Setei.ListCount - 1
                .AddItem Me.cmb_Kakaku_Ikatu_Setei.Column(0, i) & ";" & Me.cmb_Kakaku_Ikatu_Setei.Column(1, i)
            Next
        End With
    Else
        '事務手数料 2018/08/11 ins
        With Me.FKS450_S01.Form!cmb_Kakaku_New_Setei
            .RowSource = ""
            For i = 0 To cmb_Jimute_Ikatu_Setei.ListCount - 1
                .AddItem Me.cmb_Jimute_Ikatu_Setei.Column(0, i) & ";" & Me.cmb_Jimute_Ikatu_Setei.Column(1, i)
            Next
        End With
    End If
    
    '割引
    With Me.FKS450_S01.Form!cmb_Waribiki_New_Setei
        .RowSource = ""
        For i = 0 To cmb_Waribiki_Ikatu_Setei.ListCount - 1
            .AddItem Me.cmb_Waribiki_Ikatu_Setei.Column(0, i) & ";" & Me.cmb_Waribiki_Ikatu_Setei.Column(1, i)
        Next
    End With
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subSetCmb" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncTableExist
'        機能             :ACCESSテーブル存在チェック
'        IN               :テーブル名
'        OUT              :True=存在する False=存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncTableExist(strTableName As String) As Boolean

    Dim tdf As TableDef
    
On Error GoTo Exception
    
    fncTableExist = False
    
    'ローカルDB
    If dbAccess Is Nothing Then
        Set dbAccess = CurrentDb
    End If
    
    For Each tdf In dbAccess.TableDefs
        If tdf.NAME = strTableName Then
            fncTableExist = True
            Exit For
        End If
    Next tdf

    Exit Function
Exception:
    Call Err.Raise(Err.Number, "fncTableExist" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'   MODULE_NAME     :   ワークテーブルデータ削除
'   Parameter       :   lngFLG    再描画    0=しない、1=する
'   CREATE_DATE     :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDelWorkTable(ByVal lngFLG As Long)
    
    On Error GoTo Exception
    
    '警告非表示
    doCmd.SetWarnings False
    
    'SQL実行
    doCmd.RunSQL "DELETE FROM " & P_WORK_TABLE
    
    '警告表示
    doCmd.SetWarnings True
    
    If 1 = lngFLG Then
        '再描画
        FKS450_S01.Requery
        
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "subDelWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :明細 価格SQL(現在、前回、前々回)
'        機能             :fncGetLstWaribikiNewSql
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetLstKakakuSql(strYcode As String) As String
    Dim strSQL          As String
    strSQL = strSQL & " select TOP 3" & vbCrLf
    strSQL = strSQL & "     KTPST_TCODE , KTPST_STDTE , KKTHM_TNAME" & vbCrLf
    strSQL = strSQL & " from KTPS_TRAN" & vbCrLf
    strSQL = strSQL & " inner join KKTH_MAST on KTPST_TCODE = KKTHM_TCODE" & vbCrLf
    strSQL = strSQL & " where KTPST_YCODE = '" & strYcode & "'" & vbCrLf
    strSQL = strSQL & " and   KTPST_TTYPE = '" & IIf(tab_Kensaku.VALUE = 0, "0", "63") & "'" & vbCrLf '2018/08/11 ins テンプレートタイプ
    strSQL = strSQL & " order by KTPST_STDTE desc" & vbCrLf
    fncGetLstKakakuSql = strSQL
End Function

'==============================================================================*
'
'        MODULE_NAME      :明細 割引SQL(現在)
'        機能             :fncGetLstWaribikiNewSql
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetLstWaribikiNewSql(strYcode) As String
    Dim strSQL          As String
    strSQL = strSQL & " select top 1" & vbCrLf
    strSQL = strSQL & "      RLDNT_TEXT" & vbCrLf
    strSQL = strSQL & "     ,RLDNT_FROM" & vbCrLf
    strSQL = strSQL & "     ,RLDNT_TO" & vbCrLf
    strSQL = strSQL & " from" & vbCrLf
    strSQL = strSQL & "     RLDN_TRAN" & vbCrLf
    strSQL = strSQL & " where" & vbCrLf
    strSQL = strSQL & " ((  RLDNT_ENABLE = '1' and CONVERT(varchar(8),GETDATE(), 112) between RLDNT_FROM and RLDNT_TO)" & vbCrLf
    strSQL = strSQL & " or (RLDNT_ENABLE = '1' and CONVERT(varchar(8),GETDATE(), 112) < RLDNT_FROM))" & vbCrLf
    strSQL = strSQL & " and RLDNT_BUMOC = '" & P_BUMON_CODE & "'" & vbCrLf
    strSQL = strSQL & " and RLDNT_YCODE = '" & strYcode & "'" & vbCrLf & vbCrLf
    strSQL = strSQL & " Order by" & vbCrLf
    strSQL = strSQL & "      RLDNT_FROM" & vbCrLf
    strSQL = strSQL & "     ,RLDNT_GNKBN desc" & vbCrLf
    strSQL = strSQL & "     ,RLDNT_ORDER asc" & vbCrLf
    fncGetLstWaribikiNewSql = strSQL
End Function

'==============================================================================*
'
'        MODULE_NAME      :明細 割引SQL(前回、前々回)
'        機能             :fncGetLstWaribikiZenSql
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetLstWaribikiZenSql(strYcode) As String
    Dim strSQL          As String
    strSQL = strSQL & " select top 2" & vbCrLf
    strSQL = strSQL & " min (RLDNT_TEXT) as RLDNT_TEXT" & vbCrLf
    strSQL = strSQL & " ,RLDNT_FROM" & vbCrLf
    strSQL = strSQL & " ,RLDNT_TO" & vbCrLf
    strSQL = strSQL & " from" & vbCrLf
    strSQL = strSQL & " RLDN_TRAN" & vbCrLf
    strSQL = strSQL & " where" & vbCrLf
    strSQL = strSQL & " (" & vbCrLf
    strSQL = strSQL & "     RLDNT_ENABLE = '0'" & vbCrLf
    strSQL = strSQL & " or CONVERT(varchar(8),GETDATE(), 112) > RLDNT_TO" & vbCrLf
    strSQL = strSQL & " )" & vbCrLf
    strSQL = strSQL & " and RLDNT_BUMOC = '" & P_BUMON_CODE & "'" & vbCrLf
    strSQL = strSQL & " and RLDNT_YCODE = '" & strYcode & "'" & vbCrLf & vbCrLf
    strSQL = strSQL & " Group by" & vbCrLf
    strSQL = strSQL & "    RLDNT_FROM , RLDNT_TO" & vbCrLf
    strSQL = strSQL & " Order by" & vbCrLf
    strSQL = strSQL & "     RLDNT_FROM desc, RLDNT_TO desc" & vbCrLf
    fncGetLstWaribikiZenSql = strSQL
End Function

'==============================================================================*
'
'        MODULE_NAME      :subAddWork
'        機能             :作業テーブル追加
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subAddWork(objCon As Object, objRstLst As Object, objRst450 As Object)
    Dim objRstRireki    As Object
    Dim intRirekiCnt    As Integer
    Dim strSQL          As String
    Dim strYcode        As String
    Dim strText         As String
    Dim strForm         As String
    Dim strTo           As String
    Dim i               As Integer
On Error GoTo Exception
    'ワーク格納
    With objRst450
        .AddNew
        strYcode = Format(objRstLst.Fields("YARD_CODE"), "000000")
        .Fields("YARD_CODE") = strYcode
        .Fields("YARD_NAME") = objRstLst.Fields("YARD_NAME")
        .Fields("YARD_ADDR") = objRstLst.Fields("YARD_SEV1N") & objRstLst.Fields("YARD_SEV2N") & objRstLst.Fields("YARD_SEV3N") & objRstLst.Fields("YARD_ENDEN")
        '2018/08/11 トラ君・オフィス対応
        Select Case P_BUMON_CODE
        Case "8", "1" ', "H"
            '稼働率
            If objRstLst.Fields("KADOU_BUNBO") <> 0 Then
                .Fields("TANTM_TANTN") = Int((objRstLst.Fields("KADOU_BUNSI") / objRstLst.Fields("KADOU_BUNBO")) * 100)
            Else
                .Fields("TANTM_TANTN") = 0
            End If
        Case Else
            'ヤードリーダー
            .Fields("TANTM_TANTN") = objRstLst.Fields("TANTM_TANTN")
        End Select
        
        .Fields("KAKAKU_GENZAI_STATUS") = "－"
        .Fields("KAKAKU_ZENKAI_STATUS") = "－"
        .Fields("KAKAKU_ZENZENKAI_STATUS") = "－"
        '最新、前回、前々回(価格の履歴が1件以上あるときのみ検索)
        If objRstLst.Fields("KAKAKU_RIREKI_COUNT") <> 0 Then
            '価格履歴トラン
            Set objRstRireki = ADODB_Recordset(fncGetLstKakakuSql(strYcode), objCon)
            While Not objRstRireki.EOF
                intRirekiCnt = intRirekiCnt + 1
                Select Case intRirekiCnt
                Case 1 '最新
                    .Fields("KAKAKU_GENZAI_STATUS") = objRstRireki.Fields("KKTHM_TNAME") & vbCrLf & Format(objRstRireki.Fields("KTPST_STDTE"), "YYYY/MM/DD") & "～"
                    .Fields("KAKAKU_GENZAI_TCODE") = objRstRireki.Fields("KTPST_TCODE")
                    .Fields("KAKAKU_GENZAI_TNAME") = objRstRireki.Fields("KKTHM_TNAME")
                    .Fields("KAKAKU_GENZAI_STDTE") = Format(objRstRireki.Fields("KTPST_STDTE"), "YYYY/MM/DD")
                Case 2 '前回
                    .Fields("KAKAKU_ZENKAI_STATUS") = objRstRireki.Fields("KKTHM_TNAME") & vbCrLf & Format(objRstRireki.Fields("KTPST_STDTE"), "YYYY/MM/DD") & "～"
                    .Fields("KAKAKU_ZENKAI_TCODE") = objRstRireki.Fields("KTPST_TCODE")
                    .Fields("KAKAKU_ZENKAI_TNAME") = objRstRireki.Fields("KKTHM_TNAME")
                    .Fields("KAKAKU_ZENKAI_STDTE") = Format(objRstRireki.Fields("KTPST_STDTE"), "YYYY/MM/DD")
                Case 3 '前々回
                    .Fields("KAKAKU_ZENZENKAI_STATUS") = objRstRireki.Fields("KKTHM_TNAME") & vbCrLf & Format(objRstRireki.Fields("KTPST_STDTE"), "YYYY/MM/DD") & "～"
                    .Fields("KAKAKU_ZENZENKAI_TCODE") = objRstRireki.Fields("KTPST_TCODE")
                    .Fields("KAKAKU_ZENZENKAI_TNAME") = objRstRireki.Fields("KKTHM_TNAME")
                    .Fields("KAKAKU_ZENZENKAI_STDTE") = Format(objRstRireki.Fields("KTPST_STDTE"), "YYYY/MM/DD")
                End Select
                objRstRireki.MoveNext
            Wend
            objRstRireki.Close
        End If
        
        '最新(割引の履歴が1件以上あるときのみ検索)
        If objRstLst.Fields("WARIBIKI_GENZAI_CNT") <> 0 Then
            '割引履歴トラン
            Set objRstRireki = ADODB_Recordset(fncGetLstWaribikiNewSql(strYcode), objCon)
            Do While Not objRstRireki.EOF
                '最新
                strText = Nz(objRstRireki.Fields("RLDNT_TEXT"), "")
                strForm = fncDateFormat(objRstRireki.Fields("RLDNT_FROM"))
                strTo = fncDateFormat(objRstRireki.Fields("RLDNT_TO"))
                .Fields("WARIBIKI_GENZAI_STATUS") = Trim(CmMidB(strText, 1, 20)) & vbCrLf & strForm & "～" & strTo
                .Fields("WARIBIKI_GENZAI_TEXT") = strText
                .Fields("WARIBIKI_GENZAI_FROM") = strForm
                .Fields("WARIBIKI_GENZAI_TO") = strTo
                Exit Do
            Loop
            objRstRireki.Close
        End If
        
        '前回、前々回(割引の前回前々回の履歴が1件以上あるときのみ検索)
        If objRstLst.Fields("WARIBIKI_ZENKAI_CNT") <> 0 Then
            intRirekiCnt = 0
            '割引履歴トラン
            Set objRstRireki = ADODB_Recordset(fncGetLstWaribikiZenSql(strYcode), objCon)
            While Not objRstRireki.EOF
                intRirekiCnt = intRirekiCnt + 1
                strText = Nz(objRstRireki.Fields("RLDNT_TEXT"), "")
                strForm = fncDateFormat(objRstRireki.Fields("RLDNT_FROM"))
                strTo = fncDateFormat(objRstRireki.Fields("RLDNT_TO"))
                Select Case intRirekiCnt
                Case 1 '前回
                    .Fields("WARIBIKI_ZENKAI_STATUS") = Trim(CmMidB(strText, 1, 20)) & vbCrLf & strForm & "～" & strTo
                    .Fields("WARIBIKI_ZENKAI_TEXT") = strText
                    .Fields("WARIBIKI_ZENKAI_FROM") = strForm
                    .Fields("WARIBIKI_ZENKAI_TO") = strTo
                Case 2 '前々回
                    .Fields("WARIBIKI_ZENZENKAI_STATUS") = Trim(CmMidB(strText, 1, 20)) & vbCrLf & strForm & "～" & strTo
                    .Fields("WARIBIKI_ZENZENKAI_TEXT") = strText
                    .Fields("WARIBIKI_ZENZENKAI_FROM") = strForm
                    .Fields("WARIBIKI_ZENZENKAI_TO") = strTo
                End Select
                objRstRireki.MoveNext
            Wend
            objRstRireki.Close
        End If
        '更新
        .UPDATE
    End With
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subAddWork" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncDateFormat
'        機能             :日付表示(/編集)
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDateFormat(objTgtDate As Object) As String
    Dim strRet As String
    fncDateFormat = ""
    strRet = Nz(objTgtDate)
    If Trim(strRet) = "" Then Exit Function
    fncDateFormat = Mid(strRet, 1, 4) & "/" & Mid(strRet, 5, 2) & "/" & Mid(strRet, 7, 2)
End Function


'==============================================================================*
'
'        MODULE_NAME      :Form_Open
'        機能             :Form_Openイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    doCmd.Maximize

    Call MSHS010_M00(PROG_ID, "1")
    
        Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル2.BackColor = RGB(255, 0, 204)
    End Select

    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Unload
'        機能             :Form_Unloadイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    'Access側ワークテーブル削除
    Call fncClearWork(P_WORK_TABLE)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subKikanFrom_BeforeUpdate
'        機能             :subKikanFrom_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKikanFrom_BeforeUpdate(objActiveControl As Object, ByRef intCancel As Integer)
On Error GoTo Exception
    If Nz(objActiveControl.VALUE) <> "" Then
        P_DATE_VALUE = Nz(CmDateChecker(Nz(objActiveControl.VALUE)))
        If P_DATE_VALUE = "" Then
            MsgBox "実施期間はYYYY/MM/DD形式です。"
            intCancel = True
            SendKeys "{ESC}"
        End If
    End If
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subKikanFrom_BeforeUpdate" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subKikanFrom_AfterUpdate
'        機能             :subKikanFrom_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKikanFrom_AfterUpdate(objActiveControl As Object)
    objActiveControl.VALUE = P_DATE_VALUE
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subMeisaiLbl
'        機能             :明細行のラベル切り替え
'        CREATE DATE      : 2018/08/11
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subMeisaiLbl()
    Select Case P_BUMON_CODE
    Case "8", "1" ', "H"
        Me.FKS450_S01.Form!lbl_YardL_Jimute.Caption = "稼働率"
        Me.FKS450_S01.Form!txt_Yard_Leader.Visible = False
        Me.FKS450_S01.Form!txt_Kadouritu.Visible = True
    Case Else
        Me.FKS450_S01.Form!lbl_YardL_Jimute.Caption = "ﾔｰﾄﾞﾘｰﾀﾞｰ"
        Me.FKS450_S01.Form!txt_Yard_Leader.Visible = True
        Me.FKS450_S01.Form!txt_Kadouritu.Visible = False
    End Select
    FKS450_S01.Requery
End Sub

'==============================================================================*
'
'        MODULE_NAME      :tab_Kensaku_Change
'        機能             :検索モードの切替
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub tab_Kensaku_Change()

    'タブを戻す際に重複イベント回避
    If tab_Kensaku_flg = True Then
        tab_Kensaku_flg = False
        Exit Sub
    End If

    '新設定が設定されているかチェック
    If fncCheckMeisaiCmb() = False Then
        '新設定値が設定されている
        If MsgBox("明細の新設定値リストが選択されています。" & vbCrLf & "選択内容は破棄されますが宜しいでしょうか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbNo Then
            tab_Kensaku_flg = True
            tab_Kensaku.VALUE = IIf(tab_Kensaku.VALUE = 0, 1, 0)
            Exit Sub
        End If
    End If

    'ワークテーブルとデータ作成
    Call subDelWorkTable(1)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_From_1_BeforeUpdate
'        機能             :txt_Kikan_From_1_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_From_1_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subKikanFrom_BeforeUpdate(ActiveControl, Cancel)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_From_1_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_From_1_AfterUpdate
'        機能             :txt_Kikan_From_1_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_From_1_AfterUpdate()
On Error GoTo Exception
    Call subKikanFrom_AfterUpdate(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_From_1_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_From_2_BeforeUpdate
'        機能             :txt_Kikan_From_2_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_From_2_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subKikanFrom_BeforeUpdate(ActiveControl, Cancel)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_From_2_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_From_2_AfterUpdate
'        機能             :txt_Kikan_From_2_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_From_2_AfterUpdate()
On Error GoTo Exception
    Call subKikanFrom_AfterUpdate(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_From_2_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_From_3_BeforeUpdate
'        機能             :txt_Kikan_From_3_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_From_3_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subKikanFrom_BeforeUpdate(ActiveControl, Cancel)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_From_3_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_From_3_AfterUpdate
'        機能             :txt_Kikan_From_3_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_From_3_AfterUpdate()
On Error GoTo Exception
    Call subKikanFrom_AfterUpdate(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_From_3_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_From_4_BeforeUpdate
'        機能             :txt_Kikan_From_4_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_From_4_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subKikanFrom_BeforeUpdate(ActiveControl, Cancel)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_From_4_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_From_4_AfterUpdate
'        機能             :txt_Kikan_From_4_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_From_4_AfterUpdate()
On Error GoTo Exception
    Call subKikanFrom_AfterUpdate(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_From_4_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_From_5_BeforeUpdate
'        機能             :txt_Kikan_From_5_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_From_5_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subKikanFrom_BeforeUpdate(ActiveControl, Cancel)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_From_5_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_From_5_AfterUpdate
'        機能             :txt_Kikan_From_5_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_From_5_AfterUpdate()
On Error GoTo Exception
    Call subKikanFrom_AfterUpdate(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_From_5_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subKikanTo_BeforeUpdate
'        機能             :subKikanTo_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKikanTo_BeforeUpdate(objActiveControl As Object, ByRef intCancel As Integer)
On Error GoTo Exception
    If Nz(objActiveControl.VALUE) <> "" Then
        P_DATE_VALUE = Nz(CmDateChecker(Nz(objActiveControl.VALUE)))
        If P_DATE_VALUE = "" Then
            MsgBox "実施期間はYYYY/MM/DD形式です。"
            intCancel = True
            SendKeys "{ESC}"
        End If
    End If
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subKikanTo_BeforeUpdate" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subKikanTo_AfterUpdate
'        機能             :subKikanTo_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKikanTo_AfterUpdate(objActiveControl As Object)
    objActiveControl.VALUE = P_DATE_VALUE
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_To_1_AfterUpdate
'        機能             :txt_Kikan_To_1_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_To_1_AfterUpdate()
On Error GoTo Exception
    Call subKikanTo_AfterUpdate(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_To_1_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_To_1_BeforeUpdate
'        機能             :txt_Kikan_To_1_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_To_1_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subKikanTo_BeforeUpdate(ActiveControl, Cancel)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_To_1_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_To_2_AfterUpdate
'        機能             :txt_Kikan_To_2_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_To_2_AfterUpdate()
On Error GoTo Exception
    Call subKikanTo_AfterUpdate(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_To_2_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_To_2_BeforeUpdate
'        機能             :txt_Kikan_To_2_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_To_2_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subKikanTo_BeforeUpdate(ActiveControl, Cancel)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_To_2_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_To_3_AfterUpdate
'        機能             :txt_Kikan_To_3_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_To_3_AfterUpdate()
On Error GoTo Exception
    Call subKikanTo_AfterUpdate(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_To_3_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_To_3_BeforeUpdate
'        機能             :txt_Kikan_To_3_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_To_3_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subKikanTo_BeforeUpdate(ActiveControl, Cancel)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_To_3_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_To_4_AfterUpdate
'        機能             :txt_Kikan_To_4_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_To_4_AfterUpdate()
On Error GoTo Exception
    Call subKikanTo_AfterUpdate(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_To_4_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_To_4_BeforeUpdate
'        機能             :txt_Kikan_To_4_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_To_4_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subKikanTo_BeforeUpdate(ActiveControl, Cancel)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_To_4_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_To_5_AfterUpdate
'        機能             :txt_Kikan_To_5_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_To_5_AfterUpdate()
On Error GoTo Exception
    Call subKikanTo_AfterUpdate(ActiveControl)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_To_5_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kikan_To_5_BeforeUpdate
'        機能             :txt_Kikan_To_5_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kikan_To_5_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subKikanTo_BeforeUpdate(ActiveControl, Cancel)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Kikan_To_5_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CmBCheck_Char
'        機能             :文字数判断①
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subLenChk(objActiveControl As Object, ByRef intCancel As Integer, intSize As Integer)
    Dim vCode       As Variant

On Error GoTo Exception
      
    vCode = Nz(objActiveControl, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, intSize)
        Case 0
    
        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。(MAX:" & CStr(intSize) & ")", vbInformation, P_ERROR
            intCancel = True
        Case 9
    End Select

    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subLenChk" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CmBCheck_Char
'        機能             :文字数判断②
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer

On Error GoTo Exception

    CmBCheck_Char = 9

    If CmLenb(oText) > ilen Then
       CmBCheck_Char = 2
       GoTo CmBCheck_Char_Exit
    End If

    CmBCheck_Char = 0

CmBCheck_Char_Exit:
    Exit Function

Exception:
    CmBCheck_Char = 9
    Call Err.Raise(Err.Number, "CmBCheck_Char" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :txt_Mongon_1_BeforeUpdate
'        機能             :txt_Mongon_1_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mongon_1_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 42)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Mongon_1_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Mongon_2_BeforeUpdate
'        機能             :txt_Mongon_2_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mongon_2_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 42)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Mongon_2_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Mongon_3_BeforeUpdate
'        機能             :txt_Mongon_3_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mongon_3_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 42)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Mongon_2_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Mongon_4_BeforeUpdate
'        機能             :txt_Mongon_4_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mongon_4_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 42)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Mongon_2_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Mongon_5_BeforeUpdate
'        機能             :txt_Mongon_5_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mongon_5_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 42)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Mongon_5_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Biko_1_BeforeUpdate
'        機能             :txt_Biko_1_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Biko_1_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Biko_1_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Biko_2_BeforeUpdate
'        機能             :txt_Biko_2_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Biko_2_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Biko_2_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Biko_3_BeforeUpdate
'        機能             :txt_Biko_3_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Biko_3_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Biko_3_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Biko_4_BeforeUpdate
'        機能             :txt_Biko_4_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Biko_4_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Biko_4_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Biko_5_BeforeUpdate
'        機能             :txt_Biko_5_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Biko_5_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Biko_5_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_1_1_BeforeUpdate
'        機能             :txt_Sabisu_1_1_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_1_1_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 80)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_1_1_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_1_2_BeforeUpdate
'        機能             :txt_Sabisu_1_2_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_1_2_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 80)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_1_2_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_1_3_BeforeUpdate
'        機能             :txt_Sabisu_1_3_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_1_3_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 80)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_1_3_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_1_4_BeforeUpdate
'        機能             :txt_Sabisu_1_4_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_1_4_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_1_4_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_1_5_BeforeUpdate
'        機能             :txt_Sabisu_1_5_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_1_5_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_1_5_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_2_1_BeforeUpdate
'        機能             :txt_Sabisu_2_1_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_2_1_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 80)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_2_1_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_2_2_BeforeUpdate
'        機能             :txt_Sabisu_2_2_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_2_2_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 80)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_2_2_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_2_3_BeforeUpdate
'        機能             :txt_Sabisu_2_3_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_2_3_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 80)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_2_3_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_2_4_BeforeUpdate
'        機能             :txt_Sabisu_2_4_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_2_4_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_2_4_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_2_5_BeforeUpdate
'        機能             :txt_Sabisu_2_5_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_2_5_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_2_5_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_3_1_BeforeUpdate
'        機能             :txt_Sabisu_3_1_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_3_1_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 80)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_3_1_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_3_2_BeforeUpdate
'        機能             :txt_Sabisu_3_2_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_3_2_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 80)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_3_2_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_3_3_BeforeUpdate
'        機能             :txt_Sabisu_3_3_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_3_3_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 80)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_3_3_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_3_4_BeforeUpdate
'        機能             :txt_Sabisu_3_4_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_3_4_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_3_4_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_3_5_BeforeUpdate
'        機能             :txt_Sabisu_3_5_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_3_5_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_3_5_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Kikan_1_BeforeUpdate
'        機能             :txt_Sabisu_Kikan_1_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Kikan_1_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_Kikan_1_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Kikan_2_BeforeUpdate
'        機能             :txt_Sabisu_Kikan_2_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Kikan_2_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_Kikan_2_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Kikan_3_BeforeUpdate
'        機能             :txt_Sabisu_Kikan_3_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Kikan_3_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_Kikan_3_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Kikan_4_BeforeUpdate
'        機能             :txt_Sabisu_Kikan_4_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Kikan_4_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_Kikan_4_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Kikan_5_BeforeUpdate
'        機能             :txt_Sabisu_Kikan_5_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Kikan_5_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 60)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_Kikan_5_BeforeUpdate", True)
End Sub

'add strat 2018/11/15 ここから
'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Tekiyou_1_BeforeUpdate
'        機能             :txt_Sabisu_Tekiyou_1_BeforeUpdateeイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Tekiyou_1_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 3)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_Tekiyou_1_BeforeUpdate", True)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Tekiyou_1_KeyPress
'        機能             :txt_Sabisu_Tekiyou_1_KeyPressイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Tekiyou_1_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_Sabisu_Tekiyou_1, KeyAscii, 3)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Tekiyou_1_BeforeUpdate
'        機能             :txt_Sabisu_Tekiyou_1_BeforeUpdateeイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Tekiyou_2_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 3)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_Tekiyou_2_BeforeUpdate", True)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Tekiyou_2_KeyPress
'        機能             :txt_Sabisu_Tekiyou_2_KeyPressイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Tekiyou_2_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_Sabisu_Tekiyou_2, KeyAscii, 3)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Tekiyou_3_BeforeUpdate
'        機能             :txt_Sabisu_Tekiyou_3_BeforeUpdateeイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Tekiyou_3_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 3)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_Tekiyou_3_BeforeUpdate", True)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Tekiyou_3_KeyPress
'        機能             :txt_Sabisu_Tekiyou_3_KeyPressイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Tekiyou_3_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_Sabisu_Tekiyou_3, KeyAscii, 3)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Tekiyou_4_BeforeUpdate
'        機能             :txt_Sabisu_Tekiyou_4_BeforeUpdateeイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Tekiyou_4_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 3)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_Tekiyou_4_BeforeUpdate", True)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Tekiyou_4_KeyPress
'        機能             :txt_Sabisu_Tekiyou_4_KeyPressイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Tekiyou_4_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_Sabisu_Tekiyou_4, KeyAscii, 3)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Tekiyou_5_BeforeUpdate
'        機能             :txt_Sabisu_Tekiyou_5_BeforeUpdateeイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Tekiyou_5_BeforeUpdate(Cancel As Integer)
On Error GoTo Exception
    Call subLenChk(ActiveControl, Cancel, 3)
    Exit Sub
Exception:
    Call MSZZ024_M00("txt_Sabisu_Tekiyou_5_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Sabisu_Tekiyou_5_KeyPress
'        機能             :txt_Sabisu_Tekiyou_5_KeyPressイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Sabisu_Tekiyou_5_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_Sabisu_Tekiyou_5, KeyAscii, 3)
End Sub


'==============================================================================*
'
'        MODULE_NAME      :fncGet割引適用利用月数
'        機能             :割引適用利用月数の取得
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGet割引適用利用月数(aWaribikiNo As String) As Variant
    Dim objCon          As Object
    Dim objWaribiki     As Object
    Dim strSQL          As String
    Dim isException     As Boolean
    isException = False

On Error GoTo Exception

    fncGet割引適用利用月数 = ""
    
    strSQL = strSQL & " select"
    strSQL = strSQL & " case DCNTM_USE_PERIOD"
    strSQL = strSQL & " when '1' then 1"    '1:二カ月以下は 1
    strSQL = strSQL & " when '2' then 3"    '2:三カ月以上は 3
    strSQL = strSQL & " when '3' then 6"    '3:半年以上は   6
    strSQL = strSQL & " when '4' then 12"   '4:一年以上は   12
    strSQL = strSQL & " else 0"             '5:制限なしは   0
    strSQL = strSQL & " end DCNTM_USE_PERIOD"
    strSQL = strSQL & " from   DCNT_MAST"
    strSQL = strSQL & " where  DCNTM_NO = " & aWaribikiNo

    'ADOコネクト
    Set objCon = ADODB_Connection()

    'レコードセット作成
    Set objWaribiki = ADODB_Recordset(strSQL, objCon)
    
    '値の取得
    With objWaribiki
        Do While Not .EOF
           fncGet割引適用利用月数 = Nz(.Fields("DCNTM_USE_PERIOD"), 0)
           Exit Do
        Loop
    End With

    GoTo Finally
    
Exception:
    isException = True
    
Finally:
    If Not objWaribiki Is Nothing Then objWaribiki.Close: Set objWaribiki = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    If isException Then Call Err.Raise(Err.Number, "fncGet割引適用利用月数" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function







