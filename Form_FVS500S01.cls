VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS500S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　予約入力サブフォーム
'   プログラムＩＤ　：　FVS500S01
'   作　成　日　　　：  2005/05/06
'   作　成　者　　　：  ChronoSystem 三ツ野
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/10/06
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  フォームのレコードソースに直接SQLを実行するのではなく、
'                      ワークテーブルの内容を表示するように修正
'   Ver             ：  0.1
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/05/19
'   修　正　者　　　：  tajima
'   修　正　内　容　：  不具合対応
'   Ver             ：  0.2
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/06/14
'   修　正　者　　　：  tajima
'   修　正　内　容　：  カード対応
'   Ver             ：  0.3
'**********************************************
Option Compare Database
Option Explicit

Private Const P_ERROR            As String = "エラー"

' 【予約受付状態区分】
Private Const P_YUKBN_20         As String = "20"          ' 契約

' 【移動区分】
Private Const P_MOVEKBN_00       As String = "00"          ' 無し

' 【契約状態】
Private Const P_AGRE_04          As String = "04"          ' メンテ中
Private Const P_AGRE_09          As String = "09"          ' 完了

Private Const P_SELECT_ERROR     As String = "メンテ中または完了のため、" & vbCrLf & _
                                             "移動元としては選択できません。"

'==============================================================================*
'
'画面最大表示
'
'
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

'==============================================================================*
'
'        MODULE_NAME      :cmd_ACPTNO_Click
'        機能             :cmd_ACPTNO_Clickイベント
'        IN               :
'        OUT              :
'
'        内容             :サブフォームで選択した受注契約番号をメインの画面に設定
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_ACPTNO_Click()

    If Me.Parent.pIs受付作成済 = True Then Exit Sub '2006/06/14 受付作成済みなら変更できん
    
    Call subAcptnoSet(0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subAcptnoSet
'        機能             :サブフォームﾃで選択した内容をメインフォームに反映させる
'        IN               :
'        OUT              :
'
'        内容             :サブフォームで選択した受注契約番号をメインの画面に設定
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subAcptnoSet(ByVal intKbn As Integer)

    Dim objRs      As Recordset
    Dim strACPTNO  As String
    Dim strAGRE    As String

    On Error Resume Next

    Set objRs = Me.Form.RecordsetClone

    objRs.MoveLast
    objRs.MoveFirst

    ' 契約一覧のデータ件数が0件のときは処理を抜ける
    If objRs.RecordCount = 0 Then
        objRs.Close
        Set objRs = Nothing
        Exit Sub
    End If

    ' メイン画面の予約受付状態区分 = 契約 または メイン画面の移動区分 = 無し のとき
    ' 下記の処理は行わない。
    If (Nz(Me.Parent.cmb_Yukbn, "") = P_YUKBN_20 Or Nz(Me.Parent.cmb_MoveKbn, "") = P_MOVEKBN_00) Then Exit Sub

    ' 選択した受注契約番号取得
    strACPTNO = ""
    Select Case intKbn
        Case 0  ' *** 自画面からCallされた場合 ***
            strACPTNO = Nz(Me.FVS500S01_Acptno, "")

        Case 1  ' *** 予約入力画面からCallされた場合 ***
            strACPTNO = Nz(Me.Parent.txt_MOTO_ACPTNO, "")
    End Select

    If Nz(strACPTNO, "") = "" Then Exit Sub

    ' 契約状態チェック
    strAGRE = ""
    strAGRE = Format(Nz(Me.FVS500S01_AGRE, 0), "00")
    If intKbn = 0 And (strAGRE = P_AGRE_04 Or strAGRE = P_AGRE_09) Then
    '  ~~~~~~~~~~↑
    ' 2006/05/19 tajima 自画面からCallされた場合のみチェック
        ' 該当する契約データが｢メンテ中｣または｢完了｣のときは選択エラーとする
        MsgBox P_SELECT_ERROR, vbExclamation, P_ERROR
        Exit Sub
    End If

    ' メイン画面の｢移動元受注契約番号｣欄に設定
    Me.Parent.txt_MOTO_ACPTNO = strACPTNO

    Set objRs = Me.Form.RecordsetClone

    doCmd.Echo False

'    On Error Resume Next
'    objRs.MoveLast
'    objRs.MoveFirst
    Do Until objRs.EOF
        objRs.Edit

        If objRs.Fields("CARG_ACPTNO") = strACPTNO Then
            ' 選択したレコードのダミーの受注契約番号に値をセット
            objRs.Fields("CARG_ACPTNO_Dummy") = strACPTNO
            Me.Form.Bookmark = objRs.Bookmark

            Me.FVS500S01_YCODE = objRs.Fields("CARG_YCODE")  ' ヤードコード
            Me.FVS500S01_NO = objRs.Fields("CARG_NO")        ' コンテナ番号

        Else
            ' ダミーの受注契約番号をNullに更新
            objRs.Fields("CARG_ACPTNO_Dummy") = ""
        End If

        objRs.UPDATE

        objRs.MoveNext
    Loop
    objRs.Close
    Set objRs = Nothing

    doCmd.Echo True
End Sub
'****************************  ended or program ********************************
