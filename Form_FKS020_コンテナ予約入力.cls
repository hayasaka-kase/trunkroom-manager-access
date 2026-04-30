VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS020_コンテナ予約入力"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　コンテナ予約入力
'   プログラムＩＤ　：　FKS020
'   作　成　日　　　：  1999/10/25
'   作　成　者　　　：  αシステム　栗原  洋行
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：  2003/05/21
'   修　正　者　　　：  azegami@eagle-soft.co.jp
'   修　正　内　容　：  コンテナ配置図をAPIで呼び出す
'   Ver             ：  0.1
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：  2004/07/10
'   修　正　者　　　：  N.MIURA
'   修　正　内　容　：  WEB表示機能追加
'   Ver             ：  0.2
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：  2004/11/06
'   修　正　者　　　：  N.MIURA
'   修　正　内　容　：  キャンペーン表示
'   Ver             ：  0.3
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/01/28
'   修　正　者　　　：  N.MIURA
'   修　正　内　容　：  予約受付条件変更
'   Ver             ：  0.4
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/05/06
'   修　正　者　　　：  K.MITSUNO
'   修　正　内　容　：  新予約入力画面追加対応
'   Ver            ：  0.5
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/07/13
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  予約トランの結合条件修正
'   Ver             ：  0.6
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/07/17
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  外部から再表示される前に保存する
'   修　正　内　容　：  画面遷移する前にカレントを保存
'   修　正　内　容　：  タイトルに修正中かどうかを表示
'   Ver             ：  0.7
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/08/03
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  コンテナ予約ＷＫに項目(CARG_ACPTNO)を追加し、コンテナ契約ファイルの受注契約番号をセットするように修正
'   Ver             ：  0.8
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/10/19
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  コンテナの種別ではなく、コンテナの実帖を設定するように修正
'   Ver             ：  0.9
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/11/21
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  1.価格表テーブル変更に伴う項目の表示項目の修正
'                       2.SQL-Serverに直接接続するよう修正
'                       3.上記「2.」の修正に伴い、SQL文の変更
'                       4.[保存]、[取消]ボタンの使用可否の制御追加
'   Ver             ：  0.9
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/12/06
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  1.契約客先名表示箇所の修正
'                         (条件に従い｢紹介待ち｣を表示するように修正)
'   Ver             ：  1.0
'**********************************************
'修正履歴
'   修　正　日　　　：　2006/06/24
'   修　正　者　　　：  tajima
'   修　正　内　容　：  修繕呼出対応
'   Ver             ：  1.1
'**********************************************
'修正履歴
'   修　正　日　　　：　2006/09/24
'   修　正　者　　　：  tajima
'   修　正　内　容　：  修繕対応その２
'                   ：  ①地主情報削除により地主情報関係のＤＢ参照を削除
'                   ：  ②撤去コンテナは対象外
'   Ver             ：  1.2
'**********************************************
'修正履歴
'   修　正　日　　　：　2007/01/12
'   修　正　者　　　：  tajima
'   修　正　内　容　：  備考のフォントを赤字 ※ソース修正無し
'   Ver             ：  1.3
'**********************************************
'修正履歴
'   修　正　日　　　：　2007/02/08
'   修　正　者　　　：  tajima
'   修　正　内　容　：  ヤード制限対応
'   Ver             ：  1.4
'**********************************************
'修正履歴
'   修　正　日　　　：　2007/04/15
'   修　正　者　　　：  tajima
'   修　正　内　容　：  近隣で取置されたコンテナの予約情報表示
'   Ver             ：  1.5
'**********************************************
'修正履歴
'   修　正　日　　　：　2009/06/13
'   修　正　者　　　：  N.Miura
'   修　正　内　容　：  詳細ページ（ホームページ）の拡張子変更
'   Ver             ：  1.6
'**********************************************
'修正履歴
'   修　正　日　　　：　2011/09/14
'   修　正　者　　　：  M.RYU
'   修　正　内　容　：  WEBボタンリンク
'   Ver             ：  1.7
'**********************************************
'修正履歴
'   修　正　日　　　：　2012/07/04
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  ログ出力対応
'   Ver             ：  1.8
'**********************************************
'修正履歴
'   修　正　日　　　：　2013/04/18
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  ﾒﾝﾃ担当者/巡回業者を追加
'   Ver             ：  1.9
'**********************************************
'修正履歴
'   修　正　日　　　：　2013/04/30
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  メンテ/巡回業者の位置を調整
'   Ver             ：  2.0
'**********************************************
'修正履歴
'   修　正　日　　　：　2013/06/03
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  終了メッセージを表示させないように修正
'   Ver             ：  2.1
'**********************************************
'修正履歴
'   修　正　日　　　：　2014/04/07
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  ﾔｰﾄﾞﾏｽﾀのｽﾍﾟｱ区分を表示するように修正
'   Ver             ：  2.2
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2017/09/07
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  ヤード担当を表示
'   Ver             ：  2.3
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2018/03/18
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  写真フォルダーボタンを作成
'   Ver             ：  2.4
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2019/11/20
'   修　正　者　　　：  Y.WADA
'   修　正　内　容　：  トランク管理（部門8）の場合はオーナー情報、問合せボタンを表示
'   Ver             ：  2.5
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2019/12/03
'   修　正　者　　　：  Y.WADA
'   修　正　内　容　：  問合せボタンを廃止
'   Ver             ：  2.6
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2019/12/11
'   修　正　者　　　：  Y.WADA
'   修　正　内　容　：  オーナー情報の取得で部門コードが条件から抜けていたのを修正
'   Ver             ：  2.7
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2021/07/08
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  権限８を追加
'   Ver             ：  2.8
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2023/08/14
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  鍵番号を表示するように修正
'   Ver             ：  2.9
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2025/07/11
'   修　正　者　　　：  N.IMAI
'   修　正　内　容　：  備考をカーソルで選択可能にする
'   Ver             ：  3.0
'**********************************************
'修正履歴
'   修　正　日　　　：　2026/01/26
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  写真フォルダーの参照方法をHの場合のみ変更
'   Ver             ：  3.1
'**********************************************
'修正履歴
'   修　正　日　　　：　2026/04/20
'   修　正　者　　　：  N.IMAI
'   修　正　内　容　：  ヤードマスタのネット予約対象フラグがOFFの場合には、WEBボタンを使用不可へ
'   Ver             ：  3.2
'**********************************************
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FKS020"

Private strMsg               As String           'メッセージ表示用
Private intAns               As Integer          '確認入力用
Private Const strMsg_kaku = "確認"
Private Const strMsg_err = "エラー"

Private Const WORK_TAB01 = "WK_FKS020_TEMP1"    'オーナー情報用  'INSERT 2019/11/20 Y.WADA

Private objDb                As Database
'Dim objRs As Recordset  ' 2005/11/21 DEL T.SUZUKI

Private WRI_ObjWS            As Workspace
Private WRI_objDb            As Database

Private vGotText             As Variant          'フォーカスを取得した時の内容
Private bSYKYMD              As Boolean
Private vStep                As Variant
Private vDATE                As Variant

Private KYDATE               As Variant

Public bOpenFlag_FKG020      As Boolean          'フォームが読込まれたかどうかのフラグ
Public bChangeFlag_FKG020    As Boolean          '入力に変更があったかどうかのフラグ
Public sJYUNO_FKG020         As String           '複製ボタンを押した時に返り値をセットする

Private strCaption           As String
Private lngYCodeDum          As Long             'add 01-05-2000 Araki

' 2005/11/21 ADD T.SUZUKI Start
Private pobjKONT_DB              As Database             ' コンテナDB
Private pstrBumoCD               As String               ' 部門コード格納
Private Const pCStrCANCEL        As String = "9"         ' キャンセル
Private Const pCStrNotIntroduct  As String = "6"         ' 紹介不要

Private Const pCStrWaitIntroduct As String = "紹介待ち"  ' ｢紹介待ち｣
' 2005/11/21 ADD T.SUZUKI End

' 2005/12/19 ADD T.SUZUKI Start
Private Const P_INTROKBN_01      As String = "1"          '取置
Private Const P_INTROKBN_02      As String = "2"          '契約受付

Private Const P_YUKBN_01         As Integer = 1           '自社理由
Private Const P_YUKBN_02         As Integer = 2           '取置
Private Const P_YUKBN_10         As Integer = 10          '受付
' 2005/12/19 ADD T.SUZUKI End

' デバック用変数宣言 Start
' ライブラリ
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private mLngStartTime As Long
' デバック用変数宣言 End

' 2006/09/24 Add tajima
Private pobjKASE_DB              As Database             ' 加瀬DB

''==============================================================================*
''
''        MODULE_NAME      :検索ボタン
''        MODULE_ID        :cmd_search_Click
''        CREATE_DATE      :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Search_Click()
  doCmd.OpenForm "FKS521_利用状況検索", acNormal, , , , acDialog
  
  Call cmb_ヤードコード_AfterUpdate
  
  
End Sub

'DELETE 2019/12/03 Y.WADA Start
''==============================================================================*
''
''        MODULE_NAME      :問合せボタン
''        MODULE_ID        :cmd_toiawase_Click
''        CREATE_DATE      :2019/11/20 Y.WADA
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmd_toiawase_Click()
'    On Error GoTo Exception
'
'
'    '（加瀬システム）賃貸物件問合せを開く
'    Call MSZZ080_SysFormOpen("MFLY017", pstrBumoCD & " " & Nz(cmb_ヤードコード))
'
'Exit Sub
'
'Exception:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'End Sub
'DELETE 2019/12/03 Y.WADA End

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_写真_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_写真_Click()

    Dim strPathName As String

    'パスを読込
    strPathName = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = '" & "FKS020_" & pstrBumoCD & "'"), "")
    
    If Len(strPathName) > 0 And Right$(strPathName, 1) <> "\" Then
        strPathName = strPathName & "\"
    End If
    
    
    If pstrBumoCD = "H" Then
        'Hの場合にヤードコードのあとに名称を追加
        If Dir(strPathName & cmb_ヤードコード & txt_ヤード名, vbDirectory) = "" Then
            MkDir cmb_ヤードコード & txt_ヤード名
        End If
        '対象のフォルダを開く
        Shell "C:\Windows\Explorer.exe " & strPathName & cmb_ヤードコード & txt_ヤード名, vbNormalFocus
     Else
        '選択されているヤードのフォルダーが無ければ作成する。
        If Dir(strPathName & cmb_ヤードコード, vbDirectory) = "" Then
            MkDir strPathName & cmb_ヤードコード
        End If
        '対象のフォルダを開く
        Shell "C:\Windows\Explorer.exe " & strPathName & cmb_ヤードコード, vbNormalFocus
    End If
    
  

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Openイベント
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル2.BackColor = RGB(255, 0, 204)
 
    End Select
    Call MSHS010_M00(PROG_ID, "1")
    doCmd.Maximize
End Sub
'==============================================================================*
'
'        MODULE_NAME      :Form_Loadイベント
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    strCaption = Me.Caption                                                     'INSERT 20050718 K.ISHIZAKA

    doCmd.Maximize
    bOpenFlag_FKG020 = False

    Set objDb = CurrentDb

    'INSERT 2019/11/20 Y.WADA Start
    ' 部門コード取得
    pstrBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    'INSERT 2019/11/20 Y.WADA End

    'フォームの初期化処理
    Call CLR_RTN       ' 画面初期化処理
    Call Etc_Lock_Rtn  ' コントロールのロック制御処理
    Call DELWK_RTN     ' ワークテーブルのデータ削除処理

'DELETE 2019/11/20 Y.WADA Start
'' 2005/11/08 ADD T.SUZUKI Start
'    ' 部門コード取得
'    pstrBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
'' 2005/11/08 ADD T.SUZUKI Start
'DELETE 2019/11/20 Y.WADA End

    bOpenFlag_FKG020 = True

    lngYCodeDum = 0         'add 01-05-2000 Araki
End Sub
'==============================================================================*
'
'        MODULE_NAME      :Form_Activateイベント
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    'フルスクリーン化
    doCmd.Maximize
End Sub
'==============================================================================*
'
'        MODULE_NAME      :Form_Closeイベント
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()

    Dim strNm As String     'INSERT 2019/11/20 Y.WADA

    ' ワークテーブルのデータ削除処理
    Call DELWK_RTN
    
    'DELETE 2019/11/20 Y.WADA Start
    '作業テーブル削除
    strNm = subf_owner.Form.RecordSource
    subf_owner.Form.RecordSource = ""
    Call DropWorkTable_CurrentDb(strNm)
    'DELETE 2019/11/20 Y.WADA End
    
    Call MSHS010_M00(PROG_ID, "2")

    'オブジェクト変数の解放
'    Set objRs = Nothing  ' 2005/11/21 DEL T.SUZUKI
    Set objDb = Nothing
    Set WRI_ObjWS = Nothing
    Set WRI_objDb = Nothing
End Sub
'==============================================================================*
'
'        MODULE_NAME      :Form_Deactivateイベント
'        MODULE_ID        :Form_Deactivate
'        CREATE_DATE      :2005/07/17 K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Deactivate()
    On Error Resume Next
    doCmd.RunCommand acCmdSaveRecord
End Sub
'==============================================================================*
'
'        MODULE_NAME      :Form_Errorイベント
'        MODULE_ID        :Form_Error
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Error(DataErr As Integer, Response As Integer)

'    Debug.Print DataErr
'    Debug.Print Response
    Select Case DataErr
        Case 2169       '「レコードは保存されません。」という時
            DataErr = 0
            Response = 0
    End Select
End Sub
'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDownイベント
'        MODULE_ID        :Form_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next

    ' ▼2006/06/24 Add tajima
    Call CmReprMod.OpenFVS90n(KeyCode, Me.cmb_ヤードコード)
    ' ▲2006/06/24 Add tajima

    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_終了.SetFocus
        Case vbKeyF2
            KeyCode = 0
            Me.cmd_保存.SetFocus
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_ヤードコード_AfterUpdateイベント
'        MODULE_ID        :cmb_ヤードコード_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤードコード_AfterUpdate()

    Dim vCode As Variant

    On Error GoTo err_rtn

' デバック用メッセージ ------------------------------
' スタート時間セット
'Call StartTimer_FS
' デバック用メッセージ ------------------------------

    vCode = Nz(Me.cmb_ヤードコード, "")

    Call Etc_Lock_Rtn  ' 2005/11/21 ADD T.SUZUKI

    If vCode = "" Then
        Exit Sub
    End If

    Me.cmb_ヤードコード = Format(Me.cmb_ヤードコード, "000000")   'ADD 99/12/16

    'ファイルへの存在チェック
    ' *** ヤードマスタ ***
    If Chk_Yard_Mast(1) = False Then
    End If

    ' *** 地主マスタ
' ▼2006/09/24 Del tajima▼
'    If CHK_JARG_FILE(1) = False Then
'    End If
' ▲2006/09/24 Del tajima▲

'--- ADD 2000/06/13 ----------------
'''''  If KYDATE < Date Then
'''''     Exit Sub
'''''  End If
'--- ADD 2000/06/13 ----------------

    Call subCreSubf_owner(True) 'オーナー情報サブフォーム設定                     'INSERT 2019/11/20 Y.WADA

    ' ワークテーブルのデータ削除処理
    Call DELWK_RTN

    vDATE = Now

    Me.txt_date = vDATE

    Call Make_Work

    Me.サブ上.Enabled = True
    Me.サブ下.Enabled = True

    Me.サブ上.Requery
    Me.サブ下.Requery

    Me.サブ上A.Requery  ' 2005/11/21 ADD T.SUZUKI
    Me.サブ下A.Requery  ' 2005/11/21 ADD T.SUZUKI

    If CHK_CNTA_MAST2 = False Then
        Exit Sub
    End If

'2005/05/30
'予約名/電話番号変更時、予約トランに受付番号が入っていたら
'ｺﾝﾃﾅﾏｽﾀへUpdateしない処理追加の為、以下の処理をコメント
    If vStep = 2 Or vStep = 0 Then
        Me.サブ上.SetFocus
'    Me.サブ上.Form!txt_予約客先名.SetFocus
    Else
        Me.サブ下.SetFocus
'    Me.サブ下.Form!txt_予約客先名.SetFocus
    End If

    lngYCodeDum = Me.cmb_ヤードコード 'add 01-05-2000 Araki

    Me.サブ上A.Enabled = True         '2000/07/03 insert n-miua
    Me.サブ下A.Enabled = True         '2000/07/03 insert n-miua

    Me.サブ上A.Requery                '2000/07/03 insert n-miua
    Me.サブ下A.Requery                '2000/07/03 insert n-miua

'    Me.txtBUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")                                             'DELETE 20090613 N.MIURA START 'INSERT 20040710 N.MIURA START
'    If Me.txtBUMOC = "8" Then
'        Me.txtAddress = DLookup("CONT_DISKBACK", "dbo_CONT_MAST", "CONT_KEY = 1") & Val(Me.cmb_ヤードコード) & ".htm"
'    Else
'        Me.txtAddress = DLookup("CONT_DISKBACK", "dbo_CONT_MAST", "CONT_KEY = 1") & Me.cmb_ヤードコード & ".htm"
'    End If                                                                                                           'DELETE 20090613 N.MIURA ENDED 'INSERT 20040710 N.MIURA ENDED
    
    Me.txtBUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")                                             'INSERT 20090613 N.MIURA START
    
    Select Case Me.txtBUMOC
    Case "1": Me.txtAddress = DLookup("CONT_DISKBACK", "dbo_CONT_MAST", "CONT_KEY = 1") & Me.cmb_ヤードコード & ".php"
'    Case "8": Me.txtAddress = DLookup("CONT_DISKBACK", "dbo_CONT_MAST", "CONT_KEY = 1") & Val(Me.cmb_ヤードコード) & ".htm"    'DELETE 2011/09/14 M.RYU
'    Case Else: Me.txtAddress = DLookup("CONT_DISKBACK", "dbo_CONT_MAST", "CONT_KEY = 1") & Me.cmb_ヤードコード & ".htm"        'DELETE 2011/09/14 M.RYU
    Case "8": Me.txtAddress = DLookup("CONT_DISKBACK", "dbo_CONT_MAST", "CONT_KEY = 1") & Val(Me.cmb_ヤードコード) & ".html"    'INSERT 2011/09/14 M.RYU
    Case Else: Me.txtAddress = DLookup("CONT_DISKBACK", "dbo_CONT_MAST", "CONT_KEY = 1") & Me.cmb_ヤードコード & ".html"        'INSERT 2011/09/14 M.RYU
    End Select                                                                                                       'INSERT 20090613 N.MIURA ENDED

    bChangeFlag_FKG020 = False                                                    'INSERT 20050718 K.ISHIZAKA
    Me.Caption = strCaption                                                       'INSERT 20050718 K.ISHIZAKA

' デバック用メッセージ ------------------------------
'Dim lngTime As Long
' 処理時間取得
'lngTime = StopTimer_FS()

' コンテナマスタ削除データログ出力処理
'Call MSZZ003_M00(PROG_ID, 0, "処理時間：" & lngTime & "ﾐﾘ秒")
' デバック用メッセージ ------------------------------

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_ヤードコード_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤードコード_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_ヤードコード, "")

    If vCode = "" Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

    '入力チェック
    Select Case CmbCheck_Num(vCode, 6)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn

        Case 9
            MsgBox "コードを入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
    End Select

    'ファイルへの存在チェック
    ' *** ヤードマスタ
    If Chk_Yard_Mast(0) = False Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

    ' *** 地主マスタ
' ▼2006/09/24 Del tajima▼
'    If CHK_JARG_FILE(0) = False Then
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'    End If
' ▲2006/09/24 Del tajima▲

    ' *** コンテナマスタ
    If CHK_CNTA_MAST(0) = False Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_ヤードコード_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤードコード_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Num(Me.cmb_ヤードコード, KeyAscii, 6)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :受付入力(ｺﾝﾃﾅ未確定⇒ﾔｰﾄﾞ番号のみ渡す)
'        MODULE_ID        :cmd_YUke_Click
'        CREATE_DATE      :2005/05/06 K.MITSUNO
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_YUke_Click()

    Dim vHData As Variant

    If Me.SaveData() = False Then                                               'INSERT 20050718 K.ISHIZAKA
        Me.cmb_ヤードコード.SetFocus                                            'INSERT 20050718 K.ISHIZAKA
        Exit Sub                                                                'INSERT 20050718 K.ISHIZAKA
    End If                                                                      'INSERT 20050718 K.ISHIZAKA
    
    'ﾔｰﾄﾞ番号(6桁)送るデータは23byte
    vHData = Format$(Me.cmb_ヤードコード, "000000")

    '予約入力画面　起動
    doCmd.OpenForm "FVS500", , , , , , vHData
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_WEB表示_Click
'        CREATE_DATE      :2004/07/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_WEB表示_Click()
    'Dim txtSubAddress As TextBox
    Call MSZZ009_M00(Me.cmd_WEB表示, Me.txtsubAddress, Me.txtAddress)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ｷｬﾝﾍﾟｰﾝ
'        MODULE_ID        :cmd_CAMPB_Click
'        CREATE_DATE      :2004/11/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_CAMPB_Click()

    If Me.cmb_ヤードコード = "" Or IsNull(Me.cmb_ヤードコード) = True Then
        Exit Sub
    End If
    doCmd.OpenForm "FKS020S50", , , , , acDialog
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_近隣ヤード_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_近隣ヤード_Click()

'ここから                                                                       'DELETE 20050718 K.ISHIZAKA
'    '$$$ start $$$ 近隣ヤード選択機能追加 (01-05-2000 Araki)
'    If Me.cmb_ヤードコード = "" Or IsNull(Me.cmb_ヤードコード) = True Then
'        Exit Sub
'    End If
'
'    '保存処理
'    If Me.flag_ch = True Then                                                  'DELETE 20050718 K.ISHIZAKA
'        intAns = MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, strMsg_kaku)
'        If intAns = vbYes Then
'            If WRI_RTN = False Then
'                Me.cmb_ヤードコード.SetFocus
'                Exit Sub
'            End If
'            Me.flag_ch = False
'        End If
'    End If
'ここまで                                                                       'DELETE 20050718 K.ISHIZAKA
    If SaveData() = False Then                                                  'INSERT 20050718 K.ISHIZAKA
        Me.cmb_ヤードコード.SetFocus                                            'INSERT 20050718 K.ISHIZAKA
        Exit Sub                                                                'INSERT 20050718 K.ISHIZAKA
    End If                                                                      'INSERT 20050718 K.ISHIZAKA

    '近隣ヤード選択画面　起動
    doCmd.OpenForm "FKS035_近隣ヤード", , , , , acDialog

    If Format(lngYCodeDum, "000000") <> Format(Me.cmb_ヤードコード, "000000") Then
        Call cmb_ヤードコード_AfterUpdate                                       'INSERT 20050718 K.ISHIZAKA
'ここから                                                                       'DELETE 20050718 K.ISHIZAKA
'        'ファイルへの存在チェック
'        If Chk_Yard_Mast(1) = False Then
'        End If
'        If CHK_JARG_FILE(1) = False Then
'        End If
'
'        'ワーク作成
'        Call DELWK_RTN
'
'        Me.flag_ch = False
'        lngYCodeDum = Me.cmb_ヤードコード
'        vDATE = Now
'        Me.txt_date = vDATE
'
'        Call Make_Work
'
'        Me.サブ上.Enabled = True
'        Me.サブ下.Enabled = True
'        Me.サブ上.Requery
'        Me.サブ下.Requery
'
'        '2005/05/29 MITSUNO
'        '予約名/電話番号変更時、予約トランに受付番号が入っていたら
'        'ｺﾝﾃﾅﾏｽﾀへUpdateしない処理追加の為、以下の処理をコメント
'        If vStep = 2 Or vStep = 0 Then
'            Me.サブ上.SetFocus
''            Me.サブ上.Form!txt_予約客先名.SetFocus
'        Else
'            Me.サブ下.SetFocus
''            Me.サブ下.Form!txt_予約客先名.SetFocus
'        End If
'ここまで                                                                       'DELETE 20050718 K.ISHIZAKA
    End If
    '$$$ end $$$
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_配置_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_配置_Click()

    'Dim vRet1        As Variant  2003/05/21 azegami@eagle-soft.co.jp
    Dim vRET2        As Variant
    Dim strExcel     As String
    Dim RetVal       As Variant

    On Error GoTo err_rtn

    Call MSHS010_M00(PROG_ID, "3", cmb_ヤードコード)

    'Start 2003/05/21 azegami@eagle-soft.co.jp
    'If Me.cmb_ヤードコード = "" Or IsNull(Me.cmb_ヤードコード) = True Then
    '   Exit Sub
    'End If

    'Excel起動パス情報
        '[コントロールマスタ].Excel情報パス 存在チェック
        
        'vRet1 = Nz(DlookUp("EXCEL_PATH", "ＥＸＣＥＬ起動パス", ""), "")
    
        'If (vRet1 = "") Or (vRet1 = " ") Then
        '    MsgBox "EXCEL起動パスが登録されていません。", vbOK, strMsg_kaku
        '    Exit Sub
        'End If
    'End 2003/05/21 azegami@eagle-soft.co.jp

    'Excel情報取得
    '[コントロールマスタ].Excel情報パス 存在チェック
    vRET2 = Nz(DLookup("CONT_EXCEL_PATH", "dbo_CONT_MAST", ""), "")

    If (vRET2 = "") Or (vRET2 = " ") Then
        MsgBox "EXCEL情報パスが登録されていません。", vbOK, strMsg_kaku
        Exit Sub
    End If

    If Right(vRET2, 1) = "\" Then
        vRET2 = vRET2
    Else
        vRET2 = vRET2 & "\"
    End If

    '[コントロールマスタ].EXCEL情報パス のディレクトリにExcelが存在するか検索する。
    'EXCELのファイル名は "(ヤードコード).XLS"とする。
    strExcel = Dir(vRET2 & Format(Me.cmb_ヤードコード, "000000") & "*.XLS")

    If strExcel <> "" Then
        'RetVal = Shell(vRet1 & " " & vRet2 & strEXCEL, 1) 2003/05/21 azegami@eagle-soft.co.jp
        Call Y_Exec(Me.hwnd, vRET2 & strExcel, "", "", SW_SHOWNORMAL) '2003/05/21 azegami@eagle-soft.co.jp
        Exit Sub
    Else
        MsgBox "ｺﾝﾃﾅ配置EXCELが存在しません。", vbOK, strMsg_kaku
        Exit Sub
    End If

err_rtn:
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Haichi_End_Enterイベント
'        MODULE_ID        :txt_Haichi_End_Enter
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Haichi_End_Enter()
    ' ヤードコードにフォーカス移動
    Me.cmb_ヤードコード.SetFocus
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CmBCheck_Char
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer

    '文字項目のチェックを一括でする。
    
    '返り値は、
    '
    '２：桁数が溢れている時
    '９：それ以外のエラー
    '０：すべてのチェックがＯＫの時

    On Error GoTo err_rtn

    CmBCheck_Char = 9

    If CmLenb(oText) > ilen Then
        CmBCheck_Char = 2
        GoTo Exit_rtn
    End If

    CmBCheck_Char = 0

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    CmBCheck_Char = 9
    GoTo Exit_rtn
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CmBCheck_Date
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Date(ByVal oText As Variant, Optional ByVal ilen As Variant) As Integer

    '日付項目のチェックを一括でする
    
    '引数
    'oText 対象項目
    'iLen  桁数(省略時は、"yyyy/mm/dd"として10バイト)
    
    '返り値は、
    '２：桁数が溢れている時
    '３：日付と判断できない時
    '９：それ以外のエラー
    '０：すべてのチェックがＯＫの時

    Dim sYear As String
    Dim intLen As Integer

    On Error GoTo err_rtn

    CmBCheck_Date = 2

    If IsMissing(ilen) = False Then
        intLen = Nz(CmNumeric(ilen), 0)
    Else
        intLen = 10
    End If

    oText = Nz(oText, "")

    '桁数チェック
    If CmLenb(oText) > intLen Then
        Exit Function
    End If

    CmBCheck_Date = 3

    If Nz(oText, "") = "" Then
        Exit Function
    End If

    '論理チェック
    If IsDate(oText) = False Then
        Exit Function
    End If

    sYear = Format(oText, "yyyy")

    If CmLenb(sYear) <> 4 Then
        Exit Function
    End If

    CmBCheck_Date = 0

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    CmBCheck_Date = 9
    GoTo Exit_rtn
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CmBCheck_Num
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmbCheck_Num(ByVal oText As Variant, ByVal ilen As Integer, Optional ByVal isLEN As Variant) As Integer

    '数値項目のチェックを一括でする。
    '返り値は、
    '１：数値と判断できない時
    '２：桁数が溢れている時
    '９：それ以外のエラー
    '０：すべてのチェックがＯＫの時
    
    '引数 oText ←判定する値
    '     iLen  ←整数桁
    '     isLen ←小数桁(省略された時は、小数点桁数無し)

    Dim vINT    As Variant
    Dim vsINT   As Variant
    Dim i       As Integer
    
    Dim dTEXT   As Double
    Dim dFROM   As Double
    Dim dTO     As Double
    Dim vRet    As Variant
    
    On Error GoTo err_rtn
    CmbCheck_Num = 9

    If Nz(CmNumeric(oText), "") = "" Then       '数値と判断できないのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If

    If CmChkCher(CStr(Nz(oText)), "A") = False Then '全角文字が入力されているのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If

    '小数点の値が設定されている時のチェック
    vRet = Nz(InStr(Abs(oText), "."), 0)

    If IsMissing(isLEN) = False Then
        If vRet <> 0 Then
            If isLEN < CmLenb(Abs(oText)) - vRet Then    '小数点桁数が設定桁数を超えているエラー
                CmbCheck_Num = 2
                GoTo Exit_rtn
            End If
        End If

    Else
        If CmLenb(Abs(CmNumeric(oText))) > ilen Then           '文字数超過のためエラー
            CmbCheck_Num = 2
            GoTo Exit_rtn
        End If

        If vRet <> 0 Then                         '小数点の入力があるのでエラー
            CmbCheck_Num = 2
            GoTo Exit_rtn
        End If
    End If

    '桁チェック
    i = 0
    For i = 1 To ilen
        vINT = vINT & "9"
    Next

    If IsMissing(isLEN) = False Then
        vsINT = "."
        i = 0
        For i = 1 To isLEN
            vsINT = vsINT & "9"
        Next
    End If

    vINT = vINT & vsINT
    vINT = CmNumeric(vINT)

    dTEXT = CmNumeric(oText)
    dFROM = vINT * -1
    dTO = vINT

    If dFROM > dTEXT Or dTO < dTEXT Then        '論理範囲超過のためエラー
        CmbCheck_Num = 2
        GoTo Exit_rtn
    End If

    CmbCheck_Num = 0

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    CmbCheck_Num = 9
    GoTo Exit_rtn
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CMCheck_Date
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CMCheck_Date(ByVal vDATE As Variant)

    Dim sYear As String
    'Check_Dateでは'198/03といった形が変換されるための補助検査

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
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CmParEmpty
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmParEmpty(ByVal oEmpty As Variant, Optional ByVal oEmptyFull As Variant) As Variant

    Dim vText As Variant

    If IsMissing(oEmptyFull) = False Then
        vText = oEmptyFull
    Else
        vText = 0
    End If

    If IsEmpty(oEmpty) = True Then
        CmParEmpty = vText
    Else
        CmParEmpty = oEmpty
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナマスタ、契約ファイル、顧客マスタ、予約受付トラン、
'                          予約ご紹介トランの情報をワークに落す
'        MODULE_ID        :Make_Work
'        CREATE_DATE      :2005/12/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub Make_Work()

    Dim strSQL  As String
    Dim vCode   As Variant

    Dim objRs     As Recordset  ' メインデータレコードセット
    Dim objRsWk   As Recordset  ' ワークテーブルレコードセット
    Dim intCount  As Integer    ' カウント変数
    Dim objKaseRs As Recordset  ' 加瀬DBのレコードセット 2006/09/24 Add tajima
    Dim strReprText As String   ' 修繕文言               2006/09/24 Add tajima
    Dim strKotoText As String   ' 顧客問合文言           2006/09/24 Add tajima
    
    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_ヤードコード, "")

    If vCode = "" Then Exit Sub

    ' 加瀬DB接続
    If fncConnectDB = False Then Exit Sub
    If fncConnectKaseDB = False Then Exit Sub ' 2006/09/24 tajima add

    ' コンテナマスタより対象データ抽出
    strSQL = " SELECT   CNTA_MAST.CNTA_CODE, " & Chr(13)        ' ヤードコード"
    strSQL = strSQL & " CNTA_MAST.CNTA_NO, " & Chr(13)          ' コンテナ番号
    strSQL = strSQL & " CNTA_MAST.CNTA_STEP, " & Chr(13)        ' 段区分
    strSQL = strSQL & " 0 AS CNTA_CONTNA, " & Chr(13)           ' コンテナ状態
    strSQL = strSQL & " CNTA_MAST.CNTA_YCODE, " & Chr(13)       ' 予約客先ｺｰﾄﾞ
    strSQL = strSQL & " CNTA_MAST.CNTA_YNAME, " & Chr(13)       ' コンテナマスタ.予約客先名
    strSQL = strSQL & " CNTA_MAST.CNTA_YTEL, " & Chr(13)        ' コンテナマスタ.予約客先電話番号
    strSQL = strSQL & " YOYAKU.YOUKT_NAME, " & Chr(13)          ' 予約受付トラン.予約客先名
    strSQL = strSQL & " YOYAKU.YOYAKU_YTEL, " & Chr(13)         ' 予約受付トラン.予約客先電話番号
    strSQL = strSQL & " CNTA_MAST.CNTA_SIZE, " & Chr(13)        ' 実帖
    strSQL = strSQL & " ISNULL(CARG_FILE.CARG_UCODE,YOUKT_UCODE) CARG_UCODE, " & Chr(13)       ' 顧客コード
    strSQL = strSQL & " CARG_FILE.CARG_KYDATE, " & Chr(13)      ' 解約日付
    strSQL = strSQL & " CARG_FILE.CARG_EDDATE, " & Chr(13)      ' 契約満了日
    strSQL = strSQL & " CARG_FILE.CARG_DOCU1, " & Chr(13)       ' 新契約書切替区分
    strSQL = strSQL & " CARG_FILE.CARG_DOCU2, " & Chr(13)       ' 共益費区分
    strSQL = strSQL & " USER_MAST.USER_NAME, " & Chr(13)        ' 顧客名
    strSQL = strSQL & " CARG_FILE.CARG_RENTKG, " & Chr(13)      ' 月額料金
    strSQL = strSQL & " CARG_FILE.CARG_SECUKG, " & Chr(13)      ' 預り保証金
    strSQL = strSQL & " CARG_FILE.CARG_CONTNA, " & Chr(13)      ' コンテナ状態
    strSQL = strSQL & " CARG_FILE.CARG_ACPTNO, " & Chr(13)      ' 受注契約番号
    strSQL = strSQL & " CNTA_MAST.CNTA_USE, " & Chr(13)         ' 利用可否
    strSQL = strSQL & " CNTA_UKNO, " & Chr(13)                  ' 紹介待ちにした予約番号
    strSQL = strSQL & " YOYAKU.YOUKT_UKNO " & Chr(13)           ' 予約受付トラン.予約受付番号
    strSQL = strSQL & " ,CNTA_REPR_KBN " & Chr(13)              ' 修理中可否 2006/09/24 tajima add
    
    strSQL = strSQL & "  ,CASE WHEN RCPT_CARG_ACPTNO  IS NOT NULL THEN NAME_601.NAME_NAME ELSE NAME_602.NAME_NAME END KEY_TYPE " & Chr(13)
    strSQL = strSQL & "  ,CASE " & Chr(13)
    strSQL = strSQL & "          WHEN RCPT_CARG_ACPTNO IS NOT NULL AND RCPT_KAINO IS NULL  THEN RCPT_KAGI_NO " & Chr(13)
     strSQL = strSQL & "         WHEN RCPT_CARG_ACPTNO IS NOT NULL   THEN RCPT_KAINO " & Chr(13)
     strSQL = strSQL & "         WHEN CNTA_KEY_TYPE IN (10,15,16,17)  THEN (  SELECT AKKNT.AKKNT_KAINO " & Chr(13)
     strSQL = strSQL & "                                                                 FROM ( " & Chr(13)
     strSQL = strSQL & "                                                                        SELECT AKKNT_YARD, " & Chr(13)
     strSQL = strSQL & "                                                                               AKKNT_CTNO, " & Chr(13)
     strSQL = strSQL & "                                                                               AKKNT_KTYPE, " & Chr(13)
     strSQL = strSQL & "                                                                               AKKNT_HATUD, " & Chr(13)
     strSQL = strSQL & "                                                                               ROW_NUMBER() OVER(PARTITION BY AKKNT_YARD, AKKNT_CTNO, AKKNT_KTYPE, AKKNT_STYPE ORDER BY AKKNT_HATUD DESC) AS SEQ, " & Chr(13)
      strSQL = strSQL & "                                                                              AKKNT_KAINO " & Chr(13)
     strSQL = strSQL & "                                                                        FROM AKKN_TRAN " & Chr(13)
     strSQL = strSQL & "                                                                        WHERE AKKNT_YARD = CNTA_CODE " & Chr(13)
     strSQL = strSQL & "                                                                        AND   AKKNT_CTNO = CNTA_NO " & Chr(13)
     strSQL = strSQL & "                                                                        AND   AKKNT_KTYPE = CNTA_KEY_TYPE " & Chr(13)
     strSQL = strSQL & "                                                                        AND   AKKNT_STYPE = '4' " & Chr(13)
     strSQL = strSQL & "                                                                     ) AKKNT " & Chr(13)
     strSQL = strSQL & "                                                                 WHERE AKKNT.SEQ = 1 " & Chr(13)
     strSQL = strSQL & "                                                               )" & Chr(13)
    strSQL = strSQL & "     Else CNTA_KEY_NO  " & Chr(13)
    strSQL = strSQL & "     END KAGI_NO " & Chr(13)
    
    strSQL = strSQL & " ,ISNULL(YARD_NETUSE_KBN, 0) AS YARD_NETUSE_KBN " & Chr(13)      'INSERT 2026/04/20 N.IMAI
  
    ' ※ｺﾝﾃﾅﾏｽﾀ・ｺﾝﾃﾅ契約ﾌｧｲﾙ・顧客ﾏｽﾀ
    strSQL = strSQL & " FROM CNTA_MAST " & Chr(13)              ' *** ｺﾝﾃﾅﾏｽﾀ ***
    strSQL = strSQL & " INNER JOIN YARD_MAST ON YARD_CODE = CNTA_CODE " & Chr(13)  'INSERT 2026/04/20 N.IMAI
    strSQL = strSQL & " LEFT OUTER JOIN CARG_FILE ON ( CNTA_MAST.CNTA_CODE  = CARG_FILE.CARG_YCODE " & Chr(13)
    strSQL = strSQL & "                           AND  CNTA_MAST.CNTA_NO    = CARG_FILE.CARG_NO " & Chr(13)
    strSQL = strSQL & "                           AND  0                    = CARG_FILE.CARG_CONTNA) " & Chr(13)
    strSQL = strSQL & " LEFT OUTER JOIN RCPT_TRAN ON ( RCPT_CARG_ACPTNO  = CARG_FILE.CARG_ACPTNO ) " & Chr(13)
    strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST NAME_601 ON  NAME_601.NAME_ID = '060' AND NAME_601.NAME_CODE = RCPT_KAGIICD "
    strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST NAME_602 ON  NAME_602.NAME_ID = '060' AND NAME_602.NAME_CODE = CNTA_KEY_TYPE "
    strSQL = strSQL & " LEFT OUTER JOIN USER_MAST ON ( CARG_FILE.CARG_UCODE = USER_MAST.USER_CODE) " & Chr(13)
    ' ｺﾝﾃﾅﾏｽﾀ - 予約受付ﾄﾗﾝ(予約ご紹介ﾄﾗﾝ)
    strSQL = strSQL & " LEFT OUTER JOIN " & Chr(13)
    ' ----------------------------------------------------------------------------------
    ' 予約受付ﾄﾗﾝ - 予約ご紹介ﾄﾗﾝのﾃﾞｰﾀ取得
    strSQL = strSQL & " (SELECT YOUKT_UKNO, " & Chr(13)
    strSQL = strSQL & "         YOUKT_UCODE, " & Chr(13)
    strSQL = strSQL & "         YOUKT_NAME, " & Chr(13)
    strSQL = strSQL & "         ISNULL(YOUKT_TEL, " & Chr(13)
    strSQL = strSQL & "         ISNULL(YOUKT_KEITAI, " & Chr(13)
    strSQL = strSQL & "         ISNULL(YOUKT_FAX,''))) AS YOYAKU_YTEL, " & Chr(13)
    strSQL = strSQL & "         YOUKT_YUKBN, " & Chr(13)        ' 予約受付状態区分
    strSQL = strSQL & "         INTRT_YCODE, " & Chr(13)        ' 予約ご紹介トラン.ヤードコード
    strSQL = strSQL & "         INTRT_NO  " & Chr(13)           ' 予約ご紹介トラン.コンテナ番号
    strSQL = strSQL & "    FROM INTR_TRAN, " & Chr(13)          ' *** 予約ご紹介トラン
    strSQL = strSQL & "         YOUK_TRAN " & Chr(13)           ' *** 予約受付トラン
    strSQL = strSQL & "   WHERE INTRT_UKNO  = YOUKT_UKNO " & Chr(13)
    'strSQL = strSQL & "     AND INTRT_YCODE = YOUKT_YCODE " & Chr(13)  '2007/04/15 del tajima
    strSQL = strSQL & "     AND (YOUKT_YUKBN = " & P_YUKBN_01 & Chr(13)                          ' 自社理由
    strSQL = strSQL & "      OR  YOUKT_YUKBN = " & P_YUKBN_02 & Chr(13)                          ' 取置
    strSQL = strSQL & "      OR  YOUKT_YUKBN = " & P_YUKBN_10 & ") " & Chr(13)                   ' 受付
    strSQL = strSQL & "     AND (INTRT_INTROKBN = '" & P_INTROKBN_01 & "' " & Chr(13)            ' 取置
    strSQL = strSQL & "      OR  INTRT_INTROKBN = '" & P_INTROKBN_02 & "') ) YOYAKU " & Chr(13)  ' 受付
    ' ----------------------------------------------------------------------------------
    strSQL = strSQL & " ON ( CNTA_MAST.CNTA_CODE  = YOYAKU.INTRT_YCODE " & Chr(13)
    strSQL = strSQL & " AND  CNTA_MAST.CNTA_NO    = YOYAKU.INTRT_NO) " & Chr(13)
    '▼2006/09/24 tajima Rep▼
    '↓以下の条件に変更 strSql = strSql & " WHERE CNTA_MAST.CNTA_SYU <> 0 " & Chr(13) ' コンテナマスタ.コンテナ種別
    strSQL = strSQL & " WHERE CNTA_MAST.CNTA_USE <> 9 " & Chr(13) '撤去コンテナは表示対象外
    '▲2006/09/24 tajima Rep▲
    strSQL = strSQL & "   AND CNTA_MAST.CNTA_CODE = " & vCode & Chr(13)  ' コンテナマスタ.ヤードコード
    strSQL = strSQL & ";"
    Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

    If objRs.EOF = False Then

        cmd_WEB表示.Enabled = (objRs.Fields("YARD_NETUSE_KBN") <> 0)            'INSERT 2026/04/20 N.IMAI
        
        ' ワークテーブルオープン
        strSQL = " SELECT * FROM FKS020_コンテナ予約ＷＫ "
        Set objRsWk = objDb.OpenRecordset(strSQL, dbOpenDynaset)

        ' ワークテーブルへデータ登録
        With objRsWk
            Do Until objRs.EOF
                .AddNew

                .Fields("CNTA_DATE") = CDate(Format$(vDATE, "yyyy/mm/dd HH:NN:SS"))  ' 実行日時
                .Fields("CNTA_YARD") = objRs.Fields("CNTA_CODE")                     ' コンテナマスタ.ヤードコード
                .Fields("CNTA_NO") = objRs.Fields("CNTA_NO")                         ' コンテナマスタ.コンテナ番号
                .Fields("CNTA_STEP") = objRs.Fields("CNTA_STEP")                     ' コンテナマスタ.段区分
                .Fields("CNTA_CONTNA") = Nz(objRs.Fields("CNTA_CONTNA"), 0)          ' コンテナマスタ.コンテナ状態
                .Fields("CNTA_YCODE") = objRs.Fields("CNTA_YCODE")                   ' コンテナマスタ.予約客先ｺｰﾄﾞ

                '▼ 2006/09/24 tajima rep ▼
'                If Nz(objRs.Fields("YOUKT_NAME"), "") <> "" Then
'                    ' 該当する予約受付ﾄﾗﾝが存在する場合
'                    ' 予約客先名
'                    .Fields("CNTA_YNAME") = objRs.Fields("YOUKT_NAME")               ' 予約受付トラン.予約客先名
'                    .Fields("CNTA_YTEL") = objRs.Fields("YOYAKU_YTEL")               ' 連絡先(電話番号→携帯電話→FAXでセット)
'                Else
'                    ' 該当する予約受付ﾄﾗﾝが存在しない場合
'                    '
'                    .Fields("CNTA_YNAME") = objRs.Fields("CNTA_YNAME")               ' コンテナマスタ.予約客先名
'                    .Fields("CNTA_YTEL") = objRs.Fields("CNTA_YTEL")                 ' コンテナマスタ.予約客先電話番号
'                End If
'               ★以上は消して構わぬ
                ' 予約が入っていたならば連絡先には予約客先名、予約客先がない且つ契約があれば、その契約顧客の情報を設定する
                If Nz(objRs.Fields("YOUKT_NAME"), "") <> "" Then
                    .Fields("CNTA_YTEL") = objRs.Fields("YOUKT_NAME")
                ElseIf Nz(objRs.Fields("CARG_UCODE"), 0) <> 0 Then
                    .Fields("CNTA_YTEL") = CmReprMod.GetUserLatestKotoText(pobjKASE_DB, pstrBumoCD, _
                                                                           Format$(objRs.Fields("CARG_UCODE"), "000000"))
                End If
                ' 予約客先名は常に部屋修繕を表示
                .Fields("CNTA_YNAME") = CmReprMod.GetRoomLatestReprText(pobjKASE_DB, pstrBumoCD, _
                                                                        Format$(objRs.Fields("CNTA_CODE"), "000000"), _
                                                                        Format$(objRs.Fields("CNTA_NO"), "000000"))
                '▲ 2006/09/24 tajima rep ▲

                .Fields("CNTA_RENTKG") = Nz(objRs.Fields("CARG_RENTKG"), 0)          ' コンテナ契約ファイル.月額料金
                .Fields("CNTA_SECUKG") = Nz(objRs.Fields("CARG_SECUKG"), 0)          ' コンテナ契約ファイル.預り保証金
                .Fields("CNTA_SIZE") = objRs.Fields("CNTA_SIZE")                     ' コンテナマスタ.実帖
                .Fields("CNTA_KCODE") = Nz(objRs.Fields("CARG_UCODE"), 0)            ' コンテナ契約ファイル.顧客コード
                .Fields("CNTA_KYDATE") = objRs.Fields("CARG_KYDATE")                 ' コンテナ契約ファイル.解約日付
                .Fields("CNTA_EDDATE") = objRs.Fields("CARG_EDDATE")                 ' コンテナ契約ファイル.契約満了日
                .Fields("CNTA_DOCU1") = Nz(objRs.Fields("CARG_DOCU1"), 0)            ' コンテナ契約ファイル.新契約書切替区分
                .Fields("CNTA_DOCU2") = Nz(objRs.Fields("CARG_DOCU2"), 0)            ' コンテナ契約ファイル.共益費区分
                .Fields("CNTA_NAME") = objRs.Fields("USER_NAME")                     ' 顧客マスタ.顧客名
                .Fields("CNTA_CONTNA") = Nz(objRs.Fields("CARG_CONTNA"), 0)          ' コンテナ契約ファイル.ｺﾝﾃﾅ状態
                .Fields("CARG_ACPTNO") = objRs.Fields("CARG_ACPTNO")                 ' コンテナ契約ファイル.受注契約番号
                .Fields("YOUKT_UKNO") = objRs.Fields("YOUKT_UKNO")                   ' 予約受付トラン.予約受付番号

                ' ｢紹介待ち｣か否かの判断
                If Nz(objRs.Fields("CARG_UCODE"), "") = "" Then                      ' 2006/01/19 ADD T.SUZUKI
                    If Nz(objRs.Fields("CNTA_UKNO"), "") <> "" Then
                        ' コンテナマスタ.紹介待ちにした予約番号 != NULL のとき
                        .Fields("CNTA_NAME") = pCStrWaitIntroduct
                    End If
                End If                                                               ' 2006/01/19 ADD T.SUZUKI
                .Fields("CNTA_REPR_KBN") = objRs.Fields("CNTA_REPR_KBN")             ' 修理中可否 2006/09/24 tajima add
                
                .Fields("KAGI_NO") = objRs.Fields("KAGI_NO")
                .Fields("KEY_TYPE") = objRs.Fields("KEY_TYPE")
                
                
                
                .UPDATE

                objRs.MoveNext
 
            Loop
        End With

        objRsWk.Close
        Set objRsWk = Nothing
    End If
    objRs.Close
    Set objRs = Nothing

    ' 加瀬DB切断
    If fncDisConnectDB = False Then Exit Sub

Exit_rtn:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not objRsWk Is Nothing Then objRsWk.Close: Set objRsWk = Nothing
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CLR_RTN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CLR_RTN()

    Dim objControl As Object
    Dim blnVisible  As Boolean  'INSERT 2019/11/20 Y.WADA
    Dim adoDbConnection As Object
    Dim rsData          As Object
    Dim isException     As Boolean
    Dim strSQL          As String
    Dim userName        As String
    
    '画面上のテキストボックスをクリアする

    On Error Resume Next  'エラーがあっても処理続行

    For Each objControl In Me.Controls

        Select Case objControl.ControlType
            Case acTextBox
                If Nz(objControl.OldValue, "") = "" Then
                    objControl.VALUE = Null
                End If

            Case acComboBox
                If Nz(objControl.OldValue, "") = "" Then
                    objControl.VALUE = ""
                End If

            Case Else
        End Select
    Next

    Me.cmb_ヤードコード = ""
    Me.txt_ヤード名 = ""
    Me.txt_ヤード住所 = ""
    Me.txt_付近メモ = ""                                                          'INSERT 20041106 N.MIURA
    Me.txt_備考 = ""                                                              'INSERT 20041106 N.MIURA
    Me.txt_営業開始日 = ""                                                        ' 2006/09/24 tajima 地主系から変更
    Me.txt_営業終了日 = ""                                                        ' 2006/09/24 tajima 地主系から変更
    Me.txt_解約予定日 = ""                                                        ' 2006/09/24 tajima 地主系から変更
    Me.txt_メンテ担当者 = ""                                                      ' 20130418 M.HONDA INS
    Me.txt_巡回担当者 = ""                                                        ' 20130418 M.HONDA INS
    Me.txt_スペア = ""                                                            ' 2014/04/07 M.HONDA INS
    Me.txt_リーダ = ""                                                            ' 2017/09/07 M.HONDA INS
    

'    Me.cmd_CAMPB.Enabled = False                                                  'INSERT 20041106 N.MIURA  ' 2005/11/21 DEL T.SUZUKI

    '$$$ 近隣ヤード選択機能追加による追加 (01-05-2000 Araki)
    Me.flag_ch = False

    ' ワークテーブルのデータ削除処理
    Call DELWK_RTN
    
     'サブフォームの初期化
    Me.サブ上.Requery
    Me.サブ下.Requery

    Me.サブ上A.Requery                                                             ' 2005/11/21 ADD T.SUZUKI
    Me.サブ下A.Requery                                                             ' 2005/11/21 ADD T.SUZUKI
    
    '20210708 M.HONDA INS
    'ユーザID取得
    userName = ""
    userName = MSZZ000.LsGetUserName

   'ADO接続Object生成
    Set adoDbConnection = MSZZ025.ADODB_Connection()

    ' コンテナマスタ、価格表テーブルから取得
    strSQL = "SELECT ONLET_ACCEI FROM ONLE_TABL WHERE ONLET_PROGB = 'FKS020'"
    strSQL = strSQL & "AND ONLET_USERB = '" & userName & "'"

    Set rsData = MSZZ025.ADODB_Recordset(strSQL, adoDbConnection)
    ' 対象確認
    If rsData.EOF = False Then
        If "8" = Nz(rsData.Fields("ONLET_ACCEI"), "") Then
            Me.txt_ACC = Nz(rsData.Fields("ONLET_ACCEI"), "")
            Form_FKS020_コンテナ予約入力サブ上.cmd_詳細.Enabled = False
            Form_FKS020_コンテナ予約入力サブ下.cmd_詳細.Enabled = False
            Form_FKS020_コンテナ予約入力サブ上.cmd_Uke.Enabled = False
            Form_FKS020_コンテナ予約入力サブ下.cmd_Uke.Enabled = False
            Me.cmd_YUke.Enabled = False
            Me.cmd_紹介制限.Enabled = False
        End If
    End If
    '20210708 M.HONDA INS
    'INSERT 2019/11/20 Y.WADA Start
    '表示制御
    blnVisible = (pstrBumoCD = "8")     'トランク管理（部門8）の場合は表示
    If "8" = Nz(rsData.Fields("ONLET_ACCEI"), "") Then
        blnVisible = False
    
    Else
        blnVisible = True
    End If
'    cmd_toiawase.Visible = blnVisible   '問合せボタン  'DELETE 2019/12/03 Y.WADA
    lbl_owner.Visible = blnVisible      'オーナー情報ラベル
    subf_owner.Visible = blnVisible     'オーナー情報
    
    Call subCreSubf_owner(True) 'オーナー情報サブフォーム設定
    'INSERT 2019/11/20 Y.WADA End

    '変更フラグの初期化
    bChangeFlag_FKG020 = False
    Me.Caption = strCaption                                                       'INSERT 20050718 K.ISHIZAKA

    Me.cmb_ヤードコード.SetFocus
End Sub

'==============================================================================*
'
'        MODULE_NAME      :編集中に設定
'        MODULE_ID        :SetChangeFlag
'        CREATE_DATE      :2005/07/18 K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetChangeFlag()
    If bChangeFlag_FKG020 = False Then
        Me.Caption = strCaption & "(修正中)"
        bChangeFlag_FKG020 = True
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CmLockControl
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CmLockControl(ByVal ContNM As Object, ByVal iMode As Integer)

  '引数 ０  使用可
  '     １  使用不可
  Dim objControl As Object
  
  Set objControl = ContNM
  
  If iMode = 0 Then
    objControl.Enabled = True
    objControl.Locked = False
  Else
    objControl.Enabled = False
    objControl.Locked = True
  End If
  
End Sub
'==============================================================================*
'
'        MODULE_NAME      :コンテナマスタ問い合わせ
'        MODULE_ID        :CHK_CNTA_MAST
'        Parameter   IN   :iMode
'                                 0 = 存在チェック
'                                 1 = データ参照
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_CNTA_MAST(ByVal iMode As Integer) As Boolean
    
    Dim stWHERE As String
    Dim vCode   As Variant
    Dim vRet    As Variant
    Dim objRs   As Recordset  ' 205/11/18 ADD T.SUZUKI

    On Error GoTo err_rtn

    CHK_CNTA_MAST = False

    vCode = Nz(Me.cmb_ヤードコード, "")

    If vCode = "" Then
        CHK_CNTA_MAST = True
        GoTo Exit_rtn
    End If

    doCmd.Hourglass True
    doCmd.SetWarnings False

    stWHERE = ""
    stWHERE = " CNTA_CODE =" & Val(vCode)
    vRet = Nz(DLookup("CNTA_NO", "dbo_CNTA_MAST", stWHERE), "")
    
    If vRet = "" Then
        MsgBox "コンテナ登録がされておりません", vbInformation, strMsg_kaku
        GoTo Exit_rtn
    End If

    CHK_CNTA_MAST = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set objRs = Nothing
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :WRI_RTN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function WRI_RTN() As Boolean

    On Error GoTo err_rtn

    WRI_RTN = False

    doCmd.Hourglass True
    doCmd.SetWarnings False

    Set WRI_ObjWS = DBEngine(0)
    Set WRI_objDb = CurrentDb

    ' 加瀬DB接続
    If fncConnectDB = False Then Exit Function  ' 2005/11/21 ADD T.SUZUKI Start

    WRI_ObjWS.BeginTrans  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    ' コンテナマスタ更新処理
    If WRI_RTN_For_CNTA_MAST(0) = False Then
        GoTo Update_Err
    End If

    WRI_ObjWS.CommitTrans  ' ｺﾐｯﾄ
    WRI_RTN = True

    ' 加瀬DB切断
    If fncDisConnectDB = False Then Exit Function  ' 2005/11/21 ADD T.SUZUKI Start

    bChangeFlag_FKG020 = False                                                    'INSERT 20050718 K.ISHIZAKA
    Me.Caption = strCaption                                                       'INSERT 20050718 K.ISHIZAKA

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set WRI_objDb = Nothing
    Set WRI_ObjWS = Nothing
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    GoTo Exit_rtn

Update_Err:
    WRI_ObjWS.Rollback  ' ﾛｰﾙﾊﾞｯｸ
    MsgBox "更新中にエラーが発生しました。" & vbCrLf & "（処理は中断されました。）", vbCritical, strMsg_kaku
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :WRI_RTN_For_CNTA_MAST
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function WRI_RTN_For_CNTA_MAST(ByVal iMode As Integer) As Boolean

    '引数 ０：正更新
    '     １：逆更新

    'コンテナ予約ＷＫからコンテナマスタに予約をセットする。
    '解約チェックがされている時は、コンテナマスタからクリアする。
    '既に予約されている場合は、変更のあったコンテナのみ変更する。

    Dim strSQL  As String
    Dim vCode   As Variant

' 2005/11/21 ADD T.SUZUKI Start
    Dim dteWKDate As Date
    Dim objRs     As Recordset
    Dim objRsWk   As Recordset
' 2005/11/21 ADD T.SUZUKI End

    'Stop

    On Error GoTo err_rtn

    WRI_RTN_For_CNTA_MAST = False

' 2005/11/21 ADD T.SUZUKI Start
    ' 日付セット
    Dim strWkDate As String

    strWkDate = Format(vDATE, "yyyy/mm/dd HH:NN:SS")

    '契約有り、予約追加／変更の時*****************
    strSQL = " SELECT " & Chr(13)
    strSQL = strSQL & " CNTA_YARD, " & Chr(13)   ' ヤードコード
    strSQL = strSQL & " CNTA_NO, " & Chr(13)     ' コンテナ番号
    strSQL = strSQL & " CNTA_YNAME, " & Chr(13)  ' 予約客先名
    strSQL = strSQL & " CNTA_YTEL " & Chr(13)    ' 予約客先電話番号
    strSQL = strSQL & "  FROM FKS020_コンテナ予約ＷＫ " & Chr(13)
    strSQL = strSQL & " WHERE CNTA_CHANGE = 1 " & Chr(13)
    strSQL = strSQL & "   AND CNTA_KCODE  > 0 " & Chr(13)
    strSQL = strSQL & "   AND CNTA_CHECK  = 0 " & Chr(13)
    strSQL = strSQL & "   AND CNTA_DATE   = CDATE(#" & strWkDate & "#) " & Chr(13)
    strSQL = strSQL & " ORDER BY CNTA_YARD, " & Chr(13)
    strSQL = strSQL & "          CNTA_STEP, " & Chr(13)
    strSQL = strSQL & "          CNTA_NO "
    Set objRsWk = objDb.OpenRecordset(strSQL, dbReadOnly)

    Do Until objRsWk.EOF
        ' コンテナマスタオープン
        strSQL = " SELECT " & Chr(13)
        strSQL = strSQL & " CNTA_YNAME, " & Chr(13)  ' 予約客先名
        strSQL = strSQL & " CNTA_YTEL, " & Chr(13)   ' 予約客先電話番号
        strSQL = strSQL & " CNTA_UPDATE " & Chr(13)  ' 更新日
        strSQL = strSQL & "  FROM CNTA_MAST " & Chr(13)
        strSQL = strSQL & " WHERE CNTA_CODE = " & objRsWk.Fields("CNTA_YARD") & Chr(13)  ' ヤードコード
        strSQL = strSQL & "   AND CNTA_NO   = " & objRsWk.Fields("CNTA_NO")              ' コンテナ番号
        Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

        If objRs.EOF = False Then
            objRs.Edit

            objRs.Fields("CNTA_YNAME") = objRsWk.Fields("CNTA_YNAME")
            objRs.Fields("CNTA_YTEL") = objRsWk.Fields("CNTA_YTEL")
            objRs.Fields("CNTA_UPDATE") = Format(Now, "yyyy/mm/dd")

            objRs.UPDATE
        End If
        objRs.Close
        Set objRs = Nothing

        objRsWk.MoveNext
    Loop
    objRsWk.Close
    Set objRsWk = Nothing

    '契約無し、予約追加／変更の時*****************
    strSQL = " SELECT " & Chr(13)
    strSQL = strSQL & " CNTA_YARD, " & Chr(13)     ' ヤードコード
    strSQL = strSQL & " CNTA_NO, " & Chr(13)       ' コンテナ番号
    strSQL = strSQL & " CNTA_YNAME, " & Chr(13)    ' 予約客先名
    strSQL = strSQL & " CNTA_YTEL, " & Chr(13)     ' 予約客先電話番号
    strSQL = strSQL & " CNTA_RENTKG, " & Chr(13)   ' 予約賃料
    strSQL = strSQL & " CNTA_SECUKG " & Chr(13)    ' 予約保証金
    strSQL = strSQL & "  FROM FKS020_コンテナ予約ＷＫ " & Chr(13)
    strSQL = strSQL & " WHERE CNTA_CHANGE = 1 " & Chr(13)
    strSQL = strSQL & "   AND CNTA_KCODE  = 0 " & Chr(13)
    strSQL = strSQL & "   AND CNTA_CHECK  = 0 " & Chr(13)
    strSQL = strSQL & "   AND CNTA_DATE   = CDATE(#" & strWkDate & "#) " & Chr(13)
    strSQL = strSQL & " ORDER BY CNTA_YARD, " & Chr(13)
    strSQL = strSQL & "          CNTA_STEP, " & Chr(13)
    strSQL = strSQL & "          CNTA_NO "
    Set objRsWk = objDb.OpenRecordset(strSQL, dbReadOnly)

    Do Until objRsWk.EOF
        ' コンテナマスタオープン
        strSQL = " SELECT " & Chr(13)
        strSQL = strSQL & " CNTA_YNAME, " & Chr(13)
        strSQL = strSQL & " CNTA_YTEL, " & Chr(13)
        strSQL = strSQL & " CNTA_YRENTKG, " & Chr(13)
        strSQL = strSQL & " CNTA_SECUKG, " & Chr(13)
        strSQL = strSQL & " CNTA_UPDATE " & Chr(13)
        strSQL = strSQL & "  FROM CNTA_MAST " & Chr(13)
        strSQL = strSQL & " WHERE CNTA_CODE = " & objRsWk.Fields("CNTA_YARD") & Chr(13)  ' ヤードコード
        strSQL = strSQL & "   AND CNTA_NO   = " & objRsWk.Fields("CNTA_NO")              ' コンテナ番号
        Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

        If objRs.EOF = False Then
            objRs.Edit

            objRs.Fields("CNTA_YNAME") = objRsWk.Fields("CNTA_YNAME")
            objRs.Fields("CNTA_YTEL") = objRsWk.Fields("CNTA_YTEL")
            objRs.Fields("CNTA_YRENTKG") = objRsWk.Fields("CNTA_RENTKG")
            objRs.Fields("CNTA_SECUKG") = objRsWk.Fields("CNTA_SECUKG")
            objRs.Fields("CNTA_UPDATE") = Format(Now, "yyyy/mm/dd")

            objRs.UPDATE
        End If
        objRs.Close
        Set objRs = Nothing

        objRsWk.MoveNext
    Loop
    objRsWk.Close
    Set objRsWk = Nothing

    '*********************************************
    ' ２０００／０２／２３
    ' 契約有り、契約無しの２パターンのＳＱＬを１本化
    '*********************************************
    '予約取消しの時*****************
    strSQL = " SELECT " & Chr(13)
    strSQL = strSQL & " CNTA_YARD, " & Chr(13)     ' ヤードコード
    strSQL = strSQL & " CNTA_NO " & Chr(13)        ' コンテナ番号
    strSQL = strSQL & " FROM FKS020_コンテナ予約ＷＫ " & Chr(13)
    strSQL = strSQL & " WHERE CNTA_CHECK = -1 " & Chr(13)
    strSQL = strSQL & "   AND CNTA_DATE = CDATE(#" & strWkDate & "#) " & Chr(13)
    strSQL = strSQL & " ORDER BY CNTA_YARD, "
    strSQL = strSQL & "          CNTA_STEP, "
    strSQL = strSQL & "          CNTA_NO "
    Set objRsWk = objDb.OpenRecordset(strSQL, dbReadOnly)

    Do Until objRsWk.EOF
        ' コンテナマスタオープン
        strSQL = " SELECT " & Chr(13)
        strSQL = strSQL & " CNTA_YCODE, " & Chr(13)    ' ヤードコード
        strSQL = strSQL & " CNTA_YNAME, " & Chr(13)    ' 予約客先名
        strSQL = strSQL & " CNTA_YTEL, " & Chr(13)     ' 予約客先電話番号
        strSQL = strSQL & " CNTA_DEPOSIT, " & Chr(13)  ' 予約金
        strSQL = strSQL & " CNTA_YRENTKG, " & Chr(13)  ' 予約賃料
        strSQL = strSQL & " CNTA_SECUKG, " & Chr(13)   ' 予約保証金
        strSQL = strSQL & " CNTA_UPDATE " & Chr(13)    ' 更新日
        strSQL = strSQL & "  FROM CNTA_MAST " & Chr(13)
        strSQL = strSQL & " WHERE CNTA_CODE = " & objRsWk.Fields("CNTA_YARD") & Chr(13)  ' ヤードコード
        strSQL = strSQL & "   AND CNTA_NO   = " & objRsWk.Fields("CNTA_NO")              ' コンテナ番号
        Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

        If objRs.EOF = False Then
            objRs.Edit

            objRs.Fields("CNTA_YCODE") = 0
            objRs.Fields("CNTA_YNAME") = Null
            objRs.Fields("CNTA_YTEL") = Null
            objRs.Fields("CNTA_DEPOSIT") = 0
            objRs.Fields("CNTA_YRENTKG") = 0
            objRs.Fields("CNTA_SECUKG") = 0
            objRs.Fields("CNTA_UPDATE") = Format(Now, "yyyy/mm/dd")

            objRs.UPDATE
        End If
        objRs.Close
        Set objRs = Nothing

        objRsWk.MoveNext
    Loop
    objRsWk.Close
    Set objRsWk = Nothing
' 2005/11/21 ADD T.SUZUKI End



  WRI_RTN_For_CNTA_MAST = True

Exit_rtn:
    Exit Function

err_rtn:
    '*** エラー処理（Ｒｏｌｌｂａｃｋ） ***
    Dim errLoop As Error

    If DBEngine.Errors.Count > 0 Then
        For Each errLoop In DBEngine.Errors
            MsgBox "Error Number:" + CStr(errLoop.Number) + vbCrLf + errLoop.Description
        Next
    End If

  'MsgBox "ｴﾗｰ番号" & ERR.Number & vbCrLf & ERR.Description
  GoTo Exit_rtn
End Function
'==============================================================================*
'
'        MODULE_NAME      :確認と保存
'        MODULE_ID        :SaveData
'        CREATE_DATE      :20050718 K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function SaveData() As Boolean
    If (Nz(Me.cmb_ヤードコード) <> "") And bChangeFlag_FKG020 Then
        If MsgBox("修正中の内容を保存します。よろしいですか？", vbQuestion + vbYesNo, strMsg_kaku) = vbYes Then
            SaveData = WRI_RTN()
            Exit Function
        End If
    End If
    SaveData = True
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_取消_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_取消_Click()

    intAns = MsgBox("データ入力を取り消します。よろしいですか？", vbQuestion + vbYesNo + vbDefaultButton2, strMsg_kaku)
    If intAns = vbNo Then
        GoTo Exit_rtn
    End If

    Call DELWK_RTN     ' ワークテーブルのデータ削除処理
    Call CLR_RTN       ' 画面初期化処理
    Call Etc_Lock_Rtn  ' 2005/11/21 ADD T.SUZUKI

Exit_rtn:
    Me.cmb_ヤードコード.SetFocus
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_保存_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_Click()

    On Error GoTo err_rtn

' 2005/11/21 ADD T.SUZUKI Start
    ' ヤードコードが未選択の場合、処理を抜ける
    If Nz(Me.cmb_ヤードコード, "") = "" Then
        Exit Sub
    End If
' 2005/11/21 ADD T.SUZUKI End

    '確認メッセージ
    'intAns = MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, strMsg_kaku)               'DELETE 20050718 K.ISHIZAKA
    intAns = MsgBox("修正中の内容を保存します。よろしいですか？", vbQuestion + vbYesNo, strMsg_kaku)  'INSERT 20050718 K.ISHIZAKA

    If intAns = vbNo Then
        GoTo Exit_rtn
    End If

    If WRI_RTN = True Then
        Call DELWK_RTN     ' ワークテーブルのデータ削除処理
        Call CLR_RTN       ' 画面初期化処理
        Call Etc_Lock_Rtn  ' 2005/11/21 ADD T.SUZUKI
    End If

Exit_rtn:
    Me.cmb_ヤードコード.SetFocus
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :ETC_LOCK_RTN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Etc_Lock_Rtn()

    On Error Resume Next  ' 2005/11/28 ADD T.SUZUKI

    '入力区分によるコントロールのロックを設定
    Me.サブ上.Enabled = False
    Me.サブ下.Enabled = False

' 2005/11/21 REP T.SUZUKI Start
    With Me
        If Nz(.cmb_ヤードコード, "") = "" Then
            .cmd_保存.Enabled = False
            .cmd_取消.Enabled = False
    
'            .cmd_toiawase.Enabled = False       'INSERT 2019/11/20 Y.WADA  'DELETE 2019/12/03 Y.WADA
            .cmd_YUke.Enabled = False
            .cmd_WEB表示.Enabled = False
            .cmd_CAMPB.Enabled = False
            .cmd_近隣ヤード.Enabled = False
            .cmd_配置.Enabled = False
            .cmd_写真.Enabled = False  '2018/03/16 M.HONDA INS
        Else
            .cmd_保存.Enabled = True
            .cmd_取消.Enabled = True

'            .cmd_toiawase.Enabled = True       'INSERT 2019/11/20 Y.WADA   'DELETE 2019/12/03 Y.WADA
            .cmd_YUke.Enabled = True
            .cmd_WEB表示.Enabled = True
            .cmd_CAMPB.Enabled = True
            .cmd_近隣ヤード.Enabled = True
            .cmd_配置.Enabled = True
            .cmd_写真.Enabled = True   '2018/03/16 M.HONDA INS
        End If

        '20210708 M.HONDA INS
        If Me.txt_ACC = 8 Then
            .cmd_YUke.Enabled = False
            .cmd_CAMPB.Enabled = False
            .cmd_近隣ヤード.Enabled = False
            .cmd_紹介制限.Enabled = False
        End If
        '20210708 M.HONDA INS
        
     End With
'    Me.cmd_保存.Enabled = True
'    Me.cmd_取消.Enabled = True
' 2005/11/21 REP T.SUZUKI End
    Me.cmd_終了.Enabled = True
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Clickイベント
'        MODULE_ID        :cmd_終了_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
    On Error Resume Next
    doCmd.Close acForm, Me.FormName
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルのデータ削除処理
'        MODULE_ID        :DELWK_RTN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub DELWK_RTN()

    Dim strSQL     As String

    On Error Resume Next

    If IsDate(vDATE) = False Then Exit Sub  ' 2005/11/21 DEL T.SUZUKI

    'ワークの削除
    strSQL = ""
    strSQL = " delete from FKS020_コンテナ予約ＷＫ"
    strSQL = strSQL & " where CNTA_DATE =" & "CDATE(#" & Format(vDATE, "yyyy/mm/dd HH:NN:SS") & "# )"
    strSQL = strSQL & " Or CNTA_DATE <=" & "CDATE(#" & Format(vDATE, "yyyy/mm/dd HH:NN:SS") & "# )"

    objDb.Execute strSQL, dbFailOnError
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ヤードマスタ問い合わせ(上段、下段の確認)
'        MODULE_ID        :CHK_CNTA_MAST2
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_CNTA_MAST2()

    Dim strSQL       As String
    Dim vCode        As Variant
    Dim vRet, vRET2  As Variant
    Dim objRs        As Recordset  ' 2005/11/21 ADD T.SUZUKI

    On Error GoTo err_rtn

    CHK_CNTA_MAST2 = False

    vCode = Nz(Me.cmb_ヤードコード, "")

    If vCode = "" Then
        CHK_CNTA_MAST2 = True
        GoTo Exit_rtn
    End If

    doCmd.Hourglass True
    doCmd.SetWarnings False

    vStep = 2         '両方
    strSQL = ""
    strSQL = strSQL & " CNTA_CODE =" & Val(vCode)
    strSQL = strSQL & " AND CNTA_STEP = 0"
    vRet = Nz(DLookup("CNTA_NO", "dbo_CNTA_MAST", strSQL), "")

    If Nz(vRet, "") = "" Then
        vStep = 1  '下段
    End If

    strSQL = ""
    strSQL = strSQL & " CNTA_CODE =" & Val(vCode)
    strSQL = strSQL & " AND CNTA_STEP = 1"
    vRET2 = Nz(DLookup("CNTA_NO", "dbo_CNTA_MAST", strSQL), "")

    If Nz(vRET2, "") = "" Then
        vStep = 0  '上段
    End If

    If Nz(vRet, "") = "" And Nz(vRET2, "") = "" Then
        CHK_CNTA_MAST2 = False
        GoTo Exit_rtn
    End If

    CHK_CNTA_MAST2 = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set objRs = Nothing
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function
'==============================================================================*
'
'        MODULE_NAME      :ヤードマスタ問い合わせ
'        MODULE_ID        :CHK_YARD_MAST
'        Parameter    IN  :iMode
'                                0 = 存在チェック
'                                1 = データ参照
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Chk_Yard_Mast(ByVal iMode As Integer) As Boolean

    Dim strSQL As String
    Dim vCode  As Variant
    Dim vRet   As Variant
    Dim i      As Integer
    Dim vRET2  As Variant
    Dim objRs  As Recordset  ' 2005/11/21 ADD T.SUZUKI
    Dim strDataSource As String
    
    On Error GoTo err_rtn

    Chk_Yard_Mast = False

    vCode = Nz(Me.cmb_ヤードコード, "")

    If vCode = "" Then
        Chk_Yard_Mast = True
        GoTo Exit_rtn
    End If


    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))

    doCmd.Hourglass True
    doCmd.SetWarnings False

    If iMode = 0 Then
        '存在チェック
        strSQL = ""
        strSQL = strSQL & " YARD_CODE =" & Val(vCode)
        vRet = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strSQL), "")
      
        If vRet = "" Then
            MsgBox "該当するヤードマスタがありません。", vbInformation, strMsg_kaku
            'CHK_YARD_MAST = True     'Del 01-07-2000 Araki
            GoTo Exit_rtn             'Add 01-07-2000 Araki
        End If

    Else
        Me.cmd_CAMPB.Enabled = False                                              'INSERT 20041106 N.MIURA

        ' 加瀬DB接続
        If fncConnectDB = False Then Exit Function                                                  ' 2005/11/21 ADD T.SUZUKI

        strSQL = ""
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "YARD_MAST.YARD_NAME,"
        strSQL = strSQL & "YARD_MAST.YARD_ADDR_1,"
        strSQL = strSQL & "YARD_MAST.YARD_ADDR_2,"
        strSQL = strSQL & "YARD_MAST.YARD_ADDR_3,"
        strSQL = strSQL & "YARD_MAST.YARD_NOTE,"
        strSQL = strSQL & "YARD_MAST.YARD_SEV1N,"
        strSQL = strSQL & "YARD_MAST.YARD_SEV2N,"
        strSQL = strSQL & "YARD_MAST.YARD_SEV3N,"
        strSQL = strSQL & "YARD_MAST.YARD_ENDEN,"
        strSQL = strSQL & "YARD_MAST.YARD_BIKON "
        strSQL = strSQL & ",YARD_MAST.YARD_BEGIN_DAY "    '営業開始日       2006/09/24 add tajima
        strSQL = strSQL & ",YARD_MAST.YARD_END_DAY "      'ヤード解約予定日 2006/09/24 add tajima
        strSQL = strSQL & ",YARD_MAST.YARD_RENTEND_DAY "  '営業終了日       2006/09/24 add tajima
        strSQL = strSQL & ",YARD_MAST.YARD_INLIMIT_DAY INLIMIT_DAY" '紹介制限日      2007/02/08 add tajima
        strSQL = strSQL & ",YARD_MAST.YARD_INLIMIT_YCODE " '紹介制限ヤード  2007/02/08 add tajima
        strSQL = strSQL & ",INLIMIT_YARD.YARD_NAME LIMIT_NAME " '制限ヤード名2007/02/08 add tajima
        strSQL = strSQL & ",TANT_MENTE.TANTM_TANTN AS MNT_TAN  "       '' 20130418 M.HONDA INS
        strSQL = strSQL & ",TANT_JUN.TANTM_TANTN AS JUN_TAN "         '' 20130418 M.HONDA INS
        strSQL = strSQL & ",NAME_MAST.NAME_NAME AS YARD_SPAR "         '' 2014/04/07 M.HONDA INS
        strSQL = strSQL & ",TANT_TAN.TANTM_TANTN AS YARD_TANN "         '' 2017/09/07 M.HONDA INS
        
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "YARD_MAST "                                                              ' 2005/11/21 Rep T.SUZUKI
        strSQL = strSQL & "LEFT OUTER JOIN YARD_MAST AS INLIMIT_YARD ON "           '2007/02/08 add tajima
        strSQL = strSQL & "YARD_MAST.YARD_INLIMIT_YCODE = INLIMIT_YARD.YARD_CODE "  '2007/02/08 add tajima
        '' 20130418 M.HONDA START
        strSQL = strSQL & "LEFT OUTER JOIN "
        strSQL = strSQL & strDataSource & ".dbo.TANT_MAST AS TANT_MENTE ON "
        strSQL = strSQL & "YARD_MAST.YARD_MNT_TANTO = TANT_MENTE.TANTM_TANTC AND "
        strSQL = strSQL & "TANT_MENTE.TANTM_BUMOC = '" & pstrBumoCD & "'"
        strSQL = strSQL & "LEFT OUTER JOIN "
        strSQL = strSQL & strDataSource & ".dbo.TANT_MAST AS TANT_JUN ON "
        strSQL = strSQL & "YARD_MAST.YARD_JUN_TANTO = TANT_JUN.TANTM_TANTC AND "
        strSQL = strSQL & "TANT_JUN.TANTM_BUMOC = '" & pstrBumoCD & "'"
        '' 20130418 M.HONDA END
        '2017/09/07 M.HONDA INS
        strSQL = strSQL & "LEFT OUTER JOIN "
        strSQL = strSQL & strDataSource & ".dbo.TANT_MAST AS TANT_TAN ON "
        strSQL = strSQL & "YARD_MAST.YARD_TANTO = TANT_TAN.TANTM_TANTC AND "
        strSQL = strSQL & "TANT_TAN.TANTM_BUMOC = '" & pstrBumoCD & "'"
        '2017/09/07 M.HONDA INS
        
        strSQL = strSQL & "LEFT OUTER JOIN NAME_MAST ON "
        strSQL = strSQL & "NAME_ID = '061' AND "
        strSQL = strSQL & "NAME_CODE = YARD_MAST.YARD_SPARE_KEY "
        
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "YARD_MAST.YARD_CODE = " & Val(vCode)
        Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)  ' 2005/11/21 Rep T.SUZUKI

        If Not (objRs.BOF And objRs.EOF) Then
            With Me
                .txt_ヤード名 = objRs.Fields("YARD_NAME")
                .txt_ヤード住所 = objRs.Fields("YARD_ADDR_1") & _
                                 objRs.Fields("YARD_ADDR_2") & _
                                 objRs.Fields("YARD_ADDR_3")
                .txt_付近メモ = objRs.Fields("YARD_NOTE")
                .txt_備考 = objRs.Fields("YARD_BIKON")
                .txt_営業開始日 = objRs.Fields("YARD_BEGIN_DAY")    ' 2006/09/24 add tajima
                .txt_営業終了日 = objRs.Fields("YARD_RENTEND_DAY")  ' 2006/09/24 add tajima
                .txt_解約予定日 = objRs.Fields("YARD_END_DAY")      ' 2006/09/24 add tajima
                
                .txt_メンテ担当者 = objRs.Fields("MNT_TAN")      ' 20130418 M.HONDA
                .txt_巡回担当者 = objRs.Fields("JUN_TAN")        ' 20130418 M.HONDA
                
                
                .txt_スペア = objRs.Fields("YARD_SPAR")        ' 20130418 M.HONDA
                .txt_リーダ = objRs.Fields("YARD_TANN")         ' 2017/09/07 M.HONDA
                
                
                
                If Not (IsNull(objRs.Fields("YARD_SEV1N")) Or objRs.Fields("YARD_SEV1N") = "" Or objRs.Fields("YARD_SEV1N") = " ") Or _
                    Not (IsNull(objRs.Fields("YARD_SEV2N")) Or objRs.Fields("YARD_SEV2N") = "" Or objRs.Fields("YARD_SEV2N") = " ") Or _
                    Not (IsNull(objRs.Fields("YARD_SEV3N")) Or objRs.Fields("YARD_SEV3N") = "" Or objRs.Fields("YARD_SEV3N") = " ") Or _
                    Not (IsNull(objRs.Fields("YARD_ENDEN")) Or objRs.Fields("YARD_ENDEN") = "" Or objRs.Fields("YARD_ENDEN") = " ") Then
                    
                    .cmd_CAMPB.Enabled = True
                End If
                '▼ 2007/02/08 add tajima
                ' 紹介制限があって且つ、まだかかっていればその情報を表示
                If Not IsNull(objRs.Fields("INLIMIT_DAY")) And objRs.Fields("INLIMIT_DAY") >= DATE Then
                    txt_紹介制限情報.VALUE = "本ヤードは【" & Format(objRs.Fields("YARD_INLIMIT_YCODE"), "000000") & "：" & objRs.Fields("LIMIT_NAME") _
                                                & "】が" & vbCrLf & "ヤード解約の為 " & objRs.Fields("INLIMIT_DAY") _
                                                & " まで代替えヤードとして紹介制限をしています"
                    txt_紹介制限情報.Visible = True
                Else
                    txt_紹介制限情報.VALUE = ""
                    txt_紹介制限情報.Visible = False
                End If
                '▲ 2007/02/08 add tajima
            End With
        End If                                                                    'INSERT 20041106 N.MIURA ENDED

        ' 加瀬DB切断
        If fncDisConnectDB = False Then Exit Function                                               ' 2005/11/21 ADD T.SUZUKI

    End If

    Chk_Yard_Mast = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set objRs = Nothing
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function
'==============================================================================*
'
'        MODULE_NAME      :txt_END_Enterイベント
'        MODULE_ID        :txt_END_Enter
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_END_Enter()

    Dim vCode  As Variant
    
    vCode = Nz(Me.cmb_ヤードコード, "")
                    
    If vCode = "" Then
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If

'--- ADD 2000/06/13 -------------------
    If KYDATE < DATE Then
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If
'--- ADD 2000/06/13 -------------------

    '2005/05/29 MITSUNO
    '予約名/電話番号変更時、予約トランに受付番号が入っていたら
    'ｺﾝﾃﾅﾏｽﾀへUpdateしない処理追加の為、以下の処理をコメント
    Me.サブ上.SetFocus
'    Me.サブ上.Form!txt_予約客先名.SetFocus
End Sub
''==============================================================================*
''
''        MODULE_NAME      :txt_HEnd_GotFocusイベント
''        MODULE_ID        :txt_HEnd_GotFocus
''        CREATE_DATE      :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_HEnd_GotFocus()
'    Me.cmb_ヤードコード.SetFocus
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CmKeyPress_For_D
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CmKeyPress_For_D(ojCtrl As Control, KeyAscii As Integer, inLen As Integer)

    On Error GoTo CmKeyPress_For_Dec_Err

    Dim stValue As String

    With ojCtrl
        Select Case KeyAscii
            Case 8, 9, 10, 13                               ' 8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9"), Asc("."), Asc("-")   ' 0～9迄と .と - を許可する
            
                If Len(.Text) > 1 Then
                  stValue = CStr(Abs(CLng(Format$(.Text, "#0.00"))))
                End If

                If Len(stValue) > (inLen - 1) And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With

    Exit Sub

CmKeyPress_For_Dec_Err:

  KeyAscii = 0

End Sub
'==============================================================================*
'
'        MODULE_NAME      :終了メッセージを表示
'        MODULE_ID        :Form_Unload
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Err_Form_Unload

    '' 2013/06/03 M.HONDA START
'    strMsg = "終了します。よろしいですか？"
'    intAns = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, "確認")
'
'    If intAns = vbNo Then
'        Cancel = True
'        Me.cmb_ヤードコード.SetFocus
'        Exit Sub
'    End If
    '' 2013/06/03 M.HONDA END
    
    ' ワークテーブルのデータ削除処理
    Call DELWK_RTN

    Exit Sub

Err_Form_Unload:
    MsgBox Error$
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :フォーカスを先頭に戻す。
'        MODULE_ID        :txt_Start1_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Start1_GotFocus()
    Me.cmb_ヤードコード.SetFocus
End Sub
'==============================================================================*
'
'        MODULE_NAME      :サブ下_Enterイベント
'        MODULE_ID        :サブ下_Enter
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub サブ下_Enter()
    On Error Resume Next
    doCmd.GoToRecord , , acFirst
End Sub
'==============================================================================*
'
'        MODULE_NAME      :サブ上_Enterイベント
'        MODULE_ID        :サブ上_Enter
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub サブ上_Enter()
    On Error Resume Next
    doCmd.GoToRecord , , acFirst
End Sub
'==============================================================================*
'
'        MODULE_NAME      :サブ上A_Enterイベント
'        MODULE_ID        :サブ上A_Enter
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub サブ上A_Enter()
    On Error Resume Next
    doCmd.GoToRecord , , acFirst
End Sub
'==============================================================================*
'
'        MODULE_NAME      :外部から再表示させる
'        MODULE_ID        :Reload
'        CREATE_DATE      :2005/06/03 K.I
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub Reload()                                                            'DELETE 20050718 K.ISHIZAKA
'ここから                                                                       'INSERT 20050718 K.ISHIZAKA
Public Sub Reload(ByVal strYard1 As String, ByVal strYard2 As String)
    If bChangeFlag_FKG020 Then
        Exit Sub
    End If
    If (strYard1 <> Me.cmb_ヤードコード) And (strYard2 <> Me.cmb_ヤードコード) Then
        Exit Sub
    End If
'ここまで                                                                       'INSERT 20050718 K.ISHIZAKA
    Call cmb_ヤードコード_AfterUpdate
End Sub
'==============================================================================*
'
'        MODULE_NAME      :DB接続
'        MODULE_ID        :fncConnectDB
'        CREATE_DATE      :2005/11/08
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncConnectDB() As Boolean

    Dim strConnect    As String
    Dim strDataSource As String

    On Error GoTo fncConnectDB_Err

    fncConnectDB = False

    ' -----------------------------------------------------------------------------------------------------------
    ' コンテナDB接続
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & pstrBumoCD & "'"), "")

    If strDataSource = "" Then
        MsgBox "SETU_TABLの設定が不正です。", vbExclamation, PROG_ID
        Set pobjKONT_DB = Nothing
        Exit Function
    Else
        Set pobjKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(pstrBumoCD))
    End If
    ' -----------------------------------------------------------------------------------------------------------

    fncConnectDB = True

    Exit Function

fncConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function
'==============================================================================*
'
'        MODULE_NAME      :DB切断
'        MODULE_ID        :fncDisConnectDB
'        CREATE_DATE      :2005/11/08
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDisConnectDB() As Boolean

    On Error GoTo fncDisConnectDB_Err

    fncDisConnectDB = False

    ' *** コンテナDB ***
    pobjKONT_DB.Close
    Set pobjKONT_DB = Nothing
    '▼ 2006/09/24 tajima add ▼
    If Not pobjKASE_DB Is Nothing Then
      pobjKASE_DB.Close
      Set pobjKASE_DB = Nothing
    End If
    '▲ 2006/09/24 tajima add ▲

    fncDisConnectDB = True

    Exit Function

fncDisConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function
'==============================================================================*
'
'        MODULE_NAME      :KASE-DB接続
'        MODULE_ID        :fncConnectKaseDB
'        CREATE_DATE      :2006/09/24
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncConnectKaseDB() As Boolean

    Dim strDataSource As String

    On Error GoTo Exception

    fncConnectKaseDB = False

    ' -----------------------------------------------------------------------------------------------------------
    ' コンテナDB接続
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"), "")

    If strDataSource = "" Then
        MsgBox "SETU_TABLの設定が不正です。", vbExclamation, PROG_ID
        Set pobjKASE_DB = Nothing
        Exit Function
    Else
        Set pobjKASE_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00())
    End If
    ' -----------------------------------------------------------------------------------------------------------

    fncConnectKaseDB = True

    Exit Function

Exception:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_紹介制限_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_紹介制限_Click()

  On Error GoTo Exception

    ' ヤードコードが未選択の場合、処理を抜ける
    If Nz(Me.cmb_ヤードコード, "") = "" Then
        Exit Sub
    End If


    '紹介制限画面　起動
    doCmd.OpenForm "FVS200", , , , , , Format$(Me.cmb_ヤードコード, "000000")

    Exit Sub
    
Exception:
    Call MSZZ024_M00("cmd_紹介制限_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :オーナーサブフォーム設定
'        MODULE_ID        :subCreSubf_owner
'        PARM             :blnCreate        true...作業テーブル作成(I)
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCreSubf_owner(Optional blnCreate As Boolean = False)
    Dim strSQL          As String

    On Error GoTo Exception

    If Not subf_owner.Visible Then
        'オーナーサブフォームが非表示の場合は何もしない
        Exit Sub
    End If

    strSQL = strSQL & vbCrLf & "DECLARE @YMD AS VARCHAR(8) = CONVERT(varchar, GETDATE(), 112);"
    strSQL = strSQL & vbCrLf & "DECLARE @SYO1C AS VARCHAR(6) = '" & Nz(cmb_ヤードコード) & "';"
    strSQL = strSQL & vbCrLf & "SELECT  DISTINCT"
    strSQL = strSQL & vbCrLf & "    HKEIT_BUMOC"
    strSQL = strSQL & vbCrLf & ",   HKEIT_SYO1C"
    strSQL = strSQL & vbCrLf & ",   HKEIT_SHIRC"
    strSQL = strSQL & vbCrLf & ",   HKEIT_KANRI"
    strSQL = strSQL & vbCrLf & ",   cKanri.CODET_NAMEN AS HKEIT_KANRI_nm"
    strSQL = strSQL & vbCrLf & ",   SHIRM_SHIRN"
    strSQL = strSQL & vbCrLf & ",   SHIRM_YUBIB"
    strSQL = strSQL & vbCrLf & ",   SHIRM_ADR1N"
    strSQL = strSQL & vbCrLf & ",   SHIRM_ADR2N"
    strSQL = strSQL & vbCrLf & ",   SHIRM_ADR3N"
    strSQL = strSQL & vbCrLf & ",   SHIRM_TELFB"
    strSQL = strSQL & vbCrLf & ",   SHIRM_KEITB"
    strSQL = strSQL & vbCrLf & ",   SHIRM_FAXFB"
    strSQL = strSQL & vbCrLf & "FROM HKEI_TRAN"
    strSQL = strSQL & vbCrLf & "LEFT JOIN  CODE_TABL    cKanri"
    strSQL = strSQL & vbCrLf & "    ON  cKanri.CODET_SIKBC  = '221'"
    strSQL = strSQL & vbCrLf & "    AND cKanri.CODET_CODEC = HKEIT_KANRI"
    strSQL = strSQL & vbCrLf & "LEFT JOIN  SHIR_MAST    "
    strSQL = strSQL & vbCrLf & "    ON  SHIRM_BUMOC = 'ｦ'"
    strSQL = strSQL & vbCrLf & "    AND SHIRM_SHIRC = HKEIT_SHIRC"
    strSQL = strSQL & vbCrLf & "WHERE"
    strSQL = strSQL & vbCrLf & "    HKEIT_SHIRC NOT LIKE 'Z%'"
    strSQL = strSQL & vbCrLf & "AND HKEIT_SYO1C = @SYO1C"
    strSQL = strSQL & vbCrLf & "AND @YMD BETWEEN HKEIT_KAISD AND ISNULL(HKEIT_KAIYD, '99991231')"
    strSQL = strSQL & vbCrLf & "AND HKEIT_BUMOC = '" & pstrBumoCD & "'" 'INSERT 2019/12/11 Y.WADA
    subf_owner.Form.Painting = False     'フォームの描画をOFF
    If blnCreate Then
        subf_owner.Form.RecordSource = ""
        Call ADODB_CreateWorkTable_CurrentDb(WORK_TAB01, strSQL)
        subf_owner.Form.RecordSource = WORK_TAB01
    Else
        Call ADODB_Reset_CurrentDb(WORK_TAB01, strSQL)
        subf_owner.Form.Requery
    End If
    subf_owner.Form.Painting = True     'フォームの描画をON

Exit Sub
    
Exception:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
End Sub

'****************************  ended or program ********************************

' デバック用 関数群 -------------------------------------------------------------------------------
'==============================================================================*
' 開始時間セット
'==============================================================================*
Private Sub StartTimer_FS()
    mLngStartTime = timeGetTime()
End Sub

'==============================================================================*
' 処理時間取得
'==============================================================================*
Private Function StopTimer_FS() As Long
    StopTimer_FS = timeGetTime() - mLngStartTime
End Function

'Private Sub subTEST()
'
'    Dim lngI    As Long
'    Dim strTemp As String
'
'    ' スタート時間セット
'    StartTimer_FS
'
'    ' 時間稼ぎ
'    For lngI = 1 To 1000000
'        strTemp = ""
'    Next lngI
'
'    Debug.Print "計測時間：", StopTimer_FS()
'End Sub
' デバック用 関数群 -------------------------------------------------------------------------------
