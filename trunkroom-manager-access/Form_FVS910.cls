VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS910"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　修繕入力ダイアログ
'   プログラムＩＤ　：　FVS910
'   作　成　日　　　：  2006/04/22
'   作　成　者　　　：  Eagle-Soft T.SUZUKI
'**********************************************
'修正履歴
'   Ver             ：0.1
'   修　正　日　　　：2006/09/06
'   修　正　者　　　：tajima@eagle-soft
'   修　正　内　容　：①メンテナンス依頼書出力対応
'                   ：②貸止チェック削除
'                   ：③操作権限変更
'
'   Ver             ：0.2
'   修　正　日　　　：2007/01/26
'   修　正　者　　　：tajima@eagle-soft
'   修　正　内　容　：①ヤード修理入力機能追加対応
'                   ：②内容項目追加
'                   ：③ヤードの入力区分からクレームを外す
'
'   Ver             ：0.3
'   修　正　日　　　：2007/04/11
'   修　正　者　　　：tajima@eagle-soft
'   修　正　内　容　：①【紹介待ち】に対し修理・撤去登録したら紹介待ちを殺す
'                   ：②修理完了後、契約中でなければ紹介待ちとする
'
'   Ver             ：0.4
'   修　正　日　　　：2008/12/08
'   修　正　者　　　：M.RYU
'   修　正　内　容　：桁数オーバー時に、テキストボックスの内容を消さないように修正
'
'   Ver             ：0.5
'   修　正　日　　　：2011/03/10
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：顧客コメント時に表示項目を追加
'
'   Ver             ：0.6
'   修　正　日　　　：2012/06/21
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：対象が部屋のとき、統合ＤＢを更新する
'
'   Ver             ：0.7
'   修　正　日　　　：2013/06/17
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：対象がヤードのとき、クレームも選択可能とする
'                     対象が顧客のとき、時刻も入力可能とする
'
'   Ver             ：0.8
'   修　正　日　　　：2013/07/15
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：時刻がヌルのときの考慮、コロンに対応した時間のチェック
'
'   Ver             ：0.9
'   修　正　日　　　：2013/07/29
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：時刻の非表示（ヌルの時はシステム時刻を設定する）
'
'   Ver             ：1.0
'   修　正　日　　　：2013/09/24
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：登録時間・更新時間を秒まで保持する。
'
'   Ver             ：1.1
'   修　正　日　　　：2013/11/15
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：顧客ｺﾒﾝﾄの場合のみ時刻を表示
'
'   Ver             ：1.2
'   修　正　日　　　：2015/05/09
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：修繕鍵自動発番対応
'
'   Ver             ：1.3
'   修　正　日　　　：2017/11/11
'   修　正　者　　　：k.sato
'   修　正　内　容　：修繕一括修正対応
'
'   Ver             ：1.4
'   修　正　日　　　：2021/08/03
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：入力区分に「07:メンテ依頼」を追加（02:修理と同じ動作）
'
'   Ver             ：1.5
'   修　正　日　　　：2022/05/25
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：入力区分に「13:14:15:16:17」を追加（02:修理と同じ動作）
'
'   Ver             ：1.6
'   修　正　日　　　：2022/06/14
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：入力区分に「18」を追加（02:修理と同じ動作）
'
'   Ver             ：1.8
'   修　正　日　　　：2023/04/18
'   修　正　者　　　：M.IMAI
'   修　正　内　容　：CNTA_USE の更新条件を変更
'
'**********************************************
Option Compare Database
Option Explicit

' 定数宣言
Private Const P_参照モード      As String = "0"   ' 参照
Private Const P_新規モード      As String = "1"   ' 新規
Private Const P_修正モード      As String = "2"   ' 修正

Private Const P_対象区分コード  As String = "250" ' コードテーブル
Private Const P_修正区分コード  As String = "251" ' コードテーブル
Private Const P_ヤード修正区分コード  As String = "254" ' コードテーブル

Private Const P_COLOR_WHITE     As String = "16777215"   ' 背景色(白色)
Private Const P_COLOR_YELLOW    As String = "14680063"   ' 背景色(黄色)

Private Const P_DB_KONT         As Integer = 1
Private Const P_DB_KASE         As Integer = 2

' コンボボックスの中身を取得するWHERE句
'                                                                               'DELETE START 2013/06/17 K.ISHIZAKA
'Private Const P_入力区分ヤードWHERE   As String = "WHERE CODET_SIKBC = '250' AND CODET_CODEC IN(1,2,3) ORDER BY CODET_CODEC" '2007/01/26 4は外す
'Private Const P_入力区分顧客WHERE     As String = "WHERE CODET_SIKBC = '250' AND CODET_CODEC IN(3,4,9,10) ORDER BY CODET_CODEC"  '2011/03/10 9・10追加
'Private Const P_入力区分部屋WHERE     As String = "WHERE CODET_SIKBC = '250' AND CODET_CODEC IN(2,3,4,5) ORDER BY CODET_CODEC"
'                                                                               'DELETE END   2013/06/17 K.ISHIZAKA
'                                                                               'INSERT START 2013/06/17 K.ISHIZAKA
Private Const P_入力区分ヤードWHERE   As String = "WHERE CODET_SIKBC = '250' AND CODET_TOTOQ IN(1,13,14,134) ORDER BY CODET_CODEC" '2007/01/26 4は外す
Private Const P_入力区分顧客WHERE     As String = "WHERE CODET_SIKBC = '250' AND CODET_TOTOQ IN(4,14,34,134) ORDER BY CODET_CODEC"  '2011/03/10 9・10追加
Private Const P_入力区分部屋WHERE     As String = "WHERE CODET_SIKBC = '250' AND CODET_TOTOQ IN(3,13,34,134) ORDER BY CODET_CODEC"
'                                                                               'INSERT END   2013/06/17 K.ISHIZAKA

Private Const P_部屋修理区分WHERE     As String = "WHERE CODET_SIKBC = '251' ORDER BY CODET_CODEC"
Private Const P_ヤード修理区分WHERE   As String = "WHERE CODET_SIKBC = '254' ORDER BY CODET_CODEC"  '2007/01/26 add tajima

' メッセージ
Private Const P_MSG_KAKUNIN     As String = "確認"
Private Const P_MSG_ERR         As String = "エラー"

Private Const P_MSG_ERR_NODATA  As String = "存在しないコードです。"
Private Const P_MSG_ERR_DATE    As String = "日付を正しく入力してください。"
Private Const P_MSG_ERR_TYPEC   As String = "指定した入力区分は選択できません。"
Private Const P_MSG_ERR_INPUT   As String = "を入力してください。"

' Private変数宣言
Private pIs修正                 As Boolean
Private p画面標題               As String
Private pvarChangeDate          As Variant   ' 日付チェック・書換用の変数
Private pstrBumonCd             As String    ' 部門コード
Private pCntaUse                As Integer   ' コンテナ利用可否
Private pCntaReprKbn            As Integer   ' 修理中可否

Private pobjKONT_DB             As Database  ' KOMS系DB
Private pobjKASE_DB             As Database  ' KASE系DB
Private pobjDB                  As Database  ' Access系DB

' 構造体(起動パラメータ保持)
Private Type Type_SHUZEN
    修繕対象    As String                 ' 修繕対象(ヤード修繕／顧客修繕／コンテナ(部屋)修繕)
    処理モード  As Variant                ' 登録モード(参照(0) OR 新規(1) OR 修正(2))
    UNIQEC      As String                 ' ユニークコード
    YARDC       As String                 ' ヤードコード
    ROOMC       As String                 ' 部屋番号(コンテナ番号)
    KOKYC       As String                 ' 顧客コード
    契約番号    As String
    画面ID      As String                 ' 画面ID
    権限可否    As Boolean                ' Add 2006/09/06
    一括        As Boolean                ' Add 2017/11/11
End Type
Private pParamArgs As Type_SHUZEN

Private pstReprTran   As CmReprMod.Type_REPR_TRAN   ' 修繕テーブルイメージ構造体
Private pstKotoTran   As CmReprMod.Type_KOTO_TRAN   ' 顧客問合テーブルイメージ構造体
Private pstrUserName  As String                     ' ユーザID
Private blnUpdFlg     As Boolean                    ' 呼出元のリロード制御のフラグ
Private strMotokagi   As String                     ' 元鍵区分 '2015/02/27 M.HONDA INS
Private strTiho       As String                     ' 地方区分 '2015/02/27 M.HONDA INS


Private strReprKbn    As String                     ' 保持用修理区分 '2017/11/18 K.SATO INS

'==============================================================================*
'
'        MODULE_NAME      :cmb_修理区分_GotFocus イベント
'        MODULE_ID        :cmb_修理区分_GotFocus
'        CREATE_DATE      :2017/11/18
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_修理区分_GotFocus()
    
    '修理区分保管 2017-11-18 INS K.SATO
    strReprKbn = Nz(Me.cmb_修理区分, "")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    ' 適正サイズにする ヘルプによると 1cm は 567twip
    'doCmd.MoveSize , , 16.501 * 567, 9.485 * 567
    
    ' とりあえずMAXサイズで
    doCmd.Maximize

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Open イベント
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    Dim strParam()  As String
    Dim intLoopCnt  As Integer

' ↓↓↓本番用↓↓↓ -----------------------
    ' 起動パラメータ取得
    strParam = Split(Nz(Me.OpenArgs), ",")
' ↑↑↑本番用↑↑↑------------------------

' ↓↓↓検証用↓↓↓ -----------------------
    ' 起動パラメータ取得
' 【ヤード修繕】
'    strParam = Split("01,1,,100101,,,,FVS901,True", ",")  ' 新規 5/23確認
'    strParam = Split("01,2,{2BD98F1A-2BFD-4954-BE81-918F1340874E},100101,,,,FVS901,True", ",")  ' 修正 5/23確認
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_TITTLE.BackColor = RGB(255, 0, 204)
    End Select

' 【顧客修繕】
'    strParam = Split("04,1,,,,820031,,FVS902,True", ",")  ' 新規 5/23確認
'    strParam = Split("04,2,{DBD2EED6-098F-4274-8C5A-F049BC329F5C},,,820031,,FVS902,True", ",")  ' 修正 5/23確認

' 【コンテナ(部屋)修繕】
'    strParam = Split("03,1,,000012,870000,,,FVS902,True", ",")  ' 新規 5/23確認
'    strParam = Split("03,2,{D29F67AD-9840-4D42-AEA7-F178313AD5B8},000012,870000,", ",")  ' 修正 5/23確認
' ↑↑↑検証用↑↑↑------------------------

    ' パラメータ数チェック
    '2017-11-11 upd k.sato start
'    If UBound(strParam) <> 8 Then
'        Call MSZZ024_M10("OpenArgs", "起動パラメータが正しくありません。")
'        Cancel = True: Exit Sub
'    End If
    If strParam(7) = "FVS903" And UBound(strParam) <> 8 Then
        Call MSZZ024_M10("OpenArgs", "起動パラメータが正しくありません。")
        Cancel = True: Exit Sub
        End If
    If strParam(7) = "FVS904" And UBound(strParam) <> 9 Then
        Call MSZZ024_M10("OpenArgs", "起動パラメータが正しくありません。")
        Cancel = True: Exit Sub
    End If
    '2017-11-11 upd k.sato end
    ' 構造体に値をセット
    With pParamArgs
        .修繕対象 = strParam(0)     ' 処理区分
        .処理モード = strParam(1)   ' 登録モード
        .UNIQEC = strParam(2)       ' ユニークコード
        .YARDC = strParam(3)        ' ヤードコード
        .ROOMC = strParam(4)        ' 部屋番号
        .KOKYC = strParam(5)        ' 顧客コード
        .契約番号 = strParam(6)
        .画面ID = strParam(7)       ' 画面ID
        .権限可否 = CBool(strParam(8))  ' Add 2006/09/06
        '2017/11/11 add start
        .一括 = False
        If UBound(strParam) = 9 Then
            .一括 = CBool(strParam(9))
        End If
        '2017/11/11 add end
    End With
    
    Exit Sub
    
Exception:
    Call MSZZ024_M00("Form_Load", True)
End Sub


'==============================================================================*
'
'        MODULE_NAME      :Form_Load イベント
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :2006/09/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

  On Error GoTo Exception

    'ユーザID取得　※ユーザIDによってこの画面に更新が出来るかキメ
    pstrUserName = ""
    pstrUserName = MSZZ000.LsGetUserName

    ' カレントDB接続
    Set pobjDB = CurrentDb

    ' 部門コード取得
    Call subGetBumonName

    ' 画面標題設定
    p画面標題 = Me.Caption

    ' 画面タイトル設定
    Call setTitleName
'    txt_時刻.Visible = (pParamArgs.修繕対象 = CmReprMod.P_対象_顧客)           'DELETE 2013/07/29 K.ISHIZAKA 'INSERT 2013/06/17 K.ISHIZAKA
    '' 2013/11/13 M.HONDA UPD
    'txt_時刻.Visible = False                                                    'INSERT 2013/07/29 K.ISHIZAKA
    '' 2013/11/13 M.HONDA UPD
    lbl_日付.Caption = IIf(txt_時刻.Visible, "日付／時刻", "日付")              'INSERT 2013/06/17 K.ISHIZAKA

    ' 印刷ボタン使用可否設定 2006/09/06 tajima
    If pParamArgs.修繕対象 = CmReprMod.P_対象_部屋 Or _
       pParamArgs.修繕対象 = CmReprMod.P_対象_ヤード Then '2007/01/26 add tajima
        Me.cmb_印刷.Visible = True
        txt_時刻.Visible = False              ''2013/11/15 M.HONDA INS
    Else
        Me.cmb_印刷.Visible = False
        txt_時刻.Visible = True               ''2013/11/15 M.HONDA INS
    End If

    ' 画面の項目設定
    Call fncSetFrmItem

    ' 変更フラグ初期化
    Call subSetChangeFlag(False)
    
    '2015/03/25 M.HONDA INS
    '修理・鍵の場合には、再度変更させないため使用不可へ
    'If cmb_入力区分 = "02" And cmb_修理区分 = "05" Then                                                                    'DELETE 2021/08/03 N.IMAI
    If (cmb_入力区分 = CmReprMod.P_修繕_修理 Or cmb_入力区分 = CmReprMod.P_修繕_メンテ依頼) And cmb_修理区分 = "05" Then    'INSERT 2021/08/03 N.IMAI
        cmb_入力区分.Enabled = False
        cmb_修理区分.Enabled = False
        txt_内容壱サブ壱.Enabled = False
    End If
    '2015/03/25 M.HONDA INS
    
    blnUpdFlg = False

    '2017/11/11 K.SATO INS
    If pParamArgs.一括 Then
        Me.lbl_ikkatu.Visible = True
        Me.txt_ikkatu.Visible = True
        Me.txt_ikkatu.VALUE = FVS904_sub.getOnDataCSVRooms
    Else
        Me.lbl_ikkatu.Visible = False
        Me.txt_ikkatu.Visible = False
    End If

    Exit Sub
    
Exception:
    Call MSZZ024_M00("Form_Load", True)

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
    
    'F1:戻る
    Select Case KeyCode
    Case vbKeyF1
        KeyCode = 0
        Me.cmd_戻る.SetFocus
    Case vbKeyF2
        KeyCode = 0
        Me.cmd_保存.SetFocus
    Case vbKeyF3, vbKeyF5 To vbKeyF12
        KeyCode = 0
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_入力区分_AfterUpdate イベント
'        MODULE_ID        :cmb_入力区分_AfterUpdate
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_入力区分_AfterUpdate()

    Dim vCode       As Variant

  On Error GoTo Exception
  
    vCode = Nz(Me.cmb_入力区分, "")
    If vCode = "" Then
        Me.txt_入力区分 = ""
    Else
        Me.cmb_入力区分 = Format(vCode, "00")

        ' 入力区分名取得
        Me.txt_入力区分 = fncSelectName(P_対象区分コード, Me.cmb_入力区分)
    End If
    
    ' 入力区分による画面項目表示設定
    Call setDispInputItems(Me.cmb_入力区分.VALUE)

    Exit Sub
    
Exception:
    Call MSZZ024_M00("cmb_入力区分_AfterUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_入力区分_BeforeUpdate イベント
'        MODULE_ID        :cmb_入力区分_BeforeUpdate
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_入力区分_BeforeUpdate(Cancel As Integer)

    Dim strName     As String

  On Error GoTo Exception

    If Nz(Me.cmb_入力区分, "") = "" Then
        MsgBox "入力区分は必須選択です。", vbInformation, P_MSG_ERR
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If

    ' 入力区分名取得で有無チェック
    strName = fncSelectName(P_対象区分コード, Me.cmb_入力区分)
    If strName = "" Then
        MsgBox P_MSG_ERR_NODATA, vbInformation, P_MSG_ERR
        Cancel = True
        SendKeys "{ESC}"
    End If

    ' 権限が無いモノが撤去を選べないようにする 2006/09/06
    If Format(Me.cmb_入力区分.VALUE, "00") = CmReprMod.P_修繕_撤去予定 And pParamArgs.権限可否 = False Then
        MsgBox "権限がないため撤去を選ぶことはできません。", vbInformation, P_MSG_ERR
        Cancel = True
        SendKeys "{ESC}"
    End If

    ' 既に撤去済みならば撤去を選べないようにする 2006/09/06
    If Format(Me.cmb_入力区分.VALUE, "00") = CmReprMod.P_修繕_撤去予定 And pCntaUse = 9 Then
        MsgBox "既に撤去登録を行っています。" & vbCr & "撤去は複数件の登録はできません。", vbInformation, P_MSG_ERR
        Cancel = True
        SendKeys "{ESC}"
    End If

    ' 修理中は撤去を選べないようにする 2006/09/06
    If Format(Me.cmb_入力区分.VALUE, "00") = CmReprMod.P_修繕_撤去予定 And pCntaReprKbn = 1 Then
        MsgBox "修理中のため撤去出来ません。" & vbCr & "撤去をするためには全ての修理を完了として下さい。", vbInformation, P_MSG_ERR
        Cancel = True
        SendKeys "{ESC}"
    End If

    ' 既に撤去済みならば修理は選べないようにする 2006/10/12
    'If Format(Me.cmb_入力区分.VALUE, "00") = CmReprMod.P_修繕_修理 And pCntaUse = 9 Then                                                                       'DELETE 2021/08/03 N.IMAI
    If (Format(Me.cmb_入力区分.VALUE, "00") = CmReprMod.P_修繕_修理 Or Format(Me.cmb_入力区分.VALUE, "00") = CmReprMod.P_修繕_メンテ依頼) And pCntaUse = 9 Then 'INSERT 2021/08/03 N.IMAI
        MsgBox "既に撤去登録を行っています。" & vbCr & "撤去した部屋に対し修理の登録はできません。", vbInformation, P_MSG_ERR
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

Exception:
    Call MSZZ024_M00("cmb_入力区分_BeforeUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_入力区分_Change イベント
'        MODULE_ID        :cmb_入力区分_Change
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_入力区分_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_入力区分_KeyPress イベント
'        MODULE_ID        :cmb_入力区分_KeyPress
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_入力区分_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_入力区分, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :chk_修理完了_Click イベント
'        MODULE_ID        :chk_修理完了_Click
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_修理完了_Click()

  On Error GoTo Exception

    If Me.chk_修理完了.VALUE = -1 Then
        ' 修理完了チェックボックス選択時
        Me.txt_完了日.Enabled = True
        If Nz(Me.txt_完了日, "") = "" Then
            Me.txt_完了日 = Format(DATE, "YYYY/MM/DD")  ' 初期値(システム日付)セット
        End If
    Else
        ' 修理完了チェックボックス未選択時
        Me.txt_完了日.Enabled = False
        Me.txt_完了日 = ""
    End If

    Call subSetChangeFlag(True)
    Exit Sub
    
Exception:
    Call MSZZ024_M00("chk_修理完了_Click", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_修理区分_AfterUpdate イベント
'        MODULE_ID        :cmb_修理区分_AfterUpdate
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_修理区分_AfterUpdate()

    Dim vCode       As Variant
    Dim strName     As String
    Dim strBUMOC    As String
    Dim txt_新解除番号 As String
    Dim txt_変換鍵自動鍵区分 As String
  On Error GoTo Exception
  
    vCode = Nz(Me.cmb_修理区分, "")
    If vCode = "" Then
        Me.txt_修理区分 = ""
        Exit Sub
    End If
    Me.cmb_修理区分 = Format(vCode, "00")

    ' 修理区分名取得
    '▼ 2007/01/26 add tajima
    If pParamArgs.修繕対象 = CmReprMod.P_対象_ヤード Then
        strName = fncSelectName(P_ヤード修正区分コード, Me.cmb_修理区分)
    Else
        strName = fncSelectName(P_修正区分コード, Me.cmb_修理区分)
    End If
    '▲ 2007/01/26 add

    Me.txt_修理区分 = strName
    
    If pParamArgs.修繕対象 <> CmReprMod.P_対象_ヤード Then
        If pParamArgs.処理モード = 1 Then
            
            '一括のときは鍵交換を選べない INSERT 2017-11-18 K.Sato
            'If cmb_入力区分 = "02" And cmb_修理区分 = "05" And pParamArgs.画面ID = "FVS904" And Nz(txt_鍵区分, "") = "" Then                                                                   'DELETE 2021/08/03 N.IMAI
            If (cmb_入力区分 = CmReprMod.P_修繕_修理 Or cmb_入力区分 = CmReprMod.P_修繕_メンテ依頼) And cmb_修理区分 = "05" And pParamArgs.画面ID = "FVS904" And Nz(txt_鍵区分, "") = "" Then   'INSERT 2021/08/03 N.IMAI
                Me.cmb_修理区分 = Format(strReprKbn, "00")
                strName = ""
                If Nz(Me.cmb_修理区分, "") <> "" Then
                    strName = fncSelectName(P_修正区分コード, Me.cmb_修理区分)
                End If
                Me.txt_修理区分 = strName
                MsgBox "ヤード一括修繕（新規）の場合は鍵交換を選択出来ません。", vbExclamation, P_MSG_ERR
                Exit Sub
            Else
                '2015/02/27 M.HONDA INS
                '2017/11/18 K.SATO DEL ＩＦの中に移動
                'txt_変換鍵自動鍵区分 = DLookup("NAME_VALUE_TO", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & txt_鍵区分)
                'If cmb_入力区分 = "02" And cmb_修理区分 = "05" Then                                                                    'DELETE 2021/08/03 N.IMAI
                If (cmb_入力区分 = CmReprMod.P_修繕_修理 Or cmb_入力区分 = CmReprMod.P_修繕_メンテ依頼) And cmb_修理区分 = "05" Then    'INSERT 2021/08/03 N.IMAI
                    
                    txt_変換鍵自動鍵区分 = DLookup("NAME_VALUE_TO", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & txt_鍵区分)
                    
                    '鍵校の際には、自動発番を行う。
                    '発番しない条件
                    '・既に自動発番鍵となっている
                    '・ヤードが地方
                    If strMotokagi <> "15" And strTiho <> -1 And txt_変換鍵自動鍵区分 = 15 Then
                        strBUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
                        txt_新解除番号 = MSZZ070_M00(strBUMOC, Val(pParamArgs.YARDC), Val(pParamArgs.ROOMC), strMotokagi)
                    End If
                    
                    '既に自動発番されている場合には、鍵番号の変更は行わない。
                    If txt_新解除番号 <> "" Then
                        txt_新鍵 = txt_新解除番号
                        txt_内容壱 = "下記のダイヤル番号で施錠してください。"
                        txt_内容壱サブ壱 = txt_新解除番号
                    ElseIf strMotokagi = 15 Then
                        txt_新鍵 = txt_旧鍵
                        txt_内容壱 = "下記のダイヤル番号で施錠してください。"
                        txt_内容壱サブ壱 = txt_旧鍵
                    End If
                    '番号を修正できないように
                    txt_内容壱サブ壱.Enabled = False
                Else
                    txt_内容壱サブ壱.Enabled = True
                    txt_新鍵 = ""
                    txt_内容壱 = ""
                    txt_内容壱サブ壱 = ""
                    txt_内容壱サブ弐 = ""
                End If
            End If 'INSERT 2017-11-18 K.Sato
        End If
       End If
         
        '2015/02/27 M.HONDA INS
    
        '2017/11/18 K.SATO INS
        strReprKbn = Me.cmb_修理区分
    
    Exit Sub
    
Exception:
    Call MSZZ024_M00("cmb_修理区分_AfterUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_修理区分_BeforeUpdate イベント
'        MODULE_ID        :cmb_修理区分_BeforeUpdate
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_修理区分_BeforeUpdate(Cancel As Integer)

    Dim strName     As String

  On Error GoTo Exception

    ' コードが空の場合入力区分欄も空欄表示
    If fncControlCrear(Me.cmb_修理区分, Me.txt_修理区分) = False Then Exit Sub

    '  修理区分名取得
    '▼ 2007/01/26 add tajima
    If pParamArgs.修繕対象 = CmReprMod.P_対象_ヤード Then
        strName = fncSelectName(P_ヤード修正区分コード, Me.cmb_修理区分)
    Else
    '▲ 2007/01/26 add
        strName = fncSelectName(P_修正区分コード, Me.cmb_修理区分)
    End If  '2007/01/26 add
    
    If strName = "" Then
        MsgBox P_MSG_ERR_NODATA, vbExclamation, P_MSG_ERR
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub
    
Exception:
    Call MSZZ024_M00("cmb_修理区分_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_修理区分_Change イベント
'        MODULE_ID        :cmb_修理区分_Change
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_修理区分_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_修理区分_Change イベント
'        MODULE_ID        :cmb_修理区分_Change
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_修理区分_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_修理区分, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_時刻_BeforeUpdate イベント
'        MODULE_ID        :txt_時刻_BeforeUpdate
'        CREATE_DATE      :2013/06/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_時刻_BeforeUpdate(Cancel As Integer)
    Dim varVal              As Variant                                          'INSERT 2013/07/15 K.ISHIZAKA
    
    If Not fncChkTextLength(Me.txt_時刻, 6) Then
        'txt_時刻.Undo
        Cancel = True
    ElseIf Not IsNull(txt_時刻.VALUE) Then                                      'INSERT START 2013/07/15 K.ISHIZAKA
        If InStr(txt_時刻.VALUE, ":") > 0 Then
            varVal = Split(Nz(txt_時刻.VALUE), ":")
            If (UBound(varVal) > 1) Or (Val(varVal(0)) > 23) Or (Val(varVal(1)) > 59) Then
                MsgBox "内容を正しく入力してください。", vbInformation, P_MSG_ERR
                Cancel = True
            End If
        Else
            varVal = String(Len(txt_時刻.VALUE) Mod 2, "0") & txt_時刻.VALUE & String(6, "0")
            If (Val(Left(varVal, 2)) > 23) Or (Val(Mid(varVal, 3, 2)) > 59) Or (Val(Mid(varVal, 5, 2)) > 59) Then
                MsgBox "内容を正しく入力してください。", vbInformation, P_MSG_ERR
                Cancel = True
            End If
        End If
    End If                                                                      'INSERT END   2013/07/15 K.ISHIZAKA
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_時刻_Change イベント
'        MODULE_ID        :txt_時刻_Change
'        CREATE_DATE      :2013/06/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_時刻_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_時刻_KeyPress イベント
'        MODULE_ID        :txt_時刻_KeyPress
'        CREATE_DATE      :2013/06/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_時刻_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(":") Then
        With txt_時刻
            If ((CmLenb(.Text) > (6 - 1)) Or (InStr(Nz(.Text), ":") > 0)) And (.SelLength = 0) Then
                KeyAscii = 0
            End If
        End With
    Else
        Call CmKeyPress_For_Num1(Me.txt_時刻, KeyAscii, 6)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_日付_AfterUpdate イベント
'        MODULE_ID        :txt_日付_AfterUpdate
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_日付_AfterUpdate()
    Me.txt_日付 = Format(pvarChangeDate, "YYYY/MM/DD")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_日付_BeforeUpdate イベント
'        MODULE_ID        :txt_日付_BeforeUpdate
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_日付_BeforeUpdate(Cancel As Integer)
    Call subChkDate(Me.txt_日付, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_日付_Change イベント
'        MODULE_ID        :txt_日付_Change
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_日付_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_日付_GotFocus イベント
'        MODULE_ID        :txt_日付_GotFocus
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_日付_GotFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_日付, "")

    If vCode <> "" Then
        Me.txt_日付 = Format(vCode, "yy/mm/dd")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_日付_KeyPress イベント
'        MODULE_ID        :txt_日付_KeyPress
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_日付_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_日付, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_日付_LostFocus イベント
'        MODULE_ID        :txt_日付_LostFocus
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_日付_LostFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_日付, "")

    If vCode <> "" Then
        Me.txt_日付 = Format(vCode, "YYYY/MM/DD")
    End If
End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_対応期限日_AfterUpdate イベント
'        MODULE_ID        :txt_対応期限日_AfterUpdate
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_対応期限日_AfterUpdate()
    Me.txt_対応期限日 = Format(pvarChangeDate, "YYYY/MM/DD")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_対応期限日_BeforeUpdate イベント
'        MODULE_ID        :txt_対応期限日_BeforeUpdate
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_対応期限日_BeforeUpdate(Cancel As Integer)
    Call subChkDate(Me.txt_対応期限日, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_対応期限日_Change イベント
'        MODULE_ID        :txt_対応期限日_Change
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_対応期限日_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_対応期限日_GotFocus イベント
'        MODULE_ID        :txt_対応期限日_GotFocus
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_対応期限日_GotFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_対応期限日, "")

    If vCode <> "" Then
        Me.txt_対応期限日 = Format(vCode, "yy/mm/dd")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_対応期限日_KeyPress イベント
'        MODULE_ID        :txt_対応期限日_KeyPress
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_対応期限日_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_対応期限日, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_対応期限日_LostFocus イベント
'        MODULE_ID        :txt_対応期限日_LostFocus
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_対応期限日_LostFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_対応期限日, "")

    If vCode <> "" Then
        Me.txt_対応期限日 = Format(vCode, "YYYY/MM/DD")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_完了日_AfterUpdate イベント
'        MODULE_ID        :txt_完了日_AfterUpdate
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_完了日_AfterUpdate()
    Me.txt_完了日 = Format(pvarChangeDate, "YYYY/MM/DD")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_完了日_BeforeUpdate イベント
'        MODULE_ID        :txt_完了日_BeforeUpdate
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_完了日_BeforeUpdate(Cancel As Integer)
    Call subChkDate(Me.txt_完了日, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_完了日_Change イベント
'        MODULE_ID        :txt_完了日_Change
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_完了日_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_完了日_GotFocus イベント
'        MODULE_ID        :txt_完了日_GotFocus
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_完了日_GotFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_完了日, "")

    If vCode <> "" Then
        Me.txt_完了日 = Format(vCode, "yy/mm/dd")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_完了日_KeyPress イベント
'        MODULE_ID        :txt_完了日_KeyPress
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_完了日_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_完了日, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_完了日_LostFocus イベント
'        MODULE_ID        :txt_完了日_LostFocus
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_完了日_LostFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_完了日, "")

    If vCode <> "" Then
        Me.txt_完了日 = Format(vCode, "YYYY/MM/DD")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_金額_BeforeUpdate イベント
'        MODULE_ID        :txt_金額_BeforeUpdate
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_金額_BeforeUpdate(Cancel As Integer)
    If fncChkTextLength(Me.txt_金額, 9) = False Then
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_金額_Change イベント
'        MODULE_ID        :txt_金額_Change
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_金額_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_金額_GotFocus イベント
'        MODULE_ID        :txt_金額_GotFocus
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_金額_GotFocus()
    Me.txt_金額.Format = "Fixed"
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_金額_KeyPress イベント
'        MODULE_ID        :txt_金額_KeyPress
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_金額_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_金額, KeyAscii, 9)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_金額_LostFocus イベント
'        MODULE_ID        :txt_金額_LostFocus
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_金額_LostFocus()
    Me.txt_金額.Format = "Standard"
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_内容壱_BeforeUpdate イベント
'        MODULE_ID        :txt_内容壱_BeforeUpdate
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容壱_BeforeUpdate(Cancel As Integer)
    If fncChkTextLength(Me.txt_内容壱, 80) = False Then
        Cancel = True
''        SendKeys "{ESC}"      '----20091203----M.RYU----del----'修正依頼：桁数オーバーの場合、入力内容をそのまま保留
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_内容壱_Change イベント
'        MODULE_ID        :txt_内容壱_Change
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容壱_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_内容壱_KeyPress
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容壱_KeyPress(KeyAscii As Integer)
    Call CmKeyPress.CmKeyPress_For_Char2(txt_内容壱, KeyAscii, 80)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_内容弐_BeforeUpdate イベント
'        MODULE_ID        :txt_内容弐_BeforeUpdate
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容弐_BeforeUpdate(Cancel As Integer)
    If fncChkTextLength(Me.txt_内容弐, 80) = False Then
        Cancel = True
''        SendKeys "{ESC}"      '----20091203----M.RYU----del----'修正依頼：桁数オーバーの場合、入力内容をそのまま保留
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_内容弐_Change イベント
'        MODULE_ID        :txt_内容弐_Change
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容弐_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_内容弐_KeyPress
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容弐_KeyPress(KeyAscii As Integer)
    Call CmKeyPress.CmKeyPress_For_Char2(txt_内容弐, KeyAscii, 80)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_作業者_BeforeUpdate イベント
'        MODULE_ID        :txt_作業者_BeforeUpdate
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_作業者_BeforeUpdate(Cancel As Integer)
    If fncChkTextLength(Me.txt_作業者, 60) = False Then
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_作業者_Change イベント
'        MODULE_ID        :txt_作業者_Change
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_作業者_Change()
    Call subSetChangeFlag(True)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_作業者_KeyPress
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_作業者_KeyPress(KeyAscii As Integer)
    Call CmKeyPress.CmKeyPress_For_Char2(txt_作業者, KeyAscii, 60)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_作業者_AfterUpdate イベント
'        MODULE_ID        :cmb_作業者_AfterUpdate
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_作業者_AfterUpdate()

    Dim vCode   As Variant
    Dim vTanto  As Variant

  On Error GoTo Exception

    vCode = Nz(Me.cmb_作業者, "")
    If vCode = "" Then Exit Sub

    ' ファイルへの存在チェック
    If fncGetTantoData(vCode, vTanto) = False Then Exit Sub
    
    If Me.cmb_作業者 = "999" Then
        ' 業者コード = "999"のとき
        ' 業者名欄を入力可能に設定
        Me.txt_作業者.Enabled = True
        Me.txt_作業者.Locked = False
        Me.txt_作業者.BackColor = P_COLOR_WHITE  ' 背景色(白色)
        Me.txt_作業者 = ""
        Me.cmb_作業者 = Format$(Me.cmb_作業者, "000")
    Else
        ' 業者コード != "999"のとき
        ' 業者名欄を入力不可に設定
        Me.txt_作業者.Enabled = False
        Me.txt_作業者.Locked = True
        Me.txt_作業者.BackColor = P_COLOR_YELLOW  ' 背景色(黄色)
        Me.txt_作業者 = vTanto
        Me.cmb_作業者 = Format$(Me.cmb_作業者, "000")
    End If
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmb_作業者_AfterUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_作業者_BeforeUpdate イベント
'        MODULE_ID        :cmb_作業者_BeforeUpdate
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_作業者_BeforeUpdate(Cancel As Integer)

  On Error GoTo Exception
  
    'ｺｰﾄﾞが空の場合担当者名欄も空欄表示
    If fncControlCrear(Me.cmb_作業者, Me.txt_作業者) = False Then Exit Sub

    '選択された担当者データの存在チェック
    If fncGetTantoData(Nz(Me.cmb_作業者, "")) = False Then
       MsgBox "登録されていません。", vbInformation, P_MSG_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmb_作業者_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_作業者_Change イベント
'        MODULE_ID        :cmb_作業者_Change
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_作業者_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_作業者_KeyPress イベント
'        MODULE_ID        :cmb_作業者_KeyPress
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_作業者_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_作業者, KeyAscii, 3)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_入力担当_AfterUpdate イベント
'        MODULE_ID        :cmb_入力担当_AfterUpdate
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_入力担当_AfterUpdate()

    Dim vCode   As Variant
    Dim vTanto  As Variant
    
  On Error GoTo Exception

    vCode = Nz(Me.cmb_入力担当, "")
    If vCode = "" Then Exit Sub

    ' ファイルへの存在チェック
    If fncGetTantoData(vCode, vTanto) = False Then Exit Sub
    Me.txt_入力担当 = vTanto
    Me.cmb_入力担当 = Format$(Me.cmb_入力担当, "000")

    Exit Sub

Exception:
    Call MSZZ024_M00("cmb_入力担当_AfterUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_入力担当_BeforeUpdate イベント
'        MODULE_ID        :cmb_入力担当_BeforeUpdate
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_入力担当_BeforeUpdate(Cancel As Integer)

  On Error GoTo Exception
  
    'ｺｰﾄﾞが空の場合担当者名欄も空欄表示
    If fncControlCrear(Me.cmb_入力担当, Me.txt_入力担当) = False Then Exit Sub

    '選択された担当者データの存在チェック
    If fncGetTantoData(Nz(Me.cmb_入力担当, "")) = False Then
       MsgBox "登録されていません。", vbInformation, P_MSG_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If

    Exit Sub
    
Exception:
    Call MSZZ024_M00("cmb_入力担当_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_入力担当_Change イベント
'        MODULE_ID        :cmb_入力担当_Change
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_入力担当_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_入力担当_KeyPress イベント
'        MODULE_ID        :cmb_入力担当_KeyPress
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_入力担当_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_入力担当, KeyAscii, 3)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_保存_Click イベント
'        MODULE_ID        :cmd_保存_Click
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :2006/09/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_Click()

  On Error GoTo Exception

    If func保存処理() = False Then Exit Sub ' 2006/09/06 保存処理を外出化
    
    ' 保存したら戻る
    Call cmd_戻る_Click

    Exit Sub
    
Exception:
    Call MSZZ024_M00("cmd_保存_Click", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_戻る_Click イベント
'        MODULE_ID        :cmd_戻る_Click
'        CREATE_DATE      :2006/04/22
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_戻る_Click()

    ' 変更フラグがTrueの時のみ更新の確認メッセージを表示する
    If pIs修正 = True Then
        If MsgBox("修正中ですが保存せずに前画面に戻りますか？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, P_MSG_KAKUNIN) = vbNo Then
            Exit Sub
        End If
    End If

    ' 本画面を閉じる
    doCmd.Close acForm, Me.FormName

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Unload イベント
'        MODULE_ID        :Form_Unload
'        CREATE_DATE      :2006/05/27
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)
'    Call subParentReload TODO ？いらんだろ？
End Sub

'==============================================================================*
'
'        MODULE_NAME      :呼び出し元画面のリロード処理
'        MODULE_ID        :subParentReload
'        CREATE_DATE      :2006/05/27
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subParentReload()

     Dim objForm         As Form

    On Error Resume Next

    If blnUpdFlg = False Then Exit Sub  ' 変更してなければリロード不要

    If pParamArgs.画面ID <> "" Then
        Set objForm = Forms(pParamArgs.画面ID)
        If Not objForm Is Nothing Then
            Call objForm.OutSideReload
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncControlCrear
'        機能             :ｺﾝﾎﾞﾎﾞｯｸｽの内容が空の場合、名称テキストも空にする
'        PARAMETER        :IN  ctlCmb = コンボボックス
'                          IN  ctlTxt = テキストボックス
'        RETURN           :TRUE = 入力項目がNULL以外／FALSE = 入力項目がNULL
'        CREATE           :2006/04/22
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncControlCrear(ByRef ctlCmb As ComboBox, ByRef ctlTxt As TextBox) As Boolean

    If Nz(ctlCmb, "") = "" Then
        ctlTxt = ""
        Exit Function
    Else
        fncControlCrear = True
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :各入力項目修正時、フォームの標題を『修正中』に設定
'        MODULE_ID        :subSetChangeFlag
'        PARAMETER        :blnFlg(TRUE = 変更あり／FALSE = 変更なし)
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetChangeFlag(ByVal blnFLG As Boolean)

    Select Case blnFLG
        Case True
            ' *** 変更あり ***
            If pIs修正 = False Then
                Me.Caption = p画面標題 & "(修正中)"
                pIs修正 = True  ' 変更フラグ(ON)
            End If

        Case False
            ' *** 変更なし ***
            Me.Caption = p画面標題
            pIs修正 = False     ' 変更フラグ(OFF)
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :画面項目設定処理
'        MODULE_ID        :fncSetFrmItem
'        PARAMETER        :Nothing
'        RETURN           :TRUE = 正常終了／FALSE = 異常終了
'        CREATE_DATE      :2006/04/22
'        UPDATE_DATE      :2017/11/18
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSetFrmItem() As Boolean

    Dim strSQL    As String
    Dim iFirst    As Integer '2017-11-11 insert k.sato

    On Error GoTo Exception

    fncSetFrmItem = False

    '①各ＤＢへの接続
    ' コンテナDB接続
    Call subConnectDB(P_DB_KONT, pstrBumonCd)
    ' 加瀬DB接続
    Call subConnectDB(P_DB_KASE, "")

    '　入力区分コンボの設定
    Call set入力区分コンボ(pobjKASE_DB)

    Select Case pParamArgs.修繕対象
        ' ③修繕系の読込
        Case CmReprMod.P_対象_ヤード, CmReprMod.P_対象_部屋
            '③－α.プログラムパラメータを元に部門ＤＢ毎の情報を取得し表示
            Call sub修繕対象文言画面設定
            
            ' 修正（参照）モードならば修繕データを取得して画面に表示
            If pParamArgs.処理モード = P_修正モード Or pParamArgs.処理モード = P_参照モード Then
            
                '③－β.修繕情報取得
                ' 修繕トラン構造体初期化
                Call CmReprMod.InitReprTranType(pstReprTran)
                
                '③－γ.修繕情報ゲットだぜ
                
                '2017-11-11 k.sato insert start
                If pParamArgs.一括 Then
                    iFirst = FVS904_sub.getOnDataFirst
                    pParamArgs.UNIQEC = FVS904_sub.pSelectArgs(iFirst).UNIQEC
                End If
                '2017-11-11 k.sato insert end
                
                pstReprTran.UNIQEC = pParamArgs.UNIQEC
                If CmReprMod.ReadReprTran(pobjKASE_DB, pstReprTran) = False Then
                    ' 読み込む修繕が無い！
                    GoTo fncSetFrmItem_Exit
                End If
                '③－δ.読み込んだ修繕内容を画面に設定する
                Call sub修繕情報画面設定
            
            End If ' 新規ならば何もせん
            
        ' ④顧客系の読込
        Case CmReprMod.P_対象_顧客
            '④－α.プログラムパラメータを元に部門ＤＢ毎の情報を取得し表示
            Call sub顧客対象文言画面設定

            ' 修正（参照）モードならば修繕データを取得して画面に表示
            If pParamArgs.処理モード = P_修正モード Or pParamArgs.処理モード = P_参照モード Then
            
                '④－β.修繕情報取得
                ' 顧客問合トラン構造体初期化
                Call CmReprMod.InitKotoTranType(pstKotoTran)
    
                '④－γ.修繕情報ゲットだぜ
                pstKotoTran.UNIQEC = pParamArgs.UNIQEC
                If CmReprMod.ReadKotoTran(pobjKASE_DB, pstKotoTran) = False Then
                    ' 読み込む修繕が無い！
                    GoTo fncSetFrmItem_Exit
                End If
                '④－δ.読み込んだ修繕内容を画面に設定する
                Call sub顧客情報画面設定

            End If ' 新規ならば何もせん

    End Select

    ' 新規モード時の共通処理
    If pParamArgs.処理モード = P_新規モード Then
        Call setDispInputItems("")
        Me.cmb_入力区分.Enabled = True
        Me.txt_日付 = Format(DATE, "YYYY/MM/DD")
        If txt_時刻.Visible Then                                                'INSERT 2013/07/29 K.ISHIZAKA
            Me.txt_時刻 = Format(time, "HH:MM")                                 'INSERT 2013/07/29 K.ISHIZAKA
        End If                                                                  'INSERT 2013/07/29 K.ISHIZAKA
    ' 修正モード時の共通処理
    ElseIf pParamArgs.処理モード = P_修正モード Then
        Me.cmb_入力区分.Enabled = False '入力区分は変更できない
    End If

    ' DB切断
    Call subDisConnectDB(P_DB_KONT)
    Call subDisConnectDB(P_DB_KASE)

    ' 権限により入力可否を設定する
    If pParamArgs.処理モード = P_参照モード Then
      Call sub画面入力可否設定(False)
    End If
    
    fncSetFrmItem = True

fncSetFrmItem_Exit:
    If Not pobjKONT_DB Is Nothing Then: Call subDisConnectDB(P_DB_KONT)
    If Not pobjKASE_DB Is Nothing Then: Call subDisConnectDB(P_DB_KASE)

    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncSetFrmItem" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
  
End Function

'==============================================================================*
'
'        MODULE_NAME      :日付入力ﾃﾞｰﾀのチェック
'        MODULE_ID        :subChkDate
'        PARAMETER        :IN   ctlTxt = 処理対象テキスト
'                          OUT  Cancel = 状態区分
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subChkDate(ByVal txtTarget As Variant, ByRef Cancel As Integer)

    pvarChangeDate = CmDateChecker(txtTarget)
    If IsNull(pvarChangeDate) Then
        MsgBox P_MSG_ERR_DATE, vbInformation, P_MSG_ERR
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :文字列型→日付型書式に変換
'        MODULE_ID        :fncConvertYMD
'        PARAMETER        :IN strYMD    = "YYYYMMDD"
'        RETURN           :日付型文字列("YYYY/MM/DD")
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncConvertYMD(ByVal strYMD As String) As String

    Dim strYMD_WK As String

    On Error GoTo fncConvertYMD_Err

    ' 桁数チェック
    If Len(strYMD) > 8 Or Len(strYMD) < 8 Then
        fncConvertYMD = "": Exit Function
    End If

    strYMD_WK = Mid(strYMD, 1, 4) & "/" & Mid(strYMD, 5, 2) & "/" & Mid(strYMD, 7, 2)
    If IsDate(strYMD_WK) = False Then
        fncConvertYMD = "": Exit Function
    End If

    fncConvertYMD = strYMD_WK

fncConvertYMD_Exit:
    Exit Function

fncConvertYMD_Err:
    fncConvertYMD = ""
End Function

'==============================================================================*
'
'        MODULE_NAME      :Textエリア入力文字数チェック
'        MODULE_ID        :fncChkTextLength
'        PARAMETER        :IN ctlTxt    = 処理対象Text
'                          IN lngLength = 制限文字数
'        RETURN           :TRUE = チェックOK／FALSE：チェックNG
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkTextLength(ByRef ctlTxt As TextBox, ByVal lngLength As Long) As Boolean

    Dim strwork As String

  On Error GoTo Exception

    fncChkTextLength = False
    strwork = Nz(ctlTxt, "")

    Select Case fncChkChar(strwork, lngLength)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_MSG_ERR
            Exit Function
        Case 9
            MsgBox "内容を正しく入力してください。", vbInformation, P_MSG_ERR
            Exit Function
    End Select

    fncChkTextLength = True
    Exit Function

Exception:
    Call MSZZ024_M00("fncChkTextLength", True)
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :文字列のチェック
'        MODULE_ID        :fncChkChar
'        PARAMETER        :IN  oText = 処理対象Text
'                          IN  iLen  = 制限文字数
'        RETURN           :0 = チェックOK／2 = 桁溢れ／9 = 左記以外のエラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkChar(ByVal oText As Variant, ByVal ilen As Integer) As Integer

  On Error GoTo Exception

    fncChkChar = 9

    If CmLenb(oText) > ilen Then
        fncChkChar = 2
        Exit Function
    End If

    fncChkChar = 0
    Exit Function

Exception:
    fncChkChar = 9
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :データ保存時の入力項目チェック
'        MODULE_ID        :fncChkInput
'        PARAMETER        :なし
'        RETURN           :True = 入力OK／False = 入力NG
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkInput() As Boolean

    Dim strErrItem As String
    Dim intErrItem As Integer

  On Error GoTo Exception

    fncChkInput = False

    intErrItem = 0

    ' 各種入力チェック

    ' 入力区分
    If Nz(Me.cmb_入力区分, "") = "" Then
        intErrItem = 1: GoTo fncChkInput_INPUT_Err
    End If

    ' 修理の場合区分  2007/01/26 fix tajima
    'If Me.cmb_入力区分 = P_修繕_修理 Then                                          'DELETE 2021/08/03 N.IMAI
    If Me.cmb_入力区分 = P_修繕_修理 Or Me.cmb_入力区分 = P_修繕_メンテ依頼 Then    'INSERT 2021/08/03 N.IMAI
        ' 「入力区分=修正」のとき修理区分必須チェック
        If Nz(Me.cmb_修理区分, "") = "" Then
            intErrItem = 2: GoTo fncChkInput_INPUT_Err
        End If
        
    End If

    ' 日付
    If Nz(Me.txt_日付, "") = "" Then
        intErrItem = 3: GoTo fncChkInput_INPUT_Err
    End If

    ' 完了日
    If (Me.chk_修理完了.VALUE = -1 And Nz(Me.txt_完了日, "") = "") Then
        ' 「修理完了ON」かつ「完了日 = ""」のとき
        intErrItem = 4: GoTo fncChkInput_INPUT_Err
    End If

    ' 入力担当
    If Nz(Me.cmb_入力担当, "") = "" Then
        intErrItem = 5: GoTo fncChkInput_INPUT_Err
    End If

    fncChkInput = True

fncChkInput_Exit:
    Exit Function

' 入力エラー処理
fncChkInput_INPUT_Err:
    With Me
        Select Case intErrItem
            Case 1
                ' 入力区分
                strErrItem = "入力区分"
                .cmb_入力区分.SetFocus

            Case 2
                ' 修理区分
                strErrItem = "修理区分"
                .cmb_修理区分.SetFocus

            Case 3
                ' 日付
                strErrItem = "日付"
                .txt_日付.SetFocus
    
            Case 4
                ' 完了日
                strErrItem = "完了日"
                .txt_完了日.SetFocus

            Case 5
                ' 入力担当
                strErrItem = "入力担当"
                .cmb_入力担当.SetFocus
        End Select
    End With

    ' エラーメッセージ表示
    MsgBox strErrItem & P_MSG_ERR_INPUT, vbExclamation, P_MSG_ERR

    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncChkInput" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :名称取得
'        MODULE_ID        :fncSelectName
'        PARAMETER        :IN  strID     = 選択コード
'                          IN  ObjContNM = コントロール
'        RETURN           :入力区分名または修正区分名
'        CREATE_DATE      :2006/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectName(ByVal strID As String, ByVal ObjContNM As Object)

    Dim vCode    As Variant
    Dim strWhere As String
    Dim strName  As String

  On Error GoTo Exception
 
    ' 選択区分名取得
    vCode = Nz(ObjContNM, "")

    ' コードテーブルより名称を取得
    strWhere = ""
    strWhere = strWhere & " CODET_SIKBC = '" & strID & "'"
    strWhere = strWhere & " AND CODET_CODEC = " & CInt(vCode)

    strName = Nz(DLookup("CODET_NAMEN", "CODE_TABL", strWhere), "")

    fncSelectName = strName
    Exit Function
    
Exception:
    Call MSZZ024_M00("fncSelectName", True)
End Function

'==============================================================================*
'
'        MODULE_NAME      :GetTantoData
'        機能             :担当者マスタ情報取得
'        PARAMETER        :IN   vTanto = 担当者コード
'                          OUT  vTanNm = 担当者名
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetTantoData(ByVal vTanto As Variant, Optional ByRef vTanNm As Variant) As Boolean

    Dim strSQL  As String
    Dim objRs   As Recordset
    Dim isException As Boolean

  On Error GoTo Exception
    isException = False
    fncGetTantoData = False

    strSQL = "SELECT "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTC, "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTN "
    strSQL = strSQL & "FROM TANT_MAST INNER JOIN "
    strSQL = strSQL & "dbo_CONT_MAST ON "
    strSQL = strSQL & "TANT_MAST.TANTM_BUMOC = dbo_CONT_MAST.CONT_BUMOC "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTC = " & "'" & Format$(vTanto, "000") & "'"
    Set objRs = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)

    If objRs.EOF Then
        objRs.Close
        Set objRs = Nothing
        Exit Function
    Else
        If Nz(objRs("TANTM_TANTN"), "") <> "" Then vTanNm = objRs("TANTM_TANTN")
    End If
    
    fncGetTantoData = True
    GoTo Finally

Exception:
    isException = True

Finally:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If isException = True Then
        Call MSZZ024_M00("fncGetTantoData", True)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :部門コード/部門名称の取得
'        MODULE_ID        :subGetBumonName
'        PARAMETER        :なし
'        CREATE_DATE      :2006/04/22
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subGetBumonName()

    Dim strSQL  As String
    Dim objRs   As Recordset
    Dim isException As Boolean
    
  On Error GoTo Exception
    isException = False

    strSQL = "SELECT BUMOM_BUMOC "  ' 部門コード
    strSQL = strSQL & "FROM  dbo_CONT_MAST "
    strSQL = strSQL & "     ,BUMO_MAST "
    strSQL = strSQL & " WHERE CONT_BUMOC = BUMOM_BUMOC "
    Set objRs = pobjDB.OpenRecordset(strSQL, dbReadOnly)

    With objRs
        If Not .EOF Then
            pstrBumonCd = .Fields("BUMOM_BUMOC")
        Else
            pstrBumonCd = Null
        End If
    End With
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If isException = True Then
        Call MSZZ024_M00("subGetBumonName", True)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :データベースへの登録処理
'        MODULE_ID        :fncDB登録処理
'        PARAMETER        :Nothing
'        RETURN           :True = 正常終了／False = 異常終了
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDB登録処理() As Boolean

    Dim WRI_ObjWS     As Workspace
    Dim strSQL        As String
    Dim isException   As Boolean
    Dim is修理中      As Boolean
    Dim errorMessage  As String
        
  On Error GoTo Exception

    fncDB登録処理 = False
    isException = False

    ' 加瀬DB接続
    Call subConnectDB(P_DB_KASE, "")
    ' 2006/09/06 KOMS系接続 部屋修繕のときのみ開く
    If pParamArgs.修繕対象 = P_対象_部屋 Then Call subConnectDB(P_DB_KONT, pstrBumonCd)
    
    ' トランザクション開始
    Set WRI_ObjWS = DBEngine.Workspaces(0)
    WRI_ObjWS.BeginTrans

    Select Case pParamArgs.修繕対象
        Case P_対象_ヤード, P_対象_部屋
            Call subSet画面項目to修繕構造体(CStr(pParamArgs.処理モード))
            If pParamArgs.処理モード = P_新規モード Then
                is修理中 = CmReprMod.IsRepairing(pobjKASE_DB, pstrBumonCd, pstReprTran.YARDC, pstReprTran.ROOMC)
                ' 修繕トラン１件作成
                pParamArgs.UNIQEC = CmReprMod.InsertReprTran(pobjKASE_DB, pstReprTran)
                
                '2015/02/27 M.HONDA INS
                '修理完了のチェックがついている場合のみ、自動発番トランクを作成
                If chk_修理完了.VALUE = -1 Then
                    'If cmb_入力区分 = "02" And cmb_修理区分 = "05" And txt_新鍵 <> "" Then                                                                     'DELETE 2021/08/03 N.IMAI
                    If (cmb_入力区分 = CmReprMod.P_修繕_修理 Or cmb_入力区分 = CmReprMod.P_修繕_メンテ依頼) And cmb_修理区分 = "05" And txt_新鍵 <> "" Then     'INSERT 2021/08/03 N.IMAI
                        If Not SetAkknTran() Then
                            GoTo Finally
                        End If
                    End If
                End If
                '2015/02/27 M.HONDA INS
                
            Else
                is修理中 = CmReprMod.IsRepairing(pobjKASE_DB, pstrBumonCd, pstReprTran.YARDC, pstReprTran.ROOMC, pstReprTran.UNIQEC)
                ' 修繕トラン更新
                Call CmReprMod.UpdateReprTran(pobjKASE_DB, pstReprTran)
                
                '2015/02/27 M.HONDA INS
                '修理完了のチェックがついている場合のみ、自動発番トランクを作成
                If chk_修理完了.VALUE = -1 Then
                    'If cmb_入力区分 = "02" And cmb_修理区分 = "05" And txt_新鍵 <> "" Then                                                                     'DELETE 2021/08/03 N.IMAI
                    If (cmb_入力区分 = CmReprMod.P_修繕_修理 Or cmb_入力区分 = CmReprMod.P_修繕_メンテ依頼) And cmb_修理区分 = "05" And txt_新鍵 <> "" Then     'INSERT 2021/08/03 N.IMAI
                        If Not SetAkknTran() Then
                            GoTo Finally
                        End If
                    End If
                End If
                '2015/02/27 M.HONDA INS
                
            End If
            
            '2006/09/06 部屋修繕の場合、他ＤＢへの反映処理
            If pParamArgs.修繕対象 = P_対象_部屋 Then
                Call subKont系反映(is修理中)
            End If
                
        Case P_対象_顧客
            Call subSet画面項目to顧客問合構造体(CStr(pParamArgs.処理モード))
            If pParamArgs.処理モード = P_新規モード Then
                ' 顧客問合せ１件作成
                Call CmReprMod.InsertKotoTran(pobjKASE_DB, pstKotoTran)
            Else
                ' 顧客問合わせ更新
                Call CmReprMod.UpdateKotoTran(pobjKASE_DB, pstKotoTran)
            End If

    End Select

    ' コミット
    WRI_ObjWS.CommitTrans

    fncDB登録処理 = True
    GoTo Finally

Exception:
    ' ロールバック
    WRI_ObjWS.Rollback
    isException = True

Finally:
    ' 各DB切断
    If Not pobjKASE_DB Is Nothing Then: Call subDisConnectDB(P_DB_KASE)
    If Not pobjKONT_DB Is Nothing Then: Call subDisConnectDB(P_DB_KONT)
    If isException = True Then
        Call Err.Raise(Err.Number, "fncDB登録処理" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :SetAkknTran
'        機能             :自動鍵解除番号発番トラン登録処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SetAkknTran() As Boolean
    Dim strSQL          As String
    Dim isException     As Boolean
    Dim rsAkkn          As DAO.Recordset    ' 自動鍵解除番号発番トランレコードセット
   

On Error GoTo SetAkknTran_Err

    SetAkknTran = False
    isException = False

    strSQL = ""
    strSQL = strSQL & "SELECT      *" & vbCrLf
    strSQL = strSQL & "FROM        AKKN_TRAN" & vbCrLf
    strSQL = strSQL & "WHERE       AKKNT_YARD = " & Val(pstReprTran.YARDC) & vbCrLf
    strSQL = strSQL & "AND         AKKNT_CTNO = " & Val(pstReprTran.ROOMC) & vbCrLf
    strSQL = strSQL & "AND         AKKNT_KTYPE = '15' " & vbCrLf

    Set rsAkkn = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

    '自動鍵解除番号発番トランにデータ登録
    With rsAkkn
        .AddNew
        .Fields("AKKNT_YARD") = Val(pstReprTran.YARDC)                                      ' ヤードコード
        .Fields("AKKNT_CTNO") = Val(pstReprTran.ROOMC)                                      ' コンテナ番号
        .Fields("AKKNT_KTYPE") = 15                                                         ' 鍵区分
        .Fields("AKKNT_STYPE") = Len(Nz(txt_新鍵))                                          ' 副種別 INSERT 2021/08/03 N.IMAI
        .Fields("AKKNT_HATUD") = Format$(DATE, "YYYYMMDD") & Format$(time, "HHMMSS")        ' 解除番号発番日時
        .Fields("AKKNT_KAINO") = Nz(txt_新鍵)                                               ' 解除番号
        .Fields("AKKNT_INSED") = Format$(DATE, "YYYYMMDD")                                  ' 作成日付
        .Fields("AKKNT_INSEJ") = Format$(time, "HHMMSS")                                    ' 作成時刻
        .Fields("AKKNT_INSPB") = "FVS910"                                                   ' 作成プログラムID
        .Fields("AKKNT_INSUB") = LsGetUserName()                                            ' 作成ユーザーID
        .UPDATE
    End With

    SetAkknTran = True

    GoTo SetAkknTran_End

SetAkknTran_Err:
    isException = True

SetAkknTran_End:
    If Not rsAkkn Is Nothing Then rsAkkn.Close: Set rsAkkn = Nothing

    If isException = True Then
        Call Err.Raise(Err.Number, "SetAkknTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :DB接続
'        MODULE_ID        :subConnectDB
'        PARAMETER        :intKbn
'                            1 = コンテナDB
'                            2 = 加瀬DB
'                         :strBumoCD  部門コード
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subConnectDB(ByVal intKbn As Integer, ByVal strBumoCD As String)

    Dim strConnect    As String
    Dim strDataSource As String

    On Error GoTo Exception

    Select Case intKbn
        Case P_DB_KONT
            ' -----------------------------------------------------------------------------------------------------------
            ' コンテナDB接続
            strDataSource = fncGetDataSource("ODBC_DATA_SOURCE_NAME_" & strBumoCD)
            If strDataSource = "" Then
                ' テーブル[SETU_TABL]の設定不正
                Exit Sub
            End If
        
            ' DB接続
            Set pobjKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(CStr(strBumoCD)))
            ' -----------------------------------------------------------------------------------------------------------

        Case P_DB_KASE
            ' -----------------------------------------------------------------------------------------------------------
            ' 加瀬DB接続
            strDataSource = fncGetDataSource("ODBC_DATA_SOURCE_NAME")
        
            If strDataSource = "" Then
                ' テーブル[SETU_TABL]の設定不正
                Exit Sub
            End If
        
            Set pobjKASE_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00())
            ' -----------------------------------------------------------------------------------------------------------
    End Select

    Exit Sub

Exception:
    Call MSZZ024_M00("subConnectDB", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :DB切断
'        MODULE_ID        :subDisConnectDB
'        PARAMETER        :intKbn
'                            1 = コンテナDB
'                            2 = 加瀬DB
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDisConnectDB(ByVal intKbn As Integer)

  On Error GoTo Exception

    Select Case intKbn
        Case 1
            ' コンテナDB
            pobjKONT_DB.Close
            Set pobjKONT_DB = Nothing

        Case 2
            ' 加瀬DB
            pobjKASE_DB.Close
            Set pobjKASE_DB = Nothing
    End Select

    Exit Sub

Exception:
    Call MSZZ024_M00("subDisConnectDB", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ODBC接続文字列取得
'        MODULE_ID        :fncGetDataSource
'        Parameter        :strODBCNAME ODBC名
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetDataSource(ByVal strODBCNAME As String) As String

    Dim strName As String

  On Error GoTo Exception

    fncGetDataSource = ""

    strName = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = '" & strODBCNAME & "'"))

    fncGetDataSource = strName

    Exit Function

Exception:
    Call MSZZ024_M00("fncGetDataSource", True)

End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :sub修繕対象文言画面設定
'        PARAMETER        :Nothing
'        CREATE_DATE      :2006/05/19
'        NOTE             :前提としてプログラムパラメータ"SHUZEN"取得済みのこと
'        UPDATE_DATE      :2017/11/18
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub sub修繕対象文言画面設定()

    Dim strSQL As String
    Dim blnError As Boolean
    Dim rsObject As Recordset
    Dim strKagi As String
  
  On Error GoTo Exception
    blnError = False
  
    ' Select句
    strSQL = "SELECT " & Chr(13)
    strSQL = strSQL & "  YARD_NAME " & Chr(13)
    strSQL = strSQL & " ,isnull(YARD_TIHO_KBN,0) YARD_TIHO_KBN " & Chr(13)
    If pParamArgs.修繕対象 = P_対象_部屋 Then
      ' コンテナならばSelect句にコンテナ項目
      strSQL = strSQL & "  ,CNTA_SIZE " & Chr(13)
      strSQL = strSQL & "  ,STEPCD_NAME.NAME_NAME AS STEP_NAME " & Chr(13)
      strSQL = strSQL & "  ,CNTA_USE,CNTA_REPR_KBN " & Chr(13)
      strSQL = strSQL & "  ,CAST( CNTA_MAST.CNTA_KEY_TYPE AS INT ) AS CNTA_KEY_TYPE " & Chr(13)
      strSQL = strSQL & "  ,KAGI_NAME.NAME_NAME AS KAGI_NAME " & Chr(13)
      '2015/03/25 M.HONDA INS
      strSQL = strSQL & " ,isnull(( "
      strSQL = strSQL & " SELECT "
      strSQL = strSQL & "         AKKNT.AKKNT_KAINO "
      strSQL = strSQL & "     FROM "
      strSQL = strSQL & "         ( "
      strSQL = strSQL & "         SELECT "
      strSQL = strSQL & "             AKKNT_YARD, "
      strSQL = strSQL & "             AKKNT_CTNO, "
      strSQL = strSQL & "             AKKNT_KTYPE, "
      strSQL = strSQL & "             AKKNT_HATUD, "
      strSQL = strSQL & "             ROW_NUMBER() OVER (PARTITION BY AKKNT_YARD, AKKNT_CTNO, AKKNT_KTYPE ORDER BY AKKNT_HATUD DESC) AS SEQ, "
      strSQL = strSQL & "             AKKNT_KAINO "
      strSQL = strSQL & "         FROM "
      strSQL = strSQL & "             AKKN_TRAN "
      strSQL = strSQL & "         WHERE "
      strSQL = strSQL & "             AKKNT_YARD = CNTA_MAST.CNTA_CODE AND "
      strSQL = strSQL & "             AKKNT_CTNO = CNTA_MAST.CNTA_NO AND "
      strSQL = strSQL & "             AKKNT_KTYPE = CNTA_MAST.CNTA_KEY_TYPE "
      strSQL = strSQL & "         ) AKKNT "
      strSQL = strSQL & "     WHERE "
      strSQL = strSQL & "         AKKNT.SEQ = 1 "
      strSQL = strSQL & "     ) ,CNTA_KEY_NO) CNTA_KEY_NO "
      '2015/03/25 M.HONDA INS
    End If
     
    'From句
     strSQL = strSQL & " FROM YARD_MAST "
    If pParamArgs.修繕対象 = P_対象_部屋 Then
      ' コンテナならばFrom句にコンテナ項目
      strSQL = strSQL & "  INNER JOIN CNTA_MAST ON CNTA_MAST.CNTA_CODE = YARD_MAST.YARD_CODE" & Chr(13)
      strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST AS STEPCD_NAME " & Chr(13)
      strSQL = strSQL & "  ON STEPCD_NAME.NAME_CODE = CAST( CNTA_MAST.CNTA_STEP AS INT ) " & Chr(13)
      strSQL = strSQL & "  AND STEPCD_NAME.NAME_ID = '090' " & Chr(13)
      '2015/02/27 M.HONDA INS
      strSQL = strSQL & "   LEFT JOIN NAME_MAST AS KAGI_NAME " & Chr(13)
      strSQL = strSQL & "    ON KAGI_NAME.NAME_CODE = CAST( CNTA_MAST.CNTA_KEY_TYPE AS INT )  " & Chr(13)
      strSQL = strSQL & "    AND KAGI_NAME.NAME_ID = '060'  " & Chr(13)
      '2015/02/27 M.HONDA INS
    End If
  
    'Where句
     strSQL = strSQL & " WHERE YARD_CODE = " & pParamArgs.YARDC
    If pParamArgs.修繕対象 = P_対象_部屋 Then
      ' コンテナならばFrom句にコンテナ項目
      strSQL = strSQL & " AND CNTA_NO = " & pParamArgs.ROOMC
    End If
    
    Set rsObject = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    If rsObject.EOF = False Then
      '対象ラベルの文言生成
      Dim strText As String
      strText = pParamArgs.YARDC & Space(1) & rsObject.Fields("YARD_NAME")
      If pParamArgs.修繕対象 = P_対象_部屋 Then
        If Not pParamArgs.一括 Then                                                                '2017/11/11 K.SATO INS add
            '一括ではないときに部屋情報を出力                                                      '2017/11/11 K.SATO INS add
            ' コンテナならばコンテナ情報も付加
            strText = strText & " ： " & pParamArgs.ROOMC & Space(1) & rsObject.Fields("STEP_NAME")
            strText = strText & Space(1) & Format$(rsObject.Fields("CNTA_SIZE"), "0.00") & "帖"
            pCntaUse = Nz(rsObject.Fields("CNTA_USE"), 0)
            pCntaReprKbn = Nz(rsObject.Fields("CNTA_REPR_KBN"), 0)
            '2015/02/27 M.HONDA INS
            Me.txt_鍵区分 = rsObject.Fields("CNTA_KEY_TYPE")
            Me.txt_鍵名称 = rsObject.Fields("KAGI_NAME")
            strMotokagi = rsObject.Fields("CNTA_KEY_TYPE")
            Me.txt_旧鍵 = rsObject.Fields("CNTA_KEY_NO")
            '2015/02/27 M.HONDA INS
        End If                                                                                     '2017/11/11 K.SATO INS add
      Else
        pCntaUse = 0
        pCntaReprKbn = 0
      End If
      Me.txt_修繕対象 = strText
      '2015/02/27 M.HONDA INS
      strTiho = rsObject.Fields("YARD_TIHO_KBN")
      '2015/02/27 M.HONDA INS
      
    End If
     
    GoTo Finally
   
Exception:
    blnError = True
  
Finally:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
  
    If blnError Then
        Call Err.Raise(Err.Number, "sub修繕対象文言設定" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
  
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :sub顧客対象文言画面設定
'        PARAMETER        :Nothing
'        CREATE_DATE      :2006/05/19
'        NOTE             :前提としてプログラムパラメータ"SHUZEN"取得済みのこと
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub sub顧客対象文言画面設定()
  
    Dim strSQL As String
    Dim blnError As Boolean
    Dim rsObject As Recordset
    
  On Error GoTo Exception
    blnError = False
  
    ' Select文、顧客名取得
    strSQL = "SELECT USER_NAME FROM USER_MAST " & Chr(13)
    strSQL = strSQL & " WHERE USER_CODE = '" & pParamArgs.KOKYC & "'"
    
    Set rsObject = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    '対象ラベルの文言生成
    Dim strText As String
    strText = pParamArgs.KOKYC & Space(1)
    
    If rsObject.EOF = False Then
        strText = strText & rsObject.Fields("USER_NAME")
    Else
        strText = strText & "この顧客コードに該当する顧客情報はありません"
    End If
     
    Me.txt_修繕対象 = strText
    
    GoTo Finally
   
Exception:
    blnError = True
  
Finally:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
  
    If blnError Then
        Call Err.Raise(Err.Number, "sub顧客対象文言画面設定" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :sub修繕情報画面設定
'        PARAMETER        :Nothing
'        CREATE_DATE      :2006/05/19
'        NOTE             :前提として　stReprTran　に修繕データ取得済みのこと
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub sub修繕情報画面設定()

  Dim blnError As Boolean
  
  On Error GoTo Exception
  blnError = False
   
  With Me
      .cmb_入力区分 = pstReprTran.TYPEC
      Call cmb_入力区分_AfterUpdate

      .cmb_修理区分 = pstReprTran.REPRC
      Call cmb_修理区分_AfterUpdate

      .txt_日付 = IIf(Nz(pstReprTran.GENTD, "") = "", Null, _
                       fncConvertYMD(pstReprTran.GENTD))

      .txt_対応期限日 = IIf(Nz(pstReprTran.LIMTD, "") = "", Null, _
                       fncConvertYMD(Nz(pstReprTran.LIMTD)))
      
      .txt_完了日 = IIf(Nz(pstReprTran.COMPD, "") = "", Null, _
                       fncConvertYMD(Nz(pstReprTran.COMPD, "")))
      
      ' 完了日が入っていれば修理完了チェック済みとする
      .chk_修理完了.VALUE = IIf(Nz(pstReprTran.COMPD, "") = "", 0, -1)
      .txt_完了日.Enabled = IIf(Nz(pstReprTran.COMPD, "") = "", False, True)
      
      Call txt_金額_LostFocus
      .txt_金額 = pstReprTran.KINGA
      
      .txt_内容壱 = pstReprTran.CNT1N
      .txt_内容弐 = pstReprTran.CNT2N
      '▼ 2007/01/26 tajima add
      .txt_内容壱サブ壱 = pstReprTran.CNT1N1
      .txt_内容壱サブ弐 = pstReprTran.CNT1N2
      .txt_内容弐サブ壱 = pstReprTran.CNT2N1
      .txt_内容弐サブ弐 = pstReprTran.CNT2N2
      '▲ 2007/01/26 tajima add
      
      '2015/03/25 M.HONDA INS
      'If cmb_入力区分 = "02" And cmb_修理区分 = "05" Then                                                                  'DELETE 2021/08/03 N.IMAI
      If (cmb_入力区分 = CmReprMod.P_修繕_修理 Or cmb_入力区分 = CmReprMod.P_修繕_メンテ依頼) And cmb_修理区分 = "05" Then  'INSERT 2021/08/03 N.IMAI
         txt_新鍵 = txt_内容壱サブ壱
         If txt_完了日 <> "" Then
            '鍵交換で完了済みにの際には、再度データの修正はさせない。
            cmd_保存.Enabled = False
         End If
      End If
      '2015/03/25 M.HONDA INS
      
      .cmb_作業者 = pstReprTran.WORKC
      If .cmb_作業者 <> "999" Then
          Call cmb_作業者_AfterUpdate
      Else
          .txt_作業者 = pstReprTran.GYOUN
          .txt_作業者.Enabled = True                                'INSERT 2013/06/17 K.ISHIZAKA
          .txt_作業者.Locked = False                                'INSERT 2013/06/17 K.ISHIZAKA
          .txt_作業者.BackColor = P_COLOR_WHITE ' 背景色(白色)      'INSERT 2013/06/17 K.ISHIZAKA
      End If

      .cmb_入力担当 = pstReprTran.TANTC
      Call cmb_入力担当_AfterUpdate
  End With
   
  Exit Sub
   
Exception:
  Call Err.Raise(Err.Number, "sub修繕情報画面設定" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
  
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :sub顧客情報画面設定
'        PARAMETER        :Nothing
'        CREATE_DATE      :2006/05/19
'        NOTE             :前提として　stKotoTran　に修繕データ取得済みのこと
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub sub顧客情報画面設定()
  
  Dim blnError As Boolean
  
  On Error GoTo Exception
  blnError = False
   
  With Me
      .cmb_入力区分 = pstKotoTran.TYPEC
      Call cmb_入力区分_AfterUpdate

      .cmb_修理区分 = Null

      .txt_日付 = IIf(Nz(pstKotoTran.GENTD, "") = "", Null, _
                       fncConvertYMD(pstKotoTran.GENTD))
'      .txt_時刻.Value = IIf(pstKotoTran.GENTJ = "", Null, pstKotoTran.GENTJ)   'DELETE 2013/07/15 K.ISHIZAKA 'INSERT 2013/06/17 K.ISHIZAKA
        pstKotoTran.GENTJ = Replace(pstKotoTran.GENTJ, ":", "")                 'INSERT START 2013/07/15 K.ISHIZAKA
        If IsNumeric(pstKotoTran.GENTJ) Then
            If Val(pstKotoTran.GENTJ) = 0 Then
                .txt_時刻.VALUE = Null
            Else
                .txt_時刻.VALUE = Left(pstKotoTran.GENTJ, 2) & ":" & Mid(pstKotoTran.GENTJ, 3, 2)
            End If
        Else
            .txt_時刻.VALUE = pstKotoTran.GENTJ
        End If                                                                  'INSERT END   2013/07/15 K.ISHIZAKA
        
      .txt_対応期限日 = Null
      .txt_完了日 = Null
      .txt_金額 = Null
      .txt_内容壱 = pstKotoTran.CONTN
      '▼ 2007/01/26 tajima add
      .txt_内容壱サブ壱 = pstKotoTran.CNT1N
      .txt_内容壱サブ弐 = pstKotoTran.CNT2N
      .txt_内容弐サブ壱 = Null
      .txt_内容弐サブ弐 = Null
      '▲ 2007/01/26 tajima add
      .txt_内容弐 = Null
      .cmb_作業者 = Null
      .txt_作業者 = Null
      
      .cmb_入力担当 = pstKotoTran.TANTC
      Call cmb_入力担当_AfterUpdate
  End With
   
  Exit Sub
   
Exception:
  Call Err.Raise(Err.Number, "sub顧客情報画面設定" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
  
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :sub権限による画面入力可否設定
'        PARAMETER        :TRUE or FALSE
'        CREATE_DATE      :2006/05/19
'        NOTE             :前提として pIsAuthority　に権限取得済みのこと
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub sub画面入力可否設定(a更新可否 As Boolean)
   
  With Me
      .cmd_保存.Enabled = a更新可否
      .cmb_入力区分.Enabled = a更新可否
      .cmb_修理区分.Enabled = a更新可否
      .txt_日付.Enabled = a更新可否
      .txt_対応期限日.Enabled = a更新可否
      .txt_完了日.Enabled = a更新可否
      .txt_金額.Enabled = a更新可否
      .txt_内容壱.Enabled = a更新可否
      .txt_内容弐.Enabled = a更新可否
      .cmb_作業者.Enabled = a更新可否
      .txt_作業者.Enabled = a更新可否
      .cmb_入力担当.Enabled = a更新可否
      .chk_修理完了.Enabled = a更新可否
      '▼ 2007/01/26 tajima add
      .txt_内容壱サブ壱.Enabled = a更新可否
      .txt_内容壱サブ弐.Enabled = a更新可否
      .txt_内容弐サブ壱.Enabled = a更新可否
      .txt_内容弐サブ弐.Enabled = a更新可否
      '▲ 2007/01/26 tajima add
  End With
 
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subSet画面項目to修繕構造体
'        PARAMETER        :セットモード
'        CREATE_DATE      :2006/05/19
'        NOTE             :画面内容を修繕構造体 stReprTran にセットす
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSet画面項目to修繕構造体(a更新区分 As String)
          
        pstReprTran.TANTC = Me.cmb_入力担当
        pstReprTran.GENTD = Format$(Me.txt_日付, "YYYYMMDD")
        pstReprTran.LIMTD = zeroByteToNullorFormat(Me.txt_対応期限日, "YYYYMMDD")
        pstReprTran.COMPD = zeroByteToNullorFormat(Me.txt_完了日, "YYYYMMDD")
        pstReprTran.TYPEC = Me.cmb_入力区分
        pstReprTran.REPRC = zeroByteToNull(Me.cmb_修理区分)
        pstReprTran.CNT1N = zeroByteToNull(Me.txt_内容壱)
        pstReprTran.CNT2N = zeroByteToNull(Me.txt_内容弐)
        pstReprTran.KINGA = zeroByteToNull(Me.txt_金額)
        pstReprTran.WORKC = zeroByteToNull(Me.cmb_作業者)
        pstReprTran.GYOUN = zeroByteToNull(Me.txt_作業者)
        '▼ 2007/01/26 tajima add
        pstReprTran.CNT1N1 = zeroByteToNull(Me.txt_内容壱サブ壱)
        pstReprTran.CNT1N2 = zeroByteToNull(Me.txt_内容壱サブ弐)
        pstReprTran.CNT2N1 = zeroByteToNull(Me.txt_内容弐サブ壱)
        pstReprTran.CNT2N2 = zeroByteToNull(Me.txt_内容弐サブ弐)
        '▲ 2007/01/26 tajima add
        
        If a更新区分 = P_新規モード Then
          '更新の場合は既に構造体に設定済み
          pstReprTran.BUMOC = pstrBumonCd
          pstReprTran.YARDC = pParamArgs.YARDC
          pstReprTran.ROOMC = zeroByteToNull(pParamArgs.ROOMC)
          pstReprTran.INPTI = pParamArgs.修繕対象
          pstReprTran.ACPTB = zeroByteToNull(pParamArgs.契約番号)
          ' 作成情報
          pstReprTran.INSED = Format$(DATE, "YYYYMMDD")
          pstReprTran.INSEJ = Format$(time, "HHMMSS")
          pstReprTran.INSPB = Me.Form.NAME
          pstReprTran.INSUB = pstrUserName
        Else
          pstReprTran.UPDAD = Format$(DATE, "YYYYMMDD")
          pstReprTran.UPDAJ = Format$(time, "HHMMSS")
          pstReprTran.UPDPB = Me.Form.NAME
          pstReprTran.UPDUB = pstrUserName
        End If
                
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subSet画面項目to顧客問合構造体
'        PARAMETER        :セットモード
'        CREATE_DATE      :2006/05/19
'        NOTE             :画面内容を顧客問合構造体 stKotoTran にセットす
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSet画面項目to顧客問合構造体(a更新区分 As String)
    Dim varVal              As Variant                                          'INSERT 2013/07/15 K.ISHIZAKA
          
        pstKotoTran.TANTC = Me.cmb_入力担当
        pstKotoTran.GENTD = Format$(Me.txt_日付, "YYYYMMDD")
'        pstKotoTran.GENTJ = Nz(txt_時刻.Value)                                 'DELETE 2013/07/15 K.ISHIZAKA 'INSERT 2013/06/17 K.ISHIZAKA
        If IsNull(txt_時刻.VALUE) Then                                          'INSERT START 2013/07/15 K.ISHIZAKA
'            pstKotoTran.GENTJ = "000000"                                       'DELETE 2013/07/29 K.ISHIZAKA
            ''2013/09/24 M.HONDA UPD
            ''pstKotoTran.GENTJ = Format$(Time, "HHMM") & "00"                    'INSERT 2013/07/29 K.ISHIZAKA
            pstKotoTran.GENTJ = Format$(time, "HHMMSS")
            ''2013/09/24 M.HONDA UPD
        Else
            If InStr(txt_時刻.VALUE, ":") > 0 Then
                pstKotoTran.GENTJ = ""
                For Each varVal In Split(txt_時刻.VALUE, ":")
                    pstKotoTran.GENTJ = pstKotoTran.GENTJ & Format(varVal, "00")
                Next
            Else
                pstKotoTran.GENTJ = String(Len(txt_時刻.VALUE) Mod 2, "0") & txt_時刻.VALUE
            End If
            pstKotoTran.GENTJ = Left(pstKotoTran.GENTJ & "000000", 6)
        End If                                                                  'INSERT END   2013/07/15 K.ISHIZAKA
        pstKotoTran.TYPEC = Me.cmb_入力区分
        pstKotoTran.CONTN = zeroByteToNull(Me.txt_内容壱)
        '▼ 2007/01/26 tajima add
        pstKotoTran.CNT1N = zeroByteToNull(Me.txt_内容壱サブ壱)
        pstKotoTran.CNT2N = zeroByteToNull(Me.txt_内容壱サブ弐)
        '▲ 2007/01/26 tajima add
        
        If a更新区分 = P_新規モード Then
          '更新の場合は既に構造体に設定済み
          pstKotoTran.BUMOC = pstrBumonCd
          pstKotoTran.KOKYC = pParamArgs.KOKYC
          ' 作成情報
          pstKotoTran.INSED = Format$(DATE, "YYYYMMDD")
          pstKotoTran.INSEJ = Format$(time, "HHMMSS")
          pstKotoTran.INSPB = Me.Form.NAME
          pstKotoTran.INSUB = pstrUserName
        Else
          pstKotoTran.UPDAD = Format$(DATE, "YYYYMMDD")
          pstKotoTran.UPDAJ = Format$(time, "HHMMSS")
          pstKotoTran.UPDPB = Me.Form.NAME
          pstKotoTran.UPDUB = pstrUserName
        End If
                                
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :ZeroByteToNull
'        PARAMETER        :評価対象
'        RETURN           :変換結果
'        CREATE_DATE      :2006/05/19
'        NOTE             :文字列長がなければNULL返す
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function zeroByteToNull(varInput As Variant) As Variant

    zeroByteToNull = varInput
    
    If Nz(varInput, "") = "" Then
        zeroByteToNull = Null
    End If
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :zeroByteToNullorFormat
'        PARAMETER        :評価対象
'        RETURN           :変換結果
'        CREATE_DATE      :2006/05/19
'        NOTE             :文字列長がなければNULL返す、あれば指定フォーマット編集
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function zeroByteToNullorFormat(varInput As Variant, aFormat As String) As Variant
  
    If Nz(varInput, "") = "" Then
        zeroByteToNullorFormat = Null
    Else
        zeroByteToNullorFormat = Format(varInput, aFormat)
    End If
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :タイトル名の設定
'        MODULE_ID        :setTitleName
'        PARAMETER        :
'        CREATE_DATE      :2006/06/23
'        NOTE             :
'        UPDATE_DATE      :2017/11/18
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setTitleName()

      Dim titleName As String
      
      titleName = ""
      
      Select Case pParamArgs.修繕対象
          Case CmReprMod.P_対象_ヤード
              titleName = "ヤード修繕"
          Case CmReprMod.P_対象_ボックス
              titleName = "コンテナ(ボックス)修繕"
          Case CmReprMod.P_対象_部屋
              titleName = "コンテナ(部屋)修繕"
              'INSERT 2017-11-19 K.SATO START
              If pParamArgs.画面ID = "FVS904" Then
                   titleName = "ヤード一括修繕"
              End If
              'INSERT 2017-11-19 K.SATO START
          Case CmReprMod.P_対象_顧客
              titleName = "顧客ｺﾒﾝﾄ･ｸﾚｰﾑ履歴"
      End Select
          
      Select Case pParamArgs.処理モード
          Case P_新規モード
              titleName = titleName & "・新規"
          Case P_修正モード
              titleName = titleName & "・修正"
          Case P_参照モード
              titleName = titleName & "・参照"
      End Select
    
      '2017-11-11 add
      If pParamArgs.画面ID = "FVS904" Then
        If pParamArgs.一括 Then
              titleName = titleName & "（一括）"
        End If
      End If
    
    lbl_TITTLE.Caption = titleName

End Sub

'==============================================================================*
'
'        MODULE_NAME      :入力区分による画面項目の表示設定
'        MODULE_ID        :setDispInputItems
'        PARAMETER        :
'        CREATE_DATE      :2006/06/23
'        NOTE             :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setDispInputItems(ByVal a入力区分 As String)

      ' 以下は如何なる場合でも表示
      ' Me.lbl_入力区分
      ' Me.cmb_入力区分
      ' Me.txt_入力区分
      ' Me.lbl_日付
      ' Me.txt_日付
      ' Me.lbl_内容壱
      ' Me.txt_内容壱
      ' Me.lbl_入力担当
      ' Me.cmb_入力担当
      ' Me.txt_入力担当

      ' 以下は入力区分、修繕対象によって可変する
      Select Case a入力区分
      
          Case CmReprMod.P_修繕_工事
              Call set画面項目状態(True, False, True)
              ' 内容壱の標題変更
              Me.lbl_内容壱.Caption = "工事内容"
              
          Case CmReprMod.P_修繕_修理, CmReprMod.P_修繕_メンテ依頼, CmReprMod.P_修繕_組替え, CmReprMod.P_修繕_鍵返却, CmReprMod.P_修繕_社内理由, CmReprMod.P_修繕_オーナー理由, CmReprMod.P_修繕_解約予定, CmReprMod.P_修繕_未設置, CmReprMod.P_修繕_仮置き  'UPDATE 2022/02/27 N.IMAI
              ' 内容弐、修繕詳細は部屋修理の場合のみTrue,以外はFalseとする
              Dim is部屋修繕 As Boolean
              ' ▼ 2007/01/26 ヤードも部屋も入力項目は同じ
              is部屋修繕 = True 'IIf(pParamArgs.修繕対象 = CmReprMod.P_対象_部屋, True, False)
              ' ▲ 2007/01/26
              
              Call set画面項目状態(True, is部屋修繕, True, is部屋修繕, is部屋修繕)
              ' 内容壱弐の標題変更
              Me.lbl_内容壱.Caption = "依頼内容"
              Me.lbl_内容弐.Caption = "作業内容"
              
              ' 完了日の入力可否設定
              Me.txt_完了日.Enabled = IIf(Me.chk_修理完了.VALUE = 0, False, True)
              
          Case CmReprMod.P_修繕_クレーム, CmReprMod.P_修繕_コメント
              Call set画面項目状態
              ' 内容壱の標題変更
              If a入力区分 = CmReprMod.P_修繕_クレーム Then
                  Me.lbl_内容壱.Caption = "クレーム内容"
              Else
                  Me.lbl_内容壱.Caption = "コメント"
              End If
              
          Case CmReprMod.P_修繕_撤去予定
              Call set画面項目状態
              ' 内容壱の標題変更
              Me.lbl_内容壱.Caption = "備考(撤去先等)"
              
           
          Case Else
              ' 指定コード以外は真っ新に
              Call set画面項目状態
              ' 内容壱の標題変更
              Me.lbl_内容壱.Caption = "内容"
       
      End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :修繕対象による入力区分の設定
'        MODULE_ID        :set入力区分コンボ
'        PARAMETER        :
'        CREATE_DATE      :2006/06/23
'        NOTE             :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub set入力区分コンボ(dbSQLServer As DAO.Database)

    Dim wheres As String
    Dim reprWheres As String  ' 2007/01/26 add tajima
    
    wheres = ""
    reprWheres = "" ' 2007/01/26 add tajima
    Select Case pParamArgs.修繕対象
        
        Case CmReprMod.P_対象_ヤード
            wheres = P_入力区分ヤードWHERE
            reprWheres = P_ヤード修理区分WHERE  ' 2007/01/26 add tajima
        
        Case CmReprMod.P_対象_ボックス, CmReprMod.P_対象_部屋
            wheres = P_入力区分部屋WHERE
            reprWheres = P_部屋修理区分WHERE  ' 2007/01/26 add tajima
        
        Case CmReprMod.P_対象_顧客
            wheres = P_入力区分顧客WHERE
            reprWheres = "" ' 2007/01/26 add tajima
    
    End Select

    Call CmReprMod.SetCodeCmbItems(dbSQLServer, Me.cmb_入力区分, wheres)
    '▼ 2007/01/26 add tajima
    If reprWheres <> "" Then
        Call CmReprMod.SetCodeCmbItems(dbSQLServer, Me.cmb_修理区分, reprWheres)
    End If
    '▲ 2007/01/26 add
End Sub

'==============================================================================*
'
'        MODULE_NAME      :修繕対象による入力区分の設定
'        MODULE_ID        :set画面項目状態
'        PARAMETER        :
'        CREATE_DATE      :2006/06/23
'        NOTE             :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub set画面項目状態( _
                            Optional a金額関連 As Boolean = False, _
                            Optional a内容弐関連 As Boolean = False, _
                            Optional a作業者関連 As Boolean = False, _
                            Optional a修理関連 As Boolean = False, _
                            Optional a修理詳細 As Boolean = False _
                            )
                            
        Me.lbl_修理区分.Visible = a修理関連
        Me.cmb_修理区分.Visible = a修理関連
        Me.txt_修理区分.Visible = a修理関連
        Me.lbl_新鍵.Visible = a修理関連
        Me.txt_新鍵.Visible = a修理関連
        
        If a修理関連 = False Then
            Me.cmb_修理区分.VALUE = Null
            Me.txt_修理区分.VALUE = Null
            Me.txt_新鍵.VALUE = Null
        End If
        
        Me.chk_修理完了.Visible = a修理詳細
        Me.lbl_修理完了.Visible = a修理詳細
        If a修理詳細 = False Then Me.chk_修理完了.VALUE = 0
        
        Me.lbl_完了日.Visible = a修理詳細
        Me.txt_完了日.Visible = a修理詳細
        If a修理詳細 = False Then Me.txt_完了日.VALUE = Null
        
        Me.lbl_対応期限日.Visible = a修理詳細
        Me.txt_対応期限日.Visible = a修理詳細
        If a修理詳細 = False Then Me.txt_対応期限日.VALUE = Null
              
        Me.lbl_金額.Visible = a金額関連
        Me.txt_金額.Visible = a金額関連
        If a金額関連 = False Then Me.txt_金額.VALUE = Null
              
        Me.lbl_内容弐.Visible = a内容弐関連
        Me.txt_内容弐.Visible = a内容弐関連
        If a内容弐関連 = False Then Me.txt_内容弐.VALUE = Null
        '▼2007/01/26 add tajima
        Me.txt_内容弐サブ壱.Visible = a内容弐関連
        If a内容弐関連 = False Then Me.txt_内容弐サブ壱.VALUE = Null
        Me.txt_内容弐サブ弐.Visible = a内容弐関連
        If a内容弐関連 = False Then Me.txt_内容弐サブ弐.VALUE = Null
        '▲2007/01/26 add tajima

        Me.lbl_作業者.Visible = a作業者関連
        Me.cmb_作業者.Visible = a作業者関連
        Me.txt_作業者.Visible = a作業者関連
        If a修理関連 = False Then
            Me.cmb_作業者.VALUE = Null
            Me.txt_作業者.VALUE = Null
        End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :印刷ボタン押下処理
'        MODULE_ID        :cmb_印刷_Click
'        PARAMETER        :
'        CREATE_DATE      :2006/09/06
'        NOTE             :
'        UPDATE_DATE      :2017/11/18
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_印刷_Click()

 On Error GoTo Exception
 
    Dim roomCSV As String  'INSERT 2017/11/18 K.SATO

    'If Nz(Me.cmb_入力区分.VALUE, "") <> CmReprMod.P_修繕_修理 Then                                                                     'DELETE 2021/08/03 N.IMAI
    If Nz(Me.cmb_入力区分.VALUE, "") <> CmReprMod.P_修繕_修理 And Nz(Me.cmb_入力区分.VALUE, "") <> CmReprMod.P_修繕_メンテ依頼 Then     'INSERT 2021/08/03 N.IMAI
        MsgBox "修理以外は修理依頼書は印刷できません。", vbInformation, P_MSG_ERR
        Exit Sub
    End If
    
    'INSERT 2017-11-18 K.SATO start
    If pParamArgs.修繕対象 = CmReprMod.P_対象_部屋 Then
        If pParamArgs.一括 Then
            roomCSV = FVS904_sub.getOnDataCSVRooms
        End If
    End If
    'INSERT 2017-11-18 K.SATO end
    
    If pIs修正 = True Then
        If MsgBox("修正中の為、一度保存してから印刷を行います。" & vbCr & "よろしいですか？", _
                    vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
            Exit Sub
        End If
        
      If func保存処理() = False Then
            Exit Sub
        End If
        
        pParamArgs.処理モード = P_修正モード ' 一度保存したら修正モードとしておく
        
    End If
    
    If pParamArgs.修繕対象 = CmReprMod.P_対象_部屋 Then ' 2007/01/26 add tajima
        'INSERT 2017-11-18 K.SATO start
        If pParamArgs.一括 Then
            If False = Print910.fncPrintMaintenanceRequest( _
                                                Print910.P910_MODE_PREVIEW, _
                                                "", _
                                                pParamArgs.YARDC, _
                                                pParamArgs.YARDC, _
                                                False, _
                                                "", _
                                                "", _
                                                "", _
                                                "", _
                                                "", _
                                                roomCSV) Then
                MsgBox "修理依頼書が印刷できませんでした。", vbInformation, P_MSG_ERR
            End If
        'INSERT 2017-11-18 K.SATO end
        Else                                                                          'INSERT 2017-11-18 K.SATO
            If False = Print910.fncPrintMaintenanceRequest( _
                                                Print910.P910_MODE_PREVIEW, _
                                                pParamArgs.UNIQEC) Then
                MsgBox "修理依頼書が印刷できませんでした。", vbInformation, P_MSG_ERR
            End If                                                                        'INSERT 2017-11-18 K.SATO
        End If                                                                        'INSERT 2017-11-18 K.SATO
    End If
    '▼ 2007/01/26 add tajima
    '   ヤードの修理依頼書印刷
    If pParamArgs.修繕対象 = CmReprMod.P_対象_ヤード Then
        
        'INSERT 2017-11-18 K.SATO start
        If pParamArgs.一括 Then
            If False = Print910.fncPrintMaintenanceRequest( _
                                                Print910.P910_MODE_PREVIEW, _
                                                "", _
                                                pParamArgs.YARDC, _
                                                pParamArgs.YARDC, _
                                                False, _
                                                "", _
                                                "", _
                                                "", _
                                                "", _
                                                "", _
                                                roomCSV) Then
                MsgBox "修理依頼書が印刷できませんでした。", vbInformation, P_MSG_ERR
            End If
        Else
        'INSERT 2017-11-18 K.SATO end
            If False = Print901.fncPrintYardMaintenanceRequest( _
                                                Print901.P901_MODE_PREVIEW, _
                                                pParamArgs.UNIQEC) Then
                MsgBox "修理依頼書が印刷できませんでした。", vbInformation, P_MSG_ERR
            End If
        End If 'INSERT 2017-11-18 K.SATO
    End If
    '▲ 2007/01/26 add
    
    Exit Sub
Exception:
    Call MSZZ024_M00("cmb_印刷_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :保存処理
'        MODULE_ID        :func保存処理
'        PARAMETER        :
'        CREATE_DATE      :2006/09/06
'        NOTE             :保存ボタン押下の処理内容を別出しにした
'
'        UPDATE_DATE      :2017/11/11
'        NOTE             :一括修繕対応
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function func保存処理() As Boolean

    Dim loopcnt As Integer  '2017-11-11 add k.sato
    Dim wk As Integer       '2017-11-11 add k.sato
    Dim i As Integer        '2017-11-11 add k.sato

    func保存処理 = False
  
  On Error GoTo Exception

    ' 権限チェック 部屋が対象のときのみチェック
    If pParamArgs.権限可否 = False And pParamArgs.修繕対象 = CmReprMod.P_対象_部屋 Then
        If Me.cmb_入力区分.VALUE = CmReprMod.P_修繕_撤去予定 Then
            MsgBox "権限がないため撤去に対する登録変更は出来ません。", vbExclamation, P_MSG_ERR
            Exit Function
        'ElseIf Me.cmb_入力区分.VALUE = CmReprMod.P_修繕_修理 And Me.chk_修理完了.VALUE = -1 Then                                                           'DELETE 2021/08/03 N.IMAI
        ElseIf (Me.cmb_入力区分.VALUE = CmReprMod.P_修繕_修理 Or Me.cmb_入力区分.VALUE = CmReprMod.P_修繕_メンテ依頼) And Me.chk_修理完了.VALUE = -1 Then   'INSERT 2021/08/03 N.IMAI
            MsgBox "権限がないため修理完了に対する登録変更は出来ません。", vbExclamation, P_MSG_ERR
            Exit Function
        End If
    End If


    ' 入力チェック
    If fncChkInput = False Then Exit Function
    
    '2017-11-11 add start k.sato
    loopcnt = 1
    If pParamArgs.一括 Then
        loopcnt = FVS904_sub.cntOnData
    End If
    '2017-11-11 add end k.sato

    wk = 0
    For i = 1 To loopcnt   '2017-11-11 add k.sato
        
        '2017-11-11 add start k.sato
        If pParamArgs.一括 Then
            wk = FVS904_sub.getOneEntryData
            
            pParamArgs.YARDC = FVS904_sub.pSelectArgs(wk).YARDC
            pParamArgs.ROOMC = FVS904_sub.pSelectArgs(wk).ROOMC
            If pParamArgs.処理モード <> P_新規モード Then
                pParamArgs.UNIQEC = FVS904_sub.pSelectArgs(wk).UNIQEC
            End If
            
            pstReprTran.YARDC = pParamArgs.YARDC
            pstReprTran.ROOMC = pParamArgs.ROOMC
            If pParamArgs.処理モード <> P_新規モード Then
                pstReprTran.UNIQEC = pParamArgs.UNIQEC
            End If
            
        End If
        '2017-11-11 add end k.sato
        
        ' 撤去ならば部屋の利用可否チェック 2006/09/06
        If pParamArgs.修繕対象 = P_対象_部屋 And Me.cmb_入力区分.VALUE = CmReprMod.P_修繕_撤去予定 Then
            Dim errorMessage As String
            errorMessage = ""
            Call subConnectDB(P_DB_KONT, pstrBumonCd)
            If 0 = CmKomsMod.CheckLending(pobjKONT_DB, _
                                           CmKomsMod.pcintYOYAKU_UNITCHECK + pcintKEIYAKU_UNITCHECK, _
                                           CLng(pParamArgs.YARDC), _
                                           CLng(pParamArgs.ROOMC), _
                                           errorMessage) Then
                errorMessage = ""
            End If
            If Not pobjKONT_DB Is Nothing Then: Call subDisConnectDB(P_DB_KONT)
            If errorMessage <> "" Then
                MsgBox errorMessage & vbCr & "よって撤去することは出来ません。", vbExclamation, P_MSG_ERR
                Exit Function
            End If
        End If
    
        ' 登録処理
        If fncDB登録処理() = False Then Exit Function
    
        ' フラグ初期化
        Call subSetChangeFlag(False)
    
        If pParamArgs.修繕対象 = P_対象_部屋 Then                                   'INSERT START 2012/06/21 K.ISHIZAKA
            '統合ＤＢへの更新
            If Not MSTGP01_M10(pParamArgs.YARDC, Nz(pParamArgs.ROOMC, "")) Then
                Call MSZZ024_M10("MSTGP01_M10", "統合ＤＢへの更新に失敗しました。")
            End If
        End If                                                                      'INSERT END   2012/06/21 K.ISHIZAKA
        
        blnUpdFlg = True

    Next '2017-11-11 add k.sato

    ' 呼び出し元画面のリロード
    Call subParentReload

    blnUpdFlg = False
    func保存処理 = True
    
    Exit Function

Exception:
    If Not pobjKONT_DB Is Nothing Then: Call subDisConnectDB(P_DB_KONT)
    Call Err.Raise(Err.Number, "func保存処理" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :KONT系ＤＢへの修繕内容反映処理
'        MODULE_ID        :subKont系反映
'        PARAMETER        :
'        CREATE_DATE      :
'        NOTE             :KONT系DBとしてCNTA_MASTへの反映を行う
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function subKontCont() As Boolean

    Dim strSQL            As String
    Dim rsObject          As Recordset
    Dim isCNTA_Chenging   As Boolean 'コンテナマスタを変更したかのフラグ
 
  On Error GoTo Exception
    subKontCont = False

    ' ヤードコード＆コンテナコードで更新対象抽出
    strSQL = "SELECT * "
    strSQL = strSQL & " FROM CNTA_MAST WHERE CNTA_CODE = " & pstReprTran.YARDC
    strSQL = strSQL & " AND CNTA_NO = " & pstReprTran.ROOMC
    
    Set rsObject = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

    rsObject.Edit
    
    rsObject.Fields("CNTA_KEY_TYPE") = "15"
    'rsObject.Fields("CNTA_KEY_NO") = zeroByteToNull(txt_新鍵)
    rsObject.Fields("CNTA_KEY_NUM") = 3
    rsObject.Fields("CNTA_UPDATE") = Format(Now, "yyyy/mm/dd")
    rsObject.UPDATE

    rsObject.Close
    Set rsObject = Nothing
   
    subKontCont = True
   
    Exit Function
   
Exception:
    subKontCont = False
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    Call Err.Raise(Err.Number, "subKont系反映" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function


'==============================================================================*
'
'        MODULE_NAME      :KONT系ＤＢへの修繕内容反映処理
'        MODULE_ID        :subKont系反映
'        PARAMETER        :
'        CREATE_DATE      :2006/09/06
'        NOTE             :KONT系DBとしてCNTA_MASTへの反映を行う
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKont系反映(ByVal a修理中可否 As Boolean)

    Dim strSQL            As String
    Dim rsObject          As Recordset
    Dim isCNTA_Chenging   As Boolean 'コンテナマスタを変更したかのフラグ
 
  On Error GoTo Exception
    isCNTA_Chenging = False

    ' ヤードコード＆コンテナコードで更新対象抽出
    strSQL = "SELECT * "
    strSQL = strSQL & " FROM CNTA_MAST WHERE CNTA_CODE = " & pstReprTran.YARDC
    strSQL = strSQL & " AND CNTA_NO = " & pstReprTran.ROOMC
    
    Set rsObject = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

    rsObject.Edit
    Select Case Me.cmb_入力区分.VALUE
        Case CmReprMod.P_修繕_修理, CmReprMod.P_修繕_メンテ依頼, P_修繕_組替え, P_修繕_鍵返却, P_修繕_社内理由, P_修繕_オーナー理由, P_修繕_解約予定, P_修繕_未設置, CmReprMod.P_修繕_仮置き    'UPDATE 2022/02/27 N.IMAI
        
            ' 完了日が未入力または他の修繕が修理中ならば「修理中」
            If Nz(pstReprTran.COMPD, "") = "" Or a修理中可否 = True Then
                'If Nz(rsObject.Fields("CNTA_REPR_KBN"), "0") <> "1" Then                                       'DELETE 2023/04/18 N.IMAI
                If Nz(rsObject.Fields("CNTA_REPR_KBN"), "0") <> "1" Or rsObject.Fields("CNTA_USE") <> "0" Then  'INSERT 2023/04/18 N.IMAI
                    rsObject.Fields("CNTA_REPR_KBN") = 1
                    rsObject.Fields("CNTA_USE") = 0
                    rsObject.Fields("CNTA_UKNO") = Null '2007/04/11 tajima add
                    isCNTA_Chenging = True
                End If
            ' 完了日が入力済みまたは他の修繕が修理完了ならば「修理完了」
            Else
                If Nz(rsObject.Fields("CNTA_REPR_KBN"), "0") <> "0" Or a修理中可否 = False Then
                    rsObject.Fields("CNTA_REPR_KBN") = 0
                    rsObject.Fields("CNTA_USE") = 1
                    isCNTA_Chenging = True
                End If
                '修理完了後､契約中でなければ紹介待ちとする 2007/04/17 add tajima
                Dim ereMsg As String
                If False = CmKomsMod.CheckLending(pobjKONT_DB, _
                                                  pcintKEIYAKU_UNITCHECK, _
                                                  pstReprTran.YARDC, _
                                                  pstReprTran.ROOMC, _
                                                  ereMsg) Then
                    rsObject.Fields("CNTA_USE") = 0
                    rsObject.Fields("CNTA_UKNO") = "FVS910"
                    isCNTA_Chenging = True
                End If
                
                '完了時のみ鍵区分を変更する。
                '2015/03/02 M.HONDA INS
                'If cmb_入力区分 = "02" And cmb_修理区分 = "05" And txt_新鍵 <> "" Then                                                                     'DELETE 2021/08/03 N.IMAI
                If (cmb_入力区分 = CmReprMod.P_修繕_修理 Or cmb_入力区分 = CmReprMod.P_修繕_メンテ依頼) And cmb_修理区分 = "05" And txt_新鍵 <> "" Then     'INSERT 2021/08/03 N.IMAI
                    rsObject.Fields("CNTA_KEY_TYPE") = "15"
                    rsObject.Fields("CNTA_KEY_NUM") = 3
                End If
                '2015/03/02 M.HONDA INS
                
            End If
        
        Case CmReprMod.P_修繕_撤去予定
            ' 撤去予定で登録するならばコンテナは撤去
            If Nz(rsObject.Fields("CNTA_USE"), "0") <> "9" Then
                rsObject.Fields("CNTA_USE") = 9
                rsObject.Fields("CNTA_UKNO") = Null '2007/04/11 tajima add
                isCNTA_Chenging = True
            End If
    End Select
    ' 更新日設定
    If isCNTA_Chenging = True Then
        rsObject.Fields("CNTA_UPDATE") = Format(Now, "yyyy/mm/dd")
        rsObject.UPDATE
    End If


    'If cmb_入力区分 = "02" And cmb_修理区分 = "05" And txt_新鍵 <> "" And Me.chk_修理完了.VALUE = -1 Then                                                                  'DELETE 2021/08/03 N.IMAI
    If (cmb_入力区分 = CmReprMod.P_修繕_修理 Or cmb_入力区分 = CmReprMod.P_修繕_メンテ依頼) And cmb_修理区分 = "05" And txt_新鍵 <> "" And Me.chk_修理完了.VALUE = -1 Then  'INSERT 2021/08/03 N.IMAI
        ' 契約トランの抽出
        strSQL = "SELECT TOP 1  *  "
        strSQL = strSQL & " FROM CARG_FILE INNER JOIN RCPT_TRAN ON CARG_FILE.CARG_ACPTNO = RCPT_TRAN.RCPT_CARG_ACPTNO WHERE CARG_YCODE = " & pstReprTran.YARDC
        strSQL = strSQL & " AND CARG_NO = " & pstReprTran.ROOMC
        strSQL = strSQL & " ORDER BY CARG_FSDATE DESC "
        Set rsObject = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
        If rsObject.RecordCount <> 0 Then    '2015/05/09 M.HONDA INS
            rsObject.Edit
            If rsObject.Fields("CARG_AGRE") <> 9 Then
                '完了以外の場合には、鍵区分を書き換える。
                rsObject.Fields("CARG_KAGIICD") = 15
                rsObject.Fields("RCPT_KAGIICD") = 15
                rsObject.Fields("RCPT_KAINO") = txt_新鍵
                rsObject.UPDATE
            End If
        End If
        
    End If


    rsObject.Close
    Set rsObject = Nothing
   
    Exit Sub
   
Exception:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    Call Err.Raise(Err.Number, "subKont系反映" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'2007/01/26 add new method
'==============================================================================*
'
'        MODULE_NAME      :txt_内容壱サブ壱_BeforeUpdate イベント
'        MODULE_ID        :txt_内容壱サブ壱_BeforeUpdate
'        CREATE_DATE      :2007/01/26
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容壱サブ壱_BeforeUpdate(Cancel As Integer)
    If fncChkTextLength(Me.txt_内容壱サブ壱, 80) = False Then
        Cancel = True
''        SendKeys "{ESC}"      '----20091203----M.RYU----del----'修正依頼：桁数オーバーの場合、入力内容をそのまま保留
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_内容壱サブ壱_Change イベント
'        MODULE_ID        :txt_内容壱サブ壱_Change
'        CREATE_DATE      :2007/01/26
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容壱サブ壱_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :xt_内容壱サブ壱_KeyPress
'        MODULE_ID        :txt_内容壱サブ壱_KeyPress
'        NOTE             :
'        CREATE_DATE      :2007/01/26
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容壱サブ壱_KeyPress(KeyAscii As Integer)
    Call CmKeyPress.CmKeyPress_For_Char2(txt_内容壱サブ壱, KeyAscii, 80)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_内容壱サブ弐_BeforeUpdate イベント
'        MODULE_ID        :txt_内容壱サブ弐_BeforeUpdate
'        CREATE_DATE      :2007/01/26
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容壱サブ弐_BeforeUpdate(Cancel As Integer)
    If fncChkTextLength(Me.txt_内容壱サブ弐, 80) = False Then
        Cancel = True
''        SendKeys "{ESC}"      '----20091203----M.RYU----del----'修正依頼：桁数オーバーの場合、入力内容をそのまま保留
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_内容壱サブ弐_Change イベント
'        MODULE_ID        :txt_内容壱サブ弐_Change
'        CREATE_DATE      :2007/01/26
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容壱サブ弐_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :xt_内容壱サブ弐_KeyPress
'        MODULE_ID        :txt_内容壱サブ弐_KeyPress
'        NOTE             :
'        CREATE_DATE      :2007/01/26
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容壱サブ弐_KeyPress(KeyAscii As Integer)
    Call CmKeyPress.CmKeyPress_For_Char2(txt_内容壱サブ弐, KeyAscii, 80)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_内容弐サブ壱_BeforeUpdate イベント
'        MODULE_ID        :txt_内容弐サブ壱_BeforeUpdate
'        CREATE_DATE      :2007/01/26
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容弐サブ壱_BeforeUpdate(Cancel As Integer)
    If fncChkTextLength(Me.txt_内容弐サブ壱, 80) = False Then
        Cancel = True
''        SendKeys "{ESC}"      '----20091203----M.RYU----del----'修正依頼：桁数オーバーの場合、入力内容をそのまま保留
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_内容弐サブ壱_Change イベント
'        MODULE_ID        :txt_内容弐サブ壱_Change
'        CREATE_DATE      :2007/01/26
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容弐サブ壱_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :xt_内容弐サブ壱_KeyPress
'        MODULE_ID        :txt_内容弐サブ壱_KeyPress
'        NOTE             :
'        CREATE_DATE      :2007/01/26
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容弐サブ壱_KeyPress(KeyAscii As Integer)
    Call CmKeyPress.CmKeyPress_For_Char2(txt_内容弐サブ壱, KeyAscii, 80)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_内容弐サブ弐_BeforeUpdate イベント
'        MODULE_ID        :txt_内容弐サブ弐_BeforeUpdate
'        CREATE_DATE      :2007/01/26
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容弐サブ弐_BeforeUpdate(Cancel As Integer)
    If fncChkTextLength(Me.txt_内容弐サブ弐, 80) = False Then
        Cancel = True
''        SendKeys "{ESC}"      '----20091203----M.RYU----del----'修正依頼：桁数オーバーの場合、入力内容をそのまま保留
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_内容弐サブ弐_Change イベント
'        MODULE_ID        :txt_内容弐サブ弐_Change
'        CREATE_DATE      :2007/01/26
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容弐サブ弐_Change()
    Call subSetChangeFlag(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :xt_内容弐サブ弐_KeyPress
'        MODULE_ID        :txt_内容弐サブ弐_KeyPress
'        NOTE             :
'        CREATE_DATE      :2007/01/26
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_内容弐サブ弐_KeyPress(KeyAscii As Integer)
    Call CmKeyPress.CmKeyPress_For_Char2(txt_内容弐サブ弐, KeyAscii, 80)
End Sub

'****************************  ended or program ********************************
Private Sub フォームヘッダー_Click()

End Sub
