VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS460"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　（株）加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　価格テンプレートマスタメンテ
'   プログラムＩＤ　：
'   作　成　日　　　：  2017/11/07
'   作　成　者　　　：  Y.SUZUKI
'
'   更　新　日　　　：  2018/08/11
'   更　新　者　　　：  EGL
'
'**********************************************
Option Compare Database
Option Explicit

'**********************************************
' 定数宣言
'**********************************************
Private Const strPROG_ID            As String = "FKS460"

Private Const strmsg保存確認        As String = "保存します。よろしいですか？"
Private Const strmsg取消確認        As String = "データ入力を取り消します。よろしいですか？"
Private Const strmsg削除確認        As String = "データを削除します。よろしいですか？"
Private Const strmsg終了確認        As String = "終了します。よろしいですか？"

Private Const strmsg更新エラー      As String = "更新中にエラーが発生しました。（処理は中断されました。）"
Private Const strmsgクリア確認      As String = "クリアされます。よろしいですか？"
Private Const strMSGデータなし      As String = "存在しないコードです。"
Private Const strmsg重複            As String = "既に存在するコードです。"

Private Const strMSG桁数オーバー    As String = "桁数オーバーしています。"
Private Const strMSG半角入力エラー  As String = "半角英数字で入力してください。"
Private Const strMsg_kaku           As String = "確認"
Private Const strMsg_err            As String = "エラー"

Private Const str価格表             As String = "0"    '2018/08/11
Private Const str事務手数料         As String = "63"   '2018/08/11
Private Const pcsglMoneyMaxValue    As Long = 999999999  '      '2018/08/11

'**********************************************
' 変数宣言
'**********************************************
Private pstrBumoCD                  As String       ' 部門コード

Private strUSER_ID                  As String
Private strCaption                  As String       'タイトルを保持
Private blnModifyFlg                As Boolean      'True=変更、False=未編集

'==============================================================================*
'   MODULE_NAME :   cmb_テンプレートタイプの更新
'   MODULE_ID   :   cmb_テンプレートタイプ_AfterUpdate
'   CREATE_DATE :   2018/08/11
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_テンプレートタイプ_AfterUpdate()
    '空レコード作成
    Call subCreateMeisai
    '入力項目制御
    Call subInputSwKakakuJimu
    '複写用テンプレートコード初期化
    cmb_テンプレートコード2.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル2.BackColor = RGB(255, 0, 204)
    End Select

End Sub

'==============================================================================*
'   MODULE_NAME :   opg_処理選択_BeforeUpdateイベント
'   MODULE_ID   :   opg_処理選択
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opg_処理選択_BeforeUpdate(Cancel As Integer)

    On Error GoTo Exception
    
    Dim intAns      As Integer
    
    '変更フラグがTrueの時は確認メッセージを出す
    If blnModifyFlg Then
        intAns = MsgBox(strmsgクリア確認, vbQuestion + vbYesNo, strMsg_kaku)
        If vbNo = intAns Then
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        End If
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   opg_処理選択_AfterUpdateイベント
'   MODULE_ID   :   opg_処理選択
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opg_処理選択_AfterUpdate()
    
    On Error GoTo Exception
    
    Dim intMode     As Integer
    
    'コマンドボタン使用可否設定
    intMode = opg_処理選択.VALUE        ' 処理モード
    
    Call set_CtrlButton(intMode)

    '画面クリア
    Call clr_SCREEN
    
    '2018/08/11 カーソルフォーカス設定
    Select Case intMode
    Case 1
        '新規
        Me.cmb_テンプレートタイプ.Enabled = True
        'テンプレートコードにカーソルを設定
        cmb_テンプレートタイプ.SetFocus
    Case 2, 3
        Me.cmb_テンプレートタイプ.Enabled = False
        'テンプレートコードにカーソルを設定
        cmb_テンプレートコード.SetFocus
    End Select
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_AfterUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmd_保存_Clickイベント
'   MODULE_ID   :   cmd_保存
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_Click()
    '
    Dim strTcode        As String
    Dim strTNAME        As String
    Dim strKEYAD        As String
    Dim strKEYPR        As String
    
    Dim strDML          As String       '重複チェック用SQL
    
    Dim lngRTN          As Long
    Dim blnRTN          As Boolean
    
    On Error GoTo Exception
    
    Application.Echo False              '画面の描画を止める
    
    '********************
    ' 初期処理
    '********************
    '入力中のデータ確定
    doCmd.RunCommand acCmdSaveRecord
    FKS460_S01.Requery
    
    '********************
    ' 入力チェック
    '********************
    'コード - 必須
    strTcode = Nz(cmb_テンプレートコード.VALUE, "")
    If "" = strTcode Then
        MsgBox "テンプレートコードが入力されていません。", vbInformation + vbOK, strMsg_err
        cmb_テンプレートコード.SetFocus
        GoTo Exit_Sub
    End If
    
    '名称 - 必須
    strTNAME = Nz(txt_テンプレート名.VALUE, "")
    If "" = strTNAME Then
        MsgBox "テンプレート名称が入力されていません。", vbInformation + vbOK, strMsg_err
        txt_テンプレート名.SetFocus
        GoTo Exit_Sub
    End If
    
    'キーワード(地区) - 必須
    strKEYAD = Nz(txt_キーワード地区.VALUE, "")
    If "" = strKEYAD Then
        MsgBox "キーワード(地区)が入力されていません。", vbInformation + vbOK, strMsg_err
        txt_キーワード地区.SetFocus
        GoTo Exit_Sub
    End If
    
    'キーワード(価格) - 必須
    strKEYPR = Nz(txt_キーワード価格.VALUE, "")
    If "" = strKEYPR Then
        MsgBox "キーワード(価格)が入力されていません。", vbInformation + vbOK, strMsg_err
        txt_キーワード価格.SetFocus
        GoTo Exit_Sub
    End If
    
    '一覧 - 重複（段区分、用途、実帖）
    lngRTN = chk_FKS460_W01
    '2018/08/11 事務手数料の場合はNULLを許可
    'If 1 = lngRTN Then
    If 1 = lngRTN And cmb_テンプレートタイプ.VALUE = str価格表 Then
        MsgBox "実帖が 0、または入力されていない行があります。", vbInformation + vbOK, strMsg_err
        FKS460_S01.SetFocus
        GoTo Exit_Sub
    '2018/08/11 事務手数料の場合はNULLを許可
    'ElseIf 2 = lngRTN Then
    ElseIf 2 = lngRTN And cmb_テンプレートタイプ.VALUE = str価格表 Then
        MsgBox "価格が 0、または入力されていない行があります。", vbInformation + vbOK, strMsg_err
        FKS460_S01.SetFocus
        GoTo Exit_Sub
    ElseIf 3 = lngRTN Then
        MsgBox "段区分・用途・実帖の組合せが重複しています。", vbInformation + vbOK, strMsg_err
        FKS460_S01.SetFocus
        GoTo Exit_Sub
    ElseIf 4 = lngRTN Then '2018/08/11
        MsgBox "事務手数料は「初回雑費２」に設定して下さい。", vbExclamation, strMsg_err
        FKS460_S01.SetFocus
        GoTo Exit_Sub
    ElseIf 5 = lngRTN Then '2018/08/11
        MsgBox "「初回雑費２」は事務手数料を設定して下さい。", vbExclamation, strMsg_err
        FKS460_S01.SetFocus
        GoTo Exit_Sub
    ElseIf 6 = lngRTN Then '2018/08/11
        MsgBox "１つの価格内で同じ雑費区分は指定できません。", vbExclamation, strMsg_err
        FKS460_S01.SetFocus
        GoTo Exit_Sub
    End If
    
    '********************
    ' 保存処理
    '********************
    '確認メッセージ
    lngRTN = MsgBox(strmsg保存確認, vbQuestion + vbYesNo, strMsg_kaku)
    If vbNo = lngRTN Then
        GoTo Exit_Sub
    End If
    
    '価格表テンプレートマスタ(ヘッダ・価格)登録
    blnRTN = upd_THTP
    If Not blnRTN Then
        GoTo Exit_Sub
    End If
    
    '********************
    ' 終了処理
    '********************
    '画面初期化
    Call clr_SCREEN
    
Exit_Sub:
    Application.Echo True               '画面の描画を行う
    Exit Sub
    
Exception:
    Application.Echo True               '画面の描画を行う
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmd_保存_Click", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmd_削除_Clickイベント
'   MODULE_ID   :   cmd_削除
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_削除_Click()
    Dim objCon              As Object
    Dim strTcode            As String
    Dim lngRTN              As Long
    
    On Error GoTo Exception
    
    '入力中のデータ確定
    doCmd.RunCommand acCmdSaveRecord
    
    Application.Echo False              '画面の描画を止める
    
    '入力チェック
    strTcode = Nz(cmb_テンプレートコード.VALUE, "")
    If "" = strTcode Then
        GoTo Exit_Sub
    End If
    
    '変更フラグがTrueの時は確認メッセージを出す
    lngRTN = MsgBox(strmsg削除確認, vbQuestion + vbYesNo, strMsg_kaku)
    If vbNo = lngRTN Then
        GoTo Exit_Sub
    End If
    
    Set objCon = MSZZ025.ADODB_Connection(pstrBumoCD)
    
    'トランザクション開始
    objCon.BeginTrans
    
    '価格表テンプレートマスタ（ヘッダ）の該当データを削除
    Call del_TH(objCon, strTcode)
    
    '価格表テンプレートマスタ（価格）の該当データを削除
    Call del_TP(objCon, strTcode)
    
    'コミット
    objCon.CommitTrans
    
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    
    '画面クリア
    Call clr_SCREEN
    
Exit_Sub:
    Application.Echo True               '画面の描画を行う
    Exit Sub
    
Exception:
    'ロールバック
    objCon.RollbackTrans
    
    Application.Echo True               '画面の描画を行う
    
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmd_削除_Click", True)
    
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmd_取消_Clickイベント
'   MODULE_ID   :   cmd_取消
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_取消_Click()
    
    Dim lngRTN  As Long
    
    On Error GoTo Exception
    
    If blnModifyFlg Then
        lngRTN = MsgBox(strmsg取消確認, vbQuestion + vbYesNo + vbDefaultButton2, strMsg_kaku)
        If vbNo = lngRTN Then
            Exit Sub
        End If
    End If
    
    '画面クリア
    Call clr_SCREEN
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmd_削除_Click", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmd_取消_KeyDownイベント
'   MODULE_ID   :   cmd_終了
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_取消_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim intMode     As Integer
    
    On Error GoTo Exception
    
    intMode = opg_処理選択.VALUE
    
    Select Case KeyCode
        Case vbKeyReturn
            Call cmd_取消_Click
            
        Case vbKeyTab
            If 0 < Shift Then
                cmb_テンプレートコード.SetFocus
            Else
                cmd_終了.SetFocus
            End If
            
        Case vbKeyUp, vbKeyRight
            cmd_終了.SetFocus

        Case vbKeyDown, vbKeyLeft
            If 2 <> intMode Then
                cmd_保存.SetFocus
            Else
                cmd_削除.SetFocus
            End If
    End Select
    
    KeyCode = 0
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmd_取消_KeyDown", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmd_終了_Clickイベント
'   MODULE_ID   :   cmd_終了
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
    
    Dim lngRTN  As Long
    
    On Error GoTo Exception
    
    '画面リセット
    If blnModifyFlg Then
        lngRTN = MsgBox(strmsg取消確認, vbQuestion + vbYesNo + vbDefaultButton2, strMsg_kaku)
        If vbNo = lngRTN Then
            Exit Sub
        End If
    End If
    
    '画面を閉じる
    doCmd.Close acForm, FormName
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmd_終了_Click", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmd_終了_KeyDownイベント
'   MODULE_ID   :   cmd_終了
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim intMode     As Integer
    
    On Error GoTo Exception
    
    intMode = opg_処理選択.VALUE
    
    Select Case KeyCode
        Case vbKeyReturn
            Call cmd_終了_Click
            
        Case vbKeyTab
            If 0 < Shift Then
                cmb_テンプレートコード.SetFocus
                
            Else
                If 2 <> opg_処理選択.VALUE Then
                    cmd_保存.SetFocus
                Else
                    cmd_削除.SetFocus
                End If
            End If
            
        Case vbKeyLeft, vbKeyDown
            cmd_取消.SetFocus
            
        Case vbKeyRight, vbKeyUp
            If 2 <> intMode Then
                cmd_保存.SetFocus
            Else
                cmd_削除.SetFocus
            End If
    End Select
    
    KeyCode = 0
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmd_終了_KeyDown", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   Form_KeyDownイベント
'   MODULE_ID   :   From_FKS460
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 画面がアクティブになった時処理
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Exception
    
    'F1=「終了」、 F2=「保存・削除」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            cmd_終了.SetFocus
            
        Case vbKeyF2
            KeyCode = 0
            If cmd_保存.Enabled Then cmd_保存.SetFocus
            If cmd_削除.Enabled Then cmd_削除.SetFocus
            
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
            
    End Select
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("Form_KeyDown", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   Form_Activateイベント
'   MODULE_ID   :   From_FKS460
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 画面がアクティブになった時処理
Private Sub Form_Activate()
    
    On Error GoTo Exception
    
    '画面最大化
    doCmd.Maximize
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("Form_Activate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   Form_Loadイベント
'   MODULE_ID   :   From_FKS460
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 読込時処理
Private Sub Form_Load()
    
    On Error GoTo Exception
    
    Dim intMode         As Integer
    
    '初期処理
    strUSER_ID = LsGetUserName()                            'ユーザーID
    intMode = opg_処理選択.VALUE                            '処理モード
    strCaption = Caption                                    'タイトルをインスタンスに保持
    
    pstrBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")     ' 部門コード取得
    
    'テンプレートタイプセット 2018/08/11
    Call subSetCmbTmpTyp
    
    '画面初期化
    Call clr_SCREEN
    
    'コマンドボタン使用可否設定
    Call set_CtrlButton(intMode)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("Form_Load", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmb_テンプレートタイプ初期化
'   MODULE_ID   :　 subSetCmbTmpTyp
'   CREATE_DATE :   2018/08/11
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Sub subSetCmbTmpTyp()
    cmb_テンプレートタイプ.RowSource = ""
    cmb_テンプレートタイプ.AddItem "0;価格表"
    cmb_テンプレートタイプ.AddItem "63;事務手数料"
End Sub

'==============================================================================*
'   MODULE_NAME :   Form_Closeイベント
'   MODULE_ID   :   From_FKS460
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 画面終了時処理
Private Sub Form_Close()
    
    On Error GoTo Exception
    
    'ワークテーブルを削除
    Call del_FKS460_W01(0)
    
    '画面終了
    Call MSHS010_M00(strPROG_ID, "2")
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("Form_Close", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmb_テンプレートコード_BeforeUpdateイベント
'   MODULE_ID   :   cmb_テンプレートコード
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' テンプレート変更時処理
Private Sub cmb_テンプレートコード_BeforeUpdate(Cancel As Integer)
    Dim strKKTHM_TCODE  As String
    Dim intMode         As Integer
    Dim blnRTN          As Boolean
    
    On Error GoTo Exception
    
    '入力値取得
    strKKTHM_TCODE = Nz(cmb_テンプレートコード.VALUE, "")
    intMode = opg_処理選択.VALUE
    
    '入力チェック
    If "" = strKKTHM_TCODE Then
        MsgBox "テンプレートコードを入力してください。", vbInformation, strMsg_kaku
        Exit Sub
    End If
    
    '桁数チェック
    If 6 < CmLenb(strKKTHM_TCODE) Then
        MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
        GoTo Exit_rtn
    End If
    
    '入力値の存在チェック
    blnRTN = chk_KKTH_MAST_EXIST(strKKTHM_TCODE, 1)
    
    '存在する場合
    If Not blnRTN Then
        GoTo Exit_rtn
    End If
    
    Exit Sub
    
Exit_rtn:
    Cancel = True
    SendKeys "{ESC}"
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmb_テンプレートコード_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmb_テンプレートコード_AfterUpdateイベント
'   MODULE_ID   :   cmb_テンプレートコード
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' テンプレート変更時処理
Private Sub cmb_テンプレートコード_AfterUpdate()
    'コード変更時
    Dim objCon          As Object
    Dim objWrk          As Object
    Dim strDML          As String
    
    Dim intMode         As Integer
    Dim strKKTHM_TCODE  As String
    
    Dim strKKTHM_TTYPE  As String '2018/08/11
    Dim strKKTHM_TNAME  As String
    Dim strKKTHM_KEYAD  As String
    Dim strKKTHM_KEYPR  As String
    
    Dim blnRTN          As Boolean
    
    On Error GoTo Exception
    
    '********************
    ' 初期処理
    '********************
    intMode = opg_処理選択.VALUE
    strKKTHM_TCODE = Nz(cmb_テンプレートコード.VALUE, "")
    
    If "" = strKKTHM_TCODE Then
        Call clr_SCREEN
        Exit Sub
    End If

    '********************
    ' モードにより処理分岐
    '********************
    If 2 = intMode Or 3 = intMode Then
        '**********
        '修正、削除
        '**********
        
        Set objCon = MSZZ025.ADODB_Connection(pstrBumoCD)
        
        'SQL文編集
        '2018/08/11 テンプレートタイプを追加
        'strDML = "select KKTHM_TNAME, KKTHM_KEYAD, KKTHM_KEYPR from KKTH_MAST where KKTHM_TCODE = '" & strKKTHM_TCODE & "'"
        strDML = "select KKTHM_TTYPE,KKTHM_TNAME, KKTHM_KEYAD, KKTHM_KEYPR from KKTH_MAST where KKTHM_TCODE = '" & strKKTHM_TCODE & "'"
        
        '検索
        Set objWrk = MSZZ025.ADODB_Recordset(strDML, objCon)
        
        'データ取得
        With objWrk
            If Not .EOF Then
                
                strKKTHM_TTYPE = Trim(Nz(.Fields("KKTHM_TTYPE"))) '2018/08/11 テンプレートタイプ
                strKKTHM_TNAME = Trim(Nz(.Fields("KKTHM_TNAME")))
                strKKTHM_KEYAD = Trim(Nz(.Fields("KKTHM_KEYAD")))
                strKKTHM_KEYPR = Trim(Nz(.Fields("KKTHM_KEYPR")))
                
            End If
            .Close
        End With
        
        '取得したデータを画面にせってい
        
        cmb_テンプレートタイプ.VALUE = strKKTHM_TTYPE '2018/08/11 テンプレートタイプ
        txt_テンプレート名.VALUE = strKKTHM_TNAME
        txt_キーワード地区.VALUE = strKKTHM_KEYAD
        txt_キーワード価格.VALUE = strKKTHM_KEYPR

        '入力項目制御 2018/08/11
        Call subInputSwKakakuJimu

        '詳細一覧を表示
        Call cre_FKS460_W01(strKKTHM_TCODE)
        
    Else
        
        '2018/08/11 空レコード作成
        Call subCreateMeisai
        
        '入力項目制御 2018/08/11
        Call subInputSwKakakuJimu

    End If
    
    '画面項目使用可能
    Call set_ENABLED(True)
    
    '編集中フラグを落とす
    set_ModifyFlg (False)
    
    '2018/08/11 複写用テンプレートコード初期化
    cmb_テンプレートコード2.Requery
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmb_テンプレートコード_AfterUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   明細行の入力制御
'   MODULE_ID   :   subInputSwKakakuJimu
'   CREATE_DATE :   2018/08/11
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subInputSwKakakuJimu()
    If cmb_テンプレートタイプ.VALUE = str価格表 Then
        '■価格表(明細域の制御)
        FKS460_S01.Form.KKTPM_PRICE.Enabled = True
        FKS460_S01.Form.KKTPM_FZAPPI_1_CHK.Enabled = False
        FKS460_S01.Form.KKTPM_FZAPPI_2_CODE.Enabled = True
        FKS460_S01.Form.KKTPM_FZAPPI_2_COST.Enabled = True
        FKS460_S01.Form.KKTPM_FZAPPI_3_CODE.Enabled = False
        FKS460_S01.Form.KKTPM_FZAPPI_3_COST.Enabled = False
        FKS460_S01.Form.KKTPM_EZAPPI_1_CODE.Enabled = True
        FKS460_S01.Form.KKTPM_EZAPPI_1_COST.Enabled = True
        FKS460_S01.Form.KKTPM_EZAPPI_2_CODE.Enabled = False
        FKS460_S01.Form.KKTPM_EZAPPI_2_COST.Enabled = False
        FKS460_S01.Form.KKTPM_EZAPPI_3_CODE.Enabled = False
        FKS460_S01.Form.KKTPM_EZAPPI_3_COST.Enabled = False
        FKS460_S01.Form.cmd_削除.Enabled = True
        'コンボボックス制御
        Call subSetMeisaiCmb
        '空白レコード無し
        FKS460_S01.Form.AllowEdits = True                '更新の許可
        FKS460_S01.Form.AllowAdditions = True            '追加の許可
        '説明文無し
        lbl_CheckJimu.Visible = False
        
        '明細域を元の大きさにする
        FKS460_S01.HEIGHT = 4712
        FKS460_S01.WIDTH = 19593
        FKS460_S01.Form.ScrollBars = 3
        
    Else
        '■事務手数料(明細域の制御)
        FKS460_S01.Form.KKTPM_PRICE.Enabled = False
        FKS460_S01.Form.KKTPM_FZAPPI_1_CHK.Enabled = True
        FKS460_S01.Form.KKTPM_FZAPPI_2_CODE.Enabled = False
        FKS460_S01.Form.KKTPM_FZAPPI_2_COST.Enabled = False
        FKS460_S01.Form.KKTPM_FZAPPI_3_CODE.Enabled = False
        FKS460_S01.Form.KKTPM_FZAPPI_3_COST.Enabled = False
        FKS460_S01.Form.KKTPM_EZAPPI_1_CODE.Enabled = False
        FKS460_S01.Form.KKTPM_EZAPPI_1_COST.Enabled = False
        FKS460_S01.Form.KKTPM_EZAPPI_2_CODE.Enabled = False
        FKS460_S01.Form.KKTPM_EZAPPI_2_COST.Enabled = False
        FKS460_S01.Form.KKTPM_EZAPPI_3_CODE.Enabled = False
        FKS460_S01.Form.KKTPM_EZAPPI_3_COST.Enabled = False
        FKS460_S01.Form.cmd_削除.Enabled = False
        'コンボボックス制御
        Call subSetMeisaiCmb
        '空白レコード有り
        FKS460_S01.Form.AllowEdits = True                '更新の許可
        FKS460_S01.Form.AllowAdditions = False           '追加の許可
        '説明文有り
        lbl_CheckJimu.Visible = True
        '明細域を元の大きさに小さくする
        FKS460_S01.HEIGHT = 655
        FKS460_S01.Form.ScrollBars = 0
        FKS460_S01.WIDTH = 9395
    End If
    FKS460_S01.Requery
End Sub

'==============================================================================*
'   MODULE_NAME :   明細行の作成(事務手数料の場合のみ)
'   MODULE_ID   :   subCreateMeisai
'   CREATE_DATE :   2018/08/11
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCreateMeisai()
    Dim strSQL          As String
    'delete
    Call del_FKS460_W01(0)
    '価格表の場合は1行いらない
    If cmb_テンプレートタイプ.VALUE = str価格表 Then
        Exit Sub
    End If
    'insert
    strSQL = "insert into FKS460_W01 "
    strSQL = strSQL & " ("
    strSQL = strSQL & " KKTPM_TCODE," '1
    strSQL = strSQL & " KKTPM_USAGE," '2
    strSQL = strSQL & " KKTPM_SIZE," '3
    strSQL = strSQL & " KKTPM_STEP," '4
    strSQL = strSQL & " KKTPM_PRICE," '5
    strSQL = strSQL & " KKTPM_FZAPPI_1_CODE," '6
    strSQL = strSQL & " KKTPM_FZAPPI_1_COST," '7
    strSQL = strSQL & " KKTPM_FZAPPI_2_CODE," '8
    strSQL = strSQL & " KKTPM_FZAPPI_2_COST," '9
    strSQL = strSQL & " KKTPM_FZAPPI_3_CODE," '10
    strSQL = strSQL & " KKTPM_FZAPPI_3_COST," '11
    strSQL = strSQL & " KKTPM_EZAPPI_1_CODE," '12
    strSQL = strSQL & " KKTPM_EZAPPI_1_COST," '13
    strSQL = strSQL & " KKTPM_EZAPPI_2_CODE," '14
    strSQL = strSQL & " KKTPM_EZAPPI_2_COST," '15
    strSQL = strSQL & " KKTPM_EZAPPI_3_CODE," '16
    strSQL = strSQL & " KKTPM_EZAPPI_3_COST " '17
    strSQL = strSQL & " ) values ("
    strSQL = strSQL & "'" & Me.cmb_テンプレートコード.VALUE & "'," '1
    strSQL = strSQL & "-1," '2
    strSQL = strSQL & "null," '3
    strSQL = strSQL & "-1," '4
    strSQL = strSQL & "null," '5
    strSQL = strSQL & "63," '6
    strSQL = strSQL & "0,"  '7
    strSQL = strSQL & "0,"  '8
    strSQL = strSQL & "0,"  '9
    strSQL = strSQL & "0,"  '10
    strSQL = strSQL & "0,"  '11
    strSQL = strSQL & "0,"  '12
    strSQL = strSQL & "0,"  '13
    strSQL = strSQL & "0,"  '15
    strSQL = strSQL & "0,"  '15
    strSQL = strSQL & "0,"  '16
    strSQL = strSQL & "0"  '17
    strSQL = strSQL & " )"
    CurrentDb.Execute strSQL, dbFailOnError
    FKS460_S01.Requery
End Sub

'==============================================================================*
'   MODULE_NAME :   明細のコンボボックス設定
'   MODULE_ID   :   subSetMeisaiCmb
'   CREATE_DATE :   2018/08/11
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetMeisaiCmb()
    Dim strSQL          As String
On Error GoTo Exception
    
    '段区分
    strSQL = ""
    If cmb_テンプレートタイプ.VALUE = str事務手数料 Then
        strSQL = strSQL & " SELECT TOP 1 -1,null FROM dbo_NAME_MAST"
        strSQL = strSQL & " UNION"
    End If
    strSQL = strSQL & " SELECT dbo_NAME_MAST.NAME_CODE, dbo_NAME_MAST.NAME_RYAK"
    strSQL = strSQL & " FROM dbo_NAME_MAST"
    strSQL = strSQL & " WHERE (((dbo_NAME_MAST.[NAME_ID])='090') AND ((dbo_NAME_MAST.[NAME_CODE])<>3))"
    strSQL = strSQL & " ORDER BY 1"
    FKS460_S01.Form.KKTPM_STEP.RowSource = strSQL
    '用途
    strSQL = ""
    If cmb_テンプレートタイプ.VALUE = str事務手数料 Then
        strSQL = strSQL & " SELECT TOP 1 -1 ,null FROM dbo_NAME_MAST"
        strSQL = strSQL & " UNION"
    End If
    strSQL = strSQL & " SELECT dbo_NAME_MAST.NAME_CODE, dbo_NAME_MAST.NAME_RYAK"
    strSQL = strSQL & " FROM dbo_NAME_MAST"
    strSQL = strSQL & " WHERE (((dbo_NAME_MAST.NAME_CODE)<>10 And (dbo_NAME_MAST.NAME_CODE)<>30) AND ((dbo_NAME_MAST.NAME_ID)='086'))"
    strSQL = strSQL & " ORDER BY 1"
    FKS460_S01.Form.KKTPM_USAGE.RowSource = strSQL
    '初回雑費
    If cmb_テンプレートタイプ.VALUE = str価格表 Then
        '価格表
        '初回雑費
        strSQL = ""
        strSQL = strSQL & " SELECT dbo_NAME_MAST.NAME_CODE, dbo_NAME_MAST.NAME_RYAK"
        strSQL = strSQL & " FROM dbo_NAME_MAST"
        strSQL = strSQL & " WHERE dbo_NAME_MAST.NAME_ID = '096'"
        strSQL = strSQL & " AND   dbo_NAME_MAST.NAME_CODE <> 63"
        strSQL = strSQL & " ORDER BY dbo_NAME_MAST.NAME_CODE"
        FKS460_S01.Form.KKTPM_FZAPPI_1_CODE.RowSource = strSQL
    Else
        '事務手数料
        strSQL = "SELECT dbo_NAME_MAST.NAME_CODE, dbo_NAME_MAST.NAME_RYAK"
        strSQL = strSQL & " FROM dbo_NAME_MAST"
        strSQL = strSQL & " WHERE dbo_NAME_MAST.NAME_ID = '096'"
        strSQL = strSQL & " AND   dbo_NAME_MAST.NAME_CODE = 63"
        strSQL = strSQL & " ORDER BY dbo_NAME_MAST.NAME_CODE"
        FKS460_S01.Form.KKTPM_FZAPPI_1_CODE.RowSource = strSQL
    End If

    FKS460_S01.Requery
    
    Exit Sub
Exception:
    Call MSZZ024_M00("subSetMeisaiCmb", True)
End Sub
'==============================================================================*
'   MODULE_NAME :   cmb_テンプレートコード_KeyPressイベント
'   MODULE_ID   :   cmb_テンプレートコード
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' キー入力時処理
Private Sub cmb_テンプレートコード_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Exception
    
    '入力値チェック
    Call chk_ASCII_CODE(cmb_テンプレートコード, KeyAscii)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmb_テンプレートコード_KeyPress", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   txt_テンプレート名_BeforeUpdateイベント
'   MODULE_ID   :   txt_テンプレート名
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_テンプレート名_BeforeUpdate(Cancel As Integer)
    
    Dim strTemp        As String
    Dim lngRTN          As Long
    
    On Error GoTo Exception
    
    strTemp = Nz(txt_テンプレート名, "")
    
    '何も入力されていない時は、処理をしない
    If "" = strTemp Then
        Exit Sub
    End If
    
    '入力チェック
    lngRTN = CmBCheck_Char(strTemp, 20)
    If 2 = lngRTN Then
        '桁数ｴﾗｰ
        MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
        Cancel = True
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("txt_テンプレート名_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   txt_テンプレート名_AfterUpdateイベント
'   MODULE_ID   :   txt_テンプレート名
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_テンプレート名_AfterUpdate()
    
    On Error GoTo Exception
    
    '編集中フラグを立てる
    set_ModifyFlg (True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("txt_テンプレート名_AfterUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   txt_キーワード地区_BeforeUpdateイベント
'   MODULE_ID   :   txt_キーワード地区
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_キーワード地区_BeforeUpdate(Cancel As Integer)
    
    Dim strTemp         As String
    Dim lngRTN          As Long
    
    On Error GoTo Exception
    
    strTemp = Nz(txt_キーワード地区, "")
    
    '何も入力されていない時は、処理をしない
    If "" = strTemp Then
        Exit Sub
    End If
    
    '入力チェック
    lngRTN = CmBCheck_Char(strTemp, 60)
    If 2 = lngRTN Then
        '桁数ｴﾗｰ
        MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
        Cancel = True
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("txt_キーワード地区_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   txt_キーワード地区_AfterUpdateイベント
'   MODULE_ID   :   txt_キーワード地区
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_キーワード地区_AfterUpdate()
    
    On Error GoTo Exception
    
    '編集中フラグを立てる
    set_ModifyFlg (True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("txt_キーワード地区_AfterUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   txt_キーワード価格_BeforeUpdateイベント
'   MODULE_ID   :   txt_キーワード価格
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_キーワード価格_BeforeUpdate(Cancel As Integer)
    
    Dim strTemp        As String
    Dim lngRTN          As Long
    
    On Error GoTo Exception
    
    strTemp = Nz(txt_キーワード価格, "")
    
    '何も入力されていない時は、処理をしない
    If "" = strTemp Then
        Exit Sub
    End If
    
    '入力チェック
    lngRTN = CmBCheck_Char(strTemp, 60)
    If 2 = lngRTN Then
        '桁数ｴﾗｰ
        MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
        Cancel = True
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("txt_キーワード価格_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   txt_キーワード価格_AfterUpdateイベント
'   MODULE_ID   :   txt_キーワード価格
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_キーワード価格_AfterUpdate()
    
    On Error GoTo Exception
    
    '編集中フラグを立てる
    set_ModifyFlg (True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("txt_キーワード価格_AfterUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmb_テンプレートコード2_BeforeUpdateイベント
'   MODULE_ID   :   cmb_テンプレートコード2
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' テンプレート変更時処理
Private Sub cmb_テンプレートコード2_BeforeUpdate(Cancel As Integer)
    Dim strKKTHM_TCODE  As String
    Dim intMode         As Integer
    Dim blnRTN          As Boolean
    
    On Error GoTo Exception
    
    '入力値取得
    strKKTHM_TCODE = Nz(cmb_テンプレートコード2.VALUE, "")
    intMode = opg_処理選択.VALUE
    
    '入力チェック
    If "" = strKKTHM_TCODE Then
        MsgBox "テンプレートコードを入力してください。", vbInformation, strMsg_kaku
        Exit Sub
    End If
    
    '桁数チェック
    If 6 < CmLenb(strKKTHM_TCODE) Then
        MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
        GoTo Exit_rtn
    End If
    
    '入力値の存在チェック
    blnRTN = chk_KKTH_MAST_EXIST(strKKTHM_TCODE, 2)
    
    '存在しない場合
    If Not blnRTN Then
        GoTo Exit_rtn
    End If
    
    Exit Sub
    
Exit_rtn:
    Cancel = True
    SendKeys "{ESC}"
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmb_テンプレートコード2_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmb_テンプレートコード2_AfterUpdateイベント
'   MODULE_ID   :   cmb_テンプレートコード2
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' テンプレート変更時処理
Private Sub cmb_テンプレートコード2_AfterUpdate()
    'コード変更時
    Dim strDML          As String
    
    Dim intMode         As Integer
    Dim strKKTHM_TCODE  As String
    
    Dim strKKTHM_TNAME  As String
    Dim strKKTHM_KEYAD  As String
    Dim strKKTHM_KEYPR  As String
    
    Dim rsWork          As Recordset
    Dim blnRTN          As Boolean
    
    On Error GoTo Exception
    
    '********************
    ' 初期処理
    '********************
    intMode = opg_処理選択.VALUE
    strKKTHM_TCODE = Nz(cmb_テンプレートコード2.VALUE, "")

    '入力値の存在チェック
    blnRTN = chk_KKTH_MAST_EXIST(strKKTHM_TCODE, 2)
    
    '存在しない場合
    If Not blnRTN Then
        MsgBox strMSGデータなし, vbInformation, strMsg_err
        Exit Sub
        
    End If
    
    '入力されたテンプレートコードの情報を取得・画面表示
    Call get_KKTH_MAST(cmb_テンプレートコード2, txt_テンプレート名2, txt_キーワード地区2, txt_キーワード価格2)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmb_テンプレートコード2_AfterUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmb_テンプレートコード2_KeyPressイベント
'   MODULE_ID   :   cmb_テンプレートコード2
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' キー入力時処理
Private Sub cmb_テンプレートコード2_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Exception
    
    '入力値チェック
    Call chk_ASCII_CODE(cmb_テンプレートコード2, KeyAscii)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmb_テンプレートコード2_KeyPress", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   cmd_内容複写_Clickイベント
'   MODULE_ID   :   cmd_内容複写
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'価格表テンプレートマスタ（価格）（内容複写）
Private Sub cmd_内容複写_Click()

    Dim strKKTHM_TCODE  As String
    Dim strKKTHM_TCODE2 As String
    Dim strDML          As String
    
    On Error GoTo Exception
    
    '****************************************
    ' 初期処理
    '****************************************
    Application.Echo False              '画面の描画を止める
    
    'コード取得
    strKKTHM_TCODE = Nz(cmb_テンプレートコード.VALUE, "")
    strKKTHM_TCODE2 = Nz(cmb_テンプレートコード2.VALUE, "")
    
    'テンプレートコード入力の確認
    If "" = strKKTHM_TCODE2 Then
        MsgBox "テンプレートコードが入力されていません。", vbInformation, strMsg_err
        cmb_テンプレートコード2.SetFocus
        GoTo Exit_Sub
    End If
    
    '****************************************
    ' テンプレート名、キーワードをコピー
    '****************************************
    txt_テンプレート名.VALUE = txt_テンプレート名2.VALUE
    txt_キーワード地区.VALUE = txt_キーワード地区2.VALUE
    txt_キーワード価格.VALUE = txt_キーワード価格2.VALUE
    
    '****************************************
    ' ワークテーブルデータ作成
    '****************************************
    Call cre_FKS460_W01(strKKTHM_TCODE2)
    
    '編集中フラグを立てる
    set_ModifyFlg (True)
    
Exit_Sub:
    Application.Echo True               '画面の描画を行う
    Exit Sub
    
Exception:
    Application.Echo True               '画面の描画を行う
    'メッセージ・ログ出力
    Call MSZZ024_M00("cmd_内容複写_Click", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME     :   ボタンの切り替え
'   Parameter       :   intMODE（ 1=新規、2=修正、3=削除 ）
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub set_CtrlButton(ByVal intMode As Integer)
    
    On Error GoTo Exception
    
    If 3 = intMode Then
        cmd_保存.Enabled = False
        cmd_削除.Enabled = True
    Else
        cmd_保存.Enabled = True
        cmd_削除.Enabled = False
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "set_CtrlButton" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'   MODULE_NAME     :   価格表テンプレートマスタ（ヘッダ）データ削除
'   Parameter       :   strTCODE    テンプレートコード
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub del_TH(ByVal objCon As Object, ByVal strTcode As String)
    'SQL文
    Dim strDML      As String
    Dim lngCnt      As Long
    
    On Error GoTo Exception
    
    'SQL編集 - 削除
    strDML = "DELETE FROM KKTH_MAST WHERE KKTHM_TCODE = '" & strTcode & "'"
    lngCnt = ADODB_Execute(strDML, objCon)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "del_TH" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'   MODULE_NAME     :   価格表テンプレートマスタ（価格）データ削除
'   Parameter       :   strTCODE    テンプレートコード
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub del_TP(ByVal objCon As Object, ByVal strTcode As String)
    'SQLサーバー用
    Dim strDML      As String
    Dim lngCnt      As Long
    
    On Error GoTo Exception
    
    'SQL編集 - 削除
    strDML = "DELETE FROM KKTP_MAST WHERE KKTPM_TCODE = '" & strTcode & "'"
    lngCnt = ADODB_Execute(strDML, objCon)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "del_TP" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'   MODULE_NAME     :   ワークテーブルデータ削除
'   Parameter       :   lngFLG    再描画    0=しない、1=する
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub del_FKS460_W01(ByVal lngFLG As Long)
    
    On Error GoTo Exception
    
    '警告非表示
    doCmd.SetWarnings False
    
    'SQL実行
    doCmd.RunSQL "DELETE FROM FKS460_W01"
    
    '警告表示
    doCmd.SetWarnings True
    
    If 1 = lngFLG Then
        '再描画
        FKS460_S01.Requery
        
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "del_FKS460_W01" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'   MODULE_NAME     :   ワークテーブルデータ作成
'   Parameter       :   strTCODE    テンプレートコード
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'価格表テンプレートマスタ（価格）から、ワークテーブルにコピー
Private Sub cre_FKS460_W01(ByVal strTcode As String)
    'SQLサーバー用
    Dim objCon              As Object
    Dim objKKTP             As Object       'Recordset
    
    'アクセスワークテーブル用
    Dim dbAcc               As Database
    Dim rsWork              As Recordset
    
    'SQL文
    Dim strDML              As String
    
    '一時変数
    Dim lngUSAGE            As Long
    Dim dblSIZE             As Double
    Dim lngSTEP             As Long
    Dim lngPrice            As Long
    Dim lngFZAPPI_1_CODE    As Long
    Dim lngFZAPPI_1_COST    As Long
    Dim lngFZAPPI_2_CODE    As Long
    Dim lngFZAPPI_2_COST    As Long
    Dim lngFZAPPI_3_CODE    As Long
    Dim lngFZAPPI_3_COST    As Long
    Dim lngEZAPPI_1_CODE    As Long
    Dim lngEZAPPI_1_COST    As Long
    Dim lngEZAPPI_2_CODE    As Long
    Dim lngEZAPPI_2_COST    As Long
    Dim lngEZAPPI_3_CODE    As Long
    Dim lngEZAPPI_3_COST    As Long
    
    On Error GoTo Exception
    
    '****************************************
    '初期処理
    '****************************************
    ' ワークテーブルを削除
    Call del_FKS460_W01(1)
    
    Set objCon = MSZZ025.ADODB_Connection(pstrBumoCD)
    Set dbAcc = CurrentDb
    
    '****************************************
    'SQLサーバーからデータを取得
    '****************************************
    'SQL文編集 - 検索
    strDML = "select * from KKTP_MAST where KKTPM_TCODE = '" & strTcode & "'"
    Set objKKTP = MSZZ025.ADODB_Recordset(strDML, objCon)
    
    '****************************************
    'アクセスワークテーブルにデータを挿入
    '****************************************
    'SQL文編集 - 検索
    Set rsWork = dbAcc.OpenRecordset("FKS460_W01", dbOpenDynaset, dbAppendOnly)
    
    '警告非表示
    doCmd.SetWarnings False
    
    Do Until objKKTP.EOF
        
        'SQLサーバーのデータ取得
        
        '2018/08/11 空白は-1として扱う
        'lngUSAGE = objKKTP.Fields("KKTPM_USAGE") '*
        lngUSAGE = Nz(objKKTP.Fields("KKTPM_USAGE"), -1)
        
        '2018/08/11 空白は-1として扱う
        'dblSIZE = objKKTP.Fields("KKTPM_SIZE")
        dblSIZE = Nz(objKKTP.Fields("KKTPM_SIZE"), -1)
        
        '2018/08/11 空白は-1として扱う
        'lngSTEP = objKKTP.Fields("KKTPM_STEP") '*
        lngSTEP = Nz(objKKTP.Fields("KKTPM_STEP"), -1)
        
        '2018/08/11 空白は-1として扱う
        'lngPRICE = objKKTP.Fields("KKTPM_PRICE")
        lngPrice = Nz(objKKTP.Fields("KKTPM_PRICE"), -1)
        
        lngFZAPPI_1_CODE = objKKTP.Fields("KKTPM_FZAPPI_1_CODE")
        lngFZAPPI_1_COST = objKKTP.Fields("KKTPM_FZAPPI_1_COST")
        lngFZAPPI_2_CODE = objKKTP.Fields("KKTPM_FZAPPI_2_CODE")
        lngFZAPPI_2_COST = objKKTP.Fields("KKTPM_FZAPPI_2_COST")
        lngFZAPPI_3_CODE = objKKTP.Fields("KKTPM_FZAPPI_3_CODE")
        lngFZAPPI_3_COST = objKKTP.Fields("KKTPM_FZAPPI_3_COST")
        lngEZAPPI_1_CODE = objKKTP.Fields("KKTPM_EZAPPI_1_CODE")
        lngEZAPPI_1_COST = objKKTP.Fields("KKTPM_EZAPPI_1_COST")
        lngEZAPPI_2_CODE = objKKTP.Fields("KKTPM_EZAPPI_2_CODE")
        lngEZAPPI_2_COST = objKKTP.Fields("KKTPM_EZAPPI_2_COST")
        lngEZAPPI_3_CODE = objKKTP.Fields("KKTPM_EZAPPI_3_CODE")
        lngEZAPPI_3_COST = objKKTP.Fields("KKTPM_EZAPPI_3_COST")
        
        'ワークテーブルに書き込み
        With rsWork
            '行追加
            .AddNew
            
            'データ設定
            .Fields("KKTPM_TCODE") = strTcode                   'テンプレートコード
            .Fields("KKTPM_USAGE") = lngUSAGE                   '使用用途
            
            '2018/08/11 空白は-1として扱う
            '.Fields("KKTPM_SIZE") = dblSIZE
            .Fields("KKTPM_SIZE") = IIf(dblSIZE = -1, Null, dblSIZE) '実帖(整数３桁少数２桁)
            
            .Fields("KKTPM_STEP") = lngSTEP                     '段区分(0:上段、1:下段)
            
            '2018/08/11 空白は-1として扱う
            '.Fields("KKTPM_PRICE") = lngPRICE
            .Fields("KKTPM_PRICE") = IIf(lngPrice = -1, Null, lngPrice) '価格
            
            '事務手数料の場合は999時はチェックする 2018/08/11
            If cmb_テンプレートタイプ.VALUE = str事務手数料 Then
                If lngFZAPPI_1_COST = pcsglMoneyMaxValue Then
                    .Fields("KKTPM_FZAPPI_1_CHK") = True
                End If
            End If
            
            .Fields("KKTPM_FZAPPI_1_CODE") = lngFZAPPI_1_CODE   '初回雑費1コード
            .Fields("KKTPM_FZAPPI_1_COST") = lngFZAPPI_1_COST   '初回雑費1費用
            .Fields("KKTPM_FZAPPI_2_CODE") = lngFZAPPI_2_CODE   '初回雑費2コード
            .Fields("KKTPM_FZAPPI_2_COST") = lngFZAPPI_2_COST   '初回雑費2費用
            .Fields("KKTPM_FZAPPI_3_CODE") = lngFZAPPI_3_CODE   '初回雑費3コード
            .Fields("KKTPM_FZAPPI_3_COST") = lngFZAPPI_3_COST   '初回雑費3費用
            .Fields("KKTPM_EZAPPI_1_CODE") = lngEZAPPI_1_CODE   '毎月雑費1コード
            .Fields("KKTPM_EZAPPI_1_COST") = lngEZAPPI_1_COST   '毎月雑費1費用
            .Fields("KKTPM_EZAPPI_2_CODE") = lngEZAPPI_2_CODE   '毎月雑費2コード
            .Fields("KKTPM_EZAPPI_2_COST") = lngEZAPPI_2_COST   '毎月雑費2費用
            .Fields("KKTPM_EZAPPI_3_CODE") = lngEZAPPI_3_CODE   '毎月雑費3コード
            .Fields("KKTPM_EZAPPI_3_COST") = lngEZAPPI_3_COST   '毎月雑費3費用
            
            '更新
            .UPDATE
        End With
        
        '次のデータ
        objKKTP.MoveNext
    Loop
    
    objKKTP.Close
    rsWork.Close
    
    '警告表示
    doCmd.SetWarnings True
    
    Set rsWork = dbAcc.OpenRecordset("SELECT * FROM FKS460_W01", dbOpenDynaset)
    
    '再描画
    FKS460_S01.Requery
    
    Set objCon = Nothing
    Set dbAcc = Nothing
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "cre_FKS460_W01" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
    'レコードセットクリア
    If Not objCon Is Nothing Then rsWork.Close: Set objCon = Nothing
    If Not dbAcc Is Nothing Then objKKTP.Close: Set dbAcc = Nothing
    
End Sub

'==============================================================================*
'   MODULE_NAME     :   オブジェクトの使用可否制御
'   Parameter       :   blnFLG    使用可能=true、使用不可=false
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub set_ENABLED(ByVal blnFLG As Boolean)
    Dim intMode     As Integer
    Dim ctlCur      As Control
    
    On Error GoTo Exception
    
    intMode = opg_処理選択.VALUE

    '使用可否制御対象オブジェクト
    If 3 = intMode Then
        txt_テンプレート名.Enabled = False
        txt_キーワード地区.Enabled = False
        txt_キーワード価格.Enabled = False
        cmb_テンプレートコード2.Enabled = False
        txt_テンプレート名2.Enabled = False
        txt_キーワード地区2.Enabled = False
        txt_キーワード価格2.Enabled = False
        cmd_内容複写.Enabled = False
        FKS460_S01.Enabled = False
        
    Else
        txt_テンプレート名.Enabled = True
        txt_キーワード地区.Enabled = True
        txt_キーワード価格.Enabled = True
        cmb_テンプレートコード2.Enabled = blnFLG
        txt_テンプレート名2.Enabled = blnFLG
        txt_キーワード地区2.Enabled = blnFLG
        txt_キーワード価格2.Enabled = blnFLG
        cmd_内容複写.Enabled = blnFLG
        FKS460_S01.Enabled = blnFLG
        
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "set_ENABLED" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub
'==============================================================================*
'   MODULE_NAME     :   ワークテーブルデータ作成
'   Parameter       :   strTCODE    テンプレートコード
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub clr_SCREEN()
    
    On Error GoTo Exception
    
    '画面の項目をクリア
    
    '2018/08/11 テンプレートタイプの設定
    cmb_テンプレートタイプ.VALUE = IIf(pstrBumoCD = "H", str価格表, str事務手数料)
    
    '入力項目制御 2018/08/11
    Call subInputSwKakakuJimu
    
    cmb_テンプレートコード.VALUE = ""
    txt_テンプレート名.VALUE = ""
    txt_キーワード地区.VALUE = ""
    txt_キーワード価格.VALUE = ""
    
    cmb_テンプレートコード2.VALUE = ""
    txt_テンプレート名2.VALUE = ""
    txt_キーワード地区2.VALUE = ""
    txt_キーワード価格2.VALUE = ""
    
    'ワークテーブルを削除
    Call del_FKS460_W01(1)
    
    '
    cmb_テンプレートコード.Requery
    cmb_テンプレートコード2.Requery

'2018/08/11 tabindex0が優先されるためコメントアウト
'    FKS460_S01.Form.KKTPM_STEP.SetFocus
        
    '編集不可項目
    Call set_ENABLED(False)
    
    '保持データクリア
    blnModifyFlg = False        '画面データ編集の有無
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "clr_SCREEN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'   MODULE_NAME     :   価格表テンプレートマスタ（ヘッダ）存在チェック
'   Parameter       :   code（ テンプレートコード ）
'   Return Value    :   True=存在する、False=存在しない
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function chk_KKTH_MAST_EXIST(ByVal strTcode As String, ByVal intFlg As Integer) As Boolean

    Dim intMode     As Integer
    Dim varRet      As Variant

    On Error GoTo Exception

    chk_KKTH_MAST_EXIST = False
    
    intMode = opg_処理選択.VALUE
    
    varRet = Nz(DLookup("KKTHM_TCODE", "dbo_KKTH_MAST", " KKTHM_TCODE = '" & strTcode & "'"), "")
    
    If 1 = intFlg Then
        Select Case intMode
            Case 1
                '新規モードでデータが在る場合
                If "" <> varRet Then
                    MsgBox strmsg重複, vbInformation, strMsg_err
                    GoTo Exit_rtn
                End If
    
            Case 2, 3
                '修正・削除モードでデータが無い場合
                If "" = varRet Then
                    MsgBox strMSGデータなし, vbInformation, strMsg_err
                    GoTo Exit_rtn
                End If
        End Select
    ElseIf 2 = intFlg Then
        'コピー元検索でデータが無い場合
        If "" = varRet Then
            MsgBox strMSGデータなし, vbInformation, strMsg_err
            GoTo Exit_rtn
        End If
    Else
        'エラー
        GoTo Exit_rtn
    End If
    
    chk_KKTH_MAST_EXIST = True
    
Exit_rtn:
    Exit Function
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "chk_KKTH_MAST_EXIST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Function

'==============================================================================*
'   MODULE_NAME     :   ワークテーブル存在チェック
'   Parameter       :   code（ テンプレートコード ）
'   Return Value    :   True=存在する、False=存在しない
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function chk_FKS460_W01_EXIST(ByVal code As String) As Boolean

    Dim strWhere As String
    Dim vRet     As Variant

    On Error GoTo Exception

    chk_FKS460_W01_EXIST = False
    
    vRet = Nz(DLookup("KKTPM_TCODE", "FKS460_W01", " KKTPM_TCODE = '" & code & "'"), "")
    
    If "" <> vRet Then chk_FKS460_W01_EXIST = True

Exit_rtn:
    Exit Function

Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "chk_FKS460_W01_EXIST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Function

'==============================================================================*
'   MODULE_NAME     :   画面項目リセット
'   Parameter           txt_TMP_COD
'                       txt_TMP_MEI
'                       txt_KEY_ARE
'                       txt_KEY_PRI
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function get_KKTH_MAST(ByVal txt_TMP_COD As Object, ByVal txt_TMP_MEI As Object, ByVal txt_KEY_ARE As Object, ByVal txt_KEY_PRI As Object)

    Dim strDML          As String
    Dim strKKTHM_TCODE  As String
    Dim rsWork          As Recordset
    Dim strKKTHM_TNAME  As String
    Dim strKKTHM_KEYAD  As String
    Dim strKKTHM_KEYPR  As String
    
    On Error GoTo Exception
    
    '項目初期化
    txt_TMP_MEI.VALUE = ""
    txt_KEY_ARE.VALUE = ""
    txt_KEY_PRI.VALUE = ""
    
    'SQL文編集
    strKKTHM_TCODE = txt_TMP_COD.VALUE
    strDML = "select KKTHM_TNAME, KKTHM_KEYAD, KKTHM_KEYPR from dbo_KKTH_MAST where KKTHM_TCODE = '" & strKKTHM_TCODE & "'"
    
    '検索
    Set rsWork = CurrentDb().OpenRecordset(strDML, dbOpenSnapshot)
    
    'データ取得
    With rsWork
        If Not .EOF Then
            strKKTHM_TNAME = Trim(Nz(.Fields("KKTHM_TNAME")))
            strKKTHM_KEYAD = Trim(Nz(.Fields("KKTHM_KEYAD")))
            strKKTHM_KEYPR = Trim(Nz(.Fields("KKTHM_KEYPR")))
        End If
        .Close
    End With
    
    '取得したデータを画面にせってい
    txt_TMP_MEI.VALUE = strKKTHM_TNAME
    txt_KEY_ARE.VALUE = strKKTHM_KEYAD
    txt_KEY_PRI.VALUE = strKKTHM_KEYPR
    
    Exit Function
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "get_KKTH_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Function

'==============================================================================*
'   MODULE_NAME     :   ワークテーブル　重複/入力チェック
'   Parameter       :   なし
'   Return Value    :   0=正常終了、1=実帖エラー、2=価格エラー、3=重複エラー
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function chk_FKS460_W01() As Integer
    
    Dim strRTN      As String
    Dim strDML      As String
    Dim rsWork      As Recordset
    Dim intRtn      As Integer
    
    On Error GoTo Exception
    
    '******************************
    ' 初期設定
    '******************************
    '戻り値(エラーなし)の初期値を設定
    intRtn = 0
    
    '******************************
    ' 入力チェック
    '******************************
    'SQL編集して検索
    strRTN = Nz(DLookup("COUNT(*)", "FKS460_W01", "Nz(KKTPM_SIZE, 0) = 0"), "")
    
    '実帖チェック
    If 0 < CLng(strRTN) Then
        'エラーあり
        intRtn = 1
        GoTo Exit_rtn
    End If
    
    '価格チェック
    strRTN = Nz(DLookup("COUNT(*)", "FKS460_W01", "Nz(KKTPM_PRICE, 0) = 0"), "")
    If 0 < CLng(strRTN) Then
        'エラーあり
        intRtn = 2
        GoTo Exit_rtn
    End If
    
    '******************************
    ' 重複チェック
    '******************************
    'SQL編集して検索
    strDML = "SELECT CNT FROM ( SELECT COUNT(*) AS CNT FROM FKS460_W01 GROUP BY KKTPM_STEP, KKTPM_USAGE, KKTPM_SIZE ) WHERE CNT > 1"
    Set rsWork = CurrentDb().OpenRecordset(strDML, dbOpenSnapshot)
    
    '重複データが在る場合
    If Not rsWork.EOF Then
        intRtn = 3
    End If
    rsWork.Close
    
    '2018/08/11
    '初回雑費コード１＝"63" or 毎月雑費コード１="63" ならばチェック
    strRTN = Nz(DLookup("COUNT(*)", "FKS460_W01", "Nz(KKTPM_FZAPPI_1_CODE, 0) = 63 or Nz(KKTPM_EZAPPI_1_CODE, 0) = 63"), "")
    If 0 < CLng(strRTN) Then
        'エラーあり
        intRtn = 4
        GoTo Exit_rtn
    End If
    
    '2018/08/11
    '初回雑費コード２≠"63" ならばチェック
    strRTN = Nz(DLookup("COUNT(*)", "FKS460_W01", "Nz(KKTPM_FZAPPI_2_CODE, 0) <> 63"), "")
    If 0 < CLng(strRTN) Then
        'エラーあり
        intRtn = 5
        GoTo Exit_rtn
    End If
    
    '2018/08/11
    '初回雑費コード１、初回雑費コード２、毎月雑費コード１に同じコードが設定されていた場合※但し「0:無し」は重複可チェック
    strDML = "select count(*)"
    strDML = strDML & " from FKS460_W01"
    strDML = strDML & " where"
    strDML = strDML & " (KKTPM_FZAPPI_1_CODE = KKTPM_FZAPPI_2_CODE and KKTPM_FZAPPI_1_CODE <> 0 and KKTPM_FZAPPI_2_CODE <> 0)"
    strDML = strDML & " or"
    strDML = strDML & " (KKTPM_FZAPPI_1_CODE = KKTPM_EZAPPI_1_CODE and KKTPM_FZAPPI_1_CODE <> 0 and KKTPM_EZAPPI_1_CODE <> 0)"
    strDML = strDML & " or"
    strDML = strDML & " (KKTPM_FZAPPI_2_CODE = KKTPM_EZAPPI_1_CODE and KKTPM_FZAPPI_2_CODE <> 0 and KKTPM_EZAPPI_1_CODE <> 0)"
    '実行
    Set rsWork = CurrentDb().OpenRecordset(strDML, dbOpenSnapshot)
    '対象レコードがあればErr
    If rsWork.Fields(0) <> 0 Then
        intRtn = 6
    End If
    rsWork.Close
    
    If 3 = intRtn Then
        GoTo Exit_rtn
    End If
        
Exit_rtn:
    '戻り値を設定
    chk_FKS460_W01 = intRtn

    Exit Function

Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "chk_FKS460_W01" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Function

'==============================================================================*
'   MODULE_NAME     :   画面項目変更フラグ
'   Parameter       :   blnMod（ True：編集中、False：未編集 ）
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function set_ModifyFlg(ByVal blnMod As Boolean)
    
    Dim strText     As String
    
    On Error GoTo Exception
    
    '画面タイトル変更
    strText = strCaption
    
    If blnMod Then
        strText = strText & "(修正中)"
    End If
    
    blnModifyFlg = blnMod
    Caption = strText
    
    Exit Function
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "set_ModifyFlg" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Function

'==============================================================================*
'   MODULE_NAME     :   価格表テンプレートマスタ（ヘッダ）更新処理
'   Return Value    :   True：正常終了、False：異常終了
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function upd_THTP() As Boolean
    Dim objCon              As Object
    Dim dbAcc               As Database
    
    Dim strTNAME            As String
    Dim strKEYAD            As String
    Dim strKEYPR            As String
    Dim strDate             As String
    Dim strTime             As String
    Dim strDML              As String
    
    Dim strTtype            As String '2018/08/11
    Dim strTcode            As String
    Dim lngUSAGE            As Long
    Dim dblSIZE             As Double
    Dim lngSTEP             As Long
    Dim lngPrice            As Long
    Dim lngFZAPPI_1_CODE    As Long
    Dim lngFZAPPI_1_COST    As Long
    Dim lngFZAPPI_2_CODE    As Long
    Dim lngFZAPPI_2_COST    As Long
    Dim lngFZAPPI_3_CODE    As Long
    Dim lngFZAPPI_3_COST    As Long
    Dim lngEZAPPI_1_CODE    As Long
    Dim lngEZAPPI_1_COST    As Long
    Dim lngEZAPPI_2_CODE    As Long
    Dim lngEZAPPI_2_COST    As Long
    Dim lngEZAPPI_3_CODE    As Long
    Dim lngEZAPPI_3_COST    As Long
    
    Dim objKKTH             As Object
    Dim objKKTP             As Object
    Dim rsWork              As Object
    
    On Error GoTo Exception
    
    '****************************************
    '初期処理
    '****************************************
    Set objCon = MSZZ025.ADODB_Connection(pstrBumoCD)
    Set dbAcc = CurrentDb
    
    strTtype = Me.cmb_テンプレートタイプ '2018/08/11 テンプレートタイプ
    strTcode = cmb_テンプレートコード.VALUE
    strTNAME = txt_テンプレート名.VALUE
    strKEYAD = txt_キーワード地区.VALUE
    strKEYPR = txt_キーワード価格.VALUE
    strDate = Format(Now(), "yyyymmdd")
    strTime = Format(Now(), "hhmmss")
    
    'トランザクション開始
    objCon.BeginTrans
    
    '****************************************
    '価格表テンプレートマスタ（ヘッダ)
    '****************************************
    'SQL編集 - 検索
    strDML = "SELECT * FROM KKTH_MAST WHERE KKTHM_TCODE = '" & strTcode & "'"
    Set objKKTH = MSZZ025.ADODB_Recordset(strDML, objCon, adoReadWrite)
    
    '登録更新
    With objKKTH
        If .EOF Then
            '新規
            .AddNew
            .Fields("KKTHM_TCODE") = strTcode
            .Fields("KKTHM_INSED") = strDate
            .Fields("KKTHM_INSEJ") = strTime
            .Fields("KKTHM_INSPB") = strPROG_ID
            .Fields("KKTHM_INSUB") = strUSER_ID
        End If
        
        '2018/08/11 テンプレートタイプ
        .Fields("KKTHM_TTYPE") = strTtype
        
        .Fields("KKTHM_TNAME") = strTNAME
        .Fields("KKTHM_KEYAD") = strKEYAD
        .Fields("KKTHM_KEYPR") = strKEYPR
        .Fields("KKTHM_UPDAD") = strDate
        .Fields("KKTHM_UPDAT") = strTime
        .Fields("KKTHM_UPDPB") = strPROG_ID
        .Fields("KKTHM_UPDUB") = strUSER_ID
        
        .UPDATE
    End With
    objKKTH.Close
    Set objKKTH = Nothing
    
    '****************************************
    '価格表テンプレートマスタ(価格)
    '****************************************
    '価格表テンプレートマスタ(価格)の対象データを一旦削除
    Call del_TP(objCon, strTcode)
    
    'ワークテーブルのデータ取得
    ' - SQL編集
    strDML = "SELECT * FROM FKS460_W01"
    
    ' - 検索
    Set rsWork = dbAcc.OpenRecordset(strDML, dbOpenDynaset)
    
    '価格表テンプレートマスタ(価格)の枠を作成
    ' - 検索
    Set objKKTP = MSZZ025.ADODB_Recordset("KKTP_MAST", objCon, adoAppendOnly)       'adoAppendOnly, adoReadWrite
    
    'ワークテーブルのデータを価格表テンプレートマスタ(価格)に設定して更新
    Do Until rsWork.EOF
        'ワークテーブルのデータ取得
        
        '2018/08/11 NULL対応
        'lngUSAGE = rsWork.Fields("KKTPM_USAGE")
        'dblSIZE = rsWork.Fields("KKTPM_SIZE")
        'lngSTEP = rsWork.Fields("KKTPM_STEP")
        'lngPRICE = rsWork.Fields("KKTPM_PRICE")

        lngUSAGE = Nz(rsWork.Fields("KKTPM_USAGE"))
        dblSIZE = Nz(rsWork.Fields("KKTPM_SIZE"), -1)
        lngSTEP = Nz(rsWork.Fields("KKTPM_STEP"))
        lngPrice = Nz(rsWork.Fields("KKTPM_PRICE"), -1)
        
        lngFZAPPI_1_CODE = rsWork.Fields("KKTPM_FZAPPI_1_CODE")
        lngFZAPPI_1_COST = rsWork.Fields("KKTPM_FZAPPI_1_COST")
        lngFZAPPI_2_CODE = rsWork.Fields("KKTPM_FZAPPI_2_CODE")
        lngFZAPPI_2_COST = rsWork.Fields("KKTPM_FZAPPI_2_COST")
        lngFZAPPI_3_CODE = rsWork.Fields("KKTPM_FZAPPI_3_CODE")
        lngFZAPPI_3_COST = rsWork.Fields("KKTPM_FZAPPI_3_COST")
        lngEZAPPI_1_CODE = rsWork.Fields("KKTPM_EZAPPI_1_CODE")
        lngEZAPPI_1_COST = rsWork.Fields("KKTPM_EZAPPI_1_COST")
        lngEZAPPI_2_CODE = rsWork.Fields("KKTPM_EZAPPI_2_CODE")
        lngEZAPPI_2_COST = rsWork.Fields("KKTPM_EZAPPI_2_COST")
        lngEZAPPI_3_CODE = rsWork.Fields("KKTPM_EZAPPI_3_CODE")
        lngEZAPPI_3_COST = rsWork.Fields("KKTPM_EZAPPI_3_COST")
        
        'テンプレートマスタ(価格)に登録
        With objKKTP
            '行追加
            .AddNew
            
            'データ設定
            .Fields("KKTPM_TCODE") = strTcode                   'テンプレートコード
            
            '2018/08/11 空白は-1として扱う
            '.Fields("KKTPM_USAGE") = lngUSAGE                   '使用用途 *
            .Fields("KKTPM_USAGE") = IIf(lngUSAGE = -1, Null, lngUSAGE)
            
            '2018/08/11 空白は-1として扱う
            '.Fields("KKTPM_SIZE") = dblSIZE                     '実帖(整数３桁少数２桁)
            .Fields("KKTPM_SIZE") = IIf(dblSIZE = -1, Null, dblSIZE)
            
            '2018/08/11 空白は-1として扱う
            '.Fields("KKTPM_STEP") = lngSTEP                     '段区分(0:上段、1:下段) *
            .Fields("KKTPM_STEP") = IIf(lngSTEP = -1, Null, lngSTEP)
            
            '2018/08/11 空白は-1として扱う
            '.Fields("KKTPM_PRICE") = lngPRICE
            .Fields("KKTPM_PRICE") = IIf(lngPrice = -1, Null, lngPrice) '価格
            
            .Fields("KKTPM_FZAPPI_1_CODE") = lngFZAPPI_1_CODE   '初回雑費1コード
            .Fields("KKTPM_FZAPPI_1_COST") = lngFZAPPI_1_COST   '初回雑費1費用
            .Fields("KKTPM_FZAPPI_2_CODE") = lngFZAPPI_2_CODE   '初回雑費2コード
            .Fields("KKTPM_FZAPPI_2_COST") = lngFZAPPI_2_COST   '初回雑費2費用
            .Fields("KKTPM_FZAPPI_3_CODE") = lngFZAPPI_3_CODE   '初回雑費3コード
            .Fields("KKTPM_FZAPPI_3_COST") = lngFZAPPI_3_COST   '初回雑費3費用
            .Fields("KKTPM_EZAPPI_1_CODE") = lngEZAPPI_1_CODE   '毎月雑費1コード
            .Fields("KKTPM_EZAPPI_1_COST") = lngEZAPPI_1_COST   '毎月雑費1費用
            .Fields("KKTPM_EZAPPI_2_CODE") = lngEZAPPI_2_CODE   '毎月雑費2コード
            .Fields("KKTPM_EZAPPI_2_COST") = lngEZAPPI_2_COST   '毎月雑費2費用
            .Fields("KKTPM_EZAPPI_3_CODE") = lngEZAPPI_3_CODE   '毎月雑費3コード
            .Fields("KKTPM_EZAPPI_3_COST") = lngEZAPPI_3_COST   '毎月雑費3費用
            .Fields("KKTPM_INSED") = strDate                    '作成日付
            .Fields("KKTPM_INSEJ") = strTime                    '作成時刻
            .Fields("KKTPM_INSPB") = strPROG_ID                 '作成プログラムID
            .Fields("KKTPM_INSUB") = strUSER_ID                 '作成ユーザーID
            .Fields("KKTPM_UPDAD") = strDate                    '更新日時
            .Fields("KKTPM_UPDAT") = strTime                    '更新時刻
            .Fields("KKTPM_UPDPB") = strPROG_ID                 '更新プログラムID
            .Fields("KKTPM_UPDUB") = strUSER_ID                 '更新ユーザーID
            
            '更新
            .UPDATE
        End With
        
        '次のデータ
        rsWork.MoveNext
    Loop
    
    'コミット
    objCon.CommitTrans
    
    '終了処理
    If Not objKKTP Is Nothing Then objKKTP.Close: Set objKKTP = Nothing
    If Not rsWork Is Nothing Then rsWork.Close: Set rsWork = Nothing
    If Not dbAcc Is Nothing Then dbAcc.Close: Set dbAcc = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    
    upd_THTP = True
    
    Exit Function
    
Exception:
    'ロールバック
    objCon.RollbackTrans
    
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "upd_THTP" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
    'レコードセットクリア
    If Not objKKTH Is Nothing Then objKKTH.Close: Set objKKTH = Nothing
    If Not objKKTP Is Nothing Then objKKTP.Close: Set objKKTP = Nothing
    If Not rsWork Is Nothing Then rsWork.Close: Set rsWork = Nothing
    If Not dbAcc Is Nothing Then dbAcc.Close: Set dbAcc = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    
End Function

'==============================================================================*
'   MODULE_NAME     :   小数入力項目の入力チェック
'   Parameter       :   ctlTxt      = 対象となるテキストボックス
'                       varMaxValue = 最大値
'                       intLength   = 長さ
'   Return Value    :   True：正常終了、False：異常終了
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncTxtInputCheck(ByRef ctlTxt As TextBox, _
                                 ByVal varMaxValue As Variant, _
                                 ByVal intLength As Integer, _
                                 Optional ByVal intKbn As Integer = 0) As Boolean
    
    On Error GoTo Exception
    
    fncTxtInputCheck = False
    
    If IsNull(ctlTxt) Then
        fncTxtInputCheck = True
        Exit Function
    End If

    Select Case intKbn
        Case 0
            If (ctlTxt > varMaxValue Or _
                Len(Replace(Replace(ctlTxt, ".", ""), "-", "")) > intLength) Then
                Exit Function
            End If

        Case 1
            ' *** フォームのエラーイベントでのエラートラップのとき
            If (ctlTxt > varMaxValue Or _
                Len(Replace(Replace(ctlTxt.Text, ".", ""), "-", "")) > intLength) Then
                Exit Function
            End If
    End Select
    
    fncTxtInputCheck = True
    
    Exit Function
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "fncTxtInputCheck" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Function

'==============================================================================*
'   MODULE_NAME     :   文字項目のチェックを一括でする。
'   Parameter       :   oText   = 対象となるテキスト
'                       iLen    = 長さ
'   Return Value    :   0：正常終了、2：桁数が溢れ、9：それ以外のエラー
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function CmBCheck_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer
    
    On Error GoTo Exception
    
    CmBCheck_Char = 9
    
    If CmLenb(oText) > ilen Then
      CmBCheck_Char = 2
      GoTo Exit_rtn
      
    End If
    
    CmBCheck_Char = 0
    
Exit_rtn:
    Exit Function
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "CmBCheck_Char" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Function

'==============================================================================*
'   MODULE_NAME     :   押されたキーボードのキーのアスキーコードをチェック
'   Parameter       :   obj         = 対象となるオブジェクト
'                       KeyAscii    = アスキーコード
'   Return Value    :   0：正常終了、2：桁数が溢れ、9：それ以外のエラー
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub chk_ASCII_CODE(ByVal obj As Object, ByRef KeyAscii As Integer)
    
    On Error GoTo Exception
    
    ' "-, 0-9, A-Z, a-z, _" 以外は無効にする。
    Select Case KeyAscii
        Case 8, 9, 10, 13           '8(BS), 9(TAB), 10(LF), 13(CR), 127(DEL)
        Case 45 To 95, 65 To 90, 97 To 122               '0-9, a-z, A-Z
            '桁数制限
            With obj
                If CmLenb(.Text) > (6 - 1) And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            End With
            
        Case Else
            KeyAscii = 0            'キー入力無効
            
    End Select
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "chk_ASCII_CODE" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

