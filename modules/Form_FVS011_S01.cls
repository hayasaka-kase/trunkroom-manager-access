VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS011_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　空状況一覧サブ
'   プログラムＩＤ　：　FVS011_S01
'   作　成　日　　　：  2005/07/28
'   作　成　者　　　：  K.ISHIZAKA
'**********************************************
'修正履歴
'   修　正　日　　　：  20050829
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.1
'   修　正　内　容　：  解約予定本数/取り置き本数を追加
'**********************************************
'   修　正　日　　　：  2005/11/17
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.2
'   修　正　内　容　：  デザイン項目追加
'**********************************************
'   修　正　日　　　：  2005/12/05
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.3
'   修　正　内　容　：  デザイン変更
'**********************************************
'   修　正　日　　　：  2006/03/31
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.8
'   修　正　内　容　：  1.地図表示対応
'**********************************************
'   修　正　日　　　：  2006/08/31
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.9
'   修　正　内　容　：  ・ヤード非表示日を過ぎたヤードは選択リストから外す
'                       ・営業終了日を過ぎたヤードは検索対象外
'                       ・営業開始日を過ぎていなければ一覧上に表示する
'                       ・営業終了日が入力されていれば一覧上に表示する
'                       ・地主解約のチェックを外す
'                       ・コンテナ種別=0の条件をはずす
'**********************************************
'   修　正　日　　　：  2009/12/02
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  1.0
'   修　正　内　容　：  1.Noを廃止し管理部門を追加
'                       2.統合検索の時は管理部門により色を変える
'
'**********************************************
Option Compare Database
Option Explicit
'==============================================================================*
'   イベント関数
'==============================================================================*
Private Const C_GREEN_COLOR         As Long = 8454016                           'INSERT 2009/12/02 K.ISHIZAKA

'==============================================================================*
'
'       MODULE_NAME     : 詳細ボタンが押されたとき
'       MODULE_ID       : cmd_詳細_Click
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_詳細_Click()
    Dim strKey              As String                                           'INSERT 2009/12/02 K.ISHIZAKA
    Dim StrYard             As String                                           'INSERT 2009/12/02 K.ISHIZAKA
    
    strKey = Me.txt_Key.VALUE                                                   'INSERT 2009/12/02 K.ISHIZAKA
    StrYard = Format(Val(Me.txt_ヤードコード.VALUE))                            'INSERT 2009/12/02 K.ISHIZAKA
    '一覧の指定ヤードだけ最新に更新
'    If Parent.Reload(Me.txt_ヤードコード, "") = False Then                     'DELETE 2009/12/02 K.ISHIZAKA
    If Parent.Reload(strKey, StrYard, "") = False Then                          'INSERT 2009/12/02 K.ISHIZAKA
        Exit Sub
    End If
'    If Val(Me.txt_空合計) > 0 Then                                             'DELETE 2009/12/02 K.ISHIZAKA
    If Val(DLookup("SUM_AKI", "WK_FVS011_TEMP", "KEY='" & strKey & "' AND YARD_CODE=" & StrYard)) > 0 Then  'INSERT 2009/12/02 K.ISHIZAKA
'        Call Parent.ShowFVS021(Me.txt_ヤードコード)                            'DELETE 2009/12/02 K.ISHIZAKA
        Call Parent.ShowFVS021(strKey, StrYard)                                 'INSERT 2009/12/02 K.ISHIZAKA
    Else
        MsgBox "空きがなくなりました。", vbOKOnly, "確認"
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームキー押下時処理
'       MODULE_ID       : Form_KeyDown
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : KeyCode           キーコード
'                       : Shift             シフトキー状態
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Parent.KeyDown(KeyCode, Shift)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームロード
'       MODULE_ID       : Form_Load
'       CREATE_DATE     : 2009/12/02            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    Dim ctrl            As Control
    On Error GoTo ErrorHandler

    If Not fncFormIsLoaded("FKS000") Then
        txt_NO.Visible = False
        lbl_No.Visible = False
        txt_Key.Visible = True
        lbl_Key.Visible = True
        For Each ctrl In Me.Controls
            If UCase(TypeName(ctrl)) = "TEXTBOX" Then
                If ctrl.Visible Then
                    With ctrl.FormatConditions
                        .Delete
                        With .Add(acExpression, , "[txt_Key]='8'")
                            .BackColor = C_GREEN_COLOR
                            .FontBold = True
                            .ForeColor = ctrl.ForeColor
                        End With
                    End With
                End If
            End If
        Next
    End If
Exit Sub
    
ErrorHandler:       '↓自分の関数名
    Call MSZZ024_M00("Form_Load", True)
End Sub

'****************************  ended or program ********************************
