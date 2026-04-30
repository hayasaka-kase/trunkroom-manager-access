VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FTG011_S02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : コンテナ管理システム
'       SUB_SYSTEM_NAME : 統合検索
'
'       PROGRAM_NAME    : 統合空状況 条件詳細
'       PROGRAM_ID      : FTG011_S02
'       PROGRAM_KBN     : FROM
'
'       CREATE          : 2011/02/12
'       CERATER         : K.ISHZIAKA
'       Ver             : 0.0
'
'       UPDATE          : 2011/03/16
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.3
'                       : レンタル用途の関連付け検索を削除
'                         サブフォームFTG011_S03を削除
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'
'       MODULE_NAME     : キーが押されたとき
'       MODULE_ID       : chk_Fuzzy_KeyPress
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyAscii          アスキーコード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_Fuzzy_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler
    
    Select Case KeyAscii
    Case 12288, vbKeySpace
        chk_Fuzzy.VALUE = Not chk_Fuzzy.VALUE
    End Select
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("chk_Fuzzy_KeyPress", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キーが押されたとき
'       MODULE_ID       : chk_Sort_Service_KeyPress
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyAscii          アスキーコード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_Sort_Service_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler
    
    Select Case KeyAscii
    Case 12288, vbKeySpace
        chk_Sort_Service.VALUE = Not chk_Sort_Service.VALUE
    End Select
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("chk_Sort_Service_KeyPress", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キーが押されたとき
'       MODULE_ID       : chk_Sort_YardCode_KeyPress
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyAscii          アスキーコード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_Sort_YardCode_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler
    
    Select Case KeyAscii
    Case 12288, vbKeySpace
        chk_Sort_YardCode.VALUE = Not chk_Sort_YardCode.VALUE
    End Select
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("chk_Sort_YardCode_KeyPress", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : あいまい検索更新ボタン
'       MODULE_ID       : cmd_FuzzyUpdate_Click
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_FuzzyUpdate_Click()
    On Error GoTo ErrorHandler

    Call setFuzzyChar(txt_FuzzyChar.VALUE)
    MsgBox "更新しました。", vbOKOnly + vbInformation, Parent.Caption
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmd_FuzzyUpdate_Click", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キー入力時処理
'       MODULE_ID       : txt_MaxListCount_KeyPress
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyAscii          アスキーコード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaxListCount_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler
    Call CmKeyPress_For_Num1(txt_MaxListCount, KeyAscii, 3)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("txt_MaxListCount_KeyPress", True)
End Sub

'****************************  ended of program ********************************
