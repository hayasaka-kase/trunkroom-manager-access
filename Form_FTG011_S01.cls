VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FTG011_S01"
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
'       PROGRAM_NAME    : 統合空状況 空状況一覧
'       PROGRAM_ID      : FTG011_S01
'       PROGRAM_KBN     : FROM
'
'       CREATE          : 2011/02/12
'       CERATER         : K.ISHZIAKA
'       Ver             : 0.0
'
'       UPDATE          : 2011/03/10
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.2
'                       : レンタル用途の書式設定を変更
'                       : 次画面への引数「サイズ条件」を追加
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'
'       MODULE_NAME     : 詳細ボタンが押されたとき
'       MODULE_ID       : cmd_Syousai_Click
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Syousai_Click()
    Dim strYcode            As String
    Dim strUsage            As String
    On Error GoTo ErrorHandler

    strYcode = Format(Val(txt_YARD_CODE.VALUE))
    strUsage = Format(Val(txt_YARD_USAGE.VALUE))
    '一覧の指定ヤードだけ最新に更新
    If Not Parent.Reload(strYcode, strUsage, "", "") Then
        MsgBox "データの再取得に失敗しました。" & vbCrLf & "再度、検索しなおしてください。", vbOKOnly + vbExclamation, Parent.Caption
        Exit Sub
    End If
    
    '' 2011/07/27 M.HONDA START
    '' If Val(DLookup("SUM_AKI", Me.RecordSource, "YARD_CODE=" & strYcode & " AND YARD_USAGE=" & strUsage)) > 0 Then
    If Val(DLookup("SUM_AKI", Me.RecordSource, "YARD_CODE=" & strYcode & " AND YARD_USAGE=" & strUsage)) > 0 Or _
    Val(DLookup("KAIYAKU", Me.RecordSource, "YARD_CODE=" & strYcode & " AND YARD_USAGE=" & strUsage)) > 0 Then
    '' 2011/07/27 M.HONDA END
        If fncFormIsLoadedMsg("FTG021", Me.Caption) Then
            Exit Sub
        End If
        '引数は ヤードコード、レンタル用途、接続部門コード、サイズ条件
'        doCmd.OpenForm "FTG021", , , , , , Join(Array(strYcode, strUsage, txt_YARD_BUMOC.VALUE), ",") 'DELETE 2011/03/10 K.ISHIZAKA
        doCmd.OpenForm "FTG021", , , , , , Join(Array(strYcode, strUsage, txt_YARD_BUMOC.VALUE, Parent.strSizeWhere), ",") 'INSERT 2011/03/10 K.ISHIZAKA
    Else
        MsgBox "空きがありません。", vbOKOnly + vbInformation, Parent.Caption
    End If
Exit Sub

ErrorHandler:       '↓自分の関数名
    Call MSZZ024_M00("cmd_Syousai_Click", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームキー押下時処理
'       MODULE_ID       : Form_KeyDown
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyCode           キーコード
'                       : Shift             シフトキー状態
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
    Call Parent.KeyDown(KeyCode, Shift)
Exit Sub

ErrorHandler:       '↓自分の関数名
    Call MSZZ024_M00("Form_KeyDown", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームロード
'       MODULE_ID       : Form_Load
'       CREATE_DATE     : 2011/02/12            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Call setUsageBackColor(txt_YARD_USAGE, Me.Controls)
    Me.OrderBy = "NO"
    Me.OrderByOn = True
Exit Sub

ErrorHandler:       '↓自分の関数名
    Call MSZZ024_M00("Form_Load", True)
End Sub

'****************************  ended of program ********************************
