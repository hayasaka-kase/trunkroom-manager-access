Attribute VB_Name = "MSZZ054"
'****************************  start of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : サブフォーム制御
'       PROGRAM_ID      : MSZZ054
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2010/07/22
'       CERATER         : K.ISHIZAKA
'       Ver             : 0.0
'
'       UPDATE          :
'       UPDATER         :
'       Ver             :
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassNameA Lib "user32" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const GWL_STYLE = (-16)

' Scroll Bar Styles
'Private Const SBS_HORZ = &H0&
Private Const SBS_VERT = &H1&
'Private Const SBS_SIZEBOX = &H8&

' Windows Message Constant
Private Const WM_VSCROLL = &H115
'Private Const WM_HSCROLL = &H114

' Scroll Bar Commands
'Private Const SB_LINEUP = 0
'Private Const SB_LINELEFT = 0
'Private Const SB_LINEDOWN = 1
'Private Const SB_LINERIGHT = 1
'Private Const SB_PAGEUP = 2
'Private Const SB_PAGELEFT = 2
'Private Const SB_PAGEDOWN = 3
'Private Const SB_PAGERIGHT = 3
Private Const SB_THUMBPOSITION = 4
'Private Const SB_THUMBTRACK = 5
'Private Const SB_TOP = 6
'Private Const SB_LEFT = 6
'Private Const SB_BOTTOM = 7
'Private Const SB_RIGHT = 7
'Private Const SB_ENDSCROLL = 8

' GetWindow() Constants
Private Const GW_HWNDNEXT   As Long = 2
Private Const GW_CHILD      As Long = 5

'==============================================================================*
'
'       MODULE_NAME     : サブフォームの先頭行を指定
'       MODULE_ID       : SetTopRow
'       CREATE_DATE     : 2010/07/22
'       PARAM           : frm                   サブフォーム
'                       : lngIndex              行(I)
'       RETURN          : 正常(先頭行)／エラー(-1)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function SetTopRow(frm As Form, ByVal lngIndex As Long) As Long
    Dim hwnd                As Long
    Dim lngRet              As Long
    Dim lngThumb            As Long
    On Error GoTo ErrorHandler

    hwnd = GetWindowVScroll(frm)
    If hwnd = -1 Then
        SetTopRow = -1
        Exit Function
    End If

    lngThumb = MakeDWord(SB_THUMBPOSITION, CInt(lngIndex - 1))
    lngRet = SendMessage(frm.hwnd, WM_VSCROLL, ByVal lngThumb, ByVal hwnd)
    
    SetTopRow = lngIndex
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "SetTopRow" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : サブフォームの縦スクロールバー取得
'       MODULE_ID       : GetWindowVScroll
'       CREATE_DATE     : 2010/07/22
'       PARAM           : frm                   サブフォーム
'       RETURN          : 正常(ウインドウハンドル)／エラー(-1)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetWindowVScroll(frm As Form) As Long
    Dim hwnd                As Long
    On Error GoTo ErrorHandler

    hwnd = GetWindow(frm.hwnd, GW_CHILD)
    While hwnd <> 0
    
        If GetClassName(hwnd) = "scrollBar" Then
            If GetWindowLong(hwnd, GWL_STYLE) And SBS_VERT Then
                GetWindowVScroll = hwnd
                Exit Function
            End If
        End If
        
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Wend
    
    GetWindowVScroll = -1
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "GetWindowVScroll" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : クラス名の取得
'       MODULE_ID       : GetClassName
'       CREATE_DATE     : 2010/07/22
'       PARAM           : hWnd                  ウインドウハンドル(I)
'       RETURN          : クラス名
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetClassName(ByVal hwnd As Long) As String
    Dim strBuffer           As String
    Dim lngLen              As Long
    Const MAX_LEN           As Long = 255
    On Error GoTo ErrorHandler

    strBuffer = Space$(MAX_LEN)
    lngLen = GetClassNameA(hwnd, strBuffer, MAX_LEN)
    If lngLen > 0 Then
        GetClassName = Left$(strBuffer, lngLen)
    Else
        GetClassName = ""
    End If
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "GetClassName" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ダブルワードの作成
'       MODULE_ID       : MakeDWord
'       CREATE_DATE     : 2010/07/22
'       PARAM           : loWord                下位ワード(I)
'                       : hiWord                上位ワード(I)
'       RETURN          : ダブルワード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeDWord(ByVal loWord As Integer, ByVal hiWord As Integer) As Long
    MakeDWord = (hiWord * &H10000) Or (loWord And &HFFFF&)
End Function

'****************************  ended of program ********************************
