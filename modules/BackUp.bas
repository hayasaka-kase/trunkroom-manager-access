Attribute VB_Name = "BackUp"
Option Compare Database
Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'API宣言
Private Declare Function GetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'プロパティ変数（手抜き）
Public DefaultExt As String
Public DialogTitle As String
Public Filter As String
Public Flags As Long
Public OwnerhWnd As Long
Public fileName As String

'コモンダイアログの表示
Property Let Action(new_Action As Integer)
    Dim res As String
    
    '実際は内部プロシージャにて実行
    res = doCmd(new_Action)
    
    'FileNameプロパティのセット
    fileName = res

End Property

'コモンダイアログを表示する内部プロシージャ
Private Function doCmd(Action As Integer)
    Dim ret As Long
    Dim OFN As OPENFILENAME

    '構造体への格納
    With OFN
        .lStructSize = Len(OFN)
        .hWndOwner = OwnerhWnd
'        .hInstance = App.hInstance
        
        'COMDLG32.OCX互換を維持するため「|」の区切りをNullにする
        .lpstrFilter = Replace(Filter, "|", Chr(0))
        
        .lpstrFile = String(255, Chr(0))
        .nMaxFile = 255
        .lpstrFileTitle = String(255, Chr(0))
        .nMaxFileTitle = 255
        .lpstrInitialDir = vbNullString
        .lpstrTitle = DialogTitle
        .Flags = Flags
    End With
    
    'タイプによってAPIを呼び出し
    Select Case Action
    Case 1
        'ファイルを開く
    
        ret = GetOpenFileName(OFN)
        
    Case 2
        'ファイルを保存
    
        ret = GetSaveFileName(OFN)
        
    Case Else
        '未対応m(__)m
    
        doCmd = 0
    
        Exit Function
        
    End Select
    
    'ファイル名を抽出
    If ret > 0 Then
        doCmd = Left(OFN.lpstrFile, InStr(OFN.lpstrFile, Chr(0)) - 1)
    Else
        doCmd = ""
    End If

End Function

'置換ルーチン

'注意：MaeとAtoの引数の文字数が同じ場合のみ動作保証します
Private Function Replace(St As String, Mae As String, Ato As String)
Dim r As Long, e As String, pos As Long

e = St
pos = 1

Do

    r = InStr(pos, e, Mae)
    
    pos = r + 1
    
    If r = 0 Then Replace = e: Exit Function
   
    e = Left$(e, r - 1) + Ato + Mid$(e, r + Len(Mae))
    
Loop

End Function

