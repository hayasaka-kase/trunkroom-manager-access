Attribute VB_Name = "API_Function"
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　API関数定義
'   プログラムＩＤ　：　API_Function
'   作　成　日　　　：  2003/05/21
'   作　成　者　　　：  azegami@eagle-soft.co.jp
'**********************************************
'修正履歴
'   修　正　日　　　：
'   修　正　者　　　：
'   修　正　内　容　：
'**********************************************

Option Compare Database
Option Explicit

'ファイルを拡張子に関連付けされているアプリケーションで開く
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1      '通常ウインドウ
Public Const SW_SHOWMINIMIZED = 2   '最小化
Public Const SW_SHOWMAXIMIZED = 3   '最大化

Public Const ERROR_FILE_NOT_FOUND = 2&      'ファイルが見つからない
Public Const ERROR_PATH_NOT_FOUND = 3&      'パス名が見つからない
Public Const ERROR_BAD_FORMAT = 11&         'Win32用EXEではないか EXE内にエラーがある

Public Sub Y_Exec(hwnd As Long, FilePath As String, parameter As String, WorkPath As String, WindowSize As Long)
'*******************************************************************
'機能  : ShellExecute関数を呼び出し､ファイルを関連付けされている
'　　　　アプリケーションで開く
'引数　: hWnd　　　 = 呼び出し元ウインドウハンドル
'　　　　FilePath 　= 開きたいファイルのフルパス名
'　　　　WorkPath 　= 作業フォルダのフルパス名
'　　　　WindowSize = アプリケーションウインドウの大きさ
'　　　　　　　　　　　1:通常のサイズで開く
'　　　　　　　　　　　2:最小化して開く
'　　　　　　　　　　　3:最大化して開く
'*******************************************************************

On Error GoTo err_rtn

    Dim longret As Long
    Dim msg As String
    
    longret = ShellExecute(hwnd, "Open", FilePath, parameter, WorkPath, WindowSize)
    If longret < 31 Then
        Select Case longret
            Case 0
                msg = "メモリ不足です。"
            Case ERROR_FILE_NOT_FOUND
                msg = "ファイルが見つかりません。"
            Case ERROR_PATH_NOT_FOUND
                msg = "ファイルのパスが見つかりません。"
            Case Else
                msg = longret & "その他のエラー"
        End Select
        Call MsgBox(msg, 16)
    End If
    Exit Sub
    
err_rtn:
    Call MsgBox("その他のエラー", 16)

End Sub


