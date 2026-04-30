Attribute VB_Name = "Const"
Option Compare Database
Option Explicit

'****** Iniﾌｧｲﾙのｷｰ値取得用DLLの参照
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'****** Windowsパス取得 API宣言
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Type RECT
    
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
    
End Type

'/****** lha関連関数の宣言及び変数
Declare Function LHA_UnlhaGetVersion Lib "unlha32.dll" Alias "UnlhaGetVersion" () As Long
Declare Function LHA_UnlhaGetRunning Lib "unlha32.dll" Alias "UnlhaGetRunning" () As Long
Declare Function LHA_Unlha Lib "unlha32.dll" Alias "Unlha" (ByVal hwnd As Long, ByVal szCmdLine$, ByVal szOutput$, ByVal dwSize As Long) As Long
Declare Function LHA_UnlhaGetBackGroundMode Lib "unlha32.dll" Alias "UnlhaGetBackGroundMode" () As Long
Declare Function LHA_UnlhaSetBackGroundMode Lib "unlha32.dll" Alias "UnlhaSetBackGroundMode" (ByVal BackGroundMode As Long) As Long
Declare Function LHA_UnlhaCheckArchive Lib "unlha32.dll" Alias "UnlhaCheckArchive" (ByVal szFileName$, ByVal iMode As Long) As Long
Declare Function LHA_UnlhaGetFileCount Lib "unlha32.dll" Alias "UnlhaGetFileCount" (ByVal szArcFile As String) As Long

Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'■警告: 該当ファイルについての処理をスキップするだけで実行を中止する事
'        はない｡
Public Const ERROR_DISK_SPACE = &H8005
'解凍する為のディスクの空きが足りません。
Public Const ERROR_READ_ONLY = &H8006
'解凍先のファイルはリードオンリーです｡
Public Const ERROR_USER_SKIP = &H8007
'ユーザーによって解凍をスキップされました。
Public Const ERROR_UNKNOWN_TYPE = &H8008
Public Const ERROR_METHOD = &H8009
Public Const ERROR_PASSWORD_FILE = &H800A
Public Const ERROR_VERSION = &H800B
Public Const ERROR_FILE_CRC = &H800C
'格納ファイルのチェックサムが合っていません｡
Public Const ERROR_FILE_OPEN = &H800D
'解凍時にファイルを開けませんでした｡
Public Const ERROR_MORE_FRESH = &H800E
'より新しいファイルが解凍先に存在しています｡
Public Const ERROR_NOT_EXIST = &H800F
'ファイルは解凍先に存在していません。
Public Const ERROR_ALREADY_EXIST = &H8010

Public Const ERROR_TOO_MANY_FILES = &H8011

'■エラー: 致命的なエラーでその時点で実行を中止する｡
Public Const ERROR_MAKEDIRECTORY = &H8012
'ディレクトリが作成できません｡
Public Const ERROR_CANNOT_WRITE = &H8013
'解凍中に書き込みエラーが生じました｡
Public Const ERROR_HUFFMAN_CODE = &H8014
'LZH ファイルのハフマンコードが壊れています。
Public Const ERROR_COMMENT_HEADER = &H8015
'LZH ファイルのコメントヘッダが壊れています。
Public Const ERROR_HEADER_CRC = &H8016
'LZH ファイルのヘッダのチェックサムが合っていません。
Public Const ERROR_HEADER_BROKEN = &H8017
'LZH ファイルのヘッダが壊れています。
Public Const ERROR_ARC_FILE_OPEN = &H8018
'LZH ファイルを開く事が出来ません。
Public Const ERROR_NOT_ARC_FILE = &H8019
'LZH ファイル名の指定がされていません。
Public Const ERROR_CANNOT_READ = &H801A
'LZH ファイルの読み込み時に読み込みエラーが出ました。
Public Const ERROR_FILE_STYLE = &H801B
'指定されたファイルは LZH ファイルではありません。
Public Const ERROR_COMMAND_NAME = &H801C
'コマンド指定が間違っています。
Public Const ERROR_MORE_HEAP_MEMORY = &H801D
'作業用のためのヒープメモリが不足しています。
Public Const ERROR_ENOUGH_MEMORY = &H801E
'グローバルメモリが不足しています｡
Public Const ERROR_ALREADY_RUNNING = &H801F
'既に UNLHA32.DLL が動作中です。
Public Const ERROR_USER_CANCEL = &H8020
'ユーザーによって解凍を中断されました。
Public Const ERROR_HARC_ISNOT_OPENED = &H8021
'UnlhaOpenArchive() で書庫ファイルとハンドルを関連付ける前に Unlha-
'FindFirst() 等の API を使用した。
Public Const ERROR_NOT_SEARCH_MODE = &H8022
'UnlhaFindFirst() を使用する前に UnlhaFindNext() が呼ばれた。または，
'これらの API を呼び出す前に UnlhaGetFileName() 等の API が呼ばれた。
Public Const ERROR_NOT_SUPPORT = &H8023
'UNLHA32.DLL でサポートされていない API が使用されました。
Public Const ERROR_TIME_STAMP = &H8024
'日付及び時間の指定形式が間違っています｡
Public Const ERROR_TMP_OPEN = &H8025
'作業ファイルがオープンできません｡
Public Const ERROR_LONG_FILE_NAME = &H8026
'ディレクトリのパスが長すぎます｡
Public Const ERROR_ARC_READ_ONLY = &H8027
'書き込み専用属性の書庫ファイルに対する操作はできません｡
Public Const ERROR_SAME_NAME_FILE = &H8028
'すでに同じ名前のファイルが書庫に格納されています｡
Public Const ERROR_NOT_FIND_ARC_FILE = &H8029
'指定されたディレクトリには LZH ファイルがありませんでした。


'解約可能な過去の期間                                                               'ADD 20040423 K.ISHIZAKA
Public Const C_KAIYAKU_KIKAN_FROM       As Long = (-3)  '当月を含めて３ヶ月前まで   'ADD 20040423 K.ISHIZAKA
Public Const C_KAIYAKU_KIKAN_TO         As Long = (20)  '当月を含めて20ヶ月後まで   'ADD 20040423 K.ISHIZAKA


'/****** ｼｽﾃﾑ情報等の変数
Global Const gstIni = "Hsys.ini"       '/****** Iniﾌｧｲﾙ名

'/****** ﾚﾎﾟｰﾄ用のｸﾞﾛｰﾊﾞﾙ変数
Global gboPrintFlg   As Boolean             '/****** 印刷時の表示、非表示のﾌﾗｸﾞ
Global gstPrintTitle() As String            '/****** ﾚﾎﾟｰﾄの条件表示用Global変数
Global gvaPrintWhere() As Variant           '/****** 画面の範囲指定値

'/****** ﾒﾆｭｰﾊﾞｰの表示ﾓｰﾄﾞ格納
Global gstMenubar   As String               '/****** ﾃｽﾄﾓｰﾄﾞのｸﾞﾛｰﾊﾞﾙ変数

'/****** Acのﾃｰﾌﾞﾙとﾘﾝｸしている場合のﾘﾝｸ先のﾌｧｲﾙ情報を格納する変数
Global gstMDBPath   As String               '/****** ﾘﾝｸﾃｰﾌﾞﾙのﾊﾟｽ
Global gstMDBName   As String               '/****** ﾘﾝｸﾃｰﾌﾞﾙ名

Public strSize() As String   '2016/12/13 M.HONDA
Public strBox()  As String   '2016/12/13 M.HONDA

