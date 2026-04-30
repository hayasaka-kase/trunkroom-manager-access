Attribute VB_Name = "MSZZ075"
'****************************  start of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : WindowsAPIで発生したエラーの文言を取得する
'       PROGRAM_ID      : MSZZ075
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2017/01/26
'       CERATER         : K.ISHIZAKA
'       Ver             : 0.0
'
'       UPDATE          :
'       UPDATER         :
'       Ver             : 0.1
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   API宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Declare Function FormatMessageW Lib "kernel32.dll" ( _
    ByVal dwFlags As Long, _
    ByRef lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByRef lpBuffer As Long, _
    ByVal nSize As Long, _
    ByRef Arguments As Any) As Long
Private Declare Function lstrcpyW Lib "kernel32.dll" ( _
    ByVal lpString1 As Long, _
    ByVal lpString2 As Long) As Long
Private Declare Function LocalFree Lib "kernel32.dll" ( _
    ByVal hMem As Long) As Long

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER    As Long = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS     As Long = &H200
Private Const FORMAT_MESSAGE_FROM_STRING        As Long = &H400
Private Const FORMAT_MESSAGE_FROM_HMODULE       As Long = &H800
Private Const FORMAT_MESSAGE_FROM_SYSTEM        As Long = &H1000
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY     As Long = 8192
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK     As Long = 255
Private Const LANG_NEUTRAL                      As Long = &H0
Private Const SUBLANG_DEFAULT                   As Long = &H1

'==============================================================================*
'
'       MODULE_NAME     : MAKELANGIDマクロの代替え
'       MODULE_ID       : MAKELANGID
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : p                     プライマリ(I)
'                       : s                     セカンダリ(I)
'       RETURN          : ID(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MAKELANGID(ByVal p As Long, ByVal S As Long) As Long
    MAKELANGID = (CLng(CInt(S)) * 1024) Or CLng(CInt(p))
End Function

'==============================================================================*
'
'       MODULE_NAME     : WindowsAPIで発生したエラーの文言を取得する
'       MODULE_ID       : GetAPIErrorText
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : ErrorCode             エラーコード(I)
'       RETURN          : エラー文言(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetAPIErrorText(ByVal ErrorCode As Long) As String
    Dim lpBuffer            As Long
    Dim messageLength       As Long
    Dim strResult           As String

    messageLength = FormatMessageW( _
        FORMAT_MESSAGE_ALLOCATE_BUFFER Or FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
        0, ErrorCode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
        lpBuffer, 0, 0)

    If messageLength = 0 Then
        strResult = ""
    Else
        strResult = Space$(messageLength)
        Call lstrcpyW(ByVal StrPtr(strResult), ByVal lpBuffer)
        Call LocalFree(lpBuffer)
    End If
    GetAPIErrorText = strResult
End Function

'****************************  ended of program ********************************
