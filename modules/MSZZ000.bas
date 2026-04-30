Attribute VB_Name = "MSZZ000"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 未収管理サブシステム
'
'        PROGRAM_NAME    :
'        PROGRAM_ID      :
'        PROGRAM_KBN     :
'
'        CREATE          : 2001/05/11
'        CERATER         : N.MIURA
'
'        UPDATE          :
'        UPDATER         :
'==============================================================================*
Option Compare Database
Option Explicit

'コンピュータ名取得
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'ユーザ名取得
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'==============================================================================*
'
'コンピュータ名取得
'
'
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function LsGetComputerName() As String
    
    Dim lpBuffer As String, nSize As Long, RTN As Long
    
    nSize = 255
    
    lpBuffer = Space$(nSize)
    
    RTN = GetComputerName(lpBuffer, nSize)
    
    LsGetComputerName = LsNullTrim(lpBuffer)

End Function
'==============================================================================*
'
'ユーザ名取得
'
'
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function LsGetUserName() As String
    
    Dim lpBuffer As String, nSize As Long, RTN As Long
    
    nSize = 255
    
    lpBuffer = Space$(nSize)
    
    RTN = GetUserName(lpBuffer, nSize)
    
    LsGetUserName = LsNullTrim(lpBuffer)

End Function
'==============================================================================*
'
'Null文字以降を削除
'
'
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function LsNullTrim(strExp As String) As String
    Dim i As Integer
    
    i = InStr(strExp, Chr$(0))
    
    If i > 1 Then
        LsNullTrim = Left$(strExp, i - 1)
    Else
        LsNullTrim = strExp
    End If

End Function
'****************************  ended or program ********************************
