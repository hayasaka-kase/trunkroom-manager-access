Attribute VB_Name = "MSZZ013"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : ActiveX 存在チェック
'        PROGRAM_ID      : MSZZ013
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2004/12/03
'        CERATER         : K.ISHIZAKA
'        Ver             : 0.0
'
'        UPDATE          :
'        UPDATER         :
'        Ver             :
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Declare Function RegOpenKey Lib "ADVAPI32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, hKey As Long) As Long
Private Declare Function RegCloseKey Lib "ADVAPI32" (ByVal hKey As Long) As Long

Private Const HKEY_CLASSES_ROOT     As Long = &H80000000
Private Const HKEY_CURRENT_USER     As Long = &H80000001
Private Const HKEY_CURRENT_CONFIG   As Long = &H80000005
Private Const HKEY_DYN_DATA         As Long = &H80000006
Private Const HKEY_LOCAL_MACHINE    As Long = &H80000002
Private Const HKEY_USERS            As Long = &H80000003

'==============================================================================*
'
'        MODULE_NAME      :指定したAxtiveXがインストールされているかチェックする
'        MODULE_ID        :MSZZ013_M00
'        CREATE_DATE      :2004/12/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ013_M00(ByVal ActiveXName As String) As Boolean
    Dim Handle  As Long
    Dim RetVal  As Boolean

    RetVal = (RegOpenKey(HKEY_CLASSES_ROOT, ActiveXName, 0, 1, Handle) = 0)
    If RetVal Then
        Call RegCloseKey(Handle)
    End If
    MSZZ013_M00 = RetVal
End Function

'****************************  ended or program ********************************

