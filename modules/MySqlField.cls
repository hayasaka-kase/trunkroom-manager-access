VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MySqlField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************  strat of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬総合システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : WebServer MySQL フィールドクラス
'       PROGRAM_ID      : MySqlField
'       PROGRAM_KBN     : Class
'
'       CREATE          : 2012/06/16
'       CERATER         : K.ISHZIAKA
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
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private pName As String
Private pValue As String

'==============================================================================*
'
'       MODULE_NAME     : フィールド名（プロパティ）
'       MODULE_ID       : Name
'       CREATE_DATE     : 2012/06/16            K.ISHIZAKA
'       RETURN          : フィールド名(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Property Get NAME() As String
    NAME = pName
End Property

'==============================================================================*
'
'       MODULE_NAME     : フィールド値（プロパティ）
'       MODULE_ID       : Value
'       CREATE_DATE     : 2012/06/16            K.ISHIZAKA
'       RETURN          : フィールド値(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Property Get VALUE() As String
    VALUE = pValue
End Property

'==============================================================================*
'
'       MODULE_NAME     : 設定
'       MODULE_ID       : SetValue
'       CREATE_DATE     : 2012/06/16            K.ISHIZAKA
'       PARAM           : strValue              フィールド値(I)
'                       : strName               フィールド名(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetValue(ByVal strValue As String, ByVal strName As String)
    pValue = strValue
    pName = strName
End Sub

'****************************  ended of program ********************************
