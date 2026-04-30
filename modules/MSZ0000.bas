Attribute VB_Name = "MSZ0000"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 郵便番号変換制御
'
'        PROGRAM_NAME    : 郵便番号変換制御
'        PROGRAM_ID      : MSZ0000
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2004/04/09
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
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private strParentFormName   As String   'フォーム
Private strParentNYUBBName  As String   'テキスト郵便番号
Private strParentADDRNName  As String   'テキスト住所
Private strParentNYUBBVal   As String   'テキスト郵便番号
Private strParentADDRNVal   As String   'テキスト住所

'==============================================================================*
'
'        MODULE_NAME      :郵便番号ダイアログ表示
'        MODULE_ID        :MSZ0000_YUBIB
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZ0000_YUBIB(strFormName As String, strNYUBBName As String, strADDRNName As String, strNYUBBVal As String, strADDRNVal As String)
    strParentFormName = strFormName
    strParentNYUBBName = strNYUBBName
    strParentADDRNName = strADDRNName
    If InStr(strNYUBBVal, "-") > 0 Then
        strParentNYUBBVal = Left(strNYUBBVal, 3) & Mid(strNYUBBVal, 5)
    Else
        strParentNYUBBVal = strNYUBBVal
    End If
    strParentADDRNVal = strADDRNVal
    doCmd.OpenForm "MFZ0000", acNormal, , , , acDialog
End Sub

'==============================================================================*
'
'        MODULE_NAME      :住所取得
'        MODULE_ID        :MSZ0000_GET_ADDRN
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZ0000_GET_ADDRN() As String
    Dim strRet              As String

    strRet = strParentADDRNVal
    MSZ0000_GET_ADDRN = strRet
End Function

'==============================================================================*
'
'        MODULE_NAME      :郵便番号取得
'        MODULE_ID        :MSZ0000_GET_NYUBB
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZ0000_GET_NYUBB() As String
    Dim strRet              As String

    strRet = strParentNYUBBVal
    MSZ0000_GET_NYUBB = strRet
End Function

'==============================================================================*
'
'        MODULE_NAME      :郵便番号反映
'        MODULE_ID        :MSZ0000_SET
'        CREATE_DATE      :2004/04/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZ0000_SET(strNYUBB As String, strTODON As String)
    On Error GoTo Err_Handle
    If InStr(strNYUBB, "-") = 0 Then
        Forms(strParentFormName).Controls(strParentNYUBBName).VALUE = Left(strNYUBB, 3) & "-" & Mid(strNYUBB, 4)
    Else
        Forms(strParentFormName).Controls(strParentNYUBBName).VALUE = strNYUBB
    End If
    doCmd.Close
    With Forms(strParentFormName).Controls(strParentADDRNName)
        .SetFocus
        .VALUE = strTODON
        .SelStart = LenB(strTODON)
    End With
    Exit Sub

Err_Handle:
    MsgBox Err.Description
    On Error Resume Next
End Sub

'****************************  ended or program ********************************
