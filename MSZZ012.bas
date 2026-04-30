Attribute VB_Name = "MSZZ012"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME :
'
'        PROGRAM_NAME    : 日付チェック
'        PROGRAM_ID      : MSZZ012
'        PROGRAM_KBN     :
'
'        CREATE          : 2004/07/31
'        CERATER         : N.MIURA
'        Ver             : 0.0
'
'        UPDATE          :
'        UPDATER         :
'        Ver             :
'
'==============================================================================*
Option Explicit
'==============================================================================*
'
'        MODULE_NAME      :テスト
'        MODULE_ID        :TEST_MSZZ012_M00
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function TEST_MSZZ012_M00()
    
    If MSZZ012_M00("A00401") = True Then
       MsgBox ("正常")
    Else
       MsgBox ("異常")
    End If
    
    If MSZZ012_M00("200401") = True Then
       MsgBox ("正常")
    Else
       MsgBox ("異常")
    End If
    
    If MSZZ012_M00("2004011") = True Then
       MsgBox ("正常")
    Else
       MsgBox ("異常")
    End If
    
    If MSZZ012_M00("20040101") = True Then
       MsgBox ("正常")
    Else
       MsgBox ("異常")
    End If
    
    If MSZZ012_M00("200401011") = True Then
       MsgBox ("正常")
    Else
       MsgBox ("異常")
    End If
    
    If MSZZ012_M00("20040132") = True Then
       MsgBox ("正常")
    Else
       MsgBox ("異常")
    End If

End Function
'==============================================================================*
'
'        MODULE_NAME      :MAIN
'        MODULE_ID        :MSZZ012_M00
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZ012_M00(strCHYMD As String) As Boolean

    Dim WK_CHYMD             As String
    
    MSZZ012_M00 = False
    
    WK_CHYMD = ""
    
    '数値チェック
    If Not IsNumeric(strCHYMD) Then
       GoTo MSZZ012_M00_ERR
    End If
    
    '桁数チェック
    Select Case Len(strCHYMD)
    Case 6
         WK_CHYMD = strCHYMD & "01"
    Case 8
         WK_CHYMD = strCHYMD
    Case Else
       GoTo MSZZ012_M00_ERR
    End Select
    
    '日付チェック
    If (Not IsDate(Left(WK_CHYMD, 4) & "/" & Mid(WK_CHYMD, 5, 2) & "/" & Right(WK_CHYMD, 2))) Then
       GoTo MSZZ012_M00_ERR
    End If
    
    MSZZ012_M00 = True

MSZZ012_M00_ERR:
    
End Function
'****************************  ended or program ********************************

