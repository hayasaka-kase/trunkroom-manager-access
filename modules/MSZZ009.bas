Attribute VB_Name = "MSZZ009"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME :
'
'        PROGRAM_NAME    : ハイパーリンク
'        PROGRAM_ID      : MSZZ009
'        PROGRAM_KBN     :
'
'        CREATE          : 2003/12/18
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
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "MSZZ009"
'
Private hlk             As Hyperlink
'
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :MSZZ009_M00
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZ009_M00(ctlSelected As Control, _
                     strSubAddress As TextBox, _
                     Optional strAddress As TextBox)
    On Error GoTo err_MSZZ009_M00
   
    Select Case ctlSelected.ControlType
    Case acLabel, acImage, acCommandButton
          
          Set hlk = ctlSelected.Hyperlink
          
          With hlk
          If Not IsMissing(strAddress) Then
            .ADDRESS = strAddress
          Else
            .ADDRESS = ""
          End If
          
          .SubAddress = strSubAddress
                             .Follow
          .ADDRESS = ""
          .SubAddress = ""
          End With
     
    Case Else
          MsgBox "このコントロール ( '" & ctlSelected.NAME & "' ) は､ハイパーリンクをサポートしていません｡ "
    End Select

Exit_MSZZ009_M00:
    Exit Function

err_MSZZ009_M00:
  'MsgBox "ｴﾗｰ番号" & err.Number & vbCrLf & err.Description
  Err.Clear
  GoTo Exit_MSZZ009_M00
    
End Function
'****************************  ended or program ********************************

