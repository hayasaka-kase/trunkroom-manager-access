Attribute VB_Name = "MSZZ003"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    :
'        PROGRAM_ID      : MSZZ003
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2003/02/10
'        CERATER         : N.MIURA
'
'        UPDATE          : 2017/04/01
'        UPDATER         : K.SATO
'                        :
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "MSZZ003"
'
Private USER_ID         As String
Private MODL_ID         As String
'
Private WK_DATED        As String
Private WK_TIMED        As String
Private WK_ERR          As String
'
Private intFNum         As Integer

Private mstrAppPath     As String
Private mstrErrLogName  As String
Private mstrLogName     As String
'
Private intPos          As Integer
Private intPosSave      As Integer
Private strDbName       As String
'
'==============================================================================*
'   テスト用
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZ003_TEST()
    
    Call MSZZ003_M00("TEST00", "0", "")
    Call MSZZ003_M00("TEST00", "1", "XXXXXX")

End Function
'==============================================================================*
'
'        MODULE_NAME      :メイン
'        MODULE_ID        :MSZZ003
'        CREATE_DATE      :2003/02/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZ003_M00(strID As String, strKBI, strPRINN As String)
    On Error GoTo MSZZ003_M00_Err
    
    MSZZ003_M00 = False
    
    USER_ID = LsGetUserName()
    
    WK_DATED = Format(Now(), "yyyymmdd")
    WK_TIMED = Format(Now(), "HHMMSS")
    WK_ERR = 0
    
    'カレントパス
    strDbName = CurrentDb.NAME
    
    intPos = InStr(1, strDbName, "\")
    
    Do Until intPos = 0
        
        intPosSave = intPos
        intPos = InStr(intPosSave + 1, strDbName, "\")
    
    Loop
    
     
    mstrAppPath = Left$(CurrentDb.NAME, intPosSave)
    
    If Right$(mstrAppPath, 1) <> "\" Then
    
        mstrAppPath = mstrAppPath & "\"
    
    End If
 
    '実行ログファイル名
    mstrLogName = mstrAppPath & strID & ".log"


    Select Case strKBI
           Case "0" '開始ログ出力
                Call MSZZ003_M80(strID & " START " & strPRINN)
           Case "1" '終了処理
                Call MSZZ003_M80(strID & " ENDED " & strPRINN)
           'INSERT 2017/04/01 K.ASTO Start
           Case "8" '出力処理
                Call MSZZ003_M80(strID & strPRINN)
           'INSERT 2017/04/01 K.ASTO End
           Case "9" '異常終了
                Call MSZZ003_M80(strID & " ENDED " & strPRINN)
    End Select
    
    If WK_ERR = 0 Then
       
       MSZZ003_M00 = True
    
    End If

MSZZ003_M00_Exit:
    Exit Function

MSZZ003_M00_Err:
    MsgBox Error$
    Resume MSZZ003_M00_Exit

End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :MSZZ003_M80
'        CREATE_DATE      :
'        PARA             :strLog：出力する文字列
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Sub MSZZ003_M80(strLog As String)

    intFNum = FreeFile
    
    Open (mstrLogName) For Append As #intFNum
    
    Print #intFNum, Format(Now, "yyyy/mm/dd hh:nn:ss") & " " & strLog
    
    Close #intFNum

End Sub
'****************************  ended or program ********************************


