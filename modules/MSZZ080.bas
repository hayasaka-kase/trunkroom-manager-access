Attribute VB_Name = "MSZZ080"
'****************************  start of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : コマンド引数起動
'       PROGRAM_ID      : MSZZ080
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2019/11/20
'       CERATER         : Y.WADA
'       Ver             : 0.0
'
'       UPDATE          : 2020/02/22
'       UPDATER         : Y.WADA
'       Ver             : 0.1
'                         ランタイム環境で動作しないのを修正
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "MSZZ080"

Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long) 'INSERT 2020/02/22 Y.WADA
'Private acApp As Object    'DELETE 2020/02/22 Y.WADA

'DELETE 2020/02/22 Y.WADA Start
''==============================================================================*
''
''        MODULE_NAME      :他のシステムのフォームを開く
''        MODULE_ID        :MSZZ080_SysFormOpen
''        PARM             :strId        フォーム名(I)
''                         :strArgs      フォームに渡す引数(I)
''                         :strSys       システム名（省略値：KAGTOS）(I)
''        CREATE_DATE      :2019/11/20 Y.WADA
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub MSZZ080_SysFormOpen(strID As String, strArgs As String, Optional strSys As String = "KAGTOS")
'
'    Dim strFilePath As String
'
'    On Error GoTo ErrorHandler
'
'    strFilePath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & PROG_ID & "' AND INTIF_RECFB = '" & strSys & "'"), "")
'    If strFilePath = "" Then
'        Call MSZZ024_M10("INTI_FILE", "呼び出し先のシステムが設定されていません(INTI_FILE," & PROG_ID & "," & strSys & ")")
'        Exit Sub
'    End If
'    On Error Resume Next
'    '既にmdbが開かれているかもしれないので一度フォームをオープンしてみる
'    acApp.doCmd.OpenForm strID, acNormal, , , , , strArgs
''    If Err = 2046 Then
'    If Err <> 0 Then
'        Err.Clear
'        'mdbを開いてフォームオープン
'        Set acApp = CreateObject("Access.Application")
'        acApp.OpenCurrentDatabase strFilePath
'        acApp.Visible = True
'        acApp.UserControl = True
'        acApp.doCmd.OpenForm strID, acNormal, , , , , strArgs
'    End If
'
'Exit Sub
'
'ErrorHandler:                   '↓自分の関数名
'    Call Err.Raise(Err.Number, "MSZZ080_SysFormOpen" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'End Sub
'DELETE 2020/02/22 Y.WADA End

'==============================================================================*
'
'        MODULE_NAME      :他のシステムのフォームを開く
'        MODULE_ID        :MSZZ080_SysFormOpen
'        PARM             :strId        フォーム名(I)
'                         :strArgs      フォームに渡す引数(I)
'                         :strSys       システム名（省略値：KAGTOS）(I)         INTIF_RECFBの値
'                         :strPrg       プログラムＩＤ（省略値：MSZZ080）(I)    INTIF_PROGBの値
'        CREATE_DATE      :2022/02/22 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZZ080_SysFormOpen(strID As String, strArgs As String, Optional strSys As String = "KAGTOS", Optional strPrg As String = PROG_ID)

    Dim strFilePath     As String
    Dim strExePath      As String
    Dim objWsh          As Object
    Dim acApp           As Object
    Dim i               As Integer

    On Error GoTo ErrorHandler
    
    strFilePath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & strPrg & "' AND INTIF_RECFB = '" & strSys & "'"), "")
    If strFilePath = "" Then
        Call MSZZ024_M10("INTI_FILE", "呼び出し先のシステムが設定されていません(INTI_FILE," & PROG_ID & "," & strSys & ")")
        Exit Sub
    End If
    
    strExePath = SysCmd(acSysCmdAccessDir) & "MSACCESS.EXE"
    
    On Error Resume Next

    '起動中のAccessを使用
    Set acApp = GetObject(strFilePath)
    
    If Err = 0 Then
        '起動中があれば使用する
        'フォームが開いていたら閉じる
        acApp.doCmd.Close acForm, strID, acSaveNo
    
    Else
        Err.Clear
        '起動中がなければ起動する
        Set objWsh = CreateObject("WScript.Shell")
        objWsh.Run """" & strExePath & """" & " /Runtime " & """" & strFilePath & """", 1, False
        
        For i = 1 To 20
            Set acApp = GetObject(strFilePath)
            If Err = 0 Then
                Exit For
            End If
            Err.Clear
            Sleep 500 '起動を待つ
        Next
    
    End If
    
    Err.Clear
    On Error GoTo ErrorHandler

'    acApp.Visible = True
'    acApp.UserControl = True
    'フォームを開く
    acApp.doCmd.OpenForm strID, acNormal, , , , , strArgs

Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MSZZ080_SysFormOpen" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
