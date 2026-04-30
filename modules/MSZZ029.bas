Attribute VB_Name = "MSZZ029"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : 画面が開いているかチェック
'        PROGRAM_ID      : MSZZ029
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2007/03/01
'        CERATER         : S.Shibazaki
'        Ver             : 0.0
'
'        UPDATE          : 2008/05/21
'        UPDATER         : S.Shibazaki
'        Ver             : 0.1
'                        : レポートが開いているか調べる関数を追加
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'
'       MODULE_NAME     : 画面が開いているかチェック
'       MODULE_ID       : fncFormIsLoaded
'       CREATE_DATE     : 2007/03/01            S.Shibazaki
'       PARAM           : strFormId             呼び出し先フォームID(I)
'       RETURN          : True=画面が開いている False=開いていない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncFormIsLoaded(ByVal strFormID As String) As Boolean
On Error GoTo FormNothing

    fncFormIsLoaded = CurrentProject.AllForms(strFormID).IsLoaded

    Exit Function
    
FormNothing:
    fncFormIsLoaded = False
    
End Function

'==============================================================================*
'
'       MODULE_NAME     : 画面が開いていたらメッセージを表示
'       MODULE_ID       : fncFormIsLoaded
'       CREATE_DATE     : 2007/03/01            S.Shibazaki
'       PARAM           : strFormId             呼び出し先フォームID(I)
'                       : strMsgCaption         メッセージボックスタイトル(I)
'       RETURN          : True=画面が開いている False=開いていない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncFormIsLoadedMsg(ByVal strFormID As String, _
                                   ByVal strMsgCaption As String) As Boolean
    
    fncFormIsLoadedMsg = fncFormIsLoaded(strFormID)
    
    If fncFormIsLoadedMsg Then
        MsgBox "[" & Forms(strFormID).Caption & "]は既に開かれています。" & vbCrLf & _
               "閉じてから再度実行してください。", _
               vbOKOnly + vbInformation, strMsgCaption
    End If
        
End Function

'==============================================================================*
'
'       MODULE_NAME     : レポートが開いているかチェック
'       MODULE_ID       : fncReportIsLoaded
'       CREATE_DATE     : 2008/05/21            S.Shibazaki
'       PARAM           : strReportId           レポートID(I)
'       RETURN          : True=開いている False=開いていない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncReportIsLoaded(ByVal StrReportID As String) As Boolean
On Error GoTo FormNothing

    fncReportIsLoaded = CurrentProject.AllReports(StrReportID).IsLoaded

    Exit Function
    
FormNothing:
    fncReportIsLoaded = False
    
End Function

'==============================================================================*
'
'       MODULE_NAME     : レポートが開いていたらメッセージを表示
'       MODULE_ID       : fncReportIsLoadedMsg
'       CREATE_DATE     : 2008/05/21            S.Shibazaki
'       PARAM           : strReportId           レポートID(I)
'                       : strMsgCaption         メッセージボックスタイトル(I)
'       RETURN          : True=開いている False=開いていない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncReportIsLoadedMsg(ByVal StrReportID As String, _
                                   ByVal strMsgCaption As String) As Boolean
    
    fncReportIsLoadedMsg = fncReportIsLoaded(StrReportID)
    
    If fncReportIsLoadedMsg Then
        MsgBox "レポート[" & Reports(StrReportID).Caption & "]は既に開かれています。" & vbCrLf & _
               "閉じてから再度実行してください。", _
               vbOKOnly + vbInformation, strMsgCaption
    End If
        
End Function

'****************************  ended or program ********************************

