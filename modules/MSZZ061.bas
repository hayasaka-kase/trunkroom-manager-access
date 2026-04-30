Attribute VB_Name = "MSZZ061"
'****************************  strat of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬総合システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : あいまい検索
'       PROGRAM_ID      : MSZZ061
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2011/02/12
'       CERATER         : K.ISHZIAKA
'       Ver             : 0.0
'
'       UPDATE          :
'       UPDATER         :
'       Ver             :
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID       As String = "MSZZ061"

'==============================================================================*
'
'       MODULE_NAME     : あいまい検索文字列に変換する
'       MODULE_ID       : changeFuzzyString
'       CREATE_DATE     : 2011/02/12            K.ISHIZAKA
'       PARAM           : strSource             文字列(I)
'       RETURN          : あいまい検索用文字列
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function changeFuzzyString(ByVal strSource As String) As String
    Dim varChar             As Variant
    On Error GoTo ErrorHandler
    
    For Each varChar In Split(getFuzzyChar(), ";")
        strSource = Replace(strSource, varChar, "")
    Next
    changeFuzzyString = strSource
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "changeFuzzyString" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : あいまい検索文字列に変換するカラム
'       MODULE_ID       : changeFuzzySQL
'       CREATE_DATE     : 2011/02/12            K.ISHIZAKA
'       PARAM           : strColumn             カラム(I)
'       RETURN          : ＳＱＬ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function changeFuzzySQL(ByVal strColumn As String) As String
    Dim varChar             As Variant
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = strColumn
    For Each varChar In Split(getFuzzyChar(), ";")
        strSQL = "Replace(" & strSQL & ",'" & varChar & "','')"
    Next
    changeFuzzySQL = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "changeFuzzySQL" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : あいまい検索対象文字の取得
'       MODULE_ID       : getFuzzyChar
'       CREATE_DATE     : 2011/02/12            K.ISHIZAKA
'       RETURN          : あいまい検索対象文字
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function getFuzzyChar() As String
    On Error GoTo ErrorHandler
    getFuzzyChar = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='" & PROG_ID & "'"), "")
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "getFuzzyChar" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : あいまい検索対象文字の保存
'       MODULE_ID       : setFuzzyChar
'       CREATE_DATE     : 2011/02/12            K.ISHIZAKA
'       PARAM           : strFuzzyChar          あいまい検索対象文字(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub setFuzzyChar(ByVal strFuzzyChar As String)
    Dim objRst              As Recordset
    On Error GoTo ErrorHandler
    
    Set objRst = CurrentDb.OpenRecordset("SELECT * FROM INTI_FILE WHERE INTIF_PROGB='" & PROG_ID & "'", dbOpenDynaset)
    On Error GoTo ErrorHandler1
    With objRst
        If .EOF Then
            .AddNew
            On Error GoTo ErrorHandler2
            .Fields("INTIF_PROGB") = PROG_ID
            .Fields("INTIF_RECFB") = "あいまい検索対象文字"
            .Fields("INTIF_BIKON") = "セミコロン（;）で複数指定"
        Else
            .Edit
            On Error GoTo ErrorHandler2
        End If
        .Fields("INTIF_RECDB") = Replace(strFuzzyChar, "；", ";")
        .UPDATE
        On Error GoTo ErrorHandler1
        .Close
    End With
    On Error GoTo ErrorHandler
Exit Sub
    
ErrorHandler2:
    objRst.CancelUpdate
ErrorHandler1:
    objRst.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "setFuzzyChar" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'****************************  ended of program ********************************

