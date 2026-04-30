Attribute VB_Name = "MSZZ060"
'****************************  strat of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬総合システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : 用途別背景色設定
'       PROGRAM_ID      : MSZZ060
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2011/02/12
'       CERATER         : K.ISHZIAKA
'       Ver             : 0.0
'
'       UPDATE          : 2011/03/10
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.1
'                       : コンボボックスにも対応する
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID       As String = "MSZZ060"

'==============================================================================*
'
'       MODULE_NAME     : 用途カラーを取得する
'       MODULE_ID       : getUsageColor
'       CREATE_DATE     : 2011/02/12            K.ISHIZAKA
'       PARAM           : strUsage              用途(I)
'       RETURN          : カラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function getUsageColor(ByVal strUsage As String) As Long
    On Error GoTo ErrorHandler
    getUsageColor = CLng(Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='" & PROG_ID & "' AND INTIF_RECFB = '" & Format(Val(strUsage), "00") & "'"), 0))
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "getUsageColor" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 用途カラーを設定する
'       MODULE_ID       : setUsageBackColor
'       CREATE_DATE     : 2011/02/12            K.ISHIZAKA
'       PARAM           : txtUsage              用途コントロール(I)
'                       : txtBoxs               関係するコントロール達(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub setUsageBackColor(ByVal txtUsage As TextBox, ParamArray txtBoxs())  'DELETE 2011/03/10 K.ISHIZAKA
Public Sub setUsageBackColor(ByVal txtUsage As Control, ParamArray txtBoxs())   'INSERT 2011/03/10 K.ISHIZAKA
    Dim ctrl            As Variant
    Dim coTbl           As Collection
    On Error GoTo ErrorHandler

    If TypeName(txtBoxs(LBound(txtBoxs))) <> "Controls" Then
        If Not IsArray(txtBoxs(LBound(txtBoxs))) Then
            Call setUsageBackColor(txtUsage, txtBoxs)
            Exit Sub
        End If
    End If
    Set coTbl = collectionCurrentDbTable("SELECT INTIF_RECFB,INTIF_RECDB FROM INTI_FILE WHERE INTIF_PROGB='" & PROG_ID & "'")
    For Each ctrl In txtBoxs(LBound(txtBoxs))
        Call setControlBackColor(txtUsage.NAME, ctrl, coTbl)
    Next
    If UCase(TypeName(txtUsage)) = "TEXTBOX" Then                               'INSERT 2011/03/10 K.ISHIZAKA
        Call setControlBackColor(txtUsage.NAME, txtUsage, coTbl)
    End If                                                                      'INSERT 2011/03/10 K.ISHIZAKA
Exit Sub
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "setUsageBackColor" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 用途カラーを設定する
'       MODULE_ID       : setControlBackColor
'       CREATE_DATE     : 2011/02/12            K.ISHIZAKA
'       PARAM           : strUsageName          用途コントロール名(I)
'                       : ctrl                  書式設定するテキストボックス(I)
'                       : coTbl                 条件と色の情報(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setControlBackColor(ByVal strUsageName As String, ctrl As Variant, coTbl As Collection)
    Dim i               As Long
    On Error GoTo ErrorHandler
    
    If UCase(TypeName(ctrl)) = "TEXTBOX" Then
        If ctrl.Visible Then
            With ctrl.FormatConditions
                .Delete
                For i = 1 To coTbl.Count
                    With .Add(acExpression, , "[" & strUsageName & "] = '" & Replace(coTbl(i)("INTIF_RECFB"), ",", "' OR [" & strUsageName & "] = '") & "'")
                        .BackColor = coTbl(i)("INTIF_RECDB")
                        .FontBold = ctrl.FontBold
                        .ForeColor = ctrl.ForeColor
                    End With
                Next
            End With
        End If
    End If
Exit Sub
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "setControlBackColor" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 読込データをコレクション化する
'       MODULE_ID       : collectionCurrentDbTable
'       CREATE_DATE     : 2011/02/12            K.ISHIZAKA
'       PARAM           : strSQL                ＳＱＬ文(I)
'       RETURN          : テーブル(Collection)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function collectionCurrentDbTable(ByVal strSQL As String) As Collection
    Dim objRst              As Recordset
    Dim coTbl               As New Collection
    On Error GoTo ErrorHandler
    
    Set objRst = CurrentDb.OpenRecordset(strSQL, dbDenyRead)
    On Error GoTo ErrorHandler1
    While Not objRst.EOF
        coTbl.Add collectionRecord(objRst)
        objRst.MoveNext
    Wend
    objRst.Close
    On Error GoTo ErrorHandler
    Set collectionCurrentDbTable = coTbl
Exit Function
    
ErrorHandler1:
    objRst.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "collectionCurrentDbTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 読込データをコレクション化する
'       MODULE_ID       : collectionRecord
'       CREATE_DATE     : 2011/02/12            K.ISHIZAKA
'       PARAM           : objRst                レコードセットオブジェクト(I)
'       RETURN          : レコード(Collection)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function collectionRecord(objRst As Recordset) As Collection
    Dim coRec               As New Collection
    Dim fld                 As Field
    On Error GoTo ErrorHandler

    For Each fld In objRst.Fields
        coRec.Add fld.VALUE, fld.NAME
    Next
    Set collectionRecord = coRec
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "collectionRecord" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  ended of program ********************************
