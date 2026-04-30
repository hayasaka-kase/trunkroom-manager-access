Attribute VB_Name = "MSZZ038"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : PDF関数
'        PROGRAM_ID      : MSZZ038
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2007/08/14
'        CERATER         : K.ISHZIAKA
'        Ver             : 0.0
'
'        UPDATE          :
'        UPDATER         :
'        Ver             :
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const MODULE_ID     As String = "MSZZ038"

Private Sub testPDF()
    Call PDFConvert("C:\Documents and Settings\さがるまーた\My Documents\加瀬（イーグル）\インターネット予約\条件のとこ.xls")
    Call PDFConvert("C:\Documents and Settings\さがるまーた\My Documents\加瀬（イーグル）\インターネット予約\条件のとこ.xls", "C:\TEMP\PDF")
    Call PDFConvert("C:\Documents and Settings\さがるまーた\My Documents\加瀬（イーグル）\インターネット予約\条件のとこ.xls", "C:\TEMP\aaa.PDF")
End Sub

'==============================================================================*
'
'       MODULE_NAME     : PDF作成
'       MODULE_ID       : PDFConvert
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : strBookName           EXCELファイル(I)
'                       : [strOutputPath]       出力先(I)省略可
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub PDFConvert(ByVal strBookName As String, Optional strOutputPath As String = "")
    Dim xlApp               As Object
    Dim xlBook              As Object
    On Error GoTo ErrorHandler

    'EXCEL起動
    Set xlApp = CreateObject("Excel.Application")
    On Error GoTo ErrorHandler2
    'EXCELファイルを開く
    Set xlBook = xlApp.Workbooks.Open(strBookName)
    On Error GoTo ErrorHandler3
    
    Call PDFConvertEx(xlBook, strBookName, strOutputPath)
    
    'EXCELファイルを閉じる
    xlBook.Close False
    On Error GoTo ErrorHandler2
    'EXCEL終了
    xlApp.Quit
    On Error GoTo ErrorHandler
Exit Sub

ErrorHandler3:
    xlBook.Close False
ErrorHandler2:
    xlApp.Visible = True
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "PDFConvert" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : PDF作成
'       MODULE_ID       : PDFConvertEx
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : objColl               印刷対象(I)
'                       : strBookName           EXCELファイル(I)
'                       : [strOutputPath]       出力先(I)省略可
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub PDFConvertEx(objColl As Object, ByVal strBookName As String, Optional strOutputPath As String = "")
    Dim strPrinterName      As String
    Dim objAbDist           As Object
    On Error GoTo ErrorHandler

    strPrinterName = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "'"))
    If strPrinterName = "" Then
        Call MSZZ024_M10("GetPrinter", "INTI_FILEの設定不足です。")
    End If
    '拡張子のない出力先ファイル名の生成
    If strOutputPath = "" Then
        strOutputPath = Replace(strBookName, Dir(strBookName), "")
    End If
    If LCase(Right(strOutputPath, 4)) <> ".pdf" Then
        If Right(strOutputPath, 1) <> "\" Then
            strOutputPath = strOutputPath & "\"
        End If
        strOutputPath = strOutputPath & Dir(strBookName)
    End If
    strOutputPath = Left$(strOutputPath, Len(strOutputPath) - 4)

    'PSファイルとして印刷
    objColl.PrintOut ActivePrinter:=strPrinterName, PrintToFile:=True, PrToFileName:=strOutputPath & ".ps"

    'PSファイルをPDFに変換
    Set objAbDist = CreateObject("PdfDistiller.PdfDistiller.1")
    objAbDist.FileToPDF strOutputPath & ".ps", strOutputPath & ".pdf", vbNullString
    'PSファイルとLOGファイルを削除
    Kill strOutputPath & ".ps"
    Kill strOutputPath & ".log"
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "PDFConvertEx" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'****************************  ended or program ********************************
