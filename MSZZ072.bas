Attribute VB_Name = "MSZZ072"
'****************************  start of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : イメージコントロール制御
'                         ハイパーリンクの設定とそのファイルの画像もしくはアイコンを表示
'       PROGRAM_ID      : MSZZ072
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2015/07/26
'       CERATER         : K.ISHIZAKA
'       Ver             : 0.0
'
'       UPDATE          : 2021/09/07
'       UPDATER         : N.IMAI
'       Ver             : 0.1
'                       : png、jpegの表示方法を変更
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'ファイルオブジェクトに関する情報を定義する構造体
Private Type SHFILEINFOA
    hIcon                       As Long
    iIcon                       As Long
    dwAttributes                As Long
    szDisplayName(260 - 1)      As Byte
    szTypeName(80 - 1)          As Byte
End Type

'ファイルシステムオブジェクトの情報を取得
Private Declare Function SHGetFileInfo Lib "Shell32" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFOA, _
    ByVal cbFileInfo As Long, _
    ByVal uFlags As Long) As Long

'SHGetFileInfo() API 関連の定数
Private Const SHGFI_ICON        As Long = &H100&
Private Const SHGFI_LARGEICON   As Long = &H0&
Private Const SHGFI_SMALLICON   As Long = &H1&
Private Const SHGFI_USEFILEATTRIBUTES As Long = &H10&

'アイコンの解放
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

'取得したいアイコンサイズの指定（インテリセンス機能用）
Public Enum ICONSIZE
    Size32x32 = SHGFI_LARGEICON      '大きいアイコン
    Size16x16 = SHGFI_SMALLICON      '小さいアイコン
End Enum

'ビットマップファイルヘッダー構造体
Private Type BITMAPFILEHEADER
       bfType       As String * 2
       bfSize       As Long
       bfReserved1  As Integer
       bfReserved2  As Integer
       bfOffBits    As Long
End Type

'ビットマップ情報構造体
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
End Type

'ビットマップ描画関数群
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pbmi As BITMAPINFO, _
                                                       ByVal iUsage As Long, ByVal ppvBits As Long, _
                                                       ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, _
                                                ByVal nStartScan As Long, ByVal nNumScans As Long, _
                                                lpBits As Any, lpBI As BITMAPINFO, _
                                                ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hGdiObj As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function CreatePen Lib "gdi32.dll" (ByVal fnPenStyle As Long, ByVal nWidth As Long, _
                                                    ByVal crColor As Long) As Integer
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

'アイコンの描画
Private Declare Function DrawIconEx Lib "user32" _
    (ByVal hdc As Long, _
     ByVal xLeft As Long, ByVal yTop As Long, _
     ByVal hIcon As Long, _
     ByVal cxWidth As Long, ByVal cyWidth As Long, _
     ByVal istepIfAniCur As Long, _
     ByVal hbrFlickerFreeDraw As Long, _
     ByVal diFlags As Long) As Long
'DrawIcon, DrawIconEx() API 関連の定数
Private Const DI_NORMAL As Long = &H3

'描画領域
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'文字の書き込み
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, _
                                                                     ByVal nWidth As Long, _
                                                                     ByVal nEscapement As Long, _
                                                                     ByVal nOrientation As Long, _
                                                                     ByVal fnWeight As Long, _
                                                                     ByVal IfdwItalic As Long, _
                                                                     ByVal fdwUnderline As Long, _
                                                                     ByVal fdwStrikeOut As Long, _
                                                                     ByVal fdwCharSet As Long, _
                                                                     ByVal fdwOutputPrecision As Long, _
                                                                     ByVal fdwClipPrecision As Long, _
                                                                     ByVal fdwQuality As Long, _
                                                                     ByVal fdwPitchAndFamily As Long, _
                                                                     ByVal lpszFace As String) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long
Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Integer


'==============================================================================*
'
'       MODULE_NAME     : ハイパーリンクの設定とそのファイルの画像もしくはアイコンを表示
'       MODULE_ID       : SetHyperlinkDrawIcon
'       CREATE_DATE     : 2015/07/26            K.ISHIZAKA
'       PARAM           : imgCtrl               イメージ(I)
'                       : strLinkPath           ハイパーリンクパス名(I)
'                       : strLinkFile           ハイパーリンクファイル名(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetHyperlinkImage(imgCtrl As Image, ByVal strLinkPath As String, ByVal strLinkFile As String)
    On Error GoTo ErrorHandler
    
    Select Case LCase(Mid(strLinkFile, InStrRev(strLinkFile, ".")))
    'Case ".wmf", ".emf", ".dib", ".bmp", ".ico", ".cgm", ".eps", ".gif", ".jpg", ".pct", ".jng", ".wpg"                    'DELETE 2021/09/07 N.IMAI
    Case ".wmf", ".emf", ".dib", ".bmp", ".ico", ".cgm", ".eps", ".gif", ".jpg", ".pct", ".jng", ".wpg", ".jpeg", ".png"    'INSERT 2021/09/07 N.IMAI
        imgCtrl.HyperlinkAddress = strLinkPath & strLinkFile
        imgCtrl.SizeMode = acOLESizeZoom
        imgCtrl.Picture = strLinkPath & strLinkFile
    Case Else
        imgCtrl.SizeMode = acOLESizeClip
        Call SetHyperlinkDrawIcon(imgCtrl, strLinkPath, strLinkFile, Size32x32)
    End Select
Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "SetHyperlinkImage" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ハイパーリンクの設定とそのファイルのアイコンを表示
'       MODULE_ID       : SetHyperlinkDrawIcon
'       CREATE_DATE     : 2015/07/26            K.ISHIZAKA
'       PARAM           : imgCtrl               イメージ(I)
'                       : strLinkPath           ハイパーリンクパス名(I)
'                       : strLinkFile           ハイパーリンクファイル名(I)
'                       : icSize                アイコンサイズ(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetHyperlinkDrawIcon(imgCtrl As Image, ByVal strLinkPath As String, ByVal strLinkFile As String, ByVal icSize As ICONSIZE)
    Dim strTempBmpFile      As String
    Dim udtShellFileInfo    As SHFILEINFOA
    On Error GoTo ErrorHandler

    'BMP一時ファイル名作成
    strTempBmpFile = Environ("temp") & "\MSZZ072.bmp"
    
    'ファイルに関するアイコン情報を取得
    Call SHGetFileInfo(strLinkPath & strLinkFile, 0&, udtShellFileInfo, _
                       Len(udtShellFileInfo), _
                       SHGFI_ICON Or icSize Or SHGFI_USEFILEATTRIBUTES)
    On Error GoTo ErrorHandler1
    Call CreateBitmapFile(strTempBmpFile, udtShellFileInfo.hIcon, imgCtrl.WIDTH, imgCtrl.HEIGHT, strLinkFile)
    'アイコンリソース解放
    Call DestroyIcon(udtShellFileInfo.hIcon)
    On Error GoTo ErrorHandler
    
    'イメージコントロールに設定
    imgCtrl.HyperlinkAddress = strLinkPath & strLinkFile
    imgCtrl.Picture = strTempBmpFile
    
    'BMP一時ファイル削除
    Kill strTempBmpFile
Exit Sub

ErrorHandler1:
    Call DestroyIcon(udtShellFileInfo.hIcon)
ErrorHandler:
    Call Err.Raise(Err.Number, "SetHyperlinkDrawIcon" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : アイコンをビットマップとして保存
'       MODULE_ID       : CreateBitmapFile
'       CREATE_DATE     : 2015/07/26            K.ISHIZAKA
'       PARAM           : strOutputFile         ファイル名(I)
'                       : hIcon                 アイコンハンドル(I)
'                       : lngWidth              幅(I)
'                       : lngHeight             高さ(I)
'                       : strText               書き込み文字列(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CreateBitmapFile(ByVal strOutputFile As String, ByVal hIcon As Long, ByVal lngWidth As Long, ByVal lngHeight As Long, ByVal strText As String)
    Dim strFile             As String
    Dim hGdiObj             As Long
    Dim hDC0                As Long
    Dim hDC1                As Long
    Dim bmpInfo             As BITMAPINFO
    Dim bytBits()           As Byte
    On Error GoTo ErrorHandler
    
    With bmpInfo.bmiHeader
        .biSize = 40    '←不明
        .biWidth = lngWidth
        .biHeight = lngHeight
        .biPlanes = 1
        .biBitCount = 24 '32でも24でも8でも4でも好きなように(8だと256色、4だと16色しか使えないけど)
    End With
    '初期処理
    hDC0 = GetDC(0&)
    On Error GoTo ErrorHandler1
    hDC1 = CreateCompatibleDC(hDC0)
    On Error GoTo ErrorHandler2
    hGdiObj = CreateDIBSection(hDC1, bmpInfo, 0, 0, 0, 0)
    On Error GoTo ErrorHandler3
    Call SelectObject(hDC1, hGdiObj)
    'まずは背景を白色に
    Call SetWhiteBackColor(hDC1, lngWidth, lngHeight)
    'アイコンを書き込み
    Call DrawIconEx(hDC1, lngWidth \ 2 - 16, lngHeight \ 2 - 16, hIcon, 0, 0, 0, 0, DI_NORMAL)
    'ファイル名を書き込み
    Call DrawFileName(hDC1, lngWidth, lngHeight, strText)
    
    '終了処理
    Call GetDIBits(hDC1, hGdiObj, 0, lngHeight, ByVal 0&, bmpInfo, 0)
    ReDim bytBits(bmpInfo.bmiHeader.biSizeImage - 1)
    Call GetDIBits(hDC1, hGdiObj, 0, lngHeight, bytBits(0), bmpInfo, 0)
    'ファイル出力
    Call OutputBitmapFile(strOutputFile, bmpInfo, bytBits)
    'リソース解放
    Call DeleteObject(hGdiObj)
    On Error GoTo ErrorHandler2
    Call DeleteObject(hDC1)
    On Error GoTo ErrorHandler1
    Call ReleaseDC(0&, hDC0)
Exit Sub

ErrorHandler3:
    Call DeleteObject(hGdiObj)
ErrorHandler2:
    Call DeleteObject(hDC1)
ErrorHandler1:
    Call ReleaseDC(0&, hDC0)
ErrorHandler:
    Call Err.Raise(Err.Number, "OutputBitmapFile" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : アイコンをビットマップとして保存
'       MODULE_ID       : CreateBitmapFile
'       CREATE_DATE     : 2015/07/26            K.ISHIZAKA
'       PARAM           : strOutputFile         ファイル名(I)
'                       : hIcon                 アイコンハンドル(I)
'                       : lngWidth              幅(I)
'                       : lngHeight             高さ(I)
'                       : strText               書き込み文字列(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub DrawFileName(ByVal hDC1 As Long, ByVal lngWidth As Long, ByVal lngHeight As Long, ByVal strText As String)
    Dim hFont               As Long
    Dim stRct               As RECT
    Dim lpszFace            As String
    'CreateFont()用定数
    Const FW_NORMAL             As Long = 400
    'Const FW_BOLD               As Long = 700
    Const DEFAULT_CHARSET       As Long = 1
    Const OUT_DEFAULT_PRECIS    As Long = 0
    Const CLIP_DEFAULT_PRECIS   As Long = 0
    Const DEFAULT_QUALITY       As Long = 0
    Const DEFAULT_PITCH         As Long = 0
    Const FF_SCRIPT             As Long = 64
    'DrawText()用定数
    Const DT_CENTER             As Long = &H1
    Const DT_SINGLELINE         As Long = &H20
    On Error GoTo ErrorHandler

    hFont = CreateFont(12, 0, 0, 0, FW_NORMAL, _
            0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
            CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, _
            DEFAULT_PITCH Or FF_SCRIPT, lpszFace)
    On Error GoTo ErrorHandler1
    Call SelectObject(hDC1, hFont)

    stRct.Left = 10 'ワクを決めて
    stRct.Top = lngHeight \ 2 + 16
    stRct.Right = lngWidth - 10
    stRct.Bottom = lngHeight - 10
    Call SetTextColor(hDC1, RGB(0, 0, 0)) '文字色を決めて
    Call SetBkColor(hDC1, RGB(255, 255, 255)) '背景色を決めて
    Call DrawText(hDC1, strText, -1, stRct, DT_CENTER Or DT_SINGLELINE)

    Call DeleteObject(hFont)
Exit Sub

ErrorHandler1:
    Call DeleteObject(hFont)
ErrorHandler:
    Call Err.Raise(Err.Number, "DrawFileName" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ビットマップファイル作成
'       MODULE_ID       : OutputBitmapFile
'       CREATE_DATE     : 2015/07/26            K.ISHIZAKA
'       PARAM           : strOutputFile         ファイル名(I)
'                       : bmpInfo               ビットマップイメージ(I)
'                       : bytBits               終端制御文字(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub OutputBitmapFile(ByVal strOutputFile As String, bmpInfo As BITMAPINFO, bytBits() As Byte)
    Dim iFileNo             As Integer
    Dim bmpFileHd           As BITMAPFILEHEADER
    On Error GoTo ErrorHandler
    
    With bmpFileHd
        .bfType = "BM"
        .bfReserved1 = 0
        .bfReserved2 = 0
        .bfSize = Len(bmpFileHd) + Len(bmpInfo) + UBound(bytBits) + 1
        .bfOffBits = Len(bmpFileHd) + Len(bmpInfo)
    End With
    
    iFileNo = FreeFile()
    Open strOutputFile For Binary As #iFileNo
    On Error GoTo ErrorHandler1
    Put #iFileNo, , bmpFileHd
    Put #iFileNo, , bmpInfo
    Put #iFileNo, , bytBits
    Close #iFileNo
Exit Sub

ErrorHandler1:
    Close #iFileNo
ErrorHandler:
    Call Err.Raise(Err.Number, "OutputBitmapFile" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 背景色を白色に設定する
'       MODULE_ID       : SetWhiteBackColor
'       CREATE_DATE     : 2015/07/26            K.ISHIZAKA
'       PARAM           : hDC                   描画領域ハンドル(I)
'                       : lngWidth              幅(I)
'                       : lngHeight             高さ(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetWhiteBackColor(ByVal hdc As Long, ByVal lngWidth As Long, ByVal lngHeight As Long)
    Dim hPen                As Long
    Dim hBrush              As Long
    Const PS_SOLID          As Long = 0
    Const WHITE_BRUSH       As Long = 0
    On Error GoTo ErrorHandler

    hPen = CreatePen(PS_SOLID, 0, RGB(255, 255, 255))
    On Error GoTo ErrorHandler1
    Call SelectObject(hdc, hPen)
    hBrush = GetStockObject(WHITE_BRUSH)
    On Error GoTo ErrorHandler2
    Call SelectObject(hdc, hBrush)
    Call Rectangle(hdc, 0, 0, lngWidth, lngHeight)
    Call DeleteObject(hBrush)
    On Error GoTo ErrorHandler1
    Call DeleteObject(hPen)
    On Error GoTo ErrorHandler
Exit Sub

ErrorHandler2:
    Call DeleteObject(hBrush)
ErrorHandler1:
    Call DeleteObject(hPen)
ErrorHandler:
    Call Err.Raise(Err.Number, "SetWhiteBackColor" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'****************************  ended of program ********************************
