Attribute VB_Name = "MNN020"
'****************************  strat of program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : 緯度経度距離
'        PROGRAM_ID      : MNN020
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2006/02/13
'        CERATER         : K.ISHZIAKA
'        Ver             : 0.0
'
'        UPDATE          : 2007/02/16
'        UPDATER         : N.MIURA
'        Ver             : 0.1
'                          WebGoogleOpen
'                          GOOGLEマップ表示用ＨＰアドレスの取得方法の変更
'
'        UPDATE          : 2012/10/11
'        UPDATER         : K.ISHZIAKA
'        Ver             : 0.2
'                          WebGoogleMarker 地図表示の高速化対応
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   外部関数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWMAXIMIZED  As Long = 3

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID   As String = "MNN020"

'日本測地系（Tokyo Datum）
'Private Const a         As Double = 6377397.155         '長半径(メートル)
'Private Const f         As Double = (1 / 299.152813)    '扁平率

'世界測地系（日本測地系2000）
Private Const a         As Double = 6378137             '長半径(メートル)
Private Const F         As Double = (1 / 298.257222101) '扁平率

Private Const π        As Double = 3.14159265358979    '円周率
Private Const m0        As Double = 0.9999              '座標系の原点における縮尺係数

'ダイアログフォーム名
Private Const C_DIALOG_NAME As String = "FNN020"

'一時テーブル名
Public Const WK_TABLE_NAME  As String = "#WK_MNN020_TEMP"

'ADO Constant
Private Const adCmdText = 1
Private Const adCmdTable = 2
Private Const adOpenForwardOnly = 0
Private Const adOpenKeyset = 1
Private Const adLockReadOnly = 1
Private Const adLockPessimistic = 2

'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private strParentFormName   As String   'フォーム
Private strParentBangoName  As String   'テキスト座標系番号
Private strParentIdoName    As String   'テキスト緯度
Private strParentKeidoName  As String   'テキスト経度
Private strParentBangoVal   As String   'テキスト座標系番号
Private strParentIdoVal     As String   'テキスト緯度
Private strParentKeidoVal   As String   'テキスト経度
Private strParentAddrVal    As String   'テキスト住所
Private strParentInitAddr   As String

'==============================================================================*
'
'       MODULE_NAME     : テスト用
'       MODULE_ID       : TEST_ShowDialog
'       CREATE_DATE     : 2006/02/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub TEST_ShowDialog()
    'Call ShowDialog("", "", "", "", "", "", "")
    'Call ShowDialog("", "", "", "", "北海道札幌市", "10", "")
    Call ShowDialog("", "", "", "", "東京都千代田区神田神保町", "２－１３", "××ビル")
    'Call ShowDialog("", "", "", "", "東京都千代田区神田神保町", "2-13", "××ビル")
'    Call ShowDialog("", "", "", "", "神奈川県横浜市鶴見区三ッ池公園", "１", "")
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ダイアログ表示
'       MODULE_ID       : ShowDialog
'       CREATE_DATE     : 2006/02/13
'       PARAM           : strFormName           フォーム名(I)
'                       : strBangoName          座標系番号のコントロール名(I)
'                       : strIdoName            緯度のコントロール名(I)
'                       : strKeidoName          経度のコントロール名(I)
'                       : strAddrVal1           住所１(I)
'                       : strAddrVal2           住所２(I)
'                       : strAddrVal3           住所３(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub ShowDialog(ByVal strFormName As String, ByVal strBangoName As String, ByVal strIdoName As String, ByVal strKeidoName As String, ByVal strAddrVal1 As String, ByVal strAddrVal2 As String, ByVal strAddrVal3 As String)
    Dim i                   As Integer
    Dim strNumb             As String
    On Error GoTo ErrorHandler
    
    If strFormName <> "" Then
        With Forms(strFormName)
            strParentBangoVal = Nz(.Controls(strBangoName).VALUE)
            strParentIdoVal = Nz(.Controls(strIdoName).VALUE)
            strParentKeidoVal = Nz(.Controls(strKeidoName).VALUE)
        End With
    End If
    strParentAddrVal = strAddrVal1 & strAddrVal2 & strAddrVal3
    If InStr(2, strAddrVal2, "－", vbDatabaseCompare) = 2 Then
        strParentInitAddr = strAddrVal1
        strNumb = NumbConv(Left(strAddrVal2, 1), 3)
        If strNumb <> "" Then
            strParentInitAddr = strParentInitAddr & strNumb & "丁目"
            i = 3
        Else
            i = 0
        End If
    Else
        strParentInitAddr = strAddrVal1
        i = 1
    End If
    If i > 0 Then
        strNumb = NumbConv(Mid(strAddrVal2, i, 1))
        While strNumb <> ""
            strParentInitAddr = strParentInitAddr & strNumb
            i = i + 1
            strNumb = NumbConv(Mid(strAddrVal2, i, 1))
        Wend
    End If
    strParentFormName = strFormName
    strParentBangoName = strBangoName
    strParentIdoName = strIdoName
    strParentKeidoName = strKeidoName
    
    doCmd.OpenForm C_DIALOG_NAME, acNormal, , , , acDialog
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "ShowDialog" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : データ読み込み
'       MODULE_ID       : NumbConv
'       CREATE_DATE     : 2006/02/13
'       PARAM           : strNumb               数字文字(I)
'       RETURN          : 半角数字
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function NumbConv(ByVal strNumb As String, Optional iRet As Integer = 1) As String
    On Error GoTo ErrorHandler

    Select Case strNumb
    Case "0", "０", "〇"
        NumbConv = Choose(iRet, "0", "０", "〇")
    Case "1", "１", "一"
        NumbConv = Choose(iRet, "1", "１", "一")
    Case "2", "２", "二"
        NumbConv = Choose(iRet, "2", "２", "二")
    Case "3", "３", "三"
        NumbConv = Choose(iRet, "3", "３", "三")
    Case "4", "４", "四"
        NumbConv = Choose(iRet, "4", "４", "四")
    Case "5", "５", "五"
        NumbConv = Choose(iRet, "5", "５", "五")
    Case "6", "６", "六"
        NumbConv = Choose(iRet, "6", "６", "六")
    Case "7", "７", "七"
        NumbConv = Choose(iRet, "7", "７", "七")
    Case "8", "８", "八"
        NumbConv = Choose(iRet, "8", "８", "八")
    Case "9", "９", "九"
        NumbConv = Choose(iRet, "9", "９", "九")
    Case Else
        NumbConv = ""
    End Select
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "NumbConv" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : データ読み込み
'       MODULE_ID       : SetData
'       CREATE_DATE     : 2006/02/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetData() As String
    On Error GoTo ErrorHandler
    
    With Forms(C_DIALOG_NAME)
        .txtKEIBI.VALUE = strParentBangoVal
        .txtIDOPO.VALUE = strParentIdoVal
        .txtKEIPO.VALUE = strParentKeidoVal
        .txtAddress.VALUE = strParentAddrVal
    End With
    GetData = strParentInitAddr
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "GetData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : データ書き込み
'       MODULE_ID       : SetData
'       CREATE_DATE     : 2006/02/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetData()
    On Error GoTo ErrorHandler
    
    If strParentFormName <> "" Then
        With Forms(strParentFormName)
            .Controls(strParentBangoName).VALUE = Forms(C_DIALOG_NAME).txtKEIBI.VALUE
            .Controls(strParentIdoName).VALUE = Forms(C_DIALOG_NAME).txtIDOPO.VALUE
            .Controls(strParentKeidoName).VALUE = Forms(C_DIALOG_NAME).txtKEIPO.VALUE
        End With
    End If
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "SetData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 短半径
'       MODULE_ID       : b
'       CREATE_DATE     : 2006/02/13
'       RETURN          : 短半径
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function b() As Double
    On Error GoTo ErrorHandler
    b = a * (1 - F)
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "b" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 極での曲率半径
'       MODULE_ID       : c
'       CREATE_DATE     : 2006/02/13
'       RETURN          : 極での曲率半径
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function c() As Double
    On Error GoTo ErrorHandler
    c = a / (1 - F)
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "c" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 第一離心率の二乗
'       MODULE_ID       : e2
'       CREATE_DATE     : 2006/02/13
'       RETURN          : 第一離心率の二乗
'       計算精度を上げるため二乗を使用する（ルートを使用しない）
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function e2() As Double
    On Error GoTo ErrorHandler
    e2 = 2 * F - (F ^ 2)
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "e2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 第二離心率の二乗
'       MODULE_ID       : e_dash2
'       CREATE_DATE     : 2006/02/13
'       RETURN          : 第二離心率の二乗
'       計算精度を上げるため二乗を使用する（ルートを使用しない）
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function e_dash2() As Double
    On Error GoTo ErrorHandler
    e_dash2 = e2 / (1 - e2)
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "e_dash2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ラジアン値変換
'       MODULE_ID       : Radian
'       CREATE_DATE     : 2006/02/13
'       PARAM           : dd                    度(I)
'                       : mm                    分(I)
'                       : sss                   秒(I)
'       RETURN          : ラジアン値
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Radian(ByVal dd As Double, Optional ByVal mm As Double = 0, Optional ByVal sss As Double = 0) As Double
    On Error GoTo ErrorHandler
    Radian = (dd + mm / 60 + sss / 3600) / 180 * π
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Radian" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 座標系原点の緯度
'       MODULE_ID       : Getφ0
'       CREATE_DATE     : 2006/02/13
'       PARAM           : iCourse               座標系番号(I)
'       RETURN          : 座標系原点の緯度（ラジアン値）
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Getφ0(ByVal iCourse As Integer) As Double
    On Error GoTo ErrorHandler
    Getφ0 = Radian(Choose(iCourse, 33, 33, 36, 33, 36, 36, 36, 36, 36, 40, 44, 44, 44, 26, 26, 26, 26, 26, 26))
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Getφ0" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 座標系原点の経度
'       MODULE_ID       : Getλ0
'       CREATE_DATE     : 2006/02/13
'       PARAM           : iCourse               座標系番号(I)
'       RETURN          : 座標系原点の経度（ラジアン値）
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Getλ0(ByVal iCourse As Integer) As Double
    On Error GoTo ErrorHandler
    Getλ0 = Radian( _
        Choose(iCourse, 129, 131, 132, 133, 134, 136, 137, 138, 139, 140, 140, 142, 144, 142, 127, 124, 131, 136, 154), _
        Choose(iCourse, 30, 0, 10, 30, 20, 0, 10, 30, 50, 50, 15, 15, 15, 0, 30, 0, 0, 0, 0))
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Getλ0" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 平面直角座標Ｘ
'       MODULE_ID       : Xzahyou
'       CREATE_DATE     : 2006/02/13
'       PARAM           : iCourse               座標系番号(I)
'                       : φ                    緯度(I)
'                       : λ                    経度(I)
'       RETURN          : 平面直角座標Ｘ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Xzahyou(ByVal iCourse As Integer, ByVal φ As Double, ByVal λ As Double) As Double
    Dim ans                 As Double
    Dim i                   As Integer
    Dim t                   As Double
    Dim Δλ                As Double
    Dim η2                 As Double
    Dim φ0                 As Double
    Dim λ0                 As Double
    Dim S                   As Double   '赤道から緯度φまでの子午線弧長
    Dim S0                  As Double   '赤道から座標系の原点の緯度φ0までの子午線弧長
    Dim N                   As Double   '卯酉線曲率半径
    Dim chan(1 To 4)        As Double
    On Error GoTo ErrorHandler
    
    '===========
    '座標系原点の経緯度
    φ0 = Getφ0(iCourse)
    λ0 = Getλ0(iCourse)
    '===========
    '(4)縮尺係数
    Δλ = λ - λ0
    η2 = e_dash2 * (Cos(φ) ^ 2)
    t = Tan(φ)
    '===========
    '5.赤道からの子午線弧長
    S = ShigosenKotyou(φ)
    S0 = ShigosenKotyou(φ0)
    '===========
    '卯酉線曲率半径
    N = a / Sqr(1 - e2 * (Sin(φ) ^ 2))
    '===========
    chan(1) = 1
    chan(2) = 5 - (t ^ 2) + 9 * η2 + 4 * (η2 ^ 2)
    chan(3) = -61 + 58 * (t ^ 2) - (t ^ 4) - 270 * η2 + 330 * (t ^ 2) * η2
'    chan(4) = -1385 + 3111 * (t ^ 2) - 543 * (t ^ 4) + (t ^ 6) '←公開されている式はこれ
    chan(4) = -1385 + 3111 * (t ^ 2) - 543 ^ (t ^ 4) + (t ^ 6)  '←公開されている答えにあう式はこれ
    '===========
    ans = (S - S0)
    For i = 1 To 4
        ans = ans + (1 / Choose(i, 2, 24, -720, -40320)) * N * (Cos(φ) ^ (2 * i)) * t * chan(i) * (Δλ ^ (2 * i))
    Next
    Xzahyou = ans * m0
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Xzahyou" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 平面直角座標Ｙ
'       MODULE_ID       : Yzahyou
'       CREATE_DATE     : 2006/02/13
'       PARAM           : iCourse               座標系番号(I)
'                       : φ                    緯度(I)
'                       : λ                    経度(I)
'       RETURN          : 平面直角座標Ｙ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Yzahyou(ByVal iCourse As Integer, ByVal φ As Double, ByVal λ As Double) As Double
    Dim ans                 As Double
    Dim i                   As Integer
    Dim t                   As Double
    Dim Δλ                As Double
    Dim η2                 As Double
    Dim λ0                 As Double
    Dim N                   As Double   '卯酉線曲率半径
    Dim chan(1 To 4)        As Double
    On Error GoTo ErrorHandler
    
    '===========
    '座標系原点の経緯度
    λ0 = Getλ0(iCourse)
    '===========
    '(4)縮尺係数
    Δλ = λ - λ0
    η2 = e_dash2 * (Cos(φ) ^ 2)
    t = Tan(φ)
    '===========
    '卯酉線曲率半径
    N = a / Sqr(1 - e2 * (Sin(φ) ^ 2))
    '===========
    chan(1) = 1
    chan(2) = -1 + (t ^ 2) - η2
    chan(3) = -5 + 18 * (t ^ 2) - (t ^ 4) - 14 * η2 + 58 * (t ^ 2) * η2
    chan(4) = -61 + 479 * (t ^ 2) - 179 * (t ^ 4) + (t ^ 6)
    '===========
    
    ans = 0
    For i = 1 To 4
        ans = ans + (1 / Choose(i, 1, -6, -120, -5040)) * N * Cos(φ) ^ (i * 2 - 1) * chan(i) * Δλ ^ (i * 2 - 1)
    Next
    Yzahyou = ans * m0
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Yzahyou" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 測地線長
'       MODULE_ID       : KmFromZahyou
'       CREATE_DATE     : 2006/02/13
'       PARAM           : iCourse               座標系番号(I)
'                       : y1                    始点のＹ座標(I)
'                       : x1                    始点のＸ座標(I)
'                       : y2                    終点のＹ座標(I)
'                       : x2                    終点のＸ座標(I)
'       RETURN          : 測地線長（キロメートル）
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function KmFromZahyou(ByVal iCourse As Integer, ByVal y1 As Double, ByVal x1 As Double, ByVal y2 As Double, ByVal x2 As Double) As Double
    Dim ans                 As Double
    Dim φ0                 As Double
    Dim R0                  As Double
    On Error GoTo ErrorHandler
    
    '===========
    '座標系原点の経緯度
    φ0 = Getφ0(iCourse)
    '===========
    R0 = (a * Sqr(1 - e2)) / (1 - e2 * Sin(φ0) ^ 2)
    '===========
    ans = 1 + 1 / (6 * (R0 ^ 2) * (m0 ^ 2)) * ((y1 ^ 2) + y1 * y2 + (y2 ^ 2))
    ans = Sqr(((x2 - x1) ^ 2) + ((y2 - y1) ^ 2)) / (m0 * ans)
    KmFromZahyou = Format(ans / 1000, "0.0000")
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "KmFromZahyou" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 子午線弧長
'       MODULE_ID       : ShigosenKotyou
'       CREATE_DATE     : 2006/02/13
'       PARAM           : φ                    緯度(I)
'       RETURN          : 子午線弧長
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ShigosenKotyou(ByVal φ As Double) As Double
    Dim i                   As Integer
    Dim ans                 As Double
    Dim bunsi(0 To 8)       As Variant
    Dim bunbo(0 To 8)       As Variant
    On Error GoTo ErrorHandler

    'A
    bunsi(0) = Array(3, 45, 175, 11025, 43659, 693693, 19324305, 4927697775#)
    bunbo(0) = Array(4, 64, 256, 16384, 65536, 1048576, 29360128, 7516192768#)
    'B
    bunsi(1) = Array(3, 15, 525, 2205, 72765, 297297, 135270135, 547521975)
    bunbo(1) = Array(4, 16, 512, 2048, 65536, 262144, 117440512, 469762048)
    'C
    bunsi(2) = Array(0, 15, 105, 2205, 10395, 1486485, 45090045, 766530765)
    bunbo(2) = Array(1, 64, 256, 4096, 16384, 2097152, 58720256, 939524096)
    'D
    bunsi(3) = Array(0, 0, 35, 315, 31185, 165165, 45090045, 209053845)
    bunbo(3) = Array(1, 1, 512, 2048, 131072, 524288, 117440512, 469762048)
    'E
    bunsi(4) = Array(0, 0, 0, 315, 3465, 99099, 4099095, 348423075)
    bunbo(4) = Array(1, 1, 1, 16384, 65536, 1048576, 29360128, 1879048192)
    'F
    bunsi(5) = Array(0, 0, 0, 0, 693, 9009, 4099095, 26801775)
    bunbo(5) = Array(1, 1, 1, 1, 131072, 524288, 117440512, 469762048)
    'G
    bunsi(6) = Array(0, 0, 0, 0, 0, 3003, 315315, 11486475)
    bunbo(6) = Array(1, 1, 1, 1, 1, 2097152, 58720256, 939524096)
    'H
    bunsi(7) = Array(0, 0, 0, 0, 0, 0, 45045, 765765)
    bunbo(7) = Array(1, 1, 1, 1, 1, 1, 117440512, 469762048)
    'I
    bunsi(8) = Array(0, 0, 0, 0, 0, 0, 0, 765765)
    bunbo(8) = Array(1, 1, 1, 1, 1, 1, 1, 7516192768#)

    ans = (1 + ShigosenKotyouWork(bunsi(0), bunbo(0))) * φ
    For i = 1 To 8
        ans = ans + ShigosenKotyouWork(bunsi(i), bunbo(i)) / Choose((i And 1) + 1, 2 * i, -2 * i) * Sin(2 * i * φ)
    Next
    
    ShigosenKotyou = a * (1 - e2) * ans
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "ShigosenKotyou" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 子午線弧長定型算
'       MODULE_ID       : ShigosenKotyouWork
'       CREATE_DATE     : 2006/02/13
'       PARAM           : argBunsi              分子(I)
'                       : argBunbo              分母(I)
'       RETURN          : 子午線弧長定数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ShigosenKotyouWork(ByVal argBunsi As Variant, ByVal argBunbo As Variant) As Double
    Dim i                   As Integer
    Dim ans                 As Double
    Dim bunsi()             As Variant
    Dim bunbo()             As Variant
    On Error GoTo ErrorHandler
    
    bunsi = argBunsi
    bunbo = argBunbo
    ans = 0
    For i = 1 To 8
        ans = ans + CDbl(bunsi(i - 1)) / CDbl(bunbo(i - 1)) * (e2 ^ i)
    Next
    ShigosenKotyouWork = ans
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "ShigosenKotyouWork" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 測地線長
'       MODULE_ID       : KmFromDo
'       CREATE_DATE     : 2006/02/13
'       PARAM           : iCourse               座標系番号(I)
'                       : ido1                  始点の緯度(I)
'                       : keido1                始点の経度(I)
'                       : ido2                  終点の緯度(I)
'                       : keido2                終点の経度(I)
'       RETURN          : 測地線長
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function KmFromDo(ByVal iCourse As Integer, ByVal ido1 As Double, ByVal keido1 As Double, ByVal ido2 As Double, ByVal keido2 As Double) As Double
    Dim φ1                 As Double
    Dim λ1                 As Double
    Dim φ2                 As Double
    Dim λ2                 As Double
    Dim x1                  As Double
    Dim y1                  As Double
    Dim x2                  As Double
    Dim y2                  As Double
    On Error GoTo ErrorHandler
    
    '弧度法
    φ1 = Radian(ido1)
    λ1 = Radian(keido1)
    φ2 = Radian(ido2)
    λ2 = Radian(keido2)
    '平面直角座標
    y1 = Yzahyou(iCourse, φ1, λ1)
    x1 = Xzahyou(iCourse, φ1, λ1)
    y2 = Yzahyou(iCourse, φ2, λ2)
    x2 = Xzahyou(iCourse, φ2, λ2)
    
    KmFromDo = KmFromZahyou(iCourse, y1, x1, y2, x2)
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "KmFromDo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 緯度経度の地図表示
'       MODULE_ID       : WebPointGoogle
'       CREATE_DATE     : 2006/02/13
'       PARAM           : IEApp                 IEオブジェクト(I/O)
'                       : ido                   緯度(I)
'                       : keido                 経度(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub WebPointGoogle(IEApp As Object, ByVal IDO As Double, ByVal KEIDO As Double, Optional strAddr As String = "")
    On Error GoTo ErrorHandler

    Call WebGoogleOpen(IEApp, False)
    Call WebGoogleMarker(IEApp, IDO, KEIDO, "r", "", "位置の確認", strAddr, "")
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "WebPointGoogle" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 地図表示
'       MODULE_ID       : WebGoogleOpen
'       CREATE_DATE     : 2006/02/13
'       PARAM           : IEApp                 IEオブジェクト(I/O)
'                       : bHelp                 アイコンヘルプ表示(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub WebGoogleOpen(IEApp As Object, ByVal bHelp As Boolean)
    Dim strUrl              As String
    Dim Elmnt               As Object
    On Error GoTo ErrorHandler

    'strURL = "http://www18.ocn.ne.jp/~zaka/test3.html"                         'DELETE 20070216 N.MIURA
    strUrl = DLookup("CONT_MAP_URL", "dbo_CONT_MAST", "CONT_KEY=1 ")            'INSERT 20070216 N.MIURA
    
    'MsgBox (strURL)
    
    If (TypeName(IEApp) = "Nothing") Or (TypeName(IEApp) = "Object") Then
        Set IEApp = CreateObject("InternetExplorer.Application")
    End If
    Call IEApp.Navigate(strUrl)
    While IEApp.Busy
        Call Sleep(50)
    Wend
'    IEApp.Toolbar = False
'    IEApp.MenuBar = False
    While IEApp.Busy
        Call Sleep(50)
    Wend
    If Not bHelp Then
        Set Elmnt = IEApp.Document.getElementById("divHelp")
        Elmnt.Style.Display = "none"
        Set Elmnt = Nothing
    End If
    IEApp.Visible = True
    Call ShowWindow(IEApp.hwnd, SW_SHOWMAXIMIZED)
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "WebGoogleOpen" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 緯度経度によるマーカーの追加
'       MODULE_ID       : WebGoogleMarker
'       CREATE_DATE     : 2006/02/13
'       PARAM           : IEApp                 IEオブジェクト(I/O)
'                       : ido                   緯度(I)
'                       : keido                 経度(I)
'                       : strMark               マーク種類(I)
'                       : strYardCode           ヤードコード(I)
'                       : strYardName           ヤード名称(I)
'                       : strAddr               住所(I)
'                       : strNote               補足(I)
'                       : [blCenter]            地図位置中心あわせ(True/False)する:初期値/しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub WebGoogleMarker(IEApp As Object, ByVal IDO As Double, ByVal KEIDO As Double, _
    ByVal strMark As String, ByVal strYardCode As String, ByVal strYardName As String, _
    ByVal strAddr As String, ByVal strNote As String)                           'DELETE 2012/10/11 K.ISHIZAKA
Public Sub WebGoogleMarker(IEApp As Object, ByVal IDO As Double, ByVal KEIDO As Double, _
    ByVal strMark As String, ByVal strYardCode As String, ByVal strYardName As String, _
    ByVal strAddr As String, ByVal strNote As String, Optional ByVal blCenter As Boolean = True) 'INSERT 2012/10/11 K.ISHIZAKA
    Dim Elmnt               As Object
    On Error GoTo ErrorHandler

    Set Elmnt = IEApp.Document.getElementById("txtLng")
    Elmnt.VALUE = IDO
    Set Elmnt = Nothing

    Set Elmnt = IEApp.Document.getElementById("txtLat")
    Elmnt.VALUE = KEIDO
    Set Elmnt = Nothing

    Set Elmnt = IEApp.Document.getElementById("txtNumb")
    Elmnt.InnerHTML = strYardCode
    Set Elmnt = Nothing

    Set Elmnt = IEApp.Document.getElementById("txtName")
    Elmnt.InnerHTML = strYardName
    Set Elmnt = Nothing

    Set Elmnt = IEApp.Document.getElementById("txtAddr")
    Elmnt.InnerHTML = strAddr
    Set Elmnt = Nothing

    Set Elmnt = IEApp.Document.getElementById("txtNote")
    Elmnt.InnerHTML = strNote
    Set Elmnt = Nothing

'    Set Elmnt = IEApp.Document.getElementById("btnMarker")                     'DELETE 2012/10/11 K.ISHIZAKA
    Set Elmnt = IEApp.Document.getElementById(IIf(blCenter, "btnMarker", "btnMarke2")) 'INSERT 2012/10/11 K.ISHIZAKA
    Elmnt.VALUE = strMark
    Elmnt.Click
    Set Elmnt = Nothing
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "WebGoogleMarker" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 緯度経度の地図表示
'       MODULE_ID       : WebPointMSN
'       CREATE_DATE     : 2006/02/13
'       PARAM           : IEApp                 IEオブジェクト(I/O)
'                       : ido                   緯度(I)
'                       : keido                 経度(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub WebPointMSN(IEApp As Object, ByVal IDO As Double, ByVal KEIDO As Double)
    Dim strUrl              As String
    Dim jpIdo               As Double
    Dim jpKeido             As Double
    On Error GoTo ErrorHandler

'日本測地系（Tokyo Datum）を世界測地系（日本測地系2000）に変換する場合
'ido = jpIdo - jpIdo * 0.00010695 + jpKeido * 0.000017464 + 0.0046017
'keido = jpKeido - jpIdo * 0.000046038 - jpKeido * 0.000083043 + 0.01004
    
    '日本測地系（Tokyo Datum）に変換する
    jpIdo = IDO + IDO * 0.00010696 - KEIDO * 0.000017467 - 0.004602
    jpKeido = KEIDO + IDO * 0.000046047 + KEIDO * 0.000083049 - 0.010041

    strUrl = "http://map.msn.co.jp/mapmarking.armx?mode=1&la=" & FormatDo(jpKeido) & "&lg=" & FormatDo(jpIdo) & "&zm=12&smode=2"
    If (TypeName(IEApp) = "Nothing") Or (TypeName(IEApp) = "Object") Then
        Set IEApp = CreateObject("InternetExplorer.Application")
    End If
    Call IEApp.Navigate(strUrl)
    While IEApp.Busy
        Call Sleep(50)
    Wend
    IEApp.Visible = True
Exit Sub
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "WebPointMSN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 度を点区切りの度分秒にする
'       MODULE_ID       : FormatDo
'       CREATE_DATE     : 2006/02/13
'       PARAM           : dd0                   度(I)
'       RETURN          : （例）139.44.41.0
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function FormatDo(ByVal dd0 As Double) As String
    Dim dd                  As Double
    Dim mm                  As Double
    Dim sss                 As Double
    On Error GoTo ErrorHandler
    
    dd = Int(dd0)
    mm = Int((dd0 - dd) * 60)
    sss = (dd0 - dd - mm / 60) * 3600

'    FormatDo = Format(dd, "0") & "ﾟ" & Format(mm, "00") & "'" & Replace(Format(sss, "00.0000"), ".", """")
    FormatDo = Format(dd, "0") & "." & Format(mm, "00") & "." & Format(sss, "0.0")
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "FormatDo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 指定ヤードから範囲内のヤードを抽出
'       MODULE_ID       : GetNearYard
'       CREATE_DATE     : 2006/02/13
'       PARAM           : dbSqlServer           SQL-ServerのADODB.Connection(I)
'                       : strYardCode           ヤードコード(I)
'                       : dblKiro               含みたいキロメートル(I)
'       RETURN          : 件数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetNearYard(dbSQLServer As Object, ByVal strYardCode As String, ByVal dblKiro As Double) As Long
    Dim dblIdo              As Double
    Dim dblKeido            As Double
    Dim rsSqlServer         As Object
    On Error GoTo ErrorHandler
    
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    With rsSqlServer
        .Open fncSelectYard(strYardCode), dbSQLServer, adOpenForwardOnly, adLockReadOnly, adCmdText
        On Error GoTo ErrorHandler1
        If Not .EOF Then
            dblIdo = .Fields("YARD_IDO")
            dblKeido = .Fields("YARD_KEIDO")
        End If
        .Close
    End With
    On Error GoTo ErrorHandler
    If dblIdo > 0 Then
        GetNearYard = GetNearPoint(dbSQLServer, strYardCode, dblIdo, dblKeido, dblKiro)
    Else
        GetNearYard = 0
    End If
Exit Function
    
ErrorHandler1:
    rsSqlServer.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "GetNearYard" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ヤードの経緯度を抽出
'       MODULE_ID       : fncSelectYard
'       CREATE_DATE     : 2006/02/13
'       PARAM           : strYardCode           ヤードコード(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectYard(ByVal strYardCode As String) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " YARD_IDO,"
    strSQL = strSQL & " YARD_KEIDO "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " YARD_MAST "
    strSQL = strSQL & "WHERE YARD_CODE = " & strYardCode & " "
    strSQL = strSQL & "AND   YARD_IDO   > 0 "
    strSQL = strSQL & "AND   YARD_KEIDO > 0 "
    
    fncSelectYard = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncSelectYard" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 基準点から範囲内のヤードを抽出
'       MODULE_ID       : GetNearPoint
'       CREATE_DATE     : 2006/02/13
'       PARAM           : dbSqlServer           SQL-ServerのADODB.Connection(I)
'                       : strYcode              ヤードコード(I)
'                       : dblIdo                緯度(I)
'                       : dblKeido              経度(I)
'                       : dblKiro               含みたいキロメートル(I)
'       RETURN          : 件数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetNearPoint(dbSQLServer As Object, ByVal strYcode As String, ByVal dblIdo As Double, ByVal dblKeido As Double, ByVal dblKiro As Double) As Long
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = fncCreateTempTable()
    dbSQLServer.Execute strSQL, adCmdText
    strSQL = fncInsertTempTable(strYcode, dblIdo, dblKeido, dblKiro)
    GetNearPoint = fncUpdate_TEMP_KIRO(dbSQLServer, strSQL, dblIdo, dblKeido, dblKiro)
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "GetNearPoint" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 基準点から範囲内のヤードを抽出
'       MODULE_ID       : fncUpdate_TEMP_KIRO
'       CREATE_DATE     : 2006/02/13
'       PARAM           : dbSqlServer           SQL-ServerのADODB.Connection(I)
'                       : strSql                SQL文(I)
'                       : dblKiro               含みたいキロメートル(I)
'       RETURN          : 件数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdate_TEMP_KIRO(dbSQLServer As Object, ByVal strSQL As String, ByVal dblIdo As Double, ByVal dblKeido As Double, ByVal dblKiro As Double) As Long
    Dim lngCount            As Long
    Dim wkKiro              As Double
    Dim rsSqlServer         As Object
    On Error GoTo ErrorHandler
    
    dbSQLServer.Execute strSQL, , adCmdText
    lngCount = 0
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    With rsSqlServer
'        .Open fncSelectTempTable(), dbSqlServer, adOpenForwardOnly, adLockPessimistic, adCmdText
        .Open WK_TABLE_NAME, dbSQLServer, adOpenForwardOnly, adLockPessimistic, adCmdTable
        On Error GoTo ErrorHandler1
        While (Not .EOF)
            wkKiro = KmFromDo(CInt(.Fields("YARD_ZAHYOKEI")), CDbl(.Fields("YARD_IDO")), CDbl(.Fields("YARD_KEIDO")), dblIdo, dblKeido)
            If wkKiro <= dblKiro Then
                lngCount = lngCount + 1
                .Fields("YARD_KIRO") = wkKiro
            Else
                .Delete
            End If
            .UPDATE
            .MoveNext
        Wend
        .Close
    End With
    On Error GoTo ErrorHandler
    fncUpdate_TEMP_KIRO = lngCount
Exit Function
    
ErrorHandler1:
    rsSqlServer.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncUpdate_TEMP_KIRO" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用テーブル作成ＳＱＬ作成
'       MODULE_ID       : fncCreateTempTable
'       CREATE_DATE     : 2006/02/13
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateTempTable() As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = strSQL & "CREATE TABLE " & WK_TABLE_NAME & " ("
    strSQL = strSQL & " YARD_CODE       NUMERIC(6)  NOT NULL,"
    strSQL = strSQL & " YARD_NAME       VARCHAR(36),"
    strSQL = strSQL & " YARD_ADDR_1     VARCHAR(36),"
    strSQL = strSQL & " YARD_ADDR_2     VARCHAR(36),"
    strSQL = strSQL & " YARD_ADDR_3     VARCHAR(36),"
    strSQL = strSQL & " YARD_ZAHYOKEI   NUMERIC(2),"
    strSQL = strSQL & " YARD_IDO        NUMERIC(8, 6),"
    strSQL = strSQL & " YARD_KEIDO      NUMERIC(9, 6),"
    strSQL = strSQL & " YARD_KIRO       NUMERIC(7, 4) "
    strSQL = strSQL & ")"
    
    fncCreateTempTable = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncCreateTempTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用データ作成ＳＱＬ作成
'       MODULE_ID       : fncInsertTempTable
'       CREATE_DATE     : 2006/02/13
'       PARAM           : strYcode              ヤードコード(I)
'                       : dblIdo                緯度(I)
'                       : dblKeido              経度(I)
'                       : dblKiro               含みたいキロメートル(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInsertTempTable(ByVal strYcode As String, ByVal dblIdo As Double, ByVal dblKeido As Double, ByVal dblKiro As Double) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = strSQL & "INSERT"
    strSQL = strSQL & " " & WK_TABLE_NAME & " "
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " YARD_CODE,"
    strSQL = strSQL & " YARD_NAME,"
    strSQL = strSQL & " YARD_ADDR_1,"
    strSQL = strSQL & " YARD_ADDR_2,"
    strSQL = strSQL & " YARD_ADDR_3,"
    strSQL = strSQL & " YARD_ZAHYOKEI,"
    strSQL = strSQL & " YARD_IDO,"
    strSQL = strSQL & " YARD_KEIDO,"
    strSQL = strSQL & " NULL AS YARD_KIRO "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " YARD_MAST "
    strSQL = strSQL & "WHERE YARD_IDO   BETWEEN " & Format(BetweenFromDo(dblIdo, dblKiro)) & " AND " & Format(BetweenToDo(dblIdo, dblKiro)) & " "
    strSQL = strSQL & "AND   YARD_KEIDO BETWEEN " & Format(BetweenFromDo(dblKeido, dblKiro)) & " AND " & Format(BetweenToDo(dblKeido, dblKiro)) & " "
    strSQL = strSQL & "AND   YARD_ZAHYOKEI > 0 "
    strSQL = strSQL & "AND   YARD_CODE  != '" & strYcode & "' "
    
    fncInsertTempTable = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncInsertTempTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用データ読込ＳＱＬ作成
'       MODULE_ID       : fncSelectTempTable
'       CREATE_DATE     : 2006/02/13
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectTempTable() As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " * "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " " & WK_TABLE_NAME & " "
    
    fncSelectTempTable = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncSelectTempTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 基準点からの範囲指定用の最小の度数を算出
'       MODULE_ID       : BetweenFromDo
'       CREATE_DATE     : 2006/02/13
'       PARAM           : d0                    度(I)
'                       : Km0                   含みたいキロメートル(I)
'       RETURN          : 度
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function BetweenFromDo(ByVal do0 As Double, ByVal km0 As Double) As Double
    On Error GoTo ErrorHandler
    BetweenFromDo = do0 - BetweenValueDo(do0, km0)
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "BetweenFromDo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 基準点からの範囲指定用の最大の度数を算出
'       MODULE_ID       : BetweenToDo
'       CREATE_DATE     : 2006/02/13
'       PARAM           : d0                    度(I)
'                       : Km0                   含みたいキロメートル(I)
'       RETURN          : 度
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function BetweenToDo(ByVal do0 As Double, ByVal km0 As Double) As Double
    On Error GoTo ErrorHandler
    BetweenToDo = do0 + BetweenValueDo(do0, km0)
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "BetweenToDo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 基準点からの範囲指定用の差異を算出
'       MODULE_ID       : BetweenValueDo
'       CREATE_DATE     : 2006/02/13
'       PARAM           : d0                    度(I)
'                       : Km0                   含みたいキロメートル(I)
'       RETURN          : 度
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function BetweenValueDo(ByVal do0 As Double, ByVal km0 As Double) As Double
    On Error GoTo ErrorHandler
    BetweenValueDo = 1# / IIf(do0 > 90, 90#, 110#) * km0
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "BetweenValueDo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ヤードの経緯度を抽出
'       MODULE_ID       : WebYardCode
'       CREATE_DATE     : 2006/02/13
'       PARAM           : strYardCode           ヤードコード(I)
'       RETURN          : 表示した(True)／表示できなかった(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function WebYardCode(ByVal strYardCode As String) As Boolean
    Dim dbSQLServer         As Object
    Dim rsSqlServer         As Object
    Dim IEApp               As Object
    On Error GoTo ErrorHandler
    
    'SQL-Server接続
    Set dbSQLServer = CreateObject("ADODB.Connection")
    dbSQLServer.CommandTimeout = 180
    Call dbSQLServer.Open(MSZZ007_M10(DLookup("CONT_BUMOC", "dbo_CONT_MAST")))
    On Error GoTo ErrorHandler1
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    With rsSqlServer
        .Open fncSelectYard2(strYardCode), dbSQLServer, adOpenForwardOnly, adLockReadOnly, adCmdText
        On Error GoTo ErrorHandler2
        If Not .EOF Then
            Call WebGoogleOpen(IEApp, False)
            Call WebGoogleMarker(IEApp, .Fields("YARD_IDO"), .Fields("YARD_KEIDO"), "r", Format(.Fields("YARD_CODE"), "000000"), .Fields("YARD_NAME"), Nz(.Fields("YARD_ADDR_1")) & Nz(.Fields("YARD_ADDR_2")) & Nz(.Fields("YARD_ADDR_3")), Nz(.Fields("YARD_NOTE")))
            WebYardCode = True
        Else
            WebYardCode = False
        End If
        .Close
    End With
    On Error GoTo ErrorHandler1
    dbSQLServer.Close
Exit Function
    
ErrorHandler2:
    rsSqlServer.Close
ErrorHandler1:
    dbSQLServer.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "WebYardCode" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ヤードの経緯度を抽出
'       MODULE_ID       : fncSelectYard2
'       CREATE_DATE     : 2006/02/13
'       PARAM           : strYardCode           ヤードコード(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectYard2(ByVal strYardCode As String) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " YARD_CODE,"
    strSQL = strSQL & " YARD_NAME,"
    strSQL = strSQL & " YARD_ADDR_1,"
    strSQL = strSQL & " YARD_ADDR_2,"
    strSQL = strSQL & " YARD_ADDR_3,"
    strSQL = strSQL & " YARD_NOTE,"
    strSQL = strSQL & " YARD_IDO,"
    strSQL = strSQL & " YARD_KEIDO "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " YARD_MAST "
    strSQL = strSQL & "WHERE YARD_CODE  = " & strYardCode & " "
    strSQL = strSQL & "AND   YARD_IDO   > 0 "
    strSQL = strSQL & "AND   YARD_KEIDO > 0 "

    fncSelectYard2 = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncSelectYard2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : テスト用
'       MODULE_ID       : TEST_WebYardCode
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub TEST_WebYardCode()
    If WebYardCode("000012") Then
        MsgBox "OK"
    Else
        MsgBox "NG"
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : テスト用
'       MODULE_ID       : TEST_TUKUBA_TOKYO
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub TEST_TUKUBA_TOKYO()
    Dim ans                 As Double
    Dim φ1                 As Double
    Dim λ1                 As Double
    Dim φ2                 As Double
    Dim λ2                 As Double
    Dim x1                  As Double
    Dim y1                  As Double
    Dim x2                  As Double
    Dim y2                  As Double
    Dim iCourse             As Integer
    Dim IEApp               As Object
    On Error GoTo ErrorHandler
    
    '東京都（XIV系、XVIII系及びXIX系に規定する区域を除く)　福島県　栃木県　茨城県　埼玉県 千葉県 群馬県　神奈川県
    iCourse = 9
    
    'つくば
    φ1 = Radian(36, 6, 2)
    λ1 = Radian(140, 5, 28)
    
    y1 = Yzahyou(iCourse, φ1, λ1)
    x1 = Xzahyou(iCourse, φ1, λ1)
    
    '東京
    φ2 = Radian(35, 39, 18)
    λ2 = Radian(139, 44, 41)
    
    y2 = Yzahyou(iCourse, φ2, λ2)
    x2 = Xzahyou(iCourse, φ2, λ2)
    
    Debug.Print KmFromZahyou(iCourse, y1, x1, y2, x2)
Exit Sub
    
ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("TEST_TUKUBA_TOKYO", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : テスト用
'       MODULE_ID       : TEST_AllResetNYAR_MAST
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub TEST_AllResetNYAR_MAST()
    Dim strBUMOC            As String
    On Error GoTo ErrorHandler
    
    strBUMOC = "H"
    If fncAllResetNYAR_MAST(strBUMOC) Then
        MsgBox "OK"
    Else
        MsgBox "NG"
    End If
Exit Sub
    
ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("TEST_AllResetNYAR_MAST", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 近隣マスタの整備
'       MODULE_ID       : fncAllResetNYAR_MAST
'       CREATE_DATE     : 2006/02/13
'       PARAM           : strBumoc              部門コード
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncAllResetNYAR_MAST(ByVal strBUMOC As String) As Boolean
    Dim strSQL              As String
    Dim dbSQLServer         As Object       'ADODB.Connection
    Dim dblKiro             As Double
    Dim StrYard             As String
    Dim dblIdo              As Double
    Dim dblKeido            As Double
    On Error GoTo ErrorHandler
    
    'SQL-Server接続
    Set dbSQLServer = CreateObject("ADODB.Connection")
    dbSQLServer.CommandTimeout = 180
    Call dbSQLServer.Open(MSZZ007_M10(strBUMOC))
    On Error GoTo ErrorHandler1
    'コントロールマスタの取得
    dblKiro = fncSelectCTRL_MAST(dbSQLServer)
    'SQL-Serverの一時テーブル作成
    strSQL = fncCreateTempTable()
    dbSQLServer.Execute strSQL, adCmdText
    'ヤードマスタの取得
    StrYard = ""
    While fncSelectTopYard(dbSQLServer, StrYard, dblIdo, dblKeido)
        strSQL = fncInsertTempTable(StrYard, dblIdo, dblKeido, dblKiro)
        If fncUpdate_TEMP_KIRO(dbSQLServer, strSQL, dblIdo, dblKeido, dblKiro) > 0 Then
            strSQL = fncUpdate_NYAR_KIRO(StrYard)
            dbSQLServer.Execute strSQL, , adCmdText
            strSQL = fncInsert_NYAR_KIRO(StrYard)
            dbSQLServer.Execute strSQL, , adCmdText
            strSQL = "TRUNCATE TABLE " & WK_TABLE_NAME
            dbSQLServer.Execute strSQL, , adCmdText
        End If
        strSQL = fncInsert_TempTable2(StrYard)
        If fncUpdate_TEMP_KIRO(dbSQLServer, strSQL, dblIdo, dblKeido, 9999) > 0 Then
            strSQL = fncUpdate_NYAR_KIRO(StrYard)
            dbSQLServer.Execute strSQL, , adCmdText
            strSQL = "TRUNCATE TABLE " & WK_TABLE_NAME
            dbSQLServer.Execute strSQL, , adCmdText
        End If
    Wend
    dbSQLServer.Close
    On Error GoTo ErrorHandler
    fncAllResetNYAR_MAST = True
Exit Function

ErrorHandler1:
    dbSQLServer.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncAllResetNYAR_MAST" & vbRightAllow & Err.Source, Err.Description & vbCrLf & "LastSQL=" & strSQL, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : コントロールマスタ読込
'       MODULE_ID       : fncSelectCTRL_MAST
'       CREATE_DATE     : 2006/02/13
'       PARAM           : dbSqlServer           データベース
'       RETURN          : 最大キロ数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectCTRL_MAST(dbSQLServer As Object) As Double
    Dim strSQL              As String
    Dim dblKiro             As Double
    Dim rsSqlServer         As Object       'ADODB.Recordset
    On Error GoTo ErrorHandler

    strSQL = "SELECT CONT_MAX_KIRO FROM CONT_MAST WHERE CONT_KEY = 1 "
    dblKiro = 0
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    With rsSqlServer
        .Open strSQL, dbSQLServer, adOpenForwardOnly, adLockReadOnly
        On Error GoTo ErrorHandler1
        If Not .EOF Then
            dblKiro = Nz(.Fields("CONT_MAX_KIRO"), 0)
        End If
        .Close
    End With
    On Error GoTo ErrorHandler
    If dblKiro = 0 Then
        Call MSZZ024_M10("ADODB.Recordset.Open", "CONT_MAX_KIROに１以上を設定してください")
    End If
    fncSelectCTRL_MAST = dblKiro
Exit Function
    
ErrorHandler1:
    rsSqlServer.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncSelectCTRL_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ヤードマスタ読込
'       MODULE_ID       : fncSelectTopYard
'       CREATE_DATE     : 2006/02/13
'       PARAM           : dbSqlServer           データベース
'                       : strYcode              ヤードコード(I/O)
'                       : dblIdo                緯度(O)
'                       : dblKeido              経度(O)
'       RETURN          : データあり(True)／なし(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectTopYard(dbSQLServer As Object, strYcode As String, dblIdo As Double, dblKeido As Double) As Boolean
    Dim strSQL              As String
    Dim rsSqlServer         As Object       'ADODB.Recordset
    On Error GoTo ErrorHandler
    
    strSQL = fncSelect_YARD_MAST(strYcode)
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    With rsSqlServer
        .Open strSQL, dbSQLServer, adOpenForwardOnly, adLockReadOnly
        On Error GoTo ErrorHandler1
        If Not .EOF Then
            strYcode = .Fields("YARD_CODE")
            dblIdo = .Fields("YARD_IDO")
            dblKeido = .Fields("YARD_KEIDO")
            fncSelectTopYard = True
        Else
            strYcode = ""
            dblIdo = 0
            dblKeido = 0
            fncSelectTopYard = False
        End If
        .Close
    End With
Exit Function
    
ErrorHandler1:
    rsSqlServer.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncSelectTopYard" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ヤードマスタ読込ＳＱＬ作成
'       MODULE_ID       : fncSelect_YARD_MAST
'       CREATE_DATE     : 2006/02/13
'       PARAM           : strYcode              ヤードコード(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelect_YARD_MAST(ByVal strYcode As String) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = strSQL & "SELECT TOP 1 "
    strSQL = strSQL & " YARD_CODE, "
    strSQL = strSQL & " YARD_ZAHYOKEI, "
    strSQL = strSQL & " YARD_IDO, "
    strSQL = strSQL & " YARD_KEIDO "
    strSQL = strSQL & "FROM  YARD_MAST "
    strSQL = strSQL & "WHERE YARD_ZAHYOKEI > 0 "
    strSQL = strSQL & "AND   YARD_IDO      > 0 "
    strSQL = strSQL & "AND   YARD_KEIDO    > 0 "
    If strYcode <> "" Then
        strSQL = strSQL & "AND   YARD_CODE     > " & strYcode & " "
    End If
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & " YARD_CODE "

    fncSelect_YARD_MAST = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncSelect_YARD_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 近隣マスタ更新ＳＱＬ作成
'       MODULE_ID       : fncUpdate_NYAR_KIRO
'       CREATE_DATE     : 2006/02/13
'       PARAM           : strYcode              ヤードコード(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdate_NYAR_KIRO(ByVal strYcode As String) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = strSQL & "UPDATE"
    strSQL = strSQL & " NYAR_MAST "
    strSQL = strSQL & "SET"
    strSQL = strSQL & " NYAR_UPDAD = '" & Format(DATE, "yyyymmdd") & "', "
    strSQL = strSQL & " NYAR_UPDAJ = '" & Format(time, "hhmmss") & "', "
    strSQL = strSQL & " NYAR_UPDPB = '" & PROG_ID & "', "
    strSQL = strSQL & " NYAR_UPDUB = '" & Left(MSZZ000.LsGetUserName(), 8) & "', "
    strSQL = strSQL & " NYAR_KIRO  = YARD_KIRO "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " " & WK_TABLE_NAME & " "
    strSQL = strSQL & "WHERE NYAR_YCODE = '" & strYcode & "' "
    strSQL = strSQL & "AND   NYAR_NCODE = YARD_CODE "
    
    fncUpdate_NYAR_KIRO = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncUpdate_NYAR_KIRO" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 近隣マスタ作成ＳＱＬ作成
'       MODULE_ID       : fncInsert_NYAR_KIRO
'       CREATE_DATE     : 2006/02/13
'       PARAM           : strYcode              ヤードコード(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInsert_NYAR_KIRO(ByVal strYcode As String) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = strSQL & "INSERT"
    strSQL = strSQL & " NYAR_MAST "
    strSQL = strSQL & "( "
    strSQL = strSQL & " NYAR_YCODE, "
    strSQL = strSQL & " NYAR_NCODE, "
    strSQL = strSQL & " NYAR_INSED, "
    strSQL = strSQL & " NYAR_INSEJ, "
    strSQL = strSQL & " NYAR_INSPB, "
    strSQL = strSQL & " NYAR_INSUB, "
    strSQL = strSQL & " NYAR_KIRO "
    strSQL = strSQL & ") "
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " '" & strYcode & "',"
    strSQL = strSQL & " YARD_CODE,"
    strSQL = strSQL & " '" & Format(DATE, "yyyymmdd") & "', "
    strSQL = strSQL & " '" & Format(time, "hhmmss") & "', "
    strSQL = strSQL & " '" & PROG_ID & "', "
    strSQL = strSQL & " '" & Left(MSZZ000.LsGetUserName(), 8) & "', "
    strSQL = strSQL & " YARD_KIRO "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " " & WK_TABLE_NAME & " "
    strSQL = strSQL & "WHERE NOT EXISTS"
    strSQL = strSQL & "    ("
    strSQL = strSQL & "    SELECT"
    strSQL = strSQL & "        *"
    strSQL = strSQL & "    FROM"
    strSQL = strSQL & "        NYAR_MAST"
    strSQL = strSQL & "    WHERE"
    strSQL = strSQL & "        NYAR_YCODE = '" & strYcode & "'"
    strSQL = strSQL & "    AND NYAR_NCODE = YARD_CODE"
    strSQL = strSQL & "    )"
    
    fncInsert_NYAR_KIRO = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncInsert_NYAR_KIRO" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用データ作成ＳＱＬ作成
'       MODULE_ID       : fncInsert_TempTable2
'       CREATE_DATE     : 2006/02/13
'       PARAM           : strYcode              ヤードコード(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInsert_TempTable2(ByVal strYcode As String) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = strSQL & "INSERT"
    strSQL = strSQL & " " & WK_TABLE_NAME & " "
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " YARD_CODE,"
    strSQL = strSQL & " YARD_NAME,"
    strSQL = strSQL & " YARD_ADDR_1,"
    strSQL = strSQL & " YARD_ADDR_2,"
    strSQL = strSQL & " YARD_ADDR_3,"
    strSQL = strSQL & " YARD_ZAHYOKEI,"
    strSQL = strSQL & " YARD_IDO,"
    strSQL = strSQL & " YARD_KEIDO,"
    strSQL = strSQL & " NULL AS YARD_KIRO "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " YARD_MAST "
    strSQL = strSQL & "INNER JOIN "
    strSQL = strSQL & " NYAR_MAST "
    strSQL = strSQL & "ON( YARD_CODE     = NYAR_NCODE "
    strSQL = strSQL & ") "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "    NYAR_YCODE    = '" & strYcode & "' "
    strSQL = strSQL & "AND NYAR_NCODE   != '" & strYcode & "' "
    strSQL = strSQL & "AND NYAR_KIRO    IS NULL "
    strSQL = strSQL & "AND YARD_IDO      > 0 "
    strSQL = strSQL & "AND YARD_KEIDO    > 0 "
    strSQL = strSQL & "AND YARD_ZAHYOKEI > 0 "

    fncInsert_TempTable2 = strSQL
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncInsert_TempTable2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 地図用データベースのオープン
'       MODULE_ID       : OpenMapDB
'       CREATE_DATE     : 2006/02/14
'       RETURN          : データベースオブジェクト
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function OpenMapDB() As Object
    Dim dbSQLServer         As Object       'ADODB.Connection
    Dim strConn             As String
    Dim i                   As Integer
    On Error GoTo ErrorHandler

    strConn = MSZZ007_M10("H")
    i = InStr(1, strConn, "DATABASE=")
    If i <= 0 Then
        Set OpenMapDB = Nothing
        Exit Function
    End If
    strConn = Left(strConn, i + 8) & "MAP_DB"
    Set dbSQLServer = CreateObject("ADODB.Connection")
    dbSQLServer.CommandTimeout = 180
    Call dbSQLServer.Open(strConn)
    Set OpenMapDB = dbSQLServer
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "OpenMapDB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'****************************  ended of program ********************************
