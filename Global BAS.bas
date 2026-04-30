Attribute VB_Name = "Global BAS"
Option Compare Database
Option Explicit
Public Function FullFix(ByVal MyNumber As Double) As Double
'標準の Fix 関数にバグがあるので、
'それを回避する関数。
'Fix の機能を完全（Full）に満たしているので、
'FullFix という名前にした。
'使用方法は通常の Fix 関数と同様。

    Dim StrLen  As Long
    Dim TempVal As Double

    StrLen = Len(Format(MyNumber, "0.0"))
    TempVal = Val(Left(Trim(Str(MyNumber)), (StrLen - 2)))
    FullFix = TempVal

End Function



'UNLHA32.DLLのﾘﾀｰﾝｺｰﾄﾞからどういうｴﾗｰかをMsgBoxで表示する
Public Sub hrpUnlhaErrorMessage(lngErrCode As Long)
    Dim strMsg  As String
    Dim intType As Integer
    strMsg = "#:" & CStr(lngErrCode) & vbCr
    Select Case lngErrCode
    '■警告: 該当ﾌｧｲﾙについての処理をｽｷｯﾌﾟするだけで実行を中止する事はない。
        Case ERROR_DISK_SPACE
            strMsg = strMsg & "解凍する為のﾃﾞｨｽｸの空きが足りません。"
        Case ERROR_READ_ONLY
            strMsg = strMsg & "解凍先のﾌｧｲﾙはﾘｰﾄﾞｵﾝﾘｰです。"
        Case ERROR_USER_SKIP
            strMsg = strMsg & "ﾕｰｻﾞｰによって解凍をｽｷｯﾌﾟされました。"
        Case ERROR_UNKNOWN_TYPE
            strMsg = strMsg & "未知の圧縮ﾀｲﾌﾟです。"
        Case ERROR_METHOD
            strMsg = strMsg & "メソットのｴﾗｰです。"
        Case ERROR_PASSWORD_FILE
            strMsg = strMsg & "ﾊﾟｽﾜｰﾄﾞの入ったﾌｧｲﾙです。"
        Case ERROR_VERSION
            strMsg = strMsg & "ﾊﾞｰｼﾞｮﾝが違います。"
        Case ERROR_FILE_CRC
            strMsg = strMsg & "格納ﾌｧｲﾙのﾁｪｯｸｻﾑが合っていません。"
        Case ERROR_FILE_OPEN
            strMsg = strMsg & "解凍時にﾌｧｲﾙを開けませんでした。"
        Case ERROR_MORE_FRESH
            strMsg = strMsg & "より新しいﾌｧｲﾙが解凍先に存在しています。"
        Case ERROR_NOT_EXIST
            strMsg = strMsg & "ﾌｧｲﾙは解凍先に存在していません。"
        Case ERROR_ALREADY_EXIST
            strMsg = strMsg & "既にﾌｧｲﾙが存在します。"
        Case ERROR_TOO_MANY_FILES
            strMsg = strMsg & "ﾌｧｲﾙが多すぎます。"
    '■ｴﾗｰ: 致命的なｴﾗｰでその時点で実行を中止する。
        Case ERROR_MAKEDIRECTORY
            strMsg = strMsg & "ﾌｫﾙﾀﾞが作成できません。"
        Case ERROR_CANNOT_WRITE
            strMsg = strMsg & "解凍中に書き込みｴﾗｰが生じました。"
        Case ERROR_HUFFMAN_CODE
            strMsg = strMsg & "LZH ﾌｧｲﾙのﾊﾌﾏﾝｺｰﾄﾞが壊れています。"
        Case ERROR_COMMENT_HEADER
            strMsg = strMsg & "LZH ﾌｧｲﾙのｺﾒﾝﾄﾍｯﾀﾞが壊れています。"
        Case ERROR_HEADER_CRC
            strMsg = strMsg & "LZH ﾌｧｲﾙのﾍｯﾀﾞのﾁｪｯｸｻﾑが合っていません。"
        Case ERROR_HEADER_BROKEN
            strMsg = strMsg & "LZH ﾌｧｲﾙのﾍｯﾀﾞが壊れています。"
        Case ERROR_ARC_FILE_OPEN
            strMsg = strMsg & "LZH ﾌｧｲﾙを開く事が出来ません。"
        Case ERROR_NOT_ARC_FILE
            strMsg = strMsg & "LZH ﾌｧｲﾙ名の指定がされていません。"
        Case ERROR_CANNOT_READ
            strMsg = strMsg & "LZH ﾌｧｲﾙの読み込み時に読み込みｴﾗｰが出ました。"
        Case ERROR_FILE_STYLE
            strMsg = strMsg & "指定されたﾌｧｲﾙは LZH ﾌｧｲﾙではありません。"
        Case ERROR_COMMAND_NAME
            strMsg = strMsg & "ｺﾏﾝﾄﾞ指定が間違っています。"
        Case ERROR_MORE_HEAP_MEMORY
            strMsg = strMsg & "作業用のためのﾋｰﾌﾟﾒﾓﾘが不足しています。"
        Case ERROR_ENOUGH_MEMORY
            strMsg = strMsg & "ｸﾞﾛｰﾊﾞﾙﾒﾓﾘが不足しています。"
        Case ERROR_ALREADY_RUNNING
            strMsg = strMsg & "既に UNLHA32.DLL が動作中です。"
        Case ERROR_USER_CANCEL
            strMsg = strMsg & "ﾕｰｻﾞｰによって解凍を中断されました。"
        Case ERROR_HARC_ISNOT_OPENED
            strMsg = strMsg & "UnlhaOpenArchive() で書庫ﾌｧｲﾙとﾊﾝﾄﾞﾙを関連付ける前に UnlhaFindFirst() 等の API を使用した。"
        Case ERROR_NOT_SEARCH_MODE
            strMsg = strMsg & "UnlhaFindFirst() を使用する前に UnlhaFindNext() が呼ばれた。または，これらの API を呼び出す前に UnlhaGetFileName() 等の API が呼ばれた。"
        Case ERROR_NOT_SUPPORT
            strMsg = strMsg & "UNLHA32.DLL でｻﾎﾟｰﾄされていない API が使用されました。"
        Case ERROR_TIME_STAMP
            strMsg = strMsg & "日付及び時間の指定形式が間違っています。"
        Case ERROR_TMP_OPEN
            strMsg = strMsg & "作業ﾌｧｲﾙがｵｰﾌﾟﾝできません。"
        Case ERROR_LONG_FILE_NAME
            strMsg = strMsg & "ﾌｫﾙﾀﾞのﾊﾟｽが長すぎます。"
        Case ERROR_ARC_READ_ONLY
            strMsg = strMsg & "書き込み専用属性の書庫ﾌｧｲﾙに対する操作はできません。"
        Case ERROR_SAME_NAME_FILE
            strMsg = strMsg & "すでに同じ名前のﾌｧｲﾙが書庫に格納されています。"
        Case ERROR_NOT_FIND_ARC_FILE
            strMsg = strMsg & "指定されたﾌｫﾙﾀﾞには LZH ﾌｧｲﾙがありませんでした。"
        Case Else
            strMsg = strMsg & "未知のｴﾗｰです。"
    End Select
    intType = vbOKOnly Or vbCritical Or vbApplicationModal
    Call MsgBox(strMsg, intType, "在庫管理システム")

End Sub

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/  引数で指定したﾌｧｲﾙを圧縮する。                                                              _/_/_/_/_/_/
'_/_/_/_/_/  stLzhName :圧縮後のﾌｧｲﾙ名(ﾌﾙﾊﾟｽで指定)                                                     _/_/_/_/_/_/
'_/_/_/_/_/  stFileName:圧縮するﾌｧｲﾙ名(ﾌﾙﾊﾟｽで指定)                                                     _/_/_/_/_/_/
'_/_/_/_/_/  lohwnd    :ﾌｫｰﾑのﾊﾝﾄﾞﾙ                                                                    _/_/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Function Freezing(stLzhName As String, stFileName As String, lohwnd As Long)

  Dim lloRtn        As Long
  Dim BUFF          As String * 2048
  Dim lstCmdLine    As String
  Dim lstVer As String
  
  lstVer = LHA_UnlhaGetVersion()

'/****** UnLha32.dll が実行されているかのチェック
  If LHA_UnlhaGetRunning() <> 0 Then
    MsgBox "他のアプリケーションにより圧縮／解凍が実行されています!!" + vbCrLf + "他のアプリケーションが終了してから実行して下さい!!", 64, "解凍"
    Freezing = 1
    Exit Function
  End If

'/****** Unlha32のﾊﾞｯｸｸﾞﾗｳﾝﾄﾞﾓｰﾄﾞのﾁｪｯｸ
'/****** ﾊﾞｯｸｸﾞﾗｳﾝﾄﾞﾓｰﾄﾞであれば非ﾊﾞｯｸｸﾞﾗｳﾝﾄﾞﾓｰﾄﾞに変更
  If LHA_UnlhaGetBackGroundMode() <> 0 Then
    lloRtn = LHA_UnlhaSetBackGroundMode(False)
  End If

  lstCmdLine = stLzhName + " " + stFileName

  BUFF = ""
'/****** -y ｽｲｯﾁにより確認ﾒｯｾｰｼﾞ全てにYesとする。
  'lloRtn = LHA_Unlha(Me.hwnd, "a -y " + lstCmdLine, BUFF, 2048)
  lloRtn = LHA_Unlha(lohwnd, "a " + lstCmdLine, BUFF, 2048)
  If lloRtn <> 0& Then 'エラーの場合
      Call hrpUnlhaErrorMessage(lloRtn)
  End If

  Freezing = lloRtn

End Function
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/  引数で指定したﾌｧｲﾙを解凍する。                                                              _/_/_/_/_/_/
'_/_/_/_/_/  stLzhName :解凍するﾌｧｲﾙ名(ﾌﾙﾊﾟｽで指定)                                                     _/_/_/_/_/_/
'_/_/_/_/_/  stFileName:解凍するﾃﾞｨﾚｸﾄﾘ                                                                _/_/_/_/_/_/
'_/_/_/_/_/  lohwnd    :ﾌｫｰﾑのﾊﾝﾄﾞﾙ                                                                    _/_/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Function Melting(stLzhName As String, stDir As String, lohwnd As Long)

  Dim lloRtn        As Long
  Dim BUFF          As String * 2048    '/****** Unlha32 の返り値を格納する変数
  Dim lstCmdLine    As String           '/****** Unlha32 に渡すｺﾏﾝﾄﾞを格納する変数
  Dim lstVer        As String           '/****** Unlha32 のﾊﾞｰｼﾞｮﾝを格納する変数

  lstVer = LHA_UnlhaGetVersion()

'/****** UnLha32.dll が実行されているかのチェック
  If LHA_UnlhaGetRunning() <> 0 Then
    MsgBox "他のアプリケーションにより圧縮／解凍が実行されています!!" + vbCrLf + "他のアプリケーションが終了してから実行して下さい!!", 64, "圧縮・解凍"
    Melting = 1&
    Exit Function
  End If

'/****** Unlha32のﾊﾞｯｸｸﾞﾗｳﾝﾄﾞﾓｰﾄﾞのﾁｪｯｸ
'/****** ﾊﾞｯｸｸﾞﾗｳﾝﾄﾞﾓｰﾄﾞであれば非ﾊﾞｯｸｸﾞﾗｳﾝﾄﾞﾓｰﾄﾞに変更
  If LHA_UnlhaGetBackGroundMode() <> 0 Then
    lloRtn = LHA_UnlhaSetBackGroundMode(False)
  End If

'/****** unlha32の Commndの作成
  If CmRightB(stDir, 1) <> "\" Then
    lstCmdLine = stLzhName + " " + stDir + "\"
  Else
    lstCmdLine = stLzhName + " " + stDir
  End If

'/****** 返り値の格納変数の初期化
  BUFF = ""

'/****** ﾌｧｲﾙの解凍
  lloRtn = LHA_Unlha(lohwnd, "x -m1y0 " + lstCmdLine, BUFF, 2048)
  If lloRtn <> 0& Then 'エラーの場合
      Call hrpUnlhaErrorMessage(lloRtn)
  End If

  Melting = lloRtn

End Function

Private Sub CompactDB(lstDataFile As String, lstBackUpFile As String)

'/****** 圧縮先ﾌｧｲﾙの存在ﾁｪｯｸ
'/****** 若し、ﾌｧｲﾙが存在すればﾌｧｲﾙを削除
'/****** (上書きできへんのかねー、しっかし、上書きしてくれたらこんなｺｰﾄﾞ書かんでええのになぁ(;_;))
  If Dir(lstBackUpFile) <> "" And Not IsNull(Dir(lstBackUpFile)) Then
    Kill lstBackUpFile
  End If

'/****** ﾃﾞｰﾀﾍﾞｰｽの圧縮
  DBEngine.CompactDatabase lstDataFile, lstBackUpFile

End Sub
Public Function CmNumCheck(LST_IN_DATA As String, LIN_KETA As Integer, LIN_SKETA As Integer) As Integer
'---------------------------------------------
' 数値ﾁｪｯｸ関数
'---------------------------------------------
    '-----------------------------------------------------------------------------------------------------------------------------
    '　　パラメータを基に数値、桁数のチェックを行い結果を返す
    '
    'LST_IN_DATAは入力値設定
    'LIN_KETAは整数桁数
    'LIN_SKETAは小数点以下桁数
    'CmNumCheckは "-1"は　数値（正数）
    '               "0"は　ZERO
    '               "1"は　数値（負数）
    '               "2"は　数値以外
    '               "3"は  桁あふれ
    '               "9"は  パラメータエラー
    '-----------------------------------------------------------------------------------------------------------------------------
'On Error GoTo CmNumCheckERR

    Dim LIN_TEN     As Integer       '小数点の位置
    Dim LIN_ALL_LEN As Integer       'ＩＮ＿ＤＡＴＡの長さ
    Dim LIN_SU_LEN  As Integer       '小数部の長さ
    Dim LIN_SE_LEN  As Integer       '整数部の長さ
    Dim LST_FLG     As String * 1    'ＺＥＲＯサプレス，ＳＰＡＣＥサプレス用
    Dim LCU_SUJI    As Currency      '桁あふれ判定用
    Dim LST_SUJI    As String
    Dim LST_FUGOU   As String * 1
    Dim i, N        As Integer

    '初期値設定
    CmNumCheck = -1
    'パラメータチェク
    If (LIN_KETA < 1) Then
        CmNumCheck = 9
        Exit Function
    End If
    If (LIN_SKETA < 0) Then
        CmNumCheck = 9
        Exit Function
    End If
    If (LIN_KETA + LIN_SKETA) > 13 Then
        CmNumCheck = 9
        Exit Function
    End If
    'ＮＵＬＬチェク
    If LST_IN_DATA = "" Then
        CmNumCheck = 2
       Exit Function
    End If
    '空白検知
    If InStr(LST_IN_DATA, " ") <> 0 Then
        CmNumCheck = 2
        Exit Function
    End If
    '数値妥当性チェク
    If (IsNumeric(LST_IN_DATA) = False) Then
        CmNumCheck = 2
        Exit Function
    End If
    For N = 1 To i
        If IsNumeric(Mid$(LST_IN_DATA, N, 1)) = False Then
            CmNumCheck = 2
            Exit Function
        End If
    Next

    'チェク桁数算出
    i = 1
    Do Until Mid$(LST_IN_DATA, i, 1) = ""
       i = i + 1
    Loop
    If i = 1 Then
        CmNumCheck = 2
        Exit Function
    Else
        i = i - 1
    End If
    For N = 1 To i
        Select Case Mid$(LST_IN_DATA, N, 1)
            Case "+", "-", ".", " ", ""
            Case "0" To "9"
            Case Else
                CmNumCheck = 2
                Exit Function
        End Select
    Next
    '数字間のスペースチェク
    LST_FLG = ""
    For N = 1 To i
        If Mid$(LST_IN_DATA, N, 1) = " " Then
            If LST_FLG = "1" Then
               LST_FLG = "2"
            End If
        Else
            If LST_FLG = "2" Then
               CmNumCheck = 3
               Exit Function
            Else
               LST_FLG = "1"
            End If
        End If
    Next
    '符号判定
    For N = 1 To i
        If Mid$(LST_IN_DATA, N, 1) = "+" Then
           LST_FUGOU = "+"
        End If
        If Mid$(LST_IN_DATA, N, 1) = "-" Then
           LST_FUGOU = "-"
        End If
    Next
    'SPACE、","、"+"、"-"、"0"サプレス
    LST_FLG = ""
    For N = 1 To i
        If Mid$(LST_IN_DATA, N, 1) = Space$(1) Or Mid$(LST_IN_DATA, N, 1) = "," Or Mid$(LST_IN_DATA, N, 1) = "+" Or Mid$(LST_IN_DATA, N, 1) = "-" Then
        Else
            If Mid$(LST_IN_DATA, N, 1) = "0" Then
               If LST_FLG = "1" Then
                    LST_SUJI = LST_SUJI & Mid$(LST_IN_DATA, N, 1)
               End If
            Else
                LST_FLG = "1"
                LST_SUJI = LST_SUJI & Mid$(LST_IN_DATA, N, 1)
            End If
        End If
    Next
    '小数点の位置を算出
    LIN_TEN = InStr(LST_SUJI, ".")
    '小数点の有無を判定
    If (LIN_TEN <> 0) And (LIN_SKETA = 0) Then
       CmNumCheck = 3
       Exit Function
    End If
    '入力データの長さを算出
    LIN_ALL_LEN = Len(LST_SUJI)
    '入力データの小数部の長さを算出
    If LIN_TEN = 0 Then
        LIN_SU_LEN = 0
    Else
        LIN_SU_LEN = LIN_ALL_LEN - LIN_TEN
    End If
    '入力データの整数部の長さを算出
    If LIN_TEN = 0 Then
       LIN_SE_LEN = LIN_ALL_LEN
    Else
        LIN_SE_LEN = LIN_ALL_LEN - 1 - LIN_SU_LEN
    End If
    '整数部の桁数チェク
    If LIN_KETA < LIN_SE_LEN Then
        CmNumCheck = 3
        Exit Function
    End If
    '小数部の桁数チェク
    If LIN_SKETA < LIN_SU_LEN Then
        CmNumCheck = 3
        Exit Function
    End If
    '符号を判定しリターンコドを設定
    If LST_FUGOU = "-" Then
        LST_SUJI = "-" & LST_SUJI
        CmNumCheck = 1
    End If
    'ＺＥＲＯチェク
    LCU_SUJI = Val(LST_SUJI)
    If LCU_SUJI = 0 Then
        LST_SUJI = "0"
        CmNumCheck = 0
    End If

'    Exit Function
'CmNumCheckERR:
'    Call CmDicsAp_Abend("CmNumCheck でエラーが発生しました。異常終了します。", 1, 1)


End Function
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/ Unicode対応のRightBの代用                                         _/_/_/_/_/
'_/_/_/_/ stArg：この文字列式の左端から文字列が取り出されます。                _/_/_/_/_/
'_/_/_/_/ lolen：取り出す文字列の文字数                                      _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Function CmRightB(ByVal stArg As String, ByVal lolen As Long) As String
  CmRightB = CmStrConv(RightB(CmStrConv(stArg, vbFromUnicode), lolen), vbUnicode)
End Function
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/ Unicode対応のLeftBの代用                                          _/_/_/_/_/
'_/_/_/_/ stArg：この文字列式の左端から文字列が取り出されます。               _/_/_/_/_/
'_/_/_/_/ lolen：取り出す文字列の文字数                                     _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Function CmLeftB(ByVal stArg As String, ByVal lolen As Long) As String

  CmLeftB = CmStrConv(LeftB(CmStrConv(stArg, vbFromUnicode), lolen), vbUnicode)
  
End Function
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/ Unicode対応のLenBの代用                                           _/_/_/_/_/
'_/_/_/_/ stArg：任意の文字列式を指定します。                                _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Function CmLenb(ByVal stArg As Variant) As Long
    
    CmLenb = LenB(CmStrConv(stArg, vbFromUnicode))

End Function
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/ Unicode対応のMidBの代用                                           _/_/_/_/_/
'_/_/_/_/ stArg  ：任意の文字列式を指定します。                              _/_/_/_/_/
'_/_/_/_/ loStart：stArgのどの位置から文字列を取出すかを指定します。          _/_/_/_/_/
'_/_/_/_/ vaEnd  ：stArgのどの位置まで文字列を取出すかを指定します。省略可    _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Function CmMidB(ByVal stArg As String, ByVal loStart As Long, Optional vaEnd)

    If IsMissing(vaEnd) Then
      CmMidB = CmStrConv(MidB(CmStrConv(stArg, vbFromUnicode), loStart), vbUnicode)
    Else
      CmMidB = CmStrConv(MidB(CmStrConv(stArg, vbFromUnicode), loStart, vaEnd), vbUnicode)
    End If

End Function
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/_/                                                                              _/_/_/_/_/_/
'_/_/_/_/_/_/    入力した日付を年月で返す                                                    _/_/_/_/_/_/
'_/_/_/_/_/_/                                                                              _/_/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Function Cm_DateConvYM(ByVal STDATE As String) As String

  Dim stDateValue   As String     '/****** 返り値を格納する変数
  Dim inByte        As Integer    '/****** 一番最初の／が見つかった位置を格納する変数

  If STDATE = "" Or IsNull(STDATE) Then Exit Function
  If CmLenb(STDATE) > 7 Then Exit Function
  If CmRightB(STDATE, 1) = "/" Then Exit Function
  If CmLeftB(STDATE, 1) = "/" Then Exit Function

  inByte = InStr(STDATE, "/")

'/******* /の位置による年月の判定
  Select Case inByte
    Case 1
      stDateValue = ""
    Case 2:
      If CmLenb(CmMidB(STDATE, 3)) > 2 Or CmLenb(CmMidB(STDATE, 3)) = 0 Then
        stDateValue = ""
      ElseIf CmMidB(STDATE, 3) >= 1 And CmMidB(STDATE, 3) <= 12 Then
        stDateValue = "200" + CStr(STDATE)
      End If
    Case 3:
      If CmLenb(CmMidB(STDATE, 4)) > 2 Or CmLenb(CmMidB(STDATE, 3)) = 0 Then
        stDateValue = ""
      ElseIf CmMidB(STDATE, 4) >= 1 And CmMidB(STDATE, 4) <= 12 Then
        If CInt(CmMidB(STDATE, 1, 2)) < 80 Then
          stDateValue = "20" + STDATE
        Else
          stDateValue = "19" + STDATE
        End If
      End If
    Case 4
      If CmLenb(CmMidB(STDATE, 5)) > 2 Or CmLenb(CmMidB(STDATE, 5)) = 0 Then
        stDateValue = ""
      ElseIf CmMidB(STDATE, 5) >= 1 And CmMidB(STDATE, 5) <= 12 Then
        If CInt(CmMidB(STDATE, 1, 3)) < 900 Then
          stDateValue = "2" + STDATE
        Else
          stDateValue = "1" + STDATE
        End If
      End If
    Case 5
      If CmLenb(CmMidB(STDATE, 6)) > 2 Or CmLenb(CmMidB(STDATE, 6)) = 0 Then
        stDateValue = ""
      ElseIf CmMidB(STDATE, 6) >= 1 And CmMidB(STDATE, 6) <= 12 Then
        stDateValue = STDATE
      End If
    Case 0
      Select Case CmLenb(STDATE)
        Case 0:
          Exit Function
        Case 1:
          stDateValue = CStr(Year(DATE)) + "/" + CStr(STDATE)
        Case 2:
          If STDATE > 12 Or CInt(CmRightB(STDATE, 2)) < 1 Then
            If CmRightB(STDATE, 1) >= 1 And CmRightB(STDATE, 1) <= 9 Then
              If CmMidB(STDATE, 1, 1) = CmMidB(Format$(DATE, "yyyy"), 4, 1) Then
                stDateValue = Format$(DATE, "yyyy") + "/" + CmMidB(STDATE, 2, 1)
              Else
                stDateValue = "200" + CmMidB(STDATE, 1, 1) + "/" + CmMidB(STDATE, 2, 1)
              End If
            Else
              stDateValue = STDATE + "/" + CStr(Month(DATE))
            End If
          ElseIf STDATE >= 1 And STDATE <= 12 Then
            stDateValue = CStr(Year(DATE)) + "/" + STDATE
          End If
        Case 3:
          If CInt(CmRightB(STDATE, 2)) > 12 Or CInt(CmRightB(STDATE, 2)) < 1 Then
            If CmRightB(STDATE, 1) >= 1 And CmRightB(STDATE, 1) <= 9 Then
              If CInt(CmMidB(STDATE, 1, 2)) < 90 Then
                stDateValue = "20" + CmMidB(STDATE, 1, 2) + "/" + CmRightB(STDATE, 1)
              Else
                stDateValue = "19" + CmMidB(STDATE, 1, 2) + "/" + CmRightB(STDATE, 1)
              End If
            Else
              stDateValue = STDATE + "/" + CStr(Month(DATE))
            End If
          ElseIf CInt(CmRightB(STDATE, 2)) >= 1 And CInt(CmRightB(STDATE, 2)) <= 12 Then
            stDateValue = "200" + CStr(CmMidB(STDATE, 1, 1)) + "/" + CmMidB(STDATE, 2, 2)
          End If
        Case 4:
          If CInt(CmRightB(STDATE, 2)) > 12 Or CInt(CmRightB(STDATE, 2)) < 1 Then
            If CmRightB(STDATE, 1) >= 1 And CmRightB(STDATE, 1) <= 9 Then
              If CmMidB(STDATE, 1, 3) < 900 Then
                stDateValue = "2" + CmMidB(STDATE, 1, 3) + "/" + CmRightB(STDATE, 1)
              Else
                stDateValue = "1" + CmMidB(STDATE, 1, 3) + "/" + CmRightB(STDATE, 1)
              End If
            Else
              stDateValue = ""
            End If
          ElseIf CInt(CmRightB(STDATE, 2)) >= 1 Or CInt(CmRightB(STDATE, 2)) <= 12 Then
            If CInt(CmMidB(STDATE, 1, 2)) < 80 Then
              stDateValue = "20" + CStr(CmMidB(STDATE, 1, 2)) + "/" + CmMidB(STDATE, 3, 2)
            Else
              stDateValue = "19" + CStr(CmMidB(STDATE, 1, 2)) + "/" + CmMidB(STDATE, 3, 2)
            End If
          End If
        Case 5:
          If CInt(CmRightB(STDATE, 2)) > 12 Or CInt(CmRightB(STDATE, 2)) < 1 Then
            If CInt(CmRightB(STDATE, 1)) >= 1 And CInt(CmRightB(STDATE, 1)) <= 9 Then
              stDateValue = CmMidB(STDATE, 1, 4) + "/" + CmMidB(STDATE, 5, 1)
            Else
              Exit Function
            End If
          ElseIf CInt(CmRightB(STDATE, 2)) >= 1 Or CInt(CmRightB(STDATE, 2)) <= 12 Then
            If CInt(CmMidB(STDATE, 1, 3)) > 900 Then
              stDateValue = "1" + CmMidB(STDATE, 1, 3) + "/" + CmRightB(STDATE, 2)
            Else
              stDateValue = "2" + CmMidB(STDATE, 1, 3) + "/" + CmRightB(STDATE, 2)
            End If
          End If
        Case 6:
          If CInt(CmRightB(STDATE, 2)) > 12 Or CInt(CmRightB(STDATE, 2)) < 1 Then
            Exit Function
          ElseIf CInt(CmRightB(STDATE, 2)) >= 1 Or CInt(CmRightB(STDATE, 2)) <= 12 Then
            stDateValue = CmMidB(STDATE, 1, 4) + "/" + CmRightB(STDATE, 2)
          Else
            Exit Function
          End If
        Case Else
          Exit Function
      End Select
  End Select

  Cm_DateConvYM = Format$(stDateValue, "yyyy/mm")

End Function
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/ IniFileからキー値の取得                                           _/_/_/_/_/
'_/_/_/_/ stSection：iniFile のセクション                                   _/_/_/_/_/
'_/_/_/_/ stKey    ：Keyの項目                                             _/_/_/_/_/
'_/_/_/_/ stIniFile：iniFileの名前                                         _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Function GetIniString(stSection, stKey, stIniFile)
  Dim stTmp     As String * 256
  Dim stBuf    As String
  Dim inRet    As Integer

  inRet = GetPrivateProfileString(Format$(stSection), Format$(stKey), "", stTmp, 256, stIniFile)
  stBuf = CmLeftB(stTmp, inRet)

  GetIniString = stBuf
End Function
Public Function CmStrConv(st1, flag)

  CmStrConv = StrConv(st1, flag)

End Function
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/ 入力された文字列をｷｬﾝｾﾙする                                        _/_/_/_/_/
'_/_/_/_/ stArg     ：ｷｬﾝｾﾙしたい文字列                                     _/_/_/_/_/
'_/_/_/_/ inKeyAscii :KeyAscii                                             _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Sub CmStringCancel(stArg As String, inKeyAscii As Integer)

  Dim inLpCnt As Integer

    For inLpCnt = 1 To Len(stArg)
        If Asc(Mid(stArg, inLpCnt, 1)) = inKeyAscii Then
            inKeyAscii = 0
        End If
    Next

End Sub

Public Function IsLoaded(MyFormName) As Boolean
' 引数: フォーム名
' 用途: フォームが読み込まれているかどうかをチェック
' 戻り値: 引数に渡されたフォームが読み込まれている場合は、True:
'          引数に渡されたフォームが読み込まれていない場合は、False
' 参照: 『ユーザーズ ガイド』「第 25 章」
    
    Dim i

    IsLoaded = False
    For i = 0 To Forms.Count - 1
        If Forms(i).NAME = MyFormName Then
            IsLoaded = True
            Exit Function       ' フォームが見つかった時点で関数を終了します。
        End If
    Next

End Function



'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/ 引数で指定したｺﾝﾄﾛｰﾙに入力されている文字列の最後にｶｰｿﾙを移動する。   _/_/_/_/_/
'_/_/_/_/ ojControl:ｺﾝﾄﾛｰﾙ名を指定します。                                  _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Sub SetTxtBoxStart(ojControl As Control)

  Dim loLength As Long

  loLength = CmLenb(ojControl)
  
  ojControl.SelStart = loLength

End Sub
Public Function Rounded_cnv(ByVal No As Variant, ByVal FLG As Integer, Optional ByVal S As Variant) As Double
'   引数(No)の小数部を、小数部以下の処理桁数（引数(s)）を元に処理して値を返します。

'   処理方法は、引数(flg)によって「四捨五入・切り捨て・切り上げ」を判断します。
    
    '////////////////////////////////////////////////////////////////////////
    '注意事項
    '日本光機では
    '引数(flg)   1 ← 四捨五入
    '            2 ← 切捨て
    '            3 ← 切上
    '/////////////////////////////////////////////////////////////////////////

    Dim t           As Double
    Dim AddNum      As Double
    Dim TempVal     As Double
    Dim TempVal2    As Long
    Dim StrLen      As Long

    If IsNull(No) Then
        Rounded_cnv = 0
        Exit Function
    End If
    If No = 0 Then
        Rounded_cnv = 0
        Exit Function
    End If
    
'   小数以下桁数の判断
    If IsMissing(S) Then
        t = 1
    Else
        t = 10 ^ (FullFix(Abs(S)))
    End If

    TempVal = No * t

    Select Case FLG
'   切り捨ての場合
        Case 2
            AddNum = 0
'   切り上げの場合
        Case 3
            AddNum = 0.9999
'   四捨五入の場合
        Case 1
            AddNum = 0.5
    End Select

    If No < 0 Then
        AddNum = AddNum * -1
    End If

    'Fix を使用すると、２進数←→１０進数の変換誤差が発生する為、
    '文字列演算で対応。
    'TempVal = TempVal + AddNum
    'StrLen = Len(Format(TempVal, "0.0"))
    'TempVal2 = Val(Left(Trim(Str(TempVal)), (StrLen - 2)))
    'TempVal = TempVal2 / t
    Rounded_cnv = FullFix(TempVal + AddNum) / t

End Function


Public Function GetWDrive() As String
'機能： Windows のフォルダーがどこのドライブにあるかを取得する。
  
  Dim stTmp     As String * 255
  Dim stBuf     As String
  Dim loRet     As Long

  loRet = GetWindowsDirectory(stTmp, 256)
  stBuf = CmLeftB(stTmp, loRet)

  GetWDrive = stBuf
  
End Function

Public Function Check_Date(ByVal INPUTDATE As Variant) As Variant
'入力値が日付と判断できるかどうかチェック
'日付と判断できない場合は、Nullを返す
'
'対応する入力書式
'   VBで認識できる書式および、
'   yyyymmdd,yymmdd ,mmdd   ,mdd
'   yyyy/mdd,yy/mdd
'   yyyy-mdd,yy-mdd
'   yyyy mdd,yy mdd
'   geemmdd ,gemmdd
'   gee/mmdd,ge/mmdd,gee/mdd,ge/mdd
'
'   その他、多少の入力間違い（例 yy//mm/dd など）
'   は、日付に変換しようとする
'
'対応しない入力書式
'   md    （ddのみの入力と判断し、エラーを返す。
'         　m/dならば VBが日付と認識する）
'   mmd   （mddと誤認する。mm/dならば正しく認識する）
'   yymd  （mmddと誤認する。場合によってはエラー）
'   yymdd （ymmddと誤認する。場合によってはエラー）
'   gemdd （gmmddと判断し、エラーを返す）
'
'   その他、日の入力が２桁でない場合や、
'   月の入力が２桁でないものは苦手
'

    Dim GetString As String
    Dim MyCounter As Integer

    Dim modeflg As Integer
        '↑0:年 1:月 2:日 を処理中
    Dim SeparateFLG As Boolean
    Dim i As Integer

    Dim GetValue(0 To 3) As Variant
        '↑0:年 1:月 2:日 3:元号 のデータが入る
    Dim TempValue As Variant

    'VB の関数で日付と判断できる場合は、そのまま
    '処理して終了する
    If IsDate(INPUTDATE) Then
        Check_Date = Format(INPUTDATE, "yyyy/mm") '2000/05/26
        Exit Function
    End If

    '引数のチェック
    If IsNull(INPUTDATE) Or IsEmpty(INPUTDATE) Then
        GoTo NotDate_Check_Date
    End If

    INPUTDATE = StrConv(INPUTDATE, vbNarrow)
    MyCounter = Len(INPUTDATE)

    '引数の文字数が２文字以下の場合はエラー
    If MyCounter <= 2 Then
        GoTo NotDate_Check_Date
    End If

    '変数の初期化
    For i = 0 To 3
        GetValue(i) = ""
    Next i

    SeparateFLG = False
    modeflg = 2

    '入力された文字列を右から左へ検索する
    Do Until (MyCounter <= 0 Or modeflg < 0)
        GetString = Mid(INPUTDATE, MyCounter, 1)
        Select Case GetString
            Case "0" To "9"
                GetValue(modeflg) = GetString & GetValue(modeflg)
                SeparateFLG = False
                If modeflg >= 1 And Len(GetValue(modeflg)) >= 2 Then
                    modeflg = modeflg - 1
                    SeparateFLG = True
                End If
            '↓M,T,S,H 対応用
            Case "a" To "z", "A" To "Z"
                GetValue(3) = StrConv(GetString, vbUpperCase) & GetValue(3)
                SeparateFLG = False
            Case "/", "-", " ", "."
                If Len(GetValue(2)) >= 1 And SeparateFLG = False Then
                    modeflg = modeflg - 1
                End If
                SeparateFLG = True
            Case Else
                '↓不正な文字の場合は、ModeFLG < 0 にして
                '　強制的にループを終了させる
                modeflg = -1
        End Select
        MyCounter = MyCounter - 1
    Loop

    '↓ ModeFLG < 0 ならばエラー
    If modeflg < 0 Then
        GoTo NotDate_Check_Date
    End If

    '↓月、日のデータが無い場合はエラー
    If Len(GetValue(1)) < 1 Or Len(GetValue(2)) < 1 Then
        GoTo NotDate_Check_Date
    End If

    '元号の文字が入力されている場合は、文字列を年と連結する
    If Len(GetValue(0)) >= 1 And Len(GetValue(3)) >= 1 Then
        GetValue(0) = GetValue(3) & GetValue(0)
    End If

    If Len(GetValue(0)) < 1 Then
        '↓年のデータが無い場合は、システム日付の年を使う
        TempValue = Trim(Str(Year(Now()))) & "/" & GetValue(1) & "/" & GetValue(2)
    Else
        TempValue = GetValue(0) & "/" & GetValue(1) & "/" & GetValue(2)
    End If

    '↓最終的に、日付と判断できるかどうかチェック
    If IsDate(TempValue) Then
        Check_Date = Format(TempValue, "yyyy/mm/dd")
    Else
        GoTo NotDate_Check_Date
    End If

Exit_Check_Date:
    Exit Function

NotDate_Check_Date:
    Check_Date = Null
    'MsgBox "日付型の値を入力して下さい。", vbInformation, "入力確認"
    Exit Function

End Function
Public Function Explodeform(F As Form, CenterFlag As Boolean)

Const STEPS = 100
Dim FRect As RECT
Dim WIDTH As Long
Dim HEIGHT As Long
Dim i As Long
Dim X As Long, Y As Long, Cx As Long, Cy As Long
Dim hDCscreen As Long, hBrush As Long, hOldBrush

    GetWindowRect F.hwnd, FRect
    WIDTH = FRect.Right - FRect.Left
    HEIGHT = FRect.Bottom - FRect.Top
    
    hDCscreen = GetDC(0)
    hBrush = CreateSolidBrush(F.BackColor)
    hOldBrush = SelectObject(hDCscreen, hBrush)
    
    For i = 1 To STEPS
        Cx = WIDTH * (i / STEPS)
        Cy = HEIGHT * (i / STEPS)
        If CenterFlag Then
            X = FRect.Left + (WIDTH - Cx) / 2
            Y = FRect.Top + (HEIGHT - Cy) / 2
        Else
            X = FRect.Left
            Y = FRect.Top
        End If
        
        Rectangle hDCscreen, X, Y, Cx + X, Cy + Y
        
    Next i
    
    DeleteObject (hBrush)
    F.Visible = True
    
End Function
Public Function CmCreateDynaset(stsql As String, ojflds() As Object, lofieldcnt As Long, loreccnt As Long) As Object

  Dim RTN As Object
  Dim LST_SQL As String
  Dim LST_DATE As String
  Dim fnum As Integer
  Dim buf As String
  Dim ojDb As Database
  Dim i As Long
  
  
  'On Error GoTo ErrFunc

  Set ojDb = CurrentDb
  
  Set RTN = ojDb.OpenRecordset(stsql, dbOpenDynaset)
  lofieldcnt = RTN.Fields.Count
  ReDim ojflds(lofieldcnt)
  For i = 0 To lofieldcnt - 1
       Set ojflds(i) = RTN.Fields(i)
  Next i
  RTN.MoveLast
  loreccnt = RTN.RecordCount
  Set CmCreateDynaset = RTN

End Function
' @f
'
' 機　能　 :入力値の全角／半角のチェック
' 返り値　 :True/False
' 引　数　 :sStr   - 入力値
' 　　　    sKubun - 全角のみＯＫ（”Ｋ”）／半角のみＯＫ（”Ａ”）
'
' 機能説明 :引数で指定した入力値に、全角以外または半角以外の文字列が
' 　　　　  入力されている時、Ｆａｌｓｅを返す
'
Public Function CmChkCher(sStr As String, sKubun As String) As Boolean

Dim swk As String * 2
Dim iIdx As Integer
Dim iMax As Integer
        
    '**************************************************
    '*  戻り値のセット（Ｔｒｕｅ）                      *
    '**************************************************
    CmChkCher = True
    
    '**************************************************
    '*  最大文字列数の検索　　                          *
    '**************************************************
    iMax = CmLenb(Nz(sStr, ""))
    
    '**************************************************
    '*  文字列数の全角／半角のチェック                  *
    '**************************************************
    iIdx = 0
    
    Do
    
        iIdx = iIdx + 1
        If iIdx > iMax Then
            Exit Do
        End If
        
        swk = Mid(sStr, iIdx, 1)
        If IsNull(swk) Then
            Exit Do
        End If
        
        'MsgBox Asc(swk) & " :(" & iIdx & "): " & swk & " :(" & iMax & "): "
        
        If sKubun = "A" Then    '半角文字のみ許可
            Select Case Asc(swk)
                Case 0 To 255
                    '半角文字
                    CmChkCher = True
                Case Else
                    '全角文字
                    CmChkCher = False
                    Exit Do
            End Select
        Else                     '全角文字のみ許可
            Select Case Asc(swk)
                Case 0 To 255
                    '半角文字
                    CmChkCher = False
                    Exit Do
                Case Else
                    '全角文字
                    CmChkCher = True
                    iMax = iMax - 1
                    If iMax < 1 Then
                        Exit Do
                    End If
            End Select
        End If
    
    Loop
    
End Function

Public Function CmNumeric(ByVal vCode As Variant) As Variant
  '数値に判断できる時は、Cdbl型に変換する
  '数値に判断できない時は、Nullを返す
  
  '単にClngを使うと数値以外のときにエラーになるから
  'とりあえず、すべての状態で判断できるようにした
  '小数点数値は、Cdblで丸まるから注意してださい。
  
  '99/05/31 Add takasaka

  CmNumeric = Null

  If IsMissing(vCode) Then
    Exit Function
  End If
      
  If IsEmpty(vCode) Then
    Exit Function
  End If
  
  If Nz(vCode, "") = "" Then
    Exit Function
  End If
    
  If IsNumeric(vCode) = False Then
    Exit Function
  End If
  
  
  CmNumeric = CDbl(vCode)
  
End Function


Public Function CmParChar(oChar As Variant) As String
  '文字列の値をSQL文にセットする時は、
  ' "'"(シングルクォーテション)に対応するために
  'Chr$(34) & [変数] & Chr$(34) で引き渡す。
  '99/09/14   Takasaka
  
  CmParChar = Chr$(34) & oChar & Chr$(34)
  
End Function


