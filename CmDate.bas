Attribute VB_Name = "CmDate"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME :
'
'        PROGRAM_NAME    : 日付関係共通関数群
'        PROGRAM_ID      :
'        PROGRAM_KBN     :
'
'        CREATE          : 2005/07/29
'        CERATER         : S.SHIBAZAKI
'
'        UPDATE          : 2005/12/15
'        UPDATER         : S.SHIBAZAKI
'        NOTE            : 日付の最低値をチェックするように修正
'
'==============================================================================*
Option Compare Database
Option Explicit

'2005/12/15 Shibazaki 追加
Private Const P_MIN_DATE As String = "1753/01/01"

'==============================================================================*
'
' 日付正当性チェック
'
' 下記対応書式ののっとった形式でかつ、日付として正当性がある場合、
' yyyy/mm/dd形式で日付を返却する。
' 不正な日付の場合、NULLを返却する。
' パラメータ指定された日付がNULL/空文字の場合、空文字を返却する。
'
' << 対応書式 >>
' ①年月日スラッシュ区切り（年月日それぞれの桁数は問わない）
' ②月日スラッシュ区切り（月日それぞれの桁数は問わない）
' ③"yyyymmdd"
' ④"yymmdd"
' ⑤"mmdd"
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function CmDateChecker(ByVal varInput As Variant) As Variant

    CmDateChecker = Null
    
    '引数のチェック
    If Nz(varInput) = "" Then
        CmDateChecker = ""
        Exit Function
    End If
    
    Dim lngInStr            As Long
    Dim intCount            As Integer
    Dim strYear             As String
    Dim strMonth            As String
    Dim strDay              As String
    Dim intYearLen          As Integer
    Dim strDate             As String
    
    'スラッシュ区切りでの入力をばらして格納する
    Dim strValue(0 To 2)    As String
    
    intCount = 0
    strDate = varInput
    
    Do
        'スラッシュを探す
        lngInStr = InStr(strDate, "/")
        If lngInStr = 0 Then
            '見つからなければループ終了
            Exit Do
        Else
            'スラッシュの個数
            intCount = intCount + 1
            '三つ以上見つかったらループ終了
            If intCount > 2 Then
                Exit Do
            End If
            'スラッシュより前を格納
            strValue(intCount - 1) = Left(strDate, lngInStr - 1)
            'スラッシュより後ろを格納
            strValue(intCount) = Mid(strDate, lngInStr + 1)
            strDate = Mid(strDate, lngInStr + 1)
        End If
    Loop
    
    If intCount = 0 Then
        'スラッシュがない
        If Len(varInput) = 8 Then
            'yyyymmddとする
            intYearLen = 4
        ElseIf Len(varInput) = 6 Then
            'yymmddとする
            intYearLen = 2
        ElseIf Len(varInput) = 4 Then
            'mmddとする
            intYearLen = 0
        Else
            'それ以外はエラーとする
            Exit Function
        End If
        strYear = Left(varInput, intYearLen)
        strMonth = Mid(varInput, intYearLen + 1, 2)
        strDay = Right(varInput, 2)
    ElseIf intCount = 2 Then
        'スラッシュ二つある
        '年月日がスラッシュ区切りとする
        strYear = strValue(0)
        strMonth = strValue(1)
        strDay = strValue(2)
    ElseIf intCount = 1 Then
        'スラッシュ一つある
        '月日がスラッシュ区切りとする
        strYear = ""
        strMonth = strValue(0)
        strDay = strValue(1)
    Else
        'それ以外はエラーとする
        Exit Function
    End If
    
    '年が入力されていない場合、現在年を使用
    If Nz(strYear) = "" Then
        strYear = Format(Now, "yyyy")
    End If
    
    strDate = strYear & "年" & strMonth & "月" & strDay & "日"
    
    If Not IsDate(strDate) Then
        '日付として認められない
        'エラーとする
        Exit Function
    End If
    
    '2005/12/15 Shibazaki 追加↓
    '日付の最小値より以前の場合はエラーとする。
    If StrComp(Format(strDate, "yyyy/mm/dd"), P_MIN_DATE) < 0 Then
        Exit Function
    End If
    '2005/12/15 Shibazaki 追加↑
    
    CmDateChecker = Format(strDate, "yyyy/mm/dd")

End Function

