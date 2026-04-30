Attribute VB_Name = "MSZZ043"
'****************************  strat or program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬総合システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : ANSIコードに変換できる文字列かチェック
'       PROGRAM_ID      : MSZZ043
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2008/04/12
'       CERATER         : K.ISHIZAKA
'       Ver             : 0.0
'
'       UPDATE          :
'       UPDATER         :
'       Ver             :
'                       :
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'
'       MODULE_NAME     : ANSIコードに変換できる文字列かチェック
'       MODULE_ID       : MSZZ043_M00
'       CREATE_DATE     : 2008/04/12
'       PARAM           : strUnicode        Unicode文字列
'       RETURN          : (0)正常 (1～)エラーとなった文字の位置
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ0043_M00(ByVal strUnicode As String) As Long
    Dim i                   As Long
    Dim lngUnicode          As Long
    Dim ch                  As String
    On Error GoTo ErrorHandler
    
    For i = 1 To Len(strUnicode)
        ch = Mid(strUnicode, i, 1)
        lngUnicode = AscW(ch)
        Select Case lngUnicode
        Case &H0 To &H33FF                      '各種文字(アルファベット、組合せ型ハングル、仮名など)・記号 ※&H00～&HFFはASCIIコード互換
            Select Case lngUnicode
            Case &HA, &HD                           '改行コード
            Case &H20 To &H7E                       '半角
            Case &H3000                             '全角スペース
            Case &H3041 To &H3094                   '全角ひらがな
            Case &H30A1 To &H30FB                   '全角カナ
            Case Else                               'その他
                If checkAnsi(ch) = False Then
                    MSZZ0043_M00 = i
                    Exit Function
                End If
            End Select
        Case &H3400 To &H4DFF                   '完成型ハングル移動後の空き？ 漢字？
            MSZZ0043_M00 = i
            Exit Function
        Case &H4E00 To &H7FFF, &H8000 To &H9FFF '漢字(中国・台湾・日本・韓国)
            If checkAnsi(ch) = False Then
                MSZZ0043_M00 = i
                Exit Function
            End If
        Case &HA000 To &HABFF                   '拡張用予約領域
            MSZZ0043_M00 = i
            Exit Function
        Case &HAC00 To &HD7FF                   '完成型ハングル
            MSZZ0043_M00 = i
            Exit Function
        Case &HD800 To &HDFFF                   'サロゲート・ペア
            MSZZ0043_M00 = i
            Exit Function
        Case &HE000 To &HE8FF                   '私用予約領域
            MSZZ0043_M00 = i
            Exit Function
        Case &HF900 To &HFFFD                   '互換用文字など
            Select Case lngUnicode
            Case &HF900 To &HFA6F                   '漢字
                If checkAnsi(ch) = False Then
                    MSZZ0043_M00 = i
                    Exit Function
                End If
            Case &HFF00 To &HFF9F                   '半角と全角
            'Case &HFFE0 To &HFFEF                   '記号文字
            Case Else                               'その他
                If checkAnsi(ch) = False Then
                    MSZZ0043_M00 = i
                    Exit Function
                End If
            End Select
        End Select
    Next
    MSZZ0043_M00 = 0
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MSZZ0043_M00" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ANSIコードに変換できる文字かチェック
'       MODULE_ID       : checkAnsi
'       CREATE_DATE     : 2008/04/12
'       PARAM           : ch            Unicode文字
'                         半角は事前にはじいておかないとエラーになる？
'       RETURN          : 正常(True)／エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function checkAnsi(ByVal ch As String) As Boolean
    Dim i                   As Long

    On Error Resume Next
    i = Asc(StrConv(ch, vbFromUnicode))
    checkAnsi = (Err.Number = 0)
    Err.Clear
End Function

'****************************  ended or program ********************************
