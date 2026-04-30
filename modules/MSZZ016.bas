Attribute VB_Name = "MSZZ016"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : Unicode文字種チェック
'        PROGRAM_ID      : MSZZ016
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2005/03/18
'        CERATER         : K.ISHIZAKA
'        Ver             : 0.0
'
'        UPDATE          : 2007/02/14
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.1
'                        : 半角その他 ＆ 全角その他 に対応
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
        
'=== 半角 ==============================================
Public Const charTypeNumber         As Integer = &H1    '半角 数値
Public Const charTypeAlphabetU      As Integer = &H2    '半角 英字 大文字
Public Const charTypeAlphabetL      As Integer = &H4    '半角 英字 小文字
Public Const charTypeKana           As Integer = &H8    '半角 カナ
Public Const charTypeMark           As Integer = &H40   '半角 スペース' '、点'･'、括弧'('、')'、マイナス'-'
Public Const charTypeOther          As Integer = &H80   '半角 その他

Public Const charTypeAlphabet       As Integer = &H6    '半角 英字 大小
Public Const charTypeAlphaNumber    As Integer = &H7    '半角 英数字
Public Const charTypeAll            As Integer = &HFF   '半角 全て

'=== 全角 ==============================================
Public Const charTypeWideNumber     As Integer = &H100  '全角 数値
Public Const charTypeWideAlphabetU  As Integer = &H200  '全角 英字 大文字
Public Const charTypeWideAlphabetL  As Integer = &H400  '全角 英字 小文字
Public Const charTypeWideKana       As Integer = &H800  '全角 カナ
Public Const charTypeWideHiragana   As Integer = &H1000 '全角 ひらがな
Public Const charTypeWideMark       As Integer = &H4000 '全角 スペース' '、点'･'、括弧'('、')'、マイナス'-'
Public Const charTypeWideOther      As Integer = &H8000 '全角 その他

Public Const charTypeWideAlphabet       As Integer = &H600  '全角 英字 大小
Public Const charTypeWideAlphaNumber    As Integer = &H700  '全角 英数字
Public Const charTypeWideAll            As Integer = &HFF00 '全角 全て

'==============================================================================*
'
'        MODULE_NAME      :テスト用
'        MODULE_ID        :TEST_MSZZ0016_M00
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub TEST_MSZZ0016_M00()
    
    If MSZZ0016_M00("019AZazｱﾝﾊﾟﾊﾞｰ ･()-０１９ＡＺａｚアンガバパあん　・（）－", charTypeAll + charTypeWideNumber + charTypeWideAlphabetU + charTypeWideAlphabetL + charTypeWideKana + charTypeWideHiragana) = True Then
       MsgBox ("TRUE")
    Else
       MsgBox ("FALSE")
    End If

End Sub
'==============================================================================*
'
'       MODULE_NAME     : Unicode文字種チェック
'       MODULE_ID       : MSZZ016_M00
'       CREATE_DATE     : 2005/03/18
'       PARAM           : strUnicode        Unicode文字列
'                       : intOption         文字種チェック定数（組み合わせ可）
'       RETURN          : 正常(True) エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ0016_M00(ByVal strUnicode As String, ByVal intOption As Long) As Boolean
    Dim i                   As Integer
    
    For i = 1 To Len(strUnicode)
        '=== 半角 ==============================================
        Select Case AscW(Mid(strUnicode, i, 1))
        Case &H30 To &H39       '半角数値
            If (intOption And charTypeNumber) = 0 Then
                MSZZ0016_M00 = False
                Exit Function
            End If
        Case &H41 To &H5A       '半角英字 大文字
            If (intOption And charTypeAlphabetU) = 0 Then
                MSZZ0016_M00 = False
                Exit Function
            End If
        Case &H61 To &H7A       '半角英字 小文字
            If (intOption And charTypeAlphabetL) = 0 Then
                MSZZ0016_M00 = False
                Exit Function
            End If
        Case &HFF66 To &HFF9F   '半角カナ
            If (intOption And charTypeKana) = 0 Then
                MSZZ0016_M00 = False
                Exit Function
            End If
        Case &H20, &HFF65, &H28, &H29, &H2D '半角 スペース' '、点'･'、括弧'('、')'、マイナス'-'
            If (intOption And charTypeMark) = 0 Then
                MSZZ0016_M00 = False
                Exit Function
            End If
        '>>> INSERT START 2007/02/14 K.ISHIZAKA >>>
        Case &H21 To &H7E       '半角その他
            If (intOption And charTypeOther) = 0 Then
                MSZZ0016_M00 = False
                Exit Function
            End If
        '<<< INSERT END   2007/02/14 K.ISHIZAKA <<<
        '=== 全角 ==============================================
        Case &HFF10 To &HFF19   '全角数値
            If (intOption And charTypeWideNumber) = 0 Then
                MSZZ0016_M00 = False
                Exit Function
            End If
        Case &HFF21 To &HFF3A   '全角英字 大文字
            If (intOption And charTypeWideAlphabetU) = 0 Then
                MSZZ0016_M00 = False
                Exit Function
            End If
        Case &HFF41 To &HFF5A   '全角英字 小文字
            If (intOption And charTypeWideAlphabetL) = 0 Then
                MSZZ0016_M00 = False
                Exit Function
            End If
        Case &H30A1 To &H30FA   '全角カナ
            If (intOption And charTypeWideKana) = 0 Then
                MSZZ0016_M00 = False
                Exit Function
            End If
        Case &H3041 To &H3094   '全角ひらがな
            If (intOption And charTypeWideHiragana) = 0 Then
                MSZZ0016_M00 = False
                Exit Function
            End If
        Case &H3000, &H30FB, &HFF08, &HFF09, &HFF0D '全角 スペース' '、点'･'、括弧'('、')'、マイナス'-'
            If (intOption And charTypeWideMark) = 0 Then
                MSZZ0016_M00 = False
                Exit Function
            End If
        Case Else
'            MSZZ0016_M00 = False                                               'DELETE 2007/02/14 K.ISHIZAKA
'            Exit Function                                                      'DELETE 2007/02/14 K.ISHIZAKA
            '>>> INSERT START 2007/02/14 K.ISHIZAKA >>>
            If LenB(StrConv(Mid(strUnicode, i, 1), vbFromUnicode)) = 2 Then '全角その他
                If (intOption And charTypeWideOther) = 0 Then
                    MSZZ0016_M00 = False
                    Exit Function
                End If
            Else
                'ここには制御文字などが該当する
                MSZZ0016_M00 = False
                Exit Function
            End If
            '<<< INSERT END   2007/02/14 K.ISHIZAKA <<<
        End Select
    Next
    MSZZ0016_M00 = True
End Function
'****************************  ended or program ********************************
