Attribute VB_Name = "CmKeyPress"
'//////////////////////////////////////////
'キー入力の取消関数群
'//////////////////////////////////////////
Option Compare Database
Option Explicit

Public Function CmKeyPress_For_Char(ByVal objControl As Object, ByRef KeyAscii As Integer, ByVal iKey_Max As Integer)
  'キー入力制御用(漢字が入力された時の対応版)
  'といいたいところだがBeForeupdateにも、桁数チェックは必要
  '理由 テキストボックスにコピーで文字を貼り付けると制御が利かないため
  '99/05/31 ADD takasaka
  '//////////////////////////////////////////////////////////////
  '小文字は大文字へ
  '//////////////////////////////////////////////////////////////
Dim oCont      As Object
Dim IMaxLength As Integer
Dim intKanji   As Integer

  Set oCont = objControl
  
  IMaxLength = iKey_Max
  
    With oCont
        Select Case KeyAscii
            Case 8, 9, 10           '8:BS 9:TAB 10:LF
            Case vbKeyReturn
                KeyAscii = 0
            Case Else
                If KeyAscii < 0 Then
                  '全角文字の時は、Keyascii は、負の値だから
                  intKanji = 1
                End If
                If LenB(StrConv(.Text, vbFromUnicode)) - .SelLength + intKanji > IMaxLength - 1 And .SelLength = 0 Then
                    KeyAscii = 0
                End If
        End Select
        KeyAscii = Asc(StrConv(Chr(KeyAscii), vbUpperCase))   '小文字は、大文字に変換
    End With

End Function

Public Function CmKeyPress_For_Char2(ByVal objControl As Object, ByRef KeyAscii As Integer, ByVal iKey_Max As Integer)
  'キー入力制御用(漢字が入力された時の対応版)
  'といいたいところだがBeForeupdateにも、桁数チェックは必要
  '理由 テキストボックスにコピーで文字を貼り付けると制御が利かないため
  '99/05/31 ADD takasaka
  '//////////////////////////////////////////////////////////////
  '小文字は大文字へ変換しない
  '//////////////////////////////////////////////////////////////
Dim oCont      As Object
Dim IMaxLength As Integer
Dim intKanji   As Integer

  Set oCont = objControl
  
  IMaxLength = iKey_Max
  
    With oCont
        Select Case KeyAscii
            Case 8, 9, 10           '8:BS 9:TAB 10:LF
            Case vbKeyReturn
                KeyAscii = 0
            Case Else
                If KeyAscii < 0 Then
                  '全角文字の時は、Keyascii は、負の値だから
                  intKanji = 1
                End If
                If LenB(StrConv(.Text, vbFromUnicode)) - .SelLength + intKanji > IMaxLength - 1 And .SelLength = 0 Then
                    KeyAscii = 0
                End If
        End Select
    End With

End Function

Public Sub CmKeyPress_For_Tel(ojCtrl As Control, KeyAscii As Integer, inLen As Integer)
    '郵便番号や、電話番号などで活用すると良いと思う
    '99/05/31 ADD takaska
    
    On Error GoTo CmKeyPress_For_Num_Err

    Dim stValue As String

    With ojCtrl
        Select Case KeyAscii
            Case 8, 9, 10, 13                               ' 8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9"), Asc("-")   ' 0～9迄と ，と - を許可する
                ''' 入力された文字列が引数の最大文字長に達した場合、入力したキーを無効にする
                ''' 絶対値を求めることで、負数にも対応

                If Len(.Text) > (inLen - 1) And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With

    Exit Sub

CmKeyPress_For_Num_Err:

  KeyAscii = 0

End Sub


' @f
'
' 機　能　 :キーの取消 日付のみ対応
' 返り値　 :Keyacii
' 引　数　 :ojCtrl   - コントロール
' 　　　    Keyascii - Keyascii
' 　　　　  inLen    - 入力する最大文字長
'
' 機能説明 :引数で指定したコントロールに入力したAsciiCodeを有効か無効か判断し
' 　　　　  返り値をKeyasciiに格納する
'
Public Sub CmKeyPress_For_Date(ojCtrl As Control, KeyAscii As Integer, inLen As Integer)
    
    With ojCtrl
        Select Case KeyAscii
            Case 8, 9, 10, 13                     ' 8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9"), Asc("/")   ' 0～9迄、/ を許可する
                ''' 入力された文字列が引数の最大文字長に達した場合、入力したキーを無効にする
                If Len(.Text) > (inLen - 1) And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With

End Sub

' @f
'
' 機　能　 :キーの取消 日付のみ対応
' 返り値　 :Keyacii
' 引　数　 :ojCtrl   - コントロール
' 　　　    Keyascii - Keyascii
' 　　　　  inLen    - 入力する最大文字長
'
' 機能説明 :引数で指定したコントロールに入力したAsciiCodeを有効か無効か判断し
' 　　　　  返り値をKeyasciiに格納する
'
Public Sub CmKeyPress_For_Time(ojCtrl As Control, KeyAscii As Integer, inLen As Integer)
    
    With ojCtrl
        Select Case KeyAscii
            Case 8, 9, 10, 13                     ' 8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9"), Asc(":")   ' 0～9迄、/ を許可する
                ''' 入力された文字列が引数の最大文字長に達した場合、入力したキーを無効にする
                If Len(.Text) > (inLen - 1) And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With

End Sub

' @f
'
' 機　能　 :キーの取消 "0"～"9","-",","のみ許可 99/09/08 takasaka
' 返り値　 :Keyacii
' 引　数　 :ojCtrl   - コントロール
' 　　　    Keyascii - Keyascii
' 　　　　  inLen    - 入力する最大文字長
'
' 機能説明 :引数で指定したコントロールに入力したAsciiCodeを有効か無効か判断し
' 　　　　  返り値をKeyasciiに格納する
'
Public Sub CmKeyPress_For_Num(ojCtrl As Control, KeyAscii As Integer, inLen As Integer)

    On Error GoTo CmKeyPress_For_Num_Err

    Dim stValue As String

    With ojCtrl
        Select Case KeyAscii
            Case 8, 9, 10, 13                               ' 8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9"), Asc(","), Asc("-")   ' 0～9迄と ，と - を許可する
                ''' 入力された文字列が引数の最大文字長に達した場合、入力したキーを無効にする
                ''' 絶対値を求めることで、負数にも対応
                If Len(.Text) > 1 Then
                  stValue = CStr(Abs(CLng(Format$(.Text, "#0"))))
                End If

                If Len(stValue) > (inLen - 1) And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With

    Exit Sub

CmKeyPress_For_Num_Err:

  KeyAscii = 0

End Sub

' @f バグ対応中
'
' 機　能　 :キーの取消 数値(少数有)のみ対応
' 返り値　 :Keyacii
' 引　数　 :ojCtrl   - コントロール
' 　　　    Keyascii - Keyascii
' 　　　　  inLen    - 入力する最大文字長
' 　　　　  inKeta   - 少数以下桁数
'
' 機能説明 :引数で指定したコントロールに入力したAsciiCodeを有効か無効か判断し
' 　　　　  返り値をKeyasciiに格納する
'
Public Sub CmKeyPress_For_Dec(ojCtrl As Control, KeyAscii As Integer, inLen As Integer, inKeta As Integer)

    Dim stValue As String
    Dim stTen   As String
    Dim inI     As Integer

    stTen = ""

    With ojCtrl
        Select Case KeyAscii
            Case 8, 9, 10, 13                               ' 8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9"), Asc(","), Asc("-"), Asc(".")  ' 0～9迄と ，と - を許可する
                ''' 入力された文字列が引数の最大文字長に達した場合、入力したキーを無効にする
                ''' 絶対値を求めることで、負数にも対応
                If Len(.Text) > 1 Then
                  For inI = 1 To inKeta
                    stTen = stTen & "0"
                  Next

                  stValue = Trim$(CStr(Abs(Val(Format$(.Text, "#0." & stTen)) * 10 ^ Len(.Text))))
                  stValue = CmLeftB(stValue, Len(.Text) + 1)
                End If

                If Len(stValue) > (inLen - 1) And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With

End Sub

' @f
'
' 機　能　 :キーの取消 数値(整数)のみ対応
' 返り値　 :Keyacii
' 引　数　 :ojCtrl   - コントロール
' 　　　    Keyascii - Keyascii
' 　　　　  inLen    - 入力する最大文字長
'
' 機能説明 :引数で指定したコントロールに入力したAsciiCodeを有効か無効か判断し
' 　　　　  返り値をKeyasciiに格納する
'
Public Sub CmKeyPress_For_Num1(ojCtrl As Control, KeyAscii As Integer, inLen As Integer)

    On Error GoTo CmKeyPress_For_Num_Err

    Dim stValue As String

    With ojCtrl
        Select Case KeyAscii
            Case 8, 9, 10, 13                               ' 8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9")     ' 0～9迄と .を許可する
                ''' 入力された文字列が引数の最大文字長に達した場合、入力したキーを無効にする
                ''' 絶対値を求めることで、負数にも対応
                If CmLenb(.Text) > (inLen - 1) And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With

    Exit Sub

CmKeyPress_For_Num_Err:

  KeyAscii = 0

End Sub







Public Sub ETC_CUREN_RTN(ByRef cntTXT As Control, ByVal intLen As Integer, ByVal intTEN As Integer, ByRef KeyAscii As Integer)
'***************************************************************************
'   テキストボックス上で数値入力を行う時の制御関数(並木さん作成)
'
'  引数
'      cntTXT---------入力対象となるコントロール名（例　TXT_120）
'      intLEN---------入力する整数部の桁数
'      intTEN---------入力する小数部の桁数
'      KeyAscii-------KEYPERSSイベントで使用するので、ASCIIキーをいれる
' この関数は、金額入力に必要な物のみを入力できるようにして有ります。
' ただし、カンマ編集はできません。面倒
' 使用方法は、上記の通りですが、この関数は、keypressのみ使用可能です。
' 小数を入力する時は、姉妹品の”ETC_CURENsub_RTN”をkeydownイベントに
' 記述してください
'***************************************************************************

'↓****************************************  φ(..)
'マイナス入力ができるように変更をかける
'Custom 99/08/30 TAKASAKA
'↑****************************************  φ(..)


Dim intSTRAT As Integer
Dim intSEL As Integer
Dim intACS As Integer
Dim intPRI As Integer
Dim intTWK As Integer
Dim intPos As Integer

Dim intMin As Integer   '"-"が入力されているかどうかの判断

        
    '↓追加 Start   98/11/27 furusawa
    '入力可能なKeyAscii以外は、０にしてサブルーチン終了
    Select Case KeyAscii
        Case 8, 9, 10, 13
        Case Asc(0) To Asc(9), Asc("."), Asc("-") '*--99/08/30
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
    '↑追加 End
    
    
    '↓****************************************  φ(..)
    If KeyAscii = Asc("-") Then
      
      If CmLenb(cntTXT.Text) = cntTXT.SelLength Then
        Exit Sub
      End If
      
      intMin = InStr(1, cntTXT.Text, "-")
      
      If intMin = 0 Then
      
        '"-"という文字が入力されていない時の動き
        
        If cntTXT.Text <> "" Then
          
          If Left(cntTXT.Text, 1) = "." Then
            KeyAscii = 0
            cntTXT.Text = "-0" & cntTXT.Text
            Exit Sub
          Else
            KeyAscii = 0
            cntTXT.Text = "-" & cntTXT.Text
            Exit Sub
          End If
          
        Else
          Exit Sub
        End If
        
      Else
        '"-"という文字が入力されている時の動き
        KeyAscii = 0
        Exit Sub
      End If
    End If
    '↑****************************************  φ(..)
    
    
    
    '小数部桁数=0なら、小数点が入らないようにする
    If intTEN = 0 And KeyAscii = 46 Then
        KeyAscii = 0
    End If
    
    intACS = KeyAscii
    intPRI = 0
    intTWK = intTEN
        
    If KeyAscii >= 31 Then
        If InStr(CVar(cntTXT.Text), ".") <> 0 Then
            intPRI = 1
            If InStr(CVar(cntTXT.Text), ".") <= cntTXT.SelStart Then
               intACS = 46
            End If
        Else
            intTWK = 0
        End If
    End If
        
    Select Case intACS
        Case 0 To 31   '制御キー
            If KeyAscii = 8 Then
                Select Case cntTXT.SelLength
                    Case Is > 0
                        If InStr(CVar(cntTXT), ".") <> 0 Then
                            intPos = cntTXT.SelStart + cntTXT.SelLength
                            If intPos >= InStr(CVar(cntTXT.Text), ".") Then
                                cntTXT.SelStart = cntTXT.SelStart
                                cntTXT.SelLength = Len(cntTXT.Text) - cntTXT.SelStart
                            End If
                        End If
                    Case Else
                        If InStr(CVar(cntTXT), ".") <> 0 Then
                            intSTRAT = cntTXT.SelStart
                            If InStr(CVar(cntTXT.Text), ".") = intSTRAT Then
                                intPos = Len(cntTXT.Text)
                                cntTXT.SelStart = intSTRAT - 1
                                cntTXT.SelLength = intPos
                            End If
                        End If
                End Select
            End If
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 45
            '選択されてる時  45
            Select Case cntTXT.SelLength
                Case Is > 0
                    '項目にハイライトがされている時
                    intSTRAT = cntTXT.SelStart
                    intSEL = cntTXT.SelLength
                    If cntTXT.SelStart <> 0 Then
                        If KeyAscii = 45 Then
                            If InStr(CVar(cntTXT.Text), "-") <> 0 Then
                                KeyAscii = 0
                                Exit Sub
                            End If
                            cntTXT.Text = "-" & cntTXT.Text
                            cntTXT.SelStart = intSTRAT + 1
                            cntTXT.SelLength = intSEL
                            KeyAscii = 0
                        End If
                    End If
                Case Else
                    If Len(cntTXT.Text) - intTWK >= intLen + intPRI Then
                        If KeyAscii = 45 Then
                            If InStr(CVar(cntTXT.Text), "-") <> 0 Then
                                KeyAscii = 0
                                Exit Sub
                            End If
                            cntTXT.Text = "-" & cntTXT.Text
                            cntTXT.SelStart = Len(cntTXT.Text)
                            KeyAscii = 0
                        Else
                            If InStr(CVar(cntTXT.Text), "-") <> 0 Then
                                If Len(cntTXT.Text) >= intLen + 1 + intPRI Then
                                   KeyAscii = 0
                                End If
                            Else
                                 KeyAscii = 0
                            End If
                        End If
                    Else
                        If KeyAscii = 45 Then
                            If InStr(CVar(cntTXT.Text), "-") = 0 Then
                                cntTXT.Text = "-" & cntTXT.Text
                                cntTXT.SelStart = Len(cntTXT.Text)
                            End If
                            KeyAscii = 0
                        End If
                    End If
            End Select
        Case 46
            Select Case cntTXT.SelLength
                Case Is > 0
                    '項目にハイライトがされている時
                    intSTRAT = cntTXT.SelStart
                    intSEL = cntTXT.SelLength
                    If InStr(CVar(cntTXT.Text), ".") <> 0 Then
                        If intSTRAT + intSEL < InStr(CVar(cntTXT.Text), ".") Then
                           KeyAscii = 0
                        End If
                    Else
                        intPos = intSTRAT + intSEL
                        If Len(cntTXT.Text) - intPos > intTEN Then
                            KeyAscii = 0
                        End If
                    End If
                    
                Case Else
                    If KeyAscii = 46 And InStr(CVar(cntTXT.Text), ".") <> 0 Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                    intSTRAT = cntTXT.SelStart
                    intSEL = cntTXT.SelLength
                    If InStr(CVar(cntTXT.Text), ".") = 0 Then
                        intPos = intSTRAT + intSEL
                        If Len(cntTXT.Text) - intPos > intTEN Then
                            KeyAscii = 0
                        End If
                        Exit Sub
                    End If
    
                    If Len(cntTXT.Text) - InStr(CVar(cntTXT.Text), ".") >= intTEN Then
                        If KeyAscii = 45 Then
                            If InStr(CVar(cntTXT.Text), "-") <> 0 Then
                                KeyAscii = 0
                                Exit Sub
                            End If
                            cntTXT.Text = "-" & cntTXT.Text
                            cntTXT.SelStart = Len(cntTXT.Text)
                            KeyAscii = 0
                        Else
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii = 45 Then
                            If InStr(CVar(cntTXT.Text), "-") = 0 Then
                                cntTXT.Text = "-" & cntTXT.Text
                                cntTXT.SelStart = Len(cntTXT.Text)
                            End If
                            KeyAscii = 0
                        End If
                    End If
            End Select
        Case Else
            KeyAscii = 0
    End Select
    
End Sub



Public Sub ETC_CURENsub_RTN(ByRef cntTXT As Control, ByVal KeyCode As Integer, ByRef Shift As Integer)
'***************************************************************************
'   テキストボックス上で数値入力を行う時の制御関数(並木さん作成)
'
'ETC_CUREN_RTNの機能を補助するための関数
'keypressではdelkeyが発生しないためｋｅｙｄｏｗｎで補助をする。
'小数点をDELで消したときは、小数部を消すために使用
'  引数
'      cntTXT---------入力対象となるコントロール名（例　TXT_120）
'      KeyCode-------keydownイベントで使用するので、KeyCodeキーをいれる
'      Shift---------keydownイベントで使用するので、Shiftキーをいれる
'使用方法は、上記の通りですが、この関数は、keydownでのみ使用可能です。
'***************************************************************************

Dim intSTRAT As Integer
Dim intPos As Integer

    If KeyCode = vbKeyDelete Then
        Select Case cntTXT.SelLength
            Case Is > 0
                If InStr(CVar(cntTXT), ".") <> 0 Then
                    intPos = cntTXT.SelStart + cntTXT.SelLength
                    If intPos >= InStr(CVar(cntTXT.Text), ".") Then
                        cntTXT.SelStart = cntTXT.SelStart
                        cntTXT.SelLength = Len(cntTXT.Text) - cntTXT.SelStart
                    End If
                End If
        
            Case Else
                If Shift = 0 Then
                    If InStr(CVar(cntTXT), ".") <> 0 Then
                        intSTRAT = cntTXT.SelStart
                        If InStr(CVar(cntTXT.Text), ".") = intSTRAT + 1 Then
                            intPos = Len(cntTXT.Text)
                            cntTXT.SelStart = intSTRAT
                            cntTXT.SelLength = intPos
                        End If
                    End If
                Else
                    If InStr(CVar(cntTXT), ".") <> 0 Then
                        intSTRAT = cntTXT.SelStart
                        If InStr(CVar(cntTXT.Text), ".") = intSTRAT Then
                            intPos = Len(cntTXT.Text)
                            cntTXT.SelStart = intSTRAT - 1
                            cntTXT.SelLength = intPos
                        End If
                    End If
                End If
        End Select
    End If

End Sub


