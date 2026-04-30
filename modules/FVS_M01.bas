Attribute VB_Name = "FVS_M01"
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　共通モジュール
'   プログラムＩＤ　：　FVS_M01
'   作　成　日　　　：  2003/06/09
'   作　成　者　　　：  Eagle Soft 田島
'**********************************************
'修正履歴
'   修　正　日　　　：
'   修　正　者　　　：
'   修　正　内　容　：
'**********************************************
Option Compare Database
Option Explicit
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
Public Sub CmKeyPress_For_Num2(ojCtrl As Control, KeyAscii As Integer, inLen As Integer)

    On Error GoTo CmKeyPress_For_Num_Err

    Dim stValue As String

    With ojCtrl
        Select Case KeyAscii
            Case 8, 9, 10, 13                               ' 8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9")     ' 0～9迄と .を許可する
                ''' 入力された文字列が引数の最大文字長に達した場合、入力したキーを無効にする
                ''' 絶対値を求めることで、負数にも対応
                If FvsModLenb(.Text) > (inLen - 1) And .SelLength = 0 Then
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

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/ Unicode対応のLenBの代用                                           _/_/_/_/_/
'_/_/_/_/ stArg：任意の文字列式を指定します。                                _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Function FvsModLenb(ByVal stArg As Variant) As Long
    
    FvsModLenb = LenB(FVSModStrConv(stArg, vbFromUnicode))

End Function
Public Function FVSModStrConv(st1, flag)

  FVSModStrConv = StrConv(st1, flag)

End Function
Public Function FVSModNumeric(ByVal vCode As Variant) As Variant
  '数値に判断できる時は、Cdbl型に変換する
  '数値に判断できない時は、Nullを返す
  
  '単にClngを使うと数値以外のときにエラーになるから
  'とりあえず、すべての状態で判断できるようにした
  '小数点数値は、Cdblで丸まるから注意してださい。
  
  '99/05/31 Add takasaka

  FVSModNumeric = Null

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
  
  
  FVSModNumeric = CDbl(vCode)
  
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
Public Function FVSModChkCher(sStr As String, sKubun As String) As Boolean

Dim swk As String * 2
Dim iIdx As Integer
Dim iMax As Integer
        
    '**************************************************
    '*  戻り値のセット（Ｔｒｕｅ）                      *
    '**************************************************
    FVSModChkCher = True
    
    '**************************************************
    '*  最大文字列数の検索　　                          *
    '**************************************************
    iMax = FvsModLenb(Nz(sStr, ""))
    
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
                    FVSModChkCher = True
                Case Else
                    '全角文字
                    FVSModChkCher = False
                    Exit Do
            End Select
        Else                     '全角文字のみ許可
            Select Case Asc(swk)
                Case 0 To 255
                    '半角文字
                    FVSModChkCher = False
                    Exit Do
                Case Else
                    '全角文字
                    FVSModChkCher = True
                    iMax = iMax - 1
                    If iMax < 1 Then
                        Exit Do
                    End If
            End Select
        End If
    
    Loop
    
End Function

Public Function GetCurPath() As String
    
    Dim strAppPath As String
    Dim intCount   As Integer
    Dim lngTemp    As Long
    
    strAppPath = ""
    intCount = 1
    Do While 1
        lngTemp = InStr(intCount, CurrentDb.NAME, "\", 1)
        If lngTemp <> 0 Then
            intCount = lngTemp + 1
        Else
            strAppPath = strAppPath & Left(CurrentDb.NAME, intCount - 1)
            Exit Do
        End If
    Loop
    
    GetCurPath = strAppPath

End Function

Public Function DebugPrint(strPrint As String, Optional strFilename As String)

    If strFilename = "" Then
        strFilename = "Default.txt"
    End If
    
    Open GetCurPath & strFilename For Output As #1
    Print #1, strPrint
    Close #1
    
End Function
