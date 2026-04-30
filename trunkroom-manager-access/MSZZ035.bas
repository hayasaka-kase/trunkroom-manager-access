Attribute VB_Name = "MSZZ035"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : コンバート関数
'        PROGRAM_ID      : MSZZ035
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2007/08/14
'        CERATER         : K.ISHZIAKA
'        Ver             : 0.0
'
'        UPDATE          : 2011/03/02
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.1
'                        : UTF8_UrlEncode を追加
'
'        UPDATE          : 2012/02/20
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.2
'                        : UTF8_GetString を修正（復号に失敗してるのはこいつのせいだったかも）
'
'        UPDATE          : 2012/09/14
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.3
'                        : UTF8_GetString を修正（PHP暗号空白対応）
'
'        UPDATE          : 2014/07/23
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.4
'                        : UTF8_GetString を修正 Ver0.3の修正により全て半角だったときに中身が消えてしまっていた
'
'        UPDATE          : 2019/05/08
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.5
'                        : UTF8_GetBytes の戻り値の最後に余分なNULL文字がついていた
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   API宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Declare Function WideCharToMultiByte Lib "kernel32" _
    (ByVal CodePage As Long, _
     ByVal dwFlags As Long, _
     ByVal lpWideCharStr As Long, _
     ByVal cchWideChar As Long, _
     ByRef lpMultiByteStr As Any, _
     ByVal cchMultiByte As Long, _
     ByVal lpDefaultChar As String, _
     ByVal lpUsedDefaultChar As Long) As Long
     
Private Declare Function MultiByteToWideChar Lib "kernel32" _
    (ByVal CodePage As Long, _
     ByVal dwFlags As Long, _
     ByRef lpMultiByteStr As Any, _
     ByVal cchMultiByte As Long, _
     ByVal lpWideCharStr As Long, _
     ByVal cchWideChar As Long) As Long

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const CP_UTF8 = 65001  '0----+----1----+----2----+----3----+----4----+----5----+----6--3
Private Const C_BASE64_TABLE = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Private Sub TEST_UTF8_UrlEncode()
    Debug.Print UTF8_UrlEncode("石阪 キヨシ")
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 文字列をUTF-8でエンコードする（URLに使用できるように）
'       MODULE_ID       : UTF8_UrlEncode
'       CREATE_DATE     : 2011/03/02            K.ISHIZAKA
'       PARAM           : strText               文字列(I)
'       RETURN          : エンコード文字列(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UTF8_UrlEncode(ByVal strText As String) As String
    Dim objSC               As Object
    On Error GoTo ErrorHandler
    
    Set objSC = CreateObject("ScriptControl")
    objSC.Language = "Jscript"
    UTF8_UrlEncode = objSC.CodeObject.encodeURIComponent(strText)
    Set objSC = Nothing
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "UTF8_UrlEncode" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : Unicode文字列をUTF8バイト配列に変換
'       MODULE_ID       : UTF8_GetBytes
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : strText               文字列(I)
'       RETURN          : UTF8バイト配列(Byte())
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UTF8_GetBytes(ByVal strText As String) As Byte()
    Dim bytUtf8()           As Byte
    Dim lngTextLen          As Long
    Dim lngBuffLen          As Long
    Dim lngResult           As Long
    On Error GoTo ErrorHandler

    lngTextLen = Len(strText)
    If lngTextLen = 0 Then                                                      'INSERT 2019/05/08 K.ISHIZAKA
        ReDim bytUtf8(0)                                                        'INSERT 2019/05/08 K.ISHIZAKA
    Else                                                                        'INSERT 2019/05/08 K.ISHIZAKA
        lngBuffLen = lngTextLen * 3 + 1
        ReDim bytUtf8(lngBuffLen)
        lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strText), lngTextLen, bytUtf8(0), lngBuffLen, vbNullString, 0)                                           'WideCharToMultiByte で Unicode → UTF-8 変換
        
'        ReDim Preserve bytUtf8(lngResult)                                      'DELETE 2019/05/08 K.ISHIZAKA
        ReDim Preserve bytUtf8(lngResult - 1)                                   'INSERT 2019/05/08 K.ISHIZAKA
    End If                                                                      'INSERT 2019/05/08 K.ISHIZAKA
    UTF8_GetBytes = bytUtf8
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "UTF8_GetBytes" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : UTF8バイト配列をUnicode文字列に変換
'       MODULE_ID       : UTF8_GetString
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : bytUtf8()             UTF8バイト配列(I)
'       RETURN          : 文字列(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UTF8_GetString(bytUtf8() As Byte) As String
    Dim strText             As String
    Dim lngBuffLen          As Long
    Dim lngResult           As Long
    On Error GoTo ErrorHandler
    
    lngBuffLen = UBound(bytUtf8) + 1
    If lngBuffLen <= 0 Then                                                     'INSERT 2012/09/14 K.ISHIZAKA
        strText = ""                                                            'INSERT 2012/09/14 K.ISHIZAKA
    Else                                                                        'INSERT 2012/09/14 K.ISHIZAKA
        strText = String(lngBuffLen, vbNullChar)
        lngResult = MultiByteToWideChar(CP_UTF8, 0, bytUtf8(0), lngBuffLen, StrPtr(strText), lngBuffLen)

'        strText = Left$(strText, InStr(1, strText, vbNullChar) - 1)            'DELETE 2012/02/20 K.ISHIZAKA
        lngResult = InStr(1, strText, vbNullChar)                               'INSERT START 2012/02/20 K.ISHIZAKA
        If lngResult > 0 Then
            strText = Left$(strText, lngResult - 1)
'        Else                                                                   'DELETE 2014/07/23 K.ISHIZAKA 'INSERT 2012/09/14 K.ISHIZAKA
'            strText = ""                                                       'DELETE 2014/07/23 K.ISHIZAKA 'INSERT 2012/09/14 K.ISHIZAKA
        End If                                                                  'INSERT END   2012/02/20 K.ISHIZAKA
    End If                                                                      'INSERT 2012/09/14 K.ISHIZAKA
    UTF8_GetString = strText
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "UTF8_GetString" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : バイト配列をBase64文字列に変換
'       MODULE_ID       : ToBase64String
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : bstr()                バイト配列(I)
'       RETURN          : Base64文字列(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function ToBase64String(bstr() As Byte) As String
    Dim Bmode               As Integer          'バイナリデータの区切りモード(0-3)
    Dim cnt                 As Long             'バイナリデータのポインタ
    Dim B1                  As Byte             '符号化するバイナリデータ
    Dim B64str              As String
    Dim Blen                As Long
    On Error GoTo ErrorHandler

    Blen = UBound(bstr) + 1
    B64str = ""
    cnt = 0
    Bmode = 0
    
    Do Until Blen <= cnt
        ' Ｂｉｔデータの整形
        B1 = bstr(cnt)
        Select Case Bmode
            Case 0
                B1 = (&HFC And B1) \ 4                  '上位6Bit
            Case 1
                B1 = (&H3 And B1) * 16                  '下位2Bit
                cnt = cnt + 1                           '     +
                If Blen > cnt Then
                    B1 = B1 + (&HF0 And bstr(cnt)) \ 16 '上位4Bit
                End If
            Case 2
                B1 = (&HF And B1) * 4                   '下位4Bit
                cnt = cnt + 1                           '     +
                If Blen > cnt Then
                    B1 = B1 + (&HC0 And bstr(cnt)) \ 64 '上位2Bit
                End If
            Case 3
                B1 = &H3F And B1                        '下位6Bit
                cnt = cnt + 1
        End Select
        
        'Base64文字列作成（符号化）
        B64str = B64str & Mid(C_BASE64_TABLE, B1 + 1, 1)
        
        Bmode = Bmode + 1
        If Bmode > 3 Then Bmode = 0
    Loop
    
    '終端記号"="の付加
    Select Case Bmode
    Case 0                       '何も付け加えない
        B64str = B64str
    Case 1, 2                    '"=="
        B64str = B64str & "=="
    Case 3                       '"="
        B64str = B64str & "="
    End Select
    
    ToBase64String = B64str
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "ToBase64String" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : Base64文字列をバイト配列に変換
'       MODULE_ID       : FromBase64String
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : B64str                Base64文字列(I)
'       RETURN          : バイト配列(Byte())
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function FromBase64String(ByVal B64str As String) As Byte()
    Dim Bmode               As Integer          'バイナリデータの区切りモード(0-3)
    Dim ACnt                As Long             'ASCIIデータのポインタ
    Dim BCnt                As Long             'バイナリデータのポインタ
    Dim RetVal              As Integer
    Dim bstr()              As Byte
    On Error GoTo ErrorHandler
    
    BCnt = 0
    Bmode = 0
    B64str = Replace(B64str, "=", "")
    For ACnt = 0 To Len(B64str) - 1
        'ASCIIデータの数値化(0-63)
        RetVal = InStr(1, C_BASE64_TABLE, Mid$(B64str, ACnt + 1, 1), vbBinaryCompare)
        If RetVal > 0 Then
            Dim B1          As Byte             '作成されたバイナリデータ
            
            B1 = CByte(RetVal - 1)
            'Base64バイナリ変換
            Select Case Bmode
            Case 0
                ReDim Preserve bstr(BCnt)
                bstr(BCnt) = B1 * 4                 '上位6Bit
            Case 1
                bstr(BCnt) = bstr(BCnt) + (B1 \ 16) '下位2Bit
                If ((Len(B64str) - 1) > ACnt) Or ((&HF And B1) > 0) Then
                    BCnt = BCnt + 1                 '     +
                    ReDim Preserve bstr(BCnt)
                    bstr(BCnt) = (&HF And B1) * 16  '上位4Bit
                End If
            Case 2
                bstr(BCnt) = bstr(BCnt) + (B1 \ 4)  '下位4Bit
                If ((Len(B64str) - 1) > ACnt) Or ((&H3 And B1) > 0) Then
                    BCnt = BCnt + 1                 '     +
                    ReDim Preserve bstr(BCnt)
                    bstr(BCnt) = (&H3 And B1) * 64  '上位2Bit
                End If
            Case 3
                bstr(BCnt) = bstr(BCnt) + B1        '上位6Bit
                BCnt = BCnt + 1
            End Select
            Bmode = Bmode + 1
            If Bmode > 3 Then Bmode = 0
        End If
    Next
    
    FromBase64String = bstr
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "FromBase64String" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  ended or program ********************************
