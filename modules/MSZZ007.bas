Attribute VB_Name = "MSZZ007"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME :
'
'        PROGRAM_NAME    :
'        PROGRAM_ID      : MSZZ007
'        PROGRAM_KBN     :
'
'        CREATE          : 2003/11/28
'        CERATER         : S.SHIBAZAKI
'        Ver             : 0.0
'
'        UPDATE          : 2003/12/19
'        UPDATER         : S.SHIBAZAKI
'        Ver             :
'        DETAIL          : 省略可能引数として部門コード追加
'
'        UPDATE          : 2004/10/29
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.1
'        DETAIL          : ActiveX Data Object 対応
'
'==============================================================================*
Option Explicit

Private dbsMSZZ007      As Database

Const CONNECT_STRING_HEAD = "ODBC;"
Const CONNECT_STRING_HEAD_ADO = "Provider=SQLOLEDB;"                            '20041029 ADD K.ISHIZAKA

Const ODBC_DATA_SOURCE_NAME = "ODBC_DATA_SOURCE_NAME,DSN"
Const ODBC_SERVER_NAME = "ODBC_SERVER_NAME,SERVER"
Const ODBC_USER_ID = "ODBC_USER_ID,UID"
Const ODBC_PASSWORD = "ODBC_PASSWORD,PWD"
Const ODBC_DATABASE_NAME = "ODBC_DATABASE_NAME,DATABASE"

Const NULL_PARTS = "#NULL#"
'==============================================================================*
'
'        MODULE_NAME      :ODBC接続文字列取得
'        MODULE_ID        :MSZZ007_M00
'        CREATE_DATE      :2003/11/28
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Function MSZZ007_M00() As String
Function MSZZ007_M00(Optional strBumon As String) As String             'UPDATE 2003/12/19 SHIBAZAKI

    Dim strConnectString        As String
    Dim varParts                As Variant
    Dim varArrList              As String
    Dim varArray                As Variant
    Dim intCount                As Integer
    Dim intPos                  As Integer
    Dim blnOpenDatabase         As Boolean
    Dim strKeyTail              As String
    
    blnOpenDatabase = False
    
    'ADD 2003/12/19 SHIBAZAKI
    '部門コードの引数が指定されたら、SETU_TABL検索キーを下記のようにする
    '検索キー & "_" & 部門コード
    If IsNull(strBumon) Or strBumon = "" Then
        strKeyTail = ""
    Else
        strKeyTail = "_" & strBumon
    End If
    
    varArray = Array(ODBC_DATA_SOURCE_NAME, ODBC_SERVER_NAME, ODBC_USER_ID, ODBC_PASSWORD, ODBC_DATABASE_NAME)
    
    'カレントデータベース接続
    Set dbsMSZZ007 = CurrentDb
    blnOpenDatabase = True
    
    strConnectString = ""
    
    '配列には、SETU_TABL検索キーと接続文字列を作成する文字がカンマ区切りで格納。
    '例えば、"ODBC_DATA_SOURCE_NAME,DSN"ならば、
    '"ODBC_DATA_SOURCE_NAME"でSETU_TABLを検索し、"DSN"と取得文字列を"="で結合
    '取得文字列が"KASE_DB"ならば、"DSN=KASE_DB"となる。
    'ただし、部門コードの引数が指定されたら、SETU_TABL検索キーを下記のようにする
    '検索キー & "_" & 部門コード
    For intCount = 0 To UBound(varArray)
        If strConnectString <> "" Then
            strConnectString = strConnectString & ";"
        End If
        intPos = InStr(varArray(intCount), ",")
        'varParts = fncGetSetuTabl(Left(varArray(intCount), intPos - 1))
        varParts = fncGetSetuTabl(Left(varArray(intCount), intPos - 1) & strKeyTail)    'UPDATE 2003/12/19 SHIBAZAKI
        If varParts = "" Then
            strConnectString = ""
            GoTo Exit_rtn
        End If
        If varParts = NULL_PARTS Then
            varParts = ""
        End If
        strConnectString = strConnectString & Mid(varArray(intCount), intPos + 1) & "=" & varParts
    Next

    strConnectString = CONNECT_STRING_HEAD & strConnectString
    
Exit_rtn:
    If blnOpenDatabase Then
        dbsMSZZ007.Close
        Set dbsMSZZ007 = Nothing
    End If
    
    MSZZ007_M00 = strConnectString
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :ADO接続文字列取得
'        MODULE_ID        :MSZZ007_M10
'        CREATE_DATE      :2004/10/29
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZ007_M10(Optional strBumon As String) As String

    Dim strConnectString        As String
    Dim varParts                As Variant
    Dim varArrList              As String
    Dim varArray                As Variant
    Dim intCount                As Integer
    Dim intPos                  As Integer
    Dim blnOpenDatabase         As Boolean
    Dim strKeyTail              As String
    
    blnOpenDatabase = False
    
    '部門コードの引数が指定されたら、SETU_TABL検索キーを下記のようにする
    '検索キー & "_" & 部門コード
    If IsNull(strBumon) Or strBumon = "" Then
        strKeyTail = ""
    Else
        strKeyTail = "_" & strBumon
    End If
    
    varArray = Array(ODBC_SERVER_NAME, ODBC_USER_ID, ODBC_PASSWORD, ODBC_DATABASE_NAME)
    
    'カレントデータベース接続
    Set dbsMSZZ007 = CurrentDb
    blnOpenDatabase = True
    
    strConnectString = ""
    
    '配列には、SETU_TABL検索キーと接続文字列を作成する文字がカンマ区切りで格納。
    '例えば、"ODBC_DATA_SOURCE_NAME,DSN"ならば、
    '"ODBC_DATA_SOURCE_NAME"でSETU_TABLを検索し、"DSN"と取得文字列を"="で結合
    '取得文字列が"KASE_DB"ならば、"DSN=KASE_DB"となる。
    'ただし、部門コードの引数が指定されたら、SETU_TABL検索キーを下記のようにする
    '検索キー & "_" & 部門コード
    For intCount = 0 To UBound(varArray)
        If strConnectString <> "" Then
            strConnectString = strConnectString & ";"
        End If
        intPos = InStr(varArray(intCount), ",")
        varParts = fncGetSetuTabl(Left(varArray(intCount), intPos - 1) & strKeyTail)
        If varParts = "" Then
            strConnectString = ""
            GoTo Exit_rtn
        End If
        If varParts = NULL_PARTS Then
            varParts = ""
        End If
        strConnectString = strConnectString & Mid(varArray(intCount), intPos + 1) & "=" & varParts
    Next

    strConnectString = CONNECT_STRING_HEAD_ADO & strConnectString
    
Exit_rtn:
    If blnOpenDatabase Then
        dbsMSZZ007.Close
        Set dbsMSZZ007 = Nothing
    End If
    
    MSZZ007_M10 = strConnectString
End Function

Function fncGetSetuTabl(varKey As Variant) As String

    Dim strRet      As String
    Dim varParts    As Variant
    
    On Error GoTo Err_End
    strRet = ""
    
    varParts = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = '" & varKey & "'")
    If IsNull(varParts) Then
        strRet = ""
    Else
        strRet = varParts
    End If
    
    GoTo Func_End
    
Err_End:
    MsgBox Err.Description
    strRet = ""

Func_End:
    fncGetSetuTabl = strRet

End Function
'****************************  ended or program ********************************

