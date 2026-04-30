Attribute VB_Name = "MSZZ070"
'****************************  start of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : 自動鍵解除番号発番
'       PROGRAM_ID      : MSZZ070
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2014/03/05
'       CERATER         : EGL K.MIYAMOTO
'       Ver             : 0.0
'
'       UPDATE          : 2014/04/16
'       UPDATER         : MIYAMOTO
'       Ver             : 0.1
'                       : 引数にコンテナ番号を追加、乱数系列を初期化を追加、不要コメント部削除
'
'       UPDATE          : 2016/05/16
'       UPDATER         : SUGIMURA
'       Ver             : 0.2
'                       : 解除番号制御区分のトリミング処理を追加
'
'       UPDATE          : 2021/06/02
'       UPDATER         : EGL
'       Ver             : 0.3
'                       : QR番号対応
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'==============================================================================*
'
'       MODULE_NAME     : 自動鍵解除番号発番
'       MODULE_ID       : MSZZ070_M00
'       CREATE_DATE     : 2014/03/05            EGL K.MIYAMOTO
'       PARAM           : strBumonCode          部門コード(String)
'                         strYARD               ヤードコード(Long)
'                         strKEY_TYPE           鍵区分(String)
'       RETURN          : 解除番号(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'↓ UPDATE 2014/04/16 MIYAMOTO
'Public Function MSZZ070_M00(ByVal strBumonCode As String, intYARD As Long, strKEY_TYPE As String) As String
Public Function MSZZ070_M00(ByVal strBumonCode As String, intYARD As Long, intCNTA As Long, strKEY_TYPE As String) As String
'↑ UPDATE 2014/04/16 MIYAMOTO

    Dim objAdoDbConnection  As Object
    Dim rsCNTAData          As Object
    Dim rsAKNDData          As Object
    Dim rsData              As Object

    Dim intVal              As Integer
    Dim strVal              As String
    Dim booSW               As Boolean
    Dim strSQL              As String

On Error GoTo ErrorHandler

    booSW = False

'↓ INSERT 2014/04/16 MIYAMOTO
    '乱数系列を初期化
    Randomize
'↑ INSERT 2014/04/16 MIYAMOTO

    Set objAdoDbConnection = MSZZ025.ADODB_Connection(strBumonCode)

    Do Until booSW = True
        '乱数発生
        intVal = Int(Rnd * 9999)
        strVal = Format(intVal, "0000")
        '取得値が自動鍵解除番号発番トラン、及び自動鍵採番除外マスタに存在するか確認
        strSQL = ""
        strSQL = strSQL & "SELECT      AKKNT.AKKNT_KAINO" & vbCrLf
        strSQL = strSQL & "FROM        (" & vbCrLf
        strSQL = strSQL & "                SELECT      AKKNT_YARD" & vbCrLf
        strSQL = strSQL & "                ,           AKKNT_KTYPE" & vbCrLf
        strSQL = strSQL & "                ,           AKKNT_HATUD" & vbCrLf
        strSQL = strSQL & "                ,           AKKNT_KAINO" & vbCrLf
        strSQL = strSQL & "                ,           ROW_NUMBER() OVER(ORDER BY AKKNT_HATUD DESC) AS SEQ" & vbCrLf
        strSQL = strSQL & "                FROM        AKKN_TRAN" & vbCrLf      '自動鍵解除番号発番トラン
        strSQL = strSQL & "                WHERE       AKKNT_YARD = " & intYARD & vbCrLf
        strSQL = strSQL & "                AND         AKKNT_CTNO = " & intCNTA & vbCrLf                'INSERT 2014/04/16 MIYAMOTO
        strSQL = strSQL & "                AND         AKKNT_KTYPE = '" & strKEY_TYPE & "'" & vbCrLf
        strSQL = strSQL & "                AND         AKKNT_STYPE = '4'" & vbCrLf                      'INSERT 2021/06/02 EGL
        strSQL = strSQL & "            ) AKKNT" & vbCrLf
        strSQL = strSQL & "WHERE       AKKNT.SEQ <= 5" & vbCrLf
        strSQL = strSQL & "AND         AKKNT.AKKNT_KAINO = '" & strVal & "'" & vbCrLf
        strSQL = strSQL & "UNION" & vbCrLf
        strSQL = strSQL & "SELECT      AKNDM_KEYNO" & vbCrLf
        strSQL = strSQL & "FROM        AKND_MAST" & vbCrLf                      '自動鍵採番除外マスタ
        strSQL = strSQL & "WHERE       AKNDM_KEYCD = '" & strKEY_TYPE & "'" & vbCrLf
        strSQL = strSQL & "AND         AKNDM_KEYNO = '" & strVal & "'" & vbCrLf
        strSQL = strSQL & ";" & vbCrLf

        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)

        If rsData.EOF Then              '存在しない場合
            booSW = True
        End If
        rsData.Close
        Set rsData = Nothing
    Loop

    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

    MSZZ070_M00 = strVal
Exit Function

ErrorHandler:
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If
    Call Err.Raise(Err.Number, "MSZZ070_M00" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 自動鍵解除番号発番(6桁用)
'       MODULE_ID       : MSZZ070_M01
'       CREATE_DATE     : 2021/06/02            EGL
'       PARAM           : strBumonCode          部門コード(String)
'                         strYARD               ヤードコード(Long)
'                         strKEY_TYPE           鍵区分(String)
'       RETURN          : 解除番号(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ070_M01(ByVal strBumonCode As String, strKEY_TYPE As String) As String

    Dim objAdoDbConnection  As Object
    Dim rsCNTAData          As Object
    Dim rsAKNDData          As Object
    Dim rsData              As Object

    Dim intVal              As Long
    Dim strVal              As String
    Dim booSW               As Boolean
    Dim strSQL              As String

On Error GoTo ErrorHandler

    booSW = False

    '乱数系列を初期化
    Randomize

    Set objAdoDbConnection = MSZZ025.ADODB_Connection(strBumonCode)

    Do Until booSW = True
        '乱数発生
        intVal = CLng(Rnd * 999999)
        strVal = Format(intVal, "000000")

        '取得値が自動鍵解除番号発番トラン、及び自動鍵採番除外マスタに存在するか確認
        strSQL = ""
        strSQL = strSQL & "SELECT      AKKNT.AKKNT_KAINO" & vbCrLf
        strSQL = strSQL & "FROM        (" & vbCrLf
        strSQL = strSQL & "                SELECT      AKKNT_YARD" & vbCrLf
        strSQL = strSQL & "                ,           AKKNT_KTYPE" & vbCrLf
        strSQL = strSQL & "                ,           AKKNT_HATUD" & vbCrLf
        strSQL = strSQL & "                ,           AKKNT_KAINO" & vbCrLf
        strSQL = strSQL & "                ,           ROW_NUMBER() OVER(ORDER BY AKKNT_HATUD DESC) AS SEQ" & vbCrLf
        strSQL = strSQL & "                FROM        AKKN_TRAN" & vbCrLf      '自動鍵解除番号発番トラン
        strSQL = strSQL & "                WHERE       AKKNT_STYPE = '6'" & vbCrLf
        strSQL = strSQL & "                AND         AKKNT_KTYPE = '" & strKEY_TYPE & "'" & vbCrLf
        strSQL = strSQL & "            ) AKKNT" & vbCrLf
        strSQL = strSQL & "WHERE       AKKNT.SEQ <= 3" & vbCrLf
        strSQL = strSQL & "AND         AKKNT.AKKNT_KAINO = '" & strVal & "'" & vbCrLf
        strSQL = strSQL & "UNION" & vbCrLf
        strSQL = strSQL & "SELECT      AKNDM_KEYNO" & vbCrLf
        strSQL = strSQL & "FROM        AKND_MAST" & vbCrLf                      '自動鍵採番除外マスタ
        strSQL = strSQL & "WHERE       AKNDM_KEYCD = '" & strKEY_TYPE & "'" & vbCrLf
        strSQL = strSQL & "AND         AKNDM_KEYNO = '" & strVal & "'" & vbCrLf
        strSQL = strSQL & ";" & vbCrLf

        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)

        If rsData.EOF Then              '存在しない場合
            booSW = True
        End If
        rsData.Close
        Set rsData = Nothing
    Loop

    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

    MSZZ070_M01 = strVal
Exit Function

ErrorHandler:
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If
    Call Err.Raise(Err.Number, "MSZZ070_M01" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function



'↓INSERT 2016/05/16 SUGIMURA
'==============================================================================*
'
'        MODULE_NAME      :小数の解除番号制御区分のトリミング
'        MODULE_ID        :MSZZ070_KaiKey
'        IN               :第1引数:地方ヤードチェックボックス
'                         :第2引数:解除番号制御区分
'        OUT              :解除番号制御区分
'        CREATE_DATE      :2016/05/16
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ070_KaiKey(intChkYard As Integer, varNameVal As Variant) As String

    Dim strVal      As String
    Dim tmp As Variant

    'varNameValに値が入っていない場合、"0"をリターン
    If IsNull(varNameVal) = True Or varNameVal = "" Then
        MSZZ070_KaiKey = "0"
        Exit Function
    End If

    strVal = Format(varNameVal, "0.0#")
    'varNameValに小数点が含まれているか判断する
    '地方ヤードの場合（-1）、小数部のみ取得し、都市ヤードの場合は整数部を取得する
    If intChkYard = -1 Then
        tmp = Split(strVal, ".")
        MSZZ070_KaiKey = tmp(1)
    Else
        tmp = Split(strVal, ".")
        MSZZ070_KaiKey = tmp(0)
    End If
End Function
'↑INSERT 2016/05/16 SUGIMURA


'****************************  ended of program ********************************
