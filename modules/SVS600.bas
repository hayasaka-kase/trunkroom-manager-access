Attribute VB_Name = "SVS600"
Option Compare Database

'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　契約入力用共通関数
'   プログラムＩＤ　：　契約入力用関数
'   作　成　日　　　：  2005/07/13
'   作　成　者　　　：  T.SUZUKI
'**********************************************

Public pblnChangeFlg As Boolean

'==============================================================================*
'
'        MODULE_NAME      :コンテナ契約ファイル問い合わせ
'        MODULE_ID        :pfncDSP_CARG_FILE
'        引       数      ：strYARD_CODE ヤードコード
'                           strCNTA_CODE コンテナ番号
'                           strUSER_CODE 顧客コード
'        CREATE_DATE      :2005/06/17
'
'==============================================================================*
Public Function pfncDSP_CARG_FILE(strYARD_CODE As String, _
                                  strCNTA_CODE As String, _
                                  strUSER_CODE As String) As Boolean

    On Error GoTo err_rtn

    Dim objDb  As Database
    Dim objRs  As Recordset
    Dim strSQL As String

    pfncDSP_CARG_FILE = False

    Set objDb = CurrentDb

    doCmd.Hourglass True
    doCmd.SetWarnings False

    '存在チェック
    strSQL = ""
    strSQL = strSQL & "SELECT COUNT(*) "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "dbo_CARG_FILE "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "(      CARG_YCODE  = " & strYARD_CODE          ' ヤードコード
    strSQL = strSQL & " AND   CARG_NO     = " & strCNTA_CODE          ' コンテナ番号
    strSQL = strSQL & " AND   CARG_UCODE  = " & strUSER_CODE & " )"   ' 顧客コード
    strSQL = strSQL & " AND ((CARG_STDATE >  CDATE(#" & DATE & "#) "
    strSQL = strSQL & " AND ( CARG_KYDATE IS NULL ) )"
    strSQL = strSQL & " OR  ( CARG_STDATE <= CDATE(#" & DATE & "#)"
    strSQL = strSQL & " AND ( CARG_EDDATE >= CDATE(#" & DATE & "#) OR CARG_CONTNA = 0 ) "
    strSQL = strSQL & " AND ( CARG_KYDATE >= CDATE(#" & DATE & "#) OR CARG_KYDATE IS NULL OR CARG_CONTNA = 0 )))"
    Set objRs = objDb.OpenRecordset(strSQL, dbOpenSnapshot)

    If objRs.Fields(0).VALUE = 0 Then
        GoTo Exit_rtn
    End If
    pfncDSP_CARG_FILE = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set objRs = Nothing
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function
