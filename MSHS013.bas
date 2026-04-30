Attribute VB_Name = "MSHS013"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME :
'
'        PROGRAM_NAME    : KAGTOSテーブルリンク(KOMS.MDB用)
'        PROGRAM_ID      : MSHS013
'        PROGRAM_KBN     :
'
'        CREATE          : 2008/08/01
'        CERATER         : N.MIURA
'        Ver             : 0.0
'
'        CREATE          : 2009/04/01
'        CERATER         : hirano
'        Ver             : 0.1
'
'        UPDATE          : 2019/05/14
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.2
'                        : KOUZ_MAST_Jを追加
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "MSHS013"
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :MSHS013_M00
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSHS013_M00()
    On Error GoTo ERR_MSHS013_M00
    
    Dim SV_ID           As String
    
    Dim intPos          As Integer
    Dim intPosSave      As Integer
    Dim intLen          As Integer
    
    Dim strDNSNN        As String
    Dim strSERVN        As String
    Dim strDATBN        As String
    
    SV_ID = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'KASE_DB'"), "")
    
    intLen = Len(SV_ID)
    
    intPos = InStr(1, SV_ID, "\")
    strDNSNN = Mid$(SV_ID, 1, intPos - 1)
    intPosSave = intPos
        
    intPos = InStr(intPosSave + 1, SV_ID, "\")
    strSERVN = Mid$(SV_ID, intPosSave + 1, intPos - (intPosSave + 1))
    intPosSave = intPos
        
    intPos = InStr(intPosSave + 1, SV_ID, "\")
    strDATBN = Mid$(SV_ID, intPosSave + 1, intLen - (intPosSave))
    
    Call subKaseDbLink(strDNSNN, strSERVN, strDATBN)
    
    Exit Function

ERR_MSHS013_M00:
    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :加瀬DB LINK
'        MODULE_ID        :subKaseDbLink
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKaseDbLink(strDNSNN As String, _
                          strSERVN As String, _
                          strDATBN As String)
    
    Dim strUSRID        As String
    Dim strPASWD        As String
    
    strUSRID = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_USER_ID'")
    strPASWD = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_PASSWORD'")
    
    Call MSHS013_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "ADDR_TABL")
    Call MSHS013_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "BANK_TABL")
    Call MSHS013_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "BUMO_MAST")
    Call MSHS013_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "CAMP_MAST")
    Call MSHS013_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "CODE_TABL")
    Call MSHS013_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "KOUZ_MAST")
    Call MSHS013_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "KOUZ_MAST_J") 'INSERT 2019/05/14 K.ISHIZAKA
    Call MSHS013_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "NYKO_MAST")
    Call MSHS013_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "TANT_MAST")
    Call MSHS013_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "SHIR_MAST")    '2009/04/01 Add hirano
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :MSHS013_DBO_LINK
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSHS013_DBO_LINK(strDNSNN As String, _
                          strSERVN As String, _
                          strDATBN As String, _
                          strUSRID As String, _
                          strPASWD As String, _
                          strTABLID As String)
    
    Dim ret             As Integer
    Dim WK_ERRON        As String
    
    ret = MSZZ002_M00(strTABLID, WK_ERRON)
    ret = MSZZ002_M20(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "dbo." & strTABLID, strTABLID, WK_ERRON)

End Function
'****************************  ended or program ********************************

