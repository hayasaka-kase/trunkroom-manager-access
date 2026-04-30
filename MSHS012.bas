Attribute VB_Name = "MSHS012"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME :
'
'        PROGRAM_NAME    : テーブルリンク(KOMS)
'        PROGRAM_ID      : MSHS012
'        PROGRAM_KBN     :
'
'        CREATE          : 2005/01/08
'        CERATER         : N.MIURA
'        Ver             : 0.0
'
'        UPDATE          : 2005/06/02
'        UPDATER         : N.MIURA
'        Ver             : 0.1
'                        : 予約受付トラン追加
'
'        UPDATE          : 2008/07/30
'        UPDATER         : S.SHIBAZAKI
'        Ver             : 0.2
'                        : MSHS012_M10追加
'
'        UPDATE          : 2008/08/01
'        UPDATER         : N.Miura
'        Ver             : 0.3
'                        : 1.リンクテーブルの追加
'                            CBOX_MAST
'                            CBRA_RECD
'                            CBRA_TRAN
'                            CBYS_RECD
'                            CBYS_TRAN
'                            INTR_TRAN
'                            RCPT_TRAN
'                            SCRD_MAST
'                        : 2.リンクテーブルの削除
'                            KYOU_TRAN
'
'        UPDATE          : 2018/07/21
'        UPDATER         : N.IMAI
'        Ver             : 0.4
'                        : 1.リンクテーブルの追加
'                            KKTH_MAST
'                            KKTP_MAST
'                            KTPS_TRAN
'                            PVS500_LOG
'
'        UPDATE          : 2020/09/01
'        UPDATER         : Y.WADA
'        Ver             : 0.5
'                        : 1.リンクテーブルの追加
'                            CTRL_TABL
'
'        UPDATE          : 2020/09/28
'        UPDATER         : N.IMAI
'        Ver             : 0.6
'                        : 1.CTRL_TABLは「dbo_」を付加しない
'
'        UPDATE          : 2020/09/30
'        UPDATER         : N.IMAI
'        Ver             : 0.7
'                        : 1.CTRL_TABLは「dbo_」を付加するように戻す
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "MSHS012"
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :MSHS012_M00
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSHS012_M00(strDBID)
    On Error GoTo ERR_MSHS012_M00
    
    Dim SV_ID           As String
    
    Dim intPos          As Integer
    Dim intPosSave      As Integer
    Dim intLen          As Integer
    
    Dim strDNSNN        As String
    Dim strSERVN        As String
    Dim strDATBN        As String
    
    SV_ID = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = '" & strDBID & "'"), "")
    
    intLen = Len(SV_ID)
    
    intPos = InStr(1, SV_ID, "\")
    strDNSNN = Mid$(SV_ID, 1, intPos - 1)
    intPosSave = intPos
        
    intPos = InStr(intPosSave + 1, SV_ID, "\")
    strSERVN = Mid$(SV_ID, intPosSave + 1, intPos - (intPosSave + 1))
    intPosSave = intPos
        
    intPos = InStr(intPosSave + 1, SV_ID, "\")
    strDATBN = Mid$(SV_ID, intPosSave + 1, intLen - (intPosSave))
    
    Call subLink(strDNSNN, strSERVN, strDATBN, "", "")                  'INSERT 2008/07/30 SHIBAZAKI
    
    '↓DELETE 2008/07/30 SHIBAZAKI
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "CARG_DELE")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "CARG_FILE")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "CNTA_MAST")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "CNTA_RFILE")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "CONT_MAST")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "EXPE_FILE")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "FKS220_WORK_01")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "FKS220_WORK_02")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "JARG_FILE")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "JINU_MAST")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "JINU_RFILE")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "KYOU_TRAN")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "NAME_MAST")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "NYAR_MAST")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "PAID_FILE")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "PAIH_FILE")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "PRIC_TABL")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "REQU_DELE")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "REQU_FILE")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "RKS140_WORK")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "RKS170_WORK")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "RKS170_WORK_01")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "RKS170_WORK_02")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "RKS170_WORK_03")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "RKS220_WORK")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "TKS100_WORK")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "TKS110_WORK")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "TKS180_WORK")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "TKS180_WORK_01")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "TKS180_WORK_02")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "TKS190_WORK")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "TKS270_WORK")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "TKS300_WORK")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "TKS360_WORK")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "TORI_MAST")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "USER_MAST")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "YARD_MAST")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "YARD_RFILE")
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "YOUK_TRAN")            'INSERT 20050602 N.MIURA
'    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, "ZAIK_FILE")
    '↑DELETE 2008/07/30 SHIBAZAKI
    
    Exit Function

ERR_MSHS012_M00:
    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :MSHS012_DBO_LINK
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2008/07/30 SHIBAZAKI strUSRIDとstrPASWDを追加
Function MSHS012_DBO_LINK(strDNSNN As String, _
                          strSERVN As String, _
                          strDATBN As String, _
                          strUSRID As String, _
                          strPASWD As String, _
                          strTABLID As String)
    
    Dim ret             As Integer
    Dim WK_ERRON        As String
    
    'DELETE 2020/09/28 N.IMAI Start
'    ret = MSZZ002_M00("dbo_" & strTABLID, WK_ERRON)
'    'DELETE 2008/07/30 SHIBAZAKI
'    'Ret = MSZZ002_M20(strDNSNN, strSERVN, strDATBN, "dbo." & strTABLID, "dbo_" & strTABLID, WK_ERRON)
'    'INSERT 2008/07/30 SHIBAZAKI
'    ret = MSZZ002_M20(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "dbo." & strTABLID, "dbo_" & strTABLID, WK_ERRON)
    'DELETE 2020/09/28 N.IMAI End
    
    'INSERT 2020/09/28 N.IMAI Start
    'If strTABLID = "CTRL_TABL" Then
    '    ret = MSZZ002_M00(strTABLID, WK_ERRON)
    '    ret = MSZZ002_M20(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "dbo." & strTABLID, strTABLID, WK_ERRON)
    'Else
        ret = MSZZ002_M00("dbo_" & strTABLID, WK_ERRON)
        ret = MSZZ002_M20(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD, "dbo." & strTABLID, "dbo_" & strTABLID, WK_ERRON)
    'End If
    'INSERT 2020/09/28 N.IMAI End

End Function
'==============================================================================*
'
'        MODULE_NAME      :部門コードで実行する
'        MODULE_ID        :MSHS012_M10
'        CREATE_DATE      :2008/07/30 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSHS012_M10(strParamBumoc As String)
    On Error GoTo ErrHndl
    
    Dim strDNSNN        As String
    Dim strSERVN        As String
    Dim strDATBN        As String
    Dim strUSRID        As String
    Dim strPASWD        As String
    Dim strBUMOC        As String
    
    strBUMOC = IIf(strParamBumoc = "", "", "_" & strParamBumoc)
    
    strDNSNN = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATA_SOURCE_NAME" & strBUMOC & "'")
    strSERVN = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_SERVER_NAME" & strBUMOC & "'")
    strDATBN = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATABASE_NAME" & strBUMOC & "'")
    strUSRID = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_USER_ID" & strBUMOC & "'")
    strPASWD = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_PASSWORD" & strBUMOC & "'")
    
    Call subLink(strDNSNN, strSERVN, strDATBN, strUSRID, strPASWD)

ErrHndl:
    Exit Function
End Function

'==============================================================================*
'
'        MODULE_NAME      :各テーブルをリンクする
'        MODULE_ID        :subLink
'        CREATE_DATE      :2008/07/30 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subLink(strDNSNN As String, strSERVN As String, strDATBN As String, strUid As String, strPwd As String)
    On Error GoTo ErrHndl
    
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "CARG_DELE")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "CARG_FILE")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "CBOX_MAST") 'INSERT 20080808 N.MIURA
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "CBRA_RECD") 'INSERT 20080808 N.MIURA
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "CBRA_TRAN") 'INSERT 20080808 N.MIURA
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "CBYS_RECD") 'INSERT 20080808 N.MIURA
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "CBYS_TRAN") 'INSERT 20080808 N.MIURA
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "CNTA_MAST")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "CNTA_RFILE")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "CONT_MAST")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "EXPE_FILE")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "FKS220_WORK_01")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "FKS220_WORK_02")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "INTR_TRAN") 'INSERT 20080808 N.MIURA
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "JARG_FILE")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "JINU_MAST")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "JINU_RFILE")
    'Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUID, strPWD, "KYOU_TRAN")'DELETE 20080801 N.MIURA
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "NAME_MAST")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "NYAR_MAST")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "PAID_FILE")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "PAIH_FILE")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "PRIC_TABL")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "RCPT_TRAN") 'INSERT 20080808 N.MIURA
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "REQU_DELE")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "REQU_FILE")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "RKS140_WORK")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "RKS170_WORK")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "RKS170_WORK_01")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "RKS170_WORK_02")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "RKS170_WORK_03")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "RKS220_WORK")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "SCRD_MAST") 'INSERT 20080808 N.MIURA
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "TKS100_WORK")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "TKS110_WORK")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "TKS180_WORK")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "TKS180_WORK_01")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "TKS180_WORK_02")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "TKS190_WORK")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "TKS270_WORK")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "TKS300_WORK")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "TKS360_WORK")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "TORI_MAST")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "USER_MAST")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "YARD_MAST")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "YARD_RFILE")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "YOUK_TRAN")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "ZAIK_FILE")
    'insert 2018/07/21 N.IMAI Start
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "KKTH_MAST")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "KKTP_MAST")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "KTPS_TRAN")
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "PVS500_LOG")
    'insert 2018/07/21 N.IMAI End
    Call MSHS012_DBO_LINK(strDNSNN, strSERVN, strDATBN, strUid, strPwd, "CTRL_TABL")    'INSERT 2020/09/01 Y.WADA

ErrHndl:
    Exit Sub

End Sub
'****************************  ended or program ********************************

