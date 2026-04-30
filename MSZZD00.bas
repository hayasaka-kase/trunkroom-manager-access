Attribute VB_Name = "MSZZD00"
'****************************  strat of program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : DLOOK_UP
'        PROGRAM_ID      : MSZZD00
'        PROGRAM_KBN     :
'
'        CREATE          : 2003/04/15
'        CERATER         : N.MIURA
'        Ver             : 0.0
'
'        UPDATE          : 2004/04/12
'        UPDATER         : N.MIURA
'        Ver             : 0.1
'                        : 業者名称・仕入カナ
'
'        UPDATE          : 2004/06/07
'        UPDATER         : N.MIURA
'        Ver             : 0.2
'                        : 次回更新区分名称
'
'        UPDATE          : 2005/07/06
'        UPDATER         : K.KINEBUCHI
'        Ver             : 0.3
'                        : 賃料改定名称（車庫証明・賃料改定対応）
'
'        UPDATE          : 2008/12/19
'        UPDATER         : S.SHIBAZAKI
'        Ver             : 0.4
'                        : INTI_FILE
'
'        UPDATE          : 2011/02/17
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.5
'                        : MSZZD00_RECDB_Collection を追加
'                        : INTI_FILE 用で INTIF_RECFB をキーに持つ INTIF_RECDB のコレクション
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "MSZZD00"
'
'==============================================================================*
'
'        MODULE_NAME      :部門名称
'        MODULE_ID        :MSZZD00_BUMON
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_BUMON(MSZZD00_BUMOC As String) As String
'    On Error GoTo ERR_MSZZD00_BUMON
    
    MSZZD00_BUMON = Nz(DLookup("BUMOM_BUMON", "BUMO_MAST", _
                               "BUMOM_BUMOC = " & "'" & MSZZD00_BUMOC & "'"))
    
'    Exit Function
'ERR_MSZZD00_BUMON:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :カンパニー名称
'        MODULE_ID        :MSZZD00_CAMPN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_CAMPN(MSZZD00_CAMPC As String) As String
'    On Error GoTo ERR_MSZZD00_CAMPN
    
    MSZZD00_CAMPN = Nz(DLookup("CAMPM_CAMPN", "CAMP_MAST", _
                               "CAMPM_CAMPC = " & "'" & MSZZD00_CAMPC & "'"))
    
'    Exit Function
'ERR_MSZZD00_CAMPN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :担当者名
'        MODULE_ID        :MSZZD00_TANTN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_TANTN(MSZZD00_BUMOC As String, MSZZD00_TANTC As String) As String
'    On Error GoTo ERR_MSZZD00_TANTN
    
    MSZZD00_TANTN = Nz(DLookup("TANTM_TANTN", "TANT_MAST", _
                               "TANTM_BUMOC = " & "'" & MSZZD00_BUMOC & "'" & " AND " & _
                               "TANTM_TANTC = " & "'" & MSZZD00_TANTC & "'"))
    
'    Exit Function
'ERR_MSZZD00_TANTN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :仕入先名称
'        MODULE_ID        :MSZZD00_SHIRN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_SHIRN(MSZZD00_BUMOC As String, MSZZD00_SHIRC As String) As String
'    On Error GoTo ERR_MSZZD00_SHIRN
    
    MSZZD00_SHIRN = Nz(DLookup("SHIRM_SHIRN", "SHIR_MAST", _
                               "SHIRM_BUMOC = " & "'" & MSZZD00_BUMOC & "'" & " AND " & _
                               "SHIRM_SHIRC = " & "'" & MSZZD00_SHIRC & "'"))

'    Exit Function
'ERR_MSZZD00_SHIRN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :仕入先カナ
'        MODULE_ID        :MSZZD00_SHIRF
'        CREATE_DATE      :2004/04/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_SHIRF(MSZZD00_BUMOC As String, MSZZD00_SHIRC As String) As String
'    On Error GoTo ERR_MSZZD00_SHIRN
    
    MSZZD00_SHIRF = Nz(DLookup("SHIRM_SHIRF", "SHIR_MAST", _
                               "SHIRM_BUMOC = " & "'" & MSZZD00_BUMOC & "'" & " AND " & _
                               "SHIRM_SHIRC = " & "'" & MSZZD00_SHIRC & "'"))

'    Exit Function
'ERR_MSZZD00_SHIRN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :物件本体名称
'        MODULE_ID        :MSZZD00_BUKHN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_BUKHN(MSZZD00_BUKHC As String) As String
'    On Error GoTo ERR_MSZZD00_BUKHN
    
    MSZZD00_BUKHN = Nz(DLookup("BUKEM_BUKEN", "BUKE_MAST", _
                               "BUKEM_BUKHC = " & "'" & MSZZD00_BUKHC & "'"))

'    Exit Function

'ERR_MSZZD00_BUKHN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :商品名称
'        MODULE_ID        :MSZZD00_SYOHN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_SYOHN(MSZZD00_BUMOC As String, _
                       MSZZD00_SYO1C As String, _
                       MSZZD00_SYO2C As String) As String
'    On Error GoTo ERR_MSZZD00_SYOHN
    
    MSZZD00_SYOHN = Nz(DLookup("SYOHM_SYOHN", "SYOH_MAST", _
                               "SYOHM_BUMOC = " & "'" & MSZZD00_BUMOC & "'" & " AND " & _
                               "SYOHM_SYO1C = " & "'" & MSZZD00_SYO1C & "'" & " AND " & _
                               "SYOHM_SYO2C = " & "'" & MSZZD00_SYO2C & "'"))

'    Exit Function
'ERR_MSZZD00_SYOHN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :金融機関名
'        MODULE_ID        :MSZZD00_KINYN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_KINYN(MSZZD00_KINYC As String) As String
'    On Error GoTo ERR_MSZZD00_KINYN
    
    MSZZD00_KINYN = Nz(DLookup("BANKT_KINYN", "BANK_TABL", _
                            "BANKT_KINYC = " & "'" & MSZZD00_KINYC & "'"))
    
'    Exit Function
'ERR_MSZZD00_KINYN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :支店名
'        MODULE_ID        :MSZZD00_SHITN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_SHITN(MSZZD00_KINYC As String _
                     , MSZZD00_SHITC As String) As String
'    On Error GoTo ERR_MSZZD00_SHITN
    
    
    MSZZD00_SHITN = Nz(DLookup("BANKT_SHITN", "BANK_TABL", _
                            "BANKT_KINYC = " & "'" & MSZZD00_KINYC & "'" & " AND " & _
                            "BANKT_SHITC = " & "'" & MSZZD00_SHITC & "'"))
'    Exit Function
'ERR_MSZZD00_SHITN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :郵便番号
'        MODULE_ID        :MSZZD00_YUBIB
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_YUBIB(MSZZD00_ADRRN As String) As String
'    On Error GoTo ERR_MSZZD00_YUBIB
    
    MSZZD00_YUBIB = Nz(DLookup("MID(ADDRT_NYUBB,1,3)&'-'&MID(ADDRT_NYUBB,4,4)", "ADDR_TABL", _
                                "ADDRT_TODON & ADDRT_SITSN & ADDRT_CHINN = " & "'" & MSZZD00_ADRRN & "'"))
    
'    Exit Function
'ERR_MSZZD00_YUBIB:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :管理区分名称
'        MODULE_ID        :MSZZD00_KANRN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_KANRN(MSZZD00_KANRI As Integer) As String
'    On Error GoTo ERR_MSZZD00_KANRN
    
    MSZZD00_KANRN = Nz(DLookup("CODET_NAMEN", "CODE_TABL", _
                               "CODET_SIKBC = '221' " & " AND " & _
                               "CODET_CODEC = " & MSZZD00_KANRI))
'    Exit Function
'ERR_MSZZD00_KANRN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :課税区分名称
'        MODULE_ID        :MSZZD00_KAZEN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_KAZEN(MSZZD00_KANRI As Integer) As String
'    On Error GoTo ERR_MSZZD00_KAZEN
    
    MSZZD00_KAZEN = Nz(DLookup("CODET_NAMEN", "CODE_TABL", _
                               "CODET_SIKBC = '204' " & " AND " & _
                               "CODET_CODEC = " & MSZZD00_KANRI))
'    Exit Function
'ERR_MSZZD00_KAZEN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :預金種別
'        MODULE_ID        :MSZZD00_YOKIN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_YOKIN(MSZZD00_YOKII As Integer) As String
'    On Error GoTo ERR_MSZZD00_YOKIN
    
    MSZZD00_YOKIN = Nz(DLookup("CODET_NAMEN", "CODE_TABL", _
                               "CODET_SIKBC = '121' " & " AND " & _
                               "CODET_CODEC = " & MSZZD00_YOKII))
'    Exit Function
'ERR_MSZZD00_YOKIN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :支払指示書名称
'        MODULE_ID        :MSZZD00_SIIHN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_SIIHN(MSZZD00_SIIHI As Integer) As String
'    On Error GoTo ERR_MSZZD00_SIIHN
    
    MSZZD00_SIIHN = Nz(DLookup("CODET_NAMEN", "CODE_TABL", _
                               "CODET_SIKBC = '202' " & " AND " & _
                               "CODET_CODEC = " & MSZZD00_SIIHI))
'    Exit Function
'ERR_MSZZD00_SIIHN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :集金区分
'        MODULE_ID        :MSZZD00_SKBNN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_SKBNN(MSZZD00_SKBNI As Integer) As String
'    On Error GoTo ERR_MSZZD00_SKBNN

    MSZZD00_SKBNN = Nz(DLookup("CODET_NAMEN", "CODE_TABL", _
                               "CODET_SIKBC = '030' " & " AND " & _
                               "CODET_CODEC = " & MSZZD00_SKBNI))
'    Exit Function
'ERR_MSZZD00_SKBNN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :業者名称
'        MODULE_ID        :MSZZD00_GYOUN
'        CREATE_DATE      :2004/04/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_GYOUN(MSZZD00_GYOUC As String) As String
'    On Error GoTo ERR_MSZZD00_GYOUN
    
    MSZZD00_GYOUN = Nz(DLookup("SHIRM_SHIRN", "SHIR_BUMA", _
                               "SHIRM_SHIRC = " & "'" & MSZZD00_GYOUC & "'"))

'    Exit Function
'ERR_MSZZD00_GYOUN:
'    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :業者カナ
'        MODULE_ID        :MSZZD00_GYOUF
'        CREATE_DATE      :2004/04/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_GYOUF(MSZZD00_GYOUC As String) As String
    
    MSZZD00_GYOUF = Nz(DLookup("SHIRM_SHIRF", "SHIR_BUMA", _
                               "SHIRM_SHIRC = " & "'" & MSZZD00_GYOUC & "'"))

End Function
'==============================================================================*
'
'        MODULE_NAME      :次回更新名称
'        MODULE_ID        :MSZZD00_JNEXN
'        CREATE_DATE      :2004/06/07
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_JNEXN(MSZZD00_JNEXI As String) As String
    
    MSZZD00_JNEXN = Nz(DLookup("JNEXN", "CODE_C231", _
                               "JNEXI = " & "'" & MSZZD00_JNEXI & "'"))

End Function
'==============================================================================*
'
'        MODULE_NAME      :賃料改定名称
'        MODULE_ID        :MSZZD00_CHIKN
'        CREATE_DATE      :2005/07/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_CHIKN(MSZZD00_CHIKI As String) As String
    
    MSZZD00_CHIKN = Nz(DLookup("CHIKN", "CODE_C233", _
                               "CHIKI = " & "'" & MSZZD00_CHIKI & "'"))

End Function
'==============================================================================*
'
'        MODULE_NAME      :INTI_FILE
'        MODULE_ID        :MSZZD00_RECDB
'        CREATE_DATE      :2008/12/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZD00_RECDB(strINTIF_PROGB As String, strINTIF_RECFB As String) As String
On Error GoTo ErrorHandler
    
    Dim strWhere            As String
    
    strWhere = "     INTIF_PROGB = '" & strINTIF_PROGB & "'" _
             & " AND INTIF_RECFB = '" & strINTIF_RECFB & "'"
             
    MSZZD00_RECDB = Nz(DLookup("INTIF_RECDB", "INTI_FILE", strWhere), "")

    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "MSZZD00_RECDB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'       MODULE_NAME     : INTIF_RECFB をキーに持つ INTIF_RECDB のコレクション
'       MODULE_ID       : MSZZD00_RECDB_Collection
'       CREATE_DATE     : 2011/02/17            K.ISHIZAKA
'       PARAM           : strPROGB              プログラムＩＤ(I)
'       RETURN          : コレクション
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZD00_RECDB_Collection(ByVal strPROGB As String) As Collection
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = "SELECT INTIF_RECDB, INTIF_RECFB FROM INTI_FILE WHERE INTIF_PROGB = '" & strPROGB & "'"
    Set MSZZD00_RECDB_Collection = MSZZD00_Collection(strSQL, "INTIF_RECFB", "INTIF_RECDB")
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "MSZZD00_RECDB_Collection" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : テーブルをコレクションに変換する
'       MODULE_ID       : MSZZD00_Collection
'       CREATE_DATE     : 2011/02/17            K.ISHIZAKA
'       PARAM           : strSQL                ＳＱＬ文(I)
'                       : strKeyFieldName       キーとなる列名(I)
'                       : strValueFieldName     値となる列名(I)
'       RETURN          : コレクション
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZD00_Collection(ByVal strSQL As String, ByVal strKeyFieldName As String, ByVal strValueFieldName As String)
    Dim objRst              As Recordset
    Dim colItem             As New Collection
    On Error GoTo ErrorHandler
    
    Set objRst = CurrentDb.OpenRecordset(strSQL, dbOpenForwardOnly, dbReadOnly)
    On Error GoTo ErrorHandler1
    With objRst
        While Not .EOF
            colItem.Add .Fields(strValueFieldName).VALUE, .Fields(strKeyFieldName).VALUE
            .MoveNext
        Wend
        .Close
    End With
    On Error GoTo ErrorHandler
    Set MSZZD00_Collection = colItem
Exit Function

ErrorHandler1:
    objRst.Close
ErrorHandler:
    Call Err.Raise(Err.Number, "MSZZD00_Collection" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  ended of program ********************************
