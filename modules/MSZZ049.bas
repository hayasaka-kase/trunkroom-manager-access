Attribute VB_Name = "MSZZ049"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : 解約ヘルパー出力
'        PROGRAM_ID      : MSZZ049
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2009/02/01
'        CERATER         : kihara
'        Ver             : 0.0
'
'        UPDATE          : 2009/05/01
'        UPDATER         : kihara
'        Ver             : 0.1
'                        : 地図・配置シートの検索コードを変更
'                          （移動元ヤードコード→移動先ヤードコード）
'
'        UPDATE          : 2011/08/09
'        UPDATER         : M.RYU
'        Ver             : 0.2
'                        : 地図と配置図を挿入したとき、ヘッダ左右部を空白にする
'
'        UPDATE          : 2011/08/18
'        UPDATER         : M.RYU
'        Ver             : 0.3
'                        : データ取得sqlを修正
'                        : 解約ヘルパー⇒承諾書に「変更前契約№」「変更後契約№」を追加
'                        : 変更後の月額賃料を実際に表示
'
'        UPDATE          : 2011/09/30
'        UPDATER         : M.RYU
'        Ver             : 0.4
'                        : 返送日を設定、FVS220画面に返送日を入力しヘルパーに表示
'                        : 自動は「書類作成日+26日」
'
'        UPDATE          : 2012/01/11
'        UPDATER         : M.RYU
'        Ver             : 0.5
'                        : ヘルパー出力するとき、受付データなしの場合中止
'
'        UPDATE          : 2012/02/02
'        UPDATER         : M.HONDA
'        Ver             : 0.6
'                        : 電話番号がNULLの際には空白へ変換する。
'
'        UPDATE          : 2012/11/17
'        UPDATER         : M.HONDA
'        Ver             : 0.7
'                        : 移動前の段・サイズを表示するように修正。
'
'        UPDATE          : 2017/11/02
'        UPDATER         : N.IMAI
'        Ver             : 0.8
'                        : データ取得の検索条件に解約キャンセルは除外するを追加
'
'        UPDATE          : 2018/03/10
'        UPDATER         : N.IMAI
'        Ver             : 0.9
'                        : 解約用電話番号を部門毎に変更可能とする(CONT_MAST)
'
'        UPDATE          : 2018/09/28
'        UPDATER         : EGL
'        Ver             : 1.0
'                        : 分社化対応
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_HELPER_ヤード解約       As String = "解約ヘルパー.xls"

'コントロールマスタ情報・構造体
Public Type Type_CONT_MAST
    CONT_KAISYA         As String   '貸主会社名
    CONT_TEL_NO         As String   '貸主TEL
    CONT_FAX_NO         As String   '貸主FAX
    CONT_CANCEL_TEL_NO  As String   '解約専用TEL INSERT 2018/03/10 N.IMAI
'    CONT_TANNM          As String   '貸主担当  DEL 2009/4/30 KIHARA
    CONT_YUBINO         As String   '貸主郵便番号
    CONT_ADDR_1         As String   '貸主住所１
    CONT_ADDR_2         As String   '貸主住所２
End Type

'コンテナ契約情報・構造体
Public Type Type_CARG_INF
    USER_YUBINO         As String   '契約者郵便番号
    USER_ADR_1          As String   '契約者住所１
    USER_ADR_2          As String   '契約者住所２
    USER_ADR_3          As String   '契約者住所３
    USER_NAME           As String   '契約者名
    USER_TANM           As String   '契約代表者名
    CARG_UCODE          As String   '顧客コード
    USER_TEL            As String   '契約者TEL
    USER_FAX            As String   '契約者FAX
    USER_KEITAI         As String   '契約者CEL
    YARD_NAME           As String   'ヤード名
    CARG_YCODE          As String   'ヤードコード
    YARD_ADDR_1         As String   'ヤード住所１
    YARD_ADDR_2         As String   'ヤード住所２
    CARG_NO             As String   'スペースコード
    CNTA_STEP           As String   '上下段コード
    CNTA_STEP_NM        As String   '上下段コード　' INS M.HONDA 2012/11/17
    CNTA_SIZE           As String   'ｻｲｽﾞ          ' INS M.HONDA 2012/11/17
    GETSUGAKU           As String   '月額使用料
    ZAPPI               As String   '他月額料
    YARD_END_DAY        As Variant  '解約日
    CARG_ACPTNO         As String   '受注契約番号
End Type

'移動先コンテナ契約情報・構造体
Public Type Type_CARG_INF2
    IDO_YARD_NAME       As String   '移動先ヤード名
    IDO_CARG_YCODE      As String   '移動先ヤードコード
    IDO_YARD_ADDR_1     As String   '移動先ヤード住所１
    IDO_YARD_ADDR_2     As String   '移動先ヤード住所２
    IDO_CARG_NO         As String   '移動先スペースコード
    IDO_CNTA_SIZE       As String   '移動先スペースサイズ
    IDO_CNTA_STEP       As String   '移動先上下段コード
    IDO_CNTA_STEP_NM    As String   '移動先上下段
    IDO_YOTO_NM         As String   '移動先用途名
    IDO_GETSUGAKU       As String   '月額使用料
    IDO_ZAPPI           As String   '他月額料
    IDO_ACPTNO          As String   '移行後契約番号     'INSERT 2011/08/18 M.RYU
End Type

'予約受付トラン情報・構造体　   ADD 2009/04/30 KIHARA
Public Type Type_YOUK_TRAN
    YOUKT_UKTANTO       As String   '受付担当コード
End Type

'担当者マスタ情報・構造体　     ADD 2009/04/30 KIHARA
Public Type Type_TANT_MAST
    TANTM_TANTN         As String   '担当者名
End Type

'部門コード
Private strBumonCode    As String


'==============================================================================*
'
'       MODULE_NAME     : HelperPrintPreviewKyk
'       機能            : ヤード解約書のプレビュー
'       CREATE_DATE     : 2009/02/01            KIHARA
'       PARAM           : strBumoc              部門コード(I)
'                       : strYardc              ヤードコード(I)
'                       : strCARG_NO            コンテナ番号(I)
'                       : strYardcido           移動先ヤードコード(I)
'                       : strCARG_Noido         移動先コンテナ番号(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub HelperPrintPreviewKyk(ByVal strBUMOC As String, ByVal strYardc As String, ByVal strCARG_NO As String, ByVal strYardcido As String, _
                                 ByVal strCARG_Noido As String, intPrevew As Integer)
    Dim stCARG              As Type_CARG_INF
    Dim stCARG2             As Type_CARG_INF2
    Dim stCONT              As Type_CONT_MAST
    Dim stYOUK              As Type_YOUK_TRAN   'ADD 2009/04/30 KIHARA
    Dim stTANT              As Type_TANT_MAST   'ADD 2009/04/30 KIHARA

    On Error GoTo ErrorHandler
    
    '解約データ取得
'    Call GetHelperDataKyk(strBumoc, strYardc, strCARG_NO, strYardcido, strCARG_Noido, stCONT, stCARG, stCARG2)                 'DEL 2009/04/30 KIHARA
    Call GetHelperDataKyk(strBUMOC, strYardc, strCARG_NO, strYardcido, strCARG_Noido, stCONT, stCARG, stCARG2, stYOUK, stTANT)  'ADD 2009/04/30 KIHARA
    
    If Nz(stCARG.CARG_UCODE) = "" Then Exit Sub     'INSERT 2012/01/11 M.RYU
        
    'プレビュー表示
'    Call HelperPrintXXKyk(strBumoc, strYardc, stCARG, stCONT, stCARG2, intPrevew)                                              'DEL 2009/04/30 KIHARA
'    Call HelperPrintXXKyk(strBumoc, strYardc, stCARG, stCONT, stCARG2, stTANT, intPrevew)               'ADD 2009/04/30 KIHARA 'DEL 2009/05/01 KIHARA
    Call HelperPrintXXKyk(strBUMOC, strYardc, strYardcido, stCARG, stCONT, stCARG2, stTANT, intPrevew)                          'ADD 2009/05/01 KIHARA

Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "HelperPrintPreviewKyk" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : GetHelperDataKyk
'       機能            : 解約データ取得
'       CREATE_DATE     : 2009/02/01            KIHARA
'       PARAM           : strBumoc              部門コード(I)
'                       : strYardc              ヤードコード(I)
'                       : strCARG_NO            コンテナ番号(I)
'                       : strYardcido           移動先ヤードコード(I)
'                       : strCARG_Noido         移動先コンテナ番号(I)
'                       : stCONT                コントロールマスタ情報(O)
'                       : stCARG                コンテナ契約ファイル情報(O)
'                       : stCARG2               移動先コンテナ契約ファイル情報(O)
'                       : stYOUK                予約受付トラン情報(O)       'ADD 2009/04/30 KIHARA
'                       : stTANT                担当マスタ情報(O)           'ADD 2009/04/30 KIHARA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'MOD 2009/04/30 KIHARA
Private Sub GetHelperDataKyk(ByVal strBUMOC As String, ByVal strYardc As String, ByVal strCARG_NO As String, ByVal strYardcido As String, _
                            ByVal strCARG_Noido As String, stCONT As Type_CONT_MAST, stCARG As Type_CARG_INF, stCARG2 As Type_CARG_INF2, stYOUK As Type_YOUK_TRAN, stTANT As Type_TANT_MAST)
    
    Dim objCon              As Object
    Dim objCon2             As Object   'ADD 2009/04/30 KIHARA
    On Error GoTo ErrorHandler

    Set objCon = ADODB_Connection(strBUMOC)
    Set objCon2 = ADODB_Connection()    'ADD 2009/04/30 KIHARA
    On Error GoTo ErrorHandler1
    
    'コントロールマスタ情報取得
    Call Select_CONT_MASTKyk(objCon, stCONT)
    
    
    '---↓↓--Delete 2011/08/18 M.RYU------↓↓---<s>
'    'コンテナ契約ファイル情報取得
'    Call Select_CARG_FILE(objCon, strYARDC, strCARG_NO, stCARG)
'    '移動先コンテナ契約ファイル情報取得
'    Call Select_CARG_FILE2(objCon, strYARDC, strCARG_NO, strYardcido, strCARG_Noido, stCARG2)
''ADD 2009/4/30 KIHARA Start
'    '予約受付トラン情報取得
'    Call Select_YOUK_TRAN(objCon, stYOUK, stCARG)
'    '担当者マスタ情報取得
'    Call Select_TANT_MAST(objCon2, stTANT, stYOUK, strBumoc)
''ADD 2009/4/30 KIHARA End
    '---↑↑--Delete 2011/08/18 M.RYU------↑↑---<e>
    
    Call Select_Data(objCon, strBUMOC, strYardc, strCARG_NO, strYardcido, _
                     strCARG_Noido, stCARG, stCARG2, stYOUK, stTANT)   'INSERT 2011/08/18 M.RYU

    objCon.Close
    objCon2.Close                       'ADD 2009/04/30 KIHARA
    On Error GoTo ErrorHandler

Exit Sub

ErrorHandler1:
    objCon.Close
    objCon2.Close                       'ADD 2009/04/30 KIHARA
ErrorHandler:
    Call Err.Raise(Err.Number, "GetHelperDataKyk" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub


'==============================================================================*
'
'       MODULE_NAME     : Select_CONT_MASTKyk
'       機能            : コントロールマスタ情報取得
'       CREATE_DATE     : 2009/02/01            KIHARA
'       PARAM           : objCon            　　コネクション(I)
'                       : stCONT                コントロールマスタ情報(O)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Select_CONT_MASTKyk(objCon As Object, stCONT As Type_CONT_MAST)
    Dim objRst              As Object
    Dim strSQL              As String
    On Error GoTo ErrorHandler

    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CONT_KAISYA,"
    strSQL = strSQL & " CONT_CANCEL_TEL_NO,"                                    'INSERT 2018/03/10 N.IMAI
    strSQL = strSQL & " CONT_TEL_NO,"
    strSQL = strSQL & " CONT_FAX_NO "    'MOD 2009/4/30 KIHARA
'    strSQL = strSQL & " CONT_TANNM "    'DEL 2009/4/30 KIHARA
    strSQL = strSQL & ",CONT_YUBINO "    'ADD 2018/09/28 EGL
    strSQL = strSQL & ",CONT_ADDR_1 "    'ADD 2018/09/28 EGL
    strSQL = strSQL & ",CONT_ADDR_2 "    'ADD 2018/09/28 EGL
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " CONT_MAST "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & "  CONT_KEY = 1"
    
    Set objRst = ADODB_Recordset(strSQL, objCon)
    On Error GoTo ErrorHandler1
    With objRst
        stCONT.CONT_KAISYA = .Fields("CONT_KAISYA")                 '貸主会社名
        stCONT.CONT_TEL_NO = .Fields("CONT_TEL_NO")                 '貸主TEL    'INSERT 2018/03/10 N.IMAI
        stCONT.CONT_FAX_NO = .Fields("CONT_FAX_NO")                 '貸主FAX
'        stCONT.CONT_TANNM = .Fields("CONT_TANNM")                  '貸主担当  DEL 2009/4/30 KIHARA
        stCONT.CONT_CANCEL_TEL_NO = .Fields("CONT_CANCEL_TEL_NO")   '貸主TEL   INSERT 2018/03/10 N.IMAI
        stCONT.CONT_YUBINO = .Fields("CONT_YUBINO")     'ADD 2018/09/28 EGL
        stCONT.CONT_ADDR_1 = .Fields("CONT_ADDR_1")     'ADD 2018/09/28 EGL
        stCONT.CONT_ADDR_2 = Nz(.Fields("CONT_ADDR_2"), "")    'ADD 2018/09/28 EGL
        .Close
    End With
    On Error GoTo ErrorHandler
Exit Sub
    
ErrorHandler1:
    objRst.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Select_CONT_MASTKyk" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : Select_Data
'       機能            : 情報取得
'       CREATE_DATE     : 2011/08/18 M.RYU
'       PARAM           : strBumoc              部門コード(I)
'                       : strYardc              ヤードコード(I)
'                       : strCARG_NO            コンテナ番号(I)
'                       : strYardcido           移動先ヤードコード(I)
'                       : strCARG_Noido         移動先コンテナ番号(I)
'                       : stCARG                コンテナ契約ファイル情報(O)
'                       : stCARG2               移動先コンテナ契約ファイル情報(O)
'                       : stYOUK                予約受付トラン情報(O)
'                       : stTANT                担当マスタ情報(O)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Select_Data(objCon As Object, ByVal strBUMOC As String, _
                        ByVal strYardc As String, ByVal strCARG_NO As String, _
                        ByVal strYardcido As String, ByVal strCARG_Noido As String, _
                        stCARG As Type_CARG_INF, stCARG2 As Type_CARG_INF2, _
                        stYOUK As Type_YOUK_TRAN, stTANT As Type_TANT_MAST)
                             
    Dim objRst              As Object
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strSQL = fncMakeGetDataSql(strBUMOC, strYardc, strCARG_NO, strYardcido, strCARG_Noido)
    Set objRst = ADODB_Recordset(strSQL, objCon)
    On Error GoTo ErrorHandler1
    
    If objRst.EOF Then                      'INSERT 2012/01/11 M.RYU
        MsgBox "受付データがありません。"   'INSERT 2012/01/11 M.RYU
        Exit Sub                            'INSERT 2012/01/11 M.RYU
    End If
    
    With objRst
        stCARG.USER_YUBINO = .Fields("USER_YUBINO")                 '契約者郵便番号
        stCARG.USER_ADR_1 = .Fields("USER_ADR_1")                   '契約者住所１
        stCARG.USER_ADR_2 = .Fields("USER_ADR_2")                   '契約者住所２
        stCARG.USER_ADR_3 = Nz(.Fields("USER_ADR_3"), "")           '契約者住所３
        stCARG.USER_NAME = Nz(.Fields("USER_NAME"), "")             '契約者名
        stCARG.USER_TANM = Nz(.Fields("USER_TANM"), "")             '契約代表者名
        stCARG.CARG_UCODE = Format(.Fields("CARG_UCODE"), "000000") '顧客コード
        '' 2012/03/02 M.HONDA START
        ''stCARG.USER_TEL = .Fields("USER_TEL")                     '契約者TEL
        stCARG.USER_TEL = Nz(.Fields("USER_TEL"), "")               '契約者TEL
        '' 2012/03/02 M.HONDA END
        stCARG.USER_FAX = Nz(.Fields("USER_FAX"), "")               '契約者FAX
        stCARG.USER_KEITAI = Nz(.Fields("USER_KEITAI"), "")         '契約者CEL
        stCARG.CARG_YCODE = Format(.Fields("CARG_YCODE"), "000000") 'ヤードコード
        stCARG.YARD_NAME = .Fields("YARD_NAME")                     'ヤード名
        stCARG.YARD_ADDR_1 = Nz(.Fields("YARD_ADDR_1"), "")         'ヤード住所１
        stCARG.YARD_ADDR_2 = Nz(.Fields("YARD_ADDR_2"), "")         'ヤード住所２
        stCARG.CARG_NO = Format(.Fields("CARG_NO"), "000000")       'スペースコード
        stCARG.GETSUGAKU = .Fields("GETSUGAKU")                     '月額使用料
        stCARG.ZAPPI = Nz(.Fields("ZAPPI"), "")                     '他月額料
        stCARG.YARD_END_DAY = Nz(.Fields("YARD_END_DAY"))           '解約日
        stCARG.CNTA_STEP = Nz(.Fields("CNTA_STEP"))                 '上下段コード
        stCARG.CNTA_STEP_NM = .Fields("CNTA_STEP_NAME")             '移動先上下段  '' 2012/11/07 INS M.HONDA
        stCARG.CNTA_SIZE = .Fields("CNTA_SIZE")                     'ｻｲｽﾞ          '' 2012/11/07 INS M.HONDA
        stCARG.CARG_ACPTNO = Nz(.Fields("CARG_ACPTNO"))             '受注契約番号
        stCARG2.IDO_CARG_YCODE = .Fields("IDO_CARG_YCODE")           '移動先ヤードコード
        stCARG2.IDO_YARD_NAME = .Fields("IDO_YARD_NAME")             '移動先ヤード名
        stCARG2.IDO_YARD_ADDR_1 = Nz(.Fields("IDO_YARD_ADDR_1"), "") '移動先ヤード住所１
        stCARG2.IDO_YARD_ADDR_2 = Nz(.Fields("IDO_YARD_ADDR_2"), "") '移動先ヤード住所２
        stCARG2.IDO_CARG_NO = .Fields("IDO_CARG_NO")                 '移動先スペースコード
        stCARG2.IDO_CNTA_SIZE = .Fields("IDO_CNTA_SIZE")             '移動先スペースサイズ
        stCARG2.IDO_CNTA_STEP = .Fields("IDO_CNTA_STEP")             '移動先上下段コード
        stCARG2.IDO_CNTA_STEP_NM = .Fields("IDO_CNTA_STEP_NAME")     '移動先上下段
        stCARG2.IDO_YOTO_NM = .Fields("IDO_YOTO_NAME")               '移動先レンタル用途
        stCARG2.IDO_GETSUGAKU = Nz(.Fields("IDO_GETSUGAKU"), "")     '移動先月額
        stCARG2.IDO_ZAPPI = Nz(.Fields("IDO_ZAPPI"), "")             '移動先他月額
        stCARG2.IDO_ACPTNO = Nz(.Fields("IDO_ACPTNO"), "")           '変更後契約№      'INSERT 2011/08/18 M.RYU
        
        stYOUK.YOUKT_UKTANTO = .Fields("YOUKT_UKTANTO")              '受付担当コード
        stTANT.TANTM_TANTN = .Fields("TANTM_TANTN")                  '担当者名
        .Close
    End With
    On Error GoTo ErrorHandler
Exit Sub

ErrorHandler1:
    objRst.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Select_Data" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : Select_Data
'       機能            : 情報取得
'       CREATE_DATE     : 2011/08/18 M.RYU
'       PARAM           : strBumoc              部門コード(I)
'                       : strYardc              ヤードコード(I)
'                       : strCARG_NO            コンテナ番号(I)
'                       : strYardcido           移動先ヤードコード(I)
'                       : strCARG_Noido         移動先コンテナ番号(I)
'                       : stCARG                コンテナ契約ファイル情報(O)
'                       : stCARG2               移動先コンテナ契約ファイル情報(O)
'                       : stYOUK                予約受付トラン情報(O)
'                       : stTANT                担当マスタ情報(O)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSql(ByVal strBUMOC As String, _
                                   ByVal strYardc As String, ByVal strCARG_NO As String, _
                                   ByVal strYardcido As String, ByVal strCARG_Noido As String) As String
    'KASE_DB名前を取得
    Dim strKASEDBN As String
    strKASEDBN = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATABASE_NAME'")
    strKASEDBN = strKASEDBN & ".dbo."

    Dim strSQL              As String
    strSQL = " SELECT * " & Chr(13)
    strSQL = strSQL & " FROM " & Chr(13)
    
    ' 変更前データ取得SQL
    ' 【SELECT句】
    strSQL = strSQL & " ( " & Chr(13)
    strSQL = strSQL & " SELECT  USER_MAST.USER_YUBINO  " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_ADR_1   " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_ADR_2   " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_ADR_3   " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_NAME    " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_TANM    " & Chr(13)
    strSQL = strSQL & "        ,CARG_FILE.CARG_UCODE   " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_TEL     " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_FAX     " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_KEITAI  " & Chr(13)
    strSQL = strSQL & "        ,CARG_FILE.CARG_YCODE   " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_NAME    " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_ADDR_1  " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_ADDR_2  " & Chr(13)
    strSQL = strSQL & "        ,CARG_FILE.CARG_NO      " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_RENTKG AS GETSUGAKU " & Chr(13)
    strSQL = strSQL & "        ,ISNULL(RCPT_TRAN.RCPT_EZAPPI,0) + ISNULL(RCPT_TRAN.RCPT_ADD_EZAPPI1,0) + ISNULL(RCPT_TRAN.RCPT_ADD_EZAPPI2,0)  AS ZAPPI " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_END_DAY " & Chr(13)
    strSQL = strSQL & "        ,CNTA_MAST.CNTA_STEP    " & Chr(13)
    strSQL = strSQL & "        ,STEP_NAME.NAME_NAME AS  CNTA_STEP_NAME " & Chr(13)
    strSQL = strSQL & "        ,CNTA_MAST.CNTA_SIZE    " & Chr(13)
    strSQL = strSQL & "        ,CARG_FILE.CARG_ACPTNO  " & Chr(13)
    strSQL = strSQL & " FROM ((((( CARG_FILE INNER JOIN YARD_MAST ON CARG_FILE.CARG_YCODE = YARD_MAST.YARD_CODE ) " & Chr(13)  ' ヤードマスタ
    strSQL = strSQL & "   INNER JOIN CNTA_MAST ON ( CARG_FILE.CARG_YCODE = CNTA_MAST.CNTA_CODE )    " & Chr(13)
    strSQL = strSQL & "                       AND ( CARG_FILE.CARG_NO    = CNTA_MAST.CNTA_NO ) )    " & Chr(13)                  ' コンテナマスタ
    strSQL = strSQL & "   INNER JOIN USER_MAST ON CARG_FILE.CARG_UCODE   = USER_MAST.USER_CODE )    " & Chr(13)                  ' ユーザーマスタ
    strSQL = strSQL & "    LEFT JOIN RCPT_TRAN ON CARG_FILE.CARG_UKNO    = RCPT_TRAN.RCPT_NO   )    " & Chr(13)                  ' RCPT_TRAN
    strSQL = strSQL & "    LEFT  JOIN NAME_MAST STEP_NAME ON STEP_NAME.NAME_ID = '090' AND STEP_NAME.NAME_CODE = CNTA_MAST.CNTA_STEP)    " & Chr(13) '段階名前：上段、下段
    strSQL = strSQL & " WHERE CARG_FILE.CARG_YCODE = " & strYardc & "   " & Chr(13)
    strSQL = strSQL & "   AND CARG_FILE.CARG_NO = " & strCARG_NO & "    " & Chr(13)
    strSQL = strSQL & " ) Before " & Chr(13)  ' 変更前

    ' 変更後データ取得SQL
    ' 【SELECT句】
    strSQL = strSQL & " , " & Chr(13)
    strSQL = strSQL & " ( " & Chr(13)
    strSQL = strSQL & " SELECT  CNTA_MAST.CNTA_USAGE  AS IDO_CNTA_USAGE     " & Chr(13)
    strSQL = strSQL & "        ,NAME_MAST.NAME_NAME   AS IDO_YOTO_NAME      " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_YCODE  AS IDO_CARG_YCODE     " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_NAME   AS IDO_YARD_NAME      " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_ADDR_1 AS IDO_YARD_ADDR_1    " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_ADDR_2 AS IDO_YARD_ADDR_2    " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_CNO    AS IDO_CARG_NO        " & Chr(13)
    strSQL = strSQL & "        ,CNTA_MAST.CNTA_SIZE   AS IDO_CNTA_SIZE      " & Chr(13)
    strSQL = strSQL & "        ,CNTA_MAST.CNTA_STEP   AS IDO_CNTA_STEP      " & Chr(13)
    strSQL = strSQL & "        ,STEP_NAME.NAME_NAME   AS IDO_CNTA_STEP_NAME " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_RENTKG AS IDO_GETSUGAKU      " & Chr(13)
    strSQL = strSQL & "        ,ISNULL(RCPT_TRAN.RCPT_EZAPPI,0) + ISNULL(RCPT_TRAN.RCPT_ADD_EZAPPI1,0) + ISNULL(RCPT_TRAN.RCPT_ADD_EZAPPI2,0) AS IDO_ZAPPI  " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_CARG_ACPTNO  AS IDO_ACPTNO   " & Chr(13)
    strSQL = strSQL & "        ,YOUK_TRAN.YOUKT_MOTO_ACPTNO AS IDO_MOTO_ACPTNO " & Chr(13)
    strSQL = strSQL & "        ,YOUK_TRAN.YOUKT_UKTANTO                     " & Chr(13)
    strSQL = strSQL & "        ,TANT_MAST.TANTM_TANTN                       " & Chr(13)
    strSQL = strSQL & "   FROM (((((( RCPT_TRAN INNER JOIN YARD_MAST ON RCPT_TRAN.RCPT_YCODE = YARD_MAST.YARD_CODE ) " & Chr(13)  ' ヤードマスタ
    strSQL = strSQL & "   INNER JOIN CNTA_MAST ON ( RCPT_TRAN.RCPT_YCODE = CNTA_MAST.CNTA_CODE )    " & Chr(13)
    strSQL = strSQL & "                       AND ( RCPT_TRAN.RCPT_CNO   = CNTA_MAST.CNTA_NO ) )    " & Chr(13)                  ' コンテナマスタ
    'strSQL = strSQL & "   INNER JOIN YOUK_TRAN ON RCPT_TRAN.RCPT_NO      = YOUK_TRAN.YOUKT_UKNO)    " & Chr(13)                  ' 予約受付トラン  'DEL 2017/11/02 N.IMAI
    strSQL = strSQL & "   INNER JOIN YOUK_TRAN ON RCPT_TRAN.RCPT_NO      = YOUK_TRAN.YOUKT_UKNO AND YOUK_TRAN.YOUKT_YTDATE IS NULL) " & Chr(13)     'ADD 2017/11/02 N.IMAI
    strSQL = strSQL & "   INNER JOIN NAME_MAST ON CNTA_MAST.CNTA_USAGE   = NAME_MAST.NAME_CODE )    " & Chr(13)                  ' NAME_MAST
    strSQL = strSQL & "   LEFT  JOIN NAME_MAST STEP_NAME ON STEP_NAME.NAME_ID = '090' AND STEP_NAME.NAME_CODE = CNTA_MAST.CNTA_STEP)    " & Chr(13) '段階名前：上段、下段
    strSQL = strSQL & "   LEFT  JOIN " & strKASEDBN & "TANT_MAST on TANT_MAST.TANTM_BUMOC = '" & strBUMOC & "' AND YOUK_TRAN.YOUKT_UKTANTO = TANT_MAST.TANTM_TANTC )"   '担当マスタ
    strSQL = strSQL & " WHERE NAME_MAST.NAME_ID = '086' " & Chr(13)
    strSQL = strSQL & "   AND RCPT_TRAN.RCPT_YCODE = " & strYardcido & "   " & Chr(13)
    strSQL = strSQL & "   AND RCPT_TRAN.RCPT_CNO = " & strCARG_Noido & "   " & Chr(13)
    strSQL = strSQL & " ) After " & Chr(13)  ' 変更後
    strSQL = strSQL & " WHERE Before.CARG_ACPTNO = After.IDO_MOTO_ACPTNO"  ' 変更前と変更後の結合条件
    
    fncMakeGetDataSql = strSQL

End Function
'------↓↓↓↓-----Delete 2011/08/18 M.RYU--------↓↓↓↓--------<s>
''==============================================================================*
''
''       MODULE_NAME     : Select_CARG_FILE
''       機能            : コンテナ契約ファイル情報取得
''       CREATE_DATE     : 2009/02/01            KIHARA
''       PARAM           : objCon                コネクション(I)
''                       : strYardc              ヤードコード(I)
''                       : strCARG_NO            コンテナ番号(I)
''                       : stCARG                コンテナ契約ファイル情報(O)
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub Select_CARG_FILE(objCon As Object, ByVal strYARDC As String, ByVal strCARG_NO As String, stCARG As Type_CARG_INF)
'    Dim objRst              As Object
'    Dim strSQL              As String
'
'    On Error GoTo ErrorHandler
'
'    strSQL = strSQL & "SELECT"
'    strSQL = strSQL & " USER_YUBINO,"
'    strSQL = strSQL & " USER_ADR_1,"
'    strSQL = strSQL & " USER_ADR_2,"
'    strSQL = strSQL & " USER_ADR_3,"
'    strSQL = strSQL & " USER_NAME,"
'    strSQL = strSQL & " USER_TANM,"
'    strSQL = strSQL & " CARG_UCODE,"
'    strSQL = strSQL & " USER_TEL,"
'    strSQL = strSQL & " USER_FAX,"
'    strSQL = strSQL & " USER_KEITAI,"
'    strSQL = strSQL & " YARD_NAME,"
'    strSQL = strSQL & " CARG_YCODE,"
'    strSQL = strSQL & " YARD_ADDR_1,"
'    strSQL = strSQL & " YARD_ADDR_2,"
'    strSQL = strSQL & " CARG_NO,"
'    strSQL = strSQL & " ISNULL(CARG_RENTKG,0) + ISNULL(CARG_SYOZEI,0) GETSUGAKU,"
'    strSQL = strSQL & " ISNULL(RCPT_EZAPPI,0) + ISNULL(RCPT_ADD_EZAPPI1,0) + ISNULL(RCPT_ADD_EZAPPI2,0) ZAPPI,"
'    strSQL = strSQL & " YARD_END_DAY, "
''    strSQL = strSQL & " CNTA_STEP "        'DEL 2009/04/30 KIHARA
'    strSQL = strSQL & " CNTA_STEP, "        'MOD 2009/04/30 KIHARA
'    strSQL = strSQL & " CARG_ACPTNO "       'ADD 2009/04/30 KIHARA
'
'    strSQL = strSQL & "FROM"
'    strSQL = strSQL & " CARG_FILE"
'    strSQL = strSQL & " INNER JOIN USER_MAST"
'    strSQL = strSQL & " ON USER_CODE = CARG_UCODE"
'    strSQL = strSQL & " LEFT OUTER JOIN YARD_MAST"
'    strSQL = strSQL & " ON YARD_CODE = CARG_YCODE"
'    strSQL = strSQL & " LEFT OUTER JOIN RCPT_TRAN"
'    strSQL = strSQL & " ON RCPT_YCODE = CARG_YCODE"
'    strSQL = strSQL & " AND RCPT_CNO = CARG_NO "
'    strSQL = strSQL & " LEFT OUTER JOIN CNTA_MAST ON (CNTA_CODE = CARG_YCODE AND CNTA_NO = CARG_NO ) "
'
'    strSQL = strSQL & "WHERE"
'    strSQL = strSQL & " CARG_YCODE = " & strYARDC
'    strSQL = strSQL & " AND  CARG_NO = " & strCARG_NO
'
'    Set objRst = ADODB_Recordset(strSQL, objCon)
'    On Error GoTo ErrorHandler1
'
'    With objRst
'        stCARG.USER_YUBINO = .Fields("USER_YUBINO")                 '契約者郵便番号
'        stCARG.USER_ADR_1 = .Fields("USER_ADR_1")                   '契約者住所１
'        stCARG.USER_ADR_2 = .Fields("USER_ADR_2")                   '契約者住所２
'        stCARG.USER_ADR_3 = Nz(.Fields("USER_ADR_3"), "")           '契約者住所３
'        stCARG.USER_NAME = Nz(.Fields("USER_NAME"), "")             '契約者名
'        stCARG.USER_TANM = Nz(.Fields("USER_TANM"), "")             '契約代表者名
'        stCARG.CARG_UCODE = Format(.Fields("CARG_UCODE"), "000000") '顧客コード
'        stCARG.USER_TEL = .Fields("USER_TEL")                       '契約者TEL
'        stCARG.USER_FAX = Nz(.Fields("USER_FAX"), "")               '契約者FAX
'        stCARG.USER_KEITAI = Nz(.Fields("USER_KEITAI"), "")         '契約者CEL
'        stCARG.YARD_NAME = .Fields("YARD_NAME")                     'ヤード名
'        stCARG.CARG_YCODE = Format(.Fields("CARG_YCODE"), "000000") 'ヤードコード
'        stCARG.YARD_ADDR_1 = Nz(.Fields("YARD_ADDR_1"), "")         'ヤード住所１
'        stCARG.YARD_ADDR_2 = Nz(.Fields("YARD_ADDR_2"), "")         'ヤード住所２
'        stCARG.CARG_NO = Format(.Fields("CARG_NO"), "000000")       'スペースコード
'        stCARG.GETSUGAKU = .Fields("GETSUGAKU")                     '月額使用料
'        stCARG.ZAPPI = Nz(.Fields("ZAPPI"), "")                     '他月額料
'        stCARG.YARD_END_DAY = Nz(.Fields("YARD_END_DAY"))           '解約日
'        stCARG.CNTA_STEP = Nz(.Fields("CNTA_STEP"))                 '上下段コード
'        stCARG.CARG_ACPTNO = Nz(.Fields("CARG_ACPTNO"))             '受注契約番号
'        .Close
'    End With
'    On Error GoTo ErrorHandler
'Exit Sub
'
'ErrorHandler1:
'    objRst.Close
'ErrorHandler:                   '↓自分の関数名
'    Call Err.Raise(Err.Number, "Select_CARG_FILE" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'End Sub

''==============================================================================*
''
''       MODULE_NAME     : Select_CARG_FILE2
''       機能            : 移動先コンテナ契約ファイル情報取得
''       CREATE_DATE     : 2009/02/01            KIHARA
''       PARAM           : objCon                コネクション(I)
''                       : strYardc              ヤードコード(I)
''                       : strCARG_NO            コンテナ番号(I)
''                       : strYardcido           移動先ヤードコード(O)
''                       : strCARG_Noido         移動先コンテナ番号(O)
''                       : stCARG2               移動先コンテナ契約ファイル情報(O)
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub Select_CARG_FILE2(objCon As Object, ByVal strYARDC As String, ByVal strCARG_NO As String, ByVal strYardcido As String, _
'                              ByVal strCARG_Noido As String, stCARG2 As Type_CARG_INF2)
'    Dim objRst              As Object
'    Dim strSQL              As String
'    On Error GoTo ErrorHandler
'
'    strSQL = strSQL & "SELECT"
'    strSQL = strSQL & " YARD_NAME,"
'    strSQL = strSQL & " YARD_ADDR_1,"
'    strSQL = strSQL & " YARD_ADDR_2,"
'    strSQL = strSQL & " CNTA_SIZE,"
'    strSQL = strSQL & " CNTA_STEP "
'    strSQL = strSQL & " ,STEP_NAME.NAME_NAME STEP_NAME "
'    strSQL = strSQL & " ,YOTO_NAME.NAME_NAME YOTO_NAME"
'    strSQL = strSQL & " ,PRIC_PRICE GETSUGAKU"
'    strSQL = strSQL & " ,PRIC_EZAPPI ZAPPI "
'
'    strSQL = strSQL & "FROM YARD_MAST "
'    strSQL = strSQL & " INNER JOIN CNTA_MAST ON (CNTA_CODE = YARD_CODE ) "
'    strSQL = strSQL & " LEFT OUTER JOIN NAME_MAST STEP_NAME ON "
'    strSQL = strSQL & "           STEP_NAME.NAME_ID = '090' AND STEP_NAME.NAME_CODE = CNTA_MAST.CNTA_STEP "
'    strSQL = strSQL & " LEFT OUTER JOIN NAME_MAST YOTO_NAME ON "
'    strSQL = strSQL & "            YOTO_NAME.NAME_ID = '086' AND YOTO_NAME.NAME_CODE = CNTA_MAST.CNTA_USAGE "
'    strSQL = strSQL & " LEFT OUTER JOIN PRIC_TABL ON"
'    strSQL = strSQL & "             YARD_MAST.YARD_CODE = PRIC_YCODE And CNTA_MAST.CNTA_STEP = PRIC_STEP And CNTA_MAST.CNTA_SIZE = PRIC_SIZE "
'
'    strSQL = strSQL & "WHERE"
'    strSQL = strSQL & " YARD_CODE = " & strYardcido
'    strSQL = strSQL & " AND  CNTA_NO = " & strCARG_Noido
'
'    Set objRst = ADODB_Recordset(strSQL, objCon)
'    On Error GoTo ErrorHandler1
'
'    With objRst
'        stCARG2.IDO_YARD_NAME = .Fields("YARD_NAME")                '移動先ヤード名
'        stCARG2.IDO_CARG_YCODE = strYardcido                        '移動先ヤードコード
'        stCARG2.IDO_YARD_ADDR_1 = Nz(.Fields("YARD_ADDR_1"), "")    '移動先ヤード住所１
'        stCARG2.IDO_YARD_ADDR_2 = Nz(.Fields("YARD_ADDR_2"), "")    '移動先ヤード住所２
'        stCARG2.IDO_CARG_NO = strCARG_Noido                         '移動先スペースコード
'        stCARG2.IDO_CNTA_SIZE = .Fields("CNTA_SIZE")                '移動先スペースサイズ
'        stCARG2.IDO_CNTA_STEP = .Fields("CNTA_STEP")                '移動先上下段コード
'        stCARG2.IDO_CNTA_STEP_NM = .Fields("STEP_NAME")             '移動先上下段
'        stCARG2.IDO_YOTO_NM = .Fields("YOTO_NAME")                  '移動先レンタル用途
'        stCARG2.IDO_GETSUGAKU = Nz(.Fields("GETSUGAKU"), "")        '移動先月額
'        stCARG2.IDO_ZAPPI = Nz(.Fields("ZAPPI"), "")                '移動先他月額
'        .Close
'    End With
'    On Error GoTo ErrorHandler
'Exit Sub
'
'ErrorHandler1:
'    objRst.Close
'ErrorHandler:                   '↓自分の関数名
'    Call Err.Raise(Err.Number, "Select_CARG_FILE2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'End Sub
'
''==============================================================================*
''
''       MODULE_NAME     : Select_YOUK_TRAN
''       機能            : 予約受付トラン情報取得
''       CREATE_DATE     : 2009/04/30            KIHARA
''       PARAM           : objCon            　　コネクション(I)
''                       : stYOUK                予約受付トラン情報(O)
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub Select_YOUK_TRAN(objCon As Object, stYOUK As Type_YOUK_TRAN, stCARG As Type_CARG_INF)
'    Dim objRst              As Object
'    Dim strSQL              As String
'    On Error GoTo ErrorHandler
'
'    strSQL = strSQL & "SELECT"
'    strSQL = strSQL & " YOUKT_UKTANTO "
'    strSQL = strSQL & "FROM"
'    strSQL = strSQL & " YOUK_TRAN "
'    strSQL = strSQL & "WHERE"
'    strSQL = strSQL & " YOUKT_MOTO_ACPTNO = '" & stCARG.CARG_ACPTNO & "'"
'
'    Set objRst = ADODB_Recordset(strSQL, objCon)
'    On Error GoTo ErrorHandler1
'    With objRst
'        stYOUK.YOUKT_UKTANTO = .Fields("YOUKT_UKTANTO")     '受付担当コード
'        .Close
'    End With
'    On Error GoTo ErrorHandler
'Exit Sub
'
'ErrorHandler1:
'    objRst.Close
'ErrorHandler:                   '↓自分の関数名
'    Call Err.Raise(Err.Number, "Select_YOUK_TRAN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'End Sub
'
''==============================================================================*
''
''       MODULE_NAME     : Select_TANT_MAST
''       機能            : 担当者マスタ情報取得
''       CREATE_DATE     : 2009/04/30            KIHARA
''       PARAM           : objCon            　　コネクション(I)
''                       : stTANT                担当者マスタ情報(O)
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub Select_TANT_MAST(objCon2 As Object, stTANT As Type_TANT_MAST, stYOUK As Type_YOUK_TRAN, strBumoc As String)
'    Dim objRst              As Object
'    Dim strSQL              As String
'    On Error GoTo ErrorHandler
'
'    strSQL = strSQL & "SELECT"
'    strSQL = strSQL & " TANTM_TANTN "
'    strSQL = strSQL & "FROM"
'    strSQL = strSQL & " TANT_MAST "
'    strSQL = strSQL & "WHERE"
'    strSQL = strSQL & " TANTM_BUMOC = '" & strBumoc & "'"
'    strSQL = strSQL & "AND"
'    strSQL = strSQL & "  TANTM_TANTC = '" & stYOUK.YOUKT_UKTANTO & "'"
'
'    Set objRst = ADODB_Recordset(strSQL, objCon2)
'    On Error GoTo ErrorHandler1
'    With objRst
'        stTANT.TANTM_TANTN = .Fields("TANTM_TANTN")     '担当者名
'        .Close
'    End With
'    On Error GoTo ErrorHandler
'Exit Sub
'
'ErrorHandler1:
'    objRst.Close
'ErrorHandler:                   '↓自分の関数名
'    Call Err.Raise(Err.Number, "Select_TANT_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'End Sub
'------↑↑↑↑-----Delete 2011/08/18 M.RYU--------↑↑↑↑--------<e>

'==============================================================================*
'
'       MODULE_NAME     : HelperPrintXXKyk
'       機能            : 申込書をプレビュー
'       CREATE_DATE     : 2009/02/01            KIHARA
'       PARAM           : strBumoc              部門コード(I)
'                       : strYardc              ヤードコード(I)
'                       : strYardcido           移動先ヤードコード(I)   'ADD 2009/05/01 KIHARA
'                       : stCARG                コンテナ情報(I)
'                       : stCONT                コントロールマスタ情報(I)
'                       : stCARG2               移動先コンテナ情報(I)
'                       : stTANT                担当者マスタ情報(I)     'ADD 2009/04/30 KIHARA
'                       : intPrevew             1:Preview 0:Print
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'MOD 2009/05/01 KIHARA
Private Sub HelperPrintXXKyk(ByVal strBUMOC As String, ByVal strYardc As String, _
                             ByVal strYardcido As String, stCARG As Type_CARG_INF, _
                             stCONT As Type_CONT_MAST, stCARG2 As Type_CARG_INF2, _
                             stTANT As Type_TANT_MAST, intPrevew As Integer)

    Dim strPath             As String 'ヤード解約ヘルパー
    Dim strPath2            As String '地図,配置
    Dim xlApp               As Object
    Dim xlBook              As Object 'ヤード解約ヘルパー
    Dim xlBook2             As Object '地図,配置
    Dim varPrintSeets       As Variant
    Dim strShName           As String
    Dim intxlCount          As Integer 'ヤード解約ヘルパーのカウンタ
    Dim intxl2Count         As Integer '地図,配置のカウンタ
    Dim intCount            As Integer 'カウンタ
    Dim intShCheck          As Integer
    On Error GoTo ErrorHandler
    'ヘルパー情報取得
    If GetHelperFileInfo(strBUMOC, strPath, strPath2) = False Then
        GoTo ErrorHandler
    End If
    ' Excelオブジェクトを生成する
    Set xlApp = CreateObject("Excel.Application")
    On Error GoTo ErrorHandler2
    Set xlBook = xlApp.Workbooks.Open(strPath & P_HELPER_ヤード解約) 'コピー元
'    If Dir(strPath2 & Format(strYardc, "000000") & ".xls") <> "" Then      'DEL 2009/05/01 KIHARA
    If Dir(strPath2 & Format(strYardcido, "000000") & ".xls") <> "" Then    'ADD 2009/05/01 KIHARA
        'ファイル有
'        Set xlBook2 = xlApp.Workbooks.Open(strPath2 & Format(strYardc, "000000") & ".xls")     'コピー先 'DEL 2009/05/01 KIHARA
        Set xlBook2 = xlApp.Workbooks.Open(strPath2 & Format(strYardcido, "000000") & ".xls")   'コピー先 'ADD 2009/05/01 KIHARA
        'コピー先
        intxl2Count = xlBook2.Sheets.Count
        '地図シートがあるかチェック
        strShName = "地図"
        intShCheck = 0
        For intCount = 1 To intxl2Count
            'If strShName = xlBook2.worksheets(intCount).Name Then                                                  'DELETE 2017/11/02 N.IMAI
            If strShName = xlBook2.Worksheets(intCount).NAME And xlBook2.Worksheets(intCount).Visible = True Then   'INSERT 2017/11/02 N.IMAI
                intShCheck = 1
                Exit For
            End If
        Next
        If intShCheck = 1 Then
            '地図シートを追加
            xlBook2.Activate
            xlBook2.Sheets("地図").Select
            intxlCount = xlBook.Sheets.Count
            xlBook2.Sheets("地図").Copy After:=xlBook.Sheets(intxlCount) 'ヤード解約ヘルパーの最後尾シートに追加
            xlBook.Sheets("地図").PageSetup.LeftHeader = ""         'INSERT 2011/08/09 M.RYU
            xlBook.Sheets("地図").PageSetup.RightHeader = ""        'INSERT 2011/08/09 M.RYU
        End If
        '配置シートがあるかチェック
        strShName = "配置"
        intShCheck = 0
        For intCount = 1 To intxl2Count
            'If strShName = xlBook2.worksheets(intCount).Name Then                                                  'DELETE 2017/11/02 N.IMAI
            If strShName = xlBook2.Worksheets(intCount).NAME And xlBook2.Worksheets(intCount).Visible = True Then   'INSERT 2017/11/02 N.IMAI
                intShCheck = 1
                Exit For
            End If
        Next
        If intShCheck = 1 Then
            '配置シートを追加
            xlBook2.Activate
            xlBook2.Sheets("配置").Select
            intxlCount = xlBook.Sheets.Count
            xlBook2.Sheets("配置").Copy After:=xlBook.Sheets(intxlCount) 'ヤード解約ヘルパーの最後尾シートに追加
            xlBook.Sheets("配置").PageSetup.LeftHeader = ""         'INSERT 2011/08/09 M.RYU
            xlBook.Sheets("配置").PageSetup.RightHeader = ""        'INSERT 2011/08/09 M.RYU
        End If
        
        xlBook2.Close False
    End If

    '基本入力シートに値を設定
'    Call SetBaseSheetKyk(strBumoc, xlBook.Worksheets("基本入力"), stCARG, stCONT, stCARG2)         'DEL 2009/04/30 KIHARA
    Call SetBaseSheetKyk(strBUMOC, xlBook.Worksheets("基本入力"), stCARG, stCONT, stCARG2, stTANT)  'ADD 2009/04/30 KIHARA
        
    '出力対象のシートを取得する
    intxlCount = xlBook.Sheets.Count
    ReDim varPrintSeets(intxlCount - 2)
    For intCount = 1 To intxlCount - 1
        varPrintSeets(intCount - 1) = xlBook.Worksheets(intCount + 1).NAME
    Next
    
    '印刷プレビュー表示
    xlApp.Visible = True
    If intPrevew = 1 Then
        'Preview
        Call xlBook.Sheets(varPrintSeets).PrintPreview
    Else
        '印刷
        xlBook.Sheets(varPrintSeets).PrintOut
    End If
    'doCmd.SelectObject acStoredProcedure, xlBook, False
    'doCmd.RunCommand acCmdPrint
    
    'EXCELファイルを閉じる
    xlBook.Close False
    On Error GoTo ErrorHandler1
    'EXCEL終了
    xlApp.DisplayAlerts = False
    xlApp.Quit

    Exit Sub

ErrorHandler2:
    xlBook.Close False
ErrorHandler1:
    xlApp.Visible = True
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "HelperPrintXX" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : GetHelperFileInfo
'       機能            : 申込書をプレビュー
'       CREATE_DATE     : 2009/02/01            KIHARA
'       PARAM           : pstrPath              ヘルパーファイルパス
'                       : pstrPath2             コンテナExcelファイルパス
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetHelperFileInfo(ByVal strBUMOC As String, pstrPath As String, pstrPath2 As String) As Boolean
    Dim INTIF_RECFB値   As String
    
    GetHelperFileInfo = False
    On Error GoTo errorHandle1
    ' ヘルパーファイルパス取得
    pstrPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB=""FVS400"" AND INTIF_RECFB = ""HELPER_PATH"""))
    If Right(pstrPath, 1) <> "\" Then
        pstrPath = pstrPath & "\"
    End If
    
    ' コンテナExcelファイルパス取得
    'Ｅｘｃｅｌ表とファイル名の取得
    If strBUMOC = "8" Then
        INTIF_RECFB値 = "TRNK_BOOK_PATH"
    Else
        INTIF_RECFB値 = "CNTN_BOOK_PATH"
    End If
    pstrPath2 = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB=""GenYardP"" AND INTIF_RECFB = """ & INTIF_RECFB値 & """"))
    If Right(pstrPath2, 1) <> "\" Then
        pstrPath2 = pstrPath2 & "\"
    End If
    GetHelperFileInfo = True
    Exit Function
errorHandle1:
    Exit Function
End Function


'==============================================================================*
'
'       MODULE_NAME     :SetBaseSheetKyk
'       機能            :基本入力シートに値を設定する
'       CREATE_DATE     : 2009/02/01            KIHARA
'       PARAM           :strBumoc                部門コード(I)
'                       :aSheet                  値を設定するシート(I)
'                       :aCARG_INF               設定するコンテナ情報(I)
'                       :aCONT_MAST              設定するコントロールマスタ情報(I)
'                       :aCARG_INF               設定する移動先コンテナ情報(I)
'                       :aTANT_MAST              設定する担当者マスタ情報(I)   'ADD 2009/04/30 KIHARA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'MOD 2009/04/30 KIHARA
Private Sub SetBaseSheetKyk(ByVal strBUMOC As String, _
                            aSheet As Object, _
                            aCARG_INF As Type_CARG_INF, _
                            aCONT_MAST As Type_CONT_MAST, _
                            aCARG_INF2 As Type_CARG_INF2, _
                            aTANT_MAST As Type_TANT_MAST)
    
    Dim strRentuse             As String
    On Error GoTo Exception
    
    'レンタル用途をセット
    Select Case strBUMOC
    Case "0"
         strRentuse = "看板"
    Case "1"
         strRentuse = "オフィス"
    Case "8"
         strRentuse = "トランク"
    Case "H"
         strRentuse = "コンテナ"
    Case "T"
         strRentuse = "横浜鋼管"
    Case Else
        '上記以外の部門コードは例外終了する
          Call MSZZ024_M10("OutPutHelperApplication", "部門コード[" & strBUMOC & "]は解約ヘルパーには対応していません。")
    End Select
    
    With aSheet
        '★コントロールマスタの値を設定
        .Range("貸主会社名").VALUE = aCONT_MAST.CONT_KAISYA
        .Range("貸主TEL").VALUE = aCONT_MAST.CONT_TEL_NO
        .Range("貸主FAX").VALUE = aCONT_MAST.CONT_FAX_NO
'        .Range("貸主担当").VALUE = aCONT_MAST.CONT_TANNM   'DEL 2009/04/30 KIHARA
        .Range("貸主担当").VALUE = aTANT_MAST.TANTM_TANTN   'ADD 2009/04/30 KIHARA
        .Range("解約専用TEL").VALUE = aCONT_MAST.CONT_CANCEL_TEL_NO             'INSERT 2018/03/10 N.IMAI
        .Range("貸主郵便番号").VALUE = aCONT_MAST.CONT_YUBINO  'ADD 2018/09/28 EGL
        .Range("貸主住所１").VALUE = aCONT_MAST.CONT_ADDR_1    'ADD 2018/09/28 EGL
        .Range("貸主住所２").VALUE = aCONT_MAST.CONT_ADDR_2    'ADD 2018/09/28 EGL
        '★コンテナ契約の値を設定
        .Range("契約者郵便番号").VALUE = aCARG_INF.USER_YUBINO
        .Range("契約者住所１").VALUE = Nz(aCARG_INF.USER_ADR_1, "") & Nz(aCARG_INF.USER_ADR_2, "")
        .Range("契約者住所２").VALUE = Nz(aCARG_INF.USER_ADR_3, "")
        '.Range("契約者住所３").VALUE = aCARG_INF.USER_ADR_3
        .Range("契約者名").VALUE = aCARG_INF.USER_NAME
        .Range("契約代表者名").VALUE = aCARG_INF.USER_TANM
        .Range("顧客コード").VALUE = aCARG_INF.CARG_UCODE
        .Range("契約者TEL").VALUE = aCARG_INF.USER_TEL
        .Range("契約者FAX").VALUE = aCARG_INF.USER_FAX
        .Range("契約者CEL").VALUE = aCARG_INF.USER_KEITAI
        .Range("契約№").VALUE = aCARG_INF.CARG_ACPTNO                          'INSERT 2011/08/18 M.RYU
        .Range("ヤード名").VALUE = aCARG_INF.YARD_NAME
        .Range("ヤードコード").VALUE = aCARG_INF.CARG_YCODE
        .Range("ヤード住所１").VALUE = aCARG_INF.YARD_ADDR_1
        .Range("ヤード住所２").VALUE = aCARG_INF.YARD_ADDR_2
        .Range("スペースコード").VALUE = aCARG_INF.CARG_NO
        .Range("レンタル用途").VALUE = strRentuse
        .Range("スペースサイズ").VALUE = Nz(aCARG_INF.CNTA_SIZE, "")
        .Range("上下段").VALUE = Nz(aCARG_INF.CNTA_STEP_NM, "")
        .Range("月額使用料").VALUE = aCARG_INF.GETSUGAKU
        .Range("他月額料").VALUE = aCARG_INF.ZAPPI
        .Range("解約日").VALUE = Format(aCARG_INF.YARD_END_DAY, "yyyy/mm/dd")
        .Range("書類作成日").VALUE = Format(Now, "yyyy/mm/dd")
        
        '--↓↓--INSERT 2011/09/30 M.RYU----<S>
'        .Range("返送期日").VALUE = Format(DateAdd("d", 21, Now), "yyyy/mm/dd") '書類作成日+21日
        If Screen.ActiveForm.NAME = Form_FVS220.NAME Then
            If Nz(Form_FVS220.txt_Hensobi) <> "" Then
                .Range("返送期日").VALUE = Format(Form_FVS220.txt_Hensobi, "yyyy/mm/dd")
            Else
                .Range("返送期日").VALUE = Format(DateAdd("d", 26, Now), "yyyy/mm/dd") '書類作成日+26日
            End If
        Else
            .Range("返送期日").VALUE = Format(DateAdd("d", 26, Now), "yyyy/mm/dd") '書類作成日+26日
        End If
        '--↑↑--INSERT 2011/09/30 M.RYU----<E>
        
        .Range("変更後契約№").VALUE = aCARG_INF2.IDO_ACPTNO                    'INSERT 2011/08/18 M.RYU
        .Range("移動先ヤード名").VALUE = Nz(aCARG_INF2.IDO_YARD_NAME, "")
        .Range("移動先ヤードコード").VALUE = Nz(aCARG_INF2.IDO_CARG_YCODE, "")
        .Range("移動先ヤード住所１").VALUE = Nz(aCARG_INF2.IDO_YARD_ADDR_1, "")
        .Range("移動先ヤード住所２").VALUE = Nz(aCARG_INF2.IDO_YARD_ADDR_2, "")
        .Range("移動先スペースコード").VALUE = Nz(aCARG_INF2.IDO_CARG_NO, "")
        .Range("移動先レンタル用途").VALUE = Nz(aCARG_INF2.IDO_YOTO_NM, "")
        .Range("移動先スペースサイズ").VALUE = Nz(aCARG_INF2.IDO_CNTA_SIZE, "")
        .Range("移動先上下段").VALUE = Nz(aCARG_INF2.IDO_CNTA_STEP_NM, "")
'        If aCARG_INF.CNTA_STEP = aCARG_INF2.IDO_CNTA_STEP Then                 'DELETE 2011/08/18 M.RYU
'            '段コードが同じ場合
'            .Range("移動先月額使用料").VALUE = Nz(aCARG_INF.GETSUGAKU, "")     'DELETE 2011/08/18 M.RYU
'            .Range("移動先他月額料").VALUE = Nz(aCARG_INF.ZAPPI, "")           'DELETE 2011/08/18 M.RYU
'        Else                                                                   'DELETE 2011/08/18 M.RYU
'            '段コードが異なる場合
        .Range("移動先月額使用料").VALUE = Nz(aCARG_INF2.IDO_GETSUGAKU, "")
        .Range("移動先他月額料").VALUE = Nz(aCARG_INF2.IDO_ZAPPI, "")
'        End If                                                                 'DELETE 2011/08/18 M.RYU
    End With
    
    Exit Sub
    
Exception:
    
    Call Err.Raise(Err.Number, "SetBaseSheetKyk" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'****************************  ended or program ********************************

