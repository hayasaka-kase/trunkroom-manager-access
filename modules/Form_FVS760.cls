VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS760"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　メンテナンス表印刷
'   プログラムＩＤ　：　FVS760
'   作　成　日　　　：  2005/08/03
'   作　成　者　　　：  T.SUZUKI
'   Ver            ：  0.0
'**********************************************
'修正履歴
'   Ver             ：0.1
'   修　正　日      ：2005/10/31
'   修　正　者      ：T.SUZUKI
'   修　正　内　容  ：1.画面の入力項目に「解約年月日」(From～To)、「解約区分」(From～To)追加
'                    2.上記1.1の追加に伴い、データの抽出条件追加
'                    3.帳票プレビューまたはエクセル出力前の必須入力チェックを変更
'                    4.帳票の出力項目追加
'**********************************************
'修正履歴
'   Ver             ：0.2
'   修　正　日      ：2005/11/30
'   修　正　者      ：T.SUZUKI
'   修　正　内　容  ：顧客名編集時、｢顧客コード｣ + ｢顧客名｣とするよう、処理を修正
'**********************************************
'修正履歴
'   Ver             ：0.4
'   修　正　日      ：2011/06/03
'   修　正　者      ：M.RYU
'   修　正　内　容  ：FVS760_W01にヤード住所１、ヤード住所２、コンテナ用途を追加
'**********************************************
'修正履歴
'   Ver             ：0.5
'   修　正　日      ：2012/04/12
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：スロープ区分を追加
'**********************************************
'修正履歴
'   Ver             ：0.6
'   修　正　日      ：2012/08/22
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：検索条件にメンテ担当者を追加
'**********************************************
'修正履歴
'   Ver             ：0.7
'   修　正　日      ：2012/12/10
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：鍵番号を表示
'**********************************************
'修正履歴
'   Ver             ：0.8
'   修　正　日      ：2013/04/05
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：段区分を追加
'**********************************************
'修正履歴
'   Ver             ：0.9
'   修　正　日      ：2014/04/08
'   修　正　者      ：MIYAMOTO
'   修　正　内　容  ：FVS760_W01に現解除番号、旧解除番号を追加
'**********************************************
'修正履歴
'   Ver             ：1.0
'   修　正　日      ：2014/04/10
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：KASE_DBへの接続方法を変更
'**********************************************
'修正履歴
'   Ver             ：1.1
'   修　正　日      ：2014/06/24
'   修　正　者      ：MIYAMOTO
'   修　正　内　容  ：FVS760_W01にスペアキー有無文言を追加、現解除番号の取得条件変更
'**********************************************
'修正履歴
'   Ver             ：1.2
'   修　正　日      ：2014/09/30
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：解約区分をデフォルトでセットするように修正01～06
'**********************************************
'修正履歴
'   Ver             ：1.3
'   修　正　日      ：2014/10/28
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：メンテ担当者コードを表示
'**********************************************
'修正履歴
'   Ver             ：1.4
'   修　正　日      ：2015/09/28
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：屋外ライン・屋内ラインを除く
'**********************************************
'修正履歴
'   Ver             ：1.5
'   修　正　日      ：2015/09/30
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：地方の場合には、旧ダイヤルに刻印の下四桁を表示
'**********************************************
'修正履歴
'   Ver             ：1.6
'   修　正　日      ：2021/06/03
'   修　正　者      ：Y.SUZUKI
'   修　正　内　容  ：出力するエクセルファイルに「物件鍵ＮＯ」を追加
'**********************************************
'修正履歴
'   Ver             ：1.7
'   修　正　日      ：2023/05/17
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：メンテナンス表に物件番号を追加
'**********************************************
'修正履歴
'   Ver             ：1.8
'   修　正　日      ：2023/06/26
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：メンテナンス表にスペアー・特記・セキュリティー会社を追加
'**********************************************
'修正履歴
'   Ver             ：2.1
'   修　正　日      ：2024/04/26
'   修　正　者      ：N.IMAI
'   修　正　内　容  ：地方でもトランクの場合は新鍵番号を出力
'**********************************************
'修正履歴
'   Ver             ：2.2
'   修　正　日      ：2026/01/26
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：データベースの接続方法を変更
'**********************************************Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const pstrCstRPT_ID      As String = "RVS760"

Private Const pstrCstKAKUNIN     As String = "確認"
Private Const pstrCstERROR       As String = "エラー"
Private Const pstrCstWK_TABLE_NM As String = "FVS760_W01"  ' 2005/10/31 ADD T.SUZUKI
Private Const pcintKASE_DB          As Integer = 1  ' 加瀬DB
'↓ INSERT 2014/06/24 MIYAMOTO
Private Const pstrSpareKeyNasiNasi  As String = "スペア無し無し"
Private Const pstrSpareKeyNasiAri   As String = "スペア無し有り"
Private Const pstrSpareKeyAriNasi   As String = "スペア有り無し"
Private Const pstrSpareKeyAriAri    As String = "スペア有り有り"
'↑ INSERT 2014/06/24 MIYAMOTO
' カレントDB
Private pobjDB                   As Database

' コンテナDB
Private pobjKONT_DB              As Database               ' 2005/10/31 ADD T.SUZUKI
Private pobjKASE_DB              As Database     ' 加瀬DB

Private Const PROG_ID = "KOMSKAN00"
'
Private USER_ID         As String
Private WK_DATED        As String
Private WK_TIMED        As String
'==============================================================================*
'
'        MODULE_NAME      :Auto_cmd_Click
'        機能             :
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Auto_cmd_Click()

    Dim strSQL         As String
    Dim lngCnt         As Long
    Dim lngYCnt        As Long
    Dim strYard        As String
    
    Dim BUMON          As String
    Dim strErr         As String
    Dim P_PROGRAM_ID   As String
    Dim strFile        As String
    Dim rsWork         As Object
    Dim objCon         As Object
    Dim objRst         As Object
    Dim strErrMsg      As String
    Dim DateEnd        As String
    Dim strCONT_BUMOC As String
    
    On Error GoTo ErrorHandler
    
    Call MSZZ003_M00(PROG_ID, "0", "")
    
    
    BUMON = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    
    USER_ID = LsGetUserName()
    WK_DATED = Format(Now(), "yyyymmdd")
    WK_TIMED = Format(Now(), "HHMMSS")
    P_PROGRAM_ID = "KOMSKAN00"
    
    strErr = "STEP-00 解約予定完了バッチ開始"
    
    If Me.cmb_TANTO_F = "" Then
    
        MsgBox ("ﾒﾝﾃ担当者ｺｰﾄﾞは必ず指定してください。")
    
        Exit Sub
    End If
    
    
    'SQLサーバー接続
    Set objCon = ADODB_Connection(BUMON)
    
    '前月末
    DateEnd = DateSerial(Year(DATE), Month(DATE), 0)
    
    '対象データ取得
    Set objRst = ADODB_Recordset(Select_Carg(DateEnd), objCon)
    lngCnt = 0
    lngYCnt = 0
    
    While Not objRst.EOF

        
        '契約ファイル更新
         Call fncUPD_CargFile(objRst.Fields("CARG_ACPTNO"), BUMON, objRst.Fields("YARD_MNT_TANTO"), objRst.Fields("CARG_KAGIICD"), objRst.Fields("CARG_KAICD"), objCon)
         
         'コンテナマスタ更新
         Call fncUPD_CNTA_MAST(objRst.Fields("CARG_YCODE"), objRst.Fields("CARG_NO"), BUMON, objRst.Fields("CARG_ACPTNO"), objCon)
       
         strYard = objRst.Fields("CARG_YCODE")
         
         lngCnt = lngCnt + 1
         objRst.MoveNext
    
    Wend
    objRst.Close
    
    Call MSZZ003_M00(PROG_ID, "8", "メンテ完了件数 : " & lngCnt)
 
    Call MSZZ003_M00(PROG_ID, "1", "")
    
    MsgBox ("自動開放終了")
    
Exit Sub
        
ErrorHandler:                    '↓自分の関数名
    Call Err.Raise(Err.Number, "KOMSKAN00" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)


End Sub
'==============================================================================*
'
'       MODULE_NAME     : 対象取得
'       MODULE_ID       : Select_Conveni
'       CREATE_DATE     :
'       PARAM           :
'       RETURN          :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Select_Carg(WK_DATED As String) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
  
    strSQL = "SELECT * "
    strSQL = strSQL & " FROM CARG_FILE "
    strSQL = strSQL & "        INNER JOIN YARD_MAST ON "
    strSQL = strSQL & "            YARD_CODE = CARG_YCODE　AND "
    strSQL = strSQL & "            YARD_TIHO_KBN <> -1 AND "
    
    If (Nz(Me.cmb_TANTO_F, "") <> "" And Nz(Me.cmb_TANTO_T, "") <> "") Then
        strSQL = strSQL & "  YARD_MNT_TANTO BETWEEN " & Me.cmb_TANTO_F & " AND " & Me.cmb_TANTO_T
    ElseIf (Nz(Me.cmb_TANTO_F, "") <> "" And Nz(Me.cmb_TANTO_T, "") = "") Then
        strSQL = strSQL & "  YARD_MNT_TANTO >= " & Me.cmb_TANTO_F
    ElseIf (Nz(Me.cmb_TANTO_F, "") = "" And Nz(Me.cmb_TANTO_T, "") <> "") Then
        strSQL = strSQL & "  YARD_MNT_TANTO <= " & Me.cmb_TANTO_T
    End If
    
    strSQL = strSQL & " WHERE  CARG_AGRE IN ('2','3') "
    strSQL = strSQL & " AND    CARG_KAGIICD IN ('10','15') "
    strSQL = strSQL & " AND    CARG_KAICD NOT IN ('03','04','05','06','07','08') "
    strSQL = strSQL & " AND    CARG_KAITANTO NOT IN ('259','275','277','278','279','280','298','541') "
    strSQL = strSQL & " AND    CARG_KYDATE = '" & WK_DATED & "' "
    strSQL = strSQL & " ORDER BY CARG_YCODE, "
    strSQL = strSQL & "          CARG_NO"
  
    Select_Carg = strSQL
Exit Function

ErrorHandler:
     Call Err.Raise(Err.Number, "Select_Carg" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :fncUPD_CargFile
'        Parameter        :strErrMsg エラーメッセージ格納
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUPD_CargFile(strCARG_ACPTNO As String, BUMON As String, MTanto As String, KagiCode As String, CARG_KAICD As String, objCon As Object) As Boolean

    Dim strSQL      As String
    Dim objRs       As Long
'    Dim objCon      As Object

    fncUPD_CargFile = False
    
    'SQLサーバー接続
'    Set objCon = ADODB_Connection(BUMON)
    
    strSQL = " UPDATE CARG_FILE "
    strSQL = strSQL & " SET CARG_CONTNA = -1, "
    strSQL = strSQL & "     CARG_AGRE = 9, "
    strSQL = strSQL & "     CARG_MNT_REQDATE = CARG_KYDATE, "                   'メンテ依頼日
    strSQL = strSQL & "     CARG_MNT_CMPDATE = '" & Format(Now, "yyyy/mm/dd") & "', "     'メンテ完了日
    strSQL = strSQL & "     CARG_MNT_TANTO   = '" & MTanto & "', "                        'メンテ担当者
    If KagiCode = "10" Then
        strSQL = strSQL & "     CARG_KEY_LENTNUM   = '2', "                               '貸出本数
    Else
        strSQL = strSQL & "     CARG_KEY_LENTNUM   = '3', "                               '貸出本数
    End If
    
    If CARG_KAICD = "4" Then
        strSQL = strSQL & "     CARG_LIQ_REQDATE   = null, "                               '`
    End If
    
    strSQL = strSQL & "     CARG_KEY_CHGTANTO = '987', "
    strSQL = strSQL & "     CARG_KEY_CHGRIYUCD = '00', "
    strSQL = strSQL & "     CARG_UPDAD = '" & Format$(DATE, "YYYYMMDD") & "', "
    strSQL = strSQL & "     CARG_UPDAJ = '" & Format$(time, "HHMMSS") & "', "
    strSQL = strSQL & "     CARG_UPDPB = '" & PROG_ID & "', "
    strSQL = strSQL & "     CARG_UPDUB =  '" & LsGetUserName() & "' "
    strSQL = strSQL & "  WHERE CARG_ACPTNO = '" & strCARG_ACPTNO & "' "
    
    objRs = ADODB_Execute(strSQL, objCon)
 
    fncUPD_CargFile = True

    Exit Function
fncUPD_CargFile_Err:
    Call Err.Raise(Err.Number, "fncUPD_CargFile" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
 
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :fncUPD_CNTA_MAST
'        Parameter        :strErrMsg エラーメッセージ格納
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUPD_CNTA_MAST(strCnta_no As String, strCNTA_CODE As String, BUMON As String, strCARG_ACPTNO As String, objCon As Object) As Boolean

    Dim strSQL      As String
    Dim objRs       As Long
'    Dim objCon      As Object

    fncUPD_CNTA_MAST = False
    
    'SQLサーバー接続
'    Set objCon = ADODB_Connection(BUMON)
    
    strSQL = " UPDATE CNTA_MAST "
    strSQL = strSQL & " SET CNTA_USE = 0, "
    strSQL = strSQL & "     CNTA_UKNO = '@" & strCARG_ACPTNO & "' "
    strSQL = strSQL & " WHERE CNTA_CODE   = '" & strCnta_no & "'"
    strSQL = strSQL & "   AND CNTA_NO = '" & strCNTA_CODE & "'"
    
    objRs = ADODB_Execute(strSQL, objCon)

 
    fncUPD_CNTA_MAST = True

fncUPD_CNTA_MAST_Exit:
    Exit Function

fncUPD_CNTA_MAST_Err:
    Call Err.Raise(Err.Number, "fncUPD_CNTA_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
   GoTo fncUPD_CNTA_MAST_Exit
End Function


'==============================================================================*
'
'       MODULE_NAME     : カンマ区切りの接続部門コードをKOMSＤＢ名配列に変換する
'       MODULE_ID       : GetKomsDbNames
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strBUMOC              接続部門(I)カンマ区切りで複数指定可
'       RETURN          : KOMSＤＢ名配列
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetKomsDbNames(ByVal strBUMOC As String) As Variant
    Dim i                   As Long
    Dim j                   As Long
    Dim varKomsDB           As Variant
    On Error GoTo ErrorHandler

    varKomsDB = Split(strBUMOC, ",")
    j = LBound(varKomsDB)
    For i = j To UBound(varKomsDB)
        If Trim(varKomsDB(i)) <> "" Then
            varKomsDB(j) = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME_" & Trim(varKomsDB(i)) & "'"), "")
            If varKomsDB(j) = "" Then
                Call MSZZ024_M10("DLookup", "テーブル[SETU_TABL]の設定不正です。")
            End If
            j = j + 1
        End If
    Next
    ReDim Preserve varKomsDB(j - 1)
    GetKomsDbNames = varKomsDB
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "GetKomsDbNames" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'        MODULE_NAME      :cmb_メンテ担当者コード_AfterUpdate
'        機能             :メンテ担当者コード AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_TANTO_F_AfterUpdate()

    Dim vCode   As Variant
    Dim vTanto  As Variant

    On Error GoTo cmb_TANTO_F_AfterUpdate_Err

    vCode = Nz(Trim(Me.cmb_TANTO_F), "")
    
    If vCode = "" Then
        Me.cmb_TANTO_F = ""
        GoTo cmb_TANTO_F_AfterUpdate_End
    End If

    'ファイルへの存在チェック
    If GetTantoData(vCode, vTanto) = False Then GoTo cmb_TANTO_F_AfterUpdate_End

    Me.txt_TANTO_F = vTanto
    Me.cmb_TANTO_F = Format$(Me.cmb_TANTO_F, "000")

    GoTo cmb_TANTO_F_AfterUpdate_End

cmb_TANTO_F_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_TANTO_F_AfterUpdate_End:
    Exit Sub
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_メンテ担当者コード_BeforeUpdate
'        機能             :メンテ担当者コード BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_TANTO_F_BeforeUpdate(Cancel As Integer)

    'ｺｰﾄﾞが空の場合担当者名欄も空欄表示
    If ControlCrear(Me.cmb_TANTO_F, Me.txt_TANTO_F) = False Then Exit Sub

    '選択された担当者データの存在チェック
    If GetTantoData(Nz(Me.cmb_TANTO_F, "")) = False Then
       MsgBox "登録されていません。", vbInformation, pstrCstERROR
       Cancel = True
       SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ControlCrear
'        機能             :ｺﾝﾎﾞﾎﾞｯｸｽの内容が空の場合、名称ﾃｷｽﾄも空にする
'        IN               :ctlCmb = ｺﾝﾎﾞﾎﾞｯｸｽ / ctlTxt = ﾃｷｽﾄ
'        OUT              :ｺﾝﾎﾞﾎﾞｯｸｽが空だったかの判断
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ControlCrear(ByRef ctlCmb As ComboBox, ByRef ctlTxt As TextBox) As Boolean

    If Nz(ctlCmb, "") = "" Then
        ctlTxt = ""
        Exit Function
    Else
        ControlCrear = True
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :Getメンテ担当者コードData
'        機能             :担当者マスタ情報取得
'        IN               :vTanto = 担当者ｺｰﾄﾞ / vTanNm = 担当者名
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetTantoData(ByRef vTanto As Variant, Optional ByRef vTanNm As Variant) As Boolean

    Dim strSQL    As String
    Dim objRs     As Recordset
    Dim pstrBumoCD As String


    GetTantoData = False
    
    ' 部門コード取得
    pstrBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")


    ' 加瀬DB接続
    If fncConnectDB(1) = False Then Exit Function

    strSQL = " SELECT TANTM_TANTC, " & Chr(13)
    strSQL = strSQL & " TANTM_TANTN " & Chr(13)
    strSQL = strSQL & "  FROM TANT_MAST " & Chr(13)
    strSQL = strSQL & " WHERE TANTM_BUMOC = '" & pstrBumoCD & "' " & Chr(13)
    strSQL = strSQL & "   AND TANT_MAST.TANTM_TANTC = " & "'" & Format$(vTanto, "000") & "'"
    Set objRs = pobjKASE_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

    If objRs.EOF Then
        ' 該当データなしのとき
        objRs.Close
        Set objRs = Nothing
        GoTo GetTantoData_Exit
    Else
        If Nz(objRs("TANTM_TANTN"), "") <> "" Then vTanNm = objRs("TANTM_TANTN")
    End If
    objRs.Close
    Set objRs = Nothing

    ' 加瀬DB切断
    If fncDisConnectDB(pcintKASE_DB) = False Then Exit Function
    
    GetTantoData = True

GetTantoData_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not pobjKASE_DB Is Nothing Then pobjKASE_DB.Close: Set pobjKASE_DB = Nothing  ' 2005/11/08 ADD T.SUZUKI
End Function

'==============================================================================*
'
'        MODULE_NAME      :cmb_メンテ担当者コード_KeyPress
'        機能             : メンテ担当者コード　KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_TANTO_F_KeyPress(KeyAscii As Integer)

    '桁数チェック追加
    Call CmKeyPress_For_Num1(Me.cmb_TANTO_F, KeyAscii, 3)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_メンテ担当者コード_AfterUpdate
'        機能             :メンテ担当者コード AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_TANTO_T_AfterUpdate()
    Dim vCode   As Variant
    Dim vTanto  As Variant

    On Error GoTo cmb_TANTO_T_AfterUpdate_Err

    vCode = Nz(Trim(Me.cmb_TANTO_T), "")
    
    If vCode = "" Then
        Me.cmb_TANTO_T = ""
        GoTo cmb_TANTO_T_AfterUpdate_End
    End If

    'ファイルへの存在チェック
    If GetTantoData(vCode, vTanto) = False Then GoTo cmb_TANTO_T_AfterUpdate_End

    Me.txt_TANTO_T = vTanto
    Me.cmb_TANTO_T = Format$(Me.cmb_TANTO_T, "000")

    GoTo cmb_TANTO_T_AfterUpdate_End

cmb_TANTO_T_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_TANTO_T_AfterUpdate_End:
    Exit Sub
End Sub


'==============================================================================*
'
'        MODULE_NAME      :cmb_メンテ担当者コード_BeforeUpdate
'        機能             :メンテ担当者コード BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_TANTO_T_BeforeUpdate(Cancel As Integer)
    'ｺｰﾄﾞが空の場合担当者名欄も空欄表示
    If ControlCrear(Me.cmb_TANTO_T, Me.txt_TANTO_T) = False Then Exit Sub

    '選択された担当者データの存在チェック
    If GetTantoData(Nz(Me.cmb_TANTO_T, "")) = False Then
       MsgBox "登録されていません。", vbInformation, pstrCstERROR
       Cancel = True
       SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_メンテ担当者コード_KeyPress
'        機能             : メンテ担当者コード　KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_TANTO_T_KeyPress(KeyAscii As Integer)

    '桁数チェック追加
    Call CmKeyPress_For_Num1(Me.cmb_TANTO_T, KeyAscii, 3)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    doCmd.Maximize
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    ' 初期処理
    Call Frm_Init
End Sub

Private Sub Form_Open(Cancel As Integer)


    Dim strCONT_BUMOC As String
    
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_Title.BackColor = RGB(255, 0, 204)
        Me.Auto_cmd.Visible = True
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KaiyakuDate_F_AfterUpdateイベント
'        MODULE_ID        :txt_KaiyakuDate_F_AfterUpdate
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuDate_F_AfterUpdate()
    Me.txt_KaiyakuDate_F = Check_Date(Me.txt_KaiyakuDate_F)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KaiyakuDate_F_BeforeUpdateイベント
'        MODULE_ID        :txt_KaiyakuDate_F_BeforeUpdate
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuDate_F_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_KaiyakuDate_F, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KaiyakuDate_F_GotFocusイベント
'        MODULE_ID        :txt_KaiyakuDate_F_GotFocus
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuDate_F_GotFocus()
    Call Date_TextFormat(Me.txt_KaiyakuDate_F, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KaiyakuDate_F_LostFocusイベント
'        MODULE_ID        :txt_KaiyakuDate_F_LostFocus
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuDate_F_LostFocus()
    Call Date_TextFormat(Me.txt_KaiyakuDate_F, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KaiyakuDate_F_KeyPressイベント
'        MODULE_ID        :txt_KaiyakuDate_F_KeyPress
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuDate_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_KaiyakuDate_F, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KaiyakuDate_T_AfterUpdateイベント
'        MODULE_ID        :txt_KaiyakuDate_T_AfterUpdate
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuDate_T_AfterUpdate()
    Me.txt_KaiyakuDate_T = Check_Date(Me.txt_KaiyakuDate_T)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KaiyakuDate_T_BeforeUpdateイベント
'        MODULE_ID        :txt_KaiyakuDate_T_BeforeUpdate
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuDate_T_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_KaiyakuDate_T, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KaiyakuDate_T_GotFocusイベント
'        MODULE_ID        :txt_KaiyakuDate_T_GotFocus
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuDate_T_GotFocus()
    Call Date_TextFormat(Me.txt_KaiyakuDate_T, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KaiyakuDate_T_LostFocusイベント
'        MODULE_ID        :txt_KaiyakuDate_T_LostFocus
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuDate_T_LostFocus()
    Call Date_TextFormat(Me.txt_KaiyakuDate_T, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KaiyakuDate_T_KeyPressイベント
'        MODULE_ID        :txt_KaiyakuDate_T_KeyPress
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuDate_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_KaiyakuDate_T, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Mnt_ReqDate_F_AfterUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mnt_ReqDate_F_AfterUpdate()
    Me.txt_Mnt_ReqDate_F = Check_Date(Me.txt_Mnt_ReqDate_F)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Mnt_ReqDate_F_BeforeUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mnt_ReqDate_F_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_Mnt_ReqDate_F, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Mnt_ReqDate_F_GotFocus
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mnt_ReqDate_F_GotFocus()
    Call Date_TextFormat(Me.txt_Mnt_ReqDate_F, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Mnt_ReqDate_F_LostFocus
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mnt_ReqDate_F_LostFocus()
    Call Date_TextFormat(Me.txt_Mnt_ReqDate_F, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Mnt_ReqDate_F_KeyPress
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mnt_ReqDate_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_Mnt_ReqDate_F, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Mnt_ReqDate_T_AfterUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mnt_ReqDate_T_AfterUpdate()
    Me.txt_Mnt_ReqDate_T = Check_Date(Me.txt_Mnt_ReqDate_T)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Mnt_ReqDate_T_BeforeUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mnt_ReqDate_T_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_Mnt_ReqDate_T, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Mnt_ReqDate_T_GotFocus
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mnt_ReqDate_T_GotFocus()
    Call Date_TextFormat(Me.txt_Mnt_ReqDate_T, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Mnt_ReqDate_T_LostFocus
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mnt_ReqDate_T_LostFocus()
    Call Date_TextFormat(Me.txt_Mnt_ReqDate_T, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Mnt_ReqDate_T_KeyPress
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mnt_ReqDate_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_Mnt_ReqDate_T, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_F_AfterUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_F_AfterUpdate()

    Dim strCode   As String
    Dim strCARG   As String

    On Error GoTo cmb_YCODE_F_AfterUpdate_Err

    strCode = Nz(Me.cmb_YCODE_F, "")
    strCARG = ""
    If strCode = "" Then Exit Sub

    If Trim(strCode) <> "" Then
        ' ヤード名取得
        strCARG = fncGetYardData(strCode)
    End If

    Me.cmb_YCODE_F = Format(Me.cmb_YCODE_F, "000000")
    Me.txt_YCODE_F = strCARG
    Me.cmb_YCODE_T.SetFocus  ' ヤードコード(TO)にフォーカスセット

    Exit Sub

cmb_YCODE_F_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_F_BeforeUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_F_BeforeUpdate(Cancel As Integer)

    Dim strCode  As String
    Dim strCARG  As String

    On Error GoTo cmb_YCODE_F_BeforeUpdate_Err

    strCode = Nz(Me.cmb_YCODE_F, "")
    If strCode = "" Then
        Me.txt_YCODE_F = ""
        Exit Sub
    End If

    ' ヤード名取得
    strCARG = fncGetYardData(strCode)
    If strCARG = "" Then
        MsgBox "該当するヤードマスタがありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_YCODE_F_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_F_KeyPress
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_YCODE_F, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_T_AfterUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_T_AfterUpdate()

    Dim varCode   As Variant
    Dim strCARG   As String
    Dim strWhere  As String

    On Error GoTo cmb_YCODE_T_AfterUpdate_Err

    varCode = Nz(Me.cmb_YCODE_T, "")
    strCARG = ""
    If varCode = "" Then Exit Sub

    If Trim(varCode) <> "" Then
        strWhere = " YARD_CODE = " & varCode & ""
        strCARG = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strWhere), "")
    End If

    Me.cmb_YCODE_T = Format(Me.cmb_YCODE_T, "000000")
    Me.txt_YCODE_T = strCARG
    Me.chk_KaiPage.SetFocus  ' ヤード毎に改ページチェックボックスにフォーカスセット

    Exit Sub

cmb_YCODE_T_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_T_BeforeUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_T_BeforeUpdate(Cancel As Integer)

    Dim strCode  As String
    Dim strCARG  As String

    On Error GoTo cmb_YCODE_T_BeforeUpdate_Err

    strCode = Nz(Me.cmb_YCODE_T, "")
    If strCode = "" Then
        Me.txt_YCODE_T = ""
        Exit Sub
    End If

    ' ヤード名取得
    strCARG = fncGetYardData(strCode)
    If strCARG = "" Then
        MsgBox "該当するヤードマスタがありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_YCODE_T_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_T_KeyPress
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_YCODE_T, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_F_AfterUpdateイベント
'        MODULE_ID        :cmb_KAIYAKU_F_AfterUpdate
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_F_AfterUpdate()

    Dim strCode   As String
    Dim strCARG   As String
    Dim strWhere  As String

    On Error GoTo cmb_KAIYAKU_F_AfterUpdate_Err

    strCARG = ""
    strCode = Nz(Me.cmb_KAIYAKU_F, "")
    If strCode = "" Then Exit Sub

    If Trim(strCode) <> "" Then
        ' 解約区分名取得
        strCARG = fncGetKaiyakuData("089", strCode)
    End If

    Me.cmb_KAIYAKU_F = Format(Me.cmb_KAIYAKU_F, "00")
    Me.txt_KAIYAKU_F = strCARG  ' 解約区分名をセット
    Me.cmb_KAIYAKU_T.SetFocus   ' 解約区分(TO)にフォーカスセット

    Exit Sub

cmb_KAIYAKU_F_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_F_BeforeUpdateイベント
'        MODULE_ID        :cmb_KAIYAKU_F_BeforeUpdate
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_F_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strCARG  As String

    On Error GoTo cmb_KAIYAKU_F_BeforeUpdate_Err

    varCode = Nz(Me.cmb_KAIYAKU_F, "")
    If varCode = "" Then
        Me.txt_KAIYAKU_F = ""
        Exit Sub
    End If

    ' 解約区分名取得
    strCARG = fncGetKaiyakuData("089", varCode)
    If strCARG = "" Then
        MsgBox "該当する解約区分がありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_KAIYAKU_F_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_F_KeyPressイベント
'        MODULE_ID        :cmb_KAIYAKU_F_KeyPress
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_KAIYAKU_F, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_T_AfterUpdateイベント
'        MODULE_ID        :cmb_KAIYAKU_T_AfterUpdate
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_T_AfterUpdate()

    Dim strCode   As String
    Dim strCARG   As String

    On Error GoTo cmb_KAIYAKU_T_AfterUpdate_Err

    strCARG = ""
    strCode = Nz(Me.cmb_KAIYAKU_T, "")
    If strCode = "" Then Exit Sub

    If Trim(strCode) <> "" Then
        ' 解約区分名取得
        strCARG = fncGetKaiyakuData("089", strCode)
    End If

    Me.cmb_KAIYAKU_T = Format(Me.cmb_KAIYAKU_T, "00")
    Me.txt_KAIYAKU_T = strCARG  ' 解約区分名をセット
'    Me.cmb_KAIYAKU_T.SetFocus      ' 解約区分(TO)にフォーカスセット

    Exit Sub

cmb_KAIYAKU_T_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_T_BeforeUpdateイベント
'        MODULE_ID        :cmb_KAIYAKU_T_BeforeUpdate
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_T_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strCARG  As String

    On Error GoTo cmb_KAIYAKU_F_BeforeUpdate_Err

    varCode = Nz(Me.cmb_KAIYAKU_T, "")
    If varCode = "" Then
        Me.txt_KAIYAKU_T = ""
        Exit Sub
    End If

    ' 解約区分名取得
    strCARG = fncGetKaiyakuData("089", varCode)
    If strCARG = "" Then
        MsgBox "該当する解約区分がありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_KAIYAKU_F_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_T_KeyPressイベント
'        MODULE_ID        :cmb_KAIYAKU_T_KeyPress
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_KAIYAKU_T, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_Preview_Click
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Preview_Click()

    Dim strMsg As String
    Dim lngRc  As Long

    On Error GoTo cmd_Preview_Click_Err

    ' 日付入力チェック
    If fncInputCheck = False Then GoTo cmd_Preview_Click_Exit

    strMsg = "印刷内容を画面に出力します。よろしいですか？"
    lngRc = MsgBox(strMsg, vbQuestion + vbYesNo, pstrCstKAKUNIN)
    If lngRc = vbNo Then GoTo cmd_Preview_Click_Exit

    doCmd.Hourglass True

    ' 帳票出力用データ作成
    If FncCreateWork = False Then GoTo cmd_Preview_Click_Exit

    doCmd.OpenReport "RVS760", acPreview, "", ""

cmd_Preview_Click_Exit:
    doCmd.Hourglass False
    Exit Sub

cmd_Preview_Click_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo cmd_Preview_Click_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_Excel_Click
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()

    On Error GoTo cmd_Excel_Click_Err

    ' 日付入力チェック
    If fncInputCheck = False Then GoTo cmd_Excel_Click_Exit

    If (MsgBox("Excel出力を実行しますか？", vbYesNo + vbQuestion, pstrCstKAKUNIN) = vbYes) Then

        doCmd.Hourglass True

        ' 帳票出力用データ作成
        If FncCreateWork = True Then
           doCmd.OutputTo acQuery, "FVS760_Q01", "MicrosoftExcel(*.xls)", "", True, ""
        End If
    End If

cmd_Excel_Click_Exit:
    doCmd.Hourglass False
    Exit Sub

cmd_Excel_Click_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo cmd_Excel_Click_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_End_Click
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_End_Click()

    On Error Resume Next
    doCmd.Close acForm, Me.FormName
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Frm_Init
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Frm_Init()

    With Me
' 2005/10/31 ADD T.SUZUKI Start
        Me.txt_KaiyakuDate_F = ""  ' 解約年月日(From)
        Me.txt_KaiyakuDate_T = ""  ' 解約年月日(To)
' 2005/10/31 ADD T.SUZUKI End
        .txt_Mnt_ReqDate_F = ""    ' メンテナンス依頼日(From)
        .txt_Mnt_ReqDate_T = ""    ' メンテナンス依頼日(To)
        .cmb_YCODE_F = Null        ' ヤードコード(From)
        .txt_YCODE_F = Null        ' ヤード名(From)
        .cmb_YCODE_T = Null        ' ヤードコード(To)
        .txt_YCODE_T = Null        ' ヤード名(To)
        .chk_KaiPage.VALUE = -1    ' 改ページ
        .chk_cmpDate = -1          ' 完了日を除く
' 2005/10/31 ADD T.SUZUKI Start
'        .cmb_KAIYAKU_F = ""
'        .txt_KAIYAKU_F = ""
'        .cmb_KAIYAKU_T = ""
'        .txt_KAIYAKU_T = ""

        '2014/09/30 M.HONDA INS
        .cmb_KAIYAKU_F = "01"
        cmb_KAIYAKU_F_AfterUpdate
        .cmb_KAIYAKU_T = "06"
        cmb_KAIYAKU_T_AfterUpdate
        '2014/09/30 M.HONDA INS


' 2005/10/31 ADD T.SUZUKI End
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CMCheck_Date
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CMCheck_Date(ByVal vDATE As Variant)
  
    Dim sYear As String

    CMCheck_Date = Null

    If Nz(vDATE, "") = "" Then Exit Function
    If IsDate(vDATE) = False Then Exit Function

    sYear = Format(vDATE, "yyyy")
    If CmLenb(sYear) <> 4 Then Exit Function

    CMCheck_Date = vDATE
End Function

'==============================================================================*
'
'        MODULE_NAME      :入力チェック
'        MODULE_ID        :fncInputCheck
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheck() As Boolean

    fncInputCheck = False

    With Me
' 2005/10/31 REP T.SUZUKI Start
        If (Nz(.txt_KaiyakuDate_F, "") = "" And Nz(.txt_KaiyakuDate_T, "") = "") And _
           (Nz(.txt_Mnt_ReqDate_F, "") = "" And Nz(.txt_Mnt_ReqDate_T, "") = "") Then
            MsgBox "解約年月日またはメンテナンス依頼日を入力して下さい。", vbInformation, pstrCstERROR
            .txt_KaiyakuDate_F.SetFocus
            Exit Function
        End If
'        If (Nz(.txt_Mnt_ReqDate_F, "") = "" And Nz(.txt_Mnt_ReqDate_T, "") = "") Then
'            MsgBox "メンテ依頼日Fromまたはメンテ開始日Toを入力して下さい。", vbInformation, pstrCstERROR
'            .txt_Mnt_ReqDate_F.SetFocus
'            Exit Function
'        End If
' 2005/10/31 REP T.SUZUKI End
    End With

    fncInputCheck = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルの削除
'        MODULE_ID        :fncDeleteWorkTable
'        Parameter        :strTableNm = テーブル名
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDeleteWorkTable(ByVal strTableNm As String) As Boolean

    Dim strSQL As String

    On Error GoTo fncDeleteWorkTable_Err

    fncDeleteWorkTable = False

    ' ワークテーブル存在チェック
    If fncTableExist(strTableNm) = True Then

        ' ワークテーブルを削除
        pobjDB.TableDefs.Delete strTableNm
        Select Case Err.Number
            Case 0
            Case 3265
                Err.Clear
            Case Else
                MsgBox "subDropWorkTable:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
                Err.Clear
        End Select
    End If

    fncDeleteWorkTable = True

    Exit Function

fncDeleteWorkTable_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :ACCESSテーブル存在チェック
'        MODULE_ID        :fncTableExist
'        Parameter        :IN strTableName = テーブル名
'        RETURN           :True=存在する False=存在しない
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncTableExist(strTableName As String) As Boolean

    Dim tdf As TableDef

    fncTableExist = False

    For Each tdf In pobjDB.TableDefs
        If tdf.NAME = strTableName Then
            fncTableExist = True
            Exit For
        End If
    Next tdf
End Function

' 2005/10/31 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :ワークテーブルのデータ削除
''        MODULE_ID        :fncCreateWork
''        CREATE_DATE      :2005/08/03
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncDeleteWork() As Boolean
'
'    Dim strSQL As String
'
'    On Error GoTo fncDeleteWork_Err
'
'    fncDeleteWork = False
'
'    ' ワークテーブルの内容を削除
'    strSQL = "DELETE FROM FVS760_W01 ; "
'    pobjDB.Execute strSQL, dbFailOnError
'
'    fncDeleteWork = True
'
'    Exit Function
'
'fncDeleteWork_Err:
'    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
'    Err.Clear
'End Function
' 2005/10/31 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルの作成
'        MODULE_ID        :fncCreateWorkTable
'        Parameter        :strTableNm = テーブル名
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWorkTable(ByVal strTableNm As String) As Boolean

    Dim tdfNew    As TableDef
    Dim intCount  As Integer

    On Error GoTo fncCreateWorkTable_Err

    fncCreateWorkTable = False

    Set tdfNew = pobjDB.CreateTableDef(strTableNm)

    With tdfNew
        .Fields.Append .CreateField("REPORT_ID", dbText, 50)        ' レポートID
        .Fields.Append .CreateField("BUMOM_BUMOC", dbText, 3)       ' 部門コード
        .Fields.Append .CreateField("BUMOM_BUMON", dbText, 50)      ' 部門名称
        .Fields.Append .CreateField("CARG_YCODE", dbLong)           ' ヤードコード
        .Fields.Append .CreateField("YARD_NAME", dbText, 36)        ' ヤード名
        .Fields.Append .CreateField("YARD_ADDR_1", dbText, 36)      ' ヤード住所1           'INSERT 2011/06/03 M.RYU
        .Fields.Append .CreateField("YARD_ADDR_2", dbText, 36)      ' ヤード住所2           'INSERT 2011/06/03 M.RYU
        .Fields.Append .CreateField("CARG_NO", dbLong)              ' コンテナ番号
        .Fields.Append .CreateField("CARG_USAGE", dbText, 20)       ' コンテナ用途          'INSERT 2011/06/03 M.RYU
        .Fields.Append .CreateField("USER_NAME", dbText, 50)        ' 顧客コード + 顧客名
        .Fields.Append .CreateField("CARG_HOSYICD", dbText, 20)     ' 保証区分
        .Fields.Append .CreateField("CARG_KAICD", dbText, 20)       ' 解約区分
        .Fields.Append .CreateField("CARG_KAGIICD", dbText, 20)     ' 鍵(タイプ)
        .Fields.Append .CreateField("NEW_BUIL_KAGI_NO", dbText, 10)     ' 物件鍵ＮＯ（建物用の鍵番号） ' INSERT 2021/06/03 Y.SUZUKI
        .Fields.Append .CreateField("CARG_KAGIISPARE", dbText, 20)  ' 鍵(スペア)
        .Fields.Append .CreateField("CARG_KEY_RETDATE", dbText, 5)  ' 鍵返却日(MM/DD)
        .Fields.Append .CreateField("CARG_MNT_CMPDATE", dbText, 5)  ' メンテ完了日(MM/DD)
        .Fields.Append .CreateField("CARG_MNT_MEMO", dbText, 60)    ' メンテメモ
        .Fields.Append .CreateField("CARG_NIRCD", dbText, 10)       ' 荷入区分
        .Fields.Append .CreateField("CARG_SYOCD", dbText, 10)       ' 処分区分
        .Fields.Append .CreateField("CARG_HOSCD", dbText, 10)       ' 補修区分
        .Fields.Append .CreateField("CARG_AGRE", dbText, 10)        ' 契約状態
        .Fields.Append .CreateField("CARG_MNT_TANTO", dbText, 36)   ' メンテ担当者名
        .Fields.Append .CreateField("CARG_MNT_REQDATE", dbDate)     ' メンテ依頼日
        .Fields.Append .CreateField("RCPT_SLOPE", dbText, 36)       ' スロープ         'INS 2012/04/12 M.HONDA
        .Fields.Append .CreateField("TANT_TANCD", dbText, 10)        ' 担当者コード     'INS 2014/10/28 M.HONDA
        .Fields.Append .CreateField("TANTM_TANTN", dbText, 36)      ' メンテ担当者     'INS 2012/08/22 M.HONDA
        .Fields.Append .CreateField("CARG_KAIJB", dbText, 10)       ' 鍵解除番号       'INS 2012/12/10 M.HONDA
        .Fields.Append .CreateField("CNTA_STEP", dbText, 10)        ' 段区分           'INS 2013/04/05 M.HONDA
        
        .Fields.Append .CreateField("CNTA_SIZE", dbText, 10)
        
        .Fields.Append .CreateField("AKKNT_KAINO", dbText, 10)      ' 現解除番号        'INSERT 2014/04/08 MIYAMOTO
        .Fields.Append .CreateField("RCPT_KAINO", dbText, 10)       ' 旧解除番号        'INSERT 2014/04/08 MIYAMOTO
        .Fields.Append .CreateField("SPAREKEY_UMU", dbText, 100)    ' スペアキー有無文言    'INSERT 2014/06/24 MIYAMOTO
        .Fields.Append .CreateField("TINHM_BUKHC", dbText, 100)
        
        .Fields.Append .CreateField("SPARE_KEY", dbText, 100)
        .Fields.Append .CreateField("YARD_SECUCOMP", dbText, 100)
        .Fields.Append .CreateField("YARD_TOKKI", dbText, 255)
        .Fields.Append .CreateField("YARD_END_DAY", dbText, 100)
        .Fields.Append .CreateField("YARD_RENTEND_DAY", dbText, 100)
       
         For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount
    End With

    ' テーブルの追加
    pobjDB.TableDefs.Append tdfNew

    fncCreateWorkTable = True

    Exit Function

fncCreateWorkTable_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルへ対象データのInsert
'        MODULE_ID        :fncCreateWork
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function FncCreateWork() As Boolean

    Dim strSQL      As String
    Dim strWhere    As String

    Dim strSQLWK    As String

    Dim objRs       As Recordset
    Dim objWk       As Recordset

    Dim varBumoCD   As Variant  ' 部門ｺｰﾄﾞ
    Dim varBumoNm   As Variant  ' 部門名

    Dim varCode       As Variant
    Dim varName       As Variant
    Dim strBumoCD     As String
    Dim strDataSource As String

    Dim strIntiFileKey2 As String                   'INSERT 2014/06/24 MIYAMOTO
    
    On Error GoTo fncCreateWork_Err

    FncCreateWork = False

    doCmd.SetWarnings False

    ' 部門コード取得
    strBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))  '2014/06/10 M.HONDA INS

    ' DB接続
    Set pobjDB = CurrentDb

' 2005/10/31 ADD T.SUZUKI Start
    ' コンテナDB接続
    If fncConnectDB = False Then GoTo fncCreateWork_Err
' 2005/10/31 ADD T.SUZUKI End

    ' ワークテーブルの内容を削除
    If fncDeleteWorkTable(pstrCstWK_TABLE_NM) = False Then GoTo fncCreateWork_Err  ' 2005/10/31 ADD T.SUZUKI
'    If fncDeleteWork = False Then GoTo fncCreateWork_Err                          ' 2005/10/31 DEL T.SUZUKI

    ' ワークテーブルを作成
    If fncCreateWorkTable(pstrCstWK_TABLE_NM) = False Then GoTo fncCreateWork_Err  ' 2005/10/31 ADD T.SUZUKI

    ' 部門名取得
    Call subGetBumonName(pobjDB, varBumoCD, varBumoNm)

    ' コンテナ契約ファイルよりデータ取得
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " CARG_YCODE      , " & Chr(13)       ' ヤードコード
    strSQL = strSQL & " YARD_NAME       , " & Chr(13)       ' ヤード名
    strSQL = strSQL & " YARD_ADDR_1     , " & Chr(13)       ' ヤード住所１              'INSERT 2011/06/03 M.RYU
    strSQL = strSQL & " YARD_ADDR_2     , " & Chr(13)       ' ヤード住所２              'INSERT 2011/06/03 M.RYU
    strSQL = strSQL & " CARG_NO         , " & Chr(13)       ' コンテナ番号
    strSQL = strSQL & " NAME_MAST.NAME_NAME AS CARG_USAGE, " & Chr(13) 'コンテナ用途    'INSERT 2011/06/03 M.RYU
    strSQL = strSQL & " YARD_TIHO_KBN   , " & Chr(13)       ' 地方区分   '2015/09/30 M.HONDA INS
    strSQL = strSQL & " USER_NAME       , " & Chr(13)       ' 顧客名
    strSQL = strSQL & " CARG_HOSYICD    , " & Chr(13)       ' 保証区分(200)
    strSQL = strSQL & " CARG_KAICD      , " & Chr(13)       ' 解約区分(089)
    strSQL = strSQL & " CARG_KAGIICD    , " & Chr(13)       ' 鍵区分(060)
    strSQL = strSQL & " CARG_KEY_RETDATE, " & Chr(13)       ' 鍵返却日(書式：MM/DD)
    strSQL = strSQL & " CARG_MNT_CMPDATE, " & Chr(13)       ' 完了日(書式：MM/DD)
    strSQL = strSQL & " CARG_MNT_MEMO   , " & Chr(13)       ' メンテメモ
    strSQL = strSQL & " CARG_NIRCD      , " & Chr(13)       ' 荷入
    strSQL = strSQL & " CARG_SYOCD      , " & Chr(13)       ' 処分
    strSQL = strSQL & " CARG_HOSCD      , " & Chr(13)       ' 補修
    strSQL = strSQL & " CARG_AGRE       , " & Chr(13)       ' 契約状態
    strSQL = strSQL & " CARG_MNT_TANTO  , " & Chr(13)       ' メンテ担当者コード
    strSQL = strSQL & " CARG_MNT_REQDATE, " & Chr(13)       ' メンテ依頼日
    strSQL = strSQL & " CARG_KYDATE     , " & Chr(13)       ' 解約年月日
    strSQL = strSQL & " CARG_UCODE      , " & Chr(13)       ' 顧客コード
    strSQL = strSQL & " CARG_SPAREKEY   , " & Chr(13)       ' スペアキー
    strSQL = strSQL & " CARG_MNT_TANTO  , " & Chr(13)       ' メンテ担当者コード
    strSQL = strSQL & " YARD_MNT_TANTO  , " & Chr(13)       ' メンテ担当者コード
    strSQL = strSQL & " NAME_294.NAME_NAME AS RCPT_SLOPE , " & Chr(13)       ' スロープ貸出
    strSQL = strSQL & " TANTM_TANTN , " & Chr(13)       ' メンテ担当者ｺｰﾄﾞ    ' 2012/08/22 M.HONDA
    strSQL = strSQL & " CARG_KAIJB ,  " & Chr(13)       ' 鍵番号              ' 2012/12/10 M.HONDA
    strSQL = strSQL & " CNTA_STEP,     " & Chr(13)       ' 段区分              ' 2013/04/05 M.HONDA
    strSQL = strSQL & " CNTA_SIZE     " & Chr(13)
    strSQL = strSQL & " ,   ( " & Chr(13)
    strSQL = strSQL & "         SELECT      AKKNT.AKKNT_KAINO " & Chr(13)
    strSQL = strSQL & "         FROM        ( " & Chr(13)
    strSQL = strSQL & "                         SELECT      AKKNT_YARD " & Chr(13)
    strSQL = strSQL & "                         ,           AKKNT_CTNO " & Chr(13)
    strSQL = strSQL & "                         ,           AKKNT_KTYPE " & Chr(13)
    strSQL = strSQL & "                         ,           AKKNT_HATUD " & Chr(13)
    strSQL = strSQL & "                         ,           ROW_NUMBER() OVER(PARTITION BY AKKNT_YARD, AKKNT_CTNO, AKKNT_KTYPE ORDER BY AKKNT_HATUD DESC) AS SEQ " & Chr(13)
    strSQL = strSQL & "                         ,           AKKNT_KAINO " & Chr(13)
    strSQL = strSQL & "                         FROM        AKKN_TRAN " & Chr(13)
    strSQL = strSQL & "                         WHERE       AKKNT_YARD = CNTA_MAST.CNTA_CODE " & Chr(13)
    strSQL = strSQL & "                         AND         AKKNT_CTNO = CNTA_MAST.CNTA_NO " & Chr(13)
    strSQL = strSQL & "                         AND         AKKNT_KTYPE = CNTA_MAST.CNTA_KEY_TYPE " & Chr(13)
    strSQL = strSQL & "                         AND         AKKNT_STYPE = '4'" '2021/06/03 add 部屋鍵のみ対象とする
    strSQL = strSQL & "                     ) AKKNT " & Chr(13)
    strSQL = strSQL & "         WHERE       AKKNT.SEQ = 1 " & Chr(13)
    strSQL = strSQL & "     ) AKKNT_KAINO " & Chr(13)
    strSQL = strSQL & " ,   RCPT_KAINO " & Chr(13)
    strSQL = strSQL & " ,   CNTA_MAST.CNTA_KEY_TYPE AS NEW_KAGIICD" & Chr(13)
    strSQL = strSQL & " ,   ( " & Chr(13)
    strSQL = strSQL & "         SELECT      AKKNT.AKKNT_KAINO " & Chr(13)
    strSQL = strSQL & "         FROM        ( " & Chr(13)
    strSQL = strSQL & "                         SELECT      AKKNT_YARD " & Chr(13)
    strSQL = strSQL & "                         ,           AKKNT_CTNO " & Chr(13)
    strSQL = strSQL & "                         ,           AKKNT_KTYPE " & Chr(13)
    strSQL = strSQL & "                         ,           AKKNT_HATUD " & Chr(13)
    strSQL = strSQL & "                         ,           ROW_NUMBER() OVER(PARTITION BY AKKNT_YARD, AKKNT_CTNO, AKKNT_KTYPE ORDER BY AKKNT_HATUD DESC) AS SEQ " & Chr(13)
    strSQL = strSQL & "                         ,           AKKNT_KAINO " & Chr(13)
    strSQL = strSQL & "                         FROM        AKKN_TRAN " & Chr(13)
    strSQL = strSQL & "                         WHERE       AKKNT_YARD = CNTA_MAST.CNTA_CODE " & Chr(13)
    strSQL = strSQL & "                         AND         AKKNT_CTNO = CNTA_MAST.CNTA_NO " & Chr(13)
    strSQL = strSQL & "                         AND         AKKNT_KTYPE = CNTA_MAST.CNTA_KEY_TYPE " & Chr(13)
    strSQL = strSQL & "                         AND         AKKNT_STYPE = '6'" '物件鍵のみ対象とする
    strSQL = strSQL & "                     ) AKKNT " & Chr(13)
    strSQL = strSQL & "         WHERE       AKKNT.SEQ = 1 " & Chr(13)
    strSQL = strSQL & "     ) NEW_BUIL_KAGI_NO, " & Chr(13)
    strSQL = strSQL & "     TINHM_BUKHC ,         "
    
    strSQL = strSQL & "     NAME_061.NAME_NAME SPARE_KEY , "
    strSQL = strSQL & "     NAME_062.NAME_NAME YARD_SECUCOMP , "
    strSQL = strSQL & "     YARD_TOKKI, "
    strSQL = strSQL & "     CONVERT(NVARCHAR,  YARD_END_DAY, 111) AS YARD_END_DAY,  "
    strSQL = strSQL & "     CONVERT(NVARCHAR,  YARD_RENTEND_DAY, 111) AS YARD_RENTEND_DAY  "
    strSQL = strSQL & " FROM CARG_FILE                                              "
    strSQL = strSQL & " LEFT JOIN YARD_MAST                                         "
    strSQL = strSQL & "        ON CARG_FILE.CARG_YCODE = YARD_MAST.YARD_CODE        "
    strSQL = strSQL & " LEFT JOIN                                                   "
    strSQL = strSQL & strDataSource & ".dbo.TANT_MAST ON                            "
    strSQL = strSQL & " TANTM_BUMOC = '" & strBumoCD & "' AND                       "
    strSQL = strSQL & " TANTM_TANTC = YARD_MAST.YARD_MNT_TANTO                      "
    strSQL = strSQL & " LEFT JOIN ( CNTA_MAST                                       "
    strSQL = strSQL & "      LEFT JOIN NAME_MAST                                    "
    strSQL = strSQL & "             ON CNTA_MAST.CNTA_USAGE = NAME_MAST.NAME_CODE   "
    strSQL = strSQL & "             AND CNTA_USAGE NOT IN ('31')               "  '2015/09/28 M.HONDA INS
    strSQL = strSQL & "      ) ON CARG_FILE.CARG_YCODE = CNTA_MAST.CNTA_CODE        "
    strSQL = strSQL & "       AND CARG_FILE.CARG_NO = CNTA_MAST.CNTA_NO             "
    strSQL = strSQL & " LEFT JOIN USER_MAST                                         "
    strSQL = strSQL & "        ON CARG_FILE.CARG_UCODE = USER_MAST.USER_CODE        "
    strSQL = strSQL & " Left Join "
    strSQL = strSQL & " RCPT_TRAN ON "
    strSQL = strSQL & " CARG_FILE.CARG_ACPTNO = RCPT_TRAN.RCPT_CARG_ACPTNO "
    
    strSQL = strSQL & " LEFT JOIN                                                   "
    strSQL = strSQL & strDataSource & ".dbo.TINH_MAST ON                            "
    strSQL = strSQL & " TINHM_BUMOC = '" & strBumoCD & "' AND                       "
    strSQL = strSQL & " TINHM_TINTC = CARG_YCODE                          "
    
    strSQL = strSQL & "      Left Join "
    strSQL = strSQL & "        NAME_MAST NAME_061 "
    strSQL = strSQL & "    ON  NAME_061.NAME_ID = '061' "
    strSQL = strSQL & "    AND NAME_061.NAME_CODE = YARD_SPARE_KEY "
    
    strSQL = strSQL & "   Left Join "
    strSQL = strSQL & "        NAME_MAST NAME_062 "
    strSQL = strSQL & "     ON  NAME_062.NAME_ID = '062' "
    strSQL = strSQL & "     AND NAME_062.NAME_CODE = YARD_SECUCOMP, "
    
    strSQL = strSQL & " NAME_MAST NAME_294 "
    strSQL = strSQL & " WHERE  NAME_MAST.NAME_ID = 86  "    'コンテナ用途区分
    strSQL = strSQL & " and NAME_294.NAME_ID = 294 "
    strSQL = strSQL & " and NAME_294.NAME_CODE = ISNULL(RCPT_SLOPE,0) "
    
'--------↑↑↑↑---------UPDATE 2011/06/03 M.RYU------↑↑↑↑------------<E>

    ' Where句生成
    strWhere = fncCreateWhere

    If strWhere <> "" Then
        strSQL = strSQL & strWhere
    End If

    ' ソート
    strSQL = strSQL & " ORDER BY CARG_YCODE, " & Chr(13)       ' ヤードコード
    strSQL = strSQL & "          CARG_NO, " & Chr(13)          ' コンテナ番号
    strSQL = strSQL & "          CARG_MNT_REQDATE " & Chr(13)  ' メンテ依頼日

    ' コンテナ契約ファイル読込
    Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)   ' 2005/10/31 ADD T.SUZUKI
'    Set objRs = pobjDB.OpenRecordset(strSQL, dbReadOnly)                                        ' 2005/10/31 DEL T.SUZUKI

    If objRs.RecordCount = 0 Then
        MsgBox "該当するデータがありません。", vbInformation, pstrCstERROR
        objRs.Close
        Set objRs = Nothing
        Exit Function
    End If

    ' ワークテーブルへ書き込み
    strSQLWK = "SELECT * FROM FVS760_W01 "
    Set objWk = pobjDB.OpenRecordset(strSQLWK, dbOpenDynaset, dbAppendOnly)

    Do Until objRs.EOF
        With objWk
            .AddNew
            .Fields("REPORT_ID") = pstrCstRPT_ID                        ' 帳票ID
            .Fields("BUMOM_BUMOC") = varBumoCD                          ' 部門コード
            .Fields("BUMOM_BUMON") = varBumoNm                          ' 部門名
            .Fields("CARG_YCODE") = objRs.Fields("CARG_YCODE")          ' ヤードコード
            .Fields("YARD_NAME") = objRs.Fields("YARD_NAME")            ' ヤード名
            .Fields("CARG_NO") = objRs.Fields("CARG_NO")                ' コンテナ番号

            .Fields("YARD_ADDR_1") = objRs.Fields("YARD_ADDR_1")        ' ヤード住所１  'INSERT 2011/06/03 M.RYU
            .Fields("YARD_ADDR_2") = objRs.Fields("YARD_ADDR_2")        ' ヤード住所２  'INSERT 2011/06/03 M.RYU
            .Fields("CARG_USAGE") = objRs.Fields("CARG_USAGE")          ' コンテナ用途  'INSERT 2011/06/03 M.RYU

' 2005/10/31 REP T.SUZUKI Start
            ' 顧客名
            If Nz(objRs.Fields("CARG_UCODE"), "") <> "" Then
                .Fields("USER_NAME") = Format(objRs.Fields("CARG_UCODE"), "000000") & vbCrLf & objRs.Fields("USER_NAME")     ' 2005/11/30 ADD T.SUZUKI
'                .Fields("USER_NAME") = Format(objRs.Fields("CARG_UCODE"), "000000") & Space(1) & objRs.Fields("USER_NAME")  ' 2005/11/30 DEL T.SUZUKI
            End If
'            .Fields("USER_NAME") = objRs.Fields("USER_NAME")            ' 顧客名
' 2005/10/31 REP T.SUZUKI End

            ' 保証区分取得
            If Nz(objRs.Fields("CARG_HOSYICD"), "") <> "" Then
                varCode = fncGetName("200", objRs.Fields("CARG_HOSYICD"))
                .Fields("CARG_HOSYICD") = varCode                       ' 保証区分
            End If

            ' 解約区分取得
            If Nz(objRs.Fields("CARG_KAICD"), "") <> "" Then
                varCode = fncGetName("089", objRs.Fields("CARG_KAICD"))
                .Fields("CARG_KAICD") = varCode                         ' 解約区分
            End If

            ' 鍵区分取得(タイプ)
            If Nz(objRs.Fields("CARG_KAGIICD"), "") <> "" Then
                varCode = fncGetName("060", objRs.Fields("CARG_KAGIICD"))
                .Fields("CARG_KAGIICD") = varCode                       ' 鍵区分
            End If

            .Fields("NEW_BUIL_KAGI_NO") = objRs.Fields("NEW_BUIL_KAGI_NO")    ' 物件鍵ＮＯ（建物用の鍵番号）  INSERT 2021/06/03 Y.SUZUKI

' 2005/10/31 ADD T.SUZUKI Start
            ' スペアキー取得
            If Nz(objRs.Fields("CARG_SPAREKEY"), "") <> "" Then
                varCode = fncGetName("093", objRs.Fields("CARG_SPAREKEY"))
                .Fields("CARG_KAGIISPARE") = varCode                    ' スペアキー
            End If
' 2005/10/31 ADD T.SUZUKI End

            .Fields("CARG_KEY_RETDATE") = Format(objRs.Fields("CARG_KEY_RETDATE"), "MM/DD") ' 鍵返却日(MM/DD)
            .Fields("CARG_MNT_CMPDATE") = Format(objRs.Fields("CARG_MNT_CMPDATE"), "MM/DD") ' 完了日(MM/DD)
            .Fields("CARG_MNT_MEMO") = Nz(objRs.Fields("CARG_MNT_MEMO"), "")                ' メンテメモ

            ' 荷入区分
            Select Case Nz(objRs.Fields("CARG_NIRCD"), "")
                Case ""
                    .Fields("CARG_NIRCD") = "有・無"
                Case "0", "1"
                    varCode = fncGetName("093", objRs.Fields("CARG_NIRCD"))
                    .Fields("CARG_NIRCD") = varCode
                Case Else
                    .Fields("CARG_NIRCD") = Null
            End Select

            ' 処分区分
            Select Case Nz(objRs.Fields("CARG_SYOCD"), "")
                Case ""
                    .Fields("CARG_SYOCD") = "有・無"
                Case "0", "1"
                    varCode = fncGetName("093", objRs.Fields("CARG_SYOCD"))
                    .Fields("CARG_SYOCD") = varCode
                Case Else
                    .Fields("CARG_SYOCD") = Null
            End Select

            ' 補修区分
            Select Case Nz(objRs.Fields("CARG_HOSCD"), "")
                Case ""
                    .Fields("CARG_HOSCD") = "有・無"
                Case "0", "1"
                    varCode = fncGetName("093", objRs.Fields("CARG_HOSCD"))
                    .Fields("CARG_HOSCD") = varCode
                Case Else
                    .Fields("CARG_HOSCD") = Null
            End Select

            ' 貸出
            Select Case Nz(objRs.Fields("CARG_AGRE"), "")
                Case 9
                    varCode = fncGetName("094", 1)
                    .Fields("CARG_AGRE") = varCode
                Case Else
                    .Fields("CARG_AGRE") = "可・否"
            End Select
'            If objRs.Fields("CARG_AGRE") = "9" Then
'                .Fields("CARG_AGRE") = "可"
'            Else
'                .Fields("CARG_AGRE") = "可・否"
'            End If

            ' メンテ担当者コード
            varName = ""
            If Nz(objRs.Fields("CARG_MNT_TANTO"), "") <> "" Then
                varName = fncGetTantoData(objRs.Fields("CARG_MNT_TANTO"))
                .Fields("CARG_MNT_TANTO") = varName
            Else
                .Fields("CARG_MNT_TANTO") = Null
            End If

            .Fields("CARG_MNT_REQDATE") = objRs.Fields("CARG_MNT_REQDATE")  ' メンテ依頼日
            
            ' スロープ
            .Fields("RCPT_SLOPE") = Nz(objRs.Fields("RCPT_SLOPE"), "")
            
            .Fields("TANTM_TANTN") = objRs.Fields("TANTM_TANTN")
            .Fields("TANT_TANCD") = "(" & objRs.Fields("YARD_MNT_TANTO") & ") "
            
            .Fields("CARG_KAIJB") = objRs.Fields("CARG_KAIJB") '' 2012/12/10 M.HONDA
            
            '' 2013/04/05 M.HONDA
            ' 段区分取得
            If Nz(objRs.Fields("CNTA_STEP"), "") <> "" Then
                varCode = fncGetName("090", objRs.Fields("CNTA_STEP"))
                .Fields("CNTA_STEP") = varCode                       ' 鍵区分
            End If
            
'↓ INSERT 2014/04/08 MIYAMOTO
'↓ UPDATE 2014/06/24 MIYAMOTO
'            .Fields("RCPT_KAINO") = objRs.Fields("RCPT_KAINO")          '旧解除番号
'            .Fields("AKKNT_KAINO") = objRs.Fields("AKKNT_KAINO")        '現解除番号
            
            '地方の場合には、刻印下四桁を表示する。
            'If Nz(objRs.Fields("YARD_TIHO_KBN")) = -1 Then                             'DELETE 2024/04/26 N.IMAI
            If Nz(objRs.Fields("YARD_TIHO_KBN")) = -1 And varBumoCD = "H" Then          'INSERT 2024/04/26 N.IMAI
                .Fields("RCPT_KAINO") = "旧：" & Right(objRs.Fields("CARG_KAIJB"), 4)
                .Fields("AKKNT_KAINO") = "新："
                
            Else

                If Nz(objRs.Fields("RCPT_KAINO")) = "" Then
                    .Fields("RCPT_KAINO") = ""                                      '旧解除番号
                Else
                    .Fields("RCPT_KAINO") = "旧：" & objRs.Fields("RCPT_KAINO")     '旧解除番号
                End If
                If Nz(objRs.Fields("AKKNT_KAINO")) = "" Then
                    .Fields("AKKNT_KAINO") = ""                                     '現解除番号
                Else
                    .Fields("AKKNT_KAINO") = "新：" & objRs.Fields("AKKNT_KAINO")   '現解除番号
                End If
            End If
'↑ UPDATE 2014/06/24 MIYAMOTO
'↑ INSERT 2014/04/08 MIYAMOTO
'↓ INSERT 2014/06/24 MIYAMOTO
            If objRs.Fields("CARG_KAGIICD") = 15 And objRs.Fields("NEW_KAGIICD") = 15 Then
                strIntiFileKey2 = pstrSpareKeyNasiNasi
            ElseIf objRs.Fields("CARG_KAGIICD") = 15 And objRs.Fields("NEW_KAGIICD") <> 15 Then
                strIntiFileKey2 = pstrSpareKeyNasiAri
            ElseIf objRs.Fields("CARG_KAGIICD") <> 15 And objRs.Fields("NEW_KAGIICD") = 15 Then
                strIntiFileKey2 = pstrSpareKeyAriNasi
            ElseIf objRs.Fields("CARG_KAGIICD") <> 15 And objRs.Fields("NEW_KAGIICD") <> 15 Then
                strIntiFileKey2 = pstrSpareKeyAriAri
            End If
            .Fields("SPAREKEY_UMU") = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='" & pstrCstRPT_ID & "' AND INTIF_RECFB = '" & strIntiFileKey2 & "'"))
'↑ INSERT 2014/06/24 MIYAMOTO

            .Fields("TINHM_BUKHC") = objRs.Fields("TINHM_BUKHC")
            
            .Fields("SPARE_KEY") = objRs.Fields("SPARE_KEY")
            .Fields("YARD_SECUCOMP") = objRs.Fields("YARD_SECUCOMP")
            .Fields("YARD_TOKKI") = objRs.Fields("YARD_TOKKI")
            .Fields("YARD_END_DAY") = objRs.Fields("YARD_END_DAY")

            .Fields("YARD_RENTEND_DAY") = objRs.Fields("YARD_RENTEND_DAY")
            .Fields("CNTA_SIZE") = objRs.Fields("CNTA_SIZE")

            .UPDATE
        End With

        objRs.MoveNext
    Loop

' 2005/10/31 ADD T.SUZUKI Start
    ' コンテナDB切断
    If fncConnectDB = False Then GoTo fncCreateWork_Err
' 2005/10/31 ADD T.SUZUKI End

    FncCreateWork = True

fncCreateWork_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    If Not objWk Is Nothing Then
        objWk.Close
        Set objWk = Nothing
    End If

    Exit Function

fncCreateWork_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo fncCreateWork_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :Where句生成
'        MODULE_ID        :fncCreateWhere
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWhere() As String

    Dim strWhere             As String

    Dim strKaiyakuDate_From  As String  ' 2005/10/31 ADD T.SUZUKI 解約年月日(From)
    Dim strKaiyakuDate_To    As String  ' 2005/10/31 ADD T.SUZUKI 解約年月日(To)
    Dim strReqDate_From      As String  ' 2005/10/31 ADD T.SUZUKI メンテナンス依頼日(From)
    Dim strReqDate_To        As String  ' 2005/10/31 ADD T.SUZUKI メンテナンス依頼日(To)

    With Me
' 2005/10/31 ADD T.SUZUKI Start
        ' *** 解約年月日の大小チェック ***
        Call ChkDateFromTo(.txt_KaiyakuDate_F, .txt_KaiyakuDate_T)
' 2005/10/31 ADD T.SUZUKI End

        ' *** メンテ依頼日の大小チェック ***
        Call ChkDateFromTo(.txt_Mnt_ReqDate_F, .txt_Mnt_ReqDate_T)

        ' *** ヤードコードの大小チェック ***
        Call ChkYARDFromTo(.cmb_YCODE_F, .txt_YCODE_F, .cmb_YCODE_T, .txt_YCODE_T)

' 2005/10/31 ADD T.SUZUKI Start
        ' *** 解約区分の大小チェック ***
        Call ChkYARDFromTo(.cmb_KAIYAKU_F, .txt_KAIYAKU_F, .cmb_KAIYAKU_T, .txt_KAIYAKU_T)

        ' 解約年月日、メンテナンス依頼日を格納
        strKaiyakuDate_From = Nz(.txt_KaiyakuDate_F, "") ' 解約年月日(From)
        strKaiyakuDate_To = Nz(.txt_KaiyakuDate_T, "")   ' 解約年月日(To)

        strReqDate_From = Nz(.txt_Mnt_ReqDate_F, "")     ' メンテナンス依頼日(From)
        strReqDate_To = Nz(.txt_Mnt_ReqDate_T, "")       ' メンテナンス依頼日(To)

        ' 解約年月日
        If (Nz(strKaiyakuDate_From, "") <> "" And Nz(strKaiyakuDate_To, "") <> "") Then
            ' 解約年月日(From)：入力
            ' 解約年月日(To)  ：入力
            strWhere = " AND CARG_KYDATE BETWEEN '" & strKaiyakuDate_From & "' AND '" & strKaiyakuDate_To & "' "
'            strWHERE = " WHERE CARG_KYDATE BETWEEN '" & strKaiyakuDate_From & "' AND '" & strKaiyakuDate_To & "' "

        ElseIf (Nz(strKaiyakuDate_From, "") <> "" And Nz(strKaiyakuDate_To, "") = "") Then
            ' 解約年月日(From)：入力
            ' 解約年月日(To)  ：未入力
            strWhere = " AND CARG_KYDATE >= '" & strKaiyakuDate_From & "' "
'            strWHERE = " WHERE CARG_KYDATE >= '" & strKaiyakuDate_From & "' "

        ElseIf (Nz(strKaiyakuDate_From, "") = "" And Nz(strKaiyakuDate_To, "") <> "") Then
            ' 解約年月日(From)：未入力
            ' 解約年月日(To)  ：入力
            strWhere = " AND CARG_KYDATE <= '" & strKaiyakuDate_To & "' "
'            strWHERE = " WHERE CARG_KYDATE <= '" & strKaiyakuDate_To & "' "
        End If
' 2005/10/31 ADD T.SUZUKI End
' 2005/10/31 REP T.SUZUKI Start
        If (Nz(strKaiyakuDate_From, "") <> "" Or Nz(strKaiyakuDate_To, "") <> "") And _
           (Nz(strReqDate_From, "") = "" And Nz(strReqDate_To, "") = "") Then
            ' 解約年月日        ：入力
            ' メンテナンス依頼日：未入力
            ' → メンテナンス依頼日がNULLのデータを対象とする
            strWhere = strWhere & " AND CARG_MNT_REQDATE IS NULL "
'            strWHERE = " WHERE CARG_MNT_REQDATE IS NULL "

        Else
            ' メンテナンス依頼日
            If (Nz(strReqDate_From, "") <> "" And Nz(strReqDate_To, "") <> "") Then
                ' メンテナンス依頼日(From)：入力
                ' メンテナンス依頼日(To)  ：入力
                strWhere = strWhere & " AND CARG_MNT_REQDATE BETWEEN '" & strReqDate_From & "' AND '" & strReqDate_To & "' "
'                strWHERE = " WHERE CARG_MNT_REQDATE BETWEEN '" & strReqDate_From & "' AND '" & strReqDate_To & "' "
    
            ElseIf (Nz(strReqDate_From, "") <> "" And Nz(strReqDate_To, "") = "") Then
                ' メンテナンス依頼日(From)：入力
                ' メンテナンス依頼日(To)  ：未入力
                strWhere = strWhere & " AND CARG_MNT_REQDATE >= '" & strReqDate_From & "' "
'                strWHERE = " WHERE CARG_MNT_REQDATE >= '" & strReqDate_From & "' "

            ElseIf (Nz(strReqDate_From, "") = "" And Nz(strReqDate_To, "") <> "") Then
                ' メンテナンス依頼日(From)：未入力
                ' メンテナンス依頼日(To)  ：入力
                strWhere = strWhere & " AND CARG_MNT_REQDATE <= '" & strReqDate_To & "' "
'                strWHERE = " WHERE CARG_MNT_REQDATE <= '" & strReqDate_To & "' "
            End If
        End If
'        If (Nz(.txt_Mnt_ReqDate_F, "") <> "" And Nz(.txt_Mnt_ReqDate_T, "") <> "") Then
'            ' メンテナンス依頼日(From)：入力
'            ' メンテナンス依頼日(To)  ：入力
'            strWHERE = " WHERE CARG_MNT_REQDATE BETWEEN CDate(#" & .txt_Mnt_ReqDate_F & "#) And CDATE(#" & .txt_Mnt_ReqDate_T & "#)"
'
'        ElseIf (Nz(.txt_Mnt_ReqDate_F, "") <> "" And Nz(.txt_Mnt_ReqDate_T, "") = "") Then
'            ' メンテナンス依頼日(From)：入力
'            ' メンテナンス依頼日(To)  ：未入力
'            strWHERE = " WHERE CARG_MNT_REQDATE >= CDate(#" & .txt_Mnt_ReqDate_F & "#)"
'
'        ElseIf (Nz(.txt_Mnt_ReqDate_F, "") = "" And Nz(.txt_Mnt_ReqDate_T, "") <> "") Then
'            ' メンテナンス依頼日(From)：未入力
'            ' メンテナンス依頼日(To)  ：入力
'            strWHERE = " WHERE CARG_MNT_REQDATE <= CDate(#" & txt_Mnt_ReqDate_T & "#)"
'        End If
' 2005/10/31 REP T.SUZUKI End

        ' ヤードコード
        If (Nz(.cmb_YCODE_F, "") <> "" And Nz(.cmb_YCODE_T, "") <> "") Then
            ' ヤードコード(From)：入力
            ' ヤードコード(To)  ：入力
            strWhere = strWhere & " AND CARG_YCODE BETWEEN " & .cmb_YCODE_F & " AND " & .cmb_YCODE_T

        ElseIf (Nz(.cmb_YCODE_F, "") <> "" And Nz(.cmb_YCODE_T, "") = "") Then
            ' ヤードコード(From)：入力
            ' ヤードコード(To)  ：未入力
            strWhere = strWhere & " AND CARG_YCODE >= " & .cmb_YCODE_F

        ElseIf (Nz(.cmb_YCODE_F, "") = "" And Nz(.cmb_YCODE_T, "") <> "") Then
            ' ヤードコード(From)：未入力
            ' ヤードコード(To)  ：入力
            strWhere = strWhere & " AND CARG_YCODE <= " & .cmb_YCODE_T
        End If

        ' 完了を除く
        If .chk_cmpDate.VALUE = -1 Then
            ' 完了を除くチェックボックスがONのとき
            strWhere = strWhere & " AND CARG_AGRE <> 9 "  ' 完了以外
        End If

        '' 2012/08/22 M.HONDA START
        ' メンテ担当者
        If (Nz(.cmb_TANTO_F, "") <> "" And Nz(.cmb_TANTO_T, "") <> "") Then
            strWhere = strWhere & " AND YARD_MNT_TANTO BETWEEN " & .cmb_TANTO_F & " AND " & .cmb_TANTO_T
        ElseIf (Nz(.cmb_TANTO_F, "") <> "" And Nz(.cmb_TANTO_T, "") = "") Then
            strWhere = strWhere & " AND YARD_MNT_TANTO >= " & .cmb_TANTO_F
        ElseIf (Nz(.cmb_TANTO_F, "") = "" And Nz(.cmb_TANTO_T, "") <> "") Then
            strWhere = strWhere & " AND YARD_MNT_TANTO <= " & .cmb_TANTO_T
        End If
        '' 2012/08/22 M.HONDA END

        ' 完了を除く
        If .chk_cmpDate.VALUE = -1 Then
            ' 完了を除くチェックボックスがONのとき
            strWhere = strWhere & " AND CARG_AGRE <> 9 "  ' 完了以外
        End If



' 2005/10/31 ADD T.SUZUKI Start
        ' 解約区分
        If (Nz(.cmb_KAIYAKU_F, "") <> "" And Nz(.cmb_KAIYAKU_T, "") <> "") Then
            ' ヤードコード(From)：入力
            ' ヤードコード(To)  ：入力
            strWhere = strWhere & " AND CARG_KAICD BETWEEN " & .cmb_KAIYAKU_F & " AND " & .cmb_KAIYAKU_T

        ElseIf (Nz(.cmb_KAIYAKU_F, "") <> "" And Nz(.cmb_KAIYAKU_T, "") = "") Then
            ' ヤードコード(From)：入力
            ' ヤードコード(To)  ：未入力
            strWhere = strWhere & " AND CARG_KAICD >= " & .cmb_KAIYAKU_F

        ElseIf (Nz(.cmb_KAIYAKU_F, "") = "" And Nz(.cmb_KAIYAKU_T, "") <> "") Then
            ' ヤードコード(From)：未入力
            ' ヤードコード(To)  ：入力
            strWhere = strWhere & " AND CARG_KAICD <= " & .cmb_KAIYAKU_T
        End If
' 2005/10/31 ADD T.SUZUKI End
    End With
    fncCreateWhere = strWhere
End Function

'==============================================================================*
'
'        MODULE_NAME      :日付入力チェック（From/Toを逆に入力した時の対応）
'        MODULE_ID        :ChkDateFromTo
'        IN               :
'        CREATE_DATE      :2005/08/03
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ChkDateFromTo(ByRef ctlTextF As TextBox, ByRef ctlTextT As TextBox)

    Dim strDayF As String
    Dim strDayT As String
    Dim strSwap As String

    If ctlTextF <> "" And ctlTextT <> "" Then
        strDayF = Format(ctlTextF, "yyyy/mm/dd")  ' メンテ依頼日(From)
        strDayT = Format(ctlTextT, "yyyy/mm/dd")  ' メンテ依頼日(To)

        If strDayT <> "" And strDayF > strDayT Then
            ' 大小入力エラー時
            strSwap = strDayF
            strDayF = strDayT
            strDayT = strSwap

            ' 参照元へ再セット
            ctlTextF = strDayF
            ctlTextT = strDayT
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ヤードコード入力チェック（From/Toを逆に入力した時の対応）
'        MODULE_ID        :ChkYARDFromTo
'        IN               :
'        CREATE_DATE      :2005/08/03
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ChkYARDFromTo(ByRef ctlcmbF As ComboBox, ByRef ctlTextF As TextBox, _
                          ByRef ctlcmbT As ComboBox, ByRef ctlTextT As TextBox)

    Dim strYARDF       As String
    Dim strYARDT       As String
    Dim strSwap        As String

    strSwap = ""
    If Trim(ctlcmbF) <> "" And Trim(ctlcmbT) <> "" Then
        strYARDF = Format(ctlcmbF, "000000")  ' ヤードコード(From)
        strYARDT = Format(ctlcmbT, "000000")  ' ヤードコード(To)

        If strYARDF > strYARDT Then
            ' 大小入力エラー時
            ' *** ヤードコード
            strSwap = strYARDF
            strYARDF = strYARDT
            strYARDT = strSwap

            '　参照元へ再セット
            ctlcmbF = strYARDF
            ctlcmbT = strYARDT

            ' 名称
            ' *** ヤード名
            strSwap = ""
            strYARDF = ""
            strYARDT = ""

            strSwap = ctlTextF
            strYARDF = ctlTextT
            strYARDT = strSwap

            '　参照元へ再セット
            ctlTextF = strYARDF
            ctlTextT = strYARDT
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :DateCheck
'        機能             :日付入力ﾃﾞｰﾀのチェック
'        IN               :ctlTxt = 処理対象ﾃｷｽﾄ / Cancel  = 状態区分
'        OUT              :strMsg = ｴﾗｰﾒｯｾｰｼﾞ / Chk_FrmControl = 入力不備有無
'        CREATE_DATE      :2005/08/03
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub dateCheck(ByRef ctlTxt As TextBox, Cancel As Integer)

    Dim vCode As Variant
    Dim vDATE As Variant

    On Error GoTo DateCheck_Err

    vCode = Nz(ctlTxt, "")

    ' 何も入力されていない時は、処理をしない
    If vCode = "" Then GoTo DateCheck_Exit
    vCode = Check_Date(vCode)
    vCode = CMCheck_Date(vCode)

    ' 入力チェック
    Select Case CmBCheck_Date(vCode)
        Case 2
             ' 桁数エラー
             MsgBox "桁数オーバーしています。", vbExclamation, pstrCstERROR
             Cancel = True
             SendKeys "{ESC}"
             GoTo DateCheck_Exit

        Case 3, 9
             MsgBox "日付を正しく入力してください。", vbExclamation, pstrCstERROR
             Cancel = True
             SendKeys "{ESC}"
             GoTo DateCheck_Exit
        Case Else
    End Select

DateCheck_Exit:
    Exit Sub

DateCheck_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :日付項目のチェックを一括でする
'        MODULE_ID        :
'        CREATE_DATE      :
'
'引数
'oText 対象項目
'iLen  桁数(省略時は、"yyyy/mm/dd"として10バイト)
'返り値は、
'２：桁数が溢れている時
'３：日付と判断できない時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Date(ByVal oText As Variant, Optional ByVal ilen As Variant) As Integer

    Dim sYear   As String
    Dim intLen  As Integer

    On Error GoTo CmBCheck_Date_Err

    CmBCheck_Date = 2

    If IsMissing(ilen) = False Then
       intLen = Nz(CmNumeric(ilen), 0)
    Else
       intLen = 10
    End If

    oText = Nz(oText, "")
    ' 桁数チェック
    If CmLenb(oText) > intLen Then GoTo CmBCheck_Date_Exit
    CmBCheck_Date = 3
    If Nz(oText, "") = "" Then GoTo CmBCheck_Date_Exit

    ' 論理チェック
    If IsDate(oText) = False Then GoTo CmBCheck_Date_Exit

    sYear = Format(oText, "yyyy")
    If CmLenb(sYear) <> 4 Then GoTo CmBCheck_Date_Exit

    CmBCheck_Date = 0

CmBCheck_Date_Exit:
    Exit Function

CmBCheck_Date_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    CmBCheck_Date = 9
End Function

'==============================================================================*
'
'        MODULE_NAME      :Date_TextFormat
'        機能             :日付データのフォーマット
'        IN               :
'        OUT              :
'        CREATE_DATE      :2005/08/03
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Date_TextFormat(ByRef ctlTxt As TextBox, ByVal lngMode As Long)
    
    Dim vCode As Variant
    
    vCode = Nz(ctlTxt, "")
    If vCode <> "" Then
        If lngMode = 0 Then
            ctlTxt = Format(vCode, "yy/mm/dd")
        Else
            ctlTxt = Format(vCode, "yyyy/mm/dd")
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncGetYardData
'        機能             :ヤードマスタよりヤード名取得
'        IN               :strCode = ヤードコード
'        OUT              :ヤード名
'        CREATE_DATE      :2005/08/03
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetYardData(ByVal strCode As String) As String

    Dim strWhere As String
    Dim strCARG  As String

    On Error GoTo fncGetYardData_Err

    fncGetYardData = ""

    strWhere = strWhere & " dbo_YARD_MAST.YARD_CODE = " & strCode & ""
    strCARG = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strWhere), "")

    fncGetYardData = strCARG

    Exit Function

fncGetYardData_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGetKaiyakuData
'        機能             :名称マスタより解約区分名取得
'        IN               :strID   = ID
'                          strCode = CODE
'        OUT              :解約区分名
'        CREATE_DATE      :2005/08/03
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetKaiyakuData(ByVal strNAME_ID As String, ByVal strNAME_CODE As String) As String

    Dim strWhere As String
    Dim strCARG  As String

    On Error GoTo fncGetKaiyakuData_Err

    fncGetKaiyakuData = ""

    ' 名称マスタより解約区分名を取得
    strWhere = " NAME_ID = '" & strNAME_ID & "'"
    strWhere = strWhere & " AND NAME_CODE = " & strNAME_CODE
    strCARG = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")

    fncGetKaiyakuData = strCARG

    Exit Function

fncGetKaiyakuData_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :部門コード/部門名称の取得
'        MODULE_ID        :subGetBumonName
'        IN               :objDB = DB接続 /strBumonCd = 部門ｺｰﾄﾞ /strBumonNm = 部門名
'        CREATE_DATE      :2005/08/03
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subGetBumonName(ByRef objDb As Database, ByRef varBumonCd As Variant, ByRef varBumonNm As Variant)

    Dim strSQL  As String
    Dim objRs   As Recordset

    On Error GoTo subGetBumonName_Err

    strSQL = "SELECT "
    strSQL = strSQL & "BUMO_MAST.BUMOM_BUMOC, "  ' 部門コード
    strSQL = strSQL & "BUMO_MAST.BUMOM_BUMON "   ' 部門名称
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "dbo_CONT_MAST INNER JOIN BUMO_MAST ON "
    strSQL = strSQL & "dbo_CONT_MAST.CONT_BUMOC = BUMO_MAST.BUMOM_BUMOC; "
    Set objRs = objDb.OpenRecordset(strSQL, dbReadOnly)

    With objRs
        If Not .EOF Then
            varBumonCd = .Fields("BUMOM_BUMOC")
            varBumonNm = .Fields("BUMOM_BUMON")
        Else
            varBumonCd = Null
            varBumonNm = Null
        End If
    End With

subGetBumonName_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    Exit Sub

subGetBumonName_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo subGetBumonName_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :名称取得
'        MODULE_ID        :fncGetName
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetName(strNAME_ID As String, varNAME_CODE As Variant) As String

    Dim strWhere  As String
    Dim strName   As String

    On Error GoTo fncGetName_Err

    fncGetName = ""

    If Nz(varNAME_CODE, "") = "" Then Exit Function

    ' 名称マスタより略称を取得
    strWhere = " NAME_ID = '" & strNAME_ID & "'"
    strWhere = strWhere & " AND NAME_CODE = " & varNAME_CODE

    strName = ""
    strName = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")

    If strName <> "" Then
        fncGetName = strName
    End If

fncGetName_Exit:
    Exit Function

fncGetName_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo fncGetName_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGetTantoData
'        機能             :担当者マスタ情報取得
'        IN               :vTanto = 担当者ｺｰﾄﾞ / vTanNm = 担当者名
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetTantoData(ByRef vTanto As Variant) As String

    Dim strSQL  As String
    Dim objRs   As Recordset

    fncGetTantoData = ""

    strSQL = "SELECT "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTC, "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTN "
    strSQL = strSQL & "FROM TANT_MAST INNER JOIN "
    strSQL = strSQL & "dbo_CONT_MAST ON "
    strSQL = strSQL & "TANT_MAST.TANTM_BUMOC = dbo_CONT_MAST.CONT_BUMOC "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTC = " & "'" & Format$(vTanto, "000") & "'"

    Set objRs = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
    If objRs.EOF Then
        objRs.Close
        Set objRs = Nothing
        Exit Function
    Else
        If Nz(objRs("TANTM_TANTN"), "") <> "" Then fncGetTantoData = objRs("TANTM_TANTN")
    End If

fncGetTantoData_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    Exit Function

fncGetTantoData_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo fncGetTantoData_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :DB接続
'        MODULE_ID        :fncConnectDB
'        CREATE_DATE      :2005/10/19
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncConnectDB(Optional conn As Integer = 0) As Boolean

    Dim strBumoCD     As String

    Dim strConnect    As String
    Dim strDataSource As String

    On Error GoTo fncConnectDB_Err

    fncConnectDB = False

    ' 部門コード取得
    strBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")

    If conn = 0 Then

        ' -----------------------------------------------------------------------------------------------------------
        ' コンテナDB接続
        strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumoCD & "'"), "")
        
        If strDataSource = "" Then
            MsgBox "SETU_TABLの設定が不正です。", vbExclamation, pstrCstRPT_ID
            Set pobjKONT_DB = Nothing
        Else
            Set pobjKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(strBumoCD))
        End If
        ' -----------------------------------------------------------------------------------------------------------
    Else
    
        ' -----------------------------------------------------------------------------------------------------------
        ' 加瀬DB接続
        strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))

        If strDataSource = "" Then
            MsgBox "SETU_TABLの設定が不正です。", vbExclamation, pstrCstRPT_ID
            Set pobjKASE_DB = Nothing
            Exit Function
        End If
    
        Set pobjKASE_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00())
        ' -----------------------------------------------------------------------------------------------------------
    
    End If
    fncConnectDB = True

    Exit Function

fncConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :DB切断
'        MODULE_ID        :fncDisConnectDB
'        CREATE_DATE      :2005/10/19
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDisConnectDB(Optional conn As Integer = 0) As Boolean

    On Error GoTo fncDisConnectDB_Err

    fncDisConnectDB = False

    If conn = 0 Then
        ' コンテナDB
        pobjKONT_DB.Close
        Set pobjKONT_DB = Nothing
    Else
        pobjKASE_DB.Close
        Set pobjKASE_DB = Nothing
    End If
    fncDisConnectDB = True

    Exit Function

fncDisConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function




'****************************  ended or program ********************************
