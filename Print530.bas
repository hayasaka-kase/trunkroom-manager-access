Attribute VB_Name = "Print530"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : ご紹介連絡表出力処理
'
'        PROGRAM_NAME    : 予約受付検索制御
'        PROGRAM_ID      : ご紹介連絡表出力処理
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2005/09/12
'        CERATER         : T.SUZUKI
'        Ver             : 0.0
'
'==============================================================================*
'*******************************************************************************
 '修正履歴
'   修　正　日　　　：  2005/11/29
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：
'                       使用料算出方法を変更
'   Ver             ：  0.1
'*******************************************************************************
'修正履歴
'   修　正　日　　　：  2005/12/14
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：
'                       1.テーブル(FVS530_W01)作成時、【CNTA_MISETTEI】項目追加
'                       2.使用料算出を変更、価格表テーブルのデータが存在しないとき、
'                         「未設定」文字列を出力
'   Ver             ：  0.2
'*******************************************************************************
'   修　正　日　　　：  2007/04/06
'   修　正　者　　　：  tajima
'   修　正　内　容　：  営業開始日当日印刷の注意文言付加
'   Ver             ：  0.3
'*******************************************************************************
'   修　正　日　　　：  2017/04/01
'   修　正　者　　　：  K.SATO
'   修　正　内　容　：  ご紹介連絡表のメール送信機能を追加
'   Ver             ：  0.4
'*******************************************************************************
'   更　新　日　　　：  2017/12/20
'   更　新　者　　　：  N.IMAI
'                   ：  名称取得不可の場合エラーが発生するバグに対応
'   Ver             ：  0.5
'*******************************************************************************
'   更　新　日　　　：  2018/02/15
'   更　新　者　　　：  N.IMAI
'                   ：  保証会社新プラン対応
'   Ver             ：  0.6
'*******************************************************************************
'   更　新　日　　　：  2020/07/27
'   更　新　者　　　：  S.WATANABE
'                   ：  料金改定対応
'   Ver             ：  0.7
'*******************************************************************************
'   更　新　日　　　：  2023/05/30
'   更　新　者　　　：  N.IMAI
'                   ：  メールURLに部屋番号を追加
'   Ver             ：  0.8
'*******************************************************************************
'   修　正　日　　　：  2025/12/04
'   修　正　者　　　：  T.KAWABATA
'   修　正　内　容　：  YARD_MASTカラム追加によるカラム上限エラーの解消 https://kaseit.backlog.com/view/IT-1208
'   Ver　　         ：  0.9
'*******************************************************************************
'   修　正　日　　　：  2025/12/15
'   修　正　者　　　：  T.KAWABATA
'   修　正　内　容　：  紹介メール送信の不具合対応
'   Ver　　         ：  1.0
'*******************************************************************************

Option Compare Database
Option Explicit

Private Const pstrCstERROR        As String = "エラー"
Private Const pstrCstFRM_ID       As String = "FVS530"
Private Const pstrCstRPT_ID       As String = "RVS530"
Public Const pintCstRenraku       As Integer = -999
Private Const pstrCstTEL          As String = "TEL："
Private Const pstrCstFAX          As String = "FAX："
Private Const pstrCstKeitai       As String = "携帯："
Private Const PROG_ID             As String = "Print530"

Public Const pintCstViewNew       As Integer = 1
Public Const pintCstViewOnly      As Integer = 2
Public Const pintCstViewAll       As Integer = 3

Private Const pcstrBIKO_Message1 As String = "として毎月"           ' 2005/11/29 ADD T.SUZUKI
Private Const pcstrBIKO_Message2 As String = "円が含まれます。"     ' 2005/11/29 ADD T.SUZUKI

Private Const pstrCstWK_TABLE_NM  As String = "FVS530_W01"

Public Const pintCstFVS530        As Integer = 0   ' ご紹介連絡表印刷指示画面から遷移
Public Const pintCstOther         As Integer = 1   ' 上記以外の画面から遷移

Public Const pintCstPreview       As Integer = 0   ' 帳票プレビュー
Public Const pintCstExcel         As Integer = 1   ' エクセル出力
Public Const pintCstMail          As Integer = 2   ' メール送信     ' ADD 2017/04/01 K.SATO

Private Const pstrCstMISETTEI     As String = "未設定"              ' 2005/12/14ADD T.SUZUKI
Private Const pintCstCANCEL       As Integer = 2501                 ' 2005/12/14 ADD T.SUZUKI

Private pobjDB                    As Database
Private pobjForm                  As Form

Private pobjKONT_DB               As Database

'2017/04/01 K.SATO ADD Start
Private Const pstrCstKAKUNIN    As String = "確認"

'先行予約メール構造体
Private Type Print530_MOSIKOMI_MAIL_INF
    strSendTo       As String
    strCONFIRM01    As String
    strCONFIRM02    As String
    strCONFIRM03    As String
    strCONFIRM04    As String
    strCONFIRM05    As String
    strCONFIRM06    As String
    strCONFIRM07    As String
    strCONFIRM08    As String
    strMOSIKOMIURL  As String
End Type
'2017/04/01 K.SATO ADD End

'==============================================================================*
'
'        MODULE_NAME      :ご紹介連絡表出力処理
'        MODULE_ID        :pfncRptPrintFVS530
'        Parameter        :intKbn    pintCstFVS530   = ご紹介連絡表印刷指示画面から遷移
'                                    pintCstOther    = 上記以外の画面から遷移
'                          intOutPut pintCstPreview  = 帳票プレビュー
'                                    pintCstExcel    = エクセル出力
'                          intPrintKbn   = 印刷種別
'                                          (1：新規紹介／2：紹介可能のみ／3：全て)
'                          objKONT_DB    = データベースオブジェクト(コンテナDB)
'                          strYOUKT_UKNO = 予約受付番号
'        CREATE_DATE      :2005/09/12
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function pfncRptPrintFVS530(ByVal intKbn As Integer, ByVal intOutPut As Integer, _
                                   ByVal intPrintkbn As Integer, _
                                   Optional ByRef objKONT_DB As Database = Null, _
                                   Optional strYOUKT_UKNO As String = "") As Boolean

    Dim blnDBFlg  As Boolean  ' 引数あり:True / 引数なし:False

    On Error GoTo pfncRptPrintFVS530_Err

    pfncRptPrintFVS530 = False

    doCmd.Hourglass True

    If intKbn = 0 Then
        ' ご紹介連絡表印刷指示画面からの遷移の場合
        ' フォームオブジェクトの設定
        Set pobjForm = Forms(pstrCstFRM_ID)

    ElseIf intKbn = 1 Then
        ' 上記以外の画面から遷移
        If Nz(strYOUKT_UKNO, "") = "" Then
            GoTo pfncRptPrintFVS530_Exit
        End If
    End If

    ' *** DBの設定 ***
    ' カレントDB設定
    Set pobjDB = CurrentDb

    ' パラメータのコンテナDBオブジェクトが背邸されていなかった場合、
    ' コンテナDBへの接続を行う
    If Not objKONT_DB Is Nothing Then
        Set pobjKONT_DB = objKONT_DB
        blnDBFlg = True
    Else
        ' コンテナDBに接続
        If fncConnectDB = False Then Exit Function
        blnDBFlg = False
    End If

    ' 帳票出力用データ作成
    If pfncCreateWorkFVS530(intKbn, intPrintkbn, strYOUKT_UKNO) = False Then GoTo pfncRptPrintFVS530_Exit

    Select Case intOutPut
        Case 0  ' *** 帳票プレビュー ***
            ' 予約ご紹介トラン更新処理
            If fncUpdINTR_TRAN() = False Then GoTo pfncRptPrintFVS530_Exit

            ' 帳票プレビュー
            doCmd.OpenReport pstrCstRPT_ID, acPreview, "", ""

        Case 1  ' *** エクセル出力 ***
            doCmd.OutputTo acQuery, "FVS530_Q01", "MicrosoftExcel(*.xls)", "", True, ""
    
        '2017/04/01 ADD Start K.SATO
        Case 2  ' *** メール送信 ***
            MsgBox ("mail")
            Call psubMail
        '2017/04/01 ADD End K.SATO
    End Select

    pobjDB.Close
    Set pobjDB = Nothing

    If blnDBFlg = False Then
        ' コンテナDB切断
        If fncDisConnectDB = False Then GoTo pfncRptPrintFVS530_Exit
    End If

    pfncRptPrintFVS530 = True

pfncRptPrintFVS530_Exit:
    ' カレントDB
    If Not pobjDB Is Nothing Then
        pobjDB.Close
        Set pobjDB = Nothing
    End If

    ' コンテナDB
    If blnDBFlg = False Then
        If Not pobjKONT_DB Is Nothing Then
            pobjKONT_DB.Close
            Set pobjKONT_DB = Nothing
        End If
    End If

    doCmd.Hourglass False
    Exit Function

pfncRptPrintFVS530_Err:
    Select Case Err.Number
        Case pintCstCANCEL
            ' Excel出力のキャンセル時のエラー
        Case Else
            MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    End Select
    Err.Clear
    GoTo pfncRptPrintFVS530_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルへ対象データのInsert
'        MODULE_ID        :pfncCreateWorkFVS530
'        Parameter        :intKbn      0 = ご紹介連絡表での帳票出力
'                                      1 = 他画面からご紹介連絡表出力
'                          strYOUKT_UKNO = 予約受付番号(省略可)
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function pfncCreateWorkFVS530(ByVal intKbn As Integer, _
                                     ByVal intPrintkbn As Integer, _
                                     Optional strYOUKT_UKNO As String = "") As Boolean

    Dim strSQL        As String
    Dim strWhere      As String
    Dim objRs         As Recordset
    Dim objWk         As Recordset

    Dim strBumoCD     As String     ' 部門コード
    Dim strBumoNm     As String     ' 部門名

    Dim strKaisya     As String     ' 会社名
    Dim strTEL_NO     As String     ' 会社電話番号
    Dim strFAX_NO     As String     ' 会社FAX番号

    Dim strUserName   As String     ' 顧客名(顧客名 - 代表者名)
    Dim strRenraku    As String     ' 電話番号 - FAX番号 - 携帯番号
    Dim strAddr       As String     ' 住所1 - 住所2 - 住所3
    Dim strSEV        As String     ' サービス内容1 - サービス内容2 - サービス内容3

    Dim varCode       As Variant

    Dim strYMD        As String
    Dim strYYYY       As String
    Dim strMM         As String
    Dim strDD         As String

    Dim strTANTN      As String     ' 受付担当者名

' 2005/11/29 ADD T.SUZUKI Start
    Dim dblPRICE      As Double     ' プライス
    Dim dblEZAPPI     As Double     ' 毎月雑費
    Dim intEZAPPI_CD  As Integer    ' 毎月雑費コード
    Dim dblPRICE_DIFF As Double     ' 料金差額
    Dim dblKEI_RENTKG As Double     ' 使用料

    Dim strREASON     As String     ' 差額理由
    Dim strEZAPPI     As String     ' 毎月雑費名
    Dim strBIKO       As String     ' 備考

    Dim objPRIC       As Recordset  ' 価格表テーブルレコードセット
' 2005/11/29 ADD T.SUZUKI End

    On Error GoTo pfncCreateWorkFVS530_Err

    pfncCreateWorkFVS530 = False

    doCmd.SetWarnings False

    ' ワークテーブルを削除
    If pfncDeleteWorkTableFVS530() = False Then GoTo pfncCreateWorkFVS530_Err

    ' ワークテーブルを作成
    If pfncCreateWorkTableFVS530() = False Then GoTo pfncCreateWorkFVS530_Err

    ' 部門マスタ・コントロールマスタデータ取得
    Call subGetControl_Data(strBumoCD, strBumoNm, strKaisya, strTEL_NO, strFAX_NO)

    ' 予約ご紹介トランよりデータ取得
    strSQL = " SELECT   YOUKT_UCODE, " & Chr(13)                        ' 予約受付トラン.顧客コード
    strSQL = strSQL & " YOUKT_NAME, " & Chr(13)                         ' 予約受付トラン.顧客名
    strSQL = strSQL & " YOUKT_TANM, " & Chr(13)                         ' 予約受付トラン.代表者名
    strSQL = strSQL & " YOUKT_TEL, " & Chr(13)                          ' 予約受付トラン.電話番号
    strSQL = strSQL & " YOUKT_KEITAI, " & Chr(13)                       ' 予約受付トラン.携帯番号
    strSQL = strSQL & " YOUKT_FAX, " & Chr(13)                          ' 予約受付トラン.FAX番号
    strSQL = strSQL & " YOUKT_UKTANTO, " & Chr(13)                      ' 予約受付トラン.受付担当者コード
    strSQL = strSQL & " YOUKT_TKDATE, " & Chr(13)                       ' 取置期限日                      ' 2005/11/29 ADD T.SUZUKI
    strSQL = strSQL & " INTRT_UKNO, " & Chr(13)                         ' 予約受付番号
    strSQL = strSQL & " INTRT_YCODE, " & Chr(13)                        ' ヤードコード
    strSQL = strSQL & " YARD_NAME, " & Chr(13)                          ' ヤード名
    strSQL = strSQL & " YARD_YUBINO, " & Chr(13)                        ' 郵便番号
    strSQL = strSQL & " YARD_ADDR_1, " & Chr(13)                        ' 住所1
    strSQL = strSQL & " YARD_ADDR_2, " & Chr(13)                        ' 住所2
    strSQL = strSQL & " YARD_ADDR_3, " & Chr(13)                        ' 住所3
    strSQL = strSQL & " YARD_SEV1N, " & Chr(13)                         ' サービス内容1
    strSQL = strSQL & " YARD_SEV2N, " & Chr(13)                         ' サービス内容2
    strSQL = strSQL & " YARD_SEV3N, " & Chr(13)                         ' サービス内容3
    strSQL = strSQL & " YARD_ENDEN, " & Chr(13)                         ' サービス期間
    strSQL = strSQL & " YARD_BEGIN_DAY, " & Chr(13)                     ' 営業開始日 2007/04/06 add tajima
    strSQL = strSQL & " INTRT_INTRONO, " & Chr(13)                      ' 紹介番号
    strSQL = strSQL & " INTRT_NO, " & Chr(13)                           ' コンテナ番号
    strSQL = strSQL & " CNTA_USAGE, " & Chr(13)                         ' 使用用途
    strSQL = strSQL & " CNTA_SIZE, " & Chr(13)                          ' サイズ
    strSQL = strSQL & " CNTA_STEP, " & Chr(13)                          ' 段区分                         ' 2005/11/29 ADD T.SUZUKI
    strSQL = strSQL & " CNTA_PRICE_DIFF, "                              ' 料金差額                       ' 2005/11/29 ADD T.SUZUKI
    strSQL = strSQL & " CNTA_REASON, "                                  ' 差額理由                       ' 2005/11/29 ADD T.SUZUKI
    strSQL = strSQL & " INTRT_INTROKBN, " & Chr(13)                     ' ご紹介区分
    strSQL = strSQL & " INTRT_NEARKBN  " & Chr(13)                      ' 近隣フラグ
    strSQL = strSQL & " FROM INTR_TRAN, " & Chr(13)                     ' ***予約ご紹介トラン
    strSQL = strSQL & "      YOUK_TRAN, " & Chr(13)                     ' ***予約受付トラン
    strSQL = strSQL & "      YARD_MAST, " & Chr(13)                     ' ***ヤードマスタ
    strSQL = strSQL & "      CNTA_MAST  " & Chr(13)                     ' ***コンテナマスタ
    strSQL = strSQL & " WHERE INTR_TRAN.INTRT_UKNO  = YOUK_TRAN.YOUKT_UKNO " & Chr(13)
    strSQL = strSQL & "   AND INTR_TRAN.INTRT_YCODE = YARD_MAST.YARD_CODE " & Chr(13)
    strSQL = strSQL & "   AND INTR_TRAN.INTRT_YCODE = CNTA_MAST.CNTA_CODE " & Chr(13)
    strSQL = strSQL & "   AND INTR_TRAN.INTRT_NO    = CNTA_MAST.CNTA_NO " & Chr(13)

    Select Case intKbn
        Case 0
            ' Where句生成
            strWhere = fncCreateWhere(intPrintkbn)

        Case 1
            ' 受付番号
            If Nz(strYOUKT_UKNO, "") <> "" Then
                strWhere = " AND YOUKT_UKNO = '" & strYOUKT_UKNO & "' "
            End If

            ' 印刷区分
            Select Case intPrintkbn
                Case pintCstViewNew  ' 新規紹介
                    strWhere = strWhere & " AND INTRT_INTROKBN = '1' "  ' ご紹介区分(取置きした)
                    strWhere = strWhere & " AND INTRT_FOUTD IS NULL "   ' 初回出力日

                Case pintCstViewOnly  ' 紹介可能のみ
                    strWhere = strWhere & " AND INTRT_INTROKBN = '1' "  ' ご紹介区分(取置きした)

                Case pintCstViewAll  ' 全て
                    ' 条件なし

                Case Else
            End Select
    End Select

    If strWhere <> "" Then
        strSQL = strSQL & strWhere
    End If

    ' ソート
    strSQL = strSQL & " ORDER BY YOUKT_UKTANTO, " & Chr(13)  ' 受付担当者コード
    strSQL = strSQL & "          INTRT_UKNO, " & Chr(13)     ' 予約受付番号
    strSQL = strSQL & "          INTRT_NEARKBN, " & Chr(13)  ' 近隣フラグ
    strSQL = strSQL & "          INTRT_YCODE, " & Chr(13)    ' ヤードコード
    strSQL = strSQL & "          INTRT_INTRONO DESC "        ' 紹介番号

    ' 予約ご紹介トラン読込(SQL-Serverに直接)
    Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

    If objRs.EOF = True Then
'    If objRs.RecordCount = 0 Then
        MsgBox "該当するデータがありません。", vbInformation, pstrCstERROR
        objRs.Close
        Set objRs = Nothing
        Exit Function
    End If

    ' ワークテーブルへ書き込み
    strSQL = "SELECT * FROM FVS530_W01 "
    Set objWk = pobjDB.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)

    Do Until objRs.EOF
        With objWk
            .AddNew

            .Fields("REPORT_ID") = pstrCstRPT_ID                          ' 帳票ID

            ' 顧客名
            If Nz(objRs.Fields("YOUKT_NAME"), "") <> "" Then
                strUserName = objRs.Fields("YOUKT_NAME")
            End If

            ' 代表者名
            If Nz(objRs.Fields("YOUKT_TANM"), "") <> "" Then
                strUserName = strUserName & Chr(13) & Chr(10) & objRs.Fields("YOUKT_TANM")
            End If
            .Fields("YOUKT_NAME") = strUserName                           ' 顧客名

            ' 顧客連絡先 --------------------------------------------------------------------------
            strRenraku = ""

            ' 電話番号
            If Nz(objRs.Fields("YOUKT_TEL"), "") <> "" Then
                strRenraku = pstrCstTEL & objRs.Fields("YOUKT_TEL")
            End If

            ' FAX番号
            If Nz(objRs.Fields("YOUKT_FAX"), "") <> "" Then
                If Nz(strRenraku) <> "" Then
                    strRenraku = strRenraku & Chr(13) & Chr(10) & pstrCstFAX & objRs.Fields("YOUKT_FAX")
                Else
                    strRenraku = pstrCstFAX & objRs.Fields("YOUKT_FAX")
                End If
            End If

            ' 携帯番号
            If Nz(objRs.Fields("YOUKT_KEITAI"), "") <> "" Then
                If Nz(strRenraku, "") <> "" Then
                    strRenraku = strRenraku & Chr(13) & Chr(10) & pstrCstKeitai & objRs.Fields("YOUKT_KEITAI")
                Else
                    strRenraku = pstrCstKeitai & objRs.Fields("YOUKT_KEITAI")
                End If
            End If
            .Fields("RENRAKUSAKI") = strRenraku
            ' ------------------------------------------------------------------------------------

            .Fields("INTRT_UKNO") = objRs.Fields("INTRT_UKNO")            ' 予約受付番号

            ' お客様番号(顧客コード)
            If Nz(objRs.Fields("YOUKT_UCODE"), "") <> "" Then
                .Fields("YOUKT_UCODE") = objRs.Fields("YOUKT_UCODE")
            End If

            .Fields("CONT_KAISYA") = strKaisya                            ' 会社名
            .Fields("BUMOM_BUMON") = strBumoNm                            ' 部門名
            .Fields("CONT_TEL_NO") = strTEL_NO                            ' 会社TEL
            .Fields("CONT_FAX_NO") = strFAX_NO                            ' 会社FAX

            .Fields("YOUKT_UKTANTO") = objRs.Fields("YOUKT_UKTANTO")      ' 受付担当者コード

            ' 受付担当名取得
            strTANTN = pfncGetTantoData(pobjDB, objRs.Fields("YOUKT_UKTANTO"))
            .Fields("TANTM_TANTN") = strTANTN                             ' 受付担当名

            .Fields("YOUKT_TKDATE") = objRs.Fields("YOUKT_TKDATE")        ' 取置期限日     ' 2005/11/29 ADD T.SUZUKI

            .Fields("INTRT_YCODE") = objRs.Fields("INTRT_YCODE")          ' ヤードコード
            .Fields("YARD_NAME") = objRs.Fields("YARD_NAME")              ' ヤード名

            .Fields("YARD_YUBINO") = objRs.Fields("YARD_YUBINO")          ' 郵便番号

            ' 住所 -------------------------------------------------------------------------------
            strAddr = ""

            ' 住所1
            If Nz(objRs.Fields("YARD_ADDR_1"), "") <> "" Then
                strAddr = objRs.Fields("YARD_ADDR_1")
            End If

            ' 住所2
            If Nz(objRs.Fields("YARD_ADDR_2"), "") <> "" Then
                If Nz(strAddr, "") <> "" Then
                    strAddr = strAddr & Chr(13) & Chr(10) & objRs.Fields("YARD_ADDR_2")
                Else
                    strAddr = objRs.Fields("YARD_ADDR_2")
                End If
            End If

            ' 住所3
            If Nz(objRs.Fields("YARD_ADDR_3"), "") <> "" Then
                If Nz(strAddr, "") <> "" Then
                    strAddr = strAddr & Chr(13) & Chr(10) & objRs.Fields("YARD_ADDR_3")
                Else
                    strAddr = objRs.Fields("YARD_ADDR_3")
                End If
            End If
            .Fields("YARD_ADDR") = strAddr
            ' ------------------------------------------------------------------------------------

            ' キャンペーン情報 --------------------------------------------------------------------
            strSEV = ""

            ' サービス内容1
            If Nz(objRs.Fields("YARD_SEV1N"), "") <> "" Then
                strSEV = objRs.Fields("YARD_SEV1N")
            End If

            ' サービス内容2
            If Nz(objRs.Fields("YARD_SEV2N"), "") <> "" Then
                If Nz(strSEV, "") <> "" Then
                    strSEV = strSEV & Chr(13) & Chr(10) & objRs.Fields("YARD_SEV2N")
                Else
                    strSEV = objRs.Fields("YARD_SEV2N")
                End If
            End If

            ' サービス内容3
            If Nz(objRs.Fields("YARD_SEV3N"), "") <> "" Then
                If Nz(strSEV, "") <> "" Then
                    strSEV = strSEV & Chr(13) & Chr(10) & objRs.Fields("YARD_SEV3N")
                Else
                    strSEV = objRs.Fields("YARD_SEV3N")
                End If
            End If

            ' サービス期間
            If Nz(objRs.Fields("YARD_ENDEN"), "") <> "" Then
                If Nz(strSEV, "") <> "" Then
                    strSEV = strSEV & Chr(13) & Chr(10) & objRs.Fields("YARD_ENDEN")
                Else
                    strSEV = objRs.Fields("YARD_ENDEN")
                End If
            End If
            .Fields("YARD_SEV") = strSEV
            ' ------------------------------------------------------------------------------------
            .Fields("YARD_BEGIN_DAY") = objRs.Fields("YARD_BEGIN_DAY")    ' 営業開始日 2007/04/06 add tajima

            .Fields("INTRT_INTRONO") = objRs.Fields("INTRT_INTRONO")      ' ご紹介番号
            .Fields("INTRT_NO") = objRs.Fields("INTRT_NO")                ' コンテナ番号

            ' 使用用途
            If Nz(objRs.Fields("CNTA_USAGE"), "") <> "" Then
                varCode = fncGetName_MAST(0, "086", objRs.Fields("CNTA_USAGE"))
'                varCode = fncGETNAME("086", objRs.Fields("CNTA_USAGE"))
                .Fields("USAGE_NAME") = varCode                           ' 使用用途
            End If

            .Fields("CNTA_SIZE") = objRs.Fields("CNTA_SIZE")              ' サイズ

' 2005/11/29 ADD T.SUZUKI Start
            ' 初期化
            dblPRICE = 0       ' プライス
            dblEZAPPI = 0      ' 毎月雑費
            intEZAPPI_CD = 0   ' 毎月雑費コード
            dblPRICE_DIFF = 0  ' 料金差額
            dblKEI_RENTKG = 0  ' 使用料

            ' 価格表テーブルよりデータ取得
            strSQL = " SELECT   PRIC_PRICE, " & Chr(13)        ' プライス
            strSQL = strSQL & " PRIC_EZAPPI, " & Chr(13)       ' 毎月雑費
            strSQL = strSQL & " PRIC_EZAPPI_CODE  " & Chr(13)  ' 毎月雑費コード
            strSQL = strSQL & "  FROM dbo_PRIC_TABL " & Chr(13)
            strSQL = strSQL & " WHERE PRIC_STEP = " & objRs.Fields("CNTA_STEP") & Chr(13)
            strSQL = strSQL & "   AND PRIC_SIZE = " & objRs.Fields("CNTA_SIZE") & Chr(13)
            strSQL = strSQL & "   AND PRIC_USAGE = " & objRs.Fields("CNTA_USAGE") & Chr(13)
            strSQL = strSQL & "   AND PRIC_YCODE = " & objRs.Fields("INTRT_YCODE") & Chr(13)
            Set objPRIC = pobjDB.OpenRecordset(strSQL, dbReadOnly)

            If objPRIC.EOF = False Then
                dblPRICE = Nz(objPRIC.Fields("PRIC_PRICE"), 0)            ' プライス
                dblEZAPPI = Nz(objPRIC.Fields("PRIC_EZAPPI"), 0)          ' 毎月雑費
                intEZAPPI_CD = Nz(objPRIC.Fields("PRIC_EZAPPI_CODE"), 0)  ' 毎月雑費コード

' 2005/12/14 ADD T.SUZUKI Start
                dblPRICE_DIFF = Nz(objRs.Fields("CNTA_PRICE_DIFF"), 0)    ' コンテナマスタ.料金差額

                ' *** 使用料算出 ***
                ' → 使用料 = プライス ＋ 毎月雑費 ＋ 料金差額
                dblKEI_RENTKG = dblPRICE + dblEZAPPI + dblPRICE_DIFF
                .Fields("CNTA_KEI_RENTKG") = dblKEI_RENTKG                ' 使用料
                .Fields("CNTA_MISETTEI") = ""                             ' 使用料(未設定)

            Else
                ' *** 使用料算出 ***
                ' 価格表テーブルのデータが存在しない場合「未設定」をセット
                .Fields("CNTA_KEI_RENTKG") = Null                         ' 使用料
                .Fields("CNTA_MISETTEI") = pstrCstMISETTEI                ' 使用料(未設定)
' 2005/12/14 ADD T.SUZUKI End
            End If
            objPRIC.Close
            Set objPRIC = Nothing

' 2005/12/14 DEL T.SUZUKI Start
'            lngPRICE_DIFF = Nz(objRs.Fields("CNTA_PRICE_DIFF"), 0)        ' コンテナマスタ.料金差額

'            ' *** 使用料算出 ***
'            ' → 使用料 = プライス ＋ 毎月雑費 ＋ 料金差額
'            lngKEI_RENTKG = lngPRICE + lngEZAPPI + lngPRICE_DIFF
'            .Fields("CNTA_KEI_RENTKG") = lngKEI_RENTKG                    ' 使用料
' 2005/12/14 DEL T.SUZUKI End
' 2005/11/29 ADD T.SUZUKI End
'            .Fields("CNTA_KEI_RENTKG") = objRs.Fields("CNTA_KEI_RENTKG")  ' 使用料  ' 2005/11/29 DEL T.SUZUKI

' 2005/11/29 DEL T.SUZUKI Start
'            ' 作成日
'            strYMD = ""
'            If Nz(objRs.Fields("INTRT_INSED"), "") <> "" Then
'                strYYYY = Mid(objRs.Fields("INTRT_INSED"), 1, 4)
'                strMM = Mid(objRs.Fields("INTRT_INSED"), 5, 2)
'                strDD = Mid(objRs.Fields("INTRT_INSED"), 7, 2)
'                strYMD = strYYYY & "/" & strMM & "/" & strDD
'
'                If IsDate(strYMD) = True Then
'                    .Fields("INTRT_INSED") = strYMD
'                Else
'                    .Fields("INTRT_INSED") = Null
'                End If
'            Else
'                .Fields("INTRT_INSED") = Null
'            End If
' 2005/11/29 DEL T.SUZUKI End

            .Fields("INTRT_INTROKBN") = objRs.Fields("INTRT_INTROKBN")    ' ご紹介区分

            ' ご紹介区分取得
            If Nz(objRs.Fields("INTRT_INTROKBN"), "") <> "" Then
                varCode = fncGetName_MAST(0, "097", objRs.Fields("INTRT_INTROKBN"))
'                varCode = fncGETNAME("097", objRs.Fields("INTRT_INTROKBN"))
                .Fields("ZYOKYO") = varCode                               ' 状況
            End If

            .Fields("INTRT_NEARKBN") = objRs.Fields("INTRT_NEARKBN")      ' 近隣フラグ

' 2005/11/29 ADD T.SUZUKI Start
            strREASON = Nz(Trim(objRs.Fields("CNTA_REASON")), "")         ' 差額理由

            strBIKO = ""
            Select Case intEZAPPI_CD
                Case Is <> 0 ' *** 毎月雑費あり ***
                    ' 毎月雑費名称取得
                    strEZAPPI = fncGetName_MAST(0, "096", intEZAPPI_CD)   ' 毎月雑費名
                    dblEZAPPI = dblEZAPPI                                 ' 毎月雑費

                    ' 備考文字列作成
                    '↓MOD 2020/07/27 S.WATANABE
                    'strBIKO = strEZAPPI & pcstrBIKO_Message1 & Format(dblEZAPPI, "#,##0") & pcstrBIKO_Message2
                    If dblEZAPPI <> 0 Then
                        strBIKO = strEZAPPI & pcstrBIKO_Message1 & Format(dblEZAPPI, "#,##0") & pcstrBIKO_Message2
                    End If
                    '↑MOD 2020/07/27 S.WATANABE

                    If strREASON <> "" Then
                        ' 差額あり
                        strBIKO = strBIKO & vbCrLf & strREASON
                    Else
                        ' 差額なし
                    End If

                Case 0  ' *** 毎月雑費なし ***
                    ' 備考文字列作成
                    strBIKO = strREASON
            End Select

            ' 備考
            .Fields("BIKO") = strBIKO

            ' ソート用フィールド
            ' (サービス内容1・サービス内容2・サービス内容3・サービス期間のいずれかに
            '  値が入っている場合、キャンペーンヤードとする)
            If Nz(Trim(objRs.Fields("YARD_SEV1N")), "") <> "" Or _
               Nz(Trim(objRs.Fields("YARD_SEV2N")), "") <> "" Or _
               Nz(Trim(objRs.Fields("YARD_SEV3N")), "") <> "" Or _
               Nz(Trim(objRs.Fields("YARD_ENDEN")), "") <> "" Then
                .Fields("CAMPAIN") = 0  ' キャンペーンヤード
            Else
                .Fields("CAMPAIN") = 1  ' キャンペーンヤードではない
            End If
' 2005/11/29 ADD T.SUZUKI End

            .UPDATE
        End With

        objRs.MoveNext
    Loop

    pfncCreateWorkFVS530 = True

pfncCreateWorkFVS530_Exit:
    If Not objRs Is Nothing Then: objRs.Close: Set objRs = Nothing
    If Not objWk Is Nothing Then: objWk.Close: Set objWk = Nothing
    If Not objPRIC Is Nothing Then: objPRIC.Close: Set objPRIC = Nothing

    Exit Function

pfncCreateWorkFVS530_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo pfncCreateWorkFVS530_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルのデータ削除
'        MODULE_ID        :pfncDeleteWorkFVS530
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function pfncDeleteWorkTableFVS530() As Boolean
'Private Function pfncDeleteWorkTableFVS530(ByVal pobjDB As Database) As Boolean

    Dim strSQL As String

    On Error GoTo pfncDeleteWorkFVS530_Err

    pfncDeleteWorkTableFVS530 = False

    ' テーブルの存在チェック
    If TableExist(pstrCstWK_TABLE_NM) = True Then

        ' テーブル削除
        pobjDB.TableDefs.Delete pstrCstWK_TABLE_NM
        Select Case Err.Number
            Case 0
            Case 3265
                Err.Clear
            Case Else
                MsgBox "subDropWorkTable:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
                Err.Clear
        End Select
    End If

    pfncDeleteWorkTableFVS530 = True

pfncDeleteWorkFVS530_Exit:
    Exit Function

pfncDeleteWorkFVS530_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo pfncDeleteWorkFVS530_Exit
End Function

'==============================================================================*
'
' MODULE_NAME :TableExist
' 機能 :ACCESSテーブル存在チェック
' IN :テーブル名
' OUT :True=存在する False=存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function TableExist(strTableName As String) As Boolean

    Dim tdf As TableDef

    TableExist = False

    For Each tdf In pobjDB.TableDefs
        If tdf.NAME = strTableName Then
            TableExist = True
            Exit For
        End If
    Next tdf
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルの作成
'        MODULE_ID        :pfncCreateWorkFVS530
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function pfncCreateWorkTableFVS530() As Boolean

    Dim tdfNew    As TableDef
    Dim intCount  As Integer

    On Error GoTo pfncCreateWorkTableFVS530_Err

    pfncCreateWorkTableFVS530 = False

    Set tdfNew = pobjDB.CreateTableDef(pstrCstWK_TABLE_NM)
    With tdfNew
        .Fields.Append .CreateField("REPORT_ID", dbText, 50)       ' レポートID
        .Fields.Append .CreateField("YOUKT_NAME", dbText, 54)      ' 顧客名 + 代表者名
        .Fields.Append .CreateField("RENRAKUSAKI", dbText, 100)    ' 電話番号 + FAX番号 + 携帯番号
        .Fields.Append .CreateField("INTRT_UKNO", dbText, 11)      ' 予約受付番号
        .Fields.Append .CreateField("YOUKT_UCODE", dbLong)         ' 顧客コード
        .Fields.Append .CreateField("CONT_KAISYA", dbText, 30)     ' 会社名
        .Fields.Append .CreateField("BUMOM_BUMON", dbText, 50)     ' 部門名称
        .Fields.Append .CreateField("CONT_TEL_NO", dbText, 15)     ' 会社TEL
        .Fields.Append .CreateField("CONT_FAX_NO", dbText, 15)     ' 会社FAX
        .Fields.Append .CreateField("YOUKT_UKTANTO", dbText, 3)    ' 受付担当者コード
        .Fields.Append .CreateField("TANTM_TANTN", dbText, 36)     ' 受付担当者名
        .Fields.Append .CreateField("YOUKT_TKDATE", dbDate)        ' 取置期限日                                                             ' 2005/11/30 ADD T.SUZUKI
        .Fields.Append .CreateField("INTRT_YCODE", dbLong)         ' ヤードコード
        .Fields.Append .CreateField("YARD_NAME", dbText, 36)       ' ヤード名
        .Fields.Append .CreateField("INTRT_NO", dbLong)            ' コンテナ番号
        .Fields.Append .CreateField("YARD_YUBINO", dbText, 10)     ' 郵便番号
        .Fields.Append .CreateField("YARD_ADDR", dbText, 108)      ' ヤード住所(住所1 + 住所2 + 住所3)
        .Fields.Append .CreateField("YARD_SEV", dbText, 240)       ' キャンペーン(サービス内容1 + サービス内容2 + サービス内容3 + サービス期間)
        .Fields.Append .CreateField("YARD_BEGIN_DAY", dbDate)      ' 営業開始日 2007/04/06 add tajima
        .Fields.Append .CreateField("INTRT_INTRONO", dbLong)       ' 紹介番号
        .Fields.Append .CreateField("CNTA_USAGE", dbText, 50)      ' 使用用途
        .Fields.Append .CreateField("USAGE_NAME", dbText, 50)      ' 使用用途名
        .Fields.Append .CreateField("CNTA_SIZE", dbSingle)         ' サイズ（帖）
        .Fields.Append .CreateField("CNTA_KEI_RENTKG", dbDouble)   ' 使用料
        .Fields.Append .CreateField("CNTA_MISETTEI", dbText, 10)   ' 使用料(未設定)  ' 2005/12/14 ADD T.SUZUKI
'        .Fields.Append .CreateField("INTRT_INSED", dbText, 10)     ' 作成日                                                                ' 2005/11/29 DEL T.SUZUKI
        .Fields.Append .CreateField("INTRT_INTROKBN", dbText, 1)   ' ご紹介区分
        .Fields.Append .CreateField("ZYOKYO", dbText, 50)          ' 状況
        .Fields.Append .CreateField("INTRT_NEARKBN", dbText, 1)    ' 近隣フラグ
        .Fields.Append .CreateField("BIKO", dbText, 150)           ' 備考                                                                   ' 2005/11/29 ADD T.SUZUKI
        .Fields.Append .CreateField("CAMPAIN", dbInteger)          ' コンテナ番号  ' hana

        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount
    End With

    ' テーブルの追加
    pobjDB.TableDefs.Append tdfNew

    pfncCreateWorkTableFVS530 = True

pfncCreateWorkTableFVS530_Exit:
    Exit Function

pfncCreateWorkTableFVS530_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo pfncCreateWorkTableFVS530_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :部門コード/部門名称の取得
'        MODULE_ID        :subGetControl_Data
'        IN               :objDB      = DB接続
'                          strBumonCd = 部門ｺｰﾄﾞ
'                          strBumonNm = 部門名
'                          strKAISYA  = 会社名
'                          strTEL_NO  = TEL
'                          strFAX_NO  = FAX
'        CREATE_DATE      :2005/09/12
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subGetControl_Data(ByRef strBumonCd As String, ByRef strBumonNm As String, _
                               ByRef strKaisya As String, ByRef strTEL_NO As String, ByRef strFAX_NO As String)

    Dim strSQL  As String
    Dim objRs   As Recordset

    On Error GoTo subGetControl_Data_Err

    strSQL = "SELECT " & Chr(13)
    strSQL = strSQL & "BUMO_MAST.BUMOM_BUMOC, " & Chr(13)      ' 部門コード
    strSQL = strSQL & "BUMO_MAST.BUMOM_BUMON, " & Chr(13)      ' 部門名称
    strSQL = strSQL & "dbo_CONT_MAST.CONT_KAISYA, " & Chr(13)  ' 会社名
    strSQL = strSQL & "dbo_CONT_MAST.CONT_TEL_NO, " & Chr(13)  ' TEL
    strSQL = strSQL & "dbo_CONT_MAST.CONT_FAX_NO  " & Chr(13)  ' FAX
    strSQL = strSQL & "FROM " & Chr(13)
    strSQL = strSQL & "dbo_CONT_MAST INNER JOIN BUMO_MAST ON " & Chr(13)
    strSQL = strSQL & "dbo_CONT_MAST.CONT_BUMOC = BUMO_MAST.BUMOM_BUMOC; "
    Set objRs = pobjDB.OpenRecordset(strSQL, dbReadOnly)

    With objRs
        If Not .EOF Then
            strBumonCd = .Fields("BUMOM_BUMOC")
            strBumonNm = .Fields("BUMOM_BUMON")
            strKaisya = .Fields("CONT_KAISYA")
            strTEL_NO = .Fields("CONT_TEL_NO")
            strFAX_NO = .Fields("CONT_FAX_NO")

        Else
            strBumonCd = Null
            strBumonNm = Null
            strKaisya = Null
            strTEL_NO = Null
            strFAX_NO = Null
        End If
    End With

subGetControl_Data_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    Exit Sub

subGetControl_Data_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo subGetControl_Data_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :pfncGetTantoData
'        機能             :担当者マスタ情報取得
'        IN               :vTanto = 担当者ｺｰﾄﾞ / vTanNm = 担当者名
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function pfncGetTantoData(ByRef objDb As Database, ByRef vTanto As Variant) As String

    Dim strSQL  As String
    Dim objRs   As Recordset

    On Error GoTo pfncGetTantoData_Err

    pfncGetTantoData = ""

    strSQL = " SELECT   TANT_MAST.TANTM_TANTC, " & Chr(13)
    strSQL = strSQL & " TANT_MAST.TANTM_TANTN " & Chr(13)
    strSQL = strSQL & " FROM TANT_MAST INNER JOIN dbo_CONT_MAST ON " & Chr(13)
    strSQL = strSQL & "      TANT_MAST.TANTM_BUMOC = dbo_CONT_MAST.CONT_BUMOC " & Chr(13)
    strSQL = strSQL & " WHERE TANT_MAST.TANTM_TANTC = " & "'" & Format$(vTanto, "000") & "'"
    Set objRs = objDb.OpenRecordset(strSQL, dbOpenDynaset)

    If objRs.EOF Then
        objRs.Close
        Set objRs = Nothing
        Exit Function
    Else
        If Nz(objRs("TANTM_TANTN"), "") <> "" Then
            pfncGetTantoData = objRs("TANTM_TANTN")
        End If
    End If

pfncGetTantoData_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If

    Exit Function

pfncGetTantoData_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo pfncGetTantoData_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :Where句生成
'        MODULE_ID        :fncCreateWhere
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'        UPDATE_DATE      :2017/04/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWhere(Optional intPrintkbn As Integer = 0) As String

    Dim strWhere    As String
    Dim strINSED_F  As String
    Dim strINSED_T  As String
    Dim strMLDATE_F  As String                                                   'INSERT 2017/04/01 K.SATO
    Dim strMLDATE_T  As String                                                   'INSERT 2017/04/01 K.SATO
    Dim intRenraku  As Integer

    With pobjForm
        ' *** 紹介データ作成日の大小チェック ***
        Call ChkDateFromTo(.txt_INTRT_INSED_F, .txt_INTRT_INSED_T)

        ' *** 受付担当コードの大小チェック ***
        Call ChkUKECDFromTo(0, .cmb_Tantoc_F, .txt_Tantoc_F, .cmb_Tantoc_T, .txt_Tantoc_T)

        ' ** 予約受付状態区分 ***
        Call ChkUKECDFromTo(1, .cmb_Yukbn_F, .txt_Yukbn_F, .cmb_Yukbn_T, .txt_Yukbn_T)

        ' *** メール送信日の大小チェック ***
        Call ChkDateFromTo(.txt_MLDATE_F, .txt_MLDATE_T)                               'INSERT 2017/04/01 K.SATO

        strINSED_F = Replace(Nz(.txt_INTRT_INSED_F, ""), "/", "")
        strINSED_T = Replace(Nz(.txt_INTRT_INSED_T, ""), "/", "")

        strMLDATE_F = Replace(Nz(.txt_MLDATE_F, ""), "/", "")                           'INSERT 2017/04/01 K.SATO
        strMLDATE_T = Replace(Nz(.txt_MLDATE_T, ""), "/", "")                           'INSERT 2017/04/01 K.SATO

        ' 紹介データ作成日
        If (Nz(.txt_INTRT_INSED_F, "") <> "" And Nz(.txt_INTRT_INSED_T, "") <> "") Then
            ' 紹介データ作成日(From)：入力
            ' 紹介データ作成日(To)  ：入力
            strWhere = " AND INTRT_INSED BETWEEN '" & strINSED_F & "' AND '" & strINSED_T & "' "

        ElseIf (Nz(.txt_INTRT_INSED_F, "") <> "" And Nz(.txt_INTRT_INSED_T, "") = "") Then
            ' 紹介データ作成日(From)：入力
            ' 紹介データ作成日(To)  ：未入力
            strWhere = " AND INTRT_INSED >= '" & strINSED_F & "' "

        ElseIf (Nz(.txt_INTRT_INSED_F, "") = "" And Nz(.txt_INTRT_INSED_T, "") <> "") Then
            ' 紹介データ作成日(From)：未入力
            ' 紹介データ作成日(To)  ：入力
            strWhere = " AND INTRT_INSED <= '" & strINSED_T & "' "
        End If

        'INSERT 2017/04/01 K.SATO START
        ' メール送信日
        If (Nz(.txt_MLDATE_F, "") <> "" And Nz(.txt_MLDATE_T, "") <> "") Then
            ' 紹介データ作成日(From)：入力
            ' 紹介データ作成日(To)  ：入力
            strWhere = " AND convert(NVARCHAR,youkt_mldate,112) BETWEEN '" & strMLDATE_F & "' AND '" & strMLDATE_T & "' "

        ElseIf (Nz(.txt_MLDATE_F, "") <> "" And Nz(.txt_MLDATE_T, "") = "") Then
            ' 紹介データ作成日(From)：入力
            ' 紹介データ作成日(To)  ：未入力
            strWhere = " AND convert(NVARCHAR,youkt_mldate,112) >= '" & strMLDATE_F & "' "

        ElseIf (Nz(.txt_MLDATE_F, "") = "" And Nz(.txt_MLDATE_T, "") <> "") Then
            ' 紹介データ作成日(From)：未入力
            ' 紹介データ作成日(To)  ：入力
            strWhere = " AND convert(NVARCHAR,youkt_mldate,112) <= '" & strMLDATE_T & "' "
        End If
        'INSERT 2017/04/01 K.SATO END

        ' 受付担当コード
        If (Nz(.cmb_Tantoc_F, "") <> "" And Nz(.cmb_Tantoc_T, "") <> "") Then
            ' 受付担当コード(From)：入力
            ' 受付担当コード(To)  ：入力
            strWhere = strWhere & " AND YOUKT_UKTANTO BETWEEN '" & .cmb_Tantoc_F & "' AND '" & .cmb_Tantoc_T & "' "

        ElseIf (Nz(.cmb_Tantoc_F, "") <> "" And Nz(.cmb_Tantoc_T, "") = "") Then
            ' 受付担当コード(From)：入力
            ' 受付担当コード(To)  ：未入力
            strWhere = strWhere & " AND YOUKT_UKTANTO >= '" & .cmb_Tantoc_F & "' "

        ElseIf (Nz(.cmb_Tantoc_F, "") = "" And Nz(.cmb_Tantoc_T, "") <> "") Then
            ' 受付担当コード(From)：未入力
            ' 受付担当コード(To)  ：入力
            strWhere = strWhere & " AND YOUKT_UKTANTO <= '" & .cmb_Tantoc_T & "' "
        End If

        ' 印刷種別
        Select Case intPrintkbn
            Case pintCstViewNew  ' 新規紹介
                strWhere = strWhere & " AND INTRT_INTROKBN = '1' "  ' ご紹介区分(取置きした)
                strWhere = strWhere & " AND INTRT_FOUTD IS NULL "   ' 初回出力日

            Case pintCstViewOnly  ' 紹介可能のみ
                strWhere = strWhere & " AND INTRT_INTROKBN = '1' "  ' ご紹介区分(取置きした)

            Case pintCstViewAll  ' 全て
                ' 条件なし

            Case Else
        End Select

        ' 受付番号
        If Nz(.txt_Ukno, "") <> "" Then
            strWhere = strWhere & " AND YOUKT_UKNO = '" & .txt_Ukno & "' "
        End If

        ' 連絡区分
        If .cmb_Renraku <> pintCstRenraku Then
            ' -999：全て
            '    0：電話
            '    1：携帯電話
            '    2：FAX
            '   10：Eメール
            intRenraku = .cmb_Renraku

            strWhere = strWhere & " AND YOUKT_RENKBN = " & intRenraku  ' 連絡区分
        End If

        ' 予約受付状態区分
        If (Nz(.cmb_Yukbn_F, "") <> "" And Nz(.cmb_Yukbn_T, "") <> "") Then
            ' 受付担当コード(From)：入力
            ' 受付担当コード(To)  ：入力
            strWhere = strWhere & " AND YOUKT_YUKBN BETWEEN " & Val(.cmb_Yukbn_F) & " AND " & Val(.cmb_Yukbn_T)

        ElseIf (Nz(.cmb_Yukbn_F, "") <> "" And Nz(.cmb_Yukbn_T, "") = "") Then
            ' 受付担当コード(From)：入力
            ' 受付担当コード(To)  ：未入力
            strWhere = strWhere & " AND YOUKT_YUKBN >= " & Val(.cmb_Yukbn_F)

        ElseIf (Nz(.cmb_Yukbn_F, "") = "" And Nz(.cmb_Yukbn_T, "") <> "") Then
            ' 受付担当コード(From)：未入力
            ' 受付担当コード(To)  ：入力
            strWhere = strWhere & " AND YOUKT_YUKBN <= " & Val(.cmb_Yukbn_T)
        End If
    End With

    fncCreateWhere = strWhere
End Function

'==============================================================================*
'
'        MODULE_NAME      :日付入力チェック（From/Toを逆に入力した時の対応）
'        MODULE_ID        :ChkDateFromTo
'        IN               :
'        CREATE_DATE      :2005/09/12
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ChkDateFromTo(ByRef ctlTextF As TextBox, ByRef ctlTextT As TextBox)

    Dim strDayF As String
    Dim strDayT As String
    Dim strSwap As String

    If ctlTextF <> "" And ctlTextT <> "" Then
        strDayF = Format(ctlTextF, "yyyy/mm/dd")  ' 紹介データ作成日(From)
        strDayT = Format(ctlTextT, "yyyy/mm/dd")  ' 紹介データ作成日(To)

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
'        MODULE_NAME      :受付担当コード入力チェック（From/Toを逆に入力した時の対応）
'        MODULE_ID        :ChkYARDFromTo
'        Parameter        :intKbn ：0 = 受付担当者コード
'                                   1 = 予約受付状態区分
'                          ctlcmbF：該当項目(From)
'                          ctlcmbT：該当項目(To)
'        CREATE_DATE      :2005/09/12
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ChkUKECDFromTo(ByVal intKbn As Integer, _
                           ByRef ctlcmbF As ComboBox, ByRef ctlTextF As TextBox, _
                           ByRef ctlcmbT As ComboBox, ByRef ctlTextT As TextBox)

    Dim strUKECD_F  As String  ' 受付担当コード(FROM)
    Dim strUKECD_T  As String  ' 受付担当コード(TO)
    Dim strSwap     As String  ' ワーク変数

    strSwap = ""
    If Trim(ctlcmbF) <> "" And Trim(ctlcmbT) <> "" Then
        Select Case intKbn
            Case 0
                strUKECD_F = Format(ctlcmbF, "000")  ' 受付担当コード(From)
                strUKECD_T = Format(ctlcmbT, "000")  ' 受付担当コード(To)
            Case 1
                strUKECD_F = Format(ctlcmbF, "00")   ' 予約受付状態区分(From)
                strUKECD_T = Format(ctlcmbT, "00")   ' 予約受付状態区分(To)
        End Select

        If strUKECD_F > strUKECD_T Then
            ' 大小入力エラー時
            ' *** 受付担当コード
            strSwap = strUKECD_F
            strUKECD_F = strUKECD_T
            strUKECD_T = strSwap

            ' 参照元へ再セット
            ctlcmbF = strUKECD_F
            ctlcmbT = strUKECD_T

            ' 名称
            ' *** 受付担当名
            strSwap = ""
            strUKECD_F = ""
            strUKECD_T = ""

            strSwap = ctlTextF
            strUKECD_F = ctlTextT
            strUKECD_T = strSwap

            ' 参照元へ再セット
            ctlTextF = strUKECD_F
            ctlTextT = strUKECD_T
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :名称取得
'        MODULE_ID        :fncGetName_MAST
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncGetName_MAST(ByVal intKbn As Integer, ByVal strNAME_ID As String, ByVal varNAME_CODE As Variant) As String

    Dim strWhere  As String
    Dim strName   As String

    On Error GoTo fncGetName_MAST_Err

    fncGetName_MAST = ""

    If Nz(varNAME_CODE, "") <> "" Then                                          'INSERT 2017/12/20 N.IMAI
        ' 名称マスタより略称を取得
        strWhere = " NAME_ID = '" & strNAME_ID & "'"
        strWhere = strWhere & " AND NAME_CODE = " & varNAME_CODE
    
        strName = ""
        Select Case intKbn
            Case 0  ' *** 名称 ***
                strName = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")
            Case 1  ' *** 略称 ***
                strName = Nz(DLookup("NAME_RYAK", "dbo_NAME_MAST", strWhere), "")
        End Select
    
        If strName <> "" Then
            fncGetName_MAST = strName
        End If
    End If                                                                      'INSERT 2017/12/20 N.IMAI

fncGetName_MAST_Exit:
    Exit Function

fncGetName_MAST_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo fncGetName_MAST_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :予約ご紹介トランの更新処理
'        MODULE_ID        :fncUpdINTR_TRAN
'        IN               :
'        CREATE_DATE      :2005/09/12
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdINTR_TRAN() As Boolean

    Dim strSQL As String
    Dim objRs  As Recordset
    Dim objWk  As Recordset

    fncUpdINTR_TRAN = False

    strSQL = "  SELECT  INTRT_UKNO, " & Chr(13)    ' 予約受付番号
    strSQL = strSQL & " INTRT_INTRONO " & Chr(13)  ' 紹介番号
    strSQL = strSQL & " FROM FVS530_W01 " & Chr(13)
    strSQL = strSQL & " ORDER BY INTRT_UKNO "

    ' ワークテーブル読込
    Set objRs = pobjDB.OpenRecordset(strSQL, dbReadOnly)

    Do Until objRs.EOF
        strSQL = "  SELECT  INTRT_FOUTD, " & Chr(13)  ' 初回出力日
        strSQL = strSQL & " INTRT_LOUTD, " & Chr(13)  ' 最終出力日
        strSQL = strSQL & " INTRT_UPDAD, " & Chr(13)  ' 更新日付
        strSQL = strSQL & " INTRT_UPDAJ, " & Chr(13)  ' 更新時刻
        strSQL = strSQL & " INTRT_UPDPB, " & Chr(13)  ' 更新プログラムID
        strSQL = strSQL & " INTRT_UPDUB  " & Chr(13)  ' 更新ユーザーID
        strSQL = strSQL & " FROM INTR_TRAN " & Chr(13)
        strSQL = strSQL & " WHERE INTRT_UKNO    = '" & objRs.Fields("INTRT_UKNO") & "' " & Chr(13)  ' 予約受付番号
        strSQL = strSQL & "   AND INTRT_INTRONO = " & objRs.Fields("INTRT_INTRONO")                 ' 紹介番号

        ' 予約ご紹介トラン読込
        Set objWk = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

        If objWk.EOF = False Then
            With objWk
                .Edit

                ' 初回出力日
                If Nz(.Fields("INTRT_FOUTD"), "") = "" Then
                    .Fields("INTRT_FOUTD") = Format(DATE, "YYYY/MM/DD")
                End If

                .Fields("INTRT_LOUTD") = Format(DATE, "YYYY/MM/DD")  ' 最終出力日
                .Fields("INTRT_UPDAD") = Format$(DATE, "YYYYMMDD")   ' 更新日付
                .Fields("INTRT_UPDAJ") = Format$(time, "hhmmss")     ' 更新時刻
                .Fields("INTRT_UPDPB") = pstrCstFRM_ID               ' 更新プログラムID
                .Fields("INTRT_UPDUB") = MSZZ000.LsGetUserName       ' 更新ユーザーID

                .UPDATE

                objWk.Close
                Set objWk = Nothing
            End With
        End If

        objRs.MoveNext
    Loop

    fncUpdINTR_TRAN = True

fncUpdINTR_TRAN_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    If Not objWk Is Nothing Then
        objWk.Close
        Set objWk = Nothing
    End If
    Exit Function

fncUpdINTR_TRAN_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

    GoTo fncUpdINTR_TRAN_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :DB接続
'        MODULE_ID        :fncConnectDB
'        CREATE_DATE      :2005/10/19
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncConnectDB() As Boolean

    Dim strBumoCD     As String

    Dim strConnect    As String
    Dim strDataSource As String

    On Error GoTo fncConnectDB_Err

    fncConnectDB = False

    ' 部門コード取得
    strBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")

    ' -----------------------------------------------------------------------------------------------------------
    ' コンテナDB接続
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumoCD & "'"), "")
    
    If strDataSource = "" Then
        MsgBox "SETU_TABLの設定が不正です。", vbExclamation, PROG_ID
        Set pobjKONT_DB = Nothing
    Else
        Set pobjKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(strBumoCD))
    End If
    ' -----------------------------------------------------------------------------------------------------------

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
Private Function fncDisConnectDB() As Boolean

    On Error GoTo fncDisConnectDB_Err

    fncDisConnectDB = False

    ' コンテナDB
    pobjKONT_DB.Close
    Set pobjKONT_DB = Nothing

    fncDisConnectDB = True

    Exit Function

fncDisConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :メール送信ボタン押下時の処理
'        MODULE_ID        :cmd_Mail_Click
'        CREATE_DATE      :2017/04/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub psubMail()

    On Error GoTo psubMail_Err

    Dim strMsg        As String
    Dim lngRc         As Long

    Dim ecnt                As Integer
    Dim cnt                 As Integer
    Dim strBumonArr()       As String
    Dim varFilePath         As Variant
    Dim varFileName         As Variant
    Dim strUrl              As String
    Dim strMailSubject      As String
    Dim strTempMailMessage  As String
    Dim strSendMailMessage  As String
    Dim strSQL              As String
    Dim objWk               As Recordset
    Dim objWk2              As Recordset
    Dim objYo               As Recordset
    Dim objPr               As Recordset

    Dim stMAIL              As Print530_MOSIKOMI_MAIL_INF
    Dim dbSQLServer         As Object       'ADODB.Connection
    Dim rsSqlServer         As Object       'ADODB.Recordset
    Dim strConnection       As Object
    Dim strBumonCode        As String
    Dim strAddr()           As String
    Dim wkDate              As Date
    Dim strPR               As String
    Dim strZP               As String
    Dim ret                 As Boolean
    Dim strLog              As String
    Dim strWhere            As String    'ADD 2020/07/27 S.WATANABE
    Dim strSysDate          As String    'ADD 2020/07/27 S.WATANABE
    Dim intSysDateD         As Integer   'ADD 2020/07/27 S.WATANABE
    Dim strColVal           As String    'ADD 2020/07/27 S.WATANABE


    Call MSZZ003_M00(pstrCstFRM_ID, "0", "")

    wkDate = Now
    strMsg = "メールを送信します。よろしいですか？"
    lngRc = MsgBox(strMsg, vbQuestion + vbYesNo, pstrCstKAKUNIN)
    If lngRc = vbNo Then
    Else
        'テンプレートファイル格納パス取得
        varFilePath = DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB=""" & pstrCstFRM_ID & """ AND INTIF_RECFB = ""SENKO_MAIL_FILE_PATH""")
        varFileName = DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB=""" & pstrCstFRM_ID & """ AND INTIF_RECFB = ""SENKO_MAIL_FILE_NAME""")
        strMailSubject = DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB=""" & pstrCstFRM_ID & """ AND INTIF_RECFB = ""SENKO_MAIL_SUBJECT""")
        strUrl = DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB=""" & pstrCstFRM_ID & """ AND INTIF_RECFB = ""URL""")

        strSQL = "SELECT * FROM FVS530_W01 WHERE INTRT_UKNO LIKE ""E*"""
        Set objWk = pobjDB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

        If objWk.EOF = True Then
            MsgBox "該当するデータがありません。", vbInformation, pstrCstERROR
            objWk.Close
            Set objWk = Nothing
            Exit Sub
        End If

        strTempMailMessage = Nz(GetTextMail(varFilePath & varFileName), "null")

        strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))

        cnt = 0
        Do Until objWk.EOF
            strLog = ""

            strSendMailMessage = strTempMailMessage

            strSQL = ""
            'DELETE 2025/12/04 T.KAWABATA START
            'strSQL = strSQL & " SELECT * FROM INTR_TRAN, " & Chr(13)
            'DELETE 2025/12/04 T.KAWABATA END
            'INSERT 2025/12/04 T.KAWABATA START
            strSQL = strSQL & " SELECT INTR_TRAN.*, YOUK_TRAN.*, CNTA_MAST.*, " & Chr(13)
            strSQL = strSQL & " YARD_MAST.YARD_NAME, " & Chr(13)
            'DELETE 2025/12/15 T.KAWABATA START
            'strSQL = strSQL & " YARD_MAST.YARD_ADDR, " & Chr(13)
            'DELETE 2025/12/15 T.KAWABATA END
            'INSERT 2025/12/15 T.KAWABATA START
            strSQL = strSQL & " YARD_MAST.YARD_ADDR_1, " & Chr(13)
            strSQL = strSQL & " YARD_MAST.YARD_ADDR_2, " & Chr(13)
            strSQL = strSQL & " YARD_MAST.YARD_ADDR_3, " & Chr(13)
            'INSERT 2025/12/15 T.KAWABATA END
            strSQL = strSQL & " YARD_MAST.YARD_TIHO_KBN" & Chr(13)
            strSQL = strSQL & " FROM INTR_TRAN, " & Chr(13)
            'INSERT 2025/12/04 T.KAWABATA END
            strSQL = strSQL & " YOUK_TRAN, " & Chr(13)
            strSQL = strSQL & " YARD_MAST, " & Chr(13)
            strSQL = strSQL & " CNTA_MAST " & Chr(13)
            strSQL = strSQL & " WHERE INTR_TRAN.INTRT_UKNO = '" & objWk.Fields("INTRT_UKNO") & "'" & Chr(13)
            strSQL = strSQL & " AND INTR_TRAN.INTRT_YCODE = YARD_MAST.YARD_CODE " & Chr(13)
            strSQL = strSQL & " AND INTR_TRAN.INTRT_YCODE = CNTA_MAST.CNTA_CODE " & Chr(13)
            strSQL = strSQL & " AND INTR_TRAN.INTRT_NO    = CNTA_MAST.CNTA_NO " & Chr(13)
            strSQL = strSQL & " AND INTR_TRAN.INTRT_UKNO  = YOUK_TRAN.YOUKT_UKNO " & Chr(13)
            strSQL = strSQL & " AND INTR_TRAN.INTRT_YCODE = " & objWk.Fields("INTRT_YCODE") & Chr(13)
            strSQL = strSQL & " AND INTR_TRAN.INTRT_NO = " & objWk.Fields("INTRT_NO") & Chr(13)
            Set objWk2 = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

            strSQL = "SELECT * FROM YOUK_TRAN " & Chr(13)
            strSQL = strSQL & " WHERE YOUK_TRAN.YOUKT_UKNO = '" & objWk.Fields("INTRT_UKNO") & "'" & Chr(13)
            strSQL = strSQL & " AND YOUK_TRAN.YOUKT_MLDATE IS NULL " & Chr(13)
            strSQL = strSQL & " AND YOUKT_YUKBN = 2 " & Chr(13)
            Set objYo = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
            If objYo.EOF = True Then
            Else
                stMAIL.strSendTo = Nz(objWk2.Fields("YOUKT_MAIL"))      'メールアドレスは予約トランから取得
                stMAIL.strCONFIRM01 = Replace(Nz(objWk.Fields("YARD_ADDR")), Chr(13) & Chr(10), "")
                stMAIL.strCONFIRM02 = Nz(objWk.Fields("YARD_NAME"))
                stMAIL.strCONFIRM03 = Nz(objWk.Fields("USAGE_NAME"))
                stMAIL.strCONFIRM04 = Nz(objWk.Fields("CNTA_SIZE"))
                stMAIL.strCONFIRM05 = fncGetName_MAST(0, "090", objWk2.Fields("YOUKT_STEP"))

                strSQL = ""
                strSQL = strSQL & " select * from PRIC_TABL dbo_PRIC_TABL where dbo_PRIC_TABL.PRIC_STEP = " & objWk2.Fields("CNTA_STEP") & Chr(13)
                strSQL = strSQL & "      AND dbo_PRIC_TABL.PRIC_SIZE = " & objWk2.Fields("CNTA_SIZE") & Chr(13)
                strSQL = strSQL & "      AND dbo_PRIC_TABL.PRIC_USAGE = " & objWk2.Fields("CNTA_USAGE") & Chr(13)
                strSQL = strSQL & "      AND dbo_PRIC_TABL.PRIC_YCODE = " & objWk2.Fields("INTRT_YCODE") & Chr(13)
                Set objPr = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

                If objPr.EOF = True Then
                Else
                    strPR = Nz(objPr.Fields("PRIC_PRICE"))
                    strZP = Nz(objPr.Fields("PRIC_EZAPPI"))
                    stMAIL.strCONFIRM06 = Format(Nz(objPr.Fields("PRIC_PRICE"), 0), "#,##0")
                    stMAIL.strCONFIRM07 = Format(Nz(objPr.Fields("PRIC_EZAPPI"), 0), "#,##0")
                    stMAIL.strCONFIRM08 = Format(Nz(objPr.Fields("PRIC_PRICE"), 0) + Nz(objPr.Fields("PRIC_EZAPPI"), 0), "#,##0")
                End If
                objPr.Close
                Set objPr = Nothing

                strAddr = Split(Nz(objWk2.Fields("YOUKT_MAIL")), "@")
                stMAIL.strMOSIKOMIURL = strUrl
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "CODE=" & strBumonCode & Nz(objWk2.Fields("INTRT_YCODE"))
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&TYPE=" & Nz(objWk2.Fields("YOUKT_USAGE"))
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&STEP=" & Nz(objWk2.Fields("YOUKT_STEP"))
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&SIZE=" & Nz(objWk2.Fields("CNTA_SIZE"))
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&PR=" & Hex(strPR)
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&ZP=" & Hex(strZP)
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&TIHO=" & Nz(objWk2.Fields("YARD_TIHO_KBN"))
                'stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&NO="                                 'DELETE 2023/05/30 N.IMAI
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&RM=" & Nz(objWk2.Fields("YOUKT_NO"))  'INSERT 2023/05/30 N.IMAI
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&PLN=1"        'INSERT 2018/02/15 N.IMAI
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&MA=" & strAddr(0)
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&MD=" & strAddr(1)
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&UKNO=" & Nz(objWk2.Fields("INTRT_UKNO"))
                stMAIL.strMOSIKOMIURL = stMAIL.strMOSIKOMIURL & "&TADT=" & Format(Nz(objWk2.Fields("YOUKT_TADATE")), "yyyymmdd")
                
                '↓INSERT 2020/07/27 S.WATANABE
                ' 起算日の文言を求める
                strSysDate = Format(DATE$, "YYYYMMDD")
                intSysDateD = Format(DATE$, "D")
                strWhere = "CTRLT_CODE LIKE 'KISAN_PTN*'"
                strWhere = strWhere & " AND CTRLT_KEY01 <= '" & strSysDate & "'"
                strWhere = strWhere & " AND CTRLT_KEY02 >= '" & strSysDate & "'"
                strWhere = strWhere & " AND Val(CTRLT_KEY03) <= " & intSysDateD
                strWhere = strWhere & " AND Val(CTRLT_KEY04) >= " & intSysDateD
                strColVal = Nz(DLookup("CTRLT_VALUE01", "CTRL_TABL", strWhere), "")
                
                '↑INSERT 2020/07/27 S.WATANABE
                strSendMailMessage = Replace(strSendMailMessage, "%SENDTO%", stMAIL.strSendTo & "様")
                strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM01%", "  所在地　：" & stMAIL.strCONFIRM01 & Chr(13))
                strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM02%", "  物件名　：" & stMAIL.strCONFIRM02 & Chr(13))
                strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM03%", "  タイプ　：" & stMAIL.strCONFIRM03 & Chr(13))
                strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM04%", "  サイズ　：" & stMAIL.strCONFIRM04 & " 帖" & Chr(13))
                strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM05%", "  段　　　：" & stMAIL.strCONFIRM05 & Chr(13))
                '↓MOD 2020/07/27 S.WATANABE
                'strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM06%", "  使用料　：" & stMAIL.strCONFIRM06 & " 円 ※賃料発生日は申込日より4日後とさせていただきます。" & Chr(13))
                strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM06%", "  使用料　：" & stMAIL.strCONFIRM06 & " 円 ※賃料発生日は" & strColVal & "とさせていただきます。" & Chr(13))
                'strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM07%", "  共益費　：" & stMAIL.strCONFIRM07 & " 円" & Chr(13))
                'strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM08%", "  請求金額：" & stMAIL.strCONFIRM08 & " 円" & Chr(13))
                If stMAIL.strCONFIRM07 = "0" Then
                    '雑費が0円の場合は表示しない
                    strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM07%", "  請求金額：" & stMAIL.strCONFIRM08 & " 円" & Chr(13))
                    strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM08%", "" & Chr(13))
                Else
                    strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM07%", "  共益費　：" & stMAIL.strCONFIRM07 & " 円" & Chr(13))
                    strSendMailMessage = Replace(strSendMailMessage, "%CONFIRM08%", "  請求金額：" & stMAIL.strCONFIRM08 & " 円" & Chr(13))
                End If
                '↑MOD 2020/07/27 S.WATANABE
                strSendMailMessage = Replace(strSendMailMessage, "%MOSIKOMIURL%", stMAIL.strMOSIKOMIURL)
                
                strLog = " MAIL=" & Nz(objWk2.Fields("YOUKT_MAIL")) & _
                         " UKNO=" & objWk.Fields("INTRT_UKNO") & _
                         " CODE=" & strBumonCode & Nz(objWk2.Fields("INTRT_YCODE")) & _
                         " TYPE=" & Nz(objWk2.Fields("YOUKT_USAGE")) & _
                         " STEP=" & Nz(objWk2.Fields("YOUKT_STEP")) & _
                         " SIZE=" & Nz(objWk2.Fields("CNTA_SIZE"))
                
                ' メール送信処理
                ret = MSZZ021_M00_CUSTOM(pstrCstFRM_ID, stMAIL.strSendTo, "", strMailSubject, strSendMailMessage)
                
                If ret Then

                    objYo.Edit
                    objYo.Fields("YOUKT_MLDATE").VALUE = wkDate
                    objYo.UPDATE
                    cnt = 1
                    
                    strLog = " OK " & strLog

                Else
                    ecnt = 1
                    
                    strLog = " NG " & strLog
                                
                End If
                    
                Call MSZZ003_M00(pstrCstFRM_ID, "8", strLog)
            
            End If
            
            objYo.Close
            Set objYo = Nothing
            objWk2.Close
            Set objWk2 = Nothing

            objWk.MoveNext
        Loop

        objWk.Close
        Set objWk = Nothing

        If cnt = 0 And ecnt = 0 Then
            MsgBox "メールは送信済みです。", vbInformation, pstrCstERROR
        Else
            If ecnt = 0 Then
                MsgBox ("メールを送信しました")
            Else
                MsgBox ("メールを送信しました。一部エラーがあります。" & vbCrLf & "ログファイルをご確認ください。")
            End If
        End If
    End If

    Call MSZZ003_M00(pstrCstFRM_ID, "1", "")
    Exit Sub

psubMail_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Call MSZZ003_M00(pstrCstFRM_ID, "9", "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description)
    Err.Clear

End Sub


'==============================================================================*
'
'       MODULE_NAME     : メール内容の原文取得
'       MODULE_ID       : GetTextMail
'       CREATE_DATE     : 2017/04/01            K.SATO
'       PARAM           : strFileName           ファイル名(I)
'       RETURN          : 原文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetTextMail(ByVal strFilename As String) As String
    Dim strBuff             As String
    Dim strResult           As String
    Dim iFile               As Integer
    On Error GoTo ErrorHandler

    iFile = FreeFile()
    Open strFilename For Input As #iFile
    While Not EOF(iFile)
        Line Input #iFile, strBuff
        strResult = strResult & strBuff & vbCrLf
    Wend
    Close #iFile
    GetTextMail = strResult
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "GetTextMail" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'****************************  ended or program ********************************


