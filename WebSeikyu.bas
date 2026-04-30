Attribute VB_Name = "WebSeikyu"
'****************************  strat or program ********************************
'==============================================================================*
'
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : Web請求関数群
'       PROGRAM_ID      :
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2018/06/01
'       CERATER         : TFC
'       Ver             : 0.0
'
'       UPDATE          : 2019/10/19
'       CERATER         : N.IMAI
'       Ver             : 0.1
'                       : 消費税対応（Web請求書）
'
'       UPDATE          : 2022/08/04
'       CERATER         : N.IMAI
'       Ver             : 0.2
'                       : 圧着はがきCSV出力
'
'       UPDATE          : 2022/09/05
'       CERATER         : N.IMAI
'       Ver             : 0.3
'                       : 口座名義人対応
'
'       UPDATE          : 2023/02/28
'       CERATER         : N.IMAI
'       Ver             : 0.4
'                       : インボイス対応
'
'       UPDATE          : 2023/03/31
'       CERATER         : N.IMAI
'       Ver             : 0.5
'                       : 顧客住所２と住所３の間に全角スペースを入れる
'
'       UPDATE          : 2023/07/05
'       CERATER         : N.IMAI
'       Ver             : 0.6
'                       : 消費税がintのためオーバーフローする問題に対応
'
'       UPDATE          : 2023/10/31
'       CERATER         : N.IMAI
'       Ver             : 0.7
'                       : NYKO_MASTを変更すると自動入金で入金しなくなるので
'                       　2023/11/30まで用の修正
'
'==============================================================================*
'Option Explicit
'Option Compare Database

'Global変数
Private sBoundy         As String           'Boundy
Private g_dtnow         As Date             '請求書発行日時

'-----------------------------------
'       << 構造体 >>
'-----------------------------------
'-----------------------------------
'       << CONST 定数定義 >>
'-----------------------------------
'------- << HTTP構文 >>
Private Const C_HED_TOK = "X-WB-apitoken:"
Private Const C_HED_CONT = "Content-Type"
Private Const C_HED_CONT_MULTI = "multipart/form-data;"
Private Const C_HED_CONT_APL = "application/json;"
Private Const C_HED_CONT_TXT = "text/plain;"

'------- << HTTP書式 >>
Private Const adTypeBinary = 1
Private Const adTypeText = 2
Private Const adBTypeContent = 1
Private Const adBTypeBody = 2
Private Const adBTypeFooter = 3

'------- << レスポンスコード >>
Private Const C_RES_MENT = 503                  'メンテナンス中
Private Const C_RES_OK = 200                    '成功

'------- << ステータス >>
Private Const C_STS_SUCCESS = "success"         '成功
Private Const C_STS_ERROR = "error"             'エラー
Private Const C_STS_ACTIVE = "active"           '取込中
Private Const C_STS_COMP = "complete"           '正常終了

'------- << 各種CSVヘッダ情報 >>
'顧客CSV
Private Const C_CUSTCOL = "顧客コード,顧客名,郵便番号,住所１,住所２,会社名,担当者名,お客様コード" & _
                            ",発行元郵便番号,発行元住所,発行元会社名,発行元TEL,発行元FAX,法人個人区分" & _
                            ",メールアドレス,サブメールアドレス１,識別コード"
'帳票CSV
'Private Const C_PRTCOL = "日付,顧客コード,ページ,ご請求金額,振込期限日,請求年月,明細名称,請求明細No." & _
                            ",明細金額,金融機関,口座番号,口座名義,振込人名,備考"                                                                    'DELETE 2019/10/19 N.IMAI
Private Const C_PRTCOL = "日付,顧客コード,ページ,ご請求金額,振込期限日,請求年月,明細名称,請求明細No." & _
                            ",明細金額,金融機関,口座番号,口座名義,振込人名,備考,登録番号,社判,テキスト1,テキスト2,テキスト3,テキスト4,テキスト5"    'INSERT 2019/10/19 N.IMAI
Private Const C_PRTCOL2 = "日付,顧客コード,ページ,ご請求金額,振込期限日,請求年月,明細名称,請求明細No." & _
                            ",明細金額,金融機関,口座番号,口座名義,振込人名,備考,登録番号,社判,テキスト1,テキスト2,テキスト3,テキスト4,テキスト5,お客様番号,ヤード番号,部屋番号"
'------- << 定数 >>
'プログラムID
Private Const P_PROGRAM_ID  As String = "WEBINV"
'WEB請求開始年月(下記年月＜計上年月の場合、前月計上の顧客に存在しない顧客データをアップロード)
Private Const C_STARTMON    As Date = #3/1/2020#
'ログ出力有無フラグ(0:ログ出力なし,1:ログ出力あり)
Private Const C_LOG         As Integer = 1
'帳票承認フラグ(0:承認しない,1:承認する)
Private Const C_APPROVE     As Integer = 0

'==============================================================================*
'
'        PROGRAM_NAME    :プログラムID Getter
'        PROGRAM_ID      :
'        PROGRAM_KBN     :
'
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'
'        NOTE            :プログラムID Getter
'
'==============================================================================*
Public Function GetProgId() As String

    GetProgId = P_PROGRAM_ID
End Function

'==============================================================================*
'
'        PROGRAM_NAME    :ログ設定値 Getter
'        PROGRAM_ID      :
'        PROGRAM_KBN     :
'
'        CREATE          : 2018/06/30
'        CERATER         : TFC
'
'        NOTE            :ログ設定値 Getter
'
'==============================================================================*
Public Function GetLog() As Integer

    GetLog = C_LOG
End Function

'==============================================================================*
'
'        MODULE_NAME      :顧客CSVと帳票CSVを作成する
'        MODULE_ID        :FncOut_Csv
'        PARAMETER        : I/    strFileCT     顧客CSVファイル名
'                           I/    strFilePT     帳票CSVファイル名
'                           I/    strMonth      計上年月(yyyymm)
'        RETURN           : 0:正常, 0以外:異常
'        CREATE_DATE      :2018.06.30   TFC
'        UPDATE_DATE      :
'
'==============================================================================*
Public Function FncOut_Csv(strFileCT As String, strFilePT As String, strMonth As String) As Integer

    Dim iFileCT         As Integer      'ファイル番号(顧客CSV)
    Dim iFilePT         As Integer      'ファイル番号(帳票CSV)
    Dim objRs           As Object       'RecordSet Object
    Dim objRs2          As Object       'RecordSet Object
    Dim varCOLCT(17)    As Variant      '顧客CSV項目
'   Dim varCOLPT(14)    As Variant      '帳票CSV項目                            'DELETE 2019/10/19 N.IMAI
    Dim varCOLPT(21)    As Variant      '帳票CSV項目                            'INSERT 2019/10/19 N.IMAI
    Dim strSQL          As String       'SQL スクリプト文字列
    Dim varHeadCT       As Variant      '顧客CSV項目タイトル
    Dim varHeadPT       As Variant      '帳票CSV項目タイトル
    Dim dtKigen         As Date         '振込期限
    Dim strUserCode     As String       '顧客コード(部門付き)
    Dim strMsg          As String       'ログ出力用メッセージ
    Dim strUser         As String       '比較用
    Dim strCAMPC        As String       '会社コード                             'INSERT 2019/10/19 N.IMAI
    Dim strTourokuNo    As String       '請求書登録番号                         'INSERT 2019/10/19 N.IMAI
    Dim intTaxRate      As Integer      '消費税                                 'INSERT 2019/10/19 N.IMAI
    Dim lngKinga        As Long         '金額                                   'INSERT 2019/10/19 N.IMAI
    'Dim intTax          As Integer      '消費税額                              'DELETE 2023/07/05 N.IMAI 'INSERT 2019/10/19 N.IMAI
    Dim intTax          As Long         '消費税額                               'INSERT 2023/07/05 N.IMAI

    On Error GoTo ErrorHandler
    
    If C_LOG = 1 Then
        strMsg = "CSV出力処理開始 " & "計上年月:" & strMonth & " 顧客CSV:" & strFileCT & " 帳票CSV:" & strFilePT
        Call MSZZ003_M00(GetProgId(), "0", strMsg)
    End If
    
    strCAMPC = DLookup("CONT_CAMPC", "dbo_CONT_MAST", "CONT_KEY=1")                         'INSERT 2019/10/19 N.IMAI
    strTourokuNo = Nz(DLookup("CONT_SEIKYUSYO_TOUROKU_NO", "dbo_CONT_MAST", "CONT_KEY=1"))  'INSERT 2019/10/19 N.IMAI
    
    '出力データ取得
    Set objRs = CurrentDb.OpenRecordset(Get_strHeader(strMonth), dbOpenSnapshot)
    On Error GoTo ErrorHandler1
    
    If objRs.EOF = True Then
        GoTo ErrorHandler1
    End If
    
    'CSVを作成する
    iFileCT = FreeFile()
    Open strFileCT For Output Lock Write As iFileCT
    
    iFilePT = FreeFile()
    Open strFilePT For Output Lock Write As iFilePT
    
    'ヘッダ項目名セット(1行目)
    varHeadCT = Split(C_CUSTCOL, ",")
    Write #iFileCT, varHeadCT(0), varHeadCT(1), varHeadCT(2), varHeadCT(3), varHeadCT(4), _
                    varHeadCT(5), varHeadCT(6), varHeadCT(7), varHeadCT(8), varHeadCT(9), _
                    varHeadCT(10), varHeadCT(11), varHeadCT(12), varHeadCT(13), varHeadCT(14), _
                    varHeadCT(15), varHeadCT(16)
   
    varHeadPT = Split(C_PRTCOL, ",")
'   Write #iFilePT, varHeadPT(0), varHeadPT(1), varHeadPT(2), varHeadPT(3), varHeadPT(4), _
                    varHeadPT(5), varHeadPT(6), varHeadPT(7), varHeadPT(8), varHeadPT(9), _
                    varHeadPT(10), varHeadPT(11), varHeadPT(12), varHeadPT(13)                      'DELETE 2019/10/19 N.IMAI
    Write #iFilePT, varHeadPT(0), varHeadPT(1), varHeadPT(2), varHeadPT(3), varHeadPT(4), _
                    varHeadPT(5), varHeadPT(6), varHeadPT(7), varHeadPT(8), varHeadPT(9), _
                    varHeadPT(10), varHeadPT(11), varHeadPT(12), varHeadPT(13), varHeadPT(14), _
                    varHeadPT(15), varHeadPT(16), varHeadPT(17), varHeadPT(18), varHeadPT(19), _
                    varHeadPT(20)                                                                   'INSERT 2019/10/19 N.IMAI
    
    '顧客CSV
    Do Until objRs.EOF

'DELETE 2019/10/19 N.IMAI Start
'        '顧客が変わったら強制的に一行差し込み
'        If strUser <> "" And strUser <> Nz(objRs![NYKOM_BUMOC], "") & "-" & Format(objRs![USER_CODE], "000000") Then
'            '帳票CSV Write
'            Write #iFilePT, varCOLPT(0), varCOLPT(1), varCOLPT(2), varCOLPT(3), varCOLPT(4), _
'                varCOLPT(5), "", "消費税10%込み", "", varCOLPT(9), _
'                varCOLPT(10), varCOLPT(11), varCOLPT(12), varCOLPT(13)
'        End If
'DELETE 2019/10/19 N.IMAI End

        strUser = Nz(objRs![NYKOM_BUMOC], "") & "-" & Format(objRs![USER_CODE], "000000")

    
        strUserCode = Nz(objRs![NYKOM_BUMOC], "") & "-" & Format(objRs![USER_CODE], "000000")
        '前月に存在しない顧客をCSV出力
        If objRs![PRVFLG] <= 0 Then
            'ユーザコードのBreakにてCSV出力
            If varCOLCT(0) <> strUserCode Then
                varCOLCT(0) = strUserCode                               '顧客コード
                varCOLCT(1) = Nz(objRs![USER_NAME], "")                 '顧客名
                
                '連絡先担当者に半角ｽﾍﾟｰｽが入ってる可能性があるため。
                If Nz(objRs![USER_RNAME], "") = "" Or Nz(objRs![USER_RNAME], "") = " " Then
                    varCOLCT(2) = Nz(objRs![USER_YUBINO], "")           '郵便番号
                    varCOLCT(3) = Nz(objRs![USER_ADR_1], "")            '住所１
                    varCOLCT(4) = Nz(objRs![USER_ADR_2], "") & _
                                    Nz(objRs![USER_ADR_3], "")          '住所２+住所３
                    varCOLCT(14) = Nz(objRs![USER_MAIL], "")            'メールアドレス
                    varCOLCT(16) = Nz(objRs![USER_YUBINO], "")          '識別コード(郵便番号)
                Else
                    varCOLCT(2) = Nz(objRs![USER_RPOST], "")            '郵便番号
                    varCOLCT(3) = Nz(objRs![USER_RADR_1], "")           '住所１
                    varCOLCT(4) = Nz(objRs![USER_RADR_2], "") & _
                                    Nz(objRs![USER_RADR_3], "")         '住所２+住所３
                    varCOLCT(14) = Nz(objRs![USER_RMAIL], "")           'メールアドレス
                    varCOLCT(16) = Nz(objRs![USER_RPOST], "")           '識別コード(郵便番号)
                End If
                
                
                '担当者名
                varCOLCT(6) = IIf(Nz(objRs![USER_RNAME], "") = "", Nz(objRs![USER_TANM], ""), _
                                    Nz(objRs![USER_RNAME], ""))

                If varCOLCT(6) = " " Then
                    '会社名
                    varCOLCT(5) = Nz(objRs![USER_NAME], "") & IIf(Nz(objRs![USER_KKBN], "") = 1, " 御中", " 様")
                Else
                    '会社名
                    varCOLCT(5) = Nz(objRs![USER_NAME], "")
                    varCOLCT(6) = varCOLCT(6) & "様"
                End If
                
                'お客様コード
                varCOLCT(7) = "( お客様コード ： " & Nz(objRs![NYKOM_BUMOC], "") & "-" & _
                                    Format(objRs![USER_CODE], "000000") & " )"
                '発行元郵便番号
                varCOLCT(8) = "〒" & DLookup("CONT_YUBINO", "dbo_CONT_MAST", "CONT_KEY=1")
                '発行元住所
                varCOLCT(9) = DLookup("CONT_ADDR_1", "dbo_CONT_MAST", "CONT_KEY=1") & _
                                    DLookup("CONT_ADDR_2", "dbo_CONT_MAST", "CONT_KEY=1")
                '発行元会社名
                varCOLCT(10) = DLookup("CONT_KAISYA", "dbo_CONT_MAST", "CONT_KEY=1")
                '発行元TEL
                varCOLCT(11) = "T E L : " & DLookup("CONT_TEL_NO", "dbo_CONT_MAST", "CONT_KEY=1")
                '発行元FAX
                varCOLCT(12) = "F A X : " & DLookup("CONT_FAX_NO", "dbo_CONT_MAST", "CONT_KEY=1")
                '法人・個人区分
                varCOLCT(13) = " "
                'サブメールアドレス１
                varCOLCT(15) = " "
                
                '顧客CSV Write
                Write #iFileCT, varCOLCT(0), varCOLCT(1), varCOLCT(2), varCOLCT(3), varCOLCT(4), _
                                varCOLCT(5), varCOLCT(6), varCOLCT(7), varCOLCT(8), varCOLCT(9), _
                                varCOLCT(10), varCOLCT(11), varCOLCT(12), varCOLCT(13), varCOLCT(14), _
                                varCOLCT(15), varCOLCT(16)
                                
            End If
        End If
        
        'INSERT 2019/10/19 N.IMAI Start
        intTaxRate = Nz(objRs![TAX_RATE], 0)
        lngKinga = Nz(objRs![RKS170_KINGAKの合計], 0)
        intTax = Nz(objRs![TAX], 0)
        'INSERT 2019/10/19 N.IMAI End
        
        '請求WKからデータ取得
        strSQL = strOutputDetail(objRs![RKS170_KCODE], objRs![RKS170_PAGE])
        Set objRs2 = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot)
        On Error GoTo ErrorHandler2
    
        If objRs2.EOF = False Then
            '帳票CSV
            Do Until objRs2.EOF
                '発行日(当日)
                If Nz(Forms!FKS240!txt_Hkobi) = "" Then
                    varCOLPT(0) = Year(Now()) & "年" & Month(Now()) & _
                                    "月" & Day(Now()) & "日"
                Else
                    varCOLPT(0) = Left(Forms!FKS240!txt_Hkobi, 4) & "年" & Mid(Forms!FKS240!txt_Hkobi, 5, 2) & _
                                    "月" & Right(Forms!FKS240!txt_Hkobi, 2) & "日"
                End If
                '顧客コード
                varCOLPT(1) = strUserCode
                'ページ
                varCOLPT(2) = objRs![RKS170_PAGE]
                'ご請求金額(MAX10桁)
                varCOLPT(3) = Nz(objRs![RKS170_KINGAKの合計], 0)
                '振込期限(発行日の月末)
                varCOLPT(4) = Left(Forms!FKS240!txt_Fkbi, 4) & "年" & _
                                    Mid(Forms!FKS240!txt_Fkbi, 5, 2) & "月" & Right(Forms!FKS240!txt_Fkbi, 2) & "日"
                '請求年月
                varCOLPT(5) = Mid(Format(objRs2![RKS170_KJDATE], "yyyymm"), 1, 4) & "年" & _
                                    Mid(Format(objRs2![RKS170_KJDATE], "yyyymm"), 5, 2) & "月"
                '明細名称
                varCOLPT(6) = Nz(objRs2![RKS170_YARDNAME], "") & "(" & Nz(objRs2![RKS170_YCODE], "") & ")"
                '請求明細No.
                varCOLPT(7) = IIf(Nz(objRs2![RKS170_NO], "") = "", "", "No." & _
                                    (Format(Nz(objRs2![RKS170_NO], 0), "000000")))
                '明細金額(MAX8桁)
                varCOLPT(8) = Nz(objRs2![RKS170_KINGAK], 0) + Nz(objRs2![RKS170_SYOZEI], 0)
                
                If Nz(objRs![KOUZM_KOUZB], "") = "" Then
                    '金融機関
                    varCOLPT(9) = Nz(objRs![BANKT_KINYN], "") & IIf(Nz(objRs![BANKT_KINYI], "") = "2", "銀行", "") & _
                                    "(" & Nz(objRs![NYKOM_KINYC], "") & ") " & Nz(objRs![BANKT_SHITN], "") & _
                                    "(" & Nz(objRs![NYKOM_SHITC], "") & ")"
                    '口座番号
                    'varCOLPT(10) = Nz(objRs![YOKIN], "") & " " & Nz(objRs![NYKOM_KOUZB], "")
                    'INSERT 2023/10/31 N.IMAI Start
                    If Nz(objRs![NYKOM_KINYC], "") = "0010" Then
                        varCOLPT(10) = "当座預金 0534679"
                    Else
                        varCOLPT(10) = Nz(objRs![YOKIN], "") & " " & Nz(objRs![NYKOM_KOUZB], "")
                    End If
                    'INSERT 2023/10/31 N.IMAI End
                
                    '振込人名
                    varCOLPT(12) = Nz(objRs![CAMPM_CASEQ], "") & Nz(objRs![BUMOM_BUSEQ], "") & _
                                    Format(Nz(objRs![USER_CODE], 0), "000000") & Nz(objRs![USER_KANA], "")
                
                    '口座名義
                    varCOLPT(11) = Nz(objRs![NYKOM_KOUZN], "")
                    'If Nz(objRs![NYKOM_KINYC], "") = "0010" Then    'りそな
                    '    varCOLPT(11) = "株式会社 加瀬倉庫"
                    'Else
                        varCOLPT(11) = Nz(objRs![KOUZM_KOUZN], objRs![NYKOM_KOUZN])
                        'If strCAMPC <> "KAS" Then
                        If (Nz(objRs![NYKOM_BUMOC], "") <> "1" And varCOLPT(11) = "株式会社　加瀬不動産活用") _
                        Or (Nz(objRs![NYKOM_BUMOC], "") = "1" And varCOLPT(11) = "株式会社　加瀬倉庫") Then
                            varCOLPT(11) = varCOLPT(11) & "　※収納代行会社"
                        End If
                    'End If
                Else
                    '金融機関
                    varCOLPT(9) = Nz(objRs![KOUZM_KINYN], "") & "(" & Nz(objRs![KOUZM_KINYC], "") & ") " & _
                                    Nz(objRs![KOUZM_SHITN], "") & "(" & Nz(objRs![KOUZM_SHITC], "") & ")"
                    '口座番号
                    varCOLPT(10) = Nz(objRs![KOUZM_YOKIN], "") & " " & Nz(objRs![KOUZM_KOUZB], "")
                
                    '振込人名
                    varCOLPT(12) = Nz(objRs![USER_KANA], "")
                    
                    '口座名義
                    If Nz(objRs![KOUZM_KINYC], "") = "0310" Then
                        varCOLPT(11) = "株式会社　加瀬倉庫"
                    Else
                        varCOLPT(11) = Nz(objRs![KOUZM_KOUZN], objRs![NYKOM_KOUZN])
                        'varCOLPT(11) = "株式会社　加瀬不動産活用"
                        'If strCAMPC <> "KAS" Then
                    End If
                    If (Nz(objRs![NYKOM_BUMOC], "") <> "1" And varCOLPT(11) = "株式会社　加瀬不動産活用") _
                    Or (Nz(objRs![NYKOM_BUMOC], "") = "1" And varCOLPT(11) = "株式会社　加瀬倉庫") Then
                        varCOLPT(11) = varCOLPT(11) & "　※収納代行会社"
                    End If
                End If
                '口座名義
                'varCOLPT(11) = Nz(objRs![NYKOM_KOUZN], "")
                
                '備考
                varCOLPT(13) = " "
                
                varCOLPT(14) = strTourokuNo                                     'INSERT 2023/02/28 N.IMAI
                
                'INSERT 2019/10/19 N.IMAI Start
                '社判
                varCOLPT(15) = strCAMPC
                'テキスト1
                varCOLPT(16) = intTaxRate & "%対象"
                'テキスト2
                varCOLPT(17) = lngKinga
                'テキスト3
                varCOLPT(18) = "消費税"
                'テキスト4
                varCOLPT(19) = intTax
                'INSERT 2019/10/19 N.IMAI End
              
                '帳票CSV Write
'               Write #iFilePT, varCOLPT(0), varCOLPT(1), varCOLPT(2), varCOLPT(3), varCOLPT(4), _
                                varCOLPT(5), varCOLPT(6), varCOLPT(7), varCOLPT(8), varCOLPT(9), _
                                varCOLPT(10), varCOLPT(11), varCOLPT(12), varCOLPT(13)                      'DELETE 2019/10/19 N.IMAI
                Write #iFilePT, varCOLPT(0), varCOLPT(1), varCOLPT(2), varCOLPT(3), varCOLPT(4), _
                                varCOLPT(5), varCOLPT(6), varCOLPT(7), varCOLPT(8), varCOLPT(9), _
                                varCOLPT(10), varCOLPT(11), varCOLPT(12), varCOLPT(13), varCOLPT(14), _
                                varCOLPT(15), varCOLPT(16), varCOLPT(17), varCOLPT(18), varCOLPT(19), _
                                varCOLPT(20)                                                                'INSERT 2019/10/19 N.IMAI
                objRs2.MoveNext
            Loop
        End If
        
        objRs2.Close
        Set objRs2 = Nothing
        
        objRs.MoveNext
    Loop
    
'DELETE 2019/10/19 N.IMAI Start
'            '帳票CSV Write
'            Write #iFilePT, varCOLPT(0), varCOLPT(1), varCOLPT(2), varCOLPT(3), varCOLPT(4), _
'                varCOLPT(5), "", "消費税10%込み", "", varCOLPT(9), _
'                varCOLPT(10), varCOLPT(11), varCOLPT(12), varCOLPT(13)
'DELETE 2019/10/19 N.IMAI End

    Close #iFileCT
    Close #iFilePT
    
    objRs.Close
    Set objRs = Nothing
    
    If C_LOG = 1 Then
        strMsg = "CSV出力処理終了 "
        Call MSZZ003_M00(GetProgId(), "1", strMsg)
    End If
    
    FncOut_Csv = 0
        
ErrorHandler2:
    If Not (objRs2 Is Nothing) Then
        objRs2.Close
        Set objRs2 = Nothing
    End If
    If Err <> 0 Then
        Close #iFileCT
        Close #iFilePT
        If C_LOG = 1 Then
            strMsg = "CSV出力処理異常 Err.Number=" & Err.Number & _
                        " Err.Source=" & Err.Source & " Err.Description=" & Err.Description
            Call MSZZ003_M00(GetProgId(), "9", strMsg)
        End If
        Err.Clear
        FncOut_Csv = 1
    End If
    
ErrorHandler1:
    If Not (objRs Is Nothing) Then
        objRs.Close
        Set objRs = Nothing
    End If
    If Err <> 0 Then
        Close #iFileCT
        Close #iFilePT
        If C_LOG = 1 Then
            strMsg = "CSV出力処理異常 Err.Number=" & Err.Number & _
                        " Err.Source=" & Err.Source & " Err.Description=" & Err.Description
            Call MSZZ003_M00(GetProgId(), "9", strMsg)
        End If
        Err.Clear
        FncOut_Csv = 2
    End If

ErrorHandler:
    If Err <> 0 Then
        If C_LOG = 1 Then
            strMsg = "CSV出力処理異常 Err.Number=" & Err.Number & _
                        " Err.Source=" & Err.Source & " Err.Description=" & Err.Description
            Call MSZZ003_M00(GetProgId(), "9", strMsg)
        End If
        FncOut_Csv = 3
    End If
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :圧着葉書用顧客CSVと帳票CSVを作成する
'        MODULE_ID        :FncOut_Csv
'        PARAMETER        : I/    strFileCT     顧客CSVファイル名
'                           I/    strFilePT     帳票CSVファイル名
'                           I/    strMonth      計上年月(yyyymm)
'        RETURN           : 0:正常, 0以外:異常
'        CREATE_DATE      :
'        UPDATE_DATE      :
'
'==============================================================================*
Public Function FncOut_Csv2(strFileCT As String, strFilePT As String, strMonth As String) As Integer

    Dim iFileCT         As Integer      'ファイル番号(顧客CSV)
    Dim iFilePT         As Integer      'ファイル番号(帳票CSV)
    Dim objRs           As Object       'RecordSet Object
    Dim objRs2          As Object       'RecordSet Object
    Dim varCOLCT(17)    As Variant      '顧客CSV項目
    Dim varCOLPT(23)    As Variant      '帳票CSV項目
    Dim strSQL          As String       'SQL スクリプト文字列
    Dim varHeadCT       As Variant      '顧客CSV項目タイトル
    Dim varHeadPT       As Variant      '帳票CSV項目タイトル
    Dim dtKigen         As Date         '振込期限
    Dim strUserCode     As String       '顧客コード(部門付き)
    Dim strMsg          As String       'ログ出力用メッセージ
    Dim strUser         As String       '比較用
    Dim strCAMPC        As String       '会社コード
    Dim strTourokuNo    As String       '請求書登録番号
    Dim intTaxRate      As Integer      '消費税
    Dim lngKinga        As Long         '金額
    'Dim intTax          As Integer      '消費税額                              'DELETE 2023/07/05 N.IMAI 'INSERT 2019/10/19 N.IMAI
    Dim intTax          As Long         '消費税額                               'INSERT 2023/07/05 N.IMAI

    On Error GoTo ErrorHandler
    
    If C_LOG = 1 Then
        strMsg = "CSV出力処理開始 " & "計上年月:" & strMonth & " 顧客CSV:" & strFileCT & " 帳票CSV:" & strFilePT
        Call MSZZ003_M00(GetProgId(), "0", strMsg)
    End If
    
    strCAMPC = DLookup("CONT_CAMPC", "dbo_CONT_MAST", "CONT_KEY=1")
    strTourokuNo = Nz(DLookup("CONT_SEIKYUSYO_TOUROKU_NO", "dbo_CONT_MAST", "CONT_KEY=1"))
    
    '出力データ取得
    Set objRs = CurrentDb.OpenRecordset(Get_strHeader(strMonth), dbOpenSnapshot)
    On Error GoTo ErrorHandler1
    
    If objRs.EOF = True Then
        GoTo ErrorHandler1
    End If
    
    'CSVを作成する
    iFileCT = FreeFile()
    Open strFileCT For Output Lock Write As iFileCT
    
    iFilePT = FreeFile()
    Open strFilePT For Output Lock Write As iFilePT
    
    'ヘッダ項目名セット(1行目)
    varHeadCT = Split(C_CUSTCOL, ",")
    Write #iFileCT, varHeadCT(0), varHeadCT(1), varHeadCT(2), varHeadCT(3), varHeadCT(4), _
                    varHeadCT(5), varHeadCT(6), varHeadCT(7), varHeadCT(8), varHeadCT(9), _
                    varHeadCT(10), varHeadCT(11), varHeadCT(12), varHeadCT(13), varHeadCT(14), _
                    varHeadCT(15), varHeadCT(16)
   
    varHeadPT = Split(C_PRTCOL2, ",")
   Write #iFilePT, varHeadPT(0), varHeadPT(1), varHeadPT(2), varHeadPT(3), varHeadPT(4), _
                    varHeadPT(5), varHeadPT(6), varHeadPT(7), varHeadPT(8), varHeadPT(9), _
                    varHeadPT(10), varHeadPT(11), varHeadPT(12), varHeadPT(13), varHeadPT(14), _
                    varHeadPT(15), varHeadPT(16), varHeadPT(17), varHeadPT(18), varHeadPT(19), _
                    varHeadPT(20), varHeadPT(21), varHeadPT(22), varHeadPT(23)
    
    '顧客CSV
    Do Until objRs.EOF

'

        strUser = Nz(objRs![NYKOM_BUMOC], "") & "-" & Format(objRs![USER_CODE], "000000")

    
        strUserCode = Nz(objRs![NYKOM_BUMOC], "") & "-" & Format(objRs![USER_CODE], "000000")
        '前月に存在しない顧客をCSV出力
        If objRs![PRVFLG] <= 0 Then
            'ユーザコードのBreakにてCSV出力
            If varCOLCT(0) <> strUserCode Then
                varCOLCT(0) = strUserCode                               '顧客コード
                varCOLCT(1) = Nz(objRs![USER_NAME], "")                 '顧客名
                
                '連絡先担当者に半角ｽﾍﾟｰｽが入ってる可能性があるため。
                If Nz(objRs![USER_RNAME], "") = "" Or Nz(objRs![USER_RNAME], "") = " " Then
                    varCOLCT(2) = Nz(objRs![USER_YUBINO], "")           '郵便番号
                    varCOLCT(3) = Nz(objRs![USER_ADR_1], "")            '住所１
                    'varCOLCT(4) = Nz(objRs![USER_ADR_2], "") & _               'DELETE 2023/03/31 N.IMAI
                    varCOLCT(4) = Nz(objRs![USER_ADR_2], "") & "　" & _
                                    Nz(objRs![USER_ADR_3], "")          '住所２+住所３
                    varCOLCT(14) = Nz(objRs![USER_MAIL], "")            'メールアドレス
                    varCOLCT(16) = Nz(objRs![USER_YUBINO], "")          '識別コード(郵便番号)
                Else
                    varCOLCT(2) = Nz(objRs![USER_RPOST], "")            '郵便番号
                    varCOLCT(3) = Nz(objRs![USER_RADR_1], "")           '住所１
                    'varCOLCT(4) = Nz(objRs![USER_RADR_2], "") & _              'DELETE 2023/03/31 N.IMAI
                    varCOLCT(4) = Nz(objRs![USER_RADR_2], "") & "　" & _
                                    Nz(objRs![USER_RADR_3], "")         '住所２+住所３
                    varCOLCT(14) = Nz(objRs![USER_RMAIL], "")           'メールアドレス
                    varCOLCT(16) = Nz(objRs![USER_RPOST], "")           '識別コード(郵便番号)
                End If
                
                
                '担当者名
                varCOLCT(6) = IIf(Nz(objRs![USER_RNAME], "") = "", Nz(objRs![USER_TANM], ""), _
                                    Nz(objRs![USER_RNAME], ""))

                If varCOLCT(6) = " " Or varCOLCT(6) = "" Then
                    '会社名
                    varCOLCT(5) = Nz(objRs![USER_NAME], "") & IIf(Nz(objRs![USER_KKBN], "") = 1, " 御中", " 様")
                Else
                    '会社名
                    varCOLCT(5) = Nz(objRs![USER_NAME], "")
                    varCOLCT(6) = varCOLCT(6) & "様"
                End If
                
                'お客様コード
                varCOLCT(7) = "( お客様コード ： " & Nz(objRs![NYKOM_BUMOC], "") & "-" & _
                                    Format(objRs![USER_CODE], "000000") & " )"
                '発行元郵便番号
                varCOLCT(8) = "〒" & DLookup("CONT_YUBINO", "dbo_CONT_MAST", "CONT_KEY=1")
                '発行元住所
                varCOLCT(9) = DLookup("CONT_ADDR_1", "dbo_CONT_MAST", "CONT_KEY=1") & _
                                    DLookup("CONT_ADDR_2", "dbo_CONT_MAST", "CONT_KEY=1")
                '発行元会社名
                varCOLCT(10) = DLookup("CONT_KAISYA", "dbo_CONT_MAST", "CONT_KEY=1")
                '発行元TEL
                varCOLCT(11) = "T E L : " & DLookup("CONT_TEL_NO", "dbo_CONT_MAST", "CONT_KEY=1")
                '発行元FAX
                varCOLCT(12) = "F A X : " & DLookup("CONT_FAX_NO", "dbo_CONT_MAST", "CONT_KEY=1")
                '法人・個人区分
                varCOLCT(13) = " "
                'サブメールアドレス１
                varCOLCT(15) = " "
                
                '顧客CSV Write
                Write #iFileCT, varCOLCT(0), varCOLCT(1), varCOLCT(2), varCOLCT(3), varCOLCT(4), _
                                varCOLCT(5), varCOLCT(6), varCOLCT(7), varCOLCT(8), varCOLCT(9), _
                                varCOLCT(10), varCOLCT(11), varCOLCT(12), varCOLCT(13), varCOLCT(14), _
                                varCOLCT(15), varCOLCT(16)
                                
            End If
        End If
        
        intTaxRate = Nz(objRs![TAX_RATE], 0)
        lngKinga = Nz(objRs![RKS170_KINGAKの合計], 0)
        intTax = Nz(objRs![TAX], 0)
           
        '請求WKからデータ取得
        strSQL = strOutputDetail(objRs![RKS170_KCODE], objRs![RKS170_PAGE])
        Set objRs2 = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot)
        On Error GoTo ErrorHandler2
    
        If objRs2.EOF = False Then
            '帳票CSV
            Do Until objRs2.EOF
                '発行日(当日)
                If Nz(Forms!FKS240!txt_Hkobi) = "" Then
                    varCOLPT(0) = Year(Now()) & "年" & Month(Now()) & _
                                    "月" & Day(Now()) & "日"
                Else
                    varCOLPT(0) = Left(Forms!FKS240!txt_Hkobi, 4) & "年" & Mid(Forms!FKS240!txt_Hkobi, 5, 2) & _
                                    "月" & Right(Forms!FKS240!txt_Hkobi, 2) & "日"
                End If
                '顧客コード
                varCOLPT(1) = strUserCode
                'ページ
                varCOLPT(2) = objRs![RKS170_PAGE]
                'ご請求金額(MAX10桁)
                varCOLPT(3) = Nz(objRs![RKS170_KINGAKの合計], 0)
                '振込期限(発行日の月末)
                varCOLPT(4) = Left(Forms!FKS240!txt_Fkbi, 4) & "年" & _
                                    Mid(Forms!FKS240!txt_Fkbi, 5, 2) & "月" & Right(Forms!FKS240!txt_Fkbi, 2) & "日"
                '請求年月
                varCOLPT(5) = Mid(Format(objRs2![RKS170_KJDATE], "yyyymm"), 1, 4) & "年" & _
                                    Mid(Format(objRs2![RKS170_KJDATE], "yyyymm"), 5, 2) & "月"
                '明細名称
                varCOLPT(6) = Nz(objRs2![RKS170_YARDNAME], "") & "(" & Nz(objRs2![RKS170_YCODE], "") & ")"
                '請求明細No.
                varCOLPT(7) = IIf(Nz(objRs2![RKS170_NO], "") = "", "", "No." & _
                                    (Format(Nz(objRs2![RKS170_NO], 0), "000000")))
                '明細金額(MAX8桁)
                varCOLPT(8) = Nz(objRs2![RKS170_KINGAK], 0) + Nz(objRs2![RKS170_SYOZEI], 0)
                
                If Nz(objRs![KOUZM_KOUZB], "") = "" Then
                    '金融機関
                    varCOLPT(9) = Nz(objRs![BANKT_KINYN], "") & IIf(Nz(objRs![BANKT_KINYI], "") = "2", "銀行", "") & _
                                    "(" & Nz(objRs![NYKOM_KINYC], "") & ") " & Nz(objRs![BANKT_SHITN], "") & _
                                    "(" & Nz(objRs![NYKOM_SHITC], "") & ")"
                    '口座番号
                    'varCOLPT(10) = Nz(objRs![YOKIN], "") & " " & Nz(objRs![NYKOM_KOUZB], "")
                    'INSERT 2023/10/31 N.IMAI Start
                    If Nz(objRs![NYKOM_KINYC], "") = "0010" Then
                        varCOLPT(10) = "当座預金 0534679"
                    Else
                        varCOLPT(10) = Nz(objRs![YOKIN], "") & " " & Nz(objRs![NYKOM_KOUZB], "")
                    End If
                    'INSERT 2023/10/31 N.IMAI End
                
                    '振込人名
                    varCOLPT(12) = Nz(objRs![CAMPM_CASEQ], "") & Nz(objRs![BUMOM_BUSEQ], "") & _
                                    Format(Nz(objRs![USER_CODE], 0), "000000") & Nz(objRs![USER_KANA], "")
                
                    '口座名義
                    'varCOLPT(11) = Nz(objRs![NYKOM_KOUZN], "")
                    'If Nz(objRs![NYKOM_KINYC], "") = "0010" Then    'りそな
                    '    varCOLPT(11) = "株式会社 加瀬倉庫"
                    'Else
                        varCOLPT(11) = Nz(objRs![KOUZM_KOUZN], objRs![NYKOM_KOUZN])
                        'If strCAMPC <> "KAS" Then
                        If (Nz(objRs![NYKOM_BUMOC], "") <> "1" And varCOLPT(11) = "株式会社　加瀬不動産活用") _
                        Or (Nz(objRs![NYKOM_BUMOC], "") = "1" And varCOLPT(11) = "株式会社　加瀬倉庫") Then
                            varCOLPT(11) = varCOLPT(11) & "　※収納代行会社"
                        End If
                    'End If
                Else
                    '金融機関
                    varCOLPT(9) = Nz(objRs![KOUZM_KINYN], "") & "(" & Nz(objRs![KOUZM_KINYC], "") & ") " & _
                                    Nz(objRs![KOUZM_SHITN], "") & "(" & Nz(objRs![KOUZM_SHITC], "") & ")"
                    '口座番号
                    varCOLPT(10) = Nz(objRs![KOUZM_YOKIN], "") & " " & Nz(objRs![KOUZM_KOUZB], "")
                
                    '振込人名
                    varCOLPT(12) = Nz(objRs![USER_KANA], "")
                    
                    '口座名義
                    If Nz(objRs![KOUZM_KINYC], "") = "0310" Then
                        varCOLPT(11) = "株式会社　加瀬倉庫"
                    Else
                        varCOLPT(11) = Nz(objRs![KOUZM_KOUZN], objRs![NYKOM_KOUZN])
                        'varCOLPT(11) = "株式会社　加瀬不動産活用"
                        'If strCAMPC <> "KAS" Then
                    End If
                    If (Nz(objRs![NYKOM_BUMOC], "") <> "1" And varCOLPT(11) = "株式会社　加瀬不動産活用") _
                    Or (Nz(objRs![NYKOM_BUMOC], "") = "1" And varCOLPT(11) = "株式会社　加瀬倉庫") Then
                        varCOLPT(11) = varCOLPT(11) & "　※収納代行会社"
                    End If

                End If
                '口座名義
                'varCOLPT(11) = Nz(objRs![NYKOM_KOUZN], "")
                
                '備考
                varCOLPT(13) = " "
                
                varCOLPT(14) = strTourokuNo                                     'INSERT 2023/02/28 N.IMAI
                
                '社判
                varCOLPT(15) = strCAMPC
                'テキスト1
                varCOLPT(16) = intTaxRate & "%対象"
                'テキスト2
                varCOLPT(17) = lngKinga
                'テキスト3
                varCOLPT(18) = "消費税"
                'テキスト4
                varCOLPT(19) = intTax
                               
                'お客様番号
                varCOLPT(21) = Format(objRs![USER_CODE], "000000")
                'ヤード番号
                varCOLPT(22) = Nz(objRs2![RKS170_YCODE], "")
                '部屋番号
                varCOLPT(23) = Nz(objRs2![RKS170_NO], "")
                
                
                '帳票CSV Write

                Write #iFilePT, varCOLPT(0), varCOLPT(1), varCOLPT(2), varCOLPT(3), varCOLPT(4), _
                                varCOLPT(5), varCOLPT(6), varCOLPT(7), varCOLPT(8), varCOLPT(9), _
                                varCOLPT(10), varCOLPT(11), varCOLPT(12), varCOLPT(13), varCOLPT(14), _
                                varCOLPT(15), varCOLPT(16), varCOLPT(17), varCOLPT(18), varCOLPT(19), _
                                varCOLPT(20), varCOLPT(21), varCOLPT(22), varCOLPT(23)
                objRs2.MoveNext
            Loop
        End If
        
        objRs2.Close
        Set objRs2 = Nothing
        
        objRs.MoveNext
    Loop

    Close #iFileCT
    Close #iFilePT
    
    objRs.Close
    Set objRs = Nothing
    
    If C_LOG = 1 Then
        strMsg = "CSV出力処理終了 "
        Call MSZZ003_M00(GetProgId(), "1", strMsg)
    End If
    
    FncOut_Csv2 = 0
        
ErrorHandler2:
    If Not (objRs2 Is Nothing) Then
        objRs2.Close
        Set objRs2 = Nothing
    End If
    If Err <> 0 Then
        Close #iFileCT
        Close #iFilePT
        If C_LOG = 1 Then
            strMsg = "CSV出力処理異常 Err.Number=" & Err.Number & _
                        " Err.Source=" & Err.Source & " Err.Description=" & Err.Description
            Call MSZZ003_M00(GetProgId(), "9", strMsg)
        End If
        Err.Clear
        FncOut_Csv2 = 1
    End If
    
ErrorHandler1:
    If Not (objRs Is Nothing) Then
        objRs.Close
        Set objRs = Nothing
    End If
    If Err <> 0 Then
        Close #iFileCT
        Close #iFilePT
        If C_LOG = 1 Then
            strMsg = "CSV出力処理異常 Err.Number=" & Err.Number & _
                        " Err.Source=" & Err.Source & " Err.Description=" & Err.Description
            Call MSZZ003_M00(GetProgId(), "9", strMsg)
        End If
        Err.Clear
        FncOut_Csv2 = 2
    End If

ErrorHandler:
    If Err <> 0 Then
        If C_LOG = 1 Then
            strMsg = "CSV出力処理異常 Err.Number=" & Err.Number & _
                        " Err.Source=" & Err.Source & " Err.Description=" & Err.Description
            Call MSZZ003_M00(GetProgId(), "9", strMsg)
        End If
        FncOut_Csv2 = 3
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :請求書ヘッダ情報を取得するSQL
'        MODULE_ID        :Get_strHeader
'        PARAMETER        : I/ strMonth     計上年月(yyyymm)
'        CREATE_DATE      :2018/04/30 TFC
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Get_strHeader(strMonth As String) As String
    Dim strSQL  As String
    Dim dtMonth As Date


    dtMonth = CDate(Format(strMonth & "/01", "yyyy/mm/dd"))
    
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "    請求書ヘッダーＷＫ.RKS170_KCODE"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.RKS170_KINGAKの合計"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_CODE"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_NAME"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_KANA"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_TANM"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_TAKA"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_YUBINO"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_ADR_1"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_ADR_2"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_ADR_3"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_TEL"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_KEITAI"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_FAX"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_SDAY"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_KKBN"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_SKBN"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_BIKO"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_UPDATE"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.RKS170_PAGE"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_RNAME"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_RPOST"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_RADR_1"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_RADR_2"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_RADR_3"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_KINYC"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_KINYN"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_SHITC"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_SHITN"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_YOKII"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_YOKIN"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_KOUZB "
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_KOUZN "                     'INSERT 2023/02/28 N.IMAI
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.TAX_RATE "                        'INSERT 2019/10/19 N.IMAI
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.TAX "                             'INSERT 2019/10/19 N.IMAI
    strSQL = strSQL & "  , US.USER_MAIL"
    strSQL = strSQL & "  , US.USER_RMAIL"
    
    '計上年月が初回発行月以降の場合、
    '前月顧客WKに顧客コードが存在するものはPRVFLG > 0
    'If dtMonth > C_STARTMON Then
    '    strSql = strSql & " ,"
    '    strSql = strSql & "  (SELECT count(*) "
    '    strSql = strSql & "    FROM 前月顧客ＷＫ "
    '    strSql = strSql & "    WHERE 前月顧客ＷＫ.USER_CODE = 請求書ヘッダーＷＫ.USER_CODE) AS PRVFLG"
    'Else
        strSQL = strSQL & " , 0 AS PRVFLG"
    'End If
    
    strSQL = strSQL & "  , FQS240.* "
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "    請求書ヘッダーＷＫ"
    strSQL = strSQL & "  , FQS240 "
    strSQL = strSQL & "  , dbo_USER_MAST AS US "
    strSQL = strSQL & " WHERE US.USER_CODE = 請求書ヘッダーＷＫ.USER_CODE "
    strSQL = strSQL & " ORDER BY 請求書ヘッダーＷＫ.USER_CODE "
    strSQL = strSQL & " ,請求書ヘッダーＷＫ.RKS170_PAGE "

    Get_strHeader = strSQL

End Function

''==============================================================================*
''
''        MODULE_NAME      :請求書情報を取得するSQL
''        MODULE_ID        :strOutputDetail
''        CREATE_DATE      :2018/04/30 TFC
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function strOutputDetail(lkcode As Long, ipage As Integer) As String

    Dim strSQL  As String

    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & " RKS170_DATETIME"
    strSQL = strSQL & ",RKS170_REQNO"
    strSQL = strSQL & ",RKS170_YYYYMM"
    strSQL = strSQL & ",RKS170_KJDATE"
    strSQL = strSQL & ",RKS170_SDATE"
    strSQL = strSQL & ",RKS170_NKDATE"
    strSQL = strSQL & ",RKS170_KCODE"
    strSQL = strSQL & ",RKS170_NKBN"
    strSQL = strSQL & ",RKS170_YCODE"
    strSQL = strSQL & ",RKS170_NO"
    strSQL = strSQL & ",RKS170_TTANKA"
    strSQL = strSQL & ",RKS170_TUBOSU"
    strSQL = strSQL & ",RKS170_KINGAK"
    strSQL = strSQL & ",RKS170_SYOZEI"
    strSQL = strSQL & ",RKS170_SECUKG"
    strSQL = strSQL & ",RKS170_TOTAL"
    strSQL = strSQL & ",RKS170_TEKI"
    strSQL = strSQL & ",RKS170_FLG"
    strSQL = strSQL & ",RKS170_UPDATE"
    strSQL = strSQL & ",RKS170_PAGE"
    strSQL = strSQL & ",RKS170_YARDNAME"
    strSQL = strSQL & ",RKS170_ZEIFLG"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & " 請求書ＷＫ"
    strSQL = strSQL & " WHERE RKS170_KCODE=" & lkcode
    strSQL = strSQL & " AND RKS170_PAGE=" & ipage
    strSQL = strSQL & " ORDER BY "
    strSQL = strSQL & " RKS170_KCODE"
    strSQL = strSQL & ",RKS170_PAGE"

    strOutputDetail = strSQL

End Function

'==============================================================================*
'        PROGRAM_NAME    : Get関数
'        PROGRAM_ID      : Get_Domain
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'==============================================================================*
Public Function Get_Domain() As String
    
    Dim strWhere    As String       'Where句

    'アカウント取得
    strWhere = " INTIF_PROGB = '" & GetProgId() & "' AND " & _
               " INTIF_RECFB = 'DOMAIN'"
    
    Get_Domain = Nz(DLookup("INTIF_RECDB", "INTI_FILE", strWhere), "")
    
End Function

'==============================================================================*
'        PROGRAM_NAME    : Get関数
'        PROGRAM_ID      : Get_Account
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'==============================================================================*
Public Function Get_Account() As String
    
    Dim strWhere    As String       'Where句

    'アカウント取得
    strWhere = " INTIF_PROGB = '" & GetProgId() & "' AND " & _
               " INTIF_RECFB = 'ACCOUNT'"
    
    Get_Account = Nz(DLookup("INTIF_RECDB", "INTI_FILE", strWhere), "")
    
End Function

'==============================================================================*
'        PROGRAM_NAME    : Get関数
'        PROGRAM_ID      : Get_APItoken
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'==============================================================================*
Public Function Get_APItoken() As String
    
    Dim strWhere    As String       'Where句

    'APIトークン取得
    strWhere = " INTIF_PROGB = '" & GetProgId() & "' AND " & _
               " INTIF_RECFB = 'APITOKEN'"
    
    Get_APItoken = Nz(DLookup("INTIF_RECDB", "INTI_FILE", strWhere), "")
    
End Function

'==============================================================================*
'
'        PROGRAM_NAME    : URL作成
'        PROGRAM_ID      : Fnc_MakeURL
'        PROGRAM_KBN     :
'        PARAMETER       :
'                          (i/ ) pKind  API種別
'                          (i/ ) pImpId 取込ID(Optional)
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'
'        DESCRIPTION     : URL作成
'
'==============================================================================*
Public Function Fnc_MakeURL(pKind As Integer, Optional pImpId As Integer)

    Fnc_MakeURL = "HTTPS://" & Get_Domain() & "/" & Get_Account() & "/"
    
    Select Case pKind
    
        '顧客データCSV一括取込
        Case 1:
            Fnc_MakeURL = Fnc_MakeURL & "api/v1/customers/imports"
        '顧客データ一括取込状況取得
        Case 2:
            Fnc_MakeURL = Fnc_MakeURL & "api/v1/customers/imports/" & _
                            CStr(pImpId) & "/state"
        '帳票データCSV一括取込
        Case 3:
            Fnc_MakeURL = Fnc_MakeURL & "api/v1/reports/imports"
        '帳票データ一括取込状況取得
        Case 4:
            Fnc_MakeURL = Fnc_MakeURL & "api/v1/reports/imports/" & _
                            CStr(pImpId) & "/state"
        '帳票一括取込情報削除
        Case 5:
            Fnc_MakeURL = Fnc_MakeURL & "api/v1/reports/imports/" & _
                            CStr(pImpId)
        '顧客情報取得API
        Case 6:
            Fnc_MakeURL = Fnc_MakeURL & "api/v1/customers/?keyword=H-880989"
            
        Case Else
            Fnc_MakeURL = ""
    End Select
    
End Function

'==============================================================================*
'
'        PROGRAM_NAME    : 一括取込
'        PROGRAM_ID      : Fnc_BulkCapture
'        PROGRAM_KBN     :
'        PARAMETER       :
'                          (i/ ) pKind      1:顧客データCSV一括取り込み
'                                           3:帳票データCSV一括取り込み
'                          (i/ ) pstrFile   CSVPath＋ファイル名
'                          (i/o) pCaptureID 取込ID
'        RETURN          :
'                          1以上 :正常(件数)
'                           0    :Uploadデータ無し
'                          -1    :取込失敗
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'
'        DESCRIPTION     : 各一括データ取込を行い、取り込み結果情報を取得する
'
'==============================================================================*
Public Function Fnc_BulkCapture(pKind As Integer, pstrFile As String, ByRef pCaptureID As Integer) As Integer

    Dim objHttp         As Object       'HTTP Object
    Dim JSONobj         As Object       'JSON Object
    Dim tempParmStream  As Object       'リクエストパラメータ用領域
    Dim bRet            As Boolean      '関数復帰値
    Dim sendParameter   As Variant
    Dim strMsg          As String       'ログ編集領域
    Dim vMsgPam         As Variant      '取得パラメータ
    Dim strRes          As String       'レスポンスTEXT
    Dim iFileNo         As Integer      'ファイルNo
    Dim vData           As Variant      'CSVデータ
    Dim iDataCnt        As Integer      'CSVデータ件数
    Dim strKind         As String       '帳票種別
    
    On Error GoTo ErrorHandler:
    
    If pKind = 1 Then
        strKind = "顧客"
    ElseIf pKind = 3 Then
        strKind = "帳票"
    Else
        strKind = "種別無し"
    End If
    strMsg = strKind & " アップロードファイル:" & pstrFile
        
    If C_LOG = 1 Then
        Call MSZZ003_M00(GetProgId(), "0", "一括取込開始 " & strMsg)
    End If
    Fnc_BulkCapture = -1
    sBoundy = ""
    iDataCnt = 0
    '--------------------------------------
    ' 対象有無チェック
    '--------------------------------------
    iFileNo = FreeFile()
    
    Open pstrFile For Input As iFileNo

    Do Until EOF(1)
        Line Input #iFileNo, vData
        iDataCnt = iDataCnt + 1
    Loop
    
    Close #iFileNo
            
    If iDataCnt < 2 Then
        '対象データなし
        If C_LOG = 1 Then
            strMsg = strKind & "対象データが0件の為、アップロード未実施"
            Call MSZZ003_M00(GetProgId(), "1", "一括取込終了 " & strMsg)
        End If
        Fnc_BulkCapture = 0
        Exit Function
    End If

    '--------------------------------------
    ' リクエストパラメタ用領域生成
    '--------------------------------------
    Set tempParmStream = CreateObject("ADODB.Stream")
    tempParmStream.Open
    
    '--------------------------------------
    ' リクエストパラメタ作成
    '--------------------------------------
    
    '--------------------------------------
    ' JSON作成
    '   (コメントの[*]はデフォルト値)
    '--------------------------------------
    Set JSONobj = New Dictionary
    
    If pKind = 1 Then
        '顧客CSV用
        With JSONobj
            .Add "isUpdateInfo", 1                  '利用中の上書き(0:しない[*],1:する)
    '        .Add "updateBlank", 1                  '空白で上書き(1:空白で上書き,2:元の値を残す[*])
            .Add "importProcessName", GetProgId()   'メモ
    '        .Add "skipFirst", 1                    '1行目を読みこまない(0:読み込む,1:読み込まない[*])
        End With
    Else
        '帳票CSV用
        With JSONobj
    '        .Add "reporttypeId", 1                 '帳票種別ID(1[*])
            .Add "isNewIssues", 1                   '取込方法(1:新規発行,2:差替発行) 必須
            .Add "importProcessName", GetProgId()   'メモ
    '        .Add "skipFirst", 1                    '1行目を読み込まない(0:1行目も読み込む,1:1行目を読みこまない[*])
            .Add "isImmApproval", C_APPROVE         '承認フラグ(0:承認しない[*],1:承認する)
            If C_APPROVE = 1 Then
                .Add "pdfIssueType", 0              '発行方法(0:即時発行,1:予約発行) 承認フラグ=1の場合のみ有効 かつ 必須
            End If
    '        .Add "publicationDate", _
    '                "YYYY-MM-DD HH:MM:SS"          '予約発行日
    '        .Add "isSendMail", 0                   '差替発行通知メール送信
    '        .Add "isPostReplaceBill", 0            '差替帳票の郵送依頼
    '        .Add "inputUserComment", ""            'コメント
        End With
    End If
    
    ' イミディエイトウィンドウで確認（デバック用）
    Debug.Print JsonConverter.ConvertToJson(JSONobj, Whitespace:=2)

    'データフォームのパラメータ設定
    bRet = SetNomarlParameter(tempParmStream, "json", JsonConverter.ConvertToJson(JSONobj), C_HED_CONT_APL)
    
    'ファイルのパラメタ設定
    bRet = SetFileParmater(tempParmStream, "files[0]", pstrFile, "text/csv")
    
    'フッタのパラメタ設定
    bRet = SetEndParameter(tempParmStream)
    
    '--------------------------------------
    ' リクエストパラメタ取得
    '--------------------------------------
    bRet = GetSendParameter(sendParameter, tempParmStream)
    
    '--------------------------------------
    ' リクエスト
    '--------------------------------------
    Set objHttp = CreateObject("msxml2.xmlhttp")
    
    objHttp.Open "POST", Fnc_MakeURL(pKind), False

    objHttp.setRequestHeader C_HED_CONT, C_HED_CONT_MULTI & " boundary=" + getBoundy(adBTypeContent)
    objHttp.setRequestHeader C_HED_TOK, Get_APItoken()
    objHttp.SEND sendParameter
    
    strRes = objHttp.responseText
    '-----------------------------------------
    ' リクエスト発行後のレスポンス(JSON)取得
    '-----------------------------------------
    '取得パラメータ情報作成
    vMsgPam = Split("importId,status,code,url,version,accessTime,linkuri" & _
                    ",errors(0).code,errors(0).msg,errors(0).description(0).code,errors(0).description(0).msg" & _
                    ",errors(0).description(0).name,errors(0).description(0).value", ",")
    
    Set JSONobj = JsonConverter.ParseJson(strRes)
    
    'ログ用メッセージ作成
    strMsg = Fnc_GetRes(JSONobj, vMsgPam)
    
    If objHttp.STATUS <> C_RES_OK Then
        If C_LOG = 1 Then
            Call MSZZ003_M00(GetProgId(), "9", "一括取込エラー " & strKind & " " & strMsg)
        End If
        GoTo ErrorHandler
    End If
    
    '取込ID
    pCaptureID = JSONobj("importId")
    
    If C_LOG = 1 Then
        Call MSZZ003_M00(GetProgId(), "1", "一括取込終了 " & strKind & " " & strMsg)
    End If

    Fnc_BulkCapture = iDataCnt
    
ErrorHandler:
    If Err <> 0 Then
        strMsg = "Err.Number=" & Err.Number
        strMsg = strMsg & " Err.Source=" & Err.Source
        strMsg = strMsg & " Err.Description=" & Err.Description
        
        If C_LOG = 1 Then
            Call MSZZ003_M00(GetProgId(), "9", "一括取込エラー " & strKind & " " & strMsg)
        End If
    End If
    Set objHttp = Nothing
    Set JSONobj = Nothing
    Set tempParmStream = Nothing

End Function

'==============================================================================*
'
'        PROGRAM_NAME    : 一括取込状況取得
'        PROGRAM_ID      : Fnc_BulkStatus
'        PROGRAM_KBN     :
'        PARAMETER       :
'                          (i/ ) pKind      1:顧客データCSV一括取り込み
'                                           2:帳票データCSV一括取り込み
'                          (i/ ) pCapID     一括取込ID
'        RETURN          :
'                          0 :正常
'                          1 :取込中
'                          -1:エラー
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'
'        DESCRIPTION     : 一括取込状況を取得する
'
'==============================================================================*
Public Function Fnc_BulkStatus(pKind As Integer, pCapID As Integer) As Integer

    Dim objHttp         As Object       'HTTP Object
    Dim JSONobj         As Object       'JSON Object
    Dim bRet            As Boolean      '関数復帰値
    Dim strMsg          As String       'ログ編集領域
    Dim iCaptureID      As Integer      '取込ID
    Dim strSTATUS       As String       '状態
    Dim strResult       As String       '結果
    Dim strKind         As String       '帳票種別
    
    On Error GoTo ErrorHandler:
    
    Fnc_BulkStatus = -1
    
    
    If pKind = 1 Then
        strKind = "顧客"
    Else
        strKind = "帳票"
    End If

    If C_LOG = 1 Then
        strMsg = strKind & "CSV一括取込 取込ID[" & CStr(pCapID) & "]"
        Call MSZZ003_M00(GetProgId(), "0", "一括取込状況取得開始 " & strMsg)
    End If
    '--------------------------------------
    ' JSON解析
    '--------------------------------------
    Set JSONobj = New Dictionary
    
    '--------------------------------------
    ' リクエスト
    '--------------------------------------
    Set objHttp = CreateObject("msxml2.xmlhttp")
    
    objHttp.Open "GET", Fnc_MakeURL(pKind + 1, pCapID), False
    objHttp.setRequestHeader C_HED_TOK, Get_APItoken()
    objHttp.setRequestHeader "Cache-Control", "no-cache"

    objHttp.SEND

    If objHttp.STATUS <> C_RES_OK Then
        If C_LOG = 1 Then
            Call MSZZ003_M00(GetProgId(), "9", "一括取込状況取得エラー " & strKind & _
                                " STATUS:" & objHttp.STATUS)
        End If
        GoTo ErrorHandler
    End If
    
    '-----------------------------------------
    ' リクエスト発行後のレスポンス(JSON)取得
    '-----------------------------------------
    Set JSONobj = JsonConverter.ParseJson(objHttp.responseText)
    
    '取得パラメータ情報作成
    vMsgPam = Split("status,code,url,processStatus,version" & _
                    ",accessTime", ",")
    
    'ログ用メッセージ作成
    strMsg = Fnc_GetRes(JSONobj, vMsgPam)
    
    'PROCESS STATUS
    strSTATUS = JSONobj("processStatus")
    
    Select Case (strSTATUS)
        Case C_STS_COMP
            '正常終了
            Fnc_BulkStatus = 0
            
        Case C_STS_ACTIVE
            '取込中
            Fnc_BulkStatus = 1
        
        Case Default
            'エラー
            Fnc_BulkStatus = -1
    End Select
            
    If C_LOG = 1 Then
        Call MSZZ003_M00(GetProgId(), "1", "一括取込状況取得終了 " & strKind & " " & strMsg)
    End If
    
ErrorHandler:
    If Err <> 0 Then
        strMsg = "Err.Number=" & Err.Number
        strMsg = strMsg & " Err.Source=" & Err.Source
        strMsg = strMsg & " Err.Description=" & Err.Description
        
        If C_LOG = 1 Then
            Call MSZZ003_M00(GetProgId(), "9", "一括取込状況取得エラー " & strKind & " " & strMsg)
        End If
    End If
    
    Set objHttp = Nothing
    Set JSONobj = Nothing

End Function


'==============================================================================*
'
'        PROGRAM_NAME    : レスポンス内容取得
'        PROGRAM_ID      : Fnc_GetRes
'        PROGRAM_KBN     :
'        PARAMETER       :
'                          (i/ ) pJSONobj   JSONコード
'                          (i/ ) pvParm     取得したいパラメータ
'        RETURN          : レスポンス内容
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'
'        DESCRIPTION     : JSONコード内に、取得パラメータがあれば、
'                          "パラメータ名:VALUE"の形で返却する
'
'==============================================================================*
Private Function Fnc_GetRes(pJSONobj As Object, pvParm As Variant) As String

    Dim strRes      As String

    strRes = ""
    Fnc_GetRes = ""
    
    If pvParm(0) = "" Then
        Exit Function
    End If
    
    For ilp1 = 0 To UBound(pvParm)
        '該当パラメータがあれば編集して返却
        If pJSONobj.Exists(pvParm(ilp1)) = True Then
            If strRes <> "" Then
'                strRes = strRes & " " & pvParm(ilp1) & ":" & CallByName(pJSONobj, pvParm(ilp1), VbGet)
                strRes = strRes & " " & pvParm(ilp1) & ":" & pJSONobj(pvParm(ilp1))
            Else
'                strRes = pvParm(ilp1) & ":" & CallByName(pJSONobj, pvParm(ilp1), VbGet)
                strRes = pvParm(ilp1) & ":" & pJSONobj(pvParm(ilp1))
            End If
        End If
    Next
    
    Fnc_GetRes = strRes
        
End Function



'==============================================================================*
'
'        PROGRAM_NAME    : データフォームのパラメータ設定
'        PROGRAM_ID      : SetNomarlParameter
'        PARAMETER       :
'                          (i/o) tempParamStream
'                          (i/ ) fname
'                          (i/ ) fvalue
'        RETURN          : true :正常
'                          false:異常
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'
'        DESCRIPTION     : データフォームのパラメータ設定
'
'==============================================================================*
Private Function SetNomarlParameter( _
                    ByRef tempParamStream As Object, _
                    ByVal fname As String, _
                    ByVal fvalue As String, _
                    ByVal fct As String) As Boolean

    Dim params As String

    If fvalue <> "" Then

        ChangeStreamType tempParamStream, adTypeText

        params = ""
        params = params + getBoundy(adBTypeBody)
        params = params + "Content-Disposition: form-data; name=""" + fname + """" + vbCrLf
'--- TEST
'        params = params + "Content-Type: " + fct + vbCrLf
'--- TEST END
        params = params + vbCrLf
        params = params + fvalue + vbCrLf

        tempParamStream.WriteText params

    End If

    SetNomarlParameter = True
End Function


'==============================================================================*
'
'        PROGRAM_NAME    : Boundy 情報取得
'        PROGRAM_ID      : SetNomarlParameter
'        PARAMETER       :
'                          (i/o) tempParamStream
'                          (i/ ) fname
'                          (i/ ) fvalue
'        RETURN          : true:正常
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'
'        DESCRIPTION     : Boundy 情報取得
'
'==============================================================================*
Private Function getBoundy(ByVal adType As Integer) As String


    If sBoundy = "" Then

        Dim multipartChars As String: multipartChars = "-_1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim boundary As String: boundary = "----"

        Dim i, point As Integer

        For i = 1 To 20
            Randomize
            point = Int(Len(multipartChars) * Rnd + 1)
            boundary = boundary + Mid(multipartChars, point, 1)
        Next

        sBoundy = boundary + Format(Now, "yyyymmddHHMMSS")

    End If

    Select Case adType
    Case adBTypeContent
        getBoundy = sBoundy
    Case adBTypeBody
        getBoundy = "--" + sBoundy + vbCrLf
    Case adBTypeFooter
        getBoundy = vbCrLf + "--" + sBoundy + "--" + vbCrLf
    End Select

End Function

'==============================================================================*
'
'        PROGRAM_NAME    : ファイルのパラメタ設定
'        PROGRAM_ID      : SetFileParmater
'        PARAMETER       :
'                          (i/o) tempParamStream
'                          (i/ ) fname
'                          (i/ ) fvalue
'                          (i/ ) fct
'        RETURN          : true:正常
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'
'        DESCRIPTION     : ファイルのパラメタ設定
'
'==============================================================================*
Private Function SetFileParmater( _
                            ByRef tempParamStream As Object, _
                            ByVal fname As String, _
                            ByVal fvalue As String, _
                            ByVal fct As String) As Boolean

    '-------------------------------------
    ' テキストデータ
    '-------------------------------------
    ChangeStreamType tempParamStream, adTypeText

    Dim params As String
    params = ""
    params = params + getBoundy(adBTypeBody)
    params = params + "Content-Disposition: form-data; name=""" + fname + """; filename=""" + fvalue + """" + vbCrLf
    params = params + "Content-Type: " + fct + vbCrLf
    params = params + vbCrLf

    tempParamStream.WriteText params

    '-------------------------------------
    ' バイナリデータ
    '-------------------------------------
    ChangeStreamType tempParamStream, adTypeBinary

    Dim fileStream As Object
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = adTypeBinary
    fileStream.Open
    fileStream.LoadFromFile fvalue

    tempParamStream.Write fileStream.Read()

    fileStream.Close
    Set fileStream = Nothing

    SetFileParmater = True
End Function

'==============================================================================*
'
'        PROGRAM_NAME    : フッタのパラメタ設定
'        PROGRAM_ID      : SetEndParameter
'        PARAMETER       :
'                          (i/o) tempParamStream
'        RETURN          : true:正常
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'
'        DESCRIPTION     : フッタのパラメタ設定
'
'==============================================================================*
Private Function SetEndParameter( _
                    ByRef tempParamStream As Object) As Boolean

    ChangeStreamType tempParamStream, adTypeText
    tempParamStream.WriteText getBoundy(adBTypeFooter)

    SetEndParameter = True
End Function

'==============================================================================*
'
'        PROGRAM_NAME    : 送信するパラメタを取得
'        PROGRAM_ID      : GetSendParameter
'        PARAMETER       :
'                          (i/o) parameter
'                          (i/o) stream
'        RETURN          : true:正常
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'
'        DESCRIPTION     : 送信するパラメタを取得
'
'==============================================================================*
Private Function GetSendParameter( _
                    ByRef parameter As Variant, _
                    ByRef stream As Object) As Boolean

    ChangeStreamType stream, adTypeBinary
    stream.Position = 0
    parameter = stream.Read

    stream.Close
    Set stream = Nothing

    GetSendParameter = True
End Function

'==============================================================================*
'
'        PROGRAM_NAME    : パラメタ用の領域の状態を変更する
'        PROGRAM_ID      : ChangeStreamType
'        PARAMETER       :
'                          (i/o) stream
'                          (i/ ) adType
'        RETURN          : true:正常
'        CREATE          : 2018/04/30
'        CERATER         : TFC
'
'        DESCRIPTION     : パラメタ用の領域の状態を変更する
'
'==============================================================================*
Private Function ChangeStreamType( _
                    ByRef stream As Object, _
                    ByVal adType As Integer) As Boolean
    Dim p As Long
    p = stream.Position
    stream.Position = 0
    stream.Type = adType

    If adType = adTypeText Then
        stream.Charset = "UTF-8"
    End If

    stream.Position = p

    ChangeStreamType = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :部門コード取得
'        MODULE_ID        :Get_Bumoc
'        PARAMETER        :
'        CREATE_DATE      :2018/06/30 TFC
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function Get_Bumoc() As String
    
    '部門コード取得
    Get_Bumoc = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))

End Function
    
'==============================================================================*
'
'        MODULE_NAME      :CSV出力PATH取得
'        MODULE_ID        :Get_CsvPath
'        PARAMETER        :
'        CREATE_DATE      :2018/06/30 TFC
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function Get_CsvPath(strBUMOC As String) As String
    
    Dim strWhere    As String       'Path格納

    'PATH取得
    strWhere = " INTIF_PROGB = '" & GetProgId() & "' AND " & _
               " INTIF_RECFB = 'CSVPATH_" & strBUMOC & "' "
    
    Get_CsvPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", strWhere), "")

End Function


'==============================================================================*
'
'        MODULE_NAME      :テーブル存在チェック
'        MODULE_ID        :ExistTable
'        PARAMETER        :
'        CREATE_DATE      :2018/06/30 TFC
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function ExistTable(TableName As String) As Boolean
    On Error Resume Next
    ExistTable = CurrentDb.TableDefs(TableName).NAME = TableName
End Function






