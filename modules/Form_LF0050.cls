VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LF0050"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : コンテナ管理システム
'        SUB_SYSTEM_NAME :
'
'        PROGRAM_NAME    : 契約データ出力
'        PROGRAM_ID      : LF0050
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2004/08/04
'        CERATER         : N.MIURA
'        Ver             : 0.0
'
'        UPDATE          : 2004/12/10
'        UPDATER         : N.MIURA
'        Ver             : 1.0
'                        : 名称変更
'                          代引きリストｰ>契約データ出力
'
'        UPDATE          : 2005/07/04
'        UPDATER         : N.MIURA
'        Ver             : 1.1
'                        : 顧客情報の付与
'
'        UPDATE          : 2006/04/19
'        UPDATER         : S.SHIBAZAKI
'        Ver             : 1.2
'                        : １．選択条件にレンタル用途追加
'                        : ２．その他に設置場所出力有無追加
'                        : ３．以下の出力を追加
'                        : 　　・顧客住所３
'                        : 　　・キャンペーン適用満了日
'                        : 　　・用途区分
'                        : 　　・用途名称
'                        : 　　・種別
'                        : 　　・種別名称
'                        : 　　・実帖
'                        : 　　・ヤード郵便番号
'                        : 　　・ヤード住所１～３
'                        : ４．鍵区分と保証区分の桁数を１から２に変更
'
'        UPDATE          : 2009/04/01
'        UPDATER         : kihara
'        Ver             : 1.3
'                        : 条件追加：保証会社コード
'                        : 出力項目追加：保証会社コード、保証会社名
'
'        UPDATE          : 2009/05/12
'        UPDATER         : hirano
'        Ver             : 1.4
'                        : 条件追加：保証対応不整合　保証会社有で無い場合は　仕入マスタは事業所Aで検索
'
'        UPDATE          : 2009/05/14
'        UPDATER         : hirano
'        Ver             : 1.5
'                        : 保証対応不整合　保証会社が設定されていた場合、名称セットするように修正
'
'        UPDATE          : 2009/09/08
'        UPDATER         : N.Miura
'        Ver             : 1.6
'                          1.項目追加（契約状態、予約受付番号、解約区分、解約理由区分コード）
'                          2.必須入力チェック条件変更（ISNULL =>NZ）
'                          3.日付妥当チェック追加(MSZZ012)
'                          4.CARG_FILE OPENRECORDSET DB指定変更（DB2=>KONT_DB）
'                          5.保証会社名称の取得方法の変更
'                          6.顧客範囲指定から単独していに変更
'
'        UPDATE          : 2009/10/05
'        UPDATER         : N.Miura
'        Ver             : 1.7
'                          1.出力項目追加
'                            法人個人区分、段
'                          2.出力項目削除
'                            状態区分、新契約、代引き
'                          3.抽出条件削除
'                            状態”レ”更新
'                          4.抽出条件追加
'                            所在地、段、サイズ
'                          5.画面一覧表示機能追加
'
'        UPDATE          : 2010/02/12
'        UPDATER         : M.HONDA
'        Ver             : 1.8
'                          1.出力項目追加
'                            初回請求方法、受付日を追加
'
'        UPDATE          : 2010/04/17
'        UPDATER         : M.HONDA
'        Ver             : 1.9
'                          1.出力項目追加
'                            集金区分
'
'        UPDATE          : 2011/07/07
'        UPDATER         : M.RYU
'        Ver             : 2.0
'                          1.宛先印刷ボタンを追加　⇒ LF0052　に遷移
'                          2.EXCEL出力に顧客メールを追加
'
'        UPDATE          : 2012/02/09
'        UPDATER         : M.RYU
'        Ver             : EXCEL出力に代表者名を追加
'
'        UPDATE          : 2012/04/12
'        UPDATER         : M.HONDA
'        Ver             : 2.3
'        Ver             : 1.スロープ対応
'                        : 2.受付日を検索条件に追加
'
'        UPDATE          : 2012/07/07
'        UPDATER         : M.HONDA
'        Ver             : 2.4
'        Ver             : 1.使用料合計を追加
'
'        UPDATE          : 2013/02/28
'        UPDATER         : M.HONDA
'        Ver             : 2.5
'        Ver             : 1.初回雑費・毎月雑費を追加
'
'        UPDATE          : 2013/04/10
'        UPDATER         : M.HONDA
'        Ver             : 2.6
'        Ver             : 1.キャンペーン内容を追加
'
'        UPDATE          : 2013/07/19
'        UPDATER         : M.HONDA
'        Ver             : 2.7
'        Ver             : 1.東都対応
'
'        UPDATE          : 2014/04/07
'        UPDATER         : M.HONDA
'        Ver             : 2.8
'        Ver             : 1.WEBﾊﾟﾙﾏ対応
'
'        UPDATE          : 2016/02/06
'        UPDATER         : M.HONDA
'        Ver             : 2.9
'        Ver             : 1.移動データの出力可否に対応
'
'        UPDATE          : 2016/11/17
'        UPDATER         : M.HONDA
'        Ver             : 3.0
'        Ver             : 1.ToToに府中市白糸台の報告に追加。
'
'        UPDATE          : 2017/02/08
'        UPDATER         : M.HONDA
'        Ver             : 3.1
'        Ver             : 1.生年月日・性別を追加。
'
'        UPDATE          : 2018/03/17
'        UPDATER         : M.HONDA
'        Ver             : 3.2
'        Ver             : 1.初回開始日を追加。
'
'        UPDATE          : 2018/04/26
'        UPDATER         : M.HONDA
'        Ver             : 3.3
'        Ver             : 1.初回開始日を各帳票に追加。
'
'        UPDATE          : 2019/03/01
'        UPDATER         : N.IMAI
'        Ver             : 3.4
'        Ver             : 1.初回開始日を起算日に変更
'
'        UPDATE          : 2019/11/01
'        UPDATER         : N.IMAI
'        Ver             : 3.5
'        Ver             : 1.EXCELの保存形式を変更（データ数増加のため）
'
'        UPDATE          : 2020/04/04
'        UPDATER         : N.IMAI
'        Ver             : 3.6
'        Ver             : 選択条件に受付データを追加(除く場合はYOUKT_YUKBN > 10)
'
'        UPDATE          : 2020/05/09
'        UPDATER         : N.IMAI
'        Ver             : 3.7
'        Ver             : 承認日に初期値をセットするように変更
'
'        UPDATE          : 2020/11/30
'        UPDATER         : takenouchi
'        Ver             : 3.8
'                        : ご紹介コード入力対応
'                        : EXCEL出力項目追加：紹介会社名（ユーザ入力）、紹介コード（ユーザ入力）、紹介会社名（確定版）、紹介コード（確定版）
'
'        UPDATE          : 2023/11/01
'        UPDATER         : N.IMAI
'        Ver             : 3.9
'                        : LQ0050_M10、LQ0050_M30変更
'                        : 出力項目追加：顧客カナ、緊急連絡先名、緊急連絡先電話
'
'        UPDATE          : 2023/12/01
'        UPDATER         : N.IMAI
'        Ver             : 4.0
'                        : DROP TABLE でクエリが消えるので修正
'
'        UPDATE          : 2025/05/22
'        UPDATER         : K.KINEBUCHI
'        Ver             : 4.1
'                        : 日付の抽出条件考慮漏れを対応
'                          ※例）2025/05/31 12:00:00.00　←　このようにデータが登録されている場合
'                                2025/05/01 00:00:00.00 ～ 2025/05/31 00:00:00.00 で検索しているため上記「例）」が漏れてしまう
'
'        UPDATE          : 2025/07/18
'        UPDATER         : N.IMAI
'        Ver             : 4.2
'                        : 受付日のみ入力した場合エラーが発生する問題に対応
'
'        UPDATE          : 2025/08/05
'        UPDATER         : N.IMAI
'        Ver             : 4.3
'                        : 契約日を入力した場合エラーが発生する問題に対応
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "LF0050"
'
Private USER_ID         As String
Private MODL_ID         As String

Private strDbName       As String

'
Private dbs             As Database
Private KONT_DB         As Database
Private KASE_DB         As Database '2009/05/14 INS hirano
'
Private strSQL          As String
'
Private RST_CARG        As Recordset
Private RST_Q050        As Recordset
'
Private SET_BUMOC       As String
'
'==============================================================================*
'
'        MODULE_NAME      :フォームのオープン
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Form_Open
    
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル44.BackColor = RGB(255, 0, 204)
    End Select
    
    Call MSHS010_M00(PROG_ID, "1")
    
    SET_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    
    Call LF0050_DBOPEN
    
    USER_ID = LsGetUserName()
    
    Call LF0050_CLEA0

Exit_Form_Open:
    Exit Sub

Err_Form_Open:
    MsgBox Err.Description
    Resume Exit_Form_Open
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ヤード
'        MODULE_ID        :LF0050_YARDI_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_YARDI_AfterUpdate()
       
    If LF0050_YARDI = -1 Then
       Me.LF0050_YAR1C = Me.LF0050_YAR1C.ItemData(0)
       Me.LF0050_YAR2C = Me.LF0050_YAR2C.ItemData(Me.LF0050_YAR2C.ListCount - 1)
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :顧客
'        MODULE_ID        :LF0050_KOKYF_Change
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_KOKYF_Change()
    
    If Trim(LF0050_KOKYF.Text) = "" Then
        LF0050_KOKYC = ""
        LF0050_KOKYN = ""
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :顧客名
'        MODULE_ID        :LF0050_KOKYC_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_KOKYC_AfterUpdate()
    
    With LF0050_KOKYC
        If .ListIndex >= 0 Then
            LF0050_KOKYN = .Column(1, .ListIndex)
            LF0050_KOKYF = .Column(2, .ListIndex)
            .Requery
            LF0050_KOKYI.VALUE = -1
        Else
            If Not fncGetKokyaku() Then
                LF0050_KOKYC = ""
                LF0050_KOKYN = ""
                LF0050_KOKYF.SetFocus
                .SetFocus
                MsgBox "存在しない顧客コードです。", vbOKOnly + vbInformation, Me.NAME
            End If
        End If
    End With
End Sub
'==============================================================================*
'
'        MODULE_NAME      :顧客コード
'        MODULE_ID        :LF0050_KOKYC_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_KOKYC_GotFocus()
    LF0050_KOKYC.Requery
End Sub
'==============================================================================*
'
'        MODULE_NAME      :初回開始日ＦＲＯＭ
'        MODULE_ID        :LF0050_SYO1D_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_SYO1D_AfterUpdate()
    
    Call FromAfterUpdateCopy(LF0050_SYO, LF0050_SYO1D, LF0050_SYO2D, True)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :受付日ＦＲＯＭ
'        MODULE_ID        :LF0050_KEI1D_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_UKE1D_AfterUpdate()
    
    Call FromAfterUpdateCopy(LF0050_UKE, LF0050_UKE1D, LF0050_UKE2D, True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :契約日ＦＲＯＭ
'        MODULE_ID        :LF0050_KEI1D_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_KEI1D_AfterUpdate()
    
    Call FromAfterUpdateCopy(LF0050_KEIYI, LF0050_KEI1D, LF0050_KEI2D, True)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :開始日ＦＲＯＭ
'        MODULE_ID        :LF0050_KAI1D_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_KAI1D_AfterUpdate()
    
    Call FromAfterUpdateCopy(LF0050_KAISI, LF0050_KAI1D, LF0050_KAI2D, True)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :解約日ＦＲＯＭ
'        MODULE_ID        :LF0050_KAY1D_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_KAY1D_AfterUpdate()
    
    Call FromAfterUpdateCopy(LF0050_KAIYI, LF0050_KAY1D, LF0050_KAY2D, True)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :送付日ＦＲＯＭ
'        MODULE_ID        :LF0050_SOU1D_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_SOU1D_AfterUpdate()
    
    Call FromAfterUpdateCopy(LF0050_SOUFI, LF0050_SOU1D, LF0050_SOU2D, True)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :伝票番号ＦＲＯＭ
'        MODULE_ID        :LF0050_DEN1B_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_DEN1B_AfterUpdate()
    
    Call FromAfterUpdateCopy(LF0050_DENPI, LF0050_DEN1B, LF0050_DEN2B)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :保証区分
'        MODULE_ID        :LF0050_HOSII_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_HOSII_AfterUpdate()

    If LF0050_HOSII = -1 Then
       Me.LF0050_HOS1I = Me.LF0050_HOS1I.ItemData(0)
       Me.LF0050_HOS2I = Me.LF0050_HOS2I.ItemData(Me.LF0050_HOS2I.ListCount - 1)
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :保証会社
'        MODULE_ID        :LF0050_HOSYO_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_HOSYO_AfterUpdate()

    If LF0050_HOSYO = -1 Then
       Me.LF0050_HOSYO1C = Me.LF0050_HOSYO1C.ItemData(0)
       Me.LF0050_HOSYO2C = Me.LF0050_HOSYO2C.ItemData(Me.LF0050_HOSYO2C.ListCount - 1)
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :承認日ＦＲＯＭ
'        MODULE_ID        :LF0050_HOS1D_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_HOS1D_AfterUpdate()
    
    Call FromAfterUpdateCopy(LF0050_HOSDI, LF0050_HOS1D, LF0050_HOS2D, True)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :承認番号ＦＲＯＭ
'        MODULE_ID        :LF0050_HOS1B_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_HOS1B_AfterUpdate()
    
    Call FromAfterUpdateCopy(LF0050_HOSBI, LF0050_HOS1B, LF0050_HOS2B)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :広告媒体
'        MODULE_ID        :LF0050_ADMII_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_ADMII_AfterUpdate()

    If LF0050_ADMII = -1 Then
       Me.LF0050_ADM1I = Me.LF0050_ADM1I.ItemData(0)
       Me.LF0050_ADM2I = Me.LF0050_ADM2I.ItemData(Me.LF0050_ADM2I.ListCount - 1)
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :鍵
'        MODULE_ID        :LF0050_KAGII_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_KAGII_AfterUpdate()

    If LF0050_KAGII = -1 Then
       Me.LF0050_KAG1I = Me.LF0050_KAG1I.ItemData(0)
       Me.LF0050_KAG2I = Me.LF0050_KAG2I.ItemData(Me.LF0050_KAG2I.ListCount - 1)
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :段区分
'        MODULE_ID        :LF0050_STEPI_AfterUpdate
'        CREATE_DATE      :2009/10/05 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_STEPI_AfterUpdate()

    If LF0050_STEPI = -1 Then
       Me.LF0050_STE1I = Me.LF0050_STE1I.ItemData(0)
       Me.LF0050_STE2I = Me.LF0050_STE2I.ItemData(Me.LF0050_STE2I.ListCount - 1)
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :レンタル用途
'        MODULE_ID        :LF0050_USAGE_AfterUpdate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_USAGE_AfterUpdate()

    If LF0050_USAGE = -1 Then
       Me.LF0050_USA1I = Me.LF0050_USA1I.ItemData(0)
       Me.LF0050_USA2I = Me.LF0050_USA2I.ItemData(Me.LF0050_USA2I.ListCount - 1)
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :サイズ（帖）
'        MODULE_ID        :LF0050_SIZEI_AfterUpdate
'        CREATE_DATE      :2009/10/05 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_SIZ1I_AfterUpdate()
    
    Call FromAfterUpdateCopy(LF0050_SIZEI, LF0050_SIZ1I, LF0050_SIZ2I)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :実行
'        MODULE_ID        :LF0050_EXE1B_Click
'        CREATE_DATE      :2004/06/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20041210 N.MIURA
'Private Sub LF0050_EXE1B_Click()
'    On Error GoTo Err_LF0050_EXE1B_Click
'    If LF0050_INCHK0 = True Then
'       If (MsgBox("実行しますか？", vbYesNo) = vbYes) Then
'           Call LF0050_DROPTB
'           Call LF0050_SELECT
'           doCmd.OpenReport "LR0050_M00", acPreview, "", ""
'       End If
'    End If
'Exit_LF0050_EXE1B_Click:
'    Exit Sub
'Err_LF0050_EXE1B_Click:
'    MsgBox Err.Description
'    Resume Exit_LF0050_EXE1B_Click
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :EXCEL
'        MODULE_ID        :LF0050_EXE2B_Click
'        CREATE_DATE      :2004/06/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_EXE2B_Click()
    On Error GoTo Err_LF0050_EXE2B_Click
    
    Dim logCOUNT        As Long
    
    If LF0050_INCHK0 = True Then
       If (MsgBox("実行しますか？", vbYesNo) = vbYes) Then
           
           Call LF0050_DROPTB
           Call LF0050_SELECT(logCOUNT)
           
           If logCOUNT > 0 Then
              
              MsgBox ("抽出件数　：　" & Format(logCOUNT, "000000") & "　件")
              
                '' 2013/07/19 M.HONDA
                'If LF0050_TOTO = -1 Then
                If LF0050_TOTO = -1 Or LF0050_PARU = -1 Then '2014/04/07 M.HONDA UPD
                    '東都専用形式
                    'doCmd.OutputTo acQuery, "LQ0050_M40", "MicrosoftExcel(*.xls)", "", True, ""    'DELETE 2019/11/01 N.IMAI
                    Call subOutputToExcel("LQ0050_M40")                                             'INSERT 2019/11/01 N.IMAI
                Else
                    If LF0050_KOKYG = 0 Then                                          'INSERT 20050704 N.MIURA START
                      '2006/04/19 S.SHIBAZAKI INSERT&UPDATE↓
                      If LF0050_YARDG = 0 Then
                          '顧客住所なし＆ヤード設置場所出力なし
                          'doCmd.OutputTo acQuery, "LQ0050_M00", "MicrosoftExcel(*.xls)", "", True, ""  'DELETE 2019/11/01 N.IMAI
                          Call subOutputToExcel("LQ0050_M00")                                           'INSERT 2019/11/01 N.IMAI
                      Else
                          '顧客住所なし＆ヤード設置場所出力あり
                          'doCmd.OutputTo acQuery, "LQ0050_M20", "MicrosoftExcel(*.xls)", "", True, ""  'DELETE 2019/11/01 N.IMAI
                          Call subOutputToExcel("LQ0050_M20")                                           'INSERT 2019/11/01 N.IMAI
                      End If
                      '2006/04/19 S.SHIBAZAKI INSERT&UPDATE↑
                    Else
                      '2006/04/19 S.SHIBAZAKI INSERT&UPDATE↓
                      If LF0050_YARDG = 0 Then
                          '顧客住所あり＆ヤード設置場所出力なし
                          'doCmd.OutputTo acQuery, "LQ0050_M10", "MicrosoftExcel(*.xls)", "", True, ""  'DELETE 2019/11/01 N.IMAI
                          Call LF0050_M10                                                               'INSERT 2023/12/01 N.IMAI
                          Call subOutputToExcel("LQ0050_M10")                                           'INSERT 2019/11/01 N.IMAI
                      Else
                          '顧客住所あり＆ヤード設置場所出力あり
                          'doCmd.OutputTo acQuery, "LQ0050_M30", "MicrosoftExcel(*.xls)", "", True, ""  'DELETE 2019/11/01 N.IMAI
                          Call LF0050_M30                                                               'INSERT 2023/12/01 N.IMAI
                          Call subOutputToExcel("LQ0050_M30")                                           'INSERT 2019/11/01 N.IMAI
                      End If
                      '2006/04/19 S.SHIBAZAKI INSERT&UPDATE↑
                    End If                                                            'INSERT 20050704 N.MIURA ENDED
                End If
            Else
              MsgBox ("該当データ無し")
           
           End If
       End If
    End If

Exit_LF0050_EXE2B_Click:
    Exit Sub

Err_LF0050_EXE2B_Click:
    MsgBox Err.Description
    Resume Exit_LF0050_EXE2B_Click
End Sub
'==============================================================================*
'
'        MODULE_NAME      :EXCEL
'        MODULE_ID        :LF0050_EXE3B_Click
'        CREATE_DATE      :2009/10/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_EXE3B_Click()
    On Error GoTo Err_LF0050_EXE3B_Click
    
    Dim logCOUNT        As Long
    
    If fncFormIsLoadedMsg("LF0051", Me.NAME) Then
        Exit Sub
    End If
    
    If LF0050_INCHK0 = True Then
       If (MsgBox("実行しますか？", vbYesNo) = vbYes) Then
           
           Call LF0050_DROPTB
           Call LF0050_SELECT(logCOUNT)
           
           If logCOUNT > 0 Then
              doCmd.OpenForm "LF0051"
           Else
              MsgBox ("該当データ無し")
           End If
       End If
    End If

Exit_LF0050_EXE3B_Click:
    Exit Sub

Err_LF0050_EXE3B_Click:
    MsgBox Err.Description
    Resume Exit_LF0050_EXE3B_Click
End Sub
'==============================================================================*
'
'        MODULE_NAME      :閉じる
'        MODULE_ID        :LF0050_EXITB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_EXITB_Click()
'    On Error GoTo Err_LF0050_EXITB_Click                                       'DELETE 20041208 N.MIURA
    
    doCmd.Close

'Exit_LF0050_EXITB_Click:                                                       'DELETE 20041208 N.MIURA START
'    Exit Sub
'Err_LF0050_EXITB_Click:
'    MsgBox Err.Description
'    Resume Exit_LF0050_EXITB_Click                                             'DELETE 20041208 N.MIURA ENDED
End Sub
'==============================================================================*
'
'        MODULE_NAME      :フォームクローズ
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :2004/06/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    Call LF0050_DBCLOS
    Call MSHS010_M00(PROG_ID, "2")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :画面最大表示
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    On Error GoTo Err_Form_Activate
    
    doCmd.Maximize

Exit_Form_Activate:
    Exit Sub

Err_Form_Activate:
    MsgBox Err.Description
    Resume Exit_Form_Activate
End Sub
'==============================================================================*
'
'        MODULE_NAME      :初期値
'        MODULE_ID        :LF0050_CLEA0
'        CREATE_DATE      :2004/06/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_CLEA0()
    
    With Me
    .LF0050_YARDI = -1
'    .LF0050_YAR1C = "000000"                                                   'Delete 20090908 N.Miura
'    .LF0050_YAR2C = "999999"                                                   'Delete 20090908 N.Miura
    Call LF0050_YARDI_AfterUpdate                                               'Insert 20090908 N.Miura
    .LF0050_SYOZI = 0                                                           'Insert 20091005 N.Miura
    .LF0050_KOKYI = 0
'    .LF0050_KOK1C = Null                                                       'Delete 20090908 N.Miura
'    .LF0050_KOK2C = Null                                                       'Delete 20090908 N.Miura
    
    .LF0050_KEIYI = 0
    .LF0050_KEI1D = Null
    .LF0050_KEI2D = Null
    
    '' 20120412 M.HONDA START
    .LF0050_UKE = 0
    .LF0050_UKE1D = Null
    .LF0050_UKE2D = Null
    '' 20120412 M.HONDA END
    
    '' 2018/03/17 M.HONDA START
    .LF0050_SYO = 0
    .LF0050_SYO1D = Null
    .LF0050_SYO2D = Null
    '' 2018/03/17 M.HONDA END
    
    .LF0050_KAISI = 0
    .LF0050_KAI1D = Null
    .LF0050_KAI2D = Null
    
    .LF0050_KAIYI = 0
    .LF0050_KAY1D = Null
    .LF0050_KAY2D = Null
    
    .LF0050_HOSII = 0
'    .LF0050_HOS1I = Null                                                       'Delete 20090908 N.Miura
'    .LF0050_HOS2I = Null                                                       'Delete 20090908 N.Miura
    Call LF0050_HOSII_AfterUpdate                                               'Insert 20090908 N.Miura
    
    '2009/04/01 Add <S> kihara 保証会社コード追加
    .LF0050_HOSYO = 0
'    .LF0050_HOSYO1C = Null                                                     'Delete 20090908 N.Miura
'    .LF0050_HOSYO2C = Null                                                     'Delete 20090908 N.Miura
    '2009/04/01 Add <E> kihara
    Call LF0050_HOSYO_AfterUpdate                                               'Insert 20090908 N.Miura
    
    .LF0050_HOSDI = 0
'    .LF0050_HOS1D = Null                                                       'DELETE 2020/05/09 N.IMAI
'    .LF0050_HOS2D = Null                                                       'DELETE 2020/05/09 N.IMAI
    .LF0050_HOS1D = "20000101"                                                  'INSERT 2020/05/09 N.IMAI
    .LF0050_HOS2D = Format(Now, "yyyyMMdd")                                     'INSERT 2020/05/09 N.IMAI
    
    .LF0050_HOSBI = 0
    .LF0050_HOS1B = Null
    .LF0050_HOS2B = Null
    
    .LF0050_SOUFI = 0
    .LF0050_SOU1D = Null
    .LF0050_SOU2D = Null
    
    .LF0050_DENPI = 0
    .LF0050_DEN1B = Null
    .LF0050_DEN2B = Null
    
    .LF0050_ADMII = 0
'    .LF0050_ADM1I = Null                                                       'Delete 20090908 N.Miura
'    .LF0050_ADM2I = Null                                                       'Delete 20090908 N.Miura
    Call LF0050_ADMII_AfterUpdate                                               'Insert 20090908 N.Miura
    
    .LF0050_KAGII = 0
'    .LF0050_KAG1I = Null                                                       'Delete 20090908 N.Miura
'    .LF0050_KAG2I = Null                                                       'Delete 20090908 N.Miura
    Call LF0050_KAGII_AfterUpdate                                               'Insert 20090908 N.Miura
    .LF0050_STEPI = 0                                                           'Insert 20091005 N.Miura
    .LF0050_USAGE = 0                                                           'INSERT 2006/04/19 S.SHIBAZAKI
'    .LF0050_USA1I = Null                                                       'Delete 20090908 N.Miura 'INSERT 2006/04/19 S.SHIBAZAKI
'    .LF0050_USA2I = Null                                                       'Delete 20090908 N.Miura 'INSERT 2006/04/19 S.SHIBAZAKI
    Call LF0050_USAGE_AfterUpdate                                               'Insert 20090908 N.Miura
    .LF0050_SIZEI = 0                                                           'Insert 20091005 N.Miura
    Call LF0050_SIZ1I_AfterUpdate                                               'Insert 20091005 N.Miura
    .LF0050_KOKYG = 0                                                           'INSERT 20050704 N.MIURA
    .LF0050_YARDG = 0                                                           'INSERT 2006/04/19 S.SHIBAZAKI
    End With

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :LF0050_INCHK0
'        CREATE_DATE      :2004/12/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0050_INCHK0() As Boolean
    
    LF0050_INCHK0 = True
    
'サイズ（帖）                                                                  'Insert 20091005 N.Miura Strat
    If Me.LF0050_SIZEI = -1 Then
       '必須入力
       If Nz(Me.LF0050_SIZ2I, "") = "" Then
          Me.LF0050_SIZ2I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_SIZ1I, "") = "" Then
          Me.LF0050_SIZ1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '数字
       If IsNumeric(Me.LF0050_SIZ2I) = False Then
          Me.LF0050_SIZ2I.SetFocus
          LF0050_INCHK0 = False
       End If
       If IsNumeric(Me.LF0050_SIZ1I) = False Then
          Me.LF0050_SIZ1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_SIZ1I > Me.LF0050_SIZ2I Then
          Me.LF0050_SIZ1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If                                                                     'Insert 20091005 N.Miura Ended

'レンタル用途
    If Me.LF0050_USAGE = -1 Then
       '必須入力
       If Nz(Me.LF0050_USA2I, "") = "" Then
          Me.LF0050_USA2I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_USA1I, "") = "" Then
          Me.LF0050_USA1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_USA2I) > 2 Then
          Me.LF0050_USA2I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_USA1I) > 2 Then
          Me.LF0050_USA1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '範囲
       If Val(Me.LF0050_USA1I) > Val(Me.LF0050_USA2I) Then
          Me.LF0050_USA1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If
    
'段区分                                                                        'Insert 20091005 N.Miura Strat
    If Me.LF0050_STEPI = -1 Then
       '必須入力
       If Nz(Me.LF0050_STE2I, "") = "" Then
          Me.LF0050_STE2I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_STE1I, "") = "" Then
          Me.LF0050_STE1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_STE2I) > 1 Then
          Me.LF0050_STE2I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_STE1I) > 1 Then
          Me.LF0050_STE1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_STE1I > Me.LF0050_STE2I Then
          Me.LF0050_STE1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If                                                                     'Insert 20091005 N.Miura Ended

'鍵番号　CARG_KAGIICD
    If Me.LF0050_KAGII = -1 Then
       '必須入力
       If Nz(Me.LF0050_KAG2I, "") = "" Then
          Me.LF0050_KAG2I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_KAG1I, "") = "" Then
          Me.LF0050_KAG1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_KAG2I) > 2 Then
          Me.LF0050_KAG2I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_KAG1I) > 2 Then
          Me.LF0050_KAG1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_KAG1I > Me.LF0050_KAG2I Then
          Me.LF0050_KAG1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If
    
'広告媒体　CARG_ADMEI
    If Me.LF0050_ADMII = -1 Then
       '必須入力
       If Nz(Me.LF0050_ADM2I, "") = "" Then
          Me.LF0050_ADM2I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_ADM1I, "") = "" Then
          Me.LF0050_ADM1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_ADM2I) <> 2 Then
          Me.LF0050_ADM2I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_ADM1I) <> 2 Then
          Me.LF0050_ADM1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_ADM1I > Me.LF0050_ADM2I Then
          Me.LF0050_ADM1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If
    
'承認番号　CARG_HOSYB
    If Me.LF0050_HOSBI = -1 Then
       '必須入力
       If Nz(Me.LF0050_HOS2B, "") = "" Then
          Me.LF0050_HOS2B.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_HOS1B, "") = "" Then
          Me.LF0050_HOS1B.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_HOS2B) > 20 Then
          Me.LF0050_HOS2B.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_HOS1B) > 20 Then
          Me.LF0050_HOS1B.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_HOS1B > Me.LF0050_HOS2B Then
          Me.LF0050_HOS1B.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If
    
'承認日　CARG_HOSYD
    If Me.LF0050_HOSDI = -1 Then
       '必須入力
       If Nz(Me.LF0050_HOS2D, "") = "" Then
          Me.LF0050_HOS2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_HOS1D, "") = "" Then
          Me.LF0050_HOS1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_HOS2D) <> 8 Then
          Me.LF0050_HOS2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_HOS1D) <> 8 Then
          Me.LF0050_HOS1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '日付
       If MSZZ012_M00(Me.LF0050_HOS2D) = False Then
          Me.LF0050_HOS2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If MSZZ012_M00(Me.LF0050_HOS1D) = False Then
          Me.LF0050_HOS1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("日付が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_HOS1D > Me.LF0050_HOS2D Then
          Me.LF0050_HOS1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If
    
'保証会社コード　CARG_HOSYO_CD
    If Me.LF0050_HOSYO = -1 Then
       '必須入力
       If Nz(Me.LF0050_HOSYO2C, "") = "" Then
          Me.LF0050_HOSYO2C.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_HOSYO1C, "") = "" Then
          Me.LF0050_HOSYO1C.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_HOSYO2C) > 7 Then
          Me.LF0050_HOSYO2C.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_HOSYO1C) > 7 Then
          Me.LF0050_HOSYO1C.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_HOSYO1C > Me.LF0050_HOSYO2C Then
          Me.LF0050_HOSYO1C.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If
    
'保証　CARG_HOSYICD
    If Me.LF0050_HOSII = -1 Then
       '必須入力
       If Nz(Me.LF0050_HOS2I, "") = "" Then
          Me.LF0050_HOS2I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_HOS1I, "") = "" Then
          Me.LF0050_HOS1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       If Len(Me.LF0050_HOS2I) > 2 Then
          Me.LF0050_HOS2I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_HOS1I) > 2 Then
          Me.LF0050_HOS1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_HOS1I > Me.LF0050_HOS2I Then
          Me.LF0050_HOS1I.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If

'伝票番号　CARG_DAHIB
    If Me.LF0050_DENPI = -1 Then
       '必須入力
       If Nz(Me.LF0050_DEN2B, "") = "" Then
          Me.LF0050_DEN2B.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_DEN1B, "") = "" Then
          Me.LF0050_DEN1B.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_DEN2B) > 20 Then
          Me.LF0050_DEN2B.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_DEN1B) > 20 Then
          Me.LF0050_DEN1B.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_DEN1B > Me.LF0050_DEN2B Then
          Me.LF0050_DEN1B.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If
    
'送付日　CARG_MDATE
    If Me.LF0050_SOUFI = -1 Then
       '必須入力
       If Nz(Me.LF0050_SOU2D, "") = "" Then
          Me.LF0050_SOU2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_SOU1D, "") = "" Then
          Me.LF0050_SOU1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_SOU2D) <> 8 Then
          Me.LF0050_SOU2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_SOU1D) <> 8 Then
          Me.LF0050_SOU1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '日付
       If MSZZ012_M00(Me.LF0050_SOU2D) = False Then
          Me.LF0050_SOU2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If MSZZ012_M00(Me.LF0050_SOU1D) = False Then
          Me.LF0050_SOU1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("日付が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_SOU1D > Me.LF0050_SOU2D Then
          Me.LF0050_SOU1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If
    
'解約日　CARG_KYDATE
    If Me.LF0050_KAIYI = -1 Then
       '必須入力
       If Nz(Me.LF0050_KAY2D, "") = "" Then
          Me.LF0050_KAY2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_KAY1D, "") = "" Then
          Me.LF0050_KAY1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_KAY2D) <> 8 Then
          Me.LF0050_KAY2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_KAY1D) <> 8 Then
          Me.LF0050_KAY1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '日付
       If MSZZ012_M00(Me.LF0050_KAY2D) = False Then
          Me.LF0050_KAY2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If MSZZ012_M00(Me.LF0050_KAY1D) = False Then
          Me.LF0050_KAY1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("日付が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_KAY1D > Me.LF0050_KAY2D Then
          Me.LF0050_KAY1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If
    
'開始日　CARG_STDATE
    If Me.LF0050_KAISI = -1 Then
       '必須入力
       If Nz(Me.LF0050_KAI2D, "") = "" Then
          Me.LF0050_KAI2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_KAI1D, "") = "" Then
          Me.LF0050_KAI1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_KAI2D) <> 8 Then
          Me.LF0050_KAI2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_KAI1D) <> 8 Then
          Me.LF0050_KAI1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '日付
       If MSZZ012_M00(Me.LF0050_KAI2D) = False Then
          Me.LF0050_KAI2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If MSZZ012_M00(Me.LF0050_KAI1D) = False Then
          Me.LF0050_KAI1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("日付が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_KAI1D > Me.LF0050_KAI2D Then
          Me.LF0050_KAI1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If
    
'契約日　CARG_FSDATE
    If Me.LF0050_KEIYI = -1 Then
       '必須入力
       If Nz(Me.LF0050_KEI2D, "") = "" Then
          Me.LF0050_KEI2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_KEI1D, "") = "" Then
          Me.LF0050_KEI1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_KEI2D) <> 8 Then
          Me.LF0050_KEI2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_KEI1D) <> 8 Then
          Me.LF0050_KEI1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '日付
       If MSZZ012_M00(Me.LF0050_KEI2D) = False Then
          Me.LF0050_KEI2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If MSZZ012_M00(Me.LF0050_KEI1D) = False Then
          Me.LF0050_KEI1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("日付が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_KEI1D > Me.LF0050_KEI2D Then
          Me.LF0050_KEI1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If
    
    
    '2018/03/17 M.HONDA INS
    If Me.LF0050_SYO = -1 Then
       '必須入力
       If Nz(Me.LF0050_SYO2D, "") = "" Then
          Me.LF0050_SYO2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_SYO1D, "") = "" Then
          Me.LF0050_SYO1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_SYO2D) <> 8 Then
          Me.LF0050_SYO2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_SYO1D) <> 8 Then
          Me.LF0050_SYO1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '日付
       If MSZZ012_M00(Me.LF0050_SYO2D) = False Then
          Me.LF0050_SYO2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If MSZZ012_M00(Me.LF0050_SYO1D) = False Then
          Me.LF0050_SYO1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("日付が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_SYO1D > Me.LF0050_SYO2D Then
          Me.LF0050_SYO1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If
    '2018/03/17 M.HONDA INS
    
'受付日
    If Me.LF0050_UKE = -1 Then
       '必須入力
       If Nz(Me.LF0050_UKE2D, "") = "" Then
          Me.LF0050_UKE2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Nz(Me.LF0050_UKE1D, "") = "" Then
          Me.LF0050_UKE1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_UKE2D) <> 8 Then
          Me.LF0050_UKE2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Len(Me.LF0050_UKE1D) <> 8 Then
          Me.LF0050_UKE1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
       '日付
       If MSZZ012_M00(Me.LF0050_UKE2D) = False Then
          Me.LF0050_UKE2D.SetFocus
          LF0050_INCHK0 = False
       End If
       If MSZZ012_M00(Me.LF0050_UKE1D) = False Then
          Me.LF0050_UKE1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("日付が不正です。")
          Exit Function
       End If
       '範囲
       If Me.LF0050_UKE1D > Me.LF0050_UKE2D Then
          Me.LF0050_UKE1D.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("範囲指定が不正です。")
          Exit Function
       End If
    End If

'顧客　CARG_UCODE
    If Me.LF0050_KOKYI = -1 Then
       '必須入力
       If Nz(Me.LF0050_KOKYC, "") = "" Then
          Me.LF0050_KOKYC.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("必須入力項目です。")
          Exit Function
       End If
       '桁数
       If Len(Me.LF0050_KOKYC) > 6 Then
          Me.LF0050_KOKYC.SetFocus
          LF0050_INCHK0 = False
       End If
       If Not LF0050_INCHK0 Then
          MsgBox ("入力桁数が不正です。")
          Exit Function
       End If
    End If
   
'ヤードコード
    '必須入力
    If Nz(Me.LF0050_YAR2C, "") = "" Then
       Me.LF0050_YAR2C.SetFocus
       LF0050_INCHK0 = False
    End If
    If Nz(Me.LF0050_YAR1C, "") = "" Then
       Me.LF0050_YAR1C.SetFocus
       LF0050_INCHK0 = False
    End If
    If Not LF0050_INCHK0 Then
       MsgBox ("必須入力項目です。")
       Exit Function
    End If
    '桁数
    If Len(Me.LF0050_YAR2C) <> 6 Then
       Me.LF0050_YAR2C.SetFocus
       LF0050_INCHK0 = False
    End If
    If Len(Me.LF0050_YAR1C) <> 6 Then
       Me.LF0050_YAR1C.SetFocus
       LF0050_INCHK0 = False
    End If
    If Not LF0050_INCHK0 Then
       MsgBox ("入力桁数が不正です。")
       Exit Function
    End If
    '範囲
    If Me.LF0050_YAR1C > Me.LF0050_YAR2C Then
       Me.LF0050_YAR1C.SetFocus
       LF0050_INCHK0 = False
    End If
    If Not LF0050_INCHK0 Then
       MsgBox ("範囲指定が不正です。")
       Exit Function
    End If
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :LF0050_INCHK0
'        CREATE_DATE      :2004/12/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'Delete 20090908 N.Miura
'Private Function LF0050_INCHK0() As Boolean
'    LF0050_INCHK0 = True
'    '2006/04/19 S.SHIBAZAKI INSERT↓
'    'レンタル用途　CARG_USAGE
'    If Me.LF0050_USAGE = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_USA2I) Then
'          Me.LF0050_USA2I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_USA1I) Then
'          Me.LF0050_USA1I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       If Len(Me.LF0050_USA2I) > 2 Then
'          Me.LF0050_USA2I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Len(Me.LF0050_USA1I) > 2 Then
'          Me.LF0050_USA1I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_USA1I > Me.LF0050_USA2I Then
'          Me.LF0050_USA1I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_USA1I = Null
'       Me.LF0050_USA2I = Null
'    End If
'    '2006/04/19 S.SHIBAZAKI INSERT↑
'    '鍵番号　CARG_KAGIICD
'    If Me.LF0050_KAGII = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_KAG2I) Then
'          Me.LF0050_KAG2I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_KAG1I) Then
'          Me.LF0050_KAG1I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       '2006/04/19 S.SHIBAZAKI UPDATE
'       'If Len(Me.LF0050_KAG2I) <> 1 Then
'       If Len(Me.LF0050_KAG2I) > 2 Then
'          Me.LF0050_KAG2I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       '2006/04/19 S.SHIBAZAKI UPDATE
'       'If Len(Me.LF0050_KAG1I) <> 1 Then
'       If Len(Me.LF0050_KAG1I) > 2 Then
'          Me.LF0050_KAG1I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_KAG1I > Me.LF0050_KAG2I Then
'          Me.LF0050_KAG1I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_KAG1I = Null
'       Me.LF0050_KAG2I = Null
'    End If
'    '広告媒体　CARG_ADMEI
'    If Me.LF0050_ADMII = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_ADM2I) Then
'          Me.LF0050_ADM2I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_ADM1I) Then
'          Me.LF0050_ADM1I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       If Len(Me.LF0050_ADM2I) <> 2 Then
'          Me.LF0050_ADM2I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Len(Me.LF0050_ADM1I) <> 2 Then
'          Me.LF0050_ADM1I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_ADM1I > Me.LF0050_ADM2I Then
'          Me.LF0050_ADM1I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_ADM1I = Null
'       Me.LF0050_ADM2I = Null
'    End If
'    '承認番号　CARG_HOSYB
'    If Me.LF0050_HOSBI = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_HOS2B) Then
'          Me.LF0050_HOS2B.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_HOS1B) Then
'          Me.LF0050_HOS1B.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       If Len(Me.LF0050_HOS2B) > 20 Then
'          Me.LF0050_HOS2B.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Len(Me.LF0050_HOS1B) > 20 Then
'          Me.LF0050_HOS1B.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_HOS1B > Me.LF0050_HOS2B Then
'          Me.LF0050_HOS1B.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_HOS1B = Null
'       Me.LF0050_HOS2B = Null
'    End If
'    '承認日　CARG_HOSYD
'    If Me.LF0050_HOSDI = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_HOS2D) Then
'          Me.LF0050_HOS2D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_HOS1D) Then
'          Me.LF0050_HOS1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       If Len(Me.LF0050_HOS2D) <> 8 Then
'          Me.LF0050_HOS2D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Len(Me.LF0050_HOS1D) <> 8 Then
'          Me.LF0050_HOS1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_HOS1D > Me.LF0050_HOS2D Then
'          Me.LF0050_HOS1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_HOS1D = Null
'       Me.LF0050_HOS2D = Null
'    End If
'    '保証　CARG_HOSYICD
'    If Me.LF0050_HOSII = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_HOS2I) Then
'          Me.LF0050_HOS2I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_HOS1I) Then
'          Me.LF0050_HOS1I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       '2006/04/19 S.SHIBAZAKI UPDATE
'       'If Len(Me.LF0050_HOS2I) <> 1 Then
'       If Len(Me.LF0050_HOS2I) > 2 Then
'          Me.LF0050_HOS2I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       '2006/04/19 S.SHIBAZAKI UPDATE
'       'If Len(Me.LF0050_HOS1I) <> 1 Then
'       If Len(Me.LF0050_HOS1I) > 2 Then
'          Me.LF0050_HOS1I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_HOS1I > Me.LF0050_HOS2I Then
'          Me.LF0050_HOS1I.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_HOS1I = Null
'       Me.LF0050_HOS2I = Null
'    End If
'    '保証会社コード　CARG_HOSYO_CD 2009/04/01 ADD <S> kihara
'    If Me.LF0050_HOSYO = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_HOSYO2C) Then
'          Me.LF0050_HOSYO2C.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_HOSYO1C) Then
'          Me.LF0050_HOSYO1C.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       '2006/04/19 S.SHIBAZAKI UPDATE
'       'If Len(Me.LF0050_HOS2I) <> 1 Then
'       If Len(Me.LF0050_HOSYO2C) > 7 Then
'          Me.LF0050_HOSYO2C.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       '2006/04/19 S.SHIBAZAKI UPDATE
'       'If Len(Me.LF0050_HOS1I) <> 1 Then
'       If Len(Me.LF0050_HOSYO1C) > 7 Then
'          Me.LF0050_HOSYO1C.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_HOSYO1C > Me.LF0050_HOSYO2C Then
'          Me.LF0050_HOSYO1C.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_HOSYO1C = Null
'       Me.LF0050_HOSYO2C = Null
'    End If
'    '2009/04/01 ADD <E> kihara
'    '伝票番号　CARG_DAHIB
'    If Me.LF0050_DENPI = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_DEN2B) Then
'          Me.LF0050_DEN2B.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_DEN1B) Then
'          Me.LF0050_DEN1B.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       If Len(Me.LF0050_DEN2B) > 20 Then
'          Me.LF0050_DEN2B.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Len(Me.LF0050_DEN1B) > 20 Then
'          Me.LF0050_DEN1B.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_DEN1B > Me.LF0050_DEN2B Then
'          Me.LF0050_DEN1B.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_DEN1B = Null
'       Me.LF0050_DEN2B = Null
'    End If
'    '送付日　CARG_MDATE
'    If Me.LF0050_SOUFI = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_SOU2D) Then
'          Me.LF0050_SOU2D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_SOU1D) Then
'          Me.LF0050_SOU1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       If Len(Me.LF0050_SOU2D) <> 8 Then
'          Me.LF0050_SOU2D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Len(Me.LF0050_SOU1D) <> 8 Then
'          Me.LF0050_SOU1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_SOU1D > Me.LF0050_SOU2D Then
'          Me.LF0050_SOU1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_SOU1D = Null
'       Me.LF0050_SOU2D = Null
'    End If
'    '解約日　CARG_KYDATE
'    If Me.LF0050_KAIYI = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_KAY2D) Then
'          Me.LF0050_KAY2D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_KAY1D) Then
'          Me.LF0050_KAY1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       If Len(Me.LF0050_KAY2D) <> 8 Then
'          Me.LF0050_KAY2D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Len(Me.LF0050_KAY1D) <> 8 Then
'          Me.LF0050_KAY1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_KAY1D > Me.LF0050_KAY2D Then
'          Me.LF0050_KAY1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_KAY1D = Null
'       Me.LF0050_KAY2D = Null
'    End If
'    '開始日　CARG_STDATE
'    If Me.LF0050_KAISI = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_KAI2D) Then
'          Me.LF0050_KAI2D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_KAI1D) Then
'          Me.LF0050_KAI1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       If Len(Me.LF0050_KAI2D) <> 8 Then
'          Me.LF0050_KAI2D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Len(Me.LF0050_KAI1D) <> 8 Then
'          Me.LF0050_KAI1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_KAI1D > Me.LF0050_KAI2D Then
'          Me.LF0050_KAI1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_KAI1D = Null
'       Me.LF0050_KAI2D = Null
'    End If
'    '契約日　CARG_FSDATE
'    If Me.LF0050_KEIYI = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_KEI2D) Then
'          Me.LF0050_KEI2D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_KEI1D) Then
'          Me.LF0050_KEI1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       If Len(Me.LF0050_KEI2D) <> 8 Then
'          Me.LF0050_KEI2D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Len(Me.LF0050_KEI1D) <> 8 Then
'          Me.LF0050_KEI1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_KEI1D > Me.LF0050_KEI2D Then
'          Me.LF0050_KEI1D.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_KEI1D = Null
'       Me.LF0050_KEI2D = Null
'    End If
'    '顧客　CARG_UCODE
'    If Me.LF0050_KOKYI = -1 Then
'       '必須入力
'       If IsNull(Me.LF0050_KOK2C) Then
'          Me.LF0050_KOK2C.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If IsNull(Me.LF0050_KOK1C) Then
'          Me.LF0050_KOK1C.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("必須入力項目です。")
'          Exit Function
'       End If
'       '桁数
'       If Len(Me.LF0050_KOK2C) <> 6 Then
'          Me.LF0050_KOK2C.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Len(Me.LF0050_KOK1C) <> 6 Then
'          Me.LF0050_KOK1C.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("入力桁数が不正です。")
'          Exit Function
'       End If
'       '範囲
'       If Me.LF0050_KOK1C > Me.LF0050_KOK2C Then
'          Me.LF0050_KOK1C.SetFocus
'          LF0050_INCHK0 = False
'       End If
'       If Not LF0050_INCHK0 Then
'          MsgBox ("範囲指定が不正です。")
'          Exit Function
'       End If
'    Else
'       Me.LF0050_KOK1C = Null
'       Me.LF0050_KOK2C = Null
'    End If
'    'ヤード
'    '必須入力
'    If IsNull(Me.LF0050_YAR2C) Then
'       Me.LF0050_YAR2C.SetFocus
'       LF0050_INCHK0 = False
'    End If
'    If IsNull(Me.LF0050_YAR1C) Then
'       Me.LF0050_YAR1C.SetFocus
'       LF0050_INCHK0 = False
'    End If
'    If Not LF0050_INCHK0 Then
'       MsgBox ("必須入力項目です。")
'       Exit Function
'    End If
'    '桁数
'    If Len(Me.LF0050_YAR2C) <> 6 Then
'       Me.LF0050_YAR2C.SetFocus
'       LF0050_INCHK0 = False
'    End If
'    If Len(Me.LF0050_YAR1C) <> 6 Then
'       Me.LF0050_YAR1C.SetFocus
'       LF0050_INCHK0 = False
'    End If
'    If Not LF0050_INCHK0 Then
'       MsgBox ("入力桁数が不正です。")
'       Exit Function
'    End If
'    '範囲
'    If Me.LF0050_YAR1C > Me.LF0050_YAR2C Then
'       Me.LF0050_YAR1C.SetFocus
'       LF0050_INCHK0 = False
'    End If
'    If Not LF0050_INCHK0 Then
'       MsgBox ("範囲指定が不正です。")
'       Exit Function
'    End If
'End Function
'==============================================================================*
'
'        MODULE_NAME      :fncGetKokyaku
'        機能             :顧客を検索
'        IN               :
'        OUT              :True:OK False:NG(=存在しないコード)
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetKokyaku() As Boolean
On Error GoTo ErrorHandler
    Dim dbSQLServer         As Object
    Dim rsData              As Object
    Dim blnError            As Boolean
    Dim strSQL              As String
    
    blnError = False
    fncGetKokyaku = False

    If Nz(LF0050_KOKYC, "") = "" Then
        LF0050_KOKYN = ""
        LF0050_KOKYF = ""
        fncGetKokyaku = True
        GoTo ExitHandler
    End If
    
    'SQLサーバー接続
    Set dbSQLServer = ADODB_Connection(SET_BUMOC)
    
    '顧客取得
    strSQL = strSQL & " SELECT USER_KANA, "
    strSQL = strSQL & "        USER_NAME "
    strSQL = strSQL & "   FROM USER_MAST "
    strSQL = strSQL & "  WHERE USER_CODE = " & Val(LF0050_KOKYC) & " "
    
    Set rsData = ADODB_Recordset(strSQL, dbSQLServer)

    With rsData
        If Not .EOF Then
            '顧客が存在したら名称等を表示
            LF0050_KOKYN = .Fields("USER_NAME")
            LF0050_KOKYF = .Fields("USER_KANA")
            LF0050_KOKYC.Requery
            LF0050_KOKYI.VALUE = -1
    
            fncGetKokyaku = True
        End If
    End With

    GoTo ExitHandler
    
ErrorHandler:
    blnError = True

ExitHandler:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    If blnError Then
        Call Err.Raise(Err.Number, "fncGetKokyaku" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルの削除
'        MODULE_ID        :LF0050_DROPTB
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_DROPTB()
    On Error GoTo Err_LF0050_DROPTB
    
    Dim strSQL  As String
    Dim tdf     As DAO.TableDef
    
    'strSQL = "DROP TABLE WK_LQ0050_TEMP"                                       'DELETE 2023/12/01 N.IMAI
    'Call LF0050_SQLEXEC(strSQL)                                                'DELETE 2023/12/01 N.IMAI
    'INSERT 2023/12/01 N.IMAI Start
    For Each tdf In CurrentDb.TableDefs
        If tdf.NAME = "WK_LQ0050_TEMP" Then
            doCmd.DeleteObject acTable, "WK_LQ0050_TEMP"
            Exit For
        End If
    Next
    Set tdf = Nothing
    'INSERT 2023/12/01 N.IMAI End
    
    strSQL = "CREATE TABLE WK_LQ0050_TEMP "
    strSQL = strSQL & "("
'    strSQL = strSQL & "LQ0050_YCODE  Integer,"                                'Delete 20090908 N.Miura
    strSQL = strSQL & "LQ0050_YCODE   TEXT(06),"                                'Insert 20090908 N.Miura
    strSQL = strSQL & "LQ0050_YNAME   TEXT(36),"
'    strSQL = strSQL & "LQ0050_NO     Integer,"                                'Delete 20090908 N.Miura
    strSQL = strSQL & "LQ0050_NO      TEXT(06),"                                'Insert 20090908 N.Miura
    strSQL = strSQL & "LQ0050_ACPTNO  TEXT(10),"
    
    strSQL = strSQL & "LQ0050_TANTO   TEXT(3),"                                 'Insert 20100417 M.HONDA
    strSQL = strSQL & "LQ0050_TANTN   TEXT(36),"                                'Insert 20100417 M.HONDA
    
'    strSQL = strSQL & "LQ0050_UCODE  Integer,"                                'Delete 20090908 N.Miura
    strSQL = strSQL & "LQ0050_UCODE   TEXT(06),"                                'Insert 20090908 N.Miura
    strSQL = strSQL & "LQ0050_AGRE    Integer,"                                 'Insert 20090908 N.Miura
    strSQL = strSQL & "LQ0050_AGREN   TEXT(20),"                                'Insert 20090908 N.Miura
    
    strSQL = strSQL & "LQ0050_SKBN    Integer,"                                 'Insert 20100417 M.HONDA
    strSQL = strSQL & "LQ0050_SKBNN   TEXT(20),"                                'Insert 20100417 M.HONDA
    strSQL = strSQL & "LQ0050_SEIKI    Integer,"                                 'Insert 20100417 M.HONDA
    strSQL = strSQL & "LQ0050_SEIKIN   TEXT(20),"                                'Insert 20100417 M.HONDA
    
    strSQL = strSQL & "LQ0050_SLOPE   Integer,"                                'Insert 20120412 M.HONDA
    strSQL = strSQL & "LQ0050_SLOPEN  TEXT(20),"                                'Insert 20120412 M.HONDA
    strSQL = strSQL & "LQ0050_UKDATE  TEXT(08),"                                'Insert 20120412 M.HONDA
    
    strSQL = strSQL & "LQ0050_UNAME   TEXT(36),"
    strSQL = strSQL & "LQ0050_UTANM   TEXT(36),"                                'INSERT 2012/02/09 M.RYU
    strSQL = strSQL & "LQ0050_UYUBINO TEXT(20),"                                'INSERT 20050704 N.MIURA START"
    strSQL = strSQL & "LQ0050_UADR_1  TEXT(36),"
    strSQL = strSQL & "LQ0050_UADR_2  TEXT(36),"
    strSQL = strSQL & "LQ0050_UTEL    TEXT(20),"
    strSQL = strSQL & "LQ0050_UKEITAI TEXT(20),"
    strSQL = strSQL & "LQ0050_UFAX    TEXT(20),"                                'INSERT 20050704 N.MIURA ENDED"
    
    strSQL = strSQL & "LQ0050_UMAIL   TEXT(60),"                                'INSERT 2011/07/07 M.RYU
    strSQL = strSQL & "LQ0050_FSDATE  TEXT(08),"
    strSQL = strSQL & "LQ0050_STDATE  TEXT(08),"
    strSQL = strSQL & "LQ0050_EDDATE  TEXT(08),"
    strSQL = strSQL & "LQ0050_KYDATE  TEXT(08),"
'    strSQL = strSQL & "LQ0050_CONTNA  Integer,"                                'Delete 20091005 N.Miura
'    strSQL = strSQL & "LQ0050_DOCU1   Integer,"                                'Delete 20091005 N.Miura
'    strSQL = strSQL & "LQ0050_DOCU2   Integer,"                                'Delete 20091005 N.Miura
    strSQL = strSQL & "LQ0050_RENTKG  Integer,"
    strSQL = strSQL & "LQ0050_SYOZEI  Integer,"
    strSQL = strSQL & "LQ0050_GOKEI   Integer,"                                 '2012/07/07 M.HONDA INS
    strSQL = strSQL & "LQ0050_SECUKG  Integer,"
    strSQL = strSQL & "LQ0050_FRSTKG  Integer,"
    strSQL = strSQL & "LQ0050_KAGIA   Integer,"
    strSQL = strSQL & "LQ0050_UKNO    TEXT(11),"                                'Insert 20090908 N.Miura
    strSQL = strSQL & "LQ0050_FSYOZEI Integer,"
    strSQL = strSQL & "LQ0050_MDATE   TEXT(08),"
    strSQL = strSQL & "LQ0050_DAHIB   TEXT(20),"
    strSQL = strSQL & "LQ0050_KAGIICD TEXT(02),"    'TEXT(01),"                'UPDATE 2006/04/19 S.SHIBAZAKI"
    strSQL = strSQL & "LQ0050_KAGIN   TEXT(20),"
    strSQL = strSQL & "LQ0050_KAIJB   TEXT(10),"
    strSQL = strSQL & "LQ0050_HOSYICD TEXT(02),"    'TEXT(01),"                'UPDATE 2006/04/19 S.SHIBAZAKI"
    strSQL = strSQL & "LQ0050_HOSYN   TEXT(20),"
    strSQL = strSQL & "LQ0050_HOSYO_CD TEXT(06),"    '2009/04/01 add kihara
    strSQL = strSQL & "LQ0050_HOSYO_NM TEXT(25),"    '2009/04/01 add kihara
    strSQL = strSQL & "LQ0050_HOSYD   TEXT(08),"
    strSQL = strSQL & "LQ0050_HOSYB   TEXT(20),"
    strSQL = strSQL & "LQ0050_ADMEI   TEXT(02),"
    strSQL = strSQL & "LQ0050_ADMEN   TEXT(20),"
    strSQL = strSQL & "LQ0050_KAICD   Integer,"                                 'Insert 20090908 N.Miura start
    strSQL = strSQL & "LQ0050_KAICD_NAME TEXT(20),"
    strSQL = strSQL & "LQ0050_KAIRIYUCD Integer,"
    strSQL = strSQL & "LQ0050_KAIRIYUCD_NAME TEXT(20),"                         'Insert 20090908 N.Miura Ended
    strSQL = strSQL & "LQ0050_BIKO    TEXT(60)"
    
    'INSERT 2006/04/19 S.SHIBAZAKI ↓
    strSQL = strSQL & ","
    strSQL = strSQL & "LQ0050_UADR_3        TEXT(36),"
    strSQL = strSQL & "LQ0050_CMP_EXDATE    TEXT(8),"
    strSQL = strSQL & "LQ0050_USAGECD       TEXT(2),"
    strSQL = strSQL & "LQ0050_USAGEN        TEXT(20),"
    strSQL = strSQL & "LQ0050_SYUCD         TEXT(2),"
    strSQL = strSQL & "LQ0050_SYUN          TEXT(20),"
    strSQL = strSQL & "LQ0050_SIZE          TEXT(6),"
    strSQL = strSQL & "LQ0050_YYUBINO       TEXT(20),"
    strSQL = strSQL & "LQ0050_YADR_1        TEXT(36),"
    strSQL = strSQL & "LQ0050_YADR_2        TEXT(36),"
    strSQL = strSQL & "LQ0050_YADR_3        TEXT(36)"
    'INSERT 2006/04/19 S.SHIBAZAKI ↑
    
    strSQL = strSQL & ","                                                       'Insert 20091005 N.Miura Start
    strSQL = strSQL & "LQ0050_KKBN          Integer,"
    strSQL = strSQL & "LQ0050_KKBNN         TEXT(20),"
    strSQL = strSQL & "LQ0050_STEP          Integer,"
    strSQL = strSQL & "LQ0050_STEPN         TEXT(20),"
    strSQL = strSQL & "LQ0050_BUTU          TEXT(60), "                          'Insert 20091005 N.Miura Enede
    
    '' 2010/02/12 HONDA START
    strSQL = strSQL & "LQ0050_SEIKYU        TEXT(20), "
    '' 2010/02/12 HONDA END
    
    '2020/11/30 Takenouchi INS Start
    strSQL = strSQL & "LQ0050_USER_CP_CD    TEXT(8), "
    strSQL = strSQL & "LQ0050_USER_CP_TXT   TEXT(60), "
    strSQL = strSQL & "LQ0050_CP_CD         TEXT(8), "
    strSQL = strSQL & "LQ0050_CP_TXT        TEXT(60), "
    '2020/11/30 Takenouchi INS End
    
    strSQL = strSQL & "LQ0050_EZAPPI_DAILY  Integer,"
    strSQL = strSQL & "LQ0050_EZAPPI        Integer, "
    strSQL = strSQL & "LQ0050_DCRAT_TEXT    TEXT(42), "                      ' 2013/04/10 M.HONDA INS

    strSQL = strSQL & "LQ0050_BUKEM_BUKHC   TEXT(6), "
    strSQL = strSQL & "LQ0050_SYOUI         TEXT(20), "
    
    strSQL = strSQL & "LQ0050_KIDATE        TEXT(8), "  '2018/04/26 M.HONDA INS
    
    strSQL = strSQL & "LQ0050_BIRTHDAY      TEXT(8), "
    strSQL = strSQL & "LQ0050_SEIB          TEXT(20) "
    
    'INSERT 2023/11/01 N.IMAI Start
    strSQL = strSQL & ",LQ0050_UKANA         TEXT(40) "
    strSQL = strSQL & ",LQ0050_KNAME         TEXT(50) "
    strSQL = strSQL & ",LQ0050_KTEL          TEXT(20) "
    'INSERT 2023/11/01 N.IMAI End
    
    strSQL = strSQL & ",LQ0050_NUMBER         TEXT(100) "
    strSQL = strSQL & ",LQ0050_SHASYU         TEXT(100) "
    strSQL = strSQL & ",LQ0050_HAIKIRYO       TEXT(100) "
    
    
    
    strSQL = strSQL & ");"
    
    Call LF0050_SQLEXEC(strSQL)
    
Exit_LF0050_DROPTB:
    Exit Sub

Err_LF0050_DROPTB:
    MsgBox Err.Description
    Resume Exit_LF0050_DROPTB
End Sub
'==============================================================================*
'
'        MODULE_NAME      :抽出
'        MODULE_ID        :LF0050_SELECT
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0050_SELECT(logCOUNT As Long)
    On Error GoTo Err_LF0050_SELECT
    
    Call LF0050_SEL_CARG("F")
    
    Call LF0050_INSERT("O")
    
    logCOUNT = 0
    
    Do Until RST_CARG.EOF
       
       logCOUNT = logCOUNT + 1
       
       Call LF0050_INSERT("I")
       Call LF0050_SEL_CARG("S")
    Loop
    
    RST_CARG.Close
    RST_Q050.Close

Exit_LF0050_SELECT:
    Exit Function

Err_LF0050_SELECT:
    MsgBox Err.Description
    Resume Exit_LF0050_SELECT
End Function
'==============================================================================*
'
'        MODULE_NAME      :出力
'        MODULE_ID        :LF0050_INSERT
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0050_SEL_CARG(strACCEI As String) As Boolean
    On Error GoTo Err_LF0050_SEL_CARG
    
    If strACCEI = "F" Then
       
       strDbName = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME'"), "")
       
       '2020/11/30 Takenouchi UPD Start
       '' 20100417 M.HONDA START
       '' strSQL = FunSelCarg(strDbName & ".dbo.SHIR_MAST")
       'strSQL = FunSelCarg(strDbName & ".dbo.SHIR_MAST", strDbName & ".dbo.TANT_MAST", strDbName & ".dbo.BUKE_MAST", strDbName & ".dbo.TINH_MAST")
        '' 20100417 M.HONDA END
       strSQL = FunSelCarg(strDbName & ".dbo.SHIR_MAST", strDbName & ".dbo.TANT_MAST", strDbName & ".dbo.BUKE_MAST", strDbName & ".dbo.TINH_MAST", strDbName & ".dbo.PGPA_TABL")
       '2020/11/30 Takenouchi UPD End
       
       strSQL = strSQL & LF0050_SEL_WHERE       'INSERT 2011/07/07 M.RYU

       strSQL = strSQL & ";"
       
       Set RST_CARG = KONT_DB.OpenRecordset(strSQL, dbOpenForwardOnly, dbSQLPassThrough, dbReadOnly)
       
       
           '' シングルを強制的に置き換える。
       Call MSHS010_M00(PROG_ID, "3", Replace(LF0050_SEL_WHERE, "'", ""))

       
    Else
       RST_CARG.MoveNext
    End If

    Exit Function

Err_LF0050_SEL_CARG:
    MsgBox Err.Description
End Function

'==============================================================================*
'
'        MODULE_NAME      :ＳＱＬ条件設定
'        MODULE_ID        :LF0050_SEL_WHERE
'        CREATE_DATE      :2011/07/07 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0050_SEL_WHERE() As String
    Dim strW As String
    strW = ""
    
    
    '' 2013/07/19 M.HONDA START
    If LF0050_TOTO.VALUE = -1 Then
    '' 強制的にヤードコードを指定させる。
        '2016/11/17 M.HONDA UPD
        'strW = strW & "WHERE  CARG_YCODE IN (11903,12303,22001,50705) "
        strW = strW & "WHERE  CARG_YCODE IN (11903,12303,22001,50705,22901) "
        '2016/11/17 M.HONDA UPD
    ElseIf LF0050_PARU.VALUE = -1 Then
    ' 2014/04/07 M.HONDA INS
    '' 強制的にヤードコードを指定させる。
        strW = strW & "WHERE  CARG_YCODE IN (131301,131202,131201,131203,180301,180302,140102,140103,010504) "
    Else
        'ヤード
        strW = strW & "WHERE  CARG_YCODE   Between " & Val(Me.LF0050_YAR1C) & " AND "
        strW = strW & "                            " & Val(Me.LF0050_YAR2C) & " "
    End If
    '' 2013/07/19 M.HONDA END
    
    '所在地
    If Me.LF0050_SYOZI = -1 Then
       strW = strW & "AND YARD_ADDR_1 + YARD_ADDR_2 + YARD_ADDR_3 LIKE '%" & Me.LF0050_SYOZN & "%' "
    End If
    '顧客
    If Me.LF0050_KOKYI = -1 Then
       strW = strW & "AND CARG_UCODE   =       " & Val(Me.LF0050_KOKYC) & " "
    End If
    
    '2018/03/17 M.HONDA INS
    '受付日
    If Me.LF0050_SYO = -1 Then
'       strW = strW & "AND RCPT_SYOKISAN_DATE  Between " & FunConvDate(Me.LF0050_SYO1D) & " AND "                                      'DELETE 2025/05/22 K.KINEBUCHI
       strW = strW & "AND CONVERT(DATETIME,CONVERT(VARCHAR,RCPT_SYOKISAN_DATE,112)) Between " & FunConvDate(Me.LF0050_SYO1D) & " AND " 'INSERT 2025/05/22 K.KINEBUCHI
       strW = strW & "                                " & FunConvDate(Me.LF0050_SYO2D) & " "
    End If
    
    '2018/03/17 M.HONDA INS
    
    '受付日
    If Me.LF0050_UKE = -1 Then
'       strW = strW & "AND RCPT_UKDATE  Between " & FunConvDate(Me.LF0050_UKE1D) & " AND "                                              'DELETE 2025/05/22 K.KINEBUCHI
       strW = strW & "AND CONVERT(DATETIME,CONVERT(VARCHAR,RCPT_UKDATE,112))  Between " & FunConvDate(Me.LF0050_UKE1D) & " AND "        'INSERT 2025/05/22 K.KINEBUCHI
       strW = strW & "                         " & FunConvDate(Me.LF0050_UKE2D) & " "
    End If
    
    '契約日
    If Me.LF0050_KEIYI = -1 Then
'       strW = strW & "AND CARG_FSDATE  Between " & FunConvDate(Me.LF0050_KEI1D) & " AND "                                              'DELETE 2025/05/22 K.KINEBUCHI
       strW = strW & "AND CONVERT(DATETIME,CONVERT(VARCHAR,CARG_FSDATE,112))  Between " & FunConvDate(Me.LF0050_KEI1D) & " AND "        'INSERT 2025/05/22 K.KINEBUCHI
       strW = strW & "                         " & FunConvDate(Me.LF0050_KEI2D) & " "
    End If
    '開始日
    If Me.LF0050_KAISI = -1 Then
'       strW = strW & "AND CARG_STDATE  Between " & FunConvDate(Me.LF0050_KAI1D) & " AND "                                              'DELETE 2025/05/22 K.KINEBUCHI
       strW = strW & "AND CONVERT(DATETIME,CONVERT(VARCHAR,CARG_STDATE,112))  Between " & FunConvDate(Me.LF0050_KAI1D) & " AND "        'INSERT 2025/05/22 K.KINEBUCHI
       strW = strW & "                         " & FunConvDate(Me.LF0050_KAI2D) & " "
    End If
    '解約日
    If Me.LF0050_KAIYI = -1 Then
'       strW = strW & "AND CARG_KYDATE  Between " & FunConvDate(Me.LF0050_KAY1D) & " AND "                                                              'DELETE 2025/05/22 K.KINEBUCHI
       strW = strW & "AND CONVERT(DATETIME,CONVERT(VARCHAR,ISNULL(CARG_KYDATE,'2999/01/01'),112))  Between " & FunConvDate(Me.LF0050_KAY1D) & " AND "   'INSERT 2025/05/22 K.KINEBUCHI
       strW = strW & "                         " & FunConvDate(Me.LF0050_KAY2D) & " "
    End If
    '送付日
    If Me.LF0050_HOSDI = -1 Then
'       strW = strW & "AND CARG_HOSYD   Between " & FunConvDate(Me.LF0050_HOS1D) & " AND "                                              'DELETE 2025/05/22 K.KINEBUCHI
       strW = strW & "AND CONVERT(DATETIME,CONVERT(VARCHAR,CARG_HOSYD,112))   Between " & FunConvDate(Me.LF0050_HOS1D) & " AND "        'INSERT 2025/05/22 K.KINEBUCHI
       strW = strW & "                         " & FunConvDate(Me.LF0050_HOS2D) & " "
    End If
    '伝票番号
    If Me.LF0050_HOSBI = -1 Then
       strW = strW & "AND CARG_HOSYB   Between '" & Me.LF0050_HOS1B & "' AND "
       strW = strW & "                         '" & Me.LF0050_HOS2B & "' "
    End If
    '保証区分
    If Me.LF0050_HOSII = -1 Then
       strW = strW & "AND CARG_HOSYICD Between '" & Me.LF0050_HOS1I & "' AND "
       strW = strW & "                         '" & Me.LF0050_HOS2I & "' "
    End If
    '保証会社コード
    If Me.LF0050_HOSYO = -1 Then
       strW = strW & "AND CARG_HOSYO_CD Between '" & Me.LF0050_HOSYO1C & "' AND "
       strW = strW & "                          '" & Me.LF0050_HOSYO2C & "' "
    End If
    '承認日
    If Me.LF0050_SOUFI = -1 Then
'       strW = strW & "AND CARG_MDATE   Between " & FunConvDate(Me.LF0050_SOU1D) & " AND "                                              'DELETE 2025/05/22 K.KINEBUCHI
       strW = strW & "AND CONVERT(DATETIME,CONVERT(VARCHAR,CARG_MDATE,112))   Between " & FunConvDate(Me.LF0050_SOU1D) & " AND "        'INSERT 2025/05/22 K.KINEBUCHI
       strW = strW & "                         " & FunConvDate(Me.LF0050_SOU2D) & " "
    End If
    '承認番号
    If Me.LF0050_DENPI = -1 Then
       strW = strW & "AND CARG_DAHIB   Between '" & Me.LF0050_DEN1B & "' AND "
       strW = strW & "                         '" & Me.LF0050_DEN2B & "' "
    End If
    '広告媒体
    If Me.LF0050_ADMII = -1 Then
       strW = strW & "AND CARG_ADMEI   Between '" & Me.LF0050_ADM1I & "' AND "
       strW = strW & "                         '" & Me.LF0050_ADM2I & "' "
    End If
    '鍵
    If Me.LF0050_KAGII = -1 Then
       strW = strW & "AND CARG_KAGIICD Between '" & Me.LF0050_KAG1I & "' AND "
       strW = strW & "                         '" & Me.LF0050_KAG2I & "' "
    End If
     
    '段区分
    If Me.LF0050_STEPI = -1 Then
       strW = strW & "AND CNTA_STEP   Between " & Val(Me.LF0050_STE1I) & " AND "
       strW = strW & "                        " & Val(Me.LF0050_STE2I) & " "
    End If
     
    'レンタル用途
    If Me.LF0050_USAGE = -1 Then
       strW = strW & "AND CNTA_USAGE  Between " & Val(Me.LF0050_USA1I) & " AND "
       strW = strW & "                        " & Val(Me.LF0050_USA2I) & " "
    End If
     
    'サイズ（帖）
    If Me.LF0050_SIZEI = -1 Then
       strW = strW & "AND CNTA_SIZE   Between " & Val(Me.LF0050_SIZ1I) & " AND "
       strW = strW & "                        " & Val(Me.LF0050_SIZ2I) & " "
    End If

    'INSERT 2020/04/06 N.IMAI Start
    '受付データ
    Select Case Me.LF0050_UKETI
        Case 0
            
        Case 1
            strW = strW & "AND YOUKT_YUKBN > 10 "
    End Select
    'INSERT 2020/04/06 N.IMAI End
    
    '解約明細
    Select Case Me.LF0050_KAIGI
        Case 1
            If Me.LF0050_UKE = -1 Then
'                strW = strW & "AND ISNULL(CARG_KYDATE,'2999/01/01')  >= " & FunConvDate(LF0050_UKE2D)                                  'DELETE 2025/05/22 K.KINEBUCHI
'                strW = strW & "AND CONVERT(DATETIME,CONVERT(ISNULL(CARG_KYDATE,'2999/01/01'),112))  >= " & FunConvDate(LF0050_UKE2D)   'INSERT 2025/05/22 K.KINEBUCHI
                strW = strW & "AND CONVERT(DATETIME,CONVERT(VARCHAR,ISNULL(CARG_KYDATE,'2999/01/01'),112))  >= " & FunConvDate(LF0050_UKE2D)    'INSERT 2025/07/18 N.IMAI UPDATE 2025/08/05
            '2018/03/17 M.HONDA INS
            ElseIf LF0050_SYO = -1 Then
'                strW = strW & "AND ISNULL(CARG_KYDATE,'2999/01/01')  >= " & FunConvDate(LF0050_SYO2D)                                  'DELETE 2025/05/22 K.KINEBUCHI
'                strW = strW & "AND CONVERT(DATETIME,CONVERT(ISNULL(CARG_KYDATE,'2999/01/01'),112))  >= " & FunConvDate(LF0050_SYO2D)   'INSERT 2025/05/22 K.KINEBUCHI
                strW = strW & "AND CONVERT(DATETIME,CONVERT(VARCHAR,ISNULL(CARG_KYDATE,'2999/01/01'),112))  >= " & FunConvDate(LF0050_SYO2D)    'INSERT 2025/07/18 N.IMAI UPDATE 2025/08/05
            '2018/03/17 M.HONDA INS
            ElseIf LF0050_KEIYI = -1 Then
'                strW = strW & "AND ISNULL(CARG_KYDATE,'2999/01/01')  >= " & FunConvDate(LF0050_KEI2D)                                  'DELETE 2025/05/22 K.KINEBUCHI
'                strW = strW & "AND CONVERT(DATETIME,CONVERT(ISNULL(CARG_KYDATE,'2999/01/01'),112))  >= " & FunConvDate(LF0050_KEI2D)   'INSERT 2025/05/22 K.KINEBUCHI
                strW = strW & "AND CONVERT(DATETIME,CONVERT(VARCHAR,ISNULL(CARG_KYDATE,'2999/01/01'),112))  >= " & FunConvDate(LF0050_KEI2D)    'INSERT 2025/07/18 N.IMAI UPDATE 2025/08/05
            ElseIf LF0050_KAISI = -1 Then
'                strW = strW & "AND ISNULL(CARG_KYDATE,'2999/01/01')  >= " & FunConvDate(LF0050_KAI2D)                                  'DELETE 2025/05/22 K.KINEBUCHI
'                strW = strW & "AND CONVERT(DATETIME,CONVERT(ISNULL(CARG_KYDATE,'2999/01/01'),112))  >= " & FunConvDate(LF0050_KAI2D)   'INSERT 2025/05/22 K.KINEBUCHI
                strW = strW & "AND CONVERT(DATETIME,CONVERT(VARCHAR,ISNULL(CARG_KYDATE,'2999/01/01'),112))  >= " & FunConvDate(LF0050_KAI2D)    'INSERT 2025/07/18 N.IMAI UPDATE 2025/08/05
            ElseIf LF0050_KAIYI = -1 Then
                'strW = strW & "AND ISNULL(CARG_KYDATE,'2999/01/01')  >= " & FunConvDate(LF0050_KAY2D)
            Else
               strW = strW & "AND CARG_KYDATE  IS NULL "
            End If
        Case 2
            strW = strW & "AND CARG_KYDATE  IS NOT NULL "
        Case Else
    End Select

    '2016/02/06 M.HONDA INS
    Select Case Me.LF0050_IDO
    Case 1
        If (Me.LF0050_UKE = 0 Or Me.LF0050_KEIYI = 0 Or Me.LF0050_KAISI = 0 Or Me.LF0050_SYO = 0) And LF0050_KAIYI = -1 Then
            strW = strW & "            AND CARG_STDATE <= isnull(CARG_KYDATE,'2999/01/01') "
            strW = strW & "            AND  not EXISTS( "
            strW = strW & "                           SELECT * "
            strW = strW & "                           FROM   YOUK_TRAN , CARG_FILE B "
            strW = strW & "                           WHERE YOUKT_MOTO_ACPTNO = a.CARG_ACPTNO "
            strW = strW & "                           AND    ISNULL(YOUKT_MOVEKBN, 0) <>  0"
            strW = strW & "                           AND    YOUKT_YUKBN = '20' "
            strW = strW & "                           AND    YOUKT_UKNO = B.CARG_UKNO "
            strW = strW & "                           AND    B.CARG_STDATE <= isnull(B.CARG_KYDATE,'2999/01/01') "
            strW = strW & "                            ) "
        Else
            strW = strW & "        AND CARG_STDATE <= isnull(CARG_KYDATE,'2999/01/01') "
            strW = strW & "        AND NOT EXISTS( "
            strW = strW & "                SELECT * FROM YOUK_TRAN "
            strW = strW & "                WHERE YOUKT_UKNO = CARG_UKNO "
            strW = strW & "                AND   ISNULL(YOUKT_MOTO_ACPTNO,'') <> '' "
            strW = strW & "                AND   YOUKT_MOVEKBN IN ('1','2','4','5','6','7','10') "
            strW = strW & "                ) "
        End If
    Case 2
        If (Me.LF0050_UKE = 0 Or Me.LF0050_KEIYI = 0 Or Me.LF0050_KAISI = 0 Or Me.LF0050_SYO = 0) And LF0050_KAIYI = -1 Then
            strW = strW & "            AND CARG_STDATE <= isnull(CARG_KYDATE,'2999/01/01') "
            strW = strW & "            AND    EXISTS( "
            strW = strW & "                           SELECT * "
            strW = strW & "                           FROM   YOUK_TRAN , CARG_FILE B "
            strW = strW & "                           WHERE YOUKT_MOTO_ACPTNO = a.CARG_ACPTNO "
            strW = strW & "                           AND    ISNULL(YOUKT_MOVEKBN, 0) <>  0"
            strW = strW & "                           AND    YOUKT_YUKBN = '20' "
            strW = strW & "                           AND    YOUKT_UKNO = B.CARG_UKNO "
            strW = strW & "                           AND    B.CARG_STDATE <= isnull(B.CARG_KYDATE,'2999/01/01') "
            strW = strW & "                            ) "
        Else
            strW = strW & "        AND CARG_STDATE <= isnull(CARG_KYDATE,'2999/01/01') "
            strW = strW & "        AND     EXISTS( "
            strW = strW & "                SELECT * FROM YOUK_TRAN "
            strW = strW & "                WHERE  YOUKT_UKNO = CARG_UKNO "
            strW = strW & "                AND    ISNULL(YOUKT_MOTO_ACPTNO,'') <> '' "
            strW = strW & "                AND    YOUKT_MOVEKBN IN ('1','2','4','5','6','7','10') "
            strW = strW & "                ) "
        End If
    Case Else
    End Select
    '2016/02/06 M.HONDA INS

    LF0050_SEL_WHERE = strW
    
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :CONVERT DATETIME
'        MODULE_ID        :FunConvDate
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function FunConvDate(strDate As String) As String

    Dim strConvDate     As String

    strConvDate = ""
    strConvDate = strConvDate & "CONVERT(DATETIME,'"
    strConvDate = strConvDate & Mid(strDate, 1, 4) & "/ "
    strConvDate = strConvDate & Mid(strDate, 5, 2) & "/"
    strConvDate = strConvDate & Mid(strDate, 7, 2) & "')"

    FunConvDate = strConvDate

End Function
'==============================================================================*
'
'       MODULE_NAME     : FROMをTOにコピーする機能
'       MODULE_ID       : FromAfterUpdateCopy
'       CREATE_DATE     : 2009/09/08            N.Miura
'       PARAM           : ctrlChk               チェックボックス
'                       : ctrlFrom              FROM
'                       : ctrlTo                TO
'                       : [bIsDate]             日付型
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function FromAfterUpdateCopy(ctrlChk As CheckBox, ctrlFrom As Control, ctrlTo As Control, Optional ByVal bIsDate As Boolean = False) As Boolean
    Dim bCopy               As Boolean
    
    If TypeName(ctrlFrom) = "Combobox" Then
        bCopy = (ctrlFrom.ListIndex >= 0)
    Else
        bCopy = Nz(ctrlFrom.VALUE) <> ""
    End If
    If bCopy Then
        ctrlChk.VALUE = -1
        If bIsDate Then
            If MSZZ012_M00(ctrlFrom.VALUE) Then
                If Right(ctrlFrom.VALUE, 2) = "01" Then
                    ctrlTo.VALUE = Format(DateSerial(Val(Left(ctrlFrom.VALUE, 4)), Val(Mid(ctrlFrom.VALUE, 5, 2)) + 1, 0), "yyyymmdd")
                Else
                    ctrlTo.VALUE = ctrlFrom.VALUE
                End If
            Else
                ctrlTo.VALUE = ""
            End If
        Else
            ctrlTo.VALUE = ctrlFrom.VALUE
        End If
    End If
    FromAfterUpdateCopy = bCopy
End Function
'==============================================================================*
'
'        MODULE_NAME      :出力
'        MODULE_ID        :LF0050_INSERT
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'Delete 20090908 N.Miura
'Private Function LF0050_SEL_CARG(strACCEI As String) As Boolean
'    On Error GoTo Err_LF0050_SEL_CARG
'    If strACCEI = "F" Then
'       strSQL = ""
'       strSQL = strSQL & "SELECT "
'       strSQL = strSQL & "CARG_YCODE, "
'       strSQL = strSQL & "YARD_NAME, "
'       strSQL = strSQL & "CARG_NO, "
'       strSQL = strSQL & "CARG_ACPTNO, "
'       strSQL = strSQL & "CARG_UCODE, "
'       strSQL = strSQL & "USER_NAME, "
'       strSQL = strSQL & "USER_YUBINO, "                                        'INSERT 20050704 N.MIURA START
'       strSQL = strSQL & "USER_ADR_1, "
'       strSQL = strSQL & "USER_ADR_2, "
'       strSQL = strSQL & "USER_TEL, "
'       strSQL = strSQL & "USER_KEITAI, "
'       strSQL = strSQL & "USER_FAX, "                                           'INSERT 20050704 N.MIURA ENDED
'       strSQL = strSQL & "CARG_FSDATE, "
'       strSQL = strSQL & "CARG_STDATE, "
'       strSQL = strSQL & "CARG_EDDATE, "
'       strSQL = strSQL & "CARG_KYDATE, "
'       strSQL = strSQL & "CARG_CONTNA, "
'       strSQL = strSQL & "CARG_DOCU1, "
'       strSQL = strSQL & "CARG_DOCU2, "
'       strSQL = strSQL & "CARG_RENTKG, "
'       strSQL = strSQL & "CARG_SYOZEI, "
'       strSQL = strSQL & "CARG_SECUKG, "
'       strSQL = strSQL & "CARG_FRSTKG, "
'       strSQL = strSQL & "CARG_KAGIA, "
'       strSQL = strSQL & "CARG_FSYOZEI, "
'       strSQL = strSQL & "CARG_MDATE, "
'       strSQL = strSQL & "CARG_DAHIB, "
'       strSQL = strSQL & "CARG_KAGIICD, "
'       strSQL = strSQL & "NAME_C060.KAGIN, "
'       strSQL = strSQL & "CARG_KAIJB, "
'       strSQL = strSQL & "CARG_HOSYICD, "
'       strSQL = strSQL & "NAME_C200.HOSYN, "
'        'INSERT 2009/04/01 kihara ↓
'       strSQL = strSQL & "CARG_HOSYO_CD, "
'       '2009/05/14 MOD<S> hirano
'       'strSQL = strSQL & "SHIR_MAST.SHIRM_SHIRN, "
'       strSQL = strSQL & "NULL AS SHIRM_SHIRN, "
'       '2009/05/14 MOD<E> hirano
'        'INSERT 2009/04/01 kihara ↑
'       strSQL = strSQL & "CARG_HOSYD, "
'       strSQL = strSQL & "CARG_HOSYB, "
'       strSQL = strSQL & "CARG_ADMEI, "
'       strSQL = strSQL & "NAME_C070.ADMEN, "
'       strSQL = strSQL & "CARG_BIKO "
'        'INSERT 2006/04/19 S.SHIBAZAKI ↓
'        strSQL = strSQL & ", "
'        strSQL = strSQL & "USER_ADR_3, "
'        strSQL = strSQL & "CARG_CMP_EXDATE, "
'        strSQL = strSQL & "CNTA_USAGE, "
'        strSQL = strSQL & "USAGENAME.NAME_NAME AS USAGE_NAME, "
'        strSQL = strSQL & "CNTA_SYU, "
'        strSQL = strSQL & "SYUNAME.NAME_NAME AS SYU_NAME, "
'        strSQL = strSQL & "CNTA_SIZE, "
'        strSQL = strSQL & "YARD_YUBINO, "
'        strSQL = strSQL & "YARD_ADDR_1, "
'        strSQL = strSQL & "YARD_ADDR_2, "
'        strSQL = strSQL & "YARD_ADDR_3 "
'        'INSERT 2006/04/19 S.SHIBAZAKI ↑
'       '2009/04/01 MOD <S> kihara
'       'strSQL = strSQL & "FROM (((((((dbo_CARG_FILE "      '((((dbo_CARG_FILE "       'UPDATE 2006/04/19 S.SHIBAZAKI
'       '2009/05/14 MOD<S> hirano 元に戻す
'       'strSQL = strSQL & "FROM ((((((((dbo_CARG_FILE "
'       strSQL = strSQL & "FROM  (((((((dbo_CARG_FILE "
'       '2009/05/14 MOD<E> hirano
'       '2009/04/01 MOD <E> kihara
'       strSQL = strSQL & "INNER JOIN dbo_YARD_MAST ON dbo_CARG_FILE.CARG_YCODE   = dbo_YARD_MAST.YARD_CODE) "
'       strSQL = strSQL & "INNER JOIN dbo_USER_MAST ON dbo_CARG_FILE.CARG_UCODE   = dbo_USER_MAST.USER_CODE) "
'       strSQL = strSQL & "INNER JOIN NAME_C070     ON dbo_CARG_FILE.CARG_ADMEI   = NAME_C070.ADMEI        ) "
'       strSQL = strSQL & "INNER JOIN NAME_C200     ON dbo_CARG_FILE.CARG_HOSYICD = NAME_C200.HOSYI        ) "
'        'INSERT 2006/04/19 S.SHIBAZAKI ↓
'        strSQL = strSQL & "INNER JOIN dbo_CNTA_MAST "
'        strSQL = strSQL & "   ON dbo_CARG_FILE.CARG_YCODE = dbo_CNTA_MAST.CNTA_CODE "
'        strSQL = strSQL & "  AND dbo_CARG_FILE.CARG_NO    = dbo_CNTA_MAST.CNTA_NO ) "
'        strSQL = strSQL & "INNER JOIN dbo_NAME_MAST USAGENAME "
'        strSQL = strSQL & "  ON USAGENAME.NAME_CODE = dbo_CNTA_MAST.CNTA_USAGE ) "
'        strSQL = strSQL & "INNER JOIN dbo_NAME_MAST SYUNAME "
'        strSQL = strSQL & "  ON SYUNAME.NAME_CODE = dbo_CNTA_MAST.CNTA_SYU ) "
'        'INSERT 2006/04/19 S.SHIBAZAKI ↑
'       strSQL = strSQL & "INNER JOIN NAME_C060     ON dbo_CARG_FILE.CARG_KAGIICD = NAME_C060.KAGII          "
'        'INSERT 2009/04/01 kihara ↓
'        '2009/05/14 DEL<S> hirano
'        'strSQL = strSQL & ") LEFT JOIN SHIR_MAST "
'        'strSQL = strSQL & "   ON dbo_CARG_FILE.CARG_HOSYO_CD = SHIR_MAST.SHIRM_SHIRC "
'        '2009/05/14 DEL<E> hirano
'        'INSERT 2009/04/01 kihara ↑
'       'ヤード
'       strSQL = strSQL & "WHERE  CARG_YCODE   Between " & Val(Me.LF0050_YAR1C) & " AND "
'       strSQL = strSQL & "                            " & Val(Me.LF0050_YAR2C) & " "
'        'INSERT 2006/04/19 S.SHIBAZAKI ↓
'        strSQL = strSQL & "AND USAGENAME.NAME_ID = '086' "
'        strSQL = strSQL & "AND SYUNAME.NAME_ID = '040' "
'        'INSERT 2006/04/19 S.SHIBAZAKI ↑
'        'INSERT 2009/04/01 kihara ↓
'        '2009/05/12 MOD<S> hirano
'        'strSQL = strSQL & "AND ((SHIRM_BUMOC  = '" & SET_BUMOC & "' AND CARG_HOSYICD = '1') OR CARG_HOSYICD <> '1') "
'        '2009/05/14 DEL<S> hirano
'        'strSQL = strSQL & "AND ((SHIRM_BUMOC  = '" & SET_BUMOC & "' AND CARG_HOSYICD = '1') OR (CARG_HOSYICD <> '1' AND SHIRM_BUMOC  = 'A') ) "
'        '2009/05/14 DEL<E> hirano
'        '2009/05/12 MOD<E> hirano
'        'INSERT 2009/04/01 kihara ↑
'       '顧客
'       If Me.LF0050_KOKYI = 0 Then
'       Else
'          strSQL = strSQL & "AND CARG_UCODE   Between " & Val(Me.LF0050_KOK1C) & " AND "
'          strSQL = strSQL & "                         " & Val(Me.LF0050_KOK2C) & " "
'       End If
'       '契約日
'       If Me.LF0050_KEIYI = 0 Then
'       Else
'          strSQL = strSQL & "AND CARG_FSDATE  Between CDATE(#" & LF0050_chgDATE(Me.LF0050_KEI1D) & "#) AND "
'          strSQL = strSQL & "                         CDATE(#" & LF0050_chgDATE(Me.LF0050_KEI2D) & "#) "
'       End If
'       '開始日
'       If Me.LF0050_KAISI = 0 Then
'       Else
'          strSQL = strSQL & "AND CARG_STDATE  Between CDATE(#" & LF0050_chgDATE(Me.LF0050_KAI1D) & "#) AND "
'          strSQL = strSQL & "                         CDATE(#" & LF0050_chgDATE(Me.LF0050_KAI2D) & "#) "
'       End If
'       '解約日
'       If Me.LF0050_KAIYI = 0 Then
'       Else
'          strSQL = strSQL & "AND CARG_KYDATE  Between CDATE(#" & LF0050_chgDATE(Me.LF0050_KAY1D) & "#) AND "
'          strSQL = strSQL & "                         CDATE(#" & LF0050_chgDATE(Me.LF0050_KAY2D) & "#) "
'       End If
'       '送付日
'       If Me.LF0050_HOSDI = 0 Then
'       Else
'          strSQL = strSQL & "AND CARG_HOSYD   Between CDATE(#" & LF0050_chgDATE(Me.LF0050_HOS1D) & "#) AND "
'          strSQL = strSQL & "                         CDATE(#" & LF0050_chgDATE(Me.LF0050_HOS2D) & "#) "
'       End If
'       '伝票番号
'       If Me.LF0050_HOSBI = 0 Then
'       Else
'          strSQL = strSQL & "AND CARG_HOSYB   Between '" & Me.LF0050_HOS1B & "' AND "
'          strSQL = strSQL & "                         '" & Me.LF0050_HOS2B & "' "
'       End If
'       '保証区分
'       If Me.LF0050_HOSII = 0 Then
'       Else
'          strSQL = strSQL & "AND CARG_HOSYICD Between '" & Me.LF0050_HOS1I & "' AND "
'          strSQL = strSQL & "                         '" & Me.LF0050_HOS2I & "' "
'       End If
'       '保証会社コード 2009/04/01 ADD <S> kihara
'       If Me.LF0050_HOSYO = 0 Then
'       Else
'          strSQL = strSQL & "AND CARG_HOSYO_CD Between '" & Me.LF0050_HOSYO1C & "' AND "
'          strSQL = strSQL & "                         '" & Me.LF0050_HOSYO2C & "' "
'       End If
'       '2009/04/01 ADD <E> kihara
'       '承認日
'       If Me.LF0050_SOUFI = 0 Then
'       Else
'          strSQL = strSQL & "AND CARG_MDATE   Between CDATE(#" & LF0050_chgDATE(Me.LF0050_SOU1D) & "#) AND "
'          strSQL = strSQL & "                         CDATE(#" & LF0050_chgDATE(Me.LF0050_SOU2D) & "#) "
'       End If
'       '承認番号
'       If Me.LF0050_DENPI = 0 Then
'       Else
'          strSQL = strSQL & "AND CARG_DAHIB   Between '" & Me.LF0050_DEN1B & "' AND "
'          strSQL = strSQL & "                         '" & Me.LF0050_DEN2B & "' "
'       End If
'       '広告媒体
'       If Me.LF0050_ADMII = 0 Then
'       Else
'          strSQL = strSQL & "AND CARG_ADMEI   Between '" & Me.LF0050_ADM1I & "' AND "
'          strSQL = strSQL & "                         '" & Me.LF0050_ADM2I & "' "
'       End If
'       '鍵
'       If Me.LF0050_KAGII = 0 Then
'       Else
'          strSQL = strSQL & "AND CARG_KAGIICD Between '" & Me.LF0050_KAG1I & "' AND "
'          strSQL = strSQL & "                         '" & Me.LF0050_KAG2I & "' "
'       End If
'        'INSERT 2006/04/19 S.SHIBAZAKI ↓
'        If Me.LF0050_USAGE = 0 Then
'        Else
'           strSQL = strSQL & "AND CNTA_USAGE Between " & Me.LF0050_USA1I & " AND "
'           strSQL = strSQL & "                       " & Me.LF0050_USA2I & " "
'        End If
'        'INSERT 2006/04/19 S.SHIBAZAKI ↑
'       '解約明細
'       Select Case Me.LF0050_KAIGI
'       Case 1
'          strSQL = strSQL & "AND CARG_KYDATE  IS NULL "
'       Case 2
'          strSQL = strSQL & "AND CARG_KYDATE  IS NOT NULL "
'       Case Else
'       End Select
'       '状態”レ”更新
'       Select Case Me.LF0050_KAIGF
'       Case 1
'          strSQL = strSQL & "AND CARG_CONTNA  <> -1 "
'       Case 2
'          strSQL = strSQL & "AND CARG_CONTNA  =  -1 "
'       Case Else
'       End Select
'       strSQL = strSQL & ";"
'       Set RST_CARG = DBS.OpenRecordset(strSQL, dbOpenForwardOnly, dbReadOnly)
'    Else
'       RST_CARG.MoveNext
'    End If
'    Exit Function
'Err_LF0050_SEL_CARG:
'    MsgBox Err.Description
'End Function
'==============================================================================*
'
'        MODULE_NAME      :出力
'        MODULE_ID        :LF0050_INSERT
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0050_INSERT(strACCEI As String) As Boolean
    
    If strACCEI = "O" Then
       Set RST_Q050 = dbs.OpenRecordset("WK_LQ0050_TEMP", dbOpenDynaset, dbAppendOnly)
    Else
       With RST_Q050
       .AddNew
       .Fields("LQ0050_YCODE") = RST_CARG.Fields("CARG_YCODE")
       .Fields("LQ0050_YNAME") = RST_CARG.Fields("YARD_NAME")
       .Fields("LQ0050_NO") = RST_CARG.Fields("CARG_NO")
       .Fields("LQ0050_ACPTNO") = RST_CARG.Fields("CARG_ACPTNO")
       .Fields("LQ0050_UCODE") = RST_CARG.Fields("CARG_UCODE")
       .Fields("LQ0050_UNAME") = RST_CARG.Fields("USER_NAME")
       .Fields("LQ0050_UTANM") = RST_CARG.Fields("USER_TANM")       'INSERT 2012/02/09 M.RYU
       .Fields("LQ0050_AGRE") = RST_CARG.Fields("CARG_AGRE")
       .Fields("LQ0050_AGREN") = RST_CARG.Fields("AGREN")
       
       '' 20100417 M.HODNA START
       .Fields("LQ0050_TANTO") = RST_CARG.Fields("CARG_UKTANTO")
       .Fields("LQ0050_TANTN") = RST_CARG.Fields("TANTM_TANTN")
       .Fields("LQ0050_SKBN") = RST_CARG.Fields("USER_SKBN")
       .Fields("LQ0050_SKBNN") = RST_CARG.Fields("SKBN")
       .Fields("LQ0050_SEIKI") = RST_CARG.Fields("CARG_SEIKI")
       .Fields("LQ0050_SEIKIN") = RST_CARG.Fields("SEIKI")
       '' 20100417 M.HODNA END
       
       .Fields("LQ0050_UYUBINO") = RST_CARG.Fields("USER_YUBINO")
       .Fields("LQ0050_UADR_1") = RST_CARG.Fields("USER_ADR_1")
       .Fields("LQ0050_UADR_2") = RST_CARG.Fields("USER_ADR_2")
       .Fields("LQ0050_UTEL") = RST_CARG.Fields("USER_TEL")
       .Fields("LQ0050_UKEITAI") = RST_CARG.Fields("USER_KEITAI")
       .Fields("LQ0050_UFAX") = RST_CARG.Fields("USER_FAX")
       .Fields("LQ0050_KKBN") = RST_CARG.Fields("USER_KKBN")                   'Insert 20091005 N.Miura
       .Fields("LQ0050_KKBNN") = RST_CARG.Fields("KKBNN")                      'Insert 20091005 N.Miura
       
       .Fields("LQ0050_KIDATE") = RST_CARG.Fields("RCPT_SYOKISAN_DATE")        '2018/04/26 M.HONDA INS
       
       .Fields("LQ0050_FSDATE") = RST_CARG.Fields("CARG_FSDATE")
       .Fields("LQ0050_STDATE") = RST_CARG.Fields("CARG_STDATE")
       .Fields("LQ0050_EDDATE") = RST_CARG.Fields("CARG_EDDATE")
       .Fields("LQ0050_KYDATE") = RST_CARG.Fields("CARG_KYDATE")
'       .Fields("LQ0050_CONTNA") = RST_CARG.Fields("CARG_CONTNA")              'Delete 20091005 N.Miura
'       .Fields("LQ0050_DOCU1") = RST_CARG.Fields("CARG_DOCU1")                'Delete 20091005 N.Miura
'       .Fields("LQ0050_DOCU2") = RST_CARG.Fields("CARG_DOCU2")                'Delete 20091005 N.Miura
       .Fields("LQ0050_RENTKG") = RST_CARG.Fields("CARG_RENTKG")
       .Fields("LQ0050_SYOZEI") = RST_CARG.Fields("CARG_SYOZEI")
       .Fields("LQ0050_GOKEI") = RST_CARG.Fields("CARG_GOKEI")                ' 2012/07/07 M.HONDA INS
       .Fields("LQ0050_SECUKG") = RST_CARG.Fields("CARG_SECUKG")
       .Fields("LQ0050_FRSTKG") = RST_CARG.Fields("CARG_FRSTKG")
       .Fields("LQ0050_KAGIA") = RST_CARG.Fields("CARG_KAGIA")
       .Fields("LQ0050_UKNO") = RST_CARG.Fields("CARG_UKNO")
       .Fields("LQ0050_FSYOZEI") = RST_CARG.Fields("CARG_FSYOZEI")
       .Fields("LQ0050_MDATE") = RST_CARG.Fields("CARG_MDATE")
       .Fields("LQ0050_DAHIB") = RST_CARG.Fields("CARG_DAHIB")
       .Fields("LQ0050_KAGIICD") = RST_CARG.Fields("CARG_KAGIICD")
       .Fields("LQ0050_KAGIN") = RST_CARG.Fields("KAGIN")
       .Fields("LQ0050_KAIJB") = RST_CARG.Fields("CARG_KAIJB")
       .Fields("LQ0050_HOSYICD") = RST_CARG.Fields("CARG_HOSYICD")
       .Fields("LQ0050_HOSYN") = RST_CARG.Fields("HOSYN")
       .Fields("LQ0050_HOSYO_CD") = RST_CARG.Fields("CARG_HOSYO_CD")
       .Fields("LQ0050_HOSYO_NM") = RST_CARG.Fields("SHIRM_SHIRN")
       .Fields("LQ0050_HOSYD") = RST_CARG.Fields("CARG_HOSYD")
       .Fields("LQ0050_HOSYB") = RST_CARG.Fields("CARG_HOSYB")
       .Fields("LQ0050_ADMEI") = RST_CARG.Fields("CARG_ADMEI")
       .Fields("LQ0050_ADMEN") = RST_CARG.Fields("ADMEN")
       .Fields("LQ0050_KAICD") = RST_CARG.Fields("CARG_KAICD")
       .Fields("LQ0050_KAICD_NAME") = RST_CARG.Fields("KAICD_NAME")
       .Fields("LQ0050_KAIRIYUCD") = RST_CARG.Fields("CARG_KAIRIYUCD")
       .Fields("LQ0050_KAIRIYUCD_NAME") = RST_CARG.Fields("KAIRIYUCD_NAME")
       .Fields("LQ0050_BIKO") = RST_CARG.Fields("CARG_BIKO")
       .Fields("LQ0050_UADR_3") = RST_CARG.Fields("USER_ADR_3")
       .Fields("LQ0050_CMP_EXDATE") = RST_CARG.Fields("CARG_CMP_EXDATE")
       .Fields("LQ0050_STEP") = RST_CARG.Fields("CNTA_STEP")                   'Insert 20091005 N.Miura
       .Fields("LQ0050_STEPN") = RST_CARG.Fields("CNTA_STEPN")                 'Insert 20091005 N.Miura
       .Fields("LQ0050_USAGECD") = RST_CARG.Fields("CNTA_USAGE")
       .Fields("LQ0050_USAGEN") = RST_CARG.Fields("USAGE_NAME")
       .Fields("LQ0050_SYUCD") = RST_CARG.Fields("CNTA_SYU")
       .Fields("LQ0050_SYUN") = RST_CARG.Fields("SYU_NAME")
       .Fields("LQ0050_SIZE") = Format(RST_CARG.Fields("CNTA_SIZE"), "0.00")
       .Fields("LQ0050_YYUBINO") = RST_CARG.Fields("YARD_YUBINO")
       .Fields("LQ0050_YADR_1") = RST_CARG.Fields("YARD_ADDR_1")
       .Fields("LQ0050_YADR_2") = RST_CARG.Fields("YARD_ADDR_2")
       .Fields("LQ0050_YADR_3") = RST_CARG.Fields("YARD_ADDR_3")
       '' 20120704 M.HONDA START
       ''.Fields("LQ0050_BUTU") = RST_CARG.Fields("RCPT_BUTU")                   'Insert 20091005 N.Miura
       .Fields("LQ0050_BUTU") = ""
       '' 20120704 M.HONDA END
       '' 2010/02/12 M.HONDA START
       .Fields("LQ0050_UKDATE") = RST_CARG.Fields("RCPT_UKDATE")
       .Fields("LQ0050_SEIKYU") = RST_CARG.Fields("SEIKYU")
       '' 2010/02/12 M.HONDA END
       .Fields("LQ0050_UMAIL") = RST_CARG.Fields("USER_MAIL")                   'INSERT 2011/07/07 M.RYU
       '' 2012/04/12 M.HONDA START
       .Fields("LQ0050_SLOPE") = Nz(RST_CARG.Fields("RCPT_SLOPE"), 0)
       .Fields("LQ0050_SLOPEN") = RST_CARG.Fields("RCPT_SLOPEN")
       '' 2012/04/12 M.HONDA END
       
       '' 2013/02/28 M.HONDA START
       .Fields("LQ0050_EZAPPI") = Nz(RST_CARG.Fields("RCPT_EZAPPI"), 0)
       .Fields("LQ0050_EZAPPI_DAILY") = Nz(RST_CARG.Fields("RCPT_EZAPPI_DAILY"), 0)
       '' 2013/02/28 M.HONDA END
       
       .Fields("LQ0050_DCRAT_TEXT") = RST_CARG.Fields("DCRAT_TEXT") '2013/04/10 M.HONDA
       
       .Fields("LQ0050_BUKEM_BUKHC") = RST_CARG.Fields("BUKEM_BUKHC")
       .Fields("LQ0050_SYOUI") = RST_CARG.Fields("SYOUI")
       
       .Fields("LQ0050_BIRTHDAY") = RST_CARG.Fields("USER_BIRTHDAY")
       .Fields("LQ0050_SEIB") = RST_CARG.Fields("SEIB")
       
       '2020/11/30 Takenouchi INS Start
       .Fields("LQ0050_USER_CP_TXT") = RST_CARG.Fields("RCPT_USER_COUPON_TXT")
       .Fields("LQ0050_USER_CP_CD") = RST_CARG.Fields("RCPT_USER_COUPON_CD")
       .Fields("LQ0050_CP_TXT") = RST_CARG.Fields("COUPON_TXT")
       .Fields("LQ0050_CP_CD") = RST_CARG.Fields("RCPT_COUPON_CD")
       '2020/11/30 Takenouchi INS End
       
       'INSERT 2023/11/01 N.IMAI Start
       .Fields("LQ0050_UKANA") = RST_CARG.Fields("USER_KANA")
       .Fields("LQ0050_KNAME") = RST_CARG.Fields("RCPT_KREN_NAME")
       .Fields("LQ0050_KTEL") = RST_CARG.Fields("RCPT_KREN_TEL")
       'INSERT 2023/11/01 N.IMAI End
       
       .Fields("LQ0050_NUMBER") = RST_CARG.Fields("RCPT_BIKE_NUMBER")
       .Fields("LQ0050_SHASYU") = RST_CARG.Fields("RCPT_BIKE_SHASYU")
       .Fields("LQ0050_HAIKIRYO") = RST_CARG.Fields("RCPT_BIKE_HAIKIRYO")
       
       
       
       
       
       .UPDATE
       End With
    End If

End Function
'==============================================================================*
'
'        MODULE_NAME      :出力
'        MODULE_ID        :LF0050_INSERT
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function LF0050_INSERT(strACCEI As String) As Boolean
'    If strACCEI = "O" Then
'       Set RST_Q050 = DBS.OpenRecordset("WK_LQ0050_TEMP", dbOpenDynaset, dbAppendOnly)
'    Else
'       With RST_Q050
'       .AddNew
'       .Fields("LQ0050_YCODE") = RST_CARG.Fields("CARG_YCODE")
'       .Fields("LQ0050_YNAME") = RST_CARG.Fields("YARD_NAME")
'       .Fields("LQ0050_NO") = RST_CARG.Fields("CARG_NO")
'       .Fields("LQ0050_ACPTNO") = RST_CARG.Fields("CARG_ACPTNO")
'       .Fields("LQ0050_UCODE") = RST_CARG.Fields("CARG_UCODE")
'       .Fields("LQ0050_UNAME") = RST_CARG.Fields("USER_NAME")
'       .Fields("LQ0050_UYUBINO") = RST_CARG.Fields("USER_YUBINO")               'INSERT 20050704 N.MIURA START
'       .Fields("LQ0050_UADR_1") = RST_CARG.Fields("USER_ADR_1")
'       .Fields("LQ0050_UADR_2") = RST_CARG.Fields("USER_ADR_2")
'       .Fields("LQ0050_UTEL") = RST_CARG.Fields("USER_TEL")
'       .Fields("LQ0050_UKEITAI") = RST_CARG.Fields("USER_KEITAI")
'       .Fields("LQ0050_UFAX") = RST_CARG.Fields("USER_FAX")                     'INSERT 20050704 N.MIURA ENDED
'       .Fields("LQ0050_FSDATE") = Format(RST_CARG.Fields("CARG_FSDATE"), "YYYYMMDD")
'       .Fields("LQ0050_STDATE") = Format(RST_CARG.Fields("CARG_STDATE"), "YYYYMMDD")
'       .Fields("LQ0050_EDDATE") = Format(RST_CARG.Fields("CARG_EDDATE"), "YYYYMMDD")
'       .Fields("LQ0050_KYDATE") = Format(RST_CARG.Fields("CARG_KYDATE"), "YYYYMMDD")
'       .Fields("LQ0050_CONTNA") = RST_CARG.Fields("CARG_CONTNA")
'       .Fields("LQ0050_DOCU1") = RST_CARG.Fields("CARG_DOCU1")
'       .Fields("LQ0050_DOCU2") = RST_CARG.Fields("CARG_DOCU2")
'       .Fields("LQ0050_RENTKG") = RST_CARG.Fields("CARG_RENTKG")
'       .Fields("LQ0050_SYOZEI") = RST_CARG.Fields("CARG_SYOZEI")
'       .Fields("LQ0050_SECUKG") = RST_CARG.Fields("CARG_SECUKG")
'       .Fields("LQ0050_FRSTKG") = RST_CARG.Fields("CARG_FRSTKG")
'       .Fields("LQ0050_KAGIA") = RST_CARG.Fields("CARG_KAGIA")
'       .Fields("LQ0050_FSYOZEI") = RST_CARG.Fields("CARG_FSYOZEI")
'       .Fields("LQ0050_MDATE") = Format(RST_CARG.Fields("CARG_MDATE"), "YYYYMMDD")
'       .Fields("LQ0050_DAHIB") = RST_CARG.Fields("CARG_DAHIB")
'       .Fields("LQ0050_KAGIICD") = RST_CARG.Fields("CARG_KAGIICD")
'       .Fields("LQ0050_KAGIN") = RST_CARG.Fields("KAGIN")
'       .Fields("LQ0050_KAIJB") = RST_CARG.Fields("CARG_KAIJB")
'       .Fields("LQ0050_HOSYICD") = RST_CARG.Fields("CARG_HOSYICD")
'       .Fields("LQ0050_HOSYN") = RST_CARG.Fields("HOSYN")
'        'INSERT 2009/04/01 kihara ↓
'        .Fields("LQ0050_HOSYO_CD") = RST_CARG.Fields("CARG_HOSYO_CD")
'        '2009/05/14 MOD <S> hirano
'        '.Fields("LQ0050_HOSYO_NM") = RST_CARG.Fields("SHIRM_SHIRN")
'        If RST_CARG.Fields("CARG_HOSYO_CD") <> "" Then
'            .Fields("LQ0050_HOSYO_NM") = LF0050_SEL_SHIR(RST_CARG.Fields("CARG_HOSYO_CD"))
'        End If
'        '2009/05/14 MOD <E> hirano
'        'INSERT 2009/04/01 kihara ↑
'       .Fields("LQ0050_HOSYD") = Format(RST_CARG.Fields("CARG_HOSYD"), "YYYYMMDD")
'       .Fields("LQ0050_HOSYB") = RST_CARG.Fields("CARG_HOSYB")
'       .Fields("LQ0050_ADMEI") = RST_CARG.Fields("CARG_ADMEI")
'       .Fields("LQ0050_ADMEN") = RST_CARG.Fields("ADMEN")
'       .Fields("LQ0050_BIKO") = RST_CARG.Fields("CARG_BIKO")
'        'INSERT 2006/04/19 S.SHIBAZAKI ↓
'        .Fields("LQ0050_UADR_3") = RST_CARG.Fields("USER_ADR_3")
'        .Fields("LQ0050_CMP_EXDATE") = Format(RST_CARG.Fields("CARG_CMP_EXDATE"), "yyyymmdd")
'        .Fields("LQ0050_USAGECD") = RST_CARG.Fields("CNTA_USAGE")
'        .Fields("LQ0050_USAGEN") = RST_CARG.Fields("USAGE_NAME")
'        .Fields("LQ0050_SYUCD") = RST_CARG.Fields("CNTA_SYU")
'        .Fields("LQ0050_SYUN") = RST_CARG.Fields("SYU_NAME")
'        .Fields("LQ0050_SIZE") = Format(RST_CARG.Fields("CNTA_SIZE"), "0.00")
'        .Fields("LQ0050_YYUBINO") = RST_CARG.Fields("YARD_YUBINO")
'        .Fields("LQ0050_YADR_1") = RST_CARG.Fields("YARD_ADDR_1")
'        .Fields("LQ0050_YADR_2") = RST_CARG.Fields("YARD_ADDR_2")
'        .Fields("LQ0050_YADR_3") = RST_CARG.Fields("YARD_ADDR_3")
'        'INSERT 2006/04/19 S.SHIBAZAKI ↑
'       .UPDATE
'       End With
'    End If
'End Function
'==============================================================================*
'
'        MODULE_NAME      :SQLの実行
'        MODULE_ID        :LF0050_SQLEXEC
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function LF0050_SQLEXEC(strSQL As String) As Boolean
    On Error GoTo Err_LF0050_SQLEXEC

    dbs.Execute strSQL, dbFailOnError

    LF0050_SQLEXEC = True
    Exit Function

Err_LF0050_SQLEXEC:
    MsgBox Err.Description
    LF0050_SQLEXEC = False
End Function
'==============================================================================*
'
'        MODULE_NAME      :抽出クエリー
'        MODULE_ID        :FunSelCarG
'        CREATE_DATE      :2009/09/08 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2020/11/30 Takenouchi UPD Start
'Function FunSelCarg(strSHIR_MAST As String) As String
'Function FunSelCarg(strSHIR_MAST As String, strTANT_MAST As String, strBUKE_MAST As String, strTINH_MAST As String) As String
Function FunSelCarg(strSHIR_MAST As String, strTANT_MAST As String, strBUKE_MAST As String, strTINH_MAST As String, strPGPA_TABL As String) As String
'2020/11/30 Takenouchi UPD End
    
    Dim strSQL          As String

    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "  REPLACE(STR(CARG_YCODE,6),' ',0) AS CARG_YCODE,"
    strSQL = strSQL & "  YARD_NAME,"
    strSQL = strSQL & "  REPLACE(STR(CARG_NO,6),' ',0)    AS CARG_NO,"
    strSQL = strSQL & "  CARG_ACPTNO,"
    
'' 2010/04/17 M.HONDA START
    strSQL = strSQL & "  CARG_UKTANTO, "
    strSQL = strSQL & "  TAN.TANTM_TANTN, "
'' 2010/04/17 M.HONDA END
    'strSql = strSql & "  CONVERT(CHAR(8),RCPT_SYOKISAN_DATE,112) AS RCPT_SYOKISAN_DATE, "  'DELETE 2019/03/01 N.IMAI
    strSQL = strSQL & "  CONVERT(CHAR(8),RCPT_KISAN_DATE,112) AS RCPT_SYOKISAN_DATE, "      'INSERT 2019/03/01 N.IMAI
    
    strSQL = strSQL & "  REPLACE(STR(CARG_UCODE,6),' ',0) AS CARG_UCODE,"
    strSQL = strSQL & "  USER_NAME,"
    strSQL = strSQL & "  USER_TANM,"                                            'INSERT 2012/02/09 M.RYU
    strSQL = strSQL & "  CARG_AGRE,"
    strSQL = strSQL & "  NAME_C088.NAME_NAME AS AGREN,"
    
    '' 2010/04/17 M.HONDA START
    strSQL = strSQL & "  USER_SKBN,"
    strSQL = strSQL & "  NAME_C030.NAME_NAME AS SKBN,"
    strSQL = strSQL & "  CARG_SEIKI,"
    strSQL = strSQL & "  ISNULL(NAME_C201.NAME_NAME,'')  AS SEIKI,"
    '' 2010/04/17 M.HONDA END
    
    strSQL = strSQL & "  USER_YUBINO,   "
    strSQL = strSQL & "  USER_ADR_1,    "
    strSQL = strSQL & "  USER_ADR_2,    "
    strSQL = strSQL & "  USER_ADR_3,    "
    strSQL = strSQL & "  USER_TEL,      "
    strSQL = strSQL & "  USER_KEITAI,   "
    strSQL = strSQL & "  USER_FAX,      "
    strSQL = strSQL & "  USER_MAIL,     "                                      'INSERT 2011/07/07 M.RYU
    strSQL = strSQL & "  USER_KKBN,     "                                      'Insert 20091005 N.Miura
    strSQL = strSQL & "  NAME_C020.NAME_NAME                  AS KKBNN,"       'Insert 20091005 N.Miura
    
    '' 2010/02/12 M.HONDA START
    strSQL = strSQL & "  NAME_C014.NAME_NAME AS SEIKYU,"
    strSQL = strSQL & "  CONVERT(CHAR(8),RCPT_UKDATE,112) AS RCPT_UKDATE,"
    '' 2010/02/12 M.HONDA END
    
    strSQL = strSQL & "  USER_BIRTHDAY, "                                      'INS 2017/02/08 M.HONDA
    strSQL = strSQL & "  USER_SEIBETU, "                                       'INS 2017/02/08 M.HONDA
    strSQL = strSQL & "  NAME_C302.NAME_NAME                  AS SEIB, "
    
    strSQL = strSQL & "  CONVERT(CHAR(8),CARG_FSDATE,112)     AS CARG_FSDATE,"
    strSQL = strSQL & "  CONVERT(CHAR(8),CARG_STDATE,112)     AS CARG_STDATE,"
    strSQL = strSQL & "  CONVERT(CHAR(8),CARG_EDDATE,112)     AS CARG_EDDATE,"
    strSQL = strSQL & "  CONVERT(CHAR(8),CARG_KYDATE,112)     AS CARG_KYDATE,"
'    strSQL = strSQL & "  CARG_CONTNA,"                                        'Delete 20091005 N.Miura
'    strSQL = strSQL & "  CARG_DOCU1,"                                         'Delete 20091005 N.Miura
'    strSQL = strSQL & "  CARG_DOCU2,"                                         'Delete 20091005 N.Miura
    strSQL = strSQL & "  CARG_RENTKG,"
    strSQL = strSQL & "  CARG_SYOZEI,"
    
    strSQL = strSQL & "  CARG_RENTKG + CARG_SYOZEI            AS CARG_GOKEI ,"    '' 2012/07/07 M.HONDA INS
    
    strSQL = strSQL & "  CARG_SECUKG,"
    strSQL = strSQL & "  CARG_FRSTKG,"
    strSQL = strSQL & "  CARG_KAGIA,"
    strSQL = strSQL & "  CARG_UKNO,"
    strSQL = strSQL & "  CARG_FSYOZEI,"
    strSQL = strSQL & "  CONVERT(CHAR(8),CARG_MDATE,112)      AS CARG_MDATE,"
    strSQL = strSQL & "  CARG_DAHIB,"
    strSQL = strSQL & "  CARG_KAGIICD,"
    strSQL = strSQL & "  NAME_C060.NAME_NAME                  AS KAGIN,"
    strSQL = strSQL & "  CARG_KAIJB,"
    strSQL = strSQL & "  CARG_HOSYICD,"
    strSQL = strSQL & "  NAME_C200.NAME_NAME                  AS HOSYN,"
    strSQL = strSQL & "  CARG_HOSYO_CD,"
    strSQL = strSQL & "  SHIRM_SHIRN                          AS SHIRM_SHIRN,"
    strSQL = strSQL & "  CONVERT(CHAR(8),CARG_HOSYD,112)      AS CARG_HOSYD,"
    strSQL = strSQL & "  CARG_HOSYB,"
    strSQL = strSQL & "  CARG_ADMEI,"
    strSQL = strSQL & "  NAME_C070.NAME_NAME                  AS ADMEN,"
    strSQL = strSQL & "  CARG_KAICD,"
    strSQL = strSQL & "  NAME_C089.NAME_NAME                  AS KAICD_NAME,"
    strSQL = strSQL & "  CARG_KAIRIYUCD,"
    strSQL = strSQL & "  NAME_C092.NAME_NAME                  AS KAIRIYUCD_NAME,"
    strSQL = strSQL & "  CARG_BIKO,"
    strSQL = strSQL & "  CONVERT(CHAR(8),CARG_CMP_EXDATE,112) AS CARG_CMP_EXDATE,"
    strSQL = strSQL & "  CNTA_STEP,"                                           'Insert 20091005 N.Miura
    strSQL = strSQL & "  NAME_C090.NAME_NAME                  AS CNTA_STEPN,"  'Insert 20091005 N.Miura
    strSQL = strSQL & "  CNTA_USAGE,"
    strSQL = strSQL & "  RCPT_SLOPE,"                                           'Insert 20120214 M.HONDA
    strSQL = strSQL & "  NAME_C294.NAME_NAME                  AS RCPT_SLOPEN,"  'Insert 20120214 M.HONDA
    strSQL = strSQL & "  RCPT_EZAPPI, "                                         'Insert 20130228 M.HONDA
    strSQL = strSQL & "  RCPT_EZAPPI_DAILY, "                                   'Insert 20130228 M.HONDA
    strSQL = strSQL & "  NAME_C086.NAME_NAME                  AS USAGE_NAME,"
    strSQL = strSQL & "  CNTA_SYU,"
    strSQL = strSQL & "  NAME_C040.NAME_NAME                  AS SYU_NAME,"
    strSQL = strSQL & "  CNTA_SIZE,"
    strSQL = strSQL & "  YARD_YUBINO,"
    strSQL = strSQL & "  YARD_ADDR_1,"
    strSQL = strSQL & "  YARD_ADDR_2,"
    strSQL = strSQL & "  YARD_ADDR_3,"                                         'Update 20091005 N.Miura
    strSQL = strSQL & "  RCPT_BUTU, "                                           'Insert 20091005 N.Miura
    
    strSQL = strSQL & "  RCPT_BIKE_NUMBER, "
    strSQL = strSQL & "  RCPT_BIKE_SHASYU, "
    strSQL = strSQL & "  RCPT_BIKE_HAIKIRYO, "
    
    '2020/11/30 Takenouchi INS Start
    strSQL = strSQL & "  RCPT_USER_COUPON_CD, "
    strSQL = strSQL & "  RCPT_USER_COUPON_TXT, "
    strSQL = strSQL & "  RCPT_COUPON_CD, "
    strSQL = strSQL & "  PGPAT_BIKON                          AS COUPON_TXT, "
    '2020/11/30 Takenouchi INS End
    
    strSQL = strSQL & "  DCRAT_TEXT, "                                           'Insert 2013/04/10 M.HONDA
    strSQL = strSQL & "  BUKEM_BUKHC, "
    strSQL = strSQL & "  BUKEM_SYOUI, "
    strSQL = strSQL & "  NAME_C210.NAME_NAME                  AS SYOUI "
    
    'INSERT 2023/11/01 N.IMAI Start
    strSQL = strSQL & ",USER_KANA"
    strSQL = strSQL & ",RCPT_KREN_NAME"
    strSQL = strSQL & ",IIF(ISNULL(RCPT_KREN_KEITAI, '') = '', RCPT_KREN_TEL, RCPT_KREN_KEITAI) AS RCPT_KREN_TEL "
    'INSERT 2023/11/01 N.IMAI End
    
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "  CARG_FILE A "
    
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  YARD_MAST "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  YARD_CODE = CARG_YCODE "
    
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  CNTA_MAST "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  CNTA_CODE = CARG_YCODE "
    strSQL = strSQL & "AND "
    
    strSQL = strSQL & "  CNTA_NO = CARG_NO "
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  USER_MAST "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  USER_CODE = CARG_UCODE "
    
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C020 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C020.NAME_CODE = USER_KKBN "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C020.NAME_ID   = '020' "
    
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C060 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C060.NAME_CODE = CONVERT(INTEGER,CARG_KAGIICD) "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C060.NAME_ID   = '060' "
    
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C070 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C070.NAME_CODE = CONVERT(INTEGER,CARG_ADMEI) "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C070.NAME_ID   = '070' "
    
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C088 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C088.NAME_CODE = CARG_AGRE "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C088.NAME_ID   = '088' "
    
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C090 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C090.NAME_CODE = CNTA_STEP "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C090.NAME_ID   = '090' "
    
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C200 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C200.NAME_CODE = CONVERT(INTEGER,CARG_HOSYICD) "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C200.NAME_ID   = '200' "
    
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C086 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C086.NAME_CODE = CONVERT(INTEGER,CNTA_USAGE) "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C086.NAME_ID   = '086' "
    
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C040 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C040.NAME_CODE = CONVERT(INTEGER,CNTA_SYU) "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C040.NAME_ID   = '040' "
    
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C089 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C089.NAME_CODE = CARG_KAICD "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C089.NAME_ID   = '089' "
    
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C092 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C092.NAME_CODE = CARG_KAIRIYUCD "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C092.NAME_ID   = '092' "
    
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & strSHIR_MAST & " "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  SHIRM_SHIRC = CARG_HOSYO_CD "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  SHIRM_BUMOC = 'A' "
    strSQL = strSQL & "Left Join "                                             'Insert 20091005 N.Miura Start
    strSQL = strSQL & "  RCPT_TRAN "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  RCPT_CARG_ACPTNO = CARG_ACPTNO "                      'Insert 20091005 N.Miura Ended
    
    strSQL = strSQL & "Left Join "                                             'INSERT 2020/04/03 N.IMAI Start
    strSQL = strSQL & "  YOUK_TRAN "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  YOUKT_UKNO = RCPT_NO "                                'INSERT 2020/04/03 N.IMAI End
    
    '' 2010/02/12 M.HONDA START
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C014 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C014.NAME_CODE = RCPT_SEIKYU_CD "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C014.NAME_ID   = '014' "
    '' 2010/02/12 M.HONDA END
    '' 2010/02/12 M.HONDA START
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C030 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C030.NAME_CODE = USER_SKBN "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C030.NAME_ID   = '030' "
    '' 2012/04/12 M.HONDA START
    strSQL = strSQL & "INNER Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C294 "
    strSQL = strSQL & "ON "
'    strSQL = strSQL & "  NAME_C294.NAME_CODE = RCPT_SLOPE "                    'DELETE 2025/05/22 K.KINEBUCHI
'  ※ネットから取り込んだ際、NULLとなっていた（5/22にバッチでRCPT_SLOPEは「0」でセットするように修正済み）今後は必要ないと思うが、INNER JOINしているので一応NULL処理を入れておく
    strSQL = strSQL & "  NAME_C294.NAME_CODE = ISNULL(RCPT_SLOPE,0) "           'INSERT 2025/05/22 K.KINEBUCHI
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C294.NAME_ID   = '294' "
    '' 2012/04/12 M.HONDA END
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C201 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C201.NAME_CODE = USER_SKBN "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C201.NAME_ID   = '201' "
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & "( SELECT *"
    strSQL = strSQL & "  FROM  " & strTANT_MAST & " , CONT_MAST "
    strSQL = strSQL & "  WHERE  CONT_BUMOC = TANTM_BUMOC ) TAN  "
    strSQL = strSQL & "ON CARG_UKTANTO = TAN.TANTM_TANTC "
    '' 2010/02/12 M.HONDA END
    '2016/02/06 M.HONDA INS
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & strTINH_MAST & " "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  TINHM_BUMOC = '" & SET_BUMOC & "'"
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  TINHM_TINTC = CARG_YCODE "
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & strBUKE_MAST & " "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  TINHM_BUKHC = BUKEM_BUKHC "
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C210 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C210.NAME_CODE = BUKEM_SYOUI "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C210.NAME_ID   = '210' "
    '2016/02/06 M.HONDA INS
    
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & "  NAME_MAST As NAME_C302 "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  NAME_C302.NAME_CODE = USER_SEIBETU "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  NAME_C302.NAME_ID   = '302' "
    
    '2020/11/30 Takenouchi INS Start
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & strPGPA_TABL & " "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  PGPAT_PGP1B = 'FVS400_COUPON' "
    strSQL = strSQL & "AND "
    strSQL = strSQL & "  PGPAT_PGP2B = RCPT_COUPON_CD "
    '2020/11/30 Takenouchi INS End
    
    '' 2013/04/10 M.HONDA START
    '' 保証委託料の割引は除く
    '' 複数割引がある場合には、割引番号が新しい割引を表示
    strSQL = strSQL & "Left Join "
    strSQL = strSQL & "( "
    strSQL = strSQL & "SELECT DISTINCT "
    strSQL = strSQL & "       B.DCRAT_ACPTNO, "
    strSQL = strSQL & "       B.DCRAT_DCNT_NO ,"
    strSQL = strSQL & "       B.DCRAT_TEXT "
    strSQL = strSQL & "FROM ( "
    strSQL = strSQL & "      SELECT "
    strSQL = strSQL & "             DCRAT_ACPTNO, "
    strSQL = strSQL & "             MAX(DCRAT_DCNT_NO) DCRAT_DCNT_NO "
    strSQL = strSQL & "      FROM "
    strSQL = strSQL & "             DCRA_TRAN "
    strSQL = strSQL & "      WHERE "
    strSQL = strSQL & "             DCRAT_DCNT_NO NOT LIKE 'ZH%' "
    strSQL = strSQL & "      Group BY "
    strSQL = strSQL & "      DCRAT_ACPTNO "
    strSQL = strSQL & "        ) A, "
    strSQL = strSQL & "      DCRA_TRAN b "
    strSQL = strSQL & "WHERE a.DCRAT_ACPTNO = b.DCRAT_ACPTNO "
    strSQL = strSQL & "AND   A.DCRAT_DCNT_NO = B.DCRAT_DCNT_NO "
    strSQL = strSQL & ") DCRA ON "
    strSQL = strSQL & "CARG_ACPTNO = DCRA.DCRAT_ACPTNO "
    '' 2013/04/10 M.HONDA END
    FunSelCarg = strSQL

End Function
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＯＰＥＮ
'        MODULE_ID        :LF0050_DBOPEN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_DBOPEN()
    On Error GoTo Err_LF0050_DBOPEN
    
    Dim strDataSource        As String

    Set dbs = CurrentDb()
 
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & SET_BUMOC & "'"), "")
    
    If strDataSource = "" Then
        MsgBox "SETU_TABLの設定が不正です。", vbExclamation, PROG_ID
        Set KONT_DB = Nothing
    Else
        Set KONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, True, MSZZ007_M00(SET_BUMOC))
    End If
    
    '2009/05/14 INS<S> hirano
    '加瀬DB　OPEN
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"), "")
    If strDataSource = "" Then
        MsgBox "SETU_TABLの設定が不正です。", vbExclamation, PROG_ID
        Set KASE_DB = Nothing
    Else
        Set KASE_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, True, MSZZ007_M00())
    End If
    '2009/05/14 INS<E> hirano

Exit_LF0050_DBOPEN:
    Exit Sub

Err_LF0050_DBOPEN:
    MsgBox Err.Description
    Resume Exit_LF0050_DBOPEN
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＣＬＯＳＥ
'        MODULE_ID        :LF0050_DBCLOS
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_DBCLOS()
    On Error GoTo Err_LF0050_DBCLOS
    
    If Not dbs Is Nothing Then
       dbs.Close
       Set dbs = Nothing
    End If
    '2009/05/14 INS<S> hirano
    If Not KONT_DB Is Nothing Then
       KONT_DB.Close
       Set KONT_DB = Nothing
    End If
    If Not KASE_DB Is Nothing Then
       KASE_DB.Close
       Set KASE_DB = Nothing
    End If
    '2009/05/14 INS<E> hirano
Exit_LF0050_DBCLOS:
    Exit Sub

Err_LF0050_DBCLOS:
    MsgBox Err.Description
    Resume Exit_LF0050_DBCLOS
End Sub

'==============================================================================*
'
'        MODULE_NAME      :LF0052フォームを開く
'        MODULE_ID        :LF0050_EXE1B_Click
'        CREATE_DATE      :2011/07/07 RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_EXE1B_Click()
    On Error GoTo ErrHandler
    Dim strTbl As String
    Dim strBuc As String
    Dim strQ As String
                       
    strTbl = "WK_LQ0052_TEMP"
    strBuc = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1")
        
    dbs.Execute " DELETE FROM " & strTbl & "  ; ", dbFailOnError
    
    strQ = ""
    strQ = strQ & " INSERT INTO " & strTbl
    strQ = strQ & " ( "
    strQ = strQ & " BUMOC,  "
    strQ = strQ & " BUMON,  "
    strQ = strQ & " KOKYC,  "
    strQ = strQ & " KOKYN,  "
    strQ = strQ & " KOTANTN,"
    strQ = strQ & " FLGSEL  "
    strQ = strQ & " )       "
    strQ = strQ & " SELECT                  "
    strQ = strQ & "     '" & strBuc & "',   "
    strQ = strQ & "     '" & DLookup("BUMOM_BUMON", "BUMO_MAST", "BUMOM_BUMOC = '" & strBuc & "'") & "'  ,   "
    strQ = strQ & "     UMT.USER_CODE   ,   "
    strQ = strQ & "     UMT.USER_NAME   ,   "
    strQ = strQ & "     UMT.USER_TANM   ,   "
    strQ = strQ & "     FALSE               "
    strQ = strQ & " FROM ((dbo_CARG_FILE CAF    "
    strQ = strQ & " INNER JOIN dbo_YARD_MAST YMT ON CAF.CARG_YCODE = YMT.YARD_CODE)                                 "
    strQ = strQ & " INNER JOIN dbo_CNTA_MAST CMT ON CAF.CARG_NO = CMT.CNTA_NO AND CAF.CARG_YCODE = CMT.CNTA_CODE)   "
    strQ = strQ & " INNER JOIN dbo_USER_MAST UMT ON CAF.CARG_UCODE = UMT.USER_CODE                                  "
    
    strQ = strQ & LF0050_SEL_WHERE
    strQ = strQ & ";"

    dbs.Execute strQ
    
    doCmd.OpenForm "LF0052"
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :LF0050_M10
'        CREATE_DATE      :2023/12/01   N.IMAI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_M10()

    Dim strSQL  As String
    Dim db      As DAO.Database
    Dim qdf     As DAO.QueryDef
    
    Set db = CurrentDb

    For Each qdf In CurrentDb.QueryDefs
        If qdf.NAME = "LQ0050_M10" Then
            doCmd.DeleteObject acQuery, "LQ0050_M10"
            Exit For
        End If
    Next
    
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "  WK_LQ0050_TEMP.LQ0050_YCODE AS ﾔｰﾄﾞ "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_YNAME AS 名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_NO AS 番号 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_STEP AS 段区分 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_STEPN AS 段 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_AGRE AS ステータス "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_AGREN AS ステータス名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_USAGECD AS 用途区分 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_USAGEN AS 用途名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SLOPE AS スロープ区分 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SLOPEN AS スロープ名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SYUCD AS 種別 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SYUN AS 種別名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SIZE AS 実帖 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_TANTO AS 担当者コード "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_TANTN AS 担当者名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_ACPTNO AS 契約番号 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UCODE AS 顧客ｺｰﾄﾞ "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UNAME AS 顧客名 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UKANA AS 顧客カナ "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UTANM AS 代表者名 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UMAIL AS メールアドレス "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UYUBINO AS 郵便番号 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UADR_1 AS 住所1 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UADR_2 AS 住所2 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UADR_3 AS 住所3 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UTEL AS 電話 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UKEITAI AS 携帯 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UFAX AS ﾌｧｯｸｽ "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KKBN AS 個人法人区分 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KKBNN AS 個人法人区分名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_BIRTHDAY AS 生年月日 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SEIB AS 性別 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KNAME AS 緊急連絡先名 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KTEL AS 緊急連絡先電話 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SEIKYU AS 初回請求区分 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SKBN AS 集金区分 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SKBNN AS 集金区分名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SEIKI AS 請求書発行区分 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SEIKIN AS 請求書発行区分名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KIDATE AS 起算日 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UKDATE AS 受付日 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_FSDATE AS 契約日 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_STDATE AS 開始日 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_EDDATE AS 満了日 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KYDATE AS 解約日 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_CMP_EXDATE AS キャンペーン適用満了日 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_DCRAT_TEXT AS キャンペン内容 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_RENTKG AS 使用料 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SYOZEI AS 消費税 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_GOKEI AS 使用料合計 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_EZAPPI_DAILY AS 毎月雑費日割り "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_EZAPPI AS 毎月雑費 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SECUKG AS 保証金 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_FRSTKG AS 初回金額 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAGIA AS 初回鍵代 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_FSYOZEI AS 初回消費税 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UKNO AS 受付番号 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_MDATE AS 伝票送付日 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_DAHIB AS 伝票番号 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAGIICD AS 鍵種別 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAGIN AS 鍵名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAIJB AS 解除番号 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HOSYICD AS 保証区分 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HOSYN AS 保証名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HOSYO_CD AS 保証会社コード "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HOSYO_NM AS 保証会社名 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HOSYD AS 承認日 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HOSYB AS 承認番号 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_ADMEI AS 広告種別 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_ADMEN AS 広告名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_BUTU AS 収納分 "
    
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_NUMBER AS バイクナンバー "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SHASYU AS 車種 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HAIKIRYO AS 排気量 "
    
    
    
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAICD AS 解約区分 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAICD_NAME AS 解約区分名称 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAIRIYUCD AS 理由コード "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAIRIYUCD_NAME AS 理由 "
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_BIKO AS 備考  "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "  WK_LQ0050_TEMP; "
    Set qdf = db.CreateQueryDef("LQ0050_M10", strSQL)
    
    Set qdf = Nothing
    Set db = Nothing
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :LF0050_M30
'        CREATE_DATE      :2023/12/01   N.IMAI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0050_M30()

    Dim strSQL  As String
    Dim db      As DAO.Database
    Dim qdf     As DAO.QueryDef
    
    Set db = CurrentDb

    For Each qdf In CurrentDb.QueryDefs
        If qdf.NAME = "LQ0050_M30" Then
            doCmd.DeleteObject acQuery, "LQ0050_M30"
            Exit For
        End If
    Next
    
    strSQL = ""
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "  WK_LQ0050_TEMP.LQ0050_YCODE AS ﾔｰﾄﾞ"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_YNAME AS 名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_YYUBINO AS 設置場所郵便番号"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_YADR_1 AS 設置場所住所１"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_YADR_2 AS 設置場所住所２"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_YADR_3 AS 設置場所住所３"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_NO AS 番号"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_STEP AS 段区分"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_STEPN AS 段"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_AGRE AS ステータス"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_AGREN AS ステータス名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_USAGECD AS 用途区分"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_USAGEN AS 用途名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SLOPE AS スロープ区分"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SLOPEN AS スロープ名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SYUCD AS 種別"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SYUN AS 種別名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SIZE AS 実帖"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_TANTO AS 担当者コード"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_TANTN AS 担当者名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_ACPTNO AS 契約番号"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UCODE AS 顧客ｺｰﾄﾞ"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UNAME AS 顧客名"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UKANA AS 顧客カナ"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UTANM AS 代表者名"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UMAIL AS メールアドレス"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UYUBINO AS 郵便番号"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UADR_1 AS 住所1"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UADR_2 AS 住所2"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UADR_3 AS 住所3"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UTEL AS 電話"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UKEITAI AS 携帯"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UFAX AS ﾌｧｯｸｽ"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KKBN AS 個人法人区分"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KKBNN AS 個人法人区分名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_BIRTHDAY AS 生年月日"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SEIB AS 性別"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KNAME AS 緊急連絡先名"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KTEL AS 緊急連絡先電話"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SEIKYU AS 初回請求区分"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SKBN AS 集金区分"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SKBNN AS 集金区分名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SEIKI AS 請求書発行区分"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SEIKIN AS 請求書発行区分名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KIDATE AS 起算日"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UKDATE AS 受付日"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_FSDATE AS 契約日"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_STDATE AS 開始日"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_EDDATE AS 満了日"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KYDATE AS 解約日"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_CMP_EXDATE AS キャンペーン適用満了日"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_DCRAT_TEXT AS キャンペーン内容"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_RENTKG AS 使用料"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SYOZEI AS 消費税"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_GOKEI AS 使用料合計"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_EZAPPI_DAILY AS 毎月雑費日割り"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_EZAPPI AS 毎月雑費"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SECUKG AS 保証金"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_FRSTKG AS 初回金額"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAGIA AS 初回鍵代"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_FSYOZEI AS 初回消費税"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_UKNO AS 受付番号"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_MDATE AS 伝票送付日"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_DAHIB AS 伝票番号"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAGIICD AS 鍵種別"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAGIN AS 鍵名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAIJB AS 解除番号"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HOSYICD AS 保証区分"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HOSYN AS 保証名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HOSYO_CD AS 保証会社コード"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HOSYO_NM AS 保証会社名"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HOSYD AS 承認日"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_HOSYB AS 承認番号"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_ADMEI AS 広告種別"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_ADMEN AS 広告名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_USER_CP_TXT AS 紹介会社名（ユーザ入力）"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_USER_CP_CD AS 紹介コード（ユーザ入力）"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_CP_TXT AS 紹介会社名（確定版）"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_CP_CD AS 紹介コード（確定版）"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_BUTU AS 収納分"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAICD AS 解約区分"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAICD_NAME AS 解約区分名称"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAIRIYUCD AS 理由コード"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_KAIRIYUCD_NAME AS 理由"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_BIKO AS 備考"
    strSQL = strSQL & "  , WK_LQ0050_TEMP.LQ0050_SIZE AS 実帖数 "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "  WK_LQ0050_TEMP;"
    Set qdf = db.CreateQueryDef("LQ0050_M30", strSQL)
    
    Set qdf = Nothing
    Set db = Nothing
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :LF0050_CDATE
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'Delete 20090908 N.Miura
'Function LF0050_chgDATE(strDate As String) As String
''    On Error GoTo Err_LF0050_chgDATE
'
'    LF0050_chgDATE = Mid(strDate, 1, 4) & "/" & _
'                     Mid(strDate, 5, 2) & "/" & _
'                     Mid(strDate, 7, 2)
'
''Exit_LF0050_chgDATE:
''    Exit Sub
'
''Err_LF0050_chgDATE:
''    MsgBox err.Description
''    Resume Exit_LF0050_chgDATE
'End Function
'2009/05/14 INS <S> hirano
'==============================================================================*
'
'        MODULE_NAME      :出力
'        MODULE_ID        :LF0050_INSERT
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'Delete 20090908 N.Miura
'Private Function LF0050_SEL_SHIR(strSHIR_CD As String) As String
'    Dim strSQLSHIR As String
'    Dim strSHIR_NM As String
'    Dim RS_SHIR    As Recordset
'    strSQLSHIR = ""
'    strSHIR_NM = ""
'    strSQLSHIR = "SELECT SHIRM_SHIRN FROM SHIR_MAST WHERE  SHIRM_BUMOC = 'A' AND  SHIRM_SHIRC = '" & strSHIR_CD & "'"
'    Set RS_SHIR = KASE_DB.OpenRecordset(strSQLSHIR, dbOpenForwardOnly, dbReadOnly)
'    strSHIR_NM = RS_SHIR.Fields("SHIRM_SHIRN")
'    RS_SHIR.Close
'    Set RS_SHIR = Nothing
'    LF0050_SEL_SHIR = strSHIR_NM
'    Exit Function
'Err_LF0050_SEL_SHIR:
'    MsgBox Err.Description
'    If RS_SHIR Is Not Null Then
'        RS_SHIR.Close
'        Set RS_SHIR = Nothing
'    End If
'End Function
'2009/05/14 INS <E> hirano
'****************************  ended or program ********************************

