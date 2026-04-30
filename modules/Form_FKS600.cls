VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'********************************************************************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'********************************************************************************************
'   プログラム名　　：　セキュリティカードメンテ
'   プログラムＩＤ　：　FKS600
'   作　成　日　　　：  2006/05/01
'   作　成　者　　　：  tajima
'   Ver             ：  0.0
'********************************************************************************************
'修正履歴
'   修　正　日　　　：2006/06/20
'   修　正　者　　　：tajima
'   修　正　内　容　：ワークテーブルクリア処理の変更
'   Ver             ：0.1
'********************************************************************************************
'修正履歴
'   修　正　日　　　：2006/07/28
'   修　正　者　　　：tajima
'   修　正　内　容　：FVS601対応
'   Ver             ：0.2
'********************************************************************************************
'修正履歴
'   修　正　日　　　：2007/02/08
'   修　正　者　　　：tajima
'   修　正　内　容　：依頼書印刷対応
'   Ver             ：0.3
'********************************************************************************************
'修正履歴
'   修　正　日　　　：2013/06/03
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：終了メッセージを表示させないように修正。
'   Ver             ：0.4
'********************************************************************************************
'修正履歴
'   修　正　日　　　：2015/03/25
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：既に雑費が設定されてる際には、雑費を空にしない
'   Ver             ：0.5
'********************************************************************************************
'修正履歴
'   修　正　日　　　：2023/12/04
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：鍵検索機能追加
'   Ver             ：0.7
'********************************************************************************************
'修正履歴
'   修　正　日　　　：2026/04/06
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：保存時にFVS400の保存処理を実行する（依頼書印刷時は不要）
'   Ver             ：0.8
'********************************************************************************************
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private pobjMY_DB                 As Database                   ' カレントＤＢ
Private pobjKONT_DB               As Database                   ' コンテナDB
Private strConnection             As String                     ' SQL-Server接続文字列

Private pstrFormCaption           As String                     ' 画面標題格納
Private Const pstrCstSubFromName  As String = "FKS600_S01"      ' サブフォーム名
Private Const P_SUB_FORM          As String = "FKS600_S01"      ' サブフォーム名
Private Const P_WORK_TABLE        As String = "FKS600_W01"      ' ワークテーブル
Private isEditChange              As Boolean                    ' 入力に変更があったかどうかのフラグ
Private isNoticeforFVS400         As Boolean                    ' FVS400への変更通知するの可否
Private WRI_ObjWS                 As Workspace
Private WRI_objDb                 As Database
Private pstrWORKData()            As String       ' ワークテーブルデータ格納配列(セキュリティカードマスタ)

Private pFromPROG As Integer
' 構造体(起動パラメータ保持)
Private Type Type_SCRD
    FROM_ID     As String         ' 遷移元
    YARDC       As String         ' ヤードコード
    ROOMC       As String         ' 部屋番号(コンテナ番号)
    ACPTNO      As String         ' 契約No
End Type
Private pParamArgs As Type_SCRD

' サブフォームが参照するための情報
Public pFROM_IDforFKS600_S01        As String   ' 遷移元情報
'Public pRoomRowSource_forFKS600_S01 As String   ' サブフォームの部屋リストのネタ→コントロール内でクエリとした
Public pUseRoomNoforFKS600_S01      As String   ' 使用(受付契約)している部屋番号
Public pAcptNoforFKS600_S01         As String   ' 使用(受付契約)している契約番号
Public pUserNameforFKS600_S01       As String   ' 使用(受付契約)している顧客名

' 呼出元に反映させるためのネタ情報
Private pUsedCardRowSource           As String   ' RowSourceの為に編集したカード番号達
Private pBaseUsedCardCount           As Integer  ' 基本使用枚数
Private pAddUsedCardCount            As Integer  ' 追加使用枚数
Private p元AddCardPrice              As Variant  ' 元々の追加カード代(RCPT_TRANの)

'カード依頼書印刷範囲
Private p停止日FromDate              As String  '2007/02/08 add
Private p停止日ToDate                As String  '2007/02/08 add
'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_fromMENU   As String = "MENU"     ' メニューからの印
Private Const P_fromFKS440 As String = "FKS440"   ' ヤードメンテからの印
Private Const P_fromFVS400 As String = "FVS400"   ' 受付入力からの印
Private Const P_fromFVS600 As String = "FVS600"   ' 契約入力からの印
Private Const P_fromFVS601 As String = "FVS601"   ' 契約入力からの印 2006/07/28 Add

Private Const P_KAKUNIN                 As String = "確認"
Private Const P_ERROR                   As String = "エラー"
Private Const P_部屋割当無し            As String = "無し"
Private Const P_部屋割当無し値          As Integer = -1
Private Const P_用途_予備               As String = "00"
Private Const P_雑費区分_追加カード代値 As String = "84"  '名称マスタ 96:雑費区分参照

Private Const PROG_ID           As String = "FKS600"
Private Const MAX_CARD_NO       As String = "~~~~~~~~~~~" 'ニョロがバイナリ比較で一番大ｷｲﾖ

'==============================================================================*
'
'        MODULE_NAME      :ヤードコード変更後処理
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_AfterUpdate()

    Dim isException As Boolean

  On Error GoTo Exception
    isException = False
        
    ' 体裁を整える
    Me.cmb_Yard = Format(Me.cmb_Yard, "000000")
        
    ' 各ＤＢへの接続
    Set pobjMY_DB = CurrentDb
    If fncConnectDB() = False Then
        Call MSZZ024_M10("fncConnectDB", "コンテナDB接続処理が失敗しました。")
    End If
        
    ' ヤード情報を引っ張る
    Call setHedderText(P_fromMENU, Me.cmb_Yard)
    
    ' サブフォーム用のワーク生成
    Call makeFKS600_W01(P_fromMENU, Me.cmb_Yard)
        
    GoTo Finally
    
Exception:
    isException = True
    
Finally:
    If Not pobjMY_DB Is Nothing Then pobjMY_DB.Close: Set pobjMY_DB = Nothing
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing
    If isException = True Then
        Call MSZZ024_M00("cmb_Yard_AfterUpdate", True)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ヤードコード変更前処理
'        MODULE_ID        :cmb_Yard_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_BeforeUpdate(Cancel As Integer)
    
    Dim vCode As Variant
    
    On Error GoTo Exception
    
    vCode = Nz(Me.cmb_Yard, "")
    
    '入力チェック
    Select Case CmbCheck_Num(vCode, 6)
        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        Case 1, 9
            MsgBox "ヤードコードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        Case Else
    End Select
    
    'ヤードの存在チェック
    If vCode <> "" And isYardExistence(vCode) = False Then
       MsgBox "入力されたヤードは存在しません。", vbInformation, P_ERROR
       Cancel = True
       SendKeys "{ESC}"
       Exit Sub
    End If
    
    ' 変更フラグがTrueの時は入力取消の確認メッセージを表示
    If isEditChange = True Then
        If vbNo = MsgBox("データ入力を取り消します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) Then
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        End If
    End If
    
    Exit Sub
    
Exception:
    Call MSZZ024_M00("cmb_Yard_BeforeUpdate", True)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_Yard, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_カード代_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_カード代_AfterUpdate()
    Call SetFKS600EditChange
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_カード代_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_カード代_BeforeUpdate(Cancel As Integer)
    
    Dim vNum As Variant
    Dim errorMsg As String
    
  On Error GoTo Exception

    vNum = Nz(Me.txt_カード代, "")

    If vNum = "" Then Exit Sub
    
    '入力チェック
    Select Case CmbCheck_Num(vNum, 9)
        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        Case 1, 9
            MsgBox "カード代として金額を入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        Case Else
    End Select
    
    Exit Sub
    
Exception:
    Call MSZZ024_M00("txt_カード代_BeforeUpdate", True)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_カード代_GotFocus
'        NOTE             :
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_カード代_GotFocus()
    Me.txt_カード代.Format = "Fixed"  ' 書式：固定(カンマ編集なし)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_カード代_LostFocus
'        NOTE             :
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_カード代_LostFocus()
    Me.txt_カード代.Format = "Standard"  ' 書式：標準(カンマ編集あり)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_カード代_KeyPress
'        NOTE             :
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_カード代_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_カード代, KeyAscii, 9)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :4
'        NOTE             :
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function check_txt_カード代(ByVal aカード代 As Variant) As Boolean
     
    
    If Nz(aカード代, "") = "" And pAddUsedCardCount > 0 Then
        MsgBox "追加カードが" & pAddUsedCardCount & "枚あります。" & vbCr & _
               "追加カード代の入力が必要です。", vbInformation, P_ERROR
        check_txt_カード代 = False
        Exit Function
    
    ElseIf Nz(aカード代, "") <> "" And pAddUsedCardCount = 0 Then
        If vbNo = MsgBox("追加カードの指定はありません。" & vbCr & _
                         "追加カード代は無しとします。宜しいですか？", _
                          vbYesNo + vbExclamation + vbDefaultButton2, P_ERROR) Then
            check_txt_カード代 = False
            Exit Function
        End If
        txt_カード代.VALUE = Null 'カード代を未入力とする
    End If
    check_txt_カード代 = True
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :ヤードマスタの存在可否
'        MODULE_ID        :isYardExistence
'        NOTE             :パラメータで指定されたヤードコードの存在チェックを行う
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function isYardExistence(ByVal aYardCode As Long) As Boolean
        
    Dim strSQL  As String
    Dim vRet    As Variant
    
    isYardExistence = False
    On Error GoTo Exception
    
    strSQL = " YARD_CODE =" & aYardCode
    vRet = Nz(DLookup("YARD_CODE", "dbo_YARD_MAST", strSQL), "")
        
    If vRet <> "" Then
        isYardExistence = True
    End If
    Exit Function
Exception:
    Call Err.Raise(Err.Number, "isYardExistence" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
   
End Function

'==============================================================================*
'
'        MODULE_NAME      :部屋の存在可否
'        MODULE_ID        :isRoomExistence
'        NOTE             :パラメータで指定された部屋の存在チェックを行う
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function isRoomExistence(ByVal aYardCode As Long, ByVal aRoomNo As Long) As Boolean
        
    Dim strSQL  As String
    Dim vRet    As Variant
    
    isRoomExistence = False
    On Error GoTo Exception
    
    strSQL = " CNTA_CODE =" & aYardCode & " AND CNTA_NO=" & aRoomNo
    vRet = Nz(DLookup("CNTA_NO", "dbo_CNTA_MAST", strSQL), "")
        
    If vRet <> "" Then
        isRoomExistence = True
    End If
    Exit Function
Exception:
    Call Err.Raise(Err.Number, "isRoomExistence" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
   
End Function

'==============================================================================*
'
'        MODULE_NAME      :終了ボタン押下時処理
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_End_Click()

    Dim strMessage    As String
    Dim varDefBtn     As Variant
    
    On Error Resume Next
    doCmd.RunCommand acCmdSaveRecord
    
    '' 2013/06/03 M.HONDA START
    ' 変更フラグがTrueの時は更新の確認メッセージを表示、それ以外は終了の確認メッセージを表示
'    strMessage = IIf(isEditChange = True, "修正中ですが保存せずに終了しますか？", "終了します。よろしいですか？")
'    varDefBtn = vbYesNo + vbQuestion + IIf(isEditChange = True, vbDefaultButton2, vbDefaultButton1)
'
'    ' 確認メッセージ表示
'    If MsgBox(strMessage, varDefBtn, P_KAKUNIN) = vbNo Then
'        Exit Sub
'    End If
    '' 2013/06/03 M.HONDA END
    
    On Error Resume Next
    doCmd.Close acForm, Me.FormName

End Sub

'==============================================================================*
'
'        MODULE_NAME      :コンテナDB接続処理
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncConnectDB() As Boolean

    Dim strBumoCD     As String
    Dim strConnect    As String
    Dim strDataSource As String
    
    On Error GoTo Exception
    
    fncConnectDB = False
    
    ' 部門コード取得
    strBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    
    ' コンテナDB接続
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumoCD & "'"), "")
    
    If strDataSource = "" Then
        Set pobjKONT_DB = Nothing
        Call MSZZ024_M10("fncConnectDB", "SETU_TABLの設定が不正です。")
    Else
        Set pobjKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(strBumoCD))
    End If
    
    fncConnectDB = True
    Exit Function
    
Exception:
    Call Err.Raise(Err.Number, "fncConnectDB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function


'==============================================================================*
'
'        MODULE_NAME      :保存ボタン押下時
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Enter_Click()

    Dim errorMessage  As String
    Dim isException As Boolean
  
  On Error GoTo Exception
    isException = False
    
    If Nz(Me.cmb_Yard.VALUE, "") = "" Then Exit Sub
    
    errorMessage = ""
    
    ' 受付から来た刻のみ必須入力チェック
    If pParamArgs.FROM_ID = P_fromFVS400 Then
        Call getUseCardNumforWORK 'ワークで使用している追加枚数を取得
        ' カード代を入力するのかしてはいけないのかチェック
        If check_txt_カード代(txt_カード代.VALUE) = False Then
            txt_カード代.SetFocus
            Exit Sub
        End If
    End If
    
    
    '確認ﾒｯｾｰｼﾞ
    If vbNo = MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) Then Exit Sub

    ' 各ＤＢへの接続
    Set pobjMY_DB = CurrentDb
    If fncConnectDB() = False Then
        Call MSZZ024_M10("fncConnectDB", "コンテナDB接続処理が失敗しました。")
    End If
    
    ' テーブルへの更新処理
    Call writeSCRDtable(errorMessage)

    ' 選択系のみの場合の処理
    If IsSelectModeFKS600() Then
        ' 更新失敗のメッセージ　※選択系のみの場合発生しうる
        If errorMessage <> "" Then
            MsgBox "選んだカードで以下のものは更新できませんでした。" & vbCr & errorMessage, vbOKOnly, P_KAKUNIN
        End If
                        
        ' 現状の貸出カード内容を取得しておく
        Call CmKomsMod.GetContractSCards(pobjKONT_DB, pParamArgs.ACPTNO, pBaseUsedCardCount, pAddUsedCardCount, pUsedCardRowSource)
        ' 契約からの場合に呼ぶ外部プロシージャ
        If pParamArgs.FROM_ID = P_fromFVS600 Then
            Call Form_FVS600.SetCardInfomation(pBaseUsedCardCount, pAddUsedCardCount, pUsedCardRowSource)
        ElseIf pParamArgs.FROM_ID = P_fromFVS601 Then                                                     '2006/07/28 ADD tajima
            Call Form_FVS601.SetCardInfomation(pBaseUsedCardCount, pAddUsedCardCount, pUsedCardRowSource) '2006/07/28 ADD tajima
        ' 受付からの場合に呼ぶ外部プロシージャ
        ElseIf pParamArgs.FROM_ID = P_fromFVS400 Then
            Call Form_FVS400.subSCardRefresh(pBaseUsedCardCount, pAddUsedCardCount, pUsedCardRowSource, isNoticeforFVS400)
            Call Form_FVS400.fncDataSave(1, False)                              '2026/04/06 N.IMAI
        End If
                        
        If vbYes = MsgBox("保存しました。元の画面に戻りますか？", vbQuestion + vbYesNo, P_KAKUNIN) Then
            On Error Resume Next
            doCmd.Close acForm, Me.FormName
            GoTo Finally
        End If
    Else
        MsgBox "保存しました。", , P_KAKUNIN
    End If
    
    ' 再描画（初期状態に戻す）
    If pParamArgs.FROM_ID = P_fromMENU Then
        'メニューからきたときは画面クリア
        Me.cmb_Yard = Null
        Me.txt_ヤード名 = Null
        Me.txt_ヤード所在地 = Null
        Call clearFKS600_W01(P_WORK_TABLE)
    Else
        Call makeFKS600_W01(pParamArgs.FROM_ID, pParamArgs.YARDC, pParamArgs.ROOMC)
    End If
    ' サブフォーム初期設定
    Call initSubForm
    
    ' 各制御の初期化
    isEditChange = False
    isNoticeforFVS400 = False
    Me.Caption = pstrFormCaption
    
    GoTo Finally
    
Exception:
    isException = True
    
Finally:
    If Not pobjMY_DB Is Nothing Then pobjMY_DB.Close: Set pobjMY_DB = Nothing
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing
    If isException = True Then
        Call MSZZ024_M00("cmd_Enter_Click", True)
    End If
 End Sub

'==============================================================================*
'
'        MODULE_NAME      :取消ボタン押下時処理
'        MODULE_ID        :cmd_Kesi_Click
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Kesi_Click()
    
    Dim isException As Boolean

  On Error GoTo Exception
    
    isException = False
    
    If Nz(Me.cmb_Yard.VALUE, "") = "" Then Exit Sub
    
    '確認ﾒｯｾｰｼﾞ
    If vbNo = MsgBox("データ入力を取り消します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) Then
        Exit Sub
    End If

    ' 各ＤＢへの接続
    Set pobjMY_DB = CurrentDb
    If fncConnectDB() = False Then
        Call MSZZ024_M10("fncConnectDB", "コンテナDB接続処理が失敗しました。")
    End If

    If pParamArgs.FROM_ID = P_fromMENU Then
    'メニューから来た場合は全てクリア
        Me.cmb_Yard = ""
        Me.txt_ヤード名 = ""
        Me.txt_ヤード所在地 = ""
        Call clearFKS600_W01(P_WORK_TABLE)
    Else
    'メニュー以外から来た場合は検索直後の状態とする
         Call makeFKS600_W01(pParamArgs.FROM_ID, pParamArgs.YARDC, pParamArgs.ROOMC)
         ' サブフォーム初期設定
         Call initSubForm
    End If
    
    '取消処理が終わったら初期化
    isEditChange = False
    isNoticeforFVS400 = False
    Me.Caption = pstrFormCaption
    
    GoTo Finally
    
Exception:
    isException = True
    
Finally:
    If Not pobjMY_DB Is Nothing Then pobjMY_DB.Close: Set pobjMY_DB = Nothing
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing
    If isException = True Then
        Call MSZZ024_M00("cmd_Kesi_Click", True)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    Dim iParamCnt   As Integer
    Dim strParam()  As String

    On Error GoTo Exception

' ↓↓↓本番用↓↓↓ -----------------------
    ' 起動パラメータ取得
    strParam = Split(Nz(Me.OpenArgs), ",")
' ↑↑↑本番用↑↑↑------------------------
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_標題.BackColor = RGB(255, 0, 204)
    End Select

' ↓↓↓検証用↓↓↓ -----------------------
    ' 起動パラメータ取得
' 【ヤード】
'    strParam = Split(",,H066000495", ",")
' 【ヤード&コンテナ(部屋)】
'    strParam = Split("900219,,", ",")
' ↑↑↑検証用↑↑↑------------------------

    ' パラメータ数チェック
    iParamCnt = UBound(strParam)
    ' パラメータ無しはメニューから来た
    If iParamCnt = -1 Then
        pParamArgs.FROM_ID = P_fromMENU
        pParamArgs.YARDC = ""
        pParamArgs.ROOMC = ""
        pParamArgs.ACPTNO = ""
    ' パラメータが４つならば他画面から来た
    ElseIf iParamCnt = 3 Then
        pParamArgs.FROM_ID = strParam(0) ' 遷移元
        pParamArgs.YARDC = strParam(1)   ' ヤードコード
        pParamArgs.ROOMC = strParam(2)   ' 部屋番号
        pParamArgs.ACPTNO = strParam(3)  ' 契約No
    Else
        Cancel = True
        Call MSZZ024_M10("OpenArgs", "起動パラメータが正しくありません。")
        Exit Sub
    End If

    ' パラメータ整合性チェック
    Select Case pParamArgs.FROM_ID
        'メニューから
        Case P_fromMENU
            ' ★特にチェックしないが最大化は行う
            doCmd.Maximize
        'ヤードメンテから
        Case P_fromFKS440
            If pParamArgs.YARDC = "" Then
                Call MSZZ024_M10("fromFKS440-ParamCheck", "ヤードが指定されてない。")
            End If
        
        '受付入力から,契約入力から
        Case P_fromFVS400, P_fromFVS600, P_fromFVS601 '2006/07/28 "P_fromFVS601" をAdd
            If pParamArgs.YARDC = "" Or pParamArgs.ROOMC = "" Or pParamArgs.ACPTNO = "" Then
                Call MSZZ024_M10("from" & pParamArgs.FROM_ID & "-ParamCheck", "ヤードまたは部屋または契約番号が指定されてない｡ ")
            End If
    End Select

    Exit Sub

Exception:
    Call MSZZ024_M00("Form_Open", True)
    Cancel = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :画面起動時処理
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    
    Dim isException As Boolean
  
  On Error GoTo Exception
    isException = False
        
    ' 画面の標題格納
    pstrFormCaption = Me.Caption
    
    ' 変更フラグ初期化
    isEditChange = False
    isNoticeforFVS400 = False
    
    ' 各ＤＢへの接続
    Set pobjMY_DB = CurrentDb
    If fncConnectDB() = False Then
        Call MSZZ024_M10("fncConnectDB", "コンテナDB接続処理が失敗しました。")
    End If
    
    ' サブフォームが参照する値の設定
    pUseRoomNoforFKS600_S01 = pParamArgs.ROOMC
    pAcptNoforFKS600_S01 = pParamArgs.ACPTNO
    pUserNameforFKS600_S01 = "" 'setHedderTextで取得します
    
    'パラメータに従い表示モードを変える
    pFROM_IDforFKS600_S01 = pParamArgs.FROM_ID 'サブフォームが参照する情報を設定
    Select Case pParamArgs.FROM_ID
        'メニューから
        Case P_fromMENU
            ' 画面ヘッダの設定
            Me.lbl_標題.Caption = "セキュリティカードマスタ"
            Me.cmb_Yard.Enabled = True
            Me.txt_コンテナ番号.Visible = False
            Me.lbl_コンテナ番号.Visible = False
            Me.txt_段区分.Visible = False
            Me.lbl_段区分.Visible = False
            Me.txt_サイズ.Visible = False
            Me.lbl_サイズ.Visible = False
            Me.txt_カード代.Visible = False
            Me.lbl_カード代.Visible = False
            Me.lbl_ガイダンス.Visible = False
        
            ' サブフォーム用のワーク生成
            Call clearFKS600_W01(P_WORK_TABLE)
            
        'ヤードメンテから
        Case P_fromFKS440
            Me.SecurityCardsList.Enabled = True
            
            If isYardExistence(CLng(pParamArgs.YARDC)) = False Then
                Call MSZZ024_M10("IsYardExistence", "指定されたヤード[" & pParamArgs.YARDC & "]は存在しません。")
            End If
            
            ' 画面ヘッダの設定
            Me.lbl_標題.Caption = "セキュリティカードマスタ"
            Call setHedderText(P_fromFKS440, pParamArgs.YARDC)
            Me.cmb_Yard = pParamArgs.YARDC
            Me.cmb_Yard.Enabled = False
            Me.txt_コンテナ番号.Visible = False
            Me.lbl_コンテナ番号.Visible = False
            Me.txt_段区分.Visible = False
            Me.lbl_段区分.Visible = False
            Me.txt_サイズ.Visible = False
            Me.lbl_サイズ.Visible = False
            Me.txt_カード代.Visible = False
            Me.lbl_カード代.Visible = False
            Me.lbl_ガイダンス.Visible = False

            ' サブフォーム用のワーク生成
            Call makeFKS600_W01(P_fromFKS440, pParamArgs.YARDC)
            ' サブフォーム初期設定
            Call initSubForm
        
        '受付入力 & 契約入力から
        Case P_fromFVS400, P_fromFVS600, P_fromFVS601 '2006/07/28 "P_fromFVS601" をAdd
        
            If isRoomExistence(CLng(pParamArgs.YARDC), CLng(pParamArgs.ROOMC)) = False Then
                Call MSZZ024_M10("IsRoomExistence", "指定されたヤード[" & pParamArgs.YARDC & "]コンテナ[" & pParamArgs.ROOMC & "]は存在しません。")
            End If
        
            ' 画面ヘッダの設定
            Me.lbl_標題.Caption = "セキュリティカード選択"
            Call setHedderText(P_fromFVS400, pParamArgs.YARDC, pParamArgs.ROOMC)
            Me.cmb_Yard = pParamArgs.YARDC
            Me.cmb_Yard.Enabled = False
            Me.txt_コンテナ番号.Visible = True
            Me.lbl_コンテナ番号.Visible = True
            Me.txt_段区分.Visible = True
            Me.lbl_段区分.Visible = True
            Me.txt_サイズ.Visible = True
            Me.lbl_サイズ.Visible = True
            Me.lbl_ガイダンス.Visible = True
            If pParamArgs.FROM_ID = P_fromFVS400 Then ' 受付のときのみカード代入力可
                Me.txt_カード代.Visible = True
                Me.lbl_カード代.Visible = True
            Else
                Me.txt_カード代.Visible = False ' 契約のときは入力不可
                Me.lbl_カード代.Visible = False
            End If
        
            ' サブフォーム用のワーク生成
            Call makeFKS600_W01(pParamArgs.FROM_ID, pParamArgs.YARDC, pParamArgs.ROOMC)
            ' サブフォーム初期設定
            Call initSubForm
        
    End Select
          
    GoTo Finally
    
Exception:
    isException = True
    
Finally:
    If Not pobjMY_DB Is Nothing Then pobjMY_DB.Close: Set pobjMY_DB = Nothing
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing
    If isException = True Then
        Call MSZZ024_M00("Form_Load", True)
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :Form_Unload
'        機能             :Form_Unloadイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)
    
  On Error GoTo Exception
  
    Call clearFKS600_W01(P_WORK_TABLE)
    
    If Not pobjMY_DB Is Nothing Then pobjMY_DB.Close: Set pobjMY_DB = Nothing
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing
    
    Exit Sub
    
Exception:
    Call MSZZ024_M00("Form_Unload", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDownイベント
'        MODULE_ID        :Form_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDown(KeyCode, Shift)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    ' フルスクリーン化
'    If IsSelectModeFKS600() = False Then 2006/02/08 del tajima
        doCmd.Maximize
'    End If                               2006/02/08 del tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :画面のヘッダー部に表示する
'        MODULE_ID        :setHedderText
'        NOTE             :
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setHedderText(ByVal a表示種別 As String, _
                          ByVal aヤードコード As String, _
                          Optional a部屋番号 As String = "0" _
                          )

    Dim strSQL      As String
    Dim sqlSelect   As String
    Dim sqlFrom     As String
    Dim sqlWhere    As String
    Dim rsObject    As Recordset
    Dim isException As Boolean
    
  On Error GoTo Exception
    
    isException = False
    
    ' ベースとなるSQL ヤード情報取得
    sqlSelect = "SELECT "
    sqlSelect = sqlSelect & " YARD_MAST.YARD_NAME "
    sqlSelect = sqlSelect & ",YARD_MAST.YARD_ADDR_1 "
    sqlSelect = sqlSelect & ",YARD_MAST.YARD_ADDR_2 "
    sqlSelect = sqlSelect & ",YARD_MAST.YARD_ADDR_3 "

    sqlFrom = " FROM  "
  
    sqlWhere = " WHERE "
    
    If IsSelectModeFKS600(a表示種別) = False Then
        'メニュー、ヤードメンテからの場合
        ' ヤード情報のみ
        sqlFrom = sqlFrom & " YARD_MAST "
    
        sqlWhere = sqlWhere & " YARD_CODE =" & aヤードコード
        
    Else
        '受付入力、契約入力からの場合
        ' コンテナのサイズ、段区分名、使用用途名、顧客名を取得する
        sqlSelect = sqlSelect & ",CNTA_MAST.CNTA_SIZE "
        sqlSelect = sqlSelect & ",STEP_NAME.NAME_NAME AS STEP_NAME "
        sqlSelect = sqlSelect & ",USAGE_NAME.NAME_NAME AS USAGE_NAME "
        sqlSelect = sqlSelect & "," & getUserNameSubQuery() ' 顧客名を取得するサブクエリー
        
        sqlFrom = sqlFrom & "  CNTA_MAST INNER JOIN YARD_MAST ON CNTA_MAST.CNTA_CODE = YARD_MAST.YARD_CODE"
        sqlFrom = sqlFrom & "  LEFT OUTER JOIN NAME_MAST AS STEP_NAME"
        sqlFrom = sqlFrom & "    ON STEP_NAME.NAME_CODE = CNTA_MAST.CNTA_STEP"
        sqlFrom = sqlFrom & "  LEFT OUTER JOIN NAME_MAST AS USAGE_NAME"
        sqlFrom = sqlFrom & "    ON USAGE_NAME.NAME_CODE = CNTA_MAST.CNTA_USAGE"

        sqlWhere = sqlWhere & " CNTA_CODE  =" & aヤードコード
        sqlWhere = sqlWhere & " AND CNTA_NO =" & a部屋番号
        sqlWhere = sqlWhere & " AND STEP_NAME.NAME_ID = '090'"
        sqlWhere = sqlWhere & " AND USAGE_NAME.NAME_ID = '086'"
        
        If pParamArgs.FROM_ID = P_fromFVS400 Then
            sqlSelect = sqlSelect & "," & getAddCardPriceSubQuery() ' 追加カード代を取得するサブクエリー
        End If
        
    End If

    ' SQL文 Let'sコンバイン
    strSQL = sqlSelect & sqlFrom & sqlWhere
    ' SQLで情報取得し、画面に設定する
    Set rsObject = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    
    If Not (rsObject.BOF And rsObject.EOF) Then
        Me.txt_ヤード名 = rsObject.Fields("YARD_NAME")
        Me.txt_ヤード所在地 = rsObject.Fields("YARD_ADDR_1") & rsObject.Fields("YARD_ADDR_2") & rsObject.Fields("YARD_ADDR_3")
    
        If IsSelectModeFKS600(a表示種別) = True Then
            Me.txt_コンテナ番号 = a部屋番号
            Me.txt_サイズ = Format$(rsObject.Fields("CNTA_SIZE"), "0.00")
            Me.txt_段区分 = rsObject.Fields("STEP_NAME")
            pUserNameforFKS600_S01 = Nz(rsObject.Fields("USER_NAME"), "")
            If pParamArgs.FROM_ID = P_fromFVS400 Then '受付からのときのみカード代設定
                Me.txt_カード代 = rsObject.Fields("CARD_PRICE") 'Nullあり得る
                p元AddCardPrice = Me.txt_カード代
            End If
        End If
    
    End If
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "setHedderText" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルの生成
'        MODULE_ID        :makeFKS600_W01
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub makeFKS600_W01(ByVal a表示種別 As String, _
                                ByVal aヤードコード As String, _
                                Optional a部屋番号 As String = "0" _
                                )

    Dim rsSqlServer         As Recordset
    Dim rsAccess            As Recordset
    Dim i                   As Integer
    Dim strSQL              As String
    Dim isException         As Boolean
    
    On Error GoTo Exception
    isException = False
    
    'Access側ワークテーブル初期化処理
    If clearFKS600_W01(P_WORK_TABLE) = False Then
        Call MSZZ024_M10("MakeFKS600_W01", "ワークテーブル初期化が失敗しました。")
    End If
    
    ' KONT系からワークテーブルに取得する内容
    strSQL = strSQL & "SELECT SCRDM_CARD_NO"
    strSQL = strSQL & ", CASE WHEN SCRDM_ACPTNO IS NULL THEN SCRDM_OTHER_NAME"
    strSQL = strSQL & "  ELSE (SELECT USER_NAME FROM USER_MAST"
    strSQL = strSQL & "       WHERE USER_CODE = ISNULL(CARG_UCODE, RCPT_UCODE))"
    strSQL = strSQL & "  END SCRDM_OTHER_NAME"
    strSQL = strSQL & ", SCRDM_ROOM_NO"
    strSQL = strSQL & ", SCRDM_USAGE, SCRDM_LENDING_DAY, SCRDM_RETURN_DAY"
    strSQL = strSQL & ", SCRDM_STOP_DAY, SCRDM_NOTE"
    strSQL = strSQL & ", CASE WHEN SCRDM_LENDING_DAY IS NULL THEN 0 ELSE -1 END CHK_LENDING_DAY "
    strSQL = strSQL & ", CASE WHEN SCRDM_RETURN_DAY  IS NULL THEN 0 ELSE -1 END CHK_RETURN_DAY "
    strSQL = strSQL & ", CASE WHEN SCRDM_STOP_DAY    IS NULL THEN 0 ELSE -1 END CHK_STOP_DAY "
    strSQL = strSQL & ", SCRDM_ACPTNO "
    strSQL = strSQL & ", 0 AS CHK_CHANGE "  '修正フラグは最初は0
    strSQL = strSQL & ", SCRDM_ACPTNO AS MOTO_ACPTNO "
    strSQL = strSQL & ", SCRDM_RETURN_RNO AS MOTO_ROOM_NO "
    strSQL = strSQL & ", SCRDM_USAGE AS MOTO_USAGE "
    
    strSQL = strSQL & "FROM SCRD_MAST "
    strSQL = strSQL & "LEFT JOIN RCPT_TRAN "
    strSQL = strSQL & " ON RCPT_CARG_ACPTNO = SCRDM_ACPTNO "
    strSQL = strSQL & "LEFT JOIN CARG_FILE "
    strSQL = strSQL & " ON CARG_ACPTNO = SCRDM_ACPTNO "

    ' このメンテの表示種別によって条件を変える
    If IsSelectModeFKS600(a表示種別) = False Then
    ' メニューとヤードメンテからならば
        strSQL = strSQL & "WHERE SCRDM_YCODE = " & aヤードコード
        strSQL = strSQL & " ORDER BY SCRDM_CARD_NO"
    Else
    ' 受付、契約入力から
        ' 自分の契約または指定した部屋または予備の鍵が対象
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & " SCRDM_ACPTNO = '" & pParamArgs.ACPTNO & "' OR "
        strSQL = strSQL & " (SCRDM_YCODE = " & aヤードコード & " AND SCRDM_ROOM_NO = " & a部屋番号
        strSQL = strSQL & "                                      AND SCRDM_ACPTNO IS NULL ) OR "
        strSQL = strSQL & " (SCRDM_YCODE = " & aヤードコード & " AND SCRDM_USAGE = '" & P_用途_予備 & "'"
        strSQL = strSQL & "                                      AND SCRDM_ROOM_NO = " & P_部屋割当無し値 & ") "
        strSQL = strSQL & " ORDER BY SCRDM_CARD_NO"
    End If
    
    Set rsSqlServer = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
  
    ' KONT系から取得した結果をワークに書き込む
    Set rsAccess = pobjMY_DB.OpenRecordset("FKS600_W01", dbOpenDynaset, dbAppendOnly)
    Do While Not rsSqlServer.EOF
        rsAccess.AddNew
        For i = 0 To rsSqlServer.Fields.Count - 1
            If i = 2 Then
                ' "2"の SCRDM_ROOM_NO のみ特殊
                If rsSqlServer.Fields("SCRDM_ROOM_NO") = P_部屋割当無し値 Then
                    rsAccess.Fields("SCRDM_ROOM_NO") = P_部屋割当無し
                Else
                    rsAccess.Fields("SCRDM_ROOM_NO") = Format$(rsSqlServer.Fields("SCRDM_ROOM_NO"), "000000")
                End If
            ElseIf i = 14 Then
                ' "14"の MOTO_ROOM_NO のみ特殊
                If rsSqlServer.Fields("MOTO_ROOM_NO") = P_部屋割当無し値 Then
                    rsAccess.Fields("MOTO_ROOM_NO") = P_部屋割当無し
                Else
                    rsAccess.Fields("MOTO_ROOM_NO") = Format$(rsSqlServer.Fields("MOTO_ROOM_NO"), "000000")
                End If
            Else
                rsAccess.Fields(rsSqlServer.Fields(i).NAME) = rsSqlServer.Fields(i).VALUE
            End If
        Next
        rsAccess.UPDATE
        rsSqlServer.MoveNext
    Loop
    
    Me.SecurityCardsList.Visible = True
    Me.SecurityCardsList.Requery
 
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If Not rsAccess Is Nothing Then rsAccess.Close: Set rsAccess = Nothing
    If Not rsSqlServer Is Nothing Then rsSqlServer.Close: Set rsSqlServer = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "makeFKS600_W01" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ClearFKS600_W01
'        機能             :ワークテーブルクリア
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function clearFKS600_W01(strTable As String) As Boolean
      
    Dim sqlText     As String
    Dim objDb       As Database
    Dim isException As Boolean
    
 On Error GoTo Exception
    isException = False
    clearFKS600_W01 = False
    
    Set objDb = CurrentDb()

    ' ワーク削除
    ' →登録・修正・削除があり得るワークテーブルはドロップクリエイトだとロックにより
    '   ドロップできない場合があり得るのでDeleteとした
    sqlText = "DELETE FROM " & strTable
    objDb.Execute sqlText, dbFailOnError
    
    Me.SecurityCardsList.Visible = False
    clearFKS600_W01 = True
    GoTo Finally
   

Exception:
    isException = True

Finally:
    objDb.Close
    Set objDb = Nothing
        
    If isException = True Then
        Call Err.Raise(Err.Number, "clearFKS600_W01" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

' ▼ ワークテーブルをドロップするとロックする場合がある
'
'    Dim tdfNew      As TableDef
'    Dim isException As Boolean
'    Dim isConnect   As Boolean
'    Dim idx         As Integer
'Dim errlog As String
'  On Error GoTo Exception
'
'    clearFKS600_W01 = False
'    isException = False
'    isConnect = False
'errlog = errlog & "①:"
'    If pobjMY_DB Is Nothing Then
'        Set pobjMY_DB = CurrentDb
'        isConnect = True
'errlog = errlog & "関数内でワークDBを開く:"
'    End If
'errlog = errlog & "②:"
'
'    Me.SecurityCardsList.SourceObject = ""
'
'errlog = errlog & "③:"
'    'ワークテーブル削除
'    If TableExist(pobjMY_DB, strTable) Then
'errlog = errlog & "③-α:【" & strTable & "】"
'        Call pobjMY_DB.TableDefs.Delete(strTable)
'    End If
'errlog = errlog & "④:"
'
'    'ワークテーブル作成
'    Set tdfNew = pobjMY_DB.CreateTableDef(strTable)
'    With tdfNew                                                         'カラム追加
'        Call .Fields.Append(.CreateField("SCRDM_CARD_NO", dbText, 11))       'カード番号
'        Call .Fields.Append(.CreateField("SCRDM_OTHER_NAME", dbText, 36))    'その他貸出先名称
'        Call .Fields.Append(.CreateField("SCRDM_ROOM_NO", dbText, 6))        '部屋番号
'        Call .Fields.Append(.CreateField("SCRDM_USAGE", dbText, 2))          '使用用途
'        Call .Fields.Append(.CreateField("SCRDM_LENDING_DAY", dbDate))       '貸出日
'        Call .Fields.Append(.CreateField("SCRDM_RETURN_DAY", dbDate))        '返却日
'        Call .Fields.Append(.CreateField("SCRDM_STOP_DAY", dbDate))          '停止日
'        Call .Fields.Append(.CreateField("SCRDM_NOTE", dbText, 60))          '備考
'        Call .Fields.Append(.CreateField("CHK_LENDING_DAY", dbInteger))      '貸出日
'        Call .Fields.Append(.CreateField("CHK_RETURN_DAY", dbInteger))       '返却日
'        Call .Fields.Append(.CreateField("CHK_STOP_DAY", dbInteger))         '停止日
'        Call .Fields.Append(.CreateField("SCRDM_ACPTNO", dbText, 11))        '貸出先契約
'        Call .Fields.Append(.CreateField("CHK_CHANGE", dbInteger))           '変更可否
'        Call .Fields.Append(.CreateField("MOTO_ACPTNO", dbText, 11))         '元々貸出先契約
'        Call .Fields.Append(.CreateField("MOTO_ROOM_NO", dbText, 6))         '元々部屋番号
'        Call .Fields.Append(.CreateField("MOTO_USAGE", dbText, 2))           '元々使用用途
'
'
'      '項目にNULL値の設定許可をする
'        For idx = 0 To .Fields.Count - 1
'            If .Fields(idx).Type = dbText Then
'                .Fields(idx).AllowZeroLength = True
'            End If
'        Next idx
'
'    End With
'errlog = errlog & "⑤:"
'    Call pobjMY_DB.TableDefs.Append(tdfNew)
'    Me.SecurityCardsList.Visible = False
'errlog = errlog & "⑥:"
'
'    clearFKS600_W01 = True
'
'    GoTo Finally
'
'Exception:
'    isException = True
'
'Finally:
'    If isConnect = True Then
'errlog = errlog & "⑦:"
'        If Not pobjMY_DB Is Nothing Then pobjMY_DB.Close: Set pobjMY_DB = Nothing
'    End If
'errlog = errlog & "⑧:"
'    If Not tdfNew Is Nothing Then Set tdfNew = Nothing
'errlog = errlog & "⑨:"
'
'    Me.SecurityCardsList.SourceObject = P_SUB_FORM
'errlog = errlog & "⑩:"
'
'    If isException = True Then
'Call Err.Raise(Err.Number, "clearFKS600_W01:" & errlog & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
''        Call Err.Raise(Err.Number, "clearFKS600_W01" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'    End If
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :tableExist
'        機能             :ACCESSテーブル存在チェック
'        IN               :テーブル名
'        OUT              :True=存在する False=存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function TableExist(dbAccess As Database, strTableName As String) As Boolean

    Dim tdf As TableDef
    
    TableExist = False
    
    For Each tdf In dbAccess.TableDefs
        If tdf.NAME = strTableName Then
            TableExist = True
            Exit For
        End If
    Next tdf

End Function
'==============================================================================*
'
'        MODULE_NAME      :カードメンテテーブルへの書込
'        MODULE_ID        :writeSCRDtable
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub writeSCRDtable(ByRef aエラー文言 As String)
    
    Dim transObject  As Workspace
    Dim isException  As Boolean
    Dim dateText     As String
    Dim timeText     As String
    Dim userName     As String
   
    On Error GoTo Exception
    isException = False
    
    ' 追加及び更新情報の取得
    dateText = Format$(DATE, "yyyymmdd")
    timeText = Format$(time, "hhmmss")
    userName = MSZZ000.LsGetUserName
    
    'トランザクション開始
    Set transObject = DBEngine(0)
    transObject.BeginTrans

    ' メンテ系かカード選択系かで更新手順が異なる
    If IsSelectModeFKS600() = False Then
        'メンテ系(メニューから & ヤードから)
        Call writeSCRDforMainte(Me.cmb_Yard, dateText, timeText, userName)
    Else
        '選択系(受付入力から & 契約入力から)
        Call writeSCRDforSelect(pParamArgs.YARDC, pParamArgs.ROOMC, pParamArgs.ACPTNO, _
                                dateText, timeText, userName, aエラー文言)
                
        ' 追加カード代に変更があれば受付トランに書き込む
        If Nz(p元AddCardPrice, "") <> Nz(Me.txt_カード代.VALUE, "") Then
            Call writeAddCardPrice(pParamArgs.ACPTNO, dateText, timeText, userName)
            isNoticeforFVS400 = True
        End If
    End If
    
    'コミット完了
    transObject.CommitTrans
    
    GoTo Finally

Exception:
    transObject.Rollback
    isException = True

Finally:
    If Not transObject Is Nothing Then Set transObject = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "writeSCRDtable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :メンテ用のカードメンテ書込
'        MODULE_ID        :writeSCRDforMainte
'        NOTE             :このサブはメニューかヤードから来たメンテモード時専用
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub writeSCRDforMainte(ByVal aヤードコード As String, _
                               ByVal a更新日 As String, _
                               ByVal a更新時 As String, _
                               ByVal aユーザID As String)

    Dim rsAccess      As Recordset  ' ワークテーブルレコードセット
    Dim rsSqlServer   As Recordset  ' ＳＱＬサーバテーブルのレコードセット
    Dim sqlText       As String
    Dim serverCardNo  As String
    Dim accessCardNo  As String
    Dim workDatas()   As Variant    ' ワークテーブルデータ格納配列(セキュリティカードマスタ)
    Dim workCount     As Integer    ' ワークテーブルのデータ数
    Dim workIdx       As Integer    ' 格納配列の添字
    Dim isException   As Boolean
    
    On Error GoTo Exception
    isException = False
    
    ' カードマスタから更新対象に成りうるレコードを取得する
    sqlText = " SELECT * FROM SCRD_MAST WHERE SCRDM_YCODE = " & aヤードコード
    sqlText = sqlText & " ORDER BY SCRDM_CARD_NO"
    Set rsSqlServer = pobjKONT_DB.OpenRecordset(sqlText, dbOpenDynaset)

    ' ワークテーブルのデータ抽出
    ' (※画面上で削除されたセキュリティマスタデータのカード番号は｢0｣で残る)
    sqlText = " SELECT * FROM FKS600_W01 ORDER BY SCRDM_CARD_NO "
    Set rsAccess = pobjMY_DB.OpenRecordset(sqlText, dbReadOnly)

    ' 対象データ数分処理を行う
    accessCardNo = ""
    serverCardNo = ""
    workCount = 0
    Do Until (rsAccess.EOF And rsSqlServer.EOF)
        ' ワークテーブルのキー作成
        If rsAccess.EOF = True Then
            accessCardNo = MAX_CARD_NO
        Else
            accessCardNo = rsAccess.Fields("SCRDM_CARD_NO")
        End If

        ' セキュリティマスタのキー作成
        If rsSqlServer.EOF = True Then
            serverCardNo = MAX_CARD_NO
        Else
            serverCardNo = rsSqlServer.Fields("SCRDM_CARD_NO")
        End If

        ' ワークテーブルとセキュリティマスタのデータ比較
        Select Case StrComp(accessCardNo, serverCardNo, vbBinaryCompare)
            Case 0  'accessCardNo = serverCardNo
                ' ワークテーブルのデータとセキュリティマスタのデータが一致する場合
                ' 修正があった場合のみ更新をする
                If rsAccess.Fields("CHK_CHANGE") = -1 Then
                    With rsSqlServer
                        .Edit
                        
                        ' 顧客貸出中で返却日が設定されていれば返却更新を行う
                        If Nz(rsAccess.Fields("SCRDM_ACPTNO"), "") <> "" And _
                           rsAccess.Fields("CHK_RETURN_DAY") = -1 Then
                            Call setRetunCardRecord(rsAccess, rsSqlServer)
                        Else
                        ' 通常(貸出中/未貸出)更新
                            Call setMainteCardRecord(rsAccess, rsSqlServer)
                        End If
                        .Fields("SCRDM_NOTE") = rsAccess.Fields("SCRDM_NOTE")
                        .Fields("SCRDM_UPDAD") = a更新日
                        .Fields("SCRDM_UPDAJ") = a更新時
                        .Fields("SCRDM_UPDPB") = PROG_ID
                        .Fields("SCRDM_UPDUB") = aユーザID
                        .UPDATE
                    End With
                End If
                ' → 次レコードの処理を行う
                rsAccess.MoveNext     ' ワークテーブル
                rsSqlServer.MoveNext  ' セキュリティマスタ

            Case -1  'accessCardNo < serverCardNo
                ' ワークテーブルのデータがセキュリティマスタに存在しない場合
                ' → ワークテーブルのデータをセキュリティマスタに登録
    
                ' 配列にワークテーブルのデータをセット
                ReDim Preserve workDatas(8, workCount)
                workDatas(0, workCount) = aヤードコード
                workDatas(1, workCount) = accessCardNo
                workDatas(2, workCount) = rsAccess.Fields("SCRDM_ROOM_NO")
                workDatas(3, workCount) = rsAccess.Fields("SCRDM_USAGE")
                If rsAccess.Fields("SCRDM_USAGE") = "03" Then
                    workDatas(4, workCount) = rsAccess.Fields("SCRDM_OTHER_NAME")
                Else
                    workDatas(4, workCount) = Null
                End If
                workDatas(5, workCount) = rsAccess.Fields("SCRDM_LENDING_DAY")
                workDatas(6, workCount) = rsAccess.Fields("SCRDM_RETURN_DAY")
                workDatas(7, workCount) = rsAccess.Fields("SCRDM_STOP_DAY")
                workDatas(8, workCount) = rsAccess.Fields("SCRDM_NOTE")
    
                workCount = workCount + 1  ' 登録用ワークテーブルデータ数カウントアップ
                ' → 次レコードの処理を行う
                rsAccess.MoveNext    ' ワークテーブルのみ
  
            Case 1 ' accessCardNo > serverCardNo
              ' セキュリティマスタのデータがワークテーブルに存在しない場合
               ' → セキュリティマスタのデータを削除
               With rsSqlServer
                   .Edit
                   .Delete
               End With
               ' → 次レコードの処理を行う
               rsSqlServer.MoveNext   ' セキュリティマスタのみ
        End Select
    Loop
'    rsAccess.Close
'    Set rsAccess = Nothing

    ' ワークテーブルのデータ作成
    If workCount > 0 Then
        For workIdx = 0 To workCount - 1
            ' 新規追加のデータをセキュリティマスタに追加登録
            Call insertSCRD_MAST(rsSqlServer, workDatas(), workIdx, a更新日, a更新時, aユーザID)
        Next workIdx
    End If

    GoTo Finally

Exception:
    isException = True

Finally:
    If workCount > 0 Then Erase workDatas
    If Not rsAccess Is Nothing Then rsAccess.Close: Set rsAccess = Nothing
    If Not rsSqlServer Is Nothing Then rsSqlServer.Close: Set rsSqlServer = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "writeSCRDforMainte" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :カードマスタへの新規追加
'        MODULE_ID        :insertSCRD_MAST
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub insertSCRD_MAST(ByRef aRecoed As Recordset, _
                            ByVal aWorkDatas As Variant, _
                            ByVal aWorkIndex As Integer, _
                            ByVal a更新日 As String, _
                            ByVal a更新時 As String, _
                            ByVal aユーザID As String)
    
    On Error GoTo Exception
    
    With aRecoed
        .AddNew

        .Fields("SCRDM_YCODE") = CLng(aWorkDatas(0, aWorkIndex))
        .Fields("SCRDM_CARD_NO") = aWorkDatas(1, aWorkIndex)
        .Fields("SCRDM_ROOM_NO") = convertRoomNo(aWorkDatas(2, aWorkIndex))
        .Fields("SCRDM_RETURN_RNO") = .Fields("SCRDM_ROOM_NO") ' 新規なら部屋番号と同じ値を設定する
        .Fields("SCRDM_USAGE") = aWorkDatas(3, aWorkIndex)
        .Fields("SCRDM_OTHER_NAME") = aWorkDatas(4, aWorkIndex)
        If Nz(aWorkDatas(5, aWorkIndex), "") <> "" Then
            .Fields("SCRDM_LENDING_DAY") = aWorkDatas(5, aWorkIndex)
        End If
        If Nz(aWorkDatas(6, aWorkIndex), "") <> "" Then
            .Fields("SCRDM_RETURN_DAY") = aWorkDatas(6, aWorkIndex)
        End If
        If Nz(aWorkDatas(7, aWorkIndex), "") <> "" Then
            .Fields("SCRDM_STOP_DAY") = aWorkDatas(7, aWorkIndex)
        End If
        .Fields("SCRDM_NOTE") = aWorkDatas(8, aWorkIndex)
        .Fields("SCRDM_INSED") = a更新日
        .Fields("SCRDM_INSEJ") = a更新時
        .Fields("SCRDM_INSPB") = PROG_ID
        .Fields("SCRDM_INSUB") = aユーザID

        .UPDATE
    End With

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "insertSCRD_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :選択用のカードメンテ書込
'        MODULE_ID        :writeSCRDforSelect
'        NOTE             :このサブは受付か契約から来たメンテモード時専用
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub writeSCRDforSelect(ByVal aヤードコード As String, _
                               ByVal a部屋番号 As String, _
                               ByVal a契約番号 As String, _
                               ByVal a更新日 As String, _
                               ByVal a更新時 As String, _
                               ByVal aユーザID As String, _
                               ByRef aエラー文言 As String)

    Dim rsAccess      As Recordset  ' ワークテーブルレコードセット
    Dim rsSqlServer   As Recordset  ' ＳＱＬサーバテーブルのレコードセット
    Dim sqlText       As String
    Dim serverCardNo  As String
    Dim accessCardNo  As String
    Dim isException   As Boolean
    Dim whereCards    As String     ' 元テーブルから抽出の為のカード達
    Dim errorMessage  As String
    
    On Error GoTo Exception
    isException = False
    errorMessage = ""
    
    ' 呼出元に反映させる為の情報初期化
    pUsedCardRowSource = ""
    pBaseUsedCardCount = 0
    pAddUsedCardCount = 0
    
    ' ワークテーブルのデータ抽出
    sqlText = " SELECT * FROM FKS600_W01 WHERE CHK_CHANGE = -1 ORDER BY SCRDM_CARD_NO"
    Set rsAccess = pobjMY_DB.OpenRecordset(sqlText, dbReadOnly)
    If rsAccess.EOF = True Then GoTo Finally
    
    
    ' 更新カードの抽出の為の IN(....)ステートメントを作る
    ' INステートメント使うよりは、ヤードで絞るか１件づつ処理を行うようにしたほうがパフォーマンスが良いハズ
    whereCards = ""
    Do While Not rsAccess.EOF
        whereCards = whereCards & "'" & rsAccess.Fields("SCRDM_CARD_NO") & "',"
        rsAccess.MoveNext
    Loop
    rsAccess.MoveFirst
    
    If whereCards = "" Then
        Call MSZZ024_M10("pobjMY_DB.OpenRecordset", "ワークテーブルの更新データが存在しない。")
    End If
    whereCards = Left(whereCards, Len(whereCards) - 1) '末尾の "'" を抜く
    
    ' カードマスタから更新対象に成りうるレコードを取得する
    sqlText = " SELECT * FROM SCRD_MAST WHERE SCRDM_YCODE = " & aヤードコード
    sqlText = sqlText & " AND SCRDM_CARD_NO IN(" & whereCards & ")"
    sqlText = sqlText & " ORDER BY SCRDM_CARD_NO"
    Set rsSqlServer = pobjKONT_DB.OpenRecordset(sqlText, dbOpenDynaset)


    ' 対象データ数分処理を行う
    Do Until (rsAccess.EOF And rsSqlServer.EOF)
        ' ワークテーブルのキー作成
        If rsAccess.EOF = True Then
            accessCardNo = MAX_CARD_NO
        Else
            accessCardNo = rsAccess.Fields("SCRDM_CARD_NO")
        End If

        ' セキュリティマスタのキー作成
        If rsSqlServer.EOF = True Then
            serverCardNo = MAX_CARD_NO
        Else
            serverCardNo = rsSqlServer.Fields("SCRDM_CARD_NO")
        End If

        ' ワークテーブルとセキュリティマスタのデータ比較
        Select Case StrComp(accessCardNo, serverCardNo, vbBinaryCompare)
            Case 0  'accessCardNo = serverCardNo
            ' ワークテーブルのデータとセキュリティマスタのデータが一致する場合
                With rsSqlServer
                    If .Fields("SCRDM_ROOM_NO") = P_部屋割当無し値 Or _
                      (.Fields("SCRDM_ROOM_NO") = CLng(a部屋番号)) Then
                    ' 部屋未割当てか部屋割当て済みなら更新可
                        .Edit
                        
                        ' 返却日が設定されていれば返却更新を行う
                        If rsAccess.Fields("CHK_RETURN_DAY") = -1 Then
                            Call setRetunCardRecord(rsAccess, rsSqlServer)
                        Else
                        ' 通常(貸出中)更新
                            Call setSelectCardRecord(rsAccess, rsSqlServer)
                        End If
                        '↓残りは共通更新内容
                        .Fields("SCRDM_NOTE") = rsAccess.Fields("SCRDM_NOTE")
                        .Fields("SCRDM_UPDAD") = a更新日
                        .Fields("SCRDM_UPDAJ") = a更新時
                        .Fields("SCRDM_UPDPB") = PROG_ID
                        .Fields("SCRDM_UPDUB") = aユーザID
                        .UPDATE
                        
                    ElseIf .Fields("SCRDM_ROOM_NO") <> P_部屋割当無し値 And _
                            rsAccess.Fields("SCRDM_ROOM_NO") = P_部屋割当無し値 Then
                    ' 未割当てとしたが、更新時には別の部屋が割り当たっていた
                    ' →こちらの更新権限は無い特に何もしない
                    
                    Else
                    ' 使おうと思ったカードが既に別の部屋に割付いている
                        errorMessage = errorMessage & _
                            "○カードNo.【" & rsAccess.Fields("SCRDM_CARD_NO") & "】は先に別の部屋【" & _
                            Format$(.Fields("SCRDM_ROOM_NO"), "000000") & "】に使われてしまいました。" & vbCr
                    ' 使えなかったカードの情報は記録する
                    End If
                End With
                ' → 次レコードの処理を行う
                rsAccess.MoveNext     ' ワークテーブル
                rsSqlServer.MoveNext  ' セキュリティマスタ

            Case -1  'accessCardNo < serverCardNo
                ' ワークテーブルのデータがセキュリティマスタに存在しない場合
                ' → カード選択中にメンテで消してしまった！
                errorMessage = errorMessage & _
                            "●カードNo.【" & rsAccess.Fields("SCRDM_CARD_NO") & "】は別のメンテ機能により削除されてしまってます。" & vbCr
                ' → 次レコードの処理を行う
                rsAccess.MoveNext    ' ワークテーブルのみ
  
            Case 1 ' accessCardNo > serverCardNo
                ' セキュリティマスタのデータがワークテーブルに存在しない場合
                ' → 絶対ありえない！
                Call MSZZ024_M10("強制停止", "セキュリティマスタのデータがワークテーブルにない")
                ' → 次レコードの処理を行う
                rsSqlServer.MoveNext   ' セキュリティマスタのみ
        
        End Select
    Loop
'    rsAccess.Close
'    Set rsAccess = Nothing
    
    aエラー文言 = errorMessage
    GoTo Finally

Exception:
    isException = True

Finally:
    If Not rsAccess Is Nothing Then rsAccess.Close: Set rsAccess = Nothing
    If Not rsSqlServer Is Nothing Then rsSqlServer.Close: Set rsSqlServer = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "writeSCRDforSelect" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :追加カード代の記入
'        MODULE_ID        :writeAddCardPrice
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub writeAddCardPrice(ByVal a契約番号 As String, _
                              ByVal a更新日 As String, _
                              ByVal a更新時 As String, _
                              ByVal aユーザID As String)

    Dim rsSqlServer   As Recordset  ' ＳＱＬサーバテーブルのレコードセット
    Dim sqlText       As String
    Dim isException   As Boolean
    
  On Error GoTo Exception
    isException = False
    
    ' 書き込む受付トランの抽出
    sqlText = " SELECT * FROM RCPT_TRAN WHERE RCPT_CARG_ACPTNO = '" & a契約番号 & "'"
    Set rsSqlServer = pobjKONT_DB.OpenRecordset(sqlText, dbOpenDynaset)
    
    With rsSqlServer
        If Not .EOF Then
            .Edit
            If Nz(Me.txt_カード代, "") = "" Then
            ' 追加カード代に何も入力されていない
                '2015/03/25 M.HONDA INS
                If .Fields("RCPT_ADD_FZAPPI_CODE2") <> "96" Then
                    .Fields("RCPT_ADD_FZAPPI_CODE2") = Null
                    .Fields("RCPT_ADD_FZAPPI2") = Null
                End If
            Else
            ' 追加カード代に何か入力
                .Fields("RCPT_ADD_FZAPPI_CODE2") = P_雑費区分_追加カード代値
                .Fields("RCPT_ADD_FZAPPI2") = zeroByteToNull(Me.txt_カード代)
            End If
            .Fields("RCPT_UPDAD") = a更新日
            .Fields("RCPT_UPDAJ") = a更新時
            .Fields("RCPT_UPDPB") = PROG_ID
            .Fields("RCPT_UPDUB") = aユーザID
            .UPDATE
        End If
    End With

    GoTo Finally

Exception:
    isException = True

Finally:
    If Not rsSqlServer Is Nothing Then rsSqlServer.Close: Set rsSqlServer = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "writeAddCardPrice" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :カードの返却更新
'        MODULE_ID        :setRetunCardRecord
'        NOTE             :返却状態（返却日入力）した場合そのカードを未貸出に戻す
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setRetunCardRecord(ByRef anAccess As Recordset, _
                               ByRef aSqlServer As Recordset)
                               
     Dim roomNo As Long
     
     On Error GoTo Exception
     
     roomNo = convertRoomNo(anAccess.Fields("SCRDM_ROOM_NO")) '部屋番変換
     
    '用途が基本ならば基本で、基本以外ならば予備を設定
    If anAccess.Fields("SCRDM_USAGE") = "01" Then
        '基本としての登録
        aSqlServer.Fields("SCRDM_USAGE") = "01"
        aSqlServer.Fields("SCRDM_ROOM_NO") = roomNo
        aSqlServer.Fields("SCRDM_RETURN_RNO") = roomNo
    Else
        '予備としての登録
        aSqlServer.Fields("SCRDM_USAGE") = "00"
        aSqlServer.Fields("SCRDM_ROOM_NO") = aSqlServer.Fields("SCRDM_RETURN_RNO")
        ' aSqlServer.Fields("SCRDM_RETURN_RNO")←返却元は変えない
    End If
    
    '↓残りは全て初期化
    aSqlServer.Fields("SCRDM_ACPTNO") = Null
    aSqlServer.Fields("SCRDM_OTHER_NAME") = Null
    aSqlServer.Fields("SCRDM_LENDING_DAY") = Null
    aSqlServer.Fields("SCRDM_RETURN_DAY") = Null
    aSqlServer.Fields("SCRDM_STOP_DAY") = Null
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "setRetunCardRecord" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :選択カードの通常更新
'        MODULE_ID        :setSelectCardRecord
'        NOTE             :選択モードでの返却状態（返却日入力）以外のカードの更新
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setSelectCardRecord(ByRef anAccess As Recordset, _
                               ByRef aSqlServer As Recordset)
                               
    On Error GoTo Exception
     
    If anAccess.Fields("SCRDM_USAGE") = CmKomsMod.P_カード用途_基本値 Then
        ' 基本用途ならば返却先部屋番は部屋番号
        aSqlServer.Fields("SCRDM_RETURN_RNO") = anAccess.Fields("SCRDM_ROOM_NO")
        ' 追加用途ならば返却先部屋番そのまま変えない
    End If
    aSqlServer.Fields("SCRDM_ROOM_NO") = convertRoomNo(anAccess.Fields("SCRDM_ROOM_NO"))
    aSqlServer.Fields("SCRDM_USAGE") = anAccess.Fields("SCRDM_USAGE")
    aSqlServer.Fields("SCRDM_ACPTNO") = anAccess.Fields("SCRDM_ACPTNO")
    aSqlServer.Fields("SCRDM_LENDING_DAY") = anAccess.Fields("SCRDM_LENDING_DAY")
    aSqlServer.Fields("SCRDM_RETURN_DAY") = anAccess.Fields("SCRDM_RETURN_DAY")
    aSqlServer.Fields("SCRDM_STOP_DAY") = anAccess.Fields("SCRDM_STOP_DAY")
     
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "setSelectCardRecord" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :メンテカードの通常更新
'        MODULE_ID        :setMainteCardRecord
'        NOTE             :メンテモードでの返却状態（返却日入力）以外のカードの更新
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setMainteCardRecord(ByRef anAccess As Recordset, _
                               ByRef aSqlServer As Recordset)
                               
    On Error GoTo Exception
     
    aSqlServer.Fields("SCRDM_ROOM_NO") = convertRoomNo(anAccess.Fields("SCRDM_ROOM_NO"))
    aSqlServer.Fields("SCRDM_RETURN_RNO") = convertRoomNo(anAccess.Fields("SCRDM_ROOM_NO"))
    aSqlServer.Fields("SCRDM_USAGE") = anAccess.Fields("SCRDM_USAGE")
    aSqlServer.Fields("SCRDM_ACPTNO") = anAccess.Fields("SCRDM_ACPTNO")
    aSqlServer.Fields("SCRDM_LENDING_DAY") = anAccess.Fields("SCRDM_LENDING_DAY")
    aSqlServer.Fields("SCRDM_RETURN_DAY") = anAccess.Fields("SCRDM_RETURN_DAY")
    aSqlServer.Fields("SCRDM_STOP_DAY") = anAccess.Fields("SCRDM_STOP_DAY")
     
    ' 業者貸しとしたら業者名をセット
    If anAccess.Fields("SCRDM_USAGE") = "03" Then
        aSqlServer.Fields("SCRDM_OTHER_NAME") = anAccess.Fields("SCRDM_OTHER_NAME")
    Else
        aSqlServer.Fields("SCRDM_OTHER_NAME") = Null
    End If
     
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "setMainteCardRecord" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :顧客名取得サブクエリの取得
'        MODULE_ID        :getUserNameSubQuery
'        NOTE             :
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getUserNameSubQuery() As String

    Dim sqlText As String
    
    sqlText = "( SELECT USER_NAME FROM USER_MAST "
    ' 遷移元によって顧客名取得元のトランを変えないといけない
    ' 誠に遺憾ではあるが過去の契約には受付トランがない場合もありえるので
    If pParamArgs.FROM_ID = P_fromFVS400 Then
        sqlText = sqlText & " INNER JOIN RCPT_TRAN ON USER_CODE = RCPT_UCODE "
        sqlText = sqlText & " WHERE RCPT_CARG_ACPTNO = '"
    Else
        sqlText = sqlText & " INNER JOIN CARG_FILE ON USER_CODE = CARG_UCODE "
        sqlText = sqlText & " WHERE CARG_ACPTNO = '"
    End If
    sqlText = sqlText & pParamArgs.ACPTNO & "') AS USER_NAME "

    getUserNameSubQuery = sqlText

End Function

'==============================================================================*
'
'        MODULE_NAME      :追加カード代取得サブクエリの取得
'        MODULE_ID        :getAddCardPriceSubQuery
'        NOTE             :受付から来たときのみ実施
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getAddCardPriceSubQuery() As String

    Dim sqlText As String
    
    sqlText = "( SELECT RCPT_ADD_FZAPPI2 FROM RCPT_TRAN "
    sqlText = sqlText & " WHERE RCPT_CARG_ACPTNO = '" & pParamArgs.ACPTNO
    sqlText = sqlText & "') AS CARD_PRICE "

    getAddCardPriceSubQuery = sqlText

End Function

'==============================================================================*
'
'        MODULE_NAME      :部屋番号の変換
'        MODULE_ID        :convertRoomNo
'        NOTE             :ワークテーブルの部屋は文字列で且つ、"無し"があるので
'                         :それを-1に変換する
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function convertRoomNo(ByVal a部屋番号 As String) As Long

    Dim returnValue As Long
    
    returnValue = P_部屋割当無し値

    If a部屋番号 <> P_部屋割当無し Then returnValue = CLng(a部屋番号)

    convertRoomNo = returnValue

End Function
'==============================================================================*
'   外部関数
'==============================================================================*
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function CmbCheck_Num(ByVal oText As Variant, ByVal ilen As Integer, Optional ByVal isLEN As Variant) As Integer
    '********** 入力チェック処理 **********
    Dim vINT    As Variant
    Dim vsINT   As Variant
    Dim i       As Integer
    
    Dim dTEXT   As Double
    Dim dFROM   As Double
    Dim dTO     As Double
    Dim vRet    As Variant
    
    On Error GoTo err_rtn
    
    CmbCheck_Num = 9
    
    If Nz(CmNumeric(oText), "") = "" Then       '数値と判断できないのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If
    
    If CmChkCher(CStr(Nz(oText)), "A") = False Then '全角文字が入力されているのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If
    
    '小数点の値が設定されている時のチェック
    vRet = Nz(InStr(Abs(oText), "."), 0)
    
    If IsMissing(isLEN) = False Then
        If vRet <> 0 Then
            If isLEN < CmLenb(Abs(oText)) - vRet Then    '小数点桁数が設定桁数を超えているエラー
              CmbCheck_Num = 2
              GoTo Exit_rtn
            End If
        End If
    Else
        If CmLenb(Abs(CmNumeric(oText))) > ilen Then           '文字数超過のためエラー
              CmbCheck_Num = 2
              GoTo Exit_rtn
        End If
        If vRet <> 0 Then                         '小数点の入力があるのでエラー
          CmbCheck_Num = 2
          GoTo Exit_rtn
        End If
    End If

    '桁チェック
    i = 0
    For i = 1 To ilen: vINT = vINT & "9": Next
    
    If IsMissing(isLEN) = False Then
        vsINT = "."
        i = 0
        For i = 1 To isLEN
            vsINT = vsINT & "9"
        Next
    End If
    
    vINT = vINT & vsINT
    vINT = CmNumeric(vINT)
    
    dTEXT = CmNumeric(oText)
    dFROM = vINT * -1
    dTO = vINT
    
    If dFROM > dTEXT Or dTO < dTEXT Then        '論理範囲超過のためエラー
        CmbCheck_Num = 2
        GoTo Exit_rtn
    End If
    
    CmbCheck_Num = 0
    
Exit_rtn:
    Exit Function
    
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    CmbCheck_Num = 9
    GoTo Exit_rtn
End Function
'==============================================================================*
'
'       MODULE_NAME     : キー押下処理
'       MODULE_ID       : KeyDown
'       CREATE_DATE     : 2006/06/01
'       PARAM           : KeyCode           キーコード
'                       : Shift             シフトキー状態
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    'F1:戻る
    Select Case KeyCode
    Case vbKeyF1
        KeyCode = 0
        Me.cmd_End.SetFocus
    Case vbKeyF2
        KeyCode = 0
        Me.cmd_Enter.SetFocus
    Case vbKeyF3, vbKeyF5 To vbKeyF12
        KeyCode = 0
    End Select
End Sub
'==============================================================================*
'
'        MODULE_NAME      :カードメンテが選択モード有無
'        MODULE_ID        :isSelectModeFKS600
'        NOTE             :カードメンテが選択モードならばTrue,メンテモードならFalse
'        CREATE_DATE      :2006/06/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function IsSelectModeFKS600(Optional aFROM_ID As String = "") As Boolean
        
    Dim fromID As String
    
    ' 引数がなければグローバル、あればその値で評価準備ス
    If aFROM_ID = "" Then
        fromID = pParamArgs.FROM_ID
    Else
        fromID = aFROM_ID
    End If
        
    ' 実際のチェック、受付と契約から来たならば選択モード、それ以外は編集モード
    '2006/07/28 "P_fromFVS601" をAdd
    If fromID = P_fromFVS400 Or _
       fromID = P_fromFVS601 Or _
       fromID = P_fromFVS600 Then
       
        IsSelectModeFKS600 = True
    Else
        IsSelectModeFKS600 = False
    End If
End Function
'==============================================================================*
'
'        MODULE_NAME      :FKS600の編集変更処理
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetFKS600EditChange()

    If isEditChange = False Then
        Me.Caption = Me.Caption & "(修正中)"
        isEditChange = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :サブフォームの使用用途変更時の処理
'        MODULE_ID        :NoticeChangeUseType
'        NOTE             :選択モードのときのみだがガイダンスを変える必要がある
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub NoticeChangeUseType()

    Call getUseCardNumforWORK 'ガイダンスはちょっと保留
    Call setGuidance

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ガイダンスの設定
'        MODULE_ID        :setGuidance
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub setGuidance()

    Dim guidanceText  As String
    Dim guid基本の件  As String
    Dim guid追加の件  As String
    Dim guid総括      As String

    guidanceText = ""
    ' ガイダンスの生成
    guid基本の件 = "基本貸出は" & pBaseUsedCardCount & "枚、"
    If pAddUsedCardCount = 0 Then
        guid追加の件 = "追加貸出があれば必要枚数分、貸出をチェックして下さい。"
    Else
        guid追加の件 = "追加貸出は " & pAddUsedCardCount & "枚 "
    End If
    guid総括 = "現在合計" & (pBaseUsedCardCount + pAddUsedCardCount) & "枚選択しています。"
    guidanceText = guid基本の件 & guid追加の件 & guid総括
    
    lbl_ガイダンス.Caption = guidanceText

End Sub
'==============================================================================*
'
'        MODULE_NAME      :ゼロバイトはNULLで返す
'        MODULE_ID        :zeroByteToNull
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function zeroByteToNull(varInput As Variant) As Variant

    zeroByteToNull = varInput
    
    If varInput = "" Then
        zeroByteToNull = Null
    End If
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :現在の使用枚数を取得
'        MODULE_ID        :getUseCardNumforWORK
'        NOTE             :ワーク上での基本と追加の使用枚数をグローバルに取得
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub getUseCardNumforWORK()

    Dim rsAccess      As Recordset
    Dim isException   As Boolean
    Dim isConnect     As Boolean
    Dim guidanceText  As String
    Dim guid基本の件  As String
    Dim guid追加の件  As String
    Dim guid総括      As String
    Dim sqlText       As String
        
  On Error GoTo Exception
    
    isException = False
    isConnect = False
    
    If pobjMY_DB Is Nothing Then
        Set pobjMY_DB = CurrentDb
        isConnect = True
    End If
    
   ' ワークテーブルの基本と追加の件数抽出
    sqlText = "SELECT COUNT(SCRDM_USAGE) AS BASE_CNT FROM " & P_WORK_TABLE & " WHERE SCRDM_USAGE = '" _
                      & CmKomsMod.P_カード用途_基本値 & "'" _
                      & " AND SCRDM_LENDING_DAY IS NOT NULL " _
                      & " AND SCRDM_RETURN_DAY IS NULL " _
                      & " AND SCRDM_STOP_DAY IS NULL "
    Set rsAccess = pobjMY_DB.OpenRecordset(sqlText, dbReadOnly)
    If rsAccess.EOF = True Then
        pBaseUsedCardCount = 0
    Else
        pBaseUsedCardCount = Nz(rsAccess.Fields("BASE_CNT"), 0)
     End If
    
    sqlText = "SELECT COUNT(SCRDM_USAGE) AS ADD_CNT FROM " & P_WORK_TABLE & " WHERE SCRDM_USAGE = '" _
                          & CmKomsMod.P_カード用途_追加値 & "'" _
                          & " AND SCRDM_LENDING_DAY IS NOT NULL " _
                          & " AND SCRDM_RETURN_DAY IS NULL " _
                          & " AND SCRDM_STOP_DAY IS NULL "
    Set rsAccess = pobjMY_DB.OpenRecordset(sqlText, dbReadOnly)
    If rsAccess.EOF = True Then
        pAddUsedCardCount = 0
    Else
        pAddUsedCardCount = Nz(rsAccess.Fields("ADD_CNT"), 0)
    End If
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If isConnect = True Then
        If Not pobjMY_DB Is Nothing Then pobjMY_DB.Close: Set pobjMY_DB = Nothing
    End If
    
    If isException = True Then
        Call Err.Raise(Err.Number, "getUseCardNumforWORK" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :サブフォームの初期化
'        MODULE_ID        :initSubForm
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub initSubForm()
      
    On Error GoTo Exception
      ' 選択モードのときのみの設定
      If IsSelectModeFKS600() = True Then
            ' サブフォームの指定箇所にフォーカスを移す
            Me.SecurityCardsList.SetFocus
            Me.SecurityCardsList.Form.chk_LendingDay.SetFocus
            ' サブフォームのプロパティ指定
            ' 読込時の使用枚数を求めておく ※追加枚数があれば追加カード代の入力制御に使う
            Call getUseCardNumforWORK
            ' 使用枚数を元にガイダンスも設定
            ' Call setGuidance
      End If

      'メンテ／選択モード共通の設定
      Call Me.SecurityCardsList.Form.SetBaseProperty
      Me.SecurityCardsList.Requery
      Exit Sub

Exception:
      Call Err.Raise(Err.Number, "initSubForm" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

' 2007/02/08 add new method
'==============================================================================*
'
'        MODULE_NAME      :cmd_依頼書印刷_Click
'        MODULE_ID        :cmd_依頼書印刷_Click
'        NOTE             :
'        CREATE_DATE      :2007/02/08
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_依頼書印刷_Click()
    Dim date最終停止日 As Variant
    Dim is印刷  As Boolean
    
  On Error GoTo Exception
    
    ' ヤードコード未入力では印刷不可
    If Nz(Me.cmb_Yard.VALUE, "") = "" Then Exit Sub
    
    ' 修正中だったときのための、保存処理
    If isEditChange = True Then
    
        If MsgBox("修正中の為、一度保存してから印刷を行います。" & vbCr & "よろしいですか？", _
                    vbQuestion + vbYesNo, P_KAKUNIN) = vbNo Then
            Exit Sub
        End If
        
        '  保存処理実行
        If func保存処理 = False Then Exit Sub
    End If
    
    p停止日FromDate = ""
    p停止日ToDate = ""
    ' 現在編集中画面内ので最終停止日チェック
    date最終停止日 = getWorkLatestStopDay
    If Nz(date最終停止日, "") = "" Then
        MsgBox "現在の画面内にて停止したカードは存在しません。" & vbCr & "よって依頼書は印刷できません", , P_KAKUNIN
        Exit Sub
    End If
    
    ' 印刷範囲の入力
    doCmd.OpenForm "FKS610", , , , , acDialog, date最終停止日 & "," & date最終停止日
    If Nz(p停止日FromDate, "") = "" And Nz(p停止日ToDate, "") = "" Then
        Exit Sub  '両方未入力なら印刷キャンセル
    End If
    
    ' 印刷の開始
    If IsSelectModeFKS600() Then
        ' 部屋指定有り
        is印刷 = Print600.fncPrintSecurityCardRequest(Print600.P600_MODE_PREVIEW, _
                                                  Me.cmb_Yard, _
                                                  p停止日FromDate, p停止日ToDate, _
                                                  pParamArgs.ROOMC)
    Else
        ' 部屋指定無し
        is印刷 = Print600.fncPrintSecurityCardRequest(Print600.P600_MODE_PREVIEW, _
                                                  Me.cmb_Yard, _
                                                  p停止日FromDate, p停止日ToDate)
    End If
    ' 印刷結果の確認
    If is印刷 = False Then
        MsgBox "指定した日付にて停止されたカードは存在しません。", , P_KAKUNIN
    End If
    
    Exit Sub
    
Exception:
        Call MSZZ024_M00("cmd_依頼書印刷_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :func保存処理
'        MODULE_ID        :func保存処理
'        NOTE             :
'        CREATE_DATE      :2007/02/08
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function func保存処理() As Boolean

    Dim errorMessage  As String
    Dim isException As Boolean
    
  On Error GoTo Exception
    isException = False
    
    ' 受付から来た刻のみ必須入力チェック
    If pParamArgs.FROM_ID = P_fromFVS400 Then
        Call getUseCardNumforWORK 'ワークで使用している追加枚数を取得
        ' カード代を入力するのかしてはいけないのかチェック
        If check_txt_カード代(txt_カード代.VALUE) = False Then
            txt_カード代.SetFocus
            func保存処理 = False
            Exit Function
        End If
    End If
        
        
    If Not pobjMY_DB Is Nothing Then pobjMY_DB.Close: Set pobjMY_DB = Nothing
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing
        
    ' 各ＤＢへの接続
    Set pobjMY_DB = CurrentDb
    If fncConnectDB() = False Then
        Call MSZZ024_M10("fncConnectDB", "コンテナDB接続処理が失敗しました。")
    End If
    
    ' テーブルへの更新処理
    Call writeSCRDtable(errorMessage)

    ' 選択系のみの場合の処理
    If IsSelectModeFKS600() Then
        ' 更新失敗のメッセージ　※選択系のみの場合発生しうる
        If errorMessage <> "" Then
            MsgBox "選んだカードで以下のものは更新できませんでした。" & vbCr & errorMessage, vbOKOnly, P_KAKUNIN
        End If
                    
        ' 現状の貸出カード内容を取得しておく
        Call CmKomsMod.GetContractSCards(pobjKONT_DB, pParamArgs.ACPTNO, pBaseUsedCardCount, pAddUsedCardCount, pUsedCardRowSource)
        ' 契約からの場合に呼ぶ外部プロシージャ
        If pParamArgs.FROM_ID = P_fromFVS600 Then
            Call Form_FVS600.SetCardInfomation(pBaseUsedCardCount, pAddUsedCardCount, pUsedCardRowSource)
        ElseIf pParamArgs.FROM_ID = P_fromFVS601 Then
            Call Form_FVS601.SetCardInfomation(pBaseUsedCardCount, pAddUsedCardCount, pUsedCardRowSource)
        ' 受付からの場合に呼ぶ外部プロシージャ
        ElseIf pParamArgs.FROM_ID = P_fromFVS400 Then
            Call Form_FVS400.subSCardRefresh(pBaseUsedCardCount, pAddUsedCardCount, pUsedCardRowSource, isNoticeforFVS400)
        End If
    End If
        
    ' 再描画（初期状態に戻す）
    Call makeFKS600_W01(pParamArgs.FROM_ID, cmb_Yard.VALUE, pParamArgs.ROOMC)
    ' サブフォーム初期設定
    'Call initSubForm 初期設定する必要なし
    
    ' 各制御の初期化
    isEditChange = False
    isNoticeforFVS400 = False
    Me.Caption = pstrFormCaption
    
    func保存処理 = True
    GoTo Finally
    
Exception:
    isException = True
    func保存処理 = False
    
Finally:
    If Not pobjMY_DB Is Nothing Then pobjMY_DB.Close: Set pobjMY_DB = Nothing
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "func保存処理" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :WORKテーブル上の最新停止日を取得する
'        MODULE_ID        :getWorkLatestStopDay
'        NOTE             :
'        CREATE_DATE      :2007/02/08
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getWorkLatestStopDay() As Variant
    Dim isException As Boolean
    Dim rsAccess    As Recordset
    Dim sqlText     As String
  
  On Error GoTo Exception
    isException = False
    Set pobjMY_DB = CurrentDb
    sqlText = "SELECT MAX(SCRDM_STOP_DAY) FROM " & P_WORK_TABLE
    Set rsAccess = pobjMY_DB.OpenRecordset(sqlText, dbReadOnly)
    getWorkLatestStopDay = rsAccess.Fields(0)
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If Not rsAccess Is Nothing Then rsAccess.Close: Set rsAccess = Nothing
    If Not pobjMY_DB Is Nothing Then pobjMY_DB.Close: Set pobjMY_DB = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "getWorkLatestStopDay" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function
'==============================================================================*
'
'        MODULE_NAME      :外部ダイアログが入力した依頼書停止日を設定する
'        MODULE_ID        :subSetStopDay
'        NOTE             :
'        CREATE_DATE      :2007/02/08
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subSetStopDay(ByVal aFromDay As Variant, ByVal aToDay As Variant)
    p停止日FromDate = Nz(aFromDay, "")
    p停止日ToDate = Nz(aToDay, "")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :鍵のヤードを検索
'        MODULE_ID        :cmd_Search_Click
'        NOTE             :
'        CREATE_DATE      :2023/12/04
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Search_Click()

    Dim objCon  As Object
    Dim strSQL  As String
    Dim varYard As Variant

    On Error GoTo ErrorHandler

    Set objCon = ADODB_Connection(DLookup("CONT_BUMOC", "dbo_CONT_MAST"))
    
    On Error GoTo ErrorHandler1
    strSQL = ""
    strSQL = strSQL & "SELECT SCRDM_YCODE from SCRD_MAST "
    strSQL = strSQL & " WHERE SCRDM_CARD_NO = '" & txt_CARD_NO & "'"
    
    varYard = ADODB_ExecGetVariant(strSQL, objCon)
    If Nz(varYard) = "" Then
        MsgBox "見つかりませんでした。", vbExclamation
    Else
        cmb_Yard = varYard
        Call cmb_Yard_AfterUpdate
    End If

ErrorHandler1:
    objCon.Close
    
ErrorHandler:
    If Err <> 0 Then
        Call MSZZ024_M00("cmd_Search_Click", True)
    End If
End Sub
'****************************  ended or program ********************************
