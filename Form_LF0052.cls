VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LF0052"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 未収管理サブシステム
'
'        PROGRAM_NAME    : 宛名印刷
'        PROGRAM_ID      : LF0052
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2001/05/11
'        CERATER         : N.MIURA
'
'        UPDATE          : 2001/08/29
'        UPDATER         : N.MIURA
'
'        UPDATE          : 2001/11/19
'        UPDATER         : N.MIURA
'                        : ＤＢ移行
'
'        UPDATE          : 2002/03/27
'        UPDATER         : N.MIURA
'                        : ワークテーブル名変更
'
'        UPDATE          : 2002/03/28
'        UPDATER         : N.MIURA
'                        : 起動ログ出力機能追加
'
'        UPDATE          : 2002/05/31
'        UPDATER         : N.MIURA
'                        : 標準化
'
'        UPDATE          : 2003/12/10
'        UPDATER         : N.MIURA
'        Ver             : 0.7
'
'        UPDATE          : 2007/02/16
'        UPDATER         : N.MIURA
'        Ver             : 0.8
'                        : 住所３対応
'
'        UPDATE          : 2011/4/15
'        UPDATER         : M.RYU
'        Ver             : 0.9
'                        : クリアボタンを追加
'
'        UPDATE          : 2011/07/11
'        UPDATER         : M.RYU
'        Ver             : 1.0
'                          ワークテーブルWK_MQLY010_TEMP変更⇒WK_LQ0052_TEMP
'                          全選択、全解除、削除　ボタンを追加
'                          宛先印刷を修正⇒選択したレコードのみ印刷
'
'        UPDATE          : 2011/08/03
'        UPDATER         : M.RYU
'        Ver             : 1.1
'                          宛名印刷に書類送付先優先で印刷
'
'        UPDATE          : 2011/09/15
'        UPDATER         : M.RYU
'        Ver             : 1.2
'                          宛名印刷修正 代表取締役を表示
'
'        UPDATE          : 2011/10/27
'        UPDATER         : M.RYU
'        Ver             : 1.3
'                          代表取締役出力修正
'
'        UPDATE          : 2026/04/14
'        UPDATER         : N.IMAI
'        Ver             : 1.4
'                          請求書・宛名の出し方に合わせる
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "MFLY010"
'
Private USER_ID         As String
Private MODL_ID         As String
'
Private dbs             As Database
Private strSQL          As String
'
Private stDocName       As String
Private WK_KEISN        As String
Private WK_DAIHN        As String
'
Private WK_BUMOC        As String
Private WK_BUMON        As String

Private Const WK_TEMP   As String = "WK_LQ0052_TEMP"
Private Const WK_TEMP1   As String = "WK_LQ0052_TEMP1"
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
    
    Call MSHS010_M00(PROG_ID, "1")
    Call MFLY010_DBOPEN
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル6.BackColor = RGB(255, 0, 204)
    End Select
    
    USER_ID = LsGetUserName()
    WK_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1")
    WK_BUMON = DLookup("BUMOM_BUMON", "BUMO_MAST", "BUMOM_BUMOC = '" & WK_BUMOC & "'")      'INSERT 2011/07/11 M.RYU
    
Exit_Form_Open:
    Exit Sub

Err_Form_Open:
    MsgBox Err.Description
    Resume Exit_Form_Open
End Sub
'==============================================================================*
'
'        MODULE_NAME      : 全解除
'        MODULE_ID        : MFLY010_ALLCLE_Click
'        CREATE_DATE      : 2011/07/11 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_ALLCLE_Click()
On Error GoTo Err_Handler
    Dim strSQL As String
        
    If DLookup("COUNT(*)", WK_TEMP) <= 0 Then
        MsgBox "データがありません。", vbExclamation
        Exit Sub
    End If
    
    strSQL = " UPDATE " & WK_TEMP & " SET FLGSEL = FALSE ; "
    dbs.Execute strSQL, dbFailOnError
    
    Me.MFLY010_MFLY010S00.Requery
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Call MSZZ024_M00("MFLY010_ALLCLE_Click", True)
    Resume Exit_Handler
End Sub

'==============================================================================*
'
'        MODULE_NAME      :全選択
'        MODULE_ID        :MFLY010_ALLSEL_Click
'        CREATE_DATE      : 2011/07/11 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_ALLSEL_Click()
On Error GoTo Err_Handler
    Dim strSQL As String
    
    If DLookup("COUNT(*)", WK_TEMP) <= 0 Then
        MsgBox "データがありません。", vbExclamation
        Exit Sub
    End If
    
    strSQL = " UPDATE " & WK_TEMP & " SET FLGSEL = TRUE ; "
    dbs.Execute strSQL, dbFailOnError
    
    Me.MFLY010_MFLY010S00.Requery
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Call MSZZ024_M00("MFLY010_ALLSEL_Click", True)
    Resume Exit_Handler
    
End Sub
''==============================================================================*
''
''        MODULE_NAME      :クリア
''        MODULE_ID        :MFLY010_CLNB_Click
''        CREATE_DATE      :2011/04/05 M.RYU
''-------------------------DELETE 2011/07/11 M.RYU-------------------------------
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub MFLY010_CLNB_Click()
'    On Error GoTo Err_MFLY010_CLNB_Click
'
'    If (MsgBox("クリアしますか？", 4) = 7) Then
'        Exit Sub
'    End If
'
'    strsql = ""
'    strsql = strsql & "DELETE WK_MQLY010_TEMP.* "
'    strsql = strsql & "FROM WK_MQLY010_TEMP "
'    strsql = strsql & ";"
'    dbs.Execute strsql, dbFailOnError
'
'    Call MFLY010_CLEA1
'
'    MsgBox "クリア完了", vbExclamation
'
'
'Exit_MFLY010_CLNB_Click:
'    Exit Sub
'
'Err_MFLY010_CLNB_Click:
'    Call MSZZ024_M00("MFLY010_CLNB_Click", True)
'    Resume Exit_MFLY010_CLNB_Click
'
'End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :MFLY010_KOKYF_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_KOKYF_AfterUpdate()
On Error GoTo Err_MFLY010_KOKYF_AfterUpdate
    'INSERT 2011/07/11 M.RYU----<S>
    If Nz(Me.MFLY010_KOKYF) = "" Then
        Me.MFLY010_KOKYC = ""
        Me.MFLY010_KOKYN = ""
        Me.MFLY010_KOTANTN = ""
    Else
        Me.MFLY010_KOKYC = Me.MFLY010_KOKYF.Column(1)
        Me.MFLY010_KOKYN = Me.MFLY010_KOKYF.Column(2)
        Me.MFLY010_KOTANTN = Me.MFLY010_KOKYF.Column(3)
    End If                '----<E>

'    doCmd.Requery "MFLY010_KOKYC"      'DELETE 2011/07/11 M.RYU                                                                                                                                         'INSERT 20010711 N.MIURA

Exit_MFLY010_KOKYF_AfterUpdate:
    Exit Sub

Err_MFLY010_KOKYF_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_MFLY010_KOKYF_AfterUpdate


End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :MFLY010_KOKYC_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_KOKYC_AfterUpdate()
    On Error GoTo Err_MFLY010_KOKYC_AfterUpdate
    
    'INSERT 2011/07/11 M.RYU----<S>
    If Nz(Me.MFLY010_KOKYC) = "" Then
        Me.MFLY010_KOKYN = ""
        Me.MFLY010_KOKYF = ""
        Me.MFLY010_KOTANTN = ""
    Else
        Me.MFLY010_KOKYN = Me.MFLY010_KOKYC.Column(1)
        Me.MFLY010_KOKYF = Me.MFLY010_KOKYC.Column(2)
        Me.MFLY010_KOTANTN = Me.MFLY010_KOKYC.Column(3)
    End If                '----<E>
    
'----DELETE 2011/07/11 M.RYU----<E>
'    Me.MFLY010_KOKYF = DLookup("USER_KANA", "dbo_USER_MAST", "USER_CODE = MFLY010_KOKYC")
'    Me.MFLY010_KOKYN = DLookup("USER_NAME", "dbo_USER_MAST", "USER_CODE = MFLY010_KOKYC")
'    doCmd.Requery "MFLY010_MFLY010S00"
'    Me.MFLY010_INSEB.SetFocus
'----DELETE 2011/07/11 M.RYU----<E>

Exit_MFLY010_KOKYC_AfterUpdate:
    Exit Sub

Err_MFLY010_KOKYC_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_MFLY010_KOKYC_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :登録
'        MODULE_ID        :MFLY010_INSEB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_INSEB_Click()
    On Error GoTo Err_MFLY010_INSEB_Click
    
    If Nz(Me.MFLY010_KOKYC) = "" Then   'INSERT 2011/07/11 M.RYU----<s>
        MsgBox "顧客を選択してください。", vbOKOnly
        Exit Sub
    End If                                                     '----<e>
    
    If (MsgBox("登録しますか？", 4) = 7) Then
        Exit Sub
    End If
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO "
'    strsql = strsql & "WK_MQLY010_TEMP "                               'DELETE 2011/07/11 M.RYU
    strSQL = strSQL & WK_TEMP                                           'INSERT 2011/07/11 M.RYU
    strSQL = strSQL & " ( "
    strSQL = strSQL & " BUMOC,  "
    strSQL = strSQL & " BUMON,  "   'INSERT 2011/06/18 M.RYU
    strSQL = strSQL & " KOKYC,  "
    strSQL = strSQL & " KOKYN,  "   'INSERT 2011/06/18 M.RYU
    strSQL = strSQL & " KOTANTN,"   'INSERT 2011/06/18 M.RYU
    strSQL = strSQL & " FLGSEL "    'INSERT 2011/06/18 M.RYU
    strSQL = strSQL & ") "
    strSQL = strSQL & "VALUES ("
    strSQL = strSQL & " '" & WK_BUMOC & "'                          ,"
    strSQL = strSQL & " '" & WK_BUMON & "'                          , " 'INSERT 2011/06/18 M.RYU
    strSQL = strSQL & " '" & Format(Me.MFLY010_KOKYC, "000000") & "', "
    strSQL = strSQL & " '" & Nz(Me.MFLY010_KOKYN) & "'              , " 'INSERT 2011/06/18 M.RYU
    strSQL = strSQL & " '" & Nz(Me.MFLY010_KOTANTN) & "'            , " 'INSERT 2011/06/18 M.RYU
    strSQL = strSQL & " FALSE                                         " 'INSERT 2011/06/18 M.RYU
    strSQL = strSQL & ")"
    strSQL = strSQL & ";"
    
    dbs.Execute strSQL, dbFailOnError
    
    Call MFLY010_CLEA1
    
    MsgBox "登録完了", vbExclamation

Exit_MFLY010_INSEB_Click:
    Exit Sub

Err_MFLY010_INSEB_Click:
    MsgBox Err.Description
    Resume Exit_MFLY010_INSEB_Click

End Sub
'==============================================================================*
'
'        MODULE_NAME      :削除
'        MODULE_ID        :MFLY010_DELEB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_DELEB_Click()
    On Error GoTo Err_MFLY010_DELEB_Click

    If MFLY010_CheckSelect = False Then Exit Sub                                        'INSERT 2011/07/11 M.RYU
    
    If (MsgBox("削除しますか？", 4) = 7) Then
        Exit Sub
    End If
    
'    strsql = ""                                                                        'DELETE 2011/07/11 M.RYU---<S>
'    strsql = strsql & "DELETE "
'    strsql = strsql & "WK_MQLY010_TEMP.* "
'    strsql = strsql & "FROM WK_MQLY010_TEMP "
'    strsql = strsql & "WHERE "
'    strsql = strsql & "BUMOC = " & "'" & WK_BUMOC & "'" & " AND "
'    strsql = strsql & "KOKYC = " & "'" & Format(Me.MFLY010_KOKYC, "000000") & "'"
'    strsql = strsql & ";"
'    dbs.Execute strsql, dbFailOnError
'    Call MFLY010_CLEA1                                                                 'DELETE 2011/07/11 M.RYU---<E>

    dbs.Execute " DELETE FROM " & WK_TEMP & " WHERE FLGSEL = TRUE ; ", dbFailOnError    'INSERT 2011/07/11 M.RYU
    Me.MFLY010_MFLY010S00.Requery                                                      'INSERT 2011/07/11 M.RYU
    
    MsgBox "削除完了", vbExclamation


Exit_MFLY010_DELEB_Click:
    Exit Sub

Err_MFLY010_DELEB_Click:
    MsgBox Err.Description
    Resume Exit_MFLY010_DELEB_Click
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :宛名印刷
'        MODULE_ID        :MFLY010_PRINB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_PRINB_Click()
    On Error GoTo Err_MFLY010_PRINB_Click
    
    If MFLY010_CheckSelect = False Then Exit Sub            'INSERT 2011/07/11 M.RYU
    
    If (MsgBox("実行しますか？", 4) = 7) Then
        Exit Sub
    End If
    
    Call MFLY010_DROPTB
    Call MFLY010_SELECT
    
    Select Case Me.MFLY010_LISTG
           Case "0"
                '小封筒
'                stDocName = "MRLY010_M00"                  'DELETE 2011/07/11 M.RYU
                stDocName = "LR0052_M00"                    'INSERT 2011/07/11 M.RYU
           Case "1"
                '中封筒
'                stDocName = "MRLY010_M10"                  'DELETE 2011/07/11 M.RYU
                stDocName = "LR0052_M10"                    'INSERT 2011/07/11 M.RYU
    End Select
    
    doCmd.OpenReport stDocName, acPreview

Exit_MFLY010_PRINB_Click:
    Exit Sub

Err_MFLY010_PRINB_Click:
    MsgBox Err.Description
    Resume Exit_MFLY010_PRINB_Click
    
End Sub


'==============================================================================*
'
'        MODULE_NAME      :閉じる
'        MODULE_ID        :MFLY010_EXITB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_EXITB_Click()
    On Error GoTo Err_MFLY010_EXITB_Click

    Call MFLY010_DBCLOS

    Call MSHS010_M00(PROG_ID, "2")
    
    doCmd.Close

Exit_MFLY010_EXITB_Click:
    Exit Sub

Err_MFLY010_EXITB_Click:
    MsgBox Err.Description
    Resume Exit_MFLY010_EXITB_Click
    
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
'        MODULE_NAME      :ワークテーブルの削除
'        MODULE_ID        :MFLY010_DROPTB
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_DROPTB()
    On Error GoTo Err_MFLY010_DROPTB
    
    strSQL = ""
'    strsql = strsql & "DROP TABLE WK_MQLY010_TEMP1 "       'DELETE 2011/07/11 M.RYU
    strSQL = strSQL & "DROP TABLE " & WK_TEMP1 & " ; "      'DINSERT 2011/07/11 M.RYU
    strSQL = strSQL & ";"
    dbs.Execute strSQL, dbFailOnError
    

Exit_MFLY010_DROPTB:
    Exit Sub

Err_MFLY010_DROPTB:
    MsgBox Err.Description
    Resume Exit_MFLY010_DROPTB

End Sub




'==============================================================================*
'
'        MODULE_NAME      :抽出
'        MODULE_ID        :MFLY010_SELECT
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_SELECT()
    On Error GoTo Err_MFLY010_SELECT

    MODL_ID = "MFLY010_SELECT"
                
    WK_KEISN = ""
    WK_DAIHN = "代表取締役"
    
    Select Case Me.MFLY010_KEISG
        Case "0"
             WK_KEISN = "　様"
        Case "1"
             WK_KEISN = "　御中"
        Case "2"
             WK_KEISN = "　殿"
        Case "3"
             WK_KEISN = ""
    End Select
    
    strSQL = ""
    strSQL = strSQL & "SELECT "

    strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> ''  ,              "   'INSERT 2011/08/03 M.RYU
    strSQL = strSQL & "    SEL_KOKY.USER_RNAME &" & "'" & WK_KEISN & "' ,       "   'INSERT 2011/08/03 M.RYU
    
    Select Case Me.MFLY010_DAIHG
        Case "0"
            '有
            strSQL = strSQL & "IIF(SEL_KOKY.USER_KKBN = 1                       "
            strSQL = strSQL & "   ,SEL_KOKY.USER_NAME                           "
            strSQL = strSQL & "   ,SEL_KOKY.USER_NAME &" & "'" & WK_KEISN & "') "
            strSQL = strSQL & " )                                               "   'INSERT 2011/08/03 M.RYU
            strSQL = strSQL & "                      AS KOKYN,                  "
'            strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> '' ,       "   'INSERT 2011/08/03 M.RYU -- 'DELETE 2011/09/15 M.RYU
            strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> '' , '' ,  "   'INSERT 2011/09/15 M.RYU
            
            
            '↓↓↓-----2011/10/27----M.RYU----------↓↓↓---
'            strSQL = strSQL & "IIF(SEL_KOKY.USER_KKBN = 1 "
'            strSQL = strSQL & "   ," & "'" & WK_DAIHN & "'&" & "'" & WK_KEISN & "'"
'            strSQL = strSQL & "   ,NULL)                                        "
            
            strSQL = strSQL & "     IIF(SEL_KOKY.USER_TANM IS NULL OR SEL_KOKY.USER_TANM = '　' OR SEL_KOKY.USER_TANM = ' '"
            strSQL = strSQL & "        ,NULL                                    "
            strSQL = strSQL & "        ,'" & WK_DAIHN & "　'" & " & SEL_KOKY.USER_TANM & " & "'" & WK_KEISN & "') "
            '↑↑↑-----2011/10/27----M.RYU----------↑↑↑---
            
            strSQL = strSQL & " )                                               "   'INSERT 2011/08/03 M.RYU
            strSQL = strSQL & "                      AS TANTN,                  "
        Case "1"
            '無
            strSQL = strSQL & "IIF(SEL_KOKY.USER_TANM IS NULL OR SEL_KOKY.USER_TANM = '　' OR SEL_KOKY.USER_TANM = ' '"
            strSQL = strSQL & "   ,SEL_KOKY.USER_NAME &" & "'" & WK_KEISN & "'  "
            strSQL = strSQL & "   ,SEL_KOKY.USER_NAME)                          "
            strSQL = strSQL & " )                                               "   'INSERT 2011/08/03 M.RYU
            strSQL = strSQL & "                      AS KOKYN,                  "
'            strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> ''  ,      "   'INSERT 2011/08/03 M.RYU  -- 'DELETE 2011/09/15 M.RYU
            strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> '' , ''  , "   'INSERT 2011/09/15 M.RYU
            strSQL = strSQL & "     IIF(SEL_KOKY.USER_TANM IS NULL OR SEL_KOKY.USER_TANM = '　' OR SEL_KOKY.USER_TANM = ' '"
            strSQL = strSQL & "        ,NULL                                    "
            strSQL = strSQL & "        ,SEL_KOKY.USER_TANM &" & "'" & WK_KEISN & "') "
            strSQL = strSQL & " )                                               "   'INSERT 2011/08/03 M.RYU
            strSQL = strSQL & "                      AS TANTN, "
    End Select
    
''    strSQL = strSQL & "SEL_KOKY.USER_YUBINO AS YUBIB, "                           'UPDATE 2011/08/03 M.RYU------<S>
''    strSQL = strSQL & "SEL_KOKY.USER_ADR_1  AS ADR1N, "
''    strSQL = strSQL & "SEL_KOKY.USER_ADR_2  AS ADR2N, "
''    'strSql = strSql & "NULL                 AS ADR3N, " 'DELETE 20070216 N.MIURA
''    strSQL = strSQL & "SEL_KOKY.USER_ADR_3  AS ADR3N, "  'INSERT 20070216 N.MIURA
        
    strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> ''  ,                 "
    strSQL = strSQL & "    SEL_KOKY.USER_RPOST  , SEL_KOKY.USER_YUBINO ) AS YUBIB, "
    strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> ''  ,                 "
    strSQL = strSQL & "    SEL_KOKY.USER_RADR_1 , SEL_KOKY.USER_ADR_1  ) AS ADR1N, "
    strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> ''  ,                 "
    strSQL = strSQL & "    SEL_KOKY.USER_RADR_2 , SEL_KOKY.USER_ADR_2  ) AS ADR2N, "
    strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> ''  ,                 "
    strSQL = strSQL & "    SEL_KOKY.USER_RADR_3 , SEL_KOKY.USER_ADR_3  ) AS ADR3N, " 'UPDATE 2011/08/03 M.RYU------<E>

    strSQL = strSQL & "'" & WK_KEISN & "' AS KEISN, "                           'INSERT 2026/04/14 N.IMAI
    strSQL = strSQL & " SEL_KOKY.*, "                                           'INSERT 2026/04/14 N.IMAI

    Select Case Me.MFLY010_KOKYG
        Case "0"
             '有
'            strsql = strsql & "WK_MQLY010_TEMP.BUMOC & WK_MQLY010_TEMP.KOKYC AS PRINC "    'DELETE 2011/07/11 M.RYU
            strSQL = strSQL & WK_TEMP & ".BUMOC & " & WK_TEMP & ".KOKYC AS PRINC "          'INSERT 2011/07/11 M.RYU
        Case "1"
             '無
             strSQL = strSQL & "NULL     AS PRINC "
    End Select
    
    strSQL = strSQL & " INTO "
    
'    strsql = strsql & "WK_MQLY010_TEMP1 "                                          'UPDATE 2011/07/11 M.RYU------<S>
'    strSQL = strSQL & " FROM "
'    strsql = strsql & "WK_MQLY010_TEMP  "
'    strsql = strsql & "INNER JOIN SEL_KOKY ON WK_MQLY010_TEMP.KOKYC = SEL_KOKY.KOKYC "

    strSQL = strSQL & " " & WK_TEMP1 & " "
    strSQL = strSQL & " FROM " & WK_TEMP & " "
    strSQL = strSQL & " INNER JOIN SEL_KOKY ON " & WK_TEMP & ".KOKYC = SEL_KOKY.KOKYC "
    strSQL = strSQL & " WHERE " & WK_TEMP & ".FLGSEL = TRUE               "         'UPDATE 2011/07/11 M.RYU------<E>   '選択フラグを設定、画面で選択したレコードを印刷
    strSQL = strSQL & ";"
    dbs.Execute strSQL, dbFailOnError

Exit_MFLY010_SELECT:
    Exit Sub

Err_MFLY010_SELECT:
    MsgBox Err.Description
    Resume Exit_MFLY010_SELECT

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＯＰＥＮ
'        MODULE_ID        :MFLY010_DBOPEN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_DBOPEN()
    On Error GoTo Err_MFLY010_DBOPEN
    
    Set dbs = CurrentDb()
    
Exit_MFLY010_DBOPEN:
    Exit Sub

Err_MFLY010_DBOPEN:
    MsgBox Err.Description
    Resume Exit_MFLY010_DBOPEN

End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＣＬＯＳＥ
'        MODULE_ID        :MFLY010_DBCLOS
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_DBCLOS()
    On Error GoTo Err_MFLY010_DBCLOS
    
    dbs.Close
    
    Set dbs = Nothing

Exit_MFLY010_DBCLOS:
    Exit Sub

Err_MFLY010_DBCLOS:
    MsgBox Err.Description
    Resume Exit_MFLY010_DBCLOS

End Sub

'==============================================================================*
'
'        MODULE_NAME      :画面クリアー１
'        MODULE_ID        :MFLY010_CLEA1
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MFLY010_CLEA1()
    On Error GoTo MFLY010_CLEA1_Err
    
    Me.MFLY010_KOKYF = Null
    Me.MFLY010_KOKYC = Null
    Me.MFLY010_KOKYN = Null
    
    doCmd.Requery "MFLY010_KOKYF"
    doCmd.Requery "MFLY010_KOKYC"                               'INSERT 20010920 N.MIURA
    doCmd.Requery "MFLY010_MFLY010S00"
    

MFLY010_CLEA1_Exit:
    Exit Sub

MFLY010_CLEA1_Err:
    MsgBox Error$
    Resume MFLY010_CLEA1_Exit

End Sub

'==============================================================================*
'
'        MODULE_NAME      :明細ﾚｺｰﾄﾞ選択チェック
'        MODULE_ID        :MFLY010_CheckSelect
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MFLY010_CheckSelect() As Boolean
    If DLookup("COUNT(*)", WK_TEMP) <= 0 Then
        MsgBox "データがありません。", vbExclamation
        GoTo falseExit
    End If
    If DLookup("COUNT(*)", WK_TEMP, "FLGSEL = TRUE") <= 0 Then
        MsgBox "レコードを選択してください。", vbExclamation
        GoTo falseExit
    End If
    
    MFLY010_CheckSelect = True
    Exit Function
    
falseExit:
    MFLY010_CheckSelect = False
    Exit Function
End Function
'****************************  ended or program ********************************

