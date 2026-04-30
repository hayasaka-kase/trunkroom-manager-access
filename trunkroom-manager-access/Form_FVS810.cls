VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS810"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード一覧
'   プログラムＩＤ　：　FVS810
'   作　成　日　　　：  2003/06/09
'   作　成　者　　　：  Eagle Soft 田島
'**********************************************
'修正履歴
'   修　正　日　　　：  2004/12/05
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.1
'   修　正　内　容　：  コンテナ種別：未設置は対象外とする
'**********************************************


Option Compare Database
Option Explicit
    
Dim objDb                     As Database
Dim objRs                     As Recordset

Dim bKinrin     As Boolean
Dim bFound      As Boolean

Private Sub Form_Activate()
    
    'フルスクリーン化
    doCmd.Maximize

End Sub

Private Sub Form_Close()
    'ワークテーブル削除
    Call DELWK_RTN
    'オブジェクト変数の解放
    Set objRs = Nothing
    Set objDb = Nothing

End Sub

Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next
   
    'F1:戻る
    Select Case KeyCode
    Case vbKeyF1
        KeyCode = 0
        Call cmd_終了_Click
    Case vbKeyF3, vbKeyF5 To vbKeyF12
        KeyCode = 0
    End Select

End Sub

Private Sub Form_Load()

    Dim strSyu          As String
    Dim str種別条件DISP As String

    'フルスクリーン化
    doCmd.Maximize
    
    'パラメータは
    '1～3桁目:"ON "=近隣含む "OFF"=近隣含まない
    '4～7桁目:"ON "=近隣含む "OFF"=近隣含まない
    '8桁目から検索条件の住所
    txt_検索条件 = Mid(Me.OpenArgs, 8)
    
    If Trim(Left(Me.OpenArgs, 3)) = "ON" Then
        bKinrin = True
        txt_検索条件DISP = Mid(Me.OpenArgs, 8) & "(近隣ヤード含む)"
    Else
        bKinrin = False
        txt_検索条件DISP = Mid(Me.OpenArgs, 8)
    End If
    
    '種別
    strSyu = Trim$(Mid(Me.OpenArgs, 4, 4))
    txt_検索条件2 = strSyu
    If strSyu <> "0000" Then
        str種別条件DISP = "(種別："
        str種別条件DISP = str種別条件DISP & Val(Trim$(Left$(strSyu, 2))) & " ～ "
        str種別条件DISP = str種別条件DISP & Val(Trim$(Right$(strSyu, 2)))
        str種別条件DISP = str種別条件DISP & ")"
    End If
    txt_検索条件DISP = txt_検索条件DISP & str種別条件DISP
    
    'Connect Database
    Set objDb = CurrentDb
    bFound = False
    
    'ワークテーブル削除
    Call DELWK_RTN

    'サブフォーム表示用のワークテーブル作成
    Call Make_Work
    
    'サブフォームの初期化
    Me.一覧サブ.Requery
    
    '検索結果が0件でなければサブフォームにフォーカス
    If bFound = True Then
        Me.一覧サブ.SetFocus
        Me.一覧サブ.Form!cmd_詳細.SetFocus
    End If
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    'フルスクリーン化
'    doCmd.Maximize
    
End Sub

'サブフォーム表示のためのワークテーブル作成
'FVS810_W01:住所からヤードマスタを検索し該当ヤードコードを登録
'           近隣ありの場合は、上の該当ヤードの近隣ヤードコードも追加する
'FVS810_W02:FVS810_W01に登録されているヤードの空コンテナ数を算出して登録。
'
'※実際にサブフォームに表示する際は、FVS810_Q01&Q02で種別ごとの集計をしている
Private Sub Make_Work()

    Dim strSQL As String
    Dim blnAdr As Boolean
    Dim lngCnt As Long
  
On Error GoTo err_rtn
  
    blnAdr = False
    
    If txt_検索条件 <> "" Then
        'ヤードマスタを住所を検索し、該当ヤードをFVS810_W01に登録
        strSQL = ""
        strSQL = "insert into FVS810_W01 "
        strSQL = strSQL & " ( YARD_CODE,YARD_NEAR ) "
        strSQL = strSQL & " SELECT dbo_YARD_MAST.YARD_CODE, 0 "
        strSQL = strSQL & " FROM dbo_YARD_MAST "
        strSQL = strSQL & "  where dbo_YARD_MAST.YARD_ADDR_1 Like ""*" & txt_検索条件 & "*"" "
        objDb.Execute strSQL, dbFailOnError
        
        '近隣ありならば、近隣ヤードマスタを検索しFVS810_W01に追加
        If bKinrin = True Then
            strSQL = ""
            strSQL = "insert into FVS810_W01 "
            strSQL = strSQL & " ( YARD_CODE, YARD_NEAR ) "
            strSQL = strSQL & " SELECT DISTINCT dbo_NYAR_MAST.NYAR_NCODE, 1 "
            strSQL = strSQL & " FROM FVS810_W01 INNER JOIN dbo_NYAR_MAST "
            strSQL = strSQL & "   ON FVS810_W01.YARD_CODE = dbo_NYAR_MAST.NYAR_YCODE "
            strSQL = strSQL & " WHERE dbo_NYAR_MAST.NYAR_NCODE not in "
            strSQL = strSQL & "        ( SELECT YARD_CODE FROM FVS810_W01 )"
            objDb.Execute strSQL, dbFailOnError
        End If
        blnAdr = True
    End If
    
    If blnAdr Then
        If txt_検索条件2 <> "0000" Then
            strSQL = ""
            strSQL = "delete FROM FVS810_W01 "
            strSQL = strSQL & " WHERE FVS810_W01.YARD_CODE not in "
            strSQL = strSQL & "        ( SELECT CNTA_CODE FROM dbo_CNTA_MAST "
            strSQL = strSQL & "          WHERE dbo_CNTA_MAST.CNTA_SYU >= " & Left$(txt_検索条件2, 2)
            strSQL = strSQL & "          AND dbo_CNTA_MAST.CNTA_SYU <= " & Right$(txt_検索条件2, 2) & ")"
            objDb.Execute strSQL, dbFailOnError
        End If
    Else
        strSQL = ""
        strSQL = "insert into FVS810_W01 "
        strSQL = strSQL & " ( YARD_CODE,YARD_NEAR ) "
        strSQL = strSQL & " SELECT DISTINCT dbo_CNTA_MAST.CNTA_CODE, 0 "
        strSQL = strSQL & " FROM dbo_CNTA_MAST "
        strSQL = strSQL & " WHERE dbo_CNTA_MAST.CNTA_SYU >= " & Left$(txt_検索条件2, 2)
        strSQL = strSQL & " AND dbo_CNTA_MAST.CNTA_SYU <= " & Right$(txt_検索条件2, 2)
        objDb.Execute strSQL, dbFailOnError
    End If

    'FVS810_W01のヤードの段と種別ごとのコンテナ空数を算出し、FVS810_W02に登録
    strSQL = ""
    strSQL = "insert into FVS810_W02 "
    strSQL = strSQL & " ( YARD_CODE, CONTAINER_STEP,CONTAINER_SYUBETSU, YARD_NEAR, CONTAINER_COUNT ) "
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "  dbo_CNTA_MAST.CNTA_CODE,"
    strSQL = strSQL & "  dbo_CNTA_MAST.CNTA_STEP,"
    strSQL = strSQL & "  dbo_CNTA_MAST.CNTA_SYU,"
    strSQL = strSQL & "  FVS810_W01.YARD_NEAR, "
    strSQL = strSQL & "  Count(*) - COUNT(dbo_CARG_FILE.CARG_NO) AS CNTCNTA"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "  ("
    strSQL = strSQL & "   (FVS810_W01 INNER JOIN dbo_CNTA_MAST"
    strSQL = strSQL & "    ON (    FVS810_W01.YARD_CODE = dbo_CNTA_MAST.CNTA_CODE"
    strSQL = strSQL & "       )"
    strSQL = strSQL & "   ) LEFT JOIN dbo_CARG_FILE"
    strSQL = strSQL & "    ON (    dbo_CNTA_MAST.CNTA_CODE = dbo_CARG_FILE.CARG_YCODE"
    strSQL = strSQL & "        AND dbo_CNTA_MAST.CNTA_NO   = dbo_CARG_FILE.CARG_NO"
    strSQL = strSQL & "        AND dbo_CARG_FILE.CARG_KYDATE Is Null"
    strSQL = strSQL & "       )"
    strSQL = strSQL & "  ) INNER JOIN dbo_NAME_MAST"
    strSQL = strSQL & "   ON (    dbo_CNTA_MAST.CNTA_SYU = dbo_NAME_MAST.NAME_CODE"
    strSQL = strSQL & "       AND dbo_NAME_MAST.NAME_ID = ""040"" "
    strSQL = strSQL & "      )"
    '>>>>> ADD START 20041204 K.ISHIZAKA >>>>>>>>>>>>
    '未設置は対象外とする
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "  dbo_CNTA_MAST.CNTA_SYU > 0"
    '<<<<< ADD END   20041204 K.ISHIZAKA <<<<<<<<<<<<
    strSQL = strSQL & " GROUP BY"
    strSQL = strSQL & "  dbo_CNTA_MAST.CNTA_CODE,"
    strSQL = strSQL & "  FVS810_W01.YARD_NEAR, "
    strSQL = strSQL & "  dbo_CNTA_MAST.CNTA_STEP,"
    strSQL = strSQL & "  dbo_CNTA_MAST.CNTA_SYU"
    strSQL = strSQL & " HAVING Count(*) - COUNT(dbo_CARG_FILE.CARG_NO) > 0"
    strSQL = strSQL & " ORDER BY"
    strSQL = strSQL & "  dbo_CNTA_MAST.CNTA_CODE,"
    strSQL = strSQL & "  dbo_CNTA_MAST.CNTA_STEP,"
    strSQL = strSQL & "  dbo_CNTA_MAST.CNTA_SYU;"
    objDb.Execute strSQL, dbFailOnError
    
    '検索結果の有無を確認
    If objDb.RecordsAffected = 0 Then
        bFound = False
    Else
        bFound = True
    End If

    lngCnt = objDb.RecordsAffected
    
    If txt_検索条件2 <> "0000" Then
        strSQL = ""
        strSQL = "delete FROM FVS810_W02 "
        strSQL = strSQL & " WHERE FVS810_W02.YARD_CODE not in "
        strSQL = strSQL & "        ( SELECT YARD_CODE FROM FVS810_W02 "
        strSQL = strSQL & "          WHERE FVS810_W02.CONTAINER_SYUBETSU >= " & Left$(txt_検索条件2, 2)
        strSQL = strSQL & "          AND FVS810_W02.CONTAINER_SYUBETSU <= " & Right$(txt_検索条件2, 2) & ")"
        objDb.Execute strSQL, dbFailOnError
    End If
    
    'FVS810_W02 全件削除の場合
    If objDb.RecordsAffected = lngCnt Then
        bFound = False
    End If
    
GoTo Exit_rtn

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
Exit_rtn:
    Exit Sub

End Sub

Private Sub DELWK_RTN()
    Dim strSQL     As String

On Error Resume Next
    
    'ワークの削除
    strSQL = ""
    strSQL = " delete from FVS810_W01 "
  
    objDb.Execute strSQL, dbFailOnError

    strSQL = ""
    strSQL = " delete from FVS810_W02 "
  
    objDb.Execute strSQL, dbFailOnError

End Sub

Private Sub cmd_終了_Click()
On Error Resume Next
    doCmd.Close acForm, Me.FormName
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Form_Unload

    Exit Sub
    
Err_Form_Unload:
    MsgBox Error$
    Exit Sub

End Sub
