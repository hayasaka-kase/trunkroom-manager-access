VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS820"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　空コンテナ予約入力
'   プログラムＩＤ　：　FVS820
'   作　成　日　　　：  2003/06/09
'   作　成　者　　　：  Eagle Soft 田島
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：  2004/11/06
'   修　正　者　　　：  N.MIURA
'   Ver             ：  0.1
'   修　正　内　容　：  キャンペーン表示
'**********************************************
'   修　正　日　　　：  2004/12/04
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.2
'   修　正　内　容　：  コンテナ種別：未設置は対象外
'**********************************************
'**********************************************
'   修　正　日　　　：  2005/05/10
'   修　正　者　　　：  K.mitsuno
'   Ver            ：  0.3
'   修　正　内　容　：  新受付入力画面対応
'**********************************************
'**********************************************
'   修　正　日　　　：  2005/06/14
'   修　正　者　　　：  T.SUZUKI
'   Ver            ：  0.4
'   修　正　内　容　：  空検索（コンテナ）と同様、空検索（トランク）においても
'                      [予約入力]画面に遷移できるよう、処理を追加
'**********************************************
'   修　正　日　　　：  2005/07/17
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.5
'   修　正　内　容　：  外部から再表示される前に保存する
'   修　正　内　容　：  画面遷移する前にカレントを保存
'**********************************************
Option Compare Database
Option Explicit
    
Dim intAns As Integer                        '確認入力用

Const strMsg_kaku = "確認"
Const strMsg_err = "エラー"

Dim objDb                   As Database
Dim objRs                   As Recordset

Dim WRI_ObjWS               As Workspace
Dim WRI_objDb               As Database

'上下段のデータの有無　0=上下段ともデータ無し
'                    1=上段のみあり
'                    2=下段のみあり
'                    3=両方あり
Dim vStep                   As Variant

'True=上段にフォーカスあり False=なし
Dim bActiveUp               As Boolean

'True=下段にフォーカスあり False=なし
Dim bActiveDown             As Boolean

Dim bChangeFlag_FVS820      As Boolean                          'INSERT 20050718 K.ISHIZAKA
Dim strCaption              As String                           'INSERT 20050718 K.ISHIZAKA

'サブフォーム表示用のワークを作成
'FVS820_W03:契約トランの最大STDATEのデータのみを抽出して登録
'FVS820_W01:該当ヤードの空コンテナを抽出して登録
'FVS820_W02:FVS820_W01のコピー。取消処理時にFVS820_W01にコピーする。
Public Sub Make_Work()
    
    Dim strSQL As String
  
On Error GoTo err_rtn
  
    If txt_ヤードコード = "" Then
        Exit Sub
    End If
  
    strSQL = ""
    strSQL = strSQL & " INSERT INTO FVS820_W03  "
    strSQL = strSQL & " (CARG_YCODE,CARG_NO,CARG_STDATE,CARG_KYDATE)"
    strSQL = strSQL & " SELECT dbo_CARG_FILE.CARG_YCODE, "
    strSQL = strSQL & "        dbo_CARG_FILE.CARG_NO,"
    strSQL = strSQL & "        FVS820_Q01.MAXSTDATE,"
    strSQL = strSQL & "        dbo_CARG_FILE.CARG_KYDATE "
    strSQL = strSQL & "   FROM FVS820_Q01 INNER JOIN dbo_CARG_FILE "
    strSQL = strSQL & "     ON (FVS820_Q01.MAXSTDATE = dbo_CARG_FILE.CARG_FSDATE) AND "
    strSQL = strSQL & "        (FVS820_Q01.CARG_NO = dbo_CARG_FILE.CARG_NO) AND "
    strSQL = strSQL & "        (FVS820_Q01.CARG_YCODE = dbo_CARG_FILE.CARG_YCODE) "
    strSQL = strSQL & " where dbo_CARG_FILE.CARG_YCODE = " & txt_ヤードコード
    
    objDb.Execute strSQL, dbFailOnError
    
'(2003/07/01 ADD) >>>
    strSQL = ""
    strSQL = strSQL & " UPDATE FVS820_W03 SET CARG_KYDATE = NULL "
    strSQL = strSQL & " WHERE (CARG_YCODE IN "
    strSQL = strSQL & "       (SELECT  FVS820_W03.CARG_YCODE "
    strSQL = strSQL & "        FROM FVS820_W03 "
    strSQL = strSQL & "        GROUP BY FVS820_W03.CARG_YCODE,FVS820_W03.CARG_NO,FVS820_W03.CARG_STDATE "
    strSQL = strSQL & "        HAVING COUNT(*) > 1)"
    strSQL = strSQL & "  AND CARG_NO IN"
    strSQL = strSQL & "       (SELECT  FVS820_W03.CARG_NO "
    strSQL = strSQL & "        FROM FVS820_W03 "
    strSQL = strSQL & "        GROUP BY FVS820_W03.CARG_YCODE,FVS820_W03.CARG_NO,FVS820_W03.CARG_STDATE "
    strSQL = strSQL & "        HAVING COUNT(*) > 1)"
    strSQL = strSQL & "  AND CARG_STDATE IN"
    strSQL = strSQL & "       (SELECT  FVS820_W03.CARG_STDATE "
    strSQL = strSQL & "        FROM FVS820_W03 "
    strSQL = strSQL & "        GROUP BY FVS820_W03.CARG_YCODE,FVS820_W03.CARG_NO,FVS820_W03.CARG_STDATE "
    strSQL = strSQL & "        HAVING COUNT(*) > 1))"
    strSQL = strSQL & "  AND CARG_NO IN"
    strSQL = strSQL & "       (SELECT  FVS820_W03.CARG_NO "
    strSQL = strSQL & "        FROM FVS820_W03 "
    strSQL = strSQL & "        WHERE FVS820_W03.CARG_KYDATE IS NULL)"

    objDb.Execute strSQL, dbFailOnError

    strSQL = ""
    strSQL = strSQL & " DELETE FROM FVS820_W03"
    strSQL = strSQL & " WHERE (CARG_YCODE IN "
    strSQL = strSQL & "       (SELECT  FVS820_W03.CARG_YCODE "
    strSQL = strSQL & "        FROM FVS820_W03 "
    strSQL = strSQL & "        GROUP BY FVS820_W03.CARG_YCODE,FVS820_W03.CARG_NO,FVS820_W03.CARG_STDATE "
    strSQL = strSQL & "        HAVING COUNT(*) > 1)"
    strSQL = strSQL & "  AND CARG_NO IN"
    strSQL = strSQL & "       (SELECT  FVS820_W03.CARG_NO "
    strSQL = strSQL & "        FROM FVS820_W03 "
    strSQL = strSQL & "        GROUP BY FVS820_W03.CARG_YCODE,FVS820_W03.CARG_NO,FVS820_W03.CARG_STDATE "
    strSQL = strSQL & "        HAVING COUNT(*) > 1)"
    strSQL = strSQL & "  AND CARG_STDATE IN"
    strSQL = strSQL & "       (SELECT  FVS820_W03.CARG_STDATE "
    strSQL = strSQL & "        FROM FVS820_W03 "
    strSQL = strSQL & "        GROUP BY FVS820_W03.CARG_YCODE,FVS820_W03.CARG_NO,FVS820_W03.CARG_STDATE "
    strSQL = strSQL & "        HAVING COUNT(*) > 1))"
    strSQL = strSQL & "  AND CARG_NO IN"
    strSQL = strSQL & "       (SELECT  FVS820_W03.CARG_NO "
    strSQL = strSQL & "        FROM FVS820_W03 "
    strSQL = strSQL & "        WHERE FVS820_W03.CARG_KYDATE IS NOT NULL)"

    objDb.Execute strSQL, dbFailOnError
'(2003/07/01 ADD) <<<

'--- T.SUZUKI
    strSQL = ""
    strSQL = "INSERT INTO FVS820_W01 "
    strSQL = strSQL & " ( CNTA_YARD, CNTA_NO, CNTA_STEP, "
    strSQL = strSQL & " CNTA_YNAME, CNTA_YTEL, NAME_RYAK, CNTA_KYDATE, CNTA_CHANGE ) "
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & Val(txt_ヤードコード) & ","
    strSQL = strSQL & " dbo_CNTA_MAST.CNTA_NO, dbo_CNTA_MAST.CNTA_STEP, "
    strSQL = strSQL & " dbo_CNTA_MAST.CNTA_YNAME, dbo_CNTA_MAST.CNTA_YTEL, "
    strSQL = strSQL & " dbo_NAME_MAST.NAME_RYAK, FVS820_W03.CARG_KYDATE, 0 "
    strSQL = strSQL & " FROM (dbo_CNTA_MAST "
    strSQL = strSQL & " LEFT JOIN FVS820_W03 ON "
    strSQL = strSQL & " (dbo_CNTA_MAST.CNTA_NO = FVS820_W03.CARG_NO) AND "
    strSQL = strSQL & " (dbo_CNTA_MAST.CNTA_CODE = FVS820_W03.CARG_YCODE)) "
    strSQL = strSQL & " INNER JOIN dbo_NAME_MAST ON dbo_CNTA_MAST.CNTA_SYU = dbo_NAME_MAST.NAME_CODE "
    strSQL = strSQL & " Where dbo_CNTA_MAST.CNTA_CODE = " & txt_ヤードコード
    'strSql = strSql & "   AND dbo_CNTA_MAST.CNTA_SYU <= 40 "
    '>>>>> ADD START 20041204 K.ISHIZAKA >>>>>>>>>>>>
    '未設置は対象外とする
    'ちなみに上の行は既にコメントになっていた
    strSQL = strSQL & "   AND dbo_CNTA_MAST.CNTA_SYU > 0 "
    '<<<<< ADD END   20041204 K.ISHIZAKA <<<<<<<<<<<<
    strSQL = strSQL & "   AND dbo_NAME_MAST.NAME_ID = ""040"" "
    strSQL = strSQL & "   AND ( (FVS820_W03.CARG_YCODE Is Null) or "
    strSQL = strSQL & "           (FVS820_W03.CARG_NO Is not Null AND "
    strSQL = strSQL & "           FVS820_W03.CARG_KYDATE Is not Null) )"

'★下記SQLはFVS020を参考にしたもの、上記はFVS820修正前のSQL文
'    strSQL = ""
'    strSQL = "INSERT INTO FVS820_W01 "
'    strSQL = strSQL & " ( CNTA_YARD, CNTA_NO, CNTA_STEP, "
'    strSQL = strSQL & " CNTA_YNAME, CNTA_YTEL, NAME_RYAK, CNTA_KYDATE, CNTA_CHANGE ) "
'    strSQL = strSQL & " SELECT "
'    strSQL = strSQL & Val(txt_ヤードコード) & ","
'    strSQL = strSQL & " dbo_CNTA_MAST.CNTA_NO, dbo_CNTA_MAST.CNTA_STEP, "
'    strSQL = strSQL & " dbo_CNTA_MAST.CNTA_YNAME, dbo_CNTA_MAST.CNTA_YTEL, "
'    strSQL = strSQL & " dbo_NAME_MAST.NAME_RYAK, FVS820_W03.CARG_KYDATE, 0 "
'    strSQL = strSQL & " FROM (dbo_CNTA_MAST "
'    strSQL = strSQL & " LEFT JOIN FVS820_W03 ON "
'    strSQL = strSQL & " (dbo_CNTA_MAST.CNTA_NO = FVS820_W03.CARG_NO) AND "
'    strSQL = strSQL & " (dbo_CNTA_MAST.CNTA_CODE = FVS820_W03.CARG_YCODE)) "
'    strSQL = strSQL & " INNER JOIN dbo_NAME_MAST ON dbo_CNTA_MAST.CNTA_SYU = dbo_NAME_MAST.NAME_CODE "
'    strSQL = strSQL & " Where dbo_CNTA_MAST.CNTA_CODE = " & txt_ヤードコード
'    '>>>>> CHANGE START 20041204 K.ISHIZAKA >>>>>>>>>>>>>>>>>>>>>>>>>>>
'    '未設置は対象外とする
'    'strSQL = strSQL & "   AND dbo_CNTA_MAST.CNTA_SYU <= 40 "
'    strSQL = strSQL & "   AND dbo_CNTA_MAST.CNTA_SYU BETWEEN 1 AND 40 "
'    '<<<<< CHANGE END   20041204 K.ISHIZAKA <<<<<<<<<<<<<<<<<<<<<<<<<<<
'    strSQL = strSQL & "   AND dbo_NAME_MAST.NAME_ID = ""040"" "
'    strSQL = strSQL & "   AND ( (FVS820_W03.CARG_YCODE Is Null) or "
'    strSQL = strSQL & "           (FVS820_W03.CARG_NO Is not Null AND "
'    strSQL = strSQL & "           FVS820_W03.CARG_KYDATE Is not Null) )"

    objDb.Execute strSQL, dbFailOnError
'--- T.SUZUKI

    'INSERT 20050516 MITSUNO START
    '予約受付トランより予約顧客名・電話番号の取得
    strSQL = ""
    strSQL = strSQL & "UPDATE "
    strSQL = strSQL & "FVS820_W01 INNER JOIN dbo_YOUK_TRAN ON "
    strSQL = strSQL & "FVS820_W01.CNTA_NO = dbo_YOUK_TRAN.YOUKT_NO "
    strSQL = strSQL & "AND FVS820_W01.CNTA_YARD = dbo_YOUK_TRAN.YOUKT_YCODE "    'INSERT 20050713 K.ISHIZAKA
    strSQL = strSQL & "SET "
    strSQL = strSQL & "FVS820_W01.YOUKT_UKNO = dbo_YOUK_TRAN.YOUKT_UKNO, "
    strSQL = strSQL & "FVS820_W01.CNTA_YNAME = dbo_YOUK_TRAN.YOUKT_NAME, "
    '電話番号→携帯電話→FAXでセット
    strSQL = strSQL & "FVS820_W01.CNTA_YTEL = " & "Nz(dbo_YOUK_TRAN.YOUKT_TEL, Nz(dbo_YOUK_TRAN.YOUKT_KEITAI, NZ(dbo_YOUK_TRAN.YOUKT_FAX,''))) "
    
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "((dbo_YOUK_TRAN.YOUKT_YUKBN = 02) AND (dbo_YOUK_TRAN.YOUKT_TKDATE >= " & Format(DATE, "yyyy/mm/dd") & ")) "
    strSQL = strSQL & "OR (dbo_YOUK_TRAN.YOUKT_YUKBN = 10) "
    strSQL = strSQL & ";"
    
    objDb.Execute strSQL, dbFailOnError
    'INSERT 20050516 MITSUNO END
    
    'FVS820_W02作成
    Call MakeTorikeshiTable
    
    GoTo Exit_rtn
    
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    Exit Sub

End Sub
Private Sub MakeTorikeshiTable()
    
    Dim strSQL As String
  
On Error GoTo err_rtn
    
    Call DeleteWk2
    
    '取消用のテーブル作成
    strSQL = ""
    strSQL = "INSERT INTO FVS820_W02 "
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM FVS820_W01 "

    objDb.Execute strSQL, dbFailOnError
    
    GoTo Exit_rtn
    
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    Exit Sub

End Sub

Private Sub UpdateCntaChange()
    
    Dim strSQL As String
  
On Error GoTo err_rtn
    
    '変更フラグを全部OFFに変更
    strSQL = ""
    strSQL = "UPDATE FVS820_W01 "
    strSQL = strSQL & " set CNTA_CHANGE = 0 "

    objDb.Execute strSQL, dbFailOnError
    
    GoTo Exit_rtn
    
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    Exit Sub

End Sub

Private Sub CLR_RTN()
  
  Dim objControl As Object
  
  '画面上のテキストボックスをクリアする
  
On Error Resume Next  'エラーがあっても処理続行
    
    
    For Each objControl In Me.Controls
  
        Select Case objControl.ControlType
        Case acTextBox
            If Nz(objControl.OldValue, "") = "" Then
                objControl.VALUE = Null
            End If
        Case acComboBox
            If Nz(objControl.OldValue, "") = "" Then
                objControl.VALUE = ""
            End If
        Case Else
        End Select
    Next
          
    Me.txt_ヤード名 = ""
    Me.txt_ヤード住所 = ""
  
          
    Call DELWK_RTN
  
    'サブフォームの初期化
    Me.サブ上.Requery
    Me.サブ下.Requery
  
End Sub


Private Sub CmLockControl(ByVal ContNM As Object, ByVal iMode As Integer)

    '引数 ０  使用可
    '     １  使用不可
    Dim objControl As Object
  
    Set objControl = ContNM
  
    If iMode = 0 Then
        objControl.Enabled = True
        objControl.Locked = False
    Else
        objControl.Enabled = False
        objControl.Locked = True
    End If
  
End Sub

Private Function WRI_RTN() As Boolean

On Error GoTo err_rtn
  
    WRI_RTN = False
  
    doCmd.Hourglass True
    doCmd.SetWarnings False
  
    Set WRI_ObjWS = DBEngine(0)
    Set WRI_objDb = CurrentDb
  
    'トランザクション開始
    WRI_ObjWS.BeginTrans
  
    '予約設定
    If WRI_RTN_For_CNTA_MAST() = False Then
        GoTo Update_Err
    End If
      
    'コミット
    WRI_ObjWS.CommitTrans
    WRI_RTN = True
  
    bChangeFlag_FVS820 = False                                                  'INSERT 20050718 K.ISHIZAKA
    Me.Caption = strCaption                                                     'INSERT 20050718 K.ISHIZAKA
    
    GoTo Exit_rtn

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Update_Err:
    'ロールバック
    WRI_ObjWS.Rollback
    MsgBox "更新中にエラーが発生しました。" & vbCrLf & "（処理は中断されました。）", vbCritical, strMsg_kaku

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set WRI_objDb = Nothing
    Set WRI_ObjWS = Nothing

    Exit Function

End Function

Private Function WRI_RTN_For_CNTA_MAST() As Boolean
  
    'FVS820_W01からコンテナマスタに予約をセットする。
  
    Dim strSQL  As String
  
On Error GoTo err_rtn
    WRI_RTN_For_CNTA_MAST = False
  
    strSQL = ""
    strSQL = " UPDATE FVS820_W01 LEFT JOIN dbo_CNTA_MAST ON "
    
    strSQL = strSQL & " ( FVS820_W01.CNTA_YARD = dbo_CNTA_MAST.CNTA_CODE) AND"
    strSQL = strSQL & " ( FVS820_W01.CNTA_NO = dbo_CNTA_MAST.CNTA_NO )"
    
    strSQL = strSQL & " SET dbo_CNTA_MAST.CNTA_YNAME = FVS820_W01.CNTA_YNAME,"
    strSQL = strSQL & "     dbo_CNTA_MAST.CNTA_YTEL = FVS820_W01.CNTA_YTEL,"
    strSQL = strSQL & "     dbo_CNTA_MAST.CNTA_UPDATE = " & " CDATE(#" & Format(Now, "yyyy/mm/dd") & "# )"
    
    strSQL = strSQL & " WHERE FVS820_W01.CNTA_CHANGE = 1 "
    
    doCmd.RunSQL strSQL, True
    
    WRI_RTN_For_CNTA_MAST = True
    GoTo Exit_rtn

err_rtn:
    '*** エラー処理（Ｒｏｌｌｂａｃｋ） ***
    Dim errLoop As Error
    
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    If DBEngine.Errors.Count > 0 Then
        For Each errLoop In DBEngine.Errors
            MsgBox "Error Number:" + CStr(errLoop.Number) + vbCrLf + errLoop.Description
        Next
    End If

    Err.Clear
    
Exit_rtn:
    Exit Function

End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_WEB_Click
'        CREATE_DATE      :2004/07/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_WEB_Click()
    Call MSZZ009_M00(Me.cmd_WEB, Me.txtsubAddress, Me.txtAddress)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :キャンペーン表示
'        MODULE_ID        :cmd_CAMPB_Click
'        CREATE_DATE      :2004/11/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_CAMPB_Click()
    
    doCmd.OpenForm "FVS820_S50", , , , , acDialog

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_取消_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_取消_Click()

    intAns = MsgBox("データ入力を取り消します。よろしいですか？", vbQuestion + vbYesNo + vbDefaultButton2, strMsg_kaku)
    If intAns = vbNo Then
        Call SubFormSetFocus
        GoTo Exit_rtn
    End If
  
    'FVS820_W02をFVS820_W01にコピーする
    'Call DeleteWk1                                                             'DELETE 20050718 K.ISHIZAKA
    'Call DataTorikeshi                                                         'DELETE 20050718 K.ISHIZAKA
    Call DELWK_RTN                                                              'INSERT 20050718 K.ISHIZAKA
    Call MakeDisplay                                                            'INSERT 20050718 K.ISHIZAKA
    
    Call SubFormSetFocus
  
    Me.サブ上.Requery
    Me.サブ下.Requery
    
    bChangeFlag_FVS820 = False                                                  'INSERT 20050718 K.ISHIZAKA
    Me.Caption = strCaption                                                     'INSERT 20050718 K.ISHIZAKA
    
Exit_rtn:

End Sub

Public Sub DataTorikeshi()
  
    Dim strSQL As String
  
On Error GoTo err_rtn
  
    Call DeleteWk1
    
    '取消用のテーブルからコピー
    strSQL = ""
    strSQL = "INSERT INTO FVS820_W01 "
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM FVS820_W02 "

    objDb.Execute strSQL, dbFailOnError
    GoTo Exit_rtn
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
Exit_rtn:
    Exit Sub
End Sub
'Start 2003/06/23 azegami@eagle-soft.co.jp
Private Sub cmd_配置_Click()
    
    On Error GoTo err_rtn
    
    Dim vRET2        As Variant
    Dim strExcel     As String
    Dim RetVal       As Variant
  
    vRET2 = Nz(DLookup("CONT_EXCEL_PATH", "dbo_CONT_MAST", ""), "")
    
    If (vRET2 = "") Or (vRET2 = " ") Then
        MsgBox "EXCEL情報パスが登録されていません。", vbOK, strMsg_kaku
        Exit Sub
    End If

    If Right(vRET2, 1) = "\" Then
        vRET2 = vRET2
        Else
            vRET2 = vRET2 & "\"
    End If

    strExcel = Dir(vRET2 & Format(Me.txt_ヤードコード, "000000") & "*.XLS")
    
    If strExcel <> "" Then
        Call Y_Exec(Me.hwnd, vRET2 & strExcel, "", "", SW_SHOWNORMAL)
        Exit Sub
        Else
            MsgBox "ｺﾝﾃﾅ配置EXCELが存在しません。", vbOK, strMsg_kaku
            Exit Sub
    End If

err_rtn:

End Sub
'End 2003/06/23 azegami@eagle-soft.co.jp

Private Sub cmd_保存_Click()
On Error GoTo err_rtn
    
    '確認メッセージ
    'intAns = MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, strMsg_kaku) 'DELETE 20050718 K.ISHIZAKA
    intAns = MsgBox("修正中の内容を保存します。よろしいですか？", vbQuestion + vbYesNo, strMsg_kaku)  'INSERT 20050718 K.ISHIZAKA
    
    If intAns = vbNo Then
        Call SubFormSetFocus
        GoTo Exit_rtn
    End If
    
    '出力
    Call WRI_RTN
    
    '変更フラグクリア
    Call UpdateCntaChange
    '取消テーブル(FVS820_W02)再登録
    Call MakeTorikeshiTable
    
    Call SubFormSetFocus
    
    Me.サブ上.Requery
    Me.サブ下.Requery
    
    GoTo Exit_rtn

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    
Exit_rtn:
    Exit Sub

End Sub
'==============================================================================*
'
'        MODULE_NAME      :受付入力(ｺﾝﾃﾅ未確定⇒ﾔｰﾄﾞ番号のみ渡す)
'        MODULE_ID        :cmd_YUke_Click
'        CREATE_DATE      :2005/05/06 K.MITSUNO
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdYUke_Click()
    
    Dim vHData As Variant
    
    doCmd.RunCommand acCmdSaveRecord                                            'INSERT 20050718 K.ISHIZAKA
    If Me.SaveData() = False Then                                               'INSERT 20050718 K.ISHIZAKA
        Me.txt_ヤードコード.SetFocus                                            'INSERT 20050718 K.ISHIZAKA
        Exit Sub                                                                'INSERT 20050718 K.ISHIZAKA
    End If                                                                      'INSERT 20050718 K.ISHIZAKA
    
    'ﾔｰﾄﾞ番号(6桁)
    vHData = Format$(Me.txt_ヤードコード, "000000")
    
    '予約入力画面　起動
    doCmd.OpenForm "FVS500", , , , , , vHData
    
End Sub

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
    Set WRI_ObjWS = Nothing
    Set WRI_objDb = Nothing
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Deactivate
'        CREATE_DATE      :2005/07/17 K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Deactivate()
    On Error Resume Next
    doCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    
    Debug.Print DataErr
    Debug.Print Response
    Select Case DataErr
    Case 2169       '「レコードは保存されません。」という時
        DataErr = 0
        Response = 0
    End Select
    
End Sub

Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next
   
    'F1:戻る
    'F2:保存
    'F3:取消
    'F5:段切り替え
    'F8:配置
    Select Case KeyCode
    Case vbKeyF1
        KeyCode = 0
        Call cmd_終了_Click
    Case vbKeyF2
        KeyCode = 0
        Me.cmd_保存.SetFocus
        Call cmd_保存_Click
    Case vbKeyF3
        KeyCode = 0
        cmd_取消.SetFocus
        Call cmd_取消_Click
    Case vbKeyF5
        If bActiveUp = True Then
            If vStep >= 2 Then
                Me.サブ下.SetFocus
                '2005/06/01 MITSUNO
                Me.サブ下.Form!cmdUke.SetFocus
                '予約名/電話番号変更時、予約トランに受付番号が入っていたら
                'ｺﾝﾃﾅﾏｽﾀへUpdateしない処理追加の為、以下の処理をコメント
'                Me.サブ下.Form!txt_予約客先名.SetFocus
            End If
        ElseIf bActiveDown = True Then
            If vStep = 1 Or vStep = 3 Then
                Me.サブ上.SetFocus
                '2005/06/01 MITSUNO
                Me.サブ下.Form!cmdUke.SetFocus
                '予約名/電話番号変更時、予約トランに受付番号が入っていたら
                'ｺﾝﾃﾅﾏｽﾀへUpdateしない処理追加の為、以下の処理をコメント
'                Me.サブ上.Form!txt_予約客先名.SetFocus
            End If
        End If
        
    Case vbKeyF8
        KeyCode = 0
        Call cmd_配置_Click
    Case vbKeyF3, vbKeyF5 To vbKeyF12
        KeyCode = 0
    End Select

End Sub

Private Sub Form_Load()

    strCaption = Me.Caption                                                     'INSERT 20050718 K.ISHIZAKA

    'フルスクリーン化
    doCmd.Maximize
    
    txt_ヤードコード = Me.OpenArgs
    
    '解約日は現在日より以降の日付しか表示しない
    'その条件に使う
    txt_NowDate = Format(Now, "yyyy/mm/dd")
    
    Set objDb = CurrentDb
  
    'フォームの初期化処理
    Call CLR_RTN
    Call Etc_Lock_Rtn
  
    vStep = 0
    bActiveUp = False
    bActiveDown = False
    
    '表示データ作成
    Call MakeDisplay
    
End Sub

Private Sub MakeDisplay()

    '住所／ヤード名等取得
    Call GetYardData(1)
  
    'サブフォームでデータリンスするワークテーブル作成
    Call Make_Work
    
    Me.サブ上.Enabled = True
    Me.サブ下.Enabled = True

    Me.サブ上.Requery
    Me.サブ下.Requery
  
    Me.サブ上A.Enabled = True         '2000/07/03 insert n-miua
    Me.サブ下A.Enabled = True         '2000/07/03 insert n-miua
    
    Me.サブ上A.Requery                '2000/07/03 insert n-miua
    Me.サブ下A.Requery                '2000/07/03 insert n-miua
        
    Call SubFormSetFocus

End Sub

Private Sub Form_Open(Cancel As Integer)

'    doCmd.Maximize
    
End Sub
Private Sub Etc_Lock_Rtn()
    '入力区分によるコントロールのロックを設定
      
    Me.サブ上.Enabled = False
    Me.サブ下.Enabled = False
    Me.cmd_保存.Enabled = True
    Me.cmd_取消.Enabled = True
    Me.cmd_終了.Enabled = True
  
End Sub
Private Sub cmd_終了_Click()
On Error Resume Next
    doCmd.Close acForm, Me.FormName
    
End Sub

Private Sub DELWK_RTN()

    Call DeleteWk1
    Call DeleteWk2
    Call DeleteWk3

End Sub
Private Sub DeleteWk1()

    Dim strSQL     As String

On Error Resume Next
    'ワークの削除
    strSQL = ""
    strSQL = " delete from FVS820_W01"
  
    objDb.Execute strSQL, dbFailOnError

End Sub
Private Sub DeleteWk2()

    Dim strSQL     As String

On Error Resume Next
    'ワークの削除
    strSQL = ""
    strSQL = " delete from FVS820_W02 "
  
    objDb.Execute strSQL, dbFailOnError

End Sub
Private Sub DeleteWk3()

    Dim strSQL     As String

On Error Resume Next
    'ワークの削除
    strSQL = ""
    strSQL = " delete from FVS820_W03 "
  
    objDb.Execute strSQL, dbFailOnError

End Sub

Private Function CheckCntaData()
    
    Dim strSQL As String
  
On Error GoTo err_rtn
        
    vStep = 0
    
    '上段データ有無チェック
    strSQL = ""
    strSQL = " SELECT"
    strSQL = strSQL & " count(*) AS RECCOUNT"
    strSQL = strSQL & " FROM FVS820_W01"
    strSQL = strSQL & " Where"
    strSQL = strSQL & " CNTA_STEP = 0"
        
    Set objRs = objDb.OpenRecordset(strSQL, dbOpenSnapshot)
      
    If Not (objRs.BOF And objRs.EOF) Then
        If objRs.Fields("RECCOUNT") <> 0 Then
            'あり
            vStep = vStep + 1
        End If
    End If
    Set objRs = Nothing
    
    '下段データ有無チェック
    strSQL = ""
    strSQL = " SELECT"
    strSQL = strSQL & " count(*) AS RECCOUNT"
    strSQL = strSQL & " FROM FVS820_W01"
    strSQL = strSQL & " Where"
    strSQL = strSQL & " CNTA_STEP = 1"
        
    Set objRs = objDb.OpenRecordset(strSQL, dbOpenSnapshot)
      
    If Not (objRs.BOF And objRs.EOF) Then
        If objRs.Fields("RECCOUNT") <> 0 Then
            'あり
            vStep = vStep + 2
        End If
    End If
    Set objRs = Nothing
    CheckCntaData = True

    GoTo Exit_rtn

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set objRs = Nothing
    Exit Function

End Function
Private Function GetYardData(ByVal iMode As Integer) As Boolean
    'ヤードマスタ問い合わせ
    '引数 ０：存在チェック
    '     １：データ参照
    
    Dim strSQL As String
    Dim vRet   As Variant
    Dim i      As Integer
    Dim vRET2  As Variant
  
On Error GoTo err_rtn
  
    GetYardData = False
  
    If txt_ヤードコード = "" Then
        GetYardData = True
        GoTo Exit_rtn
    End If
      
    doCmd.Hourglass True
    doCmd.SetWarnings False
  
  
    If iMode = 0 Then
  
        '存在チェック
        strSQL = ""
        strSQL = strSQL & " YARD_CODE =" & Val(txt_ヤードコード)
      
        vRet = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strSQL), "")
      
        If vRet = "" Then
            MsgBox "該当するヤードマスタがありません。", vbInformation, strMsg_kaku
            'GetYardData = True     'Del 01-07-2000 Araki
            GoTo Exit_rtn             'Add 01-07-2000 Araki
        End If
     
  
    Else
        strSQL = ""                                                             'INSERT 20041106 N.MIURA START
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "YARD_NAME,"
        strSQL = strSQL & "YARD_ADDR_1,"
        strSQL = strSQL & "YARD_ADDR_2,"
        strSQL = strSQL & "YARD_ADDR_3,"
        strSQL = strSQL & "YARD_NOTE,"
        strSQL = strSQL & "YARD_SEV1N,"
        strSQL = strSQL & "YARD_SEV2N,"
        strSQL = strSQL & "YARD_SEV3N,"
        strSQL = strSQL & "YARD_ENDEN,"
        strSQL = strSQL & "YARD_BIKON "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "dbo_YARD_MAST "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & " YARD_CODE =" & Val(txt_ヤードコード)
        
        Set objRs = objDb.OpenRecordset(strSQL, dbOpenSnapshot)
      
        If Not (objRs.BOF And objRs.EOF) Then
           With Me
           .txt_ヤード名 = objRs.Fields("YARD_NAME")
           .txt_ヤード住所 = objRs.Fields("YARD_ADDR_1") & _
                             objRs.Fields("YARD_ADDR_2") & _
                             objRs.Fields("YARD_ADDR_3")
           .txt_付近メモ = objRs.Fields("YARD_NOTE")
           .txt_備考 = objRs.Fields("YARD_BIKON")
            
           If Not (IsNull(objRs.Fields("YARD_SEV1N")) Or objRs.Fields("YARD_SEV1N") = "" Or objRs.Fields("YARD_SEV1N") = " ") Or _
              Not (IsNull(objRs.Fields("YARD_SEV2N")) Or objRs.Fields("YARD_SEV2N") = "" Or objRs.Fields("YARD_SEV2N") = " ") Or _
              Not (IsNull(objRs.Fields("YARD_SEV3N")) Or objRs.Fields("YARD_SEV3N") = "" Or objRs.Fields("YARD_SEV3N") = " ") Or _
              Not (IsNull(objRs.Fields("YARD_ENDEN")) Or objRs.Fields("YARD_ENDEN") = "" Or objRs.Fields("YARD_ENDEN") = " ") Then

              .cmd_CAMPB.Enabled = True
         
           End If

           .txtBUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
           If .txtBUMOC = "8" Then
              .txtAddress = DLookup("CONT_DISKBACK", "dbo_CONT_MAST", "CONT_KEY = 1") & Val(.txt_ヤードコード) & ".htm"
           Else
              .txtAddress = DLookup("CONT_DISKBACK", "dbo_CONT_MAST", "CONT_KEY = 1") & .txt_ヤードコード & ".htm"
           End If
           End With
        End If                                                                  'INSERT 20041106 N.MIURA ENDED
        
        'strSQL = ""                                                            'DELETE 20041106 N.MIURA START
        'strSQL = " SELECT"
        'strSQL = strSQL & " YARD_NAME,"
        'strSQL = strSQL & " YARD_ADDR_1,"
        'strSQL = strSQL & " YARD_ADDR_2 "
        'strSQL = strSQL & " FROM dbo_YARD_MAST"
        'strSQL = strSQL & " Where"
        'strSQL = strSQL & " YARD_CODE =" & Val(txt_ヤードコード)
        'Set objRS = objDB.OpenRecordset(strSQL, dbOpenSnapshot)
        'If Not (objRS.BOF And objRS.EOF) Then
        '    Me.txt_ヤード名 = objRS![YARD_NAME]
        '    vRET2 = objRS![YARD_ADDR_1]
        '    vRet = objRS![YARD_ADDR_2]
        '    Me.txt_ヤード住所 = vRET2 & vRet
        '    Me.txt_備考 = objRS![YARD_NOTE]
        '    Me.cmd_CAMPB.Enabled = True
        '    Me.txtBUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")                                             'INSERT 20040710 N.MIURA START
        '    If Me.txtBUMOC = "8" Then
        '       Me.txtAddress = DLookup("CONT_DISKBACK", "dbo_CONT_MAST", "CONT_KEY = 1") & Val(Me.txt_ヤードコード) & ".htm"
        '    Else
        '       Me.txtAddress = DLookup("CONT_DISKBACK", "dbo_CONT_MAST", "CONT_KEY = 1") & Me.txt_ヤードコード & ".htm"
        '    End If                                                                                                           'INSERT 20040710 N.MIURA ENDED
        'End If                                                                 'DELETE 20041106 N.MIURA ENDED
      
    End If
    
    GetYardData = True
    GoTo Exit_rtn

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set objRs = Nothing
    Exit Function
    
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Form_Unload

    
    Call DELWK_RTN
    
    Exit Sub
    
Err_Form_Unload:
    MsgBox Error$
    Exit Sub

End Sub

Private Sub サブ下_Enter()
On Error Resume Next
    doCmd.GoToRecord , , acFirst
    bActiveDown = True
End Sub

Private Sub サブ下_Exit(Cancel As Integer)

    bActiveDown = False

End Sub

Private Sub サブ上_Enter()
On Error Resume Next
    doCmd.GoToRecord , , acFirst
    bActiveUp = True
End Sub

Private Sub サブ上_Exit(Cancel As Integer)
    
    bActiveUp = False

End Sub

Private Sub サブ上A_Enter()
On Error Resume Next
    doCmd.GoToRecord , , acFirst
End Sub

Sub SubFormSetFocus()
    
    '2005/05/29 MITSUNO
    '予約名/電話番号変更時、予約トランに受付番号が入っていたら
    'ｺﾝﾃﾅﾏｽﾀへUpdateしない処理追加の為、以下の処理をコメント
    '上下段のデータ有無チェック
    If CheckCntaData = False Then
        Exit Sub
    End If

    Select Case vStep
    Case 0  '上下段ともデータ無し
        Me.cmd_取消.SetFocus
    Case 1  '上段のみデータあり
        Me.サブ上.SetFocus
        Me.サブ上.Form!cmdUke.SetFocus
        'Me.サブ上.Form!txt_予約客先名.SetFocus
    Case 2  '下段のみデータあり
        Me.サブ下.SetFocus
        Me.サブ下.Form!cmdUke.SetFocus
        'Me.サブ下.Form!txt_予約客先名.SetFocus
    Case 3  '上下段ともにデータあり
        Me.サブ下.SetFocus
        Me.サブ下.Form!cmdUke.SetFocus
        'Me.サブ下.Form!txt_予約客先名.SetFocus
        Me.サブ上.SetFocus
        Me.サブ上.Form!cmdUke.SetFocus
        'Me.サブ上.Form!txt_予約客先名.SetFocus
    End Select


End Sub

'==============================================================================*
'
'        MODULE_NAME      :確認と保存
'        MODULE_ID        :SaveData
'        CREATE_DATE      :20050718 K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function SaveData() As Boolean
    If (Nz(Me.txt_ヤードコード) <> "") And bChangeFlag_FVS820 Then
        If MsgBox("修正中の内容を保存します。よろしいですか？", vbQuestion + vbYesNo, strMsg_kaku) = vbYes Then
            SaveData = WRI_RTN()
            Exit Function
        End If
    End If
    SaveData = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :編集中に設定
'        MODULE_ID        :SetChangeFlag
'        CREATE_DATE      :2005/07/18 K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetChangeFlag()
    If bChangeFlag_FVS820 = False Then
        Me.Caption = strCaption & "(修正中)"
        bChangeFlag_FVS820 = True
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :外部から再表示させる
'        MODULE_ID        :Reload
'        CREATE_DATE      :2005/06/03 K.I
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub Reload()                                                            'DELETE 20050718 K.ISHIZAKA
'ここから                                                                       'INSERT 20050718 K.ISHIZAKA
Public Sub Reload(ByVal strYard1 As String, ByVal strYard2 As String)
    If bChangeFlag_FVS820 Then
        Exit Sub
    End If
    If (strYard1 <> Me.txt_ヤードコード) And (strYard2 <> Me.txt_ヤードコード) Then
        Exit Sub
    End If
'ここまで                                                                       'INSERT 20050718 K.ISHIZAKA
    'ワークテーブル削除
    Call DELWK_RTN
    '表示データ作成
    Call MakeDisplay
End Sub
