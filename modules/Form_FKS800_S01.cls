VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS800_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス（サブ）
'   プログラムＩＤ　：　FKS800_S01
'   作　成　日　　　：  2008/06/02
'   作　成　者　　　：  iizuka
'   Ver             ：  0.0
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス（サブ）
'   プログラムＩＤ　：　FKS800_S01
'   作　成　日　　　：  2013/11/22
'   作　成　者　　　：  MIYAMOTO
'   Ver             ：  0.1
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス（サブ）
'   プログラムＩＤ　：　FKS800_S01
'   作　成　日　　　：  2014/01/17
'   作　成　者　　　：  MIYAMOTO
'   Ver             ：  0.2
'**********************************************
Option Compare Database
Option Explicit

Private Const PROG_ID = "FKS800_S01"
Private Const P_ROOM_WORK_TABLE         As String = "WK_FKS800"
Private Const P_BOX_LIST_WORK_TABLE     As String = "WK_FKS800_S01_TEMP"

Private pobjDB                       As Object

Private pstrBumonCode                As String       ' 部門コード

'==============================================================================*
'
'        MODULE_NAME      :cmd_移動_Click
'        MODULE_ID        :cmd_移動_Click
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_移動_Click()
    
    Dim anStrParameter As String
    Dim isException As Boolean
    isException = False

On Error GoTo Exception
    
'↓UPDATE 2013/11/22 MIYAMOTO
'    If Me.txt_状況.VALUE = "撤去" Then
    If Form_FKS800.cmb_状況.VALUE = "撤去" Then
'↑UPDATE 2013/11/22 MIYAMOTO
        MsgBox "撤去したボックスは修正できません。", vbInformation, "確認"
    Else
        
        Call Form_FKS800.dropWorkTable(2)
        Call Form_FKS800.createWorkTable(2)
        
        If fncKeiyakuCheck("移動") Then
            anStrParameter = ""
            anStrParameter = setParameter & ",移動登録,2"
            doCmd.OpenForm "FKS810", , , , , acDialog, anStrParameter
        End If
    End If
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_移動_Click", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_削除_Click
'        MODULE_ID        :cmd_削除_Click
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_削除_Click()
    
'↓INSERT 2013/11/22 MIYAMOTO
    Dim strBoxNo           As String
    Dim strEdaNo           As String
'↑INSERT 2013/11/22 MIYAMOTO
    Dim objAdoDbConnection As Object
    Dim rsData             As Object       'ADODB.Recordset
    Dim strSQL             As String
    Dim strMsg             As String
    Dim isException        As Boolean
    isException = False

On Error GoTo Exception
    
'↓INSERT 2013/11/22 MIYAMOTO
'↓UPDATE 2014/01/17 MIYAMOTO
'    strBoxNo = Mid(Me.txt_ボックス番号, 1, InStr(Me.txt_ボックス番号, " ") - 1)
'    strEdaNo = Mid(Me.txt_ボックス番号, InStr(Me.txt_ボックス番号, " ") + 1)
    strBoxNo = Me.txt_ボックス番号.VALUE
    strEdaNo = Me.txt_枝番.VALUE
'↑UPDATE 2014/01/17 MIYAMOTO
'↑INSERT 2013/11/22 MIYAMOTO
    
'↓UPDATE 2013/11/22 MIYAMOTO
'    If Me.txt_状況.VALUE = "撤去" Then
    If Form_FKS800.cmb_状況.VALUE = "撤去" Then
'↑UPDATE 2013/11/22 MIYAMOTO
        MsgBox "撤去したボックスは修正できません。", vbInformation, "確認"
    Else
        '※現状、削除では契約中のチェックは行わない。よってコメントとする。
        'If fncKeiyakuCheck("削除") Then
            strSQL = ""
            strSQL = strSQL & " SELECT * FROM CBYS_RECD "
'↓UPDATE 2013/11/22 MIYAMOTO
'            strSQL = strSQL & " WHERE CBYSR_BOXNO = '" & Me.txt_ボックス番号 & "';"
            strSQL = strSQL & " WHERE CBYSR_BOXNO = '" & strBoxNo & "' "
            strSQL = strSQL & " AND CBYSR_EDANO = '" & strEdaNo & "';"
'↑UPDATE 2013/11/22 MIYAMOTO
            Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
            Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)
            
            If rsData.EOF Then
                strMsg = "データを削除します。よろしいですか？"
            Else
                strMsg = "このボックスを移動前のヤードに戻します。よろしいですか？"
            End If
            
            rsData.Close
            objAdoDbConnection.Close
            Set rsData = Nothing
            Set objAdoDbConnection = Nothing
            
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, "確認") = vbYes Then
                '削除
                Call DEL_RTN
                Call Form_FKS800.createBoxData
                Call Form_FKS800.Requery
            End If
        'End If
    End If
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If Not rsData Is Nothing Then: rsData.Close: Set rsData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    If isException Then Call MSZZ024_M00("cmd_削除_Click", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_修正_Click
'        MODULE_ID        :cmd_修正_Click
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_修正_Click()
    
    Dim strSQL         As String
    Dim anStrParameter As String
    Dim isException    As Boolean
    isException = False

On Error GoTo Exception
    
'↓UPDATE 2013/11/22 MIYAMOTO
'    If Me.txt_状況.VALUE = "撤去" Then
    If Form_FKS800.cmb_状況.VALUE = "撤去" Then
'↑UPDATE 2013/11/22 MIYAMOTO
        MsgBox "撤去したボックスは修正できません。", vbInformation, "確認"
    Else
        Call createRoomWorkData
        
        anStrParameter = ""
        anStrParameter = setParameter & ",修正,1"
        doCmd.OpenForm "FKS810", , , , , acDialog, anStrParameter
    End If
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_修正_Click", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ダイアログ引数作成
'        MODULE_ID        :setParameter
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function setParameter() As String
    
    Dim strParameter As String
    
    On Error GoTo Exception
    
    strParameter = ""
    strParameter = Nz(Forms![FKS800]![cmb_ヤード])
    strParameter = strParameter & "," & Nz(Forms![FKS800]![lbl_ヤード名].Caption)
    strParameter = strParameter & "," & Nz(Forms![FKS800]![lbl_営業開始日].Caption)
    strParameter = strParameter & "," & Nz(Forms![FKS800]![lbl_営業終了日].Caption)
    strParameter = strParameter & "," & Nz(Forms![FKS800]![lbl_ヤード解約予定日].Caption)
'↓UPDATE 2014/01/17 MIYAMOTO
'    strParameter = strParameter & "," & Nz(Me.txt_ボックス番号.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_ボックス番号.VALUE) & " " & Nz(Me.txt_枝番.VALUE)
'↑UPDATE 2014/01/17 MIYAMOTO
    strParameter = strParameter & "," & Nz(Me.txt_タイプ_値.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_タイプ.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_内装_値.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_内装.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_サイズ_値.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_サイズ.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_ドア数_値.VALUE)
    strParameter = strParameter & "," & Replace(Nz(Me.txt_ドア数.VALUE), ",", "#")
    strParameter = strParameter & "," & Nz(Me.txt_ドア位置_値.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_ドア位置.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_ドアタイプ_値.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_ドアタイプ.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_入庫日.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_設置日.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_設置階層_値.VALUE)
    strParameter = strParameter & "," & Nz(Me.txt_備考.VALUE)
    
    setParameter = strParameter
    
    Exit Function

Exception:
  Call Err.Raise(Err.Number, "setParameter" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
  
End Function

'==============================================================================*
'
'        MODULE_NAME      :部屋割付データ作成
'        MODULE_ID        :createRoomWorkData
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub createRoomWorkData()
    
'↓INSERT 2013/11/22 MIYAMOTO
    Dim strBoxNo           As String
    Dim strEdaNo           As String
'↑INSERT 2013/11/22 MIYAMOTO
    Dim objAdoDbConnection As Object
    Dim rsData             As Object       'ADODB.Recordset
    Dim rsRoomData         As Recordset
    Dim strSQL             As String
    
    On Error GoTo Exception

    Call Form_FKS800.dropWorkTable(2)
    Call Form_FKS800.createWorkTable(2)
    
'↓INSERT 2013/11/22 MIYAMOTO
'↓UPDATE 2014/01/17 MIYAMOTO
'    strBoxNo = Mid(Nz(Me.txt_ボックス番号.VALUE), 1, InStr(Nz(Me.txt_ボックス番号.VALUE), " ") - 1)
'    strEdaNo = Mid(Nz(Me.txt_ボックス番号.VALUE), InStr(Nz(Me.txt_ボックス番号.VALUE), " ") + 1)
    strBoxNo = Me.txt_ボックス番号.VALUE
    strEdaNo = Me.txt_枝番.VALUE
'↑UPDATE 2014/01/17 MIYAMOTO
'↑INSERT 2013/11/22 MIYAMOTO
    
    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & "        CBRA_TRAN.CBRAT_CNO AS CNO"
    strSQL = strSQL & "       ,NAME_MAST1.NAME_RYAK AS NAME1"
    strSQL = strSQL & "       ,CNTA_MAST.CNTA_SIZE AS SIZE"
    strSQL = strSQL & "       ,NAME_MAST2.NAME_RYAK AS NAME2"
    strSQL = strSQL & "   FROM CBRA_TRAN , "
    strSQL = strSQL & "        CNTA_MAST "
    strSQL = strSQL & "             LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CNTA_MAST.CNTA_USAGE AND (NAME_MAST1.NAME_CODE <> 10) AND (NAME_MAST1.NAME_CODE <> 30) AND (NAME_MAST1.NAME_ID = '086')"
    strSQL = strSQL & "             LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CNTA_MAST.CNTA_STEP AND NAME_MAST2.NAME_ID = '090'"
    strSQL = strSQL & "  WHERE CBRA_TRAN.CBRAT_CNO = CNTA_MAST.CNTA_NO"
    strSQL = strSQL & "    AND CBRA_TRAN.CBRAT_YCODE = CNTA_MAST.CNTA_CODE"
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "    AND CBRA_TRAN.CBRAT_BOXNO = '" & Nz(Me.txt_ボックス番号.VALUE)
    strSQL = strSQL & "    AND CBRA_TRAN.CBRAT_BOXNO = '" & strBoxNo
    strSQL = strSQL & "'   AND CBRA_TRAN.CBRAT_EDANO = '" & strEdaNo
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & "'   AND CBRA_TRAN.CBRAT_YCODE = " & Nz(Forms![FKS800]![cmb_ヤード])
    strSQL = strSQL & " ;"
    Set rsRoomData = pobjDB.OpenRecordset(P_ROOM_WORK_TABLE, dbOpenDynaset, dbAppendOnly)
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)
    
    With rsData
        Do Until .EOF
            rsRoomData.AddNew
            rsRoomData.Fields("CNO") = .Fields("CNO")
            rsRoomData.Fields("TEXTDATA") = Format(Nz(.Fields("CNO")), "000000") & " " & _
                                                    .Fields("NAME1") & " " & _
                                            Format(Nz(.Fields("SIZE")), "0.00") & "帖 " & _
                                                    .Fields("NAME2")
            rsRoomData.UPDATE
            .MoveNext
        Loop
        rsRoomData.Close
        rsData.Close
        objAdoDbConnection.Close
        Set rsData = Nothing
        Set rsRoomData = Nothing
        Set objAdoDbConnection = Nothing
    End With

    Exit Sub

Exception:
    If Not rsData Is Nothing Then: rsData.Close: Set rsData = Nothing
    If Not rsRoomData Is Nothing Then: rsRoomData.Close: Set rsRoomData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    Call Err.Raise(Err.Number, "createRoomWorkData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
  
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_複写_Click
'        MODULE_ID        :cmd_複写_Click
'        NOTE             :
'        CREATE_DATE      :2019/04/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_複写_Click()
    
    Dim strSQL         As String
    Dim anStrParameter As String
    Dim isException    As Boolean
    isException = False

On Error GoTo Exception
    
    If Form_FKS800.cmb_状況.VALUE = "撤去" Then
        MsgBox "撤去したボックスは修正できません。", vbInformation, "確認"
    Else
        Call createRoomWorkData
        
        anStrParameter = ""
        anStrParameter = setParameter & ",複写,3"
        doCmd.OpenForm "FKS810", , , , , acDialog, anStrParameter
    End If
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_複写_Click", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        MODULE_ID        :Form_Load
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    
    Dim isException As Boolean
    isException = False

On Error GoTo Exception
    
    pstrBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If pstrBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    Set pobjDB = CurrentDb
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Load", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Close
'        MODULE_ID        :Form_Close
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    
    Dim isException As Boolean
    isException = False

On Error GoTo Exception
    
    'コネクションクローズ
    If Not pobjDB Is Nothing Then
       pobjDB.Close
       Set pobjDB = Nothing
    End If
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Close", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :契約状態チェック（移動時）
'        MODULE_ID        :fncKeiyakuCheck
'        NOTE             :
'        CREATE_DATE      :2008/06/20
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncKeiyakuCheck(strMode As String) As Boolean
  
'↓INSERT 2013/11/22 MIYAMOTO
    Dim strBoxNo           As String
    Dim strEdaNo           As String
'↑INSERT 2013/11/22 MIYAMOTO
    Dim objAdoDbConnection As Object
    Dim strSQL             As String
    Dim rsData             As Object
    Dim rsDataCno          As Object
    Dim bError             As Boolean

    fncKeiyakuCheck = False
    bError = False

    On Error GoTo Exception

    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)

'↓INSERT 2013/11/22 MIYAMOTO
'↓UPDATE 2014/01/17 MIYAMOTO
'    strBoxNo = Mid(Nz(Me.txt_ボックス番号.VALUE), 1, InStr(Nz(Me.txt_ボックス番号.VALUE), " ") - 1)
'    strEdaNo = Mid(Nz(Me.txt_ボックス番号.VALUE), InStr(Nz(Me.txt_ボックス番号.VALUE), " ") + 1)
    strBoxNo = Me.txt_ボックス番号.VALUE
    strEdaNo = Me.txt_枝番.VALUE
'↑UPDATE 2014/01/17 MIYAMOTO
'↑INSERT 2013/11/22 MIYAMOTO
    
    strSQL = strSQL & " SELECT CBRA_TRAN.CBRAT_CNO AS CNO"
    strSQL = strSQL & "  FROM CBRA_TRAN"
    strSQL = strSQL & "      ,CNTA_MAST "
    strSQL = strSQL & " WHERE CBRA_TRAN.CBRAT_CNO = CNTA_MAST.CNTA_NO"
    strSQL = strSQL & "   AND CBRA_TRAN.CBRAT_YCODE = CNTA_MAST.CNTA_CODE"
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "   AND CBRA_TRAN.CBRAT_BOXNO = '" & Nz(Me.txt_ボックス番号.VALUE) & "'"
    strSQL = strSQL & "   AND CBRA_TRAN.CBRAT_BOXNO = '" & strBoxNo & "'"
    strSQL = strSQL & "   AND CBRA_TRAN.CBRAT_EDANO = '" & strEdaNo & "'"
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   AND CBRA_TRAN.CBRAT_YCODE = '" & Nz(Forms![FKS800]![cmb_ヤード]) & "'"
    Set rsDataCno = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    If Not rsDataCno.EOF Then
        While Not rsDataCno.EOF And bError = False
            strSQL = ""
            strSQL = strSQL & "SELECT * FROM CARG_FILE "
            strSQL = strSQL & " WHERE CARG_YCODE = " & CLng(Nz(Forms![FKS800]![cmb_ヤード]))
            strSQL = strSQL & " AND CARG_NO = " & rsDataCno.Fields("CNO")
            strSQL = strSQL & " ORDER BY CARG_FSDATE DESC, CARG_UPDAD DESC, CARG_UPDAJ DESC, CARG_INSED DESC, CARG_INSEJ DESC "
            Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
            '最新のコンテナ契約を参照
            If Not rsData.EOF Then
                If rsData.Fields("CARG_AGRE") = 1 Or rsData.Fields("CARG_AGRE") = 2 Or rsData.Fields("CARG_AGRE") = 3 Then
                    '契約中データあり
                    MsgBox rsDataCno.Fields("CNO") & "が契約中のため" & strMode & "は出来ません｡" & vbCrLf & "解約入力で「メンテ中」または「完了」にして下さい。", vbInformation, "確認"
                    bError = True
                End If
            End If
            rsData.Close
            Set rsData = Nothing
            rsDataCno.MoveNext
        Wend
    End If
    rsDataCno.Close
    Set rsDataCno = Nothing
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

    fncKeiyakuCheck = Not bError

Exception:
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If
    If Not rsDataCno Is Nothing Then
        rsDataCno.Close
        Set rsDataCno = Nothing
    End If
    If Not objAdoDbConnection Is Nothing Then
        objAdoDbConnection.Close
        Set objAdoDbConnection = Nothing
    End If
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "fncKeiyakuCheck" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      : 削除処理
'        MODULE_ID        : DEL_RTN
'        PARAMETER        : 引数なし
'        戻り値           : True=正常 False=異常
'        CREATE_DATE      : 2008/**/**
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DEL_RTN() As Boolean
  
    Dim objAdoDbConnection   As Object
    Dim strSQL               As String
    Dim rsData               As Object
    Dim rsDataDel            As Object     'マスタ、トラン削除用
    Dim rsDataIns            As Object     'トラン作成用
    Dim strDate              As String     '初期日付取得用
    Dim strTime              As String     '初期時刻取得用
    Dim strBoxNo             As String     '選択ボックス番号得用
'↓INSERT 2013/11/22 MIYAMOTO
    Dim strEdaNo             As String
'↑INSERT 2013/11/22 MIYAMOTO
    Dim strWkDate            As String     'コンテナボックスヤード設置履歴トラン設置日取得用

On Error GoTo Exception

    DEL_RTN = False
    
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    objAdoDbConnection.BeginTrans

    '初期値設定
    strDate = Format(DATE, "yyyymmdd")
    strTime = Format(Now, "hhmmss")
'↓UPDATE 2013/11/22 MIYAMOTO
'    strBoxNo = Me.txt_ボックス番号
'↓UPDATE 2014/01/17 MIYAMOTO
'    strBoxNo = Mid(Me.txt_ボックス番号, 1, InStr(Me.txt_ボックス番号, " ") - 1)
'    strEdaNo = Mid(Me.txt_ボックス番号, InStr(Me.txt_ボックス番号, " ") + 1)
    strBoxNo = Me.txt_ボックス番号.VALUE
    strEdaNo = Me.txt_枝番.VALUE
'↑UPDATE 2014/01/17 MIYAMOTO
'↑UPDATE 2013/11/22 MIYAMOTO
    
    'コンテナボックスヤード設置履歴トランデータ存在チェック
    strSQL = ""
    strSQL = strSQL & " SELECT * FROM CBYS_RECD             "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & " WHERE CBYSR_BOXNO = '" & strBoxNo & "'  "
    strSQL = strSQL & " WHERE CBYSR_BOXNO = '" & strBoxNo & "'  "
    strSQL = strSQL & " AND CBYSR_EDANO = '" & strEdaNo & "'  "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " ORDER BY CBYSR_SETDATE DESC"

    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)

    If rsData.EOF Then
        'コンテナボックスヤード設置履歴トランに選択ボックス番号が存在しない場合、マスタ、トランを全て物理削除する
        'コンテナボックスマスタ削除
        '※コンテナボックスマスタ部屋割付履歴も残っていたら全て削除
        strSQL = ""
        strSQL = strSQL & " SELECT * FROM CBOX_MAST             "
'↓UPDATE 2013/11/22 MIYAMOTO
'        strSQL = strSQL & " WHERE CBOXM_BOXNO = '" & strBoxNo & "'       "
        strSQL = strSQL & " WHERE CBOXM_BOXNO = '" & strBoxNo & "'       "
        strSQL = strSQL & " AND CBOXM_EDANO = '" & strEdaNo & "'         "
'↑UPDATE 2013/11/22 MIYAMOTO

        Set rsDataDel = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)

        rsDataDel.Delete
        
        rsDataDel.Close
        Set rsDataDel = Nothing
        
        'コンテナボックスヤード設置トラン削除
        strSQL = ""
        strSQL = strSQL & " SELECT * FROM CBYS_TRAN             "
'↓UPDATE 2013/11/22 MIYAMOTO
'        strSQL = strSQL & " WHERE CBYST_BOXNO = '" & strBoxNo & "'       "
        strSQL = strSQL & " WHERE CBYST_BOXNO = '" & strBoxNo & "'       "
        strSQL = strSQL & " AND CBYST_EDANO = '" & strEdaNo & "'         "
'↑UPDATE 2013/11/22 MIYAMOTO

        Set rsDataDel = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)

        rsDataDel.Delete
        
        rsDataDel.Close
        Set rsDataDel = Nothing
        
        'コンテナボックス部屋割付トラン削除
        strSQL = ""
        strSQL = strSQL & " SELECT * FROM CBRA_TRAN             "
'↓UPDATE 2013/11/22 MIYAMOTO
'        strSQL = strSQL & " WHERE CBRAT_BOXNO = '" & strBoxNo & "'       "
        strSQL = strSQL & " WHERE CBRAT_BOXNO = '" & strBoxNo & "'       "
        strSQL = strSQL & " AND CBRAT_EDANO = '" & strEdaNo & "'         "
'↑UPDATE 2013/11/22 MIYAMOTO

        Set rsDataDel = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)

        If Not rsDataDel.EOF Then

            While Not rsDataDel.EOF
                rsDataDel.Delete
                rsDataDel.MoveNext
            Wend

        End If

        rsDataDel.Close
        Set rsDataDel = Nothing

        'コンテナボックス部屋割付履歴トラン削除
        strSQL = ""
        strSQL = strSQL & " SELECT * FROM CBRA_RECD                     "
'↓UPDATE 2013/11/22 MIYAMOTO
'        strSQL = strSQL & " WHERE CBRAR_BOXNO = '" & strBoxNo & "'      "
        strSQL = strSQL & " WHERE CBRAR_BOXNO = '" & strBoxNo & "'      "
        strSQL = strSQL & " AND CBRAR_EDANO = '" & strEdaNo & "'        "
'↑UPDATE 2013/11/22 MIYAMOTO

        Set rsDataDel = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)

        If Not rsDataDel.EOF Then

            While Not rsDataDel.EOF
                rsDataDel.Delete
                rsDataDel.MoveNext
            Wend

        End If

        rsDataDel.Close
        Set rsDataDel = Nothing

        rsData.Close
        Set rsData = Nothing
    Else
        'コンテナボックスヤード設置履歴トランに選択ボックス番号が存在した場合、各履歴トランより現トランへ情報を戻す
        'コンテナボックスヤード設置履歴トラン設置日セット
        strWkDate = rsData.Fields("CBYSR_SETDATE")

        'コンテナボックスヤード設置トランを削除し、コンテナボックスヤード設置履歴トランより
        'コンテナボックスヤード設置トランを作成する
        strSQL = ""
        strSQL = strSQL & " SELECT * FROM CBYS_TRAN             "
'↓UPDATE 2013/11/22 MIYAMOTO
'        strSQL = strSQL & " WHERE CBYST_BOXNO = '" & strBoxNo & "'       "
        strSQL = strSQL & " WHERE CBYST_BOXNO = '" & strBoxNo & "'       "
        strSQL = strSQL & " AND CBYST_EDANO = '" & strEdaNo & "'         "
'↑UPDATE 2013/11/22 MIYAMOTO

        Set rsDataDel = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)

        rsDataDel.Delete
        
        rsDataDel.Close
        Set rsDataDel = Nothing
  
        'コンテナボックスヤード設置トラン登録
        Set rsDataIns = MSZZ025.ADODB_Recordset("CBYS_TRAN", objAdoDbConnection, adoAppendOnly)

        rsDataIns.AddNew
        rsDataIns.Fields("CBYST_BOXNO") = rsData.Fields("CBYSR_BOXNO")
'↓INSERT 2013/11/22 MIYAMOTO
        rsDataIns.Fields("CBYST_EDANO") = rsData.Fields("CBYSR_EDANO")
'↑INSERT 2013/11/22 MIYAMOTO
        rsDataIns.Fields("CBYST_SYUBETSU") = rsData.Fields("CBYSR_SYUBETSU")
        rsDataIns.Fields("CBYST_YCODE") = rsData.Fields("CBYSR_YCODE")
        rsDataIns.Fields("CBYST_STATUS") = 0
        rsDataIns.Fields("CBYST_SETDATE") = rsData.Fields("CBYSR_SETDATE")
        rsDataIns.Fields("CBYST_SETFLOOR") = rsData.Fields("CBYSR_SETFLOOR")
        rsDataIns.Fields("CBYST_BIKON") = rsData.Fields("CBYSR_BIKON")
        rsDataIns.Fields("CBYST_INSED") = strDate
        rsDataIns.Fields("CBYST_INSEJ") = strTime
        rsDataIns.Fields("CBYST_INSPB") = PROG_ID
        rsDataIns.Fields("CBYST_INSUB") = Form_FKS800.gstrUserName
        rsDataIns.UPDATE

        rsDataIns.Close
        Set rsDataIns = Nothing

        rsData.Delete

        rsData.Close
        Set rsData = Nothing

        'コンテナボックス部屋割付履歴トランデータ存在チェック
        strSQL = ""
        strSQL = strSQL & " SELECT * FROM CBRA_RECD             "
'↓UPDATE 2013/11/22 MIYAMOTO
'        strSQL = strSQL & " WHERE CBRAR_BOXNO = '" & strBoxNo & "'      "
        strSQL = strSQL & " WHERE CBRAR_BOXNO = '" & strBoxNo & "'      "
        strSQL = strSQL & " AND CBRAR_EDANO = '" & strEdaNo & "'        "
'↑UPDATE 2013/11/22 MIYAMOTO
        strSQL = strSQL & "   AND CBRAR_SETDATE = '" & strWkDate & "'     "

        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)

        'コンテナボックス部屋割付トランを削除し、コンテナボックス部屋割付履歴トランより
        'コンテナボックス部屋割付トランを作成する。但し、コンテナボックス部屋割付履歴
        'トランに履歴がない場合は、コンテナボックス部屋割付トランの削除のみとする。
        strSQL = ""
        strSQL = strSQL & " SELECT * FROM CBRA_TRAN             "
'↓UPDATE 2013/11/22 MIYAMOTO
'        strSQL = strSQL & " WHERE CBRAT_BOXNO = '" & strBoxNo & "'      "
        strSQL = strSQL & " WHERE CBRAT_BOXNO = '" & strBoxNo & "'      "
        strSQL = strSQL & " AND CBRAT_EDANO = '" & strEdaNo & "'        "
'↑UPDATE 2013/11/22 MIYAMOTO

        Set rsDataDel = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)

        If Not rsDataDel.EOF Then

            While Not rsDataDel.EOF
                rsDataDel.Delete
                rsDataDel.MoveNext
            Wend
 
        End If

        rsDataDel.Close
        Set rsDataDel = Nothing

        If Not rsData.EOF Then
  
            'コンテナボックス部屋割付トラン登録
            Set rsDataIns = MSZZ025.ADODB_Recordset("CBRA_TRAN", objAdoDbConnection, adoAppendOnly)
    
            While Not rsData.EOF
                rsDataIns.AddNew
                rsDataIns.Fields("CBRAT_BOXNO") = rsData.Fields("CBRAR_BOXNO")
'↓INSERT 2013/11/22 MIYAMOTO
                rsDataIns.Fields("CBRAT_EDANO") = rsData.Fields("CBRAR_EDANO")
'↑INSERT 2013/11/22 MIYAMOTO
                rsDataIns.Fields("CBRAT_YCODE") = rsData.Fields("CBRAR_YCODE")
                rsDataIns.Fields("CBRAT_CNO") = rsData.Fields("CBRAR_CNO")
                rsDataIns.Fields("CBRAT_STATUS") = 0
                rsDataIns.Fields("CBRAT_SETDATE") = rsData.Fields("CBRAR_SETDATE")
                rsDataIns.Fields("CBRAT_INSED") = strDate
                rsDataIns.Fields("CBRAT_INSEJ") = strTime
                rsDataIns.Fields("CBRAT_INSPB") = PROG_ID
                rsDataIns.Fields("CBRAT_INSUB") = Form_FKS800.gstrUserName
                rsDataIns.UPDATE
                rsData.Delete
                rsData.MoveNext
            Wend
    
            rsDataIns.Close
            Set rsDataIns = Nothing

        End If

        rsData.Close
        Set rsData = Nothing

    End If

    DEL_RTN = True
    
    objAdoDbConnection.CommitTrans
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

Exception:

    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If

    If Not rsDataDel Is Nothing Then
        rsDataDel.Close
        Set rsDataDel = Nothing
    End If

    If Not rsDataIns Is Nothing Then
        rsDataIns.Close
        Set rsDataIns = Nothing
    End If

    If Err <> 0 Then
        objAdoDbConnection.RollbackTrans
        objAdoDbConnection.Close
        Set objAdoDbConnection = Nothing
        Call Err.Raise(Err.Number, "DEL_RTN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
