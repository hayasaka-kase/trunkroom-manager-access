Attribute VB_Name = "AccessGlobal"
Option Compare Database
Option Explicit
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/_/          テーブルの再リンク                                           _/_/_/_/_/
'_/_/_/_/_/_/ iniファイルのファイル名はグローバルで宣言してるので使用する時は修正してﾈ  _/_/_/_/_/
'_/_/_/_/_/_/ ｴﾗｰ発生時はｼｽﾃﾑを強制終了します。                                      _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Sub CmRefreshLink(stMDBPath As String, stMDBName As String)

  On Error GoTo Err_CmRefreshLink

  Dim lojWksp         As Workspace          '_/_/_/ ﾜｰｸｽﾍﾟｰｽ変数
  Dim lojdb           As Database           '_/_/_/ ﾃﾞｰﾀﾍﾞｰｽ変数
  Dim lojtable        As TableDef           '_/_/_/ TableDefｵﾌﾞｼﾞｪｸﾄ変数
  Dim lstConnect      As String             '_/_/_/ ﾘﾝｸ文字列変数
  Dim inLpCnt         As Integer            '_/_/_/ ﾙｰﾌﾟｶｳﾝﾄ変数
  Dim linRtnValue     As Integer            '_/_/_/ 関数の返り値を格納する変数
  Dim linTableCnt     As Integer            '_/_/_/ ﾃｰﾌﾞﾙ数格納変数

  On Error GoTo Err_RollBack

'/****** ﾜｰｸｽﾍﾟｰｽ及びﾃﾞｰﾀﾍﾞｰｽを設定
  Set lojWksp = DBEngine(0)
  Set lojdb = lojWksp(0)

'/****** ﾃｰﾌﾞﾙ数を取得
  linTableCnt = lojdb.TableDefs.Count

'/****** ｽﾃｰﾀｽﾊﾞｰの進行状況ｲﾝｼﾞｹｰﾀの初期化
  linRtnValue = SysCmd(acSysCmdInitMeter, "ﾃｰﾌﾞﾙﾘﾝｸ中...", linTableCnt)

'/****** ﾄﾗﾝｻﾞｸｼｮﾝの開始
  lojWksp.BeginTrans

'/****** 変数の初期化
  inLpCnt = 0
  
  For Each lojtable In lojdb.TableDefs
    
'/****** ﾃｰﾌﾞﾙの接続文字列の判定
    If lojtable.Connect <> "" And Not IsNull(lojtable.Connect) Then
      lstConnect = ""
      lstConnect = ";DATABASE=" & stMDBPath & stMDBName
      lstConnect = lstConnect & ";Table=" & lojtable.SourceTableName
'/****** ﾃｰﾌﾞﾙの再ﾘﾝｸ
      lojtable.Connect = lstConnect
      lojtable.RefreshLink
    End If
    inLpCnt = inLpCnt + 1
'/****** ｽﾃｰﾀｽﾊﾞｰｲﾝｼﾞｹｰﾀｰの更新
    linRtnValue = SysCmd(acSysCmdUpdateMeter, inLpCnt)
  Next
'/****** ｺﾐｯﾄ
  lojWksp.CommitTrans
  lojdb.Close
  lojWksp.Close

'/****** ｽﾃｰﾀｽﾊﾞｰﾃｷｽﾄの消去
  linRtnValue = SysCmd(acSysCmdSetStatus, " ")

'/****** ｽﾃｰﾀｽﾊﾞｰｸﾘｱ
  linRtnValue = SysCmd(acSysCmdClearStatus)
  
  Exit Sub

Err_RollBack:
'/****** RollBack ｾｸｼｮﾝ
  doCmd.Hourglass False
  lojWksp.Rollback
  lojdb.Close
  lojWksp.Close
  doCmd.Quit acQuitSaveAll
  Exit Sub

Err_CmRefreshLink:
'/****** ｴﾗｰ時
  doCmd.Hourglass False
  MsgBox Err.Description
  doCmd.Quit acQuitSaveAll
  Exit Sub

End Sub

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/                                                                                            _/_/_/_/_/
'_/_/_/_/_/ OpenしているFormを全て閉じる                                                                _/_/_/_/_/
'_/_/_/_/_/                                                                                            _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Sub CloseForm()

  Dim ojFrm     As Form         '_/_/_/ ﾌｫｰﾑ変数
  Dim inLpCnt   As Integer      '_/_/_/ ﾙｰﾌﾟ回数

'_/_/_/ ｱｸﾃｨﾌﾞなﾌｫｰﾑが０の場合は関数を抜ける
  If Forms.Count = 0 Then
    Exit Sub
  End If

'_/_/_/ ｱｸﾃｨﾌﾞなﾌｫｰﾑの数だけﾙｰﾌﾟ
  For inLpCnt = 0 To Forms.Count - 1
'_/_/_/ ｱｸﾃｨﾌﾞなﾌｫｰﾑのIDは閉じる度にReNumberされるため常にIDに０を指定
    doCmd.Close acForm, Forms(0)
  Next

End Sub
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/ ｵｰﾌﾟﾝしているﾌｫｰﾑ全てに配置しているﾃｷｽﾄﾎﾞｯｸｽ、ｺﾝﾎﾞﾎﾞｯｸｽの値をｸﾘｱする。                          _/_/_/_/_/
'_/_/_/_/ また、ﾃｷｽﾄﾎﾞｯｸｽ、ｺﾝﾎﾞﾎﾞｯｸｽの既定値ﾌﾟﾛﾊﾟﾃｨに値が入力されている場合は、既定値をｾｯﾄする。           _/_/_/_/_/
'_/_/_/_/ 単一ﾌｫｰﾑではﾃｽﾄ済                                                                             _/_/_/_/_/
'_/_/_/_/ 複数ﾌｫｰﾑでは未確認の為、如何なる現象が起きても保証しません。                                     _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Sub CmTxtBoxClear()

  Dim ojFrm   As Form             '_/_/_/ ﾌｫｰﾑ変数
  Dim ojCtrl  As Control          '_/_/_/ ｺﾝﾄﾛｰﾙ変数

'_/_/_/ ｱｸﾃｨﾌﾞなﾌｫｰﾑの数だけﾙｰﾌﾟ
  For Each ojFrm In Forms
'_/_/_/ ｱｸﾃｨﾌﾞなﾌｫｰﾑのｺﾝﾄﾛｰﾙ数だけﾙｰﾌﾟ
    For Each ojCtrl In ojFrm.Controls
'_/_/_/ ｺﾝﾄﾛｰﾙﾀｲﾌﾟがﾃｷｽﾄﾎﾞｯｸｽ、ｺﾝﾎﾞﾎﾞｯｸｽのみ値をｸﾘｱ
      Select Case ojCtrl.ControlType
        Case acTextBox, acComboBox
'_/_/_/ 既定値に値が入力されているかどうか判定し、値が指定されている場合は既定値を代入
'_/_/_/ それ以外は空文字を代入する
          If ojCtrl.DefaultValue = "" Or IsNull(ojCtrl.DefaultValue) Then
            ojCtrl = ""
          ElseIf CmLeftB(ojCtrl.DefaultValue, 1) = "=" Then
            ojCtrl.DefaultValue = ojCtrl.DefaultValue
            'ojCtrl.Requery
          Else
            ojCtrl = ojCtrl.DefaultValue
          End If
        Case acOptionGroup
          If ojCtrl.DefaultValue = "" Or IsNull(ojCtrl.DefaultValue) Then
            ojCtrl = ""
          Else
            ojCtrl.DefaultValue = ojCtrl.DefaultValue
          End If
      End Select
    Next
  Next

End Sub
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/ lstDataFile  :最適化したいMDBのﾌｧｲﾙ名（ﾌﾙﾊﾟｽで記述)                                            _/_/_/_/_/
'_/_/_/_/ lstBackUpFile:最適化後のﾌｧｲﾙ名(ﾌﾙﾊﾟｽで記述)                                                   _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Sub CompactDB(lstDataFile As String, lstBackUpFile As String)

'/****** 圧縮先ﾌｧｲﾙの存在ﾁｪｯｸ
'/****** 若し、ﾌｧｲﾙが存在すればﾌｧｲﾙを削除
'/****** (上書きできへんのかねー、しっかし、上書きしてくれたらこんなｺｰﾄﾞ書かんでええのになぁ(;_;))
  If Dir(lstBackUpFile) <> "" And Not IsNull(Dir(lstBackUpFile)) Then
    Kill lstBackUpFile
  End If

'/****** ﾃﾞｰﾀﾍﾞｰｽの圧縮
  DBEngine.CompactDatabase lstDataFile, lstBackUpFile
  'DBEngine                                           'DELETE 20030707 N.MIURA
End Sub


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/_/          ﾌｫｰﾑのｵｰﾌﾟﾝ/ｸﾛｰｽﾞ                                            _/_/_/_/_/
'_/_/_/_/_/_/ stOpenClose:ｵｰﾌﾟﾝするﾌｫｰﾑの名前                                       _/_/_/_/_/
'_/_/_/_/_/_/ stClFname  :ｸﾛｰｽﾞするﾌｫｰﾑの名前                                       _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Sub CmFormOpenClose(stOpFname As String, stClFname As String)

'_/_/_/ 引数stOpFnameで指定したﾌｫｰﾑを開く
  doCmd.OpenForm stOpFname
'_/_/_/ 引数stClFnameで指定したﾌｫｰﾑを閉じる
  doCmd.Close acForm, stClFname

End Sub

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/_/_/  メニューバーの非表示                                                     _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Sub CmSetMenuBar()

'_/_/_/ Iniﾌｧｲﾙから取得したﾃｽﾄﾓｰﾄﾞｾｸｼｮﾝのﾒﾆｭｰﾊﾞｰのｷｰ値を判定
  If UCase(gstMenubar) = "ON" Then
    Application.CommandBars.ActiveMenuBar.Enabled = True
  ElseIf UCase(gstMenubar) = "OFF" Then
    Application.CommandBars.ActiveMenuBar.Enabled = False
  End If

End Sub

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/       ﾌｫｰﾑのｻｲｽﾞの変更                                                        _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Sub CmMoveSize()

'_/_/_/ ﾒﾆｭｰﾊﾞｰ非表示関数の呼出
  Call CmSetMenuBar

'_/_/_/ ﾌｫｰﾑｻｲｽﾞの変更
  doCmd.MoveSize 0, 0, 9600, 4800

End Sub

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/       ﾌｫｰﾑのｻｲｽﾞの変更                                                        _/_/_/_/_/
'_/_/_/_/_/       ﾃﾞｨｽﾌﾟﾚｲのｻｲｽﾞが 800*600 で小さいﾌｫﾝﾄを基準(ﾃﾞｽｸﾄｯﾌﾟを対象)              _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Sub CmWindowSize()

'_/_/_/ ﾒﾆｭｰﾊﾞｰ非表示関数の呼出
  Call CmSetMenuBar

  'DoCmd.Maximize
'_/_/_/ ﾌｫｰﾑｻｲｽﾞの変更
  doCmd.MoveSize 0, 0, 11940, 7940

End Sub



'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/_/_/   ﾃﾞｰﾀﾍﾞｰｽの修復                                                         _/_/_/_/_/
'_/_/_/_/_/_/_/  stPath:修復MDBのﾊﾟｽ                                                     _/_/_/_/_/
'_/_/_/_/_/_/_/  stMDB :修復MDBのﾌｧｲﾙ名                                                  _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Sub CmRepairDatabase(stMDBPath As String, stMDBName As String)

  Dim errLoop As Error
  Dim stPath As String
  Dim linRtnValue As Integer

'/****** ﾊﾟｽのﾁｪｯｸ.引数の文字列の最後の文字が￥であるかどうか判定
  If Right$(stPath, 1) <> "\" Then
    stPath = stMDBPath + "\"
  Else
    stPath = stMDBPath
  End If

'_/_/_/ 引数で指定したﾊﾟｽの存在ﾁｪｯｸ
  If Dir(stPath) = "" Or IsNull(Dir(stPath)) Then
    doCmd.Quit acQuitSaveAll
    Exit Sub
  End If

'_/_/_/ ｽﾃｰﾀｽﾊﾞｰの文字列変更
  linRtnValue = SysCmd(acSysCmdSetStatus, "ﾃﾞｰﾀﾍﾞｰｽの修復中...")

  On Error GoTo Err_Repair
'/****** ﾃﾞｰﾀﾍﾞｰｽの修復
  DBEngine.RepairDatabase stPath + stMDBName

'/****** ｽﾃｰﾀｽﾊﾞｰﾃｷｽﾄの消去
  linRtnValue = SysCmd(acSysCmdSetStatus, " ")

'/****** ｽﾃｰﾀｽﾊﾞｰｸﾘｱ
  linRtnValue = SysCmd(acSysCmdClearStatus)
  
  Exit Sub

Err_Repair:

  For Each errLoop In DBEngine.Errors
    MsgBox "Repair unsuccessful!" & vbCr & "Error number: " & errLoop.Number & vbCr & errLoop.Description
  Next errLoop

  doCmd.Quit acQuitSaveAll

End Sub

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/       ﾌｫｰﾑのｻｲｽﾞの変更                                                        _/_/_/_/_/
'_/_/_/_/_/       ﾌｫｰﾑを最大化                                                            _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Sub CmWindowMax()

'_/_/_/ ﾒｰﾆｭｰﾊﾞｰ非表示関数の呼出
  Call CmSetMenuBar

'_/_/_/ ﾌｫｰﾑの最大化
  doCmd.Maximize

End Sub

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/       ｼｽﾃﾑの終了                                                              _/_/_/_/_/
'_/_/_/_/_/       stSysName：ｼｽﾃﾑの名前を指定.ﾒｯｾｰｼﾞﾎﾞｯｸｽに表示される                      _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Sub CmSysEnd(stSysName As String)

  If MsgBox(stSysName + "を終了します。" + vbCrLf + "よろしいですか？", vbQuestion + vbYesNo, stSysName) <> vbYes Then
    Exit Sub
  End If

'_/_/_/ ｼｽﾃﾑを終了
  doCmd.Quit acQuitSaveAll

End Sub


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/_/_/_/_/       ﾊﾟｽｽﾙｰｸｴﾘの実行                                                         _/_/_/_/_/
'_/_/_/_/_/       stSql：実行するSql文                                                    _/_/_/_/_/
'_/_/_/_/_/       成功したらTrueを失敗したらFalseを返す                                    _/_/_/_/_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Function ExecutePassThroughQry(stsql As String) As Boolean

  On Error GoTo Err_ExecutePassThroughQry

  Dim ojQRY       As QueryDef
  Dim ojDb        As Database
  Dim stConnect   As String
  Dim errLoop     As Error
  Dim ojWksp      As Workspace
  Dim ojCon       As Connection

  Set ojWksp = CreateWorkspace("", "sa", "", dbUseODBC)
  'Set ojdb = CurrentDb
  'Set ojQry = ojdb.CreateQueryDef("")

'/****** SQLｻｰﾊﾞｰへの接続文字列の作成
  stConnect = "ODBC;Database=" & GetIniString("odbc", "database", "ueki.ini") & ";"
  stConnect = stConnect + "UID=" & GetIniString("odbc", "UID", "ueki.ini") & ";"
  stConnect = stConnect + "PWD=" & GetIniString("odbc", "PassWord", "ueki.ini") & ";"
  stConnect = stConnect + "DSN=" & GetIniString("odbc", "dsn", "ueki.ini")
  
  Set ojCon = ojWksp.OpenConnection("", dbDriverNoPrompt, False, stConnect)
  'ojQry.Connect = stConnect
  'Set ojQry = ojCon.CreateQueryDef("")

  'ojQry.SQL = stSql

  'ojQry.Execute dbRunAsync + dbFailOnError
  ojCon.Execute stsql
  If ojCon.RecordsAffected = 0 Then
    GoTo Err_ExecutePassThroughQry
    Exit Function
  End If

  'ojdb.QueryDefs.Delete ojQry.Name
  ojCon.Close
  ojWksp.Close
  Set ojCon = Nothing
  Set ojWksp = Nothing
'  ojdb.Close

  ExecutePassThroughQry = True

  Exit Function

Err_ExecutePassThroughQry:

  ojCon.Close
  ojWksp.Close
  Set ojCon = Nothing
  Set ojWksp = Nothing
  
  If DBEngine.Errors.Count > 0 Then
    For Each errLoop In DBEngine.Errors
      MsgBox "Error Number:" + CStr(errLoop.Number) + vbCrLf + errLoop.Description
    Next
  End If

  ExecutePassThroughQry = False

  Exit Function

End Function
Public Sub CmWindowSizeR()

'_/_/_/ ﾒﾆｭｰﾊﾞｰ非表示関数の呼出
  Call CmSetMenuBar

  'DoCmd.Maximize
'_/_/_/ ﾌｫｰﾑｻｲｽﾞの変更
  doCmd.MoveSize 0, 0, 11940, 7550
End Sub

Public Function CmIsDate(ByVal vYear As Variant, ByVal vMonth As Variant, ByVal vDATE As Variant)
  '日付のチェック
    
  '日付の項目が、「年」、「月」、「日」の3項目にわかれている時に使用する。
    
  'スマイルには、DLLで判定させるらしいが
  'Accessの DateSerial 関数 は、バグが発見されているので
  '論理チェックをする
  
  '返り値 ０ 正常
  '       １ 判定不能   ←判定項目が不足の時
  '       ９ 論理エラー
  
  '前提は、1990年から2089年までを判断の対象とする
  'ここを変える時は、ZYearF ,ZYearTの値を変えてください
  
  Dim ZYearF As Integer
  Dim ZYearT As Integer
  Dim bZUru   As Boolean   '閏年の時は、True
    
  ZYearF = 1990
  ZYearT = 2089
  
On Error GoTo err_rtn
    
  CmIsDate = 1
  If Nz(CmNumeric(vYear), "") = "" Or Nz(CmNumeric(vMonth), "") = "" Or Nz(CmNumeric(vDATE), "") = "" Then
    Exit Function
  End If
  
  vYear = CInt(vYear)
  vMonth = CInt(vMonth)
  vDATE = CInt(vDATE)
  
  CmIsDate = 9
  If vYear < 0 Or vYear < 0 Or vYear < 0 Then
    Exit Function
  End If
  
  If vYear < ZYearF Or vYear > ZYearT Then
    Exit Function
  End If
      
  If vMonth < 1 Or vMonth > 12 Then
    Exit Function
  End If
  
  Select Case vMonth
    Case 1, 3, 5, 7, 8, 10, 12
      If vDATE < 1 Or vDATE > 31 Then
        Exit Function
      End If
    Case 4, 6, 9, 11
      If vDATE < 1 Or vDATE > 30 Then
        Exit Function
      End If
    Case 2

      If (vYear Mod 4) = 0 Then
        bZUru = True
        If (vYear Mod 100) = 0 And (vYear Mod 400) <> 0 Then
          bZUru = False
        End If
      End If
      
      If bZUru = False Then
        If vDATE < 1 Or vDATE > 28 Then
          Exit Function
        End If
      Else
        If vDATE < 1 Or vDATE > 29 Then
          Exit Function
        End If
      End If
        
End Select
      
  CmIsDate = 0
err_rtn:
  Exit Function
End Function

