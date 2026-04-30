Attribute VB_Name = "SVS520"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 予約受付検索
'
'        PROGRAM_NAME    : 予約受付検索制御
'        PROGRAM_ID      : XXXXXX
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2005/07/26
'        CERATER         : T.SUZUKI
'        Ver             : 0.0
'
'        UPDATE          :
'        UPDATER         :
'        Ver             :
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private pstrParentFormName   As String  ' フォーム
Private pstrParentUKNO       As String  ' 受付番号オブジェクト

Public pstrUKNO              As String  ' 受付番号

'==============================================================================*
'
'        MODULE_NAME      :検索ダイアログ表示
'        MODULE_ID        :psubSearchUKdata
'        Parameter        :strFormName       = フォーム名
'                          strParentUKNOName = コントロール名
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub psubSearchUKdata(strFormName As String, strParentUKNOName As String)

    pstrParentFormName = strFormName    ' フォーム名
    pstrParentUKNO = strParentUKNOName  ' コントロール名

    ' 受付番号初期化
    pstrUKNO = ""

    ' 検索画面を表示する
    doCmd.OpenForm "FVS520_予約受付検索", acNormal, , , , acDialog
End Sub

'==============================================================================*
'
'        MODULE_NAME      :検索結果の内容を予約受付入力画面に反映させる
'        MODULE_ID        :psubSetDataFVS500
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub psubSetDataFVS500()

    ' 遷移元画面の受付番号コンボボックスに値をセット
    With Forms(pstrParentFormName).Controls(pstrParentUKNO)
        .VALUE = pstrUKNO
    End With
End Sub
'****************************  ended or program ********************************
