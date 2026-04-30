VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_RVS530"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　精算依頼リスト印刷
'   プログラムＩＤ　：　RVS770
'   作　成　日　　　：  2005/08/10
'   作　成　者　　　：  T.SUZUKI
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/11/29
'   修　正　者　　　：　T.SUZUKI
'   修　正　内　容　：  申し込み期限日の出力を追加
'   Ver             ：　0.1
'**********************************************
'   修　正　日　　　：　2007/03/01
'   修　正　者　　　：　tajima
'   修　正　内　容　：  文言修正
'   Ver             ：　0.3
'**********************************************
'   修　正　日　　　：　2007/04/06
'   修　正　者　　　：　tajima
'   修　正　内　容　：  営業開始日当日の場合、注意文言出力
'   Ver             ：　0.4
'**********************************************
Option Compare Database
Option Explicit

Private Const pstrMessage1  As String = "  ご登録されております条件のコンテナが"
Private Const pstrMessage2  As String = "件該当がありましたので、お知らせいたします。"
Private Const pstrMessage3  As String = "詳細につきましては、弊社までお問い合わせください。"

' 2005/11/29 ADD T.SUZUKI Start
Private Const pstrMessage4  As String = "今回の申し込み期限日は"
Private Const pstrMessage5  As String = "とさせていただきます。"
' 2005/11/29 ADD T.SUZUKI End

Private pintRowCount        As Integer

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Report_Open
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Report_Open(Cancel As Integer)

'    Dim varForm_OpenArgs  As Variant
'    Dim intDataCount      As Integer
'
'    ' 画面からのパラメータを取得
'    varForm_OpenArgs = Nz(OpenArgs, "")
'    If varForm_OpenArgs <> "" Then
'        intDataCount = CInt(varForm_OpenArgs)
'    End If
'
'    Me.lbl_Message1.Caption = pstrMessage1 & intDataCount & pstrMessage2
'    Me.lbl_Message2.Caption = pstrMessage3
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :グループヘッダー1_Format
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub グループヘッダー1_Format(Cancel As Integer, FormatCount As Integer)
    ' グループヘッダーフォーマット時、該当件数をセットする。
    Me.lbl_Message1.Caption = pstrMessage1 & Me.txt_DataCount.VALUE & pstrMessage2
    Me.lbl_Message2.Caption = pstrMessage3
    Me.lbl_Message3.Caption = pstrMessage4 & Nz(Me.YOUKT_TKDATE.VALUE, "") & pstrMessage5  ' 2005/11/29 ADD T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :グループフッダー1_Format
'        CREATE_DATE      :2007/04/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub グループフッター1_Format(Cancel As Integer, FormatCount As Integer)
  
    If Nz(Me.txt_営業開始日.VALUE, "") = Format$(DATE, "yyyy/mm/dd") Then
        Me.lbl_営業開始日注意.Caption = "ご利用開始日は " & Me.txt_営業開始日.VALUE & " 以降とさせて頂きます。" & vbCrLf & _
                                      "但し､ご利用開始日については変更になる場合も御座いますのでご了承下さい "
    Else
        Me.lbl_営業開始日注意.Caption = ""
    End If

End Sub


