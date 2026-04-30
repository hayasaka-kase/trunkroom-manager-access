VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_RKS240"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : レンタル管理システム
'        SUB_SYSTEM_NAME : 売上管理サブシステム
'
'        PROGRAM_NAME    : 請求書
'        PROGRAM_ID      : RKS240
'        PROGRAM_KBN     : REPORT
'
'        CREATE          :
'        CERATER         :
'        Ver             :
'
'        UPDATE          : 2008/06/16
'        UPDATER         : N.MIURA
'        VER             : 0.6
'
'        UPDATE          : 2017/12/01
'        UPDATER         : N.IMAI
'        VER             : 1.6
'                          顧客コード順に出力
'
'        UPDATE          : 2018/01/24
'        UPDATER         : N.IMAI
'        VER             : 1.7
'                          仮想口座が存在する場合は仮想口座のみ出力する
'
'        UPDATE          : 2018/11/16
'        UPDATER         : N.IMAI
'        VER             : 1.9
'                          収納代行会社の文言追加
'
'        UPDATE          : 2018/12/15
'        UPDATER         : N.IMAI
'        VER             : 2.1
'                          送付先名称を修正
'                          上段：書類送付先担当名なければ顧客名
'                          下段：書類送付先担当名なしで法人の場合、代表者名（入力されていたら）を出力
'
'        UPDATE          : 2019/02/04
'        UPDATER         : N.IMAI
'        VER             : 2.2
'                          送付先名称を修正
'                          上段：個人の場合、書類送付先担当名なければ顧客名
'                          下段：書類送付先担当名なしで法人の場合、代表者名（入力されていたら）を出力
'                          敬称は全て「様」とする
'
'        UPDATE          : 2019/10/31
'        UPDATER         : K.ISHIZAKA
'        VER             : 2.7
'                          請求書レイアウト変更
'
'        UPDATE          : 2020/08/19
'        UPDATER         : N.IMAI
'        VER             : 2.8
'                          収納代行会社の文言を削除（非表示）
'                          消費税→内税に変更
'
'        UPDATE          : 2022/08/10
'        UPDATER         : N.IMAI
'        VER             : 3.3
'                          請求書の口座名対応
'
'        UPDATE          : 2022/08/15
'        UPDATER         : N.IMAI
'        VER             : 3.4
'                          請求書の口座名対応
'
'==============================================================================*
Option Compare Database
Option Explicit

Private Sub ページフッター_Print(Cancel As Integer, PrintCount As Integer)
    
    If Nz(Me.txtKOUZB) = "" Then
'       Me.BOX_NAME.Visible = False
       Me.LABEL_KINYN_NAME.Visible = False
       Me.LABEL_KOUZB_NAME.Visible = False
       Me.LABEL_KOUZN_NAME.Visible = False
       Me.LABEL_USER_KANA.Visible = False
       Me.txtKINYN.Visible = False
       Me.txtYOKIN.Visible = False
       Me.txtKOUZN.Visible = False
       Me.txt代行文言.Visible = False                                           'INSERT 2022/08/15 N.IMAI
       Me.txtUserKana.Visible = False
    Else
'       Me.BOX_NAME.Visible = True
       Me.LABEL_KINYN_NAME.Visible = True
       Me.LABEL_KOUZB_NAME.Visible = True
       Me.LABEL_KOUZN_NAME.Visible = True
       Me.LABEL_USER_KANA.Visible = True
       Me.txtKINYN.Visible = True
       Me.txtYOKIN.Visible = True
       Me.txtKOUZN.Visible = True
       Me.txt代行文言.Visible = True                                            'INSERT 2022/08/15 N.IMAI
       Me.txtUserKana.Visible = True
       
    End If
    'INSERT 2018/01/24 N.IMAI Start
    Me.LABEL_KINYN_NAME1.Visible = Not Me.LABEL_KINYN_NAME.Visible
    Me.LABEL_KOUZB_NAME1.Visible = Not Me.LABEL_KOUZB_NAME.Visible
    Me.LABEL_KOUZN_NAME1.Visible = Not Me.LABEL_KOUZN_NAME.Visible
    Me.LABEL_USER_KANA1.Visible = Not Me.LABEL_USER_KANA.Visible
    Me.txtKINYN1.Visible = Not Me.txtKINYN.Visible
    Me.txtYOKIN1.Visible = Not Me.txtYOKIN.Visible
    Me.txtKOUZN1.Visible = Not Me.txtKOUZN.Visible
    Me.txt代行文言2.Visible = Not Me.txt代行文言.Visible                        'INSERT 2022/08/15 N.IMAI
    Me.txtUSERKANA1.Visible = Not Me.txtUserKana.Visible
    'INSERT 2018/01/24 N.IMAI End

End Sub

Private Sub 詳細_Print(Cancel As Integer, PrintCount As Integer)
    On Error Resume Next

    If Nz(Form_FKS240.txt_日付FROM) = "" Then   'INSERT 2011/01/12 M.RYU---<S>
        Me.txt_Frkg = ""
'        Me.lb_Frkg.Visible = False             'DELETE 2012/01/17 M.RYU
        Me.txt_Frkg.Visible = False
    Else
'        Me.lb_Frkg.Visible = True              'DELETE 2012/01/17 M.RYU
        Me.txt_Frkg.Visible = True
    End If                                      'INSERT 2011/01/12 M.RYU---<E>
    
    imgStamp.Picture = Nz(txtStampPath, "")

End Sub
