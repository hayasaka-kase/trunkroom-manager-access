Attribute VB_Name = "MSZZ020"
'****************************  strat or program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : 起動時マクロ
'       PROGRAM_ID      : MSZZ020
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2005/09/13
'       CERATER         : K.ISHIZAKA
'       Ver             : 0.0
'
'       UPDATE          : 2009/12/07
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.1
'                       : 統合検索MDBのときはFVS011を起動する
'
'       UPDATE          : 2011/02/22
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.2
'                       : 統合検索MDBのときはFTG011を起動する
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'
'       MODULE_NAME     : 起動時マクロ
'       MODULE_ID       : MSZZ020_M00
'       CREATE_DATE     : 2005/09/13
'       PARAM           : strOpenForm       KAGTOS のとき"MFMESTR"
'                                           KOMS   のとき"FKS000"
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ020_M00(ByVal strOpenForm As String) As Boolean
    If InStr(UCase(Application.CurrentProject.NAME), "TOGO2002") = 1 Then       'INSERT START 2009/12/07 K.ISHIZAKA
'        doCmd.OpenForm "FVS011"                                                'DELETE 2011/02/22 K.ISHIZAKA
        doCmd.OpenForm "FTG011"                                                 'INSERT 2011/02/22 K.ISHIZAKA
    Else                                                                        'INSERT END   2009/12/07 K.ISHIZAKA
        If Command = "" Then
            doCmd.OpenForm strOpenForm
        End If
    End If                                                                      'INSERT 2009/12/07 K.ISHIZAKA
    MSZZ020_M00 = True
End Function

'****************************  ended or program ********************************

