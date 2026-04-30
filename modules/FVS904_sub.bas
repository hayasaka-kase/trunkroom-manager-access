Attribute VB_Name = "FVS904_sub"
Option Compare Database

'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　一括修繕一覧用共通関数
'   プログラムＩＤ　：　一括修繕一覧入力用関数
'   作　成　日　　　：  2017/11/13
'   作　成　者　　　：  K.SATO
'**********************************************

' 構造体(起動パラメータ保持)
Public Type Type_SHUZEN_LIST
    CHK         As Boolean            ' 選択
    UNIQEC      As String             ' ユニークコード
    YARDC       As String             ' ヤードコード
    ROOMC       As String             ' 部屋番号(コンテナ番号)
    STEP        As String             ' 段
    KOKYC       As String             ' 顧客コード
    INFO        As String             ' 情報
    ENTRY910    As Boolean            ' ９１０登録
    KEY_TYPE    As String             ' 鍵タイプ
    NAME_60_VAL As String             ' 鍵自動区分
End Type
Public pSelectArgs() As Type_SHUZEN_LIST

Public Const chkon As Boolean = True
Public Const chkoff As Boolean = False

Function cntOnData()

    Dim i As Integer
    Dim cnt As Integer

    cnt = 0
    For i = 0 To UBound(FVS904_sub.pSelectArgs)
        If FVS904_sub.pSelectArgs(i).CHK = chkon Then
            cnt = cnt + 1
        End If
    Next i
    
    cntOnData = cnt

End Function

Function chkOnDataSame()

    Dim i As Integer
    Dim wk As String
    Dim chkBoo As Boolean

    chkBoo = True
    
    For i = 0 To UBound(FVS904_sub.pSelectArgs)
        If FVS904_sub.pSelectArgs(i).CHK = chkon Then
            If wk = "" Then
                wk = FVS904_sub.pSelectArgs(i).INFO
            Else
                If wk <> FVS904_sub.pSelectArgs(i).INFO Then
                    chkBoo = False
                    Exit For
                End If
            End If
        End If
    Next i
    
    chkOnDataSame = chkBoo

End Function

Function chkOnDataFoundNull()

    Dim i As Integer
    Dim wk As String
    Dim chkBoo As Boolean

    chkBoo = True
    
    For i = 0 To UBound(FVS904_sub.pSelectArgs)
        If FVS904_sub.pSelectArgs(i).CHK = chkon Then
            If FVS904_sub.pSelectArgs(i).INFO = "" Then
                chkBoo = False
            End If
        End If
    Next i
    
    chkOnDataFoundNull = chkBoo

End Function

Function getOnDataOne()

    Dim i As Integer
    Dim wk As String

    cnt = 0
    For i = 0 To UBound(FVS904_sub.pSelectArgs)
        If FVS904_sub.pSelectArgs(i).CHK = chkon Then
            wk = FVS904_sub.pSelectArgs(i).INFO
            Exit For
        End If
    Next i
    
    getOnDataOne = wk

End Function

Function getOnDataFirst()

    Dim i As Integer
    Dim wk As Integer

    cnt = 0
    For i = 0 To UBound(FVS904_sub.pSelectArgs)
        If FVS904_sub.pSelectArgs(i).CHK = chkon Then
            wk = i
            Exit For
        End If
    Next i
    
    getOnDataFirst = wk

End Function

Sub chkOnOffData(YARDC As Variant, ROOMC As Variant, STEP As Variant, CHK As Object)

    Dim i As Integer

    For i = 0 To UBound(pSelectArgs)
        If pSelectArgs(i).YARDC = YARDC And _
           pSelectArgs(i).ROOMC = ROOMC And _
           pSelectArgs(i).STEP = STEP _
        Then
            If CHK Then
                pSelectArgs(i).CHK = chkon
            Else
                pSelectArgs(i).CHK = chkoff
            End If
            Exit For
        End If
    Next i

End Sub

Function getOneEntryData()

    Dim i As Integer
    Dim wk As Integer

    cnt = 0
    For i = 0 To UBound(FVS904_sub.pSelectArgs)
        If FVS904_sub.pSelectArgs(i).ENTRY910 = chkoff And FVS904_sub.pSelectArgs(i).CHK = chkon Then
            wk = i
            FVS904_sub.pSelectArgs(i).ENTRY910 = chkon
            Exit For
        End If
    Next i
    
    getOneEntryData = wk
        
End Function

Function getOnDataCSVRooms()

    Dim i As Integer
    Dim wk As String

    cnt = 0
    For i = 0 To UBound(FVS904_sub.pSelectArgs)
        If FVS904_sub.pSelectArgs(i).CHK = chkon Then
            If wk = "" Then
            Else
                wk = wk & " , "
            End If
            
            wk = wk & FVS904_sub.pSelectArgs(i).ROOMC
        End If
    Next i
    
    getOnDataCSVRooms = wk

End Function
'****************************  ended or program ********************************

