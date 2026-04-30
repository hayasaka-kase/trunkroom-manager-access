Attribute VB_Name = "CmReprMod"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME :
'
'        PROGRAM_NAME    : 修繕トラン用の共通関数群
'        PROGRAM_ID      : CmRperMod
'        PROGRAM_KBN     :
'
'        CREATE          : 2006/05/13
'        CERATER         : tajima
'        Ver             : 0.0
'
'        UPDATE          : 2006/09/02
'        UPDATER         : tajima
'        Ver             : 0.1
'                        : 既存機能の修繕対応
'                        : ①直近修繕取得関数用意
'
'        UPDATE          : 2007/01/26
'        UPDATER         : tajima
'        Ver             : 0.2
'                        : ①修繕、顧客問合わせの内容項目追加
'
'        UPDATE          : 2010/06/10
'        UPDATER         : SHIBAZAKI
'        Ver             : 0.3
'                        : 直近の顧客問合せを検索する際、表示区分を見る。
'
'        UPDATE          : 2013/06/17
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.4
'                        : 顧客問合わせの内容項目追加
'
'        UPDATE          : 2021/08/03
'        UPDATER         : N.IMAI
'        Ver             : 0.5
'                        : 07:メンテ依頼を追加
'
'        UPDATE          : 2022/02/27
'        UPDATER         : N.IMAI
'        Ver             : 0.6
'                        : 18:未設置、19:仮置きを追加
'
'        UPDATE          : 2025/09/12
'        UPDATER         : N.IMAI
'        Ver             : 0.7
'                        : 修理中の判定を変更
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "CmRperMod"

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 修繕関係定数
' 区分コード 252 修繕対象
Public Const P_対象_ヤード   As String = "01"
Public Const P_対象_ボックス As String = "02"
Public Const P_対象_部屋     As String = "03"
Public Const P_対象_顧客     As String = "04"
' 区分コード 250 修繕入力対象
Public Const P_修繕_工事       As String = "01"  ' 工事
Public Const P_修繕_修理       As String = "02"  ' 修理
Public Const P_修繕_コメント   As String = "03"  ' コメント
Public Const P_修繕_クレーム   As String = "04"  ' クレーム
Public Const P_修繕_撤去予定   As String = "05"  ' 撤去予定
Public Const P_修繕_メンテ依頼 As String = "07"  ' メンテ依頼                   'INSERT 2021/08/03 N.IMAI


Public Const P_修繕_組替え         As String = "13"  ' 組替え
Public Const P_修繕_鍵返却         As String = "14"  ' 鍵返却
Public Const P_修繕_社内理由       As String = "15"  ' 社内理由
Public Const P_修繕_オーナー理由   As String = "16"  ' オーナー理由
Public Const P_修繕_解約予定       As String = "17"  ' 解約予定
Public Const P_修繕_未設置         As String = "18"  ' 未設置
Public Const P_修繕_仮置き         As String = "19"  ' 仮置き                   'INSERT 2022/02/27 N.IMAI


' 修繕を呼ぶファンクションキー
Public Const P_呼出_ヤード    As Integer = vbKeyF8
Public Const P_呼出_顧客      As Integer = vbKeyF9
Public Const P_呼出_ボックス  As Integer = vbKeyF11
Public Const P_呼出_部屋      As Integer = vbKeyF12


' 修繕トラン構造体 ※Null許可項目はVariant
Public Type Type_REPR_TRAN
    UNIQEC  As String
    BUMOC   As String
    YARDC   As String
    ROOMC   As Variant
    INPTI   As String
    TANTC   As String
    GENTD   As String
    LIMTD   As Variant
    COMPD   As Variant
    TYPEC   As String
    REPRC   As Variant
    CNT1N   As Variant
    CNT1N1   As Variant '2007/01/26 add tajima
    CNT1N2   As Variant '2007/01/26 add tajima
    CNT2N   As Variant
    CNT2N1   As Variant '2007/01/26 add tajima
    CNT2N2   As Variant '2007/01/26 add tajima
    KINGA   As Variant
    WORKC   As Variant
    GYOUN   As Variant
    ACPTB   As Variant
    INSED   As String
    INSEJ   As String
    INSPB   As String
    INSUB   As String
    UPDAD   As Variant
    UPDAJ   As Variant
    UPDPB   As Variant
    UPDUB   As Variant
End Type

' 顧客問合わせトラン構造体 ※Null許可項目はVariant
Public Type Type_KOTO_TRAN
    UNIQEC  As String
    BUMOC   As String
    KOKYC   As String
    TANTC   As String
    GENTD   As String
    GENTJ   As String   'INSERT 2013/06/17 K.ISHIZAKA
    TYPEC   As String
    CONTN   As Variant
    CNT1N   As Variant '2007/01/26 add tajima
    CNT2N   As Variant '2007/01/26 add tajima
    INSED   As String
    INSEJ   As String
    INSPB   As String
    INSUB   As String
    UPDAD   As Variant
    UPDAJ   As Variant
    UPDPB   As Variant
    UPDUB   As Variant
End Type
'==============================================================================*
'
'        MODULE_NAME      :修繕トラン構造体初期化
'        MODULE_ID        :InitReprTranType
'        Parameter        :第1引数(ByRef構造体) = 修繕トラン内容
'        戻り値           : Nothing
'        CREATE_DATE      :2006/05/19
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub InitReprTranType(ByRef aREPR_TRAN As Type_REPR_TRAN)

    aREPR_TRAN.UNIQEC = ""
    aREPR_TRAN.BUMOC = ""
    aREPR_TRAN.YARDC = ""
    aREPR_TRAN.ROOMC = Null
    aREPR_TRAN.INPTI = ""
    aREPR_TRAN.TANTC = ""
    aREPR_TRAN.GENTD = ""
    aREPR_TRAN.LIMTD = Null
    aREPR_TRAN.COMPD = Null
    aREPR_TRAN.TYPEC = ""
    aREPR_TRAN.REPRC = Null
    aREPR_TRAN.CNT1N = Null
    aREPR_TRAN.CNT1N1 = Null '2007/01/26 add tajima
    aREPR_TRAN.CNT1N2 = Null '2007/01/26 add tajima
    aREPR_TRAN.CNT2N = Null
    aREPR_TRAN.CNT2N1 = Null '2007/01/26 add tajima
    aREPR_TRAN.CNT2N2 = Null '2007/01/26 add tajima
    aREPR_TRAN.KINGA = Null
    aREPR_TRAN.WORKC = Null
    aREPR_TRAN.GYOUN = Null
    aREPR_TRAN.ACPTB = Null
    aREPR_TRAN.INSED = ""
    aREPR_TRAN.INSEJ = ""
    aREPR_TRAN.INSPB = ""
    aREPR_TRAN.INSUB = ""
    aREPR_TRAN.UPDAD = Null
    aREPR_TRAN.UPDAJ = Null
    aREPR_TRAN.UPDPB = Null
    aREPR_TRAN.UPDUB = Null

End Sub
'==============================================================================*
'
'        MODULE_NAME      :顧客問合わせトラン構造体初期化
'        MODULE_ID        :InitKotoTranType
'        Parameter        :第1引数(ByRef構造体) = 修繕トラン内容
'        戻り値           : Nothing
'        CREATE_DATE      :2006/05/19
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub InitKotoTranType(ByRef aKOTO_TRAN As Type_KOTO_TRAN)

    aKOTO_TRAN.UNIQEC = ""
    aKOTO_TRAN.BUMOC = ""
    aKOTO_TRAN.KOKYC = ""
    aKOTO_TRAN.TANTC = ""
    aKOTO_TRAN.GENTD = ""
    aKOTO_TRAN.GENTJ = ""   'INSERT 2013/06/17 K.ISHIZAKA
    aKOTO_TRAN.TYPEC = ""
    aKOTO_TRAN.CONTN = ""
    aKOTO_TRAN.CNT1N = Null '2007/01/26 add tajima
    aKOTO_TRAN.CNT2N = Null '2007/01/26 add tajima
    aKOTO_TRAN.INSED = ""
    aKOTO_TRAN.INSEJ = ""
    aKOTO_TRAN.INSPB = ""
    aKOTO_TRAN.INSUB = ""
    aKOTO_TRAN.UPDAD = Null
    aKOTO_TRAN.UPDAJ = Null
    aKOTO_TRAN.UPDPB = Null
    aKOTO_TRAN.UPDUB = Null

End Sub
'==============================================================================*
'
'        MODULE_NAME      :修繕トラン作成
'        MODULE_ID        :1
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(ByRef構造体) = 修繕トラン内容
'        戻り値           :String...発行されたUniqID
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :2006/09/03 UniqIDをこっちで取得
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function InsertReprTran(dbSQLServer As DAO.Database, _
                               ByRef aREPR_TRAN As Type_REPR_TRAN _
                              ) As String
                              
    Dim strSQL     As String
    Dim strNewID   As Variant
    Dim rsObject   As Recordset 'レコードセット
     
    InsertReprTran = ""
    
    On Error GoTo Exception
  
    ' UniqIDの発行 2006/09/06
    strSQL = "SELECT NEWID()"
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenForwardOnly, dbSQLPassThrough)
    strNewID = rsObject.Fields(0).VALUE
    rsObject.Close
  
    ' 追加準備
    Set rsObject = dbSQLServer.OpenRecordset("REPR_TRAN", dbOpenDynaset, dbAppendOnly)
    
   ' レコードを追加
    With rsObject
        .AddNew
        ' 共通的を設定
        Call SetReprTran(rsObject, aREPR_TRAN)
        ' インサートのみ設定する項目
        .Fields("REPRT_UNIQEC") = strNewID        ' 発行したUniqID設定 2006/09/06
        .Fields("REPRT_INSED") = aREPR_TRAN.INSED
        .Fields("REPRT_INSEJ") = aREPR_TRAN.INSEJ
        .Fields("REPRT_INSPB") = aREPR_TRAN.INSPB
        .Fields("REPRT_INSUB") = aREPR_TRAN.INSUB
        .UPDATE
    End With
    
   rsObject.Close
   Set rsObject = Nothing
   
   ' 発行したUniqID から余計なのを省く
   ' 例）DAOだと『{guid {E5D51878-1E9B-4C18-BCC7-2E5ABC0C3F3D}}』で取得されるので
   '     余計な   ~~~~~~~                                    ~~ を省く
   Dim startIdx As Integer
   startIdx = InStr(2, strNewID, "{") + 1
   strNewID = Mid(strNewID, startIdx, Len(strNewID) - (startIdx + Len(strNewID) - InStr(2, strNewID, "}")))
   aREPR_TRAN.UNIQEC = strNewID '構造体の補完
   InsertReprTran = strNewID
   
   Exit Function
                              
Exception:
  If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
  Call Err.Raise(Err.Number, "InsertReprTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function
'==============================================================================*
'
'        MODULE_NAME      :顧客問合わせトラン作成
'        MODULE_ID        :InsertKotoTran
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(ByRef構造体) = 修繕トラン内容
'        戻り値           : True...登録完了
'                         : False...エラー
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function InsertKotoTran(dbSQLServer As DAO.Database, _
                               ByRef aKOTO_TRAN As Type_KOTO_TRAN _
                              ) As Boolean
                              
    Dim strSQL     As String
    Dim rsObject   As Recordset 'レコードセット
     
    InsertKotoTran = False
    
    On Error GoTo Exception
  
    ' 追加準備
    strSQL = "SELECT * FROM KOTO_TRAN"
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)
    
   ' レコードを追加
    With rsObject
        .AddNew
        Call SetKotoTran(rsObject, aKOTO_TRAN)
        ' インサートのみ設定する項目
        .Fields("KOTOT_INSED") = aKOTO_TRAN.INSED
        .Fields("KOTOT_INSEJ") = aKOTO_TRAN.INSEJ
        .Fields("KOTOT_INSPB") = aKOTO_TRAN.INSPB
        .Fields("KOTOT_INSUB") = aKOTO_TRAN.INSUB
        .UPDATE
    End With
    
   rsObject.Close
   Set rsObject = Nothing
   
   InsertKotoTran = True
   Exit Function
                              
Exception:
  If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
  Call Err.Raise(Err.Number, "InsertKotoTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :修繕トラン更新
'        MODULE_ID        :UpdateReprTran
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(ByRef構造体) = 修繕トラン内容
'        戻り値           : True...登録完了
'                         : False...エラー
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UpdateReprTran(dbSQLServer As DAO.Database, _
                               ByRef aREPR_TRAN As Type_REPR_TRAN _
                              ) As Boolean
                              
    Dim strSQL     As String
    Dim rsObject   As Recordset
     
    UpdateReprTran = False
    
    On Error GoTo Exception
  
    ' ユニークキーで更新対象抽出
    strSQL = "SELECT * FROM REPR_TRAN WHERE REPRT_UNIQEC = '" & aREPR_TRAN.UNIQEC & "'"
    
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)
    
    ' 対象のレコードを更新
    With rsObject
        .Edit
       ' 共通的なものを設定
        Call SetReprTran(rsObject, aREPR_TRAN)
        ' アップデートのみ設定する項目
        .Fields("REPRT_UPDAD") = aREPR_TRAN.UPDAD
        .Fields("REPRT_UPDAJ") = aREPR_TRAN.UPDAJ
        .Fields("REPRT_UPDPB") = aREPR_TRAN.UPDPB
        .Fields("REPRT_UPDUB") = aREPR_TRAN.UPDUB
        .UPDATE
    End With
    
   rsObject.Close
   Set rsObject = Nothing
   
   UpdateReprTran = True
   Exit Function
                              
Exception:
  If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
  Call Err.Raise(Err.Number, "UpdatetReprTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function
'==============================================================================*
'
'        MODULE_NAME      :顧客問合トラン更新
'        MODULE_ID        :UpdateKotoTran
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(ByRef構造体) = 顧客問合トラン内容
'        戻り値           : True...登録完了
'                         : False...エラー
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UpdateKotoTran(dbSQLServer As DAO.Database, _
                               ByRef aKOTO_TRAN As Type_KOTO_TRAN _
                              ) As Boolean
                              
    Dim strSQL     As String
    Dim rsObject   As Recordset
     
    UpdateKotoTran = False
    
    On Error GoTo Exception
  
    ' ユニークキーで更新対象抽出
    strSQL = "SELECT * FROM KOTO_TRAN WHERE KOTOT_UNIQEC = '" & aKOTO_TRAN.UNIQEC & "'"
    
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)
    
    ' 対象のレコードを更新
    With rsObject
        .Edit
       ' 共通的なものを設定
        Call SetKotoTran(rsObject, aKOTO_TRAN)
        ' アップデートのみ設定する項目
        .Fields("KOTOT_UPDAD") = aKOTO_TRAN.UPDAD
        .Fields("KOTOT_UPDAJ") = aKOTO_TRAN.UPDAJ
        .Fields("KOTOT_UPDPB") = aKOTO_TRAN.UPDPB
        .Fields("KOTOT_UPDUB") = aKOTO_TRAN.UPDUB
        .UPDATE
    End With
    
   rsObject.Close
   Set rsObject = Nothing
   
   UpdateKotoTran = True
   Exit Function
                              
Exception:
  If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
  Call Err.Raise(Err.Number, "UpdateKotoTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function
'==============================================================================*
'
'        MODULE_NAME      :修繕トラン読込
'        MODULE_ID        :ReadReprTran
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(ByRef構造体) = 修繕トラン内容※キーを設定しておくこと
'        戻り値           : True...読込成功
'                         : False...対象無し
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function ReadReprTran(dbSQLServer As DAO.Database, _
                               ByRef aREPR_TRAN As Type_REPR_TRAN _
                              ) As Boolean
                              
    Dim strSQL     As String
    Dim rsObject   As Recordset
     
    ReadReprTran = False
    
    On Error GoTo Exception
  
    ' ユニークキーで対象抽出
    strSQL = "SELECT * FROM REPR_TRAN WHERE REPRT_UNIQEC = '" & aREPR_TRAN.UNIQEC & "'"
    
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    ' 対象確認
    If rsObject.EOF = False Then
    ' 対象を構造体にセッタップ！
      With rsObject
        aREPR_TRAN.BUMOC = .Fields("REPRT_BUMOC")
        aREPR_TRAN.YARDC = .Fields("REPRT_YARDC")
        aREPR_TRAN.ROOMC = .Fields("REPRT_ROOMC")
        aREPR_TRAN.INPTI = .Fields("REPRT_INPTI")
        aREPR_TRAN.TANTC = .Fields("REPRT_TANTC")
        aREPR_TRAN.GENTD = .Fields("REPRT_GENTD")
        aREPR_TRAN.LIMTD = .Fields("REPRT_LIMTD")
        aREPR_TRAN.COMPD = .Fields("REPRT_COMPD")
        aREPR_TRAN.TYPEC = .Fields("REPRT_TYPEC")
        aREPR_TRAN.REPRC = .Fields("REPRT_REPRC")
        aREPR_TRAN.CNT1N = .Fields("REPRT_CNT1N")
        aREPR_TRAN.CNT1N1 = .Fields("REPRT_CNT1N1") '2007/01/26 add tajima
        aREPR_TRAN.CNT1N2 = .Fields("REPRT_CNT1N2") '2007/01/26 add tajima
        aREPR_TRAN.CNT2N = .Fields("REPRT_CNT2N")
        aREPR_TRAN.CNT2N1 = .Fields("REPRT_CNT2N1") '2007/01/26 add tajima
        aREPR_TRAN.CNT2N2 = .Fields("REPRT_CNT2N2") '2007/01/26 add tajima
        aREPR_TRAN.KINGA = .Fields("REPRT_KINGA")
        aREPR_TRAN.WORKC = .Fields("REPRT_WORKC")
        aREPR_TRAN.GYOUN = .Fields("REPRT_GYOUN")
        aREPR_TRAN.ACPTB = .Fields("REPRT_ACPTB")
        '作成・更新情報は使用すると思えないので読込対象外
      End With
      ReadReprTran = True
    End If
    
   rsObject.Close
   Set rsObject = Nothing
   
   Exit Function
                              
Exception:
  If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
  Call Err.Raise(Err.Number, "ReadReprTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :顧客問合トラン読込
'        MODULE_ID        :ReadKotoTran
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(ByRef構造体) = 顧客問合トラン内容※キーを設定しておくこと
'        戻り値           : True...読込成功
'                         : False...対象無し
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function ReadKotoTran(dbSQLServer As DAO.Database, _
                               ByRef aKOTO_TRAN As Type_KOTO_TRAN _
                              ) As Boolean
                              
    Dim strSQL     As String
    Dim rsObject   As Recordset
     
    ReadKotoTran = False
    
    On Error GoTo Exception
  
    ' ユニークキーで対象抽出
    strSQL = "SELECT * FROM KOTO_TRAN WHERE KOTOT_UNIQEC = '" & aKOTO_TRAN.UNIQEC & "'"
    
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    
    ' 対象確認
    If rsObject.EOF = False Then
    ' 対象を構造体にセッタップ！
      With rsObject
        aKOTO_TRAN.BUMOC = .Fields("KOTOT_BUMOC")
        aKOTO_TRAN.KOKYC = .Fields("KOTOT_KOKYC")
        aKOTO_TRAN.TANTC = .Fields("KOTOT_TANTC")
        aKOTO_TRAN.GENTD = .Fields("KOTOT_GENTD")
        aKOTO_TRAN.GENTJ = Nz(.Fields("KOTOT_GENTJ"))   'INSERT 2013/06/17 K.ISHIZAKA
        aKOTO_TRAN.TYPEC = .Fields("KOTOT_TYPEC")
        aKOTO_TRAN.CONTN = .Fields("KOTOT_CONTN")
        aKOTO_TRAN.CNT1N = .Fields("KOTOT_CNT1N") '2007/01/26 add tajima
        aKOTO_TRAN.CNT2N = .Fields("KOTOT_CNT2N") '2007/01/26 add tajima
        '作成・更新情報は使用すると思えないので読込対象外
      End With
      ReadKotoTran = True
    End If
    
   rsObject.Close
   Set rsObject = Nothing
   
   Exit Function
                              
Exception:
  If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
  Call Err.Raise(Err.Number, "ReadKotoTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :修繕トランデータ設定
'        MODULE_ID        :SetKotoTran
'        Parameter        :第1引数(Recordset) = 設定するレコードセット
'                         :第2引数(ByRef構造体) = 修繕トラン内容
'        Return           : なし
'        Note             :修繕トランのテーブル定義を変えたら本関数とType_REPR_TRANを変えること
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetKotoTran(aRsObject As Recordset, ByRef aKOTO_TRAN As Type_KOTO_TRAN)

    On Error GoTo Exception
    
    With aRsObject
        .Fields("KOTOT_BUMOC") = aKOTO_TRAN.BUMOC
        .Fields("KOTOT_KOKYC") = aKOTO_TRAN.KOKYC
        .Fields("KOTOT_TANTC") = aKOTO_TRAN.TANTC
        .Fields("KOTOT_GENTD") = aKOTO_TRAN.GENTD
        .Fields("KOTOT_GENTJ") = IIf(aKOTO_TRAN.GENTJ = "", Null, aKOTO_TRAN.GENTJ) 'INSERT 2013/06/17 K.ISHIZAKA
        .Fields("KOTOT_TYPEC") = aKOTO_TRAN.TYPEC
        .Fields("KOTOT_CONTN") = aKOTO_TRAN.CONTN
        .Fields("KOTOT_CNT1N") = aKOTO_TRAN.CNT1N  '2007/01/26 add tajima
        .Fields("KOTOT_CNT2N") = aKOTO_TRAN.CNT2N  '2007/01/26 add tajima
        '作成・更新情報はここでは設定しない
    End With
    Exit Sub
    
Exception:
  If Not aRsObject Is Nothing Then aRsObject.Close: Set aRsObject = Nothing
  Call Err.Raise(Err.Number, "SetaKotoTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :修繕トランデータ設定
'        MODULE_ID        :SetReprTran
'        Parameter        :第1引数(Recordset) = 設定するレコードセット
'                         :第2引数(ByRef構造体) = 修繕トラン内容
'        Return           : なし
'        Note             :修繕トランのテーブル定義を変えたら本関数とType_REPR_TRANを変えること
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetReprTran(aRsObject As Recordset, ByRef aREPR_TRAN As Type_REPR_TRAN)

    On Error GoTo Exception
    
    With aRsObject
        .Fields("REPRT_BUMOC") = aREPR_TRAN.BUMOC
        .Fields("REPRT_YARDC") = aREPR_TRAN.YARDC
        .Fields("REPRT_ROOMC") = aREPR_TRAN.ROOMC
        .Fields("REPRT_INPTI") = aREPR_TRAN.INPTI
        .Fields("REPRT_TANTC") = aREPR_TRAN.TANTC
        .Fields("REPRT_GENTD") = aREPR_TRAN.GENTD
        .Fields("REPRT_LIMTD") = aREPR_TRAN.LIMTD
        .Fields("REPRT_COMPD") = aREPR_TRAN.COMPD
        .Fields("REPRT_TYPEC") = aREPR_TRAN.TYPEC
        .Fields("REPRT_REPRC") = aREPR_TRAN.REPRC
        .Fields("REPRT_CNT1N") = aREPR_TRAN.CNT1N
        .Fields("REPRT_CNT1N1") = aREPR_TRAN.CNT1N1 '2007/01/26 add tajima
        .Fields("REPRT_CNT1N2") = aREPR_TRAN.CNT1N2 '2007/01/26 add tajima
        .Fields("REPRT_CNT2N") = aREPR_TRAN.CNT2N
        .Fields("REPRT_CNT2N1") = aREPR_TRAN.CNT2N1 '2007/01/26 add tajima
        .Fields("REPRT_CNT2N2") = aREPR_TRAN.CNT2N2 '2007/01/26 add tajima
        .Fields("REPRT_KINGA") = aREPR_TRAN.KINGA
        .Fields("REPRT_WORKC") = aREPR_TRAN.WORKC
        .Fields("REPRT_GYOUN") = aREPR_TRAN.GYOUN
        .Fields("REPRT_ACPTB") = aREPR_TRAN.ACPTB
        '作成・更新情報はここでは設定しない
    End With
    Exit Sub
    
Exception:
  If Not aRsObject Is Nothing Then aRsObject.Close: Set aRsObject = Nothing
  Call Err.Raise(Err.Number, "SetReprTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :修繕更新権限可否
'        MODULE_ID        :IsAuthority
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = ユーザID
'        戻り値           : True...更新可
'                         : False...更新不可
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function isAuthority(dbSQLServer As DAO.Database, _
                               aユーザID As String _
                              ) As Boolean
                              
    Dim strSQL     As String
    Dim rsObject   As Recordset
     
    isAuthority = False
    
    On Error GoTo Exception
  
    ' オンライン管理テーブルから権限取得
    strSQL = "SELECT ONLET_ACCEI FROM ONLE_TABL WHERE ONLET_PROGB = 'FVS900'"
    strSQL = strSQL & "AND ONLET_USERB = '" & aユーザID & "'"
    
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    ' 対象確認
    If rsObject.EOF = False Then
      ' オンライン更新区分が"1"ならば更新権限がある
      If "1" = Nz(rsObject.Fields("ONLET_ACCEI"), "") Then
        isAuthority = True
      End If
    End If
    rsObject.Close
    Set rsObject = Nothing
   
    Exit Function

Exception:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    Call Err.Raise(Err.Number, "IsAuthority" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :CODE_TABLE用コンボボックスへの値設定
'        MODULE_ID        :SetCodeCmbItems
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(ComboBox) = 値(アイテム)を入れる対象
'                         :第3引数(String) = 値取得のWhere句
'                         :第4引数(String) = コードのフォーマットタイプ※初期値"00"
'        NOTE             :指定したコントロールに指定したWhere内容でCODE_TABLEの
'                         :内容を設定するっちゃ
'        CREATE_DATE      :2006/05/23
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetCodeCmbItems(dbSQLServer As DAO.Database, _
                            aTargetObject As ComboBox, _
                            aWhere句 As String, _
                            Optional aFormat As String = "00")
                              
    Dim strSQL     As String
    Dim rsObject   As Recordset
    Dim strItem   As String
     
    On Error GoTo Exception
  
    strSQL = "SELECT "
    strSQL = strSQL & "  CODET_CODEC"
    strSQL = strSQL & " ,CODET_NAMEN "
    strSQL = strSQL & " FROM CODE_TABL "
    strSQL = strSQL & aWhere句
    
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    
    While Not rsObject.EOF
        strItem = Format$(rsObject.Fields("CODET_CODEC"), aFormat) & ";" & rsObject.Fields("CODET_NAMEN")
        aTargetObject.AddItem (strItem)
        rsObject.MoveNext
    Wend
    
   rsObject.Close
   Set rsObject = Nothing
   
   Exit Sub
                              
Exception:
  If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
  Call Err.Raise(Err.Number, "SetCodeCmbItems" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :修繕画面の呼出
'        MODULE_ID        :OpenFVS90n
'        Parameter        :第1引数(Integer) = KeyCode
'                         :第2引数(Variant) = aヤードコード
'                         :第3引数(Variant) = a部屋番号
'                         :第4引数(Variant) = a顧客コード
'                         :第4引数(Variant) = a契約番号
'        NOTE             :押下されたキーコードによって対応する修繕一覧を呼ぶ
'                         :それ以外のキーコードは無視
'                         :ヤード・部屋・顧客の３コードは念のためゼロサプレスする
'        CREATE_DATE      :2006/05/23
'        UPDATE_DATE      :2006/09/06 呼出をダイアログモードで
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub OpenFVS90n(KeyCode As Integer, _
                      Optional aヤードコード As Variant = Null, _
                      Optional a部屋番号 As Variant = Null, _
                      Optional a顧客コード As Variant = Null, _
                      Optional a契約番号 As Variant = Null _
                      )
   
    On Error Resume Next

    Dim prgParameter As String

    Select Case KeyCode
        'ヤード修繕を呼ぶ
        Case P_呼出_ヤード
            If Nz(aヤードコード, "") = "" Then Exit Sub
            prgParameter = Format$(aヤードコード, "000000")
            doCmd.OpenForm "FVS901", , , , , , prgParameter
        
        '部屋修繕を呼ぶ
        Case P_呼出_部屋
            If Nz(aヤードコード, "") = "" Or Nz(a部屋番号, "") = "" Then Exit Sub
            prgParameter = Format$(aヤードコード, "000000") & "," & Format$(a部屋番号, "000000")
            If Nz(a契約番号, "") <> "" Then prgParameter = prgParameter & "," & a契約番号
            doCmd.OpenForm "FVS903", , , , , , prgParameter
        
        '顧客修繕を呼ぶ
        Case P_呼出_顧客
            If Nz(a顧客コード, "") = "" Then Exit Sub
            prgParameter = Format$(a顧客コード, "000000")
            If Nz(a契約番号, "") <> "" Then prgParameter = prgParameter & "," & a契約番号
            doCmd.OpenForm "FVS902", , , , , , prgParameter
                   
   End Select
   
End Sub

'==============================================================================*
'
'        MODULE_NAME      :部屋の直近修繕内容取得
'        MODULE_ID        :GetRoomLatestReprText
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String)  = a部門コード
'                         :第3引数(Variant) = aヤードコード(6桁ゼロサプレス済み前提)
'                         :第4引数(Variant) = a部屋番号(6桁ゼロサプレス済み前提)
'                         :第5引数(Variant) = a全件対象(default...false)
'        Retrun           :取得した直近修繕文言
'        NOTE             :a全件対象をTrueで取得ですると修理完了のデータも対象とす
'        CREATE_DATE      :2006/09/02
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetRoomLatestReprText(dbSQLServer As DAO.Database, _
                                      a部門コード As String, _
                                      aヤードコード As Variant, _
                                      a部屋番号 As Variant, _
                                      Optional a全件対象 As Boolean = False _
                                      ) As String
    Dim strSQL     As String
    Dim returnText As String
    Dim rsObject   As Recordset
     
    returnText = ""
    
    On Error GoTo Exception
  
    ' 指定した部屋の最新修繕情報を１件取得するSQL
    strSQL = "SELECT TOP 1 REPRT_GENTD,REPRC_NAME.CODET_NAMEN AS REPRC_NAME, REPRT_TYPEC, REPRT_CNT1N " & Chr(13)
    strSQL = strSQL & "FROM REPR_TRAN LEFT OUTER JOIN CODE_TABL AS REPRC_NAME ON " & Chr(13)
    strSQL = strSQL & " REPRT_REPRC = REPRC_NAME.CODET_CODEC AND REPRC_NAME.CODET_SIKBC = '251' " & Chr(13)
    strSQL = strSQL & "WHERE REPRT_BUMOC = '" & a部門コード & "'" & Chr(13)
    strSQL = strSQL & " AND REPRT_YARDC = '" & aヤードコード & "'" & Chr(13)
    strSQL = strSQL & " AND REPRT_ROOMC = '" & a部屋番号 & "'" & Chr(13)
    strSQL = strSQL & " AND REPRT_INPTI = '" & P_対象_部屋 & "'" & Chr(13)
      
    ' 全件対象としないのならば修理完了は除く
    If a全件対象 = False Then
'DELETE 2021/08/03 N.IMAI Start
'        strSql = strSql & " AND ( REPRT_TYPEC <> '" & P_修繕_修理 & "'" & Chr(13)
'        strSql = strSql & "       OR ( REPRT_TYPEC = '" & P_修繕_修理 & "' AND ISNULL(REPRT_COMPD,'') = '')" & Chr(13)
'        strSql = strSql & "     )" & Chr(13)
'DELETE 2021/08/03 N.IMAI End
'        'INSERT 2021/08/03 N.IMAI Start
'        strSQL = strSQL & " AND ( "
'        strSQL = strSQL & "      (REPRT_TYPEC NOT IN ('" & P_修繕_修理 & "', '" & P_修繕_メンテ依頼 & "')) " & Chr(13)
'        strSQL = strSQL & "   OR (REPRT_TYPEC     IN ('" & P_修繕_修理 & "', '" & P_修繕_メンテ依頼 & "') AND ISNULL(REPRT_COMPD,'') = '')" & Chr(13)
'        strSQL = strSQL & "     )" & Chr(13)
'        'INSERT 2021/08/03 N.IMAI End
'        'INSERT 2023/02/27 N.IMAI Start
'        strSQL = strSQL & " AND ( "
'        strSQL = strSQL & "      (REPRT_TYPEC NOT IN ('" & P_修繕_修理 & "', '" & P_修繕_メンテ依頼 & "', '" & P_修繕_未設置 & "', '" & P_修繕_仮置き & "')) " & Chr(13)
'        strSQL = strSQL & "   OR (REPRT_TYPEC     IN ('" & P_修繕_修理 & "', '" & P_修繕_メンテ依頼 & "', '" & P_修繕_未設置 & "', '" & P_修繕_仮置き & "') AND ISNULL(REPRT_COMPD,'') = '')" & Chr(13)
'        strSQL = strSQL & "     )" & Chr(13)
'        'INSERT 2023/02/27 N.IMAI End
        'INSERT 2025/09/12 N.IMAI Start
        strSQL = strSQL & " AND ( "
        strSQL = strSQL & "      (REPRT_TYPEC NOT IN ('" & P_修繕_修理 & "','" & P_修繕_メンテ依頼 & "','" & P_修繕_組替え & "','" & P_修繕_鍵返却 & "','" & P_修繕_社内理由 & "','" & P_修繕_オーナー理由 & "','" & P_修繕_解約予定 & "','" & P_修繕_未設置 & "','" & P_修繕_仮置き & "')) " & Chr(13)
        strSQL = strSQL & "   OR (REPRT_TYPEC     IN ('" & P_修繕_修理 & "','" & P_修繕_メンテ依頼 & "','" & P_修繕_組替え & "','" & P_修繕_鍵返却 & "','" & P_修繕_社内理由 & "','" & P_修繕_オーナー理由 & "','" & P_修繕_解約予定 & "','" & P_修繕_未設置 & "','" & P_修繕_仮置き & "') AND ISNULL(REPRT_COMPD,'') = '')" & Chr(13)
        strSQL = strSQL & "     )" & Chr(13)
        'INSERT 2025/09/12 N.IMAI End
    End If
      
    ' 並び順の指定
    strSQL = strSQL & " ORDER BY REPRT_GENTD DESC, REPRT_INSED DESC, REPRT_INSEJ DESC"
    
    ' 組み立てたSQLでデータ取得
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    
    If rsObject.EOF = False Then
    ' 取得データを元に文言を組み立てる
      With rsObject
        '『発生日』を YY/MM 形式で付加
        returnText = Mid(.Fields("REPRT_GENTD"), 3, 2) & "/" & Mid(.Fields("REPRT_GENTD"), 5, 2) & " "
        ' 修理中なら『修理区分名』、それ以外なら『内容１』を付加
        If .Fields("REPRT_TYPEC") = P_修繕_修理 _
        Or .Fields("REPRT_TYPEC") = P_修繕_メンテ依頼 _
        Or .Fields("REPRT_TYPEC") = P_修繕_組替え _
        Or .Fields("REPRT_TYPEC") = P_修繕_鍵返却 _
        Or .Fields("REPRT_TYPEC") = P_修繕_社内理由 _
        Or .Fields("REPRT_TYPEC") = P_修繕_オーナー理由 _
        Or .Fields("REPRT_TYPEC") = P_修繕_解約予定 _
        Or .Fields("REPRT_TYPEC") = P_修繕_未設置 _
        Or .Fields("REPRT_TYPEC") = P_修繕_仮置き Then                      'UPDATE 2025/09/12 N.IMAI 'INSERT 2021/08/03 N.IMAI
          returnText = returnText & .Fields("REPRC_NAME")
        Else
          returnText = returnText & .Fields("REPRT_CNT1N")
        End If
      End With
    End If
    
   rsObject.Close
   Set rsObject = Nothing
   GetRoomLatestReprText = returnText
   Exit Function
                              
Exception:
  GetRoomLatestReprText = returnText
  If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
  Call Err.Raise(Err.Number, "GetRoomLatestReprText" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :顧客の直近顧客問合わせ内容取得
'        MODULE_ID        :GetUserLatestKotoText
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String)  = a部門コード
'                         :第3引数(Variant) = a顧客コード(6桁ゼロサプレス済み前提)
'        Retrun           :取得した直近問合わせ内容文言
'        NOTE             :
'        CREATE_DATE      :2006/09/02
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetUserLatestKotoText(dbSQLServer As DAO.Database, _
                                      a部門コード As String, _
                                      a顧客コード As Variant _
                                      ) As String
    Dim strSQL     As String
    Dim returnText As String
    Dim rsObject   As Recordset
     
    returnText = ""
    
    On Error GoTo Exception
  
    ' 指定した顧客の最新問合わせ情報を１件取得するSQL
    strSQL = "SELECT TOP 1 KOTOT_GENTD,KOTOT_CONTN " & Chr(13)
    strSQL = strSQL & "FROM KOTO_TRAN " & Chr(13)
    strSQL = strSQL & "WHERE KOTOT_BUMOC = '" & a部門コード & "'" & Chr(13)
    strSQL = strSQL & " AND KOTOT_KOKYC = '" & a顧客コード & "'" & Chr(13)
    strSQL = strSQL & " AND KOTOT_DISPI = 1 " & Chr(13)                         'INSERT 2010/06/10 SHIBAZAKI
      
    ' 並び順の指定
    strSQL = strSQL & " ORDER BY KOTOT_GENTD DESC, KOTOT_INSED DESC, KOTOT_INSEJ DESC"
    
    ' 組み立てたSQLでデータ取得
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    
    If rsObject.EOF = False Then
    ' 取得データを元に文言を組み立てる
      With rsObject
        '『発生日』を YY/MM 形式で付加
        returnText = Mid(.Fields("KOTOT_GENTD"), 3, 2) & "/" & Mid(.Fields("KOTOT_GENTD"), 5, 2) & " "
        ' 『内容』を付加
        returnText = returnText & .Fields("KOTOT_CONTN")
      End With
    End If
    
   rsObject.Close
   Set rsObject = Nothing
   GetUserLatestKotoText = returnText
   Exit Function
                              
Exception:
  GetUserLatestKotoText = returnText
  If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
  Call Err.Raise(Err.Number, "GetUserLatestKotoText" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function
'==============================================================================*
'
'        MODULE_NAME      :指定した部屋が修理中なのか判断する
'        MODULE_ID        :IsRepairing
'        PARAMETER        :
'        CREATE_DATE      :2006/09/06
'        NOTE             :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function IsRepairing(dbSQLServer As DAO.Database, _
                                 a部門コード As String, _
                                 aヤードコード As Variant, _
                                 a部屋番号 As Variant, _
                                 Optional a除外したいレコード As Variant = Null) As Boolean
    Dim strSQL    As String
    Dim rsObject  As Recordset
    Dim isRet     As Boolean

  On Error GoTo Exception
    isRet = False
    
    ' 修理中が何件あるか求めるSQL
    strSQL = "SELECT COUNT(*) CNT FROM REPR_TRAN WHERE"
    If Nz(a除外したいレコード, "") <> "" Then
        strSQL = strSQL & " REPRT_UNIQEC <> '" & a除外したいレコード & "' AND " & Chr(13)
    End If
    strSQL = strSQL & " REPRT_BUMOC ='" & a部門コード & "' " & Chr(13)
    strSQL = strSQL & " AND REPRT_YARDC ='" & aヤードコード & "' " & Chr(13)
    strSQL = strSQL & " AND REPRT_ROOMC ='" & a部屋番号 & "' " & Chr(13)
    strSQL = strSQL & " AND REPRT_INPTI ='" & P_対象_部屋 & "' " & Chr(13)
    'strSql = strSql & " AND REPRT_TYPEC ='" & P_修繕_修理 & "' " & Chr(13)                                 'DELETE 2021/08/03 N.IMAI
    'strSQL = strSQL & " AND REPRT_TYPEC IN('" & P_修繕_修理 & "','" & P_修繕_メンテ依頼 & "') " & Chr(13)   'INSERT 2021/08/03 N.IMAI
    'strSQL = strSQL & " AND REPRT_TYPEC IN('" & P_修繕_修理 & "','" & P_修繕_メンテ依頼 & "','" & P_修繕_未設置 & "','" & P_修繕_仮置き & "') " & Chr(13)   'INSERT 2023/02/27 N.IMAI
    strSQL = strSQL & " AND REPRT_TYPEC IN('" & P_修繕_修理 & "','" & P_修繕_メンテ依頼 & "','" & P_修繕_組替え & "','" & P_修繕_鍵返却 & "','" & P_修繕_社内理由 & "','" & P_修繕_オーナー理由 & "','" & P_修繕_解約予定 & "','" & P_修繕_未設置 & "','" & P_修繕_仮置き & "') " & Chr(13) 'INSERT 2025/09/12 N.IMAI
    
    strSQL = strSQL & " AND  ISNULL(REPRT_COMPD,'') = '' "
   
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
            
    If 0 < rsObject.Fields("CNT") Then
        isRet = True
    End If
    
    rsObject.Close
    Set rsObject = Nothing
    
    IsRepairing = isRet
    Exit Function
Exception:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    Call Err.Raise(Err.Number, "IsRepairing" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :テスト用
'        MODULE_ID        :testCmReprMod
'        Parameter        :
'        CREATE_DATE      :2006/05/13
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function testCmReprMod() As Boolean

    Dim strSQL     As String
    Dim dbObject   As Database  '加瀬DBオブジェクト
     
    Dim strDataSource As String
    Dim stReprt       As Type_REPR_TRAN
        
    testCmReprMod = False
    
    On Error GoTo Exception
  
    ' 加瀬DB接続
    strDataSource = GetDataSource("ODBC_DATA_SOURCE_NAME")

    If strDataSource = "" Then
        ' テーブル[SETU_TABL]の設定不正
      GoTo Exception
    End If

    Set dbObject = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00())
 
    stReprt.BUMOC = "P"
    stReprt.YARDC = "000012"
    stReprt.INPTI = "1"
    stReprt.TANTC = "123"
    stReprt.GENTD = "20060513"
    stReprt.TYPEC = "01"
    stReprt.INSED = "20060513"
    stReprt.INSEJ = "1221"
    stReprt.INSPB = "test"
    stReprt.INSUB = "tajima"
 
'    Call InsertReprTran(dbObject, stReprt)
 
    stReprt.UNIQEC = "CBDBA445-7B6D-4FF7-AE57-134994AD34B5"
'    Call UpdateReprTran(dbObject, stReprt)
 
 
    dbObject.Close: Set dbObject = Nothing
    
    testCmReprMod = True
    Exit Function
  
Exception:
  If Not dbObject Is Nothing Then dbObject.Close: Set dbObject = Nothing
  Call Err.Raise(Err.Number, "InsertKeyReprTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)


End Function

'==============================================================================*
'
'        MODULE_NAME      :ODBC接続文字列取得
'        MODULE_ID        :GetDataSource
'        Parameter        :anODBCNAME ODBC名
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetDataSource(ByVal anODBCNAME As String) As String

    Dim strName As String

    On Error GoTo Exception

    GetDataSource = ""

    strName = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = '" & anODBCNAME & "'"))

    GetDataSource = strName

    Exit Function

Exception:
  Call Err.Raise(Err.Number, "GetDataSource" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'****************************  ended or program ********************************





