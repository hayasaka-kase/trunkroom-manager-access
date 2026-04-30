Attribute VB_Name = "GenYardPage"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME :
'
'        PROGRAM_NAME    : ヤードWebページ作成処理
'        PROGRAM_ID      : GebYardPage
'        PROGRAM_KBN     :
'
'        CREATE          : 2008/02/15
'        CERATER         : tajima
'        Ver             : 0.0
'
'        UPDATE          : 2008/10/6
'        UPDATER         : hirano
'        Ver             : 0.1
'                        :宣言部  ヤードWebページ生成のための情報構造体 Type_YARD_WP_INFにサービス１，２，３，期間を追加
'                         Function GenerateYardPageEx セット処理追加
'
'        UPDATE          : 2009/01/30
'        UPDATER         : Suzuki
'        Ver             : 0.2
'                        :パンくずリスト追加
'
'        UPDATE          : 2009/02/27
'        UPDATER         : hirano
'        Ver             : 0.3
'                        :近隣ヤード表示数の変更10->8
'
'        UPDATE          : 2010/04/05
'        UPDATER         : M.HONDA
'        Ver             : 0.4
'                        : WEB課依頼
'                          アップロードのファイルをUTF-8対応
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "GenYardPage"
Private Const INTIF_PROGB = "GenYardP"
Private Const P_トランク部門 As String = "8"
Private Const P_配置図抽出_KEYWORD As String = "配置"
Private Const P_寸法表抽出_KEYWORD As String = "寸法"
'2009/02/27 MOD <S> hirano 近隣ヤード表示数の変更10->8
'Public Const TYPE_配列標準_SIZE As Integer = 10
Public Const TYPE_配列標準_SIZE As Integer = 8
'2009/02/27 MOD <E> hirano

' 各ファイルの格納場所 "INTI_FILE"の"BASE_PATH"を基点とした場所を差す
Private Const P_コンテナ配置図_FOLDER As String = "layout_cntn/"
Private Const P_コンテナ画像_FOLDER As String = "img_cntn/"
Private Const P_トランク配置図_FOLDER As String = "layout_trnk/"
Private Const P_トランク寸法表_FOLDER As String = "size_trnk/"
Private Const P_トランク画像_FOLDER As String = "img_trnk/"

'パンくずリスト用
Private Const P_STATES_TK As String = "東京"
Private Const P_STATES_KG As String = "神奈川"
Private Const P_STATES_CB As String = "千葉"
Private Const P_STATES_ST As String = "埼玉"
Private Const P_STATES_SZ As String = "静岡"

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private pobjKONT_DB               As Database                   ' コンテナDB

' ヤードWebページ生成のための情報構造体
' ヤード情報
Public Type Type_YARD_WP_INF
    code        As String
    NAME        As String
    ADDRESS     As String
    CAMPAIGN    As String
    NOTE        As String   '付近メモ
    HPNOTE      As String   'HP用備考
    IDO         As String
    KEIDO       As String
    SEV1N       As String   '2008/10/06 サービス１追加      Rev0.1
    SEV2N       As String   '2008/10/06 サービス２追加      Rev0.1
    SEV3N       As String   '2008/10/10 サービス３追加      Rev0.1
    ENDEN       As String   '2008/10/06 サービス期間追加    Rev0.1
End Type

' 近隣情報
Public Type Type_NYAR_WP_INF
    code        As String
    NAME        As String
End Type

' 物件情報
Public Type Type_TYPE_WP_INF
    SIZE        As String
    STEP        As String
    WIDTH       As Single
    DEPTH       As Single
    HEIGHT      As Single
    PRICE       As String
End Type

'==============================================================================*
'
'        MODULE_NAME      :ヤード情報構造体設定
'        MODULE_ID        :SetYardWPInf
'        Parameter        :
'        戻り値           :Nothing
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetYardWPInf(ByRef aYARD_WP_INF As Type_YARD_WP_INF, _
                        aCode As String, _
                        aName As String, _
                        anAddress As String, _
                        aCampaign As Variant, _
                        aNote As Variant, _
                        aHPnote As Variant, _
                        anIdo As Variant, _
                        aKeido As Variant)

    aYARD_WP_INF.code = aCode
    aYARD_WP_INF.NAME = aName
    aYARD_WP_INF.ADDRESS = anAddress
    aYARD_WP_INF.CAMPAIGN = Nz(aCampaign, "")
    aYARD_WP_INF.NOTE = Nz(aNote, "")
    aYARD_WP_INF.HPNOTE = Nz(aHPnote, "")
    aYARD_WP_INF.IDO = Nz(anIdo, "0")
    aYARD_WP_INF.KEIDO = Nz(aKeido, "0")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :近隣情報情報構造体設定
'        MODULE_ID        :SetNyarWPInf
'        Parameter        :
'        戻り値           :Nothing
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetNyarWPInf(ByRef aNEAR_WP_INF As Type_NYAR_WP_INF, _
                        aCode As String, _
                        aName As String)

    aNEAR_WP_INF.code = aCode
    aNEAR_WP_INF.NAME = aName

End Sub

'==============================================================================*
'
'        MODULE_NAME      :物件情報構造体設定
'        MODULE_ID        :SetTypeWPInf
'        Parameter        :
'        戻り値           :Nothing
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetTypeWPInf(ByRef aTYPE_WP_INF As Type_TYPE_WP_INF, _
                        aSize As String, _
                        aStep As Variant, _
                        aWidth As Variant, _
                        aDepth As Variant, _
                        aHeight As Variant, _
                        aPrice As Variant _
                       )

    aTYPE_WP_INF.SIZE = aSize
    aTYPE_WP_INF.STEP = Nz(aStep, "")
    aTYPE_WP_INF.WIDTH = Nz(aWidth, 0)
    aTYPE_WP_INF.DEPTH = Nz(aDepth, 0)
    aTYPE_WP_INF.HEIGHT = Nz(aHeight, 0)
    aTYPE_WP_INF.PRICE = Format$(aPrice, "#,##0")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ヤード情報読込
'        MODULE_ID        :ReadReprTran
'        Parameter        :第1引数(ADOコネクション）
'                         :第2引数(ByRef構造体) = ヤード情報内容※キーを設定しておくこと
'        戻り値           : True...読込成功
'                         : False...対象無し
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function ReadYardMast(aConnection As Object, _
                               ByRef aYARD_WP_INF As Type_YARD_WP_INF _
                              ) As Boolean
    Dim strSQL  As String
    Dim rsData  As Object
     
    ReadYardMast = False
    
    On Error GoTo Exception
  
    strSQL = "SELECT * FROM YARD_MAST WHERE YARD_CODE = '" & aYARD_WP_INF.code & "'"
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, aConnection)
    
    ' 対象確認
    If rsData.EOF = False Then
    ' 対象を構造体にセッタップ！
      With rsData
        Call SetYardWPInf(aYARD_WP_INF, _
                        .Fields("YARD_CODE"), _
                        .Fields("YARD_NAME"), _
                        Nz(.Fields("YARD_ADDR_1"), "") & Nz(.Fields("YARD_ADDR_2"), "") & Nz(.Fields("YARD_ADDR_1"), ""), _
                        Nz(.Fields("YARD_SEV1N"), "") & Nz(.Fields("YARD_SEV2N"), "") & Nz(.Fields("YARD_ENDEN"), ""), _
                        .Fields("YARD_NOTE"), _
                        .Fields("YARD_WP_NOTE"), _
                        .Fields("YARD_IDO"), _
                        .Fields("YARD_KEIDO") _
                    )
      End With
      ReadYardMast = True
    End If
    
   rsData.Close
   Set rsData = Nothing
   
   Exit Function
                              
Exception:
  If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
  Call Err.Raise(Err.Number, "ReadYardMast" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :近隣ヤード情報読込
'        MODULE_ID        :ReadNyarMast
'        Parameter        :第1引数(ADOコネクション）
'                         :第2引数(ヤードコード） ←取得キー
'                         :第3引数(ByRef構造体) => 近隣情報内容
'                         :第4引数(Optional) ←部門コード
'        戻り値           : True...読込成功
'                         : False...対象無し
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function ReadNyarMast(aConnection As Object, _
                               aYardCode As Long, _
                               ByRef aNYAR_WP_INF() As Type_NYAR_WP_INF, _
                               Optional aBumonCode As String = "H" _
                              ) As Boolean
                              
    Dim strSQL  As String
    Dim rsData  As Object
    Dim idx     As Integer
    Dim yardCodeFormat As String
     
    ReadNyarMast = False
    
    On Error GoTo Exception
  
    ' 近隣ヤード情報から直近を【TYPE_配列標準_SIZE】分10件まで取得
    strSQL = "SELECT TOP " & TYPE_配列標準_SIZE & Chr(13)
    strSQL = strSQL & " NYAR_NCODE ,YARD_NAME " & Chr(13)
    strSQL = strSQL & " FROM NYAR_MAST INNER JOIN YARD_MAST ON NYAR_NCODE = YARD_CODE " & Chr(13)
    strSQL = strSQL & " WHERE NYAR_YCODE = " & aYardCode & Chr(13)
    strSQL = strSQL & "   AND NYAR_NCODE <> " & aYardCode & Chr(13)
    strSQL = strSQL & "   AND ISNULL(YARD_RENTEND_DAY, '9999/12/31') > GETDATE()" & Chr(13) '営業中ヤードのみ対象
    strSQL = strSQL & "   AND YARD_NETUSE_KBN = -1 " & Chr(13)  'ネット予約可のみ対象
    strSQL = strSQL & " ORDER BY NYAR_KIRO"
    
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, aConnection)
    ' 対象取得...取得できた分且つMAX【TYPE_配列標準_SIZE】分10件件まで
    idx = 0
    yardCodeFormat = "000000" '6桁ゼロサプレス
    If aBumonCode = P_トランク部門 Then yardCodeFormat = "00000" 'トランクは5桁ゼロサプレス
        
    While Not rsData.EOF And idx < TYPE_配列標準_SIZE
        ' 対象を構造体にセッタップ！
        Call SetNyarWPInf(aNYAR_WP_INF(idx), _
                        Format$(rsData.Fields("NYAR_NCODE"), yardCodeFormat), _
                        rsData.Fields("YARD_NAME"))
        rsData.MoveNext
        idx = idx + 1
    Wend
    
   rsData.Close
   Set rsData = Nothing
   
   Exit Function
                              
Exception:
  If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
  Call Err.Raise(Err.Number, "ReadNyarMast" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :タイプ情報読込
'        MODULE_ID        :ReadTypeData
'        Parameter        :第1引数(ADOコネクション）
'                         :第2引数 ヤードコード ←取得キー
'                         :第2引数(ByRef構造体) = タイプ情報配列
'        戻り値           : True...読込成功
'                         : False...対象無し
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function ReadTypeData(aConnection As Object, _
                               aYardCode As Long, _
                               ByRef aTYPE_WP_INF() As Type_TYPE_WP_INF _
                              ) As Boolean
                              
    Dim strSQL  As String
    Dim rsData  As Object
    Dim idx     As Integer
    Dim intMax  As Integer
     
    ReadTypeData = False
    
    On Error GoTo Exception
  
    ' コンテナマスタ、価格表テーブルから取得
    strSQL = "SELECT CNTA_SIZE ,NAME_NAME, PRIC_PRICE " & Chr(13)
    strSQL = strSQL & ",MIN(CNTA_DEPTH) DEPTH, MIN(CNTA_WIDTH) WIDTH, MIN(CNTA_HEIGHT) HEIGHT" & Chr(13)
    strSQL = strSQL & " FROM CNTA_MAST LEFT OUTER JOIN PRIC_TABL ON CNTA_CODE = PRIC_YCODE " & Chr(13)
    strSQL = strSQL & "     AND CNTA_STEP = PRIC_STEP AND CNTA_USAGE = PRIC_USAGE AND CNTA_SIZE = PRIC_SIZE " & Chr(13)
    strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST ON NAME_ID = '090' AND NAME_CODE = CNTA_STEP " & Chr(13)
    strSQL = strSQL & " WHERE CNTA_CODE = " & aYardCode & Chr(13)
    strSQL = strSQL & "   AND CNTA_USE <> 9 " & Chr(13)
    strSQL = strSQL & "   AND CNTA_USAGE = 0" & Chr(13) '部門によって変わるが現状コンテナのみ対象
    strSQL = strSQL & " GROUP BY CNTA_SIZE , CNTA_STEP, NAME_NAME, PRIC_PRICE " & Chr(13)
    strSQL = strSQL & " ORDER BY CNTA_SIZE , CNTA_STEP, PRIC_PRICE"
    
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, aConnection)
    idx = 0
    intMax = TYPE_配列標準_SIZE
    While Not rsData.EOF
        ' もし標準MAX以上ならば配列サイズを増やしておく
        If intMax <= idx Then
              ReDim Preserve aTYPE_WP_INF(UBound(aTYPE_WP_INF) + TYPE_配列標準_SIZE) As GenYardPage.Type_TYPE_WP_INF
              intMax = intMax + TYPE_配列標準_SIZE
        End If
        
        ' 対象を構造体にセッタップ！
        Call SetTypeWPInf(aTYPE_WP_INF(idx), _
                        CStr(rsData.Fields("CNTA_SIZE")), _
                        rsData.Fields("NAME_NAME"), _
                        rsData.Fields("WIDTH"), _
                        rsData.Fields("DEPTH"), _
                        rsData.Fields("HEIGHT"), _
                        rsData.Fields("PRIC_PRICE") _
                        )
        rsData.MoveNext
'    Debug.Print idx & ":" & aTYPE_WP_INF(idx).SIZE & ":" & aTYPE_WP_INF(idx).PRICE
        idx = idx + 1
    Wend
    
   rsData.Close
   Set rsData = Nothing
   ReadTypeData = True
   Exit Function
                              
Exception:
  If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
  Call Err.Raise(Err.Number, "ReadTypeData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :トランクｎタイプ料金読込
'        MODULE_ID        :ReadTrnkPrices
'        Parameter        :第1引数(ADOコネクション）
'                         :第2引数 ヤードコード ←取得キー
'                         :第2引数(ByRef String) = nタイプ料金情報配列
'        戻り値           : True...読込成功
'                         : False...対象無し
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function ReadTrnkPrices(aConnection As Object, _
                               aYardCode As Long, _
                               ByRef aPrice() As String _
                              ) As Boolean
                              
    Dim strSQL  As String
    Dim rsData  As Object
     
    ReadTrnkPrices = False
    
    On Error GoTo Exception
  
    ' コンテナマスタ、価格表テーブルから４サイズ範囲毎に取得
    strSQL = "SELECT " & Chr(13)
    strSQL = strSQL & "( SELECT MIN(PRIC_PRICE) FROM  CNTA_MAST LEFT OUTER JOIN PRIC_TABL ON CNTA_CODE = PRIC_YCODE" & Chr(13)
    strSQL = strSQL & "  AND CNTA_STEP = PRIC_STEP AND CNTA_USAGE = PRIC_USAGE AND CNTA_SIZE = PRIC_SIZE " & Chr(13)
    strSQL = strSQL & "  WHERE CNTA_CODE = " & aYardCode & " AND CNTA_USE <> 9 " & Chr(13)
    strSQL = strSQL & "    AND CNTA_USAGE = 1 " & Chr(13) 'トランクで
    strSQL = strSQL & "    AND CNTA_SIZE < 1 " & Chr(13)  '1帖以下
    strSQL = strSQL & ") type1," & Chr(13)
    
    strSQL = strSQL & "( SELECT MIN(PRIC_PRICE) FROM  CNTA_MAST LEFT OUTER JOIN PRIC_TABL ON CNTA_CODE = PRIC_YCODE" & Chr(13)
    strSQL = strSQL & "  AND CNTA_STEP = PRIC_STEP AND CNTA_USAGE = PRIC_USAGE AND CNTA_SIZE = PRIC_SIZE " & Chr(13)
    strSQL = strSQL & "  WHERE CNTA_CODE = " & aYardCode & " AND CNTA_USE <> 9 " & Chr(13)
    strSQL = strSQL & "    AND CNTA_USAGE = 1 " & Chr(13) 'トランクで
    strSQL = strSQL & "    AND 1 <= CNTA_SIZE AND CNTA_SIZE < 1.5 " & Chr(13) '1～1.5帖
    strSQL = strSQL & ") type2," & Chr(13)
    
    strSQL = strSQL & "( SELECT MIN(PRIC_PRICE) FROM  CNTA_MAST LEFT OUTER JOIN PRIC_TABL ON CNTA_CODE = PRIC_YCODE" & Chr(13)
    strSQL = strSQL & "  AND CNTA_STEP = PRIC_STEP AND CNTA_USAGE = PRIC_USAGE AND CNTA_SIZE = PRIC_SIZE " & Chr(13)
    strSQL = strSQL & "  WHERE CNTA_CODE = " & aYardCode & " AND CNTA_USE <> 9 " & Chr(13)
    strSQL = strSQL & "    AND CNTA_USAGE = 1 " & Chr(13) 'トランクで
    strSQL = strSQL & "    AND 1.5 <= CNTA_SIZE AND CNTA_SIZE < 1.9  " & Chr(13) '1.5～1.9帖
    strSQL = strSQL & ") type3," & Chr(13)
    
    strSQL = strSQL & "( SELECT MIN(PRIC_PRICE) FROM  CNTA_MAST LEFT OUTER JOIN PRIC_TABL ON CNTA_CODE = PRIC_YCODE" & Chr(13)
    strSQL = strSQL & "  AND CNTA_STEP = PRIC_STEP AND CNTA_USAGE = PRIC_USAGE AND CNTA_SIZE = PRIC_SIZE " & Chr(13)
    strSQL = strSQL & "  WHERE CNTA_CODE = " & aYardCode & " AND CNTA_USE <> 9 " & Chr(13)
    strSQL = strSQL & "    AND CNTA_USAGE = 1 " & Chr(13) 'トランクで
    strSQL = strSQL & "    AND 2 <= CNTA_SIZE   " & Chr(13) '2帖以上
    strSQL = strSQL & ") type4"
    
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, aConnection)
    
    If Not rsData.EOF Then
        aPrice(0) = convertPrice(rsData.Fields("type1"))
        aPrice(1) = convertPrice(rsData.Fields("type2"))
        aPrice(2) = convertPrice(rsData.Fields("type3"))
        aPrice(3) = convertPrice(rsData.Fields("type4"))
    End If
    
    rsData.Close
    Set rsData = Nothing
    ReadTrnkPrices = True
    Exit Function
                              
Exception:
  If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
  Call Err.Raise(Err.Number, "ReadTrnkPrices" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :金額の変換処理
'        MODULE_ID        :convertPrice
'        Parameter        :第1引数(DB取得項目）
'        戻り値           : True...変換結果
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function convertPrice(aPrice As Variant) As String
    
    If Nz(aPrice, "") = "" Then
        convertPrice = "－"
    Else
        convertPrice = Format$(aPrice, "#,##0") & "円"
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :ヤード詳細ページ生成
'        MODULE_ID        :GenerateYardPage
'        Parameter        :第1引数(String)          = 部門コード
'                         :第2引数(ByRef構造体)     = ヤード情報内容
'                         :第3引数(ByRef構造体配列) = 近隣情報内容
'                         :第4引数(ByRef構造体配列) = タイプ情報内容(コンテナ用）
'                         :第5引数(String配列)      = タイプ情報内容(トランク用）
'        戻り値           : True...成功
'                         : False...失敗
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GenerateYardPage(aBumonCode As String, _
                                 aYARD_WP_INF As Type_YARD_WP_INF, _
                                 aNYAR_WP_INF() As Type_NYAR_WP_INF, _
                                 aTYPE_WP_INF() As Type_TYPE_WP_INF, _
                                 aTrnkPrices() As String _
                              ) As Boolean
    Dim sourceFile  As String
    Dim outputFile  As String
    Dim subFolder   As String
    Dim ouputPath   As String

    On Error GoTo Exception

    'テンプレートファイルパスとファイル名の取得
    sourceFile = getINTIF_RECDB("SOURCE_PATH")
    subFolder = getINTIF_RECDB("BASE_PATH")
    subFolder = Replace(subFolder, "/", "\")
    
    If aBumonCode = P_トランク部門 Then
        sourceFile = sourceFile & subFolder & "Yard8Template.htm"
    Else
        sourceFile = sourceFile & subFolder & "YardHTemplate.htm"
    End If
    '取得したファイル名の有無を確認
    If Dir$(sourceFile) = "" Then
        Call MSZZ024_M10("GenerateYardPage", sourceFile & "にテンプレートファイルが存在しません。")
    End If
    
    '出力先の取得
    ouputPath = getINTIF_RECDB("GENERATE_PATH")
    ouputPath = ouputPath & subFolder
    'ファイル名決め
    If aBumonCode = P_トランク部門 Then
        'トランクのみヤードコードはゼロサプレス５桁
        outputFile = ouputPath & Format$(CLng(aYARD_WP_INF.code), "00000") & ".htm"
    Else
        outputFile = ouputPath & aYARD_WP_INF.code & ".htm"
    End If
    
    'ページ生成本体の呼出
    GenerateYardPage = GenerateYardPageEx( _
                                           aBumonCode, _
                                           sourceFile, _
                                           outputFile, _
                                           aYARD_WP_INF, _
                                           aNYAR_WP_INF, _
                                           aTYPE_WP_INF, _
                                           aTrnkPrices)

    Exit Function
Exception:
    Call Err.Raise(Err.Number, "GenerateYardPage" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
   
'==============================================================================*
'
'        MODULE_NAME      :ヤード詳細ページ生成（本体）
'        MODULE_ID        :GenerateYardPageEx
'        Parameter        :第1引数(String)          = 部門コード
'                         :第2引数(String)          = ページ生成先＆ファイル名
'                         :第3引数(String)          = ページ生成先＆ファイル名
'                         :第4引数(ByRef構造体)     = ヤード情報内容
'                         :第5引数(ByRef構造体配列) = 近隣情報内容
'                         :第6引数(ByRef構造体配列) = タイプ情報内容(コンテナ用）
'                         :第7引数(String配列)      = タイプ情報内容(トランク用）
'        戻り値           : True...成功
'                         : False...失敗
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GenerateYardPageEx(aBumonCode As String, _
                                 aTemplatePathFile As String, _
                                 aGeneratePathFile As String, _
                                 aYARD_WP_INF As Type_YARD_WP_INF, _
                                 aNYAR_WP_INF() As Type_NYAR_WP_INF, _
                                 aTYPE_WP_INF() As Type_TYPE_WP_INF, _
                                 aTrnkPrices() As String _
                              ) As Boolean
    
    Dim isException     As Boolean  'エラー発生検知フラグ
    Dim textLine        As String   '読込→出力内容
    Dim loopControl     As Integer  '繰り返しステートメント状態
    Dim textLoopLine    As String   '繰り返し部の出力テンプレート
    Dim index           As Integer  '情報配列の添字
    
'< start > パンくず対応 2009/01/29 - EGL UK
    Dim url             As String   '
    Dim STATES          As String   '
'< end > パンくず対応 2009/01/29 - EGL UK
    
    '' 20100405 Ver0.4 M.HONDA START
    Dim txt As Object
    Dim txt2 As Object
    Const adReadLine = -2
    Const adSaveCreateOverWrite = 2
    Const adWriteLine = 1
    '' 20100405 Ver0.4 M.HONDA END
    
    GenerateYardPageEx = False
    isException = False
    
    On Error GoTo Exception
    
'< start > パンくず対応 2009/01/29 - EGL UK
    Call GetBreadCrumb(aYARD_WP_INF.ADDRESS, url, STATES)
'< end > パンくず対応 2009/01/29 - EGL UK
    
    
    '' 20100405 Ver0.4 M.HONDA START
    ''' WEB課の依頼で文字コードUTF-8対応
    '' 読込み用
    
    Set txt = CreateObject("ADODB.Stream")

    '' 書込み用
    Set txt2 = CreateObject("ADODB.Stream")

    '文字列型のオブジェクトの文字コードを指定する
    txt.Charset = "UTF-8"
    'オブジェクトのインスタンスを作成
    txt.Open

    '文字列型のオブジェクトの文字コードを指定する
    txt2.Charset = "UTF-8"
    'オブジェクトのインスタンスを作成
    txt2.Open


    'ファイルからデータを読み込む
    txt.LoadFromFile (aTemplatePathFile)

    '最終行までループする
    Do While Not txt.EOS

        '次の行を読み取る
        textLine = txt.ReadText(adReadLine)

        If 0 < InStr(textLine, "<!--@LOOP(nyar_data)-->") Then
        ' 近隣情報置き換えループ検知
            loopControl = 1 '近隣は１デス、オレは定数にはせんよ
            textLoopLine = ""
        ElseIf 0 < InStr(textLine, "<!--@LOOP(type_data)-->") Then
        ' タイプ情報置き換えループ検知
            loopControl = 2 'タイプは２デス、オレは定数にはせんよ
            textLoopLine = ""
        ElseIf 0 < InStr(textLine, "<!--@END_LOOP-->") Then
        ' ループ制御終わり、溜めた情報を全て出力
        ' ※ホントはサブルーチン化したかったけど#を引数で渡し方が不明・・
            If loopControl = 1 Then
            '近隣情報を配列分置き換える、最大【TYPE_配列標準_SIZE】件
                index = 0
                Do While aNYAR_WP_INF(index).code <> "" Or index > TYPE_配列標準_SIZE
                    textLine = textLoopLine
                    textLine = Replace$(textLine, "{nyar_code}", aNYAR_WP_INF(index).code)
                    textLine = Replace$(textLine, "{nyar_name}", aNYAR_WP_INF(index).NAME)
                    txt2.WriteText textLine, adWriteLine
                    index = index + 1
                Loop
            End If

            If loopControl = 2 Then
            'タイプ情報を配列分置き換える、最大件数はあるぶん
                index = 0
                Do While aTYPE_WP_INF(index).SIZE <> "" Or index > UBound(aTYPE_WP_INF)
                    textLine = textLoopLine
                    textLine = Replace$(textLine, "{cnta_size}", aTYPE_WP_INF(index).SIZE)
                    textLine = Replace$(textLine, "{cnta_step}", aTYPE_WP_INF(index).STEP)
                    textLine = Replace$(textLine, "{cnta_width}", aTYPE_WP_INF(index).WIDTH)
                    textLine = Replace$(textLine, "{cnta_depth}", aTYPE_WP_INF(index).DEPTH)
                    textLine = Replace$(textLine, "{cnta_height}", aTYPE_WP_INF(index).HEIGHT)
                    textLine = Replace$(textLine, "{pric_price}", aTYPE_WP_INF(index).PRICE)
                    txt2.WriteText textLine, adWriteLine
                    index = index + 1
                Loop

            End If
            loopControl = 0

        ElseIf loopControl > 0 Then
        ' テンプレートの繰返情報を溜める
            textLoopLine = textLoopLine & textLine
        Else
        '制御キーワード無し行ならば
            ' ヤード情報置き換えまたはテンプレート内容出力
            textLine = Replace$(textLine, "{yard_name}", aYARD_WP_INF.NAME)
            textLine = Replace$(textLine, "{yard_address}", aYARD_WP_INF.ADDRESS)
            textLine = Replace$(textLine, "{yard_code}", aYARD_WP_INF.code)
            textLine = Replace$(textLine, "{yard_campaign}", aYARD_WP_INF.CAMPAIGN)
            textLine = Replace$(textLine, "{yard_fukinmemo}", aYARD_WP_INF.NOTE)
            textLine = Replace$(textLine, "{yard_hpmemo}", aYARD_WP_INF.HPNOTE)
            textLine = Replace$(textLine, "{yard_ido}", aYARD_WP_INF.IDO)
            textLine = Replace$(textLine, "{yard_keido}", aYARD_WP_INF.KEIDO)
            '2008/10/06 add ▼ページ出力の際、サービス１，２，期間情報の追加
            textLine = Replace$(textLine, "{yard_sev1n}", aYARD_WP_INF.SEV1N)   'サービス１追加
            textLine = Replace$(textLine, "{yard_sev2n}", aYARD_WP_INF.SEV2N)   'サービス２追加
            textLine = Replace$(textLine, "{yard_sev3n}", aYARD_WP_INF.SEV3N)   '2008/10/10 add サービス３追加
            textLine = Replace$(textLine, "{yard_enden}", aYARD_WP_INF.ENDEN)   'サービス期間追加
            '2008/10/06 add ▲
            'トランクのみの書換、無駄にReplaceするよりもよいだろう
            If aBumonCode = P_トランク部門 Then
                textLine = Replace$(textLine, "{trnk_price(1)}", aTrnkPrices(0))
                textLine = Replace$(textLine, "{trnk_price(2)}", aTrnkPrices(1))
                textLine = Replace$(textLine, "{trnk_price(3)}", aTrnkPrices(2))
                textLine = Replace$(textLine, "{trnk_price(4)}", aTrnkPrices(3))
            End If

'< start > パンくず対応 2009/01/29 - EGL UK
            textLine = Replace$(textLine, "{state_url}", url)                   'URL
            textLine = Replace$(textLine, "{state_name}", STATES)               '都県

'< end > パンくず対応 2009/01/29 - EGL UK

            txt2.WriteText textLine, adWriteLine

        End If

    Loop

    'オブジェクトを閉じる
    txt.Close

    'メモリからオブジェクトを削除する
    Set txt = Nothing

    'オブジェクトの内容をファイルに保存
    txt2.SaveToFile (aGeneratePathFile), adSaveCreateOverWrite

    'オブジェクトを閉じる
    txt2.Close

    'メモリからオブジェクトを削除する
    Set txt2 = Nothing
    
'    Open aTemplatePathFile For Input Access Read As #1
'    Open aGeneratePathFile For Output Access Write As #2
'
'    Do While Not EOF(1)
'        Line Input #1, textLine
'
'        If 0 < InStr(textLine, "<!--@LOOP(nyar_data)-->") Then
'        ' 近隣情報置き換えループ検知
'            loopControl = 1 '近隣は１デス、オレは定数にはせんよ
'            textLoopLine = ""
'        ElseIf 0 < InStr(textLine, "<!--@LOOP(type_data)-->") Then
'        ' タイプ情報置き換えループ検知
'            loopControl = 2 'タイプは２デス、オレは定数にはせんよ
'            textLoopLine = ""
'        ElseIf 0 < InStr(textLine, "<!--@END_LOOP-->") Then
'        ' ループ制御終わり、溜めた情報を全て出力
'        ' ※ホントはサブルーチン化したかったけど#を引数で渡し方が不明・・
'            If loopControl = 1 Then
'            '近隣情報を配列分置き換える、最大【TYPE_配列標準_SIZE】件
'                index = 0
'                Do While aNYAR_WP_INF(index).CODE <> "" Or index > TYPE_配列標準_SIZE
'                    textLine = textLoopLine
'                    textLine = Replace$(textLine, "{nyar_code}", aNYAR_WP_INF(index).CODE)
'                    textLine = Replace$(textLine, "{nyar_name}", aNYAR_WP_INF(index).NAME)
'                    Print #2, textLine
'                    index = index + 1
'                Loop
'            End If
'
'            If loopControl = 2 Then
'            'タイプ情報を配列分置き換える、最大件数はあるぶん
'                index = 0
'                Do While aTYPE_WP_INF(index).Size <> "" Or index > UBound(aTYPE_WP_INF)
'                    textLine = textLoopLine
'                    textLine = Replace$(textLine, "{cnta_size}", aTYPE_WP_INF(index).Size)
'                    textLine = Replace$(textLine, "{cnta_step}", aTYPE_WP_INF(index).Step)
'                    textLine = Replace$(textLine, "{cnta_width}", aTYPE_WP_INF(index).WIDTH)
'                    textLine = Replace$(textLine, "{cnta_depth}", aTYPE_WP_INF(index).DEPTH)
'                    textLine = Replace$(textLine, "{cnta_height}", aTYPE_WP_INF(index).HEIGHT)
'                    textLine = Replace$(textLine, "{pric_price}", aTYPE_WP_INF(index).PRICE)
'                    Print #2, textLine
'                    index = index + 1
'                Loop
'
'            End If
'            loopControl = 0
'
'        ElseIf loopControl > 0 Then
'        ' テンプレートの繰返情報を溜める
'            textLoopLine = textLoopLine & textLine
'        Else
'        '制御キーワード無し行ならば
'            ' ヤード情報置き換えまたはテンプレート内容出力
'            textLine = Replace$(textLine, "{yard_name}", aYARD_WP_INF.NAME)
'            textLine = Replace$(textLine, "{yard_address}", aYARD_WP_INF.ADDRESS)
'            textLine = Replace$(textLine, "{yard_code}", aYARD_WP_INF.CODE)
'            textLine = Replace$(textLine, "{yard_campaign}", aYARD_WP_INF.CAMPAIGN)
'            textLine = Replace$(textLine, "{yard_fukinmemo}", aYARD_WP_INF.NOTE)
'            textLine = Replace$(textLine, "{yard_hpmemo}", aYARD_WP_INF.HPNOTE)
'            textLine = Replace$(textLine, "{yard_ido}", aYARD_WP_INF.IDO)
'            textLine = Replace$(textLine, "{yard_keido}", aYARD_WP_INF.KEIDO)
'            '2008/10/06 add ▼ページ出力の際、サービス１，２，期間情報の追加
'            textLine = Replace$(textLine, "{yard_sev1n}", aYARD_WP_INF.SEV1N)   'サービス１追加
'            textLine = Replace$(textLine, "{yard_sev2n}", aYARD_WP_INF.SEV2N)   'サービス２追加
'            textLine = Replace$(textLine, "{yard_sev3n}", aYARD_WP_INF.SEV3N)   '2008/10/10 add サービス３追加
'            textLine = Replace$(textLine, "{yard_enden}", aYARD_WP_INF.ENDEN)   'サービス期間追加
'            '2008/10/06 add ▲
'            'トランクのみの書換、無駄にReplaceするよりもよいだろう
'            If aBumonCode = P_トランク部門 Then
'                textLine = Replace$(textLine, "{trnk_price(1)}", aTrnkPrices(0))
'                textLine = Replace$(textLine, "{trnk_price(2)}", aTrnkPrices(1))
'                textLine = Replace$(textLine, "{trnk_price(3)}", aTrnkPrices(2))
'                textLine = Replace$(textLine, "{trnk_price(4)}", aTrnkPrices(3))
'            End If
'
''< start > パンくず対応 2009/01/29 - EGL UK
'            textLine = Replace$(textLine, "{state_url}", URL)                   'URL
'            textLine = Replace$(textLine, "{state_name}", STATES)               '都県
'
''< end > パンくず対応 2009/01/29 - EGL UK
'
'            Print #2, textLine
'        End If
'    Loop
    '' 20100405 Ver0.4 M.HONDA END
    
    GenerateYardPageEx = True
    GoTo Finally

Exception:
    isException = True
  
Finally:
    Close 'ファイルを閉じる
    If isException = True Then
        Call Err.Raise(Err.Number, "GenerateYardPageEx" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :PDFファイルの生成
'        MODULE_ID        :GeneratePDF
'        Parameter        :第1引数(String) = 部門コード
'                         :第2引数(String) = ヤードコード
'        戻り値           : True...成功
'                         : False...失敗
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GeneratePDF(aBumonCode As String, _
                            aYardCode As String _
                            ) As Boolean
    
    Dim INTIF_RECFB値   As String
    Dim sourceFile      As String
    Dim outputPath      As String
    Dim outputFile      As String
    
    GeneratePDF = False
    
    On Error GoTo Exception

    'Ｅｘｃｅｌ表とファイル名の取得
    If aBumonCode = P_トランク部門 Then
        INTIF_RECFB値 = "TRNK_BOOK_PATH"
    Else
        INTIF_RECFB値 = "CNTN_BOOK_PATH"
    End If
    sourceFile = getINTIF_RECDB(INTIF_RECFB値)
    sourceFile = sourceFile & aYardCode & ".xls"
    '取得したファイル名の有無を確認
    If Dir$(sourceFile) = "" Then
        Exit Function
    End If
    
    '出力場所とファイル名の取得
    outputPath = getINTIF_RECDB("GENERATE_PATH")
    outputFile = getINTIF_RECDB("BASE_PATH")
    outputPath = outputPath & Replace(outputFile, "/", "\")
    
    '実際の配置図＆寸法表の生成
    ' ↓まとめてやる版
    ' GeneratePDF = generatePdfEx(aBumonCode, aYardCode, sourceFile, outputPath)
        
    Dim isRet As Boolean
    If aBumonCode = P_トランク部門 Then
        isRet = generatePdfEx2(aYardCode, sourceFile, _
                               outputPath & Replace(P_トランク配置図_FOLDER, "/", "\"), _
                               P_配置図抽出_KEYWORD)
        If isRet = False Then Exit Function
        isRet = generatePdfEx2(aYardCode, sourceFile, _
                               outputPath & Replace(P_トランク寸法表_FOLDER, "/", "\"), _
                               P_寸法表抽出_KEYWORD)
    Else
        isRet = generatePdfEx2(aYardCode, sourceFile, _
                               outputPath & Replace(P_コンテナ配置図_FOLDER, "/", "\"), _
                               P_配置図抽出_KEYWORD)
    End If
        
        
    GeneratePDF = isRet
    Exit Function
Exception:
    Call Err.Raise(Err.Number, "GeneratePDF" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Function
    
   
'==============================================================================*
'
'        MODULE_NAME      :PDFファイルの生成
'        MODULE_ID        :generatePdfEx
'        Parameter        :第1引数(String) = 部門コード
'                         :第2引数(String) = ヤードコード
'                         :第3引数(String) = 生成元ファイル
'                         :第4引数(String) = 生成先パス
'        戻り値           : True...成功
'                         : False...配置図及び寸法表が無かった
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function generatePdfEx(aBumonCode As String, _
                              aYardCode As String, _
                              aSourceFile As String, _
                              aOutputPath As String _
                            ) As Boolean
    
    Dim xlApp           As Object
    Dim xlBook          As Object
    Dim varPrintSeets   As Variant
    Dim isException     As Boolean
    Dim subFolder       As String
    
    generatePdfEx = False
    
    On Error GoTo Exception
    
    Set xlApp = CreateObject("Excel.Application")    ' Excelオブジェクトを生成する
    Set xlBook = xlApp.Workbooks.Open(aSourceFile)   'ファイルオープン

    '配置図の抽出とＰＤＦ変換
    varPrintSeets = getSheetsName(xlBook, P_配置図抽出_KEYWORD)
    If varPrintSeets(0) = "" Then GoTo Finally
    If aBumonCode = P_トランク部門 Then
        subFolder = Replace(P_トランク配置図_FOLDER, "/", "\")
    Else
        subFolder = Replace(P_コンテナ配置図_FOLDER, "/", "\")
    End If
    '対象シートをPDF変換する
    Call MSZZ038.PDFConvertEx(xlBook.Sheets(varPrintSeets), xlBook.NAME, _
                            aOutputPath & subFolder & aYardCode & ".pdf")
    
    '部門トランクならば寸法表の抽出とＰＤＦ変換
    If aBumonCode = P_トランク部門 Then
        varPrintSeets = getSheetsName(xlBook, P_寸法表抽出_KEYWORD)
        If varPrintSeets(0) = "" Then GoTo Finally
        '対象シートをPDF変換する
        Call MSZZ038.PDFConvertEx(xlBook.Sheets(varPrintSeets), xlBook.NAME, _
                            aOutputPath & Replace(P_トランク寸法表_FOLDER, "/", "\") _
                            & aYardCode & ".pdf")
    End If
    
    generatePdfEx = True
    GoTo Finally

Exception:
    isException = True
  
Finally:
    If Not xlBook Is Nothing Then Set xlBook = Nothing
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = False
        xlApp.Quit
        Set xlApp = Nothing
    End If
    If isException = True Then
        Call Err.Raise(Err.Number, "generatePdfEx" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
'==============================================================================*
'
'        MODULE_NAME      :PDFファイルの生成
'        MODULE_ID        :generatePdfEx
'        Parameter        :第1引数(String) = 部門コード
'                         :第2引数(String) = ヤードコード
'                         :第3引数(String) = 生成元ファイル
'                         :第4引数(String) = 生成先パス
'        戻り値           : True...成功
'                         : False...配置図及び寸法表が無かった
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function generatePdfEx2( _
                              aYardCode As String, _
                              aSourceFile As String, _
                              aOutputPath As String, _
                              aKeyWord As String _
                            ) As Boolean
    
    Dim xlApp           As Object
    Dim xlBook          As Object
    Dim varPrintSeets   As Variant
    Dim isException     As Boolean
    
    generatePdfEx2 = False
    
    On Error GoTo Exception
    
    Set xlApp = CreateObject("Excel.Application")    ' Excelオブジェクトを生成する
    Set xlBook = xlApp.Workbooks.Open(aSourceFile)   'ファイルオープン

    '配置図の抽出とＰＤＦ変換
    varPrintSeets = getSheetsName(xlBook, aKeyWord)
    If varPrintSeets(0) = "" Then GoTo Finally
    '対象シートをPDF変換する
    Call MSZZ038.PDFConvertEx(xlBook.Sheets(varPrintSeets), xlBook.NAME, _
                            aOutputPath & aYardCode & ".pdf")
    
    generatePdfEx2 = True
    GoTo Finally

Exception:
    isException = True
  
Finally:
    If Not xlBook Is Nothing Then Set xlBook = Nothing
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = False
        xlApp.Quit
        Set xlApp = Nothing
    End If
    If isException = True Then
        Call Err.Raise(Err.Number, "generatePdfEx" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :PDF変換対象シート名の取得
'        MODULE_ID        :getSheetsName
'        Parameter        :第1引数(Object) = Excelオブジェクト
'                         :第2引数(String) = 抽出キーワード
'        戻り値           :キーワードにマッチしたシート名配列
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getSheetsName(aXlBook As Object, aKeyWord As String) As Variant

    Dim workSheet As Object
    Dim sheetName As String
    
    On Error GoTo Exception
    
    getSheetsName = Array("")   '空の配列を設定しておく
    
    For Each workSheet In aXlBook.Worksheets
        If Nz(InStr(workSheet.NAME, aKeyWord), 0) <> 0 And workSheet.Visible = True Then
            sheetName = sheetName & workSheet.NAME & "/"
        End If
    Next
    If Len(sheetName) > 0 Then
        sheetName = Left(sheetName, Len(sheetName) - 1) '末尾の"/"除去
        '抽出したシート名を列挙したものから配列にしておく
        getSheetsName = Split(sheetName, "/")
    End If

    Exit Function

Exception:
        Call Err.Raise(Err.Number, "getSheetsName" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :Webページのアップロード
'        MODULE_ID        :UploadWebPage
'        Parameter        :第1引数(String)      = 部門コード
'                         :第2引数(String)      = ヤードコード
'        戻り値           : True...成功
'                         : False...失敗
'                         :
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UploadWebPage(aBumonCode As String, _
                              aYardCode As String _
                              ) As Boolean
                              
    Dim objFtp          As Object
    Dim sourcePath      As String
    Dim basePath        As String
    Dim uploadPath      As String
    Dim yardHtmFileName     As String
    Dim sizePdfFileName     As String
    Dim layoutPdfFileName   As String
    Dim imgAFileName        As String
    Dim imgBFileName        As String
    Dim imgUploadPath   As String
    
    Dim isException As Boolean
    
    UploadWebPage = False
    isException = False
    
    On Error GoTo Exception
  
    basePath = getINTIF_RECDB("BASE_PATH")       '基点フォルダ取得
    uploadPath = getINTIF_RECDB("UPLOAD_PATH")   'UP先の基点フォルダ取得
    sourcePath = getINTIF_RECDB("GENERATE_PATH") 'UP元の基点フォルダ取得
    
    '取得した情報を元に合成
    uploadPath = uploadPath & basePath
    sourcePath = sourcePath & Replace(basePath, "/", "\")

    'ファイル名決め
    If aBumonCode = P_トランク部門 Then
        'トランクのみヤードコードはゼロサプレス５桁
''        yardHtmFileName = sourcePath & Format$(CLng(aYardCode), "00000") & ".htm" 'DELETE 2011/10/14 M.R
        layoutPdfFileName = sourcePath & Replace(P_トランク配置図_FOLDER, "/", "\") & aYardCode & ".pdf"
        sizePdfFileName = sourcePath & Replace(P_トランク寸法表_FOLDER, "/", "\") & aYardCode & ".pdf"
''        imgAFileName = sourcePath & Replace(P_トランク画像_FOLDER, "/", "\") & aYardCode & "a.jpg"    'DELETE 2011/10/14 M.R
''        imgBFileName = sourcePath & Replace(P_トランク画像_FOLDER, "/", "\") & aYardCode & "b.jpg"    'DELETE 2011/10/14 M.R
''        imgUploadPath = P_トランク画像_FOLDER             'DELETE 2011/10/14 M.R
    Else
''        yardHtmFileName = sourcePath & aYardCode & ".htm" 'DELETE 2011/10/14 M.R
        layoutPdfFileName = sourcePath & Replace(P_コンテナ配置図_FOLDER, "/", "\") & aYardCode & ".pdf"
        sizePdfFileName = ""
''        imgAFileName = sourcePath & Replace(P_コンテナ画像_FOLDER, "/", "\") & aYardCode & "a.jpg"    'DELETE 2011/10/14 M.R
''        imgBFileName = sourcePath & Replace(P_コンテナ画像_FOLDER, "/", "\") & aYardCode & "b.jpg"    'DELETE 2011/10/14 M.R
''        imgUploadPath = P_コンテナ画像_FOLDER             'DELETE 2011/10/14 M.R
    End If

    'それぞれのファイルが元フォルダに存在するかチェックする
''    If Dir$(yardHtmFileName) = "" Or Dir$(layoutPdfFileName) = "" Then GoTo Finally   'DELETE 2011/10/14 M.R
    If sizePdfFileName <> "" And Dir$(sizePdfFileName) = "" Then GoTo Finally
''    If Dir$(imgAFileName) = "" Or Dir$(imgBFileName) = "" Then GoTo Finally           'DELETE 2011/10/14 M.R
    
    'アップロードファイルの確認OK
    ' FTP出力準備
    Set objFtp = Kase3535FtpCreate()    'kase3535用の接続
    
    If aBumonCode = P_トランク部門 Then
        '配置図ＰＤＦをネットへＵＰ
        Call MSZZ037.FtpPut(objFtp, sourcePath & Replace(P_トランク配置図_FOLDER, "/", "\") & aYardCode & ".pdf", _
                                    uploadPath & P_トランク配置図_FOLDER)
        'サイズ詳細をネットへＵＰ
        Call MSZZ037.FtpPut(objFtp, sourcePath & Replace(P_トランク寸法表_FOLDER, "/", "\") & aYardCode & ".pdf", _
                                    uploadPath & P_トランク寸法表_FOLDER)
    Else
        '配置図ＰＤＦをネットへＵＰ
        Call MSZZ037.FtpPut(objFtp, sourcePath & Replace(P_コンテナ配置図_FOLDER, "/", "\") & aYardCode & ".pdf", _
                                    uploadPath & P_コンテナ配置図_FOLDER)

    End If
    
''    '画像ファイルをネットへＵＰ
''    Call MSZZ037.FtpPut(objFtp, imgAFileName, uploadPath & imgUploadPath)     'DELETE 2011/10/14 M.R
''    Call MSZZ037.FtpPut(objFtp, imgBFileName, uploadPath & imgUploadPath)     'DELETE 2011/10/14 M.R
''
''    'ヤード詳細ページをネットへＵＰ
''    Call MSZZ037.FtpPut(objFtp, yardHtmFileName, uploadPath)      'DELETE 2011/10/14 M.R
    UploadWebPage = True
    
    GoTo Finally
   
    Exit Function
                              
Exception:
    isException = True
  
Finally:
    If Not objFtp Is Nothing Then objFtp.Close: Set objFtp = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "UploadWebPage" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :CSVファイルのアップロード
'        MODULE_ID        :UploadCSVFile
'        Parameter        :第1引数(String)      = 部門コード
'                         :第2引数(String)      = ヤードコード
'        戻り値           : True...成功
'                         : False...失敗
'                         :
'        CREATE_DATE      :2009/01/23
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UploadCSVFile(ByRef asFile() As String, ByVal asUpLoadPath As String) As Boolean
                              
    Dim objFtp              As Object
    Dim lngCnt              As Long
    Dim isException         As Boolean
    
    UploadCSVFile = False
    isException = False
    
    On Error GoTo Exception
    
    'アップロードファイルの確認OK
    ' FTP出力準備
    Set objFtp = Kase3535FtpCreate()                                'kase3535用の接続
    
    'CSVファイルをアップロード
    For lngCnt = 1 To UBound(asFile)
        Call MSZZ037.FtpPut(objFtp, asFile(lngCnt), asUpLoadPath)
        
    Next
    
    UploadCSVFile = True
    
    GoTo Finally
    
    Exit Function
    
Exception:
    isException = True
    
Finally:
    If Not objFtp Is Nothing Then objFtp.Close: Set objFtp = Nothing
    
    If isException = True Then
        Call Err.Raise(Err.Number, "UploadCSVFile" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
        
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :指定した６桁ヤードコードの画像ファイルを有無を確認し
'                         :無ければ５桁のヤードコード画像を６桁のファイル名で作っておく
'        MODULE_ID        :isSixDigitsJpg
'        Parameter        :第1引数(String)      = フルパス画像ファイル名
'                         :第2引数(String)      = ヤードコード
'        戻り値           : True...あった
'                         : False..どうにもこうにも見つからず
'                         :
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function isSixDigitsJpg(anImgFile As String, aYardCode) As Boolean

    Dim fiveName As String
    
    On Error GoTo Exception
    
    isSixDigitsJpg = False
    
    If Dir$(anImgFile) <> "" Then
        isSixDigitsJpg = True
        Exit Function
    End If

    ' ６桁の画像ファイル名がないので５桁であるか確認する
    fiveName = Left$(anImgFile, InStrRev(anImgFile, "\")) & Format$(CLng(aYardCode), "00000") & _
                Right$(anImgFile, 5) '5 = "a.jpg" の桁数
    If Dir$(fiveName) = "" Then Exit Function  '5桁でもなかった

    '５桁のファイルがあるならば、それを６桁のファイル名で保存
    FileCopy fiveName, anImgFile

    isSixDigitsJpg = True
    Exit Function

Exception:
   Call Err.Raise(Err.Number, "isSixDigitsJpg" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :アップロード済み情報記録
'        MODULE_ID        :SetUploadInfo
'        Parameter        :第1引数(ADOコネクション）
'                         :第2引数(ヤードコード）
'        戻り値           :設定した日付
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function SetUploadInfo(aConnection As Object, _
                               aYardCode As Long _
                            ) As String
    Dim sqlText As String
    Dim rsData  As Object
     
    SetUploadInfo = ""
    
    On Error GoTo Exception
  
    sqlText = "SELECT * "
    sqlText = sqlText & " FROM YARD_MAST "
    sqlText = sqlText & " WHERE YARD_CODE = '" & aYardCode & "' "
    
    Set rsData = MSZZ025.ADODB_Recordset(sqlText, aConnection, adoReadWrite)
    
    If Not rsData.EOF Then
        'データ更新
        rsData.Fields("YARD_WP_UPDATE") = DATE
        rsData.Fields("YARD_NETUSE_KBN") = -1
        rsData.UPDATE
    End If
    
   rsData.Close
   Set rsData = Nothing
   SetUploadInfo = Format$(DATE, "yyyy/mm/dd")
   Exit Function
                              
Exception:
  If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
  Call Err.Raise(Err.Number, "SetUploadInfo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :ヤード詳細ページのＵＲＬ（サーバの）
'        MODULE_ID        :GetYardPageUrl
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetYardPageUrl(aBumonCode As String, _
                                aYardCode As String _
                            ) As String
                    
    Dim basePath        As String
    Dim sourcePath      As String
    Dim url             As String
                    
    On Error GoTo Exception
    
    GetYardPageUrl = ""
    
    basePath = getINTIF_RECDB("BASE_PATH")          '基点フォルダ取得
    sourcePath = getINTIF_RECDB("GENERATE_PATH")    'UP元の基点フォルダ取得
    
    'ヤード名確定
    sourcePath = sourcePath & Replace(basePath, "/", "\")
    If aBumonCode = P_トランク部門 Then
         url = sourcePath & Format$(CLng(aYardCode), "00000") & ".htm"
    Else
         url = sourcePath & aYardCode & ".htm"
    End If
    If Dir$(url) = "" Then Exit Function
    
    GetYardPageUrl = url
   
    Exit Function
Exception:
    Call Err.Raise(Err.Number, "GetYardPageUrl" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function


'==============================================================================*
'
'        MODULE_NAME      :画像ファイルの準備
'        MODULE_ID        :preparationImageFile
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function PreparationImageFile(aBumonCode As String, _
                                 aYardCode As String _
                                )
    Dim noImgFileName   As String
    Dim outputPath      As String
    Dim basePath        As String
    Dim imgAFileName  As String
    Dim imgBFileName  As String
    
    PreparationImageFile = False
    
    On Error GoTo Exception
    
    noImgFileName = getINTIF_RECDB("NO_IMAGE_FILE")     'NoImageFile位置取得
    outputPath = getINTIF_RECDB("GENERATE_PATH")        'ヤードページ生成先ファイルパス
    basePath = getINTIF_RECDB("BASE_PATH")              '基点パス
    basePath = Replace(basePath, "/", "\")              'WEB用なんで書換
    
    '取得した情報を元に合成
    noImgFileName = outputPath & noImgFileName
    outputPath = outputPath & basePath
    
    ' ファイル名決め画像データはなければno_imageをその画像名にコピーしてくる。トランクは５桁画像まで探す
    If aBumonCode = P_トランク部門 Then
        imgAFileName = outputPath & Replace(P_トランク画像_FOLDER, "/", "\") & aYardCode & "a.jpg"
        imgBFileName = outputPath & Replace(P_トランク画像_FOLDER, "/", "\") & aYardCode & "b.jpg"
        If isSixDigitsJpg(imgAFileName, aYardCode) = False Then FileCopy noImgFileName, imgAFileName
        If isSixDigitsJpg(imgBFileName, aYardCode) = False Then FileCopy noImgFileName, imgBFileName
    Else
        imgAFileName = outputPath & Replace(P_コンテナ画像_FOLDER, "/", "\") & aYardCode & "a.jpg"
        imgBFileName = outputPath & Replace(P_コンテナ画像_FOLDER, "/", "\") & aYardCode & "b.jpg"
        If Dir$(imgAFileName) = "" Then FileCopy noImgFileName, imgAFileName
        If Dir$(imgBFileName) = "" Then FileCopy noImgFileName, imgBFileName
    End If
    
    PreparationImageFile = True
    
    Exit Function

Exception:
    Call Err.Raise(Err.Number, "PreparationImageFile" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :INTI_FILEの値取得
'        MODULE_ID        :getINTIF_RECDB
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getINTIF_RECDB(aRecfb As String) As String

    Dim recdb As String

    recdb = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & INTIF_PROGB & "' AND INTIF_RECFB = '" & aRecfb & "'"))
    If recdb = "" Then
        Call MSZZ024_M10("getINTIF_RECDB", "INTI_FILE[" & aRecfb & "]の設定不足です。")
    End If
    getINTIF_RECDB = recdb

End Function

'==============================================================================*
'
'       MODULE_NAME     : ソケット作成
'       MODULE_ID       : Kase35335FtpCreate
'       CREATE_DATE     : 2008/02/29
'       PARAM           : [bPASV]               パッシブモード(I)省略時(False)
'       RETURN          : FTPオブジェクト(Object)
'
'
'==============================================================================*
'Private Const C_FTP_PROVIDER_2 = "basp21.FTP"
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Kase3535FtpCreate(Optional bPAVS As Boolean = False) As Object
    Dim objFtp              As Object
    Dim strSvr              As String
    Dim strUid              As String
    Dim strPwd              As String
    On Error GoTo ErrorHandler
    
    Set objFtp = CreateObject("basp21.FTP")
    If objFtp Is Nothing Then
        Call MSZZ024_M10("FTP.PROVIDER", "システムの設定不足です。")
    End If
    
    strSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'FTP_KASE3535_SERVER_NAME'"))
    strUid = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'FTP_KASE3535_USER_ID'"))
    strPwd = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'FTP_KASE3535_PASSWORD'"))
    If (strSvr <> "") And (strUid <> "") And (strPwd <> "") Then
        MSZZ037.FtpConnect objFtp, strSvr, strUid, strPwd, bPAVS
    Else
        Call MSZZ024_M10("getINTIF_RECDB", "INTI_FILE[FTP_KASE3535系]の設定不足です。")
    End If
    
    Set Kase3535FtpCreate = objFtp
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Kase3535FtpCreate" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :Webページの削除
'        MODULE_ID        :DeleteWebPage
'        Parameter        :第1引数(String)      = 部門コード
'                         :第2引数(String)      = ヤードコード
'        戻り値           : True...成功
'                         : False...失敗
'                         :
'        CREATE_DATE      :2009/01/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function DeleteWebPage(aBumonCode As String, _
                              aYardCode As String _
                              ) As Boolean
                              
    Dim objFtp                  As Object
    
    Dim basePath                As String
    Dim uploadPath              As String
    
    Dim yardHtmFileName         As String
    Dim sizePdfFileName         As String
    Dim layoutPdfFileName       As String
    Dim imgAFileName            As String
    Dim imgBFileName            As String
    
    Dim isException As Boolean
    
    DeleteWebPage = False
    isException = False
    
    On Error GoTo Exception
  
    basePath = getINTIF_RECDB("BASE_PATH")       '基点フォルダ取得
    uploadPath = getINTIF_RECDB("UPLOAD_PATH")   'UP先の基点フォルダ取得
    
    '取得した情報を元に合成
    uploadPath = uploadPath & basePath

    'ファイル名決め
    If aBumonCode = P_トランク部門 Then
        'トランクのみヤードコードはゼロサプレス５桁
        yardHtmFileName = uploadPath & Format$(CLng(aYardCode), "00000") & ".htm"
        layoutPdfFileName = uploadPath & P_トランク配置図_FOLDER & aYardCode & ".pdf"
        sizePdfFileName = uploadPath & P_トランク寸法表_FOLDER & aYardCode & ".pdf"
        imgAFileName = uploadPath & P_トランク画像_FOLDER & aYardCode & "a.jpg"
        imgBFileName = uploadPath & P_トランク画像_FOLDER & aYardCode & "b.jpg"
    Else
        yardHtmFileName = uploadPath & aYardCode & ".htm"
        layoutPdfFileName = uploadPath & P_コンテナ配置図_FOLDER & aYardCode & ".pdf"
        sizePdfFileName = ""
        imgAFileName = uploadPath & P_コンテナ画像_FOLDER & aYardCode & "a.jpg"
        imgBFileName = uploadPath & P_コンテナ画像_FOLDER & aYardCode & "b.jpg"
    End If
    
    ' FTP削除準備
    Set objFtp = Kase3535FtpCreate()    'kase3535用の接続
    
    '配置図ＰＤＦを削除
    Call MSZZ037.FtpDelete(objFtp, layoutPdfFileName)

    'サイズ詳細を削除
    If aBumonCode = P_トランク部門 Then
        Call MSZZ037.FtpDelete(objFtp, sizePdfFileName)

    End If
    
    '画像ファイルを削除
    Call MSZZ037.FtpDelete(objFtp, imgAFileName)
    Call MSZZ037.FtpDelete(objFtp, imgBFileName)
    
    'ヤード詳細ページを削除
    Call MSZZ037.FtpDelete(objFtp, yardHtmFileName)
    DeleteWebPage = True
    
    GoTo Finally
   
    Exit Function
                              
Exception:
    isException = True
  
Finally:
    If Not objFtp Is Nothing Then objFtp.Close: Set objFtp = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "DeleteWebPage" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :パンくず設定データ取得
'        MODULE_ID        :GenYardPage
'        Parameter        :
'        戻り値           :Nothing
'        CREATE_DATE      :2008/02/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetBreadCrumb(ByRef asAddress As String, ByRef asURL As String, ByRef asSTATES_NAME As String) As Boolean

    Dim strKey  As String           'キー
    
    GetBreadCrumb = False
    
    On Error GoTo ErrorHandler
    
    strKey = "URL_"
    
    If 0 < InStr(asAddress, P_STATES_TK) Then
        '東京
        strKey = strKey & "TOKYO"
        
    ElseIf 0 < InStr(asAddress, P_STATES_ST) Then
        '埼玉
        strKey = strKey & "SAITAMA"
    
    ElseIf 0 < InStr(asAddress, P_STATES_CB) Then
        '千葉
        strKey = strKey & "CHIBA"
    
    ElseIf 0 < InStr(asAddress, P_STATES_KG) Then
        '神奈川
        strKey = strKey & "KANAGAWA"
        
    ElseIf 0 < InStr(asAddress, P_STATES_SZ) Then
        '静岡
        strKey = strKey & "SHIZUOKA"
        
    Else
        '例外
        Exit Function
        
    End If
    
    asURL = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = '" & strKey & "'"))
    asSTATES_NAME = Nz(DLookup("SETUT_BIKON", "SETU_TABL", "SETUT_SETUB = '" & strKey & "'"))
    
    GetBreadCrumb = True
    
    Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "GetBreadCrumb" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

