Attribute VB_Name = "MSZZ074"
'****************************  start of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : AES復号化
'       PROGRAM_ID      : MSZZ074
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2017/01/26
'       CERATER         : K.ISHIZAKA
'       Ver             : 0.0
'
'       UPDATE          : 2019/03/14
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.1
'                       : 暗号化できるようにする
'                         ※新しいキーの作成はSQLServerに接続します
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   API宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Declare Function CryptAcquireContextA Lib "advapi32.dll" ( _
    ByRef phProv As Long, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByRef pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal hPubKey As Long, _
    ByVal dwFlags As Long, _
    ByRef phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" ( _
    ByVal hKey As Long) As Long
Private Declare Function CryptSetKeyParam Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal dwParam As Long, _
    ByRef pbData As Any, _
    ByVal dwFlags As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    ByRef pbData As Any, _
    ByRef pdwDataLen As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    ByRef pbData As Any, _
    ByRef pdwDataLen As Long, _
    ByVal dwBufLen As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const CRYPT_VERIFYCONTEXT       As Long = &HF0000000
Private Const ALG_TYPE_BLOCK            As Long = 1536
Private Const ALG_CLASS_DATA_ENCRYPT    As Long = 24576

Private Const ALG_SID_AES_128   As Long = 14
Private Const ALG_SID_AES_192   As Long = 15
Private Const ALG_SID_AES_256   As Long = 16

Private Const PROV_RSA_AES      As Long = 24

Private Const CALG_AES_128      As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_128
Private Const CALG_AES_192      As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_192
Private Const CALG_AES_256      As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_256

Private Const KP_IV             As Long = 1
Private Const KP_PADDING        As Long = 3
Private Const KP_MODE           As Long = 4

Private Const PKCS5_PADDING     As Long = 1
Private Const CRYPT_MODE_CBC    As Long = 1
Private Const PLAINTEXTKEYBLOB  As Long = 8
Private Const CUR_BLOB_VERSION  As Long = 2

'BLOBHEADERユーザ定義型
Private Type BLOBHEADER
    bType       As Byte
    bVersion    As Byte
    reserved    As Integer
    aiKeyAlg    As Long
End Type

'インポート用の鍵データのユーザ定義型
'
'PUBLICKEYSTRUC BLOBヘッダに続いて、鍵サイズ、鍵データが必要だが、
'鍵データについては鍵サイズによって配列サイズが変わるため、
'ロジック中で動的にメモリを確保するようにし、ここでは未定義とする
Private Type keyBlob
    hdr         As BLOBHEADER
    keySize     As Long
'    keyData()   As Byte
End Type

'鍵長定数定義
Public Enum AESKeyBits
    AES_KEY128 = 128
    AES_KEY192 = 192
    AES_KEY256 = 256
End Enum

'エラーコード定義
Private Const ERR_CRYPT_API     As Long = vbObjectError + 513   'CryptAPIエラー
Private Const ERR_KEY_LENGTH    As Long = vbObjectError + 514   '鍵長エラー
Private Const ERR_IV_LENGTH     As Long = vbObjectError + 515   'IV長エラー

'==============================================================================*
'   テスト
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub testaa()
    Dim iv                  As String
    Dim key                 As String
    Dim data                As String
    
    'config.ini に設定されいる INITIAL_VECTOR の値
    iv = "954765e246cea092578dce72f6cc68b2b283de15fc80694cce4ccb604821d70c4f412545811672d9d8da60b86e93359f00e0daa13c687111b4984f7f07a0b9b60187ee46ae0c897dd9407aad54ee5b200908e6fc6a892f4656dec9a3d77c7eaa11158c5c1d5bcf92705fccd4fab2e8d8cf4826c290d5d4944ed82b136c56"
    'LGIN_TABL の LGINT_TOKEN の値
    key = "2712f44103a02048b90a49607eba2b2ef7b3329504250046d04aa50b6ed02f5e503f22638afcd0a40af7b04b6e0a2810d426d9874077f021602c0c8aa04d7e0ab4f0fe7aac6bbfc65c600c150e6ce044190bdcc056d173e80ac2e5a2731e0d25504e3c0870e0a5ca462341d9312de34e0064b04c9f084800bfc07ae4b783"
    'USER_MAST の 暗号化された値
    data = ""
    
    'USER_MAST の戻したい値
    Debug.Print AesDecrypt(key, iv, data)
    Stop
End Sub

Private Sub testbb()
    Dim iv                  As String
    Dim key                 As String
    Dim data                As String
    'config.ini に設定されいる INITIAL_VECTOR の値
    iv = "73292784dd93a980ff2ad18c4fe7ef87ba728d7e54e2403d07cb0208d5d977b643e04da3a0f4fa2a25698ff3fdd1dac16a2d464e562639f9be94c2e8a6781de38fcc39e80802cf17fe2586c9c89ecd80b41883b99f1ca85df3831b098b1b186b11e85728718c15d6b21c32847b08a3bffe41cd441b52d3a3ef8f6d547f29"
    'キーを作成
    key = AesNewKey()
    Debug.Print key
    '平文
    data = "いしざか㈱あおいう"
    '暗号化
    data = AesEncrypt(key, iv, data)
    Debug.Print data
    '復号化
    data = AesDecrypt(key, iv, data)
    Debug.Print data
    Stop
End Sub

'==============================================================================*
'
'       MODULE_NAME     : AES復号化
'       MODULE_ID       : AesDecrypt
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : key                   キー(I)
'                       : iv                    初期ベクタ(I)
'                       : data                  暗号文字列
'       RETURN          : 平文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function AesDecrypt(ByVal key As String, ByVal iv As String, ByVal data As String) As String
    On Error GoTo ErrorHandler
    
    AesDecrypt = UTF8_GetString(MyAesDecrypt(hex2bin(key, 32), hex2bin(iv, 16), base64_decode(data), AES_KEY256))
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "AesDecrypt" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : AES暗号化
'       MODULE_ID       : AesEncrypt
'       CREATE_DATE     : 2019/03/14            K.ISHIZAKA
'       PARAM           : key                   キー(I)
'                       : iv                    初期ベクタ(I)
'                       : data                  平文(I)
'       RETURN          : 暗号文字列(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function AesEncrypt(ByVal key As String, ByVal iv As String, ByVal data As String) As String
    On Error GoTo ErrorHandler
    
    AesEncrypt = base64_encode(MyAesEncrypt(hex2bin(key, 32), hex2bin(iv, 16), UTF8_GetBytes(data), AES_KEY256))
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "AesEncrypt" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ヘックス文字列をバイナリ配列に変換する
'       MODULE_ID       : hex2bin
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : strData               ヘックス文字列(I)
'                       : iLen                  長さ(I)
'       RETURN          : バイナリ配列(Byte())
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function hex2bin(ByVal strData As String, ByVal ilen As Integer) As Byte()
    Dim i                   As Integer
    Dim byteResult()        As Byte
    On Error GoTo ErrorHandler
    
    ReDim byteResult(0 To ilen - 1)
    For i = 0 To ilen - 1
        byteResult(i) = CByte("&H" & Mid(strData, i * 2 + 1, 2))
    Next
    hex2bin = byteResult
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "hex2bin" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : BASE64文字列をバイナリ配列にデコードする
'       MODULE_ID       : base64_decode
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : strData               BASE64文字列(I)
'       RETURN          : バイナリ配列(Byte())
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function base64_decode(ByVal strData As String) As Byte()
    Dim objBase64           As Object
    On Error GoTo ErrorHandler

    Set objBase64 = CreateObject("MSXML2.DOMDocument").createElement("b64")
    objBase64.DataType = "bin.base64"
    objBase64.Text = strData
    base64_decode = objBase64.nodeTypedValue

    Set objBase64 = Nothing
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "base64_decode" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : バイナリ配列をBASE64文字列にエンコードする
'       MODULE_ID       : base64_encode
'       CREATE_DATE     : 2019/03/14            K.ISHIZAKA
'       PARAM           : byteData()            バイナリ配列(I)
'       RETURN          : BASE64文字列(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function base64_encode(ByRef byteData() As Byte) As String
    Dim objBase64 As Object
    On Error GoTo ErrorHandler

    Set objBase64 = CreateObject("MSXML2.DOMDocument").createElement("b64")
    objBase64.DataType = "bin.base64"
    objBase64.nodeTypedValue = byteData
    base64_encode = objBase64.Text

    Set objBase64 = Nothing
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "base64_encode" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : AES復号化
'       MODULE_ID       : MyAesDecrypt
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : key()                 キー(I)
'                       : iv()                  初期ベクタ(I)
'                       : data()                暗号文字列
'                       : keyBits               鍵長(I)
'       RETURN          : 復号バイナリ配列(Byte())
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MyAesDecrypt(ByRef key() As Byte, ByRef iv() As Byte, ByRef data() As Byte, ByVal keyBits As AESKeyBits) As Byte()
    Dim hProv               As Long     'CSPハンドラ
    Dim hKey                As Long     '暗号鍵ハンドラ
    Dim strResult           As String
    On Error GoTo ErrorHandler

    Dim keyLength As Long   '鍵バイト長
    keyLength = keyBits / 8 'ビット->バイト変換

    '鍵長のチェック
    If UBound(key) + 1 <> keyLength Then
        Err.Raise ERR_KEY_LENGTH, "decrypt()", "鍵長が不正です: " & UBound(key) + 1 & "byte"
    End If

    'IV長のチェック
    If UBound(iv) + 1 <> 16 Then
        Err.Raise ERR_IV_LENGTH, "decrypt()", "IV長が不正です: " & UBound(iv) + 1 & "byte"
    End If

    'CSP(Cryptographic Service Provider)のハンドルを取得
    If Not CBool(CryptAcquireContextA(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)) Then
        Err.Raise Err.LastDllError, "decrypt()->CryptAcquireContext()", GetAPIErrorText(Err.LastDllError)
    End If
    On Error GoTo ErrorHandler1

    Dim blob As keyBlob '鍵データ(ユーザ定義型)
    Dim keyData() As Byte   '鍵データ(バイト列)

    '鍵データの作成
    'keyBlobユーザ定義型に鍵データを結合したバイト列を無理やり作成する
    blob.hdr.bType = PLAINTEXTKEYBLOB
    blob.hdr.bVersion = CUR_BLOB_VERSION
    blob.hdr.reserved = 0
    blob.hdr.aiKeyAlg = CALG_AES_256
    blob.keySize = keyLength
    ReDim keyData(LenB(blob) + blob.keySize - 1)
    Call RtlMoveMemory(keyData(0), blob, LenB(blob))
    Call RtlMoveMemory(keyData(LenB(blob)), key(0), keyLength)

    '鍵のインポート
    If Not CBool(CryptImportKey(hProv, keyData(0), UBound(keyData) + 1, 0, 0, hKey)) Then
        Err.Raise Err.LastDllError, "CryptImportKey()", GetAPIErrorText(Err.LastDllError)
    End If
    On Error GoTo ErrorHandler2

    'パディング方式の設定(PKCS#5)
    If Not CBool(CryptSetKeyParam(hKey, KP_PADDING, PKCS5_PADDING, 0)) Then
        Err.Raise Err.LastDllError, "CryptSetKeyParam():KP_PADDING", GetAPIErrorText(Err.LastDllError)
    End If

    'IV(Initialization Vector)の設定
    If Not CBool(CryptSetKeyParam(hKey, KP_IV, iv(0), 0)) Then
        Err.Raise Err.LastDllError, "CryptSetKeyParam():KP_IV", GetAPIErrorText(Err.LastDllError)
    End If

    '暗号モードの設定(暗号文ブロック連鎖モード)
    If Not CBool(CryptSetKeyParam(hKey, KP_MODE, CRYPT_MODE_CBC, 0)) Then
        Err.Raise Err.LastDllError, "CryptSetKeyParam():KP_MODE", GetAPIErrorText(Err.LastDllError)
    End If

    '暗号化バイト列長
    Dim dwDataLen As Long
    dwDataLen = UBound(data) + 1

    'CryptDecryptは、引数の暗号化バイト列に復号したバイト列を戻す仕様のため
    'メソッドの引数の暗号化バイト列をローカル変数にコピーして使用する
    Dim pbData() As Byte
    ReDim pbData(dwDataLen - 1)
    Call RtlMoveMemory(pbData(0), data(0), UBound(data) + 1)

    '復号処理
    If Not CBool(CryptDecrypt(hKey, 0, True, 0, pbData(0), dwDataLen)) Then
        Err.Raise Err.LastDllError, "CryptDecrypt()", GetAPIErrorText(Err.LastDllError)
    End If

    '暗号鍵ハンドラの開放
    If Not CBool(CryptDestroyKey(hKey)) Then
        Err.Raise Err.LastDllError, "CryptDestroyKey()", GetAPIErrorText(Err.LastDllError)
    End If
    On Error GoTo ErrorHandler1

    'CSPハンドラの開放
    If Not CBool(CryptReleaseContext(hProv, 0)) Then
        Err.Raise Err.LastDllError, "CryptReleaseContext()", GetAPIErrorText(Err.LastDllError)
    End If
    On Error GoTo ErrorHandler
    
    If dwDataLen > 0 Then
        ReDim Preserve pbData(dwDataLen - 1)
    End If
    MyAesDecrypt = pbData
Exit Function

ErrorHandler2:
    '暗号鍵ハンドラの開放
    Call CryptDestroyKey(hKey)
ErrorHandler1:
    'CSPハンドラの開放
    Call CryptReleaseContext(hProv, 0)
ErrorHandler:
    Call Err.Raise(Err.Number, "MyAesDecrypt" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : AES暗号化
'       MODULE_ID       : MyAesEncrypt
'       CREATE_DATE     : 2019/03/14            K.ISHIZAKA
'       PARAM           : key()                 キー(I)
'                       : iv()                  初期ベクタ(I)
'                       : data()                復号文字列
'                       : keyBits               鍵長(I)
'       RETURN          : 暗号バイナリ配列(Byte())
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MyAesEncrypt(ByRef key() As Byte, ByRef iv() As Byte, ByRef data() As Byte, ByVal keyBits As AESKeyBits) As Byte()
    Dim hProv               As Long     'CSPハンドラ
    Dim hKey                As Long     '暗号鍵ハンドラ
    Dim strResult           As String
    Const C_BUF_LEN         As Long = 1024
    On Error GoTo ErrorHandler

    Dim keyLength As Long   '鍵バイト長
    keyLength = keyBits / 8 'ビット->バイト変換

    '鍵長のチェック
    If UBound(key) + 1 <> keyLength Then
        Err.Raise ERR_KEY_LENGTH, "decrypt()", "鍵長が不正です: " & UBound(key) + 1 & "byte"
    End If

    'IV長のチェック
    If UBound(iv) + 1 <> 16 Then
        Err.Raise ERR_IV_LENGTH, "decrypt()", "IV長が不正です: " & UBound(iv) + 1 & "byte"
    End If

    'CSP(Cryptographic Service Provider)のハンドルを取得
    If Not CBool(CryptAcquireContextA(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)) Then
        Err.Raise Err.LastDllError, "decrypt()->CryptAcquireContext()", GetAPIErrorText(Err.LastDllError)
    End If
    On Error GoTo ErrorHandler1

    Dim blob As keyBlob '鍵データ(ユーザ定義型)
    Dim keyData() As Byte   '鍵データ(バイト列)

    '鍵データの作成
    'keyBlobユーザ定義型に鍵データを結合したバイト列を無理やり作成する
    blob.hdr.bType = PLAINTEXTKEYBLOB
    blob.hdr.bVersion = CUR_BLOB_VERSION
    blob.hdr.reserved = 0
    blob.hdr.aiKeyAlg = CALG_AES_256
    blob.keySize = keyLength
    ReDim keyData(LenB(blob) + blob.keySize - 1)
    Call RtlMoveMemory(keyData(0), blob, LenB(blob))
    Call RtlMoveMemory(keyData(LenB(blob)), key(0), keyLength)

    '鍵のインポート
    If Not CBool(CryptImportKey(hProv, keyData(0), UBound(keyData) + 1, 0, 0, hKey)) Then
        Err.Raise Err.LastDllError, "CryptImportKey()", GetAPIErrorText(Err.LastDllError)
    End If
    On Error GoTo ErrorHandler2

    'パディング方式の設定(PKCS#5)
    If Not CBool(CryptSetKeyParam(hKey, KP_PADDING, PKCS5_PADDING, 0)) Then
        Err.Raise Err.LastDllError, "CryptSetKeyParam():KP_PADDING", GetAPIErrorText(Err.LastDllError)
    End If

    'IV(Initialization Vector)の設定
    If Not CBool(CryptSetKeyParam(hKey, KP_IV, iv(0), 0)) Then
        Err.Raise Err.LastDllError, "CryptSetKeyParam():KP_IV", GetAPIErrorText(Err.LastDllError)
    End If

    '暗号モードの設定(暗号文ブロック連鎖モード)
    If Not CBool(CryptSetKeyParam(hKey, KP_MODE, CRYPT_MODE_CBC, 0)) Then
        Err.Raise Err.LastDllError, "CryptSetKeyParam():KP_MODE", GetAPIErrorText(Err.LastDllError)
    End If

    '暗号化バイト列長
    Dim dwDataLen As Long
    dwDataLen = UBound(data) + 1

    'CryptDecryptは、引数の暗号化バイト列に復号したバイト列を戻す仕様のため
    'メソッドの引数の暗号化バイト列をローカル変数にコピーして使用する
    Dim pbData() As Byte
    ReDim pbData(C_BUF_LEN)
    Call RtlMoveMemory(pbData(0), data(0), UBound(data) + 1)

    '暗号処理
    If Not CBool(CryptEncrypt(hKey, 0, True, 0, pbData(0), dwDataLen, C_BUF_LEN)) Then
        Err.Raise Err.LastDllError, "CryptDecrypt()", GetAPIErrorText(Err.LastDllError)
    End If

    '暗号鍵ハンドラの開放
    If Not CBool(CryptDestroyKey(hKey)) Then
        Err.Raise Err.LastDllError, "CryptDestroyKey()", GetAPIErrorText(Err.LastDllError)
    End If
    On Error GoTo ErrorHandler1

    'CSPハンドラの開放
    If Not CBool(CryptReleaseContext(hProv, 0)) Then
        Err.Raise Err.LastDllError, "CryptReleaseContext()", GetAPIErrorText(Err.LastDllError)
    End If
    On Error GoTo ErrorHandler
    
    If dwDataLen > 0 Then
        ReDim Preserve pbData(dwDataLen - 1)
    End If
    MyAesEncrypt = pbData
Exit Function

ErrorHandler2:
    '暗号鍵ハンドラの開放
    Call CryptDestroyKey(hKey)
ErrorHandler1:
    'CSPハンドラの開放
    Call CryptReleaseContext(hProv, 0)
ErrorHandler:
    Call Err.Raise(Err.Number, "MyAesEncrypt" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : AESキー作成
'       MODULE_ID       : AesNewKey
'       CREATE_DATE     : 2019/03/14            K.ISHIZAKA
'       RETURN          : キー(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function AesNewKey() As String
    Dim strSQL              As String
    Dim objCon              As Object
    Dim i                   As Long
    On Error GoTo ErrorHandler

    strSQL = "SELECT LOWER(REPLACE(''"
    For i = 1 To 8 - 1
        strSQL = strSQL & " + CONVERT(varchar(36), NEWID())"
    Next
    strSQL = strSQL & ", '-', '0'))"

    Set objCon = ADODB_Connection("H") 'Y
    On Error GoTo ErrorHandler1
    AesNewKey = Nz(ADODB_ExecGetVariant(strSQL, objCon))
    objCon.Close
    On Error GoTo ErrorHandler
Exit Function

ErrorHandler1:
    objCon.Close
ErrorHandler:
    Call Err.Raise(Err.Number, "AesNewKey" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  ended of program ********************************




