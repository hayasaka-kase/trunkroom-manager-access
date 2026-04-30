Attribute VB_Name = "modCvsSub2Define"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : コンビニ代金回収自動入金データ登録
'
'        PROGRAM_NAME    : modCvsSubDefine
'        PROGRAM_ID      :
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2001/07/11
'        CERATER         : S.SHIBAZAKI
'
'        UPDATE          : 2003/08/15
'        UPDATER         : N.MIURA
'                        : 1,部門コード・払込データ
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'実行ファイルパス
'Public Const EXEC_PATH = "C:\Documents and Settings\KPC0007\デスクトップ\柴崎\"
'Public Const EXEC_PATH = "\\Server03\電算室\TEST_SYSTEM\KOMS3\"
'Public Const EXEC_PATH = "D:\TEST_SYSTEM\ｺﾝﾃﾅ管理ｼｽﾃﾑ\"                        'DELETE 20030815 N.MIURA
'実行ログファイル名
'Public Const LOG_NAME = "CVSSUB2.LOG"                                          'DELETE 20030815 N.MIURA
'出力ファイルパス
'Public Const OUTPUT_PATH = "C:\Documents and Settings\KPC0007\デスクトップ\柴崎\output\"
'Public Const OUTPUT_PATH = "C:\kase_SYSTEM\KOMS\output\"
'Public Const OUTPUT_PATH = "D:\TEST_SYSTEM\ｺﾝﾃﾅ管理ｼｽﾃﾑ\output\"
'Public Const OUTPUT_PATH = "\\Server05\nts\df\HARA\"
'Public Const OUTPUT_PATH = "\\Server08\TEST_nts\df\HARA\"                      'DELETE 20030815 N.MIURA
'出力ファイル名
'Public Const OUTPUT_NAME = "CVS_OUT_H.txt"                                     'DELETE 20030815 N.MIURA

'バックアップパス
Public Const BACKUP_PATH = "\BACKUP"

'明細行数
Public Const MEISAI_MAX = 11

'支払期限 True = 入金予定日   False = 無期限
Public Const SIHARAI_KIGEN = True
Public Const KIGEN_VALUE = "99999"

'メッセージポックスタイトル
Public Const MSGBOX_TITLE = "コンビニ請求書発行"

'部門マスタ検索コード
'Public Const BUMO_CODE = "H"                                                   'DELETE 20030815 N.MIURA
'カンパニーマスタ検索コード
'Public Const COMP_CODE = "KAS"                                                 'DELETE 20030815 N.MIURA
'集金区分検索
Public Const SYUUKIN_KUBUN = "7"

'請求書消費税明細
Public Const SYOUHIZEI_MEISAI = "消費税"
'請求書合計明細
Public Const GOUKEI_MEISAI = "合計"

'実行ログ
'Public Const LOG_START = "［開始］"                                            'DELETE 20030815 N.MIURA
'Public Const LOG_END = "［終了］"                                              'DELETE 20030815 N.MIURA
'****************************  ended or program ********************************

