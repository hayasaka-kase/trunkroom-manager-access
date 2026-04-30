Attribute VB_Name = "ExportAllModules"
Option Compare Database
Option Explicit

Public Sub ExportAllModules()
    Dim vbeComp As Object
    Dim exportPath As String
    Dim fileName As String
    Dim tempFile As String
    Dim extension As String
    Dim fso As Object
    
    ' 1. 出力先フォルダの設定
    Set fso = CreateObject("Scripting.FileSystemObject")
    exportPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERPROFILE%") & "\git\trunkroom-manager-access\modules\"
    
    If Not fso.FolderExists(exportPath) Then
        fso.CreateFolder exportPath
    End If

    ' 2. 各モジュールをエクスポートしてUTF-8に変換
    For Each vbeComp In Application.VBE.ActiveVBProject.VBComponents
        Select Case vbeComp.Type
            Case 1: extension = ".bas" ' 標準モジュール
            Case 2: extension = ".cls" ' クラス
            Case 3: extension = ".frm" ' フォーム
            Case 100: extension = ".cls" ' Document
            Case Else: extension = ".txt"
        End Select
        
        fileName = exportPath & vbeComp.NAME & extension
        tempFile = fileName & ".tmp" ' 一時ファイル
        
        On Error Resume Next
        ' 一旦、VBA標準のShift-JISで書き出す
        vbeComp.Export tempFile
        
        If Err.Number <> 0 Then
            Debug.Print "失敗: " & vbeComp.NAME
        Else
            ' Shift-JISのテンポラリファイルを読み込み、UTF-8(BOMなし)で保存し直す
            If ConvertToUTF8NoBOM(tempFile, fileName) Then
                fso.DeleteFile tempFile ' テンポラリを削除
                Debug.Print "成功(UTF-8): " & vbeComp.NAME
            End If
        End If
        On Error GoTo 0
    Next vbeComp

    MsgBox "GitHub対応(UTF-8)でのエクスポートが完了しました。", vbInformation
End Sub

' --- Shift-JISを読み込み、BOMなしUTF-8で保存する関数 ---
Private Function ConvertToUTF8NoBOM(ByVal src As String, ByVal dest As String) As Boolean
    On Error GoTo ErrHandler
    Dim streamIn As Object, streamOut As Object
    
    ' 入力ストリーム (Shift-JIS)
    Set streamIn = CreateObject("ADODB.Stream")
    streamIn.Type = 2 ' adTypeText
    streamIn.Charset = "Shift_JIS"
    streamIn.Open
    streamIn.LoadFromFile src
    
    ' 出力ストリーム (UTF-8)
    Set streamOut = CreateObject("ADODB.Stream")
    streamOut.Type = 2 ' adTypeText
    streamOut.Charset = "UTF-8"
    streamOut.Open
    streamOut.WriteText streamIn.ReadText
    
    ' BOM (先頭3バイト) を除去するためのバイナリ変換
    streamOut.Position = 0
    streamOut.Type = 1 ' adTypeBinary
    streamOut.Position = 3 ' BOMをスキップ
    
    Dim binaryData As Variant
    binaryData = streamOut.Read
    
    streamOut.Close
    streamOut.Open
    streamOut.Write binaryData
    streamOut.SaveToFile dest, 2 ' adSaveCreateOverWrite
    
    streamIn.Close
    streamOut.Close
    ConvertToUTF8NoBOM = True
    Exit Function

ErrHandler:
    Debug.Print "変換エラー: " & Err.Description
    ConvertToUTF8NoBOM = False
End Function
