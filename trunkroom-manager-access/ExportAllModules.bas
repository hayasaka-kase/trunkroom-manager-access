Attribute VB_Name = "ExportAllModules"
Option Compare Database

Public Sub ExportAllModules()
    Dim vbeComp As Object
    Dim exportPath As String
    Dim fileName As String
    Dim extension As String
    Dim fso As Object
    
    ' 1. 出力先フォルダの設定（デスクトップに「Trunkroom-Manager-Access」フォルダを作成）
    Set fso = CreateObject("Scripting.FileSystemObject")
    exportPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Trunkroom-Manager-Access\"
    
    If Not fso.FolderExists(exportPath) Then
        fso.CreateFolder exportPath
    End If

    ' 2. コンポーネントをループしてエクスポート
    For Each vbeComp In Application.VBE.ActiveVBProject.VBComponents
        Select Case vbeComp.Type
            Case 1 ' 標準モジュール
                extension = ".bas"
            Case 2 ' クラスモジュール (MySqlFieldなど)
                extension = ".cls"
            Case 3 ' フォーム
                extension = ".frm"
            Case 100 ' Document (Table/Report/ThisWorkbookなど)
                extension = ".cls"
            Case Else
                extension = ".txt"
        End Select
        
        fileName = exportPath & vbeComp.NAME & extension
        
        ' エクスポート実行
        On Error Resume Next
        vbeComp.Export fileName
        If Err.Number <> 0 Then
            Debug.Print "失敗: " & vbeComp.NAME
        Else
            Debug.Print "成功: " & vbeComp.NAME
        End If
        On Error GoTo 0
    Next vbeComp

    MsgBox "エクスポートが完了しました。" & vbCrLf & "保存先: " & exportPath, vbInformation
End Sub
