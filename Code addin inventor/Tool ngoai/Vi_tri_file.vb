Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Inventor
Namespace ToolInventor2025.Toolngoai
    Public Module Vitrifile
        Public Sub Vitrifile()
            Dim invApp As Inventor.Application = g_inventorApplication
            If invApp Is Nothing OrElse invApp.ActiveDocument Is Nothing Then
                MessageBox.Show("No document is open.")
                Exit Sub
            End If

            Dim doc As Inventor.Document = invApp.ActiveDocument

            ' 获取完整路径（对所有文档类型都适用）
            Dim fullPath As String = doc.FullFileName
            If String.IsNullOrEmpty(fullPath) Then
                MessageBox.Show("Document has not been saved yet.")
                Exit Sub
            End If

            ' 在资源管理器中打开并选中
            Process.Start("explorer.exe", "/select,""" & fullPath & """")
        End Sub
    End Module
End Namespace