Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep
    Public Module Ass_LG_1
        Public Sub OnExecute(ByVal Context As NameValueMap)



            ' Lấy ứng dụng Inventor đang chạy
            Dim invApp As Inventor.Application = Interop.Marshal2.GetActiveObject("Inventor.Application")

            Dim oAsmDoc As Document = invApp.ActiveDocument

            If oAsmDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("This is not an assembly!", "Visual Studio")
                Return
            End If

            Dim i As Integer = 0

            For Each oDoc As Document In oAsmDoc.AllReferencedDocuments
                Try
                    If oDoc.ModelingSettings.AdaptivelyUsedInAssembly = True Then
                        oDoc.ModelingSettings.AdaptivelyUsedInAssembly = False
                        i += 1
                    End If
                Catch ex As Exception
                    ' Bỏ qua lỗi
                End Try
            Next

            MessageBox.Show("Turned off adaptivity in " & i & " documents", "Visual Studio")
        End Sub
    End Module

End Namespace