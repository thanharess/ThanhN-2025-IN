Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Part
    Public Module Ass_Part_7
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Executed Assembly Action 11")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Assembly Action 11: " & ex.Message)
                Catch
                End Try
            End Try

            ' Lấy ứng dụng Inventor đang chạy
            Dim invApp As Inventor.Application = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)

            Dim oAsmDoc As Document = invApp.ActiveDocument

            If oAsmDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("This is not an assembly!", " ")
                Return
            End If


            ' Lặp qua tất cả Occurrences trong Assembly
            For Each occ As ComponentOccurrence In oAsmDoc.ComponentDefinition.Occurrences
                Try
                    Dim d As Document = occ.Definition.Document
                    Dim compDef As ComponentDefinition = d.ComponentDefinition

                    ' Ẩn các mặt phẳng gốc
                    compDef.WorkPlanes.Item("XY Plane").Visible = False
                    compDef.WorkPlanes.Item("YZ Plane").Visible = False
                    compDef.WorkPlanes.Item("XZ Plane").Visible = False
                Catch ex As Exception
                    ' Bỏ qua lỗi nếu có
                End Try
            Next
            oAsmDoc.Update()
            MessageBox.Show("Đã an plane goc toa do", "Visual Studio")
        End Sub

    End Module
End Namespace
