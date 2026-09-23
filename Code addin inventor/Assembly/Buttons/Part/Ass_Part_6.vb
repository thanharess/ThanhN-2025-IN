Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Part
    Public Module Ass_Part_6
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Executed Assembly Action 12")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Assembly Action 12: " & ex.Message)
                Catch
                End Try
            End Try


            ' Lấy ứng dụng Inventor đang chạy
            Dim invApp As Inventor.Application = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)

            ' Lấy document hiện tại
            Dim asmDoc As AssemblyDocument = TryCast(invApp.ActiveDocument, AssemblyDocument)

            If asmDoc Is Nothing Then
                MessageBox.Show("This is not an assembly!", "Visual Studio")
                Return
            End If

            ' Lặp qua các đối tượng có override
            For Each obj As ComponentOccurrence In asmDoc.ComponentDefinition.AppearanceOverridesObjects
                Try
                    ' Trả về màu gốc của part
                    obj.AppearanceSourceType = AppearanceSourceTypeEnum.kPartAppearance
                Catch ex As Exception
                    ' Bỏ qua lỗi nếu có
                End Try
            Next
            asmDoc.Update()
            MessageBox.Show("Đã reset màu về gốc cho các Occurrence", "Thanh")

        End Sub

    End Module
End Namespace
