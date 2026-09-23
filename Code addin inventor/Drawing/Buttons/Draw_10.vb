Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Drawing.Buttons
    Public Module Draw_5
        Public Sub OnExecute(ByVal Context As NameValueMap)


            ' Lấy ứng dụng Inventor đang chạy
            Dim inventorApp As Inventor.Application = Nothing
            Try
                inventorApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch ex As Exception
                Console.WriteLine("Không tìm thấy Inventor đang chạy: " & ex.Message)
                Return
            End Try

            ' Lấy tài liệu bản vẽ đang mở
            Dim oDrawDoc As DrawingDocument = CType(inventorApp.ActiveDocument, DrawingDocument)
            Dim oSheet As Sheet = oDrawDoc.ActiveSheet

            ' Vòng lặp chọn và xóa đường tâm
            Do
                ' Cho phép người dùng chọn đối tượng đường tâm
                Dim oViewSelect As Object = inventorApp.CommandManager.Pick(SelectionFilterEnum.kDrawingCentermarkFilter, "Chọn đường tâm : tkn")

                If oViewSelect IsNot Nothing Then
                    Try
                        oViewSelect.Delete()
                    Catch ex As Exception
                        Console.WriteLine("Không thể xóa đối tượng: " & ex.Message)
                    End Try
                End If

                Console.WriteLine("Nhấn Y để tiếp tục xóa, phím khác để thoát:")
                Dim key As ConsoleKeyInfo = Console.ReadKey()
                If key.Key <> ConsoleKey.Y Then
                    Exit Do
                End If
            Loop
        End Sub
    End Module

End Namespace
