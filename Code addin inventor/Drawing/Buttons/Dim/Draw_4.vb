Imports System
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Drawing.Buttons
    Public Module Draw_4
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                ' Kết nối tới Inventor (lấy instance đang chạy)
                Dim invApp As Inventor.Application = Nothing
                Try
                    invApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
                Catch ex As COMException
                    MessageBox.Show("Không tìm thấy Inventor đang chạy. Vui lòng mở Inventor và một Drawing trước khi chạy.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End Try

                ' Lấy document hiện hành và kiểm tra là Drawing
                Dim oDoc As Document = invApp.ActiveDocument
                If oDoc Is Nothing OrElse oDoc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Vui lòng mở một Drawing document trong Inventor trước khi chạy.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim oDrawDoc As DrawingDocument = CType(oDoc, DrawingDocument)

                ' Tên transaction (undo wrapper)
                Dim oNamer As String = "Highlight Dimension Overrides"

                ' Bắt transaction
                Dim UNDO As Transaction = Nothing
                Try
                    UNDO = invApp.TransactionManager.StartTransaction(oDoc, oNamer)
                Catch ex As Exception
                    MessageBox.Show("Không thể bắt transaction: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End Try

                Try
                    ' Lấy sheet hiện hành
                    Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                    If oSheet Is Nothing Then
                        MessageBox.Show("Không thể truy cập ActiveSheet.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ' Kết thúc transaction an toàn
                        Try
                            UNDO.End()
                        Catch
                        End Try
                        Return
                    End If

                    ' Tạo màu đỏ và màu đen
                    Dim redColor As Color = invApp.TransientObjects.CreateColor(255, 0, 0)
                    Dim blackColor As Color = invApp.TransientObjects.CreateColor(0, 0, 0)

                    Dim DimCount As Integer = 0

                    ' Duyệt tất cả DrawingDimensions trên sheet
                    For Each oDim As DrawingDimension In oSheet.DrawingDimensions
                        Try
                            ' Một số thuộc tính có thể không tồn tại cho mọi loại dimension, nên dùng Try/Catch
                            Dim isOverridden As Boolean = False
                            Dim hideVal As Boolean = False

                            Try
                                isOverridden = (oDim.OverrideModelValue <> oDim.ModelValue)
                            Catch
                                ' Nếu không có OverrideModelValue/ModelValue, giữ false
                                isOverridden = False
                            End Try

                            Try
                                hideVal = oDim.HideValue
                            Catch
                                hideVal = False
                            End Try

                            If isOverridden OrElse hideVal Then
                                ' Gán màu chữ đỏ
                                Try
                                    oDim.Text.Color = redColor
                                Catch
                                    ' Nếu không thể gán trực tiếp, bỏ qua
                                End Try
                                DimCount += 1
                            Else
                                ' Gán màu chữ đen
                                Try
                                    oDim.Text.Color = blackColor
                                Catch
                                End Try
                            End If
                        Catch exDim As Exception
                            ' Bỏ qua dimension gây lỗi, tiếp tục
                            Console.WriteLine("Lỗi xử lý dimension: " & exDim.Message)
                        End Try
                    Next

                    ' Cập nhật document (tương đương iLogicVb.DocumentUpdate)
                    Try
                        oDrawDoc.Update()
                    Catch
                        Try
                            oDrawDoc.Update2()
                        Catch
                        End Try
                    End Try

                    ' Thông báo kết quả
                    If DimCount > 0 Then
                        MessageBox.Show("Có " & DimCount & " kích thước bị ghi đè.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Không có kích thước bị ghi đè.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                    ' Kết thúc transaction (commit)
                    Try
                        UNDO.End()
                    Catch exEnd As Exception
                        ' Nếu End thất bại, cố gắng rollback (nếu API hỗ trợ) hoặc thông báo
                        MessageBox.Show("Không thể kết thúc transaction: " & exEnd.Message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Try

                Catch exInner As Exception
                    ' Nếu có lỗi trong quá trình xử lý, cố gắng kết thúc transaction rồi báo lỗi
                    Try
                        If UNDO IsNot Nothing Then UNDO.End()
                    Catch
                    End Try
                    MessageBox.Show("Lỗi khi xử lý: " & exInner.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Catch ex As Exception
                MessageBox.Show("Lỗi không mong muốn: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Module

End Namespace
