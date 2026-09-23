Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Part
    Public Module Ass_Part_1
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application
            Try
                invApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch ex As Exception
                MessageBox.Show("Inventor chưa chạy.")
                Return
            End Try

            ' Lấy Assembly đang mở
            Dim doc As AssemblyDocument = TryCast(invApp.ActiveDocument, AssemblyDocument)
            If doc Is Nothing Then
                MessageBox.Show("Không phải Assembly Document.")
                Return
            End If

            Try
                '========================================================
                ' Kiểm tra Assembly đang hoạt động
                '========================================================
                Dim oADoc As AssemblyDocument =
                    TryCast(g_inventorApplication.ActiveDocument, AssemblyDocument)

                If oADoc Is Nothing Then
                    MessageBox.Show(
                        "Vui lòng mở Assembly trước.",
                        "Chuyển sang Sheet Metal",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
                    Return
                End If


                '========================================================
                ' Cho phép chọn nhiều chi tiết liên tiếp
                ' Nhấn ESC / Cancel để kết thúc
                '========================================================
                While True

                    Dim comp As ComponentOccurrence = Nothing

                    Try
                        ' Chọn trực tiếp chi tiết trong Assembly
                        comp = TryCast(
                            g_inventorApplication.CommandManager.Pick(
                                SelectionFilterEnum.kAssemblyLeafOccurrenceFilter,
                                "Chọn chi tiết cần chuyển đổi sang thép tấm"),
                            ComponentOccurrence)

                    Catch
                        ' ESC hoặc Cancel
                        Exit While
                    End Try


                    If comp Is Nothing Then
                        Exit While
                    End If


                    '====================================================
                    ' Lấy Part Document
                    '====================================================
                    Dim oPartDoc As PartDocument = Nothing

                    Try
                        oPartDoc = TryCast(comp.Definition.Document, PartDocument)
                    Catch
                        oPartDoc = Nothing
                    End Try

                    If oPartDoc Is Nothing Then
                        MessageBox.Show(
                            "Chi tiết được chọn không phải Part.",
                            "Chuyển sang Sheet Metal",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
                        Continue While
                    End If


                    '====================================================
                    ' Chuyển Part sang Sheet Metal
                    '====================================================
                    Try

                        ' Sheet Metal SubType
                        oPartDoc.SubType =
                            "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"


                        ' Lấy Sheet Metal Component Definition
                        Dim oSheetMetalCompDef As SheetMetalComponentDefinition =
                            TryCast(oPartDoc.ComponentDefinition,
                                    SheetMetalComponentDefinition)

                        If oSheetMetalCompDef Is Nothing Then

                            MessageBox.Show(
                                "Không thể chuyển chi tiết này sang Sheet Metal.",
                                "Chuyển sang Sheet Metal",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)

                            Continue While

                        End If


                        '================================================
                        ' Không sử dụng độ dày mặc định của Sheet Metal Style
                        '================================================
                        oSheetMetalCompDef.UseSheetMetalStyleThickness = False


                        '================================================
                        ' Lấy Parameter Thickness
                        '================================================
                        Dim oThicknessParam As Parameter =
                            oSheetMetalCompDef.Thickness


                        '================================================
                        ' Đọc độ dày hiện tại
                        '
                        ' Inventor API sử dụng cm cho Value
                        ' nên nhân 10 để chuyển sang mm
                        '================================================
                        Dim currentThickness As Double =
                            oThicknessParam.Value * 10.0


                        '================================================
                        ' Nhập độ dày mới
                        '================================================
                        Dim inputThickness As String =
                            InputBox(
                                "Nhập độ dày tấm (mm):",
                                "Chuyển sang Sheet Metal",
                                currentThickness.ToString("0.##"))


                        '================================================
                        ' Nếu Cancel hoặc để trống
                        '================================================
                        If String.IsNullOrWhiteSpace(inputThickness) Then

                            ' Nếu đã chuyển SubType nhưng người dùng Cancel
                            ' thì vẫn giữ Part ở dạng Sheet Metal.
                            ' Nếu muốn quay lại Part có thể xử lý thêm ở đây.

                            Continue While

                        End If


                        '================================================
                        ' Chuyển giá trị nhập sang Double
                        '================================================
                        Dim oNewTHK As Double

                        If Not Double.TryParse(
                            inputThickness,
                            Globalization.NumberStyles.Any,
                            Globalization.CultureInfo.CurrentCulture,
                            oNewTHK) Then

                            MessageBox.Show(
                                "Độ dày nhập vào không hợp lệ.",
                                "Chuyển sang Sheet Metal",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)

                            Continue While

                        End If


                        If oNewTHK <= 0 Then

                            MessageBox.Show(
                                "Độ dày phải lớn hơn 0 mm.",
                                "Chuyển sang Sheet Metal",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)

                            Continue While

                        End If


                        '================================================
                        ' Gán Thickness
                        '
                        ' mm -> cm
                        '================================================
                        oThicknessParam.Value = oNewTHK / 10.0


                        '================================================
                        ' Update Part
                        '================================================
                        oPartDoc.Update()


                        '================================================
                        ' Update Assembly
                        '================================================
                        oADoc.Update()


                    Catch ex As Exception

                        MessageBox.Show(
                            "Không thể chuyển đổi chi tiết:" &
                            vbCrLf & comp.Name &
                            vbCrLf & vbCrLf &
                            ex.Message,
                            "Lỗi",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)

                    End Try

                End While


            Catch ex As Exception

                MessageBox.Show(
                    "Lỗi:" &
                    vbCrLf & ex.Message,
                    "Convert To Sheet Metal",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub



    End Module
End Namespace
