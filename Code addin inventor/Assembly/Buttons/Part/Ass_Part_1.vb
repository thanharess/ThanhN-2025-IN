Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Part

    Public Module Ass_Part_1

        Private Const SHEETMETAL_SUBTYPE As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

        ' Đơn vị nội bộ Inventor: cm
        Private Const MIN_VALID_THICKNESS_CM As Double = 0.001   ' 0.005 mm
        Private Const MAX_VALID_THICKNESS_CM As Double = 20.0    ' 100 mm

        '=====================================================
        ' ENTRY POINT
        '=====================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)

            '--- 1. Kết nối Inventor ---
            Dim invApp As Inventor.Application
            Try
                invApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"),
                               Inventor.Application)
            Catch
                MessageBox.Show("Inventor chưa chạy.", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End Try

            '--- 2. Kiểm tra Assembly ---
            Dim oADoc As AssemblyDocument = TryCast(invApp.ActiveDocument, AssemblyDocument)
            If oADoc Is Nothing Then
                MessageBox.Show("Vui lòng mở Assembly trước.", "Chuyển sang Sheet Metal",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            '--- 3. Vòng lặp chọn part ---
            While True

                Dim comp As ComponentOccurrence = Nothing
                Try
                    comp = TryCast(
                        invApp.CommandManager.Pick(
                            SelectionFilterEnum.kAssemblyLeafOccurrenceFilter,
                            "Chọn chi tiết (ESC để kết thúc)"),
                        ComponentOccurrence)
                Catch
                    Exit While
                End Try
                If comp Is Nothing Then Exit While

                '--- 4. Lấy PartDocument ---
                Dim oPartDoc As PartDocument = Nothing
                Try
                    oPartDoc = TryCast(comp.Definition.Document, PartDocument)
                Catch
                End Try

                If oPartDoc Is Nothing Then
                    MessageBox.Show("Chi tiết được chọn không phải Part.",
                                    "Thông báo",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Continue While
                End If

                '=================================================
                ' 5. XÁC ĐỊNH PART ĐÃ LÀ SHEET METAL CHƯA
                '=================================================
                Dim oSMDef As SheetMetalComponentDefinition = Nothing
                Dim isAlreadySM As Boolean = False

                Try
                    oSMDef = TryCast(oPartDoc.ComponentDefinition, SheetMetalComponentDefinition)
                    If oSMDef IsNot Nothing Then isAlreadySM = True
                Catch
                End Try

                '=================================================
                ' 6. NẾU CHƯA PHẢI SM → KIỂM TRA + CONVERT
                '=================================================
                If Not isAlreadySM Then

                    ' 6a) Kiểm tra khả năng chuyển
                    Dim reason As String = ""
                    If Not IsConvertibleToSheetMetal(oPartDoc, reason) Then
                        MessageBox.Show(
                            "Không thể chuyển chi tiết này." & vbCrLf & vbCrLf &
                            "Lý do: " & reason & vbCrLf & vbCrLf &
                            "Chi tiết: " & comp.Name,
                            "Chuyển sang Sheet Metal",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Continue While
                    End If

                    ' 6b) Đổi SubType
                    Try
                        oPartDoc.SubType = SHEETMETAL_SUBTYPE
                    Catch ex As Exception
                        MessageBox.Show("Không đổi được SubType: " & ex.Message,
                                        "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Continue While
                    End Try

                    ' 6c) Lấy SM ComponentDefinition sau khi convert
                    Try
                        oSMDef = TryCast(oPartDoc.ComponentDefinition,
                                         SheetMetalComponentDefinition)
                    Catch
                    End Try

                    If oSMDef Is Nothing Then
                        MessageBox.Show("Không thể chuyển sang Sheet Metal.",
                                        "Lỗi",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Continue While
                    End If
                End If

                '=================================================
                ' 7. ÉP ĐƠN VỊ mm + TẮT STYLE THICKNESS
                '=================================================
                ' ⭐ Ép document hiển thị mm (Part)
                Try
                    oPartDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits
                Catch
                End Try

                ' ⭐ Ép luôn Assembly hiển thị mm
                Try
                    oADoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits
                Catch
                End Try

                ' Không dùng thickness mặc định từ style
                Try : oSMDef.UseSheetMetalStyleThickness = False : Catch : End Try

                '=================================================
                ' 8. GỢI Ý + HỎI ĐỘ DÀY
                '=================================================
                Dim suggestedMm As Double = GetSuggestedThicknessMm(oSMDef, oPartDoc)

                Dim inputStr As String = InputBox(
                    "Nhập độ dày tấm (mm):",
                    If(isAlreadySM, "Đổi độ dày Sheet Metal", "Chuyển sang Sheet Metal"),
                    suggestedMm.ToString("0.##"))

                If String.IsNullOrWhiteSpace(inputStr) Then Continue While

                Dim userThkMm As Double
                If Not Double.TryParse(inputStr,
                                       Globalization.NumberStyles.Any,
                                       Globalization.CultureInfo.CurrentCulture,
                                       userThkMm) OrElse userThkMm <= 0 Then
                    MessageBox.Show("Độ dày không hợp lệ.", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Continue While
                End If

                '=================================================
                ' 9. GÁN ĐỘ DÀY (mm → cm)
                '=================================================
                Dim thkCm As Double = userThkMm / 10.0

                If thkCm < MIN_VALID_THICKNESS_CM OrElse thkCm > MAX_VALID_THICKNESS_CM Then
                    MessageBox.Show("Độ dày ngoài khoảng cho phép (0.05 – 200 mm).",
                                    "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Continue While
                End If

                Try
                    oSMDef.Thickness.Value = thkCm
                Catch ex As Exception
                    MessageBox.Show("Không gán được độ dày: " & ex.Message,
                                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Continue While
                End Try

                '=================================================
                ' 10. UPDATE + THÔNG BÁO
                '=================================================
                Try : oPartDoc.Update2(True) : Catch : End Try
                Try : oADoc.Update2(True) : Catch : End Try

                '  MessageBox.Show(If(isAlreadySM, "Đã đổi độ dày: ", "Đã chuyển: ") & comp.Name, "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End While

        End Sub

        '=====================================================
        ' KIỂM TRA KHẢ NĂNG CHUYỂN SANG SHEET METAL
        '=====================================================
        Private Function IsConvertibleToSheetMetal(ByVal pDoc As PartDocument,
                                                   ByRef reason As String) As Boolean
            reason = ""

            Try
                ' 1) Phải có ít nhất 1 solid body
                Dim bodies As SurfaceBodies = pDoc.ComponentDefinition.SurfaceBodies
                If bodies Is Nothing OrElse bodies.Count = 0 Then
                    reason = "Part không có khối hình học."
                    Return False
                End If

                ' 2) Đếm số mặt — bu lông/bánh răng có rất nhiều mặt
                Dim body As SurfaceBody = bodies.Item(1)
                Dim faceCount As Integer = body.Faces.Count

                If faceCount > 80 Then
                    reason = "Part quá phức tạp (" & faceCount & " mặt)."
                    Return False
                End If

                ' 3) Tỷ lệ 3 chiều — chặn khối gần lập phương
                Dim bbox As Box = body.RangeBox
                Dim dX As Double = bbox.MaxPoint.X - bbox.MinPoint.X
                Dim dY As Double = bbox.MaxPoint.Y - bbox.MinPoint.Y
                Dim dZ As Double = bbox.MaxPoint.Z - bbox.MinPoint.Z

                Dim dims() As Double = {dX, dY, dZ}
                Array.Sort(dims)

                Dim minD As Double = dims(0)
                Dim maxD As Double = dims(2)

                If maxD <= 0 Then
                    reason = "Kích thước không hợp lệ."
                    Return False
                End If

                If (minD / maxD) > 0.6 Then
                    reason = "Hình dạng quá khối / gần lập phương."
                    Return False
                End If

                Return True

            Catch ex As Exception
                reason = ex.Message
                Return False
            End Try
        End Function

        '=====================================================
        ' GỢI Ý ĐỘ DÀY (mm)
        ' Ưu tiên: Thickness hiện tại → BoundingBox → 3mm
        '=====================================================
        Private Function GetSuggestedThicknessMm(ByVal smDef As SheetMetalComponentDefinition,
                                                 ByVal pDoc As PartDocument) As Double

            ' 1) Thickness hiện tại (giá trị API luôn là cm)
            Try
                Dim tCm As Double = smDef.Thickness.Value
                If tCm >= MIN_VALID_THICKNESS_CM AndAlso tCm <= MAX_VALID_THICKNESS_CM Then
                    Return Math.Round(tCm * 10.0, 2)
                End If
            Catch
            End Try

            ' 2) Bounding box — cạnh nhỏ nhất
            Try
                Dim body As SurfaceBody = pDoc.ComponentDefinition.SurfaceBodies.Item(1)
                Dim bbox As Box = body.RangeBox
                Dim dX As Double = bbox.MaxPoint.X - bbox.MinPoint.X
                Dim dY As Double = bbox.MaxPoint.Y - bbox.MinPoint.Y
                Dim dZ As Double = bbox.MaxPoint.Z - bbox.MinPoint.Z
                Dim tCm As Double = Math.Min(dX, Math.Min(dY, dZ))

                If tCm >= MIN_VALID_THICKNESS_CM AndAlso tCm <= MAX_VALID_THICKNESS_CM Then
                    Return Math.Round(tCm * 10.0, 2)
                End If
            Catch
            End Try

            ' 3) Fallback
            Return 3.0
        End Function

    End Module
End Namespace