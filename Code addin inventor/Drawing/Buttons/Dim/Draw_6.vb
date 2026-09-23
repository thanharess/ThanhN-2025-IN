Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic

Namespace ToolInventor2025.Drawing.Buttons.Drawdim
    Public Module Dim_Linear_Total

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Vui lòng mở file Drawing (.idw)!", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument = CType(app.ActiveDocument, DrawingDocument)
                Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                Dim tg As TransientGeometry = app.TransientGeometry

                '===== 1. CHỌN NHIỀU VIEW =====
                Dim selectedViews As New List(Of DrawingView)

                Do
                    Try
                        oDrawDoc.SelectSet.Clear()
                    Catch
                    End Try

                    Dim pickedObj As Object = Nothing
                    Try
                        pickedObj = app.CommandManager.Pick(
                            SelectionFilterEnum.kDrawingViewFilter,
                            "Chọn View (ESC để kết thúc)")
                    Catch
                        Exit Do
                    End Try

                    If pickedObj Is Nothing Then Exit Do

                    Dim v As DrawingView = TryCast(pickedObj, DrawingView)
                    If v Is Nothing Then Continue Do

                    Dim dup As Boolean = False
                    For Each x As DrawingView In selectedViews
                        If x Is v Then
                            dup = True
                            Exit For
                        End If
                    Next
                    If Not dup Then selectedViews.Add(v)

                    Dim more As DialogResult = MessageBox.Show(
                        "Đã chọn " & selectedViews.Count & " view. Chọn tiếp?",
                        "Dim tổng",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question)
                    If more = DialogResult.No Then Exit Do
                Loop

                If selectedViews.Count = 0 Then
                    MessageBox.Show("Chưa chọn view nào.", "Thông báo")
                    Exit Sub
                End If

                '===== 2. HỎI KHOẢNG CÁCH ĐẶT DIM =====
                Dim sOffset As String = InputBox(
                    "Khoảng cách đặt dim ngoài view (mm):",
                    "Dim tổng", "5.5")
                If sOffset Is Nothing Then Exit Sub

                Dim offsetCM As Double = 0.55
                Double.TryParse(sOffset.Replace(",", ".").Trim(),
                                Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.InvariantCulture,
                                offsetCM)
                offsetCM = offsetCM / 10.0

                '===== 3. XỬ LÝ TỪNG VIEW =====
                Dim nOK As Integer = 0
                Dim nFail As Integer = 0
                Dim log As New System.Text.StringBuilder()

                For Each oView As DrawingView In selectedViews

                    '-------------------------------------------------
                    ' Tìm 4 cạnh ngoài
                    '-------------------------------------------------
                    Dim leftEdge As DrawingCurve = Nothing
                    Dim rightEdge As DrawingCurve = Nothing
                    Dim topEdge As DrawingCurve = Nothing
                    Dim bottomEdge As DrawingCurve = Nothing

                    Dim minX As Double = Double.MaxValue
                    Dim maxX As Double = Double.MinValue
                    Dim minY As Double = Double.MaxValue
                    Dim maxY As Double = Double.MinValue

                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType <> CurveTypeEnum.kLineCurve AndAlso
                               oCurve.CurveType <> CurveTypeEnum.kLineSegmentCurve Then Continue For

                            Dim p1 As Point2d = oCurve.StartPoint
                            Dim p2 As Point2d = oCurve.EndPoint
                            If p1 Is Nothing OrElse p2 Is Nothing Then Continue For

                            If Math.Abs(p1.X - p2.X) < 0.03 Then
                                If p1.X < minX Then
                                    minX = p1.X
                                    leftEdge = oCurve
                                End If
                                If p1.X > maxX Then
                                    maxX = p1.X
                                    rightEdge = oCurve
                                End If
                            End If

                            If Math.Abs(p1.Y - p2.Y) < 0.03 Then
                                If p1.Y > maxY Then
                                    maxY = p1.Y
                                    topEdge = oCurve
                                End If
                                If p1.Y < minY Then
                                    minY = p1.Y
                                    bottomEdge = oCurve
                                End If
                            End If

                        Catch
                        End Try
                    Next

                    If leftEdge Is Nothing OrElse rightEdge Is Nothing OrElse
                       topEdge Is Nothing OrElse bottomEdge Is Nothing Then
                        nFail += 1
                        log.AppendLine("  ✘ View " & oView.Name & ": không tìm đủ 4 cạnh")
                        Continue For
                    End If

                    If (maxX - minX) <= 0 OrElse (maxY - minY) <= 0 Then
                        nFail += 1
                        Continue For
                    End If

                    '-------------------------------------------------
                    ' DIM NGANG (W) — đặt bên trên
                    '-------------------------------------------------
                    Try
                        Dim tp As Point2d = tg.CreatePoint2d((minX + maxX) / 2, maxY + offsetCM)
                        oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                            tp,
                            oSheet.CreateGeometryIntent(leftEdge),
                            oSheet.CreateGeometryIntent(rightEdge),
                            DimensionTypeEnum.kHorizontalDimensionType)
                        nOK += 1
                        log.AppendLine("  ✔ " & oView.Name & ": dim W = " &
                                       ((maxX - minX) * 10.0).ToString("0.0#") & " mm")
                    Catch ex As Exception
                        nFail += 1
                        log.AppendLine("  ✘ " & oView.Name & " dim W: " & ex.Message)
                    End Try

                    '-------------------------------------------------
                    ' DIM DỌC (H) — đặt bên trái
                    '-------------------------------------------------
                    Try
                        Dim tp As Point2d = tg.CreatePoint2d(minX - offsetCM, (maxY + minY) / 2)
                        oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                            tp,
                            oSheet.CreateGeometryIntent(topEdge),
                            oSheet.CreateGeometryIntent(bottomEdge),
                            DimensionTypeEnum.kVerticalDimensionType)
                        nOK += 1
                        log.AppendLine("  ✔ " & oView.Name & ": dim H = " &
                                       ((maxY - minY) * 10.0).ToString("0.0#") & " mm")
                    Catch ex As Exception
                        nFail += 1
                        log.AppendLine("  ✘ " & oView.Name & " dim H: " & ex.Message)
                    End Try

                Next

                oDrawDoc.Update()

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Số view: " & selectedViews.Count & vbCrLf &
                    "Dim tạo: " & nOK & vbCrLf &
                    "Lỗi: " & nFail & vbCrLf & vbCrLf &
                    "--- Chi tiết ---" & vbCrLf &
                    log.ToString(),
                    "Dim Linear tổng",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "Dim Linear tổng",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)
            End Try

        End Sub

    End Module
End Namespace