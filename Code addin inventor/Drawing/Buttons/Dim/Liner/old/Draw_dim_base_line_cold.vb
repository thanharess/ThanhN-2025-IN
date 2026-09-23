Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq

Namespace ToolInventor2025.Drawing.Buttons.Drawdim
    Public Module Draw_dim_base_line_caaa

        Private Const TOL As Double = 0.00
        Private Const EDGE_TOL As Double = 0.03

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

                '===== CHỌN VIEW =====
                Dim selectedViews As New List(Of DrawingView)
                Do
                    Dim oSS As SelectSet = oDrawDoc.SelectSet
                    oSS.Clear()
                    Dim oView As DrawingView = Nothing
                    Try
                        oView = CType(app.CommandManager.Pick(
                            SelectionFilterEnum.kDrawingViewFilter,
                            "Chọn View (Esc hoặc Right-click để kết thúc)"), DrawingView)
                    Catch
                        Exit Do
                    End Try
                    If oView Is Nothing Then Exit Do

                    Dim already As Boolean = False
                    For Each v As DrawingView In selectedViews
                        If v Is oView Then already = True : Exit For
                    Next
                    If Not already Then selectedViews.Add(oView)
                Loop

                If selectedViews.Count = 0 Then
                    MessageBox.Show("Chưa chọn View nào.", "Thông báo")
                    Exit Sub
                End If

                '===== FORM =====
                Dim hFromRight As Boolean = False
                Dim vFromBottom As Boolean = False
                Dim hPickedEdge As DrawingCurve = Nothing
                Dim vPickedEdge As DrawingCurve = Nothing
                Dim includeHoles As Boolean = True

                Dim xNearGap As Double = 3.2   ' mm
                Dim yNearGap As Double = 3.2
                Dim minGapX As Double = 5.0
                Dim minGapY As Double = 5.0
                Dim useTier As Boolean = True   ' bật/tắt khoảng cách bậc

                If Not ShowDirectionDialog(hFromRight, vFromBottom,
                           hPickedEdge, vPickedEdge, includeHoles,
                           xNearGap, yNearGap, minGapX, minGapY, useTier, app) Then Exit Sub

                Dim countOK As Integer = 0
                Dim countFail As Integer = 0
                Dim dimmedCurves As List(Of DrawingCurve) = CollectDimmedCurves(oSheet)

                For Each oView As DrawingView In selectedViews

                    Dim allEdges As New List(Of EdgeInfo)

                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType = CurveTypeEnum.kLineCurve OrElse
                               oCurve.CurveType = CurveTypeEnum.kLineSegmentCurve Then

                                Dim p1 As Point2d = oCurve.StartPoint
                                Dim p2 As Point2d = oCurve.EndPoint
                                If p1 Is Nothing OrElse p2 Is Nothing Then Continue For

                                Dim dx As Double = Math.Abs(p1.X - p2.X)
                                Dim dy As Double = Math.Abs(p1.Y - p2.Y)
                                If dx < EDGE_TOL AndAlso dy < EDGE_TOL Then Continue For

                                Dim isV As Boolean = dx < EDGE_TOL
                                Dim isH As Boolean = dy < EDGE_TOL
                                If Not isV AndAlso Not isH Then Continue For

                                Dim ei As New EdgeInfo With {
                                    .Curve = oCurve, .IsVertical = isV, .IsHorizontal = isH,
                                    .IsHole = False, .X = (p1.X + p2.X) / 2.0, .Y = (p1.Y + p2.Y) / 2.0, .Dimmed = False}
                                allEdges.Add(ei)

                            ElseIf includeHoles AndAlso oCurve.CurveType = CurveTypeEnum.kCircleCurve Then
                                Dim c As Point2d = oCurve.CenterPoint
                                If c Is Nothing Then Continue For
                                Dim hi As New EdgeInfo With {
                                    .Curve = oCurve, .IsVertical = False, .IsHorizontal = False,
                                    .IsHole = True, .X = c.X, .Y = c.Y, .Dimmed = False}
                                allEdges.Add(hi)
                            End If
                        Catch
                        End Try
                    Next

                    If allEdges.Count = 0 Then Continue For

                    ' Dedup
                    Dim uniqueEdges As New List(Of EdgeInfo)
                    For Each e As EdgeInfo In allEdges
                        Dim dup As Boolean = False
                        For Each u As EdgeInfo In uniqueEdges
                            If e.IsHole AndAlso u.IsHole Then
                                If Math.Abs(u.X - e.X) <= EDGE_TOL AndAlso Math.Abs(u.Y - e.Y) <= EDGE_TOL Then
                                    dup = True : Exit For
                                End If
                            ElseIf (Not e.IsHole) AndAlso (Not u.IsHole) Then
                                If SameSegment(u.Curve, e.Curve) Then
                                    dup = True : Exit For
                                End If
                            End If
                        Next
                        If Not dup Then uniqueEdges.Add(e)
                    Next
                    allEdges = uniqueEdges

                    ' 4 cạnh ngoài
                    Dim leftEdge As DrawingCurve = Nothing, rightEdge As DrawingCurve = Nothing
                    Dim topEdge As DrawingCurve = Nothing, bottomEdge As DrawingCurve = Nothing
                    Dim minX As Double = Double.MaxValue, maxX As Double = Double.MinValue
                    Dim minY As Double = Double.MaxValue, maxY As Double = Double.MinValue

                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType <> CurveTypeEnum.kLineCurve AndAlso
                               oCurve.CurveType <> CurveTypeEnum.kLineSegmentCurve Then Continue For

                            Dim p1 As Point2d = oCurve.StartPoint
                            Dim p2 As Point2d = oCurve.EndPoint
                            If p1 Is Nothing OrElse p2 Is Nothing Then Continue For

                            If Math.Abs(p1.X - p2.X) < EDGE_TOL Then
                                If p1.X < minX Then minX = p1.X : leftEdge = oCurve
                                If p1.X > maxX Then maxX = p1.X : rightEdge = oCurve
                            End If
                            If Math.Abs(p1.Y - p2.Y) < EDGE_TOL Then
                                If p1.Y > maxY Then maxY = p1.Y : topEdge = oCurve
                                If p1.Y < minY Then minY = p1.Y : bottomEdge = oCurve
                            End If
                        Catch
                        End Try
                    Next

                    If leftEdge Is Nothing OrElse rightEdge Is Nothing OrElse
                       topEdge Is Nothing OrElse bottomEdge Is Nothing Then Continue For

                    If (maxX - minX) <= 0 OrElse (maxY - minY) <= 0 Then Continue For

                    Dim usedX As New List(Of Double)
                    Dim usedY As New List(Of Double)

                    ' Base
                    Dim hBaseEdge As DrawingCurve
                    Dim hBaseCoord As Double
                    If EdgeInView(hPickedEdge, oView) Then
                        hBaseEdge = hPickedEdge
                        hBaseCoord = (hPickedEdge.StartPoint.X + hPickedEdge.EndPoint.X) / 2.0
                    Else
                        hBaseEdge = If(hFromRight, rightEdge, leftEdge)
                        hBaseCoord = If(hFromRight, maxX, minX)
                    End If

                    Dim vBaseEdge As DrawingCurve
                    Dim vBaseCoord As Double
                    If EdgeInView(vPickedEdge, oView) Then
                        vBaseEdge = vPickedEdge
                        vBaseCoord = (vPickedEdge.StartPoint.Y + vPickedEdge.EndPoint.Y) / 2.0
                    Else
                        vBaseEdge = If(vFromBottom, bottomEdge, topEdge)
                        vBaseCoord = If(vFromBottom, minY, maxY)
                    End If

                    Dim hBaseOnRight As Boolean = (hBaseCoord > (minX + maxX) / 2.0)
                    Dim vBaseOnBottom As Boolean = (vBaseCoord < (minY + maxY) / 2.0)

                    Dim hDimLineY As Double = If(hBaseOnRight, minY - xNearGap, maxY + xNearGap)
                    Dim vDimLineX As Double = If(vBaseOnBottom, maxX + yNearGap, minX - yNearGap)

                    '========== DIM X (ngang) ==========
                    Dim xTargets = allEdges.Where(Function(e) e.IsVertical OrElse e.IsHole).ToList()
                    If hBaseOnRight Then
                        xTargets = xTargets.OrderByDescending(Function(e) e.X).ToList()
                    Else
                        xTargets = xTargets.OrderBy(Function(e) e.X).ToList()
                    End If

                    Dim tierIndexX As Integer = 0
                    Dim tierStepX As Double = If(useTier, 0.4, 0.0)   ' 4 mm hoặc 0

                    For Each e As EdgeInfo In xTargets
                        If Math.Abs(e.X - hBaseCoord) <= EDGE_TOL Then Continue For
                        If CurveHasDim(dimmedCurves, e.Curve) Then Continue For
                        If Math.Abs(e.X - hBaseCoord) < minGapX Then Continue For
                        If ValueExists(usedX, e.X) Then Continue For          ' ← chống trùng

                        Dim offsetY As Double = xNearGap + tierIndexX * tierStepX
                        If useTier Then tierIndexX += 1

                        Dim placeY As Double = If(hBaseOnRight, minY - offsetY, maxY + offsetY)

                        Try
                            Dim baseIntent = oSheet.CreateGeometryIntent(hBaseEdge)
                            Dim edgeIntent = MakeIntent(oSheet, e)
                            Dim placement = tg.CreatePoint2d(e.X, placeY)

                            Dim newDim = oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                placement, baseIntent, edgeIntent, DimensionTypeEnum.kHorizontalDimensionType)

                            If newDim IsNot Nothing Then
                                usedX.Add(e.X)
                                e.Dimmed = True
                                countOK += 1
                                dimmedCurves.Add(e.Curve)
                            End If
                        Catch
                            countFail += 1
                        End Try
                    Next

                    '========== DIM Y (dọc) ==========
                    Dim yTargets = allEdges.Where(Function(e) e.IsHorizontal OrElse e.IsHole).ToList()
                    If vBaseOnBottom Then
                        yTargets = yTargets.OrderBy(Function(e) e.Y).ToList()
                    Else
                        yTargets = yTargets.OrderByDescending(Function(e) e.Y).ToList()
                    End If

                    Dim tierIndexY As Integer = 0
                    Dim tierStepY As Double = If(useTier, 0.4, 0.0)

                    For Each e As EdgeInfo In yTargets
                        If Math.Abs(e.Y - vBaseCoord) <= EDGE_TOL Then Continue For
                        If CurveHasDim(dimmedCurves, e.Curve) Then Continue For
                        If Math.Abs(e.Y - vBaseCoord) < minGapY Then Continue For
                        If ValueExists(usedY, e.Y) Then Continue For

                        Dim offsetX As Double = yNearGap + tierIndexY * tierStepY
                        If useTier Then tierIndexY += 1

                        Dim placeX As Double = If(vBaseOnBottom, maxX + offsetX, minX - offsetX)

                        Try
                            Dim baseIntent = oSheet.CreateGeometryIntent(vBaseEdge)
                            Dim edgeIntent = MakeIntent(oSheet, e)
                            Dim placement = tg.CreatePoint2d(placeX, e.Y)

                            Dim newDim = oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                placement, baseIntent, edgeIntent, DimensionTypeEnum.kVerticalDimensionType)

                            If newDim IsNot Nothing Then
                                usedY.Add(e.Y)
                                e.Dimmed = True
                                countOK += 1
                                dimmedCurves.Add(e.Curve)
                            End If
                        Catch
                            countFail += 1
                        End Try
                    Next
                Next

                ' Arrange
                Try
                    ArrangeDimensions(oSheet, app)
                Catch
                End Try

                oDrawDoc.Update()

                MessageBox.Show("Hoàn tất!" & vbCrLf & vbCrLf &
                    "View: " & selectedViews.Count & vbCrLf &
                    "Dim tạo: " & countOK & vbCrLf &
                    "Lỗi: " & countFail,
                    "Dim Baseline cạnh + lỗ", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Dim Baseline cạnh + lỗ",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        '===== Helper =====
        Private Function MakeIntent(oSheet As Sheet, e As EdgeInfo) As GeometryIntent
            If e Is Nothing OrElse e.Curve Is Nothing Then Return Nothing
            If e.IsHole Then
                Return oSheet.CreateGeometryIntent(e.Curve, PointIntentEnum.kCenterPointIntent)
            Else
                Return oSheet.CreateGeometryIntent(e.Curve)
            End If
        End Function

        Private Function EdgeInView(edge As DrawingCurve, view As DrawingView) As Boolean
            If edge Is Nothing OrElse view Is Nothing Then Return False
            Try
                For Each c As DrawingCurve In view.DrawingCurves
                    If c Is edge Then Return True
                Next
            Catch
            End Try
            Return False
        End Function

        '===== FORM (đã bỏ xa gần, thêm bật/tắt bậc) =====
        Private Function ShowDirectionDialog(
            ByRef hFromRight As Boolean, ByRef vFromBottom As Boolean,
            ByRef hPickedEdge As DrawingCurve, ByRef vPickedEdge As DrawingCurve,
            ByRef includeHoles As Boolean,
            ByRef xNearGap As Double, ByRef yNearGap As Double,
            ByRef minGapX As Double, ByRef minGapY As Double,
            ByRef useTier As Boolean,
            app As Inventor.Application) As Boolean

            Dim frm As New Form With {
                .Text = "Hướng chuẩn Baseline",
                .ClientSize = New Size(420, 480),
                .StartPosition = FormStartPosition.CenterScreen,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .MaximizeBox = False, .MinimizeBox = False, .ShowInTaskbar = False}

            Dim hHolder As New EdgePickHolder()
            Dim vHolder As New EdgePickHolder()
            Dim okClicked As Boolean = False

            ' Group NGANG
            Dim gbH As New GroupBox With {.Text = "Hướng chuẩn NGANG (X)", .Bounds = New Rectangle(12, 12, 396, 110)}
            Dim rbHLeft As New RadioButton With {.Text = "Từ TRÁI → Phải (mặc định)", .Bounds = New Rectangle(15, 22, 370, 20), .Checked = True}
            Dim rbHRight As New RadioButton With {.Text = "Từ PHẢI → Trái", .Bounds = New Rectangle(15, 46, 370, 20)}
            Dim lblH As New Label With {.Text = "(chưa pick cạnh)", .Bounds = New Rectangle(15, 74, 250, 20), .ForeColor = System.Drawing.Color.Gray, .Font = New Font("Segoe UI", 8, FontStyle.Italic)}
            Dim btnPickH As New Button With {.Text = "Pick cạnh X...", .Bounds = New Rectangle(270, 70, 105, 26)}
            gbH.Controls.AddRange({rbHLeft, rbHRight, lblH, btnPickH})
            frm.Controls.Add(gbH)

            ' Group DỌC
            Dim gbV As New GroupBox With {.Text = "Hướng chuẩn DỌC (Y)", .Bounds = New Rectangle(12, 130, 396, 110)}
            Dim rbVTop As New RadioButton With {.Text = "Từ TRÊN → Dưới (mặc định)", .Bounds = New Rectangle(15, 22, 370, 20), .Checked = True}
            Dim rbVBottom As New RadioButton With {.Text = "Từ DƯỚI → Trên", .Bounds = New Rectangle(15, 46, 370, 20)}
            Dim lblV As New Label With {.Text = "(chưa pick cạnh)", .Bounds = New Rectangle(15, 74, 250, 20), .ForeColor = System.Drawing.Color.Gray, .Font = New Font("Segoe UI", 8, FontStyle.Italic)}
            Dim btnPickV As New Button With {.Text = "Pick cạnh Y...", .Bounds = New Rectangle(270, 70, 105, 26)}
            gbV.Controls.AddRange({rbVTop, rbVBottom, lblV, btnPickV})
            frm.Controls.Add(gbV)

            ' Group KHOẢNG CÁCH
            Dim gbO As New GroupBox With {.Text = "Khoảng cách dim line (mm)", .Bounds = New Rectangle(12, 250, 396, 130)}
            Dim cbXNear As New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDown, .Bounds = New Rectangle(90, 22, 90, 22)}
            cbXNear.Items.AddRange({"2.0", "2.5", "3.2", "5.0", "8.0", "10.0"})
            cbXNear.Text = xNearGap.ToString("0.0#")

            Dim cbYNear As New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDown, .Bounds = New Rectangle(265, 22, 90, 22)}
            cbYNear.Items.AddRange({"2.0", "2.5", "3.2", "5.0", "8.0", "10.0"})
            cbYNear.Text = yNearGap.ToString("0.0#")

            Dim cbMinX As New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDown, .Bounds = New Rectangle(110, 55, 90, 22)}
            cbMinX.Items.AddRange({"0", "1", "2", "5", "10", "15", "20", "30"})
            cbMinX.Text = minGapX.ToString("0.0#")

            Dim cbMinY As New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDown, .Bounds = New Rectangle(300, 55, 85, 22)}
            cbMinY.Items.AddRange({"0", "1", "2", "5", "10", "15", "20", "30"})
            cbMinY.Text = minGapY.ToString("0.0#")

            Dim chkTier As New CheckBox With {
                .Text = "Bật khoảng cách bậc dim line (mỗi dim lệch 4 mm)",
                .Bounds = New Rectangle(15, 90, 370, 22),
                .Checked = useTier}

            gbO.Controls.AddRange({
                New Label With {.Text = "X gần:", .Bounds = New Rectangle(15, 24, 70, 22)},
                cbXNear,
                New Label With {.Text = "Y gần:", .Bounds = New Rectangle(200, 24, 60, 22)},
                cbYNear,
                New Label With {.Text = "Bỏ nếu X < :", .Bounds = New Rectangle(15, 57, 90, 22)},
                cbMinX,
                New Label With {.Text = "Bỏ nếu Y < :", .Bounds = New Rectangle(210, 57, 90, 22)},
                cbMinY,
                chkTier})
            frm.Controls.Add(gbO)

            Dim chkHoles As New CheckBox With {.Text = "Bao gồm LỖ TRÒN (đo theo tâm lỗ)", .Bounds = New Rectangle(15, 390, 396, 22), .Checked = True}
            frm.Controls.Add(chkHoles)

            Dim btnOK As New Button With {.Text = "OK", .Bounds = New Rectangle(215, 430, 85, 28)}
            Dim btnCancel As New Button With {.Text = "Hủy", .Bounds = New Rectangle(310, 430, 85, 28)}
            frm.Controls.AddRange({btnOK, btnCancel})

            AddHandler btnOK.Click, Sub()
                                        okClicked = True
                                        frm.DialogResult = DialogResult.OK
                                        frm.Close()
                                    End Sub
            AddHandler btnCancel.Click, Sub()
                                            okClicked = False
                                            frm.DialogResult = DialogResult.Cancel
                                            frm.Close()
                                        End Sub

            AddHandler btnPickH.Click, Sub()
                                           frm.Hide()
                                           System.Windows.Forms.Application.DoEvents()
                                           Try
                                               Dim seg = TryCast(app.CommandManager.Pick(
                                                   SelectionFilterEnum.kDrawingCurveSegmentFilter,
                                                   "Chọn cạnh ĐỨNG làm gốc đo X"), DrawingCurveSegment)
                                               If seg IsNot Nothing Then
                                                   Dim c = seg.Parent
                                                   Dim dx = Math.Abs(c.StartPoint.X - c.EndPoint.X)
                                                   Dim dy = Math.Abs(c.StartPoint.Y - c.EndPoint.Y)
                                                   If dx < EDGE_TOL AndAlso dy > EDGE_TOL Then
                                                       hHolder.Edge = c
                                                       lblH.Text = "✔ Đã pick cạnh đứng"
                                                       lblH.ForeColor = System.Drawing.Color.Green
                                                   Else
                                                       hHolder.Edge = Nothing
                                                       lblH.Text = "✘ Không phải cạnh ĐỨNG"
                                                       lblH.ForeColor = System.Drawing.Color.Red
                                                   End If
                                               End If
                                           Catch
                                           End Try
                                           frm.Show() : frm.Activate()
                                       End Sub

            AddHandler btnPickV.Click, Sub()
                                           frm.Hide()
                                           System.Windows.Forms.Application.DoEvents()
                                           Try
                                               Dim seg = TryCast(app.CommandManager.Pick(
                                                   SelectionFilterEnum.kDrawingCurveSegmentFilter,
                                                   "Chọn cạnh NGANG làm gốc đo Y"), DrawingCurveSegment)
                                               If seg IsNot Nothing Then
                                                   Dim c = seg.Parent
                                                   Dim dx = Math.Abs(c.StartPoint.X - c.EndPoint.X)
                                                   Dim dy = Math.Abs(c.StartPoint.Y - c.EndPoint.Y)
                                                   If dy < EDGE_TOL AndAlso dx > EDGE_TOL Then
                                                       vHolder.Edge = c
                                                       lblV.Text = "✔ Đã pick cạnh ngang"
                                                       lblV.ForeColor = System.Drawing.Color.Green
                                                   Else
                                                       vHolder.Edge = Nothing
                                                       lblV.Text = "✘ Không phải cạnh NGANG"
                                                       lblV.ForeColor = System.Drawing.Color.Red
                                                   End If
                                               End If
                                           Catch
                                           End Try
                                           frm.Show() : frm.Activate()
                                       End Sub

            If frm.ShowDialog() <> DialogResult.OK OrElse Not okClicked Then Return False

            ' Parse mm → cm
            Dim tmp As Double
            If Not Double.TryParse(cbXNear.Text.Replace(",", "."), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, tmp) OrElse tmp <= 0 Then tmp = 3.2
            xNearGap = tmp / 10.0

            If Not Double.TryParse(cbYNear.Text.Replace(",", "."), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, tmp) OrElse tmp <= 0 Then tmp = 3.2
            yNearGap = tmp / 10.0

            If Not Double.TryParse(cbMinX.Text.Replace(",", "."), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, tmp) OrElse tmp < 0 Then tmp = 5.0
            minGapX = tmp / 10.0

            If Not Double.TryParse(cbMinY.Text.Replace(",", "."), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, tmp) OrElse tmp < 0 Then tmp = 5.0
            minGapY = tmp / 10.0

            useTier = chkTier.Checked
            hFromRight = rbHRight.Checked
            vFromBottom = rbVBottom.Checked
            hPickedEdge = hHolder.Edge
            vPickedEdge = vHolder.Edge
            includeHoles = chkHoles.Checked

            Return True
        End Function

        Private Sub ArrangeDimensions(oSheet As Sheet, app As Inventor.Application)
            Try
                Dim oDims = oSheet.DrawingDimensions
                If oDims Is Nothing OrElse oDims.Count = 0 Then Exit Sub
                Dim oCol = app.TransientObjects.CreateObjectCollection
                For Each oDim As DrawingDimension In oDims
                    Try
                        If TypeOf oDim Is LinearGeneralDimension OrElse TypeOf oDim Is AngularGeneralDimension Then
                            Try : oDim.CenterText() : Catch : End Try
                            oCol.Add(oDim)
                        End If
                    Catch
                    End Try
                Next
                If oCol.Count > 1 Then oDims.Arrange(oCol)
            Catch
            End Try
        End Sub

        Private Function ValueExists(list As List(Of Double), value As Double) As Boolean
            If list Is Nothing Then Return False
            For Each v In list
                If Math.Abs(v - value) <= EDGE_TOL Then Return True
            Next
            Return False
        End Function

        Private Function SameSegment(c1 As DrawingCurve, c2 As DrawingCurve) As Boolean
            If c1 Is Nothing OrElse c2 Is Nothing Then Return False
            Try
                Dim a1 = c1.StartPoint, a2 = c1.EndPoint
                Dim b1 = c2.StartPoint, b2 = c2.EndPoint
                If a1 Is Nothing OrElse a2 Is Nothing OrElse b1 Is Nothing OrElse b2 Is Nothing Then Return False
                Return (Math.Abs(a1.X - b1.X) < EDGE_TOL AndAlso Math.Abs(a1.Y - b1.Y) < EDGE_TOL AndAlso
                        Math.Abs(a2.X - b2.X) < EDGE_TOL AndAlso Math.Abs(a2.Y - b2.Y) < EDGE_TOL) OrElse
                       (Math.Abs(a1.X - b2.X) < EDGE_TOL AndAlso Math.Abs(a1.Y - b2.Y) < EDGE_TOL AndAlso
                        Math.Abs(a2.X - b1.X) < EDGE_TOL AndAlso Math.Abs(a2.Y - b1.Y) < EDGE_TOL)
            Catch
            End Try
            Return False
        End Function

        Private Function SameCircle(c1 As DrawingCurve, c2 As DrawingCurve) As Boolean
            If c1 Is Nothing OrElse c2 Is Nothing Then Return False
            Try
                Dim p1 = c1.CenterPoint, p2 = c2.CenterPoint
                If p1 Is Nothing OrElse p2 Is Nothing Then Return False
                Return Math.Abs(p1.X - p2.X) <= EDGE_TOL AndAlso Math.Abs(p1.Y - p2.Y) <= EDGE_TOL
            Catch
            End Try
            Return False
        End Function

        Private Function CollectDimmedCurves(oSheet As Sheet) As List(Of DrawingCurve)
            Dim result As New List(Of DrawingCurve)
            Try
                For Each oDim As DrawingDimension In oSheet.DrawingDimensions
                    Try
                        Dim atts = oDim.AttachedEntities
                        If atts Is Nothing Then Continue For
                        For i = 1 To atts.Count
                            Dim ent = atts.Item(i)
                            If ent Is Nothing Then Continue For
                            Dim gi = TryCast(ent, GeometryIntent)
                            If gi IsNot Nothing Then
                                Dim dc = TryCast(gi.Geometry, DrawingCurve)
                                If dc IsNot Nothing Then result.Add(dc) : Continue For
                            End If
                            Dim dc2 = TryCast(ent, DrawingCurve)
                            If dc2 IsNot Nothing Then result.Add(dc2) : Continue For
                            Dim seg = TryCast(ent, DrawingCurveSegment)
                            If seg IsNot Nothing AndAlso seg.Parent IsNot Nothing Then result.Add(seg.Parent)
                        Next
                    Catch
                    End Try
                Next
            Catch
            End Try
            Return result
        End Function

        Private Function CurveHasDim(dimmed As List(Of DrawingCurve), c As DrawingCurve) As Boolean
            If dimmed Is Nothing OrElse c Is Nothing Then Return False
            For Each dc In dimmed
                If dc Is c Then Return True
                If SameSegment(dc, c) Then Return True
                If SameCircle(dc, c) Then Return True
            Next
            Return False
        End Function

        Public Class EdgeInfo
            Public Curve As DrawingCurve
            Public IsVertical As Boolean
            Public IsHorizontal As Boolean
            Public IsHole As Boolean
            Public X As Double
            Public Y As Double
            Public Dimmed As Boolean
        End Class

        Public Class EdgePickHolder
            Public Edge As DrawingCurve
        End Class

    End Module
End Namespace