Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq

Namespace ToolInventor2025.Drawing.Buttons.Drawdim
    Public Module Draw_dim_hole_b

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

                Dim countOK As Integer = 0
                Dim countFail As Integer = 0
                Dim countZero As Integer = 0

                Const TOL As Double = 0.0
                Const RATIO As Double = 0.55
                Const ARRAY_RADIUS As Double = 30.0

                '=====================================================
                ' CHỌN NHIỀU VIEW
                '=====================================================
                Dim selectedViews As New List(Of DrawingView)
                Do
                    Dim oSS As SelectSet = oDrawDoc.SelectSet
                    oSS.Clear()

                    Dim oView As DrawingView = Nothing
                    Try
                        oView = CType(
                            app.CommandManager.Pick(
                                SelectionFilterEnum.kDrawingViewFilter,
                                "Chọn View (Esc hoặc Right-click để kết thúc)"),
                            DrawingView)
                    Catch
                        Exit Do
                    End Try

                    If oView Is Nothing Then Exit Do

                    Dim already As Boolean = False
                    For Each v As DrawingView In selectedViews
                        If v Is oView Then
                            already = True
                            Exit For
                        End If
                    Next
                    If Not already Then selectedViews.Add(oView)
                Loop

                If selectedViews.Count = 0 Then
                    MessageBox.Show("Chưa chọn View nào.", "Thông báo")
                    Exit Sub
                End If

                '=====================================================
                ' XỬ LÝ TỪNG VIEW ĐÃ CHỌN
                '=====================================================
                For Each oView As DrawingView In selectedViews

                    '===== 1. Lấy đường tròn + lọc ARRAY =====
                    Dim allHoles As New List(Of HoleInfo)
                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType <> CurveTypeEnum.kCircleCurve Then Continue For
                            Dim c As Point2d = oCurve.CenterPoint
                            If c Is Nothing Then Continue For

                            Dim isArray As Boolean = False
                            Try
                                Dim modelGeom As Object = oCurve.ModelGeometry
                                If TypeOf modelGeom Is Edge Then
                                    Dim ed As Edge = CType(modelGeom, Edge)
                                    For Each fc As Face In ed.Faces
                                        Dim feat As PartFeature = fc.CreatedByFeature
                                        If feat IsNot Nothing Then
                                            Dim tName As String = TypeName(feat).ToUpperInvariant()
                                            If tName.Contains("CIRCULARPATTERN") OrElse tName.Contains("PATTERN") Then
                                                isArray = True
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If
                            Catch
                            End Try

                            If isArray Then Continue For

                            Dim hi As New HoleInfo
                            hi.Curve = oCurve
                            hi.Center = c
                            hi.X = c.X
                            hi.Y = c.Y
                            Try
                                hi.Radius = oCurve.Radius
                            Catch
                                hi.Radius = 1.5
                            End Try
                            allHoles.Add(hi)
                        Catch
                        End Try
                    Next

                    Dim cleanHoles As New List(Of HoleInfo)
                    For Each h In allHoles
                        Dim nearby = allHoles.Where(Function(o) o IsNot h AndAlso
                            Math.Sqrt((o.X - h.X) ^ 2 + (o.Y - h.Y) ^ 2) < ARRAY_RADIUS).Count()
                        If nearby >= 3 Then Continue For
                        cleanHoles.Add(h)
                    Next
                    If cleanHoles.Count = 0 Then cleanHoles = allHoles
                    If cleanHoles.Count = 0 Then Continue For

                    '===== 2. Tìm 4 cạnh =====
                    Dim leftEdge As DrawingCurve = Nothing
                    Dim rightEdge As DrawingCurve = Nothing
                    Dim topEdge As DrawingCurve = Nothing
                    Dim bottomEdge As DrawingCurve = Nothing
                    Dim minX As Double = Double.MaxValue
                    Dim maxX As Double = Double.MinValue
                    Dim maxY As Double = Double.MinValue
                    Dim minY As Double = Double.MaxValue

                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType <> CurveTypeEnum.kLineCurve AndAlso
                               oCurve.CurveType <> CurveTypeEnum.kLineSegmentCurve Then Continue For

                            Dim p1 = oCurve.StartPoint
                            Dim p2 = oCurve.EndPoint

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

                    Dim viewW = maxX - minX
                    Dim viewH = maxY - minY
                    Dim limitX = minX + viewW * RATIO
                    Dim limitY = maxY - viewH * RATIO

                    '===== 3. Dim tổng =====
                    If leftEdge IsNot Nothing AndAlso rightEdge IsNot Nothing Then
                        Try
                            Dim tp = tg.CreatePoint2d((minX + maxX) / 2, maxY + 5.5)
                            oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                tp,
                                oSheet.CreateGeometryIntent(leftEdge),
                                oSheet.CreateGeometryIntent(rightEdge),
                                DimensionTypeEnum.kHorizontalDimensionType)
                            countOK += 1
                        Catch
                            countFail += 1
                        End Try
                    End If

                    If topEdge IsNot Nothing AndAlso bottomEdge IsNot Nothing Then
                        Try
                            Dim tp = tg.CreatePoint2d(minX - 5.5, (maxY + minY) / 2)
                            oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                tp,
                                oSheet.CreateGeometryIntent(topEdge),
                                oSheet.CreateGeometryIntent(bottomEdge),
                                DimensionTypeEnum.kVerticalDimensionType)
                            countOK += 1
                        Catch
                            countFail += 1
                        End Try
                    End If

                    '=====================================================
                    ' b1: Trái → Phải (phía trên)
                    '=====================================================
                    Dim topRow = cleanHoles.Where(Function(h) h.Y >= limitY).OrderBy(Function(h) h.X).ToList()
                    If topRow.Count > 0 AndAlso leftEdge IsNot Nothing Then
                        Dim prev = topRow.First()
                        If (prev.X - minX) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d((minX + prev.X) / 2, maxY + 3.2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    tp,
                                    oSheet.CreateGeometryIntent(leftEdge),
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                            Catch
                                countFail += 1
                            End Try
                        End If

                        For i = 1 To topRow.Count - 1
                            Dim h = topRow(i)
                            Dim dist = h.X - prev.X
                            If dist < TOL Then
                                countZero += 1
                                Continue For
                            End If
                            Try
                                Dim tp = tg.CreatePoint2d((prev.X + h.X) / 2, maxY + 2.5)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(h.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                                prev = h
                            Catch
                                countFail += 1
                            End Try
                        Next

                        If rightEdge IsNot Nothing AndAlso (maxX - prev.X) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d((prev.X + maxX) / 2, maxY + 3.2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(rightEdge),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                            Catch
                                countFail += 1
                            End Try
                        End If
                    End If

                    '=====================================================
                    ' b2: Trên → Dưới (bên trái)
                    '=====================================================
                    Dim leftCol = cleanHoles.Where(Function(h) h.X <= limitX) _
                                            .OrderByDescending(Function(h) h.Y).ToList()
                    Dim uniqueLeft As New List(Of HoleInfo)
                    Dim lastY As Double = Double.MaxValue
                    For Each h In leftCol
                        If Math.Abs(h.Y - lastY) < TOL Then Continue For
                        uniqueLeft.Add(h)
                        lastY = h.Y
                    Next

                    If uniqueLeft.Count > 0 AndAlso topEdge IsNot Nothing Then
                        Dim prev = uniqueLeft.First()
                        If (maxY - prev.Y) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d(minX - 3.2, (maxY + prev.Y) / 2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    tp,
                                    oSheet.CreateGeometryIntent(topEdge),
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                            Catch
                                countFail += 1
                            End Try
                        End If

                        For i = 1 To uniqueLeft.Count - 1
                            Dim h = uniqueLeft(i)
                            Dim dist = prev.Y - h.Y
                            If dist < TOL Then
                                countZero += 1
                                Continue For
                            End If
                            Try
                                Dim tp = tg.CreatePoint2d(minX - 2.5, (prev.Y + h.Y) / 2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(h.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                                prev = h
                            Catch
                                countFail += 1
                            End Try
                        Next

                        If bottomEdge IsNot Nothing AndAlso (prev.Y - minY) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d(minX - 3.2, (prev.Y + minY) / 2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(bottomEdge),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                            Catch
                                countFail += 1
                            End Try
                        End If
                    End If

                    '=====================================================
                    ' b3: Trái → Phải (phía dưới)
                    '=====================================================
                    Dim botRow = cleanHoles.Where(Function(h) h.Y < limitY).OrderBy(Function(h) h.X).ToList()
                    If botRow.Count > 0 AndAlso leftEdge IsNot Nothing Then
                        Dim prev = botRow.First()
                        If (prev.X - minX) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d((minX + prev.X) / 2, minY - 3.2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    tp,
                                    oSheet.CreateGeometryIntent(leftEdge),
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                            Catch
                                countFail += 1
                            End Try
                        End If

                        For i = 1 To botRow.Count - 1
                            Dim h = botRow(i)
                            Dim dist = h.X - prev.X
                            If dist < TOL Then
                                countZero += 1
                                Continue For
                            End If
                            Try
                                Dim tp = tg.CreatePoint2d((prev.X + h.X) / 2, minY - 2.5)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(h.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                                prev = h
                            Catch
                                countFail += 1
                            End Try
                        Next

                        If rightEdge IsNot Nothing AndAlso (maxX - prev.X) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d((prev.X + maxX) / 2, minY - 3.2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(rightEdge),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                            Catch
                                countFail += 1
                            End Try
                        End If
                    End If

                    '=====================================================
                    ' b4: Trên → Dưới (bên phải)
                    '=====================================================
                    Dim rightCol = cleanHoles.Where(Function(h) h.X > limitX) _
                                             .OrderByDescending(Function(h) h.Y).ToList()
                    If rightCol.Count > 0 AndAlso topEdge IsNot Nothing Then
                        Dim prev = rightCol.First()
                        If (maxY - prev.Y) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d(maxX + 3.2, (maxY + prev.Y) / 2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    tp,
                                    oSheet.CreateGeometryIntent(topEdge),
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                            Catch
                                countFail += 1
                            End Try
                        End If

                        For i = 1 To rightCol.Count - 1
                            Dim h = rightCol(i)
                            Dim dist = prev.Y - h.Y
                            If dist < TOL Then
                                countZero += 1
                                Continue For
                            End If
                            Try
                                Dim tp = tg.CreatePoint2d(maxX + 2.5, (prev.Y + h.Y) / 2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(h.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                                prev = h
                            Catch
                                countFail += 1
                            End Try
                        Next

                        If bottomEdge IsNot Nothing AndAlso (prev.Y - minY) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d(maxX + 3.2, (prev.Y + minY) / 2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(bottomEdge),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                            Catch
                                countFail += 1
                            End Try
                        End If
                    End If
                Next

                '===== ARRANGE — CHỈ DIM THUỘC VIEW ĐÃ CHỌN =====
                Try
                    ArrangeDimensions(oSheet, app, selectedViews)
                Catch
                End Try

                oDrawDoc.Update()

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Số view đã xử lý: " & selectedViews.Count & vbCrLf &
                    "Số dim: " & countOK & vbCrLf &
                    "Bỏ = 0: " & countZero & vbCrLf &
                    "Lỗi: " & countFail & vbCrLf & vbCrLf &
                    "• Chọn nhiều View liên tục" & vbCrLf &
                    "• Lọc Array bằng Feature + mật độ" & vbCrLf &
                    "• Auto Arrange Dimension (chỉ view đã chọn)",
                    "Dim Chain lỗ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "Dim Chain lỗ",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)
            End Try
        End Sub

        '=============================================================
        ' ARRANGE — CHỈ DIM THUỘC VIEW ĐÃ CHỌN
        '=============================================================
        Private Sub ArrangeDimensions(ByVal oSheet As Sheet,
                                      ByVal app As Inventor.Application,
                                      ByVal selectedViews As List(Of DrawingView))
            Try
                Dim oDims As DrawingDimensions = oSheet.DrawingDimensions
                If oDims Is Nothing OrElse oDims.Count = 0 Then Exit Sub

                Dim oCol As ObjectCollection = app.TransientObjects.CreateObjectCollection

                For Each oDim As DrawingDimension In oDims
                    Try
                        If Not IsDimInAnyView(oDim, selectedViews) Then Continue For

                        If TypeOf oDim Is LinearGeneralDimension OrElse
                           TypeOf oDim Is AngularGeneralDimension Then
                            Try
                                oDim.CenterText()
                            Catch
                            End Try
                            oCol.Add(oDim)
                        End If
                    Catch
                    End Try
                Next

                If oCol.Count > 0 Then
                    oDims.Arrange(oCol)
                End If
            Catch
            End Try
        End Sub

        '=============================================================
        ' LỌC DIM THEO VIEW ĐÃ CHỌN (Inventor 2020)
        '=============================================================
        Private Function IsDimInAnyView(ByVal oDim As DrawingDimension,
                                         ByVal views As List(Of DrawingView)) As Boolean
            Try
                Dim linDim As LinearGeneralDimension = TryCast(oDim, LinearGeneralDimension)
                If linDim IsNot Nothing Then
                    If CheckIntentBelongsToView(linDim.IntentOne, views) Then Return True
                    If CheckIntentBelongsToView(linDim.IntentTwo, views) Then Return True
                End If
            Catch
            End Try

            Try
                Dim ent As Object = Nothing
                Try
                    ent = oDim.AttachedEntity
                Catch
                End Try

                If ent IsNot Nothing Then
                    Dim parentView As DrawingView = TryCast(GetParentView(ent), DrawingView)
                    If parentView IsNot Nothing Then
                        For Each v As DrawingView In views
                            If v Is parentView Then Return True
                        Next
                    End If
                End If
            Catch
            End Try

            Try
                Dim tp As Point2d = Nothing
                Try
                    tp = oDim.Text.Origin
                Catch
                    Try
                        tp = oDim.Text.Position
                    Catch
                        Return False
                    End Try
                End Try

                If tp Is Nothing Then Return False

                Const boxTol As Double = 0.5

                For Each v As DrawingView In views
                    Try
                        Dim cx As Double = v.Position.X
                        Dim cy As Double = v.Position.Y
                        Dim hw As Double = v.Width / 2.0
                        Dim hh As Double = v.Height / 2.0

                        Dim vL As Double = cx - hw
                        Dim vR As Double = cx + hw
                        Dim vB As Double = cy - hh
                        Dim vT As Double = cy + hh

                        If tp.X >= (vL - boxTol) AndAlso tp.X <= (vR + boxTol) AndAlso
                           tp.Y >= (vB - boxTol) AndAlso tp.Y <= (vT + boxTol) Then
                            Return True
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            Return False
        End Function

        Private Function CheckIntentBelongsToView(ByVal intent As Object,
                                                   ByVal views As List(Of DrawingView)) As Boolean
            If intent Is Nothing Then Return False
            Try
                Dim gi As GeometryIntent = TryCast(intent, GeometryIntent)
                If gi Is Nothing Then Return False

                Dim geom As Object = Nothing
                Try
                    geom = gi.Geometry
                Catch
                    Return False
                End Try

                If geom Is Nothing Then Return False

                Dim parentView As DrawingView = TryCast(GetParentView(geom), DrawingView)
                If parentView Is Nothing Then Return False

                For Each v As DrawingView In views
                    If v Is parentView Then Return True
                Next
            Catch
            End Try
            Return False
        End Function

        Private Function GetParentView(ByVal obj As Object) As DrawingView
            If obj Is Nothing Then Return Nothing
            Try
                Dim p As Object = Nothing
                Try
                    p = obj.Parent
                Catch
                End Try

                Dim dv As DrawingView = TryCast(p, DrawingView)
                If dv IsNot Nothing Then Return dv

                Try
                    If p IsNot Nothing Then
                        Dim p2 As Object = p.Parent
                        dv = TryCast(p2, DrawingView)
                        If dv IsNot Nothing Then Return dv
                    End If
                Catch
                End Try
            Catch
            End Try
            Return Nothing
        End Function

        Public Class HoleInfo
            Public Curve As DrawingCurve
            Public Center As Point2d
            Public X As Double
            Public Y As Double
            Public Radius As Double
        End Class

    End Module
End Namespace