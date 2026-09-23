Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq

Namespace ToolInventor2025.Drawing.Buttons.Drawdim
    Public Module Draw_dim_hole_a

        Private Const TOL As Double = 0.0
        Private Const EDGE_TOL As Double = 0.15
        Private Const RADIUS_TOL As Double = 0.04

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim app As Inventor.Application = g_inventorApplication

            Try
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Vui lòng mở file Drawing (.idw)!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument = CType(app.ActiveDocument, DrawingDocument)
                Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                Dim tg As TransientGeometry = app.TransientGeometry

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

                Dim countOK As Integer = 0
                Dim countFail As Integer = 0
                Dim countZero As Integer = 0
                Dim countArray As Integer = 0

                '=====================================================
                ' XỬ LÝ TỪNG VIEW ĐÃ CHỌN
                '=====================================================
                For Each oView As DrawingView In selectedViews

                    Dim allHoles As New List(Of HoleInfo)
                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType <> CurveTypeEnum.kCircleCurve Then Continue For
                            Dim c As Point2d = oCurve.CenterPoint
                            If c Is Nothing Then Continue For

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

                    If allHoles.Count = 0 Then Continue For

                    Dim arrayHoles As List(Of HoleInfo) = FindCircularArrayHoles(allHoles)

                    Dim cleanHoles As New List(Of HoleInfo)
                    For Each h As HoleInfo In allHoles
                        If ContainsHole(arrayHoles, h) Then
                            countArray += 1
                        Else
                            cleanHoles.Add(h)
                        End If
                    Next

                    If cleanHoles.Count = 0 Then Continue For

                    '-------------------------------------------------
                    ' 4. TÌM BIÊN VIEW
                    '-------------------------------------------------
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
                            Select Case oCurve.CurveType
                                Case CurveTypeEnum.kLineCurve, CurveTypeEnum.kLineSegmentCurve
                                    Dim p1 As Point2d = oCurve.StartPoint
                                    Dim p2 As Point2d = oCurve.EndPoint
                                    minX = Math.Min(minX, Math.Min(p1.X, p2.X))
                                    maxX = Math.Max(maxX, Math.Max(p1.X, p2.X))
                                    minY = Math.Min(minY, Math.Min(p1.Y, p2.Y))
                                    maxY = Math.Max(maxY, Math.Max(p1.Y, p2.Y))
                                Case CurveTypeEnum.kCircleCurve
                                    Dim c As Point2d = oCurve.CenterPoint
                                    If c Is Nothing Then Continue For
                                    Dim r As Double = 1.5
                                    Try : r = oCurve.Radius : Catch : End Try
                                    minX = Math.Min(minX, c.X - r)
                                    maxX = Math.Max(maxX, c.X + r)
                                    minY = Math.Min(minY, c.Y - r)
                                    maxY = Math.Max(maxY, c.Y + r)
                            End Select
                        Catch
                        End Try
                    Next

                    Dim bestLeftDist As Double = Double.MaxValue
                    Dim bestRightDist As Double = Double.MaxValue
                    Dim bestTopDist As Double = Double.MaxValue
                    Dim bestBotDist As Double = Double.MaxValue

                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType <> CurveTypeEnum.kLineCurve AndAlso
                               oCurve.CurveType <> CurveTypeEnum.kLineSegmentCurve Then Continue For

                            Dim p1 As Point2d = oCurve.StartPoint
                            Dim p2 As Point2d = oCurve.EndPoint
                            Dim len As Double = Math.Sqrt((p2.X - p1.X) ^ 2 + (p2.Y - p1.Y) ^ 2)
                            If len < 2.0 Then Continue For

                            If Math.Abs(p1.X - p2.X) < 0.05 Then
                                Dim distL As Double = Math.Abs(p1.X - minX)
                                Dim distR As Double = Math.Abs(p1.X - maxX)
                                If distL < bestLeftDist Then
                                    bestLeftDist = distL
                                    leftEdge = oCurve
                                End If
                                If distR < bestRightDist Then
                                    bestRightDist = distR
                                    rightEdge = oCurve
                                End If
                            End If

                            If Math.Abs(p1.Y - p2.Y) < 0.05 Then
                                Dim distT As Double = Math.Abs(p1.Y - maxY)
                                Dim distB As Double = Math.Abs(p1.Y - minY)
                                If distT < bestTopDist Then
                                    bestTopDist = distT
                                    topEdge = oCurve
                                End If
                                If distB < bestBotDist Then
                                    bestBotDist = distB
                                    bottomEdge = oCurve
                                End If
                            End If
                        Catch
                        End Try
                    Next

                    If leftEdge Is Nothing OrElse rightEdge Is Nothing OrElse
                       topEdge Is Nothing OrElse bottomEdge Is Nothing Then
                        Continue For
                    End If

                    Dim viewW As Double = maxX - minX
                    Dim viewH As Double = maxY - minY
                    Dim RATIO As Double = 0.48
                    Dim limitX As Double = minX + viewW * RATIO
                    Dim limitY As Double = maxY - viewH * RATIO

                    '-------------------------------------------------
                    ' 5. DIM TỔNG
                    '-------------------------------------------------
                    Try
                        Dim tp As Point2d = tg.CreatePoint2d((minX + maxX) / 2, maxY + 5.5)
                        oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                            tp,
                            oSheet.CreateGeometryIntent(leftEdge),
                            oSheet.CreateGeometryIntent(rightEdge),
                            DimensionTypeEnum.kHorizontalDimensionType)
                        countOK += 1
                    Catch
                        countFail += 1
                    End Try

                    Try
                        Dim tp As Point2d = tg.CreatePoint2d(minX - 5.5, (maxY + minY) / 2)
                        oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                            tp,
                            oSheet.CreateGeometryIntent(topEdge),
                            oSheet.CreateGeometryIntent(bottomEdge),
                            DimensionTypeEnum.kVerticalDimensionType)
                        countOK += 1
                    Catch
                        countFail += 1
                    End Try

                    '-------------------------------------------------
                    ' 6. CLUSTER + DIM 4 CHUỖI
                    '-------------------------------------------------
                    Dim rows As List(Of List(Of HoleInfo)) = ClusterByY(cleanHoles, 3.0)
                    Dim cols As List(Of List(Of HoleInfo)) = ClusterByX(cleanHoles, 3.0)

                    Dim topRow As List(Of HoleInfo) = Nothing
                    Dim botRow As List(Of HoleInfo) = Nothing
                    If rows.Count > 0 Then
                        topRow = rows(0).OrderBy(Function(h) h.X).ToList()
                        botRow = rows(rows.Count - 1).OrderBy(Function(h) h.X).ToList()
                    End If

                    Dim leftCol As List(Of HoleInfo) = Nothing
                    Dim rightCol As List(Of HoleInfo) = Nothing
                    If cols.Count > 0 Then
                        leftCol = cols(0).OrderByDescending(Function(h) h.Y).ToList()
                        rightCol = cols(cols.Count - 1).OrderByDescending(Function(h) h.Y).ToList()
                    End If

                    If topRow IsNot Nothing AndAlso topRow.Count > 0 Then
                        DimChainHorizontal(oSheet, tg, topRow, leftEdge, rightEdge, minX, maxX, maxY + 4.5, maxY + 3.0, True, countOK, countFail, countZero)
                    End If
                    If botRow IsNot Nothing AndAlso botRow.Count > 0 Then
                        DimChainHorizontal(oSheet, tg, botRow, leftEdge, rightEdge, minX, maxX, minY - 4.5, minY - 3.0, False, countOK, countFail, countZero)
                    End If
                    If leftCol IsNot Nothing AndAlso leftCol.Count > 0 Then
                        DimChainVertical(oSheet, tg, leftCol, topEdge, bottomEdge, maxY, minY, minX - 4.5, minX - 3.0, True, countOK, countFail, countZero)
                    End If
                    If rightCol IsNot Nothing AndAlso rightCol.Count > 0 Then
                        DimChainVertical(oSheet, tg, rightCol, topEdge, bottomEdge, maxY, minY, maxX + 4.5, maxX + 3.0, False, countOK, countFail, countZero)
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
                    "Số DIM: " & countOK & vbCrLf &
                    "Lỗ Circular Array bỏ qua: " & countArray & vbCrLf &
                    "Bỏ = 0: " & countZero & vbCrLf &
                    "Lỗi: " & countFail & vbCrLf & vbCrLf &
                    "• Chọn nhiều View liên tục" & vbCrLf &
                    "• Lọc Circular Array" & vbCrLf &
                    "• Auto Arrange Dimension (chỉ view đã chọn)",
                    "Dim Chain lỗ", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Dim Chain lỗ", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

        '=====================================================
        ' HELPER: DIM NGANG
        '=====================================================
        Private Sub DimChainHorizontal(
            oSheet As Sheet, tg As TransientGeometry,
            holes As List(Of HoleInfo),
            leftEdge As DrawingCurve, rightEdge As DrawingCurve,
            minX As Double, maxX As Double,
            edgeOffset As Double, midOffset As Double,
            isTop As Boolean,
            ByRef countOK As Integer, ByRef countFail As Integer, ByRef countZero As Integer)

            If holes.Count = 0 OrElse leftEdge Is Nothing Then Exit Sub

            Dim prev As HoleInfo = holes.First()

            If (prev.X - minX) > TOL Then
                Try
                    Dim tp As Point2d = tg.CreatePoint2d((minX + prev.X) / 2, edgeOffset)
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

            For i As Integer = 1 To holes.Count - 1
                Dim h As HoleInfo = holes(i)
                Dim dist As Double = h.X - prev.X
                If dist < TOL Then
                    countZero += 1
                    Continue For
                End If
                Try
                    Dim tp As Point2d = tg.CreatePoint2d((prev.X + h.X) / 2, midOffset)
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
                    Dim tp As Point2d = tg.CreatePoint2d((prev.X + maxX) / 2, edgeOffset)
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
        End Sub

        '=====================================================
        ' HELPER: DIM DỌC
        '=====================================================
        Private Sub DimChainVertical(
            oSheet As Sheet, tg As TransientGeometry,
            holes As List(Of HoleInfo),
            topEdge As DrawingCurve, bottomEdge As DrawingCurve,
            maxY As Double, minY As Double,
            edgeOffset As Double, midOffset As Double,
            isLeft As Boolean,
            ByRef countOK As Integer, ByRef countFail As Integer, ByRef countZero As Integer)

            If holes.Count = 0 OrElse topEdge Is Nothing Then Exit Sub

            Dim prev As HoleInfo = holes.First()

            If (maxY - prev.Y) > EDGE_TOL Then
                Try
                    Dim tp As Point2d = tg.CreatePoint2d(edgeOffset, (maxY + prev.Y) / 2)
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

            For i As Integer = 1 To holes.Count - 1
                Dim h As HoleInfo = holes(i)
                Dim dist As Double = prev.Y - h.Y
                If dist < TOL Then
                    countZero += 1
                    Continue For
                End If
                Try
                    Dim tp As Point2d = tg.CreatePoint2d(midOffset, (prev.Y + h.Y) / 2)
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

            If bottomEdge IsNot Nothing AndAlso (prev.Y - minY) > EDGE_TOL Then
                Try
                    Dim tp As Point2d = tg.CreatePoint2d(edgeOffset, (prev.Y + minY) / 2)
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
        End Sub

        '=====================================================
        ' TÌM CIRCULAR ARRAY
        '=====================================================
        Private Function FindCircularArrayHoles(allHoles As List(Of HoleInfo)) As List(Of HoleInfo)
            Dim result As New List(Of HoleInfo)
            If allHoles Is Nothing OrElse allHoles.Count < 4 Then Return result

            Dim checked As New List(Of HoleInfo)

            For Each baseHole As HoleInfo In allHoles
                If ContainsHole(checked, baseHole) Then Continue For

                Dim sameSize As New List(Of HoleInfo)
                For Each h As HoleInfo In allHoles
                    If Math.Abs(h.Radius - baseHole.Radius) < RADIUS_TOL Then
                        sameSize.Add(h)
                    End If
                Next

                For Each h As HoleInfo In sameSize
                    If Not ContainsHole(checked, h) Then checked.Add(h)
                Next

                If sameSize.Count < 4 Then Continue For

                Dim found As Boolean = False
                For i As Integer = 0 To sameSize.Count - 3
                    If found Then Exit For
                    For j As Integer = i + 1 To sameSize.Count - 2
                        If found Then Exit For
                        For k As Integer = j + 1 To sameSize.Count - 1
                            Dim h1 = sameSize(i), h2 = sameSize(j), h3 = sameSize(k)
                            Dim cx As Double = 0, cy As Double = 0
                            If Not GetCircleCenter(h1.X, h1.Y, h2.X, h2.Y, h3.X, h3.Y, cx, cy) Then Continue For

                            Dim arrayR As Double = GetDistance(h1.X, h1.Y, cx, cy)
                            If arrayR < 0.8 Then Continue For

                            Dim members As New List(Of HoleInfo)
                            Dim circleTol As Double = Math.Max(0.1, arrayR * 0.015)

                            For Each h As HoleInfo In sameSize
                                Dim d As Double = GetDistance(h.X, h.Y, cx, cy)
                                If Math.Abs(d - arrayR) <= circleTol Then
                                    members.Add(h)
                                End If
                            Next

                            If members.Count >= 4 AndAlso IsCircularDistribution(members, cx, cy) Then
                                For Each h As HoleInfo In members
                                    If Not ContainsHole(result, h) Then result.Add(h)
                                Next
                                found = True
                                Exit For
                            End If
                        Next
                    Next
                Next
            Next
            Return result
        End Function

        Private Function IsCircularDistribution(holes As List(Of HoleInfo), cx As Double, cy As Double) As Boolean
            If holes.Count < 4 Then Return False
            Dim angles As New List(Of Double)
            For Each h As HoleInfo In holes
                angles.Add(Math.Atan2(h.Y - cy, h.X - cx))
            Next
            angles.Sort()
            Dim unique As Integer = 1
            For i As Integer = 1 To angles.Count - 1
                If Math.Abs(angles(i) - angles(i - 1)) > 0.03 Then unique += 1
            Next
            If unique < 4 Then Return False
            Dim totalSpan As Double = angles.Last() - angles.First()
            If totalSpan < 1.5 Then Return False
            Return True
        End Function

        Private Function GetCircleCenter(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double,
                                         ByRef cx As Double, ByRef cy As Double) As Boolean
            Dim d As Double = 2.0 * (x1 * (y2 - y3) + x2 * (y3 - y1) + x3 * (y1 - y2))
            If Math.Abs(d) < 0.000001 Then Return False
            Dim x1sq As Double = x1 * x1 + y1 * y1
            Dim x2sq As Double = x2 * x2 + y2 * y2
            Dim x3sq As Double = x3 * x3 + y3 * y3
            cx = (x1sq * (y2 - y3) + x2sq * (y3 - y1) + x3sq * (y1 - y2)) / d
            cy = (x1sq * (x3 - x2) + x2sq * (x1 - x3) + x3sq * (x2 - x1)) / d
            Return True
        End Function

        Private Function ContainsHole(list As List(Of HoleInfo), target As HoleInfo) As Boolean
            For Each h As HoleInfo In list
                If h Is target Then Return True
            Next
            Return False
        End Function

        Private Function GetDistance(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
            Return Math.Sqrt((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
        End Function

        Private Function ClusterByY(holes As List(Of HoleInfo), tol As Double) As List(Of List(Of HoleInfo))
            Dim result As New List(Of List(Of HoleInfo))
            If holes Is Nothing OrElse holes.Count = 0 Then Return result
            Dim sorted = holes.OrderByDescending(Function(h) h.Y).ToList()
            Dim current As New List(Of HoleInfo)
            current.Add(sorted(0))
            For i As Integer = 1 To sorted.Count - 1
                If Math.Abs(sorted(i).Y - current(0).Y) <= tol Then
                    current.Add(sorted(i))
                Else
                    result.Add(current.OrderBy(Function(h) h.X).ToList())
                    current = New List(Of HoleInfo)
                    current.Add(sorted(i))
                End If
            Next
            result.Add(current.OrderBy(Function(h) h.X).ToList())
            Return result
        End Function

        Private Function ClusterByX(holes As List(Of HoleInfo), tol As Double) As List(Of List(Of HoleInfo))
            Dim result As New List(Of List(Of HoleInfo))
            If holes Is Nothing OrElse holes.Count = 0 Then Return result
            Dim sorted = holes.OrderBy(Function(h) h.X).ToList()
            Dim current As New List(Of HoleInfo)
            current.Add(sorted(0))
            For i As Integer = 1 To sorted.Count - 1
                If Math.Abs(sorted(i).X - current(0).X) <= tol Then
                    current.Add(sorted(i))
                Else
                    result.Add(current.OrderByDescending(Function(h) h.Y).ToList())
                    current = New List(Of HoleInfo)
                    current.Add(sorted(i))
                End If
            Next
            result.Add(current.OrderByDescending(Function(h) h.Y).ToList())
            Return result
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