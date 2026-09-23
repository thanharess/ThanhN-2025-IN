Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq

Namespace ToolInventor2025.Drawing.Buttons.Drawdim
    Public Module Draw_2d

        Private Const TOL As Double = 0.0
        Private Const RATIO As Double = 0.5

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then

                    MessageBox.Show(
                        "Vui lòng mở file Drawing (.idw)!",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument =
                    CType(app.ActiveDocument, DrawingDocument)

                Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                Dim tg As TransientGeometry = app.TransientGeometry

                '=====================================================
                ' CHỌN NHIỀU VIEW SAU KHI CHẠY CODE
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

                    ' Tránh chọn trùng
                    Dim already As Boolean = False
                    For Each v As DrawingView In selectedViews
                        If v Is oView Then
                            already = True
                            Exit For
                        End If
                    Next

                    If Not already Then
                        selectedViews.Add(oView)
                    End If

                Loop

                If selectedViews.Count = 0 Then
                    MessageBox.Show("Chưa chọn View nào.", "Thông báo")
                    Exit Sub
                End If

                Dim countOK As Integer = 0
                Dim countFail As Integer = 0
                Dim countAdd As Integer = 0
                Dim countZero As Integer = 0

                '=====================================================
                ' XỬ LÝ TỪNG VIEW ĐÃ CHỌN
                '=====================================================
                For Each oView As DrawingView In selectedViews

                    '-------------------------------------------------
                    ' 1. LẤY TẤT CẢ LỖ TRÒN (chỉ loại trùng TÂM)
                    '-------------------------------------------------
                    Dim allHoles As New List(Of HoleInfo)

                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType <> CurveTypeEnum.kCircleCurve Then Continue For

                            Dim c As Point2d = oCurve.CenterPoint
                            If c Is Nothing Then Continue For

                            Dim duplicated As Boolean = False
                            For Each oldHole As HoleInfo In allHoles
                                If Math.Abs(oldHole.X - c.X) <= TOL AndAlso
                                   Math.Abs(oldHole.Y - c.Y) <= TOL Then
                                    duplicated = True
                                    Exit For
                                End If
                            Next

                            If duplicated Then Continue For

                            Dim hi As New HoleInfo
                            hi.Curve = oCurve
                            hi.Center = c
                            hi.X = c.X
                            hi.Y = c.Y
                            hi.Dimmed = False
                            allHoles.Add(hi)

                        Catch
                        End Try
                    Next

                    If allHoles.Count = 0 Then Continue For

                    '-------------------------------------------------
                    ' 2. TÌM 4 CẠNH NGOÀI
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
                        Continue For
                    End If

                    Dim viewW As Double = maxX - minX
                    Dim viewH As Double = maxY - minY
                    If viewW <= 0 OrElse viewH <= 0 Then Continue For

                    Dim limitX As Double = minX + viewW * RATIO
                    Dim limitY As Double = maxY - viewH * RATIO

                    '-------------------------------------------------
                    ' 3. DIM TỔNG
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
                    ' 4. B1 - HÀNG TRÊN (BASE = CẠNH TRÁI)
                    '-------------------------------------------------
                    Dim topRow As List(Of HoleInfo) =
                        allHoles.Where(Function(h) h.Y >= limitY).
                        OrderBy(Function(h) h.X).ToList()

                    If topRow.Count > 0 Then
                        Dim result As Boolean =
                            AddBaselineSet(oSheet, app, leftEdge, topRow, True, maxY + 3.2)

                        If result Then
                            For Each h As HoleInfo In topRow
                                h.Dimmed = True
                            Next
                            countOK += topRow.Count
                        Else
                            countFail += 1
                        End If
                    End If

                    '-------------------------------------------------
                    ' 5. B2 - CỘT TRÁI (BASE = CẠNH TRÊN)
                    '-------------------------------------------------
                    Dim leftCol As List(Of HoleInfo) =
                        allHoles.Where(Function(h) h.X <= limitX).
                        OrderByDescending(Function(h) h.Y).ToList()

                    If leftCol.Count > 0 Then
                        Dim result As Boolean =
                            AddBaselineSet(oSheet, app, topEdge, leftCol, False, minX - 3.2)

                        If result Then
                            For Each h As HoleInfo In leftCol
                                h.Dimmed = True
                            Next
                            countOK += leftCol.Count
                        Else
                            countFail += 1
                        End If
                    End If

                    '-------------------------------------------------
                    ' 6. B3 - HÀNG DƯỚI (BASE = CẠNH TRÁI)
                    '-------------------------------------------------
                    Dim botRow As List(Of HoleInfo) =
                        allHoles.Where(Function(h) h.Y < limitY).
                        OrderBy(Function(h) h.X).ToList()

                    If botRow.Count > 0 Then
                        Dim result As Boolean =
                            AddBaselineSet(oSheet, app, leftEdge, botRow, True, minY - 3.2)

                        If result Then
                            For Each h As HoleInfo In botRow
                                h.Dimmed = True
                            Next
                            countOK += botRow.Count
                        Else
                            countFail += 1
                        End If
                    End If

                    '-------------------------------------------------
                    ' 7. B4 - CỘT PHẢI (BASE = CẠNH TRÊN)
                    '-------------------------------------------------
                    Dim rightCol As List(Of HoleInfo) =
                        allHoles.Where(Function(h) h.X > limitX).
                        OrderByDescending(Function(h) h.Y).ToList()

                    If rightCol.Count > 0 Then
                        Dim result As Boolean =
                            AddBaselineSet(oSheet, app, topEdge, rightCol, False, maxX + 3.2)

                        If result Then
                            For Each h As HoleInfo In rightCol
                                h.Dimmed = True
                            Next
                            countOK += rightCol.Count
                        Else
                            countFail += 1
                        End If
                    End If

                    '-------------------------------------------------
                    ' 8. QUÉT LẠI LỖ CHƯA ĐƯỢC DIM
                    '-------------------------------------------------
                    Dim missingHoles As List(Of HoleInfo) =
                        allHoles.Where(Function(h) Not h.Dimmed).ToList()

                    For Each hole As HoleInfo In missingHoles
                        Dim done As Boolean = False

                        If hole.X <= limitX Then
                            done = AddSingleBaseline(oSheet, app, leftEdge, hole.Curve, True, minX - 3.2)
                        Else
                            done = AddSingleBaseline(oSheet, app, leftEdge, hole.Curve, True, hole.Y + 2.5)
                        End If

                        If done Then
                            hole.Dimmed = True
                            countOK += 1
                            countAdd += 1
                            Continue For
                        End If

                        If hole.X <= limitX Then
                            done = AddSingleBaseline(oSheet, app, topEdge, hole.Curve, False, minX - 3.2)
                        Else
                            done = AddSingleBaseline(oSheet, app, topEdge, hole.Curve, False, maxX + 3.2)
                        End If

                        If done Then
                            hole.Dimmed = True
                            countOK += 1
                            countAdd += 1
                        Else
                            countFail += 1
                        End If
                    Next

                    '-------------------------------------------------
                    ' 9. QUÉT LẦN CUỐI
                    '-------------------------------------------------
                    For Each hole As HoleInfo In allHoles
                        If hole.Dimmed Then Continue For

                        Dim ok As Boolean = False

                        ok = AddSingleBaseline(oSheet, app, leftEdge, hole.Curve, True, minX - 4.0)

                        If Not ok Then
                            ok = AddSingleBaseline(oSheet, app, topEdge, hole.Curve, False, maxX + 4.0)
                        End If

                        If ok Then
                            hole.Dimmed = True
                            countOK += 1
                            countAdd += 1
                        End If
                    Next

                Next ' End selectedViews

                '=====================================================
                ' AUTO ARRANGE DIMENSIONS (ĐÚNG API)
                '=====================================================
                Try
                    ArrangeDimensions(oSheet, app)
                Catch
                End Try

                oDrawDoc.Update()

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Số view đã xử lý: " & selectedViews.Count & vbCrLf &
                    "Số dim tạo: " & countOK & vbCrLf &
                    "Dim bổ sung: " & countAdd & vbCrLf &
                    "Lỗi: " & countFail & vbCrLf & vbCrLf &
                    "• Chọn nhiều View liên tục" & vbCrLf &
                    "• Dùng Baseline Dimension" & vbCrLf &
                    "• Ngang: Base = Cạnh Trái" & vbCrLf &
                    "• Dọc: Base = Cạnh Trên" & vbCrLf &
                    "• Auto Arrange Dimension",
                    "Dim Baseline lỗ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show(
                    "Lỗi:" & vbCrLf & ex.Message,
                    "Dim Baseline lỗ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End Try

        End Sub

        '=============================================================
        ' AUTO ARRANGE - ĐÚNG API INVENTOR
        '=============================================================
        Private Sub ArrangeDimensions(ByVal oSheet As Sheet, ByVal app As Inventor.Application)

            Try
                Dim oDims As DrawingDimensions = oSheet.DrawingDimensions
                If oDims Is Nothing OrElse oDims.Count = 0 Then Exit Sub

                Dim oCol As ObjectCollection = app.TransientObjects.CreateObjectCollection

                For Each oDim As DrawingDimension In oDims
                    Try
                        ' Linear + Angular
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

                ' Baseline Dimension Set cũng đưa vào
                For Each bSet As BaselineDimensionSet In oDims.BaselineDimensionSets
                    Try
                        oCol.Add(bSet)
                    Catch
                    End Try
                Next

                If oCol.Count > 0 Then
                    oDims.Arrange(oCol)          ' <-- Lệnh chính thức
                End If

            Catch
            End Try

        End Sub

        '=============================================================
        ' TẠO BASELINE SET
        '=============================================================
        Private Function AddBaselineSet(
            ByVal oSheet As Sheet,
            ByVal app As Inventor.Application,
            ByVal baseCurve As DrawingCurve,
            ByVal holeList As List(Of HoleInfo),
            ByVal horizontal As Boolean,
            ByVal offset As Double) As Boolean

            Try
                If baseCurve Is Nothing Then Return False
                If holeList Is Nothing OrElse holeList.Count = 0 Then Return False

                Dim intents As ObjectCollection =
                    app.TransientObjects.CreateObjectCollection()

                ' Intent đầu tiên = BASE
                intents.Add(oSheet.CreateGeometryIntent(baseCurve))

                For Each hole As HoleInfo In holeList
                    If hole Is Nothing OrElse hole.Curve Is Nothing Then Continue For
                    intents.Add(
                        oSheet.CreateGeometryIntent(
                            hole.Curve,
                            PointIntentEnum.kCenterPointIntent))
                Next

                If intents.Count < 2 Then Return False

                Dim firstHole As HoleInfo = holeList.First()
                Dim placement As Point2d

                If horizontal Then
                    placement = app.TransientGeometry.CreatePoint2d(firstHole.X, offset)
                Else
                    placement = app.TransientGeometry.CreatePoint2d(offset, firstHole.Y)
                End If

                Dim dimType As DimensionTypeEnum =
                    If(horizontal,
                       DimensionTypeEnum.kHorizontalDimensionType,
                       DimensionTypeEnum.kVerticalDimensionType)

                Dim baseSets As BaselineDimensionSets =
                    oSheet.DrawingDimensions.BaselineDimensionSets

                Dim baseSet As BaselineDimensionSet =
                    baseSets.Add(intents, placement, dimType)

                If baseSet Is Nothing Then Return False

                Try
                    baseSet.ArrangeText()
                Catch
                End Try

                Return True

            Catch
                Return False
            End Try

        End Function

        '=============================================================
        ' TẠO 1 BASELINE RIÊNG
        '=============================================================
        Private Function AddSingleBaseline(
            ByVal oSheet As Sheet,
            ByVal app As Inventor.Application,
            ByVal baseCurve As DrawingCurve,
            ByVal holeCurve As DrawingCurve,
            ByVal horizontal As Boolean,
            ByVal offset As Double) As Boolean

            Try
                If baseCurve Is Nothing OrElse holeCurve Is Nothing Then Return False

                Dim intents As ObjectCollection =
                    app.TransientObjects.CreateObjectCollection()

                intents.Add(oSheet.CreateGeometryIntent(baseCurve))
                intents.Add(
                    oSheet.CreateGeometryIntent(
                        holeCurve,
                        PointIntentEnum.kCenterPointIntent))

                Dim c As Point2d = holeCurve.CenterPoint
                If c Is Nothing Then Return False

                Dim placement As Point2d
                If horizontal Then
                    placement = app.TransientGeometry.CreatePoint2d(c.X, offset)
                Else
                    placement = app.TransientGeometry.CreatePoint2d(offset, c.Y)
                End If

                Dim dimType As DimensionTypeEnum =
                    If(horizontal,
                       DimensionTypeEnum.kHorizontalDimensionType,
                       DimensionTypeEnum.kVerticalDimensionType)

                Dim baseSets As BaselineDimensionSets =
                    oSheet.DrawingDimensions.BaselineDimensionSets

                Dim baseSet As BaselineDimensionSet =
                    baseSets.Add(intents, placement, dimType)

                If baseSet Is Nothing Then Return False

                Try
                    baseSet.ArrangeText()
                Catch
                End Try

                Return True

            Catch
                Return False
            End Try

        End Function

        '=============================================================
        ' HOLE INFO
        '=============================================================
        Public Class HoleInfo
            Public Curve As DrawingCurve
            Public Center As Point2d
            Public X As Double
            Public Y As Double
            Public Dimmed As Boolean
        End Class

    End Module
End Namespace