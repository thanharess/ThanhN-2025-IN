Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq

Namespace ToolInventor2025.Drawing.Buttons.Drawdim

    Public Module Draw_dim_base_line_c

        Private Const TOL As Double = 0.00
        Private Const EDGE_TOL As Double = 0.03

        '=========================================================
        ' LƯU THÔNG SỐ GIỮA CÁC LẦN CHẠY (đơn vị cm)
        '=========================================================
        Private lastXNearGap As Double = 0.2      ' 2 mm
        Private lastYNearGap As Double = 0.2
        Private lastMinGapX As Double = 0.5        ' 5 mm
        Private lastMinGapY As Double = 0.5
        Private lastNearMode As Boolean = True
        Private lastUseTier As Boolean = True
        Private lastIncludeHoles As Boolean = True

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

                Dim oSheet As Sheet =
                    oDrawDoc.ActiveSheet

                Dim tg As TransientGeometry =
                    app.TransientGeometry

                '=========================================================
                ' CHỌN VIEW
                '=========================================================
                Dim selectedViews As New List(Of DrawingView)

                Do

                    Dim oSS As SelectSet =
                        oDrawDoc.SelectSet

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

                    If Not already Then
                        selectedViews.Add(oView)
                    End If

                Loop

                If selectedViews.Count = 0 Then

                    MessageBox.Show(
                        "Chưa chọn View nào.",
                        "Thông báo")

                    Exit Sub

                End If

                '=========================================================
                ' FORM — load giá trị đã lưu từ lần chạy trước
                '=========================================================
                Dim hFromRight As Boolean = False
                Dim vFromBottom As Boolean = False

                Dim hPickedEdge As DrawingCurve = Nothing
                Dim vPickedEdge As DrawingCurve = Nothing

                Dim includeHoles As Boolean = lastIncludeHoles

                Dim xNearGap As Double = lastXNearGap       ' cm
                Dim yNearGap As Double = lastYNearGap

                Dim minGapX As Double = lastMinGapX
                Dim minGapY As Double = lastMinGapY

                Dim useTier As Boolean = lastUseTier
                Dim nearMode As Boolean = lastNearMode

                If Not ShowDirectionDialog(
                    hFromRight,
                    vFromBottom,
                    hPickedEdge,
                    vPickedEdge,
                    includeHoles,
                    xNearGap,
                    yNearGap,
                    minGapX,
                    minGapY,
                    useTier,
                    nearMode,
                    app) Then Exit Sub

                '=========================================================
                ' LƯU LẠI CHO LẦN CHẠY SAU
                '=========================================================
                lastXNearGap = xNearGap
                lastYNearGap = yNearGap
                lastMinGapX = minGapX
                lastMinGapY = minGapY
                lastUseTier = useTier
                lastNearMode = nearMode
                lastIncludeHoles = includeHoles

                Dim countOK As Integer = 0
                Dim countFail As Integer = 0

                Dim dimmedCurves As List(Of DrawingCurve) =
                    CollectDimmedCurves(oSheet)

                '=========================================================
                ' XỬ LÝ TỪNG VIEW
                '=========================================================
                For Each oView As DrawingView In selectedViews

                    Dim allEdges As New List(Of EdgeInfo)

                    '=====================================================
                    ' LẤY CẠNH + LỖ
                    '=====================================================
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
                                    .Curve = oCurve,
                                    .IsVertical = isV,
                                    .IsHorizontal = isH,
                                    .IsHole = False,
                                    .X = (p1.X + p2.X) / 2.0,
                                    .Y = (p1.Y + p2.Y) / 2.0,
                                    .Dimmed = False
                                }

                                allEdges.Add(ei)

                            ElseIf includeHoles AndAlso
                                   oCurve.CurveType = CurveTypeEnum.kCircleCurve Then

                                Dim c As Point2d = oCurve.CenterPoint

                                If c Is Nothing Then Continue For

                                Dim hi As New EdgeInfo With {
                                    .Curve = oCurve,
                                    .IsVertical = False,
                                    .IsHorizontal = False,
                                    .IsHole = True,
                                    .X = c.X,
                                    .Y = c.Y,
                                    .Dimmed = False
                                }

                                allEdges.Add(hi)

                            End If

                        Catch
                        End Try

                    Next

                    If allEdges.Count = 0 Then Continue For

                    '=====================================================
                    ' DEDUP
                    '=====================================================
                    Dim uniqueEdges As New List(Of EdgeInfo)

                    For Each e As EdgeInfo In allEdges

                        Dim dup As Boolean = False

                        For Each u As EdgeInfo In uniqueEdges

                            If e.IsHole AndAlso u.IsHole Then

                                If Math.Abs(u.X - e.X) <= EDGE_TOL AndAlso
                                   Math.Abs(u.Y - e.Y) <= EDGE_TOL Then

                                    dup = True
                                    Exit For

                                End If

                            ElseIf (Not e.IsHole) AndAlso (Not u.IsHole) Then

                                If SameSegment(u.Curve, e.Curve) Then

                                    dup = True
                                    Exit For

                                End If

                            End If

                        Next

                        If Not dup Then uniqueEdges.Add(e)

                    Next

                    allEdges = uniqueEdges

                    '=====================================================
                    ' 4 CẠNH NGOÀI
                    '=====================================================
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
                               oCurve.CurveType <> CurveTypeEnum.kLineSegmentCurve Then
                                Continue For
                            End If

                            Dim p1 As Point2d = oCurve.StartPoint
                            Dim p2 As Point2d = oCurve.EndPoint

                            If p1 Is Nothing OrElse p2 Is Nothing Then Continue For

                            If Math.Abs(p1.X - p2.X) < EDGE_TOL Then

                                If p1.X < minX Then
                                    minX = p1.X
                                    leftEdge = oCurve
                                End If

                                If p1.X > maxX Then
                                    maxX = p1.X
                                    rightEdge = oCurve
                                End If

                            End If

                            If Math.Abs(p1.Y - p2.Y) < EDGE_TOL Then

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

                    If leftEdge Is Nothing OrElse
                       rightEdge Is Nothing OrElse
                       topEdge Is Nothing OrElse
                       bottomEdge Is Nothing Then
                        Continue For
                    End If

                    If (maxX - minX) <= 0 OrElse (maxY - minY) <= 0 Then
                        Continue For
                    End If

                    Dim usedX As New List(Of Double)
                    Dim usedY As New List(Of Double)

                    '=====================================================
                    ' BASE X
                    '=====================================================
                    Dim hBaseEdge As DrawingCurve
                    Dim hBaseCoord As Double

                    If EdgeInView(hPickedEdge, oView) Then

                        hBaseEdge = hPickedEdge
                        hBaseCoord =
                            (hPickedEdge.StartPoint.X +
                             hPickedEdge.EndPoint.X) / 2.0

                    Else

                        hBaseEdge =
                            If(hFromRight, rightEdge, leftEdge)

                        hBaseCoord =
                            If(hFromRight, maxX, minX)

                    End If

                    '=====================================================
                    ' BASE Y
                    '=====================================================
                    Dim vBaseEdge As DrawingCurve
                    Dim vBaseCoord As Double

                    If EdgeInView(vPickedEdge, oView) Then

                        vBaseEdge = vPickedEdge
                        vBaseCoord =
                            (vPickedEdge.StartPoint.Y +
                             vPickedEdge.EndPoint.Y) / 2.0

                    Else

                        vBaseEdge =
                            If(vFromBottom, bottomEdge, topEdge)

                        vBaseCoord =
                            If(vFromBottom, minY, maxY)

                    End If

                    Dim hBaseOnRight As Boolean =
                        hBaseCoord > (minX + maxX) / 2.0

                    Dim vBaseOnBottom As Boolean =
                        vBaseCoord < (minY + maxY) / 2.0

                    '=====================================================
                    ' DIM X
                    '=====================================================
                    Dim xTargets As List(Of EdgeInfo) =
                        allEdges.Where(Function(e) e.IsVertical OrElse e.IsHole).ToList()

                    Dim filteredX As List(Of EdgeInfo) =
                        FilterTargets(xTargets, hBaseCoord, minGapX, nearMode, True)

                    Dim tierIndexX As Integer = 0
                    Dim tierStepX As Double = If(useTier, 0.0, 0.0) '0.4 đầu tên

                    For Each e As EdgeInfo In filteredX

                        If Math.Abs(e.X - hBaseCoord) <= EDGE_TOL Then Continue For
                        If CurveHasDim(dimmedCurves, e.Curve) Then Continue For
                        If ValueExists(usedX, e.X) Then Continue For

                        Dim offsetY As Double = xNearGap + tierIndexX * tierStepX
                        If useTier Then tierIndexX += 1

                        Dim placeY As Double = If(hBaseOnRight, minY - offsetY, maxY + offsetY)

                        Try

                            Dim baseIntent As GeometryIntent = oSheet.CreateGeometryIntent(hBaseEdge)
                            Dim edgeIntent As GeometryIntent = MakeIntent(oSheet, e)

                            If edgeIntent Is Nothing Then Continue For

                            Dim placement As Point2d = tg.CreatePoint2d(e.X, placeY)

                            Dim newDim As GeneralDimension =
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    placement, baseIntent, edgeIntent,
                                    DimensionTypeEnum.kHorizontalDimensionType)

                            If newDim IsNot Nothing Then
                                usedX.Add(e.X)
                                e.Dimmed = True
                                countOK += 1
                                If Not CurveHasDim(dimmedCurves, e.Curve) Then
                                    dimmedCurves.Add(e.Curve)
                                End If
                            End If

                        Catch
                            countFail += 1
                        End Try

                    Next

                    '=====================================================
                    ' DIM Y
                    '=====================================================
                    Dim yTargets As List(Of EdgeInfo) =
                        allEdges.Where(Function(e) e.IsHorizontal OrElse e.IsHole).ToList()

                    Dim filteredY As List(Of EdgeInfo) =
                        FilterTargets(yTargets, vBaseCoord, minGapY, nearMode, False)

                    Dim tierIndexY As Integer = 0
                    Dim tierStepY As Double = If(useTier, 0.0, 0.0) '0.4 đầu tên

                    For Each e As EdgeInfo In filteredY

                        If Math.Abs(e.Y - vBaseCoord) <= EDGE_TOL Then Continue For
                        If CurveHasDim(dimmedCurves, e.Curve) Then Continue For
                        If ValueExists(usedY, e.Y) Then Continue For

                        Dim offsetX As Double = yNearGap + tierIndexY * tierStepY
                        If useTier Then tierIndexY += 1

                        Dim placeX As Double = If(vBaseOnBottom, maxX + offsetX, minX - offsetX)

                        Try

                            Dim baseIntent As GeometryIntent = oSheet.CreateGeometryIntent(vBaseEdge)
                            Dim edgeIntent As GeometryIntent = MakeIntent(oSheet, e)

                            If edgeIntent Is Nothing Then Continue For

                            Dim placement As Point2d = tg.CreatePoint2d(placeX, e.Y)

                            Dim newDim As GeneralDimension =
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    placement, baseIntent, edgeIntent,
                                    DimensionTypeEnum.kVerticalDimensionType)

                            If newDim IsNot Nothing Then
                                usedY.Add(e.Y)
                                e.Dimmed = True
                                countOK += 1
                                If Not CurveHasDim(dimmedCurves, e.Curve) Then
                                    dimmedCurves.Add(e.Curve)
                                End If
                            End If

                        Catch
                            countFail += 1
                        End Try

                    Next
                Next

                '=========================================================
                ' SẮP XẾP DIM — CHỈ DIM THUỘC VIEW ĐÃ CHỌN
                '=========================================================
                Try
                    ArrangeDimensions(oSheet, app, selectedViews)
                Catch
                End Try

                oDrawDoc.Update()

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "View: " & selectedViews.Count & vbCrLf &
                    "Dim tạo: " & countOK & vbCrLf &
                    "Lỗi: " & countFail & vbCrLf & vbCrLf &
                    "--- Thông số đã áp dụng ---" & vbCrLf &
                    "Chế độ: " & If(nearMode, "GẦN", "XA") & vbCrLf &
                    "X gần: " & (xNearGap * 10.0).ToString("0.0#") & " mm" & vbCrLf &
                    "Y gần: " & (yNearGap * 10.0).ToString("0.0#") & " mm" & vbCrLf &
                    "MinGap X: " & (minGapX * 10.0).ToString("0.0#") & " mm" & vbCrLf &
                    "MinGap Y: " & (minGapY * 10.0).ToString("0.0#") & " mm",
                    "Dim Baseline cạnh + lỗ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)

            Catch ex As Exception

                MessageBox.Show(
                    "Lỗi:" & vbCrLf & ex.Message,
                    "Dim Baseline cạnh + lỗ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub

        '=============================================================
        ' TẠO GEOMETRY INTENT
        '=============================================================
        Private Function MakeIntent(
            oSheet As Sheet,
            e As EdgeInfo) As GeometryIntent

            If e Is Nothing OrElse e.Curve Is Nothing Then
                Return Nothing
            End If

            If e.IsHole Then
                Return oSheet.CreateGeometryIntent(
                    e.Curve,
                    PointIntentEnum.kCenterPointIntent)
            Else
                Return oSheet.CreateGeometryIntent(e.Curve)
            End If

        End Function

        '=============================================================
        ' KIỂM TRA CẠNH CÓ THUỘC VIEW
        '=============================================================
        Private Function EdgeInView(
            edge As DrawingCurve,
            view As DrawingView) As Boolean

            If edge Is Nothing OrElse view Is Nothing Then
                Return False
            End If

            Try
                For Each c As DrawingCurve In view.DrawingCurves
                    If c Is edge Then Return True
                Next
            Catch
            End Try

            Return False

        End Function

        '=============================================================
        ' LỌC CẠNH THEO MODE GẦN / XA
        '=============================================================
        Private Function FilterTargets(
    source As List(Of EdgeInfo),
    baseCoord As Double,
    minGap As Double,
    nearMode As Boolean,
    isX As Boolean) As List(Of EdgeInfo)

            Dim result As New List(Of EdgeInfo)

            If source Is Nothing OrElse source.Count = 0 Then
                Return result
            End If

            Dim DistToBase As Func(Of EdgeInfo, Double) =
        Function(e As EdgeInfo) As Double
            If isX Then
                Return Math.Abs(e.X - baseCoord)
            Else
                Return Math.Abs(e.Y - baseCoord)
            End If
        End Function

            ' Sắp xếp từ gần Base → xa Base, bỏ chính Base
            Dim ordered As List(Of EdgeInfo) =
        source.Where(Function(e) DistToBase(e) > 0.001).
               OrderBy(Function(e) DistToBase(e)).
               ToList()

            If ordered.Count = 0 Then Return result

            '---------------------------------------------------------
            ' Bước 1: Lọc cạnh đầu tiên (phải ≥ minGap so với Base)
            '---------------------------------------------------------
            Dim firstFound As Boolean = False

            For Each e As EdgeInfo In ordered

                Dim d As Double = DistToBase(e)

                If Not firstFound Then
                    If d < minGap Then Continue For   ' quá gần Base → bỏ

                    result.Add(e)
                    firstFound = True
                    Continue For
                End If

                '-----------------------------------------------------
                ' Các cạnh tiếp theo
                '-----------------------------------------------------
                Dim lastKept As EdgeInfo = result(result.Count - 1)
                Dim lastDist As Double = DistToBase(lastKept)
                Dim gapBetween As Double = d - lastDist

                If gapBetween >= minGap Then
                    ' Đủ xa → giữ cả hai
                    result.Add(e)
                Else
                    ' Quá gần nhau
                    If nearMode Then
                        ' GẦN → giữ cạnh GẦN hơn (bỏ e - cái đang xét)
                        ' Không thêm e
                    Else
                        ' XA → giữ cạnh XA hơn (thay lastKept bằng e)
                        result(result.Count - 1) = e
                    End If
                End If

            Next

            '---------------------------------------------------------
            ' Bước 2 (chỉ cho GẦN): Dọn lần cuối
            ' Đảm bảo không còn cặp nào < minGap, ưu tiên giữ gần
            '---------------------------------------------------------
            If nearMode AndAlso result.Count > 1 Then

                Dim cleaned As New List(Of EdgeInfo)
                cleaned.Add(result(0))

                For i As Integer = 1 To result.Count - 1
                    Dim prev = cleaned(cleaned.Count - 1)
                    Dim curr = result(i)

                    Dim gap As Double = DistToBase(curr) - DistToBase(prev)

                    If gap >= minGap Then
                        cleaned.Add(curr)
                    End If
                    ' Nếu gap < minGap → bỏ curr (giữ prev - gần hơn)
                Next

                result = cleaned

            End If

            Return result

        End Function

        '=============================================================
        ' FORM
        '=============================================================
        Private Function ShowDirectionDialog(
            ByRef hFromRight As Boolean,
            ByRef vFromBottom As Boolean,
            ByRef hPickedEdge As DrawingCurve,
            ByRef vPickedEdge As DrawingCurve,
            ByRef includeHoles As Boolean,
            ByRef xNearGap As Double,
            ByRef yNearGap As Double,
            ByRef minGapX As Double,
            ByRef minGapY As Double,
            ByRef useTier As Boolean,
            ByRef nearMode As Boolean,
            app As Inventor.Application) As Boolean

            Dim frm As New Form With {
                .Text = "Hướng chuẩn Baseline",
                .ClientSize = New Size(460, 525),
                .StartPosition = FormStartPosition.CenterScreen,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .ShowInTaskbar = False
            }

            Dim hHolder As New EdgePickHolder()
            Dim vHolder As New EdgePickHolder()

            Dim okClicked As Boolean = False

            '=========================================================
            ' GROUP NGANG
            '=========================================================
            Dim gbH As New GroupBox With {
                .Text = "Hướng chuẩn NGANG (X)",
                .Bounds = New Rectangle(12, 12, 396, 110)
            }

            Dim rbHLeft As New RadioButton With {
                .Text = "Từ TRÁI → Phải (mặc định)",
                .Bounds = New Rectangle(15, 22, 370, 20),
                .Checked = True
            }

            Dim rbHRight As New RadioButton With {
                .Text = "Từ PHẢI → Trái",
                .Bounds = New Rectangle(15, 46, 370, 20)
            }

            Dim lblH As New Label With {
                .Text = "(chưa pick cạnh)",
                .Bounds = New Rectangle(15, 74, 250, 20),
                .ForeColor = System.Drawing.Color.Gray,
                .Font = New Font("Segoe UI", 8, FontStyle.Italic)
            }

            Dim btnPickH As New Button With {
                .Text = "Pick cạnh X...",
                .Bounds = New Rectangle(270, 70, 105, 26)
            }

            gbH.Controls.AddRange({rbHLeft, rbHRight, lblH, btnPickH})
            frm.Controls.Add(gbH)

            '=========================================================
            ' GROUP DỌC
            '=========================================================
            Dim gbV As New GroupBox With {
                .Text = "Hướng chuẩn DỌC (Y)",
                .Bounds = New Rectangle(12, 130, 396, 110)
            }

            Dim rbVTop As New RadioButton With {
                .Text = "Từ TRÊN → Dưới (mặc định)",
                .Bounds = New Rectangle(15, 22, 370, 20),
                .Checked = True
            }

            Dim rbVBottom As New RadioButton With {
                .Text = "Từ DƯỚI → Trên",
                .Bounds = New Rectangle(15, 46, 370, 20)
            }

            Dim lblV As New Label With {
                .Text = "(chưa pick cạnh)",
                .Bounds = New Rectangle(15, 74, 250, 20),
                .ForeColor = System.Drawing.Color.Gray,
                .Font = New Font("Segoe UI", 8, FontStyle.Italic)
            }

            Dim btnPickV As New Button With {
                .Text = "Pick cạnh Y...",
                .Bounds = New Rectangle(270, 70, 105, 26)
            }

            gbV.Controls.AddRange({rbVTop, rbVBottom, lblV, btnPickV})
            frm.Controls.Add(gbV)

            '=========================================================
            ' GROUP KHOẢNG CÁCH
            '=========================================================
            Dim gbO As New GroupBox With {
                .Text = "Khoảng cách",
                .Bounds = New Rectangle(12, 250, 396, 165)
            }

            Dim cbXNear As New ComboBox With {
                .DropDownStyle = ComboBoxStyle.DropDown,
                .Bounds = New Rectangle(90, 22, 90, 22)
            }
            cbXNear.Items.AddRange({"1.0", "2.0", "2.5", "3.2", "5.0", "8.0", "10.0"})
            cbXNear.Text = (xNearGap * 20.0).ToString("0.0#")

            Dim cbYNear As New ComboBox With {
                .DropDownStyle = ComboBoxStyle.DropDown,
                .Bounds = New Rectangle(265, 22, 90, 22)
            }
            cbYNear.Items.AddRange({"1.0", "2.0", "2.5", "3.2", "5.0", "8.0", "10.0"})
            cbYNear.Text = (yNearGap * 20.0).ToString("0.0#")

            Dim cbMinX As New ComboBox With {
                .DropDownStyle = ComboBoxStyle.DropDown,
                .Bounds = New Rectangle(110, 55, 90, 22)
            }
            cbMinX.Items.AddRange({"0", "1", "2", "5", "10", "15", "20", "30"})
            cbMinX.Text = (minGapX * 20.0).ToString("0.0#")

            Dim cbMinY As New ComboBox With {
                .DropDownStyle = ComboBoxStyle.DropDown,
                .Bounds = New Rectangle(310, 55, 78, 22)
            }
            cbMinY.Items.AddRange({"0", "1", "2", "5", "10", "15", "20", "30"})
            cbMinY.Text = (minGapY * 20.0).ToString("0.0#")

            Dim rbNear As New RadioButton With {
                .Text = "GẦN - ưu tiên cạnh gần Base",
                .Bounds = New Rectangle(15, 88, 175, 22),
                .Checked = nearMode
            }

            Dim rbFar As New RadioButton With {
                .Text = "XA - ưu tiên cạnh xa Base",
                .Bounds = New Rectangle(195, 88, 180, 22),
                .Checked = Not nearMode
            }

            Dim chkTier As New CheckBox With {
                .Text = "Bật khoảng cách bậc dim line (mỗi dim lệch 4 mm)",
                .Bounds = New Rectangle(15, 125, 370, 22),
                .Checked = useTier
            }

            gbO.Controls.AddRange(
                {
                    New Label With {
                        .Text = "X gần (mm):",
                        .Bounds = New Rectangle(15, 24, 75, 22)
                    },
                    cbXNear,
                    New Label With {
                        .Text = "Y gần (mm):",
                        .Bounds = New Rectangle(200, 24, 65, 22)
                    },
                    cbYNear,
                    New Label With {
                        .Text = "Bỏ nếu X < (mm):",
                        .Bounds = New Rectangle(15, 57, 95, 22)
                    },
                    cbMinX,
                    New Label With {
                        .Text = "Bỏ nếu Y < (mm):",
                        .Bounds = New Rectangle(210, 57, 95, 22)
                    },
                    cbMinY, rbNear, rbFar, chkTier
                })

            frm.Controls.Add(gbO)

            '=========================================================
            ' LỖ
            '=========================================================
            Dim chkHoles As New CheckBox With {
                .Text = "Bao gồm LỖ TRÒN (đo theo tâm lỗ)",
                .Bounds = New Rectangle(15, 430, 396, 22),
                .Checked = includeHoles
            }

            frm.Controls.Add(chkHoles)

            '=========================================================
            ' BUTTON
            '=========================================================
            Dim btnOK As New Button With {
                .Text = "OK",
                .Bounds = New Rectangle(215, 470, 85, 28)
            }

            Dim btnCancel As New Button With {
                .Text = "Hủy",
                .Bounds = New Rectangle(310, 470, 85, 28)
            }
            frm.Controls.AddRange({btnOK, btnCancel})

            '=========================================================
            ' BIẾN TẠM
            '=========================================================
            Dim tempXNear As Double = xNearGap
            Dim tempYNear As Double = yNearGap
            Dim tempMinX As Double = minGapX
            Dim tempMinY As Double = minGapY
            Dim tempUseTier As Boolean = useTier
            Dim tempNearMode As Boolean = nearMode
            Dim tempIncludeHoles As Boolean = includeHoles
            Dim tempHFromRight As Boolean = False
            Dim tempVFromBottom As Boolean = False

            '=========================================================
            ' OK — PARSE NGAY TẠI ĐÂY
            '=========================================================
            AddHandler btnOK.Click,
    Sub()
        Dim tmp As Double

        ' Ép commit giá trị ComboBox
        frm.ActiveControl = btnOK
        System.Windows.Forms.Application.DoEvents()

        ' X gần
        If Not Double.TryParse(cbXNear.Text.Replace(",", ".").Trim(),
                               Globalization.NumberStyles.Any,
                               Globalization.CultureInfo.InvariantCulture, tmp) OrElse tmp <= 0 Then
            tmp = 2
        End If
        tempXNear = tmp / 10.0

        ' Y gần
        If Not Double.TryParse(cbYNear.Text.Replace(",", ".").Trim(),
                               Globalization.NumberStyles.Any,
                               Globalization.CultureInfo.InvariantCulture, tmp) OrElse tmp <= 0 Then
            tmp = 2
        End If
        tempYNear = tmp / 10.0

        ' MinGap X
        If Not Double.TryParse(cbMinX.Text.Replace(",", ".").Trim(),
                               Globalization.NumberStyles.Any,
                               Globalization.CultureInfo.InvariantCulture, tmp) OrElse tmp < 0 Then
            tmp = 2
        End If
        tempMinX = tmp / 10.0

        ' MinGap Y
        If Not Double.TryParse(cbMinY.Text.Replace(",", ".").Trim(),
                               Globalization.NumberStyles.Any,
                               Globalization.CultureInfo.InvariantCulture, tmp) OrElse tmp < 0 Then
            tmp = 2
        End If
        tempMinY = tmp / 10.0

        tempUseTier = chkTier.Checked
        tempNearMode = rbNear.Checked
        tempIncludeHoles = chkHoles.Checked
        tempHFromRight = rbHRight.Checked
        tempVFromBottom = rbVBottom.Checked

        ' Xác nhận

        okClicked = True
        frm.Close()
    End Sub

            '=========================================================
            ' CANCEL
            '=========================================================
            AddHandler btnCancel.Click,
    Sub()
        okClicked = False
        frm.Close()
    End Sub

            '=========================================================
            ' PICK X
            '=========================================================
            AddHandler btnPickH.Click,
    Sub()
        frm.Hide()
        frm.Refresh()
        System.Windows.Forms.Application.DoEvents()
        System.Threading.Thread.Sleep(200)
        System.Windows.Forms.Application.DoEvents()
        Try
            Dim seg = TryCast(
                app.CommandManager.Pick(
                    SelectionFilterEnum.kDrawingCurveSegmentFilter,
                    "Chọn cạnh ĐỨNG làm gốc đo X"),
                DrawingCurveSegment)
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
        frm.Show()
        frm.Activate()
    End Sub

            '=========================================================
            ' PICK Y
            '=========================================================
            AddHandler btnPickV.Click,
    Sub()
        frm.Hide()
        frm.Refresh()
        System.Windows.Forms.Application.DoEvents()
        System.Threading.Thread.Sleep(200)
        System.Windows.Forms.Application.DoEvents()
        Try
            Dim seg = TryCast(
                app.CommandManager.Pick(
                    SelectionFilterEnum.kDrawingCurveSegmentFilter,
                    "Chọn cạnh NGANG làm gốc đo Y"),
                DrawingCurveSegment)
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
        frm.Show()
        frm.Activate()
    End Sub

            '=========================================================
            ' MODELESS
            '=========================================================
            frm.Show()
            Do While frm.Visible
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(15)
            Loop

            If Not okClicked Then
                Return False
            End If

            ' Gán giá trị đã parse từ OK handler
            xNearGap = tempXNear
            yNearGap = tempYNear
            minGapX = tempMinX
            minGapY = tempMinY
            useTier = tempUseTier
            nearMode = tempNearMode
            includeHoles = tempIncludeHoles
            hFromRight = tempHFromRight
            vFromBottom = tempVFromBottom
            hPickedEdge = hHolder.Edge
            vPickedEdge = vHolder.Edge

            Return True

        End Function

        '=============================================================
        ' ARRANGE — CHỈ DIM THUỘC VIEW ĐÃ CHỌN (giống Draw_2b)
        '=============================================================
        Private Sub ArrangeDimensions(
            oSheet As Sheet,
            app As Inventor.Application,
            selectedViews As List(Of DrawingView))

            Try

                Dim oDims = oSheet.DrawingDimensions

                If oDims Is Nothing OrElse oDims.Count = 0 Then Exit Sub

                Dim oCol = app.TransientObjects.CreateObjectCollection

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

                If oCol.Count > 1 Then
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

        '=============================================================
        ' VALUE EXISTS
        '=============================================================
        Private Function ValueExists(
            list As List(Of Double),
            value As Double) As Boolean

            If list Is Nothing Then Return False

            For Each v In list
                If Math.Abs(v - value) <= EDGE_TOL Then Return True
            Next

            Return False

        End Function

        '=============================================================
        ' SAME SEGMENT
        '=============================================================
        Private Function SameSegment(
            c1 As DrawingCurve,
            c2 As DrawingCurve) As Boolean

            If c1 Is Nothing OrElse c2 Is Nothing Then Return False

            Try

                Dim a1 = c1.StartPoint
                Dim a2 = c1.EndPoint
                Dim b1 = c2.StartPoint
                Dim b2 = c2.EndPoint

                If a1 Is Nothing OrElse a2 Is Nothing OrElse
                   b1 Is Nothing OrElse b2 Is Nothing Then Return False

                Return (
                    Math.Abs(a1.X - b1.X) < EDGE_TOL AndAlso
                    Math.Abs(a1.Y - b1.Y) < EDGE_TOL AndAlso
                    Math.Abs(a2.X - b2.X) < EDGE_TOL AndAlso
                    Math.Abs(a2.Y - b2.Y) < EDGE_TOL
                ) OrElse (
                    Math.Abs(a1.X - b2.X) < EDGE_TOL AndAlso
                    Math.Abs(a1.Y - b2.Y) < EDGE_TOL AndAlso
                    Math.Abs(a2.X - b1.X) < EDGE_TOL AndAlso
                    Math.Abs(a2.Y - b1.Y) < EDGE_TOL
                )

            Catch
            End Try

            Return False

        End Function

        '=============================================================
        ' SAME CIRCLE
        '=============================================================
        Private Function SameCircle(c1 As DrawingCurve, c2 As DrawingCurve) As Boolean

            If c1 Is Nothing OrElse c2 Is Nothing Then Return False

            Try

                Dim p1 = c1.CenterPoint
                Dim p2 = c2.CenterPoint

                If p1 Is Nothing OrElse p2 Is Nothing Then Return False

                Return Math.Abs(p1.X - p2.X) <= EDGE_TOL AndAlso
                       Math.Abs(p1.Y - p2.Y) <= EDGE_TOL

            Catch
                Return False
            End Try

        End Function

        '=============================================================
        ' COLLECT DIMMED CURVES
        '=============================================================
        Private Function CollectDimmedCurves(
            oSheet As Sheet) As List(Of DrawingCurve)

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

                                If dc IsNot Nothing Then
                                    result.Add(dc)
                                    Continue For
                                End If

                            End If

                            Dim dc2 = TryCast(ent, DrawingCurve)

                            If dc2 IsNot Nothing Then
                                result.Add(dc2)
                                Continue For
                            End If

                            Dim seg = TryCast(ent, DrawingCurveSegment)

                            If seg IsNot Nothing AndAlso seg.Parent IsNot Nothing Then
                                result.Add(seg.Parent)
                            End If

                        Next

                    Catch
                    End Try

                Next

            Catch
            End Try

            Return result

        End Function

        '=============================================================
        ' CURVE HAS DIM
        '=============================================================
        Private Function CurveHasDim(
            dimmed As List(Of DrawingCurve),
            c As DrawingCurve) As Boolean

            If dimmed Is Nothing OrElse c Is Nothing Then Return False

            For Each dc In dimmed

                If dc Is c Then Return True
                If SameSegment(dc, c) Then Return True
                If SameCircle(dc, c) Then Return True

            Next

            Return False

        End Function

        '=============================================================
        ' EDGE INFO
        '=============================================================
        Public Class EdgeInfo

            Public Curve As DrawingCurve
            Public IsVertical As Boolean
            Public IsHorizontal As Boolean
            Public IsHole As Boolean
            Public X As Double
            Public Y As Double
            Public Dimmed As Boolean

        End Class

        '=============================================================
        ' HOLDER
        '=============================================================
        Public Class EdgePickHolder
            Public Edge As DrawingCurve
        End Class

    End Module

End Namespace