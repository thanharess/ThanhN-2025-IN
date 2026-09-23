Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq
Namespace ToolInventor2025.Drawing.Buttons.Drawdim
    Public Module Draw_dim_base_line_b_old

        Private Const TOL As Double = 0.00
        Private Const EDGE_TOL As Double = 0.03
        Private Const DIM_GAP As Double = 3.2
        Private Const TOTAL_GAP As Double = 5.5
        '=============================================================
        ' ENTRY POINT
        '=============================================================
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
                        If v Is oView Then already = True : Exit For
                    Next
                    If Not already Then selectedViews.Add(oView)
                Loop
                If selectedViews.Count = 0 Then
                    MessageBox.Show("Chưa chọn View nào.", "Thông báo")
                    Exit Sub
                End If
                '=====================================================
                ' BẢNG CHỌN HƯỚNG + PICK CẠNH BASE + TÙY CHỌN LỖ
                '=====================================================
                Dim hFromRight As Boolean = False
                Dim vFromBottom As Boolean = False
                Dim hPickedEdge As DrawingCurve = Nothing
                Dim vPickedEdge As DrawingCurve = Nothing
                Dim includeHoles As Boolean = True
                If Not ShowDirectionDialog(hFromRight, vFromBottom,
                                           hPickedEdge, vPickedEdge,
                                           includeHoles, app) Then Exit Sub
                Dim countOK As Integer = 0
                Dim countFail As Integer = 0
                Dim dimmedCurves As List(Of DrawingCurve) = CollectDimmedCurves(oSheet)
                '=====================================================
                ' XỬ LÝ TỪNG VIEW
                '=====================================================
                For Each oView As DrawingView In selectedViews
                    '-------------------------------------------------
                    ' 1. LẤY TẤT CẢ CẠNH + LỖ
                    '-------------------------------------------------
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
                                Dim ei As New EdgeInfo
                                ei.Curve = oCurve
                                ei.IsVertical = isV
                                ei.IsHorizontal = isH
                                ei.IsHole = False
                                ei.X = (p1.X + p2.X) / 2.0
                                ei.Y = (p1.Y + p2.Y) / 2.0
                                ei.Dimmed = False
                                allEdges.Add(ei)
                            ElseIf includeHoles AndAlso
                                   oCurve.CurveType = CurveTypeEnum.kCircleCurve Then
                                Dim c As Point2d = oCurve.CenterPoint
                                If c Is Nothing Then Continue For
                                Dim hi As New EdgeInfo
                                hi.Curve = oCurve
                                hi.IsVertical = False
                                hi.IsHorizontal = False
                                hi.IsHole = True
                                hi.X = c.X
                                hi.Y = c.Y
                                hi.Dimmed = False
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
                                If Math.Abs(u.X - e.X) <= EDGE_TOL AndAlso
                                   Math.Abs(u.Y - e.Y) <= EDGE_TOL Then
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
                    '-------------------------------------------------
                    ' 2. TÌM 4 CẠNH NGOÀI  (chỉ dùng cạnh, bỏ qua lỗ)
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
                    Dim viewW As Double = maxX - minX
                    Dim viewH As Double = maxY - minY
                    If viewW <= 0 OrElse viewH <= 0 Then Continue For
                    Dim usedX As New List(Of Double)
                    Dim usedY As New List(Of Double)
                    '-------------------------------------------------
                    ' 3. XÁC ĐỊNH CẠNH BASE
                    '-------------------------------------------------
                    Dim hBaseEdge As DrawingCurve
                    Dim hBaseCoord As Double
                    Dim hBaseIsPicked As Boolean = EdgeInView(hPickedEdge, oView)
                    If hBaseIsPicked Then
                        hBaseEdge = hPickedEdge
                        hBaseCoord = (hPickedEdge.StartPoint.X + hPickedEdge.EndPoint.X) / 2.0
                    Else
                        hBaseEdge = If(hFromRight, rightEdge, leftEdge)
                        hBaseCoord = If(hFromRight, maxX, minX)
                    End If
                    Dim vBaseEdge As DrawingCurve
                    Dim vBaseCoord As Double
                    Dim vBaseIsPicked As Boolean = EdgeInView(vPickedEdge, oView)
                    If vBaseIsPicked Then
                        vBaseEdge = vPickedEdge
                        vBaseCoord = (vPickedEdge.StartPoint.Y + vPickedEdge.EndPoint.Y) / 2.0
                    Else
                        vBaseEdge = If(vFromBottom, bottomEdge, topEdge)
                        vBaseCoord = If(vFromBottom, minY, maxY)
                    End If
                    Dim hBaseOnRight As Boolean = (hBaseCoord > (minX + maxX) / 2.0)
                    Dim vBaseOnBottom As Boolean = (vBaseCoord < (minY + maxY) / 2.0)
                    '-------------------------------------------------
                    ' 5. VỊ TRÍ DIM LINE
                    '-------------------------------------------------
                    Dim hDimLineY As Double
                    If hBaseOnRight Then
                        hDimLineY = minY - DIM_GAP
                    Else
                        hDimLineY = maxY + DIM_GAP
                    End If
                    Dim vDimLineX As Double
                    If vBaseOnBottom Then
                        vDimLineX = maxX + DIM_GAP
                    Else
                        vDimLineX = minX - DIM_GAP
                    End If
                    '-------------------------------------------------
                    ' 6. DIM CẠNH ĐỨNG + LỖ (X) — TỪ BASE NGANG
                    '-------------------------------------------------
                    Dim xTargets As List(Of EdgeInfo) =
                        allEdges.Where(Function(e) e.IsVertical OrElse e.IsHole).ToList()
                    If hBaseOnRight Then
                        xTargets = xTargets.OrderByDescending(Function(e) e.X).ToList()
                    Else
                        xTargets = xTargets.OrderBy(Function(e) e.X).ToList()
                    End If
                    For Each e As EdgeInfo In xTargets
                        If Math.Abs(e.X - hBaseCoord) <= EDGE_TOL Then Continue For
                        If ValueExists(usedX, e.X) Then Continue For
                        If CurveHasDim(dimmedCurves, e.Curve) Then Continue For
                        Try
                            Dim baseIntent As GeometryIntent = oSheet.CreateGeometryIntent(hBaseEdge)
                            Dim edgeIntent As GeometryIntent = MakeIntent(oSheet, e)
                            Dim placement As Point2d = tg.CreatePoint2d(e.X, hDimLineY)
                            Dim newDim As GeneralDimension =
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    placement, baseIntent, edgeIntent,
                                    DimensionTypeEnum.kHorizontalDimensionType)
                            If newDim IsNot Nothing Then
                                usedX.Add(e.X)
                                e.Dimmed = True
                                countOK += 1
                                If Not CurveHasDim(dimmedCurves, e.Curve) Then dimmedCurves.Add(e.Curve)
                            End If
                        Catch
                            countFail += 1
                        End Try
                    Next
                    '-------------------------------------------------
                    ' 7. DIM CẠNH NGANG + LỖ (Y) — TỪ BASE DỌC
                    '-------------------------------------------------
                    Dim yTargets As List(Of EdgeInfo) =
                        allEdges.Where(Function(e) e.IsHorizontal OrElse e.IsHole).ToList()
                    If vBaseOnBottom Then
                        yTargets = yTargets.OrderBy(Function(e) e.Y).ToList()
                    Else
                        yTargets = yTargets.OrderByDescending(Function(e) e.Y).ToList()
                    End If
                    For Each e As EdgeInfo In yTargets
                        If Math.Abs(e.Y - vBaseCoord) <= EDGE_TOL Then Continue For
                        If ValueExists(usedY, e.Y) Then Continue For
                        If CurveHasDim(dimmedCurves, e.Curve) Then Continue For
                        Try
                            Dim baseIntent As GeometryIntent = oSheet.CreateGeometryIntent(vBaseEdge)
                            Dim edgeIntent As GeometryIntent = MakeIntent(oSheet, e)
                            Dim placement As Point2d = tg.CreatePoint2d(vDimLineX, e.Y)
                            Dim newDim As GeneralDimension =
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    placement, baseIntent, edgeIntent,
                                    DimensionTypeEnum.kVerticalDimensionType)
                            If newDim IsNot Nothing Then
                                usedY.Add(e.Y)
                                e.Dimmed = True
                                countOK += 1
                                If Not CurveHasDim(dimmedCurves, e.Curve) Then dimmedCurves.Add(e.Curve)
                            End If
                        Catch
                            countFail += 1
                        End Try
                    Next
                    '-------------------------------------------------
                    ' 8. QUÉT LẠI
                    '-------------------------------------------------
                    For Each e As EdgeInfo In allEdges
                        ' X-axis
                        If (e.IsVertical OrElse e.IsHole) AndAlso
                           Not ValueExists(usedX, e.X) AndAlso
                           Not CurveHasDim(dimmedCurves, e.Curve) AndAlso
                           Math.Abs(e.X - hBaseCoord) > EDGE_TOL Then
                            Try
                                Dim baseIntent As GeometryIntent = oSheet.CreateGeometryIntent(hBaseEdge)
                                Dim edgeIntent As GeometryIntent = MakeIntent(oSheet, e)
                                Dim placement As Point2d = tg.CreatePoint2d(e.X, hDimLineY - 1.5)
                                Dim newDim As GeneralDimension =
                                    oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                        placement, baseIntent, edgeIntent,
                                        DimensionTypeEnum.kHorizontalDimensionType)
                                If newDim IsNot Nothing Then
                                    usedX.Add(e.X)
                                    e.Dimmed = True
                                    countOK += 1
                                    If Not CurveHasDim(dimmedCurves, e.Curve) Then dimmedCurves.Add(e.Curve)
                                End If
                            Catch
                                countFail += 1
                            End Try
                        End If
                        ' Y-axis
                        If (e.IsHorizontal OrElse e.IsHole) AndAlso
                           Not ValueExists(usedY, e.Y) AndAlso
                           Not CurveHasDim(dimmedCurves, e.Curve) AndAlso
                           Math.Abs(e.Y - vBaseCoord) > EDGE_TOL Then
                            Try
                                Dim baseIntent As GeometryIntent = oSheet.CreateGeometryIntent(vBaseEdge)
                                Dim edgeIntent As GeometryIntent = MakeIntent(oSheet, e)
                                Dim placement As Point2d = tg.CreatePoint2d(vDimLineX - 1.5, e.Y)
                                Dim newDim As GeneralDimension =
                                    oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                        placement, baseIntent, edgeIntent,
                                        DimensionTypeEnum.kVerticalDimensionType)
                                If newDim IsNot Nothing Then
                                    usedY.Add(e.Y)
                                    e.Dimmed = True
                                    countOK += 1
                                    If Not CurveHasDim(dimmedCurves, e.Curve) Then dimmedCurves.Add(e.Curve)
                                End If
                            Catch
                                countFail += 1
                            End Try
                        End If
                    Next
                Next ' End For Each selectedViews
                Try
                    ArrangeDimensions(oSheet, app)
                Catch
                End Try
                oDrawDoc.Update()
                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "View: " & selectedViews.Count & vbCrLf &
                    "Dim tạo: " & countOK & vbCrLf &
                    "Lỗi: " & countFail,
                    "Dim Baseline cạnh + lỗ",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "Dim Baseline cạnh + lỗ",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
        '=============================================================
        ' TẠO GeometryIntent phù hợp
        '=============================================================
        Private Function MakeIntent(ByVal oSheet As Sheet,
                                    ByVal e As EdgeInfo) As GeometryIntent
            If e Is Nothing OrElse e.Curve Is Nothing Then Return Nothing
            If e.IsHole Then
                Return oSheet.CreateGeometryIntent(e.Curve, PointIntentEnum.kCenterPointIntent)
            Else
                Return oSheet.CreateGeometryIntent(e.Curve)
            End If
        End Function
        '=============================================================
        ' KIỂM TRA CẠNH CÓ THUỘC VIEW
        '=============================================================
        Private Function EdgeInView(ByVal edge As DrawingCurve,
                                    ByVal view As DrawingView) As Boolean
            If edge Is Nothing OrElse view Is Nothing Then Return False
            Try
                For Each c As DrawingCurve In view.DrawingCurves
                    If c Is edge Then Return True
                Next
            Catch
            End Try
            Return False
        End Function
        '=============================================================
        ' FORM CHỌN HƯỚNG + PICK CẠNH BASE + TÙY CHỌN LỖ
        '=============================================================
        Private Function ShowDirectionDialog(
            ByRef hFromRight As Boolean,
            ByRef vFromBottom As Boolean,
            ByRef hPickedEdge As DrawingCurve,
            ByRef vPickedEdge As DrawingCurve,
            ByRef includeHoles As Boolean,
            ByVal app As Inventor.Application) As Boolean
            Dim frm As New Form()
            frm.Text = "Hướng chuẩn Baseline"
            frm.ClientSize = New Size(400, 350)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            Dim hHolder As New EdgePickHolder()
            Dim vHolder As New EdgePickHolder()
            Dim okClicked As Boolean = False
            '----------- Group NGANG -----------
            Dim gbH As New GroupBox()
            gbH.Text = "Hướng chuẩn NGANG (X)"
            gbH.SetBounds(12, 12, 376, 120)
            frm.Controls.Add(gbH)
            Dim rbHLeft As New RadioButton()
            rbHLeft.Text = "Từ TRÁI  →  Phải   (mặc định - cạnh trái ngoài)"
            rbHLeft.SetBounds(15, 22, 350, 20)
            rbHLeft.Checked = True
            gbH.Controls.Add(rbHLeft)
            Dim rbHRight As New RadioButton()
            rbHRight.Text = "Từ PHẢI  →  Trái   (cạnh phải ngoài)"
            rbHRight.SetBounds(15, 46, 350, 20)
            gbH.Controls.Add(rbHRight)
            Dim lblH As New Label()
            lblH.SetBounds(15, 74, 240, 20)
            lblH.Text = "(chưa pick cạnh cụ thể)"
            lblH.ForeColor = System.Drawing.Color.Gray
            lblH.Font = New Font("Segoe UI", 8, FontStyle.Italic)
            gbH.Controls.Add(lblH)
            Dim btnPickH As New Button()
            btnPickH.Text = "Pick cạnh X..."
            btnPickH.SetBounds(260, 72, 105, 26)
            gbH.Controls.Add(btnPickH)
            '----------- Group DỌC -----------
            Dim gbV As New GroupBox()
            gbV.Text = "Hướng chuẩn DỌC (Y)"
            gbV.SetBounds(12, 140, 376, 120)
            frm.Controls.Add(gbV)
            Dim rbVTop As New RadioButton()
            rbVTop.Text = "Từ TRÊN  →  Dưới  (mặc định - cạnh trên ngoài)"
            rbVTop.SetBounds(15, 22, 350, 20)
            rbVTop.Checked = True
            gbV.Controls.Add(rbVTop)
            Dim rbVBottom As New RadioButton()
            rbVBottom.Text = "Từ DƯỚI  →  Trên  (cạnh dưới ngoài)"
            rbVBottom.SetBounds(15, 46, 350, 20)
            gbV.Controls.Add(rbVBottom)
            Dim lblV As New Label()
            lblV.SetBounds(15, 74, 240, 20)
            lblV.Text = "(chưa pick cạnh cụ thể)"
            lblV.ForeColor = System.Drawing.Color.Gray
            lblV.Font = New Font("Segoe UI", 8, FontStyle.Italic)
            gbV.Controls.Add(lblV)
            Dim btnPickV As New Button()
            btnPickV.Text = "Pick cạnh Y..."
            btnPickV.SetBounds(260, 72, 105, 26)
            gbV.Controls.Add(btnPickV)
            '----------- Checkbox -----------
            Dim chkHoles As New CheckBox()
            chkHoles.Text = "Bao gồm LỖ TRÒN (đo theo tâm lỗ)"
            chkHoles.SetBounds(15, 270, 376, 22)
            chkHoles.Checked = True
            frm.Controls.Add(chkHoles)
            '----------- Buttons -----------
            Dim btnOK As New Button()
            btnOK.Text = "OK"
            btnOK.SetBounds(200, 305, 85, 28)
            frm.Controls.Add(btnOK)
            Dim btnCancel As New Button()
            btnCancel.Text = "Hủy"
            btnCancel.SetBounds(295, 305, 85, 28)
            frm.Controls.Add(btnCancel)
            '----------- Handlers -----------
            AddHandler btnOK.Click, Sub()
                                        okClicked = True
                                        frm.Close()
                                    End Sub
            AddHandler btnCancel.Click, Sub()
                                            okClicked = False
                                            frm.Close()
                                        End Sub
            AddHandler btnPickH.Click, Sub()
                                           frm.Hide()
                                           System.Windows.Forms.Application.DoEvents()
                                           Try
                                               Dim seg As DrawingCurveSegment = TryCast(app.CommandManager.Pick(
                                                   SelectionFilterEnum.kDrawingCurveSegmentFilter,
                                                   "Chọn cạnh ĐỨNG làm gốc đo X"), DrawingCurveSegment)
                                               If seg IsNot Nothing Then
                                                   Dim c As DrawingCurve = seg.Parent
                                                   Dim dx As Double = Math.Abs(c.StartPoint.X - c.EndPoint.X)
                                                   Dim dy As Double = Math.Abs(c.StartPoint.Y - c.EndPoint.Y)
                                                   If dx < EDGE_TOL AndAlso dy > EDGE_TOL Then
                                                       hHolder.Edge = c
                                                       lblH.Text = "✔ Đã pick cạnh đứng"
                                                       lblH.ForeColor = System.Drawing.Color.Green
                                                   Else
                                                       hHolder.Edge = Nothing
                                                       lblH.Text = "✘ Cạnh này không phải cạnh ĐỨNG"
                                                       lblH.ForeColor = System.Drawing.Color.Red
                                                   End If
                                               End If
                                           Catch
                                           End Try
                                           frm.Show()
                                           frm.Activate()
                                       End Sub
            AddHandler btnPickV.Click, Sub()
                                           frm.Hide()
                                           System.Windows.Forms.Application.DoEvents()
                                           Try
                                               Dim seg As DrawingCurveSegment = TryCast(app.CommandManager.Pick(
                                                   SelectionFilterEnum.kDrawingCurveSegmentFilter,
                                                   "Chọn cạnh NGANG làm gốc đo Y"), DrawingCurveSegment)
                                               If seg IsNot Nothing Then
                                                   Dim c As DrawingCurve = seg.Parent
                                                   Dim dx As Double = Math.Abs(c.StartPoint.X - c.EndPoint.X)
                                                   Dim dy As Double = Math.Abs(c.StartPoint.Y - c.EndPoint.Y)
                                                   If dy < EDGE_TOL AndAlso dx > EDGE_TOL Then
                                                       vHolder.Edge = c
                                                       lblV.Text = "✔ Đã pick cạnh ngang"
                                                       lblV.ForeColor = System.Drawing.Color.Green
                                                   Else
                                                       vHolder.Edge = Nothing
                                                       lblV.Text = "✘ Cạnh này không phải cạnh NGANG"
                                                       lblV.ForeColor = System.Drawing.Color.Red
                                                   End If
                                               End If
                                           Catch
                                           End Try
                                           frm.Show()
                                           frm.Activate()
                                       End Sub
            '----------- MODELESS -----------
            frm.Show()
            Do While frm.Visible
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(15)
            Loop
            If Not okClicked Then Return False
            hFromRight = rbHRight.Checked
            vFromBottom = rbVBottom.Checked
            hPickedEdge = hHolder.Edge
            vPickedEdge = vHolder.Edge
            includeHoles = chkHoles.Checked
            Return True
        End Function
        '=============================================================
        ' ARRANGE
        '=============================================================
        Private Sub ArrangeDimensions(ByVal oSheet As Sheet, ByVal app As Inventor.Application)
            Try
                Dim oDims As DrawingDimensions = oSheet.DrawingDimensions
                If oDims Is Nothing OrElse oDims.Count = 0 Then Exit Sub
                Dim oCol As ObjectCollection = app.TransientObjects.CreateObjectCollection
                For Each oDim As DrawingDimension In oDims
                    Try
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
                If oCol.Count > 1 Then oDims.Arrange(oCol)
            Catch
            End Try
        End Sub
        '=============================================================
        Private Function ValueExists(ByVal list As List(Of Double),
                                     ByVal value As Double) As Boolean
            If list Is Nothing Then Return False
            For Each v As Double In list
                If Math.Abs(v - value) <= EDGE_TOL Then Return True
            Next
            Return False
        End Function
        '=============================================================
        ' So sánh 2 cạnh theo hình học
        '=============================================================
        Private Function SameSegment(ByVal c1 As DrawingCurve,
                                     ByVal c2 As DrawingCurve) As Boolean
            If c1 Is Nothing OrElse c2 Is Nothing Then Return False
            Try
                Dim a1 As Point2d = c1.StartPoint
                Dim a2 As Point2d = c1.EndPoint
                Dim b1 As Point2d = c2.StartPoint
                Dim b2 As Point2d = c2.EndPoint
                If a1 Is Nothing OrElse a2 Is Nothing OrElse
                   b1 Is Nothing OrElse b2 Is Nothing Then Return False
                Return (Math.Abs(a1.X - b1.X) < EDGE_TOL AndAlso
                        Math.Abs(a1.Y - b1.Y) < EDGE_TOL AndAlso
                        Math.Abs(a2.X - b2.X) < EDGE_TOL AndAlso
                        Math.Abs(a2.Y - b2.Y) < EDGE_TOL) OrElse
                       (Math.Abs(a1.X - b2.X) < EDGE_TOL AndAlso
                        Math.Abs(a1.Y - b2.Y) < EDGE_TOL AndAlso
                        Math.Abs(a2.X - b1.X) < EDGE_TOL AndAlso
                        Math.Abs(a2.Y - b1.Y) < EDGE_TOL)
            Catch
            End Try
            Return False
        End Function
        '=============================================================
        ' So sánh 2 lỗ theo tâm
        '=============================================================
        Private Function SameCircle(ByVal c1 As DrawingCurve,
                                    ByVal c2 As DrawingCurve) As Boolean
            If c1 Is Nothing OrElse c2 Is Nothing Then Return False
            Try
                Dim p1 As Point2d = c1.CenterPoint
                Dim p2 As Point2d = c2.CenterPoint
                If p1 Is Nothing OrElse p2 Is Nothing Then Return False
                Return Math.Abs(p1.X - p2.X) <= EDGE_TOL AndAlso
                       Math.Abs(p1.Y - p2.Y) <= EDGE_TOL
            Catch
            End Try
            Return False
        End Function
        '=============================================================
        ' THU THẬP CURVE ĐÃ CÓ DIM
        '=============================================================
        Private Function CollectDimmedCurves(ByVal oSheet As Sheet) As List(Of DrawingCurve)
            Dim result As New List(Of DrawingCurve)
            Try
                For Each oDim As DrawingDimension In oSheet.DrawingDimensions
                    Try
                        Dim atts As ObjectCollection = oDim.AttachedEntities
                        If atts Is Nothing Then Continue For
                        For i As Integer = 1 To atts.Count
                            Dim ent As Object = atts.Item(i)
                            If ent Is Nothing Then Continue For
                            Dim gi As GeometryIntent = TryCast(ent, GeometryIntent)
                            If gi IsNot Nothing Then
                                Dim dc As DrawingCurve = TryCast(gi.Geometry, DrawingCurve)
                                If dc IsNot Nothing Then
                                    result.Add(dc)
                                    Continue For
                                End If
                            End If
                            Dim dc2 As DrawingCurve = TryCast(ent, DrawingCurve)
                            If dc2 IsNot Nothing Then result.Add(dc2) : Continue For
                            Dim seg As DrawingCurveSegment = TryCast(ent, DrawingCurveSegment)
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
        ' KIỂM TRA CURVE ĐÃ CÓ DIM
        '=============================================================
        Private Function CurveHasDim(ByVal dimmed As List(Of DrawingCurve),
                                     ByVal c As DrawingCurve) As Boolean
            If dimmed Is Nothing OrElse c Is Nothing Then Return False
            For Each dc As DrawingCurve In dimmed
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
        ' HOLDER cho lambda
        '=============================================================
        Public Class EdgePickHolder
            Public Edge As DrawingCurve
        End Class
    End Module
End Namespace