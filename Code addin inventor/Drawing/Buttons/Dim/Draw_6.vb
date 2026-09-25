Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.Drawdim

    Public Module Dim_Linear_Total

        '=========================================================
        ' OPTIONS
        '=========================================================
        Private Class DimOptions
            Public OffsetMM As Double = 5.5
            Public OffsetCM As Double = 0.55
            Public DoWidth As Boolean = True
            Public DoHeight As Boolean = True
            Public Cancelled As Boolean = True
        End Class


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
                ' 1. FORM TÙY CHỌN
                '=====================================================
                Dim opt As DimOptions = ShowOptionsForm()
                If opt Is Nothing OrElse opt.Cancelled Then Exit Sub

                '=====================================================
                ' 2. CHỌN NHIỀU VIEW (ESC kết thúc)
                '=====================================================
                Dim selectedViews As New List(Of DrawingView)

                Try : oDrawDoc.SelectSet.Clear() : Catch : End Try

                While True
                    Dim prompt As String = "Chọn Drawing View (" &
                                           selectedViews.Count.ToString() &
                                           " đã chọn). ESC để kết thúc."

                    Dim pickedObj As Object = Nothing
                    Try
                        pickedObj = app.CommandManager.Pick(
                            SelectionFilterEnum.kDrawingViewFilter, prompt)
                    Catch
                        Exit While
                    End Try

                    If pickedObj Is Nothing Then Exit While

                    Dim v As DrawingView = TryCast(pickedObj, DrawingView)
                    If v Is Nothing Then Continue While

                    Dim dup As Boolean = False
                    For Each x As DrawingView In selectedViews
                        If x Is v Then
                            dup = True
                            Exit For
                        End If
                    Next
                    If Not dup Then selectedViews.Add(v)
                End While

                If selectedViews.Count = 0 Then
                    MessageBox.Show("Chưa chọn view nào.", "Thông báo")
                    Exit Sub
                End If

                '=====================================================
                ' 3. XỬ LÝ TỪNG VIEW
                '=====================================================
                Dim nOK As Integer = 0
                Dim nFail As Integer = 0
                Dim log As New System.Text.StringBuilder()

                For Each oView As DrawingView In selectedViews
                    Try
                        ProcessOneView(oSheet, tg, oView, opt, nOK, nFail, log)
                    Catch ex As Exception
                        nFail += 1
                        log.AppendLine("  ✘ " & oView.Name & ": " & ex.Message)
                    End Try
                Next

                Try : oDrawDoc.Update() : Catch : End Try

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Số view      : " & selectedViews.Count & vbCrLf &
                    "Dim tạo      : " & nOK & vbCrLf &
                    "Lỗi / bỏ qua : " & nFail & vbCrLf & vbCrLf &
                    "--- Chi tiết ---" & vbCrLf &
                    log.ToString(),
                    "Dim Linear tổng",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "Dim Linear tổng",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '=========================================================
        ' XỬ LÝ 1 VIEW
        '=========================================================
        Private Sub ProcessOneView(ByVal oSheet As Sheet,
                                   ByVal tg As TransientGeometry,
                                   ByVal oView As DrawingView,
                                   ByVal opt As DimOptions,
                                   ByRef nOK As Integer,
                                   ByRef nFail As Integer,
                                   ByVal log As System.Text.StringBuilder)

            '=====================================================
            ' TÌM BOUNDING BOX TỪ TẤT CẢ CURVES
            '=====================================================
            Dim minX As Double = Double.MaxValue
            Dim maxX As Double = Double.MinValue
            Dim minY As Double = Double.MaxValue
            Dim maxY As Double = Double.MinValue

            Dim leftCurve As DrawingCurve = Nothing
            Dim rightCurve As DrawingCurve = Nothing
            Dim topCurve As DrawingCurve = Nothing
            Dim bottomCurve As DrawingCurve = Nothing

            Dim leftPt As Point2d = Nothing
            Dim rightPt As Point2d = Nothing
            Dim topPt As Point2d = Nothing
            Dim bottomPt As Point2d = Nothing

            Dim curveCount As Integer = 0

            For Each oCurve As DrawingCurve In oView.DrawingCurves
                Try
                    curveCount += 1

                    Dim cMinX As Double = 0
                    Dim cMinY As Double = 0
                    Dim cMaxX As Double = 0
                    Dim cMaxY As Double = 0
                    Dim cLeftPt As Point2d = Nothing
                    Dim cRightPt As Point2d = Nothing
                    Dim cTopPt As Point2d = Nothing
                    Dim cBottomPt As Point2d = Nothing

                    If Not GetCurveBounds(tg, oCurve,
                                          cMinX, cMinY, cMaxX, cMaxY,
                                          cLeftPt, cRightPt, cTopPt, cBottomPt) Then
                        Continue For
                    End If

                    '--- Cập nhật global ---
                    If cMinX < minX Then
                        minX = cMinX
                        leftCurve = oCurve
                        leftPt = cLeftPt
                    End If
                    If cMaxX > maxX Then
                        maxX = cMaxX
                        rightCurve = oCurve
                        rightPt = cRightPt
                    End If
                    If cMaxY > maxY Then
                        maxY = cMaxY
                        topCurve = oCurve
                        topPt = cTopPt
                    End If
                    If cMinY < minY Then
                        minY = cMinY
                        bottomCurve = oCurve
                        bottomPt = cBottomPt
                    End If

                Catch
                End Try
            Next

            '=====================================================
            ' KIỂM TRA
            '=====================================================
            If curveCount = 0 Then
                nFail += 1
                log.AppendLine("  ✘ " & oView.Name & ": view rỗng")
                Return
            End If

            If leftCurve Is Nothing OrElse rightCurve Is Nothing OrElse
               topCurve Is Nothing OrElse bottomCurve Is Nothing Then
                nFail += 1
                log.AppendLine("  ✘ " & oView.Name & ": không tìm đủ 4 cạnh")
                Return
            End If

            If (maxX - minX) <= 0 OrElse (maxY - minY) <= 0 Then
                nFail += 1
                log.AppendLine("  ✘ " & oView.Name & ": kích thước không hợp lệ")
                Return
            End If

            '=====================================================
            ' DIM NGANG (W)
            '=====================================================
            If opt.DoWidth Then
                Try
                    Dim tp As Point2d = tg.CreatePoint2d((minX + maxX) / 2,
                                                         maxY + opt.OffsetCM)

                    Dim intentL As GeometryIntent = Nothing
                    Dim intentR As GeometryIntent = Nothing

                    ' Ưu tiên Point2d chính xác
                    If leftPt IsNot Nothing Then
                        Try : intentL = oSheet.CreateGeometryIntent(leftCurve, leftPt) : Catch : End Try
                    End If
                    If intentL Is Nothing Then
                        Try : intentL = oSheet.CreateGeometryIntent(leftCurve) : Catch : End Try
                    End If

                    If rightPt IsNot Nothing Then
                        Try : intentR = oSheet.CreateGeometryIntent(rightCurve, rightPt) : Catch : End Try
                    End If
                    If intentR Is Nothing Then
                        Try : intentR = oSheet.CreateGeometryIntent(rightCurve) : Catch : End Try
                    End If

                    If intentL Is Nothing OrElse intentR Is Nothing Then
                        Throw New Exception("Không tạo được GeometryIntent cho W")
                    End If

                    oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                        tp, intentL, intentR,
                        DimensionTypeEnum.kHorizontalDimensionType)

                    nOK += 1
                    log.AppendLine("  ✔ " & oView.Name &
                                   " W = " & ((maxX - minX) * 10.0).ToString("0.0#") & " mm")

                Catch ex As Exception
                    nFail += 1
                    log.AppendLine("  ✘ " & oView.Name & " W: " & ex.Message)
                End Try
            End If

            '=====================================================
            ' DIM DỌC (H)
            '=====================================================
            If opt.DoHeight Then
                Try
                    Dim tp As Point2d = tg.CreatePoint2d(minX - opt.OffsetCM,
                                                         (maxY + minY) / 2)

                    Dim intentT As GeometryIntent = Nothing
                    Dim intentB As GeometryIntent = Nothing

                    If topPt IsNot Nothing Then
                        Try : intentT = oSheet.CreateGeometryIntent(topCurve, topPt) : Catch : End Try
                    End If
                    If intentT Is Nothing Then
                        Try : intentT = oSheet.CreateGeometryIntent(topCurve) : Catch : End Try
                    End If

                    If bottomPt IsNot Nothing Then
                        Try : intentB = oSheet.CreateGeometryIntent(bottomCurve, bottomPt) : Catch : End Try
                    End If
                    If intentB Is Nothing Then
                        Try : intentB = oSheet.CreateGeometryIntent(bottomCurve) : Catch : End Try
                    End If

                    If intentT Is Nothing OrElse intentB Is Nothing Then
                        Throw New Exception("Không tạo được GeometryIntent cho H")
                    End If

                    oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                        tp, intentT, intentB,
                        DimensionTypeEnum.kVerticalDimensionType)

                    nOK += 1
                    log.AppendLine("  ✔ " & oView.Name &
                                   " H = " & ((maxY - minY) * 10.0).ToString("0.0#") & " mm")

                Catch ex As Exception
                    nFail += 1
                    log.AppendLine("  ✘ " & oView.Name & " H: " & ex.Message)
                End Try
            End If

        End Sub


        '=========================================================
        ' TÍNH BOUNDING BOX CỦA 1 CURVE — HỖ TRỢ LINE + ARC + CIRCLE
        '=========================================================
        Private Function GetCurveBounds(ByVal tg As TransientGeometry,
                                        ByVal curve As DrawingCurve,
                                        ByRef cMinX As Double, ByRef cMinY As Double,
                                        ByRef cMaxX As Double, ByRef cMaxY As Double,
                                        ByRef leftPt As Point2d, ByRef rightPt As Point2d,
                                        ByRef topPt As Point2d, ByRef bottomPt As Point2d) As Boolean

            cMinX = Double.MaxValue
            cMinY = Double.MaxValue
            cMaxX = Double.MinValue
            cMaxY = Double.MinValue
            leftPt = Nothing
            rightPt = Nothing
            topPt = Nothing
            bottomPt = Nothing

            Dim hasData As Boolean = False

            Try
                For Each seg As DrawingCurveSegment In curve.Segments
                    Try
                        '=====================================================
                        ' ĐIỂM ĐẦU / CUỐI SEGMENT
                        '=====================================================
                        Dim p1 As Point2d = seg.StartPoint
                        Dim p2 As Point2d = seg.EndPoint

                        If p1 IsNot Nothing Then
                            UpdateBounds(p1, cMinX, cMinY, cMaxX, cMaxY,
                                         leftPt, rightPt, topPt, bottomPt)
                            hasData = True
                        End If

                        If p2 IsNot Nothing Then
                            UpdateBounds(p2, cMinX, cMinY, cMaxX, cMaxY,
                                         leftPt, rightPt, topPt, bottomPt)
                            hasData = True
                        End If

                        Dim geom As Object = seg.Geometry
                        If geom Is Nothing Then Continue For

                        '=====================================================
                        ' ARC 2D — SAMPLE 64 ĐIỂM, TẠO POINT2D THỰC
                        '=====================================================
                        If TypeOf geom Is Arc2d Then
                            Dim arc As Arc2d = CType(geom, Arc2d)
                            Dim cx As Double = arc.Center.X
                            Dim cy As Double = arc.Center.Y
                            Dim r As Double = arc.Radius
                            Dim a0 As Double = arc.StartAngle
                            Dim a1 As Double = arc.EndAngle

                            If a1 < a0 Then a1 += 2 * Math.PI

                            Dim N As Integer = 64
                            For i As Integer = 0 To N
                                Dim t As Double = a0 + (a1 - a0) * i / N
                                Dim px As Double = cx + r * Math.Cos(t)
                                Dim py As Double = cy + r * Math.Sin(t)

                                ' Tạo Point2d thực — không để Nothing
                                Dim samplePt As Point2d = tg.CreatePoint2d(px, py)

                                UpdateBounds(samplePt, cMinX, cMinY, cMaxX, cMaxY,
                                             leftPt, rightPt, topPt, bottomPt)
                            Next

                            hasData = True
                        End If

                        '=====================================================
                        ' CIRCLE 2D — 4 ĐIỂM CỰC TRỊ
                        '=====================================================
                        If TypeOf geom Is Circle2d Then
                            Dim circ As Circle2d = CType(geom, Circle2d)
                            Dim cx As Double = circ.Center.X
                            Dim cy As Double = circ.Center.Y
                            Dim r As Double = circ.Radius

                            UpdateBounds(tg.CreatePoint2d(cx - r, cy),
                                         cMinX, cMinY, cMaxX, cMaxY,
                                         leftPt, rightPt, topPt, bottomPt)
                            UpdateBounds(tg.CreatePoint2d(cx + r, cy),
                                         cMinX, cMinY, cMaxX, cMaxY,
                                         leftPt, rightPt, topPt, bottomPt)
                            UpdateBounds(tg.CreatePoint2d(cx, cy - r),
                                         cMinX, cMinY, cMaxX, cMaxY,
                                         leftPt, rightPt, topPt, bottomPt)
                            UpdateBounds(tg.CreatePoint2d(cx, cy + r),
                                         cMinX, cMinY, cMaxX, cMaxY,
                                         leftPt, rightPt, topPt, bottomPt)

                            hasData = True
                        End If

                    Catch
                    End Try
                Next

            Catch
                Return False
            End Try

            Return hasData

        End Function


        '=========================================================
        ' UPDATE BOUNDS TỪ 1 ĐIỂM
        '=========================================================
        Private Sub UpdateBounds(ByVal pt As Point2d,
                                 ByRef minX As Double, ByRef minY As Double,
                                 ByRef maxX As Double, ByRef maxY As Double,
                                 ByRef leftPt As Point2d, ByRef rightPt As Point2d,
                                 ByRef topPt As Point2d, ByRef bottomPt As Point2d)

            If pt Is Nothing Then Return

            If pt.X < minX Then
                minX = pt.X
                leftPt = pt
            End If

            If pt.X > maxX Then
                maxX = pt.X
                rightPt = pt
            End If

            If pt.Y < minY Then
                minY = pt.Y
                bottomPt = pt
            End If

            If pt.Y > maxY Then
                maxY = pt.Y
                topPt = pt
            End If

        End Sub


        '=========================================================
        ' FORM TÙY CHỌN
        '=========================================================
        Private Function ShowOptionsForm() As DimOptions
            Dim opt As New DimOptions()
            Dim _okConfirmed As Boolean = False

            Using frm As New Form()
                frm.Text = "Dim Linear tổng — Tùy chọn"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(600, 470)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.BackColor = Drw.Color.FromArgb(245, 245, 245)
                frm.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

                '===== HEADER =====
                Dim pnlHeader As New Panel()
                pnlHeader.Location = New Drw.Point(0, 0)
                pnlHeader.Size = New Drw.Size(600, 75)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "DIM LINEAR TỔNG"
                lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Tự động ghi kích thước W × H cho Drawing View (line + arc + circle)"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== GROUP 1: KHOẢNG CÁCH =====
                Dim gb1 As New GroupBox() With {
                    .Text = "1. Khoảng cách đặt dim",
                    .Location = New Drw.Point(15, 90),
                    .Size = New Drw.Size(570, 100),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb1)

                Dim lblOff As New Label() With {
                    .Text = "Offset ngoài view (mm):",
                    .Location = New Drw.Point(25, 40),
                    .Size = New Drw.Size(180, 25),
                    .TextAlign = Drw.ContentAlignment.MiddleLeft,
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(lblOff)

                Dim txtOffset As New System.Windows.Forms.TextBox() With {
                    .Text = "5.5",
                    .Location = New Drw.Point(215, 40),
                    .Size = New Drw.Size(120, 25),
                    .Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(txtOffset)

                '===== GROUP 2: LOẠI DIM =====
                Dim gb2 As New GroupBox() With {
                    .Text = "2. Loại dim cần tạo",
                    .Location = New Drw.Point(15, 200),
                    .Size = New Drw.Size(570, 100),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb2)

                Dim chkW As New CheckBox() With {
                    .Text = "Dim chiều rộng W (ngang) — đặt bên trên",
                    .Location = New Drw.Point(25, 30),
                    .Size = New Drw.Size(530, 25),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb2.Controls.Add(chkW)

                Dim chkH As New CheckBox() With {
                    .Text = "Dim chiều cao H (dọc) — đặt bên trái",
                    .Location = New Drw.Point(25, 60),
                    .Size = New Drw.Size(530, 25),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb2.Controls.Add(chkH)

                '===== GROUP 3: LƯU Ý =====
                Dim gb3 As New GroupBox() With {
                    .Text = "Lưu ý",
                    .Location = New Drw.Point(15, 310),
                    .Size = New Drw.Size(570, 80),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb3)

                Dim lblNote As New Label() With {
                    .Text = "• Hỗ trợ cả line + cung tròn + đường tròn" & vbCrLf &
                            "• Sau khi bấm THỰC HIỆN, click chọn từng View. ESC để kết thúc",
                    .Location = New Drw.Point(20, 22),
                    .Size = New Drw.Size(530, 50),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(80, 80, 80)}
                gb3.Controls.Add(lblNote)

                '===== NÚT THỰC HIỆN =====
                Dim btnOK As New Button()
                btnOK.Text = "THỰC HIỆN"
                btnOK.Size = New Drw.Size(160, 44)
                btnOK.Location = New Drw.Point(420, 405)
                btnOK.FlatStyle = FlatStyle.Flat
                btnOK.FlatAppearance.BorderSize = 0
                btnOK.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(60, 115, 195)
                btnOK.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(30, 80, 155)
                btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
                btnOK.ForeColor = Drw.Color.White
                btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                btnOK.Cursor = Cursors.Hand
                btnOK.UseVisualStyleBackColor = False

                AddHandler btnOK.Click, Sub()
                                            Dim offVal As Double = 5.5
                                            If Not Double.TryParse(txtOffset.Text.Replace(","c, "."c),
                                                                   Globalization.NumberStyles.Any,
                                                                   Globalization.CultureInfo.InvariantCulture,
                                                                   offVal) OrElse offVal <= 0 Then
                                                MessageBox.Show("Khoảng cách không hợp lệ.",
                                                                "Lỗi",
                                                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                Return
                                            End If

                                            If Not chkW.Checked AndAlso Not chkH.Checked Then
                                                MessageBox.Show("Phải chọn ít nhất 1 loại dim.",
                                                                "Lỗi",
                                                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                Return
                                            End If

                                            opt.OffsetMM = offVal
                                            opt.OffsetCM = offVal / 10.0
                                            opt.DoWidth = chkW.Checked
                                            opt.DoHeight = chkH.Checked
                                            opt.Cancelled = False
                                            _okConfirmed = True
                                            frm.DialogResult = DialogResult.OK
                                            frm.Close()
                                        End Sub
                frm.Controls.Add(btnOK)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(120, 44)
                btnCancel.Location = New Drw.Point(285, 405)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.FlatAppearance.BorderSize = 1
                btnCancel.FlatAppearance.BorderColor = Drw.Color.FromArgb(200, 200, 200)
                btnCancel.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(235, 235, 235)
                btnCancel.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(215, 215, 215)
                btnCancel.BackColor = Drw.Color.FromArgb(250, 250, 250)
                btnCancel.ForeColor = Drw.Color.FromArgb(60, 60, 60)
                btnCancel.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                btnCancel.Cursor = Cursors.Hand
                btnCancel.UseVisualStyleBackColor = False
                AddHandler btnCancel.Click, Sub()
                                                opt.Cancelled = True
                                                frm.DialogResult = DialogResult.Cancel
                                                frm.Close()
                                            End Sub
                frm.Controls.Add(btnCancel)

                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel

                AddHandler frm.FormClosing, Sub(sender, e)
                                                If Not _okConfirmed Then opt.Cancelled = True
                                            End Sub

                frm.ShowDialog()
            End Using

            Return opt
        End Function

    End Module

End Namespace