Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq

Namespace ToolInventor2025.Drawing.Buttons.Drawdim
    Public Module Draw_dim_chain_line_a

        Private Const TOL As Double = 0.03
        Private Const CHAIN_GAP As Double = 3.2
        Private Const TOTAL_GAP As Double = 12.0

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

                '----- CHỌN VIEW -----
                Dim selectedViews As New List(Of DrawingView)
                Do
                    Dim oSS As SelectSet = oDrawDoc.SelectSet
                    oSS.Clear()

                    Dim oView As DrawingView = Nothing
                    Try
                        oView = CType(app.CommandManager.Pick(
                            SelectionFilterEnum.kDrawingViewFilter,
                            "Chọn View (Esc/Right-click kết thúc)"), DrawingView)
                    Catch
                        Exit Do
                    End Try

                    If oView Is Nothing Then Exit Do
                    If Not selectedViews.Contains(oView) Then selectedViews.Add(oView)
                Loop

                If selectedViews.Count = 0 Then Exit Sub

                '----- FORM -----
                Dim chainTop As Boolean = True
                Dim chainLeft As Boolean = True
                Dim addTotal As Boolean = True

                If Not ShowChainForm(chainTop, chainLeft, addTotal) Then Exit Sub

                Dim nChain As Integer = 0
                Dim nTotal As Integer = 0
                Dim nFail As Integer = 0

                For Each oView As DrawingView In selectedViews
                    '===== 1. LẤY CẠNH, GỘP TRÙNG =====
                    Dim xs As New List(Of Double)
                    Dim ys As New List(Of Double)
                    Dim xCurve As New Dictionary(Of Double, DrawingCurve)
                    Dim yCurve As New Dictionary(Of Double, DrawingCurve)

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

                            Dim dx As Double = Math.Abs(p1.X - p2.X)
                            Dim dy As Double = Math.Abs(p1.Y - p2.Y)
                            If dx < TOL AndAlso dy < TOL Then Continue For

                            If dx < TOL Then
                                Dim xv As Double = (p1.X + p2.X) / 2.0
                                If p1.X < minX Then minX = p1.X
                                If p1.X > maxX Then maxX = p1.X

                                If Not HasNear(xs, xv) Then
                                    xs.Add(xv)
                                    xCurve(xv) = oCurve
                                End If

                            ElseIf dy < TOL Then
                                Dim yv As Double = (p1.Y + p2.Y) / 2.0
                                If p1.Y < minY Then minY = p1.Y
                                If p1.Y > maxY Then maxY = p1.Y

                                If Not HasNear(ys, yv) Then
                                    ys.Add(yv)
                                    yCurve(yv) = oCurve
                                End If
                            End If
                        Catch
                        End Try
                    Next

                    If xs.Count < 2 AndAlso ys.Count < 2 Then Continue For

                    xs.Sort()
                    ys.Sort()

                    '===== 2. VỊ TRÍ DIM LINE =====
                    Dim chainY As Double = If(chainTop, maxY + CHAIN_GAP, minY - CHAIN_GAP)
                    Dim totalY As Double = If(chainTop, maxY + TOTAL_GAP, minY - TOTAL_GAP)
                    Dim chainX As Double = If(chainLeft, minX - CHAIN_GAP, maxX + CHAIN_GAP)
                    Dim totalX As Double = If(chainLeft, minX - TOTAL_GAP, maxX + TOTAL_GAP)

                    '===== 3. CHAIN DIM NGANG =====
                    If xs.Count >= 2 Then
                        For i As Integer = 0 To xs.Count - 2
                            Dim x1 As Double = xs(i)
                            Dim x2 As Double = xs(i + 1)
                            If Math.Abs(x2 - x1) <= TOL Then Continue For

                            Try
                                Dim placement As Point2d =
                                    tg.CreatePoint2d((x1 + x2) / 2.0, chainY)

                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    placement,
                                    oSheet.CreateGeometryIntent(xCurve(x1)),
                                    oSheet.CreateGeometryIntent(xCurve(x2)),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                nChain += 1
                            Catch
                                nFail += 1
                            End Try
                        Next

                        If addTotal Then
                            Try
                                Dim placement As Point2d =
                                    tg.CreatePoint2d((minX + maxX) / 2.0, totalY)

                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    placement,
                                    oSheet.CreateGeometryIntent(xCurve(xs.First())),
                                    oSheet.CreateGeometryIntent(xCurve(xs.Last())),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                nTotal += 1
                            Catch
                                nFail += 1
                            End Try
                        End If
                    End If

                    '===== 4. CHAIN DIM DỌC =====
                    If ys.Count >= 2 Then
                        For i As Integer = 0 To ys.Count - 2
                            Dim y1 As Double = ys(i)
                            Dim y2 As Double = ys(i + 1)
                            If Math.Abs(y2 - y1) <= TOL Then Continue For

                            Try
                                Dim placement As Point2d =
                                    tg.CreatePoint2d(chainX, (y1 + y2) / 2.0)

                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    placement,
                                    oSheet.CreateGeometryIntent(yCurve(y1)),
                                    oSheet.CreateGeometryIntent(yCurve(y2)),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                nChain += 1
                            Catch
                                nFail += 1
                            End Try
                        Next

                        If addTotal Then
                            Try
                                Dim placement As Point2d =
                                    tg.CreatePoint2d(totalX, (minY + maxY) / 2.0)

                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    placement,
                                    oSheet.CreateGeometryIntent(yCurve(ys.First())),
                                    oSheet.CreateGeometryIntent(yCurve(ys.Last())),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                nTotal += 1
                            Catch
                                nFail += 1
                            End Try
                        End If
                    End If
                Next

                '===== ARRANGE — CHỈ DIM THUỘC VIEW ĐÃ CHỌN =====
                Try
                    Dim oDims As DrawingDimensions = oSheet.DrawingDimensions
                    Dim col As ObjectCollection = app.TransientObjects.CreateObjectCollection

                    For Each d As DrawingDimension In oDims
                        Try
                            If Not IsDimInAnyView(d, selectedViews) Then Continue For
                            d.CenterText()
                            col.Add(d)
                        Catch
                        End Try
                    Next

                    If col.Count > 1 Then oDims.Arrange(col)
                Catch
                End Try

                oDrawDoc.Update()

                MessageBox.Show(
                    "Chain hoàn tất!" & vbCrLf & vbCrLf &
                    "Chain dim: " & nChain & vbCrLf &
                    "Dim tổng: " & nTotal & vbCrLf &
                    "Lỗi: " & nFail,
                    "Chain Line cạnh",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "Chain Line cạnh",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        '=============================================================
        Private Function HasNear(ByVal list As List(Of Double),
                                 ByVal value As Double) As Boolean
            For Each v As Double In list
                If Math.Abs(v - value) <= TOL Then Return True
            Next
            Return False
        End Function

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
        Private Function ShowChainForm(ByRef chainTop As Boolean,
                                       ByRef chainLeft As Boolean,
                                       ByRef addTotal As Boolean) As Boolean

            Dim frm As New Form()
            frm.Text = "Chain Line — Hướng chuẩn"
            frm.ClientSize = New Size(400, 340)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False

            '----------- Group NGANG (X) -----------
            Dim gbH As New GroupBox()
            gbH.Text = "Hướng chuẩn NGANG (X)"
            gbH.SetBounds(12, 12, 376, 120)
            frm.Controls.Add(gbH)

            Dim rbHUp As New RadioButton()
            rbHUp.Text = "Đặt chain PHÍA TRÊN chi tiết"
            rbHUp.SetBounds(15, 22, 350, 20)
            rbHUp.Checked = True
            gbH.Controls.Add(rbHUp)

            Dim rbHDown As New RadioButton()
            rbHDown.Text = "Đặt chain PHÍA DƯỚI chi tiết"
            rbHDown.SetBounds(15, 46, 350, 20)
            gbH.Controls.Add(rbHDown)

            Dim lblH As New Label()
            lblH.SetBounds(15, 74, 350, 34)
            lblH.Text = "Chain gồm các dim liên tiếp giữa các cạnh đứng." & vbCrLf &
                        "Dim tổng (nếu bật) đặt xa hơn ở trên/dưới."
            lblH.ForeColor = System.Drawing.Color.Gray
            lblH.Font = New Font("Segoe UI", 8, FontStyle.Italic)
            gbH.Controls.Add(lblH)

            '----------- Group DỌC (Y) -----------
            Dim gbV As New GroupBox()
            gbV.Text = "Hướng chuẩn DỌC (Y)"
            gbV.SetBounds(12, 140, 376, 120)
            frm.Controls.Add(gbV)

            Dim rbVLeft As New RadioButton()
            rbVLeft.Text = "Đặt chain BÊN TRÁI chi tiết"
            rbVLeft.SetBounds(15, 22, 350, 20)
            rbVLeft.Checked = True
            gbV.Controls.Add(rbVLeft)

            Dim rbVRight As New RadioButton()
            rbVRight.Text = "Đặt chain BÊN PHẢI chi tiết"
            rbVRight.SetBounds(15, 46, 350, 20)
            gbV.Controls.Add(rbVRight)

            Dim lblV As New Label()
            lblV.SetBounds(15, 74, 350, 34)
            lblV.Text = "Chain gồm các dim liên tiếp giữa các cạnh ngang." & vbCrLf &
                        "Dim tổng (nếu bật) đặt xa hơn ở trái/phải."
            lblV.ForeColor = System.Drawing.Color.Gray
            lblV.Font = New Font("Segoe UI", 8, FontStyle.Italic)
            gbV.Controls.Add(lblV)

            '----------- Checkbox dim tổng -----------
            Dim chkTotal As New CheckBox()
            chkTotal.Text = "Thêm dim TỔNG bao ngoài (ví dụ: 584,00)"
            chkTotal.SetBounds(15, 268, 376, 22)
            chkTotal.Checked = True
            frm.Controls.Add(chkTotal)

            '----------- Buttons -----------
            Dim btnOK As New Button()
            btnOK.Text = "OK"
            btnOK.SetBounds(200, 300, 85, 28)
            btnOK.DialogResult = DialogResult.OK
            frm.Controls.Add(btnOK)

            Dim btnCancel As New Button()
            btnCancel.Text = "Hủy"
            btnCancel.SetBounds(295, 300, 85, 28)
            btnCancel.DialogResult = DialogResult.Cancel
            frm.Controls.Add(btnCancel)

            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            If frm.ShowDialog() <> DialogResult.OK Then Return False

            chainTop = rbHUp.Checked
            chainLeft = rbVLeft.Checked
            addTotal = chkTotal.Checked
            Return True
        End Function

    End Module
End Namespace