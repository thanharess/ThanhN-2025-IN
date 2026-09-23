Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq

Namespace ToolInventor2025.Drawing.Buttons.Drawdim
    Public Module Draw_dim_chain_line_b

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
                Dim includeHoles As Boolean = True

                If Not ShowChainForm(chainTop, chainLeft, addTotal, includeHoles) Then Exit Sub

                Dim nChain As Integer = 0
                Dim nTotal As Integer = 0
                Dim nFail As Integer = 0

                For Each oView As DrawingView In selectedViews
                    '===== 1. THU THẬP CẠNH + LỖ =====
                    Dim xAnchors As New List(Of AnchorInfo)
                    Dim yAnchors As New List(Of AnchorInfo)

                    Dim minX As Double = Double.MaxValue
                    Dim maxX As Double = Double.MinValue
                    Dim minY As Double = Double.MaxValue
                    Dim maxY As Double = Double.MinValue

                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType = CurveTypeEnum.kLineCurve OrElse
                               oCurve.CurveType = CurveTypeEnum.kLineSegmentCurve Then

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

                                    If Not HasNearAnchor(xAnchors, xv) Then
                                        Dim a As New AnchorInfo
                                        a.Type = AnchorType.Edge
                                        a.Coord = xv
                                        a.Curve = oCurve
                                        xAnchors.Add(a)
                                    End If

                                ElseIf dy < TOL Then
                                    Dim yv As Double = (p1.Y + p2.Y) / 2.0
                                    If p1.Y < minY Then minY = p1.Y
                                    If p1.Y > maxY Then maxY = p1.Y

                                    If Not HasNearAnchor(yAnchors, yv) Then
                                        Dim a As New AnchorInfo
                                        a.Type = AnchorType.Edge
                                        a.Coord = yv
                                        a.Curve = oCurve
                                        yAnchors.Add(a)
                                    End If
                                End If

                            ElseIf includeHoles AndAlso
                                   oCurve.CurveType = CurveTypeEnum.kCircleCurve Then

                                Dim c As Point2d = oCurve.CenterPoint
                                If c Is Nothing Then Continue For

                                If c.X < minX Then minX = c.X
                                If c.X > maxX Then maxX = c.X
                                If c.Y < minY Then minY = c.Y
                                If c.Y > maxY Then maxY = c.Y

                                If Not HasNearAnchor(xAnchors, c.X) Then
                                    Dim ax As New AnchorInfo
                                    ax.Type = AnchorType.Hole
                                    ax.Coord = c.X
                                    ax.Curve = oCurve
                                    ax.Center = c
                                    xAnchors.Add(ax)
                                End If

                                If Not HasNearAnchor(yAnchors, c.Y) Then
                                    Dim ay As New AnchorInfo
                                    ay.Type = AnchorType.Hole
                                    ay.Coord = c.Y
                                    ay.Curve = oCurve
                                    ay.Center = c
                                    yAnchors.Add(ay)
                                End If
                            End If
                        Catch
                        End Try
                    Next

                    If xAnchors.Count < 2 AndAlso yAnchors.Count < 2 Then Continue For

                    '===== 2. SẮP XẾP =====
                    xAnchors = xAnchors.OrderBy(Function(a) a.Coord).ToList()
                    yAnchors = yAnchors.OrderBy(Function(a) a.Coord).ToList()

                    '===== 3. VỊ TRÍ DIM LINE =====
                    Dim chainY As Double = If(chainTop, maxY + CHAIN_GAP, minY - CHAIN_GAP)
                    Dim totalY As Double = If(chainTop, maxY + TOTAL_GAP, minY - TOTAL_GAP)
                    Dim chainX As Double = If(chainLeft, minX - CHAIN_GAP, maxX + CHAIN_GAP)
                    Dim totalX As Double = If(chainLeft, minX - TOTAL_GAP, maxX + TOTAL_GAP)

                    '===== 4. CHAIN DIM NGANG =====
                    If xAnchors.Count >= 2 Then
                        For i As Integer = 0 To xAnchors.Count - 2
                            Dim a1 As AnchorInfo = xAnchors(i)
                            Dim a2 As AnchorInfo = xAnchors(i + 1)
                            If Math.Abs(a2.Coord - a1.Coord) <= TOL Then Continue For

                            Try
                                Dim placement As Point2d =
                                    tg.CreatePoint2d((a1.Coord + a2.Coord) / 2.0, chainY)

                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    placement,
                                    MakeIntent(oSheet, a1),
                                    MakeIntent(oSheet, a2),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                nChain += 1
                            Catch
                                nFail += 1
                            End Try
                        Next

                        If addTotal Then
                            Try
                                Dim aFirst As AnchorInfo = xAnchors.First()
                                Dim aLast As AnchorInfo = xAnchors.Last()
                                Dim placement As Point2d =
                                    tg.CreatePoint2d((aFirst.Coord + aLast.Coord) / 2.0, totalY)

                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    placement,
                                    MakeIntent(oSheet, aFirst),
                                    MakeIntent(oSheet, aLast),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                nTotal += 1
                            Catch
                                nFail += 1
                            End Try
                        End If
                    End If

                    '===== 5. CHAIN DIM DỌC =====
                    If yAnchors.Count >= 2 Then
                        For i As Integer = 0 To yAnchors.Count - 2
                            Dim a1 As AnchorInfo = yAnchors(i)
                            Dim a2 As AnchorInfo = yAnchors(i + 1)
                            If Math.Abs(a2.Coord - a1.Coord) <= TOL Then Continue For

                            Try
                                Dim placement As Point2d =
                                    tg.CreatePoint2d(chainX, (a1.Coord + a2.Coord) / 2.0)

                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    placement,
                                    MakeIntent(oSheet, a1),
                                    MakeIntent(oSheet, a2),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                nChain += 1
                            Catch
                                nFail += 1
                            End Try
                        Next

                        If addTotal Then
                            Try
                                Dim aFirst As AnchorInfo = yAnchors.First()
                                Dim aLast As AnchorInfo = yAnchors.Last()
                                Dim placement As Point2d =
                                    tg.CreatePoint2d(totalX, (aFirst.Coord + aLast.Coord) / 2.0)

                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    placement,
                                    MakeIntent(oSheet, aFirst),
                                    MakeIntent(oSheet, aLast),
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
                    "Chain Line cạnh + lỗ",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "Chain Line cạnh + lỗ",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        '=============================================================
        Private Function MakeIntent(ByVal oSheet As Sheet,
                                    ByVal a As AnchorInfo) As GeometryIntent
            If a Is Nothing OrElse a.Curve Is Nothing Then Return Nothing

            If a.Type = AnchorType.Hole Then
                Return oSheet.CreateGeometryIntent(a.Curve, PointIntentEnum.kCenterPointIntent)
            Else
                Return oSheet.CreateGeometryIntent(a.Curve)
            End If
        End Function

        '=============================================================
        Private Function HasNearAnchor(ByVal list As List(Of AnchorInfo),
                                       ByVal value As Double) As Boolean
            For Each a As AnchorInfo In list
                If Math.Abs(a.Coord - value) <= TOL Then Return True
            Next
            Return False
        End Function

        '=============================================================
        ' LỌC DIM THEO VIEW ĐÃ CHỌN (Inventor 2020)
        '=============================================================
        Private Function IsDimInAnyView(ByVal oDim As DrawingDimension,
                                         ByVal views As List(Of DrawingView)) As Boolean

            ' 1. IntentOne / IntentTwo
            Try
                Dim linDim As LinearGeneralDimension = TryCast(oDim, LinearGeneralDimension)
                If linDim IsNot Nothing Then
                    If CheckIntentBelongsToView(linDim.IntentOne, views) Then Return True
                    If CheckIntentBelongsToView(linDim.IntentTwo, views) Then Return True
                End If
            Catch
            End Try

            ' 2. AttachedEntity
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

            ' 3. Bounding box text
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
                                       ByRef addTotal As Boolean,
                                       ByRef includeHoles As Boolean) As Boolean

            Dim frm As New Form()
            frm.Text = "Chain Line — Hướng chuẩn"
            frm.ClientSize = New Size(420, 380)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False

            '----------- Group NGANG (X) -----------
            Dim gbH As New GroupBox()
            gbH.Text = "Hướng chuẩn NGANG (X)"
            gbH.SetBounds(12, 12, 396, 120)
            frm.Controls.Add(gbH)

            Dim rbHUp As New RadioButton()
            rbHUp.Text = "Đặt chain PHÍA TRÊN chi tiết"
            rbHUp.SetBounds(15, 22, 370, 20)
            rbHUp.Checked = True
            gbH.Controls.Add(rbHUp)

            Dim rbHDown As New RadioButton()
            rbHDown.Text = "Đặt chain PHÍA DƯỚI chi tiết"
            rbHDown.SetBounds(15, 46, 370, 20)
            gbH.Controls.Add(rbHDown)

            Dim lblH As New Label()
            lblH.SetBounds(15, 74, 370, 34)
            lblH.Text = "Chain đo liên tiếp giữa các cạnh đứng và tâm lỗ." & vbCrLf &
                        "Dim tổng (nếu bật) đặt xa hơn ở trên/dưới."
            lblH.ForeColor = System.Drawing.Color.Gray
            lblH.Font = New Font("Segoe UI", 8, FontStyle.Italic)
            gbH.Controls.Add(lblH)

            '----------- Group DỌC (Y) -----------
            Dim gbV As New GroupBox()
            gbV.Text = "Hướng chuẩn DỌC (Y)"
            gbV.SetBounds(12, 140, 396, 120)
            frm.Controls.Add(gbV)

            Dim rbVLeft As New RadioButton()
            rbVLeft.Text = "Đặt chain BÊN TRÁI chi tiết"
            rbVLeft.SetBounds(15, 22, 370, 20)
            rbVLeft.Checked = True
            gbV.Controls.Add(rbVLeft)

            Dim rbVRight As New RadioButton()
            rbVRight.Text = "Đặt chain BÊN PHẢI chi tiết"
            rbVRight.SetBounds(15, 46, 370, 20)
            gbV.Controls.Add(rbVRight)

            Dim lblV As New Label()
            lblV.SetBounds(15, 74, 370, 34)
            lblV.Text = "Chain đo liên tiếp giữa các cạnh ngang và tâm lỗ." & vbCrLf &
                        "Dim tổng (nếu bật) đặt xa hơn ở trái/phải."
            lblV.ForeColor = System.Drawing.Color.Gray
            lblV.Font = New Font("Segoe UI", 8, FontStyle.Italic)
            gbV.Controls.Add(lblV)

            '----------- Tùy chọn -----------
            Dim lblOpt As New Label()
            lblOpt.Text = "Tùy chọn:"
            lblOpt.SetBounds(15, 266, 100, 20)
            lblOpt.Font = New Font("Segoe UI", 8, FontStyle.Bold)
            frm.Controls.Add(lblOpt)

            Dim chkTotal As New CheckBox()
            chkTotal.Text = "Thêm dim TỔNG bao ngoài (ví dụ: 584,00)"
            chkTotal.SetBounds(15, 288, 396, 22)
            chkTotal.Checked = True
            frm.Controls.Add(chkTotal)

            Dim chkHoles As New CheckBox()
            chkHoles.Text = "Bao gồm LỖ TRÒN (đo theo tâm lỗ)"
            chkHoles.SetBounds(15, 312, 396, 22)
            chkHoles.Checked = True
            frm.Controls.Add(chkHoles)

            '----------- Buttons -----------
            Dim btnOK As New Button()
            btnOK.Text = "OK"
            btnOK.SetBounds(220, 342, 85, 28)
            btnOK.DialogResult = DialogResult.OK
            frm.Controls.Add(btnOK)

            Dim btnCancel As New Button()
            btnCancel.Text = "Hủy"
            btnCancel.SetBounds(315, 342, 85, 28)
            btnCancel.DialogResult = DialogResult.Cancel
            frm.Controls.Add(btnCancel)

            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            If frm.ShowDialog() <> DialogResult.OK Then Return False

            chainTop = rbHUp.Checked
            chainLeft = rbVLeft.Checked
            addTotal = chkTotal.Checked
            includeHoles = chkHoles.Checked
            Return True
        End Function

        '=============================================================
        Public Enum AnchorType
            Edge
            Hole
        End Enum

        Public Class AnchorInfo
            Public Type As AnchorType
            Public Coord As Double
            Public Curve As DrawingCurve
            Public Center As Point2d
        End Class

    End Module
End Namespace