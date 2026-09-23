Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq
Namespace ToolInventor2025.Drawing.Buttons.Drawdim
    Public Module Draw_dim_chain_line_c_old
        Private Const TOL As Double = 0.03
        '=========================================================
        ' LƯU THÔNG SỐ GIỮA CÁC LẦN CHẠY (đơn vị cm)
        '=========================================================
        Private lastChainGap As Double = 0.4        ' 4 mm
        Private lastTotalGap As Double = 1.2        ' 12 mm
        Private lastDeleteBelow As Double = 0.5     ' 5 mm — xóa dim < ngưỡng (0 = không xóa)
        Private lastChainTop As Boolean = True
        Private lastChainLeft As Boolean = True
        Private lastAddTotal As Boolean = True
        Private lastIncludeHoles As Boolean = True
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
                Dim chainTop As Boolean = lastChainTop
                Dim chainLeft As Boolean = lastChainLeft
                Dim addTotal As Boolean = lastAddTotal
                Dim includeHoles As Boolean = lastIncludeHoles
                Dim chainGap As Double = lastChainGap
                Dim totalGap As Double = lastTotalGap
                Dim deleteBelow As Double = lastDeleteBelow
                If Not ShowChainForm(chainTop, chainLeft, addTotal, includeHoles,
                                     chainGap, totalGap, deleteBelow) Then Exit Sub
                ' Lưu cho lần chạy sau
                lastChainTop = chainTop
                lastChainLeft = chainLeft
                lastAddTotal = addTotal
                lastIncludeHoles = includeHoles
                lastChainGap = chainGap
                lastTotalGap = totalGap
                lastDeleteBelow = deleteBelow
                Dim nChain As Integer = 0
                Dim nTotal As Integer = 0
                Dim nFail As Integer = 0
                '=========================================================
                ' TẠO CHAIN DIM
                '=========================================================
                For Each oView As DrawingView In selectedViews
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
                    xAnchors = xAnchors.OrderBy(Function(a) a.Coord).ToList()
                    yAnchors = yAnchors.OrderBy(Function(a) a.Coord).ToList()
                    Dim chainY As Double = If(chainTop, maxY + chainGap, minY - chainGap)
                    Dim totalY As Double = If(chainTop, maxY + totalGap, minY - totalGap)
                    Dim chainX As Double = If(chainLeft, minX - chainGap, maxX + chainGap)
                    Dim totalX As Double = If(chainLeft, minX - totalGap, maxX + totalGap)
                    '===== CHAIN DIM NGANG =====
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
                    '===== CHAIN DIM DỌC =====
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
                '=========================================================
                ' XÓA DIM NHỎ HƠN NGƯỠNG
                '=========================================================
                Dim nDeleted As Integer = 0
                Dim nDelFail As Integer = 0
                If deleteBelow > 0 Then
                    ' Đổi từ cm → dùng trực tiếp giá trị đã lưu (cm)
                    Dim thresholdCM As Double = deleteBelow
                    Dim toDelete As New List(Of DrawingDimension)
                    For Each oDim As DrawingDimension In oSheet.DrawingDimensions
                        Try
                            ' Chỉ xử lý dim thẳng
                            Dim linDim As LinearGeneralDimension = TryCast(oDim, LinearGeneralDimension)
                            If linDim Is Nothing Then Continue For
                            Dim valCM As Double = 0
                            Try
                                valCM = linDim.ModelValue
                            Catch
                                Try
                                    valCM = linDim.Value
                                Catch
                                    Continue For
                                End Try
                            End Try
                            If Math.Abs(valCM) < thresholdCM Then
                                toDelete.Add(oDim)
                            End If
                        Catch
                        End Try
                    Next
                    For Each d As DrawingDimension In toDelete
                        Try
                            d.Delete()
                            nDeleted += 1
                        Catch
                            nDelFail += 1
                        End Try
                    Next
                End If
                '=========================================================
                ' ARRANGE
                '=========================================================
                Try
                    Dim oDims As DrawingDimensions = oSheet.DrawingDimensions
                    Dim col As ObjectCollection = app.TransientObjects.CreateObjectCollection
                    For Each d As DrawingDimension In oDims
                        Try
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
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Chain dim: " & nChain & vbCrLf &
                    "Dim tổng: " & nTotal & vbCrLf &
                    "Đã xóa dim nhỏ: " & nDeleted & vbCrLf &
                    "Lỗi dim: " & nFail & vbCrLf &
                    "Lỗi xóa: " & nDelFail & vbCrLf & vbCrLf &
                    "--- Thông số ---" & vbCrLf &
                    "Chain gap: " & (chainGap * 10.0).ToString("0.0#") & " mm" & vbCrLf &
                    "Total gap: " & (totalGap * 10.0).ToString("0.0#") & " mm" & vbCrLf &
                    "Xóa dim < " & If(deleteBelow > 0,
                                       (deleteBelow * 10.0).ToString("0.0#") & " mm",
                                       "(không xóa)"),
                    "Chain Line",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "Chain Line",
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
        ' FORM
        '=============================================================
        Private Function ShowChainForm(ByRef chainTop As Boolean,
                                       ByRef chainLeft As Boolean,
                                       ByRef addTotal As Boolean,
                                       ByRef includeHoles As Boolean,
                                       ByRef chainGap As Double,
                                       ByRef totalGap As Double,
                                       ByRef deleteBelow As Double) As Boolean
            Dim frm As New Form()
            frm.Text = "Chain Line — Hướng chuẩn"
            frm.ClientSize = New Size(460, 550)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.ShowInTaskbar = False
            '----------- Group NGANG (X) -----------
            Dim gbH As New GroupBox()
            gbH.Text = "Hướng chuẩn NGANG (X)"
            gbH.SetBounds(12, 12, 416, 110)
            frm.Controls.Add(gbH)
            Dim rbHUp As New RadioButton()
            rbHUp.Text = "Đặt chain PHÍA TRÊN chi tiết"
            rbHUp.SetBounds(15, 22, 390, 20)
            rbHUp.Checked = chainTop
            gbH.Controls.Add(rbHUp)
            Dim rbHDown As New RadioButton()
            rbHDown.Text = "Đặt chain PHÍA DƯỚI chi tiết"
            rbHDown.SetBounds(15, 46, 390, 20)
            rbHDown.Checked = Not chainTop
            gbH.Controls.Add(rbHDown)
            Dim lblH As New Label()
            lblH.SetBounds(15, 74, 390, 30)
            lblH.Text = "Chain đo liên tiếp giữa các cạnh đứng và tâm lỗ."
            lblH.ForeColor = System.Drawing.Color.Gray
            lblH.Font = New Font("Segoe UI", 8, FontStyle.Italic)
            gbH.Controls.Add(lblH)
            '----------- Group DỌC (Y) -----------
            Dim gbV As New GroupBox()
            gbV.Text = "Hướng chuẩn DỌC (Y)"
            gbV.SetBounds(12, 130, 416, 110)
            frm.Controls.Add(gbV)
            Dim rbVLeft As New RadioButton()
            rbVLeft.Text = "Đặt chain BÊN TRÁI chi tiết"
            rbVLeft.SetBounds(15, 22, 390, 20)
            rbVLeft.Checked = chainLeft
            gbV.Controls.Add(rbVLeft)
            Dim rbVRight As New RadioButton()
            rbVRight.Text = "Đặt chain BÊN PHẢI chi tiết"
            rbVRight.SetBounds(15, 46, 390, 20)
            rbVRight.Checked = Not chainLeft
            gbV.Controls.Add(rbVRight)
            Dim lblV As New Label()
            lblV.SetBounds(15, 74, 390, 30)
            lblV.Text = "Chain đo liên tiếp giữa các cạnh ngang và tâm lỗ."
            lblV.ForeColor = System.Drawing.Color.Gray
            lblV.Font = New Font("Segoe UI", 8, FontStyle.Italic)
            gbV.Controls.Add(lblV)
            '----------- Group KHOẢNG CÁCH -----------
            Dim gbO As New GroupBox()
            gbO.Text = "Khoảng cách dim line (mm)"
            gbO.SetBounds(12, 250, 416, 100)
            frm.Controls.Add(gbO)
            Dim cbChainGap As New ComboBox()
            cbChainGap.DropDownStyle = ComboBoxStyle.DropDown
            cbChainGap.SetBounds(120, 25, 100, 22)
            cbChainGap.Items.AddRange({"1.0", "2.0", "3.2", "5.0", "8.0", "10.0"})
            cbChainGap.Text = (chainGap * 10.0).ToString("0.0#")
            gbO.Controls.Add(cbChainGap)
            Dim cbTotalGap As New ComboBox()
            cbTotalGap.DropDownStyle = ComboBoxStyle.DropDown
            cbTotalGap.SetBounds(120, 58, 100, 22)
            cbTotalGap.Items.AddRange({"8.0", "12.0", "15.0", "20.0", "25.0", "30.0"})
            cbTotalGap.Text = (totalGap * 10.0).ToString("0.0#")
            gbO.Controls.Add(cbTotalGap)
            gbO.Controls.Add(New Label With {
                .Text = "Chain gap:",
                .Bounds = New Rectangle(15, 28, 100, 22)
            })
            gbO.Controls.Add(New Label With {
                .Text = "Total gap:",
                .Bounds = New Rectangle(15, 61, 100, 22)
            })
            '----------- Group XÓA DIM NHỎ -----------
            Dim gbDel As New GroupBox()
            gbDel.Text = "Xóa dim nhỏ (mm) — để trống hoặc 0 = không xóa"
            gbDel.SetBounds(12, 358, 416, 85)
            frm.Controls.Add(gbDel)
            Dim cbDelBelow As New ComboBox()
            cbDelBelow.DropDownStyle = ComboBoxStyle.DropDown
            cbDelBelow.SetBounds(180, 30, 100, 22)
            cbDelBelow.Items.AddRange({"", "0", "1", "2", "3", "5", "8", "10"})
            If deleteBelow > 0 Then
                cbDelBelow.Text = (deleteBelow * 10.0).ToString("0.0#")
            Else
                cbDelBelow.Text = ""
            End If
            gbDel.Controls.Add(cbDelBelow)
            gbDel.Controls.Add(New Label With {
                .Text = "Xóa dim < :",
                .Bounds = New Rectangle(15, 33, 160, 22)
            })
            Dim lblDelHint As New Label()
            lblDelHint.SetBounds(290, 22, 115, 50)
            lblDelHint.Text = "VD: 5 → xóa" & vbCrLf &
                              "mọi dim < 5mm" & vbCrLf &
                              "sau khi tạo chain."
            lblDelHint.ForeColor = System.Drawing.Color.Gray
            lblDelHint.Font = New Font("Segoe UI", 7.5F, FontStyle.Italic)
            gbDel.Controls.Add(lblDelHint)
            '----------- Tùy chọn -----------
            Dim chkTotal As New CheckBox()
            chkTotal.Text = "Thêm dim TỔNG bao ngoài"
            chkTotal.SetBounds(15, 455, 396, 22)
            chkTotal.Checked = addTotal
            frm.Controls.Add(chkTotal)
            Dim chkHoles As New CheckBox()
            chkHoles.Text = "Bao gồm LỖ TRÒN (đo theo tâm lỗ)"
            chkHoles.SetBounds(15, 479, 396, 22)
            chkHoles.Checked = includeHoles
            frm.Controls.Add(chkHoles)
            '----------- Buttons -----------
            Dim btnOK As New Button()
            btnOK.Text = "OK"
            btnOK.SetBounds(230, 510, 85, 28)
            frm.Controls.Add(btnOK)
            Dim btnCancel As New Button()
            btnCancel.Text = "Hủy"
            btnCancel.SetBounds(325, 510, 85, 28)
            frm.Controls.Add(btnCancel)
            '----------- Biến tạm -----------
            Dim tempChainGap As Double = chainGap
            Dim tempTotalGap As Double = totalGap
            Dim tempDeleteBelow As Double = deleteBelow
            Dim tempChainTop As Boolean = chainTop
            Dim tempChainLeft As Boolean = chainLeft
            Dim tempAddTotal As Boolean = addTotal
            Dim tempIncludeHoles As Boolean = includeHoles
            Dim okClicked As Boolean = False
            '----------- OK handler -----------
            AddHandler btnOK.Click,
                Sub()
                    Dim tmp As Double
                    frm.ActiveControl = btnOK
                    System.Windows.Forms.Application.DoEvents()
                    ' Chain gap
                    If Not Double.TryParse(cbChainGap.Text.Replace(",", ".").Trim(),
                                           Globalization.NumberStyles.Any,
                                           Globalization.CultureInfo.InvariantCulture,
                                           tmp) OrElse tmp <= 0 Then
                        tmp = 4.0
                    End If
                    tempChainGap = tmp / 10.0
                    ' Total gap
                    If Not Double.TryParse(cbTotalGap.Text.Replace(",", ".").Trim(),
                                           Globalization.NumberStyles.Any,
                                           Globalization.CultureInfo.InvariantCulture,
                                           tmp) OrElse tmp <= 0 Then
                        tmp = 12.0
                    End If
                    tempTotalGap = tmp / 10.0
                    ' Delete below — để trống = 0 (không xóa)
                    If String.IsNullOrWhiteSpace(cbDelBelow.Text) Then
                        tempDeleteBelow = 0.0
                    ElseIf Not Double.TryParse(cbDelBelow.Text.Replace(",", ".").Trim(),
                                               Globalization.NumberStyles.Any,
                                               Globalization.CultureInfo.InvariantCulture,
                                               tmp) OrElse tmp <= 0 Then
                        tempDeleteBelow = 0.0
                    Else
                        tempDeleteBelow = tmp / 10.0
                    End If
                    tempChainTop = rbHUp.Checked
                    tempChainLeft = rbVLeft.Checked
                    tempAddTotal = chkTotal.Checked
                    tempIncludeHoles = chkHoles.Checked
                    okClicked = True
                    frm.Close()
                End Sub
            AddHandler btnCancel.Click,
                Sub()
                    okClicked = False
                    frm.Close()
                End Sub
            '----------- MODELESS -----------
            frm.Show()
            Do While frm.Visible
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(15)
            Loop
            If Not okClicked Then Return False
            chainGap = tempChainGap
            totalGap = tempTotalGap
            deleteBelow = tempDeleteBelow
            chainTop = tempChainTop
            chainLeft = tempChainLeft
            addTotal = tempAddTotal
            includeHoles = tempIncludeHoles
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