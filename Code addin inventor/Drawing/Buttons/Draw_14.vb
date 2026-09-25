Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons

    Public Module draw_14

        '==========================================================
        ' OPTIONS
        '==========================================================
        Private Class SectionOptions
            Public NewScale As Double = 0.5
            Public ScaleText As String = "0.5"
            Public Scope As Integer = 1              ' 1=Active Sheet, 2=All Sheets
            Public ForceHatch As Boolean = True
            Public ForceRebuild As Boolean = True
            Public Cancelled As Boolean = True
        End Class


        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                Dim oDrawDoc As DrawingDocument =
                    TryCast(g_inventorApplication.ActiveDocument, DrawingDocument)

                If oDrawDoc Is Nothing Then
                    MessageBox.Show("Document hiện tại không phải Drawing.",
                                    "Change Section Scale + Hatch",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                '=====================================================
                ' FORM TÙY CHỌN
                '=====================================================
                Dim opt As SectionOptions = ShowOptionsForm()
                If opt Is Nothing OrElse opt.Cancelled Then Return

                Dim newScale As Double = opt.NewScale

                If newScale <= 0 Then
                    MessageBox.Show("Scale không hợp lệ! Ví dụ: 0.5 hoặc 1/2",
                                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim count As Integer = 0
                Dim failCount As Integer = 0

                '=====================================================
                ' XỬ LÝ THEO PHẠM VI
                '=====================================================
                If opt.Scope = 1 Then
                    count = ProcessSheet(oDrawDoc.ActiveSheet, newScale,
                                         opt.ForceHatch, failCount)
                Else
                    For Each oSheet As Sheet In oDrawDoc.Sheets
                        count += ProcessSheet(oSheet, newScale,
                                              opt.ForceHatch, failCount)
                    Next
                End If

                '=====================================================
                ' FORCE UPDATE DRAWING (Inventor 2025)
                '=====================================================
                Try : oDrawDoc.Update2(True) : Catch : End Try

                If opt.ForceRebuild Then
                    ' Inventor 2025: Rebuild2() KHÔNG có tham số
                    Try : oDrawDoc.Rebuild2() : Catch : End Try
                    Try : oDrawDoc.Update2(True) : Catch : End Try
                End If

                '=====================================================
                ' KẾT QUẢ
                '=====================================================
                Dim msg As String =
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Phạm vi        : " & If(opt.Scope = 1, "Sheet hiện tại", "Toàn bộ Drawing") & vbCrLf &
                    "Mặt cắt đã xử lý: " & count.ToString() & vbCrLf &
                    "Scale mới      : " & opt.ScaleText & vbCrLf &
                    "Force Hatch    : " & If(opt.ForceHatch, "Có", "Không") & vbCrLf &
                    "Force Rebuild  : " & If(opt.ForceRebuild, "Có", "Không")

                If failCount > 0 Then
                    msg &= vbCrLf & vbCrLf &
                           "⚠ Có " & failCount.ToString() & " mặt cắt bỏ qua."
                End If

                msg &= vbCrLf & vbCrLf &
                       "Nếu Hatch vẫn sai: Save → đóng → mở lại file."

                MessageBox.Show(msg, "Change Section Scale + Hatch",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message,
                                "Change Section Scale + Hatch",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '==========================================================
        ' XỬ LÝ 1 SHEET
        '==========================================================
        Private Function ProcessSheet(ByVal oSheet As Sheet,
                                      ByVal newScale As Double,
                                      ByVal forceHatch As Boolean,
                                      ByRef failCount As Integer) As Integer

            If oSheet Is Nothing Then Return 0

            Dim count As Integer = 0

            For Each oView As DrawingView In oSheet.DrawingViews
                Try
                    If ForceSectionScaleAndHatch(oView, newScale, forceHatch) Then
                        count += 1
                    Else
                        failCount += 1
                    End If
                Catch
                    failCount += 1
                End Try
            Next

            Return count

        End Function


        '==========================================================
        ' FORCE SCALE + HATCH (Inventor 2025)
        '==========================================================
        Private Function ForceSectionScaleAndHatch(ByVal oView As DrawingView,
                                                   ByVal newScale As Double,
                                                   ByVal forceHatch As Boolean) As Boolean

            Try
                ' Chỉ xử lý Section View
                If oView.ViewType <> DrawingViewTypeEnum.kSectionDrawingViewType Then
                    Return False
                End If

                '==================================================
                ' 1. ÉP SCALE
                '==================================================
                Dim oldScale As Double = oView.Scale

                Try
                    oView.Scale = newScale
                Catch
                    Return False
                End Try

                If Not forceHatch Then Return True

                '==================================================
                ' 2. FORCE HATCH REFRESH — Inventor 2025
                '==================================================

                '--- Cách 1: Đổi SectionDepth cực nhỏ rồi trả lại ---
                Try
                    Dim secView As SectionDrawingView = TryCast(oView, SectionDrawingView)
                    If secView IsNot Nothing Then
                        Try
                            Dim oldDepth As Double = secView.Depth
                            If oldDepth > 0 Then
                                secView.Depth = oldDepth * 1.0001
                                secView.Depth = oldDepth
                            End If
                        Catch
                        End Try
                    End If
                Catch
                End Try

                '--- Cách 2: Tắt/Bật Hidden Lines ---
                Try
                    Dim oldHidden As Boolean = oView.ShowHiddenLines
                    oView.ShowHiddenLines = Not oldHidden
                    oView.ShowHiddenLines = oldHidden
                Catch
                End Try

                '--- Cách 3: Tắt/Bật Tangent Edges ---
                Try
                    Dim oldTangent As Boolean = oView.DisplayTangentEdges
                    oView.DisplayTangentEdges = Not oldTangent
                    oView.DisplayTangentEdges = oldTangent
                Catch
                End Try

                '--- Cách 4: Dịch vị trí cực nhỏ ---
                Try
                    Dim tg As TransientGeometry = g_inventorApplication.TransientGeometry
                    Dim oldPos As Point2d = oView.Position
                    oView.Position = tg.CreatePoint2d(oldPos.X + 0.002, oldPos.Y + 0.002)
                    oView.Position = oldPos
                Catch
                End Try

                '--- Cách 5: Refresh qua DrawingStyle ---
                Try
                    Dim oldStyle As DrawingViewStyleEnum = oView.ViewStyle
                    oView.ViewStyle = DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle
                    oView.ViewStyle = oldStyle
                Catch
                End Try

                '--- Cách 6: Ép Scale lần nữa (sau khi dirty) ---
                Try
                    oView.Scale = newScale
                Catch
                End Try

                '--- Cách 7: Force update document cha ---
                Try
                    Dim parent As Object = oView.Parent
                    If parent IsNot Nothing Then
                        Dim sh As Sheet = TryCast(parent, Sheet)
                        If sh IsNot Nothing Then
                            Dim doc As DrawingDocument = TryCast(sh.Parent, DrawingDocument)
                            If doc IsNot Nothing Then
                                doc.Update2(True)
                            End If
                        End If
                    End If
                Catch
                End Try

                Return True

            Catch
                Return False
            End Try

        End Function


        '==========================================================
        ' PARSE SCALE
        '==========================================================
        Private Function ParseScale(ByVal text As String) As Double

            If String.IsNullOrEmpty(text) Then Return 0

            text = text.Trim().Replace(","c, "."c)

            If text.Contains("/") Then
                Dim parts() As String = text.Split("/"c)
                If parts.Length = 2 Then
                    Dim num, den As Double
                    If Double.TryParse(parts(0), num) AndAlso
                       Double.TryParse(parts(1), den) AndAlso den <> 0 Then
                        Return num / den
                    End If
                End If
            End If

            Dim result As Double
            If Double.TryParse(text, result) Then
                Return result
            End If

            Return 0

        End Function


        '==========================================================
        ' FORM TÙY CHỌN
        '==========================================================
        Private Function ShowOptionsForm() As SectionOptions
            Dim opt As New SectionOptions()
            Dim _okConfirmed As Boolean = False

            Using frm As New Form()
                frm.Text = "Đổi Scale + Hatch Mặt cắt"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(660, 520)
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
                pnlHeader.Size = New Drw.Size(660, 75)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "ĐỔI SCALE + HATCH MẶT CẮT"
                lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Cập nhật Scale + Force refresh Hatch của Section View"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== GROUP 1: SCALE =====
                Dim gb1 As New GroupBox() With {
                    .Text = "1. Scale mới",
                    .Location = New Drw.Point(15, 90),
                    .Size = New Drw.Size(630, 100),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb1)

                Dim lblScale As New Label() With {
                    .Text = "Scale mới:",
                    .Location = New Drw.Point(25, 38),
                    .Size = New Drw.Size(120, 25),
                    .TextAlign = Drw.ContentAlignment.MiddleLeft,
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(lblScale)

                Dim txtScale As New System.Windows.Forms.TextBox() With {
                    .Text = "0.5",
                    .Location = New Drw.Point(160, 38),
                    .Size = New Drw.Size(150, 25),
                    .Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(txtScale)

                Dim lblHint As New Label() With {
                    .Text = "Ví dụ: 0.5   hoặc   1/2   (đều = 50%)",
                    .Location = New Drw.Point(160, 68),
                    .Size = New Drw.Size(440, 20),
                    .ForeColor = Drw.Color.FromArgb(140, 140, 140),
                    .Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(lblHint)

                '===== GROUP 2: PHẠM VI =====
                Dim gb2 As New GroupBox() With {
                    .Text = "2. Phạm vi áp dụng",
                    .Location = New Drw.Point(15, 200),
                    .Size = New Drw.Size(630, 110),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb2)

                Dim rdoActive As New RadioButton() With {
                    .Text = "Chỉ Sheet đang mở",
                    .Location = New Drw.Point(25, 32),
                    .Size = New Drw.Size(580, 25),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb2.Controls.Add(rdoActive)

                Dim rdoAll As New RadioButton() With {
                    .Text = "Toàn bộ Drawing (tất cả Sheet)",
                    .Location = New Drw.Point(25, 65),
                    .Size = New Drw.Size(580, 25),
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb2.Controls.Add(rdoAll)

                '===== GROUP 3: TÙY CHỌN NÂNG CAO =====
                Dim gb3 As New GroupBox() With {
                    .Text = "3. Tùy chọn nâng cao",
                    .Location = New Drw.Point(15, 320),
                    .Size = New Drw.Size(630, 100),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb3)

                Dim chkHatch As New CheckBox() With {
                    .Text = "Force refresh Hatch (khuyên dùng khi Hatch bị sai)",
                    .Location = New Drw.Point(25, 30),
                    .Size = New Drw.Size(580, 25),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb3.Controls.Add(chkHatch)

                Dim chkRebuild As New CheckBox() With {
                    .Text = "Force Rebuild toàn bộ Drawing (chậm hơn, chắc hơn)",
                    .Location = New Drw.Point(25, 62),
                    .Size = New Drw.Size(580, 25),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb3.Controls.Add(chkRebuild)

                '===== NÚT THỰC HIỆN =====
                Dim btnOK As New Button()
                btnOK.Text = "THỰC HIỆN"
                btnOK.Size = New Drw.Size(170, 46)
                btnOK.Location = New Drw.Point(475, 445)
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
                                            Dim sc As Double = ParseScale(txtScale.Text)
                                            If sc <= 0 Then
                                                MessageBox.Show("Scale không hợp lệ! Ví dụ: 0.5 hoặc 1/2",
                                                                "Lỗi",
                                                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                txtScale.Focus()
                                                Return
                                            End If

                                            opt.ScaleText = txtScale.Text.Trim()
                                            opt.NewScale = sc
                                            opt.Scope = If(rdoAll.Checked, 2, 1)
                                            opt.ForceHatch = chkHatch.Checked
                                            opt.ForceRebuild = chkRebuild.Checked
                                            opt.Cancelled = False
                                            _okConfirmed = True
                                            frm.DialogResult = DialogResult.OK
                                            frm.Close()
                                        End Sub
                frm.Controls.Add(btnOK)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(130, 46)
                btnCancel.Location = New Drw.Point(335, 445)
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