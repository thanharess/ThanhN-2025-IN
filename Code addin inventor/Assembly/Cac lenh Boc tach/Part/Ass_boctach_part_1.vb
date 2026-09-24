Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.caclenhboctach.part

    Public Module Ass_boctach_part_1

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowMainMenu()
                Case 1 : Ass_boctach_part_1a.OnExecute(Context)
                Case 2 : Ass_boctach_part_1b.OnExecute(Context)
                Case 3 : Ass_boctach_part_1c.OnExecute(Context)
                Case 4 : Ass_boctach_part_1d.OnExecute(Context)
                Case 5 : Ass_boctach_part_1e.OnExecute(Context)
                Case 6 : Ass_boctach_part_1f.OnExecute(Context)
                Case 7 : Ass_boctach_part_1g.OnExecute(Context)
                Case 8 : Ass_boctach_part_1h.OnExecute(Context)
            End Select
        End Sub

        Private Function ShowMainMenu() As Integer
            Dim result As Integer = 0

            Using frm As New Form()
                frm.Text = "Bóc tách Part"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                frm.ClientSize = New Size(560, 700)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
                frm.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
                frm.Tag = 0

                '===== HEADER =====
                Dim pnlHeader As New Panel()
                pnlHeader.Location = New System.Drawing.Point(0, 0)
                pnlHeader.Size = New Size(560, 70)
                pnlHeader.BackColor = System.Drawing.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "BÓC TÁCH PART"
                lblTitle.Font = New Font("Segoe UI", 15.0F, FontStyle.Bold, GraphicsUnit.Point)
                lblTitle.ForeColor = System.Drawing.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Chọn kiểu bóc tách"
                lblSub.Font = New Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point)
                lblSub.ForeColor = System.Drawing.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== MENU ITEMS =====
                AddMenuItem(frm, 1, "Top Lever - Tách số loại SP",
                            "Vật tư mua + tiêu chuẩn. Lọc Part", 85)

                AddMenuItem(frm, 2, "Top Lever - Tách số loại SP",
                            "PL + vật tư mua. Lọc ASS + Part", 150)

                AddMenuItem(frm, 3, "Top Lever - Bóc KL tấm",
                            "Tổng khối lượng tấm. Lọc Part", 215)

                AddMenuItem(frm, 4, "Top Lever - Bóc số lượng PL",
                            "PL + vật tư mua. Lọc Part", 280)

                AddMenuItem(frm, 5, "Top Lever - Bóc số lượng PL",
                            "PL + vật tư mua. Lọc ASS + Part", 345)

                AddMenuItem(frm, 6, "All Lever - Bóc KL tấm",
                            "Tổng khối lượng tấm toàn bộ. Lọc Part", 410)

                AddMenuItem(frm, 7, "All Lever - Lọc các loại tấm",
                            "Phân loại tấm theo độ dày. Lọc Part", 475)

                AddMenuItem(frm, 8, "All Lever - Bóc số lượng PL",
                            "PL + vật tư mua. Lọc Part", 540)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Size(120, 36)
                btnCancel.Location = New System.Drawing.Point(420, 650)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(180, 180, 180)
                btnCancel.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
                btnCancel.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
                AddHandler btnCancel.Click, Sub()
                                                frm.Tag = 0
                                                frm.Close()
                                            End Sub
                frm.Controls.Add(btnCancel)
                frm.CancelButton = btnCancel

                frm.ShowDialog()
                result = CInt(frm.Tag)
            End Using

            Return result
        End Function


        Private Sub AddMenuItem(ByVal frm As Form,
                                ByVal value As Integer,
                                ByVal title As String,
                                ByVal desc As String,
                                ByVal top As Integer)

            Dim pnl As New Panel()
            pnl.Location = New System.Drawing.Point(20, top)
            pnl.Size = New System.Drawing.Size(520, 58)
            pnl.BackColor = System.Drawing.Color.White
            pnl.BorderStyle = BorderStyle.FixedSingle
            pnl.Cursor = Cursors.Hand

            '--- Số thứ tự ---
            Dim lblNum As New Label()
            lblNum.Text = value.ToString()
            lblNum.Font = New Font("Segoe UI", 17.0F, FontStyle.Bold, GraphicsUnit.Point)
            lblNum.ForeColor = System.Drawing.Color.FromArgb(45, 100, 180)
            lblNum.Location = New System.Drawing.Point(8, 10)
            lblNum.Size = New System.Drawing.Size(48, 38)
            lblNum.TextAlign = ContentAlignment.MiddleCenter
            pnl.Controls.Add(lblNum)

            '--- Đường kẻ dọc ---
            Dim sep As New Panel()
            sep.Location = New System.Drawing.Point(62, 10)
            sep.Size = New System.Drawing.Size(1, 38)
            sep.BackColor = System.Drawing.Color.FromArgb(220, 220, 220)
            pnl.Controls.Add(sep)

            '--- Tiêu đề ---
            Dim lblTitle As New Label()
            lblTitle.Text = title
            lblTitle.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold, GraphicsUnit.Point)
            lblTitle.ForeColor = System.Drawing.Color.FromArgb(30, 30, 30)
            lblTitle.Location = New System.Drawing.Point(75, 8)
            lblTitle.AutoSize = True
            pnl.Controls.Add(lblTitle)

            '--- Mô tả ---
            Dim lblDesc As New Label()
            lblDesc.Text = desc
            lblDesc.Font = New Font("Segoe UI", 8.0F, FontStyle.Regular, GraphicsUnit.Point)
            lblDesc.ForeColor = System.Drawing.Color.FromArgb(110, 110, 110)
            lblDesc.Location = New System.Drawing.Point(75, 31)
            lblDesc.AutoSize = True
            pnl.Controls.Add(lblDesc)

            '--- Hover ---
            Dim hoverOn As EventHandler = Sub() pnl.BackColor = System.Drawing.Color.FromArgb(235, 242, 252)
            Dim hoverOff As EventHandler = Sub() pnl.BackColor = System.Drawing.Color.White

            AddHandler pnl.MouseEnter, hoverOn
            AddHandler pnl.MouseLeave, hoverOff
            AddHandler lblTitle.MouseEnter, hoverOn
            AddHandler lblTitle.MouseLeave, hoverOff
            AddHandler lblDesc.MouseEnter, hoverOn
            AddHandler lblDesc.MouseLeave, hoverOff
            AddHandler lblNum.MouseEnter, hoverOn
            AddHandler lblNum.MouseLeave, hoverOff

            '--- Click ---
            Dim clickH As EventHandler = Sub(sender, e)
                                             frm.Tag = value
                                             frm.DialogResult = DialogResult.OK
                                             frm.Close()
                                         End Sub
            AddHandler pnl.Click, clickH
            AddHandler lblTitle.Click, clickH
            AddHandler lblDesc.Click, clickH
            AddHandler lblNum.Click, clickH

            frm.Controls.Add(pnl)
        End Sub

    End Module

End Namespace