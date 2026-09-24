Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.AutoCreateDrawing

    Public Module ASS_Auto_Drawing

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowMainMenu()
                Case 1 : AutoDrawingASSTopLV.OnExecute(Context)
                Case 2 : AutoDrawingASSpartTopLV.OnExecute(Context)
                Case 3 : AutoDrawingV8.OnExecute(Context)
            End Select
        End Sub

        Private Function ShowMainMenu() As Integer
            Dim result As Integer = 0

            Using frm As New Form()
                frm.Text = "Auto Drawing Tool"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                frm.ClientSize = New Size(520, 420)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
                frm.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

                '===== HEADER =====
                Dim pnlHeader As New Panel()
                pnlHeader.Location = New System.Drawing.Point(0, 0)
                pnlHeader.Size = New Size(520, 70)
                pnlHeader.BackColor = System.Drawing.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "AUTO DRAWING"
                lblTitle.Font = New Font("Segoe UI", 15.0F, FontStyle.Bold, GraphicsUnit.Point)
                lblTitle.ForeColor = System.Drawing.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Chọn kiểu tạo bản vẽ"
                lblSub.Font = New Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point)
                lblSub.ForeColor = System.Drawing.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== MENU ITEMS =====
                AddMenuItem(frm, 1, "Drawing Top ASS",
                            "Tạo bản vẽ cụm tổng + các cụm lắp con", 90)

                AddMenuItem(frm, 2, "Drawing Top ASS + Part",
                            "Tạo bản vẽ cụm + chi tiết + Sub-Assembly cấp cao", 175)

                AddMenuItem(frm, 3, "Auto Drawing V8",
                            "Tạo bản vẽ theo cấu hình nâng cao (nhiều tùy chọn)", 260)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Size(120, 36)
                btnCancel.Location = New System.Drawing.Point(380, 360)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(180, 180, 180)
                btnCancel.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
                btnCancel.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
                AddHandler btnCancel.Click, Sub()
                                                result = 0
                                                frm.Close()
                                            End Sub
                frm.Controls.Add(btnCancel)
                frm.CancelButton = btnCancel

                frm.ShowDialog()
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
            pnl.Size = New System.Drawing.Size(480, 75)
            pnl.BackColor = System.Drawing.Color.White
            pnl.BorderStyle = BorderStyle.FixedSingle
            pnl.Cursor = Cursors.Hand

            '--- Số thứ tự ---
            Dim lblNum As New Label()
            lblNum.Text = value.ToString()
            lblNum.Font = New Font("Segoe UI", 20.0F, FontStyle.Bold, GraphicsUnit.Point)
            lblNum.ForeColor = System.Drawing.Color.FromArgb(45, 100, 180)
            lblNum.Location = New System.Drawing.Point(10, 15)
            lblNum.Size = New System.Drawing.Size(55, 45)
            lblNum.TextAlign = ContentAlignment.MiddleCenter
            pnl.Controls.Add(lblNum)

            '--- Đường phân cách ---
            Dim sep As New Panel()
            sep.Location = New System.Drawing.Point(75, 15)
            sep.Size = New System.Drawing.Size(1, 45)
            sep.BackColor = System.Drawing.Color.FromArgb(220, 220, 220)
            pnl.Controls.Add(sep)

            '--- Tiêu đề ---
            Dim lblTitle As New Label()
            lblTitle.Text = title
            lblTitle.Font = New Font("Segoe UI", 11.5F, FontStyle.Bold, GraphicsUnit.Point)
            lblTitle.ForeColor = System.Drawing.Color.FromArgb(30, 30, 30)
            lblTitle.Location = New System.Drawing.Point(90, 14)
            lblTitle.AutoSize = True
            pnl.Controls.Add(lblTitle)

            '--- Mô tả ---
            Dim lblDesc As New Label()
            lblDesc.Text = desc
            lblDesc.Font = New Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point)
            lblDesc.ForeColor = System.Drawing.Color.FromArgb(110, 110, 110)
            lblDesc.Location = New System.Drawing.Point(90, 42)
            lblDesc.AutoSize = True
            pnl.Controls.Add(lblDesc)

            '--- Hover effect ---
            Dim hoverOn As EventHandler = Sub()
                                              pnl.BackColor = System.Drawing.Color.FromArgb(235, 242, 252)
                                              pnl.BorderStyle = BorderStyle.FixedSingle
                                          End Sub
            Dim hoverOff As EventHandler = Sub()
                                               pnl.BackColor = System.Drawing.Color.White
                                           End Sub

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