Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.Drawtext

    Public Module draw_text_list_1

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowMainMenu()
                Case 1 : TextReplaceModule.OnExecute(Context)
                Case 2 : Doichuhoa.OnExecute(Context)
                Case 3 : ThayChuTrongTextModule.OnExecute(Context)
            End Select
        End Sub


        Private Function ShowMainMenu() As Integer
            Dim result As Integer = 0

            Using frm As New Form()
                frm.Text = "Auto Text — Drawing"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(680, 470)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.BackColor = Drw.Color.FromArgb(245, 245, 245)
                frm.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                frm.Tag = 0

                '===== HEADER =====
                Dim pnlHeader As New Panel()
                pnlHeader.Location = New Drw.Point(0, 0)
                pnlHeader.Size = New Drw.Size(680, 75)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "AUTO TEXT — DRAWING"
                lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Thay thế / xử lý chữ trong Drawing Note"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== NHÓM 1: TEXT REPLACE =====
                Dim gb1 As New GroupBox() With {
                    .Text = "Text Replace",
                    .Location = New Drw.Point(15, 90),
                    .Size = New Drw.Size(650, 115),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb1)

                AddCard(gb1, frm, 1,
                        "1. Text Replace",
                        "Thay thế chuỗi ký tự trong Drawing Note",
                        20, 15, 300)

                AddCard(gb1, frm, 2,
                        "2. Đổi chữ IN HOA / thường",
                        "5 kiểu: HOA, thường, Hoa đầu dòng/từ/câu",
                        20, 335, 300)

                '===== NHÓM 2: THAY CHỮ TRONG TEXT =====
                Dim gb2 As New GroupBox() With {
                    .Text = "Thay chữ trong Text (nâng cao)",
                    .Location = New Drw.Point(15, 215),
                    .Size = New Drw.Size(650, 115),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb2)

                AddCard(gb2, frm, 3,
                        "3. Thay chữ trong Text",
                        "Tìm & thay nhiều cặp — Note, Dimension, Weld Symbol...",
                        20, 15, 620)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(130, 44)
                btnCancel.Location = New Drw.Point(535, 410)
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


        '=========================================================
        ' THÊM CARD
        '=========================================================
        Private Sub AddCard(ByVal parent As GroupBox,
                            ByVal frm As Form,
                            ByVal value As Integer,
                            ByVal title As String,
                            ByVal desc As String,
                            ByVal top As Integer,
                            ByVal left As Integer,
                            ByVal width As Integer)

            Dim cardHeight As Integer = 70

            Dim pnl As New Panel()
            pnl.Location = New Drw.Point(left, top)
            pnl.Size = New Drw.Size(width, cardHeight)
            pnl.BackColor = Drw.Color.FromArgb(250, 250, 250)
            pnl.BorderStyle = BorderStyle.FixedSingle
            pnl.Cursor = Cursors.Hand
            parent.Controls.Add(pnl)

            '--- Tiêu đề ---
            Dim lblTitle As New Label()
            lblTitle.Text = title
            lblTitle.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblTitle.ForeColor = Drw.Color.FromArgb(30, 30, 30)
            lblTitle.Location = New Drw.Point(15, 8)
            lblTitle.Size = New Drw.Size(width - 25, 22)
            pnl.Controls.Add(lblTitle)

            '--- Mô tả ---
            Dim lblDesc As New Label()
            lblDesc.Text = desc
            lblDesc.Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            lblDesc.ForeColor = Drw.Color.FromArgb(110, 110, 110)
            lblDesc.Location = New Drw.Point(15, 34)
            lblDesc.Size = New Drw.Size(width - 25, 22)
            pnl.Controls.Add(lblDesc)

            '--- Hover ---
            Dim hoverOn As EventHandler = Sub()
                                              pnl.BackColor = Drw.Color.FromArgb(235, 242, 252)
                                          End Sub
            Dim hoverOff As EventHandler = Sub()
                                               pnl.BackColor = Drw.Color.FromArgb(250, 250, 250)
                                           End Sub

            AddHandler pnl.MouseEnter, hoverOn
            AddHandler pnl.MouseLeave, hoverOff
            AddHandler lblTitle.MouseEnter, hoverOn
            AddHandler lblTitle.MouseLeave, hoverOff
            AddHandler lblDesc.MouseEnter, hoverOn
            AddHandler lblDesc.MouseLeave, hoverOff

            '--- Click ---
            Dim clickH As EventHandler = Sub(sender, e)
                                             frm.Tag = value
                                             frm.DialogResult = DialogResult.OK
                                             frm.Close()
                                         End Sub
            AddHandler pnl.Click, clickH
            AddHandler lblTitle.Click, clickH
            AddHandler lblDesc.Click, clickH
        End Sub

    End Module

End Namespace