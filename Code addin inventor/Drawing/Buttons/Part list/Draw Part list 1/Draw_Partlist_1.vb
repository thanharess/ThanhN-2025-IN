Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.DrawPartList

    Public Module Draw_Partlist_1

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowMainMenu()
                Case 1 : Draw_Partlist_1a.OnExecute(Context)
                Case 2 : Draw_Partlist_1b.OnExecute(Context)
                Case 3 : Draw_Partlist_1c.OnExecute(Context)
            End Select
        End Sub

        Private Function ShowMainMenu() As Integer
            Dim result As Integer = 0

            Using frm As New Form()
                frm.Text = "Ghi Partlist — Drawing"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(640, 480)
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
                pnlHeader.Size = New Drw.Size(640, 75)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "GHI PARTLIST"
                lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Ghi thêm / thay thông tin vào Partlist ENG, VIE"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== NHÓM: GHI PARTLIST =====
                Dim gb As New GroupBox() With {
                    .Text = "Chọn chức năng",
                    .Location = New Drw.Point(15, 90),
                    .Size = New Drw.Size(610, 290),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb)

                AddCard(gb, frm, 1,
                        "Ghi Partlist VIE",
                        "Ghi / cập nhật thông tin tiếng Việt vào Partlist",
                        25)

                AddCard(gb, frm, 2,
                        "Ghi Partlist ENG",
                        "Ghi / cập nhật thông tin tiếng Anh vào Partlist",
                        110)

                AddCard(gb, frm, 3,
                        "Ghi mã chi tiết cho Part Partlist",
                        "Thêm mã chi tiết vào Part trong bảng kê",
                        195)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(140, 42)
                btnCancel.Location = New Drw.Point(485, 425)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.FlatAppearance.BorderColor = Drw.Color.FromArgb(180, 180, 180)
                btnCancel.BackColor = Drw.Color.FromArgb(245, 245, 245)
                btnCancel.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
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
                            ByVal top As Integer)

            Dim cardWidth As Integer = 560
            Dim cardHeight As Integer = 70

            Dim pnl As New Panel()
            pnl.Location = New Drw.Point(25, top)
            pnl.Size = New Drw.Size(cardWidth, cardHeight)
            pnl.BackColor = Drw.Color.FromArgb(250, 250, 250)
            pnl.BorderStyle = BorderStyle.FixedSingle
            pnl.Cursor = Cursors.Hand
            parent.Controls.Add(pnl)

            '--- Tiêu đề ---
            Dim lblTitle As New Label()
            lblTitle.Text = title
            lblTitle.Font = New Drw.Font("Segoe UI", 11.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblTitle.ForeColor = Drw.Color.FromArgb(30, 30, 30)
            lblTitle.Location = New Drw.Point(20, 12)
            lblTitle.Size = New Drw.Size(cardWidth - 40, 25)
            pnl.Controls.Add(lblTitle)

            '--- Mô tả ---
            Dim lblDesc As New Label()
            lblDesc.Text = desc
            lblDesc.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            lblDesc.ForeColor = Drw.Color.FromArgb(110, 110, 110)
            lblDesc.Location = New Drw.Point(20, 38)
            lblDesc.Size = New Drw.Size(cardWidth - 40, 22)
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