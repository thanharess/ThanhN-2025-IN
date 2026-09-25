Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.Drawdim

    Public Module draw_2

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowMainMenu()
                Case 1 : Draw_dim_base_line_a.OnExecute(Context)
                Case 2 : Draw_dim_chain_line_a.OnExecute(Context)
                Case 3 : Draw_dim_base_line_b.OnExecute(Context)
                Case 4 : Draw_dim_chain_line_b.OnExecute(Context)
                Case 5 : Draw_dim_hole_a.OnExecute(Context)
                Case 6 : Draw_dim_hole_b.OnExecute(Context)
                Case 7 : Draw_dim_hole_c.OnExecute(Context)
                Case 8 : Draw_dim_hole_d.OnExecute(Context)
                Case 9 : Draw_dim_hole_e.OnExecute(Context)
                Case 10 : Draw_dim_hole.OnExecute(Context)
                Case 11 : Draw_dim_base_line_c.OnExecute(Context)
                Case 12 : Draw_dim_chain_line_c.OnExecute(Context)
                Case 13 : Delete_small_dims.OnExecute(Context)
            End Select
        End Sub

        Private Function ShowMainMenu() As Integer
            Dim result As Integer = 0

            '=== KÍCH THƯỚC CARD — 3 CỘT ===
            Const CARD_W As Integer = 270
            Const CARD_H As Integer = 65
            Const GAP_X As Integer = 15
            Const PAD_LEFT As Integer = 15

            Using frm As New Form()
                frm.Text = "Auto Dim — Drawing"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(900, 760)
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
                pnlHeader.Size = New Drw.Size(900, 75)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "AUTO DIM — LINEAR & HOLE"
                lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Tự động ghi kích thước Base / Chain / Hole cho Drawing View"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== NHÓM 1: BASE / CHAIN DIM (4 card = 3+1) =====
                Dim gb1 As New GroupBox() With {
                    .Text = "Base / Chain Dim",
                    .Location = New Drw.Point(15, 90),
                    .Size = New Drw.Size(870, 180),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb1)

                AddCard(gb1, frm, 1, "Base Dim", "Dim theo đường cơ sở",
                        PAD_LEFT, 30, CARD_W)
                AddCard(gb1, frm, 2, "Chain Dim", "Dim theo chuỗi liên tục",
                        PAD_LEFT + (CARD_W + GAP_X), 30, CARD_W)
                AddCard(gb1, frm, 3, "Base+Hole", "Base Dim kèm Dim lỗ",
                        PAD_LEFT + 2 * (CARD_W + GAP_X), 30, CARD_W)
                AddCard(gb1, frm, 4, "Chain+Hole", "Chain Dim kèm Dim lỗ",
                        PAD_LEFT, 30 + CARD_H + 15, CARD_W)

                '===== NHÓM 2: DIM LỖ (8 card = 3+3+2) =====
                Dim gb2 As New GroupBox() With {
                    .Text = "Dim kích thước lỗ",
                    .Location = New Drw.Point(15, 280),
                    .Size = New Drw.Size(870, 255),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb2)

                '--- Hàng 1 ---
                AddCard(gb2, frm, 5,
                        "Dim kích thước lỗ",
                        "Ghi kích thước cho các lỗ",
                        PAD_LEFT, 30, CARD_W)

                AddCard(gb2, frm, 6,
                        "Dim lỗ về cạnh",
                        "Bỏ qua lỗ trùng + lỗ array",
                        PAD_LEFT + (CARD_W + GAP_X), 30, CARD_W)

                AddCard(gb2, frm, 7,
                        "Dim lỗ tương đối",
                        "Bỏ qua lỗ trùng + lỗ array",
                        PAD_LEFT + 2 * (CARD_W + GAP_X), 30, CARD_W)

                '--- Hàng 2 ---
                AddCard(gb2, frm, 8,
                        "Dim lỗ — không bỏ qua",
                        "Dim tất cả các lỗ",
                        PAD_LEFT, 30 + CARD_H + 15, CARD_W)

                AddCard(gb2, frm, 9,
                        "Dim lỗ Base Dimline",
                        "Set về cạnh",
                        PAD_LEFT + (CARD_W + GAP_X), 30 + CARD_H + 15, CARD_W)

                AddCard(gb2, frm, 10,
                        "Dim lỗ Base Dimline",
                        "Về cạnh",
                        PAD_LEFT + 2 * (CARD_W + GAP_X), 30 + CARD_H + 15, CARD_W)

                '--- Hàng 3 ---
                AddCard(gb2, frm, 11,
                        "Base + Hole + KC min",
                        "Base Dim + Hole + khoảng cách tối thiểu",
                        PAD_LEFT, 30 + 2 * (CARD_H + 15), CARD_W)

                AddCard(gb2, frm, 12,
                        "Chain + Hole + KC min",
                        "Chain Dim + Hole + khoảng cách tối thiểu",
                        PAD_LEFT + (CARD_W + GAP_X), 30 + 2 * (CARD_H + 15), CARD_W)

                '===== NHÓM 3: TIỆN ÍCH =====
                Dim gb3 As New GroupBox() With {
                    .Text = "Tiện ích",
                    .Location = New Drw.Point(15, 545),
                    .Size = New Drw.Size(870, 120),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb3)

                AddCard(gb3, frm, 13,
                        "Xóa dim nhỏ",
                        "Loại bỏ các dimension quá nhỏ",
                        PAD_LEFT, 30, CARD_W)

                Dim lblNote As New Label() With {
                    .Text = "★ Dọn dẹp các dimension nhỏ, không cần thiết trên bản vẽ" & vbCrLf &
                            "★ Giúp bản vẽ gọn gàng, dễ đọc hơn",
                    .Location = New Drw.Point(PAD_LEFT + (CARD_W + GAP_X), 40),
                    .Size = New Drw.Size(2 * CARD_W + GAP_X, 50),
                    .ForeColor = Drw.Color.FromArgb(140, 140, 140),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)}
                gb3.Controls.Add(lblNote)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(140, 42)
                btnCancel.Location = New Drw.Point(745, 680)
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
                            ByVal left As Integer,
                            ByVal top As Integer,
                            ByVal width As Integer)

            Dim pnl As New Panel()
            pnl.Location = New Drw.Point(left, top)
            pnl.Size = New Drw.Size(width, 65)
            pnl.BackColor = Drw.Color.FromArgb(250, 250, 250)
            pnl.BorderStyle = BorderStyle.FixedSingle
            pnl.Cursor = Cursors.Hand
            parent.Controls.Add(pnl)

            '--- Tiêu đề ---
            Dim lblTitle As New Label()
            lblTitle.Text = title
            lblTitle.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblTitle.ForeColor = Drw.Color.FromArgb(30, 30, 30)
            lblTitle.Location = New Drw.Point(15, 10)
            lblTitle.Size = New Drw.Size(width - 25, 22)
            pnl.Controls.Add(lblTitle)

            '--- Mô tả ---
            Dim lblDesc As New Label()
            lblDesc.Text = desc
            lblDesc.Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            lblDesc.ForeColor = Drw.Color.FromArgb(110, 110, 110)
            lblDesc.Location = New Drw.Point(15, 35)
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