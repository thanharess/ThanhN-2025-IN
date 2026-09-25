Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports Drw = System.Drawing

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep.constraint

    '=============================================================
    ' MENU CONSTRAINTS
    '=============================================================
    Public Module Ass_LG_C_1

        '── Màu sắc / kích thước ──
        Private ReadOnly COLOR_HEADER As Drw.Color = Drw.Color.FromArgb(45, 100, 180)
        Private ReadOnly COLOR_HEADER_SUB As Drw.Color = Drw.Color.FromArgb(220, 230, 245)
        Private ReadOnly COLOR_BG As Drw.Color = Drw.Color.FromArgb(245, 245, 245)
        Private ReadOnly COLOR_BTN As Drw.Color = Drw.Color.White
        Private ReadOnly COLOR_BTN_HOVER As Drw.Color = Drw.Color.FromArgb(235, 242, 252)
        Private ReadOnly COLOR_BTN_DOWN As Drw.Color = Drw.Color.FromArgb(215, 230, 250)
        Private ReadOnly COLOR_BTN_BORDER As Drw.Color = Drw.Color.FromArgb(210, 220, 235)
        Private ReadOnly COLOR_TEXT As Drw.Color = Drw.Color.FromArgb(40, 40, 40)

        Private Const FORM_W As Integer = 520
        Private Const FORM_H As Integer = 500
        Private Const HEADER_H As Integer = 75
        Private Const PAD As Integer = 20
        Private Const BTN_H As Integer = 48
        Private Const BTN_GAP As Integer = 10


        '=============================================================
        ' ENTRY
        '=============================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowSheetMetalMenu()
                Case 1
                    Ass_LG_C_1a.OnExecute(Context)
                Case 2
                    Ass_LG_C_1b.OnExecute(Context)
                Case 3
                    Ass_LG_C_1c.OnExecute(Context)
                Case 4
                    Ass_LG_C_1d.OnExecute(Context)
                Case 5
                    Ass_LG_C_1e.OnExecute(Context)
            End Select
        End Sub


        '=============================================================
        ' FORM
        '=============================================================
        Private Function ShowSheetMetalMenu() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                '── FORM ──
                form.Text = "Constraints Assembly"
                form.AutoScaleMode = AutoScaleMode.Dpi
                form.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                form.ClientSize = New Drw.Size(FORM_W, FORM_H)
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.ShowInTaskbar = False
                form.BackColor = COLOR_BG
                form.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

                '── HEADER ──
                Dim pnlHeader As New Panel() With {
                    .Location = New Drw.Point(0, 0),
                    .Size = New Drw.Size(FORM_W, HEADER_H),
                    .BackColor = COLOR_HEADER
                }
                form.Controls.Add(pnlHeader)

                Dim lblTitle As New Label() With {
                    .Text = "CONSTRAINTS ASSEMBLY",
                    .Font = New Drw.Font("Segoe UI", 14.0F, Drw.FontStyle.Bold),
                    .ForeColor = Drw.Color.White,
                    .Dock = DockStyle.Fill,
                    .TextAlign = Drw.ContentAlignment.MiddleCenter
                }
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label() With {
                    .Text = "Chọn chức năng ràng buộc",
                    .Font = New Drw.Font("Segoe UI", 9.0F),
                    .ForeColor = COLOR_HEADER_SUB,
                    .Dock = DockStyle.Bottom,
                    .Height = 20,
                    .TextAlign = Drw.ContentAlignment.MiddleCenter
                }
                pnlHeader.Controls.Add(lblSub)

                '── MENU BUTTONS ──
                Dim y As Integer = HEADER_H + PAD

                AddMenuButton(form, "Suppress, Constrain, Ground", 1, y)
                y += BTN_H + BTN_GAP

                AddMenuButton(form, "Constrain Keep Position", 2, y)
                y += BTN_H + BTN_GAP

                AddMenuButton(form, "Constrain về gốc 2 chi tiết", 3, y)
                y += BTN_H + BTN_GAP

                AddMenuButton(form, "Constrain All to Select", 4, y)
                y += BTN_H + BTN_GAP

                AddMenuButton(form, "Xóa tất cả Constraint lỗi", 5, y)

                '── NÚT HỦY ──
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(FORM_W - PAD * 2, 42)
                btnCancel.Location = New Drw.Point(PAD, FORM_H - 42 - PAD)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.FlatAppearance.BorderSize = 1
                btnCancel.FlatAppearance.BorderColor = Drw.Color.FromArgb(200, 200, 200)
                btnCancel.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(235, 235, 235)
                btnCancel.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(215, 215, 215)
                btnCancel.BackColor = Drw.Color.FromArgb(250, 250, 250)
                btnCancel.ForeColor = Drw.Color.FromArgb(60, 60, 60)
                btnCancel.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular)
                btnCancel.Cursor = Cursors.Hand
                btnCancel.UseVisualStyleBackColor = False
                btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
                AddHandler btnCancel.Click, Sub() form.Close()
                form.Controls.Add(btnCancel)

                form.CancelButton = btnCancel

                form.ShowDialog()
                result = CInt(form.Tag)
            End Using

            Return result
        End Function


        '=============================================================
        ' TẠO 1 NÚT MENU
        '=============================================================
        Private Sub AddMenuButton(ByVal form As Form,
                                   ByVal text As String,
                                   ByVal value As Integer,
                                   ByVal top As Integer)

            Dim btn As New Button()
            btn.Text = text
            btn.Left = PAD
            btn.Top = top
            btn.Width = FORM_W - PAD * 2
            btn.Height = BTN_H
            btn.TextAlign = Drw.ContentAlignment.MiddleLeft
            btn.Padding = New Padding(20, 0, 0, 0)
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 1
            btn.FlatAppearance.BorderColor = COLOR_BTN_BORDER
            btn.FlatAppearance.MouseOverBackColor = COLOR_BTN_HOVER
            btn.FlatAppearance.MouseDownBackColor = COLOR_BTN_DOWN
            btn.BackColor = COLOR_BTN
            btn.ForeColor = COLOR_TEXT
            btn.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Regular)
            btn.Cursor = Cursors.Hand
            btn.UseVisualStyleBackColor = False

            AddHandler btn.Click,
                Sub()
                    form.Tag = value
                    form.Close()
                End Sub

            AddHandler btn.MouseEnter,
                Sub()
                    btn.FlatAppearance.BorderColor = COLOR_HEADER
                End Sub

            AddHandler btn.MouseLeave,
                Sub()
                    btn.FlatAppearance.BorderColor = COLOR_BTN_BORDER
                End Sub

            form.Controls.Add(btn)
        End Sub

    End Module

End Namespace