Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor

Namespace ToolInventor2025.Toolngoai

    Public Module Button_Assembly2

        Private Const FORM_WIDTH As Integer = 760
        Private Const BTN_H As Integer = 34
        Private Const BTN_GAP As Integer = 5
        Private Const COLS As Integer = 2
        Private Const GB_PADDING As Integer = 12

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowMenub()
                Case 1 : Assembly2.Buttons.BOMcode.Ass_Bom_1.OnExecute(Context)
                Case 2 : Assembly2.Buttons.BOMcode.Ass_Bom_2.OnExecute(Context)
                Case 3 : Assembly2.Buttons.BOMcode.Ass_Bom_3.OnExecute(Context)
                Case 4 : Assembly2.Buttons.BOMcode.Ass_Bom_4.OnExecute(Context)
                Case 5 : Assembly2.Buttons.BOMcode.ass_bom_6.OnExecute(Context)
                Case 6 : Assembly2.Buttons.BOMcode.ass_bom_7.OnExecute(Context)
            End Select
        End Sub

        Private Function ShowMenub() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Tool Inventor 2025  -  BOM Tools"
                form.Width = FORM_WIDTH
                form.Height = 750
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.BackColor = System.Drawing.Color.FromArgb(245, 248, 252)
                form.Font = New Font("Segoe UI", 9.0!)
                form.Tag = 0

                Dim banner As New Label() With {
                    .Text = "📋   TỔNG HỢP CÔNG CỤ BOM",
                    .Left = 0, .Top = 0, .Width = FORM_WIDTH, .Height = 52,
                    .TextAlign = ContentAlignment.MiddleCenter,
                    .Font = New Font("Segoe UI", 13.0!, FontStyle.Bold),
                    .BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                    .ForeColor = System.Drawing.Color.White
                }
                form.Controls.Add(banner)

                Dim y As Integer = 62

                y = AddGroup(form, "📝   Tên & Đánh số", y, 1, New String() {
                    "Thay tên, STT BOM",
                    "Thay đổi đuôi tên BOM"
                })

                y = AddGroup(form, "🛒   Item & Số lượng", y, 3, New String() {
                    "Item1 Buy Top level về cuối",
                    "Ghi Item1, Thickness, Số lượng",
                    "Ghi PL Part, Stock Number"
                })

                y = AddGroup(form, "🌲   Model Browser", y, 6, New String() {
                    "Xếp Model Browser ASS"
                })

                Dim cancel As New Button() With {
                    .Text = "✕   HỦY",
                    .Left = 15, .Top = y + 3, .Width = FORM_WIDTH - 30, .Height = 38,
                    .BackColor = System.Drawing.Color.FromArgb(230, 70, 70),
                    .ForeColor = System.Drawing.Color.White,
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Segoe UI", 9.5!, FontStyle.Bold),
                    .Cursor = Cursors.Hand
                }
                cancel.FlatAppearance.BorderSize = 0
                AddHandler cancel.Click, Sub() form.Close()
                form.Controls.Add(cancel)

                form.Height = y + 75
                form.ShowDialog()
                result = CInt(form.Tag)
            End Using

            Return result
        End Function

        Private Function AddGroup(form As Form, title As String, top As Integer,
                                  startValue As Integer, texts As String()) As Integer
            Const GB_WIDTH As Integer = FORM_WIDTH - 30

            Dim Cols2 As Integer = If(texts.Length = 1, 1, COLS)
            Dim rows As Integer = CInt(Math.Ceiling(texts.Length / CDbl(Cols2)))
            Dim gbHeight As Integer = 22 + rows * (BTN_H + BTN_GAP) + 8

            Dim gb As New GroupBox() With {
                .Text = "  " & title,
                .Left = 15, .Top = top, .Width = GB_WIDTH, .Height = gbHeight,
                .Font = New Font("Segoe UI", 9.5!, FontStyle.Bold),
                .ForeColor = System.Drawing.Color.FromArgb(0, 100, 170),
                .BackColor = System.Drawing.Color.Transparent
            }
            form.Controls.Add(gb)

            Dim colWidth As Integer = CInt((GB_WIDTH - 2 * GB_PADDING - (Cols2 - 1) * 6) / Cols2)

            For i As Integer = 0 To texts.Length - 1
                Dim col As Integer = i Mod Cols2
                Dim row As Integer = i \ Cols2
                Dim value As Integer = startValue + i

                Dim b As New Button() With {
                    .Text = value.ToString("00") & ".   " & texts(i),
                    .Left = GB_PADDING + col * (colWidth + 6),
                    .Top = 22 + row * (BTN_H + BTN_GAP),
                    .Width = colWidth,
                    .Height = BTN_H,
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .BackColor = System.Drawing.Color.White,
                    .ForeColor = System.Drawing.Color.FromArgb(30, 30, 30),
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Segoe UI", 9.0!, FontStyle.Regular),
                    .Cursor = Cursors.Hand,
                    .Padding = New Padding(8, 0, 0, 0)
                }
                b.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(200, 212, 224)
                b.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(217, 236, 255)
                b.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(180, 215, 250)
                AddHandler b.Click, Sub()
                                        form.Tag = value
                                        form.Close()
                                    End Sub
                gb.Controls.Add(b)
            Next

            Return top + gbHeight + 8
        End Function

    End Module
End Namespace