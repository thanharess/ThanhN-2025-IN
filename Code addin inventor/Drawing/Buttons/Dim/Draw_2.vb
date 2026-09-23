Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Drawing.Buttons.Drawdim

    Public Module draw_2

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowSheetMetalMenu()
                Case 1
                    Draw_dim_base_line_a.OnExecute(Context)
                Case 2
                    Draw_dim_chain_line_a.OnExecute(Context)
                Case 3
                    Draw_dim_base_line_b.OnExecute(Context)
                Case 4
                    Draw_dim_chain_line_b.OnExecute(Context)
                Case 5
                    Draw_dim_hole_a.OnExecute(Context)
                Case 6
                    Draw_dim_hole_b.OnExecute(Context)
                Case 7
                    Draw_dim_hole_c.OnExecute(Context)
                Case 8
                    Draw_dim_hole_d.OnExecute(Context)
                Case 9
                    Draw_dim_hole_e.OnExecute(Context)
                Case 10
                    Draw_dim_hole.OnExecute(Context)
                Case 11
                    Draw_dim_base_line_c.OnExecute(Context)
                Case 12
                    Draw_dim_chain_line_c.OnExecute(Context)
                Case 13
                    Delete_small_dims.OnExecute(Context)
                Case 14
                    '  Draw_dim_chain_line_d.OnExecute(Context)
            End Select
        End Sub

        Private Function ShowSheetMetalMenu() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Auto dim"
                form.Width = 530
                form.Height = 600
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False

                Dim title As New Label() With {
                    .Text = "Auto dim Liner - Hole", .Left = 20, .Top = 15,
                    .Width = 470, .Height = 28
                }
                form.Controls.Add(title)

                form.Tag = 0
                AddMenuButton(form, "Base Dim view Auto", 50, 1)
                AddMenuButton(form, "Chain Dim view Auto", 85, 2)
                AddMenuButton(form, "Base Dim view Auto + Hole", 120, 3)
                AddMenuButton(form, "Chain Dim view Auto + Hole", 155, 4)
                AddMenuButton(form, "Dim kích thước lỗ", 190, 5)
                AddMenuButton(form, "Dim lỗ Dim về cạnh (Dim ít có bỏ qua lỗ nếu trùng, bỏ qua lỗ array)", 225, 6)
                AddMenuButton(form, "Dim lỗ Dim về cạnh (Dim tương đối có bỏ qua lỗ nếu trùng, bỏ qua lỗ array)", 260, 7)
                AddMenuButton(form, "Dim lỗ Dim về cạnh (Nhiều dim ko bỏ qua lỗ, dim tất cả các lỗ)", 295, 8)
                AddMenuButton(form, "Dim lỗ Base Dimline Set về cạnh", 330, 9)
                AddMenuButton(form, "Dim lỗ Base Dimline về cạnh", 365, 10)
                AddMenuButton(form, "Base Dim view Auto + Hole + KC tối thiểu)", 400, 11)
                AddMenuButton(form, "Chain Dim view Auto + Hole + KC tối thiểu)", 435, 12)
                AddMenuButton(form, "xóa dim", 470, 13)
                '    AddMenuButton(form, "Dim lỗ Base Dimline về cạnh", 505, 14)

                Dim cancelButton As New Button() With {
                    .Text = "HỦY", .Left = 20, .Top = 510, .Width = 470, .Height = 30
                }
                AddHandler cancelButton.Click, Sub() form.Close()
                form.Controls.Add(cancelButton)

                form.ShowDialog()
                result = CInt(form.Tag)
            End Using

            Return result
        End Function

        Private Sub AddMenuButton(ByVal form As Form, ByVal text As String, ByVal top As Integer, ByVal value As Integer)
            Dim button As New Button() With {
                .Text = text, .Left = 20, .Top = top, .Width = 470, .Height = 30
            }
            AddHandler button.Click, Sub()
                                         form.Tag = value
                                         form.Close()
                                     End Sub
            form.Controls.Add(button)
        End Sub

    End Module

End Namespace