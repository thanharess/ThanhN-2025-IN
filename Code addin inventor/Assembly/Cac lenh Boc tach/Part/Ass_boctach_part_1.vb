Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.caclenhboctach.part

    'Gán module này cho một nút Inventor. Các Ass_11, Ass_12, Ass_13 vẫn giữ nguyên.
    Public Module Ass_boctach_part_1

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowSheetMetalMenu()

                Case 1
                    Ass_boctach_part_1a.OnExecute(Context)
                Case 2
                    Ass_boctach_part_1b.OnExecute(Context)
                Case 3
                    Ass_boctach_part_1c.OnExecute(Context)
                Case 4
                    Ass_boctach_part_1d.OnExecute(Context)
                Case 5
                    Ass_boctach_part_1e.OnExecute(Context)
                Case 6
                    Ass_boctach_part_1f.OnExecute(Context)
                Case 7
                    Ass_boctach_part_1g.OnExecute(Context)
                Case 8
                    Ass_boctach_part_1h.OnExecute(Context)
            End Select
        End Sub

        Private Function ShowSheetMetalMenu() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Bóc tách Part"
                form.Width = 550
                form.Height = 540
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False

                Dim title As New Label() With {
                    .Text = "Bóc tách Theo lựa chọn phía dưới", .Left = 20, .Top = 15,
                    .Width = 500, .Height = 28
                }
                form.Controls.Add(title)

                form.Tag = 0
                AddMenuButton(form, "Top lever tách số loại sp vật tư mua, tiêu chuẩn. lọc part", 45, 1) 'ok
                AddMenuButton(form, "Top lever tách số loại sp PL & vật tư mua,.. lọc ASS, part", 95, 2) 'ok
                AddMenuButton(form, "Top lever Bóc tách số lượng tổng tấm. lọc part", 145, 3) 'ok
                AddMenuButton(form, "Top lever Bóc tách số lượng tổng cho PL & vật tư mua,.. lọc part", 195, 4) 'ok
                AddMenuButton(form, "Top lever Bóc tách số lượng tổng cho PL & vật tư mua,.. lọc ASS, part", 245, 5) 'ok
                AddMenuButton(form, "All lever Bóc tách số lượng tổng tấm lọc part", 295, 6) ' ok
                AddMenuButton(form, "All lever lọc các loại tấm. lọc part", 345, 7) 'ok
                AddMenuButton(form, "All lever Bóc tách số lượng tổng cho PL, vật tư mua,.. lọc part", 395, 8) 'ok


                Dim cancelButton As New Button() With {
                    .Text = "HỦY", .Left = 20, .Top = 445, .Width = 500, .Height = 32
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
                .Text = text, .Left = 20, .Top = top, .Width = 500, .Height = 42
            }
            AddHandler button.Click, Sub()
                                         form.Tag = value
                                         form.Close()
                                     End Sub
            form.Controls.Add(button)
        End Sub

    End Module

End Namespace