Option Explicit On
Option Strict Off

    Imports System.Windows.Forms
    Imports Inventor


Namespace ToolInventor2025.Assembly.Buttons.Part

    'Gán module này cho một nút Inventor. Các Ass_11, Ass_12, Ass_13 vẫn giữ nguyên.
    Module Ass_Part_list_1

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowSheetMetalMenu()

                Case 1
                    Ass_Part_2.OnExecute(Context)
                Case 2
                    Ass_Part_3.OnExecute(Context)
                Case 3
                    Ass_Part_6.OnExecute(Context)
                Case 4
                    Ass_Part_9.OnExecute(Context)
                Case 5
                    Ass_Part_8.OnExecute(Context)
                    ' Case 6
                    '   Ass_boctach_part_1f.OnExecute(Context)
                    ' Case 7
                    '     Ass_boctach_part_1g.OnExecute(Context)
                    '  Case 8
                    '       Ass_boctach_part_1h.OnExecute(Context)
            End Select
        End Sub

        Private Function ShowSheetMetalMenu() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Auto Drawing"
                form.Width = 550
                form.Height = 540
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False

                Dim title As New Label() With {
                    .Text = "Chọn kiểu tạo bản vẽ", .Left = 20, .Top = 15,
                    .Width = 490, .Height = 28
                }
                form.Controls.Add(title)

                form.Tag = 0
                AddMenuButton(form, "Đổi đơn vị Part", 45, 1) 'ok
                AddMenuButton(form, "Đổi vật liệu từ generic thành steel Part", 95, 2) 'ok
                AddMenuButton(form, "Xoá màu ghi đè lên part", 145, 3) 'ok
                AddMenuButton(form, "Thay màu part", 195, 4) 'ok
                AddMenuButton(form, "Thông số part", 245, 5) 'ok
                'AddMenuButton(form, "All lever Bóc tách số lượng tổng tấm lọc part", 295, 6) ' ok
                'AddMenuButton(form, "All lever lọc các loại tấm. lọc part", 345, 7) 'ok
                ' AddMenuButton(form, "All lever Bóc tách số lượng tổng cho PL, vật tư mua,.. lọc part", 395, 8) 'ok


                Dim cancelButton As New Button() With {
                    .Text = "Hủy", .Left = 20, .Top = 445, .Width = 490, .Height = 32
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