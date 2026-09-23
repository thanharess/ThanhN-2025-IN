Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Drawing.Buttons.DrawPartList

    Public Module Draw_Partlist_1

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowSheetMetalMenu()

                Case 1
                    Draw_Partlist_1a.OnExecute(Context)
                Case 2
                    Draw_Partlist_1b.OnExecute(Context)
                Case 3
                    Draw_Partlist_1c.OnExecute(Context)
            End Select
        End Sub

        Private Function ShowSheetMetalMenu() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Ghi Partlist"
                form.Width = 550
                form.Height = 540
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False

                Dim title As New Label() With {
                    .Text = "Ghi thêm, thay thông tin vào Partlist ENG,VIE", .Left = 20, .Top = 15,
                    .Width = 500, .Height = 28
                }
                form.Controls.Add(title)

                form.Tag = 0
                AddMenuButton(form, "Ghi Partlist VIE", 45, 1) 'ok
                AddMenuButton(form, "Ghi Partlist ENG", 95, 2) 'ok
                AddMenuButton(form, "Ghi mã chi tiết cho Part Partlist", 145, 3) 'ok
                'AddMenuButton(form, "Top lever Bóc tách số lượng tổng cho PL & vật tư mua,.. lọc part", 195, 4) 'ok
                'AddMenuButton(form, "Top lever Bóc tách số lượng tổng cho PL & vật tư mua,.. lọc ASS, part", 245, 5) 'ok
                'AddMenuButton(form, "All lever Bóc tách số lượng tổng tấm lọc part", 295, 6) ' ok
                ' AddMenuButton(form, "All lever lọc các loại tấm. lọc part", 345, 7) 'ok
                '  AddMenuButton(form, "All lever Bóc tách số lượng tổng cho PL, vật tư mua,.. lọc part", 395, 8) 'ok

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