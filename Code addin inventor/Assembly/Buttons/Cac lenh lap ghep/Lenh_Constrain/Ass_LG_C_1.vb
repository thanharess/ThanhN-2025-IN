Option Explicit On
Option Strict Off
Imports System.Diagnostics.Contracts
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor
Imports ToolInventor2025.ToolInventor2025.Assembly.Buttons.caclenhboctach.part

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep.constraint


    'Gán module này cho một nút Inventor. Các Ass_11, Ass_12, Ass_13 vẫn giữ nguyên.
    Public Module Ass_LG_C_1

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

        Private Function ShowSheetMetalMenu() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Constraints Assembly"
                form.Width = 500
                form.Height = 410
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False

                Dim title As New Label() With {
                        .Text = "Constraints", .Left = 20, .Top = 15,
                        .Width = 450, .Height = 28
                    }
                form.Controls.Add(title)

                form.Tag = 0
                AddMenuButton(form, "Suppress, contrain, Ground", 45, 1) 'ok
                AddMenuButton(form, "Contrain Keep position", 95, 2) 'ok
                AddMenuButton(form, "Contrain về gốc 2 chi tiết", 145, 3) 'ok
                AddMenuButton(form, "Contrain all to select", 195, 4) 'ok
                AddMenuButton(form, "Xóa all Constrain lỗi", 245, 5) 'ok

                Dim cancelButton As New Button() With {
                        .Text = "HỦY", .Left = 20, .Top = 310, .Width = 450, .Height = 32
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
                    .Text = text, .Left = 20, .Top = top, .Width = 450, .Height = 42
                }
            AddHandler button.Click, Sub()
                                         form.Tag = value
                                         form.Close()
                                     End Sub
            form.Controls.Add(button)
        End Sub

    End Module

End Namespace