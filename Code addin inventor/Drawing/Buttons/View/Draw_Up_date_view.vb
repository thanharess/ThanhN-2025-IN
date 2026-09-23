Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms

Namespace ToolInventor2025.Drawing.Buttons.DrawView

    Public Module Draw_Up_date_view

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                Dim oDrawDoc As DrawingDocument =
                    TryCast(g_inventorApplication.ActiveDocument, DrawingDocument)

                If oDrawDoc Is Nothing Then
                    MessageBox.Show("Document hiện tại không phải Drawing.",
                                    "Force Update Drawing Views",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim choice As String = ShowScopeDialog()
                If String.IsNullOrEmpty(choice) Then Return

                Dim count As Integer = 0

                If choice = "Active Sheet" Then
                    count = ForceUpdateActiveSheet(oDrawDoc)
                ElseIf choice = "All Sheets" Then
                    count = ForceUpdateAllSheets(oDrawDoc)
                End If

                ' Cập nhật toàn bộ document
                oDrawDoc.Update2(True)

                MessageBox.Show("Đã Force Update " & count.ToString() & " Drawing View(s)." & vbCrLf &
                                "Kiểm tra lại các view còn dấu chấm than không.",
                                "Force Update Drawing Views",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message,
                                "Force Update Drawing Views",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '==========================================================
        ' FORCE UPDATE ACTIVE SHEET
        '==========================================================
        Private Function ForceUpdateActiveSheet(ByVal oDrawDoc As DrawingDocument) As Integer

            Dim oSheet As Sheet = oDrawDoc.ActiveSheet
            If oSheet Is Nothing Then Return 0

            Dim count As Integer = 0

            For Each oView As DrawingView In oSheet.DrawingViews
                If ForceUpdateOneView(oView) Then
                    count += 1
                End If
            Next

            Return count

        End Function


        '==========================================================
        ' FORCE UPDATE ALL SHEETS
        '==========================================================
        Private Function ForceUpdateAllSheets(ByVal oDrawDoc As DrawingDocument) As Integer

            Dim count As Integer = 0

            For Each oSheet As Sheet In oDrawDoc.Sheets
                For Each oView As DrawingView In oSheet.DrawingViews
                    If ForceUpdateOneView(oView) Then
                        count += 1
                    End If
                Next
            Next

            Return count

        End Function


        '==========================================================
        ' FORCE UPDATE 1 VIEW (CÁCH ỔN ĐỊNH NHẤT)
        '==========================================================
        Private Function ForceUpdateOneView(ByVal oView As DrawingView) As Boolean

            Try
                ' Bỏ qua Overlay View
                If oView.ViewType = DrawingViewTypeEnum.kOverlayDrawingViewType Then
                    Return False
                End If

                ' ----- Cách 1: Đổi Scale cực nhỏ rồi trả lại -----
                Dim oldScale As Double = oView.Scale

                ' Chỉ thực hiện nếu Scale > 0
                If oldScale > 0 Then
                    oView.Scale = oldScale * 1.0000001
                    oView.Scale = oldScale
                End If

                ' ----- Cách 2 (dự phòng): bật/tắt Hidden Lines -----
                Try
                    Dim oldHidden As Boolean = oView.ShowHiddenLines
                    oView.ShowHiddenLines = Not oldHidden
                    oView.ShowHiddenLines = oldHidden
                Catch
                End Try

                Return True

            Catch
                Return False
            End Try

        End Function


        '==========================================================
        ' FORM CHỌN
        '==========================================================
        Private Function ShowScopeDialog() As String

            Dim frm As New Form()
            frm.Text = "Force Update Drawing Views"
            frm.Width = 300
            frm.Height = 170
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.ShowInTaskbar = False

            Dim lbl As New Label()
            lbl.Text = "Chọn phạm vi Force Update:"
            lbl.Left = 20
            lbl.Top = 15
            lbl.AutoSize = True

            Dim cbo As New ComboBox()
            cbo.Left = 20
            cbo.Top = 40
            cbo.Width = 240
            cbo.DropDownStyle = ComboBoxStyle.DropDownList
            cbo.Items.Add("Active Sheet")
            cbo.Items.Add("All Sheets")
            cbo.SelectedIndex = 0

            Dim btnOK As New Button()
            btnOK.Text = "OK"
            btnOK.Left = 100
            btnOK.Top = 80
            btnOK.Width = 70
            btnOK.DialogResult = DialogResult.OK

            Dim btnCancel As New Button()
            btnCancel.Text = "Cancel"
            btnCancel.Left = 180
            btnCancel.Top = 80
            btnCancel.Width = 70
            btnCancel.DialogResult = DialogResult.Cancel

            frm.Controls.Add(lbl)
            frm.Controls.Add(cbo)
            frm.Controls.Add(btnOK)
            frm.Controls.Add(btnCancel)

            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            If frm.ShowDialog() = DialogResult.OK Then
                Return cbo.SelectedItem.ToString()
            End If

            Return ""

        End Function

    End Module

End Namespace