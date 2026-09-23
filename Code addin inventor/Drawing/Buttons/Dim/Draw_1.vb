Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports System.Collections
Imports System.Collections.Generic
Namespace ToolInventor2025.Drawing.Buttons
    Public Module Draw_1


        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Vui lòng mở file Drawing (.idw)!", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As Inventor.DrawingDocument =
                    CType(app.ActiveDocument, Inventor.DrawingDocument)

                Dim idx As Integer = PickFromList("Thay đổi số Dim All View", New String() {
                    "0",
                    "0.1",
                    "0.12",
                    "0.123",
                    "0.1234"
                }, 2)
                If idx < 0 Then Exit Sub

                Dim linearPrec As Integer
                Dim angularPrec As Integer
                Select Case idx
                    Case 0 : linearPrec = 41729 : angularPrec = 42241
                    Case 1 : linearPrec = 41730 : angularPrec = 42242
                    Case 2 : linearPrec = 41731 : angularPrec = 42243
                    Case 3 : linearPrec = 41732 : angularPrec = 42244
                    Case 4 : linearPrec = 41733 : angularPrec = 42245
                    Case Else : Exit Sub
                End Select

                Dim changedStyles As New HashSet(Of String)
                Dim dimCount As Integer = 0

                ' Style là dùng chung → quét mọi sheet chỉ để lấy đủ style đang dùng
                For Each sh As Inventor.Sheet In oDrawDoc.Sheets
                    For Each oDim As Inventor.DrawingDimension In sh.DrawingDimensions
                        dimCount += 1
                        Try
                            If TypeOf oDim Is Inventor.GeneralDimension Then
                                Dim gDim As Inventor.GeneralDimension = CType(oDim, Inventor.GeneralDimension)
                                Dim oStyle As Inventor.DimensionStyle = gDim.Style
                                If oStyle IsNot Nothing AndAlso Not changedStyles.Contains(oStyle.Name) Then
                                    oStyle.LinearPrecision = linearPrec
                                    oStyle.AngularPrecision = angularPrec
                                    changedStyles.Add(oStyle.Name)
                                End If
                            End If
                        Catch
                        End Try
                    Next
                Next

                oDrawDoc.Update()

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf &
                    "Số thập phân: " & (New String() {"0", "0.1", "0.12", "0.123", "0.1234"})(idx) & vbCrLf &
                    "Dim đã quét: " & dimCount.ToString() & vbCrLf &
                    "Style đã đổi: " & changedStyles.Count.ToString() & vbCrLf & vbCrLf &
                    "(Style dùng chung → tất cả sheet đã áp dụng)",
                    "Đổi số Dim", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Đổi số Dim",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        Private Function PickFromList(title As String, items As String(), Optional defaultIndex As Integer = 0) As Integer
            Dim frm As New Form()
            frm.Text = title
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.Width = 400
            frm.Height = 300
            frm.ShowInTaskbar = False

            Dim lst As New ListBox()
            lst.Left = 12 : lst.Top = 12 : lst.Width = 360 : lst.Height = 200
            For Each s As String In items
                lst.Items.Add(s)
            Next
            If defaultIndex >= 0 AndAlso defaultIndex < lst.Items.Count Then
                lst.SelectedIndex = defaultIndex
            ElseIf lst.Items.Count > 0 Then
                lst.SelectedIndex = 0
            End If

            Dim btnOK As New Button() With {.Text = "OK", .Left = 200, .Top = 225, .Width = 80, .DialogResult = DialogResult.OK}
            Dim btnCancel As New Button() With {.Text = "Hủy", .Left = 290, .Top = 225, .Width = 80, .DialogResult = DialogResult.Cancel}

            frm.Controls.Add(lst)
            frm.Controls.Add(btnOK)
            frm.Controls.Add(btnCancel)
            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            If frm.ShowDialog() <> DialogResult.OK OrElse lst.SelectedIndex < 0 Then Return -1
            Return lst.SelectedIndex
        End Function

    End Module

End Namespace