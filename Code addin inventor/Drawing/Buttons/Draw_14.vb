Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports System.Drawing

Namespace ToolInventor2025.Drawing.Buttons

    Public Module draw_14

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                Dim oDrawDoc As DrawingDocument =
                    TryCast(g_inventorApplication.ActiveDocument, DrawingDocument)

                If oDrawDoc Is Nothing Then
                    MessageBox.Show("Document hiện tại không phải Drawing.",
                                    "Change Section Scale + Hatch",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim scaleText As String = ""
                Dim scope As String = ""

                If Not ShowInputDialog(scaleText, scope) Then Return

                Dim newScale As Double = ParseScale(scaleText)

                If newScale <= 0 Then
                    MessageBox.Show("Scale không hợp lệ! Ví dụ: 0.5 hoặc 1/2",
                                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim count As Integer = 0

                If scope = "Active Sheet" Then
                    count = ProcessSheet(oDrawDoc.ActiveSheet, newScale)
                Else
                    For Each oSheet As Sheet In oDrawDoc.Sheets
                        count += ProcessSheet(oSheet, newScale)
                    Next
                End If

                ' Force toàn bộ Drawing
                oDrawDoc.Update2(True)
                Try
                    oDrawDoc.Rebuild2(True)
                Catch
                End Try

                MessageBox.Show("Đã xử lý " & count.ToString() & " Mặt cắt." & vbCrLf &
                                "Scale mới: " & newScale.ToString() & vbCrLf &
                                "Nếu Hatch vẫn sai, hãy thử Save + đóng mở lại file.",
                                "Change Section Scale + Hatch",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message,
                                "Change Section Scale + Hatch",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '==========================================================
        ' XỬ LÝ 1 SHEET
        '==========================================================
        Private Function ProcessSheet(ByVal oSheet As Sheet, ByVal newScale As Double) As Integer

            If oSheet Is Nothing Then Return 0

            Dim count As Integer = 0

            For Each oView As DrawingView In oSheet.DrawingViews
                If ForceSectionScaleAndHatch(oView, newScale) Then
                    count += 1
                End If
            Next

            Return count

        End Function


        '==========================================================
        ' FORCE SCALE + HATCH MẠNH HƠN
        '==========================================================
        Private Function ForceSectionScaleAndHatch(ByVal oView As DrawingView, ByVal newScale As Double) As Boolean

            Try
                If oView.ViewType <> DrawingViewTypeEnum.kSectionDrawingViewType Then
                    Return False
                End If

                '----- 1. Đổi Scale -----
                oView.Scale = newScale

                '----- 2. Tắt hoàn toàn Hatch rồi bật lại -----
                Try
                    oView.ShowHatch = False
                Catch
                End Try

                '----- 3. Dirty bằng Hidden Lines -----
                Try
                    Dim oldHidden As Boolean = oView.ShowHiddenLines
                    oView.ShowHiddenLines = Not oldHidden
                    oView.ShowHiddenLines = oldHidden
                Catch
                End Try

                '----- 4. Dirty bằng Tangent Edges -----
                Try
                    Dim oldTangent As Boolean = oView.ShowTangentEdges
                    oView.ShowTangentEdges = Not oldTangent
                    oView.ShowTangentEdges = oldTangent
                Catch
                End Try

                '----- 5. Dịch vị trí cực nhỏ -----
                Try
                    Dim tg As TransientGeometry = g_inventorApplication.TransientGeometry
                    Dim oldPos As Point2d = oView.Position
                    oView.Position = tg.CreatePoint2d(oldPos.X + 0.002, oldPos.Y + 0.002)
                    oView.Position = oldPos
                Catch
                End Try

                '----- 6. Bật lại Hatch -----
                Try
                    oView.ShowHatch = True
                Catch
                End Try

                '----- 7. Đổi Scale lần nữa để chắc chắn -----
                oView.Scale = newScale

                Return True

            Catch
                Return False
            End Try

        End Function


        '==========================================================
        ' PARSE SCALE (hỗ trợ 0.5 và 1/2)
        '==========================================================
        Private Function ParseScale(ByVal text As String) As Double

            text = text.Trim().Replace(",", ".")

            If text.Contains("/") Then
                Dim parts() As String = text.Split("/"c)
                If parts.Length = 2 Then
                    Dim num, den As Double
                    If Double.TryParse(parts(0), num) AndAlso Double.TryParse(parts(1), den) AndAlso den <> 0 Then
                        Return num / den
                    End If
                End If
            End If

            Dim result As Double
            If Double.TryParse(text, result) Then
                Return result
            End If

            Return 0

        End Function


        '==========================================================
        ' FORM
        '==========================================================
        Private Function ShowInputDialog(ByRef scaleText As String, ByRef scope As String) As Boolean

            Dim frm As New Form()
            frm.Text = "Đổi Scale + Hatch Mặt cắt"
            frm.Size = New Size(360, 250)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.Font = New Font("Segoe UI", 9)

            Dim lbl1 As New Label() With {.Text = "Scale mới (0.5 hoặc 1/2):", .Location = New System.Drawing.Point(20, 20), .AutoSize = True}
            Dim txtScale As New System.Windows.Forms.TextBox() With {.Text = "0.5", .Location = New System.Drawing.Point(20, 45), .Width = 120}

            Dim lbl2 As New Label() With {.Text = "Phạm vi:", .Location = New System.Drawing.Point(20, 85), .AutoSize = True}
            Dim cboScope As New ComboBox() With {.Location = New System.Drawing.Point(20, 108), .Width = 300, .DropDownStyle = ComboBoxStyle.DropDownList}
            cboScope.Items.AddRange({"Active Sheet", "All Sheets"})
            cboScope.SelectedIndex = 0

            Dim btnOK As New Button() With {.Text = "OK", .Location = New System.Drawing.Point(160, 165), .Width = 70, .DialogResult = DialogResult.OK}
            Dim btnCancel As New Button() With {.Text = "Cancel", .Location = New System.Drawing.Point(240, 165), .Width = 70, .DialogResult = DialogResult.Cancel}

            frm.Controls.AddRange({lbl1, txtScale, lbl2, cboScope, btnOK, btnCancel})
            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            If frm.ShowDialog() = DialogResult.OK Then
                scaleText = txtScale.Text.Trim()
                scope = cboScope.SelectedItem.ToString()
                Return True
            End If

            Return False

        End Function

    End Module

End Namespace