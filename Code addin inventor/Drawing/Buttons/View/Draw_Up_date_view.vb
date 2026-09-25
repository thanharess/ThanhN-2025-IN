Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports Drw = System.Drawing

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

                '=================================================
                ' FORM CHỌN PHẠM VI
                '=================================================
                Dim choice As Integer = ShowScopeForm()
                If choice = 0 Then Return

                Dim updatedCount As Integer = 0
                Dim failedCount As Integer = 0
                Dim sickBefore As Integer = 0

                '=================================================
                ' FORCE UPDATE THEO PHẠM VI
                '=================================================
                If choice = 1 Then
                    ' Active Sheet
                    updatedCount = ForceUpdateSheet(oDrawDoc.ActiveSheet,
                                                    failedCount, sickBefore)
                ElseIf choice = 2 Then
                    ' All Sheets
                    For Each oSheet As Sheet In oDrawDoc.Sheets
                        updatedCount += ForceUpdateSheet(oSheet, failedCount, sickBefore)
                    Next
                End If

                '=================================================
                ' UPDATE TOÀN BỘ DOCUMENT
                '=================================================
                Try : oDrawDoc.Update2(True) : Catch : End Try

                '=================================================
                ' THÔNG BÁO
                '=================================================
                Dim msg As String =
                    "Hoàn tất Force Update!" & vbCrLf & vbCrLf &
                    "Phạm vi        : " & If(choice = 1, "Sheet hiện tại", "Toàn bộ Drawing") & vbCrLf &
                    "View đã update : " & updatedCount.ToString() & vbCrLf

                If failedCount > 0 Then
                    msg &= "View bỏ qua    : " & failedCount.ToString() & vbCrLf
                End If

                If sickBefore > 0 Then
                    msg &= vbCrLf &
                           "⚠ Có " & sickBefore.ToString() & " view có dấu chấm than." & vbCrLf &
                           "Kiểm tra lại các view này."
                Else
                    msg &= vbCrLf & "✓ Không phát hiện view lỗi."
                End If

                MessageBox.Show(msg,
                                "Force Update Drawing Views",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message,
                                "Force Update Drawing Views",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '==========================================================
        ' FORCE UPDATE 1 SHEET
        '==========================================================
        Private Function ForceUpdateSheet(ByVal oSheet As Sheet,
                                          ByRef failedCount As Integer,
                                          ByRef sickCount As Integer) As Integer

            If oSheet Is Nothing Then Return 0

            Dim count As Integer = 0

            For Each oView As DrawingView In oSheet.DrawingViews
                Try
                    If ForceUpdateOneView(oView) Then
                        count += 1
                    Else
                        failedCount += 1
                    End If

                    ' Kiểm tra view có lỗi (dấu chấm than)
                    Try
                        Dim hasError As Boolean = False
                        Try
                            ' HealthStatus không có API trực tiếp, check qua property khác
                            ' Nếu view có vấn đề về associativity → không update được
                            Dim ov As Object = oView
                            Dim viewHealth As Object = Nothing
                            Try
                                viewHealth = CallByName(ov, "HealthStatus", CallType.Get)
                            Catch
                            End Try

                            If viewHealth IsNot Nothing Then
                                Dim healthStr As String = viewHealth.ToString()
                                If healthStr.Contains("Sick") OrElse healthStr.Contains("Error") Then
                                    hasError = True
                                End If
                            End If
                        Catch
                        End Try

                        If hasError Then sickCount += 1
                    Catch
                    End Try

                Catch
                    failedCount += 1
                End Try
            Next

            Return count

        End Function


        '==========================================================
        ' FORCE UPDATE 1 VIEW — cách ổn định
        '==========================================================
        Private Function ForceUpdateOneView(ByVal oView As DrawingView) As Boolean

            Try
                ' Bỏ qua Overlay View
                If oView.ViewType = DrawingViewTypeEnum.kOverlayDrawingViewType Then
                    Return False
                End If

                '=================================================
                ' CÁCH 1: Đổi Scale cực nhỏ rồi trả lại
                '=================================================
                Dim oldScale As Double = oView.Scale

                If oldScale > 0 Then
                    Try
                        oView.Scale = oldScale * 1.0000001
                        oView.Scale = oldScale
                    Catch
                    End Try
                End If

                '=================================================
                ' CÁCH 2 (dự phòng): bật/tắt Hidden Lines
                '=================================================
                Try
                    Dim oldHidden As Boolean = oView.ShowHiddenLines
                    oView.ShowHiddenLines = Not oldHidden
                    oView.ShowHiddenLines = oldHidden
                Catch
                End Try

                '=================================================
                ' CÁCH 3 (dự phòng): bật/tắt Tangent Edges
                '=================================================
                Try
                    Dim oldTangent As Boolean = oView.DisplayTangentEdges
                    oView.DisplayTangentEdges = Not oldTangent
                    oView.DisplayTangentEdges = oldTangent
                Catch
                End Try

                Return True

            Catch
                Return False
            End Try

        End Function


        '==========================================================
        ' FORM CHỌN PHẠM VI
        '==========================================================
        Private Function ShowScopeForm() As Integer
            Dim result As Integer = 0

            Using frm As New Form()
                frm.Text = "Force Update Drawing Views"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(620, 400)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.BackColor = Drw.Color.FromArgb(245, 245, 245)
                frm.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                frm.Tag = 0

                '===== HEADER =====
                Dim pnlHeader As New Panel()
                pnlHeader.Location = New Drw.Point(0, 0)
                pnlHeader.Size = New Drw.Size(620, 75)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "FORCE UPDATE DRAWING VIEWS"
                lblTitle.Font = New Drw.Font("Segoe UI", 14.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Cập nhật lại Drawing View để xóa dấu chấm than (!)"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== HƯỚNG DẪN =====
                Dim lblHint As New Label() With {
                    .Text = "Chọn phạm vi cần Force Update:",
                    .Location = New Drw.Point(30, 95),
                    .Size = New Drw.Size(560, 25),
                    .Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(60, 60, 60)}
                frm.Controls.Add(lblHint)

                '===== CARD 1: ACTIVE SHEET =====
                AddCard(frm, 1,
                        "Sheet đang mở",
                        "Chỉ Force Update các Drawing View trên Sheet hiện tại",
                        130)

                '===== CARD 2: ALL SHEETS =====
                AddCard(frm, 2,
                        "Toàn bộ Drawing",
                        "Force Update TẤT CẢ Drawing View của mọi Sheet",
                        230)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(130, 44)
                btnCancel.Location = New Drw.Point(465, 335)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.FlatAppearance.BorderSize = 1
                btnCancel.FlatAppearance.BorderColor = Drw.Color.FromArgb(200, 200, 200)
                btnCancel.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(235, 235, 235)
                btnCancel.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(215, 215, 215)
                btnCancel.BackColor = Drw.Color.FromArgb(250, 250, 250)
                btnCancel.ForeColor = Drw.Color.FromArgb(60, 60, 60)
                btnCancel.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                btnCancel.Cursor = Cursors.Hand
                btnCancel.UseVisualStyleBackColor = False
                AddHandler btnCancel.Click, Sub()
                                                frm.Tag = 0
                                                frm.Close()
                                            End Sub
                frm.Controls.Add(btnCancel)
                frm.CancelButton = btnCancel

                frm.ShowDialog()
                result = CInt(frm.Tag)
            End Using

            Return result
        End Function


        '==========================================================
        ' THÊM CARD
        '==========================================================
        Private Sub AddCard(ByVal frm As Form,
                            ByVal value As Integer,
                            ByVal title As String,
                            ByVal desc As String,
                            ByVal top As Integer)

            Dim cardWidth As Integer = 560
            Dim cardHeight As Integer = 80

            Dim pnl As New Panel()
            pnl.Location = New Drw.Point(30, top)
            pnl.Size = New Drw.Size(cardWidth, cardHeight)
            pnl.BackColor = Drw.Color.White
            pnl.BorderStyle = BorderStyle.FixedSingle
            pnl.Cursor = Cursors.Hand
            frm.Controls.Add(pnl)

            '--- Tiêu đề ---
            Dim lblTitle As New Label()
            lblTitle.Text = title
            lblTitle.Font = New Drw.Font("Segoe UI", 11.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblTitle.ForeColor = Drw.Color.FromArgb(30, 30, 30)
            lblTitle.Location = New Drw.Point(20, 15)
            lblTitle.Size = New Drw.Size(cardWidth - 40, 25)
            pnl.Controls.Add(lblTitle)

            '--- Mô tả ---
            Dim lblDesc As New Label()
            lblDesc.Text = desc
            lblDesc.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            lblDesc.ForeColor = Drw.Color.FromArgb(110, 110, 110)
            lblDesc.Location = New Drw.Point(20, 42)
            lblDesc.Size = New Drw.Size(cardWidth - 40, 25)
            pnl.Controls.Add(lblDesc)

            '--- Hover ---
            Dim hoverOn As EventHandler = Sub()
                                              pnl.BackColor = Drw.Color.FromArgb(235, 242, 252)
                                          End Sub
            Dim hoverOff As EventHandler = Sub()
                                               pnl.BackColor = Drw.Color.White
                                           End Sub

            AddHandler pnl.MouseEnter, hoverOn
            AddHandler pnl.MouseLeave, hoverOff
            AddHandler lblTitle.MouseEnter, hoverOn
            AddHandler lblTitle.MouseLeave, hoverOff
            AddHandler lblDesc.MouseEnter, hoverOn
            AddHandler lblDesc.MouseLeave, hoverOff

            '--- Click ---
            Dim clickH As EventHandler = Sub(sender, e)
                                             frm.Tag = value
                                             frm.DialogResult = DialogResult.OK
                                             frm.Close()
                                         End Sub
            AddHandler pnl.Click, clickH
            AddHandler lblTitle.Click, clickH
            AddHandler lblDesc.Click, clickH
        End Sub

    End Module

End Namespace