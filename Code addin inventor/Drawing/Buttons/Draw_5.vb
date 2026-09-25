Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons

    Public Module Draw_5

        '=========================================================
        ' STATE cho InteractionEvents (pick + quét chuột)
        '=========================================================
        Private _ie As InteractionEvents
        Private _se As SelectEvents
        Private _collected As List(Of Object)
        Private _userDone As Boolean


        Public Sub OnExecute(ByVal Context As NameValueMap)

            '=====================================================
            ' LẤY INVENTOR APP
            '=====================================================
            Dim inventorApp As Inventor.Application = Nothing
            Try
                inventorApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"),
                                    Inventor.Application)
            Catch ex As Exception
                MessageBox.Show("Không tìm thấy Inventor đang chạy." & vbCrLf & ex.Message,
                                "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            '=====================================================
            ' KIỂM TRA DRAWING
            '=====================================================
            If inventorApp.ActiveDocument Is Nothing OrElse
               inventorApp.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                MessageBox.Show("Vui lòng mở file Drawing (.idw)!", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim oDrawDoc As DrawingDocument = CType(inventorApp.ActiveDocument, DrawingDocument)
            Dim oSheet As Sheet = oDrawDoc.ActiveSheet

            '=====================================================
            ' FORM CHỌN CHỨC NĂNG
            '=====================================================
            Dim choice As Integer = ShowCenterMarkMenu()
            If choice = 0 Then Exit Sub

            Dim deletedCount As Integer = 0
            Dim failCount As Integer = 0

            Select Case choice

                '-------------------------------------------------
                ' 1. CHỌN NHIỀU ĐƯỜNG TÂM ĐỂ XÓA (CLICK + QUÉT CHUỘT)
                '-------------------------------------------------
                Case 1

                    _collected = New List(Of Object)()
                    _userDone = False

                    Try : oDrawDoc.SelectSet.Clear() : Catch : End Try

                    Try
                        _ie = inventorApp.CommandManager.CreateInteractionEvents()
                        _ie.InteractionDisabled = False

                        _se = _ie.SelectEvents
                        _se.AddSelectionFilter(SelectionFilterEnum.kDrawingCentermarkFilter)
                        _se.WindowSelectEnabled = True       ' ⭐ cho phép quét chuột

                        AddHandler _se.OnSelect, AddressOf OnSelectHandler
                        AddHandler _ie.OnTerminate, AddressOf OnTerminateHandler

                        _ie.Start()
                    Catch ex As Exception
                        MessageBox.Show("Không khởi động được chế độ chọn: " & ex.Message,
                                        "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        CleanupInteraction()
                        Exit Sub
                    End Try

                    _ie.StatusBarText = "Click hoặc QUÉT CHUỘT để chọn đường tâm. ESC để kết thúc & xóa."

                    '--- Vòng lặp chờ user chọn xong ---
                    Do While Not _userDone
                        System.Windows.Forms.Application.DoEvents()
                        System.Threading.Thread.Sleep(20)
                    Loop

                    CleanupInteraction()

                    '--- Xóa các đối tượng đã chọn ---
                    If _collected IsNot Nothing AndAlso _collected.Count > 0 Then
                        For Each obj As Object In _collected
                            Try
                                If obj IsNot Nothing Then
                                    obj.Delete()
                                    deletedCount += 1
                                End If
                            Catch
                                failCount += 1
                            End Try
                        Next
                    End If

                    '--- Clear selection sau khi xóa ---
                    Try : oDrawDoc.SelectSet.Clear() : Catch : End Try

                '-------------------------------------------------
                ' 2. XÓA TẤT CẢ ĐƯỜNG TÂM + DẤU TÂM TRÊN SHEET
                '-------------------------------------------------
                Case 2
                    If MessageBox.Show(
                        "Xóa TẤT CẢ đường tâm + dấu tâm trên Sheet '" & oSheet.Name & "'?" & vbCrLf & vbCrLf &
                        "Thao tác này không thể hoàn tác!",
                        "Xác nhận",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning) <> DialogResult.Yes Then
                        Exit Sub
                    End If

                    '----- 1. Xóa Centerline -----
                    Try
                        Dim listCL As New List(Of Centerline)
                        For Each cl As Centerline In oSheet.Centerlines
                            listCL.Add(cl)
                        Next
                        For Each cl As Centerline In listCL
                            Try
                                cl.Delete()
                                deletedCount += 1
                            Catch
                                failCount += 1
                            End Try
                        Next
                    Catch
                    End Try

                    '----- 2. Xóa Centermark -----
                    Try
                        Dim listCM As New List(Of Centermark)
                        For Each cm As Centermark In oSheet.Centermarks
                            listCM.Add(cm)
                        Next
                        For Each cm As Centermark In listCM
                            Try
                                cm.Delete()
                                deletedCount += 1
                            Catch
                                failCount += 1
                            End Try
                        Next
                    Catch
                    End Try
            End Select

            '=====================================================
            ' UPDATE + THÔNG BÁO
            '=====================================================
            Try : oDrawDoc.Update() : Catch : End Try

            Dim msg As String =
                "Hoàn tất!" & vbCrLf & vbCrLf &
                "Đã xóa: " & deletedCount.ToString() & vbCrLf

            If failCount > 0 Then
                msg &= "Lỗi / bỏ qua: " & failCount.ToString()
            End If

            MessageBox.Show(msg, "Xóa đường tâm",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub


        '=========================================================
        ' CALLBACK: user chọn (click hoặc quét chuột)
        '=========================================================
        Private Sub OnSelectHandler(
            ByVal JustSelectedEntities As ObjectsEnumerator,
            ByVal SelectionDevice As SelectionDeviceEnum,
            ByVal ModelPosition As Inventor.Point,
            ByVal ViewPosition As Point2d,
            ByVal View As Inventor.View)

            Try
                If JustSelectedEntities Is Nothing Then Exit Sub

                For Each obj As Object In JustSelectedEntities
                    If obj Is Nothing Then Continue For
                    If _collected.Contains(obj) Then Continue For
                    _collected.Add(obj)
                Next

                '--- Cập nhật status bar ---
                Try
                    If _ie IsNot Nothing Then
                        _ie.StatusBarText = "Đã chọn " & _collected.Count.ToString() &
                                            " đường tâm. Chọn tiếp hoặc ESC để xóa."
                    End If
                Catch
                End Try
            Catch
            End Try
        End Sub


        '=========================================================
        ' CALLBACK: user kết thúc (ESC / chuột phải → Done)
        '=========================================================
        Private Sub OnTerminateHandler()
            _userDone = True
        End Sub


        '=========================================================
        ' CLEANUP InteractionEvents
        '=========================================================
        Private Sub CleanupInteraction()
            Try
                If _se IsNot Nothing Then
                    RemoveHandler _se.OnSelect, AddressOf OnSelectHandler
                End If
            Catch
            End Try

            Try
                If _ie IsNot Nothing Then
                    RemoveHandler _ie.OnTerminate, AddressOf OnTerminateHandler
                End If
            Catch
            End Try

            Try
                If _ie IsNot Nothing Then
                    _ie.Stop()
                End If
            Catch
            End Try

            _se = Nothing
            _ie = Nothing
        End Sub


        '=========================================================
        ' FORM CHỌN CHỨC NĂNG
        '=========================================================
        Private Function ShowCenterMarkMenu() As Integer
            Dim result As Integer = 0

            Using frm As New Form()
                frm.Text = "Xóa đường tâm — Drawing"
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
                lblTitle.Text = "XÓA ĐƯỜNG TÂM (CENTER MARK)"
                lblTitle.Font = New Drw.Font("Segoe UI", 14.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Chọn đường tâm cần xóa hoặc xóa toàn bộ Sheet"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== CARD 1 =====
                AddCard(frm, 1,
                        "Chọn đường tâm cần xóa (CLICK + QUÉT CHUỘT)",
                        "Click từng đường tâm hoặc quét chuột chọn nhiều. ESC để xóa",
                        100)

                '===== CARD 2 =====
                AddCard(frm, 2,
                        "Xóa TẤT CẢ đường tâm trên Sheet",
                        "Xóa toàn bộ đường tâm trên Sheet đang mở",
                        200)

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


        '=========================================================
        ' THÊM CARD
        '=========================================================
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
            lblTitle.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
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