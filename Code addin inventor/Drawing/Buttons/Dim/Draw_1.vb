Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports Drw = System.Drawing
Imports System.Collections
Imports System.Collections.Generic

Namespace ToolInventor2025.Drawing.Buttons

    Public Module Draw_1

        '=========================================================
        ' OPTIONS
        '=========================================================
        Private Class PrecisionOption
            Public Index As Integer = -1
            Public Label As String = ""
            Public Description As String = ""
            Public LinearPrec As Integer = 0
            Public AngularPrec As Integer = 0
            Public Cancelled As Boolean = True
        End Class


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

                '===== FORM CHỌN SỐ THẬP PHÂN =====
                Dim opt As PrecisionOption = ShowPrecisionForm()
                If opt Is Nothing OrElse opt.Cancelled Then Exit Sub

                Dim linearPrec As Integer = opt.LinearPrec
                Dim angularPrec As Integer = opt.AngularPrec

                Dim changedStyles As New HashSet(Of String)
                Dim dimCount As Integer = 0

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
                    "Số thập phân: " & opt.Label & vbCrLf &
                    "Dim đã quét: " & dimCount.ToString() & vbCrLf &
                    "Style đã đổi: " & changedStyles.Count.ToString() & vbCrLf & vbCrLf &
                    "(Style dùng chung → tất cả sheet đã áp dụng)",
                    "Đổi số Dim", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Đổi số Dim",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '=========================================================
        ' FORM CHỌN SỐ THẬP PHÂN
        '=========================================================
        Private Function ShowPrecisionForm() As PrecisionOption
            Dim opt As New PrecisionOption()
            Dim _okConfirmed As Boolean = False

            Using frm As New Form()
                frm.Text = "Đổi số Dim — Drawing"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(720, 520)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.BackColor = Drw.Color.FromArgb(245, 245, 245)
                frm.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

                '===== HEADER =====
                Dim pnlHeader As New Panel()
                pnlHeader.Location = New Drw.Point(0, 0)
                pnlHeader.Size = New Drw.Size(720, 75)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "ĐỔI SỐ THẬP PHÂN DIMENSION"
                lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Áp dụng cho TẤT CẢ các Sheet trong bản vẽ"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== GROUP: CHỌN SỐ THẬP PHÂN =====
                Dim gb As New GroupBox() With {
                    .Text = "Số chữ số thập phân",
                    .Location = New Drw.Point(15, 90),
                    .Size = New Drw.Size(690, 375),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb)

                '--- 5 card chọn ---
                AddPrecisionCard(gb, frm, 0,
                                 "0",
                                 "Không có số thập phân — Dim hiển thị số nguyên",
                                 25)

                AddPrecisionCard(gb, frm, 1,
                                 "0.1",
                                 "1 chữ số thập phân — ví dụ: 12.5",
                                 90)

                AddPrecisionCard(gb, frm, 2,
                                 "0.12",
                                 "2 chữ số thập phân — ví dụ: 12.50  (Khuyến nghị)",
                                 155)

                AddPrecisionCard(gb, frm, 3,
                                 "0.123",
                                 "3 chữ số thập phân — ví dụ: 12.500",
                                 220)

                AddPrecisionCard(gb, frm, 4,
                                 "0.1234",
                                 "4 chữ số thập phân — ví dụ: 12.5000",
                                 285)

                '===== NÚT =====
                Dim btnOK As New Button()
                btnOK.Text = "ÁP DỤNG"
                btnOK.Size = New Drw.Size(160, 44)
                btnOK.Location = New Drw.Point(535, 475)
                btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
                btnOK.ForeColor = Drw.Color.White
                btnOK.FlatStyle = FlatStyle.Flat
                btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                btnOK.Enabled = False

                ' ⭐ Đổi màu theo trạng thái Enabled
                AddHandler btnOK.EnabledChanged, Sub()
                                                     If btnOK.Enabled Then
                                                         btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
                                                         btnOK.ForeColor = Drw.Color.White
                                                     Else
                                                         btnOK.BackColor = Drw.Color.FromArgb(220, 220, 220)
                                                         btnOK.ForeColor = Drw.Color.FromArgb(140, 140, 140)
                                                     End If
                                                 End Sub

                ' Set màu ban đầu (đang disable)
                btnOK.BackColor = Drw.Color.FromArgb(220, 220, 220)
                btnOK.ForeColor = Drw.Color.FromArgb(140, 140, 140)

                AddHandler btnOK.Click, Sub()
                                            opt.Cancelled = False
                                            _okConfirmed = True
                                            frm.DialogResult = DialogResult.OK
                                            frm.Close()
                                        End Sub
                frm.Controls.Add(btnOK)

                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(120, 44)
                btnCancel.Location = New Drw.Point(400, 475)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                AddHandler btnCancel.Click, Sub()
                                                opt.Cancelled = True
                                                frm.DialogResult = DialogResult.Cancel
                                                frm.Close()
                                            End Sub
                frm.Controls.Add(btnCancel)

                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel

                AddHandler frm.FormClosing, Sub(sender, e)
                                                If Not _okConfirmed Then opt.Cancelled = True
                                            End Sub

                '===== LƯU CALLBACK VÀO TAG CỦA FORM =====
                ' Dùng Tag để truyền hàm enable nút OK xuống các card
                frm.Tag = New Action(Of Integer, String)(
                    Sub(idx As Integer, lbl As String)
                        opt.Index = idx
                        opt.Label = lbl

                        Select Case idx
                            Case 0 : opt.LinearPrec = 41729 : opt.AngularPrec = 42241
                            Case 1 : opt.LinearPrec = 41730 : opt.AngularPrec = 42242
                            Case 2 : opt.LinearPrec = 41731 : opt.AngularPrec = 42243
                            Case 3 : opt.LinearPrec = 41732 : opt.AngularPrec = 42244
                            Case 4 : opt.LinearPrec = 41733 : opt.AngularPrec = 42245
                        End Select

                        btnOK.Enabled = True
                    End Sub)

                frm.ShowDialog()
            End Using

            Return opt
        End Function


        '=========================================================
        ' THÊM CARD CHỌN SỐ THẬP PHÂN
        '=========================================================
        Private Sub AddPrecisionCard(ByVal parent As GroupBox,
                                     ByVal frm As Form,
                                     ByVal value As Integer,
                                     ByVal mainText As String,
                                     ByVal desc As String,
                                     ByVal top As Integer)

            Dim cardWidth As Integer = 650
            Dim cardHeight As Integer = 55

            Dim pnl As New Panel()
            pnl.Location = New Drw.Point(20, top)
            pnl.Size = New Drw.Size(cardWidth, cardHeight)
            pnl.BackColor = Drw.Color.White
            pnl.BorderStyle = BorderStyle.FixedSingle
            pnl.Cursor = Cursors.Hand
            parent.Controls.Add(pnl)

            '--- Text chính (số thập phân) ---
            Dim lblMain As New Label()
            lblMain.Text = mainText
            lblMain.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblMain.ForeColor = Drw.Color.FromArgb(45, 100, 180)
            lblMain.Location = New Drw.Point(20, 12)
            lblMain.Size = New Drw.Size(80, 32)
            lblMain.TextAlign = Drw.ContentAlignment.MiddleCenter
            pnl.Controls.Add(lblMain)

            '--- Đường kẻ dọc ---
            Dim sep As New Panel()
            sep.Location = New Drw.Point(115, 12)
            sep.Size = New Drw.Size(1, 32)
            sep.BackColor = Drw.Color.FromArgb(220, 220, 220)
            pnl.Controls.Add(sep)

            '--- Mô tả ---
            Dim lblDesc As New Label()
            lblDesc.Text = desc
            lblDesc.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            lblDesc.ForeColor = Drw.Color.FromArgb(60, 60, 60)
            lblDesc.Location = New Drw.Point(130, 18)
            lblDesc.Size = New Drw.Size(cardWidth - 150, 22)
            lblDesc.TextAlign = Drw.ContentAlignment.MiddleLeft
            pnl.Controls.Add(lblDesc)

            '--- Hover ---
            Dim hoverOn As EventHandler = Sub()
                                              pnl.BackColor = Drw.Color.FromArgb(235, 242, 252)
                                          End Sub
            Dim hoverOff As EventHandler = Sub()
                                               If pnl.Tag Is Nothing Then
                                                   pnl.BackColor = Drw.Color.White
                                               End If
                                           End Sub

            AddHandler pnl.MouseEnter, hoverOn
            AddHandler pnl.MouseLeave, hoverOff
            AddHandler lblMain.MouseEnter, hoverOn
            AddHandler lblMain.MouseLeave, hoverOff
            AddHandler lblDesc.MouseEnter, hoverOn
            AddHandler lblDesc.MouseLeave, hoverOff

            '--- Click: chọn card + đổi màu + bật nút OK ---
            Dim clickH As EventHandler = Sub(sender, e)
                                             ' Reset tất cả card khác
                                             For Each ctrl As Control In parent.Controls
                                                 If TypeOf ctrl Is Panel Then
                                                     ctrl.Tag = Nothing
                                                     ctrl.BackColor = Drw.Color.White
                                                 End If
                                             Next

                                             ' Đánh dấu card này đã chọn
                                             pnl.Tag = "selected"
                                             pnl.BackColor = Drw.Color.FromArgb(200, 220, 245)

                                             ' Gọi callback lưu giá trị
                                             Try
                                                 Dim cb = TryCast(frm.Tag, Action(Of Integer, String))
                                                 If cb IsNot Nothing Then cb.Invoke(value, mainText)
                                             Catch
                                             End Try
                                         End Sub

            AddHandler pnl.Click, clickH
            AddHandler lblMain.Click, clickH
            AddHandler lblDesc.Click, clickH
        End Sub

    End Module

End Namespace