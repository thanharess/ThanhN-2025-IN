Imports Inventor
Imports System.Windows.Forms
Imports Drw = System.Drawing
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode

    Public Module ass_bom_7

        Public Sub OnExecute(ByVal Context As NameValueMap)
            RunBOMReplace()
        End Sub


        '=====================================================
        ' OPTIONS CLASS
        '=====================================================
        Private Class BomReplaceOptions
            Public UpdatePN As Boolean = False
            Public UpdateSN As Boolean = False
            Public AllLevels As Boolean = False
            Public FindText As String = ""
            Public ReplaceText As String = ""
            Public Cancelled As Boolean = True
        End Class


        '=====================================================
        ' HÀM CHÍNH
        '=====================================================
        Public Sub RunBOMReplace()

            Dim opt As BomReplaceOptions = ShowBomReplaceForm()
            If opt Is Nothing OrElse opt.Cancelled Then Exit Sub

            RunOnModelBrowser(opt.UpdatePN, opt.UpdateSN,
                              opt.FindText, opt.ReplaceText,
                              opt.AllLevels)

        End Sub


        '=====================================================
        ' FORM TÙY CHỌN — GỘP TẤT CẢ
        '=====================================================
        '=====================================================
        ' FORM TÙY CHỌN — GỘP TẤT CẢ
        '=====================================================
        Private Function ShowBomReplaceForm() As BomReplaceOptions
            Dim opt As New BomReplaceOptions()
            Dim _okConfirmed As Boolean = False

            Using frm As New Form()
                frm.Text = "BOM Replace — Tùy chọn"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(580, 640)
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
                pnlHeader.Size = New Drw.Size(580, 70)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "BOM REPLACE — TÌM & THAY THẾ"
                lblTitle.Font = New Drw.Font("Segoe UI", 14.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Tìm & thay thế chuỗi trong Part Number / Stock Number"
                lblSub.Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== GROUP 1: LOẠI PROPERTY =====
                Dim gb1 As New GroupBox() With {
                    .Text = "Loại muốn sửa",
                    .Location = New Drw.Point(15, 85),
                    .Size = New Drw.Size(550, 100),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb1)

                Dim rdoPN As New RadioButton() With {
                    .Text = "Part Number",
                    .Location = New Drw.Point(25, 30),
                    .Size = New Drw.Size(160, 25),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(rdoPN)

                Dim rdoSN As New RadioButton() With {
                    .Text = "Stock Number",
                    .Location = New Drw.Point(200, 30),
                    .Size = New Drw.Size(160, 25),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(rdoSN)

                Dim rdoBoth As New RadioButton() With {
                    .Text = "Cả hai",
                    .Location = New Drw.Point(375, 30),
                    .Size = New Drw.Size(150, 25),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(rdoBoth)

                Dim lblNote1 As New Label() With {
                    .Text = "Chọn một hoặc nhiều property",
                    .Location = New Drw.Point(25, 60),
                    .Size = New Drw.Size(500, 20),
                    .ForeColor = Drw.Color.FromArgb(140, 140, 140),
                    .Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(lblNote1)

                '===== GROUP 2: PHẠM VI =====
                Dim gb2 As New GroupBox() With {
                    .Text = "Phạm vi chạy",
                    .Location = New Drw.Point(15, 195),
                    .Size = New Drw.Size(550, 100),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb2)

                Dim rdoTop As New RadioButton() With {
                    .Text = "Top Level  —  chỉ cấp trên + Phantom",
                    .Location = New Drw.Point(25, 30),
                    .Size = New Drw.Size(500, 25),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb2.Controls.Add(rdoTop)

                Dim rdoAll As New RadioButton() With {
                    .Text = "All Levels  —  sửa tất cả các cấp",
                    .Location = New Drw.Point(25, 60),
                    .Size = New Drw.Size(500, 25),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb2.Controls.Add(rdoAll)

                '===== GROUP 3: TÌM & THAY =====
                Dim gb3 As New GroupBox() With {
                    .Text = "Tìm & Thay thế",
                    .Location = New Drw.Point(15, 305),
                    .Size = New Drw.Size(550, 210),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb3)

                Dim lblFind As New Label() With {
                    .Text = "Chữ cần tìm:",
                    .Location = New Drw.Point(25, 40),
                    .Size = New Drw.Size(150, 25),
                    .TextAlign = Drw.ContentAlignment.MiddleLeft,
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb3.Controls.Add(lblFind)

                Dim txtFind As New System.Windows.Forms.TextBox() With {
                    .Text = "mm",
                    .Location = New Drw.Point(180, 40),
                    .Size = New Drw.Size(340, 25),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb3.Controls.Add(txtFind)

                Dim lblReplace As New Label() With {
                    .Text = "Chữ thay thế:",
                    .Location = New Drw.Point(25, 90),
                    .Size = New Drw.Size(150, 25),
                    .TextAlign = Drw.ContentAlignment.MiddleLeft,
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb3.Controls.Add(lblReplace)

                Dim txtReplace As New System.Windows.Forms.TextBox() With {
                    .Text = "L",
                    .Location = New Drw.Point(180, 90),
                    .Size = New Drw.Size(340, 25),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb3.Controls.Add(txtReplace)

                Dim lblPreview As New Label() With {
                    .Text = "Ví dụ: A100mm x B200mm  →  A100L x B200L",
                    .Location = New Drw.Point(25, 140),
                    .Size = New Drw.Size(500, 40),
                    .ForeColor = Drw.Color.FromArgb(80, 80, 80),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)}
                gb3.Controls.Add(lblPreview)

                Dim updatePreview As Action = Sub()
                                                  Try
                                                      Dim f As String = txtFind.Text
                                                      Dim r As String = txtReplace.Text
                                                      If f <> "" Then
                                                          lblPreview.Text = "Ví dụ: A100" & f & " x B200" & f &
                                                                            "  →  A100" & r & " x B200" & r
                                                      Else
                                                          lblPreview.Text = "Nhập chữ cần tìm ở trên"
                                                      End If
                                                  Catch
                                                  End Try
                                              End Sub

                AddHandler txtFind.TextChanged, Sub() updatePreview()
                AddHandler txtReplace.TextChanged, Sub() updatePreview()

                '===== NÚT CHẠY =====
                Dim btnOK As New Button()
                btnOK.Text = "CHẠY"
                btnOK.Size = New Drw.Size(160, 44)
                btnOK.Location = New Drw.Point(400, 580)
                btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
                btnOK.ForeColor = Drw.Color.White
                btnOK.FlatStyle = FlatStyle.Flat
                btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)

                ' ⭐ NÚT CHẠY: VALIDATE + LẤY DỮ LIỆU + ĐÓNG FORM
                AddHandler btnOK.Click, Sub()
                                            ' Validate
                                            If String.IsNullOrWhiteSpace(txtFind.Text) Then
                                                MessageBox.Show("Chữ cần tìm không được để trống.", "Lỗi",
                                                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                txtFind.Focus()
                                                Return
                                            End If

                                            ' Lấy dữ liệu vào opt
                                            opt.FindText = txtFind.Text
                                            opt.ReplaceText = txtReplace.Text
                                            opt.AllLevels = rdoAll.Checked

                                            If rdoPN.Checked Then
                                                opt.UpdatePN = True
                                                opt.UpdateSN = False
                                            ElseIf rdoSN.Checked Then
                                                opt.UpdatePN = False
                                                opt.UpdateSN = True
                                            Else
                                                opt.UpdatePN = True
                                                opt.UpdateSN = True
                                            End If

                                            opt.Cancelled = False
                                            _okConfirmed = True

                                            frm.DialogResult = DialogResult.OK
                                            frm.Close()
                                        End Sub
                frm.Controls.Add(btnOK)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(120, 44)
                btnCancel.Location = New Drw.Point(265, 580)
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

                ' ⭐ Bắt nút X cũng là Cancel
                AddHandler frm.FormClosing, Sub(sender, e)
                                                If Not _okConfirmed Then
                                                    opt.Cancelled = True
                                                End If
                                            End Sub

                '===== HIỆN FORM =====
                frm.ShowDialog()
            End Using

            Return opt
        End Function


        Private Sub OnOKClick(ByVal sender As Object, ByVal e As EventArgs)
            ' Để trống — logic lấy dữ liệu ở ShowDialog
        End Sub


        '=====================================================
        ' LẤY INVENTOR APPLICATION
        '=====================================================
        Private Function GetInventorApp() As Inventor.Application
            Try
                Return DirectCast(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch
                MessageBox.Show("Không tìm thấy Inventor đang chạy!")
                Return Nothing
            End Try
        End Function


        '=====================================================
        ' CHẠY TRÊN MODEL BROWSER
        '=====================================================
        Private Sub RunOnModelBrowser(updatePartNumber As Boolean,
                                      updateStockNumber As Boolean,
                                      findText As String,
                                      replaceText As String,
                                      allLevels As Boolean)

            Dim invApp As Inventor.Application = GetInventorApp()
            If invApp Is Nothing Then Exit Sub

            If invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Hãy mở Assembly trước!")
                Exit Sub
            End If

            Dim asmDoc As AssemblyDocument = invApp.ActiveDocument
            Dim countChanged As Integer = 0

            For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                ProcessOccurrence(occ, updatePartNumber, updateStockNumber, findText, replaceText, allLevels, countChanged)
            Next

            asmDoc.Update2(True)
            asmDoc.Save2(True)

            MessageBox.Show("Hoàn tất!" & vbCrLf &
                            "Số lượng item đã sửa: " & countChanged.ToString() & vbCrLf &
                            "Tìm: """ & findText & """ → Thay: """ & replaceText & """" & vbCrLf &
                            "Chế độ: " & If(allLevels, "All Levels", "Top Level"),
                            "BOM Replace")
        End Sub


        '=====================================================
        ' XỬ LÝ TỪNG OCCURRENCE
        '=====================================================
        Private Sub ProcessOccurrence(occ As ComponentOccurrence,
                                      updatePartNumber As Boolean,
                                      updateStockNumber As Boolean,
                                      findText As String,
                                      replaceText As String,
                                      allLevels As Boolean,
                                      ByRef countChanged As Integer)

            Try
                If occ.Suppressed Then Exit Sub

                Dim doc As Document = Nothing
                Try
                    doc = occ.Definition.Document
                Catch
                    Exit Sub
                End Try

                Dim isPhantom As Boolean = False
                Try
                    If occ.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                        isPhantom = True
                    End If
                Catch
                End Try

                If isPhantom Then
                    If occ.SubOccurrences IsNot Nothing Then
                        For Each subOcc As ComponentOccurrence In occ.SubOccurrences
                            ProcessOccurrence(subOcc, updatePartNumber, updateStockNumber, findText, replaceText, allLevels, countChanged)
                        Next
                    End If
                    Exit Sub
                End If

                UpdateDocumentProps(doc, updatePartNumber, updateStockNumber, findText, replaceText, countChanged)

                If allLevels Then
                    If occ.SubOccurrences IsNot Nothing Then
                        For Each subOcc As ComponentOccurrence In occ.SubOccurrences
                            ProcessOccurrence(subOcc, updatePartNumber, updateStockNumber, findText, replaceText, allLevels, countChanged)
                        Next
                    End If
                End If

            Catch
            End Try
        End Sub


        '=====================================================
        ' SỬA iPROPERTIES CỦA DOCUMENT
        '=====================================================
        Private Sub UpdateDocumentProps(doc As Document,
                                        updatePartNumber As Boolean,
                                        updateStockNumber As Boolean,
                                        findText As String,
                                        replaceText As String,
                                        ByRef countChanged As Integer)

            Try
                Dim designProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim changed As Boolean = False

                If updatePartNumber Then
                    Try
                        Dim prop As Inventor.Property = designProps.Item("Part Number")
                        Dim oldVal As String = prop.Value.ToString()

                        If oldVal.IndexOf(findText, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Dim newVal As String = Regex.Replace(oldVal, Regex.Escape(findText), replaceText, RegexOptions.IgnoreCase)
                            prop.Value = newVal
                            changed = True
                        End If
                    Catch
                    End Try
                End If

                If updateStockNumber Then
                    Try
                        Dim prop As Inventor.Property = designProps.Item("Stock Number")
                        Dim oldVal As String = prop.Value.ToString()

                        If oldVal.IndexOf(findText, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Dim newVal As String = Regex.Replace(oldVal, Regex.Escape(findText), replaceText, RegexOptions.IgnoreCase)
                            prop.Value = newVal
                            changed = True
                        End If
                    Catch
                    End Try
                End If

                If changed Then
                    doc.Save2(True)
                    countChanged += 1
                End If

            Catch
            End Try
        End Sub

    End Module

End Namespace