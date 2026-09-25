Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports Drw = System.Drawing
Imports ToolInventor2025.ToolInventor2025.Assembly.Buttons

Namespace ToolInventor2025.Drawing.Buttons.DrawPartList

    Public Module Draw_Partlist_1c

        ' Property Set ID của Document Summary Information
        Private Const DOC_SUMMARY_PROPSET As String = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}"
        ' PropId của Category = 2
        Private Const CATEGORY_PROPID As Long = 2


        '=========================================================
        ' OPTIONS CLASS
        '=========================================================
        Private Class CategoryOptions
            Public Prefix As String = ""
            Public WriteToBOM As Boolean = False
            Public Scope As Integer = 1     ' 1=first, 2=all on sheet, 3=all drawing
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

                Dim oDrawDoc As Inventor.DrawingDocument = CType(app.ActiveDocument, Inventor.DrawingDocument)
                Dim oSheet As Inventor.Sheet = oDrawDoc.ActiveSheet

                If oSheet.PartsLists.Count < 1 Then
                    MessageBox.Show("Sheet hiện tại không có Parts List.", "Thông báo",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                '=================================================
                ' FORM TÙY CHỌN — GỘP 3 BƯỚC
                '=================================================
                Dim opt As CategoryOptions = ShowOptionsForm()
                If opt Is Nothing OrElse opt.Cancelled Then Exit Sub

                Dim PREFIX As String = opt.Prefix
                Dim writeToBom As Boolean = opt.WriteToBOM
                Dim scopeIdx As Integer = opt.Scope - 1

                Dim processed As Integer = 0
                Dim totalRows As Integer = 0

                If scopeIdx = 0 Then
                    Try
                        Dim oPartList As Inventor.PartsList = oSheet.PartsLists.Item(1)
                        totalRows += ProcessOnePartsList(oPartList, PREFIX, writeToBom)
                        processed += 1
                    Catch ex As Exception
                        MessageBox.Show("Parts List 1:" & vbCrLf & ex.Message, "Cảnh báo",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Try

                ElseIf scopeIdx = 1 Then
                    For plIdx As Integer = 1 To oSheet.PartsLists.Count
                        Try
                            Dim oPartList As Inventor.PartsList = oSheet.PartsLists.Item(plIdx)
                            totalRows += ProcessOnePartsList(oPartList, PREFIX, writeToBom)
                            processed += 1
                        Catch exPL As Exception
                            MessageBox.Show("Sheet: " & oSheet.Name & vbCrLf &
                                            "Parts List: " & plIdx.ToString() & vbCrLf &
                                            exPL.Message, "Cảnh báo",
                                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End Try
                    Next

                ElseIf scopeIdx = 2 Then
                    For sheetIdx As Integer = 1 To oDrawDoc.Sheets.Count
                        Try
                            Dim oCurSheet As Inventor.Sheet = oDrawDoc.Sheets.Item(sheetIdx)

                            For plIdx As Integer = 1 To oCurSheet.PartsLists.Count
                                Try
                                    Dim oPartList As Inventor.PartsList = oCurSheet.PartsLists.Item(plIdx)
                                    totalRows += ProcessOnePartsList(oPartList, PREFIX, writeToBom)
                                    processed += 1
                                Catch exPL As Exception
                                    MessageBox.Show("Sheet: " & oCurSheet.Name & vbCrLf &
                                                    "Parts List: " & plIdx.ToString() & vbCrLf &
                                                    exPL.Message, "Cảnh báo",
                                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                End Try
                            Next
                        Catch exSheet As Exception
                            MessageBox.Show("Lỗi Sheet " & sheetIdx.ToString() & ":" & vbCrLf &
                                            exSheet.Message, "Cảnh báo",
                                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End Try
                    Next
                End If

                Try : oDrawDoc.Update() : Catch : End Try

                MessageBox.Show("Hoàn tất!" & vbCrLf &
                                "Parts List đã xử lý: " & processed.ToString() & vbCrLf &
                                "Số dòng đã ghi mã: " & totalRows.ToString() & vbCrLf &
                                "Định dạng: " & PREFIX & "STT" & vbCrLf &
                                "Ghi BOM gốc: " & If(writeToBom, "Có", "Không"),
                                "Ghi mã bản vẽ (Item → Category)",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "Ghi mã bản vẽ (Item → Category)",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '=========================================================
        ' FORM TÙY CHỌN — GỘP 3 BƯỚC
        '=========================================================
        Private Function ShowOptionsForm() As CategoryOptions
            Dim opt As New CategoryOptions()
            Dim _okConfirmed As Boolean = False

            Using frm As New Form()
                frm.Text = "Ghi mã bản vẽ — Tùy chọn"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(660, 560)
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
                pnlHeader.Size = New Drw.Size(660, 75)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "GHI MÃ BẢN VẼ — ITEM → CATEGORY"
                lblTitle.Font = New Drw.Font("Segoe UI", 14.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Ghi mã dạng PREFIX+STT vào cột Category của Parts List"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== GROUP 1: PREFIX =====
                Dim gb1 As New GroupBox() With {
                    .Text = "1. PREFIX mã bản vẽ",
                    .Location = New Drw.Point(15, 90),
                    .Size = New Drw.Size(630, 100),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb1)

                Dim lblPrefix As New Label() With {
                    .Text = "Nhập PREFIX:",
                    .Location = New Drw.Point(25, 38),
                    .Size = New Drw.Size(130, 25),
                    .TextAlign = Drw.ContentAlignment.MiddleLeft,
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(lblPrefix)

                Dim txtPrefix As New System.Windows.Forms.TextBox() With {
                    .Text = "",
                    .Location = New Drw.Point(160, 38),
                    .Size = New Drw.Size(440, 25),
                    .Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(txtPrefix)

                Dim lblHint As New Label() With {
                    .Text = "Ví dụ: 1.2.3.  →  mã sẽ ghi 1.2.3.1, 1.2.3.2, 1.2.3.3...",
                    .Location = New Drw.Point(160, 68),
                    .Size = New Drw.Size(440, 20),
                    .ForeColor = Drw.Color.FromArgb(140, 140, 140),
                    .Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(lblHint)

                '===== GROUP 2: GHI BOM =====
                Dim gb2 As New GroupBox() With {
                    .Text = "2. Ghi xuống BOM gốc?",
                    .Location = New Drw.Point(15, 200),
                    .Size = New Drw.Size(630, 130),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb2)

                Dim rdoPartsOnly As New RadioButton() With {
                    .Text = "Chỉ ghi trên Parts List (không đụng BOM)",
                    .Location = New Drw.Point(25, 32),
                    .Size = New Drw.Size(580, 25),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb2.Controls.Add(rdoPartsOnly)

                Dim lblNote1 As New Label() With {
                    .Text = "→ Mã chỉ xuất hiện trên bản vẽ, không ảnh hưởng file Part/Assembly",
                    .Location = New Drw.Point(45, 55),
                    .Size = New Drw.Size(560, 20),
                    .ForeColor = Drw.Color.FromArgb(140, 140, 140),
                    .Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)}
                gb2.Controls.Add(lblNote1)

                Dim rdoWriteBom As New RadioButton() With {
                    .Text = "Ghi cả Parts List + BOM gốc (Category iProperty)",
                    .Location = New Drw.Point(25, 82),
                    .Size = New Drw.Size(580, 25),
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb2.Controls.Add(rdoWriteBom)

                Dim lblNote2 As New Label() With {
                    .Text = "→ Ghi vào iProperty Category của từng Part/Assembly",
                    .Location = New Drw.Point(45, 105),
                    .Size = New Drw.Size(560, 20),
                    .ForeColor = Drw.Color.FromArgb(140, 140, 140),
                    .Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)}
                gb2.Controls.Add(lblNote2)

                '===== GROUP 3: PHẠM VI =====
                Dim gb3 As New GroupBox() With {
                    .Text = "3. Phạm vi áp dụng",
                    .Location = New Drw.Point(15, 340),
                    .Size = New Drw.Size(630, 130),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb3)

                Dim rdoScope(2) As RadioButton
                Dim scopeTexts As String() = {
                    "Chỉ Parts List đầu tiên trên Sheet đang mở",
                    "Tất cả Parts List trên Sheet đang mở",
                    "Tất cả Parts List của toàn bộ Drawing"
                }

                For i As Integer = 0 To 2
                    Dim rdo As New RadioButton() With {
                        .Text = scopeTexts(i),
                        .Location = New Drw.Point(25, 30 + i * 30),
                        .Size = New Drw.Size(590, 25),
                        .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point),
                        .Checked = (i = 0)}
                    gb3.Controls.Add(rdo)
                    rdoScope(i) = rdo
                Next

                '===== NÚT THỰC HIỆN =====
                Dim btnOK As New Button()
                btnOK.Text = "THỰC HIỆN"
                btnOK.Size = New Drw.Size(170, 46)
                btnOK.Location = New Drw.Point(475, 490)
                btnOK.FlatStyle = FlatStyle.Flat
                btnOK.FlatAppearance.BorderSize = 0
                btnOK.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(60, 115, 195)
                btnOK.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(30, 80, 155)
                btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
                btnOK.ForeColor = Drw.Color.White
                btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                btnOK.Cursor = Cursors.Hand
                btnOK.UseVisualStyleBackColor = False

                AddHandler btnOK.Click, Sub()
                                            If String.IsNullOrWhiteSpace(txtPrefix.Text) Then
                                                MessageBox.Show("PREFIX không được để trống.",
                                                                "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                txtPrefix.Focus()
                                                Return
                                            End If

                                            opt.Prefix = txtPrefix.Text.Trim()
                                            opt.WriteToBOM = rdoWriteBom.Checked

                                            For i As Integer = 0 To 2
                                                If rdoScope(i).Checked Then
                                                    opt.Scope = i + 1
                                                    Exit For
                                                End If
                                            Next

                                            opt.Cancelled = False
                                            _okConfirmed = True
                                            frm.DialogResult = DialogResult.OK
                                            frm.Close()
                                        End Sub
                frm.Controls.Add(btnOK)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(130, 46)
                btnCancel.Location = New Drw.Point(335, 490)
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

                frm.ShowDialog()
            End Using

            Return opt
        End Function


        '=========================================================
        ' CÁC HÀM BÊN DƯỚI GIỮ NGUYÊN 100%
        '=========================================================
        ' ProcessOnePartsList, WriteCategoryToBOM, SetCategoryProperty,
        ' FindItemColumn, FindCategoryColumn, GetCellValue, SetCell,
        ' PickFromList

        ' [Dán lại các hàm này từ code cũ của bạn vào đây]


        '=========================================================
        ' XỬ LÝ 1 PARTS LIST
        '=========================================================
        Private Function ProcessOnePartsList(
            oPartList As Inventor.PartsList,
            PREFIX As String,
            writeToBom As Boolean) As Integer

            Dim count As Integer = 0

            Dim cItem As String = FindItemColumn(oPartList)
            Dim cCategory As String = FindCategoryColumn(oPartList)

            If cItem = "" Then
                Throw New Exception("Không tìm thấy cột Item (kItemPartsListProperty) trên Parts List.")
            End If

            If cCategory = "" Then
                Throw New Exception("Không tìm thấy cột Category trên Parts List." & vbCrLf &
                                    "Hãy thêm cột Category vào Parts List trước.")
            End If

            For i As Integer = 1 To oPartList.PartsListRows.Count

                Try
                    Dim row As Inventor.PartsListRow = oPartList.PartsListRows.Item(i)

                    Dim itemValue As String = GetCellValue(row, cItem)
                    If itemValue = "" Then Continue For

                    Dim newCode As String = PREFIX & itemValue
                    Dim currentCategory As String = GetCellValue(row, cCategory)

                    ' 1. Ghi trên Parts List
                    If Not String.Equals(currentCategory, newCode, StringComparison.OrdinalIgnoreCase) Then
                        SetCell(row, cCategory, newCode)
                        count += 1
                    End If

                    ' 2. Ghi xuống BOM gốc (Category iProperty của component)
                    If writeToBom Then
                        WriteCategoryToBOM(row, newCode)
                    End If

                Catch
                End Try

            Next

            Try
                oPartList.Update()
            Catch
            End Try

            ' Nếu ghi BOM thì SaveItemOverridesToBOM (đồng bộ override nếu cần)
            If writeToBom Then
                Try
                    oPartList.SaveItemOverridesToBOM()
                Catch
                End Try
            End If

            Return count

        End Function

        '=========================================================
        ' GHI CATEGORY XUỐNG DOCUMENT GỐC (BOM)
        '=========================================================
        Private Sub WriteCategoryToBOM(row As Inventor.PartsListRow, newCode As String)

            Try
                If row.ReferencedRows Is Nothing OrElse row.ReferencedRows.Count < 1 Then
                    Exit Sub
                End If

                Dim bomRow As Inventor.BOMRow = row.ReferencedRows.Item(1).BOMRow
                If bomRow Is Nothing Then Exit Sub

                If bomRow.ComponentDefinitions.Count < 1 Then Exit Sub

                Dim refDoc As Inventor.Document = bomRow.ComponentDefinitions.Item(1).Document
                If refDoc Is Nothing Then Exit Sub

                ' Ghi vào iProperty Category (Document Summary Information)
                SetCategoryProperty(refDoc, newCode)

            Catch
            End Try

        End Sub

        '=========================================================
        ' SET CATEGORY iPROPERTY
        '=========================================================
        Private Sub SetCategoryProperty(doc As Inventor.Document, value As String)

            If doc Is Nothing OrElse value Is Nothing Then Exit Sub

            Dim newValue As String = value.Trim()
            If newValue = "" Then Exit Sub

            Try
                Dim ps As Inventor.PropertySet = doc.PropertySets.Item(DOC_SUMMARY_PROPSET)
                Dim prop As Inventor.Property = ps.ItemByPropId(CATEGORY_PROPID)

                Dim oldValue As String = ""
                Try
                    If prop.Value IsNot Nothing Then
                        oldValue = CStr(prop.Value).Trim()
                    End If
                Catch
                End Try

                If String.Equals(oldValue, newValue, StringComparison.OrdinalIgnoreCase) Then
                    Exit Sub
                End If

                prop.Value = newValue

                Try
                    doc.Update()
                Catch
                End Try

            Catch
                ' Fallback: thử theo tên
                Try
                    Dim ps2 As Inventor.PropertySet = doc.PropertySets.Item("Inventor Document Summary Information")
                    ps2.Item("Category").Value = newValue
                    Try
                        doc.Update()
                    Catch
                    End Try
                Catch
                End Try
            End Try

        End Sub

        '=========================================================
        ' TÌM CỘT ITEM theo PropertyType
        '=========================================================
        Private Function FindItemColumn(pl As Inventor.PartsList) As String

            Try
                For Each col As Inventor.PartsListColumn In pl.PartsListColumns
                    Try
                        If col.PropertyType = PropertyTypeEnum.kItemPartsListProperty Then
                            Return col.Title
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            Return ""

        End Function

        '=========================================================
        ' TÌM CỘT CATEGORY theo PropertyType + PropId
        '=========================================================
        Private Function FindCategoryColumn(pl As Inventor.PartsList) As String

            Try
                For Each col As Inventor.PartsListColumn In pl.PartsListColumns
                    Try
                        If col.PropertyType = PropertyTypeEnum.kFileProperty Then

                            Dim propSetId As String = ""
                            Dim propId As Long = 0

                            col.GetFilePropertyId(propSetId, propId)

                            If String.Equals(propSetId, DOC_SUMMARY_PROPSET, StringComparison.OrdinalIgnoreCase) AndAlso
                               propId = CATEGORY_PROPID Then

                                Return col.Title
                            End If

                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            Return ""

        End Function

        '=========================================================
        ' GET / SET CELL
        '=========================================================
        Private Function GetCellValue(row As Inventor.PartsListRow, colName As String) As String
            If colName = "" Then Return ""

            Try
                Dim v As Object = row.Item(colName).Value
                If v Is Nothing Then Return ""
                Return CStr(v).Trim()
            Catch
                Return ""
            End Try
        End Function

        Private Sub SetCell(row As Inventor.PartsListRow, colName As String, value As String)

            If colName = "" OrElse value Is Nothing Then Exit Sub

            Dim newValue As String = value.Trim()
            If newValue = "" Then Exit Sub

            Try
                Dim cell As Inventor.PartsListCell = row.Item(colName)

                Dim oldValue As String = ""
                Try
                    If cell.Value IsNot Nothing Then
                        oldValue = CStr(cell.Value).Trim()
                    End If
                Catch
                End Try

                If String.Equals(oldValue, newValue, StringComparison.OrdinalIgnoreCase) Then
                    Exit Sub
                End If

                cell.Value = newValue

                Try
                    cell.Static = True
                Catch
                End Try

            Catch
            End Try

        End Sub

        '=========================================================
        ' PICK LIST
        '=========================================================
        Private Function PickFromList(title As String, items As String(), Optional defaultIndex As Integer = 0) As Integer

            Dim frm As New Form()
            frm.Text = title
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.Width = 520
            frm.Height = 320
            frm.ShowInTaskbar = False

            Dim lst As New ListBox()
            lst.Left = 12
            lst.Top = 12
            lst.Width = 480
            lst.Height = 220

            For Each s As String In items
                lst.Items.Add(s)
            Next

            If defaultIndex >= 0 AndAlso defaultIndex < lst.Items.Count Then
                lst.SelectedIndex = defaultIndex
            ElseIf lst.Items.Count > 0 Then
                lst.SelectedIndex = 0
            End If

            Dim btnOK As New Button() With {
                .Text = "OK",
                .Left = 320,
                .Top = 245,
                .Width = 80,
                .DialogResult = DialogResult.OK
            }

            Dim btnCancel As New Button() With {
                .Text = "Hủy",
                .Left = 410,
                .Top = 245,
                .Width = 80,
                .DialogResult = DialogResult.Cancel
            }

            frm.Controls.Add(lst)
            frm.Controls.Add(btnOK)
            frm.Controls.Add(btnCancel)

            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            If frm.ShowDialog() <> DialogResult.OK OrElse lst.SelectedIndex < 0 Then
                Return -1
            End If

            Return lst.SelectedIndex

        End Function

    End Module

End Namespace