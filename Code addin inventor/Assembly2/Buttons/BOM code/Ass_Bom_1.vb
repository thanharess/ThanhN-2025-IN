Option Explicit On
Option Strict Off

        Imports Inventor
        Imports System.Windows.Forms
        Imports System.Drawing
        Imports System.Collections.Generic
        Imports System.Runtime.InteropServices

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode

        Public Module Ass_Bom_1

            Private ReadOnly BearingKeywords As String() = {
                "vòng bi", "vong bi", "bearing", "motor",
                "gối bi", "goi bi", "gối đỡ", "goi do",
                "pillow", "plummer", "ucp", "ucf", "ucfl", "khóa trục", "khoa truc"
            }

            Private ReadOnly FastenerKeywords As String() = {
                "bulong", "bu lông", "bu long", "ốc", "oc", "đai ốc", "dai oc", "vít", "vit", "ecu", "êcu", "then", "then chốt", "long đen", "long den",
                "long đen", "long den", "washer", "iso", "din", "jis", "m3", "m4", "m5", "m6", "m8", "lock collar", "locknut", "lock nut",
                "m10", "m12", "m16", "m20", "m24", "m30", "m36", "m42", "m48", "ss 2", "iso 4", "din 125", "din 127", "din 933", "din 934", "din 6912"
            }

            Private Function GetInventorApplication() As Inventor.Application
                Try
                    Return CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
                Catch ex As Exception
                    MessageBox.Show("Không lấy được Inventor đang chạy." & vbCrLf & vbCrLf & ex.Message,
                                    "BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Nothing
                End Try
            End Function

            Public Sub OnExecute(ByVal Context As NameValueMap)
                Dim invApp As Inventor.Application = Nothing
                Try
                    invApp = GetInventorApplication()
                    If invApp Is Nothing Then Exit Sub
                    Main(invApp)
                Catch ex As Exception
                    MessageBox.Show("Lỗi BOM:" & vbCrLf & vbCrLf & ex.Message,
                                    "BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Sub

            '=========================================================
            ' OPTIONS CLASS
            '=========================================================
            Private Class BomOptions
                Public IsAllLevel As Boolean = False
                Public SortAsm As Integer = 0
                Public SortPart As Integer = 0
                Public BaseText As String = ""
                Public Mode As Integer = 2
                Public ApplyPN As Boolean = False
                Public ApplySN As Boolean = False
                Public Cancelled As Boolean = True
            End Class


            '=========================================================
            ' FORM TÙY CHỌN — 1 FORM DUY NHẤT
            '=========================================================
            Private Function ShowBomOptionsForm() As BomOptions
                Dim opt As New BomOptions()

                Using frm As New Form()
                    frm.Text = "BOM Tool — Tùy chọn"
                    frm.AutoScaleMode = AutoScaleMode.None
                    frm.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                    frm.ClientSize = New Size(620, 660)
                    frm.FormBorderStyle = FormBorderStyle.FixedSingle
                    frm.StartPosition = FormStartPosition.CenterScreen
                    frm.MaximizeBox = False
                    frm.MinimizeBox = False
                    frm.ShowInTaskbar = False
                frm.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
                frm.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

                    '===== HEADER =====
                    Dim pnlHeader As New Panel()
                pnlHeader.Location = New System.Drawing.Point(0, 0)
                pnlHeader.Size = New System.Drawing.Size(620, 60)
                pnlHeader.BackColor = System.Drawing.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                    Dim lblTitle As New Label()
                    lblTitle.Text = "BOM TOOL — TÙY CHỌN"
                    lblTitle.Font = New Font("Segoe UI", 14.0F, FontStyle.Bold, GraphicsUnit.Point)
                lblTitle.ForeColor = System.Drawing.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Sắp xếp + đánh STT + cập nhật Part Number / Stock Number"
                lblSub.Font = New Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point)
                lblSub.ForeColor = System.Drawing.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 18
                lblSub.TextAlign = ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== GROUP 1: PHẠM VI =====
                Dim gb1 As New GroupBox() With {
                        .Text = "Phạm vi",
                        .Location = New System.Drawing.Point(15, 75),
                        .Size = New Size(590, 75),
                        .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                        .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                        .BackColor = System.Drawing.Color.White}
                frm.Controls.Add(gb1)

                Dim rdoTop As New RadioButton() With {
                        .Text = "Chỉ Top-level (cấp cao nhất)",
                        .Location = New System.Drawing.Point(20, 25),
                        .Size = New System.Drawing.Size(250, 25),
                        .Checked = True,
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                Dim rdoAll As New RadioButton() With {
                        .Text = "All-level (mỗi cấp đánh số riêng + kiểm tra PN)",
                        .Location = New System.Drawing.Point(280, 25),
                        .Size = New System.Drawing.Size(300, 25),
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb1.Controls.Add(rdoTop)
                gb1.Controls.Add(rdoAll)

                '===== GROUP 2: SẮP XẾP =====
                Dim gb2 As New GroupBox() With {
                        .Text = "Sắp xếp",
                        .Location = New System.Drawing.Point(15, 160),
                        .Size = New System.Drawing.Size(590, 105),
                        .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                        .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                        .BackColor = System.Drawing.Color.White}
                frm.Controls.Add(gb2)

                Dim lblAsm As New Label() With {
                        .Text = "Cụm lắp:",
                        .Location = New System.Drawing.Point(20, 30),
                        .Size = New System.Drawing.Size(90, 25),
                        .TextAlign = ContentAlignment.MiddleLeft,
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb2.Controls.Add(lblAsm)

                Dim cboAsm As New ComboBox() With {
                        .DropDownStyle = ComboBoxStyle.DropDownList,
                        .Location = New System.Drawing.Point(120, 30),
                        .Size = New System.Drawing.Size(450, 25)}
                cboAsm.Items.AddRange(New Object() {
                        "Khối lượng lớn → bé",
                        "Khối lượng bé → lớn",
                        "Tên ngắn → dài",
                        "Tên dài → ngắn",
                        "Chữ cái A → Z",
                        "Chữ cái Z → A"})
                cboAsm.SelectedIndex = 0
                gb2.Controls.Add(cboAsm)

                Dim lblPart As New Label() With {
                        .Text = "Part:",
                        .Location = New System.Drawing.Point(20, 65),
                        .Size = New Size(90, 25),
                        .TextAlign = ContentAlignment.MiddleLeft,
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb2.Controls.Add(lblPart)

                Dim cboPart As New ComboBox() With {
                        .DropDownStyle = ComboBoxStyle.DropDownList,
                        .Location = New System.Drawing.Point(120, 65),
                        .Size = New System.Drawing.Size(450, 25)}
                cboPart.Items.AddRange(New Object() {
                        "Khối lượng lớn → bé",
                        "Khối lượng bé → lớn",
                        "Tên ngắn → dài",
                        "Tên dài → ngắn",
                        "Chữ cái A → Z",
                        "Chữ cái Z → A"})
                cboPart.SelectedIndex = 0
                gb2.Controls.Add(cboPart)

                '===== GROUP 3: PART NUMBER / STOCK NUMBER =====
                Dim gb3 As New GroupBox() With {
                        .Text = "Part Number / Stock Number",
                        .Location = New System.Drawing.Point(15, 275),
                        .Size = New Size(590, 275),
                        .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                        .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                        .BackColor = System.Drawing.Color.White}
                frm.Controls.Add(gb3)

                '--- Chữ dùng để đặt tên ---
                Dim lblBase As New Label() With {
                        .Text = "Chữ dùng:",
                        .Location = New System.Drawing.Point(20, 30),
                        .Size = New System.Drawing.Size(90, 25),
                        .TextAlign = ContentAlignment.MiddleLeft,
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb3.Controls.Add(lblBase)

                Dim txtBase As New System.Windows.Forms.TextBox() With {
                        .Location = New System.Drawing.Point(120, 30),
                        .Size = New System.Drawing.Size(280, 25),
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb3.Controls.Add(txtBase)

                Dim lblHint As New Label() With {
                        .Text = "Để trống = chỉ sắp xếp, không sửa PN/SN",
                        .Location = New System.Drawing.Point(410, 50),
                        .Size = New System.Drawing.Size(170, 40),
                        .TextAlign = ContentAlignment.MiddleLeft,
                        .ForeColor = System.Drawing.Color.FromArgb(140, 140, 140),
                        .Font = New Font("Segoe UI", 8.0F, FontStyle.Italic, GraphicsUnit.Point)}
                gb3.Controls.Add(lblHint)

                '--- Cách ghi ---
                Dim lblMode As New Label() With {
                        .Text = "Cách ghi:",
                        .Location = New System.Drawing.Point(20, 80),
                        .Size = New System.Drawing.Size(90, 25),
                        .TextAlign = ContentAlignment.MiddleLeft,
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb3.Controls.Add(lblMode)

                Dim rdoModeReplace As New RadioButton() With {
                        .Text = "Thay toàn bộ (xóa cũ → chữ + STT)",
                        .Location = New System.Drawing.Point(120, 80),
                        .Size = New System.Drawing.Size(300, 25),
                        .Checked = True,
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb3.Controls.Add(rdoModeReplace)

                Dim rdoModeAfter As New RadioButton() With {
                        .Text = "Thêm chữ phía SAU",
                        .Location = New System.Drawing.Point(120, 110),
                        .Size = New System.Drawing.Size(300, 25),
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb3.Controls.Add(rdoModeAfter)

                Dim rdoModeBefore As New RadioButton() With {
                        .Text = "Thêm chữ phía TRƯỚC",
                        .Location = New System.Drawing.Point(120, 140),
                        .Size = New System.Drawing.Size(300, 25),
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb3.Controls.Add(rdoModeBefore)

                '--- Áp dụng cho ---
                Dim lblApply As New System.Windows.Forms.Label() With {
                        .Text = "Áp dụng cho:",
                        .Location = New System.Drawing.Point(20, 185),
                        .Size = New System.Drawing.Size(90, 40),
                        .TextAlign = ContentAlignment.MiddleLeft,
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb3.Controls.Add(lblApply)

                Dim rdoApplyBoth As New RadioButton() With {
                        .Text = "Cả Part Number + Stock Number",
                        .Location = New System.Drawing.Point(120, 185),
                        .Size = New Size(300, 25),
                        .Checked = True,
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb3.Controls.Add(rdoApplyBoth)

                Dim rdoApplyPN As New RadioButton() With {
                        .Text = "Chỉ Part Number",
                        .Location = New System.Drawing.Point(120, 215),
                        .Size = New Size(200, 25),
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb3.Controls.Add(rdoApplyPN)

                Dim rdoApplySN As New RadioButton() With {
                        .Text = "Chỉ Stock Number",
                        .Location = New System.Drawing.Point(320, 215),
                        .Size = New Size(200, 25),
                        .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
                gb3.Controls.Add(rdoApplySN)

                '--- Hàm bật/tắt nhóm PN/SN ---
                Dim updateEnabled As Action = Sub()
                                                  Dim hasBase As Boolean = (txtBase.Text.Trim() <> "")
                                                  rdoModeReplace.Enabled = hasBase
                                                  rdoModeAfter.Enabled = hasBase
                                                  rdoModeBefore.Enabled = hasBase
                                                  rdoApplyBoth.Enabled = hasBase
                                                  rdoApplyPN.Enabled = hasBase
                                                  rdoApplySN.Enabled = hasBase
                                              End Sub

                AddHandler txtBase.TextChanged, Sub() updateEnabled()
                updateEnabled()

                '===== NÚT =====
                Dim btnOK As New Button()
                btnOK.Text = "CHẠY"
                btnOK.Size = New Size(150, 40)
                btnOK.Location = New System.Drawing.Point(440, 565)
                btnOK.BackColor = System.Drawing.Color.FromArgb(45, 100, 180)
                btnOK.ForeColor = System.Drawing.Color.White
                btnOK.FlatStyle = FlatStyle.Flat
                    btnOK.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold, GraphicsUnit.Point)
                    btnOK.DialogResult = DialogResult.OK
                    frm.Controls.Add(btnOK)

                    Dim btnCancel As New Button()
                    btnCancel.Text = "HỦY"
                    btnCancel.Size = New Size(100, 40)
                btnCancel.Location = New System.Drawing.Point(290, 565)
                btnCancel.FlatStyle = FlatStyle.Flat
                    btnCancel.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
                    btnCancel.DialogResult = DialogResult.Cancel
                    frm.Controls.Add(btnCancel)
                    frm.AcceptButton = btnOK
                    frm.CancelButton = btnCancel

                    If frm.ShowDialog() <> DialogResult.OK Then
                        opt.Cancelled = True
                        Return opt
                    End If

                    '--- Lấy giá trị ---
                    opt.IsAllLevel = rdoAll.Checked
                    opt.SortAsm = cboAsm.SelectedIndex
                    opt.SortPart = cboPart.SelectedIndex
                    opt.BaseText = txtBase.Text.Trim()

                    If rdoModeReplace.Checked Then
                        opt.Mode = 2
                    ElseIf rdoModeAfter.Checked Then
                        opt.Mode = 3
                    Else
                        opt.Mode = 4
                    End If

                    If rdoApplyPN.Checked Then
                        opt.ApplyPN = True : opt.ApplySN = False
                    ElseIf rdoApplySN.Checked Then
                        opt.ApplyPN = False : opt.ApplySN = True
                    Else
                        opt.ApplyPN = True : opt.ApplySN = True
                    End If

                    If opt.BaseText = "" Then
                        opt.ApplyPN = False
                        opt.ApplySN = False
                    End If

                    opt.Cancelled = False
                End Using

                Return opt
            End Function


            Private Sub Main(ByVal invApp As Inventor.Application)
                Dim oAsm As AssemblyDocument = Nothing
                Try
                    oAsm = TryCast(invApp.ActiveDocument, AssemblyDocument)
                    If oAsm Is Nothing Then
                        MessageBox.Show("Vui lòng mở Assembly (.iam) trước khi chạy.",
                                        "BOM", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    End If

                    Dim oBOM As BOM = oAsm.ComponentDefinition.BOM
                    Try : oBOM.StructuredViewEnabled = True : Catch : End Try
                    Try : oBOM.StructuredViewFirstLevelOnly = False : Catch : End Try

                    Dim oBOMView As BOMView = Nothing
                    Try
                        oBOMView = oBOM.BOMViews.Item("Structured")
                    Catch
                        MessageBox.Show("Không tìm thấy Structured BOM.", "BOM",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try
                    If oBOMView Is Nothing Then Exit Sub

                    '===== FORM TÙY CHỌN =====
                    Dim opt As BomOptions = ShowBomOptionsForm()
                    If opt Is Nothing OrElse opt.Cancelled Then Exit Sub

                    Dim isAllLevel As Boolean = opt.IsAllLevel
                    Dim sortModeAsm As Integer = opt.SortAsm
                    Dim sortModePart As Integer = opt.SortPart
                    Dim baseText As String = opt.BaseText
                    Dim mode As Integer = opt.Mode
                    Dim applyPN As Boolean = opt.ApplyPN
                    Dim applySN As Boolean = opt.ApplySN

                    '===== XÁC NHẬN =====
                    Dim confirm As String =
                        "CHUẨN BỊ CHẠY" & vbCrLf &
                        "------------------" & vbCrLf &
                        If(isAllLevel, "All-level (mỗi cấp riêng + kiểm tra PN)", "Chỉ Top-level") & vbCrLf &
                        "Sắp xếp Cụm: " & SortName(sortModeAsm) & vbCrLf &
                        "Sắp xếp Part: " & SortName(sortModePart) & vbCrLf & vbCrLf

                    If baseText = "" Then
                        confirm &= "PN / SN: KHÔNG SỬA"
                    Else
                        confirm &= "Chữ: " & baseText & vbCrLf
                        Select Case mode
                            Case 2 : confirm &= "Cách: Thay toàn bộ + STT" & vbCrLf
                            Case 3 : confirm &= "Cách: Thêm phía sau" & vbCrLf
                            Case 4 : confirm &= "Cách: Thêm phía trước" & vbCrLf
                        End Select
                        If applyPN AndAlso applySN Then
                            confirm &= "Áp dụng: PN + SN"
                        ElseIf applyPN Then
                            confirm &= "Áp dụng: Chỉ Part Number"
                        Else
                            confirm &= "Áp dụng: Chỉ Stock Number"
                        End If
                    End If

                    confirm &= vbCrLf & vbCrLf & "Tiếp tục?"

                    If MessageBox.Show(confirm, "Xác nhận",
                                       MessageBoxButtons.OKCancel,
                                       MessageBoxIcon.Question) <> DialogResult.OK Then
                        Exit Sub
                    End If

                    Dim changedPN As Integer = 0
                    Dim changedSN As Integer = 0
                    Dim totalRows As Integer = 0
                    Dim listDocs As New List(Of Document)

                    If isAllLevel Then
                        ProcessLevelWithPNCheck(oBOMView.BOMRows, baseText, mode, applyPN, applySN,
                            changedPN, changedSN, totalRows, listDocs, oAsm,
                            sortModeAsm, sortModePart)
                    Else
                        Dim sortedRows As List(Of BOMRow) = SortRows(oBOMView.BOMRows, sortModeAsm, sortModePart)
                        Dim stt As Integer = 1
                        Dim pnToStt As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                        For Each row As BOMRow In sortedRows
                            System.Windows.Forms.Application.DoEvents()
                            If row Is Nothing Then Continue For

                            Dim refDoc As Document = Nothing
                            Try
                                If row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then Continue For
                                refDoc = row.ComponentDefinitions.Item(1).Document
                            Catch
                                Continue For
                            End Try
                            If refDoc Is Nothing Then Continue For

                            Try
                                If String.Equals(refDoc.FullFileName, oAsm.FullFileName, StringComparison.OrdinalIgnoreCase) Then
                                    Continue For
                                End If
                            Catch
                            End Try

                            Dim pn As String = GetProperty(refDoc, "Part Number")
                            If String.IsNullOrEmpty(pn) Then pn = refDoc.DisplayName

                            Dim thisStt As Integer
                            If pnToStt.ContainsKey(pn) Then
                                thisStt = pnToStt(pn)
                            Else
                                thisStt = stt
                                pnToStt(pn) = stt
                                stt += 1
                            End If

                            Try : row.ItemNumber = thisStt.ToString() : Catch : End Try

                            If baseText <> "" AndAlso mode >= 2 Then
                                Dim curPN As String = GetProperty(refDoc, "Part Number")
                                Dim curSN As String = GetProperty(refDoc, "Stock Number")
                                Dim newPN As String = BuildValue(curPN, baseText, thisStt, mode)
                                Dim newSN As String = BuildValue(curSN, baseText, thisStt, mode)

                                If applyPN AndAlso newPN <> "" Then
                                    If SetDesignProperty(refDoc, "Part Number", newPN) Then
                                        changedPN += 1
                                        If Not listDocs.Contains(refDoc) Then listDocs.Add(refDoc)
                                    End If
                                End If
                                If applySN AndAlso newSN <> "" Then
                                    If SetDesignProperty(refDoc, "Stock Number", newSN) Then
                                        changedSN += 1
                                        If Not listDocs.Contains(refDoc) Then listDocs.Add(refDoc)
                                    End If
                                End If
                            End If

                            totalRows += 1
                        Next
                    End If

                    For Each d As Document In listDocs
                        Try
                            If d.IsModifiable Then d.Update()
                        Catch
                        End Try
                    Next

                    Try : oBOM.Update() : Catch : End Try
                    Try : oAsm.Update2(True) : Catch : End Try

                    Dim msg As String =
                        "HOÀN TẤT" & vbCrLf &
                        "====================" & vbCrLf &
                        "Tổng dòng xử lý : " & totalRows.ToString() & vbCrLf & vbCrLf

                    If baseText = "" Then
                        msg &= "Part Number / Stock Number: KHÔNG SỬA"
                    Else
                        msg &= "Part Number đã ghi : " & changedPN.ToString() & vbCrLf &
                               "Stock Number đã ghi: " & changedSN.ToString()
                    End If

                    MessageBox.Show(msg, "BOM", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Lỗi:" & vbCrLf & vbCrLf & ex.Message,
                                    "BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Sub


            Private Function SortName(ByVal idx As Integer) As String
                Select Case idx
                    Case 0 : Return "Khối lượng lớn → bé"
                    Case 1 : Return "Khối lượng bé → lớn"
                    Case 2 : Return "Tên ngắn → dài"
                    Case 3 : Return "Tên dài → ngắn"
                    Case 4 : Return "Chữ cái A → Z"
                    Case 5 : Return "Chữ cái Z → A"
                End Select
                Return "?"
            End Function


            '=== GIỮ NGUYÊN CÁC HÀM XỬ LÝ BÊN DƯỚI ===
            ' ProcessLevelWithPNCheck, BuildValue, SetDesignProperty, GetProperty,
            ' GetSearchText, IsBearing, IsFastener, SortRows, ApplySort,
            ' GetMass, GetPartNumber

            ' [Dán lại các hàm này từ code cũ vào đây]


            ' All-level: mỗi cấp đánh số riêng + kiểm tra PN trùng trong cấp
            Private Sub ProcessLevelWithPNCheck(rows As BOMRowsEnumerator,
                                    baseText As String, mode As Integer,
                                    applyPN As Boolean, applySN As Boolean,
                                    ByRef changedPN As Integer, ByRef changedSN As Integer,
                                    ByRef totalRows As Integer,
                                    listDocs As List(Of Document),
                                    oAsm As AssemblyDocument,
                                    Optional sortModeAsm As Integer = 0,
                                    Optional sortModePart As Integer = 0)

            If rows Is Nothing Then Exit Sub

            Dim sortedRows As List(Of BOMRow) = SortRows(rows, sortModeAsm, sortModePart)
            If sortedRows Is Nothing OrElse sortedRows.Count = 0 Then Exit Sub

            Dim stt As Integer = 1
            Dim pnToStt As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

            For Each row As BOMRow In sortedRows
                System.Windows.Forms.Application.DoEvents()
                If row Is Nothing Then Continue For

                Dim refDoc As Document = Nothing
                Try
                    If row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then Continue For
                    refDoc = row.ComponentDefinitions.Item(1).Document
                Catch
                    Continue For
                End Try
                If refDoc Is Nothing Then Continue For

                Try
                    If String.Equals(refDoc.FullFileName, oAsm.FullFileName, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                Catch
                End Try

                Dim pn As String = GetProperty(refDoc, "Part Number")
                If String.IsNullOrEmpty(pn) Then pn = refDoc.DisplayName

                Dim thisStt As Integer
                If pnToStt.ContainsKey(pn) Then
                    thisStt = pnToStt(pn)          ' trùng PN → cùng STT
                Else
                    thisStt = stt
                    pnToStt(pn) = stt
                    stt += 1
                End If

                Try : row.ItemNumber = thisStt.ToString() : Catch : End Try

                If baseText <> "" AndAlso mode >= 2 Then
                    Dim curPN As String = GetProperty(refDoc, "Part Number")
                    Dim curSN As String = GetProperty(refDoc, "Stock Number")
                    Dim newPN As String = BuildValue(curPN, baseText, thisStt, mode)
                    Dim newSN As String = BuildValue(curSN, baseText, thisStt, mode)

                    If applyPN AndAlso newPN <> "" Then
                        If SetDesignProperty(refDoc, "Part Number", newPN) Then
                            changedPN += 1
                            If Not listDocs.Contains(refDoc) Then listDocs.Add(refDoc)
                        End If
                    End If
                    If applySN AndAlso newSN <> "" Then
                        If SetDesignProperty(refDoc, "Stock Number", newSN) Then
                            changedSN += 1
                            If Not listDocs.Contains(refDoc) Then listDocs.Add(refDoc)
                        End If
                    End If
                End If

                totalRows += 1

                ' Nhảy vào sub-assembly Structure (giống 1c) – chỉ để xử lý PN/STT bên trong
                If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Try
                        Dim subAsm As AssemblyDocument = CType(refDoc, AssemblyDocument)
                        Dim subBOM As BOM = subAsm.ComponentDefinition.BOM
                        Try : subBOM.StructuredViewEnabled = True : Catch : End Try
                        Try : subBOM.StructuredViewFirstLevelOnly = False : Catch : End Try
                        Try : subBOM.Update() : Catch : End Try

                        Dim subView As BOMView = Nothing
                        Try
                            subView = subBOM.BOMViews.Item("Structured")
                        Catch
                        End Try

                        If subView IsNot Nothing Then
                            '   ProcessLevelWithPNCheck(subView.BOMRows, baseText, mode, applyPN, applySN,
                            'changedPN, changedSN, totalRows, listDocs, oAsm)

                            '  ProcessLevelWithPNCheck(subView.BOMRows, baseText, mode, applyPN, applySN,
                            ' changedPN, changedSN, totalRows, listDocs, oAsm, sortMode)
                            ProcessLevelWithPNCheck(subView.BOMRows, baseText, mode, applyPN, applySN,
                            changedPN, changedSN, totalRows, listDocs, oAsm,
                            sortModeAsm, sortModePart)
                        End If
                    Catch
                    End Try
                End If
            Next
        End Sub

        Private Function BuildValue(current As String, baseText As String, stt As Integer, mode As Integer) As String
            If current Is Nothing Then current = ""
            current = current.Trim()
            Select Case mode
                Case 2 : Return baseText & stt.ToString()
                Case 3
                    If current = "" Then Return baseText
                    Return current & " " & baseText
                Case 4
                    If current = "" Then Return baseText
                    Return baseText & " " & current
            End Select
            Return ""
        End Function

        Private Function SetDesignProperty(doc As Document, propName As String, value As String) As Boolean
            Try
                If doc Is Nothing OrElse Not doc.IsModifiable Then Return False
                Dim designProps As PropertySet = Nothing
                Try
                    designProps = doc.PropertySets.Item("Design Tracking Properties")
                Catch
                    Return False
                End Try
                Try
                    Dim prop As Inventor.Property = designProps.Item(propName)
                    prop.Value = value
                    Return True
                Catch
                    Try
                        designProps.Add(value, propName)
                        Return True
                    Catch
                        Return False
                    End Try
                End Try
            Catch
                Return False
            End Try
        End Function

        Private Function GetProperty(doc As Document, propName As String) As String
            Try
                If doc Is Nothing Then Return ""
                Dim ps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim prop As Inventor.Property = ps.Item(propName)
                If prop Is Nothing OrElse prop.Value Is Nothing Then Return ""
                Return CStr(prop.Value).Trim()
            Catch
                Return ""
            End Try
        End Function

        Private Function GetSearchText(row As BOMRow) As String
            Try
                If row Is Nothing OrElse row.ComponentDefinitions Is Nothing OrElse
                   row.ComponentDefinitions.Count = 0 Then Return ""
                Dim doc As Document = row.ComponentDefinitions.Item(1).Document
                Dim pn As String = GetProperty(doc, "Part Number")
                If pn <> "" Then Return pn.Trim().ToLowerInvariant()
                Dim sn As String = GetProperty(doc, "Stock Number")
                If sn <> "" Then Return sn.Trim().ToLowerInvariant()
                Dim desc As String = GetProperty(doc, "Description")
                Return desc.Trim().ToLowerInvariant()
            Catch
                Return ""
            End Try
        End Function

        Private Function IsBearing(text As String) As Boolean
            If String.IsNullOrEmpty(text) Then Return False
            For Each kw As String In BearingKeywords
                If text.StartsWith(kw.ToLowerInvariant()) Then Return True
            Next
            Return False
        End Function

        Private Function IsFastener(text As String) As Boolean
            If String.IsNullOrEmpty(text) Then Return False
            For Each kw As String In FastenerKeywords
                If text.StartsWith(kw.ToLowerInvariant()) Then Return True
            Next
            Return False
        End Function
        Private Function SortRows(ByVal bomRows As BOMRowsEnumerator,
                          Optional ByVal sortModeAsm As Integer = 0,
                          Optional ByVal sortModePart As Integer = 0) As List(Of BOMRow)

            Dim normalAsm As New List(Of Tuple(Of BOMRow, Double, String))
            Dim purchasedAsm As New List(Of Tuple(Of BOMRow, Integer, String))
            Dim normalPart As New List(Of Tuple(Of BOMRow, Double, String))
            Dim purchasedPart As New List(Of Tuple(Of BOMRow, Integer, String))
            Dim phantomAsm As New List(Of Tuple(Of BOMRow, Double, String))
            Dim phantomPart As New List(Of Tuple(Of BOMRow, Double, String))
            Dim reference As New List(Of BOMRow)

            If bomRows Is Nothing Then Return New List(Of BOMRow)

            For Each row As BOMRow In bomRows
                If row Is Nothing Then Continue For

                Dim doc As Document = Nothing
                Try
                    If row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then
                        If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                            reference.Add(row)
                        End If
                        Continue For
                    End If
                    doc = row.ComponentDefinitions.Item(1).Document
                Catch
                    Continue For
                End Try
                If doc Is Nothing Then Continue For

                Dim isAsm As Boolean = False
                Dim isPart As Boolean = False
                Try
                    isAsm = (doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
                    isPart = (doc.DocumentType = DocumentTypeEnum.kPartDocumentObject)
                Catch
                End Try

                Dim searchText As String = GetSearchText(row)
                Dim isFast As Boolean = IsFastener(searchText)
                Dim isBear As Boolean = IsBearing(searchText)
                Dim pn As String = GetPartNumber(row)
                If String.IsNullOrEmpty(pn) Then pn = ""

                If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                    reference.Add(row)
                    Continue For
                End If

                If row.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                    Dim m As Double = GetMass(doc)
                    If isAsm Then
                        phantomAsm.Add(Tuple.Create(row, m, pn))
                    ElseIf isPart Then
                        phantomPart.Add(Tuple.Create(row, m, pn))
                    Else
                        reference.Add(row)
                    End If
                    Continue For
                End If

                If row.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                    Dim prio As Integer = 0
                    If isFast Then prio = 2
                    If isBear Then prio = 1
                    If isAsm Then
                        purchasedAsm.Add(Tuple.Create(row, prio, pn))
                    Else
                        purchasedPart.Add(Tuple.Create(row, prio, pn))
                    End If
                    Continue For
                End If

                Dim mass As Double = GetMass(doc)
                If isAsm Then
                    normalAsm.Add(Tuple.Create(row, mass, pn))
                ElseIf isPart Then
                    normalPart.Add(Tuple.Create(row, mass, pn))
                Else
                    reference.Add(row)
                End If
            Next

            '===== SẮP XẾP RIÊNG CỤM LẮP =====
            ApplySort(normalAsm, sortModeAsm)
            ApplySort(phantomAsm, sortModeAsm)

            '===== SẮP XẾP RIÊNG PART =====
            ApplySort(normalPart, sortModePart)
            ApplySort(phantomPart, sortModePart)

            ' Purchased giữ logic cũ (Bearing → Fastener → còn lại)
            purchasedAsm.Sort(Function(a, b)
                                  Dim c = a.Item2.CompareTo(b.Item2)
                                  If c <> 0 Then Return c
                                  Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                              End Function)
            purchasedPart.Sort(Function(a, b)
                                   Dim c = a.Item2.CompareTo(b.Item2)
                                   If c <> 0 Then Return c
                                   Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                               End Function)

            reference.Sort(Function(a, b) String.Compare(GetPartNumber(a), GetPartNumber(b), StringComparison.OrdinalIgnoreCase))

            Dim result As New List(Of BOMRow)
            For Each x In normalAsm : result.Add(x.Item1) : Next
            For Each x In purchasedAsm : result.Add(x.Item1) : Next
            For Each x In normalPart : result.Add(x.Item1) : Next
            For Each x In purchasedPart : result.Add(x.Item1) : Next
            For Each x In phantomAsm : result.Add(x.Item1) : Next
            For Each x In phantomPart : result.Add(x.Item1) : Next
            For Each x In reference : result.Add(x) : Next

            Return result
        End Function

        ' Hàm hỗ trợ sort theo mode
        Private Sub ApplySort(list As List(Of Tuple(Of BOMRow, Double, String)), mode As Integer)
            Select Case mode
                Case 1  ' Mass ASC (bé → lớn)
                    list.Sort(Function(a, b) a.Item2.CompareTo(b.Item2))

                Case 2  ' Tên ngắn → dài (theo độ dài, rồi A→Z)
                    list.Sort(Function(a, b)
                                  Dim c = a.Item3.Length.CompareTo(b.Item3.Length)
                                  If c <> 0 Then Return c
                                  Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                              End Function)

                Case 3  ' Tên dài → ngắn (theo độ dài, rồi Z→A)
                    list.Sort(Function(a, b)
                                  Dim c = b.Item3.Length.CompareTo(a.Item3.Length)
                                  If c <> 0 Then Return c
                                  Return String.Compare(b.Item3, a.Item3, StringComparison.OrdinalIgnoreCase)
                              End Function)

                Case 4  ' Chữ cái A → Z
                    list.Sort(Function(a, b) String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase))

                Case 5  ' Chữ cái Z → A
                    list.Sort(Function(a, b) String.Compare(b.Item3, a.Item3, StringComparison.OrdinalIgnoreCase))

                Case Else  ' 0 = Mass DESC (lớn → bé) - mặc định
                    list.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            End Select
        End Sub
        Private Function SortRows1(ByVal bomRows As BOMRowsEnumerator,
                          Optional ByVal sortMode As Integer = 0) As List(Of BOMRow)

            Dim normalAsm As New List(Of Tuple(Of BOMRow, Double, String))
            Dim purchasedAsm As New List(Of Tuple(Of BOMRow, Integer, String))
            Dim normalPart As New List(Of Tuple(Of BOMRow, Double, String))
            Dim purchasedPart As New List(Of Tuple(Of BOMRow, Integer, String))
            Dim phantomAsm As New List(Of Tuple(Of BOMRow, Double, String))
            Dim phantomPart As New List(Of Tuple(Of BOMRow, Double, String))
            Dim reference As New List(Of BOMRow)

            If bomRows Is Nothing Then Return New List(Of BOMRow)

            For Each row As BOMRow In bomRows
                If row Is Nothing Then Continue For

                Dim doc As Document = Nothing
                Try
                    If row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then
                        If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                            reference.Add(row)
                        End If
                        Continue For
                    End If
                    doc = row.ComponentDefinitions.Item(1).Document
                Catch
                    Continue For
                End Try
                If doc Is Nothing Then Continue For

                Dim isAsm As Boolean = False
                Dim isPart As Boolean = False
                Try
                    isAsm = (doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
                    isPart = (doc.DocumentType = DocumentTypeEnum.kPartDocumentObject)
                Catch
                End Try

                Dim searchText As String = GetSearchText(row)
                Dim isFast As Boolean = IsFastener(searchText)
                Dim isBear As Boolean = IsBearing(searchText)
                Dim pn As String = GetPartNumber(row)
                If String.IsNullOrEmpty(pn) Then pn = ""

                If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                    reference.Add(row)
                    Continue For
                End If

                If row.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                    Dim m As Double = GetMass(doc)
                    If isAsm Then
                        phantomAsm.Add(Tuple.Create(row, m, pn))
                    ElseIf isPart Then
                        phantomPart.Add(Tuple.Create(row, m, pn))
                    Else
                        reference.Add(row)
                    End If
                    Continue For
                End If

                If row.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                    Dim prio As Integer = 0
                    If isFast Then prio = 2
                    If isBear Then prio = 1
                    If isAsm Then
                        purchasedAsm.Add(Tuple.Create(row, prio, pn))
                    Else
                        purchasedPart.Add(Tuple.Create(row, prio, pn))
                    End If
                    Continue For
                End If

                Dim mass As Double = GetMass(doc)
                If isAsm Then
                    normalAsm.Add(Tuple.Create(row, mass, pn))
                ElseIf isPart Then
                    normalPart.Add(Tuple.Create(row, mass, pn))
                Else
                    reference.Add(row)
                End If
            Next

            '===== SẮP XẾP THEO sortMode =====
            ' 0 = Mass DESC (lớn → bé)
            ' 1 = Mass ASC  (bé → lớn)
            ' 2 = Name ASC  (ngắn → dài, rồi A→Z)
            ' 3 = Name DESC (dài → ngắn, rồi Z→A)

            Select Case sortMode
                Case 1  ' Mass ASC
                    normalAsm.Sort(Function(a, b) a.Item2.CompareTo(b.Item2))
                    normalPart.Sort(Function(a, b) a.Item2.CompareTo(b.Item2))
                    phantomAsm.Sort(Function(a, b) a.Item2.CompareTo(b.Item2))
                    phantomPart.Sort(Function(a, b) a.Item2.CompareTo(b.Item2))

                Case 2  ' Name short → long
                    normalAsm.Sort(Function(a, b)
                                       Dim c = a.Item3.Length.CompareTo(b.Item3.Length)
                                       If c <> 0 Then Return c
                                       Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                                   End Function)
                    normalPart.Sort(Function(a, b)
                                        Dim c = a.Item3.Length.CompareTo(b.Item3.Length)
                                        If c <> 0 Then Return c
                                        Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                                    End Function)
                    phantomAsm.Sort(Function(a, b)
                                        Dim c = a.Item3.Length.CompareTo(b.Item3.Length)
                                        If c <> 0 Then Return c
                                        Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                                    End Function)
                    phantomPart.Sort(Function(a, b)
                                         Dim c = a.Item3.Length.CompareTo(b.Item3.Length)
                                         If c <> 0 Then Return c
                                         Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                                     End Function)

                Case 3  ' Name long → short
                    normalAsm.Sort(Function(a, b)
                                       Dim c = b.Item3.Length.CompareTo(a.Item3.Length)
                                       If c <> 0 Then Return c
                                       Return String.Compare(b.Item3, a.Item3, StringComparison.OrdinalIgnoreCase)
                                   End Function)
                    normalPart.Sort(Function(a, b)
                                        Dim c = b.Item3.Length.CompareTo(a.Item3.Length)
                                        If c <> 0 Then Return c
                                        Return String.Compare(b.Item3, a.Item3, StringComparison.OrdinalIgnoreCase)
                                    End Function)
                    phantomAsm.Sort(Function(a, b)
                                        Dim c = b.Item3.Length.CompareTo(a.Item3.Length)
                                        If c <> 0 Then Return c
                                        Return String.Compare(b.Item3, a.Item3, StringComparison.OrdinalIgnoreCase)
                                    End Function)
                    phantomPart.Sort(Function(a, b)
                                         Dim c = b.Item3.Length.CompareTo(a.Item3.Length)
                                         If c <> 0 Then Return c
                                         Return String.Compare(b.Item3, a.Item3, StringComparison.OrdinalIgnoreCase)
                                     End Function)

                Case Else  ' 0 = Mass DESC (mặc định cũ)
                    normalAsm.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
                    normalPart.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
                    phantomAsm.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
                    phantomPart.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            End Select

            ' Purchased vẫn ưu tiên Bearing → Fastener → còn lại (giữ nguyên logic cũ)
            purchasedAsm.Sort(Function(a, b)
                                  Dim c = a.Item2.CompareTo(b.Item2)
                                  If c <> 0 Then Return c
                                  Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                              End Function)
            purchasedPart.Sort(Function(a, b)
                                   Dim c = a.Item2.CompareTo(b.Item2)
                                   If c <> 0 Then Return c
                                   Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                               End Function)

            reference.Sort(Function(a, b) String.Compare(GetPartNumber(a), GetPartNumber(b), StringComparison.OrdinalIgnoreCase))

            Dim result As New List(Of BOMRow)
            For Each x In normalAsm : result.Add(x.Item1) : Next
            For Each x In purchasedAsm : result.Add(x.Item1) : Next
            For Each x In normalPart : result.Add(x.Item1) : Next
            For Each x In purchasedPart : result.Add(x.Item1) : Next
            For Each x In phantomAsm : result.Add(x.Item1) : Next
            For Each x In phantomPart : result.Add(x.Item1) : Next
            For Each x In reference : result.Add(x) : Next

            Return result
        End Function

        Private Function GetMass(ByVal doc As Document) As Double
            Try
                If doc Is Nothing Then Return 0
                If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Return CType(doc, AssemblyDocument).ComponentDefinition.MassProperties.Mass
                ElseIf doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Return CType(doc, PartDocument).ComponentDefinition.MassProperties.Mass
                End If
            Catch
            End Try
            Return 0
        End Function

        Private Function GetPartNumber(ByVal row As BOMRow) As String
            Try
                If row Is Nothing OrElse row.ComponentDefinitions Is Nothing OrElse
                   row.ComponentDefinitions.Count = 0 Then Return ""
                Return GetProperty(row.ComponentDefinitions.Item(1).Document, "Part Number")
            Catch
                Return ""
            End Try
        End Function

        Private Function PickFromList(ByVal title As String,
                                      ByVal items As String(),
                                      Optional ByVal defaultIndex As Integer = 0) As Integer
            Dim frm As New Form()
            Try
                frm.Text = title
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.FormBorderStyle = FormBorderStyle.FixedDialog
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.Width = 460
                frm.Height = 310

                Dim lst As New ListBox()
                lst.Left = 12 : lst.Top = 12
                lst.Width = 420 : lst.Height = 200
                lst.Font = New System.Drawing.Font("Segoe UI", 10)

                For Each s As String In items
                    lst.Items.Add(s)
                Next

                If lst.Items.Count > 0 Then
                    If defaultIndex >= 0 AndAlso defaultIndex < lst.Items.Count Then
                        lst.SelectedIndex = defaultIndex
                    Else
                        lst.SelectedIndex = 0
                    End If
                End If

                Dim btnOK As New Button()
                btnOK.Text = "OK"
                btnOK.Left = 250 : btnOK.Top = 225
                btnOK.Width = 85 : btnOK.Height = 30
                btnOK.DialogResult = DialogResult.OK

                Dim btnCancel As New Button()
                btnCancel.Text = "Hủy"
                btnCancel.Left = 345 : btnCancel.Top = 225
                btnCancel.Width = 85 : btnCancel.Height = 30
                btnCancel.DialogResult = DialogResult.Cancel

                frm.Controls.Add(lst)
                frm.Controls.Add(btnOK)
                frm.Controls.Add(btnCancel)
                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel
                frm.KeyPreview = True

                AddHandler lst.DoubleClick,
                    Sub(s, e)
                        frm.DialogResult = DialogResult.OK
                        frm.Close()
                    End Sub

                AddHandler frm.KeyDown,
                    Sub(s, e)
                        If e.KeyCode = Keys.Escape Then
                            e.Handled = True
                            frm.DialogResult = DialogResult.Cancel
                            frm.Close()
                        End If
                    End Sub

                If frm.ShowDialog() <> DialogResult.OK Then Return -1
                If lst.SelectedIndex < 0 Then Return -1
                Return lst.SelectedIndex
            Finally
                If frm IsNot Nothing Then
                    Try : frm.Dispose() : Catch : End Try
                End If
            End Try
        End Function

    End Module

End Namespace



