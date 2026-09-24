Imports Inventor
Imports System.Windows.Forms
Imports Drw = System.Drawing
Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.IO

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode

    Public Module Ass_Bom_2

        Private Class BomItemOptions
            Public Prefix As String = ""
            Public Cancelled As Boolean = True
        End Class


        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                NumberAssemblyItem1()
            Catch ex As Exception
                MessageBox.Show(
                    "LỖI NumberAssemblyItem1:" & vbCrLf & vbCrLf &
                    ex.Message & vbCrLf & vbCrLf &
                    "Stack:" & vbCrLf & ex.StackTrace,
                    "Assembly2 - ERROR",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End Try
        End Sub


        '=========================================================
        ' FORM TÙY CHỌN
        '=========================================================
        Private Function ShowBomItemForm() As BomItemOptions
            Dim opt As New BomItemOptions()

            Using frm As New Form()
                frm.Text = "Đánh STT BOM — Tùy chọn"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(540, 560)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.BackColor = Drw.Color.FromArgb(245, 245, 245)
                frm.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

                '===== HEADER =====
                Dim pnlHeader As New Panel()
                pnlHeader.Location = New Drw.Point(0, 0)
                pnlHeader.Size = New Drw.Size(540, 70)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "ĐÁNH STT BOM — ITEM1"
                lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Sắp xếp BOM + ghi STT vào User Defined Properties"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== GROUP: PREFIX =====
                Dim gb As New GroupBox() With {
                    .Text = "Prefix STT",
                    .Location = New Drw.Point(15, 85),
                    .Size = New Drw.Size(510, 130),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb)

                Dim lblPrefix As New Label() With {
                    .Text = "Chữ thêm trước số:",
                    .Location = New Drw.Point(20, 35),
                    .Size = New Drw.Size(160, 25),
                    .TextAlign = Drw.ContentAlignment.MiddleLeft,
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb.Controls.Add(lblPrefix)

                Dim txtPrefix As New System.Windows.Forms.TextBox() With {
                    .Location = New Drw.Point(190, 35),
                    .Size = New Drw.Size(290, 25),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb.Controls.Add(txtPrefix)

                Dim lblExample As New Label() With {
                    .Text = "Ví dụ: TH → TH1, TH2, TH3...",
                    .Location = New Drw.Point(190, 65),
                    .Size = New Drw.Size(290, 20),
                    .ForeColor = Drw.Color.FromArgb(140, 140, 140),
                    .Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)}
                gb.Controls.Add(lblExample)

                Dim lblHint As New Label() With {
                    .Text = "Để trống = chỉ ghi số (1, 2, 3...)",
                    .Location = New Drw.Point(190, 90),
                    .Size = New Drw.Size(290, 20),
                    .ForeColor = Drw.Color.FromArgb(140, 140, 140),
                    .Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)}
                gb.Controls.Add(lblHint)

                '===== GROUP: GHI CHÚ — TO HƠN =====
                Dim gb2 As New GroupBox() With {
                    .Text = "Lưu ý",
                    .Location = New Drw.Point(15, 230),
                    .Size = New Drw.Size(510, 250),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb2)

                Dim lblNote As New Label() With {
                    .Text = "• Sắp xếp BOM theo thứ tự ưu tiên:" & vbCrLf &
                            "    1. Cụm lắp (theo khối lượng giảm dần)" & vbCrLf &
                            "    2. Part thường (theo khối lượng giảm dần)" & vbCrLf &
                            "    3. Purchased (vật tư mua ngoài)" & vbCrLf &
                            "    4. Phantom" & vbCrLf &
                            "    5. Reference (tham chiếu)" & vbCrLf &
                            "" & vbCrLf &
                            "• STT ghi vào User Defined Property tên 'item1'" & vbCrLf &
                            "• Part trùng Part Number + Revision → cùng STT" & vbCrLf &
                            "• Có thể chạy lại nhiều lần — số sẽ cập nhật theo BOM",
                    .Location = New Drw.Point(20, 30),
                    .Size = New Drw.Size(470, 200),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(60, 60, 60)}
                gb2.Controls.Add(lblNote)

                '===== NÚT =====
                Dim btnOK As New Button()
                btnOK.Text = "CHẠY"
                btnOK.Size = New Drw.Size(150, 44)
                btnOK.Location = New Drw.Point(370, 495)
                btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
                btnOK.ForeColor = Drw.Color.White
                btnOK.FlatStyle = FlatStyle.Flat
                btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                btnOK.DialogResult = DialogResult.OK
                frm.Controls.Add(btnOK)

                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(110, 44)
                btnCancel.Location = New Drw.Point(250, 495)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                btnCancel.DialogResult = DialogResult.Cancel
                frm.Controls.Add(btnCancel)
                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel

                If frm.ShowDialog() <> DialogResult.OK Then
                    opt.Cancelled = True
                    Return opt
                End If

                opt.Prefix = txtPrefix.Text.Trim()
                opt.Cancelled = False
            End Using

            Return opt
        End Function


        '==========================================================
        ' MAIN — GIỮ NGUYÊN
        '==========================================================
        Public Sub NumberAssemblyItem1()

            Dim invApp As Inventor.Application = g_inventorApplication

            If invApp Is Nothing Then
                MessageBox.Show("Không lấy được Inventor Application.", "Assembly2")
                Return
            End If

            Dim oAsm As AssemblyDocument =
                TryCast(invApp.ActiveDocument, AssemblyDocument)

            If oAsm Is Nothing Then
                MessageBox.Show(
                    "Active document không phải Assembly.",
                    "Assembly2 - item1",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
                Return
            End If

            Dim oDef As AssemblyComponentDefinition = oAsm.ComponentDefinition
            Dim oBOM As BOM = oDef.BOM

            oBOM.StructuredViewEnabled = True
            oBOM.StructuredViewFirstLevelOnly = False

            Dim oBOMView As BOMView = Nothing
            Try
                oBOMView = oBOM.BOMViews.Item("Structured")
            Catch ex As Exception
                MessageBox.Show(
                    "Không lấy được Structured BOM." & vbCrLf & ex.Message,
                    "Assembly2 - BOM ERROR",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
                Return
            End Try

            Dim opt As BomItemOptions = ShowBomItemForm()
            If opt Is Nothing OrElse opt.Cancelled Then Return

            Dim prefix As String = opt.Prefix

            Dim allDocs As New HashSet(Of Document)

            For Each occ As ComponentOccurrence In oDef.Occurrences
                CollectAllDocs(occ, allDocs)
            Next

            allDocs.Remove(oAsm)

            Dim topRows As List(Of BOMRow) = Nothing

            Try
                topRows = SortRows(oBOMView.BOMRows)
            Catch ex As Exception
                MessageBox.Show(
                    "Lỗi trong Function SortRows:" & vbCrLf & vbCrLf &
                    ex.Message & vbCrLf & vbCrLf &
                    ex.StackTrace,
                    "Assembly2 - SortRows ERROR",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
                Return
            End Try

            If topRows Is Nothing Then
                MessageBox.Show("SortRows trả về Nothing.", "Assembly2")
                Return
            End If

            Dim partKeyToDocs As New Dictionary(Of String, List(Of Document))

            For Each refDoc As Document In allDocs
                If refDoc Is Nothing Then Continue For

                Dim partNum As String = GetPartNumberFromDoc(refDoc).Trim().ToUpper()
                If partNum = "" Then partNum = GetFallbackName(refDoc)

                Dim rev As String = GetRevisionFromDoc(refDoc).Trim().ToUpper()
                If rev = "" Then rev = "?"

                Dim partKey As String = partNum & "|" & rev

                If Not partKeyToDocs.ContainsKey(partKey) Then
                    partKeyToDocs.Add(partKey, New List(Of Document))
                End If

                If Not partKeyToDocs(partKey).Contains(refDoc) Then
                    partKeyToDocs(partKey).Add(refDoc)
                End If
            Next

            Dim partKeyToSTT As New Dictionary(Of String, String)
            Dim sttCounter As Integer = 1

            For Each row As BOMRow In topRows
                If row Is Nothing Then Continue For
                If row.ComponentDefinitions Is Nothing Then Continue For
                If row.ComponentDefinitions.Count = 0 Then Continue For

                Dim refDoc As Document = Nothing
                Try
                    refDoc = row.ComponentDefinitions.Item(1).Document
                Catch
                    Continue For
                End Try

                If refDoc Is Nothing Then Continue For

                If String.Compare(refDoc.FullFileName, oAsm.FullFileName, True) = 0 Then
                    Continue For
                End If

                Dim partNum As String = GetPartNumberFromDoc(refDoc).Trim().ToUpper()
                If partNum = "" Then partNum = GetFallbackName(refDoc)

                Dim rev As String = GetRevisionFromDoc(refDoc).Trim().ToUpper()
                If rev = "" Then rev = "?"

                Dim partKey As String = partNum & "|" & rev

                Dim numericSTT As String
                If partKeyToSTT.ContainsKey(partKey) Then
                    numericSTT = partKeyToSTT(partKey)
                Else
                    numericSTT = CStr(sttCounter)
                    partKeyToSTT.Add(partKey, numericSTT)
                    sttCounter += 1
                End If

                Try
                    row.ItemNumber = numericSTT
                Catch
                End Try
            Next

            Dim updateOK As Integer = 0
            Dim updateFail As Integer = 0

            For Each kvp As KeyValuePair(Of String, String) In partKeyToSTT

                Dim key As String = kvp.Key
                Dim numericSTT As String = kvp.Value
                Dim fullSTT As String = numericSTT

                If prefix <> "" Then
                    fullSTT = prefix & fullSTT
                End If

                If partKeyToDocs.ContainsKey(key) Then
                    For Each doc As Document In partKeyToDocs(key)
                        If AddOrUpdateSTT(doc, fullSTT) Then
                            updateOK += 1
                        Else
                            updateFail += 1
                        End If
                    Next
                End If
            Next

            Try
                oAsm.Update2(True)
            Catch
            End Try

            MessageBox.Show(
                "HOÀN TẤT" & vbCrLf & vbCrLf &
                "Số nhóm STT: " & partKeyToSTT.Count & vbCrLf &
                "Số Document cập nhật: " & updateOK & vbCrLf &
                "Số Document không cập nhật: " & updateFail & vbCrLf &
                "Prefix item1: '" & prefix & "'",
                "Assembly2 - item1",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)

        End Sub


        '==========================================================
        ' CÁC HÀM BÊN DƯỚI GIỮ NGUYÊN
        '==========================================================
        Public Sub CollectAllDocs(occ As ComponentOccurrence, ByRef docsSet As HashSet(Of Document))
            If occ Is Nothing Then Return
            Try
                Dim doc As Document = occ.Definition.Document
                If doc IsNot Nothing Then docsSet.Add(doc)

                If occ.Definition.Type = ObjectTypeEnum.kAssemblyComponentDefinitionObject Then
                    Dim subDef As AssemblyComponentDefinition = TryCast(occ.Definition, AssemblyComponentDefinition)
                    If subDef IsNot Nothing Then
                        For Each subOcc As ComponentOccurrence In subDef.Occurrences
                            CollectAllDocs(subOcc, docsSet)
                        Next
                    End If
                End If
            Catch
            End Try
        End Sub


        Public Function SortRows(bomRows As BOMRowsEnumerator) As List(Of BOMRow)
            Dim subAsmMassList As New List(Of Tuple(Of BOMRow, Double))
            Dim partMassList As New List(Of Tuple(Of BOMRow, Double))
            Dim purchasedAsm As New List(Of BOMRow)
            Dim purchasedPart As New List(Of BOMRow)
            Dim phantomAsmMassList As New List(Of Tuple(Of BOMRow, Double))
            Dim phantomPartMassList As New List(Of Tuple(Of BOMRow, Double))
            Dim reference As New List(Of BOMRow)

            For Each row As BOMRow In bomRows
                If row Is Nothing Then Continue For
                If row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then
                    If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                        reference.Add(row)
                    End If
                    Continue For
                End If

                Dim refDoc As Document = Nothing
                Try
                    refDoc = row.ComponentDefinitions.Item(1).Document
                Catch
                    Continue For
                End Try
                If refDoc Is Nothing Then Continue For

                If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                    reference.Add(row)

                ElseIf row.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                    If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Dim mass As Double = 0
                        Try : mass = refDoc.ComponentDefinition.MassProperties.Mass : Catch : End Try
                        phantomAsmMassList.Add(Tuple.Create(row, mass))
                    ElseIf refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        Dim mass As Double = 0
                        Try : mass = refDoc.ComponentDefinition.MassProperties.Mass : Catch : End Try
                        phantomPartMassList.Add(Tuple.Create(row, mass))
                    End If

                ElseIf row.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                    If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        purchasedAsm.Add(row)
                    ElseIf refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        purchasedPart.Add(row)
                    End If

                Else
                    If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Dim mass As Double = 0
                        Try : mass = refDoc.ComponentDefinition.MassProperties.Mass : Catch : End Try
                        subAsmMassList.Add(Tuple.Create(row, mass))
                    ElseIf refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        Dim mass As Double = 0
                        Try : mass = refDoc.ComponentDefinition.MassProperties.Mass : Catch : End Try
                        partMassList.Add(Tuple.Create(row, mass))
                    End If
                End If
            Next

            subAsmMassList.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            partMassList.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            phantomAsmMassList.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            phantomPartMassList.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            purchasedAsm.Sort(Function(a, b) String.Compare(GetPartNumber(a), GetPartNumber(b), True))
            purchasedPart.Sort(Function(a, b) String.Compare(GetPartNumber(a), GetPartNumber(b), True))
            reference.Sort(Function(a, b) String.Compare(GetPartNumber(a), GetPartNumber(b), True))

            Dim orderedRows As New List(Of BOMRow)
            For Each item In subAsmMassList : orderedRows.Add(item.Item1) : Next
            For Each item In partMassList : orderedRows.Add(item.Item1) : Next
            For Each item In purchasedAsm : orderedRows.Add(item) : Next
            For Each item In purchasedPart : orderedRows.Add(item) : Next
            For Each item In phantomAsmMassList : orderedRows.Add(item.Item1) : Next
            For Each item In phantomPartMassList : orderedRows.Add(item.Item1) : Next
            orderedRows.AddRange(reference)

            Return orderedRows
        End Function


        Public Function GetPartNumber(row As BOMRow) As String
            Try
                If row Is Nothing Then Return ""
                If row.ComponentDefinitions Is Nothing Then Return ""
                If row.ComponentDefinitions.Count = 0 Then Return ""

                Dim refDoc As Document = Nothing
                Try
                    refDoc = row.ComponentDefinitions.Item(1).Document
                Catch
                    Return ""
                End Try
                If refDoc Is Nothing Then Return ""

                Return GetPartNumberFromDoc(refDoc)
            Catch
                Return ""
            End Try
        End Function


        Public Function GetPartNumberFromDoc(doc As Document) As String
            Try
                If doc Is Nothing Then Return ""
                Return CStr(doc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value)
            Catch
                Return ""
            End Try
        End Function


        Public Function GetRevisionFromDoc(doc As Document) As String
            Try
                If doc Is Nothing Then Return ""
                Return CStr(doc.PropertySets.Item("Design Tracking Properties").Item("Revision Number").Value)
            Catch
                Return ""
            End Try
        End Function


        Public Function GetFallbackName(doc As Document) As String
            Try
                If doc Is Nothing Then Return "?"
                If doc.DisplayName <> "" Then Return doc.DisplayName.Trim().ToUpper()
                Return "?"
            Catch
                Return "?"
            End Try
        End Function


        Public Function AddOrUpdateSTT(doc As Document, value As String) As Boolean
            Try
                If doc Is Nothing Then Return False
                If Not doc.IsModifiable Then Return False

                Dim ps As PropertySet = Nothing
                Try
                    ps = doc.PropertySets.Item("Inventor User Defined Properties")
                Catch
                    Return False
                End Try
                If ps Is Nothing Then Return False

                Dim prop As Inventor.Property = Nothing
                For Each p As Inventor.Property In ps
                    If String.Equals(p.Name, "item1", StringComparison.OrdinalIgnoreCase) Then
                        prop = p
                        Exit For
                    End If
                Next

                If prop Is Nothing Then
                    prop = ps.Add(value, "item1")
                Else
                    prop.Value = value
                End If

                Try : doc.Update() : Catch : End Try
                Try
                    If Not doc.ReadOnly Then doc.Save()
                Catch
                End Try

                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

    End Module

End Namespace