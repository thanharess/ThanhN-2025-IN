Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Lenhngoaicumlap
    Public Module Design_Assistant

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim invApp As Inventor.Application = g_inventorApplication
            Try
                If invApp.ActiveDocument Is Nothing Then
                    MessageBox.Show("Không có Document đang mở.", "Copy Assembly")
                    Exit Sub
                End If
                If invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Vui lòng mở Assembly trước!", "Copy Assembly")
                    Exit Sub
                End If

                Dim asmDoc As AssemblyDocument = CType(invApp.ActiveDocument, AssemblyDocument)
                Dim sourceRoot As String = IO.Path.GetDirectoryName(asmDoc.FullFileName)
                Dim oldMainName As String = IO.Path.GetFileNameWithoutExtension(asmDoc.FullFileName)

                Dim opt As CopyAsmOptions = Nothing
                Dim frm As New CopyAsmForm(oldMainName, sourceRoot)
                Try
                    frm.ShowDialog()
                    If frm.DialogResult <> DialogResult.OK OrElse Not frm.UserConfirmed Then
                        Exit Sub
                    End If
                    opt = frm.Options
                Finally
                    Try : frm.Dispose() : Catch : End Try
                End Try

                If opt Is Nothing Then Exit Sub
                If String.IsNullOrWhiteSpace(opt.NewMainName) Then Exit Sub
                If String.IsNullOrWhiteSpace(opt.TargetFolder) Then Exit Sub

                Dim newMainName As String = opt.NewMainName
                Dim subPrefix As String = opt.Prefix
                Dim subSuffix As String = opt.Suffix
                Dim targetFolder As String = opt.TargetFolder

                Dim fileMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                Dim processed As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim mainFolder As String = IO.Path.Combine(targetFolder, subPrefix & newMainName & subSuffix)

                CollectAssemblyTree(asmDoc, mainFolder, newMainName, subPrefix, subSuffix, fileMap, processed, True)

                For Each kvp As KeyValuePair(Of String, String) In fileMap
                    Try
                        Dim destinationFolder As String = IO.Path.GetDirectoryName(kvp.Value)
                        If Not IO.Directory.Exists(destinationFolder) Then
                            IO.Directory.CreateDirectory(destinationFolder)
                        End If
                        IO.File.Copy(kvp.Key, kvp.Value, True)
                    Catch ex As Exception
                        MessageBox.Show("Không copy được file:" & vbCrLf & vbCrLf & kvp.Key & vbCrLf & vbCrLf & ex.Message, "Lỗi Copy")
                    End Try
                Next

                If Not fileMap.ContainsKey(asmDoc.FullFileName) Then
                    Throw New Exception("Không tìm thấy Main Assembly trong File Map.")
                End If

                Dim newAsmPath As String = fileMap(asmDoc.FullFileName)
                Dim newAsm As AssemblyDocument = CType(invApp.Documents.Open(newAsmPath, True), AssemblyDocument)

                ReplaceAllReferences(newAsm, fileMap)

                Dim propertyProcessed As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                UpdateAlliProperties(newAsm, propertyProcessed)

                newAsm.Update2(True)
                newAsm.Save2(True)

                MessageBox.Show("HOÀN TẤT!" & vbCrLf & vbCrLf &
                                "Main Assembly:" & vbCrLf & newAsmPath & vbCrLf & vbCrLf &
                                "Cấu trúc:" & vbCrLf &
                                "• Assembly → Folder riêng" & vbCrLf &
                                "• Part → Không tạo Folder" & vbCrLf &
                                "• Part nằm trong Folder Assembly cha",
                                "Copy Assembly")

            Catch ex As Exception
                MessageBox.Show("CÓ LỖI:" & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, "Copy Assembly")
            End Try
        End Sub

        Public Class CopyAsmOptions
            Public NewMainName As String = ""
            Public Prefix As String = ""
            Public Suffix As String = ""
            Public TargetFolder As String = ""
        End Class

        Public Class CopyAsmForm
            Inherits Form

            Private _userConfirmed As Boolean = False
            Public ReadOnly Property UserConfirmed As Boolean
                Get
                    Return _userConfirmed
                End Get
            End Property
            Public Property Options As CopyAsmOptions

            Private txtMainName As System.Windows.Forms.TextBox
            Private txtPrefix As System.Windows.Forms.TextBox
            Private txtSuffix As System.Windows.Forms.TextBox
            Private txtFolder As System.Windows.Forms.TextBox
            Private btnBrowse As System.Windows.Forms.Button

            Private Const FORM_W As Integer = 560
            Private Const FORM_H As Integer = 400

            Public Sub New(ByVal oldMainName As String, ByVal defaultFolder As String)
                Me.Text = "Copy Assembly — Tùy chọn"
                Me.AutoScaleMode = AutoScaleMode.None
                Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                Me.FormBorderStyle = FormBorderStyle.FixedSingle
                Me.StartPosition = FormStartPosition.CenterScreen
                Me.MaximizeBox = False
                Me.MinimizeBox = False
                Me.ShowInTaskbar = False
                Me.ClientSize = New Size(FORM_W, FORM_H)
                Me.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
                Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

                Dim pnlHeader As New Panel()
                pnlHeader.Location = New System.Drawing.Point(0, 0)
                pnlHeader.Size = New Size(FORM_W, 60)
                pnlHeader.BackColor = System.Drawing.Color.FromArgb(45, 100, 180)
                Me.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "COPY ASSEMBLY"
                lblTitle.Font = New Font("Segoe UI", 14.0F, FontStyle.Bold, GraphicsUnit.Point)
                lblTitle.ForeColor = System.Drawing.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Tạo bản sao Assembly với cấu trúc thư mục mới"
                lblSub.Font = New Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point)
                lblSub.ForeColor = System.Drawing.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                Dim gb1 As New GroupBox() With {
                    .Text = "Tên file",
                    .Location = New System.Drawing.Point(15, 75),
                    .Size = New System.Drawing.Size(FORM_W - 30, 180),
                    .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                    .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                    .BackColor = System.Drawing.Color.White}
                Me.Controls.Add(gb1)

                Dim t1 As New TableLayoutPanel() With {
                    .Location = New System.Drawing.Point(20, 25),
                    .Size = New System.Drawing.Size(FORM_W - 70, 145),
                    .ColumnCount = 2,
                    .RowCount = 4,
                    .BackColor = System.Drawing.Color.White}
                t1.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 190))
                t1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
                For i As Integer = 0 To 3
                    t1.RowStyles.Add(New RowStyle(SizeType.Absolute, 34))
                Next
                gb1.Controls.Add(t1)

                t1.Controls.Add(MakeFieldLabel("Tên gốc:"), 0, 0)
                Dim lblOldVal As New Label() With {
                    .Text = oldMainName,
                    .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                    .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                    .Dock = DockStyle.Fill,
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .AutoEllipsis = True}
                t1.Controls.Add(lblOldVal, 1, 0)

                t1.Controls.Add(MakeFieldLabel("Tên MỚI (*):"), 0, 1)
                txtMainName = New System.Windows.Forms.TextBox() With {
                    .Text = oldMainName & "-2",
                    .Dock = DockStyle.Fill,
                    .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                    .Margin = New Padding(0, 4, 0, 4)}
                t1.Controls.Add(txtMainName, 1, 1)

                t1.Controls.Add(MakeFieldLabel("Prefix (đầu):"), 0, 2)
                txtPrefix = New System.Windows.Forms.TextBox() With {
                    .Text = "",
                    .Dock = DockStyle.Fill,
                    .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                    .Margin = New Padding(0, 4, 0, 4)}
                t1.Controls.Add(txtPrefix, 1, 2)

                t1.Controls.Add(MakeFieldLabel("Suffix (cuối):"), 0, 3)
                txtSuffix = New System.Windows.Forms.TextBox() With {
                    .Text = "-",
                    .Dock = DockStyle.Fill,
                    .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                    .Margin = New Padding(0, 4, 0, 4)}
                t1.Controls.Add(txtSuffix, 1, 3)

                Dim gb2 As New GroupBox() With {
                    .Text = "Thư mục đích",
                    .Location = New System.Drawing.Point(15, 265),
                    .Size = New Size(FORM_W - 30, 75),
                    .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                    .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                    .BackColor = System.Drawing.Color.White}
                Me.Controls.Add(gb2)

                txtFolder = New System.Windows.Forms.TextBox() With {
                    .Text = defaultFolder,
                    .Location = New System.Drawing.Point(20, 30),
                    .Size = New System.Drawing.Size(400, 25),
                    .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                    .ReadOnly = True,
                    .BackColor = System.Drawing.Color.FromArgb(250, 250, 250)}
                gb2.Controls.Add(txtFolder)

                btnBrowse = New System.Windows.Forms.Button() With {
                    .Text = "...",
                    .Location = New System.Drawing.Point(425, 29),
                    .Size = New Size(60, 27),
                    .FlatStyle = FlatStyle.Flat,
                    .BackColor = System.Drawing.Color.FromArgb(240, 240, 240)}
                AddHandler btnBrowse.Click, AddressOf OnBrowseFolder
                gb2.Controls.Add(btnBrowse)

                Dim btnOK As New Button()
                btnOK.Text = "TẠO BẢN SAO"
                btnOK.Size = New System.Drawing.Size(150, 40)
                btnOK.Location = New System.Drawing.Point(FORM_W - 285, FORM_H - 52)
                btnOK.BackColor = System.Drawing.Color.FromArgb(45, 100, 180)
                btnOK.ForeColor = System.Drawing.Color.White
                btnOK.FlatStyle = FlatStyle.Flat
                btnOK.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold, GraphicsUnit.Point)
                AddHandler btnOK.Click, AddressOf OnOKClick
                Me.Controls.Add(btnOK)

                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New System.Drawing.Size(100, 40)
                btnCancel.Location = New System.Drawing.Point(FORM_W - 125, FORM_H - 52)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
                AddHandler btnCancel.Click, Sub()
                                                _userConfirmed = False
                                                Me.DialogResult = DialogResult.Cancel
                                                Me.Close()
                                            End Sub
                Me.Controls.Add(btnCancel)

                Me.AcceptButton = btnOK
                Me.CancelButton = btnCancel

                AddHandler Me.FormClosing, Sub(sender, e)
                                               If Not _userConfirmed Then
                                                   Me.DialogResult = DialogResult.Cancel
                                               End If
                                           End Sub
            End Sub

            Private Function MakeFieldLabel(ByVal text As String) As Label
                Dim lbl As New Label()
                lbl.Text = text
                lbl.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
                lbl.TextAlign = ContentAlignment.MiddleLeft
                lbl.Dock = DockStyle.Fill
                lbl.AutoEllipsis = True
                lbl.Margin = New Padding(0)
                Return lbl
            End Function

            Private Sub OnBrowseFolder(ByVal sender As Object, ByVal e As EventArgs)
                Using dlg As New FolderBrowserDialog()
                    dlg.Description = "Chọn thư mục để tạo Assembly mới"
                    If IO.Directory.Exists(txtFolder.Text) Then
                        dlg.SelectedPath = txtFolder.Text
                    End If
                    If dlg.ShowDialog() = DialogResult.OK Then
                        txtFolder.Text = dlg.SelectedPath
                    End If
                End Using
            End Sub

            Private Sub OnOKClick(ByVal sender As Object, ByVal e As EventArgs)
                If String.IsNullOrWhiteSpace(txtMainName.Text) Then
                    MessageBox.Show("Tên Main Assembly không được để trống.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtMainName.Focus()
                    Return
                End If
                If String.IsNullOrWhiteSpace(txtFolder.Text) OrElse Not IO.Directory.Exists(txtFolder.Text) Then
                    MessageBox.Show("Thư mục đích không hợp lệ.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Options = New CopyAsmOptions() With {
                    .NewMainName = txtMainName.Text.Trim(),
                    .Prefix = txtPrefix.Text.Trim(),
                    .Suffix = txtSuffix.Text.Trim(),
                    .TargetFolder = txtFolder.Text.Trim()
                }
                _userConfirmed = True
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End Sub
        End Class

        Private Sub CollectAssemblyTree(
            ByVal asm As AssemblyDocument,
            ByVal currentFolder As String,
            ByVal newMainName As String,
            ByVal subPrefix As String,
            ByVal subSuffix As String,
            ByRef fileMap As Dictionary(Of String, String),
            ByRef processed As HashSet(Of String),
            ByVal isMain As Boolean)

            Try
                If Not IO.Directory.Exists(currentFolder) Then
                    IO.Directory.CreateDirectory(currentFolder)
                End If

                Dim asmName As String
                If isMain Then
                    asmName = newMainName
                Else
                    asmName = subPrefix & IO.Path.GetFileNameWithoutExtension(asm.FullFileName) & subSuffix
                End If

                Dim newAsmPath As String = IO.Path.Combine(currentFolder, asmName & IO.Path.GetExtension(asm.FullFileName))

                If Not fileMap.ContainsKey(asm.FullFileName) Then
                    fileMap.Add(asm.FullFileName, newAsmPath)
                End If

                If processed.Contains(asm.FullFileName) Then Exit Sub
                processed.Add(asm.FullFileName)

                For Each occ As ComponentOccurrence In asm.ComponentDefinition.Occurrences
                    Try
                        If occ.Suppressed Then Continue For
                        Dim refDoc As Document = Nothing
                        Try
                            refDoc = occ.Definition.Document
                        Catch
                            Continue For
                        End Try
                        If refDoc Is Nothing Then Continue For
                        If String.IsNullOrWhiteSpace(refDoc.FullFileName) Then Continue For
                        If IsPurchasedPart(refDoc) Then Continue For

                        Dim componentName As String = subPrefix & IO.Path.GetFileNameWithoutExtension(refDoc.FullFileName) & subSuffix

                        If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            Dim subFolder As String = IO.Path.Combine(currentFolder, componentName)
                            If Not IO.Directory.Exists(subFolder) Then
                                IO.Directory.CreateDirectory(subFolder)
                            End If
                            Dim subFilePath As String = IO.Path.Combine(subFolder, componentName & IO.Path.GetExtension(refDoc.FullFileName))
                            If Not fileMap.ContainsKey(refDoc.FullFileName) Then
                                fileMap.Add(refDoc.FullFileName, subFilePath)
                            End If
                            CollectAssemblyTree(CType(refDoc, AssemblyDocument), subFolder, newMainName, subPrefix, subSuffix, fileMap, processed, False)
                        Else
                            Dim partFilePath As String = IO.Path.Combine(currentFolder, componentName & IO.Path.GetExtension(refDoc.FullFileName))
                            If Not fileMap.ContainsKey(refDoc.FullFileName) Then
                                fileMap.Add(refDoc.FullFileName, partFilePath)
                            End If
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try
        End Sub

        Private Function IsPurchasedPart(ByVal doc As Document) As Boolean
            Try
                Dim desc As String = CStr(doc.PropertySets("Design Tracking Properties").Item("Description").Value)
                Return String.Equals(desc.Trim(), "purchased", StringComparison.OrdinalIgnoreCase)
            Catch
                Return False
            End Try
        End Function

        Private Sub ReplaceAllReferences(ByVal asm As AssemblyDocument, ByVal fileMap As Dictionary(Of String, String))
            Try
                ' Bottom-up: xử lý sub trước
                For Each fd As FileDescriptor In asm.File.ReferencedFileDescriptors
                    Try
                        If String.Equals(IO.Path.GetExtension(fd.FullFileName), ".iam", StringComparison.OrdinalIgnoreCase) Then
                            If fileMap.ContainsKey(fd.FullFileName) AndAlso IO.File.Exists(fileMap(fd.FullFileName)) Then
                                Dim subAsm As AssemblyDocument = CType(g_inventorApplication.Documents.Open(fileMap(fd.FullFileName), False), AssemblyDocument)
                                ReplaceAllReferences(subAsm, fileMap)
                                subAsm.Close(True)
                            End If
                        End If
                    Catch
                    End Try
                Next

                ' Relink reference của assembly hiện tại
                For Each fd As FileDescriptor In asm.File.ReferencedFileDescriptors
                    Try
                        If fileMap.ContainsKey(fd.FullFileName) Then
                            Dim newPath As String = fileMap(fd.FullFileName)
                            If IO.File.Exists(newPath) AndAlso Not String.Equals(fd.FullFileName, newPath, StringComparison.OrdinalIgnoreCase) Then
                                fd.ReplaceReference(newPath)
                            End If
                        End If
                    Catch
                    End Try
                Next

                asm.Update2(True)
                asm.Save2(True)
            Catch
            End Try
        End Sub

        Private Sub UpdateAlliProperties(ByVal asm As AssemblyDocument, ByRef processed As HashSet(Of String))
            Try
                If processed.Contains(asm.FullFileName) Then Exit Sub
                processed.Add(asm.FullFileName)
                UpdateDocProps(asm)

                For Each occ As ComponentOccurrence In asm.ComponentDefinition.Occurrences
                    Try
                        If occ.Suppressed Then Continue For
                        Dim doc As Document = occ.Definition.Document
                        If doc Is Nothing Then Continue For
                        If String.IsNullOrWhiteSpace(doc.FullFileName) Then Continue For
                        UpdateDocProps(doc)
                        If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            UpdateAlliProperties(CType(doc, AssemblyDocument), processed)
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try
        End Sub

        Private Sub UpdateDocProps(ByVal doc As Document)
            Try
                If String.IsNullOrWhiteSpace(doc.FullFileName) Then Exit Sub
                Dim newName As String = IO.Path.GetFileNameWithoutExtension(doc.FullFileName)
                Dim props As PropertySet = doc.PropertySets("Design Tracking Properties")
                props.Item("Part Number").Value = newName
                props.Item("Description").Value = newName
                doc.Save2(True)
            Catch
            End Try
        End Sub

    End Module
End Namespace