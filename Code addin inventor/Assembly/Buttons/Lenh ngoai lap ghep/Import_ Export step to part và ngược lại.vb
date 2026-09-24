Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor
Imports Drw = System.Drawing

Namespace ToolInventor2025.Assembly.Buttons.Lenhngoaicumlap

    Public Module Im_EX_step_part

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim invApp As Inventor.Application =
                CType(Interop.Marshal2.GetActiveObject("Inventor.Application"),
                      Inventor.Application)

            Dim choice As Integer = ShowMainForm()

            Select Case choice
                Case 1 : DoImportSTEP(invApp)
                Case 2 : DoExportSelectedSTEP(invApp)
                Case Else : Exit Sub
            End Select
        End Sub


        '=============================================================
        ' FORM CHÍNH
        '=============================================================
        Private Function ShowMainForm() As Integer
            Dim result As Integer = 0

            Using frm As New Form()
                frm.Text = "STEP Tool"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(500, 340)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.BackColor = Drw.Color.FromArgb(245, 245, 245)
                frm.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

                Dim pnlHeader As New Panel()
                pnlHeader.Location = New Drw.Point(0, 0)
                pnlHeader.Size = New Drw.Size(500, 70)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "STEP TOOL"
                lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Import / Export STEP"
                lblSub.Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                AddMenuItem(frm, 1,
                            "IMPORT STEP → Part",
                            "Chọn nhiều file STEP, chuyển thành Part (.ipt) hàng loạt", 90)

                AddMenuItem(frm, 2,
                            "EXPORT → STEP AP214",
                            "Chọn tùy chọn → chọn Part / Sub-Assembly → xuất STEP", 180)

                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(120, 36)
                btnCancel.Location = New Drw.Point(360, 285)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.FlatAppearance.BorderColor = Drw.Color.FromArgb(180, 180, 180)
                btnCancel.BackColor = Drw.Color.FromArgb(245, 245, 245)
                AddHandler btnCancel.Click, Sub()
                                                result = 0
                                                frm.Close()
                                            End Sub
                frm.Controls.Add(btnCancel)
                frm.CancelButton = btnCancel

                frm.ShowDialog()
                result = CInt(frm.Tag)
            End Using

            Return result
        End Function


        Private Sub AddMenuItem(ByVal frm As Form,
                                ByVal value As Integer,
                                ByVal title As String,
                                ByVal desc As String,
                                ByVal top As Integer)

            Dim pnl As New Panel()
            pnl.Location = New Drw.Point(20, top)
            pnl.Size = New Drw.Size(460, 75)
            pnl.BackColor = Drw.Color.White
            pnl.BorderStyle = BorderStyle.FixedSingle
            pnl.Cursor = Cursors.Hand

            Dim lblNum As New Label()
            lblNum.Text = value.ToString()
            lblNum.Font = New Drw.Font("Segoe UI", 20.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblNum.ForeColor = Drw.Color.FromArgb(45, 100, 180)
            lblNum.Location = New Drw.Point(10, 15)
            lblNum.Size = New Drw.Size(55, 45)
            lblNum.TextAlign = Drw.ContentAlignment.MiddleCenter
            pnl.Controls.Add(lblNum)

            Dim sep As New Panel()
            sep.Location = New Drw.Point(75, 15)
            sep.Size = New Drw.Size(1, 45)
            sep.BackColor = Drw.Color.FromArgb(220, 220, 220)
            pnl.Controls.Add(sep)

            Dim lblTitle As New Label()
            lblTitle.Text = title
            lblTitle.Font = New Drw.Font("Segoe UI", 11.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblTitle.ForeColor = Drw.Color.FromArgb(30, 30, 30)
            lblTitle.Location = New Drw.Point(90, 14)
            lblTitle.AutoSize = True
            pnl.Controls.Add(lblTitle)

            Dim lblDesc As New Label()
            lblDesc.Text = desc
            lblDesc.Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            lblDesc.ForeColor = Drw.Color.FromArgb(110, 110, 110)
            lblDesc.Location = New Drw.Point(90, 42)
            lblDesc.AutoSize = True
            pnl.Controls.Add(lblDesc)

            Dim hoverOn As EventHandler = Sub() pnl.BackColor = Drw.Color.FromArgb(235, 242, 252)
            Dim hoverOff As EventHandler = Sub() pnl.BackColor = Drw.Color.White

            AddHandler pnl.MouseEnter, hoverOn
            AddHandler pnl.MouseLeave, hoverOff
            AddHandler lblTitle.MouseEnter, hoverOn
            AddHandler lblTitle.MouseLeave, hoverOff
            AddHandler lblDesc.MouseEnter, hoverOn
            AddHandler lblDesc.MouseLeave, hoverOff
            AddHandler lblNum.MouseEnter, hoverOn
            AddHandler lblNum.MouseLeave, hoverOff

            Dim clickH As EventHandler = Sub(sender, e)
                                             frm.Tag = value
                                             frm.DialogResult = DialogResult.OK
                                             frm.Close()
                                         End Sub
            AddHandler pnl.Click, clickH
            AddHandler lblTitle.Click, clickH
            AddHandler lblDesc.Click, clickH
            AddHandler lblNum.Click, clickH

            frm.Controls.Add(pnl)
        End Sub


        '=============================================================
        ' IMPORT
        '=============================================================
        Private Sub DoImportSTEP(invApp As Inventor.Application)

            Dim sourceDoc As Inventor.PartDocument = Nothing
            Dim settingFile As String = ""
            Dim lastFolder As String = ""
            Dim outputFolder As String = ""
            Dim successCount As Integer = 0
            Dim failCount As Integer = 0
            Dim useTemplateFromSource As Boolean = False

            Try
                If invApp.ActiveDocument IsNot Nothing AndAlso
                   invApp.ActiveDocument.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
                    sourceDoc = CType(invApp.ActiveDocument, Inventor.PartDocument)
                    If sourceDoc.FullFileName <> "" Then useTemplateFromSource = True
                End If
            Catch
                sourceDoc = Nothing
                useTemplateFromSource = False
            End Try

            Try
                Dim appData As String = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.ApplicationData)
                settingFile = IO.Path.Combine(appData, "Inventor_iLogic_STEP_Import.txt")
                If IO.File.Exists(settingFile) Then
                    Dim lines() As String = IO.File.ReadAllLines(settingFile)
                    If lines.Length > 0 Then lastFolder = lines(0)
                End If
            Catch
                lastFolder = ""
            End Try

            Dim stepFiles As New List(Of String)
            Try
                Dim stepDlg As Inventor.FileDialog = Nothing
                invApp.CreateFileDialog(stepDlg)
                stepDlg.DialogTitle = "CHỌN NHIỀU FILE STEP"
                stepDlg.Filter = "STEP Files (*.step;*.stp)|*.step;*.stp|All Files (*.*)|*.*"
                stepDlg.MultiSelectEnabled = True

                If lastFolder <> "" AndAlso IO.Directory.Exists(lastFolder) Then
                    stepDlg.InitialDirectory = lastFolder
                End If

                stepDlg.ShowOpen()
                If String.IsNullOrEmpty(stepDlg.FileName) Then Exit Sub

                Dim sp() As String = stepDlg.FileName.Split("|"c)
                For Each f As String In sp
                    If IO.File.Exists(f) Then stepFiles.Add(f)
                Next

                If stepFiles.Count = 0 Then
                    MessageBox.Show("Không có file STEP hợp lệ.", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                lastFolder = IO.Path.GetDirectoryName(stepFiles(0))
            Catch ex As Exception
                MessageBox.Show("Lỗi chọn file STEP: " & ex.Message, "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try

            Try
                Dim folderDlg As New FolderBrowserDialog()
                folderDlg.Description = "Chọn thư mục lưu các file Part mới"
                folderDlg.ShowNewFolderButton = True
                If lastFolder <> "" AndAlso IO.Directory.Exists(lastFolder) Then
                    folderDlg.SelectedPath = lastFolder
                End If
                If folderDlg.ShowDialog() <> DialogResult.OK Then Exit Sub
                outputFolder = folderDlg.SelectedPath
                Try : IO.File.WriteAllText(settingFile, outputFolder) : Catch : End Try
            Catch ex As Exception
                MessageBox.Show("Lỗi chọn folder: " & ex.Message, "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try

            For Each stepFile As String In stepFiles
                Dim newDoc As Inventor.PartDocument = Nothing
                Dim importedComp As Inventor.ImportedComponent = Nothing
                Dim newFile As String = ""

                Try
                    Dim defaultName As String = IO.Path.GetFileNameWithoutExtension(stepFile) & ".ipt"
                    newFile = IO.Path.Combine(outputFolder, defaultName)

                    If useTemplateFromSource Then
                        If String.Compare(IO.Path.GetFullPath(sourceDoc.FullFileName),
                                          IO.Path.GetFullPath(newFile), True) = 0 Then
                            failCount += 1
                            Continue For
                        End If
                    End If

                    If IO.File.Exists(newFile) Then
                        Dim ans As DialogResult = MessageBox.Show(
                            "File đã tồn tại:" & vbCrLf & newFile & vbCrLf & vbCrLf & "Ghi đè?",
                            "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                        If ans = DialogResult.No Then
                            failCount += 1
                            Continue For
                        End If
                    End If

                    If useTemplateFromSource Then
                        sourceDoc.SaveAs(newFile, True)
                        Dim openDoc As Inventor.Document = invApp.Documents.Open(newFile, True)
                        newDoc = CType(openDoc, Inventor.PartDocument)
                    Else
                        Dim templatePath As String =
                            invApp.FileManager.GetTemplateFile(Inventor.DocumentTypeEnum.kPartDocumentObject)
                        newDoc = CType(invApp.Documents.Add(
                            Inventor.DocumentTypeEnum.kPartDocumentObject, templatePath, True),
                            Inventor.PartDocument)
                        newDoc.SaveAs(newFile, False)
                    End If

                    Dim compDef As Inventor.PartComponentDefinition = newDoc.ComponentDefinition
                    Dim importDef As Inventor.ImportedGenericComponentDefinition =
                        compDef.ReferenceComponents.ImportedComponents.CreateDefinition(stepFile)

                    importDef.ImportedAssemblyOrganizationType =
                        Inventor.ImportedAssemblyOrganizationTypeEnum.kImportedAsMultibodyPart
                    importedComp = compDef.ReferenceComponents.ImportedComponents.Add(importDef)
                    newDoc.Update()

                    Dim npFeatures As Inventor.NonParametricBaseFeatures =
                        compDef.Features.NonParametricBaseFeatures
                    Dim importFeature As Inventor.NonParametricBaseFeature =
                        npFeatures.Item(npFeatures.Count)

                    For i As Integer = 1 To importFeature.InputSurfaceBodies.Count
                        Dim srcBody As Inventor.SurfaceBody = importFeature.InputSurfaceBodies.Item(i)
                        If srcBody.IsSolid Then
                            Dim copied As Inventor.SurfaceBody = invApp.TransientBRep.Copy(srcBody)
                            npFeatures.Add(copied)
                        End If
                    Next
                    newDoc.Update()

                    Try
                        If importedComp IsNot Nothing Then
                            Try : importedComp.BreakLinkToFile() : Catch : End Try
                            Try : importedComp.Delete() : Catch : End Try
                        End If
                    Catch
                    End Try

                    Try
                        Dim ics = newDoc.ComponentDefinition.ReferenceComponents.ImportedComponents
                        For i As Integer = ics.Count To 1 Step -1
                            Try : ics.Item(i).Delete() : Catch : End Try
                        Next
                    Catch
                    End Try

                    Try
                        For i As Integer = newDoc.ReferencedOLEFileDescriptors.Count To 1 Step -1
                            Try : newDoc.ReferencedOLEFileDescriptors.Item(i).Delete() : Catch : End Try
                        Next
                    Catch
                    End Try

                    newDoc.Update()
                    newDoc.Save()
                    successCount += 1

                Catch ex As Exception
                    failCount += 1
                    MessageBox.Show("Lỗi xử lý file:" & vbCrLf & stepFile & vbCrLf & vbCrLf & ex.Message,
                                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Finally
                    Try
                        If newDoc IsNot Nothing Then newDoc.Close(True)
                    Catch
                    End Try
                End Try
            Next

            If sourceDoc IsNot Nothing Then
                Try : sourceDoc.Activate() : Catch : End Try
            End If

            MessageBox.Show(
                "HOÀN TẤT!" & vbCrLf & vbCrLf &
                "Thành công : " & successCount.ToString() & vbCrLf &
                "Thất bại   : " & failCount.ToString() & vbCrLf & vbCrLf &
                "Folder lưu : " & outputFolder,
                "Kết quả", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub


        '=============================================================
        ' EXPORT — LUỒNG MỚI
        '   1. Form options
        '   2. Pick chọn Part/Sub-Assembly
        '   3. Xuất STEP
        '=============================================================
        Private Sub DoExportSelectedSTEP(invApp As Inventor.Application)

            '===== 1. Kiểm tra Assembly =====
            If invApp.ActiveDocument Is Nothing OrElse
               invApp.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Mở Assembly trước khi dùng chức năng này!", "Thông báo")
                Exit Sub
            End If

            Dim asmDoc As Inventor.AssemblyDocument = CType(invApp.ActiveDocument, Inventor.AssemblyDocument)

            '===== 2. Form options =====
            Dim opt As ExportOptions = ShowExportOptionsForm()
            If opt Is Nothing OrElse opt.Cancelled Then Exit Sub

            '===== 3. PICK CHỌN PART / SUB-ASSEMBLY =====
            Try : asmDoc.SelectSet.Clear() : Catch : End Try

            Dim pickedOccs As New List(Of Inventor.ComponentOccurrence)
            Dim stopPick As Boolean = False

            '--- Tạo form điều khiển nổi ---
            Dim ctrlForm As New PickControlForm()
            AddHandler ctrlForm.OKClicked, Sub()
                                               stopPick = True
                                               Try : invApp.CommandManager.StopActiveCommand() : Catch : End Try
                                               Try : ctrlForm.Close() : Catch : End Try
                                           End Sub

            '--- Hiện form (modeless) ---
            Try
                ctrlForm.Show()
                ctrlForm.Refresh()
            Catch
            End Try

            '--- Vòng lặp Pick ---
            While Not stopPick
                Dim prompt As String = "Chọn Part / Sub-Assembly (" &
                                       pickedOccs.Count.ToString() &
                                       " đã chọn). ESC hoặc bấm KẾT THÚC."

                Dim picked As Object = Nothing
                Try
                    picked = invApp.CommandManager.Pick(
                        Inventor.SelectionFilterEnum.kAssemblyOccurrenceFilter, prompt)
                Catch
                    Exit While
                End Try

                If picked Is Nothing Then Exit While

                Dim occ As Inventor.ComponentOccurrence = TryCast(picked, Inventor.ComponentOccurrence)

                If occ Is Nothing Then
                    Try
                        Dim face As Inventor.Face = TryCast(picked, Inventor.Face)
                        If face IsNot Nothing Then occ = face.ContainingOccurrence
                    Catch
                    End Try
                End If

                If occ IsNot Nothing AndAlso Not pickedOccs.Contains(occ) Then
                    pickedOccs.Add(occ)
                    Try : ctrlForm.UpdateCount(pickedOccs.Count) : Catch : End Try
                End If
            End While

            '--- Đóng form điều khiển ---
            Try
                If Not ctrlForm.IsDisposed Then ctrlForm.Close()
                ctrlForm.Dispose()
            Catch
            End Try

            If pickedOccs.Count = 0 Then
                MessageBox.Show("Bạn chưa chọn Part / Sub-Assembly nào.", "Thông báo")
                Exit Sub
            End If
            '===== Chuyển occurrence → dictionary document =====
            Dim selectedDocs As New Dictionary(Of String, Inventor.Document)

            For Each occ As Inventor.ComponentOccurrence In pickedOccs
                Try
                    If occ.Suppressed Then Continue For
                    If occ.DefinitionDocumentType <> Inventor.DocumentTypeEnum.kPartDocumentObject AndAlso
                       occ.DefinitionDocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                        Continue For
                    End If

                    Dim doc As Inventor.Document = occ.Definition.Document
                    If doc Is Nothing OrElse String.IsNullOrEmpty(doc.FullFileName) Then Continue For

                    Dim fullPath As String = IO.Path.GetFullPath(doc.FullFileName)
                    If Not selectedDocs.ContainsKey(fullPath) Then
                        selectedDocs.Add(fullPath, doc)
                    End If
                Catch
                End Try
            Next

            If selectedDocs.Count = 0 Then
                MessageBox.Show("Không có Part / Sub-Assembly hợp lệ.", "Thông báo")
                Exit Sub
            End If

            '===== 4. Lấy STEP Translator =====
            Dim stepTranslator As Inventor.TranslatorAddIn = Nothing
            Try
                stepTranslator = CType(invApp.ApplicationAddIns.ItemById(
                    "{90AF7F40-0C01-11D5-8E83-0010B541CD80}"), Inventor.TranslatorAddIn)
            Catch
            End Try
            If stepTranslator Is Nothing Then
                MessageBox.Show("Không tìm thấy STEP Translator.", "Lỗi")
                Exit Sub
            End If

            Dim successCount As Integer = 0
            Dim failCount As Integer = 0
            Dim tempFiles As New List(Of String)

            Dim useFullSaveDialog As Boolean = (selectedDocs.Count = 1)
            Dim outputFolder As String = ""
            Dim singleSavePath As String = ""

            If useFullSaveDialog Then
                Dim firstDoc As Inventor.Document = selectedDocs.First().Value
                Dim defaultName As String =
                    IO.Path.GetFileNameWithoutExtension(firstDoc.FullFileName) & ".stp"

                Dim saveDlg As Inventor.FileDialog = Nothing
                invApp.CreateFileDialog(saveDlg)
                saveDlg.DialogTitle = "Lưu STEP " & opt.ProtocolName
                saveDlg.Filter = "STEP Files (*.stp)|*.stp|STEP Files (*.step)|*.step"
                saveDlg.FileName = defaultName
                saveDlg.ShowSave()

                If saveDlg.FileName = "" Then Exit Sub
                singleSavePath = saveDlg.FileName
                outputFolder = IO.Path.GetDirectoryName(singleSavePath)
            Else
                Dim folderDlg As New FolderBrowserDialog()
                folderDlg.Description = "Chọn thư mục lưu các file STEP"
                folderDlg.ShowNewFolderButton = True
                If folderDlg.ShowDialog() <> DialogResult.OK Then Exit Sub
                outputFolder = folderDlg.SelectedPath
            End If

            For Each kvp As KeyValuePair(Of String, Inventor.Document) In selectedDocs
                Dim doc As Inventor.Document = kvp.Value
                Dim srcPath As String = kvp.Key
                Dim exportDoc As Inventor.Document = doc
                Dim isTemp As Boolean = False

                Try
                    '--- Nếu là Sub-Assembly VÀ chọn chế độ AsmAsPart ---
                    If doc.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject AndAlso
                       opt.AsmMode = 1 Then

                        Dim tempPartPath As String = IO.Path.Combine(
                            IO.Path.GetTempPath(),
                            "TEMP_" & System.Guid.NewGuid().ToString("N").Substring(0, 8) & ".ipt")

                        Dim templatePath As String =
                            invApp.FileManager.GetTemplateFile(Inventor.DocumentTypeEnum.kPartDocumentObject)
                        Dim tempPart As Inventor.PartDocument = CType(invApp.Documents.Add(
                            Inventor.DocumentTypeEnum.kPartDocumentObject, templatePath, True),
                            Inventor.PartDocument)
                        tempPart.SaveAs(tempPartPath, False)

                        Dim derivedDef As Inventor.DerivedAssemblyDefinition =
                            tempPart.ComponentDefinition.ReferenceComponents.
                            DerivedAssemblyComponents.CreateDefinition(srcPath)
                        derivedDef.DeriveStyle = Inventor.DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
                        tempPart.ComponentDefinition.ReferenceComponents.
                            DerivedAssemblyComponents.Add(derivedDef)
                        tempPart.Update()
                        tempPart.Save2(True)

                        exportDoc = tempPart
                        isTemp = True
                        tempFiles.Add(tempPartPath)
                    End If

                    '--- Đường dẫn STEP ---
                    Dim stepFullPath As String
                    If useFullSaveDialog Then
                        stepFullPath = singleSavePath
                    Else
                        Dim baseName As String = IO.Path.GetFileNameWithoutExtension(srcPath)
                        If isTemp Then baseName &= "_PART"
                        stepFullPath = IO.Path.Combine(outputFolder, baseName & ".stp")
                    End If

                    If IO.File.Exists(stepFullPath) AndAlso Not useFullSaveDialog Then
                        Dim ans As DialogResult = MessageBox.Show(
                            "File đã tồn tại:" & vbCrLf & stepFullPath & vbCrLf & "Ghi đè?",
                            "Xác nhận", MessageBoxButtons.YesNo)
                        If ans = DialogResult.No Then
                            failCount += 1
                            Continue For
                        End If
                    End If

                    '--- Tạo context + options ---
                    Dim oContext As Inventor.TranslationContext =
                        invApp.TransientObjects.CreateTranslationContext()
                    oContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism

                    Dim oOptions As Inventor.NameValueMap = invApp.TransientObjects.CreateNameValueMap()
                    If stepTranslator.HasSaveCopyAsOptions(exportDoc, oContext, oOptions) Then

                        '--- Protocol ---
                        Try : oOptions.Value("ApplicationProtocolType") = opt.Protocol : Catch : End Try

                        '--- Đơn vị ---
                        If opt.UnitMode = 1 Then
                            Try : oOptions.Value("ExportUnit") = "mm" : Catch : End Try
                            Try : oOptions.Value("Unit") = 1 : Catch : End Try
                        ElseIf opt.UnitMode = 2 Then
                            Try : oOptions.Value("ExportUnit") = "in" : Catch : End Try
                            Try : oOptions.Value("Unit") = 2 : Catch : End Try
                        End If

                        '--- Tinh chỉnh ---
                        Try : oOptions.Value("Refine") = opt.Refine : Catch : End Try
                        Try : oOptions.Value("RefineModel") = opt.Refine : Catch : End Try
                        Try : oOptions.Value("EnableRefine") = opt.Refine : Catch : End Try
                    End If

                    Dim oData As Inventor.DataMedium = invApp.TransientObjects.CreateDataMedium()
                    oData.FileName = stepFullPath

                    stepTranslator.SaveCopyAs(exportDoc, oContext, oOptions, oData)
                    successCount += 1

                Catch ex As Exception
                    failCount += 1
                    MessageBox.Show("Lỗi xuất:" & vbCrLf & srcPath & vbCrLf & ex.Message, "Lỗi")
                Finally
                    If isTemp AndAlso exportDoc IsNot Nothing Then
                        Try : exportDoc.Close(True) : Catch : End Try
                    End If
                End Try
            Next

            For Each tmp As String In tempFiles
                Try : If IO.File.Exists(tmp) Then IO.File.Delete(tmp)
                Catch : End Try
            Next

            MessageBox.Show("EXPORT STEP " & opt.ProtocolName & " HOÀN TẤT!" & vbCrLf & vbCrLf &
                            "Định dạng   : " & opt.ProtocolName & vbCrLf &
                            "Đơn vị      : " & opt.UnitName & vbCrLf &
                            "Tinh chỉnh  : " & opt.RefineName & vbCrLf &
                            "Sub-Asm     : " & opt.SubMode & vbCrLf & vbCrLf &
                            "Thành công : " & successCount & vbCrLf &
                            "Thất bại   : " & failCount & vbCrLf & vbCrLf &
                            "Folder     : " & outputFolder, "Kết quả")
        End Sub


        '=============================================================
        ' OPTIONS CLASS
        '=============================================================
        Private Class ExportOptions
            Public Protocol As Integer = 2              ' 1=AP203, 2=AP214, 3=AP242
            Public ProtocolName As String = "AP214"
            Public UnitMode As Integer = 0              ' 0=Auto, 1=mm, 2=inch
            Public UnitName As String = "Tự động"
            Public AsmMode As Integer = 1               ' 1=AsmAsPart, 2=AsmKeepAsm
            Public SubMode As String = "Dạng Part"
            Public Refine As Boolean = True
            Public RefineName As String = "Có"
            Public Cancelled As Boolean = False
        End Class
        '=============================================================
        ' ⭐ FORM NHỎ NỔI — HIỆN TRONG LÚC PICK
        '=============================================================
        Private Class PickControlForm
            Inherits Form

            Public Event OKClicked As EventHandler

            Private lblCount As Label

            Public Sub New()
                Me.Text = "Đang chọn..."
                Me.AutoScaleMode = AutoScaleMode.None
                Me.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                Me.ClientSize = New Drw.Size(320, 130)
                Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
                Me.StartPosition = FormStartPosition.Manual
                Me.MaximizeBox = False
                Me.MinimizeBox = False
                Me.ShowInTaskbar = False
                Me.TopMost = True
                Me.BackColor = Drw.Color.FromArgb(245, 245, 245)
                Me.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

                ' Đặt góc phải trên màn hình
                Try
                    Dim wa As Drw.Rectangle = Screen.PrimaryScreen.WorkingArea
                    Me.Location = New Drw.Point(wa.Right - Me.Width - 30, wa.Top + 80)
                Catch
                    Me.Location = New Drw.Point(800, 100)
                End Try

                '===== HEADER =====
                Dim pnlHeader As New Panel()
                pnlHeader.Location = New Drw.Point(0, 0)
                pnlHeader.Size = New Drw.Size(320, 30)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                Me.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "CHỌN ĐỐI TƯỢNG XUẤT"
                lblTitle.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                '===== SỐ ĐÃ CHỌN =====
                lblCount = New Label()
                lblCount.Text = "Đã chọn: 0"
                lblCount.Location = New Drw.Point(15, 42)
                lblCount.Size = New Drw.Size(290, 25)
                lblCount.Font = New Drw.Font("Segoe UI", 11.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblCount.ForeColor = Drw.Color.FromArgb(45, 100, 180)
                lblCount.TextAlign = Drw.ContentAlignment.MiddleLeft
                Me.Controls.Add(lblCount)

                '===== NÚT KẾT THÚC =====
                Dim btnOK As New Button()
                btnOK.Text = "KẾT THÚC  (ESC)"
                btnOK.Size = New Drw.Size(290, 38)
                btnOK.Location = New Drw.Point(15, 78)
                btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
                btnOK.ForeColor = Drw.Color.White
                btnOK.FlatStyle = FlatStyle.Flat
                btnOK.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                AddHandler btnOK.Click, Sub()
                                            RaiseEvent OKClicked(Me, EventArgs.Empty)
                                        End Sub
                Me.Controls.Add(btnOK)
            End Sub

            Public Sub UpdateCount(ByVal count As Integer)
                Try
                    lblCount.Text = "Đã chọn: " & count.ToString()
                Catch
                End Try
            End Sub
        End Class

        '=============================================================
        ' FORM CHỌN KIỂU XUẤT STEP
        '=============================================================
        Private Function ShowExportOptionsForm() As ExportOptions
            Dim opt As New ExportOptions()
            opt.Cancelled = True

            Using frm As New Form()
                frm.Text = "Tùy chọn xuất STEP"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(580, 550)
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
                pnlHeader.Size = New Drw.Size(580, 60)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "TÙY CHỌN XUẤT STEP"
                lblTitle.Font = New Drw.Font("Segoe UI", 14.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Chọn định dạng, đơn vị, tinh chỉnh"
                lblSub.Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 18
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== GROUP 1: PROTOCOL =====
                Dim gb1 As New GroupBox() With {
                    .Text = "Định dạng STEP",
                    .Location = New Drw.Point(15, 75),
                    .Size = New Drw.Size(550, 105),
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb1)

                Dim rdoAP203 As New RadioButton() With {
                    .Text = "AP203  —  Không màu, kích thước cơ bản",
                    .Location = New Drw.Point(20, 25),
                    .Size = New Drw.Size(500, 22),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                Dim rdoAP214 As New RadioButton() With {
                    .Text = "AP214  —  Có màu, layer  (khuyến nghị)",
                    .Location = New Drw.Point(20, 50),
                    .Size = New Drw.Size(500, 22),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                Dim rdoAP242 As New RadioButton() With {
                    .Text = "AP242  —  Mới nhất, hỗ trợ PMI / annotation",
                    .Location = New Drw.Point(20, 75),
                    .Size = New Drw.Size(500, 22),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb1.Controls.Add(rdoAP203)
                gb1.Controls.Add(rdoAP214)
                gb1.Controls.Add(rdoAP242)

                '===== GROUP 2: ĐƠN VỊ =====
                Dim gb2 As New GroupBox() With {
                    .Text = "Đơn vị",
                    .Location = New Drw.Point(15, 190),
                    .Size = New Drw.Size(550, 75),
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb2)

                Dim rdoUnitAuto As New RadioButton() With {
                    .Text = "Tự động",
                    .Location = New Drw.Point(20, 28),
                    .Size = New Drw.Size(120, 25),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                Dim rdoUnitMm As New RadioButton() With {
                    .Text = "mm",
                    .Location = New Drw.Point(160, 28),
                    .Size = New Drw.Size(80, 25),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                Dim rdoUnitInch As New RadioButton() With {
                    .Text = "inch",
                    .Location = New Drw.Point(260, 28),
                    .Size = New Drw.Size(100, 25),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb2.Controls.Add(rdoUnitAuto)
                gb2.Controls.Add(rdoUnitMm)
                gb2.Controls.Add(rdoUnitInch)

                '===== GROUP 3: TINH CHỈNH =====
                Dim gbRefine As New GroupBox() With {
                    .Text = "Tinh chỉnh hình học",
                    .Location = New Drw.Point(15, 275),
                    .Size = New Drw.Size(550, 75),
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gbRefine)

                Dim rdoRefineYes As New RadioButton() With {
                    .Text = "Có (Refine)  —  Mượt, chính xác",
                    .Location = New Drw.Point(20, 28),
                    .Size = New Drw.Size(240, 25),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                Dim rdoRefineNo As New RadioButton() With {
                    .Text = "Không  —  Giữ nguyên mặt, nhẹ hơn",
                    .Location = New Drw.Point(280, 28),
                    .Size = New Drw.Size(260, 25),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gbRefine.Controls.Add(rdoRefineYes)
                gbRefine.Controls.Add(rdoRefineNo)

                '===== GROUP 4: SUB-ASSEMBLY =====
                Dim gb3 As New GroupBox() With {
                    .Text = "Chế độ xuất Sub-Assembly",
                    .Location = New Drw.Point(15, 360),
                    .Size = New Drw.Size(550, 100),
                    .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb3)

                Dim rdoSubPart As New RadioButton() With {
                    .Text = "Dạng PART  —  Gộp cụm thành 1 Part duy nhất (Derived)",
                    .Location = New Drw.Point(20, 25),
                    .Size = New Drw.Size(500, 22),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                Dim rdoSubAsm As New RadioButton() With {
                    .Text = "Dạng CỤM LẮP  —  Giữ nguyên cấu trúc Assembly",
                    .Location = New Drw.Point(20, 55),
                    .Size = New Drw.Size(500, 22),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb3.Controls.Add(rdoSubPart)
                gb3.Controls.Add(rdoSubAsm)

                '===== NÚT =====
                Dim btnOK As New Button()
                btnOK.Text = "TIẾP TỤC CHỌN"
                btnOK.Size = New Drw.Size(160, 40)
                btnOK.Location = New Drw.Point(400, 480)
                btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
                btnOK.ForeColor = Drw.Color.White
                btnOK.FlatStyle = FlatStyle.Flat
                btnOK.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                btnOK.DialogResult = DialogResult.OK
                frm.Controls.Add(btnOK)

                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(100, 40)
                btnCancel.Location = New Drw.Point(250, 480)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                btnCancel.DialogResult = DialogResult.Cancel
                frm.Controls.Add(btnCancel)
                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel

                If frm.ShowDialog() <> DialogResult.OK Then
                    opt.Cancelled = True
                    Return opt
                End If

                '--- Lấy giá trị ---
                If rdoAP203.Checked Then
                    opt.Protocol = 1 : opt.ProtocolName = "AP203"
                ElseIf rdoAP242.Checked Then
                    opt.Protocol = 3 : opt.ProtocolName = "AP242"
                Else
                    opt.Protocol = 2 : opt.ProtocolName = "AP214"
                End If

                If rdoUnitMm.Checked Then
                    opt.UnitMode = 1 : opt.UnitName = "mm"
                ElseIf rdoUnitInch.Checked Then
                    opt.UnitMode = 2 : opt.UnitName = "inch"
                Else
                    opt.UnitMode = 0 : opt.UnitName = "Tự động"
                End If

                If rdoRefineYes.Checked Then
                    opt.Refine = True : opt.RefineName = "Có"
                Else
                    opt.Refine = False : opt.RefineName = "Không"
                End If

                If rdoSubAsm.Checked Then
                    opt.AsmMode = 2 : opt.SubMode = "Dạng Cụm"
                Else
                    opt.AsmMode = 1 : opt.SubMode = "Dạng Part"
                End If

                opt.Cancelled = False
            End Using

            Return opt
        End Function

    End Module

End Namespace