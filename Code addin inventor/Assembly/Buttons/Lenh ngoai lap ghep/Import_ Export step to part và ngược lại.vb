Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Lenhngoaicumlap
    Public Module Im_EX_step_part
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)

            '=====================================================
            ' CHỌN CHỨC NĂNG (Yes / No / Cancel)
            '=====================================================
            Dim result As DialogResult = MessageBox.Show(
        "Chọn chức năng:" & vbCrLf & vbCrLf &
        "YES  = Import nhiều STEP → Part" & vbCrLf &
        "NO   = Export Selected → STEP AP214" & vbCrLf &
        "CANCEL = Thoát",
        "STEP Tool - Inventor 2020",
        MessageBoxButtons.YesNoCancel,
        MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                Call DoImportSTEP(invApp)
            ElseIf result = DialogResult.No Then
                Call DoExportSelectedSTEP(invApp)
            Else
                ' Cancel → thoát
                Exit Sub
            End If

        End Sub

        '=============================================================
        ' LỰA CHỌN 1 – IMPORT (y nguyên code bạn gửi)
        '=============================================================
        Private Sub DoImportSTEP(invApp As Inventor.Application)

            Dim sourceDoc As PartDocument = Nothing
            Dim settingFile As String = ""
            Dim lastFolder As String = ""
            Dim outputFolder As String = ""
            Dim successCount As Integer = 0
            Dim failCount As Integer = 0
            Dim useTemplateFromSource As Boolean = False

            ' 1. Kiểm tra file hiện tại
            Try
                If invApp.ActiveDocument IsNot Nothing AndAlso
               invApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then

                    sourceDoc = CType(invApp.ActiveDocument, PartDocument)

                    If sourceDoc.FullFileName <> "" Then
                        useTemplateFromSource = True
                    End If
                End If
            Catch
                sourceDoc = Nothing
                useTemplateFromSource = False
            End Try

            ' 2. Đọc folder nhớ
            Try
                Dim appData As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData)

                settingFile = IO.Path.Combine(appData, "Inventor_iLogic_STEP_Import.txt")

                If IO.File.Exists(settingFile) Then
                    Dim lines() As String = IO.File.ReadAllLines(settingFile)
                    If lines.Length > 0 Then lastFolder = lines(0)
                End If
            Catch
                lastFolder = ""
            End Try
<<<<<<< HEAD
            '--- 3. Chọn nhiều file STEP ----------------------------------
=======

            ' 3. Chọn nhiều file STEP
>>>>>>> c096192 (Add project files.)
            Dim stepFiles As New List(Of String)
            Try
                Dim stepDlg As Inventor.FileDialog = Nothing
                invApp.CreateFileDialog(stepDlg)

                stepDlg.DialogTitle = "CHỌN NHIỀU FILE STEP"
<<<<<<< HEAD
                stepDlg.Filter = "STEP Files (*.step;*.stp)|*.step;*.stp|All Files (*.*)|*.*"
=======
                stepDlg.Filter = "ALL FILE" '"STEP Files (*.step;*.stp)|*.step;*.stp"
>>>>>>> c096192 (Add project files.)
                stepDlg.MultiSelectEnabled = True

                If lastFolder <> "" AndAlso IO.Directory.Exists(lastFolder) Then
                    stepDlg.InitialDirectory = lastFolder
                End If

                stepDlg.ShowOpen()
<<<<<<< HEAD
                If String.IsNullOrEmpty(stepDlg.FileName) Then Exit Sub

                ' ✅ SỬA Ở ĐÂY: Inventor FileDialog không có thuộc tính 'Files'.
                ' Khi MultiSelectEnabled = True, thuộc tính 'FileName' sẽ trả về 
                ' một chuỗi chứa tất cả các file được chọn, phân tách bởi dấu '|'.
                Dim sp() As String = stepDlg.FileName.Split("|"c)
                For Each f As String In sp
=======

                If stepDlg.FileName = "" Then Exit Sub

                Dim selected() As String = stepDlg.FileName.Split("|"c)
                For Each f As String In selected
>>>>>>> c096192 (Add project files.)
                    If IO.File.Exists(f) Then stepFiles.Add(f)
                Next

                If stepFiles.Count = 0 Then
<<<<<<< HEAD
                    MessageBox.Show("Không có file STEP hợp lệ.", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
=======
                    MessageBox.Show("Không có file STEP hợp lệ.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
>>>>>>> c096192 (Add project files.)
                    Exit Sub
                End If

                lastFolder = IO.Path.GetDirectoryName(stepFiles(0))
<<<<<<< HEAD

            Catch ex As Exception
                MessageBox.Show("Lỗi chọn file STEP: " & ex.Message, "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
=======
            Catch ex As Exception
                MessageBox.Show("Lỗi chọn file STEP: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
>>>>>>> c096192 (Add project files.)
                Exit Sub
            End Try

            ' 4. Chọn folder lưu riêng
            Try
                Dim folderDlg As New FolderBrowserDialog()
                folderDlg.Description = "Chọn thư mục lưu các file Part mới"
                folderDlg.ShowNewFolderButton = True

                If lastFolder <> "" AndAlso IO.Directory.Exists(lastFolder) Then
                    folderDlg.SelectedPath = lastFolder
                End If

                If folderDlg.ShowDialog() <> DialogResult.OK Then Exit Sub

                outputFolder = folderDlg.SelectedPath

                Try
                    IO.File.WriteAllText(settingFile, outputFolder)
                Catch
                End Try
            Catch ex As Exception
                MessageBox.Show("Lỗi chọn folder: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try

            ' 5. Xử lý từng file STEP
            For Each stepFile As String In stepFiles
                Dim newDoc As PartDocument = Nothing
                Dim importedComp As ImportedComponent = Nothing
                Dim newFile As String = ""
                Dim nativeBodyCount As Integer = 0

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

                    ' Tạo file Part mới
                    If useTemplateFromSource Then
                        sourceDoc.SaveAs(newFile, True)
                        Dim openDoc As Document = invApp.Documents.Open(newFile, True)
                        newDoc = CType(openDoc, PartDocument)
                    Else
                        Dim templatePath As String = invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject)
                        newDoc = CType(invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, templatePath, True), PartDocument)
                        newDoc.SaveAs(newFile, False)
                    End If

                    ' Import STEP multi-body
                    Dim compDef As PartComponentDefinition = newDoc.ComponentDefinition
                    Dim importDef As ImportedGenericComponentDefinition =
                    compDef.ReferenceComponents.ImportedComponents.CreateDefinition(stepFile)

                    importDef.ImportedAssemblyOrganizationType = ImportedAssemblyOrganizationTypeEnum.kImportedAsMultibodyPart
                    importedComp = compDef.ReferenceComponents.ImportedComponents.Add(importDef)
                    newDoc.Update()

                    ' Copy body → Native Non-Associative
                    Dim npFeatures As NonParametricBaseFeatures = compDef.Features.NonParametricBaseFeatures
                    Dim importFeature As NonParametricBaseFeature = npFeatures.Item(npFeatures.Count)

                    For i As Integer = 1 To importFeature.InputSurfaceBodies.Count
                        Dim srcBody As SurfaceBody = importFeature.InputSurfaceBodies.Item(i)
                        If srcBody.IsSolid Then
                            Dim copied As SurfaceBody = invApp.TransientBRep.Copy(srcBody)
                            Dim nativeFeat As NonParametricBaseFeature = npFeatures.Add(copied)
                            If nativeFeat IsNot Nothing AndAlso Not nativeFeat.IsAssociative Then
                                nativeBodyCount += 1
                            End If
                        End If
                    Next
                    newDoc.Update()

                    ' Xóa ImportedComponent + 3rd Party
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

            ' 6. Kết quả
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
        ' LỰA CHỌN 2 – EXPORT SELECTED → STEP AP214
        '=============================================================
        Private Sub DoExportSelectedSTEP(invApp As Inventor.Application)

            If invApp.ActiveDocument Is Nothing OrElse
               invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Mở Assembly và chọn Part / Sub-Assembly trước!", "Thông báo")
                Exit Sub
            End If

            Dim asmDoc As AssemblyDocument = CType(invApp.ActiveDocument, AssemblyDocument)
            Dim selectedDocs As New Dictionary(Of String, Document)

            Dim selSet As SelectSet = asmDoc.SelectSet
            If selSet.Count = 0 Then
                MessageBox.Show("Hãy chọn ít nhất 1 Part hoặc Sub-Assembly!", "Thông báo")
                Exit Sub
            End If

            For Each obj As Object In selSet
                Try
                    Dim occ As ComponentOccurrence = Nothing
                    If TypeOf obj Is ComponentOccurrence Then
                        occ = CType(obj, ComponentOccurrence)
                    ElseIf TypeOf obj Is ComponentOccurrenceProxy Then
                        occ = CType(obj, ComponentOccurrenceProxy).NativeObject
                    End If

                    If occ Is Nothing OrElse occ.Suppressed Then Continue For
                    If occ.DefinitionDocumentType <> DocumentTypeEnum.kPartDocumentObject AndAlso
                       occ.DefinitionDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Continue For

                    Dim doc As Document = occ.Definition.Document
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

            ' Hỏi chế độ xuất Sub-Assembly
            Dim hasAssembly As Boolean = False
            For Each d As Document In selectedDocs.Values
                If d.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    hasAssembly = True
                    Exit For
                End If
            Next

            Dim exportAsmAsPart As Boolean = False
            If hasAssembly Then
                Dim ans As DialogResult = MessageBox.Show(
                    "Có Sub-Assembly trong lựa chọn." & vbCrLf & vbCrLf &
                    "YES = Xuất dạng PART (Derived)" & vbCrLf &
                    "NO  = Xuất dạng CỤM LẮP (Assembly)",
                    "Chế độ xuất Sub-Assembly", MessageBoxButtons.YesNoCancel)
                If ans = DialogResult.Cancel Then Exit Sub
                exportAsmAsPart = (ans = DialogResult.Yes)
            End If

            ' Lấy STEP Translator
            Dim stepTranslator As TranslatorAddIn = Nothing
            Try
                stepTranslator = CType(invApp.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}"), TranslatorAddIn)
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
                ' 1 file → hộp thoại Save giống Inventor (đổi tên + chọn thư mục)
                Dim firstDoc As Document = selectedDocs.First().Value
                Dim defaultName As String = IO.Path.GetFileNameWithoutExtension(firstDoc.FullFileName) & ".stp"

                Dim saveDlg As Inventor.FileDialog = Nothing
                invApp.CreateFileDialog(saveDlg)
                saveDlg.DialogTitle = "Lưu STEP AP214"
                saveDlg.Filter = "STEP Files (*.stp)|*.stp|STEP Files (*.step)|*.step"
                saveDlg.FileName = defaultName
                saveDlg.ShowSave()

                If saveDlg.FileName = "" Then Exit Sub
                singleSavePath = saveDlg.FileName
                outputFolder = IO.Path.GetDirectoryName(singleSavePath)
            Else
                ' Nhiều file → chọn thư mục
                Dim folderDlg As New FolderBrowserDialog()
                folderDlg.Description = "Chọn thư mục lưu các file STEP"
                folderDlg.ShowNewFolderButton = True
                If folderDlg.ShowDialog() <> DialogResult.OK Then Exit Sub
                outputFolder = folderDlg.SelectedPath
            End If

            For Each kvp As KeyValuePair(Of String, Document) In selectedDocs
                Dim doc As Document = kvp.Value
                Dim srcPath As String = kvp.Key
                Dim exportDoc As Document = doc
                Dim isTemp As Boolean = False

                Try
                    If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject AndAlso exportAsmAsPart Then
                        Dim tempPartPath As String = IO.Path.Combine(IO.Path.GetTempPath(),
                            "TEMP_" & System.Guid.NewGuid().ToString("N").Substring(0, 8) & ".ipt")

                        Dim templatePath As String = invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject)
                        Dim tempPart As PartDocument = CType(invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, templatePath, True), PartDocument)
                        tempPart.SaveAs(tempPartPath, False)

                        Dim derivedDef As DerivedAssemblyDefinition =
                            tempPart.ComponentDefinition.ReferenceComponents.DerivedAssemblyComponents.CreateDefinition(srcPath)
                        derivedDef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
                        tempPart.ComponentDefinition.ReferenceComponents.DerivedAssemblyComponents.Add(derivedDef)
                        tempPart.Update()
                        tempPart.Save2(True)

                        exportDoc = tempPart
                        isTemp = True
                        tempFiles.Add(tempPartPath)
                    End If

                    Dim stepFullPath As String
                    If useFullSaveDialog Then
                        stepFullPath = singleSavePath
                    Else
                        Dim baseName As String = IO.Path.GetFileNameWithoutExtension(srcPath)
                        If isTemp Then baseName &= "_PART"
                        stepFullPath = IO.Path.Combine(outputFolder, baseName & ".stp")
                    End If

                    If IO.File.Exists(stepFullPath) AndAlso Not useFullSaveDialog Then
                        Dim ans As DialogResult = MessageBox.Show("File đã tồn tại:" & vbCrLf & stepFullPath & vbCrLf & "Ghi đè?",
                                                                  "Xác nhận", MessageBoxButtons.YesNo)
                        If ans = DialogResult.No Then
                            failCount += 1
                            Continue For
                        End If
                    End If

                    Dim oContext As TranslationContext = invApp.TransientObjects.CreateTranslationContext()
                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                    Dim oOptions As NameValueMap = invApp.TransientObjects.CreateNameValueMap()
                    If stepTranslator.HasSaveCopyAsOptions(exportDoc, oContext, oOptions) Then
                        oOptions.Value("ApplicationProtocolType") = 3
                    End If

                    Dim oData As DataMedium = invApp.TransientObjects.CreateDataMedium()
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

            MessageBox.Show("EXPORT STEP AP214 HOÀN TẤT!" & vbCrLf & vbCrLf &
                            "Thành công : " & successCount & vbCrLf &
                            "Thất bại   : " & failCount & vbCrLf & vbCrLf &
                            "Folder     : " & outputFolder, "Kết quả")
        End Sub

    End Module
End Namespace
