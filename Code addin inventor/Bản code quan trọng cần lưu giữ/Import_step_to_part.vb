Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.OLD
    Public Module OLD1 'Import_step_to_part '|bản gốc| bản lưu trữ không liên kết vào trong code
        Public Sub OnExecute(ByVal Context As NameValueMap)

            ' Lấy ứng dụng Inventor đang chạy
            Dim invApp As Inventor.Application = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)

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

            ' 3. Chọn nhiều file STEP
            Dim stepFiles As New List(Of String)
            Try
                Dim stepDlg As Inventor.FileDialog = Nothing
                invApp.CreateFileDialog(stepDlg)

                stepDlg.DialogTitle = "CHỌN NHIỀU FILE STEP"
                stepDlg.Filter = "ALL FILE" '"STEP Files (*.step;*.stp)|*.step;*.stp"
                stepDlg.MultiSelectEnabled = True

                If lastFolder <> "" AndAlso IO.Directory.Exists(lastFolder) Then
                    stepDlg.InitialDirectory = lastFolder
                End If

                stepDlg.ShowOpen()

                If stepDlg.FileName = "" Then Exit Sub

                Dim selected() As String = stepDlg.FileName.Split("|"c)
                For Each f As String In selected
                    If IO.File.Exists(f) Then stepFiles.Add(f)
                Next

                If stepFiles.Count = 0 Then
                    MessageBox.Show("Không có file STEP hợp lệ.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                lastFolder = IO.Path.GetDirectoryName(stepFiles(0))
            Catch ex As Exception
                MessageBox.Show("Lỗi chọn file STEP: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
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


    End Module
End Namespace
