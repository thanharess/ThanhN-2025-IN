Imports Inventor
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports System.Collections.Generic

Namespace ToolInventor2025.Part.Buttons.solid

    Public Module Part_Solid_5
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                ' Get active document and validate
                Dim doc As Document = g_inventorApplication.ActiveDocument
                If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                    MessageBox.Show("Rule chỉ dùng cho Part (*.ipt)", "Lỗi")
                    Return
                End If

                Dim oDoc As PartDocument = CType(doc, PartDocument)
                If String.IsNullOrEmpty(oDoc.FullFileName) Then
                    MessageBox.Show("Vui lòng Save file trước!", "Lỗi")
                    Return
                End If

                Dim oFolderDlg As New FolderBrowserDialog()
                oFolderDlg.Description = "Chọn thư mục xuất file"
                If oFolderDlg.ShowDialog() <> DialogResult.OK Then Return
                Dim sExportFolder As String = oFolderDlg.SelectedPath

                Dim oFormatMap As New Dictionary(Of String, String) From {
                    {"STEP (*.stp)", "STEP|.stp"},
                    {"IGES (*.igs)", "IGES|.igs"},
                    {"SAT - ACIS (*.sat)", "ACIS|.sat"},
                    {"Parasolid (*.x_t)", "Parasolid|.x_t"},
                    {"JT (*.jt)", "JT|.jt"},
                    {"STL (*.stl)", "STL|.stl"},
                    {"OBJ (*.obj)", "OBJ|.obj"}
                }

                Dim aValidFormats As New List(Of String)()
                For Each kvp In oFormatMap
                    For Each addIn As ApplicationAddIn In g_inventorApplication.ApplicationAddIns
                        If addIn.DisplayName.Contains(kvp.Value.Split("|"c)(0)) Then
                            aValidFormats.Add(kvp.Key)
                            Exit For
                        End If
                    Next
                Next

                If aValidFormats.Count = 0 Then
                    MessageBox.Show("Không tìm thấy Translator nào!", "Lỗi")
                    Return
                End If

                ' Let user pick a format by entering the number
                Dim prompt As New System.Text.StringBuilder()
                prompt.AppendLine("Chọn định dạng xuất file (nhập số):")
                For i As Integer = 0 To aValidFormats.Count - 1
                    prompt.AppendFormat("{0}. {1}" & vbCrLf, i + 1, aValidFormats(i))
                Next

                Dim input As String = Interaction.InputBox(prompt.ToString(), "Export Format", "1")
                If String.IsNullOrEmpty(input) Then Return

                Dim idx As Integer
                If Not Integer.TryParse(input, idx) Then
                    MessageBox.Show("Giá trị không hợp lệ", "Lỗi")
                    Return
                End If
                If idx < 1 OrElse idx > aValidFormats.Count Then
                    MessageBox.Show("Giá trị không hợp lệ", "Lỗi")
                    Return
                End If

                Dim sChoice As String = aValidFormats(idx - 1)
                Dim parts() As String = oFormatMap(sChoice).Split("|"c)
                Dim sKey As String = parts(0)
                Dim sExt As String = parts(1)

                Dim oTrans As TranslatorAddIn = Nothing
                For Each addIn As ApplicationAddIn In g_inventorApplication.ApplicationAddIns
                    If addIn.DisplayName.Contains(sKey) Then
                        oTrans = addIn
                        Exit For
                    End If
                Next

                If oTrans Is Nothing Then
                    MessageBox.Show("Translator không tồn tại!", "Lỗi")
                    Return
                End If

                Dim oCompDef As PartComponentDefinition = oDoc.ComponentDefinition
                Dim oBodies As SurfaceBodies = oCompDef.SurfaceBodies

                Dim oContext As TranslationContext = g_inventorApplication.TransientObjects.CreateTranslationContext()
                oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                Dim oOptions As NameValueMap = g_inventorApplication.TransientObjects.CreateNameValueMap()
                Dim oData As DataMedium = g_inventorApplication.TransientObjects.CreateDataMedium()

                For i As Integer = 1 To oBodies.Count
                    Dim oNewDoc As PartDocument = g_inventorApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject, Nothing, False)

                    Dim oNewDef As PartComponentDefinition = oNewDoc.ComponentDefinition
                    Dim oDerives = oNewDef.ReferenceComponents.DerivedPartComponents
                    Dim oDef = oDerives.CreateDefinition(oDoc.FullFileName)

                    oDef.IncludeAllSolids = False
                    oDef.IncludeAllSurfaces = False
                    oDef.IncludeAllSketches = False
                    oDef.IncludeAllWorkFeatures = False

                    Dim oBody = oBodies.Item(i)

                    If oBody.IsSolid Then
                        For Each oEnt As DerivedPartEntity In oDef.Solids
                            If oEnt.ReferencedEntity Is oBody Then
                                oEnt.IncludeEntity = True
                                Exit For
                            End If
                        Next
                    Else
                        For Each oEnt As DerivedPartEntity In oDef.Surfaces
                            If oEnt.ReferencedEntity Is oBody Then
                                oEnt.IncludeEntity = True
                                Exit For
                            End If
                        Next
                    End If

                    oDerives.Add(oDef)

                    oData.FileName = System.IO.Path.Combine(sExportFolder, System.IO.Path.GetFileNameWithoutExtension(oDoc.FullFileName) & "_Body" & i & sExt)

                    oTrans.SaveCopyAs(oNewDoc, oContext, oOptions, oData)
                    oNewDoc.Close(True)
                Next

                System.Diagnostics.Process.Start("explorer.exe", sExportFolder)
                MessageBox.Show("Xuất file hoàn tất!", "Done")

            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Part Action 5: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
