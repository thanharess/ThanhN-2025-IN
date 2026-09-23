Imports Inventor
Imports System.Windows.Forms
Imports Microsoft.VisualBasic

Namespace ToolInventor2025.Part.Buttons.solid
    Public Module Part_Solid_6
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                ' Ensure active document is a part
                Dim doc As Document = g_inventorApplication.ActiveDocument
                If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                    MessageBox.Show("Please open a Part document", "iLogic")
                    Return
                End If

                Dim partDoc As PartDocument = CType(doc, PartDocument)

                ' Access the Inventor User Defined Properties set (create if missing)
                Dim customPropertySet As PropertySet = Nothing
                Try
                    customPropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
                Catch
                    ' If the property set doesn't exist, create it is not straightforward via API; show error
                    MessageBox.Show("Cannot access 'Inventor User Defined Properties' on this document.", "Error")
                    Return
                End Try

                Dim prefixPropertyName As String = "Prefix"
                Dim prefixProperty As [Property] = Nothing
                Try
                    prefixProperty = customPropertySet.Item(prefixPropertyName)
                Catch
                    ' Add the property if it doesn't exist
                    prefixProperty = customPropertySet.Add("", prefixPropertyName)
                End Try

                ' If Prefix is empty, try to seed it from Part Number property
                Try
                    Dim partNumber As String = String.Empty
                    Try
                        partNumber = CStr(partDoc.PropertySets.Item("Project").Item("Part Number").Value)
                    Catch
                        Try
                            partNumber = CStr(partDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value)
                        Catch
                            partNumber = String.Empty
                        End Try
                    End Try

                    If String.IsNullOrEmpty(CStr(prefixProperty.Value)) AndAlso Not String.IsNullOrEmpty(partNumber) Then
                        prefixProperty.Value = partNumber & "_"
                    End If
                Catch
                    ' ignore
                End Try

                ' Get prefix from user
                Dim bodyPrefix As String = Interaction.InputBox("Enter a prefix for the solid body names", "iLogic", CStr(prefixProperty.Value))

                ' Save back to custom property
                Try
                    prefixProperty.Value = bodyPrefix
                Catch
                End Try

                ' Rename all solid bodies with incrementing suffix
                Dim idx As Integer = 1
                For Each solid As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies
                    Try
                        solid.Name = String.Format("{0}{1:00}", bodyPrefix, idx)
                    Catch
                    End Try
                    idx += 1
                Next

                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Executed Part Action 6")

            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Part Action 6: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
