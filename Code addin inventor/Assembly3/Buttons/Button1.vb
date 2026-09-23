Imports Inventor
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Collections.Generic

Namespace ToolInventor2025.Assembly3.Buttons
    Public Module Button1
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                ' Run BOM numbering converted from iLogic macro
                NumberAssemblyBOM()
            Catch ex As Exception
                MessageBox.Show("Error running NumberAssemblyBOM: " & ex.Message, "Assembly2 - BOM Numbering")
            End Try
        End Sub

        Private Sub NumberAssemblyBOM()
            Dim asmDoc As AssemblyDocument = TryCast(g_inventorApplication.ActiveDocument, AssemblyDocument)
            If asmDoc Is Nothing Then
                MessageBox.Show("Active document is not an assembly.", "Assembly2 - BOM Numbering")
                Return
            End If

            Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
            Dim bom As BOM = compDef.BOM
            bom.StructuredViewEnabled = True
            bom.StructuredViewFirstLevelOnly = False
            Dim bomView As BOMView = bom.BOMViews.Item("Structured")

            Dim prefix As String = Microsoft.VisualBasic.Interaction.InputBox("Nhập prefix cho STT Part Number (ví dụ: TH, để trống nếu không dùng):", "Prefix STT Part Number", "")
            prefix = prefix.Trim()

            ' Collect all unique Documents from the entire assembly tree (excluding main assembly)
            Dim allDocs As New HashSet(Of Document)()
            For Each occ As ComponentOccurrence In compDef.Occurrences
                CollectAllDocs(occ, allDocs)
            Next
            allDocs.Remove(asmDoc)

            ' Build partKey -> list of docs map
            Dim partKeyToDocs As New Dictionary(Of String, List(Of Document))(StringComparer.OrdinalIgnoreCase)
            For Each refDoc As Document In allDocs
                Dim partNum As String = GetPartNumberFromDoc(refDoc).Trim().ToUpper()
                If partNum = String.Empty Then partNum = GetFallbackName(refDoc)
                Dim rev As String = GetRevisionFromDoc(refDoc).Trim().ToUpper()
                If rev = String.Empty Then rev = "?"
                Dim partKey As String = partNum & "|" & rev
                If Not partKeyToDocs.ContainsKey(partKey) Then
                    partKeyToDocs.Add(partKey, New List(Of Document)())
                End If
                If Not partKeyToDocs(partKey).Contains(refDoc) Then
                    partKeyToDocs(partKey).Add(refDoc)
                End If
            Next

            ' Get top-level structured BOM rows (sorted)
            Dim topRows As List(Of BOMRow) = SortRows(bomView.BOMRows)

            ' Assign ItemNumber for top-level rows (numeric, no prefix on ItemNumber)
            Dim partKeyToSTT As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim sttCounter As Integer = 1
            For Each row As BOMRow In topRows
                If row Is Nothing Then Continue For
                If row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then Continue For
                Dim refDoc As Document = Nothing
                Try
                    refDoc = row.ComponentDefinitions.Item(1).Document
                Catch
                    Continue For
                End Try
                If refDoc Is Nothing Then Continue For
                If String.Compare(refDoc.FullFileName, asmDoc.FullFileName, True) = 0 Then Continue For

                Dim partNum As String = GetPartNumberFromDoc(refDoc).Trim().ToUpper()
                If partNum = String.Empty Then partNum = GetFallbackName(refDoc)
                Dim rev As String = GetRevisionFromDoc(refDoc).Trim().ToUpper()
                If rev = String.Empty Then rev = "?"

                Dim partKey As String = partNum & "|" & rev

                Dim numericSTT As String
                If partKeyToSTT.ContainsKey(partKey) Then
                    numericSTT = partKeyToSTT(partKey)
                Else
                    numericSTT = CStr(sttCounter)
                    partKeyToSTT.Add(partKey, numericSTT)
                    sttCounter += 1
                End If

                Dim sttValue As String = numericSTT
                Try
                    row.ItemNumber = sttValue
                Catch
                End Try
            Next

            ' Write Part Number for all documents in each group (apply prefix to Part Number)
            For Each kvp As KeyValuePair(Of String, String) In partKeyToSTT
                Dim key As String = kvp.Key
                Dim numericSTT As String = kvp.Value
                Dim fullSTT As String = numericSTT
                If prefix <> String.Empty Then fullSTT = prefix & fullSTT
                If partKeyToDocs.ContainsKey(key) Then
                    For Each doc As Document In partKeyToDocs(key)
                        SetPartNumber(doc, fullSTT)
                    Next
                End If
            Next

            MessageBox.Show("Hoàn tất: Structured BOM (top-level only, không prefix) và Model Data (Part Number với prefix '" & prefix & "') đã được đánh STT.", "iLogic - Converted")
        End Sub

        Private Sub CollectAllDocs(occ As ComponentOccurrence, ByRef docsSet As HashSet(Of Document))
            Try
                If occ Is Nothing OrElse occ.Definition Is Nothing Then Return
                docsSet.Add(occ.Definition.Document)
                If occ.Definition.Type = ObjectTypeEnum.kAssemblyComponentDefinitionObject Then
                    Dim subDef As AssemblyComponentDefinition = TryCast(occ.Definition, AssemblyComponentDefinition)
                    If subDef IsNot Nothing Then
                        For Each subOcc As ComponentOccurrence In subDef.Occurrences
                            CollectAllDocs(subOcc, docsSet)
                        Next
                    End If
                End If
            Catch
                ' ignore
            End Try
        End Sub

        Private Function SortRows(bomRows As BOMRowsEnumerator) As List(Of BOMRow)
            Dim topRows As New List(Of BOMRow)()
            For Each row As BOMRow In bomRows
                If row Is Nothing Then Continue For
                ' Collect only top-level rows. Use Level property if available (level 1 = top-level).
                Try
                    If row.Level = 1 Then
                        topRows.Add(row)
                    End If
                Catch
                    ' If Level isn't available for some Inventor API version, fall back to include all rows
                    topRows.Add(row)
                End Try
            Next
            Return topRows
        End Function

        Private Function GetPartNumberFromDoc(doc As Document) As String
            Try
                Dim ps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim prop As [Property] = ps.Item("Part Number")
                If prop IsNot Nothing AndAlso prop.Value IsNot Nothing Then
                    Return Convert.ToString(prop.Value)
                End If
            Catch
            End Try
            Return String.Empty
        End Function

        Private Function GetRevisionFromDoc(doc As Document) As String
            Try
                Dim ps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim prop As [Property] = ps.Item("Revision")
                If prop Is Nothing Then
                    ' some Inventor versions use "Revision Number"
                    prop = ps.Item("Revision Number")
                End If
                If prop IsNot Nothing AndAlso prop.Value IsNot Nothing Then
                    Return Convert.ToString(prop.Value)
                End If
            Catch
            End Try
            Return String.Empty
        End Function

        Private Function GetFallbackName(doc As Document) As String
            Try
                Return System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName)
            Catch
            End Try
            Return ""
        End Function

        Private Sub SetPartNumber(doc As Document, value As String)
            Try
                Dim ps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim prop As [Property] = Nothing
                Try
                    prop = ps.Item("Part Number")
                Catch
                End Try
                If prop Is Nothing Then
                    ' try to add the property if possible
                Else
                    prop.Value = value
                End If
                ' Save/Update the document to persist the iProperty change
                Try
                    doc.Update()
                    ' If the document is not the active editable document, save it
                    If doc.ReadOnly = False Then
                        doc.Save()
                    End If
                Catch
                End Try
            Catch
            End Try
        End Sub

    End Module
End Namespace
