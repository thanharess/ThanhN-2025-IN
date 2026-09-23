Imports Inventor
Imports System.Windows.Forms
Imports System.Collections.Generic

Namespace ToolInventor2025.Assembly3.Buttons
    Public Module Button4
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                NumberAssemblyPurchasedPartNumber()
            Catch ex As Exception
                MessageBox.Show("Error running NumberAssemblyPurchasedPartNumber: " & ex.Message, "Assembly3 - Purchased PartNumber")
            End Try
        End Sub

        Private Sub NumberAssemblyPurchasedPartNumber()
            Dim asmDoc As AssemblyDocument = TryCast(g_inventorApplication.ActiveDocument, AssemblyDocument)
            If asmDoc Is Nothing Then
                MessageBox.Show("Active document is not an assembly.", "Assembly3 - Purchased PartNumber")
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

            ' Get top-level structured BOM rows (sorted using purchased-last ordering)
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

            MessageBox.Show("Hoàn tất: Structured BOM (top-level only, purchased items sorted lower) và Model Data (Part Number với prefix '" & prefix & "') đã được đánh STT.", "iLogic - Converted")
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
            Dim subAsmMassList As New List(Of Tuple(Of BOMRow, Double))
            Dim partMassList As New List(Of Tuple(Of BOMRow, Double))
            Dim purchasedAsm As New List(Of Tuple(Of BOMRow, String))
            Dim purchasedPart As New List(Of Tuple(Of BOMRow, String))
            Dim phantomAsmMassList As New List(Of Tuple(Of BOMRow, Double))
            Dim phantomPartMassList As New List(Of Tuple(Of BOMRow, Double))
            Dim reference As New List(Of Tuple(Of BOMRow, String))

            For Each Row As BOMRow In bomRows
                If Row Is Nothing Then Continue For
                If Row.ComponentDefinitions Is Nothing OrElse Row.ComponentDefinitions.Count = 0 Then
                    If Row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                        reference.Add(Tuple.Create(Row, ""))
                    End If
                    Continue For
                End If
                Dim refDoc As Document = Nothing
                Try
                    refDoc = Row.ComponentDefinitions.Item(1).Document
                Catch
                    Continue For
                End Try
                If refDoc Is Nothing Then Continue For

                Dim partNum As String = GetPartNumberFromDoc(refDoc).Trim().ToUpper()
                If partNum = String.Empty Then partNum = "?"

                If Row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                    reference.Add(Tuple.Create(Row, partNum))
                ElseIf Row.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                    Dim mass As Double = GetMassOfDocument(refDoc)
                    If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        phantomAsmMassList.Add(Tuple.Create(Row, mass))
                    Else
                        phantomPartMassList.Add(Tuple.Create(Row, mass))
                    End If
                ElseIf Row.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                    ' purchased assemblies and parts sorted by Part Number ascending
                    If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        purchasedAsm.Add(Tuple.Create(Row, partNum))
                    Else
                        purchasedPart.Add(Tuple.Create(Row, partNum))
                    End If
                Else
                    ' Normal (non-phantom, non-purchased, non-reference)
                    Dim mass As Double = GetMassOfDocument(refDoc)
                    If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        subAsmMassList.Add(Tuple.Create(Row, mass))
                    Else
                        partMassList.Add(Tuple.Create(Row, mass))
                    End If
                End If
            Next

            ' Sort groups
            subAsmMassList.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            partMassList.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            purchasedAsm.Sort(Function(a, b) String.Compare(a.Item2, b.Item2, StringComparison.OrdinalIgnoreCase))
            purchasedPart.Sort(Function(a, b) String.Compare(a.Item2, b.Item2, StringComparison.OrdinalIgnoreCase))
            phantomAsmMassList.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            phantomPartMassList.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            reference.Sort(Function(a, b) String.Compare(a.Item2, b.Item2, StringComparison.OrdinalIgnoreCase))

            ' Combine in required order: normal assemblies, normal parts, purchased assemblies, purchased parts, phantom assemblies, phantom parts, reference
            Dim result As New List(Of BOMRow)()
            For Each t In subAsmMassList
                result.Add(t.Item1)
            Next
            For Each t In partMassList
                result.Add(t.Item1)
            Next
            For Each t In purchasedAsm
                result.Add(t.Item1)
            Next
            For Each t In purchasedPart
                result.Add(t.Item1)
            Next
            For Each t In phantomAsmMassList
                result.Add(t.Item1)
            Next
            For Each t In phantomPartMassList
                result.Add(t.Item1)
            Next
            For Each t In reference
                result.Add(t.Item1)
            Next

            Return result
        End Function

        Private Function GetMassOfDocument(doc As Document) As Double
            Try
                Dim partDoc As PartDocument = TryCast(doc, PartDocument)
                If partDoc IsNot Nothing Then
                    Try
                        Return partDoc.ComponentDefinition.MassProperties.Mass
                    Catch
                    End Try
                End If
                Dim asmDoc As AssemblyDocument = TryCast(doc, AssemblyDocument)
                If asmDoc IsNot Nothing Then
                    Try
                        Return asmDoc.ComponentDefinition.MassProperties.Mass
                    Catch
                    End Try
                End If
            Catch
            End Try
            Return 0.0
        End Function

        Private Function GetPartNumberFromDoc(doc As Document) As String
            Try
                Dim ps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim prop As [Property] = Nothing
                Try
                    prop = ps.Item("Part Number")
                Catch
                End Try
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
                Dim prop As [Property] = Nothing
                Try
                    prop = ps.Item("Revision")
                Catch
                End Try
                If prop Is Nothing Then
                    Try
                        prop = ps.Item("Revision Number")
                    Catch
                    End Try
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
            Return String.Empty
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
                    Try
                        ps.Add("Part Number", value)
                    Catch
                        ' ignore add failures
                    End Try
                Else
                    prop.Value = value
                End If

                Try
                    doc.Update()
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
