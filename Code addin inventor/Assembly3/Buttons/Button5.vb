
Imports Inventor
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.IO

Namespace ToolInventor2025.Assembly3.Buttons

    Public Module Button5

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                NumberAssemblyItem1()

            Catch ex As Exception
                MessageBox.Show(
                    "LỖI NumberAssemblyItem1:" & vbCrLf & vbCrLf &
                    ex.Message & vbCrLf & vbCrLf &
                    "Stack:" & vbCrLf & ex.StackTrace,
                    "Assembly3 - ERROR",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End Try

        End Sub


        '==========================================================
        ' MAIN
        '==========================================================
        Public Sub NumberAssemblyItem1()

            Dim invApp As Inventor.Application = g_inventorApplication

            If invApp Is Nothing Then
                MessageBox.Show("Không lấy được Inventor Application.", "Assembly3")
                Return
            End If

            Dim oAsm As AssemblyDocument =
                TryCast(invApp.ActiveDocument, AssemblyDocument)

            If oAsm Is Nothing Then
                MessageBox.Show(
                    "Active document không phải Assembly.",
                    "Assembly3 - item1",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
                Return
            End If


            Dim oDef As AssemblyComponentDefinition =
                oAsm.ComponentDefinition

            Dim oBOM As BOM =
                oDef.BOM


            '======================================================
            ' ENABLE STRUCTURED BOM
            '======================================================
            oBOM.StructuredViewEnabled = True
            oBOM.StructuredViewFirstLevelOnly = False


            Dim oBOMView As BOMView = Nothing

            Try
                oBOMView = oBOM.BOMViews.Item("Structured")
            Catch ex As Exception

                MessageBox.Show(
                    "Không lấy được Structured BOM." & vbCrLf &
                    ex.Message,
                    "Assembly3 - BOM ERROR",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

                Return

            End Try


            '======================================================
            ' PREFIX
            '======================================================
            Dim prefix As String =
                InputBox(
                    "Nhập prefix cho STT item1 (ví dụ: TH, để trống nếu không dùng):",
                    "Prefix STT item1",
                    "")

            prefix = prefix.Trim()


            '======================================================
            ' COLLECT ALL DOCUMENTS
            '======================================================
            Dim allDocs As New HashSet(Of Document)


            For Each occ As ComponentOccurrence In oDef.Occurrences

                CollectAllDocs(occ, allDocs)

            Next


            ' Không đánh chính Assembly cha
            allDocs.Remove(oAsm)


            '======================================================
            ' SORT TOP LEVEL BOM
            '======================================================
            Dim topRows As List(Of BOMRow) = Nothing

            Try

                topRows = SortRows(oBOMView.BOMRows)

            Catch ex As Exception

                MessageBox.Show(
                    "Lỗi trong Function SortRows:" & vbCrLf & vbCrLf &
                    ex.Message & vbCrLf & vbCrLf &
                    ex.StackTrace,
                    "Assembly3 - SortRows ERROR",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

                Return

            End Try


            If topRows Is Nothing Then

                MessageBox.Show(
                    "SortRows trả về Nothing.",
                    "Assembly3")

                Return

            End If


            '======================================================
            ' PART NUMBER + REVISION -> DOCUMENTS
            '======================================================
            Dim partKeyToDocs As New Dictionary(Of String, List(Of Document))


            For Each refDoc As Document In allDocs

                If refDoc Is Nothing Then Continue For


                Dim partNum As String =
                    GetPartNumberFromDoc(refDoc).Trim().ToUpper()


                If partNum = "" Then
                    partNum = GetFallbackName(refDoc)
                End If


                Dim rev As String =
                    GetRevisionFromDoc(refDoc).Trim().ToUpper()


                If rev = "" Then
                    rev = "?"
                End If


                Dim partKey As String =
                    partNum & "|" & rev


                If Not partKeyToDocs.ContainsKey(partKey) Then

                    partKeyToDocs.Add(
                        partKey,
                        New List(Of Document))

                End If


                If Not partKeyToDocs(partKey).Contains(refDoc) Then

                    partKeyToDocs(partKey).Add(refDoc)

                End If

            Next


            '======================================================
            ' NUMBER BOM
            '======================================================
            Dim partKeyToSTT As New Dictionary(Of String, String)


            Dim sttCounter As Integer = 1


            For Each row As BOMRow In topRows

                If row Is Nothing Then Continue For


                If row.ComponentDefinitions Is Nothing Then
                    Continue For
                End If


                If row.ComponentDefinitions.Count = 0 Then
                    Continue For
                End If


                Dim refDoc As Document = Nothing


                Try

                    refDoc =
                        row.ComponentDefinitions.Item(1).Document

                Catch

                    Continue For

                End Try


                If refDoc Is Nothing Then Continue For


                ' Bỏ Assembly chính
                If String.Compare(
                    refDoc.FullFileName,
                    oAsm.FullFileName,
                    True) = 0 Then

                    Continue For

                End If


                Dim partNum As String =
                    GetPartNumberFromDoc(refDoc).Trim().ToUpper()


                If partNum = "" Then
                    partNum = GetFallbackName(refDoc)
                End If


                Dim rev As String =
                    GetRevisionFromDoc(refDoc).Trim().ToUpper()


                If rev = "" Then
                    rev = "?"
                End If


                Dim partKey As String =
                    partNum & "|" & rev


                Dim numericSTT As String


                If partKeyToSTT.ContainsKey(partKey) Then

                    numericSTT =
                        partKeyToSTT(partKey)

                Else

                    numericSTT =
                        CStr(sttCounter)

                    partKeyToSTT.Add(
                        partKey,
                        numericSTT)

                    sttCounter += 1

                End If


                ' ItemNumber KHÔNG prefix
                Try

                    row.ItemNumber = numericSTT

                Catch ex As Exception

                    ' Không dừng chương trình
                    ' vì ItemNumber có thể bị Inventor khóa

                End Try

            Next


            '======================================================
            ' WRITE item1
            '======================================================
            Dim updateOK As Integer = 0
            Dim updateFail As Integer = 0


            For Each kvp As KeyValuePair(Of String, String)
                In partKeyToSTT


                Dim key As String =
                    kvp.Key


                Dim numericSTT As String =
                    kvp.Value


                Dim fullSTT As String =
                    numericSTT


                If prefix <> "" Then
                    fullSTT = prefix & fullSTT
                End If


                If partKeyToDocs.ContainsKey(key) Then


                    For Each doc As Document
                        In partKeyToDocs(key)


                        If AddOrUpdateSTT(
                            doc,
                            fullSTT) Then

                            updateOK += 1

                        Else

                            updateFail += 1

                        End If

                    Next

                End If

            Next


            '======================================================
            ' UPDATE / SAVE
            '======================================================
            Try
                oAsm.Update2(True)
            Catch
            End Try


            '======================================================
            ' RESULT
            '======================================================
            MessageBox.Show(
                "HOÀN TẤT" & vbCrLf & vbCrLf &
                "Số nhóm STT: " & partKeyToSTT.Count & vbCrLf &
                "Số Document cập nhật: " & updateOK & vbCrLf &
                "Số Document không cập nhật: " & updateFail & vbCrLf &
                "Prefix item1: '" & prefix & "'",
                "Assembly3 - item1",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)

        End Sub



        '==========================================================
        ' COLLECT DOCUMENTS
        '==========================================================
        Public Sub CollectAllDocs(
            occ As ComponentOccurrence,
            ByRef docsSet As HashSet(Of Document))


            If occ Is Nothing Then Return


            Try

                Dim doc As Document =
                    occ.Definition.Document


                If doc IsNot Nothing Then

                    docsSet.Add(doc)

                End If


                If occ.Definition.Type =
                    ObjectTypeEnum.kAssemblyComponentDefinitionObject Then


                    Dim subDef As AssemblyComponentDefinition =
                        TryCast(
                            occ.Definition,
                            AssemblyComponentDefinition)


                    If subDef IsNot Nothing Then


                        For Each subOcc As ComponentOccurrence
                            In subDef.Occurrences


                            CollectAllDocs(
                                subOcc,
                                docsSet)

                        Next

                    End If

                End If


            Catch ex As Exception

                ' Không làm chết toàn bộ rule
                ' Component bị lỗi sẽ được bỏ qua

            End Try

        End Sub



        '==========================================================
        ' SORT ROWS
        '
        ' Normal Assembly
        ' Normal Part
        ' Purchased Assembly
        ' Purchased Part
        ' Phantom Assembly
        ' Phantom Part
        ' Reference
        '==========================================================
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


                If row.ComponentDefinitions Is Nothing OrElse
                   row.ComponentDefinitions.Count = 0 Then


                    If row.BOMStructure =
                        BOMStructureEnum.kReferenceBOMStructure Then

                        reference.Add(row)

                    End If


                    Continue For

                End If


                Dim refDoc As Document = Nothing


                Try

                    refDoc =
                        row.ComponentDefinitions.Item(1).Document

                Catch

                    Continue For

                End Try


                If refDoc Is Nothing Then Continue For


                '================================================
                ' REFERENCE
                '================================================
                If row.BOMStructure =
                    BOMStructureEnum.kReferenceBOMStructure Then


                    reference.Add(row)


                    '================================================
                    ' PHANTOM
                    '================================================
                ElseIf row.BOMStructure =
                    BOMStructureEnum.kPhantomBOMStructure Then


                    If refDoc.DocumentType =
                        DocumentTypeEnum.kAssemblyDocumentObject Then


                        Dim mass As Double = 0


                        Try
                            mass =
                                refDoc.ComponentDefinition.
                                MassProperties.Mass
                        Catch
                            mass = 0
                        End Try


                        phantomAsmMassList.Add(
                            Tuple.Create(row, mass))


                    ElseIf refDoc.DocumentType =
                        DocumentTypeEnum.kPartDocumentObject Then


                        Dim mass As Double = 0


                        Try
                            mass =
                                refDoc.ComponentDefinition.
                                MassProperties.Mass
                        Catch
                            mass = 0
                        End Try


                        phantomPartMassList.Add(
                            Tuple.Create(row, mass))

                    End If


                    '================================================
                    ' PURCHASED
                    '================================================
                ElseIf row.BOMStructure =
                    BOMStructureEnum.kPurchasedBOMStructure Then


                    If refDoc.DocumentType =
                        DocumentTypeEnum.kAssemblyDocumentObject Then


                        purchasedAsm.Add(row)


                    ElseIf refDoc.DocumentType =
                        DocumentTypeEnum.kPartDocumentObject Then


                        purchasedPart.Add(row)

                    End If


                    '================================================
                    ' NORMAL / INSEPARABLE
                    '================================================
                Else


                    If refDoc.DocumentType =
                        DocumentTypeEnum.kAssemblyDocumentObject Then


                        Dim mass As Double = 0


                        Try
                            mass =
                                refDoc.ComponentDefinition.
                                MassProperties.Mass
                        Catch
                            mass = 0
                        End Try


                        subAsmMassList.Add(
                            Tuple.Create(row, mass))


                    ElseIf refDoc.DocumentType =
                        DocumentTypeEnum.kPartDocumentObject Then


                        Dim mass As Double = 0


                        Try
                            mass =
                                refDoc.ComponentDefinition.
                                MassProperties.Mass
                        Catch
                            mass = 0
                        End Try


                        partMassList.Add(
                            Tuple.Create(row, mass))

                    End If

                End If

            Next



            '======================================================
            ' SORT
            '======================================================
            subAsmMassList.Sort(
                Function(a, b)
                    Return b.Item2.CompareTo(a.Item2)
                End Function)


            partMassList.Sort(
                Function(a, b)
                    Return b.Item2.CompareTo(a.Item2)
                End Function)


            phantomAsmMassList.Sort(
                Function(a, b)
                    Return b.Item2.CompareTo(a.Item2)
                End Function)


            phantomPartMassList.Sort(
                Function(a, b)
                    Return b.Item2.CompareTo(a.Item2)
                End Function)


            purchasedAsm.Sort(
                Function(a, b)
                    Return String.Compare(
                        GetPartNumber(a),
                        GetPartNumber(b),
                        True)
                End Function)


            purchasedPart.Sort(
                Function(a, b)
                    Return String.Compare(
                        GetPartNumber(a),
                        GetPartNumber(b),
                        True)
                End Function)


            reference.Sort(
                Function(a, b)
                    Return String.Compare(
                        GetPartNumber(a),
                        GetPartNumber(b),
                        True)
                End Function)



            '======================================================
            ' CREATE ORDERED LIST
            '======================================================
            Dim orderedRows As New List(Of BOMRow)


            For Each item In subAsmMassList
                orderedRows.Add(item.Item1)
            Next


            For Each item In partMassList
                orderedRows.Add(item.Item1)
            Next


            For Each item In purchasedAsm
                orderedRows.Add(item)
            Next


            For Each item In purchasedPart
                orderedRows.Add(item)
            Next


            For Each item In phantomAsmMassList
                orderedRows.Add(item.Item1)
            Next


            For Each item In phantomPartMassList
                orderedRows.Add(item.Item1)
            Next


            orderedRows.AddRange(reference)


            Return orderedRows

        End Function



        '==========================================================
        ' GET PART NUMBER FROM BOM ROW
        '==========================================================
        Public Function GetPartNumber(
            row As BOMRow) As String


            Try

                If row Is Nothing Then Return ""


                If row.ComponentDefinitions Is Nothing Then
                    Return ""
                End If


                If row.ComponentDefinitions.Count = 0 Then
                    Return ""
                End If


                Dim refDoc As Document = Nothing


                Try

                    refDoc =
                        row.ComponentDefinitions.Item(1).Document

                Catch

                    Return ""

                End Try


                If refDoc Is Nothing Then Return ""


                Return GetPartNumberFromDoc(refDoc)


            Catch

                Return ""

            End Try

        End Function



        '==========================================================
        ' GET PART NUMBER
        '==========================================================
        Public Function GetPartNumberFromDoc(
            doc As Document) As String


            Try

                If doc Is Nothing Then Return ""


                Return CStr(
                    doc.PropertySets.
                    Item("Design Tracking Properties").
                    Item("Part Number").
                    Value)


            Catch

                Return ""

            End Try

        End Function



        '==========================================================
        ' GET REVISION
        '==========================================================
        Public Function GetRevisionFromDoc(
            doc As Document) As String


            Try

                If doc Is Nothing Then Return ""


                Return CStr(
                    doc.PropertySets.
                    Item("Design Tracking Properties").
                    Item("Revision Number").
                    Value)


            Catch

                Return ""

            End Try

        End Function



        '==========================================================
        ' FALLBACK NAME
        '==========================================================
        Public Function GetFallbackName(
            doc As Document) As String


            Try

                If doc Is Nothing Then Return "?"


                If doc.DisplayName <> "" Then
                    Return doc.DisplayName.
                        Trim().
                        ToUpper()
                End If


                Return "?"


            Catch

                Return "?"

            End Try

        End Function



        '==========================================================
        ' WRITE item1
        '
        ' QUAN TRỌNG:
        ' Inventor User Defined Properties.Add
        ' = Add(Value, Name)
        '
        ' Không phải:
        ' Add(Name, Value)
        '==========================================================
        Public Function AddOrUpdateSTT(
            doc As Document,
            value As String) As Boolean


            Try

                If doc Is Nothing Then
                    Return False
                End If


                If Not doc.IsModifiable Then
                    Return False
                End If


                Dim ps As PropertySet = Nothing


                Try

                    ps =
                        doc.PropertySets.
                        Item("Inventor User Defined Properties")

                Catch

                    Return False

                End Try


                If ps Is Nothing Then
                    Return False
                End If



                Dim prop As Inventor.Property = Nothing


                '==================================================
                ' TÌM item1
                '==================================================
                For Each p As Inventor.Property In ps


                    If String.Equals(
                        p.Name,
                        "item1",
                        StringComparison.OrdinalIgnoreCase) Then


                        prop = p

                        Exit For

                    End If

                Next



                '==================================================
                ' NẾU CHƯA CÓ -> TẠO
                '
                ' ĐÚNG:
                ' ps.Add(value, "item1")
                '==================================================
                If prop Is Nothing Then


                    prop =
                        ps.Add(
                            value,
                            "item1")


                Else

                    '================================================
                    ' ĐÃ CÓ -> UPDATE
                    '================================================
                    prop.Value = value

                End If



                '==================================================
                ' UPDATE DOCUMENT
                '==================================================
                Try
                    doc.Update()
                Catch
                End Try


                '==================================================
                ' SAVE
                '==================================================
                Try

                    If Not doc.ReadOnly Then
                        doc.Save()
                    End If

                Catch
                    ' Không coi lỗi Save là lỗi tạo property
                End Try


                Return True


            Catch ex As Exception

                Return False

            End Try

        End Function


    End Module

End Namespace
