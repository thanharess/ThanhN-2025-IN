Option Explicit On
Option Strict Off
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports IO = System.IO
Imports ToolInventor2025   ' CommonKeywords

Namespace ToolInventor2025.Assembly.Buttons.caclenhboctach.part
    Public Module Ass_boctach_part_1h

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim oApp As Inventor.Application = g_inventorApplication
            Dim selectedFile As String = ""
            Dim activeDoc As Document = Nothing

            Try
                Try
                    activeDoc = oApp.ActiveDocument
                Catch
                    activeDoc = Nothing
                End Try

                If activeDoc IsNot Nothing AndAlso activeDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim result As DialogResult = MessageBox.Show("Sử dụng file lắp ghép đang mở?", "Chọn file", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If result = DialogResult.Yes Then
                        selectedFile = activeDoc.FullFileName
                    Else
                        selectedFile = SelectAssemblyFile()
                    End If
                Else
                    selectedFile = SelectAssemblyFile()
                End If

                If String.IsNullOrWhiteSpace(selectedFile) Then Exit Sub

                Dim oOrigAsmDoc As AssemblyDocument = Nothing
                Dim openedByCode As Boolean = False

                If activeDoc IsNot Nothing AndAlso String.Equals(selectedFile, activeDoc.FullFileName, StringComparison.OrdinalIgnoreCase) Then
                    oOrigAsmDoc = CType(activeDoc, AssemblyDocument)
                Else
                    oOrigAsmDoc = CType(oApp.Documents.Open(selectedFile, False), AssemblyDocument)
                    openedByCode = True
                End If

                ' Chọn loại chi tiết
                Dim partType As Integer = SimpleSelectPartType()
                If partType = 0 Then
                    If openedByCode Then oOrigAsmDoc.Close(False)
                    Exit Sub
                End If

                ' Tạo Assembly đích
                Dim targetDoc As AssemblyDocument = CType(oApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject), AssemblyDocument)
                Dim baseName As String = IO.Path.GetFileNameWithoutExtension(selectedFile)
                Dim displayName As String = baseName & "_Aggregated_Parts"
                targetDoc.DisplayName = displayName

                '=====================================================
                ' THU THẬP PARTS THEO LOẠI ĐÃ CHỌN
                ' DÙNG RECURSIVE COLLECTION ĐỂ LẤY HẾT
                '=====================================================

                Dim collectedParts As New List(Of String)
                Dim visitedAssemblies As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                ' Thu thập parts theo loại đã chọn
                CollectPartsByType(oOrigAsmDoc.ComponentDefinition, partType, collectedParts, visitedAssemblies)

                If collectedParts.Count = 0 Then
                    MessageBox.Show("Không tìm thấy chi tiết nào phù hợp.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    targetDoc.Close(True)
                    If openedByCode Then oOrigAsmDoc.Close(False)
                    Exit Sub
                End If

                Dim oTG As TransientGeometry = oApp.TransientGeometry
                Dim oMatrix As Matrix = oTG.CreateMatrix()
                Dim added As Integer = 0

                ' Thêm tất cả parts đã thu thập vào assembly mới
                For Each partFullName As String In collectedParts
                    Try
                        targetDoc.ComponentDefinition.Occurrences.Add(partFullName, oMatrix)
                        added += 1
                    Catch ex As Exception
                        ' Bỏ qua part lỗi
                    End Try
                Next

                '==========================================================
                ' UPDATE ASSEMBLY ĐÍCH + HIỆN ITEM QTY TRONG BOM
                '==========================================================
                Try
                    targetDoc.Update2(True)
                Catch
                End Try

                Try
                    Dim targetBOM As BOM = targetDoc.ComponentDefinition.BOM

                    '------------------------------------------------------
                    ' BẬT PARTS ONLY
                    '------------------------------------------------------
                    Try
                        targetBOM.PartsOnlyViewEnabled = True
                    Catch
                    End Try

                    Try
                        targetBOM.Update()
                    Catch
                    End Try

                    '------------------------------------------------------
                    ' LẤY PARTS ONLY VIEW
                    '------------------------------------------------------
                    Dim targetPartsView As BOMView = Nothing

                    Try
                        targetPartsView = targetBOM.BOMViews.Item("Parts Only")
                    Catch
                        targetPartsView = Nothing
                    End Try

                    '------------------------------------------------------
                    ' HIỆN CỘT QTY
                    '------------------------------------------------------
                    If targetPartsView IsNot Nothing Then

                        Try
                            Dim cols As Object = targetPartsView.BOMColumns

                            For i As Integer = 1 To cols.Count

                                Dim col As Object = Nothing

                                Try
                                    col = cols.Item(i)
                                Catch
                                    col = Nothing
                                End Try

                                If col Is Nothing Then Continue For

                                Dim title As String = ""
                                Dim propName As String = ""

                                Try
                                    title = CStr(col.Title)
                                Catch
                                    title = ""
                                End Try

                                Try
                                    propName = CStr(col.PropertyName)
                                Catch
                                    propName = ""
                                End Try

                                '------------------------------------------
                                ' QTY / ITEM QTY / QUANTITY
                                '------------------------------------------
                                If String.Equals(title.Trim(), "QTY", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(title.Trim(), "Item Qty", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(title.Trim(), "Quantity", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(propName.Trim(), "Quantity", StringComparison.OrdinalIgnoreCase) Then

                                    Try
                                        col.Visible = True
                                    Catch
                                    End Try

                                End If

                            Next

                        Catch
                        End Try

                    End If

                Catch
                    ' Không để lỗi BOM làm dừng chương trình
                End Try

                MessageBox.Show(
    "Done." & vbCrLf &
    "Added occurrences: " & added.ToString(),
    "Info",
    MessageBoxButtons.OK,
    MessageBoxIcon.Information
)

                If openedByCode Then oOrigAsmDoc.Close(False)

            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        '=====================================================
        ' RECURSIVE COLLECT PARTS BY TYPE
        '=====================================================
        Private Sub CollectPartsByType(
            ByVal asmDef As AssemblyComponentDefinition,
            ByVal partType As Integer,
            ByRef collectedParts As List(Of String),
            ByRef visitedAssemblies As HashSet(Of String))

            For Each oOcc As ComponentOccurrence In asmDef.Occurrences
                Try
                    ' Bỏ qua suppressed
                    If oOcc.Suppressed Then Continue For

                    ' Lấy document
                    Dim oRefDoc As Document = Nothing
                    Try
                        If oOcc.ReferencedDocumentDescriptor IsNot Nothing Then
                            oRefDoc = oOcc.ReferencedDocumentDescriptor.ReferencedDocument
                        Else
                            oRefDoc = oOcc.Definition.Document
                        End If
                    Catch
                        oRefDoc = Nothing
                    End Try

                    If oRefDoc Is Nothing Then Continue For

                    Dim refFullName As String = String.Empty
                    Try
                        refFullName = oRefDoc.FullFileName
                    Catch
                        refFullName = String.Empty
                    End Try

                    If String.IsNullOrWhiteSpace(refFullName) Then Continue For

                    ' Kiểm tra PART
                    If oRefDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        Dim oPartDoc As PartDocument = TryCast(oRefDoc, PartDocument)
                        If oPartDoc Is Nothing Then Continue For

                        Dim include As Boolean = False

                        Select Case partType
                            Case 1   ' SHEET METAL
                                include = IsSheetMetalPart(oPartDoc)

                            Case 2   ' PURCHASED
                                ' Kiểm tra BOM structure của occurrence
                                Try
                                    include = (oOcc.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure)
                                Catch
                                    include = False
                                End Try

                            Case 3   ' STANDARD LIBRARY
                                Dim pn As String = GetProperty(oRefDoc, "Part Number")
                                include = IsLibraryPart(pn)

                            Case 4   ' PURCHASED + LIBRARY
                                Dim purchased As Boolean = False
                                Try
                                    purchased = (oOcc.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure)
                                Catch
                                    purchased = False
                                End Try

                                Dim pn2 As String = GetProperty(oRefDoc, "Part Number")
                                include = purchased OrElse IsLibraryPart(pn2)
                        End Select

                        If include Then
                            ' Thêm vào list (giữ duplicate)
                            collectedParts.Add(refFullName)
                        End If

                        ' Kiểm tra SUB ASSEMBLY
                    ElseIf oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Dim oSubAsmDoc As AssemblyDocument = TryCast(oRefDoc, AssemblyDocument)
                        If oSubAsmDoc IsNot Nothing Then
                            Dim oSubAsmDef As AssemblyComponentDefinition = oSubAsmDoc.ComponentDefinition
                            If oSubAsmDef IsNot Nothing Then
                                If Not visitedAssemblies.Contains(refFullName) Then
                                    visitedAssemblies.Add(refFullName)
                                    CollectPartsByType(oSubAsmDef, partType, collectedParts, visitedAssemblies)
                                End If
                            End If
                        End If
                    End If

                Catch ex As Exception
                    ' Bỏ qua occurrence lỗi
                End Try
            Next
        End Sub

        '=====================================================
        ' KIỂM TRA SHEET METAL PART CHÍNH XÁC
        '=====================================================
        Private Function IsSheetMetalPart(ByVal partDoc As PartDocument) As Boolean
            Try
                If partDoc Is Nothing Then Return False
                If partDoc.ComponentDefinition Is Nothing Then Return False

                ' Cách 1: Kiểm tra trực tiếp SheetMetalComponentDefinition
                Dim smDef As SheetMetalComponentDefinition = TryCast(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
                If smDef IsNot Nothing Then
                    ' Kiểm tra HasFlatPattern - property này chỉ có ở sheet metal
                    Try
                        If smDef.HasFlatPattern Then
                            Return True
                        End If
                    Catch
                    End Try

                    ' Kiểm tra số lượng sheet metal features
                    Try
                        If smDef.Features.Count > 0 Then
                            ' Kiểm tra thêm xem có feature sheet metal cụ thể không
                            For Each feat As PartFeature In smDef.Features
                                ' Các sheet metal features trong Inventor API
                                If TypeOf feat Is FaceFeature OrElse
                                   TypeOf feat Is FlangeFeature OrElse
                                   TypeOf feat Is ContourFlangeFeature OrElse
                                   TypeOf feat Is FoldFeature OrElse
                                   TypeOf feat Is BendFeature OrElse
                                   TypeOf feat Is CutFeature OrElse
                                   TypeOf feat Is CornerRoundFeature OrElse
                                   TypeOf feat Is CornerChamferFeature OrElse
                                   TypeOf feat Is PunchToolFeature OrElse
                                   TypeOf feat Is FlatPattern Then
                                    Return True
                                End If
                            Next
                        End If
                    Catch
                    End Try

                    ' Nếu có ActiveSheetMetalStyle thì là sheet metal
                    Try
                        If Not String.IsNullOrEmpty(smDef.ActiveSheetMetalStyle.Name) Then
                            Return True
                        End If
                    Catch
                    End Try

                    ' Kiểm tra FlatPattern trực tiếp
                    Try
                        Dim flatPattern As FlatPattern = smDef.FlatPattern
                        If flatPattern IsNot Nothing Then
                            Return True
                        End If
                    Catch
                    End Try
                End If

                ' Cách 2: Kiểm tra qua sub type
                Try
                    If partDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then ' Sheet metal subtype
                        Return True
                    End If
                Catch
                End Try

                Return False
            Catch ex As Exception
                Return False
            End Try
        End Function

        '=====================================================
        ' KIỂM TRA KEYWORD TỪ KÝ TỰ ĐẦU TIÊN
        ' SO SÁNH TỪ TRÁI SANG PHẢI
        '=====================================================
        Private Function IsLibraryPart(ByVal pn As String) As Boolean

            If String.IsNullOrWhiteSpace(pn) Then
                Return False
            End If

            pn = pn.Trim().ToUpperInvariant()

            Try
                If ToolInventor2025.Assembly.CommonKeywords.IsBearing(pn) Then
                    Return True
                End If
            Catch
            End Try

            Try
                If ToolInventor2025.Assembly.CommonKeywords.IsFastener(pn) Then
                    Return True
                End If
            Catch
            End Try

            Try
                If ToolInventor2025.Assembly.CommonKeywords.IsStandardKeyword(pn) Then
                    Return True
                End If
            Catch
            End Try

            Return False

        End Function

        '=====================================================
        ' Dialog chọn loại
        '=====================================================
        Private Function SimpleSelectPartType() As Integer
            Dim result As Integer = 0

            Using dlg As New Form()
                dlg.Text = "CHỌN DẠNG CHI TIẾT"
                dlg.StartPosition = FormStartPosition.CenterScreen
                dlg.Size = New System.Drawing.Size(420, 220)
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog
                dlg.MaximizeBox = False
                dlg.MinimizeBox = False

                Dim b1 As New Button() With {.Text = "SHEET METAL", .Left = 20, .Top = 20, .Width = 160, .Height = 60}
                Dim b2 As New Button() With {.Text = "PURCHASED", .Left = 220, .Top = 20, .Width = 160, .Height = 60}
                Dim b3 As New Button() With {.Text = "STANDARD LIBRARY", .Left = 20, .Top = 100, .Width = 160, .Height = 60}
                Dim b4 As New Button() With {.Text = "PURCHASED + LIBRARY", .Left = 220, .Top = 100, .Width = 160, .Height = 60}

                AddHandler b1.Click, Sub()
                                         result = 1
                                         dlg.DialogResult = DialogResult.OK
                                     End Sub
                AddHandler b2.Click, Sub()
                                         result = 2
                                         dlg.DialogResult = DialogResult.OK
                                     End Sub
                AddHandler b3.Click, Sub()
                                         result = 3
                                         dlg.DialogResult = DialogResult.OK
                                     End Sub
                AddHandler b4.Click, Sub()
                                         result = 4
                                         dlg.DialogResult = DialogResult.OK
                                     End Sub

                dlg.Controls.Add(b1)
                dlg.Controls.Add(b2)
                dlg.Controls.Add(b3)
                dlg.Controls.Add(b4)

                dlg.ShowDialog()
            End Using

            Return result
        End Function

        Private Function SelectAssemblyFile() As String
            Try
                Using dlg As New OpenFileDialog()
                    dlg.Filter = "Inventor Assembly (*.iam)|*.iam|All files (*.*)|*.*"
                    If dlg.ShowDialog() = DialogResult.OK Then
                        Return dlg.FileName
                    End If
                End Using
            Catch
            End Try
            Return ""
        End Function

        Private Function GetProperty(doc As Document, propName As String) As String
            Try
                If doc Is Nothing Then Return ""
                Dim ps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim prop As Inventor.Property = ps.Item(propName)
                If prop Is Nothing OrElse prop.Value Is Nothing Then Return ""
                Return CStr(prop.Value).Trim()
            Catch
                Return ""
            End Try
        End Function

    End Module
End Namespace