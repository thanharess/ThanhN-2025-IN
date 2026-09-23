Option Explicit On
Option Strict Off
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports IO = System.IO

Namespace ToolInventor2025.Assembly.Buttons.caclenhboctach.part
    Public Module Ass_boctach_part_1d

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

                Dim partType As Integer = SimpleSelectPartType()
                If partType = 0 Then
                    If openedByCode Then oOrigAsmDoc.Close(False)
                    Exit Sub
                End If

                Dim targetDoc As AssemblyDocument = CType(oApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject), AssemblyDocument)
                Dim baseName As String = IO.Path.GetFileNameWithoutExtension(selectedFile)
                targetDoc.DisplayName = baseName & "_TopLevel_WithQty"

                '=====================================================
                ' DÙNG STRUCTURED BOM - FIRST LEVEL ONLY
                ' ĐỂ CHỈ LẤY TOP LEVEL
                '=====================================================
                Dim oBOM As BOM = oOrigAsmDoc.ComponentDefinition.BOM
                Try : oBOM.StructuredViewEnabled = True : Catch : End Try
                Try : oBOM.StructuredViewFirstLevelOnly = True : Catch : End Try
                Try : oBOM.Update() : Catch : End Try

                Dim view As BOMView = Nothing
                Try
                    view = oBOM.BOMViews.Item("Structured")
                Catch
                    view = Nothing
                End Try

                If view Is Nothing Then
                    MessageBox.Show("Cannot get Structured BOM view.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    If openedByCode Then oOrigAsmDoc.Close(False)
                    targetDoc.Close(False)
                    Exit Sub
                End If

                Dim oTG As TransientGeometry = oApp.TransientGeometry
                Dim oMatrix As Matrix = oTG.CreateMatrix()
                Dim added As Integer = 0

                ' Dictionary để theo dõi số lượng của mỗi part (cộng dồn)
                Dim partQuantities As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                For Each row As BOMRow In view.BOMRows
                    Try
                        If row Is Nothing OrElse row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then Continue For
                        Dim doc As Document = row.ComponentDefinitions.Item(1).Document
                        If doc Is Nothing Then Continue For

                        '==========================================================
                        ' CHỈ LẤY PART - KHÔNG LẤY ASSEMBLY / CỤM LẮP
                        '==========================================================
                        If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                            Continue For
                        End If

                        Dim include As Boolean = False
                        If partType = 1 Then
                            ' SHEET METAL
                            Dim part As PartDocument = CType(doc, PartDocument)
                            Dim sm As SheetMetalComponentDefinition = TryCast(part.ComponentDefinition, SheetMetalComponentDefinition)
                            include = (sm IsNot Nothing AndAlso sm.Features.Count > 0)
                        ElseIf partType = 2 Then
                            ' PURCHASED
                            include = (row.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure)
                        ElseIf partType = 3 Then
                            ' STANDARD LIBRARY
                            Dim pn As String = GetProperty(doc, "Part Number")
                            If pn Is Nothing Then pn = ""
                            pn = pn.Trim().ToUpperInvariant()
                            include = IsLibraryPart(pn)
                        ElseIf partType = 4 Then
                            ' PURCHASED + LIBRARY
                            Dim pn As String = GetProperty(doc, "Part Number")
                            If pn Is Nothing Then pn = ""
                            pn = pn.Trim().ToUpperInvariant()
                            Dim purchased As Boolean = (row.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure)
                            Dim library As Boolean = IsLibraryPart(pn)
                            include = (purchased OrElse library)
                        End If

                        If Not include Then Continue For

                        ' Lấy số lượng từ BOM row
                        Dim qty As Integer = 1
                        Try
                            qty = CInt(Math.Round(CDbl(row.ItemQuantity), 0))
                            If qty < 1 Then qty = 1
                        Catch
                            qty = 1
                        End Try

                        ' Lấy key của document
                        Dim key As String = GetDocKey(doc)

                        ' Cộng dồn số lượng
                        If partQuantities.ContainsKey(key) Then
                            partQuantities(key) += qty
                        Else
                            partQuantities.Add(key, qty)
                        End If

                    Catch ex As Exception
                        ' Bỏ qua row lỗi
                    End Try
                Next

                '=====================================================
                ' THÊM PARTS VỚI SỐ LƯỢNG ĐÃ CỘNG DỒN
                '=====================================================
                For Each kvp As KeyValuePair(Of String, Integer) In partQuantities
                    Try
                        Dim partFullName As String = kvp.Key
                        Dim qty As Integer = kvp.Value

                        ' Thêm occurrence với số lượng đã cộng dồn
                        For i As Integer = 1 To qty
                            Try
                                targetDoc.ComponentDefinition.Occurrences.Add(partFullName, oMatrix)
                                added += 1
                            Catch
                                ' Bỏ qua nếu không thêm được
                            End Try
                        Next

                    Catch ex As Exception
                        ' Bỏ qua part lỗi
                    End Try
                Next

                targetDoc.Update2(True)
                MessageBox.Show("Done. Added occurrences: " & added.ToString(), "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)

                If openedByCode Then oOrigAsmDoc.Close(False)

            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

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

        Private Function GetDocKey(doc As Document) As String
            Try
                If doc Is Nothing Then Return ""
                Return doc.FullFileName.ToLowerInvariant()
            Catch
                Return ""
            End Try
        End Function

        '==========================================================
        ' KIỂM TRA KEYWORD TỪ KÝ TỰ ĐẦU TIÊN
        ' SO SÁNH TỪ TRÁI SANG PHẢI
        '==========================================================
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
    End Module
End Namespace