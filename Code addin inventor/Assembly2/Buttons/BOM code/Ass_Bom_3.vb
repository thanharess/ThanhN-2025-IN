Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports Microsoft.VisualBasic
Imports System.Globalization
Imports System.Runtime.InteropServices

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode

    Public Module Ass_Bom_3

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim choice As Integer = PickFromList(
                "Chọn chức năng copy property",
                New String() {
                    "1 - Copy Item Number → item1 (Top-level)",
                    "1b - Copy Item Number → item1 (All-level)",
                    "1c - Copy Item Number → item0 (Structure)",
                    "2 - Copy Qty cụm (Status) + Part top (SL part)",
                    "3 - Copy Item Number + Qty cụm/Part top",
                    "4 - Ghi / sửa property PL (Sheet Metal)",
                    "5 - Copy Qty Part all-level → SL part all",
                    "6 - Copy TẤT CẢ (1+2+4+5)"
                }, 0)

            If choice < 0 Then Exit Sub

            Dim doItem1Top As Boolean = (choice = 0 OrElse choice = 4 OrElse choice = 7)
            Dim doItem1All As Boolean = (choice = 1)
            Dim doItem0 As Boolean = (choice = 2)
            Dim doQtyTop As Boolean = (choice = 3 OrElse choice = 4 OrElse choice = 7)
            Dim doPL As Boolean = (choice = 5 OrElse choice = 7)
            Dim doSLall As Boolean = (choice = 6 OrElse choice = 7)

            Try
                Dim invApp As Inventor.Application = Nothing
                Try
                    invApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
                Catch
                    MessageBox.Show("Không lấy được Inventor.", "BOM")
                    Exit Sub
                End Try

                Dim oAsm As AssemblyDocument = TryCast(invApp.ActiveDocument, AssemblyDocument)
                If oAsm Is Nothing Then
                    MessageBox.Show("Rule này chỉ chạy trong Assembly.", "BOM")
                    Exit Sub
                End If

                Dim oBOM As BOM = oAsm.ComponentDefinition.BOM
                oBOM.StructuredViewEnabled = True
                oBOM.StructuredViewFirstLevelOnly = False
                Try : oBOM.PartsOnlyViewEnabled = True : Catch : End Try
                Try : oBOM.Update() : Catch : End Try

                Dim oBOMView As BOMView = oBOM.BOMViews.Item("Structured")

                Dim countItem1 As Integer = 0
                Dim countItem0 As Integer = 0
                Dim countStatus As Integer = 0
                Dim countSLpart As Integer = 0
                Dim countPL As Integer = 0
                Dim countSLall As Integer = 0

                ' PN → giá trị
                Dim dictItem1 As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                Dim dictItem0 As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                ' 1. Lấy số từ Structured
                If doItem1Top OrElse doItem1All OrElse doItem0 Then
                    CollectFromStructured(oBOMView.BOMRows, doItem1Top, doItem1All, doItem0,
                                          dictItem1, dictItem0, True)
                End If

                ' 2. Ghi vào tất cả document (Structured + Model Data / mọi occurrence)
                '    cùng PN → cùng số
                Dim writtenDocs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                ' Ghi từ Structured trước
                WriteToAllOccurrences(oAsm.ComponentDefinition.Occurrences, dictItem1, dictItem0,
                                      writtenDocs, countItem1, countItem0)

                ' 3. Qty / PL / SLall giữ nguyên
                If doQtyTop Then
                    For Each row As BOMRow In oBOMView.BOMRows
                        If IsSkipped(row) Then Continue For
                        Dim refDoc As Document = GetDoc(row)
                        If refDoc Is Nothing Then Continue For

                        Dim qtyStr As String = GetQty(row)
                        If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            If SetStatusProperty(refDoc, qtyStr) Then countStatus += 1
                        ElseIf refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                            If SetUserProperty(refDoc, "SL part", qtyStr) Then countSLpart += 1
                        End If
                    Next
                End If

                If doPL Then
                    For Each row As BOMRow In oBOMView.BOMRows
                        If IsSkipped(row) Then Continue For
                        Dim refDoc As Document = GetDoc(row)
                        If refDoc Is Nothing Then Continue For
                        If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                            If ProcessSheetMetalPL(CType(refDoc, PartDocument)) Then countPL += 1
                        End If
                    Next
                End If

                If doSLall Then
                    Try
                        Dim partsView As BOMView = oBOM.BOMViews.Item("Parts Only")
                        For Each row As BOMRow In partsView.BOMRows
                            If IsSkipped(row) Then Continue For
                            Dim refDoc As Document = GetDoc(row)
                            If refDoc Is Nothing Then Continue For
                            If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                                Dim qtyStr As String = GetQty(row)
                                If SetUserProperty(refDoc, "SL part all", qtyStr) Then countSLall += 1
                            End If
                        Next
                    Catch
                    End Try
                End If

                Try : oBOM.Update() : Catch : End Try
                Try : oAsm.Update2(True) : Catch : End Try

                Dim msg As String = "HOÀN TẤT" & vbCrLf & "========================" & vbCrLf
                If doItem1Top OrElse doItem1All Then msg &= "item1 : " & countItem1.ToString() & vbCrLf
                If doItem0 Then msg &= "item0 : " & countItem0.ToString() & vbCrLf
                If doQtyTop Then
                    msg &= "Status (cụm) : " & countStatus.ToString() & vbCrLf
                    msg &= "SL part : " & countSLpart.ToString() & vbCrLf
                End If
                If doPL Then msg &= "PL (SheetMetal): " & countPL.ToString() & vbCrLf
                If doSLall Then msg &= "SL part all : " & countSLall.ToString() & vbCrLf

                MessageBox.Show(msg, "Assembly BOM")

            Catch ex As Exception
                MessageBox.Show("Có lỗi:" & vbCrLf & ex.Message, "Assembly BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        Private Function GetDoc(row As BOMRow) As Document
            Try
                If row Is Nothing OrElse row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then Return Nothing
                Return row.ComponentDefinitions.Item(1).Document
            Catch
                Return Nothing
            End Try
        End Function

        Private Function GetPartNumberFromDoc(doc As Document) As String
            Try
                If doc Is Nothing Then Return ""
                Dim ps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim prop As Inventor.Property = ps.Item("Part Number")
                If prop Is Nothing OrElse prop.Value Is Nothing Then Return ""
                Return CStr(prop.Value).Trim()
            Catch
                Return ""
            End Try
        End Function

        ' 1. Thu thập số từ Structured theo PN
        Private Sub CollectFromStructured(rows As BOMRowsEnumerator,
                                          doItem1Top As Boolean,
                                          doItem1All As Boolean,
                                          doItem0 As Boolean,
                                          dictItem1 As Dictionary(Of String, String),
                                          dictItem0 As Dictionary(Of String, String),
                                          isTopLevel As Boolean)

            If rows Is Nothing Then Exit Sub

            For Each row As BOMRow In rows
                If IsSkipped(row) Then Continue For

                Dim refDoc As Document = GetDoc(row)
                If refDoc Is Nothing Then Continue For

                Dim pn As String = GetPartNumberFromDoc(refDoc)
                If String.IsNullOrEmpty(pn) Then pn = refDoc.DisplayName

                Dim structItem As String = ""
                Try : structItem = row.ItemNumber : Catch : End Try
                If String.IsNullOrEmpty(structItem) Then Continue For

                If (doItem1Top AndAlso isTopLevel) OrElse doItem1All Then
                    Dim item1Val As String = structItem
                    If item1Val.Contains(".") Then
                        item1Val = item1Val.Substring(item1Val.LastIndexOf("."c) + 1)
                    End If
                    If Not dictItem1.ContainsKey(pn) Then
                        dictItem1(pn) = item1Val
                    End If
                End If

                If doItem0 Then
                    Dim item0Val As String = structItem
                    If isTopLevel AndAlso Not structItem.Contains(".") Then
                        item0Val = structItem & ".0"
                    End If
                    If Not dictItem0.ContainsKey(pn) Then
                        dictItem0(pn) = item0Val
                    End If
                End If

                Try
                    If row.ChildRows IsNot Nothing AndAlso row.ChildRows.Count > 0 Then
                        CollectFromStructured(row.ChildRows, doItem1Top, doItem1All, doItem0,
                                              dictItem1, dictItem0, False)
                    End If
                Catch
                End Try
            Next
        End Sub

        ' 2. Duyệt toàn bộ occurrence (Model Data + mọi cấp) → ghi theo PN
        Private Sub WriteToAllOccurrences(occs As ComponentOccurrences,
                                          dictItem1 As Dictionary(Of String, String),
                                          dictItem0 As Dictionary(Of String, String),
                                          writtenDocs As HashSet(Of String),
                                          ByRef countItem1 As Integer,
                                          ByRef countItem0 As Integer)

            If occs Is Nothing Then Exit Sub

            For Each occ As ComponentOccurrence In occs
                Try
                    Dim doc As Document = Nothing
                    Try : doc = occ.Definition.Document : Catch : Continue For : End Try
                    If doc Is Nothing Then Continue For

                    Dim key As String = doc.FullFileName
                    If writtenDocs.Contains(key) Then
                        ' đã ghi rồi, vẫn đệ quy con
                    Else
                        Dim pn As String = GetPartNumberFromDoc(doc)
                        If String.IsNullOrEmpty(pn) Then pn = doc.DisplayName

                        Dim written As Boolean = False

                        If dictItem1.ContainsKey(pn) Then
                            If SetUserProperty(doc, "item1", dictItem1(pn)) Then
                                countItem1 += 1
                                written = True
                            End If
                        End If

                        If dictItem0.ContainsKey(pn) Then
                            If SetUserProperty(doc, "item0", dictItem0(pn)) Then
                                countItem0 += 1
                                written = True
                            End If
                        End If

                        If written Then writtenDocs.Add(key)
                    End If

                    ' Đệ quy sub-assembly
                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Try
                            WriteToAllOccurrences(occ.SubOccurrences, dictItem1, dictItem0,
                                                  writtenDocs, countItem1, countItem0)
                        Catch
                        End Try
                    End If
                Catch
                End Try
            Next
        End Sub

        Private Function IsSkipped(row As BOMRow) As Boolean
            Try
                If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then Return True
                If row.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then Return True
            Catch
            End Try
            Return False
        End Function

        Private Function GetQty(row As BOMRow) As String
            Try
                ' Ưu tiên ItemQuantity (đúng cho thép hình – số lượng = kích thước)
                Return CStr(row.ItemQuantity)
            Catch
                Try
                    Return CStr(row.TotalQuantity)
                Catch
                    Return "1"
                End Try
            End Try
        End Function

        Private Function ProcessSheetMetalPL(partDoc As PartDocument) As Boolean
            Try
                If partDoc Is Nothing OrElse Not partDoc.IsModifiable Then Return False
                Dim smDef As SheetMetalComponentDefinition =
                    TryCast(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
                If smDef Is Nothing Then Return False

                Dim thickMM As Double = smDef.Thickness.Value * 10.0
                Dim thickStr As String = FormatThickness(thickMM)
                Dim newPrefix As String = "t" & thickStr

                Dim current As String = GetUserProperty(partDoc, "PL")
                Dim finalValue As String = ""

                If String.IsNullOrEmpty(current) Then
                    finalValue = newPrefix
                Else
                    current = current.Trim()
                    If current.StartsWith("t", StringComparison.OrdinalIgnoreCase) Then
                        Dim afterPL As String = current.Substring(1)
                        Dim oldThickStr As String = ""
                        Dim rest As String = ""
                        Dim i As Integer = 0
                        While i < afterPL.Length AndAlso (Char.IsDigit(afterPL(i)) OrElse afterPL(i) = "."c)
                            oldThickStr &= afterPL(i)
                            i += 1
                        End While
                        If i < afterPL.Length Then rest = afterPL.Substring(i)

                        Dim oldThick As Double = 0
                        Double.TryParse(oldThickStr, NumberStyles.Any, CultureInfo.InvariantCulture, oldThick)

                        If Math.Abs(oldThick - thickMM) < 0.001 Then
                            Return False
                        Else
                            finalValue = newPrefix & rest
                        End If
                    Else
                        finalValue = newPrefix
                    End If
                End If

                Return SetUserProperty(partDoc, "PL", finalValue)
            Catch
                Return False
            End Try
        End Function

        Private Function FormatThickness(value As Double) As String
            Return value.ToString("0.###", CultureInfo.InvariantCulture)
        End Function

        Private Function GetUserProperty(doc As Document, propName As String) As String
            Try
                Dim userProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
                For Each p As Inventor.Property In userProps
                    If String.Equals(p.Name, propName, StringComparison.OrdinalIgnoreCase) Then
                        If p.Value Is Nothing Then Return ""
                        Return CStr(p.Value).Trim()
                    End If
                Next
            Catch
            End Try
            Return ""
        End Function

        Private Function SetUserProperty(doc As Document, propName As String, value As String) As Boolean
            Try
                If doc Is Nothing OrElse Not doc.IsModifiable Then Return False
                Dim userProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
                Dim found As Inventor.Property = Nothing
                For Each p As Inventor.Property In userProps
                    If String.Equals(p.Name, propName, StringComparison.OrdinalIgnoreCase) Then
                        found = p
                        Exit For
                    End If
                Next
                If found Is Nothing Then
                    userProps.Add(value, propName)
                Else
                    found.Value = value
                End If
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Function SetStatusProperty(doc As Document, value As String) As Boolean
            Try
                If doc Is Nothing OrElse Not doc.IsModifiable Then Return False
                Try
                    Dim designProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                    Dim prop As Inventor.Property = designProps.Item("Status")
                    prop.Value = value
                    Return True
                Catch
                End Try
                Return SetUserProperty(doc, "Status", value)
            Catch
                Return False
            End Try
        End Function

        Private Function PickFromList(title As String, items As String(),
                                      Optional defaultIndex As Integer = 0) As Integer
            Dim frm As New Form()
            Try
                frm.Text = title
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.FormBorderStyle = FormBorderStyle.FixedDialog
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.Width = 520
                frm.Height = 380

                Dim lst As New ListBox()
                lst.Left = 12 : lst.Top = 12
                lst.Width = 480 : lst.Height = 270
                lst.Font = New System.Drawing.Font("Segoe UI", 10)

                For Each s As String In items
                    lst.Items.Add(s)
                Next

                If lst.Items.Count > 0 Then
                    If defaultIndex >= 0 AndAlso defaultIndex < lst.Items.Count Then
                        lst.SelectedIndex = defaultIndex
                    Else
                        lst.SelectedIndex = 0
                    End If
                End If

                Dim btnOK As New Button()
                btnOK.Text = "OK"
                btnOK.Left = 300 : btnOK.Top = 295
                btnOK.Width = 90 : btnOK.Height = 30
                btnOK.DialogResult = DialogResult.OK

                Dim btnCancel As New Button()
                btnCancel.Text = "Hủy"
                btnCancel.Left = 400 : btnCancel.Top = 295
                btnCancel.Width = 90 : btnCancel.Height = 30
                btnCancel.DialogResult = DialogResult.Cancel

                frm.Controls.Add(lst)
                frm.Controls.Add(btnOK)
                frm.Controls.Add(btnCancel)
                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel
                frm.KeyPreview = True

                AddHandler lst.DoubleClick, Sub(s, e)
                                                frm.DialogResult = DialogResult.OK
                                                frm.Close()
                                            End Sub

                AddHandler frm.KeyDown, Sub(s, e)
                                            If e.KeyCode = Keys.Escape Then
                                                e.Handled = True
                                                frm.DialogResult = DialogResult.Cancel
                                                frm.Close()
                                            End If
                                        End Sub

                If frm.ShowDialog() <> DialogResult.OK Then Return -1
                If lst.SelectedIndex < 0 Then Return -1
                Return lst.SelectedIndex
            Finally
                If frm IsNot Nothing Then
                    Try : frm.Dispose() : Catch : End Try
                End If
            End Try
        End Function

    End Module

End Namespace