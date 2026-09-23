Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports ToolInventor2025.ToolInventor2025.Assembly.Buttons

Namespace ToolInventor2025.Drawing.Buttons.DrawPartList

    Public Module Draw_Partlist_1c

        ' Property Set ID của Document Summary Information
        Private Const DOC_SUMMARY_PROPSET As String = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}"
        ' PropId của Category = 2
        Private Const CATEGORY_PROPID As Long = 2

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then

                    MessageBox.Show("Vui lòng mở file Drawing (.idw)!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As Inventor.DrawingDocument = CType(app.ActiveDocument, Inventor.DrawingDocument)
                Dim oSheet As Inventor.Sheet = oDrawDoc.ActiveSheet

                If oSheet.PartsLists.Count < 1 Then
                    MessageBox.Show("Sheet hiện tại không có Parts List.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                '=================================================
                ' 1. NHẬP PREFIX
                '=================================================
                Dim PREFIX As String = InputBox("Nhập PREFIX mã bản vẽ (ví dụ: 1.2.3.)", "PREFIX", "")

                If PREFIX Is Nothing OrElse PREFIX.Trim() = "" Then
                    Exit Sub
                End If
                PREFIX = PREFIX.Trim()

                '=================================================
                ' 2. GHI XUỐNG BOM GỐC?
                '=================================================
                Dim writeToBomIdx As Integer = PickFromList("Ghi xuống BOM gốc?", New String() {
                    "1 - Chỉ ghi trên Parts List (không đụng BOM)",
                    "2 - Ghi cả Parts List + BOM gốc (Category iProperty)"}, 0)

                If writeToBomIdx < 0 Then Exit Sub

                Dim writeToBom As Boolean = (writeToBomIdx = 1)

                '=================================================
                ' 3. PHẠM VI
                '=================================================
                Dim scopeIdx As Integer = PickFromList("Phạm vi", New String() {
                    "1 - Chỉ Parts List đầu trên sheet active",
                    "2 - Tất cả Parts List trên sheet active",
                    "3 - Tất cả Parts List của toàn bộ Drawing"}, 0)

                If scopeIdx < 0 Then Exit Sub

                Dim processed As Integer = 0
                Dim totalRows As Integer = 0

                If scopeIdx = 0 Then
                    Try
                        Dim oPartList As Inventor.PartsList = oSheet.PartsLists.Item(1)
                        totalRows += ProcessOnePartsList(oPartList, PREFIX, writeToBom)
                        processed += 1
                    Catch ex As Exception
                        MessageBox.Show("Parts List 1:" & vbCrLf & ex.Message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Try

                ElseIf scopeIdx = 1 Then
                    For plIdx As Integer = 1 To oSheet.PartsLists.Count
                        Try
                            Dim oPartList As Inventor.PartsList = oSheet.PartsLists.Item(plIdx)
                            totalRows += ProcessOnePartsList(oPartList, PREFIX, writeToBom)
                            processed += 1
                        Catch exPL As Exception
                            MessageBox.Show("Sheet: " & oSheet.Name & vbCrLf & "Parts List: " & plIdx.ToString() &
                                            vbCrLf & exPL.Message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End Try
                    Next

                ElseIf scopeIdx = 2 Then
                    For sheetIdx As Integer = 1 To oDrawDoc.Sheets.Count
                        Try
                            Dim oCurSheet As Inventor.Sheet = oDrawDoc.Sheets.Item(sheetIdx)

                            For plIdx As Integer = 1 To oCurSheet.PartsLists.Count
                                Try
                                    Dim oPartList As Inventor.PartsList = oCurSheet.PartsLists.Item(plIdx)
                                    totalRows += ProcessOnePartsList(oPartList, PREFIX, writeToBom)
                                    processed += 1
                                Catch exPL As Exception
                                    MessageBox.Show("Sheet: " & oCurSheet.Name & vbCrLf & "Parts List: " & plIdx.ToString() &
                                                    vbCrLf & exPL.Message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                End Try
                            Next
                        Catch exSheet As Exception
                            MessageBox.Show("Lỗi Sheet " & sheetIdx.ToString() & ":" & vbCrLf & exSheet.Message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End Try
                    Next
                End If

                Try
                    oDrawDoc.Update()
                Catch
                End Try

                MessageBox.Show("Hoàn tất!" & vbCrLf &
                                "Parts List đã xử lý: " & processed.ToString() & vbCrLf &
                                "Số dòng đã ghi mã: " & totalRows.ToString() & vbCrLf &
                                "Định dạng: " & PREFIX & "STT" & vbCrLf &
                                "Ghi BOM gốc: " & If(writeToBom, "Có", "Không"),
                                "Ghi mã bản vẽ (Item → Category)",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Ghi mã bản vẽ (Item → Category)", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        '=========================================================
        ' XỬ LÝ 1 PARTS LIST
        '=========================================================
        Private Function ProcessOnePartsList(
            oPartList As Inventor.PartsList,
            PREFIX As String,
            writeToBom As Boolean) As Integer

            Dim count As Integer = 0

            Dim cItem As String = FindItemColumn(oPartList)
            Dim cCategory As String = FindCategoryColumn(oPartList)

            If cItem = "" Then
                Throw New Exception("Không tìm thấy cột Item (kItemPartsListProperty) trên Parts List.")
            End If

            If cCategory = "" Then
                Throw New Exception("Không tìm thấy cột Category trên Parts List." & vbCrLf &
                                    "Hãy thêm cột Category vào Parts List trước.")
            End If

            For i As Integer = 1 To oPartList.PartsListRows.Count

                Try
                    Dim row As Inventor.PartsListRow = oPartList.PartsListRows.Item(i)

                    Dim itemValue As String = GetCellValue(row, cItem)
                    If itemValue = "" Then Continue For

                    Dim newCode As String = PREFIX & itemValue
                    Dim currentCategory As String = GetCellValue(row, cCategory)

                    ' 1. Ghi trên Parts List
                    If Not String.Equals(currentCategory, newCode, StringComparison.OrdinalIgnoreCase) Then
                        SetCell(row, cCategory, newCode)
                        count += 1
                    End If

                    ' 2. Ghi xuống BOM gốc (Category iProperty của component)
                    If writeToBom Then
                        WriteCategoryToBOM(row, newCode)
                    End If

                Catch
                End Try

            Next

            Try
                oPartList.Update()
            Catch
            End Try

            ' Nếu ghi BOM thì SaveItemOverridesToBOM (đồng bộ override nếu cần)
            If writeToBom Then
                Try
                    oPartList.SaveItemOverridesToBOM()
                Catch
                End Try
            End If

            Return count

        End Function

        '=========================================================
        ' GHI CATEGORY XUỐNG DOCUMENT GỐC (BOM)
        '=========================================================
        Private Sub WriteCategoryToBOM(row As Inventor.PartsListRow, newCode As String)

            Try
                If row.ReferencedRows Is Nothing OrElse row.ReferencedRows.Count < 1 Then
                    Exit Sub
                End If

                Dim bomRow As Inventor.BOMRow = row.ReferencedRows.Item(1).BOMRow
                If bomRow Is Nothing Then Exit Sub

                If bomRow.ComponentDefinitions.Count < 1 Then Exit Sub

                Dim refDoc As Inventor.Document = bomRow.ComponentDefinitions.Item(1).Document
                If refDoc Is Nothing Then Exit Sub

                ' Ghi vào iProperty Category (Document Summary Information)
                SetCategoryProperty(refDoc, newCode)

            Catch
            End Try

        End Sub

        '=========================================================
        ' SET CATEGORY iPROPERTY
        '=========================================================
        Private Sub SetCategoryProperty(doc As Inventor.Document, value As String)

            If doc Is Nothing OrElse value Is Nothing Then Exit Sub

            Dim newValue As String = value.Trim()
            If newValue = "" Then Exit Sub

            Try
                Dim ps As Inventor.PropertySet = doc.PropertySets.Item(DOC_SUMMARY_PROPSET)
                Dim prop As Inventor.Property = ps.ItemByPropId(CATEGORY_PROPID)

                Dim oldValue As String = ""
                Try
                    If prop.Value IsNot Nothing Then
                        oldValue = CStr(prop.Value).Trim()
                    End If
                Catch
                End Try

                If String.Equals(oldValue, newValue, StringComparison.OrdinalIgnoreCase) Then
                    Exit Sub
                End If

                prop.Value = newValue

                Try
                    doc.Update()
                Catch
                End Try

            Catch
                ' Fallback: thử theo tên
                Try
                    Dim ps2 As Inventor.PropertySet = doc.PropertySets.Item("Inventor Document Summary Information")
                    ps2.Item("Category").Value = newValue
                    Try
                        doc.Update()
                    Catch
                    End Try
                Catch
                End Try
            End Try

        End Sub

        '=========================================================
        ' TÌM CỘT ITEM theo PropertyType
        '=========================================================
        Private Function FindItemColumn(pl As Inventor.PartsList) As String

            Try
                For Each col As Inventor.PartsListColumn In pl.PartsListColumns
                    Try
                        If col.PropertyType = PropertyTypeEnum.kItemPartsListProperty Then
                            Return col.Title
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            Return ""

        End Function

        '=========================================================
        ' TÌM CỘT CATEGORY theo PropertyType + PropId
        '=========================================================
        Private Function FindCategoryColumn(pl As Inventor.PartsList) As String

            Try
                For Each col As Inventor.PartsListColumn In pl.PartsListColumns
                    Try
                        If col.PropertyType = PropertyTypeEnum.kFileProperty Then

                            Dim propSetId As String = ""
                            Dim propId As Long = 0

                            col.GetFilePropertyId(propSetId, propId)

                            If String.Equals(propSetId, DOC_SUMMARY_PROPSET, StringComparison.OrdinalIgnoreCase) AndAlso
                               propId = CATEGORY_PROPID Then

                                Return col.Title
                            End If

                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            Return ""

        End Function

        '=========================================================
        ' GET / SET CELL
        '=========================================================
        Private Function GetCellValue(row As Inventor.PartsListRow, colName As String) As String
            If colName = "" Then Return ""

            Try
                Dim v As Object = row.Item(colName).Value
                If v Is Nothing Then Return ""
                Return CStr(v).Trim()
            Catch
                Return ""
            End Try
        End Function

        Private Sub SetCell(row As Inventor.PartsListRow, colName As String, value As String)

            If colName = "" OrElse value Is Nothing Then Exit Sub

            Dim newValue As String = value.Trim()
            If newValue = "" Then Exit Sub

            Try
                Dim cell As Inventor.PartsListCell = row.Item(colName)

                Dim oldValue As String = ""
                Try
                    If cell.Value IsNot Nothing Then
                        oldValue = CStr(cell.Value).Trim()
                    End If
                Catch
                End Try

                If String.Equals(oldValue, newValue, StringComparison.OrdinalIgnoreCase) Then
                    Exit Sub
                End If

                cell.Value = newValue

                Try
                    cell.Static = True
                Catch
                End Try

            Catch
            End Try

        End Sub

        '=========================================================
        ' PICK LIST
        '=========================================================
        Private Function PickFromList(title As String, items As String(), Optional defaultIndex As Integer = 0) As Integer

            Dim frm As New Form()
            frm.Text = title
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.Width = 520
            frm.Height = 320
            frm.ShowInTaskbar = False

            Dim lst As New ListBox()
            lst.Left = 12
            lst.Top = 12
            lst.Width = 480
            lst.Height = 220

            For Each s As String In items
                lst.Items.Add(s)
            Next

            If defaultIndex >= 0 AndAlso defaultIndex < lst.Items.Count Then
                lst.SelectedIndex = defaultIndex
            ElseIf lst.Items.Count > 0 Then
                lst.SelectedIndex = 0
            End If

            Dim btnOK As New Button() With {
                .Text = "OK",
                .Left = 320,
                .Top = 245,
                .Width = 80,
                .DialogResult = DialogResult.OK
            }

            Dim btnCancel As New Button() With {
                .Text = "Hủy",
                .Left = 410,
                .Top = 245,
                .Width = 80,
                .DialogResult = DialogResult.Cancel
            }

            frm.Controls.Add(lst)
            frm.Controls.Add(btnOK)
            frm.Controls.Add(btnCancel)

            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            If frm.ShowDialog() <> DialogResult.OK OrElse lst.SelectedIndex < 0 Then
                Return -1
            End If

            Return lst.SelectedIndex

        End Function

    End Module

End Namespace