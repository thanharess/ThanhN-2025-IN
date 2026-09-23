Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports Microsoft.VisualBasic
Imports System.Globalization

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode

    Public Module Ass_Bom_4

        Public Sub OnExecute(ByVal Context As NameValueMap)

            '==================================================
            ' CHỌN ÁP DỤNG CHO
            '==================================================
            Dim targetIdx As Integer = PickFromList(
                "Áp dụng kiểm tra / ghi chiều dày Sheet Metal",
                New String() {
                    "Chỉ Part Number",
                    "Chỉ Stock Number",
                    "Cả Part Number và Stock Number"
                }, 2)

            If targetIdx < 0 Then Exit Sub

            Dim applyPN As Boolean = False
            Dim applySN As Boolean = False

            Select Case targetIdx
                Case 0 : applyPN = True
                Case 1 : applySN = True
                Case 2 : applyPN = True : applySN = True
            End Select


            Try
                Dim oAsm As AssemblyDocument =
                    TryCast(g_inventorApplication.ActiveDocument, AssemblyDocument)

                If oAsm Is Nothing Then
                    MessageBox.Show("Rule này chỉ chạy trong Assembly.", "BOM")
                    Exit Sub
                End If

                Dim oBOM As BOM = oAsm.ComponentDefinition.BOM
                Try : oBOM.StructuredViewEnabled = True : Catch : End Try
                Try : oBOM.StructuredViewFirstLevelOnly = False : Catch : End Try

                Dim oBOMView As BOMView = oBOM.BOMViews.Item("Structured")

                Dim countPN As Integer = 0
                Dim countSN As Integer = 0
                Dim listDocs As New List(Of Document)


                For Each row As BOMRow In oBOMView.BOMRows

                    ' Bỏ qua Reference + Phantom
                    If IsSkipped(row) Then Continue For

                    Dim refDoc As Document = Nothing
                    Try
                        refDoc = row.ComponentDefinitions.Item(1).Document
                    Catch
                        Continue For
                    End Try

                    If refDoc Is Nothing Then Continue For
                    If refDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then Continue For
                    If Not refDoc.IsModifiable Then Continue For

                    Dim partDoc As PartDocument = CType(refDoc, PartDocument)

                    ' Chỉ xử lý Sheet Metal
                    Dim smDef As SheetMetalComponentDefinition =
                        TryCast(partDoc.ComponentDefinition, SheetMetalComponentDefinition)

                    If smDef Is Nothing Then Continue For


                    '----- Lấy chiều dày thật (mm) -----
                    Dim thickMM As Double = smDef.Thickness.Value * 10.0
                    Dim thickStr As String = FormatThickness(thickMM)
                    Dim newPrefix As String = "PL" & thickStr


                    '----- Xử lý Part Number -----
                    If applyPN Then
                        Dim curPN As String = GetDesignProperty(partDoc, "Part Number")
                        Dim newPN As String = SmartUpdateThickness(curPN, newPrefix, thickMM)

                        If newPN <> "" AndAlso newPN <> curPN Then
                            If SetDesignProperty(partDoc, "Part Number", newPN) Then
                                countPN += 1
                                If Not listDocs.Contains(partDoc) Then listDocs.Add(partDoc)
                            End If
                        End If
                    End If


                    '----- Xử lý Stock Number -----
                    If applySN Then
                        Dim curSN As String = GetDesignProperty(partDoc, "Stock Number")
                        Dim newSN As String = SmartUpdateThickness(curSN, newPrefix, thickMM)

                        If newSN <> "" AndAlso newSN <> curSN Then
                            If SetDesignProperty(partDoc, "Stock Number", newSN) Then
                                countSN += 1
                                If Not listDocs.Contains(partDoc) Then listDocs.Add(partDoc)
                            End If
                        End If
                    End If

                Next


                '----- Save các file đã sửa -----
                For Each d As Document In listDocs
                    Try
                        If d.IsModifiable Then
                            d.Update()
                            d.Save2(True)
                        End If
                    Catch
                    End Try
                Next

                Try : oBOM.Update() : Catch : End Try
                Try : oAsm.Update2(True) : Catch : End Try


                '----- Thông báo -----
                Dim msg As String =
                    "HOÀN TẤT – Sheet Metal Thickness" & vbCrLf &
                    "=================================" & vbCrLf &
                    "Part Number đã sửa  : " & countPN.ToString() & vbCrLf &
                    "Stock Number đã sửa : " & countSN.ToString()

                MessageBox.Show(msg, "PL → Part Number / Stock Number")

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '==========================================================
        ' LOGIC THÔNG MINH GIỐNG LỰA CHỌN 4
        ' - Chưa có → ghi "PL" + chiều dày
        ' - Đã có "PL" + số:
        '     + Đúng chiều dày → giữ nguyên (trả về chuỗi cũ)
        '     + Sai chiều dày → chỉ sửa phần số, giữ phần còn lại
        '       Ví dụ: PL4x6x7  + dày thật 4.6 → PL4.6x6x7
        '==========================================================
        Private Function SmartUpdateThickness(current As String, newPrefix As String, realThick As Double) As String

            If current Is Nothing Then current = ""
            current = current.Trim()

            ' Chưa có gì → ghi mới
            If current = "" Then
                Return newPrefix
            End If

            ' Không bắt đầu bằng PL → ghi đè
            If Not current.StartsWith("PL", StringComparison.OrdinalIgnoreCase) Then
                Return newPrefix
            End If

            ' Tách phần sau "PL"
            Dim afterPL As String = current.Substring(2)
            Dim oldThickStr As String = ""
            Dim rest As String = ""
            Dim i As Integer = 0

            While i < afterPL.Length AndAlso (Char.IsDigit(afterPL(i)) OrElse afterPL(i) = "."c)
                oldThickStr &= afterPL(i)
                i += 1
            End While

            If i < afterPL.Length Then
                rest = afterPL.Substring(i)          ' ví dụ "x6x7"
            End If

            ' So sánh chiều dày
            Dim oldThick As Double = 0
            Double.TryParse(oldThickStr, NumberStyles.Any, CultureInfo.InvariantCulture, oldThick)

            If Math.Abs(oldThick - realThick) < 0.001 Then
                ' Đúng → giữ nguyên
                Return current
            Else
                ' Sai → sửa phần dày, giữ phần còn lại
                Return newPrefix & rest
            End If

        End Function


        '==========================================================
        ' FORMAT CHIỀU DÀY
        '==========================================================
        Private Function FormatThickness(value As Double) As String
            Return value.ToString("0.###", CultureInfo.InvariantCulture)
        End Function


        '==========================================================
        ' BỎ QUA Reference + Phantom
        '==========================================================
        Private Function IsSkipped(row As BOMRow) As Boolean
            Try
                If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then Return True
                If row.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then Return True
            Catch
            End Try
            Return False
        End Function


        '==========================================================
        ' DESIGN TRACKING PROPERTY
        '==========================================================
        Private Function GetDesignProperty(doc As Document, propName As String) As String
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


        Private Function SetDesignProperty(doc As Document, propName As String, value As String) As Boolean
            Try
                If doc Is Nothing OrElse Not doc.IsModifiable Then Return False

                Dim designProps As PropertySet = Nothing
                Try
                    designProps = doc.PropertySets.Item("Design Tracking Properties")
                Catch
                    Return False
                End Try

                Try
                    Dim prop As Inventor.Property = designProps.Item(propName)
                    prop.Value = value
                    Return True
                Catch
                    Try
                        designProps.Add(value, propName)
                        Return True
                    Catch
                        Return False
                    End Try
                End Try
            Catch
                Return False
            End Try
        End Function


        '==========================================================
        ' PICK LIST
        '==========================================================
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
                frm.Width = 420
                frm.Height = 280

                Dim lst As New ListBox()
                lst.Left = 12 : lst.Top = 12
                lst.Width = 380 : lst.Height = 170
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
                btnOK.Left = 200 : btnOK.Top = 195
                btnOK.Width = 90 : btnOK.Height = 30
                btnOK.DialogResult = DialogResult.OK

                Dim btnCancel As New Button()
                btnCancel.Text = "Hủy"
                btnCancel.Left = 300 : btnCancel.Top = 195
                btnCancel.Width = 90 : btnCancel.Height = 30
                btnCancel.DialogResult = DialogResult.Cancel

                frm.Controls.Add(lst)
                frm.Controls.Add(btnOK)
                frm.Controls.Add(btnCancel)
                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel
                frm.KeyPreview = True

                AddHandler lst.DoubleClick,
                    Sub(s, e)
                        frm.DialogResult = DialogResult.OK
                        frm.Close()
                    End Sub

                AddHandler frm.KeyDown,
                    Sub(s, e)
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