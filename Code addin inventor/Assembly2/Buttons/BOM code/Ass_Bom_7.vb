Imports Inventor
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode

    Public Module ass_bom_7

        Public Sub OnExecute(ByVal Context As NameValueMap)
            RunBOMReplace()
        End Sub

        '=====================================================
        ' HÀM CHÍNH - 1 NÚT
        '=====================================================
        Public Sub RunBOMReplace()

            Dim resultType As DialogResult = MessageBox.Show(
        "Chọn loại muốn sửa:" & vbCrLf & vbCrLf &
        "Yes = Part Number" & vbCrLf &
        "No  = Stock Number",
        "Chọn loại",
        MessageBoxButtons.YesNoCancel,
        MessageBoxIcon.Question)

            If resultType = DialogResult.Cancel Then Exit Sub

            Dim updatePartNumber As Boolean = (resultType = DialogResult.Yes)
            Dim updateStockNumber As Boolean = (resultType = DialogResult.No)

            '----- 2. Chọn Top Level / All Levels (dùng nút) -----
            Dim result As DialogResult = MessageBox.Show(
                "Chọn chế độ chạy:" & vbCrLf & vbCrLf &
                "Yes  = All Levels (sửa tất cả cấp)" & vbCrLf &
                "No   = Top Level (chỉ cấp trên + Phantom)",
                "Chọn Level",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question)

            If result = DialogResult.Cancel Then Exit Sub

            Dim allLevels As Boolean = (result = DialogResult.Yes)

            '----- 3. Nhập chữ tìm & thay -----
            Dim findText As String = InputBox("Nhập chữ cần tìm:", "Find", "mm")
            If String.IsNullOrEmpty(findText) Then Exit Sub

            Dim replaceText As String = InputBox("Nhập chữ cần thay thế:", "Replace", "L")
            If replaceText Is Nothing Then Exit Sub

            '----- 4. Chạy -----
            RunOnModelBrowser(updatePartNumber, updateStockNumber, findText, replaceText, allLevels)
        End Sub

        '=====================================================
        ' LẤY INVENTOR APPLICATION
        '=====================================================
        Private Function GetInventorApp() As Inventor.Application
            Try
                Return DirectCast(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch
                MessageBox.Show("Không tìm thấy Inventor đang chạy!")
                Return Nothing
            End Try
        End Function

        '=====================================================
        ' CHẠY TRÊN MODEL BROWSER
        '=====================================================
        Private Sub RunOnModelBrowser(updatePartNumber As Boolean,
                                      updateStockNumber As Boolean,
                                      findText As String,
                                      replaceText As String,
                                      allLevels As Boolean)

            Dim invApp As Inventor.Application = GetInventorApp()
            If invApp Is Nothing Then Exit Sub

            If invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Hãy mở Assembly trước!")
                Exit Sub
            End If

            Dim asmDoc As AssemblyDocument = invApp.ActiveDocument
            Dim countChanged As Integer = 0

            ' Duyệt từ cấp trên cùng của Model Browser
            For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                ProcessOccurrence(occ, updatePartNumber, updateStockNumber, findText, replaceText, allLevels, countChanged)
            Next

            asmDoc.Update2(True)
            asmDoc.Save2(True)

            MessageBox.Show("Hoàn tất!" & vbCrLf &
                            "Số lượng item đã sửa: " & countChanged.ToString() & vbCrLf &
                            "Tìm: """ & findText & """ → Thay: """ & replaceText & """" & vbCrLf &
                            "Chế độ: " & If(allLevels, "All Levels", "Top Level"),
                            "BOM Replace")
        End Sub

        '=====================================================
        ' XỬ LÝ TỪNG OCCURRENCE (HỖ TRỢ PHANTOM)
        '=====================================================
        Private Sub ProcessOccurrence(occ As ComponentOccurrence,
                                      updatePartNumber As Boolean,
                                      updateStockNumber As Boolean,
                                      findText As String,
                                      replaceText As String,
                                      allLevels As Boolean,
                                      ByRef countChanged As Integer)

            Try
                If occ.Suppressed Then Exit Sub

                Dim doc As Document = Nothing
                Try
                    doc = occ.Definition.Document
                Catch
                    Exit Sub
                End Try

                '----- Kiểm tra Phantom -----
                Dim isPhantom As Boolean = False
                Try
                    ' Phantom thường là kPhantomLevelOfDetail hoặc BOMStructure = kPhantomBOMStructure
                    If occ.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                        isPhantom = True
                    End If
                Catch
                End Try

                ' Nếu là Phantom → luôn đi vào bên trong
                If isPhantom Then
                    If occ.SubOccurrences IsNot Nothing Then
                        For Each subOcc As ComponentOccurrence In occ.SubOccurrences
                            ProcessOccurrence(subOcc, updatePartNumber, updateStockNumber, findText, replaceText, allLevels, countChanged)
                        Next
                    End If
                    Exit Sub   ' Phantom không sửa chính nó
                End If

                '----- Sửa thuộc tính -----
                UpdateDocumentProps(doc, updatePartNumber, updateStockNumber, findText, replaceText, countChanged)

                '----- All Levels: đi tiếp vào cấp con -----
                If allLevels Then
                    If occ.SubOccurrences IsNot Nothing Then
                        For Each subOcc As ComponentOccurrence In occ.SubOccurrences
                            ProcessOccurrence(subOcc, updatePartNumber, updateStockNumber, findText, replaceText, allLevels, countChanged)
                        Next
                    End If
                End If

            Catch
            End Try
        End Sub

        '=====================================================
        ' SỬA iPROPERTIES CỦA DOCUMENT
        '=====================================================
        Private Sub UpdateDocumentProps(doc As Document,
                                        updatePartNumber As Boolean,
                                        updateStockNumber As Boolean,
                                        findText As String,
                                        replaceText As String,
                                        ByRef countChanged As Integer)

            Try
                Dim designProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim changed As Boolean = False

                ' Part Number
                If updatePartNumber Then
                    Try
                        Dim prop As Inventor.Property = designProps.Item("Part Number")
                        Dim oldVal As String = prop.Value.ToString()

                        If oldVal.IndexOf(findText, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Dim newVal As String = Regex.Replace(oldVal, Regex.Escape(findText), replaceText, RegexOptions.IgnoreCase)
                            prop.Value = newVal
                            changed = True
                        End If
                    Catch
                    End Try
                End If

                ' Stock Number
                If updateStockNumber Then
                    Try
                        Dim prop As Inventor.Property = designProps.Item("Stock Number")
                        Dim oldVal As String = prop.Value.ToString()

                        If oldVal.IndexOf(findText, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Dim newVal As String = Regex.Replace(oldVal, Regex.Escape(findText), replaceText, RegexOptions.IgnoreCase)
                            prop.Value = newVal
                            changed = True
                        End If
                    Catch
                    End Try
                End If

                If changed Then
                    doc.Save2(True)
                    countChanged += 1
                End If

            Catch
            End Try
        End Sub

    End Module

End Namespace