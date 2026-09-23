Option Explicit On
Imports Inventor
Imports System.Windows.Forms
Imports System.Drawing

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep

    Public Module Ass_LG_3

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application =
                Interop.Marshal2.GetActiveObject("Inventor.Application")

            If invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Vui lòng mở Assembly trước!", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            Dim asmDoc As AssemblyDocument = invApp.ActiveEditDocument
            Dim oDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

            '=====================================================
            ' FORM
            '=====================================================
            Dim frm As New Form()
            frm.Text = "Ẩn Component - Inventor 2020"
            frm.Size = New Size(380, 440)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.Font = New Font("Segoe UI", 9)

            Dim btn1 As New Button() With {.Text = "1. Ẩn Referent", .Location = New System.Drawing.Point(40, 20), .Size = New Size(280, 35)}
            Dim btn2 As New Button() With {.Text = "2. Ẩn Phantom", .Location = New System.Drawing.Point(40, 60), .Size = New Size(280, 35)}
            Dim btn3 As New Button() With {.Text = "3. Ẩn Purchased (đồ mua)", .Location = New System.Drawing.Point(40, 100), .Size = New Size(280, 35)}
            Dim btn4 As New Button() With {.Text = "4. Ẩn Weldment (cụm hàn)", .Location = New System.Drawing.Point(40, 140), .Size = New Size(280, 35)}
            Dim btn5 As New Button() With {.Text = "5. Ẩn Part", .Location = New System.Drawing.Point(40, 180), .Size = New Size(280, 35)}
            Dim btn6 As New Button() With {.Text = "6. Ẩn Sheet Metal", .Location = New System.Drawing.Point(40, 220), .Size = New Size(280, 35)}
            Dim btn7 As New Button() With {.Text = "7. Ẩn TẤT CẢ các loại trên", .Location = New System.Drawing.Point(40, 260), .Size = New Size(280, 35)}
            Dim btn8 As New Button() With {.Text = "8. Hiện tất cả + chỉ ẩn Referent", .Location = New System.Drawing.Point(40, 300), .Size = New Size(280, 35), .BackColor = System.Drawing.Color.LightYellow}
            Dim btn0 As New Button() With {.Text = "0. Hiện lại tất cả", .Location = New System.Drawing.Point(40, 340), .Size = New Size(280, 35), .BackColor = System.Drawing.Color.LightGreen}

            AddHandler btn1.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "Referent")
                                       MessageBox.Show("Đã ẩn " & c & " Referent", "Hoàn tất")
                                       frm.Close()
                                   End Sub
            AddHandler btn2.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "Phantom")
                                       MessageBox.Show("Đã ẩn " & c & " Phantom", "Hoàn tất")
                                       frm.Close()
                                   End Sub
            AddHandler btn3.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "Purchased")
                                       MessageBox.Show("Đã ẩn " & c & " Purchased", "Hoàn tất")
                                       frm.Close()
                                   End Sub
            AddHandler btn4.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "Weldment")
                                       MessageBox.Show("Đã ẩn " & c & " Weldment", "Hoàn tất")
                                       frm.Close()
                                   End Sub
            AddHandler btn5.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "Part")
                                       MessageBox.Show("Đã ẩn " & c & " Part", "Hoàn tất")
                                       frm.Close()
                                   End Sub
            AddHandler btn6.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "SheetMetal")
                                       MessageBox.Show("Đã ẩn " & c & " Sheet Metal", "Hoàn tất")
                                       frm.Close()
                                   End Sub
            AddHandler btn7.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "All")
                                       MessageBox.Show("Đã ẩn " & c & " component", "Hoàn tất")
                                       frm.Close()
                                   End Sub
            AddHandler btn8.Click, Sub()
                                       ShowAll(oDef)
                                       Dim c As Integer = HideByType(oDef, "Referent")
                                       MessageBox.Show("Đã hiện tất cả + ẩn " & c & " Referent", "Hoàn tất")
                                       frm.Close()
                                   End Sub
            AddHandler btn0.Click, Sub()
                                       ShowAll(oDef)
                                       MessageBox.Show("Đã hiện lại tất cả component!", "Hoàn tất")
                                       frm.Close()
                                   End Sub

            frm.Controls.AddRange({btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8, btn0})
            frm.ShowDialog()

            asmDoc.Update2(True)

        End Sub

        '=====================================================
        ' ẨN THEO LOẠI (ĐỆ QUY TOÀN BỘ)
        '=====================================================
        Private Function HideByType(ByVal oDef As AssemblyComponentDefinition, ByVal mode As String) As Integer

            Dim count As Integer = 0
            HideRecursive(oDef.Occurrences, mode, count)
            Return count

        End Function

        Private Sub HideRecursive(ByVal occs As ComponentOccurrences, ByVal mode As String, ByRef count As Integer)

            For Each occ As ComponentOccurrence In occs
                Try
                    If occ.Suppressed Then Continue For

                    Dim shouldHide As Boolean = False

                    Select Case mode
                        Case "Referent"
                            shouldHide = IsReferent(occ)
                        Case "Phantom"
                            shouldHide = IsPhantom(occ)
                        Case "Purchased"
                            shouldHide = IsPurchased(occ)
                        Case "Weldment"
                            shouldHide = IsWeldment(occ)
                        Case "Part"
                            shouldHide = IsPart(occ)
                        Case "SheetMetal"
                            shouldHide = IsSheetMetal(occ)
                        Case "All"
                            shouldHide = IsReferent(occ) OrElse IsPhantom(occ) OrElse
                                         IsPurchased(occ) OrElse IsWeldment(occ) OrElse
                                         IsPart(occ) OrElse IsSheetMetal(occ)
                    End Select

                    If shouldHide Then
                        occ.Visible = False
                        count += 1
                    End If

                    ' Đệ quy vào sub-assembly
                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Try
                            Dim subDef As AssemblyComponentDefinition =
                                CType(occ.Definition, AssemblyComponentDefinition)
                            HideRecursive(subDef.Occurrences, mode, count)
                        Catch
                        End Try
                    End If

                Catch
                End Try
            Next

        End Sub

        '=====================================================
        ' HIỆN LẠI TẤT CẢ (ĐỆ QUY)
        '=====================================================
        Private Sub ShowAll(ByVal oDef As AssemblyComponentDefinition)
            ShowRecursive(oDef.Occurrences)
        End Sub

        Private Sub ShowRecursive(ByVal occs As ComponentOccurrences)
            For Each occ As ComponentOccurrence In occs
                Try
                    occ.Visible = True
                    If occ.Suppressed Then occ.Unsuppress()

                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Try
                            Dim subDef As AssemblyComponentDefinition =
                                CType(occ.Definition, AssemblyComponentDefinition)
                            ShowRecursive(subDef.Occurrences)
                        Catch
                        End Try
                    End If
                Catch
                End Try
            Next
        End Sub

        '=====================================================
        ' 1. REFERENT  (đã sửa mạnh)
        '=====================================================
        Private Function IsReferent(ByVal occ As ComponentOccurrence) As Boolean
            Try
                ' Ưu tiên BOM Structure
                If occ.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then Return True

                ' Content Center
                If occ.IsContentMember Then Return True

                ' Part Reference
                If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim pDoc As PartDocument = TryCast(occ.Definition.Document, PartDocument)
                    If pDoc IsNot Nothing AndAlso pDoc.ComponentDefinition.IsReferencePart Then
                        Return True
                    End If
                End If

            Catch
            End Try
            Return False
        End Function

        '=====================================================
        ' 2. PHANTOM
        '=====================================================
        Private Function IsPhantom(ByVal occ As ComponentOccurrence) As Boolean
            Try
                Return (occ.BOMStructure = BOMStructureEnum.kPhantomBOMStructure)
            Catch
            End Try
            Return False
        End Function

        '=====================================================
        ' 3. PURCHASED
        '=====================================================
        Private Function IsPurchased(ByVal occ As ComponentOccurrence) As Boolean
            Try
                If occ.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then Return True

                Dim doc As Document = occ.Definition.Document
                Dim designProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")

                Dim desc As String = ""
                Try
                    desc = designProps.Item("Description").Value.ToString()
                Catch
                End Try

                If LCase(desc).Contains("purchased") OrElse LCase(desc).Contains("đồ mua") Then
                    Return True
                End If

            Catch
            End Try
            Return False
        End Function

        '=====================================================
        ' 4. WELDMENT
        '=====================================================
        Private Function IsWeldment(ByVal occ As ComponentOccurrence) As Boolean
            Try
                If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim aDoc As AssemblyDocument = TryCast(occ.Definition.Document, AssemblyDocument)
                    If aDoc IsNot Nothing AndAlso aDoc.ComponentDefinition.IsWeldment Then
                        Return True
                    End If
                End If
            Catch
            End Try
            Return False
        End Function

        '=====================================================
        ' 5. PART THƯỜNG
        ' Điều kiện:  Là PartComponentDefinition
        '          + KHÔNG phải Sheet Metal
        '          + KHÔNG phải Reference Part
        '          + KHÔNG phải Content Center member
        '          + BOM không thuộc {Reference, Phantom, Purchased}
        '=====================================================
        Private Function IsPart(ByVal occ As ComponentOccurrence) As Boolean
            If occ Is Nothing Then Return False

            Try
                ' Bỏ qua occurrence đang suppress (tránh exception khi truy cập Definition)
                If occ.Suppressed Then Return False

                ' 1) Phải là Part (không phải Assembly/IAM)
                Dim pDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
                If pDef Is Nothing Then Return False

                ' 2) Loại trừ Sheet Metal (xử lý riêng ở IsSheetMetal)
                If pDef.IsSheetMetal Then Return False

                ' 3) Loại trừ Reference Part (part được tạo dạng tham chiếu)
                If SafeIsReferencePart(pDef) Then Return False

                ' 4) Loại trừ Content Center member (bulông, đai ốc... thư viện chuẩn)
                If SafeIsContentMember(occ) Then Return False

                ' 5) Loại trừ theo BOM Structure
                Select Case occ.BOMStructure
                    Case BOMStructureEnum.kReferenceBOMStructure,
                 BOMStructureEnum.kPhantomBOMStructure,
                 BOMStructureEnum.kPurchasedBOMStructure
                        Return False
                End Select

                ' Đủ điều kiện = Part thường
                Return True

            Catch
                Return False
            End Try
        End Function

        '=====================================================
        ' 6. SHEET METAL
        ' Điều kiện:  Là PartComponentDefinition
        '          + LÀ Sheet Metal
        '          + KHÔNG phải Reference Part
        '          + KHÔNG phải Content Center member
        '          + BOM không thuộc {Reference, Phantom, Purchased}
        '=====================================================
        Private Function IsSheetMetal(ByVal occ As ComponentOccurrence) As Boolean
            If occ Is Nothing Then Return False

            Try
                ' Bỏ qua occurrence đang suppress
                If occ.Suppressed Then Return False

                ' 1) Phải là Part
                Dim pDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
                If pDef Is Nothing Then Return False

                ' 2) BẮT BUỘC phải là Sheet Metal — check sớm để tránh truy cập prop khác
                If Not pDef.IsSheetMetal Then Return False

                ' 3) Loại trừ Reference Part
                If SafeIsReferencePart(pDef) Then Return False

                ' 4) Loại trừ Content Center member
                If SafeIsContentMember(occ) Then Return False

                ' 5) Loại trừ theo BOM Structure
                Select Case occ.BOMStructure
                    Case BOMStructureEnum.kReferenceBOMStructure,
                 BOMStructureEnum.kPhantomBOMStructure,
                 BOMStructureEnum.kPurchasedBOMStructure
                        Return False
                End Select

                ' Đủ điều kiện = Sheet Metal
                Return True

            Catch
                Return False
            End Try
        End Function

        '=====================================================
        ' HELPERS — bọc try/catch riêng để 1 prop lỗi không làm hỏng cả hàm
        '=====================================================
        Private Function SafeIsReferencePart(ByVal pDef As PartComponentDefinition) As Boolean
            If pDef Is Nothing Then Return False
            Try
                Return pDef.IsReferencePart
            Catch
                Return False
            End Try
        End Function

        Private Function SafeIsContentMember(ByVal occ As ComponentOccurrence) As Boolean
            If occ Is Nothing Then Return False
            Try
                Return True = occ.IsContentMember
            Catch
                ' Property này đôi khi ném exception với occurrence chưa resolve
                Return False
            End Try
        End Function
    End Module
End Namespace