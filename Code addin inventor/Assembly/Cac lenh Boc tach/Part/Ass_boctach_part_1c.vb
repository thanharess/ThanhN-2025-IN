Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.caclenhboctach.part
    Public Module Ass_boctach_part_1c
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim oApp As Inventor.Application = g_inventorApplication
            Dim selectedFile As String = ""

            Try

                '=====================================================
                ' BƯỚC 1: CHỌN FILE ASSEMBLY
                '=====================================================

                Dim activeDoc As Document = oApp.ActiveDocument

                If activeDoc IsNot Nothing AndAlso
                   activeDoc.DocumentType =
                   DocumentTypeEnum.kAssemblyDocumentObject Then

                    Dim result As DialogResult =
                        MessageBox.Show(
                            "Sử dụng file lắp ghép đang mở?",
                            "Chọn file",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question)

                    If result = DialogResult.Yes Then

                        selectedFile = activeDoc.FullFileName

                    Else

                        selectedFile = SelectAssemblyFile(oApp)

                    End If

                Else

                    selectedFile = SelectAssemblyFile(oApp)

                End If


                '=====================================================
                ' HỦY
                '=====================================================

                If String.IsNullOrWhiteSpace(selectedFile) Then
                    Exit Sub
                End If


                '=====================================================
                ' BƯỚC 2: MỞ ASSEMBLY GỐC
                '=====================================================

                Dim oOrigAsmDoc As AssemblyDocument = Nothing

                If activeDoc IsNot Nothing AndAlso
                   String.Equals(
                       selectedFile,
                       activeDoc.FullFileName,
                       StringComparison.OrdinalIgnoreCase) Then

                    oOrigAsmDoc =
                        CType(activeDoc, AssemblyDocument)

                Else

                    oOrigAsmDoc =
                        CType(
                            oApp.Documents.Open(
                                selectedFile,
                                False),
                            AssemblyDocument)

                End If


                '=====================================================
                ' BƯỚC 3: TẠO ASSEMBLY MỚI
                '=====================================================

                Dim oNewAsmDoc As AssemblyDocument =
                    CType(
                        oApp.Documents.Add(
                            DocumentTypeEnum.kAssemblyDocumentObject,
                            "",
                            True),
                        AssemblyDocument)


                '=====================================================
                ' TÊN DISPLAY
                '=====================================================

                Dim baseName As String =
                    IO.Path.GetFileNameWithoutExtension(
                        selectedFile)

                oNewAsmDoc.DisplayName =
                    baseName &
                    "_SheetMetal_Unfold"


                '=====================================================
                ' BƯỚC 4: DUYỆT ASSEMBLY
                '=====================================================

                Dim oAsmDef As AssemblyComponentDefinition =
                    oOrigAsmDoc.ComponentDefinition


                Dim counter As Integer = 0


                '=====================================================
                ' MATRIX = ORIGIN
                '=====================================================

                Dim oTG As TransientGeometry =
                    oApp.TransientGeometry

                Dim oMatrix As Matrix =
                    oTG.CreateMatrix()


                '=====================================================
                ' DUYỆT OCCURRENCES
                '=====================================================

                For Each oOcc As ComponentOccurrence _
                    In oAsmDef.Occurrences


                    Try

                        '---------------------------------------------
                        ' LẤY DOCUMENT
                        '---------------------------------------------

                        Dim oPartDoc As PartDocument =
                            TryCast(
                                oOcc.Definition.Document,
                                PartDocument)


                        If oPartDoc Is Nothing Then
                            Continue For
                        End If


                        '---------------------------------------------
                        ' KIỂM TRA SHEET METAL
                        '---------------------------------------------

                        Dim oPartCD As PartComponentDefinition =
                            oPartDoc.ComponentDefinition


                        Dim oSMCD As SheetMetalComponentDefinition =
                            TryCast(
                                oPartCD,
                                SheetMetalComponentDefinition)


                        If oSMCD Is Nothing Then
                            Continue For
                        End If


                        '---------------------------------------------
                        ' KIỂM TRA CÓ FEATURE
                        '---------------------------------------------

                        If oSMCD.Features.Count <= 0 Then
                            Continue For
                        End If


                        '=================================================
                        ' THÊM PART VÀO ASSEMBLY MỚI
                        ' TẠI ORIGIN
                        '=================================================

                        Dim oNewOcc As ComponentOccurrence =
                            oNewAsmDoc.ComponentDefinition.Occurrences.Add(
                                oPartDoc.FullFileName,
                                oMatrix)


                        '=================================================
                        ' LẤY PART NUMBER
                        '=================================================

                        Dim partName As String = ""


                        Try

                            partName =
                                CStr(
                                    oPartDoc.PropertySets(
                                        "Design Tracking Properties").
                                        Item("Part Number").Value)

                        Catch

                            partName = ""

                        End Try


                        '=================================================
                        ' NẾU KHÔNG CÓ PART NUMBER
                        ' DÙNG DISPLAY NAME
                        '=================================================

                        If String.IsNullOrWhiteSpace(partName) Then

                            partName =
                                IO.Path.GetFileNameWithoutExtension(
                                    oPartDoc.FullFileName)

                        End If


                        '=================================================
                        ' ĐẶT TÊN OCCURRENCE
                        '=================================================

                        If Not String.IsNullOrWhiteSpace(partName) Then

                            Try

                                oNewOcc.Name = partName

                            Catch

                                ' Không set được thì giữ tên mặc định

                            End Try

                        End If


                        counter += 1


                    Catch

                        ' Bỏ qua Part lỗi

                    End Try


                Next


                '=====================================================
                ' BƯỚC 5: UPDATE ASSEMBLY
                '=====================================================

                oNewAsmDoc.Update2(True)


                '=====================================================
                ' BƯỚC 6: CHỌN NƠI LƯU
                '=====================================================

                Dim savePath As String =
                    SelectSaveAssemblyFile(
                        oApp,
                        baseName &
                        "_SheetMetal_Unfold.iam")


                '=====================================================
                ' HỦY LƯU
                '=====================================================

                If String.IsNullOrWhiteSpace(savePath) Then

                    MessageBox.Show(
                        "Đã hủy lưu file." &
                        vbCrLf &
                        "Assembly mới vẫn đang mở.",
                        "Hủy lưu",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

                Else

                    '=================================================
                    ' SAVE AS
                    '=================================================

                    oNewAsmDoc.SaveAs(
                        savePath,
                        False)


                    MessageBox.Show(
                        "ĐÃ HOÀN TẤT!" &
                        vbCrLf & vbCrLf &
                        "File:" &
                        vbCrLf &
                        savePath &
                        vbCrLf & vbCrLf &
                        "Số chi tiết Sheet Metal: " &
                        counter.ToString() &
                        vbCrLf & vbCrLf &
                        "Tất cả được đặt tại Origin.",
                        "Sheet Metal Unfold",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

                End If


                '=====================================================
                ' ĐÓNG ASSEMBLY GỐC NẾU ĐƯỢC MỞ BỞI CODE
                '=====================================================

                If activeDoc Is Nothing OrElse
                   Not Object.ReferenceEquals(
                       oOrigAsmDoc,
                       activeDoc) Then

                    Try
                        oOrigAsmDoc.Close(False)
                    Catch
                    End Try

                End If


                '=====================================================
                ' ACTIVE ASSEMBLY MỚI
                '=====================================================

                oNewAsmDoc.Activate()


            Catch ex As Exception

                MessageBox.Show(
                    "Có lỗi xảy ra:" &
                    vbCrLf & vbCrLf &
                    ex.Message &
                    vbCrLf & vbCrLf &
                    "Chi tiết:" &
                    vbCrLf &
                    ex.StackTrace,
                    "Sheet Metal Unfold",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub


        '=============================================================
        ' CHỌN ASSEMBLY
        '=============================================================

        Private Function SelectAssemblyFile(
            ByVal oApp As Inventor.Application) As String

            Try

                Dim dlg As OpenFileDialog =
                    New OpenFileDialog()


                dlg.Title =
                    "Chọn file lắp ghép (.iam)"


                dlg.Filter =
                    "Assembly Files (*.iam)|*.iam"


                dlg.Multiselect = False


                If dlg.ShowDialog() =
                   DialogResult.OK Then

                    Return dlg.FileName

                End If


            Catch ex As Exception

                MessageBox.Show(
                    "Không thể mở hộp thoại chọn file:" &
                    vbCrLf &
                    ex.Message,
                    "Lỗi")

            End Try


            Return ""

        End Function


        '=============================================================
        ' CHỌN FILE SAVE
        '=============================================================

        Private Function SelectSaveAssemblyFile(
            ByVal oApp As Inventor.Application,
            ByVal defaultFileName As String) As String

            Try

                Dim dlg As SaveFileDialog =
                    New SaveFileDialog()


                dlg.Title =
                    "Lưu file lắp ghép mới (.iam)"


                dlg.Filter =
                    "Assembly Files (*.iam)|*.iam"


                dlg.DefaultExt = "iam"


                dlg.AddExtension = True


                dlg.FileName =
                    defaultFileName


                If dlg.ShowDialog() =
                   DialogResult.OK Then

                    Return dlg.FileName

                End If


            Catch ex As Exception

                MessageBox.Show(
                    "Không thể mở hộp thoại lưu file:" &
                    vbCrLf &
                    ex.Message,
                    "Lỗi")

            End Try


            Return ""

        End Function

    End Module

End Namespace



