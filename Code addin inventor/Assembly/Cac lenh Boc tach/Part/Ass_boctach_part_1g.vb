Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.caclenhboctach.part
    Public Module Ass_boctach_part_1g





        Public Sub OnExecute(ByVal Context As NameValueMap)



            Dim oApp As Inventor.Application = g_inventorApplication

            Dim selectedFile As String = ""

            Dim activeDoc As Document = Nothing

            Try

                activeDoc = oApp.ActiveDocument

            Catch
            End Try


            '=========================================================
            ' BƯỚC 1: CHỌN FILE ASSEMBLY
            '=========================================================

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

                    selectedFile =
                            activeDoc.FullFileName

                Else

                    selectedFile =
                            SelectAssemblyFile()

                End If

            Else

                selectedFile =
                        SelectAssemblyFile()

            End If


            '=========================================================
            ' HỦY
            '=========================================================

            If String.IsNullOrWhiteSpace(selectedFile) Then
                Exit Sub
            End If


            '=========================================================
            ' BƯỚC 2: MỞ ASSEMBLY GỐC
            '=========================================================

            Dim oOrigAsmDoc As AssemblyDocument = Nothing

            Dim openedByCode As Boolean = False


            Try

                If activeDoc IsNot Nothing AndAlso
                       String.Equals(
                           selectedFile,
                           activeDoc.FullFileName,
                           StringComparison.OrdinalIgnoreCase) Then

                    oOrigAsmDoc =
                            CType(
                                activeDoc,
                                AssemblyDocument)

                Else

                    oOrigAsmDoc =
                            CType(
                                oApp.Documents.Open(
                                    selectedFile,
                                    False),
                                AssemblyDocument)

                    openedByCode = True

                End If


                '=====================================================
                ' BƯỚC 3: TẠO ASSEMBLY MỚI
                '=====================================================

                Dim oNewAsmDoc As AssemblyDocument =
                        CType(
                            oApp.Documents.Add(
                                DocumentTypeEnum.kAssemblyDocumentObject),
                            AssemblyDocument)


                Dim baseName As String =
                        IO.Path.GetFileNameWithoutExtension(
                            selectedFile)


                oNewAsmDoc.DisplayName =
                        baseName &
                        "_SheetMetal_Unfold"


                '=====================================================
                ' BƯỚC 4:
                ' THU THẬP TẤT CẢ SHEET METAL
                '
                ' Bao gồm:
                '
                ' MAIN
                '   ├── PART
                '   ├── SUB ASM
                '   │     ├── PART
                '   │     └── SUB ASM
                '   │           └── PART
                '
                '=====================================================

                Dim sheetMetalParts As New HashSet(Of String)(
                        StringComparer.OrdinalIgnoreCase)


                Dim visitedAssemblies As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                CollectSheetMetalParts(
                        oOrigAsmDoc.ComponentDefinition,
                        sheetMetalParts,
                        visitedAssemblies)


                '=====================================================
                ' KIỂM TRA SỐ LƯỢNG
                '=====================================================

                If sheetMetalParts.Count = 0 Then

                    MessageBox.Show(
                            "Không tìm thấy Sheet Metal Part nào trong Assembly.",
                            "Sheet Metal Unfold",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

                    oNewAsmDoc.Close(True)

                    If openedByCode Then
                        oOrigAsmDoc.Close(False)
                    End If

                    Exit Sub

                End If


                '=====================================================
                ' BƯỚC 5:
                ' THÊM TẤT CẢ PART TẠI ORIGIN
                '=====================================================

                Dim oTG As TransientGeometry =
                        oApp.TransientGeometry


                Dim oMatrix As Matrix =
                        oTG.CreateMatrix()


                Dim counter As Integer = 0


                For Each partFullName As String
                        In sheetMetalParts


                    Try

                        '=================================================
                        ' THÊM OCCURRENCE
                        '=================================================

                        Dim oNewOcc As ComponentOccurrence =
                                oNewAsmDoc.ComponentDefinition.Occurrences.Add(
                                    partFullName,
                                    oMatrix)


                        '=================================================
                        ' LẤY PART NUMBER
                        '=================================================

                        Dim partName As String = ""


                        Dim oPartDoc As PartDocument = Nothing

                        Dim partWasAlreadyOpen As Boolean = False


                        Try

                            '---------------------------------------------
                            ' Kiểm tra Part đã mở chưa
                            '---------------------------------------------

                            Try

                                oPartDoc =
                                        CType(
                                            oApp.Documents.ItemByName(
                                                partFullName),
                                            PartDocument)

                                If oPartDoc IsNot Nothing Then
                                    partWasAlreadyOpen = True
                                End If

                            Catch

                                partWasAlreadyOpen = False

                            End Try


                            '---------------------------------------------
                            ' Nếu chưa mở thì mở
                            '---------------------------------------------

                            If Not partWasAlreadyOpen Then

                                oPartDoc =
                                        CType(
                                            oApp.Documents.Open(
                                                partFullName,
                                                False),
                                            PartDocument)

                            End If


                            '---------------------------------------------
                            ' Lấy Part Number
                            '---------------------------------------------

                            Try

                                partName =
                                        CStr(
                                            oPartDoc.PropertySets(
                                                "Design Tracking Properties").
                                                Item("Part Number").Value)

                            Catch

                                partName = ""

                            End Try


                            '---------------------------------------------
                            ' Không có Part Number
                            '---------------------------------------------

                            If String.IsNullOrWhiteSpace(partName) Then

                                partName =
                                        IO.Path.GetFileNameWithoutExtension(
                                            partFullName)

                            End If


                            '---------------------------------------------
                            ' Đặt tên Occurrence
                            '---------------------------------------------

                            If Not String.IsNullOrWhiteSpace(partName) Then

                                Try

                                    oNewOcc.Name =
                                            partName

                                Catch

                                End Try

                            End If


                        Catch

                            ' Nếu không lấy được Part Number
                            ' vẫn giữ occurrence


                        Finally

                            '---------------------------------------------
                            ' Chỉ đóng nếu code đã mở Part
                            '---------------------------------------------

                            If oPartDoc IsNot Nothing AndAlso
                                   Not partWasAlreadyOpen Then

                                Try
                                    oPartDoc.Close(False)
                                Catch
                                End Try

                            End If

                        End Try


                        counter += 1


                    Catch ex As Exception

                        '---------------------------------------------
                        ' Bỏ qua Part lỗi
                        '---------------------------------------------

                    End Try


                Next


                '=====================================================
                ' BƯỚC 6: UPDATE
                '=====================================================

                oNewAsmDoc.Update2(True)


                '=====================================================
                ' BƯỚC 7: CHỌN NƠI LƯU
                '=====================================================

                Dim savePath As String =
                        SelectSaveAssemblyFile(
                            baseName &
                            "_SheetMetal_Unfold.iam")


                '=====================================================
                ' HỦY LƯU
                '=====================================================

                If String.IsNullOrWhiteSpace(savePath) Then

                    MessageBox.Show(
                            "Đã hủy lưu file." &
                            vbCrLf & vbCrLf &
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
                            "Số Sheet Metal Part: " &
                            counter.ToString() &
                            vbCrLf &
                            "Tất cả được đặt tại Origin.",
                            "Sheet Metal Unfold",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

                End If


                '=====================================================
                ' ĐÓNG ASSEMBLY GỐC
                ' NẾU CODE ĐÃ MỞ NÓ
                '=====================================================

                If openedByCode AndAlso
                       oOrigAsmDoc IsNot Nothing Then

                    Try

                        oOrigAsmDoc.Close(False)

                    Catch

                    End Try

                End If


                '=====================================================
                ' ACTIVE ASSEMBLY MỚI
                '=====================================================

                Try

                    oNewAsmDoc.Activate()

                Catch

                End Try


            Catch ex As Exception

                MessageBox.Show(
                        "CÓ LỖI:" &
                        vbCrLf & vbCrLf &
                        ex.Message &
                        vbCrLf & vbCrLf &
                        ex.StackTrace,
                        "Sheet Metal Unfold",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)


                '=====================================================
                ' ĐÓNG FILE GỐC NẾU CODE MỞ
                '=====================================================

                If openedByCode AndAlso
                       oOrigAsmDoc IsNot Nothing Then

                    Try
                        oOrigAsmDoc.Close(False)
                    Catch
                    End Try

                End If

            End Try

        End Sub


        '=============================================================
        ' CHỌN ASSEMBLY
        '=============================================================

        Private Function SelectAssemblyFile() As String

            Try

                Using dlg As New OpenFileDialog()

                    dlg.Title =
                            "Chọn file lắp ghép (.iam)"

                    dlg.Filter =
                            "Assembly Files (*.iam)|*.iam"

                    dlg.Multiselect = False


                    If dlg.ShowDialog() =
                           DialogResult.OK Then

                        Return dlg.FileName

                    End If

                End Using


            Catch ex As Exception

                MessageBox.Show(
                        "Lỗi chọn Assembly:" &
                        vbCrLf &
                        ex.Message,
                        "Lỗi")

            End Try


            Return ""

        End Function


        '=============================================================
        ' CHỌN NƠI LƯU ASSEMBLY
        '=============================================================

        Private Function SelectSaveAssemblyFile(
                ByVal defaultFileName As String) As String

            Try

                Using dlg As New SaveFileDialog()

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

                End Using


            Catch ex As Exception

                MessageBox.Show(
                        "Lỗi chọn nơi lưu:" &
                        vbCrLf &
                        ex.Message,
                        "Lỗi")

            End Try


            Return ""

        End Function


        '=============================================================
        ' RECURSIVE:
        ' THU THẬP TẤT CẢ SHEET METAL PART
        '
        ' Duyệt toàn bộ Assembly + Sub Assembly
        '
        '=============================================================

        Private Sub CollectSheetMetalParts(
                ByVal asmDef As AssemblyComponentDefinition,
                ByRef sheetMetalParts As HashSet(Of String),
                ByRef visitedAssemblies As HashSet(Of String))


            For Each oOcc As ComponentOccurrence
                    In asmDef.Occurrences


                Try

                    '=================================================
                    ' BỎ QUA SUPPRESSED
                    '=================================================

                    If oOcc.Suppressed Then
                        Continue For
                    End If


                    '=================================================
                    ' LẤY DOCUMENT
                    '=================================================

                    Dim oRefDoc As Document = Nothing


                    ' Prefer safe null checks to avoid repeated NullReferenceExceptions
                    If oOcc.ReferencedDocumentDescriptor IsNot Nothing Then
                        Try
                            oRefDoc =
                                    oOcc.ReferencedDocumentDescriptor.
                                    ReferencedDocument
                        Catch
                            oRefDoc = Nothing
                        End Try
                    Else
                        Try
                            oRefDoc =
                                    oOcc.Definition.Document
                        Catch
                            oRefDoc = Nothing
                        End Try
                    End If

                    If oRefDoc Is Nothing Then
                        Continue For
                    End If

                    ' Access FullFileName inside Try/Catch because COM properties can throw
                    Dim refFullName As String = String.Empty
                    Try
                        refFullName = oRefDoc.FullFileName
                    Catch
                        refFullName = String.Empty
                    End Try

                    If String.IsNullOrWhiteSpace(refFullName) Then
                        Continue For
                    End If


                    '=================================================
                    ' NẾU LÀ PART
                    '=================================================

                    If oRefDoc.DocumentType =
                           DocumentTypeEnum.kPartDocumentObject Then


                        Dim oPartDoc As PartDocument =
                                TryCast(
                                    oRefDoc,
                                    PartDocument)


                        If oPartDoc Is Nothing Then
                            Continue For
                        End If


                        '---------------------------------------------
                        ' Kiểm tra Sheet Metal
                        '---------------------------------------------

                        Dim oPartCD As PartComponentDefinition =
                                oPartDoc.ComponentDefinition


                        Dim oSMCD As SheetMetalComponentDefinition =
                                TryCast(
                                    oPartCD,
                                    SheetMetalComponentDefinition)


                        If oSMCD IsNot Nothing Then

                            If oSMCD.Features.Count > 0 Then

                                sheetMetalParts.Add(
                                        oPartDoc.FullFileName)

                            End If

                        End If


                        '=================================================
                        ' NẾU LÀ SUB ASSEMBLY
                        '=================================================

                    ElseIf oRefDoc.DocumentType =
                               DocumentTypeEnum.kAssemblyDocumentObject Then


                        Dim oSubAsmDef As AssemblyComponentDefinition =
                                TryCast(
                                    oRefDoc.ComponentDefinition,
                                    AssemblyComponentDefinition)


                        If oSubAsmDef IsNot Nothing Then

                            ' Prevent infinite recursion by checking visited assemblies
                            Try
                                Dim subAsmFullName As String = String.Empty
                                Try
                                    subAsmFullName = oRefDoc.FullFileName
                                Catch
                                    subAsmFullName = String.Empty
                                End Try

                                If Not String.IsNullOrWhiteSpace(subAsmFullName) Then
                                    If Not visitedAssemblies.Contains(subAsmFullName) Then
                                        visitedAssemblies.Add(subAsmFullName)
                                        '-----------------------------------------
                                        ' Đệ quy xuống Level tiếp theo
                                        '-----------------------------------------
                                        CollectSheetMetalParts(
                                                oSubAsmDef,
                                                sheetMetalParts,
                                                visitedAssemblies)
                                    End If
                                End If
                            Catch
                                ' ignore recursion errors per original behavior
                            End Try

                        End If


                    End If


                Catch

                    ' Bỏ qua occurrence lỗi

                End Try


            Next

        End Sub

    End Module

End Namespace


