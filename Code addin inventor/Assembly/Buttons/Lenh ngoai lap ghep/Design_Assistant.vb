Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Lenhngoaicumlap
    Public Module Design_Assistant

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application = g_inventorApplication

            Try

                '=====================================================
                ' KIỂM TRA DOCUMENT
                '=====================================================

                If invApp.ActiveDocument Is Nothing Then

                    MessageBox.Show(
                        "Không có Document đang mở.",
                        "Copy Assembly")

                    Exit Sub

                End If


                If invApp.ActiveDocument.DocumentType <>
                   DocumentTypeEnum.kAssemblyDocumentObject Then

                    MessageBox.Show(
                        "Vui lòng mở Assembly trước!",
                        "Copy Assembly")

                    Exit Sub

                End If


                Dim asmDoc As AssemblyDocument =
                    CType(invApp.ActiveDocument, AssemblyDocument)


                '=====================================================
                ' LẤY THÔNG TIN FILE GỐC
                '=====================================================

                Dim sourceRoot As String =
                    IO.Path.GetDirectoryName(
                        asmDoc.FullFileName)

                Dim oldMainName As String =
                    IO.Path.GetFileNameWithoutExtension(
                        asmDoc.FullFileName)


                '=====================================================
                ' TÊN MAIN ASSEMBLY
                '=====================================================

                Dim newMainName As String =
    InputBoxEx(
        "TÊN GỐC: " & oldMainName & vbCrLf & vbCrLf &
        "Nhập tên MỚI cho file lắp chính:",
        "ĐỔI TÊN FILE LẮP CHÍNH",
        oldMainName & "-2")

                If newMainName Is Nothing Then Exit Sub           ' ← Cancel

                If String.IsNullOrWhiteSpace(newMainName) Then
                    MessageBox.Show("Tên Main Assembly không được để trống.",
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                '=====================================================
                ' PREFIX
                '=====================================================

                '=====================================================
                ' PREFIX
                '=====================================================

                Dim subPrefix As String =
    InputBoxEx(
        "Thêm phần tên đầu cho Cụm phụ & Part (có thể để trống):",
        "PREFIX",
        "")

                If subPrefix Is Nothing Then Exit Sub             ' ← Cancel

                '=====================================================
                ' SUFFIX
                '=====================================================



                Dim subSuffix As String =
    InputBoxEx(
        "Thêm phần tên cuối cho Cụm phụ & Part (có thể để trống):",
        "SUFFIX",
        "-")

                If subSuffix Is Nothing Then Exit Sub             ' ← Cancel


                '=====================================================
                ' CHỌN FOLDER ĐÍCH
                '=====================================================

                Dim targetFolder As String = ""

                Using dlg As New FolderBrowserDialog()

                    dlg.Description =
                        "Chọn thư mục để tạo Assembly mới"

                    dlg.SelectedPath = sourceRoot

                    If dlg.ShowDialog() <>
                       DialogResult.OK Then

                        Exit Sub

                    End If

                    targetFolder = dlg.SelectedPath

                End Using


                '=====================================================
                ' FILE MAP
                '
                ' FILE GỐC -> FILE MỚI
                '=====================================================

                Dim fileMap As New Dictionary(Of String, String)(
                    StringComparer.OrdinalIgnoreCase)


                '=====================================================
                ' FILE ĐÃ DUYỆT
                '=====================================================

                Dim processed As New HashSet(Of String)(
                    StringComparer.OrdinalIgnoreCase)


                '=====================================================
                ' FOLDER MAIN
                '=====================================================

                Dim mainFolder As String =
                    IO.Path.Combine(
                        targetFolder,
                        subPrefix &
                        newMainName &
                        subSuffix)


                '=====================================================
                ' THU THẬP CÂY ASSEMBLY
                '=====================================================

                CollectAssemblyTree(
                    asmDoc,
                    mainFolder,
                    newMainName,
                    subPrefix,
                    subSuffix,
                    fileMap,
                    processed,
                    True)


                '=====================================================
                ' COPY FILE
                '=====================================================

                For Each kvp As KeyValuePair(Of String, String)
                    In fileMap

                    Try

                        Dim destinationFolder As String =
                            IO.Path.GetDirectoryName(
                                kvp.Value)


                        If Not IO.Directory.Exists(
                            destinationFolder) Then

                            IO.Directory.CreateDirectory(
                                destinationFolder)

                        End If


                        IO.File.Copy(
                            kvp.Key,
                            kvp.Value,
                            True)


                    Catch ex As Exception

                        MessageBox.Show(
                            "Không copy được file:" &
                            vbCrLf & vbCrLf &
                            kvp.Key &
                            vbCrLf & vbCrLf &
                            ex.Message,
                            "Lỗi Copy")

                    End Try

                Next


                '=====================================================
                ' KIỂM TRA MAIN
                '=====================================================

                If Not fileMap.ContainsKey(
                    asmDoc.FullFileName) Then

                    Throw New Exception(
                        "Không tìm thấy Main Assembly trong File Map.")

                End If


                Dim newAsmPath As String =
                    fileMap(asmDoc.FullFileName)


                '=====================================================
                ' MỞ MAIN MỚI
                '=====================================================

                Dim newAsm As AssemblyDocument =
                    CType(
                        invApp.Documents.Open(
                            newAsmPath,
                            True),
                        AssemblyDocument)


                '=====================================================
                ' RELINK
                '=====================================================

                ReplaceAllReferences(
                    newAsm,
                    fileMap)


                '=====================================================
                ' UPDATE IPROPERTIES
                '=====================================================

                Dim propertyProcessed As New HashSet(Of String)(
                    StringComparer.OrdinalIgnoreCase)


                UpdateAlliProperties(
                    newAsm,
                    propertyProcessed)


                '=====================================================
                ' UPDATE + SAVE
                '=====================================================

                newAsm.Update2(True)

                newAsm.Save2(True)


                '=====================================================
                ' HOÀN TẤT
                '=====================================================

                MessageBox.Show(
                    "HOÀN TẤT!" &
                    vbCrLf & vbCrLf &
                    "Main Assembly:" &
                    vbCrLf &
                    newAsmPath &
                    vbCrLf & vbCrLf &
                    "Cấu trúc:" &
                    vbCrLf &
                    "• Assembly → Folder riêng" &
                    vbCrLf &
                    "• Part → Không tạo Folder" &
                    vbCrLf &
                    "• Part nằm trong Folder Assembly cha",
                    "Copy Assembly")


            Catch ex As Exception

                MessageBox.Show(
                    "CÓ LỖI:" &
                    vbCrLf & vbCrLf &
                    ex.Message &
                    vbCrLf & vbCrLf &
                    ex.StackTrace,
                    "Copy Assembly")

            End Try

        End Sub


        '=============================================================
        ' COLLECT ASSEMBLY TREE
        '
        ' ASSEMBLY -> TẠO FOLDER
        ' PART     -> KHÔNG TẠO FOLDER
        '
        '=============================================================

        Private Sub CollectAssemblyTree(
            ByVal asm As AssemblyDocument,
            ByVal currentFolder As String,
            ByVal newMainName As String,
            ByVal subPrefix As String,
            ByVal subSuffix As String,
            ByRef fileMap As Dictionary(Of String, String),
            ByRef processed As HashSet(Of String),
            ByVal isMain As Boolean)


            Try

                '=====================================================
                ' TẠO FOLDER ASSEMBLY
                '=====================================================

                If Not IO.Directory.Exists(
                    currentFolder) Then

                    IO.Directory.CreateDirectory(
                        currentFolder)

                End If


                '=====================================================
                ' TÊN ASSEMBLY
                '=====================================================

                Dim asmName As String


                If isMain Then

                    asmName = newMainName

                Else

                    asmName =
                        subPrefix &
                        IO.Path.GetFileNameWithoutExtension(
                            asm.FullFileName) &
                        subSuffix

                End If


                '=====================================================
                ' PATH ASSEMBLY MỚI
                '=====================================================

                Dim newAsmPath As String =
                    IO.Path.Combine(
                        currentFolder,
                        asmName &
                        IO.Path.GetExtension(
                            asm.FullFileName))


                '=====================================================
                ' MAP ASSEMBLY
                '=====================================================

                If Not fileMap.ContainsKey(
                    asm.FullFileName) Then

                    fileMap.Add(
                        asm.FullFileName,
                        newAsmPath)

                End If


                '=====================================================
                ' TRÁNH DUYỆT TRÙNG
                '=====================================================

                If processed.Contains(
                    asm.FullFileName) Then

                    Exit Sub

                End If


                processed.Add(
                    asm.FullFileName)


                '=====================================================
                ' DUYỆT COMPONENT
                '=====================================================

                For Each occ As ComponentOccurrence
                    In asm.ComponentDefinition.Occurrences


                    Try

                        '-------------------------------------------------
                        ' BỎ QUA SUPPRESSED
                        '-------------------------------------------------

                        If occ.Suppressed Then
                            Continue For
                        End If


                        '-------------------------------------------------
                        ' LẤY DOCUMENT
                        '-------------------------------------------------

                        Dim refDoc As Document = Nothing


                        Try

                            refDoc =
                                occ.Definition.Document

                        Catch

                            Continue For

                        End Try


                        If refDoc Is Nothing Then
                            Continue For
                        End If


                        If String.IsNullOrWhiteSpace(
                            refDoc.FullFileName) Then

                            Continue For

                        End If


                        '-------------------------------------------------
                        ' PURCHASED -> BỎ QUA
                        '-------------------------------------------------

                        If IsPurchasedPart(refDoc) Then

                            Continue For

                        End If


                        '-------------------------------------------------
                        ' TÊN FILE
                        '-------------------------------------------------

                        Dim componentName As String =
                            IO.Path.GetFileNameWithoutExtension(
                                refDoc.FullFileName)


                        componentName =
                            subPrefix &
                            componentName &
                            subSuffix


                        '=================================================
                        ' NẾU LÀ ASSEMBLY
                        '=================================================

                        If refDoc.DocumentType =
                           DocumentTypeEnum.kAssemblyDocumentObject Then


                            '---------------------------------------------
                            ' TẠO FOLDER RIÊNG CHO SUB ASSEMBLY
                            '---------------------------------------------

                            Dim subFolder As String =
                                IO.Path.Combine(
                                    currentFolder,
                                    componentName)


                            If Not IO.Directory.Exists(
                                subFolder) Then

                                IO.Directory.CreateDirectory(
                                    subFolder)

                            End If


                            '---------------------------------------------
                            ' PATH FILE SUB ASSEMBLY
                            '---------------------------------------------

                            Dim subFilePath As String =
                                IO.Path.Combine(
                                    subFolder,
                                    componentName &
                                    IO.Path.GetExtension(
                                        refDoc.FullFileName))


                            '---------------------------------------------
                            ' MAP
                            '---------------------------------------------

                            If Not fileMap.ContainsKey(
                                refDoc.FullFileName) Then

                                fileMap.Add(
                                    refDoc.FullFileName,
                                    subFilePath)

                            End If


                            '---------------------------------------------
                            ' ĐI XUỐNG CỤM CON
                            '---------------------------------------------

                            CollectAssemblyTree(
                                CType(
                                    refDoc,
                                    AssemblyDocument),
                                subFolder,
                                newMainName,
                                subPrefix,
                                subSuffix,
                                fileMap,
                                processed,
                                False)


                        Else


                            '=================================================
                            ' PART
                            '
                            ' KHÔNG TẠO FOLDER
                            '=================================================

                            Dim partFilePath As String =
                                IO.Path.Combine(
                                    currentFolder,
                                    componentName &
                                    IO.Path.GetExtension(
                                        refDoc.FullFileName))


                            '---------------------------------------------
                            ' MAP PART
                            '---------------------------------------------

                            If Not fileMap.ContainsKey(
                                refDoc.FullFileName) Then

                                fileMap.Add(
                                    refDoc.FullFileName,
                                    partFilePath)

                            End If


                        End If


                    Catch

                        ' Bỏ qua component lỗi

                    End Try


                Next


            Catch

                ' Bỏ qua Assembly lỗi

            End Try

        End Sub

        '=============================================================
        ' INPUTBOX CÓ PHÂN BIỆT CANCEL vs OK-EMPTY
        '   - Bấm OK     → trả về chuỗi (có thể rỗng)
        '   - Bấm Cancel → trả về Nothing
        '   - Bấm X      → trả về Nothing
        '=============================================================
        Private Function InputBoxEx(
    ByVal prompt As String,
    ByVal title As String,
    ByVal defaultValue As String) As String

            Dim frm As New Form With {
        .Text = title,
        .ClientSize = New System.Drawing.Size(460, 200),
        .StartPosition = FormStartPosition.CenterScreen,
        .FormBorderStyle = FormBorderStyle.FixedDialog,
        .MaximizeBox = False,
        .MinimizeBox = False,
        .ShowInTaskbar = False
    }

            Dim lbl As New Label With {
        .Text = prompt,
        .Bounds = New System.Drawing.Rectangle(12, 12, 436, 90),
        .Font = New System.Drawing.Font("Segoe UI", 9)
    }

            Dim txt As New System.Windows.Forms.TextBox With {
        .Text = If(defaultValue, ""),
        .Bounds = New System.Drawing.Rectangle(12, 105, 436, 24),
        .Font = New System.Drawing.Font("Segoe UI", 9)
    }
            txt.SelectAll()

            Dim btnOK As New Button With {
        .Text = "OK",
        .Bounds = New System.Drawing.Rectangle(270, 150, 85, 30),
        .DialogResult = DialogResult.OK
    }
            Dim btnCancel As New Button With {
        .Text = "Hủy",
        .Bounds = New System.Drawing.Rectangle(363, 150, 85, 30),
        .DialogResult = DialogResult.Cancel
    }

            frm.Controls.AddRange({lbl, txt, btnOK, btnCancel})
            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            ' Đóng bằng nút X → Cancel
            Dim okClicked As Boolean = False
            AddHandler btnOK.Click, Sub() okClicked = True

            Dim result As DialogResult = frm.ShowDialog()

            If result <> DialogResult.OK OrElse Not okClicked Then
                Return Nothing           ' ← Cancel / X
            End If

            Return txt.Text

        End Function

        '=============================================================
        ' KIỂM TRA PURCHASED
        '=============================================================

        Private Function IsPurchasedPart(
            ByVal doc As Document) As Boolean


            Try

                Dim desc As String =
                    CStr(
                        doc.PropertySets(
                            "Design Tracking Properties").
                            Item("Description").Value)


                Return String.Equals(
                    desc.Trim(),
                    "purchased",
                    StringComparison.OrdinalIgnoreCase)


            Catch

                Return False

            End Try

        End Function


        '=============================================================
        ' REPLACE REFERENCES
        '=============================================================

        Private Sub ReplaceAllReferences(
            ByVal asm As AssemblyDocument,
            ByVal fileMap As Dictionary(Of String, String))


            Try

                '=====================================================
                ' REPLACE REFERENCE CỦA ASSEMBLY
                '=====================================================

                For Each fd As FileDescriptor
                    In asm.File.ReferencedFileDescriptors


                    Try

                        If fileMap.ContainsKey(
                            fd.FullFileName) Then


                            Dim newPath As String =
                                fileMap(fd.FullFileName)


                            If IO.File.Exists(
                                newPath) Then

                                fd.ReplaceReference(
                                    newPath)

                            End If


                        End If


                    Catch

                    End Try

                Next


                asm.Update2(True)

                asm.Save2(True)


                '=====================================================
                ' XỬ LÝ SUB ASSEMBLY
                '=====================================================

                For Each fd As FileDescriptor
                    In asm.File.ReferencedFileDescriptors


                    Try

                        If String.Equals(
                            IO.Path.GetExtension(
                                fd.FullFileName),
                            ".iam",
                            StringComparison.OrdinalIgnoreCase) Then


                            Dim subAsm As AssemblyDocument =
                                CType(
                                    g_inventorApplication.Documents.Open(
                                        fd.FullFileName,
                                        False),
                                    AssemblyDocument)


                            ReplaceAllReferences(
                                subAsm,
                                fileMap)


                            subAsm.Close(
                                True)


                        End If


                    Catch

                    End Try

                Next


            Catch ex As Exception

                MessageBox.Show(
                    "Lỗi Relink:" &
                    vbCrLf & vbCrLf &
                    ex.Message,
                    "Relink")

            End Try

        End Sub


        '=============================================================
        ' UPDATE ALL IPROPERTIES
        '=============================================================

        Private Sub UpdateAlliProperties(
            ByVal asm As AssemblyDocument,
            ByRef processed As HashSet(Of String))


            Try

                '=====================================================
                ' TRÁNH UPDATE TRÙNG
                '=====================================================

                If processed.Contains(
                    asm.FullFileName) Then

                    Exit Sub

                End If


                processed.Add(
                    asm.FullFileName)


                '=====================================================
                ' UPDATE ASSEMBLY
                '=====================================================

                UpdateDocProps(asm)


                '=====================================================
                ' UPDATE COMPONENT
                '=====================================================

                For Each occ As ComponentOccurrence
                    In asm.ComponentDefinition.Occurrences


                    Try

                        If occ.Suppressed Then
                            Continue For
                        End If


                        Dim doc As Document =
                            occ.Definition.Document


                        If doc Is Nothing Then
                            Continue For
                        End If


                        If String.IsNullOrWhiteSpace(
                            doc.FullFileName) Then

                            Continue For

                        End If


                        '---------------------------------------------
                        ' UPDATE PART / ASSEMBLY
                        '---------------------------------------------

                        UpdateDocProps(doc)


                        '---------------------------------------------
                        ' NẾU LÀ SUB ASSEMBLY
                        '---------------------------------------------

                        If doc.DocumentType =
                           DocumentTypeEnum.kAssemblyDocumentObject Then


                            UpdateAlliProperties(
                                CType(
                                    doc,
                                    AssemblyDocument),
                                processed)


                        End If


                    Catch

                    End Try


                Next


            Catch

            End Try

        End Sub


        '=============================================================
        ' UPDATE IPROPERTIES
        '=============================================================

        Private Sub UpdateDocProps(
            ByVal doc As Document)


            Try

                If String.IsNullOrWhiteSpace(
                    doc.FullFileName) Then

                    Exit Sub

                End If


                Dim newName As String =
                    IO.Path.GetFileNameWithoutExtension(
                        doc.FullFileName)


                Dim props As PropertySet =
                    doc.PropertySets(
                        "Design Tracking Properties")


                '=====================================================
                ' PART NUMBER
                '=====================================================

                props.Item(
                    "Part Number").Value =
                    newName


                '=====================================================
                ' DESCRIPTION
                '=====================================================

                props.Item(
                    "Description").Value =
                    newName


                doc.Save2(True)


            Catch

            End Try

        End Sub

    End Module

End Namespace

