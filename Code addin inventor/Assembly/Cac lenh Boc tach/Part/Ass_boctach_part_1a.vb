
Option Explicit On

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip
Imports Inventor
Imports IO = System.IO

Namespace ToolInventor2025.Assembly.Buttons.caclenhboctach.part

    Public Module Ass_boctach_part_1a

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication
            Dim sourceDoc As AssemblyDocument = Nothing
            Dim targetDoc As AssemblyDocument = Nothing
            Dim sourceOpened As Boolean = False

            Try
                Dim partType As Integer = SelectPartType()
                If partType = 0 Then Exit Sub

                Dim selectedFile As String = GetSourceFile(app)
                If String.IsNullOrEmpty(selectedFile) Then Exit Sub

                Dim activeDoc As Document = Nothing

                Try
                    activeDoc = app.ActiveDocument
                Catch
                End Try

                If activeDoc IsNot Nothing AndAlso
                   activeDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject AndAlso
                   String.Equals(activeDoc.FullFileName,
                                 selectedFile,
                                 StringComparison.OrdinalIgnoreCase) Then

                    sourceDoc = CType(activeDoc, AssemblyDocument)

                Else
                    sourceDoc = CType(app.Documents.Open(selectedFile, False),
                                      AssemblyDocument)
                    sourceOpened = True
                End If

                Dim targetOption As Integer = SelectTargetType()

                If targetOption = 0 Then
                    If sourceOpened Then sourceDoc.Close(False)
                    Exit Sub
                End If

                If targetOption = 1 Then

                    targetDoc =
                        CType(app.Documents.Add(
                            DocumentTypeEnum.kAssemblyDocumentObject),
                            AssemblyDocument)

                    Dim baseName As String =
                        IO.Path.GetFileNameWithoutExtension(selectedFile)

                    Select Case partType
                        Case 1
                            targetDoc.DisplayName =
                                baseName & "_SheetMetal_Unfold"

                        Case 2
                            targetDoc.DisplayName =
                                baseName & "_Purchased_Parts"

                        Case 3
                            targetDoc.DisplayName =
                                baseName & "_Library_Parts"

                        Case 4
                            targetDoc.DisplayName =
                                baseName & "_Standard_Parts"
                    End Select

                Else

                    Dim targetFile As String =
                        SelectAssemblyFile("Chọn Assembly đích")

                    If String.IsNullOrEmpty(targetFile) Then

                        If sourceOpened Then
                            sourceDoc.Close(False)
                        End If

                        Exit Sub
                    End If

                    targetDoc =
                        CType(app.Documents.Open(targetFile, True),
                              AssemblyDocument)

                End If

                ' Collect unique top-level occurrences by part name (model-structure, do not sum quantities)
                Dim parts As New List(Of Tuple(Of String, String))

                Try
                    Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                    For Each occ As ComponentOccurrence In sourceDoc.ComponentDefinition.Occurrences
                        Try
                            If occ.Suppressed Then Continue For

                            Dim doc As Document = Nothing
                            If occ.ReferencedDocumentDescriptor IsNot Nothing Then
                                Try
                                    doc = occ.ReferencedDocumentDescriptor.ReferencedDocument
                                Catch
                                    doc = Nothing
                                End Try
                            Else
                                Try
                                    doc = occ.Definition.Document
                                Catch
                                    doc = Nothing
                                End Try
                            End If

                            If doc Is Nothing Then Continue For
                            If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then Continue For

                            Dim part As PartDocument = TryCast(doc, PartDocument)
                            If part Is Nothing Then Continue For

                            '----- Part Number -----
                            Dim partNum As String = ""
                            Try
                                partNum = CStr(doc.PropertySets("Design Tracking Properties").Item("Part Number").Value).Trim()
                            Catch
                                partNum = ""
                            End Try

                            Dim pn As String = If(partNum, "").Trim()
                            Dim name As String = partNum
                            If String.IsNullOrWhiteSpace(name) Then
                                name = IO.Path.GetFileNameWithoutExtension(doc.FullFileName)
                            End If

                            '----- Lọc theo CommonKeywords (giống CollectParts) -----
                            Dim include As Boolean = False

                            Select Case partType
                                Case 1   ' Sheet Metal
                                    Dim sm As SheetMetalComponentDefinition =
                        TryCast(part.ComponentDefinition, SheetMetalComponentDefinition)
                                    include = (sm IsNot Nothing AndAlso sm.Features.Count > 0)

                                Case 2   ' Purchased
                                    include = (occ.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure)

                                Case 3   ' Standard Library  (Bearing + Fastener + Standard)
                                    Dim bearing As Boolean = False
                                    Dim fastener As Boolean = False
                                    Dim standard As Boolean = False

                                    Try
                                        bearing = ToolInventor2025.Assembly.CommonKeywords.IsBearing(pn)
                                    Catch
                                        bearing = False
                                    End Try

                                    Try
                                        fastener = ToolInventor2025.Assembly.CommonKeywords.IsFastener(pn)
                                    Catch
                                        fastener = False
                                    End Try

                                    standard = IsLocalStandardPart(pn)   ' StartsWith theo CommonKeywords.StandardKeywords

                                    include = bearing OrElse fastener OrElse standard

                                Case 4   ' Purchased + Library
                                    Dim purchased As Boolean =
                        (occ.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure)

                                    Dim bearing As Boolean = False
                                    Dim fastener As Boolean = False
                                    Dim standard As Boolean = False

                                    Try
                                        bearing = ToolInventor2025.Assembly.CommonKeywords.IsBearing(pn)
                                    Catch
                                        bearing = False
                                    End Try

                                    Try
                                        fastener = ToolInventor2025.Assembly.CommonKeywords.IsFastener(pn)
                                    Catch
                                        fastener = False
                                    End Try

                                    standard = IsLocalStandardPart(pn)

                                    Dim library As Boolean = bearing OrElse fastener OrElse standard
                                    include = purchased OrElse library

                                Case Else
                                    include = False
                            End Select

                            If Not include Then Continue For
                            If String.IsNullOrWhiteSpace(part.FullFileName) Then Continue For

                            If Not seen.Contains(name) Then
                                seen.Add(name)
                                parts.Add(Tuple.Create(doc.FullFileName, name))
                            End If

                        Catch
                        End Try
                    Next

                Catch
                    ' fallback recursive
                    Dim visitedAssemblies As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    CollectParts(sourceDoc.ComponentDefinition, parts, partType, visitedAssemblies)
                End Try

                If parts.Count = 0 Then

                    MessageBox.Show(
                        "Không tìm thấy chi tiết phù hợp.",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

                    If sourceOpened Then
                        sourceDoc.Close(False)
                    End If

                    targetDoc.Close(False)
                    Exit Sub
                End If

                parts.Sort(
                    Function(a, b)
                        Return StringComparer.OrdinalIgnoreCase.Compare(
                            a.Item2,
                            b.Item2)
                    End Function)

                Dim matrix As Matrix =
                    app.TransientGeometry.CreateMatrix()

                Dim index As Integer = 1

                For Each p As Tuple(Of String, String) In parts

                    Try

                        Dim occ As ComponentOccurrence =
                            targetDoc.ComponentDefinition.Occurrences.Add(
                                p.Item1,
                                matrix)

                        Try
                            occ.Grounded = True
                        Catch
                        End Try

                        Try
                            occ.Name =
                                index.ToString("000") &
                                " - " &
                                p.Item2
                        Catch
                        End Try

                        index += 1

                    Catch
                    End Try

                Next

                targetDoc.Update2(True)

                Dim typeName As String =
                    GetTypeName(partType)

                If targetOption = 1 Then

                    Dim defaultName As String = ""

                    Select Case partType

                        Case 1
                            defaultName = "SheetMetal_Unfold.iam"

                        Case 2
                            defaultName = "Purchased_Parts.iam"

                        Case 3
                            defaultName = "Library_Parts.iam"

                        Case 4
                            defaultName = "Standard_Parts.iam"

                    End Select

                    Dim saveFile As String =
                        SaveAssemblyFile(defaultName)

                    If Not String.IsNullOrEmpty(saveFile) Then

                        targetDoc.SaveAs(
                            saveFile,
                            False)

                        MessageBox.Show(
                            "Đã tạo Assembly mới." &
                            vbCrLf &
                            "Loại: " & typeName &
                            vbCrLf &
                            "Số lượng: " &
                            parts.Count.ToString() &
                            vbCrLf &
                            "File: " & saveFile,
                            "Hoàn tất",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

                    End If

                Else

                    targetDoc.Save()

                    MessageBox.Show(
                        "Đã thêm " &
                        parts.Count.ToString() &
                        " chi tiết " &
                        typeName &
                        " vào Assembly đích.",
                        "Hoàn tất",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

                End If

                If sourceOpened Then
                    sourceDoc.Close(False)
                End If

                targetDoc.Activate()

            Catch ex As Exception

                MessageBox.Show(
                    "Lỗi:" &
                    vbCrLf &
                    ex.Message,
                    "Select Parts",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

                If sourceOpened AndAlso
                   sourceDoc IsNot Nothing Then

                    Try
                        sourceDoc.Close(False)
                    Catch
                    End Try

                End If

            End Try

        End Sub

        Private Function SelectPartType() As Integer

            Dim result As Integer = 0

            Using f As New Form()

                f.Text = "CHỌN DẠNG CHI TIẾT"
                f.StartPosition =
                    FormStartPosition.CenterScreen

                f.Size =
                    New System.Drawing.Size(620, 375)

                f.FormBorderStyle =
                    FormBorderStyle.FixedDialog

                f.MaximizeBox = False
                f.MinimizeBox = False

                Dim title As New Label()

                title.Text =
                    "CHỌN DẠNG CHI TIẾT"

                title.Font =
                    New System.Drawing.Font(
                        "Tahoma",
                        18,
                        System.Drawing.FontStyle.Bold)

                title.AutoSize = True

                title.Location =
                    New System.Drawing.Point(135, 20)

                f.Controls.Add(title)

                Dim b1 As Button =
                    MakePartButton(
                        "SHEET METAL",
                        "Tấm sheet metal",
                        25,
                        70)

                Dim b2 As Button =
                    MakePartButton(
                        "PURCHASED",
                        "Chi tiết mua",
                        315,
                        70)

                Dim b3 As Button =
                    MakePartButton(
                        "STANDARD LIBRARY",
                        "ISO / DIN / SKF",
                        25,
                        175)

                Dim b4 As Button =
                    MakePartButton(
                        "PURCHASED + LIBRARY",
                        "Cả hai loại",
                        315,
                        175)

                AddHandler b1.Click,
                    Sub(sender As Object, e As EventArgs)

                        result = 1
                        f.Close()

                    End Sub

                AddHandler b2.Click,
                    Sub(sender As Object, e As EventArgs)

                        result = 2
                        f.Close()

                    End Sub

                AddHandler b3.Click,
                    Sub(sender As Object, e As EventArgs)

                        result = 3
                        f.Close()

                    End Sub

                AddHandler b4.Click,
                    Sub(sender As Object, e As EventArgs)

                        result = 4
                        f.Close()

                    End Sub

                f.Controls.Add(b1)
                f.Controls.Add(b2)
                f.Controls.Add(b3)
                f.Controls.Add(b4)

                Dim cancel As New Button()

                cancel.Text = "HỦY"

                cancel.Size =
                    New System.Drawing.Size(100, 30)

                cancel.Location =
                    New System.Drawing.Point(260, 275)

                cancel.DialogResult =
                    DialogResult.Cancel

                f.Controls.Add(cancel)

                f.CancelButton = cancel

                f.ShowDialog()

            End Using

            Return result

        End Function

        Private Function MakePartButton(
            text As String,
            subText As String,
            x As Integer,
            y As Integer) As Button

            Dim b As New Button()

            b.Text =
                text &
                vbCrLf &
                subText

            b.Font =
                New System.Drawing.Font(
                    "Tahoma",
                    10,
                    System.Drawing.FontStyle.Bold)

            b.Size =
                New System.Drawing.Size(270, 85)

            b.Location =
                New System.Drawing.Point(x, y)

            b.FlatStyle =
                FlatStyle.Flat

            Return b

        End Function

        Private Function SelectTargetType() As Integer

            Dim result As Integer = 0

            Using f As New Form()

                f.Text = "CHỌN ASSEMBLY ĐÍCH"

                f.StartPosition =
                    FormStartPosition.CenterScreen

                f.Size =
                    New System.Drawing.Size(620, 280)

                f.FormBorderStyle =
                    FormBorderStyle.FixedDialog

                f.MaximizeBox = False
                f.MinimizeBox = False

                Dim title As New Label()

                title.Text =
                    "Chọn cách thêm chi tiết"

                title.Font =
                    New System.Drawing.Font(
                        "Tahoma",
                        13,
                        System.Drawing.FontStyle.Bold)

                title.AutoSize = True

                title.Location =
                    New System.Drawing.Point(205, 20)

                f.Controls.Add(title)

                Dim newBtn As New Button()

                newBtn.Text =
                    "TẠO ASSEMBLY MỚI" &
                    vbCrLf &
                    "Tạo file .iam mới"

                newBtn.Font =
                    New System.Drawing.Font(
                        "Tahoma",
                        10,
                        System.Drawing.FontStyle.Bold)

                newBtn.Size =
                    New System.Drawing.Size(270, 90)

                newBtn.Location =
                    New System.Drawing.Point(25, 70)

                Dim oldBtn As New Button()

                oldBtn.Text =
                    "ASSEMBLY CÓ SẴN" &
                    vbCrLf &
                    "Chọn file .iam để thêm"

                oldBtn.Font =
                    New System.Drawing.Font(
                        "Tahoma",
                        10,
                        System.Drawing.FontStyle.Bold)

                oldBtn.Size =
                    New System.Drawing.Size(270, 90)

                oldBtn.Location =
                    New System.Drawing.Point(315, 70)

                AddHandler newBtn.Click,
                    Sub(sender As Object, e As EventArgs)

                        result = 1
                        f.Close()

                    End Sub

                AddHandler oldBtn.Click,
                    Sub(sender As Object, e As EventArgs)

                        result = 2
                        f.Close()

                    End Sub

                f.Controls.Add(newBtn)
                f.Controls.Add(oldBtn)

                Dim cancel As New Button()

                cancel.Text = "HỦY"

                cancel.Size =
                    New System.Drawing.Size(100, 30)

                cancel.Location =
                    New System.Drawing.Point(255, 180)

                cancel.DialogResult =
                    DialogResult.Cancel

                f.Controls.Add(cancel)

                f.CancelButton = cancel

                f.ShowDialog()

            End Using

            Return result

        End Function

        Private Function GetSourceFile(
            app As Inventor.Application) As String

            Dim doc As Document = Nothing

            Try
                doc = app.ActiveDocument
            Catch
            End Try

            If doc IsNot Nothing AndAlso
               doc.DocumentType =
               DocumentTypeEnum.kAssemblyDocumentObject Then

                Dim r As DialogResult =
                    MessageBox.Show(
                        "Sử dụng Assembly đang mở làm nguồn?",
                        "Assembly nguồn",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question)

                If r = DialogResult.Yes Then
                    Return doc.FullFileName
                End If

                If r = DialogResult.Cancel Then
                    Return ""
                End If

            End If

            Return SelectAssemblyFile(
                "Chọn Assembly nguồn")

        End Function

        Private Function SelectAssemblyFile(
            title As String) As String

            Using dlg As New OpenFileDialog()

                dlg.Title = title
                dlg.Filter =
                    "Assembly Files (*.iam)|*.iam"

                dlg.Multiselect = False

                If dlg.ShowDialog() =
                   DialogResult.OK Then

                    Return dlg.FileName

                End If

            End Using

            Return ""

        End Function

        Private Function SaveAssemblyFile(
            defaultName As String) As String

            Using dlg As New SaveFileDialog()

                dlg.Title =
                    "Lưu Assembly mới"

                dlg.Filter =
                    "Assembly Files (*.iam)|*.iam"

                dlg.DefaultExt = "iam"
                dlg.AddExtension = True
                dlg.FileName = defaultName

                If dlg.ShowDialog() =
                   DialogResult.OK Then

                    Return dlg.FileName

                End If

            End Using

            Return ""

        End Function

        Private Function GetTypeName(
            partType As Integer) As String

            Select Case partType

                Case 1
                    Return "Sheet Metal"

                Case 2
                    Return "Purchased"

                Case 3
                    Return "Standard Library"

                Case 4
                    Return "Purchased + Library"

            End Select

            Return "Part"

        End Function

        Private Sub CollectParts(
    asmDef As AssemblyComponentDefinition,
    ByRef parts As List(Of Tuple(Of String, String)),
    partType As Integer,
    ByRef visitedAssemblies As HashSet(Of String))

            For Each occ As ComponentOccurrence In asmDef.Occurrences

                Try

                    '==========================================================
                    ' BỎ QUA SUPPRESS
                    '==========================================================
                    If occ.Suppressed Then
                        Continue For
                    End If


                    '==========================================================
                    ' LẤY DOCUMENT
                    '==========================================================
                    Dim doc As Document = Nothing

                    Try
                        If occ.ReferencedDocumentDescriptor IsNot Nothing Then
                            doc = occ.ReferencedDocumentDescriptor.ReferencedDocument
                        End If
                    Catch
                        doc = Nothing
                    End Try

                    If doc Is Nothing Then
                        Try
                            doc = occ.Definition.Document
                        Catch
                            doc = Nothing
                        End Try
                    End If

                    If doc Is Nothing Then
                        Continue For
                    End If


                    '==========================================================
                    '==========================================================
                    ' PART
                    '==========================================================
                    If doc.DocumentType =
                       DocumentTypeEnum.kPartDocumentObject Then

                        Dim part As PartDocument =
                            TryCast(doc, PartDocument)

                        If part Is Nothing Then
                            Continue For
                        End If


                        '======================================================
                        ' LẤY PART NUMBER
                        '======================================================
                        Dim partNum As String = ""

                        Try

                            partNum =
                                CStr(
                                    part.PropertySets(
                                        "Design Tracking Properties").
                                    Item("Part Number").Value
                                ).Trim()

                        Catch
                            partNum = ""
                        End Try


                        '======================================================
                        ' TÊN DÙNG ĐỂ HIỂN THỊ / SORT
                        '======================================================
                        Dim name As String = partNum

                        If String.IsNullOrWhiteSpace(name) Then

                            Try

                                name =
                                    IO.Path.GetFileNameWithoutExtension(
                                        part.FullFileName)

                            Catch

                                name = ""

                            End Try

                        End If


                        '======================================================
                        ' CHUẨN HÓA PART NUMBER
                        '======================================================
                        Dim pn As String =
                            If(partNum, "").Trim()


                        '======================================================
                        ' BIẾN LỌC
                        '======================================================
                        Dim selected As Boolean = False


                        Select Case partType


                    '==================================================
                    ' 1 - SHEET METAL
                    '==================================================
                            Case 1

                                Dim sm As SheetMetalComponentDefinition =
                                    TryCast(
                                        part.ComponentDefinition,
                                        SheetMetalComponentDefinition)

                                selected =
                                    sm IsNot Nothing AndAlso
                                    sm.Features.Count > 0


                    '==================================================
                    ' 2 - PURCHASED
                    '==================================================
                            Case 2

                                selected =
                                    occ.BOMStructure =
                                    BOMStructureEnum.kPurchasedBOMStructure


                    '==================================================
                    ' 3 - STANDARD LIBRARY
                    '==================================================
                            Case 3

                                Dim bearing As Boolean = False
                                Dim fastener As Boolean = False
                                Dim standard As Boolean = False


                                '----------------------------------------------
                                ' Bearing
                                ' Giữ nguyên CommonKeywords
                                '----------------------------------------------
                                Try
                                    bearing =
                                        ToolInventor2025.Assembly.CommonKeywords.IsBearing(pn)
                                Catch
                                    bearing = False
                                End Try


                                '----------------------------------------------
                                ' Fastener
                                ' Giữ nguyên CommonKeywords
                                '----------------------------------------------
                                Try
                                    fastener =
                                        ToolInventor2025.Assembly.CommonKeywords.IsFastener(pn)
                                Catch
                                    fastener = False
                                End Try


                                '----------------------------------------------
                                ' Standard
                                '
                                ' KHÔNG dùng CommonKeywords.IsStandardKeyword
                                ' vì file Common đang được dùng chung.
                                '
                                ' Kiểm tra riêng tại code này bằng StartsWith.
                                '----------------------------------------------
                                standard =
                                    IsLocalStandardPart(pn)


                                selected =
                                    bearing OrElse
                                    fastener OrElse
                                    standard


                    '==================================================
                    ' 4 - PURCHASED + LIBRARY
                    '==================================================
                            Case 4

                                Dim purchased As Boolean =
                                    occ.BOMStructure =
                                    BOMStructureEnum.kPurchasedBOMStructure


                                Dim bearing As Boolean = False
                                Dim fastener As Boolean = False
                                Dim standard As Boolean = False


                                '----------------------------------------------
                                ' Bearing
                                '----------------------------------------------
                                Try
                                    bearing =
                                        ToolInventor2025.Assembly.CommonKeywords.IsBearing(pn)
                                Catch
                                    bearing = False
                                End Try


                                '----------------------------------------------
                                ' Fastener
                                '----------------------------------------------
                                Try
                                    fastener =
                                        ToolInventor2025.Assembly.CommonKeywords.IsFastener(pn)
                                Catch
                                    fastener = False
                                End Try


                                '----------------------------------------------
                                ' Standard - kiểm tra local
                                '----------------------------------------------
                                standard =
                                    IsLocalStandardPart(pn)


                                Dim library As Boolean =
                                    bearing OrElse
                                    fastener OrElse
                                    standard


                                selected =
                                    purchased OrElse
                                    library


                            Case Else

                                selected = False

                        End Select


                        '======================================================
                        ' KHÔNG PHÙ HỢP
                        '======================================================
                        If Not selected Then
                            Continue For
                        End If


                        '======================================================
                        ' CHỈ NHẬN FILE PART
                        '======================================================
                        If String.IsNullOrWhiteSpace(part.FullFileName) Then
                            Continue For
                        End If


                        '======================================================
                        ' KHÔNG LẤY TRÙNG PART
                        '======================================================
                        Dim duplicate As Boolean = False

                        For Each oldPart As Tuple(Of String, String) In parts

                            If String.Equals(
                                oldPart.Item2,
                                name,
                                StringComparison.OrdinalIgnoreCase) Then

                                duplicate = True
                                Exit For

                            End If

                        Next


                        If Not duplicate Then

                            parts.Add(
                                Tuple.Create(
                                    part.FullFileName,
                                    name))

                        End If


                        '==========================================================
                        ' ASSEMBLY
                        '
                        ' KHÔNG THÊM ASSEMBLY.
                        '
                        ' Chỉ đi xuống Assembly con để tìm PART.
                        '==========================================================
                    ElseIf doc.DocumentType =
                           DocumentTypeEnum.kAssemblyDocumentObject Then

                        Dim subAsm As AssemblyDocument =
                            TryCast(doc, AssemblyDocument)

                        If subAsm Is Nothing Then
                            Continue For
                        End If


                        '======================================================
                        ' LẤY ĐƯỜNG DẪN ASSEMBLY CON
                        '======================================================
                        Dim subAsmFullName As String = ""

                        Try
                            subAsmFullName = subAsm.FullFileName
                        Catch
                            subAsmFullName = ""
                        End Try


                        If String.IsNullOrWhiteSpace(subAsmFullName) Then
                            Continue For
                        End If


                        '======================================================
                        ' CHỐNG VÒNG LẶP
                        '======================================================
                        If visitedAssemblies.Contains(subAsmFullName) Then
                            Continue For
                        End If

                        visitedAssemblies.Add(subAsmFullName)


                        '======================================================
                        ' ĐI XUỐNG ASSEMBLY CON
                        '======================================================
                        Try

                            CollectParts(
                                subAsm.ComponentDefinition,
                                parts,
                                partType,
                                visitedAssemblies)

                        Catch
                        End Try

                    End If

                Catch
                    ' Bỏ qua component lỗi
                End Try

            Next

        End Sub


        '======================================================================
        ' STANDARD KEYWORD RIÊNG CHO CODE NÀY
        '
        ' KHÔNG SỬA CommonKeywords
        '
        ' Chỉ nhận keyword ở KÝ TỰ ĐẦU TIÊN.
        '======================================================================
        Private Function IsLocalStandardPart(
            ByVal partNum As String) As Boolean

            If String.IsNullOrWhiteSpace(partNum) Then
                Return False
            End If


            Dim pn As String =
                partNum.Trim()


            '==========================================================
            ' DANH SÁCH STANDARD
            '
            ' Giữ theo CommonKeywords hiện tại
            '==========================================================
            Dim standardKeywords As String() = ToolInventor2025.Assembly.CommonKeywords.StandardKeywords


            '==========================================================
            ' CHỈ KIỂM TRA TỪ KÝ TỰ ĐẦU TIÊN
            '==========================================================
            For Each kw As String In standardKeywords

                If pn.StartsWith(
                    kw,
                    StringComparison.OrdinalIgnoreCase) Then

                    Return True

                End If

            Next


            Return False

        End Function


    End Module

End Namespace

