

Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode

    Public Module Ass_Bom_1

        Private ReadOnly BearingKeywords As String() = {
            "vòng bi", "vong bi", "bearing", "motor",
            "gối bi", "goi bi", "gối đỡ", "goi do",
            "pillow", "plummer", "ucp", "ucf", "ucfl", "khóa trục", "khoa truc"
        }

        Private ReadOnly FastenerKeywords As String() = {
            "bulong", "bu lông", "bu long", "ốc", "oc", "đai ốc", "dai oc", "vít", "vit", "ecu", "êcu", "then", "then chốt", "long đen", "long den",
            "long đen", "long den", "washer", "iso", "din", "jis", "m3", "m4", "m5", "m6", "m8", "lock collar", "locknut", "lock nut",
            "m10", "m12", "m16", "m20", "m24", "m30", "m36", "m42", "m48", "ss 2", "iso 4", "din 125", "din 127", "din 933", "din 934", "din 6912"
        }

        Private Function GetInventorApplication() As Inventor.Application
            Try
                Return CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch ex As Exception
                MessageBox.Show("Không lấy được Inventor đang chạy." & vbCrLf & vbCrLf & ex.Message,
                                "BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End Try
        End Function

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim invApp As Inventor.Application = Nothing
            Try
                invApp = GetInventorApplication()
                If invApp Is Nothing Then Exit Sub
                Main(invApp)
            Catch ex As Exception
                MessageBox.Show("Lỗi BOM:" & vbCrLf & vbCrLf & ex.Message,
                                "BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub Main(ByVal invApp As Inventor.Application)
            Dim oAsm As AssemblyDocument = Nothing
            Try
                oAsm = TryCast(invApp.ActiveDocument, AssemblyDocument)
                If oAsm Is Nothing Then
                    MessageBox.Show("Vui lòng mở Assembly (.iam) trước khi chạy.",
                                    "BOM", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                Dim oBOM As BOM = oAsm.ComponentDefinition.BOM
                Try : oBOM.StructuredViewEnabled = True : Catch : End Try
                Try : oBOM.StructuredViewFirstLevelOnly = False : Catch : End Try

                Dim oBOMView As BOMView = Nothing
                Try
                    oBOMView = oBOM.BOMViews.Item("Structured")
                Catch
                    MessageBox.Show("Không tìm thấy Structured BOM.", "BOM",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End Try
                If oBOMView Is Nothing Then Exit Sub

                Dim levelIdx As Integer = PickFromList(
                    "Chọn phạm vi",
                    New String() {
                        "1 - Chỉ Top-level",
                        "2 - All-level (mỗi cấp đánh số riêng + kiểm tra PN)"
                    }, 0)

                If levelIdx < 0 Then Exit Sub
                Dim isAllLevel As Boolean = (levelIdx = 1)



                '===== CHỌN SẮP XẾP RIÊNG CHO CỤM LẮP VÀ PART =====
                Dim sortAsmIdx As Integer = PickFromList(
    "Sắp xếp CỤM LẮP (Assembly)",
    New String() {
        "1 - Khối lượng lớn → bé",
        "2 - Khối lượng bé → lớn",
        "3 - Tên ngắn → dài",
        "4 - Tên dài → ngắn",
        "5 - Chữ cái A → Z",
        "6 - Chữ cái Z → A"
    }, 0)
                If sortAsmIdx < 0 Then Exit Sub

                Dim sortPartIdx As Integer = PickFromList(
    "Sắp xếp PART",
    New String() {
        "1 - Khối lượng lớn → bé",
        "2 - Khối lượng bé → lớn",
        "3 - Tên ngắn → dài",
        "4 - Tên dài → ngắn",
        "5 - Chữ cái A → Z",
        "6 - Chữ cái Z → A"
    }, 0)
                If sortPartIdx < 0 Then Exit Sub

                Dim sortModeAsm As Integer = sortAsmIdx
                Dim sortModePart As Integer = sortPartIdx






                Dim baseText As String = InputBox(
                    "Nhập chữ dùng cho Part Number / Stock Number." & vbCrLf & vbCrLf &
                    "Để trống = chỉ Sort + đánh STT (không sửa PN/SN).",
                    "Part Number / Stock Number", "")

                If baseText Is Nothing Then baseText = ""
                baseText = baseText.Trim()

                Dim mode As Integer = 1
                Dim applyPN As Boolean = False
                Dim applySN As Boolean = False

                If baseText <> "" Then
                    Dim modeIdx As Integer = PickFromList(
                        "Chọn cách ghi",
                        New String() {
                            "2 - Thay toàn bộ (xóa cũ → chữ + STT)",
                            "3 - Thêm chữ phía SAU",
                            "4 - Thêm chữ phía TRƯỚC"
                        }, 0)

                    If modeIdx < 0 Then Exit Sub
                    mode = modeIdx + 2

                    Dim targetIdx As Integer = PickFromList(
                        "Áp dụng cho",
                        New String() {
                            "Chỉ Part Number",
                            "Chỉ Stock Number",
                            "Cả Part Number và Stock Number"
                        }, 2)

                    If targetIdx < 0 Then Exit Sub

                    Select Case targetIdx
                        Case 0 : applyPN = True : applySN = False
                        Case 1 : applyPN = False : applySN = True
                        Case 2 : applyPN = True : applySN = True
                    End Select
                End If

                Dim confirm As String =
                    "CHUẨN BỊ CHẠY" & vbCrLf &
                    "------------------" & vbCrLf &
                    If(isAllLevel, "All-level (mỗi cấp riêng + kiểm tra PN)", "Chỉ Top-level") & vbCrLf

                If baseText = "" Then
                    confirm &= "PN / SN: KHÔNG SỬA"
                Else
                    confirm &= "Chữ: " & baseText & vbCrLf
                    Select Case mode
                        Case 2 : confirm &= "Cách: Thay toàn bộ + STT" & vbCrLf
                        Case 3 : confirm &= "Cách: Thêm phía sau" & vbCrLf
                        Case 4 : confirm &= "Cách: Thêm phía trước" & vbCrLf
                    End Select
                    If applyPN AndAlso applySN Then
                        confirm &= "Áp dụng: PN + SN"
                    ElseIf applyPN Then
                        confirm &= "Áp dụng: Chỉ Part Number"
                    Else
                        confirm &= "Áp dụng: Chỉ Stock Number"
                    End If
                End If

                confirm &= vbCrLf & vbCrLf & "Tiếp tục?"

                If MessageBox.Show(confirm, "Xác nhận",
                                   MessageBoxButtons.OKCancel,
                                   MessageBoxIcon.Question) <> DialogResult.OK Then
                    Exit Sub
                End If

                Dim changedPN As Integer = 0
                Dim changedSN As Integer = 0
                Dim totalRows As Integer = 0
                Dim listDocs As New List(Of Document)

                If isAllLevel Then
                    '  ProcessLevelWithPNCheck(oBOMView.BOMRows, baseText, mode, applyPN, applySN,
                    'changedPN, changedSN, totalRows, listDocs, oAsm)


                    ' ProcessLevelWithPNCheck(oBOMView.BOMRows, baseText, mode, applyPN, applySN,
                    'changedPN, changedSN, totalRows, listDocs, oAsm, sortMode)
                    ProcessLevelWithPNCheck(oBOMView.BOMRows, baseText, mode, applyPN, applySN,
                        changedPN, changedSN, totalRows, listDocs, oAsm,
                        sortModeAsm, sortModePart)
                Else
                    ' Top-level only
                    ' Dim sortedRows As List(Of BOMRow) = SortRows(oBOMView.BOMRows)'''''''''''''''' sua lan 1
                    ' Dim sortedRows As List(Of BOMRow) = SortRows(oBOMView.BOMRows, sortMode) '' sưa lan 2
                    Dim sortedRows As List(Of BOMRow) = SortRows(oBOMView.BOMRows, sortModeAsm, sortModePart)
                    Dim stt As Integer = 1
                    Dim pnToStt As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                    For Each row As BOMRow In sortedRows
                        System.Windows.Forms.Application.DoEvents()
                        If row Is Nothing Then Continue For

                        Dim refDoc As Document = Nothing
                        Try
                            If row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then Continue For
                            refDoc = row.ComponentDefinitions.Item(1).Document
                        Catch
                            Continue For
                        End Try
                        If refDoc Is Nothing Then Continue For

                        Try
                            If String.Equals(refDoc.FullFileName, oAsm.FullFileName, StringComparison.OrdinalIgnoreCase) Then
                                Continue For
                            End If
                        Catch
                        End Try

                        Dim pn As String = GetProperty(refDoc, "Part Number")
                        If String.IsNullOrEmpty(pn) Then pn = refDoc.DisplayName

                        Dim thisStt As Integer
                        If pnToStt.ContainsKey(pn) Then
                            thisStt = pnToStt(pn)
                        Else
                            thisStt = stt
                            pnToStt(pn) = stt
                            stt += 1
                        End If

                        Try : row.ItemNumber = thisStt.ToString() : Catch : End Try

                        If baseText <> "" AndAlso mode >= 2 Then
                            Dim curPN As String = GetProperty(refDoc, "Part Number")
                            Dim curSN As String = GetProperty(refDoc, "Stock Number")
                            Dim newPN As String = BuildValue(curPN, baseText, thisStt, mode)
                            Dim newSN As String = BuildValue(curSN, baseText, thisStt, mode)

                            If applyPN AndAlso newPN <> "" Then
                                If SetDesignProperty(refDoc, "Part Number", newPN) Then
                                    changedPN += 1
                                    If Not listDocs.Contains(refDoc) Then listDocs.Add(refDoc)
                                End If
                            End If
                            If applySN AndAlso newSN <> "" Then
                                If SetDesignProperty(refDoc, "Stock Number", newSN) Then
                                    changedSN += 1
                                    If Not listDocs.Contains(refDoc) Then listDocs.Add(refDoc)
                                End If
                            End If
                        End If

                        totalRows += 1
                    Next
                End If

                For Each d As Document In listDocs
                    Try
                        If d.IsModifiable Then d.Update()
                    Catch
                    End Try
                Next

                Try : oBOM.Update() : Catch : End Try
                Try : oAsm.Update2(True) : Catch : End Try

                Dim msg As String =
                    "HOÀN TẤT" & vbCrLf &
                    "====================" & vbCrLf &
                    "Tổng dòng xử lý : " & totalRows.ToString() & vbCrLf & vbCrLf

                If baseText = "" Then
                    msg &= "Part Number / Stock Number: KHÔNG SỬA"
                Else
                    msg &= "Part Number đã ghi : " & changedPN.ToString() & vbCrLf &
                           "Stock Number đã ghi: " & changedSN.ToString()
                End If

                MessageBox.Show(msg, "BOM", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & vbCrLf & ex.Message,
                                "BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ' All-level: mỗi cấp đánh số riêng + kiểm tra PN trùng trong cấp
        Private Sub ProcessLevelWithPNCheck(rows As BOMRowsEnumerator,
                                    baseText As String, mode As Integer,
                                    applyPN As Boolean, applySN As Boolean,
                                    ByRef changedPN As Integer, ByRef changedSN As Integer,
                                    ByRef totalRows As Integer,
                                    listDocs As List(Of Document),
                                    oAsm As AssemblyDocument,
                                    Optional sortModeAsm As Integer = 0,
                                    Optional sortModePart As Integer = 0)

            If rows Is Nothing Then Exit Sub

            Dim sortedRows As List(Of BOMRow) = SortRows(rows, sortModeAsm, sortModePart)
            If sortedRows Is Nothing OrElse sortedRows.Count = 0 Then Exit Sub

            Dim stt As Integer = 1
            Dim pnToStt As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

            For Each row As BOMRow In sortedRows
                System.Windows.Forms.Application.DoEvents()
                If row Is Nothing Then Continue For

                Dim refDoc As Document = Nothing
                Try
                    If row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then Continue For
                    refDoc = row.ComponentDefinitions.Item(1).Document
                Catch
                    Continue For
                End Try
                If refDoc Is Nothing Then Continue For

                Try
                    If String.Equals(refDoc.FullFileName, oAsm.FullFileName, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                Catch
                End Try

                Dim pn As String = GetProperty(refDoc, "Part Number")
                If String.IsNullOrEmpty(pn) Then pn = refDoc.DisplayName

                Dim thisStt As Integer
                If pnToStt.ContainsKey(pn) Then
                    thisStt = pnToStt(pn)          ' trùng PN → cùng STT
                Else
                    thisStt = stt
                    pnToStt(pn) = stt
                    stt += 1
                End If

                Try : row.ItemNumber = thisStt.ToString() : Catch : End Try

                If baseText <> "" AndAlso mode >= 2 Then
                    Dim curPN As String = GetProperty(refDoc, "Part Number")
                    Dim curSN As String = GetProperty(refDoc, "Stock Number")
                    Dim newPN As String = BuildValue(curPN, baseText, thisStt, mode)
                    Dim newSN As String = BuildValue(curSN, baseText, thisStt, mode)

                    If applyPN AndAlso newPN <> "" Then
                        If SetDesignProperty(refDoc, "Part Number", newPN) Then
                            changedPN += 1
                            If Not listDocs.Contains(refDoc) Then listDocs.Add(refDoc)
                        End If
                    End If
                    If applySN AndAlso newSN <> "" Then
                        If SetDesignProperty(refDoc, "Stock Number", newSN) Then
                            changedSN += 1
                            If Not listDocs.Contains(refDoc) Then listDocs.Add(refDoc)
                        End If
                    End If
                End If

                totalRows += 1

                ' Nhảy vào sub-assembly Structure (giống 1c) – chỉ để xử lý PN/STT bên trong
                If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Try
                        Dim subAsm As AssemblyDocument = CType(refDoc, AssemblyDocument)
                        Dim subBOM As BOM = subAsm.ComponentDefinition.BOM
                        Try : subBOM.StructuredViewEnabled = True : Catch : End Try
                        Try : subBOM.StructuredViewFirstLevelOnly = False : Catch : End Try
                        Try : subBOM.Update() : Catch : End Try

                        Dim subView As BOMView = Nothing
                        Try
                            subView = subBOM.BOMViews.Item("Structured")
                        Catch
                        End Try

                        If subView IsNot Nothing Then
                            '   ProcessLevelWithPNCheck(subView.BOMRows, baseText, mode, applyPN, applySN,
                            'changedPN, changedSN, totalRows, listDocs, oAsm)

                            '  ProcessLevelWithPNCheck(subView.BOMRows, baseText, mode, applyPN, applySN,
                            ' changedPN, changedSN, totalRows, listDocs, oAsm, sortMode)
                            ProcessLevelWithPNCheck(subView.BOMRows, baseText, mode, applyPN, applySN,
                            changedPN, changedSN, totalRows, listDocs, oAsm,
                            sortModeAsm, sortModePart)
                        End If
                    Catch
                    End Try
                End If
            Next
        End Sub

        Private Function BuildValue(current As String, baseText As String, stt As Integer, mode As Integer) As String
            If current Is Nothing Then current = ""
            current = current.Trim()
            Select Case mode
                Case 2 : Return baseText & stt.ToString()
                Case 3
                    If current = "" Then Return baseText
                    Return current & " " & baseText
                Case 4
                    If current = "" Then Return baseText
                    Return baseText & " " & current
            End Select
            Return ""
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

        Private Function GetSearchText(row As BOMRow) As String
            Try
                If row Is Nothing OrElse row.ComponentDefinitions Is Nothing OrElse
                   row.ComponentDefinitions.Count = 0 Then Return ""
                Dim doc As Document = row.ComponentDefinitions.Item(1).Document
                Dim pn As String = GetProperty(doc, "Part Number")
                If pn <> "" Then Return pn.Trim().ToLowerInvariant()
                Dim sn As String = GetProperty(doc, "Stock Number")
                If sn <> "" Then Return sn.Trim().ToLowerInvariant()
                Dim desc As String = GetProperty(doc, "Description")
                Return desc.Trim().ToLowerInvariant()
            Catch
                Return ""
            End Try
        End Function

        Private Function IsBearing(text As String) As Boolean
            If String.IsNullOrEmpty(text) Then Return False
            For Each kw As String In BearingKeywords
                If text.StartsWith(kw.ToLowerInvariant()) Then Return True
            Next
            Return False
        End Function

        Private Function IsFastener(text As String) As Boolean
            If String.IsNullOrEmpty(text) Then Return False
            For Each kw As String In FastenerKeywords
                If text.StartsWith(kw.ToLowerInvariant()) Then Return True
            Next
            Return False
        End Function
        Private Function SortRows(ByVal bomRows As BOMRowsEnumerator,
                          Optional ByVal sortModeAsm As Integer = 0,
                          Optional ByVal sortModePart As Integer = 0) As List(Of BOMRow)

            Dim normalAsm As New List(Of Tuple(Of BOMRow, Double, String))
            Dim purchasedAsm As New List(Of Tuple(Of BOMRow, Integer, String))
            Dim normalPart As New List(Of Tuple(Of BOMRow, Double, String))
            Dim purchasedPart As New List(Of Tuple(Of BOMRow, Integer, String))
            Dim phantomAsm As New List(Of Tuple(Of BOMRow, Double, String))
            Dim phantomPart As New List(Of Tuple(Of BOMRow, Double, String))
            Dim reference As New List(Of BOMRow)

            If bomRows Is Nothing Then Return New List(Of BOMRow)

            For Each row As BOMRow In bomRows
                If row Is Nothing Then Continue For

                Dim doc As Document = Nothing
                Try
                    If row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then
                        If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                            reference.Add(row)
                        End If
                        Continue For
                    End If
                    doc = row.ComponentDefinitions.Item(1).Document
                Catch
                    Continue For
                End Try
                If doc Is Nothing Then Continue For

                Dim isAsm As Boolean = False
                Dim isPart As Boolean = False
                Try
                    isAsm = (doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
                    isPart = (doc.DocumentType = DocumentTypeEnum.kPartDocumentObject)
                Catch
                End Try

                Dim searchText As String = GetSearchText(row)
                Dim isFast As Boolean = IsFastener(searchText)
                Dim isBear As Boolean = IsBearing(searchText)
                Dim pn As String = GetPartNumber(row)
                If String.IsNullOrEmpty(pn) Then pn = ""

                If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                    reference.Add(row)
                    Continue For
                End If

                If row.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                    Dim m As Double = GetMass(doc)
                    If isAsm Then
                        phantomAsm.Add(Tuple.Create(row, m, pn))
                    ElseIf isPart Then
                        phantomPart.Add(Tuple.Create(row, m, pn))
                    Else
                        reference.Add(row)
                    End If
                    Continue For
                End If

                If row.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                    Dim prio As Integer = 0
                    If isFast Then prio = 2
                    If isBear Then prio = 1
                    If isAsm Then
                        purchasedAsm.Add(Tuple.Create(row, prio, pn))
                    Else
                        purchasedPart.Add(Tuple.Create(row, prio, pn))
                    End If
                    Continue For
                End If

                Dim mass As Double = GetMass(doc)
                If isAsm Then
                    normalAsm.Add(Tuple.Create(row, mass, pn))
                ElseIf isPart Then
                    normalPart.Add(Tuple.Create(row, mass, pn))
                Else
                    reference.Add(row)
                End If
            Next

            '===== SẮP XẾP RIÊNG CỤM LẮP =====
            ApplySort(normalAsm, sortModeAsm)
            ApplySort(phantomAsm, sortModeAsm)

            '===== SẮP XẾP RIÊNG PART =====
            ApplySort(normalPart, sortModePart)
            ApplySort(phantomPart, sortModePart)

            ' Purchased giữ logic cũ (Bearing → Fastener → còn lại)
            purchasedAsm.Sort(Function(a, b)
                                  Dim c = a.Item2.CompareTo(b.Item2)
                                  If c <> 0 Then Return c
                                  Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                              End Function)
            purchasedPart.Sort(Function(a, b)
                                   Dim c = a.Item2.CompareTo(b.Item2)
                                   If c <> 0 Then Return c
                                   Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                               End Function)

            reference.Sort(Function(a, b) String.Compare(GetPartNumber(a), GetPartNumber(b), StringComparison.OrdinalIgnoreCase))

            Dim result As New List(Of BOMRow)
            For Each x In normalAsm : result.Add(x.Item1) : Next
            For Each x In purchasedAsm : result.Add(x.Item1) : Next
            For Each x In normalPart : result.Add(x.Item1) : Next
            For Each x In purchasedPart : result.Add(x.Item1) : Next
            For Each x In phantomAsm : result.Add(x.Item1) : Next
            For Each x In phantomPart : result.Add(x.Item1) : Next
            For Each x In reference : result.Add(x) : Next

            Return result
        End Function

        ' Hàm hỗ trợ sort theo mode
        Private Sub ApplySort(list As List(Of Tuple(Of BOMRow, Double, String)), mode As Integer)
            Select Case mode
                Case 1  ' Mass ASC (bé → lớn)
                    list.Sort(Function(a, b) a.Item2.CompareTo(b.Item2))

                Case 2  ' Tên ngắn → dài (theo độ dài, rồi A→Z)
                    list.Sort(Function(a, b)
                                  Dim c = a.Item3.Length.CompareTo(b.Item3.Length)
                                  If c <> 0 Then Return c
                                  Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                              End Function)

                Case 3  ' Tên dài → ngắn (theo độ dài, rồi Z→A)
                    list.Sort(Function(a, b)
                                  Dim c = b.Item3.Length.CompareTo(a.Item3.Length)
                                  If c <> 0 Then Return c
                                  Return String.Compare(b.Item3, a.Item3, StringComparison.OrdinalIgnoreCase)
                              End Function)

                Case 4  ' Chữ cái A → Z
                    list.Sort(Function(a, b) String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase))

                Case 5  ' Chữ cái Z → A
                    list.Sort(Function(a, b) String.Compare(b.Item3, a.Item3, StringComparison.OrdinalIgnoreCase))

                Case Else  ' 0 = Mass DESC (lớn → bé) - mặc định
                    list.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            End Select
        End Sub
        Private Function SortRows1(ByVal bomRows As BOMRowsEnumerator,
                          Optional ByVal sortMode As Integer = 0) As List(Of BOMRow)

            Dim normalAsm As New List(Of Tuple(Of BOMRow, Double, String))
            Dim purchasedAsm As New List(Of Tuple(Of BOMRow, Integer, String))
            Dim normalPart As New List(Of Tuple(Of BOMRow, Double, String))
            Dim purchasedPart As New List(Of Tuple(Of BOMRow, Integer, String))
            Dim phantomAsm As New List(Of Tuple(Of BOMRow, Double, String))
            Dim phantomPart As New List(Of Tuple(Of BOMRow, Double, String))
            Dim reference As New List(Of BOMRow)

            If bomRows Is Nothing Then Return New List(Of BOMRow)

            For Each row As BOMRow In bomRows
                If row Is Nothing Then Continue For

                Dim doc As Document = Nothing
                Try
                    If row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then
                        If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                            reference.Add(row)
                        End If
                        Continue For
                    End If
                    doc = row.ComponentDefinitions.Item(1).Document
                Catch
                    Continue For
                End Try
                If doc Is Nothing Then Continue For

                Dim isAsm As Boolean = False
                Dim isPart As Boolean = False
                Try
                    isAsm = (doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
                    isPart = (doc.DocumentType = DocumentTypeEnum.kPartDocumentObject)
                Catch
                End Try

                Dim searchText As String = GetSearchText(row)
                Dim isFast As Boolean = IsFastener(searchText)
                Dim isBear As Boolean = IsBearing(searchText)
                Dim pn As String = GetPartNumber(row)
                If String.IsNullOrEmpty(pn) Then pn = ""

                If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                    reference.Add(row)
                    Continue For
                End If

                If row.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                    Dim m As Double = GetMass(doc)
                    If isAsm Then
                        phantomAsm.Add(Tuple.Create(row, m, pn))
                    ElseIf isPart Then
                        phantomPart.Add(Tuple.Create(row, m, pn))
                    Else
                        reference.Add(row)
                    End If
                    Continue For
                End If

                If row.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                    Dim prio As Integer = 0
                    If isFast Then prio = 2
                    If isBear Then prio = 1
                    If isAsm Then
                        purchasedAsm.Add(Tuple.Create(row, prio, pn))
                    Else
                        purchasedPart.Add(Tuple.Create(row, prio, pn))
                    End If
                    Continue For
                End If

                Dim mass As Double = GetMass(doc)
                If isAsm Then
                    normalAsm.Add(Tuple.Create(row, mass, pn))
                ElseIf isPart Then
                    normalPart.Add(Tuple.Create(row, mass, pn))
                Else
                    reference.Add(row)
                End If
            Next

            '===== SẮP XẾP THEO sortMode =====
            ' 0 = Mass DESC (lớn → bé)
            ' 1 = Mass ASC  (bé → lớn)
            ' 2 = Name ASC  (ngắn → dài, rồi A→Z)
            ' 3 = Name DESC (dài → ngắn, rồi Z→A)

            Select Case sortMode
                Case 1  ' Mass ASC
                    normalAsm.Sort(Function(a, b) a.Item2.CompareTo(b.Item2))
                    normalPart.Sort(Function(a, b) a.Item2.CompareTo(b.Item2))
                    phantomAsm.Sort(Function(a, b) a.Item2.CompareTo(b.Item2))
                    phantomPart.Sort(Function(a, b) a.Item2.CompareTo(b.Item2))

                Case 2  ' Name short → long
                    normalAsm.Sort(Function(a, b)
                                       Dim c = a.Item3.Length.CompareTo(b.Item3.Length)
                                       If c <> 0 Then Return c
                                       Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                                   End Function)
                    normalPart.Sort(Function(a, b)
                                        Dim c = a.Item3.Length.CompareTo(b.Item3.Length)
                                        If c <> 0 Then Return c
                                        Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                                    End Function)
                    phantomAsm.Sort(Function(a, b)
                                        Dim c = a.Item3.Length.CompareTo(b.Item3.Length)
                                        If c <> 0 Then Return c
                                        Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                                    End Function)
                    phantomPart.Sort(Function(a, b)
                                         Dim c = a.Item3.Length.CompareTo(b.Item3.Length)
                                         If c <> 0 Then Return c
                                         Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                                     End Function)

                Case 3  ' Name long → short
                    normalAsm.Sort(Function(a, b)
                                       Dim c = b.Item3.Length.CompareTo(a.Item3.Length)
                                       If c <> 0 Then Return c
                                       Return String.Compare(b.Item3, a.Item3, StringComparison.OrdinalIgnoreCase)
                                   End Function)
                    normalPart.Sort(Function(a, b)
                                        Dim c = b.Item3.Length.CompareTo(a.Item3.Length)
                                        If c <> 0 Then Return c
                                        Return String.Compare(b.Item3, a.Item3, StringComparison.OrdinalIgnoreCase)
                                    End Function)
                    phantomAsm.Sort(Function(a, b)
                                        Dim c = b.Item3.Length.CompareTo(a.Item3.Length)
                                        If c <> 0 Then Return c
                                        Return String.Compare(b.Item3, a.Item3, StringComparison.OrdinalIgnoreCase)
                                    End Function)
                    phantomPart.Sort(Function(a, b)
                                         Dim c = b.Item3.Length.CompareTo(a.Item3.Length)
                                         If c <> 0 Then Return c
                                         Return String.Compare(b.Item3, a.Item3, StringComparison.OrdinalIgnoreCase)
                                     End Function)

                Case Else  ' 0 = Mass DESC (mặc định cũ)
                    normalAsm.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
                    normalPart.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
                    phantomAsm.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
                    phantomPart.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            End Select

            ' Purchased vẫn ưu tiên Bearing → Fastener → còn lại (giữ nguyên logic cũ)
            purchasedAsm.Sort(Function(a, b)
                                  Dim c = a.Item2.CompareTo(b.Item2)
                                  If c <> 0 Then Return c
                                  Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                              End Function)
            purchasedPart.Sort(Function(a, b)
                                   Dim c = a.Item2.CompareTo(b.Item2)
                                   If c <> 0 Then Return c
                                   Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                               End Function)

            reference.Sort(Function(a, b) String.Compare(GetPartNumber(a), GetPartNumber(b), StringComparison.OrdinalIgnoreCase))

            Dim result As New List(Of BOMRow)
            For Each x In normalAsm : result.Add(x.Item1) : Next
            For Each x In purchasedAsm : result.Add(x.Item1) : Next
            For Each x In normalPart : result.Add(x.Item1) : Next
            For Each x In purchasedPart : result.Add(x.Item1) : Next
            For Each x In phantomAsm : result.Add(x.Item1) : Next
            For Each x In phantomPart : result.Add(x.Item1) : Next
            For Each x In reference : result.Add(x) : Next

            Return result
        End Function

        Private Function GetMass(ByVal doc As Document) As Double
            Try
                If doc Is Nothing Then Return 0
                If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Return CType(doc, AssemblyDocument).ComponentDefinition.MassProperties.Mass
                ElseIf doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Return CType(doc, PartDocument).ComponentDefinition.MassProperties.Mass
                End If
            Catch
            End Try
            Return 0
        End Function

        Private Function GetPartNumber(ByVal row As BOMRow) As String
            Try
                If row Is Nothing OrElse row.ComponentDefinitions Is Nothing OrElse
                   row.ComponentDefinitions.Count = 0 Then Return ""
                Return GetProperty(row.ComponentDefinitions.Item(1).Document, "Part Number")
            Catch
                Return ""
            End Try
        End Function

        Private Function PickFromList(ByVal title As String,
                                      ByVal items As String(),
                                      Optional ByVal defaultIndex As Integer = 0) As Integer
            Dim frm As New Form()
            Try
                frm.Text = title
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.FormBorderStyle = FormBorderStyle.FixedDialog
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.Width = 460
                frm.Height = 310

                Dim lst As New ListBox()
                lst.Left = 12 : lst.Top = 12
                lst.Width = 420 : lst.Height = 200
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
                btnOK.Left = 250 : btnOK.Top = 225
                btnOK.Width = 85 : btnOK.Height = 30
                btnOK.DialogResult = DialogResult.OK

                Dim btnCancel As New Button()
                btnCancel.Text = "Hủy"
                btnCancel.Left = 345 : btnCancel.Top = 225
                btnCancel.Width = 85 : btnCancel.Height = 30
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



