
Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Windows.Forms
Imports Inventor
Imports ToolInventor2025.ToolInventor2025.Assembly.Buttons

Namespace ToolInventor2025.Drawing.Buttons
    Public Module Draw_61111
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication
            Try
                '=================================================
                ' KIỂM TRA DRAWING
                '=================================================
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then

                    MessageBox.Show(
                        "Vui lòng mở file Drawing (.idw)!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As Inventor.DrawingDocument = CType(app.ActiveDocument, Inventor.DrawingDocument)
                Dim oSheet As Inventor.Sheet = oDrawDoc.ActiveSheet

                If oSheet.PartsLists.Count < 1 Then
                    MessageBox.Show(
                        "Sheet hiện tại không có Parts List.",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)
                    Exit Sub
                End If
                '=================================================
                ' 1. CHẾ ĐỘ TÊN
                '=================================================
                Dim nameModeIdx As Integer =
                    PickFromList(
                        "Xử lý cột Tên",
                        New String() {
                            "1 - Part Number: không ghi đè nếu Tên đã là PN/SN",
                            "2 - Part Number: chỉ ghi khi ô Tên đang trống",
                            "3 - Stock Number: đồng bộ trực tiếp vào BOM",
                            "4 - Không sửa tên"}, 0)

                If nameModeIdx < 0 Then
                    Exit Sub
                End If

                Dim nameMode As Integer = nameModeIdx + 1

                '=================================================
                ' 2. VẬT LIỆU
                '=================================================
                Dim matDefault As String =
                    InputBox("Vật liệu mặc định (Part tự chế):", "Vật liệu", "SS400")

                If matDefault Is Nothing Then
                    matDefault = ""
                End If
                matDefault = matDefault.Trim()

                '=================================================
                ' 2b. XỬ LÝ VẬT LIỆU PURCHASED (Yes/No)
                '=================================================
                Dim result As DialogResult = MessageBox.Show(
                   "Xóa vật liệu của vật tư mua (Purchased)?" & vbCrLf &
                      "Yes = Xóa" & vbCrLf &
                          "No  = Để nguyên",
                    "Vật liệu Purchased",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                Dim clearPurchasedMaterial As Boolean = (result = DialogResult.Yes)
                '=================================================
                ' 3. PHẠM VI
                '=================================================
                Dim scopeIdx As Integer =
                    PickFromList(
                        "Phạm vi",
                        New String() {
                            "1 - Chỉ Parts List đầu trên sheet active",
                            "2 - Tất cả Parts List trên sheet active",
                            "3 - Tất cả Parts List của toàn bộ Drawing"}, 0)

                If scopeIdx < 0 Then
                    Exit Sub
                End If

                '=================================================
                ' TÊN CỘT
                '=================================================
                Dim colSTT As String = "STT"
                Dim colTen As String = "Tên chi tiết"
                Dim colTen2 As String = "Tên gọi"
                Dim colDonVi As String = "Đơn vị"
                Dim colVL As String = "Vật liệu"
                Dim colUnitQty As String = "UNIT QTY"

                Dim processed As Integer = 0


                '=================================================
                ' XỬ LÝ PHẠM VI
                '=================================================

                '-------------------------------------------------
                ' 1 - PARTS LIST ĐẦU TIÊN SHEET ACTIVE
                '-------------------------------------------------
                If scopeIdx = 0 Then

                    Try

                        Dim oPartList As Inventor.PartsList = oSheet.PartsLists.Item(1)


                        'MODE 3:
                        'Stock Number -> BOM
                        If nameMode = 3 Then
                            SyncStockNumberToBOM(oPartList)

                        End If


                        ProcessOnePartsList(oPartList, nameMode, matDefault, clearPurchasedMaterial,
                    colSTT, colTen, colTen2, colDonVi, colVL, colUnitQty)

                        processed += 1


                    Catch ex As Exception

                        MessageBox.Show("Parts List 1:" & vbCrLf & ex.Message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                    End Try

                    '-------------------------------------------------
                    ' 2 - TẤT CẢ PARTS LIST SHEET ACTIVE
                    '-------------------------------------------------
                ElseIf scopeIdx = 1 Then

                    For plIdx As Integer = 1 To oSheet.PartsLists.Count

                        Try

                            Dim oPartList As Inventor.PartsList = oSheet.PartsLists.Item(plIdx)

                            'MODE 3
                            If nameMode = 3 Then

                                SyncStockNumberToBOM(oPartList)

                            End If


                            ProcessOnePartsList(oPartList, nameMode, matDefault, clearPurchasedMaterial, colSTT, colTen, colTen2, colDonVi, colVL, colUnitQty)


                            processed += 1


                        Catch exPL As Exception

                            MessageBox.Show("Sheet: " & oSheet.Name & vbCrLf & "Parts List: " & plIdx.ToString() &
                                vbCrLf & exPL.Message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                        End Try

                    Next

                    '-------------------------------------------------
                    ' 3 - TẤT CẢ PARTS LIST TOÀN BỘ DRAWING
                    '-------------------------------------------------
                ElseIf scopeIdx = 2 Then

                    For sheetIdx As Integer = 1 To oDrawDoc.Sheets.Count

                        Try

                            Dim oCurSheet As Inventor.Sheet = oDrawDoc.Sheets.Item(sheetIdx)


                            For plIdx As Integer = 1 To oCurSheet.PartsLists.Count

                                Try

                                    Dim oPartList As Inventor.PartsList = oCurSheet.PartsLists.Item(plIdx)


                                    'MODE 3
                                    If nameMode = 3 Then

                                        SyncStockNumberToBOM(oPartList)

                                    End If


                                    ProcessOnePartsList(oPartList, nameMode, matDefault, clearPurchasedMaterial, colSTT, colTen, colTen2, colDonVi, colVL, colUnitQty)


                                    processed += 1


                                Catch exPL As Exception
                                    MessageBox.Show("Sheet: " & oCurSheet.Name & vbCrLf & "Parts List: " &
                                        plIdx.ToString() & vbCrLf & exPL.Message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                                End Try

                            Next


                        Catch exSheet As Exception

                            MessageBox.Show("Lỗi Sheet " & sheetIdx.ToString() & ":" & vbCrLf & exSheet.Message, "Cảnh báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)

                        End Try

                    Next

                End If


                '=================================================
                ' UPDATE DRAWING CUỐI CÙNG
                '=================================================
                Try
                    oDrawDoc.Update()
                Catch
                End Try


                '=================================================
                ' THÔNG BÁO
                '=================================================
                MessageBox.Show("Hoàn tất!" & vbCrLf &
    "Parts List đã xử lý: " & processed.ToString() &
    vbCrLf & "Chế độ tên: " & nameMode.ToString() &
    vbCrLf & "Nguồn tên: " & If(nameMode = 3, "Stock Number / đồng bộ từ Part Number",
        If(nameMode = 4, "Không sửa", "Part Number")) &
    vbCrLf & "VL mặc định: " & If(matDefault = "", "(không dùng)", matDefault) &
    vbCrLf & "VL Purchased: " & If(clearPurchasedMaterial, "Xóa", "Để nguyên"),
    "Override Parts List",
    MessageBoxButtons.OK,
    MessageBoxIcon.Information)


            Catch ex As Exception

                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Override Parts List", MessageBoxButtons.OK, MessageBoxIcon.Error)

            End Try

        End Sub


        '=========================================================
        ' XỬ LÝ 1 PARTS LIST
        '=========================================================
        Private Sub ProcessOnePartsList(oPartList As Inventor.PartsList, nameMode As Integer, matDefault As String,
    clearPurchasedMaterial As Boolean,          ' ← thêm dòng này
    colSTT As String, colTen As String, colTen2 As String, colDonVi As String, colVL As String, colUnitQty As String)

            '=====================================================
            ' TÌM CỘT
            '=====================================================
            Dim cSTT As String = FindColumn(oPartList, New String() {colSTT, "Item", "ITEM"})

            Dim cTen As String = FindColumn(oPartList, New String() {colTen, colTen2, "Part Number", "Stock Number", "Tên"})
            Dim nameCol As String = cTen          ' cache cột tên ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' tối ưu code l1

            Dim cDonVi As String = FindColumn(oPartList, New String() {colDonVi, "Keywords", "Unit", "ĐƠN VỊ"})

            Dim cVL As String = FindColumn(oPartList, New String() {colVL, "Material", "MATERIAL", "VẬT LIỆU"})

            Dim cUnitQty As String = FindUnitQtyColumn(oPartList)


            If cSTT = "" OrElse cDonVi = "" Then

                Throw New Exception("Không tìm thấy cột STT hoặc Đơn vị trên Parts List.")

            End If

            '=====================================================
            ' DUYỆT ROW
            '=====================================================
            For i As Integer = 1 To oPartList.PartsListRows.Count

                Try

                    Dim row As Inventor.PartsListRow = oPartList.PartsListRows.Item(i)

                    If row.ReferencedRows Is Nothing OrElse
                       row.ReferencedRows.Count < 1 Then
                        Continue For
                    End If


                    Dim bomRow As Inventor.BOMRow = row.ReferencedRows.Item(1).BOMRow

                    If bomRow Is Nothing Then
                        Continue For
                    End If


                    '=================================================
                    ' BOM STRUCTURE
                    '=================================================
                    Dim bs As Inventor.BOMStructureEnum = bomRow.BOMStructure

                    Dim isPurchased As Boolean = (bs = Inventor.BOMStructureEnum.kPurchasedBOMStructure)


                    '=================================================
                    ' DOCUMENT
                    '=================================================
                    Dim refDoc As Inventor.Document = Nothing


                    Dim docType As Inventor.DocumentTypeEnum =
                        Inventor.DocumentTypeEnum.kUnknownDocumentObject


                    Try

                        If bomRow.ComponentDefinitions.Count > 0 Then

                            refDoc = bomRow.ComponentDefinitions.Item(1).Document


                            If refDoc IsNot Nothing Then
                                docType = refDoc.DocumentType

                            End If

                        End If

                    Catch
                    End Try

                    Dim isAsm As Boolean = (docType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject)
                    Dim isPart As Boolean = (docType = Inventor.DocumentTypeEnum.kPartDocumentObject)

                    '=================================================
                    ' PN / SN
                    '=================================================
                    Dim pn As String = ""
                    Dim sn As String = ""


                    If refDoc IsNot Nothing Then

                        pn = GetProp(refDoc, "Part Number")
                        sn = GetProp(refDoc, "Stock Number")

                    End If

                    pn = If(pn, "").Trim()
                    sn = If(sn, "").Trim()

                    '=================================================
                    ' TÊN HIỆN TẠI
                    '=================================================
                    Dim currentName As String = GetCellValue(row, cTen).Trim()

                    '=================================================
                    ' MODE 1 / MODE 2
                    '
                    ' Mode 3 đã đồng bộ BOM trước đó.
                    '=================================================
                    If nameMode <> 3 Then

                        ' ApplyNameLogic(row, currentName, pn, sn, nameMode)
                        ApplyNameLogic(row, currentName, pn, sn, nameMode, nameCol) ''''''''''''''''''''''''''''''''''''''''' toi ưu lan 1
                    End If


                    '=================================================
                    ' ĐƠN VỊ
                    '=================================================
                    If isAsm Then

                        SetCell(row, cDonVi, "Bộ")


                    ElseIf isPart AndAlso Not isPurchased Then

                        Dim nameForUnit As String = currentName

                        If nameForUnit = "" Then

                            If pn <> "" Then

                                nameForUnit = pn
                            Else
                                nameForUnit = sn
                            End If
                        End If
                        Dim donVi As String = GuessUnit(nameForUnit)
                        SetCell(row, cDonVi, donVi)
                    ElseIf isPurchased Then
                        SetCell(row, cDonVi, "Cái")
                    End If


                    '=================================================
                    ' MATERIAL
                    '
                    ' ASSEMBLY:
                    '   KHÔNG ĐỤNG
                    '
                    ' PURCHASED:
                    '   KHÔNG ĐỤNG
                    '
                    ' PART TỰ CHẾ:
                    '   GHI ĐÈ
                    '=================================================
                    '    If isPart AndAlso Not isPurchased Then

                    '        If matDefault <> "" Then

                    '  SetCell(row, cVL, matDefault)

                    '           End If
                    '           End If

                    If isAsm Then
                        ' không đụng

                    ElseIf isPurchased Then
                        If clearPurchasedMaterial Then
                            ClearCell(row, cVL)          ' Yes → Xóa
                        End If
                        ' No → Để nguyên

                    ElseIf isPart Then
                        If matDefault <> "" Then
                            SetCell(row, cVL, matDefault)
                        End If
                    End If
                Catch
                    '        'Bỏ qua row lỗi
                End Try

            Next


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' code cũ ko dùng vẫn chạy ổn nhưng lâu 
            '=====================================================
            ' UPDATE PARTS LIST TRƯỚC
            '=====================================================
            ' Try
            '   oPartList.Parent.Update()

            '    Catch
            '   End Try

            '       Try
            '         oPartList.Update()
            '          Catch
            '         End Try
            '         '=====================================================
            ' UNIT QTY
            '
            ' <= 1 -> XÓA
            ' > 1  -> GIỮ
            '=====================================================
            '         If cUnitQty <> "" Then

            '        For qtyIdx As Integer = 1 To oPartList.PartsListRows.Count

            '       Try
            '       Dim qtyRow As Inventor.PartsListRow = oPartList.PartsListRows.Item(qtyIdx)
            '                  Dim qtyText As String = GetCellValue(qtyRow, cUnitQty)

            '      If qtyText <> "" Then
            '   Dim qty As Double = 0
            '      If TryParseNumber(qtyText, qty) Then
            '     If qty <= 1 Then
            '                              ClearCell(qtyRow, cUnitQty)
            '     End If
            '    End If

            '   End If

            '   Catch
            'Bỏ qua row lỗi
            'End Try

            '   Next

            '  End If

            '=====================================================
            ' SAVE OVERRIDE
            '=====================================================
            '      Try

            '   oPartList.SaveItemOverridesToBOM()

            '    Catch
            '    End Try

            '=====================================================
            ' STT
            '
            ' CỤM → PART → PURCHASED
            '=====================================================
            '    Dim stt As Integer = 1

            '    stt = NumberRows(oPartList, cSTT, stt, True, False)
            '   stt = NumberRows(oPartList, cSTT, stt, False, False)
            '   stt = NumberRows(oPartList, cSTT, stt, False, True)


            '=====================================================
            ' SORT
            '=====================================================
            '    Try
            '    oPartList.Sort(cSTT)
            '    Catch
            '    Try
            '   oPartList.Sort("Item")
            '   Catch
            '  End Try

            '  End Try
            '=====================================================
            ' SAVE LẦN CUỐI
            '=====================================================
            '    Try
            '     oPartList.SaveItemOverridesToBOM()
            '     Catch
            '     End Try
            '
            '   End Sub

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' code cũ ko dùng vẫn chạy ổn nhưng lâu 



            '=====================================================
            ' UPDATE 1 LẦN TRƯỚC KHI XỬ LÝ UNIT QTY
            '=====================================================
            Try
                oPartList.Update()
            Catch
            End Try

            '=====================================================
            ' UNIT QTY
            ' <= 1 -> XÓA
            ' > 1  -> GIỮ
            '=====================================================
            If cUnitQty <> "" Then
                For qtyIdx As Integer = 1 To oPartList.PartsListRows.Count
                    Try
                        Dim qtyRow As Inventor.PartsListRow = oPartList.PartsListRows.Item(qtyIdx)
                        Dim qtyText As String = GetCellValue(qtyRow, cUnitQty)

                        If qtyText <> "" Then
                            Dim qty As Double = 0
                            If TryParseNumber(qtyText, qty) Then
                                If qty <= 1 Then
                                    ClearCell(qtyRow, cUnitQty)
                                End If
                            End If
                        End If
                    Catch
                    End Try
                Next
            End If

            '=====================================================
            ' STT
            ' CỤM → PART → PURCHASED
            '=====================================================
            Dim stt As Integer = 1
            stt = NumberRows(oPartList, cSTT, stt, True, False)
            stt = NumberRows(oPartList, cSTT, stt, False, False)
            stt = NumberRows(oPartList, cSTT, stt, False, True)

            '=====================================================
            ' SORT
            '=====================================================
            Try
                oPartList.Sort(cSTT)
            Catch
                Try
                    oPartList.Sort("Item")
                Catch
                End Try
            End Try

            '=====================================================
            ' SAVE + UPDATE LẦN CUỐI
            '=====================================================
            Try
                oPartList.SaveItemOverridesToBOM()
            Catch
            End Try

            Try
                oPartList.Update()
            Catch
            End Try



        End Sub























        '=========================================================
        ' MODE 1 / MODE 2
        '
        ' NGUỒN CHÍNH = PART NUMBER
        '
        ' Nếu PN trống:
        '   fallback SN.
        '=========================================================
        Private Sub ApplyNameLogic(row As Inventor.PartsListRow, currentName As String, pn As String, sn As String, nameMode As Integer, nameCol As String)  ''''''''''''''''''''''''''''''''''''''''' toi ưu lan 1

            Dim cur As String = If(currentName, "").Trim()

            Dim requiredName As String = ""
            If pn <> "" Then
                requiredName = pn
            Else
                requiredName = sn
            End If

            Select Case nameMode

                '=================================================
                ' MODE 1
                '=================================================
                Case 1

                    If cur <> "" AndAlso
                       (
                           String.Equals(cur, pn, StringComparison.OrdinalIgnoreCase) OrElse
                           String.Equals(cur, sn, StringComparison.OrdinalIgnoreCase)
                       ) Then
                        Exit Sub

                    End If

                    If requiredName <> "" Then

                        ' SetCell(row, GetNameColumn(row), requiredName)

                        SetCell(row, nameCol, requiredName) ''''''''''''''''''''''''''''''''''''''''' toi ưu lan 1

                    End If
                '   ===============================================
                ' MODE 2
                '                ' CHỈ KHI Ô TÊN TRỐNG
                '=================================================
                Case 2

                    If cur <> "" Then
                        Exit Sub
                    End If

                    If requiredName <> "" Then

                        '   SetCell(row, GetNameColumn(row), requiredName)
                        SetCell(row, nameCol, requiredName) ''''''''''''''''''''''''''''''''''''''''' toi ưu lan 1
                    End If


                '=================================================
                ' MODE 4
                '=================================================
                Case 4

                    Exit Sub

            End Select

        End Sub


        '=========================================================
        ' MODE 3
        '
        ' STOCK NUMBER
        '
        ' Nếu:
        '
        ' Part Number = ABC
        ' Stock Number = ""
        '
        ' =>
        '
        ' Stock Number = ABC
        '
        ' Ghi trực tiếp vào Property của component.
        '=========================================================
        Private Sub SyncStockNumberToBOM(
            oPartList As Inventor.PartsList)

            For i As Integer = 1 To oPartList.PartsListRows.Count

                Try

                    Dim row As Inventor.PartsListRow = oPartList.PartsListRows.Item(i)

                    If row.ReferencedRows Is Nothing OrElse
                       row.ReferencedRows.Count < 1 Then

                        Continue For

                    End If

                    Dim bomRow As Inventor.BOMRow = row.ReferencedRows.Item(1).BOMRow

                    If bomRow Is Nothing Then
                        Continue For
                    End If

                    If bomRow.ComponentDefinitions.Count < 1 Then
                        Continue For
                    End If

                    Dim compDef As Inventor.ComponentDefinition = bomRow.ComponentDefinitions.Item(1)

                    If compDef Is Nothing Then
                        Continue For
                    End If

                    Dim refDoc As Inventor.Document = compDef.Document


                    If refDoc Is Nothing Then
                        Continue For
                    End If

                    '=================================================
                    ' ĐỌC PN / SN TRỰC TIẾP
                    '=================================================
                    Dim pn As String = GetProp(refDoc, "Part Number")
                    Dim sn As String = GetProp(refDoc, "Stock Number")


                    pn = If(pn, "").Trim()
                    sn = If(sn, "").Trim()


                    '=================================================
                    ' SN TRỐNG
                    '
                    ' COPY PN -> SN
                    '=================================================
                    If sn = "" AndAlso
                       pn <> "" Then

                        SetProp(refDoc, "Stock Number", pn)

                        '---------------------------------------------
                        ' UPDATE DOCUMENT NGAY
                        '---------------------------------------------
                        Try

                            refDoc.Update()

                        Catch
                        End Try

                    End If


                Catch
                    'Bỏ qua row lỗi
                End Try

            Next


            '=====================================================
            ' UPDATE BOM / PARTS LIST
            '=====================================================
            Try

                oPartList.Parent.Update()

            Catch
            End Try


            Try

                oPartList.Update()

            Catch
            End Try


            '=====================================================
            ' SAVE BOM
            '=====================================================
            Try

                oPartList.SaveItemOverridesToBOM()

            Catch
            End Try


            '=====================================================
            ' UPDATE LẦN 2
            '=====================================================
            Try

                oPartList.Parent.Update()

            Catch
            End Try


            Try

                oPartList.Update()

            Catch
            End Try

        End Sub


        '=========================================================
        ' LẤY TÊN CỘT
        '=========================================================
        Private Function GetNameColumn(row As Inventor.PartsListRow) As String

            Try

                Dim value As String = CStr(row.Item("Tên chi tiết").Value)

                Return "Tên chi tiết"

            Catch
            End Try


            Try
                Dim value As String = CStr(row.Item("Tên gọi").Value)
                Return "Tên gọi"

            Catch
            End Try

            Try

                Dim value As String = CStr(row.Item("Tên").Value)

                Return "Tên"

            Catch
            End Try

            Return "Tên chi tiết"

        End Function


        '=========================================================
        ' ĐƠN VỊ
        '=========================================================
        Private Function GuessUnit2(ten As String) As String  ''''''''''''''''''code cũ hiện ko dùng

            If String.IsNullOrEmpty(ten) Then

                Return "Cái"

            End If

            Dim t As String = ten.Trim()

            '=====================================================
            ' THANH
            '=====================================================
            If t.Length >= 3 AndAlso
               (
                   (t.StartsWith("Th", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("Tr", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("TR", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("XG", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("Ốn", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("TH", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("Câ", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("Xg", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("Pi", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("Shaf", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("U", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("C", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("I", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("H", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("V", StringComparison.OrdinalIgnoreCase) OrElse
                       t.StartsWith("L", StringComparison.OrdinalIgnoreCase)
                                             ) AndAlso
                   t.EndsWith("L", StringComparison.OrdinalIgnoreCase) OrElse
                     t.EndsWith("mm", StringComparison.OrdinalIgnoreCase)
               ) Then

                Return "Thanh"

            End If


            If t.Length >= 2 Then

                Dim c0 As Char = Char.ToUpperInvariant(t(0))


                If "TPVLHIZCUS".IndexOf(c0) >= 0 AndAlso
                 t.EndsWith("L", StringComparison.OrdinalIgnoreCase) OrElse
                     t.EndsWith("mm", StringComparison.OrdinalIgnoreCase) Then
                    Return "Thanh"

                End If
                '     Select Case c0
                ' Case "T"c   ' Thép
                ' Case "P"c   ' Phi
                '  Case "V"c   ' Vít
                '  Case "L"c   ' Lưới
                '  Case "H"c   ' Hộp
                '  Case "I"c   ' Inox
                '  Case "Z"c   ' Ký hiệu riêng
                '  Case "C"c   ' Cột
                '  Case "U"c   ' U (thanh chữ U)
                '  If t.EndsWith("L", StringComparison.OrdinalIgnoreCase) Then
                '  Return "Thanh"
                'E'nd If
                '   End Select
            End If '


            '=====================================================
            ' TẤM
            '=====================================================
            If t.StartsWith("PL", StringComparison.OrdinalIgnoreCase) OrElse
               t.StartsWith("Tô", StringComparison.OrdinalIgnoreCase) OrElse
               t.StartsWith("Tấ", StringComparison.OrdinalIgnoreCase) OrElse
               t.StartsWith("Mã", StringComparison.OrdinalIgnoreCase) OrElse
               t.StartsWith("Bi", StringComparison.OrdinalIgnoreCase) Then

                Return "Tấm"

            End If

            Return "Cái"

        End Function ''''''''''''''''''code cũ hiện ko dùng



        ''''''''''''''''''code mới đang dùng
        Private Function GuessUnit(ten As String) As String
            If String.IsNullOrEmpty(ten) Then Return "Cái"

            Dim t As String = ten.Trim()
            Dim len As Integer = t.Length
            If len = 0 Then Return "Cái"

            ' Kiểm tra kết thúc trước (nhanh nhất)
            Dim last As Char = Char.ToUpperInvariant(t(len - 1))
            Dim endsWithL As Boolean = (last = "L"c)
            Dim endsWithMM As Boolean = (len >= 2 AndAlso last = "M"c AndAlso Char.ToUpperInvariant(t(len - 2)) = "M"c)

            If endsWithL OrElse endsWithMM Then
                Dim c0 As Char = Char.ToUpperInvariant(t(0))

                ' Tiền tố 2 ký tự phổ biến
                If len >= 2 Then
                    Dim c1 As Char = Char.ToUpperInvariant(t(1))
                    If (c0 = "T"c AndAlso (c1 = "H"c OrElse c1 = "R"c)) OrElse
               (c0 = "X"c AndAlso c1 = "G"c) OrElse
               (c0 = "P"c AndAlso c1 = "I"c) OrElse
               (c0 = "S"c AndAlso c1 = "H"c) OrElse
               (c0 = "C"c AndAlso c1 = "Â"c) OrElse
               (c0 = "Ố"c AndAlso c1 = "N"c) Then
                        Return "Thanh"
                    End If
                End If

                ' Ký tự đơn
                If "TPVLHIZCUS".IndexOf(c0) >= 0 Then Return "Thanh"
            End If

            ' Tấm
            If len >= 2 Then
                Dim c0 As Char = Char.ToUpperInvariant(t(0))
                Dim c1 As Char = Char.ToUpperInvariant(t(1))
                If (c0 = "P"c AndAlso c1 = "L"c) OrElse
           (c0 = "T"c AndAlso (c1 = "Ô"c OrElse c1 = "Ấ"c)) OrElse
           (c0 = "M"c AndAlso c1 = "Ã"c) OrElse
           (c0 = "B"c AndAlso c1 = "I"c) Then
                    Return "Tấm"
                End If
            End If

            Return "Cái"
        End Function





        '=========================================================
        ' FIND COLUMN
        '=========================================================
        Private Function FindColumn(pl As Inventor.PartsList, candidates As String()) As String

            For Each name As String In candidates
                Try
                    Dim col As Inventor.PartsListColumn = pl.PartsListColumns.Item(name)

                    If col IsNot Nothing Then

                        Return name

                    End If

                Catch
                End Try

            Next

            Try

                For Each col As Inventor.PartsListColumn In
                    pl.PartsListColumns

                    For Each name As String In candidates

                        If String.Equals(col.Title, name, StringComparison.OrdinalIgnoreCase) Then

                            Return col.Title

                        End If

                    Next

                Next

            Catch
            End Try

            Return ""

        End Function


        '=========================================================
        ' GET CELL
        '=========================================================
        Private Function GetCellValue(row As Inventor.PartsListRow, colName As String) As String
            If colName = "" Then
                Return ""

            End If
            Try
                Dim v As Object = row.Item(colName).Value
                If v Is Nothing Then
                    Return ""
                End If
                Return CStr(v).Trim()
            Catch
                Return ""
            End Try

        End Function

        '=========================================================
        ' SET CELL
        '
        ' VALUE trước
        ' STATIC sau
        '=========================================================
        Private Sub SetCell(row As Inventor.PartsListRow, colName As String, value As String)

            If colName = "" Then
                Exit Sub
            End If

            If value Is Nothing Then
                Exit Sub
            End If

            Dim newValue As String = value.Trim()

            If newValue = "" Then
                Exit Sub
            End If

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

                '=================================================
                ' VALUE TRƯỚC
                '=================================================
                cell.Value = newValue

                '=================================================
                ' STATIC SAU
                '=================================================
                Try
                    cell.Static = True
                Catch
                End Try

            Catch
            End Try

        End Sub


        '=========================================================
        ' CLEAR CELL
        '
        ' UNIT QTY <= 1
        '=========================================================
        Private Sub ClearCell(row As Inventor.PartsListRow, colName As String)
            If colName = "" Then
                Exit Sub
            End If

            Try

                Dim cell As Inventor.PartsListCell = row.Item(colName)

                '=================================================
                ' XÓA VALUE
                '=================================================
                cell.Value = ""
                '=================================================
                ' STATIC
                '=================================================
                Try

                    cell.Static = True

                Catch
                End Try

            Catch

                Try
                    row.Item(colName).Value = ""
                Catch
                End Try
            End Try

        End Sub


        '=========================================================
        ' PARSE NUMBER
        '=========================================================
        Private Function TryParseNumber(value As String, ByRef result As Double) As Boolean
            result = 0

            If value Is Nothing Then
                Return False
            End If

            Dim s As String = value.Trim()

            If s = "" Then
                Return False
            End If


            '=====================================================
            ' CURRENT CULTURE
            '=====================================================
            Try
                If Double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, result) Then

                    Return True

                End If

            Catch
            End Try

            '=====================================================
            ' INVARIANT
            '=====================================================
            Try

                If Double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, result) Then
                    Return True

                End If

            Catch
            End Try


            '=====================================================
            ' ĐỔI , -> .
            '=====================================================
            Try

                Dim s2 As String = s.Replace(",", ".")

                If Double.TryParse(s2, NumberStyles.Any, CultureInfo.InvariantCulture, result) Then

                    Return True
                End If
            Catch
            End Try

            Return False

        End Function


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
            frm.Width = 500
            frm.Height = 320
            frm.ShowInTaskbar = False

            Dim lst As New ListBox()

            lst.Left = 12
            lst.Top = 12
            lst.Width = 460
            lst.Height = 220

            For Each s As String In items

                lst.Items.Add(s)

            Next

            If defaultIndex >= 0 AndAlso
               defaultIndex < lst.Items.Count Then

                lst.SelectedIndex = defaultIndex

            ElseIf lst.Items.Count > 0 Then

                lst.SelectedIndex = 0

            End If


            Dim btnOK As New Button() With {
                .Text = "OK",
                .Left = 300,
                .Top = 245,
                .Width = 80,
                .DialogResult = DialogResult.OK
            }


            Dim btnCancel As New Button() With {
                .Text = "Hủy",
                .Left = 390,
                .Top = 245,
                .Width = 80,
                .DialogResult = DialogResult.Cancel
            }


            frm.Controls.Add(lst)
            frm.Controls.Add(btnOK)
            frm.Controls.Add(btnCancel)

            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            If frm.ShowDialog() <> DialogResult.OK OrElse
               lst.SelectedIndex < 0 Then

                Return -1

            End If

            Return lst.SelectedIndex

        End Function


        '=========================================================
        ' GET PROPERTY
        '=========================================================
        Private Function GetProp(doc As Inventor.Document, propName As String) As String

            If doc Is Nothing Then

                Return ""

            End If

            Try

                Dim ps As Inventor.PropertySet = doc.PropertySets.Item("Design Tracking Properties")

                Dim v As Object = ps.Item(propName).Value


                If v Is Nothing Then

                    Return ""

                End If


                Return CStr(v).Trim()


            Catch

                Return ""

            End Try

        End Function


        '=========================================================
        ' SET PROPERTY
        '
        ' CHỈ GHI KHI PROPERTY ĐANG TRỐNG
        '=========================================================
        Private Sub SetProp(doc As Inventor.Document, propName As String, value As String)

            If doc Is Nothing Then
                Exit Sub
            End If


            If value Is Nothing Then
                Exit Sub
            End If

            Dim newValue As String = value.Trim()


            If newValue = "" Then
                Exit Sub
            End If


            Try

                Dim ps As Inventor.PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim prop As Inventor.Property = ps.Item(propName)

                Dim oldValue As String = ""


                Try

                    If prop.Value IsNot Nothing Then

                        oldValue = CStr(prop.Value).Trim()

                    End If

                Catch
                End Try


                '=================================================
                ' ĐÃ CÓ DỮ LIỆU -> KHÔNG GHI ĐÈ
                '=================================================
                If oldValue <> "" Then

                    Exit Sub

                End If


                '=================================================
                ' PROPERTY TRỐNG -> GHI
                '=================================================
                prop.Value = newValue

            Catch
            End Try

        End Sub


        '=========================================================
        ' ĐÁNH STT
        '=========================================================
        Private Function NumberRows(oPartList As Inventor.PartsList, cSTT As String, stt As Integer, preferAsm As Boolean, purchased As Boolean) As Integer


            For i As Integer = 1 To oPartList.PartsListRows.Count

                Try

                    Dim row As Inventor.PartsListRow = oPartList.PartsListRows.Item(i)
                    ' Bỏ qua nếu đã có STT rồi (giảm việc ghi đè)
                    Dim curSTT As String = GetCellValue(row, cSTT)  '''''''''''''''''''''''''''''''''''''''toi ưu lần 1 bỏ qua nếu stt đã có chữ
                    If curSTT <> "" Then Continue For              '''''''''''''''''''''''''''''''''''''''toi ưu lần 1 bỏ qua nếu stt đã có chữ

                    If row.ReferencedRows Is Nothing OrElse
                       row.ReferencedRows.Count < 1 Then

                        Continue For

                    End If


                    Dim bomRow As Inventor.BOMRow = row.ReferencedRows.Item(1).BOMRow


                    If bomRow Is Nothing Then
                        Continue For
                    End If


                    Dim isPurchased As Boolean =
                        (
                            bomRow.BOMStructure = Inventor.BOMStructureEnum.kPurchasedBOMStructure
                        )


                    If purchased <> isPurchased Then
                        Continue For

                    End If

                    If bomRow.ComponentDefinitions.Count < 1 Then
                        Continue For
                    End If

                    Dim d As Inventor.Document = bomRow.ComponentDefinitions.Item(1).Document

                    If d Is Nothing Then
                        Continue For
                    End If


                    Dim isAsm As Boolean =
                        (
                            d.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject
                        )


                    If purchased Then

                        SetCell(row, cSTT, stt.ToString())
                        stt += 1

                    ElseIf preferAsm AndAlso
                           isAsm Then

                        SetCell(row, cSTT, stt.ToString())

                        stt += 1

                    ElseIf (Not preferAsm) AndAlso
                           (Not isAsm) Then
                        SetCell(row, cSTT, stt.ToString())

                        stt += 1

                    End If

                Catch
                End Try

            Next

            Return stt

        End Function

        '=========================================================
        ' TÌM CỘT UNIT QTY THEO PROPERTY GỐC
        '
        ' KHÔNG DỰA VÀO TÊN HIỂN THỊ CỦA CỘT
        '
        ' Vì người dùng có thể đổi:
        '   UNIT QTY
        '   Số lượng
        '   SL
        '   Đơn vị
        '   ...
        '
        ' nên phải kiểm tra PropertyType của PartsListColumn.
        '=========================================================
        '    Private Function FindUnitQtyColumn(pl As Inventor.PartsList) As String

        '    Try

        '    For Each col As Inventor.PartsListColumn In
        '               pl.PartsListColumns

        '   Try
        '=================================================
        ' UNIT QUANTITY PROPERTY
        '=================================================
        '  If col.PropertyType = Inventor.PropertyTypeEnum.kUnitQuantityPartsListProperty Then

        '  Return col.Title

        '  End If
        ' Catch
        'End Try

        'Next

        'Catch
        'End Try

        '=========================================================
        ' TÌM CỘT UNIT QTY THEO PROPERTY GỐC
        ' CHỈ LẤY kUnitQuantityPartsListProperty
        ' KHÔNG FALLBACK sang Item Quantity
        '=========================================================
        Private Function FindUnitQtyColumn(pl As Inventor.PartsList) As String
            Try
                For Each col As Inventor.PartsListColumn In pl.PartsListColumns
                    Try
                        If col.PropertyType = Inventor.PropertyTypeEnum.kUnitQuantityPartsListProperty Then
                            Return col.Title
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            ' Không tìm thấy → trả về rỗng (không đụng Item Quantity)
            Return ""
        End Function

        '=========================================================
        ' FALLBACK:
        ' Một số cấu hình Inventor có thể trả về Item Quantity
        ' cho cột mà người dùng đang dùng làm UNIT QTY.
        '=========================================================
        '  Try

        '        For Each col As Inventor.PartsListColumn In
        '         pl.PartsListColumns

        '           Try

        'If col.PropertyType = Inventor.PropertyTypeEnum.kItemQuantityPartsListProperty Then

        'Return col.Title

        'End If

        ' Catch
        '          End Try

        '  Next
        'Catch
        'End Try

        'Return ""

        ' End Function

    End Module

End Namespace