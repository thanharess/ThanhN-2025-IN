
Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Windows.Forms
Imports Inventor
Imports ToolInventor2025.ToolInventor2025.Assembly.Buttons

Namespace ToolInventor2025.Drawing.Buttons.DrawPartList
    Public Module Draw_Partlist_1b
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
                Dim nameModeIdx As Integer = PickFromList("Xử lý cột Tên Theo ENG", New String() {
    "1 - Part Number: không ghi đè nếu Tên đã là PN/SN",
    "2 - Part Number: chỉ ghi khi ô Tên đang trống",
    "3 - Stock Number: đồng bộ trực tiếp vào BOM",
    "4 - Không sửa tên (dùng Part Number để đoán đơn vị)",
    "5 - Stock Number: không ghi đè nếu Tên đã là PN/SN",
    "6 - Stock Number: chỉ ghi khi ô Tên đang trống",
    "7 - Không sửa tên (dùng Stock Number để đoán đơn vị)"}, 0)

                If nameModeIdx < 0 Then
                    Exit Sub
                End If

                Dim nameMode As Integer = nameModeIdx + 1

                '=================================================
                ' 2. VẬT LIỆU
                '=================================================
                Dim matDefault As String =
                    InputBox("Vật liệu mặc định (Part tự chế):", "Material", "SS400")

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
                Dim colSTT As String = "No"
                Dim colTen As String = "Name"
                Dim colTen2 As String = "Product name"
                Dim colDonVi As String = "Unit"
                Dim colVL As String = "Material"
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
             If(nameMode = 4 OrElse nameMode = 7, "Không sửa",
             If(nameMode = 5 OrElse nameMode = 6, "Stock Number", "Part Number"))) &
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
        '
        ' MODE 7:
        '   - CHỈ ĐỌC Ô TÊN ĐANG HIỂN THỊ TRÊN PARTS LIST
        '   - KHÔNG ĐỌC PART NUMBER
        '   - KHÔNG ĐỌC STOCK NUMBER PROPERTY
        '   - KHÔNG TRUY CẬP DOCUMENT
        '   - KHÔNG TRUY CẬP BOM
        '
        ' MODE 1-6:
        '   Giữ nguyên logic cũ.
        '=========================================================
        Private Sub ProcessOnePartsList(
    oPartList As Inventor.PartsList,
    nameMode As Integer,
    matDefault As String,
    clearPurchasedMaterial As Boolean,
    colSTT As String,
    colTen As String,
    colTen2 As String,
    colDonVi As String,
    colVL As String,
    colUnitQty As String)

            '=========================================================
            ' TÌM CỘT
            '=========================================================
            Dim cSTT As String =
        FindColumn(oPartList, New String() {
            colSTT,
            "Item", "Items", "No", "No.",
            "ITEM", "ITEMS", "For items", "FOR ITEMS"
        })

            Dim cTen As String =
        FindColumn(oPartList, New String() {
            colTen,
            colTen2,
            "Part Number",
            "Stock Number",
            "Name", "Tên", "name", "Product name", "PRODUCT NAME", "tên"
        })

            Dim nameCol As String = cTen

            Dim cDonVi As String =
        FindColumn(oPartList, New String() {
            colDonVi,
            "Keywords", "KEYWORDS",
            "Unit", "unit",
            "ĐƠN VỊ"
        })

            Dim cVL As String =
        FindColumn(oPartList, New String() {
            colVL,
            "Material", "material",
            "MATERIAL",
            "VẬT LIỆU"
        })

            Dim cUnitQty As String = FindUnitQtyColumn(oPartList)


            '=========================================================
            ' KIỂM TRA CỘT
            '=========================================================
            If cSTT = "" OrElse cDonVi = "" Then

                Throw New Exception(
            "Không tìm thấy cột STT hoặc Đơn vị trên Parts List.")

            End If

            '=========================================================
            ' MODE 7 - SIÊU TỐI ƯU
            '
            ' CHỈ ĐỌC THÔNG TIN ĐANG HIỂN THỊ TRÊN PARTS LIST
            '
            ' Không:
            '   - BOM
            '   - BOMRow
            '   - ComponentDefinitions
            '   - Document
            '   - Part Number
            '   - Stock Number Property
            '   - NumberRows
            '   - Sort
            '   - SaveItemOverridesToBOM
            '   - Unit Qty
            '
            ' Chỉ:
            '   Parts List → Tên → GuessUnit → Đơn vị
            '=========================================================
            If nameMode = 7 Then

                For i As Integer = 1 To oPartList.PartsListRows.Count

                    Try

                        Dim row As Inventor.PartsListRow =
                oPartList.PartsListRows.Item(i)

                        '=================================================
                        ' ĐỌC TRỰC TIẾP Ô TÊN
                        '=================================================
                        Dim displayName As String = ""

                        Try
                            Dim v As Object = row.Item(cTen).Value

                            If v IsNot Nothing Then
                                displayName = CStr(v).Trim()
                            End If

                        Catch
                            displayName = ""
                        End Try


                        '=================================================
                        ' ĐOÁN ĐƠN VỊ
                        '=================================================
                        Dim donVi As String =
                GuessUnit(displayName)


                        '=================================================
                        ' CHỈ GHI KHI KHÁC GIÁ TRỊ HIỆN TẠI
                        '=================================================
                        Dim currentUnit As String = ""

                        Try
                            Dim oldValue As Object =
                    row.Item(cDonVi).Value

                            If oldValue IsNot Nothing Then
                                currentUnit = CStr(oldValue).Trim()
                            End If

                        Catch
                            currentUnit = ""
                        End Try


                        If Not String.Equals(
                currentUnit,
                donVi,
                StringComparison.OrdinalIgnoreCase) Then

                            Try
                                row.Item(cDonVi).Value = donVi
                            Catch
                            End Try

                        End If

                    Catch
                        ' Bỏ qua row lỗi
                    End Try

                Next


                '=========================================================
                ' CHỈ UPDATE 1 LẦN
                '=========================================================
                Try
                    oPartList.Update()
                Catch
                End Try


                '=========================================================
                ' KẾT THÚC MODE 7
                '=========================================================
                Exit Sub

            End If

            '=========================================================
            ' MODE 1-6
            '
            ' Từ đây trở xuống giữ logic BOM/Document cũ.
            '=========================================================
            For i As Integer = 1 To oPartList.PartsListRows.Count

                Try

                    Dim row As Inventor.PartsListRow =
                oPartList.PartsListRows.Item(i)


                    '=================================================
                    ' KIỂM TRA REFERENCED ROW
                    '=================================================
                    If row.ReferencedRows Is Nothing OrElse
               row.ReferencedRows.Count < 1 Then

                        Continue For

                    End If


                    Dim bomRow As Inventor.BOMRow =
                row.ReferencedRows.Item(1).BOMRow

                    If bomRow Is Nothing Then
                        Continue For
                    End If


                    '=================================================
                    ' BOM STRUCTURE
                    '=================================================
                    Dim bs As Inventor.BOMStructureEnum =
                bomRow.BOMStructure

                    Dim isPurchased As Boolean =
                (bs =
                 Inventor.BOMStructureEnum.kPurchasedBOMStructure)


                    '=================================================
                    ' DOCUMENT
                    '=================================================
                    Dim refDoc As Inventor.Document = Nothing

                    Dim docType As Inventor.DocumentTypeEnum =
                Inventor.DocumentTypeEnum.kUnknownDocumentObject


                    Try

                        If bomRow.ComponentDefinitions.Count > 0 Then

                            refDoc =
                        bomRow.ComponentDefinitions.Item(1).Document

                            If refDoc IsNot Nothing Then

                                docType =
                            refDoc.DocumentType

                            End If

                        End If

                    Catch
                    End Try


                    Dim isAsm As Boolean =
                (docType =
                 Inventor.DocumentTypeEnum.kAssemblyDocumentObject)

                    Dim isPart As Boolean =
                (docType =
                 Inventor.DocumentTypeEnum.kPartDocumentObject)

                    '=================================================
                    ' PN / SN / TÊN
                    '=================================================
                    Dim pn As String = ""
                    Dim sn As String = ""
                    Dim currentName As String = ""

                    ' Mode 1,2,4,5,6 đều cần đọc thông tin để đoán đơn vị
                    ' Chỉ Mode 1,2,5,6 mới sửa cột Tên
                    If nameMode = 1 OrElse nameMode = 2 OrElse nameMode = 4 OrElse
   nameMode = 5 OrElse nameMode = 6 Then

                        If refDoc IsNot Nothing Then
                            pn = GetProp(refDoc, "Part Number")
                            sn = GetProp(refDoc, "Stock Number")
                        End If

                        pn = If(pn, "").Trim()
                        sn = If(sn, "").Trim()

                        ' Đọc tên hiện tại trên Parts List
                        currentName = GetCellValue(row, cTen)

                        ' Chỉ Mode 1,2,5,6 mới ghi đè tên
                        If nameMode = 1 OrElse nameMode = 2 OrElse nameMode = 5 OrElse nameMode = 6 Then
                            ApplyNameLogic(row, currentName, pn, sn, nameMode, nameCol)
                        End If

                    End If

                    '=================================================
                    ' ĐƠN VỊ
                    '=================================================
                    If isAsm Then

                        SetCell(
                    row,
                    cDonVi,
                    "Set")


                    ElseIf isPart AndAlso Not isPurchased Then

                        Dim nameForUnit As String = ""

                        If nameMode = 5 OrElse nameMode = 6 Then
                            ' Ưu tiên Stock Number
                            nameForUnit = If(sn <> "", sn, pn)
                        Else
                            ' Mode 1, 2, 4: ưu tiên tên đang hiển thị
                            nameForUnit = currentName
                            If nameForUnit = "" Then
                                nameForUnit = If(pn <> "", pn, sn)
                            End If
                        End If

                        Dim donVi As String = GuessUnit(nameForUnit)
                        SetCell(row, cDonVi, donVi)

                    ElseIf isPurchased Then
                        SetCell(row, cDonVi, "Pcs")
                    End If


                    '=================================================
                    ' MATERIAL
                    '
                    ' ASSEMBLY:
                    '   Không đụng
                    '
                    ' PURCHASED:
                    '   Yes → Xóa
                    '   No  → Giữ nguyên
                    '
                    ' PART TỰ CHẾ:
                    '   Ghi vật liệu mặc định
                    '=================================================
                    If isAsm Then

                        ' Không đụng

                    ElseIf isPurchased Then

                        If clearPurchasedMaterial Then

                            ClearCell(
                        row,
                        cVL)

                        End If

                    ElseIf isPart Then

                        If matDefault <> "" Then

                            SetCell(
                        row,
                        cVL,
                        matDefault)

                        End If

                    End If


                Catch
                    ' Bỏ qua row lỗi
                End Try

            Next


            '=========================================================
            ' UPDATE TRƯỚC KHI ĐÁNH STT
            '=========================================================
            Try
                oPartList.Update()
            Catch
            End Try


            '=========================================================
            ' STT
            '
            ' CỤM → PART → PURCHASED
            '=========================================================
            Dim stt As Integer = 1

            stt =
        NumberRows(
            oPartList,
            cSTT,
            stt,
            True,
            False)

            stt =
        NumberRows(
            oPartList,
            cSTT,
            stt,
            False,
            False)

            stt =
        NumberRows(
            oPartList,
            cSTT,
            stt,
            False,
            True)


            '=========================================================
            ' SORT
            '=========================================================
            Try

                oPartList.Sort(cSTT)

            Catch

                Try

                    oPartList.Sort("Item")

                Catch
                End Try

            End Try


            '=========================================================
            ' SAVE OVERRIDE XUỐNG BOM
            '=========================================================
            Try

                oPartList.SaveItemOverridesToBOM()

            Catch
            End Try


            '=========================================================
            ' UNIT QTY
            '
            ' <= 1 → XÓA
            ' > 1  → GIỮ
            '
            ' Chỉ xóa trên Parts List.
            ' Không Save lại xuống BOM.
            '=========================================================
            If cUnitQty <> "" Then

                For qtyIdx As Integer =
            1 To oPartList.PartsListRows.Count

                    Try

                        Dim qtyRow As Inventor.PartsListRow =
                    oPartList.PartsListRows.Item(qtyIdx)

                        Dim qtyText As String =
                    GetCellValue(
                        qtyRow,
                        cUnitQty)


                        If qtyText <> "" Then

                            Dim qty As Double = 0


                            If TryParseNumber(
                        qtyText,
                        qty) Then

                                If qty <= 1 Then

                                    ClearCell(
                                qtyRow,
                                cUnitQty)

                                End If

                            End If

                        End If


                    Catch
                    End Try

                Next

            End If


            '=========================================================
            ' UPDATE LẦN CUỐI
            '=========================================================
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
        Private Sub ApplyNameLogic(row As Inventor.PartsListRow, currentName As String, pn As String, sn As String, nameMode As Integer, nameCol As String)

            Dim cur As String = If(currentName, "").Trim()
            Dim requiredName As String = ""

            Select Case nameMode
                Case 1, 2
                    ' Ưu tiên Part Number
                    requiredName = If(pn <> "", pn, sn)

                Case 5, 6
                    ' Ưu tiên Stock Number
                    requiredName = If(sn <> "", sn, pn)
            End Select

            If requiredName = "" Then Exit Sub

            Select Case nameMode
                Case 1, 5   ' Không ghi đè nếu tên đã là PN hoặc SN
                    If cur <> "" AndAlso
               (String.Equals(cur, pn, StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(cur, sn, StringComparison.OrdinalIgnoreCase)) Then
                        Exit Sub
                    End If
                    SetCell(row, nameCol, requiredName)

                Case 2, 6   ' Chỉ ghi khi ô đang trống
                    If cur = "" Then
                        SetCell(row, nameCol, requiredName)
                    End If
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

                Dim value As String = CStr(row.Item("Name").Value)

                Return "Name"

            Catch
            End Try


            Try
                Dim value As String = CStr(row.Item("Product name").Value)
                Return "Product name"

            Catch
            End Try

            Try

                Dim value As String = CStr(row.Item("Tên").Value)

                Return "Tên"

            Catch
            End Try

            Return "Name"

        End Function


        '=========================================================
        ' ĐƠN VỊ
        '=========================================================




        ''''''''''''''''''code mới đang dùng
        '=========================================================
        ' ĐOÁN ĐƠN VỊ
        '
        ' Ưu tiên:
        '   1. Thanh
        '   2. Tấm
        '   3. Cái
        '
        ' Mode 7 sử dụng Stock Number đưa vào hàm này
        '=========================================================
        '=========================================================
        ' ĐOÁN ĐƠN VỊ
        '
        ' NGUYÊN TẮC:
        '
        ' 1. So sánh từ đầu chuỗi
        ' 2. So sánh từ trái sang phải
        ' 3. Không nhận diện bằng 1 chữ cái
        ' 4. Không kiểm tra chiều dài L hoặc mm
        ' 5. Không cần phân tích kích thước
        '
        ' Ưu tiên:
        '       THANH
        '       TẤM
        '       CÁI
        '
        ' Mode 7:
        '       Input = Stock Number
        '=========================================================
        '=========================================================
        ' ĐOÁN ĐƠN VỊ - GỘP 2 KIỂU CŨ THÀNH 1
        '
        ' Ưu tiên:
        '   1. Thanh
        '   2. Tấm
        '   3. Cái
        '
        ' Logic:
        '   - Ưu tiên keyword từ đầu chuỗi (GuessUnit)
        '   - Kết hợp thêm điều kiện EndsWith L / mm (GuessUnit2)
        '=========================================================
        Private Function GuessUnit(ten As String) As String

            If String.IsNullOrEmpty(ten) Then
                Return "Pcs"
            End If

            Dim t As String = ten.Trim()
            If t = "" Then
                Return "Pcs"
            End If

            '=====================================================
            ' 1. THANH
            '=====================================================
            Dim thanhKeywords As String() = {
        "THANH", "THÉP", "TH", "TR", "ỐNG", "ONG", "PIPE", "TUBE", "THEP", "TRỤC", "TRUC", "SHAFT", "PHI", "CÂY", "CAY", "RHS", "SHS",
        "UPE", "V6", "V5", "V1", "V7", "V8", "V9", "I1", "I2", "I4", "H1", "H2", "H3", "H4",
        "XG", "SH", "ỐN", "CÂ", "C1", "C2", "C5", "C6", "C7", "U1", "U2", "U8", "C8", "C9", "u9", "u7", "u6"
    }

            ' Kiểm tra keyword từ đầu chuỗi
            For Each keyword As String In thanhKeywords
                If String.IsNullOrEmpty(keyword) Then Continue For

                If t.StartsWith(keyword, StringComparison.OrdinalIgnoreCase) Then
                    Return "Bar"
                End If
            Next

            ' Kết hợp logic cũ: chữ cái đầu + EndsWith L hoặc mm
            If t.Length >= 2 Then
                Dim c0 As Char = Char.ToUpperInvariant(t(0))

                If "TPVLHIZCUS".IndexOf(c0) >= 0 AndAlso
           (t.EndsWith("L", StringComparison.OrdinalIgnoreCase) OrElse
            t.EndsWith("mm", StringComparison.OrdinalIgnoreCase)) Then

                    Return "Bar"
                End If
            End If

            '=====================================================
            ' 2. TẤM
            '=====================================================
            Dim tamKeywords As String() = {
        "PL", "TẤM", "TAM", "TÔN", "TON", "MÃ", "MA",
        "BÌ", "BI", "PLATE", "TÔ", "TẤ", "Mái"
    }

            For Each keyword As String In tamKeywords
                If String.IsNullOrEmpty(keyword) Then Continue For

                If t.StartsWith(keyword, StringComparison.OrdinalIgnoreCase) Then
                    Return "Pcs"
                End If
            Next

            '=====================================================
            ' 3. MẶC ĐỊNH
            '=====================================================
            Return "Pcs"

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