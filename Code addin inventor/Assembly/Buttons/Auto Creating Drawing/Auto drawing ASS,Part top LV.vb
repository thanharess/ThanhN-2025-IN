Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports System.Collections
Imports System.Collections.Generic
Namespace ToolInventor2025.Assembly.Buttons.AutoCreateDrawing
    Public Module AutoDrawingASSpartTopLV

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                '=================================================
                ' 0. KIỂM TRA ASSEMBLY
                '=================================================
                If app.ActiveDocument Is Nothing OrElse
                       app.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Vui lòng mở Assembly (.iam) trước!", "Lỗi",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim asmDoc As Inventor.AssemblyDocument =
                        CType(app.ActiveDocument, Inventor.AssemblyDocument)

                Dim tg As Inventor.TransientGeometry = app.TransientGeometry

                '=================================================
                ' 1. CHẾ ĐỘ (bảng chọn)
                '=================================================
                Dim modeIdx As Integer = PickFromList("Chế độ", New String() {
                        "1 - Chỉ chi tiết (Part) cấp cao nhất",
                     "2 - Bản lắp tổng + Part + Sub-Assembly cấp cao"
                        }, 1)
                If modeIdx < 0 Then Exit Sub
                Dim mode As Integer = modeIdx + 1

                '=================================================
                ' 2. NGUỒN DRAWING (bảng chọn)
                '=================================================
                Dim srcIdx As Integer = PickFromList("Nguồn Drawing", New String() {
                        "1 - Tạo mới (template gốc Inventor)",
                        "2 - Tạo mới (chọn file template / .idw ngoài)",
                        "3 - Thêm vào drawing có sẵn"
                    }, 0)
                If srcIdx < 0 Then Exit Sub
                Dim srcMode As Integer = srcIdx + 1

                Dim drawDoc As Inventor.DrawingDocument = Nothing
                Dim isNewDrawing As Boolean = (srcMode = 1 OrElse srcMode = 2)

                If srcMode = 3 Then
                    ' Drawing có sẵn
                    Dim oFileDlg As Inventor.FileDialog = Nothing
                    app.CreateFileDialog(oFileDlg)
                    oFileDlg.Filter = "Bản vẽ Inventor (*.idw)|*.idw"
                    oFileDlg.DialogTitle = "Chọn file bản vẽ có sẵn"
                    oFileDlg.ShowOpen()
                    If String.IsNullOrEmpty(oFileDlg.FileName) Then
                        MessageBox.Show("Bạn chưa chọn file bản vẽ.", "Thông báo",
                                            MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Sub
                    End If
                    Try
                        drawDoc = CType(app.Documents.Open(oFileDlg.FileName, True), Inventor.DrawingDocument)
                    Catch ex As Exception
                        MessageBox.Show("Không mở được bản vẽ:" & vbCrLf & ex.Message, "Lỗi",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try

                ElseIf srcMode = 2 Then
                    ' Template / idw ngoài
                    Dim oFileDlg As Inventor.FileDialog = Nothing
                    app.CreateFileDialog(oFileDlg)
                    oFileDlg.Filter = "Inventor Drawing / Template (*.idw;*.idwt)|*.idw;*.idwt"
                    oFileDlg.DialogTitle = "Chọn template hoặc drawing làm mẫu"
                    oFileDlg.ShowOpen()
                    If String.IsNullOrEmpty(oFileDlg.FileName) Then Exit Sub

                    Dim sfd As New SaveFileDialog()
                    sfd.Filter = "Inventor Drawing (*.idw)|*.idw"
                    sfd.Title = "Lưu bản vẽ mới"
                    sfd.FileName = asmDoc.DisplayName & "_Drawing.idw"
                    If sfd.ShowDialog() <> DialogResult.OK Then Exit Sub

                    Try
                        Dim tmpDoc As Inventor.DrawingDocument =
                                CType(app.Documents.Open(oFileDlg.FileName, False), Inventor.DrawingDocument)
                        tmpDoc.SaveAs(sfd.FileName, True)
                        tmpDoc.Close(True)
                        drawDoc = CType(app.Documents.Open(sfd.FileName, True), Inventor.DrawingDocument)
                    Catch ex As Exception
                        MessageBox.Show("Không tạo được drawing từ template ngoài:" & vbCrLf & ex.Message,
                                            "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try

                Else
                    ' Template gốc Inventor
                    Dim sfd As New SaveFileDialog()
                    sfd.Filter = "Inventor Drawing (*.idw)|*.idw"
                    sfd.Title = "Lưu bản vẽ mới"
                    sfd.FileName = asmDoc.DisplayName & "_Drawing.idw"
                    If sfd.ShowDialog() <> DialogResult.OK Then Exit Sub

                    Try
                        Dim template As String =
                                app.FileManager.GetTemplateFile(Inventor.DocumentTypeEnum.kDrawingDocumentObject)
                        drawDoc = CType(app.Documents.Add(
                                Inventor.DocumentTypeEnum.kDrawingDocumentObject, template, True), Inventor.DrawingDocument)
                        drawDoc.SaveAs(sfd.FileName, False)
                    Catch ex As Exception
                        MessageBox.Show("Không tạo được drawing mới:" & vbCrLf & ex.Message,
                                            "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try
                End If

                '=================================================
                ' 3. KHỔ GIẤY (bảng chọn)
                '=================================================
                Dim sizeIdx As Integer = PickFromList("Khổ giấy", New String() {
                        "A0", "A1", "A2", "A3", "A4"
                    }, 3)
                If sizeIdx < 0 Then Exit Sub

                Dim sheetSizeEnum As Inventor.DrawingSheetSizeEnum
                Select Case sizeIdx
                    Case 0 : sheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA0DrawingSheetSize
                    Case 1 : sheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA1DrawingSheetSize
                    Case 2 : sheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA2DrawingSheetSize
                    Case 4 : sheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize
                    Case Else : sheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA3DrawingSheetSize
                End Select

                '=================================================
                ' 4. TỈ LỆ + SỐ PART (nhập)
                '=================================================
                Dim scaleInput As String = InputBox("Tỉ lệ view (20 = 1:20):", "Tỉ lệ", "5")
                If scaleInput = "" Then Exit Sub
                Dim userScale As Double = 1.0 / 10.0
                Try : userScale = 1.0 / CDbl(scaleInput) : Catch : End Try

                Dim partsPerSheetInput As String = InputBox("Số Part trên mỗi sheet:", "Số Part / Sheet", "4")
                If partsPerSheetInput = "" Then Exit Sub
                Dim partsPerSheet As Integer = 4
                Try : partsPerSheet = CInt(partsPerSheetInput) : Catch : End Try
                If partsPerSheet < 1 Then partsPerSheet = 1

                '=================================================
                ' 5. KIỂU VIEW (bảng chọn)
                '=================================================
                Dim viewIdx As Integer = PickFromList("Kiểu view", New String() {
                    "1 - Front",
                 "2 - Front + Right",
                         "3 - Front + Right + Iso",
                      "4 - Front + Top + Right + Iso",
                         "5 - Front + Left",
                             "6 - Front + Left + Right",
                          "7 - Front + Top + Left + Right",
                      "8 - Front + Top + Left + Right + Iso"}, 2)
                If viewIdx < 0 Then Exit Sub
                Dim viewType As Integer = viewIdx + 1

                '=================================================
                ' 6. BỘ LỌC
                '=================================================
                Dim filterSkip As Boolean =
                        (MessageBox.Show("Bỏ qua Purchased / Phantom / Reference?", "Bộ lọc",
                         MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes)

                '=================================================
                ' 7. MODE 2 → SHEET BẢN LẮP
                '=================================================
                If mode = 2 Then
                    Try
                        Dim asmSheet As Inventor.Sheet = drawDoc.Sheets.Add(sheetSizeEnum)
                        asmSheet.Size = sheetSizeEnum
                        asmSheet.Name = "Bản lắp " & asmDoc.DisplayName.Replace(".", "_")
                        ApplyBorderAndTitleBlock(drawDoc, asmSheet, sheetSizeEnum)

                        Dim cx As Double = asmSheet.Width / 2.0
                        Dim cy As Double = asmSheet.Height / 2.0
                        Dim asmViewPt As Inventor.Point2d = tg.CreatePoint2d(cx, cy)

                        Dim asmView As Inventor.DrawingView = asmSheet.DrawingViews.AddBaseView(
            asmDoc, asmViewPt, userScale,
            Inventor.ViewOrientationTypeEnum.kFrontViewOrientation,
            Inventor.DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
                        asmView.ShowLabel = True

                        AddProjectedViews(asmSheet, asmView, viewType, cx, cy, asmSheet.Width, asmSheet.Height, tg)

                        ' ★ BOM / PARTS LIST CHO CỤM
                        AddPartsListSafe(drawDoc, asmSheet, asmView, asmDoc, tg, "BẢNG KÊ CỤM TỔNG")

                    Catch ex As Exception
                        MessageBox.Show("Lỗi sheet bản lắp:" & vbCrLf & ex.Message, "Cảnh báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Try
                End If
                '=================================================
                ' 8. THU THẬP TOP LEVEL
                '    Mode 1: chỉ Part
                '    Mode 2: Part + Sub-Assembly (cấp cao nhất)
                '=================================================
                Dim allDocs As New List(Of Inventor.Document)
                Dim usedPN As New Hashtable()

                For Each occ As Inventor.ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                    Try
                        Dim refDoc As Inventor.Document = occ.Definition.Document

                        ' Mode 1: chỉ Part
                        ' Mode 2: Part + Assembly (sub)
                        If mode = 1 Then
                            If refDoc.DocumentType <> Inventor.DocumentTypeEnum.kPartDocumentObject Then
                                Continue For
                            End If
                        Else
                            ' Mode 2: chấp nhận Part và Sub-Assembly
                            If refDoc.DocumentType <> Inventor.DocumentTypeEnum.kPartDocumentObject AndAlso
               refDoc.DocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                                Continue For
                            End If
                        End If

                        ' Lọc Purchased / Phantom / Reference
                        If filterSkip Then
                            Try
                                Dim bs As Inventor.BOMStructureEnum = occ.BOMStructure
                                If bs = Inventor.BOMStructureEnum.kPurchasedBOMStructure OrElse
                   bs = Inventor.BOMStructureEnum.kReferenceBOMStructure OrElse
                   bs = Inventor.BOMStructureEnum.kPhantomBOMStructure Then
                                    Continue For
                                End If
                            Catch
                            End Try
                        End If

                        ' Lọc trùng Part Number
                        If AlreadyUsed(usedPN, refDoc) Then Continue For

                        allDocs.Add(refDoc)
                    Catch
                    End Try
                Next
                '=================================================
                ' 9. SHEET CHI TIẾT
                '=================================================
                '=================================================
                ' 9. SHEET CHI TIẾT / CỤM CON
                '=================================================
                Dim sheet As Inventor.Sheet = Nothing
                Dim sheetWidth, sheetHeight, usableW, usableH, xStart, yStart, xStep, yStep As Double
                Dim cols As Integer = If(partsPerSheet = 1, 1, 2)
                Dim partCountOnCurrentSheet As Integer = 0
                Dim globalCount As Integer = 0

                For Each refDoc As Inventor.Document In allDocs
                    Try
                        If sheet Is Nothing OrElse partCountOnCurrentSheet >= partsPerSheet Then
                            sheet = drawDoc.Sheets.Add(sheetSizeEnum)
                            sheet.Size = sheetSizeEnum

                            ' Đặt tên sheet theo loại
                            If refDoc.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                                sheet.Name = "Cụm " & drawDoc.Sheets.Count.ToString()
                            Else
                                sheet.Name = "Chi tiết " & drawDoc.Sheets.Count.ToString()
                            End If

                            ApplyBorderAndTitleBlock(drawDoc, sheet, sheetSizeEnum)

                            sheetWidth = sheet.Width
                            sheetHeight = sheet.Height
                            usableW = sheetWidth * 0.75
                            usableH = sheetHeight * 0.75
                            xStart = sheetWidth / 4.0
                            yStart = sheetHeight * 0.8
                            cols = If(partsPerSheet = 1, 1, 2)
                            Dim rowsNeeded As Integer = CInt(Math.Ceiling(partsPerSheet / CDbl(cols)))
                            xStep = usableW / cols
                            yStep = usableH / (rowsNeeded + 1)
                            partCountOnCurrentSheet = 0
                        End If

                        Dim colIndex As Integer = partCountOnCurrentSheet Mod cols
                        Dim rowIndex As Integer = partCountOnCurrentSheet \ cols
                        Dim xPos As Double = xStart + colIndex * xStep
                        Dim yPos As Double = yStart - rowIndex * yStep

                        Dim baseView As Inventor.DrawingView = sheet.DrawingViews.AddBaseView(
            refDoc, tg.CreatePoint2d(xPos, yPos), userScale,
            Inventor.ViewOrientationTypeEnum.kFrontViewOrientation,
            Inventor.DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
                        baseView.ShowLabel = True

                        baseView.ShowLabel = True

                        ' Dùng cùng logic với sheet lắp
                        AddProjectedViews(sheet, baseView, viewType, xPos, yPos, xStep * 2.5, yStep * 2.5, tg)

                        ' ★ BOM cho CỤM CON (sub-assembly)
                        If refDoc.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                            Try
                                Dim subAsm As Inventor.AssemblyDocument =
            CType(refDoc, Inventor.AssemblyDocument)

                                ' Đặt BOM gần view cụm (không đè view)
                                Dim bomPt As Inventor.Point2d =
            tg.CreatePoint2d(xPos + xStep * 0.15, yPos - yStep * 0.35)

                                AddPartsListSafeAt(drawDoc, sheet, baseView, subAsm, bomPt,
                           "BẢNG KÊ: " & subAsm.DisplayName)
                            Catch
                            End Try
                        End If

                        partCountOnCurrentSheet += 1
                        globalCount += 1
                    Catch
                    End Try
                Next

                '=================================================
                ' 10. XÓA SHEET TRẮNG ĐẦU (khi tạo drawing mới)
                '=================================================
                If isNewDrawing Then
                    DeleteBlankFirstSheets(drawDoc)
                End If

                '=================================================
                ' 11. HOÀN TẤT
                '=================================================
                drawDoc.Update()

                Dim modeText As String = If(mode = 1, "Chỉ Part", "Bản lắp + Part + Sub-Assembly")
                MessageBox.Show("Hoàn tất!" & vbCrLf & "Chế độ: " & modeText & vbCrLf & "Số mục đã vẽ: " & globalCount.ToString(), "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Lay Part Top Level",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        '=================================================
        ' XÓA SHEET TRẮNG (không có view)
        '=================================================
        Private Sub DeleteBlankFirstSheets(drawDoc As Inventor.DrawingDocument)
            Try
                ' Xóa từ cuối → đầu các sheet không có DrawingView
                For i As Integer = drawDoc.Sheets.Count To 1 Step -1
                    Try
                        Dim sh As Inventor.Sheet = drawDoc.Sheets.Item(i)
                        If sh.DrawingViews.Count = 0 Then
                            ' Không xóa nếu chỉ còn 1 sheet (Inventor cần ít nhất 1 sheet)
                            If drawDoc.Sheets.Count <= 1 Then Exit For
                            sh.Delete()
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try
        End Sub

        '=================================================
        ' BẢNG CHỌN
        '=================================================
        Private Function PickFromList(title As String, items As String(), Optional defaultIndex As Integer = 0) As Integer
            Dim frm As New Form()
            frm.Text = title
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.Width = 440
            frm.Height = 340
            frm.ShowInTaskbar = False

            Dim lst As New ListBox()
            lst.Left = 12
            lst.Top = 12
            lst.Width = 400
            lst.Height = 240
            For Each s As String In items
                lst.Items.Add(s)
            Next
            If defaultIndex >= 0 AndAlso defaultIndex < lst.Items.Count Then
                lst.SelectedIndex = defaultIndex
            ElseIf lst.Items.Count > 0 Then
                lst.SelectedIndex = 0
            End If

            Dim btnOK As New Button()
            btnOK.Text = "OK"
            btnOK.Left = 240
            btnOK.Top = 265
            btnOK.Width = 80
            btnOK.DialogResult = DialogResult.OK

            Dim btnCancel As New Button()
            btnCancel.Text = "Hủy"
            btnCancel.Left = 330
            btnCancel.Top = 265
            btnCancel.Width = 80
            btnCancel.DialogResult = DialogResult.Cancel

            frm.Controls.Add(lst)
            frm.Controls.Add(btnOK)
            frm.Controls.Add(btnCancel)
            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            If frm.ShowDialog() <> DialogResult.OK OrElse lst.SelectedIndex < 0 Then
                Return -1
            End If
            Return lst.SelectedIndex
        End Function

        '=================================================
        ' BORDER + KHUNG TÊN
        '=================================================
        Private Sub ApplyBorderAndTitleBlock(
                drawDoc As Inventor.DrawingDocument,
                sheet As Inventor.Sheet,
                sizeEnum As Inventor.DrawingSheetSizeEnum)

            Try
                Dim borderName As String = ""
                Dim titleName As String = ""

                Select Case sizeEnum
                    Case Inventor.DrawingSheetSizeEnum.kA0DrawingSheetSize
                        borderName = "NT A0" : titleName = "Khung tên SX NT A0"
                    Case Inventor.DrawingSheetSizeEnum.kA1DrawingSheetSize
                        borderName = "NT A1" : titleName = "Khung tên SX NT A1"
                    Case Inventor.DrawingSheetSizeEnum.kA2DrawingSheetSize
                        borderName = "NT A2" : titleName = "Khung tên SX NT A2"
                    Case Inventor.DrawingSheetSizeEnum.kA3DrawingSheetSize
                        borderName = "NT A3" : titleName = "Khung tên SX NT A3"
                    Case Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize
                        borderName = "NT A4" : titleName = "Khung tên SX NT A4"
                    Case Else
                        borderName = "NT A3" : titleName = "Khung tên SX NT A3"
                End Select

                Try
                    If sheet.Border IsNot Nothing Then sheet.Border.Delete()
                Catch
                End Try
                Try
                    If sheet.TitleBlock IsNot Nothing Then sheet.TitleBlock.Delete()
                Catch
                End Try

                Try
                    sheet.AddBorder(drawDoc.BorderDefinitions.Item(borderName))
                Catch
                    If sizeEnum = Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize Then
                        Try : sheet.AddBorder(drawDoc.BorderDefinitions.Item("NT A4 D")) : Catch : End Try
                    End If
                End Try

                Try
                    sheet.AddTitleBlock(drawDoc.TitleBlockDefinitions.Item(titleName))
                Catch
                    If sizeEnum = Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize Then
                        Try : sheet.AddTitleBlock(drawDoc.TitleBlockDefinitions.Item("Khung tên SX NT A4 D")) : Catch : End Try
                    End If
                End Try
            Catch
            End Try
        End Sub

        '=================================================
        ' LỌC TRÙNG PART NUMBER
        '=================================================
        Private Function GetPartNumber(doc As Inventor.Document) As String
            Try
                Dim ps As Inventor.PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim pn As String = CStr(ps.Item("Part Number").Value).Trim().ToUpper()
                If pn <> "" Then Return pn
            Catch
            End Try
            Try
                Return System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName).ToUpper()
            Catch
            End Try
            Try : Return doc.DisplayName.ToUpper() : Catch : End Try
            Return ""
        End Function

        Private Function AlreadyUsed(used As Hashtable, doc As Inventor.Document) As Boolean
            Dim key As String = GetPartNumber(doc)
            If key = "" Then
                Try : key = doc.InternalName : Catch : Return False : End Try
            End If
            If used.ContainsKey(key) Then Return True
            used.Add(key, True)
            Return False
        End Function

        '=================================================
        ' PROJECTED VIEWS
        '=================================================
        Private Sub AddProjectedViews(
    sheet As Inventor.Sheet,
    baseView As Inventor.DrawingView,
    viewType As Integer,
    cx As Double, cy As Double,
    sheetW As Double, sheetH As Double,
    tg As Inventor.TransientGeometry)

            Dim dx As Double = sheetW / 3.0
            Dim dy As Double = sheetH / 3.0

            ' Right
            If viewType = 2 OrElse viewType = 3 OrElse viewType = 4 OrElse
       viewType = 6 OrElse viewType = 7 OrElse viewType = 8 Then
                Dim rightPt As Inventor.Point2d = tg.CreatePoint2d(cx + dx, cy)
                Dim rv As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
            baseView, rightPt,
            Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
                rv.ShowLabel = True
            End If

            ' Left
            If viewType = 5 OrElse viewType = 6 OrElse viewType = 7 OrElse viewType = 8 Then
                Dim leftPt As Inventor.Point2d = tg.CreatePoint2d(cx - dx, cy)
                Dim lv As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
            baseView, leftPt,
            Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
                lv.ShowLabel = True
            End If

            ' Top (dưới Front)
            If viewType = 4 OrElse viewType = 7 OrElse viewType = 8 Then
                Dim topPt As Inventor.Point2d = tg.CreatePoint2d(cx, cy - dy)
                Dim tv As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
            baseView, topPt,
            Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
                tv.ShowLabel = True
            End If

            ' Iso
            If viewType = 3 OrElse viewType = 4 OrElse viewType = 8 Then
                Dim isoPt As Inventor.Point2d = tg.CreatePoint2d(cx + dx, cy - dy)
                sheet.DrawingViews.AddProjectedView(
            baseView, isoPt,
            Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
            End If
        End Sub
        Private Sub AddPartsListSafe(
    drawDoc As Inventor.DrawingDocument,
    sheet As Inventor.Sheet,
    baseView As Inventor.DrawingView,
    sourceAsm As Inventor.AssemblyDocument,
    tg As Inventor.TransientGeometry,
    title As String)

            Dim lastErr As String = ""

            Try
                ' 1. Bật BOM trên assembly
                Try
                    Dim bom As Inventor.BOM = sourceAsm.ComponentDefinition.BOM
                    bom.StructuredViewEnabled = True
                    bom.StructuredViewFirstLevelOnly = False
                    Try
                        bom.PartsOnlyViewEnabled = True
                    Catch
                    End Try
                Catch ex As Exception
                    lastErr = "Bật BOM: " & ex.Message
                End Try

                Try
                    sourceAsm.Update2(True)
                Catch
                End Try

                drawDoc.Update()
                System.Windows.Forms.Application.DoEvents()

                ' 2. Vị trí Parts List (góc trên phải)
                Dim pt As Inventor.Point2d =
            tg.CreatePoint2d(sheet.Width * 0.98, sheet.Height * 0.98)

                Dim pl As Inventor.PartsList = Nothing

                ' 3. Thử từng Level
                If pl Is Nothing Then
                    Try
                        pl = sheet.PartsLists.Add(baseView, pt,
                    Inventor.PartsListLevelEnum.kFirstLevelComponents)
                    Catch ex As Exception
                        lastErr = "FirstLevel: " & ex.Message
                    End Try
                End If

                If pl Is Nothing Then
                    Try
                        pl = sheet.PartsLists.Add(baseView, pt,
                    Inventor.PartsListLevelEnum.kStructuredAllLevels)
                    Catch ex As Exception
                        lastErr = "StructuredAll: " & ex.Message
                    End Try
                End If

                If pl Is Nothing Then
                    Try
                        pl = sheet.PartsLists.Add(baseView, pt,
                    Inventor.PartsListLevelEnum.kPartsOnly)
                    Catch ex As Exception
                        lastErr = "PartsOnly: " & ex.Message
                    End Try
                End If

                If pl Is Nothing Then
                    Try
                        pl = sheet.PartsLists.Add(baseView, pt)
                    Catch ex As Exception
                        lastErr = "Default: " & ex.Message
                    End Try
                End If

                ' 4. Thử Add bằng Document (không qua View)
                If pl Is Nothing Then
                    Try
                        pl = sheet.PartsLists.Add(sourceAsm, pt,
                    Inventor.PartsListLevelEnum.kFirstLevelComponents)
                    Catch ex As Exception
                        lastErr = "By Document: " & ex.Message
                    End Try
                End If

                If pl Is Nothing Then
                    MessageBox.Show(
                "Không tạo được Parts List!" & vbCrLf & vbCrLf &
                "Assembly: " & sourceAsm.DisplayName & vbCrLf &
                "Sheet: " & sheet.Name & vbCrLf & vbCrLf &
                "Lỗi: " & lastErr,
                "BOM", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                ' 5. Style
                Try
                    If drawDoc.StylesManager.PartsListStyles.Count > 0 Then
                        pl.Style = drawDoc.StylesManager.PartsListStyles.Item(1)
                    End If
                Catch
                End Try

                Try
                    pl.Title = title
                    pl.ShowTitle = True
                Catch
                End Try

                Try
                    pl.Renumber()
                Catch
                End Try

                drawDoc.Update()

            Catch ex As Exception
                MessageBox.Show("AddPartsListSafe: " & ex.Message, "BOM",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub
        Private Sub AddPartsListSafeAt(
    drawDoc As Inventor.DrawingDocument,
    sheet As Inventor.Sheet,
    baseView As Inventor.DrawingView,
    sourceAsm As Inventor.AssemblyDocument,
    pt As Inventor.Point2d,
    title As String)

            Try
                Try
                    Dim bom As Inventor.BOM = sourceAsm.ComponentDefinition.BOM
                    bom.StructuredViewEnabled = True
                    bom.StructuredViewFirstLevelOnly = False
                    Try : bom.PartsOnlyViewEnabled = True : Catch : End Try
                Catch
                End Try

                Try : sourceAsm.Update2(True) : Catch : End Try
                drawDoc.Update()
                System.Windows.Forms.Application.DoEvents()

                Dim pl As Inventor.PartsList = Nothing

                Try
                    pl = sheet.PartsLists.Add(baseView, pt,
                        Inventor.PartsListLevelEnum.kFirstLevelComponents)
                Catch
                    Try
                        pl = sheet.PartsLists.Add(baseView, pt,
                            Inventor.PartsListLevelEnum.kStructuredAllLevels)
                    Catch
                        Try
                            pl = sheet.PartsLists.Add(baseView, pt,
                                Inventor.PartsListLevelEnum.kPartsOnly)
                        Catch
                            Try : pl = sheet.PartsLists.Add(baseView, pt) : Catch : End Try
                        End Try
                    End Try
                End Try

                If pl Is Nothing Then Exit Sub

                Try
                    If drawDoc.StylesManager.PartsListStyles.Count > 0 Then
                        pl.Style = drawDoc.StylesManager.PartsListStyles.Item(1)
                    End If
                Catch
                End Try

                Try
                    pl.Title = title
                    pl.ShowTitle = True
                Catch
                End Try

                Try : pl.Renumber() : Catch : End Try

            Catch
            End Try
        End Sub
    End Module

End Namespace
