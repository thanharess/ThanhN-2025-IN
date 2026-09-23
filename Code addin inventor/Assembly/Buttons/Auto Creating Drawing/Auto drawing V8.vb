Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports System.Collections
Imports System.Collections.Generic
Namespace ToolInventor2025.Assembly.Buttons.AutoCreateDrawing
    Public Module AutoDrawingV8

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                '=================================================
                ' 0. KIỂM TRA ASSEMBLY
                '=================================================
                If app.ActiveDocument Is Nothing OrElse
               app.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Vui lòng mở file Assembly (.iam) trước!", "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim asmDoc As Inventor.AssemblyDocument =
                CType(app.ActiveDocument, Inventor.AssemblyDocument)

                Dim asmParams As Inventor.Parameters = Nothing
                Try
                    asmParams = asmDoc.ComponentDefinition.Parameters
                Catch
                    MessageBox.Show("Không thể đọc Parameters của assembly.", "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End Try

                '=================================================
                ' 1. TÊN THAM SỐ LƯU
                '=================================================
                Dim P_AsmScale As String = "iLogic_AsmScale"
                Dim P_PartScale As String = "iLogic_PartScale"
                Dim P_SheetSize As String = "iLogic_SheetSize"
                Dim P_ViewType As String = "iLogic_ViewType"
                Dim P_PartPerSheet As String = "iLogic_PartPerSheet"
                Dim P_FilterPurchased As String = "iLogic_FilterPurchased"
                Dim P_FilterPhantom As String = "iLogic_FilterPhantom"
                Dim P_BOMcreate As String = "iLogic_BOMcreate"

                '=================================================
                ' 2. ĐỌC GIÁ TRỊ ĐÃ LƯU
                '=================================================
                Dim asmScalePrev As Double = 1.0 / 20.0
                Dim partScalePrev As Double = 1.0 / 10.0
                Dim sheetSizePrev As Integer = 3
                Dim viewTypePrev As Integer = 3
                Dim partsPerSheetPrev As Integer = 4
                Dim filterPurchasedPrev As Integer = 0
                Dim filterPhantomPrev As Integer = 0
                Dim BOMcreatePrev As Integer = 1

                Try : asmScalePrev = asmParams.UserParameters.Item(P_AsmScale).Value : Catch : End Try
                Try : partScalePrev = asmParams.UserParameters.Item(P_PartScale).Value : Catch : End Try
                Try : sheetSizePrev = CInt(asmParams.UserParameters.Item(P_SheetSize).Value) : Catch : End Try
                Try : viewTypePrev = CInt(asmParams.UserParameters.Item(P_ViewType).Value) : Catch : End Try
                Try : partsPerSheetPrev = CInt(asmParams.UserParameters.Item(P_PartPerSheet).Value) : Catch : End Try
                Try : filterPurchasedPrev = CInt(asmParams.UserParameters.Item(P_FilterPurchased).Value) : Catch : End Try
                Try : filterPhantomPrev = CInt(asmParams.UserParameters.Item(P_FilterPhantom).Value) : Catch : End Try
                Try : BOMcreatePrev = CInt(asmParams.UserParameters.Item(P_BOMcreate).Value) : Catch : End Try

                EnsureParam(asmParams, P_AsmScale, asmScalePrev)
                EnsureParam(asmParams, P_PartScale, partScalePrev)
                EnsureParam(asmParams, P_SheetSize, sheetSizePrev)
                EnsureParam(asmParams, P_ViewType, viewTypePrev)
                EnsureParam(asmParams, P_PartPerSheet, partsPerSheetPrev)
                EnsureParam(asmParams, P_FilterPurchased, filterPurchasedPrev)
                EnsureParam(asmParams, P_FilterPhantom, filterPhantomPrev)
                EnsureParam(asmParams, P_BOMcreate, BOMcreatePrev)

                '=================================================
                ' 3. HỎI DÙNG CẤU HÌNH CŨ?
                '=================================================
                Dim useSaved As Boolean = False
                Dim haveSavedConfig As Boolean = False
                Try
                    Dim t = asmParams.UserParameters.Item(P_AsmScale)
                    haveSavedConfig = True
                Catch
                End Try

                If haveSavedConfig Then
                    If MessageBox.Show("Có cấu hình lưu trước. Dùng lại cấu hình cũ?",
                                   "Dùng cấu hình cũ?", MessageBoxButtons.YesNo,
                                   MessageBoxIcon.Question) = DialogResult.Yes Then
                        useSaved = True
                    End If
                End If

                '=================================================
                ' 4. ĐỌC / NHẬP CẤU HÌNH
                '=================================================
                Dim asmScale As Double = asmScalePrev
                Dim partScale As Double = partScalePrev
                Dim sheetSizeEnum As Inventor.DrawingSheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA3DrawingSheetSize
                Dim sheetSizeChoice As String = "A3"
                Dim viewType As Integer = viewTypePrev
                Dim partsPerSheet As Integer = partsPerSheetPrev
                Dim filterPurchased As Boolean = (filterPurchasedPrev = 1)
                Dim filterPhantom As Boolean = (filterPhantomPrev = 1)
                Dim BOMcreate As Boolean = (BOMcreatePrev = 1)
                Dim Xulyfileloc As Integer = 2

                If useSaved Then
                    Try : asmScale = asmParams.UserParameters.Item(P_AsmScale).Value : Catch : End Try
                    Try : partScale = asmParams.UserParameters.Item(P_PartScale).Value : Catch : End Try
                    Try : viewType = CInt(asmParams.UserParameters.Item(P_ViewType).Value) : Catch : End Try
                    Try : partsPerSheet = CInt(asmParams.UserParameters.Item(P_PartPerSheet).Value) : Catch : End Try
                    Try : filterPurchased = (CInt(asmParams.UserParameters.Item(P_FilterPurchased).Value) = 1) : Catch : End Try
                    Try : filterPhantom = (CInt(asmParams.UserParameters.Item(P_FilterPhantom).Value) = 1) : Catch : End Try
                    Try : BOMcreate = (CInt(asmParams.UserParameters.Item(P_BOMcreate).Value) = 1) : Catch : End Try
                    Try
                        Dim ss As Integer = CInt(asmParams.UserParameters.Item(P_SheetSize).Value)
                        sheetSizeEnum = SheetSizeFromInt(ss)
                    Catch
                    End Try
                Else
                    Dim sheetSizeInput As String = InputBox("Khổ giấy (A0/A1/A2/A3/A4). Enter = A3", "Khổ giấy", "A3")
                    If sheetSizeInput = "" Then sheetSizeInput = "A3"
                    sheetSizeChoice = sheetSizeInput.Trim().ToUpper()
                    sheetSizeEnum = SheetSizeFromString(sheetSizeChoice)

                    Dim asmScaleInput As String = InputBox(
                    "Tỉ lệ Assembly (20 = 1:20). Lần trước: " & Math.Round(1 / asmScalePrev, 2),
                    "Tỉ lệ Assembly", Math.Round(1 / asmScalePrev, 2).ToString())
                    If asmScaleInput = "" Then Exit Sub
                    Try : asmScale = 1.0 / CDbl(asmScaleInput) : Catch : asmScale = asmScalePrev : End Try

                    Dim partScaleInput As String = InputBox(
                    "Tỉ lệ Part (10 = 1:10). Lần trước: " & Math.Round(1 / partScalePrev, 2),
                    "Tỉ lệ Part", Math.Round(1 / partScalePrev, 2).ToString())
                    If partScaleInput = "" Then Exit Sub
                    Try : partScale = 1.0 / CDbl(partScaleInput) : Catch : partScale = partScalePrev : End Try

                    Dim viewTypeInput As String = InputBox(
                    "Kiểu view (1..4):" & vbCrLf &
                    "1: Front" & vbCrLf &
                    "2: Front + Right" & vbCrLf &
                    "3: Front + Right + Iso" & vbCrLf &
                    "4: Front + Top + Right + Iso",
                    "Kiểu view", viewTypePrev.ToString())
                    If viewTypeInput = "" Then Exit Sub
                    Try : viewType = CInt(viewTypeInput) : Catch : viewType = viewTypePrev : End Try
                    If viewType < 1 Or viewType > 4 Then viewType = viewTypePrev

                    Dim partsPerSheetInput As String = InputBox("Số chi tiết / Sheet:", "Số chi tiết / Sheet", partsPerSheetPrev.ToString())
                    If partsPerSheetInput = "" Then Exit Sub
                    Try : partsPerSheet = CInt(partsPerSheetInput) : Catch : partsPerSheet = partsPerSheetPrev : End Try
                    If partsPerSheet < 1 Then partsPerSheet = 1

                    Dim filterInput As String = InputBox(
                    "Lọc (1..8):" & vbCrLf &
                    "1=Purchased" & vbCrLf &
                    "2=Reference" & vbCrLf &
                    "3=Phantom" & vbCrLf &
                    "4=Purchased+Phantom" & vbCrLf &
                    "5=Reference+Phantom" & vbCrLf &
                    "6=Purchased+Reference" & vbCrLf &
                    "7=All" & vbCrLf &
                    "8=Inseparable (Hàn)",
                    "Bộ lọc", "2")
                    If filterInput = "" Then
                        Xulyfileloc = 2
                    Else
                        Try : Xulyfileloc = CInt(filterInput) : Catch : Xulyfileloc = 2 : End Try
                    End If
                    filterPurchased = True

                    BOMcreate = (MessageBox.Show("Tạo BOM cho các cụm lắp?", "Tạo BOM",
                             MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes)

                    Try
                        asmParams.UserParameters.Item(P_AsmScale).Value = asmScale
                        asmParams.UserParameters.Item(P_PartScale).Value = partScale
                        asmParams.UserParameters.Item(P_SheetSize).Value = SheetSizeToInt(sheetSizeChoice)
                        asmParams.UserParameters.Item(P_ViewType).Value = viewType
                        asmParams.UserParameters.Item(P_PartPerSheet).Value = partsPerSheet
                        asmParams.UserParameters.Item(P_FilterPurchased).Value = If(filterPurchased, 1, 0)
                        asmParams.UserParameters.Item(P_FilterPhantom).Value = If(filterPhantom, 1, 0)
                        asmParams.UserParameters.Item(P_BOMcreate).Value = If(BOMcreate, 1, 0)
                    Catch
                    End Try
                End If

                '=================================================
                ' 5. TẠO / MỞ DRAWING
                '=================================================
                Dim drawDoc As Inventor.DrawingDocument = Nothing
                Dim tg As Inventor.TransientGeometry = app.TransientGeometry

                Dim createChoice As String = InputBox(
                "1 = Bản vẽ mới" & vbCrLf & "2 = Bản vẽ có sẵn",
                "Kiểu bản vẽ", "1")

                If createChoice = "2" Then
                    Dim oFileDlg As Inventor.FileDialog = Nothing
                    app.CreateFileDialog(oFileDlg)
                    oFileDlg.Filter = "Bản vẽ Inventor (*.idw)|*.idw"
                    oFileDlg.DialogTitle = "Chọn file bản vẽ"
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
                Else
                    Dim sfd As New SaveFileDialog()
                    sfd.Filter = "Inventor Drawing (*.idw)|*.idw"
                    sfd.Title = "Lưu bản vẽ mới"
                    sfd.FileName = asmDoc.DisplayName & "_Drawing.idw"
                    If sfd.ShowDialog() <> DialogResult.OK Then Exit Sub

                    Dim template As String = app.FileManager.GetTemplateFile(Inventor.DocumentTypeEnum.kDrawingDocumentObject)
                    drawDoc = CType(app.Documents.Add(Inventor.DocumentTypeEnum.kDrawingDocumentObject, template, True), Inventor.DrawingDocument)
                    drawDoc.SaveAs(sfd.FileName, False)
                End If

                '=================================================
                ' 6. SHEET CỤM TỔNG
                '=================================================
                Try
                    Dim asmSheet As Inventor.Sheet = drawDoc.Sheets.Add(sheetSizeEnum)
                    asmSheet.Size = sheetSizeEnum
                    asmSheet.Name = "Bản lắp tổng " & asmDoc.DisplayName.Replace(".", "_")

                    ' ★ Border + Khung tên
                    ApplyBorderAndTitleBlock(drawDoc, asmSheet, sheetSizeEnum)

                    Dim cx As Double = asmSheet.Width / 3.0
                    Dim cy As Double = asmSheet.Height / 4.0 * 3.0
                    Dim basePt As Inventor.Point2d = tg.CreatePoint2d(cx, cy)

                    Dim baseViewTot As Inventor.DrawingView = asmSheet.DrawingViews.AddBaseView(
                    asmDoc, basePt, asmScale,
                    Inventor.ViewOrientationTypeEnum.kFrontViewOrientation,
                    Inventor.DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
                    baseViewTot.ShowLabel = True

                    AddProjectedViews(asmSheet, baseViewTot, viewType, cx, cy, asmSheet.Width, asmSheet.Height, tg)

                    If BOMcreate Then
                        AddPartsListSafe(drawDoc, asmSheet, baseViewTot, asmDoc, tg, "BẢNG KÊ VẬT TƯ")
                    End If
                Catch ex As Exception
                    MessageBox.Show("Lỗi sheet cụm tổng:" & vbCrLf & ex.Message, "Cảnh báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try

                '=================================================
                ' 7. DUYỆT CÂY – LỌC TRÙNG PART NUMBER
                '=================================================
                Dim stack As New Stack()
                stack.Push(asmDoc)

                Dim topParts As New ArrayList()
                Dim visitedAssemblies As New Hashtable()
                Dim usedPartNumbers As New Hashtable()   ' ★ lọc trùng PN

                While stack.Count > 0
                    Try
                        Dim currentAsm As Inventor.AssemblyDocument = CType(stack.Pop(), Inventor.AssemblyDocument)

                        If visitedAssemblies.ContainsKey(currentAsm.InternalName) Then
                            Continue While
                        End If
                        visitedAssemblies.Add(currentAsm.InternalName, True)

                        Dim isTopAsm As Boolean = (currentAsm.InternalName = asmDoc.InternalName)

                        '----- Sheet sub-assembly -----
                        If Not isTopAsm Then
                            Try
                                Dim sheetA As Inventor.Sheet = drawDoc.Sheets.Add(sheetSizeEnum)
                                sheetA.Size = sheetSizeEnum
                                sheetA.Name = "Bản lắp " & currentAsm.DisplayName.Replace(".", "_")

                                ' ★ Border + Khung tên
                                ApplyBorderAndTitleBlock(drawDoc, sheetA, sheetSizeEnum)

                                Dim cxA As Double = sheetA.Width / 3.0
                                Dim cyA As Double = sheetA.Height / 2.0
                                Dim basePtA As Inventor.Point2d = tg.CreatePoint2d(cxA, cyA)

                                Dim baseViewA As Inventor.DrawingView = sheetA.DrawingViews.AddBaseView(
                                currentAsm, basePtA, asmScale,
                                Inventor.ViewOrientationTypeEnum.kFrontViewOrientation,
                                Inventor.DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
                                baseViewA.ShowLabel = True

                                AddProjectedViews(sheetA, baseViewA, viewType, cxA, cyA, sheetA.Width, sheetA.Height, tg)

                                If BOMcreate Then
                                    AddPartsListSafe(drawDoc, sheetA, baseViewA, currentAsm, tg, "BẢNG KÊ VẬT TƯ")
                                End If
                            Catch
                            End Try
                        End If

                        '----- Collect parts + child asm -----
                        Dim localParts As New ArrayList()
                        Dim childAsmList As New ArrayList()

                        For Each occ As Inventor.ComponentOccurrence In currentAsm.ComponentDefinition.Occurrences
                            Try
                                Dim refDoc As Inventor.Document = occ.Definition.Document

                                If filterPurchased Then
                                    Try
                                        Dim bs As Inventor.BOMStructureEnum = occ.BOMStructure
                                        If ShouldSkip(bs, Xulyfileloc) Then Continue For
                                    Catch
                                    End Try
                                End If

                                ' ★ LỌC TRÙNG PART NUMBER
                                If AlreadyUsed(usedPartNumbers, refDoc) Then
                                    Continue For
                                End If

                                If refDoc.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                                    childAsmList.Add(CType(refDoc, Inventor.AssemblyDocument))
                                ElseIf refDoc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
                                    If isTopAsm Then
                                        topParts.Add(refDoc)
                                    Else
                                        localParts.Add(refDoc)
                                    End If
                                End If
                            Catch
                            End Try
                        Next

                        DrawPartsOnSheets(drawDoc, localParts, sheetSizeEnum, partScale, partsPerSheet, viewType, tg, "Cụm chi tiết ")

                        For i As Integer = childAsmList.Count - 1 To 0 Step -1
                            Try
                                Dim childAsm As Inventor.AssemblyDocument = CType(childAsmList(i), Inventor.AssemblyDocument)
                                If Not visitedAssemblies.ContainsKey(childAsm.InternalName) Then
                                    stack.Push(childAsm)
                                End If
                            Catch
                            End Try
                        Next

                    Catch
                    End Try
                End While

                '=================================================
                ' 8. PARTS CỦA CỤM TỔNG
                '=================================================
                DrawPartsOnSheets(drawDoc, topParts, sheetSizeEnum, partScale, partsPerSheet, viewType, tg, "Chi tiết cụm tổng ")

                '=================================================
                ' 9. HOÀN TẤT
                '=================================================
                Try
                    drawDoc.Update()
                    MessageBox.Show("Hoàn tất: Đã tạo bản vẽ theo cấu hình.", "Hoàn tất",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Hoàn tất (có lỗi cập nhật): " & ex.Message, "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Ban xuat BV",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        '=================================================
        ' PARTS LIST – ỔN ĐỊNH + STYLE
        '=================================================
        Private Sub AddPartsListSafe(
        drawDoc As Inventor.DrawingDocument,
        sheet As Inventor.Sheet,
        baseView As Inventor.DrawingView,
        sourceAsm As Inventor.AssemblyDocument,
        tg As Inventor.TransientGeometry,
        title As String)

            Try
                Try
                    Dim bom As Inventor.BOM = sourceAsm.ComponentDefinition.BOM
                    bom.StructuredViewEnabled = True
                    bom.StructuredViewFirstLevelOnly = False
                    Try : bom.PartsOnlyViewEnabled = True : Catch : End Try
                Catch
                End Try

                Try : sourceAsm.Update2(True) : Catch : Try : sourceAsm.Update() : Catch : End Try : End Try
                drawDoc.Update()
                System.Windows.Forms.Application.DoEvents()

                Dim pt As Inventor.Point2d = tg.CreatePoint2d(sheet.Width * 0.95, sheet.Height * 0.95)
                Dim pl As Inventor.PartsList = Nothing
                Dim lastError As String = ""

                If pl Is Nothing Then
                    Try
                        pl = sheet.PartsLists.Add(baseView, pt, Inventor.PartsListLevelEnum.kStructuredAllLevels)
                    Catch ex As Exception
                        lastError = ex.Message
                    End Try
                End If

                If pl Is Nothing Then
                    Try
                        pl = sheet.PartsLists.Add(baseView, pt, Inventor.PartsListLevelEnum.kFirstLevelComponents)
                    Catch ex As Exception
                        lastError = ex.Message
                    End Try
                End If

                If pl Is Nothing Then
                    Try
                        pl = sheet.PartsLists.Add(baseView, pt, Inventor.PartsListLevelEnum.kPartsOnly)
                    Catch ex As Exception
                        lastError = ex.Message
                    End Try
                End If

                If pl Is Nothing Then
                    Try
                        pl = sheet.PartsLists.Add(baseView, pt)
                    Catch ex As Exception
                        lastError = ex.Message
                    End Try
                End If

                If pl Is Nothing Then
                    Try
                        pl = sheet.PartsLists.Add(sourceAsm, pt, Inventor.PartsListLevelEnum.kFirstLevelComponents)
                    Catch ex As Exception
                        lastError = ex.Message
                    End Try
                End If

                If pl Is Nothing Then
                    MessageBox.Show("Không tạo được Parts List:" & vbCrLf & sourceAsm.DisplayName & vbCrLf & lastError,
                                "BOM", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                ' ★ ĐỔI TÊN STYLE TẠI ĐÂY (nếu có style riêng)
                Try
                    If drawDoc.StylesManager.PartsListStyles.Count > 0 Then
                        ' pl.Style = drawDoc.StylesManager.PartsListStyles.Item("TÊN_STYLE_CỦA_BẠN")
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
                drawDoc.Update()

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
            Try
                Return doc.DisplayName.ToUpper()
            Catch
            End Try
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
        ' HELPER KHÁC
        '=================================================
        Private Sub EnsureParam(params As Inventor.Parameters, name As String, value As Double)
            Try
                Dim t = params.UserParameters.Item(name)
            Catch
                Try
                    params.UserParameters.AddByValue(name, value, Inventor.UnitsTypeEnum.kUnitlessUnits)
                Catch
                End Try
            End Try
        End Sub

        Private Function SheetSizeFromString(s As String) As Inventor.DrawingSheetSizeEnum
            Select Case s.ToUpper().Trim()
                Case "A0" : Return Inventor.DrawingSheetSizeEnum.kA0DrawingSheetSize
                Case "A1" : Return Inventor.DrawingSheetSizeEnum.kA1DrawingSheetSize
                Case "A2" : Return Inventor.DrawingSheetSizeEnum.kA2DrawingSheetSize
                Case "A4" : Return Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize
                Case Else : Return Inventor.DrawingSheetSizeEnum.kA3DrawingSheetSize
            End Select
        End Function

        Private Function SheetSizeFromInt(ss As Integer) As Inventor.DrawingSheetSizeEnum
            Select Case ss
                Case 0 : Return Inventor.DrawingSheetSizeEnum.kA0DrawingSheetSize
                Case 1 : Return Inventor.DrawingSheetSizeEnum.kA1DrawingSheetSize
                Case 2 : Return Inventor.DrawingSheetSizeEnum.kA2DrawingSheetSize
                Case 4 : Return Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize
                Case Else : Return Inventor.DrawingSheetSizeEnum.kA3DrawingSheetSize
            End Select
        End Function

        Private Function SheetSizeToInt(s As String) As Integer
            Select Case s.ToUpper().Trim()
                Case "A0" : Return 0
                Case "A1" : Return 1
                Case "A2" : Return 2
                Case "A4" : Return 4
                Case Else : Return 3
            End Select
        End Function

        Private Function ShouldSkip(bs As Inventor.BOMStructureEnum, mode As Integer) As Boolean
            Select Case mode
                Case 1 : Return bs = Inventor.BOMStructureEnum.kPurchasedBOMStructure
                Case 2 : Return bs = Inventor.BOMStructureEnum.kReferenceBOMStructure
                Case 3 : Return bs = Inventor.BOMStructureEnum.kPhantomBOMStructure
                Case 4 : Return bs = Inventor.BOMStructureEnum.kPurchasedBOMStructure OrElse bs = Inventor.BOMStructureEnum.kPhantomBOMStructure
                Case 5 : Return bs = Inventor.BOMStructureEnum.kPhantomBOMStructure OrElse bs = Inventor.BOMStructureEnum.kReferenceBOMStructure
                Case 6 : Return bs = Inventor.BOMStructureEnum.kPurchasedBOMStructure OrElse bs = Inventor.BOMStructureEnum.kReferenceBOMStructure
                Case 7 : Return bs = Inventor.BOMStructureEnum.kPurchasedBOMStructure OrElse bs = Inventor.BOMStructureEnum.kPhantomBOMStructure OrElse bs = Inventor.BOMStructureEnum.kReferenceBOMStructure
                Case 8 : Return bs = Inventor.BOMStructureEnum.kInseparableBOMStructure
                Case Else : Return False
            End Select
        End Function

        Private Sub AddProjectedViews(
        sheet As Inventor.Sheet,
        baseView As Inventor.DrawingView,
        viewType As Integer,
        cx As Double, cy As Double,
        sheetW As Double, sheetH As Double,
        tg As Inventor.TransientGeometry)

            If viewType >= 2 Then
                Dim rightPt As Inventor.Point2d = tg.CreatePoint2d(cx + sheetW / 3.0, cy)
                sheet.DrawingViews.AddProjectedView(baseView, rightPt,
                Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
            End If
            If viewType = 3 OrElse viewType = 4 Then
                Dim isoPt As Inventor.Point2d = tg.CreatePoint2d(cx + sheetW / 4.0, cy - sheetH / 3.0)
                sheet.DrawingViews.AddProjectedView(baseView, isoPt,
                Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
            End If
            If viewType = 4 Then
                Dim topPt As Inventor.Point2d = tg.CreatePoint2d(cx, cy - sheetH / 3.0)
                sheet.DrawingViews.AddProjectedView(baseView, topPt,
                Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
            End If
        End Sub

        Private Sub DrawPartsOnSheets(
        drawDoc As Inventor.DrawingDocument,
        parts As ArrayList,
        sheetSizeEnum As Inventor.DrawingSheetSizeEnum,
        partScale As Double,
        partsPerSheet As Integer,
        viewType As Integer,
        tg As Inventor.TransientGeometry,
        sheetNamePrefix As String)

            If parts Is Nothing OrElse parts.Count = 0 Then Exit Sub

            Dim partSheet As Inventor.Sheet = Nothing
            Dim sheetWidth, sheetHeight, usableW, usableH, xStart, yStart, xStep, yStep As Double
            Dim partCountOnCurrentSheet As Integer = 0
            Dim cols As Integer = If(partsPerSheet = 1, 1, 2)
            Dim rowsNeeded As Integer = CInt(Math.Ceiling(partsPerSheet / CDbl(cols)))

            For Each docP As Inventor.Document In parts
                Try
                    If partSheet Is Nothing OrElse partCountOnCurrentSheet >= partsPerSheet Then
                        partSheet = drawDoc.Sheets.Add(sheetSizeEnum)
                        partSheet.Size = sheetSizeEnum
                        partSheet.Name = sheetNamePrefix & drawDoc.Sheets.Count.ToString()

                        ' ★ Border + Khung tên
                        ApplyBorderAndTitleBlock(drawDoc, partSheet, sheetSizeEnum)

                        sheetWidth = partSheet.Width
                        sheetHeight = partSheet.Height
                        usableW = sheetWidth * 0.75
                        usableH = sheetHeight * 0.75
                        xStart = sheetWidth / 4.0
                        yStart = sheetHeight * 0.8
                        xStep = usableW / cols
                        yStep = usableH / (rowsNeeded + 1)
                        partCountOnCurrentSheet = 0
                    End If

                    Dim colIndex As Integer = partCountOnCurrentSheet Mod cols
                    Dim rowIndex As Integer = partCountOnCurrentSheet \ cols
                    Dim xPos As Double = xStart + colIndex * xStep
                    Dim yPos As Double = yStart - rowIndex * yStep

                    Dim baseViewP As Inventor.DrawingView = partSheet.DrawingViews.AddBaseView(
                    docP, tg.CreatePoint2d(xPos, yPos), partScale,
                    Inventor.ViewOrientationTypeEnum.kFrontViewOrientation,
                    Inventor.DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
                    baseViewP.ShowLabel = True

                    If viewType >= 2 Then
                        Dim rightPtP As Inventor.Point2d = tg.CreatePoint2d(xPos + xStep * 0.6, yPos)
                        partSheet.DrawingViews.AddProjectedView(baseViewP, rightPtP,
                        Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseViewP.Scale)
                    End If
                    If viewType = 3 OrElse viewType = 4 Then
                        Dim isoPtP As Inventor.Point2d = tg.CreatePoint2d(xPos + xStep * 0.4, yPos - yStep * 0.66)
                        partSheet.DrawingViews.AddProjectedView(baseViewP, isoPtP,
                        Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseViewP.Scale)
                    End If
                    If viewType = 4 Then
                        Dim topPtP As Inventor.Point2d = tg.CreatePoint2d(xPos, yPos - yStep * 0.2)
                        partSheet.DrawingViews.AddProjectedView(baseViewP, topPtP,
                        Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseViewP.Scale)
                    End If

                    partCountOnCurrentSheet += 1
                Catch
                End Try
            Next
        End Sub
        Private Sub ApplyBorderAndTitleBlock(
    drawDoc As Inventor.DrawingDocument,
    sheet As Inventor.Sheet,
    sizeEnum As Inventor.DrawingSheetSizeEnum)

            Try
                Dim borderName As String = ""
                Dim titleName As String = ""

                Select Case sizeEnum
                    Case Inventor.DrawingSheetSizeEnum.kA0DrawingSheetSize
                        borderName = "NT A0"
                        titleName = "Khung tên SX NT A0"
                    Case Inventor.DrawingSheetSizeEnum.kA1DrawingSheetSize
                        borderName = "NT A1"
                        titleName = "Khung tên SX NT A1"
                    Case Inventor.DrawingSheetSizeEnum.kA2DrawingSheetSize
                        borderName = "NT A2"
                        titleName = "Khung tên SX NT A2"
                    Case Inventor.DrawingSheetSizeEnum.kA3DrawingSheetSize
                        borderName = "NT A3"
                        titleName = "Khung tên SX NT A3"
                    Case Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize
                        ' A4 thường
                        borderName = "NT A4"
                        titleName = "Khung tên SX NT A4"
                    Case Else
                        borderName = "NT A3"
                        titleName = "Khung tên SX NT A3"
                End Select

                ' Xóa border / title cũ nếu có
                Try
                    If sheet.Border IsNot Nothing Then sheet.Border.Delete()
                Catch
                End Try
                Try
                    If sheet.TitleBlock IsNot Nothing Then sheet.TitleBlock.Delete()
                Catch
                End Try

                ' Gán Border
                Try
                    Dim bd As Inventor.BorderDefinition =
                        drawDoc.BorderDefinitions.Item(borderName)
                    sheet.AddBorder(bd)
                Catch
                    ' Thử A4 D nếu A4 thường không có
                    If sizeEnum = Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize Then
                        Try
                            Dim bd2 As Inventor.BorderDefinition =
                                drawDoc.BorderDefinitions.Item("NT A4 D")
                            sheet.AddBorder(bd2)
                        Catch
                        End Try
                    End If
                End Try

                ' Gán TitleBlock
                Try
                    Dim td As Inventor.TitleBlockDefinition =
                        drawDoc.TitleBlockDefinitions.Item(titleName)
                    sheet.AddTitleBlock(td)
                Catch
                    If sizeEnum = Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize Then
                        Try
                            Dim td2 As Inventor.TitleBlockDefinition =
                                drawDoc.TitleBlockDefinitions.Item("Khung tên SX NT A4 D")
                            sheet.AddTitleBlock(td2)
                        Catch
                        End Try
                    End If
                End Try

            Catch
            End Try
        End Sub
    End Module

End Namespace