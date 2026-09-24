Option Explicit On
Option Strict Off

' ⭐ BỎ "Imports Inventor" — code đã dùng "Inventor.XXX" đầy đủ
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Collections
Imports System.Collections.Generic

Namespace ToolInventor2025.Assembly.Buttons.AutoCreateDrawing

    Public Module AutoDrawingV8

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim app As Inventor.Application = g_inventorApplication
            Try
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

                Dim P_AsmScale As String = "iLogic_AsmScale"
                Dim P_PartScale As String = "iLogic_PartScale"
                Dim P_SheetSize As String = "iLogic_SheetSize"
                Dim P_ViewType As String = "iLogic_ViewType"
                Dim P_PartPerSheet As String = "iLogic_PartPerSheet"
                Dim P_FilterPurchased As String = "iLogic_FilterPurchased"
                Dim P_FilterPhantom As String = "iLogic_FilterPhantom"
                Dim P_BOMcreate As String = "iLogic_BOMcreate"
                Dim P_Xulyfileloc As String = "iLogic_Xulyfileloc"
                Dim P_CreateMode As String = "iLogic_CreateMode"

                Dim asmScalePrev As Double = 1.0 / 20.0
                Dim partScalePrev As Double = 1.0 / 10.0
                Dim sheetSizePrev As Integer = 3
                Dim viewTypePrev As Integer = 3
                Dim partsPerSheetPrev As Integer = 4
                Dim filterPurchasedPrev As Integer = 1
                Dim filterPhantomPrev As Integer = 0
                Dim BOMcreatePrev As Integer = 1
                Dim XulyfilelocPrev As Integer = 2
                Dim createModePrev As Integer = 1

                Try : asmScalePrev = asmParams.UserParameters.Item(P_AsmScale).Value : Catch : End Try
                Try : partScalePrev = asmParams.UserParameters.Item(P_PartScale).Value : Catch : End Try
                Try : sheetSizePrev = CInt(asmParams.UserParameters.Item(P_SheetSize).Value) : Catch : End Try
                Try : viewTypePrev = CInt(asmParams.UserParameters.Item(P_ViewType).Value) : Catch : End Try
                Try : partsPerSheetPrev = CInt(asmParams.UserParameters.Item(P_PartPerSheet).Value) : Catch : End Try
                Try : filterPurchasedPrev = CInt(asmParams.UserParameters.Item(P_FilterPurchased).Value) : Catch : End Try
                Try : filterPhantomPrev = CInt(asmParams.UserParameters.Item(P_FilterPhantom).Value) : Catch : End Try
                Try : BOMcreatePrev = CInt(asmParams.UserParameters.Item(P_BOMcreate).Value) : Catch : End Try
                Try : XulyfilelocPrev = CInt(asmParams.UserParameters.Item(P_Xulyfileloc).Value) : Catch : End Try
                Try : createModePrev = CInt(asmParams.UserParameters.Item(P_CreateMode).Value) : Catch : End Try

                EnsureParam(asmParams, P_AsmScale, asmScalePrev)
                EnsureParam(asmParams, P_PartScale, partScalePrev)
                EnsureParam(asmParams, P_SheetSize, sheetSizePrev)
                EnsureParam(asmParams, P_ViewType, viewTypePrev)
                EnsureParam(asmParams, P_PartPerSheet, partsPerSheetPrev)
                EnsureParam(asmParams, P_FilterPurchased, filterPurchasedPrev)
                EnsureParam(asmParams, P_FilterPhantom, filterPhantomPrev)
                EnsureParam(asmParams, P_BOMcreate, BOMcreatePrev)
                EnsureParam(asmParams, P_Xulyfileloc, XulyfilelocPrev)
                EnsureParam(asmParams, P_CreateMode, createModePrev)

                Dim haveSavedConfig As Boolean = False
                Try
                    Dim t = asmParams.UserParameters.Item(P_AsmScale)
                    haveSavedConfig = True
                Catch
                End Try

                Dim opt As V8Options
                Using frm As New V8OptionsForm(haveSavedConfig)
                    frm.SetInitialValues(
                        SheetSizeFromInt(sheetSizePrev),
                        If(asmScalePrev > 0, 1.0 / asmScalePrev, 20.0),
                        If(partScalePrev > 0, 1.0 / partScalePrev, 10.0),
                        viewTypePrev,
                        partsPerSheetPrev,
                        XulyfilelocPrev,
                        BOMcreatePrev = 1,
                        createModePrev)
                    If frm.ShowDialog() <> DialogResult.OK Then Exit Sub
                    opt = frm.Options
                End Using

                Dim asmScale As Double = opt.AsmScale
                Dim partScale As Double = opt.PartScale
                Dim sheetSizeEnum As Inventor.DrawingSheetSizeEnum = opt.SheetSize
                Dim sheetSizeChoice As String = opt.SheetSizeName
                Dim viewType As Integer = opt.ViewType
                Dim partsPerSheet As Integer = opt.PartsPerSheet
                Dim filterPurchased As Boolean = True
                Dim filterPhantom As Boolean = (opt.Xulyfileloc = 3 OrElse opt.Xulyfileloc = 4 OrElse
                                                opt.Xulyfileloc = 5 OrElse opt.Xulyfileloc = 7)
                Dim BOMcreate As Boolean = opt.CreateBOM
                Dim Xulyfileloc As Integer = opt.Xulyfileloc
                Dim createMode As Integer = opt.CreateMode

                Try
                    asmParams.UserParameters.Item(P_AsmScale).Value = asmScale
                    asmParams.UserParameters.Item(P_PartScale).Value = partScale
                    asmParams.UserParameters.Item(P_SheetSize).Value = SheetSizeToInt(sheetSizeChoice)
                    asmParams.UserParameters.Item(P_ViewType).Value = viewType
                    asmParams.UserParameters.Item(P_PartPerSheet).Value = partsPerSheet
                    asmParams.UserParameters.Item(P_FilterPurchased).Value = If(filterPurchased, 1, 0)
                    asmParams.UserParameters.Item(P_FilterPhantom).Value = If(filterPhantom, 1, 0)
                    asmParams.UserParameters.Item(P_BOMcreate).Value = If(BOMcreate, 1, 0)
                    asmParams.UserParameters.Item(P_Xulyfileloc).Value = Xulyfileloc
                    asmParams.UserParameters.Item(P_CreateMode).Value = createMode
                Catch
                End Try

                Dim drawDoc As Inventor.DrawingDocument = Nothing
                Dim tg As Inventor.TransientGeometry = app.TransientGeometry

                If createMode = 2 Then
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

                '===== 6. SHEET CỤM TỔNG =====
                Try
                    Dim asmSheet As Inventor.Sheet = drawDoc.Sheets.Add(sheetSizeEnum)
                    asmSheet.Size = sheetSizeEnum
                    asmSheet.Name = "Bản lắp tổng " & asmDoc.DisplayName.Replace(".", "_")
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

                '===== 7. DUYỆT CÂY =====
                Dim stack As New Stack()
                stack.Push(asmDoc)

                Dim topParts As New ArrayList()
                Dim visitedAssemblies As New Hashtable()
                Dim usedPartNumbers As New Hashtable()

                While stack.Count > 0
                    Try
                        Dim currentAsm As Inventor.AssemblyDocument = CType(stack.Pop(), Inventor.AssemblyDocument)

                        If visitedAssemblies.ContainsKey(currentAsm.InternalName) Then Continue While
                        visitedAssemblies.Add(currentAsm.InternalName, True)

                        Dim isTopAsm As Boolean = (currentAsm.InternalName = asmDoc.InternalName)

                        If Not isTopAsm Then
                            Try
                                Dim sheetA As Inventor.Sheet = drawDoc.Sheets.Add(sheetSizeEnum)
                                sheetA.Size = sheetSizeEnum
                                sheetA.Name = "Bản lắp " & currentAsm.DisplayName.Replace(".", "_")
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

                                If AlreadyUsed(usedPartNumbers, refDoc) Then Continue For

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

                DrawPartsOnSheets(drawDoc, topParts, sheetSizeEnum, partScale, partsPerSheet, viewType, tg, "Chi tiết cụm tổng ")

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

        '===== HELPER (giữ nguyên) =====
        Private Sub AddPartsListSafe(drawDoc As Inventor.DrawingDocument,
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

                Try : pl = sheet.PartsLists.Add(baseView, pt, Inventor.PartsListLevelEnum.kStructuredAllLevels) : Catch : End Try
                If pl Is Nothing Then Try : pl = sheet.PartsLists.Add(baseView, pt, Inventor.PartsListLevelEnum.kFirstLevelComponents) : Catch : End Try
                If pl Is Nothing Then Try : pl = sheet.PartsLists.Add(baseView, pt, Inventor.PartsListLevelEnum.kPartsOnly) : Catch : End Try
                If pl Is Nothing Then Try : pl = sheet.PartsLists.Add(baseView, pt) : Catch : End Try
                If pl Is Nothing Then Try : pl = sheet.PartsLists.Add(sourceAsm, pt, Inventor.PartsListLevelEnum.kFirstLevelComponents) : Catch : End Try

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
                drawDoc.Update()
            Catch
            End Try
        End Sub

        Private Function GetPartNumber(doc As Inventor.Document) As String
            Try
                Dim ps As Inventor.PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim pn As String = CStr(ps.Item("Part Number").Value).Trim().ToUpper()
                If pn <> "" Then Return pn
            Catch
            End Try
            Try : Return System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName).ToUpper() : Catch : End Try
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

        Private Sub AddProjectedViews(sheet As Inventor.Sheet,
                                      baseView As Inventor.DrawingView,
                                      viewType As Integer,
                                      cx As Double, cy As Double,
                                      sheetW As Double, sheetH As Double,
                                      tg As Inventor.TransientGeometry)
            If viewType >= 2 Then
                sheet.DrawingViews.AddProjectedView(baseView, tg.CreatePoint2d(cx + sheetW / 3.0, cy),
                                                    Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
            End If
            If viewType = 3 OrElse viewType = 4 Then
                sheet.DrawingViews.AddProjectedView(baseView, tg.CreatePoint2d(cx + sheetW / 4.0, cy - sheetH / 3.0),
                                                    Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
            End If
            If viewType = 4 Then
                sheet.DrawingViews.AddProjectedView(baseView, tg.CreatePoint2d(cx, cy - sheetH / 3.0),
                                                    Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
            End If
        End Sub

        Private Sub DrawPartsOnSheets(drawDoc As Inventor.DrawingDocument,
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
                        partSheet.DrawingViews.AddProjectedView(baseViewP, tg.CreatePoint2d(xPos + xStep * 0.6, yPos),
                                                                Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseViewP.Scale)
                    End If
                    If viewType = 3 OrElse viewType = 4 Then
                        partSheet.DrawingViews.AddProjectedView(baseViewP, tg.CreatePoint2d(xPos + xStep * 0.4, yPos - yStep * 0.66),
                                                                Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseViewP.Scale)
                    End If
                    If viewType = 4 Then
                        partSheet.DrawingViews.AddProjectedView(baseViewP, tg.CreatePoint2d(xPos, yPos - yStep * 0.2),
                                                                Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseViewP.Scale)
                    End If

                    partCountOnCurrentSheet += 1
                Catch
                End Try
            Next
        End Sub

        Private Sub ApplyBorderAndTitleBlock(drawDoc As Inventor.DrawingDocument,
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

                Try : If sheet.Border IsNot Nothing Then sheet.Border.Delete()
                Catch : End Try
                Try : If sheet.TitleBlock IsNot Nothing Then sheet.TitleBlock.Delete()
                Catch : End Try

                Try
                    sheet.AddBorder(drawDoc.BorderDefinitions.Item(borderName))
                Catch
                    If sizeEnum = Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize Then
                        Try : sheet.AddBorder(drawDoc.BorderDefinitions.Item("NT A4 D"))
                        Catch : End Try
                    End If
                End Try
                Try
                    sheet.AddTitleBlock(drawDoc.TitleBlockDefinitions.Item(titleName))
                Catch
                    If sizeEnum = Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize Then
                        Try : sheet.AddTitleBlock(drawDoc.TitleBlockDefinitions.Item("Khung tên SX NT A4 D"))
                        Catch : End Try
                    End If
                End Try
            Catch
            End Try
        End Sub

    End Module


    '==========================================================
    ' CLASS OPTIONS
    '==========================================================
    Public Class V8Options
        Public SheetSize As Inventor.DrawingSheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA3DrawingSheetSize
        Public SheetSizeName As String = "A3"
        Public AsmScale As Double = 1.0 / 20.0
        Public PartScale As Double = 1.0 / 10.0
        Public ViewType As Integer = 3
        Public PartsPerSheet As Integer = 4
        Public Xulyfileloc As Integer = 2
        Public CreateBOM As Boolean = True
        Public CreateMode As Integer = 1
    End Class


    '==========================================================
    ' FORM OPTIONS — FIX CỨNG, CĂN GIỮA MÀN HÌNH
    '==========================================================
    Public Class V8OptionsForm
        Inherits Form

        Public Property Options As V8Options

        Private cboSize As ComboBox
        Private txtAsmScale As TextBox
        Private txtPartScale As TextBox
        Private cboView As ComboBox
        Private txtPartsPerSheet As TextBox
        Private cboFilter As ComboBox
        Private chkBOM As CheckBox
        Private rdoNew As RadioButton
        Private rdoExisting As RadioButton
        Private chkUseSaved As CheckBox

        Private Const FORM_W As Integer = 580
        Private Const FORM_H As Integer = 620

        Public Sub New(ByVal hasSavedConfig As Boolean)
            Me.Text = "Auto Drawing V8 — Tùy chọn"
            Me.AutoScaleMode = AutoScaleMode.None
            Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
            Me.FormBorderStyle = FormBorderStyle.FixedSingle
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.BackColor = Color.FromArgb(245, 245, 245)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.ClientSize = New Size(FORM_W, FORM_H)

            '===== HEADER =====
            Dim pnlHeader As New Panel()
            pnlHeader.Location = New Point(0, 0)
            pnlHeader.Size = New Size(FORM_W, 50)
            pnlHeader.BackColor = Color.FromArgb(45, 100, 180)
            Me.Controls.Add(pnlHeader)

            Dim lblTitle As New Label()
            lblTitle.Text = "AUTO DRAWING V8 — TÙY CHỌN"
            lblTitle.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold, GraphicsUnit.Point)
            lblTitle.ForeColor = Color.White
            lblTitle.Dock = DockStyle.Fill
            lblTitle.TextAlign = ContentAlignment.MiddleCenter
            pnlHeader.Controls.Add(lblTitle)

            '===== CHECKBOX DÙNG LẠI CẤU HÌNH =====
            chkUseSaved = New CheckBox() With {
                .Text = "Dùng lại cấu hình đã lưu (nếu có)",
                .Location = New Point(15, 60),
                .Size = New Size(300, 25),
                .ForeColor = Color.FromArgb(45, 100, 180),
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point)}
            chkUseSaved.Enabled = hasSavedConfig
            If hasSavedConfig Then chkUseSaved.Checked = True
            Me.Controls.Add(chkUseSaved)

            '===== GROUP 1: KHỔ GIẤY / TỈ LỆ =====
            Dim gb1 As New GroupBox() With {
                .Text = "Khổ giấy / Tỉ lệ",
                .Location = New Point(12, 90),
                .Size = New Size(FORM_W - 24, 145),
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                .ForeColor = Color.FromArgb(45, 100, 180),
                .BackColor = Color.White}
            Me.Controls.Add(gb1)

            Dim lblSize As New Label() With {
                .Text = "Khổ giấy:", .Location = New Point(20, 30),
                .Size = New Size(180, 25),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            gb1.Controls.Add(lblSize)

            cboSize = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Location = New Point(210, 30),
                .Size = New Size(120, 25)}
            cboSize.Items.AddRange(New Object() {"A0", "A1", "A2", "A3", "A4"})
            cboSize.SelectedIndex = 3
            gb1.Controls.Add(cboSize)

            Dim lblAsmScale As New Label() With {
                .Text = "Tỉ lệ cụm (20 = 1:20):", .Location = New Point(20, 65),
                .Size = New Size(180, 25),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            gb1.Controls.Add(lblAsmScale)

            txtAsmScale = New TextBox() With {
                .Text = "20", .Location = New Point(210, 65),
                .Size = New Size(120, 25)}
            gb1.Controls.Add(txtAsmScale)

            Dim lblPartScale As New Label() With {
                .Text = "Tỉ lệ Part (10 = 1:10):", .Location = New Point(20, 100),
                .Size = New Size(180, 25),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            gb1.Controls.Add(lblPartScale)

            txtPartScale = New TextBox() With {
                .Text = "10", .Location = New Point(210, 100),
                .Size = New Size(120, 25)}
            gb1.Controls.Add(txtPartScale)

            '===== GROUP 2: VIEW & CHIA SHEET =====
            Dim gb2 As New GroupBox() With {
                .Text = "View & Chia sheet",
                .Location = New Point(12, 245),
                .Size = New Size(FORM_W - 24, 105),
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                .ForeColor = Color.FromArgb(45, 100, 180),
                .BackColor = Color.White}
            Me.Controls.Add(gb2)

            Dim lblView As New Label() With {
                .Text = "Kiểu view:", .Location = New Point(20, 30),
                .Size = New Size(180, 25),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            gb2.Controls.Add(lblView)

            cboView = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Location = New Point(210, 30),
                .Size = New Size(320, 25)}
            cboView.Items.AddRange(New Object() {
                "1 - Front",
                "2 - Front + Right",
                "3 - Front + Right + Iso",
                "4 - Front + Top + Right + Iso"})
            cboView.SelectedIndex = 2
            gb2.Controls.Add(cboView)

            Dim lblPPS As New Label() With {
                .Text = "Số chi tiết / sheet:", .Location = New Point(20, 65),
                .Size = New Size(180, 25),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            gb2.Controls.Add(lblPPS)

            txtPartsPerSheet = New TextBox() With {
                .Text = "4", .Location = New Point(210, 65),
                .Size = New Size(120, 25)}
            gb2.Controls.Add(txtPartsPerSheet)

            '===== GROUP 3: BỘ LỌC =====
            Dim gb3 As New GroupBox() With {
                .Text = "Bộ lọc BOM Structure",
                .Location = New Point(12, 360),
                .Size = New Size(FORM_W - 24, 75),
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                .ForeColor = Color.FromArgb(45, 100, 180),
                .BackColor = Color.White}
            Me.Controls.Add(gb3)

            Dim lblFilter As New Label() With {
                .Text = "Lọc:", .Location = New Point(20, 30),
                .Size = New Size(180, 25),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            gb3.Controls.Add(lblFilter)

            cboFilter = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Location = New Point(210, 30),
                .Size = New Size(320, 25)}
            cboFilter.Items.AddRange(New Object() {
                "1 - Purchased",
                "2 - Reference",
                "3 - Phantom",
                "4 - Purchased + Phantom",
                "5 - Reference + Phantom",
                "6 - Purchased + Reference",
                "7 - All (Purchased + Phantom + Reference)",
                "8 - Inseparable (Hàn)"})
            cboFilter.SelectedIndex = 1
            gb3.Controls.Add(cboFilter)

            '===== GROUP 4: BẢNG KÊ =====
            Dim gb4 As New GroupBox() With {
                .Text = "Bảng kê",
                .Location = New Point(12, 445),
                .Size = New Size(FORM_W - 24, 55),
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                .ForeColor = Color.FromArgb(45, 100, 180),
                .BackColor = Color.White}
            Me.Controls.Add(gb4)

            chkBOM = New CheckBox() With {
                .Text = "Tạo BOM (Parts List) cho các cụm lắp",
                .Location = New Point(20, 22),
                .Size = New Size(400, 25),
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                .Checked = True}
            gb4.Controls.Add(chkBOM)

            '===== GROUP 5: KIỂU BẢN VẼ =====
            Dim gb5 As New GroupBox() With {
                .Text = "Kiểu bản vẽ",
                .Location = New Point(12, 510),
                .Size = New Size(FORM_W - 24, 55),
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                .ForeColor = Color.FromArgb(45, 100, 180),
                .BackColor = Color.White}
            Me.Controls.Add(gb5)

            rdoNew = New RadioButton() With {
                .Text = "Tạo bản vẽ MỚI",
                .Location = New Point(20, 22),
                .Size = New Size(180, 25),
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                .Checked = True}
            rdoExisting = New RadioButton() With {
                .Text = "Mở bản vẽ CÓ SẴN",
                .Location = New Point(220, 22),
                .Size = New Size(200, 25),
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            gb5.Controls.Add(rdoNew)
            gb5.Controls.Add(rdoExisting)

            '===== NÚT =====
            Dim btnOK As New Button()
            btnOK.Text = "TẠO BẢN VẼ"
            btnOK.Size = New Size(150, 38)
            btnOK.Location = New Point(FORM_W - 285, FORM_H - 42)
            btnOK.BackColor = Color.FromArgb(45, 100, 180)
            btnOK.ForeColor = Color.White
            btnOK.FlatStyle = FlatStyle.Flat
            btnOK.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold, GraphicsUnit.Point)
            AddHandler btnOK.Click, AddressOf OnOKClick
            Me.Controls.Add(btnOK)

            Dim btnCancel As New Button()
            btnCancel.Text = "HỦY"
            btnCancel.Size = New Size(100, 38)
            btnCancel.Location = New Point(FORM_W - 125, FORM_H - 42)
            btnCancel.FlatStyle = FlatStyle.Flat
            btnCancel.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            AddHandler btnCancel.Click, Sub()
                                            Me.DialogResult = DialogResult.Cancel
                                            Me.Close()
                                        End Sub
            Me.Controls.Add(btnCancel)
            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub

        Public Sub SetInitialValues(ByVal sheetEnum As Inventor.DrawingSheetSizeEnum,
                                    ByVal asmScaleInv As Double,
                                    ByVal partScaleInv As Double,
                                    ByVal viewType As Integer,
                                    ByVal partsPerSheet As Integer,
                                    ByVal xulyfileloc As Integer,
                                    ByVal createBOM As Boolean,
                                    ByVal createMode As Integer)
            Select Case sheetEnum
                Case Inventor.DrawingSheetSizeEnum.kA0DrawingSheetSize : cboSize.SelectedIndex = 0
                Case Inventor.DrawingSheetSizeEnum.kA1DrawingSheetSize : cboSize.SelectedIndex = 1
                Case Inventor.DrawingSheetSizeEnum.kA2DrawingSheetSize : cboSize.SelectedIndex = 2
                Case Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize : cboSize.SelectedIndex = 4
                Case Else : cboSize.SelectedIndex = 3
            End Select

            txtAsmScale.Text = Math.Round(asmScaleInv, 2).ToString()
            txtPartScale.Text = Math.Round(partScaleInv, 2).ToString()
            If viewType >= 1 AndAlso viewType <= 4 Then cboView.SelectedIndex = viewType - 1
            txtPartsPerSheet.Text = partsPerSheet.ToString()
            If xulyfileloc >= 1 AndAlso xulyfileloc <= 8 Then cboFilter.SelectedIndex = xulyfileloc - 1
            chkBOM.Checked = createBOM
            If createMode = 2 Then
                rdoExisting.Checked = True
            Else
                rdoNew.Checked = True
            End If
        End Sub

        Private Sub OnOKClick(ByVal sender As Object, ByVal e As EventArgs)
            Dim asmScaleVal As Double = 20
            If Not Double.TryParse(txtAsmScale.Text, asmScaleVal) OrElse asmScaleVal <= 0 Then
                MessageBox.Show("Tỉ lệ cụm không hợp lệ.", "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim partScaleVal As Double = 10
            If Not Double.TryParse(txtPartScale.Text, partScaleVal) OrElse partScaleVal <= 0 Then
                MessageBox.Show("Tỉ lệ Part không hợp lệ.", "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim partsVal As Integer = 4
            If Not Integer.TryParse(txtPartsPerSheet.Text, partsVal) OrElse partsVal < 1 Then
                MessageBox.Show("Số chi tiết/sheet không hợp lệ.", "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim sheetEnum As Inventor.DrawingSheetSizeEnum
            Dim sheetName As String
            Select Case cboSize.SelectedIndex
                Case 0 : sheetEnum = Inventor.DrawingSheetSizeEnum.kA0DrawingSheetSize : sheetName = "A0"
                Case 1 : sheetEnum = Inventor.DrawingSheetSizeEnum.kA1DrawingSheetSize : sheetName = "A1"
                Case 2 : sheetEnum = Inventor.DrawingSheetSizeEnum.kA2DrawingSheetSize : sheetName = "A2"
                Case 4 : sheetEnum = Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize : sheetName = "A4"
                Case Else : sheetEnum = Inventor.DrawingSheetSizeEnum.kA3DrawingSheetSize : sheetName = "A3"
            End Select

            Options = New V8Options() With {
                .SheetSize = sheetEnum,
                .SheetSizeName = sheetName,
                .AsmScale = 1.0 / asmScaleVal,
                .PartScale = 1.0 / partScaleVal,
                .ViewType = cboView.SelectedIndex + 1,
                .PartsPerSheet = partsVal,
                .Xulyfileloc = cboFilter.SelectedIndex + 1,
                .CreateBOM = chkBOM.Checked,
                .CreateMode = If(rdoExisting.Checked, 2, 1)
            }
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class

End Namespace