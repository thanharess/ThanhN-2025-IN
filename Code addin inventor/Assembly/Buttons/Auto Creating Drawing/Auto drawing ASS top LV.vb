Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports System.Collections
Imports System.Collections.Generic
Namespace ToolInventor2025.Assembly.Buttons.AutoCreateDrawing
    Public Module AutoDrawingASSTopLV

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                '=================================================
                ' 0. KIỂM TRA ASSEMBLY
                '=================================================
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Vui lòng mở file .iam", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim asmDoc As Inventor.AssemblyDocument =
                    CType(app.ActiveDocument, Inventor.AssemblyDocument)

                '=================================================
                ' 1. CHỌN FILE BẢN VẼ
                '=================================================
                Dim oFileDlg As Inventor.FileDialog = Nothing
                app.CreateFileDialog(oFileDlg)
                oFileDlg.Filter = "Bản vẽ Inventor (*.idw)|*.idw"
                oFileDlg.DialogTitle = "Chọn file bản vẽ để thêm sheet"
                oFileDlg.ShowOpen()
                If String.IsNullOrEmpty(oFileDlg.FileName) Then Exit Sub

                Dim drawDoc As Inventor.DrawingDocument = Nothing
                Try
                    drawDoc = CType(app.Documents.Open(oFileDlg.FileName, True), Inventor.DrawingDocument)
                Catch ex As Exception
                    MessageBox.Show("Không mở được bản vẽ:" & vbCrLf & ex.Message, "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End Try

                Dim tg As Inventor.TransientGeometry = app.TransientGeometry

                '=================================================
                ' 2. TÙY CHỌN NGƯỜI DÙNG
                '=================================================
                Dim sheetSizeChoice As String = InputBox("Chọn khổ giấy: A2, A3 hoặc A4", "Khổ giấy", "A3")
                Dim sheetSizeEnum As Inventor.DrawingSheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA3DrawingSheetSize
                If sheetSizeChoice <> "" Then
                    Select Case sheetSizeChoice.Trim().ToUpper()
                        Case "A2" : sheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA2DrawingSheetSize
                        Case "A4" : sheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize
                        Case "A0" : sheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA0DrawingSheetSize
                        Case "A1" : sheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA1DrawingSheetSize
                    End Select
                End If

                Dim scaleInput As String = InputBox("Nhập tỉ lệ view (ví dụ 20 = 1:20):", "Tỉ lệ", "20")
                Dim userScale As Double = 1.0 / 20.0
                Try : userScale = 1.0 / CDbl(scaleInput) : Catch : End Try

                Dim viewTypeInput As String = InputBox(
                    "Chọn kiểu view (1..4):" & vbCrLf &
                    "1: Front" & vbCrLf &
                    "2: Front + Right" & vbCrLf &
                    "3: Front + Right + Iso" & vbCrLf &
                    "4: Front + Top + Right + Iso",
                    "Kiểu view", "3")
                Dim viewType As Integer = 3
                Try : viewType = CInt(viewTypeInput) : Catch : End Try
                If viewType < 1 Or viewType > 4 Then viewType = 3

                Dim itemsPerSheetInput As String = InputBox("Số cụm trên mỗi sheet:", "Chia cụm", "4")
                Dim itemsPerSheet As Integer = 4
                Try : itemsPerSheet = CInt(itemsPerSheetInput) : Catch : End Try
                If itemsPerSheet < 1 Then itemsPerSheet = 1

                Dim filterPurchased As Boolean =
                    (MessageBox.Show("Bỏ qua Purchased / Phantom / Reference?", "Bộ lọc",
                     MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes)

                Dim BOMcreate As Boolean =
                    (MessageBox.Show("Tạo BOM?", "BOM",
                     MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes)

                '=================================================
                ' 3. THU THẬP CỤM (top + sub) – LỌC TRÙNG PN
                '=================================================
                Dim allAssemblies As New List(Of Inventor.AssemblyDocument)
                Dim usedPN As New Hashtable()

                allAssemblies.Add(asmDoc)
                MarkUsed(usedPN, asmDoc)

                For Each occ As Inventor.ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                    Try
                        Dim refDoc As Inventor.Document = occ.Definition.Document
                        If refDoc.DocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                            Continue For
                        End If

                        Dim Skip As Boolean = False

                        If filterPurchased Then
                            Try
                                Dim bs As Inventor.BOMStructureEnum = occ.BOMStructure
                                If bs = Inventor.BOMStructureEnum.kPurchasedBOMStructure OrElse
                                   bs = Inventor.BOMStructureEnum.kPhantomBOMStructure OrElse
                                   bs = Inventor.BOMStructureEnum.kReferenceBOMStructure Then
                                    Skip = True
                                End If
                            Catch
                            End Try
                        End If

                        If Skip Then Continue For

                        Dim subAsm As Inventor.AssemblyDocument = CType(refDoc, Inventor.AssemblyDocument)

                        ' Lọc trùng Part Number
                        If AlreadyUsed(usedPN, subAsm) Then Continue For

                        allAssemblies.Add(subAsm)
                    Catch
                    End Try
                Next

                '=================================================
                ' 4. SHEET CỤM TỔNG
                '=================================================
                Dim asmSheet As Inventor.Sheet = drawDoc.Sheets.Add(sheetSizeEnum)
                asmSheet.Size = sheetSizeEnum
                asmSheet.Name = "Bản lắp tổng " & asmDoc.DisplayName.Replace(".", "_")
                ApplyBorderAndTitleBlock(drawDoc, asmSheet, sheetSizeEnum)

                Dim centerX As Double = asmSheet.Width / 3.0
                Dim centerY As Double = asmSheet.Height / 2.0
                Dim basePt As Inventor.Point2d = tg.CreatePoint2d(centerX, centerY)

                Dim baseView As Inventor.DrawingView = asmSheet.DrawingViews.AddBaseView(
                    asmDoc, basePt, userScale,
                    Inventor.ViewOrientationTypeEnum.kFrontViewOrientation,
                    Inventor.DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
                baseView.ShowLabel = True

                AddProjectedViews(asmSheet, baseView, viewType, centerX, centerY, asmSheet.Width, asmSheet.Height, tg)

                If BOMcreate Then
                    AddPartsListSafe(drawDoc, asmSheet, baseView, asmDoc, tg, "BẢNG KÊ CỤM TỔNG")
                End If

                '=================================================
                ' 5. SHEET CỤM CON (nhiều cụm / sheet)
                '=================================================
                Dim sheet As Inventor.Sheet = Nothing
                Dim sheetWidth, sheetHeight, usableW, usableH, xStart, yStart, xStep, yStep As Double
                Dim cols As Integer = If(itemsPerSheet = 1, 1, 2)
                Dim countOnSheet As Integer = 0
                Dim globalCount As Integer = 0

                For i As Integer = 1 To allAssemblies.Count - 1
                    Dim subAsm As Inventor.AssemblyDocument = allAssemblies(i)

                    If countOnSheet = 0 Then
                        sheet = drawDoc.Sheets.Add(sheetSizeEnum)
                        sheet.Size = sheetSizeEnum
                        sheet.Name = "Cụm lắp " & drawDoc.Sheets.Count.ToString()
                        ApplyBorderAndTitleBlock(drawDoc, sheet, sheetSizeEnum)

                        sheetWidth = sheet.Width
                        sheetHeight = sheet.Height
                        usableW = sheetWidth / 5.0 * 3.0
                        usableH = sheetHeight / 4.0 * 3.0
                        xStart = sheetWidth / 3.0
                        yStart = sheetHeight / 5.0 * 4.0

                        cols = If(itemsPerSheet = 1, 1, 2)
                        Dim rowsNeeded As Integer = CInt(Math.Ceiling(itemsPerSheet / CDbl(cols)))
                        xStep = usableW / cols
                        yStep = usableH / (rowsNeeded + 1)
                    End If

                    Dim colIndex As Integer = countOnSheet Mod cols
                    Dim rowIndex As Integer = countOnSheet \ cols
                    Dim xPos As Double = xStart + colIndex * xStep
                    Dim yPos As Double = yStart - rowIndex * yStep

                    Dim baseViewSub As Inventor.DrawingView = sheet.DrawingViews.AddBaseView(
                        subAsm, tg.CreatePoint2d(xPos, yPos), userScale,
                        Inventor.ViewOrientationTypeEnum.kFrontViewOrientation,
                        Inventor.DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
                    baseViewSub.ShowLabel = True

                    If viewType >= 2 Then
                        Dim rightPt As Inventor.Point2d = tg.CreatePoint2d(xPos + xStep * 0.4, yPos)
                        Dim rightView As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
                            baseViewSub, rightPt,
                            Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseViewSub.Scale)
                        rightView.ShowLabel = True
                    End If
                    If viewType = 3 OrElse viewType = 4 Then
                        Dim isoPt As Inventor.Point2d = tg.CreatePoint2d(xPos + xStep * 0.4, yPos - yStep * 0.4)
                        sheet.DrawingViews.AddProjectedView(baseViewSub, isoPt,
                            Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseViewSub.Scale)
                    End If
                    If viewType = 4 Then
                        Dim topPt As Inventor.Point2d = tg.CreatePoint2d(xPos, yPos - yStep * 0.4)
                        Dim topView As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
                            baseViewSub, topPt,
                            Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseViewSub.Scale)
                        topView.ShowLabel = True
                    End If

                    If BOMcreate Then
                        Try
                            AddPartsListSafe(drawDoc, sheet, baseViewSub, subAsm, tg, "BẢNG KÊ: " & subAsm.DisplayName)
                        Catch
                        End Try
                    End If

                    countOnSheet += 1
                    globalCount += 1
                    If countOnSheet >= itemsPerSheet Then countOnSheet = 0
                Next

                '=================================================
                ' 6. HOÀN TẤT
                '=================================================
                drawDoc.Update()
                MessageBox.Show(
                    "Đã tạo bản vẽ cho cụm tổng và " & globalCount.ToString() & " cụm con!",
                    "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Lay cum lap",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

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
        ' PARTS LIST
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

                Try : sourceAsm.Update2(True) : Catch : End Try
                drawDoc.Update()
                System.Windows.Forms.Application.DoEvents()

                Dim pt As Inventor.Point2d = tg.CreatePoint2d(sheet.Width * 0.95, sheet.Height * 0.95)
                Dim pl As Inventor.PartsList = Nothing

                Try
                    pl = sheet.PartsLists.Add(baseView, pt, Inventor.PartsListLevelEnum.kFirstLevelComponents)
                Catch
                    Try
                        pl = sheet.PartsLists.Add(baseView, pt, Inventor.PartsListLevelEnum.kStructuredAllLevels)
                    Catch
                        Try
                            pl = sheet.PartsLists.Add(baseView, pt, Inventor.PartsListLevelEnum.kPartsOnly)
                        Catch
                            Try : pl = sheet.PartsLists.Add(baseView, pt) : Catch : End Try
                        End Try
                    End Try
                End Try

                If pl Is Nothing Then Exit Sub

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

        Private Sub MarkUsed(used As Hashtable, doc As Inventor.Document)
            Dim key As String = GetPartNumber(doc)
            If key = "" Then
                Try : key = doc.InternalName : Catch : Exit Sub : End Try
            End If
            If Not used.ContainsKey(key) Then used.Add(key, True)
        End Sub

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

            If viewType >= 2 Then
                Dim rightPt As Inventor.Point2d = tg.CreatePoint2d(cx + sheetW / 3.0, cy)
                Dim rv As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
                    baseView, rightPt, Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
                rv.ShowLabel = True
            End If
            If viewType = 3 OrElse viewType = 4 Then
                Dim isoPt As Inventor.Point2d = tg.CreatePoint2d(cx + sheetW / 3.0, cy - sheetH / 3.0)
                sheet.DrawingViews.AddProjectedView(baseView, isoPt,
                    Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
            End If
            If viewType = 4 Then
                Dim topPt As Inventor.Point2d = tg.CreatePoint2d(cx, cy - sheetH / 3.0)
                Dim tv As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
                    baseView, topPt, Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
                tv.ShowLabel = True
            End If
        End Sub

    End Module

End Namespace

