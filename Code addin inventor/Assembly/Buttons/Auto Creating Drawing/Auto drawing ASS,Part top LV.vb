Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Collections
Imports System.Collections.Generic

Namespace ToolInventor2025.Assembly.Buttons.AutoCreateDrawing

    Public Module AutoDrawingASSpartTopLV

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim app As Inventor.Application = g_inventorApplication
            Try
                '===== 0. KIỂM TRA ASSEMBLY =====
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Vui lòng mở Assembly (.iam) trước!", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim asmDoc As Inventor.AssemblyDocument =
                    CType(app.ActiveDocument, Inventor.AssemblyDocument)
                Dim tg As Inventor.TransientGeometry = app.TransientGeometry

                '===== FORM OPTIONS =====
                Dim opt As AssPartOptions
                Using frm As New AssPartOptionsForm()
                    If frm.ShowDialog() <> DialogResult.OK Then Exit Sub
                    opt = frm.Options
                End Using

                Dim mode As Integer = opt.Mode
                Dim srcMode As Integer = opt.SourceMode
                Dim sheetSizeEnum As Inventor.DrawingSheetSizeEnum = opt.SheetSize
                Dim userScale As Double = opt.Scale
                Dim partsPerSheet As Integer = opt.PartsPerSheet
                Dim viewType As Integer = opt.ViewType
                Dim filterSkip As Boolean = opt.FilterSkip

                '===== NGUỒN DRAWING =====
                Dim drawDoc As Inventor.DrawingDocument = Nothing
                Dim isNewDrawing As Boolean = (srcMode = 1 OrElse srcMode = 2)

                If srcMode = 3 Then
                    ' Drawing có sẵn
                    Dim oFileDlg As Inventor.FileDialog = Nothing
                    app.CreateFileDialog(oFileDlg)
                    oFileDlg.Filter = "Bản vẽ Inventor (*.idw)|*.idw"
                    oFileDlg.DialogTitle = "Chọn file bản vẽ có sẵn"
                    oFileDlg.ShowOpen()
                    If String.IsNullOrEmpty(oFileDlg.FileName) Then Exit Sub
                    Try
                        drawDoc = CType(app.Documents.Open(oFileDlg.FileName, True), Inventor.DrawingDocument)
                    Catch ex As Exception
                        MessageBox.Show("Không mở được bản vẽ:" & vbCrLf & ex.Message, "Lỗi",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try

                ElseIf srcMode = 2 Then
                    ' Template ngoài
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

                '===== 7. MODE 2 → SHEET BẢN LẮP =====
                If mode = 2 Then
                    Try
                        Dim asmSheet As Inventor.Sheet = drawDoc.Sheets.Add(sheetSizeEnum)
                        asmSheet.Size = sheetSizeEnum
                        asmSheet.Name = "Bản lắp " & asmDoc.DisplayName.Replace(".", "_")
                        ApplyBorderAndTitleBlock(drawDoc, asmSheet, sheetSizeEnum)

                        Dim cx As Double = asmSheet.Width / 2.0
                        Dim cy As Double = asmSheet.Height / 2.0
                        Dim asmView As Inventor.DrawingView = asmSheet.DrawingViews.AddBaseView(
                            asmDoc, tg.CreatePoint2d(cx, cy), userScale,
                            Inventor.ViewOrientationTypeEnum.kFrontViewOrientation,
                            Inventor.DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
                        asmView.ShowLabel = True
                        AddProjectedViews(asmSheet, asmView, viewType, cx, cy,
                                          asmSheet.Width, asmSheet.Height, tg)
                        AddPartsListSafe(drawDoc, asmSheet, asmView, asmDoc, tg, "BẢNG KÊ CỤM TỔNG")
                    Catch ex As Exception
                        MessageBox.Show("Lỗi sheet bản lắp:" & vbCrLf & ex.Message, "Cảnh báo",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Try
                End If

                '===== 8. THU THẬP TOP LEVEL =====
                Dim allDocs As New List(Of Inventor.Document)
                Dim usedPN As New Hashtable()

                For Each occ As Inventor.ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                    Try
                        Dim refDoc As Inventor.Document = occ.Definition.Document

                        If mode = 1 Then
                            If refDoc.DocumentType <> Inventor.DocumentTypeEnum.kPartDocumentObject Then Continue For
                        Else
                            If refDoc.DocumentType <> Inventor.DocumentTypeEnum.kPartDocumentObject AndAlso
                               refDoc.DocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then Continue For
                        End If

                        If filterSkip Then
                            Try
                                Dim bs As Inventor.BOMStructureEnum = occ.BOMStructure
                                If bs = Inventor.BOMStructureEnum.kPurchasedBOMStructure OrElse
                                   bs = Inventor.BOMStructureEnum.kReferenceBOMStructure OrElse
                                   bs = Inventor.BOMStructureEnum.kPhantomBOMStructure Then Continue For
                            Catch
                            End Try
                        End If

                        If AlreadyUsed(usedPN, refDoc) Then Continue For
                        allDocs.Add(refDoc)
                    Catch
                    End Try
                Next

                '===== 9. SHEET CHI TIẾT / CỤM CON =====
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

                        AddProjectedViews(sheet, baseView, viewType, xPos, yPos, xStep * 2.5, yStep * 2.5, tg)

                        If refDoc.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                            Try
                                Dim subAsm As Inventor.AssemblyDocument = CType(refDoc, Inventor.AssemblyDocument)
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

                '===== 10. XÓA SHEET TRẮNG =====
                If isNewDrawing Then DeleteBlankFirstSheets(drawDoc)

                '===== 11. HOÀN TẤT =====
                drawDoc.Update()
                Dim modeText As String = If(mode = 1, "Chỉ Part", "Bản lắp + Part + Sub-Assembly")
                MessageBox.Show("Hoàn tất!" & vbCrLf &
                                "Chế độ: " & modeText & vbCrLf &
                                "Số mục đã vẽ: " & globalCount.ToString(),
                                "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Lay Part Top Level",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        '===== HELPER GIỮ NGUYÊN =====
        Private Sub DeleteBlankFirstSheets(drawDoc As Inventor.DrawingDocument)
            Try
                For i As Integer = drawDoc.Sheets.Count To 1 Step -1
                    Try
                        Dim sh As Inventor.Sheet = drawDoc.Sheets.Item(i)
                        If sh.DrawingViews.Count = 0 Then
                            If drawDoc.Sheets.Count <= 1 Then Exit For
                            sh.Delete()
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try
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

        Private Sub AddProjectedViews(sheet As Inventor.Sheet,
                                      baseView As Inventor.DrawingView,
                                      viewType As Integer,
                                      cx As Double, cy As Double,
                                      sheetW As Double, sheetH As Double,
                                      tg As Inventor.TransientGeometry)
            Dim dx As Double = sheetW / 3.0
            Dim dy As Double = sheetH / 3.0

            If viewType = 2 OrElse viewType = 3 OrElse viewType = 4 OrElse
               viewType = 6 OrElse viewType = 7 OrElse viewType = 8 Then
                Dim rv As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
                    baseView, tg.CreatePoint2d(cx + dx, cy),
                    Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
                rv.ShowLabel = True
            End If

            If viewType = 5 OrElse viewType = 6 OrElse viewType = 7 OrElse viewType = 8 Then
                Dim lv As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
                    baseView, tg.CreatePoint2d(cx - dx, cy),
                    Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
                lv.ShowLabel = True
            End If

            If viewType = 4 OrElse viewType = 7 OrElse viewType = 8 Then
                Dim tv As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
                    baseView, tg.CreatePoint2d(cx, cy - dy),
                    Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
                tv.ShowLabel = True
            End If

            If viewType = 3 OrElse viewType = 4 OrElse viewType = 8 Then
                sheet.DrawingViews.AddProjectedView(
                    baseView, tg.CreatePoint2d(cx + dx, cy - dy),
                    Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
            End If
        End Sub

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

                Try : sourceAsm.Update2(True) : Catch : End Try
                drawDoc.Update()
                System.Windows.Forms.Application.DoEvents()

                Dim pt As Inventor.Point2d = tg.CreatePoint2d(sheet.Width * 0.98, sheet.Height * 0.98)
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

        Private Sub AddPartsListSafeAt(drawDoc As Inventor.DrawingDocument,
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


    '==========================================================
    ' CLASS OPTIONS
    '==========================================================
    Public Class AssPartOptions
        Public Mode As Integer = 2
        Public SourceMode As Integer = 1
        Public SheetSize As Inventor.DrawingSheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA3DrawingSheetSize
        Public Scale As Double = 1.0 / 5.0
        Public ScaleInput As Integer = 5
        Public PartsPerSheet As Integer = 4
        Public ViewType As Integer = 3
        Public FilterSkip As Boolean = True
    End Class


    '==========================================================
    ' FORM OPTIONS — dạng bảng
    '==========================================================
    '==========================================================
    ' FORM OPTIONS — CĂN THẲNG BẰNG TABLELAYOUTPANEL
    '==========================================================
    '==========================================================
    ' FORM OPTIONS — FIX CỨNG KÍCH THƯỚC + CĂN GIỮA MÀN HÌNH
    '==========================================================
    Public Class AssPartOptionsForm
        Inherits Form

        Public Property Options As AssPartOptions

        Private rdoMode1, rdoMode2 As RadioButton
        Private rdoSrc1, rdoSrc2, rdoSrc3 As RadioButton
        Private cboSize As ComboBox
        Private txtScale As System.Windows.Forms.TextBox
        Private txtPartsPerSheet As System.Windows.Forms.TextBox
        Private cboView As ComboBox
        Private chkFilter As CheckBox

        '=== KÍCH THƯỚC CỐ ĐỊNH — KHÔNG ĐỔI DÙ MÀN HÌNH NÀO ===
        Private Const FORM_W As Integer = 560
        Private Const FORM_H As Integer = 630

        Public Sub New()
            Me.Text = "Tùy chọn — Drawing ASS + Part"
            Me.AutoScaleMode = AutoScaleMode.None
            Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
            Me.FormBorderStyle = FormBorderStyle.FixedSingle
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
            Me.StartPosition = FormStartPosition.CenterScreen

            '=== FIX CỨNG CLIENT SIZE ===
            Me.ClientSize = New System.Drawing.Size(FORM_W, FORM_H)

            '===== HEADER =====
            Dim pnlHeader As New Panel()
            pnlHeader.Location = New System.Drawing.Point(0, 0)
            pnlHeader.Size = New System.Drawing.Size(FORM_W, 50)
            pnlHeader.BackColor = System.Drawing.Color.FromArgb(45, 100, 180)
            Me.Controls.Add(pnlHeader)

            Dim lblTitle As New Label()
            lblTitle.Text = "DRAWING ASS + PART — TÙY CHỌN"
            lblTitle.Font = New Font("Segoe UI", 11.5F, FontStyle.Bold, GraphicsUnit.Point)
            lblTitle.ForeColor = System.Drawing.Color.White
            lblTitle.Dock = DockStyle.Fill
            lblTitle.TextAlign = ContentAlignment.MiddleCenter
            pnlHeader.Controls.Add(lblTitle)

            '===== GROUP 1: CHẾ ĐỘ =====
            Dim gb1 As New GroupBox() With {
                .Text = "Chế độ",
                .Location = New System.Drawing.Point(12, 62),
                .Size = New System.Drawing.Size(FORM_W - 24, 75),
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                .BackColor = System.Drawing.Color.White}
            Me.Controls.Add(gb1)

            rdoMode1 = New RadioButton() With {
                .Text = "1 - Chỉ chi tiết (Part) cấp cao nhất",
                .Location = New System.Drawing.Point(15, 22),
                .AutoSize = True,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            rdoMode2 = New RadioButton() With {
                .Text = "2 - Bản lắp tổng + Part + Sub-Assembly cấp cao",
                .Location = New System.Drawing.Point(15, 46),
                .AutoSize = True,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                .Checked = True}
            gb1.Controls.Add(rdoMode1)
            gb1.Controls.Add(rdoMode2)

            '===== GROUP 2: NGUỒN DRAWING =====
            Dim gb2 As New GroupBox() With {
                .Text = "Nguồn Drawing",
                .Location = New System.Drawing.Point(12, 145),
                .Size = New System.Drawing.Size(FORM_W - 24, 100),
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                .BackColor = System.Drawing.Color.White}
            Me.Controls.Add(gb2)

            rdoSrc1 = New RadioButton() With {
                .Text = "1 - Tạo mới (template gốc Inventor)",
                .Location = New System.Drawing.Point(15, 22),
                .AutoSize = True,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                .Checked = True}
            rdoSrc2 = New RadioButton() With {
                .Text = "2 - Tạo mới (chọn file template / .idw ngoài)",
                .Location = New System.Drawing.Point(15, 46),
                .AutoSize = True,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            rdoSrc3 = New RadioButton() With {
                .Text = "3 - Thêm vào drawing có sẵn",
                .Location = New System.Drawing.Point(15, 70),
                .AutoSize = True,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            gb2.Controls.Add(rdoSrc1)
            gb2.Controls.Add(rdoSrc2)
            gb2.Controls.Add(rdoSrc3)

            '===== GROUP 3: KHỔ GIẤY / TỈ LỆ / SỐ PART =====
            Dim gb3 As New GroupBox() With {
                .Text = "Khổ giấy / Tỉ lệ / Số Part mỗi sheet",
                .Location = New System.Drawing.Point(12, 253),
                .Size = New System.Drawing.Size(FORM_W - 24, 145),
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                .BackColor = System.Drawing.Color.White}
            Me.Controls.Add(gb3)

            '--- Row 1: Khổ giấy ---
            Dim lblSize As New Label() With {
                .Text = "Khổ giấy:",
                .Location = New System.Drawing.Point(20, 30),
                .Size = New System.Drawing.Size(150, 25),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            gb3.Controls.Add(lblSize)

            cboSize = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Location = New System.Drawing.Point(175, 30),
                .Size = New System.Drawing.Size(100, 25)}
            cboSize.Items.AddRange(New Object() {"A0", "A1", "A2", "A3", "A4"})
            cboSize.SelectedIndex = 3
            gb3.Controls.Add(cboSize)

            '--- Row 2: Tỉ lệ ---
            Dim lblScale As New Label() With {
                .Text = "Tỉ lệ (20 = 1:20):",
                .Location = New System.Drawing.Point(20, 65),
                .Size = New System.Drawing.Size(150, 25),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            gb3.Controls.Add(lblScale)

            txtScale = New System.Windows.Forms.TextBox() With {
                .Text = "5",
                .Location = New System.Drawing.Point(175, 65),
                .Size = New System.Drawing.Size(100, 25)}
            gb3.Controls.Add(txtScale)

            '--- Row 3: Số Part / sheet ---
            Dim lblParts As New Label() With {
                .Text = "Số Part / sheet:",
                .Location = New System.Drawing.Point(20, 100),
                .Size = New System.Drawing.Size(150, 25),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            gb3.Controls.Add(lblParts)

            txtPartsPerSheet = New System.Windows.Forms.TextBox() With {
                .Text = "4",
                .Location = New System.Drawing.Point(175, 100),
                .Size = New System.Drawing.Size(100, 25)}
            gb3.Controls.Add(txtPartsPerSheet)

            '===== GROUP 4: KIỂU VIEW =====
            Dim gb4 As New GroupBox() With {
                .Text = "Kiểu view",
                .Location = New System.Drawing.Point(12, 406),
                .Size = New System.Drawing.Size(FORM_W - 24, 70),
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                .BackColor = System.Drawing.Color.White}
            Me.Controls.Add(gb4)

            Dim lblView As New Label() With {
                .Text = "Kiểu view:",
                .Location = New System.Drawing.Point(20, 28),
                .Size = New System.Drawing.Size(150, 25),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)}
            gb4.Controls.Add(lblView)

            cboView = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Location = New System.Drawing.Point(175, 28),
                .Size = New System.Drawing.Size(340, 25)}
            cboView.Items.AddRange(New Object() {
                "1 - Front",
                "2 - Front + Right",
                "3 - Front + Right + Iso",
                "4 - Front + Top + Right + Iso",
                "5 - Front + Left",
                "6 - Front + Left + Right",
                "7 - Front + Top + Left + Right",
                "8 - Front + Top + Left + Right + Iso"})
            cboView.SelectedIndex = 2
            gb4.Controls.Add(cboView)

            '===== GROUP 5: BỘ LỌC =====
            Dim gb5 As New GroupBox() With {
                .Text = "Bộ lọc",
                .Location = New System.Drawing.Point(12, 484),
                .Size = New System.Drawing.Size(FORM_W - 24, 55),
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                .BackColor = System.Drawing.Color.White}
            Me.Controls.Add(gb5)

            chkFilter = New CheckBox() With {
                .Text = "Bỏ qua Purchased / Phantom / Reference",
                .Location = New System.Drawing.Point(20, 22),
                .AutoSize = True,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                .Checked = True}
            gb5.Controls.Add(chkFilter)

            '===== NÚT =====
            Dim btnOK As New Button()
            btnOK.Text = "TẠO BẢN VẼ"
            btnOK.Size = New Size(150, 38)
            btnOK.Location = New System.Drawing.Point(FORM_W - 285, FORM_H - 52)
            btnOK.BackColor = System.Drawing.Color.FromArgb(45, 100, 180)
            btnOK.ForeColor = System.Drawing.Color.White
            btnOK.FlatStyle = FlatStyle.Flat
            btnOK.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold, GraphicsUnit.Point)
            AddHandler btnOK.Click, AddressOf OnOKClick
            Me.Controls.Add(btnOK)

            Dim btnCancel As New Button()
            btnCancel.Text = "HỦY"
            btnCancel.Size = New Size(100, 38)
            btnCancel.Location = New System.Drawing.Point(FORM_W - 125, FORM_H - 52)
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

        Private Sub OnOKClick(ByVal sender As Object, ByVal e As EventArgs)
            Dim scaleVal As Double = 5
            If Not Double.TryParse(txtScale.Text, scaleVal) OrElse scaleVal <= 0 Then
                MessageBox.Show("Tỉ lệ không hợp lệ.", "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim partsVal As Integer = 4
            If Not Integer.TryParse(txtPartsPerSheet.Text, partsVal) OrElse partsVal < 1 Then
                MessageBox.Show("Số Part/sheet không hợp lệ.", "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim sheetEnum As Inventor.DrawingSheetSizeEnum
            Select Case cboSize.SelectedIndex
                Case 0 : sheetEnum = Inventor.DrawingSheetSizeEnum.kA0DrawingSheetSize
                Case 1 : sheetEnum = Inventor.DrawingSheetSizeEnum.kA1DrawingSheetSize
                Case 2 : sheetEnum = Inventor.DrawingSheetSizeEnum.kA2DrawingSheetSize
                Case 4 : sheetEnum = Inventor.DrawingSheetSizeEnum.kA4DrawingSheetSize
                Case Else : sheetEnum = Inventor.DrawingSheetSizeEnum.kA3DrawingSheetSize
            End Select

            Options = New AssPartOptions() With {
                .Mode = If(rdoMode2.Checked, 2, 1),
                .SourceMode = If(rdoSrc1.Checked, 1, If(rdoSrc2.Checked, 2, 3)),
                .SheetSize = sheetEnum,
                .ScaleInput = CInt(scaleVal),
                .Scale = 1.0 / scaleVal,
                .PartsPerSheet = partsVal,
                .ViewType = cboView.SelectedIndex + 1,
                .FilterSkip = chkFilter.Checked
            }
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class

End Namespace