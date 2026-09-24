Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Collections
Imports System.Collections.Generic

Namespace ToolInventor2025.Assembly.Buttons.AutoCreateDrawing

    Public Module AutoDrawingASSTopLV

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim app As Inventor.Application = g_inventorApplication
            Try
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Vui lòng mở file .iam", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim asmDoc As Inventor.AssemblyDocument =
                    CType(app.ActiveDocument, Inventor.AssemblyDocument)

                '====== FORM OPTIONS ======
                Dim opt As TopLVOptions
                Using frm As New TopLVOptionsForm()
                    If frm.ShowDialog() <> DialogResult.OK Then Exit Sub
                    opt = frm.Options
                End Using

                Dim sheetSizeEnum As Inventor.DrawingSheetSizeEnum = opt.SheetSize
                Dim userScale As Double = opt.Scale
                Dim viewType As Integer = opt.ViewType
                Dim itemsPerSheet As Integer = opt.ItemsPerSheet
                Dim filterPurchased As Boolean = opt.FilterPurchased
                Dim BOMcreate As Boolean = opt.CreateBOM

                '====== CHỌN FILE BẢN VẼ ======
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
                ' 3. THU THẬP CỤM
                '=================================================
                Dim allAssemblies As New List(Of Inventor.AssemblyDocument)
                Dim usedPN As New Hashtable()
                allAssemblies.Add(asmDoc)
                MarkUsed(usedPN, asmDoc)

                For Each occ As Inventor.ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                    Try
                        Dim refDoc As Inventor.Document = occ.Definition.Document
                        If refDoc.DocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then Continue For

                        Dim Skip As Boolean = False
                        If filterPurchased Then
                            Try
                                Dim bs As Inventor.BOMStructureEnum = occ.BOMStructure
                                If bs = Inventor.BOMStructureEnum.kPurchasedBOMStructure OrElse
                                   bs = Inventor.BOMStructureEnum.kPhantomBOMStructure OrElse
                                   bs = Inventor.BOMStructureEnum.kReferenceBOMStructure Then Skip = True
                            Catch
                            End Try
                        End If
                        If Skip Then Continue For

                        Dim subAsm As Inventor.AssemblyDocument = CType(refDoc, Inventor.AssemblyDocument)
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
                ' 5. SHEET CỤM CON
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
                        Dim rightView As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
                            baseViewSub, tg.CreatePoint2d(xPos + xStep * 0.4, yPos),
                            Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseViewSub.Scale)
                        rightView.ShowLabel = True
                    End If
                    If viewType = 3 OrElse viewType = 4 Then
                        sheet.DrawingViews.AddProjectedView(baseViewSub,
                            tg.CreatePoint2d(xPos + xStep * 0.4, yPos - yStep * 0.4),
                            Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseViewSub.Scale)
                    End If
                    If viewType = 4 Then
                        Dim topView As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
                            baseViewSub, tg.CreatePoint2d(xPos, yPos - yStep * 0.4),
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

                drawDoc.Update()
                MessageBox.Show("Đã tạo bản vẽ cho cụm tổng và " & globalCount & " cụm con!",
                                "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Lay cum lap",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        '===== CÁC HÀM HELPER GIỮ NGUYÊN =====
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

        Private Sub AddProjectedViews(sheet As Inventor.Sheet,
                                      baseView As Inventor.DrawingView,
                                      viewType As Integer,
                                      cx As Double, cy As Double,
                                      sheetW As Double, sheetH As Double,
                                      tg As Inventor.TransientGeometry)
            If viewType >= 2 Then
                Dim rv As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
                    baseView, tg.CreatePoint2d(cx + sheetW / 3.0, cy),
                    Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
                rv.ShowLabel = True
            End If
            If viewType = 3 OrElse viewType = 4 Then
                sheet.DrawingViews.AddProjectedView(baseView,
                    tg.CreatePoint2d(cx + sheetW / 3.0, cy - sheetH / 3.0),
                    Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
            End If
            If viewType = 4 Then
                Dim tv As Inventor.DrawingView = sheet.DrawingViews.AddProjectedView(
                    baseView, tg.CreatePoint2d(cx, cy - sheetH / 3.0),
                    Inventor.DrawingViewStyleEnum.kFromBaseDrawingViewStyle, baseView.Scale)
                tv.ShowLabel = True
            End If
        End Sub

    End Module


    '==========================================================
    ' CLASS LƯU OPTIONS
    '==========================================================
    Public Class TopLVOptions
        Public SheetSize As Inventor.DrawingSheetSizeEnum = Inventor.DrawingSheetSizeEnum.kA3DrawingSheetSize
        Public Scale As Double = 1.0 / 20.0
        Public ScaleInput As Integer = 20
        Public ViewType As Integer = 3
        Public ItemsPerSheet As Integer = 4
        Public FilterPurchased As Boolean = True
        Public CreateBOM As Boolean = True
    End Class


    '==========================================================
    ' FORM OPTIONS — dạng bảng, không scale DPI
    '==========================================================
    Public Class TopLVOptionsForm
        Inherits Form

        Public Property Options As TopLVOptions

        Private cboSize As ComboBox
        Private txtScale As System.Windows.Forms.TextBox
        Private cboView As ComboBox
        Private txtItemsPerSheet As System.Windows.Forms.TextBox
        Private chkFilter As CheckBox
        Private chkBOM As CheckBox

        Public Sub New()
            Me.Text = "Tùy chọn — Drawing Top ASS"
            Me.ClientSize = New Size(560, 500)
            Me.FormBorderStyle = FormBorderStyle.FixedSingle
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.AutoScaleMode = AutoScaleMode.None
            Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
            Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)

            '===== HEADER =====
            Dim pnlHeader As New Panel()
            pnlHeader.Dock = DockStyle.Top
            pnlHeader.Height = 55
            pnlHeader.BackColor = System.Drawing.Color.FromArgb(45, 100, 180)
            Me.Controls.Add(pnlHeader)

            Dim lblTitle As New Label()
            lblTitle.Text = "DRAWING TOP ASS — TÙY CHỌN"
            lblTitle.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold, GraphicsUnit.Point)
            lblTitle.ForeColor = System.Drawing.Color.White
            lblTitle.Dock = DockStyle.Fill
            lblTitle.TextAlign = ContentAlignment.MiddleCenter
            pnlHeader.Controls.Add(lblTitle)

            '===== BẢNG =====
            Dim tbl As New TableLayoutPanel()
            tbl.Location = New System.Drawing.Point(20, 75)
            tbl.Size = New Size(520, 330)
            tbl.ColumnCount = 2
            tbl.RowCount = 6
            tbl.BackColor = System.Drawing.Color.White
            tbl.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 210))
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            For i As Integer = 0 To 5
                tbl.RowStyles.Add(New RowStyle(SizeType.Absolute, 52))
            Next
            Me.Controls.Add(tbl)

            ' Row 1 — Khổ giấy
            tbl.Controls.Add(MakeLabel("Khổ giấy:"), 0, 0)
            cboSize = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 100, .Anchor = AnchorStyles.Left}
            cboSize.Items.AddRange(New Object() {"A0", "A1", "A2", "A3", "A4"})
            cboSize.SelectedIndex = 3
            tbl.Controls.Add(cboSize, 1, 0)

            ' Row 2 — Tỉ lệ
            tbl.Controls.Add(MakeLabel("Tỉ lệ (20 = 1:20):"), 0, 1)
            txtScale = New System.Windows.Forms.TextBox() With {.Text = "20", .Width = 100, .Anchor = AnchorStyles.Left}
            tbl.Controls.Add(txtScale, 1, 1)

            ' Row 3 — Kiểu view
            tbl.Controls.Add(MakeLabel("Kiểu view:"), 0, 2)
            cboView = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 280, .Anchor = AnchorStyles.Left}
            cboView.Items.AddRange(New Object() {
                "1 - Front",
                "2 - Front + Right",
                "3 - Front + Right + Iso",
                "4 - Front + Top + Right + Iso"})
            cboView.SelectedIndex = 2
            tbl.Controls.Add(cboView, 1, 2)

            ' Row 4 — Số cụm/sheet
            tbl.Controls.Add(MakeLabel("Số cụm trên mỗi sheet:"), 0, 3)
            txtItemsPerSheet = New System.Windows.Forms.TextBox() With {.Text = "4", .Width = 100, .Anchor = AnchorStyles.Left}
            tbl.Controls.Add(txtItemsPerSheet, 1, 3)

            ' Row 5 — Bộ lọc
            tbl.Controls.Add(MakeLabel("Bộ lọc:"), 0, 4)
            chkFilter = New System.Windows.Forms.CheckBox() With {
                .Text = "Bỏ qua Purchased / Phantom / Reference",
                .Checked = True, .AutoSize = True, .Anchor = AnchorStyles.Left}
            tbl.Controls.Add(chkFilter, 1, 4)

            ' Row 6 — BOM
            tbl.Controls.Add(MakeLabel("Bảng kê:"), 0, 5)
            chkBOM = New CheckBox() With {
                .Text = "Tạo BOM (Parts List)",
                .Checked = True, .AutoSize = True, .Anchor = AnchorStyles.Left}
            tbl.Controls.Add(chkBOM, 1, 5)

            '===== NÚT =====
            Dim btnOK As New Button()
            btnOK.Text = "TẠO BẢN VẼ"
            btnOK.Size = New Size(150, 40)
            btnOK.Location = New System.Drawing.Point(230, 435)
            btnOK.BackColor = System.Drawing.Color.FromArgb(45, 100, 180)
            btnOK.ForeColor = System.Drawing.Color.White
            btnOK.FlatStyle = FlatStyle.Flat
            btnOK.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold, GraphicsUnit.Point)
            AddHandler btnOK.Click, AddressOf OnOKClick
            Me.Controls.Add(btnOK)

            Dim btnCancel As New Button()
            btnCancel.Text = "HỦY"
            btnCancel.Size = New Size(100, 40)
            btnCancel.Location = New System.Drawing.Point(400, 435)
            btnCancel.FlatStyle = FlatStyle.Flat
            AddHandler btnCancel.Click, Sub()
                                            Me.DialogResult = DialogResult.Cancel
                                            Me.Close()
                                        End Sub
            Me.Controls.Add(btnCancel)
            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub

        Private Function MakeLabel(ByVal text As String) As Label
            Dim lbl As New Label()
            lbl.Text = text
            lbl.Font = New Font("Segoe UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point)
            lbl.TextAlign = ContentAlignment.MiddleLeft
            lbl.Dock = DockStyle.Fill
            lbl.Padding = New Padding(8, 0, 0, 0)
            Return lbl
        End Function

        Private Sub OnOKClick(ByVal sender As Object, ByVal e As EventArgs)
            Dim scaleVal As Double = 20
            If Not Double.TryParse(txtScale.Text, scaleVal) OrElse scaleVal <= 0 Then
                MessageBox.Show("Tỉ lệ không hợp lệ.", "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim itemsVal As Integer = 4
            If Not Integer.TryParse(txtItemsPerSheet.Text, itemsVal) OrElse itemsVal < 1 Then
                MessageBox.Show("Số cụm/sheet không hợp lệ.", "Lỗi",
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

            Options = New TopLVOptions() With {
                .SheetSize = sheetEnum,
                .ScaleInput = CInt(scaleVal),
                .Scale = 1.0 / scaleVal,
                .ViewType = cboView.SelectedIndex + 1,
                .ItemsPerSheet = itemsVal,
                .FilterPurchased = chkFilter.Checked,
                .CreateBOM = chkBOM.Checked
            }
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class

End Namespace