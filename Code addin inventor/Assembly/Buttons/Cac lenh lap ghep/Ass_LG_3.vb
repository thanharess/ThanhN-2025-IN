Option Explicit On
Imports Inventor
Imports System.Windows.Forms
Imports Drw = System.Drawing

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep

    Public Module Ass_LG_3

        '=============================================================
        ' MÀU SẮC / KÍCH THƯỚC CHUẨN
        '=============================================================
        Private ReadOnly COLOR_HEADER As Drw.Color = Drw.Color.FromArgb(45, 100, 180)
        Private ReadOnly COLOR_HEADER_SUB As Drw.Color = Drw.Color.FromArgb(220, 230, 245)
        Private ReadOnly COLOR_BG As Drw.Color = Drw.Color.FromArgb(245, 245, 245)
        Private ReadOnly COLOR_BTN As Drw.Color = Drw.Color.White
        Private ReadOnly COLOR_BTN_HOVER As Drw.Color = Drw.Color.FromArgb(235, 242, 252)
        Private ReadOnly COLOR_BTN_DOWN As Drw.Color = Drw.Color.FromArgb(215, 230, 250)
        Private ReadOnly COLOR_BTN_BORDER As Drw.Color = Drw.Color.FromArgb(210, 220, 235)
        Private ReadOnly COLOR_TEXT As Drw.Color = Drw.Color.FromArgb(40, 40, 40)

        Private Const FORM_W As Integer = 540
        Private Const FORM_H As Integer = 630
        Private Const HEADER_H As Integer = 75
        Private Const PAD As Integer = 15
        Private Const BTN_H As Integer = 42
        Private Const BTN_GAP As Integer = 8


        '=============================================================
        ' ENTRY
        '=============================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application = Nothing
            Try
                invApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"),
                               Inventor.Application)
            Catch
            End Try

            If invApp Is Nothing Then
                MessageBox.Show("Không lấy được Inventor Application!", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Vui lòng mở Assembly trước!", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            Dim asmDoc As AssemblyDocument = invApp.ActiveEditDocument
            Dim oDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

            '=====================================================
            ' FORM
            '=====================================================
            Dim frm As New Form()
            frm.Text = "Ẩn Component"
            frm.AutoScaleMode = AutoScaleMode.Dpi
            frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
            frm.ClientSize = New Drw.Size(FORM_W, FORM_H)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.ShowInTaskbar = False
            frm.BackColor = COLOR_BG
            frm.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

            '═════════════════════════════════════════════════════
            ' HEADER
            '═════════════════════════════════════════════════════
            Dim pnlHeader As New Panel() With {
                .Location = New Drw.Point(0, 0),
                .Size = New Drw.Size(FORM_W, HEADER_H),
                .BackColor = COLOR_HEADER
            }
            frm.Controls.Add(pnlHeader)

            Dim lblTitle As New Label() With {
                .Text = "ẨN COMPONENT",
                .Font = New Drw.Font("Segoe UI", 14.0F, Drw.FontStyle.Bold),
                .ForeColor = Drw.Color.White,
                .Dock = DockStyle.Fill,
                .TextAlign = Drw.ContentAlignment.MiddleCenter
            }
            pnlHeader.Controls.Add(lblTitle)

            Dim lblSub As New Label() With {
                .Text = "Ẩn chi tiết theo loại — Inventor 2025",
                .Font = New Drw.Font("Segoe UI", 9.0F),
                .ForeColor = COLOR_HEADER_SUB,
                .Dock = DockStyle.Bottom,
                .Height = 20,
                .TextAlign = Drw.ContentAlignment.MiddleCenter
            }
            pnlHeader.Controls.Add(lblSub)

            '═════════════════════════════════════════════════════
            ' GROUP 1 — ẨN THEO LOẠI (6 nút, 2 cột × 3 hàng)
            '═════════════════════════════════════════════════════
            Dim gb1 As New GroupBox() With {
                .Text = "1. Ẩn theo loại",
                .Location = New Drw.Point(PAD, HEADER_H + PAD),
                .Size = New Drw.Size(FORM_W - PAD * 2, 175),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold),
                .ForeColor = COLOR_HEADER,
                .BackColor = Drw.Color.White
            }
            frm.Controls.Add(gb1)

            Dim btnColW As Integer = (gb1.Width - 45) \ 2
            Dim y1 As Integer = 30

            Dim btn1 As Button = MakeButton("Referent", 15, y1, btnColW)
            Dim btn2 As Button = MakeButton("Phantom", 15 + btnColW + 15, y1, btnColW)
            y1 += BTN_H + BTN_GAP

            Dim btn3 As Button = MakeButton("Purchased (đồ mua)", 15, y1, btnColW)
            Dim btn4 As Button = MakeButton("Weldment (cụm hàn)", 15 + btnColW + 15, y1, btnColW)
            y1 += BTN_H + BTN_GAP

            Dim btn5 As Button = MakeButton("Part", 15, y1, btnColW)
            Dim btn6 As Button = MakeButton("Sheet Metal", 15 + btnColW + 15, y1, btnColW)

            gb1.Controls.Add(btn1) : gb1.Controls.Add(btn2)
            gb1.Controls.Add(btn3) : gb1.Controls.Add(btn4)
            gb1.Controls.Add(btn5) : gb1.Controls.Add(btn6)

            '═════════════════════════════════════════════════════
            ' GROUP 2 — THAO TÁC TỔNG HỢP
            '═════════════════════════════════════════════════════
            Dim gb2Y As Integer = HEADER_H + PAD + 175 + PAD

            Dim gb2 As New GroupBox() With {
                .Text = "2. Thao tác tổng hợp",
                .Location = New Drw.Point(PAD, gb2Y),
                .Size = New Drw.Size(FORM_W - PAD * 2, 130),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold),
                .ForeColor = COLOR_HEADER,
                .BackColor = Drw.Color.White
            }
            frm.Controls.Add(gb2)

            Dim btn7 As Button = MakeButton("Ẩn TẤT CẢ các loại trên", 15, 30, gb2.Width - 30)
            btn7.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold)

            Dim btn8 As Button = MakeButton("Hiện tất cả + chỉ ẩn Referent", 15, 30 + BTN_H + BTN_GAP, gb2.Width - 30)
            btn8.BackColor = Drw.Color.FromArgb(255, 250, 220)
            btn8.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(255, 245, 200)

            gb2.Controls.Add(btn7)
            gb2.Controls.Add(btn8)

            '═════════════════════════════════════════════════════
            ' GROUP 3 — KHÔI PHỤC
            '═════════════════════════════════════════════════════
            Dim gb3Y As Integer = gb2Y + 130 + PAD

            Dim gb3 As New GroupBox() With {
                .Text = "3. Khôi phục",
                .Location = New Drw.Point(PAD, gb3Y),
                .Size = New Drw.Size(FORM_W - PAD * 2, 85),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold),
                .ForeColor = COLOR_HEADER,
                .BackColor = Drw.Color.White
            }
            frm.Controls.Add(gb3)

            Dim btn0 As Button = MakeButton("Hiện lại tất cả", 15, 28, gb3.Width - 30)
            btn0.BackColor = Drw.Color.FromArgb(220, 245, 225)
            btn0.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(200, 240, 210)
            btn0.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(180, 230, 195)

            gb3.Controls.Add(btn0)

            '═════════════════════════════════════════════════════
            ' NÚT DEBUG + ĐÓNG (dưới cùng)
            '═════════════════════════════════════════════════════
            Dim yBottom As Integer = FORM_H - 55

            Dim btnDebug As New Button()
            btnDebug.Text = "DEBUG: Xem SubType"
            btnDebug.Location = New Drw.Point(PAD, yBottom)
            btnDebug.Size = New Drw.Size(200, 36)
            btnDebug.FlatStyle = FlatStyle.Flat
            btnDebug.FlatAppearance.BorderSize = 1
            btnDebug.FlatAppearance.BorderColor = Drw.Color.FromArgb(150, 200, 220)
            btnDebug.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(225, 245, 250)
            btnDebug.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(200, 235, 245)
            btnDebug.BackColor = Drw.Color.FromArgb(240, 252, 255)
            btnDebug.ForeColor = Drw.Color.FromArgb(30, 100, 130)
            btnDebug.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular)
            btnDebug.Cursor = Cursors.Hand
            btnDebug.UseVisualStyleBackColor = False
            frm.Controls.Add(btnDebug)

            Dim btnClose As New Button()
            btnClose.Text = "ĐÓNG"
            btnClose.Location = New Drw.Point(FORM_W - PAD - 130, yBottom)
            btnClose.Size = New Drw.Size(130, 36)
            btnClose.FlatStyle = FlatStyle.Flat
            btnClose.FlatAppearance.BorderSize = 1
            btnClose.FlatAppearance.BorderColor = Drw.Color.FromArgb(200, 200, 200)
            btnClose.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(235, 235, 235)
            btnClose.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(215, 215, 215)
            btnClose.BackColor = Drw.Color.FromArgb(250, 250, 250)
            btnClose.ForeColor = Drw.Color.FromArgb(60, 60, 60)
            btnClose.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular)
            btnClose.Cursor = Cursors.Hand
            btnClose.UseVisualStyleBackColor = False
            btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
            frm.Controls.Add(btnClose)

            frm.CancelButton = btnClose
            '═════════════════════════════════════════════════════
            ' EVENTS — giữ form mở, bấm liên tục được
            '═════════════════════════════════════════════════════
            AddHandler btn1.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "Referent")
                                       asmDoc.Update2(True)
                                       MessageBox.Show("Đã ẩn " & c & " Referent", "Hoàn tất")
                                   End Sub

            AddHandler btn2.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "Phantom")
                                       asmDoc.Update2(True)
                                       MessageBox.Show("Đã ẩn " & c & " Phantom", "Hoàn tất")
                                   End Sub

            AddHandler btn3.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "Purchased")
                                       asmDoc.Update2(True)
                                       MessageBox.Show("Đã ẩn " & c & " Purchased", "Hoàn tất")
                                   End Sub

            AddHandler btn4.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "Weldment")
                                       asmDoc.Update2(True)
                                       MessageBox.Show("Đã ẩn " & c & " Weldment", "Hoàn tất")
                                   End Sub

            AddHandler btn5.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "Part")
                                       asmDoc.Update2(True)
                                       MessageBox.Show("Đã ẩn " & c & " Part", "Hoàn tất")
                                   End Sub

            AddHandler btn6.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "SheetMetal")
                                       asmDoc.Update2(True)
                                       MessageBox.Show("Đã ẩn " & c & " Sheet Metal", "Hoàn tất")
                                   End Sub

            AddHandler btn7.Click, Sub()
                                       Dim c As Integer = HideByType(oDef, "All")
                                       asmDoc.Update2(True)
                                       MessageBox.Show("Đã ẩn " & c & " component", "Hoàn tất")
                                   End Sub

            AddHandler btn8.Click, Sub()
                                       ShowAll(oDef)
                                       Dim c As Integer = HideByType(oDef, "Referent")
                                       asmDoc.Update2(True)
                                       MessageBox.Show("Đã hiện tất cả + ẩn " & c & " Referent", "Hoàn tất")
                                   End Sub

            AddHandler btn0.Click, Sub()
                                       ShowAll(oDef)
                                       asmDoc.Update2(True)
                                       MessageBox.Show("Đã hiện lại tất cả component!", "Hoàn tất")
                                   End Sub

            frm.ShowDialog()
        End Sub


        '=============================================================
        ' HELPER: tạo nút chuẩn style
        '=============================================================
        Private Function MakeButton(text As String, x As Integer, y As Integer, w As Integer) As Button
            Dim btn As New Button()
            btn.Text = text
            btn.Location = New Drw.Point(x, y)
            btn.Size = New Drw.Size(w, BTN_H)
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 1
            btn.FlatAppearance.BorderColor = COLOR_BTN_BORDER
            btn.FlatAppearance.MouseOverBackColor = COLOR_BTN_HOVER
            btn.FlatAppearance.MouseDownBackColor = COLOR_BTN_DOWN
            btn.BackColor = COLOR_BTN
            btn.ForeColor = COLOR_TEXT
            btn.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular)
            btn.Cursor = Cursors.Hand
            btn.UseVisualStyleBackColor = False

            AddHandler btn.MouseEnter, Sub() btn.FlatAppearance.BorderColor = COLOR_HEADER
            AddHandler btn.MouseLeave, Sub() btn.FlatAppearance.BorderColor = COLOR_BTN_BORDER

            Return btn
        End Function


        '=====================================================
        ' ẨN THEO LOẠI (ĐỆ QUY TOÀN BỘ)
        '=====================================================
        Private Function HideByType(ByVal oDef As AssemblyComponentDefinition, ByVal mode As String) As Integer
            Dim count As Integer = 0
            HideRecursive(oDef.Occurrences, mode, count)
            Return count
        End Function

        Private Sub HideRecursive(ByVal occs As ComponentOccurrences, ByVal mode As String, ByRef count As Integer)
            For Each occ As ComponentOccurrence In occs
                Try
                    If occ.Suppressed Then Continue For

                    Dim shouldHide As Boolean = False

                    Select Case mode
                        Case "Referent"
                            shouldHide = IsReferent(occ)
                        Case "Phantom"
                            shouldHide = IsPhantom(occ)
                        Case "Purchased"
                            shouldHide = IsPurchased(occ)
                        Case "Weldment"
                            shouldHide = IsWeldment(occ)
                        Case "Part"
                            shouldHide = IsPart(occ)
                        Case "SheetMetal"
                            shouldHide = IsSheetMetal(occ)
                        Case "All"
                            shouldHide = IsReferent(occ) OrElse IsPhantom(occ) OrElse
                                         IsPurchased(occ) OrElse IsWeldment(occ) OrElse
                                         IsPart(occ) OrElse IsSheetMetal(occ)
                    End Select

                    If shouldHide Then
                        occ.Visible = False
                        count += 1
                    End If

                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Try
                            Dim subDef As AssemblyComponentDefinition =
                                CType(occ.Definition, AssemblyComponentDefinition)
                            HideRecursive(subDef.Occurrences, mode, count)
                        Catch
                        End Try
                    End If
                Catch
                End Try
            Next
        End Sub

        '=====================================================
        ' HIỆN LẠI TẤT CẢ (ĐỆ QUY)
        '=====================================================
        Private Sub ShowAll(ByVal oDef As AssemblyComponentDefinition)
            ShowRecursive(oDef.Occurrences)
        End Sub

        Private Sub ShowRecursive(ByVal occs As ComponentOccurrences)
            For Each occ As ComponentOccurrence In occs
                Try
                    occ.Visible = True
                    If occ.Suppressed Then occ.Unsuppress()

                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Try
                            Dim subDef As AssemblyComponentDefinition =
                                CType(occ.Definition, AssemblyComponentDefinition)
                            ShowRecursive(subDef.Occurrences)
                        Catch
                        End Try
                    End If
                Catch
                End Try
            Next
        End Sub

        '=====================================================
        ' 1. REFERENT
        '=====================================================
        Private Function IsReferent(ByVal occ As ComponentOccurrence) As Boolean
            Try
                If occ.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then Return True
                If occ.IsContentMember Then Return True

                If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim pDoc As PartDocument = TryCast(occ.Definition.Document, PartDocument)
                    If pDoc IsNot Nothing AndAlso pDoc.ComponentDefinition.IsReferencePart Then
                        Return True
                    End If
                End If
            Catch
            End Try
            Return False
        End Function

        '=====================================================
        ' 2. PHANTOM
        '=====================================================
        Private Function IsPhantom(ByVal occ As ComponentOccurrence) As Boolean
            Try
                Return (occ.BOMStructure = BOMStructureEnum.kPhantomBOMStructure)
            Catch
            End Try
            Return False
        End Function

        '=====================================================
        ' 3. PURCHASED
        '=====================================================
        Private Function IsPurchased(ByVal occ As ComponentOccurrence) As Boolean
            Try
                If occ.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then Return True

                Dim doc As Document = occ.Definition.Document
                Dim designProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")

                Dim desc As String = ""
                Try
                    desc = designProps.Item("Description").Value.ToString()
                Catch
                End Try

                If LCase(desc).Contains("purchased") OrElse LCase(desc).Contains("đồ mua") Then
                    Return True
                End If
            Catch
            End Try
            Return False
        End Function

        '=====================================================
        ' 4. WELDMENT
        '=====================================================
        Private Function IsWeldment(ByVal occ As ComponentOccurrence) As Boolean
            Try
                If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim aDoc As AssemblyDocument = TryCast(occ.Definition.Document, AssemblyDocument)
                    If aDoc IsNot Nothing AndAlso aDoc.ComponentDefinition.IsWeldment Then
                        Return True
                    End If
                End If
            Catch
            End Try
            Return False
        End Function

        '=====================================================
        ' 5. PART THƯỜNG
        '=====================================================
        Private Function IsPart(ByVal occ As ComponentOccurrence) As Boolean
            If occ Is Nothing Then Return False
            Try
                If occ.Suppressed Then Return False
                If occ.DefinitionDocumentType <> DocumentTypeEnum.kPartDocumentObject Then Return False

                Dim pDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
                If pDef Is Nothing Then Return False

                Dim pDoc As PartDocument = TryCast(occ.Definition.Document, PartDocument)
                If pDoc IsNot Nothing AndAlso IsSheetMetalDocument(pDoc) Then Return False

                If SafeIsContentMember(occ) Then Return False
                If SafeIsReferencePart(pDef) Then Return False
                If IsExcludedBOM(occ) Then Return False

                Return True
            Catch
                Return False
            End Try
        End Function

        '=====================================================
        ' 6. SHEET METAL
        '=====================================================
        Private Function IsSheetMetal(ByVal occ As ComponentOccurrence) As Boolean
            If occ Is Nothing OrElse occ.Suppressed Then Return False
            Try
                If occ.DefinitionDocumentType <> DocumentTypeEnum.kPartDocumentObject Then Return False

                Dim pDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
                If pDef Is Nothing Then Return False

                If TypeOf pDef Is SheetMetalComponentDefinition Then
                    If SafeIsContentMember(occ) Then Return False
                    If IsExcludedBOM(occ) Then Return False
                    If SafeIsReferencePart(pDef) Then Return False
                    Return True
                End If

                Dim pDoc As PartDocument = TryCast(occ.Definition.Document, PartDocument)
                If pDoc IsNot Nothing AndAlso
                   String.Equals(pDoc.SubType, "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}", StringComparison.OrdinalIgnoreCase) Then
                    If SafeIsContentMember(occ) Then Return False
                    If IsExcludedBOM(occ) Then Return False
                    If SafeIsReferencePart(pDef) Then Return False
                    Return True
                End If
            Catch
            End Try
            Return False
        End Function

        '=====================================================
        ' HELPER: nhận diện Sheet Metal trực tiếp từ PartDocument
        '=====================================================
        Private Function IsSheetMetalDocument(ByVal pDoc As PartDocument) As Boolean
            If pDoc Is Nothing Then Return False

            Try
                If pDoc.ComponentDefinition IsNot Nothing AndAlso pDoc.ComponentDefinition.IsSheetMetal Then
                    Return True
                End If
            Catch
            End Try

            Try
                Dim smDef As SheetMetalComponentDefinition =
                    TryCast(pDoc.ComponentDefinition, SheetMetalComponentDefinition)
                If smDef IsNot Nothing Then Return True
            Catch
            End Try

            Try
                If String.Equals(pDoc.SubType, SHEETMETAL_SUBTYPE_GUID,
                         StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Catch
            End Try

            Return False
        End Function

        Private Const SHEETMETAL_SUBTYPE_GUID As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

        Private Function IsExcludedBOM(ByVal occ As ComponentOccurrence) As Boolean
            If occ Is Nothing Then Return False
            Try
                Select Case occ.BOMStructure
                    Case BOMStructureEnum.kReferenceBOMStructure,
                         BOMStructureEnum.kPhantomBOMStructure,
                         BOMStructureEnum.kPurchasedBOMStructure
                        Return True
                End Select
            Catch
            End Try
            Return False
        End Function

        Private Function SafeIsReferencePart(ByVal pDef As PartComponentDefinition) As Boolean
            If pDef Is Nothing Then Return False
            Try
                Return pDef.IsReferencePart
            Catch
                Return False
            End Try
        End Function

        Private Function SafeIsContentMember(ByVal occ As ComponentOccurrence) As Boolean
            If occ Is Nothing Then Return False
            Try
                Return True = occ.IsContentMember
            Catch
                Return False
            End Try
        End Function

        Private Sub DebugDumpSubTypes(ByVal occs As ComponentOccurrences,
                                      ByVal sb As System.Text.StringBuilder,
                                      ByRef cnt As Integer)
            For Each occ As ComponentOccurrence In occs
                Try
                    Dim doc As Document = occ.Definition.Document
                    Dim pDoc As PartDocument = TryCast(doc, PartDocument)
                    Dim subType As String = "(n/a)"
                    Dim isSM As String = "?"
                    If pDoc IsNot Nothing Then
                        Try : subType = pDoc.SubType : Catch : End Try
                        Try : isSM = pDoc.ComponentDefinition.IsSheetMetal.ToString() : Catch : End Try
                    End If
                    sb.AppendLine(occ.Name & " | SM=" & isSM & " | SubType=" & subType)
                    cnt += 1

                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Dim subDef As AssemblyComponentDefinition = CType(occ.Definition, AssemblyComponentDefinition)
                        DebugDumpSubTypes(subDef.Occurrences, sb, cnt)
                    End If
                Catch
                End Try
            Next
        End Sub

    End Module
End Namespace