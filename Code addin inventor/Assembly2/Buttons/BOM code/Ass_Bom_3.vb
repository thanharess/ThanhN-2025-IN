Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor
Imports Microsoft.VisualBasic
Imports System.Globalization
Imports System.Runtime.InteropServices

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode

    Public Module Ass_Bom_3

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim choice As Integer = ShowCopyPropertyForm()
            If choice < 0 Then Exit Sub

            Dim doItem1Top As Boolean = (choice = 0 OrElse choice = 4 OrElse choice = 7)
            Dim doItem1All As Boolean = (choice = 1)
            Dim doItem0 As Boolean = (choice = 2)
            Dim doQtyTop As Boolean = (choice = 3 OrElse choice = 4 OrElse choice = 7)
            Dim doPL As Boolean = (choice = 5 OrElse choice = 7)
            Dim doSLall As Boolean = (choice = 6 OrElse choice = 7)

            Try
                Dim invApp As Inventor.Application = Nothing
                Try
                    invApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
                Catch
                    MessageBox.Show("Không lấy được Inventor.", "BOM")
                    Exit Sub
                End Try

                Dim oAsm As AssemblyDocument = TryCast(invApp.ActiveDocument, AssemblyDocument)
                If oAsm Is Nothing Then
                    MessageBox.Show("Rule này chỉ chạy trong Assembly.", "BOM")
                    Exit Sub
                End If

                Dim oBOM As BOM = oAsm.ComponentDefinition.BOM
                oBOM.StructuredViewEnabled = True
                oBOM.StructuredViewFirstLevelOnly = False
                Try : oBOM.PartsOnlyViewEnabled = True : Catch : End Try
                Try : oBOM.Update() : Catch : End Try

                Dim oBOMView As BOMView = oBOM.BOMViews.Item("Structured")

                Dim countItem1 As Integer = 0
                Dim countItem0 As Integer = 0
                Dim countStatus As Integer = 0
                Dim countSLpart As Integer = 0
                Dim countPL As Integer = 0
                Dim countSLall As Integer = 0

                Dim dictItem1 As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                Dim dictItem0 As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                If doItem1Top OrElse doItem1All OrElse doItem0 Then
                    CollectFromStructured(oBOMView.BOMRows, doItem1Top, doItem1All, doItem0,
                                          dictItem1, dictItem0, True)
                End If

                Dim writtenDocs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                WriteToAllOccurrences(oAsm.ComponentDefinition.Occurrences, dictItem1, dictItem0,
                                      writtenDocs, countItem1, countItem0)

                If doQtyTop Then
                    For Each row As BOMRow In oBOMView.BOMRows
                        If IsSkipped(row) Then Continue For
                        Dim refDoc As Document = GetDoc(row)
                        If refDoc Is Nothing Then Continue For

                        Dim qtyStr As String = GetQty(row)
                        If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            If SetStatusProperty(refDoc, qtyStr) Then countStatus += 1
                        ElseIf refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                            If SetUserProperty(refDoc, "SL part", qtyStr) Then countSLpart += 1
                        End If
                    Next
                End If

                If doPL Then
                    For Each row As BOMRow In oBOMView.BOMRows
                        If IsSkipped(row) Then Continue For
                        Dim refDoc As Document = GetDoc(row)
                        If refDoc Is Nothing Then Continue For
                        If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                            If ProcessSheetMetalPL(CType(refDoc, PartDocument)) Then countPL += 1
                        End If
                    Next
                End If

                If doSLall Then
                    Try
                        Dim partsView As BOMView = oBOM.BOMViews.Item("Parts Only")
                        For Each row As BOMRow In partsView.BOMRows
                            If IsSkipped(row) Then Continue For
                            Dim refDoc As Document = GetDoc(row)
                            If refDoc Is Nothing Then Continue For
                            If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                                Dim qtyStr As String = GetQty(row)
                                If SetUserProperty(refDoc, "SL part all", qtyStr) Then countSLall += 1
                            End If
                        Next
                    Catch
                    End Try
                End If

                Try : oBOM.Update() : Catch : End Try
                Try : oAsm.Update2(True) : Catch : End Try

                Dim msg As String = "HOÀN TẤT" & vbCrLf & "========================" & vbCrLf
                If doItem1Top OrElse doItem1All Then msg &= "item1 : " & countItem1.ToString() & vbCrLf
                If doItem0 Then msg &= "item0 : " & countItem0.ToString() & vbCrLf
                If doQtyTop Then
                    msg &= "Status (cụm) : " & countStatus.ToString() & vbCrLf
                    msg &= "SL part : " & countSLpart.ToString() & vbCrLf
                End If
                If doPL Then msg &= "PL (SheetMetal): " & countPL.ToString() & vbCrLf
                If doSLall Then msg &= "SL part all : " & countSLall.ToString() & vbCrLf

                MessageBox.Show(msg, "Assembly BOM")

            Catch ex As Exception
                MessageBox.Show("Có lỗi:" & vbCrLf & ex.Message, "Assembly BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '=========================================================
        ' FORM CHỌN — TO HƠN
        '=========================================================
        Private Function ShowCopyPropertyForm() As Integer
            Dim result As Integer = -1

            Using frm As New Form()
                frm.Text = "Copy Property — BOM"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                frm.ClientSize = New Size(920, 660)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
                frm.Font = New Font("Segoe UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point)
                frm.Tag = -1

                '===== HEADER =====
                Dim pnlHeader As New Panel()
                pnlHeader.Location = New System.Drawing.Point(0, 0)
                pnlHeader.Size = New System.Drawing.Size(920, 80)
                pnlHeader.BackColor = System.Drawing.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "COPY PROPERTY — BOM"
                lblTitle.Font = New Font("Segoe UI", 17.0F, FontStyle.Bold, GraphicsUnit.Point)
                lblTitle.ForeColor = System.Drawing.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Chọn chức năng copy / ghi property cho Part & Cụm"
                lblSub.Font = New Font("Segoe UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point)
                lblSub.ForeColor = System.Drawing.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 22
                lblSub.TextAlign = ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== NHÓM 1: ITEM NUMBER =====
                Dim gb1 As New GroupBox() With {
                    .Text = "Item Number",
                    .Location = New System.Drawing.Point(20, 100),
                    .Size = New System.Drawing.Size(880, 155),
                    .Font = New Font("Segoe UI", 10.5F, FontStyle.Bold, GraphicsUnit.Point),
                    .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                    .BackColor = System.Drawing.Color.White}
                frm.Controls.Add(gb1)

                AddCard(gb1, frm, 0, "Copy Item Number → item1",
                        "Lấy từ BOM Top-level", 15, 35, 275)
                AddCard(gb1, frm, 1, "Copy Item Number → item1",
                        "Lấy từ BOM All-level", 302, 35, 275)
                AddCard(gb1, frm, 2, "Copy Item Number → item0",
                        "Structure (có .0)", 589, 35, 275)

                '===== NHÓM 2: QUANTITY =====
                Dim gb2 As New GroupBox() With {
                    .Text = "Quantity (Số lượng)",
                    .Location = New System.Drawing.Point(20, 265),
                    .Size = New System.Drawing.Size(880, 155),
                    .Font = New Font("Segoe UI", 10.5F, FontStyle.Bold, GraphicsUnit.Point),
                    .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                    .BackColor = System.Drawing.Color.White}
                frm.Controls.Add(gb2)

                AddCard(gb2, frm, 3, "Qty cụm + SL Part",
                        "Status + SL part top-level", 15, 35, 275)
                AddCard(gb2, frm, 4, "Item Number + Qty",
                        "Gộp chức năng 1 + 2", 302, 35, 275)
                AddCard(gb2, frm, 5, "Qty Part all-level",
                        "→ SL part all", 589, 35, 275)

                '===== NHÓM 3: ĐẶC BIỆT =====
                Dim gb3 As New GroupBox() With {
                    .Text = "Đặc biệt",
                    .Location = New System.Drawing.Point(20, 430),
                    .Size = New System.Drawing.Size(880, 155),
                    .Font = New Font("Segoe UI", 10.5F, FontStyle.Bold, GraphicsUnit.Point),
                    .ForeColor = System.Drawing.Color.FromArgb(45, 100, 180),
                    .BackColor = System.Drawing.Color.White}
                frm.Controls.Add(gb3)

                AddCard(gb3, frm, 6, "Ghi / sửa Property PL",
                        "Cho Part Sheet Metal", 15, 35, 275)
                AddCard(gb3, frm, 7, "Copy TẤT CẢ",
                        "Chạy 1 + 2 + 4 + 5", 302, 35, 275)

                Dim lblAll As New Label() With {
                    .Text = "★ Chạy toàn bộ các chức năng trên" & vbCrLf &
                            "   theo thứ tự ưu tiên",
                    .Location = New System.Drawing.Point(600, 45),
                    .Size = New System.Drawing.Size(260, 60),
                    .ForeColor = System.Drawing.Color.FromArgb(140, 140, 140),
                    .Font = New Font("Segoe UI", 9.5F, FontStyle.Italic, GraphicsUnit.Point)}
                gb3.Controls.Add(lblAll)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Size(140, 45)
                btnCancel.Location = New System.Drawing.Point(760, 600)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(180, 180, 180)
                btnCancel.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
                btnCancel.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)
                AddHandler btnCancel.Click, Sub()
                                                frm.Tag = -1
                                                frm.Close()
                                            End Sub
                frm.Controls.Add(btnCancel)
                frm.CancelButton = btnCancel

                frm.ShowDialog()
                result = CInt(frm.Tag)
            End Using

            Return result
        End Function


        '=========================================================
        ' THÊM CARD — TO HƠN
        '=========================================================
        Private Sub AddCard(ByVal parent As GroupBox,
                            ByVal frm As Form,
                            ByVal value As Integer,
                            ByVal title As String,
                            ByVal desc As String,
                            ByVal left As Integer,
                            ByVal top As Integer,
                            ByVal width As Integer)

            Dim pnl As New Panel()
            pnl.Location = New System.Drawing.Point(left, top)
            pnl.Size = New System.Drawing.Size(width, 100)
            pnl.BackColor = System.Drawing.Color.FromArgb(250, 250, 250)
            pnl.BorderStyle = BorderStyle.FixedSingle
            pnl.Cursor = Cursors.Hand
            parent.Controls.Add(pnl)

            '--- Tiêu đề ---
            Dim lblTitle As New Label()
            lblTitle.Text = title
            lblTitle.Font = New Font("Segoe UI", 10.5F, FontStyle.Bold, GraphicsUnit.Point)
            lblTitle.ForeColor = System.Drawing.Color.FromArgb(30, 30, 30)
            lblTitle.Location = New System.Drawing.Point(20, 20)
            lblTitle.Size = New System.Drawing.Size(width - 30, 26)
            pnl.Controls.Add(lblTitle)

            '--- Mô tả ---
            Dim lblDesc As New Label()
            lblDesc.Text = desc
            lblDesc.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            lblDesc.ForeColor = System.Drawing.Color.FromArgb(110, 110, 110)
            lblDesc.Location = New System.Drawing.Point(20, 52)
            lblDesc.Size = New System.Drawing.Size(width - 30, 26)
            pnl.Controls.Add(lblDesc)

            '--- Hover ---
            Dim hoverOn As EventHandler = Sub()
                                              pnl.BackColor = System.Drawing.Color.FromArgb(235, 242, 252)
                                          End Sub
            Dim hoverOff As EventHandler = Sub()
                                               pnl.BackColor = System.Drawing.Color.FromArgb(250, 250, 250)
                                           End Sub

            AddHandler pnl.MouseEnter, hoverOn
            AddHandler pnl.MouseLeave, hoverOff
            AddHandler lblTitle.MouseEnter, hoverOn
            AddHandler lblTitle.MouseLeave, hoverOff
            AddHandler lblDesc.MouseEnter, hoverOn
            AddHandler lblDesc.MouseLeave, hoverOff

            '--- Click ---
            Dim clickH As EventHandler = Sub(sender, e)
                                             frm.Tag = value
                                             frm.DialogResult = DialogResult.OK
                                             frm.Close()
                                         End Sub
            AddHandler pnl.Click, clickH
            AddHandler lblTitle.Click, clickH
            AddHandler lblDesc.Click, clickH
        End Sub


        '=========================================================
        ' CÁC HÀM XỬ LÝ — GIỮ NGUYÊN
        '=========================================================
        Private Function GetDoc(row As BOMRow) As Document
            Try
                If row Is Nothing OrElse row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then Return Nothing
                Return row.ComponentDefinitions.Item(1).Document
            Catch
                Return Nothing
            End Try
        End Function

        Private Function GetPartNumberFromDoc(doc As Document) As String
            Try
                If doc Is Nothing Then Return ""
                Dim ps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim prop As Inventor.Property = ps.Item("Part Number")
                If prop Is Nothing OrElse prop.Value Is Nothing Then Return ""
                Return CStr(prop.Value).Trim()
            Catch
                Return ""
            End Try
        End Function

        Private Sub CollectFromStructured(rows As BOMRowsEnumerator,
                                          doItem1Top As Boolean,
                                          doItem1All As Boolean,
                                          doItem0 As Boolean,
                                          dictItem1 As Dictionary(Of String, String),
                                          dictItem0 As Dictionary(Of String, String),
                                          isTopLevel As Boolean)

            If rows Is Nothing Then Exit Sub

            For Each row As BOMRow In rows
                If IsSkipped(row) Then Continue For

                Dim refDoc As Document = GetDoc(row)
                If refDoc Is Nothing Then Continue For

                Dim pn As String = GetPartNumberFromDoc(refDoc)
                If String.IsNullOrEmpty(pn) Then pn = refDoc.DisplayName

                Dim structItem As String = ""
                Try : structItem = row.ItemNumber : Catch : End Try
                If String.IsNullOrEmpty(structItem) Then Continue For

                If (doItem1Top AndAlso isTopLevel) OrElse doItem1All Then
                    Dim item1Val As String = structItem
                    If item1Val.Contains(".") Then
                        item1Val = item1Val.Substring(item1Val.LastIndexOf("."c) + 1)
                    End If
                    If Not dictItem1.ContainsKey(pn) Then
                        dictItem1(pn) = item1Val
                    End If
                End If

                If doItem0 Then
                    Dim item0Val As String = structItem
                    If isTopLevel AndAlso Not structItem.Contains(".") Then
                        item0Val = structItem & ".0"
                    End If
                    If Not dictItem0.ContainsKey(pn) Then
                        dictItem0(pn) = item0Val
                    End If
                End If

                Try
                    If row.ChildRows IsNot Nothing AndAlso row.ChildRows.Count > 0 Then
                        CollectFromStructured(row.ChildRows, doItem1Top, doItem1All, doItem0,
                                              dictItem1, dictItem0, False)
                    End If
                Catch
                End Try
            Next
        End Sub

        Private Sub WriteToAllOccurrences(occs As ComponentOccurrences,
                                          dictItem1 As Dictionary(Of String, String),
                                          dictItem0 As Dictionary(Of String, String),
                                          writtenDocs As HashSet(Of String),
                                          ByRef countItem1 As Integer,
                                          ByRef countItem0 As Integer)

            If occs Is Nothing Then Exit Sub

            For Each occ As ComponentOccurrence In occs
                Try
                    Dim doc As Document = Nothing
                    Try : doc = occ.Definition.Document : Catch : Continue For : End Try
                    If doc Is Nothing Then Continue For

                    Dim key As String = doc.FullFileName
                    If writtenDocs.Contains(key) Then
                    Else
                        Dim pn As String = GetPartNumberFromDoc(doc)
                        If String.IsNullOrEmpty(pn) Then pn = doc.DisplayName

                        Dim written As Boolean = False

                        If dictItem1.ContainsKey(pn) Then
                            If SetUserProperty(doc, "item1", dictItem1(pn)) Then
                                countItem1 += 1
                                written = True
                            End If
                        End If

                        If dictItem0.ContainsKey(pn) Then
                            If SetUserProperty(doc, "item0", dictItem0(pn)) Then
                                countItem0 += 1
                                written = True
                            End If
                        End If

                        If written Then writtenDocs.Add(key)
                    End If

                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Try
                            WriteToAllOccurrences(occ.SubOccurrences, dictItem1, dictItem0,
                                                  writtenDocs, countItem1, countItem0)
                        Catch
                        End Try
                    End If
                Catch
                End Try
            Next
        End Sub

        Private Function IsSkipped(row As BOMRow) As Boolean
            Try
                If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then Return True
                If row.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then Return True
            Catch
            End Try
            Return False
        End Function

        Private Function GetQty(row As BOMRow) As String
            Try
                Return CStr(row.ItemQuantity)
            Catch
                Try
                    Return CStr(row.TotalQuantity)
                Catch
                    Return "1"
                End Try
            End Try
        End Function

        Private Function ProcessSheetMetalPL(partDoc As PartDocument) As Boolean
            Try
                If partDoc Is Nothing OrElse Not partDoc.IsModifiable Then Return False
                Dim smDef As SheetMetalComponentDefinition =
                    TryCast(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
                If smDef Is Nothing Then Return False

                Dim thickMM As Double = smDef.Thickness.Value * 10.0
                Dim thickStr As String = FormatThickness(thickMM)
                Dim newPrefix As String = "t" & thickStr

                Dim current As String = GetUserProperty(partDoc, "PL")
                Dim finalValue As String = ""

                If String.IsNullOrEmpty(current) Then
                    finalValue = newPrefix
                Else
                    current = current.Trim()
                    If current.StartsWith("t", StringComparison.OrdinalIgnoreCase) Then
                        Dim afterPL As String = current.Substring(1)
                        Dim oldThickStr As String = ""
                        Dim rest As String = ""
                        Dim i As Integer = 0
                        While i < afterPL.Length AndAlso (Char.IsDigit(afterPL(i)) OrElse afterPL(i) = "."c)
                            oldThickStr &= afterPL(i)
                            i += 1
                        End While
                        If i < afterPL.Length Then rest = afterPL.Substring(i)

                        Dim oldThick As Double = 0
                        Double.TryParse(oldThickStr, NumberStyles.Any, CultureInfo.InvariantCulture, oldThick)

                        If Math.Abs(oldThick - thickMM) < 0.001 Then
                            Return False
                        Else
                            finalValue = newPrefix & rest
                        End If
                    Else
                        finalValue = newPrefix
                    End If
                End If

                Return SetUserProperty(partDoc, "PL", finalValue)
            Catch
                Return False
            End Try
        End Function

        Private Function FormatThickness(value As Double) As String
            Return value.ToString("0.###", CultureInfo.InvariantCulture)
        End Function

        Private Function GetUserProperty(doc As Document, propName As String) As String
            Try
                Dim userProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
                For Each p As Inventor.Property In userProps
                    If String.Equals(p.Name, propName, StringComparison.OrdinalIgnoreCase) Then
                        If p.Value Is Nothing Then Return ""
                        Return CStr(p.Value).Trim()
                    End If
                Next
            Catch
            End Try
            Return ""
        End Function

        Private Function SetUserProperty(doc As Document, propName As String, value As String) As Boolean
            Try
                If doc Is Nothing OrElse Not doc.IsModifiable Then Return False
                Dim userProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
                Dim found As Inventor.Property = Nothing
                For Each p As Inventor.Property In userProps
                    If String.Equals(p.Name, propName, StringComparison.OrdinalIgnoreCase) Then
                        found = p
                        Exit For
                    End If
                Next
                If found Is Nothing Then
                    userProps.Add(value, propName)
                Else
                    found.Value = value
                End If
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Function SetStatusProperty(doc As Document, value As String) As Boolean
            Try
                If doc Is Nothing OrElse Not doc.IsModifiable Then Return False
                Try
                    Dim designProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                    Dim prop As Inventor.Property = designProps.Item("Status")
                    prop.Value = value
                    Return True
                Catch
                End Try
                Return SetUserProperty(doc, "Status", value)
            Catch
                Return False
            End Try
        End Function

    End Module

End Namespace