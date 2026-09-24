Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports Microsoft.VisualBasic
Imports System.Globalization
Imports Drw = System.Drawing

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode

    Public Module Ass_Bom_4

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim targetIdx As Integer = ShowApplyTargetForm()
            If targetIdx < 0 Then Exit Sub

            Dim applyPN As Boolean = False
            Dim applySN As Boolean = False

            Select Case targetIdx
                Case 0 : applyPN = True
                Case 1 : applySN = True
                Case 2 : applyPN = True : applySN = True
            End Select


            Try
                Dim oAsm As AssemblyDocument =
                    TryCast(g_inventorApplication.ActiveDocument, AssemblyDocument)

                If oAsm Is Nothing Then
                    MessageBox.Show("Rule này chỉ chạy trong Assembly.", "BOM")
                    Exit Sub
                End If

                Dim oBOM As BOM = oAsm.ComponentDefinition.BOM
                Try : oBOM.StructuredViewEnabled = True : Catch : End Try
                Try : oBOM.StructuredViewFirstLevelOnly = False : Catch : End Try

                Dim oBOMView As BOMView = oBOM.BOMViews.Item("Structured")

                Dim countPN As Integer = 0
                Dim countSN As Integer = 0
                Dim listDocs As New List(Of Document)


                For Each row As BOMRow In oBOMView.BOMRows

                    If IsSkipped(row) Then Continue For

                    Dim refDoc As Document = Nothing
                    Try
                        refDoc = row.ComponentDefinitions.Item(1).Document
                    Catch
                        Continue For
                    End Try

                    If refDoc Is Nothing Then Continue For
                    If refDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then Continue For
                    If Not refDoc.IsModifiable Then Continue For

                    Dim partDoc As PartDocument = CType(refDoc, PartDocument)

                    Dim smDef As SheetMetalComponentDefinition =
                        TryCast(partDoc.ComponentDefinition, SheetMetalComponentDefinition)

                    If smDef Is Nothing Then Continue For


                    Dim thickMM As Double = smDef.Thickness.Value * 10.0
                    Dim thickStr As String = FormatThickness(thickMM)
                    Dim newPrefix As String = "PL" & thickStr


                    If applyPN Then
                        Dim curPN As String = GetDesignProperty(partDoc, "Part Number")
                        Dim newPN As String = SmartUpdateThickness(curPN, newPrefix, thickMM)

                        If newPN <> "" AndAlso newPN <> curPN Then
                            If SetDesignProperty(partDoc, "Part Number", newPN) Then
                                countPN += 1
                                If Not listDocs.Contains(partDoc) Then listDocs.Add(partDoc)
                            End If
                        End If
                    End If


                    If applySN Then
                        Dim curSN As String = GetDesignProperty(partDoc, "Stock Number")
                        Dim newSN As String = SmartUpdateThickness(curSN, newPrefix, thickMM)

                        If newSN <> "" AndAlso newSN <> curSN Then
                            If SetDesignProperty(partDoc, "Stock Number", newSN) Then
                                countSN += 1
                                If Not listDocs.Contains(partDoc) Then listDocs.Add(partDoc)
                            End If
                        End If
                    End If

                Next


                For Each d As Document In listDocs
                    Try
                        If d.IsModifiable Then
                            d.Update()
                            d.Save2(True)
                        End If
                    Catch
                    End Try
                Next

                Try : oBOM.Update() : Catch : End Try
                Try : oAsm.Update2(True) : Catch : End Try


                Dim msg As String =
                    "HOÀN TẤT – Sheet Metal Thickness" & vbCrLf &
                    "=================================" & vbCrLf &
                    "Part Number đã sửa  : " & countPN.ToString() & vbCrLf &
                    "Stock Number đã sửa : " & countSN.ToString()

                MessageBox.Show(msg, "PL → Part Number / Stock Number")

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '==========================================================
        ' FORM — CARD RỘNG FULL
        '==========================================================
        Private Function ShowApplyTargetForm() As Integer
            Dim result As Integer = -1

            Using frm As New Form()
                frm.Text = "Áp dụng kiểm tra / ghi chiều dày Sheet Metal"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(700, 500)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.BackColor = Drw.Color.FromArgb(245, 245, 245)
                frm.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                frm.Tag = -1

                '===== HEADER =====
                Dim pnlHeader As New Panel()
                pnlHeader.Location = New Drw.Point(0, 0)
                pnlHeader.Size = New Drw.Size(700, 80)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "PL → PART NUMBER / STOCK NUMBER"
                lblTitle.Font = New Drw.Font("Segoe UI", 14.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Kiểm tra & ghi chiều dày Sheet Metal vào property"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 22
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== HƯỚNG DẪN =====
                Dim lblHint As New Label() With {
                    .Text = "Chọn property sẽ ghi chiều dày tấm:",
                    .Location = New Drw.Point(20, 100),
                    .Size = New Drw.Size(660, 25),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(60, 60, 60)}
                frm.Controls.Add(lblHint)

                '===== 3 CARD — RỘNG FULL =====
                AddCard(frm, 0,
                        "Chỉ Part Number",
                        "Ghi chiều dày vào Part Number. Ví dụ: PL4.5x6x7",
                        130)

                AddCard(frm, 1,
                        "Chỉ Stock Number",
                        "Ghi chiều dày vào Stock Number. Giữ nguyên Part Number",
                        220)

                AddCard(frm, 2,
                        "Cả Part Number và Stock Number",
                        "Ghi chiều dày vào cả 2 property. Khuyến nghị dùng cho đầy đủ thông tin",
                        310)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(120, 42)
                btnCancel.Location = New Drw.Point(560, 430)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.FlatAppearance.BorderColor = Drw.Color.FromArgb(180, 180, 180)
                btnCancel.BackColor = Drw.Color.FromArgb(245, 245, 245)
                btnCancel.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
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


        '==========================================================
        ' THÊM CARD — RỘNG FULL (660px)
        '==========================================================
        Private Sub AddCard(ByVal frm As Form,
                            ByVal value As Integer,
                            ByVal title As String,
                            ByVal desc As String,
                            ByVal top As Integer)

            Dim cardLeft As Integer = 20
            Dim cardWidth As Integer = 660
            Dim cardHeight As Integer = 75

            Dim pnl As New Panel()
            pnl.Location = New Drw.Point(cardLeft, top)
            pnl.Size = New Drw.Size(cardWidth, cardHeight)
            pnl.BackColor = Drw.Color.White
            pnl.BorderStyle = BorderStyle.FixedSingle
            pnl.Cursor = Cursors.Hand
            frm.Controls.Add(pnl)

            '--- Tiêu đề ---
            Dim lblTitle As New Label()
            lblTitle.Text = title
            lblTitle.Font = New Drw.Font("Segoe UI", 11.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblTitle.ForeColor = Drw.Color.FromArgb(30, 30, 30)
            lblTitle.Location = New Drw.Point(25, 15)
            lblTitle.Size = New Drw.Size(cardWidth - 40, 25)
            pnl.Controls.Add(lblTitle)

            '--- Mô tả ---
            Dim lblDesc As New Label()
            lblDesc.Text = desc
            lblDesc.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            lblDesc.ForeColor = Drw.Color.FromArgb(110, 110, 110)
            lblDesc.Location = New Drw.Point(25, 42)
            lblDesc.Size = New Drw.Size(cardWidth - 40, 22)
            pnl.Controls.Add(lblDesc)

            '--- Hover ---
            Dim hoverOn As EventHandler = Sub()
                                              pnl.BackColor = Drw.Color.FromArgb(235, 242, 252)
                                          End Sub
            Dim hoverOff As EventHandler = Sub()
                                               pnl.BackColor = Drw.Color.White
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


        '==========================================================
        ' LOGIC — GIỮ NGUYÊN
        '==========================================================
        Private Function SmartUpdateThickness(current As String, newPrefix As String, realThick As Double) As String

            If current Is Nothing Then current = ""
            current = current.Trim()

            If current = "" Then
                Return newPrefix
            End If

            If Not current.StartsWith("PL", StringComparison.OrdinalIgnoreCase) Then
                Return newPrefix
            End If

            Dim afterPL As String = current.Substring(2)
            Dim oldThickStr As String = ""
            Dim rest As String = ""
            Dim i As Integer = 0

            While i < afterPL.Length AndAlso (Char.IsDigit(afterPL(i)) OrElse afterPL(i) = "."c)
                oldThickStr &= afterPL(i)
                i += 1
            End While

            If i < afterPL.Length Then
                rest = afterPL.Substring(i)
            End If

            Dim oldThick As Double = 0
            Double.TryParse(oldThickStr, NumberStyles.Any, CultureInfo.InvariantCulture, oldThick)

            If Math.Abs(oldThick - realThick) < 0.001 Then
                Return current
            Else
                Return newPrefix & rest
            End If

        End Function


        Private Function FormatThickness(value As Double) As String
            Return value.ToString("0.###", CultureInfo.InvariantCulture)
        End Function


        Private Function IsSkipped(row As BOMRow) As Boolean
            Try
                If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then Return True
                If row.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then Return True
            Catch
            End Try
            Return False
        End Function


        Private Function GetDesignProperty(doc As Document, propName As String) As String
            Try
                If doc Is Nothing Then Return ""
                Dim ps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim prop As Inventor.Property = ps.Item(propName)
                If prop Is Nothing OrElse prop.Value Is Nothing Then Return ""
                Return CStr(prop.Value).Trim()
            Catch
                Return ""
            End Try
        End Function


        Private Function SetDesignProperty(doc As Document, propName As String, value As String) As Boolean
            Try
                If doc Is Nothing OrElse Not doc.IsModifiable Then Return False

                Dim designProps As PropertySet = Nothing
                Try
                    designProps = doc.PropertySets.Item("Design Tracking Properties")
                Catch
                    Return False
                End Try

                Try
                    Dim prop As Inventor.Property = designProps.Item(propName)
                    prop.Value = value
                    Return True
                Catch
                    Try
                        designProps.Add(value, propName)
                        Return True
                    Catch
                        Return False
                    End Try
                End Try
            Catch
                Return False
            End Try
        End Function

    End Module

End Namespace