Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.DrawPartList

    Public Module Create_PartsList

        '=========================================================
        ' OPTIONS
        '=========================================================
        Private Class PartsListOptions
            Public StyleName As String = ""
            Public Position As Point2d = Nothing
            Public PositionLabel As String = ""
            Public Cancelled As Boolean = True
        End Class


        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                '=====================================================
                ' 1. KIỂM TRA DRAWING
                '=====================================================
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then

                    MessageBox.Show("Vui lòng mở file Drawing (.idw)!",
                                    "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument = CType(app.ActiveDocument, DrawingDocument)
                Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                Dim tg As TransientGeometry = app.TransientGeometry

                '=====================================================
                ' 2. LẤY DANH SÁCH STYLE
                '=====================================================
                Dim styleNames As List(Of String) = GetStyleNames(oDrawDoc)
                If styleNames Is Nothing OrElse styleNames.Count = 0 Then
                    MessageBox.Show("Không có Style Parts List nào.",
                                    "Thông báo",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                '=====================================================
                ' 3. LẤY DANH SÁCH VỊ TRÍ THEO KHỔ GIẤY
                '=====================================================
                Dim sheetSizeName As String = DetectSheetSize(oSheet)
                Dim positions As Dictionary(Of String, Point2d) = GetPositions(oSheet, tg)

                '=====================================================
                ' 4. FORM TÙY CHỌN (STYLE + VỊ TRÍ)
                '=====================================================
                Dim opt As PartsListOptions = ShowOptionsForm(styleNames, positions, sheetSizeName)
                If opt Is Nothing OrElse opt.Cancelled Then Exit Sub

                '=====================================================
                ' 5. CHỌN VIEW ASSEMBLY
                '=====================================================
                Dim pickedObj As Object = Nothing

                Try : oDrawDoc.SelectSet.Clear() : Catch : End Try

                Try
                    pickedObj = app.CommandManager.Pick(
                        SelectionFilterEnum.kDrawingViewFilter,
                        "Chọn view của Assembly cần tạo Parts List (ESC để hủy)")
                Catch
                    Exit Sub
                End Try

                If pickedObj Is Nothing Then Exit Sub

                Dim oView As DrawingView = TryCast(pickedObj, DrawingView)
                If oView Is Nothing Then
                    MessageBox.Show("Không phải DrawingView.",
                                    "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                '=====================================================
                ' 6. KIỂM TRA VIEW CÓ PHẢI ASSEMBLY
                '=====================================================
                Dim refDoc As Document = Nothing
                Try
                    refDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                Catch
                End Try

                If refDoc Is Nothing Then
                    MessageBox.Show("View không có model tham chiếu.",
                                    "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                If refDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("View này không tham chiếu Assembly." & vbCrLf & vbCrLf &
                                    "Parts List chỉ tạo được từ view của Assembly.",
                                    "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                '=====================================================
                ' 7. KIỂM TRA SHEET ACTIVE
                '=====================================================
                Try
                    oSheet.Activate()

                    If oSheet.PartsLists.Count > 0 Then
                        MessageBox.Show("Sheet này đã có Parts List." & vbCrLf & vbCrLf &
                                        "Không tạo thêm Parts List.",
                                        "Tạo Parts List",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    End If
                Catch
                End Try

                '=====================================================
                ' 8. TẠO PARTS LIST TẠI VỊ TRÍ AN TOÀN
                '=====================================================
                Dim safePos As Point2d = tg.CreatePoint2d(oSheet.Width / 2.0, oSheet.Height / 2.0)

                Dim oPL As PartsList = Nothing
                Try
                    oPL = oSheet.PartsLists.Add(oView, safePos)
                Catch ex As Exception
                    MessageBox.Show("Không tạo được Parts List:" & vbCrLf & vbCrLf &
                                    ex.Message & vbCrLf & vbCrLf &
                                    "View: " & oView.Name & vbCrLf &
                                    "Sheet: " & oSheet.Name,
                                    "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End Try

                If oPL Is Nothing Then
                    MessageBox.Show("Parts List trả về Nothing.",
                                    "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                '=====================================================
                ' 9. GÁN STYLE
                '=====================================================
                Try
                    oPL.Style = oDrawDoc.StylesManager.PartsListStyles.Item(opt.StyleName)
                Catch ex As Exception
                    MessageBox.Show("Không thể áp dụng Style:" & vbCrLf & vbCrLf &
                                    opt.StyleName & vbCrLf & vbCrLf & ex.Message,
                                    "Cảnh báo",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try

                '=====================================================
                ' 10. SORT
                '=====================================================
                Try : oPL.Sort("ITEM") : Catch : Try : oPL.Sort("Item") : Catch : End Try : End Try

                '=====================================================
                ' 11. UPDATE
                '=====================================================
                Try : oPL.Update() : Catch : End Try
                Try : oDrawDoc.Update() : Catch : End Try

                '=====================================================
                ' 12. DI CHUYỂN ĐẾN VỊ TRÍ ĐÃ CHỌN
                '=====================================================
                Try
                    oPL.Position = opt.Position
                Catch
                End Try

            Catch ex As Exception
                MessageBox.Show("Lỗi tổng:" & vbCrLf & vbCrLf & ex.Message,
                                "Tạo Parts List",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '=========================================================
        ' PHÁT HIỆN KHỔ GIẤY
        '=========================================================
        Private Function DetectSheetSize(ByVal oSheet As Sheet) As String
            Try
                Dim w As Double = Math.Round(oSheet.Width * 10.0, 0)   ' cm → mm
                Dim h As Double = Math.Round(oSheet.Height * 10.0, 0)

                Dim ww As Integer = CInt(w)
                Dim hh As Integer = CInt(h)

                Dim lo As Integer = Math.Min(ww, hh)
                Dim hi As Integer = Math.Max(ww, hh)

                Select Case hi
                    Case 1189 : Return "A0"
                    Case 841 : Return "A1"
                    Case 594 : Return "A2"
                    Case 420 : Return "A3"
                    Case 297 : Return "A4"
                End Select
            Catch
            End Try
            Return "?"
        End Function


        '=========================================================
        ' LẤY DANH SÁCH STYLE PARTS LIST
        '=========================================================
        Private Function GetStyleNames(ByVal oDrawDoc As DrawingDocument) As List(Of String)
            Dim result As New List(Of String)

            Try
                Dim styles As Inventor.PartsListStylesEnumerator = oDrawDoc.StylesManager.PartsListStyles
                If styles Is Nothing Then Return result

                For i As Integer = 1 To styles.Count
                    Try
                        Dim style As Inventor.PartsListStyle = styles.Item(i)
                        If style IsNot Nothing AndAlso Not String.IsNullOrEmpty(style.Name) Then
                            result.Add(style.Name)
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            Return result
        End Function


        '=========================================================
        ' LẤY DANH SÁCH VỊ TRÍ THEO KHỔ GIẤY
        '=========================================================
        Private Function GetPositions(ByVal oSheet As Sheet,
                                      ByVal tg As TransientGeometry) As Dictionary(Of String, Point2d)

            Dim positions As New Dictionary(Of String, Point2d)

            Try
                Dim w As Double = oSheet.Width
                Dim h As Double = oSheet.Height

                ' Tính offset so với mép
                Dim marginX As Double = w * 0.1
                Dim marginY As Double = h * 0.1

                ' Góc trên phải
                positions.Add("↗  Góc trên phải",
                              tg.CreatePoint2d(w - marginX, h - marginY))

                ' Góc trên trái
                positions.Add("↖  Góc trên trái",
                              tg.CreatePoint2d(marginX, h - marginY))

                ' Góc dưới phải
                positions.Add("↘  Góc dưới phải",
                              tg.CreatePoint2d(w - marginX, marginY))

                ' Góc dưới trái
                positions.Add("↙  Góc dưới trái",
                              tg.CreatePoint2d(marginX, marginY))

                ' Giữa bên phải
                positions.Add("→  Giữa cạnh phải",
                              tg.CreatePoint2d(w - marginX, h / 2))

                ' Giữa bên trái
                positions.Add("←  Giữa cạnh trái",
                              tg.CreatePoint2d(marginX, h / 2))

                ' Giữa trên
                positions.Add("↑  Giữa cạnh trên",
                              tg.CreatePoint2d(w / 2, h - marginY))

                ' Giữa dưới
                positions.Add("↓  Giữa cạnh dưới",
                              tg.CreatePoint2d(w / 2, marginY))

            Catch
            End Try

            Return positions
        End Function


        '=========================================================
        ' FORM TÙY CHỌN — GỘP STYLE + VỊ TRÍ
        '=========================================================
        Private Function ShowOptionsForm(ByVal styleNames As List(Of String),
                                         ByVal positions As Dictionary(Of String, Point2d),
                                         ByVal sheetSizeName As String) As PartsListOptions

            Dim opt As New PartsListOptions()
            Dim _okConfirmed As Boolean = False

            Using frm As New Form()
                frm.Text = "Tạo Parts List — Tùy chọn"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(700, 620)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.BackColor = Drw.Color.FromArgb(245, 245, 245)
                frm.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

                '===== HEADER =====
                Dim pnlHeader As New Panel()
                pnlHeader.Location = New Drw.Point(0, 0)
                pnlHeader.Size = New Drw.Size(700, 75)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "TẠO PARTS LIST"
                lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Khổ giấy phát hiện: " & sheetSizeName &
                              "   |   Chọn Style và vị trí đặt Parts List"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== GROUP 1: STYLE =====
                Dim gb1 As New GroupBox() With {
                    .Text = "1. Chọn Style Parts List",
                    .Location = New Drw.Point(15, 90),
                    .Size = New Drw.Size(670, 90),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb1)

                Dim cboStyle As New ComboBox() With {
                    .DropDownStyle = ComboBoxStyle.DropDownList,
                    .Location = New Drw.Point(25, 38),
                    .Size = New Drw.Size(620, 28),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                For Each n As String In styleNames
                    cboStyle.Items.Add(n)
                Next
                If cboStyle.Items.Count > 0 Then cboStyle.SelectedIndex = 0
                gb1.Controls.Add(cboStyle)

                '===== GROUP 2: VỊ TRÍ =====
                Dim gb2 As New GroupBox() With {
                    .Text = "2. Vị trí đặt Parts List trên Sheet",
                    .Location = New Drw.Point(15, 190),
                    .Size = New Drw.Size(670, 340),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb2)

                '--- 8 vị trí dạng lưới 2 hàng × 4 cột ---
                Dim positionKeys As New List(Of String)
                For Each key As String In positions.Keys
                    positionKeys.Add(key)
                Next

                Dim selectedPosKey As String = If(positionKeys.Count > 0, positionKeys(0), "")

                Dim btnCards(7) As Button

                Dim cardW As Integer = 150
                Dim cardH As Integer = 70
                Dim gapX As Integer = 10
                Dim gapY As Integer = 15
                Dim padX As Integer = 20
                Dim padY As Integer = 30

                For i As Integer = 0 To 7
                    If i >= positionKeys.Count Then Exit For

                    Dim col As Integer = i Mod 4
                    Dim row As Integer = i \ 4
                    Dim x As Integer = padX + col * (cardW + gapX)
                    Dim y As Integer = padY + row * (cardH + gapY)

                    Dim key As String = positionKeys(i)

                    Dim btn As New Button()
                    btn.Text = key
                    btn.Location = New Drw.Point(x, y)
                    btn.Size = New Drw.Size(cardW, cardH)
                    btn.FlatStyle = FlatStyle.Flat
                    btn.FlatAppearance.BorderSize = 1
                    btn.FlatAppearance.BorderColor = Drw.Color.FromArgb(200, 200, 200)
                    btn.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(235, 242, 252)
                    btn.BackColor = Drw.Color.White
                    btn.ForeColor = Drw.Color.FromArgb(60, 60, 60)
                    btn.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                    btn.Cursor = Cursors.Hand
                    btn.UseVisualStyleBackColor = False
                    btn.Tag = key

                    AddHandler btn.Click, Sub(sender As Object, e As EventArgs)
                                              Dim b As Button = CType(sender, Button)
                                              selectedPosKey = CStr(b.Tag)

                                              ' Reset tất cả card
                                              For Each bc As Button In btnCards
                                                  If bc IsNot Nothing Then
                                                      bc.BackColor = Drw.Color.White
                                                      bc.ForeColor = Drw.Color.FromArgb(60, 60, 60)
                                                      bc.FlatAppearance.BorderColor = Drw.Color.FromArgb(200, 200, 200)
                                                      bc.FlatAppearance.BorderSize = 1
                                                  End If
                                              Next

                                              ' Highlight card được chọn
                                              b.BackColor = Drw.Color.FromArgb(45, 100, 180)
                                              b.ForeColor = Drw.Color.White
                                              b.FlatAppearance.BorderSize = 0
                                          End Sub

                    gb2.Controls.Add(btn)
                    btnCards(i) = btn
                Next

                ' Chọn card đầu tiên
                If btnCards(0) IsNot Nothing Then
                    btnCards(0).BackColor = Drw.Color.FromArgb(45, 100, 180)
                    btnCards(0).ForeColor = Drw.Color.White
                    btnCards(0).FlatAppearance.BorderSize = 0
                End If

                '===== NÚT THỰC HIỆN =====
                Dim btnOK As New Button()
                btnOK.Text = "THỰC HIỆN"
                btnOK.Size = New Drw.Size(170, 46)
                btnOK.Location = New Drw.Point(515, 545)
                btnOK.FlatStyle = FlatStyle.Flat
                btnOK.FlatAppearance.BorderSize = 0
                btnOK.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(60, 115, 195)
                btnOK.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(30, 80, 155)
                btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
                btnOK.ForeColor = Drw.Color.White
                btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                btnOK.Cursor = Cursors.Hand
                btnOK.UseVisualStyleBackColor = False

                AddHandler btnOK.Click, Sub()
                                            If cboStyle.SelectedItem Is Nothing Then
                                                MessageBox.Show("Chưa chọn Style.",
                                                                "Lỗi",
                                                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                Return
                                            End If

                                            If String.IsNullOrEmpty(selectedPosKey) Then
                                                MessageBox.Show("Chưa chọn vị trí.",
                                                                "Lỗi",
                                                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                Return
                                            End If

                                            opt.StyleName = cboStyle.SelectedItem.ToString()
                                            opt.Position = positions(selectedPosKey)
                                            opt.PositionLabel = selectedPosKey
                                            opt.Cancelled = False
                                            _okConfirmed = True
                                            frm.DialogResult = DialogResult.OK
                                            frm.Close()
                                        End Sub
                frm.Controls.Add(btnOK)

                '===== NÚT HỦY =====
                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(130, 46)
                btnCancel.Location = New Drw.Point(375, 545)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.FlatAppearance.BorderSize = 1
                btnCancel.FlatAppearance.BorderColor = Drw.Color.FromArgb(200, 200, 200)
                btnCancel.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(235, 235, 235)
                btnCancel.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(215, 215, 215)
                btnCancel.BackColor = Drw.Color.FromArgb(250, 250, 250)
                btnCancel.ForeColor = Drw.Color.FromArgb(60, 60, 60)
                btnCancel.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                btnCancel.Cursor = Cursors.Hand
                btnCancel.UseVisualStyleBackColor = False
                AddHandler btnCancel.Click, Sub()
                                                opt.Cancelled = True
                                                frm.DialogResult = DialogResult.Cancel
                                                frm.Close()
                                            End Sub
                frm.Controls.Add(btnCancel)

                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel

                AddHandler frm.FormClosing, Sub(sender, e)
                                                If Not _okConfirmed Then opt.Cancelled = True
                                            End Sub

                frm.ShowDialog()
            End Using

            Return opt
        End Function

    End Module

End Namespace