Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.RegularExpressions
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons

    Public Module Draw_Rename_Sheet

        '=============================================================
        ' DANH SÁCH NGUỒN CÓ THỂ CHỌN
        '=============================================================
        Private ReadOnly SourceNames As String() = {
            "Part Number",
            "Stock Number",
            "Description",
            "Title",
            "Revision Number",
            "Project",
            "Designer",
            "Engineer",
            "Authority",
            "Cost Center",
            "User Status",
            "Vendor",
            "Checked By",
            "Date Created",
            "Mfg Approved By",
            "Eng Approved By",
            "File Name"
        }


        '=============================================================
        ' ENTRY
        '=============================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Vui lòng mở file Drawing (.idw)!", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument = CType(app.ActiveDocument, DrawingDocument)

                '=========================================================
                ' FORM CẤU HÌNH
                '=========================================================
                Dim allSheets As Boolean = True
                Dim sourceName As String = "Part Number"
                Dim appendNumber As Boolean = True
                Dim startNum As Integer = 1
                Dim padDigits As Integer = 2
                Dim customMode As Boolean = False
                Dim customBaseName As String = ""

                If Not ShowConfigDialog(oDrawDoc,
                                        allSheets,
                                        sourceName,
                                        appendNumber,
                                        startNum,
                                        padDigits,
                                        customMode,
                                        customBaseName) Then
                    Exit Sub
                End If

                '=========================================================
                ' THU THẬP SHEET
                '=========================================================
                Dim targets As New List(Of Sheet)

                If allSheets Then
                    For Each sh As Sheet In oDrawDoc.Sheets
                        targets.Add(sh)
                    Next
                Else
                    targets.Add(oDrawDoc.ActiveSheet)
                End If

                If targets.Count = 0 Then
                    MessageBox.Show("Không có sheet nào.", "Thông báo")
                    Exit Sub
                End If

                '=========================================================
                ' XỬ LÝ
                '=========================================================
                Dim nOK As Integer = 0
                Dim nFail As Integer = 0
                Dim log As New StringBuilder()

                '=========================================================
                ' MODE ĐỔI TÊN (Property hoặc Tự đặt)
                '=========================================================
                For i As Integer = 0 To targets.Count - 1

                    Dim sh As Sheet = targets(i)

                    '----- Lấy base name -----
                    Dim baseName As String = ""

                    If customMode Then
                        ' Chế độ tự đặt — dùng tên user nhập
                        baseName = customBaseName
                    Else
                        ' Chế độ theo Property
                        baseName = GetSheetModelProp(sh, sourceName)
                    End If

                    If String.IsNullOrEmpty(baseName) Then
                        log.AppendLine("  ⏭ " & sh.Name & ": không có tên gốc")
                        nFail += 1
                        Continue For
                    End If

                    '----- Ghép hậu tố -----
                    Dim newName As String

                    If appendNumber Then
                        Dim num As Integer = startNum + i
                        Dim numText As String =
                            If(padDigits > 0,
                               num.ToString(New String("0"c, padDigits)),
                               num.ToString())

                        newName = baseName & " " & numText
                    Else
                        newName = baseName
                    End If

                    newName = SanitizeName(newName)

                    If String.IsNullOrEmpty(newName) Then
                        log.AppendLine("  ⏭ Sheet " & (i + 1) & ": tên rỗng")
                        nFail += 1
                        Continue For
                    End If

                    '----- Chống trùng tên -----
                    If IsNameTaken(oDrawDoc, sh, newName) Then
                        Dim suffix As Integer = 1
                        Dim candidate As String = newName
                        While IsNameTaken(oDrawDoc, sh, candidate) AndAlso suffix < 100
                            candidate = newName & "_" & suffix
                            suffix += 1
                        End While
                        log.AppendLine("  ⚠ Trùng → " & candidate)
                        newName = candidate
                    End If

                    '----- Đổi tên -----
                    Try
                        Dim oldName As String = sh.Name
                        sh.Name = newName
                        nOK += 1
                        log.AppendLine("  ✔ " & oldName & "  →  " & newName)
                    Catch ex As Exception
                        nFail += 1
                        log.AppendLine("  ✘ Lỗi: " & ex.Message)
                    End Try
                Next

                oDrawDoc.Update()

                '=========================================================
                ' BÁO CÁO
                '=========================================================
                Dim modeLabel As String = ""
                If customMode Then
                    modeLabel = "Tự đặt — Base: " & customBaseName
                Else
                    modeLabel = "Property: " & sourceName
                End If

                MessageBox.Show(
    "Hoàn tất!" & vbCrLf & vbCrLf &
    "Phạm vi: " & If(allSheets, "Tất cả sheet", "Sheet đang mở") & vbCrLf &
    "Chế độ : " & modeLabel & vbCrLf &
    "Hậu tố: " & If(appendNumber, "Có số", "Không") & vbCrLf &
    "Đổi tên: " & nOK & " / " & targets.Count & vbCrLf &
    "Lỗi: " & nFail,
    "Đổi tên sheet",
    MessageBoxButtons.OK,
    MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "Đổi tên sheet",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)
            End Try

        End Sub


        '=============================================================
        ' FORM CẤU HÌNH
        '=============================================================
        Private Function ShowConfigDialog(
            ByVal oDrawDoc As DrawingDocument,
            ByRef allSheets As Boolean,
            ByRef sourceName As String,
            ByRef appendNumber As Boolean,
            ByRef startNum As Integer,
            ByRef padDigits As Integer,
            ByRef customMode As Boolean,
            ByRef customBaseName As String) As Boolean

            Dim frm As New Form With {
                .Text = "Đổi tên sheet",
                .AutoScaleMode = AutoScaleMode.None,
                .AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F),
                .ClientSize = New Drw.Size(620, 720),
                .StartPosition = FormStartPosition.CenterScreen,
                .FormBorderStyle = FormBorderStyle.FixedSingle,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .ShowInTaskbar = False,
                .BackColor = Drw.Color.FromArgb(245, 245, 245),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            }

            '=========================================================
            ' HEADER
            '=========================================================
            Dim pnlHeader As New Panel With {
                .Location = New Drw.Point(0, 0),
                .Size = New Drw.Size(620, 70),
                .BackColor = Drw.Color.FromArgb(45, 100, 180)
            }
            Dim lblTitle As New Label With {
                .Text = "ĐỔI TÊN SHEET",
                .Font = New Drw.Font("Segoe UI", 14.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.White,
                .Dock = DockStyle.Fill,
                .TextAlign = Drw.ContentAlignment.MiddleCenter
            }
            Dim lblSub As New Label With {
                .Text = "Đặt tên theo Property hoặc Tự đặt",
                .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(220, 230, 245),
                .Dock = DockStyle.Bottom,
                .Height = 20,
                .TextAlign = Drw.ContentAlignment.MiddleCenter
            }
            pnlHeader.Controls.Add(lblTitle)
            pnlHeader.Controls.Add(lblSub)
            frm.Controls.Add(pnlHeader)

            '=========================================================
            ' GROUP 1: CHẾ ĐỘ (2 lựa chọn)
            '=========================================================
            Dim gbMode As New GroupBox With {
                .Text = "1. Chế độ",
                .Location = New Drw.Point(15, 80),
                .Size = New Drw.Size(590, 130),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White
            }

            Dim rbNormal As New RadioButton With {
                .Text = "Đổi tên theo Property của Model",
                .Location = New Drw.Point(20, 25),
                .Size = New Drw.Size(550, 25),
                .Checked = True,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            }
            Dim rbCustom As New RadioButton With {
                .Text = "Tự đặt tên   (nhập base name bên dưới)",
                .Location = New Drw.Point(20, 55),
                .Size = New Drw.Size(550, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            }

            Dim lblCustom As New Label With {
                .Text = "Base name:",
                .Location = New Drw.Point(40, 92),
                .Size = New Drw.Size(85, 25),
                .TextAlign = Drw.ContentAlignment.MiddleLeft,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            }
            Dim txtCustom As New System.Windows.Forms.TextBox With {
                .Text = "Sheet",
                .Location = New Drw.Point(130, 92),
                .Size = New Drw.Size(430, 25),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            }

            gbMode.Controls.Add(rbNormal)
            gbMode.Controls.Add(rbCustom)
            gbMode.Controls.Add(lblCustom)
            gbMode.Controls.Add(txtCustom)
            frm.Controls.Add(gbMode)

            '=========================================================
            ' GROUP 2: PHẠM VI
            '=========================================================
            Dim gbScope As New GroupBox With {
                .Text = "2. Phạm vi",
                .Location = New Drw.Point(15, 220),
                .Size = New Drw.Size(590, 65),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White
            }
            Dim rbAll As New RadioButton With {
                .Text = "Tất cả sheet",
                .Location = New Drw.Point(20, 25),
                .Size = New Drw.Size(200, 25),
                .Checked = True,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            }
            Dim rbActive As New RadioButton With {
                .Text = "Chỉ sheet đang mở",
                .Location = New Drw.Point(280, 25),
                .Size = New Drw.Size(250, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            }
            gbScope.Controls.Add(rbAll)
            gbScope.Controls.Add(rbActive)
            frm.Controls.Add(gbScope)

            '=========================================================
            ' GROUP 3: NGUỒN ĐẶT TÊN
            '=========================================================
            Dim gbSrc As New GroupBox With {
                .Text = "3. Nguồn đặt tên  (chỉ dùng cho chế độ Property)",
                .Location = New Drw.Point(15, 295),
                .Size = New Drw.Size(590, 220),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White
            }

            Dim lbSrc As New ListBox With {
                .Location = New Drw.Point(20, 28),
                .Size = New Drw.Size(550, 180),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point),
                .IntegralHeight = False,
                .BorderStyle = BorderStyle.FixedSingle
            }
            For Each s As String In SourceNames
                lbSrc.Items.Add(s)
            Next
            lbSrc.SelectedIndex = 0
            gbSrc.Controls.Add(lbSrc)
            frm.Controls.Add(gbSrc)

            '=========================================================
            ' GROUP 4: HẬU TỐ SỐ
            '=========================================================
            Dim gbSuffix As New GroupBox With {
                .Text = "4. Hậu tố số",
                .Location = New Drw.Point(15, 525),
                .Size = New Drw.Size(590, 125),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White
            }

            Dim rbYes As New RadioButton With {
                .Text = "Có số   (""ABC-123 01"", ""ABC-123 02"", ...)",
                .Location = New Drw.Point(20, 25),
                .Size = New Drw.Size(400, 25),
                .Checked = False,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            }
            Dim rbNo As New RadioButton With {
                .Text = "Không số   (chỉ ""ABC-123"")",
                .Location = New Drw.Point(20, 50),
                .Size = New Drw.Size(400, 25),
                  .Checked = True,            ' ← THÊM DÒNG NÀY
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            }

            Dim lblStart As New Label With {
                .Text = "Số bắt đầu:",
                .Location = New Drw.Point(20, 85),
                .Size = New Drw.Size(85, 25),
                .TextAlign = Drw.ContentAlignment.MiddleLeft
            }
            Dim txtStart As New System.Windows.Forms.TextBox With {
                .Text = "1",
                .Location = New Drw.Point(110, 85),
                .Size = New Drw.Size(60, 25)
            }

            Dim lblPad As New Label With {
                .Text = "Số chữ số đệm:",
                .Location = New Drw.Point(200, 85),
                .Size = New Drw.Size(110, 25),
                .TextAlign = Drw.ContentAlignment.MiddleLeft
            }
            Dim txtPad As New System.Windows.Forms.TextBox With {
                .Text = "2",
                .Location = New Drw.Point(315, 85),
                .Size = New Drw.Size(60, 25)
            }

            Dim lblHint As New Label With {
                .Text = "(0 = không đệm, 2 = ""01"", 3 = ""001"")",
                .Location = New Drw.Point(385, 87),
                .Size = New Drw.Size(200, 20),
                .ForeColor = Drw.Color.FromArgb(140, 140, 140),
                .Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)
            }

            gbSuffix.Controls.AddRange({rbYes, rbNo, lblStart, txtStart, lblPad, txtPad, lblHint})
            frm.Controls.Add(gbSuffix)

            '=========================================================
            ' NÚT THỰC HIỆN
            '=========================================================
            Dim btnOK As New Button()
            btnOK.Text = "THỰC HIỆN"
            btnOK.Size = New Drw.Size(160, 44)
            btnOK.Location = New Drw.Point(435, 665)
            btnOK.FlatStyle = FlatStyle.Flat
            btnOK.FlatAppearance.BorderSize = 0
            btnOK.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(60, 115, 195)
            btnOK.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(30, 80, 155)
            btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
            btnOK.ForeColor = Drw.Color.White
            btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            btnOK.Cursor = Cursors.Hand
            btnOK.UseVisualStyleBackColor = False
            frm.Controls.Add(btnOK)

            Dim btnCancel As New Button()
            btnCancel.Text = "HỦY"
            btnCancel.Size = New Drw.Size(130, 44)
            btnCancel.Location = New Drw.Point(295, 665)
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
            frm.Controls.Add(btnCancel)

            '=========================================================
            ' ENABLE / DISABLE theo chế độ
            '=========================================================
            Dim updateEnable As Action =
                Sub()
                    Dim isNormal As Boolean = rbNormal.Checked
                    Dim isCustom As Boolean = rbCustom.Checked

                    ' Nhóm 3 — chỉ dùng cho chế độ Property
                    gbSrc.Enabled = isNormal
                    lbSrc.Enabled = isNormal

                    ' Ô nhập base name — chỉ dùng cho chế độ Tự đặt
                    txtCustom.Enabled = isCustom
                    lblCustom.Enabled = isCustom

                    ' Ô số — chỉ dùng khi chọn "Có số"
                    Dim hasSuffix As Boolean = rbYes.Checked
                    txtStart.Enabled = hasSuffix
                    txtPad.Enabled = hasSuffix
                    lblStart.Enabled = hasSuffix
                    lblPad.Enabled = hasSuffix
                    lblHint.Enabled = hasSuffix
                End Sub

            AddHandler rbNormal.CheckedChanged, Sub() updateEnable()
            AddHandler rbCustom.CheckedChanged, Sub() updateEnable()
            AddHandler rbYes.CheckedChanged, Sub() updateEnable()
            AddHandler rbNo.CheckedChanged, Sub() updateEnable()
            updateEnable()

            '=========================================================
            ' OK / CANCEL — dùng biến TẠM, không capture ByRef
            '=========================================================
            Dim okClicked As Boolean = False

            Dim tmpAllSheets As Boolean = True
            Dim tmpSourceName As String = "Part Number"
            Dim tmpAppendNumber As Boolean = True
            Dim tmpStartNum As Integer = 1
            Dim tmpPadDigits As Integer = 2
            Dim tmpCustomMode As Boolean = False
            Dim tmpCustomBaseName As String = "Sheet"

            AddHandler btnOK.Click,
                Sub()
                    Dim sn As Integer = 1
                    Dim pd As Integer = 2

                    ' Validate số nếu chọn hậu tố số
                    If rbYes.Checked Then
                        If Not Integer.TryParse(txtStart.Text.Trim(), sn) Then
                            MessageBox.Show("Số bắt đầu không hợp lệ.", "Lỗi",
                                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return
                        End If
                        If Not Integer.TryParse(txtPad.Text.Trim(), pd) OrElse pd < 0 Then
                            MessageBox.Show("Số chữ số đệm không hợp lệ.", "Lỗi",
                                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return
                        End If
                    End If

                    ' Validate base name nếu chế độ Tự đặt
                    If rbCustom.Checked AndAlso String.IsNullOrWhiteSpace(txtCustom.Text) Then
                        MessageBox.Show("Base name không được để trống.", "Lỗi",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        txtCustom.Focus()
                        Return
                    End If

                    tmpAllSheets = rbAll.Checked
                    If lbSrc.SelectedItem IsNot Nothing Then
                        tmpSourceName = lbSrc.SelectedItem.ToString()
                    End If
                    tmpAppendNumber = rbYes.Checked
                    tmpStartNum = sn
                    tmpPadDigits = pd
                    tmpCustomMode = rbCustom.Checked
                    tmpCustomBaseName = txtCustom.Text.Trim()

                    okClicked = True
                    frm.Close()
                End Sub

            AddHandler btnCancel.Click,
                Sub()
                    okClicked = False
                    frm.Close()
                End Sub

            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            '=========================================================
            ' HIỆN FORM
            '=========================================================
            frm.ShowDialog()

            '=========================================================
            ' GÁN RA NGOÀI SAU KHI FORM ĐÓNG
            '=========================================================
            If okClicked Then
                allSheets = tmpAllSheets
                sourceName = tmpSourceName
                appendNumber = tmpAppendNumber
                startNum = tmpStartNum
                padDigits = tmpPadDigits
                customMode = tmpCustomMode
                customBaseName = tmpCustomBaseName
            End If

            Return okClicked

        End Function


        '=============================================================
        ' LẤY PROPERTY TỪ MODEL CỦA SHEET
        '=============================================================
        Private Function GetSheetModelProp(ByVal sh As Sheet,
                                            ByVal propName As String) As String

            Try
                If sh Is Nothing OrElse sh.DrawingViews Is Nothing Then Return ""
                If sh.DrawingViews.Count = 0 Then Return ""

                For vi As Integer = 1 To sh.DrawingViews.Count

                    Dim v As DrawingView = Nothing
                    Try
                        v = sh.DrawingViews.Item(vi)
                    Catch
                        Continue For
                    End Try

                    If v Is Nothing Then Continue For

                    Dim doc As Document = Nothing
                    Try
                        doc = v.ReferencedDocumentDescriptor.ReferencedDocument
                    Catch
                    End Try

                    If doc Is Nothing Then Continue For

                    '----- Special case: File Name -----
                    If propName = "File Name" Then
                        Try
                            Dim fullName As String =
                                v.ReferencedDocumentDescriptor.FullDocumentName
                            If Not String.IsNullOrEmpty(fullName) Then
                                Dim fn As String =
                                    System.IO.Path.GetFileNameWithoutExtension(fullName)
                                If Not String.IsNullOrEmpty(fn) Then Return fn
                            End If
                        Catch
                        End Try
                        Continue For
                    End If

                    '----- Design Tracking Properties -----
                    Try
                        Dim propSets As PropertySets = doc.PropertySets
                        Dim dtp As PropertySet = propSets.Item("Design Tracking Properties")
                        Dim p As Inventor.Property = dtp.Item(propName)

                        If p IsNot Nothing AndAlso p.Value IsNot Nothing Then
                            Dim val As String = p.Value.ToString().Trim()
                            If val <> "" Then Return val
                        End If
                    Catch
                    End Try
                Next

            Catch
            End Try

            Return ""
        End Function


        '=============================================================
        ' LÀM SẠCH TÊN
        '=============================================================
        Private Function SanitizeName(ByVal raw As String) As String
            If raw Is Nothing Then Return ""
            Dim s As String = raw.Trim()
            Dim invalid As Char() = {"\"c, "/"c, ":"c, "*"c, "?"c, """"c, "<"c, ">"c, "|"c}
            For Each ch As Char In invalid
                s = s.Replace(ch, "_"c)
            Next
            s = Regex.Replace(s, "\s+", " ").Trim()
            Return s
        End Function


        Private Function IsNameTaken(ByVal oDrawDoc As DrawingDocument,
                                      ByVal exclude As Sheet,
                                      ByVal name As String) As Boolean
            Try
                For Each sh As Sheet In oDrawDoc.Sheets
                    If sh Is exclude Then Continue For
                    If String.Equals(sh.Name, name, System.StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
            Catch
            End Try
            Return False
        End Function

    End Module

End Namespace