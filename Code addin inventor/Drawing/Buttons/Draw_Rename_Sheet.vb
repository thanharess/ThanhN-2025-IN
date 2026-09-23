Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.RegularExpressions

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

                If Not ShowConfigDialog(oDrawDoc,
                                        allSheets,
                                        sourceName,
                                        appendNumber,
                                        startNum,
                                        padDigits) Then
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
                ' ĐỔI TÊN
                '=========================================================
                Dim nOK As Integer = 0
                Dim nFail As Integer = 0
                Dim log As New StringBuilder()

                For i As Integer = 0 To targets.Count - 1

                    Dim sh As Sheet = targets(i)

                    '----- Lấy base name -----
                    Dim baseName As String = GetSheetModelProp(sh, sourceName)

                    If String.IsNullOrEmpty(baseName) Then
                        log.AppendLine("  ⏭ " & sh.Name & ": không có " & sourceName)
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
                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Phạm vi: " & If(allSheets, "Tất cả sheet", "Sheet đang mở") & vbCrLf &
                    "Nguồn : " & sourceName & vbCrLf &
                    "Hậu tố: " & If(appendNumber, "Có số", "Không") & vbCrLf &
                    "Đổi tên: " & nOK & " / " & targets.Count & vbCrLf &
                    "Lỗi: " & nFail & vbCrLf & vbCrLf &
                    "--- Chi tiết ---" & vbCrLf &
                    log.ToString(),
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
            ByRef padDigits As Integer) As Boolean

            Dim frm As New Form With {
                .Text = "Đổi tên sheet",
                .ClientSize = New Size(480, 560),
                .StartPosition = FormStartPosition.CenterScreen,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .ShowInTaskbar = False
            }

            '=========================================================
            ' GROUP 1: PHẠM VI
            '=========================================================
            Dim gbScope As New GroupBox With {
                .Text = "Phạm vi",
                .Bounds = New Rectangle(12, 12, 456, 58)
            }

            Dim rbAll As New RadioButton With {
                .Text = "Tất cả sheet",
                .Bounds = New Rectangle(15, 22, 150, 22),
                .Checked = True
            }
            Dim rbActive As New RadioButton With {
                .Text = "Chỉ sheet đang mở",
                .Bounds = New Rectangle(220, 22, 220, 22)
            }
            gbScope.Controls.AddRange({rbAll, rbActive})
            frm.Controls.Add(gbScope)

            '=========================================================
            ' GROUP 2: NGUỒN ĐẶT TÊN (LIST)
            '=========================================================
            Dim gbSrc As New GroupBox With {
                .Text = "Nguồn đặt tên (chọn 1)",
                .Bounds = New Rectangle(12, 80, 456, 240)
            }

            Dim lbSrc As New ListBox With {
                .Bounds = New Rectangle(15, 22, 425, 205),
                .Font = New Font("Segoe UI", 9),
                .IntegralHeight = False
            }

            For Each s As String In SourceNames
                lbSrc.Items.Add(s)
            Next
            lbSrc.SelectedIndex = 0

            gbSrc.Controls.Add(lbSrc)
            frm.Controls.Add(gbSrc)

            '=========================================================
            ' GROUP 3: HẬU TỐ SỐ
            '=========================================================
            Dim gbSuffix As New GroupBox With {
                .Text = "Hậu tố số",
                .Bounds = New Rectangle(12, 328, 456, 150)
            }

            Dim rbYes As New RadioButton With {
                .Text = "Có số (""ABC-123 01"", ""ABC-123 02"", ...)",
                .Bounds = New Rectangle(15, 22, 350, 22),
                .Checked = True
            }
            Dim rbNo As New RadioButton With {
                .Text = "Không số (chỉ ""ABC-123"")",
                .Bounds = New Rectangle(15, 46, 350, 22)
            }

            Dim lblStart As New Label With {
                .Text = "Số bắt đầu:",
                .Bounds = New Rectangle(15, 80, 85, 22)
            }
            Dim txtStart As New System.Windows.Forms.TextBox With {
                .Text = "1",
                .Bounds = New Rectangle(105, 78, 60, 22)
            }

            Dim lblPad As New Label With {
                .Text = "Số chữ số đệm:",
                .Bounds = New Rectangle(200, 80, 100, 22)
            }
            Dim txtPad As New System.Windows.Forms.TextBox With {
                .Text = "2",
                .Bounds = New Rectangle(305, 78, 60, 22)
            }

            Dim lblHint As New Label With {
                .Text = "(0 = không đệm, 2 = ""01"", 3 = ""001"")",
                .Bounds = New Rectangle(15, 110, 400, 20),
                .ForeColor = System.Drawing.Color.Gray,
                .Font = New Font("Segoe UI", 8, FontStyle.Italic)
            }

            gbSuffix.Controls.AddRange({rbYes, rbNo, lblStart, txtStart,
                                        lblPad, txtPad, lblHint})
            frm.Controls.Add(gbSuffix)

            '=========================================================
            ' BUTTON
            '=========================================================
            Dim btnOK As New Button With {
                .Text = "Chạy",
                .Bounds = New Rectangle(255, 495, 95, 30)
            }
            Dim btnCancel As New Button With {
                .Text = "Hủy",
                .Bounds = New Rectangle(360, 495, 95, 30)
            }
            frm.Controls.AddRange({btnOK, btnCancel})

            '=========================================================
            ' ENABLE / DISABLE 2 Ô SỐ THEO RADIO
            '=========================================================
            Dim updateEnable As Action =
                Sub()
                    Dim on_ As Boolean = rbYes.Checked
                    txtStart.Enabled = on_
                    txtPad.Enabled = on_
                    lblStart.Enabled = on_
                    lblPad.Enabled = on_
                    lblHint.Enabled = on_
                End Sub

            AddHandler rbYes.CheckedChanged, Sub() updateEnable()
            AddHandler rbNo.CheckedChanged, Sub() updateEnable()
            updateEnable()

            '=========================================================
            ' BIẾN TẠM — closure, KHÔNG phải ByRef
            '=========================================================
            Dim okClicked As Boolean = False
            Dim tmpAllSheets As Boolean = True
            Dim tmpSourceName As String = "Part Number"
            Dim tmpAppendNumber As Boolean = True
            Dim tmpStartNum As Integer = 1
            Dim tmpPadDigits As Integer = 2

            AddHandler btnOK.Click,
                Sub()
                    Dim sn As Integer = 1
                    Dim pd As Integer = 2

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

                    ' Gán vào biến TẠM (không phải ByRef)
                    tmpAllSheets = rbAll.Checked
                    If lbSrc.SelectedItem IsNot Nothing Then
                        tmpSourceName = lbSrc.SelectedItem.ToString()
                    End If
                    tmpAppendNumber = rbYes.Checked
                    tmpStartNum = sn
                    tmpPadDigits = pd

                    okClicked = True
                    frm.Close()
                End Sub

            AddHandler btnCancel.Click,
                Sub()
                    okClicked = False
                    frm.Close()
                End Sub

            '=========================================================
            ' MODELESS LOOP
            '=========================================================
            frm.Show()
            Do While frm.Visible
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(15)
            Loop

            If Not okClicked Then Return False

            '=========================================================
            ' GÁN RA NGOÀI SAU KHI FORM ĐÓNG
            '=========================================================
            allSheets = tmpAllSheets
            sourceName = tmpSourceName
            appendNumber = tmpAppendNumber
            startNum = tmpStartNum
            padDigits = tmpPadDigits

            Return True

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