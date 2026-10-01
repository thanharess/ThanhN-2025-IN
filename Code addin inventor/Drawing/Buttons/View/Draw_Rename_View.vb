Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.DrawView

    Public Module Draw_Rename_View

        ' ==================== CACHE FORM ====================
        Private _frm As Form = Nothing
        Private _txtLetter As System.Windows.Forms.TextBox
        Private _cboScope As ComboBox
        Private _chkResetPerSheet As CheckBox
        Private _chkBase As CheckBox
        Private _chkProjected As CheckBox
        Private _chkSection As CheckBox
        Private _chkDetail As CheckBox
        Private _chkAuxiliary As CheckBox
        Private _chkOverlay As CheckBox

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                Dim oDrawDoc As DrawingDocument = TryCast(g_inventorApplication.ActiveDocument, DrawingDocument)
                If oDrawDoc Is Nothing Then
                    MessageBox.Show("Document hiện tại không phải Drawing.", "Rename View Label",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim startLetter As String = ""
                Dim scope As String = ""
                Dim skipBase As Boolean = True
                Dim skipProjected As Boolean = False
                Dim skipSection As Boolean = False
                Dim skipDetail As Boolean = False
                Dim skipAuxiliary As Boolean = False
                Dim skipOverlay As Boolean = True
                Dim resetPerSheet As Boolean = False

                If Not ShowInputDialog(startLetter, scope, skipBase, skipProjected, skipSection,
                                       skipDetail, skipAuxiliary, skipOverlay, resetPerSheet) Then Return

                startLetter = startLetter.Trim().ToUpper()
                If String.IsNullOrEmpty(startLetter) OrElse Not IsValidLetter(startLetter) Then
                    MessageBox.Show("Chữ cái bắt đầu không hợp lệ (chỉ dùng A-Z, AA, AB...).",
                                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                ' Tắt update tạm thời để tăng tốc
                Dim oldSilent As Boolean = g_inventorApplication.SilentOperation
                g_inventorApplication.SilentOperation = True

                Dim count As Integer = 0

                If scope = "Active Sheet" Then
                    count = RenameViewsOnSheet(oDrawDoc.ActiveSheet, startLetter,
                                               skipBase, skipProjected, skipSection,
                                               skipDetail, skipAuxiliary, skipOverlay)
                Else
                    count = RenameViewsAllSheets(oDrawDoc, startLetter,
                                                 skipBase, skipProjected, skipSection,
                                                 skipDetail, skipAuxiliary, skipOverlay,
                                                 resetPerSheet)
                End If

                g_inventorApplication.SilentOperation = oldSilent
                oDrawDoc.Update2(True)

                Dim msg As String = "Đã đặt lại tên " & count.ToString() & " View(s)" & vbCrLf &
                                    "Bắt đầu từ: " & startLetter & vbCrLf
                If skipBase Then msg &= "• Bỏ qua Base View" & vbCrLf
                If skipProjected Then msg &= "• Bỏ qua Projected View" & vbCrLf
                If skipSection Then msg &= "• Bỏ qua Section View" & vbCrLf
                If skipDetail Then msg &= "• Bỏ qua Detail View" & vbCrLf
                If skipAuxiliary Then msg &= "• Bỏ qua Auxiliary View" & vbCrLf
                If skipOverlay Then msg &= "• Bỏ qua Overlay View" & vbCrLf
                If scope = "All Sheets" AndAlso resetPerSheet Then
                    msg &= "• Reset chữ cái mỗi Sheet" & vbCrLf
                End If

                MessageBox.Show(msg, "Rename View Label", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message, "Rename View Label",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        '==========================================================
        ' RENAME 1 SHEET – indexed loop
        '==========================================================
        Private Function RenameViewsOnSheet(ByVal oSheet As Sheet,
                                            ByVal startLetter As String,
                                            ByVal skipBase As Boolean,
                                            ByVal skipProjected As Boolean,
                                            ByVal skipSection As Boolean,
                                            ByVal skipDetail As Boolean,
                                            ByVal skipAuxiliary As Boolean,
                                            ByVal skipOverlay As Boolean) As Integer

            If oSheet Is Nothing Then Return 0

            Dim currentName As String = startLetter
            Dim count As Integer = 0
            Dim views As DrawingViews = oSheet.DrawingViews
            Dim n As Integer = views.Count

            For i As Integer = 1 To n
                Dim oView As DrawingView = views.Item(i)
                Try
                    If ShouldSkipView(oView.ViewType, skipBase, skipProjected, skipSection,
                                      skipDetail, skipAuxiliary, skipOverlay) Then Continue For

                    oView.Name = currentName
                    currentName = GetNextLetter(currentName)
                    count += 1
                Catch
                End Try
            Next

            Return count
        End Function

        '==========================================================
        ' RENAME TẤT CẢ SHEET
        '==========================================================
        Private Function RenameViewsAllSheets(ByVal oDrawDoc As DrawingDocument,
                                              ByVal startLetter As String,
                                              ByVal skipBase As Boolean,
                                              ByVal skipProjected As Boolean,
                                              ByVal skipSection As Boolean,
                                              ByVal skipDetail As Boolean,
                                              ByVal skipAuxiliary As Boolean,
                                              ByVal skipOverlay As Boolean,
                                              ByVal resetPerSheet As Boolean) As Integer

            Dim total As Integer = 0
            Dim currentName As String = startLetter
            Dim sheets As Sheets = oDrawDoc.Sheets
            Dim sheetCount As Integer = sheets.Count

            For s As Integer = 1 To sheetCount
                Dim oSheet As Sheet = sheets.Item(s)

                If resetPerSheet Then currentName = startLetter

                Dim views As DrawingViews = oSheet.DrawingViews
                Dim n As Integer = views.Count

                For i As Integer = 1 To n
                    Dim oView As DrawingView = views.Item(i)
                    Try
                        If ShouldSkipView(oView.ViewType, skipBase, skipProjected, skipSection,
                                          skipDetail, skipAuxiliary, skipOverlay) Then Continue For

                        oView.Name = currentName
                        currentName = GetNextLetter(currentName)
                        total += 1
                    Catch
                    End Try
                Next
            Next

            Return total
        End Function

        '==========================================================
        ' KIỂM TRA BỎ QUA – nhận ViewType
        '==========================================================
        Private Function ShouldSkipView(ByVal viewType As DrawingViewTypeEnum,
                                        ByVal skipBase As Boolean,
                                        ByVal skipProjected As Boolean,
                                        ByVal skipSection As Boolean,
                                        ByVal skipDetail As Boolean,
                                        ByVal skipAuxiliary As Boolean,
                                        ByVal skipOverlay As Boolean) As Boolean

            Select Case viewType
                Case DrawingViewTypeEnum.kStandardDrawingViewType : Return skipBase
                Case DrawingViewTypeEnum.kProjectedDrawingViewType : Return skipProjected
                Case DrawingViewTypeEnum.kSectionDrawingViewType : Return skipSection
                Case DrawingViewTypeEnum.kDetailDrawingViewType : Return skipDetail
                Case DrawingViewTypeEnum.kAuxiliaryDrawingViewType : Return skipAuxiliary
                Case DrawingViewTypeEnum.kOverlayDrawingViewType : Return skipOverlay
                Case Else : Return False
            End Select
        End Function

        '==========================================================
        ' TĂNG CHỮ CÁI
        '==========================================================
        Private Function GetNextLetter(ByVal current As String) As String
            If String.IsNullOrEmpty(current) Then Return "A"

            Dim chars() As Char = current.ToUpper().ToCharArray()
            Dim i As Integer = chars.Length - 1

            Do While i >= 0
                If chars(i) < "Z"c Then
                    chars(i) = Chr(Asc(chars(i)) + 1)
                    Return New String(chars)
                Else
                    chars(i) = "A"c
                    i -= 1
                End If
            Loop

            Return "A" & New String(chars)
        End Function

        Private Function IsValidLetter(ByVal text As String) As Boolean
            For Each c As Char In text
                If Not Char.IsLetter(c) Then Return False
            Next
            Return True
        End Function

        '==========================================================
        ' FORM – tạo 1 lần, reuse
        '==========================================================
        Private Function ShowInputDialog(ByRef startLetter As String,
                                         ByRef scope As String,
                                         ByRef skipBase As Boolean,
                                         ByRef skipProjected As Boolean,
                                         ByRef skipSection As Boolean,
                                         ByRef skipDetail As Boolean,
                                         ByRef skipAuxiliary As Boolean,
                                         ByRef skipOverlay As Boolean,
                                         ByRef resetPerSheet As Boolean) As Boolean

            If _frm Is Nothing Then
                _frm = New Form()
                _frm.Text = "Đặt lại tên View theo chữ cái"
                _frm.Size = New Size(400, 440)
                _frm.StartPosition = FormStartPosition.CenterScreen
                _frm.FormBorderStyle = FormBorderStyle.FixedDialog
                _frm.MaximizeBox = False
                _frm.MinimizeBox = False
                _frm.Font = New Font("Segoe UI", 9)

                Dim lbl1 As New Label() With {.Text = "Chữ cái bắt đầu (A, B, AA, AB...):", .Location = New System.Drawing.Point(20, 10), .AutoSize = True}
                _txtLetter = New System.Windows.Forms.TextBox() With {.Text = "A", .Location = New System.Drawing.Point(20, 40), .Width = 80, .CharacterCasing = CharacterCasing.Upper}

                Dim lbl2 As New Label() With {.Text = "Phạm vi:", .Location = New System.Drawing.Point(20, 70), .AutoSize = True}
                _cboScope = New ComboBox() With {.Location = New System.Drawing.Point(20, 98), .Width = 340, .DropDownStyle = ComboBoxStyle.DropDownList}
                _cboScope.Items.AddRange({"Active Sheet", "All Sheets"})
                _cboScope.SelectedIndex = 0

                _chkResetPerSheet = New CheckBox() With {
                    .Text = "Reset chữ cái mỗi Sheet (chỉ dùng khi All Sheets)",
                    .Location = New System.Drawing.Point(20, 130),
                    .AutoSize = True,
                    .Checked = True
                }

                Dim lbl3 As New Label() With {.Text = "Bỏ qua loại View:", .Location = New System.Drawing.Point(20, 165), .AutoSize = True}

                _chkBase = New CheckBox() With {.Text = "Base View", .Location = New System.Drawing.Point(20, 190), .AutoSize = True, .Checked = True}
                _chkProjected = New CheckBox() With {.Text = "Projected View", .Location = New System.Drawing.Point(20, 215), .AutoSize = True, .Checked = True}
                _chkSection = New CheckBox() With {.Text = "Section View", .Location = New System.Drawing.Point(20, 240), .AutoSize = True, .Checked = False}
                _chkDetail = New CheckBox() With {.Text = "Detail View", .Location = New System.Drawing.Point(20, 265), .AutoSize = True, .Checked = False}
                _chkAuxiliary = New CheckBox() With {.Text = "Auxiliary View", .Location = New System.Drawing.Point(20, 290), .AutoSize = True, .Checked = False}
                _chkOverlay = New CheckBox() With {.Text = "Overlay View", .Location = New System.Drawing.Point(20, 315), .AutoSize = True, .Checked = True}

                Dim btnOK As New Button() With {.Text = "OK", .Location = New System.Drawing.Point(185, 350), .Height = 30, .Width = 85, .DialogResult = DialogResult.OK}
                Dim btnCancel As New Button() With {.Text = "Cancel", .Location = New System.Drawing.Point(270, 350), .Height = 30, .Width = 85, .DialogResult = DialogResult.Cancel}

                _frm.Controls.AddRange({
                    lbl1, _txtLetter, lbl2, _cboScope, _chkResetPerSheet, lbl3,
                    _chkBase, _chkProjected, _chkSection, _chkDetail, _chkAuxiliary, _chkOverlay,
                    btnOK, btnCancel
                })

                _frm.AcceptButton = btnOK
                _frm.CancelButton = btnCancel
            End If

            ' Reset giá trị mặc định mỗi lần mở
            _txtLetter.Text = "A"
            _cboScope.SelectedIndex = 0
            _chkResetPerSheet.Checked = True
            _chkBase.Checked = True
            _chkProjected.Checked = True
            _chkSection.Checked = False
            _chkDetail.Checked = False
            _chkAuxiliary.Checked = False
            _chkOverlay.Checked = True

            If _frm.ShowDialog() = DialogResult.OK Then
                startLetter = _txtLetter.Text
                scope = _cboScope.SelectedItem.ToString()
                skipBase = _chkBase.Checked
                skipProjected = _chkProjected.Checked
                skipSection = _chkSection.Checked
                skipDetail = _chkDetail.Checked
                skipAuxiliary = _chkAuxiliary.Checked
                skipOverlay = _chkOverlay.Checked
                resetPerSheet = _chkResetPerSheet.Checked
                Return True
            End If

            Return False
        End Function

    End Module
End Namespace