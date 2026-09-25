Option Explicit On
Option Strict Off

Imports Inventor
Imports System
Imports Drw = System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Namespace ToolInventor2025.Drawing.Buttons

    '=============================================================
    ' SHEET NAVIGATOR BUTTON
    ' Inventor 2025 — Toolbar nổi giữa màn hình
    ' Tự ẩn khi chuyển sang app khác, tự hiện khi quay lại Inventor
    '=============================================================
    Public Module Draw_7

        Private m_Form As ThanhNSheetNavigatorForm = Nothing

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                If m_Form IsNot Nothing AndAlso Not m_Form.IsDisposed Then
                    If Not m_Form.Visible Then m_Form.Show()
                    m_Form.BringToFront()
                    Exit Sub
                End If

                m_Form = Nothing

                Dim invApp As Inventor.Application = g_inventorApplication
                If invApp Is Nothing Then
                    MessageBox.Show("Không tìm thấy Inventor Application.",
                                    "Sheet Navigator",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If invApp.ActiveDocument Is Nothing Then
                    MessageBox.Show("Không có tài liệu đang mở.",
                                    "Sheet Navigator",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                If invApp.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Chỉ sử dụng chức năng này trong Drawing.",
                                    "Sheet Navigator",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                m_Form = New ThanhNSheetNavigatorForm(invApp)
                AddHandler m_Form.FormClosed, AddressOf NavigatorFormClosed
                m_Form.Show()
            Catch ex As Exception
                MessageBox.Show("Lỗi Sheet Navigator:" & vbCrLf & ex.Message,
                                "Sheet Navigator",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub NavigatorFormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs)
            Try
                If m_Form IsNot Nothing Then
                    RemoveHandler m_Form.FormClosed, AddressOf NavigatorFormClosed
                End If
            Catch
            End Try
            m_Form = Nothing
        End Sub

    End Module


    '=============================================================
    ' SHEET NAVIGATOR FORM — giữa màn hình, tự ẩn/hiện
    '=============================================================
    Public Class ThanhNSheetNavigatorForm
        Inherits System.Windows.Forms.Form

        '--- Win32 ---
        <DllImport("user32.dll")>
        Private Shared Function GetForegroundWindow() As IntPtr
        End Function

        '--- Inventor ---
        Private ReadOnly invApp As Inventor.Application

        '--- Controls ---
        Private lblInfo As Label
        Private txtPage As System.Windows.Forms.TextBox
        Private btnFirst As Button
        Private btnPrev As Button
        Private btnNext As Button
        Private btnLast As Button
        Private btnClose As Button

        '--- Timer ---
        Private refreshTimer As Timer
        Private foregroundTimer As Timer

        '--- Cache ---
        Private lastSheetIndex As Integer = -1
        Private lastSheetCount As Integer = -1
        Private lastDocPath As String = ""

        '=========================================================
        ' KHÔNG CƯỚP FOCUS INVENTOR
        '=========================================================
        Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
            Get
                Return True
            End Get
        End Property

        Protected Overrides ReadOnly Property CreateParams As CreateParams
            Get
                Dim cp As CreateParams = MyBase.CreateParams
                cp.ExStyle = cp.ExStyle Or &H8000000   ' WS_EX_NOACTIVATE
                cp.ExStyle = cp.ExStyle Or &H80        ' WS_EX_TOOLWINDOW
                Return cp
            End Get
        End Property

        '=========================================================
        ' CONSTRUCTOR
        '=========================================================
        Public Sub New(ByVal app As Inventor.Application)
            MyBase.New()

            If app Is Nothing Then Throw New Exception("Inventor Application không hợp lệ.")
            If app.ActiveDocument Is Nothing Then Throw New Exception("Không có document đang mở.")
            If app.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                Throw New Exception("Document hiện tại không phải Drawing.")
            End If

            invApp = app

            '--- Form setup ---
            Me.Text = ""
            Me.AutoScaleMode = AutoScaleMode.None
            Me.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
            Me.ClientSize = New Drw.Size(232, 74)
            Me.FormBorderStyle = FormBorderStyle.None
            Me.StartPosition = FormStartPosition.CenterScreen   ' ⭐ Giữa màn hình
            Me.ShowInTaskbar = False
            Me.TopMost = True
            Me.KeyPreview = True
            Me.BackColor = Drw.Color.FromArgb(45, 45, 48)
            Me.ForeColor = Drw.Color.White
            Me.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

            BuildUI()
            UpdateInfo(True)

            '--- Timer cập nhật thông tin sheet ---
            refreshTimer = New Timer()
            refreshTimer.Interval = 400
            AddHandler refreshTimer.Tick, AddressOf RefreshTimer_Tick
            refreshTimer.Start()

            '--- Timer kiểm tra foreground window (ẩn/hiện) ---
            foregroundTimer = New Timer()
            foregroundTimer.Interval = 250
            AddHandler foregroundTimer.Tick, AddressOf ForegroundTimer_Tick
            foregroundTimer.Start()
        End Sub

        '=========================================================
        ' BUILD UI
        '=========================================================
        Private Sub BuildUI()

            '═════════════════════════════════════════════════════
            ' HÀNG 1: INFO + CLOSE
            '═════════════════════════════════════════════════════
            lblInfo = New Label()
            lblInfo.Text = "Sheet 1/1"
            lblInfo.Location = New Drw.Point(10, 6)
            lblInfo.Size = New Drw.Size(186, 22)
            lblInfo.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblInfo.ForeColor = Drw.Color.FromArgb(220, 230, 245)
            lblInfo.TextAlign = Drw.ContentAlignment.MiddleLeft
            Me.Controls.Add(lblInfo)

            btnClose = New Button()
            btnClose.Text = "✕"
            btnClose.Location = New Drw.Point(202, 5)
            btnClose.Size = New Drw.Size(22, 22)
            ApplyToolbarButtonStyle(btnClose)
            btnClose.ForeColor = Drw.Color.FromArgb(255, 180, 180)
            btnClose.Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            Me.Controls.Add(btnClose)

            '═════════════════════════════════════════════════════
            ' HÀNG 2: NAVIGATOR
            '═════════════════════════════════════════════════════
            Dim yBtn As Integer = 36
            Dim btnH As Integer = 28

            btnFirst = New Button()
            btnFirst.Text = "|<"
            btnFirst.Location = New Drw.Point(10, yBtn)
            btnFirst.Size = New Drw.Size(38, btnH)
            ApplyToolbarButtonStyle(btnFirst)
            Me.Controls.Add(btnFirst)

            btnPrev = New Button()
            btnPrev.Text = "<"
            btnPrev.Location = New Drw.Point(52, yBtn)
            btnPrev.Size = New Drw.Size(38, btnH)
            ApplyToolbarButtonStyle(btnPrev)
            Me.Controls.Add(btnPrev)

            txtPage = New System.Windows.Forms.TextBox()
            txtPage.Location = New Drw.Point(94, yBtn + 3)
            txtPage.Size = New Drw.Size(44, 22)
            txtPage.TextAlign = HorizontalAlignment.Center
            txtPage.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            txtPage.BackColor = Drw.Color.FromArgb(30, 30, 34)
            txtPage.ForeColor = Drw.Color.White
            txtPage.BorderStyle = BorderStyle.FixedSingle
            Me.Controls.Add(txtPage)

            btnNext = New Button()
            btnNext.Text = ">"
            btnNext.Location = New Drw.Point(142, yBtn)
            btnNext.Size = New Drw.Size(38, btnH)
            ApplyToolbarButtonStyle(btnNext)
            Me.Controls.Add(btnNext)

            btnLast = New Button()
            btnLast.Text = ">|"
            btnLast.Location = New Drw.Point(184, yBtn)
            btnLast.Size = New Drw.Size(38, btnH)
            ApplyToolbarButtonStyle(btnLast)
            Me.Controls.Add(btnLast)

            '═════════════════════════════════════════════════════
            ' EVENTS
            '═════════════════════════════════════════════════════
            AddHandler btnFirst.Click, Sub() GoToSheet(1)
            AddHandler btnPrev.Click, Sub() GoToSheet(GetCurrentSheetIndex() - 1)
            AddHandler btnNext.Click, Sub() GoToSheet(GetCurrentSheetIndex() + 1)
            AddHandler btnLast.Click, Sub() GoToLastSheet()
            AddHandler btnClose.Click, Sub() Me.Close()

            AddHandler txtPage.KeyDown, AddressOf TxtPage_KeyDown

            '--- Scroll wheel trên form = chuyển sheet ---
            AddHandler Me.MouseWheel, AddressOf Form_MouseWheel
            AddHandler lblInfo.MouseWheel, AddressOf Form_MouseWheel
        End Sub

        '=========================================================
        ' BUTTON STYLE
        '=========================================================
        Private Sub ApplyToolbarButtonStyle(ByVal btn As Button)
            btn.BackColor = Drw.Color.FromArgb(63, 63, 70)
            btn.ForeColor = Drw.Color.White
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 1
            btn.FlatAppearance.BorderColor = Drw.Color.FromArgb(90, 90, 100)
            btn.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(80, 80, 95)
            btn.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(45, 100, 180)
            btn.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            btn.TabStop = False
            btn.Cursor = Cursors.Hand
            btn.UseVisualStyleBackColor = False
        End Sub

        '=========================================================
        ' LẤY DRAWING HIỆN TẠI
        '=========================================================
        Private Function GetCurrentDrawing() As Inventor.DrawingDocument
            Try
                If invApp Is Nothing Then Return Nothing
                If invApp.ActiveDocument Is Nothing Then Return Nothing
                If invApp.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                    Return Nothing
                End If
                Return DirectCast(invApp.ActiveDocument, Inventor.DrawingDocument)
            Catch
                Return Nothing
            End Try
        End Function

        Private Function GetCurrentSheetIndex() As Integer
            Try
                Dim oDoc As Inventor.DrawingDocument = GetCurrentDrawing()
                If oDoc Is Nothing Then Return 1
                For i As Integer = 1 To oDoc.Sheets.Count
                    If oDoc.Sheets.Item(i) Is oDoc.ActiveSheet Then Return i
                Next
            Catch
            End Try
            Return 1
        End Function

        '=========================================================
        ' UPDATE INFO — có cache
        '=========================================================
        Private Sub UpdateInfo(Optional ByVal force As Boolean = False)
            Try
                Dim oDoc As Inventor.DrawingDocument = GetCurrentDrawing()

                If oDoc Is Nothing Then
                    If force OrElse lblInfo.Text <> "Không phải Drawing" Then
                        lblInfo.Text = "Không phải Drawing"
                        txtPage.Text = ""
                        SetButtonsEnabled(False)
                    End If
                    Return
                End If

                SetButtonsEnabled(True)

                Dim currentIdx As Integer = GetCurrentSheetIndex()
                Dim total As Integer = oDoc.Sheets.Count
                Dim docPath As String = ""
                Try : docPath = oDoc.FullFileName : Catch : End Try

                If Not force AndAlso
                   currentIdx = lastSheetIndex AndAlso
                   total = lastSheetCount AndAlso
                   docPath = lastDocPath Then
                    Return
                End If

                lastSheetIndex = currentIdx
                lastSheetCount = total
                lastDocPath = docPath

                lblInfo.Text = "Sheet  " & currentIdx.ToString() & " / " & total.ToString()

                If Not txtPage.Focused Then
                    txtPage.Text = currentIdx.ToString()
                End If
            Catch
            End Try
        End Sub

        Private Sub SetButtonsEnabled(ByVal en As Boolean)
            Try
                btnFirst.Enabled = en
                btnPrev.Enabled = en
                btnNext.Enabled = en
                btnLast.Enabled = en
                txtPage.Enabled = en
            Catch
            End Try
        End Sub

        '=========================================================
        ' TIMER CẬP NHẬT INFO
        '=========================================================
        Private Sub RefreshTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
            UpdateInfo(False)
        End Sub

        '=========================================================
        ' ⭐ TIMER KIỂM TRA FOREGROUND — ẨN/HIỆN FORM
        '
        ' - Inventor ở foreground → form hiện
        ' - App khác ở foreground → form ẩn
        '=========================================================
        Private Sub ForegroundTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
            Try
                Dim fg As IntPtr = GetForegroundWindow()
                If fg = IntPtr.Zero Then Return

                Dim invHwnd As IntPtr = IntPtr.Zero
                Try : invHwnd = New IntPtr(invApp.MainFrameHWND) : Catch : End Try

                Dim myHwnd As IntPtr = IntPtr.Zero
                Try : myHwnd = Me.Handle : Catch : End Try

                '--- Nếu là cửa sổ của Inventor hoặc của form → show ---
                If fg = invHwnd OrElse fg = myHwnd Then
                    If Not Me.Visible Then
                        Me.Show()
                        UpdateInfo(True)
                    End If
                Else
                    '--- Window khác → ẩn form ---
                    If Me.Visible Then
                        Me.Hide()
                    End If
                End If
            Catch
            End Try
        End Sub

        '=========================================================
        ' GO TO SHEET
        '=========================================================
        Private Sub GoToSheet(ByVal index As Integer)
            Try
                Dim oDoc As Inventor.DrawingDocument = GetCurrentDrawing()
                If oDoc Is Nothing Then Return

                Dim total As Integer = oDoc.Sheets.Count
                If total <= 0 Then Return

                If index < 1 Then index = 1
                If index > total Then index = total

                oDoc.Sheets.Item(index).Activate()
                UpdateInfo(True)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Sheet Navigator",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

        Private Sub GoToLastSheet()
            Try
                Dim oDoc As Inventor.DrawingDocument = GetCurrentDrawing()
                If oDoc Is Nothing Then Return
                GoToSheet(oDoc.Sheets.Count)
            Catch
            End Try
        End Sub

        '=========================================================
        ' TEXTBOX KEYDOWN
        '=========================================================
        Private Sub TxtPage_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            Try
                If e.KeyCode = Keys.Enter Then
                    e.SuppressKeyPress = True
                    e.Handled = True
                    GoToSheetFromText()
                    Exit Sub
                End If

                If e.KeyCode = Keys.Escape Then
                    e.SuppressKeyPress = True
                    e.Handled = True
                    UpdateInfo(True)
                    Exit Sub
                End If

                If e.KeyCode = Keys.Up Then
                    e.SuppressKeyPress = True
                    e.Handled = True
                    GoToSheet(GetCurrentSheetIndex() - 1)
                    txtPage.SelectAll()
                    Exit Sub
                End If

                If e.KeyCode = Keys.Down Then
                    e.SuppressKeyPress = True
                    e.Handled = True
                    GoToSheet(GetCurrentSheetIndex() + 1)
                    txtPage.SelectAll()
                    Exit Sub
                End If
            Catch
            End Try
        End Sub

        Private Sub GoToSheetFromText()
            Try
                Dim pageNumber As Integer
                If Not Integer.TryParse(txtPage.Text.Trim(), pageNumber) Then
                    MessageBox.Show("Nhập số Sheet hợp lệ.", "Sheet Navigator",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtPage.SelectAll()
                    txtPage.Focus()
                    Return
                End If
                GoToSheet(pageNumber)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Sheet Navigator",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

        '=========================================================
        ' PROCESS CMD KEY
        '=========================================================
        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
            Try
                If keyData = Keys.Escape AndAlso Not txtPage.Focused Then
                    Me.Close()
                    Return True
                End If

                If keyData = Keys.Space AndAlso Not txtPage.Focused Then
                    Return True
                End If

                If keyData = Keys.Enter AndAlso txtPage.Focused Then
                    GoToSheetFromText()
                    Return True
                End If

                If keyData = Keys.PageUp Then
                    GoToSheet(GetCurrentSheetIndex() - 1)
                    Return True
                End If

                If keyData = Keys.PageDown Then
                    GoToSheet(GetCurrentSheetIndex() + 1)
                    Return True
                End If
            Catch
            End Try
            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        '=========================================================
        ' SCROLL WHEEL
        '=========================================================
        Private Sub Form_MouseWheel(ByVal sender As Object, ByVal e As MouseEventArgs)
            Try
                If e.Delta > 0 Then
                    GoToSheet(GetCurrentSheetIndex() - 1)
                ElseIf e.Delta < 0 Then
                    GoToSheet(GetCurrentSheetIndex() + 1)
                End If
            Catch
            End Try
        End Sub

        '=========================================================
        ' FORM EVENTS
        '=========================================================
        Protected Overrides Sub OnFormClosed(ByVal e As FormClosedEventArgs)
            Try
                If refreshTimer IsNot Nothing Then
                    refreshTimer.Stop()
                    refreshTimer.Dispose()
                    refreshTimer = Nothing
                End If
                If foregroundTimer IsNot Nothing Then
                    foregroundTimer.Stop()
                    foregroundTimer.Dispose()
                    foregroundTimer = Nothing
                End If
            Catch
            End Try
            MyBase.OnFormClosed(e)
        End Sub

    End Class

End Namespace