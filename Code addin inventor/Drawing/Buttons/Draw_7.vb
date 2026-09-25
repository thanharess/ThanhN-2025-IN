Option Explicit On
Option Strict Off

Imports Inventor
Imports System
Imports Drw = System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Namespace ToolInventor2025.Drawing.Buttons

    '=============================================================
    ' SHEET NAVIGATOR BUTTON — 1 hàng, 230×40
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
    ' SHEET NAVIGATOR FORM — 230×40
    '=============================================================
    Public Class ThanhNSheetNavigatorForm
        Inherits System.Windows.Forms.Form

        '--- Win32 ---
        <DllImport("user32.dll")>
        Private Shared Function GetForegroundWindow() As IntPtr
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetParent(ByVal hWnd As IntPtr) As IntPtr
        End Function

        <DllImport("user32.dll")>
        Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
        End Function

        '--- Inventor ---
        Private ReadOnly invApp As Inventor.Application

        '--- Controls ---
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
        Private lastScreenDeviceName As String = ""

        '--- Cờ ---
        Private transitioning As Boolean = False

        '--- Kích thước ---
        Private Const PAD As Integer = 5
        Private Const BTN_H As Integer = 30

        '=========================================================
        ' SHOW WITHOUT ACTIVATION
        '=========================================================
        Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
            Get
                Return False
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

            '--- Form setup — 230×40 ---
            Me.Text = ""
            Me.AutoScaleMode = AutoScaleMode.None
            Me.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
            Me.ClientSize = New Drw.Size(230, 40)          ' ⭐ 230×40
            Me.FormBorderStyle = FormBorderStyle.None
            Me.StartPosition = FormStartPosition.Manual
            Me.ShowInTaskbar = False
            Me.TopMost = True
            Me.KeyPreview = True
            Me.BackColor = Drw.Color.FromArgb(45, 45, 48)
            Me.ForeColor = Drw.Color.White
            Me.Font = New Drw.Font("Segoe UI", 6.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            Me.Opacity = 0.75

            BuildUI()
            UpdateInfo(True)

            '--- Timer cập nhật thông tin sheet ---
            refreshTimer = New Timer()
            refreshTimer.Interval = 400
            AddHandler refreshTimer.Tick, AddressOf RefreshTimer_Tick
            refreshTimer.Start()

            '--- Timer kiểm tra foreground ---
            foregroundTimer = New Timer()
            foregroundTimer.Interval = 300
            AddHandler foregroundTimer.Tick, AddressOf ForegroundTimer_Tick
            foregroundTimer.Start()
        End Sub

        '=========================================================
        ' ĐẶT VỊ TRÍ SAU KHI FORM HIỆN
        '=========================================================
        Protected Overrides Sub OnShown(ByVal e As EventArgs)
            MyBase.OnShown(e)
            PositionOnInventorScreen()
        End Sub

        '=========================================================
        ' BUILD UI — 1 HÀNG, 230×40
        '
        ' Layout:
        '   [5] |<  [3] <  [3] [ N ] [3] OK [3] > [3] >| [5] ✕ [5]
        '=========================================================
        Private Sub BuildUI()

            Dim yBtn As Integer = PAD                       ' = 5
            Dim x As Integer = PAD                          ' = 5

            '═════════════════════════════════════════════════════
            ' NÚT |<
            '═════════════════════════════════════════════════════
            btnFirst = New Button()
            btnFirst.Text = "|<"
            btnFirst.Location = New Drw.Point(x, yBtn)
            btnFirst.Size = New Drw.Size(24, BTN_H)
            ApplyToolbarButtonStyle(btnFirst)
            Me.Controls.Add(btnFirst)
            x += 24 + 3                                     ' = 32

            '═════════════════════════════════════════════════════
            ' NÚT <
            '═════════════════════════════════════════════════════
            btnPrev = New Button()
            btnPrev.Text = "<"
            btnPrev.Location = New Drw.Point(x, yBtn)
            btnPrev.Size = New Drw.Size(24, BTN_H)
            ApplyToolbarButtonStyle(btnPrev)
            Me.Controls.Add(btnPrev)
            x += 24 + 3                                     ' = 59

            '═════════════════════════════════════════════════════
            ' TEXTBOX — rộng 56
            '═════════════════════════════════════════════════════
            txtPage = New System.Windows.Forms.TextBox()
            txtPage.Location = New Drw.Point(x, yBtn + 7)  ' 30-16 = 14, /2 = 7
            txtPage.Size = New Drw.Size(56, 16)
            txtPage.TextAlign = HorizontalAlignment.Center
            txtPage.Font = New Drw.Font("Segoe UI", 8.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            txtPage.BackColor = Drw.Color.FromArgb(30, 30, 34)
            txtPage.ForeColor = Drw.Color.White
            txtPage.BorderStyle = BorderStyle.FixedSingle
            Me.Controls.Add(txtPage)
            x += 56 + 3                                     ' = 118

            '═════════════════════════════════════════════════════
            ' NÚT OK
            '═════════════════════════════════════════════════════
            Dim btnGo As New Button()
            btnGo.Text = "OK"
            btnGo.Location = New Drw.Point(x, yBtn)
            btnGo.Size = New Drw.Size(30, BTN_H)
            ApplyToolbarButtonStyle(btnGo)
            btnGo.BackColor = Drw.Color.FromArgb(45, 100, 180)
            btnGo.ForeColor = Drw.Color.White
            btnGo.Font = New Drw.Font("Segoe UI", 7.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            Me.Controls.Add(btnGo)
            x += 30 + 3                                     ' = 151

            '═════════════════════════════════════════════════════
            ' NÚT >
            '═════════════════════════════════════════════════════
            btnNext = New Button()
            btnNext.Text = ">"
            btnNext.Location = New Drw.Point(x, yBtn)
            btnNext.Size = New Drw.Size(24, BTN_H)
            ApplyToolbarButtonStyle(btnNext)
            Me.Controls.Add(btnNext)
            x += 24 + 3                                     ' = 178

            '═════════════════════════════════════════════════════
            ' NÚT >|
            '═════════════════════════════════════════════════════
            btnLast = New Button()
            btnLast.Text = ">|"
            btnLast.Location = New Drw.Point(x, yBtn)
            btnLast.Size = New Drw.Size(24, BTN_H)
            ApplyToolbarButtonStyle(btnLast)
            Me.Controls.Add(btnLast)
            x += 24 + 5                                     ' = 207 (gap 5px)

            '═════════════════════════════════════════════════════
            ' NÚT ✕
            '═════════════════════════════════════════════════════
            btnClose = New Button()
            btnClose.Text = "✕"
            btnClose.Location = New Drw.Point(x, yBtn + 8)  ' căn giữa: (30-14)/2 = 8
            btnClose.Size = New Drw.Size(18, 14)
            ApplyToolbarButtonStyle(btnClose)
            btnClose.ForeColor = Drw.Color.FromArgb(255, 180, 180)
            btnClose.Font = New Drw.Font("Segoe UI", 6.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            Me.Controls.Add(btnClose)
            '--- x = 207, + 18 = 225, padding phải = 230 - 225 = 5px ✓

            '═════════════════════════════════════════════════════
            ' EVENTS
            '═════════════════════════════════════════════════════
            AddHandler btnFirst.Click, Sub() GoToSheet(1)
            AddHandler btnPrev.Click, Sub() GoToSheet(GetCurrentSheetIndex() - 1)
            AddHandler btnNext.Click, Sub() GoToSheet(GetCurrentSheetIndex() + 1)
            AddHandler btnLast.Click, Sub() GoToLastSheet()
            AddHandler btnClose.Click, Sub() Me.Close()

            AddHandler btnGo.Click, AddressOf BtnGo_Click

            AddHandler txtPage.KeyDown, AddressOf TxtPage_KeyDown
            AddHandler txtPage.Enter, AddressOf TxtPage_Enter

            AddHandler Me.MouseWheel, AddressOf Form_MouseWheel
        End Sub

        '=========================================================
        ' TEXTBOX ENTER
        '=========================================================
        Private Sub TxtPage_Enter(ByVal sender As Object, ByVal e As EventArgs)
            Try
                txtPage.SelectAll()
            Catch
            End Try
        End Sub

        '=========================================================
        ' NÚT OK
        '=========================================================
        Private Sub BtnGo_Click(ByVal sender As Object, ByVal e As EventArgs)
            GoToSheetFromText()
        End Sub

        '=========================================================
        ' ĐẶT VỊ TRÍ FORM
        '=========================================================
        Private Sub PositionOnInventorScreen()
            Try
                Dim targetScreen As Screen = Nothing

                Try
                    Dim invHwnd As IntPtr = New IntPtr(invApp.MainFrameHWND)
                    If invHwnd <> IntPtr.Zero Then
                        targetScreen = Screen.FromHandle(invHwnd)
                    End If
                Catch
                End Try

                If targetScreen Is Nothing Then
                    Try : targetScreen = Screen.FromPoint(Cursor.Position) : Catch : End Try
                End If

                If targetScreen Is Nothing Then
                    targetScreen = Screen.PrimaryScreen
                End If

                Dim wa As Drw.Rectangle = targetScreen.WorkingArea

                Dim newX As Integer = wa.Left + (wa.Width - Me.Width) \ 2
                Dim newY As Integer = wa.Bottom - Me.Height - 50

                Me.Location = New Drw.Point(newX, newY)
                lastScreenDeviceName = targetScreen.DeviceName
            Catch
            End Try
        End Sub

        '=========================================================
        ' KIỂM TRA WINDOW
        '=========================================================
        Private Function IsWindowPartOfForm(ByVal hWnd As IntPtr) As Boolean
            If hWnd = IntPtr.Zero Then Return False

            Try
                If hWnd = Me.Handle Then Return True

                Dim h As IntPtr = hWnd
                Dim loopCount As Integer = 0

                While h <> IntPtr.Zero AndAlso loopCount < 10
                    h = GetParent(h)
                    If h = Me.Handle Then Return True
                    loopCount += 1
                End While
            Catch
            End Try

            Return False
        End Function

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
            btn.Font = New Drw.Font("Segoe UI", 7.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            btn.TabStop = False
            btn.Cursor = Cursors.Hand
            btn.UseVisualStyleBackColor = False
        End Sub

        '=========================================================
        ' LẤY DRAWING / SHEET INDEX
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
        ' UPDATE INFO
        '=========================================================
        Private Sub UpdateInfo(Optional ByVal force As Boolean = False)
            Try
                Dim oDoc As Inventor.DrawingDocument = GetCurrentDrawing()

                If oDoc Is Nothing Then
                    If force Then
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

                If Not txtPage.Focused Then
                    txtPage.Text = currentIdx.ToString()
                End If

                Try
                    Me.Text = "Sheet " & currentIdx.ToString() & " / " & total.ToString()
                Catch
                End Try
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
        ' TIMER FOREGROUND
        '=========================================================
        Private Sub ForegroundTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
            Try
                If transitioning Then Return

                Dim fg As IntPtr = GetForegroundWindow()
                If fg = IntPtr.Zero Then Return

                Dim invHwnd As IntPtr = IntPtr.Zero
                Try : invHwnd = New IntPtr(invApp.MainFrameHWND) : Catch : End Try

                Dim isInv As Boolean = (invHwnd <> IntPtr.Zero AndAlso fg = invHwnd)
                Dim isMine As Boolean = IsWindowPartOfForm(fg)

                If isInv OrElse isMine Then
                    If Not Me.Visible Then
                        Me.Show()
                        PositionOnInventorScreen()
                        UpdateInfo(True)
                    End If
                Else
                    If Me.Visible Then
                        Me.Hide()
                    End If
                    Return
                End If

                If invHwnd <> IntPtr.Zero Then
                    Dim invScreen As Screen = Nothing
                    Try : invScreen = Screen.FromHandle(invHwnd) : Catch : End Try

                    If invScreen IsNot Nothing AndAlso invScreen.DeviceName <> lastScreenDeviceName Then
                        PositionOnInventorScreen()
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

                transitioning = True

                oDoc.Sheets.Item(index).Activate()
                UpdateInfo(True)

                Try
                    Dim invHwnd As IntPtr = New IntPtr(invApp.MainFrameHWND)
                    If invHwnd <> IntPtr.Zero Then
                        SetForegroundWindow(invHwnd)
                    End If
                Catch
                End Try

                transitioning = False
            Catch ex As Exception
                transitioning = False
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
                    Exit Sub
                End If

                If e.KeyCode = Keys.Down Then
                    e.SuppressKeyPress = True
                    e.Handled = True
                    GoToSheet(GetCurrentSheetIndex() + 1)
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