Option Explicit On
Option Strict Off

Imports Inventor
Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace ToolInventor2025.Drawing.Buttons
    '=============================================================
    ' SHEET NAVIGATOR BUTTON
    ' INVENTOR 2020 - VB.NET VISUAL STUDIO
    ' MODELLESS TOOLBAR
    '=============================================================
    Public Module Draw_7
        '=========================================================
        ' FORM DUY NHẤT
        '=========================================================
        Private m_Form As ThanhNSheetNavigatorForm = Nothing
        '=========================================================
        ' MAIN BUTTON
        '========================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                '=================================================
                ' NẾU FORM ĐÃ TỒN TẠI
                ' KHÔNG TẠO FORM MỚI
                '=================================================
                If m_Form IsNot Nothing Then
                    If Not m_Form.IsDisposed Then
                        If Not m_Form.Visible Then
                            m_Form.Show()
                        End If
                        m_Form.BringToFront()
                        Exit Sub
                    Else
                        m_Form = Nothing
                    End If
                End If
                '=================================================
                ' INVENTOR APPLICATION
                '=================================================
                Dim invApp As Inventor.Application = g_inventorApplication
                If invApp Is Nothing Then
                    MessageBox.Show("Không tìm thấy Inventor Application.", "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                '=================================================
                ' KIỂM TRA DOCUMENT
                '=================================================
                If invApp.ActiveDocument Is Nothing Then
                    MessageBox.Show("Không có tài liệu đang mở.", "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
                '=================================================
                ' CHỈ DRAWING
                '=================================================
                If invApp.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Chỉ sử dụng chức năng này trong Drawing.", "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
                '=================================================
                ' TẠO FORM DUY NHẤT
                '=================================================
                m_Form = New ThanhNSheetNavigatorForm(invApp)
                '=================================================
                ' KHI FORM ĐÓNG
                '=================================================
                AddHandler m_Form.FormClosed, AddressOf NavigatorFormClosed
                '=================================================
                ' MODELLESS
                ' KHÔNG KHÓA INVENTOR
                '=================================================
                m_Form.Show()
            Catch ex As Exception
                MessageBox.Show("Lỗi Sheet Navigator:" & vbCrLf & ex.Message, "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
        '=========================================================
        ' FORM CLOSED
        '=========================================================
        Private Sub NavigatorFormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs)
            Try
                If m_Form IsNot Nothing Then
                    RemoveHandler m_Form.FormClosed, AddressOf NavigatorFormClosed
                End If
            Catch
            End Try
            '=====================================================
            ' RESET FORM
            '====================================================
            m_Form = Nothing
        End Sub
    End Module
    '=============================================================
    '    ' SHEET NAVIGATOR FORM
    '    ' INVENTOR 2020
    ' VB.NET VISUAL STUDIO
    ' MODELLESS TOOLBAR    '
    '=============================================================
    Public Class ThanhNSheetNavigatorForm
        Inherits System.Windows.Forms.Form
        '=========================================================
        ' INVENTOR
        '=========================================================
        Private ReadOnly invApp As Inventor.Application
        '=========================================================
        ' CONTROLS
        '=========================================================
        Private lblInfo As System.Windows.Forms.Label
        Private txtPage As System.Windows.Forms.TextBox
        Private btnFirst As System.Windows.Forms.Button
        Private btnPrev As System.Windows.Forms.Button
        Private btnNext As System.Windows.Forms.Button
        Private btnLast As System.Windows.Forms.Button
        Private btnGo As System.Windows.Forms.Button
        Private btnClose As System.Windows.Forms.Button
        Private chkLeft As System.Windows.Forms.CheckBox
        Private chkRight As System.Windows.Forms.CheckBox
        '=========================================================
        ' DRAG FORM
        '=========================================================
        Private isDragging As Boolean = False
        Private dragCursorPoint As System.Drawing.Point
        Private dragFormPoint As System.Drawing.Point

        '=========================================================
        ' CONSTRUCTOR
        '=========================================================
        Public Sub New(ByVal app As Inventor.Application)
            MyBase.New()
            '=====================================================
            ' INVENTOR
            '=====================================================
            invApp = app
            If invApp Is Nothing Then
                Throw New Exception("Inventor Application không hợp lệ.")
            End If
            If invApp.ActiveDocument Is Nothing Then
                Throw New Exception("Không có document đang mở.")
            End If
            If invApp.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                Throw New Exception("Document hiện tại không phải Drawing.")
            End If
            '=====================================================
            ' FORM SETUP
            '=====================================================
            Me.Text = ""
            Me.Width = 220
            Me.Height = 85
            Me.StartPosition = FormStartPosition.CenterScreen
            '=====================================================
            ' KHÔNG VIỀN
            '=====================================================
            Me.FormBorderStyle = FormBorderStyle.None
            '=====================================================
            ' KHÔNG HIỆN TASKBAR
            '=====================================================
            Me.ShowInTaskbar = False
            '=====================================================
            ' MODELLESS TOOLBAR            '
            ' FALSE:
            ' KHÔNG LUÔN CHIẾM TRÊN INVENTOR
            '=====================================================
            Me.TopMost = True
            '=====================================================
            ' ĐỘ TRONG SUỐT
            '=====================================================
            Me.Opacity = 0.55
            '=====================================================
            ' BACKGROUND
            '=====================================================
            Me.BackColor = System.Drawing.Color.FromArgb(45, 45, 48)
            '=====================================================
            ' KHÔNG CHO FORM NHẬN FOCUS KHÔNG CẦN THIẾT
            '=====================================================
            Me.KeyPreview = True
            '=====================================================
            '            ' INFO            '
            '=====================================================
            lblInfo = New System.Windows.Forms.Label()
            lblInfo.Left = 10
            lblInfo.Top = 8
            lblInfo.Width = 160
            lblInfo.Height = 18
            lblInfo.ForeColor = System.Drawing.Color.White
            lblInfo.Font = New System.Drawing.Font("Arial", 8)
            Me.Controls.Add(lblInfo)
            '=====================================================
            '            ' CLOSE            '
            '=====================================================
            btnClose = New System.Windows.Forms.Button()
            btnClose.Text = "X"
            btnClose.Left = 185
            btnClose.Top = 5
            btnClose.Width = 25
            btnClose.Height = 20
            ApplyButtonStyle(btnClose)
            Me.Controls.Add(btnClose)
            '=====================================================
            '            ' FIRST
            '            '====================================================
            btnFirst = New System.Windows.Forms.Button()
            btnFirst.Text = "|<"
            btnFirst.Left = 10
            btnFirst.Top = 30
            btnFirst.Width = 40
            btnFirst.Height = 25
            ApplyButtonStyle(btnFirst)
            Me.Controls.Add(btnFirst)
            '====================================================
            '            ' PREVIOUS            '
            '=====================================================
            btnPrev = New System.Windows.Forms.Button()
            btnPrev.Text = "<"
            btnPrev.Left = 55
            btnPrev.Top = 30
            btnPrev.Width = 35
            btnPrev.Height = 25
            ApplyButtonStyle(btnPrev)
            Me.Controls.Add(btnPrev)
            '=====================================================
            '            ' NEXT            '
            '=====================================================
            btnNext = New System.Windows.Forms.Button()
            btnNext.Text = ">"
            btnNext.Left = 95
            btnNext.Top = 30
            btnNext.Width = 35
            btnNext.Height = 25
            ApplyButtonStyle(btnNext)
            Me.Controls.Add(btnNext)
            '=====================================================
            '            ' LAST            '
            '====================================================
            btnLast = New System.Windows.Forms.Button()
            btnLast.Text = ">|"
            btnLast.Left = 135
            btnLast.Top = 30
            btnLast.Width = 40
            btnLast.Height = 25
            ApplyButtonStyle(btnLast)
            Me.Controls.Add(btnLast)
            '====================================================
            '            ' PAGE INPUT
            '            '=====================================================
            txtPage = New System.Windows.Forms.TextBox()
            txtPage.Left = 10
            txtPage.Top = 60
            txtPage.Width = 45
            txtPage.Height = 20
            txtPage.TextAlign = HorizontalAlignment.Center
            Me.Controls.Add(txtPage)
            '=====================================================
            '            ' GO            '
            '=====================================================
            btnGo = New System.Windows.Forms.Button()
            btnGo.Text = "Chuyển Sheet"
            btnGo.Left = 60
            btnGo.Top = 58
            btnGo.Width = 100
            btnGo.Height = 22
            ApplyButtonStyle(btnGo)
            Me.Controls.Add(btnGo)
            '====================================================
            '            ' DOCK LEFT            '
            '=====================================================
            chkLeft = New System.Windows.Forms.CheckBox()
            chkLeft.Text = "L"
            chkLeft.Left = 163
            chkLeft.Top = 60
            chkLeft.Width = 28
            chkLeft.ForeColor = System.Drawing.Color.White
            chkLeft.BackColor = System.Drawing.Color.Transparent
            Me.Controls.Add(chkLeft)
            '=====================================================
            '            ' DOCK RIGHT            '
            '=====================================================
            chkRight = New System.Windows.Forms.CheckBox()
            chkRight.Text = "R"
            chkRight.Left = 192
            chkRight.Top = 60
            chkRight.Width = 28
            chkRight.ForeColor = System.Drawing.Color.White
            chkRight.BackColor = System.Drawing.Color.Transparent
            Me.Controls.Add(chkRight)
            '=====================================================
            '            ' EVENTS            '
            '=====================================================
            AddHandler btnFirst.Click, AddressOf btnFirst_Click
            AddHandler btnPrev.Click, AddressOf btnPrev_Click
            AddHandler btnNext.Click, AddressOf btnNext_Click
            AddHandler btnLast.Click, AddressOf btnLast_Click
            AddHandler btnGo.Click, AddressOf btnGo_Click
            AddHandler btnClose.Click, AddressOf btnClose_Click
            AddHandler chkLeft.CheckedChanged, AddressOf chkLeft_CheckedChanged
            AddHandler chkRight.CheckedChanged, AddressOf chkRight_CheckedChanged
            '=====================================================
            ' TEXTBOX
            '=====================================================
            AddHandler txtPage.KeyDown, AddressOf txtPage_KeyDown
            '=====================================================
            ' DRAG FORM
            '=====================================================
            AddHandler Me.MouseDown, AddressOf Form_MouseDown
            AddHandler Me.MouseMove, AddressOf Form_MouseMove
            AddHandler Me.MouseUp, AddressOf Form_MouseUp
            AddHandler lblInfo.MouseDown, AddressOf Form_MouseDown
            AddHandler lblInfo.MouseMove, AddressOf Form_MouseMove
            AddHandler lblInfo.MouseUp, AddressOf Form_MouseUp
            '=====================================================
            ' UPDATE
            '=====================================================
            UpdateInfo()
        End Sub
        '=========================================================
        '        ' BUTTON STYLE
        '        '=========================================================
        Private Sub ApplyButtonStyle(ByVal btn As System.Windows.Forms.Button)
            btn.BackColor = System.Drawing.Color.FromArgb(63, 63, 70)
            btn.ForeColor = System.Drawing.Color.White
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 1
            btn.TabStop = False
        End Sub
        '=========================================================
        '        ' LẤY DRAWING HIỆN TẠI
        '        '=========================================================
        Private Function GetDrawingDocument() As Inventor.DrawingDocument
            Try
                If invApp Is Nothing Then
                    Return Nothing
                End If
                If invApp.ActiveDocument Is Nothing Then
                    Return Nothing
                End If
                If invApp.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                    Return Nothing
                End If
                Return DirectCast(invApp.ActiveDocument, Inventor.DrawingDocument)
            Catch
                Return Nothing
            End Try
        End Function
        Private Function GetCurrentDrawing() As Inventor.DrawingDocument
            Try
                If invApp Is Nothing Then
                    Return Nothing
                End If
                If invApp.ActiveDocument Is Nothing Then
                    Return Nothing
                End If
                If invApp.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then

                    Return Nothing
                End If
                Return DirectCast(invApp.ActiveDocument, Inventor.DrawingDocument)
            Catch
                Return Nothing
            End Try
        End Function
        '=========================================================
        '        ' CURRENT SHEET INDEX
        '        '=========================================================
        Private Function GetCurrentSheetIndex() As Integer
            Try
                Dim oDoc As Inventor.DrawingDocument = GetDrawingDocument()
                If oDoc Is Nothing Then
                    Return 1
                End If
                Dim i As Integer
                For i = 1 To oDoc.Sheets.Count
                    If oDoc.Sheets.Item(i) Is oDoc.ActiveSheet Then

                        Return i
                    End If
                Next
            Catch
            End Try
            Return 1
        End Function
        '=========================================================
        '        ' UPDATE INFO        '
        '=========================================================
        Private Sub UpdateInfo()
            Try
                Dim oDoc As Inventor.DrawingDocument = GetCurrentDrawing()
                If oDoc Is Nothing Then
                    lblInfo.Text = "Không phải Drawing"
                    txtPage.Text = ""
                    Exit Sub
                End If
                Dim currentIndex As Integer = GetCurrentSheetIndex()
                Dim total As Integer = oDoc.Sheets.Count
                lblInfo.Text = "Sheet " & currentIndex.ToString() & "/" & total.ToString()
                txtPage.Text = currentIndex.ToString()
            Catch
                lblInfo.Text = "Sheet"
            End Try
        End Sub
        '=========================================================
        '        ' GO TO SHEET        '
        '=========================================================
        Private Sub GoToSheet(ByVal index As Integer)
            Try
                '=====================================================
                ' LUÔN LẤY DRAWING ĐANG ACTIVE
                '=====================================================
                Dim oDoc As Inventor.DrawingDocument = GetCurrentDrawing()
                If oDoc Is Nothing Then
                    MessageBox.Show("Document hiện tại không phải Drawing.", "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
                Dim total As Integer = oDoc.Sheets.Count
                If total <= 0 Then
                    Exit Sub
                End If
                If index < 1 Then
                    index = 1
                ElseIf index > total Then
                    index = total
                End If
                '=====================================================
                ' CHUYỂN SHEET CỦA DRAWING HIỆN TẠI
                '=====================================================
                oDoc.Sheets.Item(index).Activate()
                UpdateInfo()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub
        '=========================================================        '
        ' DOCK LEFT - GIỮA MÀN HÌNH TRÁI        
        '        '=========================================================
        Private Sub chkLeft_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
            Try
                If Not chkLeft.Checked Then Exit Sub

                If chkRight.Checked Then
                    chkRight.Checked = False
                End If
                Dim screens() As Screen = Screen.AllScreens
                ' TÌM MÀN HÌNH TRÁI
                Dim leftScreen As Screen = screens(0)
                For Each s As Screen In screens
                    If s.Bounds.Left < leftScreen.Bounds.Left Then
                        leftScreen = s
                    End If
                Next
                Dim wa As System.Drawing.Rectangle = leftScreen.WorkingArea
                ' GIỮA THEO CHIỀU NGANG
                Me.Left = wa.Left + ((wa.Width - Me.Width) \ 2)
                ' CÁCH ĐÁY 100 PX
                Me.Top = wa.Bottom - Me.Height - 50
            Catch
            End Try
        End Sub
        '=========================================================
        '        ' DOCK RIGHT - GIỮA MÀN HÌNH PHẢI
        '        '=========================================================
        Private Sub chkRight_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
            Try
                If Not chkRight.Checked Then Exit Sub
                If chkLeft.Checked Then
                    chkLeft.Checked = False
                End If
                Dim screens() As Screen = Screen.AllScreens
                ' TÌM MÀN HÌNH PHẢI
                Dim rightScreen As Screen = screens(0)
                For Each s As Screen In screens
                    If s.Bounds.Right > rightScreen.Bounds.Right Then
                        rightScreen = s
                    End If
                Next
                Dim wa As System.Drawing.Rectangle = rightScreen.WorkingArea

                ' GIỮA THEO CHIỀU NGANG
                Me.Left = wa.Left + ((wa.Width - Me.Width) \ 2)
                ' CÁCH ĐÁY 100 PX
                Me.Top = wa.Bottom - Me.Height - 50
            Catch
            End Try
        End Sub
        '=========================================================
        '        ' DRAG FORM
        '        '=========================================================
        Private Sub Form_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
            Try
                '=================================================
                ' ĐANG DOCK KHÔNG KÉO
                '=================================================
                If chkLeft.Checked OrElse
                    chkRight.Checked Then
                    Exit Sub
                End If
                If e.Button <> MouseButtons.Left Then
                    Exit Sub
                End If
                isDragging = True
                dragCursorPoint = Cursor.Position
                dragFormPoint = Me.Location
            Catch
            End Try
        End Sub
        '=========================================================
        '        ' DRAG MOVE
        '        '=========================================================
        Private Sub Form_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
            Try
                If Not isDragging Then
                    Exit Sub
                End If
                Dim currentCursor As System.Drawing.Point = Cursor.Position
                Dim diffX As Integer = currentCursor.X - dragCursorPoint.X
                Dim diffY As Integer = currentCursor.Y - dragCursorPoint.Y
                Me.Location = New System.Drawing.Point(dragFormPoint.X + diffX, dragFormPoint.Y + diffY)
            Catch
            End Try
        End Sub
        '=========================================================
        '        ' DRAG UP
        '        '=========================================================
        Private Sub Form_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
            isDragging = False
        End Sub
        '=========================================================
        '        ' CLOSE
        '        '=========================================================
        Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
            Try
                Me.Close()
            Catch
            End Try
        End Sub
        '=========================================================
        '
        ' FIRST
        '
        '=========================================================
        Private Sub btnFirst_Click(ByVal sender As Object, ByVal e As EventArgs)
            GoToSheet(1)
        End Sub
        '=========================================================
        '        ' PREVIOUS        '
        '=========================================================
        Private Sub btnPrev_Click(ByVal sender As Object, ByVal e As EventArgs)
            GoToSheet(GetCurrentSheetIndex() - 1)
        End Sub

        '=========================================================
        '        ' NEXT        '
        '=========================================================
        Private Sub btnNext_Click(ByVal sender As Object, ByVal e As EventArgs)
            GoToSheet(GetCurrentSheetIndex() + 1)
        End Sub
        '=========================================================
        '        ' LAST
        '        '=========================================================
        Private Sub btnLast_Click(ByVal sender As Object, ByVal e As EventArgs)
            Try
                Dim oDoc As Inventor.DrawingDocument = GetDrawingDocument()
                If oDoc Is Nothing Then
                    Exit Sub
                End If
                GoToSheet(oDoc.Sheets.Count)
            Catch
            End Try
        End Sub

        '=========================================================
        '        ' GO BUTTON
        '        '=========================================================
        Private Sub btnGo_Click(ByVal sender As Object, ByVal e As EventArgs)
            GoToSheetFromText()
        End Sub
        '=========================================================
        '        ' TEXTBOX KEYDOWN        '
        '=========================================================
        Private Sub txtPage_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            Try
                '=================================================
                ' ENTER
                '=================================================
                If e.KeyCode = Keys.Enter Then
                    e.SuppressKeyPress = True
                    e.Handled = True
                    GoToSheetFromText()
                    Exit Sub
                End If
                '=================================================
                ' ESC
                '=================================================
                If e.KeyCode = Keys.Escape Then
                    e.SuppressKeyPress = True
                    e.Handled = True
                    Me.Close()
                    Exit Sub
                End If
                '=================================================
                ' SPACE
                '=================================================
                If e.KeyCode = Keys.Space Then
                    e.SuppressKeyPress = True
                    e.Handled = True
                    Exit Sub
                End If
            Catch
            End Try
        End Sub
        '=========================================================
        '        ' PROCESS CMD KEY
        '        '=========================================================
        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
            Try
                '=================================================
                ' ESC = CLOSE
                '=================================================
                If keyData = Keys.Escape Then
                    Me.Close()
                    Return True
                End If
                '=================================================
                ' SPACE
                ' CHẶN INVENTOR COMMAND
                '=================================================
                If keyData = Keys.Space Then
                    Return True
                End If
                '=================================================
                ' ENTER
                '=================================================
                If keyData = Keys.Enter Then
                    '---------------------------------------------
                    ' CHỈ GO KHI ĐANG NHẬP TEXTBOX
                    '---------------------------------------------
                    If Me.ActiveControl Is txtPage Then
                        GoToSheetFromText()
                    End If
                    '---------------------------------------------
                    ' KHÔNG CHO ENTER GỌI LẠI COMMAND
                    '---------------------------------------------
                    Return True
                End If
            Catch
            End Try
            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function
        '=========================================================
        '         GO TO SHEET FROM TEXT
        '       =========================================================
        Private Sub GoToSheetFromText()
            Try
                Dim pageNumber As Integer
                If Not Integer.TryParse(txtPage.Text.Trim(), pageNumber) Then
                    MessageBox.Show("Nhập số Sheet hợp lệ.", "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtPage.SelectAll()
                    txtPage.Focus()
                    Exit Sub
                End If
                GoToSheet(pageNumber)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub
        '=========================================================
        ' FORM ACTIVATED
        '=========================================================
        Protected Overrides Sub OnActivated(ByVal e As EventArgs)
            MyBase.OnActivated(e)
            Try
                UpdateInfo()
            Catch
            End Try
        End Sub
        '=========================================================
        '        ' FORM CLOSED
        '        '=========================================================
        Protected Overrides Sub OnFormClosed(ByVal e As FormClosedEventArgs)
            Try
                isDragging = False
            Catch
            End Try
            MyBase.OnFormClosed(e)
        End Sub

        Private Sub InitializeComponent()
            Me.SuspendLayout()
            '
            'ThanhNSheetNavigatorForm
            '
            Me.ClientSize = New System.Drawing.Size(284, 261)
            Me.Name = "ThanhNSheetNavigatorForm"
            Me.ResumeLayout(False)

        End Sub

        Private Sub ThanhNSheetNavigatorForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        End Sub
    End Class
End Namespace