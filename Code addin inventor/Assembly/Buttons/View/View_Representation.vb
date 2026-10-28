Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Inventor
' Alias để tránh xung đột System.Drawing <-> Inventor
Imports DrColor = System.Drawing.Color
Imports DrPoint = System.Drawing.Point
Imports DrSize = System.Drawing.Size
Imports DrRectangle = System.Drawing.Rectangle
Imports DrFont = System.Drawing.Font
Imports DrFontStyle = System.Drawing.FontStyle
Imports DrGraphics = System.Drawing.Graphics
Imports DrGraphicsPath = System.Drawing.Drawing2D.GraphicsPath
Imports DrSolidBrush = System.Drawing.SolidBrush
Imports DrPen = System.Drawing.Pen
Imports DrBrushes = System.Drawing.Brushes
Imports DrRegion = System.Drawing.Region
Imports DrContentAlignment = System.Drawing.ContentAlignment
Imports DrSmoothingMode = System.Drawing.Drawing2D.SmoothingMode
Imports DrTextRenderingHint = System.Drawing.Text.TextRenderingHint

Namespace ToolInventor2025.Assembly.Buttons.ASS_View
    Public Module View_Representation

        ' ============================================================
        '  P/INVOKE – THAY THẾ Marshal.GetActiveObject CHO .NET 8
        ' ============================================================
        <DllImport("ole32.dll", CharSet:=CharSet.Unicode, PreserveSig:=True)>
        Private Function CLSIDFromProgIDEx(<MarshalAs(UnmanagedType.LPWStr)> progId As String,
                                            ByRef clsid As System.Guid) As Integer
        End Function

        <DllImport("ole32.dll", CharSet:=CharSet.Unicode, PreserveSig:=True)>
        Private Function CLSIDFromProgID(<MarshalAs(UnmanagedType.LPWStr)> progId As String,
                                          ByRef clsid As System.Guid) As Integer
        End Function

        <DllImport("oleaut32.dll", PreserveSig:=True)>
        Private Function GetActiveObject(ByRef rclsid As System.Guid,
                                          reserved As IntPtr,
                                          <MarshalAs(UnmanagedType.IUnknown)> ByRef ppunk As Object) As Integer
        End Function

        Private Function GetInventorInstance() As Inventor.Application
            Dim clsid As System.Guid
            Dim hr As Integer
            Try
                hr = CLSIDFromProgIDEx("Inventor.Application", clsid)
            Catch
                hr = -1
            End Try
            If hr <> 0 Then hr = CLSIDFromProgID("Inventor.Application", clsid)
            If hr <> 0 Then Throw New Exception("Không tìm thấy CLSID của Inventor.Application")
            Dim obj As Object = Nothing
            hr = GetActiveObject(clsid, IntPtr.Zero, obj)
            If hr <> 0 Then Throw New Exception("Không tìm thấy phiên bản Inventor đang chạy")
            Return CType(obj, Inventor.Application)
        End Function

        ' ============================================================
        '  BIẾN MODULE-LEVEL
        ' ============================================================
        Private _selectedMode As Integer = 0

        Private Sub ModeButton_Click(sender As Object, e As EventArgs)
            Dim ctl As Control = DirectCast(sender, Control)
            _selectedMode = CInt(ctl.Tag)
            ctl.FindForm().Close()
        End Sub

        Private Sub CloseForm(sender As Object, e As EventArgs)
            DirectCast(sender, Control).FindForm().Close()
        End Sub

        Private Sub CloseBtn_MouseEnter(sender As Object, e As EventArgs)
            DirectCast(sender, Label).ForeColor = DrColor.FromArgb(220, 53, 69)
        End Sub

        Private Sub CloseBtn_MouseLeave(sender As Object, e As EventArgs)
            DirectCast(sender, Label).ForeColor = DrColor.FromArgb(173, 181, 189)
        End Sub

        Private Sub FormBorderPaintHandler(sender As Object, e As PaintEventArgs)
            Dim f = DirectCast(sender, Form)
            e.Graphics.SmoothingMode = DrSmoothingMode.AntiAlias
            Using pen As New DrPen(DrColor.FromArgb(225, 229, 234), 1)
                Using p As DrGraphicsPath = GetRoundedRectPath(
                        New DrRectangle(0, 0, f.Width - 1, f.Height - 1), 16)
                    e.Graphics.DrawPath(pen, p)
                End Using
            End Using
        End Sub

        ' ============================================================
        '  HELPER DÙNG CHUNG CHO CÁC FORM
        ' ============================================================
        Private Sub PrepareModernForm(frm As Form, width As Integer, height As Integer)
            frm.Size = New DrSize(width, height)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.None
            frm.BackColor = DrColor.FromArgb(250, 251, 252)
            frm.Padding = New Padding(1)
            frm.TopMost = True
            Using borderPath As DrGraphicsPath = GetRoundedRectPath(
                    New DrRectangle(0, 0, frm.Width, frm.Height), 16)
                frm.Region = New DrRegion(borderPath)
            End Using
            AddHandler frm.Paint, AddressOf FormBorderPaintHandler
            EnableFormDrag(frm)
        End Sub

        Private Sub EnableFormDrag(frm As Form)
            Dim dragOffset As DrPoint = DrPoint.Empty
            Dim isDragging As Boolean = False
            AddHandler frm.MouseDown, Sub(s, ev)
                                          If ev.Button = MouseButtons.Left Then
                                              isDragging = True
                                              dragOffset = ev.Location
                                          End If
                                      End Sub
            AddHandler frm.MouseMove, Sub(s, ev)
                                          If isDragging Then
                                              frm.Location = New DrPoint(
                                                  frm.Location.X + ev.X - dragOffset.X,
                                                  frm.Location.Y + ev.Y - dragOffset.Y)
                                          End If
                                      End Sub
            AddHandler frm.MouseUp, Sub(s, ev)
                                        isDragging = False
                                    End Sub
        End Sub

        Private Sub AddCloseXButton(frm As Form)
            Dim btnClose As New Label()
            btnClose.Text = "✕"
            btnClose.Font = New DrFont("Segoe UI", 12, DrFontStyle.Bold)
            btnClose.ForeColor = DrColor.FromArgb(173, 181, 189)
            btnClose.Location = New DrPoint(frm.Width - 40, 12)
            btnClose.Size = New DrSize(28, 28)
            btnClose.TextAlign = DrContentAlignment.MiddleCenter
            btnClose.Cursor = Cursors.Hand
            AddHandler btnClose.Click, AddressOf CloseForm
            AddHandler btnClose.MouseEnter, AddressOf CloseBtn_MouseEnter
            AddHandler btnClose.MouseLeave, AddressOf CloseBtn_MouseLeave
            frm.Controls.Add(btnClose)
        End Sub

        Private Function AddCancelButton(frm As Form) As Button
            Dim btnCancel As New Button()
            btnCancel.Text = "Hủy bỏ"
            btnCancel.Font = New DrFont("Segoe UI", 9)
            btnCancel.ForeColor = DrColor.FromArgb(73, 80, 87)
            btnCancel.BackColor = DrColor.FromArgb(233, 236, 239)
            btnCancel.FlatStyle = FlatStyle.Flat
            btnCancel.FlatAppearance.BorderSize = 0
            btnCancel.Size = New DrSize(100, 32)
            btnCancel.Location = New DrPoint(frm.Width - 124, frm.Height - 52)
            btnCancel.Cursor = Cursors.Hand
            AddHandler btnCancel.Click, AddressOf CloseForm
            frm.Controls.Add(btnCancel)
            Return btnCancel
        End Function

        Private Sub AddTitleLabels(frm As Form, title As String, subtitle As String)
            Dim lblTitle As New Label()
            lblTitle.Text = title
            lblTitle.Font = New DrFont("Segoe UI", 14, DrFontStyle.Bold)
            lblTitle.ForeColor = DrColor.FromArgb(33, 37, 41)
            lblTitle.Location = New DrPoint(24, 18)
            lblTitle.AutoSize = True
            frm.Controls.Add(lblTitle)

            Dim lblSub As New Label()
            lblSub.Text = subtitle
            lblSub.Font = New DrFont("Segoe UI", 9, DrFontStyle.Regular)
            lblSub.ForeColor = DrColor.FromArgb(108, 117, 125)
            lblSub.Location = New DrPoint(26, 48)
            lblSub.AutoSize = True
            frm.Controls.Add(lblSub)
        End Sub

        ' ============================================================
        '  ENTRY POINT
        ' ============================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim invApp As Inventor.Application
            Try
                invApp = GetInventorInstance()
            Catch ex As Exception
                MessageBox.Show("Không thể kết nối đến Inventor. Hãy chắc chắn Inventor đang chạy." & vbCrLf &
                                ex.Message, "Thanh N",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            Dim oDoc As Document = invApp.ActiveDocument
            If oDoc Is Nothing OrElse oDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Chỉ hỗ trợ file Assembly (.iam)!", "Thanh N",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim oAsmDoc As AssemblyDocument = CType(oDoc, AssemblyDocument)

            '=================================================
            ' BƯỚC 1: CHỌN CHỨC NĂNG
            '=================================================
            Dim selectedFunc As Integer = ShowFunctionForm()
            If selectedFunc = 0 Then Return

            '=================================================
            ' BƯỚC 2: CHỌN PHẠM VI
            '=================================================
            Dim selectedScope As Integer = ShowScopeForm()
            If selectedScope = 0 Then Return

            Dim isAllLevel As Boolean = (selectedScope = 2)

            invApp.SilentOperation = True
            Try
                Select Case selectedFunc
                    Case 1   ' Kích hoạt theo tên → cần SetOccurrence
                        Dim viewName As String = ShowInputForm(
                            "Kích hoạt View Representation",
                            "Nhập tên View Representation cần kích hoạt:",
                            "[Primary]")
                        If String.IsNullOrEmpty(viewName) Then Return

                        If isAllLevel Then
                            ActivateView_AllLevel(oAsmDoc, viewName, True)
                        Else
                            ActivateView_TopLevel(oAsmDoc, viewName)
                        End If

                    Case 2   ' [Primary] → KHÔNG cần SetOccurrence
                        If isAllLevel Then
                            ActivateView_AllLevel(oAsmDoc, "[Primary]", False)
                        Else
                            ActivateView_TopLevel(oAsmDoc, "[Primary]")
                        End If

                    Case 3   ' Default → cần SetOccurrence
                        If isAllLevel Then
                            ActivateView_AllLevel(oAsmDoc, "Default", True)
                        Else
                            ActivateView_TopLevel(oAsmDoc, "Default")
                        End If
                End Select
            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message, "Thanh N",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                invApp.SilentOperation = False
            End Try
        End Sub

        '=================================================
        ' KÍCH HOẠT VIEW - TOP LEVEL
        '=================================================
        Private Sub ActivateView_TopLevel(ByVal oAsmDoc As AssemblyDocument, ByVal viewName As String)
            Dim found As Boolean = ActivateViewInDoc(oAsmDoc, viewName)
            If found Then
                MessageBox.Show("Đã kích hoạt View """ & viewName & """ (Top Level)",
                                "Thanh N", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Không tìm thấy View: " & viewName & vbCrLf &
                                "Đã thử fallback [Primary] nếu là Default.",
                                "Thanh N", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub

        '=================================================
        ' KÍCH HOẠT VIEW - ALL LEVEL
        ' needSetOccurrence = True  → chức năng 1 & 3
        ' needSetOccurrence = False → chức năng 2 ([Primary])
        '=================================================
        Private Sub ActivateView_AllLevel(ByVal oAsmDoc As AssemblyDocument,
                                          ByVal viewName As String,
                                          ByVal needSetOccurrence As Boolean)
            Dim countSuccess As Integer = 0
            Dim countFail As Integer = 0

            ProcessAssemblyRecursive(oAsmDoc, viewName, needSetOccurrence, countSuccess, countFail)

            MessageBox.Show("Đã xử lý xong (All Level)!" & vbCrLf & vbCrLf &
                            "Thành công: " & countSuccess & vbCrLf &
                            "Không tìm thấy / Lỗi: " & countFail & vbCrLf & vbCrLf &
                            If(needSetOccurrence,
                               "Đã SetDesignViewRepresentation (Associative) trên tất cả occurrence.",
                               "Chỉ kích hoạt View [Primary] (không cần SetOccurrence)."),
                            "Thanh N", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        '=================================================
        ' ĐỆ QUY THEO CÂY OCCURRENCE
        '=================================================
        Private Sub ProcessAssemblyRecursive(ByVal oAsmDoc As AssemblyDocument,
                                             ByVal viewName As String,
                                             ByVal needSetOccurrence As Boolean,
                                             ByRef countSuccess As Integer,
                                             ByRef countFail As Integer)

            ' 1. Activate trong document hiện tại
            If ActivateViewInDoc(oAsmDoc, viewName) Then
                countSuccess += 1
            Else
                countFail += 1
            End If

            ' 2. Duyệt tất cả occurrence
            For Each oOcc As ComponentOccurrence In oAsmDoc.ComponentDefinition.Occurrences
                Try
                    If oOcc.Suppressed Then Continue For
                    If TypeOf oOcc.Definition Is VirtualComponentDefinition Then Continue For
                    If TypeOf oOcc.Definition Is WeldsComponentDefinition Then Continue For

                    ' Chỉ xử lý Sub-Assembly
                    If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then

                        ' === CHỈ chạy đoạn này khi needSetOccurrence = True (chức năng 1 & 3) ===
                        If needSetOccurrence Then
                            Try
                                oOcc.SetDesignViewRepresentation(viewName, , True)
                            Catch
                                ' Fallback nếu viewName không tồn tại trong sub
                                Try
                                    oOcc.SetDesignViewRepresentation("Default", , True)
                                Catch
                                    Try
                                        oOcc.SetDesignViewRepresentation("[Primary]", , True)
                                    Catch
                                    End Try
                                End Try
                            End Try
                        End If

                        ' Đệ quy vào sub-assembly
                        Dim subAsm As AssemblyDocument = CType(oOcc.Definition.Document, AssemblyDocument)
                        ProcessAssemblyRecursive(subAsm, viewName, needSetOccurrence, countSuccess, countFail)
                    End If
                Catch
                    countFail += 1
                End Try
            Next

            Try
                oAsmDoc.Update2(True)
            Catch
            End Try
        End Sub

        '=================================================
        ' HÀM PHỤ: KÍCH HOẠT VIEW TRONG 1 ASSEMBLY DOCUMENT
        '=================================================
        Private Function ActivateViewInDoc(ByVal oAsmDoc As AssemblyDocument, ByVal viewName As String) As Boolean
            Try
                Dim oRepMgr As RepresentationsManager = oAsmDoc.ComponentDefinition.RepresentationsManager

                ' Tìm đúng tên (không phân biệt hoa thường)
                For Each oView As DesignViewRepresentation In oRepMgr.DesignViewRepresentations
                    If String.Compare(oView.Name, viewName, True) = 0 Then
                        oView.Activate()
                        Return True
                    End If
                Next

                ' Fallback: Default → [Primary]
                If String.Compare(viewName, "Default", True) = 0 Then
                    For Each oView As DesignViewRepresentation In oRepMgr.DesignViewRepresentations
                        If String.Compare(oView.Name, "[Primary]", True) = 0 Then
                            oView.Activate()
                            Return True
                        End If
                    Next
                End If
            Catch
            End Try
            Return False
        End Function

        ' ============================================================
        '  FORM CHỌN CHỨC NĂNG
        ' ============================================================
        Private Function ShowFunctionForm() As Integer
            _selectedMode = 0
            Dim frm As New ModernForm()
            frm.Text = "View Representation"
            PrepareModernForm(frm, 480, 360)
            AddCloseXButton(frm)
            AddTitleLabels(frm, "🎨  VIEW REPRESENTATION", "Chọn chức năng cần thực hiện")

            Dim btn1 As New ModernButton()
            btn1.SetContent("🎯", "Kích hoạt theo tên", "Nhập tên View cần kích hoạt")
            btn1.Location = New DrPoint(24, 80)
            btn1.Tag = 1
            AddHandler btn1.Click, AddressOf ModeButton_Click
            frm.Controls.Add(btn1)

            Dim btn2 As New ModernButton()
            btn2.SetContent("🏠", "Về View [Primary]", "Chuyển Assembly về View [Primary]")
            btn2.Location = New DrPoint(24, 148)
            btn2.Tag = 2
            AddHandler btn2.Click, AddressOf ModeButton_Click
            frm.Controls.Add(btn2)

            Dim btn3 As New ModernButton()
            btn3.SetContent("📋", "Về View Default", "Chuyển về Default (tự fallback [Primary])")
            btn3.Location = New DrPoint(24, 216)
            btn3.Tag = 3
            AddHandler btn3.Click, AddressOf ModeButton_Click
            frm.Controls.Add(btn3)

            Dim btnCancel As Button = AddCancelButton(frm)
            frm.CancelButton = btnCancel
            frm.ActiveControl = btn1
            frm.ShowDialog()
            Return _selectedMode
        End Function

        ' ============================================================
        '  FORM CHỌN PHẠM VI
        ' ============================================================
        Private Function ShowScopeForm() As Integer
            _selectedMode = 0
            Dim frm As New ModernForm()
            frm.Text = "Chọn phạm vi áp dụng"
            PrepareModernForm(frm, 480, 290)
            AddCloseXButton(frm)
            AddTitleLabels(frm, "📐  CHỌN PHẠM VI", "Áp dụng cho Assembly hiện tại")

            Dim btn1 As New ModernButton()
            btn1.SetContent("📄", "Chỉ Top Level", "Chỉ Assembly tổng đang mở")
            btn1.Location = New DrPoint(24, 80)
            btn1.Tag = 1
            AddHandler btn1.Click, AddressOf ModeButton_Click
            frm.Controls.Add(btn1)

            Dim btn2 As New ModernButton()
            btn2.SetContent("🔄", "All Level", "Đệ quy toàn bộ + Set Occurrence (Associative)")
            btn2.Location = New DrPoint(24, 148)
            btn2.Tag = 2
            AddHandler btn2.Click, AddressOf ModeButton_Click
            frm.Controls.Add(btn2)

            Dim btnCancel As Button = AddCancelButton(frm)
            frm.CancelButton = btnCancel
            frm.ActiveControl = btn1
            frm.ShowDialog()
            Return _selectedMode
        End Function

        ' ============================================================
        '  FORM NHẬP CHUỖI – THAY THẾ InputBox
        ' ============================================================
        Private Function ShowInputForm(title As String, prompt As String, defaultValue As String) As String
            Dim result As String = Nothing
            Dim frm As New ModernForm()
            frm.Text = title
            PrepareModernForm(frm, 440, 210)
            AddCloseXButton(frm)
            AddTitleLabels(frm, "✏️  " & title.ToUpper(), "Nhập thông tin bên dưới")

            Dim lblPrompt As New Label()
            lblPrompt.Text = prompt
            lblPrompt.Font = New DrFont("Segoe UI", 9)
            lblPrompt.ForeColor = DrColor.FromArgb(75, 85, 99)
            lblPrompt.Location = New DrPoint(24, 84)
            lblPrompt.Size = New DrSize(392, 22)
            frm.Controls.Add(lblPrompt)

            Dim txt As New System.Windows.Forms.TextBox()
            txt.Text = defaultValue
            txt.Font = New DrFont("Segoe UI", 10)
            txt.Location = New DrPoint(24, 108)
            txt.Size = New DrSize(392, 28)
            txt.BorderStyle = BorderStyle.FixedSingle
            frm.Controls.Add(txt)

            Dim btnOK As New Button()
            btnOK.Text = "Đồng ý"
            btnOK.Font = New DrFont("Segoe UI", 9)
            btnOK.ForeColor = DrColor.White
            btnOK.BackColor = DrColor.FromArgb(13, 110, 253)
            btnOK.FlatStyle = FlatStyle.Flat
            btnOK.FlatAppearance.BorderSize = 0
            btnOK.Size = New DrSize(90, 32)
            btnOK.Location = New DrPoint(frm.Width - 214, frm.Height - 52)
            btnOK.Cursor = Cursors.Hand
            AddHandler btnOK.Click, Sub()
                                        result = txt.Text
                                        frm.Close()
                                    End Sub
            frm.Controls.Add(btnOK)

            Dim btnCancel As Button = AddCancelButton(frm)
            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel
            frm.ActiveControl = txt
            frm.ShowDialog()
            Return result
        End Function

        ' ============================================================
        '  HÀM VẼ BO GÓC
        ' ============================================================
        Private Function GetRoundedRectPath(rect As DrRectangle, radius As Integer) As DrGraphicsPath
            Dim path As New DrGraphicsPath()
            Dim d As Integer = radius * 2
            path.AddArc(rect.X, rect.Y, d, d, 180, 90)
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90)
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90)
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90)
            path.CloseFigure()
            Return path
        End Function

        ' ============================================================
        '  FORM CÓ BÓNG ĐỔ
        ' ============================================================
        Private Class ModernForm
            Inherits Form
            Protected Overrides ReadOnly Property CreateParams As CreateParams
                Get
                    Dim cp As CreateParams = MyBase.CreateParams
                    cp.ClassStyle = cp.ClassStyle Or &H20000   ' CS_DROPSHADOW
                    Return cp
                End Get
            End Property
        End Class

        ' ============================================================
        '  NÚT HIỆN ĐẠI (TỰ VẼ)
        ' ============================================================
        Private Class ModernButton
            Inherits Control
            Private _icon As String = ""
            Private _title As String = ""
            Private _desc As String = ""
            Private _hover As Boolean = False

            Public Sub New()
                Me.SetStyle(ControlStyles.Selectable Or
                            ControlStyles.UserPaint Or
                            ControlStyles.AllPaintingInWmPaint Or
                            ControlStyles.OptimizedDoubleBuffer, True)
                Me.TabStop = True
                Me.Cursor = Cursors.Hand
            End Sub

            Public Sub SetContent(icon As String, title As String, desc As String)
                _icon = icon
                _title = title
                _desc = desc
                Me.Size = New DrSize(432, 56)
            End Sub

            Protected Overrides Sub OnMouseEnter(e As EventArgs)
                _hover = True
                Invalidate()
                MyBase.OnMouseEnter(e)
            End Sub

            Protected Overrides Sub OnMouseLeave(e As EventArgs)
                _hover = False
                Invalidate()
                MyBase.OnMouseLeave(e)
            End Sub

            Protected Overrides Function IsInputKey(keyData As Keys) As Boolean
                If keyData = Keys.Enter OrElse keyData = Keys.Space Then Return True
                Return MyBase.IsInputKey(keyData)
            End Function

            Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
                If e.KeyCode = Keys.Enter OrElse e.KeyCode = Keys.Space Then
                    Me.OnClick(EventArgs.Empty)
                    e.Handled = True
                End If
                MyBase.OnKeyDown(e)
            End Sub

            Protected Overrides Sub OnPaint(e As PaintEventArgs)
                Dim g As DrGraphics = e.Graphics
                g.SmoothingMode = DrSmoothingMode.AntiAlias
                g.TextRenderingHint = DrTextRenderingHint.ClearTypeGridFit

                Dim rect As New DrRectangle(0, 0, Me.Width - 1, Me.Height - 1)
                Using path As DrGraphicsPath = GetRoundedRectPath(rect, 8)
                    Dim backColor As DrColor = If(_hover,
                                                   DrColor.FromArgb(239, 246, 255),
                                                   DrColor.White)
                    Dim borderColor As DrColor = If(_hover,
                                                     DrColor.FromArgb(147, 197, 253),
                                                     DrColor.FromArgb(229, 231, 235))
                    Using br As New DrSolidBrush(backColor)
                        g.FillPath(br, path)
                    End Using
                    Using pen As New DrPen(borderColor, 1)
                        g.DrawPath(pen, path)
                    End Using
                End Using

                If Me.Focused Then
                    Using path As DrGraphicsPath = GetRoundedRectPath(
                            New DrRectangle(2, 2, Me.Width - 5, Me.Height - 5), 6)
                        Using pen As New DrPen(DrColor.FromArgb(59, 130, 246), 2)
                            g.DrawPath(pen, path)
                        End Using
                    End Using
                End If

                Using fIcon As New DrFont("Segoe UI Emoji", 16)
                    g.DrawString(_icon, fIcon, DrBrushes.Black, 12.0F, 14.0F)
                End Using

                Using fTitle As New DrFont("Segoe UI", 10, DrFontStyle.Bold)
                    Using bTitle As New DrSolidBrush(DrColor.FromArgb(31, 41, 55))
                        g.DrawString(_title, fTitle, bTitle, 52.0F, 10.0F)
                    End Using
                End Using

                Using fDesc As New DrFont("Segoe UI", 8, DrFontStyle.Regular)
                    Using bDesc As New DrSolidBrush(DrColor.FromArgb(107, 114, 128))
                        g.DrawString(_desc, fDesc, bDesc, 52.0F, 32.0F)
                    End Using
                End Using
            End Sub
        End Class

    End Module
End Namespace