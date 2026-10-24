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

Namespace ToolInventor2025.Toolngoai

    Public Module delete_link_OLE

        ' ============================================================
        '  P/INVOKE – THAY THẾ Marshal.GetActiveObject CHO .NET 8
        '  Dùng System.Guid tường minh để tránh xung đột với stdole
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
        '  BIẾN MODULE-LEVEL CHO FORM
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

        ' ============================================================
        '  ENTRY POINT
        ' ============================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application
            Try
                invApp = GetInventorInstance()
            Catch ex As Exception
                MessageBox.Show("Không thể kết nối đến Inventor. Hãy chắc chắn Inventor đang chạy." & vbCrLf &
                                ex.Message, "Thanh N", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            Dim oDoc As Document = invApp.ActiveDocument

            If oDoc Is Nothing Then
                MessageBox.Show("Không có tài liệu đang mở!", "Thanh N",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If oDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject AndAlso
               oDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Chỉ hỗ trợ file Part (.ipt) hoặc Assembly (.iam)!", "Thanh N",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim countExcel As Integer = 0
            Dim countDerived As Integer = 0
            Dim countOLE As Integer = 0
            Dim countDocs As Integer = 0
            Dim modeText As String = ""

            '=================================================
            '  PART → XỬ LÝ NGAY
            '=================================================
            If oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                invApp.SilentOperation = True
                Try
                    RemoveAllParameterLinks(oDoc, countExcel, countDerived, countOLE)
                    countDocs = 1
                    oDoc.Update2(True)

                    MessageBox.Show("Đã xử lý xong Part!" & vbCrLf & vbCrLf &
                                    "Excel Tables đã xóa: " & countExcel & vbCrLf &
                                    "Derived Parameters đã xóa: " & countDerived & vbCrLf &
                                    "OLE Links đã xóa: " & countOLE,
                                    "Thanh N", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Lỗi: " & ex.Message, "Thanh N",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    invApp.SilentOperation = False
                End Try
                Return
            End If

            '=================================================
            '  ASSEMBLY → HIỆN FORM CHỌN
            '=================================================
            Dim selectedMode As Integer = ShowModeSelectionForm()
            If selectedMode = 0 Then Return

            Dim processTopOnly As Boolean = False
            Dim processOnlyAssemblies As Boolean = False
            Dim processOnlyParts As Boolean = False
            Dim processAll As Boolean = False

            Select Case selectedMode
                Case 1 : processTopOnly = True
                Case 2 : processAll = True
                Case 3 : processOnlyAssemblies = True
                Case 4 : processOnlyParts = True
            End Select

            invApp.SilentOperation = True
            Try
                Dim oAsmDoc As AssemblyDocument = CType(oDoc, AssemblyDocument)

                If processTopOnly Then
                    modeText = "Chỉ Top Level (Assembly tổng)"
                    RemoveAllParameterLinks(oAsmDoc, countExcel, countDerived, countOLE)
                    countDocs += 1
                    oAsmDoc.Update2(True)

                ElseIf processAll Then
                    modeText = "Đệ quy toàn bộ (Part + Assembly)"
                    For Each oRefDoc As Document In oAsmDoc.AllReferencedDocuments
                        Try
                            If oRefDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject OrElse
                               oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                                RemoveAllParameterLinks(oRefDoc, countExcel, countDerived, countOLE)
                                countDocs += 1
                                Try : oRefDoc.Update2(True) : Catch : End Try
                            End If
                        Catch
                        End Try
                    Next
                    RemoveAllParameterLinks(oAsmDoc, countExcel, countDerived, countOLE)
                    countDocs += 1
                    oAsmDoc.Update2(True)

                ElseIf processOnlyAssemblies Then
                    modeText = "Chỉ các Assembly (all level)"
                    For Each oRefDoc As Document In oAsmDoc.AllReferencedDocuments
                        Try
                            If oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                                RemoveAllParameterLinks(oRefDoc, countExcel, countDerived, countOLE)
                                countDocs += 1
                                Try : oRefDoc.Update2(True) : Catch : End Try
                            End If
                        Catch
                        End Try
                    Next
                    RemoveAllParameterLinks(oAsmDoc, countExcel, countDerived, countOLE)
                    countDocs += 1
                    oAsmDoc.Update2(True)

                ElseIf processOnlyParts Then
                    modeText = "Chỉ các Part (all level)"
                    For Each oRefDoc As Document In oAsmDoc.AllReferencedDocuments
                        Try
                            If oRefDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                                RemoveAllParameterLinks(oRefDoc, countExcel, countDerived, countOLE)
                                countDocs += 1
                                Try : oRefDoc.Update2(True) : Catch : End Try
                            End If
                        Catch
                        End Try
                    Next
                End If

                MessageBox.Show("Đã xử lý xong!" & vbCrLf & vbCrLf &
                                "Chế độ: " & modeText & vbCrLf &
                                "Số file đã xử lý: " & countDocs & vbCrLf &
                                "Excel Tables đã xóa: " & countExcel & vbCrLf &
                                "Derived Parameters đã xóa: " & countDerived & vbCrLf &
                                "OLE Links đã xóa: " & countOLE,
                                "Thanh N", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message, "Thanh N",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                invApp.SilentOperation = False
            End Try

        End Sub

        ' ============================================================
        '  FORM CHỌN CHẾ ĐỘ – GIAO DIỆN HIỆN ĐẠI
        ' ============================================================
        Private Sub FormBorderPaintHandler(sender As Object, e As PaintEventArgs)
            Dim f = DirectCast(sender, Form)
            e.Graphics.SmoothingMode = DrSmoothingMode.AntiAlias
            Using pen As New DrPen(DrColor.FromArgb(225, 229, 234), 1)   ' #E1E5EA
                Using p As DrGraphicsPath = GetRoundedRectPath(
                New DrRectangle(0, 0, f.Width - 1, f.Height - 1), 16)
                    e.Graphics.DrawPath(pen, p)
                End Using
            End Using
        End Sub

        Private Function ShowModeSelectionForm() As Integer

            _selectedMode = 0

            Dim frm As New ModernForm()   ' ← dùng class mới bên dưới
            frm.Text = "Xóa liên kết Parameters / OLE"
            frm.Size = New DrSize(480, 420)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.None
            frm.BackColor = DrColor.FromArgb(250, 251, 252)   ' #FAFBFC
            frm.Padding = New Padding(1)


            AddHandler frm.Paint, AddressOf FormBorderPaintHandler

            ' Cho phép đóng form bằng phím ESC
            frm.KeyPreview = True
            AddHandler frm.KeyDown, Sub(s, ev)
                                        If ev.KeyCode = Keys.Escape Then frm.Close()
                                    End Sub
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.None
            frm.BackColor = DrColor.FromArgb(250, 251, 252)   ' #FAFBFC – trắng ngà ấm
            frm.Padding = New Padding(1)

            Dim borderPath As DrGraphicsPath = GetRoundedRectPath(New DrRectangle(0, 0, frm.Width, frm.Height), 16)
            frm.Region = New DrRegion(borderPath)
            ' Vẽ viền mảnh bo góc (bắt buộc khi FormBorderStyle=None)
            AddHandler frm.Paint, AddressOf FormBorderPaintHandler




            ' Nút đóng ✕
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

            ' Tiêu đề
            Dim lblTitle As New Label()
            lblTitle.Text = "XÓA LIÊN KẾT THAM CHIẾU"
            lblTitle.Font = New DrFont("Segoe UI", 14, DrFontStyle.Bold)
            lblTitle.ForeColor = DrColor.FromArgb(33, 37, 41)
            lblTitle.Location = New DrPoint(24, 18)
            lblTitle.AutoSize = True
            frm.Controls.Add(lblTitle)

            Dim lblSub As New Label()
            lblSub.Text = "Chọn phạm vi áp dụng cho Assembly đang mở"
            lblSub.Font = New DrFont("Segoe UI", 9, DrFontStyle.Regular)
            lblSub.ForeColor = DrColor.FromArgb(108, 117, 125)
            lblSub.Location = New DrPoint(26, 48)
            lblSub.AutoSize = True
            frm.Controls.Add(lblSub)

            ' 4 nút chọn chế độ
            Dim btn1 As New ModernButton()
            btn1.SetContent("📄", "Chỉ Top Level", "Chỉ xử lý Assembly tổng (không đệ quy)")
            btn1.Location = New DrPoint(24, 80)
            btn1.Tag = 1
            AddHandler btn1.Click, AddressOf ModeButton_Click
            frm.Controls.Add(btn1)

            Dim btn2 As New ModernButton()
            btn2.SetContent("🔄", "Đệ quy toàn bộ", "Xử lý tất cả Part và Assembly con")
            btn2.Location = New DrPoint(24, 148)
            btn2.Tag = 2
            AddHandler btn2.Click, AddressOf ModeButton_Click
            frm.Controls.Add(btn2)

            Dim btn3 As New ModernButton()
            btn3.SetContent("🧩", "Chỉ các Assembly", "Chỉ xử lý các Assembly ở mọi cấp")
            btn3.Location = New DrPoint(24, 216)
            btn3.Tag = 3
            AddHandler btn3.Click, AddressOf ModeButton_Click
            frm.Controls.Add(btn3)

            Dim btn4 As New ModernButton()
            btn4.SetContent("⚙", "Chỉ các Part", "Chỉ xử lý các Part ở mọi cấp")
            btn4.Location = New DrPoint(24, 284)
            btn4.Tag = 4
            AddHandler btn4.Click, AddressOf ModeButton_Click
            frm.Controls.Add(btn4)

            ' Nút Hủy
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

            frm.ShowDialog()
            Return _selectedMode




        End Function
        ''' <summary>
        ''' Form có bóng đổ (drop shadow) khi FormBorderStyle=None
        ''' </summary>
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
        '  CONTROL NÚT HIỆN ĐẠI (TỰ VẼ)
        ' ============================================================
        Private Class ModernButton
            Inherits Control

            Private _icon As String = ""
            Private _title As String = ""
            Private _desc As String = ""
            Private _hover As Boolean = False

            Public Sub New()
                ' Cho phép Tab navigation và nhận focus bằng bàn phím
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

            ' Enter / Space kích hoạt click (giống Button)
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

                ' Nền + viền bo góc
                Using path As DrGraphicsPath = GetRoundedRectPath(rect, 8)
                    Dim backColor As DrColor = If(_hover,
                                           DrColor.FromArgb(239, 246, 255),   ' #EFF6FF
                                           DrColor.White)                     ' #FFFFFF
                    Dim borderColor As DrColor = If(_hover,
                                             DrColor.FromArgb(147, 197, 253), ' #93C5FD
                                             DrColor.FromArgb(229, 231, 235)) ' #E5E7EB
                    Using br As New DrSolidBrush(backColor)
                        g.FillPath(br, path)
                    End Using
                    Using pen As New DrPen(borderColor, 1)
                        g.DrawPath(pen, path)
                    End Using
                End Using

                ' Focus ring khi dùng bàn phím (Tab đến nút này)
                If Me.Focused Then
                    Using path As DrGraphicsPath = GetRoundedRectPath(
                    New DrRectangle(2, 2, Me.Width - 5, Me.Height - 5), 6)
                        Using pen As New DrPen(DrColor.FromArgb(59, 130, 246), 2)   ' #3B82F6
                            g.DrawPath(pen, path)
                        End Using
                    End Using
                End If

                ' Icon
                Using fIcon As New DrFont("Segoe UI Emoji", 16)
                    g.DrawString(_icon, fIcon, DrBrushes.Black, 12.0F, 14.0F)
                End Using

                ' Tiêu đề
                Using fTitle As New DrFont("Segoe UI", 10, DrFontStyle.Bold)
                    Using bTitle As New DrSolidBrush(DrColor.FromArgb(31, 41, 55))   ' #1F2937
                        g.DrawString(_title, fTitle, bTitle, 52.0F, 10.0F)
                    End Using
                End Using

                ' Mô tả
                Using fDesc As New DrFont("Segoe UI", 8, DrFontStyle.Regular)
                    Using bDesc As New DrSolidBrush(DrColor.FromArgb(107, 114, 128)) ' #6B7280
                        g.DrawString(_desc, fDesc, bDesc, 52.0F, 32.0F)
                    End Using
                End Using
            End Sub
        End Class

        ' ============================================================
        '  XỬ LÝ XÓA LIÊN KẾT
        ' ============================================================
        Private Sub RemoveAllParameterLinks(ByVal oDoc As Document,
                                            ByRef countExcel As Integer,
                                            ByRef countDerived As Integer,
                                            ByRef countOLE As Integer)

            Dim oParams As Parameters = Nothing
            Try
                oParams = oDoc.ComponentDefinition.Parameters
            Catch
                Return
            End Try
            If oParams Is Nothing Then Return

            ' 1. Xóa Excel Tables
            Try
                Dim oTables As ParameterTables = oParams.ParameterTables
                For i As Integer = oTables.Count To 1 Step -1
                    Try
                        oTables.Item(i).Delete()
                        countExcel += 1
                    Catch
                    End Try
                Next
            Catch
            End Try

            ' 2. Xóa Derived Parameters
            Try
                Dim oDerivedTables As DerivedParameterTables = oParams.DerivedParameterTables
                For i As Integer = oDerivedTables.Count To 1 Step -1
                    Try
                        Dim oTable As DerivedParameterTable = oDerivedTables.Item(i)
                        If Not oTable.HasReferenceComponent Then
                            oTable.Delete()
                            countDerived += 1
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            ' 3. Xóa OLE Links Excel
            Try
                For Each oOLE As ReferencedOLEFileDescriptor In oDoc.ReferencedOLEFileDescriptors
                    Try
                        Dim sName As String = oOLE.DisplayName.ToLower()
                        If sName.Contains(".xls") OrElse sName.Contains("excel") OrElse sName.Contains("embedding") Then
                            oOLE.Delete()
                            countOLE += 1
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

        End Sub

    End Module
End Namespace