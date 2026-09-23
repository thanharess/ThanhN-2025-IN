Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports Inventor
Imports System.Collections.Generic

'=====================================================
' KHAI BÁO WIN32 API
'=====================================================
Friend Class NativeMethods
    <DllImport("user32.dll")>
    Public Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function
End Class

Namespace ToolInventor2025.Drawing.Buttons.DrawView

    '=====================================================
    ' FORM CHỌN TỶ LỆ + LOẠI VIEW + PHẠM VI VIEW
    '=====================================================
    Public Class ScaleViewForm
        Inherits Form

        Private cboScale As ComboBox
        Private cboViewType As ComboBox
        Private txtCustom As System.Windows.Forms.TextBox
        Private chkAllViews As CheckBox
        Private btnOK As Button
        Private btnCancel As Button

        Public ReadOnly Property SelectedScale As Double
        Public ReadOnly Property SelectedViewFilter As String
        Public ReadOnly Property ApplyToAllViews As Boolean
        Public ReadOnly Property Cancelled As Boolean

        Private ReadOnly scaleList As String() = New String() {
            "2:1", "1:1", "1:2", "1:4", "1:5", "1:8", "1:10", "1:12.5", "1:15",
            "1:20", "1:25", "1:40", "1:50", "1:80", "1:100", "4:1", "5:1", "8:1", "10:1", "20:1"
        }

        Private ReadOnly viewTypeList As String() = New String() {
            "(Tất cả view)",
            "Standard", "Projected", "Section",
            "Detail", "Auxiliary", "Draft"
        }

        Public Sub New()
            _Cancelled = False
            _SelectedScale = 1.0
            _SelectedViewFilter = "(Tất cả view)"
            _ApplyToAllViews = True

            Me.Text = "Chọn tỷ lệ & View"
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ClientSize = New Size(280, 275)

            '=================================================
            ' TICK BOX NGOÀI CÙNG - PHẠM VI VIEW
            '=================================================
            chkAllViews = New CheckBox()
            chkAllViews.Text = "Áp dụng cho all View"
            chkAllViews.Location = New System.Drawing.Point(15, 12)
            chkAllViews.Size = New Size(250, 22)
            chkAllViews.Font = New Font(chkAllViews.Font, FontStyle.Regular) 'Bold)
            chkAllViews.Checked = True
            Me.Controls.Add(chkAllViews)

            Dim sep As New Label()
            sep.BorderStyle = BorderStyle.Fixed3D
            sep.Location = New System.Drawing.Point(15, 42)
            sep.Size = New Size(250, 2)
            Me.Controls.Add(sep)

            Dim lbl1 As New Label()
            lbl1.Text = "Chọn tỷ lệ có sẵn:"
            lbl1.Location = New System.Drawing.Point(15, 52)
            lbl1.Size = New Size(250, 20)
            Me.Controls.Add(lbl1)

            cboScale = New ComboBox()
            cboScale.DropDownStyle = ComboBoxStyle.DropDownList
            cboScale.Location = New System.Drawing.Point(15, 75)
            cboScale.Size = New Size(250, 25)
            For Each s As String In scaleList
                cboScale.Items.Add(s)
            Next
            cboScale.SelectedIndex = 2
            Me.Controls.Add(cboScale)

            Dim lbl2 As New Label()
            lbl2.Text = "Hoặc nhập tỷ lệ tùy ý (ví dụ 1/2, 1/10, 2):"
            lbl2.Location = New System.Drawing.Point(15, 112)
            lbl2.Size = New Size(250, 20)
            Me.Controls.Add(lbl2)

            txtCustom = New System.Windows.Forms.TextBox()
            txtCustom.Location = New System.Drawing.Point(15, 135)
            txtCustom.Size = New Size(250, 25)
            Me.Controls.Add(txtCustom)

            Dim lbl3 As New Label()
            lbl3.Text = "Loại View:"
            lbl3.Location = New System.Drawing.Point(15, 172)
            lbl3.Size = New Size(250, 20)
            lbl3.Font = New Font(lbl3.Font, FontStyle.Bold)
            Me.Controls.Add(lbl3)

            cboViewType = New ComboBox()
            cboViewType.DropDownStyle = ComboBoxStyle.DropDownList
            cboViewType.Location = New System.Drawing.Point(15, 195)
            cboViewType.Size = New Size(250, 25)
            For Each v As String In viewTypeList
                cboViewType.Items.Add(v)
            Next
            cboViewType.SelectedIndex = 0
            Me.Controls.Add(cboViewType)

            btnOK = New Button()
            btnOK.Text = "OK"
            btnOK.Location = New System.Drawing.Point(55, 235)
            btnOK.Size = New Size(80, 30)
            btnOK.DialogResult = DialogResult.OK
            Me.Controls.Add(btnOK)

            btnCancel = New Button()
            btnCancel.Text = "Hủy"
            btnCancel.Location = New System.Drawing.Point(150, 235)
            btnCancel.Size = New Size(80, 30)
            btnCancel.DialogResult = DialogResult.Cancel
            Me.Controls.Add(btnCancel)

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub

        Public Function ShowAndGet() As Boolean
            Do
                Dim result As DialogResult = Me.ShowDialog()
                If result <> DialogResult.OK Then
                    _Cancelled = True
                    Return False
                End If

                Dim input As String = txtCustom.Text.Trim()
                Dim valueToParse As String = If(input <> "", input, cboScale.SelectedItem.ToString())

                Dim s As Double
                If TryParseScale(valueToParse, s) Then
                    _SelectedScale = s
                    _SelectedViewFilter = cboViewType.SelectedItem.ToString()
                    _ApplyToAllViews = chkAllViews.Checked
                    Return True
                End If

                MessageBox.Show("Tỷ lệ không hợp lệ. Ví dụ hợp lệ: 1/2, 1, 2, 5/1.",
                                "Scale View", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Loop
        End Function

        Private Function TryParseScale(ByVal text As String, ByRef scale As Double) As Boolean
            Dim value As String = text.Trim().Replace(" ", "")
            Dim sep As Integer = value.IndexOf(":"c)
            If sep < 0 Then sep = value.IndexOf("/"c)

            If sep >= 0 Then
                Dim num As Double
                Dim den As Double
                If Not Double.TryParse(value.Substring(0, sep), num) OrElse
                   Not Double.TryParse(value.Substring(sep + 1), den) OrElse den = 0 Then Return False
                scale = num / den
            ElseIf Not Double.TryParse(value, scale) Then
                Return False
            End If
            Return scale > 0
        End Function
    End Class

    '=====================================================
    ' MODULE CHÍNH
    '=====================================================
    Public Module Draw_Change_scale_view

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                Dim invApp As Inventor.Application = g_inventorApplication

                If invApp Is Nothing Then
                    MessageBox.Show("Không tìm thấy Inventor Application.", "Scale View",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If invApp.ActiveDocument Is Nothing OrElse
                   invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Chức năng này chỉ dùng cho bản vẽ Drawing.", "Scale View",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                Dim drawingDocument As DrawingDocument =
                    CType(invApp.ActiveDocument, DrawingDocument)

                '===== HIỆN FORM CHỌN TỶ LỆ =====
                Dim form As New ScaleViewForm()
                If Not form.ShowAndGet() Then Exit Sub

                Dim scale As Double = form.SelectedScale
                Dim viewFilter As String = form.SelectedViewFilter
                Dim applyAllViews As Boolean = form.ApplyToAllViews

                '===== ĐƯA CỬA SỔ INVENTOR LÊN TRƯỚC =====
                Try
                    Dim mainHwnd As IntPtr = New IntPtr(invApp.MainFrameHWND)
                    If mainHwnd <> IntPtr.Zero Then
                        NativeMethods.SetForegroundWindow(mainHwnd)
                        System.Threading.Thread.Sleep(150)
                    End If
                Catch
                End Try

                '=====================================================
                ' LUÔN LÀM VIỆC TRÊN SHEET HIỆN TẠI
                '=====================================================
                Dim activeSheet As Sheet = drawingDocument.ActiveSheet

                '=====================================================
                ' XÁC ĐỊNH DANH SÁCH VIEW
                '=====================================================
                Dim selectedViews As New List(Of DrawingView)

                If applyAllViews Then
                    '----- Tất cả view trên sheet hiện tại -----
                    For Each v As DrawingView In activeSheet.DrawingViews
                        selectedViews.Add(v)
                    Next
                Else
                    '----- Chọn từng view bằng pick trên màn hình -----
                    Do
                        Dim oSS As SelectSet = drawingDocument.SelectSet
                        oSS.Clear()

                        Dim oView As DrawingView = Nothing
                        Try
                            oView = CType(
                                invApp.CommandManager.Pick(
                                    SelectionFilterEnum.kDrawingViewFilter,
                                    "Chọn View cần đổi tỷ lệ (Esc / Right-click để kết thúc)"),
                                DrawingView)
                        Catch
                            Exit Do
                        End Try

                        If oView Is Nothing Then Exit Do

                        Dim already As Boolean = False
                        For Each v As DrawingView In selectedViews
                            If v Is oView Then
                                already = True
                                Exit For
                            End If
                        Next

                        If Not already Then
                            selectedViews.Add(oView)
                        End If

                    Loop
                End If

                If selectedViews.Count = 0 Then
                    MessageBox.Show("Chưa chọn View nào.", "Scale View",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                '=====================================================
                ' ĐỔI TỶ LỆ CÁC VIEW ĐÃ CHỌN
                '=====================================================
                Dim okCount As Integer = 0
                Dim skipCount As Integer = 0
                Dim skipNames As New List(Of String)

                For Each drawingView As DrawingView In selectedViews

                    Try
                        Dim actualType As String = GetViewTypeName(drawingView)
                        If viewFilter <> "(Tất cả view)" AndAlso
                           Not String.Equals(actualType, viewFilter, StringComparison.OrdinalIgnoreCase) Then
                            skipCount += 1
                            skipNames.Add(drawingView.Name & " (sai loại: " & actualType & ")")
                            Continue For
                        End If

                        If drawingView.ParentView IsNot Nothing Then
                            skipCount += 1
                            skipNames.Add(drawingView.Name & " (view phụ thuộc)")
                            Continue For
                        End If

                        If drawingView.ScaleFromBase Then
                            Try
                                drawingView.ScaleFromBase = False
                            Catch
                                skipCount += 1
                                skipNames.Add(drawingView.Name & " (ScaleFromBase)")
                                Continue For
                            End Try
                        End If

                        drawingView.Scale = scale
                        okCount += 1

                    Catch ex As Exception
                        skipCount += 1
                        skipNames.Add(drawingView.Name & " (" & ex.Message & ")")
                    End Try

                Next

                drawingDocument.Update()

                '===== THÔNG BÁO KẾT QUẢ =====
                Dim scopeView As String = If(applyAllViews, "Tất cả View", "Chọn thủ công")

                Dim msg As String =
                    "Sheet: " & activeSheet.Name & vbCrLf &
                    "Phạm vi: " & scopeView & vbCrLf &
                    "Đã đổi tỷ lệ: " & okCount & " view" & vbCrLf &
                    "Bỏ qua: " & skipCount & " view"

                If skipNames.Count > 0 Then
                    msg &= vbCrLf & vbCrLf & "Chi tiết bỏ qua:" & vbCrLf &
                           String.Join(vbCrLf, skipNames.ToArray())
                End If

                MessageBox.Show(msg, "Scale View",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Không thể đổi tỷ lệ view:" & vbCrLf & ex.Message,
                                "Scale View", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        Private Function GetViewTypeName(dv As DrawingView) As String
            Try
                Select Case dv.ViewType
                    Case DrawingViewTypeEnum.kStandardDrawingViewType
                        Return "Standard"
                    Case DrawingViewTypeEnum.kProjectedDrawingViewType
                        Return "Projected"
                    Case DrawingViewTypeEnum.kSectionDrawingViewType
                        Return "Section"
                    Case DrawingViewTypeEnum.kDetailDrawingViewType
                        Return "Detail"
                    Case DrawingViewTypeEnum.kAuxiliaryDrawingViewType
                        Return "Auxiliary"
                    Case DrawingViewTypeEnum.kDraftDrawingViewType
                        Return "Draft"
                    Case Else
                        Return "Unknown"
                End Select
            Catch
                Return "Unknown"
            End Try
        End Function

    End Module
End Namespace