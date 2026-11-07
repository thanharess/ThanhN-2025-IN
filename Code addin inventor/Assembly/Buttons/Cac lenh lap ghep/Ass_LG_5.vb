Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor
Imports System.IO
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Threading

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep
    Public Module Ass_Part_Replace

        ' ====== Win32 API: đưa Inventor lên foreground để Pick chạy được ======
        <DllImport("user32.dll")>
        Private Function SetForegroundWindow(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
        End Function

        Private Const SW_RESTORE As Integer = 9

        ' ====== API helper: đưa Inventor lên trước ======
        Private Sub BringInventorToFront(invApp As Inventor.Application)
            Try
                Dim hwnd As New IntPtr(CLng(invApp.MainFrameHWND))
                ShowWindow(hwnd, SW_RESTORE)
                SetForegroundWindow(hwnd)
                Thread.Sleep(250)
            Catch
            End Try
        End Sub

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim invApp As Inventor.Application
            Try

                invApp = Interop.Marshal2.GetActiveObject("Inventor.Application")
            Catch
                MessageBox.Show("Inventor chưa chạy.", "Replace Component")
                Return
            End Try

            Dim oDoc As AssemblyDocument = TryCast(invApp.ActiveEditDocument, AssemblyDocument)
            If oDoc Is Nothing Then
                MessageBox.Show("Hãy mở / edit trong Assembly trước.", "Replace Component")
                Return
            End If

            '=================================================
            ' FORM CHỌN CHẾ ĐỘ
            '=================================================
            Dim mode As Integer = ShowModeForm()
            If mode = 0 Then Return

            '=================================================
            ' BƯỚC 1: Lấy danh sách component cần thay
            '=================================================
            Dim listOccOld As List(Of ComponentOccurrence)
            If mode = 3 OrElse mode = 4 Then
                listOccOld = CollectMultiOccurrences(invApp, oDoc)
            Else
                listOccOld = CollectSingleOccurrence(invApp, oDoc, mode)
            End If
            If listOccOld Is Nothing OrElse listOccOld.Count = 0 Then Return

            '=================================================
            ' BƯỚC 2: Chọn component mới
            ' (MessageBox tự ấn OK sau 120ms — không cần user bấm)
            '=================================================
            Dim tmrAutoOK As New System.Windows.Forms.Timer()
            tmrAutoOK.Interval = 120
            AddHandler tmrAutoOK.Tick, Sub(s, e)
                                           tmrAutoOK.Stop()
                                           tmrAutoOK.Dispose()
                                           SendKeys.SendWait("{ENTER}")
                                       End Sub
            tmrAutoOK.Start()

            MessageBox.Show("Bước 2: Click chọn 1 ASS/Part trong cửa sổ Inventor" & vbCrLf &
                "để lấy file mới dùng cho việc thay thế.",
                "Replace Component", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Đảm bảo Inventor là cửa sổ active để Pick chạy được
            Try
                oDoc.Activate()
                System.Threading.Thread.Sleep(120)
            Catch
            End Try

            Dim occNew As ComponentOccurrence = invApp.CommandManager.Pick(
    SelectionFilterEnum.kAssemblyOccurrenceFilter,
    "Bước 2: Click chọn ASS/Part mới (lấy file)...")

            If occNew Is Nothing Then Return

            occNew = ResolveOccurrence(occNew)
            If occNew Is Nothing Then
                MessageBox.Show("Không lấy được component mới hợp lệ.", "Replace Component")
                Return
            End If

            Dim newFile As String = ""
            Try
                newFile = occNew.Definition.Document.FullFileName
            Catch ex As Exception
                MessageBox.Show("Không lấy được file mới." & vbCrLf & ex.Message)
                Return
            End Try

            If String.IsNullOrEmpty(newFile) OrElse Not System.IO.File.Exists(newFile) Then
                MessageBox.Show("File mới không hợp lệ.", "Replace Component")
                Return
            End If

            '=================================================
            ' THỰC HIỆN REPLACE
            '=================================================
            Dim countSuccess As Integer = 0
            Dim countFail As Integer = 0
            Try
                oDoc.SelectSet.Clear()
                Select Case mode
                    Case 1  ' REPLACE ALL
                        Dim firstOcc As ComponentOccurrence = listOccOld(0)
                        Dim oldFile As String = SafeGetFile(firstOcc)
                        If String.IsNullOrEmpty(oldFile) Then
                            MessageBox.Show("Không lấy được file cũ.", "Replace Component")
                            Return
                        End If
                        If String.Equals(oldFile, newFile, StringComparison.OrdinalIgnoreCase) Then
                            MessageBox.Show("File mới trùng file cũ.", "Replace Component")
                            Return
                        End If
                        firstOcc.Replace(newFile, True)
                        countSuccess = 1
                        oDoc.Update2(True)

                    Case 2  ' REPLACE 1
                        Dim occ As ComponentOccurrence = listOccOld(0)
                        Dim oldFile As String = SafeGetFile(occ)
                        If String.IsNullOrEmpty(oldFile) Then
                            MessageBox.Show("Không lấy được file cũ.", "Replace Component")
                            Return
                        End If
                        If String.Equals(oldFile, newFile, StringComparison.OrdinalIgnoreCase) Then
                            MessageBox.Show("File mới trùng file cũ.", "Replace Component")
                            Return
                        End If
                        occ.Replace(newFile, False)
                        countSuccess = 1
                        oDoc.Update2(True)

                    Case 3  ' REPLACE NHIỀU
                        For Each occ As ComponentOccurrence In listOccOld
                            Try
                                If occ.Suppressed Then Continue For
                                Dim currentFile As String = SafeGetFile(occ)
                                If String.IsNullOrEmpty(currentFile) Then Continue For
                                If String.Equals(currentFile, newFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                                occ.Replace(newFile, False)
                                countSuccess += 1
                            Catch
                                countFail += 1
                            End Try
                        Next
                        oDoc.Update2(True)

                    Case 4  ' REPLACE ALL MULTI
                        Dim processedFiles As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)
                        For Each occ As ComponentOccurrence In listOccOld
                            Try
                                If occ.Suppressed Then Continue For
                                Dim currentFile As String = SafeGetFile(occ)
                                If String.IsNullOrEmpty(currentFile) Then Continue For
                                If String.Equals(currentFile, newFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                                If processedFiles.ContainsKey(currentFile) Then Continue For
                                occ.Replace(newFile, True)
                                processedFiles.Add(currentFile, True)
                                countSuccess += 1
                            Catch
                                countFail += 1
                            End Try
                        Next
                        oDoc.Update2(True)
                End Select
            Catch ex As Exception
                MessageBox.Show("Lỗi Replace:" & vbCrLf & ex.Message, "Replace Component")
            End Try
        End Sub

        '=================================================
        ' MULTI: chọn SAU khi bật lệnh (Pick từng cái, Esc = xong)
        '=================================================
        Private Function CollectMultiOccurrences(invApp As Inventor.Application, oDoc As AssemblyDocument) As List(Of ComponentOccurrence)
            Dim list As New List(Of ComponentOccurrence)

            If oDoc.SelectSet.Count > 0 Then
                AddFromSelectSet(oDoc, list)
            End If

            BringInventorToFront(invApp)

            Do
                Dim picked As Object = invApp.CommandManager.Pick(
                    SelectionFilterEnum.kAssemblyOccurrenceFilter,
                    "Chọn component cần thay (Esc = xong) | Đã chọn: " & list.Count)
                If picked Is Nothing Then Exit Do

                Dim occ As ComponentOccurrence = ResolveOccurrence(picked)
                If occ Is Nothing Then Continue Do
                If occ.Suppressed Then Continue Do

                Dim already As Boolean = False
                For Each x As ComponentOccurrence In list
                    If x Is occ Then
                        already = True
                        Exit For
                    End If
                Next
                If Not already Then list.Add(occ)
            Loop

            If list.Count = 0 Then
                MessageBox.Show("Bạn chưa chọn component nào.", "Replace Component")
            End If
            Return list
        End Function

        '=================================================
        ' SINGLE (mode 1 / 2)
        '=================================================
        Private Function CollectSingleOccurrence(invApp As Inventor.Application, oDoc As AssemblyDocument, mode As Integer) As List(Of ComponentOccurrence)
            Dim list As New List(Of ComponentOccurrence)
            Dim occOld As ComponentOccurrence = Nothing

            If oDoc.SelectSet.Count >= 1 Then
                occOld = ResolveOccurrence(oDoc.SelectSet.Item(1))
            End If

            If occOld Is Nothing Then
                BringInventorToFront(invApp)

                Dim prompt1 As String = If(mode = 1,
                    "REPLACE ALL - Chọn 1 ASS/Part (lấy file cũ)...",
                    "REPLACE 1 - Chọn đúng 1 ASS/Part cần thay...")
                Dim picked As Object = invApp.CommandManager.Pick(
                    SelectionFilterEnum.kAssemblyOccurrenceFilter, prompt1)
                occOld = ResolveOccurrence(picked)
            End If

            If occOld Is Nothing Then Return list
            list.Add(occOld)
            Return list
        End Function

        '=================================================
        ' Lấy occurrence từ SelectSet (kể cả pattern)
        '=================================================
        Private Sub AddFromSelectSet(oDoc As AssemblyDocument, ByRef list As List(Of ComponentOccurrence))
            For Each obj As Object In oDoc.SelectSet
                Dim occ As ComponentOccurrence = ResolveOccurrence(obj)
                If occ IsNot Nothing AndAlso Not occ.Suppressed Then
                    Dim already As Boolean = False
                    For Each x As ComponentOccurrence In list
                        If x Is occ Then
                            already = True
                            Exit For
                        End If
                    Next
                    If Not already Then list.Add(occ)
                End If
            Next
        End Sub

        '=================================================
        ' Xử lý khi click vào Pattern / proxy
        '=================================================
        Private Function ResolveOccurrence(obj As Object) As ComponentOccurrence
            If obj Is Nothing Then Return Nothing

            Dim occ As ComponentOccurrence = TryCast(obj, ComponentOccurrence)
            If occ IsNot Nothing Then Return occ

            Try
                Dim prop = obj.GetType().GetProperty("ParentOccurrence")
                If prop IsNot Nothing Then
                    Dim parentOcc = TryCast(prop.GetValue(obj, Nothing), ComponentOccurrence)
                    If parentOcc IsNot Nothing Then Return parentOcc
                End If
            Catch
            End Try

            Try
                Dim prop2 = obj.GetType().GetProperty("ContainingOccurrence")
                If prop2 IsNot Nothing Then
                    Dim contOcc = TryCast(prop2.GetValue(obj, Nothing), ComponentOccurrence)
                    If contOcc IsNot Nothing Then Return contOcc
                End If
            Catch
            End Try

            Return Nothing
        End Function

        Private Function SafeGetFile(occ As ComponentOccurrence) As String
            Try
                Return occ.Definition.Document.FullFileName
            Catch
                Return ""
            End Try
        End Function
        '=================================================
        ' FORM CHỌN CHẾ ĐỘ — 4 nút cứng có hover highlight
        '=================================================
        Private Function ShowModeForm() As Integer
            Dim clrBanner As System.Drawing.Color = System.Drawing.Color.FromArgb(0, 122, 204)
            Dim clrBg As System.Drawing.Color = System.Drawing.Color.FromArgb(245, 248, 252)
            Dim clrBtn As System.Drawing.Color = System.Drawing.Color.White
            Dim clrBtnHover As System.Drawing.Color = System.Drawing.Color.FromArgb(217, 236, 255)
            Dim clrBtnDown As System.Drawing.Color = System.Drawing.Color.FromArgb(180, 215, 250)
            Dim clrBorder As System.Drawing.Color = System.Drawing.Color.FromArgb(200, 212, 224)
            Dim clrText As System.Drawing.Color = System.Drawing.Color.FromArgb(30, 30, 30)
            Dim clrTextDim As System.Drawing.Color = System.Drawing.Color.FromArgb(110, 118, 128)

            Dim titles() As String = {
        "Replace All Select",
        "Replace Component Select",
        "Replace Multi Component Select",
        "Replace All Multi Component Select"
    }
            Dim descs() As String = {
        "Thay tất cả instance đang dùng cùng 1 file gốc",
        "Chỉ thay đúng 1 component được chọn",
        "Chọn nhiều component — mỗi cái chỉ thay chính nó",
        "Chọn nhiều loại — mỗi loại thay tất cả instance của nó"
    }

            Dim selectedIndex As Integer = -1
            Dim hoverIndex As Integer = -1

            '--- Pre-create font + pen + brush (1 lần, tránh GDI+ leak) ---
            Using fTitle As New Font("Segoe UI", 10.5F, FontStyle.Bold)
                Using fDesc As New Font("Segoe UI", 8.8F, FontStyle.Regular)
                    Using penNormal As New Pen(clrBorder, 1.0F)
                        Using penHover As New Pen(clrBanner, 2.0F)
                            Using brBgNormal As New SolidBrush(clrBtn)
                                Using brBgHover As New SolidBrush(clrBtnHover)
                                    Using brAccent As New SolidBrush(clrBanner)

                                        Using frm As New Form()
                                            frm.Text = "Replace Component"
                                            frm.Size = New Size(560, 540)
                                            frm.StartPosition = FormStartPosition.CenterScreen
                                            frm.FormBorderStyle = FormBorderStyle.FixedDialog
                                            frm.MaximizeBox = False
                                            frm.MinimizeBox = False
                                            frm.Font = New Font("Segoe UI", 9.5F)
                                            frm.BackColor = clrBg

                                            '--- Banner ---
                                            Dim banner As New Label() With {
                                        .Text = "🔁   REPLACE COMPONENT",
                                        .Dock = DockStyle.Top,
                                        .Height = 52,
                                        .TextAlign = ContentAlignment.MiddleCenter,
                                        .Font = New Font("Segoe UI", 12.5F, FontStyle.Bold),
                                        .BackColor = clrBanner,
                                        .ForeColor = System.Drawing.Color.White
                                    }
                                            frm.Controls.Add(banner)

                                            '--- Label hướng dẫn ---
                                            Dim lblGuide As New Label() With {
                                        .Text = "Click chọn chế độ thay thế:",
                                        .Left = 20, .Top = 64, .AutoSize = True,
                                        .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold),
                                        .ForeColor = clrText
                                    }
                                            frm.Controls.Add(lblGuide)

                                            '--- 4 nút cứng ---
                                            Dim btnW As Integer = 500
                                            Dim btnH As Integer = 68
                                            Dim btnGap As Integer = 8
                                            Dim top0 As Integer = 92

                                            For i As Integer = 0 To 3
                                                Dim idx As Integer = i

                                                Dim b As New Button() With {
                                            .Left = 20,
                                            .Top = top0 + i * (btnH + btnGap),
                                            .Width = btnW,
                                            .Height = btnH,
                                            .BackColor = clrBtn,
                                            .FlatStyle = FlatStyle.Flat,
                                            .Cursor = Cursors.Hand,
                                            .Text = ""
                                        }
                                                b.FlatAppearance.BorderSize = 0     ' viền tự vẽ trong Paint

                                                ' --- Hover tracking ---
                                                AddHandler b.MouseEnter,
                                            Sub()
                                                hoverIndex = idx
                                                b.Invalidate()
                                            End Sub
                                                AddHandler b.MouseLeave,
                                            Sub()
                                                If hoverIndex = idx Then hoverIndex = -1
                                                b.Invalidate()
                                            End Sub

                                                ' --- Paint: vẽ viền + accent bar + text ---
                                                AddHandler b.Paint,
                                            Sub(s, e)
                                                Dim g = e.Graphics
                                                Dim isHover As Boolean = (hoverIndex = idx)

                                                ' 1) Nền
                                                g.FillRectangle(If(isHover, brBgHover, brBgNormal), b.ClientRectangle)

                                                ' 2) Viền
                                                If isHover Then
                                                    g.DrawRectangle(penHover, 0, 0, b.Width - 1, b.Height - 1)
                                                Else
                                                    g.DrawRectangle(penNormal, 0, 0, b.Width - 1, b.Height - 1)
                                                End If

                                                ' 3) Thanh accent trái khi hover
                                                If isHover Then
                                                    g.FillRectangle(brAccent, 0, 0, 4, b.Height)
                                                End If

                                                ' 4) Tiêu đề (đổi màu xanh khi hover)
                                                Dim titleColor As System.Drawing.Color = If(isHover, clrBanner, clrText)
                                                Dim rTitle As New Rectangle(22, 10, b.Width - 40, 24)
                                                TextRenderer.DrawText(g,
                                                    (idx + 1).ToString() & ".  " & titles(idx),
                                                    fTitle, rTitle, titleColor,
                                                    TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)

                                                ' 5) Mô tả
                                                Dim rDesc As New Rectangle(22, 36, b.Width - 40, 22)
                                                TextRenderer.DrawText(g, descs(idx),
                                                    fDesc, rDesc, clrTextDim,
                                                    TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)

                                                ' 6) Mũi tên › bên phải khi hover
                                                If isHover Then
                                                    Using fArrow As New Font("Segoe UI", 16.0F, FontStyle.Bold)
                                                        Dim rArrow As New Rectangle(b.Width - 36, 0, 26, b.Height)
                                                        TextRenderer.DrawText(g, "›", fArrow, rArrow, clrBanner,
                                                            TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
                                                    End Using
                                                End If
                                            End Sub

                                                ' --- Click: chọn & đóng ---
                                                AddHandler b.Click,
                                            Sub()
                                                selectedIndex = idx
                                                frm.DialogResult = DialogResult.OK
                                                frm.Close()
                                            End Sub

                                                frm.Controls.Add(b)
                                            Next

                                            '--- Nút Hủy ---
                                            Dim btnCancel As New Button() With {
                                        .Text = "Hủy",
                                        .Left = 420, .Top = 440, .Width = 100, .Height = 36,
                                        .BackColor = System.Drawing.Color.FromArgb(230, 70, 70),
                                        .ForeColor = System.Drawing.Color.White,
                                        .FlatStyle = FlatStyle.Flat,
                                        .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold),
                                        .Cursor = Cursors.Hand,
                                        .DialogResult = DialogResult.Cancel
                                    }
                                            btnCancel.FlatAppearance.BorderSize = 0
                                            AddHandler btnCancel.MouseEnter, Sub() btnCancel.BackColor = System.Drawing.Color.FromArgb(210, 55, 55)
                                            AddHandler btnCancel.MouseLeave, Sub() btnCancel.BackColor = System.Drawing.Color.FromArgb(230, 70, 70)
                                            frm.Controls.Add(btnCancel)

                                            frm.CancelButton = btnCancel

                                            ' Reset hover khi form không còn focus (tránh treo highlight)
                                            AddHandler frm.Deactivate,
                                        Sub()
                                            hoverIndex = -1
                                            For Each c As Control In frm.Controls
                                                If TypeOf c Is Button Then c.Invalidate()
                                            Next
                                        End Sub

                                            Dim result As DialogResult = frm.ShowDialog()
                                            If result <> DialogResult.OK Then Return 0
                                            If selectedIndex < 0 Then Return 0
                                            Return selectedIndex + 1

                                        End Using
                                    End Using
                                End Using
                            End Using
                        End Using
                    End Using
                End Using
            End Using
        End Function
    End Module
End Namespace
