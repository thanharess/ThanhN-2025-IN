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
        ' FORM CHỌN CHẾ ĐỘ
        '=================================================
        Private Function ShowModeForm() As Integer
            Dim frm As New Form()
            frm.Text = "Replace Component"
            frm.Size = New Size(460, 310)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.Font = New Font("Segoe UI", 10)

            Dim lbl As New Label()
            lbl.Text = "Chọn chế độ thay thế:"
            lbl.Location = New System.Drawing.Point(20, 15)
            lbl.AutoSize = True
            frm.Controls.Add(lbl)

            Dim lst As New ListBox()
            lst.Location = New System.Drawing.Point(20, 45)
            lst.Size = New Size(400, 150)
            lst.Font = New Font("Segoe UI", 10.5F)
            lst.Items.Add("1. Replace All        - Thay tất cả instance cùng 1 file")
            lst.Items.Add("2. Replace 1          - Chỉ thay đúng 1 component")
            lst.Items.Add("3. Replace Nhiều      - Chọn nhiều, mỗi cái chỉ thay chính nó")
            lst.Items.Add("4. Replace All Multi  - Chọn nhiều loại → mỗi loại Replace All")
            lst.SelectedIndex = 0
            frm.Controls.Add(lst)

            Dim btnOK As New Button()
            btnOK.Text = "OK"
            btnOK.Size = New Size(100, 32)
            btnOK.Location = New System.Drawing.Point(210, 215)
            btnOK.DialogResult = DialogResult.OK
            frm.Controls.Add(btnOK)

            Dim btnCancel As New Button()
            btnCancel.Text = "Hủy"
            btnCancel.Size = New Size(100, 32)
            btnCancel.Location = New System.Drawing.Point(320, 215)
            btnCancel.DialogResult = DialogResult.Cancel
            frm.Controls.Add(btnCancel)

            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            Dim result As DialogResult = frm.ShowDialog()
            If result <> DialogResult.OK Then Return 0
            Return lst.SelectedIndex + 1
        End Function

    End Module
End Namespace