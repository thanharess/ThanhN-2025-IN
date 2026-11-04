Option Explicit On
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Frame

    '============================================================
    ' ⭐ Marshal2 – thay thế Marshal.GetActiveObject cho .NET 8+
    '    (Marshal.GetActiveObject đã bị remove từ .NET Core 3.0)
    '    Dùng P/Invoke oleaut32!GetActiveObject + ole32!CLSIDFromProgID
    '============================================================
    Friend NotInheritable Class Marshal2
        Private Sub New()
        End Sub

        <DllImport("ole32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
        Private Shared Sub CLSIDFromProgIDEx(
            <MarshalAs(UnmanagedType.LPWStr)> progId As String,
            <Out> ByRef clsid As Guid)
        End Sub                          ' ⭐ THÊM DÒNG NÀY

        <DllImport("ole32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
        Private Shared Sub CLSIDFromProgID(
            <MarshalAs(UnmanagedType.LPWStr)> progId As String,
            <Out> ByRef clsid As Guid)
        End Sub                          ' ⭐ THÊM DÒNG NÀY

        <DllImport("oleaut32.dll", PreserveSig:=False)>
        Private Shared Sub GetActiveObject(
            ByRef rclsid As Guid,
            reserved As IntPtr,
            <MarshalAs(UnmanagedType.Interface)> <Out> ByRef ppunk As Object)
        End Sub                          ' ⭐ THÊM DÒNG NÀY

        Public Shared Function GetActiveObject(progId As String) As Object
            Dim clsid As Guid
            Try
                CLSIDFromProgIDEx(progId, clsid)
            Catch
                CLSIDFromProgID(progId, clsid)
            End Try

            Dim obj As Object
            GetActiveObject(clsid, IntPtr.Zero, obj)
            Return obj
        End Function
    End Class


    Public Module Ass_Frame_1B

        Private _invApp As Inventor.Application = Nothing

        Private Const SICK_DISPLAY_STATE As Integer = 46852

        Private Enum FrameTreatmentType
            Unknown = 0
            TrimExtend = 1
            Miter = 2
            Notch = 3
            LengthenShorten = 4
            Corners = 5
        End Enum

        Private Class SickFrameInfoa
            Public Property Node As Inventor.BrowserNode
            Public Property Label As String
            Public Property FullPath As String
            Public Property ToolTip As String
            Public Property NativeObject As Object
            Public Property TreatmentType As FrameTreatmentType
        End Class

        Private _sickTreatments As New List(Of SickFrameInfoa)

        '============================================================
        ' GET INVENTOR APPLICATION
        ' ⭐ ĐÃ SỬA: Marshal → Marshal2 (do .NET 8 remove API cũ)
        '============================================================
        Private Function GetInventorApplication(ByVal Context As NameValueMap) As Inventor.Application

            Dim app As Inventor.Application = Nothing

            Try
                If Context IsNot Nothing Then
                    Try
                        app = DirectCast(Context.Item("Application"), Inventor.Application)
                    Catch
                    End Try
                End If
            Catch
            End Try

            If app IsNot Nothing Then Return app

            Try
                ' ⭐ Marshal2 thay cho Marshal.GetActiveObject
                app = DirectCast(Marshal2.GetActiveObject("Inventor.Application"),
                                 Inventor.Application)
            Catch ex As Exception
                MessageBox.Show("Không lấy được Inventor.Application." & vbCrLf & vbCrLf & ex.Message,
                                "Frame Generator", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End Try

            Return app

        End Function

        '============================================================
        ' GỌI LỆNH CÓ SẴN CỦA INVENTOR
        '============================================================
        Private Function TryExecuteCommand(ParamArray names() As String) As Boolean

            If _invApp Is Nothing Then Return False

            For Each nm As String In names
                Try
                    Dim cmd As ControlDefinition =
                        _invApp.CommandManager.ControlDefinitions.Item(nm)
                    If cmd IsNot Nothing AndAlso cmd.Enabled Then
                        cmd.Execute()
                        Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
                        Try : System.Threading.Thread.Sleep(250) : Catch : End Try
                        Return True
                    End If
                Catch
                End Try
            Next

            Return False

        End Function

        '============================================================
        ' ⭐ BẢNG TICK CHỌN LOẠI XỬ LÝ
        '============================================================
        Private Function ShowTreatmentSelector(
            ByVal trimCount As Integer,
            ByVal miterCount As Integer,
            ByVal notchCount As Integer,
            ByVal lengthenCount As Integer,
            ByVal cornersCount As Integer) As HashSet(Of FrameTreatmentType)

            Dim result As HashSet(Of FrameTreatmentType) = Nothing

            Dim f As New System.Windows.Forms.Form
            f.Text = "FRAME GENERATOR - Chọn loại xử lý"
            f.Width = 440
            f.Height = 380
            f.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            f.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            f.MaximizeBox = False
            f.MinimizeBox = False
            f.ShowInTaskbar = False

            Dim lblTitle As New System.Windows.Forms.Label
            lblTitle.Text = "Chọn loại Frame Treatment muốn xóa:"
            lblTitle.Location = New System.Drawing.Point(15, 15)
            lblTitle.AutoSize = True
            f.Controls.Add(lblTitle)

            Dim cbTrim As New System.Windows.Forms.CheckBox
            cbTrim.Text = "Trim / Extend                  (" & trimCount.ToString() & ")"
            cbTrim.Location = New System.Drawing.Point(20, 45)
            cbTrim.Width = 380
            cbTrim.Checked = True
            cbTrim.Enabled = (trimCount > 0)
            f.Controls.Add(cbTrim)

            Dim cbMiter As New System.Windows.Forms.CheckBox
            cbMiter.Text = "Miter / Mitre                  (" & miterCount.ToString() & ")"
            cbMiter.Location = New System.Drawing.Point(20, 70)
            cbMiter.Width = 380
            cbMiter.Checked = True
            cbMiter.Enabled = (miterCount > 0)
            f.Controls.Add(cbMiter)

            Dim cbNotch As New System.Windows.Forms.CheckBox
            cbNotch.Text = "Notch                          (" & notchCount.ToString() & ")"
            cbNotch.Location = New System.Drawing.Point(20, 95)
            cbNotch.Width = 380
            cbNotch.Checked = True
            cbNotch.Enabled = (notchCount > 0)
            f.Controls.Add(cbNotch)

            Dim cbLengthen As New System.Windows.Forms.CheckBox
            cbLengthen.Text = "Lengthen / Shorten             (" & lengthenCount.ToString() & ")"
            cbLengthen.Location = New System.Drawing.Point(20, 120)
            cbLengthen.Width = 380
            cbLengthen.Checked = True
            cbLengthen.Enabled = (lengthenCount > 0)
            f.Controls.Add(cbLengthen)

            Dim sep As New System.Windows.Forms.Label
            sep.Text = ""
            sep.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            sep.Location = New System.Drawing.Point(20, 155)
            sep.Width = 380
            sep.Height = 2
            f.Controls.Add(sep)

            Dim cbCorners As New System.Windows.Forms.CheckBox
            cbCorners.Text = "Insert End Cap / Corners  (" & cornersCount.ToString() & ")"
            cbCorners.Location = New System.Drawing.Point(20, 170)
            cbCorners.Width = 380
            cbCorners.Checked = False
            cbCorners.Enabled = (cornersCount > 0)
            cbCorners.Font = New System.Drawing.Font(cbCorners.Font, System.Drawing.FontStyle.Bold)
            f.Controls.Add(cbCorners)

            Dim lblNote As New System.Windows.Forms.Label
            lblNote.Text = "(Mặc định KHÔNG xóa — tick vào nếu muốn xóa)"
            lblNote.Location = New System.Drawing.Point(38, 195)
            lblNote.AutoSize = True
            lblNote.ForeColor = System.Drawing.Color.Gray
            f.Controls.Add(lblNote)

            Dim btnAll As New System.Windows.Forms.Button
            btnAll.Text = "Chọn tất cả"
            btnAll.Location = New System.Drawing.Point(20, 230)
            btnAll.Width = 100
            AddHandler btnAll.Click, Sub()
                                         If cbTrim.Enabled Then cbTrim.Checked = True
                                         If cbMiter.Enabled Then cbMiter.Checked = True
                                         If cbNotch.Enabled Then cbNotch.Checked = True
                                         If cbLengthen.Enabled Then cbLengthen.Checked = True
                                         If cbCorners.Enabled Then cbCorners.Checked = True
                                     End Sub
            f.Controls.Add(btnAll)

            Dim btnNone As New System.Windows.Forms.Button
            btnNone.Text = "Bỏ chọn"
            btnNone.Location = New System.Drawing.Point(130, 230)
            btnNone.Width = 100
            AddHandler btnNone.Click, Sub()
                                          cbTrim.Checked = False
                                          cbMiter.Checked = False
                                          cbNotch.Checked = False
                                          cbLengthen.Checked = False
                                          cbCorners.Checked = False
                                      End Sub
            f.Controls.Add(btnNone)

            Dim btnOK As New System.Windows.Forms.Button
            btnOK.Text = "OK"
            btnOK.Location = New System.Drawing.Point(200, 285)
            btnOK.Width = 90
            btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
            f.Controls.Add(btnOK)

            Dim btnCancel As New System.Windows.Forms.Button
            btnCancel.Text = "Cancel"
            btnCancel.Location = New System.Drawing.Point(300, 285)
            btnCancel.Width = 90
            btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            f.Controls.Add(btnCancel)

            f.AcceptButton = btnOK
            f.CancelButton = btnCancel

            If f.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                result = New HashSet(Of FrameTreatmentType)

                If cbTrim.Checked Then result.Add(FrameTreatmentType.TrimExtend)
                If cbMiter.Checked Then result.Add(FrameTreatmentType.Miter)
                If cbNotch.Checked Then result.Add(FrameTreatmentType.Notch)
                If cbLengthen.Checked Then result.Add(FrameTreatmentType.LengthenShorten)
                If cbCorners.Checked Then result.Add(FrameTreatmentType.Corners)

            End If

            Return result

        End Function

        '============================================================
        ' MAIN BUTTON
        '============================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try

                _invApp = GetInventorApplication(Context)
                If _invApp Is Nothing Then Return

                Dim doc As Inventor.Document = _invApp.ActiveDocument

                If doc Is Nothing Then
                    MessageBox.Show("Không có document đang mở.", "Frame Generator",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Hãy chạy trong Assembly.", "Frame Generator",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)

                ForceUpdate(asmDoc)

                ForceExpandAllBrowser(doc)

                _sickTreatments.Clear()
                ScanFrameBrowser(doc)

                If _sickTreatments.Count = 0 Then
                    MessageBox.Show("Không tìm thấy Frame Treatment bị lỗi.",
                                    "Frame Generator", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim trimCount As Integer = 0
                Dim miterCount As Integer = 0
                Dim notchCount As Integer = 0
                Dim lengthenCount As Integer = 0
                Dim cornersCount As Integer = 0

                For Each item As SickFrameInfoa In _sickTreatments
                    If item Is Nothing Then Continue For
                    Select Case item.TreatmentType
                        Case FrameTreatmentType.TrimExtend : trimCount += 1
                        Case FrameTreatmentType.Miter : miterCount += 1
                        Case FrameTreatmentType.Notch : notchCount += 1
                        Case FrameTreatmentType.LengthenShorten : lengthenCount += 1
                        Case FrameTreatmentType.Corners : cornersCount += 1
                    End Select
                Next

                Dim allowed As HashSet(Of FrameTreatmentType) =
                    ShowTreatmentSelector(trimCount, miterCount, notchCount, lengthenCount, cornersCount)

                If allowed Is Nothing Then Return
                If allowed.Count = 0 Then Return

                ProcessSickTreatments(asmDoc, allowed)

            Catch ex As Exception
                MessageBox.Show("LỖI:" & vbCrLf & vbCrLf & ex.Message, "Frame Generator",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)

            Finally
            End Try

        End Sub

        '============================================================
        ' FORCE UPDATE
        '============================================================
        Private Sub ForceUpdate(ByVal asmDoc As AssemblyDocument)

            If asmDoc Is Nothing Then Return

            Try : asmDoc.Update2(True) : Catch : End Try
            Try : asmDoc.Rebuild2() : Catch : End Try
            Try
                If _invApp IsNot Nothing Then _invApp.ActiveView.Update()
            Catch
            End Try

        End Sub

        '============================================================
        ' SCAN FRAME BROWSER
        '============================================================
        Private Sub ScanFrameBrowser(ByVal doc As Inventor.Document)

            If doc Is Nothing Then Return

            Try

                Dim pane As Inventor.BrowserPane = Nothing

                Try
                    pane = doc.BrowserPanes.Item("Model")
                Catch
                    Try
                        pane = doc.BrowserPanes.ActivePane
                    Catch
                        pane = Nothing
                    End Try
                End Try

                If pane Is Nothing Then Return
                If pane.TopNode Is Nothing Then Return

                ScanBrowserNode(pane.TopNode)

            Catch
            End Try

        End Sub

        '============================================================
        ' SCAN NODE
        '============================================================
        Private Sub ScanBrowserNode(ByVal node As Inventor.BrowserNode)

            If node Is Nothing Then Return

            Try

                Dim label As String = ""
                Dim fullPath As String = ""
                Dim tooltip As String = ""
                Dim nativeObj As Object = Nothing
                Dim isSick As Boolean = False

                Try
                    label = node.BrowserNodeDefinition.Label
                Catch
                    Try
                        label = node.FullPath
                    Catch
                        label = ""
                    End Try
                End Try

                Try
                    fullPath = node.FullPath
                Catch
                    fullPath = label
                End Try

                Try
                    tooltip = node.BrowserNodeDefinition.StateIconToolTipText
                Catch
                    tooltip = ""
                End Try

                Try
                    nativeObj = node.NativeObject
                Catch
                    nativeObj = Nothing
                End Try

                Try
                    Dim ds As Object = node.BrowserNodeDefinition.DisplayState
                    Dim stateNumber As Integer = Convert.ToInt32(ds)

                    If stateNumber = SICK_DISPLAY_STATE Then isSick = True

                    If Not isSick Then
                        Select Case stateNumber
                            Case 46852, 46853, 46854, 46855
                                isSick = True
                        End Select
                    End If
                Catch
                    isSick = False
                End Try

                If Not isSick Then
                    Try
                        If Not String.IsNullOrEmpty(tooltip) Then
                            Dim tl = tooltip.ToLowerInvariant()
                            If tl.Contains("error") OrElse tl.Contains("sick") OrElse
                               tl.Contains("fail") OrElse tl.Contains("lỗi") OrElse
                               tl.Contains("không thể") Then
                                isSick = True
                            End If
                        End If
                    Catch
                    End Try

                    If Not isSick AndAlso nativeObj IsNot Nothing Then
                        Try
                            Dim hs As Object = nativeObj.HealthStatus
                            If hs IsNot Nothing Then
                                Dim hstr = hs.ToString()
                                If hstr.Contains("Sick") OrElse hstr.Contains("Error") Then
                                    isSick = True
                                End If
                            End If
                        Catch
                        End Try
                    End If
                End If

                If isSick Then

                    Dim treatmentType As FrameTreatmentType =
                        DetectTreatmentType(label, fullPath, tooltip)

                    If treatmentType <> FrameTreatmentType.Unknown Then

                        If Not IsAlreadyAdded(node) Then

                            Dim info As New SickFrameInfoa
                            info.Node = node
                            info.Label = label
                            info.FullPath = fullPath
                            info.ToolTip = tooltip
                            info.NativeObject = nativeObj
                            info.TreatmentType = treatmentType

                            _sickTreatments.Add(info)

                        End If

                    End If

                End If

                Dim children As Object = Nothing
                Try
                    children = node.BrowserNodes
                Catch
                    children = Nothing
                End Try

                If children IsNot Nothing Then
                    Try
                        For Each child As Object In children
                            Try
                                Dim childNode As Inventor.BrowserNode =
                                    DirectCast(child, Inventor.BrowserNode)
                                If childNode IsNot Nothing Then
                                    ScanBrowserNode(childNode)
                                End If
                            Catch
                            End Try
                        Next
                    Catch
                    End Try
                End If

            Catch
            End Try

        End Sub

        '============================================================
        ' DETECT TREATMENT
        '============================================================
        Private Function DetectTreatmentType(
            ByVal label As String,
            ByVal fullPath As String,
            ByVal tooltip As String) As FrameTreatmentType

            Dim l As String = If(label, "").ToLowerInvariant()
            Dim p As String = If(fullPath, "").ToLowerInvariant()
            Dim t As String = If(tooltip, "").ToLowerInvariant()

            If l.Contains("trim") OrElse p.Contains("trim") OrElse t.Contains("trim") Then
                Return FrameTreatmentType.TrimExtend
            End If

            If l.Contains("miter") OrElse l.Contains("mitre") OrElse
               p.Contains("miter") OrElse p.Contains("mitre") OrElse
               t.Contains("miter") OrElse t.Contains("mitre") Then
                Return FrameTreatmentType.Miter
            End If

            If l.Contains("notch") OrElse p.Contains("notch") OrElse t.Contains("notch") Then
                Return FrameTreatmentType.Notch
            End If

            If l.Contains("lengthen") OrElse l.Contains("shorten") OrElse
               p.Contains("lengthen") OrElse p.Contains("shorten") OrElse
               t.Contains("lengthen") OrElse t.Contains("shorten") Then
                Return FrameTreatmentType.LengthenShorten
            End If

            If l.Contains("shop corner") OrElse p.Contains("shop corner") OrElse t.Contains("shop corner") OrElse
               l.Contains("insert end cap") OrElse p.Contains("insert end cap") OrElse t.Contains("insert end cap") OrElse
               l.Contains("sharp corners") OrElse p.Contains("sharp corners") OrElse t.Contains("sharp corners") Then
                Return FrameTreatmentType.Corners
            End If

            Return FrameTreatmentType.Unknown

        End Function

        '============================================================
        ' DUPLICATE
        '============================================================
        Private Function IsAlreadyAdded(ByVal node As Inventor.BrowserNode) As Boolean

            If node Is Nothing Then Return False

            For Each item As SickFrameInfoa In _sickTreatments
                If item Is Nothing Then Continue For
                Try
                    If Object.ReferenceEquals(item.Node, node) Then Return True
                Catch
                End Try
            Next

            Return False

        End Function

        '============================================================
        ' PROCESS
        '============================================================
        Private Sub ProcessSickTreatments(
            ByVal asmDoc As AssemblyDocument,
            ByVal allowed As HashSet(Of FrameTreatmentType))

            Dim deleteSuccess As Integer = 0
            Dim deleteFailed As Integer = 0
            Dim skipped As Integer = 0

            Dim workList As New List(Of SickFrameInfoa)

            For Each item As SickFrameInfoa In _sickTreatments
                If item IsNot Nothing Then workList.Add(item)
            Next

            Dim deleteCmd As ControlDefinition = Nothing
            Try
                deleteCmd = _invApp.CommandManager.ControlDefinitions.Item("Delete")
            Catch
                deleteCmd = Nothing
            End Try

            For i As Integer = workList.Count - 1 To 0 Step -1

                Dim item As SickFrameInfoa = workList(i)
                If item Is Nothing Then Continue For

                If Not allowed.Contains(item.TreatmentType) Then
                    skipped += 1
                    Continue For
                End If

                If item.TreatmentType = FrameTreatmentType.TrimExtend OrElse
                   item.TreatmentType = FrameTreatmentType.Miter OrElse
                   item.TreatmentType = FrameTreatmentType.Notch OrElse
                   item.TreatmentType = FrameTreatmentType.LengthenShorten OrElse
                   item.TreatmentType = FrameTreatmentType.Corners Then

                    If deleteCmd Is Nothing Then
                        deleteFailed += 1
                        Continue For
                    End If

                    Try : asmDoc.SelectSet.Clear() : Catch : End Try

                    Try
                        Dim pane As BrowserPane = asmDoc.BrowserPanes.Item("Model")
                        Try : pane.ClearSelection() : Catch : End Try
                    Catch
                    End Try

                    Dim selected As Boolean = False

                    Try
                        item.Node.Select()
                        selected = True
                    Catch
                    End Try

                    If Not selected Then
                        Try
                            item.Node.DoSelect()
                            selected = True
                        Catch
                        End Try
                    End If

                    If Not selected Then
                        deleteFailed += 1
                        Continue For
                    End If

                    Try
                        If Not deleteCmd.Enabled Then
                            deleteFailed += 1
                            Continue For
                        End If
                    Catch
                    End Try

                    Try
                        deleteCmd.Execute()
                    Catch
                        deleteFailed += 1
                        Continue For
                    End Try

                    Try : asmDoc.Update2(True) : Catch : End Try
                    Try : _invApp.ActiveView.Update() : Catch : End Try

                    Dim stillExists As Boolean = BrowserNodeStillExists(asmDoc, item.FullPath)
                    If stillExists Then
                        deleteFailed += 1
                    Else
                        deleteSuccess += 1
                    End If

                End If

            Next

            Try : asmDoc.Update2(True) : Catch : End Try
            Try : asmDoc.Rebuild2() : Catch : End Try
            Try : _invApp.ActiveView.Update() : Catch : End Try

            _sickTreatments.Clear()
            ScanFrameBrowser(asmDoc)

            Dim remainTrim As Integer = 0
            Dim remainMiter As Integer = 0
            Dim remainNotch As Integer = 0
            Dim remainLengthen As Integer = 0
            Dim remainCorners As Integer = 0

            For Each item As SickFrameInfoa In _sickTreatments
                If item Is Nothing Then Continue For
                Select Case item.TreatmentType
                    Case FrameTreatmentType.TrimExtend : remainTrim += 1
                    Case FrameTreatmentType.Miter : remainMiter += 1
                    Case FrameTreatmentType.Notch : remainNotch += 1
                    Case FrameTreatmentType.LengthenShorten : remainLengthen += 1
                    Case FrameTreatmentType.Corners : remainCorners += 1
                End Select
            Next

            Dim remainTotal As Integer = _sickTreatments.Count

            Dim result As New StringBuilder
            result.AppendLine("========== FRAME GENERATOR ==========")
            result.AppendLine()
            result.AppendLine("ĐÃ SỬA:")
            result.AppendLine("Delete thành công : " & deleteSuccess.ToString())
            result.AppendLine("Delete thất bại   : " & deleteFailed.ToString())
            If skipped > 0 Then
                result.AppendLine("Bỏ qua (không tick): " & skipped.ToString())
            End If
            result.AppendLine()
            result.AppendLine("CÒN SICK:")
            result.AppendLine("Trim / Extend     : " & remainTrim.ToString())
            result.AppendLine("Miter             : " & remainMiter.ToString())
            result.AppendLine("Notch             : " & remainNotch.ToString())
            result.AppendLine("Lengthen / Shorten: " & remainLengthen.ToString())
            result.AppendLine("Corners           : " & remainCorners.ToString())
            result.AppendLine()
            result.AppendLine("Tổng còn Sick     : " & remainTotal.ToString())
            result.AppendLine()

            If remainTotal = 0 Then
                result.AppendLine("✓ HOÀN TẤT - KHÔNG CÒN SICK.")
            Else
                result.AppendLine("⚠ VẪN CÒN " & remainTotal.ToString() & " NODE SICK.")
            End If

            _sickTreatments.Clear()

            MessageBox.Show(result.ToString(), "FRAME GENERATOR",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub

        '============================================================
        ' CHECK NODE
        '============================================================
        Private Function BrowserNodeStillExists(
            ByVal doc As Inventor.Document,
            ByVal targetPath As String) As Boolean

            If doc Is Nothing Then Return False
            If String.IsNullOrEmpty(targetPath) Then Return False

            Try
                Dim pane As BrowserPane = doc.BrowserPanes.Item("Model")
                If pane Is Nothing Then Return False
                Return SearchBrowserPath(pane.TopNode, targetPath)
            Catch
                Return False
            End Try

        End Function

        Private Function SearchBrowserPath(
            ByVal node As Inventor.BrowserNode,
            ByVal targetPath As String) As Boolean

            If node Is Nothing Then Return False

            Try
                Dim currentPath As String = ""
                Try : currentPath = node.FullPath : Catch : End Try

                If String.Equals(currentPath, targetPath, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If

                Dim children As Object = Nothing
                Try : children = node.BrowserNodes : Catch : children = Nothing : End Try

                If children IsNot Nothing Then
                    Try
                        For Each child As Object In children
                            Try
                                Dim childNode As Inventor.BrowserNode =
                                    DirectCast(child, Inventor.BrowserNode)
                                If SearchBrowserPath(childNode, targetPath) Then Return True
                            Catch
                            End Try
                        Next
                    Catch
                    End Try
                End If
            Catch
            End Try

            Return False

        End Function

        '============================================================
        ' FORCE EXPAND
        '============================================================
        Private Sub ForceExpandAllBrowser(ByVal doc As Inventor.Document)

            If doc Is Nothing Then Return

            Try
                Dim pane As Inventor.BrowserPane = Nothing
                Try
                    pane = doc.BrowserPanes.Item("Model")
                Catch
                    Try : pane = doc.BrowserPanes.ActivePane : Catch : End Try
                End Try

                If pane IsNot Nothing Then
                    Try : pane.Activate() : Catch : End Try
                End If
            Catch
            End Try

            Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
            Try : System.Threading.Thread.Sleep(100) : Catch : End Try

            TryExecuteCommand("ExpandAllCmd", "BrowserExpandAllCmd", "ExpandAll")

            Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
            Try : System.Threading.Thread.Sleep(300) : Catch : End Try

        End Sub

        '============================================================
        ' ON DEBUG
        '============================================================
        Public Sub OnDebug(ByVal Context As NameValueMap)

            Try

                _invApp = GetInventorApplication(Context)
                If _invApp Is Nothing Then Return

                Dim doc As Inventor.Document = _invApp.ActiveDocument
                If doc Is Nothing Then Return

                Dim sb As New StringBuilder
                sb.AppendLine("============================================================")
                sb.AppendLine(" FRAME GENERATOR BROWSER DEBUG - INVENTOR 2025 / .NET 8")
                sb.AppendLine("============================================================")
                sb.AppendLine()
                sb.AppendLine("DOCUMENT: " & doc.DisplayName)
                sb.AppendLine()

                DebugBrowser(doc, sb)
                ShowDebugText(sb.ToString())

            Catch ex As Exception
                MessageBox.Show(ex.Message, "Frame Debug",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        Private Sub DebugBrowser(ByVal doc As Inventor.Document, ByVal sb As StringBuilder)

            Try
                Dim pane As Inventor.BrowserPane = Nothing
                Try
                    pane = doc.BrowserPanes.Item("Model")
                Catch
                    Try : pane = doc.BrowserPanes.ActivePane : Catch : End Try
                End Try

                If pane Is Nothing Then
                    sb.AppendLine("Không lấy được Model Browser.")
                    Return
                End If

                DebugBrowserNode(pane.TopNode, sb, 0)

            Catch ex As Exception
                sb.AppendLine("BROWSER ERROR: " & ex.Message)
            End Try

        End Sub

        Private Sub DebugBrowserNode(
            ByVal node As Inventor.BrowserNode,
            ByVal sb As StringBuilder,
            ByVal level As Integer)

            If node Is Nothing Then Return

            Try

                Dim indent As String = New String(" "c, level * 2)

                Dim label As String = ""
                Dim path As String = ""
                Dim tooltip As String = ""
                Dim state As Integer = -1
                Dim nativeType As String = ""

                Try : label = node.BrowserNodeDefinition.Label : Catch : label = "?" : End Try
                Try : path = node.FullPath : Catch : path = "" : End Try
                Try : tooltip = node.BrowserNodeDefinition.StateIconToolTipText : Catch : tooltip = "" : End Try

                Try
                    Dim ds As Object = node.BrowserNodeDefinition.DisplayState
                    state = Convert.ToInt32(ds)
                Catch
                    state = -1
                End Try

                Try
                    Dim obj As Object = node.NativeObject
                    If obj IsNot Nothing Then nativeType = obj.GetType().FullName
                Catch
                    nativeType = ""
                End Try

                Dim treatmentType As FrameTreatmentType =
                    DetectTreatmentType(label, path, tooltip)

                Dim important As Boolean =
                    treatmentType <> FrameTreatmentType.Unknown OrElse
                    state = SICK_DISPLAY_STATE

                If important Then
                    sb.AppendLine()
                    sb.AppendLine(indent & "----------------------------------------")
                    sb.AppendLine(indent & "NAME: " & label)
                    sb.AppendLine(indent & "TYPE: " & GetTreatmentTypeName(treatmentType))
                    sb.AppendLine(indent & "STATE: " & state.ToString())

                    If state = SICK_DISPLAY_STATE Then
                        sb.AppendLine(indent & ">>> SICK NODE <<<")
                    End If

                    sb.AppendLine(indent & "TOOLTIP: " & tooltip)
                    sb.AppendLine(indent & "NATIVE: " & nativeType)
                    sb.AppendLine(indent & "PATH: " & path)
                End If

                Dim children As Object = Nothing
                Try : children = node.BrowserNodes : Catch : children = Nothing : End Try

                If children IsNot Nothing Then
                    Try
                        For Each child As Object In children
                            Try
                                Dim childNode As Inventor.BrowserNode =
                                    DirectCast(child, Inventor.BrowserNode)
                                If childNode IsNot Nothing Then
                                    DebugBrowserNode(childNode, sb, level + 1)
                                End If
                            Catch
                            End Try
                        Next
                    Catch
                    End Try
                End If

            Catch
            End Try

        End Sub

        Private Function GetTreatmentTypeName(ByVal treatmentType As FrameTreatmentType) As String

            Select Case treatmentType
                Case FrameTreatmentType.TrimExtend : Return "TRIM / EXTEND"
                Case FrameTreatmentType.Miter : Return "MITER / MITRE"
                Case FrameTreatmentType.Notch : Return "NOTCH"
                Case FrameTreatmentType.LengthenShorten : Return "LENGTHEN / SHORTEN"
                Case FrameTreatmentType.Corners : Return "CORNERS"
                Case Else : Return "UNKNOWN"
            End Select

        End Function

        Private Sub ShowDebugText(ByVal text As String)

            Dim f As New System.Windows.Forms.Form
            f.Text = "FRAME GENERATOR DEBUG"
            f.Width = 1200
            f.Height = 800

            Dim tb As New System.Windows.Forms.TextBox
            tb.Multiline = True
            tb.ReadOnly = True
            tb.ScrollBars = System.Windows.Forms.ScrollBars.Both
            tb.WordWrap = False
            tb.Dock = System.Windows.Forms.DockStyle.Fill
            tb.Font = New System.Drawing.Font("Consolas", 10.0F)
            tb.Text = text

            f.Controls.Add(tb)
            f.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            f.ShowDialog()

        End Sub

    End Module

End Namespace