Option Explicit On
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Frame

    Public Module Ass_Frame_1

        Private _invApp As Inventor.Application = Nothing

        Private Const SICK_DISPLAY_STATE As Integer = 46852

        '============================================================
        ' ⭐ DEBUG: bật/tắt collapse để xem kết quả
        '   True  = KHÔNG collapse (giữ browser mở để xem)
        '   False = collapse như bình thường
        '
        '   ★ Khi cần dùng collapse → đổi thành False
        '============================================================
        Private Const DISABLE_COLLAPSE As Boolean = False

        Private Enum FrameTreatmentType
            Unknown = 0
            TrimExtend = 1
            Miter = 2
            Notch = 3
            LengthenShorten = 4
            Corners = 5
        End Enum

        Private Class SickFrameInfo
            Public Property Node As Inventor.BrowserNode
            Public Property Label As String
            Public Property FullPath As String
            Public Property ToolTip As String
            Public Property NativeObject As Object
            Public Property TreatmentType As FrameTreatmentType
        End Class

        Private _sickTreatments As New List(Of SickFrameInfo)

        '============================================================
        ' BỎ QUA CÁC NHÁNH KHÔNG CHỨA FRAME TREATMENT
        '
        '   Assembly / Part:
        '     • Origin
        '     • Relationships
        '     • Representations
        '     • End of Features
        '     • Welds
        '     • Machining
        '     • Frame Reference Model
        '     • Substitutes
        '     • Reference Skeleton
        '
        '   Khớp mềm:
        '     • Model States: [Primary]  (mọi biến thể)
        '     • Sketch / Sketch Rectangular Pattern...
        '============================================================
        Private Function ShouldSkipBranch(ByVal label As String) As Boolean

            If String.IsNullOrEmpty(label) Then Return False

            Dim l As String = label.Trim().ToLowerInvariant()

            '--- Bỏ hậu tố ":N" nếu có ---
            Dim idx As Integer = l.LastIndexOf(":"c)
            If idx > 0 Then
                Dim tail As String = l.Substring(idx + 1)
                Dim n As Integer
                If Integer.TryParse(tail, n) Then l = l.Substring(0, idx).Trim()
            End If

            '--- Khớp chính xác ---
            Select Case l
                Case "origin",
                     "relationships",
                     "representations",
                     "end of features",
                     "welds",
                     "machining",
                     "frame reference model",
                     "substitutes",
                     "reference skeleton",
                     "folded model"
                    Return True
            End Select

            '--- Khớp mềm ---
            If l.StartsWith("model states") Then Return True
            If l.StartsWith("sketch") Then Return True
            If l.StartsWith("reference skeleton") Then Return True

            Return False

        End Function

        '============================================================
        ' GET INVENTOR APPLICATION
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
                app = DirectCast(Interop.Marshal2.GetActiveObject("Inventor.Application"),
                                 Inventor.Application)
            Catch ex As Exception
                MessageBox.Show("Không lấy được Inventor.Application." & vbCrLf & vbCrLf & ex.Message,
                                "Frame Generator", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End Try

            Return app

        End Function

        '============================================================
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
                Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
                Try : System.Threading.Thread.Sleep(200) : Catch : End Try

                _sickTreatments.Clear()
                ScanFrameBrowser(doc)

                If _sickTreatments.Count = 0 Then
                    MessageBox.Show("Không tìm thấy Frame Treatment bị lỗi.",
                                    "Frame Generator", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Try : ForceCollapseAll() : Catch : End Try
                    Return
                End If

                Dim trimCount As Integer = 0
                Dim miterCount As Integer = 0
                Dim notchCount As Integer = 0
                Dim lengthenCount As Integer = 0
                Dim cornersCount As Integer = 0

                For Each item As SickFrameInfo In _sickTreatments
                    If item Is Nothing Then Continue For
                    Select Case item.TreatmentType
                        Case FrameTreatmentType.TrimExtend : trimCount += 1
                        Case FrameTreatmentType.Miter : miterCount += 1
                        Case FrameTreatmentType.Notch : notchCount += 1
                        Case FrameTreatmentType.LengthenShorten : lengthenCount += 1
                        Case FrameTreatmentType.Corners : cornersCount += 1
                    End Select
                Next

                Dim total As Integer = trimCount + miterCount + notchCount + lengthenCount + cornersCount

                Dim scanInfo As New StringBuilder
                scanInfo.AppendLine("Phát hiện " & total.ToString() & " Frame Treatment bị Sick.")
                scanInfo.AppendLine()

                If trimCount > 0 Then scanInfo.AppendLine("Trim / Extend     : " & trimCount.ToString())
                If miterCount > 0 Then scanInfo.AppendLine("Miter             : " & miterCount.ToString())
                If notchCount > 0 Then scanInfo.AppendLine("Notch             : " & notchCount.ToString())
                If lengthenCount > 0 Then scanInfo.AppendLine("Lengthen / Shorten: " & lengthenCount.ToString())
                If cornersCount > 0 Then scanInfo.AppendLine("Corners           : " & cornersCount.ToString())

                scanInfo.AppendLine()
                scanInfo.AppendLine("Đang tự động sửa...")

                MessageBox.Show(scanInfo.ToString(), "FRAME GENERATOR",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)

                ProcessSickTreatments(asmDoc)

            Catch ex As Exception
                MessageBox.Show("LỖI:" & vbCrLf & vbCrLf & ex.Message, "Frame Generator",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)

            Finally
                ' ⭐ Collapse All — chỉ chạy khi DISABLE_COLLAPSE = False
                Try : ForceCollapseAll() : Catch : End Try
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

                If ShouldSkipBranch(label) Then Return

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

                            Dim info As New SickFrameInfo
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
                                If childNode IsNot Nothing Then ScanBrowserNode(childNode)
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

            For Each item As SickFrameInfo In _sickTreatments
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
        Private Sub ProcessSickTreatments(ByVal asmDoc As AssemblyDocument)

            Dim deleteSuccess As Integer = 0
            Dim deleteFailed As Integer = 0

            Dim workList As New List(Of SickFrameInfo)

            For Each item As SickFrameInfo In _sickTreatments
                If item IsNot Nothing Then workList.Add(item)
            Next

            Dim deleteCmd As ControlDefinition = Nothing
            Try
                deleteCmd = _invApp.CommandManager.ControlDefinitions.Item("Delete")
            Catch
                deleteCmd = Nothing
            End Try

            For i As Integer = workList.Count - 1 To 0 Step -1

                Dim item As SickFrameInfo = workList(i)
                If item Is Nothing Then Continue For

                If ShouldSkipBranch(item.Label) Then Continue For

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

            '--- SCAN LẠI ---
            _sickTreatments.Clear()
            ScanFrameBrowser(asmDoc)

            Dim remainTrim As Integer = 0
            Dim remainMiter As Integer = 0
            Dim remainNotch As Integer = 0
            Dim remainLengthen As Integer = 0
            Dim remainCorners As Integer = 0

            For Each item As SickFrameInfo In _sickTreatments
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
        ' FORCE EXPAND — bỏ qua các nhánh không cần
        '============================================================
        Private Sub ForceExpandAllBrowser(ByVal doc As Inventor.Document)

            If doc Is Nothing Then Return

            Dim pane As Inventor.BrowserPane = Nothing
            Try
                pane = doc.BrowserPanes.Item("Model")
            Catch
                Try : pane = doc.BrowserPanes.ActivePane : Catch : End Try
            End Try

            If pane Is Nothing Then Return
            If pane.TopNode Is Nothing Then Return

            Try : ExpandNodeRecursive(pane.TopNode, 0) : Catch : End Try

            Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
            Try : System.Threading.Thread.Sleep(150) : Catch : End Try

        End Sub

        Private Sub ExpandNodeRecursive(ByVal node As Inventor.BrowserNode, ByVal depth As Integer)

            If node Is Nothing Then Return
            If depth > 100 Then Return

            Dim lbl As String = ""
            Try : lbl = node.BrowserNodeDefinition.Label : Catch : End Try

            If ShouldSkipBranch(lbl) Then Return

            Try
                If Not node.Expanded Then node.Expanded = True
            Catch
            End Try

            Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try

            Dim childCount As Integer = 0
            Try : childCount = node.BrowserNodes.Count : Catch : Return : End Try

            For i As Integer = 1 To childCount
                Try
                    Dim child As Inventor.BrowserNode = node.BrowserNodes.Item(i)
                    ExpandNodeRecursive(child, depth + 1)
                Catch
                End Try
            Next

        End Sub

        '============================================================
        ' ⭐ FORCE COLLAPSE — ĐÓNG TOÀN BỘ CÂY THƯ MỤC
        '
        '   Code đã sẵn sàng. Muốn chạy → đổi:
        '       Private Const DISABLE_COLLAPSE As Boolean = False
        '
        '   Cơ chế: đệ quy con TRƯỚC, collapse cha SAU.
        '   Chạy 4 pass để đảm bảo COM async hoàn tất.
        '============================================================
        Private Sub ForceCollapseAll()

            If DISABLE_COLLAPSE Then Return

            If _invApp Is Nothing Then Return

            Dim doc As Inventor.Document = Nothing
            Try : doc = _invApp.ActiveDocument : Catch : End Try
            If doc Is Nothing Then Return

            Dim pane As Inventor.BrowserPane = Nothing
            Try
                pane = doc.BrowserPanes.Item("Model")
            Catch
                Try : pane = doc.BrowserPanes.ActivePane : Catch : End Try
            End Try

            If pane Is Nothing Then Return
            If pane.TopNode Is Nothing Then Return

            For attempt As Integer = 1 To 4
                Try : CollapseRecursive(pane.TopNode) : Catch : End Try
                Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
                Try : System.Threading.Thread.Sleep(120) : Catch : End Try
            Next

            Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
            Try : System.Threading.Thread.Sleep(150) : Catch : End Try

        End Sub

        Private Sub CollapseRecursive(ByVal node As Inventor.BrowserNode)

            If node Is Nothing Then Return

            Dim children As Object = Nothing
            Try : children = node.BrowserNodes : Catch : End Try

            If children IsNot Nothing Then
                Dim count As Integer = 0
                Try : count = children.Count : Catch : End Try

                For i As Integer = 1 To count
                    Try : CollapseRecursive(children.Item(i)) : Catch : End Try
                Next
            End If

            Try
                If node.Expanded Then node.Expanded = False
            Catch
            End Try

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
                sb.AppendLine(" FRAME GENERATOR BROWSER DEBUG - INVENTOR 2025")
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