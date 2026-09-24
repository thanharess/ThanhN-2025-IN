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

        '============================================================
        ' INVENTOR 2020
        '============================================================
        Private Const SICK_DISPLAY_STATE As Integer = 46852

        '============================================================
        ' FRAME TREATMENT TYPE
        '============================================================
        Private Enum FrameTreatmentType
            Unknown = 0
            TrimExtend = 1
            Miter = 2
            Notch = 3
            LengthenShorten = 4
            Corners = 5
        End Enum

        '============================================================
        ' SICK FRAME INFO
        '============================================================
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
        ' GET INVENTOR APPLICATION
        '============================================================
        Private Function GetInventorApplication(
            ByVal Context As NameValueMap) As Inventor.Application

            Dim app As Inventor.Application = Nothing

            Try

                If Context IsNot Nothing Then

                    Try
                        app = DirectCast(
                            Context.Item("Application"),
                            Inventor.Application)
                    Catch
                    End Try

                End If

            Catch
            End Try

            If app IsNot Nothing Then
                Return app
            End If

            Try

                app = DirectCast(
                    Interop.Marshal2.GetActiveObject("Inventor.Application"),
                    Inventor.Application)

            Catch ex As Exception

                MessageBox.Show(
                    "Không lấy được Inventor.Application." &
                    vbCrLf & vbCrLf &
                    ex.Message,
                    "Frame Generator",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

                Return Nothing

            End Try

            Return app

        End Function

        '============================================================
        ' MAIN BUTTON
        '============================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim savedExpandState As Dictionary(Of String, Boolean) = Nothing
            Dim doc As Inventor.Document = Nothing

            Try
                _invApp = GetInventorApplication(Context)
                If _invApp Is Nothing Then Return

                doc = _invApp.ActiveDocument

                If doc Is Nothing Then
                    MessageBox.Show("Không có document đang mở.",
                                    "Frame Generator",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Hãy chạy trong Assembly.",
                                    "Frame Generator",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)

                '--- UPDATE ĐẦU ---
                ForceUpdate(asmDoc)

                '--- LƯU TRẠNG THÁI EXPANDED ---
                savedExpandState = SaveExpandState(doc)

                '--- FORCE EXPAND ---
                ForceExpandAllBrowser(doc)
                Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
                Try : System.Threading.Thread.Sleep(200) : Catch : End Try

                '--- SCAN ---
                _sickTreatments.Clear()
                ScanFrameBrowser(doc)

                If _sickTreatments.Count = 0 Then
                    MessageBox.Show("Không tìm thấy Frame Treatment bị lỗi.",
                                    "Frame Generator",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                '--- ĐẾM ---
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

                Dim msg As New StringBuilder
                msg.AppendLine("Phát hiện " & total.ToString() & " Frame Treatment bị Sick.")
                msg.AppendLine()
                If trimCount > 0 Then msg.AppendLine("Trim / Extend     : " & trimCount.ToString())
                If miterCount > 0 Then msg.AppendLine("Miter             : " & miterCount.ToString())
                If notchCount > 0 Then msg.AppendLine("Notch             : " & notchCount.ToString())
                If lengthenCount > 0 Then msg.AppendLine("Lengthen / Shorten: " & lengthenCount.ToString())
                If cornersCount > 0 Then msg.AppendLine("Corners           : " & cornersCount.ToString())
                msg.AppendLine()
                msg.AppendLine("Tiếp tục sửa?")

                Dim result As DialogResult =
                    MessageBox.Show(msg.ToString(), "FRAME GENERATOR",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

                If result <> DialogResult.Yes Then Return

                '--- XỬ LÝ ---
                ProcessSickTreatments(asmDoc)

            Catch ex As Exception
                MessageBox.Show("LỖI:" & vbCrLf & vbCrLf & ex.Message,
                                "Frame Generator",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)


            Finally
                '====================================================
                ' ⭐ COLLAPSE ALL — dùng lệnh có sẵn của Inventor
                '====================================================
                Try
                    If doc IsNot Nothing Then
                        CollapseAllViaInventorCommand()
                    End If
                Catch
                End Try
            End Try
        End Sub

        '============================================================
        ' FORCE UPDATE
        '============================================================
        Private Sub ForceUpdate(
            ByVal asmDoc As AssemblyDocument)

            If asmDoc Is Nothing Then
                Return
            End If

            Try
                asmDoc.Update2(True)
            Catch
            End Try

            Try
                asmDoc.Rebuild2()
            Catch
            End Try

            Try
                If _invApp IsNot Nothing Then
                    _invApp.ActiveView.Update()
                End If
            Catch
            End Try

        End Sub

        '============================================================
        ' SCAN FRAME BROWSER
        '============================================================
        Private Sub ScanFrameBrowser(
            ByVal doc As Inventor.Document)

            If doc Is Nothing Then
                Return
            End If

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

                If pane Is Nothing Then
                    Return
                End If

                If pane.TopNode Is Nothing Then
                    Return
                End If

                ScanBrowserNode(pane.TopNode)

            Catch
            End Try

        End Sub

        '============================================================
        ' SCAN NODE
        '============================================================
        Private Sub ScanBrowserNode(
            ByVal node As Inventor.BrowserNode)

            If node Is Nothing Then
                Return
            End If

            Try

                Dim label As String = ""
                Dim fullPath As String = ""
                Dim tooltip As String = ""
                Dim nativeObj As Object = Nothing
                Dim isSick As Boolean = False

                '----------------------------------------------------
                ' LABEL
                '----------------------------------------------------
                Try
                    label =
                        node.BrowserNodeDefinition.Label
                Catch
                    Try
                        label = node.FullPath
                    Catch
                        label = ""
                    End Try
                End Try

                '----------------------------------------------------
                ' PATH
                '----------------------------------------------------
                Try
                    fullPath = node.FullPath
                Catch
                    fullPath = label
                End Try

                '----------------------------------------------------
                ' TOOLTIP
                '----------------------------------------------------
                Try
                    tooltip =
                        node.BrowserNodeDefinition.
                        StateIconToolTipText
                Catch
                    tooltip = ""
                End Try

                '----------------------------------------------------
                ' NATIVE
                '----------------------------------------------------
                Try
                    nativeObj = node.NativeObject
                Catch
                    nativeObj = Nothing
                End Try

                '----------------------------------------------------
                ' DISPLAY STATE — mở rộng cho Inventor 2025
                '----------------------------------------------------
                Try
                    Dim ds As Object = node.BrowserNodeDefinition.DisplayState
                    Dim stateNumber As Integer = Convert.ToInt32(ds)

                    ' Cách 1: magic number cũ
                    If stateNumber = SICK_DISPLAY_STATE Then
                        isSick = True
                    End If

                    ' Cách 2: một số giá trị sick khác có thể gặp
                    If Not isSick Then
                        Select Case stateNumber
                            Case 46852, 46853, 46854, 46855
                                isSick = True
                        End Select
                    End If
                Catch
                    isSick = False
                End Try

                '----------------------------------------------------
                ' Fallback: đọc tooltip / native object nếu chưa phát hiện
                '----------------------------------------------------
                If Not isSick Then
                    Try
                        Dim t = node.BrowserNodeDefinition.StateIconToolTipText
                        If Not String.IsNullOrEmpty(t) Then
                            Dim tl = t.ToLowerInvariant()
                            If tl.Contains("error") OrElse
                               tl.Contains("sick") OrElse
                               tl.Contains("fail") OrElse
                               tl.Contains("lỗi") OrElse
                               tl.Contains("không thể") Then
                                isSick = True
                            End If
                        End If
                    Catch
                    End Try

                    Try
                        Dim obj = node.NativeObject
                        If obj IsNot Nothing Then
                            Try
                                Dim hs As Object = obj.HealthStatus
                                If hs IsNot Nothing Then
                                    Dim hstr = hs.ToString()
                                    If hstr.Contains("Sick") OrElse hstr.Contains("Error") Then
                                        isSick = True
                                    End If
                                End If
                            Catch
                            End Try
                        End If
                    Catch
                    End Try
                End If

                '----------------------------------------------------
                ' ADD SICK
                '----------------------------------------------------
                If isSick Then

                    Dim treatmentType As FrameTreatmentType =
                        DetectTreatmentType(
                            label,
                            fullPath,
                            tooltip)

                    If treatmentType <>
                        FrameTreatmentType.Unknown Then

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

                '----------------------------------------------------
                ' CHILDREN
                '----------------------------------------------------
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
                                    DirectCast(
                                        child,
                                        Inventor.BrowserNode)

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

            Dim l As String =
                If(label, "").ToLowerInvariant()

            Dim p As String =
                If(fullPath, "").ToLowerInvariant()

            Dim t As String =
                If(tooltip, "").ToLowerInvariant()

            '========================================================
            ' TRIM
            '========================================================
            If l.Contains("trim") OrElse
               p.Contains("trim") OrElse
               t.Contains("trim") Then

                Return FrameTreatmentType.TrimExtend

            End If

            '========================================================
            ' MITER
            '========================================================
            If l.Contains("miter") OrElse
               l.Contains("mitre") OrElse
               p.Contains("miter") OrElse
               p.Contains("mitre") OrElse
               t.Contains("miter") OrElse
               t.Contains("mitre") Then

                Return FrameTreatmentType.Miter

            End If

            '========================================================
            ' NOTCH
            '========================================================
            If l.Contains("notch") OrElse
               p.Contains("notch") OrElse
               t.Contains("notch") Then

                Return FrameTreatmentType.Notch

            End If

            '========================================================
            ' LENGTHEN / SHORTEN
            '========================================================
            If l.Contains("lengthen") OrElse
               l.Contains("shorten") OrElse
               p.Contains("lengthen") OrElse
               p.Contains("shorten") OrElse
               t.Contains("lengthen") OrElse
               t.Contains("shorten") Then

                Return FrameTreatmentType.LengthenShorten

            End If

            '========================================================
            ' CORNERS
            '
            ' Browser:
            '   Shop Corner
            '
            ' Command:
            '   Insert End Cap
            '
            ' Treatment chuẩn:
            '   Sharp Corners
            '========================================================
            If l.Contains("shop corner") OrElse
               p.Contains("shop corner") OrElse
               t.Contains("shop corner") OrElse
               l.Contains("insert end cap") OrElse
               p.Contains("insert end cap") OrElse
               t.Contains("insert end cap") OrElse
               l.Contains("sharp corners") OrElse
               p.Contains("sharp corners") OrElse
               t.Contains("sharp corners") Then

                Return FrameTreatmentType.Corners

            End If

            Return FrameTreatmentType.Unknown

        End Function

        '============================================================
        ' DUPLICATE
        '============================================================
        Private Function IsAlreadyAdded(
            ByVal node As Inventor.BrowserNode) As Boolean

            If node Is Nothing Then
                Return False
            End If

            For Each item As SickFrameInfo In _sickTreatments

                If item Is Nothing Then
                    Continue For
                End If

                Try

                    If Object.ReferenceEquals(
                        item.Node,
                        node) Then

                        Return True

                    End If

                Catch
                End Try

            Next

            Return False

        End Function

        '============================================================
        ' PROCESS
        '============================================================
        Private Sub ProcessSickTreatments(
            ByVal asmDoc As AssemblyDocument)

            Dim deleteSuccess As Integer = 0
            Dim deleteFailed As Integer = 0
            Dim cornersUpdated As Integer = 0

            Dim workList As New List(Of SickFrameInfo)

            For Each item As SickFrameInfo In _sickTreatments

                If item IsNot Nothing Then
                    workList.Add(item)
                End If

            Next

            '========================================================
            ' DELETE COMMAND
            '========================================================
            Dim deleteCmd As ControlDefinition = Nothing

            Try

                deleteCmd =
                    _invApp.CommandManager.
                    ControlDefinitions.Item("Delete")

            Catch
                deleteCmd = Nothing
            End Try

            '========================================================
            ' PROCESS
            '========================================================
            For i As Integer =
                workList.Count - 1 To 0 Step -1

                Dim item As SickFrameInfo =
                    workList(i)

                If item Is Nothing Then
                    Continue For
                End If

                '====================================================
                ' CORNERS
                '
                ' KHÔNG DELETE
                ' CHỈ UPDATE / REBUILD
                '====================================================
                If item.TreatmentType =
                    FrameTreatmentType.Corners Then

                    Try
                        asmDoc.Update2(True)
                    Catch
                    End Try

                    Try
                        asmDoc.Rebuild2()
                    Catch
                    End Try

                    Try
                        _invApp.ActiveView.Update()
                    Catch
                    End Try

                    cornersUpdated += 1

                    Continue For

                End If

                '====================================================
                ' CUT TREATMENT
                '====================================================
                If item.TreatmentType =
                    FrameTreatmentType.TrimExtend OrElse
                   item.TreatmentType =
                    FrameTreatmentType.Miter OrElse
                   item.TreatmentType =
                    FrameTreatmentType.Notch OrElse
                   item.TreatmentType =
                    FrameTreatmentType.LengthenShorten Then

                    If deleteCmd Is Nothing Then

                        deleteFailed += 1
                        Continue For

                    End If

                    '------------------------------------------------
                    ' CLEAR SELECTION
                    '------------------------------------------------
                    Try
                        asmDoc.SelectSet.Clear()
                    Catch
                    End Try

                    Try

                        Dim pane As BrowserPane =
                            asmDoc.BrowserPanes.Item("Model")

                        Try
                            pane.ClearSelection()
                        Catch
                        End Try

                    Catch
                    End Try

                    '------------------------------------------------
                    ' SELECT NODE
                    '------------------------------------------------
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

                    '------------------------------------------------
                    ' DELETE
                    '------------------------------------------------
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

                    '------------------------------------------------
                    ' UPDATE
                    '------------------------------------------------
                    Try
                        asmDoc.Update2(True)
                    Catch
                    End Try

                    Try
                        _invApp.ActiveView.Update()
                    Catch
                    End Try

                    '------------------------------------------------
                    ' CHECK
                    '------------------------------------------------
                    Dim stillExists As Boolean =
                        BrowserNodeStillExists(
                            asmDoc,
                            item.FullPath)

                    If stillExists Then
                        deleteFailed += 1
                    Else
                        deleteSuccess += 1
                    End If

                End If

            Next

            '========================================================
            ' UPDATE CUỐI
            '========================================================
            Try
                asmDoc.Update2(True)
            Catch
            End Try

            Try
                asmDoc.Rebuild2()
            Catch
            End Try

            Try
                _invApp.ActiveView.Update()
            Catch
            End Try

            '========================================================
            ' SCAN LẠI
            '========================================================
            _sickTreatments.Clear()

            ScanFrameBrowser(asmDoc)

            Dim remainTrim As Integer = 0
            Dim remainMiter As Integer = 0
            Dim remainNotch As Integer = 0
            Dim remainLengthen As Integer = 0
            Dim remainCorners As Integer = 0

            For Each item As SickFrameInfo In _sickTreatments

                If item Is Nothing Then
                    Continue For
                End If

                Select Case item.TreatmentType

                    Case FrameTreatmentType.TrimExtend
                        remainTrim += 1

                    Case FrameTreatmentType.Miter
                        remainMiter += 1

                    Case FrameTreatmentType.Notch
                        remainNotch += 1

                    Case FrameTreatmentType.LengthenShorten
                        remainLengthen += 1

                    Case FrameTreatmentType.Corners
                        remainCorners += 1

                End Select

            Next

            Dim remainTotal As Integer =
                _sickTreatments.Count

            '========================================================
            ' KẾT QUẢ GỌN
            '========================================================
            Dim result As New StringBuilder

            result.AppendLine(
                "========== FRAME GENERATOR ==========")

            result.AppendLine()

            result.AppendLine(
                "ĐÃ SỬA:")

            result.AppendLine(
                "Trim / Extend     : " &
                deleteSuccess.ToString())

            result.AppendLine(
                "Corners           : " &
                cornersUpdated.ToString())

            result.AppendLine(
                "Delete thất bại   : " &
                deleteFailed.ToString())

            result.AppendLine()

            result.AppendLine(
                "CÒN SICK:")

            result.AppendLine(
                "Trim / Extend     : " &
                remainTrim.ToString())

            result.AppendLine(
                "Miter             : " &
                remainMiter.ToString())

            result.AppendLine(
                "Notch             : " &
                remainNotch.ToString())

            result.AppendLine(
                "Lengthen / Shorten: " &
                remainLengthen.ToString())

            result.AppendLine(
                "Corners           : " &
                remainCorners.ToString())

            result.AppendLine()

            result.AppendLine(
                "Tổng còn Sick     : " &
                remainTotal.ToString())

            result.AppendLine()

            If remainTotal = 0 Then

                result.AppendLine(
                    "✓ HOÀN TẤT - KHÔNG CÒN SICK.")

            Else

                result.AppendLine(
                    "⚠ VẪN CÒN " &
                    remainTotal.ToString() &
                    " NODE SICK.")

            End If

            _sickTreatments.Clear()

            ' ShowDebugText(          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'result.ToString())''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        End Sub

        '============================================================
        ' CHECK NODE
        '============================================================
        Private Function BrowserNodeStillExists(
            ByVal doc As Inventor.Document,
            ByVal targetPath As String) As Boolean

            If doc Is Nothing Then
                Return False
            End If

            If String.IsNullOrEmpty(targetPath) Then
                Return False
            End If

            Try

                Dim pane As BrowserPane =
                    doc.BrowserPanes.Item("Model")

                If pane Is Nothing Then
                    Return False
                End If

                Return SearchBrowserPath(
                    pane.TopNode,
                    targetPath)

            Catch

                Return False

            End Try

        End Function

        '============================================================
        ' SEARCH PATH
        '============================================================
        Private Function SearchBrowserPath(
            ByVal node As Inventor.BrowserNode,
            ByVal targetPath As String) As Boolean

            If node Is Nothing Then
                Return False
            End If

            Try

                Dim currentPath As String = ""

                Try
                    currentPath = node.FullPath
                Catch
                End Try

                If String.Equals(
                    currentPath,
                    targetPath,
                    StringComparison.OrdinalIgnoreCase) Then

                    Return True

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
                                    DirectCast(
                                        child,
                                        Inventor.BrowserNode)

                                If SearchBrowserPath(
                                    childNode,
                                    targetPath) Then

                                    Return True

                                End If

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
        ' ON DEBUG
        '
        ' GIỮ LẠI CHO NÚT DEBUG
        '============================================================
        Public Sub OnDebug(
            ByVal Context As NameValueMap)

            Try

                _invApp =
                    GetInventorApplication(Context)

                If _invApp Is Nothing Then
                    Return
                End If

                Dim doc As Inventor.Document =
                    _invApp.ActiveDocument

                If doc Is Nothing Then
                    Return
                End If

                Dim sb As New StringBuilder

                sb.AppendLine(
                    "============================================================")

                sb.AppendLine(
                    " FRAME GENERATOR BROWSER DEBUG - INVENTOR 2020")

                sb.AppendLine(
                    "============================================================")

                sb.AppendLine()

                sb.AppendLine(
                    "DOCUMENT: " &
                    doc.DisplayName)

                sb.AppendLine()

                DebugBrowser(
                    doc,
                    sb)

                ShowDebugText(
                    sb.ToString())

            Catch ex As Exception

                MessageBox.Show(
                    ex.Message,
                    "Frame Debug",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub

        '============================================================
        ' DEBUG BROWSER
        '============================================================
        Private Sub DebugBrowser(
            ByVal doc As Inventor.Document,
            ByVal sb As StringBuilder)

            Try

                Dim pane As Inventor.BrowserPane = Nothing

                Try
                    pane =
                        doc.BrowserPanes.Item("Model")
                Catch

                    Try
                        pane =
                            doc.BrowserPanes.ActivePane
                    Catch
                    End Try

                End Try

                If pane Is Nothing Then

                    sb.AppendLine(
                        "Không lấy được Model Browser.")

                    Return

                End If

                DebugBrowserNode(
                    pane.TopNode,
                    sb,
                    0)

            Catch ex As Exception

                sb.AppendLine(
                    "BROWSER ERROR: " &
                    ex.Message)

            End Try

        End Sub

        '============================================================
        ' DEBUG NODE
        '============================================================
        Private Sub DebugBrowserNode(
            ByVal node As Inventor.BrowserNode,
            ByVal sb As StringBuilder,
            ByVal level As Integer)

            If node Is Nothing Then
                Return
            End If

            Try

                Dim indent As String =
                    New String(
                        " "c,
                        level * 2)

                Dim label As String = ""
                Dim path As String = ""
                Dim tooltip As String = ""
                Dim state As Integer = -1
                Dim nativeType As String = ""

                Try
                    label =
                        node.BrowserNodeDefinition.Label
                Catch
                    label = "?"
                End Try

                Try
                    path = node.FullPath
                Catch
                    path = ""
                End Try

                Try
                    tooltip =
                        node.BrowserNodeDefinition.
                        StateIconToolTipText
                Catch
                    tooltip = ""
                End Try

                Try

                    Dim ds As Object =
                        node.BrowserNodeDefinition.
                        DisplayState

                    state =
                        Convert.ToInt32(ds)

                Catch

                    state = -1

                End Try

                Try

                    Dim obj As Object =
                        node.NativeObject

                    If obj IsNot Nothing Then

                        nativeType =
                            obj.GetType().FullName

                    End If

                Catch

                    nativeType = ""

                End Try

                Dim treatmentType As FrameTreatmentType =
                    DetectTreatmentType(
                        label,
                        path,
                        tooltip)

                Dim important As Boolean =
                    treatmentType <>
                    FrameTreatmentType.Unknown OrElse
                    state = SICK_DISPLAY_STATE

                If important Then

                    sb.AppendLine()

                    sb.AppendLine(
                        indent &
                        "----------------------------------------")

                    sb.AppendLine(
                        indent &
                        "NAME: " &
                        label)

                    sb.AppendLine(
                        indent &
                        "TYPE: " &
                        GetTreatmentTypeName(
                            treatmentType))

                    sb.AppendLine(
                        indent &
                        "STATE: " &
                        state.ToString())

                    If state = SICK_DISPLAY_STATE Then

                        sb.AppendLine(
                            indent &
                            ">>> SICK NODE <<<")

                    End If

                    sb.AppendLine(
                        indent &
                        "TOOLTIP: " &
                        tooltip)

                    sb.AppendLine(
                        indent &
                        "NATIVE: " &
                        nativeType)

                    sb.AppendLine(
                        indent &
                        "PATH: " &
                        path)

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
                                    DirectCast(
                                        child,
                                        Inventor.BrowserNode)

                                If childNode IsNot Nothing Then

                                    DebugBrowserNode(
                                        childNode,
                                        sb,
                                        level + 1)

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
        ' TYPE NAME
        '============================================================
        Private Function GetTreatmentTypeName(
            ByVal treatmentType As FrameTreatmentType) As String

            Select Case treatmentType

                Case FrameTreatmentType.TrimExtend
                    Return "TRIM / EXTEND"

                Case FrameTreatmentType.Miter
                    Return "MITER / MITRE"

                Case FrameTreatmentType.Notch
                    Return "NOTCH"

                Case FrameTreatmentType.LengthenShorten
                    Return "LENGTHEN / SHORTEN"

                Case FrameTreatmentType.Corners
                    Return "CORNERS"

                Case Else
                    Return "UNKNOWN"

            End Select

        End Function

        '============================================================
        ' DEBUG FORM
        '============================================================
        Private Sub ShowDebugText(
            ByVal text As String)

            Dim f As New System.Windows.Forms.Form

            f.Text =
                "FRAME GENERATOR DEBUG"

            f.Width = 1200
            f.Height = 800

            Dim tb As New System.Windows.Forms.TextBox

            tb.Multiline = True
            tb.ReadOnly = True
            tb.ScrollBars =
                System.Windows.Forms.ScrollBars.Both

            tb.WordWrap = False
            tb.Dock =
                System.Windows.Forms.DockStyle.Fill

            tb.Font =
                New System.Drawing.Font(
                    "Consolas",
                    10.0F)

            tb.Text = text

            f.Controls.Add(tb)

            f.StartPosition =
                System.Windows.Forms.FormStartPosition.CenterScreen

            f.ShowDialog()

        End Sub
        '============================================================
        ' ⭐ FORCE EXPAND TOÀN BỘ BROWSER (fix node collapsed)
        '    Inventor chỉ load con khi node được expand.
        '    Nếu không expand → scan không thấy node sick bên trong.
        '============================================================
        Private Sub ForceExpandAllBrowser(ByVal doc As Inventor.Document)
            If doc Is Nothing Then Return

            Dim pane As Inventor.BrowserPane = Nothing
            Try
                pane = doc.BrowserPanes.Item("Model")
            Catch
                Try
                    pane = doc.BrowserPanes.ActivePane
                Catch
                End Try
            End Try

            If pane Is Nothing Then Return
            If pane.TopNode Is Nothing Then Return

            Try
                ExpandNodeRecursive(pane.TopNode, 0)
            Catch
            End Try

            ' Cho UI thread xử lý nốt sự kiện expand
            Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
            Try : System.Threading.Thread.Sleep(150) : Catch : End Try
        End Sub


        Private Sub ExpandNodeRecursive(ByVal node As Inventor.BrowserNode,
                                        ByVal depth As Integer)
            If node Is Nothing Then Return
            If depth > 100 Then Return   ' tránh đệ quy vô hạn

            '----- 1. Expand chính node này -----
            Try
                If Not node.Expanded Then
                    node.Expanded = True
                End If
            Catch
            End Try

            ' Cho Inventor kịp nạp con vào browser tree
            Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try

            '----- 2. Duyệt con -----
            Dim childCount As Integer = 0
            Try
                childCount = node.BrowserNodes.Count
            Catch
                Return
            End Try

            For i As Integer = 1 To childCount
                Try
                    Dim child As Inventor.BrowserNode = node.BrowserNodes.Item(i)
                    ExpandNodeRecursive(child, depth + 1)
                Catch
                End Try
            Next
        End Sub

        '============================================================
        ' ⭐ LƯU TRẠNG THÁI EXPANDED CỦA TOÀN BỘ NODE
        '    Trả về Dictionary<FullPath, Expanded>
        '============================================================
        Private Function SaveExpandState(ByVal doc As Inventor.Document) As Dictionary(Of String, Boolean)
            Dim state As New Dictionary(Of String, Boolean)
            If doc Is Nothing Then Return state

            Dim pane As Inventor.BrowserPane = Nothing
            Try
                pane = doc.BrowserPanes.Item("Model")
            Catch
                Try : pane = doc.BrowserPanes.ActivePane : Catch : End Try
            End Try
            If pane Is Nothing Then Return state
            If pane.TopNode Is Nothing Then Return state

            Try
                SaveExpandStateRecursive(pane.TopNode, state)
            Catch
            End Try

            Return state
        End Function


        Private Sub SaveExpandStateRecursive(ByVal node As Inventor.BrowserNode,
                                             ByVal state As Dictionary(Of String, Boolean))
            If node Is Nothing Then Return

            Dim key As String = ""
            Try : key = node.FullPath : Catch : End Try

            If Not String.IsNullOrEmpty(key) Then
                Try
                    state(key) = node.Expanded
                Catch
                End Try
            End If

            Dim childCount As Integer = 0
            Try : childCount = node.BrowserNodes.Count : Catch : Return : End Try

            For i As Integer = 1 To childCount
                Try
                    Dim child As Inventor.BrowserNode = node.BrowserNodes.Item(i)
                    SaveExpandStateRecursive(child, state)
                Catch
                End Try
            Next
        End Sub


        '============================================================
        ' ⭐ RESTORE TRẠNG THÁI EXPANDED
        '============================================================
        Private Sub RestoreExpandState(ByVal doc As Inventor.Document,
                                       ByVal state As Dictionary(Of String, Boolean))
            If doc Is Nothing OrElse state Is Nothing Then Return
            If state.Count = 0 Then Return

            Dim pane As Inventor.BrowserPane = Nothing
            Try
                pane = doc.BrowserPanes.Item("Model")
            Catch
                Try : pane = doc.BrowserPanes.ActivePane : Catch : End Try
            End Try
            If pane Is Nothing Then Return
            If pane.TopNode Is Nothing Then Return

            Try
                RestoreExpandStateRecursive(pane.TopNode, state)
            Catch
            End Try

            Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
        End Sub


        Private Sub RestoreExpandStateRecursive(ByVal node As Inventor.BrowserNode,
                                                ByVal state As Dictionary(Of String, Boolean))
            If node Is Nothing Then Return

            Dim key As String = ""
            Try : key = node.FullPath : Catch : End Try

            If Not String.IsNullOrEmpty(key) AndAlso state.ContainsKey(key) Then
                Try
                    node.Expanded = state(key)
                Catch
                End Try
            End If

            Dim childCount As Integer = 0
            Try : childCount = node.BrowserNodes.Count : Catch : Return : End Try

            For i As Integer = 1 To childCount
                Try
                    Dim child As Inventor.BrowserNode = node.BrowserNodes.Item(i)
                    RestoreExpandStateRecursive(child, state)
                Catch
                End Try
            Next
        End Sub
        '============================================================
        ' ⭐ RESTORE AN TOÀN
        '    Thử restore theo FullPath.
        '    Nếu FullPath đã đổi do xóa node → collapse toàn bộ.
        '============================================================
        Private Sub RestoreExpandStateSafe(ByVal doc As Inventor.Document,
                                           ByVal state As Dictionary(Of String, Boolean))
            If doc Is Nothing Then Return

            Dim pane As Inventor.BrowserPane = Nothing
            Try
                pane = doc.BrowserPanes.Item("Model")
            Catch
                Try : pane = doc.BrowserPanes.ActivePane : Catch : End Try
            End Try
            If pane Is Nothing Then Return
            If pane.TopNode Is Nothing Then Return

            '--- Bước 1: đếm số node khớp path ---
            Dim matched As Integer = 0
            Dim totalNodes As Integer = 0

            If state IsNot Nothing AndAlso state.Count > 0 Then
                Try
                    CountMatchedNodes(pane.TopNode, state, matched, totalNodes)
                Catch
                End Try
            End If

            '--- Bước 2: nếu khớp < 50% → collapse toàn bộ ---
            Dim useCollapseAll As Boolean = False

            If state Is Nothing OrElse state.Count = 0 Then
                useCollapseAll = True
            ElseIf totalNodes > 0 AndAlso (matched * 100 \ totalNodes) < 50 Then
                useCollapseAll = True
            End If

            If useCollapseAll Then
                CollapseAllNodes(pane.TopNode)
            Else
                Try
                    RestoreExpandStateRecursive(pane.TopNode, state)
                Catch
                    CollapseAllNodes(pane.TopNode)
                End Try
            End If

            Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
        End Sub


        Private Sub CountMatchedNodes(ByVal node As Inventor.BrowserNode,
                                      ByVal state As Dictionary(Of String, Boolean),
                                      ByRef matched As Integer,
                                      ByRef total As Integer)
            If node Is Nothing Then Return

            total += 1

            Dim key As String = ""
            Try : key = node.FullPath : Catch : End Try

            If Not String.IsNullOrEmpty(key) AndAlso state.ContainsKey(key) Then
                matched += 1
            End If

            Dim childCount As Integer = 0
            Try : childCount = node.BrowserNodes.Count : Catch : Return : End Try

            For i As Integer = 1 To childCount
                Try
                    CountMatchedNodes(node.BrowserNodes.Item(i), state, matched, total)
                Catch
                End Try
            Next
        End Sub


        '============================================================
        ' ⭐ COLLAPSE TOÀN BỘ (trừ TopNode)
        '============================================================
        Private Sub CollapseAllNodes(ByVal node As Inventor.BrowserNode)
            If node Is Nothing Then Return

            Dim childCount As Integer = 0
            Try : childCount = node.BrowserNodes.Count : Catch : Return : End Try

            For i As Integer = 1 To childCount
                Try
                    Dim child As Inventor.BrowserNode = node.BrowserNodes.Item(i)
                    Try
                        If child.Expanded Then child.Expanded = False
                    Catch
                    End Try
                    CollapseAllNodes(child)
                Catch
                End Try
            Next
        End Sub
        '============================================================
        ' ⭐ GỌI LỆNH COLLAPSE ALL CỦA INVENTOR
        '    Đây là lệnh chính chủ trong menu chuột phải browser.
        '============================================================
        Private Sub CollapseAllViaInventorCommand()
            If _invApp Is Nothing Then Return

            '--- Bước 1: thử lệnh Collapse All chuẩn ---
            Dim cmd As ControlDefinition = Nothing

            Try
                cmd = _invApp.CommandManager.ControlDefinitions.Item("CollapseAllCmd")
            Catch
                cmd = Nothing
            End Try

            If cmd Is Nothing Then
                Try
                    cmd = _invApp.CommandManager.ControlDefinitions.Item("BrowserCollapseAllCmd")
                Catch
                    cmd = Nothing
                End Try
            End If

            If cmd Is Nothing Then
                Try
                    cmd = _invApp.CommandManager.ControlDefinitions.Item("CollapseAll")
                Catch
                    cmd = Nothing
                End Try
            End If

            '--- Nếu không có lệnh có sẵn → quét toàn bộ control definitions ---
            If cmd Is Nothing Then
                Try
                    Dim all As ControlDefinitions = _invApp.CommandManager.ControlDefinitions
                    For i As Integer = 1 To all.Count
                        Try
                            Dim c As ControlDefinition = all.Item(i)
                            Dim nm As String = ""
                            Try : nm = c.InternalName : Catch : End Try
                            If Not String.IsNullOrEmpty(nm) AndAlso
                               (nm.IndexOf("CollapseAll", StringComparison.OrdinalIgnoreCase) >= 0) Then
                                cmd = c
                                Exit For
                            End If
                        Catch
                        End Try
                    Next
                Catch
                End Try
            End If

            '--- Chạy lệnh ---
            If cmd IsNot Nothing Then
                Try
                    If cmd.Enabled Then
                        cmd.Execute()
                        Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
                        Try : System.Threading.Thread.Sleep(100) : Catch : End Try
                        Return
                    End If
                Catch
                End Try
            End If

            '--- Fallback: tự collapse qua code ---
            Try
                Dim doc = _invApp.ActiveDocument
                If doc IsNot Nothing Then
                    Dim pane As Inventor.BrowserPane = Nothing
                    Try
                        pane = doc.BrowserPanes.Item("Model")
                    Catch
                        Try : pane = doc.BrowserPanes.ActivePane : Catch : End Try
                    End Try
                    If pane IsNot Nothing AndAlso pane.TopNode IsNot Nothing Then
                        CollapseAllNodes(pane.TopNode)
                    End If
                End If
            Catch
            End Try
        End Sub
    End Module

End Namespace
