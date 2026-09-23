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
        Public Sub OnExecute(
            ByVal Context As NameValueMap)

            Try

                _invApp = GetInventorApplication(Context)

                If _invApp Is Nothing Then
                    Return
                End If

                Dim doc As Inventor.Document =
                    _invApp.ActiveDocument

                If doc Is Nothing Then

                    MessageBox.Show(
                        "Không có document đang mở.",
                        "Frame Generator",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Return

                End If

                If doc.DocumentType <>
                    DocumentTypeEnum.kAssemblyDocumentObject Then

                    MessageBox.Show(
                        "Hãy chạy trong Assembly.",
                        "Frame Generator",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Return

                End If

                Dim asmDoc As AssemblyDocument =
                    CType(doc, AssemblyDocument)

                '====================================================
                ' UPDATE ĐẦU
                '====================================================
                ForceUpdate(asmDoc)

                '====================================================
                ' SCAN
                '====================================================
                _sickTreatments.Clear()

                ScanFrameBrowser(doc)

                '====================================================
                ' KHÔNG CÓ SICK
                '====================================================
                If _sickTreatments.Count = 0 Then

                    MessageBox.Show(
                        "Không tìm thấy Frame Treatment bị lỗi.",
                        "Frame Generator",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

                    Return

                End If

                '====================================================
                ' ĐẾM
                '====================================================
                Dim trimCount As Integer = 0
                Dim miterCount As Integer = 0
                Dim notchCount As Integer = 0
                Dim lengthenCount As Integer = 0
                Dim cornersCount As Integer = 0

                For Each item As SickFrameInfo In _sickTreatments

                    If item Is Nothing Then
                        Continue For
                    End If

                    Select Case item.TreatmentType

                        Case FrameTreatmentType.TrimExtend
                            trimCount += 1

                        Case FrameTreatmentType.Miter
                            miterCount += 1

                        Case FrameTreatmentType.Notch
                            notchCount += 1

                        Case FrameTreatmentType.LengthenShorten
                            lengthenCount += 1

                        Case FrameTreatmentType.Corners
                            cornersCount += 1

                    End Select

                Next

                '====================================================
                ' XÁC NHẬN
                '====================================================
                Dim total As Integer =
                    trimCount +
                    miterCount +
                    notchCount +
                    lengthenCount +
                    cornersCount

                Dim msg As New StringBuilder

                msg.AppendLine(
                    "Phát hiện " &
                    total.ToString() &
                    " Frame Treatment bị Sick.")

                msg.AppendLine()

                If trimCount > 0 Then
                    msg.AppendLine(
                        "Trim / Extend     : " &
                        trimCount.ToString())
                End If

                If miterCount > 0 Then
                    msg.AppendLine(
                        "Miter              : " &
                        miterCount.ToString())
                End If

                If notchCount > 0 Then
                    msg.AppendLine(
                        "Notch              : " &
                        notchCount.ToString())
                End If

                If lengthenCount > 0 Then
                    msg.AppendLine(
                        "Lengthen / Shorten : " &
                        lengthenCount.ToString())
                End If

                If cornersCount > 0 Then
                    msg.AppendLine(
                        "Corners            : " &
                        cornersCount.ToString())
                End If

                msg.AppendLine()
                msg.AppendLine("Tiếp tục sửa?")

                Dim result As DialogResult =
                    MessageBox.Show(
                        msg.ToString(),
                        "FRAME GENERATOR",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning)

                If result <> DialogResult.Yes Then
                    Return
                End If

                '====================================================
                ' XỬ LÝ
                '====================================================
                ProcessSickTreatments(asmDoc)

            Catch ex As Exception

                MessageBox.Show(
                    "LỖI:" &
                    vbCrLf & vbCrLf &
                    ex.Message,
                    "Frame Generator",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

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
                ' DISPLAY STATE
                '----------------------------------------------------
                Try

                    Dim ds As Object =
                        node.BrowserNodeDefinition.DisplayState

                    Dim stateNumber As Integer =
                        Convert.ToInt32(ds)

                    If stateNumber = SICK_DISPLAY_STATE Then
                        isSick = True
                    End If

                Catch

                    isSick = False

                End Try

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

    End Module

End Namespace
