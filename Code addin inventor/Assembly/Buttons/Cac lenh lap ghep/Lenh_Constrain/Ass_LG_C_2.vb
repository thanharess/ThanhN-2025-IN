Option Explicit On
Option Strict Off
Imports Inventor
Imports System
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Collections.Generic

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep.constraint
    Public Module Ass_LG_C_2
        Private invApp As Inventor.Application = Nothing
        Private activeDoc As Inventor.Document = Nothing
        Private activeOccurrences As New List(Of ComponentOccurrence)
        Private isRunning As Boolean = False

        Public ReadOnly Property IsWorkFeaturesVisible As Boolean
            Get
                Return ToolInventor2025.Globals.g_workFeaturesVisible
            End Get
        End Property

        Private autoHideTimer As System.Windows.Forms.Timer = Nothing
        Private Sub StartAutoHideTimer()
            Try
                StopAutoHideTimer()
                autoHideTimer = New System.Windows.Forms.Timer()
                autoHideTimer.Interval = 5 * 60 * 1000
                AddHandler autoHideTimer.Tick, Sub(sender As Object, e As EventArgs)
                                                   Try
                                                       StopAutoHideTimer()
                                                       If isRunning Then HideAllWorkFeatures()
                                                   Catch
                                                   End Try
                                               End Sub
                autoHideTimer.Start()
            Catch
            End Try
        End Sub

        Private Sub StopAutoHideTimer()
            Try
                If autoHideTimer IsNot Nothing Then
                    autoHideTimer.Stop()
                    autoHideTimer.Dispose()
                    autoHideTimer = Nothing
                End If
            Catch
            End Try
        End Sub

        Private pollTimer As System.Windows.Forms.Timer = Nothing
        Private pollTargetCmd As ControlDefinition = Nothing
        Private pollStartTime As DateTime = DateTime.MinValue
        Private pollSawActive As Boolean = False

        Public Sub StartMonitorConstrain(cmd As ControlDefinition)
            Try
                StopPollTimer()
                pollTargetCmd = cmd
                pollStartTime = DateTime.Now
                pollSawActive = False
                pollTimer = New System.Windows.Forms.Timer()
                pollTimer.Interval = 250
                AddHandler pollTimer.Tick, AddressOf PollTimer_Tick
                pollTimer.Start()
            Catch
            End Try
        End Sub

        Private Sub PollTimer_Tick(sender As Object, e As EventArgs)
            Try
                If Not isRunning Then
                    StopPollTimer()
                    Return
                End If
                If (DateTime.Now - pollStartTime).TotalMinutes > 10 Then
                    StopPollTimer()
                    HideAllWorkFeatures()
                    Return
                End If

                Dim activeCmdName As String = ""
                Try
                    If invApp IsNot Nothing AndAlso invApp.CommandManager IsNot Nothing Then
                        activeCmdName = invApp.CommandManager.ActiveCommand
                    End If
                Catch
                End Try

                Dim isConstrainActive As Boolean =
            activeCmdName = "AssemblyInsertConstraintCmd" OrElse
            activeCmdName = "AssemblyConstraintCmd" OrElse
            activeCmdName = "AssemblyConstrainCmd"

                If isConstrainActive Then
                    pollSawActive = True
                ElseIf pollSawActive Then
                    ' Dialog đã tắt → ẩn plane + tắt nút
                    StopPollTimer()
                    System.Threading.Thread.Sleep(150)
                    HideAllWorkFeatures()
                End If
            Catch
            End Try
        End Sub

        Private Sub StopPollTimer()
            Try
                If pollTimer IsNot Nothing Then
                    RemoveHandler pollTimer.Tick, AddressOf PollTimer_Tick
                    pollTimer.Stop()
                    pollTimer.Dispose()
                    pollTimer = Nothing
                End If
                pollTargetCmd = Nothing
                pollSawActive = False
            Catch
            End Try
        End Sub

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                invApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)

                If invApp.ActiveDocument Is Nothing Then
                    MessageBox.Show("Không có Document đang mở!", "Show Planes / Axes", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                activeDoc = GetActiveEditDocument()
                If activeDoc Is Nothing Then activeDoc = invApp.ActiveDocument

                If activeDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Hãy chạy lệnh trong Assembly!", "Show Planes / Axes", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                If isRunning Then
                    HideAllWorkFeatures()
                    Exit Sub
                End If

                ToolInventor2025.Globals.g_workFeaturesVisible = True
                activeOccurrences.Clear()

                Dim picker As New OccurrencePicker(AddressOf ShowWorkFeaturesImmediately)
                Dim selectedOccurrences As List(Of ComponentOccurrence) = picker.Pick(invApp, activeDoc)

                If selectedOccurrences Is Nothing OrElse selectedOccurrences.Count = 0 Then
                    ToolInventor2025.Globals.g_workFeaturesVisible = False
                    Exit Sub
                End If

                For Each occ As ComponentOccurrence In selectedOccurrences
                    ShowWorkFeatures(occ, True, True)
                Next

                isRunning = True
                StartAutoHideTimer()

                Try : activeDoc.Update() : Catch : End Try
                Try : invApp.ActiveDocument.Update() : Catch : End Try
                Try : invApp.ActiveView.Update() : Catch : End Try
            Catch ex As Exception
                isRunning = False
                ToolInventor2025.Globals.g_workFeaturesVisible = False
                MessageBox.Show("Lỗi: " & ex.Message, "Show Planes / Axes", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Public Sub HideWorkFeaturesFromExternal()
            Try
                HideAllWorkFeatures()
            Catch
            End Try
        End Sub

        Private Sub ShowWorkFeaturesImmediately(ByVal occ As ComponentOccurrence)
            ShowWorkFeatures(occ, True, True)
            Try : activeDoc.Update() : Catch : End Try
            Try : invApp.ActiveView.Update() : Catch : End Try
        End Sub

        Private Function GetActiveEditDocument() As Inventor.Document
            Try
                Dim doc As Inventor.Document = invApp.ActiveEditDocument
                If doc IsNot Nothing Then Return doc
            Catch
            End Try
            Try
                Return invApp.ActiveDocument
            Catch
                Return Nothing
            End Try
        End Function

        Private Sub ShowWorkFeatures(ByVal occ As ComponentOccurrence, ByVal showPlanes As Boolean, ByVal showAxes As Boolean)
            Try
                If occ Is Nothing Then Exit Sub
                Dim def As ComponentDefinition = occ.Definition

                If showPlanes Then
                    For Each wp As WorkPlane In def.WorkPlanes
                        Try : wp.Visible = True : Catch : End Try
                    Next
                End If

                If showAxes Then
                    For Each wa As WorkAxis In def.WorkAxes
                        Try : wa.Visible = True : Catch : End Try
                    Next
                End If

                Try
                    If def.Document IsNot Nothing Then def.Document.Update()
                Catch
                End Try

                If Not activeOccurrences.Contains(occ) Then activeOccurrences.Add(occ)
            Catch
            End Try
        End Sub

        Private Sub HideAllWorkFeatures()
            Try
                For Each occ As ComponentOccurrence In activeOccurrences
                    Try
                        Dim def As ComponentDefinition = occ.Definition
                        For Each wp As WorkPlane In def.WorkPlanes
                            Try : wp.Visible = False : Catch : End Try
                        Next
                        For Each wa As WorkAxis In def.WorkAxes
                            Try : wa.Visible = False : Catch : End Try
                        Next
                        Try : def.Document.Update() : Catch : End Try
                    Catch
                    End Try
                Next

                Try
                    If invApp IsNot Nothing Then
                        If invApp.ActiveDocument IsNot Nothing Then invApp.ActiveDocument.Update()
                        If invApp.ActiveView IsNot Nothing Then invApp.ActiveView.Update()
                    End If
                Catch
                End Try

                activeOccurrences.Clear()
                isRunning = False
                ToolInventor2025.Globals.g_workFeaturesVisible = False   ' tắt nút
                StopAutoHideTimer()
                StopPollTimer()
            Catch
                isRunning = False
                ToolInventor2025.Globals.g_workFeaturesVisible = False
                StopAutoHideTimer()
                StopPollTimer()
            End Try
        End Sub

        Private Class OccurrencePicker
            Private interaction As InteractionEvents = Nothing
            Private selectEvents As SelectEvents = Nothing
            Private selecting As Boolean = True
            Private inventorApp As Inventor.Application = Nothing
            Private document As Inventor.Document = Nothing
            Private selectedOccurrences As New List(Of ComponentOccurrence)
            Private ReadOnly onOccurrenceSelected As Action(Of ComponentOccurrence)

            Public Sub New(ByVal callback As Action(Of ComponentOccurrence))
                onOccurrenceSelected = callback
            End Sub

            Public Function Pick(ByVal app As Inventor.Application, ByVal doc As Inventor.Document) As List(Of ComponentOccurrence)
                Try
                    inventorApp = app
                    document = doc
                    selecting = True
                    selectedOccurrences.Clear()

                    interaction = inventorApp.CommandManager.CreateInteractionEvents()
                    interaction.SelectionActive = True
                    interaction.StatusBarText = "Chọn nhiều Component để hiện Planes + Axes  |  ESC = Xong"

                    selectEvents = interaction.SelectEvents
                    selectEvents.AddSelectionFilter(SelectionFilterEnum.kAssemblyOccurrenceFilter)

                    AddHandler selectEvents.OnSelect, AddressOf SelectEvents_OnSelect
                    AddHandler interaction.OnTerminate, AddressOf Interaction_OnTerminate

                    interaction.Start()

                    Do While selecting
                        inventorApp.UserInterfaceManager.DoEvents()
                    Loop

                    Try : interaction.StatusBarText = "" : Catch : End Try
                    Try : interaction.Stop() : Catch : End Try

                    If selectEvents IsNot Nothing Then RemoveHandler selectEvents.OnSelect, AddressOf SelectEvents_OnSelect
                    If interaction IsNot Nothing Then RemoveHandler interaction.OnTerminate, AddressOf Interaction_OnTerminate

                    selectEvents = Nothing
                    interaction = Nothing
                    Return selectedOccurrences
                Catch
                    Try
                        If interaction IsNot Nothing Then interaction.Stop()
                    Catch
                    End Try
                    If selectEvents IsNot Nothing Then
                        Try : RemoveHandler selectEvents.OnSelect, AddressOf SelectEvents_OnSelect : Catch : End Try
                    End If
                    If interaction IsNot Nothing Then
                        Try : RemoveHandler interaction.OnTerminate, AddressOf Interaction_OnTerminate : Catch : End Try
                    End If
                    selectEvents = Nothing
                    interaction = Nothing
                    Return selectedOccurrences
                End Try
            End Function

            Private Sub SelectEvents_OnSelect(ByVal JustSelectedEntities As ObjectsEnumerator, ByVal SelectionDevice As SelectionDeviceEnum, ByVal ModelPosition As Inventor.Point, ByVal ViewPosition As Inventor.Point2d, ByVal CurrentView As Inventor.View)
                Try
                    If JustSelectedEntities Is Nothing OrElse JustSelectedEntities.Count <= 0 Then Exit Sub
                    For i As Integer = 1 To JustSelectedEntities.Count
                        Dim obj As Object = JustSelectedEntities.Item(i)
                        If TypeOf obj Is ComponentOccurrence Then
                            Dim occ As ComponentOccurrence = CType(obj, ComponentOccurrence)
                            If Not ContainsOccurrence(occ) Then
                                selectedOccurrences.Add(occ)
                                If onOccurrenceSelected IsNot Nothing Then onOccurrenceSelected.Invoke(occ)
                            End If
                        End If
                    Next
                    Try
                        interaction.StatusBarText = "Đã chọn " & selectedOccurrences.Count.ToString() & " Component  |  Chọn tiếp hoặc ESC = Xong"
                    Catch
                    End Try
                Catch
                End Try
            End Sub

            Private Function ContainsOccurrence(ByVal testOcc As ComponentOccurrence) As Boolean
                Try
                    If testOcc Is Nothing Then Return False
                    For Each occ As ComponentOccurrence In selectedOccurrences
                        Try
                            If Object.ReferenceEquals(occ, testOcc) Then Return True
                        Catch
                        End Try
                    Next
                    Return False
                Catch
                    Return False
                End Try
            End Function

            Private Sub Interaction_OnTerminate()
                selecting = False
            End Sub
        End Class
    End Module
End Namespace