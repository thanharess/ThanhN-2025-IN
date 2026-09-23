Option Explicit On
Option Strict Off

Imports Inventor
Imports System
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Collections.Generic

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep.constraint

    Public Module Ass_LG_C_2

        '==========================================================
        ' INVENTOR
        '==========================================================

        Private invApp As Inventor.Application = Nothing

        '==========================================================
        ' DOCUMENT ĐANG EDIT
        '==========================================================

        Private activeDoc As Inventor.Document = Nothing

        '==========================================================
        ' OCCURRENCE ĐÃ BẬT WORK FEATURES
        '==========================================================

        Private activeOccurrences As New List(Of ComponentOccurrence)

        '==========================================================
        ' USER INPUT EVENTS
        '==========================================================

        Private userInputEvents As UserInputEvents = Nothing

        '==========================================================
        ' TRẠNG THÁI
        '==========================================================

        Private isRunning As Boolean = False


        '==========================================================
        ' MAIN BUTTON
        '==========================================================

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try

                '--------------------------------------------------
                ' LẤY INVENTOR
                '--------------------------------------------------

                invApp = CType(
                        Interop.Marshal2.GetActiveObject("Inventor.Application"),
                    Inventor.Application
                )


                '--------------------------------------------------
                ' KIỂM TRA DOCUMENT
                '--------------------------------------------------

                If invApp.ActiveDocument Is Nothing Then

                    MessageBox.Show(
                        "Không có Document đang mở!",
                        "Show Planes / Axes",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    )

                    Exit Sub

                End If


                '--------------------------------------------------
                ' LẤY DOCUMENT ĐANG EDIT
                '--------------------------------------------------

                activeDoc = GetActiveEditDocument()

                If activeDoc Is Nothing Then
                    activeDoc = invApp.ActiveDocument
                End If


                '--------------------------------------------------
                ' KIỂM TRA ASSEMBLY
                '--------------------------------------------------

                If activeDoc.DocumentType <>
                    DocumentTypeEnum.kAssemblyDocumentObject Then

                    MessageBox.Show(
                        "Hãy chạy lệnh trong Assembly!",
                        "Show Planes / Axes",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    )

                    Exit Sub

                End If


                '--------------------------------------------------
                ' NẾU ĐANG CHẠY -> TẮT TẤT CẢ
                '--------------------------------------------------

                If isRunning Then

                    HideAllWorkFeatures()

                    Exit Sub

                End If


                activeOccurrences.Clear()


                '==================================================
                ' CHỌN NHIỀU OCCURRENCE
                '
                ' Click nhiều Component
                ' ESC = kết thúc chọn
                '==================================================

                'Dim picker As New OccurrencePicker()
                Dim picker As New OccurrencePicker(AddressOf ShowWorkFeaturesImmediately)

                Dim selectedOccurrences As List(Of ComponentOccurrence) =
                    picker.Pick(
                        invApp,
                        activeDoc
                    )


                '--------------------------------------------------
                ' ESC / CANCEL
                '--------------------------------------------------

                If selectedOccurrences Is Nothing Then
                    Exit Sub
                End If

                If selectedOccurrences.Count = 0 Then
                    Exit Sub
                End If


                '--------------------------------------------------
                ' HIỆN PLANE + AXIS CHO TẤT CẢ COMPONENT ĐÃ CHỌN
                '--------------------------------------------------

                For Each occ As ComponentOccurrence In selectedOccurrences

                    ShowWorkFeatures(
                        occ,
                        True,
                        True
                    )

                Next


                isRunning = True


                '--------------------------------------------------
                ' USER INPUT EVENTS
                '--------------------------------------------------

                userInputEvents =
                    invApp.CommandManager.UserInputEvents

                RemoveHandler userInputEvents.OnTerminateCommand,
                    AddressOf UserInputEvents_OnTerminateCommand

                AddHandler userInputEvents.OnTerminateCommand,
                    AddressOf UserInputEvents_OnTerminateCommand


                '--------------------------------------------------
                ' UPDATE
                '--------------------------------------------------

                Try
                    activeDoc.Update()
                Catch
                End Try

                Try
                    invApp.ActiveDocument.Update()
                Catch
                End Try

                Try
                    invApp.ActiveView.Update()
                Catch
                End Try


            Catch ex As Exception

                isRunning = False

                MessageBox.Show(
                    "Lỗi: " & ex.Message,
                    "Show Planes / Axes",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                )

            End Try

        End Sub
<<<<<<< HEAD
        Public Sub HideWorkFeaturesFromExternal()
            Try
                HideAllWorkFeatures()
            Catch
            End Try
        End Sub
=======

>>>>>>> c096192 (Add project files.)

        '==========================================================
        ' ACTIVE EDIT DOCUMENT
        '==========================================================
        Private Sub ShowWorkFeaturesImmediately(ByVal occ As ComponentOccurrence)

            ShowWorkFeatures(occ, True, True)

            Try
                activeDoc.Update()
            Catch
            End Try

            Try
                invApp.ActiveView.Update()
            Catch
            End Try

        End Sub
        Private Function GetActiveEditDocument() _
            As Inventor.Document

            Try

                Dim doc As Inventor.Document =
                    invApp.ActiveEditDocument

                If doc IsNot Nothing Then
                    Return doc
                End If

            Catch
            End Try


            Try
                Return invApp.ActiveDocument
            Catch
                Return Nothing
            End Try

        End Function


        '==========================================================
        ' HIỆN WORK PLANES / WORK AXES
        '==========================================================

        Private Sub ShowWorkFeatures(
            ByVal occ As ComponentOccurrence,
            ByVal showPlanes As Boolean,
            ByVal showAxes As Boolean)

            Try

                If occ Is Nothing Then
                    Exit Sub
                End If


                Dim def As ComponentDefinition =
                    occ.Definition


                '==================================================
                ' WORK PLANES
                '==================================================

                If showPlanes Then

                    For Each wp As WorkPlane In def.WorkPlanes

                        Try
                            wp.Visible = True
                        Catch
                        End Try

                    Next

                End If


                '==================================================
                ' WORK AXES
                '==================================================

                If showAxes Then

                    For Each wa As WorkAxis In def.WorkAxes

                        Try
                            wa.Visible = True
                        Catch
                        End Try

                    Next

                End If


                '==================================================
                ' UPDATE COMPONENT
                '==================================================

                Try

                    If def.Document IsNot Nothing Then
                        def.Document.Update()
                    End If

                Catch
                End Try


                '==================================================
                ' LƯU OCCURRENCE
                '==================================================

                If Not activeOccurrences.Contains(occ) Then
                    activeOccurrences.Add(occ)
                End If


            Catch
            End Try

        End Sub


        '==========================================================
        ' TẮT WORK FEATURES
        '==========================================================

        Private Sub HideAllWorkFeatures()

            Try

                For Each occ As ComponentOccurrence _
                    In activeOccurrences

                    Try

                        Dim def As ComponentDefinition =
                            occ.Definition


                        '------------------------------------------
                        ' PLANES
                        '------------------------------------------

                        For Each wp As WorkPlane In def.WorkPlanes

                            Try
                                wp.Visible = False
                            Catch
                            End Try

                        Next


                        '------------------------------------------
                        ' AXES
                        '------------------------------------------

                        For Each wa As WorkAxis In def.WorkAxes

                            Try
                                wa.Visible = False
                            Catch
                            End Try

                        Next


                        '------------------------------------------
                        ' UPDATE
                        '------------------------------------------

                        Try
                            def.Document.Update()
                        Catch
                        End Try

                    Catch
                    End Try

                Next


                '--------------------------------------------------
                ' UPDATE ASSEMBLY
                '--------------------------------------------------

                Try

                    If invApp IsNot Nothing Then

                        If invApp.ActiveDocument IsNot Nothing Then
                            invApp.ActiveDocument.Update()
                        End If

                        If invApp.ActiveView IsNot Nothing Then
                            invApp.ActiveView.Update()
                        End If

                    End If

                Catch
                End Try


                activeOccurrences.Clear()


                '--------------------------------------------------
                ' GỠ EVENT
                '--------------------------------------------------

                If userInputEvents IsNot Nothing Then

                    RemoveHandler userInputEvents.OnTerminateCommand,
                        AddressOf UserInputEvents_OnTerminateCommand

                End If


                isRunning = False


            Catch

                isRunning = False

            End Try

        End Sub


        '==========================================================
        ' COMMAND TERMINATE
        '==========================================================

        Private Sub UserInputEvents_OnTerminateCommand(
            ByVal CommandName As String,
            ByVal Context As NameValueMap)

            Try

                If Not isRunning Then
                    Exit Sub
                End If


                If String.IsNullOrEmpty(CommandName) Then
                    Exit Sub
                End If


                Dim cmd As String =
                    CommandName.ToUpperInvariant()


                If cmd.Contains("CONSTRAINT") OrElse
<<<<<<< HEAD
                   cmd.Contains("CONSTRAIN") OrElse
                   cmd.Contains("ASSEMBLYINSERT") Then
=======
                   cmd.Contains("CONSTRAIN") Then
>>>>>>> c096192 (Add project files.)

                    HideAllWorkFeatures()

                End If


            Catch
            End Try

        End Sub


        '################################################################
        ' OCCURRENCE PICKER - MULTI SELECT
        '################################################################

        Private Class OccurrencePicker

            Private interaction As InteractionEvents = Nothing
            Private selectEvents As SelectEvents = Nothing

            Private selecting As Boolean = True

            Private inventorApp As Inventor.Application = Nothing
            Private document As Inventor.Document = Nothing

            '----------------------------------------------------------
            ' DANH SÁCH COMPONENT ĐÃ CHỌN
            '----------------------------------------------------------

            Private selectedOccurrences As New List(Of ComponentOccurrence)


            '==========================================================
            ' PICK
            '
            ' Click nhiều Component
            ' ESC = kết thúc
            '==========================================================

            Public Function Pick(
                ByVal app As Inventor.Application,
                ByVal doc As Inventor.Document) _
                As List(Of ComponentOccurrence)

                Try

                    inventorApp = app
                    document = doc

                    selecting = True

                    selectedOccurrences.Clear()


                    '--------------------------------------------------
                    ' CREATE INTERACTION EVENTS
                    '--------------------------------------------------

                    interaction =
                        inventorApp.CommandManager.CreateInteractionEvents()


                    interaction.SelectionActive = True

                    interaction.StatusBarText =
                        "Chọn nhiều Component để hiện Planes + Axes  |  ESC = Xong"


                    '--------------------------------------------------
                    ' SELECT EVENTS
                    '--------------------------------------------------

                    selectEvents =
                        interaction.SelectEvents


                    '--------------------------------------------------
                    ' FILTER COMPONENT
                    '--------------------------------------------------

                    selectEvents.AddSelectionFilter(
                        SelectionFilterEnum.kAssemblyOccurrenceFilter
                    )


                    '--------------------------------------------------
                    ' EVENT
                    '--------------------------------------------------

                    AddHandler selectEvents.OnSelect,
                        AddressOf SelectEvents_OnSelect

                    AddHandler interaction.OnTerminate,
                        AddressOf Interaction_OnTerminate


                    '--------------------------------------------------
                    ' START
                    '--------------------------------------------------

                    interaction.Start()


                    '==================================================
                    ' CHỜ USER
                    '
                    ' KHÔNG DỪNG SAU LẦN CLICK ĐẦU TIÊN
                    '
                    ' Chỉ ESC mới kết thúc
                    '==================================================

                    Do While selecting

                        inventorApp.UserInterfaceManager.DoEvents()

                    Loop


                    '--------------------------------------------------
                    ' STOP
                    '--------------------------------------------------

                    Try
                        interaction.StatusBarText = ""
                    Catch
                    End Try


                    Try
                        interaction.Stop()
                    Catch
                    End Try


                    '--------------------------------------------------
                    ' REMOVE HANDLERS
                    '--------------------------------------------------

                    If selectEvents IsNot Nothing Then

                        RemoveHandler selectEvents.OnSelect,
                            AddressOf SelectEvents_OnSelect

                    End If


                    If interaction IsNot Nothing Then

                        RemoveHandler interaction.OnTerminate,
                            AddressOf Interaction_OnTerminate

                    End If


                    selectEvents = Nothing
                    interaction = Nothing


                    '--------------------------------------------------
                    ' TRẢ VỀ DANH SÁCH COMPONENT
                    '--------------------------------------------------

                    Return selectedOccurrences


                Catch

                    Try

                        If interaction IsNot Nothing Then
                            interaction.Stop()
                        End If

                    Catch
                    End Try


                    If selectEvents IsNot Nothing Then

                        Try
                            RemoveHandler selectEvents.OnSelect,
                                AddressOf SelectEvents_OnSelect
                        Catch
                        End Try

                    End If


                    If interaction IsNot Nothing Then

                        Try
                            RemoveHandler interaction.OnTerminate,
                                AddressOf Interaction_OnTerminate
                        Catch
                        End Try

                    End If


                    selectEvents = Nothing
                    interaction = Nothing


                    Return selectedOccurrences

                End Try

            End Function


            '==========================================================
            ' ON SELECT
            '
            ' KHÔNG set selecting = False
            '
            ' Đây chính là điểm quan trọng để chọn được nhiều Part
            '==========================================================

            Private Sub SelectEvents_OnSelect(
                ByVal JustSelectedEntities As ObjectsEnumerator,
                ByVal SelectionDevice As SelectionDeviceEnum,
                ByVal ModelPosition As Inventor.Point,
                ByVal ViewPosition As Inventor.Point2d,
                ByVal CurrentView As Inventor.View)

                Try

                    If JustSelectedEntities Is Nothing Then
                        Exit Sub
                    End If


                    If JustSelectedEntities.Count <= 0 Then
                        Exit Sub
                    End If


                    '--------------------------------------------------
                    ' LẤY CÁC COMPONENT VỪA CLICK
                    '--------------------------------------------------

                    For i As Integer = 1 To JustSelectedEntities.Count

                        Dim obj As Object =
                            JustSelectedEntities.Item(i)


                        If TypeOf obj Is ComponentOccurrence Then

                            Dim occ As ComponentOccurrence =
                                CType(
                                    obj,
                                    ComponentOccurrence
                                )


                            '------------------------------------------
                            ' KHÔNG THÊM TRÙNG
                            '------------------------------------------

                            If Not ContainsOccurrence(occ) Then

                                selectedOccurrences.Add(occ)
                                If onOccurrenceSelected IsNot Nothing Then
                                    onOccurrenceSelected.Invoke(occ)
                                End If
                            End If

                        End If

                    Next


                    '--------------------------------------------------
                    ' CẬP NHẬT STATUS BAR
                    '--------------------------------------------------

                    Try

                        interaction.StatusBarText =
                            "Đã chọn " &
                            selectedOccurrences.Count.ToString() &
                            " Component  |  Chọn tiếp hoặc ESC = Xong"

                    Catch
                    End Try


                    '--------------------------------------------------
                    ' QUAN TRỌNG:
                    '
                    ' KHÔNG:
                    '
                    ' selecting = False
                    '
                    ' Vì nếu làm vậy chỉ chọn được 1 Part.
                    '--------------------------------------------------

                Catch
                End Try

            End Sub


            '==========================================================
            ' KIỂM TRA COMPONENT ĐÃ CÓ TRONG LIST CHƯA
            '==========================================================

            Private Function ContainsOccurrence(
                ByVal testOcc As ComponentOccurrence) _
                As Boolean

                Try

                    If testOcc Is Nothing Then
                        Return False
                    End If


                    For Each occ As ComponentOccurrence _
                        In selectedOccurrences

                        Try

                            If Object.ReferenceEquals(
                                occ,
                                testOcc) Then

                                Return True

                            End If

                        Catch
                        End Try

                    Next


                    Return False

                Catch

                    Return False

                End Try

            End Function


            '==========================================================
            ' INTERACTION TERMINATE
            '
            ' ESC
            '==========================================================

            Private Sub Interaction_OnTerminate()

                selecting = False

            End Sub
            Private ReadOnly onOccurrenceSelected As Action(Of ComponentOccurrence)

            Public Sub New(ByVal callback As Action(Of ComponentOccurrence))
                onOccurrenceSelected = callback
            End Sub
        End Class

    End Module

End Namespace
