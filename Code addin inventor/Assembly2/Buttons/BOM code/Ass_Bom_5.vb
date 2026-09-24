Option Explicit On
Option Strict Off

Imports Inventor
Imports System
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Drw = System.Drawing
Imports System.Collections.Generic

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode

    Public Module ass_bom_5

        Private invApp As Inventor.Application = Nothing
        Private activeDoc As Inventor.Document = Nothing
        Private activeOccurrences As New List(Of ComponentOccurrence)
        Private userInputEvents As UserInputEvents = Nothing
        Private isRunning As Boolean = False

        '==========================================================
        ' OPTIONS — nhớ lựa chọn của user
        '==========================================================
        Private lastShowPlanes As Boolean = True
        Private lastShowAxes As Boolean = True


        '==========================================================
        ' MAIN BUTTON
        '==========================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                invApp = CType(
                    Interop.Marshal2.GetActiveObject("Inventor.Application"),
                    Inventor.Application)

                If invApp.ActiveDocument Is Nothing Then
                    MessageBox.Show("Không có Document đang mở!",
                                    "Show Planes / Axes",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                activeDoc = GetActiveEditDocument()
                If activeDoc Is Nothing Then activeDoc = invApp.ActiveDocument

                If activeDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Hãy chạy lệnh trong Assembly!",
                                    "Show Planes / Axes",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                '=================================================
                ' ĐANG CHẠY → TẮT TẤT CẢ
                '=================================================
                If isRunning Then
                    HideAllWorkFeatures()
                    Exit Sub
                End If

                '=================================================
                ' FORM CHỌN PLANES / AXES / BOTH
                '=================================================
                Dim opt As ShowOptions = ShowOptionsForm()
                If opt Is Nothing OrElse opt.Cancelled Then Exit Sub

                lastShowPlanes = opt.ShowPlanes
                lastShowAxes = opt.ShowAxes

                activeOccurrences.Clear()

                '=================================================
                ' PICK NHIỀU OCCURRENCE
                '=================================================
                Dim picker As New OccurrencePicker(
                    AddressOf ShowWorkFeaturesImmediately,
                    opt.ShowPlanes, opt.ShowAxes)

                Dim selectedOccurrences As List(Of ComponentOccurrence) =
                    picker.Pick(invApp, activeDoc)

                If selectedOccurrences Is Nothing OrElse selectedOccurrences.Count = 0 Then
                    Exit Sub
                End If

                '=================================================
                ' HOÀN TẤT — set running flag
                '=================================================
                isRunning = True

                userInputEvents = invApp.CommandManager.UserInputEvents

                RemoveHandler userInputEvents.OnTerminateCommand,
                    AddressOf UserInputEvents_OnTerminateCommand

                AddHandler userInputEvents.OnTerminateCommand,
                    AddressOf UserInputEvents_OnTerminateCommand

                Try : activeDoc.Update() : Catch : End Try
                Try : invApp.ActiveDocument.Update() : Catch : End Try
                Try : invApp.ActiveView.Update() : Catch : End Try

            Catch ex As Exception
                isRunning = False
                MessageBox.Show("Lỗi: " & ex.Message,
                                "Show Planes / Axes",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '==========================================================
        ' OPTIONS CLASS
        '==========================================================
        Private Class ShowOptions
            Public ShowPlanes As Boolean = True
            Public ShowAxes As Boolean = True
            Public Cancelled As Boolean = True
        End Class


        '==========================================================
        ' FORM CHỌN PLANES / AXES / BOTH
        '==========================================================
        Private Function ShowOptionsForm() As ShowOptions
            Dim opt As New ShowOptions()
            Dim _okConfirmed As Boolean = False

            Using frm As New Form()
                frm.Text = "Show Planes / Axes — Tùy chọn"
                frm.AutoScaleMode = AutoScaleMode.None
                frm.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
                frm.ClientSize = New Drw.Size(560, 460)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.BackColor = Drw.Color.FromArgb(245, 245, 245)
                frm.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

                '===== HEADER =====
                Dim pnlHeader As New Panel()
                pnlHeader.Location = New Drw.Point(0, 0)
                pnlHeader.Size = New Drw.Size(560, 75)
                pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
                frm.Controls.Add(pnlHeader)

                Dim lblTitle As New Label()
                lblTitle.Text = "SHOW PLANES / AXES"
                lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
                lblTitle.ForeColor = Drw.Color.White
                lblTitle.Dock = DockStyle.Fill
                lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblTitle)

                Dim lblSub As New Label()
                lblSub.Text = "Hiện Work Planes / Work Axes cho Part đã chọn"
                lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
                lblSub.Dock = DockStyle.Bottom
                lblSub.Height = 20
                lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
                pnlHeader.Controls.Add(lblSub)

                '===== GROUP: CHỌN LOẠI =====
                Dim gb As New GroupBox() With {
                    .Text = "Chọn loại hiển thị",
                    .Location = New Drw.Point(15, 90),
                    .Size = New Drw.Size(530, 190),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb)

                Dim chkPlanes As New CheckBox() With {
                    .Text = "Work Planes   —   Mặt phẳng làm việc",
                    .Location = New Drw.Point(25, 40),
                    .Size = New Drw.Size(480, 28),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb.Controls.Add(chkPlanes)

                Dim chkAxes As New CheckBox() With {
                    .Text = "Work Axes   —   Trục làm việc",
                    .Location = New Drw.Point(25, 80),
                    .Size = New Drw.Size(480, 28),
                    .Checked = True,
                    .Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
                gb.Controls.Add(chkAxes)

                Dim lblHint As New Label() With {
                    .Text = "★ Có thể chọn cả hai hoặc chỉ một loại",
                    .Location = New Drw.Point(25, 125),
                    .Size = New Drw.Size(480, 20),
                    .ForeColor = Drw.Color.FromArgb(140, 140, 140),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)}
                gb.Controls.Add(lblHint)

                Dim lblHint2 As New Label() With {
                    .Text = "★ Sau khi bấm TIẾP TỤC, chọn các Part cần hiện",
                    .Location = New Drw.Point(25, 150),
                    .Size = New Drw.Size(480, 20),
                    .ForeColor = Drw.Color.FromArgb(140, 140, 140),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)}
                gb.Controls.Add(lblHint2)

                '===== GROUP: LƯU Ý =====
                Dim gb2 As New GroupBox() With {
                    .Text = "Lưu ý",
                    .Location = New Drw.Point(15, 290),
                    .Size = New Drw.Size(530, 90),
                    .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                    .BackColor = Drw.Color.White}
                frm.Controls.Add(gb2)

                Dim lblNote As New Label() With {
                    .Text = "• ESC để kết thúc chọn Part" & vbCrLf &
                            "• Bấm lại nút lệnh để TẮT toàn bộ Planes / Axes đã hiện" & vbCrLf &
                            "• Tự động tắt khi thoát lệnh Constraint",
                    .Location = New Drw.Point(20, 22),
                    .Size = New Drw.Size(500, 65),
                    .Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point),
                    .ForeColor = Drw.Color.FromArgb(80, 80, 80)}
                gb2.Controls.Add(lblNote)

                '===== NÚT =====
                Dim btnOK As New Button()
                btnOK.Text = "TIẾP TỤC CHỌN"
                btnOK.Size = New Drw.Size(170, 44)
                btnOK.Location = New Drw.Point(370, 400)
                btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
                btnOK.ForeColor = Drw.Color.White
                btnOK.FlatStyle = FlatStyle.Flat
                btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)

                AddHandler btnOK.Click, Sub()
                                            If Not chkPlanes.Checked AndAlso Not chkAxes.Checked Then
                                                MessageBox.Show("Phải chọn ít nhất 1 loại!",
                                                                "Lỗi",
                                                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                Return
                                            End If

                                            opt.ShowPlanes = chkPlanes.Checked
                                            opt.ShowAxes = chkAxes.Checked
                                            opt.Cancelled = False
                                            _okConfirmed = True

                                            frm.DialogResult = DialogResult.OK
                                            frm.Close()
                                        End Sub
                frm.Controls.Add(btnOK)

                Dim btnCancel As New Button()
                btnCancel.Text = "HỦY"
                btnCancel.Size = New Drw.Size(120, 44)
                btnCancel.Location = New Drw.Point(235, 400)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
                AddHandler btnCancel.Click, Sub()
                                                opt.Cancelled = True
                                                frm.DialogResult = DialogResult.Cancel
                                                frm.Close()
                                            End Sub
                frm.Controls.Add(btnCancel)

                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel

                AddHandler frm.FormClosing, Sub(sender, e)
                                                If Not _okConfirmed Then opt.Cancelled = True
                                            End Sub

                frm.ShowDialog()
            End Using

            Return opt
        End Function


        '==========================================================
        ' ACTIVE EDIT DOCUMENT
        '==========================================================
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


        '==========================================================
        ' SHOW WORK FEATURES NGAY LẬP TỨC (callback sau pick)
        '==========================================================
        Private Sub ShowWorkFeaturesImmediately(ByVal occ As ComponentOccurrence)
            ShowWorkFeatures(occ, lastShowPlanes, lastShowAxes)

            Try : activeDoc.Update() : Catch : End Try
            Try : invApp.ActiveView.Update() : Catch : End Try
        End Sub


        '==========================================================
        ' HIỆN WORK PLANES / WORK AXES
        '==========================================================
        Private Sub ShowWorkFeatures(ByVal occ As ComponentOccurrence,
                                     ByVal showPlanes As Boolean,
                                     ByVal showAxes As Boolean)
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
        Private Sub UserInputEvents_OnTerminateCommand(ByVal CommandName As String,
                                                       ByVal Context As NameValueMap)
            Try
                If Not isRunning Then Exit Sub
                If String.IsNullOrEmpty(CommandName) Then Exit Sub

                Dim cmd As String = CommandName.ToUpperInvariant()

                If cmd.Contains("CONSTRAINT") OrElse cmd.Contains("CONSTRAIN") Then
                    HideAllWorkFeatures()
                End If

            Catch
            End Try
        End Sub


        '################################################################
        ' OCCURRENCE PICKER
        '################################################################
        Private Class OccurrencePicker

            Private interaction As InteractionEvents = Nothing
            Private selectEvents As SelectEvents = Nothing
            Private selecting As Boolean = True
            Private inventorApp As Inventor.Application = Nothing
            Private document As Inventor.Document = Nothing
            Private selectedOccurrences As New List(Of ComponentOccurrence)

            Private ReadOnly onOccurrenceSelected As Action(Of ComponentOccurrence)
            Private ReadOnly showPlanes As Boolean
            Private ReadOnly showAxes As Boolean

            Public Sub New(ByVal callback As Action(Of ComponentOccurrence),
                           ByVal showPlanes As Boolean,
                           ByVal showAxes As Boolean)
                onOccurrenceSelected = callback
                Me.showPlanes = showPlanes
                Me.showAxes = showAxes
            End Sub


            Public Function Pick(ByVal app As Inventor.Application,
                                 ByVal doc As Inventor.Document) As List(Of ComponentOccurrence)
                Try
                    inventorApp = app
                    document = doc
                    selecting = True
                    selectedOccurrences.Clear()

                    interaction = inventorApp.CommandManager.CreateInteractionEvents()
                    interaction.SelectionActive = True

                    Dim whatShow As String = ""
                    If showPlanes AndAlso showAxes Then
                        whatShow = "Planes + Axes"
                    ElseIf showPlanes Then
                        whatShow = "Planes"
                    Else
                        whatShow = "Axes"
                    End If

                    interaction.StatusBarText =
                        "Chọn nhiều Component để hiện " & whatShow & "  |  ESC = Xong"

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

                    If selectEvents IsNot Nothing Then
                        RemoveHandler selectEvents.OnSelect, AddressOf SelectEvents_OnSelect
                    End If

                    If interaction IsNot Nothing Then
                        RemoveHandler interaction.OnTerminate, AddressOf Interaction_OnTerminate
                    End If

                    selectEvents = Nothing
                    interaction = Nothing

                    Return selectedOccurrences

                Catch
                    Try
                        If interaction IsNot Nothing Then interaction.Stop()
                    Catch
                    End Try

                    Try
                        If selectEvents IsNot Nothing Then
                            RemoveHandler selectEvents.OnSelect, AddressOf SelectEvents_OnSelect
                        End If
                    Catch
                    End Try

                    Try
                        If interaction IsNot Nothing Then
                            RemoveHandler interaction.OnTerminate, AddressOf Interaction_OnTerminate
                        End If
                    Catch
                    End Try

                    selectEvents = Nothing
                    interaction = Nothing

                    Return selectedOccurrences
                End Try
            End Function


            Private Sub SelectEvents_OnSelect(ByVal JustSelectedEntities As ObjectsEnumerator,
                                              ByVal SelectionDevice As SelectionDeviceEnum,
                                              ByVal ModelPosition As Inventor.Point,
                                              ByVal ViewPosition As Inventor.Point2d,
                                              ByVal CurrentView As Inventor.View)
                Try
                    If JustSelectedEntities Is Nothing Then Exit Sub
                    If JustSelectedEntities.Count <= 0 Then Exit Sub

                    For i As Integer = 1 To JustSelectedEntities.Count
                        Dim obj As Object = JustSelectedEntities.Item(i)

                        If TypeOf obj Is ComponentOccurrence Then
                            Dim occ As ComponentOccurrence = CType(obj, ComponentOccurrence)

                            If Not ContainsOccurrence(occ) Then
                                selectedOccurrences.Add(occ)
                                If onOccurrenceSelected IsNot Nothing Then
                                    onOccurrenceSelected.Invoke(occ)
                                End If
                            End If
                        End If
                    Next

                    Try
                        interaction.StatusBarText =
                            "Đã chọn " & selectedOccurrences.Count.ToString() &
                            " Component  |  Chọn tiếp hoặc ESC = Xong"
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