Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports System.Runtime.InteropServices
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.Drawdim

    Friend Class NativeMethodsaaaga
        <DllImport("user32.dll")>
        Public Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
        End Function
    End Class


    '=====================================================
    ' FORM — CÓ LOGIC KHÓA TỰ ĐỘNG
    '=====================================================
    Public Class DimCleanupFormaaaga
        Inherits Form

        Public ReadOnly Property DeleteHoleDim As Boolean
        Public ReadOnly Property DeleteHoleNote As Boolean
        Public ReadOnly Property DeleteAllDim As Boolean
        Public ReadOnly Property CenterDimArrange As Boolean
        Public ReadOnly Property ArrangeDim As Boolean
        Public ReadOnly Property CenterDim As Boolean
        Public ReadOnly Property AllViews As Boolean
        Public ReadOnly Property PickViews As Boolean
        Public ReadOnly Property Cancelled As Boolean

        Private chkDeleteHoleDim As CheckBox
        Private chkDeleteHoleNote As CheckBox
        Private chkDeleteAllDim As CheckBox
        Private chkCenterDimArrange As CheckBox
        Private chkArrangeDim As CheckBox
        Private chkCenterDim As CheckBox
        Private chkSelectAll As CheckBox
        Private rdoAllViews As RadioButton
        Private rdoPickViews As RadioButton
        Private btnOK As Button
        Private btnCancel As Button

        Public Sub New()
            _Cancelled = False

            Me.Text = "Xử lý Dimension bản vẽ"
            Me.AutoScaleMode = AutoScaleMode.None
            Me.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
            Me.FormBorderStyle = FormBorderStyle.FixedSingle
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.ClientSize = New Drw.Size(660, 640)
            Me.BackColor = Drw.Color.FromArgb(245, 245, 245)
            Me.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

            '===== HEADER =====
            Dim pnlHeader As New Panel()
            pnlHeader.Location = New Drw.Point(0, 0)
            pnlHeader.Size = New Drw.Size(660, 75)
            pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
            Me.Controls.Add(pnlHeader)

            Dim lblTitle As New Label()
            lblTitle.Text = "XỬ LÝ DIMENSION BẢN VẼ"
            lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblTitle.ForeColor = Drw.Color.White
            lblTitle.Dock = DockStyle.Fill
            lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
            pnlHeader.Controls.Add(lblTitle)

            Dim lblSub As New Label()
            lblSub.Text = "Xóa / Căn chỉnh / Arrange Dimension cho Drawing View"
            lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
            lblSub.Dock = DockStyle.Bottom
            lblSub.Height = 20
            lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
            pnlHeader.Controls.Add(lblSub)

            '===== CHECKBOX CHỌN TẤT CẢ =====
            chkSelectAll = New CheckBox() With {
                .Text = "Chọn tất cả / Bỏ chọn tất cả",
                .Location = New Drw.Point(20, 85),
                .Size = New Drw.Size(620, 25),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180)}
            AddHandler chkSelectAll.CheckedChanged, AddressOf OnSelectAllChanged
            Me.Controls.Add(chkSelectAll)

            '===== GROUP 1: PHẠM VI =====
            Dim gb1 As New GroupBox() With {
                .Text = "Phạm vi áp dụng",
                .Location = New Drw.Point(15, 120),
                .Size = New Drw.Size(630, 100),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gb1)

            rdoAllViews = New RadioButton() With {
                .Text = "Tất cả View trên Sheet",
                .Location = New Drw.Point(25, 30),
                .Size = New Drw.Size(280, 25),
                .Checked = True,
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            gb1.Controls.Add(rdoAllViews)

            rdoPickViews = New RadioButton() With {
                .Text = "Chọn View riêng lẻ (multi-select)",
                .Location = New Drw.Point(25, 62),
                .Size = New Drw.Size(400, 25),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            gb1.Controls.Add(rdoPickViews)

            '===== GROUP 2: XÓA =====
            Dim gb2 As New GroupBox() With {
                .Text = "Xóa",
                .Location = New Drw.Point(15, 230),
                .Size = New Drw.Size(630, 130),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(200, 60, 60),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gb2)

            chkDeleteHoleDim = MakeCheckBox(gb2, "Xóa Dimension lỗ (Diameter)", 30)
            chkDeleteHoleNote = MakeCheckBox(gb2, "Xóa Hole / Thread Note", 60)
            chkDeleteAllDim = MakeCheckBox(gb2, "Xóa TẤT CẢ Dimension (Linear + Angular + Hole)", 90)

            ' ⭐ KHI CHỌN "XÓA TẤT CẢ" → KHÓA 2 CÁI TRÊN
            AddHandler chkDeleteAllDim.CheckedChanged, Sub()
                                                           Dim locked As Boolean = chkDeleteAllDim.Checked

                                                           If locked Then
                                                               ' Bỏ tick 2 cái trên khi chọn tất cả
                                                               chkDeleteHoleDim.Checked = False
                                                               chkDeleteHoleNote.Checked = False
                                                           End If

                                                           chkDeleteHoleDim.Enabled = Not locked
                                                           chkDeleteHoleNote.Enabled = Not locked

                                                           UpdateCheckboxVisual(chkDeleteHoleDim)
                                                           UpdateCheckboxVisual(chkDeleteHoleNote)
                                                       End Sub

            '===== GROUP 3: CĂN CHỈNH =====
            Dim gb3 As New GroupBox() With {
                .Text = "Căn chỉnh",
                .Location = New Drw.Point(15, 370),
                .Size = New Drw.Size(630, 130),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gb3)

            chkCenterDim = MakeCheckBox(gb3, "Căn Dimension về giữa (Center Text)", 30)
            chkArrangeDim = MakeCheckBox(gb3, "Arrange Dimension (tự động sắp xếp)", 60)
            chkCenterDimArrange = MakeCheckBox(gb3, "Căn về giữa + Arrange (kết hợp)", 90)

            ' ⭐ KHI CHỌN "CĂN + ARRANGE" → KHÓA 2 CÁI TRÊN
            AddHandler chkCenterDimArrange.CheckedChanged, Sub()
                                                               Dim locked As Boolean = chkCenterDimArrange.Checked

                                                               If locked Then
                                                                   chkCenterDim.Checked = False
                                                                   chkArrangeDim.Checked = False
                                                               End If

                                                               chkCenterDim.Enabled = Not locked
                                                               chkArrangeDim.Enabled = Not locked

                                                               UpdateCheckboxVisual(chkCenterDim)
                                                               UpdateCheckboxVisual(chkArrangeDim)
                                                           End Sub

            '===== NÚT =====
            btnOK = New Button()
            btnOK.Text = "THỰC HIỆN"
            btnOK.Location = New Drw.Point(495, 585)
            btnOK.Size = New Drw.Size(150, 42)
            btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
            btnOK.ForeColor = Drw.Color.White
            btnOK.FlatStyle = FlatStyle.Flat
            btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            btnOK.DialogResult = DialogResult.OK
            Me.Controls.Add(btnOK)

            btnCancel = New Button()
            btnCancel.Text = "HỦY"
            btnCancel.Location = New Drw.Point(355, 585)
            btnCancel.Size = New Drw.Size(130, 42)
            btnCancel.FlatStyle = FlatStyle.Flat
            btnCancel.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            btnCancel.DialogResult = DialogResult.Cancel
            Me.Controls.Add(btnCancel)

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub


        '=========================================================
        ' UPDATE VISUAL CHO CHECKBOX BỊ KHÓA
        '=========================================================
        Private Sub UpdateCheckboxVisual(chk As CheckBox)
            If chk Is Nothing Then Return

            If chk.Enabled Then
                chk.ForeColor = Drw.Color.FromArgb(40, 40, 40)
            Else
                chk.ForeColor = Drw.Color.FromArgb(170, 170, 170)
            End If
        End Sub


        Private Function MakeCheckBox(ByVal parent As GroupBox,
                                      ByVal text As String,
                                      ByVal top As Integer) As CheckBox
            Dim chk As New CheckBox()
            chk.Text = text
            chk.Location = New Drw.Point(20, top)
            chk.Size = New Drw.Size(590, 25)
            chk.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            chk.ForeColor = Drw.Color.FromArgb(40, 40, 40)
            parent.Controls.Add(chk)
            Return chk
        End Function


        Private Sub OnSelectAllChanged(ByVal sender As Object, ByVal e As EventArgs)
            Dim chk As Boolean = chkSelectAll.Checked

            chkDeleteHoleDim.Checked = chk
            chkDeleteHoleNote.Checked = chk
            chkDeleteAllDim.Checked = chk
            chkCenterDim.Checked = chk
            chkArrangeDim.Checked = chk
            chkCenterDimArrange.Checked = chk
        End Sub


        Public Function ShowAndGet() As Boolean
            Dim result As DialogResult = Me.ShowDialog()
            If result <> DialogResult.OK Then
                _Cancelled = True
                Return False
            End If

            _DeleteHoleDim = chkDeleteHoleDim.Checked
            _DeleteHoleNote = chkDeleteHoleNote.Checked
            _DeleteAllDim = chkDeleteAllDim.Checked
            _CenterDim = chkCenterDim.Checked
            _ArrangeDim = chkArrangeDim.Checked
            _CenterDimArrange = chkCenterDimArrange.Checked
            _AllViews = rdoAllViews.Checked
            _PickViews = rdoPickViews.Checked
            Return True
        End Function
    End Class


    '=====================================================
    ' MODULE — GIỮ NGUYÊN LOGIC
    '=====================================================
    Public Module draw_3

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                Dim invApp As Inventor.Application = g_inventorApplication

                If invApp Is Nothing Then
                    MessageBox.Show("Không tìm thấy Inventor Application.", "Dim Cleanup",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If invApp.ActiveDocument Is Nothing OrElse
                   invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Chức năng này chỉ dùng cho bản vẽ Drawing.", "Dim Cleanup",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument = CType(invApp.ActiveDocument, DrawingDocument)
                Dim oSheet As Sheet = oDrawDoc.ActiveSheet

                Dim form As New DimCleanupFormaaaga()
                If Not form.ShowAndGet() Then Exit Sub

                '=====================================================
                ' DANH SÁCH VIEW
                '=====================================================
                Dim selectedViews As New List(Of DrawingView)

                If form.AllViews Then
                    For Each v As DrawingView In oSheet.DrawingViews
                        selectedViews.Add(v)
                    Next
                Else
                    Try
                        Dim mainHwnd As IntPtr = New IntPtr(invApp.MainFrameHWND)
                        If mainHwnd <> IntPtr.Zero Then
                            NativeMethodsaaaga.SetForegroundWindow(mainHwnd)
                            System.Threading.Thread.Sleep(150)
                        End If
                    Catch
                    End Try

                    Do
                        oDrawDoc.SelectSet.Clear()
                        Dim oView As DrawingView = Nothing
                        Try
                            oView = CType(
                                invApp.CommandManager.Pick(
                                    SelectionFilterEnum.kDrawingViewFilter,
                                    "Chọn View (Esc / Right-click để kết thúc)"),
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
                        If Not already Then selectedViews.Add(oView)
                    Loop
                End If

                If selectedViews.Count = 0 Then
                    MessageBox.Show("Chưa chọn View nào.", "Dim Cleanup",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim targetDims As List(Of DrawingDimension) =
                    GetDimensionsForViews(oSheet, selectedViews)

                Dim nHoleDim As Integer = 0
                Dim nHoleNote As Integer = 0
                Dim nAllDim As Integer = 0
                Dim nCenter As Integer = 0
                Dim nArrange As Integer = 0
                Dim nFail As Integer = 0

                '=====================================================
                ' 1. XÓA DIAMETER
                '=====================================================
                If form.DeleteHoleDim Then
                    Dim toDel As New List(Of DrawingDimension)
                    For Each oDim As DrawingDimension In targetDims
                        Try
                            If TypeOf oDim Is DiameterGeneralDimension Then
                                toDel.Add(oDim)
                            End If
                        Catch
                        End Try
                    Next
                    For Each oDim As DrawingDimension In toDel
                        Try
                            oDim.Delete()
                            nHoleDim += 1
                        Catch
                            nFail += 1
                        End Try
                    Next
                End If

                '=====================================================
                ' 2. XÓA HOLE / THREAD NOTE
                '=====================================================
                If form.DeleteHoleNote Then
                    Try
                        Dim toDel As New List(Of HoleThreadNote)

                        For Each htNote As HoleThreadNote In oSheet.DrawingNotes.HoleThreadNotes
                            Try
                                If form.AllViews Then
                                    toDel.Add(htNote)
                                    Continue For
                                End If

                                Dim noteView As DrawingView = GetViewFromHoleThreadNote(htNote)
                                If noteView Is Nothing Then Continue For

                                For Each v As DrawingView In selectedViews
                                    If v Is noteView Then
                                        toDel.Add(htNote)
                                        Exit For
                                    End If
                                Next

                            Catch
                            End Try
                        Next

                        For Each htNote As HoleThreadNote In toDel
                            Try
                                htNote.Delete()
                                nHoleNote += 1
                            Catch
                                nFail += 1
                            End Try
                        Next
                    Catch
                    End Try
                End If

                '=====================================================
                ' 3. XÓA TẤT CẢ DIM
                '=====================================================
                If form.DeleteAllDim Then
                    Dim toDel As New List(Of DrawingDimension)
                    For Each oDim As DrawingDimension In targetDims
                        toDel.Add(oDim)
                    Next
                    For Each oDim As DrawingDimension In toDel
                        Try
                            oDim.Delete()
                            nAllDim += 1
                        Catch
                            nFail += 1
                        End Try
                    Next

                    Try
                        Dim notes As New List(Of HoleThreadNote)
                        For Each ht As HoleThreadNote In oSheet.DrawingNotes.HoleThreadNotes
                            notes.Add(ht)
                        Next
                        For Each ht As HoleThreadNote In notes
                            Try
                                ht.Delete()
                                nAllDim += 1
                            Catch
                                nFail += 1
                            End Try
                        Next
                    Catch
                    End Try

                    Try
                        Dim baseSets As BaselineDimensionSets =
                            oSheet.DrawingDimensions.BaselineDimensionSets
                        Dim listBs As New List(Of BaselineDimensionSet)
                        For i As Integer = 1 To baseSets.Count
                            Try
                                listBs.Add(baseSets.Item(i))
                            Catch
                            End Try
                        Next
                        For Each bs As BaselineDimensionSet In listBs
                            Try
                                bs.Delete()
                                nAllDim += 1
                            Catch
                                nFail += 1
                            End Try
                        Next
                    Catch
                    End Try

                    Try
                        Dim chainSets As ChainDimensionSets =
                            oSheet.DrawingDimensions.ChainDimensionSets
                        Dim listCs As New List(Of ChainDimensionSet)
                        For i As Integer = 1 To chainSets.Count
                            Try
                                listCs.Add(chainSets.Item(i))
                            Catch
                            End Try
                        Next
                        For Each cs As ChainDimensionSet In listCs
                            Try
                                cs.Delete()
                                nAllDim += 1
                            Catch
                                nFail += 1
                            End Try
                        Next
                    Catch
                    End Try
                End If

                targetDims = GetDimensionsForViews(oSheet, selectedViews)

                '=====================================================
                ' 4. CENTER
                '=====================================================
                If form.CenterDim AndAlso Not form.CenterDimArrange Then
                    For Each oDim As DrawingDimension In targetDims
                        Try
                            If CenterDimension(oDim) Then nCenter += 1
                        Catch
                            nFail += 1
                        End Try
                    Next
                End If

                '=====================================================
                ' 5. ARRANGE
                '=====================================================
                If form.ArrangeDim AndAlso Not form.CenterDimArrange Then
                    nArrange = ArrangeDimensions(invApp, oSheet, targetDims)
                End If

                '=====================================================
                ' 6. CENTER + ARRANGE
                '=====================================================
                If form.CenterDimArrange Then
                    For Each oDim As DrawingDimension In targetDims
                        Try
                            If CenterDimension(oDim) Then nCenter += 1
                        Catch
                            nFail += 1
                        End Try
                    Next
                    nArrange = ArrangeDimensions(invApp, oSheet, targetDims)
                End If

                oDrawDoc.Update()

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Phạm vi: " & If(form.AllViews, "Tất cả view", "View đã chọn") & vbCrLf &
                    "Số view xử lý: " & selectedViews.Count & vbCrLf & vbCrLf &
                    "Xóa Dimension lỗ: " & nHoleDim & vbCrLf &
                    "Xóa Hole/Thread Note: " & nHoleNote & vbCrLf &
                    "Xóa tất cả Dim: " & nAllDim & vbCrLf &
                    "Căn dim về giữa: " & nCenter & vbCrLf &
                    "Arrange dim: " & nArrange & vbCrLf &
                    "Lỗi / bỏ qua: " & nFail,
                    "Dim Cleanup", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Dim Cleanup",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '=====================================================
        ' CÁC HÀM BÊN DƯỚI GIỮ NGUYÊN
        '=====================================================
        Private Function GetDimensionsForViews(oSheet As Sheet,
                                               views As List(Of DrawingView)) As List(Of DrawingDimension)
            Dim result As New List(Of DrawingDimension)

            Try
                If views.Count = oSheet.DrawingViews.Count Then
                    For Each oDim As DrawingDimension In oSheet.DrawingDimensions
                        result.Add(oDim)
                    Next
                    Return result
                End If

                For Each oDim As DrawingDimension In oSheet.DrawingDimensions
                    Try
                        If DimensionBelongsToViews(oDim, views) Then
                            result.Add(oDim)
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            Return result
        End Function


        Private Function DimensionBelongsToViews(oDim As DrawingDimension,
                                                 views As List(Of DrawingView)) As Boolean
            Try
                Dim intents As Object = Nothing
                Try
                    intents = CallByName(oDim, "Intent", CallType.Get)
                Catch
                End Try

                If intents IsNot Nothing Then
                    Dim parentView As DrawingView = GetViewFromIntent(intents)
                    If parentView IsNot Nothing Then
                        For Each v As DrawingView In views
                            If v Is parentView Then Return True
                        Next
                        Return False
                    End If
                End If

                Try
                    Dim i1 As Object = CallByName(oDim, "IntentOne", CallType.Get)
                    Dim v1 As DrawingView = GetViewFromIntent(i1)
                    If v1 IsNot Nothing Then
                        For Each v As DrawingView In views
                            If v Is v1 Then Return True
                        Next
                    End If
                Catch
                End Try

                Try
                    Dim i2 As Object = CallByName(oDim, "IntentTwo", CallType.Get)
                    Dim v2 As DrawingView = GetViewFromIntent(i2)
                    If v2 IsNot Nothing Then
                        For Each v As DrawingView In views
                            If v Is v2 Then Return True
                        Next
                    End If
                Catch
                End Try

                Return False
            Catch
                Return False
            End Try
        End Function


        Private Function GetViewFromIntent(intentObj As Object) As DrawingView
            Try
                If intentObj Is Nothing Then Return Nothing

                Dim geom As Object = Nothing
                Try
                    geom = CallByName(intentObj, "Geometry", CallType.Get)
                Catch
                End Try
                If geom Is Nothing Then Return Nothing

                Try
                    Dim parent As Object = CallByName(geom, "Parent", CallType.Get)
                    If TypeOf parent Is DrawingView Then
                        Return CType(parent, DrawingView)
                    End If
                Catch
                End Try
            Catch
            End Try
            Return Nothing
        End Function


        Private Function CenterDimension(oDim As DrawingDimension) As Boolean
            Try
                If TypeOf oDim Is LinearGeneralDimension OrElse
                   TypeOf oDim Is AngularGeneralDimension Then
                    oDim.CenterText()
                    Return True
                End If
                Return False
            Catch
                Return False
            End Try
        End Function


        Private Function ArrangeDimensions(invApp As Inventor.Application,
                                           oSheet As Sheet,
                                           dims As List(Of DrawingDimension)) As Integer
            Dim count As Integer = 0
            Try
                Dim col As ObjectCollection = invApp.TransientObjects.CreateObjectCollection()

                For Each oDim As DrawingDimension In dims
                    Try
                        If TypeOf oDim Is LinearGeneralDimension OrElse
                           TypeOf oDim Is AngularGeneralDimension OrElse
                           TypeOf oDim Is DiameterGeneralDimension OrElse
                           TypeOf oDim Is RadiusGeneralDimension Then
                            Try
                                oDim.CenterText()
                            Catch
                            End Try
                            col.Add(oDim)
                        End If
                    Catch
                    End Try
                Next

                If col.Count > 0 Then
                    oSheet.DrawingDimensions.Arrange(col)
                    count = col.Count
                End If
            Catch
            End Try
            Return count
        End Function


        Private Function GetViewFromHoleThreadNote(htNote As HoleThreadNote) As DrawingView
            Try
                Dim linkedCurve As DrawingCurve = Nothing
                Try
                    linkedCurve = htNote.Edge
                Catch
                End Try

                If linkedCurve IsNot Nothing Then
                    Try
                        Dim p As Object = linkedCurve.Parent
                        If TypeOf p Is DrawingView Then
                            Return CType(p, DrawingView)
                        End If
                    Catch
                    End Try
                End If

                Try
                    Dim intent As GeometryIntent = htNote.Intent
                    If intent IsNot Nothing AndAlso intent.Geometry IsNot Nothing Then
                        Dim geom As Object = intent.Geometry
                        If TypeOf geom Is DrawingCurve Then
                            Dim p As Object = CType(geom, DrawingCurve).Parent
                            If TypeOf p Is DrawingView Then
                                Return CType(p, DrawingView)
                            End If
                        End If
                    End If
                Catch
                End Try

            Catch
            End Try
            Return Nothing
        End Function

    End Module

End Namespace