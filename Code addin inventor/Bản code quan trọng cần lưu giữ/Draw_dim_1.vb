Option Explicit On
Option Strict Off
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports Inventor

Namespace ToolInventor2025.Drawing.Buttons.Drawdim

    Friend Class NativeMethodsb
        <DllImport("user32.dll")>
        Public Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
        End Function
    End Class

    '=====================================================
    ' FORM
    '=====================================================
    Public Class DimCleanupFormb
        Inherits Form

        Private chkDeleteHoleDim As CheckBox
        Private chkDeleteHoleNote As CheckBox
        Private chkDeleteAllDim As CheckBox
        Private chkCenterDimArrange As CheckBox
        Private chkArrangeDim As CheckBox
        Private chkCenterDim As CheckBox
        Private rdoAllViews As RadioButton
        Private rdoPickViews As RadioButton
        Private btnOK As Button
        Private btnCancel As Button

        Public ReadOnly Property DeleteHoleDim As Boolean
        Public ReadOnly Property DeleteHoleNote As Boolean
        Public ReadOnly Property DeleteAllDim As Boolean
        Public ReadOnly Property CenterDimArrange As Boolean
        Public ReadOnly Property ArrangeDim As Boolean
        Public ReadOnly Property CenterDim As Boolean
        Public ReadOnly Property AllViews As Boolean
        Public ReadOnly Property PickViews As Boolean
        Public ReadOnly Property Cancelled As Boolean

        Public Sub New()
            _Cancelled = False
            Me.Text = "Xử lý Dimension bản vẽ"
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ClientSize = New Size(420, 440)

            Dim lbl As New Label()
            lbl.Text = "Chọn thao tác cần thực hiện:"
            lbl.Location = New System.Drawing.Point(15, 12)
            lbl.Size = New Size(380, 20)
            lbl.Font = New Font(lbl.Font, FontStyle.Bold)
            Me.Controls.Add(lbl)

            Dim lblScope As New Label()
            lblScope.Text = "— Phạm vi —"
            lblScope.Location = New System.Drawing.Point(15, 38)
            lblScope.Size = New Size(380, 18)
            lblScope.ForeColor = System.Drawing.Color.DarkGreen
            Me.Controls.Add(lblScope)

            rdoAllViews = New RadioButton()
            rdoAllViews.Text = "Tất cả View trên sheet"
            rdoAllViews.Location = New System.Drawing.Point(25, 58)
            rdoAllViews.Size = New Size(200, 22)
            rdoAllViews.Checked = True
            Me.Controls.Add(rdoAllViews)

            rdoPickViews = New RadioButton()
            rdoPickViews.Text = "Chọn View riêng lẻ (multi-select)"
            rdoPickViews.Location = New System.Drawing.Point(25, 82)
            rdoPickViews.Size = New Size(280, 22)
            Me.Controls.Add(rdoPickViews)

            Dim lbl1 As New Label()
            lbl1.Text = "— Xóa —"
            lbl1.Location = New System.Drawing.Point(15, 116)
            lbl1.Size = New Size(380, 18)
            lbl1.ForeColor = System.Drawing.Color.DarkRed
            Me.Controls.Add(lbl1)

            chkDeleteHoleDim = MakeCheckBox("Xóa Dimension lỗ (Diameter)", 136)
            chkDeleteHoleNote = MakeCheckBox("Xóa Hole / Thread Note", 162)
            chkDeleteAllDim = MakeCheckBox("Xóa tất cả Dimension", 188)

            Dim lbl2 As New Label()
            lbl2.Text = "— Căn chỉnh —"
            lbl2.Location = New System.Drawing.Point(15, 222)
            lbl2.Size = New Size(380, 18)
            lbl2.ForeColor = System.Drawing.Color.DarkBlue
            Me.Controls.Add(lbl2)

            chkCenterDim = MakeCheckBox("Căn dim về giữa", 242)
            chkArrangeDim = MakeCheckBox("Arrange dim (tự sắp xếp)", 268)
            chkCenterDimArrange = MakeCheckBox("Căn dim về giữa + Arrange", 294)

            btnOK = New Button()
            btnOK.Text = "Thực hiện"
            btnOK.Location = New System.Drawing.Point(220, 380)
            btnOK.Size = New Size(85, 30)
            btnOK.DialogResult = DialogResult.OK
            Me.Controls.Add(btnOK)

            btnCancel = New Button()
            btnCancel.Text = "Hủy"
            btnCancel.Location = New System.Drawing.Point(315, 380)
            btnCancel.Size = New Size(85, 30)
            btnCancel.DialogResult = DialogResult.Cancel
            Me.Controls.Add(btnCancel)

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub

        Private Function MakeCheckBox(ByVal text As String, ByVal top As Integer) As CheckBox
            Dim chk As New CheckBox()
            chk.Text = text
            chk.Location = New System.Drawing.Point(25, top)
            chk.Size = New Size(380, 22)
            Me.Controls.Add(chk)
            Return chk
        End Function

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
    ' MODULE
    '=====================================================
    Public Module Dim_delete2

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

                Dim form As New DimCleanupFormb()
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
                            NativeMethods.SetForegroundWindow(mainHwnd)
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

                ' Lấy toàn bộ dim trên sheet thuộc các view đã chọn
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
                ' 2. XÓA HOLE / THREAD NOTE  (lọc theo view đã chọn)
                '=====================================================
                If form.DeleteHoleNote Then
                    Try
                        Dim toDel As New List(Of HoleThreadNote)

                        For Each htNote As HoleThreadNote In oSheet.DrawingNotes.HoleThreadNotes
                            Try
                                ' Nếu chọn tất cả view → xóa hết
                                If form.AllViews Then
                                    toDel.Add(htNote)
                                    Continue For
                                End If

                                ' Lọc theo view: lấy curve / intent gắn với note
                                Dim noteView As DrawingView = GetViewFromHoleThreadNote(htNote)
                                If noteView Is Nothing Then
                                    ' Không xác định được view → bỏ qua khi đang PickViews
                                    Continue For
                                End If

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
                    ' Linear / Diameter / Angular... thuộc view
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

                    ' HoleThreadNotes
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

                    ' Baseline sets
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

                    ' Chain sets
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

                ' Refresh target dims sau khi xóa (cho Center/Arrange)
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
        ' LẤY DIM THUỘC CÁC VIEW ĐÃ CHỌN
        ' (Dim nằm trên Sheet, không phải trên View)
        '=====================================================
        Private Function GetDimensionsForViews(
            oSheet As Sheet,
            views As List(Of DrawingView)) As List(Of DrawingDimension)

            Dim result As New List(Of DrawingDimension)

            Try
                ' Nếu chọn tất cả view trên sheet → lấy hết
                If views.Count = oSheet.DrawingViews.Count Then
                    For Each oDim As DrawingDimension In oSheet.DrawingDimensions
                        result.Add(oDim)
                    Next
                    Return result
                End If

                ' Lọc dim theo view
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

        Private Function DimensionBelongsToViews(
            oDim As DrawingDimension,
            views As List(Of DrawingView)) As Boolean

            Try
                ' Thử lấy Intent / Geometry → Parent View
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

                ' Linear dim thường có IntentOne / IntentTwo
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

                ' Không xác định được → nếu đang AllViews đã xử lý ở trên
                ' với PickViews: bỏ qua dim không gắn view rõ
                Return False
            Catch
                Return False
            End Try
        End Function
        '=====================================================
        ' Lấy DrawingView gắn với HoleThreadNote
        '=====================================================
        Private Function GetViewFromHoleThreadNote(htNote As HoleThreadNote) As DrawingView
            Try
                ' Cách 1: Edge (DrawingCurve)
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

                ' Cách 2: Intent.Geometry
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
        Private Function GetViewFromIntent(intentObj As Object) As DrawingView
            Try
                If intentObj Is Nothing Then Return Nothing

                Dim geom As Object = Nothing
                Try
                    geom = CallByName(intentObj, "Geometry", CallType.Get)
                Catch
                End Try
                If geom Is Nothing Then Return Nothing

                ' DrawingCurve.Parent = DrawingView
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

        Private Function ArrangeDimensions(
            invApp As Inventor.Application,
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

    End Module
End Namespace