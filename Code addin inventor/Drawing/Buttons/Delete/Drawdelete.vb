Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.Drawdelete

    '=====================================================
    ' FORM CHỌN CHỨC NĂNG XÓA / ẨN — PHÂN NHÓM
    '=====================================================
    Public Class CleanupFormDelete
        Inherits Form

        '===== PROPERTIES =====
        Public ReadOnly Property HideViewLabel As Boolean
        Public ReadOnly Property DeleteBalloon As Boolean
        Public ReadOnly Property DeleteSurface As Boolean
        Public ReadOnly Property DeleteFCF As Boolean
        Public ReadOnly Property DeleteHoleDim As Boolean
        Public ReadOnly Property DeleteHoleNote As Boolean
        Public ReadOnly Property DeleteTextNote As Boolean
        Public ReadOnly Property DeleteLeaderText As Boolean
        Public ReadOnly Property DeleteWelding As Boolean
        Public ReadOnly Property Cancelled As Boolean
        Public ReadOnly Property DeleteSketchSymbol As Boolean

        '===== CONTROLS =====
        Private chkHideLabel As CheckBox
        Private chkDeleteBalloon As CheckBox
        Private chkDeleteSurface As CheckBox
        Private chkDeleteFCF As CheckBox
        Private chkDeleteHoleDim As CheckBox
        Private chkDeleteHoleNote As CheckBox
        Private chkDeleteTextNote As CheckBox
        Private chkDeleteLeaderText As CheckBox
        Private chkDeleteWelding As CheckBox
        Private chkDeleteSketchSymbol As CheckBox
        Private chkSelectAll As CheckBox
        Private btnOK As Button
        Private btnCancel As Button

        Public Sub New()
            _Cancelled = False

            Me.Text = "Dọn dẹp bản vẽ"
            Me.AutoScaleMode = AutoScaleMode.None
            Me.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
            Me.FormBorderStyle = FormBorderStyle.FixedSingle
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.ClientSize = New Drw.Size(660, 660)
            Me.BackColor = Drw.Color.FromArgb(245, 245, 245)
            Me.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

            '===== HEADER =====
            Dim pnlHeader As New Panel()
            pnlHeader.Location = New Drw.Point(0, 0)
            pnlHeader.Size = New Drw.Size(660, 75)
            pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
            Me.Controls.Add(pnlHeader)

            Dim lblTitle As New Label()
            lblTitle.Text = "DỌN DẸP BẢN VẼ DRAWING"
            lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblTitle.ForeColor = Drw.Color.White
            lblTitle.Dock = DockStyle.Fill
            lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
            pnlHeader.Controls.Add(lblTitle)

            Dim lblSub As New Label()
            lblSub.Text = "Chọn các thao tác cần thực hiện"
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

            '===== GROUP 1: CHI TIẾT KỸ THUẬT =====
            Dim gb1 As New GroupBox() With {
                .Text = "Chi tiết kỹ thuật",
                .Location = New Drw.Point(15, 120),
                .Size = New Drw.Size(630, 130),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gb1)

            chkDeleteHoleDim = MakeCheckBox(gb1, "Xóa Dimension lỗ (Diameter)", 30)
            chkDeleteHoleNote = MakeCheckBox(gb1, "Xóa Hole / Thread Note", 60)
            chkDeleteFCF = MakeCheckBox(gb1, "Xóa Feature Control Frame (GD&T)", 90)

            '===== GROUP 2: KÝ HIỆU =====
            Dim gb2 As New GroupBox() With {
                .Text = "Ký hiệu / Symbol",
                .Location = New Drw.Point(15, 260),
                .Size = New Drw.Size(630, 195),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gb2)

            chkDeleteBalloon = MakeCheckBox(gb2, "Xóa Balloon (bong bóng đánh số)", 30)
            chkDeleteSurface = MakeCheckBox(gb2, "Xóa Surface Texture Symbol", 60)
            chkDeleteSketchSymbol = MakeCheckBox(gb2, "Xóa Sketch Symbol (ký hiệu sketch)", 90)
            chkDeleteWelding = MakeCheckBox(gb2, "Xóa Welding Symbol (cần Inventor 2024+)", 120)

            '===== GROUP 3: TEXT & VIEW =====
            Dim gb3 As New GroupBox() With {
                .Text = "Text & View",
                .Location = New Drw.Point(15, 465),
                .Size = New Drw.Size(630, 130),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gb3)

            chkDeleteTextNote = MakeCheckBox(gb3, "Xóa Text Note (không leader)", 30)
            chkDeleteLeaderText = MakeCheckBox(gb3, "Xóa Leader Text (có leader)", 60)
            chkHideLabel = MakeCheckBox(gb3, "Ẩn Label của tất cả Drawing View", 90)

            '===== NÚT =====
            btnOK = New Button()
            btnOK.Text = "THỰC HIỆN"
            btnOK.Location = New Drw.Point(495, 605)
            btnOK.Size = New Drw.Size(150, 42)
            btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
            btnOK.ForeColor = Drw.Color.White
            btnOK.FlatStyle = FlatStyle.Flat
            btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            btnOK.DialogResult = DialogResult.OK
            Me.Controls.Add(btnOK)

            btnCancel = New Button()
            btnCancel.Text = "HỦY"
            btnCancel.Location = New Drw.Point(355, 605)
            btnCancel.Size = New Drw.Size(130, 42)
            btnCancel.FlatStyle = FlatStyle.Flat
            btnCancel.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            btnCancel.DialogResult = DialogResult.Cancel
            Me.Controls.Add(btnCancel)

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub


        '=========================================================
        ' TẠO CHECKBOX
        '=========================================================
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


        '=========================================================
        ' CHỌN / BỎ CHỌN TẤT CẢ
        '=========================================================
        Private Sub OnSelectAllChanged(ByVal sender As Object, ByVal e As EventArgs)
            Dim chk As Boolean = chkSelectAll.Checked

            chkDeleteHoleDim.Checked = chk
            chkDeleteHoleNote.Checked = chk
            chkDeleteFCF.Checked = chk
            chkDeleteBalloon.Checked = chk
            chkDeleteSurface.Checked = chk
            chkDeleteSketchSymbol.Checked = chk
            chkDeleteWelding.Checked = chk
            chkDeleteTextNote.Checked = chk
            chkDeleteLeaderText.Checked = chk
            chkHideLabel.Checked = chk
        End Sub


        '=========================================================
        ' SHOW VÀ LẤY KẾT QUẢ
        '=========================================================
        Public Function ShowAndGet() As Boolean
            Dim result As DialogResult = Me.ShowDialog()
            If result <> DialogResult.OK Then
                _Cancelled = True
                Return False
            End If

            _HideViewLabel = chkHideLabel.Checked
            _DeleteBalloon = chkDeleteBalloon.Checked
            _DeleteSurface = chkDeleteSurface.Checked
            _DeleteSketchSymbol = chkDeleteSketchSymbol.Checked
            _DeleteFCF = chkDeleteFCF.Checked
            _DeleteHoleDim = chkDeleteHoleDim.Checked
            _DeleteHoleNote = chkDeleteHoleNote.Checked
            _DeleteTextNote = chkDeleteTextNote.Checked
            _DeleteLeaderText = chkDeleteLeaderText.Checked
            _DeleteWelding = chkDeleteWelding.Checked
            Return True
        End Function

    End Class


    '=====================================================
    ' MODULE CHÍNH — GIỮ NGUYÊN 100%
    '=====================================================
    Public Module Draw_delete

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                Dim invApp As Inventor.Application = g_inventorApplication
                If invApp Is Nothing Then
                    MessageBox.Show("Không tìm thấy Inventor Application.", "Cleanup",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If invApp.ActiveDocument Is Nothing OrElse
                   invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Chức năng này chỉ dùng cho bản vẽ Drawing.", "Cleanup",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument = CType(invApp.ActiveDocument, DrawingDocument)
                Dim oSheet As Sheet = oDrawDoc.ActiveSheet

                Dim form As New CleanupFormDelete()
                If Not form.ShowAndGet() Then Exit Sub

                Dim nLabel As Integer = 0
                Dim nBalloon As Integer = 0
                Dim nSurface As Integer = 0
                Dim nSketchSymbol As Integer = 0
                Dim nFCF As Integer = 0
                Dim nHoleDim As Integer = 0
                Dim nHoleNote As Integer = 0
                Dim nTextNote As Integer = 0
                Dim nLeader As Integer = 0
                Dim nWelding As Integer = 0
                Dim nFail As Integer = 0

                '=====================================================
                ' 1. ẨN LABEL VIEW
                '=====================================================
                If form.HideViewLabel Then
                    For Each dv As DrawingView In oSheet.DrawingViews
                        Try
                            dv.ShowLabel = False
                            nLabel += 1
                        Catch
                            nFail += 1
                        End Try
                    Next
                End If

                '=====================================================
                ' 2. BALLOON
                '=====================================================
                If form.DeleteBalloon Then
                    Try
                        Dim toDel As New List(Of Balloon)
                        For Each b As Balloon In oSheet.Balloons
                            toDel.Add(b)
                        Next
                        For Each b As Balloon In toDel
                            Try
                                b.Delete()
                                nBalloon += 1
                            Catch
                                nFail += 1
                            End Try
                        Next
                    Catch
                    End Try
                End If

                '=====================================================
                ' 3. SURFACE TEXTURE SYMBOL
                '=====================================================
                If form.DeleteSurface Then
                    Try
                        Dim toDel As New List(Of SurfaceTextureSymbol)
                        For Each s As SurfaceTextureSymbol In oSheet.SurfaceTextureSymbols
                            toDel.Add(s)
                        Next
                        For Each s As SurfaceTextureSymbol In toDel
                            Try
                                s.Delete()
                                nSurface += 1
                            Catch
                                nFail += 1
                            End Try
                        Next
                    Catch
                    End Try
                End If

                '=====================================================
                ' 4. SKETCH SYMBOL
                '=====================================================
                If form.DeleteSketchSymbol Then
                    Try
                        Dim toDel As New List(Of SketchedSymbol)
                        For Each sk As SketchedSymbol In oSheet.SketchedSymbols
                            toDel.Add(sk)
                        Next
                        For Each sk As SketchedSymbol In toDel
                            Try
                                sk.Delete()
                                nSketchSymbol += 1
                            Catch
                                nFail += 1
                            End Try
                        Next
                    Catch
                        nFail += 1
                    End Try
                End If

                '=====================================================
                ' 5. FEATURE CONTROL FRAME
                '=====================================================
                If form.DeleteFCF Then
                    Try
                        Dim toDel As New List(Of FeatureControlFrame)
                        For Each f As FeatureControlFrame In oSheet.FeatureControlFrames
                            toDel.Add(f)
                        Next
                        For Each f As FeatureControlFrame In toDel
                            Try
                                f.Delete()
                                nFCF += 1
                            Catch
                                nFail += 1
                            End Try
                        Next
                    Catch
                    End Try
                End If

                '=====================================================
                ' 6. WELDING SYMBOL
                '=====================================================
                If form.DeleteWelding Then
                    Dim weldOk As Boolean = False
                    Try
                        Dim weldCol As Object = oSheet.WeldingSymbols
                        If weldCol IsNot Nothing Then
                            Dim cnt As Integer = CInt(weldCol.Count)
                            For i As Integer = cnt To 1 Step -1
                                Try
                                    Dim ws As Object = weldCol.Item(i)
                                    ws.Delete()
                                    nWelding += 1
                                Catch
                                    nFail += 1
                                End Try
                            Next
                            weldOk = True
                        End If
                    Catch
                        weldOk = False
                    End Try

                    If Not weldOk Then
                        MessageBox.Show(
                            "Xóa Welding Symbol bằng API chỉ hỗ trợ từ Inventor 2024." & vbCrLf & vbCrLf &
                            "Với Inventor 2020 hãy xóa thủ công bằng cách:" & vbCrLf &
                            "Shift + Right-click → chọn Welding Symbol → quét chọn → Delete.",
                            "Cleanup - Welding", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End If

                '=====================================================
                ' 7. DIAMETER DIMENSION
                '=====================================================
                If form.DeleteHoleDim Then
                    Try
                        Dim toDel As New List(Of DrawingDimension)
                        For Each oDim As DrawingDimension In oSheet.DrawingDimensions
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
                    Catch
                    End Try
                End If

                '=====================================================
                ' 8. HOLE / THREAD NOTE
                '=====================================================
                If form.DeleteHoleNote Then
                    Try
                        Dim toDel As New List(Of HoleThreadNote)
                        For Each htNote As HoleThreadNote In oSheet.DrawingNotes.HoleThreadNotes
                            toDel.Add(htNote)
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
                ' 9. TEXT NOTE & LEADER TEXT
                '=====================================================
                If form.DeleteTextNote OrElse form.DeleteLeaderText Then
                    Dim toDelNotes As New List(Of DrawingNote)
                    For Each oNote As DrawingNote In oSheet.DrawingNotes
                        Try
                            If TypeOf oNote Is HoleThreadNote Then Continue For

                            Dim hasLeader As Boolean = False
                            Try
                                If oNote.Leader IsNot Nothing Then hasLeader = True
                            Catch
                            End Try

                            If hasLeader Then
                                If form.DeleteLeaderText Then
                                    toDelNotes.Add(oNote)
                                    nLeader += 1
                                End If
                            Else
                                If form.DeleteTextNote Then
                                    toDelNotes.Add(oNote)
                                    nTextNote += 1
                                End If
                            End If
                        Catch
                            nFail += 1
                        End Try
                    Next

                    For Each oNote As DrawingNote In toDelNotes
                        Try
                            oNote.Delete()
                        Catch
                            Try
                                oSheet.DrawingNotes.Remove(oNote)
                            Catch
                                nFail += 1
                            End Try
                        End Try
                    Next
                End If

                oDrawDoc.Update()

                '===== THÔNG BÁO =====
                Dim msg As String =
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Ẩn Label view: " & nLabel & vbCrLf &
                    "Xóa Balloon: " & nBalloon & vbCrLf &
                    "Xóa Surface: " & nSurface & vbCrLf &
                    "Xóa Sketch Symbol: " & nSketchSymbol & vbCrLf &
                    "Xóa Feature Control Frame: " & nFCF & vbCrLf &
                    "Xóa Welding: " & nWelding & vbCrLf &
                    "Xóa Dimension lỗ: " & nHoleDim & vbCrLf &
                    "Xóa Hole/Thread Note: " & nHoleNote & vbCrLf &
                    "Xóa Text Note: " & nTextNote & vbCrLf &
                    "Xóa Leader Text: " & nLeader & vbCrLf &
                    "Lỗi / bỏ qua: " & nFail

                MessageBox.Show(msg, "Cleanup",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Cleanup",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Module

End Namespace