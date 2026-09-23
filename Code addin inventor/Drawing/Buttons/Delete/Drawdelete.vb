Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor

Namespace ToolInventor2025.Drawing.Buttons.Drawdelete

    '=====================================================
    ' FORM CHỌN CHỨC NĂNG XÓA / ẨN
    '=====================================================
    Public Class CleanupFormDelete
        Inherits Form

        Private chkHideLabel As CheckBox
        Private chkDeleteBalloon As CheckBox
        Private chkDeleteSurface As CheckBox
        ' Private chkDeleteDatum As CheckBox
        Private chkDeleteFCF As CheckBox
        Private chkDeleteHoleDim As CheckBox
        Private chkDeleteHoleNote As CheckBox
        Private chkDeleteTextNote As CheckBox
        Private chkDeleteLeaderText As CheckBox
        Private chkDeleteWelding As CheckBox
        Private btnOK As Button
        Private btnCancel As Button
        Private chkDeleteSketchSymbol As CheckBox
        Public ReadOnly Property HideViewLabel As Boolean
        Public ReadOnly Property DeleteBalloon As Boolean
        Public ReadOnly Property DeleteSurface As Boolean
        '  Public ReadOnly Property DeleteDatum As Boolean
        Public ReadOnly Property DeleteFCF As Boolean
        Public ReadOnly Property DeleteHoleDim As Boolean
        Public ReadOnly Property DeleteHoleNote As Boolean
        Public ReadOnly Property DeleteTextNote As Boolean
        Public ReadOnly Property DeleteLeaderText As Boolean
        Public ReadOnly Property DeleteWelding As Boolean
        Public ReadOnly Property Cancelled As Boolean
        Public ReadOnly Property DeleteSketchSymbol As Boolean
        Public Sub New()
            _Cancelled = False

            Me.Text = "Dọn dẹp bản vẽ"
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ClientSize = New Size(380, 370)

            Dim lbl As New Label()
            lbl.Text = "Chọn thao tác cần thực hiện:"
            lbl.Location = New System.Drawing.Point(15, 12)
            lbl.Size = New Size(360, 20)
            lbl.Font = New Font(lbl.Font, FontStyle.Bold)
            Me.Controls.Add(lbl)

            chkDeleteHoleDim = MakeCheckBox("Xóa Dimension lỗ (Diameter)", 40)
            chkDeleteHoleNote = MakeCheckBox("Xóa Hole / Thread Note", 68)

            chkHideLabel = MakeCheckBox("Ẩn Label của tất cả Drawing View", 96)
            chkDeleteBalloon = MakeCheckBox("Xóa Balloon (bong bóng đánh số)", 124)
            chkDeleteSurface = MakeCheckBox("Xóa Surface Texture Symbol", 152)
            ' chkDeleteDatum = MakeCheckBox("Xóa Datum Target Symbol", 180)
            chkDeleteSketchSymbol = MakeCheckBox("Xóa Sketch Symbol (ký hiệu sketch)", 180)
            chkDeleteFCF = MakeCheckBox("Xóa Feature Control Frame (GD&T)", 208)
            chkDeleteTextNote = MakeCheckBox("Xóa Text Note (không leader)", 236)
            chkDeleteLeaderText = MakeCheckBox("Xóa Leader Text (có leader)", 264)
            chkDeleteWelding = MakeCheckBox("Xóa Welding Symbol (cần 2024+)", 292)


            btnOK = New Button()
            btnOK.Text = "Thực hiện"
            btnOK.Location = New System.Drawing.Point(170, 325)
            btnOK.Size = New Size(85, 30)
            btnOK.DialogResult = DialogResult.OK
            Me.Controls.Add(btnOK)

            btnCancel = New Button()
            btnCancel.Text = "Hủy"
            btnCancel.Location = New System.Drawing.Point(265, 325)
            btnCancel.Size = New Size(85, 30)
            btnCancel.DialogResult = DialogResult.Cancel
            Me.Controls.Add(btnCancel)

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub

        Private Function MakeCheckBox(ByVal text As String, ByVal top As Integer) As CheckBox
            Dim chk As New CheckBox()
            chk.Text = text
            chk.Location = New System.Drawing.Point(20, top)
            chk.Size = New Size(360, 22)
            Me.Controls.Add(chk)
            Return chk
        End Function

        Public Function ShowAndGet() As Boolean
            Dim result As DialogResult = Me.ShowDialog()
            If result <> DialogResult.OK Then
                _Cancelled = True
                Return False
            End If

            _HideViewLabel = chkHideLabel.Checked
            _DeleteBalloon = chkDeleteBalloon.Checked
            _DeleteSurface = chkDeleteSurface.Checked
            '_DeleteDatum = chkDeleteDatum.Checked
            _DeleteSketchSymbol = chkDeleteSketchSymbol.Checked
            _DeleteFCF = chkDeleteFCF.Checked
            _DeleteHoleDim = chkDeleteHoleDim.Checked
            _DeleteHoleNote = chkDeleteHoleNote.Checked
            _DeleteTextNote = chkDeleteTextNote.Checked
            _DeleteLeaderText = chkDeleteLeaderText.Checked
            _DeleteWelding = chkDeleteWelding.Checked
            Return True
        End Function

        Private Sub InitializeComponent()
            Me.SuspendLayout()
            '
            'CleanupFormDelete
            '
            Me.ClientSize = New System.Drawing.Size(282, 253)
            Me.Name = "CleanupFormDelete"
            Me.ResumeLayout(False)

        End Sub

        Private Sub CleanupFormDelete_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        End Sub
    End Class

    '=====================================================
    ' MODULE CHÍNH
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
                ' Dim nDatum As Integer = 0
                Dim nSketchSymbol As Integer = 0
                ' (xóa Dim nDatum)
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
                ' 4. DATUM TARGET
                '=====================================================
                '  If form.DeleteDatum Then
                ' Try
                ' Dim toDel As New List(Of DatumTarget)
                'For Each d As DatumTarget In oSheet.DatumTargets.ToString
                ' toDel.Add(d)
                'Next
                '  For Each d As DatumTarget In toDel
                '   Try
                'd.Delete()
                '        nDatum += 1
                'Catch
                '           nFail += 1
                'End Try
                '   Next
                'Catch
                'End Try
                '    End If
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
                ' 7. XÓA DIAMETER DIMENSION CỦA LỖ
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
                ' 8. XÓA HOLE / THREAD NOTE
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
                            ' Bỏ qua HoleThreadNote (đã xử lý ở bước 8)
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
                  "Xóa Sketch Symbol: " & nSketchSymbol & vbCrLf &                  '  "Xóa Datum: " & nDatum & vbCrLf &
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