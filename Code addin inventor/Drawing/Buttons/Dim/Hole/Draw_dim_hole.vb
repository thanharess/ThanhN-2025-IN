Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor

Namespace ToolInventor2025.Drawing.Buttons.Drawdim

    '=====================================================
    ' CLASS CHỨA TÙY CHỌN NGƯỜI DÙNG
    '=====================================================
    Public Class DrawOptions
        Public Property SkipThreadDim As Boolean = False
        Public Property SkipNormalDim As Boolean = False
        Public Property Cancelled As Boolean = False
    End Class

    '=====================================================
    ' FORM CHỌN TÙY CHỌN
    '=====================================================
    Public Class DrawOptionsForm
        Inherits Form

        Private chkSkipThread As CheckBox
        Private chkSkipNormal As CheckBox
        Private btnOK As Button
        Private btnCancel As Button

        Public ReadOnly Property Options As DrawOptions

        Public Sub New()
            _Options = New DrawOptions()

            Me.Text = "Tùy chọn Dim lỗ ( Lỗ ren sẽ bị trùng nếu đã có dim )"
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ClientSize = New Size(360, 170)

            Dim lbl As New Label()
            lbl.Text = "Chọn loại lỗ bỏ qua khi dim:"
            lbl.Location = New System.Drawing.Point(15, 12)
            lbl.Size = New Size(320, 20)
            lbl.Font = New Font(lbl.Font, FontStyle.Bold)
            Me.Controls.Add(lbl)

            chkSkipThread = New CheckBox()
            chkSkipThread.Text = "Bỏ qua Lỗ ren (Hole/Thread Note)"
            chkSkipThread.Location = New System.Drawing.Point(20, 45)
            chkSkipThread.Size = New Size(320, 22)
            Me.Controls.Add(chkSkipThread)

            chkSkipNormal = New CheckBox()
            chkSkipNormal.Text = "Bỏ qua Lỗ thường (Diameter)"
            chkSkipNormal.Location = New System.Drawing.Point(20, 75)
            chkSkipNormal.Size = New Size(320, 22)
            Me.Controls.Add(chkSkipNormal)

            btnOK = New Button()
            btnOK.Text = "Bắt đầu"
            btnOK.Location = New System.Drawing.Point(170, 115)
            btnOK.Size = New Size(85, 30)
            btnOK.DialogResult = DialogResult.OK
            Me.Controls.Add(btnOK)

            btnCancel = New Button()
            btnCancel.Text = "Hủy"
            btnCancel.Location = New System.Drawing.Point(260, 115)
            btnCancel.Size = New Size(85, 30)
            btnCancel.DialogResult = DialogResult.Cancel
            Me.Controls.Add(btnCancel)

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub

        Public Function ShowAndGet() As DrawOptions
            Dim result As DialogResult = Me.ShowDialog()
            If result <> DialogResult.OK Then
                _Options.Cancelled = True
            Else
                _Options.SkipThreadDim = chkSkipThread.Checked
                _Options.SkipNormalDim = chkSkipNormal.Checked
            End If
            Return _Options
        End Function
    End Class

    '=====================================================
    ' MODULE CHÍNH
    '=====================================================
    Public Module Draw_dim_hole

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then

                    MessageBox.Show("Vui lòng mở file Drawing (.idw)!", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                ' ===== HIỆN FORM CHỌN TÙY CHỌN =====
                Dim optForm As New DrawOptionsForm()
                Dim opts As DrawOptions = optForm.ShowAndGet()
                If opts.Cancelled Then Exit Sub
                ' ===================================

                Dim oDrawDoc As Inventor.DrawingDocument =
                    CType(app.ActiveDocument, Inventor.DrawingDocument)

                Dim oSheet As Inventor.Sheet = oDrawDoc.ActiveSheet
                Dim tg As Inventor.TransientGeometry = app.TransientGeometry

                '=====================================================
                ' CHỌN NHIỀU VIEW
                '=====================================================
                Dim selectedViews As New List(Of DrawingView)

                Do
                    Dim oSS As SelectSet = oDrawDoc.SelectSet
                    oSS.Clear()

                    Dim oView As DrawingView = Nothing
                    Try
                        oView = CType(
                            app.CommandManager.Pick(
                                SelectionFilterEnum.kDrawingViewFilter,
                                "Chọn View (Esc hoặc Right-click để kết thúc)"),
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

                If selectedViews.Count = 0 Then
                    MessageBox.Show("Chưa chọn View nào.", "Thông báo")
                    Exit Sub
                End If

                Dim countDia As Integer = 0
                Dim countThread As Integer = 0
                Dim countFail As Integer = 0
                Dim countSkippedArray As Integer = 0
                Dim countSkippedDimmed As Integer = 0
                Dim countSkippedByOption As Integer = 0

                '=====================================================
                ' THU THẬP TẤT CẢ LỖ ĐÃ ĐƯỢC DIM
                '=====================================================
                Dim dimmedHoles As New HashSet(Of String)()
                CollectAllDimmedHoles(oSheet, dimmedHoles)

                '=====================================================
                ' XỬ LÝ TỪNG VIEW ĐÃ CHỌN
                '=====================================================
                For Each oView As Inventor.DrawingView In selectedViews

                    Try
                        If oView.DrawingCurves Is Nothing Then Continue For

                        For Each oCurve As Inventor.DrawingCurve In oView.DrawingCurves
                            Try
                                If oCurve.CurveType <> Inventor.CurveTypeEnum.kCircleCurve Then Continue For

                                '----- Bỏ qua lỗ thuộc Array -----
                                If IsPatternElement(oCurve) Then
                                    countSkippedArray += 1
                                    Continue For
                                End If

                                '----- Bỏ qua lỗ đã có dim -----
                                Dim curveKey As String = GetCenterKey(oCurve)
                                If curveKey <> "" AndAlso dimmedHoles.Contains(curveKey) Then
                                    countSkippedDimmed += 1
                                    Continue For
                                End If

                                '----- Xác định loại lỗ -----
                                Dim isThread As Boolean = IsThreadedHoleCurve(oCurve)

                                '----- Bỏ qua theo tùy chọn người dùng -----
                                If isThread AndAlso opts.SkipThreadDim Then
                                    countSkippedByOption += 1
                                    Continue For
                                End If
                                If Not isThread AndAlso opts.SkipNormalDim Then
                                    countSkippedByOption += 1
                                    Continue For
                                End If

                                '----- Tạo điểm đặt dim -----
                                Dim oIntent As Inventor.GeometryIntent =
                                    oSheet.CreateGeometryIntent(oCurve, Inventor.PointIntentEnum.kCircularLeftPointIntent)

                                Dim oPoint As Inventor.Point2d = oIntent.PointOnSheet.Copy()
                                Dim oVector As Inventor.Vector2d = oCurve.CenterPoint.VectorTo(oPoint)
                                oVector.ScaleBy(0.3)
                                oVector.AddVector(tg.CreateVector2d(oVector.X, System.Math.Abs(oVector.X)))
                                oPoint.TranslateBy(oVector)

                                Dim dimAdded As Boolean = False

                                If isThread Then
                                    ' Lỗ ren → Hole/Thread Note
                                    Try
                                        oSheet.DrawingNotes.HoleThreadNotes.Add(oPoint, oIntent)
                                        countThread += 1
                                        dimAdded = True
                                    Catch
                                        Try
                                            oSheet.DrawingDimensions.GeneralDimensions.AddDiameter(
                                                oPoint, oIntent, True, False, False)
                                            countDia += 1
                                            dimAdded = True
                                        Catch
                                            countFail += 1
                                        End Try
                                    End Try
                                Else
                                    ' Lỗ thường → Diameter
                                    Try
                                        oSheet.DrawingDimensions.GeneralDimensions.AddDiameter(
                                            oPoint, oIntent, True, False, False)
                                        countDia += 1
                                        dimAdded = True
                                    Catch
                                        countFail += 1
                                    End Try
                                End If

                                ' Đánh dấu vừa dim xong
                                If dimAdded Then
                                    Dim key As String = GetCenterKey(oCurve)
                                    If key <> "" AndAlso Not dimmedHoles.Contains(key) Then
                                        dimmedHoles.Add(key)
                                    End If
                                End If

                            Catch
                                countFail += 1
                            End Try
                        Next

                    Catch
                    End Try

                Next

                oDrawDoc.Update()

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Số view đã xử lý: " & selectedViews.Count & vbCrLf &
                    "Diameter (lỗ thường): " & countDia.ToString() & vbCrLf &
                    "Hole/Thread Note (lỗ ren): " & countThread.ToString() & vbCrLf &
                    "Bỏ qua (lỗ array): " & countSkippedArray.ToString() & vbCrLf &
                    "Bỏ qua (đã dim rồi): " & countSkippedDimmed.ToString() & vbCrLf &
                    "Bỏ qua (theo tùy chọn): " & countSkippedByOption.ToString() & vbCrLf &
                    "Lỗi / bỏ qua: " & countFail.ToString(),
                    "Dim lỗ",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Dim lỗ",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        '=================================================
        ' THU THẬP TẤT CẢ LỖ ĐÃ ĐƯỢC DIM
        '=================================================
        Private Sub CollectAllDimmedHoles(oSheet As Inventor.Sheet, ByRef dimmedHoles As HashSet(Of String))

            Try
                ' 1. Diameter dimensions
                For Each oDim As DrawingDimension In oSheet.DrawingDimensions
                    Try
                        If TypeOf oDim Is DiameterGeneralDimension Then
                            Dim diaDim As DiameterGeneralDimension = CType(oDim, DiameterGeneralDimension)
                            Try
                                Dim intent1 As GeometryIntent = diaDim.Intent
                                If intent1 IsNot Nothing AndAlso intent1.Geometry IsNot Nothing Then
                                    If TypeOf intent1.Geometry Is DrawingCurve Then
                                        Dim key As String = GetCenterKey(CType(intent1.Geometry, DrawingCurve))
                                        If key <> "" AndAlso Not dimmedHoles.Contains(key) Then
                                            dimmedHoles.Add(key)
                                        End If
                                    End If
                                End If
                            Catch
                            End Try
                        End If
                    Catch
                    End Try
                Next

                ' 2. Hole / Thread Notes — dùng Intent.PointOnSheet → tìm circle gần nhất
                For Each oNote As DrawingNote In oSheet.DrawingNotes
                    Try
                        If Not TypeOf oNote Is HoleThreadNote Then Continue For

                        Dim htNote As HoleThreadNote = CType(oNote, HoleThreadNote)

                        Dim anchorPt As Point2d = Nothing
                        Try
                            Dim intent As GeometryIntent = htNote.Intent
                            If intent IsNot Nothing Then anchorPt = intent.PointOnSheet
                        Catch
                        End Try

                        If anchorPt Is Nothing Then Continue For

                        Dim nearestKey As String = FindNearestCircleKey(oSheet, anchorPt)
                        If nearestKey <> "" AndAlso Not dimmedHoles.Contains(nearestKey) Then
                            dimmedHoles.Add(nearestKey)
                        End If

                    Catch
                    End Try
                Next

            Catch
            End Try

        End Sub

        '=================================================
        ' Key theo tọa độ tâm + bán kính
        '=================================================
        Private Function GetCenterKey(oCurve As DrawingCurve) As String
            Try
                Dim cp As Point2d = oCurve.CenterPoint
                If cp Is Nothing Then Return ""

                Dim r As Double = 0
                Try
                    r = oCurve.Radius
                Catch
                End Try

                Return "C_" & Math.Round(cp.X, 2).ToString() & "_" &
                      Math.Round(cp.Y, 2).ToString() & "_R" &
                      Math.Round(r, 2).ToString()
            Catch
                Return ""
            End Try
        End Function

        '=================================================
        ' Tìm circle gần vị trí anchor nhất
        '=================================================
        Private Function FindNearestCircleKey(oSheet As Sheet, anchorPt As Point2d) As String
            Try
                Dim bestKey As String = ""
                Dim bestDist As Double = Double.MaxValue

                For Each oView As DrawingView In oSheet.DrawingViews
                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType <> CurveTypeEnum.kCircleCurve Then Continue For

                            Dim cp As Point2d = oCurve.CenterPoint
                            If cp Is Nothing Then Continue For

                            Dim dx As Double = cp.X - anchorPt.X
                            Dim dy As Double = cp.Y - anchorPt.Y
                            Dim dist As Double = Math.Sqrt(dx * dx + dy * dy)

                            Dim r As Double = 0
                            Try
                                r = oCurve.Radius
                            Catch
                            End Try

                            ' Chỉ chấp nhận nếu dist xấp xỉ bán kính
                            If r > 0 AndAlso dist >= 0.5 * r AndAlso dist <= 2.0 * r Then
                                If dist < bestDist Then
                                    bestDist = dist
                                    bestKey = GetCenterKey(oCurve)
                                End If
                            End If
                        Catch
                        End Try
                    Next
                Next

                Return bestKey
            Catch
                Return ""
            End Try
        End Function

        '=================================================
        ' Kiểm tra lỗ thuộc Array / Pattern
        '=================================================
        Private Function IsPatternElement(oCurve As Inventor.DrawingCurve) As Boolean
            Try
                Dim modelGeom As Object = oCurve.ModelGeometry
                If modelGeom Is Nothing Then Return False

                If TypeOf modelGeom Is Inventor.Edge Then
                    Dim ed As Inventor.Edge = CType(modelGeom, Inventor.Edge)

                    For Each fc As Inventor.Face In ed.Faces
                        Try
                            Dim feat As Inventor.PartFeature = fc.CreatedByFeature
                            If feat Is Nothing Then Continue For

                            Dim tName As String = TypeName(feat).ToUpperInvariant()
                            If tName.Contains("PATTERN") Then Return True

                            If tName.Contains("HOLE") Then
                                Try
                                    Dim hf As Inventor.HoleFeature = CType(feat, Inventor.HoleFeature)
                                    If IsFeatureInsidePattern(hf) Then Return True
                                Catch
                                End Try
                            End If
                        Catch
                        End Try
                    Next
                End If
            Catch
            End Try
            Return False
        End Function

        Private Function IsFeatureInsidePattern(hf As Inventor.HoleFeature) As Boolean
            Try
                Dim partDef As Inventor.PartComponentDefinition = Nothing
                Try
                    partDef = TryCast(hf.Parent, Inventor.PartComponentDefinition)
                Catch
                End Try
                If partDef Is Nothing Then Return False

                For Each feat As Inventor.PartFeature In partDef.Features
                    Dim tName As String = TypeName(feat).ToUpperInvariant()
                    If tName.Contains("PATTERN") Then
                        Dim parentFeatures As Inventor.PartFeatures = Nothing

                        Try
                            Dim rectPat As Inventor.RectangularPatternFeature = CType(feat, Inventor.RectangularPatternFeature)
                            If rectPat.Definition IsNot Nothing Then
                                parentFeatures = rectPat.Definition.ParentFeatures
                            End If
                        Catch
                        End Try

                        If parentFeatures Is Nothing Then
                            Try
                                Dim circPat As Inventor.CircularPatternFeature = CType(feat, Inventor.CircularPatternFeature)
                                If circPat.Definition IsNot Nothing Then
                                    parentFeatures = circPat.Definition.ParentFeatures
                                End If
                            Catch
                            End Try
                        End If

                        If parentFeatures IsNot Nothing Then
                            For Each parentFeat As Inventor.PartFeature In parentFeatures
                                If parentFeat Is hf OrElse parentFeat.Name = hf.Name Then Return True
                            Next
                        End If
                    End If
                Next
            Catch
            End Try
            Return False
        End Function

        '=================================================
        ' Kiểm tra lỗ REN
        '=================================================
        Private Function IsThreadedHoleCurve(oCurve As Inventor.DrawingCurve) As Boolean
            Try
                Dim modelGeom As Object = oCurve.ModelGeometry
                If modelGeom Is Nothing Then Return False

                If TypeOf modelGeom Is Inventor.Edge Then
                    Dim ed As Inventor.Edge = CType(modelGeom, Inventor.Edge)

                    For Each fc As Inventor.Face In ed.Faces
                        Try
                            If fc.ThreadInfos IsNot Nothing AndAlso fc.ThreadInfos.Count > 0 Then
                                Return True
                            End If
                        Catch
                        End Try

                        Try
                            Dim feat As Inventor.PartFeature = fc.CreatedByFeature
                            If feat Is Nothing Then Continue For

                            Dim tName As String = TypeName(feat).ToUpperInvariant()
                            If tName.Contains("THREAD") Then Return True

                            If tName.Contains("HOLE") Then
                                Try
                                    Dim hf As Inventor.HoleFeature = CType(feat, Inventor.HoleFeature)
                                    If hf.Tapped Then Return True
                                Catch
                                End Try
                            End If
                        Catch
                        End Try
                    Next
                End If
            Catch
            End Try
            Return False
        End Function

    End Module

End Namespace