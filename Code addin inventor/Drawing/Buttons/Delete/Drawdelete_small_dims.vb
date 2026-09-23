Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic

Namespace ToolInventor2025.Drawing.Buttons.Drawdim
    Public Module Delete_small_dims

        '=============================================================
        ' XÓA DIM NHỎ HƠN NGƯỠNG
        ' - All: duyệt Sheet.DrawingDimensions
        ' - Nhiều view: lọc theo Intent / GeometryIntent + bounding box
        '=============================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim app As Inventor.Application = g_inventorApplication

            Try
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Vui lòng mở file Drawing (.idw)!", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument = CType(app.ActiveDocument, DrawingDocument)
                Dim oSheet As Sheet = oDrawDoc.ActiveSheet

                '===== 1. CHỌN PHẠM VI =====
                Dim scopeAnswer As DialogResult = MessageBox.Show(
                    "Chọn phạm vi xử lý:" & vbCrLf & vbCrLf &
                    "   [Yes]     = Tất cả view trên sheet" & vbCrLf &
                    "   [No]      = Chọn 1 hoặc nhiều view" & vbCrLf &
                    "   [Cancel]  = Hủy",
                    "Xóa dim nhỏ - Phạm vi",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question)

                If scopeAnswer = DialogResult.Cancel Then Exit Sub

                Dim processAll As Boolean = (scopeAnswer = DialogResult.Yes)
                Dim selectedViews As New List(Of DrawingView)

                '===== 2. CHỌN NHIỀU VIEW =====
                If Not processAll Then
                    Do
                        Dim pickedObj As Object = Nothing
                        Try
                            pickedObj = app.CommandManager.Pick(
                                SelectionFilterEnum.kDrawingViewFilter,
                                "Click vào view cần xóa dim. Bấm ESC để kết thúc.")
                        Catch
                        End Try

                        If pickedObj Is Nothing Then Exit Do

                        Dim v As DrawingView = TryCast(pickedObj, DrawingView)
                        If v IsNot Nothing Then
                            Dim dup As Boolean = False
                            For Each x As DrawingView In selectedViews
                                If x Is v Then
                                    dup = True
                                    Exit For
                                End If
                            Next
                            If Not dup Then selectedViews.Add(v)
                        End If

                        Dim more As DialogResult = MessageBox.Show(
                            "Đã chọn " & selectedViews.Count & " view." & vbCrLf & vbCrLf &
                            "Chọn thêm view nữa không?",
                            "Tiếp tục chọn",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question)

                        If more = DialogResult.No Then Exit Do
                    Loop

                    If selectedViews.Count = 0 Then
                        MessageBox.Show("Chưa chọn view nào. Hủy thao tác.",
                                        "Thông báo",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    End If
                End If

                '===== 3. NHẬP NGƯỠNG =====
                Dim input As String = InputBox(
                    "Nhập ngưỡng (mm)." & vbCrLf & vbCrLf &
                    "Xóa tất cả dim có giá trị NHỎ HƠN ngưỡng này." & vbCrLf &
                    "Ví dụ: nhập 5 để xóa các dim < 5mm.",
                    "Xóa dim nhỏ",
                    "5")

                If input Is Nothing OrElse input.Trim() = "" Then Exit Sub

                Dim thresholdMM As Double
                If Not Double.TryParse(input.Replace(",", ".").Trim(),
                                       Globalization.NumberStyles.Any,
                                       Globalization.CultureInfo.InvariantCulture,
                                       thresholdMM) Then
                    MessageBox.Show("Giá trị không hợp lệ.", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If thresholdMM <= 0 Then
                    MessageBox.Show("Ngưỡng phải > 0.", "Thông báo")
                    Exit Sub
                End If

                Dim thresholdCM As Double = thresholdMM / 10.0   ' Inventor dùng cm nội bộ

                '===== 4. QUÉT DIM TRÊN SHEET =====
                Dim toDelete As New List(Of DrawingDimension)
                Dim nTotal As Integer = 0
                Dim nSkipped As Integer = 0

                For Each oDim As DrawingDimension In oSheet.DrawingDimensions
                    Try
                        If Not processAll Then
                            If Not IsDimInAnyView(oDim, selectedViews) Then
                                nSkipped += 1
                                Continue For
                            End If
                        End If

                        nTotal += 1
                        AddIfSmall(oDim, thresholdCM, toDelete)
                    Catch
                    End Try
                Next

                '===== 5. XÓA =====
                Dim nDeleted As Integer = 0
                Dim nFail As Integer = 0

                For Each d As DrawingDimension In toDelete
                    Try
                        d.Delete()
                        nDeleted += 1
                    Catch
                        nFail += 1
                    End Try
                Next

                oDrawDoc.Update()

                Dim scopeText As String
                If processAll Then
                    scopeText = "Tất cả view"
                Else
                    scopeText = selectedViews.Count & " view đã chọn"
                End If

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Phạm vi: " & scopeText & vbCrLf &
                    "Tổng dim quét: " & nTotal & vbCrLf &
                    "Đã xóa: " & nDeleted & vbCrLf &
                    "Lỗi: " & nFail & vbCrLf & vbCrLf &
                    "Ngưỡng: " & thresholdMM.ToString("0.0#") & " mm",
                    "Xóa dim nhỏ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "Xóa dim nhỏ",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)
            End Try
        End Sub

        '=============================================================
        ' Kiểm tra dim có thuộc view nào trong danh sách không
        ' Ưu tiên Intent → GeometryIntent → bounding box text
        '=============================================================
        Private Function IsDimInAnyView(ByVal oDim As DrawingDimension,
                                         ByVal views As List(Of DrawingView)) As Boolean

            '--- Cách 1: LinearGeneralDimension.Intent (ổn định nhất trên 2020) ---
            Try
                Dim linDim As LinearGeneralDimension = TryCast(oDim, LinearGeneralDimension)
                If linDim IsNot Nothing Then
                    ' IntentOne
                    If CheckIntentBelongsToView(linDim.IntentOne, views) Then Return True
                    ' IntentTwo
                    If CheckIntentBelongsToView(linDim.IntentTwo, views) Then Return True
                End If
            Catch
            End Try

            '--- Cách 2: AttachedEntity (fallback) ---
            Try
                Dim ent As Object = Nothing
                Try
                    ent = oDim.AttachedEntity
                Catch
                End Try

                If ent IsNot Nothing Then
                    Dim parentView As DrawingView = TryCast(GetParentView(ent), DrawingView)
                    If parentView IsNot Nothing Then
                        For Each v As DrawingView In views
                            If v Is parentView Then Return True
                        Next
                    End If
                End If
            Catch
            End Try

            '--- Cách 3: Bounding box của text (cuối cùng) ---
            Try
                Dim tp As Point2d = Nothing
                Try
                    tp = oDim.Text.Origin
                Catch
                    Try
                        tp = oDim.Text.Position
                    Catch
                        Return False
                    End Try
                End Try

                If tp Is Nothing Then Return False

                For Each v As DrawingView In views
                    Try
                        Dim cx As Double = v.Position.X
                        Dim cy As Double = v.Position.Y
                        Dim hw As Double = v.Width / 2.0
                        Dim hh As Double = v.Height / 2.0

                        Dim vL As Double = cx - hw
                        Dim vR As Double = cx + hw
                        Dim vB As Double = cy - hh
                        Dim vT As Double = cy + hh

                        ' Thêm tolerance nhỏ vì text đôi khi nằm ngoài view một chút
                        Const tol As Double = 0.5   ' cm

                        If tp.X >= (vL - tol) AndAlso tp.X <= (vR + tol) AndAlso
                           tp.Y >= (vB - tol) AndAlso tp.Y <= (vT + tol) Then
                            Return True
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            Return False
        End Function

        '-------------------------------------------------------------
        ' Kiểm tra GeometryIntent có thuộc view đã chọn không
        '-------------------------------------------------------------
        Private Function CheckIntentBelongsToView(ByVal intent As Object,
                                                   ByVal views As List(Of DrawingView)) As Boolean
            If intent Is Nothing Then Return False

            Try
                Dim gi As GeometryIntent = TryCast(intent, GeometryIntent)
                If gi Is Nothing Then Return False

                Dim geom As Object = Nothing
                Try
                    geom = gi.Geometry
                Catch
                    Return False
                End Try

                If geom Is Nothing Then Return False

                Dim parentView As DrawingView = TryCast(GetParentView(geom), DrawingView)
                If parentView Is Nothing Then Return False

                For Each v As DrawingView In views
                    If v Is parentView Then Return True
                Next
            Catch
            End Try

            Return False
        End Function

        '-------------------------------------------------------------
        ' Lấy DrawingView cha của một geometry / entity
        '-------------------------------------------------------------
        Private Function GetParentView(ByVal obj As Object) As DrawingView
            If obj Is Nothing Then Return Nothing

            Try
                ' Nhiều object có property Parent là DrawingView
                Dim p As Object = Nothing
                Try
                    p = obj.Parent
                Catch
                End Try

                Dim dv As DrawingView = TryCast(p, DrawingView)
                If dv IsNot Nothing Then Return dv

                ' Một số trường hợp Parent là Sketch → Parent tiếp theo mới là View
                Try
                    If p IsNot Nothing Then
                        Dim p2 As Object = p.Parent
                        dv = TryCast(p2, DrawingView)
                        If dv IsNot Nothing Then Return dv
                    End If
                Catch
                End Try
            Catch
            End Try

            Return Nothing
        End Function

        '=============================================================
        ' Thêm dim vào list nếu là Linear và < ngưỡng
        '=============================================================
        Private Sub AddIfSmall(ByVal oDim As DrawingDimension,
                               ByVal thresholdCM As Double,
                               ByVal toDelete As List(Of DrawingDimension))
            Try
                Dim linDim As LinearGeneralDimension = TryCast(oDim, LinearGeneralDimension)
                If linDim Is Nothing Then Exit Sub

                Dim valCM As Double = 0
                Try
                    valCM = linDim.ModelValue
                Catch
                    Try
                        valCM = linDim.Value
                    Catch
                        Exit Sub
                    End Try
                End Try

                If Math.Abs(valCM) < thresholdCM Then
                    toDelete.Add(oDim)
                End If
            Catch
            End Try
        End Sub

    End Module
End Namespace