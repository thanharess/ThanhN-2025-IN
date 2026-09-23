Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic

Namespace ToolInventor2025.Drawing.Buttons.DrawPartList

    Public Module Create_PartsList

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                '=====================================================
                ' 1. KIỂM TRA DRAWING
                '=====================================================
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then

                    MessageBox.Show(
                        "Vui lòng mở file Drawing (.idw)!",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument =
                    CType(app.ActiveDocument, DrawingDocument)

                Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                Dim tg As TransientGeometry = app.TransientGeometry

                '=====================================================
                ' 2. CHỌN VIEW CỦA ASSEMBLY
                '=====================================================
                Dim pickedObj As Object = Nothing

                Try
                    oDrawDoc.SelectSet.Clear()
                Catch
                End Try

                Try
                    pickedObj = app.CommandManager.Pick(
                        SelectionFilterEnum.kDrawingViewFilter,
                        "Chọn view của Assembly cần tạo Parts List")
                Catch
                    Exit Sub
                End Try

                If pickedObj Is Nothing Then Exit Sub

                Dim oView As DrawingView = TryCast(pickedObj, DrawingView)
                If oView Is Nothing Then
                    MessageBox.Show(
                        "Không phải DrawingView.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
                    Exit Sub
                End If

                '=====================================================
                ' 3. KIỂM TRA VIEW CÓ PHẢI ASSEMBLY
                '=====================================================
                Dim refDoc As Document = Nothing
                Try
                    refDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                Catch
                End Try

                If refDoc Is Nothing Then
                    MessageBox.Show(
                        "View không có model tham chiếu.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
                    Exit Sub
                End If

                If refDoc.DocumentType <>
                   DocumentTypeEnum.kAssemblyDocumentObject Then

                    MessageBox.Show(
                        "View này không tham chiếu Assembly." &
                        vbCrLf & vbCrLf &
                        "Parts List chỉ tạo được từ view của Assembly.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
                    Exit Sub
                End If

                '=====================================================
                ' 4. CHỌN STYLE PARTS LIST
                '=====================================================
                Dim styleName As String = PickPartsListStyle(oDrawDoc)

                If styleName = "" Then Exit Sub

                '=====================================================
                ' 5. CHỌN VỊ TRÍ THEO KHUNG BẢN VẼ (KHÔNG NHẬP OFFSET)
                '=====================================================
                Dim chosenPos As Point2d = PickPartsListPosition(oSheet, tg)

                If chosenPos Is Nothing Then Exit Sub

                '=====================================================
                ' 6. KIỂM TRA SHEET ĐANG ACTIVE
                '=====================================================
                Try
                    oSheet.Activate()

                    If oSheet.PartsLists.Count > 0 Then
                        MessageBox.Show(
                            "Sheet này đã có Parts List." &
                            vbCrLf & vbCrLf &
                            "Không tạo thêm Parts List.",
                            "Tạo Parts List",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
                        Exit Sub
                    End If
                Catch
                End Try

                '=====================================================
                ' 7. TẠO PARTS LIST TẠI VỊ TRÍ AN TOÀN TRƯỚC
                '=====================================================
                Dim safeX As Double = oSheet.Width / 2.0
                Dim safeY As Double = oSheet.Height / 2.0
                Dim safePos As Point2d = tg.CreatePoint2d(safeX, safeY)

                Dim oPL As PartsList = Nothing

                Try
                    oPL = oSheet.PartsLists.Add(oView, safePos)
                Catch ex As Exception
                    MessageBox.Show(
                        "Không tạo được Parts List:" &
                        vbCrLf & vbCrLf &
                        ex.Message &
                        vbCrLf & vbCrLf &
                        "View: " & oView.Name &
                        vbCrLf &
                        "Sheet: " & oSheet.Name,
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
                    Exit Sub
                End Try

                If oPL Is Nothing Then
                    MessageBox.Show(
                        "Parts List trả về Nothing.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
                    Exit Sub
                End If

                '=====================================================
                ' 8. GÁN STYLE
                '=====================================================
                Try
                    oPL.Style =
                        oDrawDoc.StylesManager.PartsListStyles.Item(styleName)
                Catch ex As Exception
                    MessageBox.Show(
                        "Không thể áp dụng Style:" &
                        vbCrLf & vbCrLf &
                        styleName &
                        vbCrLf & vbCrLf &
                        ex.Message,
                        "Cảnh báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
                End Try

                '=====================================================
                ' 9. SORT
                '=====================================================
                Try
                    oPL.Sort("ITEM")
                Catch
                    Try
                        oPL.Sort("Item")
                    Catch
                    End Try
                End Try

                '=====================================================
                ' 10. UPDATE
                '=====================================================
                Try
                    oPL.Update()
                Catch
                End Try

                Try
                    oDrawDoc.Update()
                Catch
                End Try

                '=====================================================
                ' 11. DI CHUYỂN ĐẾN VỊ TRÍ ĐÃ CHỌN
                '=====================================================
                Try
                    oPL.Position = chosenPos
                Catch
                    ' Nếu vị trí không hợp lệ → giữ nguyên vị trí an toàn
                End Try

            Catch ex As Exception
                MessageBox.Show(
                    "Lỗi tổng:" & vbCrLf & vbCrLf & ex.Message,
                    "Tạo Parts List",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End Try

        End Sub

        '=============================================================
        ' CHỌN VỊ TRÍ PARTS LIST THEO KHUNG BẢN VẼ
        ' Trả về Point2d hoặc Nothing nếu hủy
        '=============================================================
        Private Function PickPartsListPosition(
            ByVal oSheet As Sheet,
            ByVal tg As TransientGeometry) As Point2d

            Try
                ' Danh sách vị trí cứng theo khung bản vẽ
                ' Tọa độ tính theo cm (Inventor internal)
                Dim positions As New Dictionary(Of String, Point2d)

                ' ---- A4 ngang (297 x 210 mm ≈ 29.7 x 21.0 cm) ----
                positions.Add("A4 - Góc trên phải", tg.CreatePoint2d(26.0, 18.5))
                positions.Add("A4 - Góc trên trái", tg.CreatePoint2d(3.5, 18.5))
                positions.Add("A4 - Góc dưới phải", tg.CreatePoint2d(26.0, 3.5))

                ' ---- A3 ngang (420 x 297 mm ≈ 42.0 x 29.7 cm) ----
                positions.Add("A3 - Góc trên phải", tg.CreatePoint2d(38.0, 26.5))
                positions.Add("A3 - Góc trên trái", tg.CreatePoint2d(4.0, 26.5))
                positions.Add("A3 - Góc dưới phải", tg.CreatePoint2d(38.0, 4.0))

                ' ---- A2 ngang (594 x 420 mm ≈ 59.4 x 42.0 cm) ----
                positions.Add("A2 - Góc trên phải", tg.CreatePoint2d(54.0, 38.0))
                positions.Add("A2 - Góc trên trái", tg.CreatePoint2d(5.0, 38.0))
                positions.Add("A2 - Góc dưới phải", tg.CreatePoint2d(54.0, 5.0))

                ' ---- A1 ngang (841 x 594 mm ≈ 84.1 x 59.4 cm) ----
                positions.Add("A1 - Góc trên phải", tg.CreatePoint2d(78.0, 54.0))
                positions.Add("A1 - Góc trên trái", tg.CreatePoint2d(6.0, 54.0))
                positions.Add("A1 - Góc dưới phải", tg.CreatePoint2d(78.0, 6.0))

                ' ---- A0 ngang (1189 x 841 mm ≈ 118.9 x 84.1 cm) ----
                positions.Add("A0 - Góc trên phải", tg.CreatePoint2d(110.0, 78.0))
                positions.Add("A0 - Góc trên trái", tg.CreatePoint2d(8.0, 78.0))
                positions.Add("A0 - Góc dưới phải", tg.CreatePoint2d(110.0, 8.0))

                ' Form chọn
                Dim frm As New Form
                frm.Text = "Chọn vị trí Parts List theo khung bản vẽ"
                frm.ClientSize = New System.Drawing.Size(420, 360)
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.FormBorderStyle = FormBorderStyle.FixedDialog
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False

                Dim lst As New ListBox
                lst.Bounds = New System.Drawing.Rectangle(12, 12, 396, 290)
                lst.Font = New System.Drawing.Font("Segoe UI", 9.0F)

                For Each key As String In positions.Keys
                    lst.Items.Add(key)
                Next

                If lst.Items.Count > 0 Then
                    lst.SelectedIndex = 0
                End If

                Dim btnOK As New Button
                btnOK.Text = "OK"
                btnOK.Bounds = New System.Drawing.Rectangle(230, 315, 85, 28)
                btnOK.DialogResult = DialogResult.OK

                Dim btnCancel As New Button
                btnCancel.Text = "Hủy"
                btnCancel.Bounds = New System.Drawing.Rectangle(323, 315, 85, 28)
                btnCancel.DialogResult = DialogResult.Cancel

                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel

                AddHandler lst.DoubleClick,
                    Sub(sender As Object, e As EventArgs)
                        If lst.SelectedItem IsNot Nothing Then
                            frm.DialogResult = DialogResult.OK
                            frm.Close()
                        End If
                    End Sub

                frm.Controls.Add(lst)
                frm.Controls.Add(btnOK)
                frm.Controls.Add(btnCancel)

                Dim result As DialogResult = frm.ShowDialog()

                If result <> DialogResult.OK OrElse lst.SelectedItem Is Nothing Then
                    Return Nothing
                End If

                Dim selectedKey As String = lst.SelectedItem.ToString()
                Return positions(selectedKey)

            Catch ex As Exception
                MessageBox.Show(
                    "Lỗi khi chọn vị trí:" & vbCrLf & vbCrLf & ex.Message,
                    "Tạo Parts List",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
                Return Nothing
            End Try

        End Function

        '=============================================================
        ' CHỌN STYLE PARTS LIST
        '=============================================================
        Private Function PickPartsListStyle(
            ByVal oDrawDoc As DrawingDocument) As String

            Try
                Dim styles As Inventor.PartsListStylesEnumerator =
                    oDrawDoc.StylesManager.PartsListStyles

                If styles Is Nothing OrElse styles.Count = 0 Then
                    Return ""
                End If

                Dim names As New List(Of String)

                For i As Integer = 1 To styles.Count
                    Try
                        Dim style As Inventor.PartsListStyle = styles.Item(i)
                        If style IsNot Nothing Then
                            names.Add(style.Name)
                        End If
                    Catch
                    End Try
                Next

                If names.Count = 0 Then Return ""
                If names.Count = 1 Then Return names(0)

                Dim frm As New Form
                frm.Text = "Chọn Style Parts List"
                frm.ClientSize = New System.Drawing.Size(420, 320)
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.FormBorderStyle = FormBorderStyle.FixedDialog
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False

                Dim lst As New ListBox
                lst.Bounds = New System.Drawing.Rectangle(12, 12, 396, 250)
                lst.Font = New System.Drawing.Font("Segoe UI", 9.0F)

                For Each n As String In names
                    lst.Items.Add(n)
                Next

                If lst.Items.Count > 0 Then
                    lst.SelectedIndex = 0
                End If

                Dim btnOK As New Button
                btnOK.Text = "OK"
                btnOK.Bounds = New System.Drawing.Rectangle(230, 275, 85, 28)
                btnOK.DialogResult = DialogResult.OK

                Dim btnCancel As New Button
                btnCancel.Text = "Hủy"
                btnCancel.Bounds = New System.Drawing.Rectangle(323, 275, 85, 28)
                btnCancel.DialogResult = DialogResult.Cancel

                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel

                AddHandler lst.DoubleClick,
                    Sub(sender As Object, e As EventArgs)
                        If lst.SelectedItem IsNot Nothing Then
                            frm.DialogResult = DialogResult.OK
                            frm.Close()
                        End If
                    End Sub

                frm.Controls.Add(lst)
                frm.Controls.Add(btnOK)
                frm.Controls.Add(btnCancel)

                Dim result As DialogResult = frm.ShowDialog()

                If result <> DialogResult.OK OrElse lst.SelectedItem Is Nothing Then
                    Return ""
                End If

                Return lst.SelectedItem.ToString()

            Catch ex As Exception
                MessageBox.Show(
                    "Lỗi khi chọn Style:" & vbCrLf & vbCrLf & ex.Message,
                    "Tạo Parts List",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
                Return ""
            End Try

        End Function

    End Module

End Namespace