Imports System.Windows.Forms
Imports System.Collections.Generic
Imports Inventor
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.DrawSheet

    ' ============================================================
    ' MODULE CHÍNH
    ' ============================================================
    Public Module Drawing_TitleBlock_Border

        ' =====================================================
        ' ENTRY POINT
        ' =====================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                Dim invApp As Inventor.Application = GetInventorApp()
                If invApp Is Nothing Then
                    MessageBox.Show("Không lấy được Inventor Application!", "Drawing Tool")
                    Return
                End If

                If invApp.ActiveDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Mở file Drawing (.idw/.dwg) trước.", "Drawing Tool")
                    Return
                End If

                Dim frm As New Form_ChooseAction()
                frm.ShowDialog()
                frm.Dispose()

            Catch ex As Exception
                Try
                    MessageBox.Show("Lỗi: " & ex.Message, "Drawing Tool")
                Catch
                End Try
            End Try
        End Sub


        ' =====================================================
        ' API CÔNG KHAI — FORM GỌI
        ' =====================================================
        Public Function GetTitleBlockNames() As List(Of String)
            Dim result As New List(Of String)
            Try
                Dim invApp = GetInventorApp()
                If invApp Is Nothing Then Return result
                If invApp.ActiveDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then Return result

                Dim drawDoc As Inventor.DrawingDocument = CType(invApp.ActiveDocument, Inventor.DrawingDocument)
                For Each def As Inventor.TitleBlockDefinition In drawDoc.TitleBlockDefinitions
                    result.Add(def.Name)
                Next
            Catch
            End Try
            Return result
        End Function

        Public Function GetBorderNames() As List(Of String)
            Dim result As New List(Of String)
            Try
                Dim invApp = GetInventorApp()
                If invApp Is Nothing Then Return result
                If invApp.ActiveDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then Return result

                Dim drawDoc As Inventor.DrawingDocument = CType(invApp.ActiveDocument, Inventor.DrawingDocument)
                For Each def As Inventor.BorderDefinition In drawDoc.BorderDefinitions
                    result.Add(def.Name)
                Next
            Catch
            End Try
            Return result
        End Function


        ' =====================================================
        ' XÓA
        ' =====================================================
        Public Sub DeleteAllTitleBlocks()
            RunDelete("Xóa Title Block — tất cả sheet", "ALL", True)
        End Sub
        Public Sub DeleteAllBorders()
            RunDelete("Xóa Border — tất cả sheet", "ALL", False)
        End Sub
        Public Sub DeleteActiveTitleBlock()
            RunDelete("Xóa Title Block — sheet active", "ACTIVE", True)
        End Sub
        Public Sub DeleteActiveBorder()
            RunDelete("Xóa Border — sheet active", "ACTIVE", False)
        End Sub
        Public Sub DeleteSelectedTitleBlock()
            RunDelete("Xóa Title Block — sheet chọn", "SELECTED", True)
        End Sub
        Public Sub DeleteSelectedBorder()
            RunDelete("Xóa Border — sheet chọn", "SELECTED", False)
        End Sub


        ' =====================================================
        ' THAY
        ' =====================================================
        Public Sub ReplaceAllTitleBlocks(tbName As String)
            RunReplace("Thay Title Block — tất cả sheet", "ALL", tbName, Nothing)
        End Sub
        Public Sub ReplaceAllBorders(bdName As String)
            RunReplace("Thay Border — tất cả sheet", "ALL", Nothing, bdName)
        End Sub
        Public Sub ReplaceActiveTitleBlock(tbName As String)
            RunReplace("Thay Title Block — sheet active", "ACTIVE", tbName, Nothing)
        End Sub
        Public Sub ReplaceActiveBorder(bdName As String)
            RunReplace("Thay Border — sheet active", "ACTIVE", Nothing, bdName)
        End Sub
        Public Sub ReplaceSelectedTitleBlock(tbName As String)
            RunReplace("Thay Title Block — sheet chọn", "SELECTED", tbName, Nothing)
        End Sub
        Public Sub ReplaceSelectedBorder(bdName As String)
            RunReplace("Thay Border — sheet chọn", "SELECTED", Nothing, bdName)
        End Sub

        Public Sub ReplaceBothAllSheets(tbName As String, bdName As String)
            RunReplace("Thay TB + Border — tất cả sheet", "ALL", tbName, bdName)
        End Sub
        Public Sub ReplaceBothActiveSheet(tbName As String, bdName As String)
            RunReplace("Thay TB + Border — sheet active", "ACTIVE", tbName, bdName)
        End Sub
        Public Sub ReplaceBothSelectedSheet(tbName As String, bdName As String)
            RunReplace("Thay TB + Border — sheet chọn", "SELECTED", tbName, bdName)
        End Sub


        ' =====================================================
        ' HÀM LÕI — XÓA
        ' =====================================================
        Private Sub RunDelete(title As String, scope As String, isTB As Boolean)

            Dim invApp As Inventor.Application = Nothing
            Dim drawDoc As Inventor.DrawingDocument = Nothing
            If Not ValidateDrawing(invApp, drawDoc, title) Then Return

            Dim count As Integer = 0
            Dim skip As Integer = 0
            Dim fail As Integer = 0

            invApp.SilentOperation = True
            Try
                Select Case scope
                    Case "ALL"
                        For Each oSheet As Inventor.Sheet In drawDoc.Sheets
                            Dim r = DeleteOnSheet(oSheet, isTB)
                            If r = 1 Then count += 1
                            If r = 0 Then skip += 1
                            If r = -1 Then fail += 1
                        Next
                    Case "ACTIVE"
                        Dim s As Inventor.Sheet = drawDoc.ActiveSheet
                        If s IsNot Nothing Then
                            Dim r = DeleteOnSheet(s, isTB)
                            If r = 1 Then count += 1
                            If r = 0 Then skip += 1
                            If r = -1 Then fail += 1
                        End If
                    Case "SELECTED"
                        Dim s As Inventor.Sheet = PickSheet(drawDoc, title)
                        If s IsNot Nothing Then
                            Dim r = DeleteOnSheet(s, isTB)
                            If r = 1 Then count += 1
                            If r = 0 Then skip += 1
                            If r = -1 Then fail += 1
                        End If
                End Select
            Catch
            End Try
            invApp.SilentOperation = False

            Try : drawDoc.Update2(True) : Catch : End Try

            Dim msg As String =
                "Đã xóa: " & count.ToString() & vbCrLf
            If skip > 0 Then msg &= "Bỏ qua (không có): " & skip.ToString() & vbCrLf
            If fail > 0 Then msg &= "Lỗi: " & fail.ToString()

            MessageBox.Show(msg, title)
        End Sub


        ''' <summary>
        ''' 1 = đã xóa | 0 = không có gì để xóa | -1 = lỗi
        ''' </summary>
        Private Function DeleteOnSheet(oSheet As Inventor.Sheet, isTB As Boolean) As Integer
            Try
                If isTB Then
                    If oSheet.TitleBlock IsNot Nothing Then
                        oSheet.TitleBlock.Delete()
                        Return 1
                    Else
                        Return 0
                    End If
                Else
                    If oSheet.Border IsNot Nothing Then
                        oSheet.Border.Delete()
                        Return 1
                    Else
                        Return 0
                    End If
                End If
            Catch
                Return -1
            End Try
        End Function


        ' =====================================================
        ' HÀM LÕI — THAY
        ' =====================================================
        Private Sub RunReplace(title As String,
                               scope As String,
                               tbName As String,
                               bdName As String)

            Dim invApp As Inventor.Application = Nothing
            Dim drawDoc As Inventor.DrawingDocument = Nothing
            If Not ValidateDrawing(invApp, drawDoc, title) Then Return

            '=================================================
            ' ⭐ KIỂM TRA DEFINITION TRƯỚC KHI CHẠY
            '=================================================
            Dim tbDef As Inventor.TitleBlockDefinition = Nothing
            Dim bdDef As Inventor.BorderDefinition = Nothing

            If tbName IsNot Nothing Then
                If Not TryGetTitleBlockDef(drawDoc, tbName, tbDef) Then
                    MessageBox.Show("Không tìm thấy Title Block mẫu: '" & tbName & "'" & vbCrLf & vbCrLf &
                                    "Có thể mẫu này chưa được load vào file." & vbCrLf &
                                    "Mở Sheet Format → chọn mẫu → OK, rồi chạy lại.",
                                    title, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
            End If

            If bdName IsNot Nothing Then
                If Not TryGetBorderDef(drawDoc, bdName, bdDef) Then
                    MessageBox.Show("Không tìm thấy Border mẫu: '" & bdName & "'" & vbCrLf & vbCrLf &
                                    "Có thể mẫu này chưa được load vào file." & vbCrLf &
                                    "Mở Sheet Format → chọn mẫu → OK, rồi chạy lại.",
                                    title, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
            End If

            '=================================================
            ' THU THẬP SHEET ĐÍCH
            '=================================================
            Dim targetSheets As New List(Of Inventor.Sheet)

            Select Case scope
                Case "ALL"
                    For Each s As Inventor.Sheet In drawDoc.Sheets
                        targetSheets.Add(s)
                    Next

                Case "ACTIVE"
                    Dim s As Inventor.Sheet = drawDoc.ActiveSheet
                    If s IsNot Nothing Then targetSheets.Add(s)

                Case "SELECTED"
                    Dim s As Inventor.Sheet = PickSheet(drawDoc, title)
                    If s IsNot Nothing Then targetSheets.Add(s)
            End Select

            If targetSheets.Count = 0 Then
                MessageBox.Show("Không có sheet nào để xử lý.", title)
                Return
            End If

            '=================================================
            ' XỬ LÝ
            '=================================================
            Dim originalSheet As Inventor.Sheet = drawDoc.ActiveSheet
            Dim result As New List(Of SheetResult)

            invApp.SilentOperation = True
            Try
                For Each oSheet As Inventor.Sheet In targetSheets

                    ' ⭐ Activate sheet trước khi thao tác (BẮT BUỘC cho Inventor 2025)
                    Try : oSheet.Activate() : Catch : End Try
                    Try : drawDoc.Update() : Catch : End Try

                    Dim r As New SheetResult() With {
                        .SheetName = oSheet.Name
                    }

                    '=============================================
                    ' TITLE BLOCK
                    '=============================================
                    If tbName IsNot Nothing Then
                        Dim r2 = ReplaceTBOnSheet_Ex(oSheet, tbDef, tbName)
                        r.TB_Status = r2
                    End If

                    '=============================================
                    ' BORDER — có kiểm tra riêng
                    '=============================================
                    If bdName IsNot Nothing Then
                        Dim r3 = ReplaceBDOnSheet_Ex(oSheet, bdDef, bdName)
                        r.BD_Status = r3
                    End If

                    result.Add(r)
                Next
            Catch
            End Try

            ' Restore sheet ban đầu
            Try
                If originalSheet IsNot Nothing Then originalSheet.Activate()
            Catch
            End Try

            invApp.SilentOperation = False
            Try : drawDoc.Update2(True) : Catch : End Try

            '=================================================
            ' BÁO CÁO KẾT QUẢ
            '=================================================
            Dim msg As New System.Text.StringBuilder()
            msg.AppendLine("Kết quả xử lý " & result.Count & " sheet:")
            msg.AppendLine()

            Dim cntTB_OK As Integer = 0
            Dim cntTB_Skip As Integer = 0
            Dim cntTB_Fail As Integer = 0
            Dim cntBD_OK As Integer = 0
            Dim cntBD_Skip As Integer = 0
            Dim cntBD_Fail As Integer = 0

            For Each r In result
                If tbName IsNot Nothing Then
                    If r.TB_Status = ReplaceStatus.Success Then cntTB_OK += 1
                    If r.TB_Status = ReplaceStatus.Skipped Then cntTB_Skip += 1
                    If r.TB_Status = ReplaceStatus.Failed Then cntTB_Fail += 1
                End If

                If bdName IsNot Nothing Then
                    If r.BD_Status = ReplaceStatus.Success Then cntBD_OK += 1
                    If r.BD_Status = ReplaceStatus.Skipped Then cntBD_Skip += 1
                    If r.BD_Status = ReplaceStatus.Failed Then cntBD_Fail += 1
                End If
            Next

            If tbName IsNot Nothing Then
                msg.AppendLine("── TITLE BLOCK ──")
                msg.AppendLine("  ✔ Thay OK   : " & cntTB_OK.ToString())
                If cntTB_Skip > 0 Then msg.AppendLine("  ⊘ Bỏ qua    : " & cntTB_Skip.ToString())
                If cntTB_Fail > 0 Then msg.AppendLine("  ✘ Lỗi       : " & cntTB_Fail.ToString())
                msg.AppendLine()
            End If

            If bdName IsNot Nothing Then
                msg.AppendLine("── BORDER ──")
                msg.AppendLine("  ✔ Thay OK   : " & cntBD_OK.ToString())
                If cntBD_Skip > 0 Then msg.AppendLine("  ⊘ Bỏ qua    : " & cntBD_Skip.ToString())
                If cntBD_Fail > 0 Then msg.AppendLine("  ✘ Lỗi       : " & cntBD_Fail.ToString())
            End If

            If cntTB_Fail > 0 OrElse cntBD_Fail > 0 Then
                msg.AppendLine()
                msg.AppendLine("⚠ Có sheet lỗi. Chi tiết:")
                For Each r In result
                    If r.TB_Status = ReplaceStatus.Failed OrElse r.BD_Status = ReplaceStatus.Failed Then
                        msg.AppendLine("  • " & r.SheetName)
                    End If
                Next
            End If

            MessageBox.Show(msg.ToString(), title,
                            MessageBoxButtons.OK,
                            If(cntTB_Fail > 0 OrElse cntBD_Fail > 0,
                               MessageBoxIcon.Warning,
                               MessageBoxIcon.Information))
        End Sub


        ' =====================================================
        ' TRẠNG THÁI KẾT QUẢ
        ' =====================================================
        Private Enum ReplaceStatus
            NotApplied = 0
            Success = 1
            Skipped = 2
            Failed = -1
        End Enum

        Private Class SheetResult
            Public SheetName As String = ""
            Public TB_Status As ReplaceStatus = ReplaceStatus.NotApplied
            Public BD_Status As ReplaceStatus = ReplaceStatus.NotApplied
        End Class


        ' =====================================================
        ' THAY TITLE BLOCK — CÓ KIỂM TRA
        ' =====================================================
        Private Function ReplaceTBOnSheet_Ex(oSheet As Inventor.Sheet,
                                             tbDef As Inventor.TitleBlockDefinition,
                                             tbName As String) As ReplaceStatus
            Try
                Dim drawDoc As Inventor.DrawingDocument = oSheet.Parent

                '=============================================
                ' KIỂM TRA: Đã có Title Block chưa
                '=============================================
                Dim hadOld As Boolean = (oSheet.TitleBlock IsNot Nothing)

                '=============================================
                ' KIỂM TRA: Title Block hiện tại có phải là mẫu cần thay không?
                ' Nếu trùng tên → bỏ qua
                '=============================================
                If hadOld Then
                    Try
                        Dim oldName As String = oSheet.TitleBlock.Definition.Name
                        If String.Equals(oldName, tbName, StringComparison.OrdinalIgnoreCase) Then
                            Return ReplaceStatus.Skipped
                        End If
                    Catch
                    End Try
                End If

                '=============================================
                ' XÓA CŨ
                '=============================================
                If hadOld Then
                    Try
                        oSheet.TitleBlock.Delete()
                        Try : drawDoc.Update() : Catch : End Try
                    Catch
                        Return ReplaceStatus.Failed
                    End Try
                End If

                '=============================================
                ' THÊM MỚI
                '=============================================
                Try
                    oSheet.AddTitleBlock(tbDef)
                    Try : drawDoc.Update() : Catch : End Try
                Catch
                    Return ReplaceStatus.Failed
                End Try

                '=============================================
                ' ⭐ VERIFY: Kiểm tra đã add thành công
                '=============================================
                Try
                    If oSheet.TitleBlock Is Nothing Then
                        Return ReplaceStatus.Failed
                    End If

                    ' Kiểm tra tên definition khớp
                    Dim newName As String = ""
                    Try : newName = oSheet.TitleBlock.Definition.Name : Catch : End Try

                    If Not String.IsNullOrEmpty(newName) AndAlso
                       Not String.Equals(newName, tbName, StringComparison.OrdinalIgnoreCase) Then
                        Return ReplaceStatus.Failed
                    End If
                Catch
                End Try

                Return ReplaceStatus.Success

            Catch
                Return ReplaceStatus.Failed
            End Try
        End Function


        ' =====================================================
        ' ⭐ THAY BORDER — KIỂM TRA CHI TIẾT
        ' =====================================================
        Private Function ReplaceBDOnSheet_Ex(oSheet As Inventor.Sheet,
                                             bdDef As Inventor.BorderDefinition,
                                             bdName As String) As ReplaceStatus
            Try
                Dim drawDoc As Inventor.DrawingDocument = oSheet.Parent

                '=============================================
                ' KIỂM TRA 1: Border definition hợp lệ
                '=============================================
                If bdDef Is Nothing Then Return ReplaceStatus.Failed

                '=============================================
                ' KIỂM TRA 2: Sheet có tồn tại Border cũ không
                '=============================================
                Dim oldBorder As Inventor.Border = Nothing
                Dim hadOld As Boolean = False

                Try
                    oldBorder = oSheet.Border
                    hadOld = (oldBorder IsNot Nothing)
                Catch
                    hadOld = False
                End Try

                '=============================================
                ' KIỂM TRA 3: Border cũ có trùng tên mẫu mới không?
                ' Nếu trùng → bỏ qua, không làm gì
                '=============================================
                If hadOld AndAlso oldBorder IsNot Nothing Then
                    Try
                        Dim oldName As String = oldBorder.Definition.Name
                        If String.Equals(oldName, bdName, StringComparison.OrdinalIgnoreCase) Then
                            Return ReplaceStatus.Skipped
                        End If
                    Catch
                    End Try
                End If

                '=============================================
                ' KIỂM TRA 4: Border cũ có bị khóa không? (nếu API hỗ trợ)
                '=============================================
                If hadOld AndAlso oldBorder IsNot Nothing Then
                    Try
                        ' Kiểm tra quyền xóa — một số Border bị khóa bởi Sheet Format
                        ' Nếu không xóa được → return Failed
                        oldBorder.Delete()
                        Try : drawDoc.Update() : Catch : End Try
                    Catch
                        ' Không xóa được Border cũ
                        Return ReplaceStatus.Failed
                    End Try
                End If

                '=============================================
                ' THÊM BORDER MỚI
                '=============================================
                Try
                    oSheet.AddBorder(bdDef)
                    Try : drawDoc.Update() : Catch : End Try
                Catch
                    Return ReplaceStatus.Failed
                End Try

                '=============================================
                ' ⭐ VERIFY 1: Kiểm tra đã có Border sau khi add
                '=============================================
                Dim newBorder As Inventor.Border = Nothing
                Try
                    newBorder = oSheet.Border
                    If newBorder Is Nothing Then
                        Return ReplaceStatus.Failed
                    End If
                Catch
                    Return ReplaceStatus.Failed
                End Try

                '=============================================
                ' ⭐ VERIFY 2: Kiểm tra tên definition của Border mới
                '=============================================
                Try
                    Dim newName As String = newBorder.Definition.Name
                    If Not String.IsNullOrEmpty(newName) AndAlso
                       Not String.Equals(newName, bdName, StringComparison.OrdinalIgnoreCase) Then
                        Return ReplaceStatus.Failed
                    End If
                Catch
                End Try

                '=============================================
                ' ⭐ VERIFY 3: Kiểm tra Border mới có khác Border cũ
                '    (nếu Border mới giống hệt Border cũ → có thể là lỗi)
                '=============================================
                If hadOld AndAlso oldBorder IsNot Nothing Then
                    Try
                        Dim oldName As String = oldBorder.Definition.Name
                        If String.Equals(oldName, bdName, StringComparison.OrdinalIgnoreCase) Then
                            ' Đã xử lý ở bước trên
                            Return ReplaceStatus.Skipped
                        End If
                    Catch
                    End Try
                End If

                Return ReplaceStatus.Success

            Catch
                Return ReplaceStatus.Failed
            End Try
        End Function


        ' =====================================================
        ' TÌM DEFINITION AN TOÀN
        ' =====================================================
        Private Function TryGetTitleBlockDef(drawDoc As Inventor.DrawingDocument,
                                             tbName As String,
                                             ByRef tbDef As Inventor.TitleBlockDefinition) As Boolean
            tbDef = Nothing
            Try
                For Each d As Inventor.TitleBlockDefinition In drawDoc.TitleBlockDefinitions
                    Try
                        If String.Equals(d.Name, tbName, StringComparison.OrdinalIgnoreCase) Then
                            tbDef = d
                            Return True
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try
            Return False
        End Function

        Private Function TryGetBorderDef(drawDoc As Inventor.DrawingDocument,
                                         bdName As String,
                                         ByRef bdDef As Inventor.BorderDefinition) As Boolean
            bdDef = Nothing
            Try
                For Each d As Inventor.BorderDefinition In drawDoc.BorderDefinitions
                    Try
                        If String.Equals(d.Name, bdName, StringComparison.OrdinalIgnoreCase) Then
                            bdDef = d
                            Return True
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try
            Return False
        End Function


        ' =====================================================
        ' CHỌN SHEET
        ' =====================================================
        Private Function PickSheet(drawDoc As Inventor.DrawingDocument, title As String) As Inventor.Sheet
            Dim names As New List(Of String)
            For Each s As Inventor.Sheet In drawDoc.Sheets
                names.Add(s.Name)
            Next

            If names.Count = 0 Then Return Nothing

            Dim sel As String = ChooseSheetDialog(names, title)
            If String.IsNullOrEmpty(sel) Then Return Nothing

            Try
                Return drawDoc.Sheets.Item(sel)
            Catch
                Return Nothing
            End Try
        End Function

        Private Function ChooseSheetDialog(names As List(Of String), title As String) As String
            Dim frm As New System.Windows.Forms.Form()
            frm.Text = title
            frm.ClientSize = New Drw.Size(340, 320)
            frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.Font = New Drw.Font("Segoe UI", 9.5F)

            Dim lst As New System.Windows.Forms.ListBox With {
                .Location = New Drw.Point(15, 15),
                .Size = New Drw.Size(310, 240),
                .Font = New Drw.Font("Segoe UI", 10.0F)
            }
            For Each s In names
                lst.Items.Add(s)
            Next
            If lst.Items.Count > 0 Then lst.SelectedIndex = 0

            Dim btnOK As New System.Windows.Forms.Button With {
                .Text = "OK",
                .Location = New Drw.Point(150, 270),
                .Size = New Drw.Size(80, 30),
                .DialogResult = System.Windows.Forms.DialogResult.OK
            }
            Dim btnCancel As New System.Windows.Forms.Button With {
                .Text = "Cancel",
                .Location = New Drw.Point(245, 270),
                .Size = New Drw.Size(80, 30),
                .DialogResult = System.Windows.Forms.DialogResult.Cancel
            }

            frm.Controls.Add(lst)
            frm.Controls.Add(btnOK)
            frm.Controls.Add(btnCancel)
            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK AndAlso lst.SelectedItem IsNot Nothing Then
                Return lst.SelectedItem.ToString()
            End If
            Return ""
        End Function


        ' =====================================================
        ' HELPERS
        ' =====================================================
        Private Function ValidateDrawing(ByRef invApp As Inventor.Application,
                                          ByRef drawDoc As Inventor.DrawingDocument,
                                          title As String) As Boolean
            invApp = GetInventorApp()
            If invApp Is Nothing Then
                MessageBox.Show("Không lấy được Inventor Application!", title)
                Return False
            End If
            If invApp.ActiveDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                MessageBox.Show("Mở file Drawing trước.", title)
                Return False
            End If
            drawDoc = CType(invApp.ActiveDocument, Inventor.DrawingDocument)
            Return True
        End Function

        Private Function GetInventorApp() As Inventor.Application
            Try
                Return CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch
                Return Nothing
            End Try
        End Function

    End Module


    ' ============================================================
    ' FORM
    ' ============================================================
    Public Class Form_ChooseAction
        Inherits System.Windows.Forms.Form

        Private _rdoAll As System.Windows.Forms.RadioButton
        Private _rdoActive As System.Windows.Forms.RadioButton
        Private _rdoSelected As System.Windows.Forms.RadioButton
        Private _rdoDelTB As System.Windows.Forms.RadioButton
        Private _rdoDelBD As System.Windows.Forms.RadioButton
        Private _rdoRepTB As System.Windows.Forms.RadioButton
        Private _rdoRepBD As System.Windows.Forms.RadioButton
        Private _rdoRepBoth As System.Windows.Forms.RadioButton
        Private _cboTitleBlock As System.Windows.Forms.ComboBox
        Private _cboBorder As System.Windows.Forms.ComboBox
        Private _lblTB As System.Windows.Forms.Label
        Private _lblBD As System.Windows.Forms.Label
        Private _btnOK As System.Windows.Forms.Button
        Private _btnCancel As System.Windows.Forms.Button

        Public Sub New()
            InitializeUI()
            LoadDefinitions()
            LoadSettings()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Drawing Tool — Chọn chức năng"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Font = New Drw.Font("Segoe UI", 9)
            Me.ClientSize = New Drw.Size(420, 560)

            Dim grpScope As New System.Windows.Forms.GroupBox With {
                .Text = "Phạm vi áp dụng",
                .Location = New Drw.Point(15, 15),
                .Size = New Drw.Size(390, 120)
            }
            _rdoAll = New System.Windows.Forms.RadioButton With {
                .Text = "Tất cả sheet",
                .Location = New Drw.Point(15, 25),
                .AutoSize = True,
                .Checked = True
            }
            _rdoActive = New System.Windows.Forms.RadioButton With {
                .Text = "Sheet đang xem",
                .Location = New Drw.Point(15, 50),
                .AutoSize = True
            }
            _rdoSelected = New System.Windows.Forms.RadioButton With {
                .Text = "Tự chọn sheet (hiện hộp thoại)",
                .Location = New Drw.Point(15, 75),
                .AutoSize = True
            }
            grpScope.Controls.Add(_rdoAll)
            grpScope.Controls.Add(_rdoActive)
            grpScope.Controls.Add(_rdoSelected)
            Me.Controls.Add(grpScope)

            Dim grpAction As New System.Windows.Forms.GroupBox With {
                .Text = "Hành động",
                .Location = New Drw.Point(15, 145),
                .Size = New Drw.Size(390, 185)
            }
            _rdoDelTB = New System.Windows.Forms.RadioButton With {
                .Text = "Xóa Title Block",
                .Location = New Drw.Point(15, 25),
                .AutoSize = True,
                .Checked = True
            }
            _rdoDelBD = New System.Windows.Forms.RadioButton With {
                .Text = "Xóa Border",
                .Location = New Drw.Point(15, 50),
                .AutoSize = True
            }
            _rdoRepTB = New System.Windows.Forms.RadioButton With {
                .Text = "Thay Title Block",
                .Location = New Drw.Point(15, 75),
                .AutoSize = True
            }
            _rdoRepBD = New System.Windows.Forms.RadioButton With {
                .Text = "Thay Border",
                .Location = New Drw.Point(15, 100),
                .AutoSize = True
            }
            _rdoRepBoth = New System.Windows.Forms.RadioButton With {
                .Text = "Thay cả Title Block + Border",
                .Location = New Drw.Point(15, 125),
                .AutoSize = True
            }

            AddHandler _rdoDelTB.CheckedChanged, AddressOf OnActionChanged
            AddHandler _rdoDelBD.CheckedChanged, AddressOf OnActionChanged
            AddHandler _rdoRepTB.CheckedChanged, AddressOf OnActionChanged
            AddHandler _rdoRepBD.CheckedChanged, AddressOf OnActionChanged
            AddHandler _rdoRepBoth.CheckedChanged, AddressOf OnActionChanged

            grpAction.Controls.Add(_rdoDelTB)
            grpAction.Controls.Add(_rdoDelBD)
            grpAction.Controls.Add(_rdoRepTB)
            grpAction.Controls.Add(_rdoRepBD)
            grpAction.Controls.Add(_rdoRepBoth)
            Me.Controls.Add(grpAction)

            Dim grpDef As New System.Windows.Forms.GroupBox With {
                .Text = "Chọn mẫu có trong file",
                .Location = New Drw.Point(15, 340),
                .Size = New Drw.Size(390, 130)
            }
            _lblTB = New System.Windows.Forms.Label With {
                .Text = "Title Block:",
                .Location = New Drw.Point(15, 30),
                .AutoSize = True
            }
            _cboTitleBlock = New System.Windows.Forms.ComboBox With {
                .Location = New Drw.Point(110, 27),
                .Size = New Drw.Size(260, 24),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }
            _lblBD = New System.Windows.Forms.Label With {
                .Text = "Border:",
                .Location = New Drw.Point(15, 70),
                .AutoSize = True
            }
            _cboBorder = New System.Windows.Forms.ComboBox With {
                .Location = New Drw.Point(110, 67),
                .Size = New Drw.Size(260, 24),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }
            grpDef.Controls.Add(_lblTB)
            grpDef.Controls.Add(_cboTitleBlock)
            grpDef.Controls.Add(_lblBD)
            grpDef.Controls.Add(_cboBorder)
            Me.Controls.Add(grpDef)

            Dim pnlBottom As New System.Windows.Forms.FlowLayoutPanel With {
                .Dock = System.Windows.Forms.DockStyle.Bottom,
                .Height = 55,
                .FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                .Padding = New System.Windows.Forms.Padding(10, 12, 15, 12),
                .WrapContents = False
            }
            _btnCancel = New System.Windows.Forms.Button With {
                .Text = "Đóng",
                .Size = New Drw.Size(90, 32),
                .Margin = New System.Windows.Forms.Padding(5, 0, 0, 0),
                .DialogResult = System.Windows.Forms.DialogResult.Cancel
            }
            _btnOK = New System.Windows.Forms.Button With {
                .Text = "Thực hiện",
                .Size = New Drw.Size(90, 32),
                .Margin = New System.Windows.Forms.Padding(5, 0, 0, 0)
            }
            AddHandler _btnOK.Click, AddressOf HandleOKClick

            pnlBottom.Controls.Add(_btnCancel)
            pnlBottom.Controls.Add(_btnOK)
            Me.Controls.Add(pnlBottom)

            Me.CancelButton = _btnCancel
            Me.AcceptButton = _btnOK

            OnActionChanged(Nothing, EventArgs.Empty)
        End Sub

        Private Sub LoadDefinitions()
            Try
                _cboTitleBlock.Items.Clear()
                For Each n In Drawing_TitleBlock_Border.GetTitleBlockNames()
                    _cboTitleBlock.Items.Add(n)
                Next
                If _cboTitleBlock.Items.Count > 0 Then _cboTitleBlock.SelectedIndex = 0

                _cboBorder.Items.Clear()
                For Each n In Drawing_TitleBlock_Border.GetBorderNames()
                    _cboBorder.Items.Add(n)
                Next
                If _cboBorder.Items.Count > 0 Then _cboBorder.SelectedIndex = 0
            Catch
            End Try
        End Sub

        Private Sub OnActionChanged(sender As Object, e As EventArgs)
            Dim needTB As Boolean = _rdoRepTB.Checked OrElse _rdoRepBoth.Checked
            Dim needBD As Boolean = _rdoRepBD.Checked OrElse _rdoRepBoth.Checked

            _cboTitleBlock.Enabled = needTB
            _cboBorder.Enabled = needBD
            _lblTB.Enabled = needTB
            _lblBD.Enabled = needBD
        End Sub

        Private Sub HandleOKClick(sender As Object, e As EventArgs)
            Dim scope As String = GetScope()
            Dim action As String = GetAction()

            Dim tbName As String = ""
            Dim bdName As String = ""
            If _cboTitleBlock.SelectedItem IsNot Nothing Then tbName = _cboTitleBlock.SelectedItem.ToString()
            If _cboBorder.SelectedItem IsNot Nothing Then bdName = _cboBorder.SelectedItem.ToString()

            If (action = "REP_TB" OrElse action = "REP_BOTH") AndAlso String.IsNullOrEmpty(tbName) Then
                MessageBox.Show("Chưa chọn Title Block mẫu!", "Drawing Tool")
                Return
            End If
            If (action = "REP_BD" OrElse action = "REP_BOTH") AndAlso String.IsNullOrEmpty(bdName) Then
                MessageBox.Show("Chưa chọn Border mẫu!", "Drawing Tool")
                Return
            End If

            SaveSettings(scope, action, tbName, bdName)

            _btnOK.Enabled = False
            _btnCancel.Enabled = False

            Try
                ExecuteAction(scope, action, tbName, bdName)
            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message, "Drawing Tool")
            End Try

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End Sub

        Private Function GetScope() As String
            If _rdoAll.Checked Then Return "ALL"
            If _rdoActive.Checked Then Return "ACTIVE"
            If _rdoSelected.Checked Then Return "SELECTED"
            Return "ALL"
        End Function

        Private Function GetAction() As String
            If _rdoDelTB.Checked Then Return "DEL_TB"
            If _rdoDelBD.Checked Then Return "DEL_BD"
            If _rdoRepTB.Checked Then Return "REP_TB"
            If _rdoRepBD.Checked Then Return "REP_BD"
            If _rdoRepBoth.Checked Then Return "REP_BOTH"
            Return "DEL_TB"
        End Function

        Private Sub ExecuteAction(scope As String, action As String, tbName As String, bdName As String)
            Select Case scope & "|" & action
                Case "ALL|DEL_TB" : Drawing_TitleBlock_Border.DeleteAllTitleBlocks()
                Case "ALL|DEL_BD" : Drawing_TitleBlock_Border.DeleteAllBorders()
                Case "ALL|REP_TB" : Drawing_TitleBlock_Border.ReplaceAllTitleBlocks(tbName)
                Case "ALL|REP_BD" : Drawing_TitleBlock_Border.ReplaceAllBorders(bdName)
                Case "ALL|REP_BOTH" : Drawing_TitleBlock_Border.ReplaceBothAllSheets(tbName, bdName)

                Case "ACTIVE|DEL_TB" : Drawing_TitleBlock_Border.DeleteActiveTitleBlock()
                Case "ACTIVE|DEL_BD" : Drawing_TitleBlock_Border.DeleteActiveBorder()
                Case "ACTIVE|REP_TB" : Drawing_TitleBlock_Border.ReplaceActiveTitleBlock(tbName)
                Case "ACTIVE|REP_BD" : Drawing_TitleBlock_Border.ReplaceActiveBorder(bdName)
                Case "ACTIVE|REP_BOTH" : Drawing_TitleBlock_Border.ReplaceBothActiveSheet(tbName, bdName)

                Case "SELECTED|DEL_TB" : Drawing_TitleBlock_Border.DeleteSelectedTitleBlock()
                Case "SELECTED|DEL_BD" : Drawing_TitleBlock_Border.DeleteSelectedBorder()
                Case "SELECTED|REP_TB" : Drawing_TitleBlock_Border.ReplaceSelectedTitleBlock(tbName)
                Case "SELECTED|REP_BD" : Drawing_TitleBlock_Border.ReplaceSelectedBorder(bdName)
                Case "SELECTED|REP_BOTH" : Drawing_TitleBlock_Border.ReplaceBothSelectedSheet(tbName, bdName)
            End Select
        End Sub

        Private Function GetConfigPath() As String
            Dim appData As String = System.Environment.GetFolderPath(
                System.Environment.SpecialFolder.ApplicationData)
            Dim dir As String = System.IO.Path.Combine(appData, "ToolInventor2025")
            If Not System.IO.Directory.Exists(dir) Then System.IO.Directory.CreateDirectory(dir)
            Return System.IO.Path.Combine(dir, "DrawingTool.cfg")
        End Function

        Private Sub SaveSettings(scope As String, action As String, tbName As String, bdName As String)
            Try
                Dim sb As New System.Text.StringBuilder()
                sb.AppendLine("Scope=" & scope)
                sb.AppendLine("Action=" & action)
                sb.AppendLine("TitleBlock=" & tbName)
                sb.AppendLine("Border=" & bdName)
                System.IO.File.WriteAllText(GetConfigPath(), sb.ToString())
            Catch
            End Try
        End Sub

        Private Sub LoadSettings()
            Try
                Dim path As String = GetConfigPath()
                If Not System.IO.File.Exists(path) Then Return

                Dim lines = System.IO.File.ReadAllLines(path)
                For Each line In lines
                    If String.IsNullOrWhiteSpace(line) Then Continue For
                    Dim idx As Integer = line.IndexOf("="c)
                    If idx < 0 Then Continue For

                    Dim key As String = line.Substring(0, idx).Trim()
                    Dim val As String = line.Substring(idx + 1).Trim()

                    Select Case key
                        Case "Scope"
                            _rdoAll.Checked = (val = "ALL")
                            _rdoActive.Checked = (val = "ACTIVE")
                            _rdoSelected.Checked = (val = "SELECTED")

                        Case "Action"
                            _rdoDelTB.Checked = (val = "DEL_TB")
                            _rdoDelBD.Checked = (val = "DEL_BD")
                            _rdoRepTB.Checked = (val = "REP_TB")
                            _rdoRepBD.Checked = (val = "REP_BD")
                            _rdoRepBoth.Checked = (val = "REP_BOTH")

                        Case "TitleBlock"
                            If _cboTitleBlock.Items.Contains(val) Then _cboTitleBlock.SelectedItem = val

                        Case "Border"
                            If _cboBorder.Items.Contains(val) Then _cboBorder.SelectedItem = val
                    End Select
                Next

                OnActionChanged(Nothing, EventArgs.Empty)

            Catch
            End Try
        End Sub

    End Class

End Namespace