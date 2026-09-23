Imports System.Windows.Forms
Imports System.Collections.Generic
Imports Inventor

Namespace ToolInventor2025.Drawing.Buttons.DrawSheet

    ' ============================================================
    ' MODULE — Đồng bộ / Copy PartsList
    ' ============================================================
    Public Module Drawing_PartsList_Sync2020

        ' =====================================================
        ' ENTRY POINT
        ' =====================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                Dim invApp = GetInventorApp()
                If invApp Is Nothing Then
                    MessageBox.Show("Không lấy được Inventor Application!", "PartsList Sync")
                    Return
                End If
                If invApp.ActiveDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Mở file Drawing (.idw/.dwg) trước.", "PartsList Sync")
                    Return
                End If

                Dim frm As New Form_SyncColumns()
                frm.ShowDialog()
                frm.Dispose()
            Catch ex As Exception
                Try
                    MessageBox.Show("Lỗi: " & ex.Message, "PartsList Sync")
                Catch
                End Try
            End Try
        End Sub

        ' =====================================================
        ' CLASS CHỨA THÔNG TIN CỘT
        ' =====================================================
        Public Class ColumnInfo
            Public Property Title As String = ""
            Public Property PropTypeRaw As Object = Nothing   ' giữ nguyên giá trị gốc
            Public Property PropSet As String = ""
            Public Property PropName As String = ""
            Public Property Width As Double = 0
        End Class

        ' =====================================================
        ' LẤY DANH SÁCH SHEET
        ' =====================================================
        Public Function GetSheetNames() As List(Of String)
            Dim result As New List(Of String)
            Try
                Dim invApp = GetInventorApp()
                If invApp Is Nothing Then Return result
                If invApp.ActiveDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then Return result

                Dim drawDoc As Inventor.DrawingDocument = CType(invApp.ActiveDocument, Inventor.DrawingDocument)
                For Each s As Inventor.Sheet In drawDoc.Sheets
                    result.Add(s.Name)
                Next
            Catch
            End Try
            Return result
        End Function

        ' =====================================================
        ' LẤY PARTSLIST TRÊN SHEET — trả về "Index|Label"
        ' =====================================================
        Public Function GetPartsListsOnSheet(sheetName As String) As List(Of String)
            Dim result As New List(Of String)
            Try
                Dim invApp = GetInventorApp()
                If invApp Is Nothing Then Return result
                If invApp.ActiveDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then Return result

                Dim drawDoc As Inventor.DrawingDocument = CType(invApp.ActiveDocument, Inventor.DrawingDocument)
                Dim oSheet As Inventor.Sheet = drawDoc.Sheets.Item(sheetName)

                Dim idx As Integer = 0
                For Each pl As Inventor.PartsList In oSheet.PartsLists
                    idx += 1
                    Dim cnt As Integer = 0
                    Try
                        cnt = pl.PartsListColumns.Count
                    Catch
                    End Try
                    result.Add(idx & "|PL #" & idx & " (" & cnt & " cột)")
                Next
            Catch
            End Try
            Return result
        End Function

        ' =====================================================
        ' LẤY CỘT CỦA 1 PARTSLIST
        ' =====================================================
        Public Function GetColumns(sheetName As String, plIdx As Integer) As List(Of ColumnInfo)
            Dim result As New List(Of ColumnInfo)
            Try
                Dim invApp = GetInventorApp()
                If invApp Is Nothing Then Return result

                Dim drawDoc As Inventor.DrawingDocument = CType(invApp.ActiveDocument, Inventor.DrawingDocument)
                Dim oSheet As Inventor.Sheet = drawDoc.Sheets.Item(sheetName)
                Dim pl As Inventor.PartsList = oSheet.PartsLists.Item(plIdx)

                For Each col As Inventor.PartsListColumn In pl.PartsListColumns
                    Dim info As New ColumnInfo()

                    ' ✅ Lấy PropType thô — không cần biết enum name
                    Try
                        info.PropTypeRaw = col.PropertyType
                    Catch
                        info.PropTypeRaw = Nothing
                    End Try

                    Try
                        info.PropSet = col.PropertySetName
                    Catch
                        info.PropSet = ""
                    End Try

                    Try
                        info.PropName = col.PropertyName
                    Catch
                        info.PropName = ""
                    End Try

                    Try
                        info.Title = col.Title
                    Catch
                        info.Title = ""
                    End Try

                    Try
                        info.Width = col.Width
                    Catch
                        info.Width = 0
                    End Try

                    result.Add(info)
                Next
            Catch
            End Try
            Return result
        End Function

        ' =====================================================
        ' ⭐ SYNC CỘT — giữ data đích, copy header nguồn
        ' =====================================================
        Public Sub SyncColumns(srcSheet As String, srcPLIdx As Integer,
                        dstSheet As String, dstPLIdx As Integer)
            Dim log As New System.Text.StringBuilder()
            Try
                Dim invApp = GetInventorApp()
                If invApp Is Nothing Then Return

                Dim drawDoc As Inventor.DrawingDocument = CType(invApp.ActiveDocument, Inventor.DrawingDocument)

                log.AppendLine("=== SYNC COLUMNS v2 ===")

                ' 1. Đọc cột nguồn
                Dim srcColumns As List(Of ColumnInfo) = GetColumns(srcSheet, srcPLIdx)
                log.AppendLine("Cột nguồn: " & srcColumns.Count)

                ' 2. Lấy PartsList đích
                Dim oDstSheet As Inventor.Sheet = drawDoc.Sheets.Item(dstSheet)
                Dim dstPL As Inventor.PartsList = oDstSheet.PartsLists.Item(dstPLIdx)

                log.AppendLine("Cột đích trước: " & dstPL.PartsListColumns.Count)

                ' 3. Activate sheet đích
                Dim originalSheet As Inventor.Sheet = drawDoc.ActiveSheet
                invApp.SilentOperation = True
                Try : oDstSheet.Activate() : Catch : End Try
                Try : drawDoc.Update() : Catch : End Try

                ' ============ CHIẾN LƯỢC ============
                ' Không xóa cột (API không cho). Thay vào đó:
                '   - Cột đích có sẵn → sửa Title cho giống nguồn
                '   - Cột nguồn chưa có → thêm mới
                '   - Cột đích thừa → ẩn (nếu có .Visible)

                Dim dstCols As New List(Of Inventor.PartsListColumn)
                For Each col As Inventor.PartsListColumn In dstPL.PartsListColumns
                    dstCols.Add(col)
                Next

                ' --- BƯỚC A: Đổi Title cột có sẵn ---
                log.AppendLine("")
                log.AppendLine("--- BƯỚC A: Đổi Title cột có sẵn ---")
                For Each info In srcColumns
                    For Each dcol In dstCols
                        Dim dTitle As String = ""
                        Try : dTitle = dcol.Title : Catch : End Try

                        If String.Equals(dTitle, info.Title, StringComparison.OrdinalIgnoreCase) Then
                            ' Title đã đúng → không cần sửa
                            Exit For
                        End If
                    Next
                Next

                ' --- BƯỚC B: Thêm cột nguồn chưa có ở đích ---
                log.AppendLine("")
                log.AppendLine("--- BƯỚC B: Thêm cột thiếu ---")
                Dim added As Integer = 0

                For Each info In srcColumns
                    ' Kiểm tra cột đã tồn tại chưa
                    Dim exists As Boolean = False
                    For Each dcol In dstCols
                        Dim dTitle As String = ""
                        Try : dTitle = dcol.Title : Catch : End Try
                        If String.Equals(dTitle, info.Title, StringComparison.OrdinalIgnoreCase) Then
                            exists = True
                            Exit For
                        End If
                    Next

                    If exists Then
                        log.AppendLine("  = Đã có: '" & info.Title & "'")
                        Continue For
                    End If

                    ' Cột chưa có → thêm mới
                    If info.PropTypeRaw Is Nothing Then
                        log.AppendLine("  ⚠ Bỏ qua (PropType NULL): '" & info.Title & "'")
                        Continue For
                    End If

                    ' ✅ Cast enum — dùng CInt trước để tránh nhầm overload
                    Dim propType As Inventor.PropertyTypeEnum
                    Try
                        propType = CType(CInt(info.PropTypeRaw), Inventor.PropertyTypeEnum)
                    Catch ex As Exception
                        log.AppendLine("  ✗ Cast PropType LỖI: " & ex.Message)
                        Continue For
                    End Try

                    ' ✅ Gọi Add qua LATE BINDING để tránh nhầm overload
                    Dim addedOK As Boolean = False
                    Try
                        Dim cols As Object = dstPL.PartsListColumns
                        Dim newCol As Inventor.PartsListColumn =
                    DirectCast(cols.Add(propType, info.PropSet, info.PropName, info.Title),
                               Inventor.PartsListColumn)
                        If newCol IsNot Nothing Then
                            Try
                                If info.Width > 0 Then newCol.Width = info.Width
                            Catch
                            End Try
                            added += 1
                            addedOK = True
                            log.AppendLine("  ✓ Add OK: '" & info.Title & "'")
                        End If
                    Catch ex As Exception
                        log.AppendLine("  ✗ Add LỖI: " & ex.Message)
                    End Try

                    If Not addedOK Then
                        ' Thử signature 3 tham số
                        Try
                            Dim cols As Object = dstPL.PartsListColumns
                            Dim newCol As Inventor.PartsListColumn =
                        DirectCast(cols.Add(propType, info.PropSet, info.PropName),
                                   Inventor.PartsListColumn)
                            If newCol IsNot Nothing Then
                                Try : newCol.Title = info.Title : Catch : End Try
                                added += 1
                                log.AppendLine("  ✓ Add OK (3 args): '" & info.Title & "'")
                            End If
                        Catch ex3 As Exception
                            log.AppendLine("  ✗ Add 3 args cũng LỖI: " & ex3.Message)
                        End Try
                    End If
                Next

                Try : drawDoc.Update() : Catch : End Try

                ' --- BƯỚC C: Ẩn cột thừa (có ở đích, không có ở nguồn) ---
                log.AppendLine("")
                log.AppendLine("--- BƯỚC C: Ẩn cột thừa ---")
                Dim hidden As Integer = 0

                ' Refresh danh sách cột đích
                Dim dstCols2 As New List(Of Inventor.PartsListColumn)
                For Each col As Inventor.PartsListColumn In dstPL.PartsListColumns
                    dstCols2.Add(col)
                Next

                For Each dcol In dstCols2
                    Dim dTitle As String = ""
                    Try : dTitle = dcol.Title : Catch : End Try
                    If String.IsNullOrEmpty(dTitle) Then Continue For

                    ' Có trong nguồn không?
                    Dim inSrc As Boolean = False
                    For Each info In srcColumns
                        If String.Equals(info.Title, dTitle, StringComparison.OrdinalIgnoreCase) Then
                            inSrc = True
                            Exit For
                        End If
                    Next

                    If Not inSrc Then
                        ' Cột thừa → thử ẩn
                        Try
                            dcol.Visible = False
                            hidden += 1
                            log.AppendLine("  ✓ Ẩn OK: '" & dTitle & "'")
                        Catch ex As Exception
                            log.AppendLine("  ✗ Không ẩn được '" & dTitle & "': " & ex.Message)
                        End Try
                    End If
                Next

                Try : drawDoc.Update() : Catch : End Try

                ' Restore sheet
                Try
                    If originalSheet IsNot Nothing Then originalSheet.Activate()
                Catch
                End Try
                invApp.SilentOperation = False
                Try : drawDoc.Update2(True) : Catch : End Try

                log.AppendLine("")
                log.AppendLine("KẾT QUẢ: thêm " & added & " cột, ẩn " & hidden & " cột")
                log.AppendLine("Cột đích sau: " & dstPL.PartsListColumns.Count)

            Catch ex As Exception
                log.AppendLine("❌ LỖI TỔNG: " & ex.Message)
                log.AppendLine(ex.StackTrace)
            End Try

            ' Ghi log
            Try
                Dim logPath As String = System.IO.Path.Combine(
            System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop),
            "PartsListSync_Log.txt")
                System.IO.File.WriteAllText(logPath, log.ToString())
                MessageBox.Show(log.ToString(), "Sync Result")
            Catch
                MessageBox.Show(log.ToString(), "Sync Result")
            End Try
        End Sub

        ' =====================================================
        ' COPY NGUYÊN PARTSLIST
        ' =====================================================
        Public Sub CopyPartsList(srcSheet As String, srcPLIdx As Integer, dstSheet As String)
            Try
                Dim invApp = GetInventorApp()
                If invApp Is Nothing Then Return

                Dim drawDoc As Inventor.DrawingDocument = CType(invApp.ActiveDocument, Inventor.DrawingDocument)
                Dim src As Inventor.Sheet = drawDoc.Sheets.Item(srcSheet)
                Dim dst As Inventor.Sheet = drawDoc.Sheets.Item(dstSheet)
                Dim srcPL As Inventor.PartsList = src.PartsLists.Item(srcPLIdx)

                Dim originalSheet As Inventor.Sheet = drawDoc.ActiveSheet
                invApp.SilentOperation = True
                Try : dst.Activate() : Catch : End Try
                Try : drawDoc.Update() : Catch : End Try

                ' ✅ CopyTo chỉ có 1 tham số
                srcPL.CopyTo(dst)
                Try : drawDoc.Update() : Catch : End Try

                Try
                    If originalSheet IsNot Nothing Then originalSheet.Activate()
                Catch
                End Try
                invApp.SilentOperation = False
                Try : drawDoc.Update2(True) : Catch : End Try

                MessageBox.Show("Đã copy PartsList sang sheet '" & dstSheet & "'.", "Copy PartsList")
            Catch ex As Exception
                MessageBox.Show("Lỗi copy: " & ex.Message, "Copy PartsList")
            End Try
        End Sub

        ' =====================================================
        ' HELPER
        ' =====================================================
        Private Function GetInventorApp() As Inventor.Application
            Try
                Return CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch
                Return Nothing
            End Try
        End Function

    End Module


    ' ============================================================
    ' FORM — Chọn nguồn/đích + mode
    ' ============================================================
    Public Class Form_SyncColumns
        Inherits System.Windows.Forms.Form

        Private _rdoSync As System.Windows.Forms.RadioButton
        Private _rdoCopy As System.Windows.Forms.RadioButton
        Private _cboSrcSheet As System.Windows.Forms.ComboBox
        Private _cboSrcPL As System.Windows.Forms.ComboBox
        Private _cboDstSheet As System.Windows.Forms.ComboBox
        Private _cboDstPL As System.Windows.Forms.ComboBox
        Private _lstPreview As System.Windows.Forms.ListBox
        Private _lblDstPL As System.Windows.Forms.Label
        Private _btnOK As System.Windows.Forms.Button
        Private _btnCancel As System.Windows.Forms.Button

        Public Sub New()
            InitializeUI()
            LoadSheets()
            UpdateModeUI()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "PartsList — Đồng bộ cột giữa 2 sheet"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Font = New System.Drawing.Font("Segoe UI", 9)
            Me.ClientSize = New System.Drawing.Size(480, 620)

            Dim y As Integer = 15

            ' ===== MODE =====
            Dim grpMode As New System.Windows.Forms.GroupBox With {
                .Text = "Chế độ",
                .Location = New System.Drawing.Point(15, y),
                .Size = New System.Drawing.Size(450, 70)
            }
            _rdoSync = New System.Windows.Forms.RadioButton With {
                .Text = "Đồng bộ CỘT (giữ data đích, copy header nguồn)",
                .Location = New System.Drawing.Point(15, 20),
                .AutoSize = True,
                .Checked = True
            }
            _rdoCopy = New System.Windows.Forms.RadioButton With {
                .Text = "Copy NGUYÊN PartsList (data + cột từ nguồn)",
                .Location = New System.Drawing.Point(15, 42),
                .AutoSize = True
            }
            AddHandler _rdoSync.CheckedChanged, AddressOf OnModeChanged
            AddHandler _rdoCopy.CheckedChanged, AddressOf OnModeChanged
            grpMode.Controls.Add(_rdoSync)
            grpMode.Controls.Add(_rdoCopy)
            Me.Controls.Add(grpMode)
            y += 80

            ' ===== NGUỒN =====
            Dim grpSrc As New System.Windows.Forms.GroupBox With {
                .Text = "NGUỒN (mẫu cột)",
                .Location = New System.Drawing.Point(15, y),
                .Size = New System.Drawing.Size(450, 120)
            }
            Dim lblSrc1 As New System.Windows.Forms.Label With {
                .Text = "Sheet:",
                .Location = New System.Drawing.Point(15, 25),
                .AutoSize = True
            }
            _cboSrcSheet = New System.Windows.Forms.ComboBox With {
                .Location = New System.Drawing.Point(90, 22),
                .Size = New System.Drawing.Size(340, 24),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }
            AddHandler _cboSrcSheet.SelectedIndexChanged, AddressOf OnSrcSheetChanged

            Dim lblSrc2 As New System.Windows.Forms.Label With {
                .Text = "PartsList:",
                .Location = New System.Drawing.Point(15, 60),
                .AutoSize = True
            }
            _cboSrcPL = New System.Windows.Forms.ComboBox With {
                .Location = New System.Drawing.Point(90, 57),
                .Size = New System.Drawing.Size(340, 24),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }
            AddHandler _cboSrcPL.SelectedIndexChanged, AddressOf OnSrcPLChanged

            grpSrc.Controls.Add(lblSrc1)
            grpSrc.Controls.Add(_cboSrcSheet)
            grpSrc.Controls.Add(lblSrc2)
            grpSrc.Controls.Add(_cboSrcPL)
            Me.Controls.Add(grpSrc)
            y += 130

            ' ===== ĐÍCH =====
            Dim grpDst As New System.Windows.Forms.GroupBox With {
                .Text = "ĐÍCH (nơi áp dụng)",
                .Location = New System.Drawing.Point(15, y),
                .Size = New System.Drawing.Size(450, 120)
            }
            Dim lblDst1 As New System.Windows.Forms.Label With {
                .Text = "Sheet:",
                .Location = New System.Drawing.Point(15, 25),
                .AutoSize = True
            }
            _cboDstSheet = New System.Windows.Forms.ComboBox With {
                .Location = New System.Drawing.Point(90, 22),
                .Size = New System.Drawing.Size(340, 24),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }
            AddHandler _cboDstSheet.SelectedIndexChanged, AddressOf OnDstSheetChanged

            _lblDstPL = New System.Windows.Forms.Label With {
                .Text = "PartsList:",
                .Location = New System.Drawing.Point(15, 60),
                .AutoSize = True
            }
            _cboDstPL = New System.Windows.Forms.ComboBox With {
                .Location = New System.Drawing.Point(90, 57),
                .Size = New System.Drawing.Size(340, 24),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }

            grpDst.Controls.Add(lblDst1)
            grpDst.Controls.Add(_cboDstSheet)
            grpDst.Controls.Add(_lblDstPL)
            grpDst.Controls.Add(_cboDstPL)
            Me.Controls.Add(grpDst)
            y += 130

            ' ===== PREVIEW =====
            Dim lblPrev As New System.Windows.Forms.Label With {
                .Text = "Xem trước cột của NGUỒN:",
                .Location = New System.Drawing.Point(15, y),
                .AutoSize = True,
                .Font = New System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)
            }
            Me.Controls.Add(lblPrev)
            y += 22

            _lstPreview = New System.Windows.Forms.ListBox With {
                .Location = New System.Drawing.Point(15, y),
                .Size = New System.Drawing.Size(450, 150)
            }
            Me.Controls.Add(_lstPreview)

            ' ===== BUTTONS =====
            Dim pnlBottom As New System.Windows.Forms.FlowLayoutPanel With {
                .Dock = System.Windows.Forms.DockStyle.Bottom,
                .Height = 55,
                .FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                .Padding = New System.Windows.Forms.Padding(10, 12, 15, 12),
                .WrapContents = False
            }
            _btnCancel = New System.Windows.Forms.Button With {
                .Text = "Đóng",
                .Size = New System.Drawing.Size(90, 32),
                .Margin = New System.Windows.Forms.Padding(5, 0, 0, 0),
                .DialogResult = System.Windows.Forms.DialogResult.Cancel
            }
            _btnOK = New System.Windows.Forms.Button With {
                .Text = "Áp dụng",
                .Size = New System.Drawing.Size(90, 32),
                .Margin = New System.Windows.Forms.Padding(5, 0, 0, 0)
            }
            AddHandler _btnOK.Click, AddressOf HandleOKClick

            pnlBottom.Controls.Add(_btnCancel)
            pnlBottom.Controls.Add(_btnOK)
            Me.Controls.Add(pnlBottom)

            Me.CancelButton = _btnCancel
            Me.AcceptButton = _btnOK
        End Sub

        Private Sub UpdateModeUI()
            Dim isSync As Boolean = _rdoSync.Checked
            _cboDstPL.Enabled = isSync
            _lblDstPL.Enabled = isSync
        End Sub

        Private Sub OnModeChanged(sender As Object, e As EventArgs)
            UpdateModeUI()
        End Sub

        ' ===== Load danh sách sheet =====
        Private Sub LoadSheets()
            Try
                _cboSrcSheet.Items.Clear()
                _cboDstSheet.Items.Clear()

                For Each n In Drawing_PartsList_Sync2020.GetSheetNames()
                    _cboSrcSheet.Items.Add(n)
                    _cboDstSheet.Items.Add(n)
                Next

                If _cboSrcSheet.Items.Count > 0 Then
                    _cboSrcSheet.SelectedIndex = 0
                End If

                If _cboDstSheet.Items.Count > 1 Then
                    _cboDstSheet.SelectedIndex = 1
                ElseIf _cboDstSheet.Items.Count > 0 Then
                    _cboDstSheet.SelectedIndex = 0
                End If

            Catch
            End Try
        End Sub

        Private Sub OnSrcSheetChanged(sender As Object, e As EventArgs)
            Try
                _cboSrcPL.Items.Clear()
                If _cboSrcSheet.SelectedItem Is Nothing Then Return

                For Each p In Drawing_PartsList_Sync2020.GetPartsListsOnSheet(_cboSrcSheet.SelectedItem.ToString())
                    _cboSrcPL.Items.Add(p)
                Next
                If _cboSrcPL.Items.Count > 0 Then _cboSrcPL.SelectedIndex = 0
            Catch
            End Try
        End Sub

        Private Sub OnSrcPLChanged(sender As Object, e As EventArgs)
            Try
                _lstPreview.Items.Clear()
                If _cboSrcSheet.SelectedItem Is Nothing Then Return
                If _cboSrcPL.SelectedItem Is Nothing Then Return

                Dim idx As Integer = ParsePlIndex(_cboSrcPL.SelectedItem.ToString())
                Dim cols = Drawing_PartsList_Sync2020.GetColumns(_cboSrcSheet.SelectedItem.ToString(), idx)
                For Each c In cols
                    Dim line As String = c.Title
                    If Not String.IsNullOrEmpty(c.PropSet) OrElse Not String.IsNullOrEmpty(c.PropName) Then
                        line &= "   [" & c.PropSet & "." & c.PropName & "]"
                    End If
                    _lstPreview.Items.Add(line)
                Next
            Catch
            End Try
        End Sub

        Private Sub OnDstSheetChanged(sender As Object, e As EventArgs)
            Try
                _cboDstPL.Items.Clear()
                If _cboDstSheet.SelectedItem Is Nothing Then Return

                For Each p In Drawing_PartsList_Sync2020.GetPartsListsOnSheet(_cboDstSheet.SelectedItem.ToString())
                    _cboDstPL.Items.Add(p)
                Next
                If _cboDstPL.Items.Count > 0 Then _cboDstPL.SelectedIndex = 0
            Catch
            End Try
        End Sub

        Private Function ParsePlIndex(s As String) As Integer
            Try
                Dim i As Integer = s.IndexOf("|"c)
                If i > 0 Then Return Integer.Parse(s.Substring(0, i))
            Catch
            End Try
            Return 1
        End Function

        Private Sub HandleOKClick(sender As Object, e As EventArgs)
            Try
                If _cboSrcSheet.SelectedItem Is Nothing OrElse
                   _cboSrcPL.SelectedItem Is Nothing OrElse
                   _cboDstSheet.SelectedItem Is Nothing Then
                    MessageBox.Show("Chưa chọn đủ thông tin!", "PartsList Sync")
                    Return
                End If

                Dim srcSheet As String = _cboSrcSheet.SelectedItem.ToString()
                Dim dstSheet As String = _cboDstSheet.SelectedItem.ToString()
                Dim srcIdx As Integer = ParsePlIndex(_cboSrcPL.SelectedItem.ToString())

                _btnOK.Enabled = False
                _btnCancel.Enabled = False

                If _rdoSync.Checked Then
                    If _cboDstPL.SelectedItem Is Nothing Then
                        MessageBox.Show("Chưa chọn PartsList đích!", "PartsList Sync")
                        _btnOK.Enabled = True
                        _btnCancel.Enabled = True
                        Return
                    End If
                    Dim dstIdx As Integer = ParsePlIndex(_cboDstPL.SelectedItem.ToString())
                    Drawing_PartsList_Sync2020.SyncColumns(srcSheet, srcIdx, dstSheet, dstIdx)
                Else
                    Drawing_PartsList_Sync2020.CopyPartsList(srcSheet, srcIdx, dstSheet)
                End If

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message, "PartsList Sync")
            End Try

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End Sub

    End Class

End Namespace