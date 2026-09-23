Imports System.Windows.Forms
Imports System.Collections.Generic
Imports Inventor

Namespace Drawing.Buttons.DrawSheet2025

    ' ============================================================
    ' MODULE — Đồng bộ PartsList cho Inventor 2025
    ' ============================================================
    Public Module Drawing_PartsList_Sync_2025

        ' =====================================================
        ' ENTRY POINT
        ' =====================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                Dim invApp = GetInventorApp()
                If invApp Is Nothing Then
                    MessageBox.Show("Không lấy được Inventor Application!", "PartsList Sync 2025")
                    Return
                End If
                If invApp.ActiveDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Mở file Drawing (.idw/.dwg) trước.", "PartsList Sync 2025")
                    Return
                End If

                Dim frm As New Form_SyncColumns2025()
                frm.ShowDialog()
                frm.Dispose()
            Catch ex As Exception
                Try
                    MessageBox.Show("Lỗi: " & ex.Message, "PartsList Sync 2025")
                Catch
                End Try
            End Try
        End Sub

        ' =====================================================
        ' CLASS CHỨA THÔNG TIN CỘT
        ' =====================================================
        Public Class ColumnInfo
            Public Property Title As String = ""
            Public Property PropTypeRaw As Object = Nothing
            Public Property PropSet As String = ""
            Public Property PropName As String = ""
            Public Property PropId As Object = Nothing
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

                    Try : info.PropTypeRaw = col.PropertyType : Catch : End Try
                    Try : info.PropSet = col.PropertySetName : Catch : End Try
                    Try : info.PropName = col.PropertyName : Catch : End Try
                    Try : info.Title = col.Title : Catch : End Try
                    Try : info.Width = col.Width : Catch : End Try

                    ' Lấy PropId cho cột kFileProperty
                    Try
                        If CType(info.PropTypeRaw, Integer) = 5 Then ' kFileProperty = 5
                            ' info.PropId = col.GetFilePropertyId'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''xóa note
                        End If
                    Catch
                    End Try

                    result.Add(info)
                Next
            Catch
            End Try
            Return result
        End Function

        ' =====================================================
        ' ⭐ SYNC CỘT — bản 2025
        ' =====================================================
        Public Sub SyncColumns(srcSheet As String, srcPLIdx As Integer,
                                dstSheet As String, dstPLIdx As Integer)
            Dim log As New System.Text.StringBuilder()
            Try
                Dim invApp = GetInventorApp()
                If invApp Is Nothing Then Return

                Dim drawDoc As Inventor.DrawingDocument = CType(invApp.ActiveDocument, Inventor.DrawingDocument)

                log.AppendLine("=== SYNC COLUMNS 2025 ===")

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

                ' ============ BƯỚC 1: XÓA HẾT CỘT CŨ ============
                log.AppendLine("")
                log.AppendLine("--- BƯỚC 1: XÓA CỘT CŨ ---")
                Dim removed As Integer = 0

                ' Xóa từ cuối lên để tránh lệch index
                For i As Integer = dstPL.PartsListColumns.Count To 1 Step -1
                    Try
                        dstPL.PartsListColumns.Item(i).Remove()
                        removed += 1
                    Catch ex As Exception
                        log.AppendLine("  ✗ Xóa cột #" & i & " LỖI: " & ex.Message)
                    End Try
                Next

                log.AppendLine("  Đã xóa: " & removed & " cột")
                Try : drawDoc.Update() : Catch : End Try

                ' ============ BƯỚC 2: THÊM CỘT MỚI TỪ NGUỒN ============
                log.AppendLine("")
                log.AppendLine("--- BƯỚC 2: THÊM CỘT MỚI ---")
                Dim added As Integer = 0
                Dim failed As New List(Of String)

                For Each info In srcColumns
                    Try
                        ' Bỏ qua cột không có PropType
                        If info.PropTypeRaw Is Nothing Then
                            log.AppendLine("  ⚠ Bỏ qua (PropType NULL): '" & info.Title & "'")
                            failed.Add(info.Title)
                            Continue For
                        End If

                        Dim propType As Inventor.PropertyTypeEnum =
                            CType(CInt(info.PropTypeRaw), Inventor.PropertyTypeEnum)

                        ' Chuẩn bị tham số
                        Dim propSet As String = info.PropSet
                        Dim propIdentifier As Object = info.PropName

                        ' ✅ Nếu là kFileProperty và có PropId → dùng PropId
                        If propType = Inventor.PropertyTypeEnum.kFileProperty AndAlso info.PropId IsNot Nothing Then
                            propIdentifier = info.PropId
                        End If

                        ' ✅ Nếu PropName rỗng → thử dùng Title
                        If String.IsNullOrEmpty(propIdentifier.ToString()) Then
                            propIdentifier = info.Title
                        End If

                        ' Gọi Add với đúng signature 2025
                        Dim newCol As Inventor.PartsListColumn = dstPL.PartsListColumns.Add(
                            propType,
                            propSet,
                            propIdentifier,
                            0,      ' TargetIndex = 0 → thêm vào cuối
                            True)   ' InsertBefore (bỏ qua vì TargetIndex = 0)

                        ' Set Title và Width
                        If newCol IsNot Nothing Then
                            Try : newCol.Title = info.Title : Catch : End Try
                            Try
                                If info.Width > 0 Then newCol.Width = info.Width
                            Catch
                            End Try

                            added += 1
                            log.AppendLine("  ✓ Add OK: '" & info.Title & "'")
                        Else
                            failed.Add(info.Title)
                            log.AppendLine("  ✗ Add FAIL (null): '" & info.Title & "'")
                        End If

                    Catch ex As Exception
                        failed.Add(info.Title)
                        log.AppendLine("  ✗ Add LỖI: '" & info.Title & "' — " & ex.Message)
                    End Try
                Next

                Try : drawDoc.Update() : Catch : End Try

                ' ============ BƯỚC 3: XỬ LÝ CỘT THẤT BẠI ============
                If failed.Count > 0 Then
                    log.AppendLine("")
                    log.AppendLine("--- BƯỚC 3: CỘT KHÔNG THỂ TÁI TẠO ---")
                    For Each f In failed
                        log.AppendLine("  ⚠ '" & f & "' — cần thêm thủ công qua Column Chooser")
                    Next
                End If

                ' Restore sheet
                Try
                    If originalSheet IsNot Nothing Then originalSheet.Activate()
                Catch
                End Try
                invApp.SilentOperation = False
                Try : drawDoc.Update2(True) : Catch : End Try

                log.AppendLine("")
                log.AppendLine("KẾT QUẢ: đã thêm " & added & " / " & srcColumns.Count & " cột")
                log.AppendLine("Cột đích sau: " & dstPL.PartsListColumns.Count)

            Catch ex As Exception
                log.AppendLine("❌ LỖI TỔNG: " & ex.Message)
                log.AppendLine(ex.StackTrace)
            End Try

            ' Ghi log ra Desktop
            Try
                Dim logPath As String = System.IO.Path.Combine(
                    System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop),
                    "PartsListSync2025_Log.txt")
                System.IO.File.WriteAllText(logPath, log.ToString())
                MessageBox.Show(log.ToString(), "Sync Result 2025")
            Catch
                MessageBox.Show(log.ToString(), "Sync Result 2025")
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

                srcPL.CopyTo(dst)
                Try : drawDoc.Update() : Catch : End Try

                Try
                    If originalSheet IsNot Nothing Then originalSheet.Activate()
                Catch
                End Try
                invApp.SilentOperation = False
                Try : drawDoc.Update2(True) : Catch : End Try

                MessageBox.Show("Đã copy PartsList sang sheet '" & dstSheet & "'.", "Copy PartsList 2025")
            Catch ex As Exception
                MessageBox.Show("Lỗi copy: " & ex.Message, "Copy PartsList 2025")
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
    ' FORM — Chọn nguồn/đích + mode (2025)
    ' ============================================================
    Public Class Form_SyncColumns2025
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
            LoadSettings()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "PartsList Sync 2025 — Đồng bộ cột"
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
                .Text = "Đồng bộ CỘT (xóa cột cũ, thêm cột mới từ nguồn)",
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
                .Text = "Sheet:", .Location = New System.Drawing.Point(15, 25), .AutoSize = True
            }
            _cboSrcSheet = New System.Windows.Forms.ComboBox With {
                .Location = New System.Drawing.Point(90, 22),
                .Size = New System.Drawing.Size(340, 24),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }
            AddHandler _cboSrcSheet.SelectedIndexChanged, AddressOf OnSrcSheetChanged

            Dim lblSrc2 As New System.Windows.Forms.Label With {
                .Text = "PartsList:", .Location = New System.Drawing.Point(15, 60), .AutoSize = True
            }
            _cboSrcPL = New System.Windows.Forms.ComboBox With {
                .Location = New System.Drawing.Point(90, 57),
                .Size = New System.Drawing.Size(340, 24),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }
            AddHandler _cboSrcPL.SelectedIndexChanged, AddressOf OnSrcPLChanged

            grpSrc.Controls.Add(lblSrc1) : grpSrc.Controls.Add(_cboSrcSheet)
            grpSrc.Controls.Add(lblSrc2) : grpSrc.Controls.Add(_cboSrcPL)
            Me.Controls.Add(grpSrc)
            y += 130

            ' ===== ĐÍCH =====
            Dim grpDst As New System.Windows.Forms.GroupBox With {
                .Text = "ĐÍCH (nơi áp dụng)",
                .Location = New System.Drawing.Point(15, y),
                .Size = New System.Drawing.Size(450, 120)
            }
            Dim lblDst1 As New System.Windows.Forms.Label With {
                .Text = "Sheet:", .Location = New System.Drawing.Point(15, 25), .AutoSize = True
            }
            _cboDstSheet = New System.Windows.Forms.ComboBox With {
                .Location = New System.Drawing.Point(90, 22),
                .Size = New System.Drawing.Size(340, 24),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }
            AddHandler _cboDstSheet.SelectedIndexChanged, AddressOf OnDstSheetChanged

            _lblDstPL = New System.Windows.Forms.Label With {
                .Text = "PartsList:", .Location = New System.Drawing.Point(15, 60), .AutoSize = True
            }
            _cboDstPL = New System.Windows.Forms.ComboBox With {
                .Location = New System.Drawing.Point(90, 57),
                .Size = New System.Drawing.Size(340, 24),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }

            grpDst.Controls.Add(lblDst1) : grpDst.Controls.Add(_cboDstSheet)
            grpDst.Controls.Add(_lblDstPL) : grpDst.Controls.Add(_cboDstPL)
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

            ' Trong InitializeUI, sau khi tạo _btnOK/_btnCancel:
            Dim btnReset As New System.Windows.Forms.Button With {
    .Text = "Reset",
    .Size = New System.Drawing.Size(70, 32),
    .Margin = New System.Windows.Forms.Padding(5, 0, 0, 0)
}
            AddHandler btnReset.Click, AddressOf HandleResetClick
            pnlBottom.Controls.Add(btnReset)
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

        Private Sub LoadSheets()
            Try
                _cboSrcSheet.Items.Clear()
                _cboDstSheet.Items.Clear()

                For Each n In Drawing_PartsList_Sync_2025.GetSheetNames()
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

                For Each p In Drawing_PartsList_Sync_2025.GetPartsListsOnSheet(_cboSrcSheet.SelectedItem.ToString())
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
                Dim cols = Drawing_PartsList_Sync_2025.GetColumns(_cboSrcSheet.SelectedItem.ToString(), idx)
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

                For Each p In Drawing_PartsList_Sync_2025.GetPartsListsOnSheet(_cboDstSheet.SelectedItem.ToString())
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
                    MessageBox.Show("Chưa chọn đủ thông tin!", "PartsList Sync 2025")
                    Return
                End If

                Dim srcSheet As String = _cboSrcSheet.SelectedItem.ToString()
                Dim dstSheet As String = _cboDstSheet.SelectedItem.ToString()
                Dim srcIdx As Integer = ParsePlIndex(_cboSrcPL.SelectedItem.ToString())

                Dim dstIdx As Integer = 0
                If _cboDstPL.SelectedItem IsNot Nothing Then
                    dstIdx = ParsePlIndex(_cboDstPL.SelectedItem.ToString())
                End If

                Dim isSync As Boolean = _rdoSync.Checked

                ' ✅ Lưu lựa chọn trước khi chạy
                SaveSettings(isSync, srcSheet, srcIdx, dstSheet, dstIdx)

                _btnOK.Enabled = False
                _btnCancel.Enabled = False

                If isSync Then
                    If _cboDstPL.SelectedItem Is Nothing Then
                        MessageBox.Show("Chưa chọn PartsList đích!", "PartsList Sync 2025")
                        _btnOK.Enabled = True
                        _btnCancel.Enabled = True
                        Return
                    End If
                    Drawing_PartsList_Sync_2025.SyncColumns(srcSheet, srcIdx, dstSheet, dstIdx)
                Else
                    Drawing_PartsList_Sync_2025.CopyPartsList(srcSheet, srcIdx, dstSheet)
                End If

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message, "PartsList Sync 2025")
            End Try

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End Sub

        Private Sub InitializeComponent()
            Me.SuspendLayout()
            '
            'Form_SyncColumns2025
            '
            Me.ClientSize = New System.Drawing.Size(284, 261)
            Me.Name = "Form_SyncColumns2025"
            Me.ResumeLayout(False)

        End Sub

        Private Sub Form_SyncColumns2025_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        End Sub


        ' =====================================================
        ' NHỚ LỆNH — lưu vào %APPDATA%\ToolInventor2025\PartsListSync2025.cfg
        ' =====================================================
        Private Function GetConfigPath() As String
            Try
                Dim appData As String = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.ApplicationData)
                Dim dir As String = System.IO.Path.Combine(appData, "ToolInventor2025")
                If Not System.IO.Directory.Exists(dir) Then System.IO.Directory.CreateDirectory(dir)
                Return System.IO.Path.Combine(dir, "PartsListSync2025.cfg")
            Catch
                ' Fallback: thư mục tạm
                Return System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PartsListSync2025.cfg")
            End Try
        End Function

        Private Sub SaveSettings(isSync As Boolean, srcSheet As String, srcIdx As Integer,
                                  dstSheet As String, dstIdx As Integer)
            Try
                Dim sb As New System.Text.StringBuilder()
                sb.AppendLine("Mode=" & If(isSync, "SYNC", "COPY"))
                sb.AppendLine("SourceSheet=" & srcSheet)
                sb.AppendLine("SourcePLIndex=" & srcIdx)
                sb.AppendLine("TargetSheet=" & dstSheet)
                sb.AppendLine("TargetPLIndex=" & dstIdx)
                System.IO.File.WriteAllText(GetConfigPath(), sb.ToString())
            Catch
            End Try
        End Sub

        Private Sub LoadSettings()
            Try
                Dim path As String = GetConfigPath()
                If Not System.IO.File.Exists(path) Then Return

                Dim mode As String = ""
                Dim srcSheet As String = ""
                Dim srcIdx As Integer = 0
                Dim dstSheet As String = ""
                Dim dstIdx As Integer = 0

                ' --- Đọc file ---
                For Each line In System.IO.File.ReadAllLines(path)
                    If String.IsNullOrWhiteSpace(line) Then Continue For
                    Dim idx As Integer = line.IndexOf("="c)
                    If idx < 0 Then Continue For

                    Dim key As String = line.Substring(0, idx).Trim()
                    Dim val As String = line.Substring(idx + 1).Trim()

                    Select Case key
                        Case "Mode" : mode = val
                        Case "SourceSheet" : srcSheet = val
                        Case "SourcePLIndex" : Integer.TryParse(val, srcIdx)
                        Case "TargetSheet" : dstSheet = val
                        Case "TargetPLIndex" : Integer.TryParse(val, dstIdx)
                    End Select
                Next

                ' --- Áp mode ---
                If mode = "COPY" Then
                    _rdoCopy.Checked = True
                Else
                    _rdoSync.Checked = True
                End If
                UpdateModeUI()

                ' --- Áp sheet nguồn ---
                If Not String.IsNullOrEmpty(srcSheet) AndAlso _cboSrcSheet.Items.Contains(srcSheet) Then
                    _cboSrcSheet.SelectedItem = srcSheet
                    ' ✅ Sau khi chọn sheet nguồn → PartsList tự load qua event

                    ' Chờ event load PL xong rồi chọn PL tương ứng
                    System.Windows.Forms.Application.DoEvents()

                    SelectPLByIndex(_cboSrcPL, srcIdx)
                End If

                ' --- Áp sheet đích ---
                If Not String.IsNullOrEmpty(dstSheet) AndAlso _cboDstSheet.Items.Contains(dstSheet) Then
                    _cboDstSheet.SelectedItem = dstSheet
                    System.Windows.Forms.Application.DoEvents()

                    SelectPLByIndex(_cboDstPL, dstIdx)
                End If

            Catch
            End Try
        End Sub

        ' Chọn item trong ComboBox theo index PartsList ("1|...", "2|...")
        Private Sub SelectPLByIndex(cbo As System.Windows.Forms.ComboBox, plIdx As Integer)
            Try
                If plIdx <= 0 Then Return
                For Each item In cbo.Items
                    Dim s As String = item.ToString()
                    Dim i As Integer = s.IndexOf("|"c)
                    If i > 0 Then
                        Dim n As Integer = 0
                        If Integer.TryParse(s.Substring(0, i), n) AndAlso n = plIdx Then
                            cbo.SelectedItem = item
                            Return
                        End If
                    End If
                Next
            Catch
            End Try
        End Sub
        Private Sub HandleResetClick(sender As Object, e As EventArgs)
            Try
                System.IO.File.Delete(GetConfigPath())
                MessageBox.Show("Đã xóa cấu hình. Mở lại form để về mặc định.", "PartsList Sync 2025")
            Catch
                MessageBox.Show("Không xóa được file cấu hình.", "PartsList Sync 2025")
            End Try
        End Sub
    End Class

End Namespace