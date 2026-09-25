Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.DrawSheet

    ' ============================================================
    ' MODULE — Đồng bộ / Copy PartsList (Inventor 2025)
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
            Public Property PropTypeRaw As Object = Nothing
            Public Property PropSet As String = ""
            Public Property PropName As String = ""
            Public Property Width As Double = 0

            ''' <summary>Key để so sánh: PropSet|PropName|Title</summary>
            Public ReadOnly Property Key As String
                Get
                    Return (PropSet & "|" & PropName & "|" & Title).ToUpper()
                End Get
            End Property
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
                    Try : cnt = pl.PartsListColumns.Count : Catch : End Try
                    result.Add(idx & "|PL #" & idx & "  (" & cnt & " cột)")
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
                    result.Add(info)
                Next
            Catch
            End Try
            Return result
        End Function


        ' =====================================================
        ' ⭐ SYNC CỘT — viết gọn, không có bước dư
        '
        ' 1. Lấy cột nguồn + cột đích
        ' 2. Cột nguồn nào chưa có ở đích → Add
        ' 3. Cột đích nào không có ở nguồn → Ẩn (nếu có thể)
        ' =====================================================
        Public Sub SyncColumns(srcSheet As String,
                       srcPLIdx As Integer,
                       dstSheet As String,
                       dstPLIdx As Integer)
            Dim log As New System.Text.StringBuilder()
            Dim invApp As Inventor.Application = Nothing
            Dim originalSheet As Inventor.Sheet = Nothing

            Try
                invApp = GetInventorApp()
                If invApp Is Nothing Then Return

                Dim drawDoc As Inventor.DrawingDocument = CType(invApp.ActiveDocument, Inventor.DrawingDocument)
                originalSheet = drawDoc.ActiveSheet

                Dim srcColumns As List(Of ColumnInfo) = GetColumns(srcSheet, srcPLIdx)
                Dim oDstSheet As Inventor.Sheet = drawDoc.Sheets.Item(dstSheet)
                Dim dstPL As Inventor.PartsList = oDstSheet.PartsLists.Item(dstPLIdx)

                log.AppendLine("Nguồn : " & srcSheet & "  PL#" & srcPLIdx & "  (" & srcColumns.Count & " cột)")
                log.AppendLine("Đích  : " & dstSheet & "  PL#" & dstPLIdx & "  (" & dstPL.PartsListColumns.Count & " cột)")
                log.AppendLine()

                invApp.SilentOperation = True
                Try : oDstSheet.Activate() : Catch : End Try
                Try : drawDoc.Update() : Catch : End Try

                '── Key nguồn ──
                Dim srcKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each info In srcColumns
                    srcKeys.Add(MakeKey(info.PropSet, info.PropName, info.Title))
                Next

                '════════════════════════════════════════
                ' B1. XÓA cột đích không có trong nguồn
                '════════════════════════════════════════
                log.AppendLine("── XÓA CỘT THỪA ──")
                Dim removed As Integer = 0
                For i As Integer = dstPL.PartsListColumns.Count To 1 Step -1
                    Try
                        Dim dcol As Inventor.PartsListColumn = dstPL.PartsListColumns.Item(i)
                        Dim dSet As String = "", dName As String = "", dTitle As String = ""
                        Try : dSet = dcol.PropertySetName : Catch : End Try
                        Try : dName = dcol.PropertyName : Catch : End Try
                        Try : dTitle = dcol.Title : Catch : End Try

                        Dim key As String = MakeKey(dSet, dName, dTitle)
                        If srcKeys.Contains(key) Then Continue For

                        dcol.Remove()
                        removed += 1
                        log.AppendLine("  ✓ Xóa: '" & dTitle & "'")
                    Catch ex As Exception
                        log.AppendLine("  ✗ Xóa lỗi: " & ex.Message)
                    End Try
                Next

                Try : drawDoc.Update() : Catch : End Try

                '════════════════════════════════════════
                ' B2. THÊM cột nguồn còn thiếu
                '════════════════════════════════════════
                log.AppendLine()
                log.AppendLine("── THÊM CỘT ──")
                Dim added As Integer = 0, skipped As Integer = 0, failed As Integer = 0

                ' Key đích sau khi xóa
                Dim dstKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each dcol As Inventor.PartsListColumn In dstPL.PartsListColumns
                    Dim dSet As String = "", dName As String = "", dTitle As String = ""
                    Try : dSet = dcol.PropertySetName : Catch : End Try
                    Try : dName = dcol.PropertyName : Catch : End Try
                    Try : dTitle = dcol.Title : Catch : End Try
                    dstKeys.Add(MakeKey(dSet, dName, dTitle))
                Next

                For Each info In srcColumns
                    Dim key As String = MakeKey(info.PropSet, info.PropName, info.Title)
                    If dstKeys.Contains(key) Then
                        skipped += 1
                        Continue For
                    End If

                    If info.PropTypeRaw Is Nothing Then
                        failed += 1
                        log.AppendLine("  ⚠ PropType NULL: '" & info.Title & "'")
                        Continue For
                    End If

                    Dim propType As Inventor.PropertyTypeEnum
                    Try
                        propType = CType(CInt(info.PropTypeRaw), Inventor.PropertyTypeEnum)
                    Catch
                        failed += 1
                        Continue For
                    End Try

                    If AddColumn(dstPL, propType, info) Then
                        added += 1
                        dstKeys.Add(key)
                        log.AppendLine("  ✓ Thêm: '" & info.Title & "'")
                    Else
                        failed += 1
                        log.AppendLine("  ✗ Thêm lỗi: '" & info.Title & "'")
                    End If
                Next

                Try : drawDoc.Update() : Catch : End Try

                '════════════════════════════════════════
                ' B3. CẬP NHẬT Title + Width cột đã khớp
                '════════════════════════════════════════
                log.AppendLine()
                log.AppendLine("── CẬP NHẬT Title/Width ──")
                Dim updated As Integer = 0
                For Each info In srcColumns
                    For Each dcol As Inventor.PartsListColumn In dstPL.PartsListColumns
                        Try
                            Dim dSet As String = "", dName As String = "", dTitle As String = ""
                            Try : dSet = dcol.PropertySetName : Catch : End Try
                            Try : dName = dcol.PropertyName : Catch : End Try
                            Try : dTitle = dcol.Title : Catch : End Try

                            If MakeKey(dSet, dName, dTitle) <> MakeKey(info.PropSet, info.PropName, info.Title) AndAlso
                       MakeKey(dSet, dName, dTitle) <> MakeKey(info.PropSet, info.PropName, dTitle) Then
                                ' Khớp theo PropSet|PropName
                                If (dSet & "|" & dName).ToUpper() <> (info.PropSet & "|" & info.PropName).ToUpper() Then
                                    Continue For
                                End If
                            End If

                            Dim changed As Boolean = False
                            If Not String.IsNullOrEmpty(info.Title) AndAlso dTitle <> info.Title Then
                                Try : dcol.Title = info.Title : changed = True : Catch : End Try
                            End If
                            If info.Width > 0 Then
                                Try : dcol.Width = info.Width : changed = True : Catch : End Try
                            End If
                            If changed Then updated += 1
                            Exit For
                        Catch
                        End Try
                    Next
                Next

                Try : drawDoc.Update2(True) : Catch : End Try

                log.AppendLine()
                log.AppendLine("══ KẾT QUẢ ══")
                log.AppendLine("  Xóa cột   : " & removed)
                log.AppendLine("  Thêm cột  : " & added)
                log.AppendLine("  Bỏ qua    : " & skipped)
                log.AppendLine("  Cập nhật  : " & updated)
                log.AppendLine("  Lỗi thêm  : " & failed)
                log.AppendLine("  Cột sau   : " & dstPL.PartsListColumns.Count)

            Catch ex As Exception
                log.AppendLine()
                log.AppendLine("❌ LỖI: " & ex.Message)
            Finally
                Try
                    If originalSheet IsNot Nothing Then originalSheet.Activate()
                Catch
                End Try
                Try
                    If invApp IsNot Nothing Then invApp.SilentOperation = False
                Catch
                End Try
            End Try

            MessageBox.Show(log.ToString(), "Sync Result", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Function MakeKey(propSet As String, propName As String, title As String) As String
            Dim ps As String = If(propSet, "").Trim()
            Dim pn As String = If(propName, "").Trim()
            If ps <> "" OrElse pn <> "" Then
                Return (ps & "|" & pn).ToUpper()
            End If
            Return ("TITLE|" & If(title, "").Trim()).ToUpper()
        End Function

        '=========================================================
        ' THÊM 1 CỘT VÀO PARTSLIST — dùng late binding
        '=========================================================
        Private Function AddColumn(ByVal pl As Inventor.PartsList,
                                    ByVal propType As Inventor.PropertyTypeEnum,
                                    ByVal info As ColumnInfo) As Boolean

            ' Thử signature 4 tham số (đầy đủ)
            Try
                Dim cols As Object = pl.PartsListColumns
                Dim newCol As Inventor.PartsListColumn =
                    DirectCast(cols.Add(propType, info.PropSet, info.PropName, info.Title),
                               Inventor.PartsListColumn)
                If newCol IsNot Nothing Then
                    Try
                        If info.Width > 0 Then newCol.Width = info.Width
                    Catch
                    End Try
                    Return True
                End If
            Catch
            End Try

            ' Thử signature 3 tham số
            Try
                Dim cols As Object = pl.PartsListColumns
                Dim newCol As Inventor.PartsListColumn =
                    DirectCast(cols.Add(propType, info.PropSet, info.PropName),
                               Inventor.PartsListColumn)
                If newCol IsNot Nothing Then
                    Try : newCol.Title = info.Title : Catch : End Try
                    Try
                        If info.Width > 0 Then newCol.Width = info.Width
                    Catch
                    End Try
                    Return True
                End If
            Catch
            End Try

            ' Thử signature chỉ PropertyType
            Try
                Dim cols As Object = pl.PartsListColumns
                Dim newCol As Inventor.PartsListColumn =
                    DirectCast(cols.Add(propType),
                               Inventor.PartsListColumn)
                If newCol IsNot Nothing Then
                    Try : newCol.Title = info.Title : Catch : End Try
                    Return True
                End If
            Catch
            End Try

            Return False
        End Function


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

                MessageBox.Show("Đã copy PartsList sang sheet '" & dstSheet & "'.",
                                "Copy PartsList", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Lỗi copy: " & ex.Message, "Copy PartsList",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub


        ' =====================================================
        ' HELPER
        ' =====================================================
        Private Function GetInventorApp() As Inventor.Application
            Try
                Return CType(Interop.Marshal2.GetActiveObject("Inventor.Application"),
                             Inventor.Application)
            Catch
                Return Nothing
            End Try
        End Function

    End Module


    ' ============================================================
    ' FORM — Chọn nguồn/đích + mode (UI mới, đồng bộ với các tool khác)
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

        Private Const FORM_W As Integer = 700
        Private Const FORM_H As Integer = 700

        Public Sub New()
            InitializeUI()
            LoadSheets()
            UpdateModeUI()
        End Sub


        Private Sub InitializeUI()
            Me.Text = "PartsList — Đồng bộ cột"
            Me.AutoScaleMode = AutoScaleMode.None
            Me.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
            Me.ClientSize = New Drw.Size(FORM_W, FORM_H)
            Me.FormBorderStyle = FormBorderStyle.FixedSingle
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.BackColor = Drw.Color.FromArgb(245, 245, 245)
            Me.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)

            '===== HEADER =====
            Dim pnlHeader As New Panel()
            pnlHeader.Location = New Drw.Point(0, 0)
            pnlHeader.Size = New Drw.Size(FORM_W, 75)
            pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
            Me.Controls.Add(pnlHeader)

            Dim lblTitle As New Label()
            lblTitle.Text = "PARTSLIST — ĐỒNG BỘ CỘT"
            lblTitle.Font = New Drw.Font("Segoe UI", 14.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblTitle.ForeColor = Drw.Color.White
            lblTitle.Dock = DockStyle.Fill
            lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
            pnlHeader.Controls.Add(lblTitle)

            Dim lblSub As New Label()
            lblSub.Text = "Copy / Sync cột giữa PartsList của 2 sheet"
            lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
            lblSub.Dock = DockStyle.Bottom
            lblSub.Height = 20
            lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
            pnlHeader.Controls.Add(lblSub)

            '===== GROUP 1: CHẾ ĐỘ =====
            Dim gb1 As New GroupBox() With {
                .Text = "1. Chế độ",
                .Location = New Drw.Point(15, 90),
                .Size = New Drw.Size(FORM_W - 30, 80),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gb1)

            _rdoSync = New System.Windows.Forms.RadioButton() With {
                .Text = "Đồng bộ CỘT   (giữ data đích, copy header nguồn)",
                .Location = New Drw.Point(20, 25),
                .Size = New Drw.Size(600, 25),
                .Checked = True,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            _rdoCopy = New System.Windows.Forms.RadioButton() With {
                .Text = "Copy NGUYÊN PartsList   (data + cột từ nguồn)",
                .Location = New Drw.Point(20, 52),
                .Size = New Drw.Size(600, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}

            AddHandler _rdoSync.CheckedChanged, AddressOf OnModeChanged
            AddHandler _rdoCopy.CheckedChanged, AddressOf OnModeChanged

            gb1.Controls.Add(_rdoSync)
            gb1.Controls.Add(_rdoCopy)

            '===== GROUP 2: NGUỒN =====
            Dim gb2 As New GroupBox() With {
                .Text = "2. NGUỒN  (mẫu cột)",
                .Location = New Drw.Point(15, 180),
                .Size = New Drw.Size(FORM_W - 30, 115),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gb2)

            Dim lblSrc1 As New System.Windows.Forms.Label() With {
                .Text = "Sheet:",
                .Location = New Drw.Point(25, 30),
                .Size = New Drw.Size(80, 25),
                .TextAlign = Drw.ContentAlignment.MiddleLeft,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            _cboSrcSheet = New System.Windows.Forms.ComboBox() With {
                .Location = New Drw.Point(110, 28),
                .Size = New Drw.Size(540, 25),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            AddHandler _cboSrcSheet.SelectedIndexChanged, AddressOf OnSrcSheetChanged

            Dim lblSrc2 As New System.Windows.Forms.Label() With {
                .Text = "PartsList:",
                .Location = New Drw.Point(25, 70),
                .Size = New Drw.Size(80, 25),
                .TextAlign = Drw.ContentAlignment.MiddleLeft,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            _cboSrcPL = New System.Windows.Forms.ComboBox() With {
                .Location = New Drw.Point(110, 68),
                .Size = New Drw.Size(540, 25),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            AddHandler _cboSrcPL.SelectedIndexChanged, AddressOf OnSrcPLChanged

            gb2.Controls.Add(lblSrc1)
            gb2.Controls.Add(_cboSrcSheet)
            gb2.Controls.Add(lblSrc2)
            gb2.Controls.Add(_cboSrcPL)

            '===== GROUP 3: ĐÍCH =====
            Dim gb3 As New GroupBox() With {
                .Text = "3. ĐÍCH  (nơi áp dụng)",
                .Location = New Drw.Point(15, 305),
                .Size = New Drw.Size(FORM_W - 30, 115),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gb3)

            Dim lblDst1 As New System.Windows.Forms.Label() With {
                .Text = "Sheet:",
                .Location = New Drw.Point(25, 30),
                .Size = New Drw.Size(80, 25),
                .TextAlign = Drw.ContentAlignment.MiddleLeft,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            _cboDstSheet = New System.Windows.Forms.ComboBox() With {
                .Location = New Drw.Point(110, 28),
                .Size = New Drw.Size(540, 25),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            AddHandler _cboDstSheet.SelectedIndexChanged, AddressOf OnDstSheetChanged

            _lblDstPL = New System.Windows.Forms.Label() With {
                .Text = "PartsList:",
                .Location = New Drw.Point(25, 70),
                .Size = New Drw.Size(80, 25),
                .TextAlign = Drw.ContentAlignment.MiddleLeft,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            _cboDstPL = New System.Windows.Forms.ComboBox() With {
                .Location = New Drw.Point(110, 68),
                .Size = New Drw.Size(540, 25),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}

            gb3.Controls.Add(lblDst1)
            gb3.Controls.Add(_cboDstSheet)
            gb3.Controls.Add(_lblDstPL)
            gb3.Controls.Add(_cboDstPL)

            '===== GROUP 4: PREVIEW =====
            Dim gb4 As New GroupBox() With {
                .Text = "Xem trước cột của NGUỒN",
                .Location = New Drw.Point(15, 430),
                .Size = New Drw.Size(FORM_W - 30, 195),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gb4)

            _lstPreview = New System.Windows.Forms.ListBox() With {
                .Location = New Drw.Point(25, 28),
                .Size = New Drw.Size(625, 150),
                .Font = New Drw.Font("Consolas", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point),
                .BorderStyle = BorderStyle.FixedSingle,
                .BackColor = Drw.Color.FromArgb(250, 250, 250)}
            gb4.Controls.Add(_lstPreview)

            '===== NÚT THỰC HIỆN =====
            _btnOK = New System.Windows.Forms.Button()
            _btnOK.Text = "ÁP DỤNG"
            _btnOK.Size = New Drw.Size(170, 46)
            _btnOK.Location = New Drw.Point(FORM_W - 195, FORM_H - 60)
            _btnOK.FlatStyle = FlatStyle.Flat
            _btnOK.FlatAppearance.BorderSize = 0
            _btnOK.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(60, 115, 195)
            _btnOK.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(30, 80, 155)
            _btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
            _btnOK.ForeColor = Drw.Color.White
            _btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            _btnOK.Cursor = Cursors.Hand
            _btnOK.UseVisualStyleBackColor = False
            AddHandler _btnOK.Click, AddressOf HandleOKClick
            Me.Controls.Add(_btnOK)

            '===== NÚT HỦY =====
            _btnCancel = New System.Windows.Forms.Button()
            _btnCancel.Text = "HỦY"
            _btnCancel.Size = New Drw.Size(130, 46)
            _btnCancel.Location = New Drw.Point(FORM_W - 340, FORM_H - 60)
            _btnCancel.FlatStyle = FlatStyle.Flat
            _btnCancel.FlatAppearance.BorderSize = 1
            _btnCancel.FlatAppearance.BorderColor = Drw.Color.FromArgb(200, 200, 200)
            _btnCancel.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(235, 235, 235)
            _btnCancel.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(215, 215, 215)
            _btnCancel.BackColor = Drw.Color.FromArgb(250, 250, 250)
            _btnCancel.ForeColor = Drw.Color.FromArgb(60, 60, 60)
            _btnCancel.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            _btnCancel.Cursor = Cursors.Hand
            _btnCancel.UseVisualStyleBackColor = False
            _btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Controls.Add(_btnCancel)

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

                For Each n In Drawing_PartsList_Sync2020.GetSheetNames()
                    _cboSrcSheet.Items.Add(n)
                    _cboDstSheet.Items.Add(n)
                Next

                If _cboSrcSheet.Items.Count > 0 Then _cboSrcSheet.SelectedIndex = 0
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
                    MessageBox.Show("Chưa chọn đủ thông tin!", "PartsList Sync",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim srcSheet As String = _cboSrcSheet.SelectedItem.ToString()
                Dim dstSheet As String = _cboDstSheet.SelectedItem.ToString()
                Dim srcIdx As Integer = ParsePlIndex(_cboSrcPL.SelectedItem.ToString())

                _btnOK.Enabled = False
                _btnCancel.Enabled = False

                If _rdoSync.Checked Then
                    If _cboDstPL.SelectedItem Is Nothing Then
                        MessageBox.Show("Chưa chọn PartsList đích!", "PartsList Sync",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
                MessageBox.Show("Lỗi: " & ex.Message, "PartsList Sync",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End Sub

    End Class

End Namespace