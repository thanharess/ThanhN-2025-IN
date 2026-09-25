Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.DrawView


    ' ============================================================
    ' MODULE — AUTO BALLOON (1 sheet hoặc TẤT CẢ sheet)
    ' ============================================================
    Public Module Draw_AutoBalloon

        '=============================================================
        ' ENTRY POINT
        '=============================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                Dim invApp = GetInventorApp()
                If invApp Is Nothing Then
                    MessageBox.Show("Không lấy được Inventor Application!", "Auto Balloon")
                    Return
                End If
                If invApp.ActiveDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Mở file Drawing (.idw/.dwg) trước.", "Auto Balloon")
                    Return
                End If

                Dim frm As New Form_AutoBalloon()
                frm.ShowDialog()
                frm.Dispose()

            Catch ex As Exception
                Try
                    MessageBox.Show("Lỗi: " & ex.Message, "Auto Balloon")
                Catch
                End Try
            End Try
        End Sub


        '=============================================================
        ' LẤY DANH SÁCH SHEET
        '=============================================================
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


        '=============================================================
        ' LẤY DRAWING VIEW TRÊN SHEET
        '=============================================================
        Public Function GetDrawingViewsOnSheet(sheetName As String) As List(Of String)
            Dim result As New List(Of String)
            Try
                Dim invApp = GetInventorApp()
                If invApp Is Nothing Then Return result

                Dim drawDoc As Inventor.DrawingDocument = CType(invApp.ActiveDocument, Inventor.DrawingDocument)
                Dim oSheet As Inventor.Sheet = drawDoc.Sheets.Item(sheetName)

                Dim idx As Integer = 0
                For Each v As Inventor.DrawingView In oSheet.DrawingViews
                    idx += 1
                    Dim label As String = "View #" & idx
                    Try : label &= "  — " & v.Name : Catch : End Try
                    result.Add(idx & "|" & label)
                Next
            Catch
            End Try
            Return result
        End Function


        '=============================================================
        ' LẤY PARTSLIST TRÊN SHEET
        '=============================================================
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
                    Try : cnt = pl.PartsListRows.Count : Catch : End Try
                    result.Add(idx & "|PL #" & idx & "  (" & cnt & " dòng)")
                Next
            Catch
            End Try
            Return result
        End Function


        '=============================================================
        ' ⭐ HELPER: LẤY OCCURRENCE TỪ DRAWING CURVE
        '   ModelGeometry là EdgeProxy/FaceProxy/VertexProxy — KHÔNG phải occ
        '=============================================================
        Private Function GetOccurrenceFromCurve(dc As Inventor.DrawingCurve) As Inventor.ComponentOccurrence
            If dc Is Nothing Then Return Nothing

            Dim mg As Object = Nothing
            Try : mg = dc.ModelGeometry : Catch : End Try
            If mg Is Nothing Then Return Nothing

            ' ─── EdgeProxy ───
            Try
                If TypeOf mg Is Inventor.EdgeProxy Then
                    Return CType(mg, Inventor.EdgeProxy).ContainingOccurrence
                End If
            Catch
            End Try

            ' ─── FaceProxy ───
            Try
                If TypeOf mg Is Inventor.FaceProxy Then
                    Return CType(mg, Inventor.FaceProxy).ContainingOccurrence
                End If
            Catch
            End Try

            ' ─── VertexProxy ───
            Try
                If TypeOf mg Is Inventor.VertexProxy Then
                    Return CType(mg, Inventor.VertexProxy).ContainingOccurrence
                End If
            Catch
            End Try

            ' ─── Fallback hiếm ───
            Try
                If TypeOf mg Is Inventor.ComponentOccurrence Then
                    Return CType(mg, Inventor.ComponentOccurrence)
                End If
            Catch
            End Try

            Return Nothing
        End Function


        '=============================================================
        ' ⭐ ĐỌC PARTSLIST — KHÔNG DÙNG TYPE PartsListCells
        '=============================================================
        Private Sub ReadPartsList(ByVal oPL As Inventor.PartsList,
                                   ByRef plItemNums As List(Of String),
                                   ByRef plPartNums As List(Of String),
                                   ByRef errMsg As String)

            plItemNums = New List(Of String)
            plPartNums = New List(Of String)
            errMsg = ""

            Try
                ' ─── Xác định cột Item / Part Number theo tên header ───
                Dim colItemIdx As Integer = 1
                Dim colPartIdx As Integer = 3
                Dim foundItemCol As Boolean = False
                Dim foundPartCol As Boolean = False

                Try
                    Dim colsObj As Object = oPL.PartsListColumns
                    Dim colCount As Integer = CInt(colsObj.Count)
                    For ci As Integer = 1 To colCount
                        Dim colObj As Object = colsObj.Item(ci)
                        Dim title As String = ""
                        Try : title = colObj.Title.ToString().Trim() : Catch : End Try

                        ' Item number column
                        If title.Equals("ITEM", StringComparison.OrdinalIgnoreCase) OrElse
                           title.Equals("ITEM NO.", StringComparison.OrdinalIgnoreCase) OrElse
                           title.Equals("ITEM NO", StringComparison.OrdinalIgnoreCase) OrElse
                           title.Equals("ITEM NUMBER", StringComparison.OrdinalIgnoreCase) OrElse
                           title.Equals("STT", StringComparison.OrdinalIgnoreCase) Then
                            colItemIdx = ci
                            foundItemCol = True
                        End If

                        ' Part number column — nhận cả STOCK NUMBER
                        If title.Equals("PART NUMBER", StringComparison.OrdinalIgnoreCase) OrElse
                           title.Equals("PART NO.", StringComparison.OrdinalIgnoreCase) OrElse
                           title.Equals("PART NO", StringComparison.OrdinalIgnoreCase) OrElse
                           title.Equals("STOCK NUMBER", StringComparison.OrdinalIgnoreCase) OrElse
                           title.Equals("STOCK NO.", StringComparison.OrdinalIgnoreCase) OrElse
                           title.Equals("STOCK NO", StringComparison.OrdinalIgnoreCase) OrElse
                           title.Equals("PN", StringComparison.OrdinalIgnoreCase) Then
                            colPartIdx = ci
                            foundPartCol = True
                        End If
                    Next
                Catch
                End Try

                ' ─── Đọc từng dòng — late binding ───
                Dim rowsObj As Object = oPL.PartsListRows
                Dim rowCount As Integer = CInt(rowsObj.Count)

                For i As Integer = 1 To rowCount
                    Dim rowObj As Object = Nothing
                    Try : rowObj = rowsObj.Item(i) : Catch : End Try
                    If rowObj Is Nothing Then Continue For

                    Dim itemNum As String = ""
                    Dim partNum As String = ""

                    ' Cột Item
                    If foundItemCol Then
                        Try
                            Dim cellObj As Object = rowObj.Item(colItemIdx)
                            If cellObj IsNot Nothing Then
                                Dim v As Object = cellObj.Value
                                If v IsNot Nothing Then itemNum = v.ToString().Trim()
                            End If
                        Catch
                        End Try
                    End If

                    ' Cột Part Number
                    Try
                        Dim cellObj As Object = rowObj.Item(colPartIdx)
                        If cellObj IsNot Nothing Then
                            Dim v As Object = cellObj.Value
                            If v IsNot Nothing Then partNum = v.ToString().Trim()
                        End If
                    Catch
                    End Try

                    plItemNums.Add(itemNum)
                    plPartNums.Add(partNum)
                Next

                ' ─── Fallback: nếu PL không có cột Item → dùng số thứ tự ───
                If Not foundItemCol Then
                    For i As Integer = 0 To plItemNums.Count - 1
                        plItemNums(i) = (i + 1).ToString()
                    Next
                Else
                    ' Nếu cột Item có nhưng toàn rỗng → fallback
                    Dim allEmpty As Boolean = True
                    For i As Integer = 0 To plItemNums.Count - 1
                        If Not String.IsNullOrEmpty(plItemNums(i)) Then
                            allEmpty = False
                            Exit For
                        End If
                    Next
                    If allEmpty Then
                        For i As Integer = 0 To plItemNums.Count - 1
                            plItemNums(i) = (i + 1).ToString()
                        Next
                    End If
                End If

            Catch ex As Exception
                errMsg = ex.Message
            End Try
        End Sub


        '=============================================================
        ' ⭐ TẠO BALLOON CHO 1 SHEET
        '=============================================================
        Public Sub CreateBalloonsForSheet(sheetName As String,
                                           viewIndex As Integer,
                                           plIndex As Integer,
                                           ByRef okCount As Integer,
                                           ByRef skipCount As Integer,
                                           ByRef failCount As Integer,
                                           ByRef errLog As String)

            okCount = 0 : skipCount = 0 : failCount = 0

            Dim invApp As Inventor.Application = Nothing
            Try
                invApp = GetInventorApp()
                If invApp Is Nothing Then Return

                Dim drawDoc As Inventor.DrawingDocument = CType(invApp.ActiveDocument, Inventor.DrawingDocument)
                Dim oSheet As Inventor.Sheet = drawDoc.Sheets.Item(sheetName)

                invApp.SilentOperation = False
                Try : oSheet.Activate() : Catch : End Try
                Try : drawDoc.Update() : Catch : End Try

                ' ─── Lấy view + PL ───
                Dim oView As Inventor.DrawingView = Nothing
                Try : oView = oSheet.DrawingViews.Item(viewIndex) : Catch : End Try
                If oView Is Nothing Then
                    errLog &= "• Sheet '" & sheetName & "': không có view #" & viewIndex & vbCrLf
                    Return
                End If

                Dim oPL As Inventor.PartsList = Nothing
                Try : oPL = oSheet.PartsLists.Item(plIndex) : Catch : End Try
                If oPL Is Nothing Then
                    errLog &= "• Sheet '" & sheetName & "': không có PartsList #" & plIndex & vbCrLf
                    Return
                End If

                ' ─── Đọc PL ───
                Dim plItemNums As New List(Of String)
                Dim plPartNums As New List(Of String)
                Dim plErr As String = ""

                ReadPartsList(oPL, plItemNums, plPartNums, plErr)

                If Not String.IsNullOrEmpty(plErr) Then
                    errLog &= "• Sheet '" & sheetName & "': đọc PL lỗi — " & plErr & vbCrLf
                    Return
                End If

                If plItemNums.Count = 0 Then
                    errLog &= "• Sheet '" & sheetName & "': PartsList không có dòng nào" & vbCrLf
                    Return
                End If

                ' ─── Map PartNumber → Item Number ───
                Dim pnToItem As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                For i As Integer = 0 To plItemNums.Count - 1
                    Dim pn As String = plPartNums(i)
                    Dim item As String = plItemNums(i)
                    If Not String.IsNullOrEmpty(pn) AndAlso Not pnToItem.ContainsKey(pn) Then
                        pnToItem.Add(pn, item)
                    End If
                Next

                If pnToItem.Count = 0 Then
                    errLog &= "• Sheet '" & sheetName & "': PL không đọc được Part Number nào" & vbCrLf
                    Return
                End If

                ' ─── Balloons collection thuộc SHEET ───
                Dim balloons As Inventor.Balloons = Nothing
                Try : balloons = oSheet.Balloons : Catch : End Try
                If balloons Is Nothing Then
                    errLog &= "• Sheet '" & sheetName & "': không truy cập được Balloons collection" & vbCrLf
                    Return
                End If

                ' ─── Quét DrawingCurves để lấy occurrence → PartNumber ───
                Dim processedPartNums As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each dc As Inventor.DrawingCurve In oView.DrawingCurves
                    Try
                        ' ⭐ Lấy occurrence ĐÚNG CÁCH
                        Dim occ As Inventor.ComponentOccurrence = GetOccurrenceFromCurve(dc)
                        If occ Is Nothing Then Continue For

                        ' Lấy PartNumber
                        Dim pn As String = ""
                        Try
                            Dim propSets As Inventor.PropertySets = occ.Definition.Document.PropertySets
                            Dim dtp As Inventor.PropertySet = propSets.Item("Design Tracking Properties")
                            pn = dtp.Item("Part Number").Value.ToString().Trim()
                        Catch
                        End Try

                        If String.IsNullOrEmpty(pn) Then
                            skipCount += 1
                            Continue For
                        End If

                        ' Tra Item Number
                        If Not pnToItem.ContainsKey(pn) Then
                            skipCount += 1
                            Continue For
                        End If

                        ' Mỗi PartNumber chỉ tạo 1 balloon (bỏ qua các occurrence cùng part)
                        If processedPartNums.Contains(pn) Then
                            skipCount += 1
                            Continue For
                        End If

                        Dim itemNum As String = pnToItem(pn)

                        ' ─── TẠO BALLOON ───
                        Try
                            Dim pos As Inventor.Point2d = GetDefaultBalloonPosition(oView, occ)
                            Dim b As Inventor.Balloon = Nothing

                            ' Thử Add trên Sheet.Balloons
                            Try
                                b = balloons.Add(itemNum, pos)
                            Catch exAdd As Exception
                                errLog &= "• Add lỗi cho " & pn & ": " & exAdd.Message & vbCrLf
                            End Try

                            If b IsNot Nothing Then
                                Try : b.Value = itemNum : Catch : End Try
                                okCount += 1
                                processedPartNums.Add(pn)
                            Else
                                failCount += 1
                                errLog &= "• " & pn & " (Item " & itemNum & "): không tạo được" & vbCrLf
                            End If
                        Catch ex As Exception
                            failCount += 1
                            errLog &= "• " & pn & ": " & ex.Message & vbCrLf
                        End Try

                    Catch
                    End Try
                Next

                ' ─── Sắp xếp quanh view ───
                Try : ArrangeBalloonsAroundView(oSheet, oView) : Catch : End Try

                Try : drawDoc.Update2(True) : Catch : End Try

            Catch ex As Exception
                errLog &= "• Sheet '" & sheetName & "': lỗi tổng — " & ex.Message & vbCrLf
            Finally
                Try
                    If invApp IsNot Nothing Then invApp.SilentOperation = False
                Catch
                End Try
            End Try
        End Sub


        '=============================================================
        ' TẠO BALLOON CHO TẤT CẢ SHEET
        '=============================================================
        Public Sub CreateBalloonsAllSheets(ByRef totalOK As Integer,
                                            ByRef totalSkip As Integer,
                                            ByRef totalFail As Integer,
                                            ByRef totalSheets As Integer,
                                            ByRef totalSheetsSkipped As Integer,
                                            ByRef errLog As String)

            totalOK = 0 : totalSkip = 0 : totalFail = 0
            totalSheets = 0 : totalSheetsSkipped = 0

            Dim invApp As Inventor.Application = Nothing
            Dim originalSheet As Inventor.Sheet = Nothing

            Try
                invApp = GetInventorApp()
                If invApp Is Nothing Then Return

                Dim drawDoc As Inventor.DrawingDocument = CType(invApp.ActiveDocument, Inventor.DrawingDocument)
                originalSheet = drawDoc.ActiveSheet

                invApp.SilentOperation = False

                For Each oSheet As Inventor.Sheet In drawDoc.Sheets
                    totalSheets += 1

                    ' ─── Tìm view có model đầu tiên ───
                    Dim viewIdx As Integer = -1
                    Try
                        For i As Integer = 1 To oSheet.DrawingViews.Count
                            Dim v As Inventor.DrawingView = oSheet.DrawingViews.Item(i)
                            Dim refDoc As Inventor.Document = Nothing
                            Try : refDoc = v.ReferencedDocumentDescriptor.ReferencedDocument : Catch : End Try
                            If refDoc IsNot Nothing Then
                                viewIdx = i
                                Exit For
                            End If
                        Next
                    Catch
                    End Try

                    ' ─── Tìm PartsList đầu tiên ───
                    Dim plIdx As Integer = -1
                    Try
                        If oSheet.PartsLists.Count > 0 Then plIdx = 1
                    Catch
                    End Try

                    If viewIdx < 1 OrElse plIdx < 1 Then
                        totalSheetsSkipped += 1
                        errLog &= "⊘ Sheet '" & oSheet.Name & "': thiếu view có model hoặc PartsList" & vbCrLf
                        Continue For
                    End If

                    Dim ok As Integer = 0, sk As Integer = 0, fl As Integer = 0
                    Dim sheetErr As String = ""

                    CreateBalloonsForSheet(oSheet.Name, viewIdx, plIdx,
                                            ok, sk, fl, sheetErr)

                    totalOK += ok
                    totalSkip += sk
                    totalFail += fl

                    If Not String.IsNullOrEmpty(sheetErr) Then
                        errLog &= sheetErr
                    End If
                Next

                Try : drawDoc.Update2(True) : Catch : End Try

            Catch ex As Exception
                errLog &= "❌ Lỗi tổng: " & ex.Message & vbCrLf
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
        End Sub


        '=============================================================
        ' VỊ TRÍ BALLOON MẶC ĐỊNH — cạnh view
        '=============================================================
        Private Function GetDefaultBalloonPosition(ByVal oView As Inventor.DrawingView,
                                                    ByVal occ As Inventor.ComponentOccurrence) As Inventor.Point2d
            Try
                Dim tg As Inventor.TransientGeometry = GetInventorApp.TransientGeometry
                Dim center As Inventor.Point2d = oView.Position
                Dim w As Double = 0, h As Double = 0
                Try : w = oView.Width : Catch : End Try
                Try : h = oView.Height : Catch : End Try
                Return tg.CreatePoint2d(center.X + w / 2 + 1.5, center.Y + h / 2 - 1.5)
            Catch
                Return Nothing
            End Try
        End Function


        '=============================================================
        ' SẮP XẾP BALLOON QUANH VIEW
        '=============================================================
        Private Sub ArrangeBalloonsAroundView(ByVal oSheet As Inventor.Sheet,
                                               ByVal oView As Inventor.DrawingView)
            Try
                Dim balloons As Inventor.Balloons = Nothing
                Try : balloons = oSheet.Balloons : Catch : End Try
                If balloons Is Nothing Then Return

                Dim count As Integer = 0
                Try : count = balloons.Count : Catch : End Try
                If count <= 1 Then Return

                Dim tg As Inventor.TransientGeometry = GetInventorApp.TransientGeometry
                Dim center As Inventor.Point2d = oView.Position
                Dim w As Double = 0, h As Double = 0
                Try : w = oView.Width : Catch : End Try
                Try : h = oView.Height : Catch : End Try

                Dim radius As Double = Math.Max(w, h) / 2 + 2.0

                For i As Integer = 1 To count
                    Try
                        Dim b As Inventor.Balloon = balloons.Item(i)
                        Dim angle As Double = Math.PI / 2 + (i - 1) * (2 * Math.PI / count)
                        Dim x As Double = center.X + radius * Math.Cos(angle)
                        Dim y As Double = center.Y + radius * Math.Sin(angle)
                        b.Position = tg.CreatePoint2d(x, y)
                    Catch
                    End Try
                Next
            Catch
            End Try
        End Sub


        '=============================================================
        ' HELPER
        '=============================================================
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
    ' FORM — giữ nguyên
    ' ============================================================
    Public Class Form_AutoBalloon
        Inherits System.Windows.Forms.Form

        Private _rdoAll As System.Windows.Forms.RadioButton
        Private _rdoOne As System.Windows.Forms.RadioButton
        Private _cboSheet As System.Windows.Forms.ComboBox
        Private _cboView As System.Windows.Forms.ComboBox
        Private _cboPL As System.Windows.Forms.ComboBox
        Private _lblSheet As System.Windows.Forms.Label
        Private _lblView As System.Windows.Forms.Label
        Private _lblPL As System.Windows.Forms.Label
        Private _lstPreview As System.Windows.Forms.ListBox
        Private _btnOK As System.Windows.Forms.Button
        Private _btnCancel As System.Windows.Forms.Button

        Private Const FORM_W As Integer = 640
        Private Const FORM_H As Integer = 620

        Public Sub New()
            InitializeUI()
            LoadSheets()
            UpdateModeUI()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Auto Balloon"
            Me.AutoScaleMode = AutoScaleMode.Dpi
            Me.AutoScaleDimensions = New Drw.SizeF(96.0F, 96.0F)
            Me.ClientSize = New Drw.Size(FORM_W, FORM_H)
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.BackColor = Drw.Color.FromArgb(245, 245, 245)
            Me.Font = New Drw.Font("Segoe UI", 9.5F)

            '── HEADER ──
            Dim pnlHeader As New Panel() With {
                .Location = New Drw.Point(0, 0),
                .Size = New Drw.Size(FORM_W, 75),
                .BackColor = Drw.Color.FromArgb(45, 100, 180)
            }
            Me.Controls.Add(pnlHeader)

            Dim lblTitle As New Label() With {
                .Text = "AUTO BALLOON",
                .Font = New Drw.Font("Segoe UI", 14.0F, Drw.FontStyle.Bold),
                .ForeColor = Drw.Color.White,
                .Dock = DockStyle.Fill,
                .TextAlign = Drw.ContentAlignment.MiddleCenter
            }
            pnlHeader.Controls.Add(lblTitle)

            Dim lblSub As New Label() With {
                .Text = "Đánh số bong bóng theo PartsList — 1 sheet hoặc tất cả",
                .Font = New Drw.Font("Segoe UI", 9.0F),
                .ForeColor = Drw.Color.FromArgb(220, 230, 245),
                .Dock = DockStyle.Bottom,
                .Height = 20,
                .TextAlign = Drw.ContentAlignment.MiddleCenter
            }
            pnlHeader.Controls.Add(lblSub)

            '── GROUP 1: PHẠM VI ──
            Dim gb1 As New GroupBox() With {
                .Text = "1. Phạm vi",
                .Location = New Drw.Point(15, 90),
                .Size = New Drw.Size(FORM_W - 30, 85),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White
            }
            Me.Controls.Add(gb1)

            _rdoAll = New System.Windows.Forms.RadioButton() With {
                .Text = "Tất cả sheet  (tự động tìm view + PL trên mỗi sheet)",
                .Location = New Drw.Point(20, 25),
                .Size = New Drw.Size(580, 25),
                .Checked = True
            }
            _rdoOne = New System.Windows.Forms.RadioButton() With {
                .Text = "Chỉ 1 sheet  (chọn bên dưới)",
                .Location = New Drw.Point(20, 52),
                .Size = New Drw.Size(580, 25)
            }
            AddHandler _rdoAll.CheckedChanged, AddressOf OnModeChanged
            AddHandler _rdoOne.CheckedChanged, AddressOf OnModeChanged

            gb1.Controls.Add(_rdoAll)
            gb1.Controls.Add(_rdoOne)

            '── GROUP 2: CHỌN ──
            Dim gb2 As New GroupBox() With {
                .Text = "2. Cấu hình",
                .Location = New Drw.Point(15, 185),
                .Size = New Drw.Size(FORM_W - 30, 170),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White
            }
            Me.Controls.Add(gb2)

            _lblSheet = New System.Windows.Forms.Label() With {
                .Text = "Sheet:", .Location = New Drw.Point(25, 35),
                .Size = New Drw.Size(90, 25),
                .TextAlign = Drw.ContentAlignment.MiddleLeft
            }
            _cboSheet = New System.Windows.Forms.ComboBox() With {
                .Location = New Drw.Point(120, 33),
                .Size = New Drw.Size(FORM_W - 175, 25),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }
            AddHandler _cboSheet.SelectedIndexChanged, AddressOf OnSheetChanged

            _lblView = New System.Windows.Forms.Label() With {
                .Text = "Drawing View:", .Location = New Drw.Point(25, 75),
                .Size = New Drw.Size(90, 25),
                .TextAlign = Drw.ContentAlignment.MiddleLeft
            }
            _cboView = New System.Windows.Forms.ComboBox() With {
                .Location = New Drw.Point(120, 73),
                .Size = New Drw.Size(FORM_W - 175, 25),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }
            AddHandler _cboView.SelectedIndexChanged, AddressOf OnViewChanged

            _lblPL = New System.Windows.Forms.Label() With {
                .Text = "PartsList:", .Location = New Drw.Point(25, 115),
                .Size = New Drw.Size(90, 25),
                .TextAlign = Drw.ContentAlignment.MiddleLeft
            }
            _cboPL = New System.Windows.Forms.ComboBox() With {
                .Location = New Drw.Point(120, 113),
                .Size = New Drw.Size(FORM_W - 175, 25),
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            }

            gb2.Controls.Add(_lblSheet) : gb2.Controls.Add(_cboSheet)
            gb2.Controls.Add(_lblView) : gb2.Controls.Add(_cboView)
            gb2.Controls.Add(_lblPL) : gb2.Controls.Add(_cboPL)

            '── GROUP 3: PREVIEW ──
            Dim gb3 As New GroupBox() With {
                .Text = "Thông tin",
                .Location = New Drw.Point(15, 365),
                .Size = New Drw.Size(FORM_W - 30, 180),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White
            }
            Me.Controls.Add(gb3)

            _lstPreview = New System.Windows.Forms.ListBox() With {
                .Location = New Drw.Point(20, 28),
                .Size = New Drw.Size(FORM_W - 70, 135),
                .Font = New Drw.Font("Consolas", 9.0F),
                .BorderStyle = BorderStyle.FixedSingle,
                .BackColor = Drw.Color.FromArgb(250, 250, 250)
            }
            gb3.Controls.Add(_lstPreview)

            '── NÚT OK ──
            _btnOK = New System.Windows.Forms.Button()
            _btnOK.Text = "TẠO BALLOON"
            _btnOK.Size = New Drw.Size(170, 44)
            _btnOK.Location = New Drw.Point(FORM_W - 195, FORM_H - 60)
            _btnOK.FlatStyle = FlatStyle.Flat
            _btnOK.FlatAppearance.BorderSize = 0
            _btnOK.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(60, 115, 195)
            _btnOK.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(30, 80, 155)
            _btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
            _btnOK.ForeColor = Drw.Color.White
            _btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold)
            _btnOK.Cursor = Cursors.Hand
            _btnOK.UseVisualStyleBackColor = False
            AddHandler _btnOK.Click, AddressOf HandleOKClick
            Me.Controls.Add(_btnOK)

            '── NÚT HỦY ──
            _btnCancel = New System.Windows.Forms.Button()
            _btnCancel.Text = "HỦY"
            _btnCancel.Size = New Drw.Size(130, 44)
            _btnCancel.Location = New Drw.Point(FORM_W - 340, FORM_H - 60)
            _btnCancel.FlatStyle = FlatStyle.Flat
            _btnCancel.FlatAppearance.BorderSize = 1
            _btnCancel.FlatAppearance.BorderColor = Drw.Color.FromArgb(200, 200, 200)
            _btnCancel.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(235, 235, 235)
            _btnCancel.BackColor = Drw.Color.FromArgb(250, 250, 250)
            _btnCancel.ForeColor = Drw.Color.FromArgb(60, 60, 60)
            _btnCancel.Font = New Drw.Font("Segoe UI", 10.0F)
            _btnCancel.Cursor = Cursors.Hand
            _btnCancel.UseVisualStyleBackColor = False
            _btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Controls.Add(_btnCancel)

            Me.AcceptButton = _btnOK
            Me.CancelButton = _btnCancel
        End Sub

        Private Sub UpdateModeUI()
            Dim useAll As Boolean = _rdoAll.Checked
            _cboSheet.Enabled = Not useAll
            _cboView.Enabled = Not useAll
            _cboPL.Enabled = Not useAll
            _lblSheet.Enabled = Not useAll
            _lblView.Enabled = Not useAll
            _lblPL.Enabled = Not useAll
        End Sub

        Private Sub OnModeChanged(sender As Object, e As EventArgs)
            UpdateModeUI()
            RefreshPreview()
        End Sub

        Private Sub LoadSheets()
            Try
                _cboSheet.Items.Clear()
                For Each n In Draw_AutoBalloon.GetSheetNames()
                    _cboSheet.Items.Add(n)
                Next
                If _cboSheet.Items.Count > 0 Then _cboSheet.SelectedIndex = 0
            Catch
            End Try
        End Sub

        Private Sub OnSheetChanged(sender As Object, e As EventArgs)
            Try
                _cboView.Items.Clear()
                _cboPL.Items.Clear()
                If _cboSheet.SelectedItem Is Nothing Then Return

                Dim sh As String = _cboSheet.SelectedItem.ToString()

                For Each v In Draw_AutoBalloon.GetDrawingViewsOnSheet(sh)
                    _cboView.Items.Add(v)
                Next
                If _cboView.Items.Count > 0 Then _cboView.SelectedIndex = 0

                For Each p In Draw_AutoBalloon.GetPartsListsOnSheet(sh)
                    _cboPL.Items.Add(p)
                Next
                If _cboPL.Items.Count > 0 Then _cboPL.SelectedIndex = 0

                RefreshPreview()
            Catch
            End Try
        End Sub

        Private Sub OnViewChanged(sender As Object, e As EventArgs)
            RefreshPreview()
        End Sub

        Private Sub RefreshPreview()
            Try
                _lstPreview.Items.Clear()

                If _rdoAll.Checked Then
                    _lstPreview.Items.Add("Chế độ: TẤT CẢ SHEET")
                    _lstPreview.Items.Add("")
                    _lstPreview.Items.Add("Trên mỗi sheet, tool sẽ tự động:")
                    _lstPreview.Items.Add("  • Tìm drawing view đầu tiên có model")
                    _lstPreview.Items.Add("  • Tìm PartsList đầu tiên")
                    _lstPreview.Items.Add("  • Đọc Item Number + Part Number từ PL")
                    _lstPreview.Items.Add("  • Quét occurrence trong view")
                    _lstPreview.Items.Add("  • Tạo balloon với số Item tương ứng")
                    _lstPreview.Items.Add("  • Sắp xếp balloon quanh view")
                    _lstPreview.Items.Add("")
                    _lstPreview.Items.Add("Sheet thiếu view/PL sẽ bị bỏ qua.")
                Else
                    _lstPreview.Items.Add("Chế độ: 1 SHEET")
                    _lstPreview.Items.Add("")
                    If _cboSheet.SelectedItem IsNot Nothing Then
                        _lstPreview.Items.Add("Sheet: " & _cboSheet.SelectedItem.ToString())
                    End If
                    If _cboView.SelectedItem IsNot Nothing Then
                        _lstPreview.Items.Add("View : " & _cboView.SelectedItem.ToString())
                    End If
                    If _cboPL.SelectedItem IsNot Nothing Then
                        _lstPreview.Items.Add("PL   : " & _cboPL.SelectedItem.ToString())
                    End If
                End If
            Catch
            End Try
        End Sub

        Private Function ParseIndex(s As String) As Integer
            Try
                Dim i As Integer = s.IndexOf("|"c)
                If i > 0 Then Return Integer.Parse(s.Substring(0, i))
            Catch
            End Try
            Return 1
        End Function

        Private Sub HandleOKClick(sender As Object, e As EventArgs)
            Try
                _btnOK.Enabled = False
                _btnCancel.Enabled = False
                Me.Cursor = Cursors.WaitCursor

                Dim ok As Integer = 0, sk As Integer = 0, fl As Integer = 0
                Dim sheetsTotal As Integer = 0, sheetsSkip As Integer = 0
                Dim errLog As String = ""

                If _rdoAll.Checked Then
                    Draw_AutoBalloon.CreateBalloonsAllSheets(
                        ok, sk, fl, sheetsTotal, sheetsSkip, errLog)

                    Dim msg As String =
                        "Hoàn tất Auto Balloon!" & vbCrLf & vbCrLf &
                        "Tổng sheet xử lý : " & (sheetsTotal - sheetsSkip).ToString() & " / " & sheetsTotal.ToString() & vbCrLf &
                        "Sheet bỏ qua    : " & sheetsSkip.ToString() & vbCrLf &
                        "" & vbCrLf &
                        "── KẾT QUẢ ──" & vbCrLf &
                        "  ✔ Tạo được : " & ok.ToString() & vbCrLf &
                        "  ⊘ Bỏ qua   : " & sk.ToString() & vbCrLf &
                        "  ✘ Lỗi      : " & fl.ToString()

                    If Not String.IsNullOrEmpty(errLog) Then
                        Dim logShort As String = errLog
                        If logShort.Length > 1500 Then logShort = logShort.Substring(0, 1500) & "...(còn nữa)"
                        msg &= vbCrLf & vbCrLf & "── Chi tiết ──" & vbCrLf & logShort
                    End If

                    MessageBox.Show(msg, "Auto Balloon",
                                    MessageBoxButtons.OK,
                                    If(fl > 0, MessageBoxIcon.Warning, MessageBoxIcon.Information))
                Else
                    If _cboSheet.SelectedItem Is Nothing OrElse
                       _cboView.SelectedItem Is Nothing OrElse
                       _cboPL.SelectedItem Is Nothing Then
                        MessageBox.Show("Chưa chọn đủ thông tin!", "Auto Balloon",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If

                    Dim sh As String = _cboSheet.SelectedItem.ToString()
                    Dim vIdx As Integer = ParseIndex(_cboView.SelectedItem.ToString())
                    Dim plIdx As Integer = ParseIndex(_cboPL.SelectedItem.ToString())

                    Draw_AutoBalloon.CreateBalloonsForSheet(
                        sh, vIdx, plIdx, ok, sk, fl, errLog)

                    Dim msg As String =
                        "Hoàn tất tạo Balloon!" & vbCrLf & vbCrLf &
                        "Sheet : " & sh & vbCrLf &
                        "View  : " & vIdx & vbCrLf &
                        "PL    : " & plIdx & vbCrLf & vbCrLf &
                        "── KẾT QUẢ ──" & vbCrLf &
                        "  ✔ Tạo được : " & ok.ToString() & vbCrLf &
                        "  ⊘ Bỏ qua   : " & sk.ToString() & vbCrLf &
                        "  ✘ Lỗi      : " & fl.ToString()

                    If Not String.IsNullOrEmpty(errLog) Then
                        Dim logShort As String = errLog
                        If logShort.Length > 1500 Then logShort = logShort.Substring(0, 1500) & "...(còn nữa)"
                        msg &= vbCrLf & vbCrLf & "── Chi tiết ──" & vbCrLf & logShort
                    End If

                    MessageBox.Show(msg, "Auto Balloon",
                                    MessageBoxButtons.OK,
                                    If(fl > 0, MessageBoxIcon.Warning, MessageBoxIcon.Information))
                End If

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message, "Auto Balloon",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Me.Cursor = Cursors.Default
            End Try

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End Sub
    End Class

End Namespace