Option Explicit On
Option Strict Off

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports Inventor
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.Drawtext

    '═══════════════════════════════════════════════════════════
    ' Enum phạm vi áp dụng
    '═══════════════════════════════════════════════════════════
    Public Enum ApplyScope
        SelectedObjects = 0
        CurrentSheet = 1
        AllSheets = 2
    End Enum


    Public Module ThayChuTrongTextModule

        '═══════════════════════════════════════════════════════════
        ' Đường dẫn file lịch sử
        '═══════════════════════════════════════════════════════════
        Friend ReadOnly HistoryFile As String =
            System.IO.Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData),
                "ToolInventor2025", "replace_history.txt")

        '═══════════════════════════════════════════════════════════
        ' STATE cho InteractionEvents
        '═══════════════════════════════════════════════════════════
        Private _inventorApp As Inventor.Application
        Private _oDrawDoc As DrawingDocument
        Private _pendingPairs As List(Of ReplacePair)
        Private _interactEvents As InteractionEvents
        Private _selectEvents As SelectEvents
        Private _collectedObjects As New List(Of Object)()


        '═══════════════════════════════════════════════════════════
        ' LỆNH CHÍNH
        '═══════════════════════════════════════════════════════════
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                _inventorApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"),
                                     Inventor.Application)
            Catch ex As Exception
                MessageBox.Show("Không tìm thấy Inventor đang chạy." & vbCrLf & ex.Message,
                                "Thay chữ trong Text",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            If _inventorApp.ActiveDocument Is Nothing OrElse
               _inventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                MessageBox.Show("Vui lòng mở một bản vẽ (IDW/DWG) trước.",
                                "Thay chữ trong Text",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            _oDrawDoc = CType(_inventorApp.ActiveDocument, DrawingDocument)

            ' ── 2. Đọc lịch sử + Đếm selection có sẵn ──
            Dim history = LoadHistory()
            Dim preSelectionCount As Integer = 0
            Try : preSelectionCount = _oDrawDoc.SelectSet.Count : Catch : End Try

            ' ── 3. Mở form ──
            Dim scope As ApplyScope
            Dim doSaveHistory As Boolean

            Using form As New ReplaceForm(preSelectionCount, history)
                If form.ShowDialog() <> DialogResult.OK Then Return
                _pendingPairs = form.Pairs
                scope = form.Scope
                doSaveHistory = form.SaveHistory
            End Using

            If _pendingPairs Is Nothing OrElse _pendingPairs.Count = 0 Then
                PostStatus("Không có cặp tìm/thay nào.")
                Return
            End If

            ' ── 4. Lưu history nếu được tick ──
            If doSaveHistory Then SaveHistory(_pendingPairs)

            ' ── 5. Xử lý theo phạm vi ──
            Select Case scope

                Case ApplyScope.SelectedObjects
                    If preSelectionCount > 0 Then
                        Dim targets As New List(Of Object)()
                        For i As Integer = 1 To preSelectionCount
                            Try : targets.Add(_oDrawDoc.SelectSet.Item(i)) : Catch : End Try
                        Next
                        ApplyAll(targets, _pendingPairs)
                        ClearSelectionAndUpdate()
                        _pendingPairs = Nothing
                    Else
                        StartInteractiveSelect()
                    End If

                Case ApplyScope.CurrentSheet
                    Dim sheet As Sheet = _oDrawDoc.ActiveSheet
                    Dim targets = CollectAllText(sheet)
                    ApplyAll(targets, _pendingPairs)
                    ClearSelectionAndUpdate()
                    _pendingPairs = Nothing
                    PostStatus($"Đã xử lý {targets.Count} đối tượng trên sheet '{sheet.Name}'.")

                Case ApplyScope.AllSheets
                    Dim allTargets As New List(Of Object)()
                    For Each sheet As Sheet In _oDrawDoc.Sheets
                        Dim t = CollectAllText(sheet)
                        allTargets.AddRange(t)
                    Next
                    ApplyAll(allTargets, _pendingPairs)
                    ClearSelectionAndUpdate()
                    _pendingPairs = Nothing
                    PostStatus($"Đã xử lý {allTargets.Count} đối tượng trên toàn bộ Drawing.")

            End Select
        End Sub


        Private Sub ClearSelectionAndUpdate()
            Try
                _oDrawDoc.SelectSet.Clear()
                _inventorApp.ActiveView.Update()
            Catch
            End Try
        End Sub


        Private Sub PostStatus(ByVal msg As String)
            Try
                _inventorApp.UserInterfaceManager.UserInteractionManager.PostStatus(msg)
            Catch
            End Try
        End Sub


        '═══════════════════════════════════════════════════════════
        ' THU THẬP TẤT CẢ ĐỐI TƯỢNG CÓ TEXT
        '
        ' ⭐ Dùng cách gọi TRỰC TIẾP sheet.WeldingSymbols (như code
        '     xóa đã chạy được trên Inventor 2025)
        '═══════════════════════════════════════════════════════════
        Private Function CollectAllText(ByVal sheet As Sheet) As List(Of Object)
            Dim result As New List(Of Object)()

            '--- DrawingNotes (GeneralNote, LeaderNote, HoleThreadNote) ---
            Try
                For Each obj As Object In sheet.DrawingNotes
                    If obj IsNot Nothing Then result.Add(obj)
                Next
            Catch
            End Try

            '--- DrawingDimensions ---
            Try
                For Each obj As Object In sheet.DrawingDimensions
                    If obj IsNot Nothing Then result.Add(obj)
                Next
            Catch
            End Try

            '--- SketchedSymbols ---
            Try
                For Each obj As Object In sheet.SketchedSymbols
                    If obj IsNot Nothing Then result.Add(obj)
                Next
            Catch
            End Try

            '═══════════════════════════════════════════════════════════
            ' ⭐ WELDING SYMBOLS — GỌI TRỰC TIẾP (không reflection)
            '═══════════════════════════════════════════════════════════
            Try
                Dim weldCol As Object = sheet.WeldingSymbols
                If weldCol IsNot Nothing Then
                    Dim cnt As Integer = 0
                    Try : cnt = CInt(weldCol.Count) : Catch : End Try

                    For i As Integer = 1 To cnt
                        Try
                            Dim ws As Object = weldCol.Item(i)
                            If ws IsNot Nothing Then
                                result.Add(ws)
                                PostStatus("  + WeldSymbol: " & ws.GetType().Name)
                            End If
                        Catch
                        End Try
                    Next
                End If
            Catch
            End Try

            '--- Sketches của sheet ---
            Try
                For Each sk As DrawingSketch In sheet.Sketches
                    Try
                        For Each tb As Inventor.TextBox In sk.TextBoxes
                            result.Add(tb)
                        Next
                    Catch
                    End Try
                Next
            Catch
            End Try

            '--- Views ---
            Try
                For Each view As DrawingView In sheet.DrawingViews

                    ' Label
                    Try
                        If view.Label IsNot Nothing Then result.Add(view.Label)
                    Catch
                    End Try

                    ' Notes trong view
                    Try
                        For Each obj As Object In view.DrawingNotes
                            If obj IsNot Nothing Then result.Add(obj)
                        Next
                    Catch
                    End Try

                    ' Dimensions trong view
                    Try
                        For Each obj As Object In view.DrawingDimensions
                            If obj IsNot Nothing Then result.Add(obj)
                        Next
                    Catch
                    End Try

                    ' ⭐ Welding symbols trong view — gọi trực tiếp
                    Try
                        Dim weldCol As Object = view.WeldingSymbols
                        If weldCol IsNot Nothing Then
                            Dim cnt As Integer = 0
                            Try : cnt = CInt(weldCol.Count) : Catch : End Try

                            For i As Integer = 1 To cnt
                                Try
                                    Dim ws As Object = weldCol.Item(i)
                                    If ws IsNot Nothing Then
                                        result.Add(ws)
                                        PostStatus("  + WeldSymbol (view): " & ws.GetType().Name)
                                    End If
                                Catch
                                End Try
                            Next
                        End If
                    Catch
                    End Try

                    ' Sketches trong view
                    Try
                        For Each sk As DrawingSketch In view.Sketches
                            Try
                                For Each tb As Inventor.TextBox In sk.TextBoxes
                                    result.Add(tb)
                                Next
                            Catch
                            End Try
                        Next
                    Catch
                    End Try
                Next
            Catch
            End Try

            Return result
        End Function


        '═══════════════════════════════════════════════════════════
        ' LỊCH SỬ - LƯU / ĐỌC / XÓA
        '═══════════════════════════════════════════════════════════
        Friend Function LoadHistory() As List(Of ReplacePair)
            Dim list As New List(Of ReplacePair)()
            Try
                If System.IO.File.Exists(HistoryFile) Then
                    For Each line In System.IO.File.ReadAllLines(HistoryFile, Encoding.UTF8)
                        If String.IsNullOrWhiteSpace(line) Then Continue For
                        Dim parts = line.Split(ControlChars.Tab)
                        If parts.Length >= 2 Then
                            list.Add(New ReplacePair(parts(0), parts(1)))
                        ElseIf parts.Length = 1 Then
                            list.Add(New ReplacePair(parts(0), ""))
                        End If
                    Next
                End If
            Catch
            End Try
            Return list
        End Function


        Friend Sub SaveHistory(ByVal pairs As List(Of ReplacePair))
            Try
                Dim dir As String = System.IO.Path.GetDirectoryName(HistoryFile)
                If Not System.IO.Directory.Exists(dir) Then System.IO.Directory.CreateDirectory(dir)

                Dim sb As New StringBuilder()
                For Each p In pairs
                    If String.IsNullOrEmpty(p.Find) Then Continue For
                    Dim f = If(p.Find, "").Replace(ControlChars.Tab, " "c)
                    Dim r = If(p.Replace, "").Replace(ControlChars.Tab, " "c)
                    sb.AppendLine(f & ControlChars.Tab & r)
                Next

                System.IO.File.WriteAllText(HistoryFile, sb.ToString(), Encoding.UTF8)
            Catch
            End Try
        End Sub


        Friend Sub ClearHistoryFile()
            Try
                If System.IO.File.Exists(HistoryFile) Then System.IO.File.Delete(HistoryFile)
            Catch
            End Try
        End Sub


        '═══════════════════════════════════════════════════════════
        ' BẮT ĐẦU CHẾ ĐỘ CHỌN TƯƠNG TÁC
        '═══════════════════════════════════════════════════════════
        Private Sub StartInteractiveSelect()
            _collectedObjects.Clear()
            Try : _oDrawDoc.SelectSet.Clear() : Catch : End Try

            Try
                _interactEvents = _inventorApp.CommandManager.CreateInteractionEvents()
                _interactEvents.InteractionDisabled = False

                _selectEvents = _interactEvents.SelectEvents
                _selectEvents.AddSelectionFilter(SelectionFilterEnum.kAllEntitiesFilter)
                _selectEvents.WindowSelectEnabled = True

                AddHandler _selectEvents.OnSelect, AddressOf HandleOnSelect
                AddHandler _interactEvents.OnTerminate, AddressOf HandleOnTerminate

                _interactEvents.Start()
            Catch ex As Exception
                MessageBox.Show("Không khởi động được chế độ chọn: " & ex.Message,
                                "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                _interactEvents = Nothing
                _selectEvents = Nothing
                Return
            End Try

            _interactEvents.StatusBarText = "Chọn đối tượng cần thay chữ. ESC để kết thúc."

            Do While _interactEvents IsNot Nothing
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(20)
            Loop
        End Sub


        Private Sub HandleOnSelect(ByVal JustSelectedEntities As ObjectsEnumerator,
                                   ByVal SelectionDevice As SelectionDeviceEnum,
                                   ByVal ModelPosition As Inventor.Point,
                                   ByVal ViewPosition As Point2d,
                                   ByVal View As Inventor.View)
            Try
                For Each obj As Object In JustSelectedEntities
                    If obj Is Nothing Then Continue For
                    If Not _collectedObjects.Contains(obj) Then
                        _collectedObjects.Add(obj)
                    End If
                Next
            Catch
            End Try
        End Sub


        Private Sub HandleOnTerminate()
            Try
                If _selectEvents IsNot Nothing Then
                    RemoveHandler _selectEvents.OnSelect, AddressOf HandleOnSelect
                End If
                If _interactEvents IsNot Nothing Then
                    RemoveHandler _interactEvents.OnTerminate, AddressOf HandleOnTerminate
                End If
            Catch
            End Try

            Try
                If _collectedObjects.Count > 0 Then
                    ApplyAll(_collectedObjects, _pendingPairs)
                    PostStatus($"Đã xử lý {_collectedObjects.Count} đối tượng.")
                Else
                    PostStatus("Không có đối tượng nào được chọn.")
                End If
            Catch
            End Try

            Try
                If _oDrawDoc IsNot Nothing Then
                    _oDrawDoc.SelectSet.Clear()
                    _inventorApp.ActiveView.Update()
                End If
            Catch
            End Try

            _collectedObjects.Clear()
            _pendingPairs = Nothing
            _selectEvents = Nothing
            _interactEvents = Nothing
        End Sub


        '═══════════════════════════════════════════════════════════
        ' ÁP DỤNG
        '═══════════════════════════════════════════════════════════
        Private Sub ApplyAll(ByVal targets As List(Of Object), ByVal pairs As List(Of ReplacePair))
            If targets Is Nothing OrElse pairs Is Nothing Then Return

            Dim applied As New HashSet(Of Object)()
            Dim count As Integer = 0

            For Each tObj In targets
                If applied.Contains(tObj) Then Continue For
                applied.Add(tObj)

                ' ⭐ Weld Symbol — xử lý riêng qua reflection
                If IsDrawingWeldingSymbol(tObj) Then
                    If ApplyReplaceToWeldSymbolReflect(tObj, pairs) Then count += 1
                    Continue For
                End If

                ' Các đối tượng thường
                Dim oldTxt As String = GetTextFromEntity(tObj)
                If String.IsNullOrEmpty(oldTxt) Then Continue For

                Dim newTxt As String = oldTxt
                For Each p In pairs
                    If String.IsNullOrEmpty(p.Find) Then Continue For
                    newTxt = ReplaceIgnoreCaseSkipTags(newTxt, p.Find, If(p.Replace, ""))
                Next

                If newTxt <> oldTxt Then
                    If SetTextToEntity(tObj, newTxt) Then count += 1
                End If
            Next

            Try : _inventorApp.ActiveView.Update() : Catch : End Try
        End Sub


        '═══════════════════════════════════════════════════════════
        ' WELD SYMBOL — Reflection
        '
        ' Vì API WeldSymbol chỉ có trên Inventor 2024+, dùng reflection
        ' để tương thích mọi phiên bản.
        '═══════════════════════════════════════════════════════════
        Private Function IsDrawingWeldingSymbol(ByVal obj As Object) As Boolean
            If obj Is Nothing Then Return False
            Try
                Return TypeOf obj Is DrawingWeldingSymbol
            Catch
                Try
                    Return obj.GetType().Name.IndexOf("Weld", StringComparison.OrdinalIgnoreCase) >= 0
                Catch
                End Try
            End Try
            Return False
        End Function

        Private Function ApplyReplaceToWeldSymbolReflect(ByVal ws As Object,
                                                 ByVal pairs As List(Of ReplacePair)) As Boolean
            If ws Is Nothing OrElse pairs Is Nothing Then Return False
            Dim anyChanged As Boolean = False

            Try
                ' Ép kiểu mạnh
                Dim symbol As DrawingWeldingSymbol = TryCast(ws, DrawingWeldingSymbol)
                If symbol Is Nothing Then Return False

                ' Symbol lấy từ model → không sửa được
                If symbol.Retrieved Then
                    PostStatus("WeldSymbol Retrieved — bỏ qua.")
                    Return False
                End If

                Dim defs As DrawingWeldingSymbolDefinitions = symbol.Definitions
                If defs Is Nothing OrElse defs.Count < 1 Then Return False

                For i As Integer = 1 To defs.Count
                    Try
                        Dim def As DrawingWeldingSymbolDefinition = defs.Item(i)
                        If def Is Nothing Then Continue For

                        Dim cur As String = ""
                        Try : cur = def.TailNote : Catch : Continue For : End Try
                        If String.IsNullOrEmpty(cur) Then Continue For

                        Dim updated As String = cur
                        For Each p In pairs
                            If String.IsNullOrEmpty(p.Find) Then Continue For
                            updated = ReplaceIgnoreCaseSkipTags(updated, p.Find, If(p.Replace, ""))
                        Next

                        If updated <> cur Then
                            def.TailNote = updated
                            anyChanged = True
                            PostStatus("TailNote: """ & cur & """ → """ & updated & """")
                        End If
                    Catch
                    End Try
                Next
            Catch ex As Exception
                PostStatus("Lỗi WeldSymbol: " & ex.Message)
            End Try

            Return anyChanged
        End Function


        '═══════════════════════════════════════════════════════════
        ' ĐỌC / GHI TEXT — các đối tượng chuẩn
        '═══════════════════════════════════════════════════════════
        Private Function GetTextFromEntity(ByVal obj As Object) As String
            Try
                If TypeOf obj Is GeneralNote Then Return CType(obj, GeneralNote).FormattedText
                If TypeOf obj Is LeaderNote Then Return CType(obj, LeaderNote).FormattedText
                If TypeOf obj Is Inventor.TextBox Then Return CType(obj, Inventor.TextBox).FormattedText

                If TypeOf obj Is DrawingDimension Then
                    Return CType(obj, DrawingDimension).Text.FormattedText
                End If

                If TypeOf obj Is DrawingViewLabel Then
                    Return CType(obj, DrawingViewLabel).FormattedText
                End If

                If TypeOf obj Is ModelGeneralNote Then
                    Return CType(obj, ModelGeneralNote).Definition.Text.FormattedText
                End If
                If TypeOf obj Is ModelLeaderNote Then
                    Return CType(obj, ModelLeaderNote).Definition.Text.FormattedText
                End If
            Catch
            End Try
            Return Nothing
        End Function


        Private Function SetTextToEntity(ByVal obj As Object, ByVal newText As String) As Boolean
            Try
                If TypeOf obj Is GeneralNote Then
                    CType(obj, GeneralNote).FormattedText = newText
                    Return True
                End If
                If TypeOf obj Is LeaderNote Then
                    CType(obj, LeaderNote).FormattedText = newText
                    Return True
                End If
                If TypeOf obj Is Inventor.TextBox Then
                    CType(obj, Inventor.TextBox).FormattedText = newText
                    Return True
                End If
                If TypeOf obj Is DrawingDimension Then
                    CType(obj, DrawingDimension).Text.FormattedText = newText
                    Return True
                End If
                If TypeOf obj Is DrawingViewLabel Then
                    CType(obj, DrawingViewLabel).FormattedText = newText
                    Return True
                End If
                If TypeOf obj Is ModelGeneralNote Then
                    CType(obj, ModelGeneralNote).Definition.Text.FormattedText = newText
                    Return True
                End If
                If TypeOf obj Is ModelLeaderNote Then
                    CType(obj, ModelLeaderNote).Definition.Text.FormattedText = newText
                    Return True
                End If
            Catch
            End Try
            Return False
        End Function


        '═══════════════════════════════════════════════════════════
        ' REPLACE AN TOÀN — bỏ qua mọi tag <...>
        '═══════════════════════════════════════════════════════════
        Private Function ReplaceIgnoreCaseSkipTags(ByVal s As String,
                                                   ByVal find As String,
                                                   ByVal replace As String) As String
            If String.IsNullOrEmpty(s) OrElse String.IsNullOrEmpty(find) Then Return s

            Dim sb As New StringBuilder(s.Length)
            Dim i As Integer = 0

            While i < s.Length
                If s(i) = "<"c Then
                    Dim j As Integer = s.IndexOf(">"c, i)
                    If j < 0 Then
                        sb.Append(s.Substring(i))
                        Exit While
                    End If
                    sb.Append(s.Substring(i, j - i + 1))
                    i = j + 1
                Else
                    Dim j As Integer = s.IndexOf("<"c, i)
                    Dim chunkEnd As Integer = If(j < 0, s.Length, j)
                    sb.Append(ReplaceIgnoreCase(s.Substring(i, chunkEnd - i), find, replace))
                    i = chunkEnd
                End If
            End While

            Return sb.ToString()
        End Function


        Private Function ReplaceIgnoreCase(ByVal str As String,
                                           ByVal oldValue As String,
                                           ByVal newValue As String) As String
            If String.IsNullOrEmpty(oldValue) OrElse String.IsNullOrEmpty(str) Then Return str

            Dim sb As New StringBuilder()
            Dim i As Integer = 0
            While i < str.Length
                If i + oldValue.Length <= str.Length AndAlso
                   String.Compare(str, i, oldValue, 0, oldValue.Length,
                                  StringComparison.OrdinalIgnoreCase) = 0 Then
                    sb.Append(newValue)
                    i += oldValue.Length
                Else
                    sb.Append(str(i))
                    i += 1
                End If
            End While
            Return sb.ToString()
        End Function


        '═══════════════════════════════════════════════════════════
        ' DEBUG — chạy để xem type name + property của weld symbol
        '═══════════════════════════════════════════════════════════
        Public Sub DebugWeldSymbols()
            Try
                Dim app As Inventor.Application =
                    CType(Interop.Marshal2.GetActiveObject("Inventor.Application"),
                          Inventor.Application)

                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Mở Drawing trước.")
                    Return
                End If

                Dim oDrawDoc As DrawingDocument = CType(app.ActiveDocument, DrawingDocument)
                Dim sheet As Sheet = oDrawDoc.ActiveSheet

                Dim sb As New StringBuilder()

                '--- Sheet.WeldingSymbols (gọi trực tiếp) ---
                Try
                    Dim weldCol As Object = sheet.WeldingSymbols
                    If weldCol Is Nothing Then
                        sb.AppendLine("Sheet.WeldingSymbols = Nothing")
                    Else
                        sb.AppendLine("Sheet.WeldingSymbols.Count = " & CStr(weldCol.Count))
                        For i As Integer = 1 To CInt(weldCol.Count)
                            Try
                                Dim ws As Object = weldCol.Item(i)
                                sb.AppendLine()
                                sb.AppendLine("Item " & i & " — Type: " & ws.GetType().FullName)

                                For Each pi As System.Reflection.PropertyInfo In ws.GetType().GetProperties()
                                    Try
                                        If Not pi.CanRead Then Continue For
                                        Dim v As Object = Nothing
                                        Try : v = pi.GetValue(ws, Nothing) : Catch : End Try
                                        If v Is Nothing Then Continue For

                                        Dim typeStr As String = v.GetType().Name
                                        sb.AppendLine("  ." & pi.Name & " (" & typeStr & ") = " &
                                                      If(pi.PropertyType Is GetType(String), """" & CStr(v) & """",
                                                         If(typeStr = "Object", "<obj>", v.ToString())))
                                    Catch
                                    End Try
                                Next
                            Catch ex As Exception
                                sb.AppendLine("  Item " & i & " lỗi: " & ex.Message)
                            End Try
                        Next
                    End If
                Catch ex As Exception
                    sb.AppendLine("Lỗi truy cập WeldingSymbols: " & ex.Message)
                End Try

                MessageBox.Show(sb.ToString(), "DEBUG WeldSymbols",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

    End Module


    '═══════════════════════════════════════════════════════════
    ' Model
    '═══════════════════════════════════════════════════════════
    Public Class ReplacePair
        Public Property Find As String
        Public Property Replace As String
        Public Sub New(ByVal f As String, ByVal r As String)
            Me.Find = f
            Me.Replace = r
        End Sub
    End Class


    '═══════════════════════════════════════════════════════════
    ' FORM — ĐẸP, ĐỒNG BỘ VỚI CÁC FORM KHÁC
    '═══════════════════════════════════════════════════════════
    Public Class ReplaceForm
        Inherits Form

        Private dgv As DataGridView
        Private btnAdd As Button
        Private btnRemove As Button
        Private btnClearHistory As Button
        Private btnOK As Button
        Private btnCancel As Button

        Private rbSelected As RadioButton
        Private rbCurrentSheet As RadioButton
        Private rbAllSheets As RadioButton
        Private chkSaveHistory As CheckBox

        Public Property Pairs As List(Of ReplacePair)
        Public Property Scope As ApplyScope
        Public Property SaveHistory As Boolean
        Private _okConfirmed As Boolean = False

        Private Const FORM_W As Integer = 720
        Private Const FORM_H As Integer = 690

        Public Sub New(ByVal preSelectionCount As Integer, ByVal history As List(Of ReplacePair))
            Me.Text = "Thay chữ trong Text"
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
            lblTitle.Text = "THAY CHỮ TRONG TEXT"
            lblTitle.Font = New Drw.Font("Segoe UI", 15.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblTitle.ForeColor = Drw.Color.White
            lblTitle.Dock = DockStyle.Fill
            lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
            pnlHeader.Controls.Add(lblTitle)

            Dim lblSub As New Label()
            lblSub.Text = "Tìm & thay thế chữ trong Drawing Note — hỗ trợ Welding Symbol"
            lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
            lblSub.Dock = DockStyle.Bottom
            lblSub.Height = 20
            lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
            pnlHeader.Controls.Add(lblSub)

            '===== GROUP 1: PHẠM VI =====
            Dim gbScope As New GroupBox() With {
                .Text = "1. Phạm vi áp dụng",
                .Location = New Drw.Point(15, 90),
                .Size = New Drw.Size(FORM_W - 30, 130),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gbScope)

            rbSelected = New RadioButton() With {
                .Text = "Đối tượng chọn trên bản vẽ",
                .Location = New Drw.Point(25, 30),
                .Size = New Drw.Size(600, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            rbCurrentSheet = New RadioButton() With {
                .Text = "Toàn bộ sheet hiện tại",
                .Location = New Drw.Point(25, 62),
                .Size = New Drw.Size(600, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            rbAllSheets = New RadioButton() With {
                .Text = "Toàn bộ tất cả sheet trong bản vẽ",
                .Location = New Drw.Point(25, 94),
                .Size = New Drw.Size(600, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}

            gbScope.Controls.AddRange(New Control() {rbSelected, rbCurrentSheet, rbAllSheets})

            If preSelectionCount > 0 Then
                rbSelected.Checked = True
                rbSelected.Text = "Đối tượng chọn trên bản vẽ  (đang có " & preSelectionCount.ToString() & " đối tượng)"
            Else
                rbCurrentSheet.Checked = True
                rbSelected.Text = "Đối tượng chọn trên bản vẽ  (chưa có — sẽ chọn sau khi OK)"
            End If

            '===== GROUP 2: CẶP TÌM / THAY =====
            Dim gbPairs As New GroupBox() With {
                .Text = "2. Các cặp TÌM / THAY THẾ",
                .Location = New Drw.Point(15, 230),
                .Size = New Drw.Size(FORM_W - 30, 340),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gbPairs)

            Dim lblHint As New Label() With {
                .Text = "Cột 'Thay bằng' để trống = XÓA cụm từ đó  |  Không phân biệt hoa thường",
                .Location = New Drw.Point(20, 25),
                .Size = New Drw.Size(650, 20),
                .ForeColor = Drw.Color.FromArgb(140, 140, 140),
                .Font = New Drw.Font("Segoe UI", 8.5F, Drw.FontStyle.Italic, Drw.GraphicsUnit.Point)}
            gbPairs.Controls.Add(lblHint)

            '--- DataGridView ---
            dgv = New DataGridView With {
                .Location = New Drw.Point(20, 52),
                .Size = New Drw.Size(650, 230),
                .AllowUserToAddRows = False,
                .AllowUserToDeleteRows = False,
                .AllowUserToResizeRows = False,
                .RowHeadersVisible = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                .EditMode = DataGridViewEditMode.EditOnEnter,
                .SelectionMode = DataGridViewSelectionMode.CellSelect,
                .MultiSelect = False,
                .BackgroundColor = Drw.Color.White,
                .BorderStyle = BorderStyle.FixedSingle,
                .GridColor = Drw.Color.FromArgb(220, 220, 220),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}

            dgv.ColumnHeadersDefaultCellStyle.BackColor = Drw.Color.FromArgb(45, 100, 180)
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Drw.Color.White
            dgv.ColumnHeadersDefaultCellStyle.Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            dgv.ColumnHeadersHeight = 30
            dgv.EnableHeadersVisualStyles = False
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Drw.Color.FromArgb(248, 250, 253)
            dgv.DefaultCellStyle.SelectionBackColor = Drw.Color.FromArgb(200, 220, 245)
            dgv.DefaultCellStyle.SelectionForeColor = Drw.Color.FromArgb(30, 30, 30)

            dgv.Columns.Add(New DataGridViewTextBoxColumn With {
                .Name = "colFind", .HeaderText = "Tìm", .FillWeight = 50
            })
            dgv.Columns.Add(New DataGridViewTextBoxColumn With {
                .Name = "colReplace", .HeaderText = "Thay bằng  (để trống = xóa)", .FillWeight = 50
            })

            If history IsNot Nothing AndAlso history.Count > 0 Then
                For Each p In history
                    dgv.Rows.Add(p.Find, p.Replace)
                Next
            Else
                dgv.Rows.Add("", "")
            End If

            gbPairs.Controls.Add(dgv)

            '--- Nút Thêm / Xóa dòng ---
            btnAdd = New Button() With {
                .Text = "Thêm dòng",
                .Location = New Drw.Point(20, 292),
                .Size = New Drw.Size(110, 30),
                .FlatStyle = FlatStyle.Flat,
                .BackColor = Drw.Color.White,
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point),
                .Cursor = Cursors.Hand}
            btnAdd.FlatAppearance.BorderColor = Drw.Color.FromArgb(45, 100, 180)
            gbPairs.Controls.Add(btnAdd)

            btnRemove = New Button() With {
                .Text = "Xóa dòng",
                .Location = New Drw.Point(140, 292),
                .Size = New Drw.Size(110, 30),
                .FlatStyle = FlatStyle.Flat,
                .BackColor = Drw.Color.White,
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point),
                .Cursor = Cursors.Hand}
            btnRemove.FlatAppearance.BorderColor = Drw.Color.FromArgb(45, 100, 180)
            gbPairs.Controls.Add(btnRemove)

            '--- Checkbox lưu history ---
            chkSaveHistory = New CheckBox() With {
                .Text = "Lưu các cặp này cho lần sau",
                .Location = New Drw.Point(280, 296),
                .Size = New Drw.Size(260, 22),
                .Checked = True,
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            gbPairs.Controls.Add(chkSaveHistory)

            '--- Nút xóa lịch sử ---
            btnClearHistory = New Button() With {
                .Text = "Xóa lịch sử",
                .Location = New Drw.Point(548, 292),
                .Size = New Drw.Size(122, 30),
                .FlatStyle = FlatStyle.Flat,
                .BackColor = Drw.Color.FromArgb(250, 250, 250),
                .ForeColor = Drw.Color.FromArgb(180, 60, 60),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point),
                .Cursor = Cursors.Hand}
            btnClearHistory.FlatAppearance.BorderColor = Drw.Color.FromArgb(200, 200, 200)
            gbPairs.Controls.Add(btnClearHistory)

            '===== NÚT THỰC HIỆN =====
            btnOK = New Button()
            btnOK.Text = "THỰC HIỆN"
            btnOK.Size = New Drw.Size(160, 46)
            btnOK.Location = New Drw.Point(FORM_W - 185, FORM_H - 60)
            btnOK.FlatStyle = FlatStyle.Flat
            btnOK.FlatAppearance.BorderSize = 0
            btnOK.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(60, 115, 195)
            btnOK.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(30, 80, 155)
            btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
            btnOK.ForeColor = Drw.Color.White
            btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            btnOK.Cursor = Cursors.Hand
            btnOK.UseVisualStyleBackColor = False
            Me.Controls.Add(btnOK)

            '===== NÚT HỦY =====
            btnCancel = New Button()
            btnCancel.Text = "HỦY"
            btnCancel.Size = New Drw.Size(130, 46)
            btnCancel.Location = New Drw.Point(FORM_W - 330, FORM_H - 60)
            btnCancel.FlatStyle = FlatStyle.Flat
            btnCancel.FlatAppearance.BorderSize = 1
            btnCancel.FlatAppearance.BorderColor = Drw.Color.FromArgb(200, 200, 200)
            btnCancel.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(235, 235, 235)
            btnCancel.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(215, 215, 215)
            btnCancel.BackColor = Drw.Color.FromArgb(250, 250, 250)
            btnCancel.ForeColor = Drw.Color.FromArgb(60, 60, 60)
            btnCancel.Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            btnCancel.Cursor = Cursors.Hand
            btnCancel.UseVisualStyleBackColor = False
            Me.Controls.Add(btnCancel)

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel

            '===== SỰ KIỆN =====
            AddHandler btnAdd.Click, Sub()
                                         Dim idx As Integer = dgv.Rows.Add("", "")
                                         dgv.CurrentCell = dgv.Rows(idx).Cells(0)
                                         dgv.BeginEdit(True)
                                     End Sub

            AddHandler btnRemove.Click, Sub()
                                            If dgv.CurrentRow IsNot Nothing Then dgv.Rows.Remove(dgv.CurrentRow)
                                            If dgv.Rows.Count = 0 Then dgv.Rows.Add("", "")
                                        End Sub

            AddHandler btnClearHistory.Click, Sub()
                                                  If MessageBox.Show(
                                                      "Xóa toàn bộ lịch sử đã lưu?" & vbCrLf &
                                                      "(Các cặp trong bảng hiện tại cũng sẽ bị xóa khỏi file)",
                                                      "Xác nhận",
                                                      MessageBoxButtons.YesNo,
                                                      MessageBoxIcon.Question) = DialogResult.Yes Then
                                                      ThayChuTrongTextModule.ClearHistoryFile()
                                                      dgv.Rows.Clear()
                                                      dgv.Rows.Add("", "")
                                                      MessageBox.Show("Đã xóa lịch sử.",
                                                                      "OK",
                                                                      MessageBoxButtons.OK,
                                                                      MessageBoxIcon.Information)
                                                  End If
                                              End Sub

            AddHandler btnOK.Click, Sub()
                                        dgv.EndEdit()

                                        Pairs = New List(Of ReplacePair)()
                                        For Each row As DataGridViewRow In dgv.Rows
                                            Dim f As String = If(row.Cells(0).Value?.ToString(), "")
                                            Dim r As String = If(row.Cells(1).Value?.ToString(), "")
                                            If Not String.IsNullOrEmpty(f) Then
                                                Pairs.Add(New ReplacePair(f, r))
                                            End If
                                        Next

                                        If Pairs.Count = 0 Then
                                            MessageBox.Show("Vui lòng nhập ít nhất 1 cặp TÌM/THAY.",
                                                            "Chưa có dữ liệu",
                                                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                            Return
                                        End If

                                        If rbSelected.Checked Then
                                            Scope = ApplyScope.SelectedObjects
                                        ElseIf rbCurrentSheet.Checked Then
                                            Scope = ApplyScope.CurrentSheet
                                        Else
                                            Scope = ApplyScope.AllSheets
                                        End If

                                        SaveHistory = chkSaveHistory.Checked
                                        _okConfirmed = True
                                        Me.DialogResult = DialogResult.OK
                                        Me.Close()
                                    End Sub

            AddHandler Me.FormClosing, Sub(sender, e)
                                           If Not _okConfirmed Then
                                               Me.DialogResult = DialogResult.Cancel
                                           End If
                                       End Sub
        End Sub
    End Class
End Namespace