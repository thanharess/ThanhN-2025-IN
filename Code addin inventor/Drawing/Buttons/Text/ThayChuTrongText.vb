Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports Inventor

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
        ' MỚI — thêm System. / System.IO.
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

            ' ── 1. Lấy Inventor ──
            Try
                _inventorApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch ex As Exception
                Console.WriteLine("Không tìm thấy Inventor: " & ex.Message)
                Return
            End Try

            If _inventorApp.ActiveDocument Is Nothing OrElse
               _inventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                Console.WriteLine("Vui lòng mở một bản vẽ (IDW/DWG) trước.")
                Return
            End If

            _oDrawDoc = CType(_inventorApp.ActiveDocument, DrawingDocument)

            ' ── 2. Đọc lịch sử + Đếm selection có sẵn ──
            Dim history = LoadHistory()
            Dim preSelectionCount As Integer = 0
            Try : preSelectionCount = _oDrawDoc.SelectSet.Count : Catch : End Try

            Console.WriteLine($"→ Có {preSelectionCount} đối tượng được chọn sẵn.")
            Console.WriteLine($"→ Lịch sử: {history.Count} cặp đã lưu.")

            ' ── 3. Mở form ──
            Dim scope As ApplyScope
            Dim doSaveHistory As Boolean

            Using form As New ReplaceForm(preSelectionCount, history)
                If form.ShowDialog() <> DialogResult.OK Then
                    Console.WriteLine("✖ Đã hủy lệnh.")
                    Return
                End If
                _pendingPairs = form.Pairs
                scope = form.Scope
                doSaveHistory = form.SaveHistory
            End Using

            If _pendingPairs Is Nothing OrElse _pendingPairs.Count = 0 Then
                Console.WriteLine("⚠️ Không có cặp tìm/thay nào.")
                Return
            End If

            ' ── 4. Lưu history nếu được tick ──
            If doSaveHistory Then
                SaveHistory(_pendingPairs)
            End If

            Console.WriteLine("=== THAYCHUTRONGTEXT ===")
            For Each p In _pendingPairs
                Console.WriteLine($"  • ""{p.Find}""  →  ""{p.Replace}""")
            Next
            Console.WriteLine($"  Phạm vi: {scope}")

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
                        ' ⚠️ OnExecute sẽ chờ trong DoEvents loop cho đến khi user ESC
                    End If

                Case ApplyScope.CurrentSheet
                    Dim sheet As Sheet = _oDrawDoc.ActiveSheet
                    Console.WriteLine($"▶ Áp dụng cho sheet: {sheet.Name}")
                    Dim targets = CollectAllText(sheet)
                    Console.WriteLine($"  Tìm thấy {targets.Count} đối tượng có text.")
                    ApplyAll(targets, _pendingPairs)
                    ClearSelectionAndUpdate()
                    _pendingPairs = Nothing

                Case ApplyScope.AllSheets
                    Console.WriteLine("▶ Áp dụng cho TẤT CẢ sheet...")
                    Dim allTargets As New List(Of Object)()
                    For Each sheet As Sheet In _oDrawDoc.Sheets
                        Console.WriteLine($"  ─ Sheet: {sheet.Name}")
                        Dim t = CollectAllText(sheet)
                        Console.WriteLine($"    {t.Count} đối tượng.")
                        allTargets.AddRange(t)
                    Next
                    Console.WriteLine($"  Tổng: {allTargets.Count} đối tượng.")
                    ApplyAll(allTargets, _pendingPairs)
                    ClearSelectionAndUpdate()
                    _pendingPairs = Nothing

            End Select
        End Sub


        Private Sub ClearSelectionAndUpdate()
            Try
                _oDrawDoc.SelectSet.Clear()
                _inventorApp.ActiveView.Update()
            Catch
            End Try
        End Sub


        '═══════════════════════════════════════════════════════════
        ' Thu thập tất cả đối tượng có text trong 1 sheet
        '═══════════════════════════════════════════════════════════
        Private Function CollectAllText(ByVal sheet As Sheet) As List(Of Object)
            Dim result As New List(Of Object)()

            ' Notes trực tiếp trên sheet
            Try
                For Each obj As Object In sheet.DrawingNotes
                    result.Add(obj)
                    If IsDrawingWeldingSymbol(obj) Then
                        Console.WriteLine("    + WeldSymbol (trong DrawingNotes)")
                    End If
                Next
            Catch
            End Try

            ' Dimensions trực tiếp trên sheet
            Try
                For Each obj As Object In sheet.DrawingDimensions
                    result.Add(obj)
                Next
            Catch
            End Try

            ' Sketched Symbols
            Try
                For Each obj As Object In sheet.SketchedSymbols
                    result.Add(obj)
                Next
            Catch
            End Try

            ' ─── Welding symbols (chỉ có từ Inventor 2024+) ───
            ' Dùng reflection để tránh lỗi build trên Inventor 2020
            Try
                Dim pi = sheet.GetType().GetProperty("WeldingSymbols")
                If pi IsNot Nothing Then
                    Dim col As Object = pi.GetValue(sheet, Nothing)
                    If col IsNot Nothing Then
                        For Each obj As Object In CType(col, IEnumerable)
                            If obj IsNot Nothing Then
                                result.Add(obj)
                                Console.WriteLine("    + DrawingWeldingSymbol")
                            End If
                        Next
                    End If
                End If
            Catch
            End Try

            ' Sketches của sheet
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

            ' Views và nội dung bên trong
            Try
                For Each view As DrawingView In sheet.DrawingViews

                    ' Label của view
                    Try
                        If view.Label IsNot Nothing Then result.Add(view.Label)
                    Catch
                    End Try

                    ' Notes trong view
                    Try
                        For Each obj As Object In view.DrawingNotes
                            result.Add(obj)
                            If IsDrawingWeldingSymbol(obj) Then              ' ← ĐÚNG
                                Console.WriteLine("      + WeldSymbol (view)")
                            End If
                        Next
                    Catch
                    End Try

                    ' Dimensions trong view
                    Try
                        For Each obj As Object In view.DrawingDimensions
                            result.Add(obj)
                        Next
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
            Catch ex As Exception
                Console.WriteLine("  ⚠️ Load history: " & ex.Message)
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
                Console.WriteLine($"  ✔ Đã lưu {pairs.Count} cặp vào: {HistoryFile}")
            Catch ex As Exception
                Console.WriteLine("  ⚠️ Save history: " & ex.Message)
            End Try
        End Sub


        Friend Sub ClearHistoryFile()
            Try
                If System.IO.File.Exists(HistoryFile) Then System.IO.File.Delete(HistoryFile)
                Console.WriteLine("  ✔ Đã xóa file lịch sử.")
            Catch ex As Exception
                Console.WriteLine("  ⚠️ Clear history: " & ex.Message)
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

            Console.WriteLine("─────────────────────────────────────────────────")
            Console.WriteLine("▶ Chọn đối tượng bằng cách CLICK hoặc QUÉT CHUỘT.")
            Console.WriteLine("▶ Nhấn ESC (hoặc chuột phải → Done) khi chọn xong.")
            Console.WriteLine("─────────────────────────────────────────────────")

            ' Giữ OnExecute không return → Inventor không kill InteractionEvents
            Do While _interactEvents IsNot Nothing
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(20)
            Loop

            Console.WriteLine("▶ Đã kết thúc chế độ chọn.")
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
                        Console.WriteLine($"  + {obj.Type}")
                    End If
                Next
            Catch ex As Exception
                Console.WriteLine("  ⚠️ OnSelect: " & ex.Message)
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
                    Console.WriteLine($"Đã chọn {_collectedObjects.Count} đối tượng. Đang áp dụng...")
                    ApplyAll(_collectedObjects, _pendingPairs)
                Else
                    Console.WriteLine("⚠️ Không có đối tượng nào được chọn.")
                End If
            Catch ex As Exception
                Console.WriteLine("  ⚠️ Apply: " & ex.Message)
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
            _interactEvents = Nothing      ' ← Cờ để DoEvents loop thoát
        End Sub


        '═══════════════════════════════════════════════════════════
        ' ÁP DỤNG
        '═══════════════════════════════════════════════════════════
        Private Sub ApplyAll(ByVal targets As List(Of Object), ByVal pairs As List(Of ReplacePair))
            If targets Is Nothing OrElse pairs Is Nothing Then Return

            Dim applied As New HashSet(Of Object)()
            Dim count As Integer = 0
            Dim skipped As Integer = 0

            For Each tObj In targets
                If applied.Contains(tObj) Then Continue For
                applied.Add(tObj)

                ' ─── Nhánh Weld Symbol (Reflection — an toàn cho mọi phiên bản) ───
                If IsDrawingWeldingSymbol(tObj) Then
                    If ApplyReplaceToWeldSymbolReflect(tObj, pairs) Then count += 1
                    Continue For
                End If

                ' ─── Các đối tượng thường ───
                Dim oldTxt As String = GetTextFromEntity(tObj)
                If String.IsNullOrEmpty(oldTxt) Then
                    skipped += 1
                    Continue For
                End If

                Dim newTxt As String = oldTxt
                For Each p In pairs
                    If String.IsNullOrEmpty(p.Find) Then Continue For
                    newTxt = ReplaceIgnoreCaseSkipTags(newTxt, p.Find, If(p.Replace, ""))
                Next

                If newTxt <> oldTxt Then
                    If SetTextToEntity(tObj, newTxt) Then count += 1
                End If
            Next

            Console.WriteLine($"✅ Hoàn tất – đã xử lý {count} đối tượng (bỏ qua {skipped} đối tượng không có text).")

            Try
                _inventorApp.ActiveView.Update()
            Catch
            End Try
        End Sub

        '═══════════════════════════════════════════════════════════
        ' XỬ LÝ WELD SYMBOL — bằng Reflection, tương thích cả 2020 & 2024+
        '
        ' Type drawing welding symbol là "Inventor.DrawingWeldingSymbol" (2024+).
        ' Trên Inventor 2020, type này không tồn tại → hàm này không bao giờ được gọi.
        '═══════════════════════════════════════════════════════════
        Private Function IsDrawingWeldingSymbol(ByVal obj As Object) As Boolean
            If obj Is Nothing Then Return False
            Try
                Dim tn As String = obj.GetType().Name
                Return tn = "DrawingWeldingSymbol" OrElse tn = "WeldSymbol"
            Catch
            End Try
            Return False
        End Function


        Private Function ApplyReplaceToWeldSymbolReflect(ByVal ws As Object,
                                                  ByVal pairs As List(Of ReplacePair)) As Boolean
            Dim anyChanged As Boolean = False

            Dim props As String() = {
        "LeftText", "RightText", "TopText", "BottomText",
        "ContourText", "TailText", "FieldWeldText", "AllAroundText"
    }

            For Each propName As String In props
                Try
                    Dim pi = ws.GetType().GetProperty(propName)
                    If pi Is Nothing Then Continue For

                    Dim wst As Object = pi.GetValue(ws, Nothing)
                    If wst Is Nothing Then Continue For

                    ' Thử FormattedText, fallback Text
                    Dim ftPi = wst.GetType().GetProperty("FormattedText")
                    If ftPi Is Nothing Then ftPi = wst.GetType().GetProperty("Text")
                    If ftPi Is Nothing Then Continue For
                    If Not ftPi.CanWrite Then Continue For

                    Dim cur As String = TryCast(ftPi.GetValue(wst, Nothing), String)
                    If String.IsNullOrEmpty(cur) Then Continue For

                    Dim updated As String = cur
                    For Each p In pairs
                        If String.IsNullOrEmpty(p.Find) Then Continue For
                        updated = ReplaceIgnoreCaseSkipTags(updated, p.Find, If(p.Replace, ""))
                    Next

                    If updated <> cur Then
                        ftPi.SetValue(wst, updated, Nothing)
                        Console.WriteLine($"    • WeldSymbol.{propName}: ""{cur}"" → ""{updated}""")
                        anyChanged = True
                    End If
                Catch ex As Exception
                    Console.WriteLine($"    ⚠️ WeldSymbol.{propName}: {ex.Message}")
                End Try
            Next

            Return anyChanged
        End Function







        '═══════════════════════════════════════════════════════════
        ' ĐỌC / GHI TEXT
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
            Catch ex As Exception
                Console.WriteLine("  ⚠️ " & ex.Message)
            End Try
            Return False
        End Function


        '═══════════════════════════════════════════════════════════
        ' REPLACE AN TOÀN — bỏ qua mọi tag <...>
        '═══════════════════════════════════════════════════════════
        Private Function ReplaceIgnoreCaseSkipTags(ByVal s As String, ByVal find As String, ByVal replace As String) As String
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


        Private Function ReplaceIgnoreCase(ByVal str As String, ByVal oldValue As String, ByVal newValue As String) As String
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
    ' FORM
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

        Public Sub New(ByVal preSelectionCount As Integer, ByVal history As List(Of ReplacePair))
            Me.Text = "Thay chữ trong Text (Inventor)"
            Me.ClientSize = New Size(650, 570)
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Font = New Font("Segoe UI", 9.0F)

            ' ── Hint ──
            Dim lblHint As New Label With {
                .Text = "• Cột 'Thay bằng' để trống = XÓA cụm từ đó" & vbCrLf &
                        "• Không phân biệt hoa thường, an toàn với tag XML",
                .Left = 12, .Top = 10, .Width = 626, .Height = 34,
                .ForeColor = System.Drawing.Color.DimGray
            }

            ' ── Scope ──
            Dim grpScope As New GroupBox With {
                .Text = "Phạm vi áp dụng",
                .Left = 12, .Top = 52, .Width = 626, .Height = 104
            }
            rbSelected = New RadioButton With {
                .Text = "Đối tượng chọn trên bản vẽ",
                .Left = 14, .Top = 22, .Width = 600
            }
            rbCurrentSheet = New RadioButton With {
                .Text = "Toàn bộ sheet hiện tại",
                .Left = 14, .Top = 46, .Width = 600
            }
            rbAllSheets = New RadioButton With {
                .Text = "Toàn bộ tất cả sheet trong bản vẽ",
                .Left = 14, .Top = 70, .Width = 600
            }
            grpScope.Controls.AddRange(New Control() {rbSelected, rbCurrentSheet, rbAllSheets})

            If preSelectionCount > 0 Then
                rbSelected.Checked = True
                rbSelected.Text = $"Đối tượng chọn trên bản vẽ (đang có {preSelectionCount} đối tượng)"
            Else
                rbCurrentSheet.Checked = True
                rbSelected.Text = "Đối tượng chọn trên bản vẽ (chưa có — sẽ chọn sau khi OK)"
            End If

            ' ── Grid label ──
            Dim lblGrid As New Label With {
                .Text = "Các cặp TÌM / THAY THẾ:",
                .Left = 12, .Top = 164, .Width = 626, .Height = 18,
                .Font = New Font(Me.Font, FontStyle.Bold)
            }

            ' ── Grid ──
            dgv = New DataGridView With {
                .Left = 12, .Top = 186, .Width = 626, .Height = 270,
                .AllowUserToAddRows = False,
                .AllowUserToDeleteRows = False,
                .AllowUserToResizeRows = False,
                .RowHeadersVisible = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                .EditMode = DataGridViewEditMode.EditOnEnter,
                .SelectionMode = DataGridViewSelectionMode.CellSelect,
                .MultiSelect = False
            }
            dgv.Columns.Add(New DataGridViewTextBoxColumn With {
                .Name = "colFind", .HeaderText = "Tìm", .FillWeight = 50
            })
            dgv.Columns.Add(New DataGridViewTextBoxColumn With {
                .Name = "colReplace", .HeaderText = "Thay bằng  (để trống = xóa)", .FillWeight = 50
            })

            ' Nạp history hoặc 1 dòng trống
            If history IsNot Nothing AndAlso history.Count > 0 Then
                For Each p In history
                    dgv.Rows.Add(p.Find, p.Replace)
                Next
            Else
                dgv.Rows.Add("", "")
            End If

            ' ── Buttons row 1 ──
            btnAdd = New Button With {.Text = "Thêm dòng", .Left = 12, .Top = 462, .Width = 110, .Height = 28}
            btnRemove = New Button With {.Text = "Xóa dòng", .Left = 130, .Top = 462, .Width = 110, .Height = 28}
            chkSaveHistory = New CheckBox With {
                .Text = "Lưu các cặp này cho lần sau",
                .Left = 260, .Top = 466, .Width = 380, .Height = 22,
                .Checked = True
            }

            ' ── Buttons row 2 ──
            btnClearHistory = New Button With {
                .Text = "Xóa lịch sử đã lưu",
                .Left = 12, .Top = 502, .Width = 180, .Height = 28
            }
            btnOK = New Button With {
                .Text = "OK", .Left = 446, .Top = 502, .Width = 90, .Height = 28,
                .DialogResult = DialogResult.OK
            }
            btnCancel = New Button With {
                .Text = "Cancel", .Left = 546, .Top = 502, .Width = 90, .Height = 28,
                .DialogResult = DialogResult.Cancel
            }

            ' ── Sự kiện ──
            AddHandler btnAdd.Click, Sub(s, e)
                                         Dim idx As Integer = dgv.Rows.Add("", "")
                                         dgv.CurrentCell = dgv.Rows(idx).Cells(0)
                                         dgv.BeginEdit(True)
                                     End Sub

            AddHandler btnRemove.Click, Sub(s, e)
                                            If dgv.CurrentRow IsNot Nothing Then dgv.Rows.Remove(dgv.CurrentRow)
                                            If dgv.Rows.Count = 0 Then dgv.Rows.Add("", "")
                                        End Sub

            AddHandler btnClearHistory.Click, Sub(s, e)
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

            AddHandler btnOK.Click, Sub(s, e)
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
                                            MessageBox.Show(
                                                "Vui lòng nhập ít nhất 1 cặp TÌM/THAY.",
                                                "Chưa có dữ liệu",
                                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                            Me.DialogResult = DialogResult.None
                                            Return
                                        End If

                                        ' Scope
                                        If rbSelected.Checked Then
                                            Scope = ApplyScope.SelectedObjects
                                        ElseIf rbCurrentSheet.Checked Then
                                            Scope = ApplyScope.CurrentSheet
                                        Else
                                            Scope = ApplyScope.AllSheets
                                        End If

                                        SaveHistory = chkSaveHistory.Checked
                                    End Sub

            ' ── Add controls ──
            Me.Controls.Add(lblHint)
            Me.Controls.Add(grpScope)
            Me.Controls.Add(lblGrid)
            Me.Controls.Add(dgv)
            Me.Controls.Add(btnAdd)
            Me.Controls.Add(btnRemove)
            Me.Controls.Add(chkSaveHistory)
            Me.Controls.Add(btnClearHistory)
            Me.Controls.Add(btnOK)
            Me.Controls.Add(btnCancel)

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub
    End Class
End Namespace