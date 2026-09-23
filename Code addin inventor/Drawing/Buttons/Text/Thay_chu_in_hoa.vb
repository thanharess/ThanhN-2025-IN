Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Drawing.Buttons.Drawtext
    Public Module Doichuhoa

        ' Nhớ lựa chọn lần trước (0..4)
        Private _lastOption As Integer = 0

        '═══════════════════════════════════════════════════════════
        ' STATE dùng chung cho InteractionEvents
        '═══════════════════════════════════════════════════════════
        Private _inventorApp As Inventor.Application
        Private _oDrawDoc As DrawingDocument
        Private _opt As Integer = 0
        Private _collected As List(Of Object)
        Private _userDone As Boolean
        Private _ie As InteractionEvents
        Private _se As SelectEvents


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
                Console.WriteLine("Vui lòng mở bản vẽ (IDW/DWG).")
                Return
            End If

            _oDrawDoc = CType(_inventorApp.ActiveDocument, DrawingDocument)

            ' ── 2. Form chọn kiểu ──
            Using form As New CaseConvertForm(_lastOption)
                If form.ShowDialog() <> DialogResult.OK Then
                    Console.WriteLine("✖ Đã hủy lệnh.")
                    Return
                End If
                _lastOption = form.SelectedIndex
            End Using

            _opt = _lastOption
            Console.WriteLine("=== DOICHUHOA - Kiểu " & _opt & " ===")

            ' ── 3. Nếu đã có selection sẵn → dùng luôn ──
            Dim preCount As Integer = 0
            Try : preCount = _oDrawDoc.SelectSet.Count : Catch : End Try

            If preCount > 0 Then
                Dim preTargets As New List(Of Object)()
                For i As Integer = 1 To preCount
                    Dim obj As Object = Nothing
                    Try : obj = _oDrawDoc.SelectSet.Item(i) : Catch : End Try
                    If obj Is Nothing Then Continue For
                    If GetTextFromEntity(obj) IsNot Nothing Then
                        preTargets.Add(obj)
                    End If
                Next

                If preTargets.Count > 0 Then
                    Console.WriteLine($"✔ Dùng {preTargets.Count} đối tượng đã chọn sẵn.")
                    ApplyCaseConvert(preTargets, _opt)
                    Try
                        _inventorApp.ActiveView.Update()
                        _oDrawDoc.SelectSet.Clear()
                        _inventorApp.ActiveView.Update()
                    Catch
                    End Try
                    Return
                End If
            End If

            ' ── 4. Chưa có → bật chế độ quét chuột ──
            _collected = New List(Of Object)()
            _userDone = False

            Try : _oDrawDoc.SelectSet.Clear() : Catch : End Try

            Try
                _ie = _inventorApp.CommandManager.CreateInteractionEvents()
                _ie.InteractionDisabled = False

                _se = _ie.SelectEvents
                _se.AddSelectionFilter(SelectionFilterEnum.kAllEntitiesFilter)
                _se.WindowSelectEnabled = True

                AddHandler _se.OnSelect, AddressOf OnSelectHandler
                AddHandler _ie.OnTerminate, AddressOf OnTerminateHandler

                _ie.Start()
            Catch ex As Exception
                Console.WriteLine("⚠️ Không khởi động được chế độ chọn: " & ex.Message)
                CleanupInteraction()
                Return
            End Try

            Console.WriteLine("─────────────────────────────────────────────────")
            Console.WriteLine("▶ Click hoặc QUÉT CHUỘT để chọn đối tượng.")
            Console.WriteLine("▶ Nhấn ESC (hoặc chuột phải → Done) để kết thúc.")
            Console.WriteLine("─────────────────────────────────────────────────")

            ' ═══════════════════════════════════════════════════════
            '  Giữ OnExecute không return để Inventor không kill IE
            ' ═══════════════════════════════════════════════════════
            Do While Not _userDone
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(30)
            Loop

            ' ── 5. Cleanup ──
            CleanupInteraction()

            ' ── 6. Áp dụng ──
            If _collected Is Nothing OrElse _collected.Count = 0 Then
                Console.WriteLine("⚠️ Không có đối tượng nào được chọn.")
                Return
            End If

            ApplyCaseConvert(_collected, _opt)

            Try
                _inventorApp.ActiveView.Update()
                _oDrawDoc.SelectSet.Clear()
                _inventorApp.ActiveView.Update()
            Catch
            End Try
        End Sub


        '═══════════════════════════════════════════════════════════
        ' CALLBACK: mỗi lần user chọn (click hoặc quét chuột)
        '═══════════════════════════════════════════════════════════
        Private Sub OnSelectHandler(
            ByVal JustSelectedEntities As ObjectsEnumerator,
            ByVal SelectionDevice As SelectionDeviceEnum,
            ByVal ModelPosition As Inventor.Point,
            ByVal ViewPosition As Point2d,
            ByVal View As Inventor.View)

            Try
                For Each obj As Object In JustSelectedEntities
                    If obj Is Nothing Then Continue For
                    If _collected.Contains(obj) Then Continue For

                    _collected.Add(obj)
                    Console.WriteLine($"  + {obj.Type}")
                Next
            Catch ex As Exception
                Console.WriteLine("  ⚠️ OnSelect: " & ex.Message)
            End Try
        End Sub


        '═══════════════════════════════════════════════════════════
        ' CALLBACK: user kết thúc (ESC / Done)
        '═══════════════════════════════════════════════════════════
        Private Sub OnTerminateHandler()
            _userDone = True     ' ← thoát vòng DoEvents
        End Sub


        '═══════════════════════════════════════════════════════════
        ' CLEANUP
        '═══════════════════════════════════════════════════════════
        Private Sub CleanupInteraction()
            Try
                If _se IsNot Nothing Then
                    RemoveHandler _se.OnSelect, AddressOf OnSelectHandler
                End If
            Catch
            End Try
            Try
                If _ie IsNot Nothing Then
                    RemoveHandler _ie.OnTerminate, AddressOf OnTerminateHandler
                End If
            Catch
            End Try
            Try
                If _ie IsNot Nothing Then
                    _ie.Stop()
                End If
            Catch
            End Try

            _se = Nothing
            _ie = Nothing
        End Sub


        '═══════════════════════════════════════════════════════════
        ' ÁP DỤNG CHUYỂN ĐỔI HOA/THƯỜNG
        '═══════════════════════════════════════════════════════════
        Private Sub ApplyCaseConvert(ByVal targets As List(Of Object), ByVal opt As Integer)
            If targets Is Nothing OrElse targets.Count = 0 Then Return

            Dim count As Integer = 0
            For Each tObj In targets
                Dim oldTxt As String = GetTextFromEntity(tObj)
                If String.IsNullOrEmpty(oldTxt) Then Continue For

                Dim newTxt As String
                Select Case opt
                    Case 1 : newTxt = CaseConvert(oldTxt, True)         ' Thường
                    Case 2 : newTxt = SentenceCase(oldTxt)              ' Hoa đầu dòng
                    Case 3 : newTxt = TitleCase(oldTxt)                 ' Hoa đầu từ
                    Case 4 : newTxt = FirstUpperRestLower(oldTxt)       ' Hoa đầu tiên
                    Case Else : newTxt = CaseConvert(oldTxt, False)     ' Hoa
                End Select

                If newTxt <> oldTxt Then
                    If SetTextToEntity(tObj, newTxt) Then count += 1
                End If
            Next

            Console.WriteLine($"✅ Hoàn tất – đã xử lý {count}/{targets.Count} đối tượng.")
        End Sub


        '═══════════════════════════════════════════════════════════
        ' ĐỌC TEXT
        '═══════════════════════════════════════════════════════════
        Private Function GetTextFromEntity(ByVal obj As Object) As String
            Try
                If TypeOf obj Is GeneralNote Then Return CType(obj, GeneralNote).Text
                If TypeOf obj Is LeaderNote Then Return CType(obj, LeaderNote).Text
                If TypeOf obj Is Inventor.TextBox Then Return CType(obj, Inventor.TextBox).Text

                If TypeOf obj Is DrawingDimension Then
                    Return CType(obj, DrawingDimension).Text.Text
                End If

                If TypeOf obj Is DrawingViewLabel Then
                    Return CType(obj, DrawingViewLabel).FormattedText
                End If

                If TypeOf obj Is ModelGeneralNote Then
                    Return CType(obj, ModelGeneralNote).Definition.Text.Text
                End If
                If TypeOf obj Is ModelLeaderNote Then
                    Return CType(obj, ModelLeaderNote).Definition.Text.Text
                End If
            Catch
            End Try
            Return Nothing
        End Function


        '═══════════════════════════════════════════════════════════
        ' GHI TEXT
        '═══════════════════════════════════════════════════════════
        Private Function SetTextToEntity(ByVal obj As Object, ByVal newText As String) As Boolean
            Try
                If TypeOf obj Is GeneralNote Then
                    CType(obj, GeneralNote).Text = newText
                    Return True
                End If
                If TypeOf obj Is LeaderNote Then
                    CType(obj, LeaderNote).Text = newText
                    Return True
                End If
                If TypeOf obj Is Inventor.TextBox Then
                    CType(obj, Inventor.TextBox).Text = newText
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
        ' 1 & 2) UPPER / lower
        '═══════════════════════════════════════════════════════════
        Private Function CaseConvert(ByVal s As String, ByVal toLower As Boolean) As String
            Dim sb As New StringBuilder(s.Length)
            Dim i As Integer = 0
            While i < s.Length
                Dim ch As Char = s(i)
                If ch = "\"c OrElse ch = "<"c Then
                    Dim consumed As Integer
                    AppendEscapeOrTag(s, i, sb, consumed)
                    i += consumed
                    Continue While
                End If
                sb.Append(If(toLower, Char.ToLowerInvariant(ch), Char.ToUpperInvariant(ch)))
                i += 1
            End While
            Return sb.ToString()
        End Function


        '═══════════════════════════════════════════════════════════
        ' 3) Hoa đầu dòng
        '═══════════════════════════════════════════════════════════
        Private Function SentenceCase(ByVal s As String) As String
            Dim sb As New StringBuilder(s.Length)
            Dim startOfLine As Boolean = True
            Dim i As Integer = 0

            While i < s.Length
                Dim ch As Char = s(i)

                If ch = "\"c OrElse ch = "<"c Then
                    Dim consumed As Integer
                    Dim tag As String = PeekEscapeOrTag(s, i, consumed)
                    sb.Append(tag)
                    i += consumed
                    If IsLineBreakTag(tag) Then startOfLine = True
                    Continue While
                End If

                If Char.IsLetter(ch) Then
                    sb.Append(If(startOfLine, Char.ToUpperInvariant(ch), Char.ToLowerInvariant(ch)))
                    startOfLine = False
                Else
                    If ch = ControlChars.Lf OrElse ch = ControlChars.Cr Then startOfLine = True
                    sb.Append(ch)
                End If
                i += 1
            End While
            Return sb.ToString()
        End Function


        '═══════════════════════════════════════════════════════════
        ' 4) Hoa đầu từ
        '═══════════════════════════════════════════════════════════
        Private Function TitleCase(ByVal s As String) As String
            Dim sb As New StringBuilder(s.Length)
            Dim startOfWord As Boolean = True
            Dim i As Integer = 0

            While i < s.Length
                Dim ch As Char = s(i)

                If ch = "\"c OrElse ch = "<"c Then
                    Dim consumed As Integer
                    Dim tag As String = PeekEscapeOrTag(s, i, consumed)
                    sb.Append(tag)
                    i += consumed
                    If IsLineBreakTag(tag) Then startOfWord = True
                    Continue While
                End If

                If Char.IsLetter(ch) Then
                    sb.Append(If(startOfWord, Char.ToUpperInvariant(ch), Char.ToLowerInvariant(ch)))
                    startOfWord = False
                Else
                    startOfWord = True
                    sb.Append(ch)
                End If
                i += 1
            End While
            Return sb.ToString()
        End Function


        '═══════════════════════════════════════════════════════════
        ' 5) Hoa đầu tiên
        '═══════════════════════════════════════════════════════════
        Private Function FirstUpperRestLower(ByVal s As String) As String
            If String.IsNullOrEmpty(s) Then Return s

            Dim lower As String = CaseConvert(s, toLower:=True)
            Dim sb As New StringBuilder(lower)
            Dim i As Integer = 0

            While i < sb.Length
                Dim ch As Char = sb(i)

                If ch = "\"c OrElse ch = "<"c Then
                    Dim consumed As Integer
                    Dim tmp As New StringBuilder()
                    AppendEscapeOrTag(sb.ToString(), i, tmp, consumed)
                    i += consumed
                    Continue While
                End If

                If Char.IsLetter(ch) Then
                    sb(i) = Char.ToUpperInvariant(ch)
                    Exit While
                End If
                i += 1
            End While
            Return sb.ToString()
        End Function


        '═══════════════════════════════════════════════════════════
        ' Đọc tag/escape (không ghi vào sb)
        '═══════════════════════════════════════════════════════════
        Private Function PeekEscapeOrTag(ByVal s As String, ByVal i As Integer, ByRef consumed As Integer) As String
            Dim tmp As New StringBuilder()
            AppendEscapeOrTag(s, i, tmp, consumed)
            Return tmp.ToString()
        End Function


        '═══════════════════════════════════════════════════════════
        ' Xử lý tag XML Inventor <...> và escape AutoCAD \...
        '═══════════════════════════════════════════════════════════
        Private Sub AppendEscapeOrTag(ByVal s As String, ByVal i As Integer, ByVal sb As StringBuilder, ByRef consumed As Integer)
            Dim ch As Char = s(i)

            ' ── Inventor XML tag: <...> ──
            If ch = "<"c Then
                Dim j As Integer = i + 1
                While j < s.Length AndAlso s(j) <> ">"c
                    j += 1
                End While
                If j < s.Length Then j += 1
                Dim len As Integer = j - i
                sb.Append(s, i, len)
                consumed = len
                Return
            End If

            ' ── AutoCAD escape: \... ──
            If ch = "\"c AndAlso i + 1 < s.Length Then
                Dim nextCh As Char = s(i + 1)

                If nextCh = "P"c OrElse nextCh = "p"c Then
                    sb.Append(s, i, 2)
                    consumed = 2
                    Return
                End If
                If nextCh = "\"c OrElse nextCh = "{"c OrElse nextCh = "}"c OrElse nextCh = "~"c Then
                    sb.Append(s, i, 2)
                    consumed = 2
                    Return
                End If

                Dim j As Integer = i + 1
                While j < s.Length AndAlso s(j) <> ";"c
                    j += 1
                End While
                If j < s.Length Then j += 1
                sb.Append(s, i, j - i)
                consumed = j - i
                Return
            End If

            sb.Append(ch)
            consumed = 1
        End Sub


        '═══════════════════════════════════════════════════════════
        ' Xác định tag xuống dòng
        '═══════════════════════════════════════════════════════════
        Private Function IsLineBreakTag(ByVal tag As String) As Boolean
            If String.IsNullOrEmpty(tag) Then Return False
            If tag = "\P" OrElse tag = "\p" Then Return True
            Dim t As String = tag.ToLower()
            Return t.Contains("<br") OrElse t.Contains("</paragraph")
        End Function

    End Module


    '═══════════════════════════════════════════════════════════
    ' FORM chọn kiểu chuyển đổi (giữ nguyên)
    '═══════════════════════════════════════════════════════════
    Public Class CaseConvertForm : Inherits Form

        Private rbHoa As RadioButton
        Private rbThuong As RadioButton
        Private rbDauDong As RadioButton
        Private rbMoiTu As RadioButton
        Private rbChuDau As RadioButton

        Private btnOK As Button
        Private btnCancel As Button

        Public Property SelectedIndex As Integer = 0

        Public Sub New(ByVal preSelected As Integer)
            Me.Text = "Chọn kiểu chuyển đổi chữ"
            Me.Width = 400
            Me.Height = 340
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False

            Dim lblTitle As New Label With {
                .Text = "Chọn kiểu chuyển đổi:",
                .Left = 20, .Top = 15, .Width = 350,
                .Font = New System.Drawing.Font(Me.Font, System.Drawing.FontStyle.Bold)
            }

            Dim pnlOpts As New Panel With {
                .Left = 20, .Top = 42, .Width = 350, .Height = 150
            }

            rbHoa = New RadioButton With {
                .Text = "1. Hoa           (IN HOA HẾT)",
                .Left = 5, .Top = 3, .Width = 340
            }
            rbThuong = New RadioButton With {
                .Text = "2. Thường        (in thường hết)",
                .Left = 5, .Top = 27, .Width = 340
            }
            rbDauDong = New RadioButton With {
                .Text = "3. Hoa đầu dòng  (chữ cái đầu MỖI DÒNG)",
                .Left = 5, .Top = 51, .Width = 340
            }
            rbMoiTu = New RadioButton With {
                .Text = "4. Hoa đầu từ    (chữ cái đầu MỖI TỪ)",
                .Left = 5, .Top = 75, .Width = 340
            }
            rbChuDau = New RadioButton With {
                .Text = "5. Hoa đầu tiên  (chỉ chữ cái đầu, còn lại thường)",
                .Left = 5, .Top = 99, .Width = 340
            }

            pnlOpts.Controls.AddRange(New Control() {rbHoa, rbThuong, rbDauDong, rbMoiTu, rbChuDau})

            Select Case preSelected
                Case 1 : rbThuong.Checked = True
                Case 2 : rbDauDong.Checked = True
                Case 3 : rbMoiTu.Checked = True
                Case 4 : rbChuDau.Checked = True
                Case Else : rbHoa.Checked = True
            End Select

            btnOK = New Button With {
                .Text = "OK",
                .Left = 195, .Top = 205, .Width = 80, .Height = 28,
                .DialogResult = DialogResult.OK
            }
            btnCancel = New Button With {
                .Text = "Cancel",
                .Left = 290, .Top = 205, .Width = 80, .Height = 28,
                .DialogResult = DialogResult.Cancel
            }

            AddHandler btnOK.Click, Sub(s, e)
                                        If rbHoa.Checked Then SelectedIndex = 0
                                        If rbThuong.Checked Then SelectedIndex = 1
                                        If rbDauDong.Checked Then SelectedIndex = 2
                                        If rbMoiTu.Checked Then SelectedIndex = 3
                                        If rbChuDau.Checked Then SelectedIndex = 4
                                    End Sub

            Me.Controls.AddRange(New Control() {lblTitle, pnlOpts, btnOK, btnCancel})
            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub
    End Class
End Namespace