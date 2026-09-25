Option Explicit On
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports Inventor
Imports Drw = System.Drawing

Namespace ToolInventor2025.Drawing.Buttons.Drawtext

    Public Module Doichuhoa

        Private _lastOption As Integer = 0
        Private _lastScope As Integer = 0

        Private _inventorApp As Inventor.Application
        Private _oDrawDoc As DrawingDocument
        Private _opt As Integer = 0
        Private _collected As List(Of Object)
        Private _userDone As Boolean
        Private _ie As InteractionEvents
        Private _se As SelectEvents


        Public Sub OnExecute(ByVal Context As NameValueMap)

            ' ── 1. Lấy Inventor ──
            Try
                _inventorApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch ex As Exception
                MessageBox.Show("Không tìm thấy Inventor đang chạy." & vbCrLf & ex.Message,
                                "Đổi chữ hoa/thường", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            If _inventorApp.ActiveDocument Is Nothing OrElse
               _inventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                MessageBox.Show("Vui lòng mở bản vẽ (IDW/DWG).",
                                "Đổi chữ hoa/thường", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            _oDrawDoc = CType(_inventorApp.ActiveDocument, DrawingDocument)

            ' ── 2. Form chọn kiểu + phạm vi ──
            Dim caseMode As Integer = 0
            Dim scope As Integer = 0

            Using form As New CaseConvertForm(_lastOption, _lastScope)
                If form.ShowDialog() <> DialogResult.OK Then Return
                caseMode = form.SelectedIndex
                scope = form.SelectedScope
                _lastOption = caseMode
                _lastScope = scope
            End Using

            _opt = caseMode

            '═══════════════════════════════════════════════════════
            ' 3A. ÁP DỤNG HÀNG LOẠT THEO PHẠM VI
            '═══════════════════════════════════════════════════════
            If scope = 1 Then
                ' Chỉ Sheet hiện tại
                Dim targets As New List(Of Object)()
                CollectAllTexts(_oDrawDoc.ActiveSheet, targets)

                If targets.Count = 0 Then
                    MessageBox.Show("Không có Text nào trên Sheet hiện tại.",
                                    "Đổi chữ hoa/thường", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ApplyCaseConvert(targets, _opt)

                Try : _inventorApp.ActiveView.Update() : Catch : End Try

                MessageBox.Show("Đã xử lý " & targets.Count.ToString() & " đối tượng trên Sheet hiện tại.",
                                "Đổi chữ hoa/thường", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            If scope = 2 Then
                ' Toàn bộ Drawing
                Dim targets As New List(Of Object)()
                For Each sheet As Sheet In _oDrawDoc.Sheets
                    CollectAllTexts(sheet, targets)
                Next

                If targets.Count = 0 Then
                    MessageBox.Show("Không có Text nào trong Drawing.",
                                    "Đổi chữ hoa/thường", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ApplyCaseConvert(targets, _opt)

                Try : _inventorApp.ActiveView.Update() : Catch : End Try

                MessageBox.Show("Đã xử lý " & targets.Count.ToString() & " đối tượng trên toàn bộ Drawing.",
                                "Đổi chữ hoa/thường", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            '═══════════════════════════════════════════════════════
            ' 3B. CHẾ ĐỘ CHỌN BẰNG CHUỘT (scope = 0)
            '═══════════════════════════════════════════════════════

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
                MessageBox.Show("Không khởi động được chế độ chọn: " & ex.Message,
                                "Đổi chữ hoa/thường", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                CleanupInteraction()
                Return
            End Try

            _ie.StatusBarText = "Click hoặc QUÉT CHUỘT để chọn. ESC để kết thúc."

            Do While Not _userDone
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(30)
            Loop

            CleanupInteraction()

            If _collected Is Nothing OrElse _collected.Count = 0 Then
                MessageBox.Show("Không có đối tượng nào được chọn.",
                                "Đổi chữ hoa/thường", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
        ' THU THẬP TẤT CẢ TEXT TRÊN SHEET
        '═══════════════════════════════════════════════════════════
        Private Sub CollectAllTexts(ByVal sheet As Sheet, ByRef targets As List(Of Object))
            If sheet Is Nothing Then Return

            '--- DrawingNotes: GeneralNote, LeaderNote, HoleThreadNote... ---
            Try
                For Each note As DrawingNote In sheet.DrawingNotes
                    Try
                        If GetTextFromEntity(note) IsNot Nothing Then
                            targets.Add(note)
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            '--- DrawingDimensions ---
            Try
                For Each dimasd As DrawingDimension In sheet.DrawingDimensions
                    Try
                        If GetTextFromEntity(dimasd) IsNot Nothing Then
                            targets.Add(dimasd)
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            '--- DrawingView Labels ---
            Try
                For Each view As DrawingView In sheet.DrawingViews
                    Try
                        Dim lbl As DrawingViewLabel = view.Label
                        If lbl IsNot Nothing Then targets.Add(lbl)
                    Catch
                    End Try
                Next
            Catch
            End Try

            '--- SketchedSymbols ---
            Try
                For Each sym As SketchedSymbol In sheet.SketchedSymbols
                    Try
                        If GetTextFromEntity(sym) IsNot Nothing Then
                            targets.Add(sym)
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try
            '--- Welding Symbols ---
            Try
                Dim weldCol As Object = sheet.WeldingSymbols
                If weldCol IsNot Nothing Then
                    Dim cnt As Integer = 0
                    Try : cnt = CInt(weldCol.Count) : Catch : End Try
                    For i As Integer = 1 To cnt
                        Try
                            Dim ws As Object = weldCol.Item(i)
                            If ws IsNot Nothing Then targets.Add(ws)
                        Catch
                        End Try
                    Next
                End If
            Catch
            End Try
        End Sub


        '═══════════════════════════════════════════════════════════
        ' CALLBACK: chọn (click / quét)
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
                Next
            Catch
            End Try
        End Sub


        Private Sub OnTerminateHandler()
            _userDone = True
        End Sub


        Private Sub CleanupInteraction()
            Try
                If _se IsNot Nothing Then RemoveHandler _se.OnSelect, AddressOf OnSelectHandler
            Catch
            End Try
            Try
                If _ie IsNot Nothing Then RemoveHandler _ie.OnTerminate, AddressOf OnTerminateHandler
            Catch
            End Try
            Try
                If _ie IsNot Nothing Then _ie.Stop()
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
                ' ── Welding Symbol ──
                If TypeOf tObj Is DrawingWeldingSymbol Then
                    Try
                        Dim ws As DrawingWeldingSymbol = CType(tObj, DrawingWeldingSymbol)
                        If ws.Retrieved Then Continue For
                        If ws.Definitions Is Nothing Then Continue For

                        For i As Integer = 1 To ws.Definitions.Count
                            Try
                                Dim def As DrawingWeldingSymbolDefinition = ws.Definitions.Item(i)
                                Dim oldTxt As String = def.TailNote
                                If String.IsNullOrEmpty(oldTxt) Then Continue For

                                Dim newTxt As String
                                Select Case opt
                                    Case 1 : newTxt = CaseConvert(oldTxt, True)
                                    Case 2 : newTxt = SentenceCase(oldTxt)
                                    Case 3 : newTxt = TitleCase(oldTxt)
                                    Case 4 : newTxt = FirstUpperRestLower(oldTxt)
                                    Case Else : newTxt = CaseConvert(oldTxt, False)
                                End Select

                                If newTxt <> oldTxt Then
                                    def.TailNote = newTxt
                                    count += 1
                                End If
                            Catch
                            End Try
                        Next
                    Catch
                    End Try
                    Continue For
                End If

                ' ── Đối tượng thường ──
                Dim oldTxt2 As String = GetTextFromEntity(tObj)
                If String.IsNullOrEmpty(oldTxt2) Then Continue For

                Dim newTxt2 As String
                Select Case opt
                    Case 1 : newTxt2 = CaseConvert(oldTxt2, True)
                    Case 2 : newTxt2 = SentenceCase(oldTxt2)
                    Case 3 : newTxt2 = TitleCase(oldTxt2)
                    Case 4 : newTxt2 = FirstUpperRestLower(oldTxt2)
                    Case Else : newTxt2 = CaseConvert(oldTxt2, False)
                End Select

                If newTxt2 <> oldTxt2 Then
                    If SetTextToEntity(tObj, newTxt2) Then count += 1
                End If
            Next
        End Sub


        '═══════════════════════════════════════════════════════════
        ' ĐỌC / GHI TEXT
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

                If TypeOf obj Is SketchedSymbol Then
                    Return CType(obj, SketchedSymbol).Definition.Text
                End If
                ' Welding Symbol
                If TypeOf obj Is DrawingWeldingSymbol Then
                    Try
                        Dim ws As DrawingWeldingSymbol = CType(obj, DrawingWeldingSymbol)
                        If ws.Retrieved Then Return Nothing
                        If ws.Definitions IsNot Nothing AndAlso ws.Definitions.Count > 0 Then
                            Return ws.Definitions.Item(1).TailNote
                        End If
                    Catch
                    End Try
                    Return Nothing
                End If
            Catch
            End Try
            Return Nothing
        End Function


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
                If TypeOf obj Is DrawingWeldingSymbol Then
                    Try
                        Dim ws As DrawingWeldingSymbol = CType(obj, DrawingWeldingSymbol)
                        If ws.Retrieved Then Return False
                        If ws.Definitions IsNot Nothing AndAlso ws.Definitions.Count > 0 Then
                            ws.Definitions.Item(1).TailNote = newText
                            Return True
                        End If
                    Catch
                    End Try
                    Return False
                End If
                If TypeOf obj Is SketchedSymbol Then
                    Try
                        CType(obj, SketchedSymbol).Definition.Text = newText
                        Return True
                    Catch
                    End Try
                End If
            Catch
            End Try
            Return False
        End Function


        '═══════════════════════════════════════════════════════════
        ' CASE CONVERTERS
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


        Private Function PeekEscapeOrTag(ByVal s As String, ByVal i As Integer, ByRef consumed As Integer) As String
            Dim tmp As New StringBuilder()
            AppendEscapeOrTag(s, i, tmp, consumed)
            Return tmp.ToString()
        End Function


        Private Sub AppendEscapeOrTag(ByVal s As String, ByVal i As Integer, ByVal sb As StringBuilder, ByRef consumed As Integer)
            Dim ch As Char = s(i)

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


        Private Function IsLineBreakTag(ByVal tag As String) As Boolean
            If String.IsNullOrEmpty(tag) Then Return False
            If tag = "\P" OrElse tag = "\p" Then Return True
            Dim t As String = tag.ToLower()
            Return t.Contains("<br") OrElse t.Contains("</paragraph")
        End Function

    End Module


    '═══════════════════════════════════════════════════════════
    ' FORM chọn kiểu + phạm vi
    '═══════════════════════════════════════════════════════════
    Public Class CaseConvertForm
        Inherits Form

        Private rbHoa As RadioButton
        Private rbThuong As RadioButton
        Private rbDauDong As RadioButton
        Private rbMoiTu As RadioButton
        Private rbChuDau As RadioButton

        Private rbPick As RadioButton
        Private rbCurrentSheet As RadioButton
        Private rbAllSheets As RadioButton

        Private btnOK As Button
        Private btnCancel As Button

        Public Property SelectedIndex As Integer = 0
        Public Property SelectedScope As Integer = 0

        Public Sub New(ByVal preSelected As Integer, ByVal preScope As Integer)
            Me.Text = "Đổi chữ hoa / thường"
            Me.AutoScaleMode = AutoScaleMode.None
            Me.ClientSize = New Drw.Size(560, 520)
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
            pnlHeader.Size = New Drw.Size(560, 70)
            pnlHeader.BackColor = Drw.Color.FromArgb(45, 100, 180)
            Me.Controls.Add(pnlHeader)

            Dim lblTitle As New Label()
            lblTitle.Text = "ĐỔI CHỮ HOA / THƯỜNG"
            lblTitle.Font = New Drw.Font("Segoe UI", 14.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            lblTitle.ForeColor = Drw.Color.White
            lblTitle.Dock = DockStyle.Fill
            lblTitle.TextAlign = Drw.ContentAlignment.MiddleCenter
            pnlHeader.Controls.Add(lblTitle)

            Dim lblSub As New Label()
            lblSub.Text = "Chọn kiểu chuyển đổi và phạm vi áp dụng"
            lblSub.Font = New Drw.Font("Segoe UI", 9.0F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)
            lblSub.ForeColor = Drw.Color.FromArgb(220, 230, 245)
            lblSub.Dock = DockStyle.Bottom
            lblSub.Height = 20
            lblSub.TextAlign = Drw.ContentAlignment.MiddleCenter
            pnlHeader.Controls.Add(lblSub)

            '===== GROUP 1: KIỂU CHUYỂN ĐỔI =====
            Dim gb1 As New GroupBox() With {
                .Text = "1. Kiểu chuyển đổi",
                .Location = New Drw.Point(15, 85),
                .Size = New Drw.Size(530, 200),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gb1)

            rbHoa = New RadioButton() With {
                .Text = "Hoa           (IN HOA HẾT)",
                .Location = New Drw.Point(20, 30),
                .Size = New Drw.Size(500, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            rbThuong = New RadioButton() With {
                .Text = "Thường        (in thường hết)",
                .Location = New Drw.Point(20, 62),
                .Size = New Drw.Size(500, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            rbDauDong = New RadioButton() With {
                .Text = "Hoa đầu dòng  (chữ cái đầu MỖI DÒNG)",
                .Location = New Drw.Point(20, 94),
                .Size = New Drw.Size(500, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            rbMoiTu = New RadioButton() With {
                .Text = "Hoa đầu từ    (chữ cái đầu MỖI TỪ)",
                .Location = New Drw.Point(20, 126),
                .Size = New Drw.Size(500, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            rbChuDau = New RadioButton() With {
                .Text = "Hoa đầu tiên  (chỉ chữ cái đầu, còn lại thường)",
                .Location = New Drw.Point(20, 158),
                .Size = New Drw.Size(500, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}

            gb1.Controls.AddRange(New Control() {rbHoa, rbThuong, rbDauDong, rbMoiTu, rbChuDau})

            '===== GROUP 2: PHẠM VI =====
            Dim gb2 As New GroupBox() With {
                .Text = "2. Phạm vi áp dụng",
                .Location = New Drw.Point(15, 295),
                .Size = New Drw.Size(530, 135),
                .Font = New Drw.Font("Segoe UI", 10.0F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point),
                .ForeColor = Drw.Color.FromArgb(45, 100, 180),
                .BackColor = Drw.Color.White}
            Me.Controls.Add(gb2)

            rbPick = New RadioButton() With {
                .Text = "Chọn đối tượng bằng chuột (click / quét)",
                .Location = New Drw.Point(20, 30),
                .Size = New Drw.Size(500, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            rbCurrentSheet = New RadioButton() With {
                .Text = "Tất cả Text trên Sheet hiện tại",
                .Location = New Drw.Point(20, 62),
                .Size = New Drw.Size(500, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}
            rbAllSheets = New RadioButton() With {
                .Text = "Tất cả Text trên TOÀN BỘ Drawing",
                .Location = New Drw.Point(20, 94),
                .Size = New Drw.Size(500, 25),
                .Font = New Drw.Font("Segoe UI", 9.5F, Drw.FontStyle.Regular, Drw.GraphicsUnit.Point)}

            gb2.Controls.AddRange(New Control() {rbPick, rbCurrentSheet, rbAllSheets})

            '--- Chọn theo pre-selected ---
            Select Case preSelected
                Case 1 : rbThuong.Checked = True
                Case 2 : rbDauDong.Checked = True
                Case 3 : rbMoiTu.Checked = True
                Case 4 : rbChuDau.Checked = True
                Case Else : rbHoa.Checked = True
            End Select

            Select Case preScope
                Case 1 : rbCurrentSheet.Checked = True
                Case 2 : rbAllSheets.Checked = True
                Case Else : rbPick.Checked = True
            End Select

            '===== NÚT THỰC HIỆN =====
            Dim btnOK As New Button()
            btnOK.Text = "THỰC HIỆN"
            btnOK.Size = New Drw.Size(150, 44)
            btnOK.Location = New Drw.Point(390, 455)
            btnOK.FlatStyle = FlatStyle.Flat
            btnOK.FlatAppearance.BorderSize = 0
            btnOK.FlatAppearance.MouseOverBackColor = Drw.Color.FromArgb(60, 115, 195)
            btnOK.FlatAppearance.MouseDownBackColor = Drw.Color.FromArgb(30, 80, 155)
            btnOK.BackColor = Drw.Color.FromArgb(45, 100, 180)
            btnOK.ForeColor = Drw.Color.White
            btnOK.Font = New Drw.Font("Segoe UI", 10.5F, Drw.FontStyle.Bold, Drw.GraphicsUnit.Point)
            btnOK.Cursor = Cursors.Hand
            btnOK.UseVisualStyleBackColor = False
            btnOK.DialogResult = DialogResult.OK
            Me.Controls.Add(btnOK)

            Dim btnCancel As New Button()
            btnCancel.Text = "HỦY"
            btnCancel.Size = New Drw.Size(120, 44)
            btnCancel.Location = New Drw.Point(260, 455)
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
            btnCancel.DialogResult = DialogResult.Cancel
            Me.Controls.Add(btnCancel)

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel

            '=== Gán giá trị khi OK ===
            AddHandler btnOK.Click, Sub()
                                        If rbHoa.Checked Then SelectedIndex = 0
                                        If rbThuong.Checked Then SelectedIndex = 1
                                        If rbDauDong.Checked Then SelectedIndex = 2
                                        If rbMoiTu.Checked Then SelectedIndex = 3
                                        If rbChuDau.Checked Then SelectedIndex = 4

                                        If rbPick.Checked Then SelectedScope = 0
                                        If rbCurrentSheet.Checked Then SelectedScope = 1
                                        If rbAllSheets.Checked Then SelectedScope = 2
                                    End Sub
        End Sub
    End Class

End Namespace