Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Inventor

Namespace ToolInventor2025.Drawing.Buttons.Drawtext
    Public Module TextReplaceModule

        '═══════════════════════════════════════════════════════════
        ' STATE dùng chung cho InteractionEvents
        '═══════════════════════════════════════════════════════════
        Private _inventorApp As Inventor.Application
        Private _oDrawDoc As DrawingDocument
        Private _srcObj As Object
        Private _srcText As String
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
            Console.WriteLine("=== TEXTREPLACE ===")

            ' ── 2. Chọn mẫu ──
            _srcObj = Nothing
            Try
                _srcObj = _inventorApp.CommandManager.Pick(
                    SelectionFilterEnum.kAllEntitiesFilter,
                    "Chọn 1 đối tượng mẫu: ")
            Catch
            End Try
            If _srcObj Is Nothing Then
                Console.WriteLine("⚠️ Không chọn được mẫu.")
                Return
            End If

            _srcText = GetTextFromEntity(_srcObj)
            If String.IsNullOrEmpty(_srcText) Then
                Console.WriteLine("⚠️ Mẫu không có text.")
                Return
            End If
            Console.WriteLine($"→ Nội dung mẫu: ""{_srcText}""")

            ' ── 3. Chuẩn bị danh sách ──
            _collected = New List(Of Object)()
            _userDone = False

            ' Xóa selection cũ
            Try : _oDrawDoc.SelectSet.Clear() : Catch : End Try

            ' ── 4. Khởi động InteractionEvents ──
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
            '  QUAN TRỌNG: Giữ OnExecute không return để Inventor
            '  không kill InteractionEvents. Thoát khi user ESC.
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

            Dim count As Integer = 0
            For Each tObj In _collected
                If SetTextToEntity(tObj, _srcText) Then count += 1
            Next
            Console.WriteLine($"✅ Đã thay {count}/{_collected.Count}")

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
                    If _srcObj IsNot Nothing AndAlso Object.ReferenceEquals(obj, _srcObj) Then Continue For
                    If _collected.Contains(obj) Then Continue For

                    _collected.Add(obj)
                    Console.WriteLine($"  + {obj.Type}")
                Next
            Catch ex As Exception
                Console.WriteLine("  ⚠️ OnSelect: " & ex.Message)
            End Try
        End Sub


        '═══════════════════════════════════════════════════════════
        ' CALLBACK: user kết thúc (ESC / Done / lệnh khác)
        '═══════════════════════════════════════════════════════════
        Private Sub OnTerminateHandler()
            _userDone = True     ' ← thoát vòng DoEvents trong OnExecute
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
        ' ĐỌC / GHI TEXT — giữ nguyên như file gốc
        '═══════════════════════════════════════════════════════════
        Private Function GetTextFromEntity(ByVal obj As Object) As String
            Try
                If TypeOf obj Is GeneralNote Then Return CType(obj, GeneralNote).Text
                If TypeOf obj Is LeaderNote Then Return CType(obj, LeaderNote).Text
                If TypeOf obj Is DrawingDimension Then Return CType(obj, DrawingDimension).Text.Text
                If TypeOf obj Is Inventor.TextBox Then Return CType(obj, Inventor.TextBox).Text

                If TypeOf obj Is DrawingViewLabel Then
                    Dim lbl As DrawingViewLabel = CType(obj, DrawingViewLabel)
                    Dim fmt As String = ""
                    Try : fmt = lbl.FormattedText : Catch : End Try
                    If Not String.IsNullOrEmpty(fmt) Then Return fmt
                    Return lbl.Text
                End If

                If TypeOf obj Is ModelGeneralNote Then
                    Return CType(obj, ModelGeneralNote).Definition.Text.Text
                End If
                If TypeOf obj Is ModelLeaderNote Then
                    Return CType(obj, ModelLeaderNote).Definition.Text.Text
                End If

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

                If TypeOf obj Is SketchedSymbol Then
                    Dim oSymbol As SketchedSymbol = CType(obj, SketchedSymbol)
                    Dim oSketch As DrawingSketch = oSymbol.Definition.Sketch
                    For Each oTB As Inventor.TextBox In oSketch.TextBoxes
                        Dim rt As String = oSymbol.GetResultText(oTB)
                        If Not String.IsNullOrEmpty(rt) Then Return rt
                    Next
                End If
            Catch
            End Try
            Return Nothing
        End Function


        Private Function SetTextToEntity(ByVal obj As Object, ByVal newText As String) As Boolean
            If obj Is Nothing OrElse newText Is Nothing Then Return False
            Try
                If TypeOf obj Is GeneralNote Then
                    Try : CType(obj, GeneralNote).FormattedText = newText
                    Catch : CType(obj, GeneralNote).Text = newText : End Try
                    Return True
                End If

                If TypeOf obj Is LeaderNote Then
                    Try : CType(obj, LeaderNote).FormattedText = newText
                    Catch : CType(obj, LeaderNote).Text = newText : End Try
                    Return True
                End If

                If TypeOf obj Is DrawingDimension Then
                    Try
                        CType(obj, DrawingDimension).Text.FormattedText = newText
                        Return True
                    Catch ex As Exception
                        Console.WriteLine("  ⚠️ Dimension: " & ex.Message)
                        Return False
                    End Try
                End If

                If TypeOf obj Is Inventor.TextBox Then
                    Try : CType(obj, Inventor.TextBox).FormattedText = newText
                    Catch : CType(obj, Inventor.TextBox).Text = newText : End Try
                    Return True
                End If

                If TypeOf obj Is DrawingViewLabel Then
                    Return RecreateViewLabel(CType(obj, DrawingViewLabel), newText)
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
                        If ws.Definitions Is Nothing OrElse ws.Definitions.Count < 1 Then Return False

                        Dim ok As Boolean = False
                        For i As Integer = 1 To ws.Definitions.Count
                            Try
                                ws.Definitions.Item(i).TailNote = newText
                                ok = True
                            Catch
                            End Try
                        Next
                        Return ok
                    Catch
                    End Try
                    Return False
                End If

                If TypeOf obj Is SketchedSymbol Then
                    Dim oSymbol As SketchedSymbol = CType(obj, SketchedSymbol)
                    Dim oSketch As DrawingSketch = oSymbol.Definition.Sketch
                    Dim updated As Boolean = False
                    For Each oTB As Inventor.TextBox In oSketch.TextBoxes
                        Try
                            oSymbol.SetPromptResultText(oTB, newText)
                            updated = True
                        Catch
                            Try : oTB.FormattedText = newText : updated = True
                            Catch
                                Try : oTB.Text = newText : updated = True
                                Catch : End Try
                            End Try
                        End Try
                    Next
                    Return updated
                End If
            Catch ex As Exception
                Console.WriteLine("  ⚠️ " & ex.Message)
            End Try
            Return False
        End Function


        Private Function RecreateViewLabel(ByVal lbl As DrawingViewLabel, ByVal newText As String) As Boolean
            Try
                Dim view As DrawingView = Nothing
                Try
                    view = CType(lbl.Parent, DrawingView)
                Catch
                End Try

                If view Is Nothing Then
                    Try
                        lbl.FormattedText = newText
                        Return True
                    Catch ex As Exception
                        Console.WriteLine("  ⚠️ Fallback: " & ex.Message)
                        Return False
                    End Try
                End If

                Dim oldPos As Point2d = Nothing
                Try : oldPos = view.LabelPosition : Catch : End Try

                Try : lbl.Delete() : Catch : End Try

                Try
                    view.ShowLabel = True
                Catch ex As Exception
                    Console.WriteLine("  ⚠️ Không bật lại ShowLabel: " & ex.Message)
                    Return False
                End Try

                Dim newLbl As DrawingViewLabel = view.Label
                If newLbl Is Nothing Then Return False

                Try
                    newLbl.FormattedText = newText
                Catch
                    Try : newLbl.FormattedText = newText : Catch : End Try
                End Try

                If oldPos IsNot Nothing Then
                    Try : view.LabelPosition = oldPos : Catch : End Try
                End If

                Return True
            Catch ex As Exception
                Console.WriteLine("  ⚠️ RecreateViewLabel: " & ex.Message)
                Return False
            End Try
        End Function

    End Module
End Namespace