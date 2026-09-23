Option Explicit On
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Drawing.Buttons
    Public Module Draw_9

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                Dim invApp As Inventor.Application = g_inventorApplication

                If invApp Is Nothing OrElse
                   invApp.ActiveDocument Is Nothing OrElse
                   invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then

                    MessageBox.Show("Chức năng này chỉ dùng trong Drawing.",
                                    "Xóa view trùng Part Number",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning)
                    Exit Sub
                End If

                Dim drawingDocument As DrawingDocument =
                    CType(invApp.ActiveDocument, DrawingDocument)

                Dim activeSheet As Sheet = drawingDocument.ActiveSheet
                Dim baseViews As New Dictionary(Of String, DrawingView)(
                    StringComparer.OrdinalIgnoreCase)

                Dim keptPartNumbers As New HashSet(Of String)(
                    StringComparer.OrdinalIgnoreCase)

                Dim duplicateBaseNames As New HashSet(Of String)(
                    StringComparer.OrdinalIgnoreCase)

                ' Tìm Base View trùng Part Number.
                For Each drawingView As DrawingView In activeSheet.DrawingViews
                    Dim baseView As DrawingView = GetBaseView(drawingView)

                    ' Mỗi base view chỉ kiểm tra một lần.
                    If baseViews.ContainsKey(baseView.Name) Then Continue For
                    baseViews.Add(baseView.Name, baseView)

                    Dim partNumber As String = GetPartNumber(baseView)
                    If String.IsNullOrWhiteSpace(partNumber) Then Continue For

                    ' Giữ base view đầu tiên; các base view sau cùng Part Number sẽ xóa.
                    If keptPartNumbers.Contains(partNumber) Then
                        duplicateBaseNames.Add(baseView.Name)
                    Else
                        keptPartNumbers.Add(partNumber)
                    End If
                Next

                If duplicateBaseNames.Count = 0 Then
                    MessageBox.Show("Không có nhóm view nào trùng Part Number.",
                                    "Xóa view trùng Part Number",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information)
                    Exit Sub
                End If

                ' Lấy toàn bộ view con của các base view trùng.
                Dim viewsToDelete As New List(Of DrawingView)

                For Each drawingView As DrawingView In activeSheet.DrawingViews
                    Dim baseView As DrawingView = GetBaseView(drawingView)

                    If duplicateBaseNames.Contains(baseView.Name) Then
                        viewsToDelete.Add(drawingView)
                    End If
                Next

                Dim message As String =
                    "Sẽ xóa " & viewsToDelete.Count & " view thuộc " &
                    duplicateBaseNames.Count & " nhóm bị trùng Part Number." &
                    vbCrLf & vbCrLf &
                    "Mỗi Part Number chỉ giữ lại Base View đầu tiên và các view liên kết với nó." &
                    vbCrLf & vbCrLf &
                    "Tiếp tục?"

                If MessageBox.Show(message,
                                   "Xóa view trùng Part Number",
                                   MessageBoxButtons.YesNo,
                                   MessageBoxIcon.Warning) <> DialogResult.Yes Then
                    Exit Sub
                End If

                ' Xóa view con trước, rồi mới xóa Base View.
                viewsToDelete.Sort(
                    Function(leftView, rightView)
                        Return GetViewDepth(rightView).CompareTo(GetViewDepth(leftView))
                    End Function)

                For Each drawingView As DrawingView In viewsToDelete
                    drawingView.Delete()
                Next

                drawingDocument.Update()

                MessageBox.Show("Đã xóa " & viewsToDelete.Count & " view trùng.",
                                "Xóa view trùng Part Number",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message,
                                "Xóa view trùng Part Number",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)
            End Try
        End Sub

        Private Function GetBaseView(ByVal drawingView As DrawingView) As DrawingView
            Dim result As DrawingView = drawingView

            While result.ParentView IsNot Nothing
                result = result.ParentView
            End While

            Return result
        End Function

        Private Function GetViewDepth(ByVal drawingView As DrawingView) As Integer
            Dim depth As Integer = 0
            Dim currentView As DrawingView = drawingView

            While currentView.ParentView IsNot Nothing
                depth += 1
                currentView = currentView.ParentView
            End While

            Return depth
        End Function

        Private Function GetPartNumber(ByVal baseView As DrawingView) As String
            Try
                Dim descriptor As DocumentDescriptor =
                    baseView.ReferencedDocumentDescriptor

                If descriptor Is Nothing OrElse descriptor.ReferenceMissing Then
                    Return String.Empty
                End If

                Dim modelDocument As Document =
                    CType(descriptor.ReferencedDocument, Document)

                Return CStr(modelDocument.PropertySets.
                    Item("Design Tracking Properties").
                    Item("Part Number").Value).Trim()

            Catch
                Return String.Empty
            End Try
        End Function

    End Module
End Namespace