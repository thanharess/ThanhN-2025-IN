Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode

    Public Module ass_bom_6

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application = Interop.Marshal2.GetActiveObject("Inventor.Application")

            Try
                If invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Chỉ chạy trên Assembly!", "Browser Rearranger", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                Dim oDoc As AssemblyDocument = CType(invApp.ActiveDocument, AssemblyDocument)
                Dim oPane As BrowserPane = oDoc.BrowserPanes.Item("AmBrowserArrangement")
                Dim oTopNode As BrowserNode = oPane.TopNode

                ' 1. Xóa folder trống
                DeleteEmptyFolders(oTopNode)

                ' 2. Gom Content Center / Purchased
                GroupContentCenter(oDoc, oPane)

                ' 3. Sắp xếp
                SortBrowserCorrectly(oDoc, oPane)

                MessageBox.Show("Đã sắp xếp xong!" & vbCrLf &
                                "1. Cụm lắp (chữ cái + tên ngắn ưu tiên)" & vbCrLf &
                                "2. Part (chữ cái + tên ngắn ưu tiên)" & vbCrLf &
                                "3. Vật tư mua / Content Center",
                                "Browser Rearranger", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message, "Browser Rearranger", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        '=========================================================
        ' XÓA FOLDER TRỐNG
        '=========================================================
        Private Sub DeleteEmptyFolders(oNode As BrowserNode)
            For i As Integer = oNode.BrowserNodes.Count To 1 Step -1
                Dim child As BrowserNode = oNode.BrowserNodes.Item(i)

                If TypeOf child.NativeObject Is BrowserFolder Then
                    If child.BrowserNodes.Count = 0 Then
                        Try
                            Dim oFolder As BrowserFolder = CType(child.NativeObject, BrowserFolder)
                            oFolder.Delete()
                        Catch
                        End Try
                    Else
                        DeleteEmptyFolders(child)
                    End If
                End If
            Next
        End Sub

        '=========================================================
        ' GOM CONTENT CENTER / PURCHASED
        '=========================================================
        Private Sub GroupContentCenter(oDoc As AssemblyDocument, oPane As BrowserPane)
            Dim oTopNode As BrowserNode = oPane.TopNode
            Dim ccFolder As BrowserFolder = Nothing

            For Each n As BrowserNode In oTopNode.BrowserNodes
                If TypeOf n.NativeObject Is BrowserFolder Then
                    If n.BrowserNodeDefinition.Label = "Content Center" Then
                        ccFolder = CType(n.NativeObject, BrowserFolder)
                        Exit For
                    End If
                End If
            Next

            If ccFolder Is Nothing Then
                Try
                    ccFolder = oTopNode.BrowserFolders.Add("Content Center")
                Catch
                    Exit Sub
                End Try
            End If

            For Each occ As ComponentOccurrence In oDoc.ComponentDefinition.Occurrences
                Try
                    If occ.IsContentMember OrElse IsPurchased(occ) Then
                        Dim oNode As BrowserNode = oPane.GetBrowserNodeFromObject(occ)
                        If oNode IsNot Nothing Then
                            ccFolder.Add(oNode)
                        End If
                    End If
                Catch
                End Try
            Next
        End Sub

        Private Function IsPurchased(occ As ComponentOccurrence) As Boolean
            Try
                Dim desc As String = occ.Definition.Document.PropertySets _
                    .Item("Design Tracking Properties").Item("Description").Value.ToString().ToLower()

                If desc.Contains("purchased") OrElse desc.Contains("content center") OrElse desc.Contains("standard") Then
                    Return True
                End If
            Catch
            End Try
            Return False
        End Function

        '=========================================================
        ' HÀM SO SÁNH TÊN NÂNG CAO
        ' 1. Theo chữ cái đầu (A-Z)
        ' 2. Cùng chữ cái đầu → tên ngắn hơn ưu tiên
        ' 3. Cùng độ dài → so sánh đầy đủ
        '=========================================================
        Private Function CompareNameAdvanced(name1 As String, name2 As String) As Integer
            name1 = name1.Trim()
            name2 = name2.Trim()

            If name1 = "" AndAlso name2 = "" Then Return 0
            If name1 = "" Then Return -1
            If name2 = "" Then Return 1

            Dim first1 As String = name1.Substring(0, 1).ToUpper()
            Dim first2 As String = name2.Substring(0, 1).ToUpper()

            Dim cmpFirst As Integer = String.Compare(first1, first2, True)
            If cmpFirst <> 0 Then Return cmpFirst

            ' Cùng chữ cái đầu → ưu tiên tên ngắn hơn
            If name1.Length < name2.Length Then Return -1
            If name1.Length > name2.Length Then Return 1

            ' Cùng độ dài → so sánh đầy đủ
            Return String.Compare(name1, name2, True)
        End Function

        '=========================================================
        ' SẮP XẾP CHÍNH
        '=========================================================
        Private Sub SortBrowserCorrectly(oDoc As AssemblyDocument, oPane As BrowserPane)
            Dim oTopNode As BrowserNode = oPane.TopNode

            Dim asmList As New List(Of BrowserNode)
            Dim partList As New List(Of BrowserNode)

            ' Phân loại
            For Each n As BrowserNode In oTopNode.BrowserNodes
                If TypeOf n.NativeObject Is ComponentOccurrence Then
                    Dim occ As ComponentOccurrence = CType(n.NativeObject, ComponentOccurrence)

                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        asmList.Add(n)
                    ElseIf occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        partList.Add(n)
                    End If
                End If
            Next

            ' Sắp xếp nâng cao
            asmList.Sort(Function(a, b) CompareNameAdvanced(a.BrowserNodeDefinition.Label, b.BrowserNodeDefinition.Label))
            partList.Sort(Function(a, b) CompareNameAdvanced(a.BrowserNodeDefinition.Label, b.BrowserNodeDefinition.Label))

            ' Tìm node Origin
            Dim originNode As BrowserNode = Nothing
            For Each n As BrowserNode In oTopNode.BrowserNodes
                If n.BrowserNodeDefinition.Label = "Origin" Then
                    originNode = n
                    Exit For
                End If
            Next
            If originNode Is Nothing Then Exit Sub

            ' Đưa lên theo thứ tự ngược để giữ đúng thứ tự đã sort
            For i As Integer = partList.Count - 1 To 0 Step -1
                Try
                    oPane.Reorder(originNode, False, partList(i))
                Catch
                End Try
            Next

            For i As Integer = asmList.Count - 1 To 0 Step -1
                Try
                    oPane.Reorder(originNode, False, asmList(i))
                Catch
                End Try
            Next
        End Sub

    End Module

End Namespace