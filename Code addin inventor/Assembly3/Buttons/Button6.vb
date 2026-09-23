Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports Microsoft.VisualBasic

Namespace ToolInventor2025.Assembly3.Buttons

    Public Module Button6

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try

                '==================================================
                ' KIỂM TRA DOCUMENT
                '==================================================
                Dim oAsm As AssemblyDocument =
                    TryCast(g_inventorApplication.ActiveDocument, AssemblyDocument)

                If oAsm Is Nothing Then
                    MessageBox.Show(
                        "Rule này chỉ chạy trong Assembly.",
                        "iLogic")
                    Exit Sub
                End If


                '==================================================
                ' BOM
                '==================================================
                Dim oBOM As BOM =
                    oAsm.ComponentDefinition.BOM

                oBOM.StructuredViewEnabled = True
                oBOM.StructuredViewFirstLevelOnly = False

                Dim oBOMView As BOMView =
                    oBOM.BOMViews.Item("Structured")


                '==================================================
                ' LẤY VÀ SẮP XẾP CÁC DÒNG TOP-LEVEL
                '==================================================
                Dim topRows As List(Of BOMRow) =
                    SortRows(oBOMView.BOMRows)


                '==================================================
                ' ĐÁNH STT
                '==================================================
                Dim i As Integer = 1

                For Each row As BOMRow In topRows

                    ' Đánh Item Number trong BOM
                    row.ItemNumber = CStr(i)


                    ' Lấy Document
                    Dim refDoc As Document = Nothing

                    Try

                        refDoc =
                            row.ComponentDefinitions.Item(1).Document

                    Catch

                        i += 1
                        Continue For

                    End Try


                    ' Ghi property item1
                    AddOrUpdateSTT(
                        refDoc,
                        CStr(i))


                    i += 1

                Next


                MessageBox.Show(
                    "Hoàn tất." & vbCrLf &
                    "Đã sắp xếp và đánh STT cho các đối tượng TOP-LEVEL.",
                    "Assembly BOM")


            Catch ex As Exception

                MessageBox.Show(
                    "Có lỗi khi thực hiện:" & vbCrLf &
                    ex.Message,
                    "Assembly BOM",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub


        '==========================================================
        ' SẮP XẾP BOM
        '
        ' Thứ tự:
        '   1. SUB-ASSEMBLY
        '   2. PART
        '   3. PURCHASED
        '==========================================================
        Private Function SortRows(
            ByVal bomRows As BOMRowsEnumerator) _
            As List(Of BOMRow)

            Dim subAsm As New List(Of BOMRow)
            Dim parts As New List(Of BOMRow)
            Dim purchased As New List(Of BOMRow)


            For Each row As BOMRow In bomRows

                Dim refDoc As Document = Nothing

                Try

                    refDoc =
                        row.ComponentDefinitions.Item(1).Document

                Catch

                    Continue For

                End Try


                Dim partNum As String =
                    GetPartNumber(row)


                '----------------------------------------------
                ' ASSEMBLY
                '----------------------------------------------
                If refDoc.DocumentType =
                    DocumentTypeEnum.kAssemblyDocumentObject Then

                    subAsm.Add(row)


                    '----------------------------------------------
                    ' PART
                    '----------------------------------------------
                ElseIf refDoc.DocumentType =
                    DocumentTypeEnum.kPartDocumentObject Then

                    If IsPurchased(partNum) Then

                        purchased.Add(row)

                    Else

                        parts.Add(row)

                    End If

                End If

            Next


            '==================================================
            ' SORT THEO PART NUMBER
            '==================================================
            subAsm.Sort(
                Function(a, b)
                    Return String.Compare(
                        GetPartNumber(a),
                        GetPartNumber(b),
                        True)
                End Function)


            parts.Sort(
                Function(a, b)
                    Return String.Compare(
                        GetPartNumber(a),
                        GetPartNumber(b),
                        True)
                End Function)


            purchased.Sort(
                Function(a, b)
                    Return String.Compare(
                        GetPartNumber(a),
                        GetPartNumber(b),
                        True)
                End Function)


            '==================================================
            ' GHÉP THỨ TỰ
            '==================================================
            Dim orderedRows As New List(Of BOMRow)

            orderedRows.AddRange(subAsm)
            orderedRows.AddRange(parts)
            orderedRows.AddRange(purchased)


            Return orderedRows

        End Function


        '==========================================================
        ' KIỂM TRA PURCHASED
        '==========================================================
        Private Function IsPurchased(
            ByVal partNum As String) As Boolean

            If String.IsNullOrEmpty(partNum) Then
                Return False
            End If


            Dim s As String =
                partNum.ToUpper()


            If s.Contains("ISO") Then Return True
            If s.Contains("DIN") Then Return True
            If s.Contains("SKF") Then Return True
            If s.Contains("PURCHASED") Then Return True


            Return False

        End Function


        '==========================================================
        ' LẤY PART NUMBER
        '==========================================================
        Private Function GetPartNumber(
            ByVal row As BOMRow) As String

            Try

                Return CStr(
                    row.ComponentDefinitions.Item(1).Document.
                    PropertySets.Item("Design Tracking Properties").
                    Item("Part Number").Value)

            Catch

                Return ""

            End Try

        End Function


        '==========================================================
        ' TẠO / CẬP NHẬT PROPERTY "item1"
        '==========================================================
        Private Sub AddOrUpdateSTT(
            ByVal doc As Document,
            ByVal value As String)

            Try

                If Not doc.IsModifiable Then
                    Exit Sub
                End If


                Dim userProps As PropertySet =
                    doc.PropertySets.Item(
                        "Inventor User Defined Properties")


                Dim sttProp As Inventor.Property = Nothing


                '----------------------------------------------
                ' TÌM ITEM1
                '----------------------------------------------
                For Each p As Inventor.Property In userProps

                    If p.Name.ToLower() = "item1" Then

                        sttProp = p
                        Exit For

                    End If

                Next


                '----------------------------------------------
                ' CHƯA CÓ → TẠO
                '----------------------------------------------
                If sttProp Is Nothing Then

                    userProps.Add(
                        value,
                        "item1")


                    '----------------------------------------------
                    ' ĐÃ CÓ → CẬP NHẬT
                    '----------------------------------------------
                Else

                    sttProp.Value = value

                End If

            Catch

                ' Không làm dừng toàn bộ chương trình
                ' nếu một document không ghi được property

            End Try

        End Sub

    End Module

End Namespace