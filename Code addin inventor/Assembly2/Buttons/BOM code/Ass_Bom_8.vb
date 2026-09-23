Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode
    Public Module Ass_Bom_8
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim invApp As Inventor.Application
            Try
                invApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch ex As Exception
                MessageBox.Show("Inventor chưa chạy.")
                Return
            End Try

            ' Lấy Assembly đang mở
            Dim doc As AssemblyDocument = TryCast(invApp.ActiveDocument, AssemblyDocument)
            If doc Is Nothing Then
                MessageBox.Show("Không phải Assembly Document.")
                Return
            End If

            Dim compDef As AssemblyComponentDefinition = doc.ComponentDefinition
            compDef.BOM.PartsOnlyViewEnabled = True
            Dim oBOMView As BOMView = compDef.BOM.BOMViews.Item("Parts Only")

            Dim count As Integer = 0

            For Each oRow As BOMRow In oBOMView.BOMRows
                For Each oCompDef As ComponentDefinition In oRow.ComponentDefinitions
                    Try
                        Dim oDoc As Document = oCompDef.Document
                        If oDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then Continue For

                        Dim partDef As PartComponentDefinition = CType(oCompDef, PartComponentDefinition)
                        If partDef.IsContentMember Then Continue For

                        If TypeOf partDef Is SheetMetalComponentDefinition Then
                            Dim smDef As SheetMetalComponentDefinition = CType(partDef, SheetMetalComponentDefinition)

                            Dim thicknessStr As String = GetThickness(smDef)
                            If thicknessStr <> "" Then
                                WriteCustom(oDoc, "PL", thicknessStr)
                            End If

                            Dim stockStr As String = GetStockNumber(smDef)
                            If stockStr <> "" Then
                                WriteStockNumber(oDoc, stockStr)
                            End If

                            count += 1
                        End If
                    Catch
                        Continue For
                    End Try
                Next
            Next

            MessageBox.Show("Hoàn tất!" & vbCrLf &
                            "Đã xử lý " & count & " tấm Sheet Metal.",
                            "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        ' =============================================
        ' HÀM THICKNESS - Có chữ "t"
        ' =============================================
        Function GetThickness(smDef As SheetMetalComponentDefinition) As String
            Try
                Dim t_cm As Double = smDef.Thickness.Value
                Dim t_mm As Double = t_cm * 10

                Return "t" & t_mm.ToString()          ' → t3, t1.5, t2 ...

            Catch
                Return ""
            End Try
        End Function


        ' =============================================
        ' HÀM STOCK NUMBER - Không có chữ "t"
        ' =============================================
        Function GetStockNumber(smDef As SheetMetalComponentDefinition) As String
            Try
                ' Tạo Flat Pattern nếu chưa có
                If Not smDef.HasFlatPattern Then
                    Try
                        smDef.Unfold()
                    Catch
                        Return ""
                    End Try
                End If

                Dim flat As FlatPattern = smDef.FlatPattern

                Dim L As Double = flat.Length * 10     ' cm -> mm
                Dim W As Double = flat.Width * 10

                ' Đảm bảo L là chiều dài lớn hơn
                If W > L Then
                    Dim temp As Double = L
                    L = W
                    W = temp
                End If

                ' Lấy độ dày không có chữ "t" để đưa vào Stock Number
                Dim t_cm As Double = smDef.Thickness.Value
                Dim t_mm As Double = t_cm * 10
                Dim thicknessNoT As String = t_mm.ToString()

                ' Stock Number: PL3x1200x800 (không có chữ t)
                Dim stockStr As String = "PL" & thicknessNoT & "x" & Math.Ceiling(L) & "x" & Math.Ceiling(W)

                Return stockStr

            Catch
                Return ""
            End Try
        End Function


        ' =============================================
        ' Hàm hỗ trợ
        ' =============================================
        Sub WriteCustom(oDoc As Document, propName As String, value As String)
            Try
                Dim customSet As PropertySet = oDoc.PropertySets.Item("Inventor User Defined Properties")
                Try
                    customSet.Item(propName).Value = value
                Catch
                    customSet.Add(value, propName)
                End Try
            Catch
            End Try
        End Sub


        Sub WriteStockNumber(oDoc As Document, value As String)
            Try
                Dim designSet As PropertySet = oDoc.PropertySets.Item("Design Tracking Properties")
                designSet.Item("Stock Number").Value = value
            Catch
            End Try
        End Sub
    End Module
End Namespace
