Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode
    Public Module Ass_Bom_12
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

            For Each oRow As BOMRow In oBOMView.BOMRows
                Dim itemNum As String = oRow.ItemNumber
                Dim qty As String = oRow.ItemQuantity

                For Each oCompDef As ComponentDefinition In oRow.ComponentDefinitions
                    Try
                        Dim oDoc As Document = oCompDef.Document
                        If oDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then Continue For

                        Dim partDef As PartComponentDefinition = CType(oCompDef, PartComponentDefinition)
                        If partDef.IsContentMember Then Continue For   ' Bỏ bulong / Content Center

                        ' Ghi Item + Qty
                        WriteCustom(oDoc, "item1", itemNum)
                        WriteCustom(oDoc, "SL Part", qty)

                        ' ================== SHEET METAL ==================
                        If TypeOf partDef Is SheetMetalComponentDefinition Then
                            Dim smDef As SheetMetalComponentDefinition = CType(partDef, SheetMetalComponentDefinition)

                            ' ===== THICKNESS - KHÔNG LÀM TRÒN =====
                            Dim t_cm As Double = smDef.Thickness.Value
                            Dim t_mm As Double = t_cm * 10
                            Dim thicknessStr As String = t_mm.ToString()
                            WriteCustom(oDoc, "PL", thicknessStr)

                            ' ===== FLAT PATTERN =====
                            If Not smDef.HasFlatPattern Then
                                Try
                                    smDef.Unfold() ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' có thẻ vô hiệu hóa
                                Catch
                                    Continue For
                                End Try
                            End If

                            Dim flat As FlatPattern = smDef.FlatPattern
                            Dim L As Double = "" 'flat.Length * 10
                            Dim W As Double = "" 'flat.Width * 10

                            If W > L Then
                                Dim temp As Double = L
                                L = W
                                W = temp
                            End If

                            ' ===== STOCK NUMBER =====
                            Dim stockStr As String = "PL" & thicknessStr & "x" & Math.Ceiling(L) & "x" & Math.Ceiling(W)
                            WriteStockNumber(oDoc, stockStr)
                        End If
                    Catch
                        Continue For
                    End Try
                Next
            Next

            MessageBox.Show("Hoàn tất!" & vbCrLf &
                            "• Thickness: giữ nguyên giá trị (không làm tròn)" & vbCrLf &
                            "• Chiều dài & rộng: đã Ceiling",
                            "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

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
