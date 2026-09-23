Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly3.Buttons
    Public Module Button9
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

                        Dim partDef As PartComponentDefinition = oCompDef

                        ' Bỏ qua bulong / Content Center
                        If partDef.IsContentMember Then Continue For

                        ' Chỉ xử lý Sheet Metal
                        If TypeOf partDef Is SheetMetalComponentDefinition Then

                            Dim smDef As SheetMetalComponentDefinition = partDef

                            ' Xử lý Thickness và thêm chữ "t" phía trước
                            Dim thicknessStr As String = GetThickness(smDef)

                            If thicknessStr <> "" Then
                                WriteCustom(oDoc, "PL", thicknessStr)
                                count = count + 1
                            End If

                        End If

                    Catch
                        Continue For
                    End Try

                Next
            Next

            MessageBox.Show("Hoàn tất!" & vbCrLf &
                    "Đã cập nhật Thickness cho " & count & " tấm Sheet Metal.",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub


        ' =============================================
        ' HÀM XỬ LÝ THICKNESS + THÊM CHỮ "t"
        ' =============================================
        Function GetThickness(smDef As SheetMetalComponentDefinition) As String
            Try
                Dim t_cm As Double = smDef.Thickness.Value      ' Giá trị nội bộ Inventor = cm
                Dim t_mm As Double = t_cm * 10                  ' Chuyển sang mm

                ' Thêm chữ "t" phía trước
                Return "t" & t_mm.ToString()

                ' Nếu sau này muốn chỉnh cách hiển thị, bạn có thể thay bằng:
                ' Return "t" & Math.Round(t_mm, 1).ToString()     ' Ví dụ: t1.5
                ' Return "t" & Math.Ceiling(t_mm).ToString()      ' Ví dụ: t2

            Catch
                Return ""
            End Try
        End Function


        ' =============================================
        ' Hàm ghi iProperty
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







    End Module
End Namespace
