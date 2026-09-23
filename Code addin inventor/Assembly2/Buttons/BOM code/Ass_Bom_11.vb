Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode
    Public Module Ass_Bom_11
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

            ' Bật Parts Only
            compDef.BOM.PartsOnlyViewEnabled = True

            Dim oBOMView As BOMView
            oBOMView = compDef.BOM.BOMViews.Item("Parts Only")

            Dim oRow As BOMRow

            For Each oRow In oBOMView.BOMRows

                Dim itemNum As String = oRow.ItemNumber
                Dim qty As String = oRow.ItemQuantity

                For Each oCompDef In oRow.ComponentDefinitions

                    Try
                        Dim oDoc As Document = oCompDef.Document

                        ' Chỉ xử lý PART
                        If oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then

                            Dim partDef As PartComponentDefinition
                            partDef = oCompDef

                            ' ❌ bỏ bulong
                            If partDef.IsContentMember Then Continue For

                            ' Ghi ItemNumber
                            WriteProp(oDoc, "item1", itemNum)

                            ' Ghi Quantity (đã cộng)
                            WriteProp(oDoc, "SL Part", qty)

                        End If

                    Catch
                    End Try

                Next

            Next

            MessageBox.Show("OK - Parts Only BOM DONE")

        End Sub


        ' ===== FUNCTION =====
        Sub WriteProp(oDoc As Document, propName As String, value As String)

            Try
                Dim customSet As PropertySet
                customSet = oDoc.PropertySets.Item("Inventor User Defined Properties")

                Dim prop As Inventor.Property

                Try
                    prop = customSet.Item(propName)
                    prop.Value = value
                Catch
                    customSet.Add(value, propName)
                End Try

            Catch
            End Try

        End Sub







    End Module
End Namespace
