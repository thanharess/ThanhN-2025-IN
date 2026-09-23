Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly3.Buttons
    Public Module Button10
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

            ' Bật BOM
            compDef.BOM.StructuredViewEnabled = True
            compDef.BOM.StructuredViewFirstLevelOnly = False

            Dim oBOMView As BOMView = compDef.BOM.BOMViews.Item(2) ' 2 = Structured

            Dim oRow As BOMRow

            For Each oRow In oBOMView.BOMRows

                Dim itemNum As String = oRow.ItemNumber

                For Each oCompDef In oRow.ComponentDefinitions

                    Try
                        ' Chỉ lấy Part
                        If TypeOf oCompDef Is PartComponentDefinition Then

                            Dim partDef As PartComponentDefinition = oCompDef

                            ' ❌ Bỏ Content Center (bulong, ốc)
                            If partDef.IsContentMember Then Continue For

                            Dim oDoc As Document = partDef.Document

                            ' ✅ TẠO iProperty nếu chưa có
                            Dim customSet As PropertySet
                            customSet = oDoc.PropertySets.Item("Inventor User Defined Properties")

                            Dim prop As Inventor.Property

                            Try
                                prop = customSet.Item("item1")
                                prop.Value = itemNum
                            Catch
                                customSet.Add(itemNum, "item1")
                            End Try

                        End If

                    Catch
                    End Try

                Next

            Next

            MessageBox.Show("Đã copy Item Number xong!")








        End Sub

    End Module
End Namespace
