Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Part
    Public Module Ass_Part_3




        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try

                ChangeGenericToSteel()

            Catch ex As Exception

                MessageBox.Show(
            "Error: " & ex.Message,
            "Assembly2 - Change Material",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error)

            End Try

        End Sub


        Private Sub ChangeGenericToSteel()

            '==========================================================
            ' Lấy Assembly đang active
            '==========================================================
            Dim oAsmDoc As AssemblyDocument =
        TryCast(
            g_inventorApplication.ActiveDocument,
            AssemblyDocument)

            If oAsmDoc Is Nothing Then

                MessageBox.Show(
            "Active document is not an assembly.",
            "Assembly2 - Change Material",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning)

                Return

            End If


            Dim changedCount As Integer = 0
            Dim errorCount As Integer = 0


            '==========================================================
            ' Duyệt tất cả Document được Assembly tham chiếu
            '==========================================================
            For Each oDoc As Document In oAsmDoc.AllReferencedDocuments

                If oDoc.DocumentType =
            DocumentTypeEnum.kPartDocumentObject Then

                    Try

                        Dim oPartDoc As PartDocument =
                    TryCast(oDoc, PartDocument)

                        If oPartDoc Is Nothing Then
                            Continue For
                        End If


                        '==================================================
                        ' Lấy Material hiện tại của Part
                        '==================================================
                        Dim currentMat As Material = Nothing

                        Try

                            currentMat =
                        oPartDoc.ComponentDefinition.Material

                        Catch

                            currentMat = Nothing

                        End Try


                        If currentMat Is Nothing Then
                            Continue For
                        End If


                        '==================================================
                        ' Kiểm tra Generic
                        '==================================================
                        If currentMat.Name.IndexOf(
                    "Generic",
                    StringComparison.OrdinalIgnoreCase) >= 0 Then


                            '================================================
                            ' Tìm Steel, Mild trong chính Part
                            '================================================
                            Dim steelMaterial As Material = Nothing

                            Try

                                steelMaterial =
                            oPartDoc.Materials.Item("Steel, Mild")

                            Catch

                                steelMaterial = Nothing

                            End Try


                            '================================================
                            ' Nếu Part đã có Steel, Mild
                            '================================================
                            If steelMaterial IsNot Nothing Then

                                oPartDoc.ComponentDefinition.Material =
                            steelMaterial

                                oPartDoc.Update()

                                changedCount += 1

                            Else

                                ' Không có Steel, Mild trong Part
                                errorCount += 1

                            End If

                        End If


                    Catch

                        errorCount += 1

                    End Try

                End If

            Next


            '==========================================================
            ' Update Assembly
            '==========================================================
            oAsmDoc.Update()


            '==========================================================
            ' Thông báo
            '==========================================================
            Dim msg As String =
        "Đã kiểm tra toàn bộ Part." &
        vbCrLf & vbCrLf &
        "Generic → Steel, Mild: " &
        changedCount.ToString()

            If errorCount > 0 Then

                msg &= vbCrLf &
               "Không xử lý được: " &
               errorCount.ToString()
            End If


            MessageBox.Show(
        msg,
        "Assembly2 - Change Material",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information)

        End Sub


    End Module
End Namespace
