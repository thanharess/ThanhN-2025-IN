Option Explicit On
Imports Inventor
Imports System.Windows.Forms
Imports System.Drawing

Namespace ToolInventor2025.Assembly.Buttons.Part

    '=====================================================
    ' MODULE RIÊNG: CHỈ ẨN SHEET METAL
    '=====================================================
    Public Module Ass_LG_SheetMetal

        '-------------------------------------------------
        ' ENTRY POINT
        '-------------------------------------------------
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application =
                Interop.Marshal2.GetActiveObject("Inventor.Application")

            If invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Vui lòng mở Assembly trước!", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            Dim asmDoc As AssemblyDocument = invApp.ActiveEditDocument
            Dim oDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

            '=================================================
            ' FORM
            '=================================================
            Dim frm As New Form()
            frm.Text = "Ẩn Sheet Metal - Inventor 2020"
            frm.Size = New Size(360, 200)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.Font = New Font("Segoe UI", 9)

            Dim btnHide As New Button() With {
                .Text = "Ẩn Sheet Metal",
                .Location = New System.Drawing.Point(40, 20),
                .Size = New Size(270, 40)
            }

            Dim btnShow As New Button() With {
                .Text = "Hiện lại tất cả Sheet Metal",
                .Location = New System.Drawing.Point(40, 70),
                .Size = New Size(270, 40),
                .BackColor = System.Drawing.Color.LightGreen
            }

            Dim lblInfo As New Label() With {
                .Text = "Quét toàn bộ Assembly (đệ quy)." & vbCrLf &
                        "Bỏ qua: Referent, Phantom, Purchased, Content Center, Reference Part.",
                .Location = New System.Drawing.Point(40, 120),
                .Size = New Size(270, 40),
                .ForeColor = System.Drawing.Color.Gray
            }

            AddHandler btnHide.Click, Sub()
                                          Dim c As Integer = HideSheetMetal(oDef)
                                          MessageBox.Show("Đã ẩn " & c & " Sheet Metal", "Hoàn tất")
                                          frm.Close()
                                      End Sub

            AddHandler btnShow.Click, Sub()
                                          Dim c As Integer = ShowSheetMetal(oDef)
                                          MessageBox.Show("Đã hiện lại " & c & " Sheet Metal", "Hoàn tất")
                                          frm.Close()
                                      End Sub

            frm.Controls.AddRange({btnHide, btnShow, lblInfo})
            frm.ShowDialog()

            asmDoc.Update2(True)

        End Sub

        '=================================================
        ' ẨN SHEET METAL (ĐỆ QUY)
        '=================================================
        Private Function HideSheetMetal(ByVal oDef As AssemblyComponentDefinition) As Integer
            Dim count As Integer = 0
            HideRecursive(oDef.Occurrences, count)
            Return count
        End Function

        Private Sub HideRecursive(ByVal occs As ComponentOccurrences, ByRef count As Integer)
            For Each occ As ComponentOccurrence In occs
                Try
                    If occ.Suppressed Then Continue For

                    If IsSheetMetal(occ) Then
                        occ.Visible = False
                        count += 1
                    End If

                    ' Đệ quy vào sub-assembly
                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Try
                            Dim subDef As AssemblyComponentDefinition =
                                CType(occ.Definition, AssemblyComponentDefinition)
                            HideRecursive(subDef.Occurrences, count)
                        Catch
                        End Try
                    End If

                Catch
                End Try
            Next
        End Sub

        '=================================================
        ' HIỆN LẠI SHEET METAL (ĐỆ QUY)
        '=================================================
        Private Function ShowSheetMetal(ByVal oDef As AssemblyComponentDefinition) As Integer
            Dim count As Integer = 0
            ShowRecursive(oDef.Occurrences, count)
            Return count
        End Function

        Private Sub ShowRecursive(ByVal occs As ComponentOccurrences, ByRef count As Integer)
            For Each occ As ComponentOccurrence In occs
                Try
                    If IsSheetMetal(occ) Then
                        occ.Visible = True
                        count += 1
                    End If

                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Try
                            Dim subDef As AssemblyComponentDefinition =
                                CType(occ.Definition, AssemblyComponentDefinition)
                            ShowRecursive(subDef.Occurrences, count)
                        Catch
                        End Try
                    End If

                Catch
                End Try
            Next
        End Sub

        '=================================================
        ' KIỂM TRA: CÓ PHẢI SHEET METAL
        ' Điều kiện:
        '   - Là PartComponentDefinition
        '   - LÀ Sheet Metal
        '   - KHÔNG phải Reference Part
        '   - KHÔNG phải Content Center member
        '   - BOM không thuộc {Reference, Phantom, Purchased}
        '=================================================
        Private Function IsSheetMetal(ByVal occ As ComponentOccurrence) As Boolean
            If occ Is Nothing Then Return False

            Try
                If occ.Suppressed Then Return False

                ' 1) Phải là Part
                Dim pDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
                If pDef Is Nothing Then Return False

                ' 2) BẮT BUỘC phải là Sheet Metal
                If Not pDef.IsSheetMetal Then Return False

                ' 3) Loại trừ Reference Part
                If SafeIsReferencePart(pDef) Then Return False

                ' 4) Loại trừ Content Center member
                If SafeIsContentMember(occ) Then Return False

                ' 5) Loại trừ theo BOM Structure
                Select Case occ.BOMStructure
                    Case BOMStructureEnum.kReferenceBOMStructure,
                         BOMStructureEnum.kPhantomBOMStructure,
                         BOMStructureEnum.kPurchasedBOMStructure
                        Return False
                End Select

                Return True

            Catch
                Return False
            End Try
        End Function

        '=================================================
        ' HELPERS — bọc try/catch riêng
        '=================================================
        Private Function SafeIsReferencePart(ByVal pDef As PartComponentDefinition) As Boolean
            If pDef Is Nothing Then Return False
            Try
                Return pDef.IsReferencePart
            Catch
                Return False
            End Try
        End Function

        Private Function SafeIsContentMember(ByVal occ As ComponentOccurrence) As Boolean
            If occ Is Nothing Then Return False
            Try
                Return True = occ.IsContentMember
            Catch
                Return False
            End Try
        End Function

    End Module
End Namespace