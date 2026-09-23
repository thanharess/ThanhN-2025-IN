Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep.constraint
    Public Module Ass_LG_C_1e

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim oApp As Inventor.Application = g_inventorApplication
            Dim activeDoc As Document = Nothing

            Try
                Try
                    activeDoc = oApp.ActiveDocument
                Catch
                    activeDoc = Nothing
                End Try

                If activeDoc Is Nothing Then
                    MessageBox.Show("Không có document nào đang mở.", "Thông báo",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                If activeDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    DeleteAssemblyConstraints(CType(activeDoc, AssemblyDocument))
                ElseIf activeDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    DeletePartConstraints(CType(activeDoc, PartDocument))
                Else
                    MessageBox.Show("Chỉ hỗ trợ Assembly hoặc Part document.", "Thông báo",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

            Catch ex As Exception
                MessageBox.Show("CÓ LỖI:" & vbCrLf & vbCrLf &
                                ex.Message & vbCrLf & vbCrLf & ex.StackTrace,
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        '=====================================================
        ' XÓA ASSEMBLY CONSTRAINTS LỖI (SICK)
        '=====================================================
        Private Sub DeleteAssemblyConstraints(ByVal asmDoc As AssemblyDocument)
            Try
                Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
                Dim constraints As AssemblyConstraints = asmDef.Constraints

                Dim errorConstraints As New List(Of AssemblyConstraint)()

                For Each constraint As AssemblyConstraint In constraints
                    Try
                        ' Chuẩn: mọi trạng thái khác UpToDate và Suppressed đều coi là sick
                        If constraint.HealthStatus <> HealthStatusEnum.kUpToDateHealth AndAlso
                           constraint.HealthStatus <> HealthStatusEnum.kSuppressedHealth Then
                            errorConstraints.Add(constraint)
                        End If
                    Catch
                        ' Bỏ qua constraint không đọc được
                    End Try
                Next

                Dim message As String =
                    "Tổng số constraints: " & constraints.Count.ToString() & vbCrLf &
                    "Constraints lỗi (phát hiện): " & errorConstraints.Count.ToString() & vbCrLf & vbCrLf

                If errorConstraints.Count = 0 Then
                    MessageBox.Show("Không tìm thấy constraint lỗi nào.", "Thông báo",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                message &= "Bạn có muốn xóa " & errorConstraints.Count.ToString() & " constraints lỗi?"

                Dim result As DialogResult =
                    MessageBox.Show(message, "Xác nhận xóa",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If result <> DialogResult.Yes Then Return

                Dim deletedCount As Integer = 0
                Dim failCount As Integer = 0

                For Each constraint As AssemblyConstraint In errorConstraints
                    Try
                        constraint.Delete()
                        deletedCount += 1
                    Catch
                        failCount += 1
                    End Try
                Next

                asmDoc.Update2(True)

                MessageBox.Show("KẾT QUẢ XÓA:" & vbCrLf & vbCrLf &
                                "Đã xóa: " & deletedCount.ToString() & " constraints" & vbCrLf &
                                "Không thể xóa: " & failCount.ToString() & " constraints",
                                "Hoàn tất",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi xử lý Assembly constraints:" & vbCrLf & ex.Message,
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        '=====================================================
        ' XÓA PART CONSTRAINTS LỖI (2D SKETCH)
        '=====================================================
        Private Sub DeletePartConstraints(ByVal partDoc As PartDocument)
            Try
                Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition
                Dim deletedCount As Integer = 0
                Dim failCount As Integer = 0

                For Each sketch As PlanarSketch In partDef.Sketches
                    Try
                        Dim geomConstraints As GeometricConstraints = sketch.GeometricConstraints
                        Dim constraintsToDelete As New List(Of GeometricConstraint)()

                        For Each geomConstraint As GeometricConstraint In geomConstraints
                            Try
                                If geomConstraint.HealthStatus <> HealthStatusEnum.kUpToDateHealth AndAlso
                                   geomConstraint.HealthStatus <> HealthStatusEnum.kSuppressedHealth Then
                                    constraintsToDelete.Add(geomConstraint)
                                End If
                            Catch
                            End Try
                        Next

                        For Each constraint As GeometricConstraint In constraintsToDelete
                            Try
                                constraint.Delete()
                                deletedCount += 1
                            Catch
                                failCount += 1
                            End Try
                        Next

                    Catch
                        ' Bỏ qua sketch lỗi
                    End Try
                Next

                partDoc.Update2(True)

                MessageBox.Show("KẾT QUẢ XÓA 2D SKETCH CONSTRAINTS:" & vbCrLf & vbCrLf &
                                "Đã xóa: " & deletedCount.ToString() & " constraints" & vbCrLf &
                                "Không thể xóa: " & failCount.ToString() & " constraints",
                                "Hoàn tất",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi xử lý Part constraints:" & vbCrLf & ex.Message,
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

    End Module
End Namespace