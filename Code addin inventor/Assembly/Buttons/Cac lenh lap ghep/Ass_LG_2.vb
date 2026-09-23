Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports System.Collections.Generic

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep

    Public Module Ass_LG_2

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim app As Inventor.Application = g_inventorApplication
            If app Is Nothing Then Exit Sub

            Dim asm As AssemblyDocument = Nothing
            Try
                asm = TryCast(app.ActiveEditDocument, AssemblyDocument)
                If asm Is Nothing Then asm = TryCast(app.ActiveDocument, AssemblyDocument)

                If asm Is Nothing Then
                    MessageBox.Show("Vui lòng mở Assembly.", "Update Bolts",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                ' 1. Thu thập tất cả Content Center parts
                Dim ccParts As New List(Of ComponentOccurrence)
                CollectContentCenterParts(asm.ComponentDefinition.Occurrences, ccParts)

                If ccParts.Count = 0 Then
                    MessageBox.Show("Không tìm thấy Content Center part (bolt/washer/nut).",
                                    "Update Bolts", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                ' 2. Select tất cả
                Try : asm.SelectSet.Clear() : Catch : End Try
                System.Windows.Forms.Application.DoEvents()

                Dim selected As Integer = 0
                For Each occ As ComponentOccurrence In ccParts
                    Try
                        asm.SelectSet.Select(occ)
                        selected += 1
                    Catch
                    End Try
                Next

                app.StatusBarText = "Đã select " & selected.ToString() & " Content Center part(s)."

                ' 3. Chạy Refresh Standard Components (Inventor 2025)
                Dim refreshCmd As ControlDefinition = Nothing
                Try
                    refreshCmd = app.CommandManager.ControlDefinitions.Item("CCV2RSCButton")
                Catch
                End Try

                If refreshCmd Is Nothing Then
                    MessageBox.Show("Không tìm thấy lệnh Refresh Standard Components.",
                                    "Update Bolts", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                refreshCmd.Execute()

                ' 4. Update
                Try : asm.Update() : Catch : End Try
                Try : app.ActiveView.Update() : Catch : End Try
                app.StatusBarText = ""

                MessageBox.Show(
                    "HOÀN TẤT!" & vbCrLf & vbCrLf &
                    "Đã select: " & selected.ToString() & " Content Center part(s)." & vbCrLf &
                    "Refresh Standard Components đã được gọi.",
                    "Update Bolts", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                Try : app.StatusBarText = "" : Catch : End Try
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "Update Bolts", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub CollectContentCenterParts(
                occurrences As ComponentOccurrences,
                result As List(Of ComponentOccurrence))

            If occurrences Is Nothing Then Exit Sub

            For Each occ As ComponentOccurrence In occurrences
                Try
                    If occ.Suppressed Then Continue For

                    If IsContentMember(occ) Then
                        result.Add(occ)
                    End If

                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Dim subDef As AssemblyComponentDefinition =
                            TryCast(occ.Definition, AssemblyComponentDefinition)
                        If subDef IsNot Nothing Then
                            CollectContentCenterParts(subDef.Occurrences, result)
                        End If
                    End If
                Catch
                End Try
            Next
        End Sub

        Private Function IsContentMember(occ As ComponentOccurrence) As Boolean
            Try
                If occ.DefinitionDocumentType <> DocumentTypeEnum.kPartDocumentObject Then Return False
                Dim partDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
                If partDef Is Nothing Then Return False
                Return partDef.IsContentMember
            Catch
                Return False
            End Try
        End Function

    End Module

End Namespace