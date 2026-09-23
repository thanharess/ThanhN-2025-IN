Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep.constraint
    Public Module Ass_LG_C_1a
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                Dim asmDoc As AssemblyDocument = TryCast(g_inventorApplication.ActiveDocument, AssemblyDocument)
                If asmDoc Is Nothing Then
                    MessageBox.Show("Hãy mở Assembly (.iam) trước khi chạy!", "ThanhN", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Select Case ShowTopMenu()
                    Case 1
                        ManageSuppressState(asmDoc)
                    Case 2
                        ManageDeleteAndGround(asmDoc)
                End Select
            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "ThanhN", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        'LỰA CHỌN 1: Suppress / Unsuppress Constraint + Joint.
        Private Sub ManageSuppressState(ByVal asmDoc As AssemblyDocument)
            Dim choice As DialogResult = MessageBox.Show(
                "Yes = Suppress tất cả Constraint và Joint." & vbCrLf &
                "No = Unsuppress tất cả Constraint và Joint.",
                "Manage Constraints / Joints", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If choice = DialogResult.Cancel Then Return

            Dim suppress As Boolean = (choice = DialogResult.Yes)
            Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
            Dim count As Integer = 0

            For Each constraint As AssemblyConstraint In compDef.Constraints
                Try
                    constraint.Suppressed = suppress
                    count += 1
                Catch
                End Try
            Next

            For Each joint As AssemblyJoint In compDef.Joints
                Try
                    joint.Suppressed = suppress
                    count += 1
                Catch
                End Try
            Next

            asmDoc.Update()
            MessageBox.Show(If(suppress, "Đã Suppress ", "Đã Unsuppress ") & count & " Constraints & Joints.", "ThanhN", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        'LỰA CHỌN 2: menu xóa / Ground.
        Private Sub ManageDeleteAndGround(ByVal asmDoc As AssemblyDocument)
            Dim action As Integer = ShowDeleteGroundMenu()
            If action = 0 Then Return

            Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
            Dim constraintCount As Integer = 0
            Dim jointCount As Integer = 0

            Select Case action
                Case 1 'Xóa cả Constraint và Joint
                    If Not Confirm("Xóa tất cả Constraint và Joint?") Then Return
                    DeleteAllConstraints(compDef, constraintCount)
                    DeleteAllJoints(compDef, jointCount)
                    MessageBox.Show("Đã xóa:" & vbCrLf & "- " & constraintCount & " Constraints" & vbCrLf & "- " & jointCount & " Joints", "ThanhN")

                Case 2 'Xóa Constraint
                    If Not Confirm("Xóa tất cả Constraint?") Then Return
                    DeleteAllConstraints(compDef, constraintCount)
                    MessageBox.Show("Đã xóa " & constraintCount & " Constraints.", "ThanhN")

                Case 3 'Xóa Joint
                    If Not Confirm("Xóa tất cả Joint?") Then Return
                    DeleteAllJoints(compDef, jointCount)
                    MessageBox.Show("Đã xóa " & jointCount & " Joints.", "ThanhN")

                Case 4 'Ground / Unground
                    SetGroundFromMenu(compDef)

                Case 5 'Xóa cả hai rồi Ground / Unground
                    If Not Confirm("Xóa tất cả Constraint và Joint, sau đó Ground/Unground toàn bộ Assembly?") Then Return
                    DeleteAllConstraints(compDef, constraintCount)
                    DeleteAllJoints(compDef, jointCount)
                    SetGroundFromMenu(compDef, constraintCount, jointCount)
            End Select

            asmDoc.Update()
        End Sub

        Private Sub SetGroundFromMenu(ByVal compDef As AssemblyComponentDefinition, Optional ByVal constraintCount As Integer = -1, Optional ByVal jointCount As Integer = -1)
            Dim groundMode As Integer = ShowGroundMenu()
            If groundMode = 0 Then Return

            Dim count As Integer = 0
            Dim ground As Boolean = (groundMode = 1)
            SetGroundRecursive(compDef.Occurrences, ground, count)

            Dim message As String = If(ground, "Đã Ground ", "Đã bỏ Ground ") & count & " components."
            If constraintCount >= 0 Then
                message = "Hoàn tất!" & vbCrLf & "Constraints: " & constraintCount & vbCrLf & "Joints: " & jointCount & vbCrLf & If(ground, "Ground: ", "Unground: ") & count
            End If
            MessageBox.Show(message, "ThanhN", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Function Confirm(ByVal message As String) As Boolean
            Return MessageBox.Show(message, "ThanhN", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes
        End Function

        Private Sub DeleteAllConstraints(ByVal compDef As AssemblyComponentDefinition, ByRef count As Integer)
            For i As Integer = compDef.Constraints.Count To 1 Step -1
                Try
                    compDef.Constraints.Item(i).Delete()
                    count += 1
                Catch
                End Try
            Next
        End Sub

        Private Sub DeleteAllJoints(ByVal compDef As AssemblyComponentDefinition, ByRef count As Integer)
            For i As Integer = compDef.Joints.Count To 1 Step -1
                Try
                    compDef.Joints.Item(i).Delete()
                    count += 1
                Catch
                End Try
            Next
        End Sub

        Private Sub SetGroundRecursive(ByVal occurrences As ComponentOccurrences, ByVal value As Boolean, ByRef count As Integer)
            For Each occurrence As ComponentOccurrence In occurrences
                Try
                    occurrence.Grounded = value
                    count += 1
                Catch
                End Try
                Try
                    If occurrence.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        SetGroundRecursive(occurrence.SubOccurrences, value, count)
                    End If
                Catch
                End Try
            Next
        End Sub

        Private Function ShowTopMenu() As Integer
            Return ShowMenu("Assembly - Constraint / Joint", "CHỌN NHÓM CHỨC NĂNG", New String() {
                "1. Suppress / Unsuppress Constraint + Joint",
                "2. Xóa Constraint / Joint / Ground"})
        End Function

        Private Function ShowDeleteGroundMenu() As Integer
            Return ShowMenu("Xóa Constraint / Joint / Ground", "CHỌN CHỨC NĂNG", New String() {
                "Xóa ALL Constraint + Joint", "Xóa ALL Constraint", "Xóa ALL Joint",
                "Ground / Bỏ Ground ALL", "Xóa Constraint + Joint + Ground"})
        End Function

        Private Function ShowGroundMenu() As Integer
            Return ShowMenu("Kiểu Ground", "CHỌN KIỂU GROUND", New String() {"GROUND ALL", "BỎ GROUND ALL"})
        End Function

        Private Function ShowMenu(ByVal title As String, ByVal caption As String, ByVal options As String()) As Integer
            Dim result As Integer = 0
            Using form As New Form()
                form.Text = title
                form.Width = 460
                form.Height = 125 + options.Length * 52
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False

                Dim label As New Label() With {.Text = caption, .Left = 20, .Top = 15, .Width = 400, .Height = 25}
                form.Controls.Add(label)

                For i As Integer = 0 To options.Length - 1
                    Dim optionIndex As Integer = i + 1
                    Dim button As New Button() With {.Text = options(i), .Left = 20, .Top = 45 + i * 50, .Width = 400, .Height = 42}
                    AddHandler button.Click, Sub()
                                                 result = optionIndex
                                                 form.Close()
                                             End Sub
                    form.Controls.Add(button)
                Next

                form.ShowDialog()
            End Using
            Return result
        End Function

    End Module

End Namespace

