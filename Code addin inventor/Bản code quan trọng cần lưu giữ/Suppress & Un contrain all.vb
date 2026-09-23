Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.OLD
    Public Module OLD2
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                ManageConstraintsAndJoints()
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Assembly2 - Manage Constraints/Joints")
            End Try
        End Sub

        Private Sub ManageConstraintsAndJoints()
            Dim asmDoc As AssemblyDocument = TryCast(g_inventorApplication.ActiveDocument, AssemblyDocument)
            If asmDoc Is Nothing Then
                MessageBox.Show("Active document is not an assembly.", "Assembly2 - Manage Constraints/Joints")
                Return
            End If

            ' Hỏi người dùng muốn suppress hay unsuppress
            Dim choice As DialogResult = MessageBox.Show("Chọn Yes để Suppress, No để Unsuppress", "Manage Constraints/Joints", MessageBoxButtons.YesNoCancel)
            If choice = DialogResult.Cancel Then Return

            Dim suppress As Boolean = (choice = DialogResult.Yes)

            Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
            Dim suppressedCount As Integer = 0
            Dim unsuppressedCount As Integer = 0

            ' Constraints
            For Each oConstraint As AssemblyConstraint In compDef.Constraints
                oConstraint.Suppressed = suppress
                If suppress Then suppressedCount += 1 Else unsuppressedCount += 1
            Next

            ' Joints
            For Each oJoint As AssemblyJoint In compDef.Joints
                oJoint.Suppressed = suppress
                If suppress Then suppressedCount += 1 Else unsuppressedCount += 1
            Next

            Dim action As String = If(suppress, "Suppressed", "Unsuppressed")
            MessageBox.Show($"{action} {suppressedCount + unsuppressedCount} Constraints & Joints", "Assembly2 - Manage Constraints/Joints")
        End Sub
    End Module

End Namespace

