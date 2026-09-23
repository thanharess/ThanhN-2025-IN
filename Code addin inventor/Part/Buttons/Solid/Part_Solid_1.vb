Imports Inventor

Namespace ToolInventor2025.Part.Buttons.solid
    Public Module Part_Solid_1
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                ' Action1: ask user for tolerance value, then select sketch dimensions and set symmetric tolerance
                Dim input As String = Microsoft.VisualBasic.Interaction.InputBox("Enter symmetric tolerance value (e.g. 0.05)", "Tolerance value", "0.05")
                If String.IsNullOrWhiteSpace(input) Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Part Action 1 cancelled by user")
                    Exit Sub
                End If

                Dim tol As Double
                If Not Double.TryParse(input, tol) Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Invalid tolerance value: " & input)
                    Exit Sub
                End If

                Dim oMsg As String = "Select sketch dimensions to apply default tolerance (Press Esc to continue)"
                While True
                    Dim oDimension As DimensionConstraint = TryCast(g_inventorApplication.CommandManager.Pick(SelectionFilterEnum.kSketchDimConstraintFilter, oMsg), DimensionConstraint)
                    ' If nothing gets selected then exit
                    If oDimension Is Nothing Then Exit While
                    oDimension.Parameter.Tolerance.SetToSymmetric(tol / 10)
                End While

                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Part Action 1 completed")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Part Action 1: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
