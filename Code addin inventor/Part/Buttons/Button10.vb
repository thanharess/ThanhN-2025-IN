Imports Inventor

Namespace ToolInventor2025.Part.Buttons
    Public Module Button10
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Executed Part Action 10")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Part Action 10: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
