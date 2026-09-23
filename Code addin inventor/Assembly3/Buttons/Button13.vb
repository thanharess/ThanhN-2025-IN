Imports Inventor

Namespace ToolInventor2025.Assembly3.Buttons
    Public Module Button13
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Executed Part Action 13")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Part Action 13: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
