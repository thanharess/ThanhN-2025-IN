Imports Inventor

Namespace ToolInventor2025.Assembly2.Buttons.BOMcode
    Public Module Ass_Bom_14
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Executed Part Action 14")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Part Action 14: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
