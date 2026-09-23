Imports Inventor

Namespace ToolInventor2025.Part.Buttons.solid
    Public Module Part_Solid_4
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                g_inventorApplication.StatusBarText = "Thành Design"

                Dim app As Inventor.Application = g_inventorApplication
                Dim doc As Document = app.ActiveDocument

                ' Pick a solid body
                Dim picked As Object = app.CommandManager.Pick(SelectionFilterEnum.kPartBodyFilter, "chọn solid")
                If picked Is Nothing Then
                    MsgBox("Chưa chọn")
                    Exit Sub
                End If

                ' Clear and select the picked item
                doc.SelectSet.Clear()
                doc.SelectSet.Select(picked)

                Dim sel As SelectSet = doc.SelectSet
                If sel.Count = 0 Then
                    MsgBox("Chưa chọn")
                    Exit Sub
                End If

                ' Get new name from user
                Dim prefix As String = Microsoft.VisualBasic.Interaction.InputBox("Đổi Tên solid Body", "Thành Design", "Prefix")
                If String.IsNullOrWhiteSpace(prefix) Then
                    MsgBox("No name entered")
                    Exit Sub
                End If

                ' Rename selected items (safe casts and per-item error handling)
                For Each item As Object In sel
                    Try
                        Dim sBody As SurfaceBody = TryCast(item, SurfaceBody)
                        If sBody IsNot Nothing Then
                            sBody.Name = prefix
                            Continue For
                        End If

                        Dim occ As ComponentOccurrence = TryCast(item, ComponentOccurrence)
                        If occ IsNot Nothing Then
                            occ.Name = prefix
                            Continue For
                        End If

                        ' Fallback: try to set Name property via reflection for other body types
                        Try
                            Dim nameProp = item.GetType().GetProperty("Name")
                            If nameProp IsNot Nothing AndAlso nameProp.CanWrite Then
                                nameProp.SetValue(item, prefix, Nothing)
                                Continue For
                            End If
                        Catch
                            ' ignore
                        End Try
                    Catch
                        ' ignore per-item errors
                    End Try
                Next

                ' Try adding a user defined property (best-effort)
                Try
                    Dim partDoc As PartDocument = TryCast(doc, PartDocument)
                    If partDoc IsNot Nothing Then
                        Dim userSet As Inventor.PropertySet = Nothing
                        For Each ps As Inventor.PropertySet In partDoc.PropertySets
                            If ps.Name.IndexOf("User Defined", StringComparison.OrdinalIgnoreCase) >= 0 OrElse ps.Name.IndexOf("Inventor User", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                userSet = ps
                                Exit For
                            End If
                        Next
                        If userSet IsNot Nothing Then
                            userSet.Add(prefix, "Prefix")
                        End If
                    End If
                Catch
                    ' ignore property addition errors
                End Try

                g_inventorApplication.StatusBarText = "Done"
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Part Action 4: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
