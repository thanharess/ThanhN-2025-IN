Namespace ToolInventor2025.My
    Public Module Settings
        Private _imageFolder As String = Nothing

        Public Property ImageFolder As String
            Get
                If _imageFolder Is Nothing Then
                    Try
                        _imageFolder = UserSettings.GetConfiguredImageFolder()
                    Catch
                        _imageFolder = Nothing
                    End Try
                End If
                Return _imageFolder
            End Get
            Set(value As String)
                _imageFolder = value
                Try
                    If Not String.IsNullOrWhiteSpace(value) Then
                        UserSettings.SetConfiguredImageFolder(value)
                    End If
                Catch
                    ' ignore
                End Try
            End Set
        End Property
    End Module

End Namespace
