Imports System
Imports System.IO

Module UserSettings
    Public Function GetConfiguredImageFolder() As String
        Try
            Dim appdata = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim cfgDir = Path.Combine(appdata, "ToolInventor2025")
            Dim cfgFile = Path.Combine(cfgDir, "imagefolder.txt")
            If File.Exists(cfgFile) Then
                Dim text = File.ReadAllText(cfgFile).Trim()
                If Not String.IsNullOrWhiteSpace(text) Then
                    Return text
                End If
            End If
        Catch
            ' ignore
        End Try
        Return Nothing
    End Function

    Public Sub SetConfiguredImageFolder(folder As String)
        Try
            Dim appdata = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim cfgDir = Path.Combine(appdata, "ToolInventor2025")
            If Not Directory.Exists(cfgDir) Then Directory.CreateDirectory(cfgDir)
            Dim cfgFile = Path.Combine(cfgDir, "imagefolder.txt")
            File.WriteAllText(cfgFile, folder)
        Catch
            ' ignore
        End Try
    End Sub
End Module
