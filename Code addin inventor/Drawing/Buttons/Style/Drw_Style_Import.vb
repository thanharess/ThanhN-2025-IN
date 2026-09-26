Option Explicit On
Option Strict Off

Imports Inventor
Imports System
Imports Env = System.Environment
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Reflection

Namespace ToolInventor2025.Drawing.Buttons.Style

    Public Module Drw_Style_Import

        Private Const STYLE_FOLDER_NAME As String = "Drawing style"

        Private ReadOnly StyleFolderMap As Dictionary(Of String, String) = BuildStyleFolderMap()

        Private Function BuildStyleFolderMap() As Dictionary(Of String, String)
            Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            map("Balloon.styxml") = "Balloon"
            map("Center mark.styxml") = "Center Mark"
            map("Datum.styxml") = "Datum"
            map("dim.styxml") = "Dimension"
            map("Feature control.styxml") = "Feature Control Frame"
            map("Hatch.styxml") = "Hatch"
            map("Hole.styxml") = "Hole"
            map("ID.styxml") = "ID"
            map("Layer.styxml") = "Layers"
            map("Leader.styxml") = "Leader"
            map("M cắt.styxml") = "Section"
            map("Partlist.styxml") = "Parts List"
            map("Suface.styxml") = "Surface Texture"
            map("Text.styxml") = "Text"
            Return map
        End Function

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                Dim oApp As Inventor.Application = g_inventorApplication
                If oApp Is Nothing Then Return

                Dim oDoc As Document = oApp.ActiveDocument
                If oDoc Is Nothing Then Return

                Dim oDrawDoc As DrawingDocument = TryCast(oDoc, DrawingDocument)
                If oDrawDoc Is Nothing Then
                    MessageBox.Show("Vui lòng mở một bản vẽ (Drawing) để import Style!", "Thông báo")
                    Return
                End If

                ImportDrawingStyles(oApp, oDrawDoc)
            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message)
            End Try
        End Sub

        Private Sub ImportDrawingStyles(ByVal oApp As Inventor.Application, ByVal oDrawDoc As DrawingDocument)
            Try
                Dim styleSourceDir As String = GetStyleSourceDirectory()
                If String.IsNullOrEmpty(styleSourceDir) OrElse Not System.IO.Directory.Exists(styleSourceDir) Then
                    MessageBox.Show("Không tìm thấy thư mục '" & STYLE_FOLDER_NAME & "' cạnh file Add-in.", "Lỗi")
                    Return
                End If

                Dim designDataPath As String = oApp.DesignDataPath
                If String.IsNullOrEmpty(designDataPath) OrElse Not System.IO.Directory.Exists(designDataPath) Then
                    MessageBox.Show("Không tìm thấy đường dẫn Design Data của Inventor.", "Lỗi")
                    Return
                End If

                Dim styxmlFiles As String() = System.IO.Directory.GetFiles(styleSourceDir, "*.styxml")
                Dim copiedCount As Integer = 0

                For Each filePath As String In styxmlFiles
                    Dim fileName As String = System.IO.Path.GetFileName(filePath)
                    Dim destSubFolder As String = ""

                    If StyleFolderMap.ContainsKey(fileName) Then
                        destSubFolder = StyleFolderMap(fileName)
                    ElseIf fileName.StartsWith("Style ", StringComparison.OrdinalIgnoreCase) Then
                        destSubFolder = "Dimension"
                    Else
                        Continue For
                    End If

                    Dim destDir As String = System.IO.Path.Combine(designDataPath, destSubFolder)
                    If Not System.IO.Directory.Exists(destDir) Then
                        System.IO.Directory.CreateDirectory(destDir)
                    End If

                    Dim destFilePath As String = System.IO.Path.Combine(destDir, fileName)

                    Try
                        System.IO.File.Copy(filePath, destFilePath, True)
                        copiedCount += 1
                    Catch ex As Exception
                        Debug.WriteLine("Lỗi copy " & fileName & ": " & ex.Message)
                    End Try
                Next

                If copiedCount = 0 Then
                    MessageBox.Show("Không có file .styxml nào được sao chép thành công.", "Thông báo")
                    Return
                End If

                Try
                    Dim stylesManager As DrawingStylesManager = oDrawDoc.StylesManager
                    Dim slm As Object = oApp.StyleLibraryManager

                    If slm Is Nothing Then
                        MessageBox.Show("Không lấy được StyleLibraryManager.", "Lỗi")
                        Return
                    End If

                    Dim activeLibrary As Object = Nothing
                    Try
                        activeLibrary = slm.ActiveStyleLibrary
                    Catch
                    End Try

                    If activeLibrary Is Nothing Then
                        MessageBox.Show("Không lấy được ActiveStyleLibrary. Kiểm tra Project đã cấu hình Style Library chưa.", "Lỗi")
                        Return
                    End If

                    Dim smObj As Object = stylesManager
                    smObj.UpdateFromStyleLibrary(activeLibrary)

                    MessageBox.Show("Đã import thành công " & copiedCount & " file Style vào bản vẽ!", "Hoàn tất")
                Catch ex As Exception
                    MessageBox.Show("Đã copy file nhưng lỗi khi cập nhật vào Drawing: " & ex.Message & vbCrLf & vbCrLf &
                                    "Vui lòng kiểm tra dự án Inventor đã set 'Use Style Library = Read-Write' chưa.", "Lỗi Cập nhật")
                End Try

            Catch ex As Exception
                MessageBox.Show("Lỗi hệ thống: " & ex.Message)
            End Try
        End Sub

        Private Function GetStyleSourceDirectory() As String
            Try
                Dim loc As String = Global.System.Reflection.Assembly.GetExecutingAssembly().Location
                If Not String.IsNullOrEmpty(loc) Then
                    Dim baseDir As String = System.IO.Path.GetDirectoryName(loc)
                    Dim targetDir As String = System.IO.Path.Combine(baseDir, STYLE_FOLDER_NAME)
                    If System.IO.Directory.Exists(targetDir) Then Return targetDir
                End If
            Catch
            End Try

            Dim roots As New List(Of String)

            Try
                Dim loc As String = Global.System.Reflection.Assembly.GetExecutingAssembly().Location
                If Not String.IsNullOrEmpty(loc) Then roots.Add(System.IO.Path.GetDirectoryName(loc))
            Catch
            End Try

            Try
                roots.Add(AppDomain.CurrentDomain.BaseDirectory.TrimEnd("\"c))
            Catch
            End Try

            Try
                For Each d As String In GetAddinFileDirs()
                    roots.Add(d)
                Next
            Catch
            End Try

            Dim uniq As New List(Of String)
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each d As String In roots
                Try
                    Dim full As String = System.IO.Path.GetFullPath(d)
                    If seen.Add(full) Then uniq.Add(full)
                Catch
                End Try
            Next

            For Each root As String In uniq
                Dim found As String = FindFolderRecursive(root, STYLE_FOLDER_NAME, 6)
                If Not String.IsNullOrEmpty(found) Then Return found
            Next

            Return ""
        End Function

        Private Function FindFolderRecursive(ByVal dir As String, ByVal folderName As String, ByVal maxDepth As Integer) As String
            If maxDepth < 0 OrElse Not System.IO.Directory.Exists(dir) Then Return Nothing

            Try
                Dim target As String = System.IO.Path.Combine(dir, folderName)
                If System.IO.Directory.Exists(target) Then Return target

                If maxDepth > 0 Then
                    For Each subDir As String In System.IO.Directory.GetDirectories(dir)
                        Dim found As String = FindFolderRecursive(subDir, folderName, maxDepth - 1)
                        If Not String.IsNullOrEmpty(found) Then Return found
                    Next
                End If
            Catch
            End Try

            Return Nothing
        End Function

        Private Function GetAddinFileDirs() As List(Of String)
            Dim result As New List(Of String)
            Try
                Dim roots As String() = {
                    System.IO.Path.Combine(Env.GetFolderPath(Env.SpecialFolder.ApplicationData), "Autodesk", "ApplicationPlugins"),
                    System.IO.Path.Combine(Env.GetFolderPath(Env.SpecialFolder.CommonApplicationData), "Autodesk", "ApplicationPlugins")
                }

                For Each root As String In roots
                    If Not System.IO.Directory.Exists(root) Then Continue For
                    For Each af As String In System.IO.Directory.GetFiles(root, "*.addin", System.IO.SearchOption.AllDirectories)
                        result.Add(System.IO.Path.GetDirectoryName(af))
                    Next
                Next
            Catch
            End Try
            Return result
        End Function

    End Module
End Namespace