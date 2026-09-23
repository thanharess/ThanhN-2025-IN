Imports Inventor

Namespace ToolInventor2025
    Public Class ToolButtonDrawing
        Private Shared Function LoadIconFromPath(path As String) As stdole.IPictureDisp
            Try
                If String.IsNullOrEmpty(path) Then Return Nothing
                If Not System.IO.File.Exists(path) Then Return Nothing
                Using bmp As New System.Drawing.Bitmap(path)
                    Dim clone As New System.Drawing.Bitmap(bmp)
                    Try
                        Return PictureDispConverter.ToIPictureDisp(clone)
                    Finally
                        clone.Dispose()
                    End Try
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        Public Shared Sub Register(controlDefs As Inventor.ControlDefinitions, addInClientID As String, buttonsList As System.Collections.Generic.List(Of ButtonDefinition), largeIcon As stdole.IPictureDisp, smallIcon As stdole.IPictureDisp)

            Dim assemblyFolder4 As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim configured As String = Nothing
            Try
                configured = My.Settings.ImageFolder
            Catch
                configured = Nothing
            End Try

            Dim iconsFolder As String = Nothing
            If Not String.IsNullOrWhiteSpace(configured) AndAlso System.IO.Directory.Exists(configured) Then
                iconsFolder = System.IO.Path.Combine(configured, "Drawing")
            Else
                iconsFolder = System.IO.Path.Combine(assemblyFolder4, "Code addin inventor", "Images", "Drawing")
            End If

#Region "Prepare icon paths"
            Dim ToolDrawLargePath1 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath1 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath2 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath2 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath3 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath3 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath4 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath4 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath5 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath5 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath6 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath6 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath7 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath7 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath8 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath8 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath9 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath9 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath10 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath10 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath11 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath11 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath12 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath12 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath13 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath13 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath14 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath14 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath15 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath15 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath16 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath16 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath17 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath17 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath18 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath18 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath19 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath19 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath20 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath20 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath21 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath21 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath22 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath22 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath23 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath23 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath24 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath24 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath25 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath25 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath26 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath26 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath27 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath27 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath28 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")

            ' Prepare IPictureDisp icons (use provided largeIcon/smallIcon as fallback)
            Dim ToolDrawLargeIcon1 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath1), LoadIconFromPath(ToolDrawLargePath1), largeIcon)
            Dim ToolDrawSmallIcon1 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath1), LoadIconFromPath(ToolDrawSmallPath1), smallIcon)
            Dim ToolDrawLargeIcon2 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath2), LoadIconFromPath(ToolDrawLargePath2), largeIcon)
            Dim ToolDrawSmallIcon2 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath2), LoadIconFromPath(ToolDrawSmallPath2), smallIcon)
            Dim ToolDrawLargeIcon3 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath3), LoadIconFromPath(ToolDrawLargePath3), largeIcon)
            Dim ToolDrawSmallIcon3 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath3), LoadIconFromPath(ToolDrawSmallPath3), smallIcon)
            Dim ToolDrawLargeIcon4 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath4), LoadIconFromPath(ToolDrawLargePath4), largeIcon)
            Dim ToolDrawSmallIcon4 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath4), LoadIconFromPath(ToolDrawSmallPath4), smallIcon)
            Dim ToolDrawLargeIcon5 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath5), LoadIconFromPath(ToolDrawLargePath5), largeIcon)
            Dim ToolDrawSmallIcon5 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath5), LoadIconFromPath(ToolDrawSmallPath5), smallIcon)
            Dim ToolDrawLargeIcon6 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath6), LoadIconFromPath(ToolDrawLargePath6), largeIcon)
            Dim ToolDrawSmallIcon6 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath6), LoadIconFromPath(ToolDrawSmallPath6), smallIcon)
            Dim ToolDrawLargeIcon7 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath7), LoadIconFromPath(ToolDrawLargePath7), largeIcon)
            Dim ToolDrawSmallIcon7 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath7), LoadIconFromPath(ToolDrawSmallPath7), smallIcon)
            Dim ToolDrawLargeIcon8 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath8), LoadIconFromPath(ToolDrawLargePath8), largeIcon)
            Dim ToolDrawSmallIcon8 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath8), LoadIconFromPath(ToolDrawSmallPath8), smallIcon)
            Dim ToolDrawLargeIcon9 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath9), LoadIconFromPath(ToolDrawLargePath9), largeIcon)
            Dim ToolDrawSmallIcon9 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath9), LoadIconFromPath(ToolDrawSmallPath9), smallIcon)
            Dim ToolDrawLargeIcon10 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath10), LoadIconFromPath(ToolDrawLargePath10), largeIcon)
            Dim ToolDrawSmallIcon10 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath10), LoadIconFromPath(ToolDrawSmallPath10), smallIcon)
            Dim ToolDrawLargeIcon11 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath11), LoadIconFromPath(ToolDrawLargePath11), largeIcon)
            Dim ToolDrawSmallIcon11 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath11), LoadIconFromPath(ToolDrawSmallPath11), smallIcon)
            Dim ToolDrawLargeIcon12 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath12), LoadIconFromPath(ToolDrawLargePath12), largeIcon)
            Dim ToolDrawSmallIcon12 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath12), LoadIconFromPath(ToolDrawSmallPath12), smallIcon)
            Dim ToolDrawLargeIcon13 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath13), LoadIconFromPath(ToolDrawLargePath13), largeIcon)
            Dim ToolDrawSmallIcon13 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath13), LoadIconFromPath(ToolDrawSmallPath13), smallIcon)
            Dim ToolDrawLargeIcon14 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath14), LoadIconFromPath(ToolDrawLargePath14), largeIcon)
            Dim ToolDrawSmallIcon14 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath14), LoadIconFromPath(ToolDrawSmallPath14), smallIcon)
            Dim ToolDrawLargeIcon15 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath15), LoadIconFromPath(ToolDrawLargePath15), largeIcon)
            Dim ToolDrawSmallIcon15 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath15), LoadIconFromPath(ToolDrawSmallPath15), smallIcon)
            Dim ToolDrawLargeIcon16 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath16), LoadIconFromPath(ToolDrawLargePath16), largeIcon)
            Dim ToolDrawSmallIcon16 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath16), LoadIconFromPath(ToolDrawSmallPath16), smallIcon)
            Dim ToolDrawLargeIcon17 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath17), LoadIconFromPath(ToolDrawLargePath17), largeIcon)
            Dim ToolDrawSmallIcon17 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath17), LoadIconFromPath(ToolDrawSmallPath17), smallIcon)
            Dim ToolDrawLargeIcon18 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath18), LoadIconFromPath(ToolDrawLargePath18), largeIcon)
            Dim ToolDrawSmallIcon18 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath18), LoadIconFromPath(ToolDrawSmallPath18), smallIcon)
            Dim ToolDrawLargeIcon19 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath19), LoadIconFromPath(ToolDrawLargePath19), largeIcon)
            Dim ToolDrawSmallIcon19 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath19), LoadIconFromPath(ToolDrawSmallPath19), smallIcon)
            Dim ToolDrawLargeIcon20 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath20), LoadIconFromPath(ToolDrawLargePath20), largeIcon)
            Dim ToolDrawSmallIcon20 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath20), LoadIconFromPath(ToolDrawSmallPath20), smallIcon)
            Dim ToolDrawLargeIcon21 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath21), LoadIconFromPath(ToolDrawLargePath21), largeIcon)
            Dim ToolDrawSmallIcon21 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath21), LoadIconFromPath(ToolDrawSmallPath21), smallIcon)
            Dim ToolDrawLargeIcon22 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath22), LoadIconFromPath(ToolDrawLargePath22), largeIcon)
            Dim ToolDrawSmallIcon22 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath22), LoadIconFromPath(ToolDrawSmallPath22), smallIcon)
            Dim ToolDrawLargeIcon23 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath23), LoadIconFromPath(ToolDrawLargePath23), largeIcon)
            Dim ToolDrawSmallIcon23 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath23), LoadIconFromPath(ToolDrawSmallPath23), smallIcon)
#End Region





        End Sub
    End Class
End Namespace
