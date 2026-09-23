Imports Inventor

Namespace ToolInventor2025.Part.Buttons.solid
    Public Module Part_Solid_3
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                ' Prompt user to pick a sketch line
                Dim prompt As String = "Select line"
                Dim picked As Object = g_inventorApplication.CommandManager.Pick(SelectionFilterEnum.kSketchCurveLinearFilter, prompt)
                If picked Is Nothing Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("No line selected; operation cancelled.")
                    Exit Sub
                End If

                Dim oSketchLine As SketchLine = TryCast(picked, SketchLine)
                If oSketchLine Is Nothing Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Selected object is not a sketch line.")
                    Exit Sub
                End If

                ' Prompt for diameter, pitch and height
                Dim diamInput As String = Microsoft.VisualBasic.Interaction.InputBox("Enter diameter (model units)", "Diameter", "5")
                If String.IsNullOrWhiteSpace(diamInput) Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Helical curve creation cancelled (no diameter)")
                    Exit Sub
                End If
                Dim diameter As Double
                If Not Double.TryParse(diamInput, diameter) Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Invalid diameter: " & diamInput)
                    Exit Sub
                End If

                Dim pitchInput As String = Microsoft.VisualBasic.Interaction.InputBox("Enter pitch (model units)", "Pitch", "1")
                If String.IsNullOrWhiteSpace(pitchInput) Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Helical curve creation cancelled (no pitch)")
                    Exit Sub
                End If
                Dim pitch As Double
                If Not Double.TryParse(pitchInput, pitch) Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Invalid pitch: " & pitchInput)
                    Exit Sub
                End If

                Dim heightInput As String = Microsoft.VisualBasic.Interaction.InputBox("Enter height (model units)", "Height", "5")
                If String.IsNullOrWhiteSpace(heightInput) Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Helical curve creation cancelled (no height)")
                    Exit Sub
                End If
                Dim height As Double
                If Not Double.TryParse(heightInput, height) Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Invalid height: " & heightInput)
                    Exit Sub
                End If

                ' Use default revolution (Nothing)
                Dim revolution As Object = Nothing

                CreateHelicalCurve(oSketchLine, diameter, pitch, revolution, height)
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Helical curve created")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Part Action 3: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub

        Private Sub CreateHelicalCurve(oSketchLine As SketchLine, diameter As Double, pitch As Double, revolution As Object, height As Double)
            Try
                Dim partDef As PartComponentDefinition = TryCast(oSketchLine.Parent.Parent, PartComponentDefinition)
                If partDef Is Nothing Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Unable to determine PartComponentDefinition from selected sketch line.")
                    Exit Sub
                End If

                Dim sketch3D As Sketch3D = partDef.Sketches3D.Add()

                Dim axisStartPoint As Point = oSketchLine.StartSketchPoint.Geometry3d
                Dim axisEndPoint As Point = oSketchLine.EndSketchPoint.Geometry3d

                Dim curveStartPoint As Point = axisStartPoint.Copy()
                curveStartPoint.TranslateBy(g_inventorApplication.TransientGeometry.CreateVector(0, 0, 1))

                Dim helicalCurveDefinition As HelicalCurveConstantShapeDefinition = sketch3D.HelicalCurves.CreateConstantShapeDefinition(
                    HelicalShapeDefinitionTypeEnum.kPitchAndHeightShapeType,
                    axisStartPoint,
                    axisEndPoint,
                    curveStartPoint,
                    diameter,
                    pitch,
                    revolution,
                    height)

                sketch3D.HelicalCurves.Add(helicalCurveDefinition)
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error creating helical curve: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
