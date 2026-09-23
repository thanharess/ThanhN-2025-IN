Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Part.Buttons.solid
    Public Module Part_Solid_2
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                ' iLogic-style Main implementation: delete fixed constraints from selected sketch
                Dim oSketchEnt As PlanarSketch
                Dim oSelSet As SelectSet
                Dim j As Integer
                oSelSet = g_inventorApplication.ActiveDocument.SelectSet

                If oSelSet.Count = 0 Then
                    MsgBox("Please select a sketch from the browser.")
                    Exit Sub
                End If

                oSketchEnt = TryCast(oSelSet.Item(1), PlanarSketch)
                If oSketchEnt Is Nothing Then
                    MsgBox("Selected item is not a sketch. Please select a sketch from the browser.")
                    Exit Sub
                End If

                Debug.Print("  geometrical Constraint count: " & oSketchEnt.GeometricConstraints.Count)
                For j = oSketchEnt.GeometricConstraints.Count To 1 Step -1
                    Dim gc As GeometricConstraint = oSketchEnt.GeometricConstraints.Item(j)
                    Debug.Print("  Constraint type: " & gc.Type)
                    ' Try deleting if constraint is a ground/fixed constraint.
                    ' Some Inventor versions use different numeric values; check both common candidates.
                    If gc.Type = 83901952 Or gc.Type = 83 Then
                        gc.Delete()
                    End If
                Next j

                MsgBox("All fixed constraints in " & oSketchEnt.Name & " are deleted")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Part Action 2: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
