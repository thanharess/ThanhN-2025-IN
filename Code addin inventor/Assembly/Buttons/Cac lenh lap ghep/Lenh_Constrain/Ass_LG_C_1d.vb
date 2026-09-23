Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep.constraint
    Public Module Ass_LG_C_1d
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                ConstrainAllComponents()
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Assembly2 - Align Origin")
            End Try
        End Sub


        Private Sub ConstrainAllComponents()
            Dim asmDoc As AssemblyDocument = TryCast(g_inventorApplication.ActiveDocument, AssemblyDocument)
            If asmDoc Is Nothing Then
                MessageBox.Show("Active document is not an assembly.", "Assembly2 - Constrain Planes & Axes")
                Return
            End If

            Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
            Dim trans As Transaction = g_inventorApplication.TransactionManager.StartTransaction(asmDoc, "Constrain Components (API)")

            ' Chọn component gốc
            Dim baseComp As ComponentOccurrence = PickComponent("Pick Base Component.")
            If baseComp Is Nothing Then trans.Abort() : Return
            baseComp.Grounded = True

            Dim basePlanes As List(Of WorkPlaneProxy) = GetComponentOriginPlaneProxies(baseComp)
            If basePlanes Is Nothing OrElse basePlanes.Count <> 3 Then
                MessageBox.Show("Failed to get 3 WorkPlaneProxy objects from Base component.", "Assembly2 - Constrain Planes & Axes")
                trans.Abort()
                Return
            End If

            Dim baseAxes As List(Of WorkAxisProxy) = GetComponentOriginAxisProxies(baseComp)
            If baseAxes Is Nothing OrElse baseAxes.Count <> 3 Then
                MessageBox.Show("Failed to get 3 WorkAxisProxy objects from Base component.", "Assembly2 - Constrain Planes & Axes")
                trans.Abort()
                Return
            End If

            ' Lặp qua các component khác
            For Each otherComp As ComponentOccurrence In compDef.Occurrences
                If otherComp Is baseComp Then Continue For
                If TypeOf otherComp.Definition Is VirtualComponentDefinition Then Continue For
                If Not ComponentPatternElementChecker(otherComp) Then Continue For

                otherComp.Grounded = True

                Dim otherPlanes As List(Of WorkPlaneProxy) = GetComponentOriginPlaneProxies(otherComp)
                If otherPlanes Is Nothing OrElse otherPlanes.Count <> 3 Then Continue For

                Dim otherAxes As List(Of WorkAxisProxy) = GetComponentOriginAxisProxies(otherComp)
                If otherAxes Is Nothing OrElse otherAxes.Count <> 3 Then Continue For

                ' Xóa constraint cũ
                DeleteConstraints(otherComp)

                ' Tạo constraint mới
                ConstrainWPlanes(basePlanes, otherPlanes)
                ConstrainWPAxes(baseAxes, otherAxes)

                otherComp.Grounded = False
            Next

            baseComp.Grounded = False
            asmDoc.Update2(True)
            trans.End()
            MessageBox.Show("Auto-constrain completed for all components.", "Assembly2 - Constrain Planes & Axes")
        End Sub

        ' Các hàm phụ trợ
        Private Function PickComponent(prompt As String) As ComponentOccurrence
            Dim obj = g_inventorApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, prompt)
            Return TryCast(obj, ComponentOccurrence)
        End Function

        Private Function GetComponentOriginPlaneProxies(
    ByVal comp As ComponentOccurrence) As List(Of WorkPlaneProxy)

            If comp Is Nothing Then Return Nothing

            Dim wps As WorkPlanes = comp.Definition.WorkPlanes
            Dim proxies As New List(Of WorkPlaneProxy)

            For i As Integer = 1 To 3

                Dim wp As WorkPlane = wps.Item(i)

                ' CreateGeometryProxy yêu cầu ByRef Object
                Dim proxyObj As Object = Nothing

                comp.CreateGeometryProxy(wp, proxyObj)

                Dim proxy As WorkPlaneProxy =
            TryCast(proxyObj, WorkPlaneProxy)

                If proxy IsNot Nothing Then
                    proxies.Add(proxy)
                End If

            Next

            Return proxies

        End Function


        Private Function GetComponentOriginAxisProxies(
    ByVal comp As ComponentOccurrence) As List(Of WorkAxisProxy)

            If comp Is Nothing Then Return Nothing

            Dim was As WorkAxes = comp.Definition.WorkAxes
            Dim proxies As New List(Of WorkAxisProxy)

            For i As Integer = 1 To 3

                Dim wa As WorkAxis = was.Item(i)

                ' CreateGeometryProxy yêu cầu ByRef Object
                Dim proxyObj As Object = Nothing

                comp.CreateGeometryProxy(wa, proxyObj)

                Dim proxy As WorkAxisProxy =
            TryCast(proxyObj, WorkAxisProxy)

                If proxy IsNot Nothing Then
                    proxies.Add(proxy)
                End If

            Next

            Return proxies

        End Function
        Private Sub DeleteConstraints(comp As ComponentOccurrence)
            If comp Is Nothing Then Exit Sub
            For Each c As AssemblyConstraint In comp.Constraints
                c.Delete()
            Next
        End Sub

        Private Function ComponentPatternElementChecker(comp As ComponentOccurrence) As Boolean
            If comp Is Nothing Then Return False
            If Not comp.IsPatternElement Then Return True
            Dim elem As OccurrencePatternElement = comp.PatternElement
            If elem.Suppressed Then Return False
            If elem.Independent Then Return True
            If TypeOf elem.Parent Is OccurrencePattern Then
                Dim patt As OccurrencePattern = elem.Parent
                If patt.OccurrencePatternElements.Item(1) Is elem Then
                    If Not patt.IsPatternElement Then Return True Else Return False
                End If
            End If
            Return False
        End Function

        Private Sub ConstrainWPlanes(basePlanes As List(Of WorkPlaneProxy), otherPlanes As List(Of WorkPlaneProxy))
            Dim oConsts As AssemblyConstraints = basePlanes(0).ContainingOccurrence.Parent.Constraints
            For Each bp In basePlanes
                For Each op In otherPlanes
                    If bp.Plane.IsPerpendicularTo(op.Plane, 0.00001) Then Continue For
                    If bp.Plane.IsParallelTo(op.Plane, 0.00001) Then
                        Dim offset = bp.Plane.DistanceTo(op.Plane.RootPoint)
                        If bp.Plane.Normal.IsEqualTo(op.Plane.Normal, 0.00001) Then
                            oConsts.AddFlushConstraint(bp, op, offset)
                        Else
                            oConsts.AddMateConstraint(bp, op, offset)
                        End If
                    Else
                        Dim angle = bp.Plane.Normal.AngleTo(op.Plane.Normal)
                        oConsts.AddAngleConstraint(bp, op, angle, AngleConstraintSolutionTypeEnum.kDirectedSolution)
                    End If
                Next
            Next
        End Sub

        Private Sub ConstrainWPAxes(baseAxes As List(Of WorkAxisProxy), otherAxes As List(Of WorkAxisProxy))
            Dim oConsts As AssemblyConstraints = baseAxes(0).ContainingOccurrence.Parent.Constraints
            For Each ba In baseAxes
                For Each oa In otherAxes
                    If ba.Line.Direction.IsPerpendicularTo(oa.Line.Direction, 0.00001) Then Continue For
                    If ba.Line.Direction.IsParallelTo(oa.Line.Direction, 0.00001) Then
                        Dim infType = InferredTypeEnum.kInferredLine
                        Dim aligned = MateConstraintSolutionTypeEnum.kAlignedSolutionType
                        Dim opposed = MateConstraintSolutionTypeEnum.kOpposedSolutionType
                        If ba.Line.IsColinearTo(oa.Line, 0.00001) Then
                            If ba.Line.Direction.IsEqualTo(oa.Line.Direction, 0.00001) Then
                                oConsts.AddMateConstraint2(ba, oa, 0, infType, infType, aligned)
                            Else
                                oConsts.AddMateConstraint2(ba, oa, 0, infType, infType, opposed)
                            End If
                        Else
                            Dim offset = ba.Line.DistanceTo(oa.Line.RootPoint)
                            If ba.Line.Direction.IsEqualTo(oa.Line.Direction, 0.00001) Then
                                oConsts.AddMateConstraint2(ba, oa, offset, infType, infType, aligned)
                            Else
                                oConsts.AddMateConstraint2(ba, oa, offset, infType, infType, opposed)
                            End If
                        End If
                    Else
                        Dim angle = ba.Line.Direction.AngleTo(oa.Line.Direction)
                        oConsts.AddAngleConstraint(ba, oa, angle, AngleConstraintSolutionTypeEnum.kDirectedSolution)
                    End If
                Next
            Next
        End Sub


    End Module
End Namespace
