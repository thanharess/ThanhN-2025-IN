Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep.constraint



    Public Module Ass_LG_C_1b
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                AutoConstrainComponents()
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Assembly2 - Auto Constrain")
            End Try
        End Sub

        Private Sub AutoConstrainComponents()
            Dim asmDoc As AssemblyDocument = TryCast(g_inventorApplication.ActiveDocument, AssemblyDocument)
            If asmDoc Is Nothing Then
                MessageBox.Show("Active document is not an assembly.", "Assembly2 - Auto Constrain")
                Return
            End If

            Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
            Dim trans As Transaction = g_inventorApplication.TransactionManager.StartTransaction(asmDoc, "Constrain Components (API)")

            ' Chọn component gốc
            Dim baseComp As ComponentOccurrence = PickComponent("Pick Base Component.")
            If baseComp Is Nothing Then trans.Abort() : Return

            Dim basePlanes As List(Of WorkPlaneProxy) = GetComponentOriginPlaneProxies(baseComp)
            If basePlanes Is Nothing OrElse basePlanes.Count <> 3 Then
                MessageBox.Show("Failed to get 3 origin planes from base component.", "Assembly2 - Auto Constrain")
                trans.Abort()
                Return
            End If

            ' Loop qua các component khác
            For Each otherComp As ComponentOccurrence In compDef.Occurrences
                If otherComp Is baseComp Then Continue For
                If TypeOf otherComp.Definition Is VirtualComponentDefinition Then Continue For
                If Not ComponentPatternElementChecker(otherComp) Then Continue For

                Dim otherPlanes As List(Of WorkPlaneProxy) = GetComponentOriginPlaneProxies(otherComp)
                If otherPlanes Is Nothing OrElse otherPlanes.Count <> 3 Then Continue For

                ' Xóa constraint cũ
                DeleteConstraints(otherComp)

                ' Tạo constraint mới
                For Each basePlane In basePlanes
                    For Each otherPlane In otherPlanes
                        If Not basePlane.Plane.IsParallelTo(otherPlane.Plane, 0.00001) Then Continue For
                        Dim offset As Double = basePlane.Plane.DistanceTo(otherPlane.Plane.RootPoint)
                        If basePlane.Plane.Normal.IsEqualTo(otherPlane.Plane.Normal, 0.00001) Then
                            compDef.Constraints.AddFlushConstraint(basePlane, otherPlane, offset)
                        Else
                            compDef.Constraints.AddMateConstraint(basePlane, otherPlane, offset)
                        End If
                    Next
                Next
            Next

            trans.End()
            MessageBox.Show("Auto-constrain completed.", "Assembly2 - Auto Constrain")
        End Sub

        ' Các hàm phụ trợ giống iLogic
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

                ' CreateGeometryProxy dùng ByRef Object
                Dim proxyObj As Object = Nothing

                comp.CreateGeometryProxy(wp, proxyObj)

                ' Chuyển Object → WorkPlaneProxy
                Dim proxy As WorkPlaneProxy =
            TryCast(proxyObj, WorkPlaneProxy)

                If proxy IsNot Nothing Then
                    proxies.Add(proxy)
                End If

            Next

            Return proxies

        End Function

        Private Sub DeleteConstraints(comp As ComponentOccurrence)
            If comp Is Nothing Then Exit Sub
            If comp.Grounded Then comp.Grounded = False
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
    End Module
End Namespace
