Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep.constraint
    Public Module Ass_LG_C_1c

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                AlignComponentToBase()
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Assembly2 - Align Origin")
            End Try
        End Sub

        Private Sub AlignComponentToBase()
            Dim asmDoc As AssemblyDocument = TryCast(g_inventorApplication.ActiveDocument, AssemblyDocument)
            If asmDoc Is Nothing Then
                MessageBox.Show("Active document is not an assembly.", "Assembly2 - Align Origin")
                Return
            End If

            Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
            Dim trans As Transaction = g_inventorApplication.TransactionManager.StartTransaction(asmDoc, "Constrain Components (API)")

            ' Chọn component gốc
            Dim baseComp As ComponentOccurrence = PickComponent("Pick Base Component.")
            If baseComp Is Nothing Then trans.Abort() : Return

            Dim basePlanes As List(Of WorkPlaneProxy) = GetComponentOriginPlaneProxies(baseComp)
            If basePlanes Is Nothing OrElse basePlanes.Count <> 3 Then
                MessageBox.Show("Failed to get 3 origin planes from base component.", "Assembly2 - Align Origin")
                trans.Abort()
                Return
            End If

            Dim baseTrans As Matrix = baseComp.Transformation

            ' Chọn component cần di chuyển
            Dim compToMove As ComponentOccurrence = PickComponent("Pick Component To Move.")
            If compToMove Is Nothing Then trans.Abort() : Return

            Dim movePlanes As List(Of WorkPlaneProxy) = GetComponentOriginPlaneProxies(compToMove)
            If movePlanes Is Nothing OrElse movePlanes.Count <> 3 Then
                MessageBox.Show("Failed to get 3 origin planes from component to move.", "Assembly2 - Align Origin")
                trans.Abort()
                Return
            End If

            ' Xóa constraint cũ
            DeleteConstraints(compToMove)

            ' Đặt transformation bằng với base
            compToMove.Transformation = baseTrans

            ' Tạo Flush constraint cho 3 mặt phẳng gốc
            For i As Integer = 0 To 2
                compDef.Constraints.AddFlushConstraint(basePlanes(i), movePlanes(i), 0)
            Next

            trans.End()
            MessageBox.Show("Component aligned to base origin.", "Assembly2 - Align Origin")
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
            For Each c As AssemblyConstraint In comp.Constraints
                c.Delete()
            Next
        End Sub






    End Module
End Namespace
