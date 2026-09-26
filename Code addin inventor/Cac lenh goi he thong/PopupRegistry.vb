Imports Inventor
Imports System.Collections.Generic

Namespace ToolInventor2025

    ''' <summary>
    ''' Registry dùng chung cho mọi popup — cả Assembly, Part, Drawing
    ''' </summary>
    Public Class PopupRegistry

        Public Class PopupDef
            Public SubButtons As New List(Of ButtonDefinition)  ' [0] = nút chính
        End Class

        ''' <summary>
        ''' Key = panelInternalName (VD: "ToolInventor2025_AssemblyPanel")
        ''' Value = danh sách popup cần tạo trên panel đó
        ''' </summary>
        Public Shared ReadOnly PendingPopups As New Dictionary(Of String, List(Of PopupDef))

        Public Shared Sub RegisterPopup(panelInternalName As String, pd As PopupDef)
            If pd Is Nothing OrElse pd.SubButtons Is Nothing OrElse pd.SubButtons.Count = 0 Then Return
            If Not PendingPopups.ContainsKey(panelInternalName) Then
                PendingPopups(panelInternalName) = New List(Of PopupDef)
            End If
            PendingPopups(panelInternalName).Add(pd)
        End Sub

    End Class

End Namespace