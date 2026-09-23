Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Part
    Public Module Ass_Part_6

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Executed Assembly Action 12")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Assembly Action 12: " & ex.Message)
                Catch
                End Try
            End Try

            Dim invApp As Inventor.Application =
                CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)

            Dim asmDoc As AssemblyDocument = TryCast(invApp.ActiveDocument, AssemblyDocument)

            If asmDoc Is Nothing Then
                MessageBox.Show("This is not an assembly!", "Visual Studio")
                Return
            End If

            Dim totalReset As Integer = 0

            ResetRecursive(asmDoc.ComponentDefinition, totalReset)

            Try : asmDoc.Update2(True) : Catch : End Try

            MessageBox.Show(
                "Đã reset màu về gốc." & vbCrLf &
                "Số occurrence đã reset: " & totalReset,
                "Hoàn tất",
                MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub


        '=====================================================
        ' ĐỆ QUY RESET APPEARANCE — KHÔNG CHECK DƯ
        '=====================================================
        Private Sub ResetRecursive(ByVal asmDef As AssemblyComponentDefinition,
                                   ByRef totalReset As Integer)

            If asmDef Is Nothing Then Return

            '--- 1. Lấy danh sách occurrence có override (copy ra mảng trước) ---
            Dim overrideList As ObjectCollection = Nothing
            Try
                overrideList = asmDef.AppearanceOverridesObjects
            Catch
            End Try

            If overrideList IsNot Nothing AndAlso overrideList.Count > 0 Then

                ' Copy ra List để tránh bị thay đổi khi đang duyệt
                Dim items As New List(Of Object)
                For i As Integer = 1 To overrideList.Count
                    Try
                        items.Add(overrideList.Item(i))
                    Catch
                    End Try
                Next

                For Each obj As Object In items
                    Try
                        Dim occ As ComponentOccurrence = TryCast(obj, ComponentOccurrence)
                        If occ Is Nothing Then Continue For

                        ' Reset thẳng — không cần kiểm tra gì thêm
                        occ.AppearanceSourceType = AppearanceSourceTypeEnum.kPartAppearance
                        totalReset += 1

                    Catch
                    End Try
                Next

            End If

            '--- 2. Đệ quy vào sub-assembly ---
            For Each occ As ComponentOccurrence In asmDef.Occurrences
                Try
                    If occ.Suppressed Then Continue For

                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then

                        Dim subAsmDef As AssemblyComponentDefinition = Nothing
                        Try
                            subAsmDef = TryCast(occ.Definition, AssemblyComponentDefinition)
                        Catch
                        End Try

                        If subAsmDef IsNot Nothing Then
                            ResetRecursive(subAsmDef, totalReset)
                        End If

                    End If
                Catch
                End Try
            Next

        End Sub

    End Module
End Namespace