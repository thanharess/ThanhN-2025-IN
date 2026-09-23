
Option Explicit On

Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor
Imports IO = System.IO

Namespace ToolInventor2025.Assembly.Buttons.Part

    Public Module Ass_Part_8

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application

            Try
                ' Lấy Inventor đang chạy
                invApp = Interop.Marshal2.GetActiveObject("Inventor.Application")
            Catch ex As Exception
                ' Nếu chưa có thì mở Inventor mới
                invApp = Activator.CreateInstance(Type.GetTypeFromProgID("Inventor.Application"))
                invApp.Visible = True
            End Try

            Dim doc As Document = invApp.ActiveDocument

            ' Vòng lặp chọn chi tiết
            Do
                Dim entity As ComponentOccurrence
                entity = invApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter,
                                                "Chọn chi tiết trong cụm lắp 🙂 - hoặc ESC để thoát")

                If entity Is Nothing Then
                    Exit Do
                End If

                ' Lấy thông tin iProperties
                Dim oOne As String = "Tên: " & entity.Name
                Dim oTwo As String = "Khối lượng: " & Math.Round(entity.MassProperties.Mass, 2) & " kg"
                Dim oThree As String = "Vật liệu: " & entity.Definition.Material.Name

                Dim propSet As PropertySet = entity.Definition.Document.PropertySets.Item("Design Tracking Properties")
                Dim oFour As String = "Designer: " & propSet.Item("Designer").Value
                Dim oFive As String = "Description: " & propSet.Item("Description").Value

                If String.IsNullOrEmpty(oFive.Replace("Description:", "").Trim()) Then
                    oFive = "Description: " & entity.Name
                End If

                ' Highlight đối tượng
                Dim oSet1 As HighlightSet = doc.CreateHighlightSet()
                oSet1.Color = invApp.TransientObjects.CreateColor(255, 125, 0)
                oSet1.AddItem(entity)

                ' Hiển thị thông tin
                MsgBox(oOne & vbCrLf & oTwo & vbCrLf & oThree & vbCrLf & oFour & vbCrLf & oFive,
                   MsgBoxStyle.Information, "Thông tin chi tiết")

                oSet1.Clear()
            Loop
        End Sub
    End Module
End Namespace