Imports Inventor

Namespace ToolInventor2025
    Public Class PartButtons
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
        Private Shared Function FindImagesFolder(startFolder As String) As String
            Dim folder As New System.IO.DirectoryInfo(startFolder)

            While folder IsNot Nothing
                Dim candidate As String =
            System.IO.Path.Combine(folder.FullName, "Images")

                If System.IO.Directory.Exists(candidate) Then
                    Return candidate
                End If

                folder = folder.Parent
            End While

            Return Nothing
        End Function

        Public Shared Sub Register(controlDefs As Inventor.ControlDefinitions, addInClientID As String, buttonsList As System.Collections.Generic.List(Of ButtonDefinition), largeIcon As stdole.IPictureDisp, smallIcon As stdole.IPictureDisp)
            ' Multi-line description/tooltips so you can edit them easily
            ' Per-button icon file paths (update these paths later to change icons per button)
            ' Resolve icons relative to the add-in assembly so installed location can vary
            ' Use absolute machine path for images instead of project-relative paths
            '  Dim iconsFolder As String = "C:\Users\thanh\source\repos\ThanhN\Code addin inventor\Images\part"

            '  Dim iconsFolder As String =
            ' "C:\Users\thanh\source\repos\ThanhN\Code addin inventor\Images\part"



            Dim assemblyFolder1 As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim configuredPart As String = Nothing
            Try
                configuredPart = My.Settings.ImageFolder
            Catch
                configuredPart = Nothing
            End Try

            Dim iconsFolder As String = Nothing
            If Not String.IsNullOrWhiteSpace(configuredPart) AndAlso System.IO.Directory.Exists(configuredPart) Then
                iconsFolder = System.IO.Path.Combine(configuredPart, "Part")
            Else
                iconsFolder = System.IO.Path.Combine(assemblyFolder1, "Code addin inventor", "Images", "Part")
            End If

#Region "Icon"
            ' Images in project: Code addin inventor\Images\Button\part (should be copied to output)
            ' Dim iconsFolder As String = System.IO.Path.Combine(assemblyFolder1, "Images", "part")

            ' Đường dẫn tới file ảnh

            Dim Part1LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part1SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part2LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part2SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part3LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part3SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part4LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part4SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part5LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part5SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part6LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part6SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part7LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part7SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part8LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part8SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part9LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part9SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part10LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part10SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part11LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part11SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part12LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part12SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part13LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part13SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part14LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part14SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")
            Dim Part15LargePath As String = System.IO.Path.Combine(iconsFolder, "i3.bmp")
            Dim Part15SmallPath As String = System.IO.Path.Combine(iconsFolder, "i3 1.bmp")


            ' Load per-button icons (fallback to provided largeIcon/smallIcon when file missing)
            Dim part1LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part1LargePath), LoadIconFromPath(Part1LargePath), largeIcon)
            Dim part1SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part1SmallPath), LoadIconFromPath(Part1SmallPath), smallIcon)
            Dim part2LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part2LargePath), LoadIconFromPath(Part2LargePath), largeIcon)
            Dim part2SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part2SmallPath), LoadIconFromPath(Part2SmallPath), smallIcon)
            Dim part3LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part3LargePath), LoadIconFromPath(Part3LargePath), largeIcon)
            Dim part3SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part3SmallPath), LoadIconFromPath(Part3SmallPath), smallIcon)
            Dim part4LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part4LargePath), LoadIconFromPath(Part4LargePath), largeIcon)
            Dim part4SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part4SmallPath), LoadIconFromPath(Part4SmallPath), smallIcon)
            Dim part5LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part5LargePath), LoadIconFromPath(Part5LargePath), largeIcon)
            Dim part5SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part5SmallPath), LoadIconFromPath(Part5SmallPath), smallIcon)
            Dim part6LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part6LargePath), LoadIconFromPath(Part6LargePath), largeIcon)
            Dim part6SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part6SmallPath), LoadIconFromPath(Part6SmallPath), smallIcon)
            Dim part7LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part7LargePath), LoadIconFromPath(Part7LargePath), largeIcon)
            Dim part7SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part7SmallPath), LoadIconFromPath(Part7SmallPath), smallIcon)
            Dim part8LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part8LargePath), LoadIconFromPath(Part8LargePath), largeIcon)
            Dim part8SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part8SmallPath), LoadIconFromPath(Part8SmallPath), smallIcon)
            Dim part9LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part9LargePath), LoadIconFromPath(Part9LargePath), largeIcon)
            Dim part9SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part9SmallPath), LoadIconFromPath(Part9SmallPath), smallIcon)
            Dim part10LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part10LargePath), LoadIconFromPath(Part10LargePath), largeIcon)
            Dim part10SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part10SmallPath), LoadIconFromPath(Part10SmallPath), smallIcon)
            Dim part11LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part11LargePath), LoadIconFromPath(Part11LargePath), largeIcon)
            Dim part11SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part11SmallPath), LoadIconFromPath(Part11SmallPath), smallIcon)
            Dim part12LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part12LargePath), LoadIconFromPath(Part12LargePath), largeIcon)
            Dim part12SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part12SmallPath), LoadIconFromPath(Part12SmallPath), smallIcon)
            Dim part13LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part13LargePath), LoadIconFromPath(Part13LargePath), largeIcon)
            Dim part13SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part13SmallPath), LoadIconFromPath(Part13SmallPath), smallIcon)
            Dim part14LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part14LargePath), LoadIconFromPath(Part14LargePath), largeIcon)
            Dim part14SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part14SmallPath), LoadIconFromPath(Part14SmallPath), smallIcon)
            Dim part15LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part15LargePath), LoadIconFromPath(Part15LargePath), largeIcon)
            Dim part15SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(Part15SmallPath), LoadIconFromPath(Part15SmallPath), smallIcon)

#End Region

            Dim partBtn1 As ButtonDefinition = controlDefs.AddButtonDefinition("Thêm dung sai vào dim sketch", "ToolInventor2025_Part_Btn1", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                    "Thêm dung sai vào dim sketch" & vbCrLf & "Mô tả: Thêm dung sai (tolerance) vào dimension trong sketch." & vbCrLf & "Ghi chú: chỉnh theo nhu cầu.",
                                                                     "Tooltip: Thêm dung sai vào dim sketch (edit)", part1SmallIcon, part1LargeIcon)
            AddHandler partBtn1.OnExecute, AddressOf Part.Buttons.solid.Part_Solid_1.OnExecute
            buttonsList.Add(partBtn1)

            Dim partBtn2 As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa fix all sketch", "ToolInventor2025_Part_Btn2", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Xóa fix all sketch" & vbCrLf & "Mô tả: Xóa trạng thái 'fix' cho tất cả sketch entities.",
                                                                               "Tooltip: Xóa fix all sketch (edit)", part2SmallIcon, part2LargeIcon)
            AddHandler partBtn2.OnExecute, AddressOf Part.Buttons.solid.Part_Solid_2.OnExecute
            buttonsList.Add(partBtn2)

            Dim partBtn3 As ButtonDefinition = controlDefs.AddButtonDefinition("Tạo coil line 3d", "ToolInventor2025_Part_Btn3", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Tạo coil line 3d" & vbCrLf & "Mô tả: Tạo đường coil 3D từ sketch hiện tại.",
                                                                               "Tooltip: Tạo coil line 3d (edit)", part3SmallIcon, part3LargeIcon)
            AddHandler partBtn3.OnExecute, AddressOf Part.Buttons.solid.Part_Solid_3.OnExecute
            buttonsList.Add(partBtn3)

            Dim partBtn4 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay tên body", "ToolInventor2025_Part_Btn4", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Thay tên body" & vbCrLf & "Mô tả: Đổi tên body theo quy tắc đặt tên của bạn.",
                                                                               "Tooltip: Thay tên body (edit)", part4SmallIcon, part4LargeIcon)
            AddHandler partBtn4.OnExecute, AddressOf Part.Buttons.solid.Part_Solid_4.OnExecute
            buttonsList.Add(partBtn4)

            Dim partBtn5 As ButtonDefinition = controlDefs.AddButtonDefinition("Export từng body sang các định dạng khác", "ToolInventor2025_Part_Btn5", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Export từng body sang các định dạng khác" & vbCrLf & "Mô tả: Xuất từng body ra file STEP/IGES/...",
                                                                               "Tooltip: Export từng body (edit)", part5SmallIcon, part5LargeIcon)
            AddHandler partBtn5.OnExecute, AddressOf Part.Buttons.solid.Part_Solid_5.OnExecute
            buttonsList.Add(partBtn5)

            Dim partBtn6 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay tên body part theo thứ tự", "ToolInventor2025_Part_Btn6", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Thay tên body part theo thứ tự" & vbCrLf & "Mô tả: Đổi tên body theo thứ tự trong danh sách.",
                                                                               "Tooltip: Thay tên body theo thứ tự (edit)", part6SmallIcon, part6LargeIcon)
            AddHandler partBtn6.OnExecute, AddressOf Part.Buttons.solid.Part_Solid_6.OnExecute
            buttonsList.Add(partBtn6)

            Dim partBtn7 As ButtonDefinition = controlDefs.AddButtonDefinition("Import step to part", "ToolInventor2025_Part_Btn7", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                              "1, Import all file to part tự lưu, xóa liên kết lưu file tự động " & vbCrLf &
                                                                               "Có thể chọn nhiều file 1 lúc",, part7SmallIcon, part7LargeIcon)
            AddHandler partBtn7.OnExecute, AddressOf Assembly.Buttons.Lenhngoaicumlap.Im_EX_step_part.OnExecute
            buttonsList.Add(partBtn7)

            Dim partBtn8 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay màu part", "ToolInventor2025_Part_Btn8", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Thay màu part" & vbCrLf & "Thay mà theo vật liệu Steel",
                                                                               Nothing, part8SmallIcon, part8LargeIcon)
            AddHandler partBtn8.OnExecute, AddressOf Part.Buttons.Part_Solid_7e.OnExecute
            buttonsList.Add(partBtn8)

            Dim partBtn9 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 9", "ToolInventor2025_Part_Btn9", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Part Action 9" & vbCrLf & "Mô tả: (chỉnh sửa) Action 9 cho Part.",
                                                                               "Tooltip: Part Action 9 (edit)", part9SmallIcon, part9LargeIcon)
            AddHandler partBtn9.OnExecute, AddressOf Part.Buttons.Button9.OnExecute
            '   buttonsList.Add(partBtn9)

            Dim partBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 10", "ToolInventor2025_Part_Btn10", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                 "Part Action 10" & vbCrLf & "Mô tả: (chỉnh sửa) Action 10 cho Part.",
                                                                                 "Tooltip: Part Action 10 (edit)", part10SmallIcon, part10LargeIcon)
            AddHandler partBtn10.OnExecute, AddressOf Part.Buttons.Button10.OnExecute
            ' buttonsList.Add(partBtn10)

            Dim partBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 11", "ToolInventor2025_Part_Btn11", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                "Part Action 11" & vbCrLf & "Mô tả: (chỉnh sửa) Action 11 cho Part.",
                                                                                "Tooltip: Part Action 11 (edit)", part11SmallIcon, part11LargeIcon)
            AddHandler partBtn11.OnExecute, AddressOf Part.Buttons.Button11.OnExecute
            '  buttonsList.Add(partBtn11)

            Dim partBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 12", "ToolInventor2025_Part_Btn12", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                "Part Action 12" & vbCrLf & "Mô tả: (chỉnh sửa) Action 12 cho Part.",
                                                                                "Tooltip: Part Action 12 (edit)", part12SmallIcon, part12LargeIcon)
            AddHandler partBtn12.OnExecute, AddressOf Part.Buttons.Button12.OnExecute
            '  buttonsList.Add(partBtn12)

            Dim partBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 13", "ToolInventor2025_Part_Btn13", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                "Part Action 13" & vbCrLf & "Mô tả: (chỉnh sửa) Action 13 cho Part.",
                                                                                "Tooltip: Part Action 13 (edit)", part13SmallIcon, part13LargeIcon)
            AddHandler partBtn13.OnExecute, AddressOf Part.Buttons.Button13.OnExecute
            '  buttonsList.Add(partBtn13)

            Dim partBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 14", "ToolInventor2025_Part_Btn14", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                "Part Action 14" & vbCrLf & "Mô tả: (chỉnh sửa) Action 14 cho Part.",
                                                                                "Tooltip: Part Action 14 (edit)", part14SmallIcon, part14LargeIcon)
            AddHandler partBtn14.OnExecute, AddressOf Part.Buttons.Button14.OnExecute
            ' buttonsList.Add(partBtn14)






            ''''''''' tool ngoài
            Dim partBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Mở nơi lưu File", "ToolInventor2025_Part_Btn15", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                Nothing,
                                                                               "Mở vị trí lưu file", part15SmallIcon, part15LargeIcon)
            AddHandler partBtn15.OnExecute, AddressOf Toolngoai.Vitrifile.Vitrifile
            buttonsList.Add(partBtn15)
        End Sub
    End Class
End Namespace
