Imports Inventor
Imports ToolInventor2025.ToolInventor2025.Drawing.Buttons.Drawtext

Namespace ToolInventor2025
    Public Class DrawingButtons
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
            Dim Dra1LargePath1 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath1 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath2 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath2 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath3 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath3 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath4 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath4 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath5 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath5 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath6 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath6 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath7 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath7 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath8 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath8 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath9 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath9 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath10 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath10 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath11 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath11 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath12 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath12 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath13 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath13 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath14 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath14 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath15 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath15 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath16 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath16 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath17 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath17 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath18 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath18 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath19 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath19 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath20 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath20 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath21 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath21 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath22 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath22 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath23 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath23 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath24 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath24 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath25 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath25 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath26 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath26 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath27 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim Dra1SmallPath27 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim Dra1LargePath28 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")

            ' Prepare IPictureDisp icons (use provided largeIcon/smallIcon as fallback)
            Dim Dra1LargeIcon1 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath1), LoadIconFromPath(Dra1LargePath1), largeIcon)
            Dim Dra1SmallIcon1 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath1), LoadIconFromPath(Dra1SmallPath1), smallIcon)
            Dim Dra1LargeIcon2 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath2), LoadIconFromPath(Dra1LargePath2), largeIcon)
            Dim Dra1SmallIcon2 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath2), LoadIconFromPath(Dra1SmallPath2), smallIcon)
            Dim Dra1LargeIcon3 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath3), LoadIconFromPath(Dra1LargePath3), largeIcon)
            Dim Dra1SmallIcon3 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath3), LoadIconFromPath(Dra1SmallPath3), smallIcon)
            Dim Dra1LargeIcon4 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath4), LoadIconFromPath(Dra1LargePath4), largeIcon)
            Dim Dra1SmallIcon4 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath4), LoadIconFromPath(Dra1SmallPath4), smallIcon)
            Dim Dra1LargeIcon5 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath5), LoadIconFromPath(Dra1LargePath5), largeIcon)
            Dim Dra1SmallIcon5 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath5), LoadIconFromPath(Dra1SmallPath5), smallIcon)
            Dim Dra1LargeIcon6 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath6), LoadIconFromPath(Dra1LargePath6), largeIcon)
            Dim Dra1SmallIcon6 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath6), LoadIconFromPath(Dra1SmallPath6), smallIcon)
            Dim Dra1LargeIcon7 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath7), LoadIconFromPath(Dra1LargePath7), largeIcon)
            Dim Dra1SmallIcon7 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath7), LoadIconFromPath(Dra1SmallPath7), smallIcon)
            Dim Dra1LargeIcon8 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath8), LoadIconFromPath(Dra1LargePath8), largeIcon)
            Dim Dra1SmallIcon8 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath8), LoadIconFromPath(Dra1SmallPath8), smallIcon)
            Dim Dra1LargeIcon9 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath9), LoadIconFromPath(Dra1LargePath9), largeIcon)
            Dim Dra1SmallIcon9 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath9), LoadIconFromPath(Dra1SmallPath9), smallIcon)
            Dim Dra1LargeIcon10 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath10), LoadIconFromPath(Dra1LargePath10), largeIcon)
            Dim Dra1SmallIcon10 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath10), LoadIconFromPath(Dra1SmallPath10), smallIcon)
            Dim Dra1LargeIcon11 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath11), LoadIconFromPath(Dra1LargePath11), largeIcon)
            Dim Dra1SmallIcon11 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath11), LoadIconFromPath(Dra1SmallPath11), smallIcon)
            Dim Dra1LargeIcon12 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath12), LoadIconFromPath(Dra1LargePath12), largeIcon)
            Dim Dra1SmallIcon12 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath12), LoadIconFromPath(Dra1SmallPath12), smallIcon)
            Dim Dra1LargeIcon13 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath13), LoadIconFromPath(Dra1LargePath13), largeIcon)
            Dim Dra1SmallIcon13 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath13), LoadIconFromPath(Dra1SmallPath13), smallIcon)
            Dim Dra1LargeIcon14 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath14), LoadIconFromPath(Dra1LargePath14), largeIcon)
            Dim Dra1SmallIcon14 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath14), LoadIconFromPath(Dra1SmallPath14), smallIcon)
            Dim Dra1LargeIcon15 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath15), LoadIconFromPath(Dra1LargePath15), largeIcon)
            Dim Dra1SmallIcon15 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath15), LoadIconFromPath(Dra1SmallPath15), smallIcon)
            Dim Dra1LargeIcon16 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath16), LoadIconFromPath(Dra1LargePath16), largeIcon)
            Dim Dra1SmallIcon16 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath16), LoadIconFromPath(Dra1SmallPath16), smallIcon)
            Dim Dra1LargeIcon17 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath17), LoadIconFromPath(Dra1LargePath17), largeIcon)
            Dim Dra1SmallIcon17 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath17), LoadIconFromPath(Dra1SmallPath17), smallIcon)
            Dim Dra1LargeIcon18 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath18), LoadIconFromPath(Dra1LargePath18), largeIcon)
            Dim Dra1SmallIcon18 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath18), LoadIconFromPath(Dra1SmallPath18), smallIcon)
            Dim Dra1LargeIcon19 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath19), LoadIconFromPath(Dra1LargePath19), largeIcon)
            Dim Dra1SmallIcon19 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath19), LoadIconFromPath(Dra1SmallPath19), smallIcon)
            Dim Dra1LargeIcon20 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath20), LoadIconFromPath(Dra1LargePath20), largeIcon)
            Dim Dra1SmallIcon20 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath20), LoadIconFromPath(Dra1SmallPath20), smallIcon)
            Dim Dra1LargeIcon21 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath21), LoadIconFromPath(Dra1LargePath21), largeIcon)
            Dim Dra1SmallIcon21 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath21), LoadIconFromPath(Dra1SmallPath21), smallIcon)
            Dim Dra1LargeIcon22 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath22), LoadIconFromPath(Dra1LargePath22), largeIcon)
            Dim Dra1SmallIcon22 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath22), LoadIconFromPath(Dra1SmallPath22), smallIcon)
            Dim Dra1LargeIcon23 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1LargePath23), LoadIconFromPath(Dra1LargePath23), largeIcon)
            Dim Dra1SmallIcon23 As stdole.IPictureDisp = If(System.IO.File.Exists(Dra1SmallPath23), LoadIconFromPath(Dra1SmallPath23), smallIcon)
#End Region


            ' Create Drawing buttons explicitly (no loop) so each button can have distinct implementation
            Dim drawingBtn1 As ButtonDefinition = controlDefs.AddButtonDefinition("Sửa số thập phân dim", "ToolInventor2025_Drawing_Btn1", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Dra1SmallIcon1, Dra1LargeIcon1)
            AddHandler drawingBtn1.OnExecute, AddressOf Drawing.Buttons.Draw_1.OnExecute
            buttonsList.Add(drawingBtn1)

            Dim DrawingBtn2 As ButtonDefinition = controlDefs.AddButtonDefinition("Auto Dim", "ToolInventor2025_Drawing_Btn2", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "Auto dim kích thước line, lỗ, lỗ ren", Dra1SmallIcon2, Dra1LargeIcon2)
            AddHandler DrawingBtn2.OnExecute, AddressOf Drawing.Buttons.Drawdim.draw_2.OnExecute
            buttonsList.Add(DrawingBtn2)

            Dim DrawingBtn3 As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa, căn chỉnh dim", "ToolInventor2025_Drawing_Btn3", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing,
                                                                                   "Xóa, căn chỉnh dim bằng arange", Dra1SmallIcon3, Dra1LargeIcon3)
            AddHandler DrawingBtn3.OnExecute, AddressOf Drawing.Buttons.Drawdim.draw_3.OnExecute
            buttonsList.Add(DrawingBtn3)

            Dim DrawingBtn4 As ButtonDefinition = controlDefs.AddButtonDefinition("Tìm Dim bị edit", "ToolInventor2025_Drawing_Btn4", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "Kiểm tra kim bị edit rồi hiện màu đỏ những dim bị chỉnh sửa", Dra1SmallIcon4, Dra1LargeIcon4)
            AddHandler DrawingBtn4.OnExecute, AddressOf Drawing.Buttons.Draw_4.OnExecute
            buttonsList.Add(DrawingBtn4)

            Dim DrawingBtn5 As ButtonDefinition = controlDefs.AddButtonDefinition("Reset part list", "ToolInventor2025_Drawing_Btn5", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Dra1SmallIcon5, Dra1LargeIcon5)
            AddHandler DrawingBtn5.OnExecute, AddressOf Drawing.Buttons.Draw_ResetPartsList.OnExecute
            buttonsList.Add(DrawingBtn5)

            Dim DrawingBtn6 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi Partlist ENG,VIE", "ToolInventor2025_Drawing_Btn6", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                  "1,2. Update Drawing Views trong sheet hoặc all sheet thay vì ấn bằng tay từng sheet 1" & vbLf &
                                                                                   "3. Ghi mã chi tiết vào category trong partlist. yêu cầu partlist phải có sẵn ô này thì mới được!", Dra1SmallIcon6, Dra1LargeIcon6)
            AddHandler DrawingBtn6.OnExecute, AddressOf Drawing.Buttons.DrawPartList.Draw_Partlist_1.OnExecute
            buttonsList.Add(DrawingBtn6)

            Dim DrawingBtn7 As ButtonDefinition = controlDefs.AddButtonDefinition("Nút chuyển Sheet", "ToolInventor2025_Drawing_Btn7", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Dra1SmallIcon7, Dra1LargeIcon7)
            AddHandler DrawingBtn7.OnExecute, AddressOf Drawing.Buttons.Draw_7.OnExecute
            buttonsList.Add(DrawingBtn7)

            Dim DrawingBtn8 As ButtonDefinition = controlDefs.AddButtonDefinition("Đổi scale view", "ToolInventor2025_Drawing_Btn8", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Dra1SmallIcon8, Dra1LargeIcon8)
            AddHandler DrawingBtn8.OnExecute, AddressOf Drawing.Buttons.DrawView.Draw_Change_scale_view.OnExecute
            buttonsList.Add(DrawingBtn8)

            Dim DrawingBtn9 As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa Part,ASS trùng lặp", "ToolInventor2025_Drawing_Btn9", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Dra1SmallIcon9, Dra1LargeIcon9)
            AddHandler DrawingBtn9.OnExecute, AddressOf Drawing.Buttons.Draw_9.OnExecute
            buttonsList.Add(DrawingBtn9)

            Dim DrawingBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa centermark, centerline", "ToolInventor2025_Drawing_Btn10", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Dra1SmallIcon10, Dra1LargeIcon10)
            AddHandler DrawingBtn10.OnExecute, AddressOf Drawing.Buttons.Draw_5.OnExecute
            buttonsList.Add(DrawingBtn10)
            '
            Dim DrawingBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("Update Views", "ToolInventor2025_Drawing_Btn11", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Update Drawing Views trong sheet hoặc all sheet thay vì ấn bằng tay từng sheet 1", Dra1SmallIcon11, Dra1LargeIcon11)
            AddHandler DrawingBtn11.OnExecute, AddressOf Drawing.Buttons.DrawView.Draw_Up_date_view.OnExecute
            buttonsList.Add(DrawingBtn11)

            Dim DrawingBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Rename sheet    ", "ToolInventor2025_Drawing_Btn12", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Thay tên sheet theo part number", Dra1SmallIcon12, Dra1LargeIcon12)
            AddHandler DrawingBtn12.OnExecute, AddressOf Drawing.Buttons.Draw_Rename_Sheet.OnExecute
            buttonsList.Add(DrawingBtn12)

            Dim DrawingBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay chữ cái view sheet", "ToolInventor2025_Drawing_Btn13", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Đổi tên view hiển thị trong sheet", Dra1SmallIcon13, Dra1LargeIcon13)
            AddHandler DrawingBtn13.OnExecute, AddressOf Drawing.Buttons.DrawView.Draw_Rename_View.OnExecute
            buttonsList.Add(DrawingBtn13)

            Dim DrawingBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Đổi scale Hatch", "ToolInventor2025_Drawing_Btn14", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Chưa ok", Dra1SmallIcon14, Dra1LargeIcon14)
            AddHandler DrawingBtn14.OnExecute, AddressOf Drawing.Buttons.draw_14.OnExecute
            buttonsList.Add(DrawingBtn14)

            Dim DrawingBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Delete", "ToolInventor2025_Drawing_Btn15", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Xóa các thứ liên quan dến text. ko hỗ trợ xóa weld cho bản 2024 trở xuống", Dra1SmallIcon15, Dra1LargeIcon15)
            AddHandler DrawingBtn15.OnExecute, AddressOf Drawing.Buttons.Drawdelete.Draw_delete.OnExecute
            buttonsList.Add(DrawingBtn15)


            Dim DrawingBtn17 As ButtonDefinition = controlDefs.AddButtonDefinition("Dim Kích thước tổng", "ToolInventor2025_Drawing_Btn17", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Dra1SmallIcon17, Dra1LargeIcon17)
            AddHandler DrawingBtn17.OnExecute, AddressOf Drawing.Buttons.Drawdim.Dim_Linear_Total.OnExecute

            buttonsList.Add(DrawingBtn17)
            Dim DrawingBtn18 As ButtonDefinition = controlDefs.AddButtonDefinition("Create_PartsList", "ToolInventor2025_Drawing_Btn18", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Dra1SmallIcon18, Dra1LargeIcon18)
            AddHandler DrawingBtn18.OnExecute, AddressOf Drawing.Buttons.DrawPartList.Create_PartsList.OnExecute

            buttonsList.Add(DrawingBtn18)

            Dim DrawingBtn19 As ButtonDefinition = controlDefs.AddButtonDefinition("Text Edit", "ToolInventor2025_Drawing_Btn19", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Thay thế văn bản" & vbLf & "Đổi chữ hoa/thường" & vbLf & "Xóa hoặc đổi chữ", Dra1SmallIcon19, Dra1LargeIcon19)
            AddHandler DrawingBtn19.OnExecute, AddressOf Drawing.Buttons.Drawtext.draw_text_list_1.OnExecute

            buttonsList.Add(DrawingBtn19)

            Dim DrawingBtn20 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay khung tên, Border", "ToolInventor2025_Drawing_Btn20", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Thay thế khung tên, Border", Dra1SmallIcon20, Dra1LargeIcon20)
            AddHandler DrawingBtn20.OnExecute, AddressOf Drawing.Buttons.DrawSheet.Drawing_TitleBlock_Border.OnExecute

            buttonsList.Add(DrawingBtn20)

            Dim DrawingBtn21 As ButtonDefinition = controlDefs.AddButtonDefinition("Đông bọ partlist", "ToolInventor2025_Drawing_Btn21", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Thay thế văn bản", Dra1SmallIcon21, Dra1LargeIcon21)
            AddHandler DrawingBtn21.OnExecute, AddressOf Drawing.Buttons.DrawSheet.Drawing_PartsList_Sync2020.OnExecute

            buttonsList.Add(DrawingBtn21)

            ''''''''' tool ngoài

            Dim DrawingBtn16 As ButtonDefinition = controlDefs.AddButtonDefinition("Mở nơi lưu File", "ToolInventor2025_Drawing_Btn16", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Mở vị trí lưu file", Dra1SmallIcon16, Dra1LargeIcon16)
            AddHandler DrawingBtn16.OnExecute, AddressOf Toolngoai.Vitrifile.Vitrifile

            buttonsList.Add(DrawingBtn16)



        End Sub
    End Class
End Namespace
