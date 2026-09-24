Imports Inventor

Namespace ToolInventor2025
    Public Class Assembly2Buttons
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

            Dim assemblyFolder3 As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim configured As String = Nothing
            Try
                configured = My.Settings.ImageFolder
            Catch
                configured = Nothing
            End Try

            Dim iconsFolder As String = Nothing
            If Not String.IsNullOrWhiteSpace(configured) AndAlso System.IO.Directory.Exists(configured) Then
                iconsFolder = System.IO.Path.Combine(configured, "Assembly")
            Else
                iconsFolder = System.IO.Path.Combine(assemblyFolder3, "Code addin inventor", "Images", "Assembly")
            End If
#Region "button"
            Dim Ass2LargePath1 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath1 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath2 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath2 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath3 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath3 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath4 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath4 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath5 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath5 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath6 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath6 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath7 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath7 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath8 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath8 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath9 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath9 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath10 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath10 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath11 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath11 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath12 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath12 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath13 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath13 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath14 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath14 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath15 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath15 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath16 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath16 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath17 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath17 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath18 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath18 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath19 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath19 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath20 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath20 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath21 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath21 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath22 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath22 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath23 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath23 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath24 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath24 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath25 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath25 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath26 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath26 As String = System.IO.Path.Combine(iconsFolder, "i18 1.bmp")
            Dim Ass2LargePath27 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath27 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2LargePath28 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")
            Dim Ass2SmallPath28 As String = System.IO.Path.Combine(iconsFolder, "i18.bmp")


            ' Load per-button icons (fallback to provided largeIcon/smallIcon when file missing)
            Dim Ass2LargeIcon1 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath1), LoadIconFromPath(Ass2LargePath1), largeIcon)
            Dim Ass2SmallIcon1 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath1), LoadIconFromPath(Ass2SmallPath1), smallIcon)
            Dim Ass2LargeIcon2 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath2), LoadIconFromPath(Ass2LargePath2), largeIcon)
            Dim Ass2SmallIcon2 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath2), LoadIconFromPath(Ass2SmallPath2), smallIcon)
            Dim Ass2LargeIcon3 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath3), LoadIconFromPath(Ass2LargePath3), largeIcon)
            Dim Ass2SmallIcon3 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath3), LoadIconFromPath(Ass2SmallPath3), smallIcon)
            Dim Ass2LargeIcon4 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath4), LoadIconFromPath(Ass2LargePath4), largeIcon)
            Dim Ass2SmallIcon4 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath4), LoadIconFromPath(Ass2SmallPath4), smallIcon)
            Dim Ass2LargeIcon5 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath5), LoadIconFromPath(Ass2LargePath5), largeIcon)
            Dim Ass2SmallIcon5 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath5), LoadIconFromPath(Ass2SmallPath5), smallIcon)
            Dim Ass2LargeIcon6 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath6), LoadIconFromPath(Ass2LargePath6), largeIcon)
            Dim Ass2SmallIcon6 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath6), LoadIconFromPath(Ass2SmallPath6), smallIcon)
            Dim Ass2LargeIcon7 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath7), LoadIconFromPath(Ass2LargePath7), largeIcon)
            Dim Ass2SmallIcon7 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath7), LoadIconFromPath(Ass2SmallPath7), smallIcon)
            Dim Ass2LargeIcon8 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath8), LoadIconFromPath(Ass2LargePath8), largeIcon)
            Dim Ass2SmallIcon8 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath8), LoadIconFromPath(Ass2SmallPath8), smallIcon)
            Dim Ass2LargeIcon9 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath9), LoadIconFromPath(Ass2LargePath9), largeIcon)
            Dim Ass2SmallIcon9 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath9), LoadIconFromPath(Ass2SmallPath9), smallIcon)
            Dim Ass2LargeIcon10 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath10), LoadIconFromPath(Ass2LargePath10), largeIcon)
            Dim Ass2SmallIcon10 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath10), LoadIconFromPath(Ass2SmallPath10), smallIcon)
            Dim Ass2LargeIcon11 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath11), LoadIconFromPath(Ass2LargePath11), largeIcon)
            Dim Ass2SmallIcon11 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath11), LoadIconFromPath(Ass2SmallPath11), smallIcon)
            Dim Ass2LargeIcon12 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath12), LoadIconFromPath(Ass2LargePath12), largeIcon)
            Dim Ass2SmallIcon12 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath12), LoadIconFromPath(Ass2SmallPath12), smallIcon)
            Dim Ass2LargeIcon13 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath13), LoadIconFromPath(Ass2LargePath13), largeIcon)
            Dim Ass2SmallIcon13 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath13), LoadIconFromPath(Ass2SmallPath13), smallIcon)
            Dim Ass2LargeIcon14 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath14), LoadIconFromPath(Ass2LargePath14), largeIcon)
            Dim Ass2SmallIcon14 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath14), LoadIconFromPath(Ass2SmallPath14), smallIcon)
            Dim Ass2LargeIcon15 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath15), LoadIconFromPath(Ass2LargePath15), largeIcon)
            Dim Ass2SmallIcon15 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath15), LoadIconFromPath(Ass2SmallPath15), smallIcon)
            Dim Ass2LargeIcon16 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath16), LoadIconFromPath(Ass2LargePath16), largeIcon)
            Dim Ass2SmallIcon16 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath16), LoadIconFromPath(Ass2SmallPath16), smallIcon)
            Dim Ass2LargeIcon17 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath17), LoadIconFromPath(Ass2LargePath17), largeIcon)
            Dim Ass2SmallIcon17 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath17), LoadIconFromPath(Ass2SmallPath17), smallIcon)
            Dim Ass2LargeIcon18 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath18), LoadIconFromPath(Ass2LargePath18), largeIcon)
            Dim Ass2SmallIcon18 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath18), LoadIconFromPath(Ass2SmallPath18), smallIcon)
            Dim Ass2LargeIcon19 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath19), LoadIconFromPath(Ass2LargePath19), largeIcon)
            Dim Ass2SmallIcon19 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath19), LoadIconFromPath(Ass2SmallPath19), smallIcon)
            Dim Ass2LargeIcon20 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath20), LoadIconFromPath(Ass2LargePath20), largeIcon)
            Dim Ass2SmallIcon20 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath20), LoadIconFromPath(Ass2SmallPath20), smallIcon)
            Dim Ass2LargeIcon21 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath21), LoadIconFromPath(Ass2LargePath21), largeIcon)
            Dim Ass2SmallIcon21 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath21), LoadIconFromPath(Ass2SmallPath21), smallIcon)
            Dim Ass2LargeIcon22 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath22), LoadIconFromPath(Ass2LargePath22), largeIcon)
            Dim Ass2SmallIcon22 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath22), LoadIconFromPath(Ass2SmallPath22), smallIcon)
            Dim Ass2LargeIcon23 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath23), LoadIconFromPath(Ass2LargePath23), largeIcon)
            Dim Ass2SmallIcon23 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath23), LoadIconFromPath(Ass2SmallPath23), smallIcon)
            Dim Ass2LargeIcon24 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath24), LoadIconFromPath(Ass2LargePath24), largeIcon)
            Dim Ass2SmallIcon24 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath24), LoadIconFromPath(Ass2SmallPath24), smallIcon)
            Dim Ass2LargeIcon25 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath25), LoadIconFromPath(Ass2LargePath25), largeIcon)
            Dim Ass2SmallIcon25 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath25), LoadIconFromPath(Ass2SmallPath25), smallIcon)
            Dim Ass2LargeIcon26 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath26), LoadIconFromPath(Ass2LargePath26), largeIcon)
            Dim Ass2SmallIcon26 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath26), LoadIconFromPath(Ass2SmallPath26), smallIcon)
            Dim Ass2LargeIcon27 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath27), LoadIconFromPath(Ass2LargePath27), largeIcon)
            Dim Ass2SmallIcon27 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath27), LoadIconFromPath(Ass2SmallPath27), smallIcon)
            Dim Ass2LargeIcon28 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2LargePath28), LoadIconFromPath(Ass2LargePath28), largeIcon)
            Dim Ass2SmallIcon28 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass2SmallPath28), LoadIconFromPath(Ass2SmallPath28), smallIcon)
#End Region



            ' Create Assembly buttons explicitly (no loop) so each button can have distinct implementation
            Dim assemblyBtna1 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay tên ,STT BOM", "ToolInventor2025_Assembly_Btna1", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass2SmallIcon1, Ass2LargeIcon1)
            AddHandler assemblyBtna1.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_1.OnExecute
            buttonsList.Add(assemblyBtna1)

            Dim assemblyBtna2 As ButtonDefinition = controlDefs.AddButtonDefinition("Item1 Buy Top lever về cuối", "ToolInventor2025_Assembly_Btna2", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass2SmallIcon2, Ass2LargeIcon2)
            AddHandler assemblyBtna2.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_2.OnExecute
            buttonsList.Add(assemblyBtna2)

            Dim assemblyBtna3 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi item1,t,SL", "ToolInventor2025_Assembly_Btna3", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass2SmallIcon3, Ass2LargeIcon3)
            AddHandler assemblyBtna3.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_3.OnExecute
            buttonsList.Add(assemblyBtna3)

            Dim assemblyBtna4 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi PL part,stocknumber", "ToolInventor2025_Assembly_Btna4", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass2SmallIcon4, Ass2LargeIcon4)
            AddHandler assemblyBtna4.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_4.OnExecute
            buttonsList.Add(assemblyBtna4)

            Dim assemblyBtna6 As ButtonDefinition = controlDefs.AddButtonDefinition("Xếp Model Browser ASS", "ToolInventor2025_Assembly_Btna6", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass2SmallIcon6, Ass2LargeIcon6)
            AddHandler assemblyBtna6.OnExecute, AddressOf Assembly2.Buttons.BOMcode.ass_bom_6.OnExecute
            buttonsList.Add(assemblyBtna6)

            Dim assemblyBtna7 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay đổi đuôi tên BOM", "ToolInventor2025_Assembly_Btna7", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                    "Đổi tên part number", Ass2SmallIcon7, Ass2LargeIcon7)
            AddHandler assemblyBtna7.OnExecute, AddressOf Assembly2.Buttons.BOMcode.ass_bom_7.OnExecute
            buttonsList.Add(assemblyBtna7)



            Dim assemblyBtna5 As ButtonDefinition = controlDefs.AddButtonDefinition("Hiện gốc tọa độ để Constrain", "ToolInventor2025_Assembly_Btna5", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                    Nothing, "ấn chọn file part hoặc Assembly để hiện gốc tọa độ" & vbCrLf &
                                                                                   " tiếp theo dùng lệnh Constrain để lắp ghép với nhau." & vbCrLf &
                                                                                    "Nếu ấn ok thì sẽ ẩn hết các mặt phẳng & trục gốc tọa độ vừa ấn hiện", Ass2SmallIcon5, Ass2LargeIcon5)
            'AddHandler assemblyBtna5.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_5.OnExecute
            ' buttonsList.Add(assemblyBtna5)



            Dim assemblyBtna8 As ButtonDefinition = controlDefs.AddButtonDefinition("5", "ToolInventor2025_Assembly_Btna8", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Code này sẽ thay tên stocknumber 
theo kích thước bao sau khi trải ", Ass2SmallIcon8, Ass2LargeIcon8)
            AddHandler assemblyBtna8.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_8.OnExecute
            ' buttonsList.Add(assemblyBtna8)

            Dim assemblyBtna9 As ButtonDefinition = controlDefs.AddButtonDefinition("9", "ToolInventor2025_Assembly_Btna9", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "
Điền thông tin chiều dày theo thickness. chỉ áp dụng với cá sheetmetal! = t1,t2,t3,t4,...", Ass2SmallIcon9, Ass2LargeIcon9)
            AddHandler assemblyBtna9.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_9.OnExecute
            '  buttonsList.Add(assemblyBtna9)

            Dim assemblyBtna10 As ButtonDefinition = controlDefs.AddButtonDefinition("6", "ToolInventor2025_Assembly_Btna10", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Code này chỉ coppy SST sang item1 cho part
toplever", Ass2SmallIcon10, Ass2LargeIcon10)
            AddHandler assemblyBtna10.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_10.OnExecute
            '  buttonsList.Add(assemblyBtna10)

            Dim assemblyBtna11 As ButtonDefinition = controlDefs.AddButtonDefinition("7", "ToolInventor2025_Assembly_Btna11", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Coppy Item, Item Qty part all leveer sang Item1, SL Part all lever
tối đa 3 lever & không copy vào cụm lắp. chỉ copy part", Ass2SmallIcon11, Ass2LargeIcon11)
            AddHandler assemblyBtna11.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_11.OnExecute
            ' buttonsList.Add(assemblyBtna11)

            Dim assemblyBtna12 As ButtonDefinition = controlDefs.AddButtonDefinition("8", "ToolInventor2025_Assembly_Btna12", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Coppy Item, Item Qty part all leveer sang Item1, SL Part, thickness PL sheetmetal all lever
tối đa 3 lever & không copy vào cụm lắp. chỉ copy part", Ass2SmallIcon12, Ass2LargeIcon12)
            AddHandler assemblyBtna12.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_12.OnExecute
            ' buttonsList.Add(assemblyBtna12)

            Dim assemblyBtna13 As ButtonDefinition = controlDefs.AddButtonDefinition("9", "ToolInventor2025_Assembly_Btna13", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass2SmallIcon13, Ass2LargeIcon13)
            AddHandler assemblyBtna13.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_13.OnExecute
            ' buttonsList.Add(assemblyBtna13)

            Dim assemblyBtna14 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 14", "ToolInventor2025_Assembly_Btna14", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass2SmallIcon14, Ass2LargeIcon14)
            AddHandler assemblyBtna14.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_14.OnExecute
            ' buttonsList.Add(assemblyBtna14)

            Dim assemblyBtna15 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 15", "ToolInventor2025_Assembly_Btna15", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass2SmallIcon15, Ass2LargeIcon15)
            AddHandler assemblyBtna15.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_15.OnExecute
            ' buttonsList.Add(assemblyBtna15)

            Dim assemblyBtna16 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 16", "ToolInventor2025_Assembly_Btna16", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass2SmallIcon16, Ass2LargeIcon16)
            'AddHandler assemblyBtna16.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_16.OnExecute
            '  buttonsList.Add(assemblyBtna16)

            Dim assemblyBtna17 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 17", "ToolInventor2025_Assembly_Btna17", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass2SmallIcon17, Ass2LargeIcon17)
            'AddHandler assemblyBtna17.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_17.OnExecute
            ' buttonsList.Add(assemblyBtna17)



        End Sub
    End Class
End Namespace
