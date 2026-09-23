Imports Inventor

Namespace ToolInventor2025
    Public Class Assembly3Buttons
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

            Dim assemblyFolder As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
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
                iconsFolder = System.IO.Path.Combine(assemblyFolder, "Code addin inventor", "Images", "Assembly")
            End If

            Dim Ass3LargePath1 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath1 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath2 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath2 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath3 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath3 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath4 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath4 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath5 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath5 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath6 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath6 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath7 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath7 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath8 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath8 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath9 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath9 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath10 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath10 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath11 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath11 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath12 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath12 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath13 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath13 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath14 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath14 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath15 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath15 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath16 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath16 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath17 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath17 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath18 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath18 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath19 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath19 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath20 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath20 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath21 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath21 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath22 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath22 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath23 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath23 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath24 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath24 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath25 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath25 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath26 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath26 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath27 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath27 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass3LargePath28 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass3SmallPath28 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")


            ' Load per-button icons (fallback to provided largeIcon/smallIcon when file missing)
            Dim Ass3LargeIcon1 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath1), LoadIconFromPath(Ass3LargePath1), largeIcon)
            Dim Ass3SmallIcon1 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath1), LoadIconFromPath(Ass3SmallPath1), smallIcon)
            Dim Ass3LargeIcon2 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath2), LoadIconFromPath(Ass3LargePath2), largeIcon)
            Dim Ass3SmallIcon2 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath2), LoadIconFromPath(Ass3SmallPath2), smallIcon)
            Dim Ass3LargeIcon3 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath3), LoadIconFromPath(Ass3LargePath3), largeIcon)
            Dim Ass3SmallIcon3 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath3), LoadIconFromPath(Ass3SmallPath3), smallIcon)
            Dim Ass3LargeIcon4 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath4), LoadIconFromPath(Ass3LargePath4), largeIcon)
            Dim Ass3SmallIcon4 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath4), LoadIconFromPath(Ass3SmallPath4), smallIcon)
            Dim Ass3LargeIcon5 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath5), LoadIconFromPath(Ass3LargePath5), largeIcon)
            Dim Ass3SmallIcon5 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath5), LoadIconFromPath(Ass3SmallPath5), smallIcon)
            Dim Ass3LargeIcon6 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath6), LoadIconFromPath(Ass3LargePath6), largeIcon)
            Dim Ass3SmallIcon6 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath6), LoadIconFromPath(Ass3SmallPath6), smallIcon)
            Dim Ass3LargeIcon7 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath7), LoadIconFromPath(Ass3LargePath7), largeIcon)
            Dim Ass3SmallIcon7 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath7), LoadIconFromPath(Ass3SmallPath7), smallIcon)
            Dim Ass3LargeIcon8 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath8), LoadIconFromPath(Ass3LargePath8), largeIcon)
            Dim Ass3SmallIcon8 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath8), LoadIconFromPath(Ass3SmallPath8), smallIcon)
            Dim Ass3LargeIcon9 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath9), LoadIconFromPath(Ass3LargePath9), largeIcon)
            Dim Ass3SmallIcon9 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath9), LoadIconFromPath(Ass3SmallPath9), smallIcon)
            Dim Ass3LargeIcon10 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath10), LoadIconFromPath(Ass3LargePath10), largeIcon)
            Dim Ass3SmallIcon10 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath10), LoadIconFromPath(Ass3SmallPath10), smallIcon)
            Dim Ass3LargeIcon11 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath11), LoadIconFromPath(Ass3LargePath11), largeIcon)
            Dim Ass3SmallIcon11 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath11), LoadIconFromPath(Ass3SmallPath11), smallIcon)
            Dim Ass3LargeIcon12 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath12), LoadIconFromPath(Ass3LargePath12), largeIcon)
            Dim Ass3SmallIcon12 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath12), LoadIconFromPath(Ass3SmallPath12), smallIcon)
            Dim Ass3LargeIcon13 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath13), LoadIconFromPath(Ass3LargePath13), largeIcon)
            Dim Ass3SmallIcon13 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath13), LoadIconFromPath(Ass3SmallPath13), smallIcon)
            Dim Ass3LargeIcon14 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath14), LoadIconFromPath(Ass3LargePath14), largeIcon)
            Dim Ass3SmallIcon14 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath14), LoadIconFromPath(Ass3SmallPath14), smallIcon)
            Dim Ass3LargeIcon15 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath15), LoadIconFromPath(Ass3LargePath15), largeIcon)
            Dim Ass3SmallIcon15 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath15), LoadIconFromPath(Ass3SmallPath15), smallIcon)
            Dim Ass3LargeIcon16 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath16), LoadIconFromPath(Ass3LargePath16), largeIcon)
            Dim Ass3SmallIcon16 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath16), LoadIconFromPath(Ass3SmallPath16), smallIcon)
            Dim Ass3LargeIcon17 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath17), LoadIconFromPath(Ass3LargePath17), largeIcon)
            Dim Ass3SmallIcon17 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath17), LoadIconFromPath(Ass3SmallPath17), smallIcon)
            Dim Ass3LargeIcon18 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath18), LoadIconFromPath(Ass3LargePath18), largeIcon)
            Dim Ass3SmallIcon18 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath18), LoadIconFromPath(Ass3SmallPath18), smallIcon)
            Dim Ass3LargeIcon19 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath19), LoadIconFromPath(Ass3LargePath19), largeIcon)
            Dim Ass3SmallIcon19 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath19), LoadIconFromPath(Ass3SmallPath19), smallIcon)
            Dim Ass3LargeIcon20 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath20), LoadIconFromPath(Ass3LargePath20), largeIcon)
            Dim Ass3SmallIcon20 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath20), LoadIconFromPath(Ass3SmallPath20), smallIcon)
            Dim Ass3LargeIcon21 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath21), LoadIconFromPath(Ass3LargePath21), largeIcon)
            Dim Ass3SmallIcon21 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath21), LoadIconFromPath(Ass3SmallPath21), smallIcon)
            Dim Ass3LargeIcon22 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath22), LoadIconFromPath(Ass3LargePath22), largeIcon)
            Dim Ass3SmallIcon22 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath22), LoadIconFromPath(Ass3SmallPath22), smallIcon)
            Dim Ass3LargeIcon23 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath23), LoadIconFromPath(Ass3LargePath23), largeIcon)
            Dim Ass3SmallIcon23 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath23), LoadIconFromPath(Ass3SmallPath23), smallIcon)
            Dim Ass3LargeIcon24 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath24), LoadIconFromPath(Ass3LargePath24), largeIcon)
            Dim Ass3SmallIcon24 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath24), LoadIconFromPath(Ass3SmallPath24), smallIcon)
            Dim Ass3LargeIcon25 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath25), LoadIconFromPath(Ass3LargePath25), largeIcon)
            Dim Ass3SmallIcon25 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath25), LoadIconFromPath(Ass3SmallPath25), smallIcon)
            Dim Ass3LargeIcon26 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath26), LoadIconFromPath(Ass3LargePath26), largeIcon)
            Dim Ass3SmallIcon26 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath26), LoadIconFromPath(Ass3SmallPath26), smallIcon)
            Dim Ass3LargeIcon27 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath27), LoadIconFromPath(Ass3LargePath27), largeIcon)
            Dim Ass3SmallIcon27 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath27), LoadIconFromPath(Ass3SmallPath27), smallIcon)
            Dim Ass3LargeIcon28 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3LargePath28), LoadIconFromPath(Ass3LargePath28), largeIcon)
            Dim Ass3SmallIcon28 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass3SmallPath28), LoadIconFromPath(Ass3SmallPath28), smallIcon)

            ' Create Assembly buttons explicitly (no loop) so each button can have distinct implementation
            Dim assemblyBtnb1 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi tên theo chuỗi Top lever Partnumber", "ToolInventor2025_Assembly_Btnb1", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass3SmallIcon1, Ass3LargeIcon1)
            AddHandler assemblyBtnb1.OnExecute, AddressOf Assembly3.Buttons.Button1.OnExecute
            buttonsList.Add(assemblyBtnb1)

            Dim assemblyBtnb2 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi tên theo chuỗi Top lever Stocknumber", "ToolInventor2025_Assembly_Btnb2", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass3SmallIcon2, Ass3LargeIcon2)
            AddHandler assemblyBtnb2.OnExecute, AddressOf Assembly3.Buttons.Button2.OnExecute
            buttonsList.Add(assemblyBtnb2)

            Dim assemblyBtnb3 As ButtonDefinition = controlDefs.AddButtonDefinition("Đánh STT cho item1 Top lever", "ToolInventor2025_Assembly_Btnb3", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass3SmallIcon3, Ass3LargeIcon3)
            AddHandler assemblyBtnb3.OnExecute, AddressOf Assembly3.Buttons.Button3.OnExecute
            buttonsList.Add(assemblyBtnb3)

            Dim assemblyBtnb4 As ButtonDefinition = controlDefs.AddButtonDefinition("Đánh STT cho cụm xếp VT partnumber Top lever", "ToolInventor2025_Assembly_Btnb4", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass3SmallIcon4, Ass3LargeIcon4)
            AddHandler assemblyBtnb4.OnExecute, AddressOf Assembly3.Buttons.Button4.OnExecute
            buttonsList.Add(assemblyBtnb4)

            Dim assemblyBtnb5 As ButtonDefinition = controlDefs.AddButtonDefinition("Đánh STT, item1 VT Buy Top lever về cuối", "ToolInventor2025_Assembly_Btnb5", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass3SmallIcon5, Ass3LargeIcon5)
            AddHandler assemblyBtnb5.OnExecute, AddressOf Assembly3.Buttons.Button5.OnExecute
            buttonsList.Add(assemblyBtnb5)

            Dim assemblyBtnb6 As ButtonDefinition = controlDefs.AddButtonDefinition("Đánh STT top lever", "ToolInventor2025_Assembly_Btnb6", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass3SmallIcon6, Ass3LargeIcon6)
            AddHandler assemblyBtnb6.OnExecute, AddressOf Assembly3.Buttons.Button6.OnExecute
            buttonsList.Add(assemblyBtnb6)

            Dim assemblyBtnb7 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi tên PLxx sheetmetal vào trong PartNB", "ToolInventor2025_Assembly_Btnb7", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass3SmallIcon7, Ass3LargeIcon7)
            AddHandler assemblyBtnb7.OnExecute, AddressOf Assembly3.Buttons.Button7.OnExecute
            buttonsList.Add(assemblyBtnb7)

            Dim assemblyBtnb8 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi tên PLxx sheetmetal vào trong StockNB", "ToolInventor2025_Assembly_Btnb8", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Code này sẽ thay tên stocknumber 
theo kích thước bao sau khi trải ", Ass3SmallIcon8, Ass3LargeIcon8)
            AddHandler assemblyBtnb8.OnExecute, AddressOf Assembly3.Buttons.Button8.OnExecute
            buttonsList.Add(assemblyBtnb8)

            Dim assemblyBtnb9 As ButtonDefinition = controlDefs.AddButtonDefinition("Điền thông tin chiều dày tấm vào PL BOM", "ToolInventor2025_Assembly_Btnb9", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "
Điền thông tin chiều dày theo thickness. chỉ áp dụng với cá sheetmetal! = t1,t2,t3,t4,...", Ass3SmallIcon9, Ass3LargeIcon9)
            AddHandler assemblyBtnb9.OnExecute, AddressOf Assembly3.Buttons.Button9.OnExecute
            buttonsList.Add(assemblyBtnb9)

            Dim assemblyBtnb10 As ButtonDefinition = controlDefs.AddButtonDefinition("STT từ Item sang item1 chỉ dánh part top lever", "ToolInventor2025_Assembly_Btnb10", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Code này chỉ coppy SST sang item1 cho part
toplever", Ass3SmallIcon10, Ass3LargeIcon10)
            AddHandler assemblyBtnb10.OnExecute, AddressOf Assembly3.Buttons.Button10.OnExecute
            buttonsList.Add(assemblyBtnb10)

            Dim assemblyBtnb11 As ButtonDefinition = controlDefs.AddButtonDefinition("Item, Item Qty part sang Item1, SL Part all lever", "ToolInventor2025_Assembly_Btnb11", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Coppy Item, Item Qty part all leveer sang Item1, SL Part all lever
tối đa 3 lever & không copy vào cụm lắp. chỉ copy part", Ass3SmallIcon11, Ass3LargeIcon11)
            AddHandler assemblyBtnb11.OnExecute, AddressOf Assembly3.Buttons.Button11.OnExecute
            buttonsList.Add(assemblyBtnb11)

            Dim assemblyBtnb12 As ButtonDefinition = controlDefs.AddButtonDefinition("Item, Item Qty part sang Item1, SL Part, PL all lever", "ToolInventor2025_Assembly_Btnb12", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Coppy Item, Item Qty part all leveer sang Item1, SL Part, thickness PL sheetmetal all lever
tối đa 3 lever & không copy vào cụm lắp. chỉ copy part", Ass3SmallIcon12, Ass3LargeIcon12)
            AddHandler assemblyBtnb12.OnExecute, AddressOf Assembly3.Buttons.Button12.OnExecute
            buttonsList.Add(assemblyBtnb12)

            Dim assemblyBtnb13 As ButtonDefinition = controlDefs.AddButtonDefinition("Item, Item Qty part sang Item1, SL Part, PL all lever", "ToolInventor2025_Assembly_Btnb13", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass3SmallIcon13, Ass3LargeIcon13)
            AddHandler assemblyBtnb13.OnExecute, AddressOf Assembly3.Buttons.Button13.OnExecute
            buttonsList.Add(assemblyBtnb13)

            Dim assemblyBtnb14 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 14", "ToolInventor2025_Assembly_Btnb14", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass3SmallIcon14, Ass3LargeIcon14)
            AddHandler assemblyBtnb14.OnExecute, AddressOf Assembly3.Buttons.Button14.OnExecute
            ' buttonsList.Add(assemblyBtnb14)

            Dim assemblyBtnb15 As ButtonDefinition = controlDefs.AddButtonDefinition("11111111111111111111", "ToolInventor2025_Assembly_Btnb15", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, Ass3SmallIcon15, Ass3LargeIcon15)
            AddHandler assemblyBtnb15.OnExecute, AddressOf Assembly3.Buttons.Button15.OnExecute
            '  buttonsList.Add(assemblyBtnb15)

            ' --- Temporary test buttons in case pulldown is missing ---
            '  Dim assemblyTest1 As ButtonDefinition = controlDefs.AddButtonDefinition("Test Action 1 (Assembly3)", "ToolInventor2025_Assembly_Test1", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Temporary test button", Ass3SmallIcon16, Ass3LargeIcon16)
            ' AddHandler assemblyTest1.OnExecute, AddressOf Assembly3.Buttons.Button1.OnExecute
            'buttonsList.Add(assemblyTest1)

            'Dim assemblyTest2 As ButtonDefinition = controlDefs.AddButtonDefinition("Test Action 2 (Assembly3)", "ToolInventor2025_Assembly_Test2", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Temporary test button", Ass3SmallIcon2, Ass3LargeIcon2)
            'AddHandler assemblyTest2.OnExecute, AddressOf Assembly3.Buttons.Button2.OnExecute
            'buttonsList.Add(assemblyTest2)

        End Sub
    End Class
End Namespace
