Imports System.Diagnostics.Contracts
Imports System.Windows.Forms
Imports Inventor


Namespace ToolInventor2025
#Region "AssemblyButtons"
    Public Class AssemblyButtons
        '════════════════════════════════════════════════════════════════════
        ' REGISTRY cho POPUP MENU
        ' Panel sẽ tạo trong AddToUserInterface → AddTabPanelButtons()
        ' nhưng ở đó không có reference tới các sub-button definitions
        ' → dùng dictionary để "gửi" popup tới đó
        '════════════════════════════════════════════════════════════════════
        Public Class PopupDef
            Public SubButtons As New List(Of ButtonDefinition)  ' Phần tử [0] = nút chính
        End Class

        ''' <summary>
        ''' Key = panelInternalName, Value = danh sách popup cần tạo trên panel đó
        ''' </summary>
        Public Shared ReadOnly PendingPopups As New Dictionary(Of String, List(Of PopupDef))
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

        Public Shared Sub Register(
    controlDefs As Inventor.ControlDefinitions,
    addInClientID As String,
    buttonsList As System.Collections.Generic.List(Of ButtonDefinition),
    largeIcon As stdole.IPictureDisp,
    smallIcon As stdole.IPictureDisp)

            Dim assemblyFolder2 As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            ' Resolve icons folder (prefer user-configured folder if set)
            Dim configured As String = Nothing
            Try
                configured = My.Settings.ImageFolder
            Catch
                configured = Nothing
            End Try

            Dim iconsFolder As String = Nothing
            If Not String.IsNullOrWhiteSpace(configured) AndAlso System.IO.Directory.Exists(configured) Then
                ' If the user configured a folder, look for an "Assembly" subfolder there to keep compatibility
                iconsFolder = System.IO.Path.Combine(configured, "Assembly")
            Else
                iconsFolder = System.IO.Path.Combine(assemblyFolder2, "Code addin inventor", "Images", "Assembly")
            End If
#End Region

#Region "Load icons from folder"
            Dim Ass1LargePath1 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath1 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath2 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath2 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath3 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath3 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath4 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath4 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath5 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath5 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath6 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath6 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath7 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath7 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath8 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath8 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath9 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath9 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath10 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath10 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath11 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath11 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath12 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath12 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath13 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath13 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath14 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath14 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath15 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath15 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath16 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath16 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath17 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath17 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath18 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath18 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath19 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath19 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath20 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath20 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath21 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath21 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath22 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath22 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath23 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath23 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath24 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath24 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath25 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath25 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath26 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath26 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath27 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath27 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath28 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath28 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath29 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath29 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath30 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath30 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath31 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath31 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath32 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath32 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath33 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath33 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath34 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath34 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath35 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath35 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")
            Dim Ass1LargePath36 As String = System.IO.Path.Combine(iconsFolder, "i39.bmp")
            Dim Ass1SmallPath36 As String = System.IO.Path.Combine(iconsFolder, "i39 1.bmp")

            ' Load per-button icons (fallback to provided largeIcon/smallIcon when file missing)
            Dim ass1LargeIcon1 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath1), LoadIconFromPath(Ass1LargePath1), largeIcon)
            Dim ass1SmallIcon1 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath1), LoadIconFromPath(Ass1SmallPath1), smallIcon)
            Dim ass1LargeIcon2 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath2), LoadIconFromPath(Ass1LargePath2), largeIcon)
            Dim ass1SmallIcon2 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath2), LoadIconFromPath(Ass1SmallPath2), smallIcon)
            Dim ass1LargeIcon3 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath3), LoadIconFromPath(Ass1LargePath3), largeIcon)
            Dim ass1SmallIcon3 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath3), LoadIconFromPath(Ass1SmallPath3), smallIcon)
            Dim ass1LargeIcon4 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath4), LoadIconFromPath(Ass1LargePath4), largeIcon)
            Dim ass1SmallIcon4 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath4), LoadIconFromPath(Ass1SmallPath4), smallIcon)
            Dim ass1LargeIcon5 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath5), LoadIconFromPath(Ass1LargePath5), largeIcon)
            Dim ass1SmallIcon5 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath5), LoadIconFromPath(Ass1SmallPath5), smallIcon)
            Dim ass1LargeIcon6 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath6), LoadIconFromPath(Ass1LargePath6), largeIcon)
            Dim ass1SmallIcon6 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath6), LoadIconFromPath(Ass1SmallPath6), smallIcon)
            Dim ass1LargeIcon7 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath7), LoadIconFromPath(Ass1LargePath7), largeIcon)
            Dim ass1SmallIcon7 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath7), LoadIconFromPath(Ass1SmallPath7), smallIcon)
            Dim ass1LargeIcon8 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath8), LoadIconFromPath(Ass1LargePath8), largeIcon)
            Dim ass1SmallIcon8 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath8), LoadIconFromPath(Ass1SmallPath8), smallIcon)
            Dim ass1LargeIcon9 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath9), LoadIconFromPath(Ass1LargePath9), largeIcon)
            Dim ass1SmallIcon9 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath9), LoadIconFromPath(Ass1SmallPath9), smallIcon)
            Dim ass1LargeIcon10 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath10), LoadIconFromPath(Ass1LargePath10), largeIcon)
            Dim ass1SmallIcon10 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath10), LoadIconFromPath(Ass1SmallPath10), smallIcon)
            Dim ass1LargeIcon11 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath11), LoadIconFromPath(Ass1LargePath11), largeIcon)
            Dim ass1SmallIcon11 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath11), LoadIconFromPath(Ass1SmallPath11), smallIcon)
            Dim ass1LargeIcon12 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath12), LoadIconFromPath(Ass1LargePath12), largeIcon)
            Dim ass1SmallIcon12 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath12), LoadIconFromPath(Ass1SmallPath12), smallIcon)
            Dim ass1LargeIcon13 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath13), LoadIconFromPath(Ass1LargePath13), largeIcon)
            Dim ass1SmallIcon13 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath13), LoadIconFromPath(Ass1SmallPath13), smallIcon)
            Dim ass1LargeIcon14 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath14), LoadIconFromPath(Ass1LargePath14), largeIcon)
            Dim ass1SmallIcon14 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath14), LoadIconFromPath(Ass1SmallPath14), smallIcon)
            Dim ass1LargeIcon15 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath15), LoadIconFromPath(Ass1LargePath15), largeIcon)
            Dim ass1SmallIcon15 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath15), LoadIconFromPath(Ass1SmallPath15), smallIcon)
            Dim ass1LargeIcon16 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath16), LoadIconFromPath(Ass1LargePath16), largeIcon)
            Dim ass1SmallIcon16 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath16), LoadIconFromPath(Ass1SmallPath16), smallIcon)
            Dim ass1LargeIcon17 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath17), LoadIconFromPath(Ass1LargePath17), largeIcon)
            Dim ass1SmallIcon17 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath17), LoadIconFromPath(Ass1SmallPath17), smallIcon)
            Dim ass1LargeIcon18 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath18), LoadIconFromPath(Ass1LargePath18), largeIcon)
            Dim ass1SmallIcon18 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath18), LoadIconFromPath(Ass1SmallPath18), smallIcon)
            Dim ass1LargeIcon19 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath19), LoadIconFromPath(Ass1LargePath19), largeIcon)
            Dim ass1SmallIcon19 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath19), LoadIconFromPath(Ass1SmallPath19), smallIcon)
            Dim ass1LargeIcon20 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath20), LoadIconFromPath(Ass1LargePath20), largeIcon)
            Dim ass1SmallIcon20 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath20), LoadIconFromPath(Ass1SmallPath20), smallIcon)
            Dim ass1LargeIcon21 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath21), LoadIconFromPath(Ass1LargePath21), largeIcon)
            Dim ass1SmallIcon21 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath21), LoadIconFromPath(Ass1SmallPath21), smallIcon)
            Dim ass1LargeIcon22 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath22), LoadIconFromPath(Ass1LargePath22), largeIcon)
            Dim ass1SmallIcon22 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath22), LoadIconFromPath(Ass1SmallPath22), smallIcon)
            Dim ass1LargeIcon23 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath23), LoadIconFromPath(Ass1LargePath23), largeIcon)
            Dim ass1SmallIcon23 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath23), LoadIconFromPath(Ass1SmallPath23), smallIcon)
            Dim ass1LargeIcon24 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath24), LoadIconFromPath(Ass1LargePath24), largeIcon)
            Dim ass1SmallIcon24 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath24), LoadIconFromPath(Ass1SmallPath24), smallIcon)
            Dim ass1LargeIcon25 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath25), LoadIconFromPath(Ass1LargePath25), largeIcon)
            Dim ass1SmallIcon25 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath25), LoadIconFromPath(Ass1SmallPath25), smallIcon)
            Dim ass1LargeIcon26 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath26), LoadIconFromPath(Ass1LargePath26), largeIcon)
            Dim ass1SmallIcon26 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath26), LoadIconFromPath(Ass1SmallPath26), smallIcon)
            Dim ass1LargeIcon27 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath27), LoadIconFromPath(Ass1LargePath27), largeIcon)
            Dim ass1SmallIcon27 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath27), LoadIconFromPath(Ass1SmallPath27), smallIcon)
            Dim ass1LargeIcon28 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath28), LoadIconFromPath(Ass1LargePath28), largeIcon)
            Dim ass1SmallIcon28 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath28), LoadIconFromPath(Ass1SmallPath28), smallIcon)
            Dim ass1LargeIcon29 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath29), LoadIconFromPath(Ass1LargePath29), largeIcon)
            Dim ass1SmallIcon29 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath29), LoadIconFromPath(Ass1SmallPath29), smallIcon)
            Dim ass1LargeIcon30 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath30), LoadIconFromPath(Ass1LargePath30), largeIcon)
            Dim ass1SmallIcon30 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath30), LoadIconFromPath(Ass1SmallPath30), smallIcon)
            Dim ass1LargeIcon31 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath31), LoadIconFromPath(Ass1LargePath31), largeIcon)
            Dim ass1SmallIcon31 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath31), LoadIconFromPath(Ass1SmallPath31), smallIcon)
            Dim ass1LargeIcon32 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath32), LoadIconFromPath(Ass1LargePath32), largeIcon)
            Dim ass1SmallIcon32 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath32), LoadIconFromPath(Ass1SmallPath32), smallIcon)
            Dim ass1LargeIcon33 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath33), LoadIconFromPath(Ass1LargePath33), largeIcon)
            Dim ass1SmallIcon33 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath33), LoadIconFromPath(Ass1SmallPath33), smallIcon)
            Dim ass1LargeIcon34 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath34), LoadIconFromPath(Ass1LargePath34), largeIcon)
            Dim ass1SmallIcon34 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath34), LoadIconFromPath(Ass1SmallPath34), smallIcon)
            Dim ass1LargeIcon35 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1LargePath35), LoadIconFromPath(Ass1LargePath35), largeIcon)
            Dim ass1SmallIcon35 As stdole.IPictureDisp = If(System.IO.File.Exists(Ass1SmallPath35), LoadIconFromPath(Ass1SmallPath35), smallIcon)
#End Region

#Region "Nut cho các assembly"

            ' Create Assembly buttons explicitly (no loop) so each button can have distinct implementation

            '''' Lệnh lắp ghép ===============

#End Region
#Region "Lệnh lắp ghép"

#Region "constrain"
            '============================== Constrain các cụm chi tiết & part về gốc tọa độ của cụm chi tiết đầu tiên chọn ==============
            '  Dim assemblyBtn28 As ButtonDefinition = controlDefs.AddButtonDefinition("Constrain, Ground, Delete", "ToolInventor2025_Assembly_Btn28", CommandTypesEnum.kShapeEditCmdType, addInClientID,
            '                                                               Nothing, "Suppress,constrain,Ground" & vbLf & "Constrain Keep position" & vbLf & "Constrain về gốc 2 chi tiết" & vbLf & "Constrain all to select" & vbLf & "Xóa all Constrain lỗi",
            '                                                                ass1SmallIcon28, ass1LargeIcon28)
            '  AddHandler assemblyBtn28.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1.OnExecute
            ' buttonsList.Add(assemblyBtn28)



            '════════════════════════════════════════════════════════════════════
            ' NÚT 28 — CONSTRAIN với POPUP MENU
            ' Đăng ký popup vào registry, sẽ được tạo khi panel hình thành
            '════════════════════════════════════════════════════════════════════

            ' ─── 1. Tạo nút chính (dùng làm main button của popup) ───
            Dim subBtn1a As ButtonDefinition = controlDefs.AddButtonDefinition("Suppress, Constrain, Ground", "ToolInventor2025_Assembly_Sub1a",
                CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Suppress, constrain, Ground tự động", ass1SmallIcon1, ass1LargeIcon1)
            AddHandler subBtn1a.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1a.OnExecute

            ' ─── 2. Tạo các nút con ───
            Dim subBtn1b As ButtonDefinition = controlDefs.AddButtonDefinition("Constrain Keep Position", "ToolInventor2025_Assembly_Sub1b", CommandTypesEnum.kShapeEditCmdType,
                addInClientID, Nothing, "Giữ nguyên vị trí các cụm & gán constrain tự động", ass1SmallIcon2, ass1LargeIcon2)
            AddHandler subBtn1b.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1b.OnExecute

            Dim subBtn1c As ButtonDefinition = controlDefs.AddButtonDefinition("Constrain về gốc 2 chi tiết", "ToolInventor2025_Assembly_Sub1c", CommandTypesEnum.kShapeEditCmdType,
                addInClientID, Nothing, "Constrain về gốc tọa độ của chi tiết chọn đầu tiên", ass1SmallIcon3, ass1LargeIcon3)
            AddHandler subBtn1c.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1c.OnExecute

            Dim subBtn1d As ButtonDefinition = controlDefs.AddButtonDefinition("Constrain All to Select", "ToolInventor2025_Assembly_Sub1d", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                "Constrain tất cả về gốc tọa độ chi tiết được chọn", ass1SmallIcon4, ass1LargeIcon4)
            AddHandler subBtn1d.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1d.OnExecute

            Dim subBtn1e As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa all Constraint lỗi", "ToolInventor2025_Assembly_Sub1e", CommandTypesEnum.kShapeEditCmdType,
                addInClientID, Nothing, "Xóa tất cả constrain lỗi trong Assembly", ass1SmallIcon27, ass1LargeIcon27)
            AddHandler subBtn1e.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1e.OnExecute

            ' ─── 3. Đăng ký popup vào registry ───
            Dim pd As New PopupDef()
            pd.SubButtons.Add(subBtn1a)     ' ← PHẦN TỬ ĐẦU = nút chính
            pd.SubButtons.Add(subBtn1b)
            pd.SubButtons.Add(subBtn1c)
            pd.SubButtons.Add(subBtn1d)
            pd.SubButtons.Add(subBtn1e)

            If Not PendingPopups.ContainsKey("ToolInventor2025_AssemblyPanel") Then
                PendingPopups("ToolInventor2025_AssemblyPanel") = New List(Of PopupDef)
            End If
            PendingPopups("ToolInventor2025_AssemblyPanel").Add(pd)



            '════════════════════════════════════════════════════════════════════
            ' NÚT 28 — CONSTRAIN với POPUP MENU
            ' Đăng ký popup vào registry, sẽ được tạo khi panel hình thành
            '════════════════════════════════════════════════════════════════════

            ' ─── 1. Tạo nút chính (dùng làm main button của popup) ───
            Dim subBtn2a As ButtonDefinition = controlDefs.AddButtonDefinition("Suppress, Constrain, Ground", "ToolInventor2025_Assembly_Sub2a",
                CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Suppress, constrain, Ground tự động", ass1SmallIcon1, ass1LargeIcon1)
            AddHandler subBtn2a.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1a.OnExecute

            ' ─── 2. Tạo các nút con ───
            Dim subBtn2b As ButtonDefinition = controlDefs.AddButtonDefinition("Constrain Keep Position", "ToolInventor2025_Assembly_Sub2b", CommandTypesEnum.kShapeEditCmdType,
                addInClientID, Nothing, "Giữ nguyên vị trí các cụm & gán constrain tự động", ass1SmallIcon2, ass1LargeIcon2)
            AddHandler subBtn2b.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1b.OnExecute

            Dim subBtn2c As ButtonDefinition = controlDefs.AddButtonDefinition("Constrain về gốc 2 chi tiết", "ToolInventor2025_Assembly_Sub2c", CommandTypesEnum.kShapeEditCmdType,
                addInClientID, Nothing, "Constrain về gốc tọa độ của chi tiết chọn đầu tiên", ass1SmallIcon3, ass1LargeIcon3)
            AddHandler subBtn2c.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1c.OnExecute

            Dim subBtn2d As ButtonDefinition = controlDefs.AddButtonDefinition("Constrain All to Select", "ToolInventor2025_Assembly_Sub2d", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                "Constrain tất cả về gốc tọa độ chi tiết được chọn", ass1SmallIcon4, ass1LargeIcon4)
            AddHandler subBtn2d.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1d.OnExecute

            Dim subBtn2e As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa all Constraint lỗi", "ToolInventor2025_Assembly_Sub2e", CommandTypesEnum.kShapeEditCmdType,
                addInClientID, Nothing, "Xóa tất cả constrain lỗi trong Assembly", ass1SmallIcon27, ass1LargeIcon27)
            AddHandler subBtn2e.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1e.OnExecute

            ' ─── 3. Đăng ký popup vào registry ───
            Dim pd1 As New PopupDef()
            pd1.SubButtons.Add(subBtn2a)     ' ← PHẦN TỬ ĐẦU = nút chính
            pd1.SubButtons.Add(subBtn2b)
            pd1.SubButtons.Add(subBtn2c)
            pd1.SubButtons.Add(subBtn2d)
            pd1.SubButtons.Add(subBtn2e)

            If Not PendingPopups.ContainsKey("ToolInventor2025_AssemblyPanel") Then
                PendingPopups("ToolInventor2025_AssemblyPanel") = New List(Of PopupDef)
            End If
            PendingPopups("ToolInventor2025_AssemblyPanel").Add(pd1)
















            Dim assemblyBtn1 As ButtonDefinition = controlDefs.AddButtonDefinition("Suppress,constrain,Ground", "ToolInventor2025_Assembly_Btn1", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing,
                                                                                   Nothing, ass1SmallIcon1, ass1LargeIcon1)
            AddHandler assemblyBtn1.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1a.OnExecute
            ' buttonsList.Add(assemblyBtn1)

            Dim assemblyBtn2 As ButtonDefinition = controlDefs.AddButtonDefinition("Constrain Keep position", "ToolInventor2025_Assembly_Btn2", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing,
                                                                                   "Code giữ nguyên vị trí các cum & gán contrain tự động" & vbLf &
                                                                                  "Không áp dùng cho cụm hàn", ass1SmallIcon2, ass1LargeIcon2)
            AddHandler assemblyBtn2.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1b.OnExecute
            ' buttonsList.Add(assemblyBtn2)

            Dim assemblyBtn3 As ButtonDefinition = controlDefs.AddButtonDefinition("Constrain về gốc 2 chi tiết", "ToolInventor2025_Assembly_Btn3", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "Constrain cụm chi tiết hoặc part về gốc tọa độ của chi tiết hoặc cụm chi tiết đầu tiên chọn", ass1SmallIcon3, ass1LargeIcon3)
            AddHandler assemblyBtn3.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1c.OnExecute
            'buttonsList.Add(assemblyBtn3)

            Dim assemblyBtn4 As ButtonDefinition = controlDefs.AddButtonDefinition("Constrain all to select", "ToolInventor2025_Assembly_Btn4", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing,
                                                                                   "Constrain tất cả cụm chi tiết & part về gốc tọa độ của chi tiết hoặc cụm chi tiết được chọn", ass1SmallIcon4, ass1LargeIcon4)
            AddHandler assemblyBtn4.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1d.OnExecute
            ' buttonsList.Add(assemblyBtn4)

            Dim assemblyBtn27 As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa all Constrain lỗi", "ToolInventor2025_Assembly_Btn27", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                          Nothing, "Xóa tất cả các constrain lỗi trong Assembly.", ass1SmallIcon27, ass1LargeIcon27)
            AddHandler assemblyBtn27.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1e.OnExecute
            ' buttonsList.Add(assemblyBtn27)
            '==============================

            Dim assemblyBtn16 As ButtonDefinition = controlDefs.AddButtonDefinition("Hiện gốc tọa độ để Constrain", "ToolInventor2025_Assembly_Btn16", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                    Nothing, "ấn chọn file part hoặc Assembly để hiện gốc tọa độ" & vbCrLf & " ti ếp theo dùng lệnh Constrain để lắp ghép với nhau." & vbCrLf &
                                                                                    "Nếu ấn ok thì sẽ ẩn hết các mặt phẳng & trục gốc tọa độ vừa ấn hiện", ass1SmallIcon16, ass1LargeIcon16)
            AddHandler assemblyBtn16.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_2.OnExecute
            buttonsList.Add(assemblyBtn16)
#End Region




            Dim assemblyBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Tat ALL Adaptive cum LG", "ToolInventor2025_Assembly_Btn10", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ass1SmallIcon10, ass1LargeIcon10)
            AddHandler assemblyBtn10.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_LG_1.OnExecute
            buttonsList.Add(assemblyBtn10)

            Dim assemblyBtn21 As ButtonDefinition = controlDefs.AddButtonDefinition("UPDATE DESIGN STANDARD", "ToolInventor2025_Assembly_Btn21", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "Up date cho các tool tinh toán tiêu chuẩn ví dụ như buloong, key,...", ass1SmallIcon21, ass1LargeIcon21)
            AddHandler assemblyBtn21.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_LG_2.OnExecute
            buttonsList.Add(assemblyBtn21)

            Dim assemblyBtn18 As ButtonDefinition = controlDefs.AddButtonDefinition("Ản file", "ToolInventor2025_Assembly_Btn18", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                    Nothing, "Tạo, ghép shetmetal to assembly all lever lấy tất cả các tấm kể cả trung tên partnumber. mục 5,6 chưa ok", ass1SmallIcon18, ass1LargeIcon18)
            AddHandler assemblyBtn18.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_LG_3.OnExecute
            buttonsList.Add(assemblyBtn18)

#End Region
            '==================================================









#Region "Lệnh với Part"
            '========= Lệnh với Part ==============

            Dim assemblyBtn6 As ButtonDefinition = controlDefs.AddButtonDefinition("Covert to sheetmetal", "ToolInventor2025_Assembly_Btn6", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing,
                                                                                   "Chuyển part thành sheet metal", ass1SmallIcon6, ass1LargeIcon6)
            AddHandler assemblyBtn6.OnExecute, AddressOf Assembly.Buttons.Part.Ass_Part_1.OnExecute
            buttonsList.Add(assemblyBtn6)

            Dim assemblyBtn7 As ButtonDefinition = controlDefs.AddButtonDefinition("Combo part 1", "ToolInventor2025_Assembly_Btn7", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Thay đổi vật liệu, màu, đơn vị part, thông số part", ass1SmallIcon7, ass1LargeIcon7)
            AddHandler assemblyBtn7.OnExecute, AddressOf Assembly.Buttons.Part.Ass_Part_list_1.OnExecute
            buttonsList.Add(assemblyBtn7)

            Dim assemblyBtn8 As ButtonDefinition = controlDefs.AddButtonDefinition("Com", "ToolInventor2025_Assembly_Btn8", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Thay đổi vật liệu, màu, đơn vị part, thông số part", ass1SmallIcon8, ass1LargeIcon8)
            '     AddHandler assemblyBtn8.OnExecute, AddressOf Assembly.Buttons.Part.Ass_LG_SheetMetal.OnExecute
            '     buttonsList.Add(assemblyBtn8)

            Dim assemblyBtn9 As ButtonDefinition = controlDefs.AddButtonDefinition("Trải ALL Sheetmetal", "ToolInventor2025_Assembly_Btn9", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ass1SmallIcon9, ass1LargeIcon9)
            AddHandler assemblyBtn9.OnExecute, AddressOf Assembly.Buttons.Part.Ass_Part_4.OnExecute
            buttonsList.Add(assemblyBtn9)

            Dim assemblyBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Save copy to replace part", "ToolInventor2025_Assembly_Btn14", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ass1SmallIcon14, ass1LargeIcon14)
            AddHandler assemblyBtn14.OnExecute, AddressOf Assembly.Buttons.Part.Ass_part_5.OnExecute
            buttonsList.Add(assemblyBtn14)

            Dim assemblyBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Xoa mau ghi de len part", "ToolInventor2025_Assembly_Btn12", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ass1SmallIcon12, ass1LargeIcon12)
            '  AddHandler assemblyBtn12.OnExecute, AddressOf Assembly.Buttons.Part.Ass_Part_6.OnExecute
            '  buttonsList.Add(assemblyBtn12)

            Dim assemblyBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("An all plane part", "ToolInventor2025_Assembly_Btn11", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ass1SmallIcon11, ass1LargeIcon11)
            AddHandler assemblyBtn11.OnExecute, AddressOf Assembly.Buttons.Part.Ass_Part_7.OnExecute
            buttonsList.Add(assemblyBtn11)

            Dim assemblyBtn22 As ButtonDefinition = controlDefs.AddButtonDefinition("Thông số part", "ToolInventor2025_Assembly_Btn22", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "Hiển thị chi tiết thông số part", ass1SmallIcon22, ass1LargeIcon22)
            '  AddHandler assemblyBtn22.OnExecute, AddressOf Assembly.Buttons.Part.Ass_Part_8.OnExecute
            ' buttonsList.Add(assemblyBtn22)

            Dim assemblyBtn17 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay màu part", "ToolInventor2025_Assembly_Btn17", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                    Nothing, "Thay đổi màu sắc của chi tiết, thay vật liệu all", ass1SmallIcon17, ass1LargeIcon17)
            ' AddHandler assemblyBtn17.OnExecute, AddressOf Assembly.Buttons.Part.Ass_Part_9.OnExecute
            '  buttonsList.Add(assemblyBtn17)


#End Region
            '===============================



            Dim assemblyBtn20 As ButtonDefinition = controlDefs.AddButtonDefinition("5", "ToolInventor2025_Assembly_Btn20", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "", ass1SmallIcon20, ass1LargeIcon20)
            'AddHandler assemblyBtn20.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.ass_15.OnExecute
            ' buttonsList.Add(assemblyBtn20)





#Region "BTVT Assembly"
            '''''''''''''''' Bóc tách vật tư tấm, mua, tiêu chuẩn ======================================


            Dim assemblyBtn5 As ButtonDefinition = controlDefs.AddButtonDefinition("Lọc tấm, mua, Standard", "ToolInventor2025_Assembly_Btn5", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ass1SmallIcon5, ass1LargeIcon5)
            AddHandler assemblyBtn5.OnExecute, AddressOf Assembly.Buttons.caclenhboctach.part.Ass_boctach_part_1.OnExecute
            buttonsList.Add(assemblyBtn5)


            '======================================

#End Region

#Region "Auto Drawing"
            ''====================================== Drawing Auto ======================================

            Dim assemblyBtn23 As ButtonDefinition = controlDefs.AddButtonDefinition("Auto Drawing", "ToolInventor2025_Assembly_Btn23", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                 Nothing, "Auto tạo bản vẽ.", ass1SmallIcon23, ass1LargeIcon23)
            AddHandler assemblyBtn23.OnExecute, AddressOf Assembly.Buttons.AutoCreateDrawing.ASS_Auto_Drawing.OnExecute
            buttonsList.Add(assemblyBtn23)

            Dim assemblyBtn24 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing top ASS", "ToolInventor2025_Assembly_Btn24", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                            Nothing, "Auto drawing cho các Assembly chỉ áp dụng cho top lever", ass1SmallIcon24, ass1LargeIcon24)
            ' AddHandler assemblyBtn24.OnExecute, AddressOf Assembly.Buttons.AutoCreateDrawing.AutoDrawingASSTopLV.OnExecute
            '  buttonsList.Add(assemblyBtn24)


            Dim assemblyBtn25 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing top Ass, Part", "ToolInventor2025_Assembly_Btn25", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                            Nothing, "Auto drawing cho Assembly & Part Top lever.", ass1SmallIcon25, ass1LargeIcon25)
            '  AddHandler assemblyBtn25.OnExecute, AddressOf Assembly.Buttons.AutoCreateDrawing.AutoDrawingASSpartTopLV.OnExecute
            '  buttonsList.Add(assemblyBtn25)

            ''''''''''''''======================================

#End Region
#Region "Frame"
            Dim assemblyBtn26 As ButtonDefinition = controlDefs.AddButtonDefinition("Xem lỗi cắt Frame", "ToolInventor2025_Assembly_Btn26", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                          Nothing, "Auto drawing cho Top lever chỉ áp dụng cho các Assembly.", ass1SmallIcon26, ass1LargeIcon26)
            AddHandler assemblyBtn26.OnExecute, AddressOf Assembly.Buttons.Frame.Ass_Frame_1.OnExecute
            buttonsList.Add(assemblyBtn26)




#End Region

#Region "Lệnh ngoài cụm lắp"
            Dim assemblyBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Import,EX step & part", "ToolInventor2025_Assembly_Btn13", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                    Nothing,
                                                                                    "1, Import all file to part tự lưu, xóa liên kết lưu file tự động " & vbCrLf & "2, Export từ Cụm lắp sang file step" & vbCrLf &
                                                                                      "Có thể chọn nhiều file 1 lúc", ass1SmallIcon13, ass1LargeIcon13)
            AddHandler assemblyBtn13.OnExecute, AddressOf Assembly.Buttons.Lenhngoaicumlap.Im_EX_step_part.OnExecute
            buttonsList.Add(assemblyBtn13)



            Dim assemblyBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Design Assistant", "ToolInventor2025_Assembly_Btn15", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ass1SmallIcon15, ass1LargeIcon15)
            AddHandler assemblyBtn15.OnExecute, AddressOf Assembly.Buttons.Lenhngoaicumlap.Design_Assistant.OnExecute
            buttonsList.Add(assemblyBtn15)


            Dim assemblyBtn19 As ButtonDefinition = controlDefs.AddButtonDefinition("Mở nơi lưu File", "ToolInventor2025_Assembly_Btn19", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "Mở vị trí lưu file", ass1SmallIcon19, ass1LargeIcon19)
            AddHandler assemblyBtn19.OnExecute, AddressOf Toolngoai.Vitrifile.Vitrifile
            buttonsList.Add(assemblyBtn19)
#End Region



        End Sub
    End Class
End Namespace
