Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Windows.Forms
Imports Inventor
Imports Microsoft.Win32

Namespace ToolInventor2025
    <ProgIdAttribute("ToolInventor2025.StandardAddInServer"),
    GuidAttribute("27785725-854b-490a-ac86-ab9dad7f3cc5")>
    Public Class Basecode
        Implements Inventor.ApplicationAddInServer

        ' AddIn client ID must be a managed String. Define it explicitly to avoid
        ' passing a COM object (System.__ComObject) where a String is expected.
        Private Const AddInClientID As String = "27785725-854b-490a-ac86-ab9dad7f3cc5"

        Private WithEvents m_uievents As UserInterfaceEvents
        Private m_partButtons As New System.Collections.Generic.List(Of ButtonDefinition)
        Private m_assemblyButtons As New System.Collections.Generic.List(Of ButtonDefinition)
        Private m_drawingButtons As New System.Collections.Generic.List(Of ButtonDefinition)
        Private m_assembly2Buttons As New System.Collections.Generic.List(Of ButtonDefinition)
        Private m_assembly3Buttons As New System.Collections.Generic.List(Of ButtonDefinition)
        '  Private WithEvents m_ToolButton As ButtonDefinition

#Region "ApplicationAddInServer Members"

        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  


        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for


        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' === Khởi tạo Context Menu ===
            ContextMenuManager.Initialize(g_inventorApplication, AddInClientID)

            ' ── Nút "Hiện Planes/Axes" — luôn hiện ──
            ' ── CHỈ đăng ký "Place Constraint" ──
            ' Nút "Hiện Planes/Axes" đã có sẵn trên Ribbon (Btn16 - "Hiện gốc tọa độ để Constrain")
            ' Không cần thêm vào context menu nữa.
            ContextMenuManager.RegisterButton(
                "Place Constraint",
                "ToolInventor2025_Context_Constraint",
                "Mở hộp thoại Place Constraint",
                AddressOf ContextMenuActions.RunPlaceConstraint,
                ContextEnv.Assembly,
                "",
                True)   ' onlyWhenPlanesVisible = True
            ' === Muốn thêm nút mới? Chỉ cần thêm 1 dòng như trên ===
            ' Connect to the user-interface events to handle a ribbon reset.
            m_uievents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents


            ' Create button definitions for Part, Assembly and Drawing: 15 buttons each.
            Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions

            ' Load icons for buttons
            Dim largeIcon As stdole.IPictureDisp '= Nothing
            Dim smallIcon As stdole.IPictureDisp ' = Nothing
            Try
                ' Resolve image folder relative to the add-in assembly so installer relocation is supported.
                ' Use the app's output directory so image files copied to output are found reliably
                Dim assemblyFolder As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)

                ' Allow an override folder (stored in My.Settings.ImageFolder). If set and exists,
                ' load images from that folder; otherwise fall back to the add-in assembly Images folder.
                Dim imagesFolder As String = Nothing
                Try
                    Dim configured As String = GetConfiguredImageFolder()

                    If Not String.IsNullOrWhiteSpace(configured) AndAlso System.IO.Directory.Exists(configured) Then
                        imagesFolder = configured
                    Else
                        imagesFolder = System.IO.Path.Combine(assemblyFolder, "Images", "Part")
                    End If
                Catch
                    imagesFolder = System.IO.Path.Combine(assemblyFolder, "Images", "Part")
                End Try

                Dim largePath As String = System.IO.Path.Combine(imagesFolder, "i3.bmp")
                Dim smallPath As String = System.IO.Path.Combine(imagesFolder, "i3 1.bmp")
                If System.IO.File.Exists(largePath) Then
                    largeIcon = LoadAndResizeIcon(largePath, 32, 32)
                Else
                    largeIcon = Nothing
                End If

                If System.IO.File.Exists(smallPath) Then
                    smallIcon = LoadAndResizeIcon(smallPath, 16, 16)
                Else
                    ' fall back to resized large icon when no small icon available
                    If largeIcon IsNot Nothing Then
                        smallIcon = LoadAndResizeIcon(largePath, 16, 16)
                    Else
                        smallIcon = Nothing
                    End If
                End If
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show("Lỗi load icon: " & ex.Message)
                largeIcon = Nothing
                smallIcon = Nothing
            End Try

            ' Create Part buttons via helper class (pass icons)
            PartButtons.Register(controlDefs, AddInClientID, m_partButtons, largeIcon, smallIcon)

            ' Create Assembly buttons via helper class
            AssemblyButtons.Register(controlDefs, AddInClientID, m_assemblyButtons, largeIcon, smallIcon)

            ' Create Drawing buttons via helper class
            DrawingButtons.Register(controlDefs, AddInClientID, m_drawingButtons, largeIcon, smallIcon)

            ' Create Assembly buttons via helper class
            Assembly2Buttons.Register(controlDefs, AddInClientID, m_assembly2Buttons, largeIcon, smallIcon)
            ' Create Assembly buttons via helper class
            Assembly3Buttons.Register(controlDefs, AddInClientID, m_assembly3Buttons, largeIcon, smallIcon)
            'm_ToolButton = controlDefs.AddButtonDefinition("Command Name", "Internal Name", CommandTypesEnum.kShapeEditCmdType, AddInClientID)
            'ToolButtons.Register(controlDefs, AddInClientID, m_ToolButtons, largeIcon, smallIcon)

            ' m_ToolButton = controlDefs.AddButtonDefinition("hhhhhh", "h Name", CommandTypesEnum.kShapeEditCmdType, AddInClientID,,, smallIcon, largeIcon)
            'ToolButtons.Register(controlDefs, AddInClientID, m_ToolButtons, largeIcon, smallIcon)

            ' Ensure the user interface is created when the add-in activates.
            ' Call AddToUserInterface unconditionally so the Assembly tab/panel is created
            ' even when firstTime is False (the ribbon can be reset at other times).
            AddToUserInterface()
        End Sub

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate
            ContextMenuManager.Shutdown()
            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            m_uievents = Nothing
            g_inventorApplication = Nothing

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        ' This property is provided to allow the AddIn to expose an API of its own to other 
        ' programs. Typically, this  would be done by implementing the AddIn's API
        ' interface in a class and returning that class object through this property.
        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ' Note:this method is now obsolete, you should use the 
        ' ControlDefinition functionality for implementing commands.
        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
        End Sub

#End Region

#Region "User interface definition"
        ' Sub where the user-interface creation is done.  This is called when
        ' the add-in loaded and also if the user interface is reset.

        ' Helper method to add a tab/panel/buttons to a ribbon.
        Private Sub AddTabPanelButtons(ribbonName As String, tabDisplayName As String, tabInternalName As String, panelDisplayName As String, panelInternalName As String, buttons As System.Collections.Generic.List(Of ButtonDefinition), Optional usePulldownOnly As Boolean = False)
            Try
                Dim ribbonObj As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item(ribbonName)
                Dim customTab As RibbonTab = Nothing
                Try
                    customTab = ribbonObj.RibbonTabs.Add(tabDisplayName, tabInternalName, AddInClientID)
                Catch
                    ' Tab likely exists - try to get it
                    Try
                        customTab = ribbonObj.RibbonTabs.Item(tabInternalName)
                    Catch
                        customTab = Nothing
                    End Try
                End Try

                If customTab IsNot Nothing Then
                    Dim customPanel As RibbonPanel = Nothing
                    Try
                        customPanel = customTab.RibbonPanels.Add(panelDisplayName, panelInternalName, AddInClientID)
                    Catch
                        Try
                            customPanel = customTab.RibbonPanels.Item(panelInternalName)
                        Catch
                            customPanel = Nothing
                        End Try
                    End Try
                    If customPanel IsNot Nothing Then

                        '══════════════════════════════════════════════════
                        ' ⭐ TẠO SPLIT BUTTON
                        '   coll[0] = nút chính — bấm trực tiếp chạy OnExecute
                        '   coll[1..] = các sub trong dropdown
                        '══════════════════════════════════════════════════
                        Try
                            If PopupRegistry.PendingPopups.ContainsKey(panelInternalName) Then
                                For Each pd In PopupRegistry.PendingPopups(panelInternalName)
                                    If pd Is Nothing OrElse pd.SubButtons Is Nothing OrElse pd.SubButtons.Count = 0 Then
                                        Continue For
                                    End If
                                    Dim mainBtn As ButtonDefinition = pd.SubButtons(0)
                                    Dim objColl As Inventor.ObjectCollection = g_inventorApplication.TransientObjects.CreateObjectCollection()
                                    For i As Integer = 1 To pd.SubButtons.Count - 1
                                        objColl.Add(pd.SubButtons(i))
                                    Next
                                    customPanel.CommandControls.AddSplitButton(mainBtn, objColl, False)
                                Next
                            End If
                        Catch popEx As Exception
                            ' Fallback: thêm nút đầu tiên dạng button thường
                            Try
                                If PopupRegistry.PendingPopups.ContainsKey(panelInternalName) Then
                                    For Each pd In PopupRegistry.PendingPopups(panelInternalName)
                                        If pd.SubButtons IsNot Nothing AndAlso pd.SubButtons.Count > 0 Then
                                            customPanel.CommandControls.AddButton(pd.SubButtons(0), True)
                                        End If
                                    Next
                                End If
                            Catch
                            End Try
                        End Try
                        '══════════════════════════════════════════════════
                        ' THÊM CÁC NÚT THƯỜNG
                        '══════════════════════════════════════════════════
                        If buttons IsNot Nothing Then
                            If Not usePulldownOnly Then
                                For Each bd As ButtonDefinition In buttons
                                    customPanel.CommandControls.AddButton(bd)
                                Next
                            Else
                                Try
                                    Dim pulldown As CommandControl = customPanel.CommandControls.AddPulldown(panelDisplayName, panelInternalName & "_Pulldown", AddInClientID)
                                    For Each bd As ButtonDefinition In buttons
                                        pulldown.Controls.AddButton(bd)
                                    Next
                                Catch pullex As Exception
                                    Try
                                        For Each bd As ButtonDefinition In buttons
                                            customPanel.CommandControls.AddButton(bd)
                                        Next
                                    Catch
                                    End Try
                                End Try
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                ' Show diagnostic so we can see why ribbon/pulldown creation failed at runtime
                Try
                    System.Windows.Forms.MessageBox.Show("AddTabPanelButtons error: " & ex.Message, "ToolInventor2025", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                Catch
                    ' ignore any failures showing the message
                End Try
            End Try
        End Sub

        Private Sub AddToUserInterface()
            ' This is where you'll add code to add buttons to the ribbon.

            '** Create a separate custom tab (not inside the Tools tab) for Assembly and Drawing.

            ' Use AddTabPanelButtons helper defined above to add tabs/panels/buttons to the ribbon.
            ' (The AddTabPanelButtons method is implemented as a Private Sub at class scope to allow Optional parameter.)
            '' Get the "Tools" tab.
            '' Get the part ribbon.

            ' Add to Assembly ribbon as a separate tab.
            AddTabPanelButtons("Assembly", "Tool Assembly", "ToolInventor2025_AssemblyTab", "Tool Assembly", "ToolInventor2025_AssemblyPanel", m_assemblyButtons, False)

            ' Add to Drawing ribbon as a separate tab.
            AddTabPanelButtons("Drawing", "Tool Drawing", "ToolInventor2025_DrawingTab", "Tool Drawing", "ToolInventor2025_DrawingPanel", m_drawingButtons, False)

            ' Add to Part ribbon as a separate tab.
            AddTabPanelButtons("Part", "Tool Part", "ToolInventor2025_PartTab", "Tool Part", "ToolInventor2025_PartPanel", m_partButtons, False)

            ' Add to Assembly ribbon as a separate tab.
            AddTabPanelButtons("Assembly", "Tool Bom", "ToolInventor2025_AssemblyTab2", "Tool Assembly 2", "ToolInventor2025_AssemblyPanel2", m_assembly2Buttons, False)
            ' Panel visible; add buttons directly to the panel (no pulldown)
            ' AddTabPanelButtons("Assembly", "Tool Bom", "ToolInventor2025_AssemblyTab3", "Tool Assembly 2", "ToolInventor2025_AssemblyPanel3", m_assembly3Buttons, False)


            '' Get the part ribbon.
            ' Dim partRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("Part")
            '' Get the "Tools" tab.
            ' Dim toolsTab As RibbonTab = partRibbon.RibbonTabs.Item("id_TabTools")
            '' Create a new panel.
            '  Dim customPanel As RibbonPanel = toolsTab.RibbonPanels.Add("jjj", "MysSampleq", AddInClientID)
            '' Add a button.
            ' customPanel.CommandControls.AddButton(m_ToolButton)



            '' Get the part ribbon.
            ' Dim assemblyRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("Assembly")
            '' Get the "Tools" tab.
            ' Dim toolsTab1 As RibbonTab = assemblyRibbon.RibbonTabs.Item("id_TabTools")
            '' Create a new panel.
            'Dim customPanel1 As RibbonPanel = toolsTab1.RibbonPanels.Add("jjj", "MysSampleq", AddInClientID)
            '' Add a button.
            ' customPanel1.CommandControls.AddButton(m_ToolButton)



            '' Get the part ribbon.
            ' Dim DrawingRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("Drawing")
            '' Get the "Tools" tab.
            ' Dim toolsTab2 As RibbonTab = DrawingRibbon.RibbonTabs.Item("id_TabTools")
            '' Create a new panel.
            'Dim customPanel2 As RibbonPanel = toolsTab2.RibbonPanels.Add("jjj", "MysSampleq", AddInClientID)
            '' Add a button.
            'customPanel2.CommandControls.AddButton(m_ToolButton)

        End Sub

        Private Sub Muievents_onresetribbonInterface(Context As NameValueMap) Handles m_uievents.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub


        ' Helpers to load images from disk and convert to IPictureDisp for Inventor
        ' AxHost.GetIPictureDispFromPicture is Protected, so expose it via a small derived helper class.
        Private NotInheritable Class AxHostHelper
            Inherits System.Windows.Forms.AxHost

            Private Sub New()
                MyBase.New(String.Empty)
            End Sub

            Public Shared Function GetIPictureDispFromImage(img As Image) As stdole.IPictureDisp
                ' Call the protected Shared method from the derived class context
                Return CType(GetIPictureDispFromPicture(img), stdole.IPictureDisp)
            End Function
        End Class

        Private Function PictureDispFromBitmap(bmp As Bitmap) As stdole.IPictureDisp
            Return AxHostHelper.GetIPictureDispFromImage(bmp)
        End Function

        Private Function LoadIconFromFile(path As String) As stdole.IPictureDisp
            If String.IsNullOrEmpty(path) Then Return Nothing
            If Not System.IO.File.Exists(path) Then Return Nothing
            Using bmp As New Bitmap(path)
                Dim clone As New Bitmap(bmp)
                Try
                    ' Use the Inventor sample converter which calls OleCreatePictureIndirect.
                    Return PictureDispConverter.ToIPictureDisp(clone)
                Finally
                    ' Dispose managed clone; OleCreatePictureIndirect takes ownership of the underlying HBITMAP
                    clone.Dispose()
                End Try
            End Using
        End Function

        Private Function LoadAndResizeIcon(path As String, width As Integer, height As Integer) As stdole.IPictureDisp
            If String.IsNullOrEmpty(path) Then Return Nothing
            If Not System.IO.File.Exists(path) Then Return Nothing

            Using src As New Bitmap(path)
                Using resized As New Bitmap(width, height)
                    Using g As Graphics = Graphics.FromImage(resized)
                        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                        g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                        g.Clear(System.Drawing.Color.Transparent)
                        g.DrawImage(src, 0, 0, width, height)
                    End Using

                    Dim clone As New Bitmap(resized)
                    Try
                        Return PictureDispConverter.ToIPictureDisp(clone)
                    Finally
                        clone.Dispose()
                    End Try
                End Using
            End Using
        End Function

        ' Button dispatcher methods removed: handlers wired directly to specific Button classes

#End Region

    End Class
End Namespace


Public Module Globals
    ' Inventor application object.
    Public g_inventorApplication As Inventor.Application

#Region "Function to get the add-in client ID."
    ' This function uses reflection to get the GuidAttribute associated with the add-in.
    Public Function AddInClientID() As String
        Dim guid As String = ""
        Try
            Dim t As Type = GetType(ToolInventor2025.Basecode)
            Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
            Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
            guid = "{" + guidAttribute.Value.ToString() + "}"
        Catch
        End Try

        Return guid
    End Function
#End Region

#Region "hWnd Wrapper Class"
    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ' This is primarily used for parenting a dialog to the Inventor window.
    '
    ' For example:
    ' myForm.Show(New WindowWrapper(g_inventorApplication.MainFrameHWND))
    '
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

        Private _hwnd As IntPtr
    End Class
#End Region

#Region "Image Converter"
    ' Class used to convert bitmaps and icons from their .Net native types into
    ' an IPictureDisp object which is what the Inventor API requires. A typical
    ' usage is shown below where MyIcon is a bitmap or icon that's available
    ' as a resource of the project.
    '
    ' Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.MyIcon)

    Public NotInheritable Class PictureDispConverter

        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)>
        Private Shared Function OleCreatePictureIndirect(
    ByVal picdesc As PICTDESC.Bitmap,
    ByRef iid As System.Guid,
    <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As IPictureDisp
        End Function

        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)>
        Private Shared Function OleCreatePictureIndirect(
    ByVal picdesc As PICTDESC.Icon,
    ByRef iid As System.Guid,
    <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As IPictureDisp
        End Function


        Shared iPictureDispGuid As System.Guid = GetType(stdole.IPictureDisp).GUID

        Private NotInheritable Class PICTDESC
            Private Sub New()
            End Sub

            'Picture Types
            Public Const PICTYPE_BITMAP As Short = 1
            Public Const PICTYPE_ICON As Short = 3

            <StructLayout(LayoutKind.Sequential)>
            Public Class Icon
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Icon))
                Friend picType As Integer = PICTDESC.PICTYPE_ICON
                Friend hicon As IntPtr = IntPtr.Zero
                Friend unused1 As Integer
                Friend unused2 As Integer

                Friend Sub New(ByVal icon As System.Drawing.Icon)
                    Me.hicon = icon.ToBitmap().GetHicon()
                End Sub
            End Class

            <StructLayout(LayoutKind.Sequential)>
            Public Class Bitmap
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
                Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
                Friend hbitmap As IntPtr = IntPtr.Zero
                Friend hpal As IntPtr = IntPtr.Zero
                Friend unused As Integer

                Friend Sub New(ByVal bitmap As System.Drawing.Bitmap)
                    Me.hbitmap = bitmap.GetHbitmap()
                End Sub
            End Class
        End Class

        Public Shared Function ToIPictureDisp(ByVal icon As System.Drawing.Icon) As stdole.IPictureDisp
            Dim pictIcon As New PICTDESC.Icon(icon)
            Return OleCreatePictureIndirect(pictIcon, iPictureDispGuid, True)
        End Function

        Public Shared Function ToIPictureDisp(ByVal bmp As System.Drawing.Bitmap) As stdole.IPictureDisp
            Dim pictBmp As New PICTDESC.Bitmap(bmp)
            Return OleCreatePictureIndirect(pictBmp, iPictureDispGuid, True)
        End Function
    End Class
#End Region

End Module
