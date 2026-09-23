Imports System.Collections.Generic
Imports Inventor

Namespace ToolInventor2025

    ' =========================================================
    ' ENUM MÔI TRƯỜNG
    ' =========================================================
    <Flags>
    Public Enum ContextEnv
        None = 0
        Part = 1
        Assembly = 2
        Drawing = 4
        All = Part Or Assembly Or Drawing
    End Enum

    ''' <summary>
    ''' Quản lý các nút tùy chỉnh trên menu chuột phải.
    ''' </summary>
    Public NotInheritable Class ContextMenuManager

        Private Sub New()
        End Sub

        ' ---------------------------------------------------------
        ' BIẾN CHUNG
        ' ---------------------------------------------------------
        Private Shared _app As Inventor.Application
        Private Shared _clientId As String
        Private Shared _userInputEvents As UserInputEvents
        Private Shared _initialized As Boolean = False

        Private Shared ReadOnly _registeredButtons As New List(Of ContextMenuButtonInfo)
        Private Shared ReadOnly _buttonDefCache As New Dictionary(Of String, ButtonDefinition)

        ' =========================================================
        ' CLASS LƯU THÔNG TIN NÚT
        ' =========================================================
        Private Class ContextMenuButtonInfo
            Public Property DisplayName As String
            Public Property InternalName As String
            Public Property Tooltip As String
            Public Property Handler As ButtonDefinitionSink_OnExecuteEventHandler
            Public Property Environment As ContextEnv
            Public Property GroupName As String
            Public Property OnlyWhenPlanesVisible As Boolean   ' ← MỚI
        End Class

        ' =========================================================
        ' KHỞI TẠO
        ' =========================================================
        Public Shared Sub Initialize(app As Inventor.Application, clientId As String)
            If app Is Nothing Then Return

            _app = app
            _clientId = clientId
            _userInputEvents = app.CommandManager.UserInputEvents

            AddHandler _userInputEvents.OnContextMenu, AddressOf OnContextMenu

            _initialized = True
        End Sub

        ' =========================================================
        ' DỌN DẸP
        ' =========================================================
        Public Shared Sub Shutdown()
            Try
                If _userInputEvents IsNot Nothing Then
                    RemoveHandler _userInputEvents.OnContextMenu, AddressOf OnContextMenu
                End If
            Catch
            End Try

            _userInputEvents = Nothing
            _app = Nothing
            _buttonDefCache.Clear()
            _registeredButtons.Clear()
            _initialized = False
        End Sub

        ' =========================================================
        ' ĐĂNG KÝ NÚT
        '
        ' displayName            : Tên hiển thị
        ' internalName           : Tên nội bộ (unique)
        ' tooltip                : Tooltip
        ' handler                : Sub xử lý khi bấm
        ' environment            : Môi trường (Part/Assembly/Drawing/All)
        ' groupName              : (tùy chọn) Tên submenu — rỗng = nút trực tiếp
        ' onlyWhenPlanesVisible  : (tùy chọn) True = chỉ hiện khi planes đang bật
        ' =========================================================
        Public Shared Sub RegisterButton(
            displayName As String,
            internalName As String,
            tooltip As String,
            handler As ButtonDefinitionSink_OnExecuteEventHandler,
            Optional environment As ContextEnv = ContextEnv.All,
            Optional groupName As String = "",
            Optional onlyWhenPlanesVisible As Boolean = False)

            If String.IsNullOrEmpty(internalName) Then Return
            If handler Is Nothing Then Return

            For Each b In _registeredButtons
                If b.InternalName = internalName Then Return
            Next

            _registeredButtons.Add(New ContextMenuButtonInfo With {
                .DisplayName = displayName,
                .InternalName = internalName,
                .Tooltip = tooltip,
                .Handler = handler,
                .Environment = environment,
                .GroupName = groupName,
                .OnlyWhenPlanesVisible = onlyWhenPlanesVisible
            })
        End Sub

        ' =========================================================
        ' XỬ LÝ SỰ KIỆN CHUỘT PHẢI
        ' =========================================================
        Private Shared Sub OnContextMenu(
            ByVal SelectionDevice As SelectionDeviceEnum,
            ByVal AdditionalInfo As NameValueMap,
            ByVal CommandBar As CommandBar)

            Try
                If Not _initialized OrElse _app Is Nothing Then Return
                If _app.ActiveDocument Is Nothing Then Return
                If CommandBar Is Nothing Then Return

                Dim currentEnv As ContextEnv = GetCurrentEnvironment(_app.ActiveDocument.DocumentType)
                If currentEnv = ContextEnv.None Then Return

                Dim controlDefs As ControlDefinitions = _app.CommandManager.ControlDefinitions

                ' ⭐ Kiểm tra planes có đang bật không (1 lần cho cả vòng lặp)
                Dim planesVisible As Boolean = False
                Try
                    planesVisible = ToolInventor2025.Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_2.IsWorkFeaturesVisible
                Catch
                End Try

                ' ── Nhóm các nút ──
                Dim normalButtons As New List(Of ContextMenuButtonInfo)
                Dim groupedButtons As New Dictionary(Of String, List(Of ContextMenuButtonInfo))

                For Each info In _registeredButtons
                    If (info.Environment And currentEnv) = 0 Then Continue For

                    ' ⭐ Bỏ qua nút nếu yêu cầu planes đang bật mà planes chưa bật
                    If info.OnlyWhenPlanesVisible AndAlso Not planesVisible Then Continue For

                    If String.IsNullOrEmpty(info.GroupName) Then
                        normalButtons.Add(info)
                    Else
                        If Not groupedButtons.ContainsKey(info.GroupName) Then
                            groupedButtons(info.GroupName) = New List(Of ContextMenuButtonInfo)
                        End If
                        groupedButtons(info.GroupName).Add(info)
                    End If
                Next

                ' ── 1. Nút bình thường ──
                For Each info In normalButtons
                    Dim btnDef As ButtonDefinition = GetOrCreateButtonDef(controlDefs, info)
                    If btnDef Is Nothing Then Continue For
                    Try : CommandBar.Controls.AddButton(btnDef) : Catch : End Try
                Next

                ' ── 2. Submenu ──
                For Each kvp In groupedButtons
                    Dim groupName As String = kvp.Key
                    Dim groupList As List(Of ContextMenuButtonInfo) = kvp.Value

                    If groupList.Count = 0 Then Continue For

                    ' Nếu submenu chỉ có 1 nút → thêm trực tiếp
                    If groupList.Count = 1 Then
                        Dim btnDef As ButtonDefinition = GetOrCreateButtonDef(controlDefs, groupList(0))
                        If btnDef IsNot Nothing Then
                            Try : CommandBar.Controls.AddButton(btnDef) : Catch : End Try
                        End If
                        Continue For
                    End If

                    Try
                        Dim groupInternalName As String = "ToolInventor2025_CtxGroup_" &
                                          groupName.Replace(" ", "_")

                        ' ⭐ Late binding để tránh warning COM overload
                        Dim popupObj As Object = CommandBar.CommandControls
                        Dim popupCtrl As Object = popupObj.AddPopup(
            groupName, groupInternalName, _clientId)
                        Dim popupControls As Object = popupCtrl.Controls

                        For Each info In groupList
                            Dim btnDef As ButtonDefinition = GetOrCreateButtonDef(controlDefs, info)
                            If btnDef Is Nothing Then Continue For
                            Try : popupControls.AddButton(btnDef) : Catch : End Try
                        Next

                    Catch
                        ' Fallback: thêm trực tiếp vào menu chính
                        For Each info In groupList
                            Dim btnDef As ButtonDefinition = GetOrCreateButtonDef(controlDefs, info)
                            If btnDef Is Nothing Then Continue For
                            Try : CommandBar.Controls.AddButton(btnDef) : Catch : End Try
                        Next
                    End Try
                Next

            Catch
            End Try
        End Sub

        ' =========================================================
        ' LẤY HOẶC TẠO ButtonDefinition
        ' =========================================================
        Private Shared Function GetOrCreateButtonDef(
            controlDefs As ControlDefinitions,
            info As ContextMenuButtonInfo) As ButtonDefinition

            If _buttonDefCache.ContainsKey(info.InternalName) Then
                Return _buttonDefCache(info.InternalName)
            End If

            Dim btnDef As ButtonDefinition = Nothing
            Try
                btnDef = controlDefs.Item(info.InternalName)
            Catch
                btnDef = Nothing
            End Try

            If btnDef Is Nothing Then
                Try
                    btnDef = controlDefs.AddButtonDefinition(
                        info.DisplayName,
                        info.InternalName,
                        CommandTypesEnum.kShapeEditCmdType,
                        _clientId,
                        Nothing,
                        info.Tooltip)
                    AddHandler btnDef.OnExecute, info.Handler
                Catch
                    Return Nothing
                End Try
            End If

            _buttonDefCache(info.InternalName) = btnDef
            Return btnDef
        End Function

        ' =========================================================
        ' XÁC ĐỊNH MÔI TRƯỜNG HIỆN TẠI
        ' =========================================================
        Private Shared Function GetCurrentEnvironment(docType As DocumentTypeEnum) As ContextEnv
            Select Case docType
                Case DocumentTypeEnum.kPartDocumentObject
                    Return ContextEnv.Part
                Case DocumentTypeEnum.kAssemblyDocumentObject
                    Return ContextEnv.Assembly
                Case DocumentTypeEnum.kDrawingDocumentObject
                    Return ContextEnv.Drawing
                Case Else
                    Return ContextEnv.None
            End Select
        End Function

    End Class

End Namespace