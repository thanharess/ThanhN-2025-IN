Imports System.Collections.Generic
Imports Inventor

Namespace ToolInventor2025

    ' =========================================================
    ' ENUM MÔI TRƯỜNG (đặt ở namespace level cho dễ gọi)
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
    ''' Quản lý các nút tùy chỉnh trên menu chuột phải (Context Menu).
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

        ' ---------------------------------------------------------
        ' DANH SÁCH NÚT ĐÃ ĐĂNG KÝ
        ' ---------------------------------------------------------
        Private Shared ReadOnly _registeredButtons As New List(Of ContextMenuButtonInfo)

        ' ---------------------------------------------------------
        ' CACHE ButtonDefinition
        ' ---------------------------------------------------------
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
        End Class

        ' =========================================================
        ' KHỞI TẠO (GỌI TRONG Activate)
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
        ' DỌN DẸP (GỌI TRONG Deactivate)
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
        ' ĐĂNG KÝ MỘT NÚT MỚI
        ' =========================================================
        Public Shared Sub RegisterButton(
            displayName As String,
            internalName As String,
            tooltip As String,
            handler As ButtonDefinitionSink_OnExecuteEventHandler,
            Optional environment As ContextEnv = ContextEnv.All)

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
                .Environment = environment
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

                For Each info In _registeredButtons
                    If (info.Environment And currentEnv) = 0 Then Continue For

                    Dim btnDef As ButtonDefinition = GetOrCreateButtonDef(controlDefs, info)
                    If btnDef Is Nothing Then Continue For

                    Try
                        CommandBar.Controls.AddButton(btnDef)
                    Catch
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