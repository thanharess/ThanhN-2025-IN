
Option Explicit On

Imports Inventor
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports IO = System.IO

Namespace ToolInventor2025.Assembly.Buttons.caclenhlapghep

    Public Module Ass_LG_2


        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication
            Dim asm As AssemblyDocument = Nothing

            Try
                asm = TryCast(app.ActiveEditDocument, AssemblyDocument)
                If asm Is Nothing Then
                    asm = TryCast(app.ActiveDocument, AssemblyDocument)
                End If

                If asm Is Nothing Then
                    MessageBox.Show("Vui lòng mở Assembly.", "Update Bolts",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                Dim autoCmd As ControlDefinition = Nothing
                Dim manualCmd As ControlDefinition = Nothing
                Dim updateCmd As ControlDefinition = Nothing

                Try : autoCmd = app.CommandManager.ControlDefinitions.Item("FDSolveAuto") : Catch : End Try
                Try : manualCmd = app.CommandManager.ControlDefinitions.Item("FDSolveManual") : Catch : End Try
                Try : updateCmd = app.CommandManager.ControlDefinitions.Item("AppUpdate") : Catch : End Try

                If autoCmd Is Nothing AndAlso manualCmd Is Nothing Then
                    MessageBox.Show("Không tìm thấy lệnh FDSolveAuto / FDSolveManual.",
                                        "Update Bolts", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim count As Integer = 0
                Dim processed As New System.Collections.Generic.HashSet(Of String)

                app.StatusBarText = "Đang Calculate Bolted Connections..."

                '====================================================
                ' 1. DUYỆT BROWSER (ĐỆ QUY)
                '====================================================
                Dim bp As BrowserPane = app.ActiveDocument.BrowserPanes.ActivePane
                If bp IsNot Nothing Then
                    ProcessNodes(bp.TopNode, asm, autoCmd, manualCmd, updateCmd, app, count, processed)
                End If

                '====================================================
                ' 2. DUYỆT OCCURRENCE (dự phòng)
                '====================================================
                For Each occ As ComponentOccurrence In asm.ComponentDefinition.Occurrences
                    Try
                        Dim nameU As String = occ.Name.ToUpperInvariant()

                        If IsBoltName(nameU) Then
                            Dim key As String = occ.Name
                            If processed.Contains(key) Then Continue For
                            processed.Add(key)

                            ForceSolveOccurrence(occ, asm, autoCmd, manualCmd, updateCmd, app)
                            count += 1
                        End If
                    Catch
                    End Try
                Next

                '====================================================
                ' 3. UPDATE CUỐI
                '====================================================
                asm.SelectSet.Clear()
                System.Windows.Forms.Application.DoEvents()

                Try
                    If updateCmd IsNot Nothing Then updateCmd.Execute()
                Catch
                End Try

                Try
                    asm.Update2(True)
                Catch
                    Try : asm.Update() : Catch : End Try
                End Try

                System.Windows.Forms.Application.DoEvents()
                app.StatusBarText = "Bolted Connections đã Calculate xong."

                MessageBox.Show(
                        "HOÀN TẤT!" & vbCrLf & vbCrLf &
                        "Đã Calculate: " & count.ToString() & " bolted connection(s)." & vbCrLf & vbCrLf &
                        "Bolt sẽ chuyển về vị trí lỗ mới.",
                        "Update Bolts",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

            Catch ex As Exception
                Try : app.ActiveDocument.SelectSet.Clear() : Catch : End Try
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                    "Update Bolts", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        '====================================================
        ' KIỂM TRA TÊN BOLT
        '====================================================
        Private Function IsBoltName(label As String) As Boolean
            If String.IsNullOrEmpty(label) Then Return False
            Return label.Contains("THEN") OrElse
                       label.Contains("BOLT") OrElse
                       label.Contains("BU LÔNG") OrElse
                       label.Contains("BULONG") OrElse
                       label.Contains("BOLTED")
        End Function

        '====================================================
        ' ĐỆ QUY BROWSER
        '====================================================
        Private Sub ProcessNodes(
                parent As BrowserNode,
                asm As AssemblyDocument,
                autoCmd As ControlDefinition,
                manualCmd As ControlDefinition,
                updateCmd As ControlDefinition,
                app As Inventor.Application,
                ByRef count As Integer,
                processed As System.Collections.Generic.HashSet(Of String))

            If parent Is Nothing Then Exit Sub

            For Each node As BrowserNode In parent.BrowserNodes
                Try
                    Dim label As String = ""
                    Try
                        label = node.BrowserNodeDefinition.Label.ToUpperInvariant()
                    Catch
                        GoTo NextChild
                    End Try

                    If IsBoltName(label) Then
                        Dim key As String = label
                        If Not processed.Contains(key) Then
                            processed.Add(key)

                            Try
                                If Not node.Expanded Then node.Expanded = True
                            Catch
                            End Try

                            ForceSolveNode(node, asm, autoCmd, manualCmd, updateCmd, app)
                            count += 1
                            app.StatusBarText = "Calculate: " & count.ToString() & " - " & label
                        End If
                    End If

NextChild:
                    ProcessNodes(node, asm, autoCmd, manualCmd, updateCmd, app, count, processed)
                Catch
                End Try
            Next
        End Sub

        '====================================================
        ' ÉP SOLVE 1 NODE
        '====================================================
        Private Sub ForceSolveNode(
                node As BrowserNode,
                asm As AssemblyDocument,
                autoCmd As ControlDefinition,
                manualCmd As ControlDefinition,
                updateCmd As ControlDefinition,
                app As Inventor.Application)

            Try
                asm.SelectSet.Clear()
                System.Windows.Forms.Application.DoEvents()

                If node.NativeObject IsNot Nothing Then
                    asm.SelectSet.Select(node.NativeObject)
                Else
                    node.DoSelect()
                End If

                System.Windows.Forms.Application.DoEvents()

                ' 1) Bật Auto (nếu có)
                If autoCmd IsNot Nothing Then
                    Try : autoCmd.Execute() : Catch : End Try
                    System.Windows.Forms.Application.DoEvents()
                End If

                ' 2) Manual Solve = Calculate (quan trọng)
                If manualCmd IsNot Nothing Then
                    Try : manualCmd.Execute() : Catch : End Try
                    System.Windows.Forms.Application.DoEvents()
                End If

                ' 3) Update
                Try
                    If updateCmd IsNot Nothing Then
                        updateCmd.Execute()
                    Else
                        asm.Update2(True)
                    End If
                Catch
                    Try : asm.Update() : Catch : End Try
                End Try

                System.Windows.Forms.Application.DoEvents()

            Catch
            End Try
        End Sub

        '====================================================
        ' ÉP SOLVE 1 OCCURRENCE
        '====================================================
        Private Sub ForceSolveOccurrence(
                occ As ComponentOccurrence,
                asm As AssemblyDocument,
                autoCmd As ControlDefinition,
                manualCmd As ControlDefinition,
                updateCmd As ControlDefinition,
                app As Inventor.Application)

            Try
                asm.SelectSet.Clear()
                System.Windows.Forms.Application.DoEvents()

                asm.SelectSet.Select(occ)
                System.Windows.Forms.Application.DoEvents()

                If autoCmd IsNot Nothing Then
                    Try : autoCmd.Execute() : Catch : End Try
                    System.Windows.Forms.Application.DoEvents()
                End If

                If manualCmd IsNot Nothing Then
                    Try : manualCmd.Execute() : Catch : End Try
                    System.Windows.Forms.Application.DoEvents()
                End If

                Try
                    If updateCmd IsNot Nothing Then
                        updateCmd.Execute()
                    Else
                        asm.Update2(True)
                    End If
                Catch
                End Try

                System.Windows.Forms.Application.DoEvents()

            Catch
            End Try
        End Sub

    End Module

End Namespace
