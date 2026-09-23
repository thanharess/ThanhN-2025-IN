Imports Inventor
Imports System.Windows.Forms

Namespace ToolInventor2025

    ''' <summary>
    ''' Chứa các hàm xử lý cho Context Menu.
    ''' </summary>
    Public Module ContextMenuActions

        ' =========================================================
        ' 1. Chạy lệnh Constrain gốc của Inventor
        ' =========================================================
        Public Sub RunPlaceConstraint(ByVal Context As NameValueMap)
            Try
                Dim app = ToolInventor2025.Globals.g_inventorApplication
                If app Is Nothing Then Return

                Dim cmd As ControlDefinition = FindConstrainCommand()
                If cmd Is Nothing Then
                    MessageBox.Show("Không tìm thấy lệnh Constrain trong Inventor.",
                                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                ' ── Chỉ cần chạy lệnh Constrain ─────────────────────
                ' Việc ẩn Planes/Axes sẽ do sự kiện OnTerminateCommand
                ' trong module Ass_LG_C_2 tự xử lý (đã có sẵn).
                cmd.Execute()

            Catch ex As Exception
                MessageBox.Show("Lỗi chạy lệnh Constrain: " & ex.Message,
                                "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Function FindConstrainCommand() As ControlDefinition
            Try
                Dim app = ToolInventor2025.Globals.g_inventorApplication
                If app Is Nothing Then Return Nothing

                Dim controlDefs = app.CommandManager.ControlDefinitions

                ' ✅ Tên chính xác của nút "Constrain" trong Inventor 2025
                ' (đã xác định từ danh sách 2459 commands)
                Dim namesToTry() As String = {
                    "AssemblyInsertConstraintCmd",   ' ← tên chính xác Inventor 2025
                    "AssemblyConstraintCmd",          ' ← fallback các bản cũ
                    "AssemblyConstrainCmd",           ' ← fallback
                    "Constraint"
                }

                For Each n As String In namesToTry
                    Try
                        Return controlDefs.Item(n)
                    Catch
                    End Try
                Next

            Catch
            End Try

            Return Nothing
        End Function
    End Module
End Namespace