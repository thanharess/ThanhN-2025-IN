Option Explicit On
Option Strict Off

Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Toolngoai

    Public Class ShortcutKeys
        Private Shared m_form As HotkeyForm

        Public Shared Sub Start()
            If m_form IsNot Nothing AndAlso Not m_form.IsDisposed Then Exit Sub
            m_form = New HotkeyForm()
            m_form.Show()
            m_form.Hide()
        End Sub

        Public Shared Sub [Stop]()
            If m_form IsNot Nothing Then
                m_form.Close()
                m_form.Dispose()
                m_form = Nothing
            End If
        End Sub

        Private Class HotkeyForm
            Inherits Form

            Private Const WM_HOTKEY As Integer = &H312
            Private Const HOTKEY_PART As Integer = 9001
            Private Const HOTKEY_ASSY As Integer = 9002
            Private Const HOTKEY_BOM As Integer = 9003
            Private Const HOTKEY_DRAW As Integer = 9004
            Private Const MOD_SHIFT As UInteger = &H4

            <DllImport("user32.dll", SetLastError:=True)>
            Private Shared Function RegisterHotKey(hWnd As IntPtr, id As Integer, fsModifiers As UInteger, vk As UInteger) As Boolean
            End Function

            <DllImport("user32.dll", SetLastError:=True)>
            Private Shared Function UnregisterHotKey(hWnd As IntPtr, id As Integer) As Boolean
            End Function

            Public Sub New()
                Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
                Me.ShowInTaskbar = False
                Me.Opacity = 0
                Me.Size = New System.Drawing.Size(0, 0)
                Me.StartPosition = FormStartPosition.Manual
                Me.Location = New System.Drawing.Point(-2000, -2000)
            End Sub

            Protected Overrides Sub OnHandleCreated(e As EventArgs)
                MyBase.OnHandleCreated(e)
                RegisterHotKey(Me.Handle, HOTKEY_PART, MOD_SHIFT, CUInt(Keys.D1))
                RegisterHotKey(Me.Handle, HOTKEY_ASSY, MOD_SHIFT, CUInt(Keys.D2))
                RegisterHotKey(Me.Handle, HOTKEY_BOM, MOD_SHIFT, CUInt(Keys.D3))
                RegisterHotKey(Me.Handle, HOTKEY_DRAW, MOD_SHIFT, CUInt(Keys.D4))
            End Sub

            Protected Overrides Sub OnHandleDestroyed(e As EventArgs)
                Try
                    UnregisterHotKey(Me.Handle, HOTKEY_PART)
                    UnregisterHotKey(Me.Handle, HOTKEY_ASSY)
                    UnregisterHotKey(Me.Handle, HOTKEY_BOM)
                    UnregisterHotKey(Me.Handle, HOTKEY_DRAW)
                Catch
                End Try
                MyBase.OnHandleDestroyed(e)
            End Sub

            Protected Overrides Sub WndProc(ByRef m As Message)
                If m.Msg = WM_HOTKEY Then
                    Try
                        Select Case m.WParam.ToInt32()
                            Case HOTKEY_PART
                                Button_Part.OnExecute(Nothing)
                            Case HOTKEY_ASSY
                                Button_Assembly.OnExecute(Nothing)
                            Case HOTKEY_BOM
                                Button_Assembly2.OnExecute(Nothing)
                            Case HOTKEY_DRAW
                                Button_Drawing.OnExecute(Nothing)
                        End Select
                    Catch
                    End Try
                    m.Result = IntPtr.Zero
                    Return
                End If
                MyBase.WndProc(m)
            End Sub
        End Class
    End Class
End Namespace