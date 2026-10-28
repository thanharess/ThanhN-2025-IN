Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor

Namespace ToolInventor2025.Toolngoai

    Public Module Button_Assembly

        '══════════════════════════════════════════════════════
        ' CẤU HÌNH GIAO DIỆN
        '══════════════════════════════════════════════════════
        Private Const FORM_WIDTH As Integer = 760
        Private Const BTN_H As Integer = 34
        Private Const BTN_GAP As Integer = 5
        Private Const COLS As Integer = 2
        Private Const GB_PADDING As Integer = 12

        '══════════════════════════════════════════════════════
        ' DISPATCH: số thứ tự → hàm xử lý
        '══════════════════════════════════════════════════════
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case qShowMenu()

                ' ── Constrain ──────────────────────────────
                Case 1 : Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1a.OnExecute(Context)
                Case 2 : Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1b.OnExecute(Context)
                Case 3 : Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1c.OnExecute(Context)
                Case 4 : Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1d.OnExecute(Context)
                Case 5 : Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1e.OnExecute(Context)

                ' ── Assembly ───────────────────────────────
                Case 6 : Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_2.OnExecute(Context)
                Case 7 : Assembly.Buttons.caclenhlapghep.Ass_LG_1.OnExecute(Context)
                Case 8 : Assembly.Buttons.caclenhlapghep.Ass_LG_2.OnExecute(Context)
                Case 9 : Assembly.Buttons.caclenhlapghep.Ass_LG_3.OnExecute(Context)
                Case 10 : Assembly.Buttons.ASS_View.View_Representation.OnExecute(Context)

                ' ── Part ───────────────────────────────────
                Case 11 : Assembly.Buttons.Part.Ass_Part_1.OnExecute(Context)
                Case 12 : Assembly.Buttons.Part.Ass_Part_list_1.OnExecute(Context)
                Case 13 : Assembly.Buttons.Part.Ass_Part_4.OnExecute(Context)
                Case 14 : Assembly.Buttons.Part.Ass_part_5.OnExecute(Context)
                Case 15 : Assembly.Buttons.Part.Ass_Part_7.OnExecute(Context)

                ' ── Vật tư / Drawing / Frame ───────────────
                Case 16 : Assembly.Buttons.caclenhboctach.part.Ass_boctach_part_1.OnExecute(Context)
                Case 17 : Assembly.Buttons.AutoCreateDrawing.ASS_Auto_Drawing.OnExecute(Context)
                Case 18 : Assembly.Buttons.Frame.Ass_Frame_1.OnExecute(Context)

                ' ── Ngoài cụm lắp ──────────────────────────
                Case 19 : Assembly.Buttons.Lenhngoaicumlap.Im_EX_step_part.OnExecute(Context)
                Case 20 : Assembly.Buttons.Lenhngoaicumlap.Design_Assistant.OnExecute(Context)
                Case 21 : Toolngoai.Vitrifile.Vitrifile()
                Case 22 : Toolngoai.delete_link_OLE.OnExecute(Context)

            End Select
        End Sub

        '══════════════════════════════════════════════════════
        ' FORM MENU
        '══════════════════════════════════════════════════════
        Private Function qShowMenu() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Tool Inventor 2025  -  Assembly Tools"
                form.Width = FORM_WIDTH
                form.Height = 750
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.BackColor = System.Drawing.Color.FromArgb(245, 248, 252)
                form.Font = New Font("Segoe UI", 9.0!)
                form.Tag = 0

                ' ── Banner ─────────────────────────────
                Dim banner As New Label() With {
                    .Text = "⚙   TỔNG HỢP CÔNG CỤ ASSEMBLY",
                    .Left = 0, .Top = 0, .Width = FORM_WIDTH, .Height = 52,
                    .TextAlign = ContentAlignment.MiddleCenter,
                    .Font = New Font("Segoe UI", 13.0!, FontStyle.Bold),
                    .BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                    .ForeColor = System.Drawing.Color.White
                }
                form.Controls.Add(banner)

                Dim y As Integer = 62

                y = AddGroup(form, "🔧   Constrain", y, 1, New String() {
                    "Suppress, Constrain, Ground",
                    "Constrain Keep Position",
                    "Constrain về gốc 2 chi tiết",
                    "Constrain All to Select",
                    "Xóa all Constraint lỗi"
                })

                y = AddGroup(form, "📦   Assembly", y, 6, New String() {
                    "Hiện gốc tọa độ để Constrain",
                    "Tắt ALL Adaptive cụm LG",
                    "UPDATE DESIGN STANDARD",
                    "Ẩn file (sheetmetal)",
                    "View Representation"
                })

                y = AddGroup(form, "🧩   Part", y, 11, New String() {
                    "Convert to Sheetmetal",
                    "Combo Part (vật liệu, màu, đơn vị)",
                    "Trải ALL Sheetmetal",
                    "Save copy to replace part",
                    "Ẩn all plane part"
                })

                y = AddGroup(form, "📊   Vật tư / Drawing / Frame", y, 16, New String() {
                    "Lọc tấm, mua, Standard",
                    "Auto Drawing",
                    "Xem lỗi cắt Frame"
                })

                y = AddGroup(form, "🌐   Ngoài cụm lắp", y, 19, New String() {
                    "Import, Export step & part",
                    "Design Assistant",
                    "Mở nơi lưu File",
                    "Xóa Link OLE"
                })

                ' ── Nút HỦY ────────────────────────────
                Dim cancel As New Button() With {
                    .Text = "✕   HỦY",
                    .Left = 15, .Top = y + 3, .Width = FORM_WIDTH - 30, .Height = 38,
                    .BackColor = System.Drawing.Color.FromArgb(230, 70, 70),
                    .ForeColor = System.Drawing.Color.White,
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Segoe UI", 9.5!, FontStyle.Bold),
                    .Cursor = Cursors.Hand
                }
                cancel.FlatAppearance.BorderSize = 0
                AddHandler cancel.Click, Sub() form.Close()
                form.Controls.Add(cancel)

                form.Height = y + 75
                form.ShowDialog()
                result = CInt(form.Tag)
            End Using

            Return result
        End Function

        '══════════════════════════════════════════════════════
        ' HELPER: vẽ 1 GroupBox với các nút chia 2 cột
        '══════════════════════════════════════════════════════
        Private Function AddGroup(form As Form, title As String, top As Integer,
                                  startValue As Integer, texts As String()) As Integer
            Const GB_WIDTH As Integer = FORM_WIDTH - 30

            Dim rows As Integer = CInt(Math.Ceiling(texts.Length / COLS))
            Dim gbHeight As Integer = 22 + rows * (BTN_H + BTN_GAP) + 8

            Dim gb As New GroupBox() With {
                .Text = "  " & title,
                .Left = 15, .Top = top, .Width = GB_WIDTH, .Height = gbHeight,
                .Font = New Font("Segoe UI", 9.5!, FontStyle.Bold),
                .ForeColor = System.Drawing.Color.FromArgb(0, 100, 170),
                .BackColor = System.Drawing.Color.Transparent
            }
            form.Controls.Add(gb)

            Dim colWidth As Integer = CInt((GB_WIDTH - 2 * GB_PADDING - (COLS - 1) * 6) / COLS)

            For i As Integer = 0 To texts.Length - 1
                Dim col As Integer = i Mod COLS
                Dim row As Integer = i \ COLS
                Dim value As Integer = startValue + i

                Dim b As New Button() With {
                    .Text = value.ToString("00") & ".   " & texts(i),
                    .Left = GB_PADDING + col * (colWidth + 6),
                    .Top = 22 + row * (BTN_H + BTN_GAP),
                    .Width = colWidth,
                    .Height = BTN_H,
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .BackColor = System.Drawing.Color.White,
                    .ForeColor = System.Drawing.Color.FromArgb(30, 30, 30),
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Segoe UI", 9.0!, FontStyle.Regular),
                    .Cursor = Cursors.Hand,
                    .Padding = New Padding(8, 0, 0, 0)
                }
                b.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(200, 212, 224)
                b.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(217, 236, 255)
                b.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(180, 215, 250)
                AddHandler b.Click, Sub()
                                        form.Tag = value
                                        form.Close()
                                    End Sub
                gb.Controls.Add(b)
            Next

            Return top + gbHeight + 8
        End Function

    End Module
End Namespace