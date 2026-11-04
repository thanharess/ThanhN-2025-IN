Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Drawing
Imports System.Collections.Generic
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
        ' CLASS MÔ TẢ 1 NÚT VÀ 1 NHÓM
        '══════════════════════════════════════════════════════
        Private Class MenuItemDef
            Public Text As String
            Public Action As Action(Of NameValueMap)
            Public Sub New(t As String, a As Action(Of NameValueMap))
                Text = t
                Action = a
            End Sub
        End Class

        Private Class MenuGroupDef
            Public Title As String
            Public Items As New List(Of MenuItemDef)
            Public Sub New(t As String)
                Title = t
            End Sub
        End Class

        '══════════════════════════════════════════════════════
        ' ⭐⭐⭐  CHỈ CẦN SỬA HÀM NÀY KHI THÊM / XÓA / ĐỔI NÚT  ⭐⭐⭐
        '
        '  • Thêm nút vào nhóm có sẵn:
        '       gX.Items.Add(New MenuItemDef("Tên nút", Sub(c) Namespace.Class.OnExecute(c)))
        '
        '  • Hàm không nhận Context:
        '       gX.Items.Add(New MenuItemDef("Tên nút", Sub(c) Namespace.Class.Run()))
        '
        '  • Thêm nhóm mới:
        '       Dim g6 As New MenuGroupDef("🎯   Tiêu đề")
        '       g6.Items.Add(New MenuItemDef("Nút 1", Sub(c) ...))
        '       groups.Add(g6)
        '
        '  • Đổi thứ tự nút  → đổi thứ tự các dòng .Items.Add()
        '  • Đổi thứ tự nhóm → đổi thứ tự các dòng groups.Add()
        '  • Xóa nút/nhóm    → xóa dòng .Items.Add() / groups.Add() tương ứng
        '
        '  Số thứ tự (01, 02, 03...) trên nút là TỰ ĐỘNG – không cần
        '  tự đánh số. Thêm / xóa / di chuyển vẫn đúng thứ tự.
        '══════════════════════════════════════════════════════
        Private Function BuildMenu() As List(Of MenuGroupDef)
            Dim groups As New List(Of MenuGroupDef)

            '── 🔧 Constrain ────────────────────────────────
            Dim g1 As New MenuGroupDef("🔧   Constrain")
            g1.Items.Add(New MenuItemDef("Suppress, Constrain, Ground",
                Sub(c) Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1a.OnExecute(c)))
            g1.Items.Add(New MenuItemDef("Constrain Keep Position",
                Sub(c) Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1b.OnExecute(c)))
            g1.Items.Add(New MenuItemDef("Constrain về gốc 2 chi tiết",
                Sub(c) Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1c.OnExecute(c)))
            g1.Items.Add(New MenuItemDef("Constrain All to Select",
                Sub(c) Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1d.OnExecute(c)))
            g1.Items.Add(New MenuItemDef("Xóa all Constraint lỗi",
                Sub(c) Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1e.OnExecute(c)))
            groups.Add(g1)

            '── 📦 Assembly ─────────────────────────────────
            Dim g2 As New MenuGroupDef("📦   Assembly")
            g2.Items.Add(New MenuItemDef("Hiện gốc tọa độ để Constrain",
                Sub(c) Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_2.OnExecute(c)))
            g2.Items.Add(New MenuItemDef("Tắt ALL Adaptive cụm LG",
                Sub(c) Assembly.Buttons.caclenhlapghep.Ass_LG_1.OnExecute(c)))
            g2.Items.Add(New MenuItemDef("UPDATE DESIGN STANDARD",
                Sub(c) Assembly.Buttons.caclenhlapghep.Ass_LG_2.OnExecute(c)))
            g2.Items.Add(New MenuItemDef("Ẩn file (sheetmetal)",
                Sub(c) Assembly.Buttons.caclenhlapghep.Ass_LG_3.OnExecute(c)))
            g2.Items.Add(New MenuItemDef("View Representation",
                Sub(c) Assembly.Buttons.ASS_View.View_Representation.OnExecute(c)))
            g2.Items.Add(New MenuItemDef("Replace Component",
                Sub(c) Assembly.Buttons.caclenhlapghep.Ass_Part_Replace.OnExecute(c)))
            groups.Add(g2)

            '── 🧩 Part ─────────────────────────────────────
            Dim g3 As New MenuGroupDef("🧩   Part")
            g3.Items.Add(New MenuItemDef("Convert to Sheetmetal",
                Sub(c) Assembly.Buttons.Part.Ass_Part_1.OnExecute(c)))
            g3.Items.Add(New MenuItemDef("Combo Part (vật liệu, màu, đơn vị)",
                Sub(c) Assembly.Buttons.Part.Ass_Part_list_1.OnExecute(c)))
            g3.Items.Add(New MenuItemDef("Trải ALL Sheetmetal",
                Sub(c) Assembly.Buttons.Part.Ass_Part_4.OnExecute(c)))
            g3.Items.Add(New MenuItemDef("Save copy to replace part",
                Sub(c) Assembly.Buttons.Part.Ass_part_5.OnExecute(c)))
            g3.Items.Add(New MenuItemDef("Ẩn all plane part",
                Sub(c) Assembly.Buttons.Part.Ass_Part_7.OnExecute(c)))
            groups.Add(g3)

            '── 📊 Vật tư / Drawing / Frame ─────────────────
            Dim g4 As New MenuGroupDef("📊   Vật tư / Drawing / Frame")
            g4.Items.Add(New MenuItemDef("Lọc tấm, mua, Standard",
                Sub(c) Assembly.Buttons.caclenhboctach.part.Ass_boctach_part_1.OnExecute(c)))
            g4.Items.Add(New MenuItemDef("Auto Drawing",
                Sub(c) Assembly.Buttons.AutoCreateDrawing.ASS_Auto_Drawing.OnExecute(c)))
            g4.Items.Add(New MenuItemDef("Xem lỗi cắt Frame",
                Sub(c) Assembly.Buttons.Frame.Ass_Frame_1.OnExecute(c)))
            g4.Items.Add(New MenuItemDef("Xem lỗi cắt Frame (chọn loại)",
                Sub(c) Assembly.Buttons.Frame.Ass_Frame_1B.OnExecute(c)))
            groups.Add(g4)

            '── 🌐 Ngoài cụm lắp ───────────────────────────
            Dim g5 As New MenuGroupDef("🌐   Ngoài cụm lắp")
            g5.Items.Add(New MenuItemDef("Import, Export step & part",
                Sub(c) Assembly.Buttons.Lenhngoaicumlap.Im_EX_step_part.OnExecute(c)))
            g5.Items.Add(New MenuItemDef("Design Assistant",
                Sub(c) Assembly.Buttons.Lenhngoaicumlap.Design_Assistant.OnExecute(c)))
            g5.Items.Add(New MenuItemDef("Mở nơi lưu File",
                Sub(c) Toolngoai.Vitrifile.Vitrifile()))
            g5.Items.Add(New MenuItemDef("Xóa Link OLE",
                Sub(c) Toolngoai.delete_link_OLE.OnExecute(c)))
            groups.Add(g5)

            Return groups
        End Function

        '══════════════════════════════════════════════════════
        ' DISPATCH – không cần Select Case nữa
        '══════════════════════════════════════════════════════
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim action As Action(Of NameValueMap) = qShowMenu()
            If action IsNot Nothing Then
                action(Context)
            End If
        End Sub

        '══════════════════════════════════════════════════════
        ' FORM MENU
        '══════════════════════════════════════════════════════
        Private Function qShowMenu() As Action(Of NameValueMap)
            Dim result As Action(Of NameValueMap) = Nothing
            Dim groups As List(Of MenuGroupDef) = BuildMenu()

            Using form As New Form()
                form.Text = "Tool Inventor 2025  -  Assembly Tools"
                form.Width = FORM_WIDTH
                form.Height = 780
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.BackColor = System.Drawing.Color.FromArgb(245, 248, 252)
                form.Font = New Font("Segoe UI", 9.0!)
                form.Tag = Nothing

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
                Dim stt As Integer = 1     ' số thứ tự tự tăng

                For Each g As MenuGroupDef In groups
                    y = AddGroup(form, g, y, stt)
                Next

                ' ── Nút HỦY ────────────────────────────
                Dim cancel As New Button() With {
                    .Text = "✕   HỦY",
                    .Left = 15, .Top = y, .Width = FORM_WIDTH - 45, .Height = 38,
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

                If form.Tag IsNot Nothing Then
                    result = CType(form.Tag, Action(Of NameValueMap))
                End If
            End Using

            Return result
        End Function

        '══════════════════════════════════════════════════════
        ' HELPER: vẽ 1 GroupBox với các nút chia 2 cột
        '══════════════════════════════════════════════════════
        Private Function AddGroup(form As Form, group As MenuGroupDef,
                                  top As Integer, ByRef startStt As Integer) As Integer
            Const GB_WIDTH As Integer = FORM_WIDTH - 30

            Dim n As Integer = group.Items.Count
            Dim rows As Integer = CInt(Math.Ceiling(n / COLS))
            Dim gbHeight As Integer = 22 + rows * (BTN_H + BTN_GAP) + 8

            Dim gb As New GroupBox() With {
                .Text = "  " & group.Title,
                .Left = 15, .Top = top, .Width = GB_WIDTH, .Height = gbHeight,
                .Font = New Font("Segoe UI", 9.5!, FontStyle.Bold),
                .ForeColor = System.Drawing.Color.FromArgb(0, 100, 170),
                .BackColor = System.Drawing.Color.Transparent
            }
            form.Controls.Add(gb)

            Dim colWidth As Integer = CInt((GB_WIDTH - 2 * GB_PADDING - (COLS - 1) * 6) / COLS)

            For i As Integer = 0 To n - 1
                Dim item As MenuItemDef = group.Items(i)
                Dim col As Integer = i Mod COLS
                Dim row As Integer = i \ COLS
                Dim value As Integer = startStt
                startStt += 1

                Dim b As New Button() With {
                    .Text = value.ToString("00") & ".   " & item.Text,
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

                Dim capturedAction As Action(Of NameValueMap) = item.Action
                AddHandler b.Click, Sub()
                                        form.Tag = capturedAction
                                        form.Close()
                                    End Sub
                gb.Controls.Add(b)
            Next

            Return top + gbHeight + 8
        End Function

    End Module
End Namespace