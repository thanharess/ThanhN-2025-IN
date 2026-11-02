Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor

Namespace ToolInventor2025.Toolngoai

    Public Module Button_Drawing

        Private Const FORM_WIDTH As Integer = 760
        Private Const BTN_H As Integer = 34
        Private Const BTN_GAP As Integer = 5
        Private Const COLS As Integer = 2
        Private Const GB_PADDING As Integer = 12

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowMenua()
                Case 1 : Drawing.Buttons.Draw_1.OnExecute(Context)
                Case 2 : Drawing.Buttons.Drawdim.draw_2.OnExecute(Context)
                Case 3 : Drawing.Buttons.Drawdim.draw_3.OnExecute(Context)
                Case 4 : Drawing.Buttons.Draw_4.OnExecute(Context)
                Case 5 : Drawing.Buttons.Drawdim.Dim_Linear_Total.OnExecute(Context)
                Case 6 : Drawing.Buttons.DrawPartList.Create_PartsList.OnExecute(Context)
                Case 7 : Drawing.Buttons.DrawPartList.Draw_Partlist_1.OnExecute(Context)
                Case 8 : Drawing.Buttons.DrawSheet.Drawing_PartsList_Sync.OnExecute(Context)
                Case 9 : Drawing.Buttons.Draw_ResetPartsList.OnExecute(Context)
                Case 10 : Drawing.Buttons.DrawView.Draw_Up_date_view.OnExecute(Context)
                Case 11 : Drawing.Buttons.DrawView.Draw_Change_scale_view.OnExecute(Context)
                Case 12 : Drawing.Buttons.DrawView.Draw_Rename_View.OnExecute(Context)
                Case 13 : Drawing.Buttons.DrawView.Draw_AutoBalloon.OnExecute(Context)
                Case 14 : Drawing.Buttons.Draw_7.OnExecute(Context)
                Case 15 : Drawing.Buttons.Draw_Rename_Sheet.OnExecute(Context)
                Case 16 : Drawing.Buttons.DrawSheet.Drawing_TitleBlock_Border.OnExecute(Context)
                Case 17 : Drawing.Buttons.Style.Drw_Style_Import.OnExecute(Context)
                Case 18 : Drawing.Buttons.Drawtext.draw_text_list_1.OnExecute(Context)
                Case 19 : Drawing.Buttons.Drawdelete.Draw_delete.OnExecute(Context)
                Case 20 : Drawing.Buttons.Draw_9.OnExecute(Context)
                Case 21 : Drawing.Buttons.Draw_5.OnExecute(Context)
                Case 22 : Drawing.Buttons.draw_14.OnExecute(Context)
                Case 23 : Toolngoai.Vitrifile.Vitrifile()
            End Select
        End Sub

        Private Function ShowMenua() As Integer
            Dim result As Integer = 0
            Using form As New Form()
                form.Text = "Tool Inventor 2025  -  Drawing Tools"
                form.Width = FORM_WIDTH
                form.Height = 980
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.BackColor = System.Drawing.Color.FromArgb(245, 248, 252)
                form.Font = New Font("Segoe UI", 9.0!)
                form.Tag = 0

                Dim banner As New Label() With {
                    .Text = "📐   TỔNG HỢP CÔNG CỤ DRAWING",
                    .Left = 0, .Top = 0, .Width = FORM_WIDTH, .Height = 52,
                    .TextAlign = ContentAlignment.MiddleCenter,
                    .Font = New Font("Segoe UI", 13.0!, FontStyle.Bold),
                    .BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                    .ForeColor = System.Drawing.Color.White
                }
                form.Controls.Add(banner)

                Dim y As Integer = 62
                y = AddGroup(form, "📏   Dimension", y, 1, New String() {
                    "Sửa số thập phân Dim", "Auto Dim (line, lỗ, ren)", "Xóa, căn chỉnh Dim", "Tìm Dim bị edit", "Dim kích thước bao"})
                y = AddGroup(form, "📋   Partlist", y, 6, New String() {
                    "Tạo Partlist", "Ghi Partlist ENG / VIE", "Đồng bộ Partlist", "Reset Partlist"})
                y = AddGroup(form, "👁   View", y, 10, New String() {
                    "Update Views", "Đổi scale View", "Thay chữ cái View", "Auto Balloon"})
                y = AddGroup(form, "📄   Sheet", y, 14, New String() {
                    "Chuyển Sheet", "Rename Sheet", "Thay khung tên, Border", "Import Drawing Style"})
                y = AddGroup(form, "✍   Text", y, 18, New String() {
                    "Text Edit", "Delete Text"})
                y = AddGroup(form, "🧹   Cleanup", y, 20, New String() {
                    "Xóa Part / ASS trùng lặp", "Xóa Centermark, Centerline"})
                y = AddGroup(form, "⚙   Khác", y, 22, New String() {
                    "Đổi scale Hatch", "Mở vị trí lưu File"})

                Dim cancel As New Button() With {
                    .Text = "✕   HỦY", .Left = 15, .Top = y + 3, .Width = FORM_WIDTH - 30, .Height = 38,
                    .BackColor = System.Drawing.Color.FromArgb(230, 70, 70), .ForeColor = System.Drawing.Color.White,
                    .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 9.5!, FontStyle.Bold), .Cursor = Cursors.Hand}
                cancel.FlatAppearance.BorderSize = 0
                AddHandler cancel.Click, Sub() form.Close()
                form.Controls.Add(cancel)

                form.Height = y + 75
                form.ShowDialog()
                result = CInt(form.Tag)
            End Using
            Return result
        End Function

        Private Function AddGroup(form As Form, title As String, top As Integer,
                                  startValue As Integer, texts As String()) As Integer
            Const GB_WIDTH As Integer = FORM_WIDTH - 30
            Dim Cols3 As Integer = If(texts.Length = 1, 1, COLS)
            Dim rows As Integer = CInt(Math.Ceiling(texts.Length / CDbl(Cols3)))
            Dim gbHeight As Integer = 22 + rows * (BTN_H + BTN_GAP) + 8

            Dim gb As New GroupBox() With {
                .Text = "  " & title, .Left = 15, .Top = top, .Width = GB_WIDTH, .Height = gbHeight,
                .Font = New Font("Segoe UI", 9.5!, FontStyle.Bold),
                .ForeColor = System.Drawing.Color.FromArgb(0, 100, 170), .BackColor = System.Drawing.Color.Transparent}
            form.Controls.Add(gb)

            Dim colWidth As Integer = CInt((GB_WIDTH - 2 * GB_PADDING - (Cols3 - 1) * 6) / Cols3)
            For i As Integer = 0 To texts.Length - 1
                Dim col As Integer = i Mod Cols3
                Dim row As Integer = i \ Cols3
                Dim value As Integer = startValue + i
                Dim b As New Button() With {
                    .Text = value.ToString("00") & ".   " & texts(i),
                    .Left = GB_PADDING + col * (colWidth + 6), .Top = 22 + row * (BTN_H + BTN_GAP),
                    .Width = colWidth, .Height = BTN_H, .TextAlign = ContentAlignment.MiddleLeft,
                    .BackColor = System.Drawing.Color.White, .ForeColor = System.Drawing.Color.FromArgb(30, 30, 30),
                    .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 9.0!), .Cursor = Cursors.Hand,
                    .Padding = New Padding(8, 0, 0, 0)}
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