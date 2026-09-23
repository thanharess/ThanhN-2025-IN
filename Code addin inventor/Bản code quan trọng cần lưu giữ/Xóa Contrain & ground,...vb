Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.OLD


    Public Module OLD3

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try

                '==========================================================
                ' KIỂM TRA ACTIVE DOCUMENT
                '==========================================================

                Dim oDoc As Document =
                        g_inventorApplication.ActiveDocument

                If oDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then

                    MessageBox.Show(
                            "Hãy mở Assembly (.iam) trước khi chạy!",
                            "ThanhN",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)

                    Return
                End If

                Dim oAssDoc As AssemblyDocument =
                        CType(oDoc, AssemblyDocument)

                Dim oAsmCompDef As AssemblyComponentDefinition =
                        oAssDoc.ComponentDefinition


                '==========================================================
                ' FORM CHỌN CHỨC NĂNG
                '==========================================================

                Dim selectedAction As Integer = ShowMainMenu()

                If selectedAction = 0 Then
                    Return
                End If


                '==========================================================
                ' 1 - XÓA ALL CONSTRAINT + JOINT
                '==========================================================

                If selectedAction = 1 Then

                    If MessageBox.Show(
                            "Xóa tất cả Constraint và Joint?",
                            "ThanhN",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning) <> DialogResult.Yes Then

                        Return
                    End If

                    Dim constraintCount As Integer = 0
                    Dim jointCount As Integer = 0

                    DeleteAllConstraints(
                            oAsmCompDef,
                            constraintCount)

                    DeleteAllJoints(
                            oAsmCompDef,
                            jointCount)

                    MessageBox.Show(
                            "Đã xóa:" & vbCrLf &
                            "- " & constraintCount & " Constraints" & vbCrLf &
                            "- " & jointCount & " Joints",
                            "ThanhN",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

                    Return
                End If


                '==========================================================
                ' 2 - XÓA ALL CONSTRAINT
                '==========================================================

                If selectedAction = 2 Then

                    If MessageBox.Show(
                            "Xóa tất cả Constraint?",
                            "ThanhN",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning) <> DialogResult.Yes Then

                        Return
                    End If

                    Dim constraintCount As Integer = 0

                    DeleteAllConstraints(
                            oAsmCompDef,
                            constraintCount)

                    MessageBox.Show(
                            "Đã xóa " &
                            constraintCount &
                            " Constraints.",
                            "ThanhN",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

                    Return
                End If


                '==========================================================
                ' 3 - XÓA ALL JOINT
                '==========================================================

                If selectedAction = 3 Then

                    If MessageBox.Show(
                            "Xóa tất cả Joint?",
                            "ThanhN",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning) <> DialogResult.Yes Then

                        Return
                    End If

                    Dim jointCount As Integer = 0

                    DeleteAllJoints(
                            oAsmCompDef,
                            jointCount)

                    MessageBox.Show(
                            "Đã xóa " &
                            jointCount &
                            " Joints.",
                            "ThanhN",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

                    Return
                End If


                '==========================================================
                ' 4 - GROUND / UNGROUND ALL
                '==========================================================

                If selectedAction = 4 Then

                    If MessageBox.Show(
                            "Tiếp tục thay đổi Ground toàn bộ Assembly?",
                            "ThanhN",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning) <> DialogResult.Yes Then

                        Return
                    End If

                    Dim groundMode As Integer =
                            ShowGroundMenu()

                    If groundMode = 0 Then
                        Return
                    End If

                    Dim qGround As Boolean

                    If groundMode = 1 Then
                        qGround = True
                    Else
                        qGround = False
                    End If

                    Dim groundCount As Integer = 0

                    SetGroundRecursive(
                            oAsmCompDef.Occurrences,
                            qGround,
                            groundCount)

                    If qGround Then

                        MessageBox.Show(
                                "Đã Ground " &
                                groundCount &
                                " components.",
                                "ThanhN",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

                    Else

                        MessageBox.Show(
                                "Đã bỏ Ground " &
                                groundCount &
                                " components.",
                                "ThanhN",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

                    End If

                    Return
                End If


                '==========================================================
                ' 5 - XÓA CONSTRAINT + JOINT + GROUND
                '==========================================================

                If selectedAction = 5 Then

                    If MessageBox.Show(
                            "Xóa tất cả Constraint và Joint," &
                            vbCrLf &
                            "sau đó Ground/Unground toàn bộ Assembly?",
                            "ThanhN",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning) <> DialogResult.Yes Then

                        Return
                    End If


                    Dim constraintCount As Integer = 0
                    Dim jointCount As Integer = 0


                    '------------------------------------------------------
                    ' XÓA CONSTRAINT
                    '------------------------------------------------------

                    DeleteAllConstraints(
                            oAsmCompDef,
                            constraintCount)


                    '------------------------------------------------------
                    ' XÓA JOINT
                    '------------------------------------------------------

                    DeleteAllJoints(
                            oAsmCompDef,
                            jointCount)


                    MessageBox.Show(
                            "Đã xóa:" & vbCrLf &
                            "- " & constraintCount & " Constraints" & vbCrLf &
                            "- " & jointCount & " Joints",
                            "ThanhN",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)


                    '------------------------------------------------------
                    ' CHỌN GROUND
                    '------------------------------------------------------

                    Dim groundMode As Integer =
                            ShowGroundMenu()

                    If groundMode = 0 Then
                        Return
                    End If


                    Dim qGround As Boolean

                    If groundMode = 1 Then
                        qGround = True
                    Else
                        qGround = False
                    End If


                    Dim groundCount As Integer = 0

                    SetGroundRecursive(
                            oAsmCompDef.Occurrences,
                            qGround,
                            groundCount)


                    If qGround Then

                        MessageBox.Show(
                                "Hoàn tất!" & vbCrLf &
                                "Constraints: " & constraintCount & vbCrLf &
                                "Joints: " & jointCount & vbCrLf &
                                "Ground: " & groundCount,
                                "ThanhN",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

                    Else

                        MessageBox.Show(
                                "Hoàn tất!" & vbCrLf &
                                "Constraints: " & constraintCount & vbCrLf &
                                "Joints: " & jointCount & vbCrLf &
                                "Unground: " & groundCount,
                                "ThanhN",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

                    End If

                    Return

                End If


            Catch ex As Exception

                MessageBox.Show(
                        "Lỗi:" & vbCrLf &
                        ex.Message,
                        "ThanhN - Ground / Constraint",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

            End Try

        End Sub


        '==================================================================
        ' FORM MENU CHÍNH
        '==================================================================

        Private Function ShowMainMenu() As Integer

            Dim result As Integer = 0

            Using frm As New Form()

                frm.Text = "Xóa Constraint / Joint / Ground"
                frm.Width = 460
                frm.Height = 390
                frm.StartPosition =
                        FormStartPosition.CenterScreen

                frm.FormBorderStyle =
                        FormBorderStyle.FixedDialog

                frm.MaximizeBox = False
                frm.MinimizeBox = False


                Dim title As New Label()

                title.Text =
                        "CHỌN CHỨC NĂNG"

                title.Left = 20
                title.Top = 15
                title.Width = 400
                title.Height = 30

                'title.Font = New Drawing.Font("Arial", 12, Drawing.FontStyle.Bold)


                frm.Controls.Add(title)


                '----------------------------------------------------------
                ' BUTTON 1
                '----------------------------------------------------------

                Dim btn1 As New Button()

                btn1.Text =
                        "Xóa ALL Constraint + Joint"

                btn1.Left = 20
                btn1.Top = 55
                btn1.Width = 400
                btn1.Height = 42

                AddHandler btn1.Click,
                        Sub()
                            result = 1
                            frm.Close()
                        End Sub

                frm.Controls.Add(btn1)


                '----------------------------------------------------------
                ' BUTTON 2
                '----------------------------------------------------------

                Dim btn2 As New Button()

                btn2.Text =
                        "Xóa ALL Constraint"

                btn2.Left = 20
                btn2.Top = 105
                btn2.Width = 400
                btn2.Height = 42

                AddHandler btn2.Click,
                        Sub()
                            result = 2
                            frm.Close()
                        End Sub

                frm.Controls.Add(btn2)


                '----------------------------------------------------------
                ' BUTTON 3
                '----------------------------------------------------------

                Dim btn3 As New Button()

                btn3.Text =
                        "Xóa ALL Joint"

                btn3.Left = 20
                btn3.Top = 155
                btn3.Width = 400
                btn3.Height = 42

                AddHandler btn3.Click,
                        Sub()
                            result = 3
                            frm.Close()
                        End Sub

                frm.Controls.Add(btn3)


                '----------------------------------------------------------
                ' BUTTON 4
                '----------------------------------------------------------

                Dim btn4 As New Button()

                btn4.Text =
                        "Ground / Bỏ Ground ALL"

                btn4.Left = 20
                btn4.Top = 205
                btn4.Width = 400
                btn4.Height = 42

                AddHandler btn4.Click,
                        Sub()
                            result = 4
                            frm.Close()
                        End Sub

                frm.Controls.Add(btn4)


                '----------------------------------------------------------
                ' BUTTON 5
                '----------------------------------------------------------

                Dim btn5 As New Button()

                btn5.Text =
                        "Xóa Constraint + Joint + Ground"

                btn5.Left = 20
                btn5.Top = 255
                btn5.Width = 400
                btn5.Height = 42

                AddHandler btn5.Click,
                        Sub()
                            result = 5
                            frm.Close()
                        End Sub

                frm.Controls.Add(btn5)


                '----------------------------------------------------------
                ' CANCEL
                '----------------------------------------------------------

                Dim btnCancel As New Button()

                btnCancel.Text = "Không làm gì"

                btnCancel.Left = 20
                btnCancel.Top = 310
                btnCancel.Width = 400
                btnCancel.Height = 35

                AddHandler btnCancel.Click,
                        Sub()
                            result = 0
                            frm.Close()
                        End Sub

                frm.Controls.Add(btnCancel)


                frm.ShowDialog()

            End Using

            Return result

        End Function


        '==================================================================
        ' FORM GROUND
        '==================================================================

        Private Function ShowGroundMenu() As Integer

            Dim result As Integer = 0

            Using frm As New Form()

                frm.Text = "Kiểu Ground"
                frm.Width = 400
                frm.Height = 260
                frm.StartPosition =
                        FormStartPosition.CenterScreen

                frm.FormBorderStyle =
                        FormBorderStyle.FixedDialog

                frm.MaximizeBox = False
                frm.MinimizeBox = False


                Dim title As New Label()

                title.Text =
                        "CHỌN KIỂU GROUND"

                title.Left = 20
                title.Top = 15
                title.Width = 330
                title.Height = 30

                'title.Font = New Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)

                frm.Controls.Add(title)


                '----------------------------------------------------------
                ' GROUND ALL
                '----------------------------------------------------------

                Dim btnGround As New Button()

                btnGround.Text =
                        "GROUND ALL"

                btnGround.Left = 20
                btnGround.Top = 55
                btnGround.Width = 330
                btnGround.Height = 45

                AddHandler btnGround.Click,
                        Sub()
                            result = 1
                            frm.Close()
                        End Sub

                frm.Controls.Add(btnGround)


                '----------------------------------------------------------
                ' UNGROUND ALL
                '----------------------------------------------------------

                Dim btnUnGround As New Button()

                btnUnGround.Text =
                        "BỎ GROUND ALL"

                btnUnGround.Left = 20
                btnUnGround.Top = 110
                btnUnGround.Width = 330
                btnUnGround.Height = 45

                AddHandler btnUnGround.Click,
                        Sub()
                            result = 2
                            frm.Close()
                        End Sub

                frm.Controls.Add(btnUnGround)


                '----------------------------------------------------------
                ' CANCEL
                '----------------------------------------------------------

                Dim btnCancel As New Button()

                btnCancel.Text =
                        "KHÔNG LÀM GÌ"

                btnCancel.Left = 20
                btnCancel.Top = 165
                btnCancel.Width = 330
                btnCancel.Height = 35

                AddHandler btnCancel.Click,
                        Sub()
                            result = 0
                            frm.Close()
                        End Sub

                frm.Controls.Add(btnCancel)


                frm.ShowDialog()

            End Using

            Return result

        End Function


        '==================================================================
        ' DELETE ALL CONSTRAINT
        '==================================================================

        Private Sub DeleteAllConstraints(
                ByVal oAsmCompDef As AssemblyComponentDefinition,
                ByRef count As Integer)

            count = 0

            For i As Integer =
                    oAsmCompDef.Constraints.Count To 1 Step -1

                Try

                    oAsmCompDef.Constraints.Item(i).Delete()

                    count += 1

                Catch

                    'Bỏ qua Constraint không thể xóa

                End Try

            Next

        End Sub


        '==================================================================
        ' DELETE ALL JOINT
        '==================================================================

        Private Sub DeleteAllJoints(
                ByVal oAsmCompDef As AssemblyComponentDefinition,
                ByRef count As Integer)

            count = 0

            For i As Integer =
                    oAsmCompDef.Joints.Count To 1 Step -1

                Try

                    oAsmCompDef.Joints.Item(i).Delete()

                    count += 1

                Catch

                    'Bỏ qua Joint không thể xóa

                End Try

            Next

        End Sub


        '==================================================================
        ' GROUND / UNGROUND RECURSIVE
        '==================================================================

        Private Sub SetGroundRecursive(
                ByVal occurrences As ComponentOccurrences,
                ByVal groundValue As Boolean,
                ByRef count As Integer)

            For Each occurrence As ComponentOccurrence In occurrences

                Try

                    occurrence.Grounded = groundValue

                    count += 1

                Catch

                End Try


                '----------------------------------------------------------
                ' Nếu là Assembly thì đi tiếp xuống các cấp bên trong
                '----------------------------------------------------------

                Try

                    If occurrence.DefinitionDocumentType =
                            DocumentTypeEnum.kAssemblyDocumentObject Then

                        SetGroundRecursive(
                                occurrence.SubOccurrences,
                                groundValue,
                                count)

                    End If

                Catch

                End Try

            Next

        End Sub

    End Module

End Namespace
