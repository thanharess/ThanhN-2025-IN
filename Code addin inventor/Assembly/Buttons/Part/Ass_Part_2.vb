Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Part
    Public Module Ass_Part_2


        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                SetDocumentUnits()

            Catch ex As Exception

                MessageBox.Show(
                    "Lỗi: " & ex.Message,
                    "Đơn vị File",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub


        Private Sub SetDocumentUnits()

            '==========================================================
            ' Danh sách lựa chọn
            '==========================================================
            Dim unitOptions As New List(Of String)

            ' --- Hệ mét (mm / cm / m - kg) ---
            unitOptions.Add("All Part Thành mm - kg")
            unitOptions.Add("All Assembly Thành mm - kg")
            unitOptions.Add("All Assembly và Part Thành mm - kg")
            unitOptions.Add("All Assembly và Part Thành cm - kg")
            unitOptions.Add("All Assembly và Part Thành m - kg")

            unitOptions.Add("")

            ' --- Hệ inch (in - lb) ---
            unitOptions.Add("All Part Thành inch - lb")
            unitOptions.Add("All Assembly Thành inch - lb")
            unitOptions.Add("All Assembly và Part Thành inch - lb")
            unitOptions.Add("All Assembly và Part Thành feet - lb")


            '==========================================================
            ' Hiện Form chọn
            '==========================================================
            Dim selectedOption As String = ShowUnitSelectionForm(unitOptions)

            If String.IsNullOrEmpty(selectedOption) Then Return


            '==========================================================
            ' Active Document
            '==========================================================
            Dim openDoc As Document = g_inventorApplication.ActiveDocument

            If openDoc Is Nothing Then
                MessageBox.Show("Không có Document đang mở.", "Đơn vị File",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If


            '==========================================================
            ' Biến thiết lập
            '==========================================================
            Dim oUOM1 As UnitsTypeEnum
            Dim oUOM2 As UnitsTypeEnum
            Dim oPrecision As Integer = 3
            Dim applyToPart As Boolean = False
            Dim applyToAssembly As Boolean = False
            Dim applyToActive As Boolean = True


            '==========================================================
            ' Phân tích lựa chọn
            '==========================================================
            Select Case selectedOption

        '----------------- HỆ MÉT -----------------
                Case "All Part Thành mm - kg"
                    oUOM1 = UnitsTypeEnum.kMillimeterLengthUnits
                    oUOM2 = UnitsTypeEnum.kKilogramMassUnits
                    oPrecision = 3
                    applyToPart = True
                    applyToAssembly = False

                Case "All Assembly Thành mm - kg"
                    oUOM1 = UnitsTypeEnum.kMillimeterLengthUnits
                    oUOM2 = UnitsTypeEnum.kKilogramMassUnits
                    oPrecision = 3
                    applyToPart = False
                    applyToAssembly = True

                Case "All Assembly và Part Thành mm - kg"
                    oUOM1 = UnitsTypeEnum.kMillimeterLengthUnits
                    oUOM2 = UnitsTypeEnum.kKilogramMassUnits
                    oPrecision = 3
                    applyToPart = True
                    applyToAssembly = True

                Case "All Assembly và Part Thành cm - kg"
                    oUOM1 = UnitsTypeEnum.kCentimeterLengthUnits
                    oUOM2 = UnitsTypeEnum.kKilogramMassUnits
                    oPrecision = 3
                    applyToPart = True
                    applyToAssembly = True

                Case "All Assembly và Part Thành m - kg"
                    oUOM1 = UnitsTypeEnum.kMeterLengthUnits
                    oUOM2 = UnitsTypeEnum.kKilogramMassUnits
                    oPrecision = 3
                    applyToPart = True
                    applyToAssembly = True

            '----------------- HỆ INCH -----------------
                Case "All Part Thành inch - lb"
                    oUOM1 = UnitsTypeEnum.kInchLengthUnits
                    oUOM2 = UnitsTypeEnum.kLbMassMassUnits
                    oPrecision = 4                       ' inch cần 4 chữ số thập phân
                    applyToPart = True
                    applyToAssembly = False

                Case "All Assembly Thành inch - lb"
                    oUOM1 = UnitsTypeEnum.kInchLengthUnits
                    oUOM2 = UnitsTypeEnum.kLbMassMassUnits
                    oPrecision = 4
                    applyToPart = False
                    applyToAssembly = True

                Case "All Assembly và Part Thành inch - lb"
                    oUOM1 = UnitsTypeEnum.kInchLengthUnits
                    oUOM2 = UnitsTypeEnum.kLbMassMassUnits
                    oPrecision = 4
                    applyToPart = True
                    applyToAssembly = True

                Case "All Assembly và Part Thành feet - lb"
                    oUOM1 = UnitsTypeEnum.kFootLengthUnits
                    oUOM2 = UnitsTypeEnum.kLbMassMassUnits
                    oPrecision = 4
                    applyToPart = True
                    applyToAssembly = True

                Case Else
                    Return

            End Select


            '==========================================================
            ' Áp dụng cho Active document (luôn luôn)
            '==========================================================
            SetUnits(openDoc, oUOM1, oUOM2, oPrecision)


            '==========================================================
            ' Áp dụng cho các document tham chiếu
            '==========================================================
            For Each docFile As Document In openDoc.AllReferencedDocuments

                Try
                    Select Case docFile.DocumentType

                        Case DocumentTypeEnum.kPartDocumentObject
                            If applyToPart Then
                                SetUnits(docFile, oUOM1, oUOM2, oPrecision)
                            End If

                        Case DocumentTypeEnum.kAssemblyDocumentObject
                            If applyToAssembly Then
                                SetUnits(docFile, oUOM1, oUOM2, oPrecision)
                            End If

                    End Select
                Catch
                End Try

            Next


            '==========================================================
            ' Update
            '==========================================================
            Try : openDoc.Update2(True) : Catch : End Try

            MessageBox.Show(
        "Đã cập nhật đơn vị:" & vbCrLf & selectedOption,
        "Đơn vị File",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information)

        End Sub


        '==============================================================
        ' Hàm thiết lập Unit cho Document
        '==============================================================
        Private Sub SetUnits(
            ByVal docFile As Document,
            ByVal lengthUnit As UnitsTypeEnum,
            ByVal massUnit As UnitsTypeEnum,
            ByVal precision As Integer)

            Try

                If docFile Is Nothing Then Return

                docFile.UnitsOfMeasure.LengthUnits = lengthUnit
                docFile.UnitsOfMeasure.MassUnits = massUnit
                docFile.UnitsOfMeasure.LengthDisplayPrecision = precision

                docFile.Update()

            Catch
                ' Bỏ qua document không thể update
            End Try

        End Sub


        '==============================================================
        ' Form chọn đơn vị
        '==============================================================
        Private Function ShowUnitSelectionForm(
            ByVal options As List(Of String)) As String

            Dim result As String = Nothing

            Using frm As New Form()

                frm.Text = "Đơn vị File"
                frm.Width = 430
                frm.Height = 190
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.FormBorderStyle = FormBorderStyle.FixedDialog
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False


                ' Label
                Dim lbl As New Label()

                lbl.Text = "Các loại đơn vị:"
                lbl.Left = 20
                lbl.Top = 20
                lbl.Width = 350

                frm.Controls.Add(lbl)


                ' ComboBox
                Dim cbo As New ComboBox()
                cbo.Left = 20
                cbo.Top = 45
                cbo.Width = 370
                cbo.DropDownStyle = ComboBoxStyle.DropDownList

                Dim firstItemIndex As Integer = -1

                For Each item As String In options

                    If String.IsNullOrEmpty(item) Then
                        ' Chèn dòng phân cách
                        cbo.Items.Add("──────────────")
                    Else
                        cbo.Items.Add(item)
                        If firstItemIndex < 0 Then firstItemIndex = cbo.Items.Count - 1
                    End If

                Next

                If firstItemIndex >= 0 Then cbo.SelectedIndex = firstItemIndex
                frm.Controls.Add(cbo)


                ' OK
                Dim btnOK As New Button()

                btnOK.Text = "OK"
                btnOK.Left = 225
                btnOK.Top = 90
                btnOK.Width = 80

                btnOK.DialogResult = DialogResult.OK

                frm.Controls.Add(btnOK)


                ' Cancel
                Dim btnCancel As New Button()

                btnCancel.Text = "Cancel"
                btnCancel.Left = 310
                btnCancel.Top = 90
                btnCancel.Width = 80

                btnCancel.DialogResult = DialogResult.Cancel

                frm.Controls.Add(btnCancel)


                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel


                If frm.ShowDialog() = DialogResult.OK Then
                    If cbo.SelectedItem IsNot Nothing Then
                        Dim s As String = cbo.SelectedItem.ToString()
                        If Not s.StartsWith("─") Then
                            result = s
                        End If
                    End If
                End If

            End Using

            Return result

        End Function

    End Module
End Namespace
