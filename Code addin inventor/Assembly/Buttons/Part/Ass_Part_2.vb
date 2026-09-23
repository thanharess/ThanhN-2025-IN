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

            unitOptions.Add("All Part Thành mm - kg")
            unitOptions.Add("All Assembly Thành mm - kg")
            unitOptions.Add("All Assembly và Part Thành mm - kg")
            unitOptions.Add("")
            unitOptions.Add("All Assembly và Part Thành cm - kg")
            unitOptions.Add("All Assembly và Part Thành m - kg")


            '==========================================================
            ' Hiện Form chọn
            '==========================================================
            Dim selectedOption As String =
                ShowUnitSelectionForm(unitOptions)

            ' Cancel
            If String.IsNullOrEmpty(selectedOption) Then
                Return
            End If


            '==========================================================
            ' Active Document
            '==========================================================
            Dim openDoc As Document =
                g_inventorApplication.ActiveDocument

            If openDoc Is Nothing Then
                MessageBox.Show(
                    "Không có Document đang mở.",
                    "Đơn vị File",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
                Return
            End If


            '==========================================================
            ' Thiết lập đơn vị
            '==========================================================
            Dim oUOM1 As UnitsTypeEnum
            Dim oUOM2 As UnitsTypeEnum
            Dim oPrecision As Integer = 3


            Select Case selectedOption

                Case "All Part Thành mm - kg"

                    oUOM1 = UnitsTypeEnum.kMillimeterLengthUnits
                    oUOM2 = UnitsTypeEnum.kKilogramMassUnits

                    ' Active document
                    openDoc.UnitsOfMeasure.LengthUnits = oUOM1
                    openDoc.UnitsOfMeasure.MassUnits = oUOM2
                    openDoc.UnitsOfMeasure.LengthDisplayPrecision = oPrecision

                    ' Chỉ Part
                    For Each docFile As Document In openDoc.AllReferencedDocuments

                        If docFile.DocumentType =
                            DocumentTypeEnum.kPartDocumentObject Then

                            SetUnits(
                                docFile,
                                oUOM1,
                                oUOM2,
                                oPrecision)

                        End If

                    Next


                Case "All Assembly Thành mm - kg"

                    oUOM1 = UnitsTypeEnum.kMillimeterLengthUnits
                    oUOM2 = UnitsTypeEnum.kKilogramMassUnits

                    ' Active document
                    openDoc.UnitsOfMeasure.LengthUnits = oUOM1
                    openDoc.UnitsOfMeasure.MassUnits = oUOM2
                    openDoc.UnitsOfMeasure.LengthDisplayPrecision = oPrecision

                    ' Chỉ Assembly
                    For Each docFile As Document In openDoc.AllReferencedDocuments

                        If docFile.DocumentType =
                            DocumentTypeEnum.kAssemblyDocumentObject Then

                            SetUnits(
                                docFile,
                                oUOM1,
                                oUOM2,
                                oPrecision)

                        End If

                    Next


                Case "All Assembly và Part Thành mm - kg"

                    oUOM1 = UnitsTypeEnum.kMillimeterLengthUnits
                    oUOM2 = UnitsTypeEnum.kKilogramMassUnits

                    SetUnits(
                        openDoc,
                        oUOM1,
                        oUOM2,
                        oPrecision)

                    ' Assembly + Part
                    For Each docFile As Document In openDoc.AllReferencedDocuments

                        SetUnits(
                            docFile,
                            oUOM1,
                            oUOM2,
                            oPrecision)

                    Next


                Case "All Assembly và Part Thành cm - kg"

                    oUOM1 = UnitsTypeEnum.kCentimeterLengthUnits
                    oUOM2 = UnitsTypeEnum.kKilogramMassUnits

                    SetUnits(
                        openDoc,
                        oUOM1,
                        oUOM2,
                        oPrecision)

                    ' Assembly + Part
                    For Each docFile As Document In openDoc.AllReferencedDocuments

                        SetUnits(
                            docFile,
                            oUOM1,
                            oUOM2,
                            oPrecision)

                    Next


                Case "All Assembly và Part Thành m - kg"

                    oUOM1 = UnitsTypeEnum.kMeterLengthUnits
                    oUOM2 = UnitsTypeEnum.kKilogramMassUnits

                    SetUnits(
                        openDoc,
                        oUOM1,
                        oUOM2,
                        oPrecision)

                    ' Assembly + Part
                    For Each docFile As Document In openDoc.AllReferencedDocuments

                        SetUnits(
                            docFile,
                            oUOM1,
                            oUOM2,
                            oPrecision)

                    Next

            End Select


            '==========================================================
            ' Update
            '==========================================================
            openDoc.Update()

            MessageBox.Show(
                "Đã cập nhật đơn vị:" &
                vbCrLf & selectedOption,
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

                For Each item As String In options

                    ' Không đưa dòng trống vào ComboBox
                    If Not String.IsNullOrEmpty(item) Then
                        cbo.Items.Add(item)
                    End If

                Next

                If cbo.Items.Count > 0 Then
                    cbo.SelectedIndex = 0
                End If

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
                        result = cbo.SelectedItem.ToString()
                    End If

                End If

            End Using

            Return result

        End Function

    End Module
End Namespace
