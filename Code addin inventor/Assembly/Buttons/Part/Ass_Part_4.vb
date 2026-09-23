Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2025.Assembly.Buttons.Part
    Public Module Ass_Part_4

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                CreateAllFlatPatterns()
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace,
                                "Assembly - Create Flat Pattern",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub CreateAllFlatPatterns()

            Dim oApp As Inventor.Application = g_inventorApplication
            Dim oActiveDoc As Document = oApp.ActiveDocument

            If oActiveDoc Is Nothing OrElse
               oActiveDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then

                MessageBox.Show("Vui lòng mở Assembly trước.",
                                "Create Flat Pattern",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
                Return
            End If

            Dim oAsmDoc As AssemblyDocument = CType(oActiveDoc, AssemblyDocument)

            Dim successCount As Integer = 0
            Dim alreadyCount As Integer = 0
            Dim skipCount As Integer = 0
            Dim errorList As New System.Collections.Generic.List(Of String)

            Dim oldSilent As Boolean = oApp.SilentOperation
            oApp.SilentOperation = True

            Try
                Dim oRefDocs As DocumentsEnumerator = oAsmDoc.AllReferencedDocuments

                For Each oRefDoc As Document In oRefDocs

                    Dim oPartDoc As PartDocument = Nothing
                    Dim needClose As Boolean = False

                    Try
                        '--------------------------------------------------
                        ' Chỉ lấy Part Sheet Metal (dùng SubType chuẩn)
                        '--------------------------------------------------
                        If oRefDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                            Continue For
                        End If

                        ' GUID Sheet Metal Part
                        If oRefDoc.SubType <> "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                            skipCount += 1
                            Continue For
                        End If

                        Dim fullPath As String = oRefDoc.FullFileName
                        If String.IsNullOrEmpty(fullPath) OrElse Not System.IO.File.Exists(fullPath) Then
                            skipCount += 1
                            Continue For
                        End If

                        '--------------------------------------------------
                        ' Mở document (True = visible tạm thời để ổn định hơn)
                        '--------------------------------------------------
                        oPartDoc = CType(oApp.Documents.Open(fullPath, True), PartDocument)
                        needClose = True

                        Dim oCompDef As SheetMetalComponentDefinition =
                            CType(oPartDoc.ComponentDefinition, SheetMetalComponentDefinition)

                        '--------------------------------------------------
                        ' Đã có Flat Pattern → bỏ qua
                        '--------------------------------------------------
                        If oCompDef.HasFlatPattern Then
                            alreadyCount += 1
                        Else
                            ' Chưa có → Unfold
                            oCompDef.Unfold()

                            ' Bắt buộc phải ExitEdit sau Unfold
                            Try
                                oCompDef.FlatPattern.ExitEdit()
                            Catch
                            End Try

                            successCount += 1
                        End If

                        ' Save
                        oPartDoc.Update2(True)
                        oPartDoc.Save2(True)

                    Catch ex As Exception
                        skipCount += 1
                        errorList.Add(oRefDoc.DisplayName & " → " & ex.Message)
                    Finally
                        '--------------------------------------------------
                        ' Đóng an toàn
                        '--------------------------------------------------
                        If needClose AndAlso oPartDoc IsNot Nothing Then
                            Try
                                ' Đảm bảo thoát Flat Pattern trước khi đóng
                                Try
                                    Dim smDef As SheetMetalComponentDefinition =
                                        TryCast(oPartDoc.ComponentDefinition, SheetMetalComponentDefinition)
                                    If smDef IsNot Nothing AndAlso smDef.HasFlatPattern Then
                                        smDef.FlatPattern.ExitEdit()
                                    End If
                                Catch
                                End Try

                                oPartDoc.Close(True)   ' True = save
                            Catch
                            End Try
                        End If
                    End Try

                    ' Cho Inventor thở một chút
                    System.Windows.Forms.Application.DoEvents()

                Next

                ' Update Assembly
                Try
                    oAsmDoc.Update2(True)
                Catch
                End Try

            Finally
                oApp.SilentOperation = oldSilent
            End Try

            '==========================================================
            ' Thông báo kết quả
            '==========================================================
            Dim msg As String =
                "Hoàn thành tạo Flat Pattern." & vbCrLf & vbCrLf &
                "Đã tạo mới: " & successCount.ToString() & vbCrLf &
                "Đã có Flat Pattern (bỏ qua): " & alreadyCount.ToString() & vbCrLf &
                "Bỏ qua / Lỗi: " & skipCount.ToString()

            If errorList.Count > 0 Then
                msg &= vbCrLf & vbCrLf & "Chi tiết lỗi (tối đa 5):" & vbCrLf
                For i As Integer = 0 To Math.Min(4, errorList.Count - 1)
                    msg &= "• " & errorList(i) & vbCrLf
                Next
            End If

            MessageBox.Show(msg,
                            "Assembly - Create Flat Pattern",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

        End Sub

    End Module
End Namespace