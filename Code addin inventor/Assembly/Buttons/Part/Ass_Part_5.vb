Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor
Imports System.IO

Namespace ToolInventor2025.Assembly.Buttons.Part
    Public Module Ass_part_5
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim invApp As Inventor.Application
            Try
                invApp = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch ex As Exception
                MessageBox.Show("Inventor chưa chạy.")
                Return
            End Try

            ' Đọc đường dẫn lần trước từ file txt
            Dim txtfile As String = "C:\Users\Public\Documents\temp.txt"
            Dim LastPath As String = ""
            If System.IO.File.Exists(txtfile) Then
                Try
                    Dim oRead As String = System.IO.File.ReadAllText(txtfile)
                    If System.IO.File.Exists(oRead) Then
                        LastPath = System.IO.Path.GetDirectoryName(oRead) & "\"
                    ElseIf System.IO.Directory.Exists(oRead) Then
                        LastPath = oRead
                        If Not LastPath.EndsWith("\") Then LastPath &= "\"
                    End If
                Catch
                End Try
            End If

            Dim oDoc As AssemblyDocument = TryCast(invApp.ActiveEditDocument, AssemblyDocument)
            If oDoc Is Nothing Then
                MessageBox.Show("Không phải Assembly Document.")
                Return
            End If

            Dim oOcc As ComponentOccurrence = Nothing
            Dim FileName As String = ""
            Dim NewPathName As String = ""

            Select Case oDoc.SelectSet.Count
                Case 0
                    oOcc = invApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Chọn File")
                    If oOcc Is Nothing Then Return
                    FileName = oOcc.Definition.Document.FullDocumentName
                    NewPathName = System.IO.Path.GetDirectoryName(oDoc.FullDocumentName) & "\" &
                                  System.IO.Path.GetFileNameWithoutExtension(FileName) & "copy"
                Case 1
                    oOcc = TryCast(oDoc.SelectSet.Item(1), ComponentOccurrence)
                    If oOcc Is Nothing Then Return
                    FileName = oOcc.Definition.Document.FullDocumentName
                    NewPathName = System.IO.Path.GetDirectoryName(FileName) & "\" &
                                  System.IO.Path.GetFileNameWithoutExtension(FileName) & "copy"
                Case Else
                    oDoc.SelectSet.Clear()
                    Return
            End Select

            Dim idwName As String = Left(FileName, Len(FileName) - 4) & ".idw"
            Dim extName As String = System.IO.Path.GetExtension(FileName).ToLower()
            Dim oFileDlg As Inventor.FileDialog = Nothing
            invApp.CreateFileDialog(oFileDlg)

            If extName = ".iam" Then
                oFileDlg.Filter = "Inventor Files (*.iam)|*.iam|All Files (*.*)|*.*"
            ElseIf extName = ".ipt" Then
                oFileDlg.Filter = "Inventor Files (*.ipt)|*.ipt|All Files (*.*)|*.*"
            End If

            oFileDlg.DialogTitle = "Replace: " & FileName
            oFileDlg.FileName = NewPathName
            oFileDlg.CancelError = True

            Dim selectedfile As String = ""
            Try
                oFileDlg.ShowSave()
                selectedfile = oFileDlg.FileName
            Catch
                selectedfile = LastPath & System.IO.Path.GetFileNameWithoutExtension(FileName) & "copy"
            End Try

            If String.IsNullOrEmpty(selectedfile) Then Return
            If System.IO.Path.GetExtension(selectedfile).ToLower() <> extName Then
                selectedfile &= extName
            End If

            Dim newIdwName As String = Left(selectedfile, Len(selectedfile) - 4) & ".idw"

            Try
                oOcc.Definition.Document.SaveAs(selectedfile, True)
                oOcc.Replace(selectedfile, False)
            Catch ex As Exception
                MessageBox.Show("Không thể SaveAs/Replace: " & ex.Message)
                Return
            End Try

            ' Copy và update IDW nếu có
            If System.IO.File.Exists(idwName) Then
                Try
                    System.IO.File.Copy(idwName, newIdwName, True)
                    Dim newIdwFile As DrawingDocument = invApp.Documents.Open(newIdwName, True)
                    If newIdwFile.Sheets.Count > 0 AndAlso newIdwFile.Sheets(1).DrawingViews.Count > 0 Then
                        newIdwFile.Sheets(1).DrawingViews(1).ReferencedFile.DocumentDescriptor.ReferencedFileDescriptor.ReplaceReference(selectedfile)
                        newIdwFile.Update()
                        newIdwFile.Save()
                    End If
                    newIdwFile.Close()
                Catch ex As Exception
                    MessageBox.Show("Không thể copy/update IDW: " & ex.Message)
                End Try
            End If

            ' Ghi lại đường dẫn vào txt
            Try
                Using iWrite As StreamWriter = System.IO.File.CreateText(txtfile)
                    iWrite.Write(selectedfile)
                End Using
            Catch
            End Try

            MessageBox.Show("Hoàn tất copy và replace!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
    End Module
End Namespace
