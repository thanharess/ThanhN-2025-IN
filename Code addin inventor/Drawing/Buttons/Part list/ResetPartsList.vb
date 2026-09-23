Imports System
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor


Namespace ToolInventor2025.Drawing.Buttons
    Public Module Draw_ResetPartsList
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application = CType(Interop.Marshal2.GetActiveObject("Inventor.Application"), Inventor.Application)
            Dim oDrawDoc As DrawingDocument = CType(invApp.ActiveDocument, DrawingDocument)

            Dim choice As MsgBoxResult
            choice = MsgBox("Yes = Reset Active Sheet" & vbCrLf &
                    "No = Reset All Sheets" & vbCrLf &
                    "Cancel = Thoát",
                    MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question,
                    "Reset Parts List")

            Select Case choice
                Case MsgBoxResult.Yes
                    Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                    ResetPartsListOnSheet(oSheet)

                Case MsgBoxResult.No
                    For Each oSheet As Sheet In oDrawDoc.Sheets
                        ResetPartsListOnSheet(oSheet)
                    Next

                Case MsgBoxResult.Cancel
                    Exit Sub
            End Select

            ' MsgBox("Xong")
        End Sub

        Private Sub ResetPartsListOnSheet(oSheet As Sheet)
            If oSheet.PartsLists.Count > 0 Then
                For Each oPartsList As PartsList In oSheet.PartsLists
                    For Each oRow As PartsListRow In oPartsList.PartsListRows
                        For Each oColumn As PartsListColumn In oPartsList.PartsListColumns
                            oRow.Item(oColumn).Static = False
                        Next
                    Next
                Next
            End If

        End Sub

    End Module
End Namespace