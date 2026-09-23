Option Explicit On
Option Strict Off

Imports Inventor
Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing

Namespace ToolInventor2025.Assembly.Buttons.Part

    Public Module Ass_Part_9

        '=========================================================
        ' 24 MATERIAL
        '=========================================================
        Private ReadOnly MaterialNames As String() = BuildMaterialNames()

        Private Function BuildMaterialNames() As String()
            Dim list As New List(Of String)
            list.Add("Steel, Mild")
            For i As Integer = 1 To 23
                list.Add("Steel, Mild " & i.ToString())
            Next
            Return list.ToArray()
        End Function

        '=========================================================
        ' ENTRY POINT - FORM 3 NÚT
        '=========================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                Dim oApp As Inventor.Application = g_inventorApplication
                If oApp Is Nothing Then Return

                Dim oDoc As Document = oApp.ActiveDocument
                If oDoc Is Nothing Then Return

                Dim oAssDoc As AssemblyDocument = TryCast(oDoc, AssemblyDocument)
                If oAssDoc Is Nothing Then
                    MessageBox.Show("Vui lòng mở Assembly!", "Thông báo")
                    Return
                End If

                '=============== FORM 3 NÚT ===============
                Dim choice As Integer = ShowChoiceForm()

                Select Case choice
                    Case 1
                        DoRandomAll(oApp, oAssDoc)
                    Case 2
                        DoPickSourceThenApplyAll(oApp, oAssDoc)
                    Case 3
                        DoPickSourceThenPickTargets(oApp, oAssDoc)
                    Case Else
                        ' User bấm Cancel / đóng form
                End Select

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message)
            End Try
        End Sub

        '=========================================================
        ' FORM CHỌN 3 CHỨC NĂNG (NÚT)
        '=========================================================
        Private Function ShowChoiceForm() As Integer
            Dim result As Integer = 0

            Using frm As New Form()
                frm.Text = "Material Tool - Inventor 2020"
                frm.FormBorderStyle = FormBorderStyle.FixedDialog
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.Size = New Size(420, 260)
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.TopMost = True

                Dim lbl As New Label()
                lbl.Text = "Chọn chức năng:"
                lbl.Font = New Font("Segoe UI", 11.0F, FontStyle.Bold)
                lbl.Location = New System.Drawing.Point(20, 15)
                lbl.AutoSize = True
                frm.Controls.Add(lbl)

                Dim btn1 As New Button()
                btn1.Text = "1. Random 24 Material cho toàn bộ Part"
                btn1.Size = New Size(360, 40)
                btn1.Location = New System.Drawing.Point(20, 50)
                btn1.Font = New Font("Segoe UI", 9.5F)
                AddHandler btn1.Click, Sub()
                                           result = 1
                                           frm.DialogResult = DialogResult.OK
                                           frm.Close()
                                       End Sub
                frm.Controls.Add(btn1)

                Dim btn2 As New Button()
                btn2.Text = "2. Chọn 1 Part mẫu → áp dụng cho TẤT CẢ Part"
                btn2.Size = New Size(360, 40)
                btn2.Location = New System.Drawing.Point(20, 100)
                btn2.Font = New Font("Segoe UI", 9.5F)
                AddHandler btn2.Click, Sub()
                                           result = 2
                                           frm.DialogResult = DialogResult.OK
                                           frm.Close()
                                       End Sub
                frm.Controls.Add(btn2)

                Dim btn3 As New Button()
                btn3.Text = "3. Chọn Part mẫu → chọn nhiều Part đích bằng tay"
                btn3.Size = New Size(360, 40)
                btn3.Location = New System.Drawing.Point(20, 150)
                btn3.Font = New Font("Segoe UI", 9.5F)
                AddHandler btn3.Click, Sub()
                                           result = 3
                                           frm.DialogResult = DialogResult.OK
                                           frm.Close()
                                       End Sub
                frm.Controls.Add(btn3)

                Dim btnCancel As New Button()
                btnCancel.Text = "Hủy"
                btnCancel.Size = New Size(80, 28)
                btnCancel.Location = New System.Drawing.Point(300, 200)
                AddHandler btnCancel.Click, Sub()
                                                result = 0
                                                frm.DialogResult = DialogResult.Cancel
                                                frm.Close()
                                            End Sub
                frm.Controls.Add(btnCancel)

                frm.ShowDialog()
            End Using

            Return result
        End Function

        '=========================================================
        ' 1. RANDOM 24 MATERIAL
        '=========================================================
        Private Sub DoRandomAll(ByVal oApp As Inventor.Application, ByVal oAssDoc As AssemblyDocument)
            Try
                Dim rand As New Random()
                Dim processedParts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                ProcessOccurrences(oApp, oAssDoc.ComponentDefinition.Occurrences, rand, processedParts)

                Try
                    oAssDoc.Update()
                Catch
                End Try

                PostStatus("Đã random Material cho Part thường.")
                MessageBox.Show("Đã random Material cho toàn bộ Part.", "Hoàn tất")
            Catch
            End Try
        End Sub

        '=========================================================
        ' 2. CHỌN 1 PART MẪU → ÁP DỤNG CHO TẤT CẢ
        '=========================================================
        Private Sub DoPickSourceThenApplyAll(ByVal oApp As Inventor.Application, ByVal oAssDoc As AssemblyDocument)
            Try
                Dim srcOcc As ComponentOccurrence = PickPart(oApp, "Chọn 1 Part MẪU (ESC để hủy)")
                If srcOcc Is Nothing Then Return

                Dim srcPart As PartDocument = TryCast(srcOcc.Definition.Document, PartDocument)
                If srcPart Is Nothing Then
                    MessageBox.Show("Part mẫu không phải Part thường!", "Thông báo")
                    Return
                End If

                Dim srcMaterial As MaterialAsset = Nothing
                Try
                    srcMaterial = srcPart.ActiveMaterial
                Catch
                End Try
                If srcMaterial Is Nothing Then
                    MessageBox.Show("Part mẫu chưa có Material!", "Thông báo")
                    Return
                End If

                Dim srcAppearance As Asset = Nothing
                Try
                    srcAppearance = srcMaterial.AppearanceAsset
                Catch
                End Try

                Dim processedParts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Try
                    processedParts.Add(srcPart.FullFileName)
                Catch
                End Try

                Dim count As Integer = 0
                ApplyRecursive(oAssDoc.ComponentDefinition.Occurrences, srcMaterial, srcAppearance, processedParts, count)

                Try
                    oAssDoc.Update()
                Catch
                End Try

                MessageBox.Show("Đã áp dụng Material """ & srcMaterial.DisplayName & """ cho " & count & " Part.", "Hoàn tất")
            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message)
            End Try
        End Sub

        '=========================================================
        ' 3. CHỌN PART MẪU → CHỌN NHIỀU PART ĐÍCH
        '    Click từng part để chọn | ESC để kết thúc / hủy
        '=========================================================
        Private Sub DoPickSourceThenPickTargets(ByVal oApp As Inventor.Application,
                                               ByVal oAssDoc As AssemblyDocument)
            Try
                '----- Bước 1: Chọn Part mẫu -----
                Dim srcOcc As ComponentOccurrence = PickPart(oApp, "Chọn 1 Part MẪU (ESC để hủy)")
                If srcOcc Is Nothing Then Return

                Dim srcPart As PartDocument = TryCast(srcOcc.Definition.Document, PartDocument)
                If srcPart Is Nothing Then
                    MessageBox.Show("Part mẫu không phải Part thường!", "Thông báo")
                    Return
                End If

                Dim srcMaterial As MaterialAsset = Nothing
                Try
                    srcMaterial = srcPart.ActiveMaterial
                Catch
                End Try
                If srcMaterial Is Nothing Then
                    MessageBox.Show("Part mẫu chưa có Material!", "Thông báo")
                    Return
                End If

                Dim srcAppearance As Asset = Nothing
                Try
                    srcAppearance = srcMaterial.AppearanceAsset
                Catch
                End Try

                '----- Bước 2: Xóa selection cũ -----
                Try
                    oApp.ActiveDocument.SelectSet.Clear()
                Catch
                End Try

                '----- Bước 3: LOOP chọn từng Part đích -----
                Dim selOccs As New List(Of ComponentOccurrence)

                Do
                    Dim prompt As String =
                        "Chọn Part đích (" & selOccs.Count &
                        " đã chọn). ESC để kết thúc."

                    Dim picked As Object = Nothing
                    Try
                        picked = oApp.CommandManager.Pick(
                            SelectionFilterEnum.kAssemblyOccurrenceFilter, prompt)
                    Catch
                    End Try

                    ' ESC hoặc không chọn gì → thoát loop
                    If picked Is Nothing Then Exit Do

                    Dim occ As ComponentOccurrence = TryCast(picked, ComponentOccurrence)

                    ' Nếu Pick trả về Face → lấy occurrence chứa nó
                    If occ Is Nothing Then
                        Dim face As Face = TryCast(picked, Face)
                        If face IsNot Nothing Then occ = face.ContainingOccurrence
                    End If

                    ' Không chọn trùng Part mẫu, không chọn trùng đã chọn
                    If occ IsNot Nothing AndAlso occ IsNot srcOcc AndAlso
                       Not selOccs.Contains(occ) Then
                        selOccs.Add(occ)
                    End If
                Loop

                '----- Bước 4: Không chọn gì → hủy êm -----
                If selOccs.Count = 0 Then
                    PostStatus("Đã hủy - không có Part đích nào được chọn.")
                    Return
                End If

                '----- Bước 5: Áp dụng -----
                Dim processedParts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Try
                    processedParts.Add(srcPart.FullFileName)
                Catch
                End Try

                Dim count As Integer = 0
                For Each occ As ComponentOccurrence In selOccs
                    Try
                        Dim pDoc As PartDocument = TryCast(occ.Definition.Document, PartDocument)
                        If pDoc IsNot Nothing Then
                            If ApplyMaterialTo(pDoc, srcMaterial, srcAppearance, processedParts) Then
                                count += 1
                            End If
                            Continue For
                        End If

                        Dim subAss As AssemblyDocument =
                            TryCast(occ.Definition.Document, AssemblyDocument)
                        If subAss IsNot Nothing Then
                            ApplyRecursive(subAss.ComponentDefinition.Occurrences,
                                           srcMaterial, srcAppearance,
                                           processedParts, count)
                        End If
                    Catch
                    End Try
                Next

                Try
                    oAssDoc.Update()
                Catch
                End Try

                MessageBox.Show("Đã áp dụng cho " & count & " Part.", "Hoàn tất")
                PostStatus("Hoàn tất - đã áp dụng Material")

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message)
            End Try
        End Sub

        '=========================================================
        ' XỬ LÝ OCCURRENCES (RANDOM)
        '=========================================================
        Private Sub ProcessOccurrences(
            ByVal oApp As Inventor.Application,
            ByVal occurrences As ComponentOccurrences,
            ByVal rand As Random,
            ByVal processedParts As HashSet(Of String))

            If occurrences Is Nothing Then Return

            For Each occ As ComponentOccurrence In occurrences
                Try
                    If occ Is Nothing Then Continue For

                    Try
                        If occ.Suppressed Then Continue For
                    Catch
                    End Try

                    Try
                        If occ.IsContentCenterMember Then Continue For
                    Catch
                    End Try

                    Dim bomStructure As BOMStructureEnum = BOMStructureEnum.kNormalBOMStructure
                    Try
                        bomStructure = occ.Definition.BOMStructure
                    Catch
                        Try
                            bomStructure = occ.BOMStructure
                        Catch
                        End Try
                    End Try

                    If bomStructure = BOMStructureEnum.kPurchasedBOMStructure Then Continue For
                    If bomStructure = BOMStructureEnum.kPhantomBOMStructure Then Continue For

                    Dim occDoc As Document = Nothing
                    Try
                        occDoc = occ.Definition.Document
                    Catch
                    End Try
                    If occDoc Is Nothing Then Continue For

                    Dim partDoc As PartDocument = TryCast(occDoc, PartDocument)
                    If partDoc IsNot Nothing Then
                        ProcessPart(oApp, partDoc, rand, processedParts)
                        Continue For
                    End If

                    Dim subAssDoc As AssemblyDocument = TryCast(occDoc, AssemblyDocument)
                    If subAssDoc IsNot Nothing Then
                        Try
                            ProcessOccurrences(oApp, subAssDoc.ComponentDefinition.Occurrences, rand, processedParts)
                        Catch
                        End Try
                    End If
                Catch
                    Continue For
                End Try
            Next
        End Sub

        '=========================================================
        ' XỬ LÝ 1 PART (RANDOM)
        '=========================================================
        Private Sub ProcessPart(
            ByVal oApp As Inventor.Application,
            ByVal partDoc As PartDocument,
            ByVal rand As Random,
            ByVal processedParts As HashSet(Of String))

            Try
                If partDoc Is Nothing Then Return

                Try
                    If partDoc.ComponentDefinition.IsContentCenterMember Then Return
                Catch
                End Try

                Try
                    Dim bom As BOMStructureEnum = partDoc.ComponentDefinition.BOMStructure
                    If bom = BOMStructureEnum.kPurchasedBOMStructure Then Return
                    If bom = BOMStructureEnum.kPhantomBOMStructure Then Return
                Catch
                End Try

                Dim fileName As String = ""
                Try
                    fileName = partDoc.FullFileName
                Catch
                End Try
                If String.IsNullOrEmpty(fileName) Then
                    Try
                        fileName = partDoc.DisplayName
                    Catch
                        Return
                    End Try
                End If

                If processedParts.Contains(fileName) Then Return
                processedParts.Add(fileName)

                Dim index As Integer = rand.Next(0, MaterialNames.Length)
                Dim materialName As String = MaterialNames(index)

                Dim materialAsset As MaterialAsset = GetMaterialAsset(oApp, partDoc, materialName)
                If materialAsset Is Nothing Then Return

                Try
                    partDoc.ActiveMaterial = materialAsset
                Catch
                    Return
                End Try

                Try
                    partDoc.AppearanceSourceType = AppearanceSourceTypeEnum.kMaterialAppearance
                Catch
                End Try

                Try
                    partDoc.Update()
                Catch
                End Try
            Catch
            End Try
        End Sub

        '=========================================================
        ' TÌM MATERIAL
        '=========================================================
        Private Function GetMaterialAsset(
            ByVal oApp As Inventor.Application,
            ByVal oPartDoc As PartDocument,
            ByVal materialName As String) As MaterialAsset

            Try
                For Each mat As MaterialAsset In oPartDoc.MaterialAssets
                    Try
                        If mat IsNot Nothing AndAlso
                           String.Equals(mat.DisplayName, materialName, StringComparison.OrdinalIgnoreCase) Then
                            Return mat
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            Try
                For Each libqq As AssetLibrary In oApp.AssetLibraries
                    Try
                        Dim libMat As MaterialAsset = Nothing
                        For Each mat As MaterialAsset In libqq.MaterialAssets
                            Try
                                If mat IsNot Nothing AndAlso
                                   String.Equals(mat.DisplayName, materialName, StringComparison.OrdinalIgnoreCase) Then
                                    libMat = mat
                                    Exit For
                                End If
                            Catch
                            End Try
                        Next
                        If libMat Is Nothing Then Continue For

                        Dim copiedMat As MaterialAsset = Nothing
                        Try
                            copiedMat = TryCast(libMat.CopyTo(oPartDoc), MaterialAsset)
                        Catch
                        End Try
                        If copiedMat Is Nothing Then Continue For

                        Try
                            Dim libAppearance As Asset = libMat.AppearanceAsset
                            If libAppearance IsNot Nothing Then
                                Dim localAppearance As Asset = FindAppearance(oPartDoc, libAppearance.DisplayName)
                                If localAppearance Is Nothing Then
                                    Try
                                        localAppearance = libAppearance.CopyTo(oPartDoc)
                                    Catch
                                        localAppearance = Nothing
                                    End Try
                                End If
                                If localAppearance IsNot Nothing Then
                                    Try
                                        copiedMat.AppearanceAsset = localAppearance
                                    Catch
                                    End Try
                                End If
                            End If
                        Catch
                        End Try

                        Return copiedMat
                    Catch
                    End Try
                Next
            Catch
            End Try

            Return Nothing
        End Function

        '=========================================================
        ' TÌM APPEARANCE
        '=========================================================
        Private Function FindAppearance(ByVal oPartDoc As PartDocument, ByVal appearanceName As String) As Asset
            Try
                For Each appAsset As Asset In oPartDoc.AppearanceAssets
                    Try
                        If appAsset Is Nothing Then Continue For
                        If String.Equals(appAsset.DisplayName, appearanceName, StringComparison.OrdinalIgnoreCase) Then
                            Return appAsset
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try
            Return Nothing
        End Function

        '=========================================================
        ' PICK 1 PART
        '=========================================================
        Private Function PickPart(ByVal oApp As Inventor.Application, ByVal prompt As String) As ComponentOccurrence
            Try
                Dim picked As Object = oApp.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, prompt)
                If picked Is Nothing Then Return Nothing

                Dim face As Face = TryCast(picked, Face)
                If face IsNot Nothing Then Return face.ContainingOccurrence

                Return Nothing
            Catch
                Return Nothing
            End Try
        End Function

        '=========================================================
        ' LẤY OCCURRENCE TỪ SELECTSET
        '=========================================================
        Private Function GetOccFromItem(ByVal item As Object) As ComponentOccurrence
            If item Is Nothing Then Return Nothing

            Try
                Dim occ As ComponentOccurrence = TryCast(item, ComponentOccurrence)
                If occ IsNot Nothing Then Return occ
            Catch
            End Try

            Try
                Dim face As Face = TryCast(item, Face)
                If face IsNot Nothing Then Return face.ContainingOccurrence
            Catch
            End Try

            Try
                Dim edge As Edge = TryCast(item, Edge)
                If edge IsNot Nothing Then Return edge.ContainingOccurrence
            Catch
            End Try

            Try
                Dim vtx As Vertex = TryCast(item, Vertex)
                If vtx IsNot Nothing Then Return vtx.ContainingOccurrence
            Catch
            End Try

            Return Nothing
        End Function

        '=========================================================
        ' ÁP DỤNG ĐỆ QUY
        '=========================================================
        Private Sub ApplyRecursive(
            ByVal occurrences As ComponentOccurrences,
            ByVal srcMaterial As MaterialAsset,
            ByVal srcAppearance As Asset,
            ByVal processedParts As HashSet(Of String),
            ByRef count As Integer)

            If occurrences Is Nothing Then Return

            For Each occ As ComponentOccurrence In occurrences
                Try
                    If occ.Suppressed Then Continue For

                    Try
                        If occ.IsContentCenterMember Then Continue For
                    Catch
                    End Try

                    Dim pDoc As PartDocument = TryCast(occ.Definition.Document, PartDocument)
                    If pDoc IsNot Nothing Then
                        If ApplyMaterialTo(pDoc, srcMaterial, srcAppearance, processedParts) Then
                            count += 1
                        End If
                        Continue For
                    End If

                    Dim subAss As AssemblyDocument = TryCast(occ.Definition.Document, AssemblyDocument)
                    If subAss IsNot Nothing Then
                        Try
                            ApplyRecursive(subAss.ComponentDefinition.Occurrences, srcMaterial, srcAppearance, processedParts, count)
                        Catch
                        End Try
                    End If
                Catch
                End Try
            Next
        End Sub

        '=========================================================
        ' ÁP DỤNG MATERIAL CHO 1 PART
        '=========================================================
        Private Function ApplyMaterialTo(
            ByVal targetDoc As PartDocument,
            ByVal srcMaterial As MaterialAsset,
            ByVal srcAppearance As Asset,
            ByVal processedParts As HashSet(Of String)) As Boolean

            Try
                If targetDoc Is Nothing OrElse srcMaterial Is Nothing Then Return False

                Try
                    If targetDoc.ComponentDefinition.IsContentCenterMember Then Return False
                Catch
                End Try

                Dim fileName As String = ""
                Try
                    fileName = targetDoc.FullFileName
                Catch
                End Try
                If String.IsNullOrEmpty(fileName) Then
                    Try
                        fileName = targetDoc.DisplayName
                    Catch
                        Return False
                    End Try
                End If

                If processedParts.Contains(fileName) Then Return False
                processedParts.Add(fileName)

                Dim tMat As MaterialAsset = FindMaterialInDoc(targetDoc, srcMaterial.DisplayName)
                If tMat Is Nothing Then
                    Try
                        tMat = TryCast(srcMaterial.CopyTo(targetDoc), MaterialAsset)
                    Catch
                    End Try
                End If
                If tMat Is Nothing Then Return False

                Try
                    targetDoc.ActiveMaterial = tMat
                Catch
                    Return False
                End Try

                If srcAppearance IsNot Nothing Then
                    Dim tApp As Asset = FindAppearance(targetDoc, srcAppearance.DisplayName)
                    If tApp Is Nothing Then
                        Try
                            tApp = srcAppearance.CopyTo(targetDoc)
                        Catch
                        End Try
                    End If
                    If tApp IsNot Nothing Then
                        Try
                            tMat.AppearanceAsset = tApp
                        Catch
                        End Try
                    End If
                End If

                Try
                    targetDoc.AppearanceSourceType = AppearanceSourceTypeEnum.kMaterialAppearance
                Catch
                End Try

                Try
                    targetDoc.Update()
                Catch
                End Try

                Return True
            Catch
                Return False
            End Try
        End Function

        '=========================================================
        ' TÌM MATERIAL TRONG DOCUMENT
        '=========================================================
        Private Function FindMaterialInDoc(ByVal doc As PartDocument, ByVal materialName As String) As MaterialAsset
            If doc Is Nothing Then Return Nothing
            Try
                For Each mat As MaterialAsset In doc.MaterialAssets
                    Try
                        If mat IsNot Nothing AndAlso
                           String.Equals(mat.DisplayName, materialName, StringComparison.OrdinalIgnoreCase) Then
                            Return mat
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try
            Return Nothing
        End Function

        '=========================================================
        ' STATUS BAR
        '=========================================================
        Private Sub PostStatus(ByVal msg As String)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus(msg)
            Catch
            End Try
        End Sub

    End Module

End Namespace