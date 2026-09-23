Option Explicit On
Option Strict Off
Imports Inventor
Imports System
Imports Env = System.Environment
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing

Namespace ToolInventor2025.Assembly.Buttons.Part
    Public Module Ass_Part_9

        Private Const ADSKLIB_FILE_NAME As String = "Mẫu sắc 1.adsklib"
        Private ReadOnly MaterialNames As String() = BuildMaterialNames()
        Private _cachedAdsklibPath As String = ""

        Private Function BuildMaterialNames() As String()
            Dim list As New List(Of String)
            For i As Integer = 1 To 24
                list.Add("Steel, Mild " & i.ToString())
            Next
            Return list.ToArray()
        End Function

        '=========================================================
        ' ENTRY POINT
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

                Dim choice As Integer = ShowChoiceForm()
                Select Case choice
                    Case 1 : DoRandomAll(oApp, oAssDoc)
                    Case 2 : DoPickSourceThenApplyAll(oApp, oAssDoc)
                    Case 3 : DoPickSourceThenPickTargets(oApp, oAssDoc)
                End Select
            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message)
            End Try
        End Sub

        '=========================================================
        ' FORM
        '=========================================================
        Private Function ShowChoiceForm() As Integer
            Dim result As Integer = 0
            Using frm As New Form()
                frm.AutoScaleMode = AutoScaleMode.None
                frm.Font = New Font("Segoe UI", 9.0F)
                frm.ClientSize = New Size(440, 245)
                frm.FormBorderStyle = FormBorderStyle.FixedSingle
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.TopMost = True
                frm.Text = "Material Tool"

                Dim lbl As New Label()
                lbl.Text = "Chọn chức năng:"
                lbl.Font = New Font("Segoe UI", 11.0F, FontStyle.Bold)
                lbl.Location = New System.Drawing.Point(20, 15)
                lbl.AutoSize = True
                frm.Controls.Add(lbl)

                Dim btn1 As New Button()
                btn1.Text = "1. Random 24 Material cho toàn bộ Part"
                btn1.Size = New Size(400, 40)
                btn1.Location = New System.Drawing.Point(20, 50)
                AddHandler btn1.Click, Sub()
                                           result = 1
                                           frm.DialogResult = DialogResult.OK
                                           frm.Close()
                                       End Sub
                frm.Controls.Add(btn1)

                Dim btn2 As New Button()
                btn2.Text = "2. Chọn 1 Part mẫu → áp dụng cho TẤT CẢ Part"
                btn2.Size = New Size(400, 40)
                btn2.Location = New System.Drawing.Point(20, 100)
                AddHandler btn2.Click, Sub()
                                           result = 2
                                           frm.DialogResult = DialogResult.OK
                                           frm.Close()
                                       End Sub
                frm.Controls.Add(btn2)

                Dim btn3 As New Button()
                btn3.Text = "3. Chọn Part mẫu → chọn nhiều Part đích bằng tay"
                btn3.Size = New Size(400, 40)
                btn3.Location = New System.Drawing.Point(20, 150)
                AddHandler btn3.Click, Sub()
                                           result = 3
                                           frm.DialogResult = DialogResult.OK
                                           frm.Close()
                                       End Sub
                frm.Controls.Add(btn3)

                Dim btnCancel As New Button()
                btnCancel.Text = "Hủy"
                btnCancel.Size = New Size(80, 28)
                btnCancel.Location = New System.Drawing.Point(340, 205)
                AddHandler btnCancel.Click, Sub()
                                                result = 0
                                                frm.DialogResult = DialogResult.Cancel
                                                frm.Close()
                                            End Sub
                frm.Controls.Add(btnCancel)
                frm.CancelButton = btnCancel
                frm.ShowDialog()
            End Using
            Return result
        End Function

        '=========================================================
        ' 1. RANDOM
        '=========================================================
        Private Sub DoRandomAll(ByVal oApp As Inventor.Application, ByVal oAssDoc As AssemblyDocument)
            Try
                Dim libb As AssetLibrary = EnsureLibraryInProject(oApp)
                If libb Is Nothing Then
                    MessageBox.Show("Không tìm thấy hoặc không add được thư viện '" & ADSKLIB_FILE_NAME & "'", "Lỗi")
                    Return
                End If

                Dim rand As New Random()
                Dim processedParts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                ProcessOccurrences(oApp, oAssDoc.ComponentDefinition.Occurrences, rand, processedParts)
                Try : oAssDoc.Update() : Catch : End Try
                PostStatus("Đã random Material cho toàn bộ Part.")
                MessageBox.Show("Đã random Material cho toàn bộ Part.", "Hoàn tất")
            Catch
            End Try
        End Sub

        '=========================================================
        ' 2. CHỌN MẪU → TẤT CẢ
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
                Try : srcMaterial = srcPart.ActiveMaterial : Catch : End Try
                If srcMaterial Is Nothing Then
                    MessageBox.Show("Part mẫu chưa có Material!", "Thông báo")
                    Return
                End If

                Dim srcAppearance As Asset = Nothing
                Try : srcAppearance = srcMaterial.AppearanceAsset : Catch : End Try

                Dim processedParts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Try : processedParts.Add(srcPart.FullFileName) : Catch : End Try

                Dim count As Integer = 0
                ApplyRecursive(oAssDoc.ComponentDefinition.Occurrences, srcMaterial, srcAppearance, processedParts, count)
                Try : oAssDoc.Update() : Catch : End Try
                MessageBox.Show("Đã áp dụng Material """ & srcMaterial.DisplayName & """ cho " & count & " Part.", "Hoàn tất")
            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message)
            End Try
        End Sub

        '=========================================================
        ' 3. CHỌN MẪU → CHỌN ĐÍCH
        '=========================================================
        Private Sub DoPickSourceThenPickTargets(ByVal oApp As Inventor.Application, ByVal oAssDoc As AssemblyDocument)
            Try
                Dim srcOcc As ComponentOccurrence = PickPart(oApp, "Chọn 1 Part MẪU (ESC để hủy)")
                If srcOcc Is Nothing Then Return

                Dim srcPart As PartDocument = TryCast(srcOcc.Definition.Document, PartDocument)
                If srcPart Is Nothing Then
                    MessageBox.Show("Part mẫu không phải Part thường!", "Thông báo")
                    Return
                End If

                Dim srcMaterial As MaterialAsset = Nothing
                Try : srcMaterial = srcPart.ActiveMaterial : Catch : End Try
                If srcMaterial Is Nothing Then
                    MessageBox.Show("Part mẫu chưa có Material!", "Thông báo")
                    Return
                End If

                Dim srcAppearance As Asset = Nothing
                Try : srcAppearance = srcMaterial.AppearanceAsset : Catch : End Try

                Try : oApp.ActiveDocument.SelectSet.Clear() : Catch : End Try

                Dim selOccs As New List(Of ComponentOccurrence)
                Do
                    Dim prompt As String = "Chọn Part đích (" & selOccs.Count & " đã chọn). ESC để kết thúc."
                    Dim picked As Object = Nothing
                    Try
                        picked = oApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, prompt)
                    Catch
                    End Try
                    If picked Is Nothing Then Exit Do

                    Dim occ As ComponentOccurrence = TryCast(picked, ComponentOccurrence)
                    If occ Is Nothing Then
                        Dim face As Face = TryCast(picked, Face)
                        If face IsNot Nothing Then occ = face.ContainingOccurrence
                    End If
                    If occ IsNot Nothing AndAlso occ IsNot srcOcc AndAlso Not selOccs.Contains(occ) Then
                        selOccs.Add(occ)
                    End If
                Loop

                If selOccs.Count = 0 Then Return

                Dim processedParts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Try : processedParts.Add(srcPart.FullFileName) : Catch : End Try

                Dim count As Integer = 0
                For Each occ As ComponentOccurrence In selOccs
                    Try
                        Dim pDoc As PartDocument = TryCast(occ.Definition.Document, PartDocument)
                        If pDoc IsNot Nothing Then
                            If ApplyMaterialTo(pDoc, srcMaterial, srcAppearance, processedParts) Then count += 1
                            Continue For
                        End If

                        Dim subAss As AssemblyDocument = TryCast(occ.Definition.Document, AssemblyDocument)
                        If subAss IsNot Nothing Then
                            ApplyRecursive(subAss.ComponentDefinition.Occurrences, srcMaterial, srcAppearance, processedParts, count)
                        End If
                    Catch
                    End Try
                Next

                Try : oAssDoc.Update() : Catch : End Try
                MessageBox.Show("Đã áp dụng cho " & count & " Part.", "Hoàn tất")
            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message)
            End Try
        End Sub

        '=========================================================
        ' ENSURE LIBRARY IN PROJECT (chỉ 1 lần)
        '=========================================================
        Private Function EnsureLibraryInProject(ByVal oApp As Inventor.Application) As AssetLibrary
            If oApp Is Nothing Then Return Nothing

            Dim libPath As String = GetAdsklibPath(ADSKLIB_FILE_NAME)
            If String.IsNullOrEmpty(libPath) OrElse Not System.IO.File.Exists(libPath) Then Return Nothing

            Dim proj As DesignProject = Nothing
            Try : proj = oApp.DesignProjectManager.ActiveDesignProject : Catch : End Try
            If proj Is Nothing Then Return Nothing

            ' Đã có chưa?
            Try
                For i As Integer = 1 To proj.MaterialLibraries.Count
                    Try
                        Dim libb As AssetLibrary = proj.MaterialLibraries.Item(i)
                        Dim libFile As String = ""
                        Try : libFile = libb.FullFileName : Catch : End Try

                        If Not String.IsNullOrEmpty(libFile) AndAlso
                           String.Equals(System.IO.Path.GetFullPath(libFile), System.IO.Path.GetFullPath(libPath), StringComparison.OrdinalIgnoreCase) Then
                            Return libb
                        End If
                        If StringEqualsLoose(libb.DisplayName, "Mẫu sắc 1") Then Return libb
                    Catch
                    End Try
                Next
            Catch
            End Try

            ' Chưa có → Add
            Try
                Return proj.MaterialLibraries.Add(libPath)
            Catch
                Try
                    Return oApp.AssetLibraries.Open(libPath)
                Catch
                    Return Nothing
                End Try
            End Try
        End Function

        '=========================================================
        ' ĐỆ QUY RANDOM
        '=========================================================
        Private Sub ProcessOccurrences(ByVal oApp As Inventor.Application,
                                       ByVal occurrences As ComponentOccurrences,
                                       ByVal rand As Random,
                                       ByVal processedParts As HashSet(Of String))
            If occurrences Is Nothing Then Return
            For Each occ As ComponentOccurrence In occurrences
                Try
                    If occ Is Nothing OrElse occ.Suppressed Then Continue For
                    Try : If occ.IsContentCenterMember Then Continue For
                    Catch : End Try

                    Dim bom As BOMStructureEnum = BOMStructureEnum.kNormalBOMStructure
                    Try : bom = occ.BOMStructure
                    Catch : End Try
                    If bom = BOMStructureEnum.kPurchasedBOMStructure OrElse bom = BOMStructureEnum.kPhantomBOMStructure Then Continue For

                    Dim partDoc As PartDocument = TryCast(occ.Definition.Document, PartDocument)
                    If partDoc IsNot Nothing Then
                        ProcessPart(oApp, partDoc, rand, processedParts)
                        Continue For
                    End If

                    Dim subAss As AssemblyDocument = TryCast(occ.Definition.Document, AssemblyDocument)
                    If subAss IsNot Nothing Then
                        ProcessOccurrences(oApp, subAss.ComponentDefinition.Occurrences, rand, processedParts)
                    End If
                Catch
                End Try
            Next
        End Sub

        Private Sub ProcessPart(ByVal oApp As Inventor.Application,
                                ByVal partDoc As PartDocument,
                                ByVal rand As Random,
                                ByVal processedParts As HashSet(Of String))
            Try
                If partDoc Is Nothing Then Return
                Try : If partDoc.ComponentDefinition.IsContentCenterMember Then Return
                Catch : End Try

                Dim fileName As String = ""
                Try : fileName = partDoc.FullFileName
                Catch : End Try
                If String.IsNullOrEmpty(fileName) Then
                    Try : fileName = partDoc.DisplayName
                    Catch : Return : End Try
                End If
                If processedParts.Contains(fileName) Then Return
                processedParts.Add(fileName)

                Dim materialName As String = MaterialNames(rand.Next(0, MaterialNames.Length))
                Dim materialAsset As MaterialAsset = GetMaterialAsset(oApp, partDoc, materialName)
                If materialAsset Is Nothing Then Return

                Try : partDoc.ActiveMaterial = materialAsset : Catch : Return : End Try
                Try : partDoc.AppearanceSourceType = AppearanceSourceTypeEnum.kMaterialAppearance : Catch : End Try

                Dim matAppearance As Asset = Nothing
                Try : matAppearance = materialAsset.AppearanceAsset : Catch : End Try
                If matAppearance IsNot Nothing Then ForceAppearanceOnBodies(partDoc, matAppearance)

                Try : partDoc.Update() : Catch : End Try
            Catch
            End Try
        End Sub

        Private Sub ForceAppearanceOnBodies(ByVal partDoc As PartDocument, ByVal app As Asset)
            If partDoc Is Nothing OrElse app Is Nothing Then Return
            Try
                For Each body As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies
                    Try
                        body.Appearance = app
                        body.AppearanceSourceType = AppearanceSourceTypeEnum.kOverrideAppearance
                    Catch
                    End Try
                Next
            Catch
            End Try
        End Sub

        '=========================================================
        ' GET MATERIAL
        '=========================================================
        Private Function GetMaterialAsset(ByVal oApp As Inventor.Application,
                                          ByVal oPartDoc As PartDocument,
                                          ByVal materialName As String) As MaterialAsset
            Dim existing As MaterialAsset = FindMaterialInDoc(oPartDoc, materialName)
            If existing IsNot Nothing Then Return existing

            Dim libb As AssetLibrary = EnsureLibraryInProject(oApp)
            If libb IsNot Nothing Then
                Dim m As MaterialAsset = FindInLibrary(libb, oPartDoc, materialName)
                If m IsNot Nothing Then Return m
            End If

            Try
                For Each lib2 As AssetLibrary In oApp.AssetLibraries
                    Dim m As MaterialAsset = FindInLibrary(lib2, oPartDoc, materialName)
                    If m IsNot Nothing Then Return m
                Next
            Catch
            End Try
            Return Nothing
        End Function

        Private Function FindInLibrary(ByVal libb As AssetLibrary,
                                       ByVal oPartDoc As PartDocument,
                                       ByVal materialName As String) As MaterialAsset
            If libb Is Nothing Then Return Nothing
            Dim libMat As MaterialAsset = Nothing
            Try
                For i As Integer = 1 To libb.MaterialAssets.Count
                    Try
                        Dim m As MaterialAsset = libb.MaterialAssets.Item(i)
                        If String.Equals(m.DisplayName, materialName, StringComparison.OrdinalIgnoreCase) Then
                            libMat = m
                            Exit For
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try
            If libMat Is Nothing Then Return Nothing

            Dim copiedMat As MaterialAsset = Nothing
            Try : copiedMat = TryCast(libMat.CopyTo(oPartDoc), MaterialAsset) : Catch : End Try
            If copiedMat Is Nothing Then Return Nothing

            Try
                Dim libAppearance As Asset = libMat.AppearanceAsset
                If libAppearance IsNot Nothing Then
                    Dim localAppearance As Asset = FindAppearance(oPartDoc, libAppearance.DisplayName)
                    If localAppearance Is Nothing Then
                        Try : localAppearance = libAppearance.CopyTo(oPartDoc) : Catch : End Try
                    End If
                    If localAppearance IsNot Nothing Then
                        Try : copiedMat.AppearanceAsset = localAppearance : Catch : End Try
                    End If
                End If
            Catch
            End Try
            Return copiedMat
        End Function

        Private Function FindAppearance(ByVal oPartDoc As PartDocument, ByVal appearanceName As String) As Asset
            Try
                For Each appAsset As Asset In oPartDoc.AppearanceAssets
                    If appAsset IsNot Nothing AndAlso String.Equals(appAsset.DisplayName, appearanceName, StringComparison.OrdinalIgnoreCase) Then
                        Return appAsset
                    End If
                Next
            Catch
            End Try
            Return Nothing
        End Function

        Private Function FindMaterialInDoc(ByVal doc As PartDocument, ByVal materialName As String) As MaterialAsset
            If doc Is Nothing Then Return Nothing
            Try
                For Each mat As MaterialAsset In doc.MaterialAssets
                    If mat IsNot Nothing AndAlso String.Equals(mat.DisplayName, materialName, StringComparison.OrdinalIgnoreCase) Then
                        Return mat
                    End If
                Next
            Catch
            End Try
            Return Nothing
        End Function

        '=========================================================
        ' PICK + APPLY
        '=========================================================
        Private Function PickPart(ByVal oApp As Inventor.Application, ByVal prompt As String) As ComponentOccurrence
            Try
                Dim picked As Object = oApp.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, prompt)
                If picked Is Nothing Then Return Nothing
                Dim face As Face = TryCast(picked, Face)
                If face IsNot Nothing Then Return face.ContainingOccurrence
            Catch
            End Try
            Return Nothing
        End Function

        Private Sub ApplyRecursive(ByVal occurrences As ComponentOccurrences,
                                   ByVal srcMaterial As MaterialAsset,
                                   ByVal srcAppearance As Asset,
                                   ByVal processedParts As HashSet(Of String),
                                   ByRef count As Integer)
            If occurrences Is Nothing Then Return
            For Each occ As ComponentOccurrence In occurrences
                Try
                    If occ.Suppressed Then Continue For
                    Try : If occ.IsContentCenterMember Then Continue For
                    Catch : End Try

                    Dim pDoc As PartDocument = TryCast(occ.Definition.Document, PartDocument)
                    If pDoc IsNot Nothing Then
                        If ApplyMaterialTo(pDoc, srcMaterial, srcAppearance, processedParts) Then count += 1
                        Continue For
                    End If

                    Dim subAss As AssemblyDocument = TryCast(occ.Definition.Document, AssemblyDocument)
                    If subAss IsNot Nothing Then
                        ApplyRecursive(subAss.ComponentDefinition.Occurrences, srcMaterial, srcAppearance, processedParts, count)
                    End If
                Catch
                End Try
            Next
        End Sub

        Private Function ApplyMaterialTo(ByVal targetDoc As PartDocument,
                                         ByVal srcMaterial As MaterialAsset,
                                         ByVal srcAppearance As Asset,
                                         ByVal processedParts As HashSet(Of String)) As Boolean
            Try
                If targetDoc Is Nothing OrElse srcMaterial Is Nothing Then Return False
                Try : If targetDoc.ComponentDefinition.IsContentCenterMember Then Return False
                Catch : End Try

                Dim fileName As String = ""
                Try : fileName = targetDoc.FullFileName : Catch : End Try
                If String.IsNullOrEmpty(fileName) Then
                    Try : fileName = targetDoc.DisplayName : Catch : Return False : End Try
                End If
                If processedParts.Contains(fileName) Then Return False
                processedParts.Add(fileName)

                Dim tMat As MaterialAsset = FindMaterialInDoc(targetDoc, srcMaterial.DisplayName)
                If tMat Is Nothing Then
                    Try : tMat = TryCast(srcMaterial.CopyTo(targetDoc), MaterialAsset) : Catch : End Try
                End If
                If tMat Is Nothing Then Return False

                Dim tApp As Asset = Nothing
                If srcAppearance IsNot Nothing Then
                    tApp = FindAppearance(targetDoc, srcAppearance.DisplayName)
                    If tApp Is Nothing Then
                        Try : tApp = srcAppearance.CopyTo(targetDoc) : Catch : End Try
                    End If
                    If tApp IsNot Nothing Then
                        Try : tMat.AppearanceAsset = tApp : Catch : End Try
                    End If
                End If

                Try : targetDoc.ActiveMaterial = tMat : Catch : Return False : End Try
                Try : targetDoc.AppearanceSourceType = AppearanceSourceTypeEnum.kMaterialAppearance : Catch : End Try
                If tApp IsNot Nothing Then ForceAppearanceOnBodies(targetDoc, tApp)
                Try : targetDoc.Update() : Catch : End Try
                Return True
            Catch
                Return False
            End Try
        End Function

        '=========================================================
        ' TÌM FILE ADSKLIB
        '=========================================================
        Private Function GetAdsklibPath(ByVal fileName As String) As String
            If Not String.IsNullOrEmpty(_cachedAdsklibPath) AndAlso System.IO.File.Exists(_cachedAdsklibPath) Then
                Return _cachedAdsklibPath
            End If

            Dim cached As String = ReadCachedPath()
            If Not String.IsNullOrEmpty(cached) AndAlso System.IO.File.Exists(cached) Then
                _cachedAdsklibPath = cached
                Return cached
            End If

            Dim roots As New List(Of String)
            Try
                Dim loc As String = System.Reflection.Assembly.GetExecutingAssembly().Location
                If Not String.IsNullOrEmpty(loc) Then roots.Add(System.IO.Path.GetDirectoryName(loc))
            Catch
            End Try
            Try
                Dim d As String = AppContext.BaseDirectory
                If Not String.IsNullOrEmpty(d) Then roots.Add(d.TrimEnd("\"c))
            Catch
            End Try
            Try
                Dim d As String = AppDomain.CurrentDomain.BaseDirectory
                If Not String.IsNullOrEmpty(d) Then roots.Add(d.TrimEnd("\"c))
            Catch
            End Try
            Try
                For Each d As String In GetAddinFileDirs()
                    If Not String.IsNullOrEmpty(d) Then roots.Add(d)
                Next
            Catch
            End Try

            Dim uniq As New List(Of String)
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each d As String In roots
                Try
                    Dim full As String = System.IO.Path.GetFullPath(d)
                    If seen.Add(full) Then uniq.Add(full)
                Catch
                End Try
            Next

            For Each root As String In uniq
                Dim found As String = FindFileRecursive(root, fileName, 6)
                If Not String.IsNullOrEmpty(found) Then
                    _cachedAdsklibPath = found
                    WriteCachedPath(found)
                    Return found
                End If
            Next

            Dim ans As DialogResult = MessageBox.Show(
                "Không tìm thấy file '" & fileName & "'." & vbCrLf & vbCrLf &
                "Bạn có muốn CHỌN FILE THỦ CÔNG không?" & vbCrLf &
                "(Đường dẫn sẽ được lưu cho lần sau)",
                "Không tìm thấy file", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If ans = DialogResult.Yes Then
                Using ofd As New OpenFileDialog()
                    ofd.Title = "Chọn file " & fileName
                    ofd.Filter = "Inventor Asset Library (*.adsklib)|*.adsklib|Tất cả (*.*)|*.*"
                    ofd.FileName = fileName
                    Try
                        ofd.InitialDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
                    Catch
                    End Try
                    If ofd.ShowDialog() = DialogResult.OK Then
                        _cachedAdsklibPath = ofd.FileName
                        WriteCachedPath(ofd.FileName)
                        Return ofd.FileName
                    End If
                End Using
            End If
            Return ""
        End Function

        Private Function FindFileRecursive(ByVal dir As String, ByVal fileName As String, ByVal maxDepth As Integer) As String
            If maxDepth < 0 OrElse Not System.IO.Directory.Exists(dir) Then Return Nothing
            Try
                Dim direct As String = System.IO.Path.Combine(dir, fileName)
                If System.IO.File.Exists(direct) Then Return direct

                For Each f As String In System.IO.Directory.GetFiles(dir, "*.adsklib")
                    If StringEqualsLoose(System.IO.Path.GetFileName(f), fileName) Then Return f
                Next

                If maxDepth > 0 Then
                    For Each subDir As String In System.IO.Directory.GetDirectories(dir)
                        Dim found As String = FindFileRecursive(subDir, fileName, maxDepth - 1)
                        If Not String.IsNullOrEmpty(found) Then Return found
                    Next
                End If
            Catch
            End Try
            Return Nothing
        End Function

        Private Function GetCacheFilePath() As String
            Try
                Dim asmPath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
                Return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(asmPath), "adsklib_path.txt")
            Catch
                Return ""
            End Try
        End Function

        Private Sub WriteCachedPath(ByVal path As String)
            Try
                Dim f As String = GetCacheFilePath()
                If Not String.IsNullOrEmpty(f) Then System.IO.File.WriteAllText(f, path, System.Text.Encoding.UTF8)
            Catch
            End Try
        End Sub

        Private Function ReadCachedPath() As String
            Try
                Dim f As String = GetCacheFilePath()
                If System.IO.File.Exists(f) Then Return System.IO.File.ReadAllText(f, System.Text.Encoding.UTF8).Trim()
            Catch
            End Try
            Return ""
        End Function

        Private Function StringEqualsLoose(ByVal a As String, ByVal b As String) As Boolean
            If a Is Nothing OrElse b Is Nothing Then Return False
            Return String.Equals(RemoveDiacritics(a), RemoveDiacritics(b), StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function RemoveDiacritics(ByVal s As String) As String
            If String.IsNullOrEmpty(s) Then Return s
            Try
                Dim normalized As String = s.Normalize(System.Text.NormalizationForm.FormD)
                Dim sb As New System.Text.StringBuilder()
                For Each c As Char In normalized
                    If System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) <> System.Globalization.UnicodeCategory.NonSpacingMark Then
                        sb.Append(c)
                    End If
                Next
                Return sb.ToString().Normalize(System.Text.NormalizationForm.FormC)
            Catch
                Return s
            End Try
        End Function

        Private Function GetAddinFileDirs() As List(Of String)
            Dim result As New List(Of String)
            Try
                Dim roots As String() = {
                    System.IO.Path.Combine(Env.GetFolderPath(Env.SpecialFolder.ApplicationData), "Autodesk", "ApplicationPlugins"),
                    System.IO.Path.Combine(Env.GetFolderPath(Env.SpecialFolder.CommonApplicationData), "Autodesk", "ApplicationPlugins")
                }
                For Each root As String In roots
                    If Not System.IO.Directory.Exists(root) Then Continue For
                    For Each af As String In System.IO.Directory.GetFiles(root, "*.addin", System.IO.SearchOption.AllDirectories)
                        result.Add(System.IO.Path.GetDirectoryName(af))
                    Next
                Next
            Catch
            End Try
            Return result
        End Function

        Private Sub PostStatus(ByVal msg As String)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus(msg)
            Catch
            End Try
        End Sub

    End Module
End Namespace