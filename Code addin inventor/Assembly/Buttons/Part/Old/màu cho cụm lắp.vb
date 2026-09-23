Option Explicit On
Option Strict Off

Imports Inventor
Imports System
Imports System.Collections.Generic

Namespace ToolInventor2025.Assembly.Buttons.Part

    Public Module màuuuuuuuu

        '=========================================================
        ' 24 MATERIAL
        '=========================================================
        Private ReadOnly MaterialNames As String() =
            BuildMaterialNames()


        Private Function BuildMaterialNames() As String()

            Dim list As New List(Of String)

            list.Add("Steel, Mild")

            For i As Integer = 1 To 23
                list.Add("Steel, Mild " & i.ToString())
            Next

            Return list.ToArray()

        End Function


        '=========================================================
        ' MAIN
        '=========================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try

                Dim oApp As Inventor.Application =
                    g_inventorApplication

                If oApp Is Nothing Then Return


                Dim oDoc As Document =
                    oApp.ActiveDocument

                If oDoc Is Nothing Then Return


                Dim oAssDoc As AssemblyDocument =
                    TryCast(oDoc, AssemblyDocument)

                If oAssDoc Is Nothing Then Return


                Dim rand As New Random()


                '=================================================
                ' PART ĐÃ XỬ LÝ
                '=================================================
                Dim processedParts As New HashSet(Of String)(
                    StringComparer.OrdinalIgnoreCase)


                '=================================================
                ' QUÉT TOÀN BỘ ASSEMBLY
                '=================================================
                ProcessOccurrences(
                    oApp,
                    oAssDoc.ComponentDefinition.Occurrences,
                    rand,
                    processedParts)


                Try
                    oAssDoc.Update()
                Catch
                End Try


                PostStatus(
                    "Đã random Material cho Part thường.")

            Catch

                ' Không hiện MsgBox
                Return

            End Try

        End Sub


        '=========================================================
        ' XỬ LÝ OCCURRENCES
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


                    '=================================================
                    ' BỎ QUA SUPPRESS
                    '=================================================
                    Try

                        If occ.Suppressed Then
                            Continue For
                        End If

                    Catch
                    End Try


                    '=================================================
                    ' BỎ QUA STANDARD / CONTENT CENTER
                    '=================================================
                    Try

                        If occ.IsContentCenterMember Then
                            Continue For
                        End If

                    Catch
                    End Try


                    '=================================================
                    ' KIỂM TRA BOM STRUCTURE
                    '
                    ' PURCHASED
                    ' PHANTOM
                    '=================================================
                    Dim bomStructure As BOMStructureEnum =
                        BOMStructureEnum.kNormalBOMStructure


                    Try

                        bomStructure =
                            occ.Definition.BOMStructure

                    Catch

                        Try
                            bomStructure = occ.BOMStructure
                        Catch
                        End Try

                    End Try


                    '=================================================
                    ' BỎ QUA PURCHASED
                    '=================================================
                    If bomStructure =
                        BOMStructureEnum.kPurchasedBOMStructure Then

                        Continue For

                    End If


                    '=================================================
                    ' BỎ QUA PHANTOM
                    '=================================================
                    If bomStructure =
                        BOMStructureEnum.kPhantomBOMStructure Then

                        Continue For

                    End If


                    '=================================================
                    ' LẤY DOCUMENT
                    '=================================================
                    Dim occDoc As Document = Nothing

                    Try

                        occDoc =
                            occ.Definition.Document

                    Catch

                        occDoc = Nothing

                    End Try


                    If occDoc Is Nothing Then
                        Continue For
                    End If


                    '=================================================
                    ' PART
                    '=================================================
                    Dim partDoc As PartDocument =
                        TryCast(occDoc, PartDocument)


                    If partDoc IsNot Nothing Then

                        ProcessPart(
                            oApp,
                            partDoc,
                            rand,
                            processedParts)

                        Continue For

                    End If


                    '=================================================
                    ' SUB ASSEMBLY
                    '=================================================
                    Dim subAssDoc As AssemblyDocument =
                        TryCast(occDoc, AssemblyDocument)


                    If subAssDoc IsNot Nothing Then

                        Try

                            ProcessOccurrences(
                                oApp,
                                subAssDoc.
                                    ComponentDefinition.
                                    Occurrences,
                                rand,
                                processedParts)

                        Catch

                            ' SubAssembly lỗi -> bỏ qua

                        End Try

                    End If


                Catch

                    '=================================================
                    ' OCCURRENCE LỖI -> BỎ QUA
                    '=================================================
                    Continue For

                End Try

            Next

        End Sub


        '=========================================================
        ' XỬ LÝ 1 PART
        '=========================================================
        Private Sub ProcessPart(
            ByVal oApp As Inventor.Application,
            ByVal partDoc As PartDocument,
            ByVal rand As Random,
            ByVal processedParts As HashSet(Of String))


            Try

                If partDoc Is Nothing Then Return


                '=================================================
                ' KIỂM TRA CONTENT CENTER LẦN NỮA
                '=================================================
                Try

                    If partDoc.ComponentDefinition.IsContentCenterMember Then
                        Return
                    End If

                Catch
                End Try


                '=================================================
                ' KIỂM TRA BOM STRUCTURE CỦA PART
                '=================================================
                Try

                    Dim bom As BOMStructureEnum =
                        partDoc.ComponentDefinition.BOMStructure


                    If bom =
                        BOMStructureEnum.kPurchasedBOMStructure Then
                        Return
                    End If


                    If bom =
                        BOMStructureEnum.kPhantomBOMStructure Then
                        Return
                    End If

                Catch
                End Try


                '=================================================
                ' LẤY TÊN FILE
                '=================================================
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


                '=================================================
                ' PART ĐÃ XỬ LÝ
                '=================================================
                If processedParts.Contains(fileName) Then
                    Return
                End If


                processedParts.Add(fileName)


                '=================================================
                ' RANDOM 1 TRONG 24 MATERIAL
                '=================================================
                Dim index As Integer =
                    rand.Next(
                        0,
                        MaterialNames.Length)


                Dim materialName As String =
                    MaterialNames(index)


                '=================================================
                ' TÌM MATERIAL
                '=================================================
                Dim materialAsset As MaterialAsset =
                    GetMaterialAsset(
                        oApp,
                        partDoc,
                        materialName)


                If materialAsset Is Nothing Then
                    Return
                End If


                '=================================================
                ' GÁN MATERIAL
                '=================================================
                Try

                    partDoc.ActiveMaterial =
                        materialAsset

                Catch

                    Return

                End Try


                '=================================================
                ' LẤY APPEARANCE ĐI KÈM MATERIAL
                '=================================================
                Try

                    Dim matAppearance As Asset =
                        materialAsset.AppearanceAsset


                    If matAppearance IsNot Nothing Then

                        Dim localAppearance As Asset =
                            FindAppearance(
                                partDoc,
                                matAppearance.DisplayName)


                        If localAppearance Is Nothing Then

                            localAppearance =
                                CopyAppearanceToDocument(
                                    oApp,
                                    partDoc,
                                    matAppearance.DisplayName)

                        End If


                        If localAppearance IsNot Nothing Then

                            Try

                                partDoc.ActiveAppearance =
                                    localAppearance

                            Catch
                            End Try

                        End If

                    End If

                Catch

                    ' Appearance lỗi -> bỏ qua

                End Try


                '=================================================
                ' DÙNG APPEARANCE CỦA MATERIAL
                '=================================================
                Try

                    partDoc.AppearanceSourceType =
                        AppearanceSourceTypeEnum.kMaterialAppearance

                Catch
                End Try


                '=================================================
                ' UPDATE
                '=================================================
                Try
                    partDoc.Update()
                Catch
                End Try


            Catch

                ' Part lỗi -> bỏ qua
                Return

            End Try

        End Sub


        '=========================================================
        ' TÌM MATERIAL
        '=========================================================
        Private Function GetMaterialAsset(
            ByVal oApp As Inventor.Application,
            ByVal oPartDoc As PartDocument,
            ByVal materialName As String) As MaterialAsset


            '=====================================================
            ' 1. DOCUMENT MATERIALS
            '=====================================================
            Try

                For Each mat As MaterialAsset In
                    oPartDoc.MaterialAssets

                    Try

                        If mat Is Nothing Then
                            Continue For
                        End If


                        If String.Equals(
                            mat.DisplayName,
                            materialName,
                            StringComparison.OrdinalIgnoreCase) Then

                            Return mat

                        End If

                    Catch
                    End Try

                Next

            Catch
            End Try


            '=====================================================
            ' 2. MATERIAL LIBRARIES
            '=====================================================
            Try

                For Each libqq As AssetLibrary In
                    oApp.AssetLibraries

                    Try

                        Dim libraryMaterial As MaterialAsset =
                            Nothing


                        '-----------------------------------------
                        ' TÌM MATERIAL
                        '-----------------------------------------
                        For Each mat As MaterialAsset In
                            libqq.MaterialAssets

                            Try

                                If mat Is Nothing Then
                                    Continue For
                                End If


                                If String.Equals(
                                    mat.DisplayName,
                                    materialName,
                                    StringComparison.OrdinalIgnoreCase) Then

                                    libraryMaterial = mat
                                    Exit For

                                End If

                            Catch
                            End Try

                        Next


                        If libraryMaterial Is Nothing Then
                            Continue For
                        End If


                        '-----------------------------------------
                        ' COPY VÀO DOCUMENT
                        '-----------------------------------------
                        Try

                            Dim copied As Asset =
                                libraryMaterial.CopyTo(
                                    oPartDoc)


                            If copied IsNot Nothing Then

                                Dim copiedMaterial As MaterialAsset =
                                    TryCast(
                                        copied,
                                        MaterialAsset)


                                If copiedMaterial IsNot Nothing Then
                                    Return copiedMaterial
                                End If

                            End If

                        Catch
                        End Try


                        '-----------------------------------------
                        ' TÌM LẠI TRONG DOCUMENT
                        '-----------------------------------------
                        Try

                            For Each mat As MaterialAsset In
                                oPartDoc.MaterialAssets

                                If String.Equals(
                                    mat.DisplayName,
                                    materialName,
                                    StringComparison.OrdinalIgnoreCase) Then

                                    Return mat

                                End If

                            Next

                        Catch
                        End Try


                    Catch

                        ' Library lỗi -> bỏ qua

                    End Try

                Next

            Catch
            End Try


            Return Nothing

        End Function


        '=========================================================
        ' TÌM APPEARANCE TRONG DOCUMENT
        '=========================================================
        Private Function FindAppearance(
            ByVal oPartDoc As PartDocument,
            ByVal appearanceName As String) As Asset


            Try

                For Each appAsset As Asset In
                    oPartDoc.AppearanceAssets

                    Try

                        If appAsset Is Nothing Then
                            Continue For
                        End If


                        If String.Equals(
                            appAsset.DisplayName,
                            appearanceName,
                            StringComparison.OrdinalIgnoreCase) Then

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
        ' COPY APPEARANCE VÀO DOCUMENT
        '=========================================================
        Private Function CopyAppearanceToDocument(
            ByVal oApp As Inventor.Application,
            ByVal oPartDoc As PartDocument,
            ByVal appearanceName As String) As Asset


            Try

                For Each libqq As AssetLibrary In
                    oApp.AssetLibraries

                    Try

                        For Each appAsset As Asset In
                            libqq.AppearanceAssets

                            Try

                                If appAsset Is Nothing Then
                                    Continue For
                                End If


                                If String.Equals(
                                    appAsset.DisplayName,
                                    appearanceName,
                                    StringComparison.OrdinalIgnoreCase) Then


                                    Try

                                        Dim copied As Asset =
                                            appAsset.CopyTo(
                                                oPartDoc)


                                        If copied IsNot Nothing Then
                                            Return copied
                                        End If

                                    Catch
                                    End Try


                                    '---------------------------------
                                    ' Tìm lại trong Document
                                    '---------------------------------
                                    Dim localAsset As Asset =
                                        FindAppearance(
                                            oPartDoc,
                                            appearanceName)


                                    If localAsset IsNot Nothing Then
                                        Return localAsset
                                    End If


                                End If

                            Catch
                            End Try

                        Next

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

                g_inventorApplication.
                    UserInterfaceManager.
                    UserInteractionManager.
                    PostStatus(msg)

            Catch
            End Try

        End Sub

    End Module

End Namespace