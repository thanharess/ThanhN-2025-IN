Option Explicit On
Option Strict Off

Imports Inventor
Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms

Namespace ToolInventor2025.Part.Buttons

    Public Module Part_Solid_7e

        '=========================================================
        ' 24 MATERIAL
        '
        ' Steel, Mild
        ' Steel, Mild 1
        ' ...
        ' Steel, Mild 23
        '=========================================================
        Private ReadOnly MaterialNames As String() = BuildMaterialNames()


        '=========================================================
        ' TẠO DANH SÁCH MATERIAL
        '=========================================================
        Private Function BuildMaterialNames() As String()

            Dim list As New List(Of String)

            ' Steel, Mild
            list.Add("Steel, Mild")

            ' Steel, Mild 1 -> Steel, Mild 23
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

                '-------------------------------------------------
                ' LẤY INVENTOR
                '-------------------------------------------------
                Dim oApp As Inventor.Application =
                    g_inventorApplication

                If oApp Is Nothing Then

                    MessageBox.Show(
                        "Không lấy được g_inventorApplication.",
                        "Part Solid",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Return

                End If


                '-------------------------------------------------
                ' DOCUMENT HIỆN TẠI
                '-------------------------------------------------
                Dim oDoc As Document =
                    oApp.ActiveEditDocument

                If oDoc Is Nothing Then

                    PostStatus(
                        "Không có tài liệu đang mở.")

                    Return

                End If


                '-------------------------------------------------
                ' PHẢI LÀ PART
                '-------------------------------------------------
                Dim oPartDoc As PartDocument =
                    TryCast(oDoc, PartDocument)

                If oPartDoc Is Nothing Then

                    MessageBox.Show(
                        "Vui lòng mở file Part (.ipt).",
                        "Part Solid",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Return

                End If


                '=================================================
                ' RANDOM 1 TRONG 24 MATERIAL
                '=================================================
                Dim rand As New Random()

                Dim index As Integer =
                    rand.Next(0, MaterialNames.Length)

                Dim chosenMaterialName As String =
                    MaterialNames(index)


                '=================================================
                ' TÌM MATERIAL
                '=================================================
                Dim materialAsset As MaterialAsset =
                    GetMaterialAsset(
                        oApp,
                        oPartDoc,
                        chosenMaterialName)


                If materialAsset Is Nothing Then

                    MessageBox.Show(
                        "Không tìm thấy Material:" &
                        vbCrLf &
                        chosenMaterialName &
                        vbCrLf & vbCrLf &
                        "Kiểm tra Inventor Material Library.",
                        "Part Solid",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Return

                End If


                '=================================================
                ' GÁN MATERIAL
                '=================================================
                Try

                    oPartDoc.ActiveMaterial =
                        materialAsset

                Catch ex As Exception

                    MessageBox.Show(
                        "Không thể gán Material:" &
                        vbCrLf &
                        chosenMaterialName &
                        vbCrLf & vbCrLf &
                        ex.Message,
                        "Part Solid",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

                    Return

                End Try


                '=================================================
                ' LẤY APPEARANCE ĐI KÈM MATERIAL
                '=================================================
                Try

                    Dim materialAppearance As Asset =
                        materialAsset.AppearanceAsset

                    If materialAppearance IsNot Nothing Then

                        '-------------------------------------------------
                        ' Appearance phải thuộc cùng Document
                        '-------------------------------------------------
                        Dim localAppearance As Asset =
                            FindAppearance(
                                oPartDoc,
                                materialAppearance.DisplayName)


                        If localAppearance IsNot Nothing Then

                            oPartDoc.ActiveAppearance =
                                localAppearance

                        Else

                            '-------------------------------------------------
                            ' Nếu appearance chưa có trong Document
                            ' thì copy từ library
                            '-------------------------------------------------
                            Dim copiedAppearance As Asset =
                                CopyAppearanceToDocument(
                                    oApp,
                                    oPartDoc,
                                    materialAppearance.DisplayName)


                            If copiedAppearance IsNot Nothing Then

                                oPartDoc.ActiveAppearance =
                                    copiedAppearance

                            End If

                        End If

                    End If

                Catch
                    ' Material vẫn được gán,
                    ' không dừng chương trình nếu appearance lỗi.
                End Try


                '=================================================
                ' ĐẢM BẢO PART DÙNG APPEARANCE CỦA MATERIAL
                '=================================================
                Try

                    oPartDoc.AppearanceSourceType =
                        AppearanceSourceTypeEnum.kMaterialAppearance

                Catch
                End Try


                '=================================================
                ' UPDATE
                '=================================================
                Try
                    oPartDoc.Update()
                Catch
                End Try


                '=================================================
                ' THÔNG BÁO
                '=================================================
                PostStatus(
                    "Đã random Material: " &
                    chosenMaterialName)


            Catch ex As Exception

                MessageBox.Show(
                    "Lỗi Part_Solid_7e:" &
                    vbCrLf & vbCrLf &
                    ex.Message &
                    vbCrLf & vbCrLf &
                    ex.StackTrace,
                    "Inventor 2020",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub


        '=========================================================
        ' TÌM MATERIAL TRONG DOCUMENT
        ' SAU ĐÓ TÌM TRONG CÁC LIBRARY
        '=========================================================
        Private Function GetMaterialAsset(
            ByVal oApp As Inventor.Application,
            ByVal oPartDoc As PartDocument,
            ByVal materialName As String) As MaterialAsset


            '=====================================================
            ' 1. TÌM TRONG DOCUMENT MATERIALS
            '=====================================================
            Try

                For Each mat As MaterialAsset In
                    oPartDoc.MaterialAssets

                    If mat Is Nothing Then
                        Continue For
                    End If

                    If String.Equals(
                        mat.DisplayName,
                        materialName,
                        StringComparison.OrdinalIgnoreCase) Then

                        Return mat

                    End If

                Next

            Catch
            End Try


            '=====================================================
            ' 2. TÌM TRONG MATERIAL LIBRARIES
            '=====================================================
            Try

                For Each liba As AssetLibrary In
                    oApp.AssetLibraries

                    Try

                        Dim libraryMaterial As MaterialAsset =
                            Nothing


                        '-----------------------------------------
                        ' Tìm theo DisplayName
                        '-----------------------------------------
                        For Each mat As MaterialAsset In
                            liba.MaterialAssets

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

                        Next


                        If libraryMaterial Is Nothing Then
                            Continue For
                        End If


                        '-----------------------------------------
                        ' COPY MATERIAL VÀO DOCUMENT
                        '
                        ' CopyTo(Document)
                        ' sẽ copy material cùng các asset
                        ' liên quan vào document.
                        '-----------------------------------------
                        Try

                            Dim copied As Asset =
                                libraryMaterial.CopyTo(
                                    oPartDoc)

                            If copied IsNot Nothing Then

                                Dim copiedMaterial As MaterialAsset =
                                    TryCast(copied, MaterialAsset)

                                If copiedMaterial IsNot Nothing Then
                                    Return copiedMaterial
                                End If

                            End If

                        Catch

                            ' Có thể material đã tồn tại
                            ' trong document.
                        End Try


                        '-----------------------------------------
                        ' Tìm lại trong Document
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
                        ' Library không đọc được -> bỏ qua
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

                    If appAsset Is Nothing Then
                        Continue For
                    End If

                    If String.Equals(
                        appAsset.DisplayName,
                        appearanceName,
                        StringComparison.OrdinalIgnoreCase) Then

                        Return appAsset

                    End If

                Next

            Catch
            End Try

            Return Nothing

        End Function


        '=========================================================
        ' COPY APPEARANCE TỪ LIBRARY VÀO DOCUMENT
        '=========================================================
        Private Function CopyAppearanceToDocument(
            ByVal oApp As Inventor.Application,
            ByVal oPartDoc As PartDocument,
            ByVal appearanceName As String) As Asset


            Try

                '-------------------------------------------------
                ' 1. TÌM TRONG CÁC LIBRARY
                '-------------------------------------------------
                For Each liba As AssetLibrary In
                    oApp.AssetLibraries

                    Try

                        For Each appAsset As Asset In
                            liba.AppearanceAssets

                            If appAsset Is Nothing Then
                                Continue For
                            End If


                            If String.Equals(
                                appAsset.DisplayName,
                                appearanceName,
                                StringComparison.OrdinalIgnoreCase) Then


                                '---------------------------------
                                ' COPY VÀO DOCUMENT
                                '---------------------------------
                                Try

                                    Dim copied As Asset =
                                        appAsset.CopyTo(
                                            oPartDoc)

                                    If copied IsNot Nothing Then
                                        Return copied
                                    End If

                                Catch

                                    ' Có thể đã tồn tại
                                End Try


                                '---------------------------------
                                ' Tìm lại
                                '---------------------------------
                                Dim localAsset As Asset =
                                    FindAppearance(
                                        oPartDoc,
                                        appearanceName)

                                If localAsset IsNot Nothing Then
                                    Return localAsset
                                End If


                            End If

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