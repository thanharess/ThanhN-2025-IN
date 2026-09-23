
Option Explicit On
Option Strict Off

Imports System
Imports System.Runtime.InteropServices

Namespace ToolInventor2025.Assembly

    Public Module CommonKeywords

        '==========================================================
        ' BEARING
        '==========================================================
        Public ReadOnly BearingKeywords As String() = {
            "vòng bi", "vong bi", "bearing", "motor", "gối bi", "goi bi", "gối đỡ", "goi do", "pillow",
            "plummer", "ucp", "ucf", "ucfl", "khóa trục", "khoa truc", "ntn", "mortor", "mottor", "UCT"}
        '==========================================================
        ' FASTENER
        '==========================================================
        Public ReadOnly FastenerKeywords As String() = {
            "bulong",
            "bu lông", "bu long", "ốc", "oc", "đai ốc", "dai oc", "vít", "vit", "ecu",
            "êcu", "then", "then chốt", "long đen", "lông đền", "long den", "washer",
            "iso", "din", "jis", "m3", "m4", "m5", "m6", "m8", "lock collar",
            "locknut", "lock nut", "m10", "m12", "m16", "then hoa",
            "m20", "m24", "m30", "m36", "m42", "m48", "ss 2", "iso 4", "din 125", "din 127",
            "din 933", "din 934", "din 6912"}

        '==========================================================
        ' STANDARD
        '==========================================================
        Public ReadOnly StandardKeywords As String() = {
            "ISO", "DIN", "SKF", "SS", "GB", "JIS", "ANSI", "BSI", "GOST", "ASTM", "UPE"}

        '==========================================================
        ' BEARING
        ' SO SÁNH TỪ KÝ TỰ ĐẦU TIÊN
        '==========================================================
        Public Function IsBearing(ByVal text As String) As Boolean

            If String.IsNullOrWhiteSpace(text) Then
                Return False
            End If

            Dim s As String = text.Trim().ToLowerInvariant()

            For Each kw As String In BearingKeywords

                If String.IsNullOrWhiteSpace(kw) Then
                    Continue For
                End If

                If s.StartsWith(
                    kw.Trim().ToLowerInvariant(),
                    StringComparison.OrdinalIgnoreCase) Then

                    Return True

                End If

            Next

            Return False

        End Function

        '==========================================================
        ' FASTENER
        ' SO SÁNH TỪ KÝ TỰ ĐẦU TIÊN
        '==========================================================
        Public Function IsFastener(ByVal text As String) As Boolean

            If String.IsNullOrWhiteSpace(text) Then
                Return False
            End If

            Dim s As String = text.Trim().ToLowerInvariant()

            For Each kw As String In FastenerKeywords

                If String.IsNullOrWhiteSpace(kw) Then
                    Continue For
                End If

                If s.StartsWith(
                    kw.Trim().ToLowerInvariant(),
                    StringComparison.OrdinalIgnoreCase) Then

                    Return True

                End If

            Next

            Return False

        End Function

        '==========================================================
        ' STANDARD
        ' SO SÁNH TỪ KÝ TỰ ĐẦU TIÊN
        ' KHÔNG DÙNG CONTAINS
        '==========================================================
        Public Function IsStandardKeyword(ByVal pn As String) As Boolean

            If String.IsNullOrWhiteSpace(pn) Then
                Return False
            End If

            Dim s As String = pn.Trim().ToUpperInvariant()

            For Each kw As String In StandardKeywords

                If String.IsNullOrWhiteSpace(kw) Then
                    Continue For
                End If

                If s.StartsWith(
                    kw.Trim().ToUpperInvariant(),
                    StringComparison.OrdinalIgnoreCase) Then

                    Return True

                End If

            Next

            Return False

        End Function

    End Module

End Namespace
