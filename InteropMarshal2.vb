Imports System.Runtime.InteropServices

Namespace Interop

    Public NotInheritable Class Marshal2
        Private Sub New()
        End Sub

        ' ---- Chuyển ProgID -> CLSID ----
        <DllImport("ole32.dll", CharSet:=CharSet.Unicode,
                   ExactSpelling:=True, PreserveSig:=True)>
        Private Shared Function CLSIDFromProgIDEx(
            <MarshalAs(UnmanagedType.LPWStr)> ByVal lpszProgID As String,
            <Out> ByRef lpclsid As System.Guid) As Integer
        End Function

        ' ---- Lấy instance đang chạy của COM object ----
        ' Đổi tên thành GetActiveObjectNative để tránh nhầm với hàm public bên dưới.
        ' Sửa dòng DllImport này:
        <DllImport("oleaut32.dll", EntryPoint:="GetActiveObject", ExactSpelling:=True, PreserveSig:=True)>
        Private Shared Function GetActiveObjectNative(
    ByRef rclsid As System.Guid,
    ByVal pvReserved As IntPtr,
    <Out> ByRef ppunk As IntPtr) As Integer
        End Function

        ''' <summary>
        ''' Tương đương Marshal.GetActiveObject của .NET Framework.
        ''' Trả về Nothing nếu không tìm thấy instance nào đang chạy.
        ''' </summary>
        Public Shared Function GetActiveObject(progId As String,
                                               Optional throwOnError As Boolean = False) As Object
            If progId Is Nothing Then Throw New ArgumentNullException(NameOf(progId))

            ' Bước 1: chuyển ProgID thành CLSID
            Dim clsid As System.Guid
            Dim hr As Integer = CLSIDFromProgIDEx(progId, clsid)
            If hr < 0 Then
                If throwOnError Then Marshal.ThrowExceptionForHR(hr)
                Return Nothing
            End If

            ' Bước 2: lấy con trỏ IUnknown của instance đang chạy
            Dim pUnk As IntPtr = IntPtr.Zero
            hr = GetActiveObjectNative(clsid, IntPtr.Zero, pUnk)
            If hr < 0 Then
                If throwOnError Then Marshal.ThrowExceptionForHR(hr)
                Return Nothing
            End If

            ' Bước 3: bọc con trỏ native thành đối tượng RCW và giải phóng con trỏ
            Try
                Return Marshal.GetObjectForIUnknown(pUnk)
            Finally
                Marshal.Release(pUnk)
            End Try
        End Function

    End Class

End Namespace