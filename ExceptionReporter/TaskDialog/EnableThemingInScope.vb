
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Security.Permissions
Imports Debug = System.Diagnostics.Debug
Imports SuppressMessage = System.Diagnostics.CodeAnalysis.SuppressMessageAttribute

Namespace Global.System.Windows.Forms
    '''//// <summary>
    '''////   This class is intended to be used with the 'using' statement, to activate an activation
    '''////   context for turning on visual theming at the beginning of a scope, and have it
    '''////   automatically deactivated when the scope is exited.
    '''//// </summary>
    '''//// <remarks>
    '''////   With a little bit of tuning, this old traditional model, still works even in Windows 7. In
    '''////   our case we need it to apply theming to the native Print dialog of the view. It may be
    '''////   useless in the future.
    '''//// </remarks>
    '''/[SuppressUnmanagedCodeSecurity()]
    '''/internal class EnableThemingInScope : IDisposable
    '''/{
    '''/    /*
    '''/     * This class is for manually loading a specific version of the ComCtl32.dll needed to display
    '''/     * TaskDialog and enable theming support.
    '''/     *
    '''/     * Code derived from http://support.microsoft.com/kb/830033/
    '''/     *
    '''/     * Source:  https://github.com/khrona/AwesomiumSharp/blob/cc91ce8885/AwesomiumSharp/EnableThemingInScope.cs
    '''/     * Updates: Correct ACTCTX implementation from https://github.com/wrouesnel/keepass/blob/3c82c82/KeePass/Native/NativeMethods.Structs.cs
    '''/     */


    Friend NotInheritable Class EnableThemingInScope
        Implements IDisposable
        '
        '         * This class is for manually loading a specific version of the ComCtl32.dll needed to display
        '         * TaskDialog and enable theming support.
        '         *
        '         * Source:  Composite of the relevant classes and members from the following files in the
        '         *          https://github.com/wrouesnel/keepass repository:
        '         *
        '         *          https://github.com/wrouesnel/keepass/blob/3c82c82/KeePass/Native/NativeMethods.cs
        '         *          https://github.com/wrouesnel/keepass/blob/3c82c82/KeePass/Native/NativeMethods.Defs.cs
        '         *          https://github.com/wrouesnel/keepass/blob/3c82c82/KeePass/Native/NativeMethods.Structs.cs
        '         *          https://github.com/wrouesnel/keepass/blob/3c82c82/KeePass/UI/EnableThemingInScope.cs
        '         *
        '         * Commit Details
        '         * Full Hash: 3c82c82f727d5c66cd331bd6e22dce1aefbd0cf6
        '         * Date:      26 March 2016
        '         * User:      wrouesnel - https://github.com/wrouesnel
        '

        ' Code derived from http://support.microsoft.com/kb/830033/



        Private Shared m_nhCtx As IntPtr? = Nothing
        Private Shared m_oSync As New Object()
        Private m_nuCookie As UIntPtr? = Nothing

        Public Sub New()
            Try
                If OSFeature.Feature.IsPresent(OSFeature.Themes) Then
                    If EnsureActCtxCreated() Then
                        Dim u As UIntPtr = UIntPtr.Zero
                        If NativeMethods.ActivateActCtx(m_nhCtx.Value, u) Then
                            m_nuCookie = u
                        End If
                    End If
                End If
            Catch generatedExceptionName As Exception
                Debug.Fail("Exception encountered while initializing " + NameOf(EnableThemingInScope))
                Debug.Assert(False)
            End Try
        End Sub

        Protected Overrides Sub Finalize()
            Try
                Debug.Assert(Not m_nuCookie.HasValue)
                Dispose(False)
            Finally
                MyBase.Finalize()
            End Try
        End Sub

        Public Shared Sub StaticDispose()
            If Not m_nhCtx.HasValue Then
                Return
            End If

            Try
                NativeMethods.ReleaseActCtx(m_nhCtx.Value)
                m_nhCtx = Nothing
            Catch generatedExceptionName As Exception
                Debug.Fail("Exception encountered while static disposing " + NameOf(EnableThemingInScope))
                Debug.Assert(False)
            End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Private Shared Function EnsureActCtxCreated() As Boolean
            SyncLock m_oSync
                If m_nhCtx.HasValue Then
                    Return True
                End If

                Dim strAsmLoc As String
                Dim p As New FileIOPermission(PermissionState.None)
                p.AllFiles = FileIOPermissionAccess.PathDiscovery
                p.Assert()
                Try
                    strAsmLoc = GetType(Object).Assembly.Location
                Finally
                    CodeAccessPermission.RevertAssert()
                End Try
                If String.IsNullOrEmpty(strAsmLoc) Then
                    Debug.Assert(False)
                    Return False
                End If

                Dim strInstDir As String = Path.GetDirectoryName(strAsmLoc)
                Dim strMfLoc As String = Path.Combine(strInstDir, "XPThemes.manifest")

                Dim ctx As New NativeMethods.ACTCTX()
                ctx.cbSize = CUInt(Marshal.SizeOf(GetType(NativeMethods.ACTCTX)))
                Debug.Assert(((IntPtr.Size = 4) AndAlso (ctx.cbSize = NativeMethods.ACTCTXSize32)) OrElse ((IntPtr.Size = 8) AndAlso (ctx.cbSize = NativeMethods.ACTCTXSize64)))

                ctx.lpSource = strMfLoc
                ctx.lpAssemblyDirectory = strInstDir
                ctx.dwFlags = NativeMethods.ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID

                m_nhCtx = NativeMethods.CreateActCtx(ctx)
                If NativeMethods.IsInvalidHandleValue(m_nhCtx.Value) Then
                    Debug.Assert(False)
                    m_nhCtx = Nothing
                    Return False
                End If
            End SyncLock

            Return True
        End Function

        <SuppressMessage("Redundancies in Symbol Declarations", "RECS0154:Parameter is never used", Justification:="Status-by-design")>
        Private Sub Dispose(bDisposing As Boolean)
            If Not m_nuCookie.HasValue Then
                Return
            End If

            Try
                If NativeMethods.DeactivateActCtx(0, m_nuCookie.Value) Then
                    m_nuCookie = Nothing
                End If
            Catch generatedExceptionName As Exception
                Debug.Assert(False)
            End Try
        End Sub

        Friend NotInheritable Class NativeMethods
            Private Sub New()
            End Sub
            Friend Const ACTCTXSize32 As UInteger = 32
            Friend Const ACTCTXSize64 As UInteger = 56
            Friend Const ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID As UInteger = &H4

            <DllImport("Kernel32.dll")>
            Friend Shared Function ActivateActCtx(hActCtx As IntPtr, ByRef lpCookie As UIntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
            End Function

            <DllImport("Kernel32.dll", CharSet:=CharSet.Auto)>
            Friend Shared Function CreateActCtx(ByRef pActCtx As ACTCTX) As IntPtr
            End Function

            <DllImport("Kernel32.dll")>
            Friend Shared Function DeactivateActCtx(dwFlags As UInteger, ulCookie As UIntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
            End Function

            <DllImport("Kernel32.dll")>
            Friend Shared Sub ReleaseActCtx(hActCtx As IntPtr)
            End Sub

            Public Shared Function IsInvalidHandleValue(p As IntPtr) As Boolean
                Dim h As Long = p.ToInt64()
                If h = -1 Then
                    Return True
                End If
                If h = &HFFFFFFFFUI Then
                    Return True
                End If

                Return False
            End Function

            <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
            Friend Structure ACTCTX
                Public cbSize As UInteger
                Public dwFlags As UInteger

                ' Not LPWStr, see source code
                <MarshalAs(UnmanagedType.LPTStr)>
                Public lpSource As String

                Public wProcessorArchitecture As UShort
                Public wLangId As UShort

                <MarshalAs(UnmanagedType.LPTStr)>
                Public lpAssemblyDirectory As String

                <MarshalAs(UnmanagedType.LPTStr)>
                Public lpResourceName As String

                <MarshalAs(UnmanagedType.LPTStr)>
                Public lpApplicationName As String

                Public hModule As IntPtr
            End Structure
        End Class
    End Class
End Namespace

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================
