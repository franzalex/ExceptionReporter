
'------------------------------------------------------------------
' <summary>
' A P/Invoke wrapper for TaskDialog. Usability was given preference to perf and size.
' </summary>
'
' <remarks/>
'------------------------------------------------------------------

Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Diagnostics.CodeAnalysis
Namespace Global.Microsoft.Samples

    ''' <summary>Class to hold native code interop declarations.</summary>
    Partial Friend NotInheritable Class UnsafeNativeMethods
        Private Sub New()
        End Sub
        ''' <summary>WM_USER taken from WinUser.h</summary>
        Friend Const WM_USER As UInteger = &H400

        ''' <summary>
        '''   The signature of the callback that receives messages from the Task Dialog when various
        '''   events occur.
        ''' </summary>
        ''' <param name="hwnd">The window handle of the</param>
        ''' <param name="msg">The message being passed.</param>
        ''' <param name="wParam">wParam which is interpreted differently depending on the message.</param>
        ''' <param name="lParam">wParam which is interpreted differently depending on the message.</param>
        ''' <param name="refData">The refrence data that was set to TaskDialog.CallbackData.</param>
        ''' <returns>A HRESULT value. The return value is specific to the message being processed.</returns>
        Friend Delegate Function TaskDialogCallback(<[In]> hwnd As IntPtr, <[In]> msg As UInteger, <[In]> wParam As UIntPtr, <[In]> lParam As IntPtr, <[In]> refData As IntPtr) As Integer

        ''' <summary>TASKDIALOG_FLAGS taken from CommCtrl.h.</summary>
        <Flags>
        Friend Enum TASKDIALOG_FLAGS
            ''' <summary>Enable hyperlinks.</summary>
            TDF_ENABLE_HYPERLINKS = &H1

            ''' <summary>Use icon handle for main icon.</summary>
            TDF_USE_HICON_MAIN = &H2

            ''' <summary>Use icon handle for footer icon.</summary>
            TDF_USE_HICON_FOOTER = &H4

            ''' <summary>Allow dialog to be cancelled, even if there is no cancel button.</summary>
            TDF_ALLOW_DIALOG_CANCELLATION = &H8

            ''' <summary>Use command links rather than buttons.</summary>
            TDF_USE_COMMAND_LINKS = &H10

            ''' <summary>Use command links with no icons rather than buttons.</summary>
            TDF_USE_COMMAND_LINKS_NO_ICON = &H20

            ''' <summary>Show expanded info in the footer area.</summary>
            TDF_EXPAND_FOOTER_AREA = &H40

            ''' <summary>Expand by default.</summary>
            TDF_EXPANDED_BY_DEFAULT = &H80

            ''' <summary>Start with verification flag already checked.</summary>
            TDF_VERIFICATION_FLAG_CHECKED = &H100

            ''' <summary>Show a progress bar.</summary>
            TDF_SHOW_PROGRESS_BAR = &H200

            ''' <summary>Show a marquee progress bar.</summary>
            TDF_SHOW_MARQUEE_PROGRESS_BAR = &H400

            ''' <summary>Callback every 200 milliseconds.</summary>
            TDF_CALLBACK_TIMER = &H800

            ''' <summary>Center the dialog on the owner window rather than the monitor.</summary>
            TDF_POSITION_RELATIVE_TO_WINDOW = &H1000

            ''' <summary>Right to Left Layout.</summary>
            TDF_RTL_LAYOUT = &H2000

            ''' <summary>No default radio button.</summary>
            TDF_NO_DEFAULT_RADIO_BUTTON = &H4000

            ''' <summary>Task Dialog can be minimized.</summary>
            TDF_CAN_BE_MINIMIZED = &H8000
        End Enum

        ''' <summary>TASKDIALOG_ELEMENTS taken from CommCtrl.h</summary>
        Friend Enum TASKDIALOG_ELEMENTS
            ''' <summary>The content element.</summary>
            TDE_CONTENT

            ''' <summary>Expanded Information.</summary>
            TDE_EXPANDED_INFORMATION

            ''' <summary>Footer.</summary>
            TDE_FOOTER

            ''' <summary>Main Instructions</summary>
            TDE_MAIN_INSTRUCTION
        End Enum

        ''' <summary>TASKDIALOG_ICON_ELEMENTS taken from CommCtrl.h</summary>
        Friend Enum TASKDIALOG_ICON_ELEMENTS
            ''' <summary>Main instruction icon.</summary>
            TDIE_ICON_MAIN

            ''' <summary>Footer icon.</summary>
            TDIE_ICON_FOOTER
        End Enum

        ''' <summary>TASKDIALOG_MESSAGES taken from CommCtrl.h.</summary>
        Friend Enum TASKDIALOG_MESSAGES As UInteger
            ' Spec is not clear on what this is for.
            ''''' <summary>Navigate page.</summary>
            ''TDM_NAVIGATE_PAGE = WM_USER + 101

            ''' <summary>Click button.</summary>
            TDM_CLICK_BUTTON = WM_USER + 102
            ' wParam = Button ID
            ''' <summary>Set Progress bar to be marquee mode.</summary>
            TDM_SET_MARQUEE_PROGRESS_BAR = WM_USER + 103
            ' wParam = 0 (nonMarque) wParam != 0 (Marquee)
            ''' <summary>Set Progress bar state.</summary>
            TDM_SET_PROGRESS_BAR_STATE = WM_USER + 104
            ' wParam = new progress state
            ''' <summary>Set progress bar range.</summary>
            TDM_SET_PROGRESS_BAR_RANGE = WM_USER + 105
            ' lParam = MAKELPARAM(nMinRange, nMaxRange)
            ''' <summary>Set progress bar position.</summary>
            TDM_SET_PROGRESS_BAR_POS = WM_USER + 106
            ' wParam = new position
            ''' <summary>Set progress bar marquee (animation).</summary>
            TDM_SET_PROGRESS_BAR_MARQUEE = WM_USER + 107
            ' wParam = 0 (stop marquee), wParam != 0 (start marquee), lparam = speed (milliseconds
            ' between repaints)
            ''' <summary>Set a text element of the Task Dialog.</summary>
            TDM_SET_ELEMENT_TEXT = WM_USER + 108
            ' wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
            ''' <summary>Click a radio button.</summary>
            TDM_CLICK_RADIO_BUTTON = WM_USER + 110
            ' wParam = Radio Button ID
            ''' <summary>Enable or disable a button.</summary>
            TDM_ENABLE_BUTTON = WM_USER + 111
            ' lParam = 0 (disable), lParam != 0 (enable), wParam = Button ID
            ''' <summary>Enable or disable a radio button.</summary>
            TDM_ENABLE_RADIO_BUTTON = WM_USER + 112
            ' lParam = 0 (disable), lParam != 0 (enable), wParam = Radio Button ID
            ''' <summary>Check or uncheck the verfication checkbox.</summary>
            TDM_CLICK_VERIFICATION = WM_USER + 113
            ' wParam = 0 (unchecked), 1 (checked), lParam = 1 (set key focus)
            ''' <summary>Update the text of an element (no effect if origially set as null).</summary>
            TDM_UPDATE_ELEMENT_TEXT = WM_USER + 114
            ' wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
            ''' <summary>
            '''   Designate whether a given Task Dialog button or command link should have a User
            '''   Account Control (UAC) shield icon.
            ''' </summary>
            TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE = WM_USER + 115
            ' wParam = Button ID, lParam = 0 (elevation not required), lParam != 0 (elevation required)
            ''' <summary>Refreshes the icon of the task dialog.</summary>
            TDM_UPDATE_ICON = WM_USER + 116
            ' wParam = icon element (TASKDIALOG_ICON_ELEMENTS), lParam = new icon (hIcon if
            ' TDF_USE_HICON_* was set, PCWSTR otherwise)
        End Enum

        '''' <summary>TaskDialog taken from commctrl.h.</summary>
        '''' <param name="hwndParent">Parent window.</param>
        '''' <param name="hInstance">Module instance to get resources from.</param>
        '''' <param name="pszWindowTitle">Title of the Task Dialog window.</param>
        '''' <param name="pszMainInstruction">The main instructions.</param>
        '''' <param name="dwCommonButtons">Common push buttons to show.</param>
        '''' <param name="pszIcon">The main icon.</param>
        '''' <param name="pnButton">The push button pressed.</param>
        ''[DllImport("ComCtl32", CharSet = CharSet.Unicode, PreserveSig = false)]
        ''public static extern void TaskDialog(
        ''    [In] IntPtr hwndParent,
        ''    [In] IntPtr hInstance,
        ''    [In] String pszWindowTitle,
        ''    [In] String pszMainInstruction,
        ''    [In] TaskDialogCommonButtons dwCommonButtons,
        ''    [In] IntPtr pszIcon,
        ''    [Out] out int pnButton);

        ''' <summary>TaskDialogIndirect taken from commctl.h</summary>
        ''' <param name="pTaskConfig">All the parameters about the Task Dialog to Show.</param>
        ''' <param name="pnButton">The push button pressed.</param>
        ''' <param name="pnRadioButton">The radio button that was selected.</param>
        ''' <param name="pfVerificationFlagChecked">
        '''   The state of the verification checkbox on dismiss of the Task Dialog.
        ''' </param>
        <DllImport("ComCtl32", CharSet:=CharSet.Unicode, PreserveSig:=False)>
        Friend Shared Sub TaskDialogIndirect(<[In]> ByRef pTaskConfig As TASKDIALOGCONFIG, <Out> ByRef pnButton As Integer, <Out> ByRef pnRadioButton As Integer, <Out> ByRef pfVerificationFlagChecked As Boolean)
        End Sub

        ''' <summary>Win32 SendMessage.</summary>
        ''' <param name="hWnd">Window handle to send to.</param>
        ''' <param name="Msg">The windows message to send.</param>
        ''' <param name="wParam">Specifies additional message-specific information.</param>
        ''' <param name="lParam">Specifies additional message-specific information.</param>
        ''' <returns>
        '''   The return value specifies the result of the message processing; it depends on the
        '''   message sent.
        ''' </returns>
        <DllImport("user32.dll")>
        Friend Shared Function SendMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
        End Function

        ''' <summary>Win32 SendMessage.</summary>
        ''' <param name="hWnd">Window handle to send to.</param>
        ''' <param name="Msg">The windows message to send.</param>
        ''' <param name="wParam">Specifies additional message-specific information.</param>
        ''' <param name="lParam">Specifies additional message-specific information as a string.</param>
        ''' <returns>
        '''   The return value specifies the result of the message processing; it depends on the
        '''   message sent.
        ''' </returns>
        <DllImport("user32.dll", EntryPoint:="SendMessage")>
        Friend Shared Function SendMessageWithString(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> lParam As String) As IntPtr
        End Function

        ''' <summary>TASKDIALOGCONFIG taken from commctl.h.</summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=1)>
        Friend Structure TASKDIALOGCONFIG
            ''' <summary>Size of the structure in bytes.</summary>
            Public cbSize As UInteger

            ' Managed code owns actual resource. Passed to native in syncronous call. No lifetime issues.
            ''' <summary>Parent window handle.</summary>
            <SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")>
            Public hwndParent As IntPtr

            ' Managed code owns actual resource. Passed to native in syncronous call. No lifetime issues.
            ''' <summary>Module instance handle for resources.</summary>
            <SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")>
            Public hInstance As IntPtr

            ''' <summary>Flags.</summary>
            Public dwFlags As TASKDIALOG_FLAGS
            ' TASKDIALOG_FLAGS (TDF_XXX) flags
            ''' <summary>Bit flags for commonly used buttons.</summary>
            Public dwCommonButtons As TaskDialogCommonButtons
            ' TASKDIALOG_COMMON_BUTTON (TDCBF_XXX) flags
            ''' <summary>Window title.</summary>
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pszWindowTitle As String            ' string or MAKEINTRESOURCE()

            ' Managed code owns actual resource. Passed to native in syncronous call. No lifetime issues.
            ''' <summary>
            '''   The Main icon. Overloaded member. Can be string, a handle, a special value or a
            '''   resource ID.
            ''' </summary>
            <SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")>
            Public MainIcon As IntPtr

            ''' <summary>Main Instruction.</summary>
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pszMainInstruction As String

            ''' <summary>Content.</summary>
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pszContent As String

            ''' <summary>Count of custom Buttons.</summary>
            Public cButtons As UInteger

            ' Managed code owns actual resource. Passed to native in syncronous call. No lifetime issues.
            ''' <summary>Array of custom buttons.</summary>
            <SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")>
            Public pButtons As IntPtr

            ''' <summary>ID of default button.</summary>
            Public nDefaultButton As Integer

            ''' <summary>Count of radio Buttons.</summary>
            Public cRadioButtons As UInteger

            ' Managed code owns actual resource. Passed to native in syncronous call. No lifetime issues.
            ''' <summary>Array of radio buttons.</summary>
            <SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")>
            Public pRadioButtons As IntPtr

            ''' <summary>ID of default radio button.</summary>
            Public nDefaultRadioButton As Integer

            ''' <summary>Text for verification check box. often "Don't ask be again".</summary>
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pszVerificationText As String

            ''' <summary>Expanded Information.</summary>
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pszExpandedInformation As String

            ''' <summary>Text for expanded control.</summary>
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pszExpandedControlText As String

            ''' <summary>Text for expanded control.</summary>
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pszCollapsedControlText As String

            ' Managed code owns actual resource. Passed to native in syncronous call. No lifetime issues.
            ''' <summary>Icon for the footer. An overloaded member link MainIcon.</summary>
            <SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")>
            Public FooterIcon As IntPtr

            ''' <summary>Footer Text.</summary>
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pszFooter As String

            ''' <summary>Function pointer for callback.</summary>
            Public pfCallback As TaskDialogCallback

            ' Managed code owns actual resource. Passed to native in syncronous call. No lifetime issues.
            ''' <summary>Data that will be passed to the call back.</summary>
            <SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")>
            Public lpCallbackData As IntPtr

            ''' <summary>Width of the Task Dialog's area in DLU's.</summary>
            Public cxWidth As UInteger
            ' width of the Task Dialog's client area in DLU's. If 0, Task Dialog will calculate the
            ' ideal width.
        End Structure
    End Class
End Namespace

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================
