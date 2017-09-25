
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
    ''' <summary>
    ''' The active Task Dialog window. Provides several methods for acting on the active TaskDialog.
    ''' You should not use this object after the TaskDialog Destroy notification callback. Doing so
    ''' will result in undefined behavior and likely crash.
    ''' </summary>
    Public Class ActiveTaskDialog
        Implements IWin32Window
        ' We don't own the window.
        ''' <summary>The Task Dialog's window handle.</summary>
        <SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")>
        Private m_handle As IntPtr

        ''' <summary>
        ''' Creates a ActiveTaskDialog.
        ''' </summary>
        ''' <param name="handle">The Task Dialog's window handle.</param>
        Friend Sub New(handle As IntPtr)
            If handle = IntPtr.Zero Then
                Throw New ArgumentNullException(NameOf(handle))
            End If

            Me.m_handle = handle
        End Sub

        ''' <summary>
        ''' The Task Dialog's window handle.
        ''' </summary>
        Public ReadOnly Property Handle() As IntPtr Implements IWin32Window.Handle
            Get
                Return Me.m_handle
            End Get
        End Property

        '' Not supported. Task Dialog Spec does not indicate what this is for.
        ''public void NavigatePage()
        ''{
        ''    // TDM_NAVIGATE_PAGE                   = WM_USER+101,
        ''    UnsafeNativeMethods.SendMessage(
        ''        this.windowHandle,
        ''        (uint)UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_NAVIGATE_PAGE,
        ''        IntPtr.Zero,
        ''        //a UnsafeNativeMethods.TASKDIALOGCONFIG value);
        ''}

        ''' <summary>
        ''' Simulate the action of a button click in the TaskDialog. This can be a DialogResult value
        ''' or the ButtonID set on a TasDialogButton set on TaskDialog.Buttons.
        ''' </summary>
        ''' <param name="buttonId">Indicates the button ID to be selected.</param>
        ''' <returns>If the function succeeds the return value is true.</returns>
        Public Function ClickButton(buttonId As Integer) As Boolean
            ' TDM_CLICK_BUTTON                    = WM_USER+102, // wParam = Button ID
            Return UnsafeNativeMethods.SendMessage(Me.m_handle,
                                               CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_CLICK_BUTTON),
                                               New IntPtr(buttonId),
                                               IntPtr.Zero) <> IntPtr.Zero
        End Function

        ''' <summary>
        ''' Used to indicate whether the hosted progress bar should be displayed in marquee mode or not.
        ''' </summary>
        ''' <param name="marquee">Specifies whether the progress bar sbould be shown in Marquee mode.
        ''' A value of true turns on Marquee mode.</param>
        ''' <returns>If the function succeeds the return value is true.</returns>
        Public Function SetMarqueeProgressBar(marquee As Boolean) As Boolean
            ' TDM_SET_MARQUEE_PROGRESS_BAR        = WM_USER+103, // wParam = 0 (nonMarque) wParam != 0 (Marquee)
            Return UnsafeNativeMethods.SendMessage(Me.m_handle,
                                               CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_SET_MARQUEE_PROGRESS_BAR),
                                               (If(marquee, New IntPtr(1), IntPtr.Zero)),
                                               IntPtr.Zero) <> IntPtr.Zero

            ' Future: get more detailed error from and throw.
        End Function

        ''' <summary>
        ''' Sets the state of the progress bar.
        ''' </summary>
        ''' <param name="newState">The state to set the progress bar.</param>
        ''' <returns>If the function succeeds the return value is true.</returns>
        Public Function SetProgressBarState(newState As ProgressBarState) As Boolean
            ' TDM_SET_PROGRESS_BAR_STATE          = WM_USER+104, // wParam = new progress state
            Return UnsafeNativeMethods.SendMessage(Me.m_handle,
                                               CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_SET_PROGRESS_BAR_STATE),
                                               New IntPtr(newState),
                                               IntPtr.Zero) <> IntPtr.Zero

            ' Future: get more detailed error from and throw.
        End Function

        ''' <summary>
        ''' Set the minimum and maximum values for the hosted progress bar.
        ''' </summary>
        ''' <param name="minRange">Minimum range value. By default, the minimum value is zero.</param>
        ''' <param name="maxRange">Maximum range value.  By default, the maximum value is 100.</param>
        ''' <returns>If the function succeeds the return value is true.</returns>
        Public Function SetProgressBarRange(minRange As Int16, maxRange As Int16) As Boolean
            ' TDM_SET_PROGRESS_BAR_RANGE          = WM_USER+105, // lParam = MAKELPARAM(nMinRange, nMaxRange)
            ' #define MAKELPARAM(l, h)      ((LPARAM)(DWORD)MAKELONG(l, h))
            ' #define MAKELONG(a, b)      ((LONG)(((WORD)(((DWORD_PTR)(a)) & 0xffff)) | ((DWORD)((WORD)(((DWORD_PTR)(b)) & 0xffff))) << 16))
            Dim lparam As IntPtr = New IntPtr((minRange And &HFFFF) Or ((maxRange And &HFFFF) << 16))
            Return UnsafeNativeMethods.SendMessage(Me.m_handle,
                                               CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_SET_PROGRESS_BAR_RANGE),
                                               IntPtr.Zero,
                                               lparam) <> IntPtr.Zero

            ' Return value is actually prior range.
        End Function

        ''' <summary>
        ''' Set the current position for a progress bar.
        ''' </summary>
        ''' <param name="newPosition">The new position.</param>
        ''' <returns>Returns the previous value if successful, or zero otherwise.</returns>
        Public Function SetProgressBarPosition(newPosition As Integer) As Integer
            ' TDM_SET_PROGRESS_BAR_POS            = WM_USER+106,
            '' wParam = new position
            Return CInt(UnsafeNativeMethods.SendMessage(Me.m_handle,
                                                    CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_SET_PROGRESS_BAR_POS),
                                                    New IntPtr(newPosition),
                                                    IntPtr.Zero))
        End Function

        ''' <summary>
        ''' Sets the animation state of the Marquee Progress Bar.
        ''' </summary>
        ''' <param name="startMarquee">true starts the marquee animation and false stops it.</param>
        ''' <param name="speed">The time in milliseconds between refreshes.</param>
        Public Sub SetProgressBarMarquee(startMarquee As Boolean, speed As UInteger)
            ' TDM_SET_PROGRESS_BAR_MARQUEE        = WM_USER+107,
            '' wParam = 0 (stop marquee), wParam != 0 (start marquee), lparam = speed (milliseconds between repaints)
            UnsafeNativeMethods.SendMessage(Me.m_handle,
                                        CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_SET_PROGRESS_BAR_MARQUEE),
                                        If(startMarquee, New IntPtr(1), IntPtr.Zero),
                                        New IntPtr(speed))
        End Sub

        ''' <summary>
        ''' Updates the content text.
        ''' </summary>
        ''' <param name="content">The new value.</param>
        ''' <returns>If the function succeeds the return value is true.</returns>
        Public Function SetContent(content As String) As Boolean
            ' TDE_CONTENT,
            ' TDM_SET_ELEMENT_TEXT                = WM_USER+108
            '' wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
            Return UnsafeNativeMethods.SendMessageWithString(Me.m_handle,
                                                         CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_SET_ELEMENT_TEXT),
                                                         New IntPtr(UnsafeNativeMethods.TASKDIALOG_ELEMENTS.TDE_CONTENT),
                                                         content) <> IntPtr.Zero
        End Function

        ''' <summary>
        ''' Updates the Expanded Information text.
        ''' </summary>
        ''' <param name="expandedInformation">The new value.</param>
        ''' <returns>If the function succeeds the return value is true.</returns>
        Public Function SetExpandedInformation(expandedInformation As String) As Boolean
            ' TDE_EXPANDED_INFORMATION,
            ' TDM_SET_ELEMENT_TEXT                = WM_USER+108  // wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
            Return UnsafeNativeMethods.SendMessageWithString(Me.m_handle,
                                                         CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_SET_ELEMENT_TEXT),
                                                         New IntPtr(UnsafeNativeMethods.TASKDIALOG_ELEMENTS.TDE_EXPANDED_INFORMATION),
                                                         expandedInformation) <> IntPtr.Zero
        End Function

        ''' <summary>
        ''' Updates the Footer text.
        ''' </summary>
        ''' <param name="footer">The new value.</param>
        ''' <returns>If the function succeeds the return value is true.</returns>
        Public Function SetFooter(footer As String) As Boolean
            ' TDE_FOOTER,
            ' TDM_SET_ELEMENT_TEXT                = WM_USER+108  // wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
            Return UnsafeNativeMethods.SendMessageWithString(Me.m_handle,
                                                          CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_SET_ELEMENT_TEXT),
                                                          New IntPtr(UnsafeNativeMethods.TASKDIALOG_ELEMENTS.TDE_FOOTER),
                                                          footer) <> IntPtr.Zero
        End Function

        ''' <summary>
        ''' Updates the Main Instruction.
        ''' </summary>
        ''' <param name="mainInstruction">The new value.</param>
        ''' <returns>If the function succeeds the return value is true.</returns>
        Public Function SetMainInstruction(mainInstruction As String) As Boolean
            ' TDE_MAIN_INSTRUCTION
            ' TDM_SET_ELEMENT_TEXT                = WM_USER+108  // wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
            Return UnsafeNativeMethods.SendMessageWithString(Me.m_handle,
                                                         CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_SET_ELEMENT_TEXT),
                                                         New IntPtr(UnsafeNativeMethods.TASKDIALOG_ELEMENTS.TDE_MAIN_INSTRUCTION),
                                                         mainInstruction) <> IntPtr.Zero
        End Function

        ''' <summary>
        ''' Simulate the action of a radio button click in the TaskDialog.
        ''' The passed buttonID is the ButtonID set on a TaskDialogButton set on TaskDialog.RadioButtons.
        ''' </summary>
        ''' <param name="buttonId">Indicates the button ID to be selected.</param>
        Public Sub ClickRadioButton(buttonId As Integer)
            ' TDM_CLICK_RADIO_BUTTON = WM_USER+110, // wParam = Radio Button ID
            UnsafeNativeMethods.SendMessage(Me.m_handle,
                                        CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_CLICK_RADIO_BUTTON),
                                        New IntPtr(buttonId),
                                        IntPtr.Zero)
        End Sub

        ''' <summary>
        ''' Enable or disable a button in the TaskDialog.
        ''' The passed buttonID is the ButtonID set on a TaskDialogButton set on TaskDialog.Buttons
        ''' or a common button ID.
        ''' </summary>
        ''' <param name="buttonId">Indicates the button ID to be enabled or diabled.</param>
        ''' <param name="enable">Enambe the button if true. Disable the button if false.</param>
        Public Sub EnableButton(buttonId As Integer, enable As Boolean)
            ' TDM_ENABLE_BUTTON = WM_USER+111, // lParam = 0 (disable), lParam != 0 (enable), wParam = Button ID
            UnsafeNativeMethods.SendMessage(Me.m_handle,
                                        CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_ENABLE_BUTTON),
                                        New IntPtr(buttonId),
                                        New IntPtr(If(Not enable, 0, 1)))
        End Sub

        ''' <summary>
        ''' Enable or disable a radio button in the TaskDialog.
        ''' The passed buttonID is the ButtonID set on a TaskDialogButton set on TaskDialog.RadioButtons.
        ''' </summary>
        ''' <param name="buttonId">Indicates the button ID to be enabled or diabled.</param>
        ''' <param name="enable">Enambe the button if true. Disable the button if false.</param>
        Public Sub EnableRadioButton(buttonId As Integer, enable As Boolean)
            ' TDM_ENABLE_RADIO_BUTTON = WM_USER+112, // lParam = 0 (disable), lParam != 0 (enable), wParam = Radio Button ID
            UnsafeNativeMethods.SendMessage(Me.m_handle,
                                        CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_ENABLE_RADIO_BUTTON),
                                        New IntPtr(buttonId),
                                        New IntPtr(If(Not enable, 0, 1)))
        End Sub

        ''' <summary>
        ''' Check or uncheck the verification checkbox in the TaskDialog.
        ''' </summary>
        ''' <param name="checkedState">The checked state to set the verification checkbox.</param>
        ''' <param name="setKeyboardFocusToCheckBox">True to set the keyboard focus to the checkbox, and fasle otherwise.</param>
        Public Sub ClickVerification(checkedState As Boolean, setKeyboardFocusToCheckBox As Boolean)
            ' TDM_CLICK_VERIFICATION = WM_USER+113, // wParam = 0 (unchecked), 1 (checked), lParam = 1 (set key focus)
            UnsafeNativeMethods.SendMessage(Me.m_handle,
                                        CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_CLICK_VERIFICATION),
                                        (If(checkedState,
                                        New IntPtr(1),
                                        IntPtr.Zero)),
                                        If(setKeyboardFocusToCheckBox, New IntPtr(1), IntPtr.Zero))
        End Sub

        ''' <summary>
        ''' Updates the content text.
        ''' </summary>
        ''' <param name="content">The new value.</param>
        Public Sub UpdateContent(content As String)
            ' TDE_CONTENT,
            ' TDM_UPDATE_ELEMENT_TEXT             = WM_USER+114, // wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
            UnsafeNativeMethods.SendMessageWithString(Me.m_handle,
                                                  CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_UPDATE_ELEMENT_TEXT),
                                                  New IntPtr(UnsafeNativeMethods.TASKDIALOG_ELEMENTS.TDE_CONTENT),
                                                  content)
        End Sub

        ''' <summary>
        ''' Updates the Expanded Information text. No effect if it was previously set to null.
        ''' </summary>
        ''' <param name="expandedInformation">The new value.</param>
        Public Sub UpdateExpandedInformation(expandedInformation As String)
            ' TDE_EXPANDED_INFORMATION,
            ' TDM_UPDATE_ELEMENT_TEXT             = WM_USER+114, // wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
            UnsafeNativeMethods.SendMessageWithString(Me.m_handle,
                                                  CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_UPDATE_ELEMENT_TEXT),
                                                  New IntPtr(UnsafeNativeMethods.TASKDIALOG_ELEMENTS.TDE_EXPANDED_INFORMATION),
                                                  expandedInformation)
        End Sub

        ''' <summary>
        ''' Updates the Footer text. No Effect if it was perviously set to null.
        ''' </summary>
        ''' <param name="footer">The new value.</param>
        Public Sub UpdateFooter(footer As String)
            ' TDE_FOOTER,
            ' TDM_UPDATE_ELEMENT_TEXT             = WM_USER+114, // wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
            UnsafeNativeMethods.SendMessageWithString(Me.m_handle,
                                                  CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_UPDATE_ELEMENT_TEXT),
                                                  New IntPtr(UnsafeNativeMethods.TASKDIALOG_ELEMENTS.TDE_FOOTER),
                                                  footer)
        End Sub

        ''' <summary>
        ''' Updates the Main Instruction.
        ''' </summary>
        ''' <param name="mainInstruction">The new value.</param>
        Public Sub UpdateMainInstruction(mainInstruction As String)
            ' TDE_MAIN_INSTRUCTION
            ' TDM_UPDATE_ELEMENT_TEXT             = WM_USER+114, // wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
            UnsafeNativeMethods.SendMessageWithString(Me.m_handle,
                                                  CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_UPDATE_ELEMENT_TEXT),
                                                  New IntPtr(UnsafeNativeMethods.TASKDIALOG_ELEMENTS.TDE_MAIN_INSTRUCTION),
                                                  mainInstruction)
        End Sub

        ''' <summary>
        ''' Designate whether a given Task Dialog button or command link should have a User Account Control (UAC) shield icon.
        ''' </summary>
        ''' <param name="buttonId">ID of the push button or command link to be updated.</param>
        ''' <param name="elevationRequired">False to designate that the action invoked by the button does not require elevation;
        ''' true to designate that the action does require elevation.</param>
        Public Sub SetButtonElevationRequiredState(buttonId As Integer, elevationRequired As Boolean)
            ' TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE = WM_USER+115, // wParam = Button ID, lParam = 0 (elevation not required), lParam != 0 (elevation required)
            UnsafeNativeMethods.SendMessage(Me.m_handle,
                                        CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE),
                                        New IntPtr(buttonId),
                                        If(elevationRequired, New IntPtr(1), IntPtr.Zero))
        End Sub

        ''' <summary>
        ''' Updates the main instruction icon. Note the type (standard via enum or
        ''' custom via Icon type) must be used when upating the icon.
        ''' </summary>
        ''' <param name="icon">Task Dialog standard icon.</param>
        Public Sub UpdateMainIcon(icon As TaskDialogIcon)
            ' TDM_UPDATE_ICON = WM_USER+116  // wParam = icon element (TASKDIALOG_ICON_ELEMENTS), lParam = new icon (hIcon if TDF_USE_HICON_* was set, PCWSTR otherwise)
            UnsafeNativeMethods.SendMessage(Me.m_handle, CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_UPDATE_ICON), New IntPtr(UnsafeNativeMethods.TASKDIALOG_ICON_ELEMENTS.TDIE_ICON_MAIN), New IntPtr(icon))
        End Sub

        ''' <summary>
        ''' Updates the main instruction icon. Note the type (standard via enum or
        ''' custom via Icon type) must be used when upating the icon.
        ''' </summary>
        ''' <param name="icon">The icon to set.</param>
        Public Sub UpdateMainIcon(icon As Icon)
            ' TDM_UPDATE_ICON = WM_USER+116  // wParam = icon element (TASKDIALOG_ICON_ELEMENTS), lParam = new icon (hIcon if TDF_USE_HICON_* was set, PCWSTR otherwise)
            UnsafeNativeMethods.SendMessage(Me.m_handle,
                                        CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_UPDATE_ICON),
                                        New IntPtr(UnsafeNativeMethods.TASKDIALOG_ICON_ELEMENTS.TDIE_ICON_MAIN),
                                        If(icon Is Nothing, IntPtr.Zero, icon.Handle))
        End Sub

        ''' <summary>
        ''' Updates the footer icon. Note the type (standard via enum or
        ''' custom via Icon type) must be used when upating the icon.
        ''' </summary>
        ''' <param name="icon">Task Dialog standard icon.</param>
        Public Sub UpdateFooterIcon(icon As TaskDialogIcon)
            ' TDM_UPDATE_ICON = WM_USER+116  // wParam = icon element (TASKDIALOG_ICON_ELEMENTS), lParam = new icon (hIcon if TDF_USE_HICON_* was set, PCWSTR otherwise)
            UnsafeNativeMethods.SendMessage(Me.m_handle,
                                        CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_UPDATE_ICON),
                                        New IntPtr(UnsafeNativeMethods.TASKDIALOG_ICON_ELEMENTS.TDIE_ICON_FOOTER),
                                        New IntPtr(icon))
        End Sub

        ''' <summary>
        ''' Updates the footer icon. Note the type (standard via enum or
        ''' custom via Icon type) must be used when upating the icon.
        ''' </summary>
        ''' <param name="icon">The icon to set.</param>
        Public Sub UpdateFooterIcon(icon As Icon)
            ' TDM_UPDATE_ICON = WM_USER+116  // wParam = icon element (TASKDIALOG_ICON_ELEMENTS), lParam = new icon (hIcon if TDF_USE_HICON_* was set, PCWSTR otherwise)
            UnsafeNativeMethods.SendMessage(Me.m_handle,
                                        CUInt(UnsafeNativeMethods.TASKDIALOG_MESSAGES.TDM_UPDATE_ICON),
                                        New IntPtr(UnsafeNativeMethods.TASKDIALOG_ICON_ELEMENTS.TDIE_ICON_FOOTER),
                                        If(icon Is Nothing, IntPtr.Zero, icon.Handle))
        End Sub
    End Class
End Namespace

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================
