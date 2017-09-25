
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
    ''' The signature of the callback that receives notifications from the Task Dialog.
    ''' </summary>
    ''' <param name="taskDialog">The active task dialog which has methods that can be performed on an active Task Dialog.</param>
    ''' <param name="args">The notification arguments including the type of notification and information for the notification.</param>
    ''' <param name="callbackData">The value set on TaskDialog.CallbackData</param>
    ''' <returns>Return value meaning varies depending on the Notification member of args.</returns>
    Public Delegate Function TaskDialogCallback(taskDialog As ActiveTaskDialog, args As TaskDialogNotificationArgs, callbackData As Object) As Boolean

    ''' <summary>
    ''' The TaskDialog common button flags used to specify the built-in buttons to show in the TaskDialog.
    ''' </summary>
    <Flags>
    Public Enum TaskDialogCommonButtons
        ''' <summary>
        ''' No common buttons.
        ''' </summary>
        None = 0

        ''' <summary>
        ''' OK common button. If selected Task Dialog will return DialogResult.OK.
        ''' </summary>
        Ok = &H1

        ''' <summary>
        ''' Yes common button. If selected Task Dialog will return DialogResult.Yes.
        ''' </summary>
        Yes = &H2

        ''' <summary>
        ''' No common button. If selected Task Dialog will return DialogResult.No.
        ''' </summary>
        No = &H4

        ''' <summary>
        ''' Cancel common button. If selected Task Dialog will return DialogResult.Cancel.
        ''' If this button is specified, the dialog box will respond to typical cancel actions (Alt-F4 and Escape).
        ''' </summary>
        Cancel = &H8

        ''' <summary>
        ''' Retry common button. If selected Task Dialog will return DialogResult.Retry.
        ''' </summary>
        Retry = &H10

        ''' <summary>
        ''' Close common button. If selected Task Dialog will return this value.
        ''' </summary>
        Close = &H20
    End Enum

    ' Type comes from CommCtrl.h
    ''' <summary>The System icons the TaskDialog supports.</summary>
    <SuppressMessage("Microsoft.Design", "CA1028:EnumStorageShouldBeInt32")>
    Public Enum TaskDialogIcon As UInteger
        ''' <summary>
        ''' No Icon.
        ''' </summary>
        None = 0

        ''' <summary>
        ''' System warning icon.
        ''' </summary>
        Warning = &HFFFF
        ' MAKEINTRESOURCEW(-1)
        ''' <summary>
        ''' System Error icon.
        ''' </summary>
        [Error] = &HFFFE
        ' MAKEINTRESOURCEW(-2)
        ''' <summary>
        ''' System Information icon.
        ''' </summary>
        Information = &HFFFD
        ' MAKEINTRESOURCEW(-3)
        ''' <summary>
        ''' Shield icon.
        ''' </summary>
        Shield = &HFFFC
        ' MAKEINTRESOURCEW(-4)
        ''' <summary>
        ''' Shield icon on a blue banner.
        ''' </summary>
        SecurityShieldBlue = &HFFFB

        ''' <summary>
        ''' Warning security shield icon. Displays an amber colored banner in addition when set as the <see cref="TaskDialog.MainIcon"/>.
        ''' </summary>
        SecurityShieldWarning = &HFFFA

        ''' <summary>
        ''' Error security shield icon. Displays a red colored banner in addition when set as the <see cref="TaskDialog.MainIcon"/>.
        ''' </summary>
        SecurityError = &HFFF9

        ''' <summary>
        ''' Success security shield icon. Displays a green colored banner in addition when set as the <see cref="TaskDialog.MainIcon"/>.
        ''' </summary>
        SecuritySuccess = &HFFF8

        ''' <summary>
        ''' Shield icon on a gray banner.
        ''' </summary>
        SecurityShieldGray = &HFFF7
    End Enum

    ''' <summary>
    ''' Task Dialog callback notifications.
    ''' </summary>
    Public Enum TaskDialogNotification
        ''' <summary>
        ''' Sent by the Task Dialog once the dialog has been created and before it is displayed.
        ''' The value returned by the callback is ignored.
        ''' </summary>
        Created = 0

        ' Spec is not clear what this is so not supporting it.
        ''''' <summary>
        '''''   Sent by the Task Dialog when a navigation has occurred. The value returned by the
        '''''   callback is ignored.
        ''''' </summary>
        ''Navigated = 1,

        ''' <summary>
        ''' Sent by the Task Dialog when the user selects a button or command link in the task dialog.
        ''' The button ID corresponding to the button selected will be available in the
        ''' TaskDialogNotificationArgs. To prevent the Task Dialog from closing, the application must
        ''' return true, otherwise the Task Dialog will be closed and the button ID returned to via
        ''' the original application call.
        ''' </summary>
        ButtonClicked = 2
        ' wParam = Button ID
        ''' <summary>
        ''' Sent by the Task Dialog when the user clicks on a hyperlink in the Task Dialog�s content.
        ''' The string containing the HREF of the hyperlink will be available in the
        ''' TaskDialogNotificationArgs. To prevent the TaskDialog from shell executing the hyperlink,
        ''' the application must return TRUE, otherwise ShellExecute will be called.
        ''' </summary>
        HyperlinkClicked = 3
        ' lParam = (LPCWSTR)pszHREF
        ''' <summary>
        ''' Sent by the Task Dialog approximately every 200 milliseconds when TaskDialog.CallbackTimer
        ''' has been set to true. The number of milliseconds since the dialog was created or the
        ''' notification returned true is available on the TaskDialogNotificationArgs. To reset
        ''' the tickcount, the application must return true, otherwise the tickcount will continue to
        ''' increment.
        ''' </summary>
        Timer = 4
        ' wParam = Milliseconds since dialog created or timer reset
        ''' <summary>
        ''' Sent by the Task Dialog when it is destroyed and its window handle no longer valid.
        ''' The value returned by the callback is ignored.
        ''' </summary>
        Destroyed = 5

        ''' <summary>
        ''' Sent by the Task Dialog when the user selects a radio button in the task dialog.
        ''' The button ID corresponding to the button selected will be available in the
        ''' TaskDialogNotificationArgs.
        ''' The value returned by the callback is ignored.
        ''' </summary>
        RadioButtonClicked = 6
        ' wParam = Radio Button ID
        ''' <summary>
        ''' Sent by the Task Dialog once the dialog has been constructed and before it is displayed.
        ''' The value returned by the callback is ignored.
        ''' </summary>
        DialogConstructed = 7

        ''' <summary>
        ''' Sent by the Task Dialog when the user checks or unchecks the verification checkbox.
        ''' The verificationFlagChecked value is available on the TaskDialogNotificationArgs.
        ''' The value returned by the callback is ignored.
        ''' </summary>
        VerificationClicked = 8
        ' wParam = 1 if checkbox checked, 0 if not, lParam is unused and always 0
        ''' <summary>
        ''' Sent by the Task Dialog when the user presses F1 on the keyboard while the dialog has focus.
        ''' The value returned by the callback is ignored.
        ''' </summary>
        Help = 9

        ''' <summary>
        ''' Sent by the task dialog when the user clicks on the dialog's expando button.
        ''' The expanded value is available on the TaskDialogNotificationArgs.
        ''' The value returned by the callback is ignored.
        ''' </summary>
        ExpandoButtonClicked = 10
        ' wParam = 0 (dialog is now collapsed), wParam != 0 (dialog is now expanded)
    End Enum

    ' Comes from CommCtrl.h PBST_* values which don't have a zero.
    ''' <summary>Progress bar state.</summary>
    <SuppressMessage("Microsoft.Design", "CA1008:EnumsShouldHaveZeroValue")>
    Public Enum ProgressBarState
        ''' <summary>
        ''' Normal.
        ''' </summary>
        Normal = 1

        ''' <summary>
        ''' Error state.
        ''' </summary>
        [Error] = 2

        ''' <summary>
        ''' Paused state.
        ''' </summary>
        Paused = 3
    End Enum

    ' Would be unused code as not required for usage.
    ''' <summary>A custom button for the TaskDialog.</summary>
    <SuppressMessage("Microsoft.Performance", "CA1815:OverrideEqualsAndOperatorEqualsOnValueTypes")>
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=1)>
    Public Structure TaskDialogButton
        ''' <summary>
        ''' The ID of the button. This value is returned by TaskDialog.Show when the button is clicked.
        ''' </summary>
        Private m_buttonId As Integer

        ''' <summary>
        ''' The string that appears on the button.
        ''' </summary>
        <MarshalAs(UnmanagedType.LPWStr)>
        Private m_buttonText As String

        ''' <summary>
        ''' Initialize the custom button.
        ''' </summary>
        ''' <param name="id">The ID of the button. This value is returned by TaskDialog.Show when
        ''' the button is clicked. Typically this will be a value in the DialogResult enum.</param>
        ''' <param name="text">The string that appears on the button.</param>
        Public Sub New(id As Integer, text As String)
            Me.m_buttonId = id
            Me.m_buttonText = text
        End Sub

        ''' <summary>
        ''' The ID of the button. This value is returned by TaskDialog.Show when the button is clicked.
        ''' </summary>
        Public Property ButtonId() As Integer
            Get
                Return Me.m_buttonId
            End Get
            Set
                Me.m_buttonId = Value
            End Set
        End Property

        ''' <summary>
        ''' The string that appears on the button.
        ''' </summary>
        Public Property ButtonText() As String
            Get
                Return Me.m_buttonText
            End Get
            Set
                Me.m_buttonText = Value
            End Set
        End Property
    End Structure

    ''' <summary>
    ''' A Task Dialog. This is like a MessageBox but with many more features. TaskDialog requires Windows Longhorn or later.
    ''' </summary>
    Public Class TaskDialog
        ''' <summary>
        ''' The string to be used for the dialog box title. If this parameter is NULL, the filename of the executable program is used.
        ''' </summary>
        Private m_windowTitle As String

        ''' <summary>
        ''' The string to be used for the main instruction.
        ''' </summary>
        Private m_mainInstruction As String

        ''' <summary>
        ''' The string to be used for the dialog�s primary content. If the EnableHyperlinks member is true,
        ''' then this string may contain hyperlinks in the form: <A HREF="executablestring">Hyperlink Text</A>.
        ''' WARNING: Enabling hyperlinks when using content from an unsafe source may cause security vulnerabilities.
        ''' </summary>
        Private m_content As String

        ''' <summary>
        ''' Specifies the push buttons displayed in the dialog box.  This parameter may be a combination of flags.
        ''' If no common buttons are specified and no custom buttons are specified using the Buttons member, the
        ''' dialog box will contain the OK button by default.
        ''' </summary>
        Private m_commonButtons As TaskDialogCommonButtons

        ''' <summary>
        ''' Specifies a built in icon for the main icon in the dialog. If this is set to none
        ''' and the CustomMainIcon is null then no main icon will be displayed.
        ''' </summary>
        Private m_mainIcon As TaskDialogIcon

        ''' <summary>
        ''' Specifies a custom in icon for the main icon in the dialog. If this is set to none
        ''' and the CustomMainIcon member is null then no main icon will be displayed.
        ''' </summary>
        Private m_customMainIcon As Icon

        ''' <summary>
        ''' Specifies a built in icon for the icon to be displayed in the footer area of the
        ''' dialog box. If this is set to none and the CustomFooterIcon member is null then no
        ''' footer icon will be displayed.
        ''' </summary>
        Private m_footerIcon As TaskDialogIcon

        ''' <summary>
        ''' Specifies a custom icon for the icon to be displayed in the footer area of the
        ''' dialog box. If this is set to none and the CustomFooterIcon member is null then no
        ''' footer icon will be displayed.
        ''' </summary>
        Private m_customFooterIcon As Icon

        ''' <summary>
        ''' Specifies the custom push buttons to display in the dialog. Use CommonButtons member for
        ''' common buttons; OK, Yes, No, Retry and Cancel, and Buttons when you want different text
        ''' on the push buttons.
        ''' </summary>
        Private m_buttons As TaskDialogButton()

        ''' <summary>
        ''' Specifies the radio buttons to display in the dialog.
        ''' </summary>
        Private m_radioButtons As TaskDialogButton()

        ''' <summary>
        ''' The flags passed to TaskDialogIndirect.
        ''' </summary>
        Private flags As UnsafeNativeMethods.TASKDIALOG_FLAGS

        ''' <summary>
        ''' Indicates the default button for the dialog. This may be any of the values specified
        ''' in ButtonId members of one of the TaskDialogButton structures in the Buttons array,
        ''' or one a DialogResult value that corresponds to a buttons specified in the CommonButtons Member.
        ''' If this member is zero or its value does not correspond to any button ID in the dialog,
        ''' then the first button in the dialog will be the default.
        ''' </summary>
        Private m_defaultButton As Integer

        ''' <summary>
        ''' Indicates the default radio button for the dialog. This may be any of the values specified
        ''' in ButtonId members of one of the TaskDialogButton structures in the RadioButtons array.
        ''' If this member is zero or its value does not correspond to any radio button ID in the dialog,
        ''' then the first button in RadioButtons will be the default.
        ''' The property NoDefaultRadioButton can be set to have no default.
        ''' </summary>
        Private m_defaultRadioButton As Integer

        ''' <summary>
        ''' The string to be used to label the verification checkbox. If this member is null, the
        ''' verification checkbox is not displayed in the dialog box.
        ''' </summary>
        Private m_verificationText As String

        ''' <summary>
        ''' The string to be used for displaying additional information. The additional information is
        ''' displayed either immediately below the content or below the footer text depending on whether
        ''' the ExpandFooterArea member is true. If the EnableHyperlinks member is true, then this string
        ''' may contain hyperlinks in the form: <A HREF="executablestring">Hyperlink Text</A>.
        ''' WARNING: Enabling hyperlinks when using content from an unsafe source may cause security vulnerabilities.
        ''' </summary>
        Private m_expandedInformation As String

        ''' <summary>
        ''' The string to be used to label the button for collapsing the expanded information. This
        ''' member is ignored when the ExpandedInformation member is null. If this member is null
        ''' and the CollapsedControlText is specified, then the CollapsedControlText value will be
        ''' used for this member as well.
        ''' </summary>
        Private m_expandedControlText As String

        ''' <summary>
        ''' The string to be used to label the button for expanding the expanded information. This
        ''' member is ignored when the ExpandedInformation member is null.  If this member is null
        ''' and the ExpandedControlText is specified, then the ExpandedControlText value will be
        ''' used for this member as well.
        ''' </summary>
        Private m_collapsedControlText As String

        ''' <summary>
        ''' The string to be used in the footer area of the dialog box. If the EnableHyperlinks member
        ''' is true, then this string may contain hyperlinks in the form: <A HREF="executablestring">
        ''' Hyperlink Text</A>.
        ''' WARNING: Enabling hyperlinks when using content from an unsafe source may cause security vulnerabilities.
        ''' </summary>
        Private m_footer As String

        ''' <summary>
        ''' The callback that receives messages from the Task Dialog when various events occur.
        ''' </summary>
        Private m_callback As TaskDialogCallback

        ''' <summary>
        ''' Reference that is passed to the callback.
        ''' </summary>
        Private m_callbackData As Object

        ''' <summary>
        ''' Specifies the width of the Task Dialog�s client area in DLU�s. If 0, Task Dialog will calculate the ideal width.
        ''' </summary>
        Private m_width As UInteger

        ''' <summary>
        ''' Creates a default Task Dialog.
        ''' </summary>
        Public Sub New()
            Me.Reset()
        End Sub

        ''' <summary>
        ''' Returns true if the current operating system supports TaskDialog. If false TaskDialog.Show should not
        ''' be called as the results are undefined but often results in a crash.
        ''' </summary>
        Public Shared ReadOnly Property IsAvailableOnThisOS() As Boolean
            Get
                Dim os As OperatingSystem = Environment.OSVersion
                If os.Platform <> PlatformID.Win32NT Then
                    Return False
                End If

                Return (os.Version.CompareTo(TaskDialog.RequiredOSVersion) >= 0)
            End Get
        End Property

        ''' <summary>
        ''' The minimum Windows version needed to support TaskDialog.
        ''' </summary>
        Public Shared ReadOnly Property RequiredOSVersion() As Version
            Get
                Return New Version(6, 0, 5243)
            End Get
        End Property

        ''' <summary>
        ''' The string to be used for the dialog box title. If this parameter is NULL, the filename of the executable program is used.
        ''' </summary>
        Public Property WindowTitle() As String
            Get
                Return Me.m_windowTitle
            End Get
            Set
                Me.m_windowTitle = Value
            End Set
        End Property

        ''' <summary>
        ''' The string to be used for the main instruction.
        ''' </summary>
        Public Property MainInstruction() As String
            Get
                Return Me.m_mainInstruction
            End Get
            Set
                Me.m_mainInstruction = Value
            End Set
        End Property

        ''' <summary>
        ''' The string to be used for the dialog�s primary content. If the EnableHyperlinks member is true,
        ''' then this string may contain hyperlinks in the form: <A HREF="executablestring">Hyperlink Text</A>.
        ''' WARNING: Enabling hyperlinks when using content from an unsafe source may cause security vulnerabilities.
        ''' </summary>
        Public Property Content() As String
            Get
                Return Me.m_content
            End Get
            Set
                Me.m_content = Value
            End Set
        End Property

        ''' <summary>
        ''' Specifies the push buttons displayed in the dialog box. This parameter may be a combination of flags.
        ''' If no common buttons are specified and no custom buttons are specified using the Buttons member, the
        ''' dialog box will contain the OK button by default.
        ''' </summary>
        Public Property CommonButtons() As TaskDialogCommonButtons
            Get
                Return Me.m_commonButtons
            End Get
            Set
                Me.m_commonButtons = Value
            End Set
        End Property

        ''' <summary>
        ''' Specifies a built in icon for the main icon in the dialog. If this is set to none
        ''' and the CustomMainIcon is null then no main icon will be displayed.
        ''' </summary>
        Public Property MainIcon() As TaskDialogIcon
            Get
                Return Me.m_mainIcon
            End Get
            Set
                Me.m_mainIcon = Value
            End Set
        End Property

        ''' <summary>
        ''' Specifies a custom in icon for the main icon in the dialog. If this is set to none
        ''' and the CustomMainIcon member is null then no main icon will be displayed.
        ''' </summary>
        Public Property CustomMainIcon() As Icon
            Get
                Return Me.m_customMainIcon
            End Get
            Set
                Me.m_customMainIcon = Value
            End Set
        End Property

        ''' <summary>
        ''' Specifies a built in icon for the icon to be displayed in the footer area of the
        ''' dialog box. If this is set to none and the CustomFooterIcon member is null then no
        ''' footer icon will be displayed.
        ''' </summary>
        Public Property FooterIcon() As TaskDialogIcon
            Get
                Return Me.m_footerIcon
            End Get
            Set
                Me.m_footerIcon = Value
            End Set
        End Property

        ''' <summary>
        ''' Specifies a custom icon for the icon to be displayed in the footer area of the
        ''' dialog box. If this is set to none and the CustomFooterIcon member is null then no
        ''' footer icon will be displayed.
        ''' </summary>
        Public Property CustomFooterIcon() As Icon
            Get
                Return Me.m_customFooterIcon
            End Get
            Set
                Me.m_customFooterIcon = Value
            End Set
        End Property

        ' Style of use is like single value. Array is of value types.
        ' Returns a reference, not a copy.
        ''' <summary>
        '''   Specifies the custom push buttons to display in the dialog. Use CommonButtons member
        '''   for common buttons; OK, Yes, No, Retry and Cancel, and Buttons when you want different
        '''   text on the push buttons.
        ''' </summary>
        <SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")>
        <SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")>
        Public Property Buttons() As TaskDialogButton()
            Get
                Return Me.m_buttons
            End Get

            Set
                If Value Is Nothing Then
                    Throw New ArgumentNullException(NameOf(Value))
                End If

                Me.m_buttons = Value
            End Set
        End Property

        ' Style of use is like single value. Array is of value types.
        ' Returns a reference, not a copy.
        ''' <summary>Specifies the radio buttons to display in the dialog.</summary>
        <SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")>
        <SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")>
        Public Property RadioButtons() As TaskDialogButton()
            Get
                Return Me.m_radioButtons
            End Get

            Set
                If Value Is Nothing Then
                    Throw New ArgumentNullException(NameOf(Value))
                End If

                Me.m_radioButtons = Value
            End Set
        End Property

        ''' <summary>
        ''' Enables hyperlink processing for the strings specified in the Content, ExpandedInformation
        ''' and FooterText members. When enabled, these members may be strings that contain hyperlinks
        ''' in the form: <A HREF="executablestring">Hyperlink Text</A>.
        ''' WARNING: Enabling hyperlinks when using content from an unsafe source may cause security vulnerabilities.
        ''' Note: Task Dialog will not actually execute any hyperlinks. Hyperlink execution must be handled
        ''' in the callback function specified by Callback member.
        ''' </summary>
        Public Property EnableHyperlinks() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_ENABLE_HYPERLINKS) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_ENABLE_HYPERLINKS, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that the dialog should be able to be closed using Alt-F4, Escape and the title bar�s
        ''' close button even if no cancel button is specified in either the CommonButtons or Buttons members.
        ''' </summary>
        Public Property AllowDialogCancellation() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_ALLOW_DIALOG_CANCELLATION) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_ALLOW_DIALOG_CANCELLATION, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that the buttons specified in the Buttons member should be displayed as command links
        ''' (using a standard task dialog glyph) instead of push buttons.  When using command links, all
        ''' characters up to the first new line character in the ButtonText member (of the TaskDialogButton
        ''' structure) will be treated as the command link�s main text, and the remainder will be treated
        ''' as the command link�s note. This flag is ignored if the Buttons member has no entires.
        ''' </summary>
        Public Property UseCommandLinks() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_USE_COMMAND_LINKS) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_USE_COMMAND_LINKS, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that the buttons specified in the Buttons member should be displayed as command links
        ''' (without a glyph) instead of push buttons. When using command links, all characters up to the
        ''' first new line character in the ButtonText member (of the TaskDialogButton structure) will be
        ''' treated as the command link�s main text, and the remainder will be treated as the command link�s
        ''' note. This flag is ignored if the Buttons member has no entires.
        ''' </summary>
        Public Property UseCommandLinksNoIcon() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_USE_COMMAND_LINKS_NO_ICON) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_USE_COMMAND_LINKS_NO_ICON, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that the string specified by the ExpandedInformation member should be displayed at the
        ''' bottom of the dialog�s footer area instead of immediately after the dialog�s content. This flag
        ''' is ignored if the ExpandedInformation member is null.
        ''' </summary>
        Public Property ExpandFooterArea() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_EXPAND_FOOTER_AREA) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_EXPAND_FOOTER_AREA, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that the string specified by the ExpandedInformation member should be displayed
        ''' when the dialog is initially displayed. This flag is ignored if the ExpandedInformation member
        ''' is null.
        ''' </summary>
        Public Property ExpandedByDefault() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_EXPANDED_BY_DEFAULT) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_EXPANDED_BY_DEFAULT, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that the verification checkbox in the dialog should be checked when the dialog is
        ''' initially displayed. This flag is ignored if the VerificationText parameter is null.
        ''' </summary>
        Public Property VerificationFlagChecked() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_VERIFICATION_FLAG_CHECKED) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_VERIFICATION_FLAG_CHECKED, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that a Progress Bar should be displayed.
        ''' </summary>
        Public Property ShowProgressBar() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_SHOW_PROGRESS_BAR) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_SHOW_PROGRESS_BAR, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that an Marquee Progress Bar should be displayed.
        ''' </summary>
        Public Property ShowMarqueeProgressBar() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_SHOW_MARQUEE_PROGRESS_BAR) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_SHOW_MARQUEE_PROGRESS_BAR, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that the TaskDialog�s callback should be called approximately every 200 milliseconds.
        ''' </summary>
        Public Property CallbackTimer() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_CALLBACK_TIMER) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_CALLBACK_TIMER, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that the TaskDialog should be positioned (centered) relative to the owner window
        ''' passed when calling Show. If not set (or no owner window is passed), the TaskDialog is
        ''' positioned (centered) relative to the monitor.
        ''' </summary>
        Public Property PositionRelativeToWindow() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_POSITION_RELATIVE_TO_WINDOW) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_POSITION_RELATIVE_TO_WINDOW, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that the TaskDialog should have right to left layout.
        ''' </summary>
        Public Property RightToLeftLayout() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_RTL_LAYOUT) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_RTL_LAYOUT, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that the TaskDialog should have no default radio button.
        ''' </summary>
        Public Property NoDefaultRadioButton() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_NO_DEFAULT_RADIO_BUTTON) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_NO_DEFAULT_RADIO_BUTTON, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates that the TaskDialog can be minimised. Works only if there if parent window is null. Will enable cancellation also.
        ''' </summary>
        Public Property CanBeMinimized() As Boolean
            Get
                Return (Me.flags And UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_CAN_BE_MINIMIZED) <> 0
            End Get
            Set
                Me.SetFlag(UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_CAN_BE_MINIMIZED, Value)
            End Set
        End Property

        ''' <summary>
        ''' Indicates the default button for the dialog. This may be any of the values specified
        ''' in ButtonId members of one of the TaskDialogButton structures in the Buttons array,
        ''' or one a DialogResult value that corresponds to a buttons specified in the CommonButtons Member.
        ''' If this member is zero or its value does not correspond to any button ID in the dialog,
        ''' then the first button in the dialog will be the default.
        ''' </summary>
        Public Property DefaultButton() As Integer
            Get
                Return Me.m_defaultButton
            End Get
            Set
                Me.m_defaultButton = Value
            End Set
        End Property

        ''' <summary>
        ''' Indicates the default radio button for the dialog. This may be any of the values specified
        ''' in ButtonId members of one of the TaskDialogButton structures in the RadioButtons array.
        ''' If this member is zero or its value does not correspond to any radio button ID in the dialog,
        ''' then the first button in RadioButtons will be the default.
        ''' The property NoDefaultRadioButton can be set to have no default.
        ''' </summary>
        Public Property DefaultRadioButton() As Integer
            Get
                Return Me.m_defaultRadioButton
            End Get
            Set
                Me.m_defaultRadioButton = Value
            End Set
        End Property

        ''' <summary>
        ''' The string to be used to label the verification checkbox. If this member is null, the
        ''' verification checkbox is not displayed in the dialog box.
        ''' </summary>
        Public Property VerificationText() As String
            Get
                Return Me.m_verificationText
            End Get
            Set
                Me.m_verificationText = Value
            End Set
        End Property

        ''' <summary>
        ''' The string to be used for displaying additional information. The additional information is
        ''' displayed either immediately below the content or below the footer text depending on whether
        ''' the ExpandFooterArea member is true. If the EnameHyperlinks member is true, then this string
        ''' may contain hyperlinks in the form: <A HREF="executablestring">Hyperlink Text</A>.
        ''' WARNING: Enabling hyperlinks when using content from an unsafe source may cause security vulnerabilities.
        ''' </summary>
        Public Property ExpandedInformation() As String
            Get
                Return Me.m_expandedInformation
            End Get
            Set
                Me.m_expandedInformation = Value
            End Set
        End Property

        ''' <summary>
        ''' The string to be used to label the button for collapsing the expanded information. This
        ''' member is ignored when the ExpandedInformation member is null. If this member is null
        ''' and the CollapsedControlText is specified, then the CollapsedControlText value will be
        ''' used for this member as well.
        ''' </summary>
        Public Property ExpandedControlText() As String
            Get
                Return Me.m_expandedControlText
            End Get
            Set
                Me.m_expandedControlText = Value
            End Set
        End Property

        ''' <summary>
        ''' The string to be used to label the button for expanding the expanded information. This
        ''' member is ignored when the ExpandedInformation member is null.  If this member is null
        ''' and the ExpandedControlText is specified, then the ExpandedControlText value will be
        ''' used for this member as well.
        ''' </summary>
        Public Property CollapsedControlText() As String
            Get
                Return Me.m_collapsedControlText
            End Get
            Set
                Me.m_collapsedControlText = Value
            End Set
        End Property

        ''' <summary>
        ''' The string to be used in the footer area of the dialog box. If the EnableHyperlinks member
        ''' is true, then this string may contain hyperlinks in the form: <A HREF="executablestring">
        ''' Hyperlink Text</A>.
        ''' WARNING: Enabling hyperlinks when using content from an unsafe source may cause security vulnerabilities.
        ''' </summary>
        Public Property Footer() As String
            Get
                Return Me.m_footer
            End Get
            Set
                Me.m_footer = Value
            End Set
        End Property

        ''' <summary>
        ''' width of the Task Dialog's client area in DLU's. If 0, Task Dialog will calculate the ideal width.
        ''' </summary>
        Public Property Width() As UInteger
            Get
                Return Me.m_width
            End Get
            Set
                Me.m_width = Value
            End Set
        End Property

        ''' <summary>
        ''' The callback that receives messages from the Task Dialog when various events occur.
        ''' </summary>
        Public Property Callback() As TaskDialogCallback
            Get
                Return Me.m_callback
            End Get
            Set
                Me.m_callback = Value
            End Set
        End Property

        ''' <summary>
        ''' Reference that is passed to the callback.
        ''' </summary>
        Public Property CallbackData() As Object
            Get
                Return Me.m_callbackData
            End Get
            Set
                Me.m_callbackData = Value
            End Set
        End Property

        ''' <summary>
        ''' Resets the Task Dialog to the state when first constructed, all properties set to their default value.
        ''' </summary>
        Public Sub Reset()
            Me.m_windowTitle = Nothing
            Me.m_mainInstruction = Nothing
            Me.m_content = Nothing
            Me.m_commonButtons = 0
            Me.m_mainIcon = TaskDialogIcon.None
            Me.m_customMainIcon = Nothing
            Me.m_footerIcon = TaskDialogIcon.None
            Me.m_customFooterIcon = Nothing
            Me.m_buttons = New TaskDialogButton(-1) {}
            Me.m_radioButtons = New TaskDialogButton(-1) {}
            Me.flags = 0
            Me.m_defaultButton = 0
            Me.m_defaultRadioButton = 0
            Me.m_verificationText = Nothing
            Me.m_expandedInformation = Nothing
            Me.m_expandedControlText = Nothing
            Me.m_collapsedControlText = Nothing
            Me.m_footer = Nothing
            Me.m_callback = Nothing
            Me.m_callbackData = Nothing
            Me.m_width = 0
        End Sub

        ''' <summary>
        ''' Creates, displays, and operates a task dialog. The task dialog contains application-defined messages, title,
        ''' verification check box, command links and push buttons, plus any combination of predefined icons and push buttons
        ''' as specified on the other members of the class before calling Show.
        ''' </summary>
        ''' <returns>The result of the dialog, either a DialogResult value for common push buttons set in the CommonButtons
        ''' member or the ButtonID from a TaskDialogButton structure set on the Buttons member.</returns>
        Public Function Show() As Integer
            Dim verificationFlagChecked As Boolean
            Dim radioButtonResult As Integer
            Return Me.Show(IntPtr.Zero, verificationFlagChecked, radioButtonResult)
        End Function

        ''' <summary>
        ''' Creates, displays, and operates a task dialog. The task dialog contains application-defined messages, title,
        ''' verification check box, command links and push buttons, plus any combination of predefined icons and push buttons
        ''' as specified on the other members of the class before calling Show.
        ''' </summary>
        ''' <param name="owner">Owner window the task Dialog will modal to.</param>
        ''' <returns>The result of the dialog, either a DialogResult value for common push buttons set in the CommonButtons
        ''' member or the ButtonID from a TaskDialogButton structure set on the Buttons member.</returns>
        Public Function Show(owner As IWin32Window) As Integer
            Dim verificationFlagChecked As Boolean
            Dim radioButtonResult As Integer
            Return Me.Show((If(owner Is Nothing, IntPtr.Zero, owner.Handle)), verificationFlagChecked, radioButtonResult)
        End Function

        ''' <summary>
        ''' Creates, displays, and operates a task dialog. The task dialog contains application-defined messages, title,
        ''' verification check box, command links and push buttons, plus any combination of predefined icons and push buttons
        ''' as specified on the other members of the class before calling Show.
        ''' </summary>
        ''' <param name="hwndOwner">Owner window the task Dialog will modal to.</param>
        ''' <returns>The result of the dialog, either a DialogResult value for common push buttons set in the CommonButtons
        ''' member or the ButtonID from a TaskDialogButton structure set on the Buttons member.</returns>
        Public Function Show(hwndOwner As IntPtr) As Integer
            Dim verificationFlagChecked As Boolean
            Dim radioButtonResult As Integer
            Return Me.Show(hwndOwner, verificationFlagChecked, radioButtonResult)
        End Function

        ''' <summary>
        ''' Creates, displays, and operates a task dialog. The task dialog contains application-defined messages, title,
        ''' verification check box, command links and push buttons, plus any combination of predefined icons and push buttons
        ''' as specified on the other members of the class before calling Show.
        ''' </summary>
        ''' <param name="owner">Owner window the task Dialog will modal to.</param>
        ''' <param name="verificationFlagChecked">Returns true if the verification checkbox was checked when the dialog
        ''' was dismissed.</param>
        ''' <returns>The result of the dialog, either a DialogResult value for common push buttons set in the CommonButtons
        ''' member or the ButtonID from a TaskDialogButton structure set on the Buttons member.</returns>
        Public Function Show(owner As IWin32Window, ByRef verificationFlagChecked As Boolean) As Integer
            Dim radioButtonResult As Integer
            Return Me.Show((If(owner Is Nothing, IntPtr.Zero, owner.Handle)), verificationFlagChecked, radioButtonResult)
        End Function

        ''' <summary>
        ''' Creates, displays, and operates a task dialog. The task dialog contains application-defined messages, title,
        ''' verification check box, command links and push buttons, plus any combination of predefined icons and push buttons
        ''' as specified on the other members of the class before calling Show.
        ''' </summary>
        ''' <param name="hwndOwner">Owner window the task Dialog will modal to.</param>
        ''' <param name="verificationFlagChecked">Returns true if the verification checkbox was checked when the dialog
        ''' was dismissed.</param>
        ''' <returns>The result of the dialog, either a DialogResult value for common push buttons set in the CommonButtons
        ''' member or the ButtonID from a TaskDialogButton structure set on the Buttons member.</returns>
        Public Function Show(hwndOwner As IntPtr, ByRef verificationFlagChecked As Boolean) As Integer
            ' We have to call a private version or PreSharp gets upset about a unsafe
            ' block in a public method. (PreSharp error 56505)
            Dim radioButtonResult As Integer
            Return Me.PrivateShow(hwndOwner, verificationFlagChecked, radioButtonResult)
        End Function

        ''' <summary>
        ''' Creates, displays, and operates a task dialog. The task dialog contains application-defined messages, title,
        ''' verification check box, command links and push buttons, plus any combination of predefined icons and push buttons
        ''' as specified on the other members of the class before calling Show.
        ''' </summary>
        ''' <param name="owner">Owner window the task Dialog will modal to.</param>
        ''' <param name="verificationFlagChecked">Returns true if the verification checkbox was checked when the dialog
        ''' was dismissed.</param>
        ''' <param name="radioButtonResult">The radio button selected by the user.</param>
        ''' <returns>The result of the dialog, either a DialogResult value for common push buttons set in the CommonButtons
        ''' member or the ButtonID from a TaskDialogButton structure set on the Buttons member.</returns>
        Public Function Show(owner As IWin32Window, ByRef verificationFlagChecked As Boolean, ByRef radioButtonResult As Integer) As Integer
            Return Me.Show((If(owner Is Nothing, IntPtr.Zero, owner.Handle)), verificationFlagChecked, radioButtonResult)
        End Function

        ''' <summary>
        ''' Creates, displays, and operates a task dialog. The task dialog contains application-defined messages, title,
        ''' verification check box, command links and push buttons, plus any combination of predefined icons and push buttons
        ''' as specified on the other members of the class before calling Show.
        ''' </summary>
        ''' <param name="hwndOwner">Owner window the task Dialog will modal to.</param>
        ''' <param name="verificationFlagChecked">Returns true if the verification checkbox was checked when the dialog
        ''' was dismissed.</param>
        ''' <param name="radioButtonResult">The radio button selected by the user.</param>
        ''' <returns>The result of the dialog, either a DialogResult value for common push buttons set in the CommonButtons
        ''' member or the ButtonID from a TaskDialogButton structure set on the Buttons member.</returns>
        Public Function Show(hwndOwner As IntPtr, ByRef verificationFlagChecked As Boolean, ByRef radioButtonResult As Integer) As Integer
            ' We have to call a private version or PreSharp gets upset about a unsafe
            ' block in a public method. (PreSharp error 56505)
            Return Me.PrivateShow(hwndOwner, verificationFlagChecked, radioButtonResult)
        End Function

        ''' <summary>
        ''' Creates, displays, and operates a task dialog. The task dialog contains application-defined messages, title,
        ''' verification check box, command links and push buttons, plus any combination of predefined icons and push buttons
        ''' as specified on the other members of the class before calling Show.
        ''' </summary>
        ''' <param name="hwndOwner">Owner window the task Dialog will modal to.</param>
        ''' <param name="verificationFlagChecked">Returns true if the verification checkbox was checked when the dialog
        ''' was dismissed.</param>
        ''' <param name="radioButtonResult">The radio button selected by the user.</param>
        ''' <returns>The result of the dialog, either a DialogResult value for common push buttons set in the CommonButtons
        ''' member or the ButtonID from a TaskDialogButton structure set on the Buttons member.</returns>
        Private Function PrivateShow(hwndOwner As IntPtr, ByRef verificationFlagChecked As Boolean, ByRef radioButtonResult As Integer) As Integer
            verificationFlagChecked = False
            radioButtonResult = 0
            Dim result As Integer = 0
            Dim config As New UnsafeNativeMethods.TASKDIALOGCONFIG()

            Try
                config.cbSize = CUInt(Marshal.SizeOf(GetType(UnsafeNativeMethods.TASKDIALOGCONFIG)))
                config.hwndParent = hwndOwner
                config.dwFlags = Me.flags
                config.dwCommonButtons = Me.m_commonButtons

                If Not String.IsNullOrEmpty(Me.m_windowTitle) Then
                    config.pszWindowTitle = Me.m_windowTitle
                End If

                config.MainIcon = New IntPtr(Me.m_mainIcon)
                If Me.m_customMainIcon IsNot Nothing Then
                    config.dwFlags = config.dwFlags Or UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_USE_HICON_MAIN
                    config.MainIcon = Me.m_customMainIcon.Handle
                End If

                If Not String.IsNullOrEmpty(Me.m_mainInstruction) Then
                    config.pszMainInstruction = Me.m_mainInstruction
                End If

                If Not String.IsNullOrEmpty(Me.m_content) Then
                    config.pszContent = Me.m_content
                End If

                Dim customButtons As TaskDialogButton() = Me.m_buttons
                If customButtons.Length > 0 Then
                    ' Hand marshal the buttons array.
                    Dim elementSize As Integer = Marshal.SizeOf(GetType(TaskDialogButton))
                    config.pButtons = Marshal.AllocHGlobal(elementSize * CInt(customButtons.Length))
                    For i As Integer = 0 To customButtons.Length - 1
                        ' Unsafe because of pointer arithmetic.
                        Dim p = config.pButtons
                        Marshal.StructureToPtr(customButtons(i), p + elementSize * i, False)

                        config.cButtons += 1UI
                    Next
                End If

                Dim customRadioButtons As TaskDialogButton() = Me.m_radioButtons
                If customRadioButtons.Length > 0 Then
                    ' Hand marshal the buttons array.
                    Dim elementSize As Integer = Marshal.SizeOf(GetType(TaskDialogButton))
                    config.pRadioButtons = Marshal.AllocHGlobal(elementSize * CInt(customRadioButtons.Length))
                    For i As Integer = 0 To customRadioButtons.Length - 1
                        ' Unsafe because of pointer arithmetic.
                        Dim p = config.pRadioButtons
                        Marshal.StructureToPtr(customRadioButtons(i), p + (elementSize * i), False)

                        config.cRadioButtons += 1UI
                    Next
                End If

                config.nDefaultButton = Me.m_defaultButton
                config.nDefaultRadioButton = Me.m_defaultRadioButton

                If Not String.IsNullOrEmpty(Me.m_verificationText) Then
                    config.pszVerificationText = Me.m_verificationText
                End If

                If Not String.IsNullOrEmpty(Me.m_expandedInformation) Then
                    config.pszExpandedInformation = Me.m_expandedInformation
                End If

                If Not String.IsNullOrEmpty(Me.m_expandedControlText) Then
                    config.pszExpandedControlText = Me.m_expandedControlText
                End If

                If Not String.IsNullOrEmpty(Me.m_collapsedControlText) Then
                    config.pszCollapsedControlText = Me.CollapsedControlText
                End If

                config.FooterIcon = New IntPtr(Me.m_footerIcon)
                If Me.m_customFooterIcon IsNot Nothing Then
                    config.dwFlags = config.dwFlags Or UnsafeNativeMethods.TASKDIALOG_FLAGS.TDF_USE_HICON_FOOTER
                    config.FooterIcon = Me.m_customFooterIcon.Handle
                End If

                If Not String.IsNullOrEmpty(Me.m_footer) Then
                    config.pszFooter = Me.m_footer
                End If

                ' If our user has asked for a callback then we need to ask for one to
                ' translate to the friendly version.
                If Me.m_callback IsNot Nothing Then
                    config.pfCallback = New UnsafeNativeMethods.TaskDialogCallback(AddressOf Me.PrivateCallback)
                End If

                ''config.lpCallbackData = this.callbackData; // How do you do this? Need to pin the ref?
                config.cxWidth = Me.m_width

                ' The call all this mucking about is here for.
                Using New EnableThemingInScope()
                    UnsafeNativeMethods.TaskDialogIndirect(config, result, radioButtonResult, verificationFlagChecked)
                End Using
            Finally
                ' Free the unmanaged memory needed for the button arrays.
                ' There is the possibility of leaking memory if the app-domain is destroyed in a non clean way
                ' and the hosting OS process is kept alive but fixing this would require using hardening techniques
                ' that are not required for the users of this class.
                If config.pButtons <> IntPtr.Zero Then
                    Dim elementSize As Integer = Marshal.SizeOf(GetType(TaskDialogButton))
                    For i = 0 To config.cButtons - 1
                        Dim p = config.pButtons
                        Marshal.DestroyStructure(p + CInt(elementSize * i), GetType(TaskDialogButton))
                    Next

                    Marshal.FreeHGlobal(config.pButtons)
                End If

                If config.pRadioButtons <> IntPtr.Zero Then
                    Dim elementSize As Integer = Marshal.SizeOf(GetType(TaskDialogButton))
                    For i = 0 To config.cRadioButtons - 1
                        Dim p = config.pRadioButtons
                        Marshal.DestroyStructure(p + CInt(elementSize * i), GetType(TaskDialogButton))
                    Next

                    Marshal.FreeHGlobal(config.pRadioButtons)
                End If
            End Try

            Return result
        End Function

        ''' <summary>
        ''' The callback from the native Task Dialog. This prepares the friendlier arguments and calls the simpler callback.
        ''' </summary>
        ''' <param name="hwnd">The window handle of the Task Dialog that is active.</param>
        ''' <param name="msg">The notification. A TaskDialogNotification value.</param>
        ''' <param name="wparam">Specifies additional notification information.  The contents of this parameter depends on the value of the msg parameter.</param>
        ''' <param name="lparam">Specifies additional notification information.  The contents of this parameter depends on the value of the msg parameter.</param>
        ''' <param name="refData">Specifies the application-defined value given in the call to TaskDialogIndirect.</param>
        ''' <returns>A HRESULT. It's not clear in the spec what a failed result will do.</returns>
        Private Function PrivateCallback(<[In]> hwnd As IntPtr, <[In]> msg As UInteger, <[In]> wparam As UIntPtr, <[In]> lparam As IntPtr, <[In]> refData As IntPtr) As Integer
            Dim callback As TaskDialogCallback = Me.m_callback
            If callback IsNot Nothing Then
                ' Prepare arguments for the callback to the user we are insulating from Inter-op casting silliness.

                ' Future: Consider reusing a single ActiveTaskDialog object and mark it as destroyed on the destroy notification.
                Dim activeDialog As New ActiveTaskDialog(hwnd)
                Dim args As New TaskDialogNotificationArgs()
                args.Notification = CType(msg, TaskDialogNotification)
                Select Case args.Notification
                    Case TaskDialogNotification.ButtonClicked, TaskDialogNotification.RadioButtonClicked
                        args.ButtonId = CInt(wparam)
                        Exit Select
                    Case TaskDialogNotification.HyperlinkClicked
                        args.Hyperlink = Marshal.PtrToStringUni(lparam)
                        Exit Select
                    Case TaskDialogNotification.Timer
                        args.TimerTickCount = CUInt(wparam)
                        Exit Select
                    Case TaskDialogNotification.VerificationClicked
                        args.VerificationFlagChecked = (wparam <> UIntPtr.Zero)
                        Exit Select
                    Case TaskDialogNotification.ExpandoButtonClicked
                        args.Expanded = (wparam <> UIntPtr.Zero)
                        Exit Select
                End Select

                Return (If(callback(activeDialog, args, Me.m_callbackData), 1, 0))
            End If

            Return 0
            ' false;
        End Function

        ''' <summary>
        ''' Helper function to set or clear a bit in the flags field.
        ''' </summary>
        ''' <param name="flag">The Flag bit to set or clear.</param>
        ''' <param name="value">True to set, false to clear the bit in the flags field.</param>
        Private Sub SetFlag(flag As UnsafeNativeMethods.TASKDIALOG_FLAGS, value As Boolean)
            If value Then
                Me.flags = Me.flags Or flag
            Else
                Me.flags = Me.flags And Not flag
            End If
        End Sub
    End Class
End Namespace

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================
