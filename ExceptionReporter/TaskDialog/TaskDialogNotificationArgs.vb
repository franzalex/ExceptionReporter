
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
Namespace Global.Microsoft.Samples

    ''' <summary>
    ''' Arguments passed to the TaskDialog callback.
    ''' </summary>
    Public Class TaskDialogNotificationArgs
        ''' <summary>
        ''' What the TaskDialog callback is a notification of.
        ''' </summary>
        Private m_notification As TaskDialogNotification

        ''' <summary>
        ''' The button ID if the notification is about a button. This a DialogResult
        ''' value or the ButtonID member of a TaskDialogButton set in the
        ''' TaskDialog.Buttons or TaskDialog.RadioButtons members.
        ''' </summary>
        Private m_buttonId As Integer

        ''' <summary>
        ''' The HREF string of the hyperlink the notification is about.
        ''' </summary>
        Private m_hyperlink As String

        ''' <summary>
        ''' The number of milliseconds since the dialog was opened or the last time the
        ''' callback for a timer notification reset the value by returning true.
        ''' </summary>
        Private m_timerTickCount As UInteger

        ''' <summary>
        ''' The state of the verification flag when the notification is about the verification flag.
        ''' </summary>
        Private m_verificationFlagChecked As Boolean

        ''' <summary>
        ''' The state of the dialog expando when the notification is about the expando.
        ''' </summary>
        Private m_expanded As Boolean

        ''' <summary>
        ''' What the TaskDialog callback is a notification of.
        ''' </summary>
        Public Property Notification() As TaskDialogNotification
            Get
                Return Me.m_notification
            End Get
            Set
                Me.m_notification = Value
            End Set
        End Property

        ''' <summary>
        ''' The button ID if the notification is about a button. This a DialogResult
        ''' value or the ButtonID member of a TaskDialogButton set in the
        ''' TaskDialog.Buttons member.
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
        ''' The HREF string of the hyperlink the notification is about.
        ''' </summary>
        Public Property Hyperlink() As String
            Get
                Return Me.m_hyperlink
            End Get
            Set
                Me.m_hyperlink = Value
            End Set
        End Property

        ''' <summary>
        ''' The number of milliseconds since the dialog was opened or the last time the
        ''' callback for a timer notification reset the value by returning true.
        ''' </summary>
        Public Property TimerTickCount() As UInteger
            Get
                Return Me.m_timerTickCount
            End Get
            Set
                Me.m_timerTickCount = Value
            End Set
        End Property

        ''' <summary>
        ''' The state of the verification flag when the notification is about the verification flag.
        ''' </summary>
        Public Property VerificationFlagChecked() As Boolean
            Get
                Return Me.m_verificationFlagChecked
            End Get
            Set
                Me.m_verificationFlagChecked = Value
            End Set
        End Property

        ''' <summary>
        ''' The state of the dialog expando when the notification is about the expando.
        ''' </summary>
        Public Property Expanded() As Boolean
            Get
                Return Me.m_expanded
            End Get
            Set
                Me.m_expanded = Value
            End Set
        End Property
    End Class
End Namespace

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================
