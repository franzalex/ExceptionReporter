
'------------------------------------------------------------------
' <summary>
' A P/Invoke wrapper for TaskDialog. Usability was given preference to perf and size.
' </summary>
'
' <remarks/>
'------------------------------------------------------------------

Imports System.Windows.Forms
Namespace Global.Microsoft.Samples

    ''' <summary>
    ''' TaskDialog wrapped in a CommonDialog class. This is required to work well in
    ''' MMC 3.0. In MMC 3.0 you must use the ShowDialog methods on the MMC classes to
    ''' correctly show a modal dialog. This class will allow you to do this and keep access
    ''' to the results of the TaskDialog.
    ''' </summary>
    Public Class TaskDialogCommonDialog
        Inherits CommonDialog
        ''' <summary>
        ''' The TaskDialog we will display.
        ''' </summary>
        Private m_taskDialog As TaskDialog

        ''' <summary>
        ''' The result of the dialog, either a DialogResult value for common push buttons set in the TaskDialog.CommonButtons
        ''' member or the ButtonID from a TaskDialogButton structure set on the TaskDialog.Buttons member.
        ''' </summary>
        Private m_taskDialogResult As Integer

        ''' <summary>
        ''' The verification flag result of the dialog. True if the verification checkbox was checked when the dialog
        ''' was dismissed.
        ''' </summary>
        Private m_verificationFlagCheckedResult As Boolean

        ''' <summary>
        ''' TaskDialog wrapped in a CommonDialog class. THis is required to work well in
        ''' MMC 2.1. In MMC 2.1 you must use the ShowDialog methods on the MMC classes to
        ''' correctly show a modal dialog. This class will allow you to do this and keep access
        ''' to the results of the TaskDialog.
        ''' </summary>
        ''' <param name="taskDialog">The TaskDialog to show.</param>
        Public Sub New(taskDialog As TaskDialog)
            If taskDialog Is Nothing Then
                Throw New ArgumentNullException("taskDialog")
            End If

            Me.m_taskDialog = taskDialog
        End Sub

        ''' <summary>
        ''' The TaskDialog to show.
        ''' </summary>
        Public ReadOnly Property TaskDialog() As TaskDialog
            Get
                Return Me.m_taskDialog
            End Get
        End Property

        ''' <summary>
        ''' The result of the dialog, either a DialogResult value for common push buttons set in the TaskDialog.CommonButtons
        ''' member or the ButtonID from a TaskDialogButton structure set on the TaskDialog.Buttons member.
        ''' </summary>
        Public ReadOnly Property TaskDialogResult() As Integer
            Get
                Return Me.m_taskDialogResult
            End Get
        End Property

        ''' <summary>
        ''' The verification flag result of the dialog. True if the verification checkbox was checked when the dialog
        ''' was dismissed.
        ''' </summary>
        Public ReadOnly Property VerificationFlagCheckedResult() As Boolean
            Get
                Return Me.m_verificationFlagCheckedResult
            End Get
        End Property

        ''' <summary>
        ''' Reset the common dialog.
        ''' </summary>
        Public Overrides Sub Reset()
            Me.m_taskDialog.Reset()
        End Sub

        ''' <summary>
        ''' The required implementation of CommonDialog that shows the Task Dialog.
        ''' </summary>
        ''' <param name="hwndOwner">Owner window. This can be null.</param>
        ''' <returns>If this method returns true, then ShowDialog will return DialogResult.OK.
        ''' If this method returns false, then ShowDialog will return DialogResult.Cancel. The
        ''' user of this class must use the TaskDialogResult member to get more information.
        ''' </returns>
        Protected Overrides Function RunDialog(hwndOwner As IntPtr) As Boolean
            Me.m_taskDialogResult = Me.m_taskDialog.Show(hwndOwner, Me.m_verificationFlagCheckedResult)
            Return (Me.m_taskDialogResult <> CInt(DialogResult.Cancel))
        End Function
    End Class
End Namespace

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================
