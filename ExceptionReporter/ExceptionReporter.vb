'*****************************************************************************
' Name:        ExceptionReporter.vb
' Description: Contains methods for capturing and logging exceptions.
' Author:      Franz Alex Gaisie-Essilfie
'
' Change Log:
'  Date        | Description
' -------------|--------------------------------------------------------------
'  2017-09-09  | Initial limited public release
'

Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms
Imports Microsoft.Samples
Imports Newtonsoft.Json

Namespace ExceptionReporter
    ''' <summary>
    ''' Contains methods for capturing and logging excceptions.
    ''' </summary>
    Public Module ExceptionReporter
        Private _preserveStackTrace As Reflection.MethodInfo

        ''' <summary>Displays a dialog that prompts a user to save an exception.</summary>
        ''' <param name="ex">The ex.</param>
        ''' <param name="exPrompt">The ex prompt.</param>
        ''' <param name="isFatal">
        '''   If set to <c>true</c>, indicattes the exception will cause the program to be terminated.
        ''' </param>
        <Extension()> Public Sub PromptToReport(ex As Exception, Optional exPrompt As String = Nothing, Optional isFatal As Boolean = False)
            ex.PreserveStackTrace()
            Dim errorDumpFile = IO.Path.Combine(IO.Path.GetTempPath(), IO.Path.GetRandomFileName())
            errorDumpFile = errorDumpFile.Remove(errorDumpFile.LastIndexOf("."c)) + ".txt"
            Dim shouldRestart As Boolean = False

            ' fill the default exception prompt if none was specified
            If String.IsNullOrWhiteSpace(exPrompt) Then
                exPrompt = "An unhandled exception has occurred."
                If isFatal Then exPrompt = exPrompt.TrimEnd("."c) & " that will cause the program to close."
            End If

            Try
                Dim exDumpFile = ""
                Dim activeAppHwnd = GetActiveApplicationWindow()
                IO.File.WriteAllText(errorDumpFile, ex.Dump())

                ' use IDs greater than 100 to avoid clashing with IDs of TaskDialogCommonButtons
                Const btnReportId As Integer = 101
                Const btnIgnoreId As Integer = 102

                If New TaskDialog() With {
                .WindowTitle = $"Exception - {Application.ProductName}",
                .MainInstruction = exPrompt,
                .Content = $"Exception Message:{ vbCrLf & ex.Message}".Trim(),
                .MainIcon = (If(isFatal, TaskDialogIcon.Error, TaskDialogIcon.Warning)),
                .Buttons = New TaskDialogButton() {New TaskDialogButton(btnReportId, "&Report to Developer"),
                                                   New TaskDialogButton(btnIgnoreId, "Do &not Report")},
                .Footer = $"Click <A HREF=""{New Uri(errorDumpFile).AbsoluteUri}"">here</A> to view the exception details.",
                .EnableHyperlinks = True,
                .Callback = AddressOf HandleHyperlinks,
                .VerificationFlagChecked = isFatal,
                .VerificationText = (If((Not isFatal), "Restart the application", Nothing))
            }.Show(activeAppHwnd, shouldRestart) = btnReportId AndAlso
                   PromptToSaveException(ex, exDumpFile) Then

                    Call New TaskDialog() With {
                    .WindowTitle = Application.ProductName,
                    .MainInstruction = "Exception Report",
                    .Content = String.Join(Environment.NewLine,
                                           {"The exception report has been saved successfully.", "",
                                            "Submit it to the developer via the usual communication channel.",
                                            "Also include the steps to reproduce the exception."}),
                    .MainIcon = TaskDialogIcon.Information,
                    .CommonButtons = TaskDialogCommonButtons.Ok,
                    .EnableHyperlinks = True,
                    .CollapsedControlText = "&Show Details",
                    .ExpandedControlText = "&Hide Details",
                    .ExpandedInformation = $"Report File: <A HREF=""{New Uri(exDumpFile).AbsoluteUri}"">{exDumpFile}</A>",
                    .ExpandedByDefault = True,
                    .Callback = AddressOf HandleHyperlinks
                }.Show(activeAppHwnd)
                End If

            Catch
                Dim tempOutput = IO.Path.Combine(Environment.ExpandEnvironmentVariables("%TEMP%"),
                                             Date.Now.ToString("'ExDump' yyyyMMdd_HHmmss'.zip'"))
                SaveExceptionToFile(ex, tempOutput)
                MessageBox.Show(String.Join(Environment.NewLine,
                                        "There was a problem while initializing the exception report dialog.", "",
                                        "This is the location of your exception report if you wish to submit it to the developer:",
                                        tempOutput),
                            My.Application.Info.ProductName,
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            Finally
                ' clean-up on complete
                IO.File.Delete(errorDumpFile)
            End Try

            If Not Debugger.IsAttached AndAlso shouldRestart Then
                Application.Restart()
                Application.ExitThread()
            End If
        End Sub

        ''' <summary>Displays a dialog box that allows the user to submit a bug report.</summary>
        ''' <param name="owner">The window to which the bug report dialog will be shown modally to.</param>
        Public Sub ShowBugReportDialog(owner As IWin32Window)
            Using lblCaption = New Label() With {
                .AutoSize = False,
                .Dock = DockStyle.Top,
                .Padding = New Padding(2, 2, 2, 6),
                .Text = String.Join(vbCrLf,
                                    {"Describe the bug that you want to report.",
                                     "Be as detailed as possible in your description."})},
            btnReport = New Button() With {
                .AutoSize = True,
                .Text = "&Report"},
            btnCancel = New Button() With {
                .AutoSize = True,
                .Text = "&Cancel"},
            errText = New TextBox() With {
                .Dock = DockStyle.Fill,
                .Multiline = True,
                .ScrollBars = ScrollBars.Vertical},
            tlp = New TableLayoutPanel() With {
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Dock = DockStyle.Bottom,
                .Padding = New Padding(0, 2, 0, 0)},
            errorDialog = New Form() With {
                .AcceptButton = btnReport,
                .CancelButton = btnCancel,
                .Font = New Font("Segoe UI", 10.0F, FontStyle.Regular),
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .Size = New Size(320, 240),
                .StartPosition = FormStartPosition.CenterParent,
                .Text = "Report a bug"}

                ' place the buttons in the table panel for a better layout
                tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent) With {.Width = 100})
                tlp.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
                tlp.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
                tlp.Controls.AddRange({btnReport, btnCancel})
                tlp.SetCellPosition(btnReport, New TableLayoutPanelCellPosition(1, 0))
                tlp.SetCellPosition(btnCancel, New TableLayoutPanelCellPosition(2, 0))

                ' add the controls to the dialog
                errorDialog.Controls.AddRange({errText, lblCaption, tlp})

                ' the caption's size needs to be readjusted to allow all the text to be shown
                lblCaption.Size = lblCaption.GetPreferredSize(New Size(errorDialog.ClientSize.Width, 0))

                ' need to manually handle the accept button
                AddHandler btnReport.Click, Sub()
                                                errorDialog.DialogResult = DialogResult.OK
                                                errorDialog.Close()
                                            End Sub


                If errorDialog.ShowDialog(owner) = DialogResult.OK Then
                    Dim ex = New ApplicationException(errText.Text)
                    ex.PromptToReport("User generated exception for submission of an error report.")
                End If
            End Using
        End Sub

        Private Function CaptureWindow(handle As IntPtr) As Bitmap
            Dim rect As Rect = Nothing
            GetWindowRect(handle, rect)
            Dim bounds = New Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top)
            Dim result = New Bitmap(bounds.Width, bounds.Height)
            Using g = Graphics.FromImage(result)
                Dim hDC = g.GetHdc()
                PrintWindow(handle, hDC, 0)
                g.ReleaseHdc(hDC)
                ''g.CopyFromScreen(bounds.Location, Point.Empty, bounds.Size)
            End Using
            Return result
        End Function

        <Extension()> Private Function Dump(Of T)(ByVal value As T) As String
            Dim builder As System.Text.StringBuilder
            If value Is Nothing Then Return "<null>"

            If (value.GetType().IsPrimitive OrElse TypeOf value Is String) Then Return value.ToString()

            Dim dict = TryCast(value, Collections.IDictionary)
            If (dict IsNot Nothing) Then
                builder = New System.Text.StringBuilder()
                If (dict.Count = 0) Then
                    Return "{ }"
                End If
                builder.AppendLine("{")
                For Each key In dict.Keys
                    Dim str = Dump(key).Replace(vbCr, "\r").Replace(vbLf, "\n")
                    Dim source = Dump(dict.Item(key)).Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                    Dim str2 As New String(" "c, (str.Length + 2))
                    builder.Append("  ").Append(str).Append(": ").Append(Enumerable.FirstOrDefault(source))
                    For Each str3 In Enumerable.Skip(source, 1)
                        builder.Append("  ").Append(str2).Append(str3)
                    Next
                Next
                builder.Append("}")
                Return builder.ToString()

            End If

            Dim items = TryCast(value, Collections.IEnumerable)
            If (items IsNot Nothing) Then
                builder = New System.Text.StringBuilder
                Dim enumerator = items.GetEnumerator()
                builder.AppendLine("[")
                Do While enumerator.MoveNext()
                    builder.Append("  ").Append(Dump(enumerator.Current)).AppendLine(",")
                Loop
                builder.Append("]")
                Dim result As String = builder.ToString()
                If (result = $"[{Environment.NewLine}]") Then Return "[ ]"
                Return result
            End If

            builder = New System.Text.StringBuilder()

            Dim flag As Boolean = TypeOf value Is Exception
            Dim info As Reflection.PropertyInfo
            For Each info In (From p In value.GetType.GetProperties
                              Order By p.Name
                              Select p)
                If (Not flag OrElse Not (info.Name = "TargetSite")) Then
                    builder.Append(info.Name).Append(": ")
                    Dim str As New String(" "c, (info.Name.Length + 2))
                    Dim source = Dump(info.GetValue(value, Nothing)).Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                    builder.AppendLine(Enumerable.FirstOrDefault(source))
                    For Each str2 In Enumerable.Skip(source, 1)
                        builder.Append(New String(" "c, (info.Name.Length + 2))).AppendLine(str2)
                    Next
                End If
            Next
            Return builder.ToString()
        End Function

        ''' <summary>Gets the active application window.</summary>
        Private Function GetActiveApplicationWindow() As IntPtr
            Dim activeHwnds = (From wnd In Application.OpenForms.OfType(Of Form)()
                               Select h = wnd.Handle).ToArray()

            If activeHwnds.Any() Then
                Dim handle As IntPtr = GetActiveWindow()
                If activeHwnds.Contains(handle) Then
                    Return handle
                Else
                    Return activeHwnds.Last()
                End If
            End If

            Return IntPtr.Zero  ' the active window doesn't belong to this application's open formss
        End Function

        Private Function GetAppInfo() As Dictionary(Of String, Object)
            Return New Dictionary(Of String, Object)() From {
            {"Name", Application.ProductName},
            {"Version", Application.ProductVersion},
            {"Company", Application.CompanyName},
            {"ExePath", Application.ExecutablePath},
            {"CurrentCulture", Application.CurrentCulture},
            {"InputLanguage", Application.CurrentInputLanguage},
            {"OpenForms", (From f In Application.OpenForms.OfType(Of Form)() Select f.GetType().FullName).ToArray()}
        }
        End Function

        Private Function GetEnvironmentInfo() As Dictionary(Of String, Object)
            Return New Dictionary(Of String, Object)() From {
            {"CommandLine", Environment.CommandLine},
            {"CurrentDirectory", Environment.CurrentDirectory},
            {"Is64BitOS", Environment.Is64BitOperatingSystem},
            {"OSVersion", Environment.OSVersion}}
        End Function

        'TODO: Change this function in order to modify the output filename
        ''' <summary>The filename of the exception dump file</summary>
        Private Function GetDumpFileName() As String
            Return $"{Application.ProductName} Exception - " & DateTime.Now.ToString(" [yyyyMMdd_HHmmss]")
        End Function

        Private Function HandleHyperlinks(ByVal td As ActiveTaskDialog, ByVal args As TaskDialogNotificationArgs, ByVal data As Object) As Boolean
            If Not (args.Notification <> TaskDialogNotification.HyperlinkClicked OrElse
                String.IsNullOrEmpty(args.Hyperlink)) Then
                Process.Start(args.Hyperlink)
            End If
            Return False
        End Function

        Private Function LogExceptionToStream(ex As Exception) As IO.MemoryStream
            Threading.Thread.Sleep(1000)
            Application.DoEvents()

            ' open a zip stream
            Dim zipStream = New IO.MemoryStream()
            Dim zip As ZipStorer = ZipStorer.Create(zipStream, "")

            ' write the texts to the zip stream

            Dim indented = Formatting.Indented
            Dim exLog = JsonConvert.SerializeObject(ex, indented)
            Dim appInfo = JsonConvert.SerializeObject(GetAppInfo(), indented)
            Dim envtInfo = JsonConvert.SerializeObject(GetEnvironmentInfo(), indented)

            For Each item In {New With {.Content = appInfo, .FileName = "AppInfo.json"},
                              New With {.Content = envtInfo, .FileName = "Environment.json"},
                              New With {.Content = exLog, .FileName = "Exception.json"},
                              New With {.Content = ex.Dump(), .FileName = "Exception.txt"}}

                Using strm = New IO.StreamWriter(New IO.MemoryStream()) With {.AutoFlush = True}
                    strm.Write(item.Content)
                    strm.BaseStream.Flush()
                    strm.BaseStream.Position = 0
                    zip.AddStream(ZipStorer.Compression.Deflate, item.FileName, strm.BaseStream, Nothing, "")
                End Using
            Next

            ' save each open form
            For i = 0 To Application.OpenForms.Count - 1
                Dim frm = Application.OpenForms(i)
                Using strm = New IO.MemoryStream()
                    CaptureWindow(frm.Handle).Save(strm, Imaging.ImageFormat.Png)
                    strm.Flush()
                    strm.Position = 0
                    zip.AddStream(ZipStorer.Compression.Deflate, $"Screenshot {i + 1} - {frm.Name}.png", strm, Nothing, "")
                End Using
            Next

            zip.Close()

            Return zipStream
        End Function

        <Extension> Private Function PreserveStackTrace(Of T As Exception)(ex As T) As T
            If _preserveStackTrace Is Nothing Then
                _preserveStackTrace = GetType(Exception).GetMethod("InternalPreserveStackTrace",
                                                               Reflection.BindingFlags.Instance Or
                                                               Reflection.BindingFlags.NonPublic)
            End If
            _preserveStackTrace.Invoke(ex, Nothing)
            Return ex
        End Function

        Private Function PromptToSaveException(ex As Exception, ByRef fileName As String) As Boolean
            Using fb = New FolderBrowserDialog() With {
            .Description = "Save Exception Log",
            .SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}

                If fb.ShowDialog() = DialogResult.OK Then
                    fileName = IO.Path.Combine(fb.SelectedPath, GetDumpFileName() + ".zip")
                    Return SaveExceptionToFile(ex, fileName)
                End If
            End Using

            ' something else happened
            fileName = String.Empty
            Return False
        End Function

        Private Function SaveExceptionToFile(ex As Exception, fileName As String) As Boolean
            Try
                IO.File.WriteAllBytes(fileName, LogExceptionToStream(ex).ToArray())
                Return True
            Catch
                Return False
            End Try
        End Function

#Region "P/Invoke"
        Private Declare Function GetActiveWindow Lib "user32.dll" () As IntPtr
        Private Declare Function GetWindowRect Lib "user32.dll" (hWnd As IntPtr, ByRef rect As Rect) As IntPtr
        Private Declare Function PrintWindow Lib "user32.dll" (hwnd As IntPtr, hDC As IntPtr, nFlags As Integer) As Boolean

        <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)>
        Private Structure Rect
            Public Left As Integer
            Public Top As Integer
            Public Right As Integer
            Public Bottom As Integer
        End Structure
#End Region
    End Module
End Namespace