Imports ExceptionReporter

Public Class frmMain
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnBugReport.Click
        ExceptionReporter.ShowBugReportDialog(Me)
    End Sub

    Private Sub btnThrowException_Click(sender As Object, e As EventArgs) Handles btnThrowException.Click
        Try
            Dim faaPhoneticSpelling = {"Alpha", "Bravo", "Charlie", "Delta", "Echo",
                                       "Foxtrot", "Golf", "Hotel", "India", "Juliet",
                                       "Kilo", "Lima", "Mike", "November", "Oscar",
                                       "Papa", "Quebec", "Romeo", "Sierra", "Tango",
                                       "Uniform", "Victor", "Whiskey", "X-Ray", "Yankee",
                                       "Zulu"}
            For i = 0 To 50
                faaPhoneticSpelling.ElementAt(i)
            Next
        Catch ex As Exception
            ex.PromptToReport()
        End Try
    End Sub
End Class
