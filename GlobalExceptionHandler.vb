Imports System.Threading

''' <summary>
''' Handles the unhandled exceptions in an application.
''' </summary>
''' <remarks></remarks>
Public Class GlobalExceptionHandler
    Private Shared inUse As Boolean = False
    Public Shared ReadOnly Property IsInUse As Boolean
        Get
            Return inUse
        End Get
    End Property

    Public Shared UnhandledExceptionHandler As UnhandledExceptionEventHandler = AddressOf UnhandledEx
    Public Shared ThreadExceptionHandler As ThreadExceptionEventHandler = AddressOf ThreadEx
    Private Shared Sub UnhandledEx(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        Dim exc As Exception
        Try
            exc = CType(e.ExceptionObject, Exception)
        Catch ex As Exception
            exc = New Exception("An exception occurred in the generic exception handler. See the InnerException for details.", ex)
        End Try
        ExceptionHandler(exc)
    End Sub
    Private Shared Sub ThreadEx(ByVal sender As Object, ByVal e As ThreadExceptionEventArgs)
        ExceptionHandler(e.Exception)
    End Sub

    ''' <summary>
    ''' Adds the event handlers for the global exception handler, provided that the debugger is not attached.
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub ExceptionHandlerInitialize()
        If Not DebuggerAttached() Then
            Try
                AddHandler AppDomain.CurrentDomain.UnhandledException, UnhandledExceptionHandler
                AddHandler Application.ThreadException, ThreadExceptionHandler
                inUse = True
            Catch ex As Exception
                ExceptionHandler(ex)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Generic application exception handler for unhandled exceptions.
    ''' </summary>
    ''' <param name="ex">The unhandled exception</param>
    ''' <remarks></remarks>
    Public Shared Sub ExceptionHandler(ByVal ex As Exception)
        ''Get the time when the exception occurred
        Dim instant As DateTime = Now

        ''Take a screenshot
        Dim screenshotMsg As String = ExceptionScreenshot(instant)

        ''Write an entry in the error log text file
        Dim logEntryMsg As String = ExceptionTextLog(ex, instant)

        ''Inform the user
        Dim msg As String = "An unhandled exception occurred and the application must now close. Further details follow." & vbCrLf & vbCrLf & vbCrLf &
            screenshotMsg & vbCrLf & vbCrLf & logEntryMsg
        MsgBox(msg, 16, "Unhandled Exception")

        ''Close the program
        Process.GetCurrentProcess.Kill()
    End Sub

    ''' <summary>
    ''' Writes an error message in the application's error log text file containing the time, type, message and stack trace of the exception.
    ''' </summary>
    ''' <param name="ex">The exception</param>
    ''' <param name="dtm">The time at which the exception occurred</param>
    ''' <returns>A string to inform the user of the success, or not, of the operation.</returns>
    ''' <remarks>The time format of the log entry is dd/mm/yyyy hh:mm:ss</remarks>
    Private Shared Function ExceptionTextLog(ByVal ex As Exception, ByVal dtm As DateTime) As String
        Dim userMsg As String
        Try
            ''Build the log entry string and set the log file name
            Dim logEntry As String = dtm.ToString & vbCrLf & ex.GetType.ToString & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & vbCrLf
            If ex.InnerException IsNot Nothing Then
                logEntry = logEntry.Remove(logEntry.Length - 2) & "InnerException: " & ex.InnerException.GetType.ToString & vbCrLf & ex.InnerException.Message &
                    vbCrLf & ex.InnerException.StackTrace & vbCrLf & vbCrLf
            End If
            Dim errLogName As String = "ErrorLog.txt"

            ''Write the log entry
            FileIO.FileSystem.WriteAllText(errLogName, logEntry, True)

            ''Return the string
            Dim errLogPath As String = FileIO.FileSystem.GetFileInfo(errLogName).FullName
            userMsg = "The following message was recorded in the application's error log at" & vbCrLf & errLogPath & ":" & vbCrLf & vbCrLf & logEntry
        Catch
            userMsg = "A message could not be recorded in the application's error log."
        End Try
        Return userMsg
    End Function

    ''' <summary>
    ''' Takes a screenshot bitmap and saves it to the application's startup directory.
    ''' </summary>
    ''' <param name="dtm">The time at which the exception occurred</param>
    ''' <returns>A string to inform the user of the success, or not, of the operation.</returns>
    ''' <remarks>The bitmap is named as follows: Exception_yyyy-mm-dd_hh-mm-ss.bmp. Note the date and time format.</remarks>
    Private Shared Function ExceptionScreenshot(ByVal dtm As DateTime) As String
        Dim userMsg As String
        Try
            ''Create the bitmap
            Dim bounds As Rectangle = Screen.PrimaryScreen.Bounds
            Dim snshot As Bitmap = New Bitmap(bounds.Width, bounds.Height)
            Dim graph As Graphics = Graphics.FromImage(snshot)
            graph.CopyFromScreen(bounds.X, bounds.Y, 0, 0, bounds.Size)

            ''Set the path to which the bitmap should be saved, including a ymdhms time format
            Dim filePath As String
            Dim m0d1(1) As String       ''Month is variable 0; day is variable 1.
            m0d1(0) = dtm.Month.ToString
            m0d1(1) = dtm.Day.ToString
            For i As Integer = 0 To 1
                If m0d1(i).Length < 2 Then
                    m0d1(i) = "0" + m0d1(i)
                End If
            Next
            Dim dateFormatted As String = dtm.Year.ToString & "-" & m0d1(0) & "-" & m0d1(1)
            dateFormatted = Text.RegularExpressions.Regex.Replace(dateFormatted, "[\/\:\ ]", "-")
            Dim timeFormatted As String = dtm.TimeOfDay.ToString.Remove(8)
            timeFormatted = Text.RegularExpressions.Regex.Replace(timeFormatted, "[:]", "-")
            filePath = Application.StartupPath & "\Exception_" & dateFormatted & "_" & timeFormatted & ".bmp"

            ''Save the bitmap
            snshot.Save(filePath)

            ''Return the string
            If FileIO.FileSystem.FileExists(filePath) Then
                userMsg = "A screenshot was taken and saved to " & vbCrLf & filePath
            Else
                userMsg = "A screenshot was not taken."
            End If
        Catch
            userMsg = "A screenshot was not taken."
        End Try
        Return userMsg
    End Function

    ''' <summary>
    ''' Throws an exception. This Sub should be used so that all calls can be removed before publishing.
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub ExceptionThrow(Optional ByVal hasInnerException As Boolean = False)
        Dim innerex As Exception
        If hasInnerException Then
            innerex = New Exception("Test InnerException")
        Else
            innerex = Nothing
        End If
        Throw New Exception("This exception was thrown deliberately as a check during debugging.", innerex)
    End Sub
End Class