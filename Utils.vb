Imports System.Security

''' <summary>
''' Miscellaneous  methods and functions.
''' </summary>
''' <remarks></remarks>
Public Module Utils
    ''' <summary>
    ''' This function should be used to ensure that all dependencies on the use of a debugger can be identified.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DebuggerAttached() As Boolean
        Return Debugger.IsAttached
    End Function

    ''' <summary>
    ''' Returns a System.Security.SecureString, built from the specified string.
    ''' </summary>
    ''' <param name="s">The string from which to build the SecureString.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SecureStrCreate(ByVal s As String) As SecureString
        Dim output As New SecureString
        Dim last As Integer = s.Length - 1
        For i As Integer = 0 To last
            output.AppendChar(s(i))
        Next
        Return output
    End Function

    ''' <summary>
    ''' Accessed through the KeyIsDown function. Detects whether the key in question is currently down or has been pressed since the last call.
    ''' Full details at http://msdn.microsoft.com/en-us/library/ms646293(v=vs.85).aspx
    ''' </summary>
    ''' <param name="vKey">The key code</param>
    ''' <returns>
    ''' -32767 if the key is currently down.
    ''' 1 if the key is not currently down but has been pressed since the last call.
    ''' 0 if neither of the above.
    ''' </returns>
    ''' <remarks>Using this function to find out whether the key has been pressed since the last call is not reliable.</remarks>
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short
    ''' <summary>Uses a Windows API to detect whether the key in question is currently down.</summary>
    ''' <param name="key">The key code</param>
    ''' <returns>True if the key in question is currently down. Otherwise false.</returns>
    ''' <exception cref="Exception">If the value returned by the windows function is not -32767, 1 or 0, an exception is thrown.</exception>
    Public Function KeyIsDown(ByVal key As Keys) As Boolean
        Select Case GetAsyncKeyState(key)
            Case -32767 : Return True
            Case 1, 0 : Return False
            Case Else : Throw New Exception("The Windows API Function GetAsyncKeyState has returned a value that is not recognised by this application.")
                Return False
        End Select
    End Function

    Public Sub Pixelate(ByRef bmp As Bitmap, ByRef indexes As Integer(,), ByVal pixelColor As Color)
        For num As Integer = 0 To (indexes.GetLength(0) - 1)
            bmp.SetPixel(indexes(num, 0), indexes(num, 1), pixelColor)
        Next
    End Sub

    ''' <summary>
    ''' Returns True if the type is numeric, otherwise False. Returns False if Boolean.
    ''' </summary>
    ''' <param name="t"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsNumericType(ByVal t As Type) As Boolean
        Select Case t
            Case GetType(Byte), GetType(Integer), GetType(Int16), GetType(Int32), GetType(Int64), GetType(UInteger), GetType(UInt16),
            GetType(UInt32), GetType(UInt64), GetType(Short), GetType(Long), GetType(UShort), GetType(ULong), GetType(Single), GetType(Double)
                Return True
            Case Else
                Return False
        End Select
    End Function
End Module