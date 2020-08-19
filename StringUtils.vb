Module StringUtils
    ''' <summary>
    ''' Returns a string for a System.Data.DataColumn.Expression, such as might be used by a DataView.RowFilter.
    ''' </summary>
    ''' <param name="strArray">The array of column names.</param>
    ''' <param name="caption">The caption for the ExpressionBuilder.</param>
    ''' <param name="startexpression">The expression to show on starting the ExpressionBuilder.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExpressionBuild(ByVal strArray As String(), Optional ByVal caption As String = Nothing, Optional ByVal startexpression As String = "") As String
        Dim eb As New ExpressionBuilder(strArray, caption, startexpression)
        eb.ShowDialog()   ''Required so that a value isn't returned until the ExpressionBuilder has been closed.
        Return ExpressionBuilder.Result
    End Function

    ''' <summary>
    ''' Inserts spaces into strings that have words, beginning with capitals, which are not separated by spaces.
    ''' </summary>
    ''' <param name="s">The string to be modified.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SpacesInsert(ByVal s As String) As String
        If String.IsNullOrWhiteSpace(s) Then
            Return s
        Else
            Dim piece(0), tmpAry() As String
            Dim lastnum As Integer = s.Length - 1
            Dim first As String = s.Remove(1)
            piece(0) = first

            ''Build the piece() array.
            For i As Integer = 1 To lastnum
                Dim c As String = s.Chars(i).ToString

                Dim cplus1 As String = ""       ''this stops cplus1 being used before it has been assigned a value.
                Dim cminus1 As String = s.Chars(i - 1).ToString
                Dim isLastChar As Boolean
                Try
                    cplus1 = s.Chars(i + 1).ToString
                Catch ex As IndexOutOfRangeException
                    isLastChar = True
                End Try

                tmpAry = piece
                Dim countBeforeAdd As Integer = piece.Count
                If IsUpperCase(c) And (Not isLastChar Or IsNumeric(c)) And Not ((IsUpperCase(cplus1) And IsUpperCase(cminus1)) Or String.IsNullOrWhiteSpace(s)) Then
                    ''Add a space, followed by the current character.
                    ReDim piece(countBeforeAdd + 1)
                    For j As Integer = 0 To countBeforeAdd - 1
                        piece(j) = tmpAry(j)
                    Next
                    piece(countBeforeAdd) = " "
                Else
                    ''Add the current character only.
                    ReDim piece(countBeforeAdd)
                    For j As Integer = 0 To countBeforeAdd - 1
                        piece(j) = tmpAry(j)
                    Next
                End If
                piece(piece.Count - 1) = c
            Next

            ''Build the output string from the piece() array.
            Dim output As String = ""
            Dim countAfterAdd As Integer = piece.Count
            For i As Integer = 0 To countAfterAdd - 1
                output += piece(i)
            Next
            Return output
        End If
    End Function
    Private Function IsUpperCase(ByVal s As Char) As Boolean
        If Equals(s, UCase(s)) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Class StringIndexed
        Implements IComparable(Of StringIndexed)
        Public Index As Integer
        Public Str As String
        Public Sub New(ByVal _str As String, ByVal _index As Integer)
            Str = _str
            Index = _index
        End Sub
        Public Overrides Function ToString() As String
            Return Str
        End Function
        Public Function CompareTo(ByVal other As StringIndexed) As Integer Implements IComparable(Of StringIndexed).CompareTo
            Return Str.CompareTo(other.Str)
        End Function
    End Class

    ''' <summary>
    ''' Returns the number of lines in the string. Only counts vbCrLf and vbNewLine characters as new line characters.
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LineCount(ByVal str As String) As Integer
        Dim crlf As Char = vbCrLf
        Dim numNewLineChars As Integer = str.Count(Function(c) (c = crlf))
        Return (1 + numNewLineChars)
    End Function
End Module
