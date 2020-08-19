Module ArrayUtils
    ''' <summary>
    ''' Adds the element to the array.
    ''' </summary>
    ''' <param name="ary">The array.</param>
    ''' <param name="value">The object to add.</param>
    ''' <remarks>If T is numeric, the new element will be appended rather than replacing any zeroes.
    ''' Otherwise, the first element which is nothing will be replaced by the new element.</remarks>
    Public Sub ArrayAddElement(Of T)(ByRef ary As T(), ByRef value As T)
        ''Find the current length of the array.
        Dim length As Integer
        Try
            length = ary.Length
        Catch ex As NullReferenceException
            length = 0
        End Try
        ''Find the required position of the new element.
        Dim newIndex As Integer = length
        If length > 0 Then
            For i As Integer = 0 To length - 1
                If ary(i) Is Nothing Then
                    newIndex = i
                    Exit For
                End If
            Next
        End If
        ''Add the new element at the required position.
        If newIndex >= length Then
            ReDim Preserve ary(length)
        End If
        ary(newIndex) = value
    End Sub

    ''' <summary>
    ''' Inserts the element into the array at the specified postion.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="ary">The array.</param>
    ''' <param name="value">The object to add.</param>
    ''' <param name="atPosition">The zero-based index of the new element.</param>
    ''' <remarks>Elements already in or after this position will be shuffled one position further.
    ''' The array size is increased by one regardless of any empty places which may be present.</remarks>
    Public Sub ArrayAddElement(Of T)(ByRef ary As T(), ByRef value As T, ByVal atPosition As Integer)
        ''Determine the array's initial length
        Dim length As Integer
        Try
            length = ary.Length
        Catch ex As NullReferenceException
            length = 0
        End Try

        ''Determine the new object's position
        Dim isFirst As Boolean = False
        Dim isLast As Boolean = False
        If (atPosition > length) Then
            Throw New IndexOutOfRangeException("Parameter 'atPosition' is larger than the size of the array.")
        ElseIf (atPosition = length) Then
            isLast = True
        ElseIf (atPosition = 0) Then
            isFirst = True
        End If

        ''Insert the new object
        Dim tmpAry(length) As T
        If (Not isFirst) Then
            For i As Integer = 0 To (atPosition - 1)
                tmpAry(i) = ary(i)
            Next
        End If
        tmpAry(atPosition) = value
        If (Not isLast) Then
            For i As Integer = (atPosition + 1) To (tmpAry.Length - 1)
                tmpAry(i) = ary(i - 1)
            Next
        End If
        ary = tmpAry
    End Sub

    Public Sub ArrayAddElement(Of T)(ByRef ary As T(,), ByVal dimension As Integer, ByRef value As T)
        Throw New NotSupportedException("Sub ArrayAddElement does not currently support arrays of more than one dimension.")
        ''To be implemented
    End Sub
End Module
