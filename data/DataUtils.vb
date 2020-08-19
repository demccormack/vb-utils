Public Module DataUtils
    Public Sub DataViewFormShow(ByRef tbl As DataTable)
        Dim dvf As New DataViewForm(tbl)
        dvf.ShowDialog()
    End Sub

    Public Function SqlReferentialRule(ByVal index As Byte) As Rule
        Select Case index
            Case 0 : Return Rule.None
            Case 1 : Return Rule.Cascade
            Case 2 : Return Rule.SetNull
            Case 3 : Return Rule.SetDefault
            Case Else : Throw New Exception("SQL Server referential integrity rule not recognised.")
        End Select
    End Function

    ''' <summary>
    ''' Finds the data type referred to in a schema table by its code number.
    ''' </summary>
    ''' <param name="code"></param>
    ''' <returns></returns>
    ''' <remarks>Throws a NullReferenceException if the data type is not recognised.</remarks>
    Public Function DataTypeFromOleDbSchema(ByVal code As Byte) As Type
        Dim dataType As Type
        Dim exceptionMsg As String
        Select Case code
            Case 2 : dataType = GetType(Short)
            Case 3 : dataType = GetType(Long)
            Case 4 : dataType = GetType(Single)
            Case 5 : dataType = GetType(Double)
            Case 6 : exceptionMsg = "Data type 'Currency' not supported"
            Case 7 : dataType = GetType(DateTime)
            Case 11 : dataType = GetType(Boolean)
            Case 17 : dataType = GetType(Byte)
            Case 72 : exceptionMsg = "Data type 'GUID' not supported"
            Case 128 : exceptionMsg = "Data type 'BigBinary'/'LongBinary'/'VarBinary' not supported"
            Case 130 : dataType = GetType(String)           ''LongText/VarChar
            Case 131 : dataType = GetType(Decimal)
        End Select
        If (exceptionMsg IsNot Nothing) Then
            Throw New Exception(exceptionMsg)
        End If
        Return dataType
    End Function

    Public Function RuleFromOleDbSchema(ByVal schemaText As String) As Rule
        Select Case schemaText
            Case "NO ACTION" : Return Rule.None
            Case "CASCADE" : Return Rule.Cascade
            Case Else : Throw New Exception("Rule not recognised for text '" & schemaText & "'")
        End Select
    End Function

    ''' <summary>
    ''' Safely combines the functions IsDBNull and String.IsNullOrWhiteSpace. Intended for use with string fields in a DataTable.
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsDBNullOrWhiteSpace(ByVal obj As Object) As Boolean
        Return (IsDBNull(obj) OrElse String.IsNullOrWhiteSpace(obj.ToString))
    End Function
End Module