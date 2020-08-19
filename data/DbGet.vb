Imports System.Data.SqlClient
Imports System.Data.OleDb

Module DbGet

    ''' <summary>
    ''' Returns a DataSet built from an SQL Server database.
    ''' </summary>
    ''' <param name="connectionStr">The connection string for the SqlConnection.</param>
    ''' <param name="FillOrderArray">The string array to be used for storing the fill order of the tables.</param>
    ''' <returns></returns>
    ''' <remarks>Can only handle databases with two or less relationship levels.</remarks>
    Public Function DbGetSql(ByVal connectionStr As String, ByRef FillOrderArray As String()) As DataSet
        Dim connection As New SqlConnection(connectionStr)
        Dim cmd As New SqlCommand(Nothing, connection)
        Dim da As New SqlDataAdapter(cmd)
        Dim output As New DataSet()

        ''Get the details of the database's tables and foreign keys.
        cmd.CommandText = "SELECT name FROM sys.sysobjects WHERE type = 'U'"
        da.Fill(output, "t")
        cmd.CommandText = "SELECT f.name AS ForeignKey, OBJECT_NAME(f.parent_object_id) AS TableName, COL_NAME(fc.parent_object_id, fc.parent_column_id) AS ColumnName, " &
            "OBJECT_NAME (f.referenced_object_id) AS RefTableName, COL_NAME(fc.referenced_object_id, fc.referenced_column_id) AS RefColumnName, " &
            "f.delete_referential_action AS DeleteAction, f.delete_referential_action_desc AS DeleteActionDesc, f.update_referential_action AS UpdateAction, " &
            "f.update_referential_action_desc AS UpdateActionDesc " &
            "FROM sys.foreign_keys AS f JOIN sys.foreign_key_columns AS fc ON f.OBJECT_ID = fc.constraint_object_id"
        da.Fill(output, "f")
        cmd.CommandText = "SELECT tbl.name as table_name, i.name AS pk_name, COL_NAME(ic.object_id,ic.column_id) AS column_name, ic.column_id AS column_index, col.is_nullable " &
            "FROM sys.indexes AS i JOIN sys.index_columns AS ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id " &
            "JOIN sys.tables AS tbl ON tbl.object_id = i.object_id " &
            "JOIN sys.columns AS col ON col.object_id = ic.object_id AND col.column_id = ic.column_id"
        da.Fill(output, "p")

        ''Build the FillOrder() array, which determines the order in which tables must be filled.
        Dim RefObj(), UnRefObj() As String
        For Each table As DataRow In output.Tables("t").Rows
            Dim IsRefObj As Boolean = False
            For Each fk As DataRow In output.Tables("f").Rows
                If fk("RefTableName") = table("name") Then
                    IsRefObj = True
                    Exit For
                End If
            Next
            If IsRefObj Then
                ArrayAddElement(RefObj, table("name"))
            Else
                ArrayAddElement(UnRefObj, table("name"))
            End If
        Next
        For i As Integer = 0 To RefObj.Count - 1
            ArrayAddElement(FillOrderArray, RefObj(i))
        Next
        For i As Integer = 0 To UnRefObj.Count - 1
            ArrayAddElement(FillOrderArray, UnRefObj(i))
        Next

        ''Fill the tables in the correct order.
        For i As Integer = 0 To FillOrderArray.Count - 1
            cmd.CommandText = "SELECT * FROM " & FillOrderArray(i)
            da.Fill(output, FillOrderArray(i))
        Next

        ''Create primary key constraints
        For Each pk As DataRow In output.Tables("p").Rows
            Dim pkcol() As DataColumn = output.Tables(pk("table_name")).PrimaryKey
            Dim newpk As DataColumn = output.Tables(pk("table_name")).Columns(pk("column_name"))
            With newpk
                .AllowDBNull = CBool(pk("is_nullable"))
                If .DataType.ToString.Remove(0, 7) <> "String" Then
                    .AutoIncrement = True
                    .AutoIncrementStep = 1
                    Dim pkDv As DataView = .Table.DefaultView
                    Dim pkColName As String = .ColumnName
                    pkDv.Sort = pkColName
                    Try
                        .AutoIncrementSeed = CLng(pkDv.ToTable.Rows(.Table.Rows.Count - 1).Item(pkColName) + 1)
                    Catch ex As InvalidOperationException
                        .AutoIncrementSeed = 0
                    End Try
                    If Not DebuggerAttached() Then
                        .ColumnMapping = MappingType.Hidden
                    End If
                End If
            End With
            ArrayAddElement(pkcol, newpk)
            output.Tables(pk("table_name")).PrimaryKey = pkcol
        Next

        ''Create foreign key constraints
        For Each fk As DataRow In output.Tables("f").Rows
            ''Define variables from which to create the constraint.
            Dim fkcName As String = fk("ForeignKey")
            Dim refTblName As String = fk("RefTableName")
            Dim refColName As String = fk("RefColumnName")
            Dim tblName As String = fk("TableName")
            Dim colName As String = fk("ColumnName")
            Dim DelIndex As Byte = fk("DeleteAction")
            Dim UpdIndex As Byte = fk("UpdateAction")
            ''Create the constraint.
            Dim pdc, fdc As DataColumn
            pdc = output.Tables(refTblName).Columns(refColName)
            fdc = output.Tables(tblName).Columns(colName)
            fdc.AllowDBNull = pdc.AllowDBNull
            Dim fkc As New ForeignKeyConstraint(fkcName, pdc, fdc)
            fkc.DeleteRule = SqlReferentialRule(DelIndex)
            fkc.UpdateRule = SqlReferentialRule(UpdIndex)
            ''Add the constraint.
            output.Tables(tblName).Constraints.Add(fkc)
        Next

        ''Create the relationships table, which is accessed through the Relationships property of IDbComm.
        Dim RelationTable As New DataTable
        RelationTable = output.Tables("f").Copy
        With RelationTable
            .Columns.Remove("DeleteAction")
            .Columns.Remove("DeleteActionDesc")
            .Columns.Remove("UpdateAction")
            .Columns.Remove("UpdateActionDesc")
            .Columns("TableName").ColumnName = "Table"
            .Columns("ColumnName").ColumnName = "Column"
            .Columns("RefTableName").ColumnName = "RefTable"
            .Columns("RefColumnName").ColumnName = "RefColumn"
            .Columns("ForeignKey").ColumnName = "Name"
            .TableName = "Relationships"
        End With
        output.Tables.Add(RelationTable)

        Return output
    End Function




    ''' <summary>
    ''' Returns a DataSet built from a MS Access database.
    ''' </summary>
    ''' <param name="connectionStr">The connection string for the OleDbConnection.</param>
    ''' <param name="FillOrderArray">The (empty) string array to be used for storing the fill order of the tables.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DbGetAccess(ByVal connectionStr As String, ByRef FillOrderArray As String()) As DataSet
        ''Create connection objects
        Dim connection As New OleDbConnection(connectionStr)
        Dim cmd As New OleDbCommand(Nothing, connection)
        Dim da As New OleDbDataAdapter(cmd)
        Dim dsSchema, output As New DataSet

        ''Get table schema
        With connection
            .Open()
            Dim dtTables As DataTable = .GetOleDbSchemaTable(OleDbSchemaGuid.Tables, {Nothing, Nothing, Nothing, "TABLE"})
            Dim dtColumns As DataTable = .GetOleDbSchemaTable(OleDbSchemaGuid.Columns, Nothing)
            Dim dtPrimaryKeys As DataTable = .GetOleDbSchemaTable(OleDbSchemaGuid.Primary_Keys, Nothing)
            Dim dtForeignKeys As DataTable = .GetOleDbSchemaTable(OleDbSchemaGuid.Foreign_Keys, Nothing)
            Dim dtUniqueConstraints As DataTable = .GetOleDbSchemaTable(OleDbSchemaGuid.Table_Constraints, {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, "UNIQUE"})
            dtUniqueConstraints.TableName = "Unique_Constraints"
            dsSchema.Tables.AddRange({dtTables, dtColumns, dtPrimaryKeys, dtForeignKeys, dtUniqueConstraints})
            .Close()
        End With
        Dim tables() As String

        ''Build the tables and fill them with columns
        Dim dvColumns As DataView = dsSchema.Tables("Columns").AsDataView
        dvColumns.Sort = "ORDINAL_POSITION"
        Dim filter As String
        For Each row As DataRow In dsSchema.Tables("Tables").Rows
            Dim tableName As String = row.Item("TABLE_NAME")
            ArrayAddElement(tables, tableName)
            output.Tables.Add(tableName)
            filter += ("TABLE_NAME = '" & row.Item("TABLE_NAME") & "' OR ")
        Next
        filter = filter.Remove(filter.LastIndexOf(" OR "))
        dvColumns.RowFilter = filter
        For Each row As DataRowView In dvColumns
            ''Get column details
            Dim tblname As String = row.Item("TABLE_NAME")
            Dim colname As String = row.Item("COLUMN_NAME")
            Dim dataTypeCode As Byte = row.Item("DATA_TYPE")
            Dim allowNull As Boolean = CBool(row.Item("IS_NULLABLE"))
            Dim defaultValue As Object = Nothing
            If CBool(row.Item("COLUMN_HASDEFAULT")) Then
                defaultValue = row.Item("COLUMN_DEFAULT")
            End If
            Dim ordinal As Integer = (row.Item("ORDINAL_POSITION") - 1)
            ''Create column
            Dim dc As New DataColumn(colname)
            dc.DataType = DataTypeFromOleDbSchema(dataTypeCode)
            dc.AllowDBNull = allowNull
            dc.DefaultValue = defaultValue
            output.Tables(tblname).Columns.Add(dc)
            dc.SetOrdinal(ordinal)
        Next

        ''Set primary keys
        For Each row As DataRow In dsSchema.Tables("Primary_Keys").Rows
            Dim tblName As String = row.Item("TABLE_NAME")
            Dim colName As String = row.Item("COLUMN_NAME")
            With output.Tables(tblName)
                ArrayAddElement(.PrimaryKey, .Columns(colName))
                .Columns(colName).AutoIncrement = True
                .Columns(colName).ColumnMapping = MappingType.Hidden
            End With
        Next

        ''Set foreign keys
        For Each row As DataRow In dsSchema.Tables("Foreign_Keys").Rows
            Dim pkTable As String = row.Item("PK_TABLE_NAME")
            Dim pkColumn As String = row.Item("PK_COLUMN_NAME")
            Dim fkTable As String = row.Item("FK_TABLE_NAME")
            Dim fkColumn As String = row.Item("FK_COLUMN_NAME")
            Dim updateRule As String = row.Item("UPDATE_RULE")
            Dim deleteRule As String = row.Item("DELETE_RULE")
            Dim name As String = row.Item("FK_NAME")
            Dim fkc As New ForeignKeyConstraint(name, output.Tables(pkTable).Columns(pkColumn), output.Tables(fkTable).Columns(fkColumn))
            fkc.UpdateRule = RuleFromOleDbSchema(updateRule)
            fkc.DeleteRule = RuleFromOleDbSchema(deleteRule)
            output.Tables(fkTable).Constraints.Add(fkc)
        Next

        ''Set any additional unique constraints. ***Can't set check constraints.***
        For Each row As DataRow In dsSchema.Tables("Unique_Constraints").Rows
            Dim tableName As String = row.Item("TABLE_NAME")
            Dim columnName As String = row.Item("CONSTRAINT_NAME")
            output.Tables(tableName).Columns(columnName).Unique = True
        Next

        ''Determine the order in which to fill the tables
        Using dtTableRelations As New DataTable
            dtTableRelations.Columns.Add(New DataColumn("Parent", GetType(String)))
            dtTableRelations.Columns.Add(New DataColumn("Child", GetType(String)))
            dtTableRelations.Columns.Add(New DataColumn("Done", GetType(Boolean)))
            For Each row As DataRow In dsSchema.Tables("Foreign_Keys").Rows
                dtTableRelations.Rows.Add({row.Item("PK_TABLE_NAME"), row.Item("FK_TABLE_NAME"), False})
            Next
            With dtTableRelations.Rows(0)
                Dim _parent As String = .Item("Parent")
                Dim _child As String = .Item("Child")
                ArrayAddElement(FillOrderArray, _parent)
                ArrayAddElement(FillOrderArray, _child, (Array.IndexOf(FillOrderArray, _parent) + 1))
                .Delete()
            End With
            For i As Integer = 0 To 0           ''this becomes a loop if the value of i is reduced by 1 later on.
                For Each currentItem As String In FillOrderArray
                    For Each row As DataRow In dtTableRelations.Rows
                        If Not row.Item("Done") Then
                            Dim indexOfCurrent As Integer = Array.IndexOf(FillOrderArray, currentItem)
                            Dim _parent As String = row.Item("Parent")
                            Dim _child As String = row.Item("Child")
                            If (_parent = currentItem) Then
                                If (Not FillOrderArray.Contains(_child)) Then
                                    ArrayAddElement(FillOrderArray, _child, indexOfCurrent + 1)
                                End If
                                row.Item("Done") = True
                            ElseIf (_child = currentItem) Then
                                If (Not FillOrderArray.Contains(_parent)) Then
                                    ArrayAddElement(FillOrderArray, _parent, indexOfCurrent)
                                End If
                                row.Item("Done") = True
                            End If
                        End If
                    Next
                Next
                For j As Integer = 0 To (dtTableRelations.Rows.Count - 1)
                    If (j > dtTableRelations.Rows.Count - 1) Then
                        Exit For
                    ElseIf dtTableRelations.Rows(j).Item("Done") Then
                        dtTableRelations.Rows(j).Delete()
                        j -= 1
                    End If
                Next
                If (dtTableRelations.Rows.Count > 0) Then
                    i -= 1
                End If
            Next
        End Using

        'Dim strFillOrder As String
        'For i As Integer = 0 To (FillOrderArray.Count - 1)
        '    strFillOrder += FillOrderArray(i) & " "
        'Next
        'MsgBox(strFillOrder)

        ''Fill the tables
        For Each table As String In FillOrderArray
            cmd.CommandText = "SELECT * FROM " & table
            da.Fill(output, table)
        Next

        ''Create the Relationships table
        Dim RelationTable As New DataTable
        With RelationTable
            .TableName = "Relationships"
            .Columns.Add(New DataColumn("Table", GetType(String)))
            .Columns.Add(New DataColumn("Column", GetType(String)))
            .Columns.Add(New DataColumn("RefTable", GetType(String)))
            .Columns.Add(New DataColumn("RefColumn", GetType(String)))
            .Columns.Add(New DataColumn("Name", GetType(String)))
            For Each row As DataRow In dsSchema.Tables("Foreign_Keys").Rows
                RelationTable.Rows.Add({row.Item("FK_TABLE_NAME"), row.Item("FK_COLUMN_NAME"), row.Item("PK_TABLE_NAME"), row.Item("PK_COLUMN_NAME"), row.Item("FK_NAME")})
            Next
            output.Tables.Add(RelationTable)
        End With

        DataViewFormShow(RelationTable)
        Return output
    End Function
End Module
