Imports System.Data.OleDb
Imports System.Data.SqlClient


''' <summary>
''' Manages communication with a MS Access or SQL Server database.
''' </summary>
''' <remarks></remarks>
Public MustInherit Class DbComm
    Private ds As DataSet
    Private dsRowChangeHandler As DataRowChangeEventHandler = AddressOf WriteBackHandler
    Private blnAutoUpdate As Boolean
    Private FillOrder() As String

#Region "Properties"
    Public ReadOnly Property DataSet As DataSet
        Get
            Return ds
        End Get
    End Property
    Public ReadOnly Property TableNames As String()
        Get
            Return FillOrder
        End Get
    End Property
    Public Event AutoUpdateChanged()
    Public Property AutoUpdate As Boolean
        Get
            Return blnAutoUpdate
        End Get
        Set(ByVal value As Boolean)
            If value = False And blnAutoUpdate = True Then
                For Each table As String In FillOrder
                    RemoveEventHandler(table)
                Next
                blnAutoUpdate = False
                RaiseEvent AutoUpdateChanged()
            ElseIf value = True And blnAutoUpdate = False Then
                blnAutoUpdate = True
                For Each table As String In FillOrder
                    AddEventHandler(table)
                Next
                RaiseEvent AutoUpdateChanged()
            End If
        End Set
    End Property
    Public ReadOnly Property Relationships As DataTable
        Get
            Return ds.Tables("Relationships")
        End Get
    End Property
    Public ReadOnly Property PrimaryKeyName(ByVal tablename As String) As String
        Get
            With ds.Tables(tablename)
                If .PrimaryKey.Count > 1 Then
                    Throw New Exception("The primary key of DataTable '" & .TableName &
                                        "' contains more than one DataColumn." & vbCrLf & "The program does not know how to respond to this.")
                Else
                    Return .PrimaryKey(0).ColumnName
                End If
            End With
        End Get
    End Property
#End Region

#Region "Constructor"
    Public Shared Function NewDbComm(ByVal oldConnectionStr As String, Optional ByVal automaticallyUpdate As Boolean = True) As DbComm
        ''Ask the user to verify the connection string.
        Dim strDbTyp As String
        Select Case DetectDbType(oldConnectionStr)
            Case DbType.MsAccess : strDbTyp = "Microsoft Access"
            Case DbType.SqlServer : strDbTyp = "SQL Server"
            Case Else : Throw New Exception("The database type was not recognised.")
        End Select
        Dim newConnectionStr As String = oldConnectionStr
        If MsgBox("Database type: " & strDbTyp & vbCrLf & vbCrLf & "Connection string:" & vbCrLf & oldConnectionStr &
                  vbCrLf & vbCrLf & "Is this OK?", MsgBoxStyle.YesNo + vbInformation, "Database Connection") = DialogResult.No AndAlso
                    MsgBox("You should only change the connection string if the application is unable to connect to the database.",
                     MsgBoxStyle.OkCancel + vbExclamation, "Warning") = DialogResult.OK Then
            Dim userDefinedConStr As String = InputBox("Enter new connection string:", "Database Connection", oldConnectionStr)
            If userDefinedConStr = "" Then
                newConnectionStr = oldConnectionStr
            Else
                newConnectionStr = userDefinedConStr
            End If
        End If

        ''Create the connection, save the connection string and return the connection object.
        Dim newDbc As DbComm
        Select Case DetectDbType(newConnectionStr)
            Case DbType.MsAccess : newDbc = New DbCommAccess(newConnectionStr, automaticallyUpdate)
            Case DbType.SqlServer : newDbc = New DbCommSql(newConnectionStr, automaticallyUpdate)
        End Select
        AppControl.ConnectionString = newConnectionStr
        Return newDbc
    End Function
    Private Enum DbType
        SqlServer
        MsAccess
        Unknown
    End Enum
    Private Shared Function DetectDbType(ByVal connectionString As String) As DbType
        If connectionString.Contains("INITIAL CATALOG") Then
            Return DbType.SqlServer
        ElseIf connectionString.Contains("PROVIDER") Then
            Return DbType.MsAccess
        Else
            Return DbType.Unknown
        End If
    End Function
#End Region

#Region "Event Handlers"
    Private Sub AddEventHandler(ByVal tblName As String)
        If blnAutoUpdate Then
            AddHandler ds.Tables(tblName).RowChanged, dsRowChangeHandler
            AddHandler ds.Tables(tblName).RowDeleted, dsRowChangeHandler
        End If
    End Sub
    Private Sub RemoveEventHandler(ByVal tblName As String)
        If blnAutoUpdate Then
            RemoveHandler ds.Tables(tblName).RowChanged, dsRowChangeHandler
            RemoveHandler ds.Tables(tblName).RowDeleted, dsRowChangeHandler
        End If
    End Sub
#End Region

#Region "Updating Database"
    Public Sub UpdateDb()
        For i As Integer = FillOrder.Count - 1 To 0 Step -1
            WriteBack(FillOrder(i))
        Next
    End Sub
    Public Sub WriteBackHandler(ByVal sender As DataTable, ByVal e As EventArgs)   ''Handles DataTable.RowChanged and DataTable.RowDeleted through the dsRowChangeHandler delegate.
        Try
            WriteBack(sender.TableName)
        Catch ex As Exception When GlobalExceptionHandler.IsInUse
            GlobalExceptionHandler.ExceptionHandler(ex)
        End Try
    End Sub
    Public MustOverride Sub WriteBack(ByVal tablename As String)

    ''' <summary>
''' Clears the tables of the DataSet before refilling them.
''' </summary>
''' <remarks></remarks>
    Private Sub ReloadData()
        For i As Integer = FillOrder.Count - 1 To 0 Step -1
            RemoveEventHandler(FillOrder(i))
            ds.Tables(FillOrder(i)).Clear()
        Next
        For i As Integer = 0 To FillOrder.Count - 1
            FillTable(FillOrder(i))
            AddEventHandler(FillOrder(i))
        Next
    End Sub
    Public MustOverride Sub FillTable(ByVal tablename As String)
#End Region

    Public Function ChildRows(ByVal parentRow As DataRow, ByVal childTable As String) As DataRow()
        Dim relDv As DataView = ds.Tables("Relationships").DefaultView
        relDv.RowFilter = "Table = '" & childTable & "' AND RefTable = '" & parentRow.Table.TableName & "'"
        If (relDv.Count <> 1) Then
            Throw New DataException("There are " & relDv.Count.ToString & " relationships linking tables '" & childTable & "' and '" & parentRow.Table.TableName & ".")
        End If
        Dim pkName As String = relDv.Item(0).Item("RefColumn")
        Dim pkVal As String = parentRow.Item(pkName)
        Dim childDv As DataView = ds.Tables(childTable).DefaultView
        childDv.RowFilter = relDv.Item(0).Item("Column") & " = '" & pkVal & "'"
        Dim rowArray() As DataRow
        For Each item As DataRowView In childDv
            ArrayAddElement(rowArray, item.Row)
        Next
        Return rowArray
    End Function


Public Class DbCommAccess
    Inherits DbComm
    Private connection As New OleDbConnection
    Private cmd As New OleDbCommand(Nothing, connection)
    Private da As New OleDbDataAdapter(cmd)

    Public Sub New(ByVal connectionStr As String, Optional ByVal automaticallyUpdate As Boolean = True)
            connection.ConnectionString = connectionStr
            ds = DbGetAccess(connectionStr, FillOrder)
            AutoUpdate = automaticallyUpdate     ''if false, the event handlers will not be added.
    End Sub

    Public Overrides Sub WriteBack(ByVal tablename As String)
        cmd.CommandText = "SELECT * FROM " & tablename
        RemoveEventHandler(tablename)       ''Prevents this Sub being called again as the database is updated.
        Dim cb As New OleDbCommandBuilder(da)
            Try
                da.Update(ds, tablename)
                AddEventHandler(tablename)
            Catch oleex As OleDbException
                AddEventHandler(tablename)
                MsgBox("Error message from database: " & vbCrLf & oleex.Message, vbExclamation, "Error")
                ReloadData()
            End Try
    End Sub

    Public Overrides Sub FillTable(ByVal tablename As String)
        cmd.CommandText = "SELECT * FROM " & tablename
        da.Fill(ds, tablename)
    End Sub
End Class


Public Class DbCommSql
    Inherits DbComm
    Private connection As New SqlConnection
    Private cmd As New SqlCommand(Nothing, connection)
    Private da As New SqlDataAdapter(cmd)

    Public Sub New(ByVal connectionStr As String, Optional ByVal automaticallyUpdate As Boolean = True)
        connection.ConnectionString = connectionStr
        ds = DbGetSql(connectionStr, FillOrder)
        AutoUpdate = automaticallyUpdate     ''if false, the event handlers will not be added.
    End Sub

    Public Overrides Sub WriteBack(ByVal tablename As String)
        cmd.CommandText = "SELECT * FROM " & tablename
        RemoveEventHandler(tablename)       ''Prevents this Sub being called again as the database is updated.
        Dim cb As New SqlCommandBuilder(da)
            Try
                da.Update(ds, tablename)
                AddEventHandler(tablename)
            Catch ex As Exception When (ex.GetType = GetType(SqlException)) Or (ex.GetType = GetType(ArgumentException))
                AddEventHandler(tablename)
                MsgBox("Error message from database: " & vbCrLf & ex.Message, vbExclamation, "Error")
                ReloadData()
            End Try
    End Sub

    Public Overrides Sub FillTable(ByVal tablename As String)
        cmd.CommandText = "SELECT * FROM " & tablename
        da.Fill(ds, tablename)
    End Sub
End Class
End Class