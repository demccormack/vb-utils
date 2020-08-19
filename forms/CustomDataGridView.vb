Public Class CustomDataGridView
    Inherits DataGridView

    Private autoGenerateColumnsMirror As Boolean = True
    Public Shadows Property AutoGenerateColumns As Boolean
        Get
            Return autoGenerateColumnsMirror
        End Get
        Set(ByVal value As Boolean)
            autoGenerateColumnsMirror = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
        EditMode = DataGridViewEditMode.EditOnEnter
        MultiSelect = False
        MyBase.AutoGenerateColumns = False
        RowTemplate = New CustomDataGridViewRow
        Dim cms As New ContextMenuStrip
        cms.Items.Add("Show Entire DataTable", Nothing, AddressOf ShowDataViewForm)
        ContextMenuStrip = cms
    End Sub

    Private Sub ShowDataViewForm()
        If (DataSourceView IsNot Nothing) Then
            Using dt As DataTable = DataSourceView.Table.Copy
                For Each column As DataColumn In dt.Columns
                    column.ColumnMapping = MappingType.Element
                Next
                Dim f As New DataViewForm(dt)
                f.Visible = False
                f.ShowDialog()
            End Using
        End If
    End Sub

    Public Shadows Property DataSource As Object
        Get
            Return MyBase.DataSource
        End Get
        Set(ByVal value As Object)
            MyBase.DataSource = value
            If AutoGenerateColumns Then
                AutoGenerateColumnSet()
            End If
        End Set
    End Property

    Private ReadOnly Property DataSourceView As DataView
        Get
            If (DataSource Is Nothing) Then
                Return Nothing
            Else
                Select Case DataSource.GetType
                    Case GetType(DataTable) : Return CType(DataSource, DataTable).DefaultView
                    Case GetType(DataView) : Return CType(DataSource, DataView)
                    Case Else : Throw New Exception("Unable to convert the DataSource to a DataView.")
                End Select
            End If
        End Get
    End Property


    ''Needed when the value in the comboboxcell changes.
    Private Sub DirtyStateChanged() Handles Me.CurrentCellDirtyStateChanged
        If IsCurrentCellDirty Then
            CType(CurrentCell.OwningRow, CustomDataGridViewRow).RequiresPencil = True
        End If
    End Sub
    ''Needed to ensure pencils are erased whenever a row loses focus through any other means than a CalendarCell.
    Private Sub CurrentRowLeave(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles Me.RowLeave
        CType(Rows(e.RowIndex), CustomDataGridViewRow).RequiresPencil = False
    End Sub
    Friend DgvWasLeft As Boolean    ''required so that an infinite loop does not occur when the dgv is re-entered.
    Private Sub DgvLosingFocus(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Leave
        If Not Equals(CurrentRow.Index, NewRowIndex) AndAlso Not Equals(CurrentCell.Value, CType(EditingControl, IDataGridViewEditingControl).GetEditingControlFormattedValue(DataGridViewDataErrorContexts.LeaveControl)) Then
            CurrentCell.Value = CType(EditingControl, IDataGridViewEditingControl).GetEditingControlFormattedValue(DataGridViewDataErrorContexts.LeaveControl)
        End If
        ClearSelection()
        EndEdit()
        DgvWasLeft = True
        CType(CurrentCell.OwningRow, CustomDataGridViewRow).RequiresPencil = False
        InvalidateCell(CurrentCell.OwningRow.HeaderCell)
    End Sub
    Private Sub DgvMultiSelectChanged(ByVal sender As Object, ByVal e As EventArgs) Handles Me.MultiSelectChanged
        If MultiSelect Then
            CType(CurrentCell.OwningRow, CustomDataGridViewRow).RequiresPencil = False
        End If
    End Sub


    Private Shared bmpPencil As Bitmap
    Public ReadOnly Property EditingPencil As Bitmap
        Get
            If (bmpPencil Is Nothing) Then
                Dim bkColor As Color = RowHeadersDefaultCellStyle.BackColor
                bmpPencil = New Bitmap(16, 16)
                Dim g As Graphics = Graphics.FromImage(bmpPencil)
                g.Clear(bkColor)
                Dim pixels As Integer(,) = {{2, 12}, {4, 12}, {6, 12}, {6, 11}, {6, 10}, {6, 9}, {6, 8}, {7, 11}, {7, 10}, {7, 7}, {7, 6}, {8, 10}, {8, 5},
                                            {8, 4}, {9, 9}, {9, 8}, {9, 3}, {9, 2}, {10, 7}, {10, 6}, {10, 4}, {10, 2}, {11, 5}, {11, 4}, {11, 2}, {12, 3}}
                Pixelate(bmpPencil, pixels, Color.Red)
            End If
            Return bmpPencil
        End Get
    End Property

    Private Sub InsertAutoTypedColumn(ByVal col As DataColumn)
        Dim newCol As DataGridViewColumn = DataGridViewColumnFromType(col.DataType)
        Dim name As String = col.ColumnName
        With newCol
            .DataPropertyName = name
            .HeaderText = name
            .Name = name
        End With
        Columns.Add(newCol)
    End Sub

    Private Sub AutoGenerateColumnSet()
        For Each col As DataGridViewColumn In Columns
            Columns.Remove(col)
        Next
        For Each column As DataColumn In DataSourceView.Table.Columns
            InsertAutoTypedColumn(column)
        Next
    End Sub

    Public Overloads Sub InsertColumnByType(ByVal colName As String)
        Dim column As DataColumn = DataSourceView.Table.Columns(colName)
        InsertAutoTypedColumn(column)
    End Sub
    Public Overloads Sub InsertColumnByType(ByVal colIndex As Integer)
        Dim column As DataColumn = DataSourceView.Table.Columns(colIndex)
        InsertAutoTypedColumn(column)
    End Sub

    Public Overloads Sub InsertComboBoxColumn(ByVal colName As String, ByVal items() As String)
        AddNewComboBoxColumn(DataSourceView.Table.Columns(colName), items)
    End Sub
    Public Overloads Sub InsertComboBoxColumn(ByVal colIndex As Integer, ByVal items() As String)
        AddNewComboBoxColumn(DataSourceView.Table.Columns(colIndex), items)
    End Sub
    Private Sub AddNewComboBoxColumn(ByVal dataCol As DataColumn, ByVal items() As String)
        Dim cbc As New DataGridViewComboBoxColumn
        Dim name As String = dataCol.ColumnName
        With cbc
            .DataPropertyName = name
            .HeaderText = name
            .Name = name
            .Items.AddRange(items)
        End With
        Columns.Add(cbc)
    End Sub

    Private Sub RowHeaders_MouseClick(ByVal sender As DataGridView, ByVal e As EventArgs) Handles Me.SelectionChanged
        If (SelectedRows.Count <> 0) Then
            EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
            MultiSelect = True
            EndEdit()
        Else
            EditMode = DataGridViewEditMode.EditOnEnter
            MultiSelect = False
        End If
    End Sub
End Class