''' <summary>
''' A form containing a read-only DataGridView for the specified DataTable. Sub New shows the form. Generally used for testing during development.
''' </summary>
''' <remarks>Although the table is passed ByRef, this class uses a copy of the table to ensure it will not be changed.</remarks>
Public Class DataViewForm
    Inherits Form
    ''' <summary>
    ''' Creates and shows the form.
    ''' </summary>
    ''' <param name="tbl">The DataTable.</param>
    ''' <remarks></remarks>
    Public Sub New(ByRef tbl As DataTable)
        Using table As DataTable = tbl.Copy
            Me.Size = New Size(800, 600)
            Me.ShowIcon = False
            Me.Text = table.TableName
            Dim dgv As New DataGridView
            With dgv
                .DataSource = table
                .Dock = DockStyle.Fill
                .ReadOnly = True
                .AllowUserToDeleteRows = False
            End With
            Controls.Add(dgv)
            Me.Show()
        End Using
    End Sub
End Class
