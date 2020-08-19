Public Class CalendarColumn
    Inherits DataGridViewColumn
    Public Sub New()
        MyBase.New(New CalendarCell)
        Me.SortMode = DataGridViewColumnSortMode.Automatic
    End Sub

    Public Overrides Property CellTemplate As System.Windows.Forms.DataGridViewCell
        Get
            Return MyBase.CellTemplate
        End Get
        Set(ByVal value As System.Windows.Forms.DataGridViewCell)
            If Not value.GetType.IsAssignableFrom(GetType(CalendarCell)) Then
                Throw New InvalidCastException("The CellTemplate must be a CalendarCell.")
            End If
            MyBase.CellTemplate = value
        End Set
    End Property
End Class