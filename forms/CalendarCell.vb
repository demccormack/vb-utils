Public Class CalendarCell
    Inherits DataGridViewTextBoxCell

    Public ctl As CalendarEditingControl
    Public Sub New()
        Style.Format = "d"
    End Sub

    Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, ByVal initialFormattedValue As Object, ByVal dataGridViewCellStyle As System.Windows.Forms.DataGridViewCellStyle)
        MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle)
        ctl = CType(DataGridView.EditingControl, CalendarEditingControl)
        ctl.OwnerCell = Me
    End Sub

    Public Overrides ReadOnly Property EditType() As Type
        Get
            ' Return the type of the editing control that CalendarCell uses.
            Return GetType(CalendarEditingControl)
        End Get
    End Property

    Public Overrides ReadOnly Property ValueType() As Type
        Get
            Return GetType(DateTime)
        End Get
    End Property

    Public Overrides ReadOnly Property DefaultNewRowValue() As Object
        Get
            Return DBNull.Value
        End Get
    End Property
End Class