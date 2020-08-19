Public Class CalendarEditingControl
    Inherits DatePickerNullable
    Implements IDataGridViewEditingControl

    Private WithEvents dataGridViewControl As DataGridView
    Private valueIsChanged As Boolean = False
    Private rowIndexNum As Integer
    Private initialCellValue As Object
    Private associatedCell As CalendarCell
    Private formClosing As Boolean
    Private WithEvents containingForm As Form
    Private bmpPencil As Bitmap
    Private changingOwnerCell As Boolean

#Region "Cell Switching"
    ''' <summary>The CalendarCell which 'owns' the control.</summary>
    ''' <remarks>Checks that this is a new assignment and takes no action if not.</remarks>
    Public Property OwnerCell As CalendarCell
        Get
            Return associatedCell
        End Get
        Set(ByVal val As CalendarCell)
            If Not ReferenceEquals(associatedCell, val) Then
                If (associatedCell IsNot Nothing) Then
                    If CType(dataGridViewControl, CustomDataGridView).DgvWasLeft Then
                        CType(dataGridViewControl, CustomDataGridView).DgvWasLeft = False
                    Else
                        UpdateDataGridView()    ''for the cell which is losing focus.
                    End If
                    If (val IsNot Nothing) AndAlso Not ReferenceEquals(associatedCell.OwningRow, val.OwningRow) Then
                        CType(associatedCell.OwningRow, CustomDataGridViewRow).RequiresPencil = False
                    End If
                End If
                changingOwnerCell = True
                If (val IsNot Nothing) AndAlso (val.RowIndex = -1) Then
                    associatedCell = dataGridViewControl.CurrentCell
                Else
                    associatedCell = val
                End If

                If associatedCell Is Nothing Then   ''...the CalendarEditingControl is disposing.
                    changingOwnerCell = False
                    initialCellValue = TextToValue()
                    InvalidateHeaderCell()
                Else
                    initialCellValue = associatedCell.Value
                    Value = Now.Date        ''Required in order to make the default date today if the control is going to be blank.
                    containingForm = EditingControlDataGridView.FindForm
                    With associatedCell
                        If (.Value Is Nothing) Or IsDBNull(.Value) Then
                            Value = .DefaultNewRowValue
                        Else
                            Value = CType(.Value, DateTime)
                        End If
                    End With
                    changingOwnerCell = False
                End If
            End If
        End Set
    End Property

    Private Sub UpdateDataGridView()
        If Not Equals(initialCellValue, Value) Then
            EditingControlDataGridView.NotifyCurrentCellDirty(True)
            OwnerCell.Value = Value
            EditingControlDataGridView.UpdateCellValue(OwnerCell.ColumnIndex, OwnerCell.RowIndex)
            initialCellValue = Nothing
        End If
    End Sub

    Private Sub EditModeChanged() Handles dataGridViewControl.EditModeChanged
        If dataGridViewControl.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2 Then
            UpdateDataGridView()
        End If
    End Sub
#End Region

#Region "Pencil Glyph"
    Private Sub InvalidateHeaderCell()
        If (dataGridViewControl.CurrentCell IsNot Nothing) Then
            dataGridViewControl.InvalidateCell(dataGridViewControl.CurrentCell.OwningRow.HeaderCell)
        End If
    End Sub
#End Region

#Region "Functionality"
    Private Sub ConditionallyReDraw() Handles Me.TextChanged
        If Not changingOwnerCell Then
            CType(OwnerCell.OwningRow, CustomDataGridViewRow).RequiresPencil = True
            InvalidateHeaderCell()
        End If
    End Sub

    Private ReadOnly Property DisplayIsAltered As Boolean
        Get
            Return Not Equals(TextToValue, initialCellValue)
        End Get
    End Property

    Private Sub EscapeKey(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Value = initialCellValue
        End If
    End Sub

    Protected Overrides Sub OnCloseUp(ByVal eventargs As System.EventArgs)
        MyBase.OnCloseUp(eventargs)
        If DisplayIsAltered Then
            InvalidateHeaderCell()
        End If
    End Sub
#End Region

#Region "Clean-up"
    Private Sub ContainingFormClosing() Handles containingForm.FormClosing
        ''Required to stop half-finished changes being written back to the DataTable.
        formClosing = True
    End Sub

    Private Sub ResetInitialCellValue() Handles dataGridViewControl.DataMemberChanged
        ''Required because the control is not disposed on changing the DataMember.
        initialCellValue = Value
    End Sub

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        ''Decide whether to write the changes back to the DataTable.
        If formClosing Then
            OwnerCell.Value = initialCellValue
        Else
            OwnerCell = Nothing
            InvalidateHeaderCell()
        End If
        ''Me.Focus()
        ''SendKeys.Send("{ESC}")
        formClosing = False
        MyBase.Dispose(disposing)
    End Sub
#End Region

#Region "Uninteresting code"
    Protected Overrides Sub OnValueChanged(ByVal eventargs As System.EventArgs)
        valueIsChanged = True
        MyBase.OnValueChanged(eventargs)
    End Sub

    Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As System.Windows.Forms.DataGridViewCellStyle) Implements System.Windows.Forms.IDataGridViewEditingControl.ApplyCellStyleToEditingControl
        Me.Font = dataGridViewCellStyle.Font
        Me.CalendarForeColor = dataGridViewCellStyle.ForeColor
        Me.CalendarMonthBackground = dataGridViewCellStyle.BackColor
    End Sub

    Public Property EditingControlDataGridView As System.Windows.Forms.DataGridView Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlDataGridView
        Get
            Return dataGridViewControl
        End Get
        Set(ByVal value As System.Windows.Forms.DataGridView)
            dataGridViewControl = value
        End Set
    End Property

    Public Property EditingControlFormattedValue As Object Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlFormattedValue
        Get
            Return ToShortDateString()
        End Get
        Set(ByVal val As Object)
            Try
                Value = DateTime.Parse(CStr(val))
            Catch ex As Exception
                Value = OwnerCell.DefaultNewRowValue
            End Try
        End Set
    End Property

    Public Property EditingControlRowIndex As Integer Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlRowIndex
        Get
            Return rowIndexNum
        End Get
        Set(ByVal value As Integer)
            rowIndexNum = value
        End Set
    End Property

    Public Property EditingControlValueChanged As Boolean Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlValueChanged
        Get
            Return valueIsChanged
        End Get
        Set(ByVal value As Boolean)
            valueIsChanged = value
        End Set
    End Property

    Public Function EditingControlWantsInputKey(ByVal keyData As System.Windows.Forms.Keys, ByVal dataGridViewWantsInputKey As Boolean) As Boolean Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlWantsInputKey
        Select Case keyData And Keys.KeyCode
            Case Keys.Left, Keys.Up, Keys.Down, Keys.Right, Keys.Home, Keys.End, Keys.PageDown, Keys.PageUp
                Return True
            Case Else
                Return Not dataGridViewWantsInputKey
        End Select
    End Function

    Public ReadOnly Property EditingPanelCursor As System.Windows.Forms.Cursor Implements System.Windows.Forms.IDataGridViewEditingControl.EditingPanelCursor
        Get
            Return MyBase.Cursor
        End Get
    End Property

    Public Function GetEditingControlFormattedValue(ByVal context As System.Windows.Forms.DataGridViewDataErrorContexts) As Object Implements System.Windows.Forms.IDataGridViewEditingControl.GetEditingControlFormattedValue
        Return ToShortDateString()
    End Function

    Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) Implements System.Windows.Forms.IDataGridViewEditingControl.PrepareEditingControlForEdit
    End Sub

    Public ReadOnly Property RepositionEditingControlOnValueChange As Boolean Implements System.Windows.Forms.IDataGridViewEditingControl.RepositionEditingControlOnValueChange
        Get
            Return False
        End Get
    End Property
#End Region
End Class