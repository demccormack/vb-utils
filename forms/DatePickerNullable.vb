Public Class DatePickerNullable
    Inherits DateTimePicker
    Private realDate As Boolean = True
    Private dateFormat As DateTimePickerFormat
    Private nullFormat As String

    Public Sub New(Optional ByVal _format As DateTimePickerFormat = DateTimePickerFormat.Short, Optional ByVal strCustomFormat As String = " ")
        MyBase.New()
        Format = _format
        dateFormat = Format
        nullFormat = strCustomFormat
    End Sub

    Protected Overridable ReadOnly Property TextToValue() As Object
        Get
            If String.IsNullOrWhiteSpace(Me.Text) Then
                Return DBNull.Value
            Else
                Return CDate(Me.Text)
            End If
        End Get
    End Property

    Public Overridable Shadows Property Value As Object
        Get
            If IsRealDate Then
                Return MyBase.Value
            Else
                Return DBNull.Value
            End If
        End Get
        Set(ByVal val As Object)
            If IsDBNull(val) Then
                IsRealDate = False
            Else
                IsRealDate = True
                MyBase.Value = Convert.ToDateTime(val)
            End If
        End Set
    End Property

    ''' <summary>Gets or sets whether the control displays null or a date value, and sets the Format accordingly.
    ''' Do not set this; instead set the Value Property to DBNull.Value or MyBase.Value.</summary>
    ''' <remarks>Private property only. Set the Value property to DBNull.Value to use this externally.</remarks>
    Private Property IsRealDate As Boolean
        Get
            Return realDate
        End Get
        Set(ByVal val As Boolean)
            realDate = val
            If val = True Then
                Format = dateFormat
                CustomFormat = Nothing
            Else
                Format = DateTimePickerFormat.Custom
                CustomFormat = nullFormat
            End If
        End Set
    End Property

    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Delete) Or (e.KeyCode = Keys.Back) Then
            Value = DBNull.Value
            OnTextChanged(EventArgs.Empty)
        End If
        MyBase.OnKeyDown(e)
    End Sub

    Protected Overrides Sub OnCloseUp(ByVal eventargs As System.EventArgs)
        If Not KeyIsDown(Keys.Escape) Then
            Value = MyBase.Value
        End If
        MyBase.OnCloseUp(eventargs)
    End Sub

    Public Function ToShortDateString() As String
        If IsRealDate Then
            Return MyBase.Value.ToShortDateString
        Else
            Return String.Empty
        End If
    End Function
End Class