Module FormUtils
    ''' <summary>
    ''' Returns a new control of the type most appropriate to the data it will handle.
    ''' </summary>
    ''' <param name="dataType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ControlFromDataType(ByVal dataType As Type) As Control
        Select Case dataType
            Case GetType(DateTime) : Return New DatePickerNullable
            Case GetType(Boolean) : Return New CheckBox
            Case Else : Return New TextBox
        End Select
    End Function

    ''' <summary>
    ''' Returns the name of the the property usually used for binding to a data source.
    ''' </summary>
    ''' <param name="controlType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PropertyUsuallyBound(ByVal controlType As Type) As String
        Select Case controlType
            Case GetType(DatePickerNullable), GetType(DateTimePicker)
                Return "Value"
            Case GetType(CheckBox)
                Return "Checked"
            Case GetType(TextBox)
                Return "Text"
            Case Else
                Throw New Exception("Type '" & controlType.ToString & "' not recognized")
        End Select
    End Function

    Public Sub ButtonCreate(ByRef container As ContainerControl, ByVal caption As String, ByVal x As Integer, ByVal y As Integer, Optional ByRef btn As Button = Nothing,
                         Optional ByRef parentArray As Button() = Nothing, Optional ByVal onClick As EventHandler = Nothing,
                         Optional ByVal anchorStyle As AnchorStyles = AnchorStyles.Right + AnchorStyles.Top, Optional ByVal width As Integer = 85,
                         Optional ByVal height As Integer = 25, Optional ByVal fntStyle As FontStyle = FontStyle.Bold, Optional ByVal fntSize As Single = 8)
        If btn Is Nothing Then
            btn = New Button
        End If
        btn.Text = caption
        btn.Location = New Point(x, y)
        btn.Size = New Size(width, height)
        btn.Font = New Font("Microsoft Sans Serif", fntSize, fntStyle, GraphicsUnit.Point)
        btn.Anchor = anchorStyle
        If parentArray IsNot Nothing Then
            ArrayAddElement(parentArray, btn)
        End If
        If onClick IsNot Nothing Then
            AddHandler btn.Click, onClick
        End If
        container.Controls.Add(btn)
    End Sub

    Public Sub AddEventHandlerToControl(ByRef ctl As Control, ByRef eh As EventHandler)
        Select Case ctl.GetType
            Case GetType(TextBox)
                Dim t As TextBox = CType(ctl, TextBox)
                AddHandler t.TextChanged, eh
            Case GetType(CheckBox)
                Dim c As CheckBox = CType(ctl, CheckBox)
                AddHandler c.CheckedChanged, eh
            Case GetType(DatePickerNullable), GetType(DateTimePicker)
                Dim dtp As DateTimePicker = CType(ctl, DateTimePicker)
                AddHandler dtp.ValueChanged, eh
        End Select
    End Sub

    Public Sub RemoveEventHandlerFromControl(ByRef ctl As Control, ByRef eh As EventHandler)
        Select Case ctl.GetType
            Case GetType(TextBox)
                Dim t As TextBox = CType(ctl, TextBox)
                RemoveHandler t.TextChanged, eh
            Case GetType(CheckBox)
                Dim c As CheckBox = CType(ctl, CheckBox)
                RemoveHandler c.CheckedChanged, eh
            Case GetType(DatePickerNullable), GetType(DateTimePicker)
                Dim dtp As DateTimePicker = CType(ctl, DateTimePicker)
                RemoveHandler dtp.ValueChanged, eh
        End Select
    End Sub

    Public Function DataGridViewColumnFromType(ByVal dataType As Type) As DataGridViewColumn
        Select Case dataType
            Case GetType(DateTime) : Return New CalendarColumn
            Case GetType(Boolean) : Return New DataGridViewCheckBoxColumn
            Case Else : Return New DataGridViewTextBoxColumn
        End Select
    End Function
End Module
