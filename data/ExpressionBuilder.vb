''' <summary>
''' A form to allow the user to write expressions which are acceptable to the System.Data.DataColumn.Expression property.
''' </summary>
''' <remarks></remarks>
Public Class ExpressionBuilder
    Inherits Form

    Public Shared Result, originalExpression As String
    Private WithEvents TbExpression, TbStrInput, TbNumInput As New TextBox
    Private WithEvents LbColumn, LbOperator As New ListBox
    Private WithEvents BtnColumn, BtnOperator, BtnUndo, BtnRedo, BtnCancel, BtnStrInput, BtnNumInput As New Button
    Public ReadOnly Property Expression As String
        Get
            Return TbExpression.Text
        End Get
    End Property

    Public Sub New(ByVal columns As String(), Optional ByVal caption As String = Nothing, Optional ByVal startexpression As String = "")
        Result = Nothing
        Setup(columns, caption)
        originalExpression = startexpression
        TbExpression.Text = originalExpression
    End Sub

    Private Shared LastAppended As String
    Private Sub AppendStr(ByVal value As String)
        TbExpression.AppendText(value)
        LastAppended = value
    End Sub

    Private isUndoing As Boolean
    Private Shared currentUndoIndex As Integer
    Private Shared TextBeforeChange(0) As String
    Private Sub BuildUndoStack() Handles TbExpression.TextChanged
        If (Not isUndoing) Then
            If (TextBeforeChange(0) = Nothing) And (TextBeforeChange.Length = 1) Then
                TextBeforeChange(0) = ""
            End If
            If (TextBeforeChange.Length = 10) Then
                Dim tmpAry(8) As String
                For i As Integer = 0 To 8
                    tmpAry(i) = TextBeforeChange(i + 1)
                Next
                TextBeforeChange = tmpAry
            End If
            ArrayAddElement(TextBeforeChange, TbExpression.Text)
            currentUndoIndex = (TextBeforeChange.Length - 2)
        End If
    End Sub

    Private Sub Undo() Handles BtnUndo.Click
        isUndoing = True
        TbExpression.Text = TextBeforeChange(currentUndoIndex)
        If (currentUndoIndex > 0) Then
            currentUndoIndex -= 1
        End If
        isUndoing = False
    End Sub

    Public Shared Function EvaluateOperator(ByVal desc As String) As String
        isContains = False
        Select Case desc
            Case "Equals" : Return " = "
            Case "Contains" : isContains = True
                Return " LIKE "
            Case "Is Greater Than" : Return " > "
            Case "Is Less Than" : Return " < "
            Case "Is Greater Than or Equal To" : Return " >= "
            Case "Is Less Than or Equal To" : Return " <= "
            Case "Does Not Equal" : Return " <> "
            Case "(", ")" : Return desc
            Case "AND", "OR" : Return " " & desc & " "
            Case Else : Return " " & desc & " "
        End Select
    End Function

    Private Sub InsertColumn() Handles BtnColumn.Click, LbColumn.DoubleClick
        Dim s As String = LbColumn.SelectedItem.ToString
        AppendStr(s)
    End Sub

    Private Sub InsertOperator() Handles BtnOperator.Click, LbOperator.DoubleClick
        Dim s As String = EvaluateOperator(LbOperator.SelectedItem.ToString)
        AppendStr(s)
    End Sub

    Private Shared isContains As Boolean
    Private Sub InsertStrInput() Handles BtnStrInput.Click
        Dim s As String
        If isContains Then
            s = "'*" & TbStrInput.Text & "*'"
        Else
            s = "'" & TbStrInput.Text & "'"
        End If
        AppendStr(s)
        TbStrInput.Text = "Text/Date"
    End Sub

    Private Sub InsertNumInput() Handles BtnNumInput.Click
        If Not IsNumeric(TbNumInput.Text) Then
            Exit Sub
        End If
        Dim s As String = TbNumInput.Text
        AppendStr(s)
        TbNumInput.Text = "Numeric Input"
    End Sub

    Private Sub tbClear(ByVal sender As TextBox, ByVal e As EventArgs) Handles TbStrInput.GotFocus, TbNumInput.GotFocus
        sender.Clear()
    End Sub

    Private Sub Setup(ByVal columns As String(), Optional ByVal caption As String = Nothing)
        FormBorderStyle = FormBorderStyle.FixedSingle
        Size = New Size(300, 300)
        If caption Is Nothing Then
            Text = "Expression Builder"
        Else
            Text = caption & " - Expression Builder"
        End If
        ShowIcon = False
        MinimizeBox = False
        MaximizeBox = False

        TbExpression.Multiline = True
        TbExpression.Dock = DockStyle.Top
        TbExpression.Height = 100
        Controls.Add(TbExpression)

        ButtonCreate(Me, "Add", -2, 245, BtnColumn, , , , 80, 23)
        ButtonCreate(Me, "Add", 105, 245, BtnOperator, , , , 80, 23)
        ButtonCreate(Me, "Undo", 212, 105, BtnUndo, , , , 80, 23)
        ButtonCreate(Me, "Cancel", 212, 131, BtnCancel, , , , 80, 23)
        ButtonCreate(Me, "Add", 212, 190, BtnStrInput, , , , 80, 23)
        ButtonCreate(Me, "Add", 212, 245, BtnNumInput, , , , 80, 23)
        Me.CancelButton = BtnCancel

        'ButtonCreate(Me, "Redo", 212, 77, btnRedo, , , , 80, 23)
        'btnRedo.BringToFront()

        LbColumn.Items.AddRange(columns)
        LbColumn.Location = New Point(0, 105)
        LbColumn.Size = New Size(80, 134)
        Controls.Add(LbColumn)

        LbOperator.Items.AddRange(New String() {"Equals", "Contains", "Is Greater Than", "Is Less Than", "Does Not Equal", "AND", "OR", "(", ")",
                                               "Is Greater Than or Equal To", "Is Less Than or Equal To"})
        LbOperator.Location = New Point(86, 105)
        LbOperator.Size = New Size(122, 134)
        Controls.Add(LbOperator)

        TbStrInput.Text = "Text/Date"
        TbStrInput.Location = New Point(214, 164)
        TbStrInput.Width = 80
        Controls.Add(TbStrInput)

        TbNumInput.Text = "Numeric Input"
        TbNumInput.Location = New Point(214, 219)
        TbNumInput.Width = 80
        Controls.Add(TbNumInput)

        Dim cms As New ContextMenuStrip
        cms.Items.Add("Remove Last", Nothing, AddressOf Undo)
        cms.Items.Add("Cancel Filter", Nothing, AddressOf CloseWithoutChange)
        ContextMenuStrip = cms
    End Sub

    Private Sub CloseWithoutChange() Handles BtnCancel.Click
        TbExpression.Text = originalExpression
        Close()
    End Sub

    Private Sub ReturnValue() Handles Me.FormClosing
        Result = Expression
    End Sub
End Class