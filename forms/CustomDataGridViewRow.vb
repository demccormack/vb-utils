Public Class CustomDataGridViewRow
    Inherits DataGridViewRow
    Private rowRequiresPencil As Boolean
    Public Property RequiresPencil As Boolean
        Get
            Return rowRequiresPencil
        End Get
        Set(ByVal value As Boolean)
            rowRequiresPencil = value
        End Set
    End Property
    Protected Overrides Sub PaintHeader(ByVal graphics As System.Drawing.Graphics, ByVal clipBounds As System.Drawing.Rectangle,
                                        ByVal rowBounds As System.Drawing.Rectangle, ByVal rowIndex As Integer,
                                        ByVal rowState As System.Windows.Forms.DataGridViewElementStates, ByVal isFirstDisplayedRow As Boolean,
                                        ByVal isLastVisibleRow As Boolean, ByVal paintParts As System.Windows.Forms.DataGridViewPaintParts)
        MyBase.PaintHeader(graphics, clipBounds, rowBounds, rowIndex, rowState, isFirstDisplayedRow, isLastVisibleRow, paintParts)
        If RequiresPencil And (clipBounds.Y <> 0) Then
            graphics.DrawImage(CType(DataGridView, CustomDataGridView).EditingPencil, New Point(clipBounds.X + 4, clipBounds.Y + 3))
        End If
    End Sub
End Class