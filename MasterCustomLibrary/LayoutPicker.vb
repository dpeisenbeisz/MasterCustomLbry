Public Class LayoutPicker

    Private m_pickedList As List(Of String)

    Public ReadOnly Property PickedList As List(Of String)
        Get
            Return m_pickedList
        End Get
    End Property

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs) Handles OkButton.Click
        Dim mylist As New List(Of String)
        For Each it As String In ListBox1.SelectedItems
            mylist.Add(it)
        Next
        m_pickedList = mylist
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Hide()
    End Sub

    Private Sub CButton_Click(sender As Object, e As EventArgs) Handles CButton.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Hide()
    End Sub
End Class