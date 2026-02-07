Public Class LayoutPicker : Inherits System.Windows.Forms.Form
    Private m_pickedList As List(Of String)
    Private m_isLayouts As Boolean

    Public ReadOnly Property PickedList As List(Of String)
        Get
            Return m_pickedList
        End Get
    End Property

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(isLayouts As Boolean)

        ' This call is required by the designer.
        InitializeComponent()
        If Not isLayouts Then
            PickerLabel.Text = "Pick named views"
            Me.Text = "View Picker"
        End If

        m_isLayouts = isLayouts

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(itemList As List(Of String), isLayouts As Boolean)

        ' This call is required by the designer.
        InitializeComponent()
        If Not isLayouts Then
            PickerLabel.Text = "Pick named views"
            Me.Text = "View Picker"
        End If

        If itemList IsNot Nothing AndAlso itemList.Count > 0 Then
            For Each it As String In itemList
                ListBox1.Items.Add(it)
            Next
        End If

        m_isLayouts = isLayouts

    End Sub

    Private Sub LayoutPicker_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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