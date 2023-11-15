Public Class PlotSettingsForm

    Private m_PageSetups As Collection
    Private m_selectedSetup As String

    Public Property PageSetups As Collection
        Get
            Return m_PageSetups
        End Get
        Set(value As Collection)
            m_PageSetups = value
        End Set
    End Property

    Public ReadOnly Property SelectedSetup As String
        Get
            Return m_selectedSetup
        End Get
    End Property

    Private Sub PlotSettingsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Sub AddSetup(setupName As String)
        BoxPageSetups.Items.Add(setupName)
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs) Handles OkButton.Click

        Dim mySetup As String
        If Not String.IsNullOrEmpty(BoxPageSetups.Text) Then
            mySetup = BoxPageSetups.Text
            m_selectedSetup = mySetup
        End If

        Me.DialogResult = Windows.Forms.DialogResult.OK

        BoxPageSetups.Items.Clear()
        Hide()

    End Sub

    Private Sub CButton_Click(sender As Object, e As EventArgs) Handles CButton.Click
        BoxPageSetups.Items.Clear()
        Hide()
    End Sub
End Class