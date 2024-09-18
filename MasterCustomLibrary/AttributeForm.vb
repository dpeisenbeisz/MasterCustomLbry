'(C) David Eisenbeisz 2023
Imports System.Windows.Forms

Public Class AttributeForm
    Inherits Form

    Private m_attCol As List(Of String)

    Public Sub New()
        MyBase.New
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Public Sub New(atList As List(Of String))
        MyBase.New
        InitializeComponent()
        m_attCol = atList
        ShowList()
    End Sub

    Public Property Attributes As List(Of String)
        Get
            Return m_attCol
        End Get
        Set(value As List(Of String))
            m_attCol = value

        End Set
    End Property

    Private Sub ShowList()

        If m_attCol IsNot Nothing AndAlso m_attCol.Count > 0 Then
            For Each str As String In m_attCol
                Dim newline() As String = str.Split(",")
                DGV.Rows.Add(newline)
            Next
        End If

    End Sub

    Private Sub ButClose_Click(sender As Object, e As EventArgs) Handles ButClose.Click
        Me.DialogResult = DialogResult.OK
        Close()
    End Sub
End Class