Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Windows.Forms

Public Class Picker : Inherits Form

    Private m_PickColl As Collection

    Public ReadOnly Property PickCol As Collection
        Get
            Return m_PickColl
        End Get
    End Property

    Private Sub OK_Click(sender As Object, e As EventArgs) Handles ButOK.Click
        Dim mySelection As ListBox.SelectedObjectCollection = BxList.SelectedItems
        Dim tempCol As New Collection
        For Each it As String In mySelection
            tempCol.Add(it)
        Next
        Me.DialogResult = DialogResult.OK
        m_PickColl = tempCol
        Hide()
    End Sub
    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles ButCancel.Click
        m_PickColl = Nothing
        Me.DialogResult = DialogResult.Cancel
        Hide()
    End Sub

    Private Sub BxLyr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BxList.SelectedIndexChanged

    End Sub
End Class