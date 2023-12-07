Imports System.Windows.Forms
Public Class ArrowTypeBox


    Private m_ArrowType As Integer


    Public ReadOnly Property ArrowType As Integer
        Get
            Return m_ArrowType
        End Get
    End Property

    Private Sub ButOK_Click(sender As Object, e As EventArgs) Handles ButOK.Click

        If ButType1.Checked Then
            m_ArrowType = 1
        ElseIf ButType2.Checked Then
            m_ArrowType = 2
        ElseIf ButType3.Checked Then
            m_ArrowType = 3
        ElseIf ButType4.Checked Then
            m_ArrowType = 4
        Else
            Exit Sub
        End If

        Me.DialogResult = DialogResult.OK

        Hide()

    End Sub

    Private Sub ButType1_CheckedChanged(sender As Object, e As EventArgs) Handles ButType1.CheckedChanged

        If ButType1.Checked Then
            m_ArrowType = 1
        ElseIf ButType2.Checked Then
            m_ArrowType = 2
        ElseIf ButType3.Checked Then
            m_ArrowType = 3
        ElseIf ButType4.Checked Then
            m_ArrowType = 4
        End If

    End Sub

    Private Sub ButType2_CheckedChanged(sender As Object, e As EventArgs) Handles ButType2.CheckedChanged
        If ButType1.Checked Then
            m_ArrowType = 1
        ElseIf ButType2.Checked Then
            m_ArrowType = 2
        ElseIf ButType3.Checked Then
            m_ArrowType = 3
        ElseIf ButType4.Checked Then
            m_ArrowType = 4
        End If

    End Sub

    Private Sub ButType3_CheckedChanged(sender As Object, e As EventArgs) Handles ButType3.CheckedChanged
        If ButType1.Checked Then
            m_ArrowType = 1
        ElseIf ButType2.Checked Then
            m_ArrowType = 2
        ElseIf ButType3.Checked Then
            m_ArrowType = 3
        ElseIf ButType4.Checked Then
            m_ArrowType = 4
        End If

    End Sub

    Private Sub ButType4_CheckedChanged(sender As Object, e As EventArgs) Handles ButType4.CheckedChanged
        If ButType1.Checked Then
            m_ArrowType = 1
        ElseIf ButType2.Checked Then
            m_ArrowType = 2
        ElseIf ButType3.Checked Then
            m_ArrowType = 3
        ElseIf ButType4.Checked Then
            m_ArrowType = 4
        End If

    End Sub
    Private Sub ButCancel_Click(sender As Object, e As EventArgs) Handles ButCancel.Click
        m_ArrowType = 0
        Me.DialogResult = DialogResult.Cancel
        Hide()
    End Sub
End Class

