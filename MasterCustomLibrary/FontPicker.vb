'(C) David Eisenbeisz 2023
Imports System
Imports System.Drawing
Imports System.Drawing.Text
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices


Public Class FontPicker
    Private f_fntName As String
    Private f_pickerType As String
    Private f_styleNm As String

    Public ReadOnly Property FontName As String
        Get
            Return f_fntName
        End Get
    End Property
    Public ReadOnly Property StyleName As String
        Get
            Return f_styleNm
        End Get
    End Property
    Public Property PickerType As String
        Get
            Return f_pickerType
        End Get
        Set(value As String)
            f_pickerType = value
        End Set

    End Property

    Private Sub SysFontPicker()

        Dim winFonts As New InstalledFontCollection
        Dim fntFams() As FontFamily = winFonts.Families
        Dim fntNameLst As New List(Of String)

        LabPickerType.Text = "Select system font:"

        Try
            For Each font As FontFamily In fntFams
                fntNameLst.Add(font.Name)
            Next font

            fntNameLst.Sort()

            For Each fontnm As String In fntNameLst
                BxFonts.Items.Add(fontnm)
            Next fontnm
        Catch
        End Try

    End Sub


    Private Sub FontPicker_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub ButOK_Click(sender As Object, e As EventArgs) Handles ButOK.Click

        If f_pickerType = "sysFont" Then
            If BxFonts.SelectedIndex > -1 Then f_fntName = BxFonts.SelectedItem.ToString
        ElseIf f_pickerType = "textstyle" Then
            If BxFonts.SelectedIndex > -1 Then f_styleNm = BxFonts.SelectedItem.ToString
        Else
        End If
        Hide()

    End Sub

    Private Sub BxFonts_SelectedIndexChanged(sender As Object, e As EventArgs)

        If f_pickerType = "sysFont" Then
            If BxFonts.SelectedIndex > -1 Then f_fntName = BxFonts.SelectedItem.ToString
        ElseIf f_pickerType = "textstyle" Then
            If BxFonts.SelectedIndex > -1 Then f_styleNm = BxFonts.SelectedItem.ToString
        Else
        End If

    End Sub

    Private Sub ButCancel_Click(sender As Object, e As EventArgs) Handles ButCancel.Click

        f_fntName = ""
        f_styleNm = ""

        Hide()

    End Sub
    Private Sub TextStylePicker()

        LabPickerType.Text = "Select Autocad textstyle:"

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        Dim mstrTSTyleList As New List(Of String)

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()
            Dim tsTbl As TextStyleTable = acTrans.GetObject(dwgDB.TextStyleTableId, OpenMode.ForRead)

            For Each objid As ObjectId In tsTbl
                Dim tstyle As TextStyleTableRecord = acTrans.GetObject(objid, OpenMode.ForRead)
                If tstyle IsNot Nothing Then
                    mstrTSTyleList.Add(tstyle.Name)
                End If
            Next

            mstrTSTyleList.Sort()

            For i As Integer = 0 To mstrTSTyleList.Count - 1
                If Not mstrTSTyleList(i) = "" Then BxFonts.Items.Add(mstrTSTyleList(i))
            Next
            acTrans.Commit()
        End Using
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BxFonts.SelectedIndexChanged

        If f_pickerType = "sysFont" Then
            If BxFonts.SelectedIndex > -1 Then f_fntName = BxFonts.SelectedItem.ToString
        ElseIf f_pickerType = "textstyle" Then
            If BxFonts.SelectedIndex > -1 Then f_styleNm = BxFonts.SelectedItem.ToString
        Else
        End If
    End Sub

    Private Sub FontPicker_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        If f_pickerType = "sysFont" Then
            SysFontPicker()
        ElseIf f_pickerType = "textstyle" Then
            TextStylePicker()
        End If

    End Sub
End Class