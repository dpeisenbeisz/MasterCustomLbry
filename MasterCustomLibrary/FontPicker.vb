'(C) David Eisenbeisz 2023
Imports System
Imports System.Drawing
Imports System.Drawing.Text
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.ComponentModel
Imports MasterCustomLibrary.AcCommon

Public Class FontPicker
    Inherits System.Windows.Forms.Form

    Private f_fntName As String
    Private ReadOnly f_pickType As PickerType
    Private f_styleNm As String
    Private f_styleID As ObjectId
    Dim f_tsTbl As TextStyleTable
    Private f_standScale As StandardScaleType
    Private f_standScaleStr As String

    Public ReadOnly Property FontName As String
        Get
            If f_pickType = PickerType.SysFont Then
                Return f_styleNm
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property StyleId As ObjectId
        Get
            If f_pickType = PickerType.TextStyle Then
                Return f_styleID
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property StyleName As String
        Get
            If f_pickType = PickerType.TextStyle Then
                Return f_styleNm
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property PickType As PickerType
        Get
            Return f_pickType
        End Get
        'Set(value As PickerType)
        '    f_pickType = value
        'End Set

    End Property

    Public ReadOnly Property StandardSclType As StandardScaleType
        Get
            Return f_standScale
        End Get
    End Property

    Public ReadOnly Property StandardSclStr As String
        Get
            Return f_standScaleStr
        End Get
    End Property

    Public Sub New(isTextStyle As Boolean)

        ' This call is required by the designer.
        InitializeComponent()

        If isTextStyle Then
            f_pickType = PickerType.TextStyle
            TextStylePicker()
        Else
            f_pickType = PickerType.SysFont
            SysFontPicker()
        End If

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(pt As PickerType)

        ' This call is required by the designer.
        InitializeComponent()
        If pt = PickerType.TextStyle Then
            TextStylePicker()
        ElseIf pt = PickerType.SysFont Then
            SysFontPicker()
        ElseIf pt = PickerType.PlotScale Then
            ScalePicker()
        End If

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub SysFontPicker()

        Dim winFonts As New InstalledFontCollection
        Dim fntFams() As FontFamily = winFonts.Families
        Dim fntNameLst As New List(Of String)

        LabPickerType.Text = "Select system font:"
        Me.Text = "System Fonts"

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

    Private Sub ScalePicker()

        LabPickerType.Text = "Select Standard Scale for layout or Cancel if none"
        Me.Text = "Standard DWG Scales"

        Try
            Dim scaletypes As List(Of StandardScaleType) = [Enum].GetValues(GetType(StandardScaleType)).Cast(Of StandardScaleType)().ToList()

            For Each scl As String In scaletypes
                BxFonts.Items.Add(scl.ToString)
            Next

        Catch

        End Try

    End Sub

    Private Sub ButOK_Click(sender As Object, e As EventArgs) Handles ButOK.Click

        If f_pickType = PickerType.SysFont Then
            If BxFonts.SelectedIndex > -1 Then f_fntName = BxFonts.SelectedItem.ToString
        ElseIf f_pickType = PickerType.TextStyle Then
            If BxFonts.SelectedIndex > -1 Then
                f_styleNm = BxFonts.SelectedItem.ToString
                f_styleID = f_tsTbl(f_styleNm)
            End If
        ElseIf f_pickType = PickerType.PlotScale Then
            If BxFonts.SelectedIndex > -1 Then
                Dim tempStr As String = BxFonts.SelectedItem
                Dim sct As StandardScaleType = [Enum].Parse(GetType(StandardScaleType), tempStr)
                f_standScale = sct
                f_standScaleStr = tempStr
            Else
                f_standScale = Nothing
                f_standScaleStr = ""
            End If
        End If

        Hide()

    End Sub

    Private Sub ButCancel_Click(sender As Object, e As EventArgs) Handles ButCancel.Click

        f_fntName = ""
        f_styleNm = ""
        f_standScaleStr = ""
        f_styleID = ObjectId.Null
        f_standScaleStr = ""
        f_standScale = Nothing

        Hide()

    End Sub
    Private Sub TextStylePicker()

        LabPickerType.Text = "Select Autocad textstyle:"
        Me.Text = "Text Styles"

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        Dim mstrTSTyleList As New List(Of String)

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()
            f_tsTbl = acTrans.GetObject(dwgDB.TextStyleTableId, OpenMode.ForRead)

            For Each objid As ObjectId In f_tsTbl
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

    Private Sub BxFonts_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BxFonts.SelectedIndexChanged

        If f_pickType = AcCommon.Enums.PickerType.SysFont Then
            If BxFonts.SelectedIndex > -1 Then f_fntName = BxFonts.SelectedItem.ToString
        ElseIf f_pickType = PickerType.TextStyle Then
            If BxFonts.SelectedIndex > -1 Then
                f_styleNm = BxFonts.SelectedItem.ToString
                f_styleID = f_tsTbl(f_styleNm)
            End If
        ElseIf f_pickType = PickerType.PlotScale Then
            If BxFonts.SelectedIndex > -1 Then
                Dim tempStr As String = BxFonts.SelectedItem
                Dim sct As StandardScaleType = [Enum].Parse(GetType(StandardScaleType), tempStr)
                f_standScale = sct
                f_standScaleStr = tempStr
            End If
        Else

        End If
    End Sub

    Private Sub FontPicker_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        If f_tsTbl IsNot Nothing Then f_tsTbl.Dispose()

    End Sub

    'Private Sub FontPicker_Activated(sender As Object, e As EventArgs) Handles Me.Activated

    '    If f_pickType = PickerType.SysFont Then
    '        SysFontPicker()
    '    ElseIf f_pickType = pickertype.TextStyle Then
    '        TextStylePicker()
    '    End If

    'End Sub

    'Private Enum PickerType
    '    SysFont
    '    TextStyle
    '    PlotScale
    'End Enum

    Private Sub FontPicker_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
