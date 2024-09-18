Imports System.IO
Imports System.Math
Imports System.Text
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
'Imports Autodesk.AutoCAD.PlottingServices
Imports System.Xml
Imports System.Xml.Schema
Imports Autodesk.AutoCAD.Colors
Imports System.Xml.Serialization


Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub




    Private Sub GetViews()

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
        Dim layouts As New Collection
        Dim vports As New List(Of ObjectId)

        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction()
            Dim layDict As DBDictionary = DwgDB.LayoutDictionaryId.GetObject(OpenMode.ForRead)
            For Each entry As DBDictionaryEntry In layDict
                Dim lay As Layout = acTrans.GetObject(entry.Value, OpenMode.ForRead) 'CType(entry.Value.GetObject(OpenMode.ForRead), Layout)
                If Not lay.LayoutName = "Model" Then
                    layouts.Add(lay.LayoutName)
                End If
            Next





            acTrans.Commit()


        End Using



    End Sub



End Class