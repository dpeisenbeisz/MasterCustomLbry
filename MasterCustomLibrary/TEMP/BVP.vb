﻿Imports System.Reflection
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsInterface
Imports System.Text


Public Class BVP

    Private ObjIds As ObjectIdCollection

    Public Sub GetBlocks()

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor
        'Dim tvs As TypedValue() = New TypedValue(0) {New TypedValue(CInt(DxfCode.LayerName), layerName)}
        Dim tv As TypedValue() = New TypedValue(0) {New TypedValue(CInt(DxfCode.Start), "INSERT")}
        Dim sf As New SelectionFilter(tv)
        Dim psr As PromptSelectionResult = ed.SelectAll(sf)
        Dim obIds As New ObjectIdCollection(psr.Value.GetObjectIds())

        ' Dim i As Integer = 0

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
            Dim upperlimit As Long = obIds.Count
            Dim blkArray As New List(Of BlkInfo)
            Dim dgv As DataGridView = DGV1
            Dim blkIds As New ObjectIdCollection
            Dim rawBlkLst As New List(Of String)

            For Each id As ObjectId In obIds
                Dim dbOb As DBObject = acTrans.GetObject(id, OpenMode.ForRead)
                If TypeOf dbOb Is BlockReference Then
                    blkIds.Add(id)
                    Dim tbref As BlockReference = TryCast(dbOb, BlockReference)
                    If tbref IsNot Nothing Then
                        rawBlkLst.Add(tbref.Name)
                        Dim bi As New BlkInfo(tbref)
                        blkArray.Add(bi)
                    End If
                End If
            Next

            Dim refCount As Integer = 0
            Dim refNumberDic As New Dictionary(Of String, Integer)
            Dim fnlBlkLst As IEnumerable(Of String) = rawBlkLst.Distinct

            'For Each st As String In rawBlkLst
            '    fnlBlkLst.Add(st)
            'Next

            'fnlBlkLst.Distinct

            For Each nameStr As String In fnlBlkLst
                For Each testStr As String In rawBlkLst
                    If nameStr = testStr Then refCount += 1
                Next
                refNumberDic.Add(nameStr, refCount)
                refCount = 0
            Next

            Dim blkCol As New Collection

            For Each id As ObjectId In blkIds
                'For Each bi As BlkInfo In blkArray
                Dim bRef As BlockReference = TryCast(acTrans.GetObject(id, OpenMode.ForRead), BlockReference)
                Dim blkInf As New BlkInfo(id, refNumberDic(bRef.Name))
                Dim curLine() As String = blkInf.CsvStr
                blkCol.Add(blkInf)
                'curLine(2) = refNumberDic(bRef.Name)
                'DGV1.Rows.Add(curLine)
            Next

            dgv.DataSource = blkCol

            'Dim bCount As New Dictionary(Of String, Long)
            'Dim x As Long = -1
            'Dim j As Long

            'For i = 0 To blkArray.Count - 1
            '    Dim binf As BlkInfo = blkArray(i)
            '    Dim testName = binf.BlkName
            '    For j = 0 To blkArray.Count - 1
            '        Dim testinf As BlkInfo = blkArray(j)
            '        If testName = testinf.BlkName Then x += 1
            '    Next
            '    bCount.Add(testName, x)
            'Next
            acTrans.Commit()
        End Using
    End Sub

    Private Sub ButOk_Click(sender As Object, e As EventArgs)
        Me.Hide()

    End Sub

    Private Sub ButCancel_Click(sender As Object, e As EventArgs)
        Me.Hide()

    End Sub

    Private Sub BlockViewPanel_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call GetBlocks()

    End Sub

    Private Sub DGV1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV1.CellContentDoubleClick

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor

        Dim rowIdx As Long = e.RowIndex
        Dim colIdx As Long = e.ColumnIndex
        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction

            Select Case colIdx

                Case Is = 7
                    Dim blkCol As Collection = DGV1.DataSource
                    Dim bi As BlkInfo = blkCol(rowIdx)
                    Dim attList As AttributeCollection = bi.Attributes
                    Dim sb As New StringBuilder
                    Dim headStr As String = "Name" & vbTab & "Value"
                    sb.AppendLine(headStr)
                    For Each atId As ObjectId In attList
                        Dim dbOb As DBObject = acTrans.GetObject(atId, OpenMode.ForRead)
                        If TypeOf dbOb Is AttributeReference Then
                            Dim atRef As AttributeReference = CType(dbOb, AttributeReference)
                            Dim atStr As String = atRef.Tag.ToString & vbTab & atRef.TextString
                            sb.AppendLine(atStr)
                        End If
                    Next
                    MessageBox.Show(sb.ToString)
                Case Is = 8
                    Dim blkCol As Collection = DGV1.DataSource
                    Dim bi As BlkInfo = blkCol(rowIdx)
                    Dim dynList As DynamicBlockReferencePropertyCollection = bi.DynamicProperties
                    Dim sb As New StringBuilder
                    Dim headStr As String = "Property" & vbTab & "Value"
                    sb.AppendLine(headStr)
                    For Each dynprop As DynamicBlockReferenceProperty In dynList
                        Dim propName As String = dynprop.PropertyName
                        Dim propValue As String = dynprop.Value.ToString
                        Dim tStr As String = propName & vbTab & propValue
                        sb.AppendLine(tStr)
                    Next
                    MessageBox.Show(sb.ToString)
                Case Else

            End Select
        End Using

    End Sub


End Class