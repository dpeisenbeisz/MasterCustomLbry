'(C) David Eisenbeisz 2023

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports System.Xml.Serialization
Public Class BlockInfo
    'Inherits CollectionBase

    Private x_Name As String
    Private x_BlockName As String
    Private x_Handle As String
    Private x_refObjId As String
    Private x_bTRid As String
    'Private x_refCount As Integer
    Private x_insPt As String
    Private x_rot As Double
    Private x_hasAttr As Boolean
    Private x_attrCol As AttributeCollection
    Private x_isDynamic As Boolean
    Private x_properties As DynamicBlockReferencePropertyCollection
    Private x_RefCount As Integer
    Private x_DynamicName As String

    Public Sub New()
        MyBase.New
    End Sub

    Public Sub New(ByRef blkRef As BlockReference)
        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        x_Name = blkRef.Name
        x_refObjId = blkRef.ObjectId.ToString
        x_Handle = blkRef.Handle.ToString
        x_bTRid = blkRef.BlockId.ToString
        x_insPt = blkRef.Position.ToString
        x_rot = blkRef.Rotation.ToString
        x_BlockName = blkRef.BlockName
        If blkRef.AttributeCollection IsNot Nothing AndAlso blkRef.AttributeCollection.Count > 0 Then
            x_hasAttr = True
            x_attrCol = blkRef.AttributeCollection
        End If
        If blkRef.IsDynamicBlock Then
            x_DynamicName = blkRef.Name
            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim bid As ObjectId = blkRef.DynamicBlockTableRecord
                Dim dyBref As BlockTableRecord = TryCast(actrans.GetObject(bid, OpenMode.ForRead), BlockTableRecord)
                If dyBref IsNot Nothing Then
                    x_Name = dyBref.Name
                End If
                actrans.Commit()
            End Using
            x_isDynamic = True
            x_properties = blkRef.DynamicBlockReferencePropertyCollection
        Else
            x_isDynamic = False
        End If
    End Sub

    Public Sub New(ByRef blkRefID As ObjectId)

        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database

        If Not blkRefID = ObjectId.Null Then
            x_refObjId = blkRefID.ToString
            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim bRef As BlockReference = actrans.GetObject(blkRefID, OpenMode.ForRead)
                Dim blkTbl As BlockTable = actrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                x_Handle = bRef.Handle.ToString
                x_Name = bRef.Name
                x_BlockName = bRef.BlockName
                x_bTRid = bRef.BlockId.ToString
                x_insPt = bRef.Position.ToString
                x_rot = bRef.Rotation.ToString
                If bRef.AttributeCollection IsNot Nothing AndAlso bRef.AttributeCollection.Count > 0 Then
                    x_hasAttr = True
                    x_attrCol = bRef.AttributeCollection
                Else
                    x_hasAttr = False
                End If
                If bRef.IsDynamicBlock Then
                    x_DynamicName = bRef.Name
                    Dim bid As ObjectId = bRef.DynamicBlockTableRecord
                    Dim dyBref As BlockTableRecord = TryCast(actrans.GetObject(bid, OpenMode.ForRead), BlockTableRecord)
                    If dyBref IsNot Nothing Then
                        x_Name = dyBref.Name
                    End If
                    x_isDynamic = True
                    x_properties = bRef.DynamicBlockReferencePropertyCollection
                Else
                    x_isDynamic = False
                End If
                actrans.Commit()
            End Using
        End If
    End Sub

    'Public Property Bindex As Long
    '    Get
    '        Return x_blkIndex
    '    End Get
    '    Set(value As Long)
    '        x_blkIndex = value
    '    End Set
    'End Property

    <XmlAttribute("RefName")>
    Public Property RefName As String
        Get
            Return x_Name
        End Get
        Set(value As String)
            x_Name = value
        End Set
    End Property

    <XmlAttribute("RefHandle")>
    Public Property RefHandle As String
        Get
            Return x_Handle
        End Get
        Set(value As String)
            x_Handle = value
        End Set
    End Property

    <XmlAttribute("Insertion")>
    Public Property Insertion As String
        Get
            Return x_insPt
        End Get
        Set(value As String)
            x_insPt = value
        End Set
    End Property

    <XmlAttribute("Rotation")>
    Public Property Rotation As Double
        Get
            Return x_rot
        End Get
        Set(value As Double)
            x_rot = value
        End Set
    End Property

    <XmlAttribute("ReferenceId")>
    Public Property ReferenceId As String
        Get
            Return x_refObjId
        End Get
        Set(value As String)
            x_refObjId = value
        End Set
    End Property

    <XmlAttribute("HasAttributes")>
    Public Property HasAttributes As Boolean
        Get
            Return x_hasAttr
        End Get
        Set(value As Boolean)
            x_hasAttr = value
        End Set
    End Property

    <XmlAttribute("IsDynamic")>
    Public Property IsDynamic As Boolean
        Get
            Return x_isDynamic
        End Get
        Set(value As Boolean)
            x_isDynamic = value
        End Set
    End Property


    <XmlAttribute("RefCount")>
    Public Property RefCount As Integer
        Get
            Return x_RefCount
        End Get
        Set(value As Integer)
            x_RefCount = value
        End Set
    End Property

    <XmlAttribute("DynamicName")>
    Public Property DynamicName As String
        Get
            Return x_DynamicName
        End Get
        Set(value As String)
            x_DynamicName = value
        End Set
    End Property

    <XmlIgnore>
    Public ReadOnly Property CsvStr As String()
        Get
            Return GetCSVstr()
        End Get
    End Property

    <XmlIgnore>
    Public Property BlockName As String
        Get
            Return x_BlockName
        End Get
        Set(value As String)
            x_BlockName = value
        End Set
    End Property

    <XmlIgnore>
    Public Property BtrId As String
        Get
            Return x_bTRid
        End Get
        Set(value As String)
            x_bTRid = value
        End Set
    End Property

    <XmlIgnore>
    Public Property Attributes As AttributeCollection
        Get
            If x_attrCol IsNot Nothing Then
                Return x_attrCol
            Else
                Return Nothing
            End If
        End Get
        Set(value As AttributeCollection)
            x_attrCol = value
        End Set
    End Property

    <XmlIgnore>
    Public Property DynamicProperties As DynamicBlockReferencePropertyCollection
        Get
            If x_properties IsNot Nothing Then
                Return x_properties
            Else
                Return Nothing
            End If
        End Get
        Set(value As DynamicBlockReferencePropertyCollection)
            x_properties = value
        End Set
    End Property

    Private Function GetCSVstr() As String()
        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim csvStr(9) As String
        csvStr(0) = x_Name
        csvStr(1) = x_refObjId.ToString
        csvStr(3) = x_BlockName
        csvStr(4) = x_bTRid.ToString
        csvStr(5) = x_insPt.ToString
        csvStr(6) = Math.Round((x_rot * 180 / Math.PI), 5).ToString & "°"
        csvStr(7) = x_hasAttr.ToString
        csvStr(8) = x_isDynamic.ToString
        Return csvStr

    End Function

End Class

