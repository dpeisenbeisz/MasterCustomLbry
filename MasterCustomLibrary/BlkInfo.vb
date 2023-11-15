Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class BlkInfo : Inherits CollectionBase

    Private x_Bname As String
    Private ReadOnly x_BlockName As String
    Private x_refObjId As ObjectId
    Private ReadOnly x_bTRid As ObjectId
    Private x_blkIndex As Long
    Private x_insPt As Point3d
    Private x_rot As Double
    Private ReadOnly x_hasAttr As Boolean
    Private ReadOnly x_attrCol As AttributeCollection
    Private ReadOnly x_isDynamic As Boolean
    Private ReadOnly x_properties As Collection

    Public Sub New()
        MyBase.New
    End Sub

    Public Sub New(ByRef blkRef As BlockReference)
        x_Bname = blkRef.Name
        x_refObjId = blkRef.ObjectId
        x_bTRid = blkRef.BlockId
        x_insPt = blkRef.Position
        x_rot = blkRef.Rotation
        If blkRef.AttributeCollection IsNot Nothing Then
            x_hasAttr = True
            x_attrCol = blkRef.AttributeCollection
        End If
        If blkRef.IsDynamicBlock Then
            x_isDynamic = True
            For Each prop As DynamicBlockReferenceProperty In blkRef.DynamicBlockReferencePropertyCollection
                x_properties.Add(prop.PropertyName)
            Next
        End If
    End Sub

    Public Sub New(ByRef blkRefID As ObjectId)

        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database

        If Not blkRefID = ObjectId.Null Then
            x_refObjId = blkRefID
            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim bRef As BlockReference = actrans.GetObject(x_refObjId, OpenMode.ForRead)
                Dim blkTbl As BlockTable = actrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                x_Bname = bRef.Name
                x_BlockName = bRef.BlockName
                x_bTRid = blkTbl(x_Bname)
                x_insPt = bRef.Position
                x_rot = bRef.Rotation
                If bRef.AttributeCollection IsNot Nothing Then
                    x_hasAttr = True
                    x_attrCol = bRef.AttributeCollection
                End If
                If bRef.IsDynamicBlock Then
                    x_isDynamic = True
                    For Each prop As DynamicBlockReferenceProperty In bRef.DynamicBlockReferencePropertyCollection
                        x_properties.Add(prop.PropertyName)
                    Next
                End If
                actrans.Commit()
            End Using
        End If
    End Sub

    Public Property Bindex As Long
        Get
            Return x_blkIndex
        End Get
        Set(value As Long)
            x_blkIndex = value
        End Set
    End Property

    Public ReadOnly Property Attributes As AttributeCollection
        Get
            If x_attrCol IsNot Nothing Then
                Return x_attrCol
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property DynamicProperties As Collection
        Get
            If x_properties IsNot Nothing Then
                Return x_properties
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property HasAttributes As Boolean
        Get
            Return x_hasAttr
        End Get
    End Property
    Public ReadOnly Property IsDynmaic As Boolean
        Get
            Return x_isDynamic
        End Get
    End Property

    Public Property Insertion As Point3d
        Get
            Return x_insPt
        End Get
        Set(value As Point3d)
            x_insPt = value
        End Set
    End Property

    Public Property Rotation As Double
        Get
            Return x_rot
        End Get
        Set(value As Double)
            x_rot = value
        End Set
    End Property

    Public Property BlkName As String
        Get
            Return x_Bname
        End Get
        Set(value As String)
            x_Bname = value
        End Set
    End Property

    Public Property RefId As ObjectId
        Get
            Return x_refObjId
        End Get
        Set(value As ObjectId)
            x_refObjId = value
        End Set
    End Property

    Public ReadOnly Property BtrID As ObjectId
        Get
            Return x_bTRid
        End Get
    End Property
    Public ReadOnly Property CsvStr As String()
        Get
            Return GetCSVstr()
        End Get
    End Property

    Private Function GetCSVstr() As String()
        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim csvStr(9) As String

        csvStr(0) = x_Bname
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

