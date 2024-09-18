Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports System.Runtime.InteropServices
Imports System.Runtime.CompilerServices

Public Class LTdata : Inherits CollectionBase

    Private ReadOnly m_symbolDash As Integer
    Private ReadOnly m_name As String
    Private ReadOnly m_numDashes As Integer
    Private ReadOnly m_shapeName As Integer
    Private ReadOnly m_textStyleId As ObjectId
    Private ReadOnly m_shapeFile As String
    Private ReadOnly m_patternLength As Double
    Private ReadOnly m_desc As String
    'Private m_comments As String
    Private ReadOnly m_toFit As String
    Private ReadOnly m_shapeNumber As Integer
    Private ReadOnly m_shapeOffset As Vector2d
    Private ReadOnly m_shapeX As Double
    Private ReadOnly m_shapeY As Double
    Private ReadOnly m_shapeScale As Double
    Private ReadOnly m_textStr As String
    Private ReadOnly m_UcsOrient As Boolean
    Private ReadOnly m_shapeRotation As Double
    Private ReadOnly m_dashlengths(11) As Double
    Private ReadOnly m_textStyleName As String
    Private ReadOnly m_hasText As Boolean
    Private ReadOnly m_hasShape As Boolean

    Public ReadOnly Property Dashes As Integer
        Get
            Return m_numDashes
        End Get
    End Property
    Public ReadOnly Property SymbolDash As Integer
        Get
            Return m_symbolDash
        End Get
    End Property

    Public ReadOnly Property Name As String
        Get
            Return m_name
        End Get
    End Property
    Public ReadOnly Property TextStyleName As String
        Get
            Return m_textStyleName
        End Get
    End Property
    Public ReadOnly Property HasText As Boolean
        Get
            Return m_hasText
        End Get
    End Property
    Public ReadOnly Property HasShape As Boolean
        Get
            Return m_hasShape
        End Get
    End Property
    Public ReadOnly Property ShapeFile As String
        Get
            Return m_shapeFile
        End Get
    End Property
    Public ReadOnly Property ShapeNumber As Integer
        Get
            Return m_shapeNumber
        End Get
    End Property
    Public ReadOnly Property PatternLength As Double
        Get
            Return m_patternLength
        End Get
    End Property
    Public ReadOnly Property Description As String
        Get
            Return m_desc
        End Get
    End Property

    Public ReadOnly Property IsFit As Boolean
        Get
            Return m_toFit
        End Get
    End Property
    Public ReadOnly Property ShapeName As String
        Get
            Return m_shapeName
        End Get
    End Property
    Public ReadOnly Property ShapeOffset As Vector2d
        Get
            Return m_shapeOffset
        End Get
    End Property

    Public ReadOnly Property ShapeX As Double
        Get
            Return m_shapeX
        End Get
    End Property
    Public ReadOnly Property ShapeY As Double
        Get
            Return m_shapeY
        End Get
    End Property
    Public ReadOnly Property ShapeScale As Double
        Get
            Return m_shapeScale
        End Get
    End Property
    Public ReadOnly Property TextString As String
        Get
            Return m_textStr
        End Get
    End Property
    Public ReadOnly Property HasUCSOrient As Boolean
        Get
            Return m_UcsOrient
        End Get
    End Property
    Public ReadOnly Property ShapeRotation As Double
        Get
            Return m_shapeRotation
        End Get
    End Property

    Public ReadOnly Property DashLengths(i As Integer) As Double
        Get
            Return m_dashlengths(i)
        End Get
    End Property
    Public ReadOnly Property TextStyleID As ObjectId
        Get
            Return m_textStyleId
        End Get
    End Property

    Public ReadOnly Property DLengths As Double()
        Get
            Return m_dashlengths
        End Get
    End Property

    Public Sub New()
        MyBase.New
    End Sub

    Public Sub New(LtRecId As ObjectId)
        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database

        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction()
            Dim ltRec As LinetypeTableRecord = TryCast(acTrans.GetObject(LtRecId, OpenMode.ForRead), LinetypeTableRecord)
            If ltRec IsNot Nothing Then
                m_hasShape = False
                m_hasText = False
                m_desc = ltRec.AsciiDescription
                'm_comments = ltRec.Comments
                m_patternLength = ltRec.PatternLength
                m_numDashes = ltRec.NumDashes
                m_name = ltRec.Name
                m_desc = ltRec.AsciiDescription

                For i As Integer = 0 To m_numDashes - 1
                    Dim sStyleId As ObjectId = ltRec.ShapeStyleAt(i)
                    'm_patternLength = ltRec.PatternLength
                    'm_toFit = ltRec.IsScaledToFit
                    'If m_toFit = True Then Debug.Print(vbLf & m_name)
                    If Not sStyleId = ObjectId.Null Then
                        Dim shpNum As Integer = ltRec.ShapeNumberAt(i)
                        If shpNum = 0 Then
                            m_textStyleId = sStyleId
                            Dim tsRec As TextStyleTableRecord = TryCast(acTrans.GetObject(sStyleId, OpenMode.ForRead), TextStyleTableRecord)
                            m_textStr = ltRec.TextAt(i)
                            m_textStyleName = tsRec.Name
                            m_hasText = True
                            m_shapeOffset = ltRec.ShapeOffsetAt(i)
                            m_shapeX = m_shapeOffset.X
                            m_shapeY = m_shapeOffset.Y
                            m_symbolDash = i
                            m_shapeScale = ltRec.ShapeScaleAt(i)
                            m_UcsOrient = ltRec.ShapeIsUcsOrientedAt(i)
                            m_shapeRotation = ltRec.ShapeRotationAt(i) * 180 / Math.PI
                            m_dashlengths(i) = 1000
                        Else
                            m_hasShape = True
                            m_shapeNumber = shpNum
                            Dim tsRec As TextStyleTableRecord = TryCast(acTrans.GetObject(sStyleId, OpenMode.ForRead), TextStyleTableRecord)
                            m_shapeFile = tsRec.FileName
                            m_shapeOffset = ltRec.ShapeOffsetAt(i)
                            m_textStyleId = sStyleId
                            m_shapeX = m_shapeOffset.X
                            m_shapeY = m_shapeOffset.Y
                            m_symbolDash = i
                            m_shapeScale = ltRec.ShapeScaleAt(i)
                            m_UcsOrient = ltRec.ShapeIsUcsOrientedAt(i)
                            m_shapeRotation = ltRec.ShapeRotationAt(i) * 180 / Math.PI
                            m_dashlengths(i) = 1000
                        End If
                    Else
                        m_dashlengths(i) = ltRec.DashLengthAt(i)
                    End If
                Next
            End If
        End Using
    End Sub

End Class

