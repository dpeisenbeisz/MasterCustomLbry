Imports Autodesk.AutoCAD.Geometry

Namespace AcCommon
    Public Class RetainingWall
        Inherits CollectionBase

        Private m_HasConcBase As Boolean
        Private m_vertFootStep As Double
        Private m_footCover As Double
        Private m_VertFactor As Double
        Private m_brickHt As Double
        Private m_brickLength As Double
        Private m_footThick As Double
        Private m_FootStepCourses As Integer
        Private m_topStepHeight As Double
        Private m_StartConcBase As Point3d
        Private m_StartTopWall As Point3d
        Private m_startFooting As Point3d
        Public Direction As Integer
        Private m_useFootPoint As Boolean

        Public Sub New()
            MyBase.New
        End Sub

        Public Sub New(vfactor As Double)
            MyBase.New
            m_VertFactor = vfactor
        End Sub
        Public Sub New(vfactor As Double, hasConcBase As Boolean, brkht As Double, brklength As Double, footcvr As Double, footingThck As Double, Optional ftStpHeight As Double = 0.667)
            MyBase.New
            m_VertFactor = vfactor
            m_HasConcBase = hasConcBase
            m_brickHt = brkht
            m_brickLength = brklength
            m_footCover = footcvr
            m_footThick = footingThck

            If hasConcBase Then
                m_vertFootStep = ftStpHeight
            Else
                m_vertFootStep = brkht
            End If

        End Sub

        Public Property HasConcreteBase As Boolean
            Get
                Return m_HasConcBase
            End Get
            Set(value As Boolean)
                m_HasConcBase = value
            End Set
        End Property
        Public Property UseFootPoint As Boolean
            Get
                Return m_useFootPoint
            End Get
            Set(value As Boolean)
                m_useFootPoint = value
            End Set
        End Property

        Public Property StartTopWall As Point3d
            Get
                Return m_StartTopWall
            End Get
            Set(value As Point3d)
                m_StartTopWall = value
            End Set
        End Property

        Public Property StartConcreteStem As Point3d
            Get
                Return m_StartConcBase
            End Get
            Set(value As Point3d)
                m_StartConcBase = value
            End Set
        End Property

        Public Property StartTopFooting As Point3d
            Get
                Return m_startFooting
            End Get
            Set(value As Point3d)
                m_startFooting = value
            End Set
        End Property
        Public Property FootingStepHeight As Double
            Get
                Return m_vertFootStep
            End Get
            Set(value As Double)
                m_vertFootStep = value
            End Set
        End Property
        Public Property FootStepCourses As Integer
            Get
                Return m_FootStepCourses
            End Get
            Set(value As Integer)
                m_FootStepCourses = value
            End Set
        End Property

        Public Property FootingCover As Double
            Get
                Return m_footCover
            End Get
            Set(value As Double)
                m_footCover = value
            End Set
        End Property

        Public Property VertFactor As Double
            Get
                Return m_VertFactor
            End Get
            Set(value As Double)
                m_VertFactor = value
            End Set
        End Property

        Public Property BrickHeight As Double
            Get
                Return m_brickHt
            End Get
            Set(value As Double)
                m_brickHt = value
            End Set
        End Property

        Public Property BrickLength As Double
            Get
                Return m_brickLength
            End Get
            Set(value As Double)
                m_brickLength = value
            End Set
        End Property

        Public Property FootingThickness As Double
            Get
                Return m_footThick
            End Get
            Set(value As Double)
                m_footThick = value
            End Set
        End Property
        Public Property TopStepHeight As Double
            Get
                Return m_topStepHeight
            End Get
            Set(value As Double)
                m_topStepHeight = value
            End Set
        End Property

        Public Property CoverInches As Double
            Get
                Return Math.Round(m_footCover * 12, 3)
            End Get
            Set(value As Double)
                m_footCover = Math.Round(value / 12, 3)
            End Set
        End Property

        Public Property TopStepInches As Double
            Set(value As Double)
                m_topStepHeight = Math.Round(value / 12, 3)
            End Set
            Get
                Return Math.Round(m_topStepHeight * 12, 3)
            End Get
        End Property

        Public Property FootingThicknessInches As Double
            Set(value As Double)
                m_footThick = Math.Round(value / 12, 3)
            End Set
            Get
                Return Math.Round(m_footThick * 12, 3)
            End Get
        End Property
        Public Property BrickLenInches As Double
            Set(value As Double)
                m_brickLength = Math.Round(value / 12, 3)
            End Set
            Get
                Return Math.Round(m_brickLength * 12, 3)
            End Get
        End Property
        Public Property BrickHtInches As Double
            Set(value As Double)
                m_brickHt = Math.Round(value / 12, 3)
            End Set
            Get
                Return Math.Round(m_brickHt * 12, 3)
            End Get
        End Property

        Public ReadOnly Property ProfBrickHt As Double
            Get
                Return m_brickHt * m_VertFactor
            End Get
        End Property

        Public ReadOnly Property ProfFootCover As Double
            Get
                Return m_footCover * m_VertFactor
            End Get
        End Property

        Public ReadOnly Property ProfFootThickness As Double
            Get
                Return m_footThick * m_VertFactor
            End Get
        End Property

        Public ReadOnly Property ProfFootStep As Double
            Get
                Return m_vertFootStep * m_VertFactor
            End Get
        End Property

        Public ReadOnly Property ProfTopStepHt As Double
            Get
                Return m_topStepHeight * m_VertFactor
            End Get
        End Property

    End Class

End Namespace

