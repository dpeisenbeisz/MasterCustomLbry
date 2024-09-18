'(C) David Eisenbeisz 2023
Public Class ArrowObj : Inherits CollectionBase
    Private m_A As Double
    Private m_B As Double
    Private m_C As Double
    Private m_D As Double
    Private m_E As Double
    Private m_F As Double
    Private m_R As Double
    'Private m_tHeight As Double
    Private m_ArrowAng As Double

    Public Sub New()
        MyBase.New
    End Sub

    'arrow types:
    '1 = single line horizontal vertical or diagonal arrow  use parameter C as necksize
    '2 = double line horizontal arrow  use parameter D as necksize
    '3 = double line vertical or diagonal arrow  use parameter C as necksize
    '4 = Vertical down arrow

    Public Sub New(arrType As Integer, sizeValue As Double, UseTextHeight As Boolean, Arrowangle As Double)

        'm_tHeight = sizeValue
        m_ArrowAng = Arrowangle

        If arrType = 2 Then
            If UseTextHeight Then
                SetDataType2(sizeValue)
            Else
                SetByNeckSizeType2(sizeValue)
            End If
        ElseIf arrType = 1 Then
            If UseTextHeight Then
                SetDataType1(sizeValue)
            Else
                SetByNeckSizeType1(sizeValue)
            End If
        ElseIf arrType = 3 Then
            If UseTextHeight Then
                SetDataType3(sizeValue)
            Else
                SetByNeckSizeType3(sizeValue)
            End If
        ElseIf arrType = 4 Then
            SetDataType4(sizeValue)
        End If
    End Sub


    Public ReadOnly Property ArrowAng As Double
        Get
            Return m_ArrowAng
        End Get
    End Property
    Public ReadOnly Property ParamA As Double
        Get
            Return m_A
        End Get
    End Property

    Public ReadOnly Property ParamB As Double
        Get
            Return m_B
        End Get
    End Property

    Public ReadOnly Property ParamC As Double
        Get
            Return m_C
        End Get
    End Property
    Public ReadOnly Property ParamD As Double
        Get
            Return m_D
        End Get
    End Property

    Public ReadOnly Property ParamE As Double
        Get
            Return m_E
        End Get
    End Property

    Public ReadOnly Property ParamF As Double
        Get
            Return m_F
        End Get
    End Property

    Public ReadOnly Property ParamR As Double
        Get
            Return m_R
        End Get
    End Property


    Public Sub SetDataType2(lettersize As Double)

        Select Case lettersize
            Case Is = 4
                m_A = 7.125
                m_B = 4.125
                m_C = 1 + 9 / 16
                m_D = 1.9375
                m_E = 7 / 16
                m_F = 6.375
                m_R = 5 / 16
            Case Is = 5
                m_A = 9
                m_B = 5 + 3 / 16
                m_C = 1 + 15 / 16
                m_D = 2 + 7 / 16
                m_E = 9 / 16
                m_F = 8 + 1 / 16
                m_R = 3 / 8
            Case Is = 6
                m_A = 10 + 11 / 16
                m_B = 6 + 3 / 16
                m_C = 2 + 5 / 16
                m_D = 2 + 15 / 16
                m_E = 5 / 8
                m_F = 9 + 9 / 16
                m_R = 1 / 2
            Case Is = 8
                m_A = 14.25
                m_B = 8.25
                m_C = 3.125
                m_D = 3.875
                m_E = 0.875
                m_F = 12.75
                m_R = 0.625
            Case Is = 10.67
                m_A = 18.75
                m_B = 10.875
                m_C = 3.75
                m_D = 5
                m_E = 1 + 5 / 16
                m_F = 17.25
                m_R = 13 / 16
            Case Is = 10
                m_A = 23 + 13 / 16
                m_B = 13 + 13 / 16
                m_C = 4.5
                m_D = 6
                m_E = 1.5
                m_F = 20.25
                m_R = 0.75
            Case Is = 12
                m_A = 23 + 13 / 16
                m_B = 13 + 13 / 16
                m_C = 4.5
                m_D = 6
                m_E = 1.5
                m_F = 20.25
                m_R = 0.75
            Case Is = 15
                m_A = 28.5
                m_B = 16.5
                m_C = 5 + 3 / 8
                m_D = 7.125
                m_E = 1.75
                m_F = 25
                m_R = 1
            Case Is = 16
                m_A = 28.5
                m_B = 16.5
                m_C = 5 + 3 / 8
                m_D = 7.125
                m_E = 1.75
                m_F = 25
                m_R = 1
            Case Else
                Call SetByNeckSizeType2(lettersize * 0.4)
        End Select
    End Sub

    Private Sub SetByNeckSizeType2(pC As Double)

        m_A = 4.56 * pC
        m_B = 2.64 * pC
        m_C = pC
        m_D = 1.24 * pC
        m_E = 0.28 * pC
        m_F = 4.08 * pC
        m_R = 0.2 * pC

    End Sub

    Private Sub SetDataType1(lettersize As Double)

        Select Case lettersize
            Case Is = 4
                m_A = 5 + 5 / 8
                m_B = 3 + 5 / 8
                m_C = 1 + 9 / 16
                m_D = 1 + 15 / 16
                m_E = 7 / 16
                m_F = 6 + 3 / 8
                m_R = 5 / 16
            Case Is = 5
                m_A = 7 + 1 / 16
                m_B = 4 + 9 / 16
                m_C = 1 + 15 / 16
                m_D = 2 + 7 / 16
                m_E = 9 / 16
                m_F = 8 + 1 / 16
                m_R = 3 / 8
            Case Is = 6
                m_A = 8 + 7 / 16
                m_B = 5 + 7 / 16
                m_C = 2 + 5 / 16
                m_D = 2 + 15 / 16
                m_E = 5 / 8
                m_F = 9 + 9 / 16
                m_R = 0.5
            Case Is = 8
                m_A = 11 + 1 / 4
                m_B = 7 + 1 / 4
                m_C = 3.125
                m_D = 3.875
                m_E = 0.875
                m_F = 12.75
                m_R = 0.625
            Case Is = 10.67
                m_A = 14.25
                m_B = 9 + 13 / 16
                m_C = 3 + 3 / 8
                m_D = 4.5
                m_E = 1 + 5 / 16
                m_F = 17.25
                m_R = 3 / 4
            Case Is = 10
                m_A = 14
                m_B = 9
                m_C = 4
                m_D = 5
                m_E = 7 / 8
                m_F = 16
                m_R = 7 / 8
            Case Is = 12
                m_A = 17.5
                m_B = 11.75
                m_C = 4 + 3 / 8
                m_D = 5 + 5 / 8
                m_E = 1.5
                m_F = 20.25
                m_R = 7 / 8
            Case Is = 15
                m_A = 21 + 7 / 8
                m_B = 14.25
                m_C = 5
                m_D = 6.75
                m_E = 1.75
                m_F = 25
                m_R = 1
            Case Else
                Call SetByNeckSizeType1(lettersize * 0.4)
        End Select

    End Sub


    Private Sub SetByNeckSizeType1(pD As Double)

        m_A = 2.8 * pD
        m_B = 1.8 * pD
        m_C = 0.8 * pD
        m_D = pD
        m_E = 0.175 * pD
        m_F = 3.2 * pD
        m_R = 0.175 * pD

    End Sub

    Public Sub SetDataType3(lettersize As Double)

        Select Case lettersize
            Case Is = 4
                m_A = 5 + 5 / 8
                m_B = 4 + 3 / 8
                m_C = 1 + 9 / 16
                m_D = 1 + 15 / 16
                m_E = 7 / 16
                m_F = 9.125
                m_R = 5 / 16
            Case Is = 5
                m_A = 7 + 1 / 16
                m_B = 5.5
                m_C = 1 + 15 / 16
                m_D = 2 + 7 / 16
                m_E = 9 / 16
                m_F = 11.5
                m_R = 3 / 8
            Case Is = 6
                m_A = 8 + 7 / 16
                m_B = 6 + 9 / 16
                m_C = 2 + 5 / 16
                m_D = 2 + 15 / 16
                m_E = 5 / 8
                m_F = 13 + 11 / 16
                m_R = 0.5
            Case Is = 8
                m_A = 11.25
                m_B = 8.75
                m_C = 3.125
                m_D = 3.875
                m_E = 0.875
                m_F = 18.25
                m_R = 0.625
            Case Is = 10.67
                m_A = 15.125
                m_B = 11 + 9 / 16
                m_C = 3.75
                m_D = 5
                m_E = 1 + 5 / 16
                m_F = 24.25
                m_R = 13 / 16
            Case Is = 10
                m_A = 18.25
                m_B = 14
                m_C = 4.5
                m_D = 6
                m_E = 1.5
                m_F = 29.25
                m_R = 0.75
            Case Is = 12
                m_A = 18.25
                m_B = 14
                m_C = 4.5
                m_D = 6
                m_E = 1.5
                m_F = 29.25
                m_R = 0.75
            Case Is = 15
                m_A = 22.25
                m_B = 17
                m_C = 5 + 3 / 8
                m_D = 7.125
                m_E = 1.75
                m_F = 35 + 5 / 8
                m_R = 1
            Case Is = 16
                m_A = 22.25
                m_B = 17
                m_C = 5 + 3 / 8
                m_D = 7.125
                m_E = 1.75
                m_F = 35 + 5 / 8
                m_R = 1
            Case Else
                Call SetByNeckSizeType3(lettersize * 0.483333)
        End Select
    End Sub

    Private Sub SetByNeckSizeType3(pc As Double)

        m_A = 3.6 * pc
        m_B = 2.8 * pc
        m_C = pc
        m_D = 1.2 * pc
        m_E = 0.28 * pc
        m_F = 5.84 * pc
        m_R = 0.2 * pc

    End Sub
    Public Sub SetDataType4(ArrowWidth As Integer)

        Select Case ArrowWidth
            Case Is = 24
                m_A = ArrowWidth
                m_B = 12
                m_C = 5
                m_D = 2
                m_E = 16.5
                m_R = 3 / 4
            Case Is = 32
                m_A = ArrowWidth
                m_B = 16
                m_C = 6.5
                m_D = 3
                m_E = 22
                m_R = 1
        End Select
    End Sub

End Class
