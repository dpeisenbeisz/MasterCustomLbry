Imports System.Math

Public Class AngleObj : Inherits CollectionBase

    Private ReadOnly m_decDegs As Double
    Private ReadOnly m_radians As Double
    Private ReadOnly m_azimuth As Double
    Private ReadOnly m_decAz As Double
    Private ReadOnly m_radBase As Double
    Private ReadOnly m_decBase As Double

    Public Sub New()
        MyBase.New
    End Sub
    Public Sub New(meas As Double, isRadians As Boolean, Optional baseAngle As Double = 0)
        'true = radians, false = decimal degrees

        If isRadians Then
            'baseAngle = (baseAngle + (3 * PI / 2)) Mod (2 * PI)
            m_azimuth = CDec(meas + baseAngle) Mod CDec(2 * PI)
            If m_azimuth < 0 Then m_azimuth += (2 * PI)
            If m_azimuth >= (2 * PI) Then m_azimuth -= (2 * PI)
            m_radians = CDec(meas) Mod CDec(2 * PI)
            If m_radians < 0 Then m_radians += (2 * PI)
            If m_radians >= (2 * PI) Then m_radians -= (2 * PI)
            m_decAz = Deg(m_azimuth)
            m_decDegs = Deg(meas)
            m_radBase = baseAngle
            m_decBase = Deg(baseAngle)
        Else
            'baseAngle = (baseAngle + 270) Mod 360
            m_decAz = CDec(meas + baseAngle) Mod 360
            If m_decAz < 0 Then m_decAz += 360
            If m_decAz >= 360 Then m_decAz -= 360
            m_decDegs = CDec(meas) Mod 360
            If m_decDegs < 0 Then m_decDegs += 360
            If m_decDegs >= 360 Then m_decDegs -= 360
            m_decBase = baseAngle
            m_azimuth = Rad(m_decAz)
            m_radians = Rad(m_decDegs)
            m_radBase = Rad(baseAngle)
        End If

    End Sub

    Public Sub New(degs As Long, mins As Long, secs As Double, Optional baseAngle As Double = 0)

        Dim tmpArr() As Double = {degs, mins, secs}
        Dim dms As Double = DMStoDEC(tmpArr)

        m_decDegs = CDec(dms) Mod 360
        If m_decDegs < 0 Then m_decDegs += 360
        If m_decDegs >= 360 Then m_decDegs -= 360
        m_decAz = CDec(dms + baseAngle) Mod 360
        If m_decAz < 0 Then m_decAz += 360
        If m_decAz >= 360 Then m_decAz -= 360
        m_decBase = baseAngle
        m_radBase = Rad(m_decBase)
        m_radians = Rad(m_decDegs)
        m_azimuth = Rad(m_decAz)

    End Sub

    Public Sub New(dMS() As Double, Optional baseAngle As Double = 0)

        Dim decDegs As Double = DMStoDEC(dMS)

        m_decDegs = CDec(decDegs) Mod 360
        If m_decDegs < 0 Then m_decDegs += 360
        If m_decDegs >= 360 Then m_decDegs -= 360
        m_decAz = CDec(decDegs + baseAngle) Mod 360
        If m_decAz < 0 Then m_decAz += 360
        If m_decAz >= 360 Then m_decAz -= 360
        m_decBase = baseAngle
        m_radBase = Rad(m_decBase)
        m_radians = Rad(m_decDegs)
        m_azimuth = Rad(m_decAz)
    End Sub

    Public Sub New(brng As String, Optional baseAngle As Double = 0)

        Dim Card As String = Left(brng, 1)
        Dim direct As String = Right(brng, 1)
        Dim NewStr As String = Mid(brng, 2, Len(brng) - 2)

        If Not Card = "N" Then
            If Not Card = "S" Then GoTo ErrSub
        End If

        If Not direct = "E" Then
            If Not direct = "W" Then GoTo ErrSub
        End If

        Dim split1() As String
        If InStr(1, NewStr, Chr(176), CompareMethod.Text) > 0 Then
            split1 = Split(NewStr, Chr(176), 2, CompareMethod.Text)
        ElseIf InStr(1, NewStr, "d", CompareMethod.Text) > 0 Then
            split1 = Split(NewStr, "d", 2, CompareMethod.Text)
        Else
            GoTo ErrSub
        End If

        Dim ang As Integer = CInt(split1(0))
        If ang > 90 Or ang < 0 Then GoTo ErrSub

        Dim minsSecs As String
        minsSecs = split1(1)
        Dim min As Integer

        Dim split2() As String
        If InStr(1, minsSecs, "'") > 0 Then
            split2 = Split(minsSecs, "'")
            min = CInt(split2(0))
        Else
            GoTo ErrSub
        End If

        'Dim secdbl As Double
        Dim sec As Double

        If Right(split2(1), 1) = Chr(34) Then
            Dim StrSec As String = Left(split2(1), Len(split2(1)) - 1)
            sec = CDbl(StrSec)
        Else
            sec = CDbl(split2(1))
        End If

        Dim dmsTmp() As Double = {ang, min, sec}

        'Dim decAng As Double = ang + mindbl + secdbl
        Dim decAz As Double = DMStoDEC(dmsTmp)

        Dim tempAz As Double
        'Dim az As Double

        Select Case Card
            Case Is = "N"
                If direct = "W" Then
                    tempAz = decAz
                ElseIf direct = "E" Then
                    tempAz = 360 - decAz
                End If

            Case Is = "S"
                If direct = "E" Then
                    tempAz = decAz + 180
                ElseIf direct = "W" Then
                    tempAz = 180 - decAz
                End If
        End Select

        Dim tempAng As Double = (tempAz - baseAngle)
        If tempAng < 0 Then tempAng += 360
        If tempAng >= 360 Then tempAng -= 360

        m_decAz = tempAz
        m_decDegs = tempAng
        m_decBase = baseAngle
        m_radians = Rad(m_decDegs)
        m_azimuth = Rad(m_decAz)
        m_radBase = Rad(m_decBase)
        Exit Sub
ErrSub:
        Throw New Exception("Surveyor's Angle not formatted correctly.")
    End Sub
    Public Function AddAngle(a2 As AngleObj) As AngleObj
        'Dim newAng As Double = (m_radians + a2.Radians) Mod (2 * PI)
        Dim newAng As Double = (m_decAz + a2.DecAzimuth) Mod 360
        If Abs(newAng) < 10 ^ -10 Then newAng = 0
        Dim a3 As New AngleObj(newAng, False, m_decBase)
        Return a3
    End Function

    Public Function SubtractAngle(a2 As AngleObj) As AngleObj
        Dim newAng As Double = (m_decAz - a2.DecAzimuth) Mod 360
        If Abs(newAng) < 10 ^ -10 Then newAng = 0
        Dim a3 As New AngleObj(newAng, False, m_decBase)
        Return a3
    End Function
    Public ReadOnly Property Degrees() As Long
        Get
            Dim tempDegs As Double = m_decDegs Mod 360
            Return tempDegs - (tempDegs Mod 1)
        End Get
    End Property
    Public Function AddAngles(angles() As AngleObj) As AngleObj
        'Dim newAng As Double = (m_radians + a2.Radians) Mod (2 * PI)
        Dim newAng As Double = m_decAz
        For i As Integer = 0 To UBound(angles)
            Dim a2 As AngleObj = angles(i)
            newAng += a2.DecAzimuth
        Next

        Dim sumAng As Double = CDec(newAng) Mod 360

        'Dim newAng As Double = (m_decAz + a2.DecAzimuth) Mod 360
        If Abs(sumAng) < 10 ^ -10 Then sumAng = 0
        Dim a3 As New AngleObj(sumAng, False, m_decBase)
        Return a3
    End Function
    Public ReadOnly Property Surveyors(Optional decimalPlaces As Integer = 2) As String
        Get
            Dim angl As Double = m_decAz Mod 360
            'Dim minDbl As Double = (angl Mod 1) * 60
            'Dim secdbl As Double = (minDbl Mod 1) * 60
            'Dim seclong As Double = Round(secdbl, 2)
            Dim card As String
            Dim direct As String
            Dim degs As String
            Dim mins As String
            Dim secs As String

            Dim adjAngl As Double

            If angl >= 0 And angl <= 90 Then
                card = "N"
                direct = "W"
                adjAngl = angl
            ElseIf angl >= 270 And angl < 360 Then
                card = "N"
                direct = "E"
                adjAngl = 360 - angl
            ElseIf angl > 90 And angl < 180 Then
                card = "S"
                direct = "W"
                adjAngl = 180 - angl
            ElseIf angl >= 180 And angl < 270 Then
                card = "S"
                direct = "E"
                adjAngl = angl - 180
            Else
                Throw New Exception("Error in angle string")
            End If

            'Dim deglong As Long = angl - (angl Mod 1)
            Dim angleParts() As Double = DECtoDMS(adjAngl)
            Dim ang As Integer = angleParts(0)
            Dim minLong As Integer = angleParts(1)
            Dim secDbl As Double = angleParts(2)

            If secDbl >= 60 Then
                minLong += 1
                secDbl -= 60
                If minLong >= 60 Then
                    minLong -= 60
                    ang += 1
                End If
            End If

            'Dim anglong As Long = adjAngl - (adjAngl Mod 1)
            'minDbl = (adjAngl Mod 1) * 60
            'Dim minlong As Long = minDbl - (minDbl Mod 1)
            'secdbl = (minDbl Mod 1) * 60

            degs = ang.ToString
            'minlong = angleParts(1)
            If minLong < 10 Then
                mins = "0" & minLong.ToString
            Else
                mins = minLong.ToString
            End If

            'secdbl = angleParts(2)
            If secDbl < 10 Then
                secs = "0" & Round(secDbl, decimalPlaces).ToString
            Else
                secs = Round(secDbl, decimalPlaces).ToString
            End If

            Return card & degs & ChrW(176) & mins & "'" & secs & ChrW(34) & direct

        End Get
    End Property
    Public ReadOnly Property Minutes() As Long
        Get
            Dim angParts() As Double = DECtoDMS(m_decDegs)
            Return CLng(angParts(1))
        End Get
    End Property

    Public ReadOnly Property Seconds(Optional decimalPlaces As Integer = 2) As Double
        Get
            Dim angParts() As Double = DECtoDMS(m_decDegs)
            Return Round(angParts(2), decimalPlaces)
        End Get
    End Property

    Public ReadOnly Property RadAzimuth As Double
        Get
            Return m_azimuth
        End Get
    End Property

    Public ReadOnly Property DecAzimuth As Double
        Get
            Return m_decAz
        End Get
    End Property
    Public ReadOnly Property BaseAngle As Double
        Get
            Return m_decBase
        End Get
    End Property

    Public ReadOnly Property RadBase As Double
        Get
            Return m_radBase
        End Get
    End Property
    Public ReadOnly Property DecimalDegrees() As Double
        Get
            Return m_decDegs
        End Get
    End Property

    Public ReadOnly Property Radians() As Double
        Get
            Return m_radians
        End Get
    End Property

    Public ReadOnly Property DMS(Optional decimalPlaces As Integer = 2) As String
        Get
            'Dim tempAng As Double = (m_decDegs Mod 360)
            'Dim degLong As Long = tempAng - (tempAng Mod 1)
            'Dim minDbl As Double = (tempAng Mod 1) * 60
            'Dim minlong As Long = minDbl - ((tempAng Mod 1) * 60) Mod 1
            'Dim secdbl As Double = (((tempAng Mod 1) * 60) Mod 1) * 60

            Dim angleParts() As Double = DECtoDMS(m_decDegs)

            Dim mins As String
            Dim secs As String

            Dim deglong As Integer = angleParts(0)

            Dim minlong As Integer = angleParts(1)
            If minlong < 10 Then
                mins = "0" & minlong.ToString
            Else
                mins = minlong.ToString
            End If

            Dim secdbl As Double = angleParts(2)
            If secdbl < 10 Then
                secs = "0" & Round(secdbl, decimalPlaces).ToString
            Else
                secs = Round(secdbl, decimalPlaces).ToString
            End If

            Return deglong.ToString & ChrW(176) & mins & "'" & secs & Chr(34)
        End Get
    End Property
    Public ReadOnly Property AutoCADdms(Optional decimalPlaces As Integer = 2) As String
        Get
            Dim angleParts() As Double = DECtoDMS(m_decDegs)

            Dim deglong As Integer = angleParts(0)
            Dim minlong As Integer = angleParts(1)
            Dim secdbl As Double = angleParts(2)

            Dim mins As String
            Dim secs As String
            If minlong < 10 Then
                mins = "0" & minlong.ToString
            Else
                mins = minlong.ToString
            End If

            If secdbl < 10 Then
                secs = "0" & Round(secdbl, decimalPlaces).ToString
            Else
                secs = Round(secdbl, decimalPlaces).ToString
            End If
            Return deglong.ToString & "d" & mins & "'" & secs & Chr(34)
        End Get
    End Property
    Private Function Rad(degs As Double) As Double
        Return degs * PI / 180
    End Function
    Private Function Deg(rads As Double) As Double
        Return rads * 180 / PI
    End Function
    Private Function DMStoDEC(dms() As Double) As Double
        If UBound(dms) = 2 Then
            Dim tempAng As Double = CDec(dms(0)) Mod 360
            Dim minDbl As Double = dms(1) / 60
            Dim secdbl As Double = dms(2) / 3600

            If secdbl >= 60 Then
                secdbl -= 60
                minDbl += 1
                If minDbl >= 60 Then
                    minDbl -= 60
                    tempAng += 1
                End If
            End If

            Return tempAng + minDbl + secdbl
        Else
            Throw New Exception("Angle array does not have 3 elements.")
        End If
    End Function
    Private Function DECtoDMS(tempAng As Double) As Double()
        'convert to decimals to prevent floating-point modulo error
        Dim decAng As Decimal = CDec(tempAng) Mod 360
        If decAng < 0 Then decAng += 360
        Dim degLong As Long = decAng - (decAng Mod 1)
        Dim minDbl As Decimal = (decAng Mod 1) * 60
        Dim secdbl As Double = (minDbl Mod 1) * 60

        If secdbl >= 60 Then
            secdbl -= 60
            minDbl += 1
            If minDbl >= 60 Then
                minDbl -= 60
                degLong += 1
            End If
        End If

        Dim minlong As Long = minDbl \ 1

        Dim returnArray() As Double = {degLong, minlong, secdbl}
        Return returnArray
    End Function
End Class
