Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.Geometry
Imports System.Math
Imports MasterCustomLibrary.AcCommon

Namespace Walls
    Public Module WallCommands

        Friend m_Wallwidth As Double
        Friend m_WDDLength As Double
        Friend m_AcLayerName As String
        Private m_dirPt As Point3d
        Private m_RWallData As RetainingWall

        <CommandMethod("WALLSTPS")>
        Public Sub WallSteps()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = CurDwg.Editor

            Dim peo As New PromptEntityOptions(vbLf & "Pick polyline for the maximum top of wall elevation.")

            Dim egId As ObjectId

            With peo
                .AllowNone = False
            End With

            Dim per As PromptEntityResult = ed.GetEntity(peo)

            If per.Status = PromptStatus.OK Then
                egId = per.ObjectId
            Else
                Exit Sub
            End If

            Dim ppo As New PromptPointOptions(vbLf & "Pick point for the top of wall at start point")
            With ppo
                .AllowNone = False
            End With

            Dim pb1 As Point3d
            Dim ppr As PromptPointResult = ed.GetPoint(ppo)

            If ppr.Status = PromptStatus.OK Then
                pb1 = ppr.Value
            Else
                Exit Sub
            End If

            Dim ppo2 As New PromptPointOptions(vbLf & "Pick approximate location where the steps will end.")

            Dim dirPt As Point3d
            Dim ppr2 As PromptPointResult = ed.GetPoint(ppo2)

            If ppr2.Status = PromptStatus.OK Then
                dirPt = ppr2.Value
            Else
                Exit Sub
            End If

            Dim pltSide As Integer

            If dirPt.X > pb1.X Then
                pltSide = 1
            Else
                pltSide = -1
            End If

            Dim endX As Double = dirPt.X

            Dim pdo2 As New PromptDistanceOptions(vbLf & "Enter or pick width of a single brick")

            With pdo2
                .AllowArbitraryInput = True
                .AllowZero = False
                .AllowNegative = False
                .AllowNone = False
            End With

            Dim pdr2 As PromptDoubleResult = ed.GetDistance(pdo2)

            Dim blen As Double

            If pdr2.Status = PromptStatus.OK Then
                blen = pdr2.Value
            Else
                Exit Sub
            End If

            Dim pdo3 As New PromptDistanceOptions(vbLf & "Enter or pick the height of a single brick")
            Dim pdr3 As PromptDoubleResult = ed.GetDistance(pdo3)

            Dim bht As Double

            If pdr3.Status = PromptStatus.OK Then
                bht = pdr3.Value
            Else
                Exit Sub
            End If

            Dim startX As Double = pb1.X
            Dim curX As Double = pb1.X
            Dim curY As Double = pb1.Y
            Dim lasty As Double = curY

            Dim refPln As New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

            Dim footPts As New Point2dCollection
            footPts.Add(pb1.Convert2d(refPln))

            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction

                Try

                    Dim curSpc As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForRead)
                    Dim baseline As DBObject = acTrans.GetObject(egId, OpenMode.ForRead)
                    Dim bl As Polyline = TryCast(baseline, Polyline)
                    'Dim tempObj As Object = GetVertexCoords(topid)

                    If bl Is Nothing OrElse bl.Closed = True Then
                        MessageBox.Show("Error.  Picked entity cannot be used for top of wall or path is closed.")
                        acTrans.Abort()
                        Exit Sub
                    End If

TryAgain:
                    If bl.Elevation <> 0 Then bl.Elevation = 0
                    Dim topfoot As New Point2dCollection
                    Dim brk As Integer = 0
                    Dim failsafe As Integer = 0
                    Dim lastx As Double = curX

                    If pltSide = 1 Then  'reference moves from left to right
                        Do
                            Dim negpt2d As New Point2d(curX, pb1.Y - 1000)
                            Dim pospt2d As New Point2d(curX, pb1.Y + 1000)

                            Using tLine As New Polyline

                                With tLine
                                    .AddVertexAt(0, negpt2d, 0, 0, 0)
                                    .AddVertexAt(1, pospt2d, 0, 0, 0)
                                End With

                                Dim pts2d As New Point3dCollection

                                bl.IntersectWith(tLine, Intersect.OnBothOperands, pts2d, IntPtr.Zero, IntPtr.Zero)

                                'check to see if refy is within 1 brick height of cury
                                If pts2d IsNot Nothing AndAlso pts2d.Count > 0 Then

                                    Debug.Print(pts2d(0).ToString)
                                    Dim refY As Double = pts2d(0).Y
                                    Dim minY As Double = refY - bht
                                    Dim elevDiff As Double = refY - curY

                                    If curY > minY Then  'cury is above minimum
                                        If curY <= refY Then  'cury is below reference line
                                            lastx = curX
                                            curX = startX + (brk * 0.5 * blen)
                                            'if next point is beyond end of bl then create one final point
                                            If curX > endX Then
                                                Dim tps1 As New Point2d(curX, curY)
                                                footPts.Add(tps1)
                                            End If
                                            brk += 1
                                        Else  'cury is above reference line
                                            'lastx = curX - blen
                                            Dim fSafe As Integer = 0
                                            lasty = curY
                                            Do
                                                curY -= bht
                                                If curY < refY Then Exit Do
                                                fSafe += 1
                                            Loop While fSafe < 50
                                            Dim tps1 As New Point2d(curX, lasty)
                                            Dim tps2 As New Point2d(curX + 0.02, curY)
                                            footPts.Add(tps1)
                                            footPts.Add(tps2)
                                        End If
                                    Else    'cury is below minimum
                                        Dim fSafe As Integer = 0
                                        lasty = curY
                                        Do
                                            curY += bht
                                            If curY > minY Then
                                                Exit Do
                                            End If
                                            fSafe += 1
                                        Loop While fSafe < 50
                                        Dim tps1 As New Point2d(curX, lasty)
                                        Dim tps2 As New Point2d(curX + 0.02, curY)
                                        footPts.Add(tps1)
                                        footPts.Add(tps2)
                                    End If

                                End If

                            End Using

                            'lastx = curX
                            failsafe += 1
                            If curX > endX Then Exit Do

                        Loop While failsafe < 10000

                        Debug.Print(failsafe)

                    Else  'reference moves from right to left
                        Do
                            Dim negpt2d As New Point2d(curX, pb1.Y - 1000)
                            Dim pospt2d As New Point2d(curX, pb1.Y + 1000)

                            Using tLine As New Polyline

                                With tLine
                                    .AddVertexAt(0, negpt2d, 0, 0, 0)
                                    .AddVertexAt(1, pospt2d, 0, 0, 0)
                                End With

                                'If tLine.Elevation <> 0 Then tLine.Elevation = 0

                                Dim pts2d As New Point3dCollection

                                bl.IntersectWith(tLine, Intersect.OnBothOperands, pts2d, IntPtr.Zero, IntPtr.Zero)

                                'If pts2d.Count <= 0 Then
                                '    tLine.Dispose()
                                '    GoTo SkipPt
                                'End If

                                'check to see if refy is within 1 brick height of cury
                                If pts2d IsNot Nothing AndAlso pts2d.Count > 0 Then

                                    'check to see if refy is within 1 brick height of cury
                                    Dim refY As Double = pts2d(0).Y
                                    Dim minY As Double = refY - bht
                                    Dim elevDiff As Double = refY - curY

                                    If curY > minY Then  'cury is above minimum
                                        If curY <= refY Then  'cury is below reference line
                                            'Dim lastx As Double = curX
                                            lastx = curX
                                            curX = startX - (brk * 0.5 * blen)
                                            'if next point is beyond end of bl then create one final point
                                            If curX < endX Then
                                                Dim tps1 As New Point2d(curX, curY)
                                                footPts.Add(tps1)
                                            End If
                                            brk += 1
                                        Else  'cury is above reference line
                                            Dim fSafe As Integer = 0
                                            lasty = curY
                                            Do
                                                curY -= bht
                                                If curY < refY Then Exit Do
                                                fSafe += 1
                                            Loop While fSafe < 50
                                            Dim tps1 As New Point2d(curX, lasty)
                                            Dim tps2 As New Point2d(curX - 0.02, curY)
                                            footPts.Add(tps1)
                                            footPts.Add(tps2)
                                        End If
                                    Else    'cury is below minimum
                                        Dim fSafe As Integer = 0
                                        lasty = curY
                                        Do
                                            curY += bht
                                            If curY > minY Then Exit Do
                                            fSafe += 1
                                        Loop While fSafe < 50
                                        Dim tps1 As New Point2d(curX, lasty)
                                        Dim tps2 As New Point2d(curX - 0.02, curY)
                                        footPts.Add(tps1)
                                        footPts.Add(tps2)
                                    End If
                                End If
SkipPt:
                            End Using

                            failsafe += 1
                            If curX < endX Then Exit Do

                        Loop While failsafe < 5000

                        Debug.Print(failsafe)

                    End If

                    bl.Dispose()

                    'Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    'Dim curspace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)
                    Using pl As New Polyline
                        For k = 0 To footPts.Count - 1
                            pl.AddVertexAt(k, footPts(k), 0, 0, 0)
                        Next
                        If Not curSpc.IsWriteEnabled Then curSpc.UpgradeOpen()
                        curSpc.AppendEntity(pl)
                        acTrans.AddNewlyCreatedDBObject(pl, True)
                    End Using

                    acTrans.Commit()

                Catch ex As Exception
                    MessageBox.Show(vbLf & "Error in footing layout command.  Check data and try again." & vbLf & vbLf & ex.Message)
                    acTrans.Abort()
                    Exit Sub
                End Try

            End Using

        End Sub

        <CommandMethod("WALLPROFILES")>
        Public Sub WallProfs()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = CurDwg.Editor

            'Dim hasConcreteBase As Boolean = False
            Dim topConcId As ObjectId
            Dim pb1 As Point3d
            'Dim pbConc As Point3d
            'Dim unadjFootStep As Double
            Dim rwData As New RetainingWall
            Dim useCurrent As Boolean = False

            If m_RWallData IsNot Nothing Then
                useCurrent = YesNoQuery(vbLf & "Do you want to use the previous retaining wall parameters?")
                If useCurrent Then rwData = m_RWallData
            End If

            If Not useCurrent Then

                Dim pdoVFact As New PromptDistanceOptions(vbLf & "Enter the vertical exaggeration factor.")

                With pdoVFact
                    If m_RWallData IsNot Nothing Then .DefaultValue = m_RWallData.VertFactor Else .DefaultValue = 1
                    .AllowArbitraryInput = True
                    .AllowZero = False
                    .AllowNegative = False
                    .AllowNone = False
                End With

                Dim pdrVFact As PromptDoubleResult = ed.GetDistance(pdoVFact)

                'Dim vFact As Double

                If pdrVFact.Status = PromptStatus.OK Then
                    rwData.VertFactor = pdrVFact.Value
                    'vFact = pdrVFact.Value
                Else
                    Exit Sub
                End If

                Dim concreteQ As PromptResult = YesNoResult("Does the wall have a reinforced concrete stem at its base?")

                If concreteQ.Status = PromptStatus.OK Then
                    Dim pRes As String = concreteQ.StringResult
                    If pRes = "Yes" Then
                        'hasConcreteBase = True
                        rwData.HasConcreteBase = True

                        Dim footopts As New PromptDistanceOptions(vbLf & "Pick or enter the vertical distance for each step in the top of the footing")
                        With footopts
                            If m_RWallData IsNot Nothing Then .DefaultValue = m_RWallData.FootingStepHeight Else .DefaultValue = 0.67
                            .AllowNegative = False
                            .AllowNone = False
                            .AllowZero = False
                        End With

                        Dim footres As PromptDoubleResult = ed.GetDistance(footopts)
                        If footres.Status = PromptStatus.OK Then
                            rwData.FootingStepHeight = footres.Value
                            'unadjFootStep = footres.Value
                        Else
                            Exit Sub
                        End If

                    End If
                End If

                Dim pdoCover As New PromptDistanceOptions(vbLf & "Enter or pick minimum distance between top of footing and ground elevation")

                With pdoCover
                    If m_RWallData IsNot Nothing Then .DefaultValue = m_RWallData.FootingCover Else .DefaultValue = 1
                    .AllowArbitraryInput = True
                    .AllowZero = False
                    .AllowNegative = False
                    .AllowNone = False
                End With

                Dim pdrCover As PromptDoubleResult = ed.GetDistance(pdoCover)

                'Dim cover As Double

                If pdrCover.Status = PromptStatus.OK Then
                    'cover = pdrCover.Value
                    rwData.FootingCover = pdrCover.Value
                Else
                    Exit Sub
                End If


                Dim pdo2 As New PromptDistanceOptions(vbLf & "Enter or pick length of a single brick in feet")

                With pdo2
                    If m_RWallData IsNot Nothing Then .DefaultValue = m_RWallData.BrickLength Else .DefaultValue = 1.333
                    .AllowArbitraryInput = True
                    .AllowZero = False
                    .AllowNegative = False
                    .AllowNone = False
                End With

                Dim pdr2 As PromptDoubleResult = ed.GetDistance(pdo2)

                'Dim blen As Double

                If pdr2.Status = PromptStatus.OK Then
                    'blen = pdr2.Value
                    rwData.BrickLength = pdr2.Value
                Else
                    Exit Sub
                End If

                Dim pdo3 As New PromptDistanceOptions(vbLf & "Enter or pick the true height of a single brick in feet")

                With pdo3
                    If m_RWallData IsNot Nothing Then .DefaultValue = m_RWallData.BrickHeight Else .DefaultValue = 0.667
                    .AllowArbitraryInput = True
                    .AllowZero = False
                    .AllowNegative = False
                    .AllowNone = False
                End With

                Dim pdr3 As PromptDoubleResult = ed.GetDistance(pdo3)

                'Dim bht As Double
                'Dim trueBht As Double

                If pdr3.Status = PromptStatus.OK Then
                    rwData.BrickHeight = pdr3.Value
                    If Not rwData.HasConcreteBase Then rwData.FootingStepHeight = pdr3.Value

                    'trueBht = pdr3.Value
                    'bht = trueBht * vFact
                Else
                    Exit Sub
                End If


                Dim pdoFoot As New PromptDistanceOptions(vbLf & "Enter or pick the minimum footing thickness.")

                With pdoFoot
                    If m_RWallData IsNot Nothing Then .DefaultValue = m_RWallData.FootingThickness
                    .AllowArbitraryInput = True
                    .AllowZero = False
                    .AllowNegative = False
                    .AllowNone = False
                End With

                Dim pdrFoot As PromptDoubleResult = ed.GetDistance(pdoFoot)

                'Dim footThk As Double

                If pdrFoot.Status = PromptStatus.OK Then
                    'footThk = pdrFoot.Value
                    rwData.FootingThickness = pdrFoot.Value
                Else
                    Exit Sub
                End If

                'Dim footstep As Double

                ''''''''''''''''''''''''''''''Pick Data

            End If

            If rwData.HasConcreteBase Then

                Dim pcw As New PromptEntityOptions(vbLf & "Pick polyline for the maximum top of concrete stem.")

                With pcw
                    .AllowNone = False
                    .SetRejectMessage(vbLf & "Must be a polyline entity.")
                    .AddAllowedClass(GetType(Polyline), True)
                End With

                Dim pcwr As PromptEntityResult = ed.GetEntity(pcw)

                If pcwr.Status = PromptStatus.OK Then
                    topConcId = pcwr.ObjectId
                Else
                    Exit Sub
                End If

            End If


            Dim peo As New PromptEntityOptions(vbLf & "Pick polyline for the maximum top of wall elevation.")

            Dim twId As ObjectId

            With peo
                .SetRejectMessage(vbLf & "Must be a polyline entity.")
                .AddAllowedClass(GetType(Polyline), True)
                .AllowNone = False
            End With

            Dim per As PromptEntityResult = ed.GetEntity(peo)

            If per.Status = PromptStatus.OK Then
                twId = per.ObjectId
            Else
                Exit Sub
            End If

            'Dim hasTopWall As Boolean = False
            'If topConcId <> twId Then hasTopWall = True

            Dim peo2 As New PromptEntityOptions(vbLf & "Pick polyline for the ground elevation on the toe side of the wall")
            With peo2
                .SetRejectMessage(vbLf & "Must be a polyline entity.")
                .AddAllowedClass(GetType(Polyline), True)
                .AllowNone = False
            End With

            Dim egID As ObjectId
            Dim per2 As PromptEntityResult = ed.GetEntity(peo2)
            If per.Status = PromptStatus.OK Then
                egID = per2.ObjectId
            Else
                Exit Sub
            End If


            'Dim ppo As New PromptPointOptions(vbLf & "Pick point for the top of concrete stem at start point")
            '    With ppo
            '        .AllowNone = False
            '    End With

            '    Dim ppr As PromptPointResult = ed.GetPoint(ppo)

            '    If ppr.Status = PromptStatus.OK Then
            '        pbConc = ppr.Value
            '    Else
            '        Exit Sub
            '    End If

            Dim ppo As New PromptPointOptions(vbLf & "Pick point for the top of wall at start point")

            With ppo
                .AllowNone = False
            End With

                Dim ppr As PromptPointResult = ed.GetPoint(ppo)

                If ppr.Status = PromptStatus.OK Then
                    pb1 = ppr.Value
                Else
                    Exit Sub
                End If

                Dim ppo2 As New PromptPointOptions(vbLf & "Pick approximate location where the wall will end.")

                Dim dirPt As Point3d
                Dim ppr2 As PromptPointResult = ed.GetPoint(ppo2)

                If ppr2.Status = PromptStatus.OK Then
                    dirPt = ppr2.Value
                Else
                    Exit Sub
                End If

            Dim pltSide As Integer


            If dirPt.X > pb1.X Then
                pltSide = 1
            Else
                pltSide = -1
            End If


            Dim endX As Double = dirPt.X

            Dim bht As Double = rwData.ProfBrickHt
            Dim blen As Double = rwData.BrickLength
            'Dim vfact As Double = rwData.VertFactor
            Dim cover As Double = rwData.ProfFootCover
            Dim footstep As Double = rwData.ProfFootCover
            Dim footthk As Double = rwData.ProfFootThickness

            Try

                ''''''''''''''''''''''''''''''''''Start Geometry

                Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction

                    If rwData.HasConcreteBase Then
                        Dim tConcStart As Point2d = TopConcPfile(pb1, topConcId, rwData, endX, pltSide, acTrans)
                        If Not tConcStart = Point2d.Origin Then
                            pb1 = New Point3d(tConcStart.X, tConcStart.Y, 0)
                        Else
                            Exit Sub
                        End If
                    End If

                    'loop variables
                    Dim startX As Double = pb1.X
                    Dim curX As Double = pb1.X
                    Dim curY As Double = pb1.Y
                    Dim lasty As Double = curY

                    Dim refPln As New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

                    Dim topPts As New Point2dCollection
                    topPts.Add(pb1.Convert2d(refPln))

                    Dim curSpc As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForRead)

                    'get the top of the wall and EG (toe elevation)
                    Dim baseline As DBObject = acTrans.GetObject(twId, OpenMode.ForRead)
                    Dim bl As Polyline = TryCast(baseline, Polyline)
                    'Dim egLine As Polyline = TryCast(acTrans.GetObject(egID, OpenMode.ForRead), Polyline)

                    'temporary coordinates for vertical line
                    Dim temppos2d As New Point2d(curX, pb1.Y - 1000)
                    Dim tempneg2d As New Point2d(curX, pb1.Y + 1000)

                    If bl Is Nothing OrElse bl.Closed = True Then
                        MessageBox.Show("Error.  Picked entity cannot be used for top of wall or polyline is closed.")
                        acTrans.Abort()
                        Exit Sub
                    End If
TryAgain:
                    If bl.Elevation <> 0 Then bl.Elevation = 0
                    'If egLine.Elevation <> 0 Then egLine.Elevation = 0
                    'If fsLine.Elevation <> 0 Then fsLine.Elevation = 0

                    Dim brk As Integer = 0
                    Dim failsafe As Integer = 0
                    Dim lastx As Double = curX

                    If pltSide = 1 Then  'reference moves from left to right
                        Do
                            Dim negpt2d As New Point2d(curX, pb1.Y - 1000)
                            Dim pospt2d As New Point2d(curX, pb1.Y + 1000)

                            Using tLine As New Polyline

                                With tLine
                                    .AddVertexAt(0, negpt2d, 0, 0, 0)
                                    .AddVertexAt(1, pospt2d, 0, 0, 0)
                                End With

                                Dim pts2d As New Point3dCollection

                                'get the intersection point of the temporary vertical line and the top of the wall
                                bl.IntersectWith(tLine, Intersect.OnBothOperands, pts2d, IntPtr.Zero, IntPtr.Zero)

                                'check to see if refy is within 1 brick height of cury
                                If pts2d IsNot Nothing AndAlso pts2d.Count > 0 Then

                                    'make minY the top of the wall (refY) - 1 brick height
                                    Dim refY As Double = pts2d(0).Y
                                    Dim minY As Double = refY - bht

                                    If curY > minY Then  'cury is above minimum 
                                        If curY <= refY Then  'cury is below reference line so this is a good point
                                            'keep track of x value for later and move down the wall
                                            lastx = curX
                                            curX = startX + (brk * 0.5 * blen)
                                            'if next point is beyond end of bl then create one final point
                                            If curX > endX Then
                                                Dim tps1 As New Point2d(curX, curY)
                                                topPts.Add(tps1)
                                            End If
                                            brk += 1
                                        Else  'cury is above reference line
                                            Dim fSafe As Integer = 0
                                            'store current y before adjusting
                                            lasty = curY
                                            'adjust height until it is within 1 brick below top of wall
                                            Do
                                                curY -= bht
                                                If curY < refY Then
                                                    Exit Do
                                                End If
                                                fSafe += 1
                                            Loop While fSafe < 50
                                            'set a point for the bottom of the previous step and the top of the current step
                                            Dim tps1 As New Point2d(curX, lasty)
                                            Dim tps2 As New Point2d(curX + 0.02, curY)
                                            topPts.Add(tps1)
                                            topPts.Add(tps2)
                                        End If
                                    Else    'cury is below minimum
                                        Dim fSafe As Integer = 0
                                        'store current y value and loop until y is within 1 brick of the top of wall
                                        lasty = curY
                                        Do
                                            curY += bht
                                            If curY > minY Then
                                                Exit Do
                                            End If
                                            fSafe += 1
                                        Loop While fSafe < 50
                                        Dim tps1 As New Point2d(curX, lasty)
                                        Dim tps2 As New Point2d(curX + 0.02, curY)
                                        topPts.Add(tps1)
                                        topPts.Add(tps2)
                                    End If

                                Else
                                    MessageBox.Show("Error.  Polyline for max wall height is not long enough.")
                                    Exit Sub
                                End If

                            End Using

                            failsafe += 1
                            If curX > endX Then Exit Do

                            'limit the loop to 5000 brick lengths
                        Loop While failsafe < 5000

                        Debug.Print(failsafe)

                    Else  'reference moves from right to left
                        Do
                            Dim negpt2d As New Point2d(curX, pb1.Y - 1000)
                            Dim pospt2d As New Point2d(curX, pb1.Y + 1000)

                            Using tLine As New Polyline

                                With tLine
                                    .AddVertexAt(0, negpt2d, 0, 0, 0)
                                    .AddVertexAt(1, pospt2d, 0, 0, 0)
                                End With

                                'If tLine.Elevation <> 0 Then tLine.Elevation = 0

                                Dim pts2d As New Point3dCollection

                                bl.IntersectWith(tLine, Intersect.OnBothOperands, pts2d, IntPtr.Zero, IntPtr.Zero)

                                'If pts2d.Count <= 0 Then
                                '    tLine.Dispose()
                                '    GoTo SkipPt
                                'End If

                                'check to see if refy is within 1 brick height of cury
                                If pts2d IsNot Nothing AndAlso pts2d.Count > 0 Then

                                    'check to see if refy is within 1 brick height of cury
                                    Dim refY As Double = pts2d(0).Y
                                    Dim minY As Double = refY - bht
                                    Dim elevDiff As Double = refY - curY

                                    If curY > minY Then  'cury is above minimum
                                        If curY <= refY Then  'cury is below reference line
                                            'Dim lastx As Double = curX
                                            lastx = curX
                                            curX = startX - (brk * 0.5 * blen)
                                            'if next point is beyond end of bl then create one final point
                                            If curX < endX Then
                                                Dim tps1 As New Point2d(curX, curY)
                                                topPts.Add(tps1)
                                            End If
                                            brk += 1
                                        Else  'cury is above reference line
                                            Dim fSafe As Integer = 0
                                            lasty = curY
                                            Do
                                                curY -= bht
                                                If curY < refY Then
                                                    'curY -= bht
                                                    Exit Do
                                                End If
                                                fSafe += 1
                                            Loop While fSafe < 50
                                            Dim tps1 As New Point2d(curX, lasty)
                                            Dim tps2 As New Point2d(curX - 0.02, curY)
                                            topPts.Add(tps1)
                                            topPts.Add(tps2)
                                        End If
                                    Else    'cury is below minimum
                                        Dim fSafe As Integer = 0
                                        lasty = curY
                                        Do
                                            curY += bht
                                            If curY > minY Then
                                                Exit Do
                                            End If
                                            fSafe += 1
                                        Loop While fSafe < 50
                                        Dim tps1 As New Point2d(curX, lasty)
                                        Dim tps2 As New Point2d(curX - 0.02, curY)
                                        topPts.Add(tps1)
                                        topPts.Add(tps2)
                                    End If

                                Else
                                    MessageBox.Show("Error.  Polyline for max wall height is not long enough.")
                                    Exit Sub
                                End If
SkipPt:
                            End Using

                            failsafe += 1
                            If curX < endX Then Exit Do

                        Loop While failsafe < 5000

                        Debug.Print(failsafe)

                    End If

                    bl.Dispose()

                    Dim tWallId As ObjectId

                    'Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    'Dim curspace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)
                    Using pl As New Polyline
                        For k = 0 To topPts.Count - 1
                            pl.AddVertexAt(k, topPts(k), 0, 0, 0)
                        Next
                        If Not curSpc.IsWriteEnabled Then curSpc.UpgradeOpen()
                        tWallId = curSpc.AppendEntity(pl)
                        acTrans.AddNewlyCreatedDBObject(pl, True)
                    End Using

                    'Dim adjCover As Double = vFact * cover

                    Dim topFtId As ObjectId = TopFooting(egID, pb1, dirPt, pltSide, rwData, acTrans)

                    If Not topFtId = ObjectId.Null Then
                        Dim tF As Polyline = acTrans.GetObject(topFtId, OpenMode.ForRead)
                        Dim sp As Point3d = tF.StartPoint
                        Dim GoodFooting As Boolean = FootBottom(topFtId, sp, rwData, pltSide)
                    Else
                        Exit Sub
                    End If

                    acTrans.Commit()
                    m_RWallData = rwData

                End Using

            Catch ex As Exception
                MessageBox.Show(vbLf & "Error in WallProfiles command.  Check data and try again." & vbLf & vbLf & ex.Message)
                Exit Sub
            End Try


        End Sub

        Public Function TopConcPfile(pb1 As Point3d, topConcId As ObjectId, rwData As RetainingWall, endX As Double, pltSide As Double, acTrans As Transaction) As Point2d

            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = CurDwg.Editor

            'loop variables
            Dim startX As Double = pb1.X
            Dim curX As Double = pb1.X
            Dim curY As Double = pb1.Y
            Dim lasty As Double = curY

            Dim refPln As New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

            Try

                Dim topPts As New Point2dCollection
                topPts.Add(pb1.Convert2d(refPln))

                Dim curSpc As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForRead)

                'get the top of the wall and EG (toe elevation)
                Dim baseline As DBObject = acTrans.GetObject(topConcId, OpenMode.ForRead)

                If TypeOf baseline IsNot Polyline Then
                    Throw New InvalidCastException
                    Return Point2d.Origin
                    Exit Function
                End If

                Dim bl As Polyline = TryCast(baseline, Polyline)

                If bl Is Nothing OrElse bl.Closed = True Then
                    MessageBox.Show("Error.  Picked entity cannot be used for top of concrete or polyline is closed.")
                    Return Point2d.Origin
                    Exit Function
                End If
TryAgain:
                If bl.Elevation <> 0 Then bl.Elevation = 0

                Dim brk As Integer = 0
                Dim failsafe As Integer = 0
                Dim lastx As Double = curX

                Dim bht As Double = rwData.ProfBrickHt
                Dim blen As Double = rwData.BrickLength

                If pltSide = 1 Then  'reference moves from left to right
                    Do
                        Dim negpt2d As New Point2d(curX, pb1.Y - 1000)
                        Dim pospt2d As New Point2d(curX, pb1.Y + 1000)

                        Using tLine As New Polyline

                            With tLine
                                .AddVertexAt(0, negpt2d, 0, 0, 0)
                                .AddVertexAt(1, pospt2d, 0, 0, 0)
                            End With

                            Dim pts2d As New Point3dCollection

                            'get the intersection point of the temporary vertical line and the top of the wall
                            bl.IntersectWith(tLine, Intersect.OnBothOperands, pts2d, IntPtr.Zero, IntPtr.Zero)

                            'check to see if refy is within 1 brick height of cury
                            If pts2d IsNot Nothing AndAlso pts2d.Count > 0 Then

                                'make minY the top of the wall (refY) - 1 brick height
                                Dim refY As Double = pts2d(0).Y
                                Dim minY As Double = refY - bht

                                If curY > minY Then  'cury is above minimum 
                                    If curY <= refY Then  'cury is below reference line so this is a good point
                                        'keep track of x value for later and move down the wall
                                        lastx = curX
                                        curX = startX + (brk * 0.5 * blen)
                                        'if next point is beyond end of bl then create one final point
                                        If curX > endX Then
                                            Dim tps1 As New Point2d(curX, curY)
                                            topPts.Add(tps1)
                                        End If
                                        brk += 1
                                    Else  'cury is above reference line
                                        Dim fSafe As Integer = 0
                                        'store current y before adjusting
                                        lasty = curY
                                        'adjust height until it is within 1 brick below top of wall
                                        Dim foundY As Boolean = False
                                        Do
                                            curY -= bht
                                            If curY < refY Then
                                                foundY = True
                                                Exit Do
                                            End If
                                            fSafe += 1
                                        Loop While fSafe < 50

                                        If foundY Then
                                            'set a point for the bottom of the previous step and the top of the current step
                                            Dim tps1 As New Point2d(curX, lasty)
                                            Dim tps2 As New Point2d(curX + 0.02, curY)
                                            topPts.Add(tps1)
                                            topPts.Add(tps2)
                                        Else
                                            MessageBox.Show("Error.  point is outside of wall ends.")
                                            Return Point2d.Origin
                                            Exit Function
                                        End If
                                    End If
                                Else    'cury is below minimum
                                    Dim fSafe As Integer = 0
                                    'store current y value and loop until y is within 1 brick of the top of wall
                                    lasty = curY
                                    Dim foundY As Boolean = False
                                    Do
                                        curY += bht
                                        If curY > minY Then
                                            foundY = True
                                            Exit Do
                                        End If
                                        fSafe += 1
                                    Loop While fSafe < 50
                                    If foundY Then
                                        'set a point for the bottom of the previous step and the top of the current step
                                        Dim tps1 As New Point2d(curX, lasty)
                                        Dim tps2 As New Point2d(curX + 0.02, curY)
                                        topPts.Add(tps1)
                                        topPts.Add(tps2)
                                    Else
                                        MessageBox.Show("Error.  point is outside of wall ends.")
                                        Return Point2d.Origin
                                        Exit Function
                                    End If
                                End If
                            Else
                                MessageBox.Show("Error.  Polyline for max concrete stem height is not long enough.")
                                Return Point2d.Origin
                                Exit Function
                            End If

                        End Using

                        failsafe += 1
                        If curX > endX Then Exit Do

                        'limit the loop to 5000 brick lengths
                    Loop While failsafe < 5000

                    Debug.Print(failsafe)

                Else  'reference moves from right to left
                    Do
                        Dim negpt2d As New Point2d(curX, pb1.Y - 1000)
                        Dim pospt2d As New Point2d(curX, pb1.Y + 1000)

                        Using tLine As New Polyline

                            With tLine
                                .AddVertexAt(0, negpt2d, 0, 0, 0)
                                .AddVertexAt(1, pospt2d, 0, 0, 0)
                            End With

                            'If tLine.Elevation <> 0 Then tLine.Elevation = 0

                            Dim pts2d As New Point3dCollection

                            bl.IntersectWith(tLine, Intersect.OnBothOperands, pts2d, IntPtr.Zero, IntPtr.Zero)

                            'If pts2d.Count <= 0 Then
                            '    tLine.Dispose()
                            '    GoTo SkipPt
                            'End If

                            'check to see if refy is within 1 brick height of cury
                            If pts2d IsNot Nothing AndAlso pts2d.Count > 0 Then

                                'check to see if refy is within 1 brick height of cury
                                Dim refY As Double = pts2d(0).Y
                                Dim minY As Double = refY - bht
                                Dim elevDiff As Double = refY - curY

                                If curY > minY Then  'cury is above minimum
                                    If curY <= refY Then  'cury is below reference line
                                        'Dim lastx As Double = curX
                                        lastx = curX
                                        curX = startX - (brk * 0.5 * blen)
                                        'if next point is beyond end of bl then create one final point
                                        If curX < endX Then
                                            Dim tps1 As New Point2d(curX, curY)
                                            topPts.Add(tps1)
                                        End If
                                        brk += 1
                                    Else  'cury is above reference line
                                        Dim fSafe As Integer = 0
                                        Dim foundY As Boolean = False
                                        lasty = curY
                                        Do
                                            curY -= bht
                                            If curY < refY Then
                                                foundY = True
                                                Exit Do
                                            End If
                                            fSafe += 1
                                        Loop While fSafe < 50

                                        If foundY Then
                                            Dim tps1 As New Point2d(curX, lasty)
                                            Dim tps2 As New Point2d(curX - 0.02, curY)
                                            topPts.Add(tps1)
                                            topPts.Add(tps2)
                                        Else
                                            MessageBox.Show("Error.  point is outside of wall ends.")
                                            Return Point2d.Origin
                                            Exit Function
                                        End If

                                    End If
                                Else    'cury is below minimum
                                    Dim fSafe As Integer = 0
                                    Dim foundY As Boolean = False
                                    lasty = curY
                                    Do
                                        curY += bht
                                        If curY > minY Then
                                            foundY = True
                                            Exit Do
                                        End If
                                        fSafe += 1
                                    Loop While fSafe < 50

                                    If foundY Then
                                        Dim tps1 As New Point2d(curX, lasty)
                                        Dim tps2 As New Point2d(curX - 0.02, curY)
                                        topPts.Add(tps1)
                                        topPts.Add(tps2)
                                    Else
                                        MessageBox.Show("Error.  point is outside of wall ends.")
                                        Return Point2d.Origin
                                        Exit Function
                                    End If

                                End If
                            Else
                                MessageBox.Show("Error.  Polyline for max concrete stem height is not long enough.")
                                Return Point2d.Origin
                                Exit Function
                            End If
SkipPt:
                        End Using

                        failsafe += 1
                        If curX < endX Then Exit Do

                    Loop While failsafe < 5000

                    Debug.Print(failsafe)

                End If

                bl.Dispose()

                Dim tWallId As ObjectId

                'Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                'Dim curspace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)
                Using pl As New Polyline
                    For k = 0 To topPts.Count - 1
                        pl.AddVertexAt(k, topPts(k), 0, 0, 0)
                    Next
                    If Not curSpc.IsWriteEnabled Then curSpc.UpgradeOpen()
                    tWallId = curSpc.AppendEntity(pl)
                    acTrans.AddNewlyCreatedDBObject(pl, True)
                End Using

                Return topPts(0)

            Catch ex As Exception
                MessageBox.Show(vbLf & "Error in TopConcPfile function.  Check data and try again." & vbLf & vbLf & ex.Message)
                Return Point2d.Origin
                Exit Function
            End Try

        End Function


        Friend Function TopConcStem(egId As ObjectId, twid As ObjectId, tcID As ObjectId, pb1 As Point3d, pltside As Integer, endX As Double, vfact As Double, bht As Double, blen As Double, acTrans As Transaction) As Point2d
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = CurDwg.Editor

            Try

                'loop variables
                Dim startX As Double = pb1.X
                Dim curX As Double = pb1.X
                Dim curY As Double = pb1.Y
                Dim lasty As Double = curY

                Dim refPln As New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

                Dim tconcPts As New Point2dCollection
                tconcPts.Add(pb1.Convert2d(refPln))

                Dim curSpc As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForRead)

                'get the top of the concrete stem, top of wall, and EG (toe elevation)
                Dim baseline As DBObject = acTrans.GetObject(tcID, OpenMode.ForRead)
                Dim bl As Polyline = TryCast(baseline, Polyline)
                Dim twLine As Polyline = TryCast(acTrans.GetObject(twid, OpenMode.ForRead), Polyline)

                'temporary coordinates for vertical line
                Dim temppos2d As New Point2d(curX, pb1.Y - 1000)
                Dim tempneg2d As New Point2d(curX, pb1.Y + 1000)

                If bl Is Nothing OrElse bl.Closed = True Then
                    MessageBox.Show("Error.  Picked entity cannot be used or polyline is closed.")
                    Return Nothing
                    Exit Function
                End If
TryAgain:
                If bl.Elevation <> 0 Then bl.Elevation = 0

                Dim twPts As New Point3dCollection
                'Dim tfootPts As New Point3dCollection

                Dim brk As Integer = 0
                Dim failsafe As Integer = 0
                Dim lastx As Double = curX
                Dim startTopPt As Point2d

                If pltside = 1 Then  'reference moves from left to right
                    Dim firstPt As Boolean = True
                    Dim ptsTopConc As New Point3dCollection

                    Do
                        Dim negpt2d As New Point2d(curX, pb1.Y - 1000)
                        Dim pospt2d As New Point2d(curX, pb1.Y + 1000)

                        Using tempLine As New Polyline

                            With tempLine
                                .AddVertexAt(0, negpt2d, 0, 0, 0)
                                .AddVertexAt(1, pospt2d, 0, 0, 0)
                            End With

                            'get the intersection point of the temporary vertical line and the top of the concrete stem
                            tempLine.IntersectWith(bl, Intersect.OnBothOperands, ptsTopConc, IntPtr.Zero, IntPtr.Zero)

                            'check to see if refy is within 1 brick height of cury
                            If ptsTopConc IsNot Nothing AndAlso ptsTopConc.Count > 0 Then

                                'make minY the top of the concrete (refY) - 1 brick height
                                Dim reftopY As Double = ptsTopConc(0).Y
                                Dim mintopConcY As Double = reftopY - bht

                                If curY > mintopConcY Then  'cury is above minY
                                    If curY <= reftopY Then  'cury is below reference line so this is a good point
                                        'keep track of x value for later
                                        lastx = curX
                                        curX = startX + (brk * 0.5 * blen)
                                        'if next point is beyond end of bl then create one final point

                                        If curX > endX Then
                                            Dim lastpt As New Point2d(curX, curY)
                                            tconcPts.Add(lastpt)
                                        End If
                                        brk += 1
                                    Else  'cury is above reference line
                                        Dim fSafe As Integer = 0
                                        'store current y before adjusting
                                        lasty = curY
                                        'adjust height until it is within 1 brick below top of wall
                                        Do
                                            curY -= bht
                                            If curY < reftopY Then
                                                Exit Do
                                            End If
                                            fSafe += 1
                                        Loop While fSafe < 50
                                        'set a point for the bottom of the previous step and the top of the current step
                                        Dim tps1 As New Point2d(curX, lasty)
                                        Dim tps2 As New Point2d(curX + 0.02, curY)
                                        tconcPts.Add(tps1)
                                        tconcPts.Add(tps2)
                                    End If

                                Else    'cury is below minimum
                                    Dim fSafe As Integer = 0
                                    'store current y value and loop until y is within 1 brick of the top of wall
                                    lasty = curY
                                    Do
                                        curY += bht
                                        If curY > mintopConcY Then
                                            Exit Do
                                        End If
                                        fSafe += 1
                                    Loop While fSafe < 50
                                    Dim tps1 As New Point2d(curX, lasty)
                                    Dim tps2 As New Point2d(curX + 0.02, curY)
                                    tconcPts.Add(tps1)
                                    tconcPts.Add(tps2)
                                End If

                                ''store a point for top of wall...
                                'If firstPt Then
                                '    Dim ptsTopWall As New Point3dCollection
                                '    tempLine.IntersectWith(twLine, Intersect.OnBothOperands, ptsTopWall, IntPtr.Zero, IntPtr.Zero)

                                '    If ptsTopWall IsNot Nothing AndAlso ptsTopWall.Count > 0 Then
                                '        Dim curtwY As Double = curY
                                '        'make minY the top of the wall (refY) - 1 brick height

                                '        Dim reftopWallY As Double = ptsTopWall(0).Y
                                '        Dim mintopWallY As Double = reftopWallY - bht

                                '        'curtwy is always below minimum unless very short wall
                                '        Dim fSafe As Integer = 0
                                '        'store current y value and loop until y is within 1 brick of the top of wall
                                '        Dim lasttwy As Double = curtwY
                                '        Do
                                '            curtwY += bht
                                '            If curtwY > mintopWallY Then
                                '                Exit Do
                                '            End If
                                '            fSafe += 1
                                '        Loop While fSafe < 50

                                '        startTopPt = New Point2d(curX, curtwY)
                                '    End If
                                '    firstPt = False
                                'End If

                                failsafe += 1
                                If curX > endX Then Exit Do
                                'limit the loop to 5000 brick lengths
                            End If

                        End Using

                    Loop While failsafe < 5000
                    Debug.Print(failsafe)

                Else  'reference moves from right to left
                    Dim firstPt As Boolean = True
                    Dim ptsTopConc As New Point3dCollection

                    Do
                        Dim negpt2d As New Point2d(curX, pb1.Y - 1000)
                        Dim pospt2d As New Point2d(curX, pb1.Y + 1000)

                        Using tempLine As New Polyline

                            With tempLine
                                .AddVertexAt(0, negpt2d, 0, 0, 0)
                                .AddVertexAt(1, pospt2d, 0, 0, 0)
                            End With

                            'get the intersection point of the temporary vertical line and the top of the concrete stem
                            tempLine.IntersectWith(bl, Intersect.OnBothOperands, ptsTopConc, IntPtr.Zero, IntPtr.Zero)

                            'check to see if refy is within 1 brick height of cury
                            If ptsTopConc IsNot Nothing AndAlso ptsTopConc.Count > 0 Then

                                'make minY the top of the concrete (refY) - 1 brick height
                                Dim reftopY As Double = ptsTopConc(0).Y
                                Dim mintopConcY As Double = reftopY - bht

                                If curY > mintopConcY Then  'cury is above minY
                                    If curY <= reftopY Then  'cury is below reference line so this is a good point
                                        'keep track of x value for later
                                        lastx = curX
                                        curX = startX - (brk * 0.5 * blen)
                                        'if next point is beyond end of bl then create one final point

                                        If curX > endX Then
                                            Dim lastpt As New Point2d(curX, curY)
                                            tconcPts.Add(lastpt)
                                        End If
                                        brk += 1
                                    Else  'cury is above reference line
                                        Dim fSafe As Integer = 0
                                        'store current y before adjusting
                                        lasty = curY
                                        'adjust height until it is within 1 brick below top of wall
                                        Do
                                            curY -= bht
                                            If curY < reftopY Then
                                                Exit Do
                                            End If
                                            fSafe += 1
                                        Loop While fSafe < 50
                                        'set a point for the bottom of the previous step and the top of the current step
                                        Dim tps1 As New Point2d(curX, lasty)
                                        Dim tps2 As New Point2d(curX - 0.02, curY)
                                        tconcPts.Add(tps1)
                                        tconcPts.Add(tps2)
                                    End If

                                Else    'cury is below minimum
                                    Dim fSafe As Integer = 0
                                    'store current y value and loop until y is within 1 brick of the top of wall
                                    lasty = curY
                                    Do
                                        curY += bht
                                        If curY > mintopConcY Then
                                            Exit Do
                                        End If
                                        fSafe += 1
                                    Loop While fSafe < 50
                                    Dim tps1 As New Point2d(curX, lasty)
                                    Dim tps2 As New Point2d(curX - 0.02, curY)
                                    tconcPts.Add(tps1)
                                    tconcPts.Add(tps2)
                                End If

                                ''store a point for top of wall...
                                'If firstPt Then
                                '    Dim ptsTopWall As New Point3dCollection
                                '    tempLine.IntersectWith(twLine, Intersect.OnBothOperands, ptsTopWall, IntPtr.Zero, IntPtr.Zero)

                                '    If ptsTopWall IsNot Nothing AndAlso ptsTopWall.Count > 0 Then
                                '        Dim curtwY As Double = curY
                                '        'make minY the top of the wall (refY) - 1 brick height

                                '        Dim reftopWallY As Double = ptsTopWall(0).Y
                                '        Dim mintopWallY As Double = reftopWallY - bht

                                '        'curtwy is always below minimum unless very short wall
                                '        Dim fSafe As Integer = 0
                                '        'store current y value and loop until y is within 1 brick of the top of wall
                                '        Dim lasttwy As Double = curtwY
                                '        Do
                                '            curtwY += bht
                                '            If curtwY > mintopWallY Then
                                '                Exit Do
                                '            End If
                                '            fSafe += 1
                                '        Loop While fSafe < 50

                                '        startTopPt = New Point2d(curX, curtwY)
                                '    End If
                                '    firstPt = False
                                'End If

                                failsafe += 1
                                If curX < endX Then Exit Do
                                'limit the loop to 5000 brick lengths
                            End If

                        End Using

                    Loop While failsafe < 5000
                    Debug.Print(failsafe)

                End If

                bl.Dispose()

                Dim tConcid As ObjectId

                'Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                'Dim curspace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)
                Using pl As New Polyline
                    For k = 0 To tconcPts.Count - 1
                        pl.AddVertexAt(k, tconcPts(k), 0, 0, 0)
                    Next
                    If Not curSpc.IsWriteEnabled Then curSpc.UpgradeOpen()
                    tConcid = curSpc.AppendEntity(pl)
                    acTrans.AddNewlyCreatedDBObject(pl, True)
                End Using

                Return startTopPt

            Catch ex As Exception
                MessageBox.Show(vbLf & "Error in TopConcPfile function.  Check data and try again." & vbLf & vbLf & ex.Message)
                Return Point2d.Origin
                Exit Function
            End Try




        End Function

        Friend Function TopFooting(egId As ObjectId, twStart As Point3d, dirPt As Point3d, pltside As Integer, rwData As RetainingWall, acTrans As Transaction) As ObjectId
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = CurDwg.Editor

            Dim curX As Double = twStart.X
            Dim testPt1 As New Point2d(curX, twStart.Y - 1000)
            Dim testPt2 As New Point2d(curX, twStart.Y + 1000)

            Dim tempLine As New Polyline

            With tempLine
                .AddVertexAt(0, testPt1, 0, 0, 0)
                .AddVertexAt(1, testPt2, 0, 0, 0)
            End With

            If tempLine.Elevation <> 0 Then tempLine.Elevation = 0

            Dim startX As Double = twStart.X
            Dim endX As Double = dirPt.X
            Dim lastx As Double = curX

            Dim refPln As New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

            Dim curSpc As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForRead)
            Dim baseline As DBObject = acTrans.GetObject(egId, OpenMode.ForRead)
            Dim tbl As Polyline = TryCast(baseline, Polyline)
            'Dim tempObj As Object = GetVertexCoords(topid)

            If tbl Is Nothing Or tbl.Closed = True Then
                MessageBox.Show("Error.  Picked entity cannot be used for toe side ground elevation or polyline is closed.")
                Return Nothing
                Exit Function
            End If

            Dim cover As Double = rwData.ProfFootCover
            Dim bht As Double = rwData.ProfBrickHt
            Dim blen As Double = rwData.BrickLength

TryAgain:
            Dim ofPt As New Point3d(tbl.StartPoint.X, tbl.StartPoint.Y - cover, 0)
            Dim transVect As Vector3d = tbl.StartPoint.GetVectorTo(ofPt)

            Dim bl As Polyline = tbl.Clone
            bl.TransformBy(Matrix3d.Displacement(transVect))
            If bl.Elevation <> 0 Then bl.Elevation = 0

            If Not curSpc.IsWriteEnabled Then curSpc.UpgradeOpen()
            curSpc.AppendEntity(bl)
            acTrans.AddNewlyCreatedDBObject(bl, True)

            Dim tfAtStart As New Point3dCollection
            bl.IntersectWith(tempLine, Intersect.OnBothOperands, tfAtStart, IntPtr.Zero, IntPtr.Zero)

            Dim rY As Double

            If tfAtStart IsNot Nothing AndAlso tfAtStart.Count > 0 Then
                rY = twStart.Y
                Dim teststart As Point3d = tfAtStart(0)
                Do
                    rY -= bht
                Loop Until rY < teststart.Y
            Else
                ed.WriteMessage(vbLf & "Toe side ground line does not reach the start of the wall.  Lengthen the ground reference line and try again")
                Return ObjectId.Null
                Exit Function
            End If

            Dim pb1 As New Point3d(twStart.X, rY, 0)
            Dim curY As Double = rY
            Dim lasty As Double = curY

            Dim footPts As New Point2dCollection
            footPts.Add(pb1.Convert2d(refPln))

            tempLine.Dispose()
            tbl.Dispose()

            'Dim cont As Boolean = True
            Dim brk As Integer = 0
            Dim failsafe As Integer = 0

            If pltside > 0 Then  'reference moves from left to right
                Do
                    Dim negpt2d As New Point2d(curX, curY - 1000)
                    Dim pospt2d As New Point2d(curX, curY + 1000)

                    Using tLine As New Polyline

                        With tLine
                            .AddVertexAt(0, negpt2d, 0, 0, 0)
                            .AddVertexAt(1, pospt2d, 0, 0, 0)
                        End With

                        If tLine.Elevation <> 0 Then tLine.Elevation = 0

                        Dim pts2d As New Point3dCollection

                        bl.IntersectWith(tLine, Intersect.OnBothOperands, pts2d, IntPtr.Zero, IntPtr.Zero)

                        If pts2d.Count <= 0 Then
                            MessageBox.Show("Ground line on toe side is not long enough.  Check input data and try again.")
                            Exit Do
                        End If

                        'check to see if refy is within 1 brick height of cury
                        If pts2d IsNot Nothing AndAlso pts2d.Count > 0 Then

                            Debug.Print(pts2d(0).ToString)
                            Dim refY As Double = pts2d(0).Y
                            Dim minY As Double = refY - bht
                            Dim elevDiff As Double = refY - curY

                            If curY > minY Then  'cury is above minimum
                                If curY <= refY Then  'cury is below reference line
                                    'Dim lastx As Double = curX
                                    lastx = curX
                                    curX = startX + (brk * 0.5 * blen)
                                    'if next point is beyond end of bl then create one final point
                                    If curX > endX Then
                                        Dim tps1 As New Point2d(curX, curY)
                                        footPts.Add(tps1)
                                    End If
                                    brk += 1
                                Else  'cury is above reference line
                                    'lastx = curX - blen
                                    Dim fSafe As Integer = 0
                                    lasty = curY
                                    Do
                                        curY -= bht
                                        If curY < refY Then
                                            'curY -= bht
                                            Exit Do
                                        End If
                                        fSafe += 1
                                    Loop While fSafe < 50
                                    Dim tps1 As New Point2d(curX, lasty)
                                    Dim tps2 As New Point2d(curX + 0.02, curY)
                                    footPts.Add(tps1)
                                    footPts.Add(tps2)
                                End If
                            Else    'cury is below minimum
                                Dim fSafe As Integer = 0
                                lasty = curY
                                Do
                                    curY += bht
                                    If curY > minY Then
                                        Exit Do
                                    End If
                                    fSafe += 1
                                Loop While fSafe < 50
                                Dim tps1 As New Point2d(curX, lasty)
                                Dim tps2 As New Point2d(curX + 0.02, curY)
                                footPts.Add(tps1)
                                footPts.Add(tps2)
                            End If

                        End If

                    End Using

                    'lastx = curX
                    failsafe += 1
                    If curX > endX Then Exit Do

                Loop While failsafe < 5000

            Else  'reference moves from right to left
                Do
                    Dim negpt2d As New Point2d(curX, curY - 1000)
                    Dim pospt2d As New Point2d(curX, curY + 1000)

                    Using tLine As New Polyline

                        With tLine
                            .AddVertexAt(0, negpt2d, 0, 0, 0)
                            .AddVertexAt(1, pospt2d, 0, 0, 0)
                        End With

                        If tLine.Elevation <> 0 Then tLine.Elevation = 0

                        Dim pts2d As New Point3dCollection

                        bl.IntersectWith(tLine, Intersect.OnBothOperands, pts2d, IntPtr.Zero, IntPtr.Zero)

                        If pts2d.Count <= 0 Then
                            MessageBox.Show("Ground line on toe side is not long enough.  Check input data and try again.")
                            Exit Do
                        End If

                        'check to see if refy is within 1 brick height of cury
                        If pts2d IsNot Nothing AndAlso pts2d.Count > 0 Then

                            'check to see if refy is within 1 brick height of cury
                            Dim refY As Double = pts2d(0).Y
                            Dim minY As Double = refY - bht
                            Dim elevDiff As Double = refY - curY

                            If curY > minY Then  'cury is above minimum
                                If curY <= refY Then  'cury is below reference line
                                    'Dim lastx As Double = curX
                                    lastx = curX
                                    curX = startX - (brk * 0.5 * blen)
                                    'if next point is beyond end of bl then create one final point
                                    If curX < endX Then
                                        Dim tps1 As New Point2d(curX, curY)
                                        footPts.Add(tps1)
                                    End If
                                    brk += 1
                                Else  'cury is above reference line
                                    Dim fSafe As Integer = 0
                                    lasty = curY
                                    Do
                                        curY -= bht
                                        If curY < refY Then
                                            'curY -= bht
                                            Exit Do
                                        End If
                                        fSafe += 1
                                    Loop While fSafe < 50
                                    Dim tps1 As New Point2d(curX, lasty)
                                    Dim tps2 As New Point2d(curX - 0.02, curY)
                                    footPts.Add(tps1)
                                    footPts.Add(tps2)
                                End If
                            Else    'cury is below minimum
                                Dim fSafe As Integer = 0
                                lasty = curY
                                Do
                                    curY += bht
                                    If curY > minY Then
                                        Exit Do
                                    End If
                                    fSafe += 1
                                Loop While fSafe < 50
                                Dim tps1 As New Point2d(curX, lasty)
                                Dim tps2 As New Point2d(curX - 0.02, curY)
                                footPts.Add(tps1)
                                footPts.Add(tps2)
                            End If
                        End If
SkipPt:
                    End Using

                    failsafe += 1
                    If curX < endX Then Exit Do

                Loop While failsafe < 10000

            End If

            Dim tpFtID As ObjectId

            'Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
            'Dim curspace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)
            'Using acTrans2 As Transaction = DwgDB.TransactionManager.StartTransaction
            'Dim mSpace As BlockTableRecord = acTrans2.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)
            Dim pl As New Polyline
            For k = 0 To footPts.Count - 1
                pl.AddVertexAt(k, footPts(k), 0, 0, 0)
            Next
            tpFtID = curSpc.AppendEntity(pl)
            acTrans.AddNewlyCreatedDBObject(pl, True)

            Return tpFtID

        End Function

        <CommandMethod("FTLO")>
        Public Sub FootingLayout()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = CurDwg.Editor

            Dim peo As New PromptEntityOptions(vbLf & "Pick polyline for the ground surface elevation.")

            Dim egId As ObjectId

            With peo
                .AllowNone = False
            End With

            Dim per As PromptEntityResult = ed.GetEntity(peo)

            If per.Status = PromptStatus.OK Then
                egId = per.ObjectId
            Else
                Exit Sub
            End If

            Dim ppo As New PromptPointOptions(vbLf & "Pick point for the top of footing at start point")
            With ppo
                .AllowNone = False
            End With

            Dim pb1 As Point3d
            Dim ppr As PromptPointResult = ed.GetPoint(ppo)

            If ppr.Status = PromptStatus.OK Then
                pb1 = ppr.Value
            Else
                Exit Sub
            End If

            ppo.Message = vbLf & "Pick approximate location where the steps will end."

            Dim dirPt As Point3d
            ppr = ed.GetPoint(ppo)

            If ppr.Status = PromptStatus.OK Then
                dirPt = ppr.Value
            Else
                Exit Sub
            End If

            Dim pltSide As Integer

            If dirPt.X > pb1.X Then
                pltSide = 1
            Else
                pltSide = -1
            End If

            Dim pdoVFact As New PromptDistanceOptions(vbLf & "Enter the vertical exaggeration factor.")

            With pdoVFact
                .DefaultValue = 1
                .AllowArbitraryInput = True
                .AllowZero = False
                .AllowNegative = False
                .AllowNone = False
            End With

            Dim pdrVFact As PromptDoubleResult = ed.GetDistance(pdoVFact)

            Dim vFact As Double

            If pdrVFact.Status = PromptStatus.OK Then
                vFact = pdrVFact.Value

            Else
                Exit Sub
            End If

            Dim endX As Double = dirPt.X

            Dim pdo1 As New PromptDistanceOptions(vbLf & "Enter or pick minimum distance between top of footing and ground elevation adjusted for vertical scale.")

            With pdo1
                .AllowArbitraryInput = True
                .AllowZero = False
                .AllowNegative = False
                .AllowNone = False
            End With

            Dim pdr As PromptDoubleResult = ed.GetDistance(pdo1)

            Dim cover As Double

            If pdr.Status = PromptStatus.OK Then
                Dim cvr As Double = pdr.Value
                cover = cvr * vFact
            Else
                Exit Sub
            End If

            pdo1.Message = vbLf & "Enter or pick width of a single brick"

            With pdo1
                .AllowArbitraryInput = True
                .AllowZero = False
                .AllowNegative = False
                .AllowNone = False
            End With

            pdr = ed.GetDistance(pdo1)

            Dim blen As Double

            If pdr.Status = PromptStatus.OK Then
                blen = pdr.Value
            Else
                Exit Sub
            End If

            pdo1.Message = vbLf & "Enter of pick the true height of a single brick"
            pdr = ed.GetDistance(pdo1)

            Dim bht As Double

            If pdr.Status = PromptStatus.OK Then
                Dim tbht As Double = pdr.Value
                bht = tbht * vFact
            Else
                Exit Sub
            End If

            Dim startX As Double = pb1.X
            Dim curX As Double = pb1.X
            Dim curY As Double = pb1.Y
            Dim lastx As Double = curX
            Dim lasty As Double = curY

            Dim refPln As New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

            Dim footPts As New Point2dCollection
            footPts.Add(pb1.Convert2d(refPln))

            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction

                Try

                    Dim curSpc As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForRead)
                    Dim baseline As DBObject = acTrans.GetObject(egId, OpenMode.ForRead)
                    Dim tbl As Polyline = TryCast(baseline, Polyline)
                    'Dim tempObj As Object = GetVertexCoords(topid)

                    If tbl Is Nothing Or tbl.Closed = True Then
                        MessageBox.Show("Error.  Picked entity cannot be used for top of wall or path is closed.")
                        acTrans.Abort()
                        Exit Sub
                    End If

TryAgain:
                    'Dim spX As Double = tbl.StartPoint.X
                    'Dim epX As Double = tbl.EndPoint.X

                    'If pltSide < 0 Then
                    '    If spX < epX Then
                    '        If Not tbl.IsWriteEnabled Then tbl.UpgradeOpen()
                    '        tbl.ReverseCurve()
                    '        'tempObj = GetVertexCoords(bl.ObjectId)
                    '        GoTo TryAgain
                    '    End If
                    'Else
                    '    If spX > epX Then
                    '        If Not tbl.IsWriteEnabled Then tbl.UpgradeOpen()
                    '        tbl.ReverseCurve()
                    '        'tempObj = GetVertexCoords(bl.ObjectId)
                    '        GoTo TryAgain
                    '    End If
                    'End If

                    Dim ofPt As New Point3d(tbl.StartPoint.X, tbl.StartPoint.Y - cover, 0)
                    Dim transVect As Vector3d = tbl.StartPoint.GetVectorTo(ofPt)

                    Dim bl As Polyline = tbl.Clone
                    bl.TransformBy(Matrix3d.Displacement(transVect))
                    If bl.Elevation <> 0 Then bl.Elevation = 0

                    If Not curSpc.IsWriteEnabled Then curSpc.UpgradeOpen()
                    curSpc.AppendEntity(bl)
                    acTrans.AddNewlyCreatedDBObject(bl, True)

                    tbl.Dispose()

                    'Dim cont As Boolean = True
                    Dim brk As Integer = 0
                    Dim failsafe As Integer = 0

                    If pltSide > 0 Then  'reference moves from left to right
                        Do
                            Dim negpt2d As New Point2d(curX, pb1.Y - 1000)
                            Dim pospt2d As New Point2d(curX, pb1.Y + 1000)

                            Using tLine As New Polyline

                                With tLine
                                    .AddVertexAt(0, negpt2d, 0, 0, 0)
                                    .AddVertexAt(1, pospt2d, 0, 0, 0)
                                End With


                                If tLine.Elevation <> 0 Then tLine.Elevation = 0

                                Dim pts2d As New Point3dCollection

                                bl.IntersectWith(tLine, Intersect.OnBothOperands, pts2d, IntPtr.Zero, IntPtr.Zero)

                                If pts2d.Count <= 0 Then
                                    MessageBox.Show("No valid footings.  Check input data and try again.")
                                    Exit Sub
                                End If

                                'check to see if refy is within 1 brick height of cury
                                If pts2d IsNot Nothing AndAlso pts2d.Count > 0 Then

                                    Debug.Print(pts2d(0).ToString)
                                    Dim refY As Double = pts2d(0).Y
                                    Dim minY As Double = refY - bht
                                    Dim elevDiff As Double = refY - curY

                                    If curY > minY Then  'cury is above minimum
                                        If curY <= refY Then  'cury is below reference line
                                            'Dim lastx As Double = curX
                                            lastx = curX
                                            curX = startX + (brk * 0.5 * blen)
                                            'if next point is beyond end of bl then create one final point
                                            If curX > endX Then
                                                Dim tps1 As New Point2d(curX, curY)
                                                footPts.Add(tps1)
                                            End If
                                            brk += 1
                                        Else  'cury is above reference line
                                            'lastx = curX - blen
                                            Dim fSafe As Integer = 0
                                            lasty = curY
                                            Do
                                                curY -= bht
                                                If curY < refY Then
                                                    'curY -= bht
                                                    Exit Do
                                                End If
                                                fSafe += 1
                                            Loop While fSafe < 50
                                            Dim tps1 As New Point2d(curX, lasty)
                                            Dim tps2 As New Point2d(curX + 0.02, curY)
                                            footPts.Add(tps1)
                                            footPts.Add(tps2)
                                        End If
                                    Else    'cury is below minimum
                                        Dim fSafe As Integer = 0
                                        lasty = curY
                                        Do
                                            curY += bht
                                            If curY > minY Then
                                                Exit Do
                                            End If
                                            fSafe += 1
                                        Loop While fSafe < 50
                                        Dim tps1 As New Point2d(curX, lasty)
                                        Dim tps2 As New Point2d(curX + 0.02, curY)
                                        footPts.Add(tps1)
                                        footPts.Add(tps2)
                                    End If

                                End If

                            End Using

                            'lastx = curX
                            failsafe += 1
                            If curX > endX Then Exit Do

                        Loop While failsafe < 5000

                    Else  'reference moves from right to left
                        Do
                            Dim negpt2d As New Point2d(curX, pb1.Y - 1000)
                            Dim pospt2d As New Point2d(curX, pb1.Y + 1000)

                            Using tLine As New Polyline

                                With tLine
                                    .AddVertexAt(0, negpt2d, 0, 0, 0)
                                    .AddVertexAt(1, pospt2d, 0, 0, 0)
                                End With

                                If tLine.Elevation <> 0 Then tLine.Elevation = 0

                                Dim pts2d As New Point3dCollection

                                bl.IntersectWith(tLine, Intersect.OnBothOperands, pts2d, IntPtr.Zero, IntPtr.Zero)

                                'check to see if refy is within 1 brick height of cury
                                If pts2d IsNot Nothing AndAlso pts2d.Count > 0 Then

                                    'check to see if refy is within 1 brick height of cury
                                    Dim refY As Double = pts2d(0).Y
                                    Dim minY As Double = refY - bht
                                    Dim elevDiff As Double = refY - curY

                                    If curY > minY Then  'cury is above minimum
                                        If curY <= refY Then  'cury is below reference line
                                            'Dim lastx As Double = curX
                                            lastx = curX
                                            curX = startX - (brk * 0.5 * blen)
                                            'if next point is beyond end of bl then create one final point
                                            If curX < endX Then
                                                Dim tps1 As New Point2d(curX, curY)
                                                footPts.Add(tps1)
                                            End If
                                            brk += 1
                                        Else  'cury is above reference line
                                            Dim fSafe As Integer = 0
                                            lasty = curY
                                            Do
                                                curY -= bht
                                                If curY < refY Then
                                                    'curY -= bht
                                                    Exit Do
                                                End If
                                                fSafe += 1
                                            Loop While fSafe < 50
                                            Dim tps1 As New Point2d(curX, lasty)
                                            Dim tps2 As New Point2d(curX - 0.02, curY)
                                            footPts.Add(tps1)
                                            footPts.Add(tps2)
                                        End If
                                    Else    'cury is below minimum
                                        Dim fSafe As Integer = 0
                                        lasty = curY
                                        Do
                                            curY += bht
                                            If curY > minY Then
                                                Exit Do
                                            End If
                                            fSafe += 1
                                        Loop While fSafe < 50
                                        Dim tps1 As New Point2d(curX, lasty)
                                        Dim tps2 As New Point2d(curX - 0.02, curY)
                                        footPts.Add(tps1)
                                        footPts.Add(tps2)
                                    End If
                                End If
SkipPt:
                            End Using

                            failsafe += 1
                            If curX < endX Then Exit Do

                        Loop While failsafe < 10000

                    End If

                    bl.Dispose()

                    'Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    'Dim curspace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)
                    Using pl As New Polyline
                        For k = 0 To footPts.Count - 1
                            pl.AddVertexAt(k, footPts(k), 0, 0, 0)
                        Next
                        If Not curSpc.IsWriteEnabled Then curSpc.UpgradeOpen()
                        curSpc.AppendEntity(pl)
                        acTrans.AddNewlyCreatedDBObject(pl, True)
                    End Using

                    acTrans.Commit()

                Catch ex As Exception
                    MessageBox.Show(vbLf & "Error in footing layout command.  Check data and try again." & vbLf & vbLf & ex.Message)
                    'acTrans.Abort()
                    Exit Sub
                End Try

            End Using

        End Sub
        <CommandMethod("FTBOT")>
        Public Sub FootingBottom()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = CurDwg.Editor

            Dim peo As New PromptEntityOptions(vbLf & "Pick polyline for the top of footing.")

            Dim topId As ObjectId

            With peo
                .AllowNone = False
            End With

            Dim per As PromptEntityResult = ed.GetEntity(peo)

            If per.Status = PromptStatus.OK Then
                topId = per.ObjectId
            Else
                Exit Sub
            End If

            Dim ppo As New PromptPointOptions(vbLf & "Pick point for the top of footing at start point")
            With ppo
                .AllowNone = False
            End With

            Dim pb1 As Point3d
            Dim ppr As PromptPointResult = ed.GetPoint(ppo)

            If ppr.Status = PromptStatus.OK Then
                pb1 = ppr.Value
            Else
                Exit Sub
            End If

            Dim pdo1 As New PromptDistanceOptions(vbLf & "Enter or pick the minimum footing thickness.")

            With pdo1
                .AllowArbitraryInput = True
                .AllowZero = False
                .AllowNegative = False
                .AllowNone = False
            End With

            Dim pdr1 As PromptDoubleResult = ed.GetDistance(pdo1)

            Dim cvr As Double

            If pdr1.Status = PromptStatus.OK Then
                cvr = pdr1.Value
            Else
                Exit Sub
            End If

            Dim pdo2 As New PromptDistanceOptions(vbLf & "Enter the vertical exaggeration factor.")

            With pdo2
                .DefaultValue = 5
                .AllowArbitraryInput = True
                .AllowZero = False
                .AllowNegative = False
                .AllowNone = False
            End With

            Dim pdr2 As PromptDoubleResult = ed.GetDistance(pdo2)

            Dim vScale As Double

            If pdr2.Status = PromptStatus.OK Then
                vScale = pdr2.Value
            Else
                Exit Sub
            End If

            Dim refPln As New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)
            Dim cover As Double = cvr * vScale


            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction

                Try

                    Dim baseline As DBObject = acTrans.GetObject(topId, OpenMode.ForRead)
                    Dim bl As Polyline = TryCast(baseline, Polyline)
                    'Dim tempObj As Object = GetVertexCoords(topid)

                    If bl Is Nothing OrElse bl.Closed = True Then
                        MessageBox.Show("Error.  Picked entity cannot be used for top of footing path.")
                        acTrans.Abort()
                        Exit Sub
                    End If

                    If pb1.GetVectorTo(bl.StartPoint).Length > pb1.GetVectorTo(bl.EndPoint).Length Then
                        If Not bl.IsWriteEnabled Then bl.UpgradeOpen()
                        bl.ReverseCurve()
                        'tempObj = GetVertexCoords(bl.ObjectId)
                    End If

                    Dim spX As Double = bl.StartPoint.X
                    Dim epX As Double = bl.EndPoint.X
                    Dim pltSide As Integer

                    If spX < epX Then
                        pltSide = 1
                    Else
                        pltSide = -1
                    End If

                    'If ide < 0 Then  'bl oriented from right to left
                    '    If spX < epX Then
                    '        If Not bl.IsWriteEnabled Then bl.UpgradeOpen()
                    '        bl.ReverseCurve()
                    '        'tempObj = GetVertexCoords(bl.ObjectId)
                    '        GoTo TryAgain
                    '    End If
                    'Else  'bl oriented from left to right
                    '    If spX > epX Then
                    '        If Not bl.IsWriteEnabled Then bl.UpgradeOpen()
                    '        bl.ReverseCurve()
                    '        'tempObj = GetVertexCoords(bl.ObjectId)
                    '        GoTo TryAgain
                    '    End If
                    'End If

                    Dim tempObj As Object = GetVertices(bl.ObjectId)
                    Dim numverts As Long = bl.NumberOfVertices
                    Dim ptcol As Point2dCollection

                    If TypeOf tempObj Is Point3dCollection Then
                        Dim tempPoints As Point3dCollection = TryCast(tempObj, Point3dCollection)
                        If tempPoints IsNot Nothing Then
                            ptcol = AcCommon.ConvertPoints2d(tempPoints)
                        Else
                            Exit Sub
                        End If
                    ElseIf TypeOf tempObj Is Point2dCollection Then
                        ptcol = CType(tempObj, Point2dCollection)
                        If ptcol Is Nothing Then Exit Sub
                    Else
                        Exit Sub
                    End If

                    'Dim ofPt As New Point3d(bl.StartPoint.X, bl.StartPoint.Y - cover, 0)
                    'Dim transVect As Vector3d = bl.StartPoint.GetVectorTo(ofPt)

                    'Dim bl As Polyline = tbl.Clone
                    'bl.TransformBy(Matrix3d.Displacement(transVect))
                    'If bl.Elevation <> 0 Then bl.Elevation = 0

                    'If Not curSpc.IsWriteEnabled Then curSpc.UpgradeOpen()
                    'curSpc.AppendEntity(bl)
                    'acTrans.AddNewlyCreatedDBObject(bl, True)

                    'Dim cont As Boolean = True

                    Dim linePts As New Point2dCollection
                    Dim firstpt As New Point2d(pb1.X, pb1.Y - cover)
                    linePts.Add(firstpt)

                    If pltSide > 0 Then  'reference moves from left to right
                        For i = 0 To numverts - 1
                            If i = numverts - 1 Then
                                linePts.Add(New Point2d(bl.EndPoint.X, bl.EndPoint.Y - cover))
                            Else
                                Dim curPt As Point2d = ptcol(i)
                                Dim nextPt As Point2d = ptcol(i + 1)

                                Dim curY As Double = ptcol(i).Y
                                Dim nextY As Double = ptcol(i + 1).Y
                                Dim olap As Double = 2 * Abs((curY - nextY) / vScale)
                                If olap < cvr Then olap = cvr

                                If nextY > curY Then 'current point is at a step up
                                    Dim stepBotX As Double = curPt.X + olap
                                    Dim stepTopX As Double = curPt.X + (2 * olap)
                                    Dim botStep As New Point2d(stepBotX, curY - cover)
                                    Dim topstep As New Point2d(stepTopX, nextY - cover)
                                    If Not botStep = firstpt Then linePts.Add(botStep)
                                    linePts.Add(topstep)
                                ElseIf nextY < curY Then    'current point is a step down
                                    Dim stepBotX As Double = curPt.X - olap
                                    Dim stepTopX As Double = curPt.X - (2 * olap)
                                    Dim topStep As New Point2d(stepTopX, curY - cover)
                                    Dim botstep As New Point2d(stepBotX, nextY - cover)
                                    If Not topStep = firstpt Then linePts.Add(topStep)
                                    linePts.Add(botstep)
                                Else   'current vertex is not a step
                                End If
                            End If
                        Next
                    Else  'reference moves from right to left
                        For i = 0 To numverts - 1
                            If i = numverts - 1 Then
                                linePts.Add(New Point2d(bl.EndPoint.X, bl.EndPoint.Y - cover))
                            Else
                                Dim curPt As Point2d = ptcol(i)
                                Dim nextPt As Point2d = ptcol(i + 1)

                                Dim curY As Double = ptcol(i).Y
                                Dim nextY As Double = ptcol(i + 1).Y
                                Dim olap As Double = 2 * Abs((curY - nextY) / vScale)
                                If olap < cvr Then olap = cvr

                                If nextY < curY Then 'current point is at a step down
                                    Dim stepTopX As Double = curPt.X + (2 * olap)
                                    Dim stepBotX As Double = curPt.X + olap
                                    Dim topstep As New Point2d(stepTopX, curY - cover)
                                    Dim botStep As New Point2d(stepBotX, nextY - cover)
                                    If Not topstep = firstpt Then linePts.Add(topstep)
                                    linePts.Add(botStep)
                                ElseIf nextY > curY Then    'current point is a step up
                                    Dim stepBotX As Double = curPt.X - olap
                                    Dim stepTopX As Double = curPt.X - (2 * olap)
                                    Dim botStep As New Point2d(stepBotX, curY - cover)
                                    Dim topstep As New Point2d(stepTopX, nextY - cover)
                                    If Not botStep = firstpt Then linePts.Add(botStep)
                                    linePts.Add(topstep)
                                Else   'current vertex is not a step
                                End If
                            End If
                        Next
                    End If

                    If Not bl.IsDisposed Then bl.Dispose()

                    'Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    'Dim curspace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)

                    Dim curSpc As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForRead)

                    Using pl As New Polyline
                        For k = 0 To linePts.Count - 1
                            pl.AddVertexAt(k, linePts(k), 0, 0, 0)
                        Next
                        If Not curSpc.IsWriteEnabled Then curSpc.UpgradeOpen()
                        curSpc.AppendEntity(pl)
                        acTrans.AddNewlyCreatedDBObject(pl, True)
                    End Using

                    acTrans.Commit()

                Catch ex As Exception
                    MessageBox.Show(vbLf & "Error in footing layout command.  Check data and try again." & vbLf & vbLf & ex.Message)
                    'acTrans.Abort()

                End Try
            End Using
        End Sub
    End Module

    Public Module WallCommon
        Friend Function FootBottom(topID As ObjectId, pb1 As Point3d, rwData As RetainingWall, pltside As Integer) As Boolean

            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = CurDwg.Editor

            Dim refPln As New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction

                Try

                    Dim baseline As DBObject = acTrans.GetObject(topID, OpenMode.ForRead)
                    Dim bl As Polyline = TryCast(baseline, Polyline)
                    'Dim tempObj As Object = GetVertexCoords(topid)

                    If bl Is Nothing OrElse bl.Closed = True Then
                        MessageBox.Show("Error.  Top of Footing not defined.")
                        Return False
                        Exit Function
                    End If

                    'If ide < 0 Then  'bl oriented from right to left
                    '    If spX < epX Then
                    '        If Not bl.IsWriteEnabled Then bl.UpgradeOpen()
                    '        bl.ReverseCurve()
                    '        'tempObj = GetVertexCoords(bl.ObjectId)
                    '        GoTo TryAgain
                    '    End If
                    'Else  'bl oriented from left to right
                    '    If spX > epX Then
                    '        If Not bl.IsWriteEnabled Then bl.UpgradeOpen()
                    '        bl.ReverseCurve()
                    '        'tempObj = GetVertexCoords(bl.ObjectId)
                    '        GoTo TryAgain
                    '    End If
                    'End If

                    Dim tempObj As Object = GetVertices(bl.ObjectId)
                    Dim numverts As Long = bl.NumberOfVertices
                    Dim ptcol As Point2dCollection

                    If TypeOf tempObj Is Point3dCollection Then
                        Dim tempPoints As Point3dCollection = TryCast(tempObj, Point3dCollection)
                        If tempPoints IsNot Nothing Then
                            ptcol = ConvertPoints2d(tempPoints)
                        Else
                            Return False
                            Exit Function
                        End If
                    ElseIf TypeOf tempObj Is Point2dCollection Then
                        ptcol = CType(tempObj, Point2dCollection)
                        If ptcol Is Nothing Then
                            Return False
                            Exit Function
                        End If
                    Else
                        Return False
                        Exit Function
                    End If

                    'Dim ofPt As New Point3d(bl.StartPoint.X, bl.StartPoint.Y - cover, 0)
                    'Dim transVect As Vector3d = bl.StartPoint.GetVectorTo(ofPt)

                    'Dim bl As Polyline = tbl.Clone
                    'bl.TransformBy(Matrix3d.Displacement(transVect))
                    'If bl.Elevation <> 0 Then bl.Elevation = 0

                    'If Not curSpc.IsWriteEnabled Then curSpc.UpgradeOpen()
                    'curSpc.AppendEntity(bl)
                    'acTrans.AddNewlyCreatedDBObject(bl, True)

                    'Dim cont As Boolean = True

                    Dim cover As Double = rwData.ProfFootThickness
                    Dim vScale As Double = rwData.VertFactor
                    Dim ftThick As Double = rwData.FootingThickness

                    Dim linePts As New Point2dCollection
                    Dim firstpt As New Point2d(pb1.X, pb1.Y - cover)

                    linePts.Add(firstpt)

                    If pltside > 0 Then  'reference moves from left to right
                        For i = 0 To numverts - 1
                            If i = numverts - 1 Then
                                linePts.Add(New Point2d(bl.EndPoint.X, bl.EndPoint.Y - cover))
                            Else
                                Dim curPt As Point2d = ptcol(i)
                                Dim nextPt As Point2d = ptcol(i + 1)

                                Dim curY As Double = ptcol(i).Y
                                Dim nextY As Double = ptcol(i + 1).Y
                                Dim olap As Double = 2 * Abs((curY - nextY) / vScale)
                                If olap < ftThick Then olap = ftThick

                                If nextY > curY Then 'current point is at a step up
                                    Dim stepBotX As Double = curPt.X + olap
                                    Dim stepTopX As Double = curPt.X + (2 * olap)
                                    Dim botStep As New Point2d(stepBotX, curY - cover)
                                    Dim topstep As New Point2d(stepTopX, nextY - cover)
                                    If Not botStep = firstpt Then linePts.Add(botStep)
                                    linePts.Add(topstep)
                                ElseIf nextY < curY Then    'current point is a step down
                                    Dim stepBotX As Double = curPt.X - olap
                                    Dim stepTopX As Double = curPt.X - (2 * olap)
                                    Dim topStep As New Point2d(stepTopX, curY - cover)
                                    Dim botstep As New Point2d(stepBotX, nextY - cover)
                                    If Not topStep = firstpt Then linePts.Add(topStep)
                                    linePts.Add(botstep)
                                Else   'current vertex is not a step
                                End If
                                End If
                        Next
                    Else  'reference moves from right to left
                        For i = 0 To numverts - 1
                            If i = numverts - 1 Then
                                linePts.Add(New Point2d(bl.EndPoint.X, bl.EndPoint.Y - cover))
                            Else
                                Dim curPt As Point2d = ptcol(i)
                                Dim nextPt As Point2d = ptcol(i + 1)

                                Dim curY As Double = ptcol(i).Y
                                Dim nextY As Double = ptcol(i + 1).Y
                                Dim olap As Double = 2 * Abs((curY - nextY) / vScale)
                                If olap < ftThick Then olap = ftThick

                                If nextY < curY Then 'current point is at a step down
                                    Dim stepTopX As Double = curPt.X + olap
                                    Dim stepBotX As Double = curPt.X + (2 * olap)
                                    Dim topstep As New Point2d(stepTopX, curY - cover)
                                    Dim botStep As New Point2d(stepBotX, nextY - cover)
                                    If Not topstep = firstpt Then linePts.Add(topstep)
                                    linePts.Add(botStep)
                                ElseIf nextY > curY Then    'current point is a step up
                                    Dim stepBotX As Double = curPt.X - olap
                                    Dim stepTopX As Double = curPt.X - (2 * olap)
                                    Dim botStep As New Point2d(stepBotX, curY - cover)
                                    Dim topstep As New Point2d(stepTopX, nextY - cover)
                                    If Not botStep = firstpt Then linePts.Add(botStep)
                                    linePts.Add(topstep)
                                Else   'current vertex is not a step
                                End If
                            End If
                        Next
                    End If

                    If Not bl.IsDisposed Then bl.Dispose()

                    'Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    'Dim curspace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)

                    Dim curSpc As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForRead)

                    Using pl As New Polyline
                        For k = 0 To linePts.Count - 1
                            pl.AddVertexAt(k, linePts(k), 0, 0, 0)
                        Next
                        If Not curSpc.IsWriteEnabled Then curSpc.UpgradeOpen()
                        curSpc.AppendEntity(pl)
                        acTrans.AddNewlyCreatedDBObject(pl, True)
                    End Using

                    acTrans.Commit()

                Catch ex As Exception
                    MessageBox.Show(vbLf & "Error in FootBottom function.  Check data and try again." & vbLf & vbLf & ex.Message)
                    'acTrans.Abort()
                    Return False
                    Exit Function
                End Try
            End Using

            Return True

        End Function

        <CommandMethod("WallLine")>
        Public Sub WallLine()

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim sLayerName As String = SetLayName()

            Dim width As Double
            Dim dashLength As Double
            Dim goodWall As Boolean = False

            'If Width and dashlength are already set, then ask if they should be used.

            Dim useWallData As Boolean = GetWallData()

            'if true, make the wall line and exit
            If useWallData Then
                If m_Wallwidth > 0 And m_WDDLength > 0 Then
                    width = m_Wallwidth
                    dashLength = m_WDDLength
                    goodWall = MakeWallLine(width, dashLength, sLayerName)
                    If Not goodWall Then
                        MessageBox.Show("Error in MakeWallLine Function.")
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            Else
                'If not, proceed with declaring variables
dashLengthinput:
                'get the dash length
                Dim pDblOpts As New PromptDoubleOptions(vbLf & "Input the length of the wall dash-gap in drawing units.")
                With pDblOpts
                    .AllowZero = False
                    .AllowNegative = False
                    .AllowNone = False
                End With

                'pStrOpts.AllowSpaces = False
                Dim pDblRes As PromptDoubleResult = ed.GetDouble(pDblOpts)

                If pDblRes.Status = PromptStatus.OK Then
                    dashLength = pDblRes.Value
                    If dashLength <= 0 Then GoTo EndFunc
                Else
                    ed.WriteMessage(vbLf & "Input cancelled. Ending command.")
                    Exit Sub
                End If

                m_WDDLength = dashLength

wallWidthInput:

                'get the wall width
                pDblOpts = New PromptDoubleOptions(vbLf & "Input the Width of the wall in drawing units.")
                With pDblOpts
                    .AllowZero = False
                    .AllowNegative = False
                    .AllowNone = False
                End With
                pDblRes = ed.GetDouble(pDblOpts)

                If pDblRes.Status = PromptStatus.OK Then
                    width = pDblRes.Value
                    If width <= 0 Then GoTo EndFunc
                Else
                    ed.WriteMessage(vbLf & "Input cancelled. Ending command.")
                    Exit Sub
                End If

                m_Wallwidth = width

Skipit:
                goodWall = MakeWallLine(m_Wallwidth, m_WDDLength, sLayerName)
EndFunc:
            End If


            If Not goodWall Then
                MsgBox("ERROR. Check picked entity or input data.")
            End If

        End Sub

        Private Function MakeWallLine(ByVal width As Double, ByVal dD As Double, Optional ByVal sLayerName As String = "0") As Boolean

            Dim ofdist = width / 2  'set the half-width of the line

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim DwgLtScale As Double = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("LTSCALE")
            Dim dDL As Double = dD / DwgLtScale

            Dim ltName As String
            If Not WLisLoaded("WALLLINE") Then
                MessageBox.Show("Error Loading Linetype. Manually load WALLLINE linetype from EESCustomLinetypes.lin and invoke command again.")
                Return False
                Exit Function
            Else
                ltName = "WALLLINE"
            End If

Retry:
            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()

                Dim options As New PromptEntityOptions(String.Format(vbLf & "Select an entity for wall centerline:"))
                'options.SetRejectMessage(vbLf & "The selected object is not a 2D polyline.")
                'options.AddAllowedClass(GetType(Polyline), True)
                Dim result As PromptEntityResult = ed.GetEntity(options)

                If result.Status = PromptStatus.OK Then
                    Dim cLineID = result.ObjectId
                    Dim cLinetemp As Curve = acTrans.GetObject(cLineID, OpenMode.ForRead)

                    Dim cLine As Curve = cLinetemp.GetOrthoProjectedCurve(New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis))

                    If cLine Is Nothing Then
                        Return False
                        Exit Function
                    End If

                    'Try
                    'create collections of objects for the offset curves
                    Dim acOffColl1 As DBObjectCollection = cLinetemp.GetOffsetCurves(ofdist)
                    Dim acOffColl2 As DBObjectCollection = cLinetemp.GetOffsetCurves(-ofdist)

                    'convert the collections to curves
                    Dim Offent1 As Entity = TryCast(acOffColl1(0), Entity)
                    Dim Offent2 As Entity = TryCast(acOffColl2(0), Entity)

                    'put them on the layer supplied as a parameter
                    Offent1.Layer = sLayerName
                    Offent2.Layer = sLayerName

                    'open model space for write
                    Dim curSpace As BlockTableRecord = acTrans.GetObject(dwgDB.CurrentSpaceId, OpenMode.ForWrite)

                    'add offset objects to Database and transaction
                    curSpace.AppendEntity(Offent1)
                    curSpace.AppendEntity(Offent2)
                    acTrans.AddNewlyCreatedDBObject(Offent1, True)
                    acTrans.AddNewlyCreatedDBObject(Offent2, True)

                    'Determine the type of offset objects and invoke the correct end closing routine
                    Select Case Offent1.GetType
                        Case Is = GetType(Polyline)
                            If Not cLine.Closed Then CloseLines(Offent1, Offent2, "B")
                        Case Is = GetType(Polyline2d)
                            If Not cLine.Closed Then CloseLines(Offent1, Offent2, "B")
                        Case Else
                            CloseCurves(Offent1, Offent2, sLayerName, "B")
                    End Select

                    Select Case cLine.GetType
                        Case Is = GetType(Polyline2d)
                            Dim cclone As Polyline2d = TryCast(cLine.Clone, Polyline2d)
                            If Not curSpace.IsWriteEnabled Then curSpace.UpgradeOpen()
                            curSpace.AppendEntity(cclone)
                            acTrans.AddNewlyCreatedDBObject(cclone, True)
                            If Not cclone.IsWriteEnabled Then cLine.UpgradeOpen()
                            cclone.ConstantWidth = width
                            cclone.Layer = sLayerName
                            cclone.Linetype = ltName
                            cclone.LinetypeScale = dDL
                            acTrans.Commit()

                        Case Is = GetType(Polyline)
                            Dim cclone As Polyline = TryCast(cLine.Clone, Polyline)
                            If Not curSpace.IsWriteEnabled Then curSpace.UpgradeOpen()
                            curSpace.AppendEntity(cclone)
                            acTrans.AddNewlyCreatedDBObject(cclone, True)
                            If Not cclone.IsWriteEnabled Then cclone.UpgradeOpen()
                            cclone.ConstantWidth = width
                            cclone.Layer = sLayerName
                            cclone.Linetype = ltName
                            cclone.LinetypeScale = dDL
                            acTrans.Commit()

                        Case Else
                            MessageBox.Show("Selected Entity cannot have width." & vbLf & "Convert to a 2d Polyline before creating wall.")
                            Return False
                            acTrans.Abort()
                            Exit Function
                    End Select

                    'Catch ex As Exception
                    'MessageBox.Show("Error creating offset wall lines.  Abort routine.")
                    '    Return False
                    '    acTrans.Abort()
                    '    Exit Function
                    'End Try
                End If
                Return True
            End Using
        End Function

        Private Function GetWallData() As Boolean

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            If m_Wallwidth > 0 AndAlso m_WDDLength > 0 Then
                Dim pMsg As String = vbLf & "Use current values for wall width (" & m_Wallwidth & ") and dash (" & m_WDDLength & ")?  (Y)es or (N)o?"
                Dim pKWopts As New PromptKeywordOptions("")
                With pKWopts
                    .Keywords.Add("Y")
                    .Keywords.Add("N")
                    .AppendKeywordsToMessage = True
                    .Message = pMsg
                    .AllowNone = False
                    .AllowArbitraryInput = False
                End With

                Dim pr As PromptResult = ed.GetKeywords(pKWopts)
                Dim pkStr As String = ""

                If pr.Status = PromptStatus.OK Then
                    pkStr = pr.StringResult

                    If pkStr = "Y" Then
                        Return True
                        Exit Function
                    End If
                Else
                    ed.WriteMessage(vbLf & "Command cancelled.")
                    Return False
                    Exit Function
                End If
            End If

            Return False

        End Function

        Private Function SetLayName() As String

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim sLayerName As String = m_AcLayerName

            If Not sLayerName = "" Then
                Dim pMsg As String = vbLf & "Use current wall layer (" & m_AcLayerName & ") for new wall?  (Y)es or (N)o"
                Dim pKWopts As New PromptKeywordOptions("")
                With pKWopts
                    .Keywords.Add("Y")
                    .Keywords.Add("N")
                    .AppendKeywordsToMessage = True
                    .Message = pMsg
                    .AllowNone = False
                    .AllowArbitraryInput = False
                End With

                Dim pr As PromptResult = ed.GetKeywords(pKWopts)
                Dim pkStr As String = ""

                If pr.Status = PromptStatus.OK Then
                    pkStr = pr.StringResult

                    If pkStr = "Y" Then
                        Return m_AcLayerName
                        Exit Function
                    End If
                Else
                    ed.WriteMessage(vbLf & "Command cancelled.")
                    Return Nothing
                    Exit Function
                End If
            End If

            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim lyrtbl As LayerTable = actrans.GetObject(DwgDB.LayerTableId, OpenMode.ForRead)
                Dim lyrList As New List(Of String)
                For Each lyrID As ObjectId In lyrtbl
                    Dim ltr As LayerTableRecord = TryCast(actrans.GetObject(lyrID, OpenMode.ForRead), LayerTableRecord)
                    If ltr IsNot Nothing Then
                        lyrList.Add(ltr.Name)
                        lyrList.Sort()
                    End If
                Next

                Dim laySel As New Picker
                With laySel
                    .TopLabel.Text = "Select layer for new wall."
                    For Each st As String In lyrList
                        .BxList.Items.Add(st)
                    Next
                    .BxList.SelectionMode = Windows.Forms.SelectionMode.One
                End With

                Dim dR As DialogResult = laySel.ShowDialog
                Dim res As New Collection

                If dR = DialogResult.OK Then
                    res = laySel.PickCol
                Else
                    res.Add("0")
                End If

                laySel.Dispose()

                If Not String.IsNullOrEmpty(res(0)) Then
                    sLayerName = res(0)
                    m_AcLayerName = sLayerName
                Else
                    sLayerName = "0"
                    m_AcLayerName = "0"
                End If

                Return sLayerName
                actrans.Commit()

            End Using

        End Function

        Public Sub CloseLines(Ent1 As Entity, Ent2 As Entity, Optional closer As String = "")

            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Dim myString As String = closer

            If myString = "" Then

                Dim pKeyOpts As New PromptKeywordOptions("")
                With pKeyOpts
                    .Message = vbLf & "Close the new feature?  Start (S) End (E) Both (B) or none (N)"
                    .Keywords.Add("S")
                    .Keywords.Add("E")
                    .Keywords.Add("B")
                    .Keywords.Add("N")
                    .AllowNone = False
                    .AppendKeywordsToMessage = True
                End With

                Dim pKeyRes As PromptResult = ed.GetKeywords(pKeyOpts)
                myString = pKeyRes.StringResult
            End If

            'If per.Status = PromptStatus.OK Then

            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim obj1 As DBObject = tr.GetObject(Ent1.Id, OpenMode.ForRead)
                Dim obj2 As DBObject = tr.GetObject(Ent2.Id, OpenMode.ForRead)
                Dim lwp1 As Polyline = TryCast(obj1, Polyline)
                Dim lwp2 As Polyline = TryCast(obj2, Polyline)
                If lwp1 IsNot Nothing Then
                    Dim Spt1 As Point3d = lwp1.StartPoint
                    Dim Ept1 As Point3d = lwp1.EndPoint
                    Dim Spt2 As Point3d = lwp2.StartPoint
                    Dim Ept2 As Point3d = lwp2.EndPoint

                    Dim currSpace As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead)

                    Select Case myString
                        Case Is = "B"
                            Dim startLine As New Polyline
                            With startLine
                                .AddVertexAt(0, New Point2d(Spt1.X, Spt1.Y), 0, 0, 0)
                                .AddVertexAt(1, New Point2d(Spt2.X, Spt2.Y), 0, 0, 0)
                            End With
                            If currSpace.IsWriteEnabled = False Then currSpace.UpgradeOpen()
                            currSpace.AppendEntity(startLine)
                            tr.AddNewlyCreatedDBObject(startLine, True)

                            Dim endline As New Polyline
                            With endline
                                .AddVertexAt(0, New Point2d(Ept1.X, Ept1.Y), 0, 0, 0)
                                .AddVertexAt(1, New Point2d(Ept2.X, Ept2.Y), 0, 0, 0)
                            End With
                            If currSpace.IsWriteEnabled = False Then currSpace.UpgradeOpen()
                            currSpace.AppendEntity(endline)
                            tr.AddNewlyCreatedDBObject(endline, True)
                            lwp1.JoinEntity(startLine)
                            lwp1.JoinEntity(endline)
                            lwp1.JoinEntity(lwp2)
                            lwp1.Closed = True
                            lwp2.Erase()
                            endline.Erase()
                            startLine.Erase()

                        Case Is = "E"
                            Dim endline As New Polyline
                            With endline
                                .AddVertexAt(0, New Point2d(Ept1.X, Ept1.Y), 0, 0, 0)
                                .AddVertexAt(1, New Point2d(Ept2.X, Ept2.Y), 0, 0, 0)
                            End With
                            If currSpace.IsWriteEnabled = False Then currSpace.UpgradeOpen()
                            currSpace.AppendEntity(endline)
                            tr.AddNewlyCreatedDBObject(endline, True)
                            lwp1.JoinEntity(endline)
                            lwp1.JoinEntity(lwp2)
                            lwp2.Erase()
                            endline.Erase()

                        Case Is = "S"
                            Dim Startline As New Polyline
                            With Startline
                                .AddVertexAt(0, New Point2d(Spt1.X, Spt1.Y), 0, 0, 0)
                                .AddVertexAt(1, New Point2d(Spt2.X, Spt2.Y), 0, 0, 0)
                            End With
                            If currSpace.IsWriteEnabled = False Then currSpace.UpgradeOpen()
                            currSpace.AppendEntity(Startline)
                            tr.AddNewlyCreatedDBObject(Startline, True)
                            lwp1.JoinEntity(Startline)
                            lwp1.JoinEntity(lwp2)
                            lwp2.Erase()
                            Startline.Erase()
                        Case Is = "N"
                    End Select
                Else

                End If
                tr.Commit()
            End Using
            '            End If
        End Sub

        Friend Sub CloseCurves(Ent1 As Entity, Ent2 As Entity, sLayerName As String, Optional closer As String = "")

            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Dim Mystring As String = closer
            If Mystring = "" Then

                Dim pKeyOpts As New PromptKeywordOptions("")
                With pKeyOpts
                    .Message = vbLf & "Close the new feature?  Start (S) End (E) Both (B) or none (N)"
                    .Keywords.Add("S")
                    .Keywords.Add("E")
                    .Keywords.Add("B")
                    .Keywords.Add("N")
                    .AllowNone = False
                End With

                Dim pKeyRes As PromptResult = ed.GetKeywords(pKeyOpts)
                Mystring = pKeyRes.StringResult
            End If

            'If per.Status = PromptStatus.OK Then

            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim obj1 As DBObject = tr.GetObject(Ent1.Id, OpenMode.ForRead)
                Dim obj2 As DBObject = tr.GetObject(Ent2.Id, OpenMode.ForRead)
                Dim lwp1 As Curve = TryCast(obj1, Curve)
                Dim lwp2 As Curve = TryCast(obj2, Curve)

                If lwp1 IsNot Nothing And lwp2 IsNot Nothing Then
                    Dim Spt1 As Point3d = lwp1.StartPoint
                    Dim Ept1 As Point3d = lwp1.EndPoint
                    Dim Spt2 As Point3d = lwp2.StartPoint
                    Dim Ept2 As Point3d = lwp2.EndPoint

                    Dim currspace As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead)

                    Select Case Mystring
                        Case Is = "B"
                            Dim Startline As New Polyline
                            With Startline
                                .AddVertexAt(0, New Point2d(Spt1.X, Spt1.Y), 0, 0, 0)
                                .AddVertexAt(1, New Point2d(Spt2.X, Spt2.Y), 0, 0, 0)
                                .Layer = sLayerName
                            End With
                            If currspace.IsWriteEnabled = False Then currspace.UpgradeOpen()
                            currspace.AppendEntity(Startline)
                            tr.AddNewlyCreatedDBObject(Startline, True)

                            Dim endline As New Polyline
                            With endline
                                .AddVertexAt(0, New Point2d(Ept1.X, Ept1.Y), 0, 0, 0)
                                .AddVertexAt(1, New Point2d(Ept2.X, Ept2.Y), 0, 0, 0)
                                .Layer = sLayerName
                            End With

                            If currspace.IsWriteEnabled = False Then currspace.UpgradeOpen()
                            currspace.AppendEntity(endline)
                            tr.AddNewlyCreatedDBObject(endline, True)

                        Case Is = "E"
                            Dim endline As New Polyline
                            With endline
                                .AddVertexAt(0, New Point2d(Ept1.X, Ept1.Y), 0, 0, 0)
                                .AddVertexAt(1, New Point2d(Ept2.X, Ept2.Y), 0, 0, 0)
                                .Layer = sLayerName
                            End With
                            If currspace.IsWriteEnabled = False Then currspace.UpgradeOpen()
                            currspace.AppendEntity(endline)
                            tr.AddNewlyCreatedDBObject(endline, True)

                        Case Is = "S"
                            Dim Startline As New Polyline
                            With Startline
                                .AddVertexAt(0, New Point2d(Spt1.X, Spt1.Y), 0, 0, 0)
                                .AddVertexAt(1, New Point2d(Spt2.X, Spt2.Y), 0, 0, 0)
                                .Layer = sLayerName
                            End With
                            If currspace.IsWriteEnabled = False Then currspace.UpgradeOpen()
                            currspace.AppendEntity(Startline)
                            tr.AddNewlyCreatedDBObject(Startline, True)
                        Case Is = "N"

                    End Select

                    tr.Commit()
                Else
                End If
            End Using
            '            End If
        End Sub

        Private Function WLisLoaded(ltName As String) As Boolean

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor
            Dim lTFilePath = "\\EESSERVER\datadisk\CAD\Lin-Shp\EESCustomLinetypes.lin"
            'Dim wLTfound = False

            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()
                Dim ltTable As LinetypeTable = acTrans.GetObject(dwgDB.LinetypeTableId, OpenMode.ForRead)

                If Not ltTable.Has(ltName) Then
                    Try
                        ltTable.UpgradeOpen()
                        dwgDB.LoadLineTypeFile("WALLLINE", lTFilePath)

                    Catch ex As Exception
                        acTrans.Abort()
                        Return False
                        Exit Function
                    End Try
                End If
                acTrans.Commit()
            End Using
            Return True
        End Function

    End Module

    Public Class RetainingWall
        Inherits CollectionBase

        Private m_HasConcBase As Boolean
        Private m_vertFootStep As Double
        Private m_footCover As Double
        Private m_VertFactor As Double
        Private m_brickHt As Double
        Private m_brickLength As Double
        Private m_footThick As Double

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

        Public Property FootingStepHeight As Double
            Get
                Return m_vertFootStep
            End Get
            Set(value As Double)
                m_vertFootStep = value
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

    End Class

End Namespace
