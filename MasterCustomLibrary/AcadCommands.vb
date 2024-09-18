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
Imports MasterCustomLibrary.AcCommon
Imports Autodesk.AutoCAD.Internal


'project and file (c) David Eisenbeisz 2023

Namespace AcCommands
    Public Module ArcCommands

        Public refPlane As New Plane(Point3d.Origin, Vector3d.ZAxis)

        <CommandMethod("AngChk")>
        Public Sub CheckAngles()
            Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgdb As Database = acDwg.Database
            Dim ed As Editor = acDwg.Editor

            Dim peo As New PromptEntityOptions(vbLf & "Select linear entity")
            With peo
                .SetRejectMessage(vbLf & "Selected entity must be a polyline or a line.")
                .AddAllowedClass(GetType(Line), True)
                .AddAllowedClass(GetType(Polyline), True)
            End With

            Dim objId As ObjectId
            Dim peR As PromptEntityResult = ed.GetEntity(peo)
            If peR.Status = PromptStatus.OK Then
                objId = peR.ObjectId
            Else
                ed.WriteMessage(vbLf & "Command Cancelled.")
                Exit Sub
            End If

            Using acTrans As Transaction = dwgdb.TransactionManager.StartTransaction
                Dim dbobj As DBObject = acTrans.GetObject(objId, OpenMode.ForRead)

                If TypeOf dbobj Is Line Then
                    Dim ln As Line = TryCast(dbobj, Line)
                    Dim sP As Point3d = ln.StartPoint
                    Dim eP As Point3d = ln.EndPoint
                    Dim sp2D As Point2d = sP.Convert2d(refPlane)
                    Dim ep2D As Point2d = eP.Convert2d(refPlane)

                    Dim v2d As Vector2d = sp2D.GetVectorTo(ep2D)
                    Dim ornt1 As Double = v2d.Angle
                    Dim ang As New AngleObj(ornt1, True, 0)
                    Dim sb As New StringBuilder

                    With sb
                        .AppendLine(ang.DecAzimuth)
                        .AppendLine(ang.DecimalDegrees)
                        .AppendLine(ang.AutoCADdms)
                        .AppendLine(ang.DMS)
                        .AppendLine(ang.Surveyors)
                        .AppendLine()
                    End With

                    ed.WriteMessage(sb.ToString)
                End If
                acTrans.Commit()
            End Using
        End Sub

        <CommandMethod("LISTARCDATA")>
        Public Sub ListArcData()
            'by David Eisenbeisz

            Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = acDwg.Database
            Dim ed As Editor = acDwg.Editor
            Dim mypt1 As Point3d
            Dim mypt2 As Point3d
            Dim mypt3 As Point3d
            Dim myArcID As ObjectId
            Dim aD As ArcData

            Dim pp0 As New PromptPointOptions(vbLf & "Select first point on arc or press escape to select arc entity: ")
            Dim pp0Res As PromptPointResult = ed.GetPoint(pp0)
            If pp0Res.Status = PromptStatus.OK Then
                mypt1 = pp0Res.Value
                Dim pp1 As New PromptPointOptions(vbLf & "Select second point on arc: ")
                Dim pp1Res As PromptPointResult = ed.GetPoint(pp1)
                If pp1Res.Status = PromptStatus.OK Then
                    mypt2 = pp1Res.Value
                    Dim pp2 As New PromptPointOptions(vbLf & "Select third point on arc: ")
                    Dim pp2Res As PromptPointResult = ed.GetPoint(pp2)
                    If pp2Res.Status = PromptStatus.OK Then
                        mypt3 = pp2Res.Value
                    Else
                        ed.WriteMessage(vbLf & "Command Cancelled.")
                        Exit Sub
                    End If
                Else
                    ed.WriteMessage(vbLf & "Command Cancelled.")
                    Exit Sub
                End If
                aD = New ArcData(mypt1, mypt2, mypt3)

            Else
                Dim peo As New PromptEntityOptions(vbLf & "Select arc entity")
                With peo
                    .SetRejectMessage(vbLf & "Selected entity must be an arc.")
                    .AddAllowedClass(GetType(Arc), True)
                End With
                Dim peR As PromptEntityResult = ed.GetEntity(peo)
                If peR.Status = PromptStatus.OK Then
                    myArcID = peR.ObjectId
                Else
                    ed.WriteMessage(vbLf & "Command Cancelled.")
                    Exit Sub
                End If

                Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                    Dim myArc As Arc = acTrans.GetObject(myArcID, OpenMode.ForRead)
                    'Dim acArc As New CircularArc2d(New Point2d(myArc.StartPoint.X, myArc.StartPoint.Y), New Point2d(myArc.EndPoint.X, myArc.EndPoint.Y), GetArcBulge(myArc), False)
                    Dim midPt As Point3d = myArc.GetPointAtDist(myArc.Length / 2)
                    Dim tpt1 As New Point2d(myArc.StartPoint.X, myArc.StartPoint.Y)
                    Dim tpt3 As New Point2d(myArc.EndPoint.X, myArc.EndPoint.Y)
                    Dim tpt2 As New Point2d(midPt.X, midPt.Y)
                    Dim acArc As New CircularArc2d(tpt1, tpt2, tpt3)
                    aD = New ArcData(acArc, mypt1.Z)
                End Using
            End If

            Dim sb As New StringBuilder
            With sb
                .AppendLine(vbLf & "Start Point: " & aD.StartPoint.ToString)
                .AppendLine("End Point: " & aD.EndPoint.ToString)
                .AppendLine("Point on Arc: " & aD.MidPoint.ToString)
                .AppendLine("Center 3D Point: " & aD.Center3d.ToString)
                .AppendLine("Center 2D Point: " & aD.Center2d.ToString)
                .AppendLine("Radius: " & aD.Radius.ToString)
                .AppendLine("Chord: " & aD.Chord.ToString)
                .AppendLine("Delta: " & aD.Delta.ToString)
                .AppendLine("Length: " & aD.Length.ToString)
                .AppendLine("Bulge: " & aD.Bulge.ToString)
                .AppendLine("Is Clockwise: " & aD.IsClockwise.ToString)
                .AppendLine("Start Angle: " & aD.StartAngle.ToString)
                .AppendLine("End Angle: " & aD.EndAngle.ToString)
                .AppendLine()
                .AppendLine("SubDelta1: " & aD.SubDelta1.ToString)
                .AppendLine("SubDelta2: " & aD.SubDelta2.ToString)
                .AppendLine("SubChord1: " & aD.SubChord1.ToString)
                .AppendLine("SubChord2: " & aD.SubChord2.ToString)
                .AppendLine("sublength1: " & aD.SubLength1.ToString)
                .AppendLine("sublength2: " & aD.SubLength2.ToString)
                .AppendLine("SubBulge1: " & aD.SubBulge1.ToString)
                .AppendLine("SubBulge2: " & aD.SubBulge2.ToString)
                .AppendLine("Mid Angle: " & aD.MidAngle.ToString)
            End With

            MessageBox.Show(sb.ToString)
            ed.WriteMessage(vbLf & sb.ToString & vbLf)

        End Sub

    End Module


    Public Module BlockCommands

        <CommandMethod("LABLKSCUST", CommandFlags.UsePickSet Or CommandFlags.Redraw Or CommandFlags.Modal)>
        Public Sub LabelBlocksGeneric()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim SelResult As PromptSelectionResult = ed.SelectImplied()

            Dim ppo As New PromptPointOptions(vbLf & "Pick upper left corner of first cell")
            With ppo
                .AllowNone = False
                .AllowArbitraryInput = True
            End With

            Dim p1 As Point3d
            Dim p2 As Point3d
            Dim userPick As Boolean = True

            Dim ppr As PromptPointResult = ed.GetPoint(ppo)

            If ppr.Status = PromptStatus.OK Then
                p1 = ppr.Value
            Else
                Exit Sub
            End If

            Dim pco As New PromptCornerOptions(vbLf & "Pick the lower right corner of first cell.", p1)
            With pco
                .UseDashedLine = True
                .AllowArbitraryInput = True
            End With

            Dim pcr As PromptPointResult = ed.GetCorner(pco)

            If pcr.Status = PromptStatus.OK Then
                p2 = pcr.Value
            Else
                Exit Sub
            End If

            'Dim pdo1 As New PromptDistanceOptions(vbLf & "Pick or enter the width of a single grid (page) in drawing units ")
            'With pdo1
            '    .AllowNegative = False
            '    .AllowNone = False
            '    .AllowArbitraryInput = True
            'End With

            'Dim pdr1 As PromptDoubleResult = ed.GetDistance(pdo1)

            'Dim pageWdth As Double

            'If pdr1.Status = PromptStatus.OK Then
            '    pageWdth = pdr1.Value
            'Else
            '    Exit Sub
            'End If

            Dim pio1 As New PromptIntegerOptions(vbLf & "Enter the number of cells (blocks) in each row.")
            With pio1
                .AllowNegative = False
                .AllowNone = False
                .AllowArbitraryInput = True
            End With

            Dim pir1 As PromptIntegerResult = ed.GetInteger(pio1)

            Dim cCols As Integer

            If pir1.Status = PromptStatus.OK Then
                cCols = pir1.Value
            Else
                Exit Sub
            End If

            Dim pio2 As New PromptIntegerOptions(vbLf & "Enter the number of cells (blocks) in each column.")
            With pio2
                .AllowNegative = False
                .AllowNone = False
                .AllowArbitraryInput = True
            End With

            Dim cRows As Integer

            Dim pir2 As PromptIntegerResult = ed.GetInteger(pio2)

            If pir2.Status = PromptStatus.OK Then
                cRows = pir2.Value
            Else
                Exit Sub
            End If

            Dim vert As Double = p2.Y - p1.Y
            Dim horiz As Double = p2.X - p1.X

            Dim pdo2 As New PromptDistanceOptions(vbLf & "Pick or enter the horizontal distance between grids (pages) or press escape key if not multiple pages.")
            With pdo2
                .AllowNegative = False
                .AllowNone = False
                .AllowArbitraryInput = True
            End With

            Dim pdr2 As PromptDoubleResult = ed.GetDistance(pdo2)

            Dim margWidth As Double
            Dim multiPages As Boolean = False

            If pdr2.Status = PromptStatus.OK Then
                margWidth = pdr2.Value
                multiPages = True
            Else
                margWidth = 0
            End If

            Dim numPages As Integer

            If multiPages Then
                Dim pio3 As New PromptIntegerOptions(vbLf & "Enter the number of grids (pages).")
                With pio3
                    .AllowNegative = False
                    .AllowNone = False
                    .AllowArbitraryInput = True
                End With

                Dim pir3 As PromptIntegerResult = ed.GetInteger(pio3)

                If pir3.Status = PromptStatus.OK Then
                    numPages = pir3.Value
                Else
                    Exit Sub
                End If
            End If

            Dim pWidth = horiz * cCols + margWidth

            If SelResult.Status = PromptStatus.Error Then
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select blocks references to label:")}
                SelResult = ed.GetSelection(Seloptions)
            Else
                ed.SetImpliedSelection(New ObjectId(-1) {})
            End If

            If SelResult.Status = PromptStatus.OK Then

                Dim acSSet As SelectionSet = SelResult.Value
                Dim MyobjIDs As ObjectId() = acSSet.GetObjectIds

                Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction()
                    'Dim i As Integer = 0
                    For Each objID As ObjectId In MyobjIDs
                        Dim blklist As New List(Of String)
                        Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForRead)
                        If TypeOf dbObj Is BlockReference Then
                            Dim bref As BlockReference = TryCast(dbObj, BlockReference)
                            Dim bName As String = ""
                            If bref.IsDynamicBlock Then
                                Dim bID As ObjectId = bref.DynamicBlockTableRecord
                                Dim dynBlk As BlockTableRecord = actrans.GetObject(bID, OpenMode.ForRead)
                                If dynBlk IsNot Nothing Then bName = dynBlk.Name
                            Else
                                bName = bref.Name
                            End If

                            Dim bPos As Point3d = bref.Position
                            Dim posX As Double = bPos.X
                            Dim posY As Double = bPos.Y

                            Dim sht As Long

                            If posX > 8.5 Then
                                If posX > 17 Then
                                    If posX > 25.5 Then
                                        If posX > 34 Then
                                            If posX > 42.5 Then
                                                sht = 5
                                            Else
                                                sht = 4
                                            End If
                                        Else
                                            sht = 3
                                        End If
                                    Else
                                        sht = 2
                                    End If
                                Else
                                    sht = 1
                                End If
                            Else
                                sht = 0
                            End If

                            Dim adjX As Double = posX - 8.5 * sht
                            Dim bcoll As Long

                            If adjX > 1.6 Then
                                If adjX > 3.2 Then
                                    If adjX > 4.8 Then
                                        If adjX > 6.4 Then
                                            bcoll = 4
                                        Else
                                            bcoll = 3
                                        End If
                                    Else
                                        bcoll = 2
                                    End If
                                Else
                                    bcoll = 1
                                End If
                            Else
                                bcoll = 0
                            End If

                            Dim tY As Double

                            If posY > 0 Then
                                tY = 0
                            Else
                                If posY < -1.5 Then
                                    If posY < -3 Then
                                        If posY < -4.5 Then
                                            If posY < -6 Then
                                                If posY < -7.5 Then
                                                    If posY < -9 Then
                                                    Else
                                                        tY = -9
                                                    End If
                                                Else
                                                    tY = -7.5
                                                End If
                                            Else
                                                tY = -6
                                            End If
                                        Else
                                            tY = -4.5
                                        End If
                                    Else
                                        tY = -3
                                    End If
                                Else
                                    tY = -1.5
                                End If
                            End If

                            Dim tempX As Double

                            Select Case bcoll
                                Case Is = 0
                                    tempX = 0.8
                                Case Is = 1
                                    tempX = 2.4
                                Case Is = 2
                                    tempX = 4
                                Case Is = 3
                                    tempX = 5.6
                                Case Is = 4
                                    tempX = 7.2
                            End Select

                            Dim tX As Double = (sht * 8.5) + tempX

                            Dim mdlSpace As BlockTableRecord = actrans.GetObject(dwgDB.CurrentSpaceId, OpenMode.ForWrite)

                            Using dt As New DBText
                                With dt
                                    .Annotative = 0
                                    .Height = 0.1
                                    .Position = New Point3d(tX, tY, 0)
                                    '.TextStyleName = ltName
                                    .Rotation = 0
                                    .TextString = bName
                                    .WidthFactor = 1
                                    .Justify = AttachmentPoint.BottomCenter
                                    .AlignmentPoint = New Point3d(tX, tY, 0)
                                End With
                                mdlSpace.AppendEntity(dt)
                                actrans.AddNewlyCreatedDBObject(dt, True)
                            End Using
                        End If
                    Next
                    actrans.Commit()
                End Using
            End If
        End Sub


        <CommandMethod("LABLKS", CommandFlags.UsePickSet Or CommandFlags.Redraw Or CommandFlags.Modal)>
        Public Sub LabelBlocks()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim SelResult As PromptSelectionResult = ed.SelectImplied()
            If SelResult.Status = PromptStatus.Error Then
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select blocks references to label:")}
                SelResult = ed.GetSelection(Seloptions)
            Else
                ed.SetImpliedSelection(New ObjectId(-1) {})
            End If

            If SelResult.Status = PromptStatus.OK Then

                Dim acSSet As SelectionSet = SelResult.Value
                Dim MyobjIDs As ObjectId() = acSSet.GetObjectIds

                Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction()
                    'Dim i As Integer = 0
                    For Each objID As ObjectId In MyobjIDs
                        Dim blklist As New List(Of String)
                        Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForRead)
                        If TypeOf dbObj Is BlockReference Then
                            Dim bref As BlockReference = TryCast(dbObj, BlockReference)
                            Dim bName As String = ""
                            If bref.IsDynamicBlock Then
                                Dim bID As ObjectId = bref.DynamicBlockTableRecord
                                Dim dynBlk As BlockTableRecord = actrans.GetObject(bID, OpenMode.ForRead)
                                If dynBlk IsNot Nothing Then bName = dynBlk.Name
                            Else
                                bName = bref.Name
                            End If

                            Dim bPos As Point3d = bref.Position
                            Dim posX As Double = bPos.X
                            Dim posY As Double = bPos.Y

                            Dim sht As Long

                            If posX > 8.5 Then
                                If posX > 17 Then
                                    If posX > 25.5 Then
                                        If posX > 34 Then
                                            If posX > 42.5 Then
                                                sht = 5
                                            Else
                                                sht = 4
                                            End If
                                        Else
                                            sht = 3
                                        End If
                                    Else
                                        sht = 2
                                    End If
                                Else
                                    sht = 1
                                End If
                            Else
                                sht = 0
                            End If

                            Dim adjX As Double = posX - 8.5 * sht
                            Dim bcoll As Long

                            If adjX > 1.6 Then
                                If adjX > 3.2 Then
                                    If adjX > 4.8 Then
                                        If adjX > 6.4 Then
                                            bcoll = 4
                                        Else
                                            bcoll = 3
                                        End If
                                    Else
                                        bcoll = 2
                                    End If
                                Else
                                    bcoll = 1
                                End If
                            Else
                                bcoll = 0
                            End If

                            Dim tY As Double

                            If posY > 0 Then
                                tY = 0
                            Else
                                If posY < -1.5 Then
                                    If posY < -3 Then
                                        If posY < -4.5 Then
                                            If posY < -6 Then
                                                If posY < -7.5 Then
                                                    If posY < -9 Then
                                                    Else
                                                        tY = -9
                                                    End If
                                                Else
                                                    tY = -7.5
                                                End If
                                            Else
                                                tY = -6
                                            End If
                                        Else
                                            tY = -4.5
                                        End If
                                    Else
                                        tY = -3
                                    End If
                                Else
                                    tY = -1.5
                                End If
                            End If

                            Dim tempX As Double

                            Select Case bcoll
                                Case Is = 0
                                    tempX = 0.8
                                Case Is = 1
                                    tempX = 2.4
                                Case Is = 2
                                    tempX = 4
                                Case Is = 3
                                    tempX = 5.6
                                Case Is = 4
                                    tempX = 7.2
                            End Select

                            Dim tX As Double = (sht * 8.5) + tempX

                            Dim mdlSpace As BlockTableRecord = actrans.GetObject(dwgDB.CurrentSpaceId, OpenMode.ForWrite)

                            Using dt As New DBText
                                With dt
                                    .Annotative = 0
                                    .Height = 0.1
                                    .Position = New Point3d(tX, tY, 0)
                                    '.TextStyleName = ltName
                                    .Rotation = 0
                                    .TextString = bName
                                    .WidthFactor = 1
                                    .Justify = AttachmentPoint.BottomCenter
                                    .AlignmentPoint = New Point3d(tX, tY, 0)
                                End With
                                mdlSpace.AppendEntity(dt)
                                actrans.AddNewlyCreatedDBObject(dt, True)
                            End Using
                        End If
                    Next
                    actrans.Commit()
                End Using
            End If
        End Sub

        <CommandMethod("RENBLKS")>
        Public Sub RenameAllBlocks()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim allBlks As Boolean = True

            Dim pko3 As New PromptKeywordOptions(vbLf & "Rename all blocks?")
            With pko3
                .Keywords.Add("Yes")
                .Keywords.Add("No")
                .AllowNone = False
                .AppendKeywordsToMessage = True
                .AllowArbitraryInput = False
            End With

            Dim pkR3 As PromptResult = ed.GetKeywords(pko3)

            If pkR3.Status = PromptStatus.OK Then
                If pkR3.StringResult = "No" Then
                    allBlks = False
                End If
            Else
                Exit Sub
            End If

            Dim bList As New SortedDictionary(Of String, ObjectId)
            Dim blkColl As New Collection

            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForWrite)

                For Each bID As ObjectId In blkTbl
                    Dim btr As BlockTableRecord = acTrans.GetObject(bID, OpenMode.ForRead)
                    If Not btr.IsLayout And Not Left(btr.Name, 1) = "*" Then
                        Dim bName As String = btr.Name
                        Dim btrID As ObjectId = blkTbl(bName)
                        bList.Add(bName, btrID)
                    End If
                Next

                If allBlks Then
                    For Each btrName As String In bList.Keys
                        Dim btr As BlockTableRecord = acTrans.GetObject(blkTbl(btrName), OpenMode.ForWrite)
                        If Not btr.IsLayout And Not Left(btr.Name, 1) = "*" Then
                            blkColl.Add(btr.Name)
                        End If
                    Next
                Else
                    blkColl = PickBlocks()
                End If
                acTrans.Commit()
            End Using

            Dim pko As New PromptKeywordOptions(vbLf & "Add prefix/suffix to existing block names or replace string?")
            With pko
                .Keywords.Add("Prefix")
                .Keywords.Add("Suffix")
                .Keywords.Add("Replace")
                .AllowNone = False
                .AppendKeywordsToMessage = True
                .AllowArbitraryInput = False
            End With

            Dim pkR As PromptResult = ed.GetKeywords(pko)

            Dim usePrefix As Boolean = False
            Dim repStr As Boolean = False

            If pkR.Status = PromptStatus.OK Then
                If pkR.StringResult = "Prefix" Then
                    usePrefix = True
                ElseIf pkR.StringResult = "Replace" Then
                    repStr = True
                End If
            Else
                Exit Sub
            End If

            Dim oldStr As String = ""

            If repStr Then

                Dim pso0 As New PromptStringOptions(vbLf & "Enter string to replace in block names: ") With {.AllowSpaces = True}

                Dim psr0 As PromptResult = ed.GetString(pso0)

                If psr0.Status = PromptStatus.OK Then
                    oldStr = psr0.StringResult
                    If String.IsNullOrEmpty(oldStr) Then Exit Sub
                Else
                    Exit Sub
                End If

            End If

            Dim pso As New PromptStringOptions(vbLf & "Enter string to add to block names: ")
            With pso
                pso.AllowSpaces = True
                If repStr Then .Message = vbLf & "Enter new string:"
            End With

            Dim addStr As String

            Dim psr As PromptResult = ed.GetString(pso)

            If psr.Status = PromptStatus.OK Then
                addStr = psr.StringResult
                If String.IsNullOrEmpty(addStr) Then Exit Sub
            Else
                Exit Sub
            End If

            Dim pko2 As New PromptKeywordOptions(vbLf & "Confirm renaming for each block?")
            With pko2
                .Keywords.Add("Yes")
                .Keywords.Add("No")
                .AllowNone = False
                .AppendKeywordsToMessage = True
                .AllowArbitraryInput = False
            End With

            Dim pkR2 As PromptResult = ed.GetKeywords(pko2)

            Dim checkNames As Boolean = False

            If pkR2.Status = PromptStatus.OK Then
                If pkR2.StringResult = "Yes" Then
                    checkNames = True
                End If
            Else
                Exit Sub
            End If

            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForWrite)

                For Each bn As String In blkColl
                    Dim btr As BlockTableRecord = acTrans.GetObject(blkTbl(bn), OpenMode.ForRead)
                    If Not btr.IsLayout And Not Left(btr.Name, 1) = "*" Then
                        btr.UpgradeOpen()
                        Dim oldName As String = btr.Name
                        Dim newName As String

                        If usePrefix Then
                            newName = addStr & oldName
                        ElseIf repStr And Not usePrefix Then
                            newName = Strings.Replace(oldName, oldStr, addStr, 1, 1)
                        Else
                            newName = oldName & addStr
                        End If
                        If Not oldName = newName Then
                            If checkNames Then
                                Dim newMsg As String = vbLf & "Change block name from " & oldName & " to " & newName & "?"
                                pko2.Message = newMsg
                                Dim pkr4 As PromptResult = ed.GetKeywords(pko2)
                                If pkr4.Status = PromptStatus.OK Then
                                    If pkr4.StringResult = "Yes" Then
                                        RenameBlock(oldName, newName)
                                    End If
                                Else
                                    Exit Sub
                                End If
                            Else
                                RenameBlock(oldName, newName)
                            End If
                        End If
                    End If
                Next
                acTrans.Commit()
            End Using
        End Sub

        <CommandMethod("CBTZ", CommandFlags.UsePickSet Or CommandFlags.Redraw Or CommandFlags.Modal)>
        Public Sub ChangeBlocksToLayerZero()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim SelResult As PromptSelectionResult = ed.SelectImplied()
            If SelResult.Status = PromptStatus.Error Then
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select blocks to change subentity layers to zero:")}
                SelResult = ed.GetSelection(Seloptions)
            Else
                ed.SetImpliedSelection(New ObjectId(-1) {})
            End If

            If SelResult.Status = PromptStatus.OK Then

                Dim acSSet As SelectionSet = SelResult.Value
                Dim MyobjIDs As ObjectId() = acSSet.GetObjectIds

                Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction()
                    Dim blkTbl As BlockTable = actrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                    Dim ltbl As LayerTable = actrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)

                    'Dim i As Integer = 0
                    For Each objID As ObjectId In MyobjIDs
                        Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForRead)
                        Dim blkList As New List(Of String)

                        If TypeOf dbObj Is BlockReference Then
                            Dim parentBref As BlockReference = TryCast(dbObj, BlockReference)
                            Dim parentName As String = parentBref.Name
                            Dim ParentObjID As ObjectId = blkTbl(parentName)
                            Dim ParentBTR As BlockTableRecord = actrans.GetObject(ParentObjID, OpenMode.ForWrite)
                            Dim sb As New StringBuilder
                            For Each obID As ObjectId In ParentBTR
                                Dim dObj As DBObject = actrans.GetObject(obID, OpenMode.ForRead)
                                If TypeOf dObj IsNot BlockReference Then
                                    If TypeOf dObj Is Entity Then
                                        Dim ent As Entity = CType(dObj, Entity)
                                        If Not ent.LayerId = dwgDB.LayerZero Then
                                            ent.UpgradeOpen()
                                            Dim entlay As String = ent.Layer
                                            Dim entlayID As ObjectId = ent.LayerId
                                            Dim layTblRec As LayerTableRecord = actrans.GetObject(entlayID, OpenMode.ForRead)
                                            If ent.Color.IsByLayer Then
                                                ent.Color = layTblRec.Color
                                                If ent.PlotStyleName = "ByLayer" Then
                                                    ent.PlotStyleName = layTblRec.PlotStyleName
                                                Else
                                                    SetEntityPlotStyle(ent.ObjectId, "Normal")
                                                End If
                                            End If
                                            If ent.Linetype = "ByLayer" Then ent.LinetypeId = layTblRec.LinetypeObjectId
                                            ent.LayerId = dwgDB.LayerZero
                                        Else
                                        End If
                                    End If
                                Else
                                    Dim thisBr As BlockReference = CType(dObj, BlockReference)
                                    blkList.Add(thisBr.Name)
                                End If
                            Next

                            Dim fnlBlkLst As IEnumerable(Of String) = blkList.Distinct
                            If fnlBlkLst.Count > 0 Then
                                'sb.AppendLine(blkList(0))
                                For m As Integer = 0 To fnlBlkLst.Count - 1
                                    sb.AppendLine(fnlBlkLst(m))
                                Next
                                'Dim blkList2 As New List(Of String)
                                If Not String.IsNullOrEmpty(sb.ToString) Then MessageBox.Show("Block definition for " & parentName & " contains the following unprocessed blocks:" & vbLf & sb.ToString)
                            End If
                            parentBref.RecordGraphicsModified(True)
                        End If
                    Next
                    actrans.Commit()
                    ed.Regen()

                End Using
            End If
            ed.WriteMessage(vbLf & "All blocks updated.")

        End Sub

        <CommandMethod("CBBL", CommandFlags.UsePickSet Or CommandFlags.Redraw Or CommandFlags.Modal)>
        Public Sub ChangeBlocksToByLayer()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim SelResult As PromptSelectionResult = ed.SelectImplied()
            If SelResult.Status = PromptStatus.Error Then
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select blocks to change subenty color to ByLayer:")}
                SelResult = ed.GetSelection(Seloptions)
            Else
                ed.SetImpliedSelection(New ObjectId(-1) {})
            End If

            If SelResult.Status = PromptStatus.OK Then

                Dim acSSet As SelectionSet = SelResult.Value
                Dim MyobjIDs As ObjectId() = acSSet.GetObjectIds

                Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction()
                    Dim blkTbl As BlockTable = actrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                    Dim ltbl As LayerTable = actrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)

                    'Dim i As Integer = 0
                    For Each objID As ObjectId In MyobjIDs
                        Dim blklist As New List(Of String)
                        Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForRead)
                        If TypeOf dbObj Is BlockReference Then
                            Dim parentBref As BlockReference = TryCast(dbObj, BlockReference)
                            Dim parentName As String = parentBref.Name
                            Dim ParentObjID As ObjectId = blkTbl(parentName)
                            Dim ParentBTR As BlockTableRecord = actrans.GetObject(ParentObjID, OpenMode.ForWrite)
                            Dim sb As New StringBuilder
                            For Each obID As ObjectId In ParentBTR
                                Dim dObj As DBObject = actrans.GetObject(obID, OpenMode.ForRead)
                                If TypeOf dObj IsNot BlockReference Then
                                    If TypeOf dObj Is Entity Then
                                        Dim ent As Entity = CType(dObj, Entity)
                                        ent.UpgradeOpen()
                                        'Dim entlay As String = ent.Layer
                                        If Not ent.Color.IsByLayer Then ent.Color = Color.FromColorIndex(ColorMethod.ByAci, 256)
                                        If Not ent.PlotStyleName = "ByLayer" Then ent.PlotStyleName = "ByLayer"
                                    End If
                                Else
                                    Dim thisBr As BlockReference = CType(dObj, BlockReference)
                                    blklist.Add(thisBr.Name)
                                End If
                            Next

                            Dim fnlBlkLst As IEnumerable(Of String) = blklist.Distinct

                            If fnlBlkLst.Count > 0 Then
                                'sb.AppendLine(blkList(0))
                                For m As Integer = 0 To fnlBlkLst.Count - 1
                                    sb.AppendLine(fnlBlkLst(m))
                                Next
                                'Dim blkList2 As New List(Of String)
                                If Not String.IsNullOrEmpty(sb.ToString) Then MessageBox.Show("Block definition for " & parentName & " contains the following unprocessed block references:" & vbLf & sb.ToString)
                            End If
                        End If
                    Next
                    actrans.Commit()
                End Using
                ed.Regen()
            End If

            ed.WriteMessage(vbLf & "All blocks updated.")

        End Sub

        <CommandMethod("CPSBL", CommandFlags.UsePickSet Or CommandFlags.Redraw Or CommandFlags.Modal)>
        Public Sub ChangeBlocksPstylesByLayer()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim SelResult As PromptSelectionResult = ed.SelectImplied()
            If SelResult.Status = PromptStatus.Error Then
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select blocks to change subenty color to ByLayer:")}
                SelResult = ed.GetSelection(Seloptions)
            Else
                ed.SetImpliedSelection(New ObjectId(-1) {})
            End If

            If SelResult.Status = PromptStatus.OK Then

                Dim acSSet As SelectionSet = SelResult.Value
                Dim MyobjIDs As ObjectId() = acSSet.GetObjectIds

                Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction()
                    Dim blkTbl As BlockTable = actrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                    Dim ltbl As LayerTable = actrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)

                    'Dim i As Integer = 0
                    For Each objID As ObjectId In MyobjIDs
                        Dim blklist As New List(Of String)
                        Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForRead)
                        If TypeOf dbObj Is BlockReference Then
                            Dim parentBref As BlockReference = TryCast(dbObj, BlockReference)
                            Dim parentName As String = parentBref.Name
                            Dim ParentObjID As ObjectId = blkTbl(parentName)
                            Dim ParentBTR As BlockTableRecord = actrans.GetObject(ParentObjID, OpenMode.ForWrite)
                            Dim sb As New StringBuilder
                            For Each obID As ObjectId In ParentBTR
                                Dim dObj As DBObject = actrans.GetObject(obID, OpenMode.ForRead)
                                If TypeOf dObj IsNot BlockReference Then
                                    If TypeOf dObj Is Entity Then
                                        Dim ent As Entity = CType(dObj, Entity)
                                        ent.UpgradeOpen()
                                        'Dim entlay As String = ent.Layer
                                        'If Not ent.Color.IsByLayer Then ent.Color = Color.FromColorIndex(ColorMethod.ByAci, 256)
                                        If Not ent.PlotStyleName = "ByLayer" Then ent.PlotStyleName = "ByLayer"
                                    End If
                                Else
                                    Dim thisBr As BlockReference = CType(dObj, BlockReference)
                                    blklist.Add(thisBr.Name)
                                End If
                            Next

                            Dim fnlBlkLst As IEnumerable(Of String) = blklist.Distinct

                            If fnlBlkLst.Count > 0 Then
                                'sb.AppendLine(blkList(0))
                                For m As Integer = 0 To fnlBlkLst.Count - 1
                                    sb.AppendLine(fnlBlkLst(m))
                                Next
                                'Dim blkList2 As New List(Of String)
                                If Not String.IsNullOrEmpty(sb.ToString) Then MessageBox.Show("Block definition for " & parentName & " contains the following unprocessed block references:" & vbLf & sb.ToString)
                            End If
                        End If
                    Next
                    actrans.Commit()
                End Using
                ed.Regen()
            End If

            ed.WriteMessage(vbLf & "All blocks updated.")

        End Sub


        <CommandMethod("SetDwgsBase")>
        Public Sub SetDwgsBase()

            Dim filepath As String = ""
            Dim blkFolder As String
            Dim folderpicker As New FolderBrowserDialog
            With folderpicker
                .Description = "Select folder containing drawing files to change."
                .ShowNewFolderButton = False
            End With

            If folderpicker.ShowDialog = DialogResult.OK Then
                filepath = folderpicker.SelectedPath
                blkFolder = folderpicker.SelectedPath
            End If

skipit:
            Dim uR2 As Integer = MsgBox("Recurse subfolders?", vbYesNoCancel)
            Dim sO As New FileIO.SearchOption
            If uR2 = vbYes Then
                sO = Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories
            ElseIf uR2 = vbYes Then
                sO = Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly
            Else
                Exit Sub
            End If

            If filepath <> "" Then
                For Each foundfile As String In My.Computer.FileSystem.GetFiles(filepath, sO, "*.dwg")
                    ChangeBase(foundfile)
                Next
            End If
        End Sub

        <CommandMethod("BLKDATA")>
        Public Sub ShowBlkData()

            Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = acDwg.Database
            Dim ed As Editor = acDwg.Editor

            Dim blkCol As New AcBlocks
            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction

                Dim blktbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                Dim curSpace As BlockTableRecord = acTrans.GetObject(dwgDB.CurrentSpaceId, OpenMode.ForRead)
                Dim obIds As New ObjectIdCollection

                For Each obid In curSpace
                    Dim myob As DBObject = acTrans.GetObject(obid, OpenMode.ForRead)
                    If TypeOf myob Is BlockReference Then
                        obIds.Add(obid)
                    End If
                Next

                Dim upperlimit As Long = obIds.Count

                Dim blkIds As New ObjectIdCollection
                Dim rawBlkLst As New List(Of String)

                For Each id As ObjectId In obIds
                    Dim dbOb As DBObject = acTrans.GetObject(id, OpenMode.ForRead)
                    If TypeOf dbOb Is BlockReference Then
                        blkIds.Add(id)
                        Dim tbref As BlockReference = TryCast(dbOb, BlockReference)
                        If tbref IsNot Nothing Then
                            Dim bi As New BlockInfo(tbref)
                            rawBlkLst.Add(bi.RefName)
                            blkCol.Add(bi)
                        End If
                    End If
                Next

                Dim refCount As Integer = 0

                For Each blkdata As BlockInfo In blkCol
                    Dim nameStr = blkdata.RefName
                    For Each testStr As String In rawBlkLst
                        If nameStr = testStr Then refCount += 1
                    Next
                    blkdata.RefCount = refCount
                    refCount = 0
                Next

                blkCol.Sort()
                acTrans.Commit()

            End Using

            Dim tmpXmlName As String = Path.GetTempFileName()
            'Dim tmpXmlName As String = "H:\VS Repos\MasterCustomLbry\BlockList.xml"

            Dim sets As New XmlWriterSettings
            With sets
                .ConformanceLevel = ConformanceLevel.Document
                .Encoding = Encoding.UTF8
                .Indent = True
                .NewLineOnAttributes = True
                .IndentChars = "   "
            End With

            Using xw As XmlWriter = XmlWriter.Create(tmpXmlName, sets)
                Dim xmlS As New XmlSerializer(GetType(AcBlocks), "http://tempuri.org/BlkInfoSchema.xsd")
                xmlS.Serialize(xw, blkCol)
                xw.Flush()
            End Using

            Dim bvp As New BlkViewPanel With {.XmlFileName = tmpXmlName}

            bvp.CreateDGVxml()
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(bvp)

            'Kill(bvp.xmlFileName)
            'End Using

        End Sub



        <CommandMethod("INSALL")>
        Public Sub InsertAllDwgFilesInFolder()

            Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = acDwg.Database
            Dim ed As Editor = acDwg.Editor

            ed.WriteMessage(vbLf & "This command will insert all drawing files from a designated folder.")

            Dim xDist As Double
            Dim yDist As Double

            Dim ptOpts As New PromptPointOptions(vbLf & "Select 2 points for size of each block cell. First point: ")
            Dim ptRes As PromptPointResult = ed.GetPoint(ptOpts)

            If ptRes.Status = PromptStatus.OK Then
                Dim pt1 As Point3d = ptRes.Value
                Dim ptopts2 As New PromptPointOptions(vbLf & "Second point: ")
                With ptopts2
                    .BasePoint = pt1
                    .UseBasePoint = True
                    .UseDashedLine = True
                End With
                Dim ptRes2 As PromptPointResult = ed.GetPoint(ptopts2)

                If ptRes2.Status = PromptStatus.OK Then
                    Dim pt2 As Point3d = ptRes2.Value
                    xDist = Abs(pt2.X - pt1.X)
                    yDist = Abs(pt2.Y - pt1.Y)
                    GoTo SkipIt
                Else
                    ed.WriteMessage(vbLf & "Command Canceled.")
                    Exit Sub
                End If
            End If

            Dim pdo As New PromptDoubleOptions(vbLf & "Enter width of block cell in dwg units: ")
            With pdo
                .AllowZero = False
            End With
            Dim pdoRes As PromptDoubleResult = ed.GetDouble(pdo)
            If pdoRes.Status = PromptStatus.OK Then
                xDist = pdoRes.Value
                Dim pdo2 As New PromptDoubleOptions(vbLf & "Enter height of block cell in dwg units: ")
                With pdo2
                    .AllowZero = False
                End With
                Dim pdo2Res As PromptDoubleResult = ed.GetDouble(pdo2)
                If pdo2Res.Status = PromptStatus.OK Then
                    yDist = pdo2Res.Value
                End If
            Else
                ed.WriteMessage(vbLf & "Command Canceled.")
                Exit Sub
            End If

SkipIt:

            Dim colNo As Integer

            Dim pdo3 As New PromptIntegerOptions(vbLf & "How many columns in each row?")
            With pdo3
                .AllowZero = False
            End With
            Dim pdo3Res As PromptIntegerResult = ed.GetInteger(pdo3)
            If pdo3Res.Status = PromptStatus.OK Then
                colNo = pdo3Res.Value
            Else
                ed.WriteMessage(vbLf & "Command Canceled.")
                Exit Sub
            End If


            If xDist = 0 OrElse yDist = 0 Then
                ed.WriteMessage(vbLf & "No cell size entered.  Command Canceled.")
                Exit Sub
            End If

            Dim blkFldr As String = GetMyFolderName()
            If String.IsNullOrEmpty(blkFldr) Then Exit Sub

            Dim files() As FileInfo = AllDwgFilesInFolder(blkFldr, False)
            Dim pathList As New Dictionary(Of String, String)

            For Each fi As FileInfo In files
                pathList.Add(Path.GetFileNameWithoutExtension(fi.Name), fi.FullName)
            Next

            Dim initialInsPtx As Double = 0.5 * xDist
            Dim insPty As Double = 0.5 * yDist
            Dim insptX As Double = 0.5 * xDist
            Dim insPt As Point3d

            Dim i As Integer = 0
            Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction

                Dim blkTbl As BlockTable = actrans.GetObject(dwgDB.BlockTableId, OpenMode.ForWrite)
                Dim cSpace As BlockTableRecord = actrans.GetObject(dwgDB.CurrentSpaceId, OpenMode.ForWrite)

                For Each ky As String In pathList.Keys

                    Try
                        Dim btrID As ObjectId = InsertDwg(pathList(ky), ky)
                        insptX = initialInsPtx + (i * xDist)

                        insPt = New Point3d(insptX, insPty, 0)

                        Dim bDef As BlockTableRecord = actrans.GetObject(btrID, OpenMode.ForRead)
                        Dim bRef As New BlockReference(insPt, btrID)
                        cSpace.AppendEntity(bRef)
                        actrans.AddNewlyCreatedDBObject(bRef, True)

                        If bDef.HasAttributeDefinitions Then
                            For Each id As ObjectId In bDef
                                Dim dbOb As DBObject = actrans.GetObject(id, OpenMode.ForRead)
                                If TypeOf dbOb Is AttributeDefinition Then
                                    Dim attdef As AttributeDefinition = dbOb
                                    Using attref As New AttributeReference
                                        attref.SetAttributeFromBlock(attdef, bRef.BlockTransform)
                                        attref.Position = attdef.Position.TransformBy(bRef.BlockTransform)
                                        attref.TextString = attref.Tag.ToString
                                        bRef.AttributeCollection.AppendAttribute(attref)
                                        actrans.AddNewlyCreatedDBObject(attref, True)
                                    End Using
                                End If
                            Next
                        End If

                        If i = colNo - 1 Then
                            i = 0
                            insPty -= yDist
                        Else
                            i += 1
                        End If

                    Catch ex As Exception
                        ed.WriteMessage(vbLf & "Error in InsertAllDwgFilesInFolder function.")
                        Exit Sub
                    End Try
                Next
                actrans.Commit()
            End Using

        End Sub

        <CommandMethod("THUMBS", CommandFlags.Session)>
        Public Sub CreateThumbnails()
            Dim docMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim curDwg As Document = docMgr.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            'Dim ed As Editor = curDwg.Editor

            Dim blkFldr As String = GetMyFolderName()
            If String.IsNullOrEmpty(blkFldr) Then Exit Sub

            Dim files() As FileInfo = AllDwgFilesInFolder(blkFldr, False)

            Dim pathList As New Dictionary(Of String, String)

            Dim pckColl As New List(Of String)

            For Each fi As FileInfo In files
                pathList.Add(fi.Name, fi.FullName)
                pckColl.Add(fi.Name)
            Next

            Dim pckr As New Picker

            For Each nm As String In pckColl
                pckr.BxList.Items.Add(nm)
            Next

            With pckr
                .TopLabel.Text = "Pick DWG files to update thumbnail"
                .Text = "BLock Picker"
            End With

            Dim fileList As Collection

            pckr.ShowDialog()

            If pckr.DialogResult = DialogResult.OK Then
                fileList = pckr.PickCol
            Else
                pckr.Dispose()
                Exit Sub
            End If

            For Each ky As String In fileList
                'Dim lspName As String = pathList(ky).Replace("\", "/")
                Dim tempDoc As Document = DocumentCollectionExtension.Add(docMgr, pathList(ky))

                Using tempLock As DocumentLock = tempDoc.LockDocument
                    docMgr.MdiActiveDocument = tempDoc
                    Dim ed2 As Editor = tempDoc.Editor
                    Dim tempDb As Database = tempDoc.Database

                    Dim pMin As Point3d
                    Dim pMax As Point3d
                    'If tempDb.TileMode = True Then
                    pMin = tempDb.Extmin
                    pMax = tempDb.Extmax
                    'End If

                    Using acView As ViewTableRecord = ed2.GetCurrentView()
                        Dim eExtents As Extents3d
                        '' Translate WCS coordinates to DCS
                        Dim matWCS2DCS As Matrix3d
                        matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection)
                        matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) * matWCS2DCS
                        matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist, acView.ViewDirection, acView.Target) * matWCS2DCS

                        Using acLine As New Line(pMin, pMax)
                            eExtents = New Extents3d(acLine.Bounds.Value.MinPoint, acLine.Bounds.Value.MaxPoint)
                        End Using
                        '' Calculate the ratio between the width and height of the current view
                        Dim dViewRatio As Double = (acView.Width / acView.Height)
                        '' Tranform the extents of the view
                        matWCS2DCS = matWCS2DCS.Inverse()
                        eExtents.TransformBy(matWCS2DCS)

                        Dim dWidth As Double
                        Dim dHeight As Double
                        Dim pNewCentPt As Point2d

                        '' Calculate the new width and height of the current view
                        dWidth = eExtents.MaxPoint.X - eExtents.MinPoint.X
                        dHeight = eExtents.MaxPoint.Y - eExtents.MinPoint.Y
                        '' Get the center of the view
                        pNewCentPt = New Point2d(((eExtents.MaxPoint.X + eExtents.MinPoint.X) * 0.5), ((eExtents.MaxPoint.Y + eExtents.MinPoint.Y) * 0.5))

                        'dHeight = dWidth / dViewRatio
                        If dWidth >= dHeight Then
                            dHeight = dWidth / dViewRatio
                        Else
                            dWidth = dHeight / dViewRatio
                        End If
                        'dHeight = dWidth * dViewRatio

                        '' Resize and scale the view
                        acView.Width = dWidth
                        acView.Height = dHeight
                        '' Set the center of the view
                        acView.CenterPoint = pNewCentPt
                        '' Set the current view
                        ed2.SetCurrentView(acView)
                    End Using

                    Dim bm As Drawing.Bitmap = tempDoc.CapturePreviewImage(800, 600)
                    tempDb.ThumbnailBitmap = bm

                    tempDb.SaveAs(pathList(ky), True, DwgVersion.AC1027, Nothing)

                    'Dim ocmd As Object = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CMDECHO")
                    'tempDoc.SendStringToExecute("(setvar ""CMDECHO"" 0)" & "(command ""_.SAVEAS"" """" """ & lspName & """)" & "(setvar ""CMDECHO"" " + ocmd.ToString() & ")" & "(princ) ", False, False, False)
                    'tempDoc.SendStringToExecute("(command ""_.SAVEAS"" """" """ & lspName & """)" & "(princ) ", False, False, True)
                    'tempDoc.SendStringToExecute("(command ""_.SAVE "" & "")" & "(princ) ", False, False, True)
                    'tempDoc.SendStringToExecute("(SETVAR \" & ChrW(34) & "CMDECHO\" & ChrW(34) & "0) ") & "(command \" & chrw(34) & "_.SAVEAS\" & chrw(34) & " \" & chrw(34) & "\ " & chrw(34) 

                End Using
                tempDoc.CloseAndDiscard
            Next

        End Sub

        <CommandMethod("PATF")>
        Public Sub PurgeAllDwgFilesInFolder()
            Dim acDwgMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim curDwg As Document = acDwgMgr.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            ed.WriteMessage(vbLf & "This command will purge non-standard textstyles from all drawing files In a designated folder.")

            Dim blkFldr As String = GetMyFolderName()
            'Dim blkfldr As String = "\\EESServer\datadisk\CAD\BLOCKS\ROAD\PVMT\Design"
            If String.IsNullOrEmpty(blkFldr) Then Exit Sub
            If Not Directory.Exists(blkFldr) Then Exit Sub

            Dim files() As FileInfo = AllDwgFilesInFolder(blkFldr, False)
            Dim pathList As New Dictionary(Of String, String)

            For Each fi As FileInfo In files
                Dim dwgName As String = Path.GetFileNameWithoutExtension(fi.Name)
                If Not InStr(dwgName.ToUpper, "MASTER") Or Not InStr(dwgName.ToUpper, "DETAIL") Then
                    pathList.Add(dwgName, fi.FullName)
                Else
                    ed.WriteMessage(vbLf & "Dwg " & dwgName & " Not processed.")
                    Dim cont As Boolean = YesNoQuery(vbLf & "Continue processing?")
                    If Not cont Then Exit Sub
                End If
            Next

            For Each ky As String In pathList.Keys
                Try
                    Dim acDoc As Document = acDwgMgr.Open(pathList(ky), False)
                    Using actrans As Transaction = acDoc.TransactionManager.StartTransaction
                        Using docLock As DocumentLock = acDoc.LockDocument
                            Dim acdb As Database = acDoc.Database
                            Dim tst As TextStyleTable = actrans.GetObject(acdb.TextStyleTableId, OpenMode.ForWrite)
                            Dim cStyleID As ObjectId = acdb.Textstyle
                            Dim cStyle As TextStyleTableRecord = actrans.GetObject(cStyleID, OpenMode.ForRead)
                            If Not cStyle.Name = "Standard" Then
                                Dim tempID As ObjectId = tst("Standard")
                                If Not tempID.IsNull Then acdb.Textstyle = tempID
                            End If

                            For Each stId As ObjectId In tst
                                Dim prgList As New ObjectIdCollection
                                Dim myStyle As TextStyleTableRecord = actrans.GetObject(stId, OpenMode.ForWrite)
                                If Not myStyle.Name = "Standard" Then
                                    prgList.Add(stId)
                                    acdb.Purge(prgList)
                                    Try
                                        myStyle.Erase(True)
                                    Catch ex As Exception
                                        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Error:" & vbLf & ex.Message)
                                    End Try
                                End If
                            Next
                        End Using

                        actrans.Commit()

                    End Using

                    acDoc.CloseAndSave(pathList(ky))

                Catch ex As Exception
                End Try
            Next

        End Sub


        <CommandMethod("WBTF")>
        Public Sub WriteBLocksToFolder()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            Dim sName As String = curDwg.Name
            Dim savef As String = Path.GetDirectoryName(sName) '& "\testFldr\"

            Dim dr As DialogResult = MessageBox.Show("This command will write all user blocks in this drawing to a directory.  Do you want to continue?", "Alert", MessageBoxButtons.YesNo)
            If dr = DialogResult.No Then Exit Sub

            Dim saveFldr As String = GetMyFolderName(savef)

            If String.IsNullOrEmpty(saveFldr) Then Exit Sub

            If Not Directory.Exists(saveFldr) Then
                Directory.CreateDirectory(saveFldr)
            End If

            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim blktbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                For Each bID As ObjectId In blktbl
                    Dim btr As BlockTableRecord = acTrans.GetObject(bID, OpenMode.ForRead)
                    'Dim fname1 As String = Path
                    Dim fName As String
                    If Not btr.IsLayout Then
                        If Not String.IsNullOrEmpty(btr.Name) And Left(btr.Name, 1) <> "*" And Left(btr.Name, 1) <> "_" And Left(btr.Name, 2) <> "A$" Then
                            fName = saveFldr & "\" & btr.Name & ".dwg"
                            'Dim tdb As new Database(False, True)
                            Dim tdb As Database
                            tdb = dwgDB.Wblock(btr.ObjectId)
                            tdb.SaveAs(fName, Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1027)
                            ed.WriteMessage(vbLf & btr.Name & " written to output folder")
                        End If
                    End If
                Next
                acTrans.Commit()
            End Using
            ed.WriteMessage(vbLf & "All blocks saved to designated folder")
        End Sub

        <CommandMethod("CBU")>
        Public Sub ChangeBlockUnits()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            Dim sName As String = curDwg.Name
            Dim savef As String = Path.GetDirectoryName(sName) '& "\testFldr\"

            Dim dr As DialogResult = MessageBox.Show("This command will change the units of all selected blocks.  Do you want to continue?", "Alert", MessageBoxButtons.YesNo)
            If dr = DialogResult.No Then Exit Sub

            Dim pko As New PromptKeywordOptions("Select the new units:")
            With pko
                .Keywords.Add("Feet")
                .Keywords.Add("Inches")
                .Keywords.Add("Meters")
                .Keywords.Add("Millimeters")
                .Keywords.Add("Undefined")
                .AppendKeywordsToMessage = True
                .AllowNone = False
            End With

            Dim pkr As PromptResult = ed.GetKeywords(pko)

            Dim kw As String
            If pkr.Status = PromptStatus.OK Then
                kw = pkr.StringResult
            Else
                Exit Sub
            End If

            Dim uv As UnitsValue
            Select Case kw
                Case Is = "Feet"
                    uv = UnitsValue.Feet
                Case Is = "Inches"
                    uv = UnitsValue.Inches
                Case Is = "Meters"
                    uv = UnitsValue.Meters
                Case Is = "Millimeters"
                    uv = UnitsValue.Millimeters
                Case Is = "Undefined"
                    uv = UnitsValue.Undefined
            End Select


            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim blktbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                Dim bList As New SortedDictionary(Of String, ObjectId)

                For Each bID As ObjectId In blktbl
                    Dim btr As BlockTableRecord = acTrans.GetObject(bID, OpenMode.ForRead)
                    If Not btr.IsLayout Then
                        Dim bName As String = btr.Name
                        Dim btrID As ObjectId = blktbl(bName)
                        bList.Add(bName, btrID)
                    End If
                Next

                Dim bPicker As New Picker
                With bPicker
                    .BxList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
                    .TopLabel.Text = "Select blocks to update"
                    .Text = "BLock Picker"
                    For Each blkNm As String In bList.Keys
                        .BxList.Items.Add(blkNm)
                    Next
                End With

                bPicker.ShowDialog()

                Dim blkColl As Collection

                If bPicker.DialogResult = DialogResult.Cancel Then
                    Exit Sub
                Else
                    blkColl = bPicker.PickCol
                End If

                Dim selBlks As New SortedDictionary(Of String, ObjectId)
                For Each blk As String In blkColl
                    selBlks.Add(blk, bList(blk))
                Next

                'dim dv As DwgVersion
                'dv = DwgVersion(curDwg.Name)

                For Each blkStr As String In selBlks.Keys
                    Using btr As BlockTableRecord = acTrans.GetObject(selBlks(blkStr), OpenMode.ForWrite)
                        btr.Units = uv
                    End Using
                Next
                acTrans.Commit()
            End Using
            ed.WriteMessage(vbLf & "All blocks updated")
        End Sub

        <CommandMethod("CPMLT", CommandFlags.UsePickSet Or CommandFlags.Redraw Or CommandFlags.Modal)>
        Public Sub ChangePvmntMarkingsLineTypes()
            'by David Eisenbeisz (c)2023

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction()
                Dim blkTbl As BlockTable = actrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                'Dim acSSet As SelectionSet = SelResult.Value
                'Dim MyobjIDs As ObjectId() = acSSet.GetObjectIds

                'Dim i As Integer = 0
                For Each btrID As ObjectId In blkTbl
                    Dim bTR As BlockTableRecord = actrans.GetObject(btrID, OpenMode.ForRead)
                    If Not bTR.IsLayout Then
                        For Each subID As ObjectId In bTR
                            Dim dbObj As DBObject = actrans.GetObject(subID, OpenMode.ForRead)
                            If TypeOf dbObj Is Entity Then
                                Dim ent As Entity = TryCast(dbObj, Entity)
                                If ent IsNot Nothing Then
                                    If ent.Linetype = "PMRem" Then
                                        ent.UpgradeOpen()
                                        ent.Linetype = "PMREMOVE"
                                        ent.LinetypeScale = 1
                                        If TypeOf ent Is Polyline Then
                                            Dim pL As Polyline = TryCast(ent, Polyline)
                                            If pL IsNot Nothing Then pL.Plinegen = True
                                        End If
                                    ElseIf ent.Linetype = "HIDDEN2" Then
                                        ent.UpgradeOpen()
                                        ent.Linetype = "PMREMOVE"
                                        ent.LinetypeScale = 1
                                        If TypeOf ent Is Polyline Then
                                            Dim pL As Polyline = TryCast(ent, Polyline)
                                            If pL IsNot Nothing Then pL.Plinegen = True
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                Next
                actrans.Commit()
            End Using
        End Sub


        <CommandMethod("CPM", CommandFlags.UsePickSet Or CommandFlags.Redraw Or CommandFlags.Modal)>
        Public Sub ChangePvmntMarkings()
            'by David Eisenbeisz (c)2023

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Dim SelResult As PromptSelectionResult = ed.SelectImplied()
            If SelResult.Status = PromptStatus.Error Then
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select dynamic pavment marking blocks to change:")}
                SelResult = ed.GetSelection(Seloptions)
            Else
                ed.SetImpliedSelection(New ObjectId(-1) {})
            End If

            If SelResult.Status = PromptStatus.OK Then

                Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction()

                    Dim blkTbl As BlockTable = actrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    Dim acSSet As SelectionSet = SelResult.Value
                    Dim MyobjIDs As ObjectId() = acSSet.GetObjectIds

                    'Dim i As Integer = 0
                    For Each objID As ObjectId In MyobjIDs
                        Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForRead)
                        If TypeOf dbObj Is BlockReference Then
                            Dim parentBref As BlockReference = TryCast(dbObj, BlockReference)
                            Dim parentName As String = parentBref.Name
                            Dim parentPos As Point3d = parentBref.Position
                            Dim parentOrient As Double = parentBref.Rotation
                            Dim parentScale As Scale3d = parentBref.ScaleFactors
                            Dim ParentObjID As ObjectId = blkTbl(parentName)
                            Dim ParentBTR As BlockTableRecord = actrans.GetObject(ParentObjID, OpenMode.ForRead)
                            Dim baseName As String = ""
                            Dim newName As String = ""
                            Dim curState As String = ""
                            Dim newState As String = ""
                            Dim sepChar As String = ""
                            If InStr(parentName, "Existing") > 0 Or InStr(parentName, "Proposed") > 0 Or InStr(parentName, "Remove") > 0 Then
                                'Dim nameParts() As String
                                If InStr(parentName, "_") > 0 Then
                                    sepChar = "_"
                                ElseIf InStr(parentName, " ") > 0 Then
                                    sepChar = " "
                                End If

                                If InStr(parentName, "_") Or InStr(parentName, " ") Then
                                    Dim nameParts = Split(parentName, sepChar)
                                    If nameParts.Count > 2 Then
                                        Dim tempName(nameParts.Count - 1) As String
                                        For j = 0 To nameParts.Count - 2
                                            ReDim Preserve tempName(j + 1)
                                            tempName(j) = nameParts(j)
                                        Next
                                        baseName = Join(tempName, sepChar)
                                        curState = nameParts(nameParts.Count - 1)
                                    ElseIf nameParts.Count = 2 Then
                                        Dim tempname() As String = Split(sepChar)
                                        curState = nameParts(1)
                                        baseName = nameParts(0)
                                    ElseIf nameParts.Count < 2 Then
                                        baseName = nameParts(0)
                                        curState = "Existing"
                                    End If
                                Else
                                    baseName = parentName
                                    curState = "Existing"
                                End If

                                If Not String.IsNullOrEmpty(curState) AndAlso curState = "Existing" Then newState = "Proposed"
                                If Not String.IsNullOrEmpty(curState) AndAlso curState = "Proposed" Then newState = "Remove"
                                If Not String.IsNullOrEmpty(curState) AndAlso curState = "Remove" Then newState = "Existing"
                                newName = (baseName & sepChar & newState)
                            End If

                            Dim curSpace As BlockTableRecord = actrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)

                            Dim newBTRid As ObjectId

                            If BlkExists(newName) Then
                                newBTRid = blkTbl(newName)
                                Using newBref As New BlockReference(parentPos, newBTRid)
                                    With newBref
                                        .Position = parentPos
                                        .Rotation = parentOrient
                                    End With
                                    curSpace.AppendEntity(newBref)
                                    actrans.AddNewlyCreatedDBObject(newBref, True)
                                End Using
                                If Not parentBref.IsWriteEnabled Then parentBref.UpgradeOpen()
                                parentBref.Erase()
                            Else
                                Using newBTR As New BlockTableRecord
                                    If Not newBTR.IsWriteEnabled Then newBTR.UpgradeOpen()
                                    newBTR.Name = newName

                                    blkTbl.UpgradeOpen()
                                    newBTRid = blkTbl.Add(newBTR)
                                    actrans.AddNewlyCreatedDBObject(newBTR, True)

                                    For Each subID As ObjectId In ParentBTR
                                        Dim subObj As DBObject = actrans.GetObject(subID, OpenMode.ForRead)
                                        If TypeOf subObj Is BlockReference Then
                                            Dim SubRef As BlockReference = CType(subObj, BlockReference)
                                            Dim subrefName As String = SubRef.Name
                                            Dim subrefPos As Point3d = SubRef.Position
                                            Dim subRefScale As Scale3d = SubRef.ScaleFactors
                                            Dim newSubRef As New BlockReference(subrefPos, SubRef.DynamicBlockTableRecord)
                                            Dim newSubRefId As ObjectId = newBTR.AppendEntity(newSubRef)
                                            actrans.AddNewlyCreatedDBObject(newSubRef, True)
                                            Dim myResult As Boolean = ChangeVisState(actrans, newSubRefId, newState, "Visibility1")
                                            Debug.Print(myResult.ToString)
                                            newSubRef.ScaleFactors = subRefScale
                                        End If
                                    Next
                                End Using

                                Using newBref As New BlockReference(parentPos, newBTRid)
                                    With newBref
                                        .Position = parentPos
                                        .Rotation = parentOrient
                                    End With
                                    curSpace.AppendEntity(newBref)
                                    actrans.AddNewlyCreatedDBObject(newBref, True)
                                End Using
                                parentBref.UpgradeOpen()
                                parentBref.Erase()
                            End If
                        End If
                        actrans.Commit()
                    Next
                End Using
            End If
        End Sub

        <CommandMethod("CPMW")>
        Public Sub CreatePavementMarkingWords()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Dim pso As New PromptStringOptions(vbLf & "Input the word to be created (numbers and spaces Ok, no symbols allowed). ")
            With pso
                .AllowSpaces = True
            End With

            Dim psr As PromptResult = ed.GetString(pso)
            Dim myStr As String

            If psr.Status = PromptStatus.OK Then
                myStr = psr.StringResult
            Else
                Exit Sub
            End If

            Dim pko As New PromptKeywordOptions(vbLf & "Is this marking Existing, Proposed, or to be Removed?")
            With pko
                .Keywords.Add("Existing")
                .Keywords.Add("Proposed")
                .Keywords.Add("Remove")
                .AppendKeywordsToMessage = True
                .AllowNone = False
                .AllowArbitraryInput = False
            End With

            Dim pkr As PromptResult = ed.GetKeywords(pko)
            Dim vis As String

            If pkr.Status = PromptStatus.OK Then
                vis = pkr.StringResult
            Else
                ed.WriteMessage(vbLf & "Command ended.")
                Exit Sub
            End If

            Dim wrdBlkName As String = "P-" & myStr
            Dim repBlk As Boolean

            If BlkExists(wrdBlkName) Then
                Dim pko2 As New PromptKeywordOptions(vbLf & "block " & wrdBlkName & " already exists.  Do you want to replace existing block?")
                With pko2
                    .Keywords.Add("Y")
                    .Keywords.Add("N")
                    .AppendKeywordsToMessage = True
                    .AllowNone = False
                    .AllowArbitraryInput = False
                End With

                Dim pkr2 As PromptResult = ed.GetKeywords(pko2)
                If pkr2.Status = PromptStatus.OK Then

                    If pkr2.StringResult = "N" Then

                        repBlk = False
                        Dim newName As String

                        Dim pso2 As New PromptStringOptions(vbLf & "Input alternate name for block to be created.")
                        With pso2
                            .AllowSpaces = False
                        End With

                        Dim psr2 As PromptResult = ed.GetString(pso2)

                        If psr2.Status = PromptStatus.OK Then
                            newName = psr2.StringResult
                        Else

                            Exit Sub
                        End If

                        If BlkExists(newName) Then
                            ed.WriteMessage(vbLf & "Block with that name already exists. Ending Command.")
                            Exit Sub
                        Else
                            wrdBlkName = newName
                        End If

                    Else
                        repBlk = True
                    End If
                Else
                    ed.WriteMessage(vbLf & "Command ended.")
                    Exit Sub
                End If
            End If


            Dim letBlkPath As String = "\\EESServer\datadisk\CAD\BLOCKS\ROAD\PVMT\Design"

TryAgain:

            If Not My.Computer.FileSystem.DirectoryExists(letBlkPath) Then
                MessageBox.Show("Default folder does not exist.  Select the folder where the individual letter blocks are stored.")
                letBlkPath = GetMyFolderName()
                If String.IsNullOrEmpty(letBlkPath) OrElse Not My.Computer.FileSystem.DirectoryExists(letBlkPath) Then
                    ed.WriteMessage(vbLf & "Invalid Path.  Ending Command.")
                    Exit Sub
                Else
                    GoTo TryAgain
                End If
            End If

            letBlkPath &= "\"

            If Not My.Computer.FileSystem.FileExists(letBlkPath & "P-A.dwg") Then
                ed.WriteMessage(vbLf & "Invalid Path.  Letter files not found.  Ending Command.")
                Exit Sub
            End If

            Dim letters(myStr.Length - 1) As String
            For j As Integer = 1 To myStr.Length
                letters(j - 1) = Mid(myStr, j, 1)
            Next

            Try

                Dim cPos As Double = 0.0

                Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                    Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    Dim curSpace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForRead)
                    Dim newWrdBlk As BlockTableRecord

                    If repBlk Then
                        Dim wbID As ObjectId = ClearBlk(wrdBlkName, acTrans)
                        newWrdBlk = acTrans.GetObject(wbID, OpenMode.ForWrite)
                    Else
                        newWrdBlk = New BlockTableRecord With {.Name = wrdBlkName}
                    End If

                    Dim midDist As Double
                    Dim ltrID As ObjectId

                    For i = 0 To myStr.Length - 1

                        If Not letters(i) = " " Then
                            Dim letBlkName As String = "P-" & letters(i)
                            Dim letBlkFullName As String = letBlkPath & letBlkName & ".dwg"

                            If Not BlkExists(letBlkName) Then
                                If My.Computer.FileSystem.FileExists(letBlkFullName) Then
                                    ltrID = InsertDwg(letBlkFullName, letBlkName, acTrans)
                                Else
                                    ed.WriteMessage("Letter block file " & letBlkName & " does not exist at the provided path.")
                                    Exit Sub
                                End If
                            Else
                                ltrID = blkTbl(letBlkName)
                            End If

                            Dim inspt As New Point3d(cPos, 0, 0)

                            'Dim letBlkDef As BlockTableRecord = acTrans.GetObject(ltrID, OpenMode.ForRead)
                            Dim letbRef As New BlockReference(inspt, ltrID)

                            If repBlk Then
                                newWrdBlk.AppendEntity(letbRef)
                                acTrans.AddNewlyCreatedDBObject(letbRef, True)
                            Else
                                newWrdBlk.AppendEntity(letbRef)
                            End If


                        End If

                        If IsNumeric(letters(i)) Then
                            Select Case letters(i)
                                Case Is = "1"
                                    cPos += 0.33333 * 7
                                Case Is = "2"
                                    cPos += 0.33333 * 9
                                Case Is = "4"
                                    cPos += 0.33333 * 9
                                Case Else
                                    cPos += 0.33333 * 8
                            End Select

                            midDist = (cPos - 0.666667) / 2

                        Else
                            Select Case letters(i).ToUpper
                                Case Is = "I"
                                    cPos += 0.33333 * 2
                                Case Is = " "
                                    cPos += 0.33333 * 4
                                Case Else
                                    cPos += 0.33333 * 5
                            End Select

                            midDist = (cPos - 0.33333) / 2

                        End If
                    Next

                    blkTbl.UpgradeOpen()
                    Dim wrdBlkId As ObjectId

                    If repBlk Then
                        wrdBlkId = blkTbl(wrdBlkName)
                    Else
                        wrdBlkId = blkTbl.Add(newWrdBlk)
                    End If

                    newWrdBlk = acTrans.GetObject(wrdBlkId, OpenMode.ForWrite)

                    For Each id As ObjectId In newWrdBlk
                        Dim obj As DBObject = acTrans.GetObject(id, OpenMode.ForRead)
                        If TypeOf obj Is BlockReference Then
                            Dim ltrRef As BlockReference = CType(obj, BlockReference)
                            ltrRef.UpgradeOpen()

                            Dim orgPt As New Point3d(0, 0, 0)
                            Dim cPt As New Point3d(midDist, 4, 0)

                            Dim dispVect As Vector3d = cPt.GetVectorTo(orgPt)
                            ltrRef.TransformBy(Matrix3d.Displacement(dispVect))

                            Dim bId As ObjectId = ltrRef.ObjectId

                            If vis = "Existing" Then
                                ChangeVisState(bId, "Existing", "Visibility1")
                            ElseIf vis = "Proposed" Then
                                ChangeVisState(bId, "Proposed", "Visibility1")
                            ElseIf vis = "Remove" Then
                                ChangeVisState(bId, "Remove", "Visibility1")
                            End If
                        End If
                    Next

                    If repBlk Then
                        Dim bRefIds As ObjectIdCollection = newWrdBlk.GetBlockReferenceIds(False, True)
                        For Each brefid As ObjectId In bRefIds
                            Dim br As BlockReference = acTrans.GetObject(brefid, OpenMode.ForWrite, False, True)
                            br.RecordGraphicsModified(True)
                        Next
                    End If

                    Dim ppo As New PromptPointOptions(vbLf & "Pick point for location of pavement marking.")

                    Dim ppr As PromptPointResult = ed.GetPoint(ppo)
                    Dim iP As Point3d

                    If ppr.Status = PromptStatus.OK Then
                        iP = ppr.Value
                        Dim wrdBref As New BlockReference(iP, wrdBlkId)

                        curSpace.UpgradeOpen()
                        curSpace.AppendEntity(wrdBref)
                        acTrans.AddNewlyCreatedDBObject(wrdBref, True)
                    End If

                    acTrans.Commit()

                End Using

            Catch ex As Exception
                Dim msgStr As String = "Fatal Error:" & vbLf & ex.Message & vbLf & ex.Source.ToString & vbLf
                MessageBox.Show(msgStr)
            End Try

        End Sub

        <CommandMethod("SEALSIG")>
        Public Sub SealSig()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            'start layout manager and transaction
            Dim lm As LayoutManager = LayoutManager.Current
            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction

                'ask what to do
                Dim pkO As New PromptKeywordOptions(vbLf & "Show, Hide, or Toggle the Engineer's signature?")
                With pkO
                    .Keywords.Add("Show")
                    .Keywords.Add("Hide")
                    .Keywords.Add("Toggle")
                    .AllowNone = False
                    .AppendKeywordsToMessage = True
                End With

                Dim pkResult As String
                Dim pkr As PromptResult = ed.GetKeywords(pkO)

                If pkr.Status = PromptStatus.OK Then
                    pkResult = pkr.StringResult
                Else
                    pkResult = "Toggle"
                End If

                'pick the layouts to be adjusted
                Dim lPicker As New LayoutPicker
                Dim loList As SortedDictionary(Of Integer, String) = LayoutTabList()

                'add layout names to the picker form
                For Each lN As String In loList.Values
                    Dim lName As String = lN
                    If Not lName = "Model" Then
                        lPicker.ListBox1.Items.Add(lName)
                    End If
                Next

                lPicker.PickerLabel.Text = "Select layouts to update engineer's seal:"

                'create a list variable and show the form
                Dim layoutLst As List(Of String)
                lPicker.ShowDialog()

                If lPicker.DialogResult = DialogResult.Cancel Then
                    lPicker.Dispose()
                    Exit Sub
                Else
                    layoutLst = lPicker.PickedList
                    lPicker.Dispose()
                End If

                Dim layDict As DBDictionary = dwgDB.LayoutDictionaryId.GetObject(OpenMode.ForRead)
                'Dim layID As ObjectId = layDict(layoutLst(0))
                'Dim lo As Layout = acTrans.GetObject(layID, OpenMode.ForWrite)

                'run through the layout list
                If layoutLst.Count > 0 Then
                    For Each loName As String In layoutLst

                        'get the layout ID
                        Dim loID As ObjectId = layDict(loName)
                        Dim lo As Layout = acTrans.GetObject(loID, OpenMode.ForRead)

                        'get the layout's paperspace block table record 
                        Dim curSpace As BlockTableRecord = acTrans.GetObject(lo.BlockTableRecordId, OpenMode.ForRead)

                        'cycle through objects in the layout's paperspace
                        For Each id As ObjectId In curSpace
                            Dim dbObj As DBObject = acTrans.GetObject(id, OpenMode.ForRead)
                            'if object is a blockreference, then check if it is dynamic
                            If TypeOf dbObj Is BlockReference Then
                                Dim bref As BlockReference = CType(dbObj, BlockReference)
                                If bref.IsDynamicBlock Then
                                    'get dynamic BTR
                                    Dim bid As ObjectId = bref.DynamicBlockTableRecord
                                    Dim dyBref As BlockTableRecord = TryCast(acTrans.GetObject(bid, OpenMode.ForRead), BlockTableRecord)
                                    If dyBref IsNot Nothing Then
                                        'if the name has "seal" in it, it is probably the right block
                                        Dim nm As String = dyBref.Name
                                        If InStr(nm.ToUpper, "SEAL") <> 0 Then
                                            dyBref.UpgradeOpen()
                                            'cycle dynamic properties to find "Visibility1"
                                            For Each prop As DynamicBlockReferenceProperty In bref.DynamicBlockReferencePropertyCollection
                                                If prop.PropertyName = "Visibility1" Then
                                                    If pkResult = "Toggle" Then
                                                        If prop.Value = "Signature" Then
                                                            ChangeVisState(id, "No Signature", "Visibility1")
                                                        Else
                                                            ChangeVisState(id, "Signature", "Visibility1")
                                                        End If
                                                    ElseIf pkResult = "Show" Then
                                                        ChangeVisState(id, "Signature", "Visibility1")
                                                    Else
                                                        ChangeVisState(id, "No Signature", "Visibility1")
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Next
                End If
                acTrans.Commit()
            End Using
        End Sub

    End Module

    Public Module GeometryCommands

        Private m_area As Double

        <CommandMethod("CBYA")>
        Public Sub CircleByArea()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Try
                Dim pdo1 As New PromptDoubleOptions(vbLf & "Enter the area of the circle.")
                With pdo1
                    .AllowNegative = False
                    .AllowNone = False
                    .AllowArbitraryInput = True
                    .AllowZero = False
                    If m_area > 0 Then .DefaultValue = m_area
                End With

                Dim pdr1 As PromptDoubleResult = ed.GetDouble(pdo1)

                Dim ar As Double

                If pdr1.Status = PromptStatus.OK Then
                    ar = pdr1.Value
                    m_area = ar
                Else
                    Exit Sub
                End If

                Dim ppo As New PromptPointOptions(vbLf & "Select the center of the circle.")
                With ppo
                    .AllowNone = False
                End With

                Dim ppr As PromptPointResult = ed.GetPoint(ppo)
                Dim ptP As Point3d

                If ppr.Status = PromptStatus.OK Then
                    ptP = ppr.Value
                Else
                    Exit Sub
                End If

                Dim r As Double = (ar / Math.PI) ^ 0.5

                If Not DwgDB.TileMode Then
                    ed.WriteMessage(vbLf & "Must be in Model space to use this command.")
                    Exit Sub
                End If

                Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction
                    Dim blktbl As BlockTable = actrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    Dim mdlSpace As BlockTableRecord = actrans.GetObject(blktbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    Dim circ As New Circle(ptP, Vector3d.ZAxis, r)

                    mdlSpace.AppendEntity(circ)
                    actrans.AddNewlyCreatedDBObject(circ, True)

                    actrans.Commit()
                End Using

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        <CommandMethod("SBYA")>
        Public Sub SquareByArea()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Try
                Dim pdo1 As New PromptDoubleOptions(vbLf & "Enter the area of the square.")
                With pdo1
                    .AllowNegative = False
                    .AllowNone = False
                    .AllowArbitraryInput = True
                    .AllowZero = False
                    If m_area > 0 Then .DefaultValue = m_area
                End With

                Dim pdr1 As PromptDoubleResult = ed.GetDouble(pdo1)

                Dim ar As Double

                If pdr1.Status = PromptStatus.OK Then
                    ar = pdr1.Value
                    m_area = ar
                Else
                    Exit Sub
                End If

                Dim ppo As New PromptPointOptions(vbLf & "Select the center of the square.")
                With ppo
                    .AllowNone = False
                End With

                Dim ppr As PromptPointResult = ed.GetPoint(ppo)
                Dim ptP As Point3d

                If ppr.Status = PromptStatus.OK Then
                    ptP = ppr.Value
                Else
                    Exit Sub
                End If

                Dim s As Double = ar ^ 0.5

                If Not DwgDB.TileMode Then
                    ed.WriteMessage(vbLf & "Must be in Model space to use this command.")
                    Exit Sub
                End If

                Dim p1 As New Point2d(ptP.X - (s / 2), ptP.Y - (s / 2))
                Dim p2 As New Point2d(ptP.X - (s / 2), ptP.Y + (s / 2))
                Dim p3 As New Point2d(ptP.X + (s / 2), ptP.Y + (s / 2))
                Dim p4 As New Point2d(ptP.X + (s / 2), ptP.Y - (s / 2))

                Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction
                    Dim blktbl As BlockTable = actrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    Dim mdlSpace As BlockTableRecord = actrans.GetObject(blktbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    Dim pl As New Polyline
                    With pl
                        .AddVertexAt(0, p1, 0, 0, 0)
                        .AddVertexAt(0, p2, 0, 0, 0)
                        .AddVertexAt(0, p3, 0, 0, 0)
                        .AddVertexAt(0, p4, 0, 0, 0)
                        .Closed = True
                    End With

                    mdlSpace.AppendEntity(pl)
                    actrans.AddNewlyCreatedDBObject(pl, True)

                    actrans.Commit()
                End Using

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try


        End Sub


        <CommandMethod("ETAN")>
        Public Sub ExteriorTangent2Circles()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Try
                Dim peo1 As New PromptEntityOptions(vbLf & "Select first circle or arc")

                With peo1
                    .SetRejectMessage(vbLf & "Must pick a circle or arc.")
                    .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Circle), True)
                    .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Arc), True)
                    .AllowNone = False
                End With

                Dim dbObId As ObjectId

                Dim pr1 As PromptEntityResult = ed.GetEntity(peo1)
                If pr1.Status = PromptStatus.OK Then
                    dbObId = pr1.ObjectId
                Else
                    Exit Sub
                End If

                Dim peo2 As New PromptEntityOptions(vbLf & "Select second circle or arc")

                With peo2
                    .SetRejectMessage(vbLf & "Must pick a circle or arc.")
                    .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Circle), True)
                    .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Arc), True)
                    .AllowNone = False
                End With

                Dim dbObId2 As ObjectId

                Dim pr2 As PromptEntityResult = ed.GetEntity(peo2)
                If pr2.Status = PromptStatus.OK Then
                    dbObId2 = pr2.ObjectId
                Else
                    Exit Sub
                End If

                If dbObId = dbObId2 Then
                    ed.WriteMessage(vbLf & "Error.  Different circles must be selected.")
                    Exit Sub
                End If

                Dim pko As New PromptKeywordOptions(vbLf & "Draw nodes at tangent points?")
                With pko
                    .Keywords.Add("Yes")
                    .Keywords.Add("No")
                    .AppendKeywordsToMessage = True
                    .AllowArbitraryInput = False
                    .AllowNone = False
                End With

                Dim pkr As PromptResult = ed.GetKeywords(pko)
                Dim showPts As Boolean

                If pkr.Status = PromptStatus.OK Then
                    If pkr.StringResult = "Yes" Then
                        showPts = True
                    Else
                        showPts = False
                    End If
                Else
                    Exit Sub
                End If

                Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction

                    Dim dbObj1 As DBObject = acTrans.GetObject(dbObId, OpenMode.ForRead)
                    Dim dbObj2 As DBObject = acTrans.GetObject(dbObId2, OpenMode.ForRead)
                    Dim pCol As Point2dCollection = ExtTan2Circles(dbObj1, dbObj2, True)

                    If pCol.Count < 3 Then Exit Sub

                    Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    If pCol.Count >= 4 Then
                        Dim pline1 As New Polyline
                        With pline1
                            .AddVertexAt(0, pCol(0), 0, 0, 0)
                            .AddVertexAt(1, pCol(1), 0, 0, 0)
                        End With
                        mdlSpace.AppendEntity(pline1)
                        acTrans.AddNewlyCreatedDBObject(pline1, True)

                        Dim pline2 As New Polyline
                        With pline2
                            .AddVertexAt(0, pCol(2), 0, 0, 0)
                            .AddVertexAt(1, pCol(3), 0, 0, 0)
                        End With
                        mdlSpace.AppendEntity(pline2)
                        acTrans.AddNewlyCreatedDBObject(pline2, True)
                    End If

                    If pCol.Count = 3 Then
                        Dim pline3 As New Polyline
                        With pline3
                            .AddVertexAt(0, pCol(0), 0, 0, 0)
                            .AddVertexAt(1, pCol(1), 0, 0, 0)
                        End With
                        mdlSpace.AppendEntity(pline3)
                        acTrans.AddNewlyCreatedDBObject(pline3, True)

                    ElseIf pCol.Count = 7 Then
                        Dim pline3 As New Polyline
                        With pline3
                            .AddVertexAt(0, pCol(4), 0, 0, 0)
                            .AddVertexAt(1, pCol(5), 0, 0, 0)
                        End With
                        mdlSpace.AppendEntity(pline3)
                        acTrans.AddNewlyCreatedDBObject(pline3, True)
                    End If

                    If showPts Then

                        If pCol.Count >= 4 Then

                            Dim p1 As Point2d = pCol(0)
                            Dim p2 As Point2d = pCol(1)
                            Dim p3 As Point2d = pCol(2)
                            Dim p4 As Point2d = pCol(3)

                            'create nodes for the tangent points
                            Dim dbp1 As New DBPoint(New Point3d(p1.X, p1.Y, 0))
                            Dim dbp2 As New DBPoint(New Point3d(p2.X, p2.Y, 0))
                            Dim dbp3 As New DBPoint(New Point3d(p3.X, p3.Y, 0))
                            Dim dbp4 As New DBPoint(New Point3d(p4.X, p4.Y, 0))

                            mdlSpace.AppendEntity(dbp1)
                            acTrans.AddNewlyCreatedDBObject(dbp1, True)

                            mdlSpace.AppendEntity(dbp2)
                            acTrans.AddNewlyCreatedDBObject(dbp2, True)

                            mdlSpace.AppendEntity(dbp3)
                            acTrans.AddNewlyCreatedDBObject(dbp3, True)

                            mdlSpace.AppendEntity(dbp4)
                            acTrans.AddNewlyCreatedDBObject(dbp4, True)
                        End If

                        If pCol.Count = 3 Then
                            Dim p5 As Point2d = pCol(2)
                            Dim dbp5 As New DBPoint(New Point3d(p5.X, p5.Y, 0))
                            mdlSpace.AppendEntity(dbp5)
                            acTrans.AddNewlyCreatedDBObject(dbp5, True)
                        End If

                        If pCol.Count = 7 Then
                            Dim p5 As Point2d = pCol(6)
                            Dim dbp5 As New DBPoint(New Point3d(p5.X, p5.Y, 0))
                            mdlSpace.AppendEntity(dbp5)
                            acTrans.AddNewlyCreatedDBObject(dbp5, True)
                        End If



                    End If
                    acTrans.Commit()

                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        <CommandMethod("ITAN")>
        Public Sub InteriorTangent2Circles()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Try
                Dim peo1 As New PromptEntityOptions(vbLf & "Select first circle or arc")

                With peo1
                    .SetRejectMessage(vbLf & "Must pick a circle or arc.")
                    .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Circle), True)
                    .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Arc), True)
                    .AllowNone = False
                End With

                Dim dbObId As ObjectId

                Dim pr1 As PromptEntityResult = ed.GetEntity(peo1)
                If pr1.Status = PromptStatus.OK Then
                    dbObId = pr1.ObjectId
                Else
                    Exit Sub
                End If

                Dim peo2 As New PromptEntityOptions(vbLf & "Select second circle or arc")

                With peo2
                    .SetRejectMessage(vbLf & "Must pick a circle or arc.")
                    .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Circle), True)
                    .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Arc), True)
                    .AllowNone = False
                End With

                Dim dbObId2 As ObjectId

                Dim pr2 As PromptEntityResult = ed.GetEntity(peo2)
                If pr2.Status = PromptStatus.OK Then
                    dbObId2 = pr2.ObjectId
                Else
                    Exit Sub
                End If
                Dim pko As New PromptKeywordOptions(vbLf & "Draw nodes at tangent points?")
                With pko
                    .Keywords.Add("Yes")
                    .Keywords.Add("No")
                    .AppendKeywordsToMessage = True
                    .AllowArbitraryInput = False
                    .AllowNone = False
                End With

                Dim pkr As PromptResult = ed.GetKeywords(pko)
                Dim showPts As Boolean

                If pkr.Status = PromptStatus.OK Then
                    If pkr.StringResult = "Yes" Then
                        showPts = True
                    Else
                        showPts = False
                    End If
                Else
                    Exit Sub
                End If

                Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                    Dim dbObj1 As DBObject = acTrans.GetObject(dbObId, OpenMode.ForRead)
                    Dim dbObj2 As DBObject = acTrans.GetObject(dbObId2, OpenMode.ForRead)
                    Dim pCol As Point2dCollection = IntTan2Circles(dbObj1, dbObj2, True)

                    If pCol Is Nothing OrElse pCol.Count < 4 Then
                        Exit Sub
                    Else
                        Dim p1 As Point2d = pCol(0)
                        Dim p2 As Point2d = pCol(1)
                        Dim p3 As Point2d = pCol(2)
                        Dim p4 As Point2d = pCol(3)

                        Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                        Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        Dim pline1 As New Polyline
                        With pline1
                            .AddVertexAt(0, p1, 0, 0, 0)
                            .AddVertexAt(1, p4, 0, 0, 0)
                        End With
                        mdlSpace.AppendEntity(pline1)
                        acTrans.AddNewlyCreatedDBObject(pline1, True)

                        Dim pline2 As New Polyline
                        With pline2
                            .AddVertexAt(0, p2, 0, 0, 0)
                            .AddVertexAt(1, p3, 0, 0, 0)
                        End With
                        mdlSpace.AppendEntity(pline2)
                        acTrans.AddNewlyCreatedDBObject(pline2, True)

                        If showPts Then

                            If pCol.Count = 4 Then
                                'create nodes for the tangent points
                                Dim dbp1 As New DBPoint(New Point3d(p1.X, p1.Y, 0))
                                Dim dbp2 As New DBPoint(New Point3d(p2.X, p2.Y, 0))
                                Dim dbp3 As New DBPoint(New Point3d(p3.X, p3.Y, 0))
                                Dim dbp4 As New DBPoint(New Point3d(p4.X, p4.Y, 0))

                                mdlSpace.AppendEntity(dbp1)
                                acTrans.AddNewlyCreatedDBObject(dbp1, True)

                                mdlSpace.AppendEntity(dbp2)
                                acTrans.AddNewlyCreatedDBObject(dbp2, True)

                                mdlSpace.AppendEntity(dbp3)
                                acTrans.AddNewlyCreatedDBObject(dbp3, True)

                                mdlSpace.AppendEntity(dbp4)
                                acTrans.AddNewlyCreatedDBObject(dbp4, True)
                            End If
                        End If
                    End If

                    acTrans.Commit()

                End Using

            Catch ex As Exception
                ed.WriteMessage(vbLf & "Fatal error in sub InteriorTangent2Circles.")
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        <CommandMethod("TANPT")>
        Public Sub TangentsFromPoint()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Dim peo1 As New PromptEntityOptions(vbLf & "Select circle or arc")

            With peo1
                .SetRejectMessage(vbLf & "Must pick a circle or arc.")
                .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Circle), True)
                .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Arc), True)
                .AllowNone = False
            End With

            Dim cObId As ObjectId

            Dim pr1 As PromptEntityResult = ed.GetEntity(peo1)

            If pr1.Status = PromptStatus.OK Then
                cObId = pr1.ObjectId
            Else
                Exit Sub
            End If

            Dim ppo As New PromptPointOptions(vbLf & "Select a point not inside the circle or arc.")
            With ppo
                .AllowNone = False
            End With

            Dim ppr As PromptPointResult = ed.GetPoint(ppo)
            Dim ptP As Point3d

            If ppr.Status = PromptStatus.OK Then
                ptP = ppr.Value
            Else
                Exit Sub
            End If

            Dim pko As New PromptKeywordOptions(vbLf & "Draw nodes at tangent points?")
            With pko
                .Keywords.Add("Yes")
                .Keywords.Add("No")
                .AppendKeywordsToMessage = True
                .AllowArbitraryInput = False
                .AllowNone = False
            End With

            Dim pkr As PromptResult = ed.GetKeywords(pko)
            Dim showPts As Boolean

            If pkr.Status = PromptStatus.OK Then
                If pkr.StringResult = "Yes" Then
                    showPts = True
                Else
                    showPts = False
                End If
            Else
                Exit Sub
            End If

            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction

                Dim dbObj As DBObject = acTrans.GetObject(cObId, OpenMode.ForRead)
                Dim ptCol As Point2dCollection = GetTangentPoints(ptP, dbObj)

                If ptCol IsNot Nothing AndAlso ptCol.Count > 0 Then
                    Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    If ptCol.Count = 1 Then
                        ed.WriteMessage(vbLf & "Picked point is the tangent point.")
                        If showPts Then
                            Dim dbp0 As New DBPoint(ptP)
                            mdlSpace.AppendEntity(dbp0)
                            acTrans.AddNewlyCreatedDBObject(dbp0, True)
                        End If

                    ElseIf ptCol.Count = 2 Then
                        Dim ptp2D As New Point2d(ptP.X, ptP.Y)
                        Dim ptp3D As New Point3d(ptP.X, ptP.Y, 0)
                        Dim pt13D As New Point3d(ptCol(0).X, ptCol(0).Y, 0)
                        Dim pt23D As New Point3d(ptCol(1).X, ptCol(1).Y, 0)

                        Using pline1 As New Polyline
                            With pline1
                                .AddVertexAt(0, ptCol(0), 0, 0, 0)
                                .AddVertexAt(1, ptp2D, 0, 0, 0)
                                .AddVertexAt(2, ptCol(1), 0, 0, 0)
                            End With
                            mdlSpace.AppendEntity(pline1)
                            acTrans.AddNewlyCreatedDBObject(pline1, True)
                        End Using

                        If showPts Then
                            Dim dbp0 As New DBPoint(ptp3D)
                            mdlSpace.AppendEntity(dbp0)
                            acTrans.AddNewlyCreatedDBObject(dbp0, True)
                            If ptCol.Count > 1 Then
                                Dim dbp1 As New DBPoint(pt13D)
                                Dim dbp2 As New DBPoint(pt23D)
                                mdlSpace.AppendEntity(dbp1)
                                acTrans.AddNewlyCreatedDBObject(dbp1, True)
                                mdlSpace.AppendEntity(dbp2)
                                acTrans.AddNewlyCreatedDBObject(dbp2, True)
                            End If
                        End If
                    End If
                End If
                acTrans.Commit()
            End Using
        End Sub


        <CommandMethod("testVects")>
        Public Sub TESTVECTS()

            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            'pick first point
            Dim pt1Pick As New PromptPointOptions("")
            With pt1Pick
                .AllowNone = False
                .Message = vbLf & "Pick first 3d point."
            End With

            'declare first point
            Dim pt1Res As PromptPointResult = ed.GetPoint(pt1Pick)
            Dim wPt1 As Point3d

            'if point is okay, then store it, if not, restore the UCS and crash out
            If pt1Res.Status = PromptStatus.OK Then
                wPt1 = pt1Res.Value
            ElseIf pt1Res.Status = PromptStatus.Cancel Then
                Exit Sub
            End If

            'declare second point
            Dim wPt2 As Point3d
            Dim pt2Pick As New PromptPointOptions("")

NextPoint:
            'pick second point
            With pt2Pick
                .AllowNone = True
                .Message = vbLf & "Pick Next 3D point."
            End With

            Dim pt2Res As PromptPointResult = ed.GetPoint(pt2Pick)

            'if it is okay, store it.  If user right clicks or hits escape, end the program
            If pt2Res.Status = PromptStatus.OK Then
                wPt2 = pt2Res.Value
            ElseIf pt2Res.Status = PromptStatus.Cancel Or pt2Res.Status = PromptStatus.None Then
                CurDwg.Dispose()
                DwgDB.Dispose()
                Exit Sub
            End If

            'if both points are identical, crash out
            If wPt2 = wPt1 Then
                MsgBox("Picked points are identical.  Ending program.")
                Exit Sub
            End If

            Dim testVect As Vector3d = wPt1.GetVectorTo(wPt2)
            Dim newvect As Vector3d = UnitVector3d(testVect)

            If Not newvect.IsUnitLength Then
                MsgBox("not a unit vector")
            Else
                MsgBox("good unit vector")
            End If

            'MakeNewUCS(wPt1, wPt2, "testUCS")
            'Dim newmatrix As Matrix3d = GetUCSMatrix3D(wPt1, wPt2)

        End Sub


        <CommandMethod("RgnCentroid")>
        Public Sub RegionCentroid()

            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Dim pEO As New PromptEntityOptions(vbLf & "Select hatch or region.")
            With pEO
                .SetRejectMessage(vbLf & "Not a region.  Try again.")
                .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Region), True)
                .AllowNone = False
            End With

            Dim peR As PromptEntityResult = ed.GetEntity(pEO)
            Dim tRegID As ObjectId
            If peR.Status = PromptStatus.OK Then
                tRegID = peR.ObjectId
            Else
                Exit Sub
            End If

            Dim curUCSMatrix As Matrix3d = ed.CurrentUserCoordinateSystem
            Dim curUCS As CoordinateSystem3d = curUCSMatrix.CoordinateSystem3d

            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim blktbl As BlockTable = TryCast(acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim curSpace As BlockTableRecord = TryCast(acTrans.GetObject(blktbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

                Dim tReg As Autodesk.AutoCAD.DatabaseServices.Region = acTrans.GetObject(tRegID, OpenMode.ForRead)

                Dim tCentroid As Point2d = tReg.AreaProperties(curUCS.Origin, curUCS.Xaxis, curUCS.Yaxis).Centroid
                Dim tC As New Point3d(tCentroid.X, tCentroid.Y, 0)

                Dim circ As New Circle(tC, curUCS.Zaxis, 0.25)

                curSpace.AppendEntity(circ)
                acTrans.AddNewlyCreatedDBObject(circ, True)

                acTrans.Commit()

            End Using
        End Sub

        <CommandMethod("LV")>
        Public Sub ListVertices()
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Dim per As PromptEntityResult = ed.GetEntity("Select a polyline")

            If per.Status = PromptStatus.OK Then
                Dim tr As Transaction = db.TransactionManager.StartTransaction()

                Using tr
                    Dim obj As DBObject = tr.GetObject(per.ObjectId, OpenMode.ForRead)
                    Dim lwp As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(obj, Autodesk.AutoCAD.DatabaseServices.Polyline)

                    If lwp IsNot Nothing Then
                        Dim vn As Integer = lwp.NumberOfVertices

                        For i As Integer = 0 To vn - 1
                            Dim pt As Point2d = lwp.GetPoint2dAt(i)
                            ed.WriteMessage(vbLf & pt.ToString())
                        Next
                    Else
                        Dim p2d As Polyline2d = TryCast(obj, Polyline2d)

                        If p2d IsNot Nothing Then
                            For Each vId As ObjectId In p2d
                                Dim v2d As Vertex2d = CType(tr.GetObject(vId, OpenMode.ForRead), Vertex2d)
                                ed.WriteMessage(vbLf & v2d.Position.ToString())
                            Next
                        Else
                            Dim p3d As Polyline3d = TryCast(obj, Polyline3d)
                            If p3d IsNot Nothing Then
                                For Each vId As ObjectId In p3d
                                    Dim v3d As PolylineVertex3d = CType(tr.GetObject(vId, OpenMode.ForRead), PolylineVertex3d)
                                    ed.WriteMessage(vbLf & v3d.Position.ToString())
                                Next
                            End If
                        End If
                    End If
                    tr.Commit()
                End Using
            End If
        End Sub

        Public Function GetVertexCoords(dbObjId As ObjectId) As Object
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Dim retVal As Object

            Using tr
                Dim obj As DBObject = tr.GetObject(dbObjId, OpenMode.ForRead)

                If TypeOf obj Is Polyline Then
                    Dim lwp As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(obj, Autodesk.AutoCAD.DatabaseServices.Polyline)
                    Dim vn As Integer = lwp.NumberOfVertices
                    Dim retCol As New Point2dCollection
                    For i As Integer = 0 To vn - 1
                        Dim pt As Point2d = lwp.GetPoint2dAt(i)
                        ed.WriteMessage(vbLf & pt.ToString())
                        retCol.Add(pt)
                    Next
                    retVal = TryCast(retCol, Object)

                ElseIf TypeOf obj Is Polyline2d Then
                    Dim p2d As Polyline2d = TryCast(obj, Polyline2d)
                    Dim retCol As New Point3dCollection
                    For Each vId As ObjectId In p2d
                        Dim v2d As Vertex2d = CType(tr.GetObject(vId, OpenMode.ForRead), Vertex2d)
                        Dim pt As Point3d = v2d.Position
                        ed.WriteMessage(vbLf & pt.ToString())
                        retCol.Add(pt)
                    Next
                    retVal = TryCast(retCol, Object)

                ElseIf TypeOf obj Is Polyline3d Then
                    Dim p3d As Polyline3d = TryCast(obj, Polyline3d)
                    Dim retCol As New Point3dCollection
                    For Each vId As ObjectId In p3d
                        Dim v3d As PolylineVertex3d = CType(tr.GetObject(vId, OpenMode.ForRead), PolylineVertex3d)
                        Dim pt As Point3d = v3d.Position
                        ed.WriteMessage(vbLf & pt.ToString())
                        retCol.Add(pt)
                    Next
                    retVal = TryCast(retCol, Object)
                Else
                    retVal = Nothing
                End If

                Return retVal
                tr.Commit()
            End Using

        End Function



    End Module

    Public Module LayerCommands

        Private ReadOnly m_layIDList As ObjectIdCollection
        Private m_layList As List(Of String)

        <CommandMethod("FRZVPLAY", CommandFlags.UsePickSet Or CommandFlags.Redraw Or CommandFlags.Modal)>
        Public Sub FreezeVPLayers()
            'by David Eisenbeisz (c)2024
            'freezes viewport layers by picking entities

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Dim SelResult As PromptSelectionResult = ed.SelectImplied()

            'If Not SelResult.Status = PromptStatus.OK Then Exit Sub

            If SelResult.Status = PromptStatus.Error Then
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select objects to freeze their layers in current viewport")}
                'Seloptions.MessageForAdding = String.Format(vbLf & "Select centerline(s)")
                SelResult = ed.GetSelection(Seloptions)
            Else
                ed.SetImpliedSelection(New ObjectId(-1) {})
            End If

            If Not SelResult.Status = PromptStatus.OK Then Exit Sub

            If Not IsInLayoutViewport() Then
                ed.WriteMessage(vbLf & "Command must be run from a paperspace viewport.  End Command.")
                Exit Sub
            End If

            Dim selSet As SelectionSet = SelResult.Value
            Dim obIds() As ObjectId = selSet.GetObjectIds

            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim fzLayers As New ObjectIdCollection
                For i = 0 To obIds.Count - 1
                    Dim ent As Entity = TryCast(actrans.GetObject(obIds(i), OpenMode.ForRead), Entity)
                    If ent IsNot Nothing Then
                        Dim entLayID As ObjectId = ent.LayerId
                        fzLayers.Add(entLayID)
                        ent.Dispose()
                    End If
                Next

                'Dim cVPID As ObjectId = DwgDB.CurrentViewportTableRecordId
                Dim cVpid As ObjectId = ed.CurrentViewportObjectId
                'Dim lo As BlockTableRecord = actrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForRead)
                Dim cVP As Viewport = actrans.GetObject(cVpid, OpenMode.ForWrite)
                cVP.FreezeLayersInViewport(fzLayers.GetEnumerator)
                actrans.Commit()

            End Using

        End Sub

        <CommandMethod("LAOFF", CommandFlags.UsePickSet Or CommandFlags.Redraw Or CommandFlags.Modal)>
        Public Sub TurnOffLayers()
            'by David Eisenbeisz (c)2024
            'temporarily turns off layers and stores them in a list to turn back on

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Dim SelResult As PromptSelectionResult = ed.SelectImplied()

            'If Not SelResult.Status = PromptStatus.OK Then Exit Sub

            If SelResult.Status = PromptStatus.Error Then
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select objects to turn off their layers.")}
                'Seloptions.MessageForAdding = String.Format(vbLf & "Select centerline(s)")
                SelResult = ed.GetSelection(Seloptions)
            Else
                ed.SetImpliedSelection(New ObjectId(-1) {})
            End If

            If Not SelResult.Status = PromptStatus.OK Then Exit Sub

            Dim selSet As SelectionSet = SelResult.Value
            Dim obIds() As ObjectId = selSet.GetObjectIds

            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim offLayers As New List(Of String)
                For i = 0 To obIds.Count - 1
                    Dim ent As Entity = TryCast(actrans.GetObject(obIds(i), OpenMode.ForRead), Entity)
                    If ent IsNot Nothing Then
                        Dim entLay As String = ent.Layer
                        offLayers.Add(entLay)
                    End If
                Next

                Dim tempList As New List(Of String)

                If m_layList IsNot Nothing AndAlso m_layList.Count > 0 Then
                    If Not YesNoQuery(vbLf & "Do you want to clear the current list of turned off layers?  If Yes, only currently selected layers will be queued for turn on command.") Then
                        tempList = m_layList
                    End If
                End If

                Dim lrTbl As LayerTable = actrans.GetObject(DwgDB.LayerTableId, OpenMode.ForRead)

                For Each lyrName As String In offLayers
                    If lrTbl.Has(lyrName) Then
                        Dim ltr As LayerTableRecord = TryCast(actrans.GetObject(lrTbl(lyrName), OpenMode.ForRead), LayerTableRecord)
                        If ltr IsNot Nothing Then
                            Debug.Print(lyrName)
                            tempList.Add(lyrName)
                        End If
                    End If
                Next

                m_layList = tempList.Distinct.ToList

                For Each s As String In m_layList
                    Dim ltr As LayerTableRecord = TryCast(actrans.GetObject(lrTbl(s), OpenMode.ForWrite), LayerTableRecord)
                    ltr.IsOff = True
                Next
                actrans.Commit()
            End Using

        End Sub

        <CommandMethod("LAON")>
        Public Sub TurnOnLayers()
            'by David Eisenbeisz (c)2024
            'turns on layers that were temporarily turned off

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor


            If m_layList Is Nothing OrElse m_layList.Count <= 0 Then
                ed.WriteMessage(vbLf & "No layers stored in list.  Use LAOFF command to turn off layers and add to list.")
                Exit Sub
            End If

            Dim templist As List(Of String) = m_layList

            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim lyrTbl As LayerTable = actrans.GetObject(DwgDB.LayerTableId, OpenMode.ForRead)
                For Each layName As String In templist
                    If lyrTbl.Has(layName) Then
                        Dim ltr As LayerTableRecord = TryCast(actrans.GetObject(lyrTbl(layName), OpenMode.ForWrite), LayerTableRecord)
                        If ltr IsNot Nothing Then ltr.IsOff = False
                    End If
                Next
                m_layList.Clear()
                actrans.Commit()
            End Using
        End Sub

        <CommandMethod("ExpLayers")>
        Public Sub ExportLayers()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database
            Dim lyrs As New AcdLayers("AcdLayers")
            Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim lTbl As LayerTable = actrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)
                For Each layID As ObjectId In lTbl
                    Dim lTR As LayerTableRecord = TryCast(actrans.GetObject(layID, OpenMode.ForRead), LayerTableRecord)
                    If lTR IsNot Nothing Then
                        If Not lTR.IsDependent Then
                            Dim lName As String = lTR.Name
                            Dim aLayer As New AcdLayer(lName)
                            lyrs.Add(aLayer)
                        End If
                    End If
                Next
            End Using
            FormLayersXML(lyrs)
        End Sub

        <CommandMethod("ImpLayers")>
        Public Sub ImpLayers()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database
            Dim xmlNameStr As String = GetMyXMLFileName()
            If String.IsNullOrEmpty(xmlNameStr) Then Exit Sub
            Dim mySchemaPath As String = "MasterCustomLibrary.LayerSchema1.xsd"

            Dim ss As XmlSchemaSet = GetLayerSchemaSet("")

            Dim sets As New XmlReaderSettings

            Dim badPstyles As New Dictionary(Of String, String)
            Dim hasBadPstyles As Boolean = False

            If ss IsNot Nothing Then
                With sets
                    .ValidationType = ValidationType.Schema
                    .Schemas = ss
                    .CloseInput = True
                    .IgnoreWhitespace = True
                    .Async = False
                End With
            Else
                With sets
                    .CloseInput = True
                    .IgnoreWhitespace = True
                    .Async = False
                End With
            End If


            Dim layset As AcdLayers
            Using xr As XmlReader = XmlReader.Create(xmlNameStr, sets)
                'Dim xsets As New XmlSerializerNamespaces
                'xsets.Add("xs", "http://tempuri.org/LayerSchema1.xsd")
                Dim xmlS As New XmlSerializer(GetType(AcdLayers), "http://tempuri.org/LayerSchema1.xsd")
                layset = xmlS.Deserialize(xr)
            End Using
            Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim lTbl As LayerTable = actrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)
                Dim lz As LayerTableRecord = TryCast(actrans.GetObject(dwgDB.LayerZero, OpenMode.ForRead), LayerTableRecord)

                For i = 0 To layset.Count - 1
                    Dim aLayer As AcdLayer = layset.AcdLayer(i)
                    If String.IsNullOrEmpty(aLayer.Name) Then GoTo Skip
                    Try
                        If lTbl.Has(aLayer.Name) Then
                            Dim lTR As LayerTableRecord = TryCast(actrans.GetObject(lTbl(aLayer.Name), OpenMode.ForRead), LayerTableRecord)
                            If lTR = lz Then GoTo Skip
                            lTR.UpgradeOpen()

                            If aLayer.Remove And Not String.IsNullOrEmpty(aLayer.MergeWith) Then
                                If lTbl.Has(aLayer.MergeWith) Then MergeThenDeleteLayer(aLayer.Name, aLayer.MergeWith, True)
                            ElseIf aLayer.Remove And String.IsNullOrEmpty(aLayer.MergeWith) Then
                                DeleteMyLayer(aLayer.Name)
                            ElseIf Not aLayer.Remove And Not String.IsNullOrEmpty(aLayer.MergeWith) Then
                                If lTbl.Has(aLayer.MergeWith) Then MergeThenDeleteLayer(aLayer.Name, aLayer.MergeWith, False)
                            Else
                                With lTR
                                    .ViewportVisibilityDefault = aLayer.VpVisDefault
                                    .IsOff = aLayer.IsOff
                                    .IsFrozen = aLayer.IsFrozen
                                    .IsLocked = aLayer.IsLocked
                                    .IsPlottable = aLayer.IsPlottable
                                    .IsHidden = aLayer.IsHidden
                                    Dim acdColor As Color = GetColor(aLayer.Color)
                                    .Color = acdColor
                                    .Transparency = GetTransparencyAlpha(CInt(aLayer.Transparency))
                                    .LinetypeObjectId = GetLTId(aLayer.Linetype)
                                    '.PlotStyleName = aLayer.PlotStyle
                                    If Not dwgDB.PlotStyleMode Then
                                        Try
                                            .PlotStyleName = aLayer.PlotStyle
                                        Catch
                                            Err.Clear()
                                            .PlotStyleName = "Normal"
                                            hasBadPstyles = True
                                            badPstyles.Add(aLayer.Name, aLayer.PlotStyle)
                                        End Try
                                    End If
                                    .Description = aLayer.Description
                                End With
                            End If
                        Else
                            Dim lTR As New LayerTableRecord
                            With lTR
                                .Name = aLayer.Name
                                .ViewportVisibilityDefault = aLayer.VpVisDefault
                                .IsOff = aLayer.IsOff
                                .IsFrozen = aLayer.IsFrozen
                                .IsLocked = aLayer.IsLocked
                                .IsPlottable = aLayer.IsPlottable
                                .IsHidden = aLayer.IsHidden
                            End With
                            lTbl.UpgradeOpen()
                            Dim ltrID As ObjectId = lTbl.Add(lTR)
                            actrans.AddNewlyCreatedDBObject(lTR, True)

                            Dim lrec As LayerTableRecord = actrans.GetObject(ltrID, OpenMode.ForWrite)
                            With lrec
                                Dim acdColor As Color = GetColor(aLayer.Color)
                                .Color = acdColor
                                .Transparency = GetTransparencyAlpha(CInt(aLayer.Transparency))
                                .LinetypeObjectId = GetLTId(aLayer.Linetype)
                                .Description = aLayer.Description
                                If Not dwgDB.PlotStyleMode Then
                                    Try
                                        .PlotStyleName = aLayer.PlotStyle
                                    Catch
                                        Err.Clear()
                                        .PlotStyleName = "Normal"
                                        hasBadPstyles = True
                                        badPstyles.Add(aLayer.Name, aLayer.PlotStyle)
                                    End Try
                                End If
                            End With
                        End If

                    Catch ex As Exception
                        If ex.ErrorStatus = ErrorStatus.AmbiguousInput Then
                            Err.Clear()
                        End If
                    End Try
Skip:
                Next
                actrans.Commit()
            End Using

            If hasBadPstyles Then
                Dim dwgName As String = curDwg.Name
                Dim folderName As String = Path.GetDirectoryName(dwgName)
                Dim tFileName As String = folderName & "\BadLayers.txt"
                If File.Exists(tFileName) Then Kill(tFileName)

                Using sr As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(tFileName, False)
                    sr.WriteLine("Layer:Plotstyle")
                    For Each ky As String In badPstyles.Keys
                        Dim outStr As String = ky & ":" & badPstyles(ky)
                        sr.WriteLine(outStr)
                    Next
                End Using

                MessageBox.Show("One or more plotstyles could not be assigned to the imported layers. " _
                            & "This occurs when a plotstyle is assigned to a layer before any objects in the drawing " _
                            & "are assigned that plotstyle.  To work around this bug, ensure that all needed plotstyles " _
                            & "are assigned to temporary objects before importing the layers.  A text file has been " _
                            & "created in the drawing folder that has the names of the failed layers and the plotstyles that could not be assigned.  " _
                            & "The Normal plotstyle has been assigned to these layers.")
            End If

        End Sub

        <CommandMethod("EML")>
        Public Sub EraseMyLayers()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim myLayers As Collection
            Dim myLayer As String

            'declare layer picker dialog
            Dim layBox As New Picker

            'add layer names to text box in layer dialog
            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                With layBox
                    Dim LayList As New List(Of String)
                    Using layTbl As LayerTable = acTrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)
                        Dim lyrRec As LayerTableRecord
                        For Each layID As ObjectId In layTbl
                            lyrRec = acTrans.GetObject(layID, OpenMode.ForRead)
                            LayList.Add(lyrRec.Name)
                        Next
                        LayList.Sort()
                        For Each st As String In LayList
                            .BxList.Items.Add(st)
                        Next
                    End Using
                End With
            End Using

            'show layer dialog and get the layer names from it
            layBox.ShowDialog()
            If layBox.DialogResult = DialogResult.OK Then
                myLayers = layBox.PickCol
            Else
                Exit Sub
            End If

            'dispose of the dialog if it is still in memory
            If layBox IsNot Nothing Then layBox.Dispose()

            For Each myLayer In myLayers
                Call DeleteMyLayer(myLayer)
            Next
        End Sub

        <CommandMethod("MDL")>
        Public Sub MergeLyrs()

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim delLayerCol As Collection

            Dim LayList As New List(Of String)

            'declare layer picker dialog
            Dim layBox As New Picker

            'add layer names to text box in layer dialog
            With layBox
                .TopLabel.Text = "Select Source Layer"
                .Text = "Layer Picker"
                .BxList.SelectionMode = System.Windows.Forms.SelectionMode.One
                Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                    Using layTbl As LayerTable = acTrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)
                        Dim lyrRec As LayerTableRecord
                        For Each layID As ObjectId In layTbl
                            lyrRec = acTrans.GetObject(layID, OpenMode.ForRead)
                            LayList.Add(lyrRec.Name)
                        Next
                    End Using
                End Using
                LayList.Sort()
                For Each st As String In LayList
                    .BxList.Items.Add(st)
                Next
            End With

            'show layer dialog and get the layer name from it
            layBox.ShowDialog()
            If layBox.DialogResult = DialogResult.OK Then
                delLayerCol = layBox.PickCol
            Else
                Exit Sub
            End If

            Dim delLayer As String
            If delLayerCol.Count > 0 Then
                delLayer = delLayerCol(1)
            Else
                Exit Sub
            End If

            With layBox
                .TopLabel.Text = "Select Destination Layer"
                .BxList.ClearSelected()
            End With

            Dim mergeLayerCol As Collection

            'show layer dialog and get the layer name from it
            layBox.ShowDialog()
            If layBox.DialogResult = DialogResult.OK Then
                mergeLayerCol = layBox.PickCol
            Else
                Exit Sub
            End If

            Dim mergeLayer As String
            If mergeLayerCol.Count > 0 Then
                mergeLayer = mergeLayerCol(1)
            Else
                Exit Sub
            End If

            Dim deleteLay As Boolean = YesNoQuery("Do you want to delete the source layer?")
            If deleteLay Then
                MergeThenDeleteLayer(delLayer, mergeLayer, True)
            Else
                MergeThenDeleteLayer(delLayer, mergeLayer, False)
            End If
        End Sub

    End Module

    Public Module PlotLayoutCommands

        <CommandMethod("ListStyleTables")>
        Public Sub PstylesFiles()

            ' Get the current document and database, and start a transaction
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            acDoc.Editor.WriteMessage(vbLf & "Plotstyle Files: ")

            For Each plotStyle As String In PlotSettingsValidator.Current.GetPlotStyleSheetList()
                ' Output the names of the available plot styles
                acDoc.Editor.WriteMessage(vbLf & "  " & plotStyle)

            Next
        End Sub

        <CommandMethod("ListPstyles")>
        Public Sub ListPstyles()

            ' Get the current document and database, and start a transaction
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction()

                'Open Plotstyle table for read
                Dim PSdict As DictionaryWithDefaultDictionary = actrans.GetObject(DwgDB.PlotStyleNameDictionaryId, OpenMode.ForRead)
                Dim i As Integer = PSdict.Count
                Dim PScoll(i) As String

                ed.WriteMessage(vbLf & "Plot styles: ")
                'Dim j As Integer = 0
                For Each DicItem As DBDictionaryEntry In PSdict
                    ed.WriteMessage(vbLf & DicItem.Key.ToString & "   " & DicItem.Value.ToString)
                Next DicItem

            End Using
            Call DumpPltStyleTables()

        End Sub

        <CommandMethod("ShowPStyles")>
        Public Sub PStyleSamples()

            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = CurDwg.Editor

            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction

                Dim cSpace As BlockTableRecord = actrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)

                'Dim fS As New FileStream(tFileName, FileMode.Create, FileAccess.Write)
                Dim PSdict As DictionaryWithDefaultDictionary = actrans.GetObject(DwgDB.PlotStyleNameDictionaryId, OpenMode.ForWrite)
                Dim psList As New Collection

                For Each ps As DBDictionaryEntry In PSdict
                    psList.Add(ps.Key)
                Next

                Dim sX As Double = 0
                Dim eX As Double = 100
                Dim y As Double = -10

                For Each str As String In psList
                    y += 10
                    Dim sPt As New Point2d(sX, y)
                    Dim ePt As New Point2d(eX, y)
                    Dim tpt As New Point3d(sX, y + 0.5, 0)

                    Using pl As New Polyline
                        pl.AddVertexAt(0, sPt, 0, 0, 0)
                        pl.AddVertexAt(1, ePt, 0, 0, 0)
                        pl.PlotStyleName = str
                        cSpace.AppendEntity(pl)
                        actrans.AddNewlyCreatedDBObject(pl, True)
                    End Using

                    Using dt As New DBText
                        With dt
                            .Annotative = 0
                            .Height = 1
                            .Rotation = 0
                            .TextString = str
                            .WidthFactor = 1
                            .Justify = AttachmentPoint.BottomLeft
                            .AlignmentPoint = tpt
                        End With
                        cSpace.AppendEntity(dt)
                        actrans.AddNewlyCreatedDBObject(dt, True)
                    End Using
                Next

                actrans.Commit()

            End Using

        End Sub


        <CommandMethod("IMVS")>
        Public Sub ImageMasterViews()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database
            Dim i As Integer = 0

            Dim ppo As New PromptPointOptions(vbLf & "Pick upper left corner of first cell")
            With ppo
                .AllowNone = False
                .AllowArbitraryInput = True
            End With

            Dim p1 As Point3d
            Dim p2 As Point3d
            Dim userPick As Boolean = True

            Dim ppr As PromptPointResult = ed.GetPoint(ppo)

            If ppr.Status = PromptStatus.OK Then
                p1 = ppr.Value

                Dim pco As New PromptCornerOptions(vbLf & "Pick the lower right corner of first cell.", p1)
                With pco
                    .UseDashedLine = True
                    .AllowArbitraryInput = True
                End With

                Dim pcr As PromptPointResult = ed.GetCorner(pco)

                If pcr.Status = PromptStatus.OK Then
                    p2 = pcr.Value
                Else
                    userPick = False
                End If
            Else
                userPick = False
            End If

            Dim vert As Double
            Dim horiz As Double
            Dim cCols As Integer
            Dim cRows As Integer

            If userPick Then
                horiz = p2.X - p1.X
                vert = p2.Y - p1.Y
            Else
                Dim pdo As New PromptDistanceOptions(vbLf & "Input or pick the horizontal distance for each cell")
                With pdo
                    .AllowNegative = True
                    .Only2d = True
                    .AllowNone = False
                End With

                Dim pdr As PromptDoubleResult = ed.GetDistance(pdo)

                If pdr.Status = PromptStatus.OK Then
                    horiz = pdr.Value
                Else
                    Exit Sub
                End If

                Dim pdo2 As New PromptDistanceOptions(vbLf & "Input or pick the vertical distance for each cell")
                Dim pdr2 As PromptDoubleResult = ed.GetDistance(pdo2)

                If pdr2.Status = PromptStatus.OK Then
                    vert = pdr2.Value
                Else
                    Exit Sub
                End If
            End If


            Dim pdo3 As New PromptIntegerOptions(vbLf & "Input the number of cells per row")
            Dim pdr3 As PromptIntegerResult = ed.GetInteger(pdo3)

            If pdr3.Status = PromptStatus.OK Then
                cCols = pdr3.Value
            Else
                Exit Sub
            End If

            Dim pdo4 As New PromptIntegerOptions(vbLf & "Input the number of cells per column")
            Dim pdr4 As PromptIntegerResult = ed.GetInteger(pdo4)

            If pdr4.Status = PromptStatus.OK Then
                cRows = pdr4.Value
            Else
                Exit Sub
            End If

            Dim makeLayouts As Boolean
            Dim mlStr As String
            Dim pko1 As New PromptKeywordOptions(vbLf & "Do you want to create layouts?")
            With pko1
                .Keywords.Add("Y")
                .Keywords.Add("N")
                .AppendKeywordsToMessage = True
            End With
            Dim pkr1 As PromptResult = ed.GetKeywords(pko1)

            If pkr1.Status = PromptStatus.OK Then
                mlStr = pkr1.StringResult
            Else
                mlStr = "N"
            End If

            If mlStr = "Y" Then
                makeLayouts = True
            Else
                makeLayouts = False
            End If

            Dim cCells As Integer = cRows * cCols

            Dim pko As New PromptKeywordOptions(vbLf & "This command will create " & cCells.ToString & " named views.  Proceed?")
            With pko
                .Keywords.Add("Y")
                .Keywords.Add("N")
                .AppendKeywordsToMessage = True
            End With

            Dim pkr As PromptResult = ed.GetKeywords(pko)

            Dim uR As String
            If pkr.Status = PromptStatus.OK Then
                uR = pkr.StringResult
            Else
                uR = "N"
            End If

            If uR = "N" Then Exit Sub

            Dim hSize As Double = 0
            Dim vSize As Double = 0

            If makeLayouts Then
                Dim pdo5 As New PromptDoubleOptions(vbLf & "Input the width of each Layout (inches)")
                Dim pdr5 As PromptDoubleResult = ed.GetDouble(pdo5)

                If pdr5.Status = PromptStatus.OK Then
                    hSize = pdr5.Value
                Else
                    Exit Sub
                End If

                Dim pdo6 As New PromptDoubleOptions(vbLf & "Input the height of each Layout (inches)")
                Dim pdr6 As PromptDoubleResult = ed.GetDouble(pdo6)

                If pdr6.Status = PromptStatus.OK Then
                    vSize = pdr6.Value
                Else
                    Exit Sub
                End If
            End If

            Dim vpLayer As String

            If LayerExists("vports") Then
                vpLayer = "vports"
            Else
                vpLayer = AddNewLayer("vports", 203)
            End If

            'Dim showForm As Int16 = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("SHOWPAGESETUPFORM")
            Dim cVp As Int16 = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("LAYOUTCREATEVIEWPORT")
            Dim showForm As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("SHOWPAGESETUPFORNEWLAYOUTS")
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("SHOWPAGESETUPFORNEWLAYOUTS", 0)
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("LAYOUTCREATEVIEWPORT", 1)

            Try

                Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction

                    Dim shtNm As String = ""

                    Dim pSet As PlotSettings = Nothing
                    Dim psetVal As PlotSettingsValidator
                    If makeLayouts Then

                        Dim mySetup As String = GetPlotSetup()
                        Dim myPltr As String
                        Dim plsets As DBDictionary = acTrans.GetObject(dwgDB.PlotSettingsDictionaryId, OpenMode.ForRead)

                        If String.IsNullOrEmpty(mySetup) Then
                            Dim setName As String
                            Dim psoPl As New PromptStringOptions(vbLf & "Enter name for new plot settings.")
                            Dim psrPl As PromptResult = ed.GetString(psoPl)
                            If psrPl.Status = PromptStatus.OK Then
                                setName = psrPl.StringResult
                            Else
                                Exit Sub
                            End If
                            If plsets.Contains(setName) Then
                                pSet = plsets.GetAt(setName).GetObject(OpenMode.ForWrite)
                            Else
                                pSet = New PlotSettings(False) With {.PlotSettingsName = setName}
                                pSet.AddToPlotSettingsDictionary(dwgDB)
                                acTrans.AddNewlyCreatedDBObject(pSet, True)
                            End If
                            psetVal = PlotSettingsValidator.Current
                            psetVal.RefreshLists(pSet)
                            myPltr = GetPrinter(psetVal)
                            psetVal.SetPlotConfigurationName(pSet, myPltr, Nothing)
                        Else
                            psetVal = PlotSettingsValidator.Current
                            pSet = plsets.GetAt(mySetup).GetObject(OpenMode.ForWrite)
                            psetVal.RefreshLists(pSet)
                            myPltr = pSet.PlotConfigurationName
                        End If
                        shtNm = GetSheetName(psetVal, pSet)
                    End If

                    Dim vtb As ViewTable = acTrans.GetObject(dwgDB.ViewTableId, OpenMode.ForWrite)
                    Dim viewList As New List(Of Integer)

                    Dim lastNo As Integer = 0
                    Dim startNo As Integer

                    For Each vID As ObjectId In vtb
                        Dim tempV As ViewTableRecord = acTrans.GetObject(vID, OpenMode.ForRead)

                        If IsNumeric(tempV.Name) Then
                            Dim tInt As Integer = CInt(tempV.Name)
                            If tInt > lastNo Then lastNo = tInt
                        End If
                    Next

                    Dim pdo7 As New PromptIntegerOptions(vbLf & "Last numbered vew is " & lastNo.ToString & ". Input the starting view number")
                    With pdo7
                        .DefaultValue = lastNo + 1
                    End With
                    Dim pdr7 As PromptIntegerResult = ed.GetInteger(pdo7)

                    If pdr7.Status = PromptStatus.OK Then
                        startNo = pdr7.Value
                    Else
                        Exit Sub
                    End If

                    i = startNo - 1

                    Dim starthoriz As Double = p1.X + (horiz / 2)
                    Dim startvert As Double = p1.Y + (vert / 2)

                    Dim endHoriz As Double = p1.X + (horiz * cCols) - (horiz / 2)
                    Dim endVert As Double = p1.Y + (vert * cRows) - (vert / 2)

                    For y = 0 To cRows - 1
                        Dim yCtr As Double = startvert + (vert * y)
                        For x As Integer = 0 To cCols - 1
                            Dim xCtr As Double = starthoriz + (horiz * x)
                            'For y As Integer = startvert To endVert Step vert
                            '    For x As Integer = starthoriz To endHoriz Step horiz
                            i += 1
                            Dim vtr As ViewTableRecord
                            If vtb.Has(i.ToString) Then
                                vtr = acTrans.GetObject(vtb(i.ToString), OpenMode.ForWrite)
                                With vtr
                                    .Width = Abs(horiz)
                                    .Height = Abs(vert)
                                    .CenterPoint = New Point2d(xCtr, yCtr)
                                End With
                            Else
                                vtr = New ViewTableRecord
                                With vtr
                                    .Name = i.ToString
                                    .Width = Abs(horiz)
                                    .Height = Abs(vert)
                                    .CenterPoint = New Point2d(xCtr, yCtr)
                                End With
                                Dim vtrid As ObjectId = vtb.Add(vtr)
                                acTrans.AddNewlyCreatedDBObject(vtr, True)
                            End If

                            If makeLayouts AndAlso pSet IsNot Nothing AndAlso shtNm IsNot Nothing Then
                                Dim ptrName As String = pSet.PlotConfigurationName
                                Dim layoutID As ObjectId = CreateVP(vtr, vpLayer, hSize, vSize, acTrans, pSet, shtNm)
                            End If
                        Next
                    Next

                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("SHOWPAGESETUPFORNEWLAYOUTS", showForm)
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("LAYOUTCREATEVIEWPORT", cVp)
                    acTrans.Commit()

                End Using

            Catch ex As Exception
                ed.WriteMessage(vbLf & "Fatal Error in sub ImageMasterViews.")
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("SHOWPAGESETUPFORNEWLAYOUTS", showForm)
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("LAYOUTCREATEVIEWPORT", cVp)
            End Try

        End Sub

        <CommandMethod("RLO")>
        Public Sub RenameLayouts()

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            Dim fName As String = GetMyCSVFileName()
            If String.IsNullOrEmpty(fName) Then Exit Sub

            Dim nmList As New List(Of String)
            Dim fileLine As IEnumerable

            For Each fileLine In File.ReadLines(fName)
                If Path.GetExtension(fName) = ".csv" Then
                    Dim names() As String = Split(fileLine, ",")
                    For i As Integer = 0 To names.Count - 1
                        nmList.Add(names(i))
                    Next
                Else
                    nmList.Add(fileLine)
                End If
            Next

            Dim lPicker As New LayoutPicker
            Dim loList As SortedDictionary(Of Integer, String) = LayoutTabList()

            For Each lN As String In loList.Values
                Dim lName As String = lN
                If Not lName = "Model" Then
                    lPicker.ListBox1.Items.Add(lName)
                End If
            Next

            lPicker.PickerLabel.Text = "Select layouts to update:"

            Dim layoutLst As List(Of String)
            lPicker.ShowDialog()

            If lPicker.DialogResult = DialogResult.Cancel Then
                Exit Sub
            Else
                layoutLst = lPicker.PickedList
            End If

            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim layDict As DBDictionary = dwgDB.LayoutDictionaryId.GetObject(OpenMode.ForRead)
                If layoutLst.Count <= nmList.Count Then
                    For j As Integer = 0 To layoutLst.Count - 1
                        Dim loName As String = layoutLst(j)
                        Dim loID As ObjectId
                        Try
                            If Not layDict(loName) = ObjectId.Null Then
                                loID = (layDict(loName))
                                Dim lo As Layout = acTrans.GetObject(loID, OpenMode.ForWrite)
                                lo.LayoutName = nmList(j)
                            End If
                        Catch ex As Exception
                            ed.WriteMessage(vbLf & "Layout does not exist or error in layout list.  Retry Command.")
                            Exit For
                        End Try
                    Next
                Else
                    ed.WriteMessage(vbLf & "Name list must be longer or equal to number of layouts.")
                    Exit Sub
                End If
                acTrans.Commit()
            End Using

        End Sub

        <CommandMethod("ALTS")>
        Public Sub AdjustLayouts()
            'Command to change the sheet sizes for layouts based upon the IMVS method.  Allows user to pick the layouts to be adjusted.

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            Try

                'start layout manager and transaction
                Dim lm As LayoutManager = LayoutManager.Current
                Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction

                    Dim hSize As Double
                    Dim vSize As Double

                    'get layout sizes from user
                    Dim pdo5 As New PromptDoubleOptions(vbLf & "Input the new width of each Layout Viewport (inches)")
                    Dim pdr5 As PromptDoubleResult = ed.GetDouble(pdo5)

                    If pdr5.Status = PromptStatus.OK Then
                        hSize = pdr5.Value
                    Else
                        Exit Sub
                    End If

                    Dim pdo6 As New PromptDoubleOptions(vbLf & "Input the new height of each Layout Viewport (inches)")
                    Dim pdr6 As PromptDoubleResult = ed.GetDouble(pdo6)

                    If pdr6.Status = PromptStatus.OK Then
                        vSize = pdr6.Value
                    Else
                        Exit Sub
                    End If

                    'get first view to be adjusted (must be numbered views)
                    Dim pdo7 As New PromptIntegerOptions(vbLf & "Input starting view number")
                    Dim pdr7 As PromptIntegerResult = ed.GetInteger(pdo7)

                    Dim startView As Integer

                    If pdr7.Status = PromptStatus.OK Then
                        startView = pdr7.Value
                    Else
                        Exit Sub
                    End If

                    Dim i As Integer = startView - 1

                    'get the view table record
                    Dim vtb As ViewTable = acTrans.GetObject(dwgDB.ViewTableId, OpenMode.ForRead)

                    'pick the layouts to be adjusted
                    Dim lPicker As New LayoutPicker
                    Dim loList As SortedDictionary(Of Integer, String) = LayoutTabList()

                    'add layout names to the picker form
                    For Each lN As String In loList.Values
                        Dim lName As String = lN
                        If Not lName = "Model" Then
                            lPicker.ListBox1.Items.Add(lName)
                        End If
                    Next

                    lPicker.PickerLabel.Text = "Select layouts to update:"

                    'create a list variable and show the form
                    Dim layoutLst As List(Of String)
                    lPicker.ShowDialog()

                    If lPicker.DialogResult = DialogResult.Cancel Then
                        Exit Sub
                    Else
                        layoutLst = lPicker.PickedList
                    End If

                    Dim layDict As DBDictionary = dwgDB.LayoutDictionaryId.GetObject(OpenMode.ForRead)
                    Dim layID As ObjectId = layDict(layoutLst(0))
                    Dim lo As Layout = acTrans.GetObject(layID, OpenMode.ForWrite)

                    'create a plot settings validator to setup layouts
                    Dim pset As PlotSettings
                    Dim psetval As PlotSettingsValidator
                    Dim mySetup As String = GetPlotSetup()
                    Dim myPltr As String
                    Dim plsets As DBDictionary = acTrans.GetObject(dwgDB.PlotSettingsDictionaryId, OpenMode.ForRead)

                    'if there is no existing setup, then create one
                    If mySetup = "" Then
                        Dim setName As String
                        Dim psoPl As New PromptStringOptions(vbLf & "Enter name for new plot settings.")
                        Dim psrPl As PromptResult = ed.GetString(psoPl)
                        If psrPl.Status = PromptStatus.OK Then
                            setName = psrPl.StringResult
                        Else
                            Exit Sub
                        End If
                        If plsets.Contains(setName) Then
                            pset = plsets.GetAt(setName).GetObject(OpenMode.ForWrite)
                        Else
                            pset = New PlotSettings(False) With {.PlotSettingsName = setName}
                            pset.AddToPlotSettingsDictionary(dwgDB)
                            acTrans.AddNewlyCreatedDBObject(pset, True)
                        End If
                        pset.CopyFrom(lo)
                        psetval = PlotSettingsValidator.Current
                        psetval.RefreshLists(pset)
                        myPltr = GetPrinter(psetval)
                        psetval.SetPlotConfigurationName(pset, myPltr, Nothing)
                    Else
                        'if there is a current plot settings validator, use it
                        psetval = PlotSettingsValidator.Current
                        pset = plsets.GetAt(mySetup).GetObject(OpenMode.ForWrite)
                        psetval.RefreshLists(pset)
                        myPltr = pset.PlotConfigurationName
                    End If

                    'create the viewport on the viewports layer
                    Dim vpLayer As String

                    If LayerExists("vports") Then
                        vpLayer = "vports"
                    Else
                        vpLayer = AddNewLayer("vports", 203)
                    End If

                    'Get the sheet name to apply to the layout
                    Dim mySheet As String = GetSheetName(psetval, pset)
                    If mySheet = "" Then Exit Sub
                    'add the plot settings, printer, and sheet name to the plot settings validator
                    psetval.SetPlotConfigurationName(pset, myPltr, mySheet)

                    'run through the layout list
                    If layoutLst.Count > 0 Then
                        For Each loName As String In layoutLst
                            lm.CurrentLayout = loName
                            Dim loID As ObjectId = layDict(loName)
                            lo = acTrans.GetObject(loID, OpenMode.ForWrite)

                            'get the second viewport on the layout (first viewport is paperspace)
                            Dim vpIDs As ObjectIdCollection = lo.GetViewports
                            Dim vp As Autodesk.AutoCAD.DatabaseServices.Viewport = acTrans.GetObject(vpIDs(1), OpenMode.ForWrite)
                            Dim curSpace As BlockTableRecord = acTrans.GetObject(dwgDB.CurrentSpaceId, OpenMode.ForWrite)

                            'change the viewport dimensions to fill the sheet
                            'Dim vp As New Viewport
                            vp.SetDatabaseDefaults()
                            vp.CenterPoint = New Point3d(hSize / 2, vSize / 2, 0)
                            vp.Height = vSize
                            vp.Width = hSize
                            vp.Layer = vpLayer

                            lo.CopyFrom(pset)
                            'Dim pSetVal As PlotSettingsValidator = PlotSettingsValidator.Current
                            Dim check As String = pset.PlotConfigurationName
                            'Debug.Print(check)
                            psetval.SetPlotConfigurationName(pset, myPltr, mySheet)
                            psetval.SetPlotType(pset, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout)
                            psetval.SetPlotRotation(pset, PlotRotation.Degrees000)
                            psetval.SetZoomToPaperOnUpdate(pset, True)
                            i += 1

                            'if there is a numbered view corresponding to the layout view, use it
                            If vtb.Has(i.ToString) Then
                                Dim vtr As ViewTableRecord = acTrans.GetObject(vtb(i.ToString), OpenMode.ForRead)
                                ed.SwitchToModelSpace()
                                ed.SetCurrentView(vtr)
                                ed.SwitchToPaperSpace()
                            End If

                        Next
                    End If
                    acTrans.Commit()
                End Using

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub





        <CommandMethod("MSPT")>
        Public Sub ModelSpacePlot()
            'plots all block references in model space

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            Try

                Dim lm As LayoutManager = LayoutManager.Current
                Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction

                    Dim BlkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                    Dim mdlSpace As BlockTableRecord = acTrans.GetObject(BlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    Dim pset As PlotSettings
                    Dim psetval As PlotSettingsValidator
                    Dim mySetup As String = GetPlotSetup()
                    Dim plsets As DBDictionary = acTrans.GetObject(dwgDB.PlotSettingsDictionaryId, OpenMode.ForRead)
                    'Dim setName As String = "ModelLayout"
                    Dim setName As String = mySetup

                    pset = plsets.GetAt(setName).GetObject(OpenMode.ForWrite)
                    psetval = PlotSettingsValidator.Current
                    Dim myPltr As String = GetPrinter(psetval)
                    psetval.RefreshLists(pset)
                    psetval.SetUseStandardScale(pset, True)
                    psetval.SetStdScaleType(pset, StdScaleType.ScaleToFit)
                    psetval.SetPlotRotation(pset, PlotRotation.Degrees000)
                    Dim mySheet As String = GetSheetName(psetval, pset)
                    psetval.SetPlotConfigurationName(pset, myPltr, mySheet)

                    For Each obID As ObjectId In mdlSpace
                        Dim minPt2d As Point2d
                        Dim maxPt2d As Point2d
                        Dim dbObj As DBObject = acTrans.GetObject(obID, OpenMode.ForRead)
                        Dim myFileName As String
                        If TypeOf dbObj Is BlockReference Then
                            Dim bRef As BlockReference = CType(dbObj, BlockReference)
                            If Not bRef.Name = "inspt" Then
                                Dim geoExt As Extents3d = bRef.GeometricExtents
                                Dim minPt As Point3d = geoExt.MinPoint
                                Dim maxpt As Point3d = geoExt.MaxPoint
                                minPt2d = New Point2d(minPt.X, minPt.Y)
                                maxPt2d = New Point2d(maxpt.X, maxpt.Y)
                                Dim curDwgName As String = curDwg.Name
                                Dim curDwgPath As String = Path.GetDirectoryName(curDwgName)
                                Dim newPath As String = curDwgPath & "\PNG\PNGtest\"
                                Dim blkName As String
                                If bRef.IsDynamicBlock Then
                                    Dim blkBTRid As ObjectId = bRef.DynamicBlockTableRecord
                                    Dim blkBTR As BlockTableRecord = acTrans.GetObject(blkBTRid, OpenMode.ForRead)
                                    Dim pName As String = ""
                                    Dim fState As String = ""
                                    For Each prop As DynamicBlockReferenceProperty In bRef.DynamicBlockReferencePropertyCollection
                                        If prop.PropertyName = "Visibility1" Then
                                            pName = " " & prop.Value
                                        ElseIf prop.PropertyName = "Flip state1" Then
                                            fState = " " & prop.Value
                                        End If
                                    Next
                                    blkName = blkBTR.Name & pName & fState
                                Else
                                    blkName = bRef.Name
                                End If
                                Debug.Print(blkName)
                                blkName = ReturnValidFileName(blkName)
                                Dim outFile As String = newPath & blkName
                                outFile &= ".png"
                                myFileName = GetNewFileName(outFile)
                                psetval.SetPlotWindowArea(pset, New Extents2d(minPt2d, maxPt2d))
                                PlotDirect(pset, psetval, myFileName)
                            End If
                        End If
                    Next
                    acTrans.Commit()
                End Using

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

    End Module
    Public Module MiscCommands

        <CommandMethod("LLTS")>
        Public Sub ListLinetypes()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            Dim i As Integer = 0
            Dim ly As Double = 0
            Dim lx As Double = 200
            Dim l2y As Double = -100
            Dim l2x As Double = 225

            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim ltTbl As LinetypeTable = acTrans.GetObject(dwgDB.LinetypeTableId, OpenMode.ForWrite)
                Dim blkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                For Each obID As ObjectId In ltTbl
                    Using ltRec As LinetypeTableRecord = acTrans.GetObject(obID, OpenMode.ForRead)

                        Dim ltName As String = ltRec.Name
                        Dim ltRecId = obID
                        Using pl As New Polyline(2)
                            pl.AddVertexAt(0, New Point2d(0, ly), 0, 0, 0)
                            pl.AddVertexAt(1, New Point2d(lx, ly), 0, 0, 0)
                            pl.Linetype = ltName
                            pl.Plinegen = True
                            mdlSpace.AppendEntity(pl)
                            acTrans.AddNewlyCreatedDBObject(pl, True)
                        End Using
                        Using dt As New DBText
                            With dt
                                .Annotative = 0
                                .Height = 1
                                '.Position = New Point3d(0, ly + 0.1, 0)
                                '.TextStyleName = ltName
                                .Rotation = 0
                                .TextString = ltName
                                .WidthFactor = 1
                                .Justify = AttachmentPoint.BottomLeft
                                .AlignmentPoint = New Point3d(0, ly + 0.1, 0)
                            End With
                            mdlSpace.AppendEntity(dt)
                            acTrans.AddNewlyCreatedDBObject(dt, True)
                        End Using
                        Using pl2 As New Polyline(2)
                            pl2.AddVertexAt(0, New Point2d(l2x, -1.1), 0, 0, 0)
                            pl2.AddVertexAt(1, New Point2d(l2x, -201.1), 0, 0, 0)
                            pl2.Linetype = ltName
                            pl2.Plinegen = True
                            mdlSpace.AppendEntity(pl2)
                            acTrans.AddNewlyCreatedDBObject(pl2, True)
                        End Using
                        Using dt2 As New DBText
                            With dt2
                                .Annotative = 0
                                .Height = 1
                                '.Position = New Point3d(0, ly + 0.1, 0)
                                '.TextStyleName = ltName
                                .Rotation = 0
                                .TextString = ltName
                                .WidthFactor = 1
                                .Justify = AttachmentPoint.BottomCenter
                                .AlignmentPoint = New Point3d(l2x, -1, 0)
                            End With
                            mdlSpace.AppendEntity(dt2)
                            acTrans.AddNewlyCreatedDBObject(dt2, True)
                        End Using

                        ly -= 10
                        l2x += 20
                    End Using
                Next
                acTrans.Commit()
            End Using
        End Sub

        <CommandMethod("LLAYS")>
        Public Sub ListLayers()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            Dim i As Integer = 0
            Dim ly As Double = 0
            Dim lx As Double = 0
            Dim miny As Long = -220

            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim layTbl As LayerTable = dwgDB.LayerTableId.GetObject(OpenMode.ForRead)
                Dim blkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim ltTbl As LinetypeTable = acTrans.GetObject(dwgDB.LinetypeTableId, OpenMode.ForRead)
                For Each obID As ObjectId In layTbl
                    Dim layRec As LayerTableRecord = acTrans.GetObject(obID, OpenMode.ForRead)
                    Dim layName As String = layRec.Name
                    Dim layRecId As ObjectId = obID
                    Using pl As New Polyline(2)
                        pl.AddVertexAt(0, New Point2d(lx, ly), 0, 0, 0)
                        pl.AddVertexAt(1, New Point2d(lx + 50, ly), 0, 0, 0)
                        pl.Plinegen = True
                        mdlSpace.AppendEntity(pl)
                        acTrans.AddNewlyCreatedDBObject(pl, True)
                        pl.LayerId = obID
                    End Using
                    Using dt As New DBText
                        With dt
                            .Annotative = 0
                            .Height = 1
                            '.Position = New Point3d(0, ly + 0.1, 0)
                            '.TextStyleName = ltName
                            .LayerId = obID
                            .Rotation = 0
                            .TextString = layName
                            .WidthFactor = 1
                            .Justify = AttachmentPoint.BottomLeft
                            .AlignmentPoint = New Point3d(lx, ly + 0.1, 0)
                        End With
                        mdlSpace.AppendEntity(dt)
                        acTrans.AddNewlyCreatedDBObject(dt, True)
                    End Using
                    If ly <= miny Then
                        ly = 0
                        lx += 60
                    Else
                        ly -= 10
                    End If
                Next
                acTrans.Commit()
            End Using

        End Sub


        <CommandMethod("LSTS")>
        Public Sub ListTextStyles()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim tst As TextStyleTable = actrans.GetObject(dwgDB.TextStyleTableId, OpenMode.ForRead)
                For Each tsID As ObjectId In tst
                    Dim tstr As TextStyleTableRecord = actrans.GetObject(tsID, OpenMode.ForRead)
                    Dim chkName As String = tstr.FileName
                    Dim tsName As String = tstr.Name
                    Dim hasFields As Boolean = tstr.HasFields
                    Dim isShape As Boolean = tstr.IsShapeFile
                    If Not tsName = "" Then
                        Debug.Print(vbCrLf & tsName & " HasFields: " & hasFields.ToString & " IsShape: " & isShape.ToString)
                    Else
                        Debug.Print(vbCrLf & chkName & " HasFields: " & hasFields.ToString & " IsShape: " & isShape.ToString)
                    End If
                Next
            End Using
        End Sub

        <CommandMethod("EXLTS")>
        Public Sub ExtractLineTypes()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database
            Dim linetypes As New List(Of LTdata)

            Dim tFileName As String = Path.GetTempFileName

            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim ltTbl As LinetypeTable = acTrans.GetObject(dwgDB.LinetypeTableId, OpenMode.ForWrite)
                Dim blkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                For Each obID As ObjectId In ltTbl
                    Dim ltd As New LTdata(obID)
                    linetypes.Add(ltd)
                Next
                Using sr As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(tFileName, False)
                    For j As Integer = 0 To linetypes.Count - 1
                        Dim ltype As LTdata = linetypes(j)
                        Dim numDashes As Double = ltype.Dashes
                        Dim dLengths() As Double = ltype.DLengths
                        Dim shapeDash As Integer = 0
                        Dim shapeNum As Integer = 0
                        Dim shapefile As String = ""
                        Dim shapestyleid As ObjectId
                        Dim ltDesc As String = ltype.Description
                        Dim ltName As String = ltype.Name
                        If ltName <> "Continuous" AndAlso ltName <> "ByLayer" AndAlso ltName <> "ByBlock" Then
                            Dim sb As New StringBuilder
                            sb.Append("*" & UCase(ltName) & "," & ltDesc)
                            sr.WriteLine(sb.ToString)
                            Dim sb2 As New StringBuilder
                            sb2.Append("A,")
                            For k As Integer = 0 To numDashes - 1
                                Dim dl As Double = dLengths(k)
                                If dl = 1000 Then
                                    If ltype.HasShape Then
                                        shapeDash = ltype.SymbolDash
                                        shapeNum = ltype.ShapeNumber
                                        shapefile = ltype.ShapeFile
                                        shapestyleid = ltype.TextStyleID
                                        Dim Charactr As Char = Chr(shapeNum)
                                        Dim shp As New Shape(New Point3d(0, 0, 0), 1, 0, 1)
                                        mdlSpace.AppendEntity(shp)
                                        acTrans.AddNewlyCreatedDBObject(shp, True)

                                        shp.StyleId = shapestyleid
                                        shp.ShapeNumber = ltype.ShapeNumber
                                        Dim shpName As String = shp.Name
                                        'Dim shpName As String = DxfCode.ShapeName
                                        sb2.Append("[" & shpName & "," & shapefile & ",S=" & ltype.ShapeScale.ToString)
                                        If ltype.HasUCSOrient Then
                                            sb2.Append("," & "A=" & ltype.ShapeRotation.ToString)
                                        Else
                                            sb2.Append("," & "U=" & ltype.ShapeRotation.ToString)
                                        End If

                                        If ltype.ShapeX <> 0 Then sb2.Append("," & ltype.ShapeX.ToString)
                                        If ltype.ShapeY <> 0 Then sb2.Append("," & ltype.ShapeY.ToString)
                                        sb2.Append("],")
                                        shp.Erase()

                                    ElseIf ltype.HasText Then
                                        Dim textStr As String = ltype.TextString
                                        Dim textstyleID As ObjectId = ltype.TextStyleID
                                        Dim textStyleName As String = ltype.TextStyleName
                                        sb2.Append("[" & Chr(34) & textStr & Chr(34) & "," & textStyleName & ",S=" & ltype.ShapeScale.ToString)

                                        If ltype.HasUCSOrient Then
                                            sb2.Append("," & "U=" & ltype.ShapeRotation.ToString)
                                        Else
                                            sb2.Append("," & "R=" & ltype.ShapeRotation.ToString)
                                        End If
                                        If ltype.ShapeX <> 0 Then sb2.Append("," & ltype.ShapeX.ToString)
                                        If ltype.ShapeY <> 0 Then sb2.Append("," & ltype.ShapeY.ToString)
                                        sb2.Append("],")
                                    End If
                                Else
                                    If IsNumeric(dl) Then sb2.Append(dl.ToString & ",")
                                End If
                            Next

                            Dim ltDef As String
                            Dim line2 As String = sb2.ToString
                            If Strings.Right(line2, 1) = "," Then
                                ltDef = Strings.Left(line2, Strings.Len(line2) - 1)
                            Else
                                ltDef = line2
                            End If
                            sr.WriteLine(ltDef.ToString)
                            sr.WriteLine()
                        End If
                    Next
                End Using
                acTrans.Commit()

            End Using
            Dim saveFName As String = GetASaveFileName("lin", "txt")
            If saveFName Is Nothing Then
                Exit Sub
            ElseIf String.IsNullOrEmpty(saveFName) Then
                Exit Sub
            ElseIf File.Exists(saveFName) Then
                Kill(saveFName)
                File.Copy(tFileName, saveFName)
                Kill(tFileName)
            Else
                File.Copy(tFileName, saveFName)
                Kill(tFileName)
            End If

        End Sub

        <CommandMethod("MAH")>
        Public Sub MkArrowHead()
            'Creates a standard sign arrow head

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            'declare two points for final arrow position and orientation
            Dim p1 As Point3d
            Dim p2 As Point3d

            'get the tip location and store as p1
            Dim ppo As New PromptPointOptions(vbLf & "Pick tip of arrowhead")
            With ppo
                .AllowNone = False
                .AllowArbitraryInput = True
            End With

            Dim ppr As PromptPointResult = ed.GetPoint(ppo)

            If ppr.Status = PromptStatus.OK Then
                p1 = ppr.Value
            Else
                ed.WriteMessage(vbLf & "Command cancelled.")
                Exit Sub
            End If

            Dim pko As New PromptKeywordOptions(vbLf & "Do you want to draw the arrow shaft?")
            With pko
                .AllowNone = False
                .Keywords.Add("Yes")
                .Keywords.Add("No")
                .AppendKeywordsToMessage = True
                .AllowArbitraryInput = False
            End With

            Dim drawShaft As Boolean

            Dim pkr As PromptResult = ed.GetKeywords(pko)
            If pkr.Status = PromptStatus.OK Then
                If pkr.StringResult = "Yes" Then
                    drawShaft = True
                Else
                    drawShaft = False
                End If
            Else
                ed.WriteMessage(vbLf & "Command cancelled.")
                Exit Sub
            End If


            'get second point for arrow orientation
            Dim ppo2 As New PromptPointOptions(vbLf & "Pick a point for the direction of the shaft")
            With ppo2
                If drawShaft Then .Message = vbLf & "Pick the end of the arrow shaft"
                .UseBasePoint = True
                .BasePoint = p1
                .UseDashedLine = True
            End With

            Dim ppr2 As PromptPointResult = ed.GetCorner(ppo2)

            If ppr2.Status = PromptStatus.OK Then
                p2 = ppr2.Value
            Else
                Exit Sub
            End If

            Dim arrLen As Double = p1.DistanceTo(p2)

            'create a vector2d from p1 & p2 to get the final orientation angle for the arrowhead
            Dim arrowVect As Vector3d = p1.GetVectorTo(p2)
            Dim arrowVect2d As Vector2d = arrowVect.Convert2d(New Plane(Point3d.Origin, Vector3d.ZAxis))
            Dim arrowAng As Double = arrowVect2d.Angle

            'get the width of the arrow shaft - this is the critical parameter for the arrowhead
            Dim pdo As New PromptDistanceOptions(vbLf & "Enter or pick the width of the shaft.")
            With pdo
                .AllowNegative = False
                .Only2d = True
                .AllowNone = False
            End With

            Dim pdr As PromptDoubleResult = ed.GetDistance(pdo)

            Dim paramA As Double

            If pdr.Status = PromptStatus.OK Then
                paramA = pdr.Value
            Else
                Exit Sub
            End If

            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                'using an origin located on the cenerline of the arrow at the end of the arrow shaft,
                'get temporary arrow coordinate geometry per Federal Sign Manual Appendix
                Dim paramB As Double = 1.21 * paramA
                Dim paramC As Double = 2 * paramA
                Dim paramE As Double = 0.21 * paramA
                Dim theta1 As Double = Asin(paramE / (paramB - (paramA / 2)))
                Dim paramD As Double = (paramA / 2) * Tan(theta1)
                Dim ptO As Point3d = Point3d.Origin
                Dim s1 As New Point2d(ptO.X, paramA / 2)
                Dim s2 As New Point2d(ptO.X, -paramA / 2)
                Dim cen1 As New Point3d(ptO.X, paramB, 0)
                Dim ptPx As Double = paramE * Cos(theta1)
                Dim ptPy As Double = paramB - (paramE * Sin(theta1))
                Dim ptP2d As New Point2d(ptPx, ptPy)
                Dim ptR As New Point3d(-paramC, ptO.Y, ptO.Z)
                Dim ptR2d As New Point2d(ptR.X, ptR.Y)
                Dim negPtP As New Point2d(ptPx, -ptPy)


                'store the vector from the tip of the temporary arrow to its final location
                Dim moveVect As Vector3d = ptR.GetVectorTo(p1)

                'create a temporary circle for finding tangents
                Dim c1 As New Circle(cen1, Vector3d.ZAxis, paramE)

                'get the tangent points from ptR to c1
                Dim iPts As Point2dCollection = GetTangentPoints(ptR, c1)

                'make sure that the tangent points exist and weed out the one that does not apply
                Dim tanPt As Point2d
                If iPts IsNot Nothing AndAlso iPts.Count = 2 Then
                    If iPts(0).Y > 0 Then
                        tanPt = iPts(0)
                    Else
                        tanPt = iPts(1)
                    End If
                ElseIf iPts IsNot Nothing AndAlso iPts.Count = 1 Then
                    ed.WriteMessage(vbLf & "Error in Sub MkArrowHead")
                    Exit Sub
                ElseIf iPts Is Nothing Then
                    ed.WriteMessage(vbLf & "Error in Sub MkArrowHead")
                    Exit Sub
                End If

                'set the opposite tangent point
                Dim negTanPt As New Point2d(tanPt.X, -tanPt.Y)

                'get the bulge value for the polyline
                Dim ahv1 As Vector2d = ptR2d.GetVectorTo(tanPt)
                Dim ahv2 As Vector2d = ptP2d.GetVectorTo(s1)
                Dim bulgeAng As Double = ahv1.GetAngleTo(ahv2)
                Dim myblg As Double = Tan(bulgeAng / 4)
                Dim ptF As New Point2d(arrLen - paramC, paramA / 2)
                Dim negPtF As New Point2d(arrLen - paramC, -paramA / 2)

                'create a polyline
                If drawShaft Then
                    Using pl0 As New Polyline
                        pl0.AddVertexAt(0, ptF, 0, 0, 0)
                        pl0.AddVertexAt(1, s1, 0, 0, 0)
                        pl0.AddVertexAt(2, ptP2d, myblg, 0, 0)
                        pl0.AddVertexAt(3, tanPt, 0, 0, 0)
                        pl0.AddVertexAt(4, ptR2d, 0, 0, 0)
                        pl0.AddVertexAt(5, negTanPt, myblg, 0, 0)
                        pl0.AddVertexAt(6, negPtP, 0, 0, 0)
                        pl0.AddVertexAt(7, s2, 0, 0, 0)
                        pl0.AddVertexAt(8, negPtF, 0, 0, 0)
                        pl0.Closed = True

                        'move and rotate the polyline to its final position
                        pl0.TransformBy(Matrix3d.Displacement(moveVect))
                        pl0.TransformBy(Matrix3d.Rotation(arrowAng, Vector3d.ZAxis, p1))

                        'add it to the current space
                        Dim curSpace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)
                        curSpace.AppendEntity(pl0)
                        acTrans.AddNewlyCreatedDBObject(pl0, True)
                    End Using

                Else
                    Using pl0 As New Polyline
                        pl0.AddVertexAt(0, s1, 0, 0, 0)
                        pl0.AddVertexAt(1, ptP2d, myblg, 0, 0)
                        pl0.AddVertexAt(2, tanPt, 0, 0, 0)
                        pl0.AddVertexAt(3, ptR2d, 0, 0, 0)
                        pl0.AddVertexAt(4, negTanPt, myblg, 0, 0)
                        pl0.AddVertexAt(5, negPtP, 0, 0, 0)
                        pl0.AddVertexAt(6, s2, 0, 0, 0)

                        'move and rotate the polyline to its final position
                        pl0.TransformBy(Matrix3d.Displacement(moveVect))
                        pl0.TransformBy(Matrix3d.Rotation(arrowAng, Vector3d.ZAxis, p1))

                        'add it to the current space
                        Dim curSpace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)
                        curSpace.AppendEntity(pl0)
                        acTrans.AddNewlyCreatedDBObject(pl0, True)
                    End Using
                End If

                'dispose of the temporary circle
                c1.Dispose()

                acTrans.Commit()
            End Using

        End Sub



        <CommandMethod("MDAR")>
        Public Sub MakeDirectionalArrows()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            'get arrow type from user
            Dim atb As New ArrowTypeBox
            atb.ShowDialog()
            Dim aType As Integer = atb.ArrowType

            'test for good type
            If aType = 0 Then Exit Sub

            'get the final tip location and store as p1
            Dim p1 As Point3d
            Dim ppo As New PromptPointOptions(vbLf & "Pick tip of arrowhead")
            With ppo
                .AllowNone = False
                .AllowArbitraryInput = True
            End With

            Dim ppr As PromptPointResult = ed.GetPoint(ppo)

            If ppr.Status = PromptStatus.OK Then
                p1 = ppr.Value
            Else
                ed.WriteMessage(vbLf & "Command cancelled.")
                Exit Sub
            End If

            'get second point for orientation of arrows type 1 and 3
            Dim arrowAng As Double
            Dim p2 As Point3d

            If aType = 1 Or aType = 3 Then
                Dim ppo2 As New PromptPointOptions(vbLf & "Pick a point for the direction of the arrow shaft")
                With ppo2
                    .UseBasePoint = True
                    .BasePoint = p1
                    .UseDashedLine = True
                End With

                Dim ppr2 As PromptPointResult = ed.GetCorner(ppo2)

                If ppr2.Status = PromptStatus.OK Then
                    p2 = ppr2.Value
                Else
                    Exit Sub
                End If

                'create a vector2d from p1 & p2 to get the final orientation angle for the arrowhead
                Dim arrowVect As Vector3d = p1.GetVectorTo(p2)
                Dim arrowVect2d As Vector2d = arrowVect.Convert2d(New Plane(Point3d.Origin, Vector3d.ZAxis))
                arrowAng = arrowVect2d.Angle

                'pick right or left for arrow type 2
            ElseIf aType = 2 Then
                Dim pko As New PromptKeywordOptions(vbLf & "Arrow pointing Left or Right?")
                With pko
                    .Keywords.Add("Left")
                    .Keywords.Add("Right")
                    .AppendKeywordsToMessage = True
                    .AllowNone = False
                    .AllowArbitraryInput = False
                End With

                Dim pkr As PromptResult = ed.GetKeywords(pko)

                If pkr.Status = PromptStatus.OK Then
                    If pkr.StringResult = "Left" Then
                        arrowAng = 0
                    Else
                        arrowAng = PI
                    End If
                Else
                    ed.WriteMessage(vbLf & "Command ended.")
                    Exit Sub
                End If

                'type 4 always points down
            ElseIf aType = 4 Then
                arrowAng = PI / 2
            Else
                Exit Sub
            End If


            'Set arrow parameters by letter height or best parameter dimension.  See California Sign Manual Appendix for standard letter heights and parameter definitions.
            Dim useLetters As Boolean
            Dim pko2 As New PromptKeywordOptions("")
            Dim msgText2 As String = ""

            'Type 1 & 3 use parameter C
            If aType = 1 Or aType = 3 Then
                With pko2
                    .Message = vbLf & "Set size by letter height or Parameter C?"
                    .Keywords.Add("Height")
                    .Keywords.Add("C")
                    .AppendKeywordsToMessage = True
                    .AllowNone = False
                    .AllowArbitraryInput = False
                End With

                'Type 2 uses parameter D
            ElseIf aType = 2 Then
                With pko2
                    .Message = vbLf & "Set size by letter height or Parameter D?"
                    .Keywords.Add("Height")
                    .Keywords.Add("D")
                    .AppendKeywordsToMessage = True
                    .AllowNone = False
                    .AllowArbitraryInput = False
                End With

                'Type 4 has only two sizes
            ElseIf aType = 4 Then
                useLetters = False
                With pko2
                    .Message = vbLf & "Arrow width (parameter A) 24 or 32?"
                    .Keywords.Add("24")
                    .Keywords.Add("32")
                    .AppendKeywordsToMessage = True
                    .AllowNone = False
                    .AllowArbitraryInput = False
                End With
            End If

            Dim pkr2 As PromptResult = ed.GetKeywords(pko2)
            Dim sizeValue As Double

            If pkr2.Status = PromptStatus.OK Then
                If pkr2.StringResult = "Height" Then
                    useLetters = True
                    msgText2 = vbLf & "Enter value for text height:"
                ElseIf pkr2.StringResult = "C" Then
                    msgText2 = vbLf & "Enter value for Parameter C:"
                ElseIf pkr2.StringResult = "D" Then
                    msgText2 = vbLf & "Enter value for Parameter D:"
                Else
                    If IsNumeric(pkr2.StringResult) Then
                        sizeValue = CDbl(pkr2.StringResult)
                    Else
                        ed.WriteMessage(vbLf & "Invalid arrow size.  Ending Command.")
                        Exit Sub
                    End If
                    useLetters = False
                End If
            Else
                ed.WriteMessage(vbLf & "Command ended.")
                Exit Sub
            End If

            If Not aType = 4 Then
                Dim pdbo As New PromptDoubleOptions(msgText2)
                With pdbo
                    .AllowNegative = False
                    .AllowZero = False
                    .DefaultValue = 8.0
                End With

                Dim pdbr As PromptDoubleResult = ed.GetDouble(pdbo)

                If pdbr.Status = PromptStatus.OK Then
                    sizeValue = pdbr.Value
                Else
                    Exit Sub
                End If
            End If

            'create a new arrow object and populate fields
            Dim ao As ArrowObj
            ao = New ArrowObj(aType, sizeValue, useLetters, arrowAng)

            'send to appropriate sub to create the arrow polyline
            If aType = 4 Then
                MkDirectionalArrowType4(p1, ao)
            Else
                MkDirectionalArrowType1(p1, ao)
            End If

        End Sub

        <CommandMethod("RABCS", (CommandFlags.Modal Or CommandFlags.UsePickSet))>
        Public Sub RemoveAllButCurrentScale()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDb As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor
            Dim ocm As ObjectContextManager = dwgDb.ObjectContextManager
            Dim occ As ObjectContextCollection = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES")
            Dim pso As PromptSelectionOptions = New PromptSelectionOptions()
            pso.MessageForAdding = vbLf & "Select annotative objects"
            Dim psr As PromptSelectionResult = ed.GetSelection(pso)
            If psr.Status <> PromptStatus.OK Then Return
            Dim objCount As Integer = 0, scaCount As Integer = 0
            Dim scalesRemovedForObject As Boolean = False
            Dim tr As Transaction = curDwg.TransactionManager.StartTransaction()

            Using tr

                If Not occ.HasContext(dwgDb.Cannoscale.Name) Then
                    ed.WriteMessage(vbLf & "Cannot find current annotation scale.")
                    Return
                End If

                Dim curCtxt As ObjectContext = occ.GetContext(dwgDb.Cannoscale.Name)

                For Each so As SelectedObject In psr.Value
                    Dim id As ObjectId = so.ObjectId
                    Dim obj As DBObject = tr.GetObject(id, OpenMode.ForRead)

                    If obj.Annotative = AnnotativeStates.[True] AndAlso obj.HasContext(curCtxt) Then
                        obj.UpgradeOpen()

                        For Each oc As ObjectContext In occ

                            If obj.HasContext(oc) AndAlso oc.Name <> dwgDb.Cannoscale.Name Then
                                obj.RemoveContext(oc)
                                scaCount += 1
                                scalesRemovedForObject = True
                            End If
                        Next

                        If scalesRemovedForObject Then
                            objCount += 1
                            scalesRemovedForObject = False
                        End If
                    End If
                Next

                tr.Commit()
                ed.WriteMessage(vbLf & "{0} scales removed from {1} objects.", scaCount, objCount)
            End Using
        End Sub

    End Module

End Namespace

