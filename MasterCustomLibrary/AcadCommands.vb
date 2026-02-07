Imports System.CodeDom
Imports System.ComponentModel.Design
Imports System.IO
Imports System.Math
Imports System.Reflection
Imports System.Security.RightsManagement
Imports System.Text
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Windows.Media.Animation
Imports System.Windows.Media.Effects
'Imports System.Windows.Shapes
Imports System.Xml
Imports System.Xml.Schema
Imports System.Xml.Serialization
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.ApplicationServices.DatabaseExtension
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsSystem
Imports Autodesk.AutoCAD.Internal
Imports Autodesk.AutoCAD.Runtime
Imports MasterCustomLibrary.AcCommon
Imports Microsoft.VisualBasic.FileIO

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
                .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Polyline), True)
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

        <CommandMethod("LISTARCDATA", CommandFlags.UsePickSet)>
        Public Sub ListArcData()
            'by David Eisenbeisz

            Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = acDwg.Database
            Dim ed As Editor = acDwg.Editor
            Dim mypt1 As Point3d
            Dim mypt2 As Point3d
            Dim mypt3 As Point3d
            Dim myArcID As ObjectId
            Dim aD As New ArcData
            Dim pickedFirst As Boolean
            Dim arcCol As New Collection

            Dim SelResult As PromptSelectionResult = ed.SelectImplied()
            Dim acSSet As SelectionSet
            Dim entObjID As ObjectId

            If SelResult.Status = PromptStatus.OK Then
                acSSet = SelResult.Value
                If acSSet.Count > 0 Then
                    Dim MyobjIDs() As ObjectId = acSSet.GetObjectIds
                    entObjID = MyobjIDs(0)
                End If
                pickedFirst = True
            Else
                Dim pp0 As New PromptPointOptions(vbLf & "Select first point on arc or press escape to select arc entities: ")
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
            End If

            If pickedFirst Then
                If Not entObjID = ObjectId.Null Then
                    Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                        Dim myEnt As Entity = acTrans.GetObject(entObjID, OpenMode.ForRead)
                        If TypeOf myEnt Is Arc Then
                            Dim myArc As Arc = TryCast(myEnt, Arc)
                            If myArc Is Nothing Then Exit Sub
                            Dim midPt As Point3d = myArc.GetPointAtDist(myArc.Length / 2)
                            Dim tpt1 As New Point2d(myArc.StartPoint.X, myArc.StartPoint.Y)
                            Dim tpt3 As New Point2d(myArc.EndPoint.X, myArc.EndPoint.Y)
                            Dim tpt2 As New Point2d(midPt.X, midPt.Y)
                            Dim acArc As New CircularArc2d(tpt1, tpt2, tpt3)
                            aD = New ArcData(acArc, mypt1.Z)
                        End If
                    End Using
                End If
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

        Friend m_fldr As String

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

        <CommandMethod("CHBLKCOLOR")>
        Public Sub ChangeBlockColor()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim docMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            ed.WriteMessage("Select folder containing blocks to change:")
            Dim fldr As String = GetMyFolderName()

            Dim fi As FileInfo() = AllDwgFilesInFolder(fldr, False)

            If fi Is Nothing Then Exit Sub
            Dim oldClr As Color = Nothing
            Dim newClr As Color = Nothing

            Try
                ed.WriteMessage(vbLf & "Pick color to replace...")

                Dim clrdia As New Autodesk.AutoCAD.Windows.ColorDialog
                Dim clrRes As DialogResult = clrdia.ShowDialog
                Dim clrStr As String = ""

                If clrRes = DialogResult.OK Then
                    oldClr = clrdia.Color
                End If

                ed.WriteMessage(vbLf & "Pick new color...")

                clrdia = New Autodesk.AutoCAD.Windows.ColorDialog
                clrRes = clrdia.ShowDialog

                If clrRes = DialogResult.OK Then
                    newClr = clrdia.Color
                End If

            Catch
                ed.WriteMessage("Bad color.  Command cancelled")
                Exit Sub
            End Try


            If newClr IsNot Nothing And oldClr IsNot Nothing Then

                Dim changeDic As New Dictionary(Of String, Long)

                For i As Integer = 0 To fi.Length - 1

                    'Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction

                    'create a new autocad drawing database to be used out-of-process
                    Dim sourceDb As New Database(False, True)

                    'read dwg file into database
                    Try
                        sourceDb.ReadDwgFile(fi(i).FullName, System.IO.FileShare.Read, False, "")
                        sourceDb.CloseInput(True)
                    Catch __unusedException1__ As System.Exception
                        MessageBox.Show(vbLf & "Unable to read drawing file.")
                        Return
                    End Try

                    'Dim tempDoc As Document = DocumentCollectionExtension.Add(docMgr, fi(i).FullName)

                    Dim j As Long = 0

                    Using actrans2 As Transaction = sourceDb.TransactionManager.StartTransaction

                        Dim blktbl As BlockTable = actrans2.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
                        For Each btrId As ObjectId In blktbl
                            Dim btr As BlockTableRecord = actrans2.GetObject(btrId, OpenMode.ForWrite)
                            For Each entId As ObjectId In btr
                                Dim ent As Entity = actrans2.GetObject(entId, OpenMode.ForRead)
                                If ent.Color = oldClr Then
                                    ent.UpgradeOpen()
                                    ent.Color = newClr
                                    j += 1
                                End If
                            Next
                        Next

                        Try
                            If j > 0 Then sourceDb.SaveAs(fi(i).FullName, False, Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1027, sourceDb.SecurityParameters)
                        Catch ex As Exception
                            MessageBox.Show("Unable to save drawing file." & vbLf & ex.Message)
                        End Try

                        changeDic.Add(fi(i).Name, j)
                        actrans2.Commit()
                    End Using

                    'acTrans.Commit()
                    'End Using

                Next

                For Each k As String In changeDic.Keys
                    ed.WriteMessage(vbLf & k.ToString & vbTab & changeDic(k).ToString & " Entities changed color")
                Next

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
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select blocks to change entity plotstyles to ByLayer:")}
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

            Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim filepath As String = ""
            Dim blkFolder As String
            Dim folderpicker As New FolderBrowserDialog
            With folderpicker
                .Description = "Select folder containing drawing files to change."
                .ShowNewFolderButton = False
                If Not String.IsNullOrEmpty(m_fldr) Then .SelectedPath = m_fldr
            End With

            If folderpicker.ShowDialog = DialogResult.OK Then
                filepath = folderpicker.SelectedPath
                blkFolder = folderpicker.SelectedPath
                m_fldr = blkFolder
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
            Dim ed As Editor = curDwg.Editor
            'Dim ed As Editor = curDwg.Editor

            ed.WriteMessage(vbLf & "This command will create thumbnail images for selected drawing files in a picked folder.")
            Dim uR As Boolean = YesNoQuery(vbLf & "Do you want to proceed?")
            If Not uR Then Exit Sub

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
                .Text = "DWG Picker"
            End With

            Dim fileList As Collection

            pckr.ShowDialog()

            If pckr.DialogResult = DialogResult.OK Then
                fileList = pckr.PickCol
            Else
                pckr.Dispose()
                Exit Sub
            End If

            Dim fileCount As Integer = 0

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

                    Dim dViewRatio As Double
                    Dim imHt As Double
                    Dim imwidth As Double

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
                        dViewRatio = (acView.Width / acView.Height)
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
                        If dHeight >= dWidth / dViewRatio Then
                            dWidth = dHeight * dViewRatio
                            imHt = 600
                            imwidth = 600 * dViewRatio
                        Else
                            dHeight = dWidth / dViewRatio
                            imwidth = 800
                            imHt = 800 / dViewRatio
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

                    Dim bm As Drawing.Bitmap = tempDoc.CapturePreviewImage(imwidth, imHt)
                    tempDb.ThumbnailBitmap = bm

                    tempDb.SaveAs(pathList(ky), True, DwgVersion.AC1027, Nothing)

                    'Dim ocmd As Object = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CMDECHO")
                    'tempDoc.SendStringToExecute("(setvar ""CMDECHO"" 0)" & "(command ""_.SAVEAS"" """" """ & lspName & """)" & "(setvar ""CMDECHO"" " + ocmd.ToString() & ")" & "(princ) ", False, False, False)
                    'tempDoc.SendStringToExecute("(command ""_.SAVEAS"" """" """ & lspName & """)" & "(princ) ", False, False, True)
                    'tempDoc.SendStringToExecute("(command ""_.SAVE "" & "")" & "(princ) ", False, False, True)
                    'tempDoc.SendStringToExecute("(SETVAR \" & ChrW(34) & "CMDECHO\" & ChrW(34) & "0) ") & "(command \" & chrw(34) & "_.SAVEAS\" & chrw(34) & " \" & chrw(34) & "\ " & chrw(34) 
                End Using

                fileCount += 1
                tempDoc.CloseAndDiscard
            Next

            ed.WriteMessage(vbLf & fileCount & " files processed.")

        End Sub

        <CommandMethod("AUDITSF")>
        Public Sub AuditSelectFiles()
            Dim acDwgMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim curDwg As Document = acDwgMgr.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            ed.WriteMessage(vbLf & "This command will audit all selected drawing files.")
            Dim uR As Boolean = YesNoQuery(vbLf & "Do you want to proceed?")
            If Not uR Then Exit Sub

            'Dim blkFldr As String = "//EESServer/datadisk/cad/blocks/road/reg colored/design"

            Dim myFiles() As String = GetMyFileNames("Dwg FIles (*.dwg)|*.DWG|", "Select Drawings to purge layers")
            Dim pathList As New Dictionary(Of String, String)

            If myFiles IsNot Nothing AndAlso myFiles.Length > 0 Then
                For Each fName As String In myFiles
                    pathList.Add(Path.GetFileNameWithoutExtension(fName), fName)
                Next
            Else
                Exit Sub
            End If

            'Dim blkFldr As String = GetMyFolderName()
            'Dim blkfldr As String = "\\EESServer\datadisk\CAD\BLOCKS\ROAD\PVMT\Design"
            'If String.IsNullOrEmpty(blkFldr) Then Exit Sub
            'If Not Directory.Exists(blkFldr) Then Exit Sub
            'Dim files() As FileInfo = AllDwgFilesInFolder(blkFldr, False)

            Try

                For Each ky As String In pathList.Keys
                    Dim fi As New FileInfo(pathList(ky))
                    Dim acDB As New Database(False, True)
                    Dim thisFile As String = pathList(ky)
                    ed.WriteMessage(vbLf & ky)
                    'read dwg file into database
                    Try
                        acDB.ReadDwgFile(thisFile, System.IO.FileShare.Read, False, "")
                        acDB.CloseInput(True)
                    Catch __unusedException1__ As System.Exception
                        ed.WriteMessage(vbLf & "Unable to read drawing file.")
                        Exit Sub
                    End Try

                    Audit(acDB, True, True)
                    acDB.SaveAs(thisFile, True, Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1027, acDB.SecurityParameters)
                Next

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        <CommandMethod("FNESTBKJ")>
        Public Sub FindNestedRef()
            Dim acDwgMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim curDwg As Document = acDwgMgr.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim findName As String
            Dim pso As New PromptStringOptions(vbLf & "Enter the block name for search")

            With pso
                .AllowSpaces = True
            End With

            Dim psr As PromptResult = ed.GetString(pso)
            If psr.Status = PromptStatus.OK Then
                Debug.Print(psr.Status)
                findName = psr.StringResult
            Else
                Exit Sub
            End If

            If String.IsNullOrEmpty(findName) Then Exit Sub
            Dim blkDic As New Dictionary(Of String, Integer)
            Dim j As Integer

            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim blkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                'Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                For Each obID As ObjectId In blkTbl
                    'Dim dbObj As DBObject = acTrans.GetObject(obID, OpenMode.ForRead)
                    Dim curBTR As BlockTableRecord = acTrans.GetObject(obID, OpenMode.ForRead)
                    j = 0
                    For Each obid2 As ObjectId In curBTR
                        Dim dbobj As DBObject = acTrans.GetObject(obid2, OpenMode.ForRead)
                        If TypeOf dbobj Is BlockReference Then
                            Dim Bref As BlockReference = CType(dbobj, BlockReference)
                            Dim testName As String = Bref.BlockName
                            If findName = testName Then
                                j += 1
                            End If
                        End If
                    Next
                    If j > 0 Then
                        blkDic(curBTR.Name) = j
                    End If
                Next
            End Using

            If blkDic.Keys.Count > 0 Then
                For Each ky As String In blkDic.Keys
                    ed.WriteMessage(vbLf & "Block: " & ky & "  Count: " & blkDic(ky).ToString)
                Next
            Else
                ed.WriteMessage(vbLf & "No nested references found.")
            End If


        End Sub

        <CommandMethod("FBLKTXT")>
        Public Sub FindBlocksWithStyle()
            Dim acDwgMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim curDwg As Document = acDwgMgr.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            'Dim pso As New PromptStringOptions(vbLf & "Enter the style for the search")

            'With pso
            '    .AllowSpaces = True
            'End With

            Dim stName As String

            Using sp As New FontPicker(True)
                sp.ShowDialog()
                If sp.DialogResult = DialogResult.OK Then
                    stName = sp.StyleName
                Else
                    stName = ""
                End If

            End Using

            'Dim psr As PromptResult = ed.GetString(pso)
            'If psr.Status = PromptStatus.OK Then
            '    Debug.Print(psr.Status)
            '    stName = psr.StringResult
            'Else
            '    Exit Sub
            'End If

            'Debug.Print(stName)

            If String.IsNullOrEmpty(stName) Then Exit Sub
            Dim hits As New Dictionary(Of String, Integer)

            Dim entCount As Integer
            Dim bName As String

            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim blkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                Dim tst As TextStyleTable = acTrans.GetObject(dwgDB.TextStyleTableId, OpenMode.ForRead)
                Dim oldTStyleId As ObjectId = tst(stName)
                'Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                For Each obID As ObjectId In blkTbl
                    'Dim dbObj As DBObject = acTrans.GetObject(obID, OpenMode.ForRead)
                    Dim curBTR As BlockTableRecord = acTrans.GetObject(obID, OpenMode.ForRead)
                    bName = curBTR.Name
                    entCount = 0
                    For Each entID As ObjectId In curBTR
                        Dim myEnt As DBObject = acTrans.GetObject(entID, OpenMode.ForRead)
                        If TypeOf myEnt Is DBText Then
                            Using myText As DBText = TryCast(myEnt, DBText)
                                If myText IsNot Nothing Then
                                    If myText.TextStyleId = oldTStyleId Then
                                        entCount += 1
                                    End If
                                End If
                            End Using
                        ElseIf TypeOf myEnt Is MText Then
                            Using myMtext As MText = TryCast(myEnt, MText)
                                If myMtext IsNot Nothing Then
                                    If myMtext.TextStyleId = oldTStyleId Then
                                        entCount += 1
                                    End If
                                End If
                            End Using
                        ElseIf TypeOf myEnt Is AttributeReference Then
                            Using myAtt As AttributeReference = TryCast(myEnt, AttributeReference)
                                If myAtt IsNot Nothing Then
                                    If myAtt.TextStyleId = oldTStyleId Then
                                        entCount += 1
                                    End If
                                End If
                            End Using
                        ElseIf TypeOf myEnt Is Dimension Then
                            Using myDim As Dimension = TryCast(myEnt, Dimension)
                                If myDim IsNot Nothing Then
                                    If myDim.TextStyleId = oldTStyleId Then
                                        entCount += 1
                                    End If
                                End If
                            End Using
                        ElseIf TypeOf myEnt Is MLeader Then
                            Using myDim As MLeader = TryCast(myEnt, MLeader)
                                If myDim IsNot Nothing Then
                                    If myDim.TextStyleId = oldTStyleId Then
                                        entCount += 1
                                    End If
                                End If
                            End Using
                        End If
                    Next
                    Dim dsT As DimStyleTable = acTrans.GetObject(dwgDB.DimStyleTableId, OpenMode.ForRead)
                    For Each objId As ObjectId In dsT
                        Using dRec As DimStyleTableRecord = TryCast(acTrans.GetObject(objId, OpenMode.ForRead), DimStyleTableRecord)
                            If dRec IsNot Nothing Then
                                If dRec.Dimtxsty = oldTStyleId Then
                                    entCount += 1
                                End If
                            End If
                        End Using
                    Next
                    hits(bName) = entCount

                Next

                'Dim myEnt As DBObject = acTrans.GetObject(entID, OpenMode.ForRead)
                '        Dim myStyleName As String
                '        If TypeOf myEnt Is DBText Then
                '            Dim myText As DBText = CType(myEnt, DBText)
                '            myStyleName = myText.TextStyleName
                '            If myStyleName = stName Then
                '                entCount += 1
                '            End If
                '        ElseIf TypeOf myEnt Is MText Then
                '            Dim myMtext As MText = CType(myEnt, MText)
                '            If myMtext.TextStyleName = stName Then
                '                entCount += 1
                '            End If
                '        ElseIf TypeOf myEnt Is AttributeReference Then
                '            Dim myMtext As AttributeReference = CType(myEnt, AttributeReference)
                '            If myMtext.TextStyleName = stName Then
                '                entCount += 1
                '            End If
                '        ElseIf TypeOf myEnt Is MLeader Then
                '            Dim myLead As MLeader = CType(myEnt, MLeader)
                '            If myLead.TextStyleId = tst(stName) Then
                '                entCount += 1
                '            End If
                '        End If
                '    Next
                '    hits(bName) = entCount
                'Next
                acTrans.Commit()
            End Using

            Dim newHits As Boolean = False

            If hits.Count > 0 Then
                For Each ky As String In hits.Keys
                    If hits(ky) > 0 Then
                        ed.WriteMessage(vbLf & "block: " & ky & ", hits: " & hits(ky).ToString)
                        newHits = True
                    End If
                Next
            End If

            If Not newHits Then ed.WriteMessage(vbLf & "No entities with textstyle " & stName & " found.")

        End Sub


        <CommandMethod("PATF")>
        Public Sub PurgeAllTexstylesInFolder()
            Dim acDwgMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim curDwg As Document = acDwgMgr.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            ed.WriteMessage(vbLf & "This command will purge all unused non-standard textstyles from all selected dwg files.")
            Dim uR As Boolean = YesNoQuery(vbLf & "Do you want to proceed?")

            If Not uR Then Exit Sub

            'Dim blkFldr As String = "//EESServer/datadisk/cad/blocks/road/reg colored/design/temptest"

            'Dim blkFldr As String = GetMyFolderName()
            'Dim blkfldr As String = "\\EESServer\datadisk\CAD\BLOCKS\ROAD\PVMT\Design"
            'If String.IsNullOrEmpty(blkFldr) Then Exit Sub
            'If Not Directory.Exists(blkFldr) Then Exit Sub

            'Dim files() As FileInfo = AllDwgFilesInFolder(blkFldr, False)

            Dim myFiles() As String = GetMyFileNames("Dwg FIles (*.dwg)|*.DWG|", "Select Drawings to purge layers")
            Dim pathList As New Dictionary(Of String, String)

            If myFiles IsNot Nothing And myFiles.Length > 0 Then
                For Each fName As String In myFiles
                    pathList.Add(Path.GetFileNameWithoutExtension(fName), fName)
                Next
            Else
                Exit Sub
            End If

            'For Each fi As FileInfo In files
            '    Dim dwgName As String = Path.GetFileNameWithoutExtension(fi.Name)
            '    If Not dwgName.ToUpper.Contains("MASTER") Or Not dwgName.ToUpper.Contains("DETAIL") Then
            '        pathList.Add(dwgName, fi.FullName)
            '    Else
            '        'ed.WriteMessage(vbLf & "Dwg " & dwgName & " Not processed.")
            '        'Dim cont As Boolean = YesNoQuery(vbLf & "Continue processing?")
            '        'If Not cont Then Exit Sub
            '    End If
            'Next

            For Each ky As String In pathList.Keys
                Try
                    'Dim acDoc As Document = acDwgMgr.Open(pathList(ky), False)
                    'Dim acdb As Database = acDoc.Database
                    ed.WriteMessage(vbLf & "Processing: " & ky)

                    Using acDB As New Database(False, True)

                        'read dwg file into database
                        Try
                            acDB.ReadDwgFile(pathList(ky), System.IO.FileShare.Read, False, "")
                            acDB.CloseInput(True)
                        Catch __unusedException1__ As System.Exception
                            ed.WriteMessage(vbLf & "Unable to read drawing file.")
                            Exit Sub
                        End Try

                        'Dim acDoc As Document = acDwgMgr.Open(pathList(ky), False)
                        Using actrans As Transaction = acDB.TransactionManager.StartTransaction
                            'Using docLock As DocumentLock = acDoc.LockDocument
                            'Dim acdb As Database = acDoc.Database
                            Dim tst As TextStyleTable = actrans.GetObject(acDB.TextStyleTableId, OpenMode.ForWrite)

                            'make the current textstyle Standard
                            Dim cStyleID As ObjectId = acDB.Textstyle
                            Dim cStyle As TextStyleTableRecord = actrans.GetObject(cStyleID, OpenMode.ForRead)

                            If Not cStyle.Name = "Standard" Then
                                Dim tempID As ObjectId = tst("Standard")
                                If Not tempID.IsNull Then acDB.Textstyle = tempID
                            End If

                            Dim prgList As New ObjectIdCollection
                            Dim transDic As Dictionary(Of String, String) = RGtranslateDic()

                            'cycle through the textstyletable and get each textstyle
                            For Each stId As ObjectId In tst
                                Dim myStyle As TextStyleTableRecord = actrans.GetObject(stId, OpenMode.ForWrite)
                                Dim stlName As String = myStyle.Name

                                'if it is not Standard, then collect the text objects that use this style
                                If Not stlName = "Standard" Then
                                    Dim textObjs As ObjectIdCollection = GetDBTextWithStyle(stlName, acDB)

                                    'if the style has no text referencing it, add it to the purge list
                                    If textObjs IsNot Nothing Then
                                        If textObjs.Count = 0 Then
                                            prgList.Add(stId)
                                        End If
                                    Else
                                        Continue For
                                    End If
                                End If
                            Next

                            'if there are styles to purge from this drawing, purge them
                            Try
                                If prgList.Count > 0 Then acDB.Purge(prgList)
                            Catch ex As Exception
                                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Error:" & vbLf & ex.Message)
                                actrans.Abort()
                                Continue For
                                'acDoc.CloseAndDiscard
                            End Try
                            'ed.WriteMessage(vbLf & "style " & stlName & " has been purged from " & ky)
                            'End Using
                            'acDoc.CloseAndSave(pathList(ky))
                            actrans.Commit()
                            acDB.SaveAs(pathList(ky), True, Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1027, acDB.SecurityParameters)
                        End Using
                    End Using

                Catch ex As Exception
                    ed.WriteMessage(ex.Message)
                    Continue For
                End Try
            Next

        End Sub

        <CommandMethod("CHTEXTST")>
        Public Sub ChTextStyles()
            Dim acDwgMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim curDwg As Document = acDwgMgr.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            ed.WriteMessage(vbLf & "This command will change the textstyle for all blocks and entities in current drawing.")
            Dim uR As Boolean = YesNoQuery(vbLf & "Do you want to proceed?")
            If Not uR Then Exit Sub

            ed.WriteMessage(vbLf & "Pick textstyle to be changed: ")
            Dim oldStName As String

            Using fp As New FontPicker(True)
                fp.Text = "Textstyles"
                fp.LabPickerType.Text = "Pick a textstyle to change:"
                fp.ShowDialog()

                If fp.DialogResult = DialogResult.OK Then
                    oldStName = fp.StyleName
                Else
                    Exit Sub
                End If
            End Using

            ed.WriteMessage(vbLf & "Pick a new textstyle: ")
            Dim newStName As String

            Using fp2 As New FontPicker(True)
                fp2.Text = "Textstyles"
                fp2.LabPickerType.Text = "Pick a new textstyle:"
                fp2.ShowDialog()

                If fp2.DialogResult = DialogResult.OK Then
                    newStName = fp2.StyleName
                Else
                    Exit Sub
                End If
            End Using

            If oldStName = "" Or newStName = "" Then Exit Sub

            Dim dtCount As Integer
            Dim mtCount As Integer
            Dim attCount As Integer
            Dim dimCount As Integer
            Dim leadCount As Integer
            Dim dimStyCount As Integer

            Try
                Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                    Dim blkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                    Dim tst As TextStyleTable = acTrans.GetObject(dwgDB.TextStyleTableId, OpenMode.ForRead)
                    Dim oldtStyleId As ObjectId = tst(oldStName)
                    Dim newtStyleId As ObjectId = tst(newStName)
                    For Each bTRid As ObjectId In blkTbl
                        Dim bTR As BlockTableRecord = acTrans.GetObject(bTRid, OpenMode.ForRead)
                        'If Not bTR.IsLayout Then
                        For Each entID As ObjectId In bTR
                            Dim myEnt As DBObject = acTrans.GetObject(entID, OpenMode.ForRead)
                            If TypeOf myEnt Is DBText Then
                                Using myText As DBText = TryCast(myEnt, DBText)
                                    If myText IsNot Nothing Then
                                        If myText.TextStyleId = oldtStyleId Then
                                            If Not myText.IsWriteEnabled Then myText.UpgradeOpen()
                                            myText.TextStyleId = newtStyleId
                                            dtCount += 1
                                        End If
                                    End If
                                End Using
                            ElseIf TypeOf myEnt Is MText Then
                                Using myMtext As MText = TryCast(myEnt, MText)
                                    If myMtext IsNot Nothing Then
                                        If myMtext.TextStyleId = oldtStyleId Then
                                            If Not myMtext.IsWriteEnabled Then myMtext.UpgradeOpen()
                                            myMtext.TextStyleId = newtStyleId
                                            mtCount += 1
                                        End If
                                    End If
                                End Using
                            ElseIf TypeOf myEnt Is AttributeReference Then
                                Using myAtt As AttributeReference = TryCast(myEnt, AttributeReference)
                                    If myAtt IsNot Nothing Then
                                        If myAtt.TextStyleId = oldtStyleId Then
                                            If Not myAtt.IsWriteEnabled Then myAtt.UpgradeOpen()
                                            myAtt.TextStyleId = newtStyleId
                                            attCount += 1
                                        End If
                                    End If
                                End Using
                            ElseIf TypeOf myEnt Is Dimension Then
                                Using myDim As Dimension = TryCast(myEnt, Dimension)
                                    If myDim IsNot Nothing Then
                                        If myDim.TextStyleId = oldtStyleId Then
                                            If Not myDim.IsWriteEnabled Then myDim.UpgradeOpen()
                                            myDim.TextStyleId = newtStyleId
                                            dimCount += 1
                                        End If
                                    End If
                                End Using
                            ElseIf TypeOf myEnt Is MLeader Then
                                Using myDim As MLeader = TryCast(myEnt, MLeader)
                                    If myDim IsNot Nothing Then
                                        If myDim.TextStyleId = oldtStyleId Then
                                            If Not myDim.IsWriteEnabled Then myDim.UpgradeOpen()
                                            myDim.TextStyleId = newtStyleId
                                            leadCount += 1
                                        End If
                                    End If
                                End Using
                            End If
                        Next
                        'End If
                    Next

                    Dim dsT As DimStyleTable = acTrans.GetObject(dwgDB.DimStyleTableId, OpenMode.ForRead)
                    For Each obId As ObjectId In dsT
                        Using dRec As DimStyleTableRecord = TryCast(acTrans.GetObject(obId, OpenMode.ForRead), DimStyleTableRecord)
                            If dRec IsNot Nothing Then
                                If dRec.Dimtxsty = oldtStyleId Then
                                    If Not dRec.IsWriteEnabled Then dRec.UpgradeOpen()
                                    dRec.Dimtxsty = newtStyleId
                                    dimStyCount += 1
                                End If
                            End If
                        End Using
                    Next

                    ed.WriteMessage(vbLf & "Changed Objects: ")
                    ed.WriteMessage(vbLf & "Text entities: " & dtCount)
                    ed.WriteMessage(vbLf & "MText entities: " & mtCount)
                    ed.WriteMessage(vbLf & "Attribute References: " & attCount)
                    ed.WriteMessage(vbLf & "Dimension entities: " & dimCount)
                    ed.WriteMessage(vbLf & "MLeader entities: " & leadCount)
                    ed.WriteMessage(vbLf & "Dimension Styles: " & dimStyCount)

                    acTrans.Commit()
                End Using

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub


        <CommandMethod("URGS")>
        Public Sub UpdateRoadgeekStyles()
            Dim acDwgMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim curDwg As Document = acDwgMgr.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor
            ed.WriteMessage(vbLf & "This command will update all Roadgeek 2000 fonts to Roadgeek 2005 fonts for all selected dwgs in a directory.")
            Dim uR As Boolean = YesNoQuery(vbLf & "Do you want to proceed?")
            If Not uR Then Exit Sub

            'Dim blkFldr As String = "//EESServer/datadisk/cad/blocks/road/reg colored/design/temptest"

            Dim myFiles() As String = GetMyFileNames("Dwg FIles (*.dwg)|*.DWG|", vbLf & "Select Drawings to purge layers")

            'Dim blkFldr As String = GetMyFolderName()
            'Dim blkfldr As String = "\\EESServer\datadisk\CAD\BLOCKS\ROAD\PVMT\Design"
            'If String.IsNullOrEmpty(blkFldr) Then Exit Sub
            'If Not Directory.Exists(blkFldr) Then Exit Sub

            'Dim files() As FileInfo = AllDwgFilesInFolder(blkFldr, False)

            Dim pathList As New Dictionary(Of String, String)

            If myFiles IsNot Nothing And myFiles.Length > 0 Then
                For Each fName As String In myFiles
                    pathList.Add(Path.GetFileNameWithoutExtension(fName), fName)
                Next
            Else
                Exit Sub
            End If

            'For Each fi As FileInfo In files
            '    Dim dwgName As String = Path.GetFileNameWithoutExtension(fi.Name)
            '    If Not dwgName.ToUpper.Contains("MASTER") Or Not dwgName.ToUpper.Contains("DETAIL") Then
            '        pathList.Add(dwgName, fi.FullName)
            '    Else
            '        'ed.WriteMessage(vbLf & "Dwg " & dwgName & " Not processed.")
            '        'Dim cont As Boolean = YesNoQuery(vbLf & "Continue processing?")
            '        'If Not cont Then Exit Sub
            '    End If
            'Next

            For Each ky As String In pathList.Keys
                Try
                    'Dim acDoc As Document = acDwgMgr.Open(pathList(ky), False)
                    'Dim acdb As Database = acDoc.Database
                    ed.WriteMessage(vbLf & "Processing: " & ky)

                    Using acDB As New Database(False, True)

                        'read dwg file into database
                        Try
                            acDB.ReadDwgFile(pathList(ky), System.IO.FileShare.Read, False, "")
                            acDB.CloseInput(True)
                        Catch __unusedException1__ As System.Exception
                            ed.WriteMessage(vbLf & "Unable to read drawing file.")
                            Exit Sub
                        End Try

                        'Dim acDoc As Document = acDwgMgr.Open(pathList(ky), False)
                        Using actrans As Transaction = acDB.TransactionManager.StartTransaction

                            Dim tst As TextStyleTable = actrans.GetObject(acDB.TextStyleTableId, OpenMode.ForWrite)
                            Dim rgFiles As Dictionary(Of String, String) = GetRGFilesDic()
                            Dim rgTrans As Dictionary(Of String, String) = RGtranslateDic()

                            For Each stId As ObjectId In tst
                                Dim myStyle As TextStyleTableRecord = actrans.GetObject(stId, OpenMode.ForWrite)
                                Dim stlName As String = myStyle.Name
                                Dim myFnt As String = myStyle.FileName
                                Dim renamePending As Boolean = False
                                Dim nameExists As Boolean = False
                                Dim fontcode As String = ""

                                If myFnt.ToUpper.Contains("ROADGEEK") Then
                                    If Not Left(stlName, 6).ToUpper = "SERIES" Then
                                        Dim charPos As Integer = myFnt.IndexOf("Series")
                                        fontcode = myFnt.Substring(charPos)
                                        If rgTrans.Keys.Contains(fontcode) Then
                                            If Not tst.Has(rgTrans(fontcode)) Then
                                                Dim renStyle As Boolean = YesNoQuery(vbLf & "Do you want to rename style " & stlName & " to " & rgTrans(fontcode) & "?")
                                                If renStyle Then renamePending = True
                                            Else
                                                nameExists = True
                                            End If
                                        End If
                                    End If

                                    Dim fontUpdated As Boolean = False

                                    Select Case myFnt
                                        Case Is = rgFiles("RG_2000B")
                                            myStyle.FileName = rgFiles("RG_2005B")
                                            fontUpdated = True
                                        Case Is = rgFiles("RG_2000C")
                                            myStyle.FileName = rgFiles("RG_2005C")
                                            fontUpdated = True
                                        Case Is = rgFiles("RG_2000D")
                                            myStyle.FileName = rgFiles("RG_2005D")
                                            fontUpdated = True
                                        Case Is = rgFiles("RG_2000E")
                                            myStyle.FileName = rgFiles("RG_2005E")
                                            fontUpdated = True
                                        Case Is = rgFiles("RG_2000F")
                                            myStyle.FileName = rgFiles("RG_2005F")
                                            fontUpdated = True
                                        Case Else
                                            fontUpdated = False
                                    End Select

                                    If renamePending Then
                                        myStyle.Name = rgTrans(fontcode)
                                        If fontUpdated Then
                                            ed.WriteMessage(vbLf & "Style named " & stlName & " in drawing " & ky & " has been renamed to " & rgTrans(fontcode) & " and the font file updated to Roadgeek 2005")
                                        Else
                                            ed.WriteMessage(vbLf & "Style named " & stlName & " in drawing " & ky & " has been renamed to " & rgTrans(fontcode) & ".")
                                        End If
                                    Else
                                        If nameExists Then ed.WriteMessage(vbLf & stlName & " in drawing " & ky & " could not be renamed to " & rgTrans(fontcode) & " because style exists.")
                                        If fontUpdated Then
                                            ed.WriteMessage(vbLf & "Font file for style named " & stlName & " in drawing " & ky & " has been updated To Roadgeek 2005")
                                        End If
                                    End If
                                End If
                            Next

                            actrans.Commit()
                            acDB.SaveAs(pathList(ky), True, Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1027, acDB.SecurityParameters)
                        End Using
                    End Using

                Catch ex As Exception
                    ed.WriteMessage(ex.Message)
                    Continue For
                End Try

            Next

        End Sub

        <CommandMethod("PALF")>
        Public Sub PurgeAllLayersFromFIles()
            Dim acDwgMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim curDwg As Document = acDwgMgr.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            ed.WriteMessage(vbLf & "This command will purge unused layers from all selected drawing files In a folder.")
            Dim uR As Boolean = YesNoQuery(vbLf & "Do you want To proceed?")
            If Not uR Then Exit Sub

            Dim myFiles() As String = GetMyFileNames("Dwg FIles (*.dwg)|*.DWG|", "Select Drawings To purge layers")

            'Dim blkFldr As String = GetMyFolderName()
            'Dim blkfldr As String = "\\EESServer\datadisk\CAD\BLOCKS\ROAD\PVMT\Design"
            'If String.IsNullOrEmpty(blkFldr) Then Exit Sub
            'If Not Directory.Exists(blkFldr) Then Exit Sub

            'Dim files() As FileInfo = AllDwgFilesInFolder(blkFldr, False)

            Dim pathList As New Dictionary(Of String, String)

            If myFiles IsNot Nothing And myFiles.Length > 0 Then
                For Each fName As String In myFiles
                    pathList.Add(Path.GetFileNameWithoutExtension(fName), fName)
                Next
            Else
                Exit Sub
            End If

            'For Each fi As FileInfo In files
            '    If fi.Extension = ".dwg" Then
            '        Dim dwgName As String = Path.GetFileNameWithoutExtension(fi.Name)
            '        If Not dwgName.ToUpper.Contains("MASTER") Or Not dwgName.ToUpper.Contains("DETAIL") Then
            '            pathList.Add(dwgName, fi.FullName)
            '        Else
            '            ed.WriteMessage(vbLf & "Dwg " & dwgName & " Not processed.")
            '            Dim cont As Boolean = YesNoQuery(vbLf & "Continue processing?")
            '            If Not cont Then
            '                Exit Sub
            '            Else
            '                Continue For
            '            End If
            '        End If
            '    End If
            'Next

            For Each ky As String In pathList.Keys
                Try
                    Using acDB As New Database(False, True)

                        'read dwg file into database
                        Try
                            acDB.ReadDwgFile(pathList(ky), System.IO.FileShare.Read, False, "")
                            acDB.CloseInput(True)
                            Debug.Print(vbLf & ky)
                        Catch __unusedException1__ As System.Exception
                            ed.WriteMessage(vbLf & "Unable To read drawing file: " & ky)
                            Continue For
                        End Try

                        'Dim acDoc As Document = acDwgMgr.Open(pathList(ky), False)
                        Using actrans As Transaction = acDB.TransactionManager.StartTransaction
                            'Using docLock As DocumentLock = acDoc.LockDocument
                            'DocumentCollectionExtension.Open(acDwgMgr, pathList(ky), False)
                            'Using acDoc As Document = acDwgMgr.GetDocument(acDB)
                            Dim lTbl As LayerTable = actrans.GetObject(acDB.LayerTableId, OpenMode.ForRead)
                            Dim cLayID As ObjectId = acDB.Clayer
                            If Not cLayID = acDB.LayerZero Then acDB.Clayer = acDB.LayerZero
                            'Dim cLayr As LayerTableRecord = actrans.GetObject(cLayID, OpenMode.ForRead)
                            'Dim tempID As ObjectId = acDB.LayerZero
                            'If Not tempID.IsNull Then acDB.Clayer = tempID
                            'End If

                            Dim prgList As New ObjectIdCollection
                            For Each layId As ObjectId In lTbl
                                If Not layId = dwgDB.LayerZero Then
                                    Dim myLyr As LayerTableRecord = actrans.GetObject(layId, OpenMode.ForRead)
                                    Dim lyrObIds As ObjectIdCollection = GetEntitiesOnLayer(myLyr.Name)
                                    If lyrObIds Is Nothing OrElse lyrObIds.Count = 0 Then
                                        prgList.Add(layId)
                                    End If
                                    'If Not ObjectsExistOnLayer(myLyr.Name, acDB) Then
                                    '        prgList.Add(layId)
                                    '    End If
                                End If
                            Next

                            If prgList.Count > 0 Then
                                Try
                                    acDB.Purge(prgList)
                                Catch ex As Exception
                                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Error:" & vbLf & ex.Message)
                                End Try

                                'End Using
                                actrans.Commit()
                                acDB.SaveAs(pathList(ky), False, Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1027, acDB.SecurityParameters)
                                ed.WriteMessage(vbLf & pathList(ky) & " purged and saved.")

                                'acDoc.CloseAndSave(pathList(ky))
                                'Else
                                '    actrans.Abort()
                                '    'acDoc.CloseAndDiscard
                            Else
                                ed.WriteMessage(vbLf & pathList(ky) & " not purged or saved.")
                            End If


                        End Using

                    End Using

                    'acDoc.CloseAndSave(pathList(ky))
                Catch ex As Exception
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Error:" & vbLf & ex.Message)
                    MessageBox.Show(ex.Message)
                    Continue For
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
                Try
                    Dim blktbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                    For Each bID As ObjectId In blktbl
                        Dim btr As BlockTableRecord = acTrans.GetObject(bID, OpenMode.ForRead)
                        'Dim fname1 As String = Path
                        Dim fName As String
                        Try
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
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                            Continue For
                        End Try
                    Next

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                acTrans.Commit()
            End Using

            ed.WriteMessage(vbLf & "All blocks saved to designated folder")

        End Sub

        <CommandMethod("CBU")>
        Public Sub ChangeBlockUnits()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            'Dim sName As String = curDwg.Name
            'Dim savef As String = Path.GetDirectoryName(sName) '& "\testFldr\"

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

            Try

                Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                    Dim blktbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                    Dim bList As New SortedDictionary(Of String, ObjectId)

                    For Each bID As ObjectId In blktbl
                        Dim btr As BlockTableRecord = TryCast(acTrans.GetObject(bID, OpenMode.ForRead), BlockTableRecord)
                        If btr IsNot Nothing Then
                            Try
                                If Not btr.IsLayout Then
                                    Dim bName As String = btr.Name
                                    Dim btrID As ObjectId = blktbl(bName)
                                    bList.Add(bName, btrID)
                                End If
                            Catch ex As Exception
                                ed.WriteMessage(vbLf & "Error in block " & btr.Name)
                                Continue For
                            End Try
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
                        Try
                            Using btr As BlockTableRecord = acTrans.GetObject(selBlks(blkStr), OpenMode.ForWrite)
                                btr.Units = uv
                            End Using
                        Catch ex As Exception
                            ed.WriteMessage(vbLf & "Error setting units for block " & blkStr)
                            Continue For
                        End Try
                    Next
                    acTrans.Commit()
                End Using

            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error. " & ex.Message)
                Exit Sub
            End Try

            ed.WriteMessage(vbLf & "All blocks updated")

        End Sub

        <CommandMethod("CPMLT", CommandFlags.UsePickSet)>
        Public Sub ChangePvmntMarkingsLineTypes()
            'by David Eisenbeisz (c)2023

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction()
                Dim blkTbl As BlockTable = actrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)

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
                                        If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                            Dim pL As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(ent, Autodesk.AutoCAD.DatabaseServices.Polyline)
                                            If pL IsNot Nothing Then pL.Plinegen = True
                                        End If
                                    ElseIf ent.Linetype = "HIDDEN2" Then
                                        ent.UpgradeOpen()
                                        ent.Linetype = "PMREMOVE"
                                        ent.LinetypeScale = 1
                                        If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                            Dim pL As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(ent, Autodesk.AutoCAD.DatabaseServices.Polyline)
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
                        Dim lo As Layout = TryCast(acTrans.GetObject(loID, OpenMode.ForRead), Layout)

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
                                                        If prop.Value.ToString = "Signature" Then
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

        Friend m_area As Double
        Private m_handMaxX As Double
        Private m_handMinX As Double
        'Private m_handMinY As Double
        Private m_handMaxY As Double
        Private m_polarity As Boolean
        Private m_xtra As Double


        <CommandMethod("HANDLINES", CommandFlags.UsePickSet)>
        Public Sub HandLines()
            'by David Eisenbeisz

            Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = acDwg.Database
            Dim ed As Editor = acDwg.Editor
            Dim arcCol As New Collection

            'check to see if objects are already selected
            Dim SelResult As PromptSelectionResult = ed.SelectImplied()

            If SelResult.Status = PromptStatus.Error Then
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select objects to hand draw:")}
                With Seloptions
                    .AllowDuplicates = False
                    .RejectObjectsFromNonCurrentSpace = True
                    .MessageForRemoval = "Invalid object.  Object not selected."
                    .RejectObjectsOnLockedLayers = True

                End With
                SelResult = ed.GetSelection(Seloptions)
            Else
                ed.SetImpliedSelection(New ObjectId(-1) {})
            End If

            Dim hDistMin As Double
            Dim hDistMax As Double
            'Dim hAmpMin As Double
            Dim hAmpMax As Double

            'if saved parameters are present, ask if those should be used
            If m_handMinX <> 0 And m_handMaxX <> 0 And m_handMaxY <> 0 Then
                Dim upDateParams As Boolean = YesNoQuery(vbLf & "Update current wobble parameters?")
                If upDateParams = False Then
                    hDistMin = m_handMinX
                    hDistMax = m_handMaxX
                    'hAmpMin = m_handMinY
                    hAmpMax = m_handMaxY
                    GoTo ROUGHLINE
                End If
            End If

            'get the minimum frequency of the wobble
            Dim pdMinX As New PromptDoubleOptions(vbLf & "Enter minimum frequency of wobble in model space units")
            With pdMinX
                .AllowNegative = False
                .AllowZero = False
                If m_handMinX > 0 Then .DefaultValue = m_handMinX
            End With

            Dim pdrMinX As PromptDoubleResult = ed.GetDouble(pdMinX)
            If pdrMinX.Status = PromptStatus.OK Then
                hDistMin = pdrMinX.Value
                m_handMinX = hDistMin
            Else
                Exit Sub
            End If

            m_handMaxX = m_handMinX * 3
            hDistMax = m_handMaxX

            'enter wobble distance
            Dim pdoMaxY As New PromptDoubleOptions(vbLf & "Enter the distance in current space units for wobble deviation from the original line.")
            'Dim pdoMaxY As New PromptDoubleOptions(vbLf & "Enter the maximum amplitude of roughness in model space units")
            With pdoMaxY
                .AllowNegative = False
                .AllowZero = False
                If m_handMaxY > 0 Then .DefaultValue = m_handMaxY
            End With

            Dim pdrMaxY As PromptDoubleResult = ed.GetDouble(pdoMaxY)
            If pdrMaxY.Status = PromptStatus.OK Then
                hAmpMax = pdrMaxY.Value
                m_handMaxY = hAmpMax
            Else
                Exit Sub
            End If

ROUGHLINE:
            'Dim hdist As Double = (hDistMin + hDistMax) / 2
            'Dim hamp As Double = (hAmpMin + hAmpMax) / 2

            'if no objects are selected, then select objects
            Dim acSSet As SelectionSet
            Dim myObjIds() As ObjectId

            If SelResult.Status = PromptStatus.OK Then
                acSSet = SelResult.Value
                myObjIds = acSSet.GetObjectIds
            Else
                Exit Sub
            End If

            Dim isCircle As Boolean = False

            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction

                For Each obID As ObjectId In myObjIds
                    Dim acEnt As Entity = TryCast(acTrans.GetObject(obID, OpenMode.ForRead), Entity)
                    'If acEnt Is Nothing Then Continue For

                    Dim myPts As New Point3dCollection
                    Dim myLTId As ObjectId = acEnt.LinetypeId

                    'if object is a polyline
                    If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                        Dim myPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = CType(acEnt, Autodesk.AutoCAD.DatabaseServices.Polyline)

                        'if polyline is closed, treat it like a circle
                        If myPoly.Closed Then isCircle = True

                        myPts.Add(myPoly.StartPoint)
                        m_polarity = True
                        m_xtra = 0

                        'vertices - 2 = max index of segments
                        For i As Integer = 0 To myPoly.NumberOfVertices - 2
                            Dim blg As Double = myPoly.GetBulgeAt(i)
                            'if segment is a curve
                            If blg = 0 Then
                                Dim lineSeg As LineSegment2d = myPoly.GetLineSegment2dAt(i)
                                'If lineSeg.Length < hDistMin Then Continue For

                                If lineSeg.Length > hDistMax * 40000 Then
                                    MessageBox.Show("Wobble frequency is too short for this object.  Use a higher frequency or break object into smaller pieces.")
                                    Exit Sub
                                End If
                                Dim ls As Boolean
                                If i = myPoly.NumberOfVertices - 2 Then ls = True

                                Dim subPts As Point3dCollection = GetLineSegWobble(lineSeg, ls)
                                If subPts Is Nothing Then Exit Sub

                                For Each pt As Point3d In subPts
                                    myPts.Add(pt)
                                Next
                            Else
                                'segment not a curve
                                Dim arcSeg As CircularArc2d = myPoly.GetArcSegment2dAt(i)
                                Dim tSPt As Point3d = myPoly.GetPointAtParameter(i)
                                Dim arcReversed As Boolean
                                If New Point2d(tSPt.X, tSPt.Y) = arcSeg.StartPoint Then
                                    arcReversed = False
                                Else
                                    arcReversed = True
                                End If

                                Dim arcLen As Double = (arcSeg.EndAngle - arcSeg.StartAngle) * arcSeg.Radius
                                '(myPoly.GetParameterAtPoint(CPoint3d(arcSeg.StartPoint)), myPoly.GetParameterAtPoint(CPoint3d(arcSeg.EndPoint)))
                                If arcLen < hDistMin Then Continue For
                                If arcLen > hDistMax * 40000 Then
                                    MessageBox.Show("Wobble frequency is too short for this object.  Use a higher frequency or break object into smaller pieces.")
                                    Exit Sub
                                End If

                                Dim ls As Boolean

                                If i = myPoly.NumberOfVertices - 2 Then ls = True
                                Dim subpts As Point3dCollection = GetCircArcWobble(arcSeg, ls, arcReversed)

                                If subpts Is Nothing Then Exit Sub

                                For Each pt As Point3d In subpts
                                    myPts.Add(pt)
                                Next

                            End If
                        Next

                        'if type of object is an arc
                    ElseIf TypeOf acEnt Is Arc Then
                        Dim acArc As Arc = CType(acEnt, Arc)
                        Dim aPolyID As ObjectId = Arc2poly(acArc.ObjectId)
                        If aPolyID = ObjectId.Null Then Exit Sub

                        'convert to polyline and extract CircularArc2d
                        Dim myPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = acTrans.GetObject(aPolyID, OpenMode.ForRead)
                        Dim arcSeg As CircularArc2d = myPoly.GetArcSegment2dAt(0)
                        Dim arcLen As Double = (arcSeg.EndAngle - arcSeg.StartAngle) * arcSeg.Radius

                        'get arc direction
                        Dim arcReversed As Boolean
                        If arcSeg.IsClockWise Then
                            arcReversed = True
                        Else
                            arcReversed = False
                        End If

                        '(myPoly.GetParameterAtPoint(CPoint3d(arcSeg.StartPoint)), myPoly.GetParameterAtPoint(CPoint3d(arcSeg.EndPoint)))
                        If arcLen < hDistMin Then Continue For
                        If arcLen > hDistMax * 40000 Then
                            MessageBox.Show("Wobble frequency is too short for this object.  Use a higher frequency or break object into smaller pieces.")
                            Exit Sub
                        End If

                        Dim subpts As Point3dCollection = GetCircArcWobble(arcSeg, True, arcReversed)

                        For Each pt As Point3d In subpts
                            myPts.Add(pt)
                        Next

                        'if type of object is a line
                    ElseIf TypeOf acEnt Is Line Then
                        Using myLine As Autodesk.AutoCAD.DatabaseServices.Line = CType(acEnt, Line)
                            Using mypoly As New Autodesk.AutoCAD.DatabaseServices.Polyline
                                mypoly.AddVertexAt(0, New Point2d(myLine.StartPoint.X, myLine.StartPoint.Y), 0, 0, 0)
                                mypoly.AddVertexAt(1, New Point2d(myLine.EndPoint.X, myLine.EndPoint.Y), 0, 0, 0)

                                Dim lineSeg As LineSegment2d = mypoly.GetLineSegment2dAt(0)
                                If lineSeg.Length < hDistMin Then Continue For

                                If lineSeg.Length > hDistMax * 40000 Then
                                    MessageBox.Show("Wobble frequency is too short for this object.  Use a higher frequency or break object into smaller pieces.")
                                    Exit Sub
                                End If

                                Dim subPts As Point3dCollection = GetLineSegWobble(lineSeg, True)

                                For Each pt As Point3d In subPts
                                    myPts.Add(pt)
                                Next
                            End Using
                        End Using

                        'if type of object is a Circle
                    ElseIf TypeOf acEnt Is Circle Then
                        isCircle = True
                        Dim mycirc As Circle = CType(acEnt, Circle)
                        Dim rad As Double = mycirc.Radius
                        Dim xDist As Double = 0

                        Dim ctr As Point3d = mycirc.Center
                        Dim failsafe As Integer = 0
                        Dim lastone As Boolean = False
                        Try
                            Dim newAmp As Single = m_handMaxY
                            Dim rad1 As Double = rad + newAmp
                            Dim radneg1 As Double = rad - newAmp
                            Dim circumf As Double = 2 * PI * rad
                            'Dim totalDist As Double = 0

                            Do
                                Dim tempXmin = m_handMinX * 100
                                Dim tempXmax = m_handMaxX * 100
                                Dim tempVal As Integer = CInt(Int((tempXmax * Rnd()) + tempXmin))
                                Dim randTerm = tempVal / 100
                                xDist += randTerm
                                'Dim rotang As Double = xDist / rad

                                'Randomize()
                                'Dim tempymin = m_handMinY * 100
                                'Dim tempymax = m_handMaxY * 100
                                'Dim tempY As Integer = CInt(Int((tempymax * Rnd()) + tempymin))
                                'newAmp = tempY / 100

                                'Dim tmpLn As Line = l2d.Clone
                                Dim curAng As Double = xDist / rad

                                If xDist > circumf Then
                                    'If curAng > 2 * PI Then
                                    'rotAng = rotAng + (curAng - (2 * PI))
                                    curAng = 2 * PI
                                    lastone = True
                                    rad1 = rad
                                    radneg1 = rad
                                    'm_polarity = Not m_polarity
                                End If

                                If m_polarity Then
                                    Dim newX As Double = ctr.X + (rad1 * Cos(curAng))
                                    Dim newY As Double = ctr.Y + (rad1 * Sin(curAng))
                                    myPts.Add(New Point3d(newX, newY, 0))
                                Else
                                    Dim newX As Double = ctr.X + (radneg1 * Cos(curAng))
                                    Dim newY As Double = ctr.Y + (radneg1 * Sin(curAng))
                                    myPts.Add(New Point3d(newX, newY, 0))
                                End If

                                m_polarity = Not m_polarity
                                failsafe += 1

                                If lastone Then Exit Do

                                If failsafe = 40000 Then
                                    MessageBox.Show("Object too long or wobble period to short.  Break apart object or make wobble period longer.")
                                    Exit Sub
                                End If
                            Loop While failsafe < 40000

                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                            Exit Sub
                        End Try

                    Else
                        MessageBox.Show("Entity cannot be used to create handline.  Try converting entity to polyline.")
                        acTrans.Abort()
                        Exit Sub
                    End If

                    'if all points are created successfully...
                    Dim bt As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                    Dim mdlSpc As BlockTableRecord = acTrans.GetObject(dwgDB.CurrentSpaceId, OpenMode.ForWrite)

                    Dim mkSpline As Boolean = YesNoQuery("Do you want to create a spline? (No creates a standard polyline using line segments only)")

                    If mkSpline Then
                        Using newSpline As New Autodesk.AutoCAD.DatabaseServices.Spline(myPts, KnotParameterizationEnum.SqrtChord, 3, 0.05)
                            mdlSpc.AppendEntity(newSpline)
                            acTrans.AddNewlyCreatedDBObject(newSpline, True)
                            newSpline.LinetypeId = myLTId
                        End Using
                    Else
                        Using newPline As New Polyline
                            Dim pts2d As Point2dCollection = ConvertPoints2d(myPts)
                            For i = 0 To pts2d.Count - 1
                                Dim pt As Point2d = pts2d(i)
                                newPline.AddVertexAt(i, pt, 0, 0, 0)
                            Next
                            If isCircle Then newPline.Closed = True
                            mdlSpc.AppendEntity(newPline)
                            acTrans.AddNewlyCreatedDBObject(newPline, True)
                            newPline.LinetypeId = myLTId
                        End Using
                    End If
                Next

                acTrans.Commit()

            End Using

        End Sub


        Public Function GetLineSegWobble(lineseg As LineSegment2d, lastSeg As Boolean) As Point3dCollection

            Dim myPts As New Point3dCollection

            'create a temporary line entity from linesegment2d
            Dim ln As New Line(CPoint3d(lineseg.StartPoint), CPoint3d(lineseg.EndPoint))
            Dim ang As Double = ln.Angle
            Dim segLen As Double = ln.Length

            Dim failSafe As Integer = 0

            'add first point at start of line
            myPts.Add(ln.StartPoint)
            Dim newAmp As Double = m_handMaxY
            Dim lnType As Integer = 0

            'determine direction of the line using start and end points
            If ln.StartPoint.X >= ln.EndPoint.X Then
                If ln.StartPoint.Y >= ln.EndPoint.Y Then
                    lnType = 1
                Else
                    lnType = 2
                End If
            ElseIf ln.StartPoint.X < ln.EndPoint.X Then
                If ln.StartPoint.Y <= ln.EndPoint.Y Then
                    lnType = 1
                Else
                    lnType = 2
                End If
            End If

            'get a random distance that is between min and max x distances
            Randomize()
            Dim tempXmin = m_handMinX * 100
            Dim tempXmax = m_handMaxX * 100
            Dim tempVal As Integer = CInt(Int((tempXmax * Rnd()) + tempXmin))
            Dim randTerm = tempVal / 100

            Dim xDist As Double = randTerm

            Try
                'if there is any distance left over from previous segment, use that distance for first point
                If m_xtra > 0 Then
                    xDist = m_xtra
                    m_xtra = 0
                End If

                Do
                    'if the current distance is longer than the segment, save extra distance and exit or create point at end of the line
                    If xDist > segLen Then
                        If lastSeg Then
                            myPts.Add(ln.EndPoint)
                        Else
                            m_xtra = xDist - segLen
                            m_polarity = Not m_polarity
                        End If
                        Exit Do
                    End If

                    Dim newPt As Point3d = ln.GetPointAtDist(xDist)

                    'add points to collection according to the direction of the line
                    If lnType = 1 Then
                        If m_polarity Then
                            Dim newX As Double = newPt.X + newAmp * Sin(ang)
                            Dim newY As Double = newPt.Y - newAmp * Cos(ang)
                            myPts.Add(New Point3d(newX, newY, 0))
                        Else
                            Dim newX As Double = newPt.X - newAmp * Sin(ang)
                            Dim newY As Double = newPt.Y + newAmp * Cos(ang)
                            myPts.Add(New Point3d(newX, newY, 0))
                        End If
                    ElseIf lnType = 2 Then
                        If m_polarity Then
                            Dim newX As Double = newPt.X - newAmp * Sin(ang)
                            Dim newY As Double = newPt.Y - newAmp * Cos(ang)
                            myPts.Add(New Point3d(newX, newY, 0))
                        Else
                            Dim newX As Double = newPt.X + newAmp * Sin(ang)
                            Dim newY As Double = newPt.Y + newAmp * Cos(ang)
                            myPts.Add(New Point3d(newX, newY, 0))
                        End If
                    End If

                    'get next distance and adjust by random parameter
                    Randomize()
                    tempVal = CInt(Int((tempXmax * Rnd()) + tempXmin))
                    randTerm = tempVal / 100
                    xDist += randTerm

                    m_polarity = Not m_polarity

                    failSafe += 1

                    If failSafe = 40000 Then
                        MessageBox.Show("Object too long or wobble period to short.  Break apart object or make wobble period longer.")
                        Return Nothing
                        ln.Dispose()
                        Exit Function
                    End If

                Loop While failSafe < 40000

                Return myPts
                ln.Dispose()
                Exit Function


            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Return Nothing
                ln.Dispose()
                Exit Function
            End Try

        End Function


        Public Function GetCircArcWobble(arcSeg As CircularArc2d, lastseg As Boolean, isReversed As Boolean) As Point3dCollection

            Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = acDwg.Database

            Dim ctr As Point2d = arcSeg.Center
            Dim rad As Double = arcSeg.Radius
            'Dim arcLen As Double = (arcSeg.EndAngle - arcSeg.StartAngle) * rad
            Dim arcLen As Double = arcSeg.GetLength(arcSeg.GetParameterOf(arcSeg.StartPoint), arcSeg.GetParameterOf(arcSeg.EndPoint))

            Dim endVect As Vector2d = arcSeg.EndPoint.GetVectorTo(ctr)
            Dim endAng As Double = arcSeg.ReferenceVector.Angle + arcSeg.EndAngle
            Dim startVect As Vector2d = arcSeg.StartPoint.GetVectorTo(ctr)
            Dim stAng As Double = arcSeg.StartAngle + arcSeg.ReferenceVector.Angle

            Dim failsafe As Integer = 0
            Dim rad1 As Double = rad + m_handMaxY
            Dim radNeg1 As Double = rad - m_handMaxY

            Dim xDist As Double
            Dim mypts As New Point3dCollection

            'Debug.Print(arcSeg.IsClockWise.ToString)
            'Debug.Print("ReferenceAngle: " & Round(arcSeg.ReferenceVector.Angle * 180 / PI, 2) & "degrees")
            'Debug.Print("StartAngle: " & Round(stAng * 180 / PI, 2) & " degrees")
            'Debug.Print("EndAngle: " & Round(endAng * 180 / PI, 2) & " degrees")
            'Debug.Print("")

            'use function parameter to establish start and end angles
            If isReversed Then
                Dim tngl As Double = stAng
                stAng = endAng
                endAng = tngl
            End If

            'begin wobble algorithm
            Try
                Dim curAng As Double
                Dim rotAng As Double

                'first see if there is any distance left over from the previous segment by checking m_xtra
                If m_xtra > 0 Then
                    xDist = m_xtra
                    m_xtra = 0
                    'rotAng = xDist / rad
                Else
                    xDist = 0
                End If

                'use the xdist to calculate change in angle if there is any distance left over
                rotAng = xDist / rad

                'check arc direction and add or subtract angle from start of arc
                If arcSeg.IsClockWise Then
                    curAng = stAng - rotAng
                Else
                    curAng = stAng + rotAng
                End If

                'use polarity to determine wobble direction (radius distance)
                'add first point to the collection
                If m_polarity Then
                    Dim newX As Double = ctr.X + (rad1 * Cos(curAng))
                    Dim newY As Double = ctr.Y + (rad1 * Sin(curAng))
                    mypts.Add(New Point3d(newX, newY, 0))
                    m_polarity = Not m_polarity
                Else
                    Dim newX As Double = ctr.X + (radNeg1 * Cos(curAng))
                    Dim newY As Double = ctr.Y + (radNeg1 * Sin(curAng))
                    mypts.Add(New Point3d(newX, newY, 0))
                    m_polarity = Not m_polarity
                End If

                'main wobble loop
                Do
                    'initialize random number gneerator
                    'multiply min and max x by 100 before applying random number between 0 and 1
                    Randomize()
                    Dim tempXmin = m_handMinX * 100
                    Dim tempXmax = m_handMaxX * 100
                    Dim tempVal As Integer = CInt(Int((tempXmax * Rnd()) + tempXmin))
                    'divide by 100 to get final distance
                    Dim randTerm = tempVal / 100
                    xDist += randTerm
                    'convert distance to an angle
                    rotAng = xDist / rad

                    'if new distance is longer than the total arc length, exit the loop but save the extra distance for the next segment
                    If xDist > arcLen Then
                        'if last segment, use end point of arc
                        If lastseg Then
                            mypts.Add(CPoint3d(arcSeg.EndPoint))
                        Else
                            m_xtra = xDist - arcLen
                        End If
                        Exit Do
                    End If

                    'if xdist is within arc, check arc direction and add or subtract angle from start of arc
                    If arcSeg.IsClockWise Then
                        curAng = stAng - rotAng
                    Else
                        curAng = stAng + rotAng
                    End If

                    'use polarity to determine wobble direction (radius distance)
                    'add next point to the collection
                    If m_polarity Then
                        Dim newX As Double = ctr.X + (rad1 * Cos(curAng))
                        Dim newY As Double = ctr.Y + (rad1 * Sin(curAng))
                        mypts.Add(New Point3d(newX, newY, 0))
                    Else
                        Dim newX As Double = ctr.X + (radNeg1 * Cos(curAng))
                        Dim newY As Double = ctr.Y + (radNeg1 * Sin(curAng))
                        mypts.Add(New Point3d(newX, newY, 0))
                    End If

                    'reverse polarity
                    m_polarity = Not m_polarity

                    failsafe += 1

                    If failsafe = 40000 Then
                        MessageBox.Show("Object too long or wobble period to short.  Break apart object or make roughness period longer.")
                        Return Nothing
                        Exit Function
                    End If

                Loop While failsafe < 40000

                Return mypts
                Exit Function

            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Return Nothing
                Exit Function
            End Try

        End Function


        <CommandMethod("MKST")>
        Public Sub MkSt()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Try

                Dim ppo As New PromptPointOptions(vbLf & "Select the center of the inscribed circle.")
                With ppo
                    .AllowNone = False
                End With

                Dim ppr As PromptPointResult = ed.GetPoint(ppo)
                Dim cPt As Point3d

                If ppr.Status = PromptStatus.OK Then
                    cPt = ppr.Value
                Else
                    Exit Sub
                End If

                Dim pdro As New PromptDistanceOptions(vbLf & "Enter or pick the radius of the inscribed circle.")
                With pdro
                    .AllowNegative = False
                    .AllowNone = False
                    .UseBasePoint = True
                    .BasePoint = cPt
                    .UseDashedLine = True
                    .AllowArbitraryInput = True
                    .AllowZero = False
                End With

                Dim pdr1 As PromptDoubleResult = ed.GetDistance(pdro)
                Dim rad As Double

                If pdr1.Status = PromptStatus.OK Then
                    rad = pdr1.Value
                Else
                    Exit Sub
                End If

                Dim pio As New PromptIntegerOptions(vbLf & "Enter the number of star points (5 minimum)")
                With pio
                    .AllowNegative = False
                    .AllowZero = False
                    .LowerLimit = 5
                    .AllowArbitraryInput = True
                End With

                Dim pts As Integer
                Dim pir As PromptIntegerResult = ed.GetInteger(pio)

                If pir.Status = PromptStatus.OK Then
                    pts = pir.Value
                Else
                    Exit Sub
                End If

                Dim centAngle As Double = (2 * PI) / pts
                Dim tips As New Point2dCollection

                For i As Integer = 0 To pts - 1
                    Dim tempX As Double = rad * Cos(i * centAngle)
                    Dim tempy As Double = rad * Sin(i * centAngle)
                    tips.Add(New Point2d(tempX, tempy))
                Next

                'figure out the type of star to make
                Dim sType As Integer = 1
                Dim starPts As New Point2dCollection
                Dim pl As New Autodesk.AutoCAD.DatabaseServices.Polyline
                Dim transVect As Vector3d = New Point3d(0, 0, 0).GetVectorTo(cPt)

                Dim openShape As Boolean = YesNoQuery(vbLf & "Plot star as open shape?")
                Dim linLst As New List(Of Line2d)
                Dim acLinLst As New List(Of Line)

                Dim typeMax As Integer = pts \ 2 - 2

                If pts > 6 Then

                    'get the type of star to draw  The more points there are the more types of stars are possible
                    Dim pio2 As New PromptIntegerOptions("")

                    Dim msg As String
                    If pts < 9 Then
                        msg = vbLf & "Create Type 1 or type 2 star?"
                        pio2.UpperLimit = 2
                    ElseIf pts > 8 And pts < 11 Then
                        msg = vbLf & "Create Type 1, 2, or 3 star?"
                        pio2.UpperLimit = 3
                    Else
                        msg = vbLf & "Create Type 1, 2, 3, or up to type " & typeMax.ToString & " star?"
                    End If

                    With pio2
                        .Message = msg
                        .AllowNegative = False
                        .AllowZero = False
                        .LowerLimit = 1
                        .AllowArbitraryInput = True
                    End With

                    Dim pir2 As PromptIntegerResult = ed.GetInteger(pio2)

                    If pir2.Status = PromptStatus.OK Then
                        sType = pir2.Value
                    Else
                        Exit Sub
                    End If
                Else
                    'only one type for 5 or 6 pointed star
                    sType = 1
                End If

                If sType = 1 Then

                    For p As Integer = 0 To pts - 1
                        Dim r As Integer = (p + 2) Mod pts
                        linLst.Add(New Line2d(tips(p), tips(r)))
                        acLinLst.Add(New Line(New Point3d(tips(p).X, tips(p).Y, 0), New Point3d(tips(r).X, tips(r).Y, 0)))
                        'z = (z + 2) Mod pts
                    Next

                    'Dim j As Integer = 0

                    If openShape Then
                        For j As Integer = 0 To pts - 1
                            starPts.Add(tips(j))
                            Dim testLn1 As Line2d = linLst(j)
                            Dim m As Integer = (j + pts - 1) Mod pts
                            Dim testLn2 As Line2d = linLst(m)
                            Dim tempPt As Point2d() = testLn1.IntersectWith(testLn2)
                            starPts.Add(tempPt(0))
                        Next
                    End If

                ElseIf sType = 2 Then

                    For z As Integer = 0 To pts - 1
                        Debug.Print(tips(z).ToString)
                        Dim r As Integer = (z + 3) Mod pts
                        linLst.Add(New Line2d(tips(z), tips(r)))
                        acLinLst.Add(New Line(New Point3d(tips(z).X, tips(z).Y, 0), New Point3d(tips(r).X, tips(r).Y, 0)))
                    Next

                    'Dim j As Integer = 0

                    If openShape Then
                        For j As Integer = 0 To pts - 1
                            starPts.Add(tips(j))
                            Dim testLn1 As Line2d = linLst(j)
                            Dim s As Integer = (j + pts - 2) Mod pts
                            Dim testLn2 As Line2d = linLst(s)
                            Dim tempPt As Point2d() = testLn1.IntersectWith(testLn2)
                            starPts.Add(tempPt(0))
                        Next
                    End If

                ElseIf sType = 3 Then

                    For z As Integer = 0 To pts - 1
                        'Debug.Print(tips(z).ToString)
                        Dim r As Integer = (z + 4) Mod pts
                        linLst.Add(New Line2d(tips(z), tips(r)))
                        acLinLst.Add(New Line(New Point3d(tips(z).X, tips(z).Y, 0), New Point3d(tips(r).X, tips(r).Y, 0)))
                    Next

                    'Dim j As Integer = 0

                    If openShape Then
                        For j As Integer = 0 To pts - 1
                            starPts.Add(tips(j))
                            Dim testLn1 As Line2d = linLst(j)
                            Dim s As Integer = (j + pts - 3) Mod pts
                            Dim testLn2 As Line2d = linLst(s)
                            Dim tempPt As Point2d() = testLn1.IntersectWith(testLn2)
                            starPts.Add(tempPt(0))
                        Next
                    End If

                Else

                    For z As Integer = 0 To pts - 1
                        'Debug.Print(tips(z).ToString)
                        Dim r As Integer = (z + (sType + 1)) Mod pts
                        linLst.Add(New Line2d(tips(z), tips(r)))
                        acLinLst.Add(New Line(New Point3d(tips(z).X, tips(z).Y, 0), New Point3d(tips(r).X, tips(r).Y, 0)))
                    Next

                    'Dim j As Integer = 0

                    If openShape Then
                        For j As Integer = 0 To pts - 1
                            starPts.Add(tips(j))
                            Dim testLn1 As Line2d = linLst(j)
                            Dim s As Integer = (j + pts - sType) Mod pts
                            Dim testLn2 As Line2d = linLst(s)
                            Dim tempPt As Point2d() = testLn1.IntersectWith(testLn2)
                            starPts.Add(tempPt(0))
                        Next
                    End If

                End If

                If openShape Then
                    For x As Integer = 0 To starPts.Count - 1
                        pl.AddVertexAt(x, starPts(x), 0, 0, 0)
                    Next

                    pl.Closed = True
                    pl.TransformBy(Matrix3d.Displacement(transVect))

                    If pts Mod 2 = 1 Then
                        Dim rotAng As Double = centAngle / 4
                        pl.TransformBy(Matrix3d.Rotation(rotAng, Vector3d.ZAxis, cPt))
                    End If

                    Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                        Dim blktbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                        Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blktbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        mdlSpace.AppendEntity(pl)
                        acTrans.AddNewlyCreatedDBObject(pl, True)
                        acTrans.Commit()
                    End Using

                    If pl IsNot Nothing Then pl.Dispose()
                Else
                    Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                        For p As Integer = 0 To acLinLst.Count - 1
                            Dim blktbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                            Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blktbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                            Using myLine As Line = acLinLst(p)
                                'Dim acLine As New Line(New Point3d(myLine.StartPoint.X, myLine.StartPoint.Y, 0), New Point3d(myLine.EndPoint.X, myLine.EndPoint.Y, 0))
                                myLine.TransformBy(Matrix3d.Displacement(transVect))
                                If pts Mod 2 = 1 Then
                                    Dim rotAng As Double = centAngle / 4
                                    myLine.TransformBy(Matrix3d.Rotation(rotAng, Vector3d.ZAxis, cPt))
                                End If
                                mdlSpace.AppendEntity(myLine)
                                acTrans.AddNewlyCreatedDBObject(myLine, True)
                            End Using
                        Next
                        acTrans.Commit()
                    End Using
                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        <CommandMethod("MKSTE")>
        Public Sub MkStarEllipse()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Try

                Dim ppo As New PromptPointOptions(vbLf & "Select the center of the inscribed ellipse.")
                With ppo
                    .AllowNone = False
                End With

                Dim ppr As PromptPointResult = ed.GetPoint(ppo)
                Dim cPt As Point3d

                If ppr.Status = PromptStatus.OK Then
                    cPt = ppr.Value
                Else
                    Exit Sub
                End If

                Dim cPt2d As New Point2d(cPt.X, cPt.Y)

                Dim pdrao As New PromptPointOptions(vbLf & "pick the radius of the major axis of the ellipse")
                With pdrao
                    .AllowNone = False
                    .UseBasePoint = True
                    .BasePoint = cPt
                    .UseDashedLine = True
                    .AllowArbitraryInput = True
                End With

                Dim pdra1 As PromptPointResult = ed.GetPoint(pdrao)
                Dim rada As Double
                Dim ptA As Point2d
                Dim elAng As Double

                If pdra1.Status = PromptStatus.OK Then
                    Dim ptATemp As Point3d = pdra1.Value
                    ptA = New Point2d(ptATemp.X, ptATemp.Y)
                    rada = ptA.GetDistanceTo(cPt2d)
                    elAng = GetAngleFromXaxis(cPt, ptATemp)
                Else
                    Exit Sub
                End If

                Dim pdrbo As New PromptPointOptions(vbLf & "Enter or pick the radius of the minor axis of the ellipse")
                With pdrbo
                    .AllowNone = False
                    .UseBasePoint = True
                    .BasePoint = cPt
                    .UseDashedLine = True
                    .AllowArbitraryInput = True
                End With

                Dim pdrb1 As PromptPointResult = ed.GetPoint(pdrbo)
                Dim ptB As Point2d
                Dim radb As Double

                If pdrb1.Status = PromptStatus.OK Then
                    Dim ptBTemp As Point3d = pdrb1.Value
                    ptB = New Point2d(ptBTemp.X, ptBTemp.Y)
                    radb = ptB.GetDistanceTo(cPt2d)
                Else
                    Exit Sub
                End If

                Dim pio As New PromptIntegerOptions(vbLf & "Enter the number of star points (5 minimum)")
                With pio
                    .AllowNegative = False
                    .AllowZero = False
                    .LowerLimit = 5
                    .AllowArbitraryInput = True
                End With

                Dim pts As Integer
                Dim pir As PromptIntegerResult = ed.GetInteger(pio)

                If pir.Status = PromptStatus.OK Then
                    pts = pir.Value
                Else
                    Exit Sub
                End If

                Dim centAngle As Double = (2 * PI) / pts
                Dim tips As New Point2dCollection

                For i As Integer = 0 To pts - 1
                    Dim ca As Double = (centAngle * i) + centAngle / 4
                    Dim radAtPt As Double = (rada * radb) / Sqrt((rada ^ 2 * Sin(ca) ^ 2) + (radb ^ 2 * Cos(ca) ^ 2))
                    Dim tempx As Double = radAtPt * Cos(ca)
                    Dim tempy As Double = radAtPt * Sin(ca)
                    tips.Add(New Point2d(tempx, tempy))
                Next

                'figure out the type of star to make
                Dim sType As Integer = 1
                Dim starPts As New Point2dCollection
                Dim pl As New Autodesk.AutoCAD.DatabaseServices.Polyline
                Dim transVect As Vector3d = New Point3d(0, 0, 0).GetVectorTo(cPt)


                If pts > 6 Then

                    'get the type of star to draw  The more points there are the more types of stars are possible

                    Dim pio2 As New PromptIntegerOptions("")

                    Dim msg As String
                    If pts < 9 Then
                        msg = vbLf & "Type 1 or type 2 star?"
                        pio2.UpperLimit = 2
                    ElseIf pts > 8 And pts < 11 Then
                        msg = vbLf & "Type 1, 2, or 3 star?"
                        pio2.UpperLimit = 3
                    Else
                        msg = vbLf & "Type 1, 2, 3, or 4 star?"
                        pio2.UpperLimit = 4
                    End If

                    With pio2
                        .Message = msg
                        .AllowNegative = False
                        .AllowZero = False
                        .LowerLimit = 1
                        .AllowArbitraryInput = True
                    End With

                    Dim pir2 As PromptIntegerResult = ed.GetInteger(pio2)

                    If pir2.Status = PromptStatus.OK Then
                        sType = pir2.Value
                    Else
                        Exit Sub
                    End If
                Else
                    'only one type for 5 or 6 pointed star
                    sType = 1
                End If

                If sType = 1 Then
                    Dim linLst As New List(Of Line2d)

                    For p As Integer = 0 To pts - 1
                        Debug.Print(tips(p).ToString)
                        Dim r As Integer = (p + 2) Mod pts
                        linLst.Add(New Line2d(tips(p), tips(r)))
                        'z = (z + 2) Mod pts
                    Next

                    'Dim j As Integer = 0

                    For j As Integer = 0 To pts - 1
                        starPts.Add(tips(j))
                        Dim testLn1 As Line2d = linLst(j)
                        Dim m As Integer = (j + pts - 1) Mod pts
                        Dim testLn2 As Line2d = linLst(m)
                        Dim tempPt As Point2d() = testLn1.IntersectWith(testLn2)
                        starPts.Add(tempPt(0))
                    Next

                ElseIf sType = 2 Then

                    Dim linLst As New List(Of Line2d)

                    For z As Integer = 0 To pts - 1
                        Debug.Print(tips(z).ToString)
                        Dim r As Integer = (z + 3) Mod pts
                        linLst.Add(New Line2d(tips(z), tips(r)))
                    Next

                    'Dim j As Integer = 0

                    For j As Integer = 0 To pts - 1
                        starPts.Add(tips(j))
                        Dim testLn1 As Line2d = linLst(j)
                        Dim s As Integer = (j + pts - 2) Mod pts
                        Dim testLn2 As Line2d = linLst(s)
                        Dim tempPt As Point2d() = testLn1.IntersectWith(testLn2)
                        starPts.Add(tempPt(0))
                    Next

                ElseIf sType = 3 Then
                    Dim linLst As New List(Of Line2d)

                    For z As Integer = 0 To pts - 1
                        Debug.Print(tips(z).ToString)
                        Dim r As Integer = (z + 4) Mod pts
                        linLst.Add(New Line2d(tips(z), tips(r)))
                    Next

                    'Dim j As Integer = 0

                    For j As Integer = 0 To pts - 1
                        starPts.Add(tips(j))
                        Dim testLn1 As Line2d = linLst(j)
                        Dim s As Integer = (j + pts - 3) Mod pts
                        Dim testLn2 As Line2d = linLst(s)
                        Dim tempPt As Point2d() = testLn1.IntersectWith(testLn2)
                        starPts.Add(tempPt(0))
                    Next

                Else

                    Dim linLst As New List(Of Line2d)

                    For z As Integer = 0 To pts - 1
                        Debug.Print(tips(z).ToString)
                        Dim r As Integer = (z + 5) Mod pts
                        linLst.Add(New Line2d(tips(z), tips(r)))
                    Next

                    'Dim j As Integer = 0

                    For j As Integer = 0 To pts - 1
                        starPts.Add(tips(j))
                        Dim testLn1 As Line2d = linLst(j)
                        Dim s As Integer = (j + pts - 4) Mod pts
                        Dim testLn2 As Line2d = linLst(s)
                        Dim tempPt As Point2d() = testLn1.IntersectWith(testLn2)
                        starPts.Add(tempPt(0))
                    Next
                End If

                For x As Integer = 0 To starPts.Count - 1
                    pl.AddVertexAt(x, starPts(x), 0, 0, 0)
                Next

                pl.Closed = True
                pl.TransformBy(Matrix3d.Displacement(transVect))

                'If pts Mod 2 = 1 Then
                '    Dim rotAng As Double = centAngle / 4
                '    pl.TransformBy(Matrix3d.Rotation(rotAng, Vector3d.ZAxis, cPt))
                'End If

                Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                    Dim blktbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blktbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    mdlSpace.AppendEntity(pl)
                    acTrans.AddNewlyCreatedDBObject(pl, True)
                    acTrans.Commit()
                End Using

                If pl IsNot Nothing Then pl.Dispose()

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        <CommandMethod("MKSTAR")>
        Public Sub MkStar()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Try

                Dim ppo As New PromptPointOptions(vbLf & "Select the center of the inscribed circle.")
                With ppo
                    .AllowNone = False
                End With

                Dim ppr As PromptPointResult = ed.GetPoint(ppo)
                Dim cPt As Point3d

                If ppr.Status = PromptStatus.OK Then
                    cPt = ppr.Value
                Else
                    Exit Sub
                End If

                Dim pdro As New PromptDistanceOptions(vbLf & "Enter or pick the radius of the inscribed circle.")
                With pdro
                    .AllowNegative = False
                    .AllowNone = False
                    .UseBasePoint = True
                    .BasePoint = cPt
                    .UseDashedLine = True
                    .AllowArbitraryInput = True
                    .AllowZero = False
                End With

                Dim pdr1 As PromptDoubleResult = ed.GetDistance(pdro)
                Dim rad As Double

                If pdr1.Status = PromptStatus.OK Then
                    rad = pdr1.Value
                Else
                    Exit Sub
                End If

                Dim pio As New PromptIntegerOptions(vbLf & "Enter the number of star points (5 minimum)")
                With pio
                    .AllowNegative = False
                    .AllowZero = False
                    .LowerLimit = 5
                    .AllowArbitraryInput = True
                End With

                Dim pts As Integer
                Dim pir As PromptIntegerResult = ed.GetInteger(pio)

                If pir.Status = PromptStatus.OK Then
                    pts = pir.Value
                Else
                    Exit Sub
                End If

                Dim centAngle As Double = (2 * PI) / pts
                Dim tips As New Point3dCollection

                For i As Integer = 0 To pts - 1
                    Dim tempX As Double = rad * Cos(i * centAngle)
                    Dim tempy As Double = rad * Sin(i * centAngle)
                    tips.Add(New Point3d(tempX, tempy, 0))
                Next

                Dim transVect As Vector3d = New Point3d(0, 0, 0).GetVectorTo(cPt)

                Dim linLst As New List(Of Line)

                Dim incr As Integer
                Dim figs As Integer = 1
                Dim nolines As Integer


                If pts Mod 5 > 0 AndAlso pts \ 2 > 5 Then
                    incr = pts \ 2 - 1
                ElseIf pts Mod 4 > 0 AndAlso pts \ 2 > 4 Then
                    incr = pts \ 2 - 1
                ElseIf pts Mod 3 > 0 AndAlso pts \ 2 > 3 Then
                    incr = pts \ 2 - 1
                ElseIf pts Mod 2 > 0 AndAlso pts \ 2 > 2 Then
                    incr = pts \ 2 - 1
                ElseIf pts = 5 Then
                    incr = 2
                Else
                    If pts Mod 5 = 0 Then
                        nolines = pts / 5 And pts / 2 > 5
                        incr = 5
                    ElseIf pts Mod 4 = 0 And pts / 2 > 4 Then
                        nolines = pts / 4
                        incr = 4
                    ElseIf pts Mod 3 = 0 And pts / 2 > 3 Then
                        nolines = pts / 3
                        incr = 3
                    Else
                        nolines = pts / 2
                        incr = 2
                    End If
                    figs = pts / nolines
                End If

                Dim startPt As Integer = 0

                If figs > 1 Then
                    For f = 0 To figs - 1
                        For p As Integer = 0 To (pts / figs) - 1
                            Dim r As Integer = (startPt + p + incr) Mod pts
                            linLst.Add(New Line(tips((startPt + p) Mod pts), tips(r)))
                            startPt += 1
                        Next
                    Next
                Else
                    For p As Integer = 0 To pts - 1
                        Dim r As Integer = (p + incr) Mod pts
                        linLst.Add(New Line(tips(p), tips(r)))
                    Next
                End If

                Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                    Dim blktbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blktbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    For x As Integer = 0 To linLst.Count - 1
                        Dim subLine As Line = linLst(x)
                        Dim rotAng As Double
                        If pts Mod 2 > 0 Then
                            rotAng = centAngle / 4
                        Else
                            rotAng = centAngle / 2
                        End If
                        subLine.TransformBy(Matrix3d.Displacement(transVect))
                        subLine.TransformBy(Matrix3d.Rotation(rotAng, Vector3d.ZAxis, cPt))
                        mdlSpace.AppendEntity(subLine)
                        acTrans.AddNewlyCreatedDBObject(subLine, True)
                    Next
                    acTrans.Commit()

                End Using

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

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

        <CommandMethod("ND")>
        Public Sub NewDistance()
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            'Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Dim ppo As New PromptPointOptions(vbLf & "First point:")
            With ppo
                .AllowNone = False
                .AllowArbitraryInput = True
            End With
            Dim ppr As PromptPointResult = ed.GetPoint(ppo)

            Dim p1 As Point3d
            If ppr.Status = PromptStatus.OK Then
                p1 = ppr.Value
            Else
                Exit Sub
            End If

            With ppo
                .Message = vbLf & "Second point"
                .AllowNone = False
                .AllowArbitraryInput = True
                .UseBasePoint = True
                .BasePoint = p1
            End With

            Dim ppr2 = ed.GetPoint(ppo)

            Dim p2 As Point3d
            If ppr2.Status = PromptStatus.OK Then
                p2 = ppr2.Value
            Else
                Exit Sub
            End If

            Dim refzPlane As New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)
            Dim refyPlane As New Plane(New Point3d(0, 0, 0), Vector3d.YAxis)
            Dim refxPlane As New Plane(New Point3d(0, 0, 0), Vector3d.XAxis)

            If Not p1 = p2 Then
                Dim v3d As Vector3d = p1.GetVectorTo(p2)
                Dim dist3d As Double = v3d.Length

                Dim deltaX As Double = v3d.X
                Dim deltaY As Double = v3d.Y
                Dim deltaZ As Double = v3d.Z

                Dim vxy2d As Vector2d = v3d.Convert2d(refzPlane)
                Dim vxz2d As Vector2d = v3d.Convert2d(refyPlane)
                Dim vyz2d As Vector2d = v3d.Convert2d(refxPlane)
                Dim xyAng As Double = vxy2d.Angle
                Dim xyAngDegs As Double = CDegs(xyAng)
                Dim dist2D As Double = vxy2d.Length
                Dim xZang As Double = Vector2d.YAxis.GetAngleTo(vyz2d)
                Dim xZangDegs As Double = CDegs(xZang)
                Dim yZang As Double = Vector2d.XAxis.GetAngleTo(vxz2d)
                Dim yZangDegs As Double = CDegs(xZang)
                Dim slp As Double = deltaZ / dist2D

                Dim prec As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("LUPREC")
                Dim angPrec As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("AUPREC")

                Dim sb As New StringBuilder()
                sb.AppendLine()
                sb.AppendLine("3D Distance: " & Round(dist3d, prec).ToString & "  2D Distance: " & Round(dist2D, prec).ToString & "  Slope: " & Round(slp, prec).ToString)
                sb.AppendLine("Delta X: " & Round(deltaX, prec).ToString & "  Delta Y: " & Round(deltaY, prec).ToString & "  Delta Z: " & Round(deltaZ, prec).ToString)
                sb.AppendLine("Angle in XY Plane: " & Round(xyAngDegs, angPrec).ToString & " degrees")
                sb.AppendLine("Angle in XZ Plane: " & Round(xZangDegs, angPrec).ToString & " degrees")
                sb.AppendLine("Angle in YZ Plane: " & Round(yZangDegs, angPrec).ToString & " degrees")

                ed.WriteMessage(sb.ToString)
            Else
                Exit Sub
            End If

        End Sub

    End Module


    Public Module LayerCommands

        Private ReadOnly m_layIDList As ObjectIdCollection
        Private m_layIdDic As Dictionary(Of ObjectId, List(Of ObjectId))
        'Private m_layList As List(Of ObjectId)

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

        <CommandMethod("PICKVPFRZLAY", CommandFlags.UsePickSet)>
        Public Sub PickVPFrzLayer()
            'by David Eisenbeisz (c)2024
            'freezes viewport layers by picking entities and viewports

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

            Dim selSet As SelectionSet = SelResult.Value
            Dim obIds() As ObjectId = selSet.GetObjectIds

            'Dim layoutDic As SortedDictionary(Of Integer, String) = LayoutTabList()

            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction

                Dim fzLayers As New ObjectIdCollection

                For i = 0 To obIds.Count - 1
                    Dim ent As Entity = TryCast(acTrans.GetObject(obIds(i), OpenMode.ForRead), Entity)
                    If ent IsNot Nothing Then
                        Dim entLayID As ObjectId = ent.LayerId
                        fzLayers.Add(entLayID)
                        ent.Dispose()
                    End If
                Next

                Dim layoutDbDict As DBDictionary = acTrans.GetObject(DwgDB.LayoutDictionaryId, OpenMode.ForRead)

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

                For Each layo As String In layoutLst
                    Dim layoutID As ObjectId = layoutDbDict(layo)
                    Dim lo As Layout = acTrans.GetObject(layoutID, OpenMode.ForRead)

                    Dim vpIds As ObjectIdCollection = lo.GetViewports

                    Dim dr As PromptResult = YesNoResult(vbLf & "Do you want to freeze paperspace layers?")

                    Dim startnumb As Integer

                    If dr.Status = PromptStatus.OK Then
                        If dr.StringResult = "Yes" Then
                            startnumb = 0
                        Else
                            startnumb = 1
                        End If
                    End If

                    For m As Integer = startnumb To vpIds.Count - 1
                        Dim cVp As Viewport = acTrans.GetObject(vpIds(m), OpenMode.ForWrite)
                        cVp.FreezeLayersInViewport(fzLayers.GetEnumerator)
                    Next
                Next
                acTrans.Commit()
            End Using

        End Sub


        <CommandMethod("LAOFF", CommandFlags.UsePickSet Or CommandFlags.Redraw Or CommandFlags.Modal)>
        Public Sub TurnOffLayers()
            'by David Eisenbeisz (c)2024
            'temporarily turns off layers and stores them by space ID to turn on again.

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

            Dim cSpaceID As ObjectId = DwgDB.CurrentSpaceId

            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim offLayers As New List(Of ObjectId)
                For i = 0 To obIds.Count - 1
                    Dim ent As Entity = TryCast(actrans.GetObject(obIds(i), OpenMode.ForRead), Entity)
                    If ent IsNot Nothing Then
                        Dim entLayID As ObjectId = ent.LayerId
                        offLayers.Add(entLayID)
                    End If
                Next

                'If offLayers IsNot Nothing AndAlso offLayers.Count > 0 Then
                'If Not YesNoQuery(vbLf & "Do you want to clear the current list of turned off layers?  If Yes, only currently selected layers will be queued for turn on command.") Then
                'm_layIdDic(cSpaceID) = offLayers
                'tempList = m_layList
                'End If
                'End If

                Dim lrTbl As LayerTable = actrans.GetObject(DwgDB.LayerTableId, OpenMode.ForRead)

                Dim tempList As New List(Of ObjectId)

                For Each lyrId As ObjectId In offLayers
                    If lrTbl.Has(lyrId) Then tempList.Add(lyrId)
                Next

                Dim sortList As List(Of ObjectId) = tempList.Distinct.ToList

                For Each lID As ObjectId In sortList
                    Try
                        Dim ltr As LayerTableRecord = TryCast(actrans.GetObject(lID, OpenMode.ForWrite), LayerTableRecord)
                        ltr.IsOff = True
                    Catch ex As Exception
                        sortList.Remove(lID)
                        Continue For
                    End Try
                Next

                If m_layIdDic Is Nothing Then m_layIdDic = New Dictionary(Of ObjectId, List(Of ObjectId))
                m_layIdDic(cSpaceID) = sortList

                actrans.Commit()
            End Using

        End Sub

        <CommandMethod("LAON")>
        Public Sub TurnOnLayers()
            'by David Eisenbeisz (c)2024
            'turns on layers that were temporarily turned off
            'used with LAOFF command above

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Dim curSpcID As ObjectId = DwgDB.CurrentSpaceId
            Dim layIDList As List(Of ObjectId)

            If m_layIdDic IsNot Nothing AndAlso m_layIdDic.Keys.Contains(curSpcID) Then
                layIDList = m_layIdDic(curSpcID)
            Else
                ed.WriteMessage(vbLf & "No layers for current space stored in list.  Use LAOFF command or run LAON again from space where layers were turned off.")
                Exit Sub
            End If

            If layIDList.Count <= 0 Then
                ed.WriteMessage(vbLf & "No layers to turn on stored in list.  Use LAOFF command to turn off layers and add to list.")
                Exit Sub
            End If

            'Dim templist As List(Of String) = m_layList

            Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim lyrTbl As LayerTable = actrans.GetObject(DwgDB.LayerTableId, OpenMode.ForRead)
                For Each layId As ObjectId In layIDList
                    Try
                        If lyrTbl.Has(layId) Then
                            Dim ltr As LayerTableRecord = TryCast(actrans.GetObject(layId, OpenMode.ForWrite), LayerTableRecord)
                            If ltr IsNot Nothing Then
                                ltr.IsOff = False
                                'layIDList.Remove(layId)
                            End If
                        End If
                    Catch ex As Exception
                        Continue For
                    End Try
                Next
                m_layIdDic.Remove(curSpcID)
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
            Dim badLinetypes As New Dictionary(Of String, String)
            Dim hasBadPstyles As Boolean = False
            Dim hasBadLinetypes As Boolean = False

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
                Dim ltypeTbl As LinetypeTable = actrans.GetObject(dwgDB.LinetypeTableId, OpenMode.ForRead)
                Dim lz As LayerTableRecord = TryCast(actrans.GetObject(dwgDB.LayerZero, OpenMode.ForRead), LayerTableRecord)

                For i = 0 To layset.Count - 1
                    Dim aLayer As AcdLayer = layset.AcdLayer(i)
                    If String.IsNullOrEmpty(aLayer.Name) Then GoTo Skip
                    Try
                        If lTbl.Has(aLayer.Name) Then
                            Dim lTR As LayerTableRecord = TryCast(actrans.GetObject(lTbl(aLayer.Name), OpenMode.ForRead), LayerTableRecord)
                            If lTR = lz Then GoTo Skip
                            lTR.UpgradeOpen()

                            If aLayer.Remove Then
                                If Not String.IsNullOrEmpty(aLayer.MergeWith) Then
                                    If lTbl.Has(aLayer.MergeWith) Then MergeThenDeleteLayer(aLayer.Name, aLayer.MergeWith, True)
                                ElseIf String.IsNullOrEmpty(aLayer.MergeWith) Then
                                    DeleteMyLayer(aLayer.Name)
                                End If
                            ElseIf Not aLayer.Remove Then

                                If Not String.IsNullOrEmpty(aLayer.MergeWith) Then
                                    If lTbl.Has(aLayer.MergeWith) Then MergeThenDeleteLayer(aLayer.Name, aLayer.MergeWith, False)
                                End If
                            End If

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
                                Dim myLtId As ObjectId = GetLTId(aLayer.Linetype)
                                If myLtId = ObjectId.Null Then
                                    ed.WriteMessage(vbLf & "Using Continuous linetype for layer " & aLayer.Name)
                                    .LinetypeObjectId = ltypeTbl("Continuous")
                                    badLinetypes.Add(aLayer.Name, aLayer.Linetype)
                                    hasBadLinetypes = True
                                Else
                                    .LinetypeObjectId = myLtId
                                End If

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
                                Dim myLtId As ObjectId = GetLTId(aLayer.Linetype)
                                If myLtId = ObjectId.Null Then
                                    ed.WriteMessage(vbLf & "Using Continuous linetype for layer " & aLayer.Name)
                                    .LinetypeObjectId = ltypeTbl("Continuous")
                                    badLinetypes.Add(aLayer.Name, aLayer.Linetype)
                                    hasBadLinetypes = True
                                Else
                                    .LinetypeObjectId = myLtId
                                End If

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
                        If ex.ErrorStatus = Autodesk.AutoCAD.Runtime.ErrorStatus.AmbiguousInput Then
                            Err.Clear()
                        End If
                    End Try
Skip:
                Next
                actrans.Commit()
            End Using

            If hasBadPstyles Or hasBadLinetypes Then
                Dim dwgName As String = curDwg.Name
                Dim folderName As String = Path.GetDirectoryName(dwgName)
                Dim tFileName As String = folderName & "\BadLayers.txt"
                If File.Exists(tFileName) Then Kill(tFileName)

                Using sr As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(tFileName, False)
                    sr.WriteLine("Bad Plotstyles (changed to Normal)")
                    sr.WriteLine("Layer:Plotstyle")
                    For Each ky As String In badPstyles.Keys
                        Dim outStr As String = ky & ":" & badPstyles(ky)
                        sr.WriteLine(outStr)
                    Next
                    sr.WriteLine()
                    sr.WriteLine("Bad Linetypes (changed to Continuous)")
                    sr.WriteLine("Layer:Linetype")
                    For Each ky As String In badLinetypes.Keys
                        Dim outStr As String = ky & ":" & badLinetypes(ky)
                        sr.WriteLine(outStr)
                    Next
                End Using

                Dim erMsg As String = ""

                If hasBadPstyles Then
                    erMsg = "One or more plotstyles could not be assigned to the imported layers. " _
                             & "This occurs when a plotstyle is assigned to a layer before any objects in the drawing " _
                             & "are assigned that plotstyle.  To work around this bug, ensure that all needed plotstyles " _
                             & "are assigned to temporary objects before importing the layers.  A text file has been " _
                             & "created in the drawing folder that has the names of the failed layers and the plotstyles that could not be assigned.  " _
                             & "The Normal plotstyle has been assigned to these layers.  Run this command again after the missing plotsytles have been assigned to an object." & vbLf
                End If

                If hasBadLinetypes Then
                    erMsg = erMsg & "One or more linetypes could not be assigned to the imported layers.  The linetypes for these layers have been changed to Continuous " _
                               & "A text file has been created In the drawing folder that has the names Of the failed layers And the linetypes that could Not be assigned."
                End If

                MessageBox.Show(erMsg)

            End If

        End Sub

        <CommandMethod("UpdateLayers")>
        Public Sub UpdateExistingLayers()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database
            Dim xmlNameStr As String = GetMyXMLFileName()
            If String.IsNullOrEmpty(xmlNameStr) Then Exit Sub
            Dim mySchemaPath As String = "MasterCustomLibrary.LayerSchema1.xsd"

            Dim ss As XmlSchemaSet = GetLayerSchemaSet("")

            Dim sets As New XmlReaderSettings

            Dim badPstyles As New Dictionary(Of String, String)
            Dim badLinetypes As New Dictionary(Of String, String)
            Dim hasBadPstyles As Boolean = False
            Dim hasBadLinetypes As Boolean = False

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
                Dim ltypeTbl As LinetypeTable = actrans.GetObject(dwgDB.LinetypeTableId, OpenMode.ForRead)
                Dim lz As LayerTableRecord = TryCast(actrans.GetObject(dwgDB.LayerZero, OpenMode.ForRead), LayerTableRecord)

                For i = 0 To layset.Count - 1
                    Dim aLayer As AcdLayer = layset.AcdLayer(i)
                    If String.IsNullOrEmpty(aLayer.Name) Then GoTo Skip
                    Try
                        If Not lTbl.Has(aLayer.Name) Then
                            Continue For

                        ElseIf lTbl.Has(aLayer.Name) Then
                            Dim lTR As LayerTableRecord = TryCast(actrans.GetObject(lTbl(aLayer.Name), OpenMode.ForRead), LayerTableRecord)
                            If lTR = lz Then GoTo Skip
                            lTR.UpgradeOpen()

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
                                Dim myLtId As ObjectId = GetLTId(aLayer.Linetype)
                                If myLtId = ObjectId.Null Then
                                    ed.WriteMessage(vbLf & "Using Continuous linetype for layer " & aLayer.Name)
                                    .LinetypeObjectId = ltypeTbl("Continuous")
                                    badLinetypes.Add(aLayer.Name, aLayer.Linetype)
                                    hasBadLinetypes = True
                                Else
                                    .LinetypeObjectId = myLtId
                                End If
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

                    Catch ex As Exception
                        If ex.ErrorStatus = Autodesk.AutoCAD.Runtime.ErrorStatus.AmbiguousInput Then
                            Err.Clear()
                        End If
                    End Try
Skip:
                Next
                actrans.Commit()
            End Using

            If hasBadPstyles Or hasBadLinetypes Then
                Dim dwgName As String = curDwg.Name
                Dim folderName As String = Path.GetDirectoryName(dwgName)
                Dim tFileName As String = folderName & "\BadLayers.txt"
                If File.Exists(tFileName) Then Kill(tFileName)

                Using sr As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(tFileName, False)
                    sr.WriteLine("Bad Plotstyles (changed to Normal)")
                    sr.WriteLine("Layer:Plotstyle")
                    For Each ky As String In badPstyles.Keys
                        Dim outStr As String = ky & ":" & badPstyles(ky)
                        sr.WriteLine(outStr)
                    Next
                    sr.WriteLine()
                    sr.WriteLine("Bad Linetypes (changed to Continuous)")
                    sr.WriteLine("Layer:Linetype")
                    For Each ky As String In badLinetypes.Keys
                        Dim outStr As String = ky & ":" & badLinetypes(ky)
                        sr.WriteLine(outStr)
                    Next
                End Using

                Dim erMsg As String = ""

                If hasBadPstyles Then
                    erMsg = "One or more plotstyles could not be assigned to the imported layers. " _
                             & "This occurs when a plotstyle is assigned to a layer before any objects in the drawing " _
                             & "are assigned that plotstyle.  To work around this bug, ensure that all needed plotstyles " _
                             & "are assigned to temporary objects before importing the layers.  A text file has been " _
                             & "created in the drawing folder that has the names of the failed layers and the plotstyles that could not be assigned.  " _
                             & "The Normal plotstyle has been assigned to these layers.  Run this command again after the missing plotsytles have been assigned to an object." & vbLf
                End If

                If hasBadLinetypes Then
                    erMsg = erMsg & "One or more linetypes could not be assigned to the imported layers.  The linetypes for these layers has been changed to Continuous " _
                               & "A text file has been created In the drawing folder that has the names Of the failed layers And the linetypes that could Not be assigned."
                End If

                MessageBox.Show(erMsg)

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
                DeleteMyLayer(myLayer)
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
                delLayer = delLayerCol(1).ToString
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
                mergeLayer = mergeLayerCol(1).ToString
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
        Private m_pixScale As Double

        <CommandMethod("LLTOF")>
        Public Sub LayoutNamesToFile()

            ' Get the current document and database, and start a transaction

            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgdb As Database = acDoc.Database

            Dim layoDic As SortedDictionary(Of Integer, String) = LayoutTabList()
            'Dim tFile As String = getasavefilename("CSV File (*.csv)|*.csv|Text File (*.txt)|*.txt|", "Select or Create Text File")

            Dim tFile As String = GetASaveFileName("csv", "txt")
            If Not tFile = "" Then
                Try
                    Using wr As TextWriter = New StreamWriter(tFile, False)
                        Dim i As Integer
                        For i = 0 To layoDic.Keys.Count
                            If layoDic.Keys.Contains(i) Then
                                Dim lineStr As String = i & "," & layoDic(i)
                                wr.WriteLine(lineStr)
                            End If
                        Next
                        wr.Close()
                    End Using
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    Exit Sub
                End Try
            End If
        End Sub

        <CommandMethod("VNTOF")>
        Public Sub ViewNamesToFile()

            ' Get the current document and database, and start a transaction

            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgdb As Database = acDoc.Database

            Using acTrans As Transaction = dwgdb.TransactionManager.StartTransaction

                Dim vtb As ViewTable = acTrans.GetObject(dwgdb.ViewTableId, OpenMode.ForRead)
                'Dim tFile As String = getasavefilename("CSV File (*.csv)|*.csv|Text File (*.txt)|*.txt|", "Select or Create Text File")

                Dim tFile As String = GetASaveFileName("csv", "txt")
                If Not tFile = "" Then
                    Try
                        Using wr As TextWriter = New StreamWriter(tFile, False)
                            For Each objID As ObjectId In vtb
                                Dim vtr As ViewTableRecord = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim cp As Point2d = vtr.CenterPoint
                                Dim tlPoint As New Point2d(cp.X - vtr.Width / 2, cp.Y + vtr.Height / 2)
                                Dim trPoint As New Point2d(cp.X + vtr.Width / 2, cp.Y + vtr.Height / 2)
                                Dim brPoint As New Point2d(cp.X + vtr.Width / 2, cp.Y - vtr.Height / 2)
                                Dim lineArray(7) As String
                                lineArray(0) = vtr.Name
                                lineArray(1) = Round(tlPoint.X, 3).ToString
                                lineArray(2) = Round(tlPoint.Y, 3).ToString
                                lineArray(3) = Round(brPoint.X, 3).ToString
                                lineArray(4) = Round(brPoint.Y, 3).ToString
                                lineArray(5) = Round(trPoint.X, 3).ToString
                                lineArray(6) = Round(trPoint.Y, 3).ToString
                                Dim lineStr As String = Join(lineArray, ",")
                                wr.WriteLine(lineStr)
                            Next
                            wr.Close()
                        End Using
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                        Exit Sub
                    End Try
                End If
                acTrans.Commit()
            End Using

        End Sub


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

            Dim pca As PickCellsArgs = PickCells()
            If pca Is Nothing Then Exit Sub

            Dim horiz As Double = pca.Horizontal
            Dim vert As Double = pca.Vertical
            Dim p1 As Point3d = pca.Point1
            Dim p2 As Point3d = pca.Point2

            'Dim ppo As New PromptPointOptions(vbLf & "Pick upper left corner of first cell")
            'With ppo
            '    .AllowNone = False
            '    .AllowArbitraryInput = True
            'End With

            'Dim p1 As Point3d
            'Dim p2 As Point3d
            'Dim userPick As Boolean = True

            'Dim ppr As PromptPointResult = ed.GetPoint(ppo)

            'If ppr.Status = PromptStatus.OK Then
            '    p1 = ppr.Value

            '    Dim pco As New PromptCornerOptions(vbLf & "Pick the lower right corner of first cell.", p1)
            '    With pco
            '        .UseDashedLine = True
            '        .AllowArbitraryInput = True
            '    End With

            '    Dim pcr As PromptPointResult = ed.GetCorner(pco)

            '    If pcr.Status = PromptStatus.OK Then
            '        p2 = pcr.Value
            '    Else
            '        userPick = False
            '    End If
            'Else
            '    userPick = False
            'End If

            'Dim vert As Double
            'Dim horiz As Double

            'If userPick Then
            '    horiz = p2.X - p1.X
            '    vert = p2.Y - p1.Y
            'Else
            '    Dim pdo As New PromptDistanceOptions(vbLf & "Input or pick the horizontal distance for each cell")
            '    With pdo
            '        .AllowNegative = True
            '        .Only2d = True
            '        .AllowNone = False
            '    End With

            '    Dim pdr As PromptDoubleResult = ed.GetDistance(pdo)

            '    If pdr.Status = PromptStatus.OK Then
            '        horiz = pdr.Value
            '    Else
            '        Exit Sub
            '    End If

            '    Dim pdo2 As New PromptDistanceOptions(vbLf & "Input or pick the vertical distance for each cell")
            '    Dim pdr2 As PromptDoubleResult = ed.GetDistance(pdo2)

            '    If pdr2.Status = PromptStatus.OK Then
            '        vert = pdr2.Value
            '    Else
            '        Exit Sub
            '    End If
            'End If

            Dim cCols As Integer
            Dim cRows As Integer

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

            If LayerExists("Vports") Then
                vpLayer = "Vports"
            Else
                vpLayer = AddNewLayer("Vports", 3)
            End If

            'Dim showForm As Int16 = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("SHOWPAGESETUPFORM")
            Dim cVp As Int16 = DirectCast(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("LAYOUTCREATEVIEWPORT"), Int16)
            Dim showForm As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("SHOWPAGESETUPFORNEWLAYOUTS")
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("SHOWPAGESETUPFORNEWLAYOUTS", 0)
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("LAYOUTCREATEVIEWPORT", 0)

            Try

                Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction

                    Dim shtNm As String = ""

                    Dim pSet As PlotSettings = Nothing
                    Dim psetVal As PlotSettingsValidator = PlotSettingsValidator.Current
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
                            'psetVal = PlotSettingsValidator.Current
                            psetVal.RefreshLists(pSet)
                            myPltr = GetPrinter(psetVal)
                            psetVal.SetPlotConfigurationName(pSet, myPltr, Nothing)

                            'psetVal.SetPlotPaperUnits(pSet, PlotPaperUnit.Inches)
                        Else
                            'psetVal = PlotSettingsValidator.Current
                            pSet = plsets.GetAt(mySetup).GetObject(OpenMode.ForWrite)
                            psetVal.RefreshLists(pSet)
                            myPltr = pSet.PlotConfigurationName
                        End If
                        shtNm = GetSheetName(psetVal, pSet)
                    End If

                    Dim usePixels As Boolean = YesNoQuery(vbLf & "Will the layouts be plotted as raster images?")

                    If usePixels Then
                        Dim pixScale As Integer = GetIntegerValue(vbLf & "Input the number of pixels per paperspace inch: ")
                        Dim cs As New CustomScale(pixScale, 1)
                        psetVal.SetCustomPrintScale(pSet, cs)
                    Else
                        psetVal.SetStdScale(pSet, True)
                        psetVal.SetStdScaleType(pSet, StdScaleType.StdScale1To1)
                        psetVal.SetPlotOrigin(pSet, New Point2d(0, 0))
                        psetVal.SetPlotType(pSet, PlotType.Layout)
                    End If

                    Dim vtb As ViewTable = acTrans.GetObject(dwgDB.ViewTableId, OpenMode.ForWrite)

                    Dim useNumbViews As Boolean = YesNoQuery("Use numbered views?")
                    Dim viewNames As New List(Of String)
                    Dim viewList As New List(Of Integer)
                    Dim transDic As New Dictionary(Of String, String)

                    If useNumbViews Then

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

                        i = startNo

                    Else
                        Dim fName As String = GetMyCSVFileName()
                        If String.IsNullOrEmpty(fName) Then Exit Sub
                        Dim translist As New List(Of String())

                        Using sr As New TextFieldParser(fName)
                            sr.TextFieldType = FileIO.FieldType.Delimited
                            sr.SetDelimiters(",")
                            Dim curRow() As String = sr.ReadFields
                            translist.Add(curRow)
                            While Not sr.EndOfData
                                Try
                                    curRow = sr.ReadFields
                                    translist.Add(curRow)
                                Catch ex As FileIO.MalformedLineException
                                    MsgBox("Line " & curRow.ToString & "is not valid and will be skipped.")
                                End Try
                            End While
                            sr.Close()
                        End Using

                        For j As Integer = 0 To translist.Count - 1
                            Try
                                Dim curLine() As String = translist(j)
                                Dim vN As String = curLine(1)
                                Dim loN As String = curLine(0)
                                If vtb.Has(vN) Then Continue For
                                transDic(loN) = vN
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                                Continue For
                            End Try
                        Next
                    End If

                    Dim starthoriz As Double = p1.X + (horiz / 2)
                    Dim startvert As Double = p1.Y + (vert / 2)

                    Dim endHoriz As Double = p1.X + (horiz * cCols) - (horiz / 2)
                    Dim endVert As Double = p1.Y + (vert * cRows) - (vert / 2)
                    Dim c As Integer = 0

                    For y = 0 To cRows - 1
                        Dim yCtr As Double = startvert + (vert * y)
                        For x As Integer = 0 To cCols - 1
                            Dim xCtr As Double = starthoriz + (horiz * x)
                            'For y As Integer = startvert To endVert Step vert
                            '    For x As Integer = starthoriz To endHoriz Step horiz
                            Dim vtr As ViewTableRecord

                            If useNumbViews Then
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
                            Else
                                Dim myVName As String = transDic(transDic.Keys(c))
                                If vtb.Has(myVName) Then
                                    vtr = acTrans.GetObject(vtb(myVName.ToString), OpenMode.ForWrite)
                                    With vtr
                                        .Width = Abs(horiz)
                                        .Height = Abs(vert)
                                        .CenterPoint = New Point2d(xCtr, yCtr)
                                    End With
                                Else
                                    vtr = New ViewTableRecord
                                    With vtr
                                        .Name = myVName
                                        .Width = Abs(horiz)
                                        .Height = Abs(vert)
                                        .CenterPoint = New Point2d(xCtr, yCtr)
                                    End With
                                    Dim vtrid As ObjectId = vtb.Add(vtr)
                                    acTrans.AddNewlyCreatedDBObject(vtr, True)
                                End If
                            End If

                            i += 1

                            If makeLayouts AndAlso pSet IsNot Nothing AndAlso shtNm IsNot Nothing Then
                                'Dim ptrName As String = pSet.PlotConfigurationName
                                Dim layoutID As ObjectId
                                If useNumbViews Then
                                    layoutID = CreateVP(vtr, vpLayer, hSize, vSize, acTrans, pSet, shtNm)
                                Else
                                    layoutID = CreateVP(vtr, vpLayer, hSize, vSize, acTrans, pSet, shtNm, transDic.Keys(c))
                                End If

                                c += 1

                                If layoutID = ObjectId.Null Then
                                    MessageBox.Show(vbLf & "Error creating layout for " & vtr.Name & ".")
                                    Continue For
                                End If
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

        <CommandMethod("RENVIEWS")>
        Public Sub RenameViews()

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            ed.WriteMessage(vbLf & "This command uses a csv file (in the form NewName, OldName) to translate existing view names to new view names.")

            Dim fName As String = GetMyCSVFileName()
            If String.IsNullOrEmpty(fName) Then Exit Sub
            Dim transList As New List(Of String())

            Using sr As New TextFieldParser(fName)
                sr.TextFieldType = FileIO.FieldType.Delimited
                sr.SetDelimiters(",")
                Dim curRow() As String = sr.ReadFields
                transList.Add(curRow)
                While Not sr.EndOfData
                    Try
                        curRow = sr.ReadFields
                        transList.Add(curRow)
                    Catch ex As FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                    End Try
                End While
                sr.Close()
            End Using

            Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim vtb As ViewTable = actrans.GetObject(dwgDB.ViewTableId, OpenMode.ForRead)

                For i As Integer = 0 To transList.Count - 1
                    Try
                        Dim curLine() As String = transList(i)
                        Dim oldVn As String = curLine(1)
                        Dim newVn As String = curLine(0)
                        If oldVn = newVn Then Continue For
                        Dim vtr As ViewTableRecord = TryCast(actrans.GetObject(vtb(oldVn), OpenMode.ForWrite), ViewTableRecord)
                        If vtr IsNot Nothing Then
                            vtr.Name = newVn
                        End If
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                        Continue For
                    End Try
                Next

                actrans.Commit()
            End Using

        End Sub



        <CommandMethod("RLOLD")>
        Public Sub RenameLayoutsOld()

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
            Dim revList As New Dictionary(Of String, Integer)
            For Each k As Integer In loList.Keys
                revList(loList(k)) = k
            Next


            For Each z As Integer In loList.Keys
                Try
                    If loList.Keys.Contains(z) Then
                        Dim lName As String = loList(z)
                        If Not lName = "Model" Then
                            lPicker.ListBox1.Items.Add(lName)
                        End If
                    End If

                Catch ex As Exception
                    Continue For
                End Try

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

                    Dim hSize As Double
                    Dim vSize As Double

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

                    Dim viewsForChange As New List(Of String)
                    Dim i As Integer

                    Dim useNumbViews As Boolean = YesNoQuery("Use numbered views?")

                    Using vtb As ViewTable = acTrans.GetObject(dwgDB.ViewTableId, OpenMode.ForRead)
                        Dim tempViewDic As New Dictionary(Of Integer, String)
                        Dim viewDic As New Dictionary(Of String, String)

                        If useNumbViews Then
                            'get first view to be adjusted (must be numbered views)
                            Dim pdo7 As New PromptIntegerOptions(vbLf & "Input starting view number")
                            Dim pdr7 As PromptIntegerResult = ed.GetInteger(pdo7)

                            Dim startView As Integer

                            If pdr7.Status = PromptStatus.OK Then
                                startView = pdr7.Value
                            Else
                                Exit Sub
                            End If

                            i = startView - 1
                            'get the view table record

                        Else
                            'Dim vtb As ViewTable = acTrans.GetObject(dwgDB.ViewTableId, OpenMode.ForRead)
                            Dim viewList As New List(Of String)
                            For Each viewID As ObjectId In vtb
                                Dim myVtr As ViewTableRecord = TryCast(acTrans.GetObject(viewID, OpenMode.ForRead), ViewTableRecord)
                                If myVtr IsNot Nothing Then viewList.Add(myVtr.Name)
                            Next

                            Using vPicker As New LayoutPicker(viewList, False)
                                vPicker.ShowDialog()

                                If Not vPicker.DialogResult = DialogResult.OK Then
                                    acTrans.Abort()
                                    Exit Sub
                                End If


                                If vPicker.PickedList IsNot Nothing AndAlso vPicker.PickedList.Count > 0 Then
                                    viewsForChange = vPicker.PickedList
                                Else
                                    acTrans.Abort()
                                    Exit Sub
                                End If

                            End Using

                        End If

                        'pick the layouts to be adjusted
                        Dim loList As SortedDictionary(Of Integer, String) = LayoutTabList()
                        Dim layoutLst As List(Of String)

                        Using lPicker As New LayoutPicker
                            'add layout names to the picker form
                            For Each lN As String In loList.Values
                                Dim lName As String = lN
                                If Not lName = "Model" Then
                                    lPicker.ListBox1.Items.Add(lName)
                                End If
                            Next

                            lPicker.PickerLabel.Text = "Select layouts to update:"

                            'create a list variable and show the form
                            lPicker.ShowDialog()

                            If lPicker.DialogResult = DialogResult.Cancel Then
                                Exit Sub
                            Else
                                layoutLst = lPicker.PickedList
                            End If
                        End Using

                        If layoutLst.Count > viewsForChange.Count Then
                            ed.WriteMessage(vbLf & "Error.  Number of views must be equal or greater than the number of layouts to change.")
                            acTrans.Dispose()
                            Exit Sub
                        End If

                        Dim n As Integer = i
                        Dim numberedDic As New Dictionary(Of String, Integer)

                        If useNumbViews Then
                            For x As Integer = 0 To layoutLst.Count - 1
                                viewDic.Add(layoutLst(x), n)
                                n += 1
                            Next
                        Else
                            For x As Integer = 0 To layoutLst.Count - 1
                                viewDic.Add(layoutLst(x), viewsForChange(x))
                            Next
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
                            For x As Integer = 0 To layoutLst.Count - 1
                                viewDic(x) = viewsForChange(x)
                            Next

                            'if there is a current plot settings validator, use it
                            psetval = PlotSettingsValidator.Current
                            pset = plsets.GetAt(mySetup).GetObject(OpenMode.ForWrite)
                            psetval.RefreshLists(pset)
                            myPltr = pset.PlotConfigurationName
                        End If


                        'Dim stdScale As StandardScaleType = pset.StdScaleType
                        'Dim custScale As CustomScale = pset.CustomPrintScale
                        'Dim tempScaleStr() As String

                        'If stdScale = StandardScaleType.CustomScale Then
                        '    ReDim tempScaleStr(1)
                        '    tempScaleStr(0) = custScale.Numerator & ":" & custScale.Denominator
                        'Else
                        '    tempScaleStr = [Enum].GetNames(stdScale.GetType)
                        'End If

                        'Dim msg As String = vbLf & "Layout scale = " & tempScaleStr(0) & ".  Do you want to change it?"
                        'Dim yn As Boolean = YesNoQuery(msg)
                        'Dim numerat As Integer
                        'Dim denom As Integer
                        'Dim bailout As Boolean
                        'Dim scl As CustomScale

                        'If yn = True Then

                        '    Dim sp As New FontPicker(PickerType.PlotScale)
                        '    sp.ShowDialog()

                        '    If sp.DialogResult = DialogResult.OK Then

                        '        If sp.StandardSclStr = "" Then
                        '            Dim pio As New PromptIntegerOptions(vbLf & "Enter the number of modelspace units:")

                        '            Dim pir As PromptIntegerResult = ed.GetInteger(pio)
                        '            If pir.Status = PromptStatus.OK Then
                        '                numerat = pir.Value
                        '                bailout = False
                        '            Else
                        '                bailout = True
                        '            End If

                        '            If Not bailout Then
                        '                Dim newMsg As String = vbLf & "Enter the number of paperspace units:"
                        '                pio.Message = newMsg
                        '                pir = ed.GetInteger(pio)

                        '                If pir.Status = PromptStatus.OK Then
                        '                    denom = pir.Value
                        '                    bailout = False
                        '                Else
                        '                    bailout = True
                        '                End If
                        '            End If
                        '        End If
                        '    End If
                        'Else
                        '    bailout = True
                        'End If

                        'If Not bailout Then
                        '    scl = New CustomScale(numerat, denom)
                        'End If

                        'create the viewport on the viewports layer
                        Dim vpLayer As String

                        If LayerExists("Vports") Or LayerExists("vports") Or LayerExists("VPORTS") Then
                            vpLayer = "Vports"
                        Else
                            vpLayer = AddNewLayer("Vports", 3)
                        End If

                        'Get the sheet name to apply to the layout
                        Dim mySheet As String = GetSheetName(psetval, pset)
                        If mySheet = "" Then Exit Sub
                        'add the plot settings, printer, and sheet name to the plot settings validator
                        psetval.SetPlotConfigurationName(pset, myPltr, mySheet)

                        'run through the layout list
                        If layoutLst.Count > 0 Then
                            For q As Integer = 0 To layoutLst.Count - 1
                                'For Each loName As String In layoutLst
                                Dim loName As String = layoutLst(q)
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
                                'If Not bailout Then psetval.SetCustomPrintScale(pset, scl)

                                'i += 1
                                'if there is a numbered view corresponding to the layout view, use it
                                If useNumbViews Then
                                    If vtb.Has(i.ToString) Then
                                        Using vtr As ViewTableRecord = TryCast(acTrans.GetObject(vtb(i.ToString), OpenMode.ForWrite), ViewTableRecord)
                                            If vtb.Has(loName) Then
                                                Dim tempname As String = GetValidViewName(loName)
                                                Using tempVTR As ViewTableRecord = acTrans.GetObject(vtb(loName), OpenMode.ForWrite)
                                                    tempVTR.Name = tempname
                                                    vtr.Name = loName
                                                    For Each ky As String In viewDic.Keys
                                                        Dim tname As String = viewDic(ky)
                                                        If tname = loName Then
                                                            viewDic(ky) = tempname
                                                            Exit For
                                                        End If
                                                    Next
                                                End Using
                                            Else
                                                vtr.Name = loName
                                            End If
                                            ed.SwitchToModelSpace()
                                            ed.SetCurrentView(vtr)
                                            ed.SwitchToPaperSpace()
                                        End Using
                                    End If
                                    i += 1
                                Else
                                    If vtb.Has(viewDic(loName)) Then
                                        Using vtr As ViewTableRecord = acTrans.GetObject(vtb(viewDic(loName)), OpenMode.ForWrite)
                                            If vtb.Has(loName) Then
                                                Dim tempname As String = GetValidViewName(loName)
                                                Using tempVTR As ViewTableRecord = acTrans.GetObject(vtb(loName), OpenMode.ForWrite)
                                                    tempVTR.Name = tempname
                                                    vtr.Name = loName
                                                    For Each ky As String In viewDic.Keys
                                                        Dim tname As String = viewDic(ky)
                                                        If tname = loName Then
                                                            viewDic(ky) = tempname
                                                            Exit For
                                                        End If
                                                    Next
                                                End Using
                                            Else
                                                vtr.Name = loName
                                            End If
                                            ed.SwitchToModelSpace()
                                            ed.SetCurrentView(vtr)
                                            ed.SwitchToPaperSpace()
                                        End Using
                                    End If
                                End If
                            Next
                        End If
                    End Using
                    acTrans.Commit()
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        Public Function GetValidViewName(tempName As String)
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim vtb As ViewTable = acTrans.GetObject(dwgDB.ViewTableId, OpenMode.ForRead)
                Dim newName As String = tempName
                If vtb.Has(newName) Then
                    Do
                        If Left(newName, 4) = "temp" Then
                            Dim numb As Integer = CInt(Right(newName, Len(newName) - 4))
                            newName = "temp" & (numb + 1).ToString
                        Else
                            newName = "temp0"
                        End If
                    Loop Until vtb.Has(newName) = False
                End If

                If Not newName = "" Then
                    Return newName
                Else
                    Return ""
                End If

                acTrans.Commit()

            End Using

        End Function


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
                                            pName = " " & prop.Value.ToString
                                        ElseIf prop.PropertyName = "Flip state1" Then
                                            fState = " " & prop.Value.ToString
                                        End If
                                    Next
                                    blkName = blkBTR.Name & pName & fState
                                Else
                                    blkName = bRef.Name.ToString
                                End If
                                Dim bNameChars() As Char = blkName.ToCharArray
                                blkName = ReturnValidFileName(bNameChars)
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

        Friend m_myClr As Autodesk.AutoCAD.Colors.Color



        Public Sub AuditDwg(curDwgName As String)
            Dim acDwgMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim curDwg As Document = acDwgMgr.Open(curDwgName, False)

            Using docLock As DocumentLock = curDwg.LockDocument
                Dim dwgdb As Database = curDwg.Database
                Dim auditVar As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("AUDITCTL")
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("AUDITCTL", 1)

                Try
                    Using actrans As Transaction = dwgdb.TransactionManager.StartTransaction
                        DatabaseExtension.Audit(dwgdb, True, True)
                        actrans.Commit()
                        'curDwg.CloseAndSave(curDwg.Name)
                    End Using

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                Finally
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("AUDITCTL", auditVar)
                    curDwg.CloseAndSave(curDwg.Name)
                End Try
            End Using

        End Sub


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
                Dim ltTbl As LinetypeTable = TryCast(acTrans.GetObject(dwgDB.LinetypeTableId, OpenMode.ForWrite), LinetypeTable)
                Dim blkTbl As BlockTable = TryCast(acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim mdlSpace As BlockTableRecord = TryCast(acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                For Each obID As ObjectId In ltTbl
                    Using ltRec As LinetypeTableRecord = TryCast(acTrans.GetObject(obID, OpenMode.ForRead), LinetypeTableRecord)
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
                                        shp.Dispose()

                                    ElseIf ltype.HasText Then
                                        Dim textStr As String = ltype.TextString
                                        Dim textstyleID As ObjectId = ltype.TextStyleID
                                        Dim textStyleName As String = ltype.TextStyleName
                                        sb2.Append("[" & Chr(34) & textStr & Chr(34) & "," & textStyleName & ",S=" & ltype.ShapeScale.ToString)

                                        If ltype.HasUCSOrient Then
                                            sb2.Append("," & "A=" & ltype.ShapeRotation.ToString)
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
            Dim pso As New PromptSelectionOptions() With {.MessageForAdding = vbLf & "Select annotative objects"}
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


        <CommandMethod("NEWTEXTMASK", CommandFlags.UsePickSet)>
        Public Sub NewTextMask()

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDb As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim SelResult As PromptSelectionResult = ed.SelectImplied()

            If SelResult.Status = PromptStatus.Error Then
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select dtext entities to mask:")}
                SelResult = ed.GetSelection(Seloptions)
            Else
                ed.SetImpliedSelection(New ObjectId(-1) {})
            End If

            If SelResult.Status = PromptStatus.OK Then

                Dim acSSet As SelectionSet = SelResult.Value
                Dim MyobjIDs As ObjectId() = acSSet.GetObjectIds

                Dim ofDist As Double = 0.05

                Dim pdo As New PromptDistanceOptions(vbLf & "Enter or pick the distance offset distance for mask:")
                With pdo
                    .AllowNegative = False
                    .DefaultValue = ofDist
                    .AllowArbitraryInput = True
                    .Only2d = True
                End With

                Dim pdr As PromptDoubleResult = ed.GetDistance(pdo)
                If pdr.Status = PromptStatus.OK Then
                    ofDist = pdr.Value
                Else
                    Exit Sub
                End If

                Dim bgClr As Autodesk.AutoCAD.Colors.Color
                If m_myClr IsNot Nothing Then
                    bgClr = m_myClr
                Else
                    ed.WriteMessage(vbLf & "Select color for background mask:")
                    bgClr = PickColor()
                    If bgClr Is Nothing Then bgClr = Color.FromRgb(255, 255, 255)
                    m_myClr = bgClr
                End If

                Using actrans As Transaction = dwgDb.TransactionManager.StartTransaction()
                    Dim blkTbl As BlockTable = actrans.GetObject(dwgDb.BlockTableId, OpenMode.ForRead)
                    Dim curSpace As BlockTableRecord = actrans.GetObject(dwgDb.CurrentSpaceId, OpenMode.ForWrite)
                    Dim gd As DBDictionary = CType(actrans.GetObject(dwgDb.GroupDictionaryId, OpenMode.ForRead), DBDictionary)

                    For Each objID As ObjectId In MyobjIDs
                        Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForRead)
                        If TypeOf dbObj Is DBText Then
                            Dim textEnt As DBText = CType(dbObj, DBText)
                            Dim textRot As Double = textEnt.Rotation
                            Dim textinspt As Point3d = textEnt.Position

                            textEnt.UpgradeOpen()
                            textEnt.Rotation = 0
                            Dim textext As Extents3d = textEnt.GeometricExtents

                            Dim llPt As New Point3d(textext.MinPoint.X - ofDist, textext.MinPoint.Y - ofDist, 0)
                            Dim urPt As New Point3d(textext.MaxPoint.X + ofDist, textext.MaxPoint.Y + ofDist, 0)
                            Dim ulPt As New Point3d(llPt.X, urPt.Y, 0)
                            Dim lrPt As New Point3d(urPt.X, llPt.Y, 0)
                            Dim sld As New Solid(llPt, lrPt, ulPt, urPt)

                            'Dim lLeft As New DBPoint(llPt)
                            'Dim lright As New DBPoint(lrPt)
                            'Dim uLeft As New DBPoint(ulPt)
                            'Dim uRight As New DBPoint(urPt)

                            'curSpace.AppendEntity(lLeft)
                            'actrans.AddNewlyCreatedDBObject(lLeft, True)
                            'curSpace.AppendEntity(uLeft)
                            'actrans.AddNewlyCreatedDBObject(uLeft, True)
                            'curSpace.AppendEntity(uRight)
                            'actrans.AddNewlyCreatedDBObject(uRight, True)
                            'curSpace.AppendEntity(lright)
                            'actrans.AddNewlyCreatedDBObject(lright, True)

                            Dim sldObjId As ObjectId = curSpace.AppendEntity(sld)
                            actrans.AddNewlyCreatedDBObject(sld, True)

                            sld.TransformBy(Matrix3d.Rotation(textRot, Vector3d.ZAxis, textinspt))
                            If bgClr IsNot Nothing Then sld.Color = bgClr
                            'textEnt.Rotation = textRot
                            textEnt.TransformBy(Matrix3d.Rotation(textRot, Vector3d.ZAxis, textinspt))

                            Dim sldColl As New ObjectIdCollection
                            sldColl.Add(sldObjId)

                            Dim orderDic As DrawOrderTable = actrans.GetObject(curSpace.DrawOrderTableId, OpenMode.ForWrite)
                            orderDic.MoveBelow(sldColl, objID)

                            Dim grp As New Group()

                            gd.UpgradeOpen()
                            Dim grpId As ObjectId = gd.SetAt("*", grp)
                            actrans.AddNewlyCreatedDBObject(grp, True)

                            sldColl.Add(objID)

                            grp.InsertAt(0, sldColl)
                            grp.Selectable = True
                        ElseIf TypeOf dbObj Is MText Then
                            MessageBox.Show("Mtext entities cannot be masked with this command.  Use built-in autocad methods for MTEXT instead.")
                        End If
                    Next
                    actrans.Commit()
                End Using
            End If
        End Sub

        <CommandMethod("SETMASKCOLOR")>
        Public Sub SetMaskColor()

            Dim cd As New Autodesk.AutoCAD.Windows.ColorDialog()
            Dim cr As System.Windows.Forms.DialogResult = cd.ShowDialog

            If cr = DialogResult.OK Then
                Dim clr As Autodesk.AutoCAD.Colors.Color = cd.Color
                m_myClr = clr
            End If

        End Sub

        <CommandMethod("LCTA")>
        Public Sub ListCommandsFromThisAssembly()
            Dim dm As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim ed As Editor = dm.MdiActiveDocument.Editor
            Dim asm As Assembly = Assembly.GetExecutingAssembly()
            Dim cmds As String() = GetCommands(asm, False)

            For Each cmd As String In cmds
                ed.WriteMessage(cmd & vbLf)
            Next
        End Sub

        <CommandMethod("LFTA")>
        Public Sub ListfunctionsFromThisAssembly()
            Dim dm As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim ed As Editor = dm.MdiActiveDocument.Editor
            Dim asm As Assembly = Assembly.GetExecutingAssembly()
            Dim cmds As String() = GetFunctions(asm)
            For Each cmd As String In cmds
                ed.WriteMessage(cmd & vbLf)
            Next
        End Sub

        <CommandMethod("LANC")>
        Public Sub ListAllNetCommands()
            'Dim cmds As StringCollection = New StringCollection()
            Dim dm As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim ed As Editor = dm.MdiActiveDocument.Editor
            Dim asms As Assembly() = AppDomain.CurrentDomain.GetAssemblies()

            For Each asm As Assembly In asms
                If Not asm.FullName.Contains("Microsoft.Expression.Interactions") Then
                    'If asm.FullName.Contains("PresentationFramework") Then GoTo Skipit
                    Try
                        Dim cmds() As String = GetCommands(asm, False)
                        If cmds.Length > 0 Then ed.WriteMessage(vbLf & vbLf & asm.FullName.ToString)
                        For Each cmd As String In cmds
                            ed.WriteMessage(cmd & vbLf)
                        Next
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString & vbLf & ex.ErrorStatus.ToString)
                        Return
                    End Try
                End If
Skipit:
            Next

        End Sub


        '<CommandMethod("LC")>

        <CommandMethod("MUTCDCOLORS")>
        Public Sub ListMUTCDcolors()

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor

            Try
                Dim pkwOpts As New PromptKeywordOptions(vbCrLf & "Enter the MUTCD Color to translate: ")

                With pkwOpts
                    .Keywords.Add("Red")
                    .Keywords.Add("Yellow")
                    .Keywords.Add("Green")
                    .Keywords.Add("Blue")
                    .Keywords.Add("Brown")
                    .Keywords.Add("Orange")
                    .Keywords.Add("Purple")
                    .Keywords.Add("Pink")
                    .Keywords.Add("YellowGreen")
                    .Keywords.Add("RedPavement")
                    .Keywords.Add("GreenPavement")
                    .AppendKeywordsToMessage = True
                End With

                Dim pkwRes As PromptResult = ed.GetKeywords(pkwOpts)
                Dim clrString As String = ""

                If pkwRes.Status = PromptStatus.OK Then
                    clrString = pkwRes.StringResult
                Else
                    Exit Sub
                End If

                If String.IsNullOrEmpty(clrString) Then Exit Sub

                Select Case clrString
                    Case Is = "Red"
                        ed.WriteMessage(vbLf & "Pantone:1805C; RGB:175,39,47; ACI:242")
                    Case Is = "Yellow"
                        ed.WriteMessage(vbLf & "Pantone:122C;  RGB:254,209,65;  ACI:53")
                    Case Is = "Green"
                        ed.WriteMessage(vbLf & "Pantone:342C;  RGB:0,103,71;  ACI:116")
                    Case Is = "Blue"
                        ed.WriteMessage(vbLf & "Pantone:294C;  RGB:0,47,108;  ACI:156")
                    Case Is = "Orange"
                        ed.WriteMessage(vbLf & "Pantone:152C;  RGB:229,114,0;  ACI:30")
                    Case Is = "YellowGreen"
                        ed.WriteMessage(vbLf & "Pantone:382C;  RGB:196,214,0;  ACI:52")
                    Case Is = "Purple"
                        ed.WriteMessage(vbLf & "Pantone:519C;  RGB:89,49,95;  ACI:219")
                    Case Is = "Brown"
                        ed.WriteMessage(vbLf & "Pantone:469C;  RGB:105,63,35;  ACI:27")
                    Case Is = "Pink"
                        ed.WriteMessage(vbLf & "Pantone:198C;  RGB:223,70,97;  ACI:13")
                    Case Is = "GreenPavement"
                        ed.WriteMessage(vbLf & "Pantone:802C;  RGB:68,214,44;  ACI:82")
                    Case Is = "RedPavement"
                        ed.WriteMessage(vbLf & "Pantone:485C;  RGB:218,41,28;  ACI:22")
                End Select

                Exit Sub

            Catch ex As Exception
                ed.WriteMessage("Fatal Error.")
            End Try

        End Sub

        <CommandMethod("CLRTOENT", CommandFlags.UsePickSet)>
        Public Sub AssignColorToEntity()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDb As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim SelResult As PromptSelectionResult = ed.SelectImplied()

            If SelResult.Status = PromptStatus.Error Then
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select entities to change their color from ByLayer to direct assignment.")}
                SelResult = ed.GetSelection(Seloptions)
            Else
                ed.SetImpliedSelection(New ObjectId(-1) {})
            End If

            If SelResult.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = SelResult.Value
                Dim MyobjIDs As ObjectId() = acSSet.GetObjectIds
                Try

                    Using actrans As Transaction = dwgDb.TransactionManager.StartTransaction()
                        Dim lyrTbl As LayerTable = actrans.GetObject(dwgDb.LayerTableId, OpenMode.ForRead)

                        For Each objID As ObjectId In MyobjIDs
                            Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForRead)
                            If TypeOf dbObj Is Entity Then
                                Dim ent As Entity = TryCast(dbObj, Entity)
                                If ent Is Nothing Then Continue For
                                If Not ent.IsWriteEnabled Then ent.UpgradeOpen()
                                Dim lyrId As ObjectId = ent.LayerId
                                Dim ltr As LayerTableRecord = actrans.GetObject(lyrId, OpenMode.ForRead)
                                Dim clr As Color = ltr.Color
                                ent.Color = clr
                            End If
                        Next
                        actrans.Commit()
                    End Using

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

            End If

        End Sub

        <CommandMethod("CTARGET", CommandFlags.UsePickSet)>
        Public Sub MatchColorTarget()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDb As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor
            Dim SelResult As PromptSelectionResult = ed.SelectImplied()

            If SelResult.Status = PromptStatus.Error Then
                Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select target entities to change color:")}
                SelResult = ed.GetSelection(Seloptions)
            Else
                ed.SetImpliedSelection(New ObjectId(-1) {})
            End If

            Dim peo2 As New PromptEntityOptions(vbLf & "Select source entity to copy color from:")
            With peo2
                .SetRejectMessage(vbLf & "Object must be an AutoCAD entity")
                .AddAllowedClass(GetType(Entity), False)
                .AllowNone = False
            End With

            Dim per2 As PromptEntityResult = ed.GetEntity(peo2)

            Dim fromEntId As ObjectId
            If per2.Status = PromptStatus.OK Then
                fromEntId = per2.ObjectId
            Else
                Exit Sub
            End If
            Try
                If SelResult.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = SelResult.Value
                    Dim MyobjIDs As ObjectId() = acSSet.GetObjectIds

                    Using actrans As Transaction = dwgDb.TransactionManager.StartTransaction()
                        Dim frmEnt As Entity = actrans.GetObject(fromEntId, OpenMode.ForRead)

                        For Each toEntId As ObjectId In MyobjIDs
                            Dim toEnt As Entity = actrans.GetObject(toEntId, OpenMode.ForWrite)
                            toEnt.Color = frmEnt.Color
                        Next
                        actrans.Commit()
                    End Using
                Else
                    Exit Sub
                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        <CommandMethod("CUSTOMHELP")>
        Public Sub CustomHelp()
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database
            Dim menuPath As String

            Try
                menuPath = HostApplicationServices.Current.FindFile("CustomNetAssemblyHelp.chm", dwgDB, FindFileHint.[Default])
            Catch ex As Exception
                menuPath = ""
            End Try

            If String.IsNullOrEmpty(menuPath) Then
                MessageBox.Show("Help file Not found.  Select the folder where the menu file is stored")

                Dim fp As String = GetMyFolderName()
                If fp IsNot Nothing Then
                    Dim menuP(1) As String
                    menuP(0) = fp
                    menuP(1) = "CustomNetAssemblyHelp.chm"
                    menuPath = Path.Combine(menuP)
                End If
            End If

            If File.Exists(menuPath) Then Help.ShowHelp(Nothing, menuPath, 0)

        End Sub

        Public Function GetRGFilesDic() As Dictionary(Of String, String)

            Dim myDic As New Dictionary(Of String, String)

            myDic("RG_2000B") = "Roadgeek 2000 Series B.TTF"
            myDic("RG_2000C") = "Roadgeek 2000 Series C.TTF"
            myDic("RG_2000D") = "Roadgeek 2000 Series D.TTF"
            myDic("RG_2000E") = "Roadgeek 2000 Series E.TTF"
            myDic("RG_2000F") = "Roadgeek 2000 Series F.TTF"

            myDic("RG_20051BW") = "Roadgeek2005BlendB1W.ttf"
            myDic("RG_20051B") = "Roadgeek_2005_Series_5.ttf"
            myDic("RG_20052B") = "Roadgeek_2005_Series_6.ttf"
            myDic("RG_20053B") = "Roadgeek_2005_Series.ttf"
            myDic("RG_2005B") = "Roadgeek_2005_Series_0.ttf"
            myDic("RG_2005C") = "Roadgeek_2005_Series_1.ttf"
            myDic("RG_2005D") = "Roadgeek_2005_Series_2.ttf"
            myDic("RG_2005E") = "Roadgeek_2005_Series_3.ttf"
            myDic("RG_2005EM") = "Roadgeek_2005_Series_EM.ttf"
            myDic("RG_2005F") = "Roadgeek_2005_Series_4.ttf"

            Return myDic


        End Function

        Public Function GetRGFontsDic() As Dictionary(Of String, String)

            Dim myDic As New Dictionary(Of String, String)

            myDic("RG_2000B") = "Roadgeek 2000 Series B"
            myDic("RG_2000C") = "Roadgeek 2000 Series C"
            myDic("RG_2000D") = "Roadgeek 2000 Series D"
            myDic("RG_2000E") = "Roadgeek 2000 Series E"
            myDic("RG_2000F") = "Roadgeek 2000 Series F"
            myDic("RG_20051BW") = "Roadgeek 2005 Blend B1W"
            myDic("RG_20051B") = "Roadgeek 2005 Series 1B"
            myDic("RG_20052B") = "Roadgeek 2005 Series 2B"
            myDic("RG_20053B") = "Roadgeek 2005 Series 3B"
            myDic("RG_2005B") = "Roadgeek 2005 Series B"
            myDic("RG_2005C") = "Roadgeek 2005 Series C"
            myDic("RG_2005D") = "Roadgeek 2005 Series D"
            myDic("RG_2005E") = "Roadgeek 2005 Series E"
            myDic("RG_2005F") = "Roadgeek 2005 Series F"
            myDic("RG_2005EM") = "Roadgeek 2005 Series EM"

            Return myDic

        End Function

        Public Function RGtranslateDic() As Dictionary(Of String, String)
            Dim myDic As New Dictionary(Of String, String)

            myDic("Series B.ttf") = "Series B"
            myDic("Series C.ttf") = "Series C"
            myDic("Series D.ttf") = "Series D"
            myDic("Series E.ttf") = "Series E"
            myDic("Series F.ttf") = "Series F"

            myDic("Series_0.ttf") = "Series B"
            myDic("Series_1.ttf") = "Series C"
            myDic("Series_2.ttf") = "Series D"
            myDic("Series_3.ttf") = "Series E"
            myDic("Series_4.ttf") = "Series F"
            myDic("Series_5.ttf") = "Series 1B"
            myDic("Series_6.ttf") = "Series 2B"
            myDic("Series.ttf") = "Series 3B"
            myDic("Series_EM.ttf") = "Series EM"

            Return myDic

        End Function



        Public Enum EHelp_MasterCustomLIbraryHelp

            HELP_CommandList = 0
            HELP_ArcCommands = 2
            HELP_LISTARCDATA = 3
            HELP_BlockCommands = 4
            HELP_LABLKCUST = 5
            HELP_LABLKS = 6
            HELP_CHBLKCOLOR = 7
            HELP_RENBLKS = 8
            HELP_CBTZ = 9
            HELP_CBBL = 10
            HELP_CPSBL = 11
            HELP_SetDwgsBase = 12
            HELP_BLKDATA = 13
            HELP_INSALL = 14
            HELP_THUMBS = 15
            HELP_PATF = 16
            HELP_WBTF = 17
            HELP_CBU = 18
            HELP_CPMLT = 19
            HELP_CPM = 20
            HELP_CPMW = 21
            HELP_SealSig = 22
            HELP_GeometryCommands = 23
            HELP_MKST = 24
            HELP_MKSTE = 25
            HELP_MKSTAR = 26
            HELP_CBYA = 27
            HELP_SBYA = 28
            HELP_ETAN = 29
            HELP_ITAN = 30
            HELP_TANPT = 31
            HELP_TESTVECTS = 32
            HELP_RgnCentroid = 33
            HELP_LV = 34
            HELP_LayerCommands = 35
            HELP_FRZVPLAY = 36
            HELP_LAOFF = 37
            HELP_LAON = 38
            HELP_ExpLayers = 39
            HELP_ImpLayers = 40
            HELP_EML = 41
            HELP_MDL = 42
            HELP_PlotLayoutCommands = 43
            HELP_ListStyleTables = 44
            HELP_ListPstyles = 45
            HELP_ShowPstyles = 46
            HELP_IMVS = 47
            HELP_RLO = 48
            HELP_ALTS = 49
            HELP_MSPT = 50
            HELP_MiscCommands = 51
            HELP_LLTS = 52
            HELP_LLAYS = 53
            HELP_LSTS = 54
            HELP_EXLTS = 55
            HELP_MAH = 56
            HELP_MDAR = 57
            HELP_RABCS = 58
            HELP_LCTA = 59
            HELP_LANC = 60
            HELP_MUTCDCOLORS = 61
            HELP_Wallsvb = 62
            HELP_WALLSTPS = 63
            HELP_WALLPROFILES = 64
            HELP_FTLO = 65
            HELP_FTBOT = 66
            HELP_WallLine = 67
        End Enum



    End Module



End Namespace

