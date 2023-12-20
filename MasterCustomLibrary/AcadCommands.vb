Imports System.IO
Imports System.Math
Imports System.Text
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.PlottingServices
Imports System.Xml
Imports Autodesk.AutoCAD.Colors
Imports System.Xml.Serialization
Imports System.Runtime.CompilerServices
Imports GI = Autodesk.AutoCAD.GraphicsInterface

Public Module ArcCommands

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

        ed.WriteMessage(vbLf & sb.ToString & vbLf)

    End Sub

End Module


Public Module BlockCommands

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
                    If TypeOf dbObj Is BlockReference Then
                        Dim parentBref As BlockReference = TryCast(dbObj, BlockReference)
                        Dim parentName As String = parentBref.Name
                        Dim ParentObjID As ObjectId = blkTbl(parentName)
                        Dim ParentBTR As BlockTableRecord = actrans.GetObject(ParentObjID, OpenMode.ForWrite)
                        For Each obID As ObjectId In ParentBTR
                            Dim dObj As DBObject = actrans.GetObject(obID, OpenMode.ForRead)
                            If TypeOf dObj Is Entity Then
                                Dim ent As Entity = CType(dObj, Entity)
                                If Not ent.LayerId = dwgDB.LayerZero Then
                                    ent.UpgradeOpen()
                                    'Dim entlay As String = ent.Layer
                                    Dim entlayID As ObjectId = ent.LayerId
                                    Dim layTblRec As LayerTableRecord = actrans.GetObject(entlayID, OpenMode.ForRead)
                                    If ent.Color.IsByLayer Then ent.Color = layTblRec.Color
                                    If ent.PlotStyleName = "ByLayer" Then ent.PlotStyleName = layTblRec.PlotStyleName
                                    If ent.Linetype = "ByLayer" Then ent.LinetypeId = layTblRec.LinetypeObjectId
                                    ent.LayerId = dwgDB.LayerZero
                                Else
                                End If
                            End If
                        Next
                    End If
                Next
                actrans.Commit()
            End Using
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

        Dim bvp As New BlockViewPanel
        bvp.ShowDialog()
        bvp.Dispose()
    End Sub


    <CommandMethod("INSALL")>
    Public Sub InsertAllDwgFilesInFolder()

        Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = acDwg.Database
        Dim ed As Editor = acDwg.Editor

        ed.WriteMessage(vbLf & "This command will insert all blocks in a designated folder.")

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
        Else
            ed.WriteMessage(vbLf & "Command Canceled.")
            Exit Sub
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
        Dim insptX As Double
        Dim insPt As Point3d

        Dim i As Integer = 0
        Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction

            Dim blkTbl As BlockTable = actrans.GetObject(dwgDB.BlockTableId, OpenMode.ForWrite)
            Dim cSpace As BlockTableRecord = actrans.GetObject(dwgDB.CurrentSpaceId, OpenMode.ForWrite)

            For Each ky As String In pathList.Keys

                Try
                    Dim btrID As ObjectId = InsertDwg(pathList(ky), ky)

                    If i = 10 Then
                        insPty -= yDist
                        insptX = initialInsPtx
                        i = 1
                    End If

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

                    insptX += xDist
                    i += 1
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error in InsertAllDwgFilesInFolder function.")
                    Exit Sub
                End Try
            Next
            actrans.Commit()
        End Using

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
                .BxList.SelectionMode = Windows.Forms.SelectionMode.MultiExtended
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
                        If InStr(parentName, "Existing") > 0 Or InStr(parentName, "Proposed") > 0 Or InStr(parentName, "Remove") > 0 Then
                            Dim nameParts() As String = Split(parentName, "_")
                            If nameParts.Count > 2 Then
                                Dim tempName(nameParts.Count - 1) As String
                                For j = 0 To nameParts.Count - 2
                                    ReDim Preserve tempName(j + 1)
                                    tempName(j) = nameParts(j)
                                Next
                                baseName = Join(tempName, "_")
                                curState = nameParts(nameParts.Count - 1)
                            ElseIf nameParts.Count = 2 Then
                                Dim tempname() As String = Split("_")
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

                        newName = (baseName & "_" & newState)
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
                            parentBref.UpgradeOpen()
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
                                'If ChangeVisState(subID, newState, "Visibility1") = False Then
                                '            ed.WriteMessage(vbLf & "Error in parent block defnition.  Block reference unchanged.")
                                '            actrans.Abort()
                                '            Exit Sub
                                '        End If
                                '    End If
                                'Next
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

                    Dim pso2 As New PromptStringOptions(vbLf & "Input alternate block name to be created.")
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

End Module

Public Module GeometryCommands

    Public Sub ExtTangent2Circles()
        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        Try
            Dim peo1 As New PromptEntityOptions(vbLf & "Select first circle or arc")

            With peo1
                .SetRejectMessage(vbLf & "Must pick a circle or arc.")
                .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Circle), True)
                .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Arc), True)
            End With

            Dim dbObId As ObjectId

            Dim pr1 As PromptEntityResult = ed.GetEntity(peo1)
            If pr1.Status = PromptStatus.OK Then
                dbObId = pr1.ObjectId
            End If

            Dim peo2 As New PromptEntityOptions(vbLf & "Select second circle or arc")

            With peo2
                .SetRejectMessage(vbLf & "Must pick a circle or arc.")
                .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Circle), True)
                .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Arc), True)
            End With

            Dim dbObId2 As ObjectId

            Dim pr2 As PromptEntityResult = ed.GetEntity(peo2)
            If pr2.Status = PromptStatus.OK Then
                dbObId2 = pr2.ObjectId
            End If

            Dim tc1 As Circle
            Dim tc2 As Circle

            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim dbObj1 As DBObject = acTrans.GetObject(dbObId, OpenMode.ForRead)
                If TypeOf dbObj1 Is Autodesk.AutoCAD.DatabaseServices.Circle Then
                    tc1 = CType(dbObj1, Circle)
                ElseIf TypeOf dbObj1 Is Arc Then
                    Dim a1 As Arc = CType(dbObj1, Arc)
                    tc1 = New Circle(a1.Center, Vector3d.ZAxis, a1.Radius)
                    'a1.Dispose()
                Else
                    acTrans.Abort()
                    Exit Sub
                End If

                Dim dbObj2 As DBObject = acTrans.GetObject(dbObId2, OpenMode.ForRead)
                If TypeOf dbObj2 Is Autodesk.AutoCAD.DatabaseServices.Circle Then
                    tc2 = CType(dbObj2, Circle)
                ElseIf TypeOf dbObj2 Is Arc Then
                    Dim a2 As Arc = CType(dbObj2, Arc)
                    tc2 = New Circle(a2.Center, Vector3d.ZAxis, a2.Radius)
                    'a2.Dispose()
                Else
                    Exit Sub
                End If

                Dim cen1 As New Point3d(tc1.Center.X, tc1.Center.Y, 0)
                Dim cen2 As New Point3d(tc2.Center.X, tc2.Center.Y, 0)
                Dim c1 As New Circle(cen1, Vector3d.ZAxis, tc1.Radius)
                Dim c2 As New Circle(cen2, Vector3d.ZAxis, tc2.Radius)

                If c1.Radius > c2.Radius Then
                    Dim tempC As Circle = c1
                    c1 = c2
                    c2 = tempC
                End If

                Dim raddiff = c2.Radius - c1.Radius

                Dim c3 As New Circle(c2.Center, Vector3d.ZAxis, raddiff)

                Dim ctLine As New Line(c2.Center, c1.Center)
                Dim halfline As Double = ctLine.Length / 2
                'Dim quarterline As Double = ctLine.Length / 4
                Dim midpt As Point3d = ctLine.GetPointAtDist(halfline)
                'Dim pointS As Point3d = ctLine.GetPointAtDist(raddiff)
                'Dim c3 As New Circle(c2.Center, Vector3d.ZAxis, c1.Radius)
                Dim c4 As New Circle(midpt, Vector3d.ZAxis, halfline)

                Dim pointJ As New Point3dCollection
                c4.IntersectWith(c3, Intersect.OnBothOperands, pointJ, IntPtr.Zero, IntPtr.Zero)

                If pointJ.Count = 0 Then Exit Sub

                Dim line1 As New Line(c2.Center, pointJ(0))
                Dim line2 As New Line(c2.Center, pointJ(1))

                Dim pointL1 As New Point3dCollection
                line1.IntersectWith(c2, Intersect.ExtendThis, pointL1, IntPtr.Zero, IntPtr.Zero)

                If pointL1.Count = 0 Then Exit Sub
                Dim testpt1 As Point3d = pointL1(0)
                Dim testpt2 As Point3d = pointL1(1)

                Dim pTanC2A As Point3d

                If c1.Center.GetVectorTo(testpt1).Length < c1.Center.GetVectorTo(testpt2).Length Then
                    pTanC2A = testpt1
                Else
                    pTanC2A = testpt2
                End If

                Dim pointL2 As New Point3dCollection

                line2.IntersectWith(c2, Intersect.ExtendThis, pointL2, IntPtr.Zero, IntPtr.Zero)

                If pointL2.Count = 0 Then Exit Sub
                testpt1 = pointL2(0)
                testpt2 = pointL2(1)

                Dim pTanC2B As Point3d

                If c1.Center.GetVectorTo(testpt1).Length < c1.Center.GetVectorTo(testpt2).Length Then
                    pTanC2B = testpt1
                Else
                    pTanC2B = testpt2
                End If

                Dim clPtS As Point3d = ctLine.GetPointAtDist(raddiff)
                Dim distSJ As Double = clPtS.GetVectorTo(pointJ(0)).Length

                Dim C5 As New Circle(c1.Center, Vector3d.ZAxis, raddiff)

                Dim ctrlineC5 As New Point3dCollection
                ctLine.IntersectWith(C5, Intersect.ExtendThis, ctrlineC5, IntPtr.Zero, IntPtr.Zero)

                testpt1 = ctrlineC5(0)
                testpt2 = ctrlineC5(1)

                Dim ptZ As New Point3d

                If c2.Center.GetVectorTo(testpt1).Length < c2.Center.GetVectorTo(testpt2).Length Then
                    ptZ = testpt2
                Else
                    ptZ = testpt1
                End If

                Dim c6 As New Circle(ptZ, Vector3d.ZAxis, distSJ)

                Dim pointT As New Point3dCollection

                c6.IntersectWith(C5, Intersect.OnBothOperands, pointT, IntPtr.Zero, IntPtr.Zero)

                If pointT.Count = 0 Then Exit Sub
                Dim line3 As New Line(c1.Center, pointT(0))
                Dim line4 As New Line(c1.Center, pointT(1))

                Dim c1L3 As New Point3dCollection
                line3.IntersectWith(c1, Intersect.OnBothOperands, c1L3, IntPtr.Zero, IntPtr.Zero)

                If c1L3.Count = 0 Then Exit Sub
                testpt1 = c1L3(0)
                'testpt2 = c1L3(1)

                Dim pTanC1A As New Point3d
                pTanC1A = New Point3d(testpt1.X, testpt1.Y, 0)

                If c2.Center.GetVectorTo(testpt1).Length < c2.Center.GetVectorTo(testpt2).Length Then
                    pTanC1A = testpt2
                Else
                    pTanC1A = testpt1
                End If

                Dim c1L4 As New Point3dCollection
                line4.IntersectWith(c1, Intersect.OnBothOperands, c1L4, IntPtr.Zero, IntPtr.Zero)

                If c1L4.Count = 0 Then Exit Sub
                testpt1 = c1L4(0)
                'testpt2 = c1L4(1)



                Dim pTanC1B As New Point3d
                pTanC1B = New Point3d(testpt1.X, testpt1.Y, 0)

                'If c2.Center.GetVectorTo(testpt1).Length < c2.Center.GetVectorTo(testpt2).Length Then
                '    pTanC1B = testpt2
                'Else
                '    pTanC1B = testpt1
                'End If

                Dim test1 As New Line(pTanC1A, pTanC2A)
                Dim test2 As New Line(pTanC1B, pTanC2B)

                Dim finals As New Point3dCollection
                test1.IntersectWith(test2, Intersect.OnBothOperands, finals, IntPtr.Zero, IntPtr.Zero)

                Dim test3 As New Line(pTanC1A, pTanC2B)
                Dim test4 As New Line(pTanC1B, pTanC2A)

                Dim final1 As Line
                Dim final2 As Line

                If finals.Count > 0 Then
                    final1 = New Line(pTanC1A, pTanC2B)
                    final2 = New Line(pTanC1B, pTanC2A)
                Else
                    final1 = test1
                    final2 = test2
                End If

                Dim curspace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)

                curspace.AppendEntity(test1)
                acTrans.AddNewlyCreatedDBObject(test1, True)

                curspace.AppendEntity(test2)
                acTrans.AddNewlyCreatedDBObject(test2, True)

                curspace.AppendEntity(test3)
                acTrans.AddNewlyCreatedDBObject(test3, True)

                curspace.AppendEntity(test4)
                acTrans.AddNewlyCreatedDBObject(test4, True)

                'curspace.AppendEntity(final1)
                'acTrans.AddNewlyCreatedDBObject(final1, True)

                'curspace.AppendEntity(final2)
                'acTrans.AddNewlyCreatedDBObject(final2, True)

                acTrans.Commit()

            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    <CommandMethod("ETAN")>
    Public Sub ETangent2Circles()
        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        'Try
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

        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
            Dim dbObj1 As DBObject = acTrans.GetObject(dbObId, OpenMode.ForRead)
            Dim dbObj2 As DBObject = acTrans.GetObject(dbObId2, OpenMode.ForRead)
            Dim pCol As Point2dCollection = ExtTan2Circles(dbObj1, dbObj2)

            If pCol Is Nothing OrElse pCol.Count < 4 Then
                MessageBox.Show("Radius of entity 1 must be smaller than radius of entity 2.  Cancel and try again.")
                Exit Sub
            End If

            Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
            Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
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
            acTrans.Commit()
        End Using

        'Catch

        'End Try

    End Sub

    <CommandMethod("ITAN")>
    Public Sub IntenalTangent2Circles()
        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        'Try
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

        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
            Dim dbObj1 As DBObject = acTrans.GetObject(dbObId, OpenMode.ForRead)
            Dim dbObj2 As DBObject = acTrans.GetObject(dbObId2, OpenMode.ForRead)
            Dim pCol As Point2dCollection = IntTan2Circles(dbObj1, dbObj2)

            If pCol Is Nothing OrElse pCol.Count < 4 Then
                MessageBox.Show("Radius of entity 1 must be smaller than radius of entity 2.  Cancel and try again.")
                Exit Sub
            End If

            Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
            Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim pline1 As New Polyline
            With pline1
                .AddVertexAt(0, pCol(0), 0, 0, 0)
                .AddVertexAt(1, pCol(3), 0, 0, 0)
            End With
            mdlSpace.AppendEntity(pline1)
            acTrans.AddNewlyCreatedDBObject(pline1, True)

            Dim pline2 As New Polyline
            With pline2
                .AddVertexAt(0, pCol(1), 0, 0, 0)
                .AddVertexAt(1, pCol(2), 0, 0, 0)
            End With
            mdlSpace.AppendEntity(pline2)
            acTrans.AddNewlyCreatedDBObject(pline2, True)
            acTrans.Commit()
        End Using

        'Catch

        'End Try

    End Sub


    <CommandMethod("XTAN")>
    Public Sub IntTangent2Circles()
        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        Try
            Dim peo1 As New PromptEntityOptions(vbLf & "Select first circle or arc")

            With peo1
                .SetRejectMessage(vbLf & "Must pick a circle or arc.")
                .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Circle), True)
                .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Arc), True)
            End With

            Dim dbObId As ObjectId

            Dim pr1 As PromptEntityResult = ed.GetEntity(peo1)
            If pr1.Status = PromptStatus.OK Then
                dbObId = pr1.ObjectId
            End If

            Dim peo2 As New PromptEntityOptions(vbLf & "Select second circle or arc")

            With peo2
                .SetRejectMessage(vbLf & "Must pick a circle or arc.")
                .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Circle), True)
                .AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Arc), True)
            End With

            Dim dbObId2 As ObjectId

            Dim pr2 As PromptEntityResult = ed.GetEntity(peo2)
            If pr2.Status = PromptStatus.OK Then
                dbObId2 = pr2.ObjectId
            End If

            Dim c1 As Circle
            Dim c2 As Circle

            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim dbObj1 As DBObject = acTrans.GetObject(dbObId, OpenMode.ForRead)
                If TypeOf dbObj1 Is Autodesk.AutoCAD.DatabaseServices.Circle Then
                    c1 = CType(dbObj1, Circle)
                ElseIf TypeOf dbObj1 Is Arc Then
                    Dim a1 As Arc = CType(dbObj1, Arc)
                    c1 = New Circle(a1.Center, Vector3d.ZAxis, a1.Radius)
                    a1.Dispose()
                Else
                    If dbObj1 IsNot Nothing Then dbObj1.Dispose()
                    acTrans.Abort()
                    Exit Sub
                End If

                Dim dbObj2 As DBObject = acTrans.GetObject(dbObId2, OpenMode.ForRead)
                If TypeOf dbObj2 Is Autodesk.AutoCAD.DatabaseServices.Circle Then
                    c2 = CType(dbObj2, Circle)
                ElseIf TypeOf dbObj2 Is Arc Then
                    Dim a2 As Arc = CType(dbObj2, Arc)
                    c2 = New Circle(a2.Center, Vector3d.ZAxis, a2.Radius)
                    a2.Dispose()
                Else
                    If dbObj2 IsNot Nothing Then dbObj1.Dispose()
                    acTrans.Abort()
                    Exit Sub
                End If

                If c1.Radius > c2.Radius Then
                    Dim tempC As Circle = c1
                    c1 = c2
                    c2 = tempC
                End If

                Dim radsum = c2.Radius + c1.Radius

                Dim c3 As New Circle(c2.Center, Vector3d.ZAxis, radsum)

                Dim ctLine As New Line(c2.Center, c1.Center)
                Dim halfline As Double = ctLine.Length / 2
                'Dim quarterline As Double = ctLine.Length / 4
                Dim midpt As Point3d = ctLine.GetPointAtDist(halfline)
                'Dim pointS As Point3d = ctLine.GetPointAtDist(quarterline)
                Dim c4 As New Circle(midpt, Vector3d.ZAxis, halfline)
                Dim pointJ As New Point3dCollection

                c4.IntersectWith(c3, Intersect.OnBothOperands, pointJ, IntPtr.Zero, IntPtr.Zero)
                Dim line1 As New Line(c2.Center, pointJ(0))
                Dim line2 As New Line(c2.Center, pointJ(1))

                Dim testpt1 As New Point3d
                Dim testpt2 As New Point3d

                Dim pointL1 As New Point3dCollection
                line1.IntersectWith(c2, Intersect.ExtendThis, pointL1, IntPtr.Zero, IntPtr.Zero)
                testpt1 = pointL1(0)
                testpt2 = pointL1(1)

                Dim pTanC2A As Point3d

                If c1.Center.GetVectorTo(testpt1).Length < c1.Center.GetVectorTo(testpt2).Length Then
                    pTanC2A = testpt1
                Else
                    pTanC2A = testpt2
                End If

                Dim pointL2 As New Point3dCollection
                line2.IntersectWith(c2, Intersect.ExtendThis, pointL2, IntPtr.Zero, IntPtr.Zero)
                testpt1 = pointL2(0)
                testpt2 = pointL2(1)

                Dim pTanC2B As Point3d

                If c1.Center.GetVectorTo(testpt1).Length < c1.Center.GetVectorTo(testpt2).Length Then
                    pTanC2B = testpt1
                Else
                    pTanC2B = testpt2
                End If

                Dim clPtS As Point3d = ctLine.GetPointAtDist(c3.Radius)
                Dim ptSptJ As Double = clPtS.GetVectorTo(pointJ(0)).Length

                Dim C5 As New Circle(c1.Center, Vector3d.ZAxis, radsum)
                Dim ctrlineC5 As New Point3dCollection
                ctLine.IntersectWith(C5, Intersect.OnBothOperands, ctrlineC5, IntPtr.Zero, IntPtr.Zero)

                'testpt1 = ctrlineC5(0)
                'testpt2 = ctrlineC5(1)

                Dim ptZ As Point3d = ctrlineC5(0)

                'If c2.Center.GetVectorTo(testpt1).Length < c2.Center.GetVectorTo(testpt2).Length Then
                '    ptZ = testpt2
                'Else
                '    ptZ = testpt1
                'End If

                Dim c6 As New Circle(ptZ, Vector3d.ZAxis, ptSptJ)

                Dim pointT As New Point3dCollection

                C5.IntersectWith(c6, Intersect.OnBothOperands, pointT, IntPtr.Zero, IntPtr.Zero)
                Dim line3 As New Line(c1.Center, pointT(0))
                Dim line4 As New Line(c1.Center, pointT(1))

                Dim c1L3 As New Point3dCollection
                line3.IntersectWith(c1, Intersect.OnBothOperands, c1L3, IntPtr.Zero, IntPtr.Zero)

                testpt1 = c1L3(0)
                testpt2 = c1L3(1)

                Dim pTanC1A As New Point3d

                If c2.Center.GetVectorTo(testpt1).Length < c2.Center.GetVectorTo(testpt2).Length Then
                    pTanC1A = testpt1
                Else
                    pTanC1A = testpt2
                End If

                Dim c1L4 As New Point3dCollection
                line4.IntersectWith(c1, Intersect.OnBothOperands, c1L4, IntPtr.Zero, IntPtr.Zero)

                testpt1 = c1L4(0)
                testpt2 = c1L4(1)

                Dim pTanC1B As New Point3d

                If c2.Center.GetVectorTo(testpt1).Length < c2.Center.GetVectorTo(testpt2).Length Then
                    pTanC1B = testpt1
                Else
                    pTanC1B = testpt2
                End If

                Dim test1 As New Line(pTanC1A, pTanC2B)
                Dim test2 As New Line(pTanC1B, pTanC2A)

                Dim finals As New Point3dCollection
                test1.IntersectWith(test2, Intersect.OnBothOperands, finals, IntPtr.Zero, IntPtr.Zero)

                Dim final1 As Line
                Dim final2 As Line

                If finals.Count <= 0 Then
                    final1 = New Line(pTanC1A, pTanC2A)
                    final2 = New Line(pTanC1B, pTanC2B)
                Else
                    final1 = test1
                    final2 = test2
                End If

                Dim curspace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)

                curspace.AppendEntity(final1)
                acTrans.AddNewlyCreatedDBObject(final1, True)

                curspace.AppendEntity(final2)
                acTrans.AddNewlyCreatedDBObject(final2, True)

                acTrans.Commit()

            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
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

End Module

Public Module LayerCommands

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
        Dim sets As New XmlReaderSettings
        With sets
            .CloseInput = True
            .IgnoreWhitespace = True
            .Async = False
        End With
        Dim layset As AcdLayers
        Using xr As XmlReader = XmlReader.Create(xmlNameStr, sets)
            'Dim xsets As New XmlSerializerNamespaces
            'xsets.Add("xs", "http://tempuri.org/XMLSchema1.xsd")
            Dim xmlS As New XmlSerializer(GetType(AcdLayers), "http://tempuri.org/XMLSchema1.xsd")
            layset = xmlS.Deserialize(xr)
        End Using
        Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
            Dim lTbl As LayerTable = actrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)
            Dim lz As LayerTableRecord = TryCast(actrans.GetObject(dwgDB.LayerZero, OpenMode.ForRead), LayerTableRecord)
            For i = 0 To layset.Count - 1
                Dim aLayer As AcdLayer = layset.AcdLayer(i)
                If lTbl.Has(aLayer.Name) Then
                    Dim lTR As LayerTableRecord = TryCast(actrans.GetObject(lTbl(aLayer.Name), OpenMode.ForRead), LayerTableRecord)
                    If lTR = lz Then GoTo Skip
                    If lTR.Name = aLayer.Name Then
                        lTR.UpgradeOpen()
                        If aLayer.Remove And Not String.IsNullOrEmpty(aLayer.MergeWith) Then
                            MergeThenDeleteLayer(aLayer.Name, aLayer.MergeWith, True)
                        ElseIf aLayer.Remove And String.IsNullOrEmpty(aLayer.Name) Then
                            DeleteMyLayer(aLayer.Name)
                        ElseIf Not aLayer.Remove And Not String.IsNullOrEmpty(aLayer.MergeWith) Then
                            MergeThenDeleteLayer(aLayer.Name, aLayer.MergeWith, False)
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
                                If dwgDB.PlotStyleMode Then
                                    Try
                                        .PlotStyleName = aLayer.PlotStyle
                                    Catch
                                        Err.Clear()
                                        .PlotStyleName = "Normal"
                                    End Try
                                End If
                                .Description = aLayer.Description
                            End With
                        End If
Skip:
                    End If
                Else
                    Dim ltr As New LayerTableRecord
                    With ltr
                        .Name = aLayer.Name
                        .ViewportVisibilityDefault = aLayer.VpVisDefault
                        .IsOff = aLayer.IsOff
                        .IsFrozen = aLayer.IsFrozen
                        .IsLocked = aLayer.IsLocked
                        .IsPlottable = aLayer.IsPlottable
                        .IsHidden = aLayer.IsHidden
                        .Description = aLayer.Description
                    End With
                    lTbl.UpgradeOpen()
                    Dim ltrID As ObjectId = lTbl.Add(ltr)
                    actrans.AddNewlyCreatedDBObject(ltr, True)
                    Dim lrec As LayerTableRecord = actrans.GetObject(ltrID, OpenMode.ForWrite)
                    With lrec
                        Dim acdColor As Color = GetColor(aLayer.Color)
                        .Color = acdColor
                        .Transparency = GetTransparencyAlpha(CInt(aLayer.Transparency))
                        .LinetypeObjectId = GetLTId(aLayer.Linetype)
                        If dwgDB.PlotStyleMode Then
                            Try
                                .PlotStyleName = aLayer.PlotStyle
                            Catch
                                Err.Clear()
                                .PlotStyleName = "Normal"
                            End Try
                        End If
                    End With
                End If
            Next
            actrans.Commit()
        End Using
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
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("LAYOUTCREATEVIEWPORT", 0)

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
        'Command to change the sheet sizes for layouts based upon the IMVS method.  Requires numbered views corresponding with the number of layouts to be adjusted.

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = curDwg.Editor
        Dim dwgDB As Database = curDwg.Database

        'Try

        Dim lm As LayoutManager = LayoutManager.Current
        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction

            Dim hSize As Double
            Dim vSize As Double

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

            Dim pdo7 As New PromptIntegerOptions(vbLf & "Input starting view number")
            Dim pdr7 As PromptIntegerResult = ed.GetInteger(pdo7)

            Dim startView As Integer

            If pdr7.Status = PromptStatus.OK Then
                startView = pdr7.Value
            Else
                Exit Sub
            End If

            Dim i As Integer = startView - 1

            Dim vtb As ViewTable = acTrans.GetObject(dwgDB.ViewTableId, OpenMode.ForRead)

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

            Dim layDict As DBDictionary = dwgDB.LayoutDictionaryId.GetObject(OpenMode.ForRead)
            Dim layID As ObjectId = layDict(layoutLst(0))
            Dim lo As Layout = acTrans.GetObject(layID, OpenMode.ForWrite)

            'Dim pko As New PromptKeywordOptions(vbLf & "Do you want to pick a new plotter?")
            'With pko
            '    .Keywords.Add("Y")
            '    .Keywords.Add("N")
            '    .AppendKeywordsToMessage = True
            'End With

            'Dim pkr As PromptResult = ed.GetKeywords(pko)

            'Dim newPltr As Boolean
            'Dim myPltr As String
            'Dim uR As String
            'If pkr.Status = PromptStatus.OK Then
            '    uR = pkr.StringResult
            'Else
            '    uR = "N"
            'End If
            'If uR = "Y" Then
            '    newPltr = True
            '    myPltr = GetPrinter()
            'Else
            '    newPltr = False
            '    myPltr = ""
            'End If
            Dim pset As PlotSettings
            Dim psetval As PlotSettingsValidator
            Dim mySetup As String = GetPlotSetup()
            Dim myPltr As String
            Dim plsets As DBDictionary = acTrans.GetObject(dwgDB.PlotSettingsDictionaryId, OpenMode.ForRead)

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
                psetval = PlotSettingsValidator.Current
                pset = plsets.GetAt(mySetup).GetObject(OpenMode.ForWrite)
                psetval.RefreshLists(pset)
                myPltr = pset.PlotConfigurationName
            End If

            Dim vpLayer As String

            If LayerExists("vports") Then
                vpLayer = "vports"
            Else
                vpLayer = AddNewLayer("vports", 203)
            End If

            Dim mySheet As String = GetSheetName(psetval, pset)
            If mySheet = "" Then Exit Sub
            psetval.SetPlotConfigurationName(pset, myPltr, mySheet)

            If layoutLst.Count > 0 Then
                For Each loName As String In layoutLst
                    lm.CurrentLayout = loName
                    Dim loID As ObjectId = layDict(loName)
                    lo = acTrans.GetObject(loID, OpenMode.ForWrite)

                    Dim vpIDs As ObjectIdCollection = lo.GetViewports
                    Dim vp As Autodesk.AutoCAD.DatabaseServices.Viewport = acTrans.GetObject(vpIDs(1), OpenMode.ForWrite)
                    Dim curSpace As BlockTableRecord = acTrans.GetObject(dwgDB.CurrentSpaceId, OpenMode.ForWrite)

                    'Dim vp As New Viewport
                    vp.SetDatabaseDefaults()
                    vp.CenterPoint = New Point3d(hSize / 2, vSize / 2, 0)
                    vp.Height = vSize
                    vp.Width = hSize
                    vp.Layer = vpLayer


                    'lo.CopyFrom(pSet)
                    'Dim pSetVal As PlotSettingsValidator = PlotSettingsValidator.Current
                    'pSetVal.SetPlotType(pSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout)

                    'pSetVal.RefreshLists(pSet)

                    'Dim newPicker As New PlotSettingsForm
                    'Dim myPaperSizes As StringCollection = pSetVal.GetCanonicalMediaNameList(pSet)
                    'For Each nm As String In myPaperSizes
                    '    newPicker.AddSetup(nm)
                    'Next
                    'newPicker.Label1.Text = "Select media for layout"
                    'newPicker.Text = "Media Picker"
                    'newPicker.ShowDialog()

                    'Dim mySheet As String
                    'Dim mp As String
                    'If newPicker.DialogResult = DialogResult.OK Then
                    '    mySheet = newPicker.SelectedSetup
                    'Else
                    '    mySheet = "300x300"
                    'End If

                    'Dim thisList As StringCollection = pSetVal.GetPlotDeviceList
                    'For Each st As String In thisList
                    '    Debug.Print(st)
                    'Next

                    lo.CopyFrom(pset)
                    'Dim pSetVal As PlotSettingsValidator = PlotSettingsValidator.Current
                    Dim check As String = pset.PlotConfigurationName
                    Debug.Print(check)
                    psetval.SetPlotConfigurationName(pset, myPltr, mySheet)
                    psetval.SetPlotType(pset, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout)
                    psetval.SetPlotRotation(pset, PlotRotation.Degrees000)
                    psetval.SetZoomToPaperOnUpdate(pset, True)
                    i += 1

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

        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try

    End Sub


    <CommandMethod("MSPT")>
    Public Sub ModelSpacePlot()
        'plots all block references in model space

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = curDwg.Editor
        Dim dwgDB As Database = curDwg.Database

        'Try

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

        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try

    End Sub

    Public Sub PlotDirect(pSet As PlotSettings, psetVal As PlotSettingsValidator, ByVal toFile As String)

        '' Get the current document and database
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgdb As Database = acDoc.Database
        Dim CurVar As Object
        'Get the BACKGROUNDPLOT SYSTEM VARIABLE
        CurVar = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("BackGroundPlot")
        'SET BACKGROUNDPLOT SYSTEM VARIABLE TO ZERO
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BackGroundPlot", 0)
        ''Start as transaction
        Using acTrans As Transaction = dwgdb.TransactionManager.StartTransaction()
            '' Referring to the Layout Manager
            Dim acLayoutMgr As LayoutManager
            acLayoutMgr = LayoutManager.Current
            '' Get the current layout and output its name in the Command Line window
            Dim acLayout As Layout
            acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead)

            '' Get the PlotInfo from the layout
            Dim acPlInfo As New PlotInfo With {.Layout = acLayout.ObjectId}
            'acPlInfo.Layout = acLayout.ObjectId

            '' Set the PlotSettings object
            psetVal.RefreshLists(pSet)

            '' Center the plot
            psetVal.SetPlotCentered(pSet, True)

            '' Set the plot info as an override
            acPlInfo.OverrideSettings = pSet

            Dim PlInfoVdr As New PlotInfoValidator
            With PlInfoVdr
                .MediaMatchingPolicy = MatchingPolicy.MatchEnabled
                .Validate(acPlInfo)
            End With


            '' Check whether a plot job is in progress
            If PlotFactory.ProcessPlotState = Autodesk.AutoCAD.PlottingServices.
                    ProcessPlotState.NotPlotting Then

                Using acPlEng As PlotEngine = PlotFactory.CreatePublishEngine()

                    '' Track the plot progress with a Progress dialog
                    Dim acPlProgDlg As New PlotProgressDialog(False, 1, True)

                    Using (acPlProgDlg)
                        '' Define the status messages to display when plotting starts
                        acPlProgDlg.PlotMsgString(PlotMessageIndex.DialogTitle) = "Plot Progress"
                        acPlProgDlg.PlotMsgString(PlotMessageIndex.CancelJobButtonMessage) = "Cancel Job"
                        acPlProgDlg.PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage) = "Cancel Sheet"
                        acPlProgDlg.PlotMsgString(PlotMessageIndex.SheetSetProgressCaption) = "Sheet Set Progress"
                        acPlProgDlg.PlotMsgString(PlotMessageIndex.SheetProgressCaption) = "Sheet Progress"

                        '' Set the plot progress range
                        acPlProgDlg.LowerPlotProgressRange = 0
                        acPlProgDlg.UpperPlotProgressRange = 100
                        acPlProgDlg.PlotProgressPos = 0

                        '' Display the Progress dialog
                        acPlProgDlg.OnBeginPlot()
                        acPlProgDlg.IsVisible = True

                        '' Start to plot the layout
                        acPlEng.BeginPlot(acPlProgDlg, Nothing)

                        '' Define the plot output
                        acPlEng.BeginDocument(acPlInfo, acDoc.Name, Nothing, 1, True, toFile)

                        '' Display information about the current plot
                        acPlProgDlg.PlotMsgString(PlotMessageIndex.Status) = "Plotting: " & acDoc.Name & " - " & acLayout.LayoutName

                        '' Set the sheet progress range
                        acPlProgDlg.OnBeginSheet()
                        acPlProgDlg.LowerSheetProgressRange = 0
                        acPlProgDlg.UpperSheetProgressRange = 100
                        acPlProgDlg.SheetProgressPos = 0

                        '' Plot the first sheet/layout
                        Dim acPlPageInfo As New PlotPageInfo
                        acPlEng.BeginPage(acPlPageInfo, acPlInfo, True, Nothing)

                        acPlEng.BeginGenerateGraphics(Nothing)
                        acPlEng.EndGenerateGraphics(Nothing)

                        '' Finish plotting the sheet/layout
                        acPlEng.EndPage(Nothing)
                        acPlProgDlg.SheetProgressPos = 100
                        acPlProgDlg.OnEndSheet()

                        '' Finish plotting the document
                        acPlEng.EndDocument(Nothing)

                        '' Finish the plot
                        acPlProgDlg.PlotProgressPos = 100
                        acPlProgDlg.OnEndPlot()
                        acPlEng.EndPlot(Nothing)

                    End Using
                End Using
            End If
        End Using
        'REVERT BACK TO THE ORIGINAL BACKGROUNDPLOT SYSTEM VARIABLE
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BackGroundPlot", CurVar)
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

    Public Function MkArrowHead(p1 As Point3d, p2 As Point3d, paramA As Double) As ObjectId

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        'create a vector2d from p1 & p2 to get the final orientation angle for the arrowhead
        Dim arrowVect As Vector3d = p1.GetVectorTo(p2)
        Dim arrowVect2d As Vector2d = arrowVect.Convert2d(New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis))
        Dim arrowAng As Double = arrowVect2d.Angle

        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
            'using an origin located on the cenerline of the arrow at the intersection of the arrowhead base lines
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
                ed.WriteMessage(vbLf & "Error in MkArrowHead Function")
                Return Nothing
                Exit Function
            ElseIf iPts Is Nothing Then
                ed.WriteMessage(vbLf & "Error in MkArrowHead Function")
                Return Nothing
                Exit Function
            End If

            'set the opposite tangent point
            Dim negTanPt As New Point2d(tanPt.X, -tanPt.Y)

            'get the bulge value for the polyline
            Dim ahv1 As Vector2d = ptR2d.GetVectorTo(tanPt)
            Dim ahv2 As Vector2d = ptP2d.GetVectorTo(s1)
            Dim bulgeAng As Double = ahv1.GetAngleTo(ahv2)
            Dim myblg As Double = Tan(bulgeAng / 4)

            'create a polyline
            Dim pl0 As New Polyline
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
            Dim plObjId As ObjectId = curSpace.AppendEntity(pl0)
            acTrans.AddNewlyCreatedDBObject(pl0, True)

            Return plobjid

            'dispose of the temporary circle
            c1.Dispose()
            acTrans.Commit()

        End Using

    End Function


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

    Public Sub MkDirectionalArrowType1(p1 As Point3d, aobj As ArrowObj)
        'Creates directional arrow types 1 through 3
        'per California Sign Specifications Appendix

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
            'set the vertices points
            Dim paramA As Double = aobj.ParamA
            Dim paramB As Double = aobj.ParamB
            Dim paramC As Double = aobj.ParamC
            Dim paramD As Double = aobj.ParamD
            Dim paramE As Double = aobj.ParamE
            Dim paramF As Double = aobj.ParamF
            Dim paramR As Double = aobj.ParamR
            Dim arrowAng As Double = aobj.ArrowAng
            Dim ptO As Point3d = Point3d.Origin
            Dim s12d As New Point2d(ptO.X, paramC / 2)
            Dim s22d As New Point2d(ptO.X, -paramC / 2)
            Dim cen1 As New Point3d(paramE - paramR, (paramA / 2) - paramR, ptO.Z)
            Dim cen2 As New Point3d(-(paramB - 2 * paramR), ptO.Y, ptO.Z)

            Dim ptQ As Point3d
            Dim ptR As Point3d

            'offset temporary line between circle centers by the radius to get tangent points.
            'make sure the direction is correct.
            Using tempLn As New Line(cen1, cen2)
                Dim pR As Double = paramR
                Dim chkVal As Double = 0
                Dim tl As Line
                Do
                    Dim oSet As DBObjectCollection = tempLn.GetOffsetCurves(pR)
                    tl = TryCast(oSet(0), Line)
                    chkVal = tl.EndPoint.Y
                    pR *= -1
                Loop Until chkVal > 0
                ptQ = tl.StartPoint
                ptR = tl.EndPoint
            End Using

            'create temporary circle for finding tangents
            Using c1 As New Circle(cen1, Vector3d.ZAxis, paramR)

                'get the tangent points
                Dim ptQ2d As New Point2d(ptQ.X, ptQ.Y)
                Dim ptR2d As New Point2d(ptR.X, ptR.Y)

                Dim tanPts2 As Point2dCollection = GetTangentPoints(New Point3d(s12d.X, s12d.Y, ptO.Z), c1)

                Dim ptP2d As Point2d

                'make sure that the tangent points exist and weed out the one that does not apply
                If tanPts2 IsNot Nothing And tanPts2.Count = 2 Then
                    If tanPts2(0).X > tanPts2(1).X Then
                        ptP2d = tanPts2(0)
                    Else
                        ptP2d = tanPts2(1)
                    End If
                Else
                    ed.WriteMessage(vbLf & "Error.  External tangents not found.")
                    Exit Sub
                End If

                Dim ptT As New Point2d(paramF - paramB + paramE, paramD / 2)
                Dim negPtT As New Point2d(ptT.X, -ptT.Y)

                'set the opposite tangent points
                Dim negPtR As New Point2d(ptR2d.X, -ptR2d.Y)
                Dim negPtQ As New Point2d(ptQ2d.X, -ptQ2d.Y)
                Dim negPtP As New Point2d(ptP2d.X, -ptP2d.Y)

                'get the bulge values for the polyline
                Dim vect1 As Vector2d = s12d.GetVectorTo(ptP2d)
                Dim vect2 As Vector2d = ptQ2d.GetVectorTo(ptR2d)
                Dim vect3 As Vector2d = negPtR.GetVectorTo(negPtQ)
                Dim ang1 As Double = vect1.GetAngleTo(vect2)
                Dim ang2 As Double = vect2.GetAngleTo(vect3)
                Dim blg1 As Double = Tan(ang1 / 4)
                Dim blg2 As Double = Tan(ang2 / 4)

                'set a point at the arrow tip
                Dim tipPt As New Point3d(paramE - paramB, 0, ptO.Z)

                'store the vector from the tip of the temporary arrow to its final location
                Dim moveVect As Vector3d = tipPt.GetVectorTo(p1)

                'create a polyline
                Using pl As New Polyline
                    pl.AddVertexAt(0, ptT, 0, 0, 0)
                    pl.AddVertexAt(1, s12d, 0, 0, 0)
                    pl.AddVertexAt(2, ptP2d, blg1, 0, 0)
                    pl.AddVertexAt(3, ptQ2d, 0, 0, 0)
                    pl.AddVertexAt(4, ptR2d, blg2, 0, 0)
                    pl.AddVertexAt(5, negPtR, 0, 0, 0)
                    pl.AddVertexAt(6, negPtQ, blg1, 0, 0)
                    pl.AddVertexAt(7, negPtP, 0, 0, 0)
                    pl.AddVertexAt(8, s22d, 0, 0, 0)
                    pl.AddVertexAt(9, negPtT, 0, 0, 0)
                    pl.Closed = True

                    'move and rotate the polyline to its final position
                    pl.TransformBy(Matrix3d.Displacement(moveVect))
                    pl.TransformBy(Matrix3d.Rotation(arrowAng, Vector3d.ZAxis, p1))

                    'add it to the current space
                    Dim curSpace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)
                    curSpace.AppendEntity(pl)
                    acTrans.AddNewlyCreatedDBObject(pl, True)
                End Using
                'dispose of the temporary circle
            End Using
            acTrans.Commit()
        End Using
    End Sub

    Public Sub MkDirectionalArrowType4(p1 As Point3d, aobj As ArrowObj)
        'Creates a 1 line horizontal, vertical, or diagonal directional arrow
        'per California Sign Specifications Appendix

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        'make sure arrow object is not nothing
        Dim ao As ArrowObj
        If aobj IsNot Nothing Then
            ao = aobj
        Else
            Exit Sub
        End If

        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
            'using an origin located on the cenerline of the arrow at the end of the arrow shaft,
            'get temporary arrow coordinate geometry per California/Federal sign specifications
            Dim paramA As Double = ao.ParamA
            Dim paramB As Double = ao.ParamB
            Dim paramC As Double = ao.ParamC
            Dim paramD As Double = ao.ParamD
            Dim paramE As Double = ao.ParamE
            Dim paramR As Double = ao.ParamR
            'Dim theta1 As Double = Asin(paramE / (paramB - (paramA / 2)))
            'Dim paramD As Double = (paramA / 2) * Tan(theta1)
            Dim ptO As Point3d = Point3d.Origin
            Dim s12d As New Point2d(0, paramC / 2)
            Dim s22d As New Point2d(0, -paramC / 2)
            Dim cen1 As New Point3d(paramD - paramR, (paramA / 2) - paramR, ptO.Z)
            Dim ptR As New Point3d(-(paramB - paramD), 0, 0)
            Dim ptR2d As New Point2d(ptR.X, ptR.Y)
            Dim ptS3d As New Point3d(s12d.X, s12d.Y, 0)
            Dim ptT2d As New Point2d(paramE - paramB, paramC / 2)
            Dim negPtT As New Point2d(ptT2d.X, -ptT2d.Y)

OtherSide:
            'create temporary circle for finding tangents
            Dim c1 As New Circle(cen1, Vector3d.ZAxis, paramR)

            'get the tangent points from the tip and the shaft
            Dim tanPts As Point2dCollection = GetTangentPoints(ptS3d, c1)
            Dim tanPts2 As Point2dCollection = GetTangentPoints(ptR, c1)

            Dim ptP2d As Point2d
            'make sure that the tangent points exist and weed out the one that does not apply
            If tanPts IsNot Nothing And tanPts.Count = 2 Then
                If tanPts(0).X > tanPts(1).X Then
                    ptP2d = tanPts(0)
                Else
                    ptP2d = tanPts(1)
                End If
            Else
                ed.WriteMessage(vbLf & "Error.  External tangents not found.")
                Exit Sub
            End If

            Dim ptQ2d As Point2d
            If tanPts2 IsNot Nothing And tanPts2.Count = 2 Then
                If tanPts2(0).Y > tanPts2(1).Y Then
                    ptQ2d = tanPts2(0)
                Else
                    ptQ2d = tanPts2(1)
                End If
            Else
                ed.WriteMessage(vbLf & "Error.  External tangents not found.")
                Exit Sub
            End If

            'set the opposite tangent points
            Dim negPtQ As New Point2d(ptQ2d.X, -ptQ2d.Y)
            Dim negPtP As New Point2d(ptP2d.X, -ptP2d.Y)

            'get the bulge value for the polyline
            Dim vect1 As Vector2d = s12d.GetVectorTo(ptP2d)
            Dim vect2 As Vector2d = ptQ2d.GetVectorTo(ptR2d)
            Dim ang1 As Double = vect1.GetAngleTo(vect2)
            Dim blg1 As Double = Tan(ang1 / 4)

            'store the vector from the tip of the temporary arrow to its final location
            Dim moveVect As Vector3d = ptR.GetVectorTo(p1)

            'create a polyline
            Using pl As New Polyline
                pl.AddVertexAt(0, ptT2d, 0, 0, 0)
                pl.AddVertexAt(1, s12d, 0, 0, 0)
                pl.AddVertexAt(2, ptP2d, blg1, 0, 0)
                pl.AddVertexAt(3, ptQ2d, 0, 0, 0)
                pl.AddVertexAt(4, ptR2d, 0, 0, 0)
                pl.AddVertexAt(5, negPtQ, blg1, 0, 0)
                pl.AddVertexAt(6, negPtP, 0, 0, 0)
                pl.AddVertexAt(7, s22d, 0, 0, 0)
                pl.AddVertexAt(8, negPtT, 0, 0, 0)
                pl.Closed = True

                'move and rotate the polyline to its final position
                pl.TransformBy(Matrix3d.Displacement(moveVect))
                pl.TransformBy(Matrix3d.Rotation(ao.ArrowAng, Vector3d.ZAxis, p1))

                'add it to the current space
                Dim curSpace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)
                curSpace.AppendEntity(pl)
                acTrans.AddNewlyCreatedDBObject(pl, True)
            End Using

            'dispose of the temporary circle
            c1.Dispose()

            acTrans.Commit()
        End Using

    End Sub

End Module
