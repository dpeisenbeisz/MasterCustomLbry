Imports System.IO
'Imports System.Xml.Serialization
'Imports System.Xml
Imports System.Math
Imports System.Text
'Imports System.Reflection
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
'Imports Autodesk.AutoCAD.GraphicsInterface
Imports C3dAs = Autodesk.Civil.ApplicationServices
Imports C3dDb = Autodesk.Civil.DatabaseServices
'Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.PlottingServices
Imports System.Xml
Imports Autodesk.AutoCAD.Colors
Imports System.Xml.Serialization
Imports System.Runtime.CompilerServices
Imports GI = Autodesk.AutoCAD.GraphicsInterface


'Imports System.Collections.Specialized
'Imports Autodesk.AutoCAD.GraphicsInterface
'Imports System.Runtime.CompilerServices


'Module StringExtension
'    <Extension()>
'    Function RemoveInvalidChars(ByVal originalString As String) As String
'        Dim finalString As String = String.Empty

'        If Not String.IsNullOrEmpty(originalString) Then
'            Return String.Concat(originalString.Split(Path.GetInvalidFileNameChars()))
'        End If

'        Return finalString
'    End Function
'End Module


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

    <CommandMethod("SetDwgsBase")>
    Public Sub SetDwgsBase()

        Dim filepath As String = ""
        'If BlkFolder <> "" Then
        '    Dim uResp1 As Integer = MsgBox("Does this folder contain the dwg files you want to change?" & vbLf &
        '       BlkFolder, vbYesNoCancel)
        '    If uResp1 = vbCancel Then Exit Sub

        '    If uResp1 = vbYes Then
        '        filepath = BlkFolder
        '        GoTo skipit
        '    End If
        'End If

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
        Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction

            Dim blkTbl As BlockTable = actrans.GetObject(DwgDB.BlockTableId, OpenMode.ForWrite)
            Dim cSpace As BlockTableRecord = actrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForWrite)

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

        'dim dv As DwgVersion
        'dv = DwgVersion(curDwg.Name)

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

        'Dim SelResult As PromptSelectionResult = ed.SelectImplied()
        'If SelResult.Status = PromptStatus.Error Then
        '    Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select dynamic pavment marking blocks to change:")}
        '    SelResult = ed.GetSelection(Seloptions)
        'Else
        '    ed.SetImpliedSelection(New ObjectId(-1) {})
        'End If

        'If SelResult.Status = PromptStatus.OK Then

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

        Dim pso As New PromptStringOptions(vbLf & "Input the word to be created (no spaces allowed). ")
        With pso
            .AllowSpaces = False
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
            Exit Sub
        End If

        Dim wrdName As String = "P-" & myStr

        If BlkExists(wrdName) Then
            ed.WriteMessage(vbLf & "Block with that name already exists. Ending command.")
            Exit Sub
        End If

        Dim blkPath As String = "\\EESServer\datadisk\CAD\BLOCKS\ROAD\PVMT\Design\"

        Dim letters(myStr.Length - 1) As String
        For j As Integer = 1 To myStr.Length
            letters(j - 1) = Mid(myStr, j, 1)
        Next

        Dim cPos As Double = 0.0
        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
            Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
            Dim curSpace As BlockTableRecord = acTrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForRead)
            Dim newBlk As New BlockTableRecord With {.Name = "P-" & myStr}
            Dim MidDist As Double

            For i = 0 To myStr.Length - 1
                Dim blkName As String = "P-" & letters(i)
                Dim blkFullName As String = blkPath & blkName & ".dwg"
                Dim ltrID As ObjectId

                If Not BlkExists(blkName) Then
                    ltrID = InsertDwg(blkFullName, blkName)
                Else
                    ltrID = blkTbl(blkName)
                End If

                Dim inspt As New Point3d(cPos, 0, 0)

                Dim bDef As BlockTableRecord = acTrans.GetObject(ltrID, OpenMode.ForRead)
                Dim bRef As New BlockReference(inspt, ltrID)

                newBlk.AppendEntity(bRef)
                'acTrans.AddNewlyCreatedDBObject(bRef, True)

                If Not letters(i).ToUpper = "I" Then
                    cPos += 0.33333 * 5
                Else
                    cPos += 0.33333 * 2
                End If

                MidDist = (cPos - 0.33333) / 2
            Next

            blkTbl.UpgradeOpen()
            Dim wrdblkId As ObjectId = blkTbl.Add(newBlk)

            For Each id As ObjectId In newBlk
                Dim obj As DBObject = acTrans.GetObject(id, OpenMode.ForRead)
                If TypeOf obj Is BlockReference Then
                    Dim ltrBlk As BlockReference = CType(obj, BlockReference)
                    ltrBlk.UpgradeOpen()

                    Dim orgPt As New Point3d(0, 0, 0)
                    Dim cPt As New Point3d(MidDist, 4, 0)

                    Dim dispVect As Vector3d = cPt.GetVectorTo(orgPt)
                    ltrBlk.TransformBy(Matrix3d.Displacement(dispVect))

                    Dim ltrID As ObjectId = ltrBlk.ObjectId
                    If vis = "Existing" Then
                        ChangeVisState(ltrID, "Existing", "Visibility1")
                    ElseIf vis = "Proposed" Then
                        ChangeVisState(ltrID, "Proposed", "Visibility1")
                    ElseIf vis = "Remove" Then
                        ChangeVisState(ltrID, "Remove", "Visibility1")
                    End If
                End If
            Next


            Dim ppo As New PromptPointOptions(vbLf & "Pick point for location of pavement marking.")

            Dim ppr As PromptPointResult = ed.GetPoint(ppo)
            Dim iP As Point3d

            If ppr.Status = PromptStatus.OK Then
                iP = ppr.Value
            Else
                Exit Sub
            End If

            Dim wrdBref As New BlockReference(iP, wrdblkId)

            curSpace.UpgradeOpen()
            curSpace.AppendEntity(wrdBref)
            acTrans.AddNewlyCreatedDBObject(wrdBref, True)

            acTrans.Commit()
        End Using

    End Sub

End Module

Public Module GeometryCommands

    <CommandMethod("ETAN")>
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
                    Exit Sub
                End If

                If c1.Radius > c2.Radius Then
                    Dim tempC As Circle = c1
                    c1 = c2
                    c2 = tempC
                End If

                Dim raddiff = c2.Radius - c1.Radius

                Dim c3 As New Circle(c2.Center, Vector3d.ZAxis, raddiff)

                Dim ctLine As New Line(c2.Center, c1.Center)
                Dim halfline As Double = ctLine.Length / 2
                Dim quarterline As Double = ctLine.Length / 4
                Dim midpt As Point3d = ctLine.GetPointAtDist(halfline)
                Dim pointS As Point3d = ctLine.GetPointAtDist(quarterline)
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

                Dim C5 As New Circle(c1.Center, Vector3d.ZAxis, c3.Radius)
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

                If c2.Center.GetVectorTo(testpt1).Length < c1.Center.GetVectorTo(testpt2).Length Then
                    pTanC1A = testpt2
                Else
                    pTanC1A = testpt1
                End If

                Dim c1L4 As New Point3dCollection
                line4.IntersectWith(c1, Intersect.OnBothOperands, c1L4, IntPtr.Zero, IntPtr.Zero)

                testpt1 = c1L4(0)
                testpt2 = c1L4(1)

                Dim pTanC1B As New Point3d

                If c2.Center.GetVectorTo(testpt1).Length < c1.Center.GetVectorTo(testpt2).Length Then
                    pTanC1B = testpt2
                Else
                    pTanC1B = testpt1
                End If

                Dim test1 As New Line(pTanC1A, pTanC2A)
                Dim test2 As New Line(pTanC1B, pTanC2B)

                Dim finals As New Point3dCollection
                test1.IntersectWith(test2, Intersect.OnBothOperands, finals, IntPtr.Zero, IntPtr.Zero)

                Dim intPt As Point3d
                Dim final1 As Line
                Dim final2 As Line

                Try
                    intPt = finals(0)
                    final1 = New Line(pTanC1A, pTanC2B)
                    final2 = New Line(pTanC1B, pTanC2A)
                Catch
                    final1 = test1
                    final2 = test2
                End Try


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

    '<CommandMethod("REVPLINES", CommandFlags.UsePickSet Or CommandFlags.Redraw Or CommandFlags.Modal)>
    'Public Sub REVPLINES()

    '    Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '    Dim DwgDB As Database = CurDwg.Database
    '    Dim ed As Editor = CurDwg.Editor

    '    Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction()

    '        Dim SelResult As PromptSelectionResult = ed.SelectImplied()

    '        If SelResult.Status = PromptStatus.Error Then
    '            Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select polylines to reverse: ")}
    '            'Seloptions.MessageForAdding = String.Format(vbLf & "Select polylines to reverse: ")
    '            SelResult = ed.GetSelection(Seloptions)
    '        Else
    '            ed.SetImpliedSelection(New ObjectId(-1) {})
    '        End If

    '        If SelResult.Status = PromptStatus.OK Then

    '            Dim acSSet As SelectionSet = SelResult.Value
    '            Dim myobjIDs As ObjectId() = acSSet.GetObjectIds

    '            'Dim i As Integer = 0
    '            For Each objID As ObjectId In myobjIDs

    '                Dim obj As Object = actrans.GetObject(objID, OpenMode.ForWrite)
    '                Dim cur = TryCast(obj, Curve)

    '                If cur IsNot Nothing Then
    '                    Try
    '                        cur.ReverseCurve()
    '                    Catch
    '                        ed.WriteMessage(vbLf & "Could not reverse object of type {0}.", cur.[GetType]().Name)
    '                    End Try
    '                Else
    '                    ed.WriteMessage(vbLf & "Could not reverse object.")
    '                End If
    '            Next
    '        End If
    '        actrans.Commit()
    '    End Using
    'End Sub


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

    <CommandMethod("SF2FL")>
    Public Sub FIGS2Feats()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim curDwg As Document = docCol.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim cvlDwg As C3dAs.CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim ed As Editor = curDwg.Editor

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
            Dim psOptions As New PromptSelectionOptions

            With psOptions
                .MessageForAdding = vbLf & "Select survey figures to convert to feature lines: "
                .AllowSubSelections = True
            End With

            Dim acSSPrompt As PromptSelectionResult = ed.GetSelection(psOptions)
            Dim ObjColl As New ObjectIdCollection

            Dim resColl As New ObjectIdCollection

            '' If the prompt status is OK, objects were selected
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                '' Step through the objects in the selection set
                For Each acSSObj As SelectedObject In acSSet
                    '' Check to make sure a valid SelectedObject object was returned
                    If Not IsDBNull(acSSObj) Then
                        Dim tempObj As Object = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        If TypeOf tempObj Is C3dDb.SurveyFigure Then
                            ObjColl.Add(acSSObj.ObjectId)
                        End If
                    End If
                Next
            End If

            Dim blkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForWrite)
            Dim curSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            Dim pKO As New PromptKeywordOptions("")

            With pKO
                .Keywords.Add("Yes")
                .Keywords.Add("No")
                .Message = vbLf & "Delete survey figures after converting to feature lines?"
                .AppendKeywordsToMessage = True
                .AllowArbitraryInput = False
                .AllowNone = False
            End With
            Dim eraseOrigObj As Boolean

            Dim pkRes As PromptResult = ed.GetKeywords(pKO)
            If pkRes.Status = PromptStatus.OK Then
                If pkRes.StringResult = "Yes" Then eraseOrigObj = True
                If pkRes.StringResult = "No" Then eraseOrigObj = False
            End If

            For Each lineID As ObjectId In ObjColl
                Dim lineObj As C3dDb.SurveyFigure = TryCast(acTrans.GetObject(lineID, OpenMode.ForRead), C3dDb.SurveyFigure)

                If lineObj IsNot Nothing Then
                    Dim LinePts As Point3dCollection = lineObj.GetPoints(0)
                    Dim objSite As ObjectId = lineObj.SiteId
                    Dim newPolyId As ObjectId = Fig2Poly2D(lineID)
                    Dim tempFlId As ObjectId
                    tempFlId = C3dDb.FeatureLine.Create(newPolyId.ToString, newPolyId)
                    Dim newPoly As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(newPolyId, OpenMode.ForWrite)
                    newPoly.Erase()
                    Dim noOfPts As Integer = LinePts.Count
                    Dim newFeat As C3dDb.FeatureLine = acTrans.GetObject(tempFlId, OpenMode.ForWrite)
                    For i = 0 To noOfPts - 1
                        newFeat.SetPointElevation(i, LinePts(i).Z)
                    Next
                End If
            Next

            If eraseOrigObj = True Then
                For Each obID As ObjectId In ObjColl
                    Dim origFig As Autodesk.AutoCAD.DatabaseServices.Entity = TryCast(acTrans.GetObject(obID, OpenMode.ForWrite), Autodesk.AutoCAD.DatabaseServices.Entity)
                    origFig.Erase()
                Next
            End If
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


    '<CommandMethod("REVPCurves")>
    'Public Sub RevPCurve()
    '    Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '    Dim DwgDB As Database = CurDwg.Database
    '    'Dim acTrans As Transaction = DwgDB.TransactionManager.StartTransaction()
    '    Dim ed As Editor = CurDwg.Editor

    '    Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction

    '        Dim SelResult As PromptSelectionResult = ed.SelectImplied()
    '        If SelResult.Status = PromptStatus.Error Then
    '            Dim Seloptions As New PromptSelectionOptions With {.MessageForAdding = String.Format(vbLf & "Select polyline curves to reverse")}
    '            'Seloptions.MessageForAdding = String.Format(vbLf & "Select centerlines")
    '            SelResult = ed.GetSelection(Seloptions)
    '        Else
    '            ed.SetImpliedSelection(New ObjectId(-1) {})
    '        End If

    '        Dim acSSet As SelectionSet = SelResult.Value
    '        Dim MyobjIDs As ObjectId() = acSSet.GetObjectIds

    '        For Each objID As ObjectId In MyobjIDs

    '            Dim obj As Object = acTrans.GetObject(objID, OpenMode.ForWrite)
    '            Dim Cur = TryCast(obj, Curve)
    '            If Cur IsNot Nothing Then
    '                Try
    '                    Cur.ReverseCurve()
    '                Catch
    '                    ed.WriteMessage(vbLf & "Could not reverse object of type {0}.", Cur.[GetType]().Name)
    '                End Try
    '            End If
    '        Next
    '        acTrans.Commit()
    '    End Using
    'End Sub


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

            'For k = 0 To i - 1
            '    'For Each plotStyle As String In PlotSettingsValidator.Current.GetPlotStyleSheetList()
            '    ' Output the names of the available plot styles
            '    ed.WriteMessage(vbLf & PScoll(k))
            'Next k
        End Using
        Call DumpPltStyleTables()

    End Sub



    <CommandMethod("ShowPStyles")>
    Public Sub PStyleSamples()

        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = CurDwg.Editor

        'Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction

        '    Dim cSpace As BlockTableRecord = actrans.GetObject(DwgDB.CurrentSpaceId, OpenMode.ForRead)

        '    Dim layMgr As LayoutManager = LayoutManager.Current
        '    Dim curLayout As Layout = actrans.GetObject(layMgr.GetLayoutId(layMgr.CurrentLayout), OpenMode.ForWrite)
        '    Dim plInfo As New PlotInfo()
        '    plInfo.Layout = curLayout.ObjectId
        '    Dim plSettings As New PlotSettings(curLayout.ModelType)
        '    plSettings.CopyFrom(curLayout)
        '    Dim psv As PlotSettingsValidator = PlotSettingsValidator.Current
        '    'psv.SetCurrentStyleSheet(plSettings, "EES Standard 5.0 10-Scale.stb")
        '    'psv.SetCurrentStyleSheet(plSettings, "\\eesserver\datadisk\CAD\plot styles\PsTest.stb")
        '    psv.SetCurrentStyleSheet(plSettings, "PsTest.stb")
        '    curLayout.UpgradeOpen()
        '    curLayout.CopyFrom(plSettings)
        '    actrans.Commit()
        'End Using

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

        'Dim showSetup As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("SHOWPAGESETUPFORM")
        ' Dim createView As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("LAYOUTCREATEVIEWPORT")

        'Using dkLck As DocumentLock = curDwg.LockDocument

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

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction

            'Dim showForm As Int16 = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("SHOWPAGESETUPFORM")
            'Dim cVp As Int16 = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("LAYOUTCREATEVIEWPORT")

            Dim showForm As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("SHOWPAGESETUPFORNEWLAYOUTS")
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("SHOWPAGESETUPFORNEWLAYOUTS", 0)
            'Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("LAYOUTCREATEVIEWPORT", 1)

            Dim shtNm As String = ""

            Dim pSet As PlotSettings = Nothing
            Dim psetVal As PlotSettingsValidator
            If makeLayouts Then

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
            'Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("LAYOUTCREATEVIEWPORT", cVp)
            acTrans.Commit()
        End Using
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

    <CommandMethod("CountSheets")>
    Public Sub CountSheets()

        Dim myTabs As SortedDictionary(Of Integer, String) = LayoutTabList()
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        ed.WriteMessage(vbCrLf & "Number of Layouts:  " & myTabs.Count) '' Number of Tabs not counting Model Space Tab
        For Each v In myTabs.Values
            ed.WriteMessage(vbCrLf & v.ToString)
        Next

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

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        'declare two points for final arrow position and orientation
        Dim p1 As Point3d
        Dim p2 As Point3d

        'get the tip location and store in p1
        Dim ppo As New PromptPointOptions(vbLf & "Pick tip of arrowhead")
        With ppo
            .AllowNone = False
            .AllowArbitraryInput = True
        End With

        Dim ppr As PromptPointResult = ed.GetPoint(ppo)

        If ppr.Status = PromptStatus.OK Then
            p1 = ppr.Value

            'get second point for arrow orientation
            Dim ppo2 As New PromptPointOptions(vbLf & "Pick direction of arrow shaft")
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
        Else
            Exit Sub
        End If

        'create a vector2d from p1 & p2 to get the final orientation angle for the arrowhead
        Dim arrowVect As Vector3d = p1.GetVectorTo(p2)
        Dim arrowVect2d As Vector2d = arrowVect.Convert2d(New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis))
        Dim arrowAng As Double = arrowVect2d.Angle

        'get the width of the arrow shaft - this is the critical parameter for the arrowhead
        Dim pdo As New PromptDistanceOptions(vbLf & "Input or pick width of arrow shaft.")
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
            'using an origin located on the cenerline of the arrow at the intersection of the arrowhead base lines
            'get temporary arrow coordinate geometry per Federal Sign Manual Appendix
            Dim paramB As Double = 1.21 * paramA
            Dim paramC As Double = 2 * paramA
            Dim paramE As Double = 0.21 * paramA
            Dim theta1 As Double = (Atan(0.21 / 0.71))
            Dim paramD As Double = (paramA * Tan(theta1)) / 2
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

            'Dim tempCRad As Double = Sqrt(paramB ^ 2 + paramC ^ 2)
            'Dim tempPt As New Point2d(-tempCRad, ptO.Y)
            'Dim iPts As Point2dCollection = FindTangentPoints(paramE, tempPt, New Point2d(cen1.X, cen1.Y))

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


            'Dim trnsvct As Vector3d = New Point3d(tempPt.X, tempPt.Y, 0).GetVectorTo(ptR)
            'Dim rotVect As Vector2d = ptR2d.GetVectorTo(New Point2d(cen1.X, cen1.Y))
            'Dim rotAng As Double = Vector2d.XAxis.GetAngleTo(rotVect)

            'Dim mirrLine As New Line3d(ptR, ptO)

            'Dim pl1 As New Polyline
            'pl1.AddVertexAt(0, s1, 0, 0, 0)
            'pl1.AddVertexAt(1, ptP, 0, 0, 0)

            'Dim pl2 As New Polyline
            'pl2.AddVertexAt(0, tempPt, 0, 0, 0)
            'pl2.AddVertexAt(0, ptR2d, 0, 0, 0)
            'pl2.AddVertexAt(1, tanPt, 0, 0, 0)
            'pl2.TransformBy(Matrix3d.Displacement(trnsvct))
            'pl2.TransformBy(Matrix3d.Rotation(rotAng, Vector3d.ZAxis, ptR))

            'Dim ahV13d As Vector3d = pl2.StartPoint.GetVectorTo(pl2.EndPoint)
            'Dim ahv1 As Vector2d = ahV13d.Convert2d(New Plane(Point3d.Origin, Vector3d.ZAxis))

            'get the bulge value for the polyline
            Dim ahv1 As Vector2d = ptR2d.GetVectorTo(tanPt)
            Dim ahv2 As Vector2d = ptP2d.GetVectorTo(s1)
            Dim bulgeAng As Double = ahv1.GetAngleTo(ahv2)
            Dim myblg As Double = Tan(bulgeAng / 4)

            'Dim pp1 As Point2d = New Point2d(pl2.EndPoint.X, pl2.EndPoint.Y)
            'Dim pp2 As Point2d = New Point2d(pl2.StartPoint.X, pl2.StartPoint.Y)
            'Dim pp3 As Point2d = New Point2d(pp1.X, -pp1.Y)
            'Dim pp4 As Point2d = New Point2d(ptP.X, -ptP.Y)
            'Dim pp5 As Point2d = New Point2d(pl1.EndPoint.X, -pl1.EndPoint.Y)

            'create a polyline
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
                With curSpace
                    .AppendEntity(pl0)
                    acTrans.AddNewlyCreatedDBObject(pl0, True)
                End With
            End Using

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
            Dim theta1 As Double = (Atan(0.21 / 0.71))
            Dim paramD As Double = (paramA * Tan(theta1)) / 2
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

            'Dim tempCRad As Double = Sqrt(paramB ^ 2 + paramC ^ 2)
            'Dim tempPt As New Point2d(-tempCRad, ptO.Y)
            'Dim iPts As Point2dCollection = FindTangentPoints(paramE, tempPt, New Point2d(cen1.X, cen1.Y))

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


            'Dim trnsvct As Vector3d = New Point3d(tempPt.X, tempPt.Y, 0).GetVectorTo(ptR)
            'Dim rotVect As Vector2d = ptR2d.GetVectorTo(New Point2d(cen1.X, cen1.Y))
            'Dim rotAng As Double = Vector2d.XAxis.GetAngleTo(rotVect)

            'Dim mirrLine As New Line3d(ptR, ptO)

            'Dim pl1 As New Polyline
            'pl1.AddVertexAt(0, s1, 0, 0, 0)
            'pl1.AddVertexAt(1, ptP, 0, 0, 0)

            'Dim pl2 As New Polyline
            'pl2.AddVertexAt(0, tempPt, 0, 0, 0)
            'pl2.AddVertexAt(0, ptR2d, 0, 0, 0)
            'pl2.AddVertexAt(1, tanPt, 0, 0, 0)
            'pl2.TransformBy(Matrix3d.Displacement(trnsvct))
            'pl2.TransformBy(Matrix3d.Rotation(rotAng, Vector3d.ZAxis, ptR))

            'Dim ahV13d As Vector3d = pl2.StartPoint.GetVectorTo(pl2.EndPoint)
            'Dim ahv1 As Vector2d = ahV13d.Convert2d(New Plane(Point3d.Origin, Vector3d.ZAxis))

            'get the bulge value for the polyline
            Dim ahv1 As Vector2d = ptR2d.GetVectorTo(tanPt)
            Dim ahv2 As Vector2d = ptP2d.GetVectorTo(s1)
            Dim bulgeAng As Double = ahv1.GetAngleTo(ahv2)
            Dim myblg As Double = Tan(bulgeAng / 4)

            'Dim pp1 As Point2d = New Point2d(pl2.EndPoint.X, pl2.EndPoint.Y)
            'Dim pp2 As Point2d = New Point2d(pl2.StartPoint.X, pl2.StartPoint.Y)
            'Dim pp3 As Point2d = New Point2d(pp1.X, -pp1.Y)
            'Dim pp4 As Point2d = New Point2d(ptP.X, -ptP.Y)
            'Dim pp5 As Point2d = New Point2d(pl1.EndPoint.X, -pl1.EndPoint.Y)

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

            With curSpace
                .AppendEntity(pl0)
                acTrans.AddNewlyCreatedDBObject(pl0, True)
            End With

            Return pl0.ObjectId

            'dispose of the temporary circle
            c1.Dispose()
            acTrans.Commit()

        End Using

    End Function

End Module
