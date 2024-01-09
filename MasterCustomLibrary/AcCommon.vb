Imports System.IO
'Imports System.Reflection
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
'Imports Autodesk.AutoCAD.GraphicsInterface
'Imports C3dAs = Autodesk.Civil.ApplicationServices
'Imports C3dDb = Autodesk.Civil.DatabaseServices
'Imports SD = System.Drawing
'Imports System.Runtime
Imports System.Xml
Imports System.Text
Imports Autodesk.AutoCAD.Colors
Imports System.Xml.Serialization
Imports System.Math
Imports System.Collections.Specialized
Imports Autodesk.AutoCAD.PlottingServices

Public Module Layers

    Public Function AddNewLayer(ByVal sLayerName As String, ByVal ColorNo As Integer) As String

        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim MyColor As Color = Color.FromColorIndex(ColorMethod.ByAci, ColorNo)

        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction()
            Dim acLyrTbl As LayerTable = acTrans.GetObject(DwgDB.LayerTableId, OpenMode.ForRead)
            If acLyrTbl.Has(sLayerName) Then
                Return sLayerName
                Exit Function
            Else
                Try
                    Using AcLyrTblRec As New LayerTableRecord()
                        With AcLyrTblRec
                            AcLyrTblRec.Color = MyColor
                            AcLyrTblRec.Name = sLayerName
                        End With
                        'promote layer table for write
                        If acLyrTbl.IsWriteEnabled = False Then acLyrTbl.UpgradeOpen()
                        '' Append the new layer to the Layer table and the transaction
                        acLyrTbl.Add(AcLyrTblRec)
                        acTrans.AddNewlyCreatedDBObject(AcLyrTblRec, True)
                        acTrans.Commit()
                    End Using
                    Return sLayerName
                Catch ex As Exception
                    Return "0"
                    MessageBox.Show(ex.Message)
                    Exit Function
                End Try
            End If
        End Using
    End Function
    Public Function LayerExists(ByVal LayerName As String) As Boolean

        'checks to see if a layer exists in the current drawing
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        Dim result As Boolean = False

        Try
            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim lyrTbl As LayerTable
                lyrTbl = acTrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)
                If lyrTbl.Has(LayerName) Then result = True
            End Using
        Catch ex As Exception
            MessageBox.Show("Error." & vbLf & ex.Message)
            Return Nothing
            Exit Function
        End Try

        Return result

    End Function

    Friend Sub FormLayersXML(lyrs As AcdLayers)
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        If Not curDwg.IsNamedDrawing Then
            MessageBox.Show("DWG must be saved before exporting layers.  Save file and try again.")
            Exit Sub
        End If
        Dim ed As Editor = curDwg.Editor
        Dim dwgDB As Database = curDwg.Database
        Dim dwgName As String = curDwg.Name
        Dim dwgBase As String = Path.GetFileNameWithoutExtension(dwgName)
        Dim folderName As String = Path.GetDirectoryName(dwgName)
        Dim baseName As String = folderName & "\" & dwgBase & "-"
        Dim xmlNameStr As String = baseName & lyrs.LayerSetName & ".xml"
        Dim fi As New FileInfo(xmlNameStr)
        If fi.Exists Then Kill(xmlNameStr)
        Dim sets As New XmlWriterSettings
        With sets
            .ConformanceLevel = ConformanceLevel.Document
            .Encoding = Encoding.UTF8
            .Indent = True
            .NewLineOnAttributes = True
            .IndentChars = "   "
        End With
        Using xw As XmlWriter = XmlWriter.Create(xmlNameStr, sets)
            'Dim xsets As New XmlSerializerNamespaces
            'xsets.Add("xs", "http://tempuri.org/XMLSchema1.xsd")
            Dim xmlS As New XmlSerializer(GetType(AcdLayers), "http://tempuri.org/XMLSchema1.xsd")
            xmlS.Serialize(xw, lyrs)
            xw.Flush()
        End Using
        'Dim pretty As String
        '    Using sr As New StreamReader(xmlNameStr)
        '        Dim ugly As String = sr.ReadToEnd()
        '        pretty = PrettyXml(ugly, sets)
        '    End Using
    End Sub

    Public Function GetEntitiesOnLayer(ByVal layerName As String) As ObjectIdCollection

        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim tvs As TypedValue() = New TypedValue(0) {New TypedValue(CInt(DxfCode.LayerName), layerName)}
        Dim sf As New SelectionFilter(tvs)
        Dim psr As PromptSelectionResult = ed.SelectAll(sf)

        If psr.Status = PromptStatus.OK Then
            Dim objIdset As New ObjectIdCollection(psr.Value.GetObjectIds())
            If objIdset.Count > 0 Then
                Return objIdset
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function

    Public Sub MergeThenDeleteLayer(lName As String, mergeLayer As String, removeLayer As Boolean)

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor
        'Dim lyrIDcoll As New Dictionary(Of String, ObjectId)

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()
            Dim acLyrTbl As LayerTable = acTrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)

            If acLyrTbl.Has(lName) Then

                'Dim blkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                'Dim mdlspace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                Dim purgeColl As New ObjectIdCollection

                Dim mergeList As ObjectIdCollection
                mergeList = GetEntitiesOnLayer(lName)

                Dim uR As Boolean

                If mergeList IsNot Nothing AndAlso mergeList.Count > 0 Then
                    Dim ents As Long = mergeList.Count

                    If ents > 0 Then
                        Dim pkeyO As New PromptKeywordOptions(vbLf & ents.ToString & " entities on layer " & lName & " will be moved to layer " & mergeLayer & ".  Continue?")
                        With pkeyO
                            .Keywords.Add("Y")
                            .Keywords.Add("N")
                            .AppendKeywordsToMessage = True
                        End With

                        Dim pkeyR As PromptResult = ed.GetKeywords(pkeyO)

                        If pkeyR.Status = PromptStatus.OK Then
                            If pkeyR.StringResult = "Y" Then
                                uR = True
                            Else
                                uR = False
                            End If
                        Else
                            uR = False
                        End If
                    End If

                    If uR = True Then
                        For k As Integer = 0 To mergeList.Count - 1
                            Dim myEnt As Entity = CType(acTrans.GetObject(mergeList(k), OpenMode.ForWrite), Entity)
                            If myEnt IsNot Nothing Then
                                myEnt.Layer = mergeLayer
                            End If
                        Next
                    Else
                        Exit Sub
                    End If
                End If

                If removeLayer Then
                    If acLyrTbl.IsWriteEnabled = False Then acLyrTbl.UpgradeOpen()
                    Dim lyrRec As LayerTableRecord = TryCast(acTrans.GetObject(acLyrTbl(lName), OpenMode.ForWrite), LayerTableRecord)
                    If lyrRec IsNot Nothing Then
                        purgeColl.Add(acLyrTbl(lName))
                        dwgDB.Purge(purgeColl)
                        Try
                            lyrRec.Erase(True)
                        Catch ex As Exception
                            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Error:" & vbLf & ex.Message)
                        End Try
                    End If
                End If
            End If
            acTrans.Commit()
        End Using
    End Sub
    Public Sub DeleteMyLayer(lName As String)

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor
        'Dim lyrIDcoll As New Dictionary(Of String, ObjectId)

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()
            Dim acLyrTbl As LayerTable = acTrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)

            If acLyrTbl.Has(lName) Then

                Dim DelList As ObjectIdCollection
                DelList = GetEntitiesOnLayer(lName)
                Dim uR As Boolean

                If DelList IsNot Nothing AndAlso DelList.Count > 0 Then
                    Dim ents As Long = DelList.Count
                    If ents > 0 Then
                        Dim pkeyO As New PromptKeywordOptions(vbLf & "Drawing contains " & ents.ToString & " entities on layer " & lName & ". Coninue deleting layer?")
                        With pkeyO
                            .Keywords.Add("Y")
                            .Keywords.Add("N")
                            .AppendKeywordsToMessage = True
                        End With

                        Dim pkeyR As PromptResult = ed.GetKeywords(pkeyO)

                        If pkeyR.Status = PromptStatus.OK Then
                            If pkeyR.StringResult = "Y" Then
                                uR = True
                            Else
                                uR = False
                            End If
                        Else
                            uR = False
                        End If
                    End If

                    If uR Then
                        For k As Integer = 0 To DelList.Count - 1
                            Dim myEnt As Entity = CType(acTrans.GetObject(DelList(k), OpenMode.ForWrite), Entity)
                            If myEnt IsNot Nothing Then
                                myEnt.Erase(True)
                            End If
                        Next
                    Else
                        Exit Sub
                    End If
                End If
            End If

            If acLyrTbl.IsWriteEnabled = False Then acLyrTbl.UpgradeOpen()

            Dim lyrRec As LayerTableRecord = TryCast(acTrans.GetObject(acLyrTbl(lName), OpenMode.ForWrite), LayerTableRecord)
            Dim purgeColl As New ObjectIdCollection
            If lyrRec IsNot Nothing Then
                purgeColl.Add(acLyrTbl(lName))
                dwgDB.Purge(purgeColl)
                Try
                    lyrRec.Erase(True)
                Catch ex As Exception
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Error:" & vbLf & ex.Message)
                End Try
            End If
            acTrans.Commit()
        End Using
    End Sub

    Public Function ExistLyrSet(ByVal LayName As String) As Boolean
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()
            Dim lyrTbl As LayerTable = acTrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)
            If lyrTbl.Has(LayName) Then
                dwgDB.Clayer = lyrTbl(LayName)
            Else
                Return False
                acTrans.Abort()
                Exit Function
            End If
            acTrans.Commit()
        End Using
        Return True
    End Function

    Public Sub NormalPstyleToLayer(ByVal layID As ObjectId)

        Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = acDwg.Database
        If dwgDB.PlotStyleMode = False Then
            'Using locker As DocumentLock = acDwg.LockDocument
            Using trans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim dict As DictionaryWithDefaultDictionary = trans.GetObject(dwgDB.PlotStyleNameDictionaryId, OpenMode.ForRead)
                Dim ltr As Entity = trans.GetObject(layID, OpenMode.ForWrite)
                ltr.PlotStyleName = "Normal"
                ltr.PlotStyleNameId = dict.Item("Normal")
                'Autodesk.AutoCAD.DatabaseServices.Entity.PlotStyleNameId As Autodesk.AutoCAD.DatabaseServices.PlotStyleDescriptor
                trans.Commit()
            End Using
            'End Using
        End If
    End Sub

    Public Sub SetEntityPlotStyle(ByVal entID As ObjectId, PstyleName As String)

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        If dwgDB.PlotStyleMode = False Then
            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim dict As DictionaryWithDefaultDictionary = acTrans.GetObject(dwgDB.PlotStyleNameDictionaryId, OpenMode.ForRead)
                Dim ent As Entity = TryCast(acTrans.GetObject(entID, OpenMode.ForWrite), Entity)
                If ent IsNot Nothing Then
                    ent.PlotStyleName = PstyleName
                End If
                acTrans.Commit()
            End Using
        End If
    End Sub


    Public Sub SetEntityPlotStyle(ByVal entID As ObjectId, PstyleName As String, acTrans As Transaction)

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        If dwgDB.PlotStyleMode = False Then
            Dim dict As DictionaryWithDefaultDictionary = acTrans.GetObject(dwgDB.PlotStyleNameDictionaryId, OpenMode.ForRead)
            Dim ent As Entity = TryCast(acTrans.GetObject(entID, OpenMode.ForWrite), Entity)
            If ent IsNot Nothing Then
                ent.PlotStyleName = PstyleName
                'ent.PlotStyleNameId = dict.Item(PstyleName)
            End If
            acTrans.Commit()
        End If
    End Sub



End Module


Public Module Blocks

    Public Function PickBlocks() As Collection

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor

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
                .TopLabel.Text = "Select blocks"
                .Text = "BLock Picker"
                For Each blkNm As String In bList.Keys
                    .BxList.Items.Add(blkNm)
                Next
            End With

            bPicker.ShowDialog()

            Dim blkColl As Collection

            If bPicker.DialogResult = DialogResult.Cancel Then
                Return Nothing
                Exit Function
            Else
                blkColl = bPicker.PickCol
            End If

            acTrans.Commit()
            Return blkColl
        End Using

    End Function





    Public Function BLockEntsToLayerZero(brefID As ObjectId) As ObjectIdCollection

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor

        Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction()
            Dim blkTbl As BlockTable = actrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
            Dim ltbl As LayerTable = actrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)

            Dim bref As BlockReference = actrans.GetObject(brefID, OpenMode.ForRead)
            Dim btrID As ObjectId = bref.BlockId

            Dim btr As BlockTableRecord = actrans.GetObject(btrID, OpenMode.ForWrite)
            Dim idCol As New ObjectIdCollection

            'Dim i As Integer = 0
            For Each obID As ObjectId In btr
                Dim dObj As DBObject = actrans.GetObject(obID, OpenMode.ForRead)
                If TypeOf dObj IsNot BlockReference Then
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
                        End If
                    End If
                Else
                    idCol.Add(obID)
                End If
            Next
            Return idCol
            actrans.Commit()
        End Using
    End Function

    Public Sub ClearBlk(bName As String)
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        Try
            If Not String.IsNullOrEmpty(bName) Then
                If BlkExists(bName) Then
                    Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                        Dim blkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                        Dim myBTR As BlockTableRecord = acTrans.GetObject(blkTbl(bName), OpenMode.ForWrite)
                        For Each id As ObjectId In myBTR
                            Dim dbobj As DBObject = acTrans.GetObject(id, OpenMode.ForWrite)
                            If TypeOf dbobj Is Entity Then
                                Dim ent As Entity = TryCast(dbobj, Entity)
                                If ent IsNot Nothing Then ent.Erase()
                            End If
                        Next
                        acTrans.Commit()
                    End Using
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Function ClearBlk(bName As String, acTrans As Transaction) As ObjectId
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        Try
            If Not String.IsNullOrEmpty(bName) Then
                If BlkExists(bName) Then
                    Dim blkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
                    Dim myBTR As BlockTableRecord = acTrans.GetObject(blkTbl(bName), OpenMode.ForWrite)
                    For Each id As ObjectId In myBTR
                        Dim dbobj As DBObject = acTrans.GetObject(id, OpenMode.ForWrite)
                        If TypeOf dbobj Is Entity Then
                            Dim ent As Entity = TryCast(dbobj, Entity)
                            If ent IsNot Nothing Then ent.Erase()
                        End If
                    Next
                    Return myBTR.ObjectId
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Private Function BTRfromRef(bRefId As ObjectId) As ObjectId
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim btrObjId As ObjectId

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
            Dim bName As String
            Dim bRef As BlockReference = acTrans.GetObject(bRefId, OpenMode.ForRead)
            If bRef.IsDynamicBlock Then
                Dim appBTR As BlockTableRecord = acTrans.GetObject(bRef.DynamicBlockTableRecord, OpenMode.ForRead)
                bName = appBTR.Name
            Else
                bName = bRef.Name
            End If
            Dim blkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
            btrObjId = blkTbl(bName)
            acTrans.Commit()
        End Using
        Return btrObjId
    End Function

    Public Function BlkExists(bNameStr As String) As Boolean

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim result As Boolean = False

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()
            'open the block table for read
            Dim acBlkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)

            'test for block name and return true if it is found
            If acBlkTbl.Has(bNameStr) Then result = True
        End Using
        Return result

    End Function

    Public Function GetBTRObjID(blkName As String) As ObjectId

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        Dim blkObjId As ObjectId

        Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
            'get the pd BTR from the pd name
            Dim blkTbl As BlockTable = actrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
            If blkTbl.Has(blkName) Then
                blkObjId = blkTbl(blkName)
            Else
                blkObjId = ObjectId.Null
            End If
            Return blkObjId
        End Using

    End Function
    Public Sub ImportBlk(bPath As String, bName As String)
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor

        Try
            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Using tmpDb As New Database(False, True)
                    tmpDb.ReadDwgFile(bPath, IO.FileShare.Read, True, Nothing)
                    DwgDB.Insert(bName, tmpDb, True)
                End Using
                acTrans.Commit()
            End Using

        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            ed.WriteMessage(vbLf & "Error in sub ImportBlk: " & ex.Message)
        End Try

    End Sub
    Public Sub DelPurgeBlock(bName As String)
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Try
            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                If BlkExists(bName) Then
                    Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    Dim delRefs As Boolean = True
                    Dim refsExist As Boolean = False
                    Dim refList As New ObjectIdCollection
                    For Each btrId As ObjectId In blkTbl
                        Dim btr As BlockTableRecord = acTrans.GetObject(btrId, OpenMode.ForRead)
                        For Each entID As ObjectId In btr
                            Dim ob As DBObject = acTrans.GetObject(entID, OpenMode.ForRead)
                            If TypeOf ob Is BlockReference Then
                                Dim bRef As BlockReference = TryCast(ob, BlockReference)
                                If bRef.Name = bName Then
                                    If Not refsExist Then
                                        Dim msgStr As String = "There are existing block references for " & bName & " in this drawing.  Erase references and purge block?"
                                        Dim dr As DialogResult = MessageBox.Show(msgStr, "References Exist", MessageBoxButtons.YesNo)
                                        If dr = DialogResult.Yes Then
                                            refsExist = True
                                            delRefs = True
                                            refList.Add(entID)
                                        Else
                                            refsExist = True
                                            delRefs = False
                                        End If
                                    Else
                                        refList.Add(entID)
                                    End If
                                End If
                            End If
                        Next
                    Next

                    If refsExist Then
                        If delRefs Then
                            For Each id As ObjectId In refList
                                Dim tempRef As Entity = acTrans.GetObject(id, OpenMode.ForWrite)
                                tempRef.Erase()
                            Next
                        Else
                            acTrans.Abort()
                            Exit Sub
                        End If
                    End If

                    PurgeBlk(bName, acTrans)
                End If
                acTrans.Commit()
            End Using

        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            MessageBox.Show(ex.Message.ToString)
        End Try

    End Sub

    Public Sub RenameBlock(bName As String, newName As String)
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        If BlkExists(bName) Then
            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = acTrans.GetObject(blkTbl(bName), OpenMode.ForWrite)
                btr.Name = newName
                acTrans.Commit()
            End Using
        End If
    End Sub

    Public Sub RenameBlock(bName As String, newName As String, acTrans As Transaction)
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        If BlkExists(bName) Then
            Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(blkTbl(bName), OpenMode.ForWrite)
            btr.Name = newName
        End If
    End Sub

    Public Sub PurgeBlk(bName As String)
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Try
            Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                If BlkExists(bName) Then
                    Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                    'Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim blkId As ObjectId = blkTbl(bName)
                    Dim blkCol As New ObjectIdCollection
                    blkCol.Add(blkId)
                    DwgDB.Purge(blkCol)
                End If
                acTrans.Commit()
            End Using
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub

    Public Sub PurgeBlk(bName As String, acTrans As Transaction)
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Try
            If BlkExists(bName) Then
                Dim blkTbl As BlockTable = acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead)
                'Dim mdlSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim blkId As ObjectId = blkTbl(bName)
                Dim blkCol As New ObjectIdCollection
                blkCol.Add(blkId)
                DwgDB.Purge(blkCol)
            End If
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub
    Public Sub InsertDwgBlock(fName As String, safeBname As String)

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        Using newDwgDb As New Database
            If Not BlkExists(safeBname) Then
                newDwgDb.ReadDwgFile(fName, IO.FileShare.Read, True, Nothing)
                dwgDB.Insert(safeBname, newDwgDb, True)
            End If
        End Using

    End Sub

    Public Function InsertDwg(fName As String, safebname As String) As ObjectId

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim tempID As ObjectId

        Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
            Using newDwgDb As New Database(False, True)
                Dim blkTbl As BlockTable = actrans.GetObject(dwgDB.BlockTableId, OpenMode.ForWrite, False)
                If Not blkTbl.Has(safebname) Then
                    newDwgDb.ReadDwgFile(fName, IO.FileShare.Read, True, Nothing)
                    tempID = dwgDB.Insert(safebname, newDwgDb, True)
                Else
                    tempID = blkTbl(safebname)
                End If
            End Using
            actrans.Commit()
        End Using

        Return tempID

    End Function

    Public Function InsertDwg(fName As String, safebname As String, tr As Transaction) As ObjectId

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim tempID As ObjectId

        Using newDwgDb As New Database(False, True)
            Dim blkTbl As BlockTable = tr.GetObject(dwgDB.BlockTableId, OpenMode.ForWrite, False)
            If Not blkTbl.Has(safebname) Then
                newDwgDb.ReadDwgFile(fName, IO.FileShare.Read, True, Nothing)
                tempID = dwgDB.Insert(safebname, newDwgDb, True)
            Else
                tempID = blkTbl(safebname)
            End If
        End Using

        Return tempID

    End Function

    Public Function SelBlock() As String

        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = CurDwg.Editor
        Dim blkobjID As ObjectId

        Dim pEO As New PromptEntityOptions("Select an inserted block")
        With pEO
            .AllowNone = False
            .AllowObjectOnLockedLayer = True
            .SetRejectMessage("Select a block reference")
            .AddAllowedClass(GetType(BlockReference), True)
        End With

        Dim pRes As PromptEntityResult = ed.GetEntity(pEO)

        If pRes.Status = PromptStatus.OK Then
            blkobjID = pRes.ObjectId
        Else
            Return "Cancel"
            Exit Function
        End If

        Dim blkRef As BlockReference
        Dim blkName As String

        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
            blkRef = acTrans.GetObject(blkobjID, OpenMode.ForRead)
            blkName = blkRef.Name
            acTrans.Commit()
        End Using

        Return blkName

    End Function

    Public Sub ChangeBase(ByVal sourcefilename As String)
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = curDwg.Editor

        'create a new autocad drawing database to be used out-of-process
        Dim sourceDb As New Database(False, True)

        'read dwg file into database
        Try
            sourceDb.ReadDwgFile(sourcefilename, System.IO.FileShare.Read, False, "")
        Catch __unusedException1__ As System.Exception
            MessageBox.Show(vbLf & "Unable to read drawing file.")
            Return
        End Try

        'Set a point at the origin and compare to base point of block file
        Dim compPoint As New Point3d(0, 0, 0)
        Dim curInsBase As Point3d = sourceDb.Insbase

        If curInsBase <> compPoint Then
            'if not the same, ask what to do
            Dim prkOpts As New PromptKeywordOptions("")
            With prkOpts
                .Message = "Change insertion base of " & Path.GetFileName(sourcefilename) & " to (0,0,0)?"
                .Keywords.Add("Yes")
                .Keywords.Add("No")
                .AppendKeywordsToMessage = True
                .AllowArbitraryInput = False
                .AllowNone = False
            End With

            Dim prkres As PromptResult = ed.GetKeywords(prkOpts)
            Dim kw As String
            If Not prkres.Status = PromptStatus.OK Then
                Exit Sub
            Else
                kw = prkres.StringResult

                'Dim uResp As Integer = MsgBox("Insertion base of " & vbLf & sourcefilename & vbLf & " is " & curInsBase.ToString _
                '& vbLf & " change it to (0, 0, 0)?", vbYesNo)
                'If uResp = vbNo Then Exit Sub
                'If uResp = vbYes Then
                sourceDb.Insbase = compPoint
                Try
                    sourceDb.SaveAs(sourcefilename, True, Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1027, sourceDb.SecurityParameters)
                Catch ex As Exception
                    MessageBox.Show("Unable to save drawing file." & vbLf & ex.Message)
                End Try
            End If

        End If
        sourceDb.Dispose()

    End Sub

    Public Function GetBlkName(bNameStr As String) As String

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        Dim bName As String = bNameStr

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()
            'open the block table for read
            Dim acBlkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
            Dim i As Integer = 0

            Dim tempBName As String = bName
            Dim newBName As String = tempBName & CStr(i)
            bName = newBName

            'test for block name and change it if it is found
            If acBlkTbl.Has(newBName) Then
                Do While acBlkTbl.Has(newBName)
                    newBName = tempBName & CStr(i)
                    i += 1
                Loop
                bName = newBName
            End If
            acTrans.Commit()

        End Using

        Return bName

    End Function

    Public Function GetNewFileName(fName As String, Optional filetype As String = "") As String

        Dim fi As New FileInfo(fName)
        Dim di As New DirectoryInfo(fi.DirectoryName)
        Dim myfiles() As FileInfo = di.GetFiles
        Dim stripName As String = Path.ChangeExtension(fi.Name, vbNullString)
        Dim outName As String

        If System.IO.File.Exists(fName) Then
            Dim newfile As String
            If String.IsNullOrEmpty(filetype) Then
                newfile = di.FullName & "\" & stripName & "-2.png"
            Else
                newfile = di.FullName & "\" & stripName & "." & filetype
            End If

            Dim j As Integer = 3

            If String.IsNullOrEmpty(filetype) Then
                Do While System.IO.File.Exists(newfile)
                    newfile = di.FullName & "\" & stripName & "-" & j.ToString & ".png"
                    j += 1
                Loop
            Else
                Do While System.IO.File.Exists(newfile)
                    newfile = di.FullName & "\" & stripName & "-" & j.ToString & "." & filetype
                    j += 1
                Loop
            End If
            outName = newfile
        Else
            outName = fName
        End If

        Return outName

    End Function

    Public Function SelBlockRef() As ObjectId
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor
        Dim blkobjID As ObjectId

        Dim pEO As New PromptEntityOptions(vbLf & "Select an inserted block")
        With pEO
            .AllowNone = False
            .AllowObjectOnLockedLayer = True
            .SetRejectMessage(vbLf & "Invalid selection.  Select a block reference.")
            .AddAllowedClass(GetType(BlockReference), True)
        End With

        Dim pRes As PromptEntityResult = ed.GetEntity(pEO)

        If pRes.Status = PromptStatus.OK Then
            blkobjID = pRes.ObjectId
        Else
            Return ObjectId.Null
            Exit Function
        End If

        Return blkobjID
    End Function

    Public Function SubBlockExists(blkName As String, sbName As String) As Boolean
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor
        Dim ret As Boolean = False

        Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
            Dim blkTbl As BlockTable = actrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
            If blkTbl.Has(blkName) Then
                Dim bTR As BlockTableRecord = actrans.GetObject(blkTbl(blkName), OpenMode.ForRead)
                Dim appBname As String
                For Each objID In bTR
                    Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is BlockReference Then
                        Dim idref As BlockReference = CType(dbObj, BlockReference)
                        If idref.IsDynamicBlock Then
                            Dim appBTR As BlockTableRecord = actrans.GetObject(idref.DynamicBlockTableRecord, OpenMode.ForRead)
                            appBname = appBTR.Name
                        Else
                            appBname = idref.Name
                        End If
                        If appBname = sbName Then
                            ret = True
                            Exit For
                        End If
                    End If
                Next
            End If
            actrans.Commit()
        End Using
        Return ret

    End Function

    Public Function ChangeVisState(bkRefID As ObjectId, visState As String, PropName As String) As Boolean

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor
        Dim retValue As Boolean

        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
            Dim bRef As BlockReference = TryCast(acTrans.GetObject(bkRefID, OpenMode.ForWrite), BlockReference)
            If bRef.IsDynamicBlock Then
                For Each prop As DynamicBlockReferenceProperty In bRef.DynamicBlockReferencePropertyCollection
                    If prop.PropertyName.ToUpper() = PropName.ToUpper() Then
                        Try
                            prop.Value = visState
                            retValue = True
                            Exit For
                        Catch ex As Exception
                            retValue = False
                        End Try
                    End If
                Next
            Else
                retValue = False
            End If
            acTrans.Commit()
        End Using

        Return retValue

    End Function

    Public Function ChangeVisState(acTrans As Transaction, bkRefID As ObjectId, visState As String, PropName As String) As Boolean

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor
        Dim retValue As Boolean

        Dim bRef As BlockReference = TryCast(acTrans.GetObject(bkRefID, OpenMode.ForWrite), BlockReference)
        If bRef.IsDynamicBlock Then
            For Each prop As DynamicBlockReferenceProperty In bRef.DynamicBlockReferencePropertyCollection
                If prop.PropertyName.ToUpper() = PropName.ToUpper() Then
                    Try
                        prop.Value = visState
                        retValue = True
                        Exit For
                    Catch ex As Exception
                        retValue = False
                        Exit For
                    End Try
                End If
            Next
        Else
            retValue = False
        End If

        Return retValue

    End Function

End Module


Public Module FileUtilities

    Public Function ReturnValidFileName(Value() As Char) As String
        Dim TempStringBuilder As New StringBuilder
        For i = 0 To Value.Length - 1
            If Not ("^\.@~".Contains(Value(i))) And Not Value(i) = ChrW(34) Then
                TempStringBuilder.Append(Value(i))
            End If
        Next
        Return TempStringBuilder.ToString
    End Function


    Public Function GetMyXMLFileName() As String

        'function for getting filename of xml data file.
        Try
            Dim fName As String = ""

            'declare a new open file dialog
            Dim fDialog As New OpenFileDialog()
            With fDialog
                .Reset()
                .Filter = "xml (*.xml)|*.xml|" _
                & "Text Files (*.txt)|*.txt"
                .FilterIndex = 1
                '.InitialDirectory = "\\EESSERVER\datadisk\LITIGATION\Active Cases"
                .Title = "Select properly formatted XML file"
                .CheckFileExists = True
            End With

            'get the dialog result or return nothing
            Dim xmlResult As DialogResult = fDialog.ShowDialog()

            If xmlResult = DialogResult.Cancel Then
                MessageBox.Show("Error. XML file not selected")
                Return ""
                Exit Function

            ElseIf xmlResult = DialogResult.OK Then
                fName = fDialog.FileName
            Else
                MessageBox.Show("Error. XML file not selected")
                Return ""
                Exit Function
            End If

skipit:
            'return the name and path of the file
            Return fName

        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "Error selecting XML file.")
            Return ""
            Exit Function
        End Try

    End Function

    Public Function GetFileNameSameAsDWG(Optional filetype As String = "") As String

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgName As String = curDwg.Name
        Dim dwgBase As String = Path.GetFileNameWithoutExtension(dwgName)
        Dim folderName As String = Path.GetDirectoryName(dwgName)
        Dim baseName As String = folderName & "\" & dwgBase
        Dim fileNameStr As String
        If Not String.IsNullOrEmpty(filetype) Then
            fileNameStr = baseName & filetype
        Else
            fileNameStr = baseName & ".dwg"
        End If

        Return fileNameStr

    End Function


    Public Function GetASaveFileName(ext As String) As String

        'function for getting filename
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgName As String = curDwg.Name
        Dim folderName As String = Path.GetDirectoryName(dwgName)

        Try
            Dim fName As String = ""

            'declare a new open file dialog
            Dim fDialog As New SaveFileDialog()
            Dim filterString As String = ext & " Files " & "(*." & ext & ")|*." & ext & "|All Files (*.*)|*.*"

            With fDialog
                .Reset()
                .Filter = filterString
                .InitialDirectory = folderName
                .Title = "Enter file name"
                .CheckFileExists = True
            End With

            'get the dialog result or return nothing
            Dim fileResult As DialogResult = fDialog.ShowDialog()

            If fileResult = DialogResult.Cancel Then
                MessageBox.Show("Error. File not selected")
                Return Nothing
                Exit Function
            ElseIf fileResult = DialogResult.OK Then
                fName = fDialog.FileName
            Else
                MessageBox.Show("Error. File name not provided.")
                Return ""
                Exit Function
            End If

skipit:
            'return the name and path of the file
            Return fName

        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "Error selecting file.")
            Return ""
            Exit Function
        End Try

    End Function

    Public Function GetMyCSVFileName() As String

        'function for getting filename of csv data file
        Try
            Dim fName As String = ""

            'declare a new open file dialog
            Dim fDialog As New OpenFileDialog()
            With fDialog
                .Reset()
                .Filter = "csv (*.csv)|*.csv|" _
                & "Text Files (*.txt)|*.txt"
                '.InitialDirectory = "H:\VS Repos"
                '.InitialDirectory = "\\EESSERVER\datadisk\LITIGATION\Active Cases"
                .Title = "Select properly formatted CSV file"
                .CheckFileExists = True
            End With

            'get the dialog result or return nothing
            Dim xmlResult As DialogResult = fDialog.ShowDialog()

            If xmlResult = DialogResult.Cancel Then
                MessageBox.Show("Error. File not selected")
                Return Nothing
                Exit Function

            ElseIf xmlResult = DialogResult.OK Then
                fName = fDialog.FileName
            Else
                MessageBox.Show("Error. File not selected")
                Return ""
                Exit Function
            End If

skipit:
            'return the name and path of the file
            Return fName

        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "Error selecting file.")
            Return ""
            Exit Function
        End Try

    End Function

    Public Function GetMyFolderName(Optional sfName As String = "") As String

        'function for getting folder name as string

        Try
            Dim fName As String = ""
            Dim retval As String

            'declare a new open file dialog
            Dim fDialog As New FolderBrowserDialog()
            With fDialog
                .Reset()
                .ShowNewFolderButton = True
                If Not sfName = "" Then .SelectedPath = sfName
            End With

            'get the dialog result or return nothing
            Dim fldrResult As DialogResult = fDialog.ShowDialog()

            If fldrResult = DialogResult.Cancel Then
                retval = ""
            ElseIf fldrResult = DialogResult.OK Then
                fName = fDialog.SelectedPath
                retval = fName
            Else
                MessageBox.Show("Error. Folder not selected")
                retval = ""
            End If

            'return the name and path of the file
            Return retval

        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "Error selecting folder.")
            Return ""
            Exit Function
        End Try

    End Function


    Public Function GetMyXSDFileName() As String

        'function for getting filename of schema file.
        Try
            Dim fName As String = ""

            'declare a new open file dialog
            Dim fDialog As New OpenFileDialog()
            With fDialog
                .Reset()
                .Filter = "Schema Files (*.xsd)|*.xsd|" _
                & "Text Files (*.txt)|*.txt"
                .FilterIndex = 1
                '.InitialDirectory = "\\EESSERVER\datadisk\Civil\ENGINEERING\Traffic Studies\"
                .Title = "Select properly formatted XSD file"
                .CheckFileExists = True
            End With

            'get the dialog result or return nothing
            Dim xsdResult As DialogResult = fDialog.ShowDialog()

            If xsdResult = DialogResult.Cancel Then
                MessageBox.Show("Error. XSD file not selected")
                Return ""
                Exit Function

            ElseIf xsdResult = DialogResult.OK Then
                fName = fDialog.FileName
            Else
                MessageBox.Show("Error. XSD file not selected")
                Return ""
                Exit Function
            End If

skipit:
            'return the name and path of the file
            Return fName

        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "Error selecting XSD file.")
            Return ""
            Exit Function
        End Try

    End Function

    Function PrettyXml(ByVal str As String, ByVal Optional settings As XmlWriterSettings = Nothing) As String
        If String.IsNullOrWhiteSpace(str) Then Return str

        Try
            Dim xml As New XmlDocument
            xml.LoadXml(str)
            Dim sets As New XmlWriterSettings
            If settings Is Nothing Then
                settings = New XmlWriterSettings
                With settings
                    .Indent = True
                    .NewLineOnAttributes = True
                    .Encoding = Encoding.UTF8
                    .IndentChars = "    "
                End With
                sets = settings
            Else
                sets = settings
            End If
            Using sw As New StringWriter
                Debug.Print(sw.Encoding.ToString)
                Using textWriter As XmlWriter = XmlWriter.Create(sw, sets)
                    xml.Save(textWriter)
                End Using
                sw.Flush()
                Return sw.ToString()
            End Using
        Catch
        End Try

        Return str

    End Function

    Public Function AllDwgFilesInFolder(fpath As String, recurseDir As Boolean) As FileInfo()

        Dim di As New DirectoryInfo(fpath)
        Dim dFiles() As FileInfo

        If Not di.Exists Then
            Return Nothing
            Exit Function
        End If

        If recurseDir Then
            dFiles = di.GetFiles("*.dwg", SearchOption.AllDirectories)
        Else
            dFiles = di.GetFiles("*.dwg", SearchOption.TopDirectoryOnly)
        End If

        Return dFiles

    End Function



End Module



Public Module TextDIms

    Public Function NewDimStyle(dimName As String) As ObjectId
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = curDwg.Database

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim DimTabb As DimStyleTable = CType(trans.GetObject(db.DimStyleTableId, OpenMode.ForRead), DimStyleTable)
            Dim dimId As ObjectId = ObjectId.Null

            If Not DimTabb.Has(dimName) Then
                DimTabb.UpgradeOpen()
                Dim newRecord As New DimStyleTableRecord() With {.Name = dimName}

                dimId = DimTabb.Add(newRecord)
                trans.AddNewlyCreatedDBObject(newRecord, True)
            Else
                dimId = DimTabb(dimName)
            End If

            Dim DimTabbRecaord As DimStyleTableRecord = CType(trans.GetObject(dimId, OpenMode.ForRead), DimStyleTableRecord)

            If DimTabbRecaord.ObjectId <> db.Dimstyle Then
                db.Dimstyle = DimTabbRecaord.ObjectId
                db.SetDimstyleData(DimTabbRecaord)
            End If

            Return dimId
            trans.Commit()

        End Using
    End Function

    Public Function AddMyTstyle(ByVal tStyleName As String) As String

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor
        Dim returnStr As String = ""

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()

            Dim hasStyle As Boolean = False
            'open the textstyle table for read
            Dim tsTbl As SymbolTable = acTrans.GetObject(dwgDB.TextStyleTableId, OpenMode.ForRead)
            Dim curTextStyle As TextStyleTableRecord = acTrans.GetObject(dwgDB.Textstyle, OpenMode.ForRead)
            'MsgBox("typeface: " & vbLf & curTextStyle.Font.TypeFace.ToString & vbLf _
            '       & "pitch and family: " & curTextStyle.Font.PitchAndFamily.ToString & vbLf _
            '       & "CharacterSet: " & curTextStyle.Font.CharacterSet.ToString)

            Dim prmptOpts1 As New PromptKeywordOptions("Use the current (D)rawing textStyle (default), (P)ick text entity for textstyle, (S)elect existing textstyle, or (C)reate new textstyle from system font?")
            With prmptOpts1
                .Keywords.Add("D")
                .Keywords.Add("P")
                .Keywords.Add("S")
                .Keywords.Add("C")
                .AppendKeywordsToMessage = True
                .AllowNone = True
            End With

            Dim PrmptRes1 As PromptResult = ed.GetKeywords(prmptOpts1)

            Dim uTResp As String

            If PrmptRes1.Status = PromptStatus.OK Then
                uTResp = PrmptRes1.StringResult
            ElseIf PrmptRes1.Status = PromptStatus.None Then
                uTResp = "D"
            ElseIf PrmptRes1.Status = PromptStatus.Cancel Then
                Return Nothing
                Exit Function
            Else
                uTResp = "D"
            End If

            If uTResp = "D" Then
                tStyleName = curTextStyle.Name

            ElseIf uTResp = "P" Then
                Dim prmptopts2 As New PromptEntityOptions(vbLf & "Select text entity for textstyle.")
                With prmptopts2
                    .SetRejectMessage(vbLf & "Not a Dtext or Mtext entity.  Try again.")
                    .AddAllowedClass(GetType(MText), False)
                    .AddAllowedClass(GetType(DBText), False)
                    .SetRejectMessage(vbLf & "Not a Dtext or Mtext entity.  Try again.")
                    .AllowNone = True
                End With

                Dim prmptres2 As PromptEntityResult = ed.GetEntity(prmptopts2)
                Dim tObjId As ObjectId

                If prmptres2.Status = PromptStatus.OK Then

                    tObjId = prmptres2.ObjectId

                    Dim tSTyleRec As TextStyleTableRecord
                    Dim tempObj As DBObject = acTrans.GetObject(tObjId, OpenMode.ForRead)
                    Dim tempDBText As DBText = TryCast(tempObj, DBText)
                    If tempDBText IsNot Nothing Then
                        tSTyleRec = acTrans.GetObject(tempDBText.TextStyleId, OpenMode.ForRead)
                        tStyleName = tSTyleRec.Name
                    Else
                        Dim tempMtext As MText = TryCast(tempObj, MText)
                        If tempMtext IsNot Nothing Then
                            tSTyleRec = acTrans.GetObject(tempMtext.TextStyleId, OpenMode.ForRead)
                            tStyleName = tSTyleRec.Name
                            ed.WriteMessage(vbLf & "Using textstyle: " & tStyleName)
                        End If
                    End If
                Else
                    ed.WriteMessage(vbLf & "Using current textstyle: " & curTextStyle.Name)
                    tStyleName = curTextStyle.Name
                End If

            ElseIf uTResp = "S" Then
                Dim tempStyleName As String = PickStyleName()
                If String.IsNullOrEmpty(tempStyleName) Then
                    tStyleName = curTextStyle.Name
                Else
                    tStyleName = tempStyleName
                End If

            ElseIf uTResp = "C" Then

                Dim pso As New PromptStringOptions(vbLf & "Input textstyle name:")
                With pso
                    .AllowSpaces = False
                    .DefaultValue = tStyleName
                End With

                Dim currentName As String = ""
tryAgain:
                Dim prmptres3 As PromptResult = ed.GetString(pso)

                If prmptres3.Status = PromptStatus.OK Then
                    Dim sName As String = prmptres3.StringResult
                    If tsTbl.Has(sName) Then
                        ed.WriteMessage(vbLf & "Textstyle " & sName & " exists. Use another name or cancel to use existing textstyle.")
                        currentName = sName
                        pso.DefaultValue = currentName
                        GoTo tryAgain
                    Else
                        Dim tempStyleName As String = TStyleCreate(sName)
                        If tempStyleName = "" Then
                            tStyleName = curTextStyle.Name
                        Else
                            tStyleName = tempStyleName
                        End If
                    End If
                Else
                    If currentName = "" Then
                        tStyleName = curTextStyle.Name
                    Else
                        tStyleName = currentName
                    End If
                End If
            Else
                tStyleName = curTextStyle.Name
            End If

            'MsgBox("textstyleName = " & tStyleName)

            acTrans.Commit()

        End Using

        Return tStyleName

    End Function

    Public Function PickSysFnt() As String

        Dim fntPicker As New FontPicker
        Dim fntName As String
        'Dim styleNm As String

        fntPicker.Text = "System Font Picker"
        fntPicker.PickerType = "sysfont"

        'Dim styleNameStr As String

        fntPicker.ShowDialog()

        If fntPicker.DialogResult = DialogResult.OK Then
            'MsgBox("Picked Font Name is " & fntPicker.FontName)
            fntName = fntPicker.FontName
        Else
            MessageBox.Show("error picking font name.")
            Return ""
            Exit Function
        End If

        Return fntName

    End Function

    Public Function TStyleCreate(tempStyleName As String, Optional sysFontName As String = "") As String

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()

            Dim tsTbl As SymbolTable = acTrans.GetObject(dwgDB.TextStyleTableId, OpenMode.ForRead)
            Dim curTextStyle As TextStyleTableRecord = acTrans.GetObject(dwgDB.Textstyle, OpenMode.ForRead)
            Dim tStyleName As String = tempStyleName

            For Each id As ObjectId In tsTbl
                Dim symbol As TextStyleTableRecord = CType(acTrans.GetObject(id, OpenMode.ForRead), TextStyleTableRecord)
                If symbol.Name = tStyleName Then
                    ed.WriteMessage("Textstyle exists.  Using existing textstyle.")
                    Return tStyleName
                    Exit Function
                End If
            Next

            Try
                If sysFontName = "" Then sysFontName = PickSysFnt()

                If sysFontName = "" Then
                    Return ""
                    Exit Function
                End If

                Dim myTS As New TextStyleTableRecord With {.Name = tStyleName}
                'myTS.Name = tStyleName
                tsTbl.UpgradeOpen()
                tsTbl.Add(myTS)

                If myTS.IsWriteEnabled = False Then myTS.UpgradeOpen()
                With myTS
                    .Annotative = AnnotativeStates.False
                    .XScale = 0.8
                    .TextSize = 0
                End With

                myTS.Font = New Autodesk.AutoCAD.GraphicsInterface.FontDescriptor(sysFontName, False, False, curTextStyle.Font.CharacterSet, curTextStyle.Font.PitchAndFamily)
                acTrans.AddNewlyCreatedDBObject(myTS, True)
                acTrans.Commit()

            Catch ex As Exception
                'Messagebox.show(ex.ErrorStatus & vbLf & ex.Message)
                MessageBox.Show("Error creating new textstyle.  Using current drawing textstyle." & vbLf & ex.ErrorStatus & vbLf & ex.Message)
                Return ""
                acTrans.Abort()
                Exit Function
            End Try
            Return tStyleName
        End Using
    End Function

    Public Function PickStyleName() As String

        Dim fntPicker As New FontPicker
        Dim styleName As String

        fntPicker.Text = "Textstyle Picker"
        fntPicker.PickerType = "textstyle"

        fntPicker.ShowDialog()

        If fntPicker.DialogResult = DialogResult.OK Then
            'Messagebox.show("Picked Font Name is " & fntPicker.FontName)
            styleName = fntPicker.FontName
        Else
            'Messagebox.show("error picking style.")
            Return ""
            Exit Function
        End If

        Return styleName

    End Function

    Public Function SetMyDimStyle(dsName As String) As ObjectId

        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        'Dim cdArrowID As ObjectId = MkCDArrow("CDArrow")

        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction

            Dim dSTbl As DimStyleTable = TryCast(acTrans.GetObject(DwgDB.DimStyleTableId, OpenMode.ForRead), DimStyleTable)
            Dim dsId As ObjectId = ObjectId.Null

            Dim blktbl As BlockTable = TryCast(acTrans.GetObject(DwgDB.BlockTableId, OpenMode.ForRead), BlockTable)

            Dim dsRec0 As ObjectId = DwgDB.Dimstyle
            Dim cDimRec As DimStyleTableRecord

            If dSTbl.Has(dsName) = False Then
                If dSTbl.IsWriteEnabled = False Then dSTbl.UpgradeOpen()
                cDimRec = New DimStyleTableRecord() With {.Name = dsName}
                dsId = dSTbl.Add(cDimRec)
                acTrans.AddNewlyCreatedDBObject(cDimRec, True)
            Else
                cDimRec = acTrans.GetObject(dSTbl(dsName), OpenMode.ForWrite)
                dsId = dSTbl(dsName)
            End If

            Dim dsCopy As DimStyleTableRecord = DwgDB.GetDimstyleData

            With cDimRec
                .Dimarcsym = 2
                .Dimse1 = True
                .Dimse2 = True
                '.Dimblk1 = GetArrowObjectId("_None")
                .Dimblk2 = ObjectId.Null
            End With

            Return cDimRec.ObjectId

            acTrans.Commit()
        End Using

    End Function

End Module

Public Module Properties
    Public Function DwgVersion(filename As String) As String
        Using reader As New StreamReader(filename)
            Select Case reader.ReadLine().Substring(0, 6)
                Case "ac1032"
                    Return "Autocad 2018"
                Case "AC1027"
                    Return "AutoCAD 2013"
                Case "AC1024"
                    Return "AutoCAD 2010"
                Case "AC1021"
                    Return "AutoCAD 2007"
                Case "AC1018"
                    Return "AutoCAD 2004"
                Case "AC1015"
                    Return "AutoCAD 2000"
                Case "AC1014"
                    Return "AutoCAD R14"
                Case Else
                    Return "Prior AutoCAD R14"
            End Select
        End Using
    End Function

    Friend Function GetTransparency(trans As Integer) As Transparency
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = curDwg.Editor
        Dim dwgDB As Database = curDwg.Database
        If trans > 0 Then
            Dim alpha As Byte = CByte(255 * ((100 - trans) / 100))
            Dim tp As New Transparency(alpha)
            Return tp
        Else
            Dim tp As New Transparency(0)
            Return tp
        End If
    End Function

    Public Function GetTransparencyAlpha(ByVal idx As Integer) As Transparency
        Dim transparencies As Dictionary(Of Integer, Byte) = TransToAlpha()
        If transparencies.ContainsKey(idx) Then
            Return New Autodesk.AutoCAD.Colors.Transparency(transparencies(idx))
        Else
            Return New Autodesk.AutoCAD.Colors.Transparency(Byte.MaxValue)
        End If
    End Function

    Friend Function TransToAlpha() As Dictionary(Of Integer, Byte)
        Dim transp As New Dictionary(Of Integer, Byte)
        transp(0) = 255
        transp(1) = 252
        transp(2) = 249
        transp(3) = 247
        transp(4) = 244
        transp(5) = 242
        transp(6) = 239
        transp(7) = 237
        transp(8) = 234
        transp(9) = 232
        transp(10) = 229
        transp(11) = 226
        transp(12) = 224
        transp(13) = 221
        transp(14) = 219
        transp(15) = 216
        transp(16) = 214
        transp(17) = 211
        transp(18) = 209
        transp(19) = 206
        transp(20) = 204
        transp(21) = 201
        transp(22) = 198
        transp(23) = 196
        transp(24) = 193
        transp(25) = 191
        transp(26) = 188
        transp(27) = 186
        transp(28) = 183
        transp(29) = 181
        transp(30) = 178
        transp(31) = 175
        transp(32) = 173
        transp(33) = 170
        transp(34) = 168
        transp(35) = 165
        transp(36) = 163
        transp(37) = 160
        transp(38) = 158
        transp(39) = 155
        transp(40) = 153
        transp(41) = 150
        transp(42) = 147
        transp(43) = 145
        transp(44) = 142
        transp(45) = 140
        transp(46) = 137
        transp(47) = 135
        transp(48) = 132
        transp(49) = 130
        transp(50) = 127
        transp(51) = 124
        transp(52) = 122
        transp(53) = 119
        transp(54) = 117
        transp(55) = 114
        transp(56) = 112
        transp(57) = 109
        transp(58) = 107
        transp(59) = 104
        transp(60) = 102
        transp(61) = 99
        transp(62) = 96
        transp(63) = 94
        transp(64) = 91
        transp(65) = 89
        transp(66) = 86
        transp(67) = 84
        transp(68) = 81
        transp(69) = 79
        transp(70) = 76
        transp(71) = 73
        transp(72) = 71
        transp(73) = 68
        transp(74) = 66
        transp(75) = 63
        transp(76) = 61
        transp(77) = 58
        transp(78) = 56
        transp(79) = 53
        transp(80) = 51
        transp(81) = 48
        transp(82) = 45
        transp(83) = 43
        transp(84) = 40
        transp(85) = 38
        transp(86) = 35
        transp(87) = 33
        transp(88) = 30
        transp(89) = 28
        transp(90) = 25
        transp(91) = 22
        transp(92) = 20
        transp(93) = 17
        transp(94) = 15
        transp(95) = 12
        transp(96) = 10
        transp(97) = 7
        transp(98) = 5
        transp(99) = 2
        transp(100) = 0
        Return transp
    End Function

    Public Function GetColor(colorStr As String) As Color
        If String.IsNullOrEmpty(colorStr) Then
            Return Nothing
            Exit Function
        End If
        Dim acdColor As Color

        If InStr(colorStr, ",") > 0 Then
            Dim rgb() As String = Split(colorStr, ",")
            acdColor = Color.FromRgb(CInt(rgb(0)), CInt(rgb(1)), CInt(rgb(2)))
        ElseIf InStr(colorStr, "PANTONE") > 0 Then
            Dim ptone() As String = Split(colorStr, "_")
            acdColor = Color.FromNames(ptone(0), ptone(1))
        ElseIf IsNumeric(colorStr) Then
            acdColor = Color.FromColorIndex(ColorMethod.ByAci, CInt(colorStr))
        Else
            Select Case UCase(colorStr)
                Case Is = "RED"
                    acdColor = Color.FromColorIndex(ColorMethod.ByAci, 1)
                Case Is = "YELLOW"
                    acdColor = Color.FromColorIndex(ColorMethod.ByAci, 2)
                Case Is = "GREEN"
                    acdColor = Color.FromColorIndex(ColorMethod.ByAci, 3)
                Case Is = "CYAN"
                    acdColor = Color.FromColorIndex(ColorMethod.ByAci, 4)
                Case Is = "BLUE"
                    acdColor = Color.FromColorIndex(ColorMethod.ByAci, 5)
                Case Is = "MAGENTA"
                    acdColor = Color.FromColorIndex(ColorMethod.ByAci, 6)
                Case Is = "WHITE"
                    acdColor = Color.FromColorIndex(ColorMethod.ByAci, 7)
                Case Else
                    acdColor = Color.FromColorIndex(ColorMethod.ByAci, 7)
            End Select
        End If
        Return acdColor

    End Function
    Public Function GetLTId(ltName As String) As ObjectId
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = curDwg.Editor
        Dim dwgDB As Database = curDwg.Database
        Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
            Dim ltTbl As LinetypeTable = actrans.GetObject(dwgDB.LinetypeTableId, OpenMode.ForRead)
            Dim obID As ObjectId
            If ltTbl.Has(ltName) Then
                obID = ltTbl(ltName)
                Return obID
            Else
                Return ltTbl("Continuous")
            End If
        End Using
    End Function

    Public Function GetMatId(ltName As String) As ObjectId
        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = curDwg.Editor
        Dim dwgDB As Database = curDwg.Database
        Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
            Dim ltTbl As LinetypeTable = actrans.GetObject(dwgDB.LinetypeTableId, OpenMode.ForRead)
            Dim obID As ObjectId
            If ltTbl.Has(ltName) Then
                obID = ltTbl(ltName)
                Return obID
            Else
                Return Nothing
            End If
        End Using
    End Function


End Module

Public Module MathGeometry

    'Private BlkFolder As String
    'Private FntName As String
    'Private StylNm As String

    'Const AutoCADver As String = "AutoCAD.Application.21"
    'Const Civil3dVer As String = "AeccXUiLand.Aeccapplication.11.0"

    Public Function GetTangentPoints(ptP As Point3d, c1 As Circle, Optional verbose As Boolean = False) As Point2dCollection

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = curDwg.Editor

        'declare a collection for output
        Dim points As New Point2dCollection

        'get the center and radius of the circle
        Dim cCen As Point3d = c1.Center
        Dim r As Double = c1.Radius

        'pull the 2D coordinates from the center point and the external point
        Dim px As Double = ptP.X
        Dim py As Double = ptP.Y
        Dim cx As Double = cCen.X
        Dim cy As Double = cCen.Y

        Dim ptP2d As New Point2d(px, py)
        Dim cCen2d As New Point2d(cx, cy)
        Dim v2d As Vector2d = ptP2d.GetVectorTo(cCen2d)
        'get angle from x-axis to a line between the external point and circle center
        Dim gamma As Double = v2d.Angle

        'Delta x and Delta y
        Dim dx = cx - px
        Dim dy = cy - py

        'if the same point, no tangents
        If dx = 0 And dy = 0 Then
            If verbose Then ed.WriteMessage(vbLf & "Error. Point cannot be the circle center.")
            Return Nothing
            Exit Function
        End If

        'get distance from ptP to center of the circle using acad function
        Dim pc = ptP2d.GetDistanceTo(cCen2d)
        'Dim pc = Sqrt(dx ^ 2 + dy ^ 2)

        'establish a temporary coordinate system with (0,0) at the external point
        'and the circle center located on the x-axis at x = pc
        'get the angle from the temp x-axis to a tangent point
        Dim theta As Double = Asin(r / pc)

        'test for no tangents or one tangent
        If pc < r Then 'no tangents
            If verbose Then ed.WriteMessage(vbLf & "Error. Point cannot be inside of the circle. No tangent points.")
            Return Nothing
            Exit Function
        ElseIf pc = r Then 'one tangent = ptP
            points.Add(New Point2d(ptP.X, ptP.Y))
            Return points
            Exit Function
        End If

        'get the coordinates of the tangent points relative to the temporary coordinate system
        Dim tp1x As Double = pc - (r * Sin(theta))
        Dim tp2x As Double = tp1x
        Dim tp1y As Double = r * Cos(theta)
        Dim tp2y As Double = -tp1y

        'rotate coordinates of tp1 and tp2 to final orientation
        Dim point1x As Double = (tp1x * Cos(gamma)) - (tp1y * Sin(gamma))
        Dim point1y As Double = (tp1x * Sin(gamma)) + (tp1y * Cos(gamma))
        Dim point2x As Double = (tp2x * Cos(gamma)) - (tp2y * Sin(gamma))
        Dim point2y As Double = (tp2x * Sin(gamma)) + (tp2y * Cos(gamma))

        'translate coordinates from temporary to world coordinate system
        Dim f1x As Double = point1x + ptP.X
        Dim f1y As Double = point1y + ptP.Y
        Dim f2x As Double = point2x + ptP.X
        Dim f2y As Double = point2y + ptP.Y

        'create 2d points and add to return collection
        Dim pt1 As New Point2d(f1x, f1y)
        Dim pt2 As New Point2d(f2x, f2y)

        points.Add(pt1)
        points.Add(pt2)

        Return points

    End Function

    Public Function GetTangentPoints(ptP As Point3d, dbObj As DBObject) As Point2dCollection

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = curDwg.Editor

        Dim c1 As Circle

        'substitute circles for arcs if arcs are picked
        If TypeOf dbObj Is Autodesk.AutoCAD.DatabaseServices.Circle Then
            c1 = CType(dbObj, Circle)
        ElseIf TypeOf dbObj Is Arc Then
            Dim a1 As Arc = CType(dbObj, Arc)
            c1 = New Circle(a1.Center, Vector3d.ZAxis, a1.Radius)
            a1.Dispose()
        Else
            ed.WriteMessage(vbLf & "Entity is not a circle or circular arc.  Command ended.")
            Return Nothing
            Exit Function
        End If

        'declare a collection for output
        Dim points As New Point2dCollection

        'get the center and radius of the circle
        Dim cCen As Point3d = c1.Center
        Dim r As Double = c1.Radius

        'pull the 2D coordinates from the center point and the external point
        Dim px As Double = ptP.X
        Dim py As Double = ptP.Y
        Dim cx As Double = cCen.X
        Dim cy As Double = cCen.Y

        Dim ptP2d As New Point2d(px, py)
        Dim cCen2d As New Point2d(cx, cy)
        Dim v2d As Vector2d = ptP2d.GetVectorTo(cCen2d)
        'get angle from x-axis to a line between the external point and circle center
        Dim gamma As Double = v2d.Angle

        'Delta x and Delta y
        Dim dx = cx - px
        Dim dy = cy - py

        'if the same point, no tangents
        If dx = 0 And dy = 0 Then
            ed.WriteMessage(vbLf & "Error. Point cannot be the circle center.")
            Return Nothing
            Exit Function
        End If

        'get distance from ptP to center of the circle using acad function
        Dim pc = ptP2d.GetDistanceTo(cCen2d)

        'establish a temporary coordinate system with (0,0) at the external point
        'and the circle center located on the x-axis at x = pc
        'get the angle from the temp x-axis to a tangent point
        Dim theta As Double = Asin(r / pc)

        'test for no tangents or one tangent
        If pc < r Then 'no tangents
            ed.WriteMessage(vbLf & "Error. Point cannot be inside of the circle. No tangent points.")
            Return Nothing
            Exit Function
        ElseIf pc = r Then 'one tangent = ptP
            points.Add(New Point2d(ptP.X, ptP.Y))
            Return points
            Exit Function
        End If

        'get the coordinates of the tangent points relative to the temporary x-axis
        Dim tp1x As Double = pc - (r * Sin(theta))
        Dim tp2x As Double = tp1x
        Dim tp1y As Double = r * Cos(theta)
        Dim tp2y As Double = -tp1y

        'rotate coordinates of tp1 and tp2 to final orientation
        Dim point1x As Double = (tp1x * Cos(gamma)) - (tp1y * Sin(gamma))
        Dim point1y As Double = (tp1y * Cos(gamma)) + (tp1x * Sin(gamma))
        Dim point2x As Double = (tp2x * Cos(gamma)) - (tp2y * Sin(gamma))
        Dim point2y As Double = (tp2y * Cos(gamma)) + (tp2x * Sin(gamma))

        'translate coordinates from temporary to world coordinate system
        Dim f1x As Double = point1x + ptP.X
        Dim f1y As Double = point1y + ptP.Y
        Dim f2x As Double = point2x + ptP.X
        Dim f2y As Double = point2y + ptP.Y

        'create points and add to return collection
        Dim pt1 As New Point2d(f1x, f1y)
        Dim pt2 As New Point2d(f2x, f2y)
        points.Add(pt1)
        points.Add(pt2)

        Return points

    End Function

    Public Function GetAngleFromXaxis(StartPt As Point3d, RefPt As Point3d) As Double
        'Function returns the angle defined by the x-axis and a vector defined by two 3D points
        'in the xy plane
        Try
            Dim SP2d = New Point2d(StartPt.X, StartPt.Y)
            Dim RP2D = New Point2d(RefPt.X, RefPt.Y)
            Dim vect2d As Vector2d = SP2d.GetVectorTo(RP2D)
            'Dim v3d As Vector3d = StartPt.GetVectorTo(RefPt)
            Dim InsAng As Double = vect2d.Angle
            Return InsAng
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function ExtTan2Circles(dbObj1 As DBObject, DBObj2 As DBObject, Optional verbose As Boolean = False) As Point2dCollection

        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        Dim c1 As Circle
        Dim c2 As Circle

        Try
            'substitute circles for arcs if arcs are picked
            If TypeOf dbObj1 Is Autodesk.AutoCAD.DatabaseServices.Circle Then
                c1 = CType(dbObj1, Circle)
            ElseIf TypeOf dbObj1 Is Arc Then
                Dim a1 As Arc = CType(dbObj1, Arc)
                c1 = New Circle(a1.Center, Vector3d.ZAxis, a1.Radius)
                a1.Dispose()
            Else
                If verbose Then ed.WriteMessage(vbLf & "C1 is not a circle or circular arc.  Command ended.")
                If dbObj1 IsNot Nothing Then dbObj1.Dispose()
                Return Nothing
                Exit Function
            End If

            If TypeOf DBObj2 Is Autodesk.AutoCAD.DatabaseServices.Circle Then
                c2 = CType(DBObj2, Circle)
            ElseIf TypeOf DBObj2 Is Arc Then
                Dim a2 As Arc = CType(DBObj2, Arc)
                c2 = New Circle(a2.Center, Vector3d.ZAxis, a2.Radius)
                a2.Dispose()
            Else
                If verbose Then ed.WriteMessage(vbLf & "C2 is not a circle or circular arc.  Command ended.")
                If DBObj2 IsNot Nothing Then DBObj2.Dispose()
                Return Nothing
                Exit Function
            End If

            'make the smaller circle c1
            If c1.Radius > c2.Radius Then
                Dim ctemp As Circle = c1
                c1 = c2
                c2 = ctemp
            End If

            'create 2D points for the picked circles' centers
            Dim c12D As New Point2d(c1.Center.X, c1.Center.Y)
            Dim c22D As New Point2d(c2.Center.X, c2.Center.Y)

            'get the angle between the x-axis and line between the circle centers
            'create a 2D vector from center c1 to center c2
            Dim v2d As Vector2d = c12D.GetVectorTo(c22D)
            Dim rotangle = v2d.Angle

            'round the distance between circle centers to 10 decimals to prevent floating point errors
            Dim cenDist As Double = Round(v2d.Length, 10)

            Dim r1 As Double = c1.Radius
            Dim r2 As Double = c2.Radius

            'create temp circles on the x-axis at the same distance apart
            Dim tc1 As New Circle(New Point3d(0, 0, 0), Vector3d.ZAxis, r1)
            Dim tc2 As New Circle(New Point3d(cenDist, 0, 0), Vector3d.ZAxis, r2)

            Dim tangents As Integer
            Dim singTanPts As New Point2dCollection
            Dim raddiff As Double = Round(r2 - r1, 10)
            Dim radsum As Double = Round(r2 + r1, 10)

            'test for the number of tangents
            If cenDist < raddiff Then
                tangents = 0
                If verbose Then ed.WriteMessage(vbLf & "C1 is completely inside of C2. Circles have no common tangents.")
                Return Nothing
                Exit Function
            ElseIf cenDist = raddiff Then
                tangents = 1
                If verbose Then ed.WriteMessage(vbLf & "C1 is inside of and touching C2. Circles have one common exterior tangent at the point of their intersection.")
                'get tangent for touching circles
                singTanPts = GetSingleTangent(c1, c2)
            ElseIf cenDist > raddiff AndAlso cenDist < radsum Then
                tangents = 2
                If verbose Then ed.WriteMessage(vbLf & "C1 intersects with C2. Circles have two common exterior tangents.")
            ElseIf cenDist = radsum Then
                tangents = 3
                If verbose Then ed.WriteMessage(vbLf & "C1 is outside of and touching C2. Circles have 3 common exterior tangents.")
                'get tangent for touching circles
                singTanPts = GetSingleTangent(c1, c2)
            ElseIf cenDist > radsum Then
                tangents = 4
            End If

            Dim p1 As Point2d
            Dim p2 As Point2d
            Dim p3 As Point2d
            Dim p4 As Point2d

            'temp circle center coordinates
            Dim tc1x As Double = tc1.Center.X
            Dim tc1y As Double = tc1.Center.Y
            Dim tc2x As Double = tc2.Center.X
            Dim tc2y As Double = tc2.Center.Y

            'calc the angle between the x-axis and the tangent lines (temp circles)
            Dim alpha As Double = Asin((r2 - r1) / cenDist)
            If alpha < 0 Then alpha = Asin((r1 - r2) / cenDist)

            Dim cen1 As New Point2d(tc1x, tc1y)
            Dim cen2 As New Point2d(tc2x, tc2y)

            'create an output collection
            Dim tPts As New Point2dCollection

            'get displacement vector from temp circles to final positions
            Dim dispVect1 As Vector3d = tc1.Center.GetVectorTo(c1.Center)

            'calculate the coordinates for the tangent points
            'if circles are the same size then the points are simply offsets of the centerline
            If Not tangents = 1 Then

                Dim p1xt As Double
                Dim p1yt As Double
                Dim p2xt As Double
                Dim p2yt As Double
                Dim p3xt As Double
                Dim p3yt As Double
                Dim p4xt As Double
                Dim p4yt As Double

                If Not r1 = r2 Then
                    p1xt = tc2x - r2 * Sin(alpha)
                    p1yt = tc2y + r2 * Cos(alpha)
                    p2xt = tc1x - r1 * Sin(alpha)
                    p2yt = tc1y + r1 * Cos(alpha)
                    p3xt = p1xt
                    p3yt = -p1yt
                    p4xt = p2xt
                    p4yt = -p2yt
                Else
                    p1xt = tc1x
                    p1yt = r1
                    p2xt = tc2x
                    p2yt = r2
                    p3xt = tc1x
                    p3yt = -r1
                    p4xt = tc2x
                    p4yt = -r1
                End If

                'rotate coordinates

                Dim p1x As Double = p1xt * Cos(rotangle) - p1yt * Sin(rotangle)
                Dim p1y As Double = p1xt * Sin(rotangle) + p1yt * Cos(rotangle)
                Dim p2x As Double = p2xt * Cos(rotangle) - p2yt * Sin(rotangle)
                Dim p2y As Double = p2xt * Sin(rotangle) + p2yt * Cos(rotangle)
                Dim p3x As Double = p3xt * Cos(rotangle) - p3yt * Sin(rotangle)
                Dim p3y As Double = p3xt * Sin(rotangle) + p3yt * Cos(rotangle)
                Dim p4x As Double = p4xt * Cos(rotangle) - p4yt * Sin(rotangle)
                Dim p4y As Double = p4xt * Sin(rotangle) + p4yt * Cos(rotangle)

                'translate coordinates
                p1x += dispVect1.X
                p1y += dispVect1.Y
                p2x += dispVect1.X
                p2y += dispVect1.Y
                p3x += dispVect1.X
                p3y += dispVect1.Y
                p4x += dispVect1.X
                p4y += dispVect1.Y

                'final 2d points
                p1 = New Point2d(p1x, p1y)
                p2 = New Point2d(p2x, p2y)
                p3 = New Point2d(p3x, p3y)
                p4 = New Point2d(p4x, p4y)

                tPts.Add(p1)
                tPts.Add(p2)
                tPts.Add(p3)
                tPts.Add(p4)

            End If

            'add points if the circles are touching
            If singTanPts IsNot Nothing AndAlso singTanPts.Count = 3 Then
                tPts.Add(singTanPts(0))
                tPts.Add(singTanPts(1))
                tPts.Add(singTanPts(2))
            End If

            'dispose temp circles
            tc1.Dispose()
            tc2.Dispose()

            Return tPts

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return Nothing
        End Try


    End Function

    Public Function GetSingleTangent(c1 As Circle, c2 As Circle) As Point2dCollection

        'routine for getting the tangent of two touching circles at their single intersection point

        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        Dim cen1 As New Point2d(c1.Center.X, c1.Center.Y)
        Dim cen2 As New Point2d(c2.Center.X, c2.Center.Y)

        Dim v2d As Vector2d = cen1.GetVectorTo(cen2)
        Dim cDist As Double = v2d.Length
        Dim rotAngle As Double = v2d.Angle

        Dim r1 As Double = c1.Radius
        Dim r2 As Double = c2.Radius

        Dim tc1 As New Circle(New Point3d(0, 0, 0), Vector3d.ZAxis, r1)
        Dim tc2 As New Circle(New Point3d(cDist, 0, 0), Vector3d.ZAxis, r2)

        'get displacement vector from temp circles to final positions
        Dim dispVect1 As Vector3d = tc1.Center.GetVectorTo(c1.Center)

        Dim pt1x As Double
        Dim pt1y As Double = -50
        Dim pt2x As Double
        Dim pt2y As Double = 50
        Dim pt3x As Double
        Dim pt3y As Double = 0

        If cDist < r2 Then
            pt1x = -r1
            pt2x = -r1
            pt3x = -r1
        ElseIf cDist > r2 Then
            pt1x = r1
            pt2x = r1
            pt3x = r1
        Else
            ed.WriteMessage(vbLf & "Error.  Circles do not touch at a single point.")
            Return Nothing
            Exit Function
        End If

        Dim p1x As Double = pt1x * Cos(rotAngle) - pt1y * Sin(rotAngle)
        Dim p1y As Double = pt1x * Sin(rotAngle) + pt1y * Cos(rotAngle)
        Dim p2x As Double = pt2x * Cos(rotAngle) - pt2y * Sin(rotAngle)
        Dim p2y As Double = pt2x * Sin(rotAngle) + pt2y * Cos(rotAngle)
        Dim p3x As Double = pt3x * Cos(rotAngle) - pt3y * Sin(rotAngle)
        Dim p3y As Double = pt3x * Sin(rotAngle) + pt3y * Cos(rotAngle)

        'translate coordinates
        p1x += dispVect1.X
        p1y += dispVect1.Y
        p2x += dispVect1.X
        p2y += dispVect1.Y
        p3x += dispVect1.X
        p3y += dispVect1.Y

        'final 2d points
        Dim p1 As New Point2d(p1x, p1y)
        Dim p2 As New Point2d(p2x, p2y)
        Dim p3 As New Point2d(p3x, p3y)

        Dim tPts As New Point2dCollection

        With tPts
            .Add(p1)
            .Add(p2)
            .Add(p3)
        End With


        tc1.Dispose()
        tc2.Dispose()

        Return tPts
    End Function

    Public Function IsArc(objId As ObjectId) As Boolean
        'test any object to see if it is an arc
        Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = acDwg.Database

        Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
            Dim blineObj As DBObject = TryCast(actrans.GetObject(objId, OpenMode.ForRead), DBObject)
            If blineObj IsNot Nothing Then
                If TypeOf blineObj Is Arc Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
            actrans.Commit()
        End Using

    End Function

    Public Function IntTan2Circles(dbObj1 As DBObject, DBObj2 As DBObject, Optional verbose As Boolean = False) As Point2dCollection

        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        Dim c1 As Circle
        Dim c2 As Circle

        Try
            'substitute circles for arcs if arcs are picked
            If TypeOf dbObj1 Is Autodesk.AutoCAD.DatabaseServices.Circle Then
                c1 = CType(dbObj1, Circle)
            ElseIf TypeOf dbObj1 Is Arc Then
                Dim a1 As Arc = CType(dbObj1, Arc)
                c1 = New Circle(a1.Center, Vector3d.ZAxis, a1.Radius)
                a1.Dispose()
            Else
                If verbose Then ed.WriteMessage(vbLf & "C1 is not a circle or circular arc.  Command ended.")
                Return Nothing
                Exit Function
            End If

            If TypeOf DBObj2 Is Autodesk.AutoCAD.DatabaseServices.Circle Then
                c2 = CType(DBObj2, Circle)
            ElseIf TypeOf DBObj2 Is Arc Then
                Dim a2 As Arc = CType(DBObj2, Arc)
                c2 = New Circle(a2.Center, Vector3d.ZAxis, a2.Radius)
                a2.Dispose()
            Else
                If verbose Then ed.WriteMessage(vbLf & "C2 is not a circle or circular arc.  Command ended.")
                Return Nothing
                Exit Function
            End If

            'make c1 the smaller circle
            If c1.Radius > c2.Radius Then
                Dim ctemp As Circle = c1
                c1 = c2
                c2 = ctemp
                'Return Nothing
                'Exit Function
            End If

            '2D circle centers
            Dim c12D As New Point2d(c1.Center.X, c1.Center.Y)
            Dim c22D As New Point2d(c2.Center.X, c2.Center.Y)

            'get the angle from the x-axis to the line between circle centers
            Dim v2d As Vector2d = c12D.GetVectorTo(c22D)
            Dim rotangle = v2d.Angle
            Dim cenDist As Double = Round(c12D.GetDistanceTo(c22D), 10)

            'test for the number of tangents
            If cenDist = 0 Then
                If verbose Then ed.WriteMessage(vbLf & "Circles are cocentric.  No common tangents.")
                Return Nothing
                Exit Function
            End If

            Dim tangents As Integer
            Dim raddiff As Double = Round(c2.Radius - c1.Radius, 10)
            Dim radsum As Double = Round(c2.Radius + c1.Radius, 10)

            If cenDist < raddiff Then
                tangents = 0
                If verbose Then ed.WriteMessage(vbLf & "C1 is completely inside of C2. Circles have no common tangents.")
                Return Nothing
                Exit Function
            ElseIf cenDist = raddiff Then
                tangents = 1
                If verbose Then ed.WriteMessage(vbLf & "C1 is inside of and touching C2. Circles have no common interior tangents.")
                Return Nothing
                Exit Function
            ElseIf cenDist > raddiff And cenDist < radsum Then
                tangents = 2
                If verbose Then ed.WriteMessage(vbLf & "C1 intersects with C2. Circles have no common interior tangents.")
                Return Nothing
                Exit Function
            ElseIf cenDist = radsum Then
                tangents = 3
                If verbose Then ed.WriteMessage(vbLf & "C1 is outside of and touching C2. Circles have no common interior tangents.")
                Return Nothing
                Exit Function
            ElseIf cenDist > radsum Then
                tangents = 4
                If verbose Then ed.WriteMessage(vbLf & "Circles have 2 common internal tangents.")
            End If

            Dim r1 As Double = c1.Radius
            Dim r2 As Double = c2.Radius

            'create temp circles on the x-axis
            Dim tc1 As New Circle(New Point3d(0, 0, 0), Vector3d.ZAxis, r1)
            Dim tc2 As New Circle(New Point3d(cenDist, 0, 0), Vector3d.ZAxis, r2)

            Dim tc1x As Double = 0
            Dim tc1y As Double = 0
            Dim tc2x As Double = cenDist
            Dim tc2y As Double = 0

            Dim p1 As Point2d
            Dim p2 As Point2d
            Dim p3 As Point2d
            Dim p4 As Point2d

            'calc tangent point coordinates
            If Not r1 = r2 Then
                Dim d1 As Double = (r1 * cenDist) / (r1 + r2)
                Dim theta As Double = Asin(r1 / d1)
                Dim tpc1x As Double = r1 * Sin(theta)
                Dim tpc1y As Double = r1 * Cos(theta)
                Dim d2 As Double = cenDist - d1

                Dim alpha As Double = Asin(r2 / d2)
                Dim tpc2x As Double = cenDist - (r2 * Sin(alpha))
                Dim tpc2y As Double = r2 * Cos(alpha)
                Dim cen1 As New Point2d(tc1x, tc1y)
                Dim cen2 As New Point2d(tc2x, tc2y)

                p1 = New Point2d(tpc1x, tpc1y)
                p2 = New Point2d(tpc2x, tpc2y)
                p3 = New Point2d(tpc1x, -tpc1y)
                p4 = New Point2d(tpc2x, -tpc2y)

            Else
                Dim theta As Double = Asin(r1 / (cenDist / 2))
                Dim tpc1x As Double = r1 * Sin(theta)
                Dim tpc1y As Double = r1 * Cos(theta)
                Dim tpc2x As Double = cenDist - (r1 * Sin(theta))
                Dim tpc2y As Double = tpc1y

                p1 = New Point2d(tpc1x, tpc1y)
                p2 = New Point2d(tpc2x, tpc2y)
                p3 = New Point2d(tpc1x, -tpc1y)
                p4 = New Point2d(tpc2x, -tpc1y)
            End If

            'rotate points
            Dim p1x As Double = p1.X * Cos(rotangle) - p1.Y * Sin(rotangle)
            Dim p1y As Double = p1.X * Sin(rotangle) + p1.Y * Cos(rotangle)
            Dim p2x As Double = p2.X * Cos(rotangle) - p2.Y * Sin(rotangle)
            Dim p2y As Double = p2.X * Sin(rotangle) + p2.Y * Cos(rotangle)
            Dim p3x As Double = p3.X * Cos(rotangle) - p3.Y * Sin(rotangle)
            Dim p3y As Double = p3.X * Sin(rotangle) + p3.Y * Cos(rotangle)
            Dim p4x As Double = p4.X * Cos(rotangle) - p4.Y * Sin(rotangle)
            Dim p4y As Double = p4.X * Sin(rotangle) + p4.Y * Cos(rotangle)

            Dim dV As Vector3d = tc1.Center.GetVectorTo(c1.Center)

            'move points
            p1x += dV.X
            p1y += dV.Y
            p2x += dV.X
            p2y += dV.Y
            p3x += dV.X
            p3y += dV.Y
            p4x += dV.X
            p4y += dV.Y

            Dim p1f As New Point2d(p1x, p1y)
            Dim p2f As New Point2d(p2x, p2y)
            Dim p3f As New Point2d(p3x, p3y)
            Dim p4f As New Point2d(p4x, p4y)

            'create an output collection
            Dim tPts As New Point2dCollection

            With tPts
                .Add(p1f)
                .Add(p2f)
                .Add(p3f)
                .Add(p4f)
            End With

            Return tPts

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return Nothing
        End Try

    End Function

    Public Function Arc2poly(arcID As ObjectId) As ObjectId

        'converts an arc to an equivalent polyline

        Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        'Dim ed As Editor = acDwg.Editor
        Dim dwgDB As Database = acDwg.Database
        Dim poly As New Polyline
        Try
            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()

                Dim cArc As Arc = TryCast(acTrans.GetObject(arcID, OpenMode.ForWrite), Arc)

                If cArc IsNot Nothing Then
                    Dim bt As BlockTable = CType(dwgDB.BlockTableId.GetObject(OpenMode.ForRead), BlockTable)
                    Dim btr As BlockTableRecord = CType(acTrans.GetObject(dwgDB.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)

                    With poly
                        .AddVertexAt(0, New Point2d(cArc.StartPoint.X, cArc.StartPoint.Y), GetArcBulge(cArc), 0, 0)
                        .AddVertexAt(1, New Point2d(cArc.EndPoint.X, cArc.EndPoint.Y), 0, 0, 0)
                        .LayerId = cArc.LayerId
                    End With

                    btr.AppendEntity(poly)
                    acTrans.AddNewlyCreatedDBObject(poly, True)
                    cArc.[Erase]()
                Else
                    Return Nothing
                    acTrans.Abort()
                    Exit Function
                End If

                acTrans.Commit()
            End Using

        Catch ex As Exception
            MessageBox.Show("Error in Arc2Poly function." & vbLf & ex.Message)
            Return Nothing
            Exit Function
        End Try

        Return poly.ObjectId
    End Function
    Public Function GetArcBulge(ByVal arc As Arc) As Double
        Dim deltaAng As Double = arc.EndAngle - arc.StartAngle
        'If deltaAng < 0 Then deltaAng += 2 * Math.PI
        Return Tan(deltaAng * 0.25)
    End Function

    Public Function GetTangentBulge(v1 As Vector2d, v2 As Vector2d) As Double

        Dim blg As Double
        If v1 = v2 Then
            Return 0
        Else
            Dim ang As Double = v1.GetAngleTo(v2)
            blg = Tan(ang / 4)
        End If

        Return blg

    End Function

    Public Function GetTangentBulge(l1 As Line, l2 As Line) As Double

        'get tangent bulge between two lines

        Dim blg As Double

        If l1 = l2 Then
            Return 0
        Else
            Dim p1 As New Point2d(l1.StartPoint.X, l1.StartPoint.Y)
            Dim p2 As New Point2d(l1.EndPoint.X, l1.EndPoint.Y)
            Dim p3 As New Point2d(l2.StartPoint.X, l2.StartPoint.Y)
            Dim p4 As New Point2d(l2.EndPoint.X, l2.EndPoint.Y)

            Dim v1 As Vector2d = p1.GetVectorTo(p2)
            Dim v2 As Vector2d = p3.GetVectorTo(p4)

            Dim ang As Double = v1.GetAngleTo(v2)
            blg = Tan(ang / 4)
        End If

        Return blg

    End Function

    Public Function GetTangentBulge(vert1 As Vertex2d, vert2 As Vertex2d, vert3 As Vertex2d, vert4 As Vertex2d) As Double

        'get tangent bulge between polyline segments

        Dim blg As Double

        Dim p1 As New Point2d(vert1.Position.X, vert1.Position.Y)
        Dim p2 As New Point2d(vert2.Position.X, vert2.Position.Y)
        Dim p3 As New Point2d(vert3.Position.X, vert3.Position.Y)
        Dim p4 As New Point2d(vert4.Position.X, vert4.Position.Y)

        Dim v1 As Vector2d = p1.GetVectorTo(p2)
        Dim v2 As Vector2d = p3.GetVectorTo(p4)

        Dim ang As Double = v1.GetAngleTo(v2)
        blg = Tan(ang / 4)

        Return blg

    End Function

    Public Function GetTangentBulge(v1 As Point3d, v2 As Point3d, v3 As Point3d, v4 As Point3d) As Double
        'get tangent bulge using vertex points

        Dim blg As Double

        Dim p1 As New Point2d(v1.X, v1.Y)
        Dim p2 As New Point2d(v2.X, v2.y)
        Dim p3 As New Point2d(v3.X, v3.Y)
        Dim p4 As New Point2d(v4.X, v4.Y)

        Dim vect1 As Vector2d = p1.GetVectorTo(p2)
        Dim vect2 As Vector2d = p3.GetVectorTo(p4)

        Dim ang As Double = vect1.GetAngleTo(vect2)
        blg = Tan(ang / 4)

        Return blg

    End Function


    Public Function OffsetSideByPoint(clineID As ObjectId, offdist As Double, Optional PrptMessage As String = "") As Integer

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = curDwg.Editor
        Dim dwgDB As Database = curDwg.Database

        'declare return variable
        Dim offside As Integer

        'lock the drawing file
        'Using dl As DocumentLock = curDwg.LockDocument()

        'start transaction
        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()

            Try
                'get reference point
                Dim ppOPts As New PromptPointOptions("")
                With ppOPts
                    If String.IsNullOrEmpty(PrptMessage) Then
                        .Message = (vbLf & "Select any point on the desired side of the reference line.")
                    Else
                        .Message = PrptMessage
                    End If
                    .AllowNone = False
                End With

                Dim ppR As PromptPointResult = ed.GetPoint(ppOPts)
                Dim pRef As New Point3d

                'if user cancels, return 10
                If ppR.Status = PromptStatus.Cancel Then
                    Return 10
                    Exit Function
                Else
                    pRef = ppR.Value
                End If

                'try to cast picked object as a curve.  If it doesn't work, then it can't be offset
                Dim pl As Curve = TryCast(acTrans.GetObject(clineID, OpenMode.ForRead), Curve)
                If pl IsNot Nothing Then
                    'get closest point on curve to picked point
                    Dim cp As Point3d = pl.GetClosestPointTo(ppR.Value, False)

                    'create offset collections on each side of the reference a short distance from the original line
                    Dim offsetPos As DBObjectCollection = pl.GetOffsetCurves(offdist)
                    Dim offsetNeg As DBObjectCollection = pl.GetOffsetCurves(-offdist)
                    Dim curvepos As Curve
                    Dim curveneg As Curve
                    Dim pPos As Point3d
                    Dim pNeg As Point3d
                    Dim distPos As Double
                    Dim distNeg As Double

                    'cast both offset collections as curves
                    For i As Integer = 0 To offsetPos.Count - 1
                        curvepos = CType(offsetPos(i), Curve)
                        curveneg = CType(offsetNeg(i), Curve)

                        'see which curve is closer to picked point (pRef)
                        pPos = curvepos.GetClosestPointTo(pRef, False)
                        pNeg = curveneg.GetClosestPointTo(pRef, False)
                        distPos = pPos.DistanceTo(pRef)
                        distNeg = pNeg.DistanceTo(pRef)

                        'return 1 for the positive side and -1 for the negative side
                        If distPos < distNeg Then
                            offside = 1
                        ElseIf distNeg < distPos Then
                            offside = -1
                        Else
                            offside = 0
                        End If
                    Next
                Else
                    'if something goes wrong, return 10
                    ed.WriteMessage(vbLf & "Selected entity cannot be used as reference line.  End Command.")
                    Return 10
                    Exit Function
                End If
                acTrans.Commit()

            Catch ex As Exception
                MessageBox.Show("Fatal Error." & ex.Message)
                Return 10
            End Try

        End Using
        Return offside

    End Function

    Public Function Roundmult(numb As Double, rMult As Long) As Long

        Dim remRad As Double = numb Mod rMult
        Dim radlong As Long

        If remRad < 0.5 * rMult Then
            radlong = CLng(numb - remRad)
        Else
            radlong = CLng(numb + (rMult - remRad))
        End If
        Return radlong

    End Function

    Public Function CRads(degs As Double) As Double
        Return (Math.PI / 180) * degs
    End Function

    Public Function CDegs(rads As Double) As Double
        Return (180 / Math.PI) * rads
    End Function

    Public Function UnitVector3d(vect As Vector3d) As Vector3d

        Dim unitVect As Vector3d
        Try
            'UnitVector components
            Dim xnrml As Double = vect.X / (vect.X ^ 2 + vect.Y ^ 2 + vect.Z ^ 2) ^ (0.5)
            Dim ynrml As Double = vect.Y / (vect.X ^ 2 + vect.Y ^ 2 + vect.Z ^ 2) ^ (0.5)
            Dim znrml As Double = vect.Z / (vect.X ^ 2 + vect.Y ^ 2 + vect.Z ^ 2) ^ (0.5)
            'unit vector
            unitVect = New Vector3d(xnrml, ynrml, znrml)
        Catch ex As Exception
            Return Nothing
            Exit Function
        End Try
        Return unitVect

    End Function

    Public Function MakeNewUCS(orgPt As Point3d, xAxPt As Point3d, ucsName As String) As Boolean

        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        Dim cUCS As Matrix3d = ed.CurrentUserCoordinateSystem

        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction()
            Try
                Dim acUCSTbl As UcsTable = DirectCast(acTrans.GetObject(DwgDB.UcsTableId, OpenMode.ForWrite), UcsTable)
                Dim tempUcsRec As UcsTableRecord

                If acUCSTbl.Has(ucsName) = False Then
                    tempUcsRec = New UcsTableRecord With {.Name = ucsName}
                    acUCSTbl.Add(tempUcsRec)
                    acTrans.AddNewlyCreatedDBObject(tempUcsRec, True)
                Else
                    Dim kwPick As New PromptKeywordOptions("")
                    With kwPick
                        .AllowNone = False
                        .Message = vbLf & "UCS exists. Overwrite?"
                        .Keywords.Add("Y")
                        .Keywords.Add("N")
                        .AppendKeywordsToMessage = True
                    End With

                    Dim kwRes As PromptResult = ed.GetKeywords(kwPick)
                    Dim uOpt As String
                    If kwRes.Status = PromptStatus.OK Then
                        uOpt = kwRes.StringResult
                    Else
                        Return False
                        ed.CurrentUserCoordinateSystem = cUCS
                        CurDwg.Dispose()
                        DwgDB.Dispose()
                        Exit Function
                    End If

                    If uOpt = "Y" Then
                        tempUcsRec = DirectCast(acTrans.GetObject(acUCSTbl(ucsName), OpenMode.ForWrite), UcsTableRecord)
                    Else
                        ucsName = GetNewUCSName(ucsName)
                        tempUcsRec = New UcsTableRecord With {.Name = ucsName}
                        acUCSTbl.Add(tempUcsRec)
                        acTrans.AddNewlyCreatedDBObject(tempUcsRec, True)
                    End If

                    tempUcsRec = DirectCast(acTrans.GetObject(acUCSTbl(ucsName), OpenMode.ForWrite), UcsTableRecord)
                    Dim curUCSname = tempUcsRec.Name
                End If

                tempUcsRec.Origin = orgPt
                Dim tempvec As Vector3d = orgPt.GetVectorTo(xAxPt)
                Dim curZaxis As New Vector3d(0, 0, 1)
                Dim curXaxis As New Vector3d(1, 0, 0)
                Dim curYaxis As New Vector3d(0, 1, 0)

                'Dim temporthoZ As Vector3d = (tempvec.DotProduct(curZaxis) / tempvec.DotProduct(tempvec)) * tempvec
                'Dim temporthoX As Vector3d = (tempvec.DotProduct(curXaxis) / tempvec.DotProduct(tempvec)) * tempvec
                'Dim temporthoY As Vector3d = (tempvec.DotProduct(curYaxis) / tempvec.DotProduct(tempvec)) * tempvec

                Dim temporthoY As Vector3d = (tempvec.GetPerpendicularVector)
                Dim tempxunit As Vector3d = UnitVector3d(tempvec)
                Dim temporthoX As Vector3d = (tempvec.DotProduct(curXaxis) / tempvec.DotProduct(tempvec)) * tempvec
                Dim temporthoZ As Vector3d = (tempvec.DotProduct(curZaxis) / tempvec.DotProduct(tempvec)) * tempvec


                'Dim tempXaxis As New Vector3d(tempvec.X, tempvec.Y,0)
                Dim tempXaxis As Vector3d = temporthoX
                'Dim tempZaxis As New Vector3d(0, 0, 1) 'tempXaxis.TransformBy(Matrix3d.Rotation(PI / 2, tempXaxis, orgPt))
                Dim tempZaxis As Vector3d = temporthoZ

                'Dim tempYaxis As Vector3d = tempXaxis.TransformBy(Matrix3d.Rotation(PI / 2, tempZaxis,
                'orgPt))

                'Dim tempYaxis As Vector3d = tempXaxis.TransformBy(Matrix3d.Rotation(PI / 2, tempZaxis, orgPt))
                Dim tempYaxis As Vector3d = temporthoY

                'calculate components of axis unit vectors for new UCS matrix
                'x axis vector
                Dim xnrmlx As Double = tempXaxis.X / tempXaxis.Length ''(tempXaxis.X ^ 2 + tempXaxis.Y ^ 2 + tempXaxis.Z ^ 2) ^ (0.5)
                Dim xnrmly As Double = tempXaxis.Y / tempXaxis.Length '(tempXaxis.X ^ 2 + tempXaxis.Y ^ 2 + tempXaxis.Z ^ 2) ^ (0.5)
                Dim xnrmlz As Double = tempXaxis.Z / tempXaxis.Length '(tempXaxis.X ^ 2 + tempXaxis.Y ^ 2 + tempXaxis.Z ^ 2) ^ (0.5)
                'y axis vector
                Dim ynrmlx As Double = tempYaxis.X / tempYaxis.Length '(tempYaxis.X ^ 2 + tempYaxis.Y ^ 2 + tempYaxis.Z ^ 2) ^ (0.5)
                Dim ynrmly As Double = tempYaxis.Y / tempYaxis.Length '(tempYaxis.X ^ 2 + tempYaxis.Y ^ 2 + tempYaxis.Z ^ 2) ^ (0.5)
                Dim ynrmlz As Double = tempYaxis.Z / tempYaxis.Length '(tempYaxis.X ^ 2 + tempYaxis.Y ^ 2 + tempYaxis.Z ^ 2) ^ (0.5)

                Dim znrmlx As Double = tempZaxis.X / tempZaxis.Length '(tempZaxis.X ^ 2 + tempZaxis.Y ^ 2 + tempZaxis.Z ^ 2) ^ (0.5)
                Dim znrmly As Double = tempZaxis.Y / tempZaxis.Length '(tempZaxis.X ^ 2 + tempZaxis.Y ^ 2 + tempZaxis.Z ^ 2) ^ (0.5)
                Dim znrmlz As Double = tempZaxis.Z / tempZaxis.Length '(tempZaxis.X ^ 2 + tempZaxis.Y ^ 2 + tempZaxis.Z ^ 2) ^ (0.5)

                Dim xAxUnitVect As New Vector3d(xnrmlx, xnrmly, xnrmlz)
                Dim yAxUnitVect As New Vector3d(ynrmlx, ynrmly, ynrmlz)
                Dim zAxUnitVect As New Vector3d(znrmlx, znrmly, znrmlz)

                Dim tempValue As Double = tempvec.DotProduct(temporthoY)

                If tempXaxis.IsPerpendicularTo(tempZaxis) Then 'And xAxUnitVect.IsPerpendicularTo(zAxUnitVect) And yAxUnitVect.IsPerpendicularTo(zAxUnitVect) Then
                    MessageBox.Show("Orthogonal Vectors")
                Else
                    MessageBox.Show("not orthogonal. dot product = " & tempValue.ToString)
                    Return Nothing
                    Exit Function
                End If

                tempUcsRec.XAxis = xAxUnitVect
                tempUcsRec.YAxis = yAxUnitVect

                'open the active viewport
                Dim acVportRec As ViewportTableRecord
                acVportRec = DirectCast(acTrans.GetObject(ed.ActiveViewportId, OpenMode.ForWrite), ViewportTableRecord)

                acVportRec.IconAtOrigin = True
                acVportRec.IconEnabled = True

                'turn off ucsfollow mode
                If acVportRec.UcsFollowMode Then
                    acVportRec.UcsFollowMode = False
                End If

                'Set the temp UCS current
                acVportRec.SetUcs(tempUcsRec.ObjectId)
                ed.UpdateTiledViewportsFromDatabase()

                acTrans.Commit()

            Catch ex As Exception
                Return False
                ed.CurrentUserCoordinateSystem = cUCS
                CurDwg.Dispose()
                DwgDB.Dispose()
                Exit Function
            End Try

        End Using

        Return True

        CurDwg.Dispose()
        DwgDB.Dispose()

    End Function
    Public Function GetUCSMatrix2D(oPt As Point3d, xPt As Point3d) As Matrix2d

        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        'Dim tempUCS As Matrix3d
        'Dim NewUCSMatrix As Matrix2d
        Dim orgPt As New Point2d(oPt.X, oPt.Y)
        Dim xAxPt As New Point2d(xPt.X, xPt.Y)

        Try
            Using acTrans1 As Transaction = DwgDB.TransactionManager.StartTransaction()

                Dim tempXaxis As Vector2d = orgPt.GetVectorTo(xAxPt)
                'Dim tempYaxis As Vector2d = tempXaxis.GetPerpendicularVector
                Dim tempYaxis As Vector2d = tempXaxis.RotateBy(PI / 2)

                'calculate components of unit vectors for new UCS matrix
                'x axis vector
                Dim xnrmlx As Double = tempXaxis.X / (tempXaxis.X ^ 2 + tempXaxis.Y ^ 2) ^ (0.5)
                Dim xnrmly As Double = tempXaxis.Y / (tempXaxis.X ^ 2 + tempXaxis.Y ^ 2) ^ (0.5)
                'y axis vector
                Dim ynrmlx As Double = tempYaxis.X / (tempYaxis.X ^ 2 + tempYaxis.Y ^ 2) ^ (0.5)
                Dim ynrmly As Double = tempYaxis.Y / (tempYaxis.X ^ 2 + tempYaxis.Y ^ 2) ^ (0.5)

                'create coordinate system matrix
                Dim tempMat2d As New Matrix2dBuilder

                tempMat2d.ElementAt(0, 0) = xnrmlx
                tempMat2d.ElementAt(0, 1) = ynrmlx
                tempMat2d.ElementAt(0, 2) = orgPt.X
                tempMat2d.ElementAt(1, 0) = xnrmly
                tempMat2d.ElementAt(1, 1) = ynrmly
                tempMat2d.ElementAt(1, 2) = orgPt.Y
                tempMat2d.ElementAt(3, 0) = 0
                tempMat2d.ElementAt(3, 1) = 0
                tempMat2d.ElementAt(3, 2) = 1

                acTrans1.Commit()
                Return tempMat2d.ToMatrix2d

            End Using

        Catch ex As Exception
            Return Nothing
            Exit Function
        End Try

    End Function
    Public Function GetUCSMatrix3D(orgPt As Point3d, xAxPt As Point3d) As Matrix3d

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Dim NewUCSMatrix As Matrix3d

        Try
            Using acTrans1 As Transaction = DwgDB.TransactionManager.StartTransaction()

                Dim tempXaxis As Vector3d = orgPt.GetVectorTo(New Point3d(xAxPt.X, xAxPt.Y, orgPt.Z))
                Dim tempYaxis As Vector3d = tempXaxis.GetPerpendicularVector
                'Dim tempYaxis As Vector3d = tempXaxis.RotateBy(PI / 2, Vector3d.ZAxis)
                Dim tempZaxis As Vector3d = tempXaxis.CrossProduct(tempYaxis)
                'Dim tempZaxis As New Vector3d(0, 0, 1)

                'calculate components of unit vectors for new UCS matrix
                'x axis vector
                Dim xnrmlx As Double = tempXaxis.X / (tempXaxis.X ^ 2 + tempXaxis.Y ^ 2 + tempXaxis.Z ^ 2) ^ (0.5)
                Dim xnrmly As Double = tempXaxis.Y / (tempXaxis.X ^ 2 + tempXaxis.Y ^ 2 + tempXaxis.Z ^ 2) ^ (0.5)
                Dim xnrmlz As Double = tempXaxis.Z / (tempXaxis.X ^ 2 + tempXaxis.Y ^ 2 + tempXaxis.Z ^ 2) ^ (0.5)
                'y axis vector
                Dim ynrmlx As Double = tempYaxis.X / (tempYaxis.X ^ 2 + tempYaxis.Y ^ 2 + tempYaxis.Z ^ 2) ^ (0.5)
                Dim ynrmly As Double = tempYaxis.Y / (tempYaxis.X ^ 2 + tempYaxis.Y ^ 2 + tempYaxis.Z ^ 2) ^ (0.5)
                Dim ynrmlz As Double = tempYaxis.Z / (tempYaxis.X ^ 2 + tempYaxis.Y ^ 2 + tempYaxis.Z ^ 2) ^ (0.5)
                'z axis vector
                Dim znrmlx As Double = tempZaxis.X / (tempZaxis.X ^ 2 + tempZaxis.Y ^ 2 + tempZaxis.Z ^ 2) ^ (0.5)
                Dim znrmly As Double = tempZaxis.Y / (tempZaxis.X ^ 2 + tempZaxis.Y ^ 2 + tempZaxis.Z ^ 2) ^ (0.5)
                Dim znrmlz As Double = tempZaxis.Z / (tempZaxis.X ^ 2 + tempZaxis.Y ^ 2 + tempZaxis.Z ^ 2) ^ (0.5)

                'create coordinate system matrix
                Dim tempMat3d As New Matrix3dBuilder

                tempMat3d.ElementAt(0, 0) = xnrmlx
                tempMat3d.ElementAt(0, 1) = ynrmlx
                tempMat3d.ElementAt(0, 2) = znrmlx 'should be zero
                tempMat3d.ElementAt(0, 3) = orgPt.X
                tempMat3d.ElementAt(1, 0) = xnrmly
                tempMat3d.ElementAt(1, 1) = ynrmly
                tempMat3d.ElementAt(1, 2) = znrmly 'should be zero
                tempMat3d.ElementAt(1, 3) = orgPt.Y
                tempMat3d.ElementAt(2, 0) = xnrmlz 'should be zero
                tempMat3d.ElementAt(2, 1) = ynrmlz 'should be zero
                tempMat3d.ElementAt(2, 2) = znrmlz 'should be 1
                tempMat3d.ElementAt(2, 3) = orgPt.Z 'should be zero
                tempMat3d.ElementAt(3, 0) = 0
                tempMat3d.ElementAt(3, 1) = 0
                tempMat3d.ElementAt(3, 2) = 0
                tempMat3d.ElementAt(3, 3) = 1

                NewUCSMatrix = tempMat3d.ToMatrix3d

                acTrans1.Commit()
            End Using

        Catch ex As Exception
            Return Nothing
            Exit Function
        End Try

        Return NewUCSMatrix

    End Function

    Public Function GetNewUCSName(uName As String) As String

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor

        'store supplied name in a new variable
        Dim tempname As String = uName

        'start a transaction
        Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction()

            'Open the UCS table for
            Dim acUCSTbl As UcsTable
            acUCSTbl = DirectCast(acTrans.GetObject(DwgDB.UcsTableId, OpenMode.ForRead), UcsTable)

            'declare an index
            Dim i As Integer = 1

            'find a unique UCS name and show it to the user
            Do While acUCSTbl.Has(uName)
                uName = tempname & i
                i += 1
            Loop
            ed.WriteMessage(vbLf & "The new UCS name is " & uName)
UCSExists:
            acTrans.Commit()
        End Using
        Return uName

    End Function

    Public Function RegionCentroid(tReg As Region) As Point2d
        Try
            Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = CurDwg.Database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim curUCSMatrix As Matrix3d = ed.CurrentUserCoordinateSystem
            Dim curUCS As CoordinateSystem3d = curUCSMatrix.CoordinateSystem3d
            Dim tCentroid As Point2d = tReg.AreaProperties(curUCS.Origin, curUCS.Xaxis, curUCS.Yaxis).Centroid
            Dim tC As New Point2d(tCentroid.X, tCentroid.Y)
            Return tC
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function GetVertices(cvObId As ObjectId) As Point3dCollection

        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim dwgDB As Database = doc.Database

        Dim tempPts As New Point3dCollection

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()
            Dim obj As DBObject = acTrans.GetObject(cvObId, OpenMode.ForRead)

            If TypeOf obj Is Polyline Then
                Dim lwp As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(obj, Autodesk.AutoCAD.DatabaseServices.Polyline)
                If lwp IsNot Nothing Then
                    Dim vn As Integer = lwp.NumberOfVertices
                    For i As Integer = 0 To vn - 1
                        Dim pt As Point2d = lwp.GetPoint2dAt(i)
                        Dim pt3d As New Point3d(pt.X, pt.Y, 0)
                        tempPts.Add(pt3d)
                    Next
                End If

            ElseIf TypeOf obj Is Polyline2d Then
                Dim p2d As Polyline2d = TryCast(obj, Polyline2d)
                If p2d IsNot Nothing Then
                    For Each vId As ObjectId In p2d
                        Dim v2d As Vertex2d = CType(acTrans.GetObject(vId, OpenMode.ForRead), Vertex2d)
                        Dim pt As Point3d = v2d.Position
                        tempPts.Add(pt)
                    Next
                End If

            ElseIf TypeOf obj Is Polyline3d Then
                Dim p3d As Polyline3d = TryCast(obj, Polyline3d)
                If p3d IsNot Nothing Then
                    For Each vId As ObjectId In p3d
                        Dim v3d As PolylineVertex3d = CType(acTrans.GetObject(vId, OpenMode.ForRead), PolylineVertex3d)
                        Dim pt As Point3d = v3d.Position
                        tempPts.Add(pt)
                    Next
                End If
            Else
                ed.WriteMessage(vbLf & "Error.  Bad DBObject sent to GetVertices function.")
                tempPts = Nothing
            End If
            acTrans.Commit()

        End Using

        Return tempPts

    End Function

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

            Return plObjId

            'dispose of the temporary circle
            c1.Dispose()
            acTrans.Commit()

        End Using

    End Function
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
                ed.WriteMessage(vbLf & "Error calculating tangent points.")
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

Public Module PlotLayout

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

    Public Function GetPstyleFiles() As List(Of String)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim pList As New List(Of String)
        For Each pStyle As String In PlotSettingsValidator.Current.GetPlotStyleSheetList()
            pList.Add(pStyle)
        Next
        Return pList
    End Function

    Public Function LayoutTabList() As SortedDictionary(Of Integer, String)

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        Dim layAndTab As New SortedDictionary(Of Integer, String)
        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()
            Dim layDict As DBDictionary = dwgDB.LayoutDictionaryId.GetObject(OpenMode.ForRead)
            For Each entry As DBDictionaryEntry In layDict
                Dim lay As Layout = CType(entry.Value.GetObject(OpenMode.ForRead), Layout)
                If Not lay.LayoutName = "Model" Then
                    layAndTab.Add(lay.TabOrder, lay.LayoutName)
                End If
            Next
            acTrans.Commit()
        End Using
        Return layAndTab
    End Function

    Public Function GetSheetName(psetVal As PlotSettingsValidator, pset As PlotSettings) As String

        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
        Dim mySheet As String

        If Not String.IsNullOrEmpty(pset.CanonicalMediaName) Then

            Dim pkoMedia As New PromptKeywordOptions(vbLf & "Do you want to use the default media " & pset.CanonicalMediaName & "?")
            With pkoMedia
                .AllowArbitraryInput = False
                .Keywords.Add("Yes")
                .Keywords.Add("No")
                .AppendKeywordsToMessage = True
                .AllowNone = False
            End With

            Dim pkrMedia As PromptResult = ed.GetKeywords(pkoMedia)

            If pkrMedia.Status = PromptStatus.OK Then
                If pkrMedia.StringResult = "Yes" Then
                    mySheet = pset.CanonicalMediaName
                Else
                    mySheet = ""
                End If
            Else
                mySheet = Nothing
            End If
        Else
            Dim newPicker As New PlotSettingsForm
            Dim myPaperSizes As StringCollection = psetVal.GetCanonicalMediaNameList(pset)
            For Each nm As String In myPaperSizes
                newPicker.AddSetup(nm)
            Next
            newPicker.Label1.Text = "Select media for layout"
            newPicker.Text = "Media Picker"
            newPicker.ShowDialog()

            If newPicker.DialogResult = DialogResult.OK Then
                mySheet = newPicker.SelectedSetup
            Else
                mySheet = Nothing
            End If

            newPicker.Dispose()

        End If

        Return mySheet

    End Function

    Public Function GetPstylesFiles() As List(Of String)

        ' Get the current document and database, and start a transaction
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim psCol As New List(Of String)
        For Each plotStyleTable As String In PlotSettingsValidator.Current.GetPlotStyleSheetList()
            psCol.Add(plotStyleTable.ToString)
        Next
        Return psCol

    End Function

    Public Sub DumpPltStyleTables()

        Dim CurDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim DwgDB As Database = CurDwg.Database
        Dim ed As Editor = CurDwg.Editor

        Using actrans As Transaction = DwgDB.TransactionManager.StartTransaction

            Dim layMgr As LayoutManager = LayoutManager.Current
            Dim curLayout As Layout = actrans.GetObject(layMgr.GetLayoutId(layMgr.CurrentLayout), OpenMode.ForRead)
            Dim plInfo As New PlotInfo() With {.Layout = curLayout.ObjectId}
            Dim tFileName As String = Path.GetTempFileName
            ed.WriteMessage(vbLf & tFileName)

            'Dim fS As New FileStream(tFileName, FileMode.Create, FileAccess.Write)
            Dim PSdict As DictionaryWithDefaultDictionary = actrans.GetObject(DwgDB.PlotStyleNameDictionaryId, OpenMode.ForRead)

            Dim output(2) As String
            Using sr As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(tFileName, False)
                sr.WriteLine(tFileName)
                For Each ps As DBDictionaryEntry In PSdict
                    output(0) = ps.Key
                    Dim ob As DBObject = actrans.GetObject(ps.Value, OpenMode.ForRead)
                    Dim typeNm As String = ob.GetType().ToString
                    output(1) = typeNm
                    output(2) = ps.Value.ToString
                    Dim prtstr = Join(output, ",")
                    sr.WriteLine(prtstr)
                Next
            End Using

            Dim newFileName As String = GetFileNameSameAsDWG(".txt")
            Dim uR As DialogResult
TryAgain:
            If File.Exists(newFileName) Then
                uR = MessageBox.Show("Default File at path " & vbLf & newFileName & vbLf & " Exists. Do you want to replace it?", "File Exists", MessageBoxButtons.YesNoCancel)
                If uR = DialogResult.Cancel Then
                    Kill(tFileName)
                    Exit Sub
                ElseIf uR = DialogResult.No Then
                    newFileName = GetASaveFileName("*.txt")
                    GoTo TryAgain
                ElseIf uR = DialogResult.Yes Then
                    Kill(newFileName)
                    File.Copy(tFileName, newFileName)
                    Kill(tFileName)
                End If
            Else
                File.Copy(tFileName, newFileName)
                Kill(tFileName)
            End If

        End Using

    End Sub

    Public Function CreateVP(vtr As ViewTableRecord, vpLayerName As String, hsize As Double, vsize As Double, acTrans As Transaction, pset As PlotSettings, sheetNm As String) As ObjectId

        'used by the ImageMasterViews sub

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = curDwg.Editor
        Dim dwgDB As Database = curDwg.Database

        Dim lm As LayoutManager = LayoutManager.Current
        Dim layId As ObjectId = lm.CreateLayout(vtr.Name)

        lm.CurrentLayout = vtr.Name
        Dim myptr = pset.PlotConfigurationName

        Dim blkTbl As BlockTable = acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead)
        Dim curSpace As BlockTableRecord = acTrans.GetObject(dwgDB.CurrentSpaceId, OpenMode.ForWrite)
        'Dim curSpace As BlockTableRecord = acTrans.GetObject(blkTbl(vtr.Name), OpenMode.ForWrite)

        Dim lo As Layout = acTrans.GetObject(layId, OpenMode.ForWrite)

        Dim vpIDs As ObjectIdCollection = lo.GetViewports
        Dim vp As Autodesk.AutoCAD.DatabaseServices.Viewport

        If vpIDs.Count > 0 Then
            vp = acTrans.GetObject(vpIDs(1), OpenMode.ForWrite)
            vp.SetDatabaseDefaults()
            vp.CenterPoint = New Point3d(hsize / 2, vsize / 2, 0)
            vp.Height = vsize
            vp.Width = hsize
        Else
            vp = New Autodesk.AutoCAD.DatabaseServices.Viewport
            vp.SetDatabaseDefaults()
            vp.CenterPoint = New Point3d(hsize / 2, vsize / 2, 0)
            vp.Height = vsize
            vp.Width = hsize
            curSpace.AppendEntity(vp)
            acTrans.AddNewlyCreatedDBObject(vp, True)
        End If

        vp.On = True

        If Not String.IsNullOrEmpty(vpLayerName) Then vp.Layer = vpLayerName

        lo.CopyFrom(pset)
        ed.SwitchToModelSpace()
        ed.SetCurrentView(vtr)
        ed.SwitchToPaperSpace()

        Dim pSetVal As PlotSettingsValidator = PlotSettingsValidator.Current
        pSetVal.SetPlotConfigurationName(pset, myptr, sheetNm)
        pSetVal.SetPlotType(pset, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout)
        'pSetVal.SetPlotRotation(pset, PlotRotation.Degrees000)
        pSetVal.SetZoomToPaperOnUpdate(pset, True)

        Return layId
    End Function

    Public Function GetPlotSetup() As String

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database

        Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()
            Dim plsets As DBDictionary = acTrans.GetObject(dwgDB.PlotSettingsDictionaryId, OpenMode.ForWrite)
            Dim SetPicker As New PlotSettingsForm
            For Each plSet As DBDictionaryEntry In plsets
                SetPicker.AddSetup(plSet.Key)
            Next

            SetPicker.ShowDialog()

            Dim mySetup As String

            If SetPicker.DialogResult = DialogResult.OK Then
                mySetup = SetPicker.SelectedSetup
            Else
                mySetup = ""
            End If
            acTrans.Commit()

            Return mySetup

        End Using

    End Function

    Public Function GetPrinter(psetval As PlotSettingsValidator) As String

        Dim myPlotters As StringCollection = psetval.GetPlotDeviceList
        Dim myPtr As String

        If myPlotters.Count <= 0 Then
            myPtr = ""
        Else
            Dim pkr As New Picker
            With pkr
                .TopLabel.Text = "Select Printer Definition"
                .BxList.SelectionMode = System.Windows.Forms.SelectionMode.One
            End With

            For i As Integer = 0 To myPlotters.Count - 1
                pkr.BxList.Items.Add(myPlotters(i))
            Next

            pkr.ShowDialog()

            If pkr.DialogResult = DialogResult.OK Then
                Dim myColl As Collection = pkr.PickCol
                If myColl.Count = 1 Then
                    myPtr = pkr.PickCol(1).ToString
                Else
                    myPtr = ""
                End If
            Else
                myPtr = ""
            End If

        End If

        Return myPtr

    End Function

End Module


