Imports System.Math
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports System.IO
Imports System.Text
Imports System.Xml.Serialization
Imports System.Xml
Imports Autodesk.AutoCAD.Colors
Imports MasterCustomLibrary.AcdLayerControl

Namespace CustomLibrary
    Public Module Macros
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
                        Dim lName As String = lTR.Name
                        Dim aLayer As New AcdLayer(lName)
                        lyrs.Add(aLayer)
                    End If
                Next
            End Using
            FormXML(lyrs)
        End Sub

        Private Sub FormXML(lyrs As AcdLayerControl.AcdLayers)
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
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
                .IndentChars = "    "
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
                    Dim testname As String = aLayer.Name
                    If lTbl.Has(aLayer.Name) Then
                        Dim lTR As LayerTableRecord = TryCast(actrans.GetObject(lTbl(aLayer.Name), OpenMode.ForRead), LayerTableRecord)
                        If lTR = lz Then GoTo Skip
                        If lTR.Name = testname Then
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
                                .Transparency = GetTransparency(aLayer.Transparency)
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
                            .Transparency = GetTransparency(aLayer.Transparency)
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

        Private Function GetTransparency(trans As Double) As Transparency
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database
            Dim alpha As Byte = 255 * ((100 - trans) / 100)
            Dim tp As New Transparency(alpha)
            Return tp
        End Function

        Private Function GetColor(colorStr As String) As Color
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
        Private Function GetLTId(ltName As String) As ObjectId
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

        Private Function GetMatId(ltName As String) As ObjectId
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

        <CommandMethod("TESTANGLES")>
        Public Sub TestAngles()

            Dim degs As Long = 263
            Dim mins As Long = 32
            Dim secs As Double = 32.75
            Dim decDegs = degs + (mins / 60) + (secs / 3600)
            Dim srvy As String = "S6d27'27.25W"
            'Dim srvy As String = "S6d32'32.75E"
            Dim nm As String = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name

            Dim pth As String = Path.GetDirectoryName(nm)
            Dim tFileName As String = pth & "\" & "TempFile"
            Dim angob1 As New AngleObj(decDegs, False, 270)
            WriteStuff(angob1, "50-degree angle object", tFileName & "1")
            Dim angRad As Double = decDegs * PI / 180
            Dim angob2 As New AngleObj(angRad, True, (3 * PI / 2))
            WriteStuff(angob2, "radian angle object", tFileName & "2")
            Dim angTxt As String = srvy
            Dim angob3 As New AngleObj(angTxt, 270)
            WriteStuff(angob3, srvy & " Bearing object", tFileName & "3")
            Dim angob4 As New AngleObj(degs, mins, secs, 270)
            WriteStuff(angob4, "309, 32, 32.75 angle object", tFileName & "4")
            Dim angob5 As AngleObj = angob1.AddAngle(angob2)
            WriteStuff(angob5, "Add 1 and 2", tFileName & "5")
            Dim angob6 As AngleObj = angob5.SubtractAngle(angob1)
            WriteStuff(angob6, "Subtract 1 from 5", tFileName & "6")
        End Sub


        Private Sub WriteStuff(ao As AngleObj, IDstring As String, tfilename As String)

            Dim tf As String = Path.ChangeExtension(tfilename, ".txt")

            Using sW As New System.IO.StreamWriter(tf)
                With sW
                    .WriteLine(IDstring)
                    .WriteLine("Decimal Degrees: " & ao.DecimalDegrees.ToString)
                    .WriteLine("Radians:  " & ao.Radians.ToString)
                    .WriteLine("DMS:  " & ao.DMS)
                    .WriteLine("Surveyors:  " & ao.Surveyors())
                    .WriteLine("Degs:  " & ao.Degrees.ToString)
                    .WriteLine("Minutes:  " & ao.Minutes.ToString)
                    .WriteLine("Seconds:  " & ao.Seconds.ToString)
                    .WriteLine("Autocad DMS: " & ao.AutoCADdms)
                    .WriteLine("Azimuth:  " & ao.DecAzimuth)
                    .WriteLine("RadAzimuth:  " & ao.RadAzimuth)
                    .WriteLine("BaseAngle:  " & ao.BaseAngle)
                    .WriteLine("RadBase:  " & ao.RadBase)
                    .WriteLine(vbCrLf)
                End With
            End Using


        End Sub

        <CommandMethod("LISTACARCDATA")>
        Public Sub ListAcArcData()
            Using curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim ed As Editor = curDwg.Editor
                Dim dwgDB As Database = curDwg.Database

                Dim myArcId As ObjectId
                Dim pp As Point3d
                Dim tanPT As Point3d

                Dim peo As New PromptEntityOptions(vbLf & "Select arc entity")
                With peo
                    .SetRejectMessage(vbLf & "Selected entity must be an arc.")
                    .AddAllowedClass(GetType(Arc), True)
                End With
                Dim peR As PromptEntityResult = ed.GetEntity(peo)
                If peR.Status = PromptStatus.OK Then
                    myArcId = peR.ObjectId
                    pp = peR.PickedPoint
                Else
                    ed.WriteMessage(vbLf & "Command Cancelled.")
                    Exit Sub
                End If

                Dim sb As New StringBuilder

                Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                    Dim myArc As Arc = acTrans.GetObject(myArcId, OpenMode.ForRead)

                    With sb
                        .AppendLine(vbLf & "Start Point: " & myArc.StartPoint.ToString)
                        .AppendLine("End Point: " & myArc.EndPoint.ToString)
                        .AppendLine("Center: " & myArc.Center.ToString)
                        .AppendLine("Total Angle: " & myArc.TotalAngle.ToString)
                        .AppendLine("Radius: " & myArc.Radius.ToString)
                        .AppendLine("Length: " & myArc.Length.ToString)
                        .AppendLine("Normal: " & myArc.Normal.ToString)
                        .AppendLine("Start Angle: " & myArc.StartAngle.ToString)
                        .AppendLine("End Angle: " & myArc.EndAngle.ToString)

                        Dim midPt As Point3d = myArc.GetPointAtParameter(myArc.GetParameterAtDistance(myArc.Length / 2))
                        Dim tpt1 As New Point2d(myArc.StartPoint.X, myArc.StartPoint.Y)
                        Dim tpt2 As New Point2d(myArc.EndPoint.X, myArc.EndPoint.Y)
                        Dim tpt3 As New Point2d(midPt.X, midPt.Y)
                        Dim cA As New CircularArc2d(tpt1, tpt2, tpt3)
                        Dim crv As Curve2d = TryCast(cA, Curve2d)

                        Dim myPolyID As ObjectId = Arc2poly(myArc.ObjectId)
                        Dim myPoly As Polyline = acTrans.GetObject(myPolyID, OpenMode.ForRead)
                        Dim d1 As Vector3d = myPoly.GetFirstDerivative(myArc.GetDistanceAtParameter(myArc.GetParameterAtPoint(tanPT)))
                        .AppendLine("1st Derivative: " & d1.ToString)
                        Dim d2 As Vector3d = myPoly.GetSecondDerivative(myArc.GetDistanceAtParameter(myArc.GetParameterAtPoint(tanPT)))
                        .AppendLine("2nd Derivative: " & d2.ToString)
                    End With
                End Using
                ed.WriteMessage(vbLf & sb.ToString & vbLf)
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

        Public Function OffsetSideByPoint(clineID As ObjectId, offdist As Double, Optional PrptMessage As String = "") As Integer

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database

            'declare return variable
            Dim offside As Integer

            'start transaction
            Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()

                Try
                    'get reference point
                    Dim ppOPts As New PromptPointOptions("")
                    With ppOPts
                        If PrptMessage = "" Then
                            .Message = (vbLf & "Select any point on the side of the baseline where the offsets are positive (+Y).")
                        Else
                            .Message = PrptMessage
                        End If
                        .AllowNone = False
                    End With

                    Dim ppR As PromptPointResult = ed.GetPoint(ppOPts)
                    Dim pRef As New Point3d

                    'if user cancels, return 10
                    If ppR.Status = PromptStatus.Cancel Then
                        ed.WriteMessage(vbLf & "User Cancelled.")
                        Return 10
                        Exit Function
                    Else
                        pRef = ppR.Value
                    End If

                    'try to cast picked object as a curve.  If it doesn't work, then it can't be offset
                    Using pl As Curve = TryCast(acTrans.GetObject(clineID, OpenMode.ForRead), Curve)
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
                            ed.WriteMessage(vbLf & "Selected entity cannot be used as baseline.  End Command.")
                            Return 10
                            Exit Function
                        End If
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Fatal Error." & ex.Message)
                    Return 10
                End Try
                acTrans.Dispose()
            End Using
            Return offside
        End Function

        Public Function BlkExists(bNameStr As String) As Boolean

            'function to test for existing block by its name
            Try
                Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim dwgDB As Database = curDwg.Database

                Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction()

                    'open the block table for read
                    Dim acBlkTbl As BlockTable = DirectCast(acTrans.GetObject(dwgDB.BlockTableId, OpenMode.ForRead), BlockTable)

                    'test for block name and return true if it is found
                    If acBlkTbl.Has(bNameStr) Then
                        Return True
                    Else
                        Return False
                    End If
                    'nothing left to do, dispose of the transaction
                    acTrans.Dispose()
                End Using

            Catch ex As Exception
                Return False
            End Try

        End Function

        Public Function Arc2poly(arcID As ObjectId) As ObjectId
            Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            'Dim ed As Editor = acDwg.Editor
            Dim dwgDB As Database = acDwg.Database
            Dim poly As Polyline

            Try
                poly = New Polyline()
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
            If deltaAng < 0 Then deltaAng += 2 * Math.PI
            Return Math.Tan(deltaAng * 0.25)
        End Function

        Public Function IsArc(objId As ObjectId) As Boolean
            Dim acDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = acDwg.Database

            Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim blineObj As Object = TryCast(actrans.GetObject(objId, OpenMode.ForRead), Object)
                Dim objtype As Type = blineObj.GetType
                If objtype = GetType(Arc) Then
                    Return True
                Else
                    Return False
                End If
                actrans.Commit()
            End Using

        End Function

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
                    MsgBox("Error. XML file not selected")
                    Return ""
                    Exit Function

                ElseIf xmlResult = DialogResult.OK Then
                    fName = fDialog.FileName
                Else
                    MsgBox("Error. XML file not selected")
                    Return ""
                    Exit Function
                End If

skipit:
                'return the name and path of the file
                Return fName

            Catch ex As Exception
                MsgBox(ex.Message & vbLf & "Error selecting XML file.")
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
                    MsgBox("Error. CSV file not selected")
                    Return Nothing
                    Exit Function

                ElseIf xmlResult = DialogResult.OK Then
                    fName = fDialog.FileName
                Else
                    MsgBox("Error. CSV file not selected")
                    Return ""
                    Exit Function
                End If

skipit:
                'return the name and path of the file
                Return fName

            Catch ex As Exception
                MsgBox(ex.Message & vbLf & "Error selecting CSV file.")
                Return ""
                Exit Function
            End Try

        End Function


        Public Function GetMyFolderName() As String

            'function for getting folder name as string

            Try
                Dim fName As String = ""
                Dim retval As String

                'declare a new open file dialog
                Dim fDialog As New FolderBrowserDialog()
                With fDialog
                    .Reset()
                    .ShowNewFolderButton = True
                End With

                'get the dialog result or return nothing
                Dim fldrResult As DialogResult = fDialog.ShowDialog()

                If fldrResult = DialogResult.Cancel Then
                    retval = ""
                ElseIf fldrResult = DialogResult.OK Then
                    fName = fDialog.SelectedPath
                    retval = fName
                Else
                    MsgBox("Error. Folder not selected")
                    retval = ""
                End If

                'return the name and path of the file
                Return retval

            Catch ex As Exception
                MsgBox(ex.Message & vbLf & "Error selecting folder.")
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
                    MsgBox("Error. XSD file not selected")
                    Return ""
                    Exit Function

                ElseIf xsdResult = DialogResult.OK Then
                    fName = fDialog.FileName
                Else
                    MsgBox("Error. XSD file not selected")
                    Return ""
                    Exit Function
                End If

skipit:
                'return the name and path of the file
                Return fName

            Catch ex As Exception
                MsgBox(ex.Message & vbLf & "Error selecting XSD file.")
                Return ""
                Exit Function
            End Try

        End Function
        Public Function CRads(degs As Double) As Double
            Return (Math.PI / 180) * degs
        End Function

        Public Function CDegs(rads As Double) As Double
            Return (180 / Math.PI) * rads
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

        Public Function SelBlockRef() As ObjectId
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor
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
                Return ObjectId.Null
                Exit Function
            End If

            'Dim blkRef As BlockReference
            'Dim blkName As String

            'Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
            '    blkRef = DirectCast(acTrans.GetObject(blkobjID, OpenMode.ForRead), BlockReference)
            '    blkName = blkRef.Name
            '    acTrans.Commit()
            'End Using
            Return blkobjID
        End Function

        Public Function GetUCSMatrix3D(orgPt As Point3d, xAxPt As Point3d) As Matrix3d

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim NewUCSMatrix As Matrix3d

            Try
                Using acTrans1 As Transaction = DwgDB.TransactionManager.StartTransaction()

                    Dim tempXaxis As Vector3d = orgPt.GetVectorTo(New Point3d(xAxPt.X, xAxPt.Y, orgPt.Z))
                    Dim tempYaxis As Vector3d = tempXaxis.GetPerpendicularVector
                    'Dim tempZaxis As Vector3d = tempXaxis.CrossProduct(tempYaxis)
                    Dim tempZaxis As New Vector3d(0, 0, 1)

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

        Public Function MakeNewUCS(orgPt As Point3d, xAxPt As Point3d, ucsName As String) As Boolean

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

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
                        ed.WriteMessage("UCS Exists.  UCS not created.")
                        Return False
                        Exit Function
                    End If

                    tempUcsRec.Origin = orgPt
                    Dim tempvec As Vector3d = orgPt.GetVectorTo(xAxPt)
                    Dim WZaxis As New Vector3d(0, 0, 1)
                    Dim WXaxis As New Vector3d(1, 0, 0)
                    Dim WYaxis As New Vector3d(0, 1, 0)

                    Dim temporthoY As Vector3d = (tempvec.GetPerpendicularVector)
                    Dim tempxunit As Vector3d = UnitVector3d(tempvec)
                    Dim temporthoX As Vector3d = (tempvec.DotProduct(WXaxis) / tempvec.DotProduct(tempvec)) * tempvec
                    Dim temporthoZ As Vector3d = (tempvec.DotProduct(WZaxis) / tempvec.DotProduct(tempvec)) * tempvec

                    'Dim tempXaxis As New Vector3d(tempvec.X, tempvec.Y,0)
                    Dim tempXaxis As Vector3d = temporthoX
                    'Dim tempZaxis As New Vector3d(0, 0, 1) 'tempXaxis.TransformBy(Matrix3d.Rotation(PI / 2, tempXaxis, orgPt))
                    Dim tempZaxis As Vector3d = temporthoZ

                    'Dim tempYaxis As Vector3d = tempXaxis.TransformBy(Matrix3d.Rotation(PI / 2, tempZaxis,
                    'orgPt))

                    'Dim tempYaxis As Vector3d = tempXaxis.TransformBy(Matrix3d.Rotation(PI / 2, tempZaxis, orgPt))
                    Dim tempYaxis As Vector3d = temporthoY

                    'calculate components of x & y axis unit vectors for new UCS matrix (z axis = 0)
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
                        MsgBox("Orthogonal Vectors")
                    Else
                        MsgBox("not orthogonal. dot product = " & tempValue.ToString)
                        Return False
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
                    curDwg.Dispose()
                    DwgDB.Dispose()
                    Exit Function
                End Try

            End Using

            Return True

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

        <CommandMethod("testVects")>
        Public Sub TESTVECTS()

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

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
                curDwg.Dispose()
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
            MakeNewUCS(wPt1, wPt2, "testUCS")

            Dim newmatrix As Matrix3d = GetUCSMatrix3D(wPt1, wPt2)

        End Sub

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

        Public Sub ImportBlk(bPath As String, bName As String)
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

            Dim sourcefilename As String = bPath

            Try
                Using acTrans As Transaction = DwgDB.TransactionManager.StartTransaction
                    Using tmpDb As New Database(False, True)
                        tmpDb.ReadDwgFile(sourcefilename, IO.FileShare.Read, True, Nothing)
                        DwgDB.Insert(bName, tmpDb, True)
                    End Using
                    acTrans.Commit()
                End Using

            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                ed.WriteMessage(vbLf & "Error in sub ImportBlk: " & ex.Message)
            End Try

        End Sub

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


        <CommandMethod("INSALL")>
        Public Sub InsertAllDwgFilesInFolder()

            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim DwgDB As Database = curDwg.Database
            Dim ed As Editor = curDwg.Editor

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

        Public Sub DeleteMyLayer(lName As String)

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

                    Dim DelList As ObjectIdCollection
                    DelList = GetEntitiesOnLayer(lName)

                    For k As Integer = 0 To DelList.Count - 1
                        Dim myEnt As Entity = CType(acTrans.GetObject(DelList(k), OpenMode.ForWrite), Entity)
                        If myEnt IsNot Nothing Then
                            myEnt.Erase(True)
                        End If
                    Next
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
                acTrans.Commit()
            End Using
        End Sub
        Public Function GetEntitiesOnLayer(ByVal layerName As String) As ObjectIdCollection
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim tvs As TypedValue() = New TypedValue(0) {New TypedValue(CInt(DxfCode.LayerName), layerName)}
            Dim sf As New SelectionFilter(tvs)
            Dim psr As PromptSelectionResult = ed.SelectAll(sf)

            If psr.Status = PromptStatus.OK Then
                Return New ObjectIdCollection(psr.Value.GetObjectIds())
            Else
                Return New ObjectIdCollection()
            End If

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

    End Module

End Namespace
