'(C) David Eisenbeisz 2023
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
'Imports Autodesk.AutoCAD.Runtime
Imports System.Text
'Imports System.Xml
'Imports System.Xml.Serialization
'Imports System.Xml.Schema

Public Class BlkViewPanel
    Inherits Form

    'Private ObjIds As ObjectIdCollection
    Private ReadOnly m_acBlocks As AcBlocks
    Private m_FullList As Boolean
    Private m_xmlFileName As String
    Private m_BlkDS As DataSet
    Private ReadOnly m_dialogResult As DialogResult
    Const m_xsdFileResource As String = "MasterCustomLibrary.BlkInfoSchema.xsd"

    Public Property FullList As Boolean
        Get
            Return m_FullList
        End Get
        Set(value As Boolean)
            m_FullList = value
        End Set
    End Property
    Public Property XmlFileName As String
        Get
            Return m_xmlFileName
        End Get
        Set(value As String)
            m_xmlFileName = value
        End Set
    End Property
    Public Property BlkDS As DataSet
        Get
            Return m_BlkDS
        End Get
        Set(value As DataSet)
            m_BlkDS = value
        End Set
    End Property

    Public ReadOnly Property Result As DialogResult
        Get
            Return m_dialogResult
        End Get
    End Property

    Public Sub CreateDGVxml()

        Dim blockDS As New DataSet("Block_Info")
        Dim schemaFilePath As String = AcCommon.GetEmbeddedResource(m_xsdFileResource)

        blockDS.ReadXmlSchema(schemaFilePath)
        blockDS.ReadXml(m_xmlFileName, XmlReadMode.Auto)

        For Each col As DataGridViewColumn In DGV1.Columns
            col.DataPropertyName = col.Name
        Next

        m_BlkDS = blockDS

        DGV1.AutoGenerateColumns = False
        DGV1.DataSource = blockDS
        DGV1.DataMember = "BlockInfo"

        Kill(schemaFilePath)

    End Sub

    Private Sub ButOk_Click(sender As Object, e As EventArgs) Handles ButOk.Click
        'm_dialogResult = DialogResult.OK
        If IO.File.Exists(m_xmlFileName) Then Kill(m_xmlFileName)
        Me.Close()
    End Sub

    Private Sub ButCancel_Click(sender As Object, e As EventArgs)
        'm_dialogResult = DialogResult.Cancel
        If IO.File.Exists(m_xmlFileName) Then Kill(m_xmlFileName)
        Me.Close()
    End Sub

    Private Sub BlockViewPanel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Private Sub DGV1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV1.CellContentClick

        Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim dwgDB As Database = curDwg.Database
        Dim ed As Editor = curDwg.Editor

        If DGV1.CurrentCell Is Nothing OrElse e.RowIndex < 0 Then
            Exit Sub
        Else
            Dim rowIdx As Integer = e.RowIndex
            Dim colIdx As Integer = e.ColumnIndex
            Dim colName As String = DGV1.Columns(colIdx).Name

            Dim rH As String = DGV1(DGV1.Columns("RefHandle").Index, rowIdx).Value
            Dim longRH As Long = Convert.ToInt64(rH, 16)
            Dim hand As New Handle(longRH)
            Dim refId As ObjectId = dwgDB.GetObjectId(False, hand, 0)

            If colName = "RefName" Then

                Dim sb As New StringBuilder

                'If AcCommon.OtherMethods.IsInModel Then
                Dim vtr As ViewTableRecord = ed.GetCurrentView
                    Dim vRatio As Double = (vtr.Width / vtr.Height)
                    Debug.Print("VTRwidth:  " & vtr.Width.ToString)
                    Debug.Print("VTRheight:  " & vtr.Height.ToString)

                    Dim bRef As BlockReference

                    Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction
                        Dim bRefObj As DBObject = TryCast(acTrans.GetObject(refId, OpenMode.ForRead), DBObject)
                        If TypeOf bRefObj Is BlockReference Then
                            bRef = TryCast(CType(bRefObj, BlockReference), BlockReference)

                        Else
                            Exit Sub
                        End If
                        If bRef IsNot Nothing Then
                            Dim objExt As Extents3d = bRef.GeometricExtents
                            Dim minPt As Point3d = bRef.Bounds.Value.MinPoint
                            Dim maxPt As Point3d = bRef.Bounds.Value.MaxPoint

                            Dim wdth As Double = maxPt.X - minPt.X
                            Dim ht As Double = maxPt.Y - minPt.Y

                            Dim vWidth As Double
                            Dim vHt As Double
                            Dim sclFact As Double

                            If wdth >= ht Then
                                vWidth = wdth
                                vHt = vWidth / vRatio
                                Debug.Print("width controls.")
                                'If vWidth >= vtr.Width Then
                                sclFact = vtr.Width / vWidth
                                'Else
                                'sclFact = vtr.Width / vWidth
                                'End If
                            Else
                                vHt = ht
                                vWidth = vHt * vRatio
                                Debug.Print("Ht controls.")
                                'If vHt >= vtr.Height Then
                                'sclFact = vtr.Height / vHt
                                'Else
                                sclFact = vtr.Height / vHt
                                'End If
                            End If

                            Debug.Print("SclFact:  " & sclFact.ToString)

                            'Dim ln As New Polyline(2)
                            'ln.AddVertexAt(0, New Point2d(objExt.MinPoint.X, objExt.MinPoint.Y), 0, 0, 0)
                            'ln.AddVertexAt(0, New Point2d(objExt.MaxPoint.X, objExt.MaxPoint.Y), 0, 0, 0)

                            'Dim midp As Point3d = ln.GetPointAtDist(ln.Length / 2)
                            Dim ctr As New Point3d(minPt.X + (Width / 2), minPt.Y + (ht / 2), 0)

                            Zoom(minPt, maxPt, Point3d.Origin, 1)

                            'ln.Dispose()

                            'Dim wdth As Double = 50
                            'Dim ht As Double = 50
                            'Dim ctr3d As Point3d = bRef.Position
                            'Dim ctrNew As New Point2d(ctr3d.X, ctr3d.Y)

                            'If ht > (wdth * vRatio) Then wdth = ht / vRatio
                            'Debug.Print(ctr.ToString)
                            'Debug.Print(midp.ToString)
                            'Debug.Print(insPt.ToString)
                            'Debug.Print(ctrNew.ToString)
                            Debug.Print(objExt.ToString)
                            Debug.Print(sclFact.ToString)

                            vtr.Dispose()

                            'Dim vtr2 As New ViewTableRecord
                            'With vtr2
                            '    .Name = "TestView"
                            '    '.CenterPoint = midp.Convert2d(New Plane(Point3d.Origin, Vector3d.ZAxis))
                            '    .CenterPoint = ctr
                            '    .Height = ht
                            '    .Width = wdth
                            'End With

                            'Dim vt As ViewTable = acTrans.GetObject(dwgDB.ViewTableId, OpenMode.ForWrite)
                            'Dim vtID As ObjectId = vt.Add(vtr2)

                            'ed.SetCurrentView(vtr2)

                        End If
                        acTrans.Commit()
                    End Using

                    'ed.Regen()
EndIt:
                    'End If

                ElseIf colName = "HasAttributes" Then
                    Dim bref As BlockReference
                Using acTrans As Transaction = dwgDB.TransactionManager.StartTransaction

                    Dim bRefObj As DBObject = TryCast(acTrans.GetObject(refId, OpenMode.ForRead), DBObject)
                    If TypeOf bRefObj Is BlockReference Then
                        bref = TryCast(CType(bRefObj, BlockReference), BlockReference)
                    Else
                        Exit Sub
                    End If

                    Dim csvList As New List(Of String)

                    If bref IsNot Nothing Then
                        Dim atCol As AttributeCollection = TryCast(bref.AttributeCollection, AttributeCollection)
                        If atCol IsNot Nothing AndAlso atCol.Count > 0 Then
                            For Each atId As ObjectId In atCol
                                Dim atdbOb As DBObject = acTrans.GetObject(atId, OpenMode.ForRead)
                                If TypeOf atdbOb Is AttributeReference Then
                                    Dim at As AttributeReference = CType(atdbOb, AttributeReference)
                                    Dim csvline As String = at.Tag.ToString & "," & at.TextString
                                    csvList.Add(csvline)
                                End If
                            Next
                            Using attForm As New AttributeForm(csvList)
                                Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(attForm)
                            End Using
                        End If
                    End If
                End Using
            End If
        End If
    End Sub


    Public Sub Zoom(ByVal pMin As Point3d, ByVal pMax As Point3d,
ByVal pCenter As Point3d, ByVal dFactor As Double)

        '' Get the current document and database
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim nCurVport As Integer = System.Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"))
        '' Get the extents of the current space when no points 
        '' or only a center point is provided
        '' Check to see if Model space is current

        If acCurDb.TileMode = True Then
            If pMin.Equals(New Point3d()) = True And pMax.Equals(New Point3d()) = True Then
                pMin = acCurDb.Extmin
                pMax = acCurDb.Extmax
            End If
        Else
            '' Check to see if Paper space is current
            If nCurVport = 1 Then
                If pMin.Equals(New Point3d()) = True And pMax.Equals(New Point3d()) = True Then
                    pMin = acCurDb.Pextmin
                    pMax = acCurDb.Pextmax
                End If
            Else
                '' Get the extents of Model space
                If pMin.Equals(New Point3d()) = True And pMax.Equals(New Point3d()) = True Then
                    pMin = acCurDb.Extmin
                    pMax = acCurDb.Extmax
                End If
            End If
        End If
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction
            '' Get the current view
            Using acView As ViewTableRecord = acDoc.Editor.GetCurrentView()
                Dim eExtents As Extents3d
                '' Translate WCS coordinates to DCS
                Dim matWCS2DCS As Matrix3d
                matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection)
                matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) * matWCS2DCS
                matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist, acView.ViewDirection, acView.Target) * matWCS2DCS

                '' If a center point is specified, define the 
                '' min and max point of the extents
                '' for Center and Scale modes
                If pCenter.DistanceTo(Point3d.Origin) <> 0 Then
                    pMin = New Point3d(pCenter.X - (acView.Width / 2), pCenter.Y - (acView.Height / 2), 0)
                    pMax = New Point3d((acView.Width / 2) + pCenter.X, (acView.Height / 2) + pCenter.Y, 0)
                End If
                '' Create an extents object using a line
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
                '' Check to see if a center point was provided (Center and Scale modes)
                If pCenter.DistanceTo(Point3d.Origin) <> 0 Then
                    dWidth = acView.Width
                    dHeight = acView.Height
                    If dFactor = 0 Then
                        pCenter = pCenter.TransformBy(matWCS2DCS)
                    End If
                    pNewCentPt = New Point2d(pCenter.X, pCenter.Y)
                Else '' Working in Window, Extents and Limits mode
                    '' Calculate the new width and height of the current view
                    dWidth = eExtents.MaxPoint.X - eExtents.MinPoint.X
                    dHeight = eExtents.MaxPoint.Y - eExtents.MinPoint.Y
                    '' Get the center of the view
                    pNewCentPt = New Point2d(((eExtents.MaxPoint.X + eExtents.MinPoint.X) * 0.5), ((eExtents.MaxPoint.Y + eExtents.MinPoint.Y) * 0.5))
                End If
                '' Check to see if the new width fits in current window
                dHeight = dWidth / dViewRatio
                '' Resize and scale the view

                If dFactor <> 0 Then
                    acView.Height = dHeight * dFactor
                    acView.Width = dWidth * dFactor
                End If
                '' Set the center of the view
                acView.CenterPoint = pNewCentPt
                '' Set the current view
                acDoc.Editor.SetCurrentView(acView)
            End Using
            '' Commit the changes
            acTrans.Commit()
        End Using
    End Sub


End Class