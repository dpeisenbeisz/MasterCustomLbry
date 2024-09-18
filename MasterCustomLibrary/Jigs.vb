Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry

Public Class LineJigger
    Inherits EntityJig

    Public mEndPoint As Point3d

    Public Sub New(ByVal ent As Line)
        MyBase.New(ent)
    End Sub

    Protected Overrides Function Sampler(prompts As JigPrompts) As SamplerStatus

        Dim prOptions1 As New JigPromptPointOptions(vbLf & "Next point:")

        With prOptions1
            .BasePoint = CType(Entity, Line).StartPoint
            .UseBasePoint = True
            .UserInputControls = UserInputControls.Accept3dCoordinates Or UserInputControls.AnyBlankTerminatesInput Or UserInputControls.GovernedByOrthoMode Or UserInputControls.GovernedByUCSDetect Or UserInputControls.UseBasePointElevation Or UserInputControls.InitialBlankTerminatesInput Or UserInputControls.NullResponseAccepted
        End With

        Dim prResult1 As PromptPointResult = prompts.AcquirePoint(prOptions1)
        If prResult1.Status = PromptStatus.Cancel Then Return SamplerStatus.Cancel

        If prResult1.Value = mEndPoint Then
            Return SamplerStatus.NoChange
        Else
            mEndPoint = prResult1.Value
            Return SamplerStatus.OK
        End If
    End Function

    Public Shared Function lineJig() As Boolean
        Dim curDwg As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = curDwg.Editor
        Try
            Dim dwgDb As Database = curDwg.Database

            Dim ppr As PromptPointResult = ed.GetPoint(vbLf & "Start point")
            If ppr.Status <> PromptStatus.OK Then Return False
            Dim pt As Point3d = ppr.Value
            Dim ent As New Line(pt, pt)

            ent.TransformBy(ed.CurrentUserCoordinateSystem)
            Dim jigger As New LineJigger(ent)
            Dim pr As PromptResult = ed.Drag(jigger)

            If pr.Status = PromptStatus.OK Then

                Using acTrans As Transaction = dwgDb.TransactionManager.StartTransaction()
                    Dim blktbl As BlockTable = CType(acTrans.GetObject(dwgDb.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim btr As BlockTableRecord = CType(acTrans.GetObject(blktbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                    btr.AppendEntity(jigger.Entity)
                    acTrans.AddNewlyCreatedDBObject(jigger.Entity, True)
                    acTrans.Commit()
                End Using
            Else
                ent.Dispose()
                Return False
            End If
            Return True
        Catch ex As Exception
            ed.WriteMessage(ex.Message)
            Return False
        End Try
    End Function

    Protected Overrides Function Update() As Boolean
        CType(Entity, Line).EndPoint = mEndPoint
        Return True
    End Function





End Class



Public Class Testers

    <CommandMethod("TestEntityJigger12")>
    Public Shared Sub TestEntityJigger12_Method()
        Dim curDwg As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = curDwg.Editor
        Dim dwgDb As Database = curDwg.Database

        If LineJigger.lineJig() Then
            ed.WriteMessage(vbLf & "A line segment has been successfully jigged and added to the database." & vbLf)
        Else
            ed.WriteMessage(vbLf & "It failed to jig and add a line segment to the database." & vbLf)
        End If
    End Sub



End Class





