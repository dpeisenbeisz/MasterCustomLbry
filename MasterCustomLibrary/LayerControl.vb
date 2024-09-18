Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports System.Xml.Serialization

Namespace AcCommon
    <XmlRoot("AcdLayers")>
    Public Class AcdLayers
        Implements ICollection

        Public LayerSetName As String
        Private Layers() As AcdLayer
        Public Sub New()
            MyBase.New
            ReDim Layers(-1)
        End Sub
        Public Sub New(setName As String)
            MyBase.New
            LayerSetName = setName
            ReDim Layers(-1)
        End Sub

        Public Sub Add(aLayer As AcdLayer, i As Integer)
            If i < Layers.Count - 1 Then
                Layers(i) = aLayer
            End If
        End Sub

        Public Sub Add(alayer As AcdLayer)
            ReDim Preserve Layers(Layers.Count)
            Layers(Layers.Count - 1) = alayer
        End Sub

        Default Public ReadOnly Property Item(ByVal index As Integer) As AcdLayer
            Get
                Return Layers(index)
            End Get
        End Property

        <XmlElement("AcdLayer")>
        Public ReadOnly Property AcdLayer(ByVal index As Integer) As AcdLayer
            Get
                Return Layers(index)
            End Get
        End Property
        Public ReadOnly Property Count As Integer Implements ICollection.Count
            Get
                Return Layers.Count
            End Get
        End Property

        Public ReadOnly Property SyncRoot As Object Implements ICollection.SyncRoot
            Get
                Return Me
            End Get
        End Property

        Public ReadOnly Property IsSynchronized As Boolean Implements ICollection.IsSynchronized
            Get
                Return False
            End Get
        End Property

        Public Sub CopyTo(aList As Array, index As Integer) Implements ICollection.CopyTo
            Layers.CopyTo(aList, index)
        End Sub

        Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return Layers.GetEnumerator
        End Function
    End Class

    Public Class AcdLayer
        'Private disposedValue As Boolean
        Private ReadOnly c_lTR As LayerTableRecord
        Private ReadOnly c_ltrID As String
        Private c_isUsed As Boolean
        Private c_nm As String
        Private c_isOff As Boolean
        Private c_viewPortVisibilityDefault As Boolean
        Private c_isFrzn As Boolean
        Private c_isLocked As Boolean
        Private c_color As String
        Private c_Linetype As String
        Private c_linetypeID As ObjectId
        Private c_materialID As String
        Private c_isPlottable As Boolean
        Private c_lw As String
        Private c_trans As String
        Private c_pStyle As String
        Private c_pstyleID As String
        Private c_desc As String
        Private c_newVpFreeze As Boolean
        Private c_isHidden As Boolean
        Private ReadOnly c_HasVPoverrides As Dictionary(Of ObjectId, Boolean)
        Private ReadOnly c_vpOverrides As Dictionary(Of ObjectId, LayerViewportProperties)
        Private ReadOnly c_HasOverrides As Boolean
        Private c_Remove As Boolean
        Private c_MergeWith As String
        Private c_hasEntities As Boolean
        Private c_EntityCount As Long

        Public Sub New()
            MyBase.New
        End Sub

        Public Sub New(nm As String)
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Using lTable As LayerTable = actrans.GetObject(dwgDB.LayerTableId, OpenMode.ForRead)
                    Using ltr As LayerTableRecord = actrans.GetObject(lTable(nm), OpenMode.ForRead)
                        With ltr
                            c_nm = .Name
                            c_ltrID = .ObjectId.ToString
                            c_isUsed = .IsUsed
                            c_isOff = .IsOff
                            c_viewPortVisibilityDefault = .ViewportVisibilityDefault
                            c_isFrzn = .IsFrozen
                            c_isLocked = .IsLocked
                            Dim tempName As String
                            Dim bkName As String
                            Dim ColorStr As String
                            If .Color.HasColorName Then
                                tempName = .Color.ColorName
                                If .Color.HasBookName Then
                                    bkName = .Color.BookName
                                    ColorStr = tempName & "_" & bkName
                                    c_color = ColorStr
                                End If
                            Else
                                c_color = .Color.ToString
                            End If
                            c_linetypeID = .LinetypeObjectId
                            c_Linetype = GetLTName(.LinetypeObjectId)
                            c_isPlottable = .IsPlottable
                            Dim linewt As Int64 = .LineWeight
                            Dim lwDesc As String = [Enum].GetName(GetType(LineWeight), linewt)  '[Enum].Parse(GetType(LineWeight), linewt, True)
                            c_lw = lwDesc
                            Dim trans As String = GetTransparencyIndex(ltr).ToString
                            'c_trans = Mid(trans, 2, Len(trans) - 2)
                            c_trans = trans
                            'If dwgDB.PlotStyleMode Then c_pStyle = .PlotStyleName
                            c_pStyle = .PlotStyleName
                            c_pstyleID = .PlotStyleNameId.ToString
                            c_desc = .Description
                            c_isHidden = .IsHidden
                            c_HasOverrides = .HasOverrides

                            Dim vps As ObjectIdCollection = GetVPIds()
                            For Each id As ObjectId In vps
                                If .HasViewportOverrides(id) Then
                                    c_HasVPoverrides.Add(id, .HasViewportOverrides(id))
                                    c_vpOverrides.Add(id, .GetViewportOverrides(id))
                                End If
                            Next

                        End With

                        Dim entityIDColl As ObjectIdCollection = GetEntitiesOnLayer(ltr.Name)
                        If entityIDColl IsNot Nothing AndAlso entityIDColl.Count > 0 Then
                            c_hasEntities = True
                            c_EntityCount = entityIDColl.Count
                        Else
                            c_EntityCount = 0
                            c_hasEntities = False
                        End If
                        actrans.Commit()
                    End Using
                End Using
            End Using
        End Sub
        Public Sub New(ltrID As ObjectId)
            MyBase.New
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim dwgDB As Database = curDwg.Database
            Dim layerID As ObjectId = ltrID
            Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim ltr As LayerTableRecord = TryCast(actrans.GetObject(layerID, OpenMode.ForRead), LayerTableRecord)
                With ltr
                    c_nm = .Name
                    c_ltrID = .ObjectId.ToString
                    c_isUsed = .IsUsed
                    c_isOff = .IsOff
                    c_viewPortVisibilityDefault = .ViewportVisibilityDefault
                    c_isFrzn = .IsFrozen
                    c_isLocked = .IsLocked
                    Dim tempName As String
                    Dim bkName As String
                    Dim ColorStr As String
                    If .Color.HasColorName Then
                        tempName = .Color.ColorName
                        If .Color.HasBookName Then
                            bkName = .Color.BookName
                            ColorStr = tempName & "_" & bkName
                            c_color = ColorStr
                        End If
                    Else
                        c_color = .Color.ToString
                    End If
                    c_Linetype = GetLTName(.LinetypeObjectId)
                    c_linetypeID = .LinetypeObjectId
                    c_isPlottable = .IsPlottable
                    c_lw = .LineWeight
                    Dim trans As String = GetTransparencyIndex(ltr).ToString
                    'c_trans = Mid(trans, 2, Len(trans) - 2)
                    c_trans = trans
                    'If dwgDB.PlotStyleMode Then c_pStyle = .PlotStyleName
                    c_pStyle = .PlotStyleName
                    c_pstyleID = .PlotStyleNameId.ToString
                    c_desc = .Description
                    c_isHidden = .IsHidden
                    c_HasOverrides = .HasOverrides
                    Dim vps As ObjectIdCollection = GetVPIds()
                    For Each id As ObjectId In vps
                        If .HasViewportOverrides(id) Then
                            c_HasVPoverrides.Add(id, .HasViewportOverrides(id))
                            c_vpOverrides.Add(id, .GetViewportOverrides(id))
                        End If
                    Next

                    Dim entityIDColl As ObjectIdCollection = GetEntitiesOnLayer(ltr.Name)
                    If entityIDColl IsNot Nothing And entityIDColl.Count > 0 Then
                        c_hasEntities = True
                        c_EntityCount = entityIDColl.Count
                    Else
                        c_EntityCount = 0
                        c_hasEntities = False
                    End If

                End With
            End Using
        End Sub

        Public Sub New(ltr As LayerTableRecord)

            With ltr
                Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                'Dim dwgDB As Database = curDwg.Database
                c_nm = .Name
                c_ltrID = .ObjectId.ToString
                c_isUsed = .IsUsed
                c_isOff = .IsOff
                c_viewPortVisibilityDefault = .ViewportVisibilityDefault
                c_isFrzn = .IsFrozen
                c_isLocked = .IsLocked
                Dim tempName As String
                Dim bkName As String
                Dim ColorStr As String
                If .Color.HasColorName Then
                    tempName = .Color.ColorName
                    If .Color.HasBookName Then
                        bkName = .Color.BookName
                        ColorStr = tempName & "_" & bkName
                        c_color = ColorStr
                    End If
                Else
                    c_color = .Color.ToString
                End If
                c_Linetype = GetLTName(.LinetypeObjectId)
                c_linetypeID = .LinetypeObjectId
                c_isPlottable = .IsPlottable
                c_lw = .LineWeight
                Dim trans As String = GetTransparencyIndex(ltr).ToString
                'c_trans = Mid(trans, 2, Len(trans) - 2)
                c_trans = trans
                'If curDwg.Database.PlotStyleMode Then c_pStyle = .PlotStyleName
                c_pStyle = .PlotStyleName
                c_pstyleID = .PlotStyleNameId.ToString
                c_desc = .Description
                c_isHidden = .IsHidden
                c_HasOverrides = .HasOverrides

                Dim vps As ObjectIdCollection = GetVPIds()
                For Each id As ObjectId In vps
                    If .HasViewportOverrides(id) Then
                        c_HasVPoverrides.Add(id, .HasViewportOverrides(id))
                        c_vpOverrides.Add(id, .GetViewportOverrides(id))
                    End If
                Next

                Dim entityIDColl As ObjectIdCollection = GetEntitiesOnLayer(ltr.Name)
                If entityIDColl IsNot Nothing And entityIDColl.Count > 0 Then
                    c_hasEntities = True
                    c_EntityCount = entityIDColl.Count
                Else
                    c_EntityCount = 0
                    c_hasEntities = False
                End If
            End With
            ltr.Dispose()
        End Sub

        <XmlAttribute("Name")>
        Public Property Name() As String
            Get
                Return c_nm
            End Get
            Set(ByVal value As String)
                c_nm = value
            End Set
        End Property

        <XmlAttribute("IsUsed")>
        Public Property IsUsed() As Boolean
            Get
                Return c_isUsed
            End Get
            Set(ByVal value As Boolean)
                c_isUsed = value
            End Set
        End Property

        <XmlAttribute("HasEntities")>
        Public Property HasEntities() As Boolean
            Get
                Return c_hasEntities
            End Get
            Set(ByVal value As Boolean)
                c_hasEntities = value
            End Set
        End Property

        <XmlAttribute("EntityCount")>
        Public Property EntityCount() As Long
            Get
                Return c_EntityCount
            End Get
            Set(ByVal value As Long)
                c_EntityCount = value
            End Set
        End Property

        <XmlAttribute("VpVisDefault")>
        Public Property VpVisDefault() As Boolean
            Get
                Return c_viewPortVisibilityDefault
            End Get
            Set(ByVal value As Boolean)
                c_viewPortVisibilityDefault = value
            End Set
        End Property
        <XmlAttribute("IsOff")>
        Public Property IsOff() As Boolean
            Get
                Return c_isOff
            End Get
            Set(ByVal value As Boolean)
                c_isOff = value
            End Set
        End Property

        <XmlAttribute("IsFrozen")>
        Public Property IsFrozen() As Boolean
            Get
                Return c_isFrzn
            End Get
            Set(ByVal value As Boolean)
                c_isFrzn = value
            End Set
        End Property

        <XmlAttribute("IsLocked")>
        Public Property IsLocked() As Boolean
            Get
                Return c_isLocked
            End Get
            Set(ByVal value As Boolean)
                c_isLocked = value
            End Set
        End Property

        <XmlAttribute("IsPlottable")>
        Public Property IsPlottable() As Boolean
            Get
                Return c_isPlottable
            End Get
            Set(ByVal value As Boolean)
                c_isPlottable = value
            End Set
        End Property

        <XmlIgnore()>
        Public ReadOnly Property HasOverrides() As Boolean
            Get
                Return c_HasOverrides
            End Get
        End Property

        <XmlAttribute("IsHidden")>
        Public Property IsHidden() As Boolean
            Get
                Return c_isHidden
            End Get
            Set(ByVal value As Boolean)
                c_isHidden = value
            End Set
        End Property

        <XmlIgnore()>
        Public ReadOnly Property HasVPOverrides(vpID As ObjectId) As Boolean
            Get
                Return c_HasVPoverrides(vpID)
            End Get
        End Property

        <XmlAttribute("NewViewportFreeze")>
        Public Property NewVpFreeze() As Boolean
            Get
                Return c_newVpFreeze
            End Get
            Set(ByVal value As Boolean)
                c_newVpFreeze = value
            End Set
        End Property
        <XmlAttribute("Remove")>
        Public Property Remove() As Boolean
            Get
                Return c_Remove
            End Get
            Set(ByVal value As Boolean)
                c_Remove = value
            End Set
        End Property


        <XmlAttribute("Color")>
        Public Property Color() As String
            Get
                Return c_color
            End Get
            Set(ByVal value As String)
                c_color = value
            End Set
        End Property

        <XmlAttribute("Linetype")>
        Public Property Linetype() As String
            Get
                Return c_Linetype
            End Get
            Set(ByVal value As String)
                c_Linetype = value
            End Set
        End Property

        <XmlAttribute("MergeWith")>
        Public Property MergeWith() As String
            Get
                Return c_MergeWith
            End Get
            Set(ByVal value As String)
                c_MergeWith = value
            End Set
        End Property

        <XmlIgnore()>
        Public Property MaterialId() As String
            Get
                Return c_materialID
            End Get
            Set(ByVal value As String)
                c_materialID = value
            End Set
        End Property

        <XmlAttribute("Lineweight")>
        Public Property LineWeight() As String
            Get
                Return c_lw
            End Get
            Set(ByVal value As String)
                c_lw = value
            End Set
        End Property

        <XmlAttribute("Transparency")>
        Public Property Transparency() As String
            Get
                Return c_trans
            End Get
            Set(ByVal value As String)
                c_trans = value
            End Set
        End Property

        <XmlAttribute("PlotStyle")>
        Public Property PlotStyle() As String
            Get
                Return c_pStyle
            End Get
            Set(ByVal value As String)
                c_pStyle = value
            End Set
        End Property

        <XmlIgnore()>
        Public Property PlotStyleID() As String
            Get
                Return c_pstyleID
            End Get
            Set(ByVal value As String)
                c_pstyleID = value
            End Set
        End Property

        <XmlAttribute("Description")>
        Public Property Description() As String
            Get
                Return c_desc
            End Get
            Set(ByVal value As String)
                c_desc = value
            End Set
        End Property
        '<XmlAttribute("Alpha")>
        'Public Property Alpha() As String
        '    Get
        '        Return c_transAlpha
        '    End Get
        '    Set(ByVal value As String)
        '        c_transAlpha = value
        '    End Set
        'End Property
        Private Function GetTransparencyIndex(ByVal ltr As LayerTableRecord) As Integer
            If Not ltr.Transparency.IsInvalid AndAlso ltr.Transparency.IsByAlpha Then
                Dim smallestDiff As Integer = Byte.MaxValue, foundValue As Integer = 0
                Dim transparencies As Dictionary(Of Integer, Byte) = TransToAlpha()
                Dim alpha As Byte = ltr.Transparency.Alpha
                For Each key As Integer In transparencies.Keys
                    Dim testval As Byte = transparencies(key)
                    If Abs(testval - alpha) = 0 Then
                        Debug.Print("Key = " & key)
                        Debug.Print("TestVal = " & transparencies(key))
                        foundValue = key
                        Exit For
                    ElseIf Abs(testval - alpha) < smallestDiff Then
                        Debug.Print("Key = " & key)
                        Debug.Print("TestVal = " & transparencies(key))
                        smallestDiff = Math.Abs(testval - alpha)
                        foundValue = key
                    End If
                Next
                Return foundValue
            Else
                Return 0
            End If
        End Function

        Private Function GetLTName(ltID As ObjectId) As String
            Dim curDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = curDwg.Editor
            Dim dwgDB As Database = curDwg.Database
            Dim ltName As String
            Using actrans As Transaction = dwgDB.TransactionManager.StartTransaction
                Dim ltTbl As LinetypeTable = actrans.GetObject(dwgDB.LinetypeTableId, OpenMode.ForRead)
                If ltID.IsValid Then
                    Dim ob As Object = actrans.GetObject(ltID, OpenMode.ForRead)
                    If TypeOf ob Is LinetypeTableRecord Then
                        Dim ltr As LinetypeTableRecord = actrans.GetObject(ltID, OpenMode.ForRead)
                        ltName = ltr.Name
                    Else
                        ltName = "Continuous"
                    End If
                Else
                    ltName = "Continuous"
                End If
            End Using
            Return ltName
        End Function
        Private Function GetVPIds() As ObjectIdCollection
            Dim Dwgdb As Database = HostApplicationServices.WorkingDatabase
            Dim vpIds As New ObjectIdCollection
            Using acTrans As Transaction = Dwgdb.TransactionManager.StartTransaction()
                Dim vpTable As ViewportTable = CType(acTrans.GetObject(Dwgdb.ViewportTableId, OpenMode.ForRead), ViewportTable)
                For Each id As ObjectId In vpTable
                    Dim vpTR As ViewportTableRecord = CType(acTrans.GetObject(id, OpenMode.ForRead), ViewportTableRecord)
                    Dim vpID As ObjectId = vpTR.ObjectId
                    vpIds.Add(vpID)
                Next
                acTrans.Commit()
            End Using
            Return vpIds
        End Function



        '    Protected Overridable Sub Dispose(disposing As Boolean)
        '        If Not disposedValue Then
        '            If disposing Then
        '                ' TODO: dispose managed state (managed objects)
        '            End If

        '            ' TODO: free unmanaged resources (unmanaged objects) and override finalizer
        '            ' TODO: set large fields to null
        '            disposedValue = True
        '        End If
        '    End Sub

        '    ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
        '    ' Protected Overrides Sub Finalize()
        '    '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        '    '     Dispose(disposing:=False)
        '    '     MyBase.Finalize()
        '    ' End Sub

        '    Public Sub Dispose() Implements IDisposable.Dispose
        '        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        '        Dispose(disposing:=True)
        '        GC.SuppressFinalize(Me)
        '    End Sub

        '    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
        '        Throw New NotImplementedException()
        '    End Function
    End Class

End Namespace
