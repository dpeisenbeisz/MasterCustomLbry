'(c) David Eisenbeisz 2023

Imports System.Xml.Serialization

<XmlRoot("AcBlocks")>
Public Class AcBlocks
    'Inherits CollectionBase
    Implements ICollection

    Public c_BlockSetName As String
    Private m_Blocks() As BlockInfo
    Public Sub New()
        MyBase.New
        ReDim m_Blocks(-1)
    End Sub
    Public Sub New(setName As String)
        MyBase.New
        c_BlockSetName = setName
        ReDim m_Blocks(-1)
    End Sub

    Public Sub Add(Blk As BlockInfo, i As Integer)
        If i < m_Blocks.Count - 1 Then
            m_Blocks(i) = Blk
        Else
            ReDim Preserve m_Blocks(i + 1)
            m_Blocks(i) = Blk
        End If
    End Sub

    Public Sub Add(blk As BlockInfo)
        ReDim Preserve m_Blocks(m_Blocks.Count)
        m_Blocks(m_Blocks.Count - 1) = blk
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As BlockInfo
        Get
            Return m_Blocks(index)
        End Get
    End Property

    <XmlElement("BlockInfo")>
    Public ReadOnly Property BlockInfo(ByVal index As Integer) As BlockInfo
        Get
            Return m_Blocks(index)
        End Get
    End Property
    Public ReadOnly Property Count As Integer Implements ICollection.Count
        Get
            Return m_Blocks.Count
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
        m_Blocks.CopyTo(aList, index)
    End Sub

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return m_Blocks.GetEnumerator
    End Function
    Private Function Compare(x As BlockInfo, y As BlockInfo) As Integer
        Return x.RefName.CompareTo(y.RefName)
    End Function
    Public Sub Sort()
        Dim tBlks() As BlockInfo = m_Blocks
        Array.Sort(Of BlockInfo)(tBlks, New Comparison(Of BlockInfo)(AddressOf Compare))
        m_Blocks = tBlks
    End Sub
End Class

