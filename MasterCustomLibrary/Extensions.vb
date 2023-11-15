Imports System.Runtime.CompilerServices
Imports System.IO

Public Module StringExtension
    <Extension()>
    Public Function RemoveInvalidChars(ByVal originalString As String) As String
        Dim finalString As String = String.Empty

        If Not String.IsNullOrEmpty(originalString) Then
            Return String.Concat(originalString.Split(Path.GetInvalidFileNameChars()))
        End If

        Return finalString
    End Function
End Module
