Namespace Core

    Public Class Sheets : Inherits AbstractExcelSubCollection(Of Worksheet) : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Function Add(Optional ByVal before As Worksheet = Nothing, Optional ByVal after As Worksheet = Nothing, Optional ByVal count As Integer = -1, Optional ByVal type As Object = Nothing) As Worksheet
            Dim args As New List(Of Object)
            If before IsNot Nothing Then
                args.Add(New NamedParameter("Before", before.ComObject))
            End If
            If after IsNot Nothing Then
                args.Add(New NamedParameter("After", after.ComObject))
            End If
            If 0 < count Then
                args.Add(New NamedParameter("Count", count))
            End If
            If type IsNot Nothing Then
                args.Add(New NamedParameter("Type", type))
            End If
            Return New Worksheet(Me, InvokeMethod("Add", args.ToArray))
        End Function

    End Class
End Namespace