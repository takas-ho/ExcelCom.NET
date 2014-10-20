Public Class Charts

End Class
Namespace Core
    Public Class Charts : Inherits AbstractExcelSubCollection(Of Chart) : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Protected Overrides Function DetectIndex(ByVal item As Chart) As Integer
            Return item.Index
        End Function

        Protected Overrides Function DetectIndex(ByVal name As String) As Integer
            For i As Integer = 0 To Me.Count - 1
                If name.Equals(InternalItems(i).Name) Then
                    Return i
                End If
            Next
            Return -1
        End Function

        Public Function Add(Optional ByVal Before As Object = Nothing, Optional ByVal After As Object = Nothing, Optional ByVal Count As Object = Nothing) As Chart
            Dim args As New List(Of Object)
            If Before IsNot Nothing Then
                args.Add(New NamedParameter("Before", Before))
            End If
            If After IsNot Nothing Then
                args.Add(New NamedParameter("After", After))
            End If
            If Count IsNot Nothing Then
                args.Add(New NamedParameter("Count", Count))
            End If
            Dim comObject As Object = InvokeMethod("Add", args.ToArray)
            If comObject Is Nothing Then
                Return Nothing
            End If
            Return New Chart(Me, comObject)
        End Function

    End Class
End Namespace