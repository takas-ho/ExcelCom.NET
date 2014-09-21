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

            Dim result As Worksheet = New Worksheet(Me, InvokeMethod("Add", args.ToArray))

            If before IsNot Nothing Then
                InternalItems.Insert(before.Index, result)
            ElseIf after IsNot Nothing Then
                If after.Index < Me.Count - 1 Then
                    InternalItems.Insert(after.Index + 1, result)
                Else
                    InternalItems.Add(result)
                End If
            Else
                InternalItems.Insert(0, result)
            End If
            Return result
        End Function

        Protected Overrides Function DetectIndex(ByVal item As Worksheet) As Integer
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

    End Class
End Namespace