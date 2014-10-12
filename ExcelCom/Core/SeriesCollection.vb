Namespace Core

    Public Class SeriesCollection : Inherits AbstractExcelSubCollection(Of Series) : Implements IExcelObject

        Public Enum XlRowCol
            xlColumns = 2
            xlRows = 1
        End Enum

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Protected Overrides Function DetectIndex(ByVal item As Series) As Integer
            For i As Integer = 0 To Me.Count - 1
                If item.Name.Equals(InternalItems(i).Name) Then
                    Return i
                End If
            Next
            Return -1
        End Function

        Protected Overrides Function DetectIndex(ByVal name As String) As Integer
            For i As Integer = 0 To Me.Count - 1
                If name.Equals(InternalItems(i).Name) Then
                    Return i
                End If
            Next
            Return -1
        End Function

        Public Function Add(ByVal Source As Object, Optional ByVal Rowcol As XlRowCol = XlRowCol.xlColumns, _
                            Optional ByVal SeriesLabels As Object = Nothing, Optional ByVal CategoryLabels As Object = Nothing, Optional ByVal Replace As Object = Nothing) As Series
            Dim args As New List(Of Object)
            args.Add(Source)
            args.Add(New NamedParameter("Rowcol", Rowcol))
            If SeriesLabels IsNot Nothing Then
                args.Add(New NamedParameter("SeriesLabels", SeriesLabels))
            End If
            If CategoryLabels IsNot Nothing Then
                args.Add(New NamedParameter("CategoryLabels", CategoryLabels))
            End If
            If Replace IsNot Nothing Then
                args.Add(New NamedParameter("Replace", Replace))
            End If
            Dim comObject As Object = InvokeMethod("Add", args.ToArray)
            If comObject Is Nothing Then
                Return Nothing
            End If
            Dim result As Series = New Series(Me, comObject)
            InternalItems.Add(result)
            Return result
        End Function

        Public Function NewSeries() As Series
            Dim comObject As Object = InvokeMethod("NewSeries")
            If comObject Is Nothing Then
                Return Nothing
            End If
            Dim result As Series = New Series(Me, comObject)
            InternalItems.Add(result)
            Return result
        End Function

    End Class
End Namespace