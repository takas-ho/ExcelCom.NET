Namespace Core

    Public Class Sheets : Inherits AbstractExcelSubObject : Implements IExcelObject

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

        Public ReadOnly Property Count() As Integer
            Get
                Return InvokeGetProperty(Of Integer)("Count")
            End Get
        End Property

        Private ReadOnly _items As New Dictionary(Of Integer, Worksheet)
        Default Public ReadOnly Property Item(ByVal index As Integer) As Worksheet
            Get
                If Not _items.ContainsKey(index) Then
                    _items.Add(index, New Worksheet(Me, InvokeGetProperty("Item", RuleUtil.ConvIndexDotNET2VBA(index))))
                End If
                Return _items(index)
            End Get
        End Property

    End Class
End Namespace