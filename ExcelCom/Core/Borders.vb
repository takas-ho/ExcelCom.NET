Namespace Core
    Public Class Borders : Inherits AbstractExcelSubObject : Implements IExcelObject, IExcelCollection(Of Border)

        Private ReadOnly _items As Dictionary(Of Border.XlBordersIndex, Border)

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)

            _items = New Dictionary(Of Border.XlBordersIndex, Border)
        End Sub

        Public ReadOnly Property Count() As Integer
            Get
                Return InvokeGetProperty(Of Integer)("Count")
            End Get
        End Property

        Default Public Overridable ReadOnly Property Item(ByVal index As Border.XlBordersIndex) As Border
            Get
                If Not _items.ContainsKey(index) Then
                    Dim comObject As Object = InvokeGetProperty("Item", index)
                    If comObject Is Nothing Then
                        Return Nothing
                    End If
                    _items.Add(index, New Border(Me, comObject))
                End If
                Return _items(index)
            End Get
        End Property

        Public ReadOnly Property InternalItems() As List(Of Border) Implements IExcelCollection(Of Border).InternalItems
            Get
                Return New List(Of Border)(_items.Values)
            End Get
        End Property
    End Class
End Namespace