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

        Public Property Color() As Integer
            Get
                Return Convert.ToInt32(InvokeGetProperty("Color"))
            End Get
            Set(ByVal value As Integer)
                InvokeSetProperty("Color", value)
            End Set
        End Property

        Public Property ColorIndex() As Integer
            Get
                Return Convert.ToInt32(InvokeGetProperty("ColorIndex"))
            End Get
            Set(ByVal value As Integer)
                InvokeSetProperty("ColorIndex", value)
            End Set
        End Property

        Public Property LineStyle() As Border.XlLineStyle
            Get
                Return InvokeGetProperty(Of Border.XlLineStyle)("LineStyle")
            End Get
            Set(ByVal value As Border.XlLineStyle)
                InvokeSetProperty("LineStyle", value)
            End Set
        End Property

        Public Property Weight() As Border.XlBorderWeight
            Get
                Return InvokeGetProperty(Of Border.XlBorderWeight)("Weight")
            End Get
            Set(ByVal value As Border.XlBorderWeight)
                InvokeSetProperty("Weight", value)
            End Set
        End Property

    End Class
End Namespace