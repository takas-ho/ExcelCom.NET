Namespace Core
    Public Class Borders : Inherits AbstractExcelSubObject : Implements IExcelObject, IExcelCollection(Of Border)

        Private ReadOnly _items As Dictionary(Of XlBordersIndex, Border)

#Region "Xl定数"
        ''' <summary>罫線の場所</summary>
        Public Enum XlBordersIndex
            ''' <summary>セル範囲の各セルの左上隅から右下隅への罫線</summary>
            xlDiagonalDown = 5
            ''' <summary>セル範囲の各セルの左下隅から右上隅への罫線</summary>
            xlDiagonalUp = 6
            ''' <summary>セル範囲の左側の罫線</summary>
            xlEdgeLeft = 7
            ''' <summary>セル範囲の上側の罫線</summary>
            xlEdgeTop = 8
            ''' <summary>セル範囲の下側の罫線</summary>
            xlEdgeBottom = 9
            ''' <summary>セル範囲の右側の罫線</summary>
            xlEdgeRight = 10
            ''' <summary>セル範囲の外枠を除く、すべてのセルの垂直方向の罫線</summary>
            xlInsideVertical = 11
            ''' <summary>セル範囲の外枠を除く、すべてのセルの水平方向の罫線</summary>
            xlInsideHorizontal = 12
        End Enum
#End Region

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)

            _items = New Dictionary(Of XlBordersIndex, Border)
        End Sub

        Public ReadOnly Property Count() As Integer
            Get
                Return InvokeGetProperty(Of Integer)("Count")
            End Get
        End Property

        Default Public Overridable ReadOnly Property Item(ByVal index As XlBordersIndex) As Border
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