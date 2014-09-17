Imports System.Reflection

Namespace Core

    Public Class AbstractExcelSubCollection(Of T As AbstractExcelSubObject) : Inherits AbstractExcelSubObject : Implements IExcelObject

        Private ReadOnly _items As List(Of T)

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)

            _items = New List(Of T)
            For i As Integer = 0 To Count - 1
                _items.Add(Nothing)
            Next
        End Sub

        Friend ReadOnly Property InternalItems() As List(Of T)
            Get
                Return _items
            End Get
        End Property

        Public ReadOnly Property Count() As Integer
            Get
                Return InvokeGetProperty(Of Integer)("Count")
            End Get
        End Property

        Public Function IndexOf(ByVal item As T) As Integer
            Return _items.IndexOf(item)
        End Function

        Default Public ReadOnly Property Item(ByVal index As Integer) As T
            Get
                If _items(index) Is Nothing Then
                    Dim constructorInfo As ConstructorInfo = GetType(T).GetConstructor(New System.Type() {GetType(IExcelObject), GetType(Object)})
                    Dim comObject As Object = InvokeGetProperty("Item", RuleUtil.ConvIndexDotNET2VBA(index))
                    If comObject Is Nothing Then
                        Return Nothing
                    End If
                    _items(index) = DirectCast(constructorInfo.Invoke(New Object() {Me, comObject}), T)
                End If
                Return _items(index)
            End Get
        End Property

    End Class
End Namespace