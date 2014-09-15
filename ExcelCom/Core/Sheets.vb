Namespace Core

    Public Class Sheets : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public ReadOnly Property Count() As Long
            Get
                Return InvokeGetProperty(Of Long)("Count")
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