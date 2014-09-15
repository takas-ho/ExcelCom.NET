Namespace Core

    Public Class Workbooks : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Function Add() As Workbook
            _items.Clear()
            Return New Workbook(Me, InvokeMethod("Add"))
        End Function

        Public ReadOnly Property Count() As Integer
            Get
                Return InvokeGetProperty(Of Integer)("Count")
            End Get
        End Property

        Private ReadOnly _items As New Dictionary(Of Integer, Workbook)
        Default Public ReadOnly Property Item(ByVal index As Integer) As Workbook
            Get
                If Not _items.ContainsKey(index) Then
                    _items.Add(index, New Workbook(Me, InvokeGetProperty("Item", RuleUtil.ConvIndexDotNET2VBA(index))))
                End If
                Return _items(index)
            End Get
        End Property

        Public Function Open(ByVal fileName As String, Optional ByVal updateLinks As Boolean = False) As Workbook
            _items.Clear()
            Return New Workbook(Me, InvokeMethod("Open", fileName, updateLinks))
        End Function

    End Class
End Namespace