﻿Namespace Core
    Public Class Range : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Function Cells() As Range
            Return New Range(Me, InvokeGetProperty("Cells"))
        End Function

        Default Public ReadOnly Property Item(ByVal row As Integer, ByVal column As Integer) As Range
            Get
                Return New Range(Me, InvokeGetProperty("Item", RuleUtil.ConvIndexDotNET2VBA(row), RuleUtil.ConvIndexDotNET2VBA(column)))
            End Get
        End Property

        Public Function Columns() As Range
            Return New Range(Me, InvokeGetProperty("Columns"))
        End Function

        Public ReadOnly Property Count() As Long
            Get
                Return InvokeGetProperty(Of Long)("Count")
            End Get
        End Property

        Public Function Range(ByVal rangeStr As String) As Range
            Return InternalRange(rangeStr)
        End Function

        Public Function Range(ByVal startRange As String, ByVal endRange As String) As Range
            Return InternalRange(startRange, endRange)
        End Function

        Public Function Range(ByVal startRange As Range, ByVal endRange As Range) As Range
            Return InternalRange(startRange.ComObject, endRange.ComObject)
        End Function

        Private Function InternalRange(ByVal cell1 As Object, Optional ByVal cell2 As Object = Nothing) As Range
            Dim args As Object() = If(cell2 Is Nothing, New Object() {cell1}, New Object() {cell1, cell2})
            Return New Range(Me, InvokeGetProperty("Range", args))
        End Function

        Public Function Rows() As Range
            Return New Range(Me, InvokeGetProperty("Rows"))
        End Function

        Public Sub [Select]()
            InvokeMethod("Select")
        End Sub

        Public ReadOnly Property Text() As String
            Get
                Return InvokeGetProperty(Of String)("Text")
            End Get
        End Property

        Public Property Value() As Object
            Get
                Return InvokeGetProperty("Value")
            End Get
            Set(ByVal value As Object)
                InvokeSetProperty("Value", value)
            End Set
        End Property

    End Class
End Namespace