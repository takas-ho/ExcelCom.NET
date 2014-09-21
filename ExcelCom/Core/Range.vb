Namespace Core
    Public Class Range : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Sub Copy(Optional ByVal destination As Object = Nothing)
            Dim args As New List(Of Object)
            If destination IsNot Nothing Then
                args.Add(New NamedParameter("Destination", destination))
            End If
            InvokeMethod("Copy", args.ToArray)
        End Sub

        Public Function Cells() As Range
            Return New Range(Me, InvokeGetProperty("Cells"))
        End Function

        Default Public ReadOnly Property Item(ByVal index As Integer) As Range
            Get
                Return New Range(Me, InvokeGetProperty("Item", RuleUtil.ConvIndexDotNET2VBA(index)))
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal row As Integer, ByVal column As Integer) As Range
            Get
                Return New Range(Me, InvokeGetProperty("Item", RuleUtil.ConvIndexDotNET2VBA(row), RuleUtil.ConvIndexDotNET2VBA(column)))
            End Get
        End Property

        Public ReadOnly Property Column() As Integer
            Get
                Return RuleUtil.ConvIndexVBA2DotNET(InvokeGetProperty(Of Integer)("Column"))
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

        Public Function Delete(Optional ByVal shift As Object = Nothing) As Object
            Dim args As New List(Of Object)
            If shift IsNot Nothing Then
                args.Add(New NamedParameter("Shift", shift))
            End If
            Return InvokeMethod("Delete", args.ToArray)
        End Function

        Public Sub Insert(Optional ByVal shift As Object = Nothing)
            Dim args As New List(Of Object)
            If shift IsNot Nothing Then
                args.Add(New NamedParameter("Shift", shift))
            End If
            InvokeMethod("Insert", args.ToArray)
        End Sub

        Public Property NumberFormatLocal() As Object
            Get
                Return InvokeGetProperty("NumberFormatLocal")
            End Get
            Set(ByVal value As Object)
                InvokeSetProperty("NumberFormatLocal", value)
            End Set
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

        Public ReadOnly Property Row() As Integer
            Get
                Return RuleUtil.ConvIndexVBA2DotNET(InvokeGetProperty(Of Integer)("Row"))
            End Get
        End Property

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