Namespace Core

    Public Class Worksheet : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Enum XlSheetVisibility
            xlSheetHidden = 0
            xlSheetVeryHidden = 2
            xlSheetVisible = -1
        End Enum

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Function Cells() As Range
            Return New Range(Me, InvokeGetProperty("Cells"))
        End Function

        Public Function Columns() As Range
            Return New Range(Me, InvokeGetProperty("Columns"))
        End Function

        Public ReadOnly Property Index() As Integer
            Get
                Return RuleUtil.ConvIndexVBA2DotNET(InvokeGetProperty(Of Integer)("Index"))
            End Get
        End Property

        Public Property Name() As String
            Get
                Return InvokeGetProperty(Of String)("Name")
            End Get
            Set(ByVal value As String)
                InvokeSetProperty("Name", value)
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

        Public Function Rows() As Range
            Return New Range(Me, InvokeGetProperty("Rows"))
        End Function

        Public Sub [Select]()
            InvokeMethod("Select")
        End Sub

        Public Property Visible() As XlSheetVisibility
            Get
                Return InvokeGetProperty(Of XlSheetVisibility)("Visible")
            End Get
            Set(ByVal value As XlSheetVisibility)
                InvokeSetProperty("Visible", value)
            End Set
        End Property

    End Class
End Namespace