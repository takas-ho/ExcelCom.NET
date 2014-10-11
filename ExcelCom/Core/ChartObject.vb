Namespace Core
    Public Class ChartObject : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Private _chart As Chart
        Public ReadOnly Property Chart() As Chart
            Get
                If _chart Is Nothing Then
                    _chart = New Chart(Me, InvokeGetProperty("Chart"))
                End If
                Return _chart
            End Get
        End Property

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

    End Class
End Namespace
