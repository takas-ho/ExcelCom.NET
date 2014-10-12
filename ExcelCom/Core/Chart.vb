Namespace Core
    Public Class Chart : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

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

        Private _seriesCollection As SeriesCollection
        Public Function SeriesCollection() As SeriesCollection
            If _seriesCollection Is Nothing Then
                _seriesCollection = New SeriesCollection(Me, InvokeMethod("SeriesCollection"))
            End If
            Return _seriesCollection
        End Function

    End Class
End Namespace