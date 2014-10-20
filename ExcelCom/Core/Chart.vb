Namespace Core
    Public Class Chart : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Enum XlChartLocation
            xlLocationAsNewSheet = 1
            xlLocationAsObject = 2
            xlLocationAutomatic = 3
        End Enum

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Property HasLegend() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("HasLegend")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("HasLegend", value)
            End Set
        End Property

        Public ReadOnly Property Index() As Integer
            Get
                Return RuleUtil.ConvIndexVBA2DotNET(InvokeGetProperty(Of Integer)("Index"))
            End Get
        End Property

        'Public Function Location(ByVal where As XlChartLocation, Optional ByVal name As String = Nothing) As Chart
        '    Dim args As New List(Of Object)
        '    args.Add(New NamedParameter("Where", where))
        '    If Name IsNot Nothing Then
        '        args.Add(New NamedParameter("Name", name))
        '    End If
        '    Dim comObject As Object = InvokeMethod("Location", args.ToArray)
        '    If comObject Is Nothing Then
        '        Return Nothing
        '    End If
        '    Return New Chart(Me, comObject)
        'End Function

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