Namespace Core
    Public Class Comment : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public ReadOnly Property Author() As String
            Get
                Return InvokeGetProperty(Of String)("Author")
            End Get
        End Property

        Public Function Text(Optional ByVal aText As Object = Nothing, Optional ByVal start As Object = Nothing, Optional ByVal overwrite As Object = Nothing) As String
            Dim args As New List(Of Object)
            If aText IsNot Nothing Then
                args.Add(New NamedParameter("Text", aText))
            End If
            If start IsNot Nothing Then
                args.Add(New NamedParameter("Start", start))
            End If
            If overwrite IsNot Nothing Then
                args.Add(New NamedParameter("Overwrite", overwrite))
            End If
            Return InvokeMethod(Of String)("Text", args.ToArray)
        End Function

        Private _shape As Shape
        Public ReadOnly Property Shape() As Shape
            Get
                If _shape Is Nothing Then
                    _shape = New Shape(Me, InvokeGetProperty("Shape"))
                End If
                Return _shape
            End Get
        End Property

    End Class
End Namespace