Namespace Core
    Public Class Comment : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

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

    End Class
End Namespace