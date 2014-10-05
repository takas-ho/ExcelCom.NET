Namespace Core

    Public Class Window : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public ReadOnly Property Index() As Integer
            Get
                Return RuleUtil.ConvIndexVBA2DotNET(InvokeGetProperty(Of Integer)("Index"))
            End Get
        End Property

    End Class
End Namespace