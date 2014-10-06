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

        Public Property FreezePanes() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("FreezePanes")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("FreezePanes", value)
            End Set
        End Property

    End Class
End Namespace