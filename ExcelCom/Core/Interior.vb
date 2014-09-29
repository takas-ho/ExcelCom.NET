Namespace Core
    Public Class Interior : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        'Public Property Bold() As Boolean
        '    Get
        '        Return InvokeGetProperty(Of Boolean)("Bold")
        '    End Get
        '    Set(ByVal value As Boolean)
        '        InvokeSetProperty("Bold", value)
        '    End Set
        'End Property

        'Public Property Color() As Integer
        '    Get
        '        Return Convert.ToInt32(InvokeGetProperty("Color"))
        '    End Get
        '    Set(ByVal value As Integer)
        '        InvokeSetProperty("Color", value)
        '    End Set
        'End Property

        Public Property ColorIndex() As Integer
            Get
                Return Convert.ToInt32(InvokeGetProperty("ColorIndex"))
            End Get
            Set(ByVal value As Integer)
                InvokeSetProperty("ColorIndex", value)
            End Set
        End Property

    End Class
End Namespace