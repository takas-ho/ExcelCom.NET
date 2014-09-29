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

        'Public Property ColorIndex() As Integer
        '    Get
        '        Return Convert.ToInt32(InvokeGetProperty("ColorIndex"))
        '    End Get
        '    Set(ByVal value As Integer)
        '        InvokeSetProperty("ColorIndex", value)
        '    End Set
        'End Property

        'Public Property Italic() As Boolean
        '    Get
        '        Return InvokeGetProperty(Of Boolean)("Italic")
        '    End Get
        '    Set(ByVal value As Boolean)
        '        InvokeSetProperty("Italic", value)
        '    End Set
        'End Property

        'Public Property Name() As String
        '    Get
        '        Return InvokeGetProperty(Of String)("Name")
        '    End Get
        '    Set(ByVal value As String)
        '        InvokeSetProperty("Name", value)
        '    End Set
        'End Property

        'Public Property Shadow() As Boolean
        '    Get
        '        Return InvokeGetProperty(Of Boolean)("Shadow")
        '    End Get
        '    Set(ByVal value As Boolean)
        '        InvokeSetProperty("Shadow", value)
        '    End Set
        'End Property

        'Public Property Size() As Double
        '    Get
        '        Return InvokeGetProperty(Of Double)("Size")
        '    End Get
        '    Set(ByVal value As Double)
        '        InvokeSetProperty("Size", value)
        '    End Set
        'End Property

        '''' <summary>取消線 on/off</summary>
        'Public Property Strikethrough() As Boolean
        '    Get
        '        Return InvokeGetProperty(Of Boolean)("Strikethrough")
        '    End Get
        '    Set(ByVal value As Boolean)
        '        InvokeSetProperty("Strikethrough", value)
        '    End Set
        'End Property

        '''' <summary>下付き文字 on/off</summary>
        'Public Property Subscript() As Boolean
        '    Get
        '        Return InvokeGetProperty(Of Boolean)("Subscript")
        '    End Get
        '    Set(ByVal value As Boolean)
        '        InvokeSetProperty("Subscript", value)
        '    End Set
        'End Property

        '''' <summary>上付き文字 on/off</summary>
        'Public Property Superscript() As Boolean
        '    Get
        '        Return InvokeGetProperty(Of Boolean)("Superscript")
        '    End Get
        '    Set(ByVal value As Boolean)
        '        InvokeSetProperty("Superscript", value)
        '    End Set
        'End Property

        'Public Property Underline() As XlUnderlineStyle
        '    Get
        '        Return InvokeGetProperty(Of XlUnderlineStyle)("Underline")
        '    End Get
        '    Set(ByVal value As XlUnderlineStyle)
        '        InvokeSetProperty("Underline", value)
        '    End Set
        'End Property

    End Class
End Namespace