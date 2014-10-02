Namespace Core
    Public Class Border : Inherits AbstractExcelSubObject : Implements IExcelObject

#Region "Xl定数"
        ''' <summary>罫線の場所</summary>
        Public Enum XlBordersIndex
            ''' <summary>セル範囲の各セルの左上隅から右下隅への罫線</summary>
            xlDiagonalDown = 5
            ''' <summary>セル範囲の各セルの左下隅から右上隅への罫線</summary>
            xlDiagonalUp = 6
            ''' <summary>セル範囲の左側の罫線</summary>
            xlEdgeLeft = 7
            ''' <summary>セル範囲の上側の罫線</summary>
            xlEdgeTop = 8
            ''' <summary>セル範囲の下側の罫線</summary>
            xlEdgeBottom = 9
            ''' <summary>セル範囲の右側の罫線</summary>
            xlEdgeRight = 10
            ''' <summary>セル範囲の外枠を除く、すべてのセルの垂直方向の罫線</summary>
            xlInsideVertical = 11
            ''' <summary>セル範囲の外枠を除く、すべてのセルの水平方向の罫線</summary>
            xlInsideHorizontal = 12
        End Enum

        ''' <summary>罫線の線種</summary>
        Public Enum XlLineStyle
            ''' <summary>実線</summary>
            xlContinuous = 1
            ''' <summary>一点鎖線</summary>
            xlDashDot = 4
            ''' <summary>二点鎖線</summary>
            xlDashDotDot = 5
            ''' <summary>斜線</summary>
            xlSlantDashDot = 13
            ''' <summary>破線</summary>
            xlDash = -4115
            ''' <summary>点線</summary>
            xlDot = -4118
            ''' <summary>二重線</summary>
            xlDouble = -4119
            ''' <summary>線なし</summary>
            xlLineStyleNone = -4142
        End Enum
#End Region

        Public Sub New(ByVal parent As IExcelCollection(Of Border), ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        'Public Property ColorIndex() As Integer
        '    Get
        '        Return Convert.ToInt32(InvokeGetProperty("ColorIndex"))
        '    End Get
        '    Set(ByVal value As Integer)
        '        InvokeSetProperty("ColorIndex", value)
        '    End Set
        'End Property

        Public Property LineStyle() As XlLineStyle
            Get
                Return InvokeGetProperty(Of XlLineStyle)("LineStyle")
            End Get
            Set(ByVal value As XlLineStyle)
                InvokeSetProperty("LineStyle", value)
            End Set
        End Property

    End Class
End Namespace