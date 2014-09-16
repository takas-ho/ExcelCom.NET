Namespace Core
    ''' <summary>
    ''' 名前付き引数を表すクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class NamedParameter
        ''' <summary>引数名</summary>
        Public ReadOnly Name As String
        ''' <summary>引数値</summary>
        Public ReadOnly Value As Object

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">引数名</param>
        ''' <param name="value">引数値</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal value As Object)
            Me.Name = name
            Me.Value = value
        End Sub
    End Class
End Namespace