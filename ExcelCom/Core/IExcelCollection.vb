Namespace Core
    ''' <summary>
    ''' Excelオブジェクトのコレクションを表すインターフェース
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <remarks></remarks>
    Public Interface IExcelCollection(Of T) : Inherits IExcelObject

        ''' <summary>要素[]</summary>
        ReadOnly Property InternalItems() As List(Of T)
    End Interface
End Namespace