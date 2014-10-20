Namespace Core
    ''' <summary>
    ''' Excelオブジェクトを表すインターフェース
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IExcelObject
        ''' <summary>
        ''' 生成したComObjectを管理に追加する
        ''' </summary>
        ''' <param name="comObject">ComObject</param>
        ''' <remarks></remarks>
        Sub AddToManager(ByVal comObject As Object)
    End Interface
End Namespace