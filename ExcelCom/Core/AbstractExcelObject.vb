Namespace Core
    ''' <summary>
    ''' Excelオブジェクトの基底クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class AbstractExcelObject

        ''' <summary>COMオブジェクト本体</summary>
        Protected Friend ReadOnly ComObject As Object

        Private ReadOnly comType As Type

        Protected Sub New(ByVal comObject As Object)
            Me.ComObject = comObject
            Me.comType = comObject.GetType
        End Sub

        ''' <summary>
        ''' プロパティ値を取得する
        ''' </summary>
        ''' <param name="name">プロパティ名</param>
        ''' <param name="args">引数[]</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Protected Friend Function InvokeGetProperty(ByVal name As String, ByVal ParamArray args As Object()) As Object
            Return InvokeGetProperty(Of Object)(name, args)
        End Function

        ''' <summary>
        ''' プロパティ値を取得する
        ''' </summary>
        ''' <typeparam name="T">戻り値型</typeparam>
        ''' <param name="name">プロパティ名</param>
        ''' <param name="args">引数[]</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Protected Friend Function InvokeGetProperty(Of T)(ByVal name As String, ByVal ParamArray args As Object()) As T
            Dim value As Object = comType.InvokeMember(name, Reflection.BindingFlags.GetProperty Or Reflection.BindingFlags.Public, Nothing, ComObject, ResolveArgs(args))
            Return Cast(Of T)(value)
        End Function

        ''' <summary>
        ''' プロパティ値を設定する
        ''' </summary>
        ''' <param name="name">プロパティ名</param>
        ''' <param name="args">引数[]</param>
        ''' <remarks></remarks>
        Protected Friend Sub InvokeSetProperty(ByVal name As String, ByVal ParamArray args As Object())
            comType.InvokeMember(name, Reflection.BindingFlags.SetProperty Or Reflection.BindingFlags.Public, Nothing, ComObject, ResolveArgs(args))
        End Sub

        ''' <summary>
        ''' メソッドを実行する
        ''' </summary>
        ''' <param name="name">プロパティ名</param>
        ''' <param name="args">引数[]</param>
        ''' <returns>戻り値</returns>
        ''' <remarks></remarks>
        Protected Friend Function InvokeMethod(ByVal name As String, ByVal ParamArray args As Object()) As Object
            Return InvokeMethod(Of Object)(name, args)
        End Function

        ''' <summary>
        ''' メソッドを実行する
        ''' </summary>
        ''' <typeparam name="T">戻り値型</typeparam>
        ''' <param name="name">プロパティ名</param>
        ''' <param name="args">引数[]</param>
        ''' <returns>戻り値</returns>
        ''' <remarks></remarks>
        Protected Friend Function InvokeMethod(Of T)(ByVal name As String, ByVal ParamArray args As Object()) As T
            Dim value As Object = comType.InvokeMember(name, Reflection.BindingFlags.InvokeMethod Or Reflection.BindingFlags.Public, Nothing, ComObject, ResolveArgs(args))
            Return Cast(Of T)(value)
        End Function

        Private Function Cast(Of T)(ByVal value As Object) As T
            If GetType(T).IsEnum Then
                Return DirectCast(DirectCast(CInt(value), Object), T)
            End If
            Return DirectCast(value, T)
        End Function

        Private Function ResolveArgs(ByVal args As Object()) As Object()
            If args Is Nothing OrElse args.Length = 0 Then
                Return Nothing
            End If
            Return args
        End Function

    End Class
End Namespace