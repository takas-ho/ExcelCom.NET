Namespace Core

    Public Class Worksheet : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Enum XlSheetVisibility
            xlSheetHidden = 0
            xlSheetVeryHidden = 2
            xlSheetVisible = -1
        End Enum

        Private Shadows ReadOnly parent As IExcelCollection(Of Worksheet)

        Public Sub New(ByVal parent As IExcelCollection(Of Worksheet), ByVal comObject As Object)
            MyBase.New(parent, comObject)
            Me.parent = parent
        End Sub

        Public Sub Calculate()
            InvokeMethod("Calculate")
        End Sub

        Public Function Cells() As Range
            Return New Range(Me, InvokeGetProperty("Cells"))
        End Function

        Private _chartObjects As ChartObjects
        Public Function ChartObjects() As ChartObjects
            If _chartObjects Is Nothing Then
                _chartObjects = New ChartObjects(Me, InvokeGetProperty("ChartObjects"))
            End If
            Return _chartObjects
        End Function

        Public Function Columns() As Range
            Return New Range(Me, InvokeGetProperty("Columns"))
        End Function

        Public Sub Copy(Optional ByVal before As Worksheet = Nothing, Optional ByVal after As Worksheet = Nothing)
            Dim args As New List(Of Object)
            If before IsNot Nothing Then
                args.Add(New NamedParameter("Before", before.ComObject))
            End If
            If after IsNot Nothing Then
                args.Add(New NamedParameter("After", after.ComObject))
            End If
            InvokeMethod("Copy", args.ToArray)

            If before IsNot Nothing Then
                Dim index As Integer = parent.InternalItems.IndexOf(before)
                If 0 <= index Then
                    parent.InternalItems.Insert(index, Nothing)
                End If
            End If
            If after IsNot Nothing Then
                Dim index As Integer = parent.InternalItems.IndexOf(after)
                If 0 <= index AndAlso index < parent.InternalItems.Count - 1 Then
                    parent.InternalItems.Insert(index + 1, Nothing)
                End If
            End If
        End Sub

        Public ReadOnly Property Index() As Integer
            Get
                Return RuleUtil.ConvIndexVBA2DotNET(InvokeGetProperty(Of Integer)("Index"))
            End Get
        End Property

        Public Property Name() As String
            Get
                Return InvokeGetProperty(Of String)("Name")
            End Get
            Set(ByVal value As String)
                InvokeSetProperty("Name", value)
            End Set
        End Property

        Public Sub PrintOut(Optional ByVal printerName As String = Nothing, Optional ByVal preview As Boolean = False, _
                            Optional ByVal copyCount As Integer = 1, Optional ByVal isCollate As Boolean = True)
            Dim args As New List(Of Object)
            If Not String.IsNullOrEmpty(printerName) Then
                args.Add(New NamedParameter("ActivePrinter", printerName))
            End If
            args.Add(New NamedParameter("Preview", preview))
            args.Add(New NamedParameter("Copies", copyCount))
            args.Add(New NamedParameter("Collate", isCollate))
            InvokeMethod("PrintOut", args.ToArray)
        End Sub

        ''' <summary>
        ''' 保護する
        ''' </summary>
        ''' <param name="password">パスワード</param>
        ''' <param name="drawingObjects">描画オブジェクトを保護する場合、true</param>
        ''' <param name="contents">シートの内容を保護する場合、true</param>
        ''' <param name="scenarios">シナリオを保護する場合、true</param>
        ''' <param name="userInterfaceOnly">マクロからの変更は可能にする場合、true</param>
        ''' <remarks></remarks>
        Public Sub Protect(Optional ByVal password As String = Nothing, Optional ByVal drawingObjects As Boolean = False, _
                           Optional ByVal contents As Boolean = True, Optional ByVal scenarios As Boolean = True, _
                           Optional ByVal userInterfaceOnly As Boolean = False)
            Dim args As New List(Of Object)
            If Not String.IsNullOrEmpty(password) Then
                args.Add(New NamedParameter("Password", password))
            End If
            args.Add(New NamedParameter("DrawingObjects", drawingObjects))
            args.Add(New NamedParameter("Contents", contents))
            args.Add(New NamedParameter("Scenarios", scenarios))
            args.Add(New NamedParameter("UserInterfaceOnly", userInterfaceOnly))
            InvokeMethod("Protect", args.ToArray)
        End Sub

        Public Function Range(ByVal rangeStr As String) As Range
            Return InternalRange(rangeStr)
        End Function

        Public Function Range(ByVal startRange As String, ByVal endRange As String) As Range
            Return InternalRange(startRange, endRange)
        End Function

        Public Function Range(ByVal startRange As Range, ByVal endRange As Range) As Range
            Return InternalRange(startRange.ComObject, endRange.ComObject)
        End Function

        Private Function InternalRange(ByVal cell1 As Object, Optional ByVal cell2 As Object = Nothing) As Range
            Dim args As Object() = If(cell2 Is Nothing, New Object() {cell1}, New Object() {cell1, cell2})
            Return New Range(Me, InvokeGetProperty("Range", args))
        End Function

        Public Function Rows() As Range
            Return New Range(Me, InvokeGetProperty("Rows"))
        End Function

        Public Sub [Select]()
            InvokeMethod("Select")
        End Sub

        Private _shapes As Shapes
        Public Function Shapes() As Shapes
            If _shapes Is Nothing Then
                _shapes = New Shapes(Me, InvokeGetProperty("Shapes"))
            End If
            Return _shapes
        End Function

        ''' <summary>
        ''' 保護解除する
        ''' </summary>
        ''' <param name="password">パスワード</param>
        ''' <remarks></remarks>
        Public Sub Unprotect(Optional ByVal password As String = Nothing)
            Dim args As New List(Of Object)
            If Not String.IsNullOrEmpty(password) Then
                args.Add(New NamedParameter("Password", password))
            End If
            InvokeMethod("Unprotect", args.ToArray)
        End Sub

        Public Property Visible() As XlSheetVisibility
            Get
                Return InvokeGetProperty(Of XlSheetVisibility)("Visible")
            End Get
            Set(ByVal value As XlSheetVisibility)
                InvokeSetProperty("Visible", value)
            End Set
        End Property

    End Class
End Namespace