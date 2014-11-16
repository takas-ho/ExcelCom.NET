Namespace Core
    Public Class Range : Inherits AbstractExcelSubObject : Implements IExcelObject

#Region "Xl定数"

        Public Enum XlSearchDirection
            xlNext = 1
            xlPrevious = 2
        End Enum
        ''' <summary>検索方法</summary>
        Public Enum XlLookAt
            ''' <summary>全てが一致するセルを検索</summary>
            xlWhole = 1
            ''' <summary>一部が一致するセルを検索</summary>
            xlPart = 2
        End Enum
        ''' <summary>検索方向</summary>
        Public Enum XlSearchOrder
            ''' <summary>行</summary>
            xlByRows = 1
            ''' <summary>列</summary>
            xlByColumns = 2
        End Enum
        ''' <summary>横配置</summary>
        Public Enum XlHAlign
            ''' <summary>標準</summary>
            xlHAlignGeneral = 1
            ''' <summary>繰り返し</summary>
            xlHAlignFill = 5
            ''' <summary>選択範囲内</summary>
            xlHAlignCenterAcrossSelection = 7
            ''' <summary>中央揃え</summary>
            xlHAlignCenter = -4108
            ''' <summary>均等割り付け</summary>
            xlHAlignDistributed = -4117
            ''' <summary>両端揃え</summary>
            xlHAlignJustify = -4130
            ''' <summary>左詰め</summary>
            xlHAlignLeft = -4131
            ''' <summary>右詰め</summary>
            xlHAlignRight = -4152
        End Enum

        ''' <summary>縦配置</summary>
        Public Enum XlVAlign
            ''' <summary>下詰め</summary>
            xlVAlignBottom = -4107
            ''' <summary>中央揃え</summary>
            xlVAlignCenter = -4108
            ''' <summary>均等割り付け</summary>
            xlVAlignDistributed = -4117
            ''' <summary>両端揃え</summary>
            xlVAlignJustify = -4130
            ''' <summary>上詰め</summary>
            xlVAlignTop = -4160
        End Enum

        Public Enum XlAutoFilterOperator
            xlAnd = 1
            xlOr = 2
            xlBottom10Items = 4
            xlBottom10Percent = 6
            xlTop10Items = 3
            xlTop10Percent = 5
        End Enum

        ''' <summary>セルタイプ</summary>
        Public Enum XlCellType
            xlCellTypeAllFormatConditions = -4172
            xlCellTypeAllValidation = -4174
            xlCellTypeBlanks = 4
            xlCellTypeComments = -4144
            xlCellTypeConstants = 2
            xlCellTypeFormulas = -4123
            ''' <summary>データ入力最終セル</summary>
            xlCellTypeLastCell = 11
            xlCellTypeSameFormatConditions = -4173
            xlCellTypeSameValidation = -4175
            xlCellTypeVisible = 12
        End Enum

        ''' <summary>
        ''' 文字列の向きを指定します
        ''' </summary>
        Public Enum XlOrientation
            ''' <summary>下向き</summary>
            xlDownward = -4170
            ''' <summary>水平</summary>
            xlHorizontal = -4128
            ''' <summary>上向き</summary>
            xlUpward = -4171
            ''' <summary>垂直 (縦書き)</summary>
            xlVertical = -4166
        End Enum

        Public Enum XlBorderWeight
            xlHairline = 1
            xlMedium = -4138
            xlThick = 4
            xlThin = 2
        End Enum

        Public Enum XlColorIndex
            xlColorIndexAutomatic = -4105
            xlColorIndexNone = -4142
        End Enum
#End Region

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Function AddComment(Optional ByVal text As String = Nothing) As Comment
            Dim args As New List(Of Object)
            If text IsNot Nothing Then
                args.Add(New NamedParameter("Text", text))
            End If
            Dim comObject As Object = InvokeMethod("AddComment", args.ToArray)
            If comObject Is Nothing Then
                Return Nothing
            End If
            Return New Comment(Me, comObject)
        End Function

        Public Function AutoFit() As Object
            Return InvokeMethod("AutoFit")
        End Function

        Public Function AutoFilter(Optional ByVal columnIndex As Integer = -1, Optional ByVal criteria1 As String = Nothing, _
                                   Optional ByVal [operator] As XlAutoFilterOperator = XlAutoFilterOperator.xlAnd, _
                                   Optional ByVal criteria2 As String = Nothing, Optional ByVal visibleDropDown As Boolean = True) As Object
            Dim args As New List(Of Object)
            If 0 <= columnIndex Then
                args.Add(New NamedParameter("Field", RuleUtil.ConvIndexDotNET2VBA(columnIndex)))
            End If
            If Not String.IsNullOrEmpty(criteria1) Then
                args.Add(New NamedParameter("Criteria1", criteria1))
            End If
            args.Add(New NamedParameter("Operator", [operator]))
            If Not String.IsNullOrEmpty(criteria2) Then
                args.Add(New NamedParameter("Criteria2", criteria2))
            End If
            args.Add(New NamedParameter("VisibleDropDown", visibleDropDown))
            Return InvokeMethod("AutoFilter", args.ToArray)
        End Function

        Public Function Borders() As Borders
            Return New Borders(Me, InvokeGetProperty("Borders"))
        End Function

        Public Function BorderAround(Optional ByVal LineStyle As Object = Nothing, Optional ByVal Weight As XlBorderWeight = XlBorderWeight.xlThin, _
                                     Optional ByVal ColorIndex As XlColorIndex = XlColorIndex.xlColorIndexAutomatic, Optional ByVal Color As Object = Nothing) As Boolean
            Dim args As New List(Of Object)
            If LineStyle IsNot Nothing Then
                args.Add(New NamedParameter("LineStyle", LineStyle))
            End If
            args.Add(New NamedParameter("Weight", Weight))
            args.Add(New NamedParameter("ColorIndex", ColorIndex))
            If Color IsNot Nothing Then
                args.Add(New NamedParameter("Color", Color))
            End If
            Return InvokeMethod(Of Boolean)("BorderAround", args.ToArray)
        End Function

        Public Function Cells() As Range
            Return New Range(Me, InvokeGetProperty("Cells"))
        End Function

        Public Function Clear() As Boolean
            Return InvokeMethod(Of Boolean)("Clear")
        End Function

        Public Sub ClearComments()
            InvokeMethod("ClearComments")
        End Sub

        Public Function ClearContents() As Boolean
            Return InvokeMethod(Of Boolean)("ClearContents")
        End Function

        Public Function ClearFormats() As Boolean
            Return InvokeMethod(Of Boolean)("ClearFormats")
        End Function

        Public ReadOnly Property Column() As Integer
            Get
                Return RuleUtil.ConvIndexVBA2DotNET(InvokeGetProperty(Of Integer)("Column"))
            End Get
        End Property

        Public Function Columns() As Range
            Return New Range(Me, InvokeGetProperty("Columns"))
        End Function

        Public Property ColumnWidth() As Double
            Get
                Return InvokeGetProperty(Of Double)("ColumnWidth")
            End Get
            Set(ByVal value As Double)
                InvokeSetProperty("ColumnWidth", value)
            End Set
        End Property

        Public Function Comment() As Comment
            Dim comObject = InvokeGetProperty("Comment")
            If comObject Is Nothing Then
                Return Nothing
            End If
            Return New Comment(Me, comObject)
        End Function

        Public Sub Copy(Optional ByVal destination As Object = Nothing)
            Dim args As New List(Of Object)
            If destination IsNot Nothing Then
                args.Add(New NamedParameter("Destination", destination))
            End If
            InvokeMethod("Copy", args.ToArray)
        End Sub

        Public ReadOnly Property Count() As Long
            Get
                Return InvokeGetProperty(Of Integer)("Count")
            End Get
        End Property

        Public Function Delete(Optional ByVal shift As Object = Nothing) As Object
            Dim args As New List(Of Object)
            If shift IsNot Nothing Then
                args.Add(New NamedParameter("Shift", shift))
            End If
            Return InvokeMethod("Delete", args.ToArray)
        End Function

        Public Function EntireColumn() As Range
            Dim comObject As Object = InvokeGetProperty("EntireColumn")
            If comObject Is Nothing Then
                Return Nothing
            End If
            Return New Range(Me, comObject)
        End Function

        Public Function EntireRow() As Range
            Dim comObject As Object = InvokeGetProperty("EntireRow")
            If comObject Is Nothing Then
                Return Nothing
            End If
            Return New Range(Me, comObject)
        End Function

        Public Function Find(ByVal What As Object, Optional ByVal After As Range = Nothing, Optional ByVal LookIn As Object = Nothing, _
                             Optional ByVal LookAt As XlLookAt = 0, Optional ByVal SearchOrder As XlSearchOrder = 0, _
                             Optional ByVal SearchDirection As XlSearchDirection = 0, _
                             Optional ByVal MatchCase As Boolean = False, Optional ByVal MatchByte As Boolean = False) As Range
            Dim args As New List(Of Object)
            args.Add(What)
            If After IsNot Nothing Then
                args.Add(New NamedParameter("After", After.ComObject))
            End If
            If LookIn IsNot Nothing Then
                args.Add(New NamedParameter("LookIn", LookIn))
            End If
            If LookAt <> 0 Then
                args.Add(New NamedParameter("LookAt", LookAt))
            End If
            If SearchOrder <> 0 Then
                args.Add(New NamedParameter("SearchOrder", SearchOrder))
            End If
            If SearchDirection <> 0 Then
                args.Add(New NamedParameter("SearchDirection", SearchDirection))
            End If
            args.Add(New NamedParameter("MatchCase", MatchCase))
            args.Add(New NamedParameter("MatchByte", MatchByte))
            Dim result As Object = InvokeMethod("Find", args.ToArray)
            If result Is Nothing Then
                Return Nothing
            End If
            Return New Range(Me, result)
        End Function

        Public Function Font() As Font
            Return New Font(Me, InvokeGetProperty("Font"))
        End Function

        Public Property Formula() As String
            Get
                Return InvokeGetProperty(Of String)("Formula")
            End Get
            Set(ByVal value As String)
                InvokeSetProperty("Formula", value)
            End Set
        End Property

        Public Property FormulaLocal() As String
            Get
                Return InvokeGetProperty(Of String)("FormulaLocal")
            End Get
            Set(ByVal value As String)
                InvokeSetProperty("FormulaLocal", value)
            End Set
        End Property

        Public Property FormulaR1C1() As String
            Get
                Return InvokeGetProperty(Of String)("FormulaR1C1")
            End Get
            Set(ByVal value As String)
                InvokeSetProperty("FormulaR1C1", value)
            End Set
        End Property

        Public Property FormulaR1C1Local() As String
            Get
                Return InvokeGetProperty(Of String)("FormulaR1C1Local")
            End Get
            Set(ByVal value As String)
                InvokeSetProperty("FormulaR1C1Local", value)
            End Set
        End Property

        Public ReadOnly Property HasFormula() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("HasFormula")
            End Get
        End Property

        Public Property Hidden() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("Hidden")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("Hidden", value)
            End Set
        End Property

        Public Property HorizontalAlignment() As XlHAlign
            Get
                Return InvokeGetProperty(Of XlHAlign)("HorizontalAlignment")
            End Get
            Set(ByVal value As XlHAlign)
                InvokeSetProperty("HorizontalAlignment", value)
            End Set
        End Property

        Public Sub Insert(Optional ByVal shift As Object = Nothing)
            Dim args As New List(Of Object)
            If shift IsNot Nothing Then
                args.Add(New NamedParameter("Shift", shift))
            End If
            InvokeMethod("Insert", args.ToArray)
        End Sub

        Public Function Interior() As Interior
            Return New Interior(Me, InvokeGetProperty("Interior"))
        End Function

        Default Public ReadOnly Property Item(ByVal index As Integer) As Range
            Get
                Return New Range(Me, InvokeGetProperty("Item", RuleUtil.ConvIndexDotNET2VBA(index)))
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal row As Integer, ByVal column As Integer) As Range
            Get
                Return New Range(Me, InvokeGetProperty("Item", RuleUtil.ConvIndexDotNET2VBA(row), RuleUtil.ConvIndexDotNET2VBA(column)))
            End Get
        End Property

        Public Sub Merge(Optional ByVal across As Boolean = False)
            Dim args As New List(Of Object)
            args.Add(New NamedParameter("Across", across))
            InvokeMethod("Merge", args.ToArray)
        End Sub

        Public Function MergeArea() As Range
            Dim comObject As Object = InvokeGetProperty("MergeArea")
            If comObject Is Nothing Then
                Return Nothing
            End If
            Return New Range(Me, comObject)
        End Function

        Public Property MergeCells() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("MergeCells")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("MergeCells", value)
            End Set
        End Property

        Public Property NumberFormatLocal() As Object
            Get
                Return InvokeGetProperty("NumberFormatLocal")
            End Get
            Set(ByVal value As Object)
                InvokeSetProperty("NumberFormatLocal", value)
            End Set
        End Property

        Public Property Orientation() As XlOrientation
            Get
                Return InvokeGetProperty(Of XlOrientation)("Orientation")
            End Get
            Set(ByVal value As XlOrientation)
                InvokeSetProperty("Orientation", value)
            End Set
        End Property

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

        Public ReadOnly Property Row() As Integer
            Get
                Return RuleUtil.ConvIndexVBA2DotNET(InvokeGetProperty(Of Integer)("Row"))
            End Get
        End Property

        Public Function Rows() As Range
            Return New Range(Me, InvokeGetProperty("Rows"))
        End Function

        Public Property RowHeight() As Double
            Get
                Return InvokeGetProperty(Of Double)("RowHeight")
            End Get
            Set(ByVal value As Double)
                InvokeSetProperty("RowHeight", value)
            End Set
        End Property

        Public Sub [Select]()
            InvokeMethod("Select")
        End Sub

        Public Property ShrinkToFit() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("ShrinkToFit")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("ShrinkToFit", value)
            End Set
        End Property

        Public Function SpecialCells(ByVal type As XlCellType, Optional ByVal value As Object = Nothing) As Range
            Return New Range(Me, InvokeMethod("SpecialCells", type))
        End Function

        Public ReadOnly Property Text() As String
            Get
                Return InvokeGetProperty(Of String)("Text")
            End Get
        End Property

        Public Sub UnMerge()
            InvokeMethod("UnMerge")
        End Sub

        Public Property Value() As Object
            Get
                Return InvokeGetProperty("Value")
            End Get
            Set(ByVal value As Object)
                InvokeSetProperty("Value", value)
            End Set
        End Property

        Public Property Value2() As Object
            Get
                Return InvokeGetProperty("Value2")
            End Get
            Set(ByVal value As Object)
                InvokeSetProperty("Value2", value)
            End Set
        End Property

        Public Property VerticalAlignment() As XlVAlign
            Get
                Return InvokeGetProperty(Of XlVAlign)("VerticalAlignment")
            End Get
            Set(ByVal value As XlVAlign)
                InvokeSetProperty("VerticalAlignment", value)
            End Set
        End Property

        Public ReadOnly Property Width() As Double
            Get
                Return InvokeGetProperty(Of Double)("Width")
            End Get
        End Property

        Public Property WrapText() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("WrapText")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("WrapText", value)
            End Set
        End Property

    End Class
End Namespace