﻿Namespace Core
    Public Class Range : Inherits AbstractExcelSubObject : Implements IExcelObject

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

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Function AutoFit() As Object
            Return InvokeMethod("AutoFit")
        End Function

        Public Sub Copy(Optional ByVal destination As Object = Nothing)
            Dim args As New List(Of Object)
            If destination IsNot Nothing Then
                args.Add(New NamedParameter("Destination", destination))
            End If
            InvokeMethod("Copy", args.ToArray)
        End Sub

        Public Function Cells() As Range
            Return New Range(Me, InvokeGetProperty("Cells"))
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

        Public ReadOnly Property Count() As Long
            Get
                Return InvokeGetProperty(Of Long)("Count")
            End Get
        End Property

        Public Function Delete(Optional ByVal shift As Object = Nothing) As Object
            Dim args As New List(Of Object)
            If shift IsNot Nothing Then
                args.Add(New NamedParameter("Shift", shift))
            End If
            Return InvokeMethod("Delete", args.ToArray)
        End Function

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

        Public Property NumberFormatLocal() As Object
            Get
                Return InvokeGetProperty("NumberFormatLocal")
            End Get
            Set(ByVal value As Object)
                InvokeSetProperty("NumberFormatLocal", value)
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

        Public ReadOnly Property Text() As String
            Get
                Return InvokeGetProperty(Of String)("Text")
            End Get
        End Property

        Public Property Value() As Object
            Get
                Return InvokeGetProperty("Value")
            End Get
            Set(ByVal value As Object)
                InvokeSetProperty("Value", value)
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