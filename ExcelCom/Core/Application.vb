Imports System.Runtime.InteropServices

Namespace Core

    Public Class Application : Inherits AbstractExcelObject : Implements IDisposable, IExcelObject

        Private ReadOnly comObjects As New List(Of Object)
#Region "Xl定数"
        ''' <summary>計算方法</summary>
        Public Enum XlCalculation
            ''' <summary>自動</summary>
            xlCalculationAutomatic = -4105
            ''' <summary>手動</summary>
            xlCalculationManual = -4135
            ''' <summary>テーブル以外自動</summary>
            xlCalculationSemiautomatic = 2
        End Enum

        Public Enum XlSaveAsAccessMode
            xlExclusive = 3
            xlNoChange = 1
            xlShared = 2
        End Enum

        Public Enum XlCutCopyMode
            [False] = 0
            xlCopy = 1
            xlCut = 2
        End Enum
#End Region

        Public Sub New()
            MyBase.New(CreateObject("Excel.Application"))
            DisplayAlerts = False
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dim count As Integer = comObjects.Count
            For subtrahend As Integer = 1 To count
                Marshal.FinalReleaseComObject(comObjects(count - subtrahend))
            Next
            Quit()
            DisplayAlerts = True
            Marshal.FinalReleaseComObject(ComObject)
        End Sub

        Public Sub AddToManager(ByVal comObject As Object) Implements IExcelObject.AddToManager
            If Not comObject.GetType.Name.Contains("ComObject") Then
                Throw New ArgumentException("ComObject型じゃない", "comObject")
            End If
            comObjects.Add(comObject)
        End Sub

        Private Sub Quit()
            InvokeMethod("Quit")
        End Sub

        Public Function ActiveCell() As Range
            Dim comObject As Object = InvokeGetProperty("ActiveCell")
            If comObject Is Nothing Then
                Return Nothing
            End If
            Return New Range(Me, comObject)
        End Function

        Public Function ActiveSheet() As Worksheet
            Dim worksheet As Worksheet = InternalActiveSheet()
            If worksheet Is Nothing Then
                Return Nothing
            End If
            Return ActiveWorkbook.Sheets(worksheet.Index)
        End Function

        Private Function InternalActiveSheet() As Worksheet
            Dim comObject As Object = InvokeGetProperty("ActiveSheet")
            If comObject Is Nothing Then
                Return Nothing
            End If
            Return New Worksheet(InternalActiveWorkbook.Sheets, comObject)
        End Function

        Public Function ActiveWorkbook() As Workbook
            ' InternalActiveWorkbookで作った Workbookは、#Workbooks値の内部Itemとインスタンス違いだから公開しちゃいけない
            Dim index As Integer = DetectIndexOfActiveWorkbook()
            If index < 0 Then
                Return Nothing
            End If
            Return Workbooks(index)
        End Function

        Private Function InternalActiveWorkbook() As Workbook
            Dim comObject As Object = InvokeGetProperty("ActiveWorkbook")
            If comObject Is Nothing Then
                Return Nothing
            End If
            Return New Workbook(Workbooks, comObject)
        End Function

        Private Function DetectIndexOfActiveWorkbook() As Integer
            Dim workbook As Workbook = InternalActiveWorkbook()
            If workbook IsNot Nothing Then
                For i As Integer = 0 To Workbooks.Count - 1
                    If workbook.Name.Equals(Workbooks(i).Name) Then
                        Return i
                    End If
                Next
            End If
            Return -1
        End Function

        Private _workbooks As Workbooks
        Public Function Workbooks() As Workbooks
            If _workbooks Is Nothing Then
                _workbooks = New Workbooks(Me, InvokeGetProperty("Workbooks"))
            End If
            Return _workbooks
        End Function

        Public Function ActiveWindow() As Window
            ' InternalActiveWindowで作った Windowは、#Windows値の内部Itemとインスタンス違いだから公開しちゃいけない
            Dim window As Window = InternalActiveWindow()
            If window Is Nothing Then
                Return Nothing
            End If
            Return Windows(window.Index)
        End Function

        Private Function InternalActiveWindow() As Window
            Dim comObject As Object = InvokeGetProperty("ActiveWindow")
            If comObject Is Nothing Then
                Return Nothing
            End If
            Return New Window(Windows, comObject)
        End Function

        Private _windows As Windows
        Public Function Windows() As Windows
            If _windows Is Nothing Then
                _windows = New Windows(Me, InvokeGetProperty("Windows"))
            End If
            Return _windows
        End Function

        Public ReadOnly Property Build() As Double
            Get
                Return InvokeGetProperty(Of Double)("Build")
            End Get
        End Property

        ''' <summary>計算方法</summary>
        ''' <remarks>※Workbookを開かないとエラー</remarks>
        Public Property Calculation() As XlCalculation
            Get
                Return InvokeGetProperty(Of XlCalculation)("Calculation")
            End Get
            Set(ByVal value As XlCalculation)
                InvokeSetProperty("Calculation", value)
            End Set
        End Property

        Public Property Caption() As String
            Get
                Return InvokeGetProperty(Of String)("Caption")
            End Get
            Set(ByVal value As String)
                InvokeSetProperty("Caption", value)
            End Set
        End Property

        Private _charts As Charts
        Public ReadOnly Property Charts() As Charts
            Get
                If _charts Is Nothing Then
                    _charts = New Charts(Me, InvokeGetProperty("Charts"))
                End If
                Return _charts
            End Get
        End Property

        Public Property CutCopyMode() As XlCutCopyMode
            Get
                Return InvokeGetProperty(Of XlCutCopyMode)("CutCopyMode")
            End Get
            Set(ByVal value As XlCutCopyMode)
                InvokeSetProperty("CutCopyMode", value <> XlCutCopyMode.False)
            End Set
        End Property

        Public Property DisplayAlerts() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("DisplayAlerts")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("DisplayAlerts", value)
            End Set
        End Property

        Public ReadOnly Property Name() As String
            Get
                Return InvokeGetProperty(Of String)("Name")
            End Get
        End Property

        Public ReadOnly Property Path() As String
            Get
                Return InvokeGetProperty(Of String)("Path")
            End Get
        End Property

        Public ReadOnly Property PathSeparator() As String
            Get
                Return InvokeGetProperty(Of String)("PathSeparator")
            End Get
        End Property

        Public ReadOnly Property ProductCode() As String
            Get
                Return InvokeGetProperty(Of String)("ProductCode")
            End Get
        End Property

        Public Property ScreenUpdating() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("ScreenUpdating")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("ScreenUpdating", value)
            End Set
        End Property

        Public Property StandardFont() As String
            Get
                Return InvokeGetProperty(Of String)("StandardFont")
            End Get
            Set(ByVal value As String)
                InvokeSetProperty("StandardFont", value)
            End Set
        End Property

        Public Property StandardFontSize() As Double
            Get
                Return InvokeGetProperty(Of Double)("StandardFontSize")
            End Get
            Set(ByVal value As Double)
                InvokeSetProperty("StandardFontSize", value)
            End Set
        End Property

        Public ReadOnly Property StartupPath() As String
            Get
                Return InvokeGetProperty(Of String)("StartupPath")
            End Get
        End Property

        Public ReadOnly Property TemplatesPath() As String
            Get
                Return InvokeGetProperty(Of String)("TemplatesPath")
            End Get
        End Property

        Public Property UserName() As String
            Get
                Return InvokeGetProperty(Of String)("UserName")
            End Get
            Set(ByVal value As String)
                InvokeSetProperty("UserName", value)
            End Set
        End Property

        Public ReadOnly Property Value() As String
            Get
                Return InvokeGetProperty(Of String)("Value")
            End Get
        End Property

        Public ReadOnly Property Version() As String
            Get
                Return InvokeGetProperty(Of String)("Version")
            End Get
        End Property

        Public Property Visible() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("Visible")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("Visible", value)
            End Set
        End Property

    End Class
End Namespace