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
            comObjects.Add(comObject)
        End Sub

        Private Sub Quit()
            InvokeMethod("Quit")
        End Sub

        Public Function ActiveSheet() As Worksheet
            Dim comObject As Object = InvokeGetProperty("ActiveSheet")
            If comObject Is Nothing Then
                Return Nothing
            End If
            Return New Worksheet(Me, comObject)
        End Function

        Public Function ActiveWorkbook() As Workbook
            Select Case Workbooks.Count
                Case 0
                    Return Nothing
                Case 1
                    Return Workbooks(0)
                Case Else
                    Throw New NotImplementedException("必要なら実装する")
            End Select
        End Function

        Private _workbooks As Workbooks
        Public Function Workbooks() As Workbooks
            If _workbooks Is Nothing Then
                _workbooks = New Workbooks(Me, InvokeGetProperty("Workbooks"))
            End If
            Return _workbooks
        End Function

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

        Public Property DisplayAlerts() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("DisplayAlerts")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("DisplayAlerts", value)
            End Set
        End Property

        Public Property ScreenUpdating() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("ScreenUpdating")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("ScreenUpdating", value)
            End Set
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