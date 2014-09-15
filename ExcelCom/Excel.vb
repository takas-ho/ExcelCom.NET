Imports ExcelCom.Core
Imports System.IO

Public Class Excel : Implements IDisposable

    Private ReadOnly app As Application
    Private fileName As String

    Public Sub New()
        Me.New(Nothing)
    End Sub

    Public Sub New(ByVal fileName As String)
        Me.app = New Application
        app.Visible = False

        If Not String.IsNullOrEmpty(fileName) Then
            If File.Exists(fileName) Then
                OpenBook(fileName)
            Else
                CreateBook(fileName)
            End If
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        app.Dispose()
    End Sub

    Private Sub CreateBook(ByVal fileName As String)
        Me.fileName = fileName
        app.Workbooks.Add()
        app.Calculation = Application.XlCalculation.xlCalculationManual
    End Sub

    Public Sub OpenBook(ByVal fileName As String)
        Me.fileName = fileName
        app.Workbooks.Open(fileName, True)
        app.Calculation = Application.XlCalculation.xlCalculationManual
    End Sub

End Class
