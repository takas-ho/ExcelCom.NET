Option Strict Off

Imports System.Runtime.InteropServices.Marshal
Imports NUnit.Framework

Public MustInherit Class LearningExcelComObjectTest

    Private Class TestingApplication : Implements IDisposable
        Private xls As Object

        Public Sub New()
            xls = CreateObject("Excel.Application")
        End Sub

        Public Function Workbooks() As TestingDisposeWorkbooks
            Return New TestingDisposeWorkbooks(xls.Workbooks)
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            FinalReleaseComObject(xls)
        End Sub
    End Class
    Private Class TestingDisposeWorkbooks : Implements IDisposable
        Private workbooks As Object

        Public Sub New(ByVal workbooks As Object)
            Me.workbooks = workbooks
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            FinalReleaseComObject(workbooks)
        End Sub
    End Class
    Private Class TestingException : Inherits Exception
        Public Sub New()
        End Sub

        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
    End Class

    Private Sub AssertNotExistsExcelPropcess()
        System.Threading.Thread.Sleep(100)
        If Process.GetProcessesByName("Excel").Length = 0 Then
            Return
        End If
        Throw New TestingException("Excelプロセスが残っているか、Excelを起動しっぱなしか")
    End Sub

    Public Class Hoge : Inherits LearningExcelComObjectTest

        <Test()> Public Sub ComObject利用中に例外が発生しても_Usingで囲んでいれば解放される()
            Try
                Using x As New TestingApplication
                    Throw New TestingException
                End Using

            Catch ex As TestingException
                AssertNotExistsExcelPropcess()
            End Try
        End Sub

        <Test()> Public Sub ComObjectから別のComObjectを派生させた時_Disposeだと終了しない()
            Dim workbooks As TestingDisposeWorkbooks
            Using x As New TestingApplication
                workbooks = x.Workbooks
                Debug.Print(workbooks.ToString)
            End Using
            Try
                AssertNotExistsExcelPropcess()
                Assert.Fail()
            Catch expected As TestingException
                Assert.IsTrue(True, "Workbooksは明示的にCloseしていないからAssertとなる")
            Finally
                workbooks.Dispose()
            End Try
        End Sub

    End Class

End Class
