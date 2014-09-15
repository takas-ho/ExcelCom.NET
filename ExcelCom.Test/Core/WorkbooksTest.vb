Imports NUnit.Framework

Namespace Core

    Public MustInherit Class WorkbooksTest

        Private sut As Application

        <SetUp()> Public Sub SetUp()
            sut = New Application
        End Sub

        <TearDown()> Public Sub TearDown()
            If sut Is Nothing Then
                Return
            End If
            sut.Dispose()
        End Sub

        Public Class たちTest : Inherits WorkbooksTest

            <Test()> Public Sub Count_最初は0()
                Assert.That(sut.Workbooks.Count, [Is].EqualTo(0))
            End Sub

            <Test()> Public Sub Count_Addすれば1になる()
                sut.Workbooks.Add()
                Assert.That(sut.Workbooks.Count, [Is].EqualTo(1))
            End Sub

        End Class

        Public Class ExcelObjectたちTest : Inherits WorkbooksTest

            <Test()> Public Sub AddのWorkbookが閉じられること()
                Dim workbook As Workbook = sut.Workbooks.Add

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub ItemのWorkbookが閉じられること()
                sut.Workbooks.Add()
                Dim workbook As Workbook = sut.Workbooks.Item(0)

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

    End Class
End Namespace