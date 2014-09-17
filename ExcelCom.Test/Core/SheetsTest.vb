Imports NUnit.Framework

Namespace Core

    Public MustInherit Class SheetsTest

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

        Public Class プロパティたちTest : Inherits SheetsTest

            <Test()> Public Sub Count_最初は0超_シート数はローカルPCの設定で変わる()
                Dim workbook As Workbook = sut.Workbooks.Add
                Assert.That(workbook.Sheets.Count, [Is].GreaterThan(0))
            End Sub

        End Class

        Public Class ExcelObjectたちTest : Inherits SheetsTest

            <Test()> Public Sub Worksheetが閉じられること()
                Dim workbook As Workbook = sut.Workbooks.Add
                Dim worksheet As Worksheet = workbook.Sheets.Item(0)

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

    End Class
End Namespace