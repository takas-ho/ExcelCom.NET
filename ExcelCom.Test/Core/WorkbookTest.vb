Imports NUnit.Framework

Namespace Core

    Public MustInherit Class WorkbookTest

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

        Public Class ExcelObjectたちTest : Inherits WorkbookTest

            <Test()> Public Sub Sheetsが閉じられること()
                Dim workbook As Workbook = sut.Workbooks.Add
                Dim sheets As Sheets = workbook.Sheets

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

    End Class
End Namespace