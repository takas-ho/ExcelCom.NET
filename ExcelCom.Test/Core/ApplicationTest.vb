Imports NUnit.Framework

Namespace Core

    Public MustInherit Class ApplicationTest

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

        Public Class ExcelObjectたちTest : Inherits ApplicationTest

            <Test()> Public Sub Workbooksが閉じられること()
                Dim workbooks As Workbooks = sut.Workbooks

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

        Public Class PropertyたちTest : Inherits ApplicationTest

            <Test()> Public Sub Calculation(<Values(Application.XlCalculation.xlCalculationAutomatic, Application.XlCalculation.xlCalculationManual, Application.XlCalculation.xlCalculationSemiautomatic)> ByVal value As Application.XlCalculation)
                Dim book As Workbook = sut.Workbooks.Add
                sut.Calculation = value
                Assert.That(sut.Calculation, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub DisplayAlerts(<Values(True, False)> ByVal value As Boolean)
                sut.DisplayAlerts = value
                Assert.That(sut.DisplayAlerts, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub Visible(<Values(True, False)> ByVal value As Boolean)
                sut.Visible = value
                Assert.That(sut.Visible, [Is].EqualTo(value))
            End Sub

        End Class

    End Class
End Namespace