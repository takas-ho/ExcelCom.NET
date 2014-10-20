Imports NUnit.Framework

Namespace Core

    Public MustInherit Class SheetsTest

        Private sut As Application
        Private workbook As Workbook

        <SetUp()> Public Sub SetUp()
            sut = New Application
            workbook = sut.Workbooks.Add
        End Sub

        <TearDown()> Public Sub TearDown()
            If sut Is Nothing Then
                Return
            End If
            sut.Dispose()
        End Sub

        Public Class プロパティたちTest : Inherits SheetsTest

            <Test()> Public Sub Count_最初は0超_シート数はローカルPCの設定で変わる()
                Assert.That(workbook.Sheets.Count, [Is].GreaterThan(0))
            End Sub

        End Class

        Public Class Add_Test : Inherits SheetsTest

            <Test()> Public Sub Addすれば1増える()
                Dim baseCount As Integer = workbook.Sheets.Count
                workbook.Sheets.Add()
                Assert.That(workbook.Sheets.Count, [Is].EqualTo(baseCount + 1))
            End Sub

            <Test()> Public Sub Addしたら先頭に追加される()
                Dim addedSheet As Worksheet = workbook.Sheets.Add()
                addedSheet.Name = "XyZ"
                Assert.That(workbook.Sheets(0).Name, [Is].EqualTo("XyZ"))
            End Sub

            <Test()> Public Sub after引数を指定したら_指定sheetの後ろに追加する()
                Dim firstSheet As Worksheet = workbook.Sheets.Add()

                Dim actual As Worksheet = workbook.Sheets.Add(after:=firstSheet)
                actual.Name = "ABC"

                Assert.That(workbook.Sheets(1).Name, [Is].EqualTo("ABC"), "firstSheetの一つ後ろに追加")
            End Sub

        End Class

        Public Class Item_Test : Inherits SheetsTest

            <Test()> Public Sub AddしたSheetとName指定と同じインスタンスである()
                Dim sheet As Worksheet = workbook.Sheets.Add
                Assert.That(workbook.Sheets(sheet.Name), [Is].SameAs(sheet))
            End Sub

            <Test()> Public Sub Nameとindexとで同じインスタンスである()
                Dim sheet As Worksheet = workbook.Sheets.Add
                Assert.That(workbook.Sheets(sheet.Name), [Is].SameAs(workbook.Sheets(0)))
            End Sub

        End Class

        Public Class ExcelObjectたちTest : Inherits SheetsTest

            <Test()> Public Sub Worksheetが閉じられること()
                Dim worksheet As Worksheet = workbook.Sheets.Item(0)

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

    End Class
End Namespace