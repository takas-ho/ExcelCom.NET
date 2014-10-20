Imports System.Runtime.InteropServices
Imports NUnit.Framework

Namespace Core

    Public MustInherit Class AbstractExcelObjectTest

        Private Class TestingExcelObject : Inherits AbstractExcelObject

            Public Sub New()
                MyBase.New(CreateObject("Excel.Application"))
            End Sub

            Public Function GetProperty(ByVal name As String, ByVal ParamArray args As Object()) As Object
                Return InvokeGetProperty(name, args)
            End Function

            Public Sub SetProperty(ByVal name As String, ByVal ParamArray args As Object())
                InvokeSetProperty(name, args)
            End Sub

            Public Function DoMethod(ByVal name As String, ByVal ParamArray args As Object()) As Object
                Return InvokeMethod(name, args)
            End Function

            Public Sub Close()
                Marshal.FinalReleaseComObject(ComObject)
            End Sub

        End Class

        Public Class [Default] : Inherits AbstractExcelObjectTest

            Private sut As TestingExcelObject

            <SetUp()> Public Sub SetUp()
                sut = New TestingExcelObject
            End Sub

            <TearDown()> Public Sub TearDown()
                sut.Close()
            End Sub

            <Test()> Public Sub InvokeGetProperty_プロパティ値を取得できる()
                Assert.That(sut.GetProperty("Height"), [Is].GreaterThan(0))
            End Sub

            <Test()> Public Sub InvokeGetProperty_プロパティ値を取得できる_大文字小文字違いでも実行できる()
                Assert.That(sut.GetProperty("height"), [Is].GreaterThan(0))
            End Sub

            <Test()> Public Sub InvokeSetProperty_プロパティ値を設定できる()
                Assert.That(sut.GetProperty("Caption"), [Is].Not.EqualTo("HOGE"))

                sut.SetProperty("Caption", "HOGE")

                Assert.That(sut.GetProperty("Caption"), [Is].EqualTo("HOGE"))
            End Sub

            <Test()> Public Sub InvokeMethod_メソッドを実行できる()
                Assert.That(sut.DoMethod("Evaluate", "=2*3"), [Is].EqualTo(6))
            End Sub

            <Test()> Public Sub InvokeMethod_NamedParameterで名前付き引数を指定できる()
                Assert.That(sut.DoMethod("CheckSpelling", "ONE", New NamedParameter("IgnoreUppercase", True)), [Is].True)
            End Sub

            <Test()> Public Sub InvokeMethod_NamedParameterで名前付き引数を指定できる2()
                Assert.That(sut.DoMethod("CheckSpelling", "ONEe", New NamedParameter("IgnoreUppercase", True)), [Is].False)
            End Sub

        End Class

    End Class
End Namespace