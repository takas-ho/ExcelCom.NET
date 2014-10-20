Public Class TestUtil

    Public Shared Sub AssertNotExistsExcelPropcess()
        System.Threading.Thread.Sleep(100)
        If Process.GetProcessesByName("Excel").Length = 0 Then
            Return
        End If
        Throw New InvalidProgramException("Excelプロセスが残っているか、Excelを起動しっぱなしか")
    End Sub

End Class
