Public Class PayPalPaymentService
    Dim timer As Timers.Timer
    Protected Overrides Sub OnStart(ByVal args() As String)
        'System.Diagnostics.Debugger.Launch()
        timer = New Timers.Timer()
        timer.Interval = 60000
        AddHandler timer.Elapsed, AddressOf TriggerPayPayPaymentUPdate
        timer.Enabled = True
        Dim strFile As String = String.Format("C:\Astro\ServiceLogs\PayPalPaymentVarification_{0}.txt", DateTime.Today.ToString("ddMMMyyyy"))
        IO.File.AppendAllText(strFile, String.Format("Service Started at-- {0}{1}", DateTime.Now, Environment.NewLine))
    End Sub

    Protected Overrides Sub OnStop()
        timer.Enabled = False
        Dim strFile As String = String.Format("C:\Astro\ServiceLogs\PayPalPaymentVarification_{0}.txt", DateTime.Today.ToString("ddMMMyyyy"))
        IO.File.AppendAllText(strFile, String.Format("Service Stopped at-- {0}{1}", DateTime.Now, Environment.NewLine))
    End Sub

    Private Sub TriggerPayPayPaymentUPdate(obj As Object, e As EventArgs)
        Try
            Dim paymntobj As PayPayPaymentUPdate = New PayPayPaymentUPdate()
            paymntobj.PaymentCheck()
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\PayPalPaymentVarification_ErrorLog_{0}.txt", DateTime.Today.ToString("ddMMMyyyy"))
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        End Try
    End Sub
End Class
