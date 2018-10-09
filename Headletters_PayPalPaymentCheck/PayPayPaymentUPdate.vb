Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports Dapper
Imports DapperExtensions
Imports Newtonsoft.Json
Imports SendGrid
Imports SendGrid.Helpers.Mail

Public Class PayPayPaymentUPdate

    Public Function PaymentCheck()
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        'Dim connstr = "data source=NITESH-PC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;"
        Dim connstr = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[);"

        Using cn As New SqlConnection(connstr)

            cn.Open()
            Dim js As New System.Web.Script.Serialization.JavaScriptSerializer

            Dim InvoicePrimaryList = cn.GetList(Of INVOICE_PRIMARY).Where(Function(s) s.PAID_STATUS.Trim() = "N").ToList()
            If (InvoicePrimaryList.Any()) Then

                Dim ADmin = cn.GetList(Of ADBADMIN).FirstOrDefault(Function(s) s.ADBADMINUSERID = "TransactionBatch")
                For Each primary As INVOICE_PRIMARY In InvoicePrimaryList
                    Dim request = cn.GetList(Of PAYMENT_REQUEST_FILE).FirstOrDefault(Function(s) s.ID = primary.PAYMENT_REQUEST)
                    Dim url = "https://api.sandbox.paypal.com/v1/invoicing/invoices/" + primary.INVOICE
                    Dim InvoiceDetail = GetRequest(url)
                    Dim responseInvoiceDetailObj = js.Deserialize(Of InvoiceDetailModel)(InvoiceDetail)
                    If (responseInvoiceDetailObj.status = "PAID") Then
                        Dim batchId = ADmin.BATCHID + 1
                        Dim transactionId = ADmin.TRANSACTIONID + 1

                        primary.PAID_AMOUNT = responseInvoiceDetailObj.paid_amount.paypal.value
                        primary.PAID_STATUS = "Y"
                        primary.PAID_DATE = DateTime.UtcNow

                        primary.PAYMENTTRANSACTION = responseInvoiceDetailObj.payments.FirstOrDefault().transaction_id
                        Dim TopUpInvoice = cn.GetList(Of INVOICE_SECONDARY).FirstOrDefault(Function(s) s.INVOICE.Trim() = primary.INVOICE.Trim() And s.TYPE.Trim() = 1)
                        Dim Tax = cn.GetList(Of INVOICE_SECONDARY).FirstOrDefault(Function(s) s.INVOICE.Trim() = primary.INVOICE.Trim() And s.TYPE.Trim() = 1)
                        Dim ServiceCharge = cn.GetList(Of INVOICE_SECONDARY).FirstOrDefault(Function(s) s.INVOICE.Trim() = primary.INVOICE.Trim() And s.TYPE.Trim() = 1)


                        TopUpInvoice.PAID_AMOUNT = responseInvoiceDetailObj.items.FirstOrDefault.unit_price.value
                        TopUpInvoice.PAID_STATUS = "Y"

                        Dim TopUpInvoiceDate = DateTime.UtcNow.ToString("yyyy-MM-dd HH':'mm':'ss")



                        updateInvoicePrimary(TopUpInvoice.PAID_AMOUNT, TopUpInvoiceDate, TopUpInvoice.PAID_STATUS.ToString(), TopUpInvoice.UID, TopUpInvoice.TYPE, TopUpInvoice.INVOICE)

                        Tax.PAID_AMOUNT = (responseInvoiceDetailObj.items.FirstOrDefault.unit_price.value * responseInvoiceDetailObj.items.FirstOrDefault.tax.percent) / 100
                        Tax.PAID_STATUS = "Y"


                        updateInvoicePrimary(Tax.PAID_AMOUNT, TopUpInvoiceDate, Tax.PAID_STATUS.ToString(), Tax.UID, Tax.TYPE, Tax.INVOICE)


                        ServiceCharge.PAID_AMOUNT = responseInvoiceDetailObj.custom.amount.value
                        ServiceCharge.PAID_STATUS = "Y"

                        updateInvoicePrimary(ServiceCharge.PAID_AMOUNT, TopUpInvoiceDate, ServiceCharge.PAID_STATUS.ToString(), ServiceCharge.UID, ServiceCharge.TYPE, ServiceCharge.INVOICE)


                        Dim User = cn.GetList(Of UPROF).FirstOrDefault(Function(s) s.USERID.Trim() = primary.UID.Trim())
                        Dim Account = cn.GetList(Of ACCOUNT).FirstOrDefault(Function(s) s.ACCOUNT.Trim() = User.USERACCOUNTNO.Trim())
                        Dim TXN = New ACCOUNT_TXN()
                        TXN.TXNACCOUNTUSERID = primary.UID
                        TXN.TXNACCOUNT = User.USERACCOUNTNO
                        TXN.TXNACCOUNTCCY = primary.CURRENCY
                        ' Need to Add type 
                        TXN.TXNTYPE = 2
                        TXN.TXNDATE = DateTime.UtcNow
                        TXN.TXNTIME = DateTime.UtcNow
                        TXN.TXNAMOUNT = primary.PAID_AMOUNT
                        TXN.BATCHID = batchId
                        TXN.TRANSACTIONID = transactionId
                        TXN.TXNCODE = "1111"
                        TXN.TXNIREF = responseInvoiceDetailObj.number
                        TXN.TXNEREF = primary.INVOICE
                        TXN.EXTREF = responseInvoiceDetailObj.links.LastOrDefault.href
                        TXN.NARRATION = "Amount Credit via Paypal"
                        'MONTH
                        Dim month = ""
                        'YEAR
                        Dim year = Date.Now.Year.ToString()
                        If (Date.Now.Month < 10) Then
                            month = "0" + Date.Now.Month.ToString
                        Else
                            month = Date.Now.Month.ToString
                        End If

                        TXN.TXNHID = "999"
                        TXN.TXNRQID = "999"
                        TXN.TXNACCOUNTSTMTSEQ = month + year.Substring(year.Length - 2)

                        TXN.TXNREVERSED = "n"

                        cn.Insert(Of ACCOUNT_TXN)(TXN)

                        Dim TXNTax = New ACCOUNT_TXN()
                        TXNTax.TXNACCOUNTUSERID = primary.UID
                        TXNTax.TXNACCOUNT = User.USERACCOUNTNO
                        TXNTax.TXNACCOUNTCCY = primary.CURRENCY
                        ' Need to Add type 
                        TXNTax.TXNTYPE = 2
                        TXNTax.TXNDATE = DateTime.UtcNow
                        TXNTax.TXNTIME = DateTime.UtcNow
                        TXNTax.TXNAMOUNT = Tax.PAID_AMOUNT
                        TXNTax.BATCHID = batchId
                        TXNTax.TRANSACTIONID = transactionId
                        TXNTax.TXNCODE = "1111"
                        TXNTax.TXNIREF = responseInvoiceDetailObj.number
                        TXNTax.TXNEREF = primary.INVOICE
                        TXNTax.EXTREF = responseInvoiceDetailObj.links.LastOrDefault.href

                        TXNTax.NARRATION = "Amount for Tax debit via paypal"
                        TXNTax.TXNHID = "999"
                        TXNTax.TXNRQID = "999"
                        TXNTax.TXNACCOUNTSTMTSEQ = month + year.Substring(year.Length - 2)

                        TXNTax.TXNREVERSED = "n"

                        cn.Insert(Of ACCOUNT_TXN)(TXNTax)

                        If (ServiceCharge.PAID_AMOUNT > 0) Then

                            Dim TXNServiceCharge = New ACCOUNT_TXN()
                            TXNServiceCharge.TXNACCOUNTUSERID = primary.UID
                            TXNServiceCharge.TXNACCOUNT = User.USERACCOUNTNO
                            TXNServiceCharge.TXNACCOUNTCCY = primary.CURRENCY
                            ' Need to Add type 
                            TXNServiceCharge.TXNTYPE = 2
                            TXNServiceCharge.TXNDATE = DateTime.UtcNow
                            TXNServiceCharge.TXNTIME = DateTime.UtcNow
                            TXNServiceCharge.TXNAMOUNT = ServiceCharge.PAID_AMOUNT
                            TXNServiceCharge.BATCHID = batchId
                            TXNServiceCharge.TRANSACTIONID = transactionId
                            TXNServiceCharge.TXNCODE = "1111"
                            TXNServiceCharge.TXNIREF = responseInvoiceDetailObj.number
                            TXNServiceCharge.TXNEREF = primary.INVOICE
                            TXNServiceCharge.EXTREF = responseInvoiceDetailObj.links.LastOrDefault.href

                            TXNServiceCharge.NARRATION = "Amount debit for service tax via paypal"
                            TXNServiceCharge.TXNHID = "999"
                            TXNServiceCharge.TXNRQID = "999"
                            TXNServiceCharge.TXNACCOUNTSTMTSEQ = month + year.Substring(year.Length - 2)

                            TXNServiceCharge.TXNREVERSED = "n"

                            cn.Insert(Of ACCOUNT_TXN)(TXNServiceCharge)
                        End If
                        Account.ACCOUNTBAL = Account.ACCOUNTBAL + TopUpInvoice.PAID_AMOUNT


                        updateAccount(Account.ACCOUNTBAL.ToString(), Account.ACCOUNT)
                        updatePrimaryStatusInvoice(primary.UID, primary.INVOICE, primary.PAID_AMOUNT, primary.PAID_STATUS, TopUpInvoiceDate, primary.PAYMENTTRANSACTION)


                        Dim htmlBuilder As StringBuilder = New StringBuilder()
                        htmlBuilder.Append(vbCrLf + "<!DOCTYPE html><html lang='en'><head>")
                        htmlBuilder.Append(vbCrLf + "    <meta charset='UTF-8'>")
                        htmlBuilder.Append(vbCrLf + "    <title>Chart</title>")
                        htmlBuilder.Append(vbCrLf + "    <link href='https://fonts.googleapis.com/css?family=Roboto+Slab' rel='stylesheet'>   ")
                        htmlBuilder.Append(vbCrLf + "</head>")
                        htmlBuilder.Append(vbCrLf + "<style>")
                        htmlBuilder.Append(vbCrLf + "</style>")
                        htmlBuilder.Append(vbCrLf + "<body>")
                        htmlBuilder.Append(vbCrLf + "<div style='width:640px;margin:0 auto 40px;border:1px solid #e3e3e3;padding:15px 0 0;border-radius: 3px;box-shadow: 0 3px 5px 0px #E1DAD1;'>")
                        htmlBuilder.Append(vbCrLf + "    <table width='640' cellpadding='0' cellspacing='0' border='0' class='wrapper' style=font - family: Roboto Slab, serif;max-width:880px; width:100%;' bgcolor='#FFFFFF'>")
                        htmlBuilder.Append(vbCrLf + "        <tr>")
                        htmlBuilder.Append(vbCrLf + "        <td height='10' style='font-size:10px; line-height:10px;'>&nbsp;</td>")
                        htmlBuilder.Append(vbCrLf + "        </tr>")
                        htmlBuilder.Append(vbCrLf + "        <tr>")
                        htmlBuilder.Append(vbCrLf + "        <td align='center' valign='top'>")
                        htmlBuilder.Append(vbCrLf + "            <table width='570' cellpadding='0' cellspacing='0' border='0' class='container'>")
                        htmlBuilder.Append(vbCrLf + "            <tr>")
                        htmlBuilder.Append(vbCrLf + "                <td align='left' valign='top'>")
                        htmlBuilder.Append(vbCrLf + "                    <table cellpadding='0' cellspacing='0' border='0'>")
                        htmlBuilder.Append(vbCrLf + "                        <tr>")
                        htmlBuilder.Append(vbCrLf + "                            <td><img src='http://49.50.103.132/LetterHead/Mail/Logo.png' width='160' alt='' title=''></td>")
                        htmlBuilder.Append(vbCrLf + "                            <td width='2' height='10' style='background:#F26F0B;font-size:2px; line-height:10px;'>&nbsp;</td>")
                        htmlBuilder.Append(vbCrLf + "                            <td valign='center'>")
                        htmlBuilder.Append(vbCrLf + "                                <h2 style='padding:0 10px;margin:0px 10px;font-size:14px;color:#F26F0B;text-transform: uppercase;'>Headletters</h2>")
                        htmlBuilder.Append(vbCrLf + "                                <h2 style='padding:0 10px;margin:0px 10px;font-size:14px;color:#F26F0B;text-transform: uppercase;'>1643 DUNDAS ST W APT 27</h2>")
                        htmlBuilder.Append(vbCrLf + "                                <h2 style='padding:0 10px;margin:0px 10px;font-size:14px;color:#F26F0B;text-transform: uppercase;'>TORONTO ON M6K 1V2PN- 1800 -555-5555</h2>")
                        htmlBuilder.Append(vbCrLf + "                            </td>")
                        htmlBuilder.Append(vbCrLf + "                        </tr>")
                        htmlBuilder.Append(vbCrLf + "                    </table>")
                        htmlBuilder.Append(vbCrLf + "                </td>")
                        htmlBuilder.Append(vbCrLf + "")
                        htmlBuilder.Append(vbCrLf + "            </tr>")
                        htmlBuilder.Append(vbCrLf + "            </table>")
                        htmlBuilder.Append(vbCrLf + "        </td>")
                        htmlBuilder.Append(vbCrLf + "        </tr>")
                        htmlBuilder.Append(vbCrLf + "        <tr>")
                        htmlBuilder.Append(vbCrLf + "        <td height='10' style='font-size:10px; line-height:10px;'>&nbsp;</td>")
                        htmlBuilder.Append(vbCrLf + "        </tr>")
                        htmlBuilder.Append(vbCrLf + "        <tr>")
                        htmlBuilder.Append(vbCrLf + "            <td align='left' valign='top'>")
                        htmlBuilder.Append(vbCrLf + "                <table cellpadding='0' cellspacing='0' border='0' width='100%' bgcolor='#F9FAFF'>")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td>")
                        htmlBuilder.Append(vbCrLf + "                            <h2 style='padding:0 20px;color:#2A2D4A;font-size:20px;'>Thanks for using Our Horoscope Services,<br>" + User.USERNAME + "</h2>")
                        htmlBuilder.Append(vbCrLf + "                        </td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "                </table>")
                        htmlBuilder.Append(vbCrLf + "            </td>")
                        htmlBuilder.Append(vbCrLf + "        </tr>")
                        htmlBuilder.Append(vbCrLf + "        <tr>")
                        htmlBuilder.Append(vbCrLf + "            <td>")
                        htmlBuilder.Append(vbCrLf + "                <h3 style='padding:0 20px;color:#2A2D4A;margin:10px 0 0;font-size:16px;'>Money Added Successfully</h3>")
                        htmlBuilder.Append(vbCrLf + "            </td>")
                        htmlBuilder.Append(vbCrLf + "        </tr>")
                        htmlBuilder.Append(vbCrLf + "        <tr>")
                        htmlBuilder.Append(vbCrLf + "            <td>")
                        htmlBuilder.Append(vbCrLf + "                <h3 style='padding:0 20px;color:#2A2D4A;margin:0 0 8px;font-size:16px;'>" + primary.CURRENCY + " " + TopUpInvoice.PAID_AMOUNT.ToString() + "  <img style='height:18px;width:18px;position: relative;top:3px;margin-left:3px;' src='http://49.50.103.132/LetterHead/mail/verify.png' alt='' title=''></h3>")
                        htmlBuilder.Append(vbCrLf + "            </td>")
                        htmlBuilder.Append(vbCrLf + "        </tr>")
                        htmlBuilder.Append(vbCrLf + "        <tr>")
                        htmlBuilder.Append(vbCrLf + "                <td>")
                        htmlBuilder.Append(vbCrLf + "                    <h3 style='padding:0 20px;color:#2A2D4A;margin:0 0 8px;font-size:12px;'>" + responseInvoiceDetailObj.invoice_date + " | Transaction Id :" + primary.PAYMENTTRANSACTION.ToString() + "</h3>")
                        htmlBuilder.Append(vbCrLf + "                </td>")
                        htmlBuilder.Append(vbCrLf + "            </tr>")
                        htmlBuilder.Append(vbCrLf + "        <tr>")
                        htmlBuilder.Append(vbCrLf + "            <td height='1' bgcolor='#4D5163' style='font-size:10px; line-height:1px;'>&nbsp;</td>")
                        htmlBuilder.Append(vbCrLf + "        </tr>")
                        htmlBuilder.Append(vbCrLf + "")
                        htmlBuilder.Append(vbCrLf + "        <tr>")
                        htmlBuilder.Append(vbCrLf + "            <td>")
                        htmlBuilder.Append(vbCrLf + "                <table cellpadding='0' cellspacing='0' border='0' width='100%'>")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td><h3 style='padding:0 20px;color:#2A2D4A;margin:10px 0;font-size:16px;'>Recharge</h3></td>")
                        htmlBuilder.Append(vbCrLf + "                        <td><h3 style='padding:0 20px;color:#2A2D4A;margin:10px 0;font-size:16px;'>" + TopUpInvoice.CURRENCY + " " + TopUpInvoice.PAID_AMOUNT.ToString() + "</h3></td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "                    <!-- <tr>")
                        htmlBuilder.Append(vbCrLf + "        <td height='5' style='font-size:10px; line-height:10px;'>&nbsp;</td>")
                        htmlBuilder.Append(vbCrLf + "    </tr> -->")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td><h3 style='padding:0 20px;color:#2A2D4A;margin:0 0 10px;font-size:16px;'>Tax</h3></td>")
                        htmlBuilder.Append(vbCrLf + "                        <td><h3 style='padding:0 20px;color:#2A2D4A;margin:0 0 10px;font-size:16px;'>" + Tax.CURRENCY + " " + Tax.PAID_AMOUNT.ToString() + "</h3></td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td><h3 style='padding:0 20px;color:#2A2D4A;margin:0 0 10px;font-size:16px;'>Service Charge</h3></td>")
                        htmlBuilder.Append(vbCrLf + "                        <td><h3 style='padding:0 20px;color:#2A2D4A;margin:0 0 10px;font-size:16px;'>" + ServiceCharge.CURRENCY + " " + ServiceCharge.PAID_AMOUNT.ToString() + "</h3></td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td colspan='2' height='1' bgcolor='#4D5163' style='font-size:10px; line-height:1px;'>&nbsp;</td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td><h3 style='padding:0 20px;color:#2A2D4A;margin:10px 0;font-size:16px;'>Total Paid</h3></td>")
                        htmlBuilder.Append(vbCrLf + "                        <td><h3 style='padding:0 20px;color:#2A2D4A;margin:10px 0;font-size:16px;'>" + primary.CURRENCY + " " + primary.PAID_AMOUNT.ToString() + "</h3></td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td height='1' colspan='2' bgcolor='#4D5163' style='font-size:10px; line-height:1px;'>&nbsp;</td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td height='5' style='font-size:10px; line-height:10px;'>&nbsp;</td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td><h3 style='padding:0 20px;color:#2A2D4A;margin:10px 0 0;font-size:16px;'>Your Updated Wallet Balance :</h3></td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td><h3 style='padding:0 20px;color:#2A2D4A;margin:5px 0;font-size:16px;'><img style='height:18px;width:18px;margin-right:3px;position: relative;top:3px;' src='http://49.50.103.132/LetterHead/mail/wallet.png' alt='' title=''> " + primary.CURRENCY + " " + Account.ACCOUNTBAL.ToString() + "</h3></td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td height='15' style='font-size:10px; line-height:15px;'>&nbsp;</td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td height='5' colspan='2' bgcolor='#583921' style='font-size:10px; line-height:5px;'>&nbsp;</td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "                    <tr>")
                        htmlBuilder.Append(vbCrLf + "                        <td height='10' colspan='2' bgcolor='#F26F0B' style='font-size:10px; line-height:10px;'>&nbsp;</td>")
                        htmlBuilder.Append(vbCrLf + "                    </tr>")
                        htmlBuilder.Append(vbCrLf + "")
                        htmlBuilder.Append(vbCrLf + "")
                        htmlBuilder.Append(vbCrLf + "")
                        htmlBuilder.Append(vbCrLf + "                </table>")
                        htmlBuilder.Append(vbCrLf + "            </td>")
                        htmlBuilder.Append(vbCrLf + "        </tr>")
                        htmlBuilder.Append(vbCrLf + "    </table>  ")
                        htmlBuilder.Append(vbCrLf + "")
                        htmlBuilder.Append(vbCrLf + "")
                        htmlBuilder.Append(vbCrLf + "</div>")
                        htmlBuilder.Append(vbCrLf + "")
                        htmlBuilder.Append(vbCrLf + "")
                        htmlBuilder.Append(vbCrLf + "</body>")
                        htmlBuilder.Append(vbCrLf + "")
                        htmlBuilder.Append(vbCrLf + "</html>")
                        Dim template = htmlBuilder.ToString()
                        SendMail(request.EMAIL.Trim(), template)
                        'SendMail(primary.UID.Trim(), template)


                    End If
                    SendMail(primary.UID.Trim(), "abhishek testing")
                    If (responseInvoiceDetailObj.status = "CANCELLED") Then
                        primary.PAID_STATUS = "C"

                        updatePrimaryCancleInvoiceStatus(primary.PAID_STATUS, primary.UID, primary.INVOICE)
                        Dim SecondaryList = cn.GetList(Of INVOICE_SECONDARY).Where(Function(s) s.INVOICE.Trim() = primary.INVOICE.Trim()).ToList()
                        For Each sec As INVOICE_SECONDARY In SecondaryList
                            sec.PAID_STATUS = "C"

                            updateSecondaryCancelInvoiceStatus(sec.PAID_STATUS, sec.UID, sec.INVOICE)
                        Next
                    End If
                Next
            End If


            cn.Close()

        End Using


        Return Nothing
    End Function



    Public Function SendMail(sender As String, template As String) As String
        Dim apiKey = "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"
        Dim client = New SendGridClient(apiKey)
        Dim from = New EmailAddress("kirti.vashishtha@enukesoftware.com", "Kirti Vashishtha")
        Dim subject = "Payment Reciept"
        Dim toSend = New EmailAddress(sender, sender)
        Dim plainTextContent = "Letter Head"

        Dim msg = MailHelper.CreateSingleEmail(from, toSend, subject, plainTextContent, template)
        Dim response = client.SendEmailAsync(msg).Result
        Return response.StatusCode
    End Function


    Public Function GetRequest(url As String) As String
        Dim model = "grant_type=client_credentials"

        Dim byts1 = Encoding.UTF8.GetBytes(model)
        Dim js As New System.Web.Script.Serialization.JavaScriptSerializer
        Dim responseToken = SendTokenRequesttype(New Uri("https://api.sandbox.paypal.com/v1/oauth2/token"), byts1, "application/x-www-form-urlencoded", "POST")
        Dim responseTokenObj = js.Deserialize(Of AccessTokenAPIModel)(responseToken)
        Dim result As String
        Using client As HttpClient = New HttpClient()
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + responseTokenObj.access_token)
            Using response As HttpResponseMessage = client.GetAsync(url).Result

                Using content As HttpContent = response.Content

                    result = content.ReadAsStringAsync().Result



                End Using
            End Using
        End Using


        Return result

    End Function

    Public Function SendRequesttype(uri As Uri, jsonDataBytes As Byte(), contentType As String, method As String) As String
        Dim model = "grant_type=client_credentials"

        Dim byts1 = Encoding.UTF8.GetBytes(model)
        Dim js As New System.Web.Script.Serialization.JavaScriptSerializer
        Dim responseToken = SendTokenRequesttype(New Uri("https://api.sandbox.paypal.com/v1/oauth2/token"), byts1, "application/x-www-form-urlencoded", "POST")
        Dim responseTokenObj = js.Deserialize(Of AccessTokenAPIModel)(responseToken)

        Dim req As WebRequest = WebRequest.Create(uri)
        req.ContentType = contentType
        req.Method = method
        req.ContentLength = jsonDataBytes.Length
        req.Headers.Add("Authorization", "Bearer " + responseTokenObj.access_token)

        req.Headers.Add("contentType", "application/json")

        Dim stream = req.GetRequestStream()
        stream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
        stream.Close()

        Dim response = req.GetResponse().GetResponseStream()

        Dim reader As New StreamReader(response)
        Dim res = reader.ReadToEnd()
        reader.Close()
        response.Close()

        Return res
    End Function

    Public Function SendTokenRequesttype(uri As Uri, jsonDataBytes As Byte(), contentType As String, method As String) As String
        Dim req As WebRequest = WebRequest.Create(uri)
        req.ContentType = contentType
        req.Method = method
        req.ContentLength = jsonDataBytes.Length
        Dim username = "AWwEcFUNZnIuI7SXPgpOq7MmwasKciaVyYArWLhmxYiununTkQF6W9hbdbqpaL57u7lO767ni_uzAAWX"
        Dim password = "EEEv8ntRTZofjqLkm9SIpOvdA3RTz0intXOgvkwxdQ1kbxg_a8I9mY844xmFJDflSuwAfPeRU2XwNK_m"
        Dim encoded = System.Convert.ToBase64String(System.Text.Encoding.GetEncoding("ISO-8859-1").GetBytes(username + ":" + password))
        req.Headers.Add("Authorization", "Basic " + encoded)
        req.Headers.Add("contentType", "application/x-www-form-urlencoded")
        Dim stream = req.GetRequestStream()
        stream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
        stream.Close()

        Dim response = req.GetResponse().GetResponseStream()

        Dim reader As New StreamReader(response)
        Dim res = reader.ReadToEnd()
        reader.Close()
        response.Close()

        Return res
    End Function


    Public Function updateInvoicePrimary(PAID_AMOUNT As String, dateData As String, Status As String, UId As String, type As String, Invoice As String)


        Dim connstr = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[);"
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand

        Try

            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "UPDATE INVOICE_SECONDARY  SET PAID_AMOUNT=" + PAID_AMOUNT.ToString() + ",PAID_DATE='" + dateData.ToString() + "',PAID_STATUS='" + Status.ToString() + "' WHERE UID ='" + UId + "' AND TYPE ='" + type + "' ANd INVOICE ='" + Invoice + "'"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
        Finally
            con.Close()
        End Try
        Return Nothing
    End Function









    Public Function updateAccount(Amount As String, Account As String)

        Dim connstr = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[);"
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand

        Try

            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con

            cmd.CommandText = "UPDATE ACCOUNT  SET ACCOUNTBAL=" + Amount + " WHERE ACCOUNT ='" + Account.Trim() + "'"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
        Finally
            con.Close()
        End Try
        Return Nothing
    End Function


    Public Function updatePrimaryStatusInvoice(UserId As String, Invoice As String, Amount As String, Status As String, DateData As String, Transaction As String)

        Dim connstr = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[);"
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand

        Try

            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con

            cmd.CommandText = "UPDATE INVOICE_PRIMARY  SET PAID_AMOUNT=" + Amount.ToString() + ",PAID_DATE='" + DateData.ToString() + "',PAID_STATUS='" + Status.ToString() + "' WHERE UID ='" + UserId + "' And INVOICE='" + Invoice + "'"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
        Finally
            con.Close()
        End Try
        Return Nothing
    End Function

    Public Function updatePrimaryCancleInvoiceStatus(Status As String, UId As String, Invoice As String)

        Dim connstr = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[);"
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand

        Try

            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "UPDATE INVOICE_PRIMARY  SET PAID_STATUS='" + Status.ToString() + "' WHERE UID ='" + UId + "' ANd INVOICE ='" + Invoice + "'"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
        Finally
            con.Close()
        End Try
        Return Nothing
    End Function
    Public Function updateSecondaryCancelInvoiceStatus(Status As String, UId As String, Invoice As String)

        Dim connstr = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[);"
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand

        Try

            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "UPDATE INVOICE_SECONDARY  SET PAID_STATUS='" + Status.ToString() + "' WHERE UID ='" + UId + "' ANd INVOICE ='" + Invoice + "'"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
        Finally
            con.Close()
        End Try
        Return Nothing
    End Function


End Class






Public Class AccessTokenRequestModel
    Public Property grant_type As String


End Class
Public Class AccessTokenAPIModel
    Public Property scope As String
    Public Property nonce As String
    Public Property access_token As String
    Public Property token_type As String
    Public Property app_id As String
    Public Property expires_in As Decimal
End Class

Public Class Phone
    Public Property country_code As String
    Public Property national_number As String
End Class

Public Class Address
    Public Property line1 As String
    Public Property city As String
    Public Property state As String
    Public Property postal_code As String
    Public Property country_code As String
End Class

Public Class MerchantInfo
    Public Property email As String
    Public Property first_name As String
    Public Property last_name As String
    Public Property business_name As String
    Public Property phone As Phone = New Phone()
    Public Property address As Address = New Address()
End Class

Public Class BillingInfo
    Public Property email As String
    Public Property first_name As String
    Public Property last_name As String
End Class

Public Class CcInfo
    Public Property email As String
End Class
Public Class Custom
    Public Property label As String
    Public Property amount As Amount = New Amount
End Class
Public Class UnitPrice
    Public Property currency As String
    Public Property value As String
End Class

Public Class Amount
    Public Property currency As String
    Public Property value As String
End Class

Public Class Tax
    Public Property name As String
    Public Property percent As Decimal

End Class

Public Class Item
    Public Property name As String
    Public Property quantity As Decimal
    Public Property unit_price As UnitPrice = New UnitPrice()
    Public Property tax As Tax = New Tax()
End Class
Public Class ItemResponse
    Public Property name As String
    Public Property quantity As Decimal
    Public Property unit_price As UnitPrice = New UnitPrice()
    Public Property tax As TaxAmount = New TaxAmount()
End Class

Public Class TotalAmount
    Public Property currency As String
    Public Property value As String
End Class

Public Class MetadataLink
    Public Property created_date As String
    Public Property payer_view_url As String
End Class
Public Class Metadata
    Public Property created_date As String

End Class

Public Class Link
    Public Property rel As String
    Public Property href As String
    Public Property method As String
End Class

Public Class PaymentLinkResponse
    Public Property id As String
    Public Property number As String
    Public Property template_id As String
    Public Property status As String
    Public Property merchant_info As MerchantInfo = New MerchantInfo()
    Public Property cc_info As List(Of CcInfo) = New List(Of CcInfo)
    Public Property billing_info As List(Of BillingInfo) = New List(Of BillingInfo)
    Public Property items As List(Of ItemResponse) = New List(Of ItemResponse)
    Public Property invoice_date As String
    Public Property tax_calculated_after_discount As Boolean
    Public Property tax_inclusive As Boolean
    Public Property note As String
    Public Property custom As Custom = New Custom()
    Public Property logo_url As String
    Public Property total_amount As TotalAmount = New TotalAmount()
    Public Property metadata As Metadata = New Metadata()
    Public Property allow_tip As Boolean

    Public Property links As List(Of Link) = New List(Of Link)

End Class

Public Class Payment
    Public Property type As String
    Public Property transaction_id As String
    Public Property transaction_type As String
    Public Property method As String


End Class


Public Class PaymentLinkRequest
    Public Property merchant_info As MerchantInfo = New MerchantInfo
    Public Property logo_url As String
    Public Property cc_info As List(Of CcInfo) = New List(Of CcInfo)
    Public Property billing_info As List(Of BillingInfo) = New List(Of BillingInfo)
    Public Property items As List(Of Item) = New List(Of Item)
    Public Property custom As Custom = New Custom()
    Public Property note As String
    Public Property terms As String
End Class

Public Class PAYMENT_REQUEST_FILE
    Public Property UID As String
    Public Property HID As String
    Public Property PAYMENT_REF As String
    Public Property AMOUNT As Nullable(Of Decimal)
    Public Property EMAIL As String
    Public Property [DATE] As Nullable(Of Date)
    Public Property TIME As String
    Public Property ID As Integer
    Public Property Currency As String
    Public Property IsProcessed As String

End Class



Public Class INVOICE_PRIMARY
    Public Property UID As String
    Public Property HID As String
    Public Property PAYMENT_REQUEST As String
    Public Property INVOICE As String
    Public Property INV_DATE As Nullable(Of Date)
    Public Property PAYMENT_LINK_REF As String
    Public Property TOTAL_INVOICE_AMOUNT As Nullable(Of Decimal)
    Public Property PAYPAL_REF As String
    Public Property CURRENCY As String
    Public Property PAID_AMOUNT As Nullable(Of Decimal)
    Public Property PAID_DATE As Nullable(Of Date)
    Public Property PAID_STATUS As String
    Public Property PAYMENTTRANSACTION As String

End Class

Public Class INVOICE_SECONDARY
    Public Property UID As String
    Public Property HID As String
    Public Property PAYMENT_REQUEST As String
    Public Property INVOICE As String
    Public Property INV_DATE As Nullable(Of Date)
    Public Property PAYMENT_LINK_REF As String
    Public Property TOTAL_INVOICE_AMOUNT As Nullable(Of Decimal)
    Public Property PAYPAL_REF As String
    Public Property CURRENCY As String
    Public Property PAID_AMOUNT As Nullable(Of Decimal)
    Public Property PAID_DATE As Nullable(Of Date)
    Public Property PAID_STATUS As String
    Public Property TYPE As String

End Class


Public Class ACCOUNT
    Public Property ACCOUNTUSERID As String
    Public Property ACCOUNT As String
    Public Property ACCOUNTCCY As String
    Public Property ACCOUNTBAL As Decimal
    Public Property ACCOUNTCREDIT As Decimal
    Public Property ACCOUNTDEBIT As Decimal
    Public Property ACCOUNTSTMTID As Decimal
End Class

Public Class TaxAmount
    Public Property name As String
    Public Property percent As Decimal
    Public Property amount As Amount
End Class



Public Class InvoiceDetailModel
    Public Property id As String
    Public Property number As String
    Public Property template_id As String
    Public Property status As String
    Public Property merchant_info As MerchantInfo = New MerchantInfo()
    Public Property cc_info As List(Of CcInfo) = New List(Of CcInfo)
    Public Property billing_info As List(Of BillingInfo) = New List(Of BillingInfo)
    Public Property items As List(Of ItemResponse) = New List(Of ItemResponse)
    Public Property invoice_date As String
    Public Property custom As Custom = New Custom()
    Public Property tax_calculated_after_discount As Boolean
    Public Property tax_inclusive As Boolean
    Public Property note As String
    Public Property logo_url As String
    Public Property total_amount As TotalAmount = New TotalAmount()
    Public Property metadata As MetadataLink = New MetadataLink()
    Public Property allow_tip As Boolean
    Public Property links As List(Of Link) = New List(Of Link)
    Public Property payments As List(Of Payment) = New List(Of Payment)
    Public Property paid_amount As PaidAmount = New PaidAmount
End Class
Public Class PaidAmount
    Public Property paypal As Paypal = New Paypal()
    Public Property other As Other = New Other
End Class
Public Class Paypal
    Public Property currency As String
    Public Property value As String
End Class

Public Class Other
    Public Property currency As String
    Public Property value As String
End Class


Public Class UPROF
    Public Property USERID As String
    Public Property USERNAME As String
    Public Property USEREMAIL As String
    Public Property USERIDD As String
    Public Property USERMOBILE As String
    Public Property USERPHOTO As String
    Public Property USERACCOUNTNO As String
    Public Property UCOUNTRY As String
    Public Property UCURRENCY As String
    Public Property UCHARGE As String
    Public Property USERPDATE As Nullable(Of Date)
    Public Property USERPPLANG As String
    Public Property USERLASTHOROID As Decimal
    Public Property PASSWORD As String
    Public Property TOKENFACEBOOK As String
    Public Property TOKENGOOGLE As String
    Public Property TOKENYAHOO As String
    Public Property USERTCFLAG As String
    Public Property TOUCHID As String
    Public Property TOKEN As String
    Public Property USERTC As String

End Class

Public Class ADBADMIN
    Public Property ADBADMINUSERID As String
    Public Property ADBADMINPW As String
    Public Property ADBACCOUNT As Nullable(Of Decimal)
    Public Property ADBUSERUNIQUE As Nullable(Of Decimal)
    Public Property STMNTSEQ As String
    Public Property TRANSACTIONID As Nullable(Of Decimal)
    Public Property BATCHID As Nullable(Of Decimal)
    Public Property ChargeFileLink As String
    Public Property TCLink As String


End Class


Public Class ACCOUNT_TXN
    Public Property TXNACCOUNTUSERID As String
    Public Property TXNACCOUNT As String
    Public Property TXNACCOUNTCCY As String
    Public Property TXNTYPE As String
    Public Property TAXCODE As String
    Public Property TXNCODE As Decimal
    Public Property TXNDATE As Date
    Public Property TXNTIME As String
    Public Property TXNAMOUNT As Decimal
    Public Property TXNIREF As Decimal
    Public Property TXNEREF As String
    Public Property TXNACCOUNTSTMTSEQ As String
    Public Property TXNTAXCODE As String
    Public Property TXNHID As String
    Public Property TXNRQID As String
    Public Property TXNREVERSED As String
    Public Property EXTREF As String
    Public Property ID As Integer
    Public Property BATCHID As Decimal
    Public Property TRANSACTIONID As Decimal
    Public Property NARRATION As String

End Class


