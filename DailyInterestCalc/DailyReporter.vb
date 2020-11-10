Imports System.Net.Mail
Imports System.Data.SqlClient
Module DailyReporter
    Dim LenderActivated(20000) As Integer

    Public Function fnDBStringField(ByVal sField) As String
        If IsDBNull(sField) Then
            fnDBStringField = " "
        Else
            fnDBStringField = Trim(CStr(sField))
        End If
    End Function
    Public Function fnDBIntField(ByVal sField) As String
        If IsDBNull(sField) Then
            fnDBIntField = 0
        Else
            fnDBIntField = CInt(sField)
        End If
    End Function

    Public Sub SendSimpleMail(sEmail As String, sSubject As String, sBody As String)
        Dim sPW As String = System.Configuration.ConfigurationManager.AppSettings("DailyReporterPW")
        Dim sUSR As String = Configuration.ConfigurationManager.AppSettings("DailyReporterUSR")
        Dim MyMailMessage As New MailMessage() With {
            .From = New MailAddress(sUSR),
            .Subject = sSubject,
            .IsBodyHtml = True,
            .Body = "<table><tr><td>" + sBody + "</table></td></tr>"
        }
        MyMailMessage.To.Add(sEmail)

        Dim SMTPServer As New SmtpClient("smtp.office365.com") With {
            .Credentials = New System.Net.NetworkCredential(sUSR, sPW),
            .Port = 587,
            .EnableSsl = True
        }

        Try
            SMTPServer.Send(MyMailMessage)
        Catch ex As Exception
            SendErrorMessage(ex)
        End Try
        SMTPServer = Nothing
        MyMailMessage = Nothing
    End Sub


    Public Function ExecuteIT() As String
        Dim sUsers, MySQL, sHTML As String
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dt As New DataTable
        Dim FBSQLEnv As String = System.Configuration.ConfigurationManager.AppSettings("RunFBSQL")
        sUsers = Configuration.ConfigurationManager.AppSettings("EmailList")

        'MySQL = "select u.userid, u.DateCreated regdate, u.firstname, u.lastname, u.address1, u.address2, u.town, u.county, u.postcode, u.activated, u.companyname, u.individorg, uh.newdatecreated actdate, uh.activated activ, u.howhear
        '        from users_history uh, users u
        '        where u.isactive = 0
        '        and uh.newdatecreated > @p1
        '        and u.Activated <= 6
        '        and uh.userid = u.userid
        '        order by uh.newdatecreated desc"

        MySQL = "select u.userid, u.DateCreated regdate, u.firstname, u.lastname, u.address1, u.address2, u.town, u.county, u.postcode, u.activated, u.companyname, u.individorg, u.datecreated actdate, u.activated activ, u.howhear
                from get_user_accounts_history  u
                where u.isactive = 0
                and u.datecreated > @p1
                and u.Activated <= 6
                order by u.datecreated desc"
        If FBSQLEnv = "FB" Then
            Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, New FirebirdSql.Data.FirebirdClient.FbConnection(Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString))
            Adaptor.SelectCommand.Parameters.Clear()
            'Adaptor.SelectCommand.Parameters.Add("@p1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = Now.AddMonths(-1)
            Adaptor.SelectCommand.Parameters.Add("@p1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = Now.AddDays(-5)

            Adaptor.Fill(dt)
        Else
            Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                Try
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter()


                    Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                    con.Open()
                    cmd.Parameters.Clear()
                    With cmd.Parameters
                        .Add(New SqlParameter("@p1", Now.AddDays(-15)))
                    End With
                    adapter.SelectCommand = cmd


                    adapter.Fill(dt)

                    con.Close()
                    con.Dispose()
                Catch ex As Exception
                Finally

                End Try
            End Using
        End If

        sHTML = "<html><body><head>
                <style>
                table {
                    font-family: arial, sans-serif;
                    border-collapse: collapse;
                    width: 100%;
                }

                td, th {
                    border: 1px solid #dddddd;
                    text-align: left;
                    padding: 8px;
                }

                tr:nth-child(even) {
                    background-color: #dddddd;
                }
                </style>
                </head>
                <table>
                  <tr>
                    <th style='font-size:30px' colspan=6>Daily Registration Report</th>
                    <th style='text-align:center' colspan=2>Activation History</th>
                  </tr>
                  <tr>
                    <th>Registration</th>
                    <th>Individ/Co.</th>
                    <th>Name</th>
                    <th>Address</th>
                    <th>Post Code</th>
                    <th>Level</th>
                    <th>Date</th>
                    <th>Level</th>
                  </tr>"

        For Each ThisRow As DataRow In dt.Rows
            Dim iActiv As Integer = fnDBIntField(ThisRow("Activ"))
            Dim iUserID As Integer = fnDBIntField(ThisRow("UserID"))
            If iActiv <> LenderActivated(iUserID) Then
                sHTML &= "<tr><td>" & fnDBStringField(ThisRow("regdate")) & "</td>"
                sHTML &= "<td>" & fnDBStringField(ThisRow("companyname")) & "</td>"
                sHTML &= "<td>" & fnDBStringField(ThisRow("Firstname")) & " " & fnDBStringField(ThisRow("Lastname")) & "</td>"
                sHTML &= "<td>" & fnDBStringField(ThisRow("Address1")) & " " & fnDBStringField(ThisRow("Address2")) & " " & fnDBStringField(ThisRow("Town")) &
                          " " & fnDBStringField(ThisRow("County")) & "</td>"
                sHTML &= "<td>" & fnDBStringField(ThisRow("PostCode")) & "</td>" & "<td>" & fnDBStringField(ThisRow("Activated")) & "</td>"
                sHTML &= "<td>" & fnDBStringField(ThisRow("actdate")) & "</td>" & "<td>" & fnDBStringField(ThisRow("activ")) & "</td></tr>"
                sHTML &= vbNewLine
                LenderActivated(iUserID) = iActiv
            End If
        Next

        sHTML &= "</table></body></html>"

        SendSimpleMail(sUsers, "Daily Report", sHTML)

        ExecuteIT = sHTML
    End Function

    Sub Main()
        ExecuteIT()
    End Sub

    Sub SendErrorMessage(ByVal ThisException As Exception)
        Dim errorPW As String = Configuration.ConfigurationManager.AppSettings("ErrorPW")
        Dim errorUSR As String = Configuration.ConfigurationManager.AppSettings("ErrorUSR")
        Dim mm As New MailMessage() With {
            .From = New MailAddress(errorUSR),
            .Subject = "An Error Has Occurred!",
            .IsBodyHtml = True,
            .Priority = MailPriority.High
        }
        mm.To.Add("web@investandfund.com")

        mm.Body =
            "<html>" & vbCrLf &
            "<body>" & vbCrLf &
            "<h1>An Error Has Occurred!</h1>" & vbCrLf &
            "<table cellpadding=""5"" cellspacing=""0"" border=""1"">" & vbCrLf &
            ItemFormat("Time of Error", DateTime.Now.ToString("dd/MM/yyyy HH:mm:sss"))

        Try
            mm.Body += ItemFormat("Exception Type", ThisException.GetType().ToString())
        Catch ex As Exception
            mm.Body += ItemFormat("Exception Type", "Could not get exception type")
        End Try

        Try
            mm.Body += ItemFormat("Message", ThisException.Message)
        Catch ex As Exception
            mm.Body += ItemFormat("Message", "Could not get message")
        End Try

        Try
            mm.Body += ItemFormat("File Name", "DailyReport.vb")
        Catch ex As Exception
            mm.Body += ItemFormat("File Name", "Could not get file name")
        End Try

        Try
            mm.Body += ItemFormat("Line Number", New StackTrace(ThisException, True).GetFrame(0).GetFileLineNumber)
        Catch ex As Exception
            mm.Body += ItemFormat("Line Number", "Could not get line number")
        End Try

        mm.Body +=
            "</table>" & vbCrLf &
            "</body>" & vbCrLf &
            "</html>"

        Dim smtp As New SmtpClient("smtp.office365.com") With {
            .Credentials = New System.Net.NetworkCredential(errorUSR, errorPW),
            .EnableSsl = True,
            .Port = 587
        }
        smtp.Send(mm)

    End Sub

    Public Function ItemFormat(ByVal Title As String, ByVal Message As String) As String
        Return "  <tr>" & vbCrLf &
                "  <tdtext-align: right;font-weight: bold"">" & Title & ":</td>" & vbCrLf &
                "  <td>" & Message & "</td>" & vbCrLf &
                "  </tr>" & vbCrLf
    End Function

End Module
