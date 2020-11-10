Imports System.Net.Mail
Imports System.Data.SqlClient
Module DailyInterestCalc
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




    Public Function ExecuteIT() As String
        Dim MySQL As String
        Dim Adaptor, Ad2 As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dt, dt1 As New DataTable
        Dim FBSQLEnv As String = System.Configuration.ConfigurationManager.AppSettings("RunFBSQL")
        Dim dr As DataRow
        Dim iCounter As Integer
        MySQL = " select a.ACCOUNTID,  u.ISACTIVE from accounts as a , users as u
                    where a.userid = u.userid
                    and u.isactive = 0"
        If FBSQLEnv = "FB" Then
            Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, New FirebirdSql.Data.FirebirdClient.FbConnection(Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString))
            Adaptor.SelectCommand.Parameters.Clear()


            Adaptor.Fill(dt)
        Else
            Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                Try
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter()


                    Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                    con.Open()
                    cmd.Parameters.Clear()

                    adapter.SelectCommand = cmd


                    adapter.Fill(dt)

                    con.Close()
                    con.Dispose()
                Catch ex As Exception
                Finally

                End Try
            End Using
        End If



        For Each ThisRow As DataRow In dt.Rows

            Dim iaccountid As Integer = fnDBIntField(ThisRow("ACCOUNTID"))
            MySQL = "  select sum(f.amount) as TheTotal, f.accountid
                         from fin_trans f, orders o, loan_holdings lh, loans l 
                         where f.transtype in (1303, 1413) 
                              and f.accountid=@p1

                              and o.orderid = f.orderid 
                              and lh.loan_holdings_id = o.lh_id 
                              and l.loanid = lh.loanid          
                              and l.isActive = 0
                         group by f.accountid"
            If FBSQLEnv = "FB" Then
                Ad2 = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, New FirebirdSql.Data.FirebirdClient.FbConnection(Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString))
                Ad2.SelectCommand.Parameters.Clear()
                Ad2.SelectCommand.Parameters.Add("@p1", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = iaccountid

                Ad2.Fill(dt1)
                iCounter = dt1.Rows.Count
                If iCounter > 0 Then
                    dr = dt1.Rows(0)

                    Dim CmdSql = New FirebirdSql.Data.FirebirdClient.FbCommand
                    Dim myConn = New FirebirdSql.Data.FirebirdClient.FbConnection(Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString)
                    Try
                        Dim taccountid As Integer = fnDBIntField(dr("ACCOUNTID"))
                        Dim tTOTAL As Integer = fnDBIntField(dr("TheTotal"))
                        MySQL = "  insert into GROSSINTERESTAMOUNT (ACCOUNTID, GROSSINTERESTAMOUNT, RUNDATE) values (@ACCOUNTID, @GROSSINTERESTAMOUNT,@RUNDATE)"

                        myConn.Open()
                        CmdSql.Connection = myConn

                        CmdSql.CommandType = Data.CommandType.Text
                        CmdSql.CommandText = MySQL
                        CmdSql.Parameters.Clear()
                        CmdSql.Parameters.Add("@ACCOUNTID", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = taccountid
                        CmdSql.Parameters.Add("@GROSSINTERESTAMOUNT", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = tTOTAL
                        CmdSql.Parameters.Add("@RunDate", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = Now
                        CmdSql.ExecuteNonQuery()

                    Catch ex As Exception

                    Finally
                        CmdSql = Nothing
                        myConn.Close()
                        myConn = Nothing
                        dt1.Clear()

                    End Try

                Else
                    Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                        Try
                            Dim adapter As SqlDataAdapter = New SqlDataAdapter()


                            Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                            con.Open()
                            cmd.Parameters.Clear()
                            With cmd.Parameters
                                .Add(New SqlParameter("@p1", iaccountid))
                            End With
                            adapter.SelectCommand = cmd


                            adapter.Fill(dt1)

                            con.Close()
                            con.Dispose()
                        Catch ex As Exception
                        Finally

                        End Try
                    End Using
                End If
            End If

        Next


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
            mm.Body += ItemFormat("File Name", "DailyInterestCalc.vb")
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
