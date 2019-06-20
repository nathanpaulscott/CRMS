Module module_scheduled_tasks
    Public Sub periodic_check_launcher(ByVal form As CRMS.main_form, ByVal local As local_machine, ByVal crms_control As crms_controller, ByVal format As cr_sheet_format, ByRef err As String)
        '0	Pending Resubmission	                Resubmitted	1
        '1	Opened	                                Pending Approval,Approved,Not Approved	1
        '1	Resubmitted	                            Pending Approval,Approved,Not Approved	1
        '2	Pending Approval	                    Approved,Not Approved	1
        '3	Approved	                            Pending Execution Planning,Execution Planned, Execution Planning Failed,Approved,Not Approved	1
        '3	Not Approved	                        Pending Resubmission,Resubmitted,Approved,Not Approved	1
        '4	Pending Execution Planning	            Execution Planned, Execution Planning Failed	1
        '5	Execution Planned	                    Pending Execution,Execution Complete, Execution Complete Pending Attachments, Execution Rejected	1
        '5	Execution Planning Failed	            Pending Resubmission,Resubmitted	1
        '6	Pending Execution	                    Execution Complete, Execution Complete Pending Attachments, Execution Rejected	1
        '7	Execution Complete	                    Pending Review,Closed, Review Failed	1
        '7	Execution Complete Pending Attachments	Execution Complete, Pending Execution	1
        '7	Execution Rejected	                    Pending Resubmission,Resubmitted	1
        '8	Pending Review	                        Closed, Review Failed	1
        '9	Closed	                                End	0
        '9	Review Failed	                        Pending Execution	1
        '10	Cancelled	                            End	0

        Dim tx_svr As New smtp_server
        Dim rx_svr_imap As New imap_server
        Dim db As New mysql_server
        Try
            Dim main_err As String = ""

            'initialise the server objects
            '---------------------------------

            rx_svr_imap.Behavior.MessageFetchMode = MessageFetchMode.Minimal
            rx_svr_imap.nat_server_folder = "inbox"
            rx_svr_imap.Host = form.TextBox_imap_server.Text
            rx_svr_imap.Port = Int(Val(form.TextBox_imap_port.Text))
            rx_svr_imap.SslProtocol = Int(Val(form.TextBox_imap_encryption_type.Text))               'SslProtocols.Tls12 'SslProtocols.Default
            rx_svr_imap.UseSsl = form.CheckBox_imap_ssl.Checked
            rx_svr_imap.nat_username = form.TextBox_imap_user.Text
            rx_svr_imap.nat_password = form.TextBox_imap_pass.Text
            rx_svr_imap.nat_cred_regular = New ImapX.Authentication.PlainCredentials(rx_svr_imap.nat_username, rx_svr_imap.nat_password)
            '            Debug.Listeners.Add(New TextWriterTraceListener(local.base_path & "\test.txt"))
            '            Debug.AutoFlush = True
            '            rx_svr_imap.IsDebug = True

            tx_svr.Host = form.TextBox_smtp_server.Text
            tx_svr.Port = form.TextBox_smtp_port.Text
            tx_svr.EnableSsl = form.CheckBox_smtp_ssl.Checked
            Dim a As New NetworkCredential
            a.UserName = form.TextBox_smtp_user.Text
            a.Password = form.TextBox_smtp_pass.Text
            tx_svr.nat_username = a.UserName
            tx_svr.nat_password = a.Password
            tx_svr.Credentials = a
            tx_svr.Timeout = Int(Val(form.TextBox_smtp_timeout.Text))

            db.address = form.TextBox_mysql_server.Text
            db.port = form.TextBox_mysql_port.Text
            db.username = form.TextBox_mysql_user.Text
            db.password = form.TextBox_mysql_password.Text
            db.schema = form.TextBox_mysql_db.Text
            '-----------------------------------------------------------------

            'connect to mysql DB
            '--------------------------
            debug.writeline(Now.ToLongTimeString & ": " & "Connecting to DB....")
            module_db.mysql_connect(db)

            'check HDD space first
            '--------------------------
            err = ""
            debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Doing HDD Check....")
            check_hdd_space(db, tx_svr, format, local, err)
            If Not err Like "" Then main_err = err

            'fix crs with intermediate status's, this could have happened because of previous app crashes or email server issues etc.
            'the output of this will be that all CRs are at 'Pending XXXXX' status
            '------------------------------------------------------------------------------------------------
            err = ""
            debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Doing hung process check....")
            process_hung_status(tx_svr, db, format, local, err)
            If Not err Like "" Then main_err = err

            'now deal with late crs
            '----------------------------
            err = ""
            debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Doing nag analysis....")
            cr_nagging(tx_svr, db, format, local, err)
            If Not err Like "" Then main_err = err

            'clear sentbox and old inbox msgs of email account => uses IMAP for this
            '---------------------------------------------------------------------
            debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Clearing sentbox....")
            err = ""
            clear_mails(rx_svr_imap, format, local, err)
            If Not err Like "" Then main_err = err
            err = main_err
get_out:
        Catch ex As Exception
            err = "ER: Error performing the CR status review, details: " & ex.ToString
        Finally
            If Not err Like "" Then
                Dim subj As String = ""
                Dim body As String = ""
                subj = "CRMS: Periodic Check Error Report"
                body = "The app encountered an error during the periodic check process.<BR>Details: " & err
                body = body & "<P>Thanks<P>"
                auto_reply_general(tx_svr, format.dev_email, "", subj, body, {""}, err)
            End If
            module_db.mysql_disconnect(db)
            tx_svr.Dispose()
            rx_svr_imap.Dispose()
            GC.Collect()
            debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Done....")
        End Try
    End Sub






    Public Sub check_hdd_space(ByVal db As mysql_server, ByVal tx_svr As smtp_server, ByVal format As cr_sheet_format, ByVal local As local_machine, ByRef err As String)
        Try
            Dim combined_flag As Boolean = False
            If Path.GetPathRoot(local.base_path) Like local.db_drive Then combined_flag = True
            For Each drive As DriveInfo In My.Computer.FileSystem.Drives
                If drive.DriveType = DriveType.Fixed And (drive.RootDirectory.FullName = Path.GetPathRoot(local.base_path) Or drive.RootDirectory.FullName = local.db_drive) Then
                    Dim freespace As Double = Math.Round(drive.AvailableFreeSpace / (1073741824), 1)

                    Dim t_array() As String = {format.dev_email}
                    Dim qrows = From row In format.ds_allow.Tables("administrators") Let a = row.Field(Of String)("email") Select a
                    If qrows.Count > 0 Then t_array = t_array.Union(qrows.ToArray).ToArray
                    Dim to_list As String = Join(t_array, ",")

                    If combined_flag Then
                        If freespace < local.base_path_freespace_warning + local.db_drive_freespace_warning Then
                            add2log(db, Now, "warning", "Low Server HDD Space Warning (" & Path.GetPathRoot(local.base_path) & " = " & freespace & "GB) sent to: " & to_list, err)
                            If Not err Like "" Then GoTo get_out

                            Dim subj As String = "CRMS: Server HDD Space Critically Low Warning!"
                            Dim body As String = "The space on the following server drive is critically low (" & Path.GetPathRoot(local.base_path) & " = " & freespace & "GB).  Please take action now or the change request management system will not be able to function.<BR>"
                            err = ""
                            auto_reply_general(tx_svr, to_list, "", subj, body, {}, err)
                        End If

                    Else
                        If freespace < local.base_path_freespace_warning Then
                            add2log(db, Now, "warning", "Low Server HDD Space Warning (" & Path.GetPathRoot(local.base_path) & " = " & freespace & "GB) sent to: " & to_list, err)
                            If Not err Like "" Then GoTo get_out

                            Dim subj As String = "CRMS: Server HDD Space Critically Low Warning!"
                            Dim body As String = "The space on the following server drive is critically low (" & Path.GetPathRoot(local.base_path) & " = " & freespace & "GB).  Please take action now or the change request management system will not be able to function.<BR>"
                            err = ""
                            auto_reply_general(tx_svr, to_list, "", subj, body, {}, err)
                        End If

                        If freespace < local.db_drive_freespace_warning Then
                            add2log(db, Now, "warning", "Low Server HDD Space Warning (" & local.db_drive & " = " & freespace & "GB) sent to: " & to_list, err)
                            If Not err Like "" Then GoTo get_out

                            Dim subj As String = "CRMS: Server HDD Space Critically Low Warning!"
                            Dim body As String = "The space on the following server drive is critically low (" & local.db_drive & " = " & freespace & "GB).  Please take action now or the change request management system will not be able to function.<BR>"
                            err = ""
                            auto_reply_general(tx_svr, to_list, "", subj, body, {}, err)
                        End If
                    End If
                End If
            Next
get_out:
        Catch ex As Exception
            err = "ER: There was an error checking the HDD freespace, details: " & ex.ToString
        End Try
    End Sub




    'note: these are mainly needed when we use POP as POP doesn't maintain the mailboxes on the server properly.
    Public Sub clear_mails(ByVal rx_svr_imap As imap_server, ByVal format As cr_sheet_format, ByVal local As local_machine, ByRef err As String)
        Try
            'connect to imap svr
            '--------------------------
            clean_discon_from_imap(rx_svr_imap)
            If Not rx_svr_imap.IsConnected And Not rx_svr_imap.IsAuthenticated Then
                debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Connecting....")
            Else
                debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Can't pre-Disconnect....")
                Throw New Exception("Can't pre-disconnect from IMAP sever")
            End If

            clean_connect_to_imap(rx_svr_imap, local)
            If rx_svr_imap.IsConnected And rx_svr_imap.IsAuthenticated Then
                debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Clearing Mailboxes....")

                'Does the sentbox
                '-------------------
                Dim i As Integer = 0
                Dim j As Integer = 0
                Dim msg_query = From msg In rx_svr_imap.Folders.Sent.Messages Select msg
                i = msg_query.Count
                If i > 0 Then
                    For j = 1 To i
                        msg_query(0).Remove()
                    Next
                End If

                'Does the inbox - removes msg of more than 72 hrs old - seen or not
                '-------------------------------------------------------------------
                msg_query = From msg In rx_svr_imap.Folders.Inbox.Messages
                            Where Not msg.Date.HasValue Or msg.Date.Value < Now.AddHours(-72)
                            Select msg
                i = msg_query.Count
                If i > 0 Then
                    For j = 1 To i
                        msg_query(0).Remove()
                    Next
                End If


            ElseIf rx_svr_imap.IsConnected Then
                debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Couldn't Login")
                Throw New Exception("Can't login to IMAP sever")
            Else
                debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Couldn't Connect")
                Throw New Exception("Can't connect to IMAP sever")
            End If

            If rx_svr_imap.IsConnected Then
                debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Disconnecting....")

                clean_discon_from_imap(rx_svr_imap)
                If Not rx_svr_imap.IsConnected And Not rx_svr_imap.IsAuthenticated Then
                    debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Idle")
                Else
                    debug.writeline(Now.ToLongTimeString & ": " & "Periodic Status Check: Couldn't Disconnect")
                    Throw New Exception("Can't final discon from IMAP server")
                End If
            End If

        Catch ex As Exception
            err = "ER: Error connecting to IMAP server"
        End Try
    End Sub



    Public Sub cr_nagging(ByVal tx_svr As smtp_server, ByVal db As mysql_server, format As cr_sheet_format, ByVal local As local_machine, ByRef err As String)
        Try
            Dim main_err As String = ""
            Dim dt As New System.Data.DataTable
            Dim sqltext As String = ""
            Dim time_now As DateTime = Now

            'NOTE: open_date is written to the DB via the split sub when it reads the sheet vals to DS then writes to DB, so we do not see it in the logic code, the open date is a full date/time value
            'NOTE: when a resubmission is accepted, it is written to the DB in the same way as for a new CR, so the open date gets overwritten to the resub accept date, this is ok
            'NOTE: approved_date is just 'Now' => the time we got the approval => written to the DB in the logic sub
            'NOTE: planned_ex_date is dimmed before the process_ex_coord_cr sub and set in the process_ex_coord_cr sub, it is the latest planned_ex_date set to 6 hours after 'Now' if the date is today and 6pm if it is another day.  This is then a full date/time value.  It is written the DB in the logic sub.
            'NOTE: ex_date is written to the DB in the logic sub and to the xl sheet in the process ex cr form sub, we get it from 'Now' in the logic sub and we have to be careful to only write it once if the cr_status is not pending attachments, ex_date is a dull date/time value and is the time we recieved the ex_results email, not the actual ex date of any of the sub CRs  
            'NOTE: closed_date is just 'Now' => the time we got the review response => written to the DB in the logic sub
            'yes this is all very confusing and down to poor SW design.

            'get crs which are late for resubmission
            '---------------------------------------
            time_now = Now
            err = ""
            clear_dt(dt)
            sqltext = make_time_sql("requester", "^(Pending[[:space:]]Resubmission)", "resubmission_lifetime", "resubmission_nag_period", "last_activity_date", time_now, db)       'time ref is the last activity date as this would be the date of approval rejection or ex plannign rejection or ex rejection
            sqlquery(False, db, sqltext, dt, err)
            If err Like "" Then
                For Each row As System.Data.DataRow In dt.Rows
                    Dim subj As String = "CRMS: CR Resubmission Request (" & row("cr_id") & ") - OUTSTANDING RESUBMISSION REMINDER"
                    Dim body As String = "You have an outstanding resubmission request (" & row("overdue") & " hours overdue),<BR>Please reply to this mail (or the original resubmit request) with your modified resubmit form ASAP or cancel the CR if you do not wish to continue.<BR>The resubmit form is attached to the original request (subject => CRMS: CR Resubmission Request (" & row("cr_id") & ")) sent to you on " & DateTime.Parse(row("date_request_sent")).ToLongDateString & " at " & DateTime.Parse(row("date_request_sent")).ToShortTimeString & "hrs.<P>Thanks<BR>"
                    auto_reply_general(tx_svr, c2e(row("to_list")), "", subj, body, {""}, err)
                    update_cr_common_table_date(db, row("cr_id"), "last_nag_date", time_now, False, err)
                Next
            End If
            If Not err Like "" Then
                main_err = err
            End If

            'get crs which are late for approval
            '----------------------------------------------
            time_now = Now
            err = ""
            clear_dt(dt)
            sqltext = make_time_sql("approver", "^(Pending[[:space:]]Approval)", "approval_lifetime", "approval_nag_period", "last_activity_date", time_now, db)
            sqlquery(False, db, sqltext, dt, err)
            If err Like "" Then
                For Each row As System.Data.DataRow In dt.Rows
                    Dim subj As String = "CRMS: CR Approval Request (" & row("cr_id") & ") - OUTSTANDING APPROVAL REMINDER"
                    Dim body As String = "You have an outstanding approval request (" & row("overdue") & " hours overdue),<BR>Please reply to this mail with a yes/ok or no/nok answer in the email body ASAP.<BR>The CR form is attached to the original request (subject => CRMS: CR Approval Request (" & row("cr_id") & ")) sent to you on " & DateTime.Parse(row("date_request_sent")).ToLongDateString & " at " & DateTime.Parse(row("date_request_sent")).ToShortTimeString & "hrs.<P>Thanks<BR>"
                    auto_reply_general(tx_svr, c2e(row("to_list")), "", subj, body, {""}, err)
                    update_cr_common_table_date(db, row("cr_id"), "last_nag_date", time_now, False, err)
                Next
            End If
            If Not err Like "" Then
                main_err = err
            End If

            'get crs which are late for execution planning
            '-------------------------------------------
            time_now = Now
            err = ""
            clear_dt(dt)
            sqltext = make_time_sql("execution_coordinator", "^(Pending[[:space:]]Execution[[:space:]]Planning)", "execution_planning_lifetime", "execution_planning_nag_period", "approval_date", time_now, db)
            sqlquery(False, db, sqltext, dt, err)
            If err Like "" Then
                For Each row As System.Data.DataRow In dt.Rows
                    Dim subj As String = "CRMS: CR Execution Planning Request (" & row("cr_id") & ") - OUTSTANDING EXECUTION PLANNING REMINDER"
                    Dim body As String = "You have an outstanding execution planning request (" & row("overdue") & " hours overdue).<BR>Please reply to this mail with the completed CR form ASAP.<BR>The CR form is attached to the original request (subject => CRMS: CR Execution Planning Request (" & row("cr_id") & ")) sent to you on " & DateTime.Parse(row("date_request_sent")).ToLongDateString & " at " & DateTime.Parse(row("date_request_sent")).ToShortTimeString & "hrs.<P>Thanks<BR>"
                    auto_reply_general(tx_svr, c2e(row("to_list")), "", subj, body, {""}, err)
                    update_cr_common_table_date(db, row("cr_id"), "last_nag_date", time_now, False, err)
                Next
            End If
            If Not err Like "" Then
                main_err = err
            End If

            'get crs which are late for execution complete with complete attchments
            '-----------------------------------------------------------------
            time_now = Now
            err = ""
            clear_dt(dt)
            sqltext = make_time_sql("executors", "^(Pending[[:space:]]Execution)", "execution_lifetime", "execution_nag_period", "planned_execution_date", time_now, db)
            sqlquery(False, db, sqltext, dt, err)
            If err Like "" Then
                For Each row As System.Data.DataRow In dt.Rows
                    Dim subj As String = "CRMS: CR Execution Request (" & row("cr_id") & ") - OUTSTANDING EXECUTION REMINDER"
                    Dim body As String = "You have an outstanding execution request (" & row("overdue") & " hours overdue).<BR>Please reply to this mail with the completed CR form and verification attchments ASAP.<BR>The CR form is attached to the original request (subject => CRMS: CR Execution Request (" & row("cr_id") & ")) sent to you on " & DateTime.Parse(row("date_request_sent")).ToLongDateString & " at " & DateTime.Parse(row("date_request_sent")).ToShortTimeString & "hrs.<P>Thanks<BR>"
                    Dim to_list As String = check_email_list(True, row("to_list"), format)
                    auto_reply_general(tx_svr, to_list, "", subj, body, {""}, err)
                    update_cr_common_table_date(db, row("cr_id"), "last_nag_date", time_now, False, err)
                Next
            End If
            If Not err Like "" Then
                main_err = err
            End If

            'get crs which are executed but missing attachments
            '-------------------------------------------------------
            time_now = Now
            err = ""
            clear_dt(dt)
            sqltext = make_time_sql("executors", "^Execution[[:space:]]Complete[[:space:]]Pending[[:space:]]Attachments", "missing_ex_attach_lifetime", "missing_ex_attach_nag_period", "execution_date", time_now, db)
            sqlquery(False, db, sqltext, dt, err)
            If err Like "" Then
                For Each row As System.Data.DataRow In dt.Rows
                    Dim subj As String = "CRMS: CR Execution Complete But Missing Attachments (" & row("cr_id") & ") - MISSING ATTACHMENTS REMINDER"
                    Dim body As String = "You have outstanding execution attachments (" & row("overdue") & " hours overdue).<BR>Please reply to this mail with the missing attachments ASAP.<P>Thanks<BR>"
                    Dim to_list As String = check_email_list(True, row("to_list"), format)
                    auto_reply_general(tx_svr, to_list, "", subj, body, {""}, err)
                    update_cr_common_table_date(db, row("cr_id"), "last_nag_date", time_now, False, err)
                Next
            End If
            If Not err Like "" Then
                main_err = err
            End If

            'get crs which are late for review
            '-----------------------------------
            time_now = Now
            err = ""
            clear_dt(dt)
            sqltext = make_time_sql("requester", "^(Pending[[:space:]]Review)", "review_lifetime", "review_nag_period", "last_activity_date", time_now, db)
            sqlquery(False, db, sqltext, dt, err)
            If err Like "" Then
                For Each row As System.Data.DataRow In dt.Rows
                    Dim subj As String = "CRMS: CR Review Request (" & row("cr_id") & ") - OUTSTANDING REVIEW REMINDER"
                    Dim body As String = "You have an outstanding review request (" & row("overdue") & " hours overdue).<BR>Please reply to this mail with a yes/ok or no/nok answer in the email body ASAP.<BR>The CR form and executor attachments are attached to the original request (subject => CRMS: CR Review Request (" & row("cr_id") & ")) sent to you on " & DateTime.Parse(row("date_request_sent")).ToLongDateString & " at " & DateTime.Parse(row("date_request_sent")).ToShortTimeString & "hrs.<P>Thanks<BR>"
                    auto_reply_general(tx_svr, c2e(row("to_list")), "", subj, body, {""}, err)
                    update_cr_common_table_date(db, row("cr_id"), "last_nag_date", time_now, False, err)
                Next
            End If
            If Not err Like "" Then
                main_err = err
            End If
            err = main_err
        Catch ex As Exception
            err = "ER: Error performing the CR status review, details: " & ex.ToString
        End Try
    End Sub







    Public Sub process_hung_status(ByVal tx_svr As smtp_server, ByVal db As mysql_server, format As cr_sheet_format, ByVal local As local_machine, ByRef err As String)
        Try
            Dim time_now As DateTime = Now
            Dim dt As New System.Data.DataTable
            Dim sqltext As String = ""

            'For each of these that need to be pushed, it is typically because the email send failed when sending the request for the next part.  So we just have to send the mail again and change the status

            'get crs which are opened/resubmitted, they need to go to approval
            '---------------------------------------------------------
            sqltext = "SELECT * FROM " & db.schema & ".cr_common WHERE cr_status REGEXP '^(Opened)|(Resubmitted)';"
            sqlquery(False, db, sqltext, dt, err)
            If Not err Like "" Then GoTo get_out
            For Each row As System.Data.DataRow In dt.Rows
                Dim subj As String = " - resend"
                Dim body As String = "NOTE: Due to a network error, this approval request may not have been sent previously so it is being resent now.<BR>If you already have this request for this CR, please ignore this mail.<P>"
                subj = "CRMS: CR Approval Request (" & row("cr_id") & ")" & subj
                body = body & "Please check the attached CR for approval.<BR>"
                body = body & "Acceptable Responses:<BR>----------------------------------------------------" & array2html_bullet_list({"You accept the CR => reply to the email and type 'ok/yes/accepted' in the email body.", "You do not accept the CR => reply to the email and type 'not ok/nok/no/not accepted' in the email body.<BR>You can add any text after this for your reason or notes."})
                body = body & "<P>Thanks<P>"
                repair_cr_state(row, "Pending Approval", "CR approval request resend,", subj, body, "approver", time_now, "requester", tx_svr, db, format, local, err)
            Next

            'get crs which are approved, they need to go to pending ex planning
            '----------------------------------------------
            clear_dt(dt)
            sqltext = "SELECT * FROM " & db.schema & ".cr_common WHERE cr_status REGEXP '^(Approved)';"
            sqlquery(False, db, sqltext, dt, err)
            If Not err Like "" Then GoTo get_out
            For Each row As System.Data.DataRow In dt.Rows
                Dim subj As String = " - resend"
                Dim body As String = "NOTE: Due to a network error, this approval request may not have been sent previously so it is being resent now.<BR>If you already have this request for this CR, please ignore this mail.<P>"
                subj = "CRMS: CR Execution Planning Request (" & row("cr_id") & ")" & subj
                body = body & "Please open the attched CR and enter the expected/planned date of execution in the yellow field, then close/save/attach/reply with the CR."
                body = body & "<BR>Acceptable Responses:<BR>-----------------------------------------------------" & array2html_bullet_list({"You have no issues => reply to the email and attach the updated cr form.", "You have issues => reply to the email and type 'not ok/nok/no' in the email body.<BR>You can add any text after this for your reason or notes."})
                body = body & "<P>Thanks<P>"
                repair_cr_state(row, "Pending Execution Planning", "CR execution planning request resend,", subj, body, "execution_coordinator", time_now, "requester", tx_svr, db, format, local, err)
            Next

            'get crs which are not approved, they need to go to pending resub
            '----------------------------------------------
            clear_dt(dt)
            sqltext = "SELECT * FROM " & db.schema & ".cr_common WHERE cr_status REGEXP '^(Not[[:space:]]Approved)';"
            sqlquery(False, db, sqltext, dt, err)
            If Not err Like "" Then GoTo get_out
            For Each row As System.Data.DataRow In dt.Rows
                Dim subj As String = " - resend"
                Dim body As String = "NOTE: Due to a network error, this approval request may not have been sent previously so it is being resent now.<BR>If you already have this request for this CR, please ignore this mail.<P>"
                subj = "CRMS: CR Resubmission Request (" & row("cr_id") & ") - !!Approval Rejection!!" & subj
                body = body & "Your CR (" & row("cr_id") & ") was not approved.<BR>If you want to continue, make modifications to the resubmit form and resubmit it."
                body = body & "<P>Thanks<P>"
                repair_cr_state(row, "Pending Resubmission", "CR approval rejection note re-sent to: ", subj, body, "requester", time_now, "resub", tx_svr, db, format, local, err)
            Next

            'get crs which are execution planned, they need to go to pending ex
            '----------------------------------------------
            clear_dt(dt)
            sqltext = "SELECT * FROM " & db.schema & ".cr_common WHERE cr_status REGEXP '^(Execution[[:space:]]Planned)';"
            sqlquery(False, db, sqltext, dt, err)
            If Not err Like "" Then GoTo get_out
            For Each row As System.Data.DataRow In dt.Rows
                Dim subj As String = " - resend"
                Dim body As String = "NOTE: Due to a network error, this approval request may not have been sent previously so it is being resent now.<BR>If you already have this request for this CR, please ignore this mail.<P>"
                subj = "CRMS: CR Execution Request (" & row("cr_id") & ")" & subj
                body = body & "Please execute the attached CR by the planned date indicated.<BR>"
                body = body & "After execution, please complete all yellow fields in the CR then close/save/attach/reply with the CR<BR>"
                body = body & "For the reviewer to accept the result of the CR, please attach the verification files with:" & array2html_bullet_list({"Parameter/Hardware CRs - CR id in the file name", "RF Basic/RF Re-engineering CRs - nodename in the filename"})
                body = body & "<P>Thanks<P>"
                repair_cr_state(row, "Pending Execution", "CR execution request resend,", subj, body, "executors", time_now, "requester", tx_svr, db, format, local, err)
            Next

            'get crs which are execution planning failed, they need to go to pending resub
            '----------------------------------------------
            clear_dt(dt)
            sqltext = "SELECT * FROM " & db.schema & ".cr_common WHERE cr_status REGEXP '^(Execution[[:space:]]Planning[[:space:]]Failed)';"
            sqlquery(False, db, sqltext, dt, err)
            If Not err Like "" Then GoTo get_out
            For Each row As System.Data.DataRow In dt.Rows
                Dim subj As String = " - resend"
                Dim body As String = "NOTE: Due to a network error, this approval request may not have been sent previously so it is being resent now.<BR>If you already have this request for this CR, please ignore this mail.<P>"
                subj = "CRMS: CR Resubmission Request (" & row("cr_id") & ") - !!Execution Planning Rejection!!" & subj
                body = body & "Your CR (" & row("cr_id") & ") could not have it's execution planned.<BR>If you want to continue, make modifications to the resubmit form and resubmit it."
                body = body & "<P>Thanks<P>"
                repair_cr_state(row, "Pending Resubmission", "CR execution planning rejection note re-sent to: ", subj, body, "requester", time_now, "resub", tx_svr, db, format, local, err)
            Next

            'get crs which are executed, they need to go to pending review
            '----------------------------------------------
            clear_dt(dt)
            sqltext = "SELECT * FROM " & db.schema & ".cr_common WHERE cr_status REGEXP '^(Execution[[:space:]]Complete)';"
            sqlquery(False, db, sqltext, dt, err)
            If Not err Like "" Then GoTo get_out
            For Each row As System.Data.DataRow In dt.Rows
                Dim subj As String = " - resend"
                Dim body As String = "NOTE: Due to a network error, this approval request may not have been sent previously so it is being resent now.<BR>If you already have this request for this CR, please ignore this mail.<P>"
                subj = "CRMS: CR Review Request (" & row("cr_id") & ")" & subj
                body = body & "Please review the attached CR form which has been executed."
                body = body & "<BR>Acceptable Responses:<BR>"
                body = body & "-----------------------------------------------------"
                body = body & array2html_bullet_list({"You have no issues => reply to the email and type 'ok/yes/accepted' in the email body.", "You have issues => reply to the email and type 'not ok/nok/no/not accepted' in the email body.<BR>You can add any text after this for your reason or notes."})
                body = body & "<P>Thanks<P>"
                repair_cr_state(row, "Pending Review", "CR review request re-sent to: ", subj, body, "requester", time_now, "executor", tx_svr, db, format, local, err)
            Next

            'get crs which are not executed, they need to go to pending resub
            '----------------------------------------------
            clear_dt(dt)
            sqltext = "SELECT * FROM " & db.schema & ".cr_common WHERE cr_status REGEXP '^(Execution[[:space:]]Rejected)';"
            sqlquery(False, db, sqltext, dt, err)
            If Not err Like "" Then GoTo get_out
            For Each row As System.Data.DataRow In dt.Rows
                Dim subj As String = " - resend"
                Dim body As String = "NOTE: Due to a network error, this approval request may not have been sent previously so it is being resent now.<BR>If you already have this request for this CR, please ignore this mail.<P>"
                subj = "CRMS: CR Resubmission Request (" & row("cr_id") & ") - !!Execution Rejection!!" & subj
                body = body & "Your CR (" & row("cr_id") & ") could not be executed.<BR>If you want to continue, make modifications to the resubmit form and resubmit it."
                body = body & "<P>Thanks<P>"
                repair_cr_state(row, "Pending Resubmission", "CR execution rejection note re-sent to: ", subj, body, "requester", time_now, "resub", tx_svr, db, format, local, err)
            Next

            'get crs which are review failed, they need to go to pending ex
            '----------------------------------------------
            clear_dt(dt)
            sqltext = "SELECT * FROM " & db.schema & ".cr_common WHERE cr_status REGEXP '^(Review[[:space:]]Failed)';"
            sqlquery(False, db, sqltext, dt, err)
            If Not err Like "" Then GoTo get_out
            For Each row As System.Data.DataRow In dt.Rows
                Dim subj As String = " - resend"
                Dim body As String = "NOTE: Due to a network error, this approval request may not have been sent previously so it is being resent now.<BR>If you already have this request for this CR, please ignore this mail.<P>"
                subj = "CRMS: CR Execution Request (" & row("cr_id") & ") - !!Review Failed!!" & subj
                body = body & "Please NOTE!  The CR (" & row("cr_id") & ") sent in for review (" & row("requester") & "), was not accepted.<BR>"
                body = body & "Please review, rectify the issues and send the correct information back for further review."
                body = body & "<P>Thanks<P>"
                repair_cr_state(row, "Pending Execution", "CR review failure note sent to: ", subj, body, "executors", time_now, "reex", tx_svr, db, format, local, err)
            Next
get_out:
        Catch ex As Exception
            err = "ER: Error performing the CR status review, details: " & ex.ToString
        End Try
    End Sub



    Public Sub repair_cr_state(ByVal row As System.Data.DataRow, ByVal target_status As String, ByVal log_entry_Text As String, ByVal subj As String, ByVal body As String, ByVal to_spec As String, ByVal time_now As DateTime, ByVal attach_type As String, ByVal tx_svr As smtp_server, ByVal db As mysql_server, ByVal format As cr_sheet_format, ByVal local As local_machine, err As String)
        Dim cr_form_type As String = row("cr_form_type")
        Dim cr_id As String = row("cr_id")
        Dim cc_list As String = ""
        Dim cc_list_final As String = ""
        Dim user_list As String = ""
        Dim requester As String = ""
        Dim approver As String = ""
        Dim execution_coordinator As String = ""
        Dim executors As String = ""
        Dim to_list As String = ""
        Try
            'find the CR form
            '------------------------
            'NOTE: I AM NOT CHECKING THAT THE CR FORM HAS  THE CORRECT FORMATTING FOR THE NEXT STAGE, i AM ASSUMING IT WAS DONE PREVIOUSLY AS THAT IS WHY WE HAD THAT STATUS, SO i THINK
            'IT IS SAFE ENOUGH TO JUST SEND IT.  IF THEY GET A SKETCHY CR FORM, JUST TELL THEM TO CANCEL IT AND START AGAIN.
            Dim cr_form As String = local.base_path & local.cr & "\" & cr_id & "\" & cr_id & ".xlsb"
            If Not FileIO.FileSystem.FileExists(cr_form) Then
                err = "ER: CR form doesn't exist"       'need to put code after the caller to deal with this case, do not get_out, just put the status to cancelled and skip this one
                GoTo get_out
            End If

            'don't think I need this, if a CR is at an intermediate state, it will have a completed form always, if it doesn't, let the user sort it out, they can cancel it
            'preps the cr_form if required
            '-------------------------------
            '            If target_status Like "Pending Execution Planning" Then
            '            Dim approval_date As DateTime = Now
            '            If Not row("approval_date").ToString Like "" Then           'test this if the date in the dB is null
            '            approval_date = row("approval_date")
            '            End If
            '            prepare_cr_form_for_excoord(approval_date, cr_form, cr_form_type, format, err)
            '            If Not err Like "" Then GoTo get_out
            '            ElseIf target_status Like "Pending Execution" Then
            '            Dim planned_execution_date As DateTime = Now
            '            If Not row("planned_execution_date").ToString Like "" Then           'test this if the date in the dB is null
            '            planned_execution_date = row("planned_execution_date")
            '            End If
            '            Dim executors_raw As String = ""
            '            Dim dt As New System.Data.DataTable
            '            process_ex_coord_cr_form(cr_form, planned_execution_date, executors_raw, dt, cr_form_type, format, err)
            '            If Not err Like "" Then GoTo get_out
            '            update_cr_data(db, dt, "cr_data_" & cr_form_type, err)
            '            If Not err Like "" Then GoTo get_out
            '            End If

            'get the attachments
            '----------------------
            Dim cr_attach() As String = {}
            If attach_type Like "" Or attach_type Like "reex" Then
                cr_attach = {}

            ElseIf attach_type Like "resub" Then
                'for the resub case
                '----------------------
                'check if resub form is there, if not make it again
                '----------------------------------------------
                If FileIO.FileSystem.FileExists(local.base_path & local.cr & "\" & cr_id & "\Resubmit_Form_" & cr_id & ".xlsb") Then
                    cr_form = local.base_path & local.cr & "\" & cr_id & "\Resubmit_Form_" & cr_id & ".xlsb"
                    cr_attach = {}
                Else
                    Dim cr_resub_form As String = ""
                    pre_process_for_resub(cr_resub_form, cr_id, cr_form, local, err)
                    If Not err Like "" Then GoTo get_out
                    create_resubmit_form(cr_resub_form, cr_form_type, format, local, err)
                    If Not err Like "" Then GoTo get_out
                    cr_form = cr_resub_form
                    cr_attach = {}
                End If

            Else
                'for the regular case, using executor attachments or requester attachments
                '----------------------
                cr_attach = FileIO.FileSystem.GetFiles(Path.GetDirectoryName(cr_form), FileIO.SearchOption.SearchTopLevelOnly, attach_type & "*.zip").ToArray
                If cr_attach Is Nothing Then cr_attach = {}

            End If

            'get the email addresses
            '-------------------------------
            Dim cc_mask() As Boolean = {}
            If to_spec Like "requester" And attach_type Like "resub" Then : cc_mask = format.cc_mask_resubmission_request
            ElseIf to_spec Like "requester" And attach_type Like "executor" Then : cc_mask = format.cc_mask_review_request
            ElseIf to_spec Like "approver" Then : cc_mask = format.cc_mask_approval_request
            ElseIf to_spec Like "execution_coordinator" Then : cc_mask = format.cc_mask_execution_planning_request
            ElseIf to_spec Like "executors" Then : cc_mask = format.cc_mask_execution_request
            ElseIf to_spec Like "all" Then : cc_mask = format.cc_mask_everyone
            End If
            get_user_emails(cc_mask, user_list, requester, approver, execution_coordinator, executors, cc_list, cr_id, format, db, err)
            If Not err Like "" Then GoTo get_out
            If to_spec Like "requester" Then : If Not requester Like "" Then to_list = requester
            ElseIf to_spec Like "approver" Then : If Not approver Like "" Then to_list = approver
            ElseIf to_spec Like "execution_coordinator" Then : If Not execution_coordinator Like "" Then to_list = execution_coordinator
            ElseIf to_spec Like "executors" Then : If Not executors Like "" Then to_list = executors
            ElseIf to_spec Like "all" Then : to_list = user_list
            End If
            If to_list Like "" Then
                err = "ER: There is no 'to' email address"
                GoTo get_out
            End If
            get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
            If Not err Like "" Then GoTo get_out

            'update_log
            '--------------
            add2log(db, time_now, cr_id, log_entry_Text & " sent to " & to_list, err)
            If Not err Like "" Then GoTo get_out

            'resend the request
            '----------------------
            Dim cr_form_zipped As String = ""
            check_size_and_zip(cr_form, cr_form_zipped, format, local, err)
            If Not err Like "" Then GoTo get_out

            Dim log_s As String = ""
            body = body & get_log_string(log_s, db, cr_id, err)
            If cr_attach.Count = 0 Then
                auto_reply_general(tx_svr, to_list, cc_list_final, subj, body, {If(Not cr_form_zipped Like "", cr_form_zipped, cr_form)}, err)
                If Not err Like "" Then GoTo get_out
            Else
                Dim j As Integer = 0
                For Each attachment In cr_attach
                    If j = 0 Then
                        auto_reply_general(tx_svr, to_list, cc_list_final, subj & If(cr_attach.Count = 1, "", " => part " & j + 1 & " of " & cr_attach.Count), body, {If(Not cr_form_zipped Like "", cr_form_zipped, cr_form), attachment}, err)
                    Else
                        auto_reply_general(tx_svr, to_list, cc_list_final, subj & " => part " & j + 1 & " of " & cr_attach.Count, body, {attachment}, err)
                    End If
                    If Not err Like "" Then GoTo get_out
                    j += 1
                Next
            End If

            'update the status and add to the log
            '--------------------------------------
            update_cr_common_table(db, cr_id, "cr_status", target_status, err)
            If Not err Like "" Then GoTo get_out
            update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
            If Not err Like "" Then GoTo get_out
            update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)
            If Not err Like "" Then GoTo get_out

            If attach_type Like "resub" Then
                update_cr_common_table_date(db, cr_id, "approval_date", time_now, True, err)
                If Not err Like "" Then GoTo get_out
                update_cr_common_table_date(db, cr_id, "planned_execution_date", time_now, True, err)
                If Not err Like "" Then GoTo get_out
                update_cr_common_table_date(db, cr_id, "execution_date", time_now, True, err)
                If Not err Like "" Then GoTo get_out

            ElseIf attach_type Like "reex" Then
                update_cr_common_table_date(db, cr_id, "execution_date", time_now, True, err)
                If Not err Like "" Then GoTo get_out

            End If
            'need to do a check on the val for the next status, do we need to enter a date.time, say for pending execution, we need to check that planned execution date is filled etc.

get_out:
        Catch ex As Exception
            err = "ER: Error repairing the CR, details: " & ex.ToString
        End Try

        Try
            If Not err Like "" Then
                'update log
                '-----------------------
                add2log(db, time_now, cr_id, "CR cancelled by the server due to inconsistencies in the CR data.", err)

                If Not user_list Like "" Then
                    'cancel the CR
                    '-------------------
                    subj = "CRMS: CR Cancellation Note"
                    body = "The CR has been cancelled due to inconsistencies in the CR data, sorry for the inconvenience.  To continue please create the CR again.<BR>"
                    Dim log_s As String = ""
                    body = body & get_log_string(log_s, db, cr_id, err)
                    auto_reply_general(tx_svr, user_list, cc_list_final, subj, body, {""}, err)
                End If

                'update the status and add to the log
                '--------------------------------------
                update_cr_common_table(db, cr_id, "cr_status", "Cancelled", err)
                update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)

            End If
        Catch ex As Exception
        End Try
    End Sub



    'This finds the CRs if the given status who are over their lifetime from the timeref or if there is a last_nag_date value, then also over their nag date period from the last_nag_date
    Public Function make_time_sql(ByVal to_list As String, ByVal cr_status_regex As String, ByVal time_limit As String, ByVal time_limit_nag As String, ByVal time_ref As String, ByVal time_now As DateTime, ByVal db As mysql_server) As String
        On Error Resume Next
        Dim sql As New StringBuilder("")
        sql.Append("SELECT cr_id,")
        sql.Append(to_list & " as to_list,")
        sql.Append(time_ref & " as date_request_sent,")
        sql.Append("TIMESTAMPDIFF(HOUR," & time_ref & ",'" & time_now.ToString("yyyy-MM-dd HH:mm:ss") & "') - IF(" & time_limit & " IS NULL, (SELECT " & time_limit & " FROM " & db.schema & ".cr_types where cr_types.cr_type = cr_common.cr_type LIMIT 1)," & time_limit & ") as overdue ")
        sql.Append("FROM " & db.schema & ".cr_common ")
        sql.Append("WHERE cr_status REGEXP '" & cr_status_regex & "' ")
        sql.Append("AND " & time_ref & " IS NOT NULL ")
        sql.Append("AND TIMESTAMPDIFF(HOUR," & time_ref & ",'" & time_now.ToString("yyyy-MM-dd HH:mm:ss") & "') > IF(" & time_limit & " IS NULL, (SELECT " & time_limit & " FROM " & db.schema & ".cr_types where cr_types.cr_type = cr_common.cr_type LIMIT 1)," & time_limit & ") ")
        sql.Append("AND IF(last_nag_date IS NOT NULL, TIMESTAMPDIFF(HOUR,last_nag_date,'" & time_now.ToString("yyyy-MM-dd HH:mm:ss") & "') > IF(" & time_limit_nag & " IS NULL, (SELECT " & time_limit_nag & " FROM " & db.schema & ".cr_types where cr_types.cr_type = cr_common.cr_type LIMIT 1)," & time_limit_nag & "),last_nag_date IS NULL);")
        Return sql.ToString
    End Function




    Public Function find_crs_in_time_period(ByVal cr_status_regex As String, ByVal back_time As String, ByVal time_now As DateTime, ByVal db As mysql_server) As String
        On Error Resume Next
        Dim sql As New StringBuilder("")
        sql.Append("SELECT * FROM " & db.schema & ".cr_common ")
        sql.Append("WHERE cr_status REGEXP '" & cr_status_regex & "' ")
        '        sql.Append("AND open_date " & time_ref & " IS NOT NULL ")
        '       sql.Append("AND TIMESTAMPDIFF(HOUR," & time_ref & ",'" & time_now.ToString("yyyy-MM-dd HH:mm:ss") & "') > IF(" & time_limit & " IS NULL, (SELECT " & time_limit & " FROM " & db.schema & ".cr_types where cr_types.cr_type = cr_common.cr_type LIMIT 1)," & time_limit & ") ")
        '      sql.Append("AND IF(last_nag_date IS NOT NULL, TIMESTAMPDIFF(HOUR,last_nag_date,'" & time_now.ToString("yyyy-MM-dd HH:mm:ss") & "') > IF(" & time_limit_nag & " IS NULL, (SELECT " & time_limit_nag & " FROM " & db.schema & ".cr_types where cr_types.cr_type = cr_common.cr_type LIMIT 1)," & time_limit_nag & "),last_nag_date IS NULL);")
        Return sql.ToString
    End Function
End Module
