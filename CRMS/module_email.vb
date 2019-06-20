Option Explicit On
Option Compare Text
Module module_email


    'Rx email
    '----------------------------------
    Public Sub get_email_launcher(ByVal form As CRMS.main_form, ByVal local As local_machine, ByVal crms_control As crms_controller, ByVal format As cr_sheet_format, ByRef outer_err As String)
        'create the imap, smtp and mysql server objects
        '-----------------------------------------------
        Dim err As String = ""
        Dim done As Boolean = False
        Dim tx_svr As New smtp_server
        Dim rx_svr_imap As New imap_server
        Dim rx_svr_pop As New pop_server
        Dim db As New mysql_server
        Try
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
            '           Debug.AutoFlush = True
            '          rx_svr_imap.IsDebug = True

            rx_svr_pop.nat_server = form.TextBox_pop_server.Text
            rx_svr_pop.nat_port = Int(Val(form.TextBox_pop_port.Text))
            '            rx_svr_pop.nat_encrypt_type = Int(Val(form.TextBox_pop_encryption_type.Text))
            rx_svr_pop.nat_use_ssl = form.CheckBox_pop_use_ssl.Checked
            rx_svr_pop.nat_username = form.TextBox_pop_user.Text
            rx_svr_pop.nat_password = form.TextBox_pop_pass.Text

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

            'start work
            '---------------------------
            'connect to mysql DB
            '--------------------------
            debug.writeline(Now.ToLongTimeString & ": " & "Connecting to DB....")
            module_db.mysql_connect(db)

            'connect to imap svr
            '--------------------------
            If format.use_imap Then
                clean_discon_from_imap(rx_svr_imap)
                If Not rx_svr_imap.IsConnected And Not rx_svr_imap.IsAuthenticated Then
                    debug.writeline(Now.ToLongTimeString & ": " & "Connecting....")
                Else
                    debug.writeline(Now.ToLongTimeString & ": " & "Can't pre-Disconnect....")
                    Throw New Exception("Can't pre-disconnect from IMAP sever")
                End If

                clean_connect_to_imap(rx_svr_imap, local)
                If rx_svr_imap.IsConnected And rx_svr_imap.IsAuthenticated Then
                    debug.writeline(Now.ToLongTimeString & ": " & "Checking Mail....")

                    'this analyses and processes emails from the server
                    '----------------------------------
                    module_crms_logic.get_email(db, rx_svr_imap, rx_svr_pop, tx_svr, local, crms_control, format, err)
                    If err Like "" Then done = True

                ElseIf rx_svr_imap.IsConnected Then
                    debug.writeline(Now.ToLongTimeString & ": " & "Couldn't Login")
                    Throw New Exception("Can't login to IMAP sever")
                Else
                    debug.writeline(Now.ToLongTimeString & ": " & "Couldn't Connect")
                    Throw New Exception("Can't connect to IMAP sever")
                End If

            Else
                'this analyses and processes emails from the server
                'for pop we have to discon and recon each time we want to scan the inbox, so no point doing it here, do it just before you check
                '----------------------------------
                module_crms_logic.get_email(db, rx_svr_imap, rx_svr_pop, tx_svr, local, crms_control, format, err)
                If err Like "" Then done = True
            End If

        Catch ex As Exception
            If Not done Then outer_err = "Error in the get email launcher sub, details: " & ex.ToString

        Finally
            debug.writeline(Now.ToLongTimeString & ": " & "Disconnecting from DB....")
            module_db.mysql_disconnect(db)

            If format.use_imap Then
                If rx_svr_imap.IsConnected Then
                    debug.writeline(Now.ToLongTimeString & ": " & "Disconnecting....")
                    clean_discon_from_imap(rx_svr_imap)
                    If Not rx_svr_imap.IsConnected And Not rx_svr_imap.IsAuthenticated Then
                        debug.writeline(Now.ToLongTimeString & ": " & "Idle")
                    Else
                        debug.writeline(Now.ToLongTimeString & ": " & "Couldn't Disconnect")
                    End If
                End If
            Else
                If rx_svr_pop.Connected Then
                    clean_discon_from_pop(rx_svr_pop)
                    If Not rx_svr_pop.Connected Then
                        debug.writeline(Now.ToLongTimeString & ": " & "Idle")
                    Else
                        debug.writeline(Now.ToLongTimeString & ": " & "Couldn't Disconnect")
                    End If
                End If
            End If
            tx_svr.Dispose()
            rx_svr_imap.Dispose()
            rx_svr_pop.Dispose()
            tx_svr = Nothing
            rx_svr_imap = Nothing
            rx_svr_pop = Nothing
            db = Nothing
            GC.Collect()
        End Try
    End Sub




    'email server connect and disconnect
    '------------------------------------------
    Public Sub clean_connect_to_pop(ByRef rx_svr_pop As pop_server, ByVal local As local_machine)
        On Error Resume Next
        local.check_directories(local)
        rx_svr_pop.Connect(rx_svr_pop.nat_server, rx_svr_pop.nat_port, rx_svr_pop.nat_use_ssl)
        If rx_svr_pop.Connected Then
            rx_svr_pop.Authenticate(rx_svr_pop.nat_username, rx_svr_pop.nat_password)
        End If
    End Sub

    Public Sub clean_discon_from_pop(ByRef rx_svr_pop As pop_server)
        On Error Resume Next
        If rx_svr_pop.Connected Then
            rx_svr_pop.Disconnect()
        End If
    End Sub



    'email server connect and disconnect
    '------------------------------------------
    Public Sub clean_connect_to_imap(ByRef rx_svr_imap As imap_server, ByVal local As local_machine)
        On Error Resume Next
        local.check_directories(local)
        If rx_svr_imap.Connect(rx_svr_imap.Host, rx_svr_imap.Port, rx_svr_imap.UseSsl) Then
            rx_svr_imap.Login(rx_svr_imap.nat_cred_regular)
        End If
    End Sub

    Public Sub clean_discon_from_imap(ByRef rx_svr_imap As imap_server)
        On Error Resume Next
        If rx_svr_imap.IsConnected Then
            rx_svr_imap.Logout()
            rx_svr_imap.Disconnect()
        End If
    End Sub




    'tx_mail
    Public Sub auto_reply_general(ByVal tx_svr As smtp_server, ByVal to_addr As String, ByVal cc_addr As String, ByVal subject As String, ByVal msg As String, ByVal attach() As String, ByRef err As String)
        Dim retry_counter As Integer = 0
retry:
        err = ""
        send_email_sync(tx_svr, subject, msg, tx_svr.nat_username, "CRMS Huawei-Telkomsel", to_addr, cc_addr, attach, err)
        If Not err Like "" Then
            retry_counter += 1
            debug.writeline(Now.ToLongTimeString & ": " & "Network issues are preventing email sending, retrying 10x....")
            System.Threading.Thread.Sleep(tx_svr.email_retry_backoff * 1000)
            If retry_counter < 10 Then GoTo retry
        End If
    End Sub




    Public Sub send_email_sync(ByVal tx_svr As smtp_server, ByVal msg_sub As String, ByVal msg_body As String, ByVal from_addr As String, ByVal from_name As String, ByVal to_addr As String, ByVal cc_addr As String, ByVal attch() As String, ByRef err As String)
        debug.writeline(Now.ToLongTimeString & ": " & "Sending Email Start....")
        Dim from_mail_addr As New System.Net.Mail.MailAddress(from_addr, from_name, System.Text.Encoding.UTF8)
        Dim message As New System.Net.Mail.MailMessage()
        Try
            message.Subject = msg_sub
            message.SubjectEncoding = System.Text.Encoding.UTF8
            message.Body = msg_body
            message.BodyEncoding = System.Text.Encoding.UTF8
            message.From = from_mail_addr
            message.To.Add(to_addr)
            message.IsBodyHtml = True
            If Not cc_addr Like "" Then message.CC.Add(cc_addr)

            'add all attachments if any exist
            '----------------------------------
            Try
                If Not attch Is Nothing AndAlso Not attch.Length = 0 Then
                    For Each item In attch
                        If IO.File.Exists(item) Then
                            Dim retry_cnt As Integer = 0
                            Try
retry:
                                Dim attachment As New System.Net.Mail.Attachment(item)
                                message.Attachments.Add(attachment)
                            Catch ex As System.IO.IOException
                                retry_cnt = retry_cnt + 1
                                If retry_cnt < 10 Then
                                    debug.writeline(Now.ToLongTimeString & ": " & "Having issues attaching a file, trying 10x then giving up.")
                                    Module_misc.kill_proc(item, err)
                                    System.Threading.Thread.Sleep(1000)
                                    GoTo retry
                                End If
                                debug.writeline(Now.ToLongTimeString & ": " & "giving up......")
                                err = "ER: We couldn't add an attachment, tried killing processes, but failed, details: " & err
                                GoTo get_out
                            End Try
                        End If
                    Next
                End If
            Catch ex As Exception
                err = "ER: Could not attach files to email, giving up: " & attch(0)
                GoTo get_out
            End Try

            'send the msg
            '-----------------
            Try
                tx_svr.Send(message)
            Catch ex As Exception
                err = "ER: Could not send the email, looks like a server connection issue, details: " & ex.ToString
                GoTo get_out
            End Try
get_out:
        Catch ex As Exception
            If err = "" Then
                err = "ER: Genreral error sending email, details: " & ex.ToString
            End If
        Finally
            debug.writeline(Now.ToLongTimeString & ": " & "Sending Email Stop....")
            message.Dispose()
        End Try
    End Sub







    'this will dl and unpack all attachments for IMAP
    '-------------------------------------------------
    Public Sub dl_attachments_imap(ByVal msg As ImapX.Message, ByVal unzip_attch_flag As Boolean, ByVal local As local_machine, ByVal tx_svr As smtp_server, ByRef err As String)
        debug.writeline(Now.ToLongTimeString & ": " & "Downloading Attachments Start....")
        Dim i As Integer = 0
        For Each item In msg.Attachments
            Dim filename As String = Path.GetFileNameWithoutExtension(item.FileName)
            Dim ext As String = Path.GetExtension(item.FileName)
            Dim final_filename As String = ""
            Try
                'this dl and saves the attachment, too critical processes, so they need the retry algorithm
                '---------------------------------------------------------
                Try
                    item.GetTextData()
                    item.Download()
                Catch ex As Exception
                    err = "XREJ: Can't DL attachment from email server, giving up..."
                    GoTo get_out
                End Try

                Try
                    'tests if the attachment file already exists, if so, there is an issue in the total attachment list, where the user has attached duplicate file names
                    'maybe one was inside a zip file and the other wasn't
                    '---------------------------------------------------------------------
                    Dim retry_cnt As Integer = 0
retry:
                    final_filename = filename & If(retry_cnt = 0, "", "(" & retry_cnt & ")") & ext
                    If IO.File.Exists(local.base_path & local.inbox & "\" & final_filename) Then
                        retry_cnt += 1
                        If retry_cnt < 100 Then
                            GoTo retry
                        End If
                        err = "XREJ: Some of your attachments have duplicate filenames, can't process, giving up."
                        GoTo get_out
                    End If
                    item.Save(local.base_path & local.inbox, final_filename)
                Catch ex As Exception
                    err = "XREJ: Can't save an attachment to the HDD...."
                    GoTo get_out
                End Try

                'this unzips/rars the attachment if it is zipped/rarred, unzips to whatever directory structure is in the zip and then deletes the original zip file
                '------------------------------------------------------------------------------------------------------------------------------------------
                If unzip_attch_flag Then
                    unpack_attachments(final_filename, local, err)
                    If Not err Like "" Then GoTo get_out
                End If
do_not_unzip:
            Catch ex As Exception
                If err = "" Then
                    err = "XREJ: There was a general error downloading mail attachments, giving up, details: " & ex.ToString
                End If
                GoTo get_out
            End Try
        Next
get_out:
        debug.writeline(Now.ToLongTimeString & ": " & "Downloading Attachments Stop....")
    End Sub





    'DL and unpack attachments for the POP case
    '---------------------------------------------
    Public Sub dl_attachments_pop(ByVal msg As OpenPop.Mime.Message, ByRef msgx As CRMS.email_msg, ByVal unzip_attch_flag As Boolean, ByVal local As local_machine, ByVal tx_svr As smtp_server, ByRef err As String)
        debug.writeline(Now.ToLongTimeString & ": " & "Downloading Attachments Start....")
        Dim i As Integer = 0
        'this writes the attachments to the HDD
        '-----------------------------------------
        Dim query = From item As MessagePart In msg.FindAllAttachments
                    Where item.IsAttachment And Not item.ContentDisposition.Inline
                    Select item
        For Each item As MessagePart In query
            Dim filename As String = item.FileName
            Dim ext As String = Path.GetExtension(item.FileName)
            Dim final_filename As String = local.base_path & local.inbox & "\" & item.FileName
            Try
                'this dl and saves the attachment, too critical processes, so they need the retry algorithm
                '---------------------------------------------------------
                System.IO.File.WriteAllBytes(final_filename, item.Body)

                'fills the attachments array
                '-------------------------------
                Dim ts As New FileInfo(final_filename)
                Dim temp As New email_msg.attachment(final_filename, Val(ts.Length / 1000000.0))
                Dim count As Integer = msgx.attachments.Count + 1
                ReDim Preserve msgx.attachments(count)
                msgx.attachments(count - 1) = temp

                'this unzips/rars the attachment if it is zipped/rarred, unzips to whatever directory structure is in the zip and then deletes the original zip file
                '------------------------------------------------------------------------------------------------------------------------------------------
                If unzip_attch_flag Then
                    unpack_attachments(filename, local, err)
                    If Not err Like "" Then GoTo get_out
                End If
do_not_unzip:
            Catch ex As Exception
                If err = "" Then
                    err = "XREJ: There was a general error downloading mail attachments, giving up, details: " & ex.ToString
                End If
                GoTo get_out
            End Try
        Next
get_out:
        debug.writeline(Now.ToLongTimeString & ": " & "Downloading Attachments Stop....")
    End Sub




    'this does the unzipping etc.
    Public Sub unpack_attachments(ByVal final_filename As String, ByVal local As local_machine, ByRef err As String)
        Try
            Dim temp_file As String = local.base_path & local.inbox & "\" & final_filename
            Dim temp_dir As String = local.base_path & local.inbox
            If Regex.IsMatch(final_filename, "\.zip$", RegexOptions.IgnoreCase) Then
                'tests if we had an error unzipping
                module_zip.unzip_file(temp_file, temp_dir, err)
                'I disable err handling  for this, as I do not want to exit if it failed
                err = ""
                '                    If Not err Like "" Then
                'I do not exit on unzip fail, we just will not process the zip file attachments, but I have it stable now anyway,
                'I unzip and if I see a file that already exists, I rename the new one
                '                        Regex.Replace(err, "^ER", "XREJ")
                '                        GoTo get_out
                '                End If

                'deletes the original zip file
                '------------------------------
                force_delete_file(temp_file, err)
                If Not err Like "" Then GoTo get_out

            ElseIf Regex.IsMatch(final_filename, "\.rar$", RegexOptions.IgnoreCase) Then
                'tests if we had an error unipping
                module_zip.unrar_file(temp_file, temp_dir, err)
                'I disable err handling  for this, as I do not want to exit if it failed
                err = ""
                '                    If Not err Like "" Then
                'I do not exit on unzip fail, we just will not process the zip file attachments, but I have it stable now anyway,
                'I unzip and if I see a file that already exists, I rename the new one
                '                        Regex.Replace(err, "^ER", "XREJ")
                '                        GoTo get_out
                '                End If

                'deletes the original rar file
                '------------------------------
                force_delete_file(temp_file, err)
                If Not err Like "" Then GoTo get_out
            End If
get_out:
        Catch ex As Exception
            err = "ER: Error unpacking attachments, details: " & ex.ToString
        End Try
    End Sub


End Module
