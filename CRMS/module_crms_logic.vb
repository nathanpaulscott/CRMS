Option Explicit On
Option Compare Text

'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------

'later
'updating engpar
'need to hold engpar and update it on cr closure
'need to handle planning emails to add planned sites, change site status to on-air/cancelled
'change site info.
'then when I update engpar data with a closed CR, if I can't find the site in engpar or the status is not on-air, I can send a warning to planning about it.
'do more input checks to make sure 3G and cr_objective are from allowed list (check as a pair!!)  Also have this for ex.coord and cr_type



Module module_crms_logic
    Public Sub get_email(ByVal db As mysql_server, ByVal rx_svr_imap As imap_server, ByVal rx_svr_pop As pop_server, ByVal tx_svr As smtp_server, ByVal local As local_machine, ByVal logic_control As crms_controller, ByVal format As cr_sheet_format, ByRef err As String)
        Try
            'set the current thread culture to en-US to override the server settings
            '---------------------------------------------------
            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", False)
            Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US", False)

            err = ""

            'Update all the allowed values dataset, need to do this before each mailserver connect and dl
            '-------------------------------------------------------------------------------------------
            allowed_values_db2ds(db, format, err)
            If Not err Like "" Then
                err = "ER: We couldn't update the allowed value tables from the DB, can't continue, giving up..."
                GoTo get_out
            End If

            'sends the opening mail
            '------------------------
            If format.send_opening_mail Then
                format.send_opening_mail = False
                Dim subj As String = ""
                Dim body As String = ""
                Dim attachment As String = If(FileIO.FileSystem.FileExists(local.base_path & "\restart_info.csv"), local.base_path & "\restart_info.csv", "")
                subj = "CRMS: Startup Email"
                body = "The app has just started up on: " & Now.ToLongDateString & " : " & Now.ToLongTimeString
                body = body & "<P>Thanks<P>"
                auto_reply_general(tx_svr, format.dev_email, "", subj, body, {attachment}, err)
                If Not attachment Like "" Then
                    FileIO.FileSystem.DeleteFile(local.base_path & "\restart_info.csv")
                End If
                If Not err Like "" Then GoTo get_out
            End If

            'this queries the msgs to download the attachments
            '---------------------------------------------
            Do
                'this is my msg class to hold all the msg data
                '----------------------------------
                Dim msgx As New CRMS.email_msg

                'we connect and get a msg here
                '------------------------------
                If Not format.use_imap Then
                    Try
                        '#######################################################3
                        'POP
                        '#######################################################3
                        Debug.WriteLine(Now.ToLongTimeString & ": " & "Connecting...")

                        Dim process_flag As Boolean = True
                        clean_discon_from_pop(rx_svr_pop)
                        clean_connect_to_pop(rx_svr_pop, local)
                        Dim count As Integer = rx_svr_pop.GetMessageCount
                        If count = 0 Then Exit Do

                        'first cleanup the inbox folder
                        '-------------------------------------
                        clean_dir(local.base_path & local.inbox, err)
                        If Not err Like "" Then Exit Do

                        '                For cnt = 1 To count
                        '               rx_svr_pop.DeleteMessage(cnt)
                        '              Next  
                        'the oldest is always 1, then it goes up from there, 0 doesn't exist

                        Debug.WriteLine(Now.ToLongTimeString & ": " & "Doing oldest msg of " & count & " on the server...")

                        Dim msgpop As OpenPop.Mime.Message
                        Try
                            msgpop = rx_svr_pop.GetMessage(1)
                            msgx.msg_date = msgpop.Headers.Date
                            Dim t_from = New email_msg.mail_spec(msgpop.Headers.From.DisplayName, msgpop.Headers.From.Address)
                            msgx.from = t_from
                            msgx.to_list = (From item In msgpop.Headers.To
                                            Let a = New email_msg.mail_spec(item.DisplayName, item.Address)
                                            Where item.HasValidMailAddress And Not item.Address Like tx_svr.nat_username
                                            Select a).Distinct.ToArray
                            msgx.cc_list = (From item In msgpop.Headers.Cc
                                            Let a = New email_msg.mail_spec(item.DisplayName, item.Address)
                                            Where item.HasValidMailAddress And Not item.Address Like tx_svr.nat_username
                                            Select a).Distinct.ToArray
                            msgx.subject_raw = msgpop.Headers.Subject

                            'read the bodies, can crash, so need to protect
                            '-------------------------------------------
                            msgx.body_text_raw = ""
                            Try
                                msgx.body_text_raw = msgpop.FindFirstPlainTextVersion.GetBodyAsText
                            Catch ex As Exception
                                Debug.WriteLine(Now.ToLongTimeString & ": " & "No text in the body error, no big deal...")
                            End Try
                            msgx.body_html_raw = ""
                            Try
                                msgx.body_html_raw = msgpop.FindFirstHtmlVersion.GetBodyAsText
                            Catch ex As Exception
                                Debug.WriteLine(Now.ToLongTimeString & ": " & "No text in the body error, no big deal...")
                            End Try

                            'dl to attachments to the HDD and fill the attachments array
                            '--------------------------------------------------
                            msgx.attachments = {}
                            dl_attachments_pop(msgpop, msgx, True, local, tx_svr, err)
                            If Not err Like "" Then Throw New Exception
                            msgx.attachments = (From item As email_msg.attachment In msgx.attachments
                                                Where Not item Is Nothing
                                                Select item).ToArray

                        Catch ex As Exception
                            process_flag = False
                        End Try

                        Try
                            rx_svr_pop.DeleteMessage(1)
                        Catch ex As Exception
                        End Try
                        clean_discon_from_pop(rx_svr_pop)

                        If process_flag Then process_mail(msgx, db, tx_svr, local, format, err)

                    Catch ex As Exception
                        err = "ER: Error connecting and downloading POP msgs, details: " & ex.ToString
                        GoTo get_out
                    End Try

                Else
                    '#######################################################3
                    'IMAP
                    '#######################################################3
                    If Not (rx_svr_imap.IsConnected And rx_svr_imap.IsAuthenticated) Then
                        clean_discon_from_imap(rx_svr_imap)
                        clean_connect_to_imap(rx_svr_imap, local)
                        If Not (rx_svr_imap.IsConnected And rx_svr_imap.IsAuthenticated) Then Exit Do
                    End If
                    rx_svr_imap.Folders.Inbox.Messages.Download("ALL", MessageFetchMode.Minimal)
                    Dim query = From item As ImapX.Message In rx_svr_imap.Folders.Inbox.Messages
                                Where Not item.Seen
                                Order By item.Date
                                Select item
                    If query.Count = 0 Then Exit Do
                    Dim msg_tot As Integer = query.Count
                    Dim msg As ImapX.Message = query.First

                    'first cleanup the inbox folder
                    '-------------------------------------
                    clean_dir(local.base_path & local.inbox, err)
                    If Not err Like "" Then Exit Do

                    'keep the status as unseen
                    '----------------------------
                    Debug.WriteLine(Now.ToLongTimeString & ": " & "Doing msg: 1 of " & msg_tot)

                    'here we simply analyse the subject and decide what to do with the mail, if process, we process, after this sub we simply delete the mail
                    '-----------------------------------------------------------
                    msg.Seen = True

                    'fill values in my email object
                    '---------------------------------
                    msgx.msg_date = msg.Date.Value
                    msgx.subject_raw = msg.Subject
                    msgx.body_text_raw = msg.Body.Text
                    msgx.body_html_raw = msg.Body.Html
                    msgx.from = New email_msg.mail_spec(msg.From.DisplayName, msg.From.Address)
                    msgx.to_list = (From item In msg.To
                                    Let a = New email_msg.mail_spec(item.DisplayName, item.Address)
                                    Where format.IsValidEmail(item.Address) And Not item.Address Like tx_svr.nat_username
                                    Select a).ToArray
                    msgx.cc_list = (From item In msg.Cc
                                    Let a = New email_msg.mail_spec(item.DisplayName, item.Address)
                                    Where format.IsValidEmail(item.Address) And Not item.Address Like tx_svr.nat_username
                                    Select a).ToArray
                    msgx.attachments = (From item In msg.Attachments
                                        Let a = New email_msg.attachment(local.base_path & local.inbox & "\" & item.FileName, item.FileSize)
                                        Where Not item.FileName Like ""
                                        Select a).ToArray

                    dl_attachments_imap(msg, True, local, tx_svr, err)
                    If Not err Like "" Then    'there was an error with the attachments attachments, so reject the CR
                        GoTo get_out
                    End If

                    process_mail(msgx, db, tx_svr, local, format, err)

                    'remove seen emails
                    '-------------------------
                    If Not (rx_svr_imap.IsConnected And rx_svr_imap.IsAuthenticated) Then
                        clean_discon_from_imap(rx_svr_imap)
                        clean_connect_to_imap(rx_svr_imap, local)
                        If Not (rx_svr_imap.IsConnected And rx_svr_imap.IsAuthenticated) Then GoTo skip_remove_seen
                        rx_svr_imap.Folders.Inbox.Messages.Download("ALL")
                    End If
                    Dim query2 = From item As ImapX.Message In rx_svr_imap.Folders.Inbox.Messages
                                    Where item.Seen
                                    Select item
                    Dim i As Integer = query2.Count
                    Dim j As Integer = 0
                    If i > 0 Then
                        For j = 1 To i
                            query2(0).Remove()
                        Next
                    End If
skip_remove_seen:
                End If

                'we only exit on a system error, for rejected CRs, we aleady sent an autoreply and we still process messages
                '----------------------------------------
                If Regex.IsMatch(err, "^([A-Z]+)(REJ:)(\sSee\sreturned\sCR\sform\s)(\()((.*\(.*\))|(.*))(\))(\()(.*)(\))$") Then
                    Dim new_subj As String = Regex.Replace(err, "^([A-Z]+)(REJ:)(\sSee\sreturned\sCR\sform\s)(\()", "", RegexOptions.IgnoreCase)
                    new_subj = Trim(Regex.Replace(new_subj, "\)\(.*$", "", RegexOptions.IgnoreCase))
                    Dim file_out As String = Regex.Replace(err, "^.*\)\(", "", RegexOptions.IgnoreCase)
                    file_out = Trim(Regex.Replace(file_out, "\)$", "", RegexOptions.IgnoreCase))
                    If new_subj Like "" Or Not FileIO.FileSystem.FileExists(file_out) Then
                        err = "ER: Bad data in CR rejection string [" & err & "]"
                        GoTo report_error
                    End If

                    Dim cr_form_zipped As String = ""
                    err = ""
                    check_size_and_zip(file_out, cr_form_zipped, format, local, err)
                    If Not err Like "" Then GoTo get_out

                    Dim subj As String = new_subj
                    Dim body As String = "Dear " & msgx.from.displayname & ",<P>There were errors with the data you gave in your CR form, please see returned CR error form.<BR>To continue you should:" & _
                    array2html_bullet_list({"Open the attachment", "Fix the highlighted errors", "Save the file, any name is ok", "Reply to this mail, attaching the new file"})
                    body = body & "<P>.<P>On " & msgx.msg_date.ToLongDateString & " at " & msgx.msg_date.ToLongTimeString & ", " & msgx.from.displayname & " &lt;" & msgx.from.address & "&gt; wrote:<BR>" '& If(msgx.Body.HasHtml, "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msgx.Body.Html & "</div>", "")
                    err = ""
                    auto_reply_general(tx_svr, msgx.from.address, "", subj, body, {If(Not cr_form_zipped Like "", cr_form_zipped, file_out)}, err)

                ElseIf Regex.IsMatch(err, "^EMAILREJ:\s") Then
                    Dim subj As String = "CRMS: Server-Email Connectivity Issues"
                    Dim body As String = "Dear " & msgx.from.displayname & ",<P>There were email server connectivity issues preventing your mail being processing:<BR>Error Message => " & err & "<BR>"
                    body = Regex.Replace(body, "^EMAILREJ:\s", "", RegexOptions.IgnoreCase)
                    err = ""
                    body = body & "<P>.<P>On " & msgx.msg_date.ToLongDateString & " at " & msgx.msg_date.ToLongTimeString & ", " & msgx.from.displayname & " &lt;" & msgx.from.address & "&gt; wrote:<BR>" '& If(msgx.Body.HasHtml, "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msgx.Body.Html & "</div>", "")
                    auto_reply_general(tx_svr, msgx.from.address, "", subj, body, {}, err)

                ElseIf Regex.IsMatch(err, "^([A-Z]+)(REJ:\s)") Then
                    Dim x As String = Regex.Replace(err, "^([A-Z]+)(REJ:\s)", "", RegexOptions.IgnoreCase)
                    Dim subj As String = "CRMS: Your Mail Had Issues"
                    Dim body As String = "Dear " & msgx.from.displayname & ",<P>Your email was not processed because there were issues with it.<BR>The issues are described as following:<BR>Error Message => " & x & "<BR>"
                    body = body & "<P>.<P>On " & msgx.msg_date.ToLongDateString & " at " & msgx.msg_date.ToLongTimeString & ", " & msgx.from.displayname & " &lt;" & msgx.from.address & "&gt; wrote:<BR>" '& If(msgx.Body.HasHtml, "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msgx.Body.Html & "</div>", "")
                    err = ""
                    auto_reply_general(tx_svr, msgx.from.address, "", subj, body, {}, err)

                ElseIf Not err Like "" Then                'If Regex.IsMatch(err, "^ER:\s") Then
report_error:
                    'to the admin
                    '-------------------
                    Dim x As String = Trim(Regex.Replace(err, "^ER:\s", "", RegexOptions.IgnoreCase))
                    Dim subj As String = "CRMS: Internal Error Warning" & " (Subj: " & msgx.subject_raw & ", From: " & msgx.from.address & ")"
                    Dim body As String = "Some error occurred:<BR>"
                    body = body & x & "<BR>"
                    body = body & "<P>.<P>On " & msgx.msg_date.ToLongDateString & " at " & msgx.msg_date.ToLongTimeString & ", " & msgx.from.displayname & " &lt;" & msgx.from.address & "&gt; wrote:<BR>" '& If(msgx.Body.HasHtml, "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msgx.Body.Html & "</div>", "")
                    err = ""
                    auto_reply_general(tx_svr, format.dev_email, "", subj, body, {}, err)

                    'to the user
                    '-----------------
                    If format.send_internal_errors_to_user Then
                        subj = "CRMS: Mail Not Processed" & " (" & msgx.subject_raw & ")"
                        body = "Dear " & msgx.from.displayname & ",<P>Your email was not processed because of an internal error, details are as follows:<BR>" & x & "<BR>An error report has been generated and forwarded to the system administrtor for analysis."
                        body = body & "<P>.<P>On " & msgx.msg_date.ToLongDateString & " at " & msgx.msg_date.ToLongTimeString & ", " & msgx.from.displayname & " &lt;" & msgx.from.address & "&gt; wrote:<BR>" ' & If(msgx.Body.HasHtml, "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msgx.Body.Html & "</div>", "")
                        err = ""
                        auto_reply_general(tx_svr, msgx.from.address, "", subj, body, {}, err)
                    End If
                End If
            Loop
get_out:

            'this deletes the downloaded msgs to keep the mailbox small
            '-------------------------------------------------
            '            Try
            '            Dim i As Integer = 0
            '           Dim j As Integer = 0
            '          Dim msg_query = From msg In rx_svr_imap.Folders.Inbox.Messages
            '            Where msg.Seen
            '                   Select msg
            '            i = msg_query.Count
            '           If i > 0 Then
            '            For j = 1 To i
            '            msg_query(0).Remove()
            '           Next
            '            End If
            '       Catch ex As Exception
            '            err = "ER: error clearing inbox: " & ex.ToString
            '       End Try
        Catch ex As Exception
            err = "ER: Error in the get email sub: " & ex.ToString
        End Try
    End Sub




    'determines what to do with the incoming emails, mainly analyses the subject for key text codes, but also checks for attachments
    '------------------------------------------------------------------------------------------------------------------------------------
    Public Sub process_mail(ByRef msg As CRMS.email_msg, ByVal db As mysql_server, ByVal tx_svr As smtp_server, ByVal local As local_machine, ByVal format As cr_sheet_format, ByRef err As String)
        Try
            'this decides what to do with the emails
            '--------------------------------------------------------
            Dim time_now As DateTime = Now
            Dim msg_signature As String = ""
            Dim dt As New System.Data.DataTable
            Dim cleansubject As String = Trim(Regex.Replace(msg.subject_raw, "[^-A-Za-z0-9\s_\.:\(\)\n\r\+\\/\[\]';,<>\?\{\}=@!#\$%\^&\*""~`]", ""))
            Dim cc_list As String = ""
            For Each item In msg.cc_list
                If Len(item.address) > 0 And format.IsValidEmail(item.address) Then
                    cc_list = cc_list & item.address & ","
                End If
            Next
            If Len(cc_list) > 0 Then cc_list = Left(cc_list, Len(cc_list) - 1)

            'does the reading of the body from text, complicated
            '--------------------------------------------------
            Dim cleanbody As String = ""
            Try
                cleanbody = msg.body_text_raw
                cleanbody = Trim(Regex.Replace(cleanbody, "[^-A-Za-z0-9\s_\.:\(\)\n\r\+\\/\[\]';,<>\?\{\}=@!#\$%\^&\*""~`]", ""))
            Catch ex As Exception
                cleanbody = ""
            End Try

            'does the logic processing here
            '-----------------------------------------------------------------
            If Regex.IsMatch(cleansubject, "^((new)|(cancel))$", RegexOptions.IgnoreCase) Or Regex.IsMatch(cleansubject, "^CRMS:\sNew\sCR\sRetry\sRequest$", RegexOptions.IgnoreCase) _
            Or (Regex.IsMatch(cleansubject, "^(.*)(\(.*\))(.*)$", RegexOptions.IgnoreCase) And Regex.IsMatch(cleanbody, "^cancel", RegexOptions.IgnoreCase)) _
            Then
                'we have a cr related email
                'Now we test for the body contents
                If Regex.IsMatch(cleansubject, "^new$", RegexOptions.IgnoreCase) And msg.attachments.Count = 0 Then
                    'respond to the emailer with a blank CR creation form => Q: to test for allowed email addresses or not
                    '----------------------------------------------------------------------------------------------
                    'Uppdates the allowed values in the blank form as the DB may have changed
                    '-------------------------------------------------------------------
                    allowed_values_db2ds(db, format, err)
                    If Not err Like "" Then GoTo get_out
                    allowed_values_ds2xl(local.base_path & local.cr_blank_request_form_dir & "\" & local.cr_form, "init", format, local, err)
                    If Not err Like "" Then GoTo get_out

                    Dim subj As String = "CRMS: CR Creation Form"
                    Dim body As String = "Dear " & msg.from.displayname & ",<P>Please find the Blank CR Form attached.<BR>To open a new CR, please:"
                    body = body & array2html_bullet_list({"Add your data to the form", "Save the file, name is up to you", "Attach the file to a mail with the subject = 'new'", "Send to " & tx_svr.nat_username})
                    body = body & "<P>Thanks<P>"
                    err = ""
                    body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                    auto_reply_general(tx_svr, msg.from.address, "", subj, body, {local.base_path & local.cr_blank_request_form_dir & "\" & local.cr_form}, err)
                    If Not err Like "" Then
                        GoTo get_out
                    End If

                ElseIf Regex.IsMatch(cleansubject, "^((new)|(CRMS:\sNew\sCR\sRetry\sRequest))$", RegexOptions.IgnoreCase) And msg.attachments.Count > 0 Then
                    '######################################################################################33
                    'procedure for the repeated CR check
                    '-------------------------------------
                    'if the requester sends the same email several times, the tool will not know that these new CR forms are repeated, so it has to have a way to distinguish.
                    'check the duplicate field in the cr_common table, if match, then reject CR....can check for identical CRs to stop doubling up => check for attachment file size and other file attributes as well as requester email.
                    'check for partical matches in CR content from the same requester => if the CR changes the same aspect of too many nodes that are in otehr open CRs, give a warning to the requester and only proceed if they respond with ok/proceed/lanjut etc...
                    'for partial matches, param CR =>  check node_paramloc_mo_param, for phy cr check cr_type_node.  Check against other open CRs from this requester.  Can calc % of CR_sub_IDs in this new request that are similar to other open CRs
                    'if more than 20-30 or whatever %, go through the warning procedure with the requester.


                    'need to  rfb and non-rfb and create a separate cr for each, but if they came from the same CR creation form, then the CR_ID will be the same except for a "p" at the end
                    '1) do check that it is from an allowed originator
                    '--------------------------------
                    Dim s_from As String = msg.from.address
                    Dim qrows = From row In format.ds_allow.Tables("requesters").AsEnumerable
                                Where row("email") Like s_from
                                Select row
                    If qrows.Count = 0 Then
                        err = "CRREJ: New CR Rejection.  Your email does not have CR requesting rights.<BR>Thanks"
                        GoTo get_out
                    End If

                    '2) basic attachement check
                    '----------------------------
                    If msg.attachments.Count = 0 Then
                        err = "CRREJ: Your email doesn't have any attachments.<BR>Thanks"
                        GoTo get_out
                    End If
                    check_attachments(msg, err)
                    If Regex.IsMatch(err, "^er:", RegexOptions.IgnoreCase) Then
                        GoTo get_out
                    ElseIf Not err Like "" Then
                        err = "CRREJ: New CR Rejection.  " & err
                        GoTo get_out
                    End If

                    '2.1) check the the msg is not a resend duplicate, maybe the requester had an itchy trigger finger or something
                    'if the user has used new force, then we do not check for duplicates and add it regardless
                    '-----------------------------------------------------------------------------------------
                    Dim msg_sig_cr_id() As String = {}
                    msg_signature = msg.from.address & "."
                    For Each item In msg.attachments
                        msg_signature = msg_signature & item.filename & "." & item.size & "."
                    Next
                    If Len(msg_signature) > 0 Then msg_signature = Left(msg_signature, Len(msg_signature) - 1)

                    If Regex.IsMatch(cleansubject, "^new$", RegexOptions.IgnoreCase) Then
                        Try
                            clear_dt(dt)
                            sqlquery(False, db, "select cr_id, cr_status from " & db.schema & ".cr_common where msg_signature LIKE '" & msg_signature & "' AND cr_status NOT REGEXP '^((closed)|(cancelled))$';", dt, err)
                            If Not err Like "" Then GoTo get_out
                            If dt.Rows.Count > 0 Then
                                'this filters values from a col of a datatable and puts it in an array, linq is soo good.
                                Dim q_temp = From row In dt.AsEnumerable()
                                                Let a = local.base_path & local.cr & "\" & row.Item("cr_id").ToString & "\" & row.Item("cr_id").ToString & ".xlsb"
                                                Let b = row.Item("cr_id").ToString
                                                Where FileIO.FileSystem.FileExists(a)
                                                Select b
                                msg_sig_cr_id = q_temp.ToArray
                            End If
                        Catch ex As Exception
                            err = "ER: an error occurred checking the msg_signature, details: " & ex.ToString
                            GoTo get_out
                        End Try
                    End If

                    '3.2) Find the cr form
                    '--------------------------------
                    Dim raw_cr_form As String = ""
                    Dim input_file As String = ""
                    If FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.xlsb").Count = 0 Then
                        err = "CRREJ: New CR Rejection.  Your email doesn't have any acceptable CR forms.<BR>Thanks"
                        GoTo get_out
                    Else
                        Dim output() As String = {raw_cr_form}
                        find_xl_cr_file({}, input_file, output, False, format, local, db, err)
                        If output.Count = 0 Then
                            err = "CRREJ: New CR Rejection.  Your email doesn't have any acceptable CR forms.<BR>Thanks"
                            GoTo get_out
                        ElseIf output(0) Like "" Then
                            err = "CRREJ: New CR Rejection.  Your email doesn't have any acceptable CR forms.<BR>Thanks"
                            GoTo get_out
                        Else
                            raw_cr_form = output(0)
                        End If
                    End If

                    '3.3 check cr form format and get the cr_id prefix (tech, region, team)
                    '---------------------------------
                    'create dts to hold the cr form data to minimise me having to keep opening xl
                    '--------------------------------------------------------------
                    Dim ds As New System.Data.DataSet
                    ds.Tables.Add("init_com")
                    ds.Tables.Add("init_data")

                    Dim tech As String = ""
                    Dim region As String = ""
                    Dim region_short As String = ""
                    Dim team_short As String = ""
                    Dim requester As String = msg.from.address
                    Dim approver As String = ""
                    Dim date_received As Date = msg.msg_date
                    Dim cr_id_index As Integer = 0
                    Dim a_cr_type() As String = {""}
                    Dim data_ok As String = "not finished testing"
                    check_requester_cr_format(data_ok, "", ds, raw_cr_form, "", a_cr_type, "", "", "", date_received, requester, tech, region_short, team_short, approver, format, local, err)
                    If data_ok = "nok" Or Not err Like "" Then
                        GoTo get_out
                    End If
                    err = ""

                    Dim a_cr_id(a_cr_type.Count - 1) As String
                    Dim a_cr_type_short(a_cr_type.Count - 1) As String
                    Dim a_cr_form_type(a_cr_type.Count - 1) As String
                    Dim a_cr_path(a_cr_type.Count - 1) As String
                    Dim a_cr_form(a_cr_type.Count - 1) As String

                    '3.4 get cr index from the DB
                    '---------------------------------
                    Dim prefix As String = UCase(region_short & tech & team_short)
                    prefix = UCase(Regex.Replace(prefix, "[^0-9A-Z]", "", RegexOptions.IgnoreCase))
                    cr_id_index = get_cr_id_index(db, prefix, err)
                    If Not err Like "" Then GoTo get_out

                    'creates the new cr_ids and short cr_types as new CR paths and split the form into the different cr types
                    '--------------------------------------------------------------------------
                    Dim i As Integer = 0
                    For Each item In a_cr_type
                        Dim qrow = From row In format.ds_allow.Tables("cr_types").AsEnumerable
                                    Where row.Item("cr_type").ToString Like item
                                    Select row
                        If qrow.Count = 0 Then
                            err = "ER: error finding cr_type in cr_types table"
                            GoTo get_out
                        End If
                        a_cr_form_type(i) = qrow(0)("cr_form_type")
                        a_cr_type_short(i) = qrow(0)("cr_type_short")
                        a_cr_id(i) = prefix & UCase(a_cr_type_short(i)) & cr_id_index
                        a_cr_path(i) = local.base_path & local.cr & "\" & a_cr_id(i)
                        force_make_dir(a_cr_path(i), err)
                        If Not err Like "" Then GoTo get_out

                        'split raw cr form into component crs and put in own directory, it also does the componenet attachment processing and zipping here
                        '-----------------------------------------------------------------------------------------------------------------
                        a_cr_form(i) = ""
                        split_cr_request_form(raw_cr_form, a_cr_form(i), ds, a_cr_id(i), a_cr_type(i), a_cr_type_short(i), a_cr_form_type(i), a_cr_path(i), requester, tech, format, local, err)
                        If Not err Like "" Then GoTo get_out
                        'at this point any attachments for each cr_form_type are in the cr dir and multi-part zipped
                        i += 1
                    Next

                    '3.8 add new cr to the db
                    '-------------------------------
                    Try
                        'merge the hdw and rfr tables as they have the same cr_form_type
                        '------------------------------------------------------------
                        If Not ds.Tables("rfr_com") Is Nothing And Not ds.Tables("rfr_data") Is Nothing Then
                            If ds.Tables("rfr_com").Rows.Count > 0 AndAlso ds.Tables("rfr_data").Rows.Count > 0 Then
                                ds.Tables("rfr_com").TableName = "oth_com"
                                ds.Tables("rfr_data").TableName = "oth_data"
                                If Not ds.Tables("hdw_com") Is Nothing And Not ds.Tables("hdw_data") Is Nothing Then
                                    If ds.Tables("hdw_com").Rows.Count > 0 And ds.Tables("hdw_data").Rows.Count > 0 Then
                                        ds.Tables("oth_com").Merge(ds.Tables("hdw_com"))
                                        ds.Tables("oth_data").Merge(ds.Tables("hdw_data"))
                                        ds.Tables.Remove("hdw_com")
                                        ds.Tables.Remove("hdw_data")
                                    End If
                                End If
                            End If
                        ElseIf Not ds.Tables("hdw_com") Is Nothing And Not ds.Tables("hdw_data") Is Nothing Then
                            If ds.Tables("hdw_com").Rows.Count > 0 And ds.Tables("hdw_data").Rows.Count > 0 Then
                                ds.Tables("hdw_com").TableName = "oth_com"
                                ds.Tables("hdw_data").TableName = "oth_data"
                            End If
                        End If

                        i = 0
                        For Each item In a_cr_form
                            time_now = Now
                            new_cr2db(a_cr_form_type(i), ds, db, a_cr_id(i), err)
                            If Not err Like "" Then GoTo get_out

                            update_cr_common_table(db, a_cr_id(i), "cr_status", "Opened", err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table_date(db, a_cr_id(i), "last_activity_date", time_now, False, err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table_date(db, a_cr_id(i), "last_nag_date", time_now, True, err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table(db, a_cr_id(i), "msg_signature", msg_signature, err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table(db, a_cr_id(i), "cc_list", cc_list, err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table(db, a_cr_id(i), "cr_type_short", a_cr_type_short(i), err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table(db, a_cr_id(i), "cr_form_type", a_cr_form_type(i), err)
                            If Not err Like "" Then GoTo get_out
                            add2log(db, time_now, a_cr_id(i), "CR opened by " & msg.from.address, err)
                            If Not err Like "" Then GoTo get_out
                            i += 1
                        Next
                    Catch ex As Exception
                        If err = "" Then
                            err = "ER: DB error, error adding data to db: " & ex.ToString
                        End If
                        GoTo get_out
                    End Try

                    'here we send the ack to the requester and update the DB => it is either attachments ok or not
                    'note I just send the same files for the missing attachment case, the missing attachments will be shown in the forms
                    '---------------------------------------------------------------------
                    Try
                        '3.7 autoreply to requester that CR was accepted
                        '------------------------------------------------
                        If a_cr_form.Count = 0 Then
                            err = "ER: There are no output cr forms!?  Some internal error."
                            GoTo get_out
                        End If

                        'replies to the requester either good or bad based on the attachment status for each CR
                        '---------------------------------------------------------------
                        Dim cr_string As String = ""
                        Dim subj As String = ""
                        Dim body As String = ""
                        i = 0
                        For Each item In a_cr_form
                            Dim cr_form_zipped As String = ""
                            check_size_and_zip(item, cr_form_zipped, format, local, err)
                            If Not err Like "" Then GoTo get_out

                            subj = "CRMS: CR Acceptance " & "(" & a_cr_id(i) & ")"
                            body = "Dear " & msg.from.displayname & ",<P>Your CR request has been accepted and forwarded to " & approver & " for approval. Please find the the attached CR form:" & array2html_bullet_list({Path.GetFileName(item)})
                            body = body & "<P>Thanks<P>"
                            If msg_sig_cr_id.Count > 0 Then
                                body = body & "<BR>WARNING!  Suspected duplicate CR, there are other open CRs from you with very similar profiles to this one:" & array2html_bullet_list(msg_sig_cr_id)
                                body = body & "You can cancel the unwanted CRs.<P>"
                            End If
                            body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                            auto_reply_general(tx_svr, requester, cc_list, subj, body, If(Not cr_form_zipped Like "", {cr_form_zipped}, {item}), err)
                            If Not err Like "" Then GoTo get_out
                            i += 1
                        Next
                    Catch ex As Exception
                        err = "ER: Error in the autoreply CR acceptance part of the analyse email sub: " & ex.ToString
                        GoTo get_out
                    End Try

                    '3.9 send the CRs to the approver (if attach_ok and data_ok and err like "")
                    '------------------------------------------------------------------------
                    Try
                        If a_cr_form.Count = 0 Then
                            err = "ER: There are no output cr forms!?  Some internal error."
                            GoTo get_out
                        End If
                        i = 0
                        For Each item In a_cr_form
                            Dim cr_attach() As String = {}
                            If IO.File.Exists(a_cr_form(i)) Then
                                cr_attach = FileIO.FileSystem.GetFiles(a_cr_path(i), FileIO.SearchOption.SearchTopLevelOnly, "requester attachments*.zip").ToArray
                                If cr_attach Is Nothing Then cr_attach = {}
                            Else
                                err = "ER: The cr form is not on the HDD!?  Some internal error."
                                GoTo get_out
                            End If

                            'get email address 
                            '---------------------
                            Dim cc_list_final As String = ""
                            Dim user_list As String = ""
                            Dim execution_coordinator As String = ""
                            Dim executors As String = ""
                            cc_list = ""            'can reset as it is in the DB now        
                            get_user_emails(format.cc_mask_approval_request, user_list, requester, approver, execution_coordinator, executors, cc_list, a_cr_id(i), format, db, err)
                            If Not err Like "" Then GoTo get_out
                            Dim to_list As String = approver
                            get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
                            If Not err Like "" Then GoTo get_out

                            'add to log
                            '--------------------
                            time_now = Now
                            add2log(db, time_now, a_cr_id(i), "CR sent for approval to " & to_list, err)
                            If Not err Like "" Then GoTo get_out

                            'send the approval request
                            '------------------------
                            Dim cr_form_zipped As String = ""
                            check_size_and_zip(a_cr_form(i), cr_form_zipped, format, local, err)
                            If Not err Like "" Then GoTo get_out

                            Dim subj As String = "CRMS: CR Approval Request (" & a_cr_id(i) & ")"
                            Dim user_name() As String = get_user_name("approver", a_cr_id(i), format, db)
                            Dim body As String = "Dear " & If(user_name.Count > 0 AndAlso Not user_name.First Like "", Join(user_name, ", "), "Approver") & ",<P>Please check the following attached CR for approval:" & array2html_bullet_list({Path.GetFileName(a_cr_form(i))}) & "<BR>"
                            body = body & "Acceptable Responses:<BR>"
                            body = body & "-----------------------------------------------------"
                            body = body & array2html_bullet_list({"You accept the CR => reply to the email and type 'ok/yes/accepted' in the email body.", "You do not accept the CR => reply to the email and type 'not ok/nok/no/not accepted' in the email body.<BR>You can add any text after this for your reason or notes."})
                            body = body & "<P>Thanks<P>"
                            Dim log_s As String = ""
                            body = body & get_log_string(log_s, db, a_cr_id(i), err)
                            body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                            If cr_attach.Count = 0 Then
                                auto_reply_general(tx_svr, to_list, cc_list_final, subj, body, {If(Not cr_form_zipped Like "", cr_form_zipped, a_cr_form(i))}, err)
                                If Not err Like "" Then GoTo get_out
                            Else
                                Dim j As Integer = 0
                                For Each attachment In cr_attach
                                    If j = 0 Then
                                        auto_reply_general(tx_svr, to_list, cc_list_final, subj & If(cr_attach.Count = 1, "", " => part " & j + 1 & " of " & cr_attach.Count), body, {If(Not cr_form_zipped Like "", cr_form_zipped, a_cr_form(i)), attachment}, err)
                                    Else
                                        auto_reply_general(tx_svr, to_list, cc_list_final, subj & " => part " & j + 1 & " of " & cr_attach.Count, body, {attachment}, err)
                                    End If
                                    If Not err Like "" Then GoTo get_out
                                    j += 1
                                Next
                            End If

                            'Update the cr status in the db after sending successfully to the approver
                            '----------------------------------------------------------
                            update_cr_common_table(db, a_cr_id(i), "cr_status", "Pending Approval", err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table_date(db, a_cr_id(i), "last_activity_date", time_now, False, err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table_date(db, a_cr_id(i), "last_nag_date", time_now, True, err)
                            If Not err Like "" Then GoTo get_out

                            i += 1
                        Next
                    Catch ex As Exception
                        If err = "" Then
                            err = "ER: Error in the autoreply CR acceptance part of the analyse email sub: " & ex.ToString
                        End If
                        GoTo get_out
                    End Try
                    '###########################################################################################################################
                    'End of new CR processing
                    '###########################################################################################################################

                ElseIf Regex.IsMatch(cleansubject, "^cancel$", RegexOptions.IgnoreCase) Then
                    'this deals with requester CR cancel requests
                    'first check it is actually one of these msgs
                    '---------------------------------------------
                    '1) do check that it is from an allowed originator
                    '--------------------------------
                    Dim s_from As String = msg.from.address
                    Dim qrows = From row In format.ds_allow.Tables("requesters").AsEnumerable
                                Where row("email") Like s_from
                                Select row
                    If qrows.Count = 0 Then
                        err = "CANCELREJ: Cancellation Rejection.  Your email does not have CR cancellation rights.<BR>Thanks"
                        GoTo get_out
                    End If

                    '2) basic attachement check
                    '----------------------------
                    If msg.attachments.Count = 0 Then
                        err = "CANCELREJ: If you type cancel in the subject, you must attach the CR forms you wish to cancel.  Your email doesn't have any attachments, giving up.<BR>Thanks"
                        GoTo get_out
                    End If
                    check_attachments(msg, err)
                    If Regex.IsMatch(err, "^er:", RegexOptions.IgnoreCase) Then
                        GoTo get_out
                    ElseIf Not err Like "" Then
                        err = "CANCELREJ: Cancellation Rejection.  Your email doesn't have any acceptable CR forms.  You must attach the forms you which to cancel.<BR>Thanks..." & err
                        GoTo get_out
                    End If

                    '3.2) Find the cr forms
                    '---------------------------
                    Dim cancel_forms() As String = {}
                    Dim cancel_cr_id() As String = {}
                    Dim input_file As String = ""
                    If FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.xlsb").Count = 0 Then
                        err = "CANCELREJ: Cancellation Rejection.  Your email doesn't have any acceptable CR forms.  You must attach the forms you which to cancel.<BR>Thanks"
                        GoTo get_out
                    Else
                        find_xl_cr_file(cancel_cr_id, input_file, cancel_forms, True, format, local, db, err)
                        If cancel_forms.Count = 0 Then
                            err = "CANCELREJ: Cancellation Rejection.  Your email doesn't have any acceptable CR forms.<BR>Thanks"
                            GoTo get_out
                        ElseIf cancel_forms(0) = "" Then
                            err = "CANCELREJ: Cancellation Rejection.  Your email doesn't have any acceptable CR forms.<BR>Thanks"
                            GoTo get_out
                        End If
                    End If

                    'filters the list
                    '----------------------------
                    Dim fail_string As String = " - couldn't cancel as either the cr_id doesn't exist, it is already cancelled or closed or you are not the requester."
                    Dim ok_string As String = " - sucessfully cancelled."
                    Dim i As Integer = 0
                    For Each item In cancel_cr_id
                        'user and state check
                        '------------------------
                        Dim cr_status As String = ""
                        Dim cr_type As String = ""
                        Dim cr_type_short As String = ""
                        Dim cr_form_Type As String = ""
                        Dim cc_list_old As String = ""
                        Dim requester As String = ""
                        Dim approver As String = ""
                        Dim execution_coordinator As String = ""
                        Dim executors As String = ""
                        get_cr_id_data(item, cr_status, cr_type, cr_type_short, cr_form_Type, cc_list_old, requester, approver, execution_coordinator, executors, db, err)
                        If Not err Like "" Then GoTo get_out
                        If cr_type Like "" Or Not msg.from.address Like c2e(requester) Then
                            cancel_cr_id(i) = cancel_cr_id(i) & fail_string
                        Else
                            Dim query = From row In format.ds_allow.Tables("state_control").AsEnumerable
                                        Where row("state") Like cr_status
                                        Select row
                            If query.Count = 0 Then
                                cancel_cr_id(i) = cancel_cr_id(i) & fail_string
                            ElseIf query(0)("can_cancel") = 0 Then
                                cancel_cr_id(i) = cancel_cr_id(i) & fail_string
                            Else
                                cancel_cr_id(i) = cancel_cr_id(i) & ok_string
                            End If
                        End If
                        i += 1
                    Next

                    'takes out the rejected cr_ids and takes out repeats from the cc_list
                    '----------------------------------
                    Dim cancel_cr_id_filter() As String = {}
                    Try
                        Dim qrow = From item In cancel_cr_id.AsEnumerable
                                    Let a = Regex.Replace(item, "\s-\ssucessfully\scancelled\.", "")
                                    Where Not Regex.IsMatch(item, "\s-\scouldn't\scancel\sas\seither\sthe\scr_id\sdoesn't\sexist,\sit\sis\salready\scancelled\sor\sclosed\sor\syou\sare\snot\sthe\srequester\.$") And Not item Like ""
                                    Select a
                        cancel_cr_id_filter = qrow.Distinct.ToArray
                    Catch ex As Exception
                        err = "ER: internal error in the cancel sub, details: " & ex.ToString
                        GoTo get_out
                    End Try
                    If cancel_cr_id_filter Is Nothing Then
                        err = "CANCELREJ: Cancellation Rejection.  Your email doesn't have any acceptable CR forms.  You must attach the forms you which to cancel, the status must not be closed or cancelled and you must be the requester.<BR>Thanks"
                        GoTo get_out
                    ElseIf cancel_cr_id_filter.Count = 0 Then
                        err = "CANCELREJ: Cancellation Rejection.  Your email doesn't have any acceptable CR forms.  You must attach the forms you which to cancel, the status must not be closed or cancelled and you must be the requester.<BR>Thanks"
                        GoTo get_out
                    End If

                    'goes through the final list and processes them
                    '---------------------------------------------
                    For Each cr_id In cancel_cr_id_filter
                        'get data again
                        '------------------------
                        Dim cr_status As String = ""
                        Dim cr_type As String = ""
                        Dim cr_type_short As String = ""
                        Dim cr_form_Type As String = ""
                        Dim cc_list_old As String = ""
                        Dim requester As String = ""
                        Dim approver As String = ""
                        Dim execution_coordinator As String = ""
                        Dim executors As String = ""
                        get_cr_id_data(cr_id, cr_status, cr_type, cr_type_short, cr_form_Type, cc_list_old, requester, approver, execution_coordinator, executors, db, err)
                        If Not err Like "" Then GoTo get_out
                        'no need to validate as it was aready done in the last step

                        'get email address 
                        '---------------------
                        Dim cc_list_final As String = ""
                        Dim user_list As String = ""
                        cc_list = ""            'can reset as it is in the DB now            
                        get_user_emails(format.cc_mask_everyone, user_list, requester, approver, execution_coordinator, executors, cc_list, cr_id, format, db, err)
                        If Not err Like "" Then GoTo get_out
                        Dim to_list As String = user_list
                        get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
                        If Not err Like "" Then GoTo get_out

                        'send ack first
                        '-----------------
                        'send the CR cancel note
                        '-----------------------------
                        time_now = Now
                        Dim subj As String = "CRMS: CR Cancel Acknowledgment"
                        Dim body As String = "Dear " & msg.from.displayname & ",<P>The following CR has been cancelled:" & array2html_bullet_list({cr_id})
                        body = body & "<P>Thanks<P>"
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                        auto_reply_general(tx_svr, msg.from.address, "", subj, body, {""}, err)
                        If Not err Like "" Then GoTo get_out

                        'add to log
                        '------------------
                        add2log(db, time_now, cr_id, "CR cancellation note sent to: " & to_list, err)
                        If Not err Like "" Then GoTo get_out

                        'send the CR cancel note
                        '-----------------------------
                        subj = "CRMS: CR Cancel Note"
                        body = "Please NOTE!  The following CR has been cancelled by (" & msg.from.address & ").  Please stop all work on this CR immediately." & array2html_bullet_list({cr_id})
                        body = body & "<P>Thanks<P>"
                        Dim log_s As String = ""
                        body = body & get_log_string(log_s, db, cr_id, err)
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                        auto_reply_general(tx_svr, to_list, cc_list_final, subj, body, {""}, err)
                        If Not err Like "" Then GoTo get_out

                        'ok, it is good, so update the status and add to the log
                        '----------------------------------------------------
                        update_cr_common_table(db, cr_id, "cr_status", "Cancelled", err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)
                        If Not err Like "" Then GoTo get_out
                        combine_cc_list(cc_list, cc_list_old, format, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table(db, cr_id, "cc_list", cc_list, err)
                        If Not err Like "" Then GoTo get_out
skip:
                    Next
                    '###########################################################################################################################
                    'End of CR multicancel processing
                    '###########################################################################################################################

                ElseIf (Regex.IsMatch(cleansubject, "^(.*)(\(.*\))(.*)$", RegexOptions.IgnoreCase) And Regex.IsMatch(cleanbody, "^cancel", RegexOptions.IgnoreCase)) Then
                    Dim cr_id As String = Regex.Replace(cleansubject, "^(.*)(\()", "", RegexOptions.IgnoreCase)
                    cr_id = Regex.Replace(cr_id, "(\))(.*)$", "", RegexOptions.IgnoreCase)
                    If cr_id Is Nothing Or cr_id = "" Or Len(cr_id) < 5 Then
                        err = "CANCELREJ: Could not read the cr_id from your mail subject => (" & cleansubject & ")"
                        GoTo get_out
                    End If
                    Dim input_file As String = ""

                    Dim cr_status As String = ""
                    Dim cr_type As String = ""
                    Dim cr_type_short As String = ""
                    Dim cr_form_Type As String = ""
                    Dim cc_list_old As String = ""
                    Dim requester As String = ""
                    Dim approver As String = ""
                    Dim execution_coordinator As String = ""
                    Dim executors As String = ""
                    get_cr_id_data(cr_id, cr_status, cr_type, cr_type_short, cr_form_Type, cc_list_old, requester, approver, execution_coordinator, executors, db, err)
                    If Not err Like "" Then GoTo get_out
                    If cr_type Like "" Then
                        err = "XREJ: The cr_id you gave doesn't exist, sorry. (" & cleansubject & ")"
                        GoTo get_out
                    End If

                    'validate email is from allowed user for this change
                    '------------------------------------------------------
                    If Not msg.from.address Like c2e(requester) Then
                        err = "XREJ: Your email doesn't have rights to perform this action, sorry."
                        GoTo get_out
                    End If

                    'filters the list
                    '----------------------------
                    Dim query = From row In format.ds_allow.Tables("state_control").AsEnumerable
                                Where row("state") Like cr_status
                                Select row
                    If query.Count = 0 OrElse query(0)("can_cancel") = 0 Then
                        err = "CANCELREJ: Cancellation Rejection.  You can not cancel your CR, the status is: " & cr_status
                        GoTo get_out
                    End If

                    'updates the cc_list in the DB
                    '----------------------------------
                    combine_cc_list(cc_list, cc_list_old, format, err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table(db, cr_id, "cc_list", cc_list, err)
                    If Not err Like "" Then GoTo get_out

                    'get email address 
                    '---------------------
                    Dim cc_list_final As String = ""
                    Dim user_list As String = ""
                    cc_list = ""            'can reset as it is in the DB now              
                    get_user_emails(format.cc_mask_everyone, user_list, requester, approver, execution_coordinator, executors, cc_list, cr_id, format, db, err)
                    If Not err Like "" Then GoTo get_out
                    Dim to_list As String = user_list
                    get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
                    If Not err Like "" Then GoTo get_out

                    'send ack first
                    '-----------------
                    'send the CR cancel note
                    '-----------------------------
                    time_now = Now
                    Dim subj As String = "CRMS: CR Cancel Acknowledgment"
                    Dim body As String = "Dear " & msg.from.displayname & ",<P>The following CR has been cancelled:" & array2html_bullet_list({cr_id})
                    body = body & "<P>Thanks<P>"
                    body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                    auto_reply_general(tx_svr, msg.from.address, "", subj, body, {""}, err)
                    If Not err Like "" Then GoTo get_out

                    'add to log
                    '------------------------
                    add2log(db, time_now, cr_id, "CR cancellation note sent to: " & to_list, err)
                    If Not err Like "" Then GoTo get_out

                    'send the CR cancel note
                    '-----------------------------
                    subj = "CRMS: CR Cancel Note"
                    body = "Please NOTE!  The following CR has been cancelled by (" & msg.from.address & ").  Please stop all work on this CR immediately." & array2html_bullet_list({cr_id})
                    body = body & "<P>Thanks<P>"
                    Dim log_s As String = ""
                    body = body & get_log_string(log_s, db, cr_id, err)
                    body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                    auto_reply_general(tx_svr, to_list, cc_list_final, subj, body, {""}, err)
                    If Not err Like "" Then GoTo get_out

                    'ok, it is good, so update the status
                    '----------------------------------------------------
                    update_cr_common_table(db, cr_id, "cr_status", "Cancelled", err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)
                    If Not err Like "" Then GoTo get_out
                    '###########################################################################################################################
                    'End of CR single cancel processing
                    '###########################################################################################################################
                Else
                    err = "XREJ: Sorry, I do not undestand the subject text."
                    GoTo get_out
                End If
                '###########################################################################################################################
                'End of do not understand subject processing
                '###########################################################################################################################


                '#####################################################
            ElseIf Regex.IsMatch(cleansubject, "^update\sapp.*$", RegexOptions.IgnoreCase) Then
                '#####################################################
                'here we get the new version of the app and install it

                'check it is from the developer
                '---------------------------------
                If Not msg.from.address Like format.dev_email Then
                    err = "XREJ: Your email address doesn't have rights for this, sorry."
                    GoTo get_out
                End If

                'check the password
                '---------------------
                If Not Regex.IsMatch(Regex.Replace(cleansubject, "^update\sapp\s", "", RegexOptions.IgnoreCase), format.x_factor_email) Then
                    err = "XREJ: Sorry you can't perform this operation, your password is incorrect, sorry."
                    GoTo get_out
                End If

                'basic attachement check
                '----------------------------
                If msg.attachments.Count = 0 Then
                    err = "XREJ: Your email doesn't have any attachments.<BR>Thanks"
                    GoTo get_out
                End If
                check_attachments(msg, err)
                If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then
                    GoTo get_out
                ElseIf Not err Like "" Then
                    err = "RESUBREJ:  " & err
                    GoTo get_out
                End If

                'does specific checks on the attachments for this task
                '-------------------------------------------------------
                Dim attachments_ok As Boolean = True
                For Each item In FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly)
                    If Not Path.GetFileName(item) Like "setup.ext" And Not Path.GetFileName(item) Like "CRMS.application" Then
                        attachments_ok = False
                    End If
                Next
                For Each item In FileIO.FileSystem.GetDirectories(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly)
                    Dim dirInfo As New System.IO.DirectoryInfo(item)
                    Dim dir As String = dirInfo.Name
                    If Not dir Like "Application Files" Then
                        attachments_ok = False
                    End If
                Next
                If Not attachments_ok Then
                    err = "XREJ: Attachments are not correct for this operation, giving up...."
                    GoTo get_out
                End If

                'next rename the exe file so it will work
                '----------------------------------------
                Dim file2run As String = local.base_path & local.inbox & "\setup.exe"
                Try
                    FileIO.FileSystem.MoveFile(local.base_path & local.inbox & "\setup.ext", file2run, True)
                Catch ex As Exception
                    err = "XREJ: Couldn't rename the exe file, giving up...."
                    GoTo get_out
                End Try

                'ack
                '--------------
                Dim subj As String = ""
                Dim body As String = ""
                subj = "CRMS: Update Ready to Go"
                body = "Your update has been checked and is ready to go, standby...."
                body = body & "<P>Thanks<P>"
                auto_reply_general(tx_svr, format.dev_email, "", subj, body, {""}, err)
                If Not err Like "" Then GoTo get_out

                'launch the install
                '-----------------------
                System.Threading.Thread.Sleep(3000)
                Process.Start(file2run)
                Forms.Application.Exit()
                '###########################################################################################################################
                'End of update app processing
                '###########################################################################################################################

            ElseIf Regex.IsMatch(cleansubject, "^update\sfiles.*$", RegexOptions.IgnoreCase) Then
                '#####################################################
                'this lets us update files onthe HDD of the server via email

                'check it is from the developer
                '---------------------------------
                If Not msg.from.address Like format.dev_email Then
                    err = "XREJ: Your email address doesn't have rights for this, sorry."
                    GoTo get_out
                End If

                'check the password
                '---------------------
                If Not Regex.IsMatch(Regex.Replace(cleansubject, "^update\sfiles\s", "", RegexOptions.IgnoreCase), format.x_factor_email) Then
                    err = "XREJ: Sorry you can't perform this operation, your password is incorrect, sorry."
                    GoTo get_out
                End If

                'basic attachement check
                '----------------------------
                If Not msg.attachments.Count = 1 Then
                    err = "XREJ: Only support 1 attachment per request right now.<BR>Thanks"
                    GoTo get_out
                End If
                check_attachments(msg, err)
                If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then
                    GoTo get_out
                ElseIf Not err Like "" Then
                    err = "XREJ:  " & err
                    GoTo get_out
                End If

                'gets the target path from the body
                '------------------------------------------
                Dim target_path As String = ""
                target_path = Trim(Regex.Replace(cleanbody, "((\t)|(\r\n)|(\n\r)(\r)|(\n))+", "", RegexOptions.IgnoreCase And RegexOptions.Singleline))
                If target_path.Length = 0 Then
                    err = "XREJ: There is no acceptable target_path in your email body text."
                    GoTo get_out
                ElseIf Not FileIO.FileSystem.DirectoryExists(Path.GetDirectoryName(target_path)) Then
                    err = "XREJ: The target path doesn't exist.  Tool doesn't support creating paths right now."
                    GoTo get_out
                End If


                'does the file update
                '-----------------------------
                Dim source_path As String = ""
                Try
                    source_path = FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly).First()
                    If Not FileIO.FileSystem.FileExists(source_path) Then
                        err = "XREJ: There was an issue finding the attachement on the HDD."
                        GoTo get_out
                    End If
                    FileIO.FileSystem.MoveFile(source_path, target_path, True)
                Catch ex As Exception
                    err = "ER: there was an issue overwriting the target file, details: " & ex.ToString
                    GoTo get_out
                End Try

                'ack
                '--------------
                Dim subj As String = ""
                Dim body As String = ""
                subj = "CRMS: File Update Done"
                body = "The file: " & target_path & "<BR>has been overwritten by the file: " & source_path & "<P>Thanks.<BR>"
                auto_reply_general(tx_svr, format.dev_email, "", subj, body, {""}, err)
                If Not err Like "" Then GoTo get_out
                '###########################################################################################################################
                'End of update files
                '###########################################################################################################################

            ElseIf Regex.IsMatch(cleansubject, "^clean\sdir\s\.*", RegexOptions.IgnoreCase) Then
                '#####################################################
                'this lets clean any dir we want => be careful

                'check it is from the developer
                '---------------------------------
                If Not msg.from.address Like format.dev_email Then
                    err = "XREJ: Your email address doesn't have rights for this, sorry."
                    GoTo get_out
                End If

                'check the password
                '---------------------
                If Not Regex.IsMatch(Trim(Regex.Replace(cleansubject, "^clean\sdir\s", "", RegexOptions.IgnoreCase)), format.x_factor_email) Then
                    err = "XREJ: Sorry you can't perform this operation, your password is incorrect, sorry."
                    GoTo get_out
                End If

                'gets the target dir from the body
                '--------------------------------
                Dim target_dir As String = ""
                target_dir = Trim(Regex.Replace(cleanbody, "((\t)|(\r\n)|(\n\r)(\r)|(\n))+", "", RegexOptions.IgnoreCase And RegexOptions.Singleline))
                If target_dir.Length = 0 Then
                    err = "XREJ: There is no acceptable directory in your email body text."
                    GoTo get_out
                ElseIf Not FileIO.FileSystem.DirectoryExists(target_dir) Then
                    err = "XREJ: The target direcotry doesn't exist."
                    GoTo get_out
                ElseIf target_dir Like "*:\" Or target_dir Like local.base_path Then       'it is a root dir, do not do it
                    err = "XREJ: This is a root directory, I will not do that."
                    GoTo get_out
                End If

                'does the work
                '-----------------
                clean_dir(target_dir, err)
                If Not err Like "" Then GoTo get_out

                'ack
                '--------------
                Dim subj As String = ""
                Dim body As String = ""
                subj = "CRMS: Directory is Clean"
                body = "The directory: " & target_dir & "<BR>has been cleaned and removed from the recycle bin.<P>Thanks.<BR>"
                auto_reply_general(tx_svr, format.dev_email, "", subj, body, {""}, err)
                If Not err Like "" Then GoTo get_out
                '###########################################################################################################################
                'End of clean dir
                '###########################################################################################################################

            ElseIf Regex.IsMatch(cleansubject, "(CRMS:\s)(CR\sResubmission)((\sRetry\s)|(\s))(Request)(\s*)(\()(.*)(\))", RegexOptions.IgnoreCase) Then
                'this is processing the incoming resubmitted CR
                '---------------------------------------------------
                Dim cr_id As String = Regex.Replace(cleansubject, "^(.*)(CRMS:\s)(CR\sResubmission)((\sRetry\s)|(\s))(Request)(\s*)(\()", "", RegexOptions.IgnoreCase)
                cr_id = Regex.Replace(cr_id, "(\))(.*)$", "", RegexOptions.IgnoreCase)
                If cr_id Is Nothing Or cr_id = "" Or Len(cr_id) < 5 Then
                    err = "RESUBREJ: Could not read the cr_id from your mail subject => (" & cleansubject & ")"
                    GoTo get_out
                End If
                Dim input_file As String = ""
                Dim cr_form = local.base_path & local.cr & "\" & cr_id & "\" & cr_id & ".xlsb"
                Dim cr_form_temp As String = local.base_path & local.inbox & "\" & cr_id & ".xlsb"
                Dim cr_status As String = ""
                Dim cr_type As String = ""
                Dim cr_type_short As String = ""
                Dim cr_form_type As String = ""
                Dim cc_list_old As String = ""
                Dim requester As String = ""
                Dim approver As String = ""
                Dim execution_coordinator As String = ""
                Dim executors As String = ""
                get_cr_id_data(cr_id, cr_status, cr_type, cr_type_short, cr_form_type, cc_list_old, requester, approver, execution_coordinator, executors, db, err)
                If Not err Like "" Then GoTo get_out
                If cr_type Like "" Then
                    err = "XREJ: The cr_id you gave doesn't exist, sorry. (" & cleansubject & ")"
                    GoTo get_out
                End If

                'validate email is from allowed user for this change
                '------------------------------------------------------
                If Not msg.from.address Like c2e(requester) Then
                    err = "XREJ: Your email doesn't have rights to perform this action, sorry."
                    GoTo get_out
                End If

                'state check
                '--------------
                If Not cr_status Like "Pending Resubmission" Then      'main state
                    If Not (cr_status Like "Not Approved" Or cr_status Like "Execution Planning Failed" Or cr_status Like "Execution Rejected" Or cr_status Like "Resubmitted") Then 'intermediate stuck states from email error
                        err = "RESUBREJ: This type of email can only be accepted while the CR has the status 'Pending Resubmission', the current status is: '" & cr_status & "', sorry."
                        GoTo get_out
                    End If
                End If

                'basic attachement check
                '----------------------------
                If msg.attachments.Count = 0 Then
                    err = "RESUBREJ: Your email doesn't have any attachments.<BR>Thanks"
                    GoTo get_out
                End If
                check_attachments(msg, err)
                If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then
                    GoTo get_out
                ElseIf Not err Like "" Then
                    err = "RESUBREJ:  " & err
                    GoTo get_out
                End If

                Dim msg_sig_cr_id() As String = {}
                Try
                    'get the msg signature and check for duplicates, for the resub case
                    '--------------------------------------------------------
                    msg_signature = msg.from.address & "."
                    For Each item In msg.attachments
                        msg_signature = msg_signature & item.filename & "." & item.size & "."
                    Next
                    If Len(msg_signature) > 0 Then msg_signature = Left(msg_signature, Len(msg_signature) - 1)
                    Try
                        clear_dt(dt)
                        sqlquery(False, db, "SELECT cr_id, cr_status FROM " & db.schema & ".cr_common WHERE msg_signature LIKE '" & msg_signature & "' AND cr_status NOT REGEXP '^((closed)|(cancelled)|(pending[[:space:]]resubmission))$';", dt, err)
                        If Not err Like "" Then GoTo get_out
                        If dt.Rows.Count > 0 And Not Regex.IsMatch(cleanbody, "^force", RegexOptions.IgnoreCase) Then
                            'this filters values from a col of a datatable and puts it in an array, linq is soo good.
                            Dim q_temp = From row In dt.AsEnumerable()
                                                    Let a = local.base_path & local.cr & "\" & row.Item("cr_id").ToString & "\" & row.Item("cr_id").ToString & ".xlsb"
                                                    Let b = row.Item("cr_id").ToString
                                                    Where FileIO.FileSystem.FileExists(a)
                                                    Select b
                            msg_sig_cr_id = q_temp.ToArray
                        End If
                    Catch ex As Exception
                        err = "ER: an error occurred checking the msg_signature, details: " & ex.ToString
                        GoTo get_out
                    End Try

                    'find the cr file and attachments and zip the attachments
                    '-----------------------------------------------------------
                    'the normal case where we are pending resubmission
                    '----------------------------------------------------
                    'Find the cr form, chooses the newest one that matches the cr_id
                    '-------------------------------------------------------------
                    If FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.xlsb").Count = 0 Then
                        err = "RESUBREJ: Resubmission Rejection.  Your email doesn't have any acceptable CR forms.  You must attach an acceptable CR form to resubmit.<BR>Thanks"
                        GoTo get_out
                    Else
                        Dim output() As String = {}
                        find_xl_cr_file({cr_id}, input_file, output, False, format, local, db, err)
                        If output.Count = 0 Then
                            err = "RESUBREJ: Resubmission Rejection.  Your email doesn't have any acceptable CR forms.<BR>Thanks"
                            GoTo get_out
                        ElseIf output(0) Like "" Then
                            err = "RESUBREJ: Resubmission Rejection.  Your email doesn't have any acceptable CR forms.<BR>Thanks"
                            GoTo get_out
                        Else
                            cr_form_temp = output(0)
                        End If
                    End If

                    'find the attachments and zips the attachments into the cr dir, 
                    'if there are none, there is no zip file in the cr dir and no requester attachments dir
                    '--------------------------------------------------------------------------
                    process_attachments("requester", cr_id, local, err)
                    If Not err Like "" Then GoTo get_out

                    'check and prepare the cr_form for the approver
                    '-------------------------------------------------
                    Dim ds As New System.Data.DataSet
                    ds.Tables.Add("com")
                    ds.Tables.Add("det")

                    Dim data_ok As String = "not finished testing"
                    Dim date_received As Date = msg.msg_date
                    approver = ""
                    requester = msg.from.address
                    'note: as we know the cr_id here, the next sub will know the caller is for a resub or a pending attachments as opposed to a new cr form
                    'also, if cr_form_temp is blank, then it knows the caller is  pending atatchments
                    check_requester_cr_format(data_ok, cr_id, ds, cr_form_temp, cr_form, {""}, cr_type, cr_type_short, cr_form_type, date_received, requester, "", "", "", approver, format, local, err)
                    If data_ok = "nok" Or Not err Like "" Then GoTo get_out
                    err = ""

                    'we overwrite the cr_form with the resubmitted data for the plain resubmit case and update the DB by deleteing then adding as its a resub
                    '--------------------------------------------------------------------------------------------------------
                    Try
                        force_delete_file(cr_form, err)
                        If Not err Like "" Then GoTo get_out
                        IO.File.Move(cr_form_temp, cr_form)
                    Catch ex As Exception
                        err = "ER: Error overwriting cr_form(processing resubmit)."
                        GoTo get_out
                    End Try

                    Try
                        'delete the old cr data from the common and details tables
                        '----------------------------------------------------------
                        delete_cr_common(db, cr_id, err)
                        If Not err Like "" Then GoTo get_out
                        delete_cr_detail(db, "cr_data_" & cr_form_type, cr_id, err)
                        If Not err Like "" Then GoTo get_out

                        'add the new data
                        '--------------------
                        new_cr2db(cr_form_type, ds, db, cr_id, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table(db, cr_id, "msg_signature", msg_signature, err)
                        If Not err Like "" Then GoTo get_out
                        combine_cc_list(cc_list, cc_list_old, format, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table(db, cr_id, "cc_list", cc_list, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table(db, cr_id, "cr_type_short", cr_type_short, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table(db, cr_id, "cr_form_type", cr_form_type, err)
                        If Not err Like "" Then GoTo get_out


                        'here we update the DB => it is either attachments ok or not
                        '---------------------------------------------------------------------
                        'update the status and add to the log
                        '----------------------------------------------------
                        time_now = Now
                        update_cr_common_table(db, cr_id, "cr_status", "Resubmitted", err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        add2log(db, time_now, cr_id, "CR resubmitted by " & requester, err)
                        If Not err Like "" Then GoTo get_out

                    Catch ex As Exception
                        err = "ER: Error resubmitting data to db: " & ex.ToString
                        GoTo get_out
                    End Try

                    'we only get here on successfull analysis (data_ok ="ok" and err = "")
                    'autoreply to requester that CR was finally accepted
                    '-------------------------------------------------------
                    Try
                        If Not FileIO.FileSystem.DirectoryExists(Path.GetDirectoryName(cr_form)) AndAlso Not FileIO.FileSystem.FileExists(cr_form) Then
                            err = "ER: The cr form is not on the HDD!?  Some internal error."
                            GoTo get_out
                        End If

                        'send the CRs to the approver
                        '----------------------------------
                        'get the attachments
                        '----------------------
                        Dim cr_attach() As String = {}
                        cr_attach = FileIO.FileSystem.GetFiles(Path.GetDirectoryName(cr_form), FileIO.SearchOption.SearchTopLevelOnly, "requester attachments*.zip").ToArray
                        If cr_attach Is Nothing Then cr_attach = {}

                        'get email address 
                        '---------------------
                        Dim cc_list_final As String = ""
                        Dim user_list As String = ""
                        cc_list = ""            'can reset as it is in the DB now                    
                        get_user_emails(format.cc_mask_approval_request, user_list, requester, approver, execution_coordinator, executors, cc_list, cr_id, format, db, err)
                        If Not err Like "" Then GoTo get_out
                        Dim to_list As String = approver
                        get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
                        If Not err Like "" Then GoTo get_out

                        'replies to the requester either good or bad based on the attachment status for each CR
                        '---------------------------------------------------------------
                        Dim cr_form_zipped As String = ""
                        check_size_and_zip(cr_form, cr_form_zipped, format, local, err)
                        If Not err Like "" Then GoTo get_out

                        Dim subj As String = ""
                        Dim body As String = ""
                        subj = "CRMS: CR Acceptance " & "(" & cr_id & ")"
                        body = "Dear " & msg.from.displayname & ",<P>Your CR request has been recorded and forwarded to " & approver & " for approval.<P>Please find the attached CR form:" & array2html_bullet_list({Path.GetFileName(cr_form)})
                        body = body & "<P>Thanks<P>"
                        If msg_sig_cr_id.Count > 0 Then
                            body = body & "<BR>WARNING!  Suspected duplicate CR, there are other open CRs from you with very similar profiles to this one:" & array2html_bullet_list(msg_sig_cr_id)
                            body = body & "You can cancel the unwanted CRs.<P>"
                        End If
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                        auto_reply_general(tx_svr, requester, "", subj, body, {If(Not cr_form_zipped Like "", cr_form_zipped, cr_form)}, err)
                        If Not err Like "" Then GoTo get_out

                        'add to log
                        '--------------------
                        time_now = Now
                        add2log(db, time_now, cr_id, "CR approval request sent to " & to_list, err)
                        If Not err Like "" Then GoTo get_out

                        'send the approval request
                        '--------------------------
                        cr_form_zipped = ""
                        check_size_and_zip(cr_form, cr_form_zipped, format, local, err)
                        If Not err Like "" Then GoTo get_out

                        subj = "CRMS: CR Approval Request (" & cr_id & ")" & If(cr_status = "Resubmitted", " - Resubmission", "")
                        Dim user_name() As String = get_user_name("approver", cr_id, format, db)
                        body = "Dear " & If(user_name.Count > 0 AndAlso Not user_name.First Like "", Join(user_name, ", "), "Approver") & ",<P>Please review the attached CR (" & cr_id & ") for approval.<BR>" & _
                        body = body & "Acceptable Responses:"
                        body = body & "-----------------------------------------------------<BR>"
                        body = body & array2html_bullet_list({"You have no issues => reply to the email and type 'ok/yes/accepted' in the email body.", "You have issues => reply to the email and type 'not ok/nok/no/not accepted' in the email body.<BR>You can add any text after this for your reason or notes."})
                        body = body & "<P>Thanks<P>"
                        Dim log_s As String = ""
                        body = body & get_log_string(log_s, db, cr_id, err)
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                        If cr_attach.Count = 0 Then
                            auto_reply_general(tx_svr, to_list, cc_list_final, subj, body, {If(Not cr_form_zipped Like "", cr_form_zipped, cr_form)}, err)
                            If Not err Like "" Then GoTo get_out
                        Else
                            Dim j As Integer = 0
                            For Each item In cr_attach
                                If j = 0 Then
                                    auto_reply_general(tx_svr, to_list, cc_list_final, subj & If(cr_attach.Count = 1, "", " => part " & j + 1 & " of " & cr_attach.Count), body, {If(Not cr_form_zipped Like "", cr_form_zipped, cr_form), item}, err)
                                Else
                                    auto_reply_general(tx_svr, to_list, cc_list_final, subj & " => part " & j + 1 & " of " & cr_attach.Count, body, {item}, err)
                                End If
                                If Not err Like "" Then GoTo get_out
                                j += 1
                            Next
                        End If

                        'update the status and add to the log
                        '--------------------------------------
                        update_cr_common_table(db, cr_id, "cr_status", "Pending Approval", err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)
                        If Not err Like "" Then GoTo get_out

                    Catch ex As Exception
                        err = "ER: Error sending autoreply for CR submission, details: " & ex.ToString
                        GoTo get_out
                    End Try
                Catch ex As Exception
                    If err = "" Then
                        err = "ER: Error processing the CR resubmission request mail: " & ex.ToString
                    End If
                    GoTo get_out
                End Try
                '###########################################################################################################################
                'End of CR resubmission/pending attachment processing
                '###########################################################################################################################

            ElseIf Regex.IsMatch(cleansubject, "(CRMS:\sCR\sApproval\sRequest)(\s*)(\()(.*)(\))", RegexOptions.IgnoreCase) Then
                'this is the case where the approver approves a CR
                '----------------------------------------------------
                'first check it is actually one of these msgs
                '---------------------------------------------
                Dim cr_id As String = Regex.Replace(cleansubject, "^(.*)(CRMS:\sCR\sApproval\sRequest)(\s*)(\()", "", RegexOptions.IgnoreCase)
                cr_id = Regex.Replace(cr_id, "(\))(.*)$", "", RegexOptions.IgnoreCase)
                If cr_id Is Nothing Or cr_id = "" Or Len(cr_id) < 5 Then
                    err = "APPREJ: Could not read the cr_id from your mail subject => (" & cleansubject & ")"
                    GoTo get_out
                End If
                Dim input_file As String = ""
                Dim cr_form = local.base_path & local.cr & "\" & cr_id & "\" & cr_id & ".xlsb"
                Dim cr_form_temp As String = local.base_path & local.inbox & "\" & cr_id & ".xlsb"
                Dim cr_status As String = ""
                Dim cr_type As String = ""
                Dim cr_type_short As String = ""
                Dim cr_form_type As String = ""
                Dim cc_list_old As String = ""
                Dim requester As String = ""
                Dim approver As String = ""
                Dim execution_coordinator As String = ""
                Dim executors As String = ""
                get_cr_id_data(cr_id, cr_status, cr_type, cr_type_short, cr_form_type, cc_list_old, requester, approver, execution_coordinator, executors, db, err)
                If Not err Like "" Then GoTo get_out
                If cr_type Like "" Then
                    err = "XREJ: The cr_id you gave doesn't exist, sorry. (" & cleansubject & ")"
                    GoTo get_out
                End If

                'validate email is from allowed user for this change
                '------------------------------------------------------
                If Not msg.from.address Like c2e(approver) Then
                    err = "XREJ: Your email doesn't have rights to perform this action, sorry."
                    GoTo get_out
                End If

                'state check
                '--------------
                If Not cr_status = "Pending Approval" Then      'main state
                    If Not (cr_status = "Resubmitted" Or cr_status = "Opened" Or cr_status = "Approved" Or cr_status = "Not Approved") Then 'intermediate stuck states from email error
                        err = "APPREJ: This CR can only be approved while it has the status 'Pending Approval', the current status is: '" & cr_status & "', sorry."
                        GoTo get_out
                    End If
                End If

                'analyse the body text
                '---------------------------
                If Regex.IsMatch(cleanbody, "^((ok)|(approved)|(yes)|(lanjut)|(carry\son))", RegexOptions.IgnoreCase) Then
                    Try
                        'ok, it is good, so update the status and add to the log
                        '--------------------------------------
                        time_now = Now
                        update_cr_common_table(db, cr_id, "cr_status", "Approved", err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "approval_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        combine_cc_list(cc_list, cc_list_old, format, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table(db, cr_id, "cc_list", cc_list, err)
                        If Not err Like "" Then GoTo get_out
                        add2log(db, time_now, cr_id, "CR approved by " & msg.from.address, err)
                        If Not err Like "" Then GoTo get_out

                        'send the CRs to the execution coordinator
                        '--------------------------------------------
                        'find the cr file, using the HDD version here, also the attachments
                        '-----------------------------------------------------------------
                        cr_form_temp = ""
                        If Not FileIO.FileSystem.FileExists(cr_form) Then
                            err = "ER: the cr_form was not found on the server HDD, internal error, can't continue...."
                            GoTo get_out
                        End If
                        Dim cr_attach() As String = FileIO.FileSystem.GetFiles(Path.GetDirectoryName(cr_form), FileIO.SearchOption.SearchTopLevelOnly, "requester attachments*.zip").ToArray
                        If cr_attach Is Nothing Then cr_attach = {}

                        'prep form for excoord
                        '----------------------------------------------------
                        Dim approval_date As DateTime = time_now
                        prepare_cr_form_for_excoord(approval_date, cr_form, cr_form_type, format, local, err)
                        If Not err Like "" Then GoTo get_out

                        'get email address 
                        '---------------------
                        Dim cc_list_final As String = ""
                        Dim user_list As String = ""
                        cc_list = ""            'can reset as it is in the DB now             
                        get_user_emails(format.cc_mask_execution_planning_request, user_list, requester, approver, execution_coordinator, executors, cc_list, cr_id, format, db, err)
                        If Not err Like "" Then GoTo get_out
                        Dim to_list As String = execution_coordinator
                        get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
                        If Not err Like "" Then GoTo get_out

                        'send the ack
                        '------------------
                        Dim subj As String = "CRMS: Acknowledgement (" & cr_id & ")"
                        Dim body As String = "Dear " & msg.from.displayname & ",<P>Your Acceptance reply was recorded and forwarded to " & execution_coordinator & " for execution planning."
                        body = body & "<P>Thanks<P>"
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                        auto_reply_general(tx_svr, msg.from.address, "", subj, body, {}, err)
                        If Not err Like "" Then GoTo get_out

                        'add to log
                        '----------------------
                        time_now = Now
                        add2log(db, time_now, cr_id, "CR execution planning request sent to " & to_list, err)
                        If Not err Like "" Then GoTo get_out

                        'send the exec planning request
                        '-------------------------------
                        Dim cr_form_zipped As String = ""
                        check_size_and_zip(cr_form, cr_form_zipped, format, local, err)
                        If Not err Like "" Then GoTo get_out

                        subj = "CRMS: CR Execution Planning Request (" & cr_id & ")"
                        Dim user_name() As String = get_user_name("execution_coordinator", cr_id, format, db)
                        body = "Dear " & If(user_name.Count > 0 AndAlso Not user_name.First Like "", Join(user_name, ", "), "Execution Coordinator") & ",<P>Please open the CR and enter the expected/planned date of execution in the yellow field, then close/save/attach/reply with the CR:" & array2html_bullet_list({Path.GetFileName(cr_form)})
                        body = body & "<BR>Acceptable Responses:<BR>"
                        body = body & "-----------------------------------------------------"
                        body = body & array2html_bullet_list({"You have no issues => reply to the email and attach the updated cr form.", "You have issues => reply to the email and type 'not ok/nok/no' in the email body.<BR>You can add any text after this for your reason or notes."})
                        body = body & "<P>Thanks<P>"
                        Dim log_s As String = ""
                        body = body & get_log_string(log_s, db, cr_id, err)
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
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
                        update_cr_common_table(db, cr_id, "cr_status", "Pending Execution Planning", err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)
                        If Not err Like "" Then GoTo get_out

                    Catch ex As Exception
                        If err = "" Then
                            err = "ER: Error processing the approver reply mail: " & ex.ToString
                        End If
                        GoTo get_out
                    End Try
                    '###########################################################################################################################
                    'End of approved CR processing
                    '###########################################################################################################################

                ElseIf Regex.IsMatch(cleanbody, "^((nok)|(not\sapproved)|(not\saccepted)|(no)|(tolak)|(fuck\soff))", RegexOptions.IgnoreCase) Then
                    'create the resub form in the cr dir and delete any attachments
                    '-----------------------------------------------------------
                    cr_form_temp = ""
                    Dim cr_resub_form As String = ""
                    pre_process_for_resub(cr_resub_form, cr_id, cr_form, local, err)
                    If Not err Like "" Then GoTo get_out
                    create_resubmit_form(cr_resub_form, cr_form_type, format, local, err)
                    If Not err Like "" Then GoTo get_out

                    'ok, it is good, so update the status and add to the log
                    '----------------------------------------------------
                    time_now = Now
                    update_cr_common_table(db, cr_id, "cr_status", "Not Approved", err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                    If Not err Like "" Then GoTo get_out
                    combine_cc_list(cc_list, cc_list_old, format, err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table(db, cr_id, "cc_list", cc_list, err)
                    If Not err Like "" Then GoTo get_out
                    add2log(db, time_now, cr_id, "CR not approved by " & msg.from.address, err)
                    If Not err Like "" Then GoTo get_out

                    'get email address 
                    '---------------------
                    Dim cc_list_final As String = ""
                    Dim user_list As String = ""
                    cc_list = ""            'can reset as it is in the DB now
                    get_user_emails(format.cc_mask_resubmission_request, user_list, requester, approver, execution_coordinator, executors, cc_list, cr_id, format, db, err)
                    If Not err Like "" Then GoTo get_out
                    Dim to_list As String = requester
                    get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
                    If Not err Like "" Then GoTo get_out

                    'send the ack
                    '------------------
                    Dim subj As String = "CRMS: Acknowledgement (" & cr_id & ")"
                    Dim body As String = "Dear " & msg.from.displayname & ",<P>Your Acceptance reply was recorded, the CR has been sent back to " & requester & " for resubmission."
                    body = body & "<P>Thanks<P>"
                    body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                    auto_reply_general(tx_svr, msg.from.address, "", subj, body, {}, err)
                    If Not err Like "" Then GoTo get_out

                    'add to log
                    '--------------------------
                    time_now = Now
                    add2log(db, time_now, cr_id, "CR approval rejection note sent to: " & to_list, err)
                    If Not err Like "" Then GoTo get_out

                    'send note to requester
                    '------------------------
                    Dim cr_form_zipped As String = ""
                    check_size_and_zip(cr_resub_form, cr_form_zipped, format, local, err)
                    If Not err Like "" Then GoTo get_out

                    subj = "CRMS: CR Resubmission Request (" & cr_id & ") - !!Approval Rejection!!"
                    Dim user_name() As String = get_user_name("requester", cr_id, format, db)
                    body = "Dear " & If(user_name.Count > 0 AndAlso Not user_name.First Like "", Join(user_name, ", "), "Requester") & ",<P>Your CR (" & cr_id & ") was not approved.<BR>If you want to continue, make modifications to the resubmit form and resubmit it."
                    body = body & "<P>Thanks<P>"
                    Dim log_s As String = ""
                    body = body & get_log_string(log_s, db, cr_id, err)
                    body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                    auto_reply_general(tx_svr, to_list, cc_list_final, subj, body, {If(Not cr_form_zipped Like "", cr_form_zipped, cr_resub_form)}, err)
                    If Not err Like "" Then GoTo get_out

                    'delete the resubmit form as we do not need it any more
                    '--------------------------------------------------
                    FileIO.FileSystem.DeleteFile(cr_resub_form)

                    'ok, it is good, so update the status and add to the log
                    '----------------------------------------------------
                    update_cr_common_table(db, cr_id, "cr_status", "Pending Resubmission", err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)
                    If Not err Like "" Then GoTo get_out

                    '###########################################################################################################################
                    'End of not approved CR processing
                    '###########################################################################################################################
                Else
                    err = "XREJ: Sorry, I do not undestand the body text."
                    GoTo get_out
                End If

            ElseIf Regex.IsMatch(cleansubject, "(CRMS:\sCR\sExecution\sPlanning)((\sRetry\s)|(\s))(Request)(\s*)(\()", RegexOptions.IgnoreCase) Then
                'this is the case where the execution coordinator replies with the planned dates and executors
                'they just have to reply with the modified CR form attached, no extra body text or subject text.
                '----------------------------------------------------------------------------
                Dim cr_id As String = Regex.Replace(cleansubject, "^(.*)(CRMS:\sCR\sExecution\sPlanning)((\sRetry\s)|(\s))(Request)(\s*)(\()", "", RegexOptions.IgnoreCase)
                cr_id = Regex.Replace(cr_id, "(\))(.*)$", "", RegexOptions.IgnoreCase)
                If cr_id Is Nothing Or cr_id = "" Or Len(cr_id) < 5 Then
                    err = "EXCOORDREJ: Could not read the cr_id from your mail subject => (" & cleansubject & ")"
                    GoTo get_out
                End If
                Dim input_file As String = ""
                Dim cr_form = local.base_path & local.cr & "\" & cr_id & "\" & cr_id & ".xlsb"
                Dim cr_form_temp As String = local.base_path & local.inbox & "\" & cr_id & ".xlsb"
                Dim cr_status As String = ""
                Dim cr_type As String = ""
                Dim cr_type_short As String = ""
                Dim cr_form_type As String = ""
                Dim cc_list_old As String = ""
                Dim requester As String = ""
                Dim approver As String = ""
                Dim execution_coordinator As String = ""
                Dim executors As String = ""
                get_cr_id_data(cr_id, cr_status, cr_type, cr_type_short, cr_form_type, cc_list_old, requester, approver, execution_coordinator, executors, db, err)
                If Not err Like "" Then GoTo get_out
                If cr_type Like "" Then
                    err = "XREJ: The cr_id you gave doesn't exist, sorry. (" & cleansubject & ")"
                    GoTo get_out
                End If

                'validate email is from allowed user for this change
                '------------------------------------------------------
                If cr_type_short Like "prm" Then
                    'for the parameter case, we allow any ex coord to answer the mail
                    Dim test_email As String = msg.from.address
                    Dim qrows = From row In format.ds_allow.Tables("prm_ex_coord")
                                Where row.Field(Of String)("email") Like test_email
                                Select row
                    If qrows.Count = 0 Then
                        err = "XREJ: Your email doesn't have rights to perform this action, sorry."
                        GoTo get_out
                    End If
                Else
                    If Not msg.from.address Like c2e(execution_coordinator) Then
                        err = "XREJ: Your email doesn't have rights to perform this action, sorry."
                        GoTo get_out
                    End If
                End If

                'state check
                '--------------
                If Not cr_status = "Pending Execution Planning" Then      'main state
                    If Not (cr_status = "Approved" Or cr_status = "Execution Planning Failed" Or cr_status = "Execution Planned") Then 'intermediate stuck states from email error
                        err = "EXCOORDREJ: This CR can only have it's execution planned while it has the status 'Pending Execution Planning', the current status is: '" & cr_status & "', sorry."
                        GoTo get_out
                    End If
                End If

                If Not Regex.IsMatch(cleanbody, "^((nok)|(not\saccepted)|(no)|(tolak)|(fuck\soff))", RegexOptions.IgnoreCase) Then
                    'this is the positive case where the execution was planned, the dates are in the attached CR form
                    'check the cr_id is in the db and get the cr_form_type(no need to check it if the cr_id exists)
                    '------------------------------------------------------------------------------
                    'basic attachement check
                    '----------------------------
                    If msg.attachments.Count = 0 Then
                        err = "EXCOORDREJ: Your email doesn't have any attachments.<BR>Thanks"
                        GoTo get_out
                    End If
                    check_attachments(msg, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then
                        GoTo get_out
                    ElseIf Not err Like "" Then
                        err = "EXCOORDREJ:  " & err
                        GoTo get_out
                    End If

                    Try
                        'Find the cr form, chooses the newest one that matches the cr_id
                        '-------------------------------------------------------------
                        If FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.xlsb").Count = 0 Then
                            err = "EXCOORDREJ: Your email doesn't have any acceptable CR forms.  You must attach an acceptable CR form.<BR>Thanks"
                            GoTo get_out
                        Else
                            Dim output() As String = {}
                            find_xl_cr_file({cr_id}, input_file, output, False, format, local, db, err)
                            If output.Count = 0 Then
                                err = "EXCOORDREJ: Your email doesn't have any acceptable CR forms.<BR>Thanks"
                                GoTo get_out
                            ElseIf output(0) Like "" Then
                                err = "EXCOORDREJ: Your email doesn't have any acceptable CR forms.<BR>Thanks"
                                GoTo get_out
                            Else
                                cr_form_temp = output(0)
                            End If
                        End If

                        'check and prepare the cr_form for the executor
                        '---------------------------------------------
                        Dim executors_raw As String = ""
                        Dim planned_ex_date As DateTime = Now
                        clear_dt(dt)
                        Dim data_ok As String = "not finished testing"
                        process_ex_coord_cr_form(data_ok, cr_id, cr_form_temp, planned_ex_date, executors_raw, dt, cr_form_type, format, local, err)
                        If data_ok = "nok" Or Not err Like "" Then
                            GoTo get_out
                        End If
                        err = ""
                        'at this point the planned_ex_date is either 6 hours after 'Now' if it is today or at 6pm on the future date chosen
                        'executors_raw is the cleaned executors comma delimited list or just the ex_coord if the ex list was empty

                        'if there was any data error, we already got out, so here we are good to continue
                        'we overwrite the old cr_form with the new one
                        '--------------------------------------------
                        Try
                            If IO.File.Exists(cr_form) Then
                                force_delete_file(cr_form, err)
                                If Not err Like "" Then GoTo get_out
                            End If
                            IO.File.Copy(cr_form_temp, cr_form)
                        Catch ex As Exception
                            err = "ER: Error overwriting old cr_form (processing ex coord's email)."
                            GoTo get_out
                        End Try

                        'Update the correct table in the DB with the data from the exec.coord (2 cols only in the data table)
                        '------------------------------------------------------------------------------------------------------
                        update_cr_data(db, dt, "cr_data_" & cr_form_type, err)
                        If Not err Like "" Then GoTo get_out

                        'ok, it is good, so update the status and add to the log
                        '----------------------------------------------------
                        time_now = Now
                        update_cr_common_table(db, cr_id, "cr_status", "Execution Planned", err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "planned_execution_date", planned_ex_date, False, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table(db, cr_id, "executors", executors_raw, err)
                        If Not err Like "" Then GoTo get_out
                        combine_cc_list(cc_list, cc_list_old, format, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table(db, cr_id, "cc_list", cc_list, err)
                        If Not err Like "" Then GoTo get_out
                        add2log(db, time_now, cr_id, "CR execution date planned by " & msg.from.address, err)
                        If Not err Like "" Then GoTo get_out

                        'send the CRs to the executer or execution coordinator if there is no email for the executor
                        '---------------------------------------------------------------------------
                        'get the attachments
                        '----------------------
                        Dim cr_attach() As String = {}
                        cr_attach = FileIO.FileSystem.GetFiles(Path.GetDirectoryName(cr_form), FileIO.SearchOption.SearchTopLevelOnly, "requester attachments*.zip").ToArray
                        If cr_attach Is Nothing Then cr_attach = {}


                        'get the other email address for the cc list
                        '--------------------------------------------------
                        Dim cc_list_final As String = ""
                        Dim user_list As String = ""
                        cc_list = ""            'can reset as it is in the DB now             
                        get_user_emails(format.cc_mask_execution_request, user_list, requester, approver, execution_coordinator, executors, cc_list, cr_id, format, db, err)
                        If Not err Like "" Then GoTo get_out
                        Dim to_list As String = If(executors Like "", execution_coordinator, executors)
                        get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
                        If Not err Like "" Then GoTo get_out

                        'send the ack
                        '------------------
                        Dim subj As String = "CRMS: Acknowledgement (" & cr_id & ")"
                        Dim body As String = "Dear " & msg.from.displayname & ",<P>Your planned execution date has been recorded and the CR has been forwarded to " & executors & " for execution."
                        body = body & "<P>Thanks<P>"
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                        auto_reply_general(tx_svr, msg.from.address, "", subj, body, {}, err)
                        If Not err Like "" Then GoTo get_out

                        'add to log
                        '-------------------------
                        time_now = Now
                        add2log(db, time_now, cr_id, "CR execution request sent to " & to_list, err)
                        If Not err Like "" Then GoTo get_out

                        'send the exec request
                        '--------------------------
                        Dim cr_form_zipped As String = ""
                        check_size_and_zip(cr_form, cr_form_zipped, format, local, err)
                        If Not err Like "" Then GoTo get_out

                        subj = "CRMS: CR Execution Request (" & cr_id & ")"
                        Dim user_name() As String = get_user_name("executors", cr_id, format, db)
                        body = "Dear " & If(user_name.Count > 0 AndAlso Not user_name.First Like "", Join(user_name, ", "), "Executors") & ",<P>Please execute the attached CR by the planned date indicated.<BR>" & array2html_bullet_list({Path.GetFileName(cr_form)})
                        body = body & "After execution, please complete all yellow fields in the CR then close/save/attach/reply with the CR<BR>"
                        body = body & "For the reviewer to accept the result of the CR, please attach the verification files with:" & array2html_bullet_list({"Parameter/Hardware CRs - CR id in the file name", "RF Basic/RF Re-engineering CRs - nodename in the filename"})
                        body = body & "<P>Thanks<P>"
                        Dim log_s As String = ""
                        body = body & get_log_string(log_s, db, cr_id, err)
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
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
                        update_cr_common_table(db, cr_id, "cr_status", "Pending Execution", err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)
                        If Not err Like "" Then GoTo get_out

                    Catch ex As Exception
                        If err = "" Then
                            err = "ER: Error processing the ex.coord reply mail: " & ex.ToString
                        End If
                        GoTo get_out
                    End Try
                    '###########################################################################################################################
                    'End of execution planned results processing
                    '###########################################################################################################################

                ElseIf Regex.IsMatch(cleanbody, "^((nok)|(not\saccepted)|(no)|(tolak)|(fuck\soff))", RegexOptions.IgnoreCase) Then
                    Try
                        'create the resub form in the cr dir and delete any attachments
                        '-----------------------------------------------------------
                        cr_form_temp = ""
                        Dim cr_resub_form As String = ""
                        pre_process_for_resub(cr_resub_form, cr_id, cr_form, local, err)
                        If Not err Like "" Then GoTo get_out
                        create_resubmit_form(cr_resub_form, cr_form_type, format, local, err)
                        If Not err Like "" Then GoTo get_out

                        'ok, it is good, so update the status and add to the log
                        '----------------------------------------------------
                        time_now = Now
                        update_cr_common_table(db, cr_id, "cr_status", "Execution Planning Failed", err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        combine_cc_list(cc_list, cc_list_old, format, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table(db, cr_id, "cc_list", cc_list, err)
                        If Not err Like "" Then GoTo get_out
                        add2log(db, time_now, cr_id, "CR execution date planning rejected by " & msg.from.address, err)
                        If Not err Like "" Then GoTo get_out

                        'send a note to the requester with the resub form
                        '------------------------------------------------
                        'get the other email address for the cc list
                        '--------------------------------------------------
                        Dim cc_list_final As String = ""
                        Dim user_list As String = ""
                        cc_list = ""            'can reset as it is in the DB now              
                        get_user_emails(format.cc_mask_resubmission_request, user_list, requester, approver, execution_coordinator, executors, cc_list, cr_id, format, db, err)
                        If Not err Like "" Then GoTo get_out
                        Dim to_list As String = requester
                        get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
                        If Not err Like "" Then GoTo get_out

                        'send the ack
                        '------------------
                        Dim subj As String = "CRMS: Acknowledgement (" & cr_id & ")"
                        Dim body As String = "Dear " & msg.from.displayname & ",<P>Your reply was recorded, the CR has been sent back to " & requester & " for resubmission."
                        body = body & "<P>Thanks<P>"
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                        auto_reply_general(tx_svr, msg.from.address, "", subj, body, {}, err)
                        If Not err Like "" Then GoTo get_out

                        'add to log
                        '----------------------
                        time_now = Now
                        add2log(db, time_now, cr_id, "CR execution planning rejection note sent to: " & to_list, err)
                        If Not err Like "" Then GoTo get_out

                        'send
                        '-------
                        Dim cr_form_zipped As String = ""
                        check_size_and_zip(cr_resub_form, cr_form_zipped, format, local, err)
                        If Not err Like "" Then GoTo get_out

                        subj = "CRMS: CR Resubmission Request (" & cr_id & ") - !!Execution Planning Rejection!!"
                        Dim user_name() As String = get_user_name("requester", cr_id, format, db)
                        body = "Dear " & If(user_name.Count > 0 AndAlso Not user_name.First Like "", Join(user_name, ", "), "Requester") & ",<P>Please NOTE!  The CR (" & cr_id & ") sent in for execution planning by (" & msg.from.address & "), was not able to be planned.<BR>"
                        body = body & "Please review, make any changes to the resubmit form and resubmit it to proceed with the CR."
                        body = body & "<P>Thanks<P>"
                        Dim log_s As String = ""
                        body = body & get_log_string(log_s, db, cr_id, err)
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                        auto_reply_general(tx_svr, to_list, cc_list_final, subj, body, {If(Not cr_form_zipped Like "", cr_form_zipped, cr_resub_form)}, err)
                        If Not err Like "" Then GoTo get_out

                        'delete the resubmit form as we do not need it any more
                        '--------------------------------------------------
                        FileIO.FileSystem.DeleteFile(cr_resub_form)

                        'ok, it is good, so update the status and add to the log
                        '----------------------------------------------------
                        update_cr_common_table(db, cr_id, "cr_status", "Pending Resubmission", err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)        'reset value to null
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "approval_date", time_now, True, err)        'reset value to null
                        If Not err Like "" Then GoTo get_out

                    Catch ex As Exception
                        If err = "" Then
                            err = "ER: Error processing the ex.coord reply mail: " & ex.ToString
                        End If
                        GoTo get_out
                    End Try
                    '###########################################################################################################################
                    'End of failed execution planned results processing
                    '###########################################################################################################################

                Else
                    err = "XREJ: Sorry, I do not undestand the body text."
                    GoTo get_out
                End If


            ElseIf Regex.IsMatch(cleansubject, "(CRMS:\sCR\sExecution\s)((Request)|(Retry\sRequest)|(Complete\sBut\sMissing\sAttachments))(\s*)(\()(.*)(\))", RegexOptions.IgnoreCase) Then
                'this is the case where the executor has finished the CR and has returned the completed CR form.
                'if the executor could not finish any sub crs, they would have set the execution status to "fail"
                'first check it is actually one of these msgs
                '---------------------------------------------
                'check the cr_id is in the db and get the cr_form_type(no need to check it if the cr_id exists)
                '------------------------------------------------------------------------------
                Dim cr_id As String = Regex.Replace(cleansubject, "^(.*)(CRMS:\sCR\sExecution\s)((Request)|(Retry\sRequest)|(Complete\sBut\sMissing\sAttachments))(\s*)(\()", "", RegexOptions.IgnoreCase)
                cr_id = Regex.Replace(cr_id, "(\))(.*)$", "", RegexOptions.IgnoreCase)
                If cr_id Is Nothing Or cr_id = "" Or Len(cr_id) < 5 Then
                    err = "EXREJ: Could not read a valid cr_id from your mail subject => (" & cleansubject & ")"
                    GoTo get_out
                End If
                Dim input_file As String = ""
                Dim cr_form = local.base_path & local.cr & "\" & cr_id & "\" & cr_id & ".xlsb"
                Dim cr_form_temp As String = local.base_path & local.inbox & "\" & cr_id & ".xlsb"
                Dim cr_status As String = ""
                Dim cr_type As String = ""
                Dim cr_type_short As String = ""
                Dim cr_form_Type As String = ""
                Dim cc_list_old As String = ""
                Dim requester As String = ""
                Dim approver As String = ""
                Dim execution_coordinator As String = ""
                Dim executors As String = ""
                get_cr_id_data(cr_id, cr_status, cr_type, cr_type_short, cr_form_Type, cc_list_old, requester, approver, execution_coordinator, executors, db, err)
                If Not err Like "" Then GoTo get_out
                If cr_type Like "" Then
                    err = "XREJ: The cr_id you gave doesn't exist, sorry. (" & cleansubject & ")"
                    GoTo get_out
                End If

                'check the msg is from the executor or the excoord and checks the cr_id
                '-------------------------------------------------------------------------
                Dim temp_ex() As String = Strings.Split(executors, ",")
                Dim temp_found As Boolean = False
                For Each item In temp_ex
                    If msg.from.address Like c2e(item) Then
                        temp_found = True
                    End If
                Next
                If Not temp_found And Not msg.from.address Like c2e(execution_coordinator) Then
                    err = "XREJ: Your email doesn't have rights to perform this action, sorry."
                    GoTo get_out
                End If

                'state check - we include the state before the pending state as this could happen if the email failed to get logged as sent, but it was sent anyway, happens with sketchy network connetions
                '----------------------------------------------------------------------
                If Not (cr_status Like "Pending Execution" Or cr_status Like "Execution Complete Pending Attachments") Then      'main state
                    If Not (cr_status Like "Execution Planned" Or cr_status Like "Execution Complete" Or cr_status Like "Execution Rejected") Then 'intermediate stuck states from email error
                        err = "EXREJ: This CR can only be executed while it has the status 'Pending Execution' or 'Execution Complete Pending Attachments', the current status is: '" & cr_status & "', sorry."
                        GoTo get_out
                    End If
                End If

                If Not Regex.IsMatch(cleanbody, "^((nok)|(no)|(tolak)|(not\sok))", RegexOptions.IgnoreCase) Then
                    'this is the positive case where the execution done
                    'check the cr_id is in the db and get the cr_form_type(no need to check it if the cr_id exists)
                    '------------------------------------------------------------------------------
                    'basic attachement check
                    '----------------------------
                    If msg.attachments.Count = 0 Then
                        err = "EXREJ: Your email doesn't have any attachments.<BR>Thanks"
                        GoTo get_out
                    End If
                    check_attachments(msg, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then
                        GoTo get_out
                    ElseIf Not err Like "" Then
                        err = "EXREJ:  " & err
                        GoTo get_out
                    End If

                    Try
                        If Not cr_status Like "Execution Complete Pending Attachments" Then
                            'the normal case where we are pending execution
                            '----------------------------------------
                            'Find the cr form, chooses the newest one that matches the cr_id
                            '-------------------------------------------------------------
                            If FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.xlsb").Count = 0 Then
                                err = "EXREJ: Your email doesn't have any acceptable CR forms.  You must attach an acceptable CR form.<BR>Thanks"
                                GoTo get_out
                            Else
                                Dim output() As String = {}
                                find_xl_cr_file({cr_id}, input_file, output, False, format, local, db, err)
                                If output.Count = 0 Then
                                    err = "EXREJ: Your email doesn't have any acceptable CR forms.<BR>Thanks"
                                    GoTo get_out
                                ElseIf output(0) Like "" Then
                                    err = "EXREJ: Your email doesn't have any acceptable CR forms.<BR>Thanks"
                                    GoTo get_out
                                Else
                                    cr_form_temp = output(0)
                                End If
                            End If

                            'find the cr file and attachments and zips the attachments and puts it in the cr dir
                            'if there are no attachments, it just doesn't make any zip file nor executors directory
                            '---------------------------------------------------------------------
                            process_attachments("executor", cr_id, local, err)
                            If Not err Like "" Then GoTo get_out

                        ElseIf cr_status Like "Execution Complete Pending Attachments" Then
                            'this is done on repeated ex result return attempts when they send more attachments, it is not done the first time when the status = pending execution
                            '------------------------------------------------------------------------------------------
                            'only look for more attachments and add them to the executor attachments zip files
                            '---------------------------------------------------------------------------------
                            'get the cr_form
                            '-------------------
                            cr_form_temp = ""
                            If Not FileIO.FileSystem.FileExists(cr_form) Then
                                err = "ER: the cr_form was not found on the server HDD, internal error, can't continue...."
                                GoTo get_out
                            End If

                            'puts all the attachments in the executor attachments dir of the inbox for checking and also multipart zips them to the cr dir
                            '------------------------------------------------------------------------------------------------------
                            process_attachments_combine("executor", cr_id, local, err)
                            If Not err Like "" Then GoTo get_out

                        End If

                        'check and prepare the cr_form for the reviewer
                        'if we are missing attachments, it will show here with err = EXREJ: See returned CR form......
                        '----------------------------------------------------------------------------
                        Dim attach_ok As String = "not finished testing"
                        Dim data_ok As String = "not finished testing"
                        Dim cr_form_to_pass As String = If(cr_status Like "Execution Complete Pending Attachments", cr_form, cr_form_temp)
                        Dim ex_date As DateTime = Now
                        Dim executors_raw As String = ""
                        clear_dt(dt)
                        process_ex_cr_form(executors_raw, cr_id, cr_status, cr_form_to_pass, dt, cr_type, cr_form_Type, ex_date, attach_ok, data_ok, local, format, err)
                        If data_ok = "nok" Or Not err Like "" Then
                            GoTo get_out
                        End If
                        err = ""
                        'NOTE: ex_date is not change dby the sub, it is still = now, the detailed dates are in the details table, the cr_common ex_date is just now which is used for nagging etc.
                        'executors_raw is the cleaned executors comma delimited list or just the ex_coord if the ex list was empty

                        'the normal case where we are pending execution, we write the cr form to the cr dir
                        '-----------------------------------------------------------------------
                        If Not cr_status Like "Execution Complete Pending Attachments" Then
                            'if we are good, we overwrite the old cr_form with the new one
                            '--------------------------------------------------------------
                            Try
                                If IO.File.Exists(cr_form) Then
                                    force_delete_file(cr_form, err)
                                    If Not err Like "" Then GoTo get_out
                                End If
                                FileIO.FileSystem.CopyFile(cr_form_temp, cr_form, True)
                            Catch ex As Exception
                                err = "ER: Error overwriting old cr_form (processing ex's email)."
                                GoTo get_out
                            End Try
                        End If

                        'Update the correct table in the DB with the data from the executor
                        '----------------------------------------------------------------------
                        update_cr_data(db, dt, "cr_data_" & cr_form_Type, err)
                        If Not err Like "" Then GoTo get_out

                        'here we update the DB => it is either attachments ok or not
                        '---------------------------------------------------------------------
                        If attach_ok Like "nok" Then
                            Dim file_out As String = ""
                            If Regex.IsMatch(cr_type, "^((RF\sRe-engineering)|(RF Basic))$", RegexOptions.IgnoreCase) Then
                                file_out = local.base_path & local.inbox & "\" & cr_id & "_errors.xlsb"
                                If Not FileIO.FileSystem.FileExists(file_out) Then
                                    err = "ER: error file not found in inbox."
                                    GoTo get_out
                                End If
                            End If

                            'update the status and add to the log
                            '----------------------------------------------------
                            time_now = Now
                            update_cr_common_table(db, cr_id, "cr_status", "Execution Complete Pending Attachments", err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                            If Not err Like "" Then GoTo get_out
                            If Not cr_status Like "Execution Complete Pending Attachments" Then
                                'we only set the execution date if this is the first time we process this mail, if it is not the first time, we do not change ex date
                                '----------------------------------------------------------------------------
                                update_cr_common_table(db, cr_id, "executors", executors_raw, err)
                                If Not err Like "" Then GoTo get_out
                                update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)
                                If Not err Like "" Then GoTo get_out
                                update_cr_common_table_date(db, cr_id, "execution_date", ex_date, False, err)
                                If Not err Like "" Then GoTo get_out
                            End If
                            combine_cc_list(cc_list, cc_list_old, format, err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table(db, cr_id, "cc_list", cc_list, err)
                            If Not err Like "" Then GoTo get_out
                            add2log(db, time_now, cr_id, If(cr_status Like "Execution Complete Pending Attachments", "CR executed, more execution attachments added (but still missing attachments)", "CR executed (but missing attachments)") & " by " & msg.from.address, err)
                            If Not err Like "" Then GoTo get_out

                            'send autoreply to executor
                            '-------------------------------
                            Dim cr_form_zipped As String = ""
                            check_size_and_zip(file_out, cr_form_zipped, format, local, err)
                            If Not err Like "" Then GoTo get_out

                            Dim subj As String = ""
                            Dim body As String = ""
                            subj = "CRMS: CR Execution Complete But Missing Attachments " & "(" & cr_id & ")"
                            If cr_status Like "Execution Complete Pending Attachments" Then
                                body = "Dear " & msg.from.displayname & ",<P>You are still missing attachments.  " & If(file_out Like "", "You must attach proof the CR was executed, at least 1 attachment is required", "You must attach proof the CR was executed, see returned form for outstanding attachments") & ", please send them ASAP so the CR can be sent for review.<BR>Just reply to this mail with your attachments.<BR>"
                            Else
                                body = "Dear " & msg.from.displayname & ",<P>" & If(file_out Like "", "You need to attach proof the CR was executed, at least 1 attachment is required", "You need to attach proof the CR was executed, see returned form for missing attachments") & ", please send them ASAP so the CR can be sent for review.<BR>Just reply to this mail with your attachments.<BR>"
                            End If
                            body = body & "Please note the following attachment rules:" & array2html_bullet_list({"For Parameter and Hardware CRs, you need at least 1 attachment, any filename", "For RF CRs, attach the site audit photos as a zip file, the filename should contain the site name, eg. KENDALMW"})
                            If Regex.IsMatch(cr_type, "^((RF\sRe-engineering)|(RF Basic))$", RegexOptions.IgnoreCase) Then
                                body = body & "Please find the attached CR error form:<BR>" & Path.GetFileName(file_out)
                            End If
                            body = body & "<P>Thanks<P>"
                            Dim log_s As String = ""
                            body = body & get_log_string(log_s, db, cr_id, err)
                            body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                            auto_reply_general(tx_svr, msg.from.address, "", subj, body, {If(Not cr_form_zipped Like "", cr_form_zipped, file_out)}, err)
                            If Not err Like "" Then GoTo get_out

                        ElseIf attach_ok = "ok" Then
                            'send the CRs and attachments to the requester for review
                            '-------------------------------------------------------------
                            'update the status and add to the log
                            '----------------------------------------------------
                            time_now = Now
                            update_cr_common_table(db, cr_id, "cr_status", "Execution Complete", err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                            If Not err Like "" Then GoTo get_out
                            If Not cr_status Like "Execution Complete Pending Attachments" Then
                                'we only set the execution date if this is the first time we process this mail, if it is not the first time, we do not change ex date
                                '-------------------------------------------------------------------------------------------------------------
                                update_cr_common_table(db, cr_id, "executors", executors_raw, err)
                                If Not err Like "" Then GoTo get_out
                                update_cr_common_table_date(db, cr_id, "execution_date", ex_date, False, err)
                                If Not err Like "" Then GoTo get_out
                            End If
                            combine_cc_list(cc_list, cc_list_old, format, err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table(db, cr_id, "cc_list", cc_list, err)
                            If Not err Like "" Then GoTo get_out
                            add2log(db, time_now, cr_id, "CR executed (attachments ok) by " & msg.from.address, err)
                            If Not err Like "" Then GoTo get_out

                            'get the reviewers email addr
                            '--------------------------------
                            'get the other email address for the cc list
                            '--------------------------------------------------
                            Dim cc_list_final As String = ""
                            Dim user_list As String = ""
                            cc_list = ""            'can reset as it is in the DB now               
                            get_user_emails(format.cc_mask_review_request, user_list, requester, approver, execution_coordinator, executors, cc_list, cr_id, format, db, err)
                            If Not err Like "" Then GoTo get_out
                            Dim to_list As String = requester
                            get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
                            If Not err Like "" Then GoTo get_out

                            'get the attachments, only get here if all attachments are there (or if all cr fail, then there would be no attachments)
                            '---------------------------------------------------------------------------------------------------
                            Dim cr_attach() As String = {}
                            cr_attach = FileIO.FileSystem.GetFiles(Path.GetDirectoryName(cr_form), FileIO.SearchOption.SearchTopLevelOnly, "executor attachments*.zip").ToArray
                            If cr_attach Is Nothing Then cr_attach = {}

                            'send the ack
                            '------------------
                            Dim subj As String = "CRMS: Acknowledgement (" & cr_id & ")"
                            Dim body As String = "Dear " & msg.from.displayname & ",<P>Your executed CR and verification files were recorded and forwarded to " & requester & " for review and acceptance."
                            body = body & "<P>Thanks<P>"
                            body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                            auto_reply_general(tx_svr, msg.from.address, "", subj, body, {}, err)
                            If Not err Like "" Then GoTo get_out

                            'add to log
                            '------------------
                            time_now = Now
                            add2log(db, time_now, cr_id, "CR review request" & If(cr_attach.Count <= 1, "", ", with " & cr_attach.Count & " parts,") & " sent to " & to_list, err)
                            If Not err Like "" Then GoTo get_out

                            'send the review request with the cr form and the first executor attachment, the others will follow if there are any.  
                            'requester attachments are not sent as the requester would have them
                            '-----------------------------------------------------------------------------
                            Dim cr_form_zipped As String = ""
                            check_size_and_zip(cr_form, cr_form_zipped, format, local, err)
                            If Not err Like "" Then GoTo get_out

                            subj = "CRMS: CR Review Request (" & cr_id & ")"
                            Dim user_name() As String = get_user_name("requester", cr_id, format, db)
                            body = "Dear " & If(user_name.Count > 0 AndAlso Not user_name.First Like "", Join(user_name, ", "), "Requester") & ",<P>Please review the attached CR form which has been executed." & array2html_bullet_list({Path.GetFileName(cr_form)})
                            body = body & "<BR>Acceptable Responses:<BR>"
                            body = body & "-----------------------------------------------------"
                            body = body & array2html_bullet_list({"You have no issues => reply to the email and type 'ok/yes/accepted' in the email body.", "You have issues => reply to the email and type 'not ok/nok/no/not accepted' in the email body.<BR>You can add any text after this for your reason or notes."})
                            body = body & "<P>Thanks<P>"
                            Dim log_s As String = ""
                            body = body & get_log_string(log_s, db, cr_id, err)
                            body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                            If cr_attach.Count = 0 Then
                                auto_reply_general(tx_svr, to_list, cc_list_final, subj, body, {If(Not cr_form_zipped Like "", cr_form_zipped, cr_form)}, err)
                                If Not err Like "" Then GoTo get_out
                            Else
                                Dim j As Integer = 0
                                For Each item In cr_attach
                                    If j = 0 Then
                                        auto_reply_general(tx_svr, to_list, cc_list_final, subj & If(cr_attach.Count = 1, "", " => part " & j + 1 & " of " & cr_attach.Count), body, {If(Not cr_form_zipped Like "", cr_form_zipped, cr_form), item}, err)
                                    Else
                                        auto_reply_general(tx_svr, to_list, cc_list_final, subj & " => part " & j + 1 & " of " & cr_attach.Count, body, {item}, err)
                                    End If
                                    If Not err Like "" Then GoTo get_out
                                    j += 1
                                Next
                            End If

                            'update the status and add to the log
                            '--------------------------------------
                            update_cr_common_table(db, cr_id, "cr_status", "Pending Review", err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                            If Not err Like "" Then GoTo get_out
                            update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)
                            If Not err Like "" Then GoTo get_out

                        End If
                    Catch ex As Exception
                        err = "ER: Error processing the executor reply mail: " & ex.ToString
                        GoTo get_out
                    End Try
                    '###########################################################################################################################
                    'End of CR execution results processing
                    '###########################################################################################################################

                ElseIf Regex.IsMatch(cleanbody, "^((nok)|(no)|(tolak)|(not\sok)|(fuck\soff))", RegexOptions.IgnoreCase) Then
                    Try
                        'create the resub form in the cr dir and delete any attachments
                        '-----------------------------------------------------------
                        cr_form_temp = ""
                        Dim cr_resub_form As String = ""
                        pre_process_for_resub(cr_resub_form, cr_id, cr_form, local, err)
                        If Not err Like "" Then GoTo get_out
                        create_resubmit_form(cr_resub_form, cr_form_Type, format, local, err)
                        If Not err Like "" Then GoTo get_out

                        'ok, it is good, so update the status and add to the log
                        '----------------------------------------------------
                        time_now = Now
                        update_cr_common_table(db, cr_id, "cr_status", "Execution Rejected", err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        combine_cc_list(cc_list, cc_list_old, format, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table(db, cr_id, "cc_list", cc_list, err)
                        If Not err Like "" Then GoTo get_out
                        add2log(db, time_now, cr_id, "CR execution rejected by the executor: " & msg.from.address, err)
                        If Not err Like "" Then GoTo get_out

                        'send a note to the requester with the resub form
                        '------------------------------------------------
                        'get the other email address for the cc list
                        '--------------------------------------------------
                        Dim cc_list_final As String = ""
                        Dim user_list As String = ""
                        cc_list = ""            'can reset as it is in the DB now              
                        get_user_emails(format.cc_mask_resubmission_request, user_list, requester, approver, execution_coordinator, executors, cc_list, cr_id, format, db, err)
                        If Not err Like "" Then GoTo get_out
                        Dim to_list As String = requester
                        get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
                        If Not err Like "" Then GoTo get_out

                        'send the ack
                        '------------------
                        Dim subj As String = "CRMS: Acknowledgement (" & cr_id & ")"
                        Dim body As String = "Dear " & msg.from.displayname & ",<P>Your reply was recorded, the CR has been sent back to " & requester & " for resubmission."
                        body = body & "<P>Thanks<P>"
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                        auto_reply_general(tx_svr, msg.from.address, "", subj, body, {}, err)
                        If Not err Like "" Then GoTo get_out

                        'add to log
                        '------------------
                        time_now = Now
                        add2log(db, time_now, cr_id, "CR execution rejection note sent to: " & to_list, err)
                        If Not err Like "" Then GoTo get_out

                        'send
                        '-------
                        Dim cr_form_zipped As String = ""
                        check_size_and_zip(cr_resub_form, cr_form_zipped, format, local, err)
                        If Not err Like "" Then GoTo get_out

                        subj = "CRMS: CR Resubmission Request (" & cr_id & ") - !!Execution Rejection!!"
                        Dim user_name() As String = get_user_name("requester", cr_id, format, db)
                        body = "Dear " & If(user_name.Count > 0 AndAlso Not user_name.First Like "", Join(user_name, ", "), "Requester") & ",<P>Please NOTE!  The CR sent in for execution by (" & msg.from.address & "), was not able to be executed." & array2html_bullet_list({Path.GetFileName(cr_form)})
                        body = body & "<BR>Please review, make any changes to the resubmit form and resubmit it to proceed with the CR."
                        body = body & "<P>Thanks<P>"
                        Dim log_s As String = ""
                        body = body & get_log_string(log_s, db, cr_id, err)
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                        auto_reply_general(tx_svr, to_list, cc_list_final, subj, body, {If(Not cr_form_zipped Like "", cr_form_zipped, cr_resub_form)}, err)
                        If Not err Like "" Then GoTo get_out

                        'delete the resubmit form as we do not need it any more
                        '--------------------------------------------------
                        FileIO.FileSystem.DeleteFile(cr_resub_form)

                        'ok, it is good, so update the status and add to the log
                        '----------------------------------------------------
                        update_cr_common_table(db, cr_id, "cr_status", "Pending Resubmission", err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "approval_date", time_now, True, err)        'reset value to null
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "planned_execution_date", time_now, True, err)        'reset value to null
                        If Not err Like "" Then GoTo get_out

                    Catch ex As Exception
                        If err = "" Then
                            err = "ER: Error processing the ex reply mail: " & ex.ToString
                        End If
                        GoTo get_out
                    End Try
                    '###########################################################################################################################
                    'End of CR rejected execution processing
                    '###########################################################################################################################

                Else
                    err = "XREJ: Sorry, I do not undestand the body text."
                    GoTo get_out
                End If

            ElseIf Regex.IsMatch(cleansubject, "(CRMS:\sCR\sReview\sRequest)(\s*)(\()(.*)(\))", RegexOptions.IgnoreCase) Then
                'this is the case where the requester has reviewed the execution results with a yes/no response."
                'check the cr_id is in the db and get the cr_form_type(no need to check it if the cr_id exists)
                '------------------------------------------------------------------------------
                Dim cr_id As String = Regex.Replace(cleansubject, "^(.*)(CRMS:\sCR\sReview\sRequest)(\s*)(\()", "", RegexOptions.IgnoreCase)
                cr_id = Regex.Replace(cr_id, "(\))(.*)$", "", RegexOptions.IgnoreCase)
                If cr_id Is Nothing Or cr_id = "" Or Len(cr_id) < 5 Then
                    err = "REVREJ: Could not read the cr_id from your mail subject => (" & cleansubject & ")"
                    GoTo get_out
                End If
                Dim input_file As String = ""
                Dim cr_form = local.base_path & local.cr & "\" & cr_id & "\" & cr_id & ".xlsb"
                Dim cr_form_temp As String = local.base_path & local.inbox & "\" & cr_id & ".xlsb"
                Dim cr_status As String = ""
                Dim cr_type As String = ""
                Dim cr_type_short As String = ""
                Dim cr_form_Type As String = ""
                Dim cc_list_old As String = ""
                Dim requester As String = ""
                Dim approver As String = ""
                Dim execution_coordinator As String = ""
                Dim executors As String = ""
                get_cr_id_data(cr_id, cr_status, cr_type, cr_type_short, cr_form_Type, cc_list_old, requester, approver, execution_coordinator, executors, db, err)
                If Not err Like "" Then GoTo get_out
                If cr_type Like "" Then
                    err = "XREJ: The cr_id you gave doesn't exist, sorry. (" & cleansubject & ")"
                    GoTo get_out
                End If

                'validate email is from allowed user for this change
                '------------------------------------------------------
                If Not msg.from.address Like c2e(requester) Then
                    err = "XREJ: Your email doesn't have rights to perform this action, sorry."
                    GoTo get_out
                End If

                'state check
                '--------------
                If Not cr_status = "Pending Review" Then      'main state
                    If Not (cr_status = "Execution Complete" Or cr_status = "Review Failed") Then 'intermediate stuck states from email error, if it is closed, then leave it
                        err = "REVREJ: This CR can only be reviewed while it has the status 'Pending Review', the current status is: '" & cr_status & "', sorry."
                        GoTo get_out
                    End If
                End If

                If Regex.IsMatch(cleanbody, "^((ok)|(closed)|(yes)|(terima))", RegexOptions.IgnoreCase) Then
                    If Not FileIO.FileSystem.FileExists(cr_form) Then
                        err = "ER: Server HDD error, can't find cr form on the server"
                        GoTo get_out
                    End If

                    'does the final formatting on the cr_form
                    '------------------------------------------
                    time_now = Now
                    finalise_cr_form(cr_form, cr_form_Type, time_now, format, local, err)
                    If Not err Like "" Then GoTo get_out

                    'combine the cc_lists
                    '----------------------------
                    combine_cc_list(cc_list, cc_list_old, format, err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table(db, cr_id, "cc_list", cc_list, err)
                    If Not err Like "" Then GoTo get_out

                    Try
                        'get the other email address for the cc list
                        '--------------------------------------------------
                        Dim cc_list_final As String = ""
                        Dim user_list As String = ""
                        cc_list = ""            'can reset as it is in the DB now             
                        get_user_emails(format.cc_mask_everyone, user_list, requester, approver, execution_coordinator, executors, cc_list, cr_id, format, db, err)
                        If Not err Like "" Then GoTo get_out
                        If executors Like "" Then
                            executors = msg.from.address
                        End If
                        Dim to_list As String = executors
                        get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
                        If Not err Like "" Then GoTo get_out

                        'send the ack
                        '------------------
                        Dim subj As String = "CRMS: Acknowledgement (" & cr_id & ")"
                        Dim body As String = "Dear " & msg.from.displayname & ",<P>Your reply was recorded.  The CR has been closed."
                        body = body & "<P>Thanks<P>"
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                        auto_reply_general(tx_svr, msg.from.address, "", subj, body, {}, err)
                        If Not err Like "" Then GoTo get_out

                        'add to log
                        '---------------------
                        add2log(db, time_now, cr_id, "CR closed note sent to: " & to_list, err)
                        If Not err Like "" Then GoTo get_out

                        'send the CR closed note
                        '--------------------------
                        Dim cr_form_zipped As String = ""
                        check_size_and_zip(cr_form, cr_form_zipped, format, local, err)
                        If Not err Like "" Then GoTo get_out

                        subj = "CRMS: CR Closed Note (" & cr_id & ")"
                        body = "The CR (" & cr_id & ") has been closed.  Please find the the attached CR form:" & array2html_bullet_list({Path.GetFileName(cr_form)})
                        body = body & "<P>Thanks<P>"
                        Dim log_s As String = ""
                        body = body & get_log_string(log_s, db, cr_id, err)
                        body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                        auto_reply_general(tx_svr, to_list, cc_list_final, subj, body, {If(Not cr_form_zipped Like "", cr_form_zipped, cr_form)}, err)
                        If Not err Like "" Then GoTo get_out

                        'ok, it is good, so update the status and add to the log
                        '----------------------------------------------------
                        update_cr_common_table(db, cr_id, "cr_status", "Closed", err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)
                        If Not err Like "" Then GoTo get_out
                        update_cr_common_table_date(db, cr_id, "closed_date", time_now, False, err)
                        If Not err Like "" Then GoTo get_out
                    Catch ex As Exception
                        If err = "" Then
                            err = "ER: Error processing the reviewer reply mail: " & ex.ToString
                        End If
                        GoTo get_out
                    End Try
                    '###########################################################################################################################
                    'End of CR review success processing
                    '###########################################################################################################################

                ElseIf Regex.IsMatch(cleanbody, "^((nok)|(not\sclosed)|(no)|(tolak)|(fuck\soff))", RegexOptions.IgnoreCase) Then
                    'delete the executor attachments if there are any as they need to resubmit, the actual cr_form will get overwritten so leave it there.
                    '---------------------------------------------------------------------------------------------------------
                    For Each item In FileIO.FileSystem.GetFiles(local.base_path & local.cr & "\" & cr_id, FileIO.SearchOption.SearchTopLevelOnly, "executor attachments*.zip")
                        force_delete_file(item, err)
                        If Not err Like "" Then GoTo get_out
                    Next

                    'put the CR to review failed state
                    'update the status and add to the log
                    '--------------------------------------
                    time_now = Now
                    update_cr_common_table(db, cr_id, "cr_status", "Review Failed", err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                    If Not err Like "" Then GoTo get_out
                    combine_cc_list(cc_list, cc_list_old, format, err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table(db, cr_id, "cc_list", cc_list, err)
                    If Not err Like "" Then GoTo get_out
                    add2log(db, time_now, cr_id, "CR review failed by " & msg.from.address, err)
                    If Not err Like "" Then GoTo get_out

                    'get email address 
                    '---------------------
                    Dim cc_list_final As String = ""
                    Dim user_list As String = ""
                    cc_list = ""            'can reset as it is in the DB now  
                    get_user_emails(format.cc_mask_execution_request, user_list, requester, approver, execution_coordinator, executors, cc_list, cr_id, format, db, err)
                    If Not err Like "" Then GoTo get_out
                    If executors Like "" Then
                        executors = msg.from.address
                    End If
                    Dim to_list As String = executors
                    get_cc_final_lists(cc_list, to_list, cc_list_final, format, err)
                    If Not err Like "" Then GoTo get_out

                    'send the ack
                    '------------------
                    time_now = Now
                    Dim subj As String = "CRMS: Acknowledgement (" & cr_id & ")"
                    Dim body As String = "Dear " & msg.from.displayname & ",<P>Your reply was recorded.  The CR was forwarded to " & executors & " for follow up."
                    body = body & "<P>Thanks<P>"
                    body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                    auto_reply_general(tx_svr, msg.from.address, "", subj, body, {}, err)
                    If Not err Like "" Then GoTo get_out

                    'add to log
                    '---------------------
                    add2log(db, time_now, cr_id, "CR review failure note sent to: " & to_list, err)
                    If Not err Like "" Then GoTo get_out

                    'send the execution reject note
                    '-----------------------------
                    subj = "CRMS: CR Execution Request (" & cr_id & ") - !!Review Failed!!"
                    Dim user_name() As String = get_user_name("executors", cr_id, format, db)
                    body = "Dear " & If(user_name.Count > 0 AndAlso Not user_name.First Like "", Join(user_name, ", "), "Executors") & ",<P>Please NOTE!  The CR (" & cr_id & ") sent in for review by (" & msg.from.address & "), was not accepted.<BR>"
                    body = body & "Please review, rectify the issues and send the correct information back for further review."
                    body = body & "<P>Thanks<P>"
                    Dim log_s As String = ""
                    body = body & get_log_string(log_s, db, cr_id, err)
                    body = body & "<P>.<P>On " & msg.msg_date.ToLongDateString & " at " & msg.msg_date.ToLongTimeString & ", " & msg.from.displayname & " &lt;" & msg.from.address & "&gt; wrote:<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw)
                    auto_reply_general(tx_svr, to_list, cc_list_final, subj, body, {}, err)
                    If Not err Like "" Then GoTo get_out

                    'ok, it is good, so update the status and add to the log
                    '----------------------------------------------------
                    update_cr_common_table(db, cr_id, "cr_status", "Pending Execution", err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table_date(db, cr_id, "last_activity_date", time_now, False, err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table_date(db, cr_id, "last_nag_date", time_now, True, err)
                    If Not err Like "" Then GoTo get_out
                    update_cr_common_table_date(db, cr_id, "execution_date", time_now, True, err)            'this sets the execution date to dbnull (second param from right = True)
                    If Not err Like "" Then GoTo get_out

                Else
                    err = "XREJ: Sorry, I do not understand the text in your email body."
                    GoTo get_out
                End If
                '###########################################################################################################################
                'End of CR review failure processing
                '###########################################################################################################################

            ElseIf Regex.IsMatch(cleansubject, "^((mod\suser)|(add\suser)|(add\sexecutor)|(add\santenna))", RegexOptions.IgnoreCase) Then
                Dim add_ex_flag As Boolean = False
                If Regex.IsMatch(cleansubject, "^add\sexecutor", RegexOptions.IgnoreCase) Then add_ex_flag = True
                Dim add_ant_flag As Boolean = False
                If Regex.IsMatch(cleansubject, "^add\santenna", RegexOptions.IgnoreCase) Then add_ant_flag = True
                Dim mod_user_flag As Boolean = False
                If Regex.IsMatch(cleansubject, "^mod\suser", RegexOptions.IgnoreCase) Then mod_user_flag = True
                Dim add_user_flag As Boolean = False
                If Regex.IsMatch(cleansubject, "^add\suser", RegexOptions.IgnoreCase) Then add_user_flag = True

                'this is for the admin to add users to the DB
                '--------------------------------
                'check it is from admin
                '-------------------------------
                clear_dt(dt)
                If Not add_ex_flag And Not add_ant_flag Then
                    sqlquery(False, db, "SELECT email FROM " & db.schema & ".people WHERE administrator='1' AND email='" & msg.from.address & "';", dt, err)
                Else
                    sqlquery(False, db, "SELECT email FROM " & db.schema & ".people WHERE (rfb_ex_coord='1' OR rfr_ex_coord='1' OR hdw_ex_coord='1' OR prm_ex_coord='1' OR administrator='1') AND email='" & msg.from.address & "';", dt, err)
                End If
                If Not err Like "" Then GoTo get_out
                If dt.Rows.Count = 0 Then
                    err = "ADDREJ: Your email address doesn't have the rights to add users, sorry."
                    GoTo get_out
                End If

                'basic attachement check
                '----------------------------
                If msg.attachments.Count = 0 Then
                    err = "XREJ: There must be ata least 1 .txt attachment to process your mail."
                    GoTo get_out
                End If
                check_attachments(msg, err)
                If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then
                    GoTo get_out
                ElseIf Not err Like "" Then
                    err = "ADDREJ:  " & err
                    GoTo get_out
                End If

                Dim dir_out As String = local.base_path & local.inbox & "\attachments"
                If FileIO.FileSystem.DirectoryExists(dir_out) Then
                    clean_dir(dir_out, err)
                    If Not err Like "" Then GoTo get_out
                Else
                    FileIO.FileSystem.CreateDirectory(dir_out)
                End If

                'get the data from the .txt file
                '------------------------------------
                Dim data() As String = {}
                Try
                    'get the sql array from the txt file, only looks in the first .txt file found
                    '----------------------------------------------
                    Dim file As String = ""
                    If FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.txt").Count > 0 Then
                        file = FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.txt").First
                    End If
                    data = get_data_from_txt(file, format, err)
                    If data.Count = 0 Then
                        err = "XREJ: Email error, no data was found in your file, giving up...."
                        GoTo get_out
                    End If
                Catch ex As Exception
                    err = "ER: Error getting attachments, details: " & ex.ToString
                    GoTo get_out
                End Try
                'at this point the data array has all the commands, whatever they are

                'processes the data
                '-----------------------
                Dim users_added() As String = {}
                Dim users_string As String = ""
                Dim ants_added() As String = {}
                Dim ant_string As String = ""
                If Not add_ant_flag Then
                    Try
                        'get the DB people/antenna table
                        '-------------------------
                        Dim people_dt As New System.Data.DataTable
                        sqlquery(False, db, "select * from " & db.schema & ".people;", people_dt, err)
                        If Not err Like "" Then GoTo get_out

                        'filter the users array to only those that are no in the DB and have valid email addresses and other valid input
                        '----------------------------------------------------------------------------------------
                        users_string = ""
                        For Each item In data
                            'takes out the terminating semicolon
                            '---------------------------------------------------
                            If Not Regex.IsMatch(Trim(item), ";$") Then GoTo skip_user
                            item = Regex.Replace(Trim(item), ";$", "")

                            'checks the format of the user string first, skips if it is no good
                            '---------------------------------------------------
                            If item Like "" Then GoTo skip_user
                            Dim item_split() As String = Regex.Split(item, ",|=>")
                            For i = 0 To item_split.Count - 1
                                item_split(i) = Trim(item_split(i))
                                If item_split(i) = "" Then GoTo skip_user
                            Next

                            'check the data first
                            '--------------------------
                            If add_ex_flag Then
                                If Not item_split.Count = 2 Then GoTo skip_user
                                item_split(0) = Regex.Replace(item_split(0), "((\()|(\))|(,))", "_")        'gets rid of brackets and commas as it makes things messy
                                If Not format.IsValidEmail(Trim(item_split(1))) Then GoTo skip_user

                            ElseIf mod_user_flag Then
                                If Not item_split.Count = 4 Then GoTo skip_user
                                item_split(0) = Regex.Replace(item_split(0), "((\()|(\))|(,))", "_")        'gets rid of brackets and commas as it makes things messy
                                item_split(2) = Regex.Replace(item_split(2), "((\()|(\))|(,))", "_")        'gets rid of brackets and commas as it makes things messy
                                If Not format.IsValidEmail(Trim(item_split(1))) Then GoTo skip_user
                                If Not format.IsValidEmail(Trim(item_split(3))) Then GoTo skip_user

                            ElseIf add_user_flag Then
                                If Not item_split.Count = 11 Then GoTo skip_user
                                item_split(0) = Regex.Replace(item_split(0), "((\()|(\))|(,))", "_")        'gets rid of brackets and commas as it makes things messy
                                For i = 2 To 10
                                    If Not Regex.IsMatch(Trim(item_split(i)), "^[01]$") Then GoTo skip_user
                                Next
                                If Not format.IsValidEmail(Trim(item_split(1))) Then GoTo skip_user
                            End If

                            'do the DB updates
                            '------------------
                            Try
                                If add_user_flag Or mod_user_flag Then
                                    Dim query = From person In people_dt.AsEnumerable
                                                Where person("name").ToString Like Trim(item_split(0)) And person("email").ToString Like Trim(item_split(1))
                                                Select person
                                    If query.Count = 0 And mod_user_flag Then
                                        users_string = users_string & "ERROR! => old name/email doesn't exist: " & item_split(0) & "(" & item_split(1) & ") => " & item_split(2) & "(" & item_split(3) & "),"
                                        GoTo skip_user
                                    ElseIf query.Count > 0 Then
                                        'ok, it is good, so update the DB and update the log
                                        '------------------------------------------------
                                        clear_dt(dt)
                                        Dim sqltext As String = ""
                                        If add_user_flag Then
                                            sqltext = "UPDATE `" & db.schema & "`.`people` SET `requester`='" & item_split(2) & "', `approver`='" & item_split(3) & "', `rfb_ex_coord`='" & item_split(4) & "', `rfr_ex_coord`='" & item_split(5) & "', `hdw_ex_coord`='" & item_split(6) & "', `prm_ex_coord`='" & item_split(7) & "', `executor`='" & item_split(8) & "', `query`='" & item_split(9) & "', `anyquery`='0', `administrator`='" & item_split(10) & "' WHERE `name`='" & item_split(0) & "' AND `email`='" & item_split(1) & "';"

                                        ElseIf mod_user_flag Then
                                            Dim query2 = From person In people_dt.AsEnumerable
                                                        Where person("name").ToString Like Trim(item_split(2)) And person("email").ToString Like Trim(item_split(3))
                                                        Select person
                                            If query2.Count = 0 Then
                                                sqltext = "UPDATE `" & db.schema & "`.`people` SET `name`='" & item_split(2) & "', `email`='" & item_split(3) & "' WHERE `name`='" & item_split(0) & "' AND `email`='" & item_split(1) & "';"

                                            Else
                                                users_string = users_string & "ERROR! => new name/email already exists: " & item_split(0) & "(" & item_split(1) & ") => " & item_split(2) & "(" & item_split(3) & "),"
                                                GoTo skip_user
                                            End If
                                        End If
                                        sqlquery(False, db, sqltext, dt, err)
                                        If Not err Like "" Then GoTo skip_user
                                        users_string = users_string & If(add_user_flag, item_split(0) & "(" & item_split(1) & ")", item_split(0) & "(" & item_split(1) & ") => " & item_split(2) & "(" & item_split(3) & ")") & " - user updated,"

                                        'add to the log
                                        '---------------------
                                        time_now = Now
                                        add2log(db, time_now, "", "User updated by " & msg.from.address & ".  User: " & If(add_user_flag, item_split(0) & "(" & item_split(1) & ")", item_split(0) & "(" & item_split(1) & ") => " & item_split(2) & "(" & item_split(3) & ")"), err)
                                        If Not err Like "" Then GoTo get_out

                                        'now skip as we already updated, there will be no insert for this user
                                        GoTo skip_user
                                    End If
                                End If
                            Catch ex As Exception
                                GoTo skip_user
                            End Try

                            'do the DB inserts
                            '--------------------
                            Try
                                If add_user_flag Or add_ex_flag Then
                                    Dim query = From person In people_dt.AsEnumerable
                                                Where person("name").ToString Like Trim(item_split(0)) And person("email").ToString Like Trim(item_split(1))
                                                Select person
                                    If query.Count > 0 Then
                                        users_string = users_string & "ERROR! => new name/email already exists: " & item_split(0) & "(" & item_split(1) & "),"
                                        GoTo skip_user
                                    ElseIf query.Count = 0 Then
                                        'ok, it is good, so update the DB and update the log
                                        '------------------------------------------------
                                        clear_dt(dt)
                                        Dim sqltext As String = ""
                                        If add_user_flag Then
                                            sqltext = "INSERT INTO " & db.schema & ".people " & _
                                                        "(name, email, requester, approver, rfb_ex_coord, rfr_ex_coord, hdw_ex_coord, prm_ex_coord, executor, query, administrator) " & _
                                                        "VALUES ('" & item_split(0) & "','" & item_split(1) & "','" & item_split(2) & "', '" & item_split(3) & "', '" & item_split(4) & "', '" & item_split(5) & "','" & item_split(6) & "', '" & item_split(7) & "', '" & item_split(8) & "', '" & item_split(9) & "', '" & item_split(10) & "');"
                                        ElseIf add_ex_flag Then
                                            sqltext = "INSERT INTO " & db.schema & ".people " & _
                                                        "(name, email, requester, approver, rfb_ex_coord, rfr_ex_coord, hdw_ex_coord, prm_ex_coord, executor, query, administrator) " & _
                                                        "VALUES ('" & item_split(0) & "','" & item_split(1) & "','0','0','0','0','0','0','1','0','0');"
                                        End If
                                        sqlquery(False, db, sqltext, dt, err)
                                        If Not err Like "" Then GoTo skip_user
                                        users_string = users_string & item_split(0) & " (" & item_split(1) & ") - user added,"

                                        'add to the log
                                        '---------------------
                                        time_now = Now
                                        add2log(db, time_now, "", "New " & If(add_ex_flag, "executor", "user") & " added by " & msg.from.address & ".  New " & If(add_ex_flag, "Executor", "User") & ": " & item_split(0) & "(" & item_split(1) & ")", err)
                                        If Not err Like "" Then GoTo get_out
                                    End If
                                End If
                            Catch ex As Exception
                                GoTo skip_user
                            End Try
skip_user:
                        Next
                        people_dt.Dispose()
                        If users_string.Length > 0 Then
                            users_string = Left(users_string, Len(users_string) - 1)
                            users_added = Strings.Split(users_string, ",").Distinct.ToArray
                        End If
skip_no_users:
                    Catch ex As Exception
                        err = "ER: some error adding users, details: " & ex.ToString
                        GoTo get_out
                    End Try

                ElseIf add_ant_flag Then
                    'get the antennas from the body text
                    '------------------------------------
                    Try
                        'get the DB antenna table
                        '-------------------------
                        Dim ant_dt As New System.Data.DataTable
                        sqlquery(False, db, "select * from " & db.schema & ".antennas;", ant_dt, err)
                        If Not err Like "" Then GoTo get_out

                        'validate the input values
                        '-------------------------------
                        ant_string = ""
                        For Each row In data
                            'takes out the terminating semicolon
                            '---------------------------------------------------
                            If Not Regex.IsMatch(Trim(row), ";$") Then GoTo skip_ant
                            row = Regex.Replace(Trim(row), ";$", "")

                            'checks the format of the user string first, skips if it is no good
                            '---------------------------------------------------
                            If row Like "" Then GoTo skip_ant
                            Dim row_split() As String = Strings.Split(row, ",")

                            If Not row_split.Count = 13 Then GoTo skip_ant
                            Dim binary_list() As Integer = {2, 3, 4, 7, 11}
                            For Each index In binary_list
                                If Not Regex.IsMatch(Trim(row_split(index)), "^[01]$") Then GoTo skip_ant
                            Next

                            Dim query = From antenna In ant_dt.AsEnumerable
                                        Where antenna("antenna").ToString Like Trim(row_split(0))
                                        Select antenna
                            If query.Count > 0 Then GoTo skip_ant

                            'ok, it is good, so update the DB and update the log
                            '------------------------------------------------
                            clear_dt(dt)
                            Dim sqltext As New StringBuilder("")
                            Dim temp_string As New StringBuilder("")
                            sqltext.Append("INSERT INTO `" & db.schema & "`.`antennas` (`antenna`, `manufacturer`, `900`, `1800`, `2100`, `edt_min`, `edt_max`, `has_mdt`, `hbw`, `vbw`, `gain_dbi`, `dual_beam`, `comment`) ")
                            sqltext.Append("VALUES (")
                            For Each item In row_split
                                sqltext.Append("'" & item & "',")
                                temp_string.Append(item & "-")
                            Next
                            sqltext.Replace(",", ");", sqltext.Length - 1, 1)
                            sqlquery(False, db, sqltext.ToString, dt, err)
                            If Not err Like "" Then GoTo skip_ant
                            temp_string.Replace("-", ",", temp_string.Length - 1, 1)
                            ant_string = ant_string & temp_string.ToString

                            'add to the log
                            '---------------------
                            time_now = Now
                            add2log(db, time_now, "", "New antenna added by " & msg.from.address & ".  New antenna: " & temp_string.ToString, err)
                            If Not err Like "" Then GoTo get_out
skip_ant:
                        Next
                        ant_dt.Dispose()
                        If ant_string.Length > 0 Then
                            ant_string = Left(ant_string, Len(ant_string) - 1)
                            ants_added = Strings.Split(ant_string, ",").Distinct.ToArray
                        End If
skip_no_ants:
                    Catch ex As Exception
                        err = "ER: some error adding antennas, details: " & ex.ToString
                        GoTo get_out
                    End Try
                End If


                'go through .xlsb files and process if they are indeed cr forms, output will be cr_attach array
                '---------------------------------------------------------------------
                Dim cr_attach() As String = {}
                Try
                    Dim s_cr_attach As String = ""
                    Dim files() As String = IO.Directory.GetFiles(local.base_path & local.inbox, "*.xlsb", IO.SearchOption.TopDirectoryOnly)
                    If files.Count = 0 Then
                        GoTo skip_attach_user
                    End If
                    For Each item In files
                        allowed_values_ds2xl(item, "", format, local, err)
                        If Not err Like "" Then GoTo get_out
                        s_cr_attach = s_cr_attach & item & ","
                    Next
                    If Len(s_cr_attach) > 1 Then s_cr_attach = Left(s_cr_attach, Len(s_cr_attach) - 1)
                    cr_attach = Strings.Split(s_cr_attach, ",")
                    If s_cr_attach Like "" Then
                        cr_attach = {}
                    Else
                        For Each item In cr_attach
                            FileIO.FileSystem.MoveFile(item, dir_out & "\" & Path.GetFileName(item))
                        Next
                        For Each item In FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.zip").ToArray
                            force_delete_file(item, err)
                            If Not err Like "" Then GoTo get_out
                        Next
                        zip_dir(dir_out, 10, 20, local.base_path & local.inbox & "\" & "attachments.zip", err)
                    End If
                Catch ex As Exception
                    GoTo skip_attach_user
                End Try
skip_attach_user:

                Try
                    'Send the ack
                    '---------------------
                    Dim attach() As String = FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.zip").ToArray
                    Dim cr_attach_name() = (From item In cr_attach
                                            Let a = Path.GetFileName(item)
                                            Select a).ToArray
                    Dim subj As String = "CRMS: Processing Completed"
                    Dim body As New StringBuilder("")
                    If Not add_ant_flag Then
                        If users_added.Count > 0 Then
                            body.Append("The following " & If(add_ex_flag, "executors", "users") & " were added to(updated in) the DB:")
                            body.Append(array2html_bullet_list(users_added))
                        Else
                            body.Append("There were no " & If(add_ex_flag, "executors", "users") & " added(updated).<BR>")
                        End If

                    ElseIf add_ant_flag Then
                        If ants_added.Count > 0 Then
                            body.Append("The following antennas were added to the DB:")
                            body.Append(array2html_bullet_list(ants_added))
                        Else
                            body.Append("There were no antennas added.<BR>")
                        End If

                    End If
                    If cr_attach.Count > 0 Then
                        body.Append("<BR>The attached cr forms have been updated with the current allowed values:")
                        body.Append(array2html_bullet_list(cr_attach_name))
                    Else
                        body.Append("<BR>There were no forms to attach.")
                    End If

                    body.Append("<BR>Thanks<BR>")
                    body.Append("<BR>-----------------------------------------------------------------------------------------------------------<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw))
                    auto_reply_general(tx_svr, msg.from.address, "", subj, body.ToString, attach, err)
                    If Not err Like "" Then GoTo get_out

                Catch ex As Exception
                    If err = "" Then
                        err = "ER: Error processing add " & If(add_ant_flag, "antennas", If(add_ex_flag, "executors", "users")) & ", details: " & ex.ToString
                    End If
                    GoTo get_out
                End Try
                '###########################################################################################################################
                'End of add user/antenna processing
                '###########################################################################################################################

            ElseIf Regex.IsMatch(cleansubject, "^update\sform", RegexOptions.IgnoreCase) Then
                'this is for anyone to update the allowed values in the cr form allowed values sheet
                '------------------------------------------------------------------------------------
                'check it is from an requester, execution coordinator, executors or admin
                '--------------------------------------------------------------
                clear_dt(dt)
                sqlquery(False, db, "SELECT email FROM " & db.schema & ".people WHERE (rfb_ex_coord='1' OR rfr_ex_coord='1' OR hdw_ex_coord='1' OR prm_ex_coord='1' OR executor='1' OR requester ='1' OR administrator='1') AND email='" & msg.from.address & "';", dt, err)
                If Not err Like "" Then GoTo get_out
                If dt.Rows.Count = 0 Then
                    err = "UPDATEREJ: Your email address doesn't have the rights to update the cr form, sorry."
                    GoTo get_out
                End If

                'get the cr forms from the attachments
                '------------------------------------
                Try
                    'basic attachement check
                    '----------------------------
                    If msg.attachments.Count = 0 Then
                        err = "UPDATEREJ: Your email doesn't have any attachments.<BR>Thanks"
                        GoTo get_out
                    End If
                    check_attachments(msg, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then
                        GoTo get_out
                    ElseIf Not err Like "" Then
                        err = "UPDATEREJ:  " & err
                        GoTo get_out
                    End If

                    'go through .xlsb/m files and process if they are indeed cr forms, output will be cr_attach array
                    '---------------------------------------------------------------------
                    Dim dir_out As String = local.base_path & local.inbox & "\attachments"
                    If FileIO.FileSystem.DirectoryExists(dir_out) Then
                        clean_dir(dir_out, err)
                        If Not err Like "" Then GoTo get_out
                    Else
                        FileIO.FileSystem.CreateDirectory(dir_out)
                    End If

                    Dim s_cr_attach As String = ""
                    Dim files() As String = IO.Directory.GetFiles(local.base_path & local.inbox, "*.xlsb", IO.SearchOption.TopDirectoryOnly)
                    If files.Count = 0 Then
                        err = "UPDATEREJ: Your mail has no form attachments."
                        GoTo get_out
                    End If

                    'update the allowed value tables first
                    '---------------------------------------
                    allowed_values_db2ds(db, format, err)
                    If Not err Like "" Then GoTo get_out

                    'update the files
                    '----------------------
                    For Each item In files
                        allowed_values_ds2xl(item, "", format, local, err)
                        If Not err Like "" Then GoTo get_out
                        s_cr_attach = s_cr_attach & item & ","
                    Next
                    If Len(s_cr_attach) > 1 Then s_cr_attach = Left(s_cr_attach, Len(s_cr_attach) - 1)
                    Dim cr_attach() As String = Strings.Split(s_cr_attach, ",")
                    If s_cr_attach Like "" Then
                        cr_attach = {}
                    Else
                        For Each item In cr_attach
                            FileIO.FileSystem.MoveFile(item, dir_out & "\" & Path.GetFileName(item))
                        Next
                        For Each item In FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.zip").ToArray
                            force_delete_file(item, err)
                            If Not err Like "" Then GoTo get_out
                        Next
                        zip_dir(dir_out, 10, 20, local.base_path & local.inbox & "\" & "attachments.zip", err)
                    End If

                    'Send the ack
                    '---------------------
                    Dim attach() As String = FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.zip").ToArray
                    Dim cr_attach_name() = (From item In cr_attach
                                        Let a = Path.GetFileName(item)
                                        Select a).ToArray
                    Dim subj As String = "CRMS: Allowed Values Updated"
                    Dim body As New StringBuilder("")
                    If cr_attach.Count > 0 Then
                        body.Append("The attached cr forms have been updated with the current allowed values:")
                        body.Append(array2html_bullet_list(cr_attach_name))
                    Else
                        body.Append("There were no forms to attach.")
                    End If
                    body.Append("<BR>Thanks<BR>")
                    body.Append("<BR>-----------------------------------------------------------------------------------------------------------<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw))
                    auto_reply_general(tx_svr, msg.from.address, "", subj, body.ToString, attach, err)
                    If Not err Like "" Then GoTo get_out

                Catch ex As Exception
                    err = "ER: Error updating allowed values in cr_forms, details: " & ex.ToString
                    GoTo get_out
                End Try
                '###########################################################################################################################
                'End of update forms processing
                '###########################################################################################################################

            ElseIf Regex.IsMatch(cleansubject, "^((anyquery)|(query))", RegexOptions.IgnoreCase) Then
                Dim date_conv As Boolean = True        'False => i'm thinking to convert the date for any query at all
                Dim sql() As String = {}
                Dim sql_err As String = ""
                Dim freespace As String = ""
                Dim cr_id() As String = {}
                Dim dir_out1 As String = local.base_path & local.db_outgoing
                clean_dir(dir_out1, err)
                Dim dir_out2 As String = local.base_path & local.db_outgoing & "\sql command output"
                Dim anyquery_flag As Boolean = If(Regex.IsMatch(cleansubject, "^anyquery", RegexOptions.IgnoreCase), True, False)

                'check it is from a qualified person
                '-----------------------------------------
                clear_dt(dt)
                If anyquery_flag Then : sqlquery(False, db, "SELECT email FROM " & db.schema & ".people WHERE anyquery='1' AND email='" & msg.from.address & "';", dt, err)
                Else : sqlquery(False, db, "SELECT email FROM " & db.schema & ".people WHERE query='1' AND email='" & msg.from.address & "';", dt, err)
                End If
                If Not err Like "" Then GoTo get_out
                If dt.Rows.Count = 0 Then
                    err = "QUERYREJ: Your email address doesn't have rights for this, sorry."
                    GoTo get_out
                End If

                'First check for key text in the body
                '-------------------------------------------------------------------------------------------
                'made this change 15jun, need to test
                'made this change 15jun, need to test
                'made this change 15jun, need to test
                'made this change 15jun, need to test
                'made this change 15jun, need to test
                'made this change 15jun, need to test
                'made this change 15jun, need to test
                'made this change 15jun, need to test
                Dim body_text As String = Trim(Regex.Replace(cleanbody, "((\r\n)|(\n)|(\r)|(\s{2,10}))", " ", RegexOptions.Singleline))
                '                Dim body_text As String = Trim(Regex.Replace(cleanbody, "((\r\n)|(\n)|(\r))", " ", RegexOptions.Singleline))
                '              body_text = Regex.Replace(body_text, "\s{2,10}", " ")

                'checks for easy query text in the body, in which case we do not search for query text in the attachments
                '-----------------------------------------------------------------------------------------------------------
                If Regex.IsMatch(body_text, "(^show\sme\s)(.*)(\scrs)", RegexOptions.IgnoreCase) Then
                    'these query the cr_common table
                    '-------------------------------
                    If Regex.IsMatch(body_text, "^show\sme\sall\scrs", RegexOptions.IgnoreCase) Then : sql = {"select * from " & db.schema & ".cr_common where cr_common.cr_status not like ''"}
                    ElseIf Regex.IsMatch(body_text, "^show\sme\sopen\scrs", RegexOptions.IgnoreCase) Then : sql = {"select * from " & db.schema & ".cr_common where Not cr_common.cr_status like 'Closed' and Not cr_common.cr_status like 'Cancelled'"}
                    ElseIf Regex.IsMatch(body_text, "^show\sme\sclosed\scrs", RegexOptions.IgnoreCase) Then : sql = {"select * from " & db.schema & ".cr_common where cr_common.cr_status like 'Closed'"}
                    ElseIf Regex.IsMatch(body_text, "^show\sme\scancelled\scrs", RegexOptions.IgnoreCase) Then : sql = {"select * from " & db.schema & ".cr_common where cr_common.cr_status like 'Cancelled'"}
                        '                ElseIf Regex.IsMatch(body_text, "^show\sme\sall\scrs\sthis\smonth\sfrom\s", RegexOptions.IgnoreCase) Then : sql = {"SELECT * from " & db.schema & ".cr_common WHERE cr_common.requester like '%" & Regex.Replace(body_text, "^show\sme\sall\scrs\sthis\smonth\sfrom\s", "", RegexOptions.IgnoreCase) & "%' AND open_date"}
                    End If

                    'adds conditions
                    '------------------
                    If Regex.IsMatch(body_text, "(\sfrom\s'[^']*')", RegexOptions.IgnoreCase) Then
                        sql = {sql(0) & " and cr_common.requester like '%" & Regex.Replace(body_text, "(^(.*)\sfrom\s')|('(.*)$)", "", RegexOptions.IgnoreCase) & "%'"}
                    End If

                    'adds in the details data
                    If Regex.IsMatch(body_text, "\sfor\s'[^']*'", RegexOptions.IgnoreCase) Then
                        'this takes off the start of the query before the 'where'
                        sql(0) = Regex.Replace(sql(0), "^(.*)\.cr_common\swhere\scr_common\.cr_status\s", "where cr_common.cr_status ", RegexOptions.IgnoreCase)

                        'this creates the new start of the query before the 'where'
                        sql(0) = "SELECT cr_common.*,B.* FROM " & db.schema & ".cr_common " & _
                                 "INNER JOIN " & _
                                    "(SELECT cr_sub_id,cr_type,node_type,node,nbr_node,parameter,proposed_setting,rollback_setting,NULL AS cur_az,NULL AS cur_mdt,NULL AS cur_edt,NULL as pro_az,NULL as pro_mdt,NULL as pro_edt,requester_comments,execution_coordinator,planned_execution_date,executor,NULL AS act_az,NULL AS act_mdt,NULL AS act_edt,NULL AS fin_az,NULL AS fin_mdt,NULL AS fin_edt,NULL AS fin_ht,NULL AS fin_antenna,NULL AS fin_coax_len,execution_status,execution_date,executor_comments " & _
                                    "FROM " & db.schema & ".cr_data_prm " & _
                                    "UNION " & _
                                    "SELECT cr_sub_id,cr_type,node_type,node,NULL AS nbr_node,NULL AS parameter,NULL AS proposed_setting,NULL AS rollback_setting,NULL AS cur_az,NULL AS cur_mdt,NULL AS cur_edt,NULL as pro_az,NULL as pro_mdt,NULL as pro_edt,requester_comments,execution_coordinator,planned_execution_date,executor,NULL AS act_az,NULL AS act_mdt,NULL AS act_edt,NULL AS fin_az,NULL AS fin_mdt,NULL AS fin_edt,NULL AS fin_ht,NULL AS fin_antenna,NULL AS fin_coax_len,execution_status,execution_date,executor_comments " & _
                                    "FROM " & db.schema & ".cr_data_oth " & _
                                    "UNION " & _
                                    "SELECT cr_sub_id,cr_type,node_type,node,NULL AS nbr_node,NULL AS parameter,NULL AS proposed_setting,NULL AS rollback_setting,cur_az,cur_mdt,cur_edt,pro_az,pro_mdt,pro_edt,requester_comments,execution_coordinator,planned_execution_date,executor,act_az,act_mdt,act_edt,fin_az,fin_mdt,fin_edt,fin_ht,fin_antenna,fin_coax_len,execution_status,execution_date,executor_comments " & _
                                    "FROM " & db.schema & ".cr_data_rfb) as B " & _
                                    "On B.cr_sub_id REGEXP concat('^',cr_common.cr_id,'\.[0-9]+$') " & _
                                 sql(0) & If(Regex.IsMatch(body_text, "\sfor\s'all\snodes'", RegexOptions.IgnoreCase), "", " AND B.node LIKE '" & Regex.Replace(body_text, "(^(.*)\sfor\s')|('(.*)$)", "", RegexOptions.IgnoreCase) & "%'")
                    End If
                    '                    If Regex.IsMatch(body_text, "(\sbetween\s)", RegexOptions.IgnoreCase) Then sql = {sql(0) & " and open_date like '%" & Regex.Replace(body_text, "((^(.*)\scrs\sfrom\s)|((\sbetween\s)(.*)))", "", RegexOptions.IgnoreCase) & "%'"}
                    '                    If Regex.IsMatch(body_text, "(\sthis week\s)", RegexOptions.IgnoreCase) Then sql = {sql(0) & " and open_date like '%" & Regex.Replace(body_text, "((^(.*)\scrs\sfrom\s)|((\sbetween\s)(.*)))", "", RegexOptions.IgnoreCase) & "%'"}
                    '                    If Regex.IsMatch(body_text, "(\sthis month\s)", RegexOptions.IgnoreCase) Then sql = {sql(0) & " and open_date like '%" & Regex.Replace(body_text, "((^(.*)\scrs\sfrom\s)|((\sbetween\s)(.*)))", "", RegexOptions.IgnoreCase) & "%'"}

                    'completes the command string
                    '---------------------------------
                    For i = 0 To sql.Length - 1
                        If sql(i).Length > 0 Then sql(i) = sql(i) & ";"
                    Next

                ElseIf Regex.IsMatch(body_text, "^show\sme\speople", RegexOptions.IgnoreCase) Then : sql = {"select * from " & db.schema & ".people;"}      'other table queries
                ElseIf Regex.IsMatch(body_text, "(^show\sme\scr\sfiles\s)(.*)", RegexOptions.IgnoreCase) Then : sql = {"cr files"}   'these are for cases of non sql queries
                ElseIf Regex.IsMatch(body_text, "^show\sme\shdd\sspace", RegexOptions.IgnoreCase) Then : sql = {"HDD space"}   'these are for cases of non sql queries
                End If

                'deals with the non sql queries
                '---------------------------------------
                If sql.Length > 0 AndAlso sql(0) Like "HDD space" Then
                    'the HDD space query
                    '---------------------------------------
                    sql = {"nosql: " & sql(0)}
                    For Each drive As DriveInfo In My.Computer.FileSystem.Drives
                        Try
                            freespace = freespace & " Freespace on drive: " & drive.Name & " = " & Regex.Replace(Str(Math.Round(drive.AvailableFreeSpace / (1073741824), 1)), ",", ".") & "GB,"
                        Catch ex As Exception
                        End Try
                    Next
                    If freespace.Length > 0 Then freespace = Left(freespace, Len(freespace) - 1)

                ElseIf sql.Length > 0 AndAlso sql(0) Like "cr files" Then
                    'the cr files query
                    '---------------------------------------
                    sql = {"nosql: " & sql(0)}
                    cr_id = Strings.Split(Trim(Regex.Replace(body_text, "(^show\sme\scr\sfiles)(\s)*", "", RegexOptions.IgnoreCase)), ",")
                    cr_id = (From item In cr_id
                             Let a = Trim(item)
                             Where Not a Like "" And FileIO.FileSystem.FileExists(local.base_path & local.cr & "\" & a & "\" & a & ".xlsb")
                             Select a).Distinct.ToArray
                    If cr_id.Count = 0 Then
                        err = "XREJ: Could not find any valid cr_ids in your cr_id list, either they do nto exist or they have been removed from the server HDD, you gave '" & Trim(Regex.Replace(body_text, "(^show\sme\scr\sfiles)(\s)*", "", RegexOptions.IgnoreCase)) & "'"
                        GoTo get_out
                    End If
                    For Each item In cr_id
                        zip_dir(local.base_path & local.cr & "\" & item, 10, 20, dir_out1 & "\" & item & ".zip", err)
                        If Not err Like "" Then
                            err = ""
                        End If
                    Next
                End If
                If sql.Length > 0 Then GoTo no_sql_attach

                'if not, then do the attachment check
                '---------------------------------------
                Try
                    'basic attachement check
                    '----------------------------
                    If msg.attachments.Count = 0 Then
                        GoTo no_sql_attach
                    End If
                    check_attachments(msg, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then
                        GoTo get_out
                    ElseIf Not err Like "" Then
                        err = "XREJ:  " & err
                        GoTo get_out
                    End If


                    'get the sql array from the txt file, only looks in the first .txt file found
                    '----------------------------------------------
                    Dim file As String = ""
                    If FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.txt").Count > 0 Then
                        file = FileIO.FileSystem.GetFiles(local.base_path & local.inbox, FileIO.SearchOption.SearchTopLevelOnly, "*.txt").First
                    End If
                    sql = get_data_from_txt(file, format, err)
                    If sql.Count = 0 Then
                        err = "XREJ: Email error, no sql commands were found in your file, giving up...."
                        GoTo get_out
                    End If
                Catch ex As Exception
                    err = "ER: Error getting attachments, details: " & ex.ToString
                    GoTo get_out
                End Try

no_sql_attach:

                'runs the querys
                '------------------
                Try
                    If Regex.IsMatch(sql(0), "nosql:\s", RegexOptions.IgnoreCase) Then GoTo skip_sql

                    FileIO.FileSystem.CreateDirectory(dir_out2)
                    Dim i As Integer = 1
                    Dim sql2run() = (From item In sql.AsEnumerable Where Not Regex.IsMatch(sql.ToString, "^nosql:\s", RegexOptions.IgnoreCase) And Not sql.ToString Like "" Select item).Distinct.ToArray
                    For Each item In sql2run
                        item = Trim(item)
                        If item Like "" Then GoTo skip_anyquery

                        If Not Regex.IsMatch(item, "^select\s.*;$", RegexOptions.IgnoreCase) And anyquery_flag Then
                            'THE NON SELECT CASE
                            anyquery_db(True, db, item.ToString, err)
                            If Not err Like "" Then
                                sql_err = If(sql_err Like "", "", sql_err & ", ") & "NEW ERROR: " & item.ToString & " => msg: " & err
                                err = ""
                                GoTo skip_anyquery
                            Else
                                sql_err = If(sql_err Like "", "", sql_err & ", ") & "NEW ERROR: SQL OK: " & item.ToString
                            End If

                        ElseIf Regex.IsMatch(item, "^select\s.*;$", RegexOptions.IgnoreCase) Then
                            'THE SELECT CASE
                            clear_dt(dt)
                            sqlquery(True, db, item.ToString, dt, err)
                            If Not err Like "" Then
                                sql_err = If(sql_err Like "", "", sql_err & ", ") & "NEW ERROR: " & item.ToString & " => msg: " & err
                                err = ""
                                GoTo skip_anyquery
                            Else
                                sql_err = If(sql_err Like "", "", sql_err & ", ") & "NEW ERROR: SQL OK: " & item.ToString
                            End If

                            'find any cols that are datetime => if the DB col format is sql datetime, the corresponding col format in the dt will also be datetime, but vb datetime, so we can use this as a filter
                            'any cols I find will be converted when I write to csv.
                            '-------------------------------------------------------------------------------------------------------------------------
                            If dt.Rows.Count > 0 Then
                                Dim date_convert_cols() As Integer = {}
                                If date_conv Then date_convert_cols = (From col As System.Data.DataColumn In dt.Columns Let a = col.Ordinal Where col.DataType = GetType(DateTime) Select a).ToArray

                                'writes the output to csv
                                '-------------------------
                                Dim file_out As String = "query_output_" & i
                                dt2csv(date_convert_cols, dir_out2, file_out, dt, err)
                                If Not err Like "" Then
                                    err = ""
                                    GoTo skip_anyquery
                                End If
                                i += 1
                            End If
                        End If
skip_anyquery:
                    Next

                    'this zips the output for sending
                    '-------------------------------------
                    If FileIO.FileSystem.GetFiles(dir_out2, FileIO.SearchOption.SearchTopLevelOnly, "*.csv").Count > 0 Then
                        zip_dir(dir_out2, 10, 20, dir_out1 & "\" & "sql command output.zip", err)
                        If Not err Like "" Then
                            err = ""
                        End If
                    End If
                Catch ex As Exception
                    err = "ER: Internal error running sql, deatils: " & ex.ToString
                    GoTo get_out
                End Try

skip_sql:
                Try
                    'ok, it is good, update the log
                    '----------------------------------
                    time_now = Now
                    add2log(db, time_now, "", "Query run by " & msg.from.address, err)
                    If Not err Like "" Then GoTo get_out

                    'Send the results
                    '----------------
                    Dim zip_attach() As String = FileIO.FileSystem.GetFiles(dir_out1, FileIO.SearchOption.SearchTopLevelOnly, "*.zip").ToArray
                    Dim zip_attach_name() = (From item In zip_attach Let a = Path.GetFileName(item) Select a).ToArray
                    Dim subj As String = "CRMS: Query Results"
                    Dim body As New StringBuilder("")
                    If zip_attach.Count > 0 Then
                        body.Append("<BR>The results of your query are attached:")
                        body.Append(array2html_bullet_list(zip_attach_name))
                    Else
                        body.Append("<BR>There were no results to attach.")
                    End If

                    If Not Regex.IsMatch(sql(0), "^nosql:\s", RegexOptions.IgnoreCase) Then
                        body.Append("<BR>SQL Run Results:")
                        sql_err = Regex.Replace(sql_err, "^NEW ERROR:\s", "")
                        Dim query() = (From item In Strings.Split(sql_err, ", NEW ERROR: ") Where Not item Like "" Select item).ToArray
                        body.Append(array2html_bullet_list(query))
                    End If

                    If Regex.IsMatch(sql(0), "^nosql:\sHDD\sspace", RegexOptions.IgnoreCase) Then
                        body.Append("<P>HDD Freespace Details:")
                        body.Append(array2html_bullet_list(Strings.Split(freespace, ",")))
                    End If

                    body.Append("<BR>Thanks<BR>")
                    body.Append("<BR>-----------------------------------------------------------------------------------------------------------<BR>" & If(Not msg.body_html_raw Like "", "<div style=""text-align:justify;padding-left:10px;padding-right:5px;"">" & msg.body_html_raw & "</div>", msg.body_text_raw))
                    auto_reply_general(tx_svr, msg.from.address, "", subj, body.ToString, zip_attach, err)
                    If Not err Like "" Then GoTo get_out

                Catch ex As Exception
                    If err = "" Then
                        err = "ER: Error processing sql query, details: " & ex.ToString
                    End If
                    GoTo get_out
                End Try
                '###########################################################################################################################
                'End of sql command processing
                '###########################################################################################################################

            ElseIf Regex.IsMatch(cleansubject, "^admin\s[A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9]$", RegexOptions.IgnoreCase) Then
                'Code to deal with admin commands in the body

            ElseIf Regex.IsMatch(cleansubject, "^report\s[A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9]$", RegexOptions.IgnoreCase) Then
                'Code to deal with report requests in the body

            Else
                err = "XREJ: Sorry, I do not understand subject text."
                GoTo get_out
            End If
get_out:
        Catch ex As Exception
            err = "ER: Error in the analyse email sub: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub



    '#############################################################################
    Friend Sub clear_dt(ByRef dt As System.Data.DataTable)
        On Error Resume Next
        dt.Columns.Clear()
        dt.Rows.Clear()
        dt.Dispose()
    End Sub
    '#############################################################################



    'sends a dt to a csv file
    'note: I convert any cols in the date_convert_cols array to the OAdate for viewing in XL.
    '---------------------------
    Friend Sub dt2csv(ByVal date_convert_cols() As Integer, ByVal path As String, ByVal name As String, ByVal dt As System.Data.DataTable, ByRef err As String)
        Try
            'This does the export row by row to csv
            '-----------------------------------------
            Dim x, y As Integer
            x = dt.Rows.Count
            y = dt.Columns.Count
            Dim colcounter As Integer = 0
            Dim rowcounter As Integer = 0
            Dim writer As New StreamWriter(path & "\" & name & "_body.csv", False)
            For rowcounter = 0 To x - 1
                Dim sb As New StringBuilder()
                For colcounter = 0 To y - 1
                    If dt.Rows(rowcounter)(colcounter) Is DBNull.Value Then
                        sb.Append("""" & "" & """" & ",")
                    ElseIf date_convert_cols.Contains(colcounter) Then
                        Try
                            sb.Append("""" & dt.Rows(rowcounter).Field(Of DateTime)(colcounter).ToOADate & """" & ",")
                        Catch ex As Exception
                            sb.Append("""" & "" & """" & ",")
                        End Try
                    Else
                        sb.Append("""" & dt.Rows(rowcounter)(colcounter).ToString & """" & ",")
                    End If
                Next
                sb.Replace(",", "", sb.Length - 1, 1)
                writer.WriteLine(sb.ToString)
            Next
            writer.Close()
            writer = Nothing
        Catch ex As Exception
            err = "ER: Some error writing body to csv... " & ex.ToString
            GoTo get_out
        End Try

        Try
            'this writes the final headers
            '---------------------------------------------------------------------
            Dim writer As New StreamWriter(path & "\" & name & ".csv", False)
            Dim sb As New StringBuilder()
            For Each column As DataColumn In dt.Columns
                sb.Append(column.ColumnName.ToString & ",")
            Next
            sb.Replace(",", "", sb.Length - 1, 1)
            writer.WriteLine(sb.ToString)

            'We append the body csv to the header csv
            '--------------------------------------------
            Dim reader As New StreamReader(path & "\" & name & "_body.csv")
            Do While Not reader.EndOfStream
                writer.WriteLine(reader.ReadLine())
            Loop
            writer.Close()
            reader.Close()
            Kill(path & "\" & name & "_body.csv")
        Catch ex As Exception
            err = "ER: There was an error adding the headers to csv... " & ex.ToString
            GoTo get_out
        End Try
get_out:
    End Sub



    Friend Sub check_attachments(ByVal msg As CRMS.email_msg, ByRef err As String)
        Try
            ' check we have no blank attachment spaces
            '-----------------------------------------
            For Each item In msg.Attachments
                If item Is Nothing Or item.FileName = "" Then
                    err = "Some of your attachments are getting stopped by the email server threat protection, do not send double zipped files, xl protected workbooks or any encrypted files.<BR>Thanks"
                    GoTo get_out
                End If
            Next
get_out:
        Catch ex As Exception
            err = "ER: There was an error doing basic attachment checks, details: " & ex.ToString
        End Try
    End Sub



    Friend Sub process_attachments_combine(ByVal type As String, ByVal cr_id As String, ByVal local As local_machine, ByRef err As String)
        'puts all the attachments in the executor/requester attachments dir of the inbox for checking
        '-----------------------------------------------------------------------------------------
        Try
            If FileIO.FileSystem.DirectoryExists(local.base_path & local.inbox & "\" & type & " attachments") Then
                clean_dir(local.base_path & local.inbox & "\" & type & " attachments", err)
                If Not err Like "" Then GoTo get_out
            Else
                FileIO.FileSystem.CreateDirectory(local.base_path & local.inbox & "\" & type & " attachments")
            End If
            If FileIO.FileSystem.DirectoryExists(local.base_path & local.inbox & "\" & type & " attachments temp") Then
                clean_dir(local.base_path & local.inbox & "\" & type & " attachments temp", err)
                If Not err Like "" Then GoTo get_out
            Else
                FileIO.FileSystem.CreateDirectory(local.base_path & local.inbox & "\" & type & " attachments temp")
            End If
        Catch ex As Exception
            err = "ER: Error creating dir in inbox when combining " & type & "s attachments, details: " & ex.ToString
        End Try

        Dim cr_dir As String = local.base_path & local.cr & "\" & cr_id
        Dim inbox_final_dir As String = local.base_path & local.inbox & "\" & type & " attachments"
        Dim inbox_temp_dir As String = local.base_path & local.inbox & "\" & type & " attachments temp"
        Try
            For Each item In IO.Directory.GetFiles(local.base_path & local.inbox, "*", IO.SearchOption.TopDirectoryOnly)
                If Path.GetFileName(item) Like cr_id & ".xlsb" Or Path.GetFileName(item) Like "~$*" Then
                    'I do not care if the executor/requester has re-sent the cr form, I already have this data and it's locked down in the DB, so this one is not needed, I'll use my stored version
                    force_delete_file(item, err)
                    If Not err Like "" Then GoTo get_out
                Else
                    FileIO.FileSystem.MoveFile(item, inbox_temp_dir & "\" & Path.GetFileName(item), True)
                End If
            Next
            For Each item In IO.Directory.GetDirectories(local.base_path & local.inbox, "*", IO.SearchOption.TopDirectoryOnly)
                Dim dirInfo As New System.IO.DirectoryInfo(item)
                Dim dir As String = dirInfo.Name
                If Not dir = type & " attachments" And Not dir = type & " attachments temp" Then
                    FileIO.FileSystem.MoveDirectory(item, inbox_temp_dir & "\" & dir, True)
                End If
            Next

            'move old cr dir attachments to the final attchment dir
            '--------------------------------------------------------
            For Each item In FileIO.FileSystem.GetFiles(cr_dir, FileIO.SearchOption.SearchTopLevelOnly, type & " attachments*.zip")
                unzip_file(item, inbox_final_dir, err)
                If Not err Like "" Then GoTo get_out
                force_delete_file(item, err)
                If Not err Like "" Then GoTo get_out
            Next

            'move all new attachments to the executor attachments dir, overwriting on files/dirs with the same names as the original attachments if any exist
            '--------------------------------------------------------
            For Each item In IO.Directory.GetFiles(inbox_temp_dir, "*", IO.SearchOption.TopDirectoryOnly)
                FileIO.FileSystem.MoveFile(item, inbox_final_dir & "\" & Path.GetFileName(item), True)
            Next
            For Each item In IO.Directory.GetDirectories(inbox_temp_dir, "*", IO.SearchOption.TopDirectoryOnly)
                Dim dirInfo As New System.IO.DirectoryInfo(item)
                Dim dir As String = dirInfo.Name
                FileIO.FileSystem.MoveDirectory(item, inbox_final_dir & "\" & dir, True)
            Next
            IO.Directory.Delete(inbox_temp_dir, True)

            'zip all files and put back in the cr directory, if there are any files to zip
            '------------------------------------------------------------------
            If (IO.Directory.GetFiles(inbox_final_dir, "*", IO.SearchOption.TopDirectoryOnly).Count + IO.Directory.GetDirectories(inbox_final_dir, "*", IO.SearchOption.TopDirectoryOnly).Count) = 0 Then
                IO.Directory.Delete(inbox_final_dir, True)
            Else
                zip_dir(inbox_final_dir, 10, 20, cr_dir & "\" & type & " attachments.zip", err)
                If Not err = "" Then GoTo get_out
            End If
        Catch ex As Exception
            err = "ER: Error finding attachments in " & type & "s email."
        End Try
get_out:
    End Sub





    Friend Sub process_attachments(ByVal type As String, ByVal cr_id As String, ByVal local As local_machine, ByRef err As String)
        Try
            If FileIO.FileSystem.DirectoryExists(local.base_path & local.inbox & "\" & type & " attachments") Then
                clean_dir(local.base_path & local.inbox & "\" & type & " attachments", err)
                If Not err Like "" Then GoTo get_out
            Else
                FileIO.FileSystem.CreateDirectory(local.base_path & local.inbox & "\" & type & " attachments")
            End If
        Catch ex As Exception
            err = "ER: couldn't create dir for attachments in inbox, details: " & ex.ToString
        End Try

        Dim cr_dir As String = local.base_path & local.cr & "\" & cr_id
        Dim inbox_final_dir As String = local.base_path & local.inbox & "\" & type & " attachments"
        Try
            For Each item In IO.Directory.GetFiles(local.base_path & local.inbox, "*", IO.SearchOption.TopDirectoryOnly)
                If Path.GetFileName(item) Like "~$*" Then
                    'delete any open file shadows
                    force_delete_file(item, err)
                    If Not err Like "" Then GoTo get_out
                ElseIf Not Path.GetFileName(item) Like "~$*" And Not Path.GetFileName(item) Like cr_id & ".xlsb" Then
                    FileIO.FileSystem.MoveFile(item, inbox_final_dir & "\" & Path.GetFileName(item), True)
                End If
            Next
            For Each item In IO.Directory.GetDirectories(local.base_path & local.inbox, "*", IO.SearchOption.TopDirectoryOnly)
                Dim dirInfo As New System.IO.DirectoryInfo(item)
                Dim dir As String = dirInfo.Name
                If Not dir = type & " attachments" Then
                    FileIO.FileSystem.MoveDirectory(item, inbox_final_dir & "\" & dir, True)
                End If
            Next

            'if there are no files/dirs in the requester attachments dir, then delete it, otherwise multi part zip it to the cr dir
            '----------------------------------------------------------------------------------------------------------------------
            If IO.Directory.GetFiles(inbox_final_dir, "*", IO.SearchOption.TopDirectoryOnly).Count + IO.Directory.GetDirectories(inbox_final_dir, "*", IO.SearchOption.TopDirectoryOnly).Count = 0 Then
                IO.Directory.Delete(inbox_final_dir, True)
            Else
                'zips the attachments
                '-------------------------
                zip_dir(inbox_final_dir, 10, 20, cr_dir & "\" & type & " attachments.zip", err)
                If Not err = "" Then GoTo get_out
            End If
        Catch ex As Exception
            err = "ER: Error zipping attachments."
        End Try
get_out:
    End Sub




    Friend Function get_user_name(ByVal col As String, ByVal cr_id As String, ByVal format As cr_sheet_format, ByVal db As mysql_server) As String()
        Try
            Dim err As String = ""
            Dim name() As String = {}
            Dim dt As New System.Data.DataTable
            sqlquery(False, db, "SELECT " & col & " as user FROM " & db.schema & ".cr_common WHERE cr_id LIKE '" & cr_id & "';", dt, err)
            If Not err Like "" Then
                GoTo get_out
            ElseIf dt.Rows.Count = 0 Then
                err = "ER: DB error, can't find cr_id in DB."
                GoTo get_out
            End If
            Dim t_a() As String = Strings.Split(dt.Rows(0).Field(Of String)("user"), ",")
            If dt.Rows(0).Field(Of String)("user") Like "" Then t_a = {}
            name = (From item In t_a
                    Let a = c2e(item), b = Regex.Replace(item, "\s\(.*", "")
                    Where format.IsValidEmail(a)
                    Select b).ToArray
get_out:
            Return name
        Catch ex As Exception
            Return {}
        End Try
    End Function








    Public Sub get_cr_id_data(ByVal cr_id As String, ByRef cr_status As String, ByRef cr_type As String, ByRef cr_type_short As String, ByRef cr_form_type As String, ByRef cc_list As String, ByRef requester As String, ByRef approver As String, ByRef excoord As String, ByRef executors As String, ByVal db As mysql_server, ByRef err As String)
        Try
            cr_status = ""
            cr_type = ""
            cr_type_short = ""
            cr_form_type = ""
            cc_list = ""
            requester = ""
            approver = ""
            excoord = ""
            executors = ""
            Dim dt As New System.Data.DataTable
            Dim sqltext As String = "SELECT * FROM " & db.schema & ".cr_common WHERE cr_id = '" & cr_id & "';"
            sqlquery(False, db, sqltext, dt, err)
            If dt.Rows.Count > 0 Then
                cr_status = dt.Rows(0).Field(Of String)("cr_status")
                cr_type = dt.Rows(0).Field(Of String)("cr_type")
                cr_type_short = dt.Rows(0).Field(Of String)("cr_type_short")
                cr_form_type = dt.Rows(0).Field(Of String)("cr_form_type")
                cc_list = dt.Rows(0).Field(Of String)("cc_list")
                requester = dt.Rows(0).Field(Of String)("requester")
                approver = dt.Rows(0).Field(Of String)("approver")
                excoord = dt.Rows(0).Field(Of String)("execution_coordinator")
                executors = dt.Rows(0).Field(Of String)("executors")
            End If
            dt.Dispose()
        Catch ex As Exception
            err = "ER: sql error getting cr_id from the cr_common table, details: " & ex.ToString
        End Try
    End Sub





    Friend Sub get_user_emails(ByVal cc_mask() As Boolean, ByRef user_list As String, ByRef requester As String, ByRef approver As String, ByRef execution_coordinator As String, ByRef executors As String, ByRef cc_list As String, ByVal cr_id As String, ByVal format As cr_sheet_format, ByVal db As mysql_server, ByRef err As String)
        'NOTE: BECAUSE WE ARE STORING THE COMBINED NAME IN THE CR FORM AND THE CR COMMON AND CR DATA TABLES, WE NEED TO CONVERT IT TO JUST THE EMAIL PORTION 
        'WHEN WE READ OUT FROM THESE FIELDS IN THE FORM/DB TABLE
        Try
            user_list = ""
            requester = ""
            approver = ""
            execution_coordinator = ""
            executors = ""
            Dim dt As New System.Data.DataTable
            sqlquery(False, db, "SELECT requester,approver,execution_coordinator,executors,cc_list FROM " & db.schema & ".cr_common WHERE cr_id LIKE '" & cr_id & "';", dt, err)
            If Not err Like "" Then
                GoTo get_out
            ElseIf dt.Rows.Count = 0 Then
                err = "ER: DB error, can't find cr_id in DB."
                GoTo get_out
            End If
            If format.IsValidEmail(c2e(dt.Rows(0).Field(Of String)("requester"))) Then
                requester = c2e(dt.Rows(0).Field(Of String)("requester"))
                user_list = user_list & requester & ","
            End If
            If format.IsValidEmail(c2e(dt.Rows(0).Field(Of String)("approver"))) Then
                approver = c2e(dt.Rows(0).Field(Of String)("approver"))
                user_list = user_list & approver & ","
            End If
            If format.IsValidEmail(c2e(dt.Rows(0).Field(Of String)("execution_coordinator"))) Then
                execution_coordinator = c2e(dt.Rows(0).Field(Of String)("execution_coordinator"))
                user_list = user_list & execution_coordinator & ","
            End If
            Dim t_a() As String = Strings.Split(dt.Rows(0).Field(Of String)("executors"), ",")
            If dt.Rows(0).Field(Of String)("executors") Like "" Then t_a = {}
            For Each item In t_a
                If format.IsValidEmail(c2e(item)) Then
                    executors = executors & c2e(item) & ","
                    user_list = user_list & c2e(item) & ","
                End If
            Next
            If Not user_list Like "" Then user_list = check_email_list(True, Left(user_list, Len(user_list) - 1), format)
            If Not executors Like "" Then executors = check_email_list(True, Left(executors, Len(executors) - 1), format)

            'does the cc_list which is in email only format => we check the email addresses, then we take out any users
            '---------------------------------------------------------------------------------------------
            Dim cc_out_s As String = ""
            Dim cc_out() As String = (From item In Strings.Split(dt.Rows(0).Field(Of String)("cc_list"), ",").AsEnumerable
                                        Where format.IsValidEmail(item)
                                        Select item).ToArray
            If cc_out.Length > 0 Then
                cc_out = cc_out.Except(Strings.Split(user_list, ",")).ToArray
                cc_out_s = If(cc_out.Length > 0, Join(cc_out, ","), "")
            Else
                cc_out_s = ""
            End If

            'now we make the final cc_list from the given mask, basically we add whoever we have added in the mask to the final cc_list
            '--------------------------------------------------------------------------------
            Dim cc_final() As String = {}
            Dim t_string As String = ""
            For i = 0 To 4
                If i = 0 Then : t_string = requester
                ElseIf i = 1 Then : t_string = approver
                ElseIf i = 2 Then : t_string = execution_coordinator
                ElseIf i = 3 Then : t_string = executors
                ElseIf i = 4 Then : t_string = cc_out_s
                End If
                If Not t_string Like "" And cc_mask(i) Then
                    cc_final = cc_final.Union(Strings.Split(t_string, ",")).ToArray
                End If
            Next
            cc_list = If(cc_final.Length > 0, Join(cc_final, ","), "")
            If Not cc_list Like "" Then cc_list = check_email_list(False, cc_list, format)
get_out:
        Catch ex As Exception
            err = "ER: error getting users from cr_common, details: " & ex.ToString
        End Try
    End Sub







    Friend Sub get_cc_final_lists(ByVal cc_list As String, ByVal to_list As String, ByRef cc_list_final As String, ByVal format As cr_sheet_format, ByRef err As String)
        'NOTE: THIS IS WORKING IN EMAIL ONLY LISTS, NOT COMBINED NAME LISTS!!  
        'This just takes out the to_list from the cc_list
        'also apply the cc_mask here
        Try
            cc_list_final = ""
            Dim cc_array() As String = Strings.Split(cc_list, ",").Distinct.ToArray
            If cc_list Like "" Then cc_array = {}
            Dim to_array() As String = Strings.Split(to_list, ",").Distinct.ToArray
            If to_list Like "" Then to_array = {}
            If Not cc_array Is Nothing AndAlso cc_array.Length > 0 Then
                cc_array = cc_array.Except(to_array).ToArray
                If Not cc_array Is Nothing AndAlso cc_array.Length > 0 Then cc_list_final = check_email_list(True, Join(cc_array, ","), format)
            End If
        Catch ex As Exception
            err = "ER: error getting final cc lists, details: " & ex.ToString
        End Try
    End Sub





    Friend Sub pre_process_for_resub(ByRef cr_resub_form As String, ByVal cr_id As String, ByVal cr_form As String, ByVal local As local_machine, ByRef err As String)
        Try
            'create the resub form in the cr dir and delete any attachments
            '-----------------------------------------------------------
            If Not FileIO.FileSystem.DirectoryExists(Path.GetDirectoryName(cr_form)) AndAlso Not FileIO.FileSystem.FileExists(cr_form) Then
                err = "ER: the cr_form was not found on the server HDD, internal error, can't continue...."
                GoTo get_out
            End If
            cr_resub_form = Path.GetDirectoryName(cr_form) & "\" & "Resubmit_Form_" & cr_id & ".xlsb"
            If FileIO.FileSystem.FileExists(cr_resub_form) Then
                force_delete_file(cr_resub_form, err)
                If Not err Like "" Then GoTo get_out
            End If
            FileIO.FileSystem.CopyFile(cr_form, cr_resub_form, True)

            'delete the requester attachments if there are any
            '---------------------------------------------
            For Each item In FileIO.FileSystem.GetFiles(local.base_path & local.cr & "\" & cr_id, FileIO.SearchOption.SearchTopLevelOnly, "requester attachments*.zip")
                force_delete_file(item, err)
                If Not err Like "" Then GoTo get_out
            Next
get_out:
        Catch ex As Exception
            err = "ER: error creating resub form, details: " & ex.ToString
        End Try
    End Sub





    Friend Sub combine_cc_list(ByRef cc_list As String, ByVal cc_list_old As String, ByVal format As cr_sheet_format, ByRef err As String)
        'NOTE: THIS IS WORKING IN EMAIL ONLY LISTS, NOT COMBINED NAME LISTS!!
        Try
            Dim a() As String = Strings.Split(cc_list, ",")
            If cc_list Like "" Then a = {}
            Dim a_old() As String = Strings.Split(cc_list_old, ",")
            If cc_list_old Like "" Then a_old = {}
            a = a.Union(a_old).ToArray
            a = a.Distinct.ToArray
            cc_list = check_email_list(True, Join(a, ","), format)
        Catch ex As Exception
            err = "ER: error combining the cc_list, details: " & ex.ToString
        End Try
    End Sub






    'this will remove bad addresses, for the clean option, it will remove repeats and just return a list of email addresses without other text
    'for the not clean option it only removes repeats and nonvalid addresses
    Friend Function check_email_list(ByVal clean_flag As Boolean, ByVal list As String, ByVal format As cr_sheet_format) As String
        Try
            If Not list Is Nothing And Not list = "" Then
                If clean_flag Then
                    Dim query = From item In Strings.Split(list, ",").AsEnumerable
                                Let email = c2e(Trim(item))
                                Where format.IsValidEmail(email)
                                Select email
                    If query.Count > 0 Then
                        Return Join(query.Distinct.ToArray, ",")
                    Else
                        Return ""
                    End If
                Else
                    Dim query = From item In Strings.Split(list, ",").AsEnumerable
                                Let a = Trim(item)
                                Where format.IsValidEmail(c2e(a))
                                Select a
                    If query.Count > 0 Then
                        Return Join(query.Distinct.ToArray, ",")
                    Else
                        Return ""
                    End If
                End If
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function





    Friend Function array2html_bullet_list(ByVal a() As String) As String
        On Error Resume Next
        Dim s As String = ""
        If Not a Is Nothing AndAlso a.Length > 0 Then
            s = "<UL>"
        End If
        For Each item In a
            s = s & "<LI>" & item & "</LI>"
        Next
        If Not a Is Nothing AndAlso a.Length > 0 Then
            s = s & "</UL>"
        End If
        Return s
    End Function


    '#####################################################################
    '#####################################################################
    '#####################################################################
    '#####################################################################
    'not tested since I changed the email format stored in the tool to name (email)
    '#####################################################################
    '#####################################################################
    '#####################################################################
    '#####################################################################
    Friend Function get_executors_from_txt(ByVal dir As String, ByVal format As cr_sheet_format, ByRef err As String) As String
        Try
            Dim s As String = ""
            Dim attach() As String = FileIO.FileSystem.GetFiles(dir, FileIO.SearchOption.SearchTopLevelOnly, "*.csv").ToArray
            If Not attach Is Nothing Then
                If attach.Count = 0 Then
                    err = "ADDEXREJ: no acceptable .csv files found, you must attach a .csv file with the executors to add, 1 executor per line: Name, Email"
                    GoTo get_out
                Else
                    For Each file In attach
                        If Is_File_Open(file) Then GoTo file_skip
                        Dim input() As String = {}
                        Using reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(file)
                            reader.TextFieldType = FileIO.FieldType.Delimited
                            reader.SetDelimiters(",")
                            reader.HasFieldsEnclosedInQuotes = True   'this will treat any fields with enclosing quotes as 1 field regardless of what is in it
                            While Not reader.EndOfData
                                input = reader.ReadFields()
                                If Not input.Count = 2 Then
                                    GoTo line_skip
                                Else
                                    If format.IsValidEmail(input(1)) Then
                                        s = s & input(0) & "(" & input(1) & "),"
                                    End If
                                End If
line_skip:
                            End While
                        End Using
file_skip:
                    Next
                End If
            End If
            If Not s = "" Then
                s = Left(s, Len(s) - 1)
            End If
get_out:
            Return s
        Catch ex As Exception
            err = "ER: error getting users from csv, details " & ex.ToString
            Return ""
        End Try
    End Function



    'this reads in instructions from a .txt file
    'NOTE: they must be terminated with a semicolon for this to work though, otherwise, they will be ignored.
    'Aug15 => added in ability to handle multiline sql statements => keeps concatenating until it gets an ';'
    '----------------------------------------------------------------------------------------------
    Friend Function get_data_from_txt(ByVal file As String, ByVal format As cr_sheet_format, ByRef err As String) As String()
        Try
            Dim data() As String = {}
            Dim s_data As String = ""
            If Not file Is Nothing AndAlso Not file = "" AndAlso Not Is_File_Open(file) Then
                Dim input As String = ""
                Using reader As New StreamReader(file)
                    s_data = reader.ReadToEnd
                End Using
                s_data = Regex.Replace(s_data, "((\r\n)|(\n)|(\r))+", "")
                data = Regex.Split(s_data, ";")
                data = data.Except({"", vbCrLf}).Distinct.ToArray

                'this ensures the items selected have ; at the end, are not blanks and do not have any other weird chars at the end
                Dim data_out() As String = (From item In data Let a = Trim(item) & ";" Where Not item Like "" Select a).Distinct.ToArray
                data = data_out
            End If
            Return data
        Catch ex As Exception
            err = "ER: error getting data from txt, details " & ex.ToString
            Return {}
        End Try
    End Function



    'keep this for later
    Friend Function get_data_from_csv(ByVal file As String, ByVal format As cr_sheet_format, ByRef err As String) As String()
        Try
            Dim data() As String = {}
            Dim s_data As String = ""
            If Not file Is Nothing AndAlso Not file = "" AndAlso Not Is_File_Open(file) Then
                Dim input() As String = {}
                Using reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(file)
                    reader.TextFieldType = FileIO.FieldType.Delimited
                    reader.SetDelimiters(",")
                    reader.HasFieldsEnclosedInQuotes = True   'this will treat any fields with enclosing quotes as 1 field regardless of what is in it
                    While Not reader.EndOfData
                        input = reader.ReadFields
                        If Not input.Count = 0 Then
                            'you have to really process a csv line by line as the array length could vary, unless you specify that it doesn't
                            'you would typically read it into a dt or db or something.....  up to you
                        Else
                            GoTo line_skip
                        End If
line_skip:
                    End While
                End Using
            End If
            Return data
        Catch ex As Exception
            err = "ER: error getting data from csv, details " & ex.ToString
            Return {}
        End Try
    End Function


    'checks if the sum of the size of the input files is above a threshold and zips them.  It uses the CRMS\temp dir
    Public Sub check_size_and_zip(ByVal cr_form As String, ByRef cr_form_zipped As String, ByVal format As cr_sheet_format, ByVal local As local_machine, ByRef err As String)
        Dim total_MB As Double = 0
        Try
            cr_form_zipped = ""
            If Not FileIO.FileSystem.FileExists(cr_form) Then GoTo get_out
            Dim fileinfo As New FileInfo(cr_form)
            total_MB = fileinfo.Length / 1000000.0
            If total_MB > format.cr_form_size_zip_limit Then
                zip_file(cr_form, local.base_path & "\temp\cr_form.zip", False, err)
                If Not err Like "" Then GoTo get_out
                cr_form_zipped = local.base_path & "\temp\cr_form.zip"
            End If
get_out:
        Catch ex As Exception
            err = "ER: error checking attachment size and zipping them if required"
        End Try
    End Sub




End Module

