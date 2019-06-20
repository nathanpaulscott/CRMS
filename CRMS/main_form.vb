Option Explicit On
Option Compare Text

Public Class main_form
    'we create instances of all the critical components of the app
    '------------------------------------------------------------------
    Public local As New CRMS.local_machine
    Public crms_control As New CRMS.crms_controller
    Public format As New cr_sheet_format
    Public time_old As Integer = 0
    Public time_new As Integer = 0
    Public busy_flag As Boolean = False
    Public google_drive_login_data As New google_drive_creds

    '----------------------------------------------------------------------
    '----------------------------------------------------------------------
    'Main Form loader
    '------------------------
    Public Sub Test_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'On Error Resume Next
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", False)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US", False)
        Dim err As String = ""

        'initialise all vars on form load
        '-------------------------------------
        'create the allowed value datatables
        '--------------------------------------
        format.create_datatables(format.ds_allow)

        'check if there is a saved settings file, if so, open it and read in defaults, otherwise, use hardcoded defaults
        '------------------------------------------------------------------------------------
        If Not FileIO.FileSystem.DirectoryExists(local.base_path) Then
            MsgBox("The settings directory doesn't exist, please ensure it exists before running the app, if you have a settings file, ensure it is called (crms_settings.csv) and is in this directory: " & local.base_path)
            Forms.Application.Exit()
        ElseIf FileIO.FileSystem.FileExists(local.base_path & "\" & local.settings_file) Then
            read_from_settings()
        Else    'there is no settings file, so use hardcoded vals
            MsgBox("The settings file doesn't exist, falling back to hardcoded defaults.  Settings file: " & local.base_path & "\" & local.settings_file)
            set_hardcoded_vals()
        End If
        crms_control.check_mail_period = Val(Me.TextBox_check_mail_period.Text)
        crms_control.check_cr_status_times = Strings.Split(Me.TextBox_cr_status_times.Text, ",")
        format.use_imap = Me.CheckBox_use_imap.Checked

        'check the blank cr creation form exists
        '----------------------------------------
        If Not FileIO.FileSystem.FileExists(local.base_path & local.cr_blank_request_form_dir & "\" & local.cr_form) Then
            MsgBox("There is no blank CR form for the app to use, there must be a blank CR creation form named (" & local.base_path & local.cr_blank_request_form_dir & "\" & local.cr_form & ") on the HDD, impossible to continue, exiting...")
            Forms.Application.Exit()
        End If

        Me.Update()

        'do date and HW checks
        '---------------------
        'use this to see the nic data
        '        For Each nic As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()
        '        debug.writeline(Now.ToLongTimeString & ": " & "Name: " & nic.Name & ", MAC: " & nic.GetPhysicalAddress.ToString & ", other: " & nic.Id & ", other: " & nic.NetworkInterfaceType & ", other: " & nic.Description)
        '        Next
        run_checks()

        'we do DB vs sheet_format consistency check first to ensure no admin related issues exist that would lead to sql errors later
        '-------------------------------------------------------------------------------------------------------------------
        format_vs_db_check(format, err)
        If Not err Like "" Then
            MsgBox("There are inconsistencies between the DB tables and the tables defined in the app, exiting...details: " & err)
            Forms.Application.Exit()
        End If

        'recreate the blank template forms, just incase the blank cr form has changed
        '-----------------------------------------------------------------------
        Dim form_types() As String = {"prm", "rfb", "oth"}
        For Each item In form_types
            split_blank_cr_form(local.base_path & local.cr_blank_request_form_dir & "\" & local.cr_form, item, format, local, err)
            If Not err Like "" Then
                MsgBox("Can't create the blank cr templates, exiting...details: " & err)
                Forms.Application.Exit()
            End If
        Next

        'all is good, so start the timers
        '-------------------------------
        Timer_check_mail.Start()
        Timer_get_email_timeout_timer.Stop()

        'use this for testing
        '        crms_control.check_cr_status_times = {"2", "6", "10", "14", "18", "22"}


        '      Dim tx_svr As New smtp_server
        '     Dim rx_svr_imap As New imap_server
        '      system_analysis(Me, local, crms_control, format, rx_svr_imap, rx_svr_pop, tx_svr, err)

    End Sub





    'events and timers
    '-------------------
    Friend Sub Timer_check_mail_Tick(sender As Object, e As EventArgs) Handles Timer_check_mail.Tick
        On Error Resume Next
        crms_control.check_mail_timer += 1
        Me.Label_timer.Text = "Check Mail Timer: " & crms_control.check_mail_timer & " s"
        Me.Update()
        If crms_control.check_mail_timer >= crms_control.check_mail_period Then
            'do date and HW checks
            '---------------------
            run_checks()

            Timer_check_mail.Stop()

            'check if we need to run the CR supervisor now
            '-------------------------------------------------
            If Not Now.DayOfWeek = DayOfWeek.Saturday And Not Now.DayOfWeek = DayOfWeek.Sunday Then
                For Each checkpoint In crms_control.check_cr_status_times
                    Dim hour_now As Integer = Hour(Now)
                    If hour_now = Val(checkpoint) Then
                        time_old = time_new
                        time_new = hour_now
                        If Not time_old = time_new Then
                            crms_control.check_mail_timer = 0
                            crms_control.get_mail_timeout_cnt = 0
                            Label_timer.Text = "Running CR Review Processes, Please Wait..."
                            Me.Update()
                            Timer_get_email_timeout_timer.Start()
                            Dim err As String = ""
                            If format.scheduled_task_switch Then periodic_check_launcher(Me, local, crms_control, format, err)
                        End If
                    End If
                Next
            End If

            'check if we need to run the Google Drive Downloader now
            'basically if we are not downloading files, we check again if there are new files to DL
            '------------------------------------------------------------
            '            If Not BackgroundWorker_get_files_from_google_drive.IsBusy Then
            '            google_drive_login_data.clientid = Me.TextBox_GD_clientId.Text
            '           google_drive_login_data.clientsecret = Me.TextBox_GD_clientsecret.Text
            '            BackgroundWorker_get_files_from_google_drive.RunWorkerAsync()
            '        End If

            'check mail
            '-----------------
            Timer_get_email_timeout_timer.Stop()
            crms_control.check_mail_timer = 0
            crms_control.get_mail_timeout_cnt = 0
            Label_timer.Text = "Check Mail Timer: Checking Mail..."
            Label_timeout_timer.Text = "Get Mail Timeout Timer: " & crms_control.get_mail_timeout_cnt & " s"
            Me.Update()
            Timer_get_email_timeout_timer.Start()
            BackgroundWorker_get_email.RunWorkerAsync()

        End If
    End Sub


    Friend Sub BackgroundWorker_get_email_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker_get_email.DoWork
        On Error Resume Next
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", False)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US", False)

        Dim outer_err As String = ""
        Debug.WriteLine(Now.ToLongTimeString & ": " & "periodic email check-start")
        Dim retry_counter As Integer = 0
retry:
        outer_err = ""
        module_email.get_email_launcher(Me, local, crms_control, format, outer_err)
        If Not outer_err Like "" Then
            retry_counter += 1
            Debug.WriteLine(Now.ToLongTimeString & ": " & "Network Issues are preventing connection to email server, retrying 10x.....")
            System.Threading.Thread.Sleep(format.email_retry_backoff * 1000)
            If retry_counter < 10 Then
                GoTo retry
            End If
        End If
        GC.Collect()
    End Sub

    Private Sub BackgroundWorker_get_email_progresschanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker_get_email.ProgressChanged
        On Error Resume Next

    End Sub


    Private Sub BackgroundWorker_get_email_complete(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker_get_email.RunWorkerCompleted
        On Error Resume Next
        crms_control.get_mail_connect_times += 1
        If crms_control.get_mail_connect_times > crms_control.get_mail_connect_times_max Then
            Label_timer.Text = "App is restarting..."
            Me.Update()
            Dim file As New StreamWriter("C:\CRMS\restart_info.csv", False)
            file.WriteLine("""The app was restarted at " & Now.ToLongDateString & " : " & Now.ToLongTimeString & " after 100 get mail connections, just to clear out issues, not caused by errors.""")
            file.Close()
            file.Dispose()
            System.Threading.Thread.Sleep(3000)
            Forms.Application.Restart()
        End If
        Timer_get_email_timeout_timer.Stop()
        Timer_check_mail.Start()
        crms_control.get_mail_timeout_cnt = 0
        Me.Label_timeout_timer.Text = "Get Mail Timeout Timer: " & crms_control.get_mail_timeout_cnt & " s"
        Me.Update()
        Debug.WriteLine(Now.ToLongTimeString & ": " & "periodic email check-stop")
        busy_flag = True
    End Sub




    'trying this other code that should id if we have gone non-responsive and restart
    Private Sub Timer_get_email_timeout_timer_Tick(sender As Object, e As EventArgs) Handles Timer_get_email_timeout_timer.Tick
        On Error Resume Next
        crms_control.get_mail_timeout_cnt += 1
        Me.Label_timeout_timer.Text = "Get Mail Timeout Timer: " & crms_control.get_mail_timeout_cnt & " s"
        Me.Update()
        If crms_control.get_mail_timeout_cnt >= crms_control.get_mail_timeout_cnt_max Then
            crms_control.get_mail_timeout_cnt = 0
            Label_timer.Text = "Get Mail Timeout Timer: Timeout!!  Restarting App..."
            Me.Update()
            '######################################################
            '                KillNonResponsiveImageProcessByName("CRMS")
            '######################################################
            Dim file As New StreamWriter("C:\CRMS\restart_info.csv", False)
            file.WriteLine("""The app was restarted at " & Now.ToLongDateString & " : " & Now.ToLongTimeString & " due to get email timeout, most likely caused by a network error.""")
            file.Close()
            file.Dispose()
            System.Threading.Thread.Sleep(3000)
            Forms.Application.Restart()
        End If
    End Sub




    'background worker for getting files from GD
    '-----------------------------------------------
    Private Sub BackgroundWorker_get_files_from_google_drive_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker_get_files_from_google_drive.DoWork
        On Error Resume Next
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", False)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US", False)

        Dim outer_err As String = ""
        Debug.WriteLine(Now.ToLongTimeString & ": " & "############## G O O G L E   D R I V E   I S   S T A R T I N G   D O W N L O A D I N G ###############")
        module_google_drive.get_files_from_drive(google_drive_login_data.clientid, google_drive_login_data.clientsecret, local.base_path & "\GD incoming", BackgroundWorker_get_files_from_google_drive, outer_err)
        If Not outer_err Like "" Then
            Debug.WriteLine(Now.ToLongTimeString & ": " & "############## G O O G L E   D R I V E   H A D   D O W N L O A D   E R R O R S ###############")
        End If
        GC.Collect()

    End Sub

    Private Sub BackgroundWorker_get_files_from_google_drive_progresschanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker_get_files_from_google_drive.ProgressChanged
        On Error Resume Next
        Debug.WriteLine(Now.ToLongTimeString & ": " & "############## G O O G L E   D R I V E   H A S   D O N E  " & e.ProgressPercentage & " %  O F   T O T A L   D A T A ###############: " & e.UserState.ToString)
    End Sub

    Private Sub BackgroundWorker_get_files_from_google_drive_complete(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker_get_files_from_google_drive.RunWorkerCompleted
        On Error Resume Next
        Debug.WriteLine(Now.ToLongTimeString & ": " & "############## G O O G L E   D R I V E   H A S   F I N I S H E D   D O W N L O A D I N G ###############")
    End Sub








    '----------------------------------------------------------------------
    '----------------------------------------------------------------------
    '----------------------------------------------------------------------
    '----------------------------------------------------------------------
    'settings
    'check that the path is still valid after the user has selected the textbox and left it
    '-----------------------------------------------------------------------------------------
    Private Sub TextBox_base_path_leave(sender As Object, e As EventArgs) Handles TextBox_base_path.Leave
        On Error Resume Next
        If Not Me.TextBox_base_path.Text = local.old_base_path AndAlso Not is_path_valid(Me.TextBox_base_path.Text) Then
            MsgBox("The path entered doesn't exist")
            Me.TextBox_base_path.Text = local.old_base_path
        ElseIf Not Me.TextBox_base_path.Text = local.old_base_path AndAlso is_path_valid(Me.TextBox_base_path.Text) Then
            'then we have changed the base path and must re-create all the component folders
            '----------------------------------------------------------------------------------
            MsgBox("You have changed your base path, all component folders will be created if they do not yet exist.")
            local.base_path = Me.TextBox_base_path.Text
            local.check_directories(local)
        End If
    End Sub


    Private Sub Button_base_browse_Click(sender As Object, e As EventArgs) Handles Button_base_browse.Click
        On Error Resume Next
        Dim s1 As String = ""
        Dim err As String = ""
        s1 = get_path_from_user(Me.TextBox_base_path.Text, "Choose Email Working Directory", err)
        If s1 <> "" Then
            Me.TextBox_base_path.Text = s1
        End If
    End Sub


    'Change Settings
    'updates all the app settings when the user presses the submit button, if a background process has launched to get emails, these changes will take effect the next time it is run
    '-------------------------------------------------------------------------
    Private Sub Button_settings_submit_Click(sender As Object, e As EventArgs) Handles Button_settings_submit.Click
        On Error Resume Next
        crms_control.check_mail_period = Val(TextBox_check_mail_period.Text)
        crms_control.check_cr_status_times = Strings.Split(TextBox_cr_status_times.Text, ",")
        local.base_path = TextBox_base_path.Text
        format.use_imap = Me.CheckBox_use_imap.Checked
        write2settings()
    End Sub



    'writes to the settings file, if it exists before the write, it is deleted
    '---------------------------------------------------------------
    Private Sub write2settings()
        Try
            'write settings to settings file
            '-------------------------------
            If Not FileIO.FileSystem.DirectoryExists(local.base_path) Then
                MsgBox("The settings directory doesn't exist, please ensure it exists before running the app, if you have a settings file, ensure it is called (crms_settings.csv) and is in this directory: " & local.base_path & "... exiting app...")
                Forms.Application.Exit()
            ElseIf FileIO.FileSystem.FileExists(local.base_path & "\" & local.settings_file) Then
                Dim err As String = ""
                force_delete_file(local.base_path & "\" & local.settings_file, err)
                If Not err Like "" Then
                    MsgBox("The settings file is already open and can't be cleaned: " & local.base_path & "\" & local.settings_file & "... exiting app...")
                    Forms.Application.Exit()
                End If
            End If

            Dim writer As New System.IO.StreamWriter(local.base_path & "\" & local.settings_file)
            Dim t_s As String = ""
            writer.WriteLine("Start Settings...")
            writer.WriteLine("""Timer_check_mail.Interval (ms)"",""" & 1000 & """")
            writer.WriteLine("""Timer_get_email_timeout_timer.Interval (ms)"",""" & 1000 & """")
            writer.WriteLine("""Me.TextBox_base_path.Text"",""" & local.base_path & """")
            writer.WriteLine("""Me.TextBox_check_mail_period.Text"",""" & crms_control.check_mail_period & """")
            writer.WriteLine("""Me.TextBox_cr_status_times.Text""," & """" & Join(crms_control.check_cr_status_times, ",") & """")
            writer.WriteLine("""Me.CheckBox_use_imap.Checked""," & """" & Me.CheckBox_use_imap.Checked & """")
            writer.WriteLine("""Me.TextBox_imap_server.Text"",""" & Me.TextBox_imap_server.Text & """")
            writer.WriteLine("""Me.TextBox_imap_user.Text"",""" & Me.TextBox_imap_user.Text & """")
            writer.WriteLine("""Me.TextBox_imap_pass.Text"",""" & Me.TextBox_imap_pass.Text & """")
            writer.WriteLine("""Me.TextBox_imap_port.Text"",""" & Me.TextBox_imap_port.Text & """")
            writer.WriteLine("""Me.TextBox_imap_encryption_type.Text"",""" & Me.TextBox_imap_encryption_type.Text & """")
            writer.WriteLine("""Me.CheckBox_imap_ssl.Checked"",""" & Me.CheckBox_imap_ssl.Checked & """")
            writer.WriteLine("""Me.TextBox_pop_server.Text"",""" & Me.TextBox_pop_server.Text & """")
            writer.WriteLine("""Me.TextBox_pop_user.Text"",""" & Me.TextBox_pop_user.Text & """")
            writer.WriteLine("""Me.TextBox_pop_pass.Text"",""" & Me.TextBox_pop_pass.Text & """")
            writer.WriteLine("""Me.TextBox_pop_port.Text"",""" & Me.TextBox_pop_port.Text & """")
            writer.WriteLine("""Me.CheckBox_pop_use_ssl.Checked"",""" & Me.CheckBox_pop_use_ssl.Checked & """")
            writer.WriteLine("""Me.TextBox_smtp_server.Text"",""" & Me.TextBox_smtp_server.Text & """")
            writer.WriteLine("""Me.TextBox_smtp_user.Text"",""" & Me.TextBox_smtp_user.Text & """")
            writer.WriteLine("""Me.TextBox_smtp_pass.Text"",""" & Me.TextBox_smtp_pass.Text & """")
            writer.WriteLine("""Me.TextBox_smtp_timeout.Text"",""" & Me.TextBox_smtp_timeout.Text & """")
            writer.WriteLine("""Me.TextBox_smtp_port.Text"",""" & Me.TextBox_smtp_port.Text & """")
            writer.WriteLine("""Me.CheckBox_smtp_ssl.Checked"",""" & Me.CheckBox_smtp_ssl.Checked & """")
            writer.WriteLine("""Me.TextBox_mysql_server.Text"",""" & Me.TextBox_mysql_server.Text & """")
            writer.WriteLine("""Me.TextBox_mysql_port.Text"",""" & Me.TextBox_mysql_port.Text & """")
            writer.WriteLine("""Me.TextBox_mysql_user.Text"",""" & Me.TextBox_mysql_user.Text & """")
            writer.WriteLine("""Me.TextBox_mysql_password.Text"",""" & Me.TextBox_mysql_password.Text & """")
            writer.WriteLine("""Me.TextBox_mysql_db.Text"",""" & Me.TextBox_mysql_db.Text & """")
            writer.WriteLine("""Me.TextBox_GD_clientId.Text"",""" & Me.TextBox_GD_clientId.Text & """")
            writer.WriteLine("""Me.TextBox_GD_clientsecret.Text"",""" & Me.TextBox_GD_clientsecret.Text & """")

            writer.WriteLine("Stop Settings...")
            If Not writer Is Nothing Then writer.Close()
            writer.Dispose()

        Catch ex As Exception
            MsgBox("Some general error writing the settings to the HDD: " & local.base_path & "\" & local.settings_file & "... exiting app...")
            Forms.Application.Exit()
        End Try
    End Sub




    'reads from the settings file into the app vars
    '----------------------------------------------
    Private Sub read_from_settings()
        Try
            Dim input1() As String
            Dim input2(30) As String
            Dim i As Integer = 0

            Using reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(local.base_path & "\" & local.settings_file)
                reader.TextFieldType = FileIO.FieldType.Delimited
                reader.SetDelimiters(",")
                reader.HasFieldsEnclosedInQuotes = True   'this will treat any fields with enclosing quotes as 1 field regardless of what is in it

                'reads the file
                '---------------------
                Dim read_body As Boolean = False
                While Not reader.EndOfData
                    input1 = reader.ReadFields()
                    If input1.Count = 0 Then
                        GoTo skip
                    End If

                    '################################################
                    'test for a section header first
                    '----------------------------------
                    If Not read_body AndAlso Regex.IsMatch(Trim(input1(0)), "^(start\ssettings)(\.){3}$", RegexOptions.IgnoreCase) Then
                        read_body = True
                    ElseIf read_body AndAlso Regex.IsMatch(Trim(input1(0)), "^(stop\ssettings)(\.){3}$", RegexOptions.IgnoreCase) Then
                        read_body = False
                        Exit While
                    ElseIf read_body Then
                        input2(i) = input1(1)
                        i += 1
                    End If
skip:
                End While
            End Using

            'reads the vals to app vars
            '-------------------------------
            i = 0
            Timer_check_mail.Interval = Int(Val(input2(i)))
            i += 1
            Timer_get_email_timeout_timer.Interval = Int(Val(input2(i)))
            i += 1
            Me.TextBox_base_path.Text = input2(i)
            i += 1
            Me.TextBox_check_mail_period.Text = Int(Val(input2(i)))
            i += 1
            Me.TextBox_cr_status_times.Text = input2(i)
            i += 1
            Me.CheckBox_use_imap.Checked = If(Regex.IsMatch(input2(i), "True", RegexOptions.IgnoreCase), True, False)
            i += 1
            Me.TextBox_imap_server.Text = input2(i)
            i += 1
            Me.TextBox_imap_user.Text = input2(i)
            i += 1
            Me.TextBox_imap_pass.Text = input2(i)
            i += 1
            Me.TextBox_imap_port.Text = Int(Val(input2(i)))
            i += 1
            Me.TextBox_imap_encryption_type.Text = Int(Val(input2(i)))
            i += 1
            Me.CheckBox_imap_ssl.Checked = If(Regex.IsMatch(input2(i), "True", RegexOptions.IgnoreCase), True, False)
            i += 1
            Me.TextBox_pop_server.Text = input2(i)
            i += 1
            Me.TextBox_pop_user.Text = input2(i)
            i += 1
            Me.TextBox_pop_pass.Text = input2(i)
            i += 1
            Me.TextBox_pop_port.Text = Int(Val(input2(i)))
            i += 1
            Me.CheckBox_pop_use_ssl.Checked = If(Regex.IsMatch(input2(i), "True", RegexOptions.IgnoreCase), True, False)
            i += 1
            Me.TextBox_smtp_server.Text = input2(i)
            i += 1
            Me.TextBox_smtp_user.Text = input2(i)
            i += 1
            Me.TextBox_smtp_pass.Text = input2(i)
            i += 1
            Me.TextBox_smtp_timeout.Text = Int(Val(input2(i)))
            i += 1
            Me.TextBox_smtp_port.Text = Int(Val(input2(i)))
            i += 1
            Me.CheckBox_smtp_ssl.Checked = If(Regex.IsMatch(input2(i), "True", RegexOptions.IgnoreCase), True, False)
            i += 1
            Me.TextBox_mysql_server.Text = input2(i)
            i += 1
            Me.TextBox_mysql_port.Text = Int(Val(input2(i)))
            i += 1
            Me.TextBox_mysql_user.Text = input2(i)
            i += 1
            Me.TextBox_mysql_password.Text = input2(i)
            i += 1
            Me.TextBox_mysql_db.Text = input2(i)
            i += 1
            Me.TextBox_GD_clientId.Text = input2(i)
            i += 1
            Me.TextBox_GD_clientsecret.Text = input2(i)

        Catch ex As Exception
            MsgBox("The settings file is corrupted, falling back to internal defaults...")
            set_hardcoded_vals()
        End Try
    End Sub




    'does an app tables vs DB tables consistency check
    '----------------------------------------------------
    Private Sub format_vs_db_check(ByVal format As cr_sheet_format, ByRef err As String)
        Try
            Dim db As New CRMS.mysql_server
            db.address = Me.TextBox_mysql_server.Text
            db.port = Me.TextBox_mysql_port.Text
            db.username = Me.TextBox_mysql_user.Text
            db.password = Me.TextBox_mysql_password.Text
            db.schema = Me.TextBox_mysql_db.Text
            mysql_connect(db)

            Try
                Dim table As String = "cr_common"
                Dim t_array() As String = format.common_hdr_name.ToArray
                For Each item In t_array
                    t_array(Array.IndexOf(t_array, item)) = LCase(Regex.Replace(item, "\s", "_", RegexOptions.IgnoreCase))
                Next
                Dim dt As New System.Data.DataTable
                Dim sqltext As String = ""
                sqltext = "SELECT * FROM " & db.schema & "." & table & " WHERE cr_id like '-9999';"
                sqlquery(False, db, sqltext, dt, err)
                Dim i As Integer = 0
                Dim max As Integer = t_array.Count
                For Each col As System.Data.DataColumn In dt.Columns
                    If i > max - 1 Then
                        Exit For
                    End If
                    Dim test As String = col.ColumnName
                    If Array.Find(t_array, Function(s) s = test) Is Nothing Then
                        err = "Discrepancy found, col '" & test & "' doesn't exist in the app table '" & table & "'"
                        GoTo get_out
                    Else
                        'remove the matched col from the array, so we do not find it again
                        '------------------------------------------------------------
                        t_array(Array.FindIndex(t_array, Function(s) s = test)) = "throw this fucker out"
                        t_array = t_array.Except({"throw this fucker out"}).ToArray()
                    End If
                    i += 1
                Next
            Catch ex As Exception
            End Try

            Try
                Dim table As String = "cr_data_prm"
                Dim t_array() As String = format.detail_hdr_name_prm.ToArray
                For Each item In t_array
                    t_array(Array.IndexOf(t_array, item)) = LCase(Regex.Replace(item, "\s", "_", RegexOptions.IgnoreCase))
                Next
                Dim dt As New System.Data.DataTable
                Dim sqltext As String = ""
                sqltext = "SELECT * FROM " & db.schema & "." & table & " WHERE cr_sub_id like '-9999';"
                sqlquery(False, db, sqltext, dt, err)
                For Each col As System.Data.DataColumn In dt.Columns
                    Dim test As String = col.ColumnName
                    If Array.Find(t_array, Function(s) s = test) Is Nothing Then
                        err = "Discrepancy found, col '" & test & "' doesn't exist in the app table '" & table & "'"
                        GoTo get_out
                    Else
                        'remove the matched col from the array, so we do not find it again
                        '------------------------------------------------------------
                        t_array(Array.FindIndex(t_array, Function(s) s = test)) = "throw this fucker out"
                        t_array = t_array.Except({"throw this fucker out"}).ToArray()
                    End If
                Next
            Catch ex As Exception
            End Try

            Try
                Dim table As String = "cr_data_oth"
                Dim t_array() As String = format.detail_hdr_name_oth.ToArray
                For Each item In t_array
                    t_array(Array.IndexOf(t_array, item)) = LCase(Regex.Replace(item, "\s", "_", RegexOptions.IgnoreCase))
                Next
                Dim dt As New System.Data.DataTable
                Dim sqltext As String = ""
                sqltext = "SELECT * FROM " & db.schema & "." & table & " WHERE cr_sub_id like '-9999';"
                sqlquery(False, db, sqltext, dt, err)
                For Each col As System.Data.DataColumn In dt.Columns
                    Dim test As String = col.ColumnName
                    If Array.Find(t_array, Function(s) s = test) Is Nothing Then
                        err = "Discrepancy found, col '" & test & "' doesn't exist in the app table '" & table & "'"
                        GoTo get_out
                    Else
                        'remove the matched col from the array, so we do not find it again
                        '------------------------------------------------------------
                        t_array(Array.FindIndex(t_array, Function(s) s = test)) = "throw this fucker out"
                        t_array = t_array.Except({"throw this fucker out"}).ToArray()
                    End If
                Next
            Catch ex As Exception
            End Try

            Try
                Dim table As String = "cr_data_rfb"
                Dim t_array() As String = format.detail_hdr_name_rfb.ToArray
                Dim i As Integer = 0
                Dim pref() As String = {"cur_", "pro_", "act_", "fin_", "fin_"}
                For Each item In t_array
                    t_array(Array.IndexOf(t_array, item)) = LCase(Regex.Replace(item, "\s", "_", RegexOptions.IgnoreCase))
                Next
                For Each item In t_array
                    If Regex.IsMatch(item, "^(az)|(mdt)|(edt)|(ht)|(antenna)|(coax_len)$", RegexOptions.IgnoreCase) Then
                        t_array(Array.IndexOf(t_array, item)) = pref(i) & item
                        If Regex.IsMatch(item, "^edt$", RegexOptions.IgnoreCase) Then
                            i += 1
                        End If
                    End If
                Next
                Dim dt As New System.Data.DataTable
                Dim sqltext As String = ""
                sqltext = "SELECT * FROM " & db.schema & "." & table & " WHERE cr_sub_id like '-9999';"
                sqlquery(False, db, sqltext, dt, err)
                For Each col As System.Data.DataColumn In dt.Columns
                    Dim test As String = col.ColumnName
                    If Array.Find(t_array, Function(s) s = test) Is Nothing Then
                        err = "Discrepancy found, col '" & test & "' doesn't exist in the app table '" & table & "'"
                        GoTo get_out
                    Else
                        'remove the matched col from the array, so we do not find it again
                        '------------------------------------------------------------
                        t_array(Array.FindIndex(t_array, Function(s) s = test)) = "throw this fucker out"
                        t_array = t_array.Except({"throw this fucker out"}).ToArray()
                    End If
                Next
            Catch ex As Exception
            End Try
get_out:
        Catch ex As Exception
            err = "ER: There was a general error checking app tables vs the DB: " & ex.ToString
        End Try
    End Sub


    'does some checks
    Private Sub run_checks()
        'do date and HW checks
        '---------------------
        Try
            If Now > format.end_date Then
                MsgBox("App is out of date, contact nathan.scott.rf@gmail.com to extend...")
                Forms.Application.Exit()
            End If
            Dim check_pass As Boolean = False
            For Each nic As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()
                For Each Item In format.tied
                    If nic.Name & "," & nic.GetPhysicalAddress.ToString & "," & nic.Description Like Item Then
                        check_pass = True
                        Exit For
                    End If
                Next
            Next
            If Not check_pass Then
                MsgBox("App can only run on the machine it was intended for, contact nathan.scott.rf@gmail.com to change...")
                Forms.Application.Exit()
            End If
        Catch ex As Exception
        End Try
    End Sub



    Private Sub set_hardcoded_vals()
        On Error Resume Next
        Timer_check_mail.Interval = 1000       'in ms
        Timer_get_email_timeout_timer.Interval = 1000       'in ms

        Me.TextBox_base_path.Text = local.base_path

        Me.CheckBox_use_imap.Checked = False

        Me.TextBox_check_mail_period.Text = crms_control.check_mail_period
        Me.TextBox_cr_status_times.Text = Join(crms_control.check_cr_status_times, ",")
        Me.TextBox_check_mail_period.Text = crms_control.check_mail_period

        Me.TextBox_imap_server.Text = "imap.gmail.com"
        Me.TextBox_imap_user.Text = "crms.hw.tsel@gmail.com"
        Me.TextBox_imap_pass.Text = "huawei2015"
        '        Me.TextBox_imap_server.Text = "popsus.huawei.com"
        '        Me.TextBox_imap_user.Text = "tsel_cr@ms.huawei.com"
        '        Me.TextBox_imap_pass.Text = "p7-j4#mQ"
        Me.TextBox_imap_port.Text = 993
        Me.TextBox_imap_encryption_type.Text = 3072         'for huawei ms server, use 240 for gmail, but 3072 also works
        Me.CheckBox_imap_ssl.Checked = True

        Me.TextBox_pop_server.Text = "pop.gmail.com"
        Me.TextBox_pop_user.Text = "crms.hw.tsel@gmail.com"
        Me.TextBox_pop_pass.Text = "huawei2015"
        '        Me.TextBox_pop_server.Text = "popsus.huawei.com"
        '        Me.TextBox_pop_user.Text = "tsel_cr@ms.huawei.com"
        '        Me.TextBox_pop_pass.Text = "p7-j4#mQ"
        Me.TextBox_pop_port.Text = 995
        Me.CheckBox_pop_use_ssl.Checked = True

        Me.TextBox_smtp_server.Text = "smtp.gmail.com"
        Me.TextBox_smtp_user.Text = "crms.hw.tsel@gmail.com"
        Me.TextBox_smtp_pass.Text = "huawei2015"
        '        Me.TextBox_smtp_server.Text = "smtpsus.huawei.com"
        '        Me.TextBox_smtp_user.Text = "tsel_cr@ms.huawei.com"
        '        Me.TextBox_smtp_pass.Text = "p7-j4#mQ"
        Me.TextBox_smtp_timeout.Text = "60000"
        Me.TextBox_smtp_port.Text = 587   '465
        Me.CheckBox_smtp_ssl.Checked = True

        Me.TextBox_mysql_server.Text = "127.0.0.1"
        Me.TextBox_mysql_port.Text = 3306
        Me.TextBox_mysql_user.Text = "crms_user"
        Me.TextBox_mysql_password.Text = "nathan"
        Me.TextBox_mysql_db.Text = "crms_db"

        Me.TextBox_GD_clientId.Text = google_drive_login_data.clientid
        Me.TextBox_GD_clientsecret.Text = google_drive_login_data.clientsecret
    End Sub






    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next
        Me.WindowState = FormWindowState.Minimized
        Me.ShowInTaskbar = False
        NotifyIcon1.Visible = True
        Me.Hide()
        '        Me.Show()
        '       Me.WindowState = FormWindowState.Normal
        '      NotifyIcon1.Visible = False
        '     Me.ShowInTaskbar = True
    End Sub



    Private Sub notifyicon1_baloon(sender As Object, e As EventArgs) Handles NotifyIcon1.BalloonTipClicked
        On Error Resume Next
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        NotifyIcon1.Visible = False
        Me.ShowInTaskbar = True
    End Sub


    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        On Error Resume Next
        NotifyIcon1.ShowBalloonTip(2000, "CRMS: Really, just stop clicking me", "I'm still running so leave me alone and go away :)", ToolTipIcon.Info)
    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        On Error Resume Next
        If Me.WindowState = FormWindowState.Minimized Then
            Me.WindowState = FormWindowState.Minimized
            Me.ShowInTaskbar = False
            NotifyIcon1.Visible = True
            Me.Hide()
        End If
    End Sub


    'makes the x button disabled => you can't close the app, only hide it
    '------------------------------------------------
    Private Const CP_NOCLOSE_BUTTON As Integer = &H200
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim myCp As CreateParams = MyBase.CreateParams
            myCp.ClassStyle = myCp.ClassStyle Or CP_NOCLOSE_BUTTON
            Return myCp
        End Get
    End Property


End Class
