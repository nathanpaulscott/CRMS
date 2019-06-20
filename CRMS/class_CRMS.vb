Public Class email_msg
    Public Class mail_spec
        Public displayname As String
        Public address As String
        Public Sub New(ByVal namespec As String, ByVal addressspec As String)
            displayname = namespec
            address = addressspec
        End Sub
    End Class
    Public Class attachment
        Public filename As String
        Public size As Double
        Public Sub New(ByVal filenamespec As String, ByVal sizespec As Double)
            filename = filenamespec
            size = sizespec
        End Sub
    End Class
    Public from As mail_spec
    Public to_list() As mail_spec
    Public cc_list() As mail_spec
    Public msg_date As DateTime
    Public subject_raw As String
    Public body_text_raw As String
    Public body_html_raw As String
    Public cleansubject As String
    Public cleanbody_text As String
    Public cleanbody_html As String
    Public attachments() As attachment
End Class

Public Class imap_server
    Inherits ImapX.ImapClient
    Public nat_delete_after_dl_flag As Boolean
    Public nat_username As String
    Public nat_password As String
    Public nat_use_idle_mode As Boolean
    Public nat_server_folder As String
    Public nat_cred_regular As ImapX.Authentication.PlainCredentials
    Public nat_cred_oauth2 As ImapX.Authentication.OAuth2Credentials
End Class


Public Class pop_server
    Inherits Pop3Client
    Public nat_server As String
    Public nat_port As Integer
    Public nat_use_ssl As Boolean
    Public nat_username As String
    Public nat_password As String
End Class

Public Class smtp_server
    Inherits SmtpClient
    Public nat_send_async_flag As Boolean
    Public ReadOnly nat_is_connected_flag As Boolean
    Public send_retry_times As Integer = 100000000
    Public send_retry_timeout As Integer = 60
    Public nat_username = ""
    Public nat_password = ""
    Public email_retry_backoff As Integer = 10           's
End Class

Public Class local_machine
    Public base_path As String = "C:\CRMS"
    Public old_base_path As String = "C:\CRMS"
    Public site_binders As String = "\site binders"
    Public inbox As String = "\inbox"
    Public cr As String = "\cr"
    Public outbox As String = "\outbox"
    Public sent As String = "\sent"
    Public db_incoming As String = "\db incoming"
    Public db_outgoing As String = "\db outgoing"
    Public temp As String = "\temp"
    Public cr_blank_request_form_dir As String = "\CR Creation Form Store"
    Public settings_file As String = "crms_settings.csv"
    Public cr_form As String = "New CR Form.xlsb"
    Public base_path_freespace_warning As Integer = 1       'GB
    Public db_drive_freespace_warning As Integer = 1     'GB
    Public db_drive As String = "C:\"

    Public Sub check_directories(ByVal local As local_machine)
        If Not FileIO.FileSystem.DirectoryExists(local.base_path & local.inbox) Then FileIO.FileSystem.CreateDirectory(local.base_path & local.inbox)
        If Not FileIO.FileSystem.DirectoryExists(local.base_path & local.cr) Then FileIO.FileSystem.CreateDirectory(local.base_path & local.cr)
        If Not FileIO.FileSystem.DirectoryExists(local.base_path & local.outbox) Then FileIO.FileSystem.CreateDirectory(local.base_path & local.outbox)
        If Not FileIO.FileSystem.DirectoryExists(local.base_path & local.sent) Then FileIO.FileSystem.CreateDirectory(local.base_path & local.sent)
        If Not FileIO.FileSystem.DirectoryExists(local.base_path & local.db_incoming) Then FileIO.FileSystem.CreateDirectory(local.base_path & local.db_incoming)
        If Not FileIO.FileSystem.DirectoryExists(local.base_path & local.db_outgoing) Then FileIO.FileSystem.CreateDirectory(local.base_path & local.db_outgoing)
        If Not FileIO.FileSystem.DirectoryExists(local.base_path & local.site_binders) Then FileIO.FileSystem.CreateDirectory(local.base_path & local.site_binders)
        If Not FileIO.FileSystem.DirectoryExists(local.base_path & local.temp) Then FileIO.FileSystem.CreateDirectory(local.base_path & local.temp)
    End Sub
End Class


Public Class mysql_server
    Public address As String
    Public port As Integer
    Public username As String
    Public password As String
    Public schema As String
    Public table As String
    Public mysql_con_string As String
    Public mysql_con As New MySqlConnection
    Public local_folder As String
    Public local_dgv As DataGridView
    Public local_dt As System.Data.DataTable
    Public big_command_timeout As Integer = 60 * 3   '3 mins
    Protected Friend mysql_max_rows_per_command As Integer = 1000       'connection craps itself any bigger than this, can play wiht max packet size on both sides of the link, but not worth it
    Public ReadOnly is_connected As Boolean
End Class



Public Class google_drive_creds
    Public clientid As String = "178984549041-c32cqer2tk1g7qt5tulfn5ct93gc73f4.apps.googleusercontent.com"
    Public clientsecret As String = "5It9tQP7grGdquWX3J-JxFkH"
    '    Public clientid As String = "134685934627-s9p9o2e5fvgsiv50o0uu6kqua6plim7a.apps.googleusercontent.com"
    '   Public clientsecret As String = "GXnihzFLhVPVN5UNY-mckrwj"
End Class


Public Class crms_controller
    Public check_mail_timer As Integer = 0      'in secs
    Public check_mail_period As Double = 30    'in secs
    Public check_cr_status_times() As String = {"6", "12", "18"}    'oclock
    Public mail_svr_retry_period As Integer = 10    'in secs
    Public hdd_retry_period As Integer = 10    'in secs
    Public mysql_db_retry_period As Integer = 10    'in secs
    Public check_mail_retry_times As Integer = 10   'times
    Public get_mail_timeout_cnt As Integer = 0                'in secs => the actual timer count
    Public get_mail_timeout_cnt_max As Integer = 60 * 20     'in secs => (20 mins) this sets the timeout for one get email run, if it takes longer than this, the app is restarted
    Public get_mail_connect_times As Integer = 0            'the check email counter
    Public get_mail_connect_times_max As Integer = 10000      'this sets the number of times we run the get mail sub before restarting the app to clean out weird bugs, 100 is about 1 hr with minimial load, so I have it as 100 hrs
End Class



Public Class cr_sheet_format
    'main debug controls
    Protected Friend debug_xl As Boolean = False
    Protected Friend send_opening_mail As Boolean = True
    Protected Friend scheduled_task_switch As Boolean = True
    Protected Friend send_internal_errors_to_user As Boolean = True

    'other VIP controls
    Protected Friend sheet_row_limit_rfb As Integer = 1000
    Protected Friend sheet_row_limit_global As Integer = 200000
    Protected Friend email_retry_backoff As Integer = 10           's
    Protected Friend use_imap As Boolean = False
    Protected Friend dev_email As String = "nathan.scott.rf@gmail.com"
    Protected Friend admin_x As String = "test1234"
    Protected Friend pm_x As String = "test4321"
    Protected Friend end_date As Date = New DateTime(2020, 1, 1)                  'DateTime.Now.AddMonths(1)
    Protected Friend tied() As String = {"Local Area Connection,D4BED9437105,Realtek PCIe GBE Family Controller", "Local Area Connection,448A5B7B8F62,Intel(R) Ethernet Connection I217-LM"}
    Protected Friend x_factor As String = "hjv234jc4i8yi54chugknGGgfjy5TGI4ghif"
    Protected Friend x_factor_email As String = "wwek9relkm4UemrgJdfgkEegk48h"
    Protected Friend planned_ex_time_delay_same_day As Integer = 6       'hrs => same day planned CRs have this time to execute before nagging starts =>eg. 6 hours + now
    Protected Friend planned_ex_time_future As Integer = 18              'hundred hrs => future dated ex planned CRs have this as the cuttoff time on the planned date
    Protected Friend cr_form_size_zip_limit As Integer = 1                'any cr_form above this size in MB needs to be zipped
    Protected Friend max_cells_per_chunk As Integer = 500 * 1000          'normally 1e6 is ok, bu tI am being conservative, this is for read/write from memory DT to/from XL

    'cc control for {requester, approver, ex.planner, executors, cc_list}
    'this will allow or stop that user from being cc'd at that particular stage
    '------------------------------------------------------------------
    Public cc_mask_resubmission_request() As Boolean = {False, False, False, False, False}
    Public cc_mask_approval_request() As Boolean = {True, False, False, False, True}
    Public cc_mask_execution_planning_request() As Boolean = {True, False, False, False, True}
    Public cc_mask_execution_request() As Boolean = {True, False, False, False, True}
    Public cc_mask_review_request() As Boolean = {True, False, False, False, True}
    Public cc_mask_everyone() As Boolean = {True, True, True, True, True}

    'these need to be set relative to "A2" and the actual offsets on the cr form
    '---------------------------------------------------------
    Public common_hdr_col As Integer = 1
    Public common_data_col As Integer = 3
    Public common_row_start As Integer = 2

    Public detail_hdr_row_start As Integer = 23
    Public detail_data_row_start As Integer = 25
    Public detail_col_start As Integer = 1

    Public row_hide_for_new_fail() As Integer = {2, 7, 8, 12, 14, 15, 16, 17, 18, 19, 20}
    Public col_hide_for_new_fail() As Integer = {17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30}

    Public row_hide_for_app() As Integer = {15, 17, 18, 19, 20}
    Public col_hide_prm_for_app() As Integer = {11, 12, 13, 14, 15, 16}
    Public col_hide_rfb_for_app() As Integer = {13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26}
    Public col_hide_oth_for_app() As Integer = {7, 8, 9, 10, 11}

    Public row_hide_for_ex_coord() As Integer = {15, 18, 19, 20}
    Public col_hide_prm_for_ex_coord() As Integer = {13, 14, 15, 16}
    Public col_hide_rfb_for_ex_coord() As Integer = {15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26}
    Public col_hide_oth_for_ex_coord() As Integer = {9, 10, 11}
    Public col_unprotect_prm_for_ex_coord() As Integer = {11, 12}
    Public col_unprotect_rfb_for_ex_coord() As Integer = {13, 14}
    Public col_unprotect_oth_for_ex_coord() As Integer = {7, 8}

    Public row_hide_for_ex() As Integer = {19, 20}
    Public col_hide_prm_for_ex() As Integer = {16}
    Public col_hide_rfb_for_ex() As Integer = {}
    Public col_hide_oth_for_ex() As Integer = {}
    Public col_unprotect_prm_for_ex() As Integer = {12, 13, 14, 15}
    Public col_unprotect_rfb_for_ex() As Integer = {14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26}
    Public col_unprotect_oth_for_ex() As Integer = {8, 9, 10, 11}

    Public row_hide_for_rev() As Integer = {20}

    Public row_clear_for_resub() As Integer = {14, 15, 16, 17, 18, 19, 20}
    Public row_hide_for_resub() As Integer = {7, 8, 12, 14, 15, 16, 17, 18, 19, 20}
    Public row_hide_for_bad_resub() As Integer = {15, 17, 18, 19, 20}
    Public row_unprotect_for_resub() As Integer = {4, 9, 10, 11, 13}
    Public col_clear_hide_prm_for_resub() As Integer = {11, 12, 13, 14, 15, 16}
    Public col_clear_hide_rfb_for_resub() As Integer = {13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26}
    Public col_clear_hide_oth_for_resub() As Integer = {7, 8, 9, 10, 11}
    Public col_unprotect_prm_for_resub() As Integer = {2, 3, 4, 5, 6, 7, 8, 9, 10}
    Public col_unprotect_rfb_for_resub() As Integer = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}
    Public col_unprotect_oth_for_resub() As Integer = {2, 3, 4, 5, 6}


    'these are the header names, the cr form must have headers in this order
    '----------------------------------
    Public common_hdr_name() As String = {"CR_ID", _
                                       "Technology", _
                                       "CR Objective", _
                                       "Team", _
                                       "Region", _
                                       "CR Type", _
                                       "Node Types", _
                                       "Issue Description", _
                                       "Expected Impact", _
                                       "Risk", _
                                       "Requester", _
                                       "Approver", _
                                       "Execution Coordinator", _
                                       "Executors", _
                                       "Open Date", _
                                       "Approval Date", _
                                       "Planned Execution Date", _
                                       "Execution Date", _
                                       "Closed Date"}

    'these are the header data restrictions
    '----------------------------------
    Public common_data_restriction() As String = {"", _
                                       "tablematch,tech,tech", _
                                       "tablematch,depends on tech,cr_objective", _
                                       "tablematch,teams,team", _
                                       "tablematch,geo,region", _
                                       "", _
                                       "", _
                                       "not blank", _
                                       "not blank", _
                                       "tablematch,risk,risk", _
                                       "", _
                                       "tablematch,approvers,combined_name", _
                                       "", _
                                       "", _
                                       "", _
                                       "", _
                                       "", _
                                       "", _
                                       ""}

    'these are the detailed data header names, the cr form must have headers in this order
    '----------------------------------
    Public detail_hdr_name() As String = {"CR_sub_ID", _
                                        "CR Type", _
                                        "Node Type", _
                                        "Node", _
                                        "Nbr Node", _
                                        "Parameter", _
                                        "Proposed Setting", _
                                        "Rollback Setting", _
                                        "AZ", _
                                        "MDT", _
                                        "EDT", _
                                        "AZ", _
                                        "MDT", _
                                        "EDT", _
                                        "Requester Comments", _
                                        "Execution Coordinator", _
                                        "Planned Execution Date", _
                                        "Executor", _
                                        "AZ", _
                                        "MDT", _
                                        "EDT", _
                                        "AZ", _
                                        "MDT", _
                                        "EDT", _
                                        "Ht", _
                                        "Antenna", _
                                        "Coax Len", _
                                        "Execution Status", _
                                        "Execution Date", _
                                        "Executor Comments"}

    'these are the detailed data header data restrictions when the requester opens a new cr
    '----------------------------------------------------------------------------------
    Public detail_data_restriction_initial() As String = {"", _
                                            "tablematch,cr_types,cr_type", _
                                            "tablematch,node_types,node_type", _
                                            "not blank", _
                                            "", _
                                            "", _
                                            "cr type is parameter and not blank", _
                                            "cr type is parameter and not blank", _
                                            "cr type is rf basic and val: 0, 359", _
                                            "cr type is rf basic and val: 0, 50", _
                                            "cr type is rf basic and 2val: 0, 50", _
                                            "cr type is rf basic and val: 0, 359", _
                                            "cr type is rf basic and val: 0, 50", _
                                            "cr type is rf basic and 2val: 0, 50", _
                                            "", _
                                            "tablematch,depends on cr_type,combined_name", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            ""}

    'these are the detailed data header names for the parameter and HW cr type
    '-----------------------------------------------------------------
    Public detail_hdr_name_prm() As String = {"CR_sub_ID", _
                                        "CR Type", _
                                        "Node Type", _
                                        "Node", _
                                        "Nbr Node", _
                                        "Parameter", _
                                        "Proposed Setting", _
                                        "Rollback Setting", _
                                        "Requester Comments", _
                                        "Execution Coordinator", _
                                        "Planned Execution Date", _
                                        "Executor", _
                                        "Execution Status", _
                                        "Execution Date", _
                                        "Executor Comments"}

    'these are the detailed data header data restrictions when the requester resubmits this type of cr_form
    '-----------------------------------------------------------------------
    Public detail_data_restriction_prm_resubmit() As String = {"", _
                                            "tablematch,cr_types,cr_type", _
                                            "tablematch,node_types,node_type", _
                                            "not blank", _
                                            "", _
                                            "", _
                                            "cr type is parameter and not blank", _
                                            "cr type is parameter and not blank", _
                                            "", _
                                            "tablematch,depends on cr_type,combined_name", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            ""}
    'these are the detailed data header names for the parameter/hw CR type
    '-----------------------------------------------------------------------
    Public detail_data_restriction_prm_from_ex() As String = {"", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "tablematch,execution_status,status", _
                                        "not blank if execution_status is executed", _
                                        "not blank if execution_status is fail and there is no attachment"}

    'these are the detailed data header names for the parameter and HW cr type
    '-----------------------------------------------------------------
    Public detail_hdr_name_oth() As String = {"CR_sub_ID", _
                                        "CR Type", _
                                        "Node Type", _
                                        "Node", _
                                        "Requester Comments", _
                                        "Execution Coordinator", _
                                        "Planned Execution Date", _
                                        "Executor", _
                                        "Execution Status", _
                                        "Execution Date", _
                                        "Executor Comments"}

    'these are the detailed data header data restrictions when the requester resubmits this type of cr_form
    '-----------------------------------------------------------------------
    Public detail_data_restriction_oth_resubmit() As String = {"", _
                                            "tablematch,cr_types,cr_type", _
                                            "tablematch,node_types,node_type", _
                                            "not blank", _
                                            "", _
                                            "tablematch,depends on cr_type,combined_name", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            ""}

    'these are the detailed data header names for the parameter/hw CR type
    '-----------------------------------------------------------------------
    Public detail_data_restriction_oth_from_ex() As String = {"", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "", _
                                        "either comment or attachment must exist"}

    'these are the detailed data header names for the RF CR type
    '-----------------------------------------------------------------------
    Public detail_hdr_name_rfb() As String = {"CR_sub_ID", _
                                        "CR Type", _
                                        "Node Type", _
                                        "Node", _
                                        "AZ", _
                                        "MDT", _
                                        "EDT", _
                                        "AZ", _
                                        "MDT", _
                                        "EDT", _
                                        "Requester Comments", _
                                        "Execution Coordinator", _
                                        "Planned Execution Date", _
                                        "Executor", _
                                        "AZ", _
                                        "MDT", _
                                        "EDT", _
                                        "AZ", _
                                        "MDT", _
                                        "EDT", _
                                        "Ht", _
                                        "Antenna", _
                                        "Coax Len", _
                                        "Execution Status", _
                                        "Execution Date", _
                                        "Executor Comments"}

    'these are the detailed data header data restrictions when the requester resubmits this type of cr_form
    '----------------------------------
    Public detail_data_restriction_rfb_resubmit() As String = {"", _
                                            "tablematch,cr_types,cr_type", _
                                            "tablematch,node_types,node_type", _
                                            "not blank", _
                                            "cr type is rf basic and val: 0, 359", _
                                            "cr type is rf basic and val: 0, 50", _
                                            "cr type is rf basic and 2val: 0, 50", _
                                            "cr type is rf basic and val: 0, 359", _
                                            "cr type is rf basic and val: 0, 50", _
                                            "cr type is rf basic and 2val: 0, 50", _
                                            "", _
                                            "tablematch,depends on cr_type,combined_name", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            ""}

    'these are the detailed data header data restrictions when the executor sends back the completed CR form with data filled out
    '----------------------------------
    Public detail_data_restriction_rfb_from_ex() As String = {"", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "", _
                                            "depends on executor and val: 0, 359", _
                                            "depends on executor and val: 0, 50", _
                                            "depends on executor and 2val: 0, 50", _
                                            "depends on executor and val: 0, 359", _
                                            "depends on executor and val: 0, 50", _
                                            "depends on executor and 2val: 0, 50", _
                                            "depends on executor and val: 0,200", _
                                            "depends on executor and tablematch,antennas,lookup", _
                                            "depends on executor and val: 0,200", _
                                            "tablematch,execution_status,status", _
                                            "not blank if execution_status is executed", _
                                            "not blank if execution_status is fail and there is no attachment"}


    'these are the datatable definitions to hold all the allowed values
    '--------------------------------------------------------
    Public ds_allow As New System.Data.DataSet

    Public Sub create_datatables(ByRef ds As System.Data.DataSet)
        ds_allow.Tables.Add("state_control")
        ds_allow.Tables.Add("antennas")
        ds_allow.Tables.Add("geo")
        ds_allow.Tables.Add("risk")
        ds_allow.Tables.Add("teams")
        ds_allow.Tables.Add("tech")
        ds_allow.Tables.Add("cr_obj_2g")
        ds_allow.Tables.Add("cr_obj_3g")
        ds_allow.Tables.Add("cr_obj_4g")
        ds_allow.Tables.Add("cr_types")
        ds_allow.Tables.Add("node_types")
        ds_allow.Tables.Add("requesters")
        ds_allow.Tables.Add("approvers")
        ds_allow.Tables.Add("rfb_ex_coord")
        ds_allow.Tables.Add("rfr_ex_coord")
        ds_allow.Tables.Add("hdw_ex_coord")
        ds_allow.Tables.Add("prm_ex_coord")
        ds_allow.Tables.Add("executors")
        ds_allow.Tables.Add("administrators")
    End Sub


    'verifying syntax of email addresses
    'this is what I originally used, but it doesn't work
    '    Protected Friend email_pattern As New Regex("^[-_a-z0-9]+(\.[-a-z0-9]+)@[-a-z0-9]+(\.[-a-z0-9]+)*(\.[a-z]{2,4})$", RegexOptions.IgnoreCase)

    'this is from microsoft
    '############################################################
    Dim email_validation As Boolean = False
    Public Function IsValidEmail(strIn As String) As Boolean
        email_validation = False
        If String.IsNullOrEmpty(strIn) Then Return False

        ' Use IdnMapping class to convert Unicode domain names. 
        Try
            strIn = Regex.Replace(strIn, "(@)(.+)$", AddressOf Me.DomainMapper, RegexOptions.None, TimeSpan.FromMilliseconds(200))
        Catch e As RegexMatchTimeoutException
            Return False
        End Try

        If email_validation Then Return False

        ' Return true if strIn is in valid e-mail format. 
        Try
            Return Regex.IsMatch(strIn,
                   "^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                   "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                   RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250))
        Catch e As RegexMatchTimeoutException
            Return False
        End Try
    End Function

    Private Function DomainMapper(match As Match) As String
        ' IdnMapping class with default property values. 
        Dim idn As New IdnMapping()

        Dim domainName As String = match.Groups(2).Value
        Try
            domainName = idn.GetAscii(domainName)
        Catch e As ArgumentException
            email_validation = True
        End Try
        Return match.Groups(1).Value + domainName
    End Function
    '############################################################

End Class