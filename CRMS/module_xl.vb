Option Compare Text

'            For Each item In System.Drawing.Printing.PrinterSettings.InstalledPrinters
'               if regex.ismatch(item,"nitro") then
'                   xlapp.ActivePrinter = item
'               end if
'            Next
'           xlsheet.PrintOutEx(1, 1, 1, False, xlapp.ActivePrinter, False, False, "c:\text.pdf", False)
'             to print, you have to manually find the code and port for a particular printer in xl, then hardcode it into the tool, it can not be done automatically.


Module xl
    'this checks the initial cr request form format meets the cr format requirements and that all attachments mentioned, exist in the email attachments.
    '------------------------------------------------------------------------
    Public Sub check_requester_cr_format(ByRef data_ok As String, ByVal cr_id As String, ByRef ds As System.Data.DataSet, ByVal cr_form_temp As String, ByVal cr_form As String, ByRef a_cr_type() As String, ByVal cr_type As String, ByVal cr_type_short As String, ByVal cr_form_type As String, ByVal date_received As Date, ByVal requester As String, ByRef tech As String, ByRef province_short As String, ByRef team_short As String, ByRef approver As String, ByVal format As cr_sheet_format, ByVal local As local_machine, ByRef err As String)
        Try
            'I think we should just transparently pass attachments and not check names.  It is up to the next guy to reject if they need attachments
            'I'm also not going to write it to the CR sheet

            'opens a new instance of XL
            '----------------------------------------------
            Dim xlapp As New Excel.Application
            xlapp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow
            xlapp.Visible = format.debug_xl
            xlapp.DisplayAlerts = False
            xlapp.UserControl = False
            xlapp.IgnoreRemoteRequests = True
            xlapp.Interactive = False
            Dim xlbook As Excel.Workbook = Nothing
            Dim cr_path As String = ""
            Dim inbox_path As String = local.base_path & local.inbox
            Dim cr_form_to_open As String = ""
            Dim col_index As Integer = 0

            Try
                'sets the resubmit flag, this determines how the sub will act whether it is to process a new cr_re_form or a resubmit form
                'there are subtle differences
                'NOTE: all pending attachments status CRs (new CR or resub CR) will be resub format as we know the cr_id already.  Only the completely new first time CRs have no cr_id.  Resub only (not pending) will have a value in cr_form_temp whereas the pending cases will be blank as we will be using the cr_form in the cr dir
                '----------------------------------
                Dim resubmit As Boolean = False
                If Not cr_id Like "" Then
                    resubmit = True
                    cr_path = Path.GetDirectoryName(cr_form)
                End If

                'fills the common/data_hdr_name arrays and the common/data_restriction_arrays to use
                '------------------------------------------------------------------------------
                Dim com_hdr_name() As String = {}
                Dim com_restrict() As String = {}
                com_hdr_name = format.common_hdr_name
                com_restrict = format.common_data_restriction

                Dim det_hdr_name() As String = {}
                Dim det_restrict() As String = {}
                If resubmit Then
                    det_hdr_name = get_string_array_from_name("detail_hdr_name_" & cr_form_type, format, err)
                    If Not err Like "" Then GoTo get_out
                    det_restrict = get_string_array_from_name("detail_data_restriction_" & cr_form_type & "_resubmit", format, err)
                    If Not err Like "" Then GoTo get_out
                Else
                    det_hdr_name = get_string_array_from_name("detail_hdr_name", format, err)
                    If Not err Like "" Then GoTo get_out
                    det_restrict = get_string_array_from_name("detail_data_restriction_initial", format, err)
                    If Not err Like "" Then GoTo get_out
                End If

                'converts the col name from the xl sheet name to the DB name
                'need this as the datatable cols are as per the DB cols
                '-------------------------------------------------------------
                Dim com_hdr_name_db() As String = com_hdr_name.ToArray
                Dim det_hdr_name_db() As String = det_hdr_name.ToArray
                Try
                    For i = 0 To com_hdr_name_db.Count - 1
                        com_hdr_name_db(i) = Regex.Replace(Strings.LCase(com_hdr_name_db(i)), "\s", "_")
                    Next
                    Dim colconvert_name() As String = {"cur", "pro", "act", "fin"}
                    Dim colconvert_cnt As Integer = 0
                    For i = 0 To det_hdr_name_db.Count - 1
                        det_hdr_name_db(i) = Regex.Replace(Strings.LCase(det_hdr_name_db(i)), "\s", "_")
                        If Regex.IsMatch(det_hdr_name_db(i), "^(az)|(mdt)|(edt)$") Then
                            det_hdr_name_db(i) = colconvert_name(colconvert_cnt) & "_" & det_hdr_name_db(i)
                            If Regex.IsMatch(det_hdr_name_db(i), "edt$") Then
                                colconvert_cnt = colconvert_cnt + 1
                            End If
                        ElseIf Regex.IsMatch(det_hdr_name_db(i), "^(ht)|(antenna)|(coax_len)$") Then
                            det_hdr_name_db(i) = "fin_" & det_hdr_name_db(i)
                        End If
                    Next
                Catch ex As Exception
                End Try

                'this opens the xl file
                '-------------------
                Try
                    Debug.WriteLine(Now.ToLongTimeString & ": opening file")

                    cr_form_to_open = cr_form_temp
                    xlapp.DisplayAlerts = False
                    xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                Catch ex As COMException
                    If ex.HResult = -2146827284 Then
                        fix_bad_xlsb_file(cr_form_to_open, xlapp, format, local, err)
                        If Not err Like "" Then GoTo get_out
                        Try
                            xlapp.DisplayAlerts = False
                            xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                        Catch exx As Exception
                            err = "ER: Internal error opening xl file after fixing, details: " & exx.ToString
                            GoTo get_out
                        End Try
                    Else
                        err = "ER: Internal COM error opening xl file, details: " & ex.ToString
                        GoTo get_out
                    End If
                Catch ex As Exception
                    err = "ER: Internal error opening xl file, details: " & ex.ToString
                    GoTo get_out
                End Try

                'this checks the CR sheet exists
                '--------------------------
                Try
                    Dim xlsheet_temp As Excel.Worksheet = xlbook.Worksheets("CR")
                Catch ex As Exception
                    If resubmit Then err = "RESUBREJ: " Else err = "CRREJ: "
                    err = err & "Form Rejection.  The CR form (" & cr_form_to_open & ") doesn't have a sheet called 'CR'.<BR>Thanks"
                    GoTo get_out
                End Try
                Dim xlsheet As Excel.Worksheet = xlbook.Worksheets("CR")
                xlsheet.Activate()

                'unprotects book and sheet and unhides cells
                '--------------------------------------------
                Debug.WriteLine(Now.ToLongTimeString & ": unhiding/unprotecting cells")

                Try
                    xlsheet.Unprotect(format.x_factor)
                    xlsheet.Range("A1").Value2 = "z"        'this disables the worksheet change macro or reset macro
                    unhide_and_lock_all(format.detail_hdr_row_start, det_hdr_name.Count, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out
                Catch ex As Exception
                    err = "ER: some error unprotecting or unhiding and locking cells, form check sub"
                    GoTo get_out
                End Try

                'this clear cells after the last detailed col ("executor comments") - needed to  deal with older forms
                '-------------------------------------------------------------------------------
                Try
                    Dim r1 As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + det_hdr_name.Count), xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + det_hdr_name.Count + 100))
                    r1.EntireColumn.Clear()
                Catch ex As Exception
                End Try

                'This checks all the header cells are present with the correct headers
                '-----------------------------------------------------------------------
                Debug.WriteLine(Now.ToLongTimeString & ": checking headers")

                'This checks all the header cells are present with the correct headers
                '-----------------------------------------------------------------------
                For col_index = 0 To com_hdr_name.Count - 1
                    Dim xlcell As Excel.Range = xlsheet.Range(xlsheet.Cells(format.common_row_start + col_index, format.common_hdr_col), xlsheet.Cells(format.common_row_start + col_index, format.common_hdr_col))
                    If Not o2s(xlcell.Value2) Like com_hdr_name(col_index) Then
                        '###########################################
                        'this converts the old vals to the new vals for pro and con
                        If o2s(xlcell.Value2) Like "Pros" And com_hdr_name(col_index) Like "Expected Impact" Then : xlcell.Value2 = "Expected Impact"
                        ElseIf o2s(xlcell.Value2) = "Cons" And com_hdr_name(col_index) Like "Risk" Then : xlcell.Value2 = "Risk"
                            '###########################################
                        Else
                            If resubmit Then err = "RESUBREJ: " Else err = "CRREJ: "
                            err = err & "CR form doesn't have the required headers => " & com_hdr_name(col_index)
                            GoTo get_out
                        End If
                    End If
                Next

                For col_index = 0 To det_hdr_name.Count - 1
                    Dim xlcell As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + col_index), xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + col_index))
                    If Not o2s(xlcell.Value2) Like det_hdr_name(col_index) Then
                        If resubmit Then err = "RESUBREJ: " Else err = "CRREJ: "
                        err = err & "CR form doesn't have the required headers => " & det_hdr_name(col_index)
                        GoTo get_out
                    End If
                Next

                'This checks the common values meet requirements
                '-----------------------------------------------
                Debug.WriteLine(Now.ToLongTimeString & ": checking common vals")

                err = ""
                Try
                    tech = ""
                    For col_index = 0 To com_hdr_name.Count - 1
                        Dim xlcell As Excel.Range = xlsheet.Range(xlsheet.Cells(format.common_row_start + col_index, format.common_data_col), xlsheet.Cells(format.common_row_start + col_index, format.common_data_col))
                        Dim test_string = Trim(o2s(xlcell.Value2))
                        Dim raw_restrict As String = com_restrict(col_index)
                        Dim data_restrict() As String = Split(raw_restrict, ",")

                        If raw_restrict Like "not blank" Then
                            If test_string Like "" Then
                                err = "ERROR!! - Must be filled"
                                error_format_cell(True, xlcell, err & " - " & test_string)
                            End If

                        ElseIf Regex.IsMatch(raw_restrict, "^tablematch,tech,tech$") Then     'we need to check the data against the given datatable
                            tech = test_string
                            If Not is_allowed_val(test_string, data_restrict(1), data_restrict(2), format) Then
                                tech = "bad technology"
                                err = "ERROR!! - Not allowed value"
                                error_format_cell(True, xlcell, err & " - " & test_string)
                            End If

                        ElseIf Regex.IsMatch(raw_restrict, "^tablematch,depends\son\stech,cr_objective$", RegexOptions.IgnoreCase) Then     'we need to check the data against the given datatable
                            'sets the table name for the cr_obj based on the tech
                            If Not tech Like "" And Not tech Like "bad technology" Then
                                If Not is_allowed_val(test_string, "cr_obj_" & LCase(tech), data_restrict(2), format) Then
                                    error_format_cell(True, xlcell, test_string)
                                End If
                            End If

                        ElseIf Regex.IsMatch(raw_restrict, "^tablematch,") Then     'we need to check the data against the given datatable
                            If Not is_allowed_val(test_string, data_restrict(1), data_restrict(2), format) Then
                                err = "ERROR!! - Not allowed value"
                                error_format_cell(True, xlcell, err & " - " & test_string)
                            End If
                        End If

                        'This reads some important header info for tool use
                        '----------------------------------------------------
                        If Regex.IsMatch(com_hdr_name(col_index), "^Approver$", RegexOptions.IgnoreCase) Then
                            approver = c2e(test_string)
                            If Not format.IsValidEmail(approver) Then
                                err = "ERROR!! - Must contain a valid email address"
                                error_format_cell(True, xlcell, err & " - " & test_string)
                            End If

                        ElseIf Regex.IsMatch(com_hdr_name(col_index), "^Team$", RegexOptions.IgnoreCase) Then
                            team_short = ""
                            Dim qrows = From row As System.Data.DataRow In format.ds_allow.Tables("teams")
                                        Where row("team") Like test_string
                                        Select row
                            If qrows.Count > 0 Then
                                team_short = qrows(0)("team_short")
                            End If

                        ElseIf Regex.IsMatch(com_hdr_name(col_index), "^region$", RegexOptions.IgnoreCase) Then
                            province_short = ""
                            Dim qrows = From row As System.Data.DataRow In format.ds_allow.Tables("geo")
                                        Where row("region") Like test_string
                                        Select row
                            If qrows.Count > 0 Then
                                province_short = qrows(0)("province_short")
                            End If

                        End If
                    Next
                Catch ex As Exception
                    err = "ER: General error checking common header vals, details: " & ex.ToString
                    GoTo get_out
                End Try
                If Not err Like "" Then
                    data_ok = "nok"
                    err = ""
                End If

                'Then we check the detail values col by col
                '-----------------------------------------------
                Dim index_crs As Integer = Array.IndexOf(det_hdr_name, "CR_sub_ID")
                Dim index_cr_type As Integer = Array.IndexOf(det_hdr_name, "CR Type")
                Dim index_node_type As Integer = Array.IndexOf(det_hdr_name, "Node Type")
                Dim index_node As Integer = Array.IndexOf(det_hdr_name, "Node")
                Dim index_propval As Integer = Array.IndexOf(det_hdr_name, "Proposed Setting")
                Dim index_comment As Integer = Array.IndexOf(det_hdr_name, "Requester Comments")
                Dim index_ex_coord As Integer = Array.IndexOf(det_hdr_name, "Execution Coordinator")
                Dim index_attach As Integer = 0
                If resubmit And cr_type_short Like "prm" Then
                    index_attach = -1
                Else
                    index_attach = Array.IndexOf(det_hdr_name, "Requester Attachments")
                End If

                'First we find the row limit for the cr form
                '----------------------------------------------
                err = ""
                Dim last_data_row As Integer = 0
                Try
                    With xlsheet
                        last_data_row = .Columns("A:A").offset(0, index_cr_type).entirecolumn.Find(What:="*", After:=.Cells(index_cr_type + 1), LookAt:=XlLookAt.xlWhole, LookIn:=XlFindLookIn.xlValues, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False).Row
                    End With
                Catch ex As Exception
                    last_data_row = 0
                End Try
                If last_data_row = 0 Then
                    err = "XREJ: There is some issue with you CR form, the format is messed up, can't find the last data row, be careful not to change the format."
                    GoTo get_out
                ElseIf last_data_row - format.detail_data_row_start + 1 > format.sheet_row_limit_global Then
                    err = "XREJ: Currently, a max of " & Math.Round(format.sheet_row_limit_global / 1000, 1) & "K rows can be processed in the new CR form, please split your CR."
                    GoTo get_out
                ElseIf resubmit Then
                    If Regex.IsMatch(cr_type, "^((RF\sRe-engineering)|(RF Basic))$", RegexOptions.IgnoreCase) AndAlso last_data_row - format.detail_data_row_start + 1 > format.sheet_row_limit_rfb Then
                        err = "XREJ: Currently, a max of " & Math.Round(format.sheet_row_limit_rfb / 1000, 1) & "K rows is supported for CR's of the type 'RF Re-engineering' or 'RF Basic', please split your CR."
                        GoTo get_out
                    End If
                End If


                '########################################################################################
                '########################################################################################
                '########################################################################################
                'at this point we would read the detailed data into a DS and process it there
                'load the form into the dataset
                '--------------------------------
                Debug.WriteLine(Now.ToLongTimeString & ": loading form to memory for checking")

                Dim ds_test As New System.Data.DataSet
                ds_test.Tables.Add("com")
                ds_test.Tables.Add("det")
                load_xl2ds(com_hdr_name.Count, last_data_row, ds_test, xlsheet, "", format, err)
                If Not err Like "" Then GoTo get_out
                'at this point ds_test.tables("com") and ds_test.tables("det") have the sheet data loaded with the database headers
                'so from here on you have to work in ds_temp with the database col names for the checking, if it checks out then you just throw ds_temp, 
                '########################################################################################
                '########################################################################################
                '########################################################################################

                Dim cr_type_range As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start - 1, index_cr_type + 1), xlsheet.Cells(last_data_row, index_cr_type + 1))
                Dim node_type_range As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start - 1, index_node_type + 1), xlsheet.Cells(last_data_row, index_node_type + 1))

                'check that there is only one execution coordinator per cr_type, reject if more than 1, they need to put in a separate CR
                '-----------------------------------------------------------------------------------------------------------------------------
                Debug.WriteLine(Now.ToLongTimeString & ": 1 ex_coord per cr type")

                err = ""
                Try
                    For Each test_cr_type In {"Parameter", "Hardware", "RF Basic", "RF Re-engineering"}
                        Dim qrows = From row In ds_test.Tables("det")
                                    Where row.Field(Of String)("cr_type") Like test_cr_type
                                    Select row
                        Dim test As String = ""
                        For Each row In qrows
                            If test = "" Then
                                test = row.Field(Of String)("execution_coordinator")
                            ElseIf Not row.Field(Of String)("execution_coordinator") = test Then
                                err = "ERROR!! - You can only have one execution coordinator per cr type - "
                                row("execution_coordinator") = err & row.Field(Of String)("execution_coordinator")
                            End If
                        Next
                    Next
                Catch ex As Exception
                    err = "ER: error checking ex_coord vals, details: " & ex.ToString
                    GoTo get_out
                End Try
                If Not err Like "" Then
                    data_ok = "nok"
                    err = ""
                End If

                'only for the resubmit case, we check that there are no new changes of different types to the cr_form_type.
                '--------------------------------------------------------------------------------------------------
                err = ""
                Try
                    If resubmit Then
                        Dim qrows = From row In ds_test.Tables("det")
                                    Where Not row.Field(Of String)("cr_type") Like cr_type
                                    Select row
                        If qrows.Count > 0 Then
                            err = "ERROR!! - You can't change the cr type, it must be: " & cr_type & " - "
                            For Each row In qrows
                                row("cr_type") = err & row.Field(Of String)("cr_type")
                            Next
                        End If
                    End If
                Catch ex As Exception
                    err = "ER: error checking ex_coord vals, details: " & ex.ToString
                    GoTo get_out
                End Try
                If Not err Like "" Then
                    data_ok = "nok"
                    err = ""
                End If

                'then we go through the header cols 1 by 1
                '-----------------------------------------
                err = ""
                For col_index = 0 To det_hdr_name_db.Count - 1
                    'first set the search range for each header col
                    '--------------------------------------------------
                    Dim raw_restrict As String = det_restrict(col_index)
                    Dim data_restrict() As String = Split(raw_restrict, ",")
                    Dim col As String = det_hdr_name_db(col_index)
                    Debug.WriteLine(Now.ToLongTimeString & ": checking detailed col: " & col)

                    'check values comply
                    '-----------------------------------
                    If raw_restrict Like "not blank" Then
                        Try
                            Dim qrows = From row In ds_test.Tables("det")
                                        Where row.Field(Of String)(col) Like ""
                                        Select row
                            If qrows.Count > 0 Then
                                err = "ERROR!! - Must be filled"
                                For Each row In qrows
                                    row(col) = err & " - " & row.Field(Of String)(col)
                                Next
                            End If
                        Catch ex As Exception
                            err = "ER: General error checking values (not blank), details: " & ex.ToString
                            GoTo get_out
                        End Try

                    ElseIf Regex.IsMatch(raw_restrict, "^tablematch,", RegexOptions.IgnoreCase) Then     'we need to check the data against the given datatable
                        '##########################################################
                        Try
                            If Regex.IsMatch(raw_restrict, "^tablematch,depends\son\scr_type,combined_name$", RegexOptions.IgnoreCase) Then     'we need to check the data against the given datatable
                                'sets the table name for the execution coordinator based on the cr_type
                                For Each test_cr_type In {"Parameter", "Hardware", "RF Basic", "RF Re-engineering"}
                                    Dim qrows0 = From row In format.ds_allow.Tables("cr_types")
                                                Where row.Field(Of String)("cr_type") Like test_cr_type
                                                Select row
                                    Dim qrows1 = From row In ds_test.Tables("det")
                                                    Where row.Field(Of String)("cr_type") Like test_cr_type
                                                    Select row
                                    Dim t_table As String = qrows0.First.Field(Of String)("cr_type_short")
                                    Dim test_vals() As String = (From row In format.ds_allow.Tables(t_table & "_ex_coord")
                                                                Let a = row.Field(Of String)(data_restrict(2))
                                                                Select a).Distinct.ToArray
                                    Dim qrows2 = From row In qrows1.AsEnumerable
                                                 Where Not test_vals.Any(Function(s) row.Field(Of String)(col).Contains(s))
                                                 Select row
                                    If qrows2.Count > 0 Then
                                        err = "ERROR!! - Not allowed value"
                                        For Each row In qrows2
                                            row(col) = err & " - " & row.Field(Of String)(col)
                                        Next
                                    End If
                                Next

                            Else
                                Dim test_vals() As String = (From row In format.ds_allow.Tables(data_restrict(1))
                                                            Let a = row.Field(Of String)(data_restrict(2))
                                                            Select a).Distinct.ToArray
                                Dim qrows = From row In ds_test.Tables("det")
                                            Where Not test_vals.Any(Function(s) row.Field(Of String)(col).Contains(s))
                                            Select row
                                If qrows.Count > 0 Then
                                    err = "ERROR!! - Not allowed value"
                                    For Each row In qrows
                                        row(col) = err & " - " & row.Field(Of String)(col)
                                    Next
                                End If
                            End If
                        Catch ex As Exception
                            err = "ER: General error checking values (tablematch), details: " & ex.ToString
                            GoTo get_out
                        End Try

                    ElseIf Regex.IsMatch(raw_restrict, "^cr\stype\sis\sparameter\sand\snot\sblank$", RegexOptions.IgnoreCase) Then
                        '##########################################################
                        Try
                            Dim qrows = From row In ds_test.Tables("det")
                                        Where Regex.IsMatch(row.Field(Of String)("cr_type"), "^Parameter$", RegexOptions.IgnoreCase) AndAlso row.Field(Of String)(col) Like ""
                                        Select row
                            If qrows.Count > 0 Then
                                err = "ERROR!! - Must be filled"
                                For Each row In qrows
                                    row(col) = err & " - " & row.Field(Of String)(col)
                                Next
                            End If
                        Catch ex As Exception
                            err = "ER: General error checking values (not blank), details: " & ex.ToString
                            GoTo get_out
                        End Try

                    ElseIf Regex.IsMatch(raw_restrict, "^cr\stype\sis\srf\sbasic\s") And Regex.IsMatch(raw_restrict, "\sand\s2val:\s") Then
                        '##########################################################
                        Try
                            Dim t_restrict As String = Regex.Replace(raw_restrict, "^.*val:\s", "", RegexOptions.IgnoreCase)
                            Dim temp_restrict() As Integer = (From item In Split(t_restrict, ",") Let a As Integer = Val(Trim(item)) Select a).ToArray
                            Dim qrows = From row In ds_test.Tables("det")
                                        Let a = Strings.Split(row.Field(Of String)(col)), b As String = Regex.Replace(a(0), "[^0-9]", ""), c As String = If(a.Count = 2, Regex.Replace(a(1), "[^0-9]", ""), Nothing), bi As Integer = Val(b), ci As Integer = If(Not c Is Nothing, Val(c), Nothing)
                                        Where Regex.IsMatch(row.Field(Of String)("cr_type"), "^RF\sBasic$", RegexOptions.IgnoreCase) _
                                        AndAlso (b Like "" Or bi < temp_restrict(0) Or bi > temp_restrict(1) Or If(Not c Is Nothing, c Like "" Or ci < temp_restrict(0) Or ci > temp_restrict(1), Nothing))
                                        Select row
                            If qrows.Count > 0 Then
                                err = "ERROR!! - Outside of allowed range"
                                For Each row In qrows
                                    row(col) = err & " - " & row.Field(Of String)(col)
                                Next
                            End If

                        Catch ex As Exception
                            err = "ER: General error checking values, details: " & ex.ToString
                            GoTo get_out
                        End Try

                    ElseIf Regex.IsMatch(raw_restrict, "^cr\stype\sis\srf\sbasic\s") And Regex.IsMatch(raw_restrict, "\sand\sval:\s") Then
                        '##########################################################
                        Try
                            Dim t_restrict As String = Regex.Replace(raw_restrict, "^.*val:\s", "", RegexOptions.IgnoreCase)
                            Dim temp_restrict() As Integer = (From item In Split(t_restrict, ",") Let a As Integer = Val(Trim(item)) Select a).ToArray
                            Dim qrows = From row In ds_test.Tables("det")
                                        Let a As String = Regex.Replace(row.Field(Of String)(col), "[^0-9]", ""), ai As Integer = Val(a)
                                        Where Regex.IsMatch(row.Field(Of String)("cr_type"), "^RF\sBasic$", RegexOptions.IgnoreCase) _
                                        AndAlso (a Like "" Or ai < temp_restrict(0) Or ai > temp_restrict(1))
                                        Select row
                            If qrows.Count > 0 Then
                                err = "ERROR!! - Outside of allowed range"
                                For Each row In qrows
                                    row(col) = err & " - " & row.Field(Of String)(col)
                                Next
                            End If
                        Catch ex As Exception
                            err = "ER: General error checking values, details: " & ex.ToString
                            GoTo get_out
                        End Try
                    End If
                Next
                If Not err Like "" Then
                    data_ok = "nok"
                    err = ""
                End If
                If data_ok Like "not finished testing" Then data_ok = "ok"

                'get the cr_types in an array, need this to get the cr_ids from the DB
                '------------------------------------------------------------------------
                If Not resubmit Then
                    Try
                        a_cr_type = {}
                        a_cr_type = (From row In ds_test.Tables("det")
                                     Let a = row.Field(Of String)("cr_type")
                                     Select a).Distinct.ToArray
                        If a_cr_type.Count = 0 Then
                            err = "ER: There are no cr_types found, internal error"
                            GoTo get_out
                        End If
                    Catch ex As Exception
                        err = "ER: General error getting cr_types, details: " & ex.ToString
                        GoTo get_out
                    End Try
                End If

                '#######################################################################################
                '#######################################################################################
                'if we got a format error exit now
                '-----------------------------------
                If data_ok = "nok" Then
                    Debug.WriteLine(Now.ToLongTimeString & ": data is nok so cleaning up")

                    'writes the details table with errors data back to the xlsheet for the error file
                    '-----------------------------------------------------------------------
                    err = ""
                    dt2xlrange(last_data_row, ds_test.Tables("det"), xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out
                    'we only do the error formatting on small to med CRs as it involves filtering which gets sketchy on large ranges, L will throw and behave badly
                    'for big crs we just write the error data 
                    set_all_errors_cell2red_unlock(last_data_row, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out

                    'hides fields again to the original cr request format, they are all protected, this is ok
                    '-------------------------------------------------------------------------------
                    If resubmit Then
                        'the case of a resub that has bad data
                        Dim col_array() As Integer = get_integer_array_from_name("col_clear_hide_" & cr_form_type & "_for_resub", format, err)
                        If Not err Like "" Then GoTo get_out
                        hide_rows_and_cols(format.row_hide_for_bad_resub, col_array, xlsheet, format, err)
                        If Not err Like "" Then GoTo get_out
                    Else
                        'the case of a new form that has not been split yet with bad data or format
                        hide_rows_and_cols(format.row_hide_for_new_fail, format.col_hide_for_new_fail, xlsheet, format, err)
                        If Not err Like "" Then GoTo get_out
                    End If

                    'protects the sheet and book
                    '------------------------------------------------------------------------
                    xlsheet.Range("C2").Select()
                    xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True)

                    'saves the error file
                    '---------------------------
                    Dim file_out As String = ""
                    file_out = inbox_path & "\" & Path.GetFileNameWithoutExtension(cr_form_to_open) & "_errors.xlsb"
                    xlsheet.Range("C2").Select()
                    xlapp.DisplayAlerts = False
                    xlbook.SaveAs(file_out)
                    If resubmit Then
                        err = "XREJ: See returned CR form (CRMS: CR Resubmission Retry Request (" & cr_id & "))(" & file_out & ")"
                    Else
                        err = "XREJ: See returned CR form (CRMS: New CR Retry Request)(" & file_out & ")"
                    End If
                    GoTo get_out
                End If
                '#######################################################################################
                '#######################################################################################


                'does some actions specific to each case now
                Debug.WriteLine(Now.ToLongTimeString & ": doing specific actions prior to split")

                If resubmit Then
                    'this is only done for resubmits
                    'we need to prepare the file for the approver before exiting
                    '-----------------------------------------------
                    'clears all cells after the last data row
                    '---------------------------------------------------------------------
                    Try
                        Dim clr_range As Excel.Range = xlsheet.Range(xlsheet.Cells(last_data_row + 1, 1), xlsheet.Cells(1048576, 1))
                        clr_range.EntireRow.Clear()
                    Catch ex As Exception
                    End Try

                    'fill out the sub_ids in the detailed table => in memory and the XL sheet
                    '----------------------------------------------------------------------------------------------
                    Try
                        Dim i As Integer = 0
                        For Each row As System.Data.DataRow In ds_test.Tables("det").Rows
                            i += 1
                            row("cr_sub_id") = cr_id & "." & i
                        Next
                        dt2xlrange(last_data_row, ds_test.Tables("det"), xlsheet, format, err)
                        If Not err Like "" Then GoTo get_out
                    Catch ex As Exception
                        err = "RESUBREJ: Couldn't fill cr_ids in the CR form...."
                        GoTo get_out
                    End Try

                    'this fills some common header fields on the sheet and the com dt
                    '---------------------------------------------------------------
                    Try
                        For col_index = 0 To com_hdr_name.Count - 1
                            Dim xlcell As Excel.Range = xlsheet.Range(xlsheet.Cells(format.common_row_start + col_index, format.common_data_col), xlsheet.Cells(format.common_row_start + col_index, format.common_data_col))
                            Dim test_string = Trim(o2s(xlcell.Value2))
                            If Regex.IsMatch(com_hdr_name(col_index), "^CR_ID$", RegexOptions.IgnoreCase) Then
                                xlcell.Value2 = cr_id
                                ds_test.Tables("com").Rows(0)("cr_id") = o2s(xlcell.Value2)

                            ElseIf Regex.IsMatch(com_hdr_name(col_index), "^Requester$", RegexOptions.IgnoreCase) Then
                                Dim qrows() As String = (From row In format.ds_allow.Tables("requesters") Let a = row.Field(Of String)("combined_name") Where row.Field(Of String)("email") Like requester Select a).ToArray
                                xlcell.Value2 = qrows.First
                                ds_test.Tables("com").Rows(0)("requester") = o2s(xlcell.Value2)

                            ElseIf Regex.IsMatch(com_hdr_name(col_index), "^Execution\sCoordinator$", RegexOptions.IgnoreCase) Then
                                Dim qrows() As String = (From row In ds_test.Tables("det") Let a = row.Field(Of String)("execution_coordinator") Select a).Distinct.ToArray
                                xlcell.Value2 = qrows.First
                                ds_test.Tables("com").Rows(0)("execution_coordinator") = o2s(xlcell.Value2)

                            ElseIf Regex.IsMatch(com_hdr_name(col_index), "^CR\sType$", RegexOptions.IgnoreCase) Then
                                xlcell.Value2 = cr_type    'o2s(xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_cr_type + 1), xlsheet.Cells(format.detail_data_row_start, index_cr_type + 1)).Value2)
                                ds_test.Tables("com").Rows(0)("cr_type") = o2s(xlcell.Value2)

                            ElseIf Regex.IsMatch(com_hdr_name(col_index), "^Open\sDate$", RegexOptions.IgnoreCase) Then
                                date_format_cell(xlcell)
                                xlcell.Value2 = Now.ToOADate
                                ds_test.Tables("com").Rows(0)("open_date") = o2s(xlcell.Value2)

                            ElseIf Regex.IsMatch(com_hdr_name(col_index), "^Node\sTypes$", RegexOptions.IgnoreCase) Then
                                Dim qrows() As String = (From row In ds_test.Tables("det") Let a = row.Field(Of String)("node_type") Select a).Distinct.ToArray
                                xlcell.Value2 = Join(qrows, ", ")
                                ds_test.Tables("com").Rows(0)("node_types") = o2s(xlcell.Value2)

                            End If
                        Next
                    Catch ex As Exception
                        err = "RESUBREJ: Couldn't add common header vals to the CR form...."
                        GoTo get_out
                    End Try

                    'sets the hide and unlock config for the next stage (approver) => there is no unlocking for the next stage
                    '---------------------------------------------------------------------------------------------
                    Dim col_array() As Integer = get_integer_array_from_name("col_hide_" & cr_form_type & "_for_app", format, err)
                    If Not err Like "" Then GoTo get_out
                    hide_rows_and_cols(format.row_hide_for_app, col_array, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out

                    Debug.WriteLine(Now.ToLongTimeString & ": copying test DTs to output DTs")

                    'copy the test tables to the output tables
                    '------------------------------------------------
                    If Not ds.Tables("com") Is Nothing Then ds.Tables.Remove("com")
                    Dim dt_com As New System.Data.DataTable
                    dt_com = ds_test.Tables("com").Copy
                    dt_com.TableName = "com"
                    ds.Tables.Add(dt_com)
                    dt_com = ds.Tables("com")

                    If Not ds.Tables("det") Is Nothing Then ds.Tables.Remove("det")
                    Dim dt_det As New System.Data.DataTable
                    dt_det = ds_test.Tables("det").Copy
                    dt_det.TableName = "det"
                    ds.Tables.Add(dt_det)
                    dt_det = ds.Tables("det")

                ElseIf Not resubmit Then
                    'remove the reset button
                    '------------------------
                    For Each item As Excel.Shape In xlsheet.Shapes
                        If item.Name = "CommandButton1" Then
                            item.Delete()
                            Exit For
                        End If
                    Next

                    Debug.WriteLine(Now.ToLongTimeString & ": copying test DTs to output DTs")

                    'copy the test tables to the init tables
                    '------------------------------------------------
                    If Not ds.Tables("init_com") Is Nothing Then ds.Tables.Remove("init_com")
                    Dim dt_com As New System.Data.DataTable
                    dt_com = ds_test.Tables("com").Copy
                    dt_com.TableName = "init_com"
                    ds.Tables.Add(dt_com)
                    dt_com = ds.Tables("init_com")

                    If Not ds.Tables("init_data") Is Nothing Then ds.Tables.Remove("init_data")
                    Dim dt_det As New System.Data.DataTable
                    dt_det = ds_test.Tables("det").Copy
                    dt_det.TableName = "init_data"
                    ds.Tables.Add(dt_det)
                    dt_det = ds.Tables("init_data")
                End If

                Debug.WriteLine(Now.ToLongTimeString & ": final formatting after checking initial form")

                'resets the formatting of the common and detail cells on the sheet, do not reset format for the requester attachments col (last one)
                '---------------------------------------------------------------------------------------------------------------
                Dim t_rng As Excel.Range = xlsheet.Range(xlsheet.Cells(format.common_row_start, format.common_data_col), xlsheet.Cells(format.common_row_start + com_hdr_name.Count - 1, format.common_data_col))
                normal_format_cell_resub(t_rng)
                t_rng = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, 1), xlsheet.Cells(last_data_row, det_hdr_name.Count - 1))
                normal_format_cell_resub(t_rng)

                'protect and save the cr form
                '-----------------------------
                xlsheet.Range("C2").Select()
                xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True)
                xlapp.DisplayAlerts = False
                xlbook.SaveAs(cr_form_to_open)

get_out:
            Catch ex As Exception
                err = "ER: General error checking the cr file '" & cr_form_to_open & "'.  Details: " & ex.ToString
            Finally
                Try
                    xlapp.DisplayAlerts = False
                    xlbook.Close()
                Finally
                    releaseObject(xlbook)
                    xlapp.UserControl = True
                    xlapp.Interactive = True
                    xlapp.IgnoreRemoteRequests = False
                    xlapp.Quit()
                    releaseObject(xlapp) 'this releases the com object
                End Try
            End Try
        Catch ex As Exception
            err = "ER: Error opening the XL application, details: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub








    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################


    'this splits the incoming CR form into the physical and parameter components
    '-------------------------------------------------------------------------
    Public Sub split_cr_request_form(ByVal file_in As String, ByRef file_out As String, ByRef ds As System.Data.DataSet, ByVal cr_id As String, ByVal cr_type As String, ByVal cr_type_short As String, ByVal cr_form_type As String, ByVal cr_path As String, ByVal requester As String, ByVal tech As String, ByVal format As cr_sheet_format, ByVal local As local_machine, ByRef err As String)
        Try
            'opens a new instance of XL
            '----------------------------------------------
            Dim xlapp As New Excel.Application
            xlapp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow
            xlapp.Visible = format.debug_xl
            xlapp.DisplayAlerts = False
            xlapp.UserControl = False
            xlapp.IgnoreRemoteRequests = True
            xlapp.Interactive = False
            Dim xlbook As Excel.Workbook = Nothing
            Dim temp_range As Excel.Range

            Try
                Dim com_hdr_name() As String = {}
                com_hdr_name = format.common_hdr_name
                Dim det_hdr_name() As String = {}
                det_hdr_name = get_string_array_from_name("detail_hdr_name_" & cr_form_type, format, err)
                If Not err Like "" Then GoTo get_out

                'Open the XL file
                '------------------------
                Dim cr_form_to_open As String = file_in
                Try
                    Debug.WriteLine(Now.ToLongTimeString & ": opening xl file for split: " & cr_type)
                    xlapp.DisplayAlerts = False
                    xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                Catch ex As COMException
                    If ex.HResult = -2146827284 Then
                        fix_bad_xlsb_file(cr_form_to_open, xlapp, format, local, err)
                        If Not err Like "" Then GoTo get_out
                        Try
                            xlapp.DisplayAlerts = False
                            xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                        Catch exx As Exception
                            err = "ER: Internal error opening xl file after fixing, details: " & exx.ToString
                            GoTo get_out
                        End Try
                    Else
                        err = "ER: Internal COM error opening xl file, details: " & ex.ToString
                        GoTo get_out
                    End If
                Catch ex As Exception
                    err = "ER: Internal error opening xl file, details: " & ex.ToString
                    GoTo get_out
                End Try

                Dim xlsheet As Excel.Worksheet = xlbook.Worksheets("CR")
                xlsheet.Activate()

                'unprotects book and sheet and unhides cells
                '--------------------------------------------
                Try
                    Debug.WriteLine(Now.ToLongTimeString & ": unprotecting book")

                    xlsheet.Unprotect(format.x_factor)
                    xlsheet.Range("A1").Value2 = "z"        'this disables the worksheet change macro or reset macro
                    unhide_and_lock_all(format.detail_hdr_row_start, format.detail_hdr_name.Count, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out
                Catch ex As Exception
                    err = "ER: some error unprotecting or unhiding and locking cells, form check sub"
                    GoTo get_out
                End Try

                Debug.WriteLine(Now.ToLongTimeString & ": clearing cells")

                'this clear cells after the last detailed col ("executor comments") - needed to  deal with older forms
                '-------------------------------------------------------------------------------
                temp_range = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + format.detail_hdr_name.Count), xlsheet.Cells(format.detail_hdr_row_start + 2, format.detail_col_start + format.detail_hdr_name.Count + 100))
                temp_range.Clear()

                'remove all details contents as we will over write later in the sub from dts
                '-----------------------------------------------------------------------
                temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, 1), xlsheet.Cells(1048576, format.detail_hdr_name.Count))
                temp_range.ClearContents()

                Debug.WriteLine(Now.ToLongTimeString & ": creating final DTs")

                'copy the init tables to the final tables, removing unwanted row in the details table
                '-----------------------------------------------------------------------------------
                If Not ds.Tables(cr_type_short & "_com") Is Nothing Then ds.Tables.Remove(cr_type_short & "_com")
                Dim dt_com As New System.Data.DataTable
                dt_com = ds.Tables("init_com").Copy
                dt_com.TableName = cr_type_short & "_com"
                ds.Tables.Add(dt_com)
                dt_com = ds.Tables(cr_type_short & "_com")

                If Not ds.Tables(cr_type_short & "_data") Is Nothing Then ds.Tables.Remove(cr_type_short & "_data")
                Dim dt_det As System.Data.DataTable = (From row In ds.Tables("init_data") Where row.Field(Of String)("cr_type") Like cr_type Select row).CopyToDataTable
                dt_det.TableName = cr_type_short & "_data"
                ds.Tables.Add(dt_det)
                dt_det = ds.Tables(cr_type_short & "_data")

                'find the row limit for the new cr form and clear all cells after this
                '----------------------------------------------------------------
                Dim last_data_row As Integer = format.detail_data_row_start + dt_det.Rows.Count - 1
                If Regex.IsMatch(cr_type, "^((RF\sRe-engineering)|(RF Basic))$", RegexOptions.IgnoreCase) AndAlso dt_det.Rows.Count > format.sheet_row_limit_rfb Then
                    err = "XREJ: Currently, a max of " & Math.Round(format.sheet_row_limit_rfb / 1000, 1) & "K rows is supported for CR's of the type 'RF Re-engineering' or 'RF Basic', please split your CR."
                    GoTo get_out
                End If
                temp_range = xlsheet.Range(xlsheet.Cells(last_data_row + 1, 1), xlsheet.Cells(1048576, det_hdr_name.Count))
                temp_range.Clear()

                Debug.WriteLine(Now.ToLongTimeString & ": filling ids")

                'fill out the sub_ids in the detailed table => in memory and the XL sheet
                '----------------------------------------------------------------------------------------------
                Try
                    Dim i As Integer = 0
                    For Each row As System.Data.DataRow In dt_det.Rows
                        i += 1
                        row("cr_sub_id") = cr_id & "." & i
                    Next
                Catch ex As Exception
                    err = "RESUBREJ: Couldn't fill cr_ids in the CR form...."
                    GoTo get_out
                End Try

                Debug.WriteLine(Now.ToLongTimeString & ": filling common header")

                'this fills some common header fields on the sheet and the com dt
                '---------------------------------------------------------------
                Try
                    For col_index = 0 To com_hdr_name.Count - 1
                        Dim xlcell As Excel.Range = xlsheet.Range(xlsheet.Cells(format.common_row_start + col_index, format.common_data_col), xlsheet.Cells(format.common_row_start + col_index, format.common_data_col))
                        Dim test_string = Trim(o2s(xlcell.Value2))
                        If Regex.IsMatch(com_hdr_name(col_index), "^CR_ID$", RegexOptions.IgnoreCase) Then
                            xlcell.Value2 = cr_id
                            dt_com.Rows(0)("cr_id") = o2s(xlcell.Value2)

                        ElseIf Regex.IsMatch(com_hdr_name(col_index), "^Requester$", RegexOptions.IgnoreCase) Then
                            Dim qrows() As String = (From row In format.ds_allow.Tables("requesters") Let a = row.Field(Of String)("combined_name") Where row.Field(Of String)("email") Like requester Select a).ToArray
                            xlcell.Value2 = qrows.First
                            dt_com.Rows(0)("requester") = o2s(xlcell.Value2)

                        ElseIf Regex.IsMatch(com_hdr_name(col_index), "^Execution\sCoordinator$", RegexOptions.IgnoreCase) Then
                            Dim qrows() As String = (From row In dt_det Let a = row.Field(Of String)("execution_coordinator") Select a).Distinct.ToArray
                            xlcell.Value2 = qrows.First
                            dt_com.Rows(0)("execution_coordinator") = o2s(xlcell.Value2)

                        ElseIf Regex.IsMatch(com_hdr_name(col_index), "^CR\sType$", RegexOptions.IgnoreCase) Then
                            xlcell.Value2 = cr_type    'o2s(xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_cr_type + 1), xlsheet.Cells(format.detail_data_row_start, index_cr_type + 1)).Value2)
                            dt_com.Rows(0)("cr_type") = o2s(xlcell.Value2)

                        ElseIf Regex.IsMatch(com_hdr_name(col_index), "^Open\sDate$", RegexOptions.IgnoreCase) Then
                            date_format_cell(xlcell)
                            Dim xltime As Double = Now.ToOADate
                            xlcell.Value2 = xltime
                            dt_com.Rows(0)("open_date") = xltime.ToString

                        ElseIf Regex.IsMatch(com_hdr_name(col_index), "^Node\sTypes$", RegexOptions.IgnoreCase) Then
                            Dim qrows() As String = (From row In dt_det Let a = row.Field(Of String)("node_type") Select a).Distinct.ToArray
                            xlcell.Value2 = Join(qrows, ", ")
                            dt_com.Rows(0)("node_types") = o2s(xlcell.Value2)

                        End If
                    Next
                Catch ex As Exception
                    err = "RESUBREJ: Couldn't add common header vals to the CR form...."
                    GoTo get_out
                End Try

                Debug.WriteLine(Now.ToLongTimeString & ": deleting cols")

                'delete unwanted cols from the DT and sheet
                '-----------------------------------------------
                Try
                    If cr_form_type = "rfb" Then
                        xlsheet.Range("E:H").EntireColumn.Delete()
                        dt_det.Columns.Remove("nbr_node")
                        dt_det.Columns.Remove("parameter")
                        dt_det.Columns.Remove("proposed_setting")
                        dt_det.Columns.Remove("rollback_setting")

                    ElseIf cr_form_type = "prm" Then
                        xlsheet.Range("I:N").EntireColumn.Delete()
                        dt_det.Columns.Remove("cur_az")
                        dt_det.Columns.Remove("cur_mdt")
                        dt_det.Columns.Remove("cur_edt")
                        dt_det.Columns.Remove("pro_az")
                        dt_det.Columns.Remove("pro_mdt")
                        dt_det.Columns.Remove("pro_edt")
                        xlsheet.Range("M:U").EntireColumn.Delete()
                        dt_det.Columns.Remove("act_az")
                        dt_det.Columns.Remove("act_mdt")
                        dt_det.Columns.Remove("act_edt")
                        dt_det.Columns.Remove("fin_az")
                        dt_det.Columns.Remove("fin_mdt")
                        dt_det.Columns.Remove("fin_edt")
                        dt_det.Columns.Remove("fin_ht")
                        dt_det.Columns.Remove("fin_antenna")
                        dt_det.Columns.Remove("fin_coax_len")

                    Else
                        xlsheet.Range("E:N").EntireColumn.Delete()
                        dt_det.Columns.Remove("nbr_node")
                        dt_det.Columns.Remove("parameter")
                        dt_det.Columns.Remove("proposed_setting")
                        dt_det.Columns.Remove("rollback_setting")
                        dt_det.Columns.Remove("cur_az")
                        dt_det.Columns.Remove("cur_mdt")
                        dt_det.Columns.Remove("cur_edt")
                        dt_det.Columns.Remove("pro_az")
                        dt_det.Columns.Remove("pro_mdt")
                        dt_det.Columns.Remove("pro_edt")
                        xlsheet.Range("I:Q").EntireColumn.Delete()
                        dt_det.Columns.Remove("act_az")
                        dt_det.Columns.Remove("act_mdt")
                        dt_det.Columns.Remove("act_edt")
                        dt_det.Columns.Remove("fin_az")
                        dt_det.Columns.Remove("fin_mdt")
                        dt_det.Columns.Remove("fin_edt")
                        dt_det.Columns.Remove("fin_ht")
                        dt_det.Columns.Remove("fin_antenna")
                        dt_det.Columns.Remove("fin_coax_len")

                    End If
                Catch ex As Exception
                    err = "ER: Error deleting cols of the cr " & cr_form_type & " file, details: " & ex.ToString
                    GoTo get_out
                End Try

                Debug.WriteLine(Now.ToLongTimeString & ": write data to XL")

                'write the details table to XL, the common is already in-sync
                '----------------------------------------------------------
                dt2xlrange(last_data_row, dt_det, xlsheet, format, err)
                If Not err Like "" Then GoTo get_out

                'we then move ALL attachments to the appropriate CR dir
                '------------------------------------------------------
                Try
                    Debug.WriteLine(Now.ToLongTimeString & ": doing attachments")

                    'create requester attachments dir in the cr dir
                    '-----------------------------------------------------
                    If FileIO.FileSystem.DirectoryExists(cr_path & "\requester attachments") Then
                        clean_dir(cr_path & "\requester attachments", err)
                        If Not err Like "" Then GoTo get_out
                    Else
                        FileIO.FileSystem.CreateDirectory(cr_path & "\requester attachments")
                    End If
                    Dim req_attach_path As String = cr_path & "\requester attachments"

                    'this just moves all attachments to each cr dir as I want to pass everything the user puts in, regardless
                    '----------------------------------------------------------------------------------------------------
                    For Each item In FileIO.FileSystem.GetFiles(Path.GetDirectoryName(file_in))
                        If Not Path.GetFileName(item) Like "~$*" AndAlso Not item Like file_in AndAlso Not FileIO.FileSystem.FileExists(req_attach_path & "\" & Path.GetFileName(item)) Then
                            FileIO.FileSystem.CopyFile(item, req_attach_path & "\" & Path.GetFileName(item), True)
                        End If
                    Next
                    For Each item In FileIO.FileSystem.GetDirectories(Path.GetDirectoryName(file_in))
                        Dim dirInfo As New System.IO.DirectoryInfo(item)
                        Dim dir As String = dirInfo.Name
                        If Not FileIO.FileSystem.DirectoryExists(req_attach_path & "\" & dir) Then
                            FileIO.FileSystem.CopyDirectory(item, req_attach_path & "\" & dir, True)
                        End If
                    Next

                    'zip the attachments and delete the requester attachments dir
                    '------------------------------------------------------------
                    If (FileIO.FileSystem.GetFiles(req_attach_path).Count + FileIO.FileSystem.GetDirectories(req_attach_path).Count) > 0 Then
                        zip_dir(req_attach_path, 10, 20, cr_path & "\requester attachments.zip", err)
                        If Not err Like "" Then GoTo get_out
                    End If
                    FileIO.FileSystem.DeleteDirectory(req_attach_path, DeleteDirectoryOption.DeleteAllContents)
                Catch ex As Exception
                    err = "ER: General error zipping the cr form '" & file_in & "'.  Details: " & ex.ToString
                    GoTo get_out
                End Try

                'resets the formatting of the common and detail cells on the sheet
                '------------------------------------------------------------
                're-adjust the size of the common fields
                '--------------------------------------------
                Debug.WriteLine(Now.ToLongTimeString & ": finalising formatting")

                Dim t_off As Integer = 0
                If cr_form_type = "rfb" Then
                    t_off = 10
                ElseIf cr_form_type = "prm" Then
                    t_off = 7
                Else
                    t_off = 8
                End If
                Dim t_rng As Excel.Range = xlsheet.Range(xlsheet.Cells(1, 1), xlsheet.Cells(1, format.common_data_col + t_off))
                t_rng.UnMerge()
                t_rng.Merge(t_rng.MergeCells)
                normal_format_cell_resub(t_rng)

                For col_index = 0 To com_hdr_name.Count - 1
                    Dim xlcell As Excel.Range = xlsheet.Range(xlsheet.Cells(format.common_row_start + col_index, format.common_data_col), xlsheet.Cells(format.common_row_start + col_index, format.common_data_col))
                    t_rng = xlsheet.Range(xlsheet.Cells(format.common_row_start + col_index, format.common_data_col), xlsheet.Cells(format.common_row_start + col_index, format.common_data_col + t_off))
                    t_rng.UnMerge()
                    t_rng.Merge(t_rng.MergeCells)
                    normal_format_cell_resub(t_rng)
                Next

                Debug.WriteLine(Now.ToLongTimeString & ": borders")

                'does the common borders
                '---------------------------
                t_rng = xlsheet.Range(xlsheet.Cells(format.common_row_start, format.common_data_col), xlsheet.Cells(format.common_row_start + com_hdr_name.Count - 1, format.common_data_col + t_off))
                set_border_common_data(t_rng)

                'does detail format reset
                '------------------------------
                t_rng = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, 1), xlsheet.Cells(last_data_row, det_hdr_name.Count))
                normal_format_cell_resub(t_rng)

                'sets the hide and unlock config for the next stage (approver) => there is no unlocking for the next stage
                '-------------------------------------------------------------------------------------------------------------
                Dim col_array() As Integer = get_integer_array_from_name("col_hide_" & cr_form_type & "_for_app", format, err)
                If Not err Like "" Then GoTo get_out
                hide_rows_and_cols(format.row_hide_for_app, col_array, xlsheet, format, err)

                'protects the sheet and book
                '-----------------------------
                xlsheet.Range("C2").Select()
                xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True)

                Debug.WriteLine(Now.ToLongTimeString & ": save and exit")

                'save the new cr form
                '-----------------------------
                Try
                    file_out = cr_path & "\" & cr_id & ".xlsb"
                    xlapp.DisplayAlerts = False
                    xlbook.SaveAs(file_out)
                Catch ex As Exception
                    err = "ER: Error saving the " & cr_form_type & " cr form '" & file_in & "'.  Details: " & ex.ToString
                End Try
get_out:
            Catch ex As Exception
                err = "ER: General error creating the " & cr_form_type & " cr form '" & file_in & "'.  Details: " & ex.ToString
            Finally
                Try
                    xlapp.DisplayAlerts = False
                    xlbook.Close()
                Finally
                    releaseObject(xlbook)
                    xlapp.UserControl = True
                    xlapp.Interactive = True
                    xlapp.IgnoreRemoteRequests = False
                    xlapp.Quit()
                    releaseObject(xlapp) 'this releases the com object
                End Try
            End Try
        Catch ex As Exception
            err = "ER: Error opening the XL application, details: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub





    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    'this splits the blank cr form into the 3 blank cr form type templates (prm, rfb, oth)
    '--------------------------------------------------------------------------------
    Public Sub split_blank_cr_form(ByVal file_in As String, ByVal cr_form_type As String, ByVal format As cr_sheet_format, ByVal local As local_machine, ByRef err As String)
        Try
            'opens a new instance of XL
            '----------------------------------------------
            Dim xlapp As New Excel.Application
            xlapp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow
            xlapp.Visible = format.debug_xl
            xlapp.DisplayAlerts = False
            xlapp.UserControl = False
            xlapp.IgnoreRemoteRequests = True
            xlapp.Interactive = False
            Dim xlbook As Excel.Workbook = Nothing

            Try
                Dim com_hdr_name() As String = {}
                com_hdr_name = format.common_hdr_name
                Dim det_hdr_name() As String = {}
                det_hdr_name = get_string_array_from_name("detail_hdr_name_" & cr_form_type, format, err)
                If Not err Like "" Then GoTo get_out

                'Open the blank cr form
                '------------------------
                Dim cr_form_to_open As String = file_in
                Try
                    Debug.WriteLine(Now.ToLongTimeString & ": opening blank cr form for splitting: " & cr_form_type)
                    xlapp.DisplayAlerts = False
                    xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                Catch ex As Exception
                    err = "ER: Internal error opening blank cr_form, details: " & ex.ToString
                    GoTo get_out
                End Try

                Dim xlsheet As Excel.Worksheet = xlbook.Worksheets("CR")
                xlsheet.Activate()

                'unprotects book and sheet and unhides cells
                '--------------------------------------------
                Try
                    Debug.WriteLine(Now.ToLongTimeString & ": unprotecting book")

                    xlsheet.Unprotect(format.x_factor)
                    xlsheet.Range("A1").Value2 = "z"        'this disables the worksheet change macro or reset macro
                    unhide_and_lock_all(format.detail_hdr_row_start, format.detail_hdr_name.Count, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out
                Catch ex As Exception
                    err = "ER: some error unprotecting or unhiding and locking cells, form check sub"
                    GoTo get_out
                End Try

                'remove the reset button
                '------------------------
                For Each item As Excel.Shape In xlsheet.Shapes
                    If item.Name = "CommandButton1" Then
                        item.Delete()
                        Exit For
                    End If
                Next

                Debug.WriteLine(Now.ToLongTimeString & ": deleting cols")

                'delete unwanted cols from the DT and sheet
                '-----------------------------------------------
                Try
                    If cr_form_type = "rfb" Then
                        xlsheet.Range("E:H").EntireColumn.Delete()
                    ElseIf cr_form_type = "prm" Then
                        xlsheet.Range("I:N").EntireColumn.Delete()
                        xlsheet.Range("M:U").EntireColumn.Delete()
                    Else
                        xlsheet.Range("E:N").EntireColumn.Delete()
                        xlsheet.Range("I:Q").EntireColumn.Delete()
                    End If
                Catch ex As Exception
                    err = "ER: Error deleting cols of the cr " & cr_form_type & " file, details: " & ex.ToString
                    GoTo get_out
                End Try

                'resets the formatting of the common and detail cells on the sheet
                '------------------------------------------------------------
                're-adjust the size of the common fields
                '--------------------------------------------
                Debug.WriteLine(Now.ToLongTimeString & ": finalising formatting")

                Dim t_off As Integer = 0
                If cr_form_type = "rfb" Then : t_off = 10
                ElseIf cr_form_type = "prm" Then : t_off = 7
                Else : t_off = 8
                End If
                Dim t_rng As Excel.Range = xlsheet.Range(xlsheet.Cells(1, 1), xlsheet.Cells(1, format.common_data_col + t_off))
                t_rng.UnMerge()
                t_rng.Merge(t_rng.MergeCells)
                normal_format_cell_resub(t_rng)

                For col_index = 0 To com_hdr_name.Count - 1
                    Dim xlcell As Excel.Range = xlsheet.Range(xlsheet.Cells(format.common_row_start + col_index, format.common_data_col), xlsheet.Cells(format.common_row_start + col_index, format.common_data_col))
                    t_rng = xlsheet.Range(xlsheet.Cells(format.common_row_start + col_index, format.common_data_col), xlsheet.Cells(format.common_row_start + col_index, format.common_data_col + t_off))
                    t_rng.UnMerge()
                    t_rng.Merge(t_rng.MergeCells)
                    normal_format_cell_resub(t_rng)
                Next

                Debug.WriteLine(Now.ToLongTimeString & ": borders")

                'does the common borders
                '---------------------------
                t_rng = xlsheet.Range(xlsheet.Cells(format.common_row_start, format.common_data_col), xlsheet.Cells(format.common_row_start + com_hdr_name.Count - 1, format.common_data_col + t_off))
                set_border_common_data(t_rng)

                'protects the sheet and book
                '-----------------------------
                xlsheet.Range("C2").Select()
                xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True)

                Debug.WriteLine(Now.ToLongTimeString & ": save and exit")

                'save the blank template form
                '-----------------------------
                Try
                    Dim file_out As String = local.base_path & local.cr_blank_request_form_dir & "\blank_template_" & cr_form_type & ".xlsb"
                    If FileIO.FileSystem.FileExists(file_out) Then
                        force_delete_file(file_out, err)
                        If Not err Like "" Then GoTo get_out
                    End If
                    xlapp.DisplayAlerts = False
                    xlbook.SaveAs(file_out)
                Catch ex As Exception
                    err = "ER: Error saving the " & cr_form_type & " blank template cr form, details: " & ex.ToString
                End Try
get_out:
            Catch ex As Exception
                err = "ER: General error creating the " & cr_form_type & " blank template cr form, details: " & ex.ToString
            Finally
                Try
                    xlapp.DisplayAlerts = False
                    xlbook.Close()
                Finally
                    releaseObject(xlbook)
                    xlapp.UserControl = True
                    xlapp.Interactive = True
                    xlapp.IgnoreRemoteRequests = False
                    xlapp.Quit()
                    releaseObject(xlapp) 'this releases the com object
                End Try
            End Try
        Catch ex As Exception
            err = "ER: Error opening the XL application, details: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub






    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    'prepares the cr_form to send to the ex.coord
    '--------------------------------------------
    Public Sub prepare_cr_form_for_excoord(ByVal approval_date As DateTime, ByVal cr_form As String, ByVal cr_form_type As String, ByVal format As cr_sheet_format, ByVal local As local_machine, ByVal err As String)
        Try
            Dim xlapp As New Excel.Application
            xlapp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow
            xlapp.Visible = format.debug_xl
            xlapp.DisplayAlerts = False
            xlapp.UserControl = False
            xlapp.IgnoreRemoteRequests = True
            xlapp.Interactive = False
            Dim xlbook As Excel.Workbook = Nothing

            Try
                'find the com and detail header name arrays
                '--------------------------------------------
                Dim com_hdr_name() As String = {}
                com_hdr_name = format.common_hdr_name
                Dim det_hdr_name() As String = get_string_array_from_name("detail_hdr_name_" & cr_form_type, format, err)
                If Not err Like "" Then GoTo get_out

                'this opens the xl file
                '-------------------
                Dim cr_form_to_open As String = cr_form
                Try
                    Debug.WriteLine(Now.ToLongTimeString & ": open XL")
                    xlapp.DisplayAlerts = False
                    xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                Catch ex As COMException
                    If ex.HResult = -2146827284 Then
                        fix_bad_xlsb_file(cr_form_to_open, xlapp, format, local, err)
                        If Not err Like "" Then GoTo get_out
                        Try
                            xlapp.DisplayAlerts = False
                            xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                        Catch exx As Exception
                            err = "ER: Internal error opening xl file after fixing, details: " & exx.ToString
                            GoTo get_out
                        End Try
                    Else
                        err = "ER: Internal COM error opening xl file, details: " & ex.ToString
                        GoTo get_out
                    End If
                Catch ex As Exception
                    err = "ER: Internal error opening xl file, details: " & ex.ToString
                    GoTo get_out
                End Try

                Dim xlsheet As Excel.Worksheet
                Try
                    xlsheet = xlbook.Worksheets("CR")
                Catch ex As Exception
                    err = "APPREJ: Can't find the CR form...."
                    GoTo get_out
                End Try
                xlsheet.Activate()

                Debug.WriteLine(Now.ToLongTimeString & ": unprotect")

                'unprotects the sheet - no need to do security checks as this sheet has not left the HDD
                '-----------------------------------------------
                Try
                    xlsheet.Unprotect(format.x_factor)
                    xlsheet.Range("A1").Value2 = "z"        'this disables the worksheet change macro or reset macro
                    unhide_and_lock_all(format.detail_hdr_row_start, 100, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out
                Catch ex As Exception
                    err = "ER: some error unprotecting or unhiding and locking cells, form check sub"
                    GoTo get_out
                End Try

                Debug.WriteLine(Now.ToLongTimeString & ": clearing extra cols")

                'this clear cells after the last detailed col ("executor comments") - needed to  deal with older forms
                '-------------------------------------------------------------------------------
                Try
                    Dim r1 As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + det_hdr_name.Count), xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + det_hdr_name.Count + 100))
                    r1.EntireColumn.Clear()
                Catch ex As Exception
                End Try

                Debug.WriteLine(Now.ToLongTimeString & ": doing other formatting")

                'setup ranges
                '------------------------------------
                Dim index_cr_type As Integer = Array.IndexOf(det_hdr_name, "CR Type")
                Dim index_planned_ex_date As Integer = Array.IndexOf(det_hdr_name, "Planned Execution Date")
                Dim index_app_date As Integer = Array.IndexOf(format.common_hdr_name, "Approval Date")

                'First we find the row limit for the cr form
                '----------------------------------------------
                Dim last_data_row As Integer = 0
                Try
                    With xlsheet
                        last_data_row = .Columns("A:A").offset(0, index_cr_type).entirecolumn.Find(What:="*", After:=.Cells(index_cr_type + 1), LookAt:=XlLookAt.xlWhole, LookIn:=XlFindLookIn.xlValues, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False).Row
                    End With
                Catch ex As Exception
                    last_data_row = 0
                End Try
                If last_data_row = 0 Then
                    err = "APPREJ: The CR form format has been corrupted, can't find the CR Type column...."
                    GoTo get_out
                End If

                'sets date formats for the date cells
                '---------------------------------------------
                Dim temp_range As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_planned_ex_date + 1), xlsheet.Cells(last_data_row, index_planned_ex_date + 1))
                date_format_cell(temp_range)

                'adds in the approval date to the header
                '-------------------------------------
                temp_range = xlsheet.Range(xlsheet.Cells(format.common_row_start + index_app_date, format.common_data_col), xlsheet.Cells(format.common_row_start + index_app_date, format.common_data_col))
                temp_range.Value2 = approval_date.ToOADate
                date_format_cell(temp_range)

                'resets the formatting of the common and detail cells on the sheet
                '------------------------------------------------------------
                Dim t_rng As Excel.Range = xlsheet.Range(xlsheet.Cells(format.common_row_start, format.common_data_col), xlsheet.Cells(format.common_row_start + com_hdr_name.Count - 1, format.common_data_col))
                normal_format_cell_resub(t_rng)
                t_rng = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, 1), xlsheet.Cells(last_data_row, det_hdr_name.Count))
                normal_format_cell_resub(t_rng)

                'sets the hide and unlock config for the next stage (ex_coord)
                '------------------------------------------------------------------
                Dim col_array() As Integer = get_integer_array_from_name("col_hide_" & cr_form_type & "_for_ex_coord", format, err)
                If Not err Like "" Then GoTo get_out
                hide_rows_and_cols(format.row_hide_for_ex_coord, col_array, xlsheet, format, err)
                col_array = get_integer_array_from_name("col_unprotect_" & cr_form_type & "_for_ex_coord", format, err)
                If Not err Like "" Then GoTo get_out
                unlock_and_input_format({}, col_array, last_data_row, cr_form_type, False, xlsheet, format, err)

                'protects the sheet and book
                '----------------------------------
                xlsheet.Range("C2").Select()
                xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True)

                Debug.WriteLine(Now.ToLongTimeString & ": save and exit")

                xlapp.DisplayAlerts = False
                xlbook.SaveAs(cr_form)
get_out:
            Catch ex As Exception
                err = "ER: General error preparing cr_form for executor, details: " & ex.ToString
            Finally
                Try
                    xlapp.DisplayAlerts = False
                    xlbook.Close()
                Finally
                    releaseObject(xlbook)
                    xlapp.UserControl = True
                    xlapp.Interactive = True
                    xlapp.IgnoreRemoteRequests = False
                    xlapp.Quit()
                    releaseObject(xlapp) 'this releases the com object
                End Try
            End Try
        Catch ex As Exception
            err = "ER: Error opening the XL application, details: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub










    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    'this checks the format of the incoming cr_form from the execution coordinator.  
    'if the format is ok, it checks all values meant to be filled are indeed filled
    Public Sub process_ex_coord_cr_form(ByRef data_ok As String, ByVal cr_id As String, ByVal cr_form As String, ByRef planned_ex_date As Date, ByRef executors_raw As String, ByRef dt As System.Data.DataTable, ByVal cr_form_type As String, ByVal format As cr_sheet_format, ByVal local As local_machine, ByRef err As String)
        Try
            'opens a new instance of XL
            '----------------------------------------------
            Dim xlapp As New Excel.Application
            xlapp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow
            xlapp.Visible = format.debug_xl
            xlapp.DisplayAlerts = False
            xlapp.UserControl = False
            xlapp.IgnoreRemoteRequests = True
            xlapp.Interactive = False
            Dim xlbook As Excel.Workbook = Nothing
            Dim cr_path As String = Path.GetDirectoryName(cr_form)

            Try
                'find the com and detail header name arrays
                '--------------------------------------------
                Dim com_hdr_name() As String = {}
                com_hdr_name = format.common_hdr_name
                Dim det_hdr_name() As String = get_string_array_from_name("detail_hdr_name_" & cr_form_type, format, err)
                If Not err Like "" Then GoTo get_out

                'this opens the xl file
                '-------------------
                Dim cr_form_to_open As String = cr_form
                Try
                    Debug.WriteLine(Now.ToLongTimeString & ": open XL")
                    xlapp.DisplayAlerts = False
                    xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                Catch ex As COMException
                    If ex.HResult = -2146827284 Then
                        fix_bad_xlsb_file(cr_form_to_open, xlapp, format, local, err)
                        If Not err Like "" Then GoTo get_out
                        Try
                            xlapp.DisplayAlerts = False
                            xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                        Catch exx As Exception
                            err = "ER: Internal error opening xl file after fixing, details: " & exx.ToString
                            GoTo get_out
                        End Try
                    Else
                        err = "ER: Internal COM error opening xl file, details: " & ex.ToString
                        GoTo get_out
                    End If
                Catch ex As Exception
                    err = "ER: Internal error opening xl file, details: " & ex.ToString
                    GoTo get_out
                End Try

                'this checks the CR sheet exists
                '--------------------------
                Dim xlsheet As Excel.Worksheet
                Try
                    xlsheet = xlbook.Worksheets("CR")
                Catch ex As Exception
                    err = "EXCOORDREJ: CR form rejection.  The CR form (" & Path.GetFileName(cr_form) & ") doesn't have a sheet called 'CR'.<BR>Thanks"
                    GoTo get_out
                End Try

                'unprotects the sheet and book
                'first checks that the sheet is protected, if it is not, then someone has switched the sheet.
                '-------------------------------------------------------------------------------------------
                If Not (xlsheet.ProtectContents And xlsheet.ProtectScenarios) Then
                    err = "EXCOORDREJ: CR form is not genuine.  The CR form (" & Path.GetFileName(cr_form) & ") appears to be fake, please use the form that was emailed to you in the planned date request'.<BR>Thanks"
                    GoTo get_out
                End If

                Debug.WriteLine(Now.ToLongTimeString & ": unprotecting")

                'unprotect, unhide and lock
                '------------------------------
                Try
                    xlsheet.Unprotect(format.x_factor)
                    xlsheet.Range("A1").Value2 = "z"        'this disables the worksheet change macro or reset macro
                    unhide_and_lock_all(format.detail_hdr_row_start, 100, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out
                Catch ex As Exception
                    err = "ER: some error unprotecting or unhiding and locking cells, form check sub"
                    GoTo get_out
                End Try
                'at this point, I know the form has not changed or been switched as it has passed the protection test

                Debug.WriteLine(Now.ToLongTimeString & ": clear extra cols")

                'this clear cells after the last detailed col ("executor comments") - needed to  deal with older forms
                '-------------------------------------------------------------------------------
                Try
                    Dim r1 As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + det_hdr_name.Count), xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + det_hdr_name.Count + 100))
                    r1.EntireColumn.Clear()
                Catch ex As Exception
                End Try

                'setup some ranges
                '--------------------------
                Dim comindex_ex_pl_date As Integer = Array.IndexOf(com_hdr_name, "Planned Execution Date")
                Dim comindex_ex_coord As Integer = Array.IndexOf(com_hdr_name, "Execution Coordinator")
                Dim comindex_ex As Integer = Array.IndexOf(com_hdr_name, "Executors")

                Dim index_cr_type As Integer = Array.IndexOf(det_hdr_name, "CR Type")
                Dim index_planned_ex_date As Integer = Array.IndexOf(det_hdr_name, "Planned Execution Date")
                Dim index_executor As Integer = Array.IndexOf(det_hdr_name, "Executor")
                Dim index_ex_date As Integer = Array.IndexOf(det_hdr_name, "Execution Date")

                'First we find the row limit for the cr form
                '----------------------------------------------
                Dim last_data_row As Integer = 0
                Try
                    With xlsheet
                        last_data_row = .Columns("A:A").offset(0, index_cr_type).entirecolumn.Find(What:="*", After:=.Cells(index_cr_type + 1), LookAt:=XlLookAt.xlWhole, LookIn:=XlFindLookIn.xlValues, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False).Row
                    End With
                Catch ex As Exception
                    last_data_row = 0
                End Try
                If last_data_row = 0 Then
                    err = "EXCOORDREJ: The CR form format has been corrupted, can't find the CR Type column...."
                    GoTo get_out
                End If

                Debug.WriteLine(Now.ToLongTimeString & ": reading to DT")

                'reads the edited range into a dt for checking and updating the DB
                '-----------------------------------------------------------
                Dim temp_range As Excel.Range
                Try
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_planned_ex_date + 1), xlsheet.Cells(last_data_row, index_executor + 1))
                    Dim cr_sub_id_range As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, format.detail_col_start), xlsheet.Cells(last_data_row, format.detail_col_start))
                    xlrange2dt({"planned_execution_date", "executor"}, cr_sub_id_range, temp_range, dt, format, err)
                    If Not err Like "" Then GoTo get_out
                Catch ex As Exception
                    err = "ER: Tool error reading xl to dt, checking ex coord returned values for executor, details: " & ex.ToString
                    GoTo get_out
                End Try

                Debug.WriteLine(Now.ToLongTimeString & ": check planned date")

                'checks the cells that should have been filled are indeed filled
                '----------------------------------------------------------
                err = ""
                Try
                    For Each row As System.Data.DataRow In dt.Rows
                        Dim ts1 As String = row.Field(Of String)("planned_execution_date")
                        If ts1 Like "" Then
                            err = "ERROR!! - Must be a date"
                            row("planned_execution_date") = err & " - " & ts1
                        Else
                            Dim ts2 As DateTime = Now
                            Dim temp_err As String = ""
                            get_xl_date(ts1, ts2, temp_err)            'more robust sub to handle all types of weird date input
                            If Not temp_err Like "" Then
                                err = "ERROR!! - " & temp_err
                                row("planned_execution_date") = err & " - " & ts1

                            ElseIf ts2.Date < Now.Date Then
                                err = "ERROR!! - Must not be a past date"
                                row("planned_execution_date") = err & " - " & ts2.ToLongDateString

                            Else
                                If ts2.Date > planned_ex_date.Date Then
                                    planned_ex_date = ts2.Date
                                End If
                                row("planned_execution_date") = ts2.Date.ToOADate
                            End If
                        End If
                    Next

                    'this sets the planned_ex_date time to 6pm on the chosen date or 6hours after now if the chosen date is today, before we set it, 
                    '-------------------------------------------------------------------------------------
                    If planned_ex_date.Date = Now.Date Then
                        'if it gets here, planned ex date = now with time
                        Dim timespan As New TimeSpan(format.planned_ex_time_delay_same_day, 0, 0)
                        planned_ex_date = planned_ex_date + timespan
                    Else
                        'If it gets here, planned ex date only has date
                        Dim timespan As New TimeSpan(format.planned_ex_time_future, 0, 0)
                        planned_ex_date = planned_ex_date.Date + timespan
                    End If
                Catch ex As Exception
                    err = "ER: Error checking dates, details: " & ex.ToString
                    GoTo get_out
                End Try
                If Not err Like "" Then
                    data_ok = "nok"
                    err = ""
                End If

                Debug.WriteLine(Now.ToLongTimeString & ": check ex")

                'checks the executors - allows the ex coord to enter themselves at this stage if they do not know which executor will work the job
                '------------------------------------------------------------------------------
                err = ""
                temp_range = xlsheet.Range(xlsheet.Cells(format.common_row_start + comindex_ex_coord, format.common_data_col), xlsheet.Cells(format.common_row_start + comindex_ex_coord, format.common_data_col))
                Dim execution_coordinator As String = o2s(temp_range.Value2)
                Try
                    If execution_coordinator Like "" Then
                        err = "ERROR!! - Must be filled"
                        error_format_cell(False, temp_range, err & " - " & execution_coordinator)
                    Else
                        Dim qrows = From row In dt Where Not format.IsValidEmail(c2e(row.Field(Of String)("executor"))) Select row
                        For Each row In qrows
                            err = "ERROR!! - Must contain a valid email address'"
                            row("executor") = err & " - " & row.Field(Of String)("executor")
                        Next
                        qrows = From row In qrows Where format.IsValidEmail(c2e(row.Field(Of String)("executor"))) Select row
                        'test_vals includes the ex coord here
                        Dim test_vals() As String = (From row In format.ds_allow.Tables("executors") Select row.Field(Of String)("combined_name")).Distinct.ToArray
                        Dim vals_cnt As Integer = test_vals.Count
                        ReDim Preserve test_vals(vals_cnt)
                        test_vals(vals_cnt) = execution_coordinator
                        test_vals = test_vals.Distinct.ToArray
                        qrows = From row In qrows Where Not test_vals.Any(Function(s) row.Field(Of String)("executor").Contains(s)) Select row
                        For Each row In qrows
                            err = "ERROR!! - Unknown Executor, you must enter a registered executor or the execution coordinator if you do not know yet - to register an executor, please send mail: Subject => Add Executor and Body => name,email; => one executor per line"
                            row("executor") = err & " - " & row.Field(Of String)("executor")
                        Next
                    End If
                Catch ex As Exception
                    err = "ER: Tool error checking executors, details: " & ex.ToString
                    GoTo get_out
                End Try
                If Not err Like "" Then
                    data_ok = "nok"
                    err = ""
                End If
                If data_ok Like "not finished testing" Then data_ok = "ok"

                '#######################################################################################
                '#######################################################################################
                'if we got a format error exit now
                '-----------------------------------
                If data_ok = "nok" Then
                    Debug.WriteLine(Now.ToLongTimeString & ": data is nok so cleaning up")
                    'writes the details table with errors data back to the xlsheet for the error file
                    '-----------------------------------------------------------------------
                    err = ""
                    ds2xl_1col(xlsheet, dt, 1, index_planned_ex_date + 1, format.detail_data_row_start, err)
                    If Not err Like "" Then GoTo get_out
                    ds2xl_1col(xlsheet, dt, 2, index_executor + 1, format.detail_data_row_start, err)
                    If Not err Like "" Then GoTo get_out

                    'Reset the detail formats before doing the error formatting
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, 1), xlsheet.Cells(last_data_row, det_hdr_name.Count))
                    normal_format_cell_resub(temp_range)
                    'we only do the error formatting on small to med CRs as it involves filtering which gets sketchy on large ranges, L will throw and behave badly
                    'for big crs we just write the error data 
                    set_all_errors_cell2red_unlock(last_data_row, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out

                    'sets date format for the execution date in the detail data
                    '---------------------------------------------------------
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_planned_ex_date + 1), xlsheet.Cells(last_data_row, index_planned_ex_date + 1))
                    date_format_cell(temp_range)
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_date + 1), xlsheet.Cells(last_data_row, index_ex_date + 1))
                    date_format_cell(temp_range)

                    'hides fields again to the original ex_coord format - all protected though
                    '--------------------------------------------------------------------
                    Dim x_a() As Integer = get_integer_array_from_name("col_hide_" & cr_form_type & "_for_ex_coord", format, err)
                    If Not err Like "" Then GoTo get_out
                    hide_rows_and_cols(format.row_hide_for_ex_coord, x_a, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out

                    'protects the sheet and book
                    '----------------------------------
                    xlsheet.Range("C2").Select()
                    xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True)

                    Dim file_out As String = ""
                    file_out = cr_path & "\" & Path.GetFileNameWithoutExtension(cr_form) & "_errors.xlsb"
                    xlapp.DisplayAlerts = False
                    xlbook.SaveAs(file_out)
                    err = "XREJ: See returned CR form (CRMS: CR Execution Planning Retry Request (" & cr_id & "))(" & file_out & ")"
                    GoTo get_out
                End If
                '#######################################################################################
                '#######################################################################################

                Debug.WriteLine(Now.ToLongTimeString & ": getting ex list")

                'we get a list of the executors and add the list to the common values
                '--------------------------------------------------------------
                executors_raw = ""
                Try
                    Dim qrows() As String = (From row In dt Let a = row.Field(Of String)("executor") Select a).Distinct.ToArray
                    executors_raw = Join(qrows, ",")
                Catch ex As Exception
                    err = "ER: Tool error getting executors, details: " & ex.ToString
                    GoTo get_out
                End Try
                'this checks and cleans the executors list => takes out doubles, yes even though I already did that and checks all values contain valid email addresses
                executors_raw = check_email_list(False, executors_raw, format)
                'sets the list to the ex coord if it is a blank list
                executors_raw = If(executors_raw Like "", execution_coordinator, executors_raw)
                'this adds the executors list to the common vals
                temp_range = xlsheet.Range(xlsheet.Cells(format.common_row_start + comindex_ex, format.common_data_col), xlsheet.Cells(format.common_row_start + comindex_ex, format.common_data_col))
                temp_range.Value2 = executors_raw

                Debug.WriteLine(Now.ToLongTimeString & ": filling header")

                'adds in the ex_coord_planned_date to the header
                '-------------------------------------------------
                temp_range = xlsheet.Range(xlsheet.Cells(format.common_row_start + comindex_ex_pl_date, format.common_data_col), xlsheet.Cells(format.common_row_start + comindex_ex_pl_date, format.common_data_col))
                temp_range.Value2 = planned_ex_date.ToOADate
                date_format_cell(temp_range)

                'writes the planned ex date back to the details table as date may have been reformatted during checking
                '-------------------------------------------------------------------------------------------------------
                ds2xl_1col(xlsheet, dt, 1, index_planned_ex_date + 1, format.detail_data_row_start, err)
                If Not err Like "" Then GoTo get_out

                'sets date format for the execution date in the detail data
                '---------------------------------------------------------
                temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_planned_ex_date + 1), xlsheet.Cells(last_data_row, index_planned_ex_date + 1))
                date_format_cell(temp_range)
                temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_date + 1), xlsheet.Cells(last_data_row, index_ex_date + 1))
                date_format_cell(temp_range)

                Debug.WriteLine(Now.ToLongTimeString & ": finalising formatting")

                'resets the formatting of the common and detail cells on the sheet
                '------------------------------------------------------------
                Dim t_rng As Excel.Range = xlsheet.Range(xlsheet.Cells(format.common_row_start, format.common_data_col), xlsheet.Cells(format.common_row_start + com_hdr_name.Count - 1, format.common_data_col))
                normal_format_cell_resub(t_rng)
                t_rng = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, 1), xlsheet.Cells(last_data_row, det_hdr_name.Count))
                normal_format_cell_resub(t_rng)

                'sets the hide and unlock config for the next stage (ex)
                '----------------------------------------------------------------------------------------------
                Dim col_array() As Integer = get_integer_array_from_name("col_hide_" & cr_form_type & "_for_ex", format, err)
                If Not err Like "" Then GoTo get_out
                hide_rows_and_cols(format.row_hide_for_ex, col_array, xlsheet, format, err)
                col_array = get_integer_array_from_name("col_unprotect_" & cr_form_type & "_for_ex", format, err)
                If Not err Like "" Then GoTo get_out
                unlock_and_input_format({}, col_array, last_data_row, cr_form_type, False, xlsheet, format, err)

                'protects the sheet and book
                '----------------------------------
                xlsheet.Range("C2").Select()
                xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True)

                Debug.WriteLine(Now.ToLongTimeString & ": save and exit")

                'save the file
                '--------------------
                xlapp.DisplayAlerts = False
                xlbook.SaveAs(cr_form)
get_out:
            Catch ex As Exception
                err = "ER: General error checking the cr file '" & cr_form & "'.  Details: " & ex.ToString
            Finally
                Try
                    xlapp.DisplayAlerts = False
                    xlbook.Close()
                Finally
                    releaseObject(xlbook)
                    xlapp.UserControl = True
                    xlapp.Interactive = True
                    xlapp.IgnoreRemoteRequests = False
                    xlapp.Quit()
                    releaseObject(xlapp) 'this releases the com object
                End Try
            End Try
        Catch ex As Exception
            err = "ER: Error opening the XL application, details: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub











    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    'this checks the format of the incoming cr_form from the executor
    'if the format is ok, it checks all values meant to be filled are indeed filled
    'NOTE THAT HERE i ONLY USE THE IN-MEMORY DT FOR CHECKING THE VALS THAT ARE IN THE PRM FORM, THE AZ AND MDT ETC ARE STILL CHECKED IN XL AND I LIMIT THE CR ROWS TO 1000 FOR RFB AND RFR CRS
    'AS IT IS NOT REALISTIC THERE WILL EVER BE MORE THAN 100 OR SO.  I CAN CHANGE THE CODE LATER TO CHECK IN MEMORY THOUGH IF REQUIRED.
    Public Sub process_ex_cr_form(ByRef executors_raw As String, ByVal cr_id As String, ByVal cr_status As String, ByVal cr_form As String, ByRef dt As System.Data.DataTable, ByVal cr_type As String, ByVal cr_form_type As String, ByVal ex_date As Date, ByRef attach_ok As String, ByRef data_ok As String, ByVal local As local_machine, format As cr_sheet_format, ByRef err As String)
        Try
            'opens a new instance of XL
            '----------------------------------------------
            Dim xlapp As New Excel.Application
            xlapp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow
            xlapp.Visible = format.debug_xl
            xlapp.DisplayAlerts = False
            xlapp.UserControl = False
            xlapp.IgnoreRemoteRequests = True
            xlapp.Interactive = False
            Dim xlbook As Excel.Workbook = Nothing
            Dim inbox_path As String = local.base_path & local.inbox
            Dim temp_range As Excel.Range

            Try
                'find the com and detail header name arrays
                '--------------------------------------------
                Dim com_hdr_name() As String = {}
                com_hdr_name = format.common_hdr_name
                Dim det_hdr_name() As String = get_string_array_from_name("detail_hdr_name_" & cr_form_type, format, err)
                If Not err Like "" Then GoTo get_out

                'this opens the xl file
                '-------------------
                Dim cr_form_to_open As String = cr_form
                Try
                    Debug.WriteLine(Now.ToLongTimeString & ": open XL")

                    'do this during debugging
                    If cr_status Like "Execution Complete Pending Attachments" Then
                        pre_open_non_inbox_cr_file_conflict_resolution(cr_form, err)
                        If Not err Like "" Then GoTo get_out
                    End If

                    xlapp.DisplayAlerts = False
                    xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                Catch ex As COMException
                    If ex.HResult = -2146827284 Then
                        fix_bad_xlsb_file(cr_form_to_open, xlapp, format, local, err)
                        If Not err Like "" Then GoTo get_out
                        Try
                            xlapp.DisplayAlerts = False
                            xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                        Catch exx As Exception
                            err = "ER: Internal error opening xl file after fixing, details: " & exx.ToString
                            GoTo get_out
                        End Try
                    Else
                        err = "ER: Internal COM error opening xl file, details: " & ex.ToString
                        GoTo get_out
                    End If
                Catch ex As Exception
                    err = "ER: Internal error opening xl file, details: " & ex.ToString
                    GoTo get_out
                End Try

                'this checks the CR sheet exists
                '--------------------------
                Dim xlsheet As Excel.Worksheet
                Try
                    xlsheet = xlbook.Worksheets("CR")
                Catch ex As Exception
                    err = "EXREJ: CR form rejection.  The CR form (" & Path.GetFileName(cr_form) & ") doesn't have a sheet called 'CR'.<BR>Thanks"
                    GoTo get_out
                End Try

                'first checks that the sheet is protected, if it is not, then someone has switched the sheet.
                '-------------------------------------------------------------------------------------------
                If Not (xlsheet.ProtectContents And xlsheet.ProtectScenarios) Then
                    err = "EXREJ: CR form is not genuine.  The CR form (" & Path.GetFileName(cr_form) & ") appears to be fake, please use the form that was emailed to you in the execution request'.<BR>Thanks"
                    GoTo get_out
                End If

                Debug.WriteLine(Now.ToLongTimeString & ": unprotecting")

                'unprotect, unhide and lock
                '------------------------------
                Try
                    xlsheet.Unprotect(format.x_factor)
                    xlsheet.Range("A1").Value2 = "z"        'this disables the worksheet change macro or reset macro
                    unhide_and_lock_all(format.detail_hdr_row_start, 100, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out
                Catch ex As Exception
                    err = "EXREJ: CR form is not genuine.  The CR form (" & Path.GetFileName(cr_form) & ") appears to be fake, please use the form that was emailed to you in the execution request'.<BR>Thanks"
                    GoTo get_out
                End Try
                'at this point, I know the form has not changed or been switched as it has passed the protection test

                Debug.WriteLine(Now.ToLongTimeString & ": clear extra cols")

                'this clear cells after the last detailed col ("executor comments") - needed to  deal with older forms
                '-------------------------------------------------------------------------------
                Try
                    Dim r1 As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + det_hdr_name.Count), xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + det_hdr_name.Count + 100))
                    r1.EntireColumn.Clear()
                Catch ex As Exception
                End Try

                'setup some ranges
                '--------------------------
                Dim comindex_ex_coord As Integer = Array.IndexOf(com_hdr_name, "Execution Coordinator")
                Dim comindex_ex As Integer = Array.IndexOf(com_hdr_name, "Executors")
                Dim comindex_ex_date As Integer = Array.IndexOf(com_hdr_name, "Execution Date")

                Dim index_cr_type As Integer = Array.IndexOf(det_hdr_name, "CR Type")
                Dim index_executor As Integer = Array.IndexOf(det_hdr_name, "Executor")
                Dim index_ex_status As Integer = Array.IndexOf(det_hdr_name, "Execution Status")
                Dim index_ex_date As Integer = Array.IndexOf(det_hdr_name, "Execution Date")
                Dim index_ex_comments As Integer = Array.IndexOf(det_hdr_name, "Executor Comments")
                Dim hdr_range As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, 1), xlsheet.Cells(format.detail_hdr_row_start, det_hdr_name.Count))

                'First we find the row limit for the cr form
                '----------------------------------------------
                Dim last_data_row As Integer = 0
                Try
                    With xlsheet
                        last_data_row = .Columns("A:A").offset(0, index_cr_type).entirecolumn.Find(What:="*", After:=.Cells(index_cr_type + 1), LookAt:=XlLookAt.xlWhole, LookIn:=XlFindLookIn.xlValues, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False).Row
                    End With
                Catch ex As Exception
                    last_data_row = 0
                End Try
                If last_data_row = 0 Then
                    err = "EXREJ: The CR form format has been corrupted, can't find the CR Type column...."
                    GoTo get_out
                End If

                Debug.WriteLine(Now.ToLongTimeString & ": check attachments")

                'check the attachments
                '----------------------------------------------------------------
                '----------------------------------------------------------------
                '----------------------------------------------------------------
                err = ""
                Dim cr_sub_id_col As Integer = find_col(hdr_range, "CR_sub_ID")
                Dim cr_type_col As Integer = find_col(hdr_range, "CR Type")
                Dim node_col As Integer = find_col(hdr_range, "Node")
                Dim comment_col As Integer = find_col(hdr_range, "Executor Comments")
                Dim status_col As Integer = find_col(hdr_range, "Execution Status")
                Dim comment_range As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, comment_col), xlsheet.Cells(last_data_row, comment_col))
                If LCase(cr_type) Like "parameter" Or LCase(cr_type) = "hardware" Then
                    Dim f_string As String = ""
                    If FileIO.FileSystem.DirectoryExists(inbox_path & "\executor attachments") Then     'it may not exist if there were no attachments
                        For Each item In IO.Directory.GetFiles(inbox_path & "\executor attachments", "*", IO.SearchOption.TopDirectoryOnly)
                            Dim file1 As String = LCase(Path.GetFileNameWithoutExtension(item))
                            If Not item Like cr_form AndAlso Not item Like "~$" & cr_form Then
                                f_string = f_string & Path.GetFileName(item) & ","
                            End If
                        Next
                        For Each item In IO.Directory.GetDirectories(inbox_path & "\executor attachments", "*", IO.SearchOption.TopDirectoryOnly)
                            Dim dir_info As New System.IO.DirectoryInfo(item)
                            Dim dir_s As String = dir_info.Name
                            f_string = f_string & dir_s & "\,"
                        Next
                    End If
                    If Not f_string Like "" Then
                        f_string = Left(f_string, Len(f_string) - 1)
                    End If

                    If f_string Like "" Then
                        err = "EXREJ: Must give at least 1 attachment to prove CR was done"
                    End If

                ElseIf LCase(cr_type) Like "rf basic" Or LCase(cr_type) = "rf re-engineering" Then
                    Try
                        'still do this in XL as for these cr types the total rows is not going to be more than 100
                        normal_format_cell_resub(comment_range)
                        normal_format_cell_resub(comment_range.Offset(0, 1))
                        For Each cell As Excel.Range In comment_range
                            Dim cr_sub_id As String = LCase(Trim(o2s(cell.Offset(0, cr_sub_id_col - comment_col).Value2)))
                            Dim comment As String = LCase(Trim(o2s(cell.Value2)))
                            Dim status As String = LCase(Trim(o2s(cell.Offset(0, status_col - comment_col).Value2)))
                            Dim node As String = text2regex(Regex.Replace(LCase(Trim(o2s(cell.Offset(0, node_col - comment_col).Value2))), "^.*_", ""))
                            Dim node2 As String = Regex.Replace(node, "[0-9]{1,2}$", "", RegexOptions.IgnoreCase)
                            If node = "" Then
                                GoTo skip
                            End If
                            Dim f_string As String = ""
                            Dim p1 As String = "^(.*[-_\. ])?"
                            Dim p2 As String = "([-_\. ].*)?$"
                            If FileIO.FileSystem.DirectoryExists(inbox_path & "\executor attachments") Then     'it may not exist if there were no attachments
                                For Each item In IO.Directory.GetFiles(inbox_path & "\executor attachments", "*", IO.SearchOption.TopDirectoryOnly)
                                    Dim file1 As String = LCase(Path.GetFileNameWithoutExtension(item))
                                    If Not item Like cr_form AndAlso Not item Like "~$" & cr_form AndAlso Regex.IsMatch(file1, p1 & "((" & node & ")|(" & node2 & "))" & p2) Then
                                        f_string = f_string & Path.GetFileName(item) & ","
                                    End If
                                Next
                                For Each item In IO.Directory.GetDirectories(inbox_path & "\executor attachments", "*", IO.SearchOption.TopDirectoryOnly)
                                    Dim dir_info As New System.IO.DirectoryInfo(item)
                                    Dim dir_s As String = dir_info.Name
                                    If Regex.IsMatch(dir_s, p1 & "((" & node & ")|(" & node2 & "))" & p2) Then
                                        f_string = f_string & dir_s & "\,"
                                    End If
                                Next
                            End If
                            If Not f_string Like "" Then
                                f_string = Left(f_string, Len(f_string) - 1)
                            End If

                            If f_string Like "" AndAlso status Like "Fail" AndAlso comment Like "" Then
                                err = "ERROR!! - Must give an attachment or comment if the status is 'Fail'."
                                'need to error format cell as it is outside of the detail range (after the ex comments)
                                error_format_cell(False, cell.Offset(0, 1), err)
                            ElseIf f_string Like "" AndAlso status Like "Executed" Then
                                err = "ERROR!! - Must give an attachment if the status is 'Executed'."
                                'need to error format cell as it is outside of the detail range (after the ex comments)
                                error_format_cell(False, cell.Offset(0, 1), err)
                            End If
skip:
                        Next
                    Catch ex As Exception
                        err = "ER: General error checking attachments, details: " & ex.ToString
                        GoTo get_out
                    End Try
                End If
                If Not err Like "" Then
                    attach_ok = "nok"
                    err = ""
                End If
                If attach_ok Like "not finished testing" Then attach_ok = "ok"

                'checks the cells that should have been filled are indeed filled and with values that are acceptable
                '--------------------------------------------------------------------------------------------------------
                'do this in XL as for this cr type the number of rows will not be > 100
                'do the 9 fields before the execution status
                '----------------------------------------
                err = ""
                If cr_form_type = "rfb" Then
                    Debug.WriteLine(Now.ToLongTimeString & ": checking ex values in XL for RFB")

                    'the critical thing here are the col offsets, so if you change the col order, you have to update it here, not so user freindly
                    '------------------------------------------
                    'NOTE: i CAN NOT USE RETURN_ERR HERE!!!!!!
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_status + 1 - 9), xlsheet.Cells(last_data_row, index_ex_status + 1 - 9))
                    check_az_values(temp_range, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then GoTo get_out
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_status + 1 - 8), xlsheet.Cells(last_data_row, index_ex_status + 1 - 8))
                    check_mdt_values(format, (8 - 2), temp_range, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then GoTo get_out
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_status + 1 - 7), xlsheet.Cells(last_data_row, index_ex_status + 1 - 7))
                    check_edt_values(format, (7 - 2), temp_range, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then GoTo get_out
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_status + 1 - 6), xlsheet.Cells(last_data_row, index_ex_status + 1 - 6))
                    check_az_values(temp_range, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then GoTo get_out
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_status + 1 - 5), xlsheet.Cells(last_data_row, index_ex_status + 1 - 5))
                    check_mdt_values(format, (5 - 2), temp_range, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then GoTo get_out
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_status + 1 - 4), xlsheet.Cells(last_data_row, index_ex_status + 1 - 4))
                    check_edt_values(format, (4 - 2), temp_range, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then GoTo get_out
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_status + 1 - 3), xlsheet.Cells(last_data_row, index_ex_status + 1 - 3))
                    check_ht_coax_len_values(temp_range, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then GoTo get_out
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_status + 1 - 2), xlsheet.Cells(last_data_row, index_ex_status + 1 - 2))
                    check_antenna_values(format, temp_range, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then GoTo get_out
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_status + 1 - 1), xlsheet.Cells(last_data_row, index_ex_status + 1 - 1))
                    check_ht_coax_len_values(temp_range, err)
                    If Regex.IsMatch(err, "^ER:", RegexOptions.IgnoreCase) Then GoTo get_out
                End If
                If Not err Like "" Then
                    data_ok = "nok"
                    err = ""
                End If

                Debug.WriteLine(Now.ToLongTimeString & ": read to DT")

                'reads the edited range into a dt for checking and updating the DB
                '--------------------------------------------------------------
                Try
                    Dim cr_sub_id_range As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, format.detail_col_start), xlsheet.Cells(last_data_row, format.detail_col_start))
                    Dim cols() As String = {}
                    If cr_form_type = "rfb" Then
                        temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_status + 1 - 10), xlsheet.Cells(last_data_row, index_ex_comments + 1))
                        cols = {"executor", "act_az", "act_mdt", "act_edt", "fin_az", "fin_mdt", "fin_edt", "fin_ht", "fin_antenna", "fin_coax_len", "execution_status", "execution_date", "executor_comments"}
                    Else
                        temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_status + 1 - 1), xlsheet.Cells(last_data_row, index_ex_comments + 1))
                        cols = {"executor", "execution_status", "execution_date", "executor_comments"}
                    End If
                    xlrange2dt(cols, cr_sub_id_range, temp_range, dt, format, err)
                Catch ex As Exception
                    err = "ER: Tool error reading xl to dt, checking ex returned values for review, details: " & ex.ToString
                    GoTo get_out
                End Try

                Debug.WriteLine(Now.ToLongTimeString & ": check ex")

                'check executors again as I have now allowed this to be edited by the executor themselves as maybe the executor was changed
                'but I only accept registed executors this time, no ex_coordinators
                '-------------------------------------------------------------------------
                err = ""
                Try
                    Dim qrows = From row In dt Where Not format.IsValidEmail(c2e(row.Field(Of String)("executor"))) Select row
                    For Each row In qrows
                        err = "ERROR!! - Must contain a valid email address'"
                        row("executor") = err & " - " & row.Field(Of String)("executor")
                    Next
                    qrows = From row In dt Where format.IsValidEmail(c2e(row.Field(Of String)("executor"))) Select row
                    Dim test_vals() As String = (From row In format.ds_allow.Tables("executors") Select row.Field(Of String)("combined_name")).Distinct.ToArray
                    qrows = From row In qrows Where Not test_vals.Any(Function(s) row.Field(Of String)("executor").Contains(s)) Select row
                    For Each row In qrows
                        err = "ERROR!! - Unknown Executor, you must enter a registered executor - to register an executor, please send mail: Subject => Add Executor and Body => name,email; => one executor per line"
                        row("executor") = err & " - " & row.Field(Of String)("executor")
                    Next
                Catch ex As Exception
                    err = "ER: Tool error checking executors, details: " & ex.ToString
                    GoTo get_out
                End Try
                If Not err Like "" Then
                    data_ok = "nok"
                    err = ""
                End If

                'we get a list of the executors and add the list to the common values
                '-----------------------------------------------
                executors_raw = ""
                Try
                    Dim qrows() As String = (From row In dt Let a = row.Field(Of String)("executor") Select a).Distinct.ToArray
                    executors_raw = Join(qrows, ",")
                Catch ex As Exception
                    err = "ER: Tool error getting executors, details: " & ex.ToString
                    GoTo get_out
                End Try
                'this checks and cleans the executors list => takes out doubles, yes even though I already did that and checks all values contain valid email addresses
                executors_raw = check_email_list(False, executors_raw, format)

                'this adds the executors list to the common vals
                '-----------------------------------
                err = ""
                temp_range = xlsheet.Range(xlsheet.Cells(format.common_row_start + comindex_ex, format.common_data_col), xlsheet.Cells(format.common_row_start + comindex_ex, format.common_data_col))
                If executors_raw Like "" Then
                    err = "ERROR!! - Invalid executors list, can not proceed, there is not one executor with a valid email"
                    'need to error format cell as it is in the common vals
                    error_format_cell(True, temp_range, err)
                Else
                    temp_range.Value2 = executors_raw
                End If
                If Not err Like "" Then
                    data_ok = "nok"
                    err = ""
                End If

                Debug.WriteLine(Now.ToLongTimeString & ": check ex status")

                'check execution status
                '--------------------------
                err = ""
                Try
                    Dim qrows = From row In dt Let a = row.Field(Of String)("execution_status") Where Not (a Like "Executed" Or a Like "Fail") Select row
                    For Each row In qrows
                        err = "ERROR!! - CR status must be either 'Executed' = it is done, or 'Fail' = it couldn't be done.  NOTE: for both cases you must supply attachments for review."
                        row("execution_status") = err & " - " & row.Field(Of String)("execution_status")
                    Next
                Catch ex As Exception
                    err = "ER: Tool error checking execution status, details: " & ex.ToString
                    GoTo get_out
                End Try
                If Not err Like "" Then
                    data_ok = "nok"
                    err = ""
                End If

                Debug.WriteLine(Now.ToLongTimeString & ": check ex date")

                'check execution date - we just make sure it is a date, that is all, the execution date we put in the common field is the date we get the cr form, as that is what we care about
                '------------------------
                err = ""
                Try
                    For Each row As System.Data.DataRow In dt.Rows
                        Dim ts1 As String = row.Field(Of String)("execution_date")
                        If ts1 Like "" Then
                            err = "ERROR!! - Must be a date"
                            row("execution_date") = err & " - " & ts1
                        Else
                            Dim ts2 As DateTime
                            Dim temp_err As String = ""
                            get_xl_date(ts1, ts2, temp_err)            'more robust sub to handle all types of weird date input
                            If Not temp_err Like "" Then
                                err = "ERROR!! - " & temp_err
                                row("execution_date") = err & " - " & ts1
                            Else
                                row("execution_date") = ts2.Date.ToOADate
                            End If
                        End If
                    Next
                Catch ex As Exception
                    err = "ER: Error checking execution date, details: " & ex.ToString
                    GoTo get_out
                End Try
                If Not err Like "" Then
                    data_ok = "nok"
                    err = ""
                End If
                If data_ok Like "not finished testing" Then data_ok = "ok"

                'if we got a format error, we make the error file to show the user and get out
                '----------------------------------------------------------------------------------
                If data_ok Like "nok" Then
                    Debug.WriteLine(Now.ToLongTimeString & ": data is nok so cleaning up")

                    'writes the details table with errors data back to the xlsheet for the error file (only do ex, ex_status and ex_date as the others where done directly in XL as they are onyl for RFB which are always small
                    '-----------------------------------------------------------------------
                    err = ""
                    If cr_form_type = "rfb" Then
                        ds2xl_1col(xlsheet, dt, 1, index_executor + 1, format.detail_data_row_start, err)
                        If Not err Like "" Then GoTo get_out
                        ds2xl_1col(xlsheet, dt, 11, index_ex_status + 1, format.detail_data_row_start, err)
                        If Not err Like "" Then GoTo get_out
                        ds2xl_1col(xlsheet, dt, 12, index_ex_date + 1, format.detail_data_row_start, err)
                        If Not err Like "" Then GoTo get_out
                    Else
                        ds2xl_1col(xlsheet, dt, 1, index_executor + 1, format.detail_data_row_start, err)
                        If Not err Like "" Then GoTo get_out
                        ds2xl_1col(xlsheet, dt, 2, index_ex_status + 1, format.detail_data_row_start, err)
                        If Not err Like "" Then GoTo get_out
                        ds2xl_1col(xlsheet, dt, 3, index_ex_date + 1, format.detail_data_row_start, err)
                        If Not err Like "" Then GoTo get_out
                    End If

                    'Reset the detail formats before doing the error formatting
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, 1), xlsheet.Cells(last_data_row, det_hdr_name.Count))
                    normal_format_cell_resub(temp_range)
                    'we only do the error formatting on small to med CRs as it involves filtering which gets sketchy on large ranges, XL will throw and behave badly
                    'for big crs we just write the error data 
                    set_all_errors_cell2red_unlock(last_data_row, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out

                    'sets date format for the execution date in the detail data
                    '---------------------------------------------------------
                    temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_date + 1), xlsheet.Cells(last_data_row, index_ex_date + 1))
                    date_format_cell(temp_range)

                    'hides fields again to the review format as I need to see attachments, they are all protected, this is ok
                    '-------------------------------------------------------------------------------
                    hide_rows_and_cols(format.row_hide_for_rev, {}, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out

                    'protects the sheet and book
                    '----------------------------------
                    xlsheet.Range("C2").Select()
                    xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True)

                    'makes the error file
                    '----------------------
                    Dim file_out As String = ""
                    file_out = inbox_path & "\" & Path.GetFileNameWithoutExtension(cr_form) & "_errors.xlsb"
                    xlapp.DisplayAlerts = False
                    xlbook.SaveAs(file_out)
                    err = "XREJ: See returned CR form (CRMS: CR Execution Retry Request (" & cr_id & "))(" & file_out & ")"
                    GoTo get_out
                End If

                Debug.WriteLine(Now.ToLongTimeString & ": writing data to XL")

                'sets the execution date in the common header, we actually just use now, basically the time we recieve and process the executor form, 
                'the detailed dates are in the details table in the DB
                '-------------------------------------------------------------------------------------------------------------------------------
                temp_range = xlsheet.Range(xlsheet.Cells(format.common_row_start + comindex_ex_date, format.common_data_col), xlsheet.Cells(format.common_row_start + comindex_ex_date, format.common_data_col))
                temp_range.Value2 = ex_date.ToOADate
                date_format_cell(temp_range)

                'writes the execution date back to XL as it may have changed after format checking in the datatable, all other cols do not need ot be written back if the sheet is ok
                '-----------------------------------------------------------------------------------------------
                If cr_form_type = "rfb" Then
                    ds2xl_1col(xlsheet, dt, 12, index_ex_date + 1, format.detail_data_row_start, err)
                    If Not err Like "" Then GoTo get_out
                Else
                    ds2xl_1col(xlsheet, dt, 3, index_ex_date + 1, format.detail_data_row_start, err)
                    If Not err Like "" Then GoTo get_out
                End If
                temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index_ex_date + 1), xlsheet.Cells(last_data_row, index_ex_date + 1))
                date_format_cell(temp_range)

                Debug.WriteLine(Now.ToLongTimeString & ": final formatting")

                'resets the formatting of the common and detail cells on the sheet
                '--------------------------------------------------------------------
                temp_range = xlsheet.Range(xlsheet.Cells(format.common_row_start, format.common_data_col), xlsheet.Cells(format.common_row_start + com_hdr_name.Count - 1, format.common_data_col))
                normal_format_cell_resub(temp_range)
                temp_range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, 1), xlsheet.Cells(last_data_row, det_hdr_name.Count))
                normal_format_cell_resub(temp_range)

                'sets the hide and unlock config for the next stage (ex) - no hiding or unlocking of cols for the reviewer
                '-------------------------------------------------------------------------------------------------
                hide_rows_and_cols(format.row_hide_for_rev, {}, xlsheet, format, err)
                If Not err Like "" Then GoTo get_out

                Debug.WriteLine(Now.ToLongTimeString & ": save and exit")

                'protect and save save the output file
                '------------------------------------
                xlsheet.Range("C2").Select()
                xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True)
                xlapp.DisplayAlerts = False
                xlbook.SaveAs(cr_form)

                'saves the error file if required
                '-------------------------------------
                If attach_ok = "nok" And Regex.IsMatch(cr_type, "^((RF\sRe-engineering)|(RF Basic))$", RegexOptions.IgnoreCase) Then
                    Debug.WriteLine(Now.ToLongTimeString & ": extra save and exit if have pending ex attachments")

                    xlsheet.Unprotect(format.x_factor)

                    Dim file_out As String = ""
                    file_out = inbox_path & "\" & Path.GetFileNameWithoutExtension(cr_form) & "_errors.xlsb"
                    If FileIO.FileSystem.FileExists(file_out) Then
                        force_delete_file(file_out, err)
                        If Not err Like "" Then GoTo get_out
                    End If

                    'lock everything and protect the sheet and book and save
                    '---------------------------------------------
                    xlsheet.Cells.Locked = True
                    xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True)
                    xlsheet.Range("C2").Select()
                    xlapp.DisplayAlerts = False
                    xlbook.SaveAs(file_out)
                End If
get_out:
            Catch ex As Exception
                err = "ER: General error checking the executor cr form '" & cr_form & "'.  Details: " & ex.ToString
            Finally
                Try
                    xlapp.DisplayAlerts = False
                    xlbook.Close()
                Finally
                    releaseObject(xlbook)
                    xlapp.UserControl = True
                    xlapp.Interactive = True
                    xlapp.IgnoreRemoteRequests = False
                    xlapp.Quit()
                    releaseObject(xlapp) 'this releases the com object
                End Try
            End Try
        Catch ex As Exception
            err = "ER: Error opening the XL application, details: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub










    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    'this creates the resubmit form from an existing cr_form.
    'it removes all added data and resets back no-fill and it changes the fields lock status so the requester can edit appropriate fields
    Public Sub create_resubmit_form(ByRef cr_resub_form As String, ByVal cr_form_type As String, ByVal format As cr_sheet_format, ByVal local As local_machine, ByRef err As String)
        Try
            'opens a new instance of XL
            '----------------------------------------------
            Dim xlapp As New Excel.Application
            xlapp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow
            xlapp.Visible = format.debug_xl
            xlapp.DisplayAlerts = False
            xlapp.UserControl = False
            xlapp.IgnoreRemoteRequests = True
            xlapp.Interactive = False
            Dim xlbook As Excel.Workbook = Nothing
            Dim cr_path As String = Path.GetDirectoryName(cr_resub_form)
            Dim temp_range As Excel.Range

            Try
                'find the com and detail header name arrays
                '--------------------------------------------
                Dim com_hdr_name() As String = {}
                com_hdr_name = format.common_hdr_name
                Dim det_hdr_name() As String = get_string_array_from_name("detail_hdr_name_" & cr_form_type, format, err)
                If Not err Like "" Then GoTo get_out

                'this opens the xl file
                '-------------------
                Dim cr_form_to_open As String = cr_resub_form
                Try
                    Debug.WriteLine(Now.ToLongTimeString & ": open XL")

                    'just in case as it is not in the inbox
                    pre_open_non_inbox_cr_file_conflict_resolution(cr_resub_form, err)
                    If Not err Like "" Then GoTo get_out

                    xlapp.DisplayAlerts = False
                    xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                Catch ex As COMException
                    If ex.HResult = -2146827284 Then
                        fix_bad_xlsb_file(cr_form_to_open, xlapp, format, local, err)
                        If Not err Like "" Then GoTo get_out
                        Try
                            xlapp.DisplayAlerts = False
                            xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                        Catch exx As Exception
                            err = "ER: Internal error opening xl file after fixing, details: " & exx.ToString
                            GoTo get_out
                        End Try
                    Else
                        err = "ER: Internal COM error opening xl file, details: " & ex.ToString
                        GoTo get_out
                    End If
                Catch ex As Exception
                    err = "ER: Internal error opening xl file, details: " & ex.ToString
                    GoTo get_out
                End Try

                'this checks the CR sheet exists
                '--------------------------
                Dim xlsheet As Excel.Worksheet
                Try
                    xlsheet = xlbook.Worksheets("CR")
                Catch ex As Exception
                    err = "EXREJ: CR form rejection.  The CR form (" & Path.GetFileName(cr_resub_form) & ") doesn't have a sheet called 'CR'.<BR>Thanks"
                    GoTo get_out
                End Try

                'first checks that the sheet is protected, if it is not, then someone has switched the sheet.
                '-------------------------------------------------------------------------------------------
                If Not (xlsheet.ProtectContents And xlsheet.ProtectScenarios) Then
                    err = "ER: Error trying to test the CR resubmit form when creating it, it appears to be not genuine (" & Path.GetFileName(cr_resub_form) & ")"
                    GoTo get_out
                End If

                Debug.WriteLine(Now.ToLongTimeString & ": unprotect")

                'unprotect, unhide and lock
                '------------------------------
                Try
                    xlsheet.Unprotect(format.x_factor)
                    xlsheet.Range("A1").Value2 = "z"        'this disables the worksheet change macro or reset macro
                    unhide_and_lock_all(format.detail_hdr_row_start, 100, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out
                Catch ex As Exception
                    err = "ER: some error unprotecting or unhiding and locking cells, form check sub"
                    GoTo get_out
                End Try
                'at this point, I know the form has not changed or been switched as it has passed the protection test

                'this clear cells after the last detailed col ("executor comments") - needed to  deal with older forms
                '-------------------------------------------------------------------------------
                temp_range = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + det_hdr_name.Count), xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + det_hdr_name.Count + 100))
                temp_range.EntireColumn.Clear()

                'First we find the row limit for the cr form
                '----------------------------------------------
                Dim index_cr_type As Integer = Array.IndexOf(det_hdr_name, "CR Type")
                Dim last_data_row As Integer = 0
                Try
                    With xlsheet
                        last_data_row = .Columns("A:A").offset(0, index_cr_type).entirecolumn.Find(What:="*", After:=.Cells(index_cr_type + 1), LookAt:=XlLookAt.xlWhole, LookIn:=XlFindLookIn.xlValues, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False).Row
                    End With
                Catch ex As Exception
                    last_data_row = 0
                End Try
                If last_data_row = 0 Then
                    err = "EXREJ: The CR form format has been corrupted, can't find the CR Type column...."
                    GoTo get_out
                End If

                'does the borders
                '--------------------
                Dim detail_rng = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, format.detail_col_start), xlsheet.Cells(1048576, det_hdr_name.Count))
                set_border_detail_data(detail_rng)

                'resets the formatting of the common and detail cells on the sheet
                '------------------------------------------------------------
                Dim t_rng As Excel.Range = xlsheet.Range(xlsheet.Cells(format.common_row_start, format.common_data_col), xlsheet.Cells(format.common_row_start + com_hdr_name.Count - 1, format.common_data_col))
                normal_format_cell_resub(t_rng)
                t_rng = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, 1), xlsheet.Cells(last_data_row, det_hdr_name.Count))
                normal_format_cell_resub(t_rng)

                'clears the common fields for ex_coord, executors as well as dates for open,approval,ex_planning,ex,close
                '-------------------------------------------------------------------------------
                Try
                    For Each item In format.row_clear_for_resub
                        t_rng = xlsheet.Range(xlsheet.Cells(item, format.common_data_col), xlsheet.Cells(item, format.common_data_col))
                        t_rng.MergeArea.ClearContents()
                    Next
                Catch ex As Exception
                    err = "ER: some error clearing common fields, details: " & ex.ToString
                    GoTo get_out
                End Try

                'clears the detail fields after Execution Coordinator
                '------------------------------------------------------
                Dim col_array() As Integer = get_integer_array_from_name("col_clear_hide_" & cr_form_type & "_for_resub", format, err)
                If Not err Like "" Then GoTo get_out
                Try
                    For Each col In col_array
                        t_rng = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, col), xlsheet.Cells(last_data_row, col))
                        t_rng.ClearContents()
                    Next
                Catch ex As Exception
                    err = "ER: some error clearing detailed fields, details: " & ex.ToString
                    GoTo get_out
                End Try

                'sets the hide and unlock config for the next stage (back to start)
                '-------------------------------------------------------------------
                col_array = get_integer_array_from_name("col_clear_hide_" & cr_form_type & "_for_resub", format, err)
                If Not err Like "" Then GoTo get_out
                hide_rows_and_cols(format.row_hide_for_resub, col_array, xlsheet, format, err)
                col_array = get_integer_array_from_name("col_unprotect_" & cr_form_type & "_for_resub", format, err)
                If Not err Like "" Then GoTo get_out
                unlock_and_input_format(format.row_unprotect_for_resub, col_array, 1048576, cr_form_type, True, xlsheet, format, err)

                'also need to copy the validation down to the bottom
                '-----------------------------------------------
                t_rng = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, 1), xlsheet.Cells(format.detail_data_row_start, det_hdr_name.Count))
                t_rng.Copy()
                t_rng = t_rng.Resize(t_rng.Rows.Count + 1048576 - format.detail_data_row_start, t_rng.Columns.Count)
                t_rng.PasteSpecial(XlPasteType.xlPasteValidation)

                'protects the sheet
                '----------------------------------
                xlsheet.Range("C2").Select()
                xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True)

                'save the file
                '--------------------
                xlapp.DisplayAlerts = False
                xlbook.SaveAs(cr_resub_form)
get_out:
            Catch ex As Exception
                err = "ER: General error creating resub file '" & cr_resub_form & "'.  Details: " & ex.ToString
            Finally
                Try
                    xlapp.DisplayAlerts = False
                    xlbook.Close()
                Finally
                    releaseObject(xlbook)
                    xlapp.UserControl = True
                    xlapp.Interactive = True
                    xlapp.IgnoreRemoteRequests = False
                    xlapp.Quit()
                    releaseObject(xlapp) 'this releases the com object
                End Try
            End Try
        Catch ex As Exception
            err = "ER: Error opening the XL application, details: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub










    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    Public Sub finalise_cr_form(ByRef cr_form As String, ByVal cr_form_type As String, ByVal closed_date As Date, ByVal format As cr_sheet_format, ByVal local As local_machine, ByRef err As String)
        Try
            'opens a new instance of XL
            '----------------------------------------------
            Dim xlapp As New Excel.Application
            xlapp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow
            xlapp.Visible = format.debug_xl
            xlapp.DisplayAlerts = False
            xlapp.UserControl = False
            xlapp.IgnoreRemoteRequests = True
            xlapp.Interactive = False
            Dim xlbook As Excel.Workbook = Nothing
            Dim cr_path As String = Path.GetDirectoryName(cr_form)

            Try
                'find the com and detail header name arrays
                '--------------------------------------------
                Dim com_hdr_name() As String = {}
                com_hdr_name = format.common_hdr_name
                Dim det_hdr_name() As String = get_string_array_from_name("detail_hdr_name_" & cr_form_type, format, err)
                If Not err Like "" Then GoTo get_out

                'this opens the xl file
                '-------------------
                Dim cr_form_to_open As String = cr_form
                Try
                    Debug.WriteLine(Now.ToLongTimeString & ": open XL")
                    xlapp.DisplayAlerts = False
                    xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                Catch ex As COMException
                    If ex.HResult = -2146827284 Then
                        fix_bad_xlsb_file(cr_form_to_open, xlapp, format, local, err)
                        If Not err Like "" Then GoTo get_out
                        Try
                            xlapp.DisplayAlerts = False
                            xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                        Catch exx As Exception
                            err = "ER: Internal error opening xl file after fixing, details: " & exx.ToString
                            GoTo get_out
                        End Try
                    Else
                        err = "ER: Internal COM error opening xl file, details: " & ex.ToString
                        GoTo get_out
                    End If
                Catch ex As Exception
                    err = "ER: Internal error opening xl file, details: " & ex.ToString
                    GoTo get_out
                End Try

                'this checks the CR sheet exists
                '--------------------------
                Dim xlsheet As Excel.Worksheet
                Try
                    xlsheet = xlbook.Worksheets("CR")
                Catch ex As Exception
                    err = "FINALREJ: CR form rejection.  The CR form (" & Path.GetFileName(cr_form) & ") doesn't have a sheet called 'CR'.<BR>Thanks"
                    GoTo get_out
                End Try

                'first checks that the sheet is protected, if it is not, then someone has switched the sheet.
                '-------------------------------------------------------------------------------------------
                If Not (xlsheet.ProtectContents And xlsheet.ProtectScenarios) Then
                    err = "ER: Error trying to test the CR resubmit form when creating it, it appears to be not genuine (" & Path.GetFileName(cr_form) & ")"
                    GoTo get_out
                End If


                'unprotect, unhide and lock
                '------------------------------
                Try
                    xlsheet.Unprotect(format.x_factor)
                    xlsheet.Range("A1").Value2 = "z"        'this disables the worksheet change macro or reset macro
                    unhide_and_lock_all(format.detail_hdr_row_start, 100, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out
                Catch ex As Exception
                    err = "ER: some error unprotecting or unhiding and locking cells, form check sub"
                    GoTo get_out
                End Try
                'at this point, I know the form has not changed or been switched as it has passed the protection test

                'sets the closed date
                '-----------------------
                Dim comindex_closed_date As Integer = Array.IndexOf(com_hdr_name, "Closed Date")
                Dim xlcell As Excel.Range = xlsheet.Range(xlsheet.Cells(format.common_row_start + comindex_closed_date, format.common_data_col), xlsheet.Cells(format.common_row_start + comindex_closed_date, format.common_data_col))
                xlcell.Value2 = closed_date.ToOADate
                date_format_cell(xlcell)

                'protects the sheet
                '----------------------------------
                xlsheet.Range("C2").Select()
                xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True)

                'save the file
                '--------------------
                xlapp.DisplayAlerts = False
                xlbook.SaveAs(cr_form)
get_out:
            Catch ex As Exception
                err = "ER: General error checking the cr file '" & cr_form & "'.  Details: " & ex.ToString
            Finally
                Try
                    xlapp.DisplayAlerts = False
                    xlbook.Close()
                Finally
                    releaseObject(xlbook)
                    xlapp.UserControl = True
                    xlapp.Interactive = True
                    xlapp.IgnoreRemoteRequests = False
                    xlapp.Quit()
                    releaseObject(xlapp) 'this releases the com object
                End Try
            End Try
        Catch ex As Exception
            err = "ER: Error opening excel, details: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub




    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    '###############################################################################################################################
    'this updates the allowed values in the given cr form
    'I do not support file fixing here, if I can't open the file, then udah
    Public Sub allowed_values_ds2xl(ByVal file_in As String, ByVal cr_form_type As String, ByVal format As cr_sheet_format, ByVal local As local_machine, ByRef err As String)
        Try
            'opens a new instance of XL
            '----------------------------------------------
            Dim xlapp As New Excel.Application
            xlapp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow
            xlapp.Visible = format.debug_xl
            xlapp.DisplayAlerts = False
            xlapp.UserControl = False
            xlapp.IgnoreRemoteRequests = True
            xlapp.Interactive = False
            Dim xlbook As Excel.Workbook = Nothing

            Try
                'find the com and detail header name arrays
                '--------------------------------------------
                Dim com_hdr_name() As String = {}
                com_hdr_name = format.common_hdr_name
                Dim det_hdr_name() As String = {}
                If cr_form_type = "" Then
                    'do nothing
                ElseIf cr_form_type Like "init" Then
                    det_hdr_name = get_string_array_from_name("detail_hdr_name", format, err)
                    If Not err Like "" Then GoTo get_out
                Else
                    det_hdr_name = get_string_array_from_name("detail_hdr_name_" & cr_form_type, format, err)
                    If Not err Like "" Then GoTo get_out
                End If

                'this opens the xl file
                '-------------------
                Dim cr_form_to_open As String = file_in
                Try
                    Debug.WriteLine(Now.ToLongTimeString & ": open XL")

                    'sometimes this file could be open as it is not from the inbox, so if it is, kill it
                    pre_open_non_inbox_cr_file_conflict_resolution(file_in, err)
                    If Not err Like "" Then GoTo get_out

                    xlapp.DisplayAlerts = False
                    xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                Catch ex As Exception
                    err = "ER: Internal error opening xl file, details: " & ex.ToString
                    GoTo get_out
                End Try

                'this checks the allowed values sheet exists
                '--------------------------
                Dim xlsheet As Excel.Worksheet
                Try
                    xlsheet = xlbook.Worksheets("Allowed Values")
                Catch ex As Exception
                    err = "UPDATEREJ: The CR form (" & Path.GetFileName(file_in) & ") is missing the allowed values sheet.<BR>Thanks"
                    GoTo get_out
                End Try

                'first checks that the sheet is protected, if it is not, then someone has switched the sheet.
                '-------------------------------------------------------------------------------------------
                If Not (xlsheet.ProtectContents And xlsheet.ProtectScenarios) Then
                    err = "ER: It appears the CR form has been tampered with (" & Path.GetFileName(file_in) & ")"
                    GoTo get_out
                End If

                'unprotect, unhide and lock
                '------------------------------
                Try
                    xlsheet.Unprotect(format.x_factor)
                    If Not err Like "" Then GoTo get_out
                Catch ex As Exception
                    err = "ER: some error unprotecting or unhiding and locking cells, form check sub"
                    GoTo get_out
                End Try
                'at this point, I know the form has not changed or been switched as it has passed the protection test

                'Here we update the allowed values
                '-------------------------------------
                Dim hdr_rng As Excel.Range = xlsheet.Range(xlsheet.Cells(1, 1), xlsheet.Cells(1, 100))
                For Each dt As System.Data.DataTable In format.ds_allow.Tables
                    If Regex.IsMatch(dt.TableName, "^(state_control)|(administrators)$", RegexOptions.IgnoreCase) Then GoTo skip

                    Dim col As Integer = find_col(hdr_rng, dt.TableName)
                    If col > 0 Then
                        'reset vals
                        '----------------
                        Dim val_rng As Excel.Range = xlsheet.Range(xlsheet.Cells(2, col), xlsheet.Cells(1048576, col))
                        val_rng.ClearContents()

                        'write new vals
                        '----------------
                        Dim read_col As Integer = 0
                        Dim view As New DataView(dt)
                        view.Sort = dt.Columns(0).ColumnName
                        If Regex.IsMatch(dt.TableName, "(^(requesters)|(approvers)|(executors))|(_ex_coord)$", RegexOptions.IgnoreCase) Then
                            view.Sort = dt.Columns("combined_name").ColumnName
                            read_col = 2
                        End If
                        Dim dt_new As System.Data.DataTable = view.ToTable()
                        ds2xl_1col(xlsheet, dt_new, read_col, col, 2, err)
                        If Not err Like "" Then GoTo get_out
                    End If
skip:
                Next

                'protects the sheet
                '----------------------------------
                'xlsheet.Range("A1").Select()   this throws an error for some reason, but only sometimes
                xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True)

                'save the file
                '--------------------
                xlapp.DisplayAlerts = False
                xlbook.SaveAs(file_in)
get_out:
            Catch ex As Exception
                err = "ER: General error checking the cr file '" & file_in & "'.  Details: " & ex.ToString
            Finally
                Try
                    xlapp.DisplayAlerts = False
                    xlbook.Close()
                Finally
                    releaseObject(xlbook)
                    xlapp.UserControl = True
                    xlapp.Interactive = True
                    xlapp.IgnoreRemoteRequests = False
                    xlapp.Quit()
                    releaseObject(xlapp) 'this releases the com object
                End Try
            End Try
        Catch ex As Exception
            err = "ER: Error opening excel, details: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub






    'this modifies the colour and value of the cell that was in error
    '---------------------------------------------------------------------
    Public Sub error_format_cell(ByVal header_flag As Boolean, ByRef xlcell As Excel.Range, ByVal msg As String)
        On Error Resume Next
        With xlcell
            .Interior.ColorIndex = 3
            .Interior.Pattern = Excel.XlPattern.xlPatternSolid
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.ColorIndex = 2
            .Font.Bold = True
            .Value2 = msg
            If header_flag Then
                .MergeArea.Locked = False
            Else
                .Locked = False
            End If
        End With
    End Sub



    Public Sub error_format_range(ByRef rng As Excel.Range)
        On Error Resume Next
        With rng
            .Interior.ColorIndex = 3
            .Interior.Pattern = Excel.XlPattern.xlPatternSolid
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.ColorIndex = 2
            .Font.Bold = True
        End With
    End Sub



    'this modifies the colour and value of a normal cell
    '-----------------------------------------------------
    Public Sub normal_format_cell_w_msg(ByRef xlcell As Excel.Range, ByVal msg As String)
        On Error Resume Next
        With xlcell
            .Interior.Pattern = Excel.XlPattern.xlPatternNone
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.Strikethrough = False
            .Font.Superscript = False
            .Font.Subscript = False
            .Font.OutlineFont = False
            .Font.Shadow = False
            .Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
            .Font.ThemeFont = XlThemeFont.xlThemeFontNone
            .Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic
            .Font.Bold = True
            .Font.TintAndShade = 0
            .Value2 = msg
        End With
    End Sub


    'this modifies the colour and value of a normal cell
    '-----------------------------------------------------
    Public Sub normal_format_cell(ByRef xlcell As Excel.Range)
        On Error Resume Next
        With xlcell
            .Interior.Pattern = Excel.XlPattern.xlPatternNone
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.Strikethrough = False
            .Font.Superscript = False
            .Font.Subscript = False
            .Font.OutlineFont = False
            .Font.Shadow = False
            .Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
            .Font.ThemeFont = XlThemeFont.xlThemeFontNone
            .Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic
            .Font.Bold = True
            .Font.TintAndShade = 0
        End With
    End Sub



    'this modifies the colour and value of a normal cell
    '-----------------------------------------------------
    Public Sub normal_format_cell_resub(ByRef xlcell As Excel.Range)
        On Error Resume Next
        With xlcell
            .Interior.Pattern = Excel.XlPattern.xlPatternNone
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic
            .Font.Bold = False
            .Font.TintAndShade = 0
        End With
    End Sub


    'set cell to input format => yellow
    '-------------------------------------
    Public Sub input_format_cell_resub(ByRef xlcell As Excel.Range)
        On Error Resume Next
        With xlcell
            .Interior.Pattern = XlPattern.xlPatternSolid
            .Interior.PatternColorIndex = XlPattern.xlPatternAutomatic
            .Interior.Color = 10092543
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic
            .Font.Bold = False
            .Font.TintAndShade = 0
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
            .WrapText = False
        End With
    End Sub



    'makes a cell use a date format
    '--------------------------------
    Public Sub date_format_cell(ByRef xlcell As Excel.Range)
        On Error Resume Next
        With xlcell
            .NumberFormat = "[$-409]d/mmm/yy;@"
        End With

        'this converts any numbers stored as text to numbers
        '-----------------------------------------
        xlcell.Value = xlcell.Value
    End Sub




    'find col from a range of cells
    '-----------------------------------------
    Public Function find_col(ByVal range As Excel.Range, ByVal test_string As String) As Integer
        On Error Resume Next
        find_col = -1
        find_col = range.Find(What:=test_string, After:=range.Cells(1), LookAt:=XlLookAt.xlWhole, LookIn:=XlFindLookIn.xlValues, SearchOrder:=XlSearchOrder.xlByColumns, SearchDirection:=XlSearchDirection.xlNext, MatchCase:=False).Column
    End Function




    'does the detail borders
    '------------------------
    Public Sub set_border_detail_data(ByRef com_range As Excel.Range)
        On Error Resume Next
        With com_range
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
            .WrapText = False
        End With
        With com_range.Borders(Excel.XlBordersIndex.xlEdgeLeft)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .TintAndShade = -0.249946592608417
        End With
        With com_range.Borders(Excel.XlBordersIndex.xlEdgeTop)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .TintAndShade = -0.249946592608417
        End With
        With com_range.Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .TintAndShade = -0.249946592608417
        End With
        With com_range.Borders(Excel.XlBordersIndex.xlEdgeRight)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .TintAndShade = -0.249946592608417
        End With
        With com_range.Borders(Excel.XlBordersIndex.xlInsideHorizontal)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .TintAndShade = -0.249946592608417
        End With
        With com_range.Borders(Excel.XlBordersIndex.xlInsideVertical)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .TintAndShade = -0.249946592608417
        End With
    End Sub




    'does the common borders
    '------------------------
    Public Sub set_border_common_data(ByRef com_range As Excel.Range)
        On Error Resume Next
        With com_range
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
            .WrapText = False
        End With
        With com_range.Borders(Excel.XlBordersIndex.xlEdgeLeft)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .TintAndShade = -0.249946592608417
        End With
        With com_range.Borders(Excel.XlBordersIndex.xlEdgeTop)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .TintAndShade = -0.249946592608417
        End With
        With com_range.Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .TintAndShade = -0.249946592608417
        End With
        With com_range.Borders(Excel.XlBordersIndex.xlEdgeRight)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .TintAndShade = -0.249946592608417
        End With
        With com_range.Borders(Excel.XlBordersIndex.xlInsideHorizontal)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .TintAndShade = -0.249946592608417
        End With
    End Sub


    'does the colours
    '--------------------
    Public Sub set_color_common_data(ByRef com_Range As Excel.Range, ByVal color As String)
        On Error Resume Next
        With com_Range.Font
            If color <> "grey" Then
                .Bold = False
                .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic
                .TintAndShade = 0
            Else
                .Bold = False
                .ThemeColor = XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0.499984740745262
            End If
        End With
        With com_Range.Interior
            If color = "green" Then
                .Pattern = XlPattern.xlPatternSolid
                .PatternColorIndex = XlPattern.xlPatternAutomatic
                .ThemeColor = XlThemeColor.xlThemeColorAccent3
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            ElseIf color = "blue" Then
                .Pattern = XlPattern.xlPatternSolid
                .PatternColorIndex = XlPattern.xlPatternAutomatic
                .ThemeColor = XlThemeColor.xlThemeColorAccent1
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            ElseIf color = "grey" Then
                .Pattern = XlPattern.xlPatternSolid
                .PatternColorIndex = XlPattern.xlPatternAutomatic
                .ThemeColor = XlThemeColor.xlThemeColorDark1
                .TintAndShade = -0.149998474074526
                .PatternTintAndShade = 0
            ElseIf color = "white" Then
                .Pattern = XlPattern.xlPatternNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            ElseIf color = "yellow" Then
                .Pattern = XlPattern.xlPatternSolid
                .PatternColorIndex = XlPattern.xlPatternAutomatic
                .Color = 10092543
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End If
        End With
    End Sub




    'This determines if a value is in the given table and field
    '---------------------------------------------------------------
    Public Function is_allowed_val(ByVal test_string As String, ByVal table_name As String, ByVal col_name As String, ByVal format As cr_sheet_format) As Boolean
        Try
            Dim qrow = From row In format.ds_allow.Tables(table_name)
                        Where row.Field(Of String)(col_name) Like test_string
                        Select row
            If qrow.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function





    'This loads the common and detailed datatable into memory for faster processing and also to transfer to the DB 
    'it does the detailed one in chunks to avoid the XL interop memory error
    '------------------------------------------------------------
    Public Sub load_xl2ds(ByVal com_hdr_cnt As Integer, ByVal last_data_row As Integer, ByRef ds As System.Data.DataSet, ByVal xlsheet As Excel.Worksheet, ByVal type As String, ByVal format As cr_sheet_format, ByRef err As String)
        Try
            'load common table first
            '-------------------------
            Dim tablename As String = ""
            If ds.Tables("com") Is Nothing Then tablename = type & "_com" Else tablename = "com"
            With ds.Tables(tablename)
                Dim r1 = xlsheet.Range(xlsheet.Cells(format.common_row_start, format.common_hdr_col), xlsheet.Cells(format.common_row_start + com_hdr_cnt - 1, format.common_hdr_col))
                'create the columns
                '----------------------
                For Each cell As Excel.Range In r1
                    Dim t_string As String = Trim(o2s(cell.Value2))
                    t_string = Regex.Replace(Strings.LCase(t_string), "\s", "_")
                    If .Columns(t_string) Is Nothing Then
                        .Columns.Add(t_string, System.Type.GetType("System.String"))
                        .Columns.Item(t_string).AllowDBNull = False
                        .Columns.Item(t_string).DefaultValue = ""
                    End If
                Next
                r1 = xlsheet.Range(xlsheet.Cells(format.common_row_start, format.common_data_col), xlsheet.Cells(format.common_row_start + com_hdr_cnt - 1, format.common_data_col))
                Dim t_array(,) As Object = r1.Value2
                Dim a(t_array.GetUpperBound(0) - 1) As String
                For i = 1 To t_array.GetUpperBound(0)
                    a(i - 1) = t_array(i, 1)
                Next
                .Rows.Add(a)
            End With
        Catch ex As Exception
            err = "ER: Error filling common datatables from xl"
            Exit Sub
        End Try

        Try
            'load data table next
            '-------------------------
            Dim r1 As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start), xlsheet.Cells(format.detail_hdr_row_start, 100))
            Dim last_data_col As Integer = r1.Find(What:="Executor Comments", After:=r1.Cells(1), LookAt:=XlLookAt.xlWhole, LookIn:=XlFindLookIn.xlValues, SearchOrder:=XlSearchOrder.xlByColumns, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False).Column
            r1 = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start), xlsheet.Cells(format.detail_hdr_row_start, last_data_col))
            Dim t_name() As String = {"cur", "pro", "act", "fin"}
            Dim cnt As Integer = 0
            Dim tablename As String = ""
            If ds.Tables("det") Is Nothing Then tablename = type & "_data" Else tablename = "det"
            With ds.Tables(tablename)
                'create the columns
                '----------------------
                For Each cell As Excel.Range In r1
                    Dim t_string As String = Trim(o2s(cell.Value2))
                    'converts the col name from the xl sheet name to the DB name
                    '-------------------------------------------------------------
                    t_string = Regex.Replace(Strings.LCase(t_string), "\s", "_")
                    If Regex.IsMatch(t_string, "^(az)|(mdt)|(edt)$") Then
                        t_string = t_name(cnt) & "_" & t_string
                        If Regex.IsMatch(t_string, "edt$") Then
                            cnt = cnt + 1
                        End If
                    ElseIf Regex.IsMatch(t_string, "^(ht)|(antenna)|(coax_len)$") Then
                        t_string = "fin_" & t_string
                    End If

                    'creates the col
                    '-----------------------
                    If .Columns(t_string) Is Nothing Then
                        .Columns.Add(t_string, System.Type.GetType("System.String"))
                        .Columns.Item(t_string).DefaultValue = ""
                    End If
                Next

                'does the data rows by chunk
                '-----------------------------
                Dim max_cells_per_chunk As Integer = format.max_cells_per_chunk
                Dim total_cols As Integer = last_data_col - format.detail_col_start + 1
                Dim row_chunk As Integer = Int(max_cells_per_chunk / total_cols)
                Dim chunk_cnt As Integer = 0
                Do
                    r1 = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start + chunk_cnt * row_chunk, format.detail_col_start), xlsheet.Cells(Math.Min(last_data_row, format.detail_data_row_start + (chunk_cnt + 1) * row_chunk - 1), last_data_col))
                    Dim t_array(,) As Object = r1.Value2
                    Dim i, j As Integer
                    For i = 1 To t_array.GetUpperBound(0)
                        Dim a(t_array.GetUpperBound(1) - 1) As String
                        For j = 1 To t_array.GetUpperBound(1)
                            a(j - 1) = t_array(i, j)
                        Next
                        .Rows.Add(a)
                    Next
                    chunk_cnt += 1
                Loop While last_data_row >= format.detail_data_row_start + chunk_cnt * row_chunk
            End With
        Catch ex As Exception
            err = "ER: Error filling detail datatables from xl, details: " & ex.ToString
        End Try
    End Sub





    'generic read an excel range to a datatable with the column names given in the call
    '-------------------------------------------------------------------
    Public Sub xlrange2dt(ByVal col() As String, ByVal r1 As Excel.Range, ByVal r2 As Excel.Range, ByRef dt As System.Data.DataTable, ByVal format As cr_sheet_format, ByRef err As String)
        Try
            If r2.Columns.Count <> col.Count Then
                err = "ER: number of cols doesn't match the input range"
                GoTo get_out
            End If

            With dt
                'add the cols
                '----------------------
                If .Columns("cr_sub_id") Is Nothing Then
                    .Columns.Add("cr_sub_id", System.Type.GetType("System.String"))
                    .Columns.Item("cr_sub_id").DefaultValue = ""
                End If
                For Each item In col
                    If .Columns(item) Is Nothing Then
                        .Columns.Add(item, System.Type.GetType("System.String"))
                        .Columns.Item(item).DefaultValue = ""
                    End If
                Next

                'add and fill the rows using chunks
                '---------------------------------------
                Dim max_cells_per_chunk As Integer = format.max_cells_per_chunk
                Dim total_cols As Integer = r1.Columns.Count + r2.Columns.Count
                Dim row_chunk As Integer = Int(max_cells_per_chunk / total_cols)
                Dim chunk_cnt As Integer = 0

                'decompose the incoming 2 ranges into the row and col limits
                '------------------------------------------------------------
                Dim t_rng As Excel.Range
                t_rng = r1.Columns(1)
                Dim r1_first_data_col As Integer = t_rng.Column
                t_rng = r1.Columns(r1.Columns.Count)
                Dim r1_last_data_col As Integer = t_rng.Column
                t_rng = r1.Rows(1)
                Dim r1_first_data_row As Integer = t_rng.Row
                t_rng = r1.Rows(r1.Rows.Count)
                Dim r1_last_data_row As Integer = t_rng.Row
                t_rng = r2.Columns(1)
                Dim r2_first_data_col As Integer = t_rng.Column
                t_rng = r2.Columns(r2.Columns.Count)
                Dim r2_last_data_col As Integer = t_rng.Column
                t_rng = r2.Rows(1)
                Dim r2_first_data_row As Integer = t_rng.Row
                t_rng = r2.Rows(r2.Rows.Count)
                Dim r2_last_data_row As Integer = t_rng.Row
                If Not r1_first_data_col = r1_last_data_col Then
                    err = "ER: r1 can only have 1 column"
                End If
                If r1_first_data_col >= r2_first_data_col And r1_first_data_col <= r2_last_data_col Then
                    err = "ER: r1 col can not be part of r2 cols"
                End If
                If Not r1_first_data_row = r2_first_data_row And Not r1_last_data_row = r2_last_data_row Then
                    err = "ER: r1 and r2 must have the same rows"
                End If
                If Not err Like "" Then GoTo get_out

                Dim xlsheet As Excel.Worksheet = r2.Worksheet
                Dim rng_chunk, rng_chunk_hdr As Excel.Range
                Do
                    rng_chunk_hdr = xlsheet.Range(xlsheet.Cells(r1_first_data_row + chunk_cnt * row_chunk, r1_first_data_col), xlsheet.Cells(Math.Min(r1_last_data_row, r1_first_data_row + (chunk_cnt + 1) * row_chunk - 1), r1_last_data_col))
                    rng_chunk = xlsheet.Range(xlsheet.Cells(r2_first_data_row + chunk_cnt * row_chunk, r2_first_data_col), xlsheet.Cells(Math.Min(r2_last_data_row, r2_first_data_row + (chunk_cnt + 1) * row_chunk - 1), r2_last_data_col))

                    Dim t_array_hdr(,) As Object
                    'note due to a massive bug with XL, when we read a range into a 2D array, it looks like it has gone into a zero based array,, i.e. array(0,0) is the to left cell, but actually internally to call the top left cell, we have to call array(1,1), i.e. it is a 1 based array, but displays as a ero based array
                    'this is a known issue and if you follow it, everything works fine.
                    If rng_chunk_hdr.Cells.Count >= 2 Then
                        t_array_hdr = rng_chunk_hdr.Value2
                    Else
                        t_array_hdr = {{"", ""}, {"", rng_chunk_hdr.Value2}}
                    End If

                    Dim t_array(,) As Object
                    'note due to a massive bug with XL, when we read a range into a 2D array, it looks like it has gone into a zero based array,, i.e. array(0,0) is the to left cell, but actually internally to call the top left cell, we have to call array(1,1), i.e. it is a 1 based array, but displays as a ero based array
                    'this is a known issue and if you follow it, everything works fine.
                    If rng_chunk.Cells.Count >= 2 Then
                        t_array = rng_chunk.Value2
                    Else
                        t_array = {{"", ""}, {"", rng_chunk.Value2}}
                    End If

                    Dim i, j As Integer
                    For i = 1 To t_array.GetUpperBound(0)   'the rows
                        Dim a(t_array.GetUpperBound(1)) As String
                        a(0) = t_array_hdr(i, 1)
                        For j = 1 To t_array.GetUpperBound(1)   'the cols
                            a(j) = t_array(i, j)
                        Next
                        .Rows.Add(a)
                    Next
                    chunk_cnt += 1
                Loop While r2_last_data_row >= r2_first_data_row + chunk_cnt * row_chunk
            End With
get_out:
        Catch ex As Exception
            err = "ER: Error filling datatable, details: " & ex.ToString
        End Try
    End Sub





    Public Sub set_all_errors_cell2red_unlock(ByVal last_data_row As Integer, ByVal xlsheet As Excel.Worksheet, ByVal format As cr_sheet_format, ByRef err As String)
        Try
            Dim r1 As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start), xlsheet.Cells(format.detail_hdr_row_start, 100))
            Dim last_data_col As Integer = r1.Find(What:="Executor Comments", After:=r1.Cells(1), LookAt:=XlLookAt.xlWhole, LookIn:=XlFindLookIn.xlValues, SearchOrder:=XlSearchOrder.xlByColumns, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False).Column
            If last_data_row > 2000 Then
                r1 = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, format.detail_col_start), xlsheet.Cells(last_data_row, last_data_col))
                For Each col As Excel.Range In r1.Columns
                    Dim test As Excel.Range = col.Find(What:="ERROR!! *", After:=col.Cells(1), LookAt:=XlLookAt.xlWhole, LookIn:=XlFindLookIn.xlValues, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext, MatchCase:=False)
                    If Not test Is Nothing AndAlso test.Row > 0 Then
                        col.Locked = False
                    End If
                Next

            Else
                Dim r2 = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start), xlsheet.Cells(format.detail_data_row_start - 1, last_data_col))
                Dim r2_temp = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, last_data_col + 20), xlsheet.Cells(format.detail_data_row_start - 1, last_data_col + (last_data_col + 20 - format.detail_col_start)))
                r2.Copy(r2_temp)
                r1 = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start - 1, format.detail_col_start), xlsheet.Cells(last_data_row, last_data_col))
                For Each col As Excel.Range In r1.Columns
                    col.AutoFilter(Field:=1, Criteria1:="=ERROR!! *", Operator:=Excel.XlAutoFilterOperator.xlFilterValues)
                    Dim test As Integer = col.SpecialCells(XlCellType.xlCellTypeVisible).Cells.Count
                    If test > 1 Then
                        error_format_range(col.SpecialCells(XlCellType.xlCellTypeVisible))
                    End If
                    xlsheet.AutoFilterMode = False
                    If test > 1 Then
                        col.Offset(1, 0).Resize(col.Rows.Count - 1, col.Columns.Count).Locked = False
                    End If
                Next
                r2_temp.Copy(r2)
                r2_temp.Clear()
            End If
get_out:
        Catch ex As Exception
            err = "ER: general setting error cells and unlocking, details: " & ex.ToString
        End Try
    End Sub





    'writes a datatable to an excel sheet, doesn't care about anything, just writes starting from the (row_start, col_start)
    Public Sub dt2xlsheet_generic(ByVal xlvalue2_flag As Boolean, ByVal row_start As Integer, ByVal col_start As Integer, ByVal dt As System.Data.DataTable, ByVal xlsheet As Excel.Worksheet, ByVal format As cr_sheet_format, ByRef err As String)
        Dim i, j As Integer
        Try
            Dim last_data_row As Integer = dt.Rows.Count + row_start - 1
            Dim last_data_col As Integer = dt.Columns.Count + col_start - 1
            Dim r1 As Excel.Range = xlsheet.Range(xlsheet.Cells(row_start, col_start), xlsheet.Cells(last_data_row, last_data_col))
            Dim cnt As Integer = 0
            Dim tablename As String = ""
            With dt
                'does the data rows by chunk
                '-----------------------------
                Dim max_cells_per_chunk As Integer = format.max_cells_per_chunk
                Dim total_cols As Integer = dt.Columns.Count
                Dim row_chunk As Integer = Int(max_cells_per_chunk / total_cols)
                Dim chunk_cnt As Integer = 0
                Do
                    r1 = xlsheet.Range(xlsheet.Cells(row_start + chunk_cnt * row_chunk, col_start), xlsheet.Cells(Math.Min(last_data_row, row_start + (chunk_cnt + 1) * row_chunk - 1), last_data_col))
                    Dim rowcnt As Integer = Math.Min(last_data_row, row_start + (chunk_cnt + 1) * row_chunk - 1) - (row_start + chunk_cnt * row_chunk) + 1
                    Dim colcnt As Integer = .Columns.Count
                    Dim t_array(rowcnt - 1, colcnt - 1) As String
                    For i = 0 To rowcnt - 1
                        For j = 0 To colcnt - 1
                            t_array(i, j) = .Rows(chunk_cnt * row_chunk + i).Field(Of String)(j)
                        Next
                    Next
                    Try
                        If xlvalue2_flag Then
                            r1.Value2 = t_array
                        Else
                            r1.Value = t_array
                        End If
                    Catch ex As Exception
                        err = "ER: error writing data to XL, details: " & ex.ToString
                        GoTo get_out
                    End Try
                    chunk_cnt += 1
                Loop While last_data_row >= row_start + chunk_cnt * row_chunk
            End With
get_out:
        Catch ex As Exception
            err = "ER: general error writing to XL generic: " & ex.ToString
        End Try
    End Sub






    'puts the ds back into the cr form => only writes the data table right now, the common data is done directly in XL
    Public Sub dt2xlrange(ByVal last_data_row As Integer, ByRef dt As System.Data.DataTable, ByVal xlsheet As Excel.Worksheet, ByVal format As cr_sheet_format, ByRef err As String)
        Dim i, j As Integer
        Try
            Dim r1 As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start), xlsheet.Cells(format.detail_hdr_row_start, 100))
            Dim last_data_col As Integer = r1.Find(What:="Executor Comments", After:=r1.Cells(1), LookAt:=XlLookAt.xlWhole, LookIn:=XlFindLookIn.xlValues, SearchOrder:=XlSearchOrder.xlByColumns, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False).Column
            r1 = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start), xlsheet.Cells(format.detail_hdr_row_start, last_data_col))
            Dim cnt As Integer = 0
            Dim tablename As String = ""
            With dt
                'tests the column count only, doesn't check the col names
                '--------------------------------------------------------
                If Not r1.Columns.Count = .Columns.Count And Not r1.Rows.Count = .Rows.Count Then
                    err = "ER: unmatched xl range and datatable error"
                    GoTo get_out
                End If

                'does the data rows by chunk
                '-----------------------------
                Dim max_cells_per_chunk As Integer = format.max_cells_per_chunk
                Dim total_cols As Integer = last_data_col - format.detail_col_start + 1
                Dim row_chunk As Integer = Int(max_cells_per_chunk / total_cols)
                Dim chunk_cnt As Integer = 0
                Do
                    r1 = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start + chunk_cnt * row_chunk, format.detail_col_start), xlsheet.Cells(Math.Min(last_data_row, format.detail_data_row_start + (chunk_cnt + 1) * row_chunk - 1), last_data_col))
                    Dim rowcnt As Integer = Math.Min(last_data_row, format.detail_data_row_start + (chunk_cnt + 1) * row_chunk - 1) - (format.detail_data_row_start + chunk_cnt * row_chunk) + 1
                    Dim colcnt As Integer = .Columns.Count
                    Dim t_array(rowcnt - 1, colcnt - 1) As String
                    For i = 0 To rowcnt - 1
                        For j = 0 To colcnt - 1
                            t_array(i, j) = .Rows(chunk_cnt * row_chunk + i).Field(Of String)(j)
                        Next
                    Next
                    Try
                        r1.Value2 = t_array
                    Catch ex As Exception
                        err = "ER: error writing data to XL, details: " & ex.ToString
                        GoTo get_out
                    End Try
                    chunk_cnt += 1
                Loop While last_data_row >= format.detail_data_row_start + chunk_cnt * row_chunk
            End With
get_out:
        Catch ex As Exception
            err = "ER: general error writing to XL: " & ex.ToString
        End Try
    End Sub






    'writes a col of data to a col in xl 
    'Use this to write to XL one col at a time
    '-------------------------------------------------
    Public Sub ds2xl_1col(ByVal xlsheet As Excel.Worksheet, ByVal dt As System.Data.DataTable, ByVal dtcol As Integer, ByVal xlcol As Integer, ByVal xl_startrow As Integer, ByRef err As String)
        Try
            Dim a(dt.Rows.Count - 1, 0) As String
            Dim r1 As Excel.Range = xlsheet.Range(xlsheet.Cells(xl_startrow, xlcol), xlsheet.Cells(xl_startrow + a.GetLength(0) - 1, xlcol))
            Dim i As Integer = 0
            For Each row As System.Data.DataRow In dt.Rows
                a(i, 0) = row.Field(Of String)(dtcol)
                i += 1
            Next
            r1.Value2 = a
        Catch ex As Exception
            err = "ER: some kind of error has occured updating values to xl, details: " & ex.ToString
        End Try
    End Sub




    Public Function get_string_array_from_name(x, format, err) As String()
        Dim found As Boolean = False
        Dim t_array() As String = {}
        Try
            Dim fields As FieldInfo() = format.GetType().GetFields()
            For Each fld As FieldInfo In fields
                Dim name As String = fld.Name
                If name Like x Then
                    t_array = fld.GetValue(format)
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                Throw New Exception
            End If
            Return t_array
        Catch ex As Exception
            err = "ER: internal error finding the detail_hdr_name_ array, details: " & ex.ToString
            t_array = {}
            Return t_array
        End Try
    End Function





    Public Function get_integer_array_from_name(x, format, err) As Integer()
        Dim found As Boolean = False
        Dim t_array() As Integer = {}
        Try
            Dim fields As FieldInfo() = format.GetType().GetFields()
            For Each fld As FieldInfo In fields
                Dim name As String = fld.Name
                If name Like x Then
                    t_array = fld.GetValue(format)
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                Throw New Exception
            End If
            Return t_array
        Catch ex As Exception
            err = "ER: internal error finding the detail_protect_hide_array, details: " & ex.ToString
            t_array = {}
            Return t_array
        End Try
    End Function




    Public Sub check_mdt_values(ByVal format As cr_sheet_format, ByVal ant_offset As Integer, ByVal temp_range As Excel.Range, ByRef err As String)
        Try
            For Each cell As Excel.Range In temp_range
                Dim qrows = From row In format.ds_allow.Tables("antennas")
                                Where row.Field(Of String)("lookup") Like o2s(cell.Offset(0, ant_offset).Value2)
                                Select row
                If qrows.Count = 0 Then GoTo skip
                Dim has_mdt As Boolean = False
                If qrows.First.Item("has_mdt") = 1 Then has_mdt = True
                Dim ts As String = o2s(cell.Value2)
                Dim ti As Integer = Int(Val(ts))
                If has_mdt Then
                    If Not (IsNumeric(ts) AndAlso Regex.IsMatch(ts, "^[-]?[0-9]{1,2}$", RegexOptions.IgnoreCase) AndAlso (ti >= -45 And ti <= 45)) Then
                        err = "ERROR!! - MDT must be between -45 and 45 degrees."
                        cell.Value2 = err & " - " & ts
                    End If
                Else
                    If Not ts Like "0" Then
                        err = "ERROR!! - MDT must be 0 for this antenna type."
                        cell.Value2 = err & " - " & ts
                    End If
                End If
skip:
            Next
get_out:
        Catch ex As Exception
            err = "ER: some error checking MDT vals, details " & ex.ToString
        End Try
    End Sub



    Public Sub check_edt_values(ByVal format As cr_sheet_format, ByVal ant_offset As Integer, ByVal temp_range As Excel.Range, ByRef err As String)
        Try
            For Each cell As Excel.Range In temp_range
                Dim qrows = From row In format.ds_allow.Tables("antennas")
                            Where row.Field(Of String)("lookup") Like o2s(cell.Offset(0, ant_offset).Value2)
                            Select row
                If qrows.Count = 0 Then GoTo skip
                Dim dbeam As Boolean = False
                If qrows.First.Item("dual_beam") = 1 Then dbeam = True
                Dim min, max As Integer
                min = qrows.First.Item("edt_min")
                max = qrows.First.Item("edt_max")
                Dim ts_raw As String = o2s(cell.Value2)
                Dim ts() As String = Strings.Split(ts_raw, ",")
                If dbeam And Not ts.Count = 2 Then
                    err = "ERROR!! - This is a dual beam antenna, you must give 2 EDT values per physical unit."
                    cell.Value2 = err & " - " & ts_raw
                    GoTo skip
                ElseIf Not dbeam And Not ts.Count = 1 Then
                    err = "ERROR!! - This is a single beam antenna, you must give 1 EDT value per physical unit."
                    cell.Value2 = err & " - " & ts_raw
                    GoTo skip
                End If
                If Regex.IsMatch(ts_raw, "^[0-9]{1,2}([,][0-9]{1,2})?$", RegexOptions.IgnoreCase) Then
                    If dbeam And ts.Count = 2 Then
                        If Not ((Val(ts(0)) >= min And Val(ts(0)) <= max) AndAlso (Val(ts(1)) >= min And Val(ts(1)) <= max)) Then
                            err = "ERROR!! - EDT must be between: " & min & " and " & max & " for this antenna type"
                            cell.Value2 = err & " - " & ts_raw
                        End If
                    Else
                        If Not (Val(ts(0)) >= min And Val(ts(0)) <= max) Then
                            err = "ERROR!! - EDT must be between: " & min & " and " & max & " for this antenna type"
                            cell.Value2 = err & " - " & ts_raw
                        End If
                    End If
                Else
                    err = "ERROR!! - EDT must be of the format 'integer' for single beam antennas and 'integer,integer' for dual beam antennas"
                    cell.Value2 = err & " - " & ts_raw
                End If
skip:
            Next
get_out:
        Catch ex As Exception
            err = "ER: some error checking EDT vals, details " & ex.ToString
        End Try
    End Sub



    Public Sub check_az_values(ByVal temp_range As Excel.Range, ByRef err As String)
        Try
            For Each cell As Excel.Range In temp_range
                Dim ts As String = o2s(cell.Value2)
                Dim ti As Integer = Int(Val(ts))
                If Not (IsNumeric(ts) AndAlso Regex.IsMatch(ts, "^[0-9]{1,3}$", RegexOptions.IgnoreCase) AndAlso (ti >= 0 And ti <= 359)) Then
                    err = "ERROR!! - Azimuth must be between 0 and 359 degrees."
                    cell.Value2 = err & " - " & ts
                End If
            Next
        Catch ex As Exception
            err = "ER: Tool error checking ex returned az values for review, details: " & ex.ToString
        End Try
    End Sub



    Public Sub check_antenna_values(ByVal format As cr_sheet_format, ByVal temp_range As Excel.Range, ByRef err As String)
        Try
            For Each cell As Excel.Range In temp_range
                Dim ts As String = o2s(cell.Value2)
                Dim qrow = From row In format.ds_allow.Tables("antennas")
                            Where row.Field(Of String)("lookup") Like ts
                            Select row
                If qrow.Count = 0 Then
                    err = "ERROR!! - Unknown antenna type, if you have an antenna that is not specified, please get the execution coordinator to add it to the database."
                    cell.Value2 = err & " - " & ts
                End If
            Next
        Catch ex As Exception
            err = "ER: Tool error checking ex returned ant values for review, details: " & ex.ToString
        End Try
    End Sub



    Public Sub check_ht_coax_len_values(ByVal temp_range As Excel.Range, ByRef err As String)
        Try
            For Each cell As Excel.Range In temp_range
                Dim ts As String = o2s(cell.Value2)
                Dim ti As Integer = Int(Val(ts))
                If Not (IsNumeric(ts) AndAlso Regex.IsMatch(ts, "^[0-9]{1,4}$", RegexOptions.IgnoreCase) And (ti >= 0 And ti <= 1000)) Then
                    err = "ERROR!! - Value must be between 0 and 1000m."
                    cell.Value2 = err & " - " & ts
                End If
            Next
        Catch ex As Exception
            err = "ER: Tool error checking ex returned ht/coaxlen values for review, details: " & ex.ToString
        End Try
    End Sub





    'this unhides the whole sheet and locks all cells
    '---------------------------------------------------
    Public Sub unhide_and_lock_all(ByVal rowmax As Integer, ByVal colmax As Integer, ByVal xlsheet As Excel.Worksheet, ByVal format As cr_sheet_format, ByRef err As String)
        Try
            Dim t_rng As Excel.Range
            With xlsheet
                t_rng = .Range(.Cells(1, 1), .Cells(rowmax + 1, 1))
                t_rng.EntireRow.Hidden = False

                t_rng = .Range(.Cells(1, 1), .Cells(1, colmax + 1))
                t_rng.EntireColumn.Hidden = False

                .Cells.Locked = True
            End With
        Catch ex As Exception
            err = "ER: some internal error unhiding and locking: " & ex.ToString
        End Try
    End Sub




    'this hides particular rows or columns
    '-----------------------------------------
    Public Sub hide_rows_and_cols(ByVal rows() As Integer, ByVal cols() As Integer, ByVal xlsheet As Excel.Worksheet, ByVal format As cr_sheet_format, ByRef err As String)
        Try
            For Each item In rows
                xlsheet.Range(xlsheet.Cells(item, format.common_data_col), xlsheet.Cells(item, format.common_data_col)).EntireRow.Hidden = True
            Next
            For Each item In cols
                xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, item), xlsheet.Cells(format.detail_data_row_start, item)).EntireColumn.Hidden = True
            Next
        Catch ex As Exception
            err = "ER: some error hiding fields, details: " & ex.ToString
        End Try
    End Sub


    'this unlocks particular common cells or detail fields
    '-------------------------------------------------------
    Public Sub unlock_and_input_format(ByVal rows() As Integer, ByVal cols() As Integer, ByVal last_data_row As Integer, ByVal cr_form_type As String, ByVal skip_format As Boolean, ByVal xlsheet As Excel.Worksheet, ByVal format As cr_sheet_format, ByRef err As String)
        Try
            Dim t_rng As Excel.Range
            For Each item In rows
                t_rng = xlsheet.Range(xlsheet.Cells(item, format.common_data_col), xlsheet.Cells(item, format.common_data_col))
                t_rng.MergeArea.Locked = False
                If Not skip_format Then
                    input_format_cell_resub(t_rng)
                End If
            Next
            For Each item In cols
                t_rng = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, item), xlsheet.Cells(last_data_row, item))
                t_rng.Locked = False
                If Not skip_format Then
                    input_format_cell_resub(t_rng)
                End If
            Next
        Catch ex As Exception
            err = "ER: some error unprotecting fields, details: " & ex.ToString
        End Try
    End Sub





    Public Sub pre_open_non_inbox_cr_file_conflict_resolution(ByVal file As String, ByRef err As String)
        Try
            For Each item In FileIO.FileSystem.GetFiles(Path.GetDirectoryName(file), FileIO.SearchOption.SearchTopLevelOnly, "~$" & Path.GetFileName(file))
                force_delete_file(item, err)
                If Not err Like "" Then GoTo get_out
            Next
get_out:
        Catch ex As Exception
            err = "ER: some error doing the pre_open conflict resolution process, details: " & ex.ToString
        End Try
    End Sub



    Public Sub find_xl_cr_file(ByRef cr_id() As String, ByRef input_file As String, ByRef cr_form() As String, ByVal find_all_flag As Boolean, ByVal format As cr_sheet_format, ByVal local As local_machine, ByVal db As mysql_server, ByRef err As String)
        Try
            Dim xlapp As New Excel.Application
            xlapp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow
            xlapp.Visible = format.debug_xl
            xlapp.DisplayAlerts = False
            xlapp.UserControl = False
            xlapp.IgnoreRemoteRequests = True
            xlapp.Interactive = False
            Dim xlbook As Excel.Workbook = Nothing
            Dim cr_path As String = ""
            Dim inbox_path As String = local.base_path & local.inbox
            Dim file_date As Date = DateTime.MinValue
            If cr_id Is Nothing Or cr_id.Length = 0 Then
                cr_id = {""}
            End If
            Dim s_cr_id As String = ""
            Try
                Dim file_array() As String = (From item In FileIO.FileSystem.GetFiles(inbox_path, FileIO.SearchOption.SearchTopLevelOnly, "*.xlsb")
                                             Where Not item Like "~$*"
                                             Select item).Distinct.ToArray
                For Each item In file_array
                    'this opens the xl file => tries normal load, if it fails, tries repair load, if that fails, tries extractdata load, if that succeeds, at least we can id the file is a cr file or not, if it fails, we skip
                    '-------------------------------------------------------------------------
                    'this opens the xl file
                    '-------------------
                    Dim cr_form_to_open As String = item
                    Try
                        xlapp.DisplayAlerts = False
                        xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                    Catch ex As COMException
                        If ex.HResult = -2146827284 Then
                            fix_bad_xlsb_file(cr_form_to_open, xlapp, format, local, err)
                            If Not err Like "" Then GoTo skip
                            Try
                                xlapp.DisplayAlerts = False
                                xlbook = xlapp.Workbooks.Open(cr_form_to_open, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                            Catch exx As Exception
                                GoTo skip
                            End Try
                        Else
                            GoTo skip
                        End If
                    Catch ex As Exception
                        GoTo skip
                    End Try

                    'this checks the CR sheet exists
                    '--------------------------
                    Try
                        Dim x As Excel.Worksheet = xlbook.Worksheets("CR")
                    Catch ex As Exception
                        GoTo skip
                    End Try
                    Dim xlsheet As Excel.Worksheet = xlbook.Worksheets("CR")
                    xlsheet.Activate()
                    With xlsheet
                        'unprotects sheet
                        '--------------------
                        Try
                            If Not (xlsheet.ProtectContents And xlsheet.ProtectScenarios) Then
                                GoTo skip
                            End If
                            xlsheet.Unprotect(format.x_factor)
                            xlsheet.Range("A1").Value2 = "z"        'this disables the worksheet change macro or reset macro
                        Catch ex As Exception
                            GoTo skip
                        End Try

                        If cr_id(0) Like "" And Not find_all_flag Then
                            'searching for the initial form, we do the header check as we have no cr_id yet
                            'This checks all the header cells are present with the correct headers
                            '-----------------------------------------------------------------------
                            For col_index = 0 To format.common_hdr_name.Count - 1
                                Dim xlcell As Excel.Range = xlsheet.Range(xlsheet.Cells(format.common_row_start + col_index, format.common_hdr_col), xlsheet.Cells(format.common_row_start + col_index, format.common_hdr_col))
                                If Not o2s(xlcell.Value2) Like format.common_hdr_name(col_index) Then
                                    GoTo skip
                                End If
                            Next
                            For col_index = 0 To format.detail_hdr_name.Count - 1
                                Dim xlcell As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + col_index), xlsheet.Cells(format.detail_hdr_row_start, format.detail_col_start + col_index))
                                If Not o2s(xlcell.Value2) Like format.detail_hdr_name(col_index) Then
                                    GoTo skip
                                End If
                            Next
                            If FileIO.FileSystem.GetFileInfo(item).LastWriteTime > file_date Then
                                input_file = item
                                file_date = FileIO.FileSystem.GetFileInfo(item).LastWriteTime
                            End If

                        ElseIf Not cr_id(0) = "" And Not find_all_flag Then
                            'Does the cr_id check for all other cases
                            '---------------------------------------------------
                            Dim xlrng As Excel.Range = .Range(.Cells(format.common_row_start, format.common_data_col), .Cells(format.common_row_start, format.common_data_col))
                            If o2s(xlrng.Value2) Like cr_id(0) Then
                                If FileIO.FileSystem.GetFileInfo(item).LastWriteTime > file_date Then
                                    input_file = item
                                    file_date = FileIO.FileSystem.GetFileInfo(item).LastWriteTime
                                End If
                            End If
                        ElseIf find_all_flag Then
                            'gets all the .xlsb/m files with valid cr_ids
                            '---------------------------------------------------
                            Dim xlrng As Excel.Range = .Range(.Cells(format.common_row_start, format.common_data_col), .Cells(format.common_row_start, format.common_data_col))
                            If is_valid_cr_id(o2s(xlrng.Value2), db) Then
                                input_file = input_file & item & ","
                                s_cr_id = s_cr_id & o2s(xlrng.Value2) & ","
                            End If
                        End If
                    End With
skip:
                    xlapp.DisplayAlerts = False
                    xlbook.Close()
                    releaseObject(xlbook)
                Next
                If Not input_file Like "" Then
                    If cr_id(0) Like "" And Not find_all_flag Then
                        Dim x As String = local.base_path & local.inbox & "\New CR Form.xlsb"
                        If Not input_file Like x Then
                            If FileIO.FileSystem.FileExists(x) Then
                                force_delete_file(x, err)
                                If Not err = "" Then GoTo get_out
                            End If
                            FileIO.FileSystem.MoveFile(input_file, x)
                        End If
                        cr_form = {x}

                    ElseIf Not cr_id(0) = "" And Not find_all_flag Then
                        Dim x As String = local.base_path & local.inbox & "\" & cr_id(0) & ".xlsb"
                        If Not input_file Like x Then
                            If FileIO.FileSystem.FileExists(x) Then
                                force_delete_file(x, err)
                                If Not err = "" Then GoTo get_out
                            End If
                            FileIO.FileSystem.MoveFile(input_file, x)
                        End If
                        cr_form = {x}

                    ElseIf find_all_flag Then
                        input_file = Trim(Left(input_file, Len(input_file) - 1))
                        s_cr_id = Trim(Left(s_cr_id, Len(s_cr_id) - 1))
                        cr_form = Strings.Split(input_file, ",").ToArray
                        cr_id = Strings.Split(s_cr_id, ",").ToArray
                    End If
                End If
get_out:
            Catch ex As Exception
                err = "ER: error finding the cr form, details: " & ex.ToString
            Finally
                xlapp.UserControl = True
                xlapp.Interactive = True
                xlapp.IgnoreRemoteRequests = False
                xlapp.Quit()
                releaseObject(xlapp) 'this releases the com object
                GC.Collect()
                If Not input_file Like "" Then
                    If cr_id(0) Like "" And Not find_all_flag Then
                        If Not cr_form(0) Like local.base_path & local.inbox & "\New CR Form.xlsb" Then
                            cr_form = {}
                        End If
                    ElseIf Not cr_id(0) = "" And Not find_all_flag Then
                        If Not cr_form(0) Like local.base_path & local.inbox & "\" & cr_id(0) & ".xlsb" Then
                            cr_form = {}
                        End If
                    ElseIf find_all_flag Then
                        If cr_form(0) Like "" Then
                            cr_form = {}
                            cr_id = {}
                        End If
                    End If
                End If
            End Try
        Catch ex As Exception
            err = "ER: error opening the xl application, details: " & ex.ToString
        End Try
    End Sub



    Public Function text2regex(ByVal s As String) As String
        On Error Resume Next
        'backslash must go first
        s = Regex.Replace(s, "\\", "\\")

        'then the meta chars
        s = Regex.Replace(s, "\.", "\.")
        s = Regex.Replace(s, "\^", "\^")
        s = Regex.Replace(s, "\$", "\$")
        s = Regex.Replace(s, "\(", "\(")
        s = Regex.Replace(s, "\)", "\)")
        s = Regex.Replace(s, "\[", "\[")
        s = Regex.Replace(s, "\{", "\{")
        s = Regex.Replace(s, "\*", "\*")
        s = Regex.Replace(s, "\+", "\+")
        s = Regex.Replace(s, "\?", "\?")
        s = Regex.Replace(s, "\|", "\|")
        Return s
    End Function




    'This converts the read object from excel to a string, need this as if the excel cell is blank, the returned object is nothing which will cause an error if cast to a string variable.  Hence need this to convert nothing to ""
    Public Function o2s(ByVal obj As Object) As String
        On Error Resume Next
        Dim str As String
        If obj Is Nothing Then
            str = ""
        Else
            str = obj.ToString
        End If
        Return str
    End Function


    'converts nathan's combined bame => name(email) to email only
    Public Function c2e(ByVal combined_name As String) As String
        On Error Resume Next
        c2e = Trim(Regex.Replace(combined_name, "((^.*\()|(\)$))", ""))
    End Function




    'so basically, if the date can not be converted to OAdate by XL, then it attempts to parse from text first with US style, then with indo style, if they both fail we throw an error.
    Public Sub get_xl_date(ByVal ts1 As String, ByRef ts2 As DateTime, ByRef err As String)
        Dim xl_double As Double
        If Double.TryParse(ts1, xl_double) Then
            Try
                ts2 = DateTime.FromOADate(xl_double)
            Catch ex As Exception
                GoTo try_string
            End Try
        Else
try_string:
            Try
                If Not DateTime.TryParse(ts1, CultureInfo.CreateSpecificCulture("en-US"), DateTimeStyles.None, ts2) Then
                    If Not DateTime.TryParse(ts1, CultureInfo.CreateSpecificCulture("id-ID"), DateTimeStyles.None, ts2) Then
                        err = "Unsupported Date Format"
                    End If
                End If
            Catch ex As Exception
                err = "Unsupported Date Format"
            End Try
        End If
    End Sub




    'this takes a corrupted .xlsb file (that has been saved in a newer version fo XL and fixes it so it is backwards compatible), 
    'if successfull the same file will be overwritten with the fixed file, the output file has no hidden rows/cols but is protected, it is meant to be fed into the XL processing subs
    '#########################################################################################################
    'actually I do not need this sub anymore, but it has a good way to open XL files using ACE.oledb which I could use in the future, so keep it
    'actually I do not need this sub anymore, but it has a good way to open XL files using ACE.oledb which I could use in the future, so keep it
    'actually I do not need this sub anymore, but it has a good way to open XL files using ACE.oledb which I could use in the future, so keep it
    'NOTE: Interface the user gets an error where they can't open an .xlsb file and they have to repair it which messes up the formatting, tell them upgrade their MS office 2007 to SP2.
    '#########################################################################################################
    Public Sub fix_bad_xlsb_file(ByVal file As String, ByVal xlapp As Excel.Application, ByVal format As cr_sheet_format, local As local_machine, ByRef err As String)
        Dim xlbook As Excel.Workbook = Nothing
        Dim detail_cols As Integer = 0
        Dim det_hdr_name() As String = {}
        Dim cr_form_type As String = ""

        'if we get here we are a green light to fix a corrupted file
        '------------------------------------------------------------
        'this reads the file into a datatable using ACE.OLDEDB, fucking awsome
        '------------------------------------------------------------
        '        Dim stSQL As String = "SELECT F4 FROM [CR$a1:az200000] Where F2='Parameter'"
        '        Dim stSQL As String = "SELECT * FROM [CR$a1:az200000] where not F2 like '' and not F2 like ':'"
        '        Dim stSQL As String = "SELECT * FROM [CR$a1:az200000] where not F2 like '' and F2 like ':'"
        Dim stSQL As String = "SELECT * FROM [CR$a1:ae200025]"
        Dim stCon As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & file & ";Extended Properties=""Excel 12.0;HDR=NO,IMEX=1"";"
        Dim cnt As New OleDb.OleDbConnection(stCon)
        Dim cmd As New OleDbCommand(stSQL, cnt)
        Dim adp As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        Try
            'fill the ds fromt he sheet in question
            '--------------------------------------
            cnt.Open()
            adp.Fill(ds)

            'this find the approriate blank template based on teh cr form type which we get from the cr_id
            '-------------------------------------------------------------------------------------------
            Dim template_form As String = ""
            Dim table_row_start_offset As Integer = 0
            Dim cr_id As String = "-9999"
            For table_row_start_offset = 0 To format.common_row_start - 1
                If ds.Tables(0).Rows(format.common_row_start - 1 - table_row_start_offset).Field(Of String)(format.common_hdr_col - 1) Like "CR_ID" Then
                    cr_id = ds.Tables(0).Rows(format.common_row_start - 1 - table_row_start_offset).Field(Of String)(format.common_data_col - 1)
                    Exit For
                End If
            Next
            If cr_id Like "-9999" Then
                err = "ER: Corrupted .xlsb file"
                GoTo get_out
            ElseIf cr_id Like "" Then
                det_hdr_name = get_string_array_from_name("detail_hdr_name", format, err)
                detail_cols = det_hdr_name.Count
                template_form = local.base_path & local.cr_blank_request_form_dir & "\" & local.cr_form
                cr_form_type = "init"
            ElseIf Regex.IsMatch(cr_id, "^(.*)(prm)([0-9]+)$", RegexOptions.IgnoreCase) Then
                det_hdr_name = get_string_array_from_name("detail_hdr_name_prm", format, err)
                detail_cols = det_hdr_name.Count
                template_form = local.base_path & local.cr_blank_request_form_dir & "\blank_template_prm.xlsb"
                cr_form_type = "prm"
            ElseIf Regex.IsMatch(cr_id, "^(.*)(rfb)([0-9]+)$", RegexOptions.IgnoreCase) Then
                det_hdr_name = get_string_array_from_name("detail_hdr_name_rfb", format, err)
                detail_cols = det_hdr_name.Count
                template_form = local.base_path & local.cr_blank_request_form_dir & "\blank_template_rfb.xlsb"
                cr_form_type = "rfb"
            ElseIf Regex.IsMatch(cr_id, "^(.*)((rfr)|(hdw))([0-9]+)$", RegexOptions.IgnoreCase) Then
                det_hdr_name = get_string_array_from_name("detail_hdr_name_oth", format, err)
                detail_cols = det_hdr_name.Count
                template_form = local.base_path & local.cr_blank_request_form_dir & "\blank_template_oth.xlsb"
                cr_form_type = "oth"
            Else
                err = "ER: Corrupted .xlsb file"
                GoTo get_out
            End If

            'converts all data fields to OAdate numbers in the DT before writing to XL => note if it was blank or it couldn't convert, it just leaves it as is
            '---------------------------------------------------------------------------
            'Does the common dates
            '-----------------------
            Try
                Dim qrows = From row In ds.Tables(0)
                            Let a = row.Field(Of String)(format.common_hdr_col - 1)
                            Where Not a Is Nothing AndAlso Regex.IsMatch(a, "\sDate", RegexOptions.IgnoreCase)
                            Select row
                For Each row In qrows
                    Dim ts1 As String = row.Field(Of String)(format.common_data_col - 1)
                    If Not ts1 Like "" Then
                        Dim ts2 As DateTime = Now
                        Dim temp_err As String = ""
                        get_xl_date(ts1, ts2, temp_err)            'more robust sub to handle all types of weird date input
                        If temp_err Like "" Then
                            row(format.common_data_col - 1) = ts2.Date.ToOADate
                        End If
                    End If
                Next
            Catch ex As Exception
                err = "ER: Couldn't read the common dates while fixing corrupted .xlsb file"
                GoTo get_out
            End Try

            'Does the detail dates
            '-----------------------
            Try
                Dim ts1 As String = ""
                Dim ts2 As DateTime = Now
                Dim index_cr_type As Integer = Array.IndexOf(det_hdr_name, "CR Type")
                Dim index_ex_date As Integer = Array.IndexOf(det_hdr_name, "Execution Date")
                Dim index_ex_planned_date As Integer = Array.IndexOf(det_hdr_name, "Planned Execution Date")
                Dim qrows = From row In ds.Tables(0)
                            Let a = row.Field(Of String)(format.detail_col_start + index_cr_type - 1)
                            Where Not a Is Nothing AndAlso Regex.IsMatch(a, "^(Parameter)|(RF\sBasic)|(RF\sRe-engineering)|(Hardware)$", RegexOptions.IgnoreCase)
                            Select row
                For Each row In qrows
                    ts1 = row.Field(Of String)(format.detail_col_start + index_ex_date - 1)
                    If Not ts1 Like "" Then
                        Dim temp_err As String = ""
                        get_xl_date(ts1, ts2, temp_err)            'more robust sub to handle all types of weird date input
                        If temp_err Like "" Then
                            row(format.detail_col_start + index_ex_date - 1) = ts2.Date.ToOADate
                        End If
                    End If
                    ts1 = row.Field(Of String)(format.detail_col_start + index_ex_planned_date - 1)
                    If Not ts1 Like "" Then
                        Dim temp_err As String = ""
                        get_xl_date(ts1, ts2, temp_err)            'more robust sub to handle all types of weird date input
                        If temp_err Like "" Then
                            row(format.detail_col_start + index_ex_planned_date - 1) = ts2.Date.ToOADate
                        End If
                    End If
                Next
            Catch ex As Exception
                err = "ER: Couldn't read the detail dates while fixing corrupted .xlsb file"
                GoTo get_out
            End Try

            'open the blank template and fill with the dataset data
            '-------------------------------------------------------
            xlapp.DisplayAlerts = False
            xlbook = xlapp.Workbooks.Open(template_form, CorruptLoad:=XlCorruptLoad.xlRepairFile)
            Dim xlsheet As Excel.Worksheet = xlbook.Sheets("CR")
            xlsheet.Activate()
            xlsheet.Unprotect(format.x_factor)
            xlsheet.Range("A1").Value2 = "z"        'this disables the worksheet change macro or reset macro

            'write the Datatable to the CR sheet here
            '------------------------------------------
            dt2xlsheet_generic(True, 1 + table_row_start_offset, 1, ds.Tables(0), xlsheet, format, err)
            If Not err Like "" Then GoTo get_out

            'close the ACE.OLEDB connection
            '---------------------------------
            cnt.Close()
            cnt = Nothing
            adp.Dispose()
            ds.Dispose()

            'fix the date ranges
            '-----------------------
            'First we find the row limit for the cr form
            '----------------------------------------------
            err = ""
            Dim last_data_row As Integer = 0
            Dim index = Array.IndexOf(det_hdr_name, "CR Type")
            Try
                With xlsheet
                    last_data_row = .Columns("A:A").offset(0, index).entirecolumn.Find(What:="*", After:=.Cells(index + 1), LookAt:=XlLookAt.xlWhole, LookIn:=XlFindLookIn.xlValues, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False).Row
                End With
            Catch ex As Exception
                last_data_row = 0
            End Try
            If last_data_row < format.detail_data_row_start Then
                Throw New Exception("the form has no cr data rows")
            End If

            Dim r1 As Excel.Range
            index = Array.IndexOf(format.common_hdr_name, "Open Date")
            r1 = xlsheet.Range(xlsheet.Cells(format.common_row_start + index, format.common_data_col), xlsheet.Cells(format.common_row_start + index, format.common_data_col))
            date_format_cell(r1)
            index = Array.IndexOf(format.common_hdr_name, "Approval Date")
            r1 = xlsheet.Range(xlsheet.Cells(format.common_row_start + index, format.common_data_col), xlsheet.Cells(format.common_row_start + index, format.common_data_col))
            date_format_cell(r1)
            index = Array.IndexOf(format.common_hdr_name, "Planned Execution Date")
            r1 = xlsheet.Range(xlsheet.Cells(format.common_row_start + index, format.common_data_col), xlsheet.Cells(format.common_row_start + index, format.common_data_col))
            date_format_cell(r1)
            index = Array.IndexOf(format.common_hdr_name, "Execution Date")
            r1 = xlsheet.Range(xlsheet.Cells(format.common_row_start + index, format.common_data_col), xlsheet.Cells(format.common_row_start + index, format.common_data_col))
            date_format_cell(r1)
            index = Array.IndexOf(format.common_hdr_name, "Closed Date")
            r1 = xlsheet.Range(xlsheet.Cells(format.common_row_start + index, format.common_data_col), xlsheet.Cells(format.common_row_start + index, format.common_data_col))
            date_format_cell(r1)
            index = Array.IndexOf(det_hdr_name, "Execution Date")
            r1 = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index + 1), xlsheet.Cells(last_data_row, index + 1))
            date_format_cell(r1)
            index = Array.IndexOf(det_hdr_name, "Planned Execution Date")
            r1 = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, index + 1), xlsheet.Cells(last_data_row, index + 1))
            date_format_cell(r1)

            'sets the hide and unlock config for the next stage
            '----------------------------------------------------------
            'first get the current cr_status from the DB, then work out what to do with it
            'this code is not needed right now, as I just do not offer file fix support for update forms case, all other cases do not need this part, so currently cr_status is always = ""
            'this code is not needed right now, as I just do not offer file fix support for update forms case, all other cases do not need this part, so currently cr_status is always = ""
            'this code is not needed right now, as I just do not offer file fix support for update forms case, all other cases do not need this part, so currently cr_status is always = ""
            'this code is not needed right now, as I just do not offer file fix support for update forms case, all other cases do not need this part, so currently cr_status is always = ""
            Dim cr_status As String = ""
            If Not cr_status Like "" Then
                If cr_status Like "Pending Approval" Then
                    Dim col_array() As Integer = get_integer_array_from_name("col_hide_" & cr_form_type & "_for_app", format, err)
                    If Not err Like "" Then GoTo get_out
                    hide_rows_and_cols(format.row_hide_for_app, col_array, xlsheet, format, err)

                ElseIf cr_status Like "Pending Resubmission" Then
                    'sets the hide and unlock config for the next stage (back to start)
                    '-------------------------------------------------------------------
                    Dim col_array() As Integer = get_integer_array_from_name("col_clear_hide_" & cr_form_type & "_for_resub", format, err)
                    If Not err Like "" Then GoTo get_out
                    hide_rows_and_cols(format.row_hide_for_resub, col_array, xlsheet, format, err)
                    col_array = get_integer_array_from_name("col_unprotect_" & cr_form_type & "_for_resub", format, err)
                    If Not err Like "" Then GoTo get_out
                    unlock_and_input_format(format.row_unprotect_for_resub, col_array, 1048576, cr_form_type, True, xlsheet, format, err)

                    'also need to copy the validation down to the bottom
                    '-----------------------------------------------
                    Dim t_rng As Excel.Range = xlsheet.Range(xlsheet.Cells(format.detail_data_row_start, 1), xlsheet.Cells(format.detail_data_row_start, det_hdr_name.Count))
                    t_rng.Copy()
                    t_rng = t_rng.Resize(t_rng.Rows.Count + 1048576 - format.detail_data_row_start, t_rng.Columns.Count)
                    t_rng.PasteSpecial(XlPasteType.xlPasteValidation)

                ElseIf cr_status Like "Pending Execution Planning" Then
                    Dim col_array() As Integer = get_integer_array_from_name("col_hide_" & cr_form_type & "_for_ex_coord", format, err)
                    If Not err Like "" Then GoTo get_out
                    hide_rows_and_cols(format.row_hide_for_ex_coord, col_array, xlsheet, format, err)
                    col_array = get_integer_array_from_name("col_unprotect_" & cr_form_type & "_for_ex_coord", format, err)
                    If Not err Like "" Then GoTo get_out
                    unlock_and_input_format({}, col_array, last_data_row, cr_form_type, False, xlsheet, format, err)

                ElseIf cr_status Like "Pending Execution" Then
                    'sets the hide and unlock config for the next stage (ex)
                    '----------------------------------------------------------------------------------------------
                    Dim col_array() As Integer = get_integer_array_from_name("col_hide_" & cr_form_type & "_for_ex", format, err)
                    If Not err Like "" Then GoTo get_out
                    hide_rows_and_cols(format.row_hide_for_ex, col_array, xlsheet, format, err)
                    col_array = get_integer_array_from_name("col_unprotect_" & cr_form_type & "_for_ex", format, err)
                    If Not err Like "" Then GoTo get_out
                    unlock_and_input_format({}, col_array, last_data_row, cr_form_type, False, xlsheet, format, err)

                ElseIf cr_status Like "Pending Review" Then
                    hide_rows_and_cols(format.row_hide_for_rev, {}, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out

                ElseIf cr_status Like "Execution Complete Pending Attachments" Then
                    hide_rows_and_cols(format.row_hide_for_rev, {}, xlsheet, format, err)
                    If Not err Like "" Then GoTo get_out

                End If
            End If

            'protects the sheet
            '----------------------------------
            xlsheet.Range("C2").Select()
            xlsheet.Protect(format.x_factor, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True)

            'save and exit
            '-------------------
            xlapp.DisplayAlerts = False
            xlbook.SaveAs(file)
get_out:
        Catch ex As Exception
            err = "ER: Error fixing bad .xlsb file, details: " & ex.ToString
        Finally
            Try
                If Not cnt Is Nothing Then
                    cnt.Close()
                    cnt = Nothing
                End If
                If Not adp Is Nothing Then
                    adp.Dispose()
                End If
                If Not ds Is Nothing Then
                    ds.Dispose()
                End If
            Catch ex As Exception
                Debug.WriteLine("not an error just closing some crap down")
            End Try
            Try
                xlapp.DisplayAlerts = False
                xlbook.Close()
            Catch ex As Exception
            End Try
            Try
                releaseObject(xlbook)
            Catch ex As Exception
            End Try
            GC.Collect()
        End Try
    End Sub
End Module


