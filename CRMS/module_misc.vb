Module Module_misc


    'this spilts a comma delimited string
    '-----------------------------------------
    Public Function split_comma_delimited_string(ByVal expression As String, ByVal delimiter As String, ByVal qualifier As String, ByVal ignoreCase As Boolean) As String()
        Try
            Dim QualifierState As Boolean = False
            Dim StartIndex As Integer = 0
            Dim Values As New System.Collections.ArrayList

            For CharIndex As Integer = 0 To expression.Length - 1
                If Not qualifier Is Nothing AndAlso String.Compare(expression.Substring(CharIndex, qualifier.Length), qualifier, ignoreCase) = 0 Then
                    QualifierState = Not QualifierState
                ElseIf Not QualifierState AndAlso Not delimiter Is Nothing AndAlso String.Compare(expression.Substring(CharIndex, delimiter.Length), delimiter, ignoreCase) = 0 Then
                    Values.Add(expression.Substring(StartIndex, CharIndex - StartIndex))
                    StartIndex = CharIndex + 1
                End If
            Next

            If StartIndex < expression.Length Then Values.Add(expression.Substring(StartIndex, expression.Length - StartIndex))

            Dim returnValues(Values.Count - 1) As String
            Values.CopyTo(returnValues)
            Return returnValues
        Catch ex As Exception
            Return {}
        End Try
    End Function




    '--------------------------------------------------------------------------------
    'This sub gets a filename and location dialog to the user and returns the full filespec
    '------------------------------------------------------------------------------------
    Public Function get_openfilename_from_user(ByVal s_prompt As String, ByRef err As String) As String
        get_openfilename_from_user = ""

        Try
            Dim getfileandlocation As New OpenFileDialog()

            'This gets the name and location of the new db
            '-----------------------------------------------
            getfileandlocation.Filter = "All Files|*.*"
            getfileandlocation.Title = s_prompt
            getfileandlocation.FilterIndex = 2
            getfileandlocation.RestoreDirectory = True

            If getfileandlocation.ShowDialog() = DialogResult.OK AndAlso Not getfileandlocation.FileName = "" Then
                get_openfilename_from_user = getfileandlocation.FileName
            End If
        Catch ex As Exception
            err = "ER: " & ex.ToString
        End Try
    End Function


    'This sub gives the the choose filename and location dialog to the user and returns the full filespec
    '------------------------------------------------------------------------------------
    Public Function get_savefilename_from_user(ByVal s_ext As String, ByVal s_prompt As String, ByRef err As String) As String
        get_savefilename_from_user = ""

        Try
            Dim getfileandlocation As New System.Windows.Forms.SaveFileDialog()
            Dim s1, s2, s3 As String

            'This gets the name and location of the new db
            '-----------------------------------------------
            getfileandlocation.Filter = "|*" & s_ext & "|All Files|*.*"
            getfileandlocation.Title = s_prompt
            getfileandlocation.FilterIndex = 2
            getfileandlocation.RestoreDirectory = True

            If getfileandlocation.ShowDialog() = DialogResult.OK AndAlso Not getfileandlocation.FileName = "" Then
                s1 = getfileandlocation.FileName
                s2 = Path.GetFileNameWithoutExtension(s1) & s_ext
                s3 = Path.GetDirectoryName(s1)
                get_savefilename_from_user = s3 & "\" & s2
            End If
        Catch ex As Exception
            MsgBox("ER: " & ex.ToString)
        End Try
    End Function






    'This sub gives the the choose directory dialog to the user and returns the full path
    '------------------------------------------------------------------------------------
    Public Function get_path_from_user(ByVal s_ext As String, ByVal s_prompt As String, ByRef err As String) As String
        get_path_from_user = ""
        Try
            Dim s1 As String
            Dim getpath As New System.Windows.Forms.FolderBrowserDialog
            getpath.ShowNewFolderButton = True
            getpath.Description = s_prompt
            getpath.SelectedPath = s_ext

            If getpath.ShowDialog() = DialogResult.OK Then
                s1 = getpath.SelectedPath
                If s1 <> "" Then
                    get_path_from_user = s1
                End If
            End If
        Catch ex As Exception
            MsgBox("ER: " & ex.ToString)
        End Try
    End Function




    'This tests if a given string is a valid path or not
    '----------------------------------------------------------------
    Public Function is_path_valid(ByVal path_s As String) As Boolean
        On Error Resume Next
        is_path_valid = False
        If System.IO.Directory.Exists(path_s) Then
            is_path_valid = True
        End If
    End Function




    'Very useful function, can tell you if the file you specified is open or not
    '-------------------------------------------------------------------------------
    Public Function Is_File_Open(ByVal fullname As String) As Boolean
        Try
            Dim fs As FileStream = Nothing
            If FileIO.FileSystem.FileExists(fullname) Then
                Try
                    fs = System.IO.File.Open(fullname, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                    fs.Close()
                    Is_File_Open = False
                Catch ex As Exception
                    Is_File_Open = True
                Finally
                    fs.Dispose()
                    GC.Collect()
                End Try
            Else
                Is_File_Open = False
            End If
        Catch ex As Exception
            Is_File_Open = True
        End Try
    End Function








    'Release interop objects
    'this is important to release these interop connections to instances of software
    '----------------------------
    Public Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub



    'this kills all processes of this type
    'use "Excel" for xl and "Access" for access
    '--------------------------------------------------
    Public Sub cleanup_processes(ByVal s_app As String)
        On Error Resume Next
        Dim proc As System.Diagnostics.Process

        For Each proc In System.Diagnostics.Process.GetProcessesByName(s_app)
            proc.Kill()
        Next
    End Sub




    'this checks for HDD space, returns the free space available in GB
    '----------------------------------------------------------------
    '    Public Function check_hdd_space(ByVal path As String) As Double
    'no time to do this now

    'End Function




    'this makes a dir no matter what, if it exists, it cleans it, if it doesn't it creates it
    '-----------------------------------------------------------------------------------
    Sub force_make_dir(ByVal path As String, ByRef err As String)
        Try
            If IO.Directory.Exists(path) Then
                clean_dir(path, err)
            Else
                FileIO.FileSystem.CreateDirectory(path)
            End If
        Catch ex As Exception
            err = "ER: Error creating directory (" & path & "), details: " & ex.ToString
        End Try
    End Sub





    'this recursively removes all files then all directories within the given directory, but keeps the directory
    '----------------------------------------------------------------------------------------------------------
    Sub clean_dir(ByVal path As String, ByRef err As String)
        Try
            For Each file_item In IO.Directory.GetFiles(path)
                force_delete_file(file_item, err)
                If Not err = "" Then GoTo get_out
            Next
            For Each dir_item In IO.Directory.GetDirectories(path)
                clean_dir(dir_item, err)
                If Not err = "" Then GoTo get_out
                FileIO.FileSystem.DeleteDirectory(dir_item, DeleteDirectoryOption.DeleteAllContents)
            Next
get_out:
        Catch ex As Exception
            err = "ER: error deleting directory, details " & ex.ToString
        End Try
    End Sub





    'this deletes a file no matter what
    '--------------------------------------
    Sub force_delete_file(ByVal file As String, ByRef err As String)
        Dim retry_cnt As Integer = 0
retry:
        Try
            FileIO.FileSystem.DeleteFile(file, UIOption.OnlyErrorDialogs, RecycleOption.DeletePermanently, UICancelOption.DoNothing)
        Catch ex As System.IO.FileNotFoundException
            GoTo get_out
        Catch ex As System.IO.IOException
            retry_cnt = retry_cnt + 1
            If retry_cnt < 10 Then
                debug.writeline(Now.ToLongTimeString & ": " & "Having issues deleting " & file & ", trying 10x then giving up....")
                kill_proc(file, err)
                If Not err Like "" Then GoTo get_out
                System.Threading.Thread.Sleep(400)
                GoTo retry
            End If
        Catch ex As Exception
            err = "ER: error deleting files, details: " & ex.ToString
        End Try
get_out:
    End Sub









    'this closes whatever process is locking a file, no matter what
    '---------------------------------------------
    Sub force_close_file(ByVal file As String, ByRef err As String)
        Try
            If Is_File_Open(file) Then
                For i = 1 To 100
                    kill_proc(file, err)
                    If Not err Like "" Then GoTo get_out
                    If Not Is_File_Open(file) Then
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            err = "ER: error closing files, details: " & ex.ToString
            GoTo get_out
        End Try
get_out:
    End Sub









    '---------------------------------------------------------------
    '---------------------------------------------------------------
    '---------------------------------------------------------------
    Public Sub kill_proc(ByVal file As String, ByRef err As String)
        On Error Resume Next
        'This gets the processes that are locking a file, very fucking useful
        'to use it do the following
        '-----------------------------------
        Dim x As IList(Of String) = {file}
        Dim y As IList(Of Process) = get_proc.GetProcessesUsingFiles(x)
        For Each proc In y
            If Regex.IsMatch(proc.ProcessName, "^(vs)|(CRMS)", RegexOptions.IgnoreCase) Then
                err = "ER: your app is hanging resources and locking files."
            Else
                proc.Kill()
            End If
        Next

    End Sub

    Public Class get_proc

        <DllImport("rstrtmgr.dll", CharSet:=CharSet.Unicode)> _
        Private Shared Function RmStartSession(ByRef pSessionHandle As UInteger, dwSessionFlags As Integer, strSessionKey As String) As Integer
        End Function

        <DllImport("rstrtmgr.dll")> _
        Private Shared Function RmEndSession(pSessionHandle As UInteger) As Integer
        End Function

        <DllImport("rstrtmgr.dll", CharSet:=CharSet.Unicode)> _
        Private Shared Function RmRegisterResources(pSessionHandle As UInteger, nFiles As UInt32, rgsFilenames As String(), nApplications As UInt32, <[In]()> rgApplications As RM_UNIQUE_PROCESS(), nServices As UInt32, rgsServiceNames As String()) As Integer
        End Function

        <DllImport("rstrtmgr.dll")> _
        Private Shared Function RmGetList(dwSessionHandle As UInteger, ByRef pnProcInfoNeeded As UInteger, ByRef pnProcInfo As UInteger, <[In](), Out()> rgAffectedApps As RM_PROCESS_INFO(), ByRef lpdwRebootReasons As UInteger) As Integer
        End Function

        Private Const RmRebootReasonNone As Integer = 0
        Private Const CCH_RM_MAX_APP_NAME As Integer = 255
        Private Const CCH_RM_MAX_SVC_NAME As Integer = 63

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure RM_UNIQUE_PROCESS
            Public dwProcessId As Integer
            Public ProcessStartTime As ComTypes.FILETIME
        End Structure

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
        Private Structure RM_PROCESS_INFO
            Public Process As RM_UNIQUE_PROCESS
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CCH_RM_MAX_APP_NAME + 1)> _
            Public strAppName As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CCH_RM_MAX_SVC_NAME + 1)> _
            Public strServiceShortName As String
            Public ApplicationType As RM_APP_TYPE
            Public AppStatus As UInteger
            Public TSSessionId As UInteger
            <MarshalAs(UnmanagedType.Bool)> _
            Public bRestartable As Boolean
        End Structure

        Private Enum RM_APP_TYPE
            RmUnknownApp = 0
            RmMainWindow = 1
            RmOtherWindow = 2
            RmService = 3
            RmExplorer = 4
            RmConsole = 5
            RmCritical = 1000
        End Enum

        Public Shared Function GetProcessesUsingFiles(filePaths As IList(Of String)) As IList(Of Process)
            Dim sessionHandle As UInteger
            Dim processes As New List(Of Process)()

            ' Create a restart manager session
            Dim rv As Integer = RmStartSession(sessionHandle, 0, Guid.NewGuid().ToString())
            If rv <> 0 Then
                Throw New Win32Exception()
            End If
            Try
                ' Let the restart manager know what files we’re interested in
                Dim pathStrings As String() = New String(filePaths.Count - 1) {}
                filePaths.CopyTo(pathStrings, 0)
                rv = RmRegisterResources(sessionHandle, CUInt(pathStrings.Length), pathStrings, 0, Nothing, 0, _
                 Nothing)
                If rv <> 0 Then
                    Throw New Win32Exception()
                End If

                ' Ask the restart manager what other applications are using those files
                Const ERROR_MORE_DATA As Integer = 234
                Dim pnProcInfoNeeded As UInteger = 0, pnProcInfo As UInteger = 0, lpdwRebootReasons As UInteger = RmRebootReasonNone
                rv = RmGetList(sessionHandle, pnProcInfoNeeded, pnProcInfo, Nothing, lpdwRebootReasons)
                If rv = ERROR_MORE_DATA Then
                    ' Create an array to store the process results
                    Dim processInfo As RM_PROCESS_INFO() = New RM_PROCESS_INFO(pnProcInfoNeeded - 1) {}
                    pnProcInfo = CUInt(processInfo.Length)

                    ' Get the list
                    rv = RmGetList(sessionHandle, pnProcInfoNeeded, pnProcInfo, processInfo, lpdwRebootReasons)
                    If rv = 0 Then
                        ' Enumerate all of the results and add them to the
                        ' list to be returned
                        For i As Integer = 0 To pnProcInfo - 1
                            Try
                                processes.Add(Process.GetProcessById(processInfo(i).Process.dwProcessId))
                                ' in case the process is no longer running
                            Catch generatedExceptionName As ArgumentException
                            End Try
                        Next
                    Else
                        Throw New Win32Exception()
                    End If
                ElseIf rv <> 0 Then
                    Throw New Win32Exception()
                End If
            Finally
                ' Close the resource manager
                RmEndSession(sessionHandle)
            End Try
            Return processes
        End Function
    End Class



    '######################################################################################################
    '######################################################################################################
    '######################################################################################################
    '######################################################################################################
    'csv stuff
    'Reads a csv into a datatable => uses ado as it is faster than the msdn csv parser
    Private Function read_csv2dt_ado(ByVal filename As String, ByVal pathname As String) As System.Data.DataTable
        Dim dt_out As New System.Data.DataTable
        Dim first_row As Boolean = True
        Dim dt_col_cnt As Integer = 0

        Try
            If Not pathname & filename Is Nothing AndAlso Not pathname & filename = "" AndAlso System.IO.File.Exists(pathname & filename) AndAlso Not Is_File_Open(pathname & filename) Then
                Try
                    'setup connection to CSV file
                    '------------------------------
                    Dim sql_cmd As String
                    sql_cmd = "SELECT * FROM [" & (filename) & "];"

                    Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (pathname) & ";Extended Properties='text;HDR=YES;IMEX=1;FMT=Delimited(,)';")
                    'Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & (pathname) & ";Extended Properties='text;HDR=YES;IMEX=1;FMT=Delimited(,)';")
                    Try
                        Dim Command As System.Data.OleDb.OleDbCommand = New System.Data.OleDb.OleDbCommand(sql_cmd, conn)
                    Catch ex As Exception
                        MsgBox("Error: " & ex.ToString)
                    End Try

                    'Open the Connection
                    '----------------------
                    conn.Open()

                    'Create a dataset to store the file contents
                    '-----------------------------------------------
                    Dim da As New OleDbDataAdapter(sql_cmd, conn)
                    Dim dt_query As New System.Data.DataTable

                    'fill datatable
                    '-----------------
                    da.Fill(dt_out)

                Catch ex As Exception
                    MsgBox("Error: " & ex.ToString)
                End Try
            End If
        Catch ex As Exception
        End Try
        Return dt_out
    End Function


    'Reads a csv into a datatable => using mdsn parser
    Private Function read_csv2dt_msdn(ByVal full_filename As String) As System.Data.DataTable
        Try
            Dim dt As New System.Data.DataTable
            Dim hdr As Boolean = False
            If Not full_filename Is Nothing AndAlso Not full_filename = "" AndAlso System.IO.File.Exists(full_filename) AndAlso Not Is_File_Open(full_filename) Then
                Using reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(full_filename)
                    reader.TextFieldType = FileIO.FieldType.Delimited
                    reader.SetDelimiters(",")
                    reader.HasFieldsEnclosedInQuotes = True   'this will treat any fields with enclosing quotes as 1 field regardless of what is in it
                    While Not reader.EndOfData
                        If Not hdr Then
                            hdr = True
                            Dim t_a() As String = reader.ReadFields
                            For Each item In t_a
                                dt.Columns.Add(item, System.Type.GetType("System.String"))
                                dt.Columns.Item(item).AllowDBNull = False
                                dt.Columns.Item(item).DefaultValue = ""
                            Next
                        Else
                            dt.Rows.Add(reader.ReadFields)
                        End If
                    End While
                End Using
                Return dt
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function





    Private Sub write_dt2csv(ByVal outputfilename As String, ByVal file_path As String, ByRef dt As System.Data.DataTable, ByRef error_msg As String)
        If dt Is Nothing Then GoTo skip
        Try
            'This does the export row by row to csv
            '-----------------------------------------
            Using writer As New StreamWriter(file_path & outputfilename, False)
                writer.WriteLine("""" & Join((From col As System.Data.DataColumn In dt.Columns Select col.ColumnName).ToArray, """,""") & """")
                For Each row As System.Data.DataRow In dt.Rows
                    writer.WriteLine("""" & Join((From item In row.ItemArray Select o2s(item)).ToArray, """,""") & """")
                Next
            End Using
        Catch ex As Exception
            MsgBox("There was an error writing to csv... " & ex.ToString)
        End Try
skip:
    End Sub








    Public Sub adosql2csv(ByVal file_path As String, ByVal file_path_csv As String, ByRef err As String)
        Dim file_path_out = file_path

        'use jet as it seems to be faster than ACE even though ACE is newer, need ACE for reading .xlsx/b files
        Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (file_path_csv) & ";Extended Properties='text;HDR=YES;IMEX=1;FMT=Delimited(,)';")
        'Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & (file_path_csv) & ";Extended Properties='text;HDR=YES;IMEX=1;FMT=Delimited(,)';")
        Dim dt_query As New System.Data.DataTable


        '####################################################################
        '####################################################################
        '####################################################################
        'NOTE: ADO HAS A MAJOR ISSUE WITH COL LENGTH, SEEMS TO BE AROUND 70 CHARS, THEN IT FUCKS UP, THIS WILL MESS UP ANY CSV READING OR QUERYING.
        'MY WORKAROUND IS TO OPEN THE CSV VIA MSDN CSV PARSER, REDUCE THE LENGTH OF THE COL NAME i WANT TO USE AND OVERWRITE THE CSV FILE, THEN EDIT IT

        'use this code to reduce log col names => takes string array => 1) csv file path to change, 2) col to change, 3) new col name, if you have more cols in teh same file to change, just put
        'them in the next value pairs => slot 3 and 4, 5 and 6 etc

        '        Dim t_a(2) As String        'set the bound to 2 min and increments of for each new col you want to reduce
        '       t_a(0) = file_path_csv & "COMBCELL_ALGO_LICENSE.csv"
        '      t_a(1) = "ADD UCELLALGOSWITCH => NBMLDCALGOSWITCH: DLLOAD_BASED_PCPICH_PWR_ADJ_SWITCH"
        '     t_a(2) = "NBMLDCALGOSWITCH: DLLOAD_BASED_PCPICH_PWR_ADJ_SWITCH"
        '    Try
        '        If Not t_a(0) Is Nothing AndAlso Not t_a(0) = "" AndAlso System.IO.File.Exists(t_a(0)) AndAlso Not Is_File_Open(t_a(0)) Then
        'Dim dt As New System.Data.DataTable
        '        dt = read_csv2dt_msdn(t_a(0))
        '       For i = 1 To t_a.GetUpperBound(0) / 2
        '        Dim query = From col As System.Data.DataColumn In dt.Columns Where col.ColumnName = t_a(2 * i - 1) Select col.ColumnName
        '       If query.Count > 0 Then dt.Columns(t_a(2 * i - 1)).ColumnName = t_a(2 * i)
        '      Next
        '     write_dt2csv(Path.GetFileName(t_a(0)), Path.GetDirectoryName(t_a(0)) & "\", dt, "")
        '    End If
        '   Catch ex As Exception
        '        MsgBox("Error reducing csv col name length: " & ex.ToString)
        '       End Try
        '####################################################################
        '####################################################################
        '####################################################################



        'setup connection to CSV file
        '------------------------------
        Dim sql_txt As String
        '        Dim sql_txt As String = "SELECT " & _
        '                                   "* " & _
        '                              "FROM " & _
        '                                 "[query_output_1.csv] As A, " & _
        '                                "[query_output_2.csv] As B1, " & _
        '                               "[query_output_3.csv] As B1, " & _
        '                              "[query_output_4.csv] As B1, " & _
        '                             "B1 INNER JOIN A1 ON (B1.[NodeId] = A1.[ADD UINTRAFREQNCELL => RNCID]) AND (B1.[ADD UCELLSETUP => CELLID] = A1.[ADD UINTRAFREQNCELL => CELLID]), " & _
        '                                        "B1.[ADD UCELLSETUP => CELLNAME] As Host_Cell, " & _
        '                                       "B2.[ADD UCELLSETUP => CELLNAME] As Nbr_Cell, " & _
        '                                      "A1.[ADD UINTRAFREQNCELL => RNCID] & ""_"" & A1.[ADD UINTRAFREQNCELL => CELLID] As Host_CellId, " & _
        '                                     "A1.[ADD UINTRAFREQNCELL => NCELLRNCID] & ""_"" & A1.[ADD UINTRAFREQNCELL => NCELLID] As Nbr_CellId, " & _
        '                                    "B1.[ADD UCELLSETUP => UARFCNDOWNLINK] As Host_Freq, " & _
        '                                   "B2.[ADD UCELLSETUP => UARFCNDOWNLINK] As Nbr_Freq " & _
        '                           "FROM " & _
        '                                 "[Init_UINTRAFREQNCELL.csv] As A1, " & _
        '                                "[COMBCELL_GENERAL.csv] As B2, " & _
        '                               "[COMBCELL_GENERAL.csv] As B1, " & _
        '                              "B1 INNER JOIN A1 ON (B1.[NodeId] = A1.[ADD UINTRAFREQNCELL => RNCID]) AND (B1.[ADD UCELLSETUP => CELLID] = A1.[ADD UINTRAFREQNCELL => CELLID]), " & _
        '                             "B2 INNER JOIN A1 ON (B2.[NodeId] = A1.[ADD UINTRAFREQNCELL => NCELLRNCID]) AND (B2.[ADD UCELLSETUP => CELLID] = A1.[ADD UINTRAFREQNCELL => NCELLID]) " & _
        '                     "WHERE " & _
        '                            "(B1.[ADD UCELLSETUP => UARFCNDOWNLINK] = 10638 " & _
        '                           "OR B1.[ADD UCELLSETUP => UARFCNDOWNLINK] = 10663 " & _
        '                          "OR B1.[ADD UCELLSETUP => UARFCNDOWNLINK] = 10613) " & _
        '                         "AND " & _
        '                        "(B2.[ADD UCELLSETUP => UARFCNDOWNLINK] = 10638 " & _
        '                       "OR B2.[ADD UCELLSETUP => UARFCNDOWNLINK] = 10663 " & _
        '                      "OR B2.[ADD UCELLSETUP => UARFCNDOWNLINK] = 10613);"

        'Create the data adapter to execute the query
        '-----------------------------------------------
        Dim da As OleDbDataAdapter
        Try
            'run
            da = New OleDbDataAdapter(sql_txt, conn)
            conn.Open()
            da.Fill(dt_query)

            Call write_dt2csv("query_1.csv", file_path_out, dt_query, err)

        Catch ex As Exception
            err = "SQL Error: " & ex.ToString
        Finally
            'cleanup
            If Not conn Is Nothing Then
                If Not conn.State = ConnectionState.Closed Then conn.Close()
                conn.Dispose()
            End If
            If Not da Is Nothing Then da.Dispose()
            If Not dt_query Is Nothing Then dt_query.Dispose()
            GC.Collect()
        End Try
    End Sub






    Public Sub linq2dt_from_csv(ByVal file_path As String, ByVal file_path_csv As String, ByRef err As String)
        Dim file_path_out = file_path
        Try
            Dim query_name = "query_1"
            Dim string1 As String = ""
            Dim writer1 As New StreamWriter(file_path_out & query_name & ".csv", False)

            Dim dt1 As New System.Data.DataTable
            Dim dt2_1 As New System.Data.DataTable
            Dim dt2_2 As New System.Data.DataTable

            ' read in csv tables to dts
            ' for really big data csvs, you would have to read in chunks to dt, but then the linq wouldn't work as it needs all the data, kind of fucked here.
            '-------------------------------------------------------------------------
            'dt1 = read_csv2dt_msdn(file_path_csv & file2)      'slow
            'dt2_1 = read_csv2dt_msdn(file_path_csv & file1)    'slow
            dt1 = read_csv2dt_ado("put the csv file name here", file_path_csv)
            dt2_1 = read_csv2dt_ado("put the csv file name here", file_path_csv)
            dt2_2 = dt2_1

            Dim results = From a In dt1
                          Join b1 In dt2_1 On b1.Item("Nodeid") Equals a.Item("ADD UINTRAFREQNCELL => RNCID") And b1.Item("ADD UCELLSETUP => CELLID") Equals a.Item("ADD UINTRAFREQNCELL => CELLID")
                          Join b2 In dt2_2 On b2.Item("Nodeid") Equals a.Item("ADD UINTRAFREQNCELL => NCELLRNCID") And b2.Item("ADD UCELLSETUP => CELLID") Equals a.Item("ADD UINTRAFREQNCELL => NCELLID")
                          Where (b1.Item("ADD UCELLSETUP => UARFCNDOWNLINK") = "10638" Or b1.Item("ADD UCELLSETUP => UARFCNDOWNLINK") = "10663" Or b1.Item("ADD UCELLSETUP => UARFCNDOWNLINK") = "10613") _
                            And (b2.Item("ADD UCELLSETUP => UARFCNDOWNLINK") = "10638" Or b2.Item("ADD UCELLSETUP => UARFCNDOWNLINK") = "10663" Or b2.Item("ADD UCELLSETUP => UARFCNDOWNLINK") = "10613")
                          Select New With {.HOST_CELL = a.Item("ADD UINTRAFREQNCELL => RNCID") & "_" & a.Item("ADD UINTRAFREQNCELL => CELLID"), .HOST_SYS = b1.Item("ADD UCELLSETUP => UARFCNDOWNLINK"), .NBR_CELL = a.Item("ADD UINTRAFREQNCELL => NCELLRNCID") & "_" & a.Item("ADD UINTRAFREQNCELL => NCELLID"), .NBR_SYS = b2.Item("ADD UCELLSETUP => UARFCNDOWNLINK"), .HOST_NAME = b1.Item("ADD UCELLSETUP => CELLNAME"), .NBR_NAME = b2.Item("ADD UCELLSETUP => CELLNAME")}

            'Write the headers first
            '-----------------------------
            string1 = "HOST_CELL,HOST_SYS,NBR_CELL,NBR_SYS,CELLNAME_SOURCE,CELLNAME_TARGET"
            writer1.WriteLine(string1)
            Dim host_sys As String = ""
            Dim nbr_sys As String = ""

            For Each a In results
                Try
                    If a.HOST_SYS.ToString = "10638" Then : host_sys = "3G_F1"
                    ElseIf a.HOST_SYS.ToString = "10663" Then : host_sys = "3G_F2"
                    ElseIf a.HOST_SYS.ToString = "10613" Then : host_sys = "3G_F3"
                    Else : host_sys = a.HOST_SYS.ToString
                    End If
                    If a.NBR_SYS.ToString = "10638" Then : nbr_sys = "3G_F1"
                    ElseIf a.NBR_SYS.ToString = "10663" Then : nbr_sys = "3G_F2"
                    ElseIf a.NBR_SYS.ToString = "10613" Then : nbr_sys = "3G_F3"
                    Else : nbr_sys = a.NBR_SYS.ToString
                    End If

                    string1 = a.HOST_CELL & "," & host_sys & "," & a.NBR_CELL & "," & nbr_sys & ",""" & a.HOST_NAME & """,""" & a.NBR_NAME & """"
                    writer1.WriteLine(string1)
                Catch ex As Exception
                    MsgBox("Some error writing to csv... " & ex.Message)
                End Try
            Next
            writer1.Close()
            writer1 = Nothing
        Catch ex As Exception
            MsgBox("LINQ to dt from csv Error: " & ex.ToString)
        End Try
    End Sub

End Module
