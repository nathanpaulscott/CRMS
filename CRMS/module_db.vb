Module module_db


    'db connect-disconnect
    '-----------------------------------------------
    Public Sub mysql_connect(ByRef db As mysql_server)
        On Error Resume Next
        If db.mysql_con.State = ConnectionState.Open Then
            mysql_disconnect(db)
        End If

        If Not db.mysql_con.State = ConnectionState.Open Then
            db.mysql_con_string = "server=" & db.address & ";user id=" & db.username & ";database=" & db.schema & ";port=" & db.port & ";password=" & db.password & ";"
            db.mysql_con.ConnectionString = db.mysql_con_string
            On Error Resume Next
            db.mysql_con.Open()
        End If
    End Sub

    Public Sub mysql_disconnect(ByRef db As mysql_server)
        On Error Resume Next
        If Not db.mysql_con.State = ConnectionState.Closed Then
            On Error Resume Next
            db.mysql_con.Close()
        End If
    End Sub


    Public Function check_mysql_connection(ByVal db As mysql_server) As Boolean
        Dim test As Boolean = False
        Try
            test = db.mysql_con.Ping
        Catch ex As Exception
            test = False
        End Try
        Return test
    End Function



    'this runs the sql command on the DB, the reader optionn is good for running trough a couple of returned results, but it is better to dump to a file or the notepad or to a gridview
    '------------------------------------------------------------------------------------------------------------
    Public Sub sqlquery(ByVal user As Boolean, ByVal db As mysql_server, ByVal sqltext As String, ByRef dt As System.Data.DataTable, ByRef err As String)
        Try
            'this puts the queried data into a dt object
            '------------------------------------------
            Dim da As New MySqlDataAdapter(sqltext, db.mysql_con)
            da.Fill(dt)
            da.Dispose()
        Catch ex As Exception
            If user Then
                err = "SQLREJ: Your sql command returned an error from the server, details: " & ex.ToString
            Else
                err = "ER: There was an error in the sql query sub, details: " & ex.ToString
            End If
        End Try
    End Sub




    'update/insert or delete from db, basically run nonqueries
    '------------------------------------------------------------------------------------------------------------
    Public Sub anyquery_db(ByVal user As Boolean, ByVal db As mysql_server, ByVal sqltext As String, ByRef err As String)
        Try
            Dim da As New MySqlDataAdapter(sqltext, db.mysql_con)
            da.SelectCommand.ExecuteNonQuery()
            da.Dispose()
        Catch ex As Exception
            If user Then
                err = "SQLREJ: Your sql command returned an error from the server, details: " & ex.ToString
            Else
                err = "ER: There was an error in the sql query sub, details: " & ex.ToString
            End If
        End Try
    End Sub



    'This updates the datatables in the in-memory dataset (ds_allow) from the mysl database for all allowed values
    '------------------------------------------------------------------------------------------------------
    Public Sub allowed_values_db2ds(ByVal db As mysql_server, ByVal format As cr_sheet_format, ByRef err As String)
        Try
            For Each dt As System.Data.DataTable In format.ds_allow.Tables
                clear_dt(dt)
                If Regex.IsMatch(dt.TableName, "^(requesters)|(approvers)|(executors)|(administrators)$") Then
                    Dim sqltext As String = "SELECT DISTINCT name,email,cast(concat(trim(name),' (',trim(email),')') as char(200)) as combined_name " & _
                                            "FROM " & db.schema & ".people " & _
                                            "WHERE " & Left(dt.TableName, Len(dt.TableName) - 1) & " = '1' " & _
                                            "ORDER BY email;"
                    sqlquery(False, db, sqltext, dt, err)

                    'this only keeps those values with valid emails, admin puts in incorrect emails sometimes => typos
                    '------------------------------------------------------------------------
                    Dim dtTemp As System.Data.DataTable = dt.Clone
                    Dim qrows = From row In dt
                                Where format.IsValidEmail(c2e(row("combined_name")))
                                Select row
                    For Each row In qrows
                        dtTemp.ImportRow(row)
                    Next
                    dt.Clear()
                    dt.Merge(dtTemp)
                    dtTemp.Dispose()

                ElseIf Regex.IsMatch(dt.TableName, "_ex_coord$") Then
                    Dim sqltext As String = "SELECT DISTINCT name,email,cast(concat(trim(name),' (',trim(email),')') as char(200)) as combined_name " & _
                                            "FROM " & db.schema & ".people " & _
                                            "WHERE " & dt.TableName & " = '1' " & _
                                            "ORDER BY email;"
                    sqlquery(False, db, sqltext, dt, err)

                    'this only keeps those values with valid emails, admin puts in incorrect emails sometimes => typos
                    '--------------------------------------------
                    Dim dtTemp As System.Data.DataTable = dt.Clone
                    Dim qrows = From row In dt
                                Where format.IsValidEmail(c2e(row("combined_name")))
                                Select row
                    For Each row In qrows
                        dtTemp.ImportRow(row)
                    Next
                    dt.Clear()
                    dt.Merge(dtTemp)
                    dtTemp.Dispose()

                ElseIf Regex.IsMatch(dt.TableName, "^geo$") Then
                    Dim sqltext As String = "SELECT DISTINCT cast(concat(trim(province),', ',trim(regency)) as char(200)) as region,province_short " & _
                                            "FROM " & db.schema & ".geo " & _
                                            "ORDER BY region;"
                    sqlquery(False, db, sqltext, dt, err)

                ElseIf Regex.IsMatch(dt.TableName, "^antennas$") Then
                    Dim sqltext As String = "SELECT DISTINCT cast(concat(trim(`manufacturer`),'_',trim(`antenna`),'_h',trim(`hbw`),'v',trim(`vbw`),'g',trim(`gain_dbi`)) as char(200)) as `lookup`, `antenna`,`manufacturer`,`900`,`1800`,`2100`,`edt_min`,`edt_max`,`has_mdt`,`hbw`,`vbw`,`gain_dbi`,`dual_beam`,`comment` " & _
                                            "FROM " & db.schema & ".antennas " & _
                                            "ORDER BY lookup;"
                    sqlquery(False, db, sqltext, dt, err)

                Else
                    Dim sqltext As String = "SELECT DISTINCT * FROM " & db.schema & "." & dt.TableName & " ORDER BY 1;"
                    sqlquery(False, db, sqltext, dt, err)

                End If
                If Not err Like "" Then GoTo get_out
            Next
get_out:
        Catch ex As Exception
            err = "ER: error updating allowed value tables from DB, details: " & ex.ToString
        End Try
    End Sub



    'this gets the new cr_ids
    '-------------------------
    Public Function get_cr_id_index(ByVal db As mysql_server, ByVal prefix As String, ByRef err As String) As Integer
        Try
            Dim index As Integer = 0
            Dim prefix_len As Integer = Len(prefix) + 1 + 3
            Dim dt As New System.Data.DataTable
            sqlquery(False, db, "select MAX(CAST(SUBSTRING(cr_id," & prefix_len & ") AS SIGNED)) from " & db.schema & ".cr_common where cr_id REGEXP '^" & prefix & "((PRM)|(HDW)|(RFR)|(RFB))[0-9]+$';", dt, err)
            If Not IsDBNull(dt.Rows(0)(0)) Then
                index = Int(Val(dt.Rows(0)(0))) + 1
            Else
                index = 1
            End If
            Return index
        Catch ex As Exception
            err = "ER: There was an error getting the new CR_ID, details: " & ex.ToString
            Return 0
        End Try
    End Function




    'this writes the new cr data to the DB
    '------------------------------------
    Public Sub new_cr2db(ByVal type As String, ByVal ds As System.Data.DataSet, ByVal db As mysql_server, ByVal cr_id As String, ByRef err As String)
        Try
            'add cr to cr common
            '---------------------
            Try
                Dim tablename As String = ""
                If ds.Tables("com") Is Nothing Then tablename = type & "_com" Else tablename = "com"
                With ds.Tables(tablename)
                    If Not ds.Tables(tablename).Rows.Count = 0 Then
                        Dim sqlcmd As New MySqlCommand("", db.mysql_con)
                        Dim tx_c As String = "("
                        For Each col As System.Data.DataColumn In .Columns
                            tx_c = tx_c & col.ColumnName & ","
                        Next
                        tx_c = Left(tx_c, Len(tx_c) - 1) & ")"

                        Dim tx_v As String = ""
                        Dim i As Integer = 0
                        Dim qrows = From row In .AsEnumerable()
                                    Where row.Field(Of String)("cr_id") Like cr_id
                        If qrows Is Nothing Or qrows.Count = 0 Then
                            err = "ER: your input dataset doesn't contain the cr_id specified."
                            GoTo get_out
                        End If
                        For Each item In qrows(0).ItemArray
                            tx_v = tx_v & "@p" & i & ","
                            If Not Regex.IsMatch(.Columns(i).ColumnName, "_date", RegexOptions.IgnoreCase) Then
                                sqlcmd.Parameters.AddWithValue("@p" & i, item)
                            Else
                                If Not IsNumeric(item) Then
                                    sqlcmd.Parameters.AddWithValue("@p" & i, DBNull.Value)
                                Else
                                    Dim d As Double = 0
                                    If Double.TryParse(item, d) Then
                                        Try
                                            sqlcmd.Parameters.AddWithValue("@p" & i, DateTime.FromOADate(d))
                                        Catch ex As Exception
                                            sqlcmd.Parameters.AddWithValue("@p" & i, DBNull.Value)
                                        End Try
                                    Else
                                        sqlcmd.Parameters.AddWithValue("@p" & i, DBNull.Value)
                                    End If
                                    '                                    sqlcmd.Parameters.AddWithValue("@p" & i, DateTime.FromOADate(Val(item.ToString)))
                                End If
                            End If
                            i += 1
                        Next
                        tx_v = Left(tx_v, Len(tx_v) - 1)
                        sqlcmd.CommandText = "INSERT IGNORE INTO " & db.schema & ".cr_common" & tx_c & " VALUES(" & tx_v & ");"
                        sqlcmd.CommandType = CommandType.Text
                        sqlcmd.CommandTimeout = db.big_command_timeout
                        sqlcmd.ExecuteNonQuery()
                        sqlcmd.Dispose()
                    End If
                End With
            Catch ex As Exception
                err = "ER: can't update common table in db: " & type & ", details: " & ex.ToString
                GoTo get_out
            End Try

            'add cr to cr data
            '---------------------
            'upgraded to write in chunks of 100K rows instead of all
            Try
                Dim tablename As String = ""
                If ds.Tables("det") Is Nothing Then tablename = type & "_data" Else tablename = "det"
                With ds.Tables(tablename)
                    If Not ds.Tables(tablename).Rows.Count = 0 Then
                        Dim qrows = From row In .AsEnumerable()
                                    Where row.Field(Of String)("cr_sub_id") Like cr_id & "*"
                                    Select row
                        Dim total_rows As Integer = qrows.Count
                        If qrows Is Nothing Or total_rows = 0 Then
                            err = "ER: your input dataset doesn't contain the cr_id you have specified!?"
                            GoTo get_out
                        End If
                        Dim chunk_size As Integer = db.mysql_max_rows_per_command
                        Dim chunk_cnt As Integer = 0
                        Dim row_index As Integer = 0
                        Dim last_row As Integer = total_rows - 1
                        Do
                            '########################################################
                            Dim sqlcmd As New MySqlCommand("", db.mysql_con)
                            Dim tx_v As New StringBuilder("INSERT IGNORE INTO " & db.schema & ".cr_data_" & type & " VALUES")
                            Dim i As Integer = 0
                            For row_index = chunk_cnt * chunk_size To Math.Min(last_row, (chunk_cnt + 1) * chunk_size - 1)       'Each row As System.Data.DataRow In qrows
                                tx_v.Append("(")
                                Dim j As Integer = 0
                                For Each item In .Rows(row_index).ItemArray
                                    tx_v.Append("@p" & i & ",")
                                    If Not Regex.IsMatch(.Columns(j).ColumnName, "_date", RegexOptions.IgnoreCase) Then
                                        sqlcmd.Parameters.AddWithValue("@p" & i, item)
                                    Else
                                        If Not IsNumeric(item) Then
                                            sqlcmd.Parameters.AddWithValue("@p" & i, DBNull.Value)
                                        Else
                                            Dim d As Double = 0
                                            If Double.TryParse(item, d) Then
                                                Try
                                                    sqlcmd.Parameters.AddWithValue("@p" & i, DateTime.FromOADate(d))
                                                Catch ex As Exception
                                                    sqlcmd.Parameters.AddWithValue("@p" & i, DBNull.Value)
                                                End Try
                                            Else
                                                sqlcmd.Parameters.AddWithValue("@p" & i, DBNull.Value)
                                            End If
                                            '                                    sqlcmd.Parameters.AddWithValue("@p" & i, DateTime.FromOADate(Val(item.ToString)))
                                        End If
                                    End If
                                    j += 1
                                    i += 1
                                Next
                                tx_v.Replace(",", "),", tx_v.Length - 1, 1)
                            Next
                            tx_v.Replace(",", ";", tx_v.Length - 1, 1)
                            sqlcmd.CommandText = tx_v.ToString
                            sqlcmd.CommandType = CommandType.Text
                            sqlcmd.CommandTimeout = db.big_command_timeout
                            sqlcmd.ExecuteNonQuery()
                            sqlcmd.Dispose()
                            '################################################################
                            chunk_cnt += 1
                        Loop While last_row >= chunk_cnt * chunk_size
                    End If
                End With
            Catch ex As Exception
                err = "ER: can't update data table in db: " & type & ", details: " & ex.ToString
                GoTo get_out
            End Try
get_out:
        Catch ex As Exception
            err = "ER: error updating db, details: " & ex.ToString
        End Try
    End Sub






    'this updates cr detailed data with an insert-update on duplicate key command
    'this is 1000s of times faster than the row per row update command
    'only problem is that if a key doesn't exist, it will add it to the DB
    '--------------------------------------------------------------------
    Public Sub update_cr_data(ByVal db As mysql_server, ByVal dt As System.Data.DataTable, ByVal cr_data_table As String, ByRef err As String)
        'INSERT IGNORE INTO table (id,Col1,Col2) VALUES (1,1,1),(2,2,3),(3,9,3),(4,10,12) ON DUPLICATE KEY UPDATE Col1=VALUES(Col1),Col2=VALUES(Col2);
        Try
            With dt
                Dim total_rows As Integer = .Rows.Count
                If Not total_rows = 0 Then
                    If Not .Columns(0).ColumnName Like "cr_sub_id" Then
                        Throw New Exception
                    End If
                    Dim chunk_size As Integer = db.mysql_max_rows_per_command
                    Dim chunk_cnt As Integer = 0
                    Dim row_index As Integer = 0
                    Dim last_dt_row As Integer = total_rows - 1
                    Do
                        Dim sqlcmd As New MySqlCommand("", db.mysql_con)
                        Dim tx_c As New StringBuilder(" (")
                        For Each col As System.Data.DataColumn In .Columns
                            tx_c.Append(col.ColumnName & ",")
                        Next
                        tx_c.Replace(",", ")", tx_c.Length - 1, 1)

                        Dim tx_c2 As New StringBuilder(" ON DUPLICATE KEY UPDATE ")
                        For Each col As System.Data.DataColumn In .Columns
                            If col.Ordinal > 0 Then
                                tx_c2.Append(col.ColumnName & "=VALUES(" & col.ColumnName & "),")
                            End If
                        Next
                        tx_c2.Replace(",", "", tx_c2.Length - 1, 1)

                        Dim tx_v As New StringBuilder(" VALUES ")
                        Dim i As Integer = 0
                        For row_index = chunk_cnt * chunk_size To Math.Min(last_dt_row, (chunk_cnt + 1) * chunk_size - 1)       'Each row As System.Data.DataRow In qrows
                            'For Each row As System.Data.DataRow In .Rows
                            tx_v.Append("(")
                            Dim j As Integer = 0
                            For Each item In .Rows(row_index).ItemArray
                                tx_v.Append("@p" & i & ",")
                                If Not Regex.IsMatch(.Columns(j).ColumnName, "_date", RegexOptions.IgnoreCase) Then
                                    sqlcmd.Parameters.AddWithValue("@p" & i, item)
                                Else
                                    If Not IsNumeric(item) Then
                                        sqlcmd.Parameters.AddWithValue("@p" & i, DBNull.Value)
                                    Else
                                        Dim d As Double = 0
                                        If Double.TryParse(item, d) Then
                                            Try
                                                sqlcmd.Parameters.AddWithValue("@p" & i, DateTime.FromOADate(d))
                                            Catch ex As Exception
                                                sqlcmd.Parameters.AddWithValue("@p" & i, DBNull.Value)
                                            End Try
                                        Else
                                            sqlcmd.Parameters.AddWithValue("@p" & i, DBNull.Value)
                                        End If
                                        '                                    sqlcmd.Parameters.AddWithValue("@p" & i, DateTime.FromOADate(Val(item.ToString)))
                                    End If
                                End If
                                j += 1
                                i += 1
                            Next
                            tx_v.Replace(",", "),", tx_v.Length - 1, 1)
                        Next
                        tx_v.Replace(",", "", tx_v.Length - 1, 1)
                        sqlcmd.CommandText = "INSERT IGNORE INTO " & db.schema & "." & cr_data_table & tx_c.ToString & tx_v.ToString & tx_c2.ToString & ";"
                        sqlcmd.CommandType = CommandType.Text
                        sqlcmd.CommandTimeout = db.big_command_timeout
                        sqlcmd.ExecuteNonQuery()
                        sqlcmd.Dispose()
                        '################################################################
                        chunk_cnt += 1
                    Loop While last_dt_row >= chunk_cnt * chunk_size
                End If
            End With
get_out:
        Catch ex As Exception
            err = "ER: can't update data table in db: " & cr_data_table & ", details: " & ex.ToString
            GoTo get_out
        End Try
    End Sub







    'adds executors to the DB
    Public Sub add_people2db(ByVal db As mysql_server, ByVal dt As System.Data.DataTable, ByVal format As cr_sheet_format, ByRef err As String)
        Try
            If Not dt.Rows.Count = 0 Then
                Dim sqlcmd As New MySqlCommand("", db.mysql_con)
                Dim tx_v As New StringBuilder("INSERT IGNORE INTO " & db.schema & ".people (name,email,requester,approver,rfb_ex_coord,rfr_ex_coord,hdw_ex_coord,prm_ex_coord,executor,query,anyquery,administrator) VALUES")
                Dim i As Integer = 0
                For Each row As System.Data.DataRow In dt.Rows
                    Dim j As Integer = 0
                    tx_v.Append("(")
                    For Each item In row.ItemArray
                        tx_v.Append("@p" & i & ",")
                        sqlcmd.Parameters.AddWithValue("@p" & i, item)
                        j += 1
                        i += 1
                    Next
                    tx_v.Replace(",", "),", tx_v.Length - 1, 1)
                Next
                tx_v.Replace(",", ";", tx_v.Length - 1, 1)
                sqlcmd.CommandText = tx_v.ToString
                sqlcmd.CommandType = CommandType.Text
                sqlcmd.CommandTimeout = 60
                sqlcmd.ExecuteNonQuery()
                sqlcmd.Dispose()
                allowed_values_db2ds(db, format, err)
                If Not err Like "" Then GoTo get_out
            End If
get_out:
        Catch ex As Exception
            err = "ER: can't add executors to the people table in the db, details: " & ex.ToString
        End Try
    End Sub





    'this updates the cr status for the given cr_ids
    '----------------------------------------------------
    Public Sub update_cr_common_table(ByVal db As mysql_server, ByVal cr_id As String, ByVal col As String, ByVal val As String, ByRef err As String)
        'modify data in 1 col
        '----------------------------
        Try
            Dim sqlcmd As New MySqlCommand("", db.mysql_con)
            sqlcmd.CommandText = "UPDATE IGNORE " & db.schema & ".cr_common SET " & col & " = @val WHERE cr_id = @crid;"
            sqlcmd.Parameters.AddWithValue("@val", val)
            sqlcmd.Parameters.AddWithValue("@crid", cr_id)
            sqlcmd.CommandType = CommandType.Text
            sqlcmd.CommandTimeout = 60
            sqlcmd.ExecuteNonQuery()
            sqlcmd.Dispose()
        Catch ex As Exception
            err = "ER: can't update cr status in db, details: " & ex.ToString
        End Try
    End Sub




    'only used for date updates
    Public Sub update_cr_common_table_date(ByVal db As mysql_server, ByVal cr_id As String, ByVal col As String, ByVal val As DateTime, ByVal null_flag As Boolean, ByRef err As String)
        'modify data in 1 col
        '----------------------------
        Try
            Dim sqlcmd As New MySqlCommand("", db.mysql_con)
            sqlcmd.CommandText = "UPDATE IGNORE " & db.schema & ".cr_common SET " & col & " = @val WHERE cr_id = @crid;"
            sqlcmd.Parameters.AddWithValue("@val", If(null_flag, DBNull.Value, val))
            sqlcmd.Parameters.AddWithValue("@crid", cr_id)
            sqlcmd.CommandType = CommandType.Text
            sqlcmd.CommandTimeout = 60
            sqlcmd.ExecuteNonQuery()
            sqlcmd.Dispose()
        Catch ex As Exception
            err = "ER: can't update cr status in db, details: " & ex.ToString
        End Try
    End Sub



    'adds an entry to the log
    '----------------------------------------------------
    Public Sub add2log(ByVal db As mysql_server, ByVal time_now As DateTime, ByVal cr_id As String, ByVal log_msg As String, ByRef err As String)
        'add event to log
        '---------------------
        Try
            Dim sqlcmd As New MySqlCommand("", db.mysql_con)
            sqlcmd.CommandText = "INSERT IGNORE INTO " & db.schema & ".log (date,cr_id,event) VALUES(@date,@cr_id,@msg);"
            sqlcmd.Parameters.AddWithValue("@date", time_now)
            sqlcmd.Parameters.AddWithValue("@cr_id", cr_id)
            sqlcmd.Parameters.AddWithValue("@msg", log_msg)
            sqlcmd.CommandType = CommandType.Text
            sqlcmd.CommandTimeout = 60
            sqlcmd.ExecuteNonQuery()
            sqlcmd.Dispose()
        Catch ex As Exception
            err = "ER: can't update log in db, details: " & ex.ToString
        End Try
    End Sub




    'gets the current log data for a cr_id for the email chain
    '----------------------------------------------------------
    Public Function get_log_string(ByVal new_string As String, ByVal db As mysql_server, cr_id As String, ByRef err As String) As String
        Try
            Dim dt As New System.Data.DataTable
            Dim sqltext As String = "select * from " & db.schema & ".log where cr_id = '" & cr_id & "' order by date desc;"
            sqlquery(False, db, sqltext, dt, err)
            Dim x As String = ""
            x = x & "<BR>--------------------------------------------------------------------------------------------------------"
            x = x & "<BR>CR History to Date"
            x = x & "<BR>--------------------------------------------------------------------------------------------------------"
            x = x & "<BR>" & new_string
            For Each row As System.Data.DataRow In dt.Rows
                x = x & "<BR>" & row.Item(0).ToString & " | " & row.Item(1).ToString & " | " & row.Item(2).ToString
            Next
            x = x & "<BR>--------------------------------------------------------------------------------------------------------"
            Return x
        Catch ex As Exception
            err = "ER: error adding log to email, details: " & ex.ToString
            Return ""
        End Try
    End Function



    Public Sub delete_cr_common(ByVal db As mysql_server, ByVal cr_id As String, ByRef err As String)
        'delete cr from db
        '-------------------
        Try
            Dim sqlcmd As New MySqlCommand("", db.mysql_con)
            sqlcmd.CommandText = "DELETE IGNORE FROM " & db.schema & ".cr_common WHERE cr_id = @crid;"
            sqlcmd.Parameters.AddWithValue("@crid", cr_id)
            sqlcmd.CommandType = CommandType.Text
            sqlcmd.CommandTimeout = 60
            sqlcmd.ExecuteNonQuery()
            sqlcmd.Dispose()
        Catch ex As Exception
            err = "ER: can't delete from db, details: " & ex.ToString
        End Try
    End Sub




    Public Sub delete_cr_detail(ByVal db As mysql_server, ByVal table As String, ByVal cr_id As String, ByRef err As String)
        'delete cr from db
        '-------------------
        Try
            Dim sqlcmd As New MySqlCommand("", db.mysql_con)
            sqlcmd.CommandText = "DELETE IGNORE FROM " & db.schema & "." & table & " WHERE cr_sub_id LIKE @crid;"
            sqlcmd.Parameters.AddWithValue("@crid", cr_id & "%")
            sqlcmd.CommandType = CommandType.Text
            sqlcmd.CommandTimeout = db.big_command_timeout
            sqlcmd.ExecuteNonQuery()
            sqlcmd.Dispose()
        Catch ex As Exception
            err = "ER: can't delete from db, details: " & ex.ToString
        End Try
    End Sub

    Public Sub delete_cr_log(ByVal db As mysql_server, ByVal cr_id As String, ByRef err As String)
        'delete cr from db
        '-------------------
        Try
            Dim sqlcmd As New MySqlCommand("", db.mysql_con)
            sqlcmd.CommandText = "DELETE IGNORE FROM " & db.schema & ".log WHERE cr_id = @crid;"
            sqlcmd.Parameters.AddWithValue("@crid", cr_id)
            sqlcmd.CommandType = CommandType.Text
            sqlcmd.CommandTimeout = db.big_command_timeout
            sqlcmd.ExecuteNonQuery()
            sqlcmd.Dispose()
        Catch ex As Exception
            err = "ER: can't delete from db, details: " & ex.ToString
        End Try
    End Sub




    Public Function is_valid_cr_id(ByVal cr_id As String, ByVal db As mysql_server) As Boolean
        Try
            Dim err As String = ""
            Dim dt As New System.Data.DataTable
            Dim sqltext As String = "SELECT cr_id FROM " & db.schema & ".cr_common WHERE cr_id = '" & cr_id & "';"
            sqlquery(False, db, sqltext, dt, err)
            If dt.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
            dt.Dispose()
        Catch ex As Exception
            Return False
        End Try
    End Function
End Module
