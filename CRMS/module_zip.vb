Module module_zip

    'zips one file, if the archive doesn't exist, it creates it, if it does, it adds it
    '---------------------------------------------------------------
    Public Sub zip_file(ByVal input_file As String, ByVal output_file As String, ByVal add_flag As Boolean, ByRef err As String)
        If Not add_flag Then
            'this deletes the output file if it already exists, an error here will kill the sub as we can not continue
            Try
                If IO.File.Exists(output_file) Then
                    force_delete_file(output_file, err)
                    If Not err Like "" Then GoTo get_out
                End If
            Catch ex As Exception
                err = "ER: output file already exists and can not be deleted, zip can ont continue, internal server issue, details: " & ex.ToString
                GoTo get_out
            End Try
        End If

        Try
            Using zipfilestream As FileStream = New FileStream(output_file, FileMode.OpenOrCreate)
                Using zipfile As Compression.ZipArchive = New Compression.ZipArchive(zipfilestream, Compression.ZipArchiveMode.Update)
                    Dim newentry As System.IO.Compression.ZipArchiveEntry = zipfile.CreateEntryFromFile(input_file, Path.GetFileName(input_file), Compression.CompressionLevel.Optimal)
                End Using
            End Using
        Catch ex As Exception
            err = "ER: some error zipping a file, details: " & ex.ToString
        End Try
get_out:
    End Sub







    'this zips an entire directory into several zip files of max size as given, the inputs for the multi case are a dir with the files to be zipped 
    'and the output zip file name as well as max part size in MB and max part count
    'if it goes above the max part count, it just makes one big zip file
    'it removes any conflicting named zip files in the output file location first
    '------------------------------------------------------------------
    Public Sub zip_dir(ByVal input_dir As String, ByVal max_size As Double, ByVal max_num As Integer, ByVal output_file As String, ByRef err As String)
        'this deletes the output files if they already exist, an error here will kill the sub as we can not continue
        Try
            If max_size = 0 Then
                '#################################################
                '#################################################
                '#################################################
                '#################################################
just_do_1_zip_file:
                'for the non split case
                '-----------------------
                If IO.File.Exists(output_file) Then
                    force_delete_file(output_file, err)
                    If Not err Like "" Then GoTo get_out
                End If
                Try
                    System.IO.Compression.ZipFile.CreateFromDirectory(input_dir, output_file, Compression.CompressionLevel.Optimal, False)
                Catch ex As Exception
                    err = "ER: Error zipping attachment!!! """ & input_dir & """ => " & ex.ToString
                End Try
                '#################################################
                '#################################################
                '#################################################
                '#################################################

            Else
                'for the split case
                '-------------------
                'get the temp_dir sorted and ready
                '------------------------------------
                Dim temp_dir As String = Path.GetDirectoryName(output_file) & "\temp"
                If FileIO.FileSystem.DirectoryExists(temp_dir) Then
                    clean_dir(temp_dir, err)
                    If Not err Like "" Then GoTo get_out
                Else
                    FileIO.FileSystem.CreateDirectory(temp_dir)
                End If

                'check the output zip files are not there
                '------------------------------------
                Dim i As Integer
                Dim t_file As String = ""
                For i = 1 To Math.Max(1, max_num)
                    t_file = Path.GetDirectoryName(output_file) & "\" & Path.GetFileNameWithoutExtension(output_file) & "_" & i & ".zip"
                    If IO.File.Exists(t_file) Then
                        force_delete_file(t_file, err)
                        If Not err Like "" Then GoTo get_out
                    End If
                Next

                'add the contents to successive zip files up to the max size, if the total number goes above the max num, I just zip to 1 big file
                '----------------------------------------------------------------------------------------------------------------------
                i = 1
                'for the files
                '--------------------
                For Each item In FileIO.FileSystem.GetFiles(input_dir)
                    'copy file to test dir
                    '---------------------
                    Dim item_copy As String = temp_dir & "\" & Path.GetFileName(item)
                    FileIO.FileSystem.CopyFile(item, item_copy)

                    'zip files in test_dir
                    '---------------------
                    t_file = Path.GetDirectoryName(output_file) & "\" & Path.GetFileNameWithoutExtension(output_file) & "_" & i & ".zip"
                    zip_dir(temp_dir, 0, 0, t_file, err)
                    If Not err Like "" Then GoTo get_out

                    'test size
                    '-------------
                    Dim t_file_info As New FileInfo(t_file)
                    If t_file_info.Length / 1000000 > max_size Then
                        If Not i = max_num Then
                            If (FileIO.FileSystem.GetFiles(temp_dir).Count + FileIO.FileSystem.GetDirectories(temp_dir).Count) > 1 Then
                                'take last file out, rezip, clean and recopy last file to temp dir to start the next file
                                force_delete_file(item_copy, err)
                                If Not err Like "" Then GoTo get_out
                                zip_dir(temp_dir, 0, 0, t_file, err)
                                If Not err Like "" Then GoTo get_out
                                clean_dir(temp_dir, err)
                                If Not err Like "" Then GoTo get_out
                                FileIO.FileSystem.CopyFile(item, item_copy)
                            End If
                            'start a new zip file
                            i += 1
                        Else
                            'there are too many parts, so cancel and put all in one big zip file
                            For i = 1 To Math.Max(1, max_num)
                                t_file = Path.GetDirectoryName(output_file) & "\" & Path.GetFileNameWithoutExtension(output_file) & "_" & i & ".zip"
                                If IO.File.Exists(t_file) Then
                                    force_delete_file(t_file, err)
                                    If Not err Like "" Then GoTo get_out
                                End If
                            Next
                            clean_dir(temp_dir, err)
                            If Not err Like "" Then GoTo get_out
                            FileIO.FileSystem.DeleteDirectory(temp_dir, DeleteDirectoryOption.DeleteAllContents)
                            GoTo just_do_1_zip_file
                        End If
                    End If
                Next

                'for the dirs 
                '--------------
                For Each item In FileIO.FileSystem.GetDirectories(input_dir)
                    'copy dir to test dir
                    '---------------------
                    Dim item_copy As String = temp_dir & "\" & Path.GetFileName(If(Right(item, 1) Like "\", item = Left(item, Len(item) - 1), item))
                    FileIO.FileSystem.CopyDirectory(item, item_copy)

                    'zip files/dirs in test_dir
                    '--------------------------
                    t_file = Path.GetDirectoryName(output_file) & "\" & Path.GetFileNameWithoutExtension(output_file) & "_" & i & ".zip"
                    zip_dir(temp_dir, 0, 0, t_file, err)
                    If Not err Like "" Then GoTo get_out

                    'test size
                    '-------------
                    Dim t_file_info As New FileInfo(t_file)
                    If t_file_info.Length / 1000000 > max_size Then
                        If Not i = max_num Then
                            If (FileIO.FileSystem.GetFiles(temp_dir).Count + FileIO.FileSystem.GetDirectories(temp_dir).Count) > 1 Then
                                'take last dir out and rezip 
                                FileIO.FileSystem.DeleteDirectory(item_copy, DeleteDirectoryOption.DeleteAllContents)
                                zip_dir(temp_dir, 0, 0, t_file, err)
                                If Not err Like "" Then GoTo get_out
                                clean_dir(temp_dir, err)
                                If Not err Like "" Then GoTo get_out
                                FileIO.FileSystem.CopyDirectory(item, item_copy)
                            End If
                            'start a new zip file
                            i += 1
                        Else
                            'there are too many parts, so cancel and put all in one big zip file
                            For i = 1 To Math.Max(1, max_num)
                                t_file = Path.GetDirectoryName(output_file) & "\" & Path.GetFileNameWithoutExtension(output_file) & "_" & i & ".zip"
                                If IO.File.Exists(t_file) Then
                                    force_delete_file(t_file, err)
                                    If Not err Like "" Then GoTo get_out
                                End If
                            Next
                            clean_dir(temp_dir, err)
                            If Not err Like "" Then GoTo get_out
                            FileIO.FileSystem.DeleteDirectory(temp_dir, DeleteDirectoryOption.DeleteAllContents)
                            GoTo just_do_1_zip_file
                        End If
                    End If
                Next

                'delete the temp directory
                '-------------------------------
                clean_dir(temp_dir, err)
                If Not err Like "" Then GoTo get_out
                FileIO.FileSystem.DeleteDirectory(temp_dir, DeleteDirectoryOption.DeleteAllContents)
            End If
get_out:
        Catch ex As Exception
            err = "ER: General error zipping!!! => " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub








    'This unzips the given zip file to the given directory using the compression library
    'we overwrite conflicting files
    '-------------------------------------------------------------------
    Public Sub unzip_file(ByVal input_file As String, ByVal output_dir As String, ByRef err As String)
        Try
            'this is the easy way, but it will crash if there are files/dirs already on the HDD with the same name
            '            ZipFile.ExtractToDirectory(input_file, output_dir)

            'so use the hard way, file by file and dir by dir, need to dispose the zip file though after this
            Using archive As Compression.ZipArchive = ZipFile.OpenRead(input_file)
                For Each item In archive.Entries
                    If Right(item.FullName, 1) Like "/" Then
                        'it is a dir, so we manually create it if it is not there
                        '----------------------------------------------------------
                        Dim dir As String = Left(item.FullName, Len(item.FullName) - 1)
                        If Not FileIO.FileSystem.DirectoryExists(output_dir & "\" & dir) Then
                            FileIO.FileSystem.CreateDirectory(output_dir & "\" & dir)
                        End If
                    Else
                        'it is a file so we extract it with an overwrite option, on error we just continue
                        '----------------------------------
                        Try
                            item.ExtractToFile(output_dir & "\" & item.FullName, True)
                        Catch ex As Exception
                        End Try
                    End If
                Next
            End Using
        Catch ex As Exception
            err = "ER: Error unzipping attachment!!! """ & input_file & """ => " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub







    'this will unrar a .rar file and put it in the specified folder
    '-------------------------------------------------------
    Public Sub unrar_file(ByVal input_file As String, ByVal output_dir As String, ByRef err As String)
        Try
            Dim rar_file As SharpCompress.Archive.IArchive = SharpCompress.Archive.ArchiveFactory.Open(input_file)
            '            rar_file.ExtractAllEntries.WriteAllToDirectory(output_dir, ExtractOptions.ExtractFullPath)
            For Each item In rar_file.Entries
                'we only take action on file entries, do not need to do anything in the case the entry is a dir for this RAR tool (different from the ZIP tool)
                '-------------------------------------------------------------------------------------------------------------
                If Not item.IsDirectory Then
                    'it is a file so we extract it with an overwrite option, on error we just continue
                    Try
                        item.WriteToDirectory(output_dir, ExtractOptions.ExtractFullPath + ExtractOptions.Overwrite)
                    Catch ex As Exception
                    End Try
                End If
            Next
            rar_file.Dispose()
        Catch ex As Exception
            err = "ER: Error unrarring attachment!!! """ & input_file & """ => " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub


End Module
