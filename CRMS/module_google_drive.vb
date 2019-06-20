Module module_google_drive
    Public GD_Service As DriveService = New DriveService
    Public Sub GD_CreateService(ByVal id As String, ByVal secret As String, ByRef err As String)
        Try
            Dim MyUserCredential As UserCredential = GoogleWebAuthorizationBroker.AuthorizeAsync(New ClientSecrets() With {.ClientId = id, .ClientSecret = secret}, {DriveService.Scope.Drive}, "user", CancellationToken.None).Result
            GD_Service = New DriveService(New BaseClientService.Initializer() With {.HttpClientInitializer = MyUserCredential, .ApplicationName = "Google Drive VB Dot Net"})
        Catch ex As Exception
            err = "ER: Error connecting to google drive, details: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub




    Public Sub get_files_from_drive(ByVal clientid As String, ByVal clientsecret As String, ByVal path As String, ByVal bw As BackgroundWorker, ByRef err As String)
        Try
            clean_dir(path, err)
            If Not err Like "" Then GoTo get_out

            'creates the google service if it has not been created
            '-------------------------------------------
            If Not GD_Service.ApplicationName Like "Google Drive VB Dot Net" Then
                GD_CreateService(clientid, clientsecret, err)
                If Not err Like "" Then GoTo get_out
            End If

            'this gets the current file list on GD
            'right now I only know how to get the complete file list in subdirs of the GD accout, getting the dirs and making a tree list seems to be rather difficult, also querying
            'files within a specific dir seems to be difficult, will do it later
            '----------------------------------------------------------------
            Dim drive_file_list As Google.Apis.Drive.v2.Data.FileList = GD_get_itemlist(10000)

            'this filters the file list to get what we want to download and does the DL
            'we can put any filtering criteria in our LINQ statement, name, size, type, exention
            'should also be able to put in the relative path, but I currently do not know how to do this
            '---------------------------------------------------------------------
            If Not drive_file_list Is Nothing Then
                Dim files2get = From item In drive_file_list.Items
                                Let a = item.OriginalFilename, b = item.FileSize, c = item.MimeType
                                Where Not a Is Nothing And Not c Like "application/vnd.google-apps.folder"
                                Select item
                '                                Where Not a Is Nothing AndAlso b < 10000000000.0 AndAlso Not Regex.IsMatch(a, "(unknown)|("")", RegexOptions.IgnoreCase) AndAlso Regex.IsMatch(System.IO.Path.GetExtension(a).Length.ToString, "^[45]$", RegexOptions.IgnoreCase) AndAlso Not item.MarkedViewedByMeDate.HasValue
                '                               Select item

                Dim total_size As Double = 0
                Dim mb_done As Double = 0
                For Each item In files2get
                    total_size = total_size + item.FileSize / 1000000.0
                Next
                For Each item In files2get
                    bw.ReportProgress(100 * mb_done / total_size, "Next file: " & item.FileSize / 1000000.0 & " MB")

                    'this downloads the file, passes the background worker to allow progress reports, however, as I am using sync DL from GD, it will not give status reports within 1 file, only between files
                    '-----------------------------------------------------------------------------------------
                    GD_download_sync(item.DownloadUrl, path & "\" & item.OriginalFilename, bw, err)
                    'Ignore errors

                    'I trash the file after downloading, could also delete or set a flag.
                    '--------------------------------------------------
                    TrashFile(GD_Service, item.Id, err)
                    'Ignore errors

                    mb_done = mb_done + item.FileSize / 1000000.0
                Next
            End If
get_out:
        Catch ex As Exception
            err = "ER: Couldn't get files, details: " & ex.ToString
        End Try
    End Sub




    Public Sub TrashFile(service As DriveService, fileId As [String], ByRef err As String)
        Try
            service.Files.Trash(fileId).Execute()
        Catch ex As Exception
            err = "ER: Couldn't delete file, details: " & ex.ToString
        End Try
    End Sub


    Public Function GD_get_itemlist(ByVal maxlist As Integer) As Google.Apis.Drive.v2.Data.FileList
        Try
            Dim Request = GD_Service.Files.List
            '            Request.Q = "mimeType != 'application/vnd.google-apps.folder' and trashed = false"
            Request.Q = "trashed = false"
            Request.MaxResults = maxlist
            Dim Results = Request.Execute
            Return Results
        Catch ex As Exception
            Return Nothing
        Finally
            GC.Collect()
        End Try
    End Function


 


    Public Sub GD_download_sync(ByVal url As String, ByVal full_filename As String, ByVal bw As BackgroundWorker, ByRef err As String)
        Try
            '            Debug.WriteLine(Now.ToLongTimeString & ": " & "G O O G L E   D R I V E: Downloading " & Path.GetFileName(full_filename) & " Start")

            Dim Downloader = New MediaDownloader(GD_Service)
            Downloader.ChunkSize = 256 * 1024     'KB

            Using FileStream = New System.IO.FileStream(full_filename, System.IO.FileMode.Create, System.IO.FileAccess.Write)
                Dim Progress = Downloader.Download(url, FileStream)
                'Debug.WriteLine(Now.ToLongTimeString & ": " & "G O O G L E   D R I V E: Downloading " & Path.GetFileName(full_filename) & " sdfsdfdgsdfgdfghad errors")

                'do not need anything here, it automatically waits until there is finalisation.
                If Not Progress.Status = DownloadStatus.Completed Then
                    '                    Debug.WriteLine(Now.ToLongTimeString & ": " & "G O O G L E   D R I V E: Downloading " & Path.GetFileName(full_filename) & " had errors")
                Else
                    '                   Debug.WriteLine(Now.ToLongTimeString & ": " & "G O O G L E   D R I V E: Downloading " & Path.GetFileName(full_filename) & " Done")
                End If
            End Using
get_out:
        Catch ex As Exception
            err = "ER: Error downloading from google drive, details: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub








    '####################################################################################################################3
    '####################################################################################################################3
    '####################################################################################################################3
    '####################################################################################################################3
    'This has not been tested, will not work
    '-------------------------------------------
    Public Sub GD_Upload(ByVal FilePath As String, ByRef err As String)
        Try
            Debug.WriteLine("                                ""Google Drive Upload Start")

            Dim TheFile As New Google.Apis.Drive.v2.Data.File()
            TheFile.Title = "My document"
            TheFile.Description = "A test document"
            TheFile.MimeType = "text/plain"

            Dim ByteArray As Byte() = System.IO.File.ReadAllBytes(FilePath)
            Dim Stream As New System.IO.MemoryStream(ByteArray)

            Dim UploadRequest As FilesResource.InsertMediaUpload = GD_Service.Files.Insert(TheFile, Stream, TheFile.MimeType)

            Debug.WriteLine("                                ""Google Drive Upload Finished")
        Catch ex As Exception
            err = "ER: Error uploading to google drive, details: " & ex.ToString
        Finally
            GC.Collect()
        End Try
    End Sub


    'This has not been tested, will not work
    '-------------------------------------------
    Public Async Sub GD_download_async(ByVal url As String, ByVal full_filename As String)
        Try
            Debug.WriteLine(Now.ToLongTimeString & ": " & "G O O G L E   D R I V E: Downloading " & Path.GetFileName(full_filename) & " Start")

            Dim Downloader = New MediaDownloader(GD_Service)
            Downloader.ChunkSize = 256 * 1024     'KB

            ' figure out the right file type
            Using FileStream = New System.IO.FileStream(full_filename, System.IO.FileMode.Create, System.IO.FileAccess.Write)
                Dim Progress = Downloader.DownloadAsync(url, FileStream)

                Dim r As Google.Apis.Download.IDownloadProgress = Await Progress

                If r.Status = DownloadStatus.Completed Then
                    Debug.WriteLine("                                ""Google Drive Download Finished")
                ElseIf r.Status = DownloadStatus.Failed Then
                    Debug.WriteLine("                                ""Google Drive Download Timed Out....")
                ElseIf r.Status = DownloadStatus.Downloading Then
                    Debug.WriteLine("                                ""Google Drive: downloading file....")
                ElseIf r.Status = DownloadStatus.NotStarted Then
                    Debug.WriteLine("                                ""Google Drive: can't start....")
                    GoTo get_out
                End If
            End Using
get_out:
        Catch ex As Exception
            ' err = "ER: Error downloading from google drive, details: " & ex.ToString
        Finally
            '            GC.Collect()
        End Try
    End Sub

End Module
