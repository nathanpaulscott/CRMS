<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class main_form
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(main_form))
        Me.Timer_check_mail = New System.Windows.Forms.Timer(Me.components)
        Me.TabPage_email = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TextBox_GD_clientsecret = New System.Windows.Forms.TextBox()
        Me.TextBox_GD_clientId = New System.Windows.Forms.TextBox()
        Me.Labelx = New System.Windows.Forms.Label()
        Me.Label_GD_clientId = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBox_pop_use_ssl = New System.Windows.Forms.CheckBox()
        Me.TextBox_pop_port = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TextBox_pop_pass = New System.Windows.Forms.TextBox()
        Me.TextBox_pop_user = New System.Windows.Forms.TextBox()
        Me.TextBox_pop_server = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox_mysql_connect = New System.Windows.Forms.GroupBox()
        Me.TextBox_mysql_port = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.TextBox_mysql_db = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.TextBox_mysql_password = New System.Windows.Forms.TextBox()
        Me.TextBox_mysql_user = New System.Windows.Forms.TextBox()
        Me.TextBox_mysql_server = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox_settings = New System.Windows.Forms.GroupBox()
        Me.CheckBox_use_imap = New System.Windows.Forms.CheckBox()
        Me.Label_cr_status_check_period = New System.Windows.Forms.Label()
        Me.TextBox_cr_status_times = New System.Windows.Forms.TextBox()
        Me.Label_check_mail2 = New System.Windows.Forms.Label()
        Me.Label_timeout_timer = New System.Windows.Forms.Label()
        Me.TextBox_check_mail_period = New System.Windows.Forms.TextBox()
        Me.Label_timer = New System.Windows.Forms.Label()
        Me.Label_check_mail_period = New System.Windows.Forms.Label()
        Me.Button_settings_submit = New System.Windows.Forms.Button()
        Me.Button_base_browse = New System.Windows.Forms.Button()
        Me.TextBox_base_path = New System.Windows.Forms.TextBox()
        Me.Label_base_dl_dir = New System.Windows.Forms.Label()
        Me.GroupBox_smtp = New System.Windows.Forms.GroupBox()
        Me.TextBox_smtp_timeout = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CheckBox_smtp_ssl = New System.Windows.Forms.CheckBox()
        Me.TextBox_smtp_port = New System.Windows.Forms.TextBox()
        Me.Label_smtp_port = New System.Windows.Forms.Label()
        Me.TextBox_smtp_pass = New System.Windows.Forms.TextBox()
        Me.TextBox_smtp_user = New System.Windows.Forms.TextBox()
        Me.TextBox_smtp_server = New System.Windows.Forms.TextBox()
        Me.Label_smtp_pass = New System.Windows.Forms.Label()
        Me.Label_smtp_user = New System.Windows.Forms.Label()
        Me.Label_smtp_server = New System.Windows.Forms.Label()
        Me.GroupBox_pop = New System.Windows.Forms.GroupBox()
        Me.Label_warning = New System.Windows.Forms.Label()
        Me.TextBox_imap_encryption_type = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CheckBox_imap_ssl = New System.Windows.Forms.CheckBox()
        Me.TextBox_imap_port = New System.Windows.Forms.TextBox()
        Me.Label_imap_port = New System.Windows.Forms.Label()
        Me.TextBox_imap_pass = New System.Windows.Forms.TextBox()
        Me.TextBox_imap_user = New System.Windows.Forms.TextBox()
        Me.TextBox_imap_server = New System.Windows.Forms.TextBox()
        Me.Label_imap_pass = New System.Windows.Forms.Label()
        Me.Label_imap_user = New System.Windows.Forms.Label()
        Me.Label_imap_server = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.Timer_get_email_timeout_timer = New System.Windows.Forms.Timer(Me.components)
        Me.BackgroundWorker_get_email = New System.ComponentModel.BackgroundWorker()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.BackgroundWorker_get_files_from_google_drive = New System.ComponentModel.BackgroundWorker()
        Me.TabPage_email.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox_mysql_connect.SuspendLayout()
        Me.GroupBox_settings.SuspendLayout()
        Me.GroupBox_smtp.SuspendLayout()
        Me.GroupBox_pop.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Timer_check_mail
        '
        Me.Timer_check_mail.Enabled = True
        Me.Timer_check_mail.Interval = 1000
        '
        'TabPage_email
        '
        Me.TabPage_email.Controls.Add(Me.GroupBox2)
        Me.TabPage_email.Controls.Add(Me.Label2)
        Me.TabPage_email.Controls.Add(Me.GroupBox1)
        Me.TabPage_email.Controls.Add(Me.Label6)
        Me.TabPage_email.Controls.Add(Me.Label5)
        Me.TabPage_email.Controls.Add(Me.GroupBox_mysql_connect)
        Me.TabPage_email.Controls.Add(Me.GroupBox_settings)
        Me.TabPage_email.Controls.Add(Me.GroupBox_smtp)
        Me.TabPage_email.Controls.Add(Me.GroupBox_pop)
        Me.TabPage_email.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_email.Name = "TabPage_email"
        Me.TabPage_email.Size = New System.Drawing.Size(1075, 515)
        Me.TabPage_email.TabIndex = 8
        Me.TabPage_email.Text = "IMAP-SMTP Email"
        Me.TabPage_email.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.TextBox_GD_clientsecret)
        Me.GroupBox2.Controls.Add(Me.TextBox_GD_clientId)
        Me.GroupBox2.Controls.Add(Me.Labelx)
        Me.GroupBox2.Controls.Add(Me.Label_GD_clientId)
        Me.GroupBox2.Location = New System.Drawing.Point(542, 10)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(257, 152)
        Me.GroupBox2.TabIndex = 82
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "DL from Google Drive"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 86)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(193, 13)
        Me.Label7.TabIndex = 45
        Me.Label7.Text = "NOTE: for this to work the client ID and"
        '
        'TextBox_GD_clientsecret
        '
        Me.TextBox_GD_clientsecret.Location = New System.Drawing.Point(82, 51)
        Me.TextBox_GD_clientsecret.Name = "TextBox_GD_clientsecret"
        Me.TextBox_GD_clientsecret.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TextBox_GD_clientsecret.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_GD_clientsecret.TabIndex = 44
        Me.TextBox_GD_clientsecret.UseSystemPasswordChar = True
        '
        'TextBox_GD_clientId
        '
        Me.TextBox_GD_clientId.Location = New System.Drawing.Point(82, 25)
        Me.TextBox_GD_clientId.Name = "TextBox_GD_clientId"
        Me.TextBox_GD_clientId.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_GD_clientId.TabIndex = 43
        '
        'Labelx
        '
        Me.Labelx.AutoSize = True
        Me.Labelx.Location = New System.Drawing.Point(9, 58)
        Me.Labelx.Name = "Labelx"
        Me.Labelx.Size = New System.Drawing.Size(64, 13)
        Me.Labelx.TabIndex = 41
        Me.Labelx.Text = "ClientSecret"
        '
        'Label_GD_clientId
        '
        Me.Label_GD_clientId.AutoSize = True
        Me.Label_GD_clientId.Location = New System.Drawing.Point(9, 28)
        Me.Label_GD_clientId.Name = "Label_GD_clientId"
        Me.Label_GD_clientId.Size = New System.Drawing.Size(42, 13)
        Me.Label_GD_clientId.TabIndex = 39
        Me.Label_GD_clientId.Text = "ClientId"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 433)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(311, 34)
        Me.Label2.TabIndex = 81
        Me.Label2.Text = "management system."
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CheckBox_pop_use_ssl)
        Me.GroupBox1.Controls.Add(Me.TextBox_pop_port)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.TextBox_pop_pass)
        Me.GroupBox1.Controls.Add(Me.TextBox_pop_user)
        Me.GroupBox1.Controls.Add(Me.TextBox_pop_server)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 168)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(265, 153)
        Me.GroupBox1.TabIndex = 78
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "DL from POP Server"
        '
        'CheckBox_pop_use_ssl
        '
        Me.CheckBox_pop_use_ssl.AutoSize = True
        Me.CheckBox_pop_use_ssl.Location = New System.Drawing.Point(110, 58)
        Me.CheckBox_pop_use_ssl.Name = "CheckBox_pop_use_ssl"
        Me.CheckBox_pop_use_ssl.Size = New System.Drawing.Size(46, 17)
        Me.CheckBox_pop_use_ssl.TabIndex = 71
        Me.CheckBox_pop_use_ssl.Text = "SSL"
        Me.CheckBox_pop_use_ssl.UseVisualStyleBackColor = True
        '
        'TextBox_pop_port
        '
        Me.TextBox_pop_port.Location = New System.Drawing.Point(68, 56)
        Me.TextBox_pop_port.Name = "TextBox_pop_port"
        Me.TextBox_pop_port.Size = New System.Drawing.Size(36, 20)
        Me.TextBox_pop_port.TabIndex = 48
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(44, 59)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(26, 13)
        Me.Label11.TabIndex = 47
        Me.Label11.Text = "Port"
        '
        'TextBox_pop_pass
        '
        Me.TextBox_pop_pass.Location = New System.Drawing.Point(68, 112)
        Me.TextBox_pop_pass.Name = "TextBox_pop_pass"
        Me.TextBox_pop_pass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TextBox_pop_pass.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_pop_pass.TabIndex = 44
        Me.TextBox_pop_pass.UseSystemPasswordChar = True
        '
        'TextBox_pop_user
        '
        Me.TextBox_pop_user.Location = New System.Drawing.Point(68, 82)
        Me.TextBox_pop_user.Name = "TextBox_pop_user"
        Me.TextBox_pop_user.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_pop_user.TabIndex = 43
        '
        'TextBox_pop_server
        '
        Me.TextBox_pop_server.Location = New System.Drawing.Point(68, 30)
        Me.TextBox_pop_server.Name = "TextBox_pop_server"
        Me.TextBox_pop_server.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_pop_server.TabIndex = 28
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(27, 115)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(30, 13)
        Me.Label12.TabIndex = 41
        Me.Label12.Text = "Pass"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(27, 85)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(29, 13)
        Me.Label13.TabIndex = 39
        Me.Label13.Text = "User"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(27, 33)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(38, 13)
        Me.Label14.TabIndex = 37
        Me.Label14.Text = "Server"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(8, 387)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(457, 34)
        Me.Label6.TabIndex = 80
        Me.Label6.Text = "It is running the change request"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.OrangeRed
        Me.Label5.Location = New System.Drawing.Point(8, 341)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(471, 34)
        Me.Label5.TabIndex = 79
        Me.Label5.Text = "Please do not shutdown this app"
        '
        'GroupBox_mysql_connect
        '
        Me.GroupBox_mysql_connect.Controls.Add(Me.TextBox_mysql_port)
        Me.GroupBox_mysql_connect.Controls.Add(Me.Label18)
        Me.GroupBox_mysql_connect.Controls.Add(Me.TextBox_mysql_db)
        Me.GroupBox_mysql_connect.Controls.Add(Me.Label17)
        Me.GroupBox_mysql_connect.Controls.Add(Me.TextBox_mysql_password)
        Me.GroupBox_mysql_connect.Controls.Add(Me.TextBox_mysql_user)
        Me.GroupBox_mysql_connect.Controls.Add(Me.TextBox_mysql_server)
        Me.GroupBox_mysql_connect.Controls.Add(Me.Label19)
        Me.GroupBox_mysql_connect.Controls.Add(Me.Label16)
        Me.GroupBox_mysql_connect.Controls.Add(Me.Label4)
        Me.GroupBox_mysql_connect.Location = New System.Drawing.Point(279, 168)
        Me.GroupBox_mysql_connect.Name = "GroupBox_mysql_connect"
        Me.GroupBox_mysql_connect.Size = New System.Drawing.Size(257, 153)
        Me.GroupBox_mysql_connect.TabIndex = 77
        Me.GroupBox_mysql_connect.TabStop = False
        Me.GroupBox_mysql_connect.Text = "Connect to MySQL DB"
        '
        'TextBox_mysql_port
        '
        Me.TextBox_mysql_port.Location = New System.Drawing.Point(168, 30)
        Me.TextBox_mysql_port.Name = "TextBox_mysql_port"
        Me.TextBox_mysql_port.Size = New System.Drawing.Size(63, 20)
        Me.TextBox_mysql_port.TabIndex = 48
        Me.TextBox_mysql_port.Text = "3306"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(137, 33)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(26, 13)
        Me.Label18.TabIndex = 47
        Me.Label18.Text = "Port"
        '
        'TextBox_mysql_db
        '
        Me.TextBox_mysql_db.Location = New System.Drawing.Point(71, 120)
        Me.TextBox_mysql_db.Name = "TextBox_mysql_db"
        Me.TextBox_mysql_db.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_mysql_db.TabIndex = 46
        Me.TextBox_mysql_db.Text = "crms_db"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(27, 123)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(46, 13)
        Me.Label17.TabIndex = 45
        Me.Label17.Text = "Schema"
        '
        'TextBox_mysql_password
        '
        Me.TextBox_mysql_password.Location = New System.Drawing.Point(71, 90)
        Me.TextBox_mysql_password.Name = "TextBox_mysql_password"
        Me.TextBox_mysql_password.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_mysql_password.TabIndex = 44
        Me.TextBox_mysql_password.UseSystemPasswordChar = True
        '
        'TextBox_mysql_user
        '
        Me.TextBox_mysql_user.Location = New System.Drawing.Point(71, 60)
        Me.TextBox_mysql_user.Name = "TextBox_mysql_user"
        Me.TextBox_mysql_user.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_mysql_user.TabIndex = 43
        Me.TextBox_mysql_user.Text = "crms_user"
        '
        'TextBox_mysql_server
        '
        Me.TextBox_mysql_server.Location = New System.Drawing.Point(71, 30)
        Me.TextBox_mysql_server.Name = "TextBox_mysql_server"
        Me.TextBox_mysql_server.Size = New System.Drawing.Size(63, 20)
        Me.TextBox_mysql_server.TabIndex = 28
        Me.TextBox_mysql_server.Text = "127.0.0.1"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(27, 93)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(30, 13)
        Me.Label19.TabIndex = 41
        Me.Label19.Text = "Pass"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(27, 63)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(29, 13)
        Me.Label16.TabIndex = 39
        Me.Label16.Text = "User"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(27, 33)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(38, 13)
        Me.Label4.TabIndex = 37
        Me.Label4.Text = "Server"
        '
        'GroupBox_settings
        '
        Me.GroupBox_settings.Controls.Add(Me.CheckBox_use_imap)
        Me.GroupBox_settings.Controls.Add(Me.Label_cr_status_check_period)
        Me.GroupBox_settings.Controls.Add(Me.TextBox_cr_status_times)
        Me.GroupBox_settings.Controls.Add(Me.Label_check_mail2)
        Me.GroupBox_settings.Controls.Add(Me.Label_timeout_timer)
        Me.GroupBox_settings.Controls.Add(Me.TextBox_check_mail_period)
        Me.GroupBox_settings.Controls.Add(Me.Label_timer)
        Me.GroupBox_settings.Controls.Add(Me.Label_check_mail_period)
        Me.GroupBox_settings.Controls.Add(Me.Button_settings_submit)
        Me.GroupBox_settings.Controls.Add(Me.Button_base_browse)
        Me.GroupBox_settings.Controls.Add(Me.TextBox_base_path)
        Me.GroupBox_settings.Controls.Add(Me.Label_base_dl_dir)
        Me.GroupBox_settings.Location = New System.Drawing.Point(805, 10)
        Me.GroupBox_settings.Name = "GroupBox_settings"
        Me.GroupBox_settings.Size = New System.Drawing.Size(261, 312)
        Me.GroupBox_settings.TabIndex = 75
        Me.GroupBox_settings.TabStop = False
        Me.GroupBox_settings.Text = "Settings"
        '
        'CheckBox_use_imap
        '
        Me.CheckBox_use_imap.AutoSize = True
        Me.CheckBox_use_imap.Location = New System.Drawing.Point(29, 168)
        Me.CheckBox_use_imap.Name = "CheckBox_use_imap"
        Me.CheckBox_use_imap.Size = New System.Drawing.Size(71, 17)
        Me.CheckBox_use_imap.TabIndex = 77
        Me.CheckBox_use_imap.Text = "Use Imap"
        Me.CheckBox_use_imap.UseVisualStyleBackColor = True
        '
        'Label_cr_status_check_period
        '
        Me.Label_cr_status_check_period.AutoSize = True
        Me.Label_cr_status_check_period.Location = New System.Drawing.Point(26, 63)
        Me.Label_cr_status_check_period.Name = "Label_cr_status_check_period"
        Me.Label_cr_status_check_period.Size = New System.Drawing.Size(120, 13)
        Me.Label_cr_status_check_period.TabIndex = 57
        Me.Label_cr_status_check_period.Text = "Check CR Status Times"
        '
        'TextBox_cr_status_times
        '
        Me.TextBox_cr_status_times.Location = New System.Drawing.Point(152, 60)
        Me.TextBox_cr_status_times.Name = "TextBox_cr_status_times"
        Me.TextBox_cr_status_times.Size = New System.Drawing.Size(87, 20)
        Me.TextBox_cr_status_times.TabIndex = 56
        '
        'Label_check_mail2
        '
        Me.Label_check_mail2.AutoSize = True
        Me.Label_check_mail2.Location = New System.Drawing.Point(211, 33)
        Me.Label_check_mail2.Name = "Label_check_mail2"
        Me.Label_check_mail2.Size = New System.Drawing.Size(12, 13)
        Me.Label_check_mail2.TabIndex = 55
        Me.Label_check_mail2.Text = "s"
        '
        'Label_timeout_timer
        '
        Me.Label_timeout_timer.AutoSize = True
        Me.Label_timeout_timer.Location = New System.Drawing.Point(26, 281)
        Me.Label_timeout_timer.Name = "Label_timeout_timer"
        Me.Label_timeout_timer.Size = New System.Drawing.Size(136, 13)
        Me.Label_timeout_timer.TabIndex = 74
        Me.Label_timeout_timer.Text = "Get Mail Timeout Timer: 0 s"
        '
        'TextBox_check_mail_period
        '
        Me.TextBox_check_mail_period.Location = New System.Drawing.Point(161, 30)
        Me.TextBox_check_mail_period.Name = "TextBox_check_mail_period"
        Me.TextBox_check_mail_period.Size = New System.Drawing.Size(44, 20)
        Me.TextBox_check_mail_period.TabIndex = 53
        '
        'Label_timer
        '
        Me.Label_timer.AutoSize = True
        Me.Label_timer.Location = New System.Drawing.Point(26, 256)
        Me.Label_timer.Name = "Label_timer"
        Me.Label_timer.Size = New System.Drawing.Size(109, 13)
        Me.Label_timer.TabIndex = 73
        Me.Label_timer.Text = "Check Mail Timer: 0 s"
        '
        'Label_check_mail_period
        '
        Me.Label_check_mail_period.AutoSize = True
        Me.Label_check_mail_period.Location = New System.Drawing.Point(26, 33)
        Me.Label_check_mail_period.Name = "Label_check_mail_period"
        Me.Label_check_mail_period.Size = New System.Drawing.Size(93, 13)
        Me.Label_check_mail_period.TabIndex = 54
        Me.Label_check_mail_period.Text = "Check Mail Period"
        '
        'Button_settings_submit
        '
        Me.Button_settings_submit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_settings_submit.Location = New System.Drawing.Point(29, 202)
        Me.Button_settings_submit.Name = "Button_settings_submit"
        Me.Button_settings_submit.Size = New System.Drawing.Size(117, 45)
        Me.Button_settings_submit.TabIndex = 52
        Me.Button_settings_submit.Text = "Submit Changes"
        Me.Button_settings_submit.UseVisualStyleBackColor = True
        '
        'Button_base_browse
        '
        Me.Button_base_browse.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_base_browse.Location = New System.Drawing.Point(186, 130)
        Me.Button_base_browse.Margin = New System.Windows.Forms.Padding(3, 2, 3, 3)
        Me.Button_base_browse.Name = "Button_base_browse"
        Me.Button_base_browse.Size = New System.Drawing.Size(54, 24)
        Me.Button_base_browse.TabIndex = 60
        Me.Button_base_browse.Text = "Browse"
        Me.Button_base_browse.UseVisualStyleBackColor = True
        '
        'TextBox_base_path
        '
        Me.TextBox_base_path.Location = New System.Drawing.Point(29, 133)
        Me.TextBox_base_path.Name = "TextBox_base_path"
        Me.TextBox_base_path.Size = New System.Drawing.Size(148, 20)
        Me.TextBox_base_path.TabIndex = 59
        Me.TextBox_base_path.TabStop = False
        '
        'Label_base_dl_dir
        '
        Me.Label_base_dl_dir.AutoSize = True
        Me.Label_base_dl_dir.Location = New System.Drawing.Point(26, 117)
        Me.Label_base_dl_dir.Name = "Label_base_dl_dir"
        Me.Label_base_dl_dir.Size = New System.Drawing.Size(60, 13)
        Me.Label_base_dl_dir.TabIndex = 58
        Me.Label_base_dl_dir.Text = "Email Path:"
        '
        'GroupBox_smtp
        '
        Me.GroupBox_smtp.Controls.Add(Me.TextBox_smtp_timeout)
        Me.GroupBox_smtp.Controls.Add(Me.Label3)
        Me.GroupBox_smtp.Controls.Add(Me.CheckBox_smtp_ssl)
        Me.GroupBox_smtp.Controls.Add(Me.TextBox_smtp_port)
        Me.GroupBox_smtp.Controls.Add(Me.Label_smtp_port)
        Me.GroupBox_smtp.Controls.Add(Me.TextBox_smtp_pass)
        Me.GroupBox_smtp.Controls.Add(Me.TextBox_smtp_user)
        Me.GroupBox_smtp.Controls.Add(Me.TextBox_smtp_server)
        Me.GroupBox_smtp.Controls.Add(Me.Label_smtp_pass)
        Me.GroupBox_smtp.Controls.Add(Me.Label_smtp_user)
        Me.GroupBox_smtp.Controls.Add(Me.Label_smtp_server)
        Me.GroupBox_smtp.Location = New System.Drawing.Point(279, 9)
        Me.GroupBox_smtp.Name = "GroupBox_smtp"
        Me.GroupBox_smtp.Size = New System.Drawing.Size(257, 153)
        Me.GroupBox_smtp.TabIndex = 63
        Me.GroupBox_smtp.TabStop = False
        Me.GroupBox_smtp.Text = "UL to SMTP Server"
        '
        'TextBox_smtp_timeout
        '
        Me.TextBox_smtp_timeout.Location = New System.Drawing.Point(156, 56)
        Me.TextBox_smtp_timeout.Name = "TextBox_smtp_timeout"
        Me.TextBox_smtp_timeout.Size = New System.Drawing.Size(40, 20)
        Me.TextBox_smtp_timeout.TabIndex = 78
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(196, 59)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 13)
        Me.Label3.TabIndex = 77
        Me.Label3.Text = "Timeout"
        '
        'CheckBox_smtp_ssl
        '
        Me.CheckBox_smtp_ssl.AutoSize = True
        Me.CheckBox_smtp_ssl.Location = New System.Drawing.Point(111, 57)
        Me.CheckBox_smtp_ssl.Name = "CheckBox_smtp_ssl"
        Me.CheckBox_smtp_ssl.Size = New System.Drawing.Size(46, 17)
        Me.CheckBox_smtp_ssl.TabIndex = 72
        Me.CheckBox_smtp_ssl.Text = "SSL"
        Me.CheckBox_smtp_ssl.UseVisualStyleBackColor = True
        '
        'TextBox_smtp_port
        '
        Me.TextBox_smtp_port.Location = New System.Drawing.Point(68, 55)
        Me.TextBox_smtp_port.Name = "TextBox_smtp_port"
        Me.TextBox_smtp_port.Size = New System.Drawing.Size(37, 20)
        Me.TextBox_smtp_port.TabIndex = 48
        '
        'Label_smtp_port
        '
        Me.Label_smtp_port.AutoSize = True
        Me.Label_smtp_port.Location = New System.Drawing.Point(36, 58)
        Me.Label_smtp_port.Name = "Label_smtp_port"
        Me.Label_smtp_port.Size = New System.Drawing.Size(26, 13)
        Me.Label_smtp_port.TabIndex = 47
        Me.Label_smtp_port.Text = "Port"
        '
        'TextBox_smtp_pass
        '
        Me.TextBox_smtp_pass.Location = New System.Drawing.Point(68, 108)
        Me.TextBox_smtp_pass.Name = "TextBox_smtp_pass"
        Me.TextBox_smtp_pass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TextBox_smtp_pass.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_smtp_pass.TabIndex = 44
        Me.TextBox_smtp_pass.UseSystemPasswordChar = True
        '
        'TextBox_smtp_user
        '
        Me.TextBox_smtp_user.Location = New System.Drawing.Point(68, 82)
        Me.TextBox_smtp_user.Name = "TextBox_smtp_user"
        Me.TextBox_smtp_user.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_smtp_user.TabIndex = 43
        '
        'TextBox_smtp_server
        '
        Me.TextBox_smtp_server.Location = New System.Drawing.Point(68, 30)
        Me.TextBox_smtp_server.Name = "TextBox_smtp_server"
        Me.TextBox_smtp_server.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_smtp_server.TabIndex = 28
        '
        'Label_smtp_pass
        '
        Me.Label_smtp_pass.AutoSize = True
        Me.Label_smtp_pass.Location = New System.Drawing.Point(27, 115)
        Me.Label_smtp_pass.Name = "Label_smtp_pass"
        Me.Label_smtp_pass.Size = New System.Drawing.Size(30, 13)
        Me.Label_smtp_pass.TabIndex = 41
        Me.Label_smtp_pass.Text = "Pass"
        '
        'Label_smtp_user
        '
        Me.Label_smtp_user.AutoSize = True
        Me.Label_smtp_user.Location = New System.Drawing.Point(27, 85)
        Me.Label_smtp_user.Name = "Label_smtp_user"
        Me.Label_smtp_user.Size = New System.Drawing.Size(29, 13)
        Me.Label_smtp_user.TabIndex = 39
        Me.Label_smtp_user.Text = "User"
        '
        'Label_smtp_server
        '
        Me.Label_smtp_server.AutoSize = True
        Me.Label_smtp_server.Location = New System.Drawing.Point(27, 33)
        Me.Label_smtp_server.Name = "Label_smtp_server"
        Me.Label_smtp_server.Size = New System.Drawing.Size(38, 13)
        Me.Label_smtp_server.TabIndex = 37
        Me.Label_smtp_server.Text = "Server"
        '
        'GroupBox_pop
        '
        Me.GroupBox_pop.Controls.Add(Me.Label_warning)
        Me.GroupBox_pop.Controls.Add(Me.TextBox_imap_encryption_type)
        Me.GroupBox_pop.Controls.Add(Me.Label1)
        Me.GroupBox_pop.Controls.Add(Me.CheckBox_imap_ssl)
        Me.GroupBox_pop.Controls.Add(Me.TextBox_imap_port)
        Me.GroupBox_pop.Controls.Add(Me.Label_imap_port)
        Me.GroupBox_pop.Controls.Add(Me.TextBox_imap_pass)
        Me.GroupBox_pop.Controls.Add(Me.TextBox_imap_user)
        Me.GroupBox_pop.Controls.Add(Me.TextBox_imap_server)
        Me.GroupBox_pop.Controls.Add(Me.Label_imap_pass)
        Me.GroupBox_pop.Controls.Add(Me.Label_imap_user)
        Me.GroupBox_pop.Controls.Add(Me.Label_imap_server)
        Me.GroupBox_pop.Location = New System.Drawing.Point(8, 9)
        Me.GroupBox_pop.Name = "GroupBox_pop"
        Me.GroupBox_pop.Size = New System.Drawing.Size(265, 153)
        Me.GroupBox_pop.TabIndex = 23
        Me.GroupBox_pop.TabStop = False
        Me.GroupBox_pop.Text = "DL from IMAP Server"
        '
        'Label_warning
        '
        Me.Label_warning.AutoSize = True
        Me.Label_warning.Location = New System.Drawing.Point(27, 275)
        Me.Label_warning.Name = "Label_warning"
        Me.Label_warning.Size = New System.Drawing.Size(0, 13)
        Me.Label_warning.TabIndex = 77
        '
        'TextBox_imap_encryption_type
        '
        Me.TextBox_imap_encryption_type.Location = New System.Drawing.Point(156, 55)
        Me.TextBox_imap_encryption_type.Name = "TextBox_imap_encryption_type"
        Me.TextBox_imap_encryption_type.Size = New System.Drawing.Size(40, 20)
        Me.TextBox_imap_encryption_type.TabIndex = 76
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(196, 58)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 75
        Me.Label1.Text = "Encrypt"
        '
        'CheckBox_imap_ssl
        '
        Me.CheckBox_imap_ssl.AutoSize = True
        Me.CheckBox_imap_ssl.Location = New System.Drawing.Point(110, 58)
        Me.CheckBox_imap_ssl.Name = "CheckBox_imap_ssl"
        Me.CheckBox_imap_ssl.Size = New System.Drawing.Size(46, 17)
        Me.CheckBox_imap_ssl.TabIndex = 71
        Me.CheckBox_imap_ssl.Text = "SSL"
        Me.CheckBox_imap_ssl.UseVisualStyleBackColor = True
        '
        'TextBox_imap_port
        '
        Me.TextBox_imap_port.Location = New System.Drawing.Point(68, 56)
        Me.TextBox_imap_port.Name = "TextBox_imap_port"
        Me.TextBox_imap_port.Size = New System.Drawing.Size(36, 20)
        Me.TextBox_imap_port.TabIndex = 48
        '
        'Label_imap_port
        '
        Me.Label_imap_port.AutoSize = True
        Me.Label_imap_port.Location = New System.Drawing.Point(44, 59)
        Me.Label_imap_port.Name = "Label_imap_port"
        Me.Label_imap_port.Size = New System.Drawing.Size(26, 13)
        Me.Label_imap_port.TabIndex = 47
        Me.Label_imap_port.Text = "Port"
        '
        'TextBox_imap_pass
        '
        Me.TextBox_imap_pass.Location = New System.Drawing.Point(68, 112)
        Me.TextBox_imap_pass.Name = "TextBox_imap_pass"
        Me.TextBox_imap_pass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TextBox_imap_pass.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_imap_pass.TabIndex = 44
        Me.TextBox_imap_pass.UseSystemPasswordChar = True
        '
        'TextBox_imap_user
        '
        Me.TextBox_imap_user.Location = New System.Drawing.Point(68, 82)
        Me.TextBox_imap_user.Name = "TextBox_imap_user"
        Me.TextBox_imap_user.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_imap_user.TabIndex = 43
        '
        'TextBox_imap_server
        '
        Me.TextBox_imap_server.Location = New System.Drawing.Point(68, 30)
        Me.TextBox_imap_server.Name = "TextBox_imap_server"
        Me.TextBox_imap_server.Size = New System.Drawing.Size(160, 20)
        Me.TextBox_imap_server.TabIndex = 28
        '
        'Label_imap_pass
        '
        Me.Label_imap_pass.AutoSize = True
        Me.Label_imap_pass.Location = New System.Drawing.Point(27, 115)
        Me.Label_imap_pass.Name = "Label_imap_pass"
        Me.Label_imap_pass.Size = New System.Drawing.Size(30, 13)
        Me.Label_imap_pass.TabIndex = 41
        Me.Label_imap_pass.Text = "Pass"
        '
        'Label_imap_user
        '
        Me.Label_imap_user.AutoSize = True
        Me.Label_imap_user.Location = New System.Drawing.Point(27, 85)
        Me.Label_imap_user.Name = "Label_imap_user"
        Me.Label_imap_user.Size = New System.Drawing.Size(29, 13)
        Me.Label_imap_user.TabIndex = 39
        Me.Label_imap_user.Text = "User"
        '
        'Label_imap_server
        '
        Me.Label_imap_server.AutoSize = True
        Me.Label_imap_server.Location = New System.Drawing.Point(27, 33)
        Me.Label_imap_server.Name = "Label_imap_server"
        Me.Label_imap_server.Size = New System.Drawing.Size(38, 13)
        Me.Label_imap_server.TabIndex = 37
        Me.Label_imap_server.Text = "Server"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage_email)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Multiline = True
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1083, 541)
        Me.TabControl1.TabIndex = 26
        '
        'Timer_get_email_timeout_timer
        '
        Me.Timer_get_email_timeout_timer.Interval = 1000
        '
        'BackgroundWorker_get_email
        '
        Me.BackgroundWorker_get_email.WorkerReportsProgress = True
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Warning
        Me.NotifyIcon1.BalloonTipText = "Change Request Management System.  Do not close."
        Me.NotifyIcon1.BalloonTipTitle = "CRMS - Do Not Close"
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "CRMS - Do Not Close"
        Me.NotifyIcon1.Visible = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(9, 107)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(218, 13)
        Me.Label8.TabIndex = 46
        Me.Label8.Text = "client secret must be generated on the target"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(9, 128)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(223, 13)
        Me.Label9.TabIndex = 47
        Me.Label9.Text = "google drive account by enabling the .net API"
        '
        'BackgroundWorker_get_files_from_google_drive
        '
        Me.BackgroundWorker_get_files_from_google_drive.WorkerReportsProgress = True
        '
        'main_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(1083, 541)
        Me.Controls.Add(Me.TabControl1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "main_form"
        Me.Text = "CRMS (Change Request Management System)"
        Me.TabPage_email.ResumeLayout(False)
        Me.TabPage_email.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox_mysql_connect.ResumeLayout(False)
        Me.GroupBox_mysql_connect.PerformLayout()
        Me.GroupBox_settings.ResumeLayout(False)
        Me.GroupBox_settings.PerformLayout()
        Me.GroupBox_smtp.ResumeLayout(False)
        Me.GroupBox_smtp.PerformLayout()
        Me.GroupBox_pop.ResumeLayout(False)
        Me.GroupBox_pop.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Timer_check_mail As System.Windows.Forms.Timer
    Friend WithEvents TabPage_email As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox_smtp As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox_smtp_ssl As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox_smtp_port As System.Windows.Forms.TextBox
    Friend WithEvents Label_smtp_port As System.Windows.Forms.Label
    Friend WithEvents TextBox_smtp_pass As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_smtp_user As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_smtp_server As System.Windows.Forms.TextBox
    Friend WithEvents Label_smtp_pass As System.Windows.Forms.Label
    Friend WithEvents Label_smtp_user As System.Windows.Forms.Label
    Friend WithEvents Label_smtp_server As System.Windows.Forms.Label
    Friend WithEvents GroupBox_pop As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox_imap_ssl As System.Windows.Forms.CheckBox
    Friend WithEvents Button_base_browse As System.Windows.Forms.Button
    Friend WithEvents TextBox_base_path As System.Windows.Forms.TextBox
    Friend WithEvents Label_base_dl_dir As System.Windows.Forms.Label
    Friend WithEvents TextBox_imap_port As System.Windows.Forms.TextBox
    Friend WithEvents Label_imap_port As System.Windows.Forms.Label
    Friend WithEvents TextBox_imap_pass As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_imap_user As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_imap_server As System.Windows.Forms.TextBox
    Friend WithEvents Label_imap_pass As System.Windows.Forms.Label
    Friend WithEvents Label_imap_user As System.Windows.Forms.Label
    Friend WithEvents Label_imap_server As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents GroupBox_settings As System.Windows.Forms.GroupBox
    Friend WithEvents Button_settings_submit As System.Windows.Forms.Button
    Friend WithEvents Label_check_mail2 As System.Windows.Forms.Label
    Friend WithEvents TextBox_check_mail_period As System.Windows.Forms.TextBox
    Friend WithEvents Label_check_mail_period As System.Windows.Forms.Label
    Friend WithEvents Label_timer As System.Windows.Forms.Label
    Friend WithEvents GroupBox_mysql_connect As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox_mysql_port As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents TextBox_mysql_db As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TextBox_mysql_password As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_mysql_user As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_mysql_server As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label_cr_status_check_period As System.Windows.Forms.Label
    Friend WithEvents TextBox_cr_status_times As System.Windows.Forms.TextBox
    Friend WithEvents Timer_get_email_timeout_timer As System.Windows.Forms.Timer
    Friend WithEvents BackgroundWorker_get_email As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label_timeout_timer As System.Windows.Forms.Label
    Friend WithEvents TextBox_imap_encryption_type As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_smtp_timeout As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label_warning As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox_pop_use_ssl As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox_pop_port As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TextBox_pop_pass As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_pop_user As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_pop_server As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CheckBox_use_imap As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox_GD_clientsecret As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_GD_clientId As System.Windows.Forms.TextBox
    Friend WithEvents Labelx As System.Windows.Forms.Label
    Friend WithEvents Label_GD_clientId As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker_get_files_from_google_drive As System.ComponentModel.BackgroundWorker

End Class
