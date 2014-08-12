Imports Microsoft.Win32
Imports System.IO

Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim joblist As Integer
    Dim last_thread_run As Integer

    Dim WithEvents Worker1 As Worker
    Public Delegate Sub WorkerhHandler(ByVal Result As String)
    Public Delegate Sub WorkerProgresshHandler(ByVal FoldersCreated As Long)

    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False
    Private application_exit As Boolean = False
    Private shutting_down As Boolean = False

    Private InitialRootDirectory As String = ""
    Private RootDirectory As String = ""

    Private lastDepartments, lastCourses, lastAssignments As String
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PictureBox6 As System.Windows.Forms.PictureBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
    End Sub

    Public Sub New(ByVal splash As Splash_Screen)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents filefoldertxt As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents statsDepartment As System.Windows.Forms.Label
    Friend WithEvents lstDepartment As MTGCComboBox
    Friend WithEvents lstAssignments As MTGCComboBox
    Friend WithEvents lstCourses As MTGCComboBox
    Friend WithEvents statsCourses As System.Windows.Forms.Label
    Friend WithEvents statsAssignments As System.Windows.Forms.Label
    Friend WithEvents lblProceed As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents yearopen As System.Windows.Forms.TextBox
    Friend WithEvents monthopen As System.Windows.Forms.TextBox
    Friend WithEvents dayopen As System.Windows.Forms.TextBox
    Friend WithEvents houropen As System.Windows.Forms.TextBox
    Friend WithEvents minuteopen As System.Windows.Forms.TextBox
    Friend WithEvents minuteclose As System.Windows.Forms.TextBox
    Friend WithEvents hourclose As System.Windows.Forms.TextBox
    Friend WithEvents dayclose As System.Windows.Forms.TextBox
    Friend WithEvents monthclose As System.Windows.Forms.TextBox
    Friend WithEvents yearclose As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Button5 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main_Screen))
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.filefoldertxt = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.statsDepartment = New System.Windows.Forms.Label
        Me.lstDepartment = New MTGCComboBox
        Me.lstAssignments = New MTGCComboBox
        Me.lstCourses = New MTGCComboBox
        Me.statsCourses = New System.Windows.Forms.Label
        Me.statsAssignments = New System.Windows.Forms.Label
        Me.lblProceed = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.yearopen = New System.Windows.Forms.TextBox
        Me.monthopen = New System.Windows.Forms.TextBox
        Me.dayopen = New System.Windows.Forms.TextBox
        Me.houropen = New System.Windows.Forms.TextBox
        Me.minuteopen = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.minuteclose = New System.Windows.Forms.TextBox
        Me.hourclose = New System.Windows.Forms.TextBox
        Me.dayclose = New System.Windows.Forms.TextBox
        Me.monthclose = New System.Windows.Forms.TextBox
        Me.yearclose = New System.Windows.Forms.TextBox
        Me.Label28 = New System.Windows.Forms.Label
        Me.Label29 = New System.Windows.Forms.Label
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.Button6 = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PictureBox6 = New System.Windows.Forms.PictureBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label30 = New System.Windows.Forms.Label
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        Me.ImageList1.Images.SetKeyName(2, "")
        Me.ImageList1.Images.SetKeyName(3, "")
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ContextMenu1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Resting..."
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Display Main Screen"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "Exit Application"
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.White
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(320, 366)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(141, 53)
        Me.Button3.TabIndex = 82
        Me.Button3.Text = "Proceed"
        Me.ToolTip1.SetToolTip(Me.Button3, "Start operation")
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Gainsboro
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(256, 130)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(56, 20)
        Me.Button1.TabIndex = 80
        Me.Button1.Text = "SET"
        Me.ToolTip1.SetToolTip(Me.Button1, "Set the Root Folder")
        Me.Button1.UseVisualStyleBackColor = False
        '
        'filefoldertxt
        '
        Me.filefoldertxt.BackColor = System.Drawing.Color.White
        Me.filefoldertxt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.filefoldertxt.ForeColor = System.Drawing.Color.SteelBlue
        Me.filefoldertxt.Location = New System.Drawing.Point(16, 130)
        Me.filefoldertxt.Name = "filefoldertxt"
        Me.filefoldertxt.Size = New System.Drawing.Size(246, 20)
        Me.filefoldertxt.TabIndex = 79
        Me.ToolTip1.SetToolTip(Me.filefoldertxt, "Base folder to review filenames")
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.Transparent
        Me.Label19.ForeColor = System.Drawing.Color.Black
        Me.Label19.Location = New System.Drawing.Point(136, 56)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(144, 16)
        Me.Label19.TabIndex = 76
        Me.ToolTip1.SetToolTip(Me.Label19, "Time stamp of when operation was completed")
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.ForeColor = System.Drawing.Color.Black
        Me.Label15.Location = New System.Drawing.Point(8, 56)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(80, 16)
        Me.Label15.TabIndex = 75
        Me.Label15.Text = "Analysis End:"
        Me.ToolTip1.SetToolTip(Me.Label15, "Time stamp of when operation was completed")
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(136, 40)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(144, 16)
        Me.Label12.TabIndex = 74
        Me.Label12.Text = "0"
        Me.ToolTip1.SetToolTip(Me.Label12, "Number of files renamed")
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.ForeColor = System.Drawing.Color.Black
        Me.Label11.Location = New System.Drawing.Point(136, 24)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(144, 16)
        Me.Label11.TabIndex = 73
        Me.ToolTip1.SetToolTip(Me.Label11, "Time stamp of when operation was launched")
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(136, 8)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(144, 16)
        Me.Label10.TabIndex = 72
        Me.ToolTip1.SetToolTip(Me.Label10, "Time stamp of when application was opened")
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(8, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(136, 16)
        Me.Label6.TabIndex = 70
        Me.Label6.Text = "Student Folders Created:"
        Me.ToolTip1.SetToolTip(Me.Label6, "Student Folders Created")
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(8, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 16)
        Me.Label5.TabIndex = 69
        Me.Label5.Text = "Analysis Start:"
        Me.ToolTip1.SetToolTip(Me.Label5, "Time stamp of when operation was launched")
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(8, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(144, 16)
        Me.Label3.TabIndex = 68
        Me.Label3.Text = "Program Launched:"
        Me.ToolTip1.SetToolTip(Me.Label3, "Time stamp of when application was opened")
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(320, 130)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(56, 20)
        Me.Button2.TabIndex = 126
        Me.Button2.Text = "BROWSE"
        Me.ToolTip1.SetToolTip(Me.Button2, "Set the Root Folder")
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(200, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.LimeGreen
        Me.Label1.Location = New System.Drawing.Point(412, 428)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 16)
        Me.Label1.TabIndex = 66
        Me.Label1.Text = "Waiting"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox5
        '
        Me.PictureBox5.BackColor = System.Drawing.Color.FromArgb(CType(CType(200, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.PictureBox5.BackgroundImage = CType(resources.GetObject("PictureBox5.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox5.ForeColor = System.Drawing.Color.Black
        Me.PictureBox5.Location = New System.Drawing.Point(388, 428)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox5.TabIndex = 65
        Me.PictureBox5.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BackColor = System.Drawing.Color.FromArgb(CType(CType(200, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.PictureBox4.BackgroundImage = CType(resources.GetObject("PictureBox4.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox4.ForeColor = System.Drawing.Color.Black
        Me.PictureBox4.Location = New System.Drawing.Point(372, 428)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox4.TabIndex = 64
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackColor = System.Drawing.Color.FromArgb(CType(CType(200, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.PictureBox3.BackgroundImage = CType(resources.GetObject("PictureBox3.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox3.ForeColor = System.Drawing.Color.Black
        Me.PictureBox3.Location = New System.Drawing.Point(356, 428)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox3.TabIndex = 63
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(200, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox2.ForeColor = System.Drawing.Color.Black
        Me.PictureBox2.Location = New System.Drawing.Point(340, 428)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox2.TabIndex = 62
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(200, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.ForeColor = System.Drawing.Color.Black
        Me.PictureBox1.Location = New System.Drawing.Point(324, 428)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox1.TabIndex = 61
        Me.PictureBox1.TabStop = False
        '
        'Label24
        '
        Me.Label24.BackColor = System.Drawing.Color.Transparent
        Me.Label24.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.Color.Black
        Me.Label24.Location = New System.Drawing.Point(16, 114)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(192, 16)
        Me.Label24.TabIndex = 81
        Me.Label24.Text = "ROOT HAND-IN FOLDER"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'statsDepartment
        '
        Me.statsDepartment.ForeColor = System.Drawing.Color.DimGray
        Me.statsDepartment.Location = New System.Drawing.Point(288, 170)
        Me.statsDepartment.Name = "statsDepartment"
        Me.statsDepartment.Size = New System.Drawing.Size(320, 40)
        Me.statsDepartment.TabIndex = 90
        '
        'lstDepartment
        '
        Me.lstDepartment.BackColor = System.Drawing.Color.White
        Me.lstDepartment.BorderStyle = MTGCComboBox.TipiBordi.FlatXP
        Me.lstDepartment.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.lstDepartment.ColumnNum = 1
        Me.lstDepartment.ColumnWidth = "121"
        Me.lstDepartment.DisplayMember = "Text"
        Me.lstDepartment.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.lstDepartment.DropDownArrowBackColor = System.Drawing.Color.FromArgb(CType(CType(136, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.lstDepartment.DropDownBackColor = System.Drawing.Color.FromArgb(CType(CType(193, Byte), Integer), CType(CType(210, Byte), Integer), CType(CType(238, Byte), Integer))
        Me.lstDepartment.DropDownForeColor = System.Drawing.Color.Black
        Me.lstDepartment.DropDownStyle = MTGCComboBox.CustomDropDownStyle.DropDown
        Me.lstDepartment.DropDownWidth = 141
        Me.lstDepartment.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstDepartment.ForeColor = System.Drawing.Color.SteelBlue
        Me.lstDepartment.GridLineColor = System.Drawing.Color.Transparent
        Me.lstDepartment.GridLineHorizontal = False
        Me.lstDepartment.GridLineVertical = False
        Me.lstDepartment.HighlightBorderColor = System.Drawing.Color.Black
        Me.lstDepartment.HighlightBorderOnMouseEvents = True
        Me.lstDepartment.LoadingType = MTGCComboBox.CaricamentoCombo.ComboBoxItem
        Me.lstDepartment.Location = New System.Drawing.Point(16, 170)
        Me.lstDepartment.ManagingFastMouseMoving = True
        Me.lstDepartment.ManagingFastMouseMovingInterval = 30
        Me.lstDepartment.Name = "lstDepartment"
        Me.lstDepartment.NormalBorderColor = System.Drawing.Color.Black
        Me.lstDepartment.Size = New System.Drawing.Size(264, 21)
        Me.lstDepartment.TabIndex = 92
        Me.lstDepartment.TabStop = False
        '
        'lstAssignments
        '
        Me.lstAssignments.BackColor = System.Drawing.Color.White
        Me.lstAssignments.BorderStyle = MTGCComboBox.TipiBordi.FlatXP
        Me.lstAssignments.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.lstAssignments.ColumnNum = 1
        Me.lstAssignments.ColumnWidth = "121"
        Me.lstAssignments.DisplayMember = "Text"
        Me.lstAssignments.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.lstAssignments.DropDownArrowBackColor = System.Drawing.Color.FromArgb(CType(CType(136, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.lstAssignments.DropDownBackColor = System.Drawing.Color.FromArgb(CType(CType(193, Byte), Integer), CType(CType(210, Byte), Integer), CType(CType(238, Byte), Integer))
        Me.lstAssignments.DropDownForeColor = System.Drawing.Color.Black
        Me.lstAssignments.DropDownStyle = MTGCComboBox.CustomDropDownStyle.DropDown
        Me.lstAssignments.DropDownWidth = 141
        Me.lstAssignments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstAssignments.ForeColor = System.Drawing.Color.SteelBlue
        Me.lstAssignments.GridLineColor = System.Drawing.Color.Transparent
        Me.lstAssignments.GridLineHorizontal = False
        Me.lstAssignments.GridLineVertical = False
        Me.lstAssignments.HighlightBorderColor = System.Drawing.Color.Black
        Me.lstAssignments.HighlightBorderOnMouseEvents = True
        Me.lstAssignments.LoadingType = MTGCComboBox.CaricamentoCombo.ComboBoxItem
        Me.lstAssignments.Location = New System.Drawing.Point(16, 250)
        Me.lstAssignments.ManagingFastMouseMoving = True
        Me.lstAssignments.ManagingFastMouseMovingInterval = 30
        Me.lstAssignments.Name = "lstAssignments"
        Me.lstAssignments.NormalBorderColor = System.Drawing.Color.Black
        Me.lstAssignments.Size = New System.Drawing.Size(264, 21)
        Me.lstAssignments.TabIndex = 99
        Me.lstAssignments.TabStop = False
        '
        'lstCourses
        '
        Me.lstCourses.BackColor = System.Drawing.Color.White
        Me.lstCourses.BorderStyle = MTGCComboBox.TipiBordi.FlatXP
        Me.lstCourses.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.lstCourses.ColumnNum = 1
        Me.lstCourses.ColumnWidth = "121"
        Me.lstCourses.DisplayMember = "Text"
        Me.lstCourses.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.lstCourses.DropDownArrowBackColor = System.Drawing.Color.FromArgb(CType(CType(136, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.lstCourses.DropDownBackColor = System.Drawing.Color.FromArgb(CType(CType(193, Byte), Integer), CType(CType(210, Byte), Integer), CType(CType(238, Byte), Integer))
        Me.lstCourses.DropDownForeColor = System.Drawing.Color.Black
        Me.lstCourses.DropDownStyle = MTGCComboBox.CustomDropDownStyle.DropDown
        Me.lstCourses.DropDownWidth = 141
        Me.lstCourses.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstCourses.ForeColor = System.Drawing.Color.SteelBlue
        Me.lstCourses.GridLineColor = System.Drawing.Color.Transparent
        Me.lstCourses.GridLineHorizontal = False
        Me.lstCourses.GridLineVertical = False
        Me.lstCourses.HighlightBorderColor = System.Drawing.Color.Black
        Me.lstCourses.HighlightBorderOnMouseEvents = True
        Me.lstCourses.LoadingType = MTGCComboBox.CaricamentoCombo.ComboBoxItem
        Me.lstCourses.Location = New System.Drawing.Point(16, 210)
        Me.lstCourses.ManagingFastMouseMoving = True
        Me.lstCourses.ManagingFastMouseMovingInterval = 30
        Me.lstCourses.Name = "lstCourses"
        Me.lstCourses.NormalBorderColor = System.Drawing.Color.Black
        Me.lstCourses.Size = New System.Drawing.Size(264, 21)
        Me.lstCourses.TabIndex = 95
        Me.lstCourses.TabStop = False
        '
        'statsCourses
        '
        Me.statsCourses.ForeColor = System.Drawing.Color.DimGray
        Me.statsCourses.Location = New System.Drawing.Point(288, 210)
        Me.statsCourses.Name = "statsCourses"
        Me.statsCourses.Size = New System.Drawing.Size(320, 40)
        Me.statsCourses.TabIndex = 97
        '
        'statsAssignments
        '
        Me.statsAssignments.ForeColor = System.Drawing.Color.DimGray
        Me.statsAssignments.Location = New System.Drawing.Point(288, 250)
        Me.statsAssignments.Name = "statsAssignments"
        Me.statsAssignments.Size = New System.Drawing.Size(320, 29)
        Me.statsAssignments.TabIndex = 101
        '
        'lblProceed
        '
        Me.lblProceed.ForeColor = System.Drawing.Color.DarkGray
        Me.lblProceed.Location = New System.Drawing.Point(12, 460)
        Me.lblProceed.Name = "lblProceed"
        Me.lblProceed.Size = New System.Drawing.Size(610, 24)
        Me.lblProceed.TabIndex = 121
        Me.lblProceed.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(16, 154)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(192, 16)
        Me.Label4.TabIndex = 123
        Me.Label4.Text = "DEPARTMENT FOLDER"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.Transparent
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Black
        Me.Label17.Location = New System.Drawing.Point(16, 194)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(192, 16)
        Me.Label17.TabIndex = 124
        Me.Label17.Text = "COURSE FOLDER"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(16, 234)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(192, 16)
        Me.Label13.TabIndex = 125
        Me.Label13.Text = "PROJECT TITLE / FOLDER"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.Description = "Select the Hand-in root directory"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Azure
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label19)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Location = New System.Drawing.Point(12, 366)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(288, 82)
        Me.Panel1.TabIndex = 127
        '
        'yearopen
        '
        Me.yearopen.BackColor = System.Drawing.Color.White
        Me.yearopen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.yearopen.ForeColor = System.Drawing.Color.SteelBlue
        Me.yearopen.Location = New System.Drawing.Point(103, 288)
        Me.yearopen.MaxLength = 4
        Me.yearopen.Name = "yearopen"
        Me.yearopen.Size = New System.Drawing.Size(32, 20)
        Me.yearopen.TabIndex = 129
        Me.yearopen.Text = "2006"
        '
        'monthopen
        '
        Me.monthopen.BackColor = System.Drawing.Color.White
        Me.monthopen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.monthopen.ForeColor = System.Drawing.Color.SteelBlue
        Me.monthopen.Location = New System.Drawing.Point(135, 288)
        Me.monthopen.MaxLength = 2
        Me.monthopen.Name = "monthopen"
        Me.monthopen.Size = New System.Drawing.Size(24, 20)
        Me.monthopen.TabIndex = 130
        Me.monthopen.Text = "22"
        '
        'dayopen
        '
        Me.dayopen.BackColor = System.Drawing.Color.White
        Me.dayopen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.dayopen.ForeColor = System.Drawing.Color.SteelBlue
        Me.dayopen.Location = New System.Drawing.Point(159, 288)
        Me.dayopen.MaxLength = 2
        Me.dayopen.Name = "dayopen"
        Me.dayopen.Size = New System.Drawing.Size(24, 20)
        Me.dayopen.TabIndex = 131
        Me.dayopen.Text = "33"
        '
        'houropen
        '
        Me.houropen.BackColor = System.Drawing.Color.White
        Me.houropen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.houropen.ForeColor = System.Drawing.Color.SteelBlue
        Me.houropen.Location = New System.Drawing.Point(191, 288)
        Me.houropen.MaxLength = 2
        Me.houropen.Name = "houropen"
        Me.houropen.Size = New System.Drawing.Size(24, 20)
        Me.houropen.TabIndex = 132
        Me.houropen.Text = "08"
        '
        'minuteopen
        '
        Me.minuteopen.BackColor = System.Drawing.Color.White
        Me.minuteopen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.minuteopen.ForeColor = System.Drawing.Color.SteelBlue
        Me.minuteopen.Location = New System.Drawing.Point(215, 288)
        Me.minuteopen.MaxLength = 2
        Me.minuteopen.Name = "minuteopen"
        Me.minuteopen.Size = New System.Drawing.Size(24, 20)
        Me.minuteopen.TabIndex = 133
        Me.minuteopen.Text = "44"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(103, 304)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 16)
        Me.Label2.TabIndex = 134
        Me.Label2.Text = "YEAR"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Black
        Me.Label14.Location = New System.Drawing.Point(159, 304)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(24, 16)
        Me.Label14.TabIndex = 136
        Me.Label14.Text = "DAY"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.Color.Transparent
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Black
        Me.Label18.Location = New System.Drawing.Point(191, 304)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(24, 16)
        Me.Label18.TabIndex = 137
        Me.Label18.Text = "HH"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.Color.Transparent
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Black
        Me.Label20.Location = New System.Drawing.Point(215, 304)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(24, 16)
        Me.Label20.TabIndex = 138
        Me.Label20.Text = "MM"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label21
        '
        Me.Label21.BackColor = System.Drawing.Color.Transparent
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.Black
        Me.Label21.Location = New System.Drawing.Point(449, 304)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(24, 16)
        Me.Label21.TabIndex = 148
        Me.Label21.Text = "MM"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label22
        '
        Me.Label22.BackColor = System.Drawing.Color.Transparent
        Me.Label22.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.Color.Black
        Me.Label22.Location = New System.Drawing.Point(425, 304)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(24, 16)
        Me.Label22.TabIndex = 147
        Me.Label22.Text = "HH"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label23
        '
        Me.Label23.BackColor = System.Drawing.Color.Transparent
        Me.Label23.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.ForeColor = System.Drawing.Color.Black
        Me.Label23.Location = New System.Drawing.Point(393, 304)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(24, 16)
        Me.Label23.TabIndex = 146
        Me.Label23.Text = "DAY"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.Color.Transparent
        Me.Label26.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.Black
        Me.Label26.Location = New System.Drawing.Point(337, 304)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(32, 16)
        Me.Label26.TabIndex = 144
        Me.Label26.Text = "YEAR"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'minuteclose
        '
        Me.minuteclose.BackColor = System.Drawing.Color.White
        Me.minuteclose.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.minuteclose.ForeColor = System.Drawing.Color.SteelBlue
        Me.minuteclose.Location = New System.Drawing.Point(449, 288)
        Me.minuteclose.MaxLength = 2
        Me.minuteclose.Name = "minuteclose"
        Me.minuteclose.Size = New System.Drawing.Size(24, 20)
        Me.minuteclose.TabIndex = 143
        Me.minuteclose.Text = "44"
        '
        'hourclose
        '
        Me.hourclose.BackColor = System.Drawing.Color.White
        Me.hourclose.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.hourclose.ForeColor = System.Drawing.Color.SteelBlue
        Me.hourclose.Location = New System.Drawing.Point(425, 288)
        Me.hourclose.MaxLength = 2
        Me.hourclose.Name = "hourclose"
        Me.hourclose.Size = New System.Drawing.Size(24, 20)
        Me.hourclose.TabIndex = 142
        Me.hourclose.Text = "08"
        '
        'dayclose
        '
        Me.dayclose.BackColor = System.Drawing.Color.White
        Me.dayclose.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.dayclose.ForeColor = System.Drawing.Color.SteelBlue
        Me.dayclose.Location = New System.Drawing.Point(393, 288)
        Me.dayclose.MaxLength = 2
        Me.dayclose.Name = "dayclose"
        Me.dayclose.Size = New System.Drawing.Size(24, 20)
        Me.dayclose.TabIndex = 141
        Me.dayclose.Text = "33"
        '
        'monthclose
        '
        Me.monthclose.BackColor = System.Drawing.Color.White
        Me.monthclose.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.monthclose.ForeColor = System.Drawing.Color.SteelBlue
        Me.monthclose.Location = New System.Drawing.Point(369, 288)
        Me.monthclose.MaxLength = 2
        Me.monthclose.Name = "monthclose"
        Me.monthclose.Size = New System.Drawing.Size(24, 20)
        Me.monthclose.TabIndex = 140
        Me.monthclose.Text = "22"
        '
        'yearclose
        '
        Me.yearclose.BackColor = System.Drawing.Color.White
        Me.yearclose.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.yearclose.ForeColor = System.Drawing.Color.SteelBlue
        Me.yearclose.Location = New System.Drawing.Point(337, 288)
        Me.yearclose.MaxLength = 4
        Me.yearclose.Name = "yearclose"
        Me.yearclose.Size = New System.Drawing.Size(32, 20)
        Me.yearclose.TabIndex = 139
        Me.yearclose.Text = "2006"
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.Color.Transparent
        Me.Label28.ForeColor = System.Drawing.Color.Black
        Me.Label28.Location = New System.Drawing.Point(15, 296)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(80, 16)
        Me.Label28.TabIndex = 150
        Me.Label28.Text = "Open Window:"
        '
        'Label29
        '
        Me.Label29.BackColor = System.Drawing.Color.Transparent
        Me.Label29.ForeColor = System.Drawing.Color.Black
        Me.Label29.Location = New System.Drawing.Point(249, 296)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(80, 16)
        Me.Label29.TabIndex = 151
        Me.Label29.Text = "Close Window:"
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Checked = True
        Me.RadioButton2.Location = New System.Drawing.Point(16, 70)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(103, 17)
        Me.RadioButton2.TabIndex = 158
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Windows Server"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(15, 93)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(89, 17)
        Me.RadioButton1.TabIndex = 157
        Me.RadioButton1.Text = "Novell Server"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(133, 304)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(32, 16)
        Me.Label7.TabIndex = 152
        Me.Label7.Text = "MTH"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.Color.Transparent
        Me.Label25.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.Color.Black
        Me.Label25.Location = New System.Drawing.Point(367, 304)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(32, 16)
        Me.Label25.TabIndex = 153
        Me.Label25.Text = "MTH"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.White
        Me.Button5.Location = New System.Drawing.Point(484, 366)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(136, 24)
        Me.Button5.TabIndex = 153
        Me.Button5.Text = "Refresh Student Folders"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.White
        Me.Button4.Location = New System.Drawing.Point(484, 425)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(136, 23)
        Me.Button4.TabIndex = 154
        Me.Button4.Text = "Kill Processes"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.SteelBlue
        Me.Label8.Location = New System.Drawing.Point(506, 133)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(116, 17)
        Me.Label8.TabIndex = 156
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.White
        Me.Button6.Location = New System.Drawing.Point(484, 396)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(136, 23)
        Me.Button6.TabIndex = 157
        Me.Button6.Text = "Batch Create"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "txt"
        Me.OpenFileDialog1.FileName = "InputFile"
        Me.OpenFileDialog1.Filter = "Text Files|*.txt|All Files|*.*"
        Me.OpenFileDialog1.Title = "Please select the input text file"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HelpToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(634, 24)
        Me.MenuStrip1.TabIndex = 158
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(40, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(48, 20)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'PictureBox6
        '
        Me.PictureBox6.Image = Global.Student_Handin_Project_Creator.My.Resources.Resources.Splash_Image
        Me.PictureBox6.Location = New System.Drawing.Point(181, 24)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(453, 106)
        Me.PictureBox6.TabIndex = 0
        Me.PictureBox6.TabStop = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.Color.SteelBlue
        Me.Label9.Location = New System.Drawing.Point(397, 133)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(103, 13)
        Me.Label9.TabIndex = 159
        Me.Label9.Text = "Current Time Stamp:"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(12, 41)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(64, 16)
        Me.Label16.TabIndex = 160
        Me.Label16.Text = "Settings"
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.Location = New System.Drawing.Point(12, 338)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(83, 16)
        Me.Label30.TabIndex = 161
        Me.Label30.Text = "Controllers"
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(200, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(634, 492)
        Me.Controls.Add(Me.Label30)
        Me.Controls.Add(Me.dayclose)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.minuteclose)
        Me.Controls.Add(Me.hourclose)
        Me.Controls.Add(Me.monthclose)
        Me.Controls.Add(Me.yearclose)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.RadioButton1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.yearopen)
        Me.Controls.Add(Me.PictureBox5)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.PictureBox4)
        Me.Controls.Add(Me.monthopen)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.dayopen)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.houropen)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.minuteopen)
        Me.Controls.Add(Me.PictureBox6)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.lblProceed)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.filefoldertxt)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.statsDepartment)
        Me.Controls.Add(Me.Label25)
        Me.Controls.Add(Me.lstDepartment)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.lstAssignments)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.lstCourses)
        Me.Controls.Add(Me.statsCourses)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.statsAssignments)
        Me.Controls.Add(Me.Label29)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "Main_Screen"
        Me.ShowInTaskbar = False
        Me.Text = "Student Handin Project Creator"
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    Private current_light As Integer = 0
    Private current_colour As Integer = 0
    Private currently_working As Boolean = False
    ' Private continuenextbatch As Boolean = False
    '    Private filefolder As String


    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Student Handin Project Creator's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Public Sub Load_Registry_Values()
        Try
            Dim configflag As Boolean
            configflag = False
            Dim str As String
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim keys() As String = oReg.GetSubKeyNames()
            System.Array.Sort(keys)

            For Each str In keys
                If str.Equals("Software\CodeUnit\Student Handin Project Creator") = True Then
                    keyflag1 = True
                    Exit For
                End If
            Next str

            If keyflag1 = False Then
                oReg.CreateSubKey("Software\CodeUnit\Student Handin Project Creator")
            End If

            keyflag1 = False

            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\CodeUnit\Student Handin Project Creator", True)

            str = oKey.GetValue("RootDirectory")
            If Not IsNothing(str) And Not (str = "") Then
                filefoldertxt.Text = str
            Else
                configflag = True
                oKey.SetValue("RootDirectory", "\\Comlab\Vol2\handin")
                filefoldertxt.Text = "\\Comlab\Vol2\handin"
            End If


            str = oKey.GetValue("lstDepartments")
            If Not IsNothing(str) And Not (str = "") Then
                lastDepartments = str
            Else
                configflag = True
                oKey.SetValue("lstDepartments", "")
                lastDepartments = ""
            End If

            str = oKey.GetValue("lstCourses")
            If Not IsNothing(str) And Not (str = "") Then
                lastCourses = str
            Else
                configflag = True
                oKey.SetValue("lstCourses", "")
                lastCourses = ""
            End If



            str = oKey.GetValue("lstAssignments")
            If Not IsNothing(str) And Not (str = "") Then
                lastAssignments = str
            Else
                configflag = True
                oKey.SetValue("lstAssignments", "")
                lastAssignments = ""
            End If
            oKey.Close()
            oReg.Close()

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Save_Registry_Values()
        Try
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\CodeUnit\Student Handin Project Creator", True)

            oKey.SetValue("RootDirectory", filefoldertxt.Text)
            oKey.SetValue("lstDepartments", lstDepartment.Text)
            oKey.SetValue("lstCourses", lstCourses.Text)
            oKey.SetValue("lstAssignments", lstAssignments.Text)

            oKey.Close()
            oReg.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_green_lights()
        Try
            Label1.ForeColor = Color.LimeGreen
            Label1.Text = "Waiting"
            '  Label7.Text = "Resting..."

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 0
            PictureBox1.Image = ImageList1.Images(1)
            PictureBox2.Image = ImageList1.Images(1)
            PictureBox3.Image = ImageList1.Images(1)
            PictureBox4.Image = ImageList1.Images(1)
            PictureBox5.Image = ImageList1.Images(1)

            Select Case current_light
                Case 0

                    PictureBox1.Image = ImageList1.Images(0)
                Case 1

                    PictureBox2.Image = ImageList1.Images(0)
                Case 2

                    PictureBox3.Image = ImageList1.Images(0)
                Case 3

                    PictureBox4.Image = ImageList1.Images(0)
                Case 4

                    PictureBox5.Image = ImageList1.Images(0)
                Case 5

                    PictureBox1.Image = ImageList1.Images(0)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_orange_lights()
        Try
            Label1.ForeColor = Color.DarkOrange
            Label1.Text = "Working"

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 1
            PictureBox1.Image = ImageList1.Images(3)
            PictureBox2.Image = ImageList1.Images(3)
            PictureBox3.Image = ImageList1.Images(3)
            PictureBox4.Image = ImageList1.Images(3)
            PictureBox5.Image = ImageList1.Images(3)
            Select Case current_light
                Case 0
                    PictureBox1.Image = ImageList1.Images(2)
                Case 1
                    PictureBox2.Image = ImageList1.Images(2)
                Case 2
                    PictureBox3.Image = ImageList1.Images(2)
                Case 3
                    PictureBox4.Image = ImageList1.Images(2)
                Case 4
                    PictureBox5.Image = ImageList1.Images(2)
                Case 5
                    PictureBox1.Image = ImageList1.Images(2)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_lights()
        Try
            If current_colour = 1 Then
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(3)
                        PictureBox2.Image = ImageList1.Images(2)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(3)
                        PictureBox3.Image = ImageList1.Images(2)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(3)
                        PictureBox4.Image = ImageList1.Images(2)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(3)
                        PictureBox5.Image = ImageList1.Images(2)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                End Select
            Else
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(1)
                        PictureBox2.Image = ImageList1.Images(0)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(1)
                        PictureBox3.Image = ImageList1.Images(0)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(1)
                        PictureBox4.Image = ImageList1.Images(0)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(1)
                        PictureBox5.Image = ImageList1.Images(0)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                End Select

            End If

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            run_lights()
            Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
            If application_exit = True Then
                dataloaded = True
                splash_loader.Visible = False
                Me.Close()
            End If
            If joblist > 0 Then
                If currently_working = False Then
                    'Threading.Thread.Sleep(1000) 
                    force_check()
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.Text = My.Application.Info.ProductName & " " & My.Application.Info.Version.Major & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ""
            Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
            Label10.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
            Timer2.Start()
            ' dataloaded = True

            application_exit = False
            'Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
            'Timer2.Start()

            yearopen.Text = Format(Now(), "yyyy")
            yearclose.Text = Format(Now(), "yyyy")
            monthopen.Text = "01"
            monthclose.Text = "12"
            dayopen.Text = "01"
            dayclose.Text = "31"
            houropen.Text = "08"
            hourclose.Text = "16"
            minuteopen.Text = "00"
            minuteclose.Text = "00"


            Dim rekey As RegistryKey
            Try
                rekey = My.Computer.Registry.LocalMachine.OpenSubKey("Software\CodeUnit\Student Handin Project Creator", False)

                If rekey Is Nothing Then
                    rekey = My.Computer.Registry.LocalMachine.CreateSubKey("Software\CodeUnit\Student Handin Project Creator")
                    rekey = My.Computer.Registry.LocalMachine.OpenSubKey("Software\CodeUnit\Student Handin Project Creator", False)
                End If

            Catch ex As Exception
                rekey = My.Computer.Registry.LocalMachine.CreateSubKey("Software\CodeUnit\Student Handin Project Creator")
                rekey = My.Computer.Registry.LocalMachine.OpenSubKey("Software\CodeUnit\Student Handin Project Creator", False)

            End Try
            Dim radionow As String = rekey.GetValue("ServerType")
            If radionow Is Nothing Then
                radionow = "True"
            End If
            If radionow = "True" Then
                RadioButton2.Checked = True
                RadioButton1.Checked = False
            Else
                RadioButton2.Checked = False
                RadioButton1.Checked = True
            End If

            RootDirectory = (rekey.GetValue("RootDirectory"))
            rekey.Close()
            rekey = Nothing
            If RootDirectory Is Nothing Then
                RootDirectory = ""
            End If
            If RootDirectory.StartsWith("Fail") = True Then

                Dim Display_Message1 As New Display_Message("An usable Handin directory cannot be located.")
                Display_Message1.ShowDialog()
                'application_exit = True
            End If

            'RootDirectory = "C:\inetpub\WWWROOT\Services\Student Handin System\Handin"
            InitialRootDirectory = RootDirectory
            filefoldertxt.Text = InitialRootDirectory
            If RootDirectory.StartsWith("\\") Then

                RootDirectory = Worker1.MapDrive(RootDirectory)
                If Not (RootDirectory).IndexOf("Fail") = -1 Then
                    Dim Display_Message1 As New Display_Message("An usable Handin directory cannot be located.")
                    Display_Message1.ShowDialog()
                    'application_exit = True
                End If
            End If


            ' MsgBox(RootDirectory)
            If My.Computer.FileSystem.DirectoryExists(RootDirectory) Then


                Dim dinfo As DirectoryInfo = New DirectoryInfo(RootDirectory)
                Dim yearfound As Boolean = False
                Dim dinfo2 As DirectoryInfo
                For Each dinfo2 In dinfo.GetDirectories
                    If dinfo2.Name = Format(Now(), "yyyy") Then
                        yearfound = True
                        Exit For
                    End If
                Next
                If yearfound = False Then
                    dinfo = New DirectoryInfo(RootDirectory)
                    dinfo.CreateSubdirectory(Format(Now(), "yyyy"))

                End If
                lstDepartment.Enabled = True
                dinfo = New DirectoryInfo((RootDirectory & "\").Replace("\\", "\") & Format(Now(), "yyyy"))
                lstDepartment.Items.Clear()
                lstDepartment.Text = ""
                For Each dinfo2 In dinfo.GetDirectories
                    lstDepartment.Items.Add(New MTGCComboBoxItem(dinfo2.Name))
                Next
                If lstDepartment.Items.Count > 0 Then
                    lstDepartment.Enabled = True
                    lstDepartment.SelectedIndex = 0
                Else

                    statsDepartment.Text = ""

                    statsCourses.Text = ""

                    statsAssignments.Text = ""

                End If
                If lstDepartment.Items.Count = 1 Then
                    statsDepartment.Text = lstDepartment.Items.Count & " Department available to choose from."
                Else
                    statsDepartment.Text = lstDepartment.Items.Count & " Departments available to choose from."
                End If
                dinfo = Nothing
                dinfo2 = Nothing
                Load_Registry_Values()
                lstDepartment.Text = lastDepartments
                lstCourses.Text = lastCourses
                lstAssignments.Text = lastAssignments
            Else
                Dim Display_Message1 As New Display_Message("An usable Handin directory cannot be located.")
                Display_Message1.ShowDialog()
            End If
            dataloaded = True
            '**splash_loader.Opacity = 0
            '**splash_loader.WindowState = FormWindowState.Minimized
            '**splash_loader.Visible = False

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub exit_application()

        Try

            shutting_down = True
            Me.Opacity = 0
            Me.WindowState = FormWindowState.Minimized
            Me.Visible = False
            Timer2.Stop()
            If RootDirectory.IndexOf("Fail") = -1 Then


                If InitialRootDirectory.StartsWith("\\") Then
                    RootDirectory = Worker1.UnMapDrive(RootDirectory)
                End If
            End If
            If Worker1.WorkerThread Is Nothing = False Then
                Worker1.WorkerThread.Abort()
                Worker1.Dispose()
            End If
            NotifyIcon1.Dispose()
            Save_Registry_Values()

            Dim rekey As RegistryKey
            Try
                rekey = My.Computer.Registry.LocalMachine.OpenSubKey("Software\CodeUnit\Student Handin Project Creator", True)

                If rekey Is Nothing Then
                    rekey = My.Computer.Registry.LocalMachine.CreateSubKey("Software\CodeUnit\Student Handin Project Creator")
                    rekey = My.Computer.Registry.LocalMachine.OpenSubKey("Software\CodeUnit\Student Handin Project Creator", True)
                End If

            Catch ex As Exception
                rekey = My.Computer.Registry.LocalMachine.CreateSubKey("Software\CodeUnit\Student Handin Project Creator")
                rekey = My.Computer.Registry.LocalMachine.OpenSubKey("Software\CodeUnit\Student Handin Project Creator", True)

            End Try
            If RadioButton2.Checked = True Then
                rekey.SetValue("ServerType", "True")
            Else
                rekey.SetValue("ServerType", "False")
            End If

            rekey.Close()
            rekey = Nothing


            Application.Exit()
        Catch ex As Exception
            Error_Handler(ex, "Shutting Down Application")
        Finally
            Application.Exit()
        End Try
    End Sub

    Private Sub Main_Screen_closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            exit_application()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Public Sub WorkerHandler(ByVal Result As String)
        Try

            If joblist > 0 Then
                joblist = joblist - 1
            End If
            'MsgBox("Worker Handler: Work Complete, joblist = " & joblist)

            currently_working = False
            ' continuenextbatch = True
            Label19.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
            NotifyIcon1.Text = "Resting... "
            LoadCourseFolders((RootDirectory & "\" & Format(Now, "yyyy") & "\" & lstDepartment.Text).Replace("\\", "\"))
            LoadAssignmentFolders((RootDirectory & "\" & Format(Now, "yyyy") & "\" & lstDepartment.Text & "\" & lstCourses.Text).Replace("\\", "\"))
            'Label12.Text = Worker1.foldercount
            run_green_lights()



            'If joblist > 0 Then
            '    Threading.Thread.Sleep(1000) 'wait 30 seconds
            '    force_check()
            'End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerProgressHandler(ByVal FoldersCreated As Long)
        Try
            Label12.Text = (FoldersCreated).ToString

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_worker(Optional ByVal refresh As Boolean = False)

        run_orange_lights()
        Label11.Text = ""
        Label12.Text = ""
        Label19.Text = ""
        Label11.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")

        Worker1.refresh = refresh
        Worker1.rootdirectory = RootDirectory
        Worker1.department = lstDepartment.Text
        Worker1.course = lstCourses.Text
        Worker1.project = lstAssignments.Text
        Worker1.opendate = New Date(CInt(yearopen.Text), CInt(monthopen.Text), CInt(dayopen.Text), CInt(houropen.Text), CInt(minuteopen.Text), 0)
        Worker1.closedate = New Date(CInt(yearclose.Text), CInt(monthclose.Text), CInt(dayclose.Text), CInt(hourclose.Text), CInt(minuteclose.Text), 0)
        If RadioButton1.Checked = True Then
            Worker1.novellserver = True
        Else
            Worker1.novellserver = False
        End If
        'MsgBox("Run Worker: Choose Thread")
        If joblist > 0 Then
            Worker1.ChooseThreads(last_thread_run)
            If last_thread_run = 1 Then
                last_thread_run = 2
            Else
                last_thread_run = 1
            End If
        Else
            Worker1.ChooseThreads(1)
        End If

        currently_working = True
    End Sub



    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub show_application()
        Try
            Me.Opacity = 1

            Me.BringToFront()
            Me.Refresh()
            Me.WindowState = FormWindowState.Normal

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub NotifyIcon1_dblclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        show_application()
    End Sub
    Private Sub NotifyIcon1_snglclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        show_application()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        show_application()
    End Sub

    Private Sub Main_Screen_resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try

            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                Me.Opacity = 0
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub force_check()

        Try

            'If Worker1.WorkerThread Is Nothing = False Then
            '    Worker1.WorkerThread.Abort()
            '    Worker1.Dispose()
            'End If

            If joblist > 0 Then



                Dim reader As StreamReader = New StreamReader(OpenFileDialog1.FileName, False)
                Dim inputline As String = ""
                Dim createfolders As Boolean = False

                'MsgBox("Force Check: joblist = " & joblist)
                For i As Integer = 1 To joblist
                    inputline = reader.ReadLine
                Next
                'MsgBox("Force Check: inputline = " & inputline)

                reader.Close()
                reader = Nothing


                lstDepartment.SelectedIndex = -1
                lstCourses.SelectedIndex = -1
                lstAssignments.SelectedIndex = -1
                yearopen.Text = ""
                monthopen.Text = ""
                dayopen.Text = ""
                houropen.Text = ""
                minuteopen.Text = ""
                yearclose.Text = ""
                monthclose.Text = ""
                dayclose.Text = ""
                hourclose.Text = ""
                minuteclose.Text = ""
                createfolders = False

                If inputline.Length > 0 Then
                    Dim inputs As String() = inputline.Split(";")
                    If (inputs.Length) = 5 Then

                        If lstDepartment.Items.Contains(inputs(0)) = True Then
                            lstDepartment.SelectedIndex = lstDepartment.Items.IndexOf(inputs(0))
                            lstDepartment.Text = inputs(0)
                        Else
                            lstDepartment.Items.Add(inputs(0))
                            lstDepartment.SelectedIndex = lstDepartment.Items.IndexOf(inputs(0))
                            lstDepartment.Text = inputs(0)
                        End If
                        If lstCourses.Items.Contains(inputs(1)) = True Then
                            lstCourses.SelectedIndex = lstCourses.Items.IndexOf(inputs(1))
                            lstCourses.Text = inputs(1)
                        Else
                            lstCourses.Items.Add(inputs(1))
                            lstCourses.SelectedIndex = lstCourses.Items.IndexOf(inputs(1))
                            lstCourses.Text = inputs(1)
                        End If
                        If lstAssignments.Items.Contains(inputs(2)) = True Then
                            lstAssignments.SelectedIndex = lstAssignments.Items.IndexOf(inputs(2))
                            lstAssignments.Text = inputs(2)
                        Else
                            lstAssignments.Items.Add(inputs(2))
                            lstAssignments.SelectedIndex = lstAssignments.Items.IndexOf(inputs(2))
                            lstAssignments.Text = inputs(2)
                        End If
                        If inputs(3).Length = 12 Then
                            yearopen.Text = inputs(3).Substring(0, 4)
                            monthopen.Text = inputs(3).Substring(4, 2)
                            dayopen.Text = inputs(3).Substring(6, 2)
                            houropen.Text = inputs(3).Substring(8, 2)
                            minuteopen.Text = inputs(3).Substring(10, 2)
                        End If
                        If inputs(4).Length = 12 Then
                            yearclose.Text = inputs(4).Substring(0, 4)
                            monthclose.Text = inputs(4).Substring(4, 2)
                            dayclose.Text = inputs(4).Substring(6, 2)
                            hourclose.Text = inputs(4).Substring(8, 2)
                            minuteclose.Text = inputs(4).Substring(10, 2)
                        End If
                        If lstDepartment.Text.Length > 0 And lstCourses.Text.Length > 0 And lstAssignments.Text.Length > 0 Then
                            If yearopen.Text.Length = 4 And yearclose.Text.Length = 4 Then
                                If monthopen.Text.Length = 2 And dayopen.Text.Length = 2 And houropen.Text.Length = 2 And minuteopen.Text.Length = 2 Then
                                    If monthclose.Text.Length = 2 And dayclose.Text.Length = 2 And hourclose.Text.Length = 2 And minuteclose.Text.Length = 2 Then
                                        createfolders = True
                                    End If
                                End If
                            End If
                        End If

                    End If
                End If
                If createfolders = True Then
                    Dim opendate As Date
                    Try
                        opendate = New Date(CInt(yearopen.Text), CInt(monthopen.Text), CInt(dayopen.Text), CInt(houropen.Text), CInt(minuteopen.Text), 0)
                    Catch ex As Exception
                        MsgBox("You need to enter appropriate values for the opening time/date value. """ & (yearopen.Text) & "/" & (monthopen.Text) & "/" & (dayopen.Text) & " " & (houropen.Text) & ":" & (minuteopen.Text) & ":" & "00"" is not a valid date/time object", MsgBoxStyle.Information, "Invalid Inputs Detected")
                        Exit Sub
                    End Try
                    Dim closedate As Date
                    Try
                        closedate = New Date(CInt(yearclose.Text), CInt(monthclose.Text), CInt(dayclose.Text), CInt(hourclose.Text), CInt(minuteclose.Text), 0)
                    Catch ex As Exception
                        MsgBox("You need to enter appropriate values for the closing time/date value. """ & (yearclose.Text) & "/" & (monthclose.Text) & "/" & (dayclose.Text) & " " & (hourclose.Text) & ":" & (minuteclose.Text) & ":" & "00"" is not a valid date/time object", MsgBoxStyle.Information, "Invalid Inputs Detected")
                        Exit Sub
                    End Try

                    If closedate < opendate Then
                        MsgBox("The opening date cannot be less that the closing date. Please ensure you have filled these fields in correctly", MsgBoxStyle.Information, "Invalid Inputs Detected")
                        Exit Sub
                    End If
                    If lstDepartment.Text.Length > 0 And lstCourses.Text.Length > 0 And lstAssignments.Text.Length > 0 Then


                        NotifyIcon1.Text = "Generating Folders..."
                        'MsgBox("Force Check: Currently_Working " & currently_working)
                        If currently_working = False Then
                            'MsgBox("Force Check: Run Job")
                            'run_worker(True)
                            run_worker(False)
                        End If
                    Else
                        MsgBox("You need to enter values in the Department, Course and Project fields in order to generate a valid project folder", MsgBoxStyle.Information, "Invalid Inputs Detected")
                    End If
                Else
                    If joblist > 0 Then
                        joblist = joblist - 1
                        force_check()
                    End If
                End If
            Else
                Dim opendate As Date
                Try
                    opendate = New Date(CInt(yearopen.Text), CInt(monthopen.Text), CInt(dayopen.Text), CInt(houropen.Text), CInt(minuteopen.Text), 0)
                Catch ex As Exception
                    MsgBox("You need to enter appropriate values for the opening time/date value. """ & (yearopen.Text) & "/" & (monthopen.Text) & "/" & (dayopen.Text) & " " & (houropen.Text) & ":" & (minuteopen.Text) & ":" & "00"" is not a valid date/time object", MsgBoxStyle.Information, "Invalid Inputs Detected")
                    Exit Sub
                End Try
                Dim closedate As Date
                Try
                    closedate = New Date(CInt(yearclose.Text), CInt(monthclose.Text), CInt(dayclose.Text), CInt(hourclose.Text), CInt(minuteclose.Text), 0)
                Catch ex As Exception
                    MsgBox("You need to enter appropriate values for the closing time/date value. """ & (yearclose.Text) & "/" & (monthclose.Text) & "/" & (dayclose.Text) & " " & (hourclose.Text) & ":" & (minuteclose.Text) & ":" & "00"" is not a valid date/time object", MsgBoxStyle.Information, "Invalid Inputs Detected")
                    Exit Sub
                End Try

                If closedate < opendate Then
                    MsgBox("The opening date cannot be less that the closing date. Please ensure you have filled these fields in correctly", MsgBoxStyle.Information, "Invalid Inputs Detected")
                    Exit Sub
                End If
                If lstDepartment.Text.Length > 0 And lstCourses.Text.Length > 0 And lstAssignments.Text.Length > 0 Then


                    NotifyIcon1.Text = "Generating Folders..."
                    If currently_working = False Then
                        run_worker()
                    End If
                Else
                    MsgBox("You need to enter values in the Department, Course and Project fields in order to generate a valid project folder", MsgBoxStyle.Information, "Invalid Inputs Detected")
                End If
            End If





        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        force_check()
    End Sub

    Private Sub Refresh_Folders()
        Try
            lstDepartment.Items.Clear()
            lstDepartment.Text = ""
            lstCourses.Items.Clear()
            lstCourses.Text = ""
            lstAssignments.Items.Clear()
            lstAssignments.Text = ""
            RootDirectory = filefoldertxt.Text
            InitialRootDirectory = RootDirectory

            If RootDirectory.StartsWith("\\") Then
                RootDirectory = Worker1.MapDrive(RootDirectory)
            End If
            If Not (RootDirectory).IndexOf("Fail") = -1 Then
                Dim Display_Message1 As New Display_Message("An usable Handin directory cannot be located and as such this application will now shut down.")
                Display_Message1.ShowDialog()
                ' application_exit = True
            End If
            If My.Computer.FileSystem.DirectoryExists(RootDirectory) = True Then

                Dim dinfo As DirectoryInfo = New DirectoryInfo(RootDirectory)

                Dim dinfo2 As DirectoryInfo
                If dinfo.Exists = True Then
                    Dim yearfound As Boolean = False

                    For Each dinfo2 In dinfo.GetDirectories
                        If dinfo2.Name = Format(Now(), "yyyy") Then
                            yearfound = True
                            Exit For
                        End If
                    Next
                    If yearfound = False Then
                        dinfo = New DirectoryInfo(RootDirectory)
                        dinfo.CreateSubdirectory(Format(Now(), "yyyy"))
                        'CommunicationLabel1.Text = "Sorry, but no valid project hand-in folders have been created for " & Format(Now(), "yyyy") & ". "
                        'CommunicationLabel2.Text = ""
                        ' lstDepartment.Enabled = False
                        '                statsDepartment.Text = ""
                        '               lstCourses.Enabled = False
                        '              statsCourses.Text = ""
                        '             statsCoursesTotal.Text = ""
                        '            lstAssignments.Enabled = False
                        '           statsAssignments.Text = ""
                        '          statsAssignmentsTotal.Text = ""
                    End If
                    lstDepartment.Enabled = True
                    dinfo = New DirectoryInfo((RootDirectory & "\").Replace("\\", "\") & Format(Now(), "yyyy"))
                    ' lstDepartmentCover.Text = ""
                    lstDepartment.Items.Clear()
                    lstDepartment.Text = ""
                    For Each dinfo2 In dinfo.GetDirectories
                        lstDepartment.Items.Add(New MTGCComboBoxItem(dinfo2.Name))
                    Next
                    'End If
                    If lstDepartment.Items.Count > 0 Then
                        lstDepartment.Enabled = True
                        lstDepartment.SelectedIndex = 0
                    Else
                        '  lstDepartment.Enabled = False
                        statsDepartment.Text = ""
                        ' lstCourses.Enabled = False
                        statsCourses.Text = ""
                        ' statsCoursesTotal.Text = ""
                        'lstAssignments.Enabled = False
                        statsAssignments.Text = ""
                        '  statsAssignmentsTotal.Text = ""
                    End If
                    If lstDepartment.Items.Count = 1 Then
                        statsDepartment.Text = lstDepartment.Items.Count & " Department available to choose from."
                    Else
                        statsDepartment.Text = lstDepartment.Items.Count & " Departments available to choose from."
                    End If

                Else
                    Dim Display_Message1 As New Display_Message("An usable Handin directory cannot be located and as such this application will now shut down.")
                    Display_Message1.ShowDialog()
                End If

                dinfo = Nothing
                dinfo2 = Nothing
            End If
        Catch ex As Exception
            Error_Handler(ex, "Refresh Folders")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Refresh_Folders()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            If Worker1.WorkerThread Is Nothing = False Then
                Worker1.WorkerThread.Abort()
                Worker1.Dispose()
            End If
        Catch ex As Exception
            Error_Handler(ex)
        Finally
            WorkerHandler("Killed")
        End Try
    End Sub


    Private Sub lstDepartment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstDepartment.SelectedIndexChanged
        Try
            ' lstDepartmentCover.Text = lstDepartment.SelectedItem.text
            lblProceed.Text = lstDepartment.Text
            LoadCourseFolders((RootDirectory & "\" & Format(Now, "yyyy") & "\" & lstDepartment.Text).Replace("\\", "\"))
        Catch ex As Exception
            Error_Handler(ex, "lstDepartment_SelectedIndexChanged")
        End Try
    End Sub

    Private Sub lstCourses_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstCourses.SelectedIndexChanged
        Try
            ' lstCoursesCover.Text = lstCourses.SelectedItem.text
            lblProceed.Text = lstDepartment.Text & "\" & lstCourses.Text
            LoadAssignmentFolders((RootDirectory & "\" & Format(Now, "yyyy") & "\" & lstDepartment.Text & "\" & lstCourses.Text).Replace("\\", "\"))
        Catch ex As Exception
            Error_Handler(ex, "lstCourses_SelectedIndexChanged")
        End Try
    End Sub

    Private Sub LoadCourseFolders(ByVal targetdir As String)
        Try
            Dim dinfo2 As DirectoryInfo
            Dim dinfo As DirectoryInfo = New DirectoryInfo(targetdir)
            Dim missedcourses As Integer = 0
            If dinfo.Exists = True Then
                lstCourses.Items.Clear()
                '  lstCoursesCover.Text = ""
                lstCourses.Text = ""
                lstAssignments.Items.Clear()
                ' lstAssignmentsCover.Text = ""
                lstAssignments.Text = ""
                For Each dinfo2 In dinfo.GetDirectories
                    ' If RegisteredCourses.Items.IndexOf(dinfo2.Name) > -1 Then
                    lstCourses.Items.Add(New MTGCComboBoxItem(dinfo2.Name))
                    'Else
                    '   missedcourses = missedcourses + 1
                    'End If


                Next

            End If

            If lstCourses.Items.Count > 0 Then
                lstCourses.Enabled = True
                '  lstCourses.SelectedIndex = 0

                If lstCourses.Items.Count = 1 Then
                    statsCourses.Text = lstCourses.Items.Count & " Course available to choose from under " & lstDepartment.Text & "."
                Else
                    statsCourses.Text = lstCourses.Items.Count & " Courses available to choose from under " & lstDepartment.Text & "."
                End If

            Else
                If lstCourses.Items.Count = 1 Then
                    statsCourses.Text = lstCourses.Items.Count & " Course available to choose from under " & lstDepartment.Text & "."
                Else
                    statsCourses.Text = lstCourses.Items.Count & " Courses available to choose from under " & lstDepartment.Text & "."
                End If
                statsAssignments.Text = ""

            End If
            dinfo2 = Nothing
            dinfo = Nothing
        Catch ex As Exception
            Error_Handler(ex, "LoadCourseFolders")
        End Try
    End Sub


    Private Sub lstAssignments_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstAssignments.SelectedIndexChanged
        Try
            ' lstAssignmentsCover.Text = lstAssignments.SelectedItem.text
            lblProceed.Text = lstDepartment.Text & "\" & lstCourses.Text & "\" & lstAssignments.Text

            'Worker1.department = lstDepartment.Text
            'Worker1.course = lstCourses.Text
            'Worker1.project = lstAssignments.Text

            Dim dinfo1 As FileInfo = New FileInfo((RootDirectory & "\" & Format(Now(), "yyyy") & "\" & lstDepartment.Text & "\" & lstCourses.Text & "\" & lstAssignments.Text & "\access.ini").Replace("\\", "\"))
            If dinfo1.Exists = True Then

                Dim filereader As StreamReader = New StreamReader(dinfo1.FullName)
                Dim lineread As String = ""
                Dim opentime As String = ""
                Dim closetime As String = ""
                While Not filereader.Peek = -1
                    lineread = filereader.ReadLine
                    If lineread.ToLower.StartsWith("open:") Then
                        opentime = lineread.Replace("open:", "")
                    End If
                    If lineread.ToLower.StartsWith("close:") Then
                        closetime = lineread.Replace("close:", "")
                    End If
                End While
                filereader.Close()
                filereader = Nothing
                If opentime.Length = 12 And closetime.Length = 12 Then
                    yearopen.Text = opentime.Substring(0, 4)
                    monthopen.Text = opentime.Substring(4, 2)
                    dayopen.Text = opentime.Substring(6, 2)
                    houropen.Text = opentime.Substring(8, 2)
                    minuteopen.Text = opentime.Substring(10, 2)

                    yearclose.Text = closetime.Substring(0, 4)
                    monthclose.Text = closetime.Substring(4, 2)
                    dayclose.Text = closetime.Substring(6, 2)
                    hourclose.Text = closetime.Substring(8, 2)
                    minuteclose.Text = closetime.Substring(10, 2)
                End If
            End If
            dinfo1 = Nothing
        Catch ex As Exception
            Error_Handler(ex, "lstAssignments_SelectedIndexChanged")
        End Try
    End Sub

    Private Sub LoadAssignmentFolders(ByVal targetdir As String)
        Try
            Dim dinfo2 As DirectoryInfo
            Dim dinfo As DirectoryInfo = New DirectoryInfo(targetdir)

            If dinfo.Exists = True Then
                lstAssignments.Items.Clear()
                '  lstAssignmentsCover.Text = ""
                lstAssignments.Text = ""
                For Each dinfo2 In dinfo.GetDirectories

                    lstAssignments.Items.Add(New MTGCComboBoxItem(dinfo2.Name))

                Next

            End If

            If lstAssignments.Items.Count > 0 Then
                lstAssignments.Enabled = True
                'lstAssignments.SelectedIndex = 0

                If lstAssignments.Items.Count = 1 Then
                    statsAssignments.Text = lstAssignments.Items.Count & " Assignment available to choose from under " & lstCourses.Text & "."
                Else
                    statsAssignments.Text = lstAssignments.Items.Count & " Assignments available to choose from under " & lstCourses.Text & "."
                End If

            Else

                If lstAssignments.Items.Count = 1 Then
                    statsAssignments.Text = lstAssignments.Items.Count & " Assignment available to choose from under " & lstCourses.Text & "."
                Else
                    statsAssignments.Text = lstAssignments.Items.Count & " Assignments available to choose from under " & lstCourses.Text & "."
                End If
                ' lstAssignments.Enabled = False

                '                statsAssignmentsTotal.Text = ""
            End If
            dinfo2 = Nothing
            dinfo = Nothing
        Catch ex As Exception
            Error_Handler(ex, "LoadAssignmentFolders")
        End Try
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Dim result As DialogResult = FolderBrowserDialog1.ShowDialog
            If result = Windows.Forms.DialogResult.OK Then
                filefoldertxt.Text = FolderBrowserDialog1.SelectedPath
                Refresh_Folders()
            End If

        Catch ex As Exception
            Error_Handler(ex)
        End Try

    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            If lstDepartment.Text.Length > 0 And lstCourses.Text.Length > 0 And lstAssignments.Text.Length > 0 Then
                NotifyIcon1.Text = "Generating Folders..."
                If currently_working = False Then
                    run_worker(True)
                End If
            Else
                MsgBox("You need to enter values in the Department, Course and Project fields in order to generate a valid project folder", MsgBoxStyle.Information, "Invalid Inputs Detected")
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        MsgBox("The batch creation function allows you to automate the creation of a number of handin folders all at once. Please ensure the input text file you choose to load conforms to the following standard: #department#;#course#;#project#;#open_yyyymmddHHMM#;#close_yyyymmddHHMM#", MsgBoxStyle.Information, "Start Batch Process")
        Try
            If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                joblist = 0
                Dim reader As StreamReader = New StreamReader(OpenFileDialog1.FileName, False)
                Dim inputline As String = ""
                Dim createfolders As Boolean = False
                'continuenextbatch = True
                While reader.Peek <> -1

                    inputline = reader.ReadLine
                    joblist = joblist + 1
                End While
                reader.Close()
                reader = Nothing
                'continuenextbatch = False
                last_thread_run = 1
                force_check()
            End If
        Catch ex As Exception
            Error_Handler(ex, "Batch Process")
        End Try

    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Try
            Label2.Text = "About displayed"
            AboutBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display About Screen")
        End Try
    End Sub

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Try
            Label2.Text = "Help displayed"
            HelpBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display Help Screen")
        End Try
    End Sub
End Class
