VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form SetupConverter 
   Caption         =   "Setup Inno"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar progBar 
      Height          =   735
      Left            =   10680
      TabIndex        =   9
      Top             =   9000
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
   End
   Begin TabDlg.SSTab tabSetup 
      Height          =   10695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   18865
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Setup"
      TabPicture(0)   =   "SetupConverter.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Files"
      TabPicture(1)   =   "SetupConverter.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lstFiles"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "AddFile"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "AddPath"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Last"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "SaveList"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Inno Setup File"
      TabPicture(2)   =   "SetupConverter.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtInno"
      Tab(2).ControlCount=   1
      Begin VB.Frame fra 
         Height          =   10095
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   14775
         Begin VB.CommandButton Icon 
            Caption         =   "Select"
            Height          =   375
            Left            =   9720
            TabIndex        =   61
            Top             =   8160
            Width           =   855
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   16
            Left            =   1920
            TabIndex        =   59
            Top             =   8160
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   41
            Tag             =   "AppPublisher"
            Text            =   "txtSetup"
            Top             =   1320
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   40
            Tag             =   "AppName"
            Text            =   "txtSetup"
            Top             =   840
            Width           =   7815
         End
         Begin VB.CheckBox chkShowGroup 
            Caption         =   "Always Show Group On Ready Page"
            Height          =   375
            Left            =   10920
            TabIndex        =   39
            Tag             =   "AlwaysShowGroupOnReadyPage"
            Top             =   2160
            Width           =   3375
         End
         Begin VB.CheckBox chkShowDir 
            Caption         =   "Always Show Directory Ready Page"
            Height          =   375
            Left            =   10920
            TabIndex        =   38
            Tag             =   "AlwaysShowDirOnReadyPage"
            Top             =   1680
            Width           =   3375
         End
         Begin VB.CheckBox chkUNC 
            Caption         =   "Allow UNC Path"
            Height          =   375
            Left            =   10920
            TabIndex        =   37
            Tag             =   "AllowUNCPath"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CheckBox chkAllowRoot 
            Caption         =   "Allow Root Directory"
            Height          =   375
            Left            =   10920
            TabIndex        =   36
            Tag             =   "AllowRootDirectory"
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chkAdmin 
            Caption         =   "Administrator Privileges Required"
            Height          =   375
            Left            =   10920
            TabIndex        =   35
            Tag             =   "AdminPrivilegesRequired"
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   34
            Tag             =   "OutputBaseFileName"
            Text            =   "txtSetup"
            Top             =   360
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   33
            Tag             =   "AppPublisherURL"
            Text            =   "txtSetup"
            Top             =   1800
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   4
            Left            =   1920
            TabIndex        =   32
            Tag             =   "AppVersion"
            Text            =   "txtSetup"
            Top             =   2280
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   5
            Left            =   1920
            TabIndex        =   31
            Tag             =   "AppVerName"
            Text            =   "txtSetup"
            Top             =   2760
            Width           =   7815
         End
         Begin VB.CheckBox chkCreateDir 
            Caption         =   "Create Application Directory"
            Height          =   375
            Left            =   10920
            TabIndex        =   30
            Tag             =   "CreateAppDir"
            Top             =   3120
            Width           =   2655
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   6
            Left            =   1920
            TabIndex        =   29
            Tag             =   "DefaultDirName"
            Text            =   "txtSetup"
            Top             =   3240
            Width           =   7815
         End
         Begin VB.CheckBox chkWarning 
            Caption         =   "Display Directory Existence Warning"
            Height          =   375
            Left            =   10920
            TabIndex        =   28
            Tag             =   "DirExistsWarning"
            Top             =   2640
            Width           =   3375
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   7
            Left            =   1920
            TabIndex        =   27
            Tag             =   "InfoBeforeFile"
            Text            =   "txtSetup"
            Top             =   3720
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   8
            Left            =   1920
            TabIndex        =   26
            Tag             =   "InfoAfterFile"
            Text            =   "txtSetup"
            Top             =   4200
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   9
            Left            =   1920
            TabIndex        =   25
            Tag             =   "LicenseFile"
            Text            =   "txtSetup"
            Top             =   4680
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   10
            Left            =   1920
            TabIndex        =   24
            Tag             =   "Password"
            Text            =   "txtSetup"
            Top             =   5160
            Width           =   7815
         End
         Begin VB.CheckBox chkRestart 
            Caption         =   "Restart If Needed"
            Height          =   375
            Left            =   10920
            TabIndex        =   23
            Tag             =   "RestartIfNeededByRun"
            Top             =   4560
            Width           =   2055
         End
         Begin VB.CheckBox chkUninstall 
            Caption         =   "The Application is Uninstallable"
            Height          =   375
            Left            =   10920
            TabIndex        =   22
            Tag             =   "Uninstallable"
            Top             =   3600
            Width           =   3015
         End
         Begin VB.CheckBox chkUserInfor 
            Caption         =   "Show User Information Page"
            Height          =   375
            Left            =   10920
            TabIndex        =   21
            Tag             =   "UserInfoPage"
            Top             =   4080
            Width           =   2775
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   11
            Left            =   1920
            TabIndex        =   20
            Tag             =   "AppCopyright"
            Text            =   "txtSetup"
            Top             =   5640
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   12
            Left            =   1920
            TabIndex        =   19
            Tag             =   "WizardImageFile"
            Text            =   "txtSetup"
            Top             =   6120
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   13
            Left            =   1920
            TabIndex        =   18
            Tag             =   "WizardSmallImageFile"
            Text            =   "txtSetup"
            Top             =   6600
            Width           =   7815
         End
         Begin VB.CommandButton InforBef 
            Caption         =   "Select"
            Height          =   375
            Left            =   9720
            TabIndex        =   17
            Top             =   3720
            Width           =   855
         End
         Begin VB.CommandButton InforAfter 
            Caption         =   "Select"
            Height          =   375
            Left            =   9720
            TabIndex        =   16
            Top             =   4200
            Width           =   855
         End
         Begin VB.CommandButton License 
            Caption         =   "Select"
            Height          =   375
            Left            =   9720
            TabIndex        =   15
            Top             =   4680
            Width           =   855
         End
         Begin VB.CommandButton ImageFile 
            Caption         =   "Select"
            Height          =   375
            Left            =   9720
            TabIndex        =   14
            Top             =   6120
            Width           =   855
         End
         Begin VB.CommandButton SmallImage 
            Caption         =   "Select"
            Height          =   375
            Left            =   9720
            TabIndex        =   13
            Top             =   6600
            Width           =   855
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   14
            Left            =   1920
            TabIndex        =   12
            Tag             =   "DefaultGroupName"
            Text            =   "txtSetup"
            Top             =   7200
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   15
            Left            =   1920
            TabIndex        =   11
            Text            =   "txtSetup"
            Top             =   7680
            Width           =   7815
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Icon File"
            Height          =   225
            Index           =   17
            Left            =   120
            TabIndex        =   60
            Top             =   8160
            Width           =   690
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Publisher"
            Height          =   225
            Index           =   8
            Left            =   120
            TabIndex        =   58
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   225
            Index           =   7
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Width           =   510
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Output Setup File"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   1410
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Application Publisher"
            Height          =   225
            Index           =   9
            Left            =   7080
            TabIndex        =   55
            Top             =   3840
            Width           =   1740
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Publisher URL"
            Height          =   225
            Index           =   10
            Left            =   120
            TabIndex        =   54
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Version"
            Height          =   225
            Index           =   11
            Left            =   120
            TabIndex        =   53
            Top             =   2280
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Version Name"
            Height          =   225
            Index           =   13
            Left            =   120
            TabIndex        =   52
            Top             =   2760
            Width           =   1185
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Directory Name"
            Height          =   225
            Index           =   14
            Left            =   120
            TabIndex        =   51
            Top             =   3240
            Width           =   1275
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Information Before "
            Height          =   225
            Index           =   15
            Left            =   120
            TabIndex        =   50
            Top             =   3720
            Width           =   1560
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Information After"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   4200
            Width           =   1335
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "License File"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   48
            Top             =   4680
            Width           =   1005
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   225
            Index           =   3
            Left            =   120
            TabIndex        =   47
            Top             =   5160
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Copyrights"
            Height          =   225
            Index           =   4
            Left            =   120
            TabIndex        =   46
            Top             =   5640
            Width           =   885
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wizard Image"
            Height          =   225
            Index           =   5
            Left            =   120
            TabIndex        =   45
            Top             =   6120
            Width           =   1125
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wizard Small Image"
            Height          =   225
            Index           =   6
            Left            =   120
            TabIndex        =   44
            Top             =   6600
            Width           =   1650
         End
         Begin VB.Line Line1 
            X1              =   10680
            X2              =   10680
            Y1              =   120
            Y2              =   6960
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Group Name"
            Height          =   225
            Index           =   12
            Left            =   120
            TabIndex        =   43
            Top             =   7200
            Width           =   1065
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Executable"
            Height          =   225
            Index           =   16
            Left            =   120
            TabIndex        =   42
            Top             =   7680
            Width           =   900
         End
      End
      Begin VB.CommandButton SaveList 
         Caption         =   "Save File List"
         Height          =   375
         Left            =   13560
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Last 
         Caption         =   "Load List"
         Height          =   375
         Left            =   13560
         TabIndex        =   7
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton AddPath 
         Caption         =   "Add Path"
         Height          =   375
         Left            =   13560
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add File"
         Height          =   375
         Left            =   13560
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton AddFile 
         Caption         =   "Add File"
         Height          =   375
         Left            =   13560
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin RichTextLib.RichTextBox txtInno 
         Height          =   10095
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   17806
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"SetupConverter.frx":0054
      End
      Begin MSComctlLib.ListView lstFiles 
         Height          =   10095
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   17806
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Source"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Destination"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Operation"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Shared"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.ComboBox cboFiles 
      Height          =   345
      Left            =   12600
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   6600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   12720
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuVbSetup 
         Caption         =   "Open Vb Setup File"
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "Convert"
      End
      Begin VB.Menu mnuInnoSetup 
         Caption         =   "Inno Setup"
      End
      Begin VB.Menu dr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjects 
         Caption         =   "<Prev Projects>"
         Index           =   0
      End
      Begin VB.Menu de 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "SetupConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private StrSource As String
Private SetupFile As String
Private colFiles As Collection
Private lastPath As String
Private StrIssFile As String
Private lstRun As Collection

Private Sub AddPath_Click()
    Dim sPath As String
    
    
RestartAgain:
    sPath = StringBrowseForFolder(Me.hWnd, "Select Path To Add")
    If Len(sPath) = 0 Then Exit Sub
    
    SaveReg "lastpath", sPath
    lastPath = sPath
    
    Dim pFiles As New Collection
    Dim dPath As String
    Dim dest As String
    Dim tFiles As Long
    Dim cFiles As Long
    Dim xFile As String
    Dim rFile(1 To 4) As String
    Dim nDest As String
    Dim nDestPos As Long
    
    Screen.MousePointer = vbHourglass
    CollectionOfFolders sPath, pFiles
    Screen.MousePointer = vbDefault
    
    dPath = StringGetFileToken(pFiles.Item(1), "p")
    dPath = MvField(dPath, -1, "\")
    
    dest = InputBox("Please enter the starting point of the destination path in each file:", "Destination Path Start", dPath)
    Resp = MsgBox("You have selected the destination path as '" & dest & "'. Is this correct?", vbYesNo + vbQuestion + vbApplicationModal, "Destination Path - " & dest)
    If Resp = vbNo Then GoTo RestartAgain
    
    dest = LCase$(dest)
    Screen.MousePointer = vbHourglass
    tFiles = pFiles.Count
    For cFiles = 1 To tFiles
        xFile = LCase$(pFiles.Item(cFiles))
        If boolFileExists(xFile) = False Then GoTo NextFile
        
        nDestPos = InStr(1, xFile, "\" & dest)
        If nDestPos > 0 Then
            nDest = Mid$(xFile, nDestPos)
            nDest = "{app}" & StringGetFileToken(nDest, "p")
            
            rFile(1) = StringProperCase(xFile)
            rFile(2) = StringProperCase(nDest)
            rFile(3) = ""
            Call LstViewUpdate(rFile, lstFiles, "")
            
            Select Case InStr(1, xFile, ".exe")
            Case 0
            Case Else
                If InStr(1, LCase$(xFile), LCase$(txtSetup(15).Text)) = 0 Then
                    lstRun.Add nDest & "\" & StringGetFileToken(xFile, "f")
                End If
            End Select
        End If
NextFile:
    Next
    CodeLibraryNew.LstViewRemoveDuplicates lstFiles
    LstViewAutoResize lstFiles
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    AppTitle = App.Title
    
    CodeLibraryNew.CleanAllControls Me
    tabSetup.Tab = 0
        
    LstBoxFromMV cboFiles, ReadReg("lastfiles")
    CodeLibraryNew.ReloadMenu mnuProjects, cboFiles
End Sub

Private Sub Icon_Click()
    txtSetup(16).Text = DialogOpenAPI(Me, StringFileFilters, "Select Icon File", lastPath, "*.ico")
    If Len(txtSetup(16).Text) > 0 Then
        Dim spRec(1 To 4) As String
        spRec(1) = txtSetup(16).Text
        spRec(2) = "{app}"
        LstViewUpdate spRec, lstFiles, ""
        CodeLibraryNew.LstViewRemoveDuplicates lstFiles
        LstViewAutoResize lstFiles
    End If
    
End Sub

Private Sub ImageFile_Click()
        txtSetup(12).Text = DialogOpenAPI(Me, StringFileFilters, "Select Image File", lastPath, "*.bmp")

End Sub

Private Sub InforAfter_Click()
        txtSetup(8).Text = DialogOpenAPI(cDialog, StringFileFilters, "Select Information After File", lastPath, "*.txt")

End Sub

Private Sub InforBef_Click()
    txtSetup(7).Text = DialogOpenAPI(cDialog, StringFileFilters, "Select Information Before File", lastPath, "*.txt")
End Sub

Private Sub Last_Click()
    If boolFileExists(App.Path & "\FileList.txt") = True Then
        Screen.MousePointer = vbHourglass
        LstViewFromFile lstFiles, App.Path & "\FileList.txt"
        LstViewAutoResize lstFiles
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub License_Click()
    txtSetup(9).Text = DialogOpenAPI(cDialog, StringFileFilters, "Select License File", lastPath, "*.txt")

End Sub

Private Sub mnuInnoSetup_Click()
    Call boolViewFile(StrIssFile)
End Sub

Private Sub mnuProjects_Click(Index As Integer)
    SetupFile = mnuProjects(Index).Caption
    ReadVbSetupFile
    Set lstRun = New Collection
End Sub

Private Sub mnuVbSetup_Click()
    ReadVbSetupFile True
    Set lstRun = New Collection
End Sub


Private Sub mnuConvert_Click()
    tabSetup.Tab = 2
    txtInno.Text = ""
    
    
    StrIssFile = App.Path & "\" & txtSetup(1).Text
    If boolDirExists(StrIssFile) = False Then MkDir StrIssFile
    
    StrIssFile = StrIssFile & "\" & txtSetup(1).Text & ".iss"
    Screen.MousePointer = vbHourglass
    
    Dim lstDetails As Collection
    
    Dim lstRow() As String
    Dim lstLine As String
    Dim strFlags As String
    
    Dim frmTot As Long
    Dim frmCnt As Long
    Dim frmTag As String
    Dim frmType As String
    Dim frmLine As String
    
    Set lstDetails = New Collection
    
    lstDetails.Add "; Script generated by the Inno Setup Script Wizard."
    lstDetails.Add "; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!"
    lstDetails.Add ""
    lstDetails.Add "[Setup]"
    
    frmTot = Me.Controls.Count - 1
    For frmCnt = 0 To frmTot
        frmType = TypeName(Me.Controls(frmCnt))
        frmTag = Me.Controls(frmCnt).Tag
        If frmTag = "" Then GoTo NextEntry
        
        Select Case frmType
        Case "TextBox"
            If Me.Controls(frmCnt).Text <> "" Then
                frmLine = frmTag & "=" & Me.Controls(frmCnt).Text
                lstDetails.Add frmLine
            End If
        Case "CheckBox"
            If Me.Controls(frmCnt).Value = 1 Then
                frmLine = frmTag & "=yes"
            Else
                frmLine = frmTag & "=no"
            End If
            lstDetails.Add frmLine
        End Select
        
NextEntry:
    Next
        
    lstDetails.Add ""
    lstDetails.Add "[Tasks]"
    lstDetails.Add "Name: " & Quote & "desktopicon" & Quote & "; Description: " & Quote & "Create a &desktop icon" & Quote & "; GroupDescription: " & Quote & "Additional icons:" & Quote
    lstDetails.Add ""
    lstDetails.Add "[Files]"

    frmTot = lstFiles.ListItems.Count
    For frmCnt = 1 To frmTot
        LstViewGetRow lstRow, lstFiles, frmCnt
        strFlags = LCase$(Trim$(lstRow(3) & " " & lstRow(4)))
        lstRow(1) = UCase$(lstRow(1))
        
        Select Case UCase$(lstRow(3))
        Case "REGSERVER", "REGTYPELIB"
            strFlags = strFlags & " noregerror"
        End Select
        
        If InStr(1, lstRow(1), "STKIT") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "COMCAT") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "ASYCFILT") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "OLEPRO") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "OLEAUT") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "STDOLE") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "MSVBVM") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "MSVCRT") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        
        If lstRow(4) = "SHAREDFILE" Then
            If InStr(1, strFlags, "uninsneveruninstall") = 0 Then strFlags = "uninsneveruninstall " & strFlags
        End If
        
        lstLine = "Source: " & Quote & StringProperCase(lstRow(1)) & Quote & "; DestDir: " & Quote & lstRow(2) & Quote & "; Flags: " & strFlags
        lstDetails.Add lstLine
        
        
    Next
    lstDetails.Add ""
    lstDetails.Add ";NOTE: Don't use " & Quote & "Flags: ignoreversion" & Quote & " on any shared system files"
    lstDetails.Add ""
    lstDetails.Add "[Icons]"
    lstDetails.Add "Name: " & Quote & "{group}\" & txtSetup(14).Text & Quote & "; Filename: " & Quote & "{app}\" & txtSetup(15).Text & Quote
    lstDetails.Add "Name: " & Quote & "{group}\Uninstall " & txtSetup(1).Text & Quote & "; Filename: " & Quote & "{uninstallexe}" & Quote
    If txtSetup(16).Text = "" Then
        lstDetails.Add "Name: " & Quote & "{userdesktop}\" & txtSetup(1).Text & Quote & "; Filename: " & Quote & "{app}\" & txtSetup(15).Text & Quote & "; Tasks: desktopicon"
    Else
        lstDetails.Add "Name: " & Quote & "{userdesktop}\" & txtSetup(1).Text & Quote & "; Filename: " & Quote & "{app}\" & txtSetup(15).Text & Quote & "; IconFilename: " & Quote & "{app}\" & StringGetFileToken(txtSetup(16).Text, "f") & Quote & "; Tasks: desktopicon"
    End If
    
    lstDetails.Add ""
    lstDetails.Add "[Run]"
    lstDetails.Add "Filename: " & Quote & "{app}\" & txtSetup(15).Text & Quote & "; Description: " & Quote & "Launch " & txtSetup(1).Text & Quote & "; Flags: nowait postinstall skipifsilent"
    For frmCnt = 1 To lstRun.Count
        lstDetails.Add "Filename: " & Quote & lstRun.Item(frmCnt) & Quote & ";Parameters: " & Quote & "/q" & Quote
    Next
    
    txtInno.Text = MvFromCollection(lstDetails, NL)
    txtInno.SaveFile StrIssFile, rtfText
    Screen.MousePointer = vbDefault
    Beep
End Sub

Private Sub Form_Activate()
    AppTitle = App.Title
End Sub


Function StrAbbreviate(ByVal StrData As String) As String
    Dim lngTot As Long
    Dim lngCnt As Long
    Dim spData() As String
    Dim strNew As String
    
    strNew = ""
    
    StringParse spData, StrData, " "
    lngTot = UBound(spData)
    For lngCnt = 1 To lngTot
        strNew = StringsConcat(strNew, Left$(spData(lngCnt), 1))
    Next
    StrAbbreviate = strNew
End Function


Private Sub SaveList_Click()
    Screen.MousePointer = vbHourglass
    LstViewToFile progBar, lstFiles, App.Path & "\FileList.txt"
    Screen.MousePointer = vbDefault
End Sub

Private Sub SmallImage_Click()
        txtSetup(13).Text = DialogOpenAPI(cDialog, StringFileFilters, "Select Small Image File", lastPath, "*.bmp")

End Sub

Private Sub txtSetup_Validate(Index As Integer, Cancel As Boolean)
    txtSetup(Index).Text = StringProperCase(txtSetup(Index).Text)
    
    Select Case Index
    Case 2
        txtSetup(11).Text = txtSetup(2).Text
    End Select
    
End Sub

Private Sub ReadVbSetupFile(Optional Prompt As Boolean = False)
    Dim txtCnt As Integer
       
    lastPath = ReadReg("lastpath")
    If lastPath = "" Then lastPath = App.Path
    
    If Prompt = True Then
        SetupFile = DialogOpenAPI(Me, StringFileFilters, "Select VB Setup File", lastPath, "*.lst")
        If Len(SetupFile) = 0 Then Exit Sub
    End If
    
    Caption = "Setup Inno - " & SetupFile
    lastPath = StringGetFileToken(SetupFile, "p")
    SaveReg "lastpath", lastPath
    
    For txtCnt = 0 To txtSetup.Count - 1
        txtSetup(txtCnt).Text = ""
    Next
    lstFiles.ListItems.Clear
    txtInno.Text = ""
    
    StrSource = StringsConcat(lastPath, "\Support")
    
    If boolDirExists(StrSource) = False Then
        MsgBox "The support directory for the setup files does not exists." & vbCr & _
        "This directory is usually created by the VB setup program.", vbOKOnly + vbExclamation + vbApplicationModal, StrSource & " Error"
        Exit Sub
    End If
            
    LstBoxUpdate cboFiles, SetupFile
    SaveReg "lastfiles", LstBoxToMV(cboFiles, VM)
    CodeLibraryNew.ReloadMenu mnuProjects, cboFiles

    Screen.MousePointer = vbHourglass
    Set colFiles = New Collection
    Dim intFile As String
    Dim StrLine As String
    Dim equalPos As Long
    Dim sHead As String
    Dim sRest As String
    Dim sFile As String
    Dim sOper As String
    Dim sDest As String
    Dim sShared As String
    Dim arrLine() As String
    Dim lstLine(1 To 4) As String
    
    intFile = FreeFile
    Open SetupFile For Input Access Read As #intFile
    Do Until EOF(intFile)
        Line Input #intFile, StrLine
        StrLine = LCase$(Trim$(StrLine))
        If Len(StrLine) = 0 Then GoTo NextLine
            
        equalPos = InStr(1, StrLine, "=")
        Select Case equalPos
        Case 0
            GoTo NextLine
        Case Else
            sHead = Left$(MvField(StrLine, 1, "="), 4)
            sRest = MvField(StrLine, 2, "=")
            
            Select Case sHead
            Case "file"
                StringParse arrLine, sRest, ","
                ReDim Preserve arrLine(4)
                
                sFile = StringProperCase(Mid$(arrLine(1), 2))
                sDest = arrLine(2)
                sOper = Replace(arrLine(3), "$(dllselfregister)", "Regserver")
                sOper = Replace(sOper, "$(tlbregister)", "Regtypelib")
                sShared = Replace(arrLine(4), "$(shared)", "Sharedfile")
                
                sDest = Replace(sDest, "$(winsyspathsysfile)", "{sys}")
                sDest = Replace(sDest, "$(apppath)", "{app}")
                sDest = Replace(sDest, "$(winsyspath)", "{sys}")
                sDest = Replace(sDest, "$(winpath)", "{win}")
                sDest = Replace(sDest, "$(msdaopath)", "{dao}")
                
                lstLine(1) = StringsConcat(StrSource, "\", sFile)
                lstLine(2) = sDest
                lstLine(3) = sOper
                lstLine(4) = sShared
                
                LstViewUpdate lstLine, lstFiles, ""
            Case "titl"
                txtSetup(1).Text = StringProperCase(sRest)
                txtSetup(6).Text = "{sd}\" & StrAbbreviate(sRest)
                txtSetup(5).Text = StringProperCase(sRest)
            Case "appe"
                txtSetup(0).Text = StringGetFileToken(StringProperCase(sRest), "fo")
                txtSetup(15).Text = txtSetup(0).Text
            Case "grou"
                txtSetup(14).Text = StringProperCase(sRest)
            End Select
        End Select
NextLine:
    Loop
    Close #intFile
    LstViewAutoResize lstFiles
    Screen.MousePointer = vbDefault
End Sub


Public Sub LstViewToFile(progBar As Object, lstView As Object, ByVal strFile As String, Optional Delim As String = "")
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim lstTot As Long
    Dim lstCnt As Long
    Dim lstStr() As String
    Dim retStr As String
    Dim intFile As Integer
    Dim colStr As String
    
    intFile = FreeFile
    Open strFile For Output Access Write As #intFile
    retStr = ""
    If Len(Delim) = 0 Then Delim = VM
    lstTot = lstView.ListItems.Count
    colStr = LstViewColNames(lstView)
    colStr = Replace(colStr, ",", Delim)
    Print #intFile, colStr
    For lstCnt = 1 To lstTot
        LstViewGetRow lstStr, lstView, lstCnt
        retStr = MvFromArray(lstStr, Delim)
        Print #intFile, retStr
    Next
    Close #intFile
End Sub


