VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Permutations and Combinations"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTxtRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   6840
      TabIndex        =   107
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5640
      TabIndex        =   27
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtStatus 
      Height          =   4815
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   495
      Left            =   4080
      TabIndex        =   16
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00000000&
      Caption         =   "Stop"
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTabOptions 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Output"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmOutput"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmOMode"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmFilename"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Characters"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmCharacters"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Length"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmlength"
      Tab(2).Control(1)=   "frmlength2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "P && C"
      TabPicture(3)   =   "frmMain.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frmMode"
      Tab(3).Control(1)=   "frmCalculations"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "BruteForce"
      TabPicture(4)   =   "frmMain.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "SSTabBF"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "About"
      TabPicture(5)   =   "frmMain.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "SSTabAbout"
      Tab(5).ControlCount=   1
      Begin VB.Frame frmFilename 
         Caption         =   "Filename "
         Height          =   855
         Left            =   360
         TabIndex        =   67
         Top             =   2520
         Width           =   4335
         Begin VB.TextBox txtFilename 
            Height          =   285
            Left            =   1680
            TabIndex        =   68
            Text            =   "C:\permute.txt"
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label lblPromptOutFile 
            Caption         =   "Please input filename"
            Height          =   375
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame frmlength2 
         Caption         =   "Length"
         Height          =   855
         Left            =   -74640
         TabIndex        =   64
         Top             =   3000
         Width           =   4335
         Begin VB.TextBox txtLength 
            Height          =   285
            Left            =   2280
            MaxLength       =   1
            TabIndex        =   65
            Text            =   "3"
            Top             =   330
            Width           =   255
         End
         Begin VB.Label lblPromptLength 
            Caption         =   "Length of Combinations"
            Height          =   375
            Left            =   240
            TabIndex        =   66
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame frmCalculations 
         Caption         =   "Calculations"
         Height          =   2415
         Left            =   -74640
         TabIndex        =   55
         Top             =   2040
         Width           =   4335
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   120
            TabIndex        =   106
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox txtTotalWords 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   72
            Text            =   "No time to complete"
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox txtTime 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   63
            Text            =   "No time to complete"
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtTotalSize 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   61
            Text            =   "No time to complete"
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtNoPerm 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   60
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtNoComb 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   59
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label19 
            Caption         =   "Total No of Words"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "Estimated Time"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "Size of Output File"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Total No of Combinations"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label12 
            Caption         =   "Total No of Permutations"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   600
            Width           =   1935
         End
      End
      Begin TabDlg.SSTab SSTabBF 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   7011
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Setup"
         TabPicture(0)   =   "frmMain.frx":00A8
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtBFSetup"
         Tab(0).Control(1)=   "frmKeys"
         Tab(0).Control(2)=   "cmdReset"
         Tab(0).Control(3)=   "cmdRepeat"
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Help"
         TabPicture(1)   =   "frmMain.frx":00C4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label1"
         Tab(1).Control(1)=   "Label2"
         Tab(1).Control(2)=   "Label3"
         Tab(1).Control(3)=   "Label4"
         Tab(1).Control(4)=   "Label5"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Key Codes"
         TabPicture(2)   =   "frmMain.frx":00E0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtHelp"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Bugs"
         TabPicture(3)   =   "frmMain.frx":00FC
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "Label20"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "Label21"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "frmParse"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).ControlCount=   3
         Begin VB.Frame frmParse 
            Caption         =   "Select Parse mode"
            Height          =   735
            Left            =   120
            TabIndex        =   77
            Top             =   480
            Width           =   4215
            Begin VB.OptionButton optPmode2 
               Caption         =   "Mode 2"
               Height          =   255
               Left            =   1560
               TabIndex        =   79
               ToolTipText     =   "This Process is slower but more reliable"
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton optPmode1 
               Caption         =   "Mode 1"
               Height          =   255
               Left            =   120
               TabIndex        =   78
               ToolTipText     =   "This process is faster but less reliable"
               Top             =   360
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.CommandButton cmdRepeat 
            Caption         =   "Repeat"
            Height          =   375
            Left            =   -73920
            TabIndex        =   76
            Top             =   3360
            Width           =   855
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "Reset !"
            Height          =   375
            Left            =   -74760
            TabIndex        =   70
            Top             =   3360
            Width           =   735
         End
         Begin VB.Frame frmKeys 
            Caption         =   "Commonly used keys"
            Height          =   3495
            Left            =   -72840
            TabIndex        =   36
            Top             =   360
            Width           =   2295
            Begin VB.CommandButton cmdPercent 
               Caption         =   "%"
               Height          =   375
               Left            =   1560
               TabIndex        =   75
               Top             =   3000
               Width           =   615
            End
            Begin VB.CommandButton cmdPlus 
               Caption         =   "+"
               Height          =   375
               Left            =   840
               TabIndex        =   74
               Top             =   3000
               Width           =   615
            End
            Begin VB.CommandButton cmdCaret 
               Caption         =   "^"
               Height          =   375
               Left            =   120
               TabIndex        =   73
               Top             =   3000
               Width           =   615
            End
            Begin VB.CommandButton cmdBkSp 
               Caption         =   "BkSp"
               Height          =   375
               Left            =   1560
               TabIndex        =   54
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton cmdCAPS 
               Caption         =   "CAPS"
               Height          =   375
               Left            =   840
               TabIndex        =   53
               Top             =   1680
               Width           =   615
            End
            Begin VB.CommandButton cmdIns 
               Caption         =   "Ins"
               Height          =   375
               Left            =   1560
               TabIndex        =   52
               Top             =   720
               Width           =   615
            End
            Begin VB.CommandButton cmdDel 
               Caption         =   "Del"
               Height          =   375
               Left            =   840
               TabIndex        =   51
               Top             =   720
               Width           =   615
            End
            Begin VB.CommandButton cmdPgDn 
               Caption         =   "PgDn"
               Height          =   375
               Left            =   1560
               TabIndex        =   50
               Top             =   2160
               Width           =   615
            End
            Begin VB.CommandButton cmdDown 
               Caption         =   "Down"
               Height          =   375
               Left            =   840
               TabIndex        =   49
               Top             =   2160
               Width           =   615
            End
            Begin VB.CommandButton cmdEnd 
               Caption         =   "End"
               Height          =   375
               Left            =   120
               TabIndex        =   48
               Top             =   2160
               Width           =   615
            End
            Begin VB.CommandButton cmdRight 
               Caption         =   "Right"
               Height          =   375
               Left            =   1560
               TabIndex        =   47
               Top             =   1680
               Width           =   615
            End
            Begin VB.CommandButton cmdLeft 
               Caption         =   "Left"
               Height          =   375
               Left            =   120
               TabIndex        =   46
               Top             =   1680
               Width           =   615
            End
            Begin VB.CommandButton cmdPgUp 
               Caption         =   "PgUp"
               Height          =   375
               Left            =   1560
               TabIndex        =   45
               Top             =   1200
               Width           =   615
            End
            Begin VB.CommandButton cmdHome 
               Caption         =   "Home"
               Height          =   375
               Left            =   120
               TabIndex        =   44
               Top             =   1200
               Width           =   615
            End
            Begin VB.CommandButton cmdUp 
               Caption         =   "Up"
               Height          =   375
               Left            =   840
               TabIndex        =   43
               Top             =   1200
               Width           =   615
            End
            Begin VB.CommandButton cmdEscape 
               Caption         =   "Esc"
               Height          =   375
               Left            =   840
               TabIndex        =   42
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox chkAlt 
               Caption         =   "Alt"
               Height          =   255
               Left            =   1560
               TabIndex        =   41
               Top             =   2640
               Width           =   615
            End
            Begin VB.CheckBox chkCtrl 
               Caption         =   "Ctrl"
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   2640
               Width           =   615
            End
            Begin VB.CheckBox chkShift 
               Caption         =   "Shift"
               Height          =   255
               Left            =   840
               TabIndex        =   39
               Top             =   2640
               Width           =   735
            End
            Begin VB.CommandButton cmdEnter 
               Caption         =   "Enter"
               Height          =   375
               Left            =   120
               TabIndex        =   38
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton cmdTab 
               Caption         =   "Tab"
               Height          =   375
               Left            =   120
               TabIndex        =   37
               Top             =   720
               Width           =   615
            End
         End
         Begin VB.TextBox txtHelp 
            Height          =   3255
            Left            =   -74640
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Text            =   "frmMain.frx":0118
            Top             =   480
            Width           =   4095
         End
         Begin VB.TextBox txtBFSetup 
            Height          =   2655
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label21 
            Caption         =   $"frmMain.frx":0420
            Height          =   615
            Left            =   120
            TabIndex        =   81
            Top             =   2760
            Width           =   4215
         End
         Begin VB.Label Label20 
            Caption         =   $"frmMain.frx":04BC
            Height          =   1455
            Left            =   120
            TabIndex        =   80
            Top             =   1320
            Width           =   4215
         End
         Begin VB.Label Label5 
            Caption         =   "Anyway atleast you can set it up so that some keys are automatically sent to window where cut copy cannot be used."
            Height          =   735
            Left            =   -74880
            TabIndex        =   33
            Top             =   2880
            Width           =   4335
         End
         Begin VB.Label Label4 
            Caption         =   "So those combinations should be entered below. All the best. Although chances are very low."
            Height          =   495
            Left            =   -74880
            TabIndex        =   32
            Top             =   2400
            Width           =   4335
         End
         Begin VB.Label Label3 
            Caption         =   $"frmMain.frx":064D
            Height          =   975
            Left            =   -74880
            TabIndex        =   31
            Top             =   1260
            Width           =   4335
         End
         Begin VB.Label Label2 
            Caption         =   "Bruteforce mode will repeatedly put in the permuted word to the current active window."
            Height          =   375
            Left            =   -74880
            TabIndex        =   30
            Top             =   780
            Width           =   4335
         End
         Begin VB.Label Label1 
            Caption         =   "Let me first Explain to you the purpose of these windows"
            Height          =   375
            Left            =   -74880
            TabIndex        =   29
            Top             =   480
            Width           =   4095
         End
      End
      Begin VB.Frame frmOMode 
         Caption         =   "Select Output Mode"
         Height          =   855
         Left            =   360
         TabIndex        =   24
         Top             =   3480
         Width           =   4335
         Begin VB.OptionButton optOverWrite 
            Caption         =   "Start with a new file (Overwrite)"
            Height          =   255
            Left            =   1560
            TabIndex        =   26
            Top             =   360
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton optAppend 
            Caption         =   "Append to file"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frmMode 
         Caption         =   "Select Mode to work with"
         Height          =   1335
         Left            =   -74640
         TabIndex        =   21
         Top             =   600
         Width           =   4335
         Begin VB.OptionButton chkAllPossibleMode 
            Caption         =   "All the damned possibilities"
            Height          =   195
            Left            =   240
            TabIndex        =   97
            ToolTipText     =   "Includes repeated letters. For eg to crack PASS. There are two S's. So you have to use this mode"
            Top             =   1080
            Width           =   3855
         End
         Begin VB.OptionButton optCPMode 
            Caption         =   "Both Combinations and Permutations"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            ToolTipText     =   "Al the combinations possible along with the Permutations"
            Top             =   720
            Value           =   -1  'True
            Width           =   3375
         End
         Begin VB.OptionButton optCombMode 
            Caption         =   "Combinations only"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            ToolTipText     =   "Only the Combinations"
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.Frame frmlength 
         Caption         =   "Select length of output Combinations"
         Height          =   2055
         Left            =   -74640
         TabIndex        =   13
         Top             =   600
         Width           =   4335
         Begin VB.OptionButton optVarying 
            Caption         =   "All Smaller and equal to fixed length"
            Height          =   375
            Left            =   360
            TabIndex        =   15
            ToolTipText     =   "For eg if you have specified a 4 letter word. All 4 letter words + 3 letter words + 2 letter word + 1 letter words are outputted"
            Top             =   1080
            Width           =   3375
         End
         Begin VB.OptionButton optFixed 
            Caption         =   "Fixed Length"
            Height          =   375
            Left            =   360
            TabIndex        =   14
            ToolTipText     =   "Just the fixed size as given in the length of combinations"
            Top             =   480
            Value           =   -1  'True
            Width           =   3495
         End
      End
      Begin VB.Frame frmCharacters 
         Caption         =   "Select Characters to use in combinations"
         Height          =   3735
         Left            =   -74640
         TabIndex        =   6
         Top             =   600
         Width           =   4335
         Begin VB.TextBox txtcustom 
            Height          =   285
            Left            =   1440
            TabIndex        =   20
            Text            =   "ABCD"
            Top             =   3240
            Width           =   2535
         End
         Begin VB.CheckBox chkCustom 
            Caption         =   "Custom"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            ToolTipText     =   "The custom part is added last. So you may give your one extra chars here"
            Top             =   3240
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkExcludeCC 
            Caption         =   "Exclude Control Characters"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "Usually in passwords the 1st 32 chars are not used since they are control characters."
            Top             =   2760
            Value           =   1  'Checked
            Width           =   3375
         End
         Begin VB.CheckBox chkAll255 
            Caption         =   "All ASCII upto 255"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   2280
            Width           =   3735
         End
         Begin VB.CheckBox chkAll127 
            Caption         =   "All upto ASCII 127"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1800
            Width           =   3735
         End
         Begin VB.CheckBox chkAllDigits 
            Caption         =   "All Digits"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1320
            Width           =   3735
         End
         Begin VB.CheckBox chkAllLow 
            Caption         =   "All Low Case Letters"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   3615
         End
         Begin VB.CheckBox chkAllUp 
            Caption         =   "All Upcase letters"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame frmOutput 
         Caption         =   "Select Output Option"
         Height          =   1815
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   4335
         Begin VB.OptionButton optByCode 
            Caption         =   "Power user: Manipulate sourcecode"
            Height          =   255
            Left            =   360
            TabIndex        =   98
            ToolTipText     =   $"frmMain.frx":075C
            Top             =   1440
            Width           =   3495
         End
         Begin VB.OptionButton optOutSendkeys 
            Caption         =   "To current Application (Brute Force)"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            ToolTipText     =   $"frmMain.frx":07F7
            Top             =   1080
            Width           =   3495
         End
         Begin VB.OptionButton optOutFile 
            Caption         =   "To output file on disk"
            Height          =   255
            Left            =   360
            TabIndex        =   4
            ToolTipText     =   "Be sure to specify the filename properly. I havent givent the eror check options"
            Top             =   720
            Width           =   3495
         End
         Begin VB.OptionButton optOutTextBox 
            Caption         =   "To inbuilt Multiline Text Box"
            Height          =   255
            Left            =   360
            TabIndex        =   3
            ToolTipText     =   "Output is directed to the text box along side"
            Top             =   360
            Value           =   -1  'True
            Width           =   3615
         End
      End
      Begin TabDlg.SSTab SSTabAbout 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   82
         Top             =   480
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   7011
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "By"
         TabPicture(0)   =   "frmMain.frx":08DF
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "frmAboutProgrammer"
         Tab(0).Control(1)=   "frmAboutContact"
         Tab(0).Control(2)=   "frmAboutOther"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Bugs"
         TabPicture(1)   =   "frmMain.frx":08FB
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frmAboutBugs"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Miscelaneous"
         TabPicture(2)   =   "frmMain.frx":0917
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frmAboutDedication"
         Tab(2).Control(1)=   "frmAboutMisc"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Improvement"
         TabPicture(3)   =   "frmMain.frx":0933
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "Label24"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "Label25"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Label26"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).ControlCount=   3
         Begin VB.Frame frmAboutOther 
            Caption         =   "Other Means"
            Height          =   1095
            Left            =   -74880
            TabIndex        =   104
            Top             =   2760
            Width           =   4335
            Begin VB.Label Label27 
               Caption         =   $"frmMain.frx":094F
               Height          =   735
               Left            =   240
               TabIndex        =   105
               Top             =   240
               Width           =   3975
            End
         End
         Begin VB.Frame frmAboutContact 
            Caption         =   "Contact"
            Height          =   975
            Left            =   -74850
            TabIndex        =   94
            Top             =   1680
            Width           =   4335
            Begin VB.Label Label11 
               Caption         =   "Web Site http://www.jofu.8m.com"
               Height          =   255
               Left            =   120
               TabIndex        =   96
               Top             =   600
               Width           =   3135
            End
            Begin VB.Label Label10 
               Caption         =   "email josephninan@crosswinds.net   liju_trv@yahoo.com"
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   240
               Width           =   4095
            End
         End
         Begin VB.Frame frmAboutProgrammer 
            Caption         =   "Programmer"
            Height          =   1305
            Left            =   -74850
            TabIndex        =   89
            Top             =   330
            Width           =   4335
            Begin VB.Label Label9 
               Caption         =   "Papanamcode, Trivandrum-18, Kerala, India"
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   960
               Width           =   3615
            End
            Begin VB.Label Label8 
               Caption         =   "Sree Chitra Thirunal College of Engineering"
               Height          =   255
               Left            =   120
               TabIndex        =   92
               Top             =   720
               Width           =   3735
            End
            Begin VB.Label Label7 
               Caption         =   "2nd Year BTech Computer Science and Engineering"
               Height          =   255
               Left            =   120
               TabIndex        =   91
               Top             =   480
               Width           =   3855
            End
            Begin VB.Label Label6 
               Caption         =   "Source code developed by Joseph Ninan"
               Height          =   255
               Left            =   120
               TabIndex        =   90
               Top             =   240
               Width           =   3375
            End
         End
         Begin VB.Frame frmAboutBugs 
            Caption         =   "Bugs detected so far"
            Height          =   3375
            Left            =   -74850
            TabIndex        =   87
            Top             =   450
            Width           =   4335
            Begin VB.Label Label28 
               Caption         =   $"frmMain.frx":09D9
               Height          =   975
               Left            =   120
               TabIndex        =   108
               Top             =   2280
               Width           =   4095
            End
            Begin VB.Label Label23 
               Caption         =   "Not sure whether All smaller words is compatible with the all damned words option. I have to check it out."
               Height          =   495
               Left            =   120
               TabIndex        =   100
               Top             =   1800
               Width           =   4095
            End
            Begin VB.Label Label22 
               Caption         =   $"frmMain.frx":0AE6
               Height          =   855
               Left            =   120
               TabIndex        =   99
               Top             =   960
               Width           =   4095
            End
            Begin VB.Label Label15 
               Caption         =   $"frmMain.frx":0B9E
               Height          =   615
               Left            =   120
               TabIndex        =   88
               Top             =   240
               Width           =   4095
            End
         End
         Begin VB.Frame frmAboutMisc 
            Caption         =   "Miscelaneous"
            Height          =   855
            Left            =   -74850
            TabIndex        =   85
            Top             =   450
            Width           =   4335
            Begin VB.Label Label16 
               Caption         =   "Time spent ont this program: Three days: Dec 29, 30, 31, 1999"
               Height          =   375
               Left            =   120
               TabIndex        =   86
               Top             =   240
               Width           =   3975
            End
         End
         Begin VB.Frame frmAboutDedication 
            Caption         =   "Dedication"
            Height          =   855
            Left            =   -74850
            TabIndex        =   83
            Top             =   1440
            Width           =   4335
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "To Someone whom i still love very much - I love her so much"
               ForeColor       =   &H000000C0&
               Height          =   495
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Width           =   3855
            End
         End
         Begin VB.Label Label26 
            Caption         =   $"frmMain.frx":0C37
            Height          =   1215
            Left            =   240
            TabIndex        =   103
            Top             =   2280
            Width           =   4455
         End
         Begin VB.Label Label25 
            Caption         =   $"frmMain.frx":0D77
            Height          =   615
            Left            =   240
            TabIndex        =   102
            Top             =   1680
            Width           =   4455
         End
         Begin VB.Label Label24 
            Caption         =   $"frmMain.frx":0E0D
            Height          =   1215
            Left            =   240
            TabIndex        =   101
            Top             =   480
            Width           =   4455
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Source code developed by Joseph Ninan
'2nd Year BTech Computer Science and Engineering
'Sree Chitra Thirunal College of Engineering
'Papanamcode, Trivandrum-18
'Affliated to University of Kerala
'Residential Address
'Liju Bhavan, Muttampuram Lane, Sreekariyam PO
'Trivandrum
'Kerala state
'India
'PIN 695017
'Tel No 0091-471-449977
'email josephninan@crosswinds.net   liju_trv@yahoo.com
'Web Site http://www.jofu.8m.com

Private Sub cmdBkSp_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{BACKSPACE}"

End Sub

Private Sub cmdCAPS_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{CAPSLOCK}"
End Sub

Private Sub cmdCaret_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{^}"
End Sub

Private Sub cmdClear_Click()
result = ""
frmMain.txtStatus.Text = ""
End Sub

Private Sub cmdDel_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{DELETE}"
End Sub

Private Sub cmdDown_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{DOWN}"
End Sub

Private Sub cmdEnd_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{END}"
End Sub

Private Sub cmdEnter_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{ENTER}"
End Sub

Private Sub cmdEscape_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{ESC}"
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdGenerate_Click()
StartTime = Time
leng = Val(frmMain.txtLength.Text)
    Initialize
    readinput   ' The results of this function are stored in char(0 to totalchar-1)

Select Case OChoice
Case 1:
    'frmMain.txtStatus.Text = ""
Case 2:
    If optOverWrite.Value = True Then
        Open frmMain.txtFilename.Text For Output As #1
    Else
        Open frmMain.txtFilename.Text For Append As #1
    End If
Case 3:
    dummy = MsgBox("Change the active window within the three seconds after pressing OK", vbOKOnly, "Bruteforce Alert")
    curtime = Timer
    While (Timer - curtime) < 3
        DoEvents
    Wend
End Select


If optFixed.Value = True Then
    GenOutput (leng)  ' This function uses the letters in char(1 to totalchar-1) to make combinations of length frmmain.txtlength.text
Else
    For MainCount = 1 To leng
        GenOutput (MainCount)
    Next MainCount
End If
'Timer1.Enabled = False
Select Case OChoice
Case 2:
    Close #1
End Select

End Sub
Public Sub readinput()
TotalChar = 0
NextBlock = 0
If chkAllUp.Value = 1 Then
    For i = 1 To TOTALUP
        char(NextBlock) = Chr(64 + i)
        NextBlock = NextBlock + 1
    Next i
End If
If chkAllLow.Value = 1 Then
    For i = 1 To TOTALLOW
        char(NextBlock) = Chr(96 + i)
        NextBlock = NextBlock + 1
    Next i
End If
If chkAllDigits.Value = 1 Then
  For i = 1 To TOTALDIGITS
        char(NextBlock) = Chr(47 + i)
        NextBlock = NextBlock + 1
    Next i
End If
If chkExcludeCC.Value = 1 Then ASCIIStart = 32 Else ASCIIStart = 0
Counter = 0
If chkAll127.Value = 1 Then
    For i = ASCIIStart To 127
        char(Counter) = Chr(i)
        Counter = Counter + 1
    Next i
    NextBlock = Counter
End If
Counter = 0
If chkAll255.Value = 1 Then
    For i = ASCIIStart To 255
        char(Counter) = Chr(i)
        Counter = Counter + 1
    Next i
    NextBlock = Counter
End If
If chkCustom.Value = 1 Then
    For i = 1 To Len(frmMain.txtcustom.Text)
        char(NextBlock) = Mid(frmMain.txtcustom.Text, i, 1)
        NextBlock = NextBlock + 1
    Next i
End If
TotalChar = NextBlock - 1
If chkAllPossibleMode.Value = True Then
    L = Val(txtLength.Text)
    For i = TotalChar + 1 To (TotalChar + 1) * L
        char(i) = char(i - TotalChar - 1)
    Next i
    TotalChar = (TotalChar + 1) * L - 1
End If
Debug.Print
Debug.Print "The next set of debug results"
For i = 0 To TotalChar: Debug.Print char(i); ;: Next i
Debug.Print
Debug.Print "Total 0 - "; TotalChar


End Sub

Private Sub cmdHome_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{HOME}"
End Sub

Private Sub cmdIns_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{INSERT}"
End Sub

Private Sub cmdLeft_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{LEFT}"
End Sub

Private Sub cmdPercent_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{%}"
End Sub

Private Sub cmdPgDn_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{PGDN}"
End Sub

Private Sub cmdPgUp_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{PGUP}"
End Sub

Private Sub cmdPlus_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{+}"
End Sub

Private Sub cmdRefresh_Click()
readinput
n = Val(TotalChar + 1)
r = Val(txtLength.Text)
txtNoComb.Text = Factorial(n) / Factorial(n - r) / Factorial(r)
txtNoPerm.Text = Factorial(n) / Factorial(n - r)

End Sub

Private Sub cmdRepeat_Click()
BFKeys = InputBox("Please input the key which you have to repeat", "Bruteforce - Repeat Key")
bftimes = InputBox("Please input the number of times you need to repeat this sequence", "Bruteforce - Repeat Times")
frmMain.txtBFSetup.Text = frmMain.txtBFSetup.Text & vbCrLf & "{" & BFKeys & " " & bftimes & "}"

End Sub

Private Sub cmdReset_Click()
txtBFSetup.Text = ""
txtBFSetup.Refresh
End Sub

Private Sub cmdRight_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{RIGHT}"
End Sub

Private Sub cmdStop_Click()
End
End Sub

Private Sub cmdTab_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{TAB}"
End Sub

Private Sub cmdTxtRefresh_Click()
txtStatus.Text = result
txtStatus.Refresh
End Sub

Private Sub cmdUp_Click()
txtBFSetup.Text = txtBFSetup.Text & vbCrLf & "{UP}"
End Sub


Private Sub Form_Load()
SSTabBF.Tab = 0
SSTabOptions.Tab = 0
SSTabAbout.Tab = 0
End Sub

Private Sub SSTabOptions_Click(PreviousTab As Integer)
If SSTabOptions.Tab = 3 Then cmdRefresh.Value = True
End Sub



Public Function Factorial(fn) As Double
fact = 1
For i = 1 To fn
    fact = fact * i
Next i
Factorial = fact
End Function

Private Sub txtBFSetup_GotFocus()
If chkCtrl.Value = 1 Then txtBFSetup.Text = txtBFSetup.Text & "^"
If chkShift.Value = 1 Then txtBFSetup.Text = txtBFSetup.Text & "+"
If chkAlt.Value = 1 Then txtBFSetup.Text = txtBFSetup.Text & "%"
End Sub

