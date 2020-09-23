VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6C410F08-CC5D-11D3-AFB0-B1F01529B83B}#1.10#0"; "AS-CTL2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   574
   StartUpPosition =   2  'CenterScreen
   Begin ASCmCtl2.asxStatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      Top             =   4740
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SizeGripView    =   2
      Style           =   1
      SimpleText      =   "Web Shelf 5 - Ready"
   End
   Begin TabDlg.SSTab sstTabs 
      Height          =   4575
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   529
      TabCaption(0)   =   "Images"
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lsvImgStats"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdFindPath"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtPath"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdRefresh"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Dimensions"
      TabPicture(1)   =   "frmMain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblImagesHigh"
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(2)=   "lblImagesAcross"
      Tab(1).Control(3)=   "Label11"
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(5)=   "lblNoImages"
      Tab(1).Control(6)=   "sliWSImagesAcross"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Appearance"
      TabPicture(2)   =   "frmMain.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label30"
      Tab(2).Control(1)=   "Label29"
      Tab(2).Control(2)=   "Label28"
      Tab(2).Control(3)=   "Label27"
      Tab(2).Control(4)=   "Label2"
      Tab(2).Control(5)=   "sliWSCellSpacing"
      Tab(2).Control(6)=   "sliWSCellPadding"
      Tab(2).Control(7)=   "sliBorderSize"
      Tab(2).Control(8)=   "txtWSTitle"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Fonts"
      TabPicture(3)   =   "frmMain.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label4"
      Tab(3).Control(1)=   "Frame4"
      Tab(3).Control(2)=   "Frame3"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Resizing"
      TabPicture(4)   =   "frmMain.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label14"
      Tab(4).Control(1)=   "Frame1"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Text"
      TabPicture(5)   =   "frmMain.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label8"
      Tab(5).Control(1)=   "Label31"
      Tab(5).Control(2)=   "Frame5"
      Tab(5).Control(3)=   "chkShowDimensions"
      Tab(5).Control(4)=   "chkShowFileType"
      Tab(5).Control(5)=   "chkShowFileSize"
      Tab(5).Control(6)=   "chkShowFileName"
      Tab(5).Control(7)=   "asxOContainer"
      Tab(5).ControlCount=   8
      TabCaption(6)   =   "User Info"
      TabPicture(6)   =   "frmMain.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label19"
      Tab(6).Control(1)=   "Label5"
      Tab(6).Control(2)=   "cmdSaveUserInfo"
      Tab(6).Control(3)=   "cmdOpenUserInfo"
      Tab(6).Control(4)=   "Frame2"
      Tab(6).Control(5)=   "fraUserOpts"
      Tab(6).ControlCount=   6
      Begin VB.Frame fraUserOpts 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   78
         Top             =   720
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox txtUserName 
            Height          =   285
            Left            =   1440
            TabIndex        =   82
            Top             =   240
            Width           =   4815
         End
         Begin VB.TextBox txtUserCopyright 
            Height          =   285
            Left            =   1440
            TabIndex        =   81
            Top             =   600
            Width           =   4815
         End
         Begin VB.TextBox txtUserComments 
            Height          =   615
            Left            =   1440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   80
            Top             =   960
            Width           =   4815
         End
         Begin VB.CheckBox chkUserIncludeDate 
            Caption         =   "Include the date the Web Shelf was created"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   1680
            Width           =   4215
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Name / Company"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   285
            Width           =   1230
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Image copyright"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   645
            Width           =   1170
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Comments:"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   975
            Width           =   810
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   6105
      End
      Begin VB.CommandButton cmdFindPath 
         Height          =   280
         Left            =   6360
         Picture         =   "frmMain.frx":0506
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   320
      End
      Begin VB.Frame Frame2 
         Caption         =   "Information Positioning"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74760
         TabIndex        =   73
         Top             =   2880
         Width           =   6375
         Begin VB.OptionButton optNoUser 
            Caption         =   "No user information!"
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optUserAlignTop 
            Caption         =   "Top"
            Height          =   195
            Left            =   2760
            TabIndex        =   36
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optUserAlignBottom 
            Caption         =   "Bottom"
            Height          =   195
            Left            =   4440
            TabIndex        =   37
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdOpenUserInfo 
         Caption         =   "Import..."
         Height          =   375
         Left            =   -74760
         TabIndex        =   38
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveUserInfo 
         Caption         =   "Save..."
         Height          =   375
         Left            =   -73320
         TabIndex        =   39
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Frame asxOContainer 
         Caption         =   "Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -72120
         TabIndex        =   70
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
         Begin VB.OptionButton optOAscending 
            Caption         =   "Ascending"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optODescending 
            Caption         =   "Descending"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkShowFileName 
         Caption         =   "Show filename"
         Height          =   195
         Left            =   -74760
         TabIndex        =   26
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkShowFileSize 
         Caption         =   "Show file size"
         Height          =   195
         Left            =   -71640
         TabIndex        =   27
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CheckBox chkShowFileType 
         Caption         =   "Show file type (bitmap, GIF or JPEG)"
         Height          =   195
         Left            =   -74760
         TabIndex        =   28
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CheckBox chkShowDimensions 
         Caption         =   "Show image dimensions"
         Height          =   195
         Left            =   -71640
         TabIndex        =   29
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Frame Frame5 
         Height          =   1095
         Left            =   -74400
         TabIndex        =   69
         Top             =   2400
         Width           =   2055
         Begin VB.OptionButton optNoImageSort 
            Caption         =   "No sorting"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optSortByName 
            Caption         =   "Sort by filename"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton optSortBySize 
            Caption         =   "Sort by file size"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   64
         Top             =   840
         Width           =   6375
         Begin VB.OptionButton optNoResize 
            Caption         =   "Leave images as they are"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optResize 
            Caption         =   "Scale all images... (recommended)"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   480
            Width           =   2775
         End
         Begin ASCmCtl2.asxContainer Container1 
            Height          =   2175
            Left            =   480
            Top             =   720
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   3836
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            CaptionStyle    =   0
            Begin VB.CheckBox chkThumbs 
               Caption         =   "Create thumbnail images"
               Height          =   195
               Left            =   240
               TabIndex        =   24
               Top             =   1560
               Width           =   3735
            End
            Begin VB.TextBox txtImgToWidth 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   840
               TabIndex        =   19
               Text            =   "80"
               Top             =   165
               Width           =   375
            End
            Begin VB.TextBox txtImgToHeight 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1920
               TabIndex        =   20
               Text            =   "60"
               Top             =   165
               Width           =   375
            End
            Begin VB.CheckBox chkLimitResize 
               Caption         =   "Don't resize an image if its dimensions are smaller"
               Height          =   195
               Left            =   240
               TabIndex        =   21
               Top             =   480
               Value           =   1  'Checked
               Width           =   3855
            End
            Begin VB.CheckBox chkLinktoFullSize 
               Caption         =   "Link to the full-scale image if resized"
               Height          =   195
               Left            =   240
               TabIndex        =   25
               Top             =   1800
               Value           =   1  'Checked
               Width           =   3975
            End
            Begin ASCmCtl2.asxContainer Container2 
               Height          =   735
               Left            =   480
               Top             =   720
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   1296
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   ""
               CaptionStyle    =   0
               Begin VB.OptionButton optNoResize2 
                  Caption         =   "If BOTH are smaller"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   23
                  Top             =   405
                  Width           =   3015
               End
               Begin VB.OptionButton optNoResize1 
                  Caption         =   "If the Width OR Height is smaller"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   22
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   3015
               End
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Width"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   67
               Top             =   210
               Width           =   495
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Height"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1320
               TabIndex        =   66
               Top             =   210
               Width           =   555
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "pixels"
               Height          =   195
               Left            =   2400
               TabIndex        =   65
               Top             =   210
               Width           =   405
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "User and Image information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74760
         TabIndex        =   57
         Top             =   960
         Width           =   6375
         Begin VB.ComboBox cboUserFontSize 
            Height          =   315
            ItemData        =   "frmMain.frx":0650
            Left            =   1680
            List            =   "frmMain.frx":065A
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1040
            Width           =   855
         End
         Begin VB.CommandButton cmdChooseFont 
            Caption         =   "Choose"
            Height          =   255
            Index           =   0
            Left            =   4560
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdChooseFont 
            Caption         =   "Choose"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   4560
            TabIndex        =   12
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Primary font"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   240
            TabIndex        =   62
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Secondary font"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   240
            TabIndex        =   61
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Font size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   240
            TabIndex        =   60
            Top             =   1080
            Width           =   750
         End
         Begin VB.Label lblUserFont 
            AutoSize        =   -1  'True
            Caption         =   "Default browser font"
            Height          =   195
            Index           =   0
            Left            =   1680
            TabIndex        =   59
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblUserFont 
            AutoSize        =   -1  'True
            Caption         =   "Default browser font"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   58
            Top             =   720
            Width           =   1500
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Web Shelf title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74760
         TabIndex        =   51
         Top             =   2880
         Width           =   6375
         Begin VB.CommandButton cmdChooseTitleFont 
            Caption         =   "Choose"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   4560
            TabIndex        =   15
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdChooseTitleFont 
            Caption         =   "Choose"
            Height          =   255
            Index           =   0
            Left            =   4560
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cboTitleFontSize 
            Height          =   315
            ItemData        =   "frmMain.frx":066A
            Left            =   1680
            List            =   "frmMain.frx":067D
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1040
            Width           =   855
         End
         Begin VB.Label lblTitleFont 
            AutoSize        =   -1  'True
            Caption         =   "Default browser font"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   56
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblTitleFont 
            AutoSize        =   -1  'True
            Caption         =   "Default browser font"
            Height          =   195
            Index           =   0
            Left            =   1680
            TabIndex        =   55
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Font size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   240
            TabIndex        =   54
            Top             =   1080
            Width           =   750
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Secondary font"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   240
            TabIndex        =   53
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Primary font"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   240
            TabIndex        =   52
            Top             =   360
            Width           =   1065
         End
      End
      Begin VB.TextBox txtWSTitle 
         Height          =   285
         Left            =   -74760
         TabIndex        =   7
         Text            =   "A Collection of Images"
         Top             =   1200
         Width           =   6135
      End
      Begin MSComctlLib.Slider sliWSImagesAcross 
         Height          =   255
         Left            =   -74760
         TabIndex        =   6
         Top             =   1800
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         Min             =   1
         SelStart        =   6
         Value           =   6
      End
      Begin MSComctlLib.Slider sliBorderSize 
         Height          =   255
         Left            =   -74760
         TabIndex        =   8
         Top             =   2160
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         SelStart        =   4
         Value           =   4
      End
      Begin MSComctlLib.Slider sliWSCellPadding 
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   2880
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         SelStart        =   2
         Value           =   2
      End
      Begin MSComctlLib.Slider sliWSCellSpacing 
         Height          =   255
         Left            =   -74760
         TabIndex        =   10
         Top             =   3600
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.ListView lsvImgStats 
         Height          =   1815
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Object.Width           =   6703
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Images to be used"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   77
         Top             =   480
         Width           =   2130
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Path of images to include:"
         Height          =   195
         Left            =   240
         TabIndex        =   76
         Top             =   840
         Width           =   1860
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User information"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         TabIndex        =   75
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   $"frmMain.frx":069F
         Height          =   435
         Left            =   -74760
         TabIndex        =   74
         Top             =   4080
         Width           =   6195
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sorting"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         TabIndex        =   72
         Top             =   2040
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Textual Information"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         TabIndex        =   71
         Top             =   480
         Width           =   2205
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resizing the selected images"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         TabIndex        =   68
         Top             =   480
         Width           =   3300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Shelf's use of Fonts"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         TabIndex        =   63
         Top             =   480
         Width           =   2625
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Shelf's Appearance"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         TabIndex        =   50
         Top             =   480
         Width           =   2850
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Title of Web Shelf (appears on top)"
         Height          =   195
         Left            =   -74760
         TabIndex        =   49
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Edge of Web Shelf size"
         Height          =   195
         Left            =   -74760
         TabIndex        =   48
         Top             =   1920
         Width           =   1650
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Cell padding"
         Height          =   195
         Left            =   -74760
         TabIndex        =   47
         Top             =   2640
         Width           =   870
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Cell spacing"
         Height          =   195
         Left            =   -74760
         TabIndex        =   46
         Top             =   3360
         Width           =   840
      End
      Begin VB.Label lblNoImages 
         AutoSize        =   -1  'True
         Caption         =   $"frmMain.frx":073C
         ForeColor       =   &H00000080&
         Height          =   390
         Left            =   -74760
         TabIndex        =   45
         Top             =   840
         Width           =   6375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Shelf's Dimensions"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         TabIndex        =   44
         Top             =   480
         Width           =   2640
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Number of images across"
         Height          =   195
         Left            =   -74760
         TabIndex        =   43
         Top             =   1560
         Width           =   1800
      End
      Begin VB.Label lblImagesAcross 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71760
         TabIndex        =   42
         Top             =   1515
         Width           =   135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Calculated number of images high"
         Height          =   195
         Left            =   -74760
         TabIndex        =   41
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label lblImagesHigh 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71760
         TabIndex        =   40
         Top             =   2235
         Width           =   135
      End
   End
   Begin ASCmCtl2.asxActionBar asxActionBar 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   5106
      ShowGroupHeaders=   0   'False
      PlaySounds      =   0   'False
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GroupFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolTipBackColor=   16765650
      SmallIcons      =   "asxImageList1"
      GroupCount      =   1
      Group1Caption   =   "Options"
      Group1ClickStyle=   1
      Group1IconSize  =   1
      Group1Selected  =   -1  'True
      Group1Style     =   2
      Group1ItemCount =   8
      Group1Item1Caption=   "New"
      Group1Item1Key  =   "New"
      Group1Item2Caption=   "Open"
      Group1Item2Key  =   "Open"
      Group1Item3Caption=   "Save"
      Group1Item3Key  =   "Save"
      Group1Item4Caption=   "Save As"
      Group1Item4Key  =   "Save As"
      Group1Item5Caption=   "Make!"
      Group1Item5Key  =   "Generate"
      Group1Item6Caption=   "Help"
      Group1Item6Key  =   "Help"
      Group1Item7Caption=   "About"
      Group1Item7Key  =   "About"
      Group1Item8Caption=   "Quit"
      Group1Item8Key  =   "Quit"
   End
   Begin MSComDlg.CommonDialog dlgWebShelf 
      Left            =   600
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".wsp"
      DialogTitle     =   "Save Web Shelf parameters"
      Filter          =   "Web Shelf definitions|*.wsp"
      Flags           =   2656262
   End
   Begin MSComDlg.CommonDialog dlgUserInfo 
      Left            =   720
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".wsi"
      DialogTitle     =   "Open user information settings"
      Filter          =   "Web Shelf User Information|*.wsi"
      Flags           =   2656262
   End
   Begin ASCmCtl2.asxImageList asxImageList1 
      Left            =   120
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Image imgDimension 
      Height          =   375
      Left            =   1200
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'  frmMain.frm for Web Shelf 5
'  created by Andrew (aka BadHart)
'  Source submitted to Planet Source Code

'Believe me, this was a nightmare to comment!

Private Sub asxActionBar_ItemClick(GroupItem As ASCmCtl2.GroupItem)
    'We pressed a button on the action bar, so we need to find out which
    'and act accordingly. We can tell by the key.
    
    Dim a As String
    
    Select Case GroupItem.Key
    Case "Generate" 'Generate a web page
    
        'First check if we have any images at all.
        If ImageCount = 0 Then
            MsgBox "You cannot make a web page until you have allocated images."
            sstTabs.Tab = 5
        Else
            frmOptions.Show vbModal
        End If
        
    Case "Quit" 'Exit the program
        Unload Me
        
    Case "About" 'Show the About form
        frmAbout.Show vbModal
        
    Case "New"
        'If the existing settings have changed, then verify a new Web Shelf.
        If Changed Then
            If MsgBox("Woah! Do you want to save your existing settings?", _
                vbYesNo + vbQuestion, "New Web Shelf") = vbYes Then
                SaveSettings dlgWebShelf.filename
            End If
            NewSettings
        Else
            NewSettings
        End If
        
    Case "Save"
        On Error GoTo whoops
        With dlgWebShelf
            If .filename = Empty Then
                .DialogTitle = "Save Web Shelf parameters"
                .ShowSave
            End If
            Status.SimpleText = "Saving Web Shelf parameters..."
            a = SaveSettings(.filename)
            Status.SimpleText = IIf(a = Empty, "Web Shelf 5 - settings saved.", a)
            Changed = False
        End With
    
    Case "Save As"
        On Error GoTo whoops
        With dlgWebShelf
            .DialogTitle = "Save Web Shelf parameters"
            .ShowSave
            Status.SimpleText = "Saving Web Shelf parameters..."
            a = SaveSettings(.filename)
            Status.SimpleText = IIf(a = Empty, "Web Shelf 5 - settings saved.", a)
            Changed = False
        End With

    Case "Open"
        On Error GoTo whoops
        With dlgWebShelf
            .DialogTitle = "Open Web Shelf parameters"
            .ShowOpen
            Status.SimpleText = "Getting Web Shelf parameters..."
            a = OpenSettings(.filename)
            Status.SimpleText = IIf(a = Empty, "Web Shelf 5 - Ready", a)
        End With
        
    End Select
    
whoops:
    asxActionBar.Refresh
End Sub

Private Sub cboTitleFontSize_Click()
    Changed = True
End Sub

Private Sub cboUserFontSize_Click()
    Changed = True
End Sub

Private Sub chkLimitResize_Click()
    'Enable or disable further options.
    Container2.Visible = (chkLimitResize.Value = 1)
    Changed = True
End Sub

Private Sub chkLinktoFullSize_Click()
    Changed = True
End Sub

Private Sub chkShowDimensions_Click()
    Changed = True
End Sub

Private Sub chkShowFileName_Click()
    Changed = True
End Sub

Private Sub chkShowFileSize_Click()
    Changed = True
End Sub

Private Sub chkShowFileType_Click()
    Changed = True
End Sub

Private Sub chkThumbs_Click()
    Changed = True
End Sub

Private Sub cmdChooseFont_Click(Index As Integer)
    'Set the flags: no initial font selected and only screen fonts. Courtesy of
    'one of my other programs, Come On! Dialog; download it at
    'badhart.tripod.com/badsoft.
    
    On Error GoTo whoops
    
    With dlgUserInfo
        .Flags = 524289
        .ShowFont
        'We only need the font name.
        lblUserFont(Index) = .FontName
        'If the first font is not the default, enable choosing of the other font.
        If Index = 0 Then
            lblUserFont(1).Enabled = True
            cmdChooseFont(1).Enabled = True
        End If
        Changed = True
    End With
    
whoops:
End Sub

Private Sub cmdChooseTitleFont_Click(Index As Integer)
    'Same as before.
    On Error GoTo whoops
    
    With dlgUserInfo
        .Flags = 524289
        .ShowFont
        'We only need the font name.
        lblTitleFont(Index) = .FontName
        'If the first font is not the default, enable choosing of the other font.
        If Index = 0 Then
            lblTitleFont(1).Enabled = True
            cmdChooseTitleFont(1).Enabled = True
        End If
        Changed = True
    End With
    
whoops:
End Sub

Private Sub cmdFindPath_Click()
    'Looks for a folder on the computer, and if we get a new path then
    'change the selected path.
    Dim NewPath As String
    NewPath = BrowseForFolder(hWnd, _
        "Select an folder with web graphics in them...")
        
    'Only do this part if we have images.
    If Len(NewPath) > 0 Then
        txtPath.Text = NewPath
        ListImages NewPath
        
        If NoImages Then
            MsgBox "No images found in folder!"
            Exit Sub
        End If
    
    End If
    
End Sub

Private Sub cmdOpenUserInfo_Click()
    'Open our user information. But where is it?
    Dim a As Integer, temp As String, v As Integer
    On Error GoTo whoops
    
    With dlgUserInfo
        .Flags = 2656262 'courtesy of Come On! Dialog
        .ShowOpen
        a = FreeFile
        'Open the information file...
        Open .filename For Input As #a
        
        'Now put all of the pre-entered fields into the respective boxes.
        Input #a, temp:         frmMain.txtUserName = temp
        Input #a, temp:         frmMain.txtUserCopyright = temp
        Input #a, temp:         frmMain.txtUserComments = temp
        Input #a, v:                frmMain.chkUserIncludeDate = v
        
        'Close the file (v. important!)
        Close #a
        
        'If we've set user information to none, them we automatically set it to
        'bottom.
        If optNoUser Then optUserAlignBottom = True
    End With
    
    Beep 'wake the user up.
    
whoops:
End Sub

Private Sub cmdRefresh_Click()
    'Refreshes the image statistics.
    ListImages txtPath
    
    'Make any changes to the slider.
    With frmMain.sliWSImagesAcross
        .Enabled = Not NoImages
        .Max = ImageCount()
        
        'Set a default size for images across. This should arouse
        'calculation of the height.
        .Max = Limit(.Max, 2, 12)
        sliWSImagesAcross_Change
    End With

End Sub

Private Sub cmdSaveUserInfo_Click()
    On Error GoTo whoops
    With dlgUserInfo
        .Flags = 2656262    'guess who!
        .Filter = "Web Shelf User Information|*.wsi"
        .DialogTitle = "Save user information settings"
        .ShowSave
        'Time to save the details...
        Dim a As Integer
        a = FreeFile
        Open .filename For Output As #a
        Write #a, frmMain.txtUserName
        Write #a, frmMain.txtUserCopyright
        Write #a, frmMain.txtUserComments
        Write #a, frmMain.chkUserIncludeDate
        Close #a
    End With
    MsgBox "User information saved.", vbInformation
whoops:
End Sub

Private Sub Form_Load()
    'As soon as the form loads we set the defaults.
    Caption = App.Title & " " & App.Major
    cboUserFontSize.Text = "10pt"
    cboTitleFontSize.Text = "18pt"
    Changed = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If the existing settings have changed, then ask to quit.
    If Changed Then
        Select Case MsgBox("Woah! Do you want to save your existing settings?", _
            vbYesNoCancel + vbQuestion, "Exit the program")
        Case vbYes 'answered Yes
            SaveSettings dlgWebShelf.filename
            End
        Case vbNo 'answered No
            End
        Case vbCancel 'answered Cancel
            Cancel = 1 'stops the main form from unloading.
        End Select
    Else
        End
    End If
End Sub

Private Sub lblTitleFont_Click(Index As Integer)
    'A click changes the font to the default.
    lblTitleFont(Index) = DefaultFont
    If Index = 0 Then
        lblTitleFont(1) = DefaultFont
        lblTitleFont(1).Enabled = False
        cmdChooseTitleFont(1).Enabled = False
        Changed = True
    End If
End Sub

Private Sub lblUserFont_Click(Index As Integer)
    lblUserFont(Index) = DefaultFont
    If Index = 0 Then
        lblUserFont(1) = DefaultFont
        lblUserFont(1).Enabled = False
        cmdChooseFont(1).Enabled = False
        Changed = True
    End If
End Sub

Private Sub optNoImageSort_Click()
    asxOContainer.Visible = False
    Changed = True
End Sub

Private Sub optNoResize_Click()
    'Disable or make invisible the resize options.
    Container1.Visible = False
    Changed = True
End Sub

Private Sub optNoResize1_Click()
    Changed = True
End Sub

Private Sub optNoResize2_Click()
    Changed = True
End Sub

Private Sub optNoUser_Click()
    fraUserOpts.Visible = False
    Changed = True
End Sub

Private Sub optOAscending_Click()
    Changed = True
End Sub

Private Sub optODescending_Click()
    Changed = True
End Sub

Private Sub optResize_Click()
    Container1.Visible = True
    Changed = True
End Sub

Private Sub optSortByName_Click()
    asxOContainer.Visible = True
    Changed = True
End Sub

Private Sub optSortBySize_Click()
    asxOContainer.Visible = True
    Changed = True
End Sub

Private Sub optUserAlignBottom_Click()
    fraUserOpts.Visible = True
    Changed = True
End Sub

Private Sub optUserAlignTop_Click()
    fraUserOpts.Visible = True
    Changed = True
End Sub

Private Sub sliBorderSize_Change()
    Changed = True

End Sub

Private Sub sliWSCellPadding_Change()
    Changed = True
End Sub

Private Sub sliWSCellSpacing_Change()
    Changed = True
End Sub

Private Sub sliWSImagesAcross_Change()
    'To find the height of the Web Shelf in images we need to apply a very
    'tricky calculation. Four times before had I needed to figure out what it
    'was.
    Dim c As Integer, m As Integer
    c = sliWSImagesAcross.Value
    lblImagesAcross = c
    m = ImageCount()
    'Do the calculation and then show it.
    lblImagesHigh = m \ c + IIf(m Mod c > 0, 1, 0)
    Changed = True
End Sub

Private Sub txtImgToHeight_Change()
    Changed = True
End Sub

Private Sub txtImgToHeight_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    Else
        Changed = True
    End If
End Sub

Private Sub txtImgToWidth_Change()
    Changed = True
End Sub

Private Sub txtImgToWidth_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    Else
        Changed = True
    End If
End Sub

Private Sub txtWSTitle_Change()
    Changed = True
End Sub
