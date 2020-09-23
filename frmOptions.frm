VERSION 5.00
Object = "{6C410F08-CC5D-11D3-AFB0-B1F01529B83B}#1.10#0"; "AS-CTL2.OCX"
Begin VB.Form frmOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Web Shelf"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   262
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   2400
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2040
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin ASCmCtl2.asxProgressBar asxProgress 
      Height          =   135
      Left            =   120
      Top             =   3720
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   238
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filename"
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
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Text            =   "images"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   ".html"
         Height          =   195
         Left            =   3480
         TabIndex        =   13
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Web Shelf filename:"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   285
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Okay"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Generation Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7095
      Begin VB.CheckBox chkRename 
         Caption         =   "Change to ""Internet friendly"" filenames"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   3375
      End
      Begin VB.Frame fraOptions 
         Height          =   1095
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox txtNewFolder 
            Height          =   285
            Left            =   2160
            TabIndex        =   16
            Top             =   720
            Width           =   1935
         End
         Begin ASCmCtl2.asxButton cmdBrowse 
            Height          =   315
            Left            =   5280
            TabIndex        =   5
            Top             =   400
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            Picture         =   "frmOptions.frx":0000
            CaptionAlignment=   5
            PictureAlignment=   3
            Caption         =   "Browse"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtPath 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   400
            Width           =   5055
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "New folder name (optional)"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   765
            Width           =   1950
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Choose an existing folder..."
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   160
            Width           =   2010
         End
      End
      Begin VB.OptionButton optNewDir 
         Caption         =   "Create in another directory"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton optOldDir 
         Caption         =   "Create in original directory"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status of generation"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   1485
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'  frmOptions.frm for Web Shelf 5
'  created by Andrew (aka BadHart)
'  Source submitted to Planet Source Code

Private Sub cmdBrowse_Click()
    'Look for a folder using the BrowseForFolder function.
    Dim temp As String
    
    temp = BrowseForFolder(Me.hWnd, _
        "Select an existing folder where the Web Shelf is to be created and images to be copied...")
    
    'Accept the selected path if there is one.
    If temp <> Empty Then
        txtPath = temp
    End If
End Sub

Private Sub cmdCancel_Click()
    'Just cancel and hide the form.
    frmOptions.Hide
End Sub

Private Sub cmdOkay_Click()
    'Start the generation!
    'Note that the Web Shelf HTML file is also renamed if we so require.
    If Me.chkRename = True Then
        txtFilename = IFriendly(txtFilename)
    End If
    
    'The whole process starts here...
    DoGenerate
    'Now close the form.
    Me.Hide
End Sub

Private Sub Form_Activate()
    'Get rid of the text near the status bar.
    lblStatus.Caption = Empty
    asxProgress.Value = 0
End Sub

Private Sub optNewDir_Click()
    'Reveals the extra options.
    fraOptions.Visible = True
End Sub

Private Sub optOldDir_Click()
    'Hides the options.
    fraOptions.Visible = False
End Sub
