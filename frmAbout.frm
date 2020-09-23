VERSION 5.00
Object = "{6C410F08-CC5D-11D3-AFB0-B1F01529B83B}#1.10#0"; "AS-CTL2.OCX"
Object = "{ED6EFBE9-2DE5-11D2-9B4A-006097731E48}#1.0#0"; "HLINK.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4170
   ClientLeft      =   2340
   ClientTop       =   1650
   ClientWidth     =   5640
   ClipControls    =   0   'False
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   278
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ASCmCtl2.asxHeader asxHeader1 
      Height          =   255
      Left            =   0
      Top             =   720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      Style           =   2
      GradientStartColor=   16777215
      GradientEndColor=   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "The Good News"
   End
   Begin HyperLinkControl.HLink HLink3 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmAbout.frx":0442
      MousePointer    =   14
      Caption         =   "badhart@hotpop.com"
      URL             =   "mailto:badhart@hotpop.com"
   End
   Begin HyperLinkControl.HLink HLink2 
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmAbout.frx":075C
      MousePointer    =   14
      Caption         =   "starts at badhart.tripod.com"
      URL             =   "http://badhart.tripod.com"
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
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
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   3720
      Width           =   1260
   End
   Begin HyperLinkControl.HLink HLink4 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmAbout.frx":0A76
      MousePointer    =   14
      Caption         =   "http://users.globalnet.co.uk/~ariad"
      URL             =   "http://users.globalnet.co.uk/~ariad"
   End
   Begin HyperLinkControl.HLink HLink5 
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   2280
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmAbout.frx":0D90
      MousePointer    =   14
      Caption         =   "http://www.tbrown.dircon.co.uk/vb"
      URL             =   "http://www.tbrown.dircon.co.uk/vb"
   End
   Begin ASCmCtl2.asxHeader asxHeader3 
      Height          =   255
      Left            =   0
      Top             =   2760
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      Style           =   2
      GradientStartColor=   16777215
      GradientEndColor=   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Author's Info"
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "©2000 BadSoft, All rights reserved."
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   2595
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tom Brown"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ariad Software"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This program uses Ariad's Common Controls 2 and the Hyperlink control by Tom Brown."
      Height          =   390
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   5520
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":10AA
      Height          =   585
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   5445
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Electronic Mail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BadSoft Online"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3705
      TabIndex        =   1
      Top             =   120
      Width           =   1860
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   195
      Left            =   5040
      TabIndex        =   2
      Top             =   480
      Width           =   525
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'  frmAbout.frm for Web Shelf 5
'  created by Andrew (aka BadHart)
'  Source submitted to Planet Source Code

Private Sub cmdOK_Click()
    'Close and unload the form.
    Unload Me
End Sub

Private Sub Form_Load()
    'This displays all of the program information.
    'Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & _
        App.Revision & " by Andrew (aka BadHart)"
    lblTitle.Caption = App.Title & " " & App.Major
    imgIcon.Picture = Me.Icon
End Sub

