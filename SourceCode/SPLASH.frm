VERSION 5.00
Begin VB.Form SPLASH 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   525
      Top             =   1470
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   105
      Top             =   1470
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      Picture         =   "SPLASH.frx":0000
      ScaleHeight     =   1215
      ScaleWidth      =   7560
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filing Cabinet Control"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   36
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   825
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   6570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EDP IPT"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   6600
         TabIndex        =   1
         Top             =   720
         Width           =   960
      End
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING....."
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   585
      TabIndex        =   7
      Top             =   2580
      Width           =   2355
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adi Jaya Sarana"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   2880
      TabIndex        =   6
      Top             =   1260
      Width           =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail: ajs.indonesia@gmail.com"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   120
      Left            =   2385
      TabIndex        =   5
      Top             =   1845
      Width           =   2790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COPYRIGHT 2008"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   0
      Left            =   3150
      TabIndex        =   4
      Top             =   1620
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "thomas_edp2006@yahoo.co.id"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   1
      Left            =   2610
      TabIndex        =   3
      Top             =   2100
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   180
      Picture         =   "SPLASH.frx":0B77
      Top             =   2565
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   -60
      Top             =   2415
      Width           =   7680
   End
End
Attribute VB_Name = "SPLASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
MakeTransparent Me.hwnd, 0
Timer1.Enabled = True
Timer2.Enabled = False
A = 0
End Sub

Private Sub Timer1_Timer()
    A = A + 10
    AA = AA + 10
    If A <= 600 Then
        MakeTransparent Me.hwnd, A
        Label17.ForeColor = AA
    Else
        Timer1.Enabled = False
        Timer2.Enabled = True
        A = 600
    End If
End Sub

Private Sub Timer2_Timer()
    A = A - 10
    If A >= 0 Then
        MakeTransparent Me.hwnd, A
    Else
        Timer1.Enabled = False
        Timer2.Enabled = False
        Unload Me
        LOGIN.Show
    End If
End Sub
