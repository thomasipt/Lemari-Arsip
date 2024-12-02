VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form FRM_SOS 
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "FRM_SOS.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1575
      Top             =   8460
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3750
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   8100
      Width           =   4515
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1080
      Top             =   8460
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   8460
   End
   Begin XPControls.XPButton CmdKeluar 
      Height          =   480
      Left            =   10035
      TabIndex        =   1
      Top             =   1620
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   847
      Caption         =   "EXIT"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
      SoundOnClick    =   1
      ColorScheme     =   1
      ColorBegin      =   16777215
      ColorEnd        =   16777088
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Not Connected..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4965
      TabIndex        =   2
      Top             =   7065
      Width           =   2025
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   4365
      Picture         =   "FRM_SOS.frx":15F942
      Stretch         =   -1  'True
      Top             =   5535
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   7200
      Picture         =   "FRM_SOS.frx":15FC4C
      Stretch         =   -1  'True
      Top             =   5535
      Width           =   480
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   225
      Shape           =   3  'Circle
      Top             =   8595
      Width           =   285
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   90
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   555
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   112
      Shape           =   3  'Circle
      Top             =   6120
      Width           =   510
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   135
      Shape           =   3  'Circle
      Top             =   6705
      Width           =   465
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   420
      Left            =   157
      Shape           =   3  'Circle
      Top             =   7245
      Width           =   420
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   180
      Shape           =   3  'Circle
      Top             =   7740
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   330
      Left            =   202
      Shape           =   3  'Circle
      Top             =   8190
      Width           =   330
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   7200
      Picture         =   "FRM_SOS.frx":15FF56
      Stretch         =   -1  'True
      Top             =   4275
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   4365
      Picture         =   "FRM_SOS.frx":160260
      Stretch         =   -1  'True
      Top             =   4275
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4365
      Picture         =   "FRM_SOS.frx":16056A
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7200
      Picture         =   "FRM_SOS.frx":160874
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   480
   End
End
Attribute VB_Name = "FRM_SOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private SqlPass As String
Private tUser As rdoResultset
Private tMasuk As rdoResultset

Private RDEl, RSave, RSave2, RSave3, RSave4, RCari, RCari2 As rdoResultset
Private SDel, SSave, SSave2, SSave3, SSave4, SCari, SCari2 As String

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_FINDSTRING = &H14C
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Sub cmdKELUAR_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
'Screen.MousePointer = vbHourglass
'FRM_PROSES.Show

Unload Me
MAINMENU.Show

Unload FRM_PROSES
Set FRM_PROSES = Nothing
'Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=LOCKER", rdDriverNoPrompt, False, CN)

Image1.Visible = False
Image4.Visible = False
Image5.Visible = False
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False

Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
A = 0

Combo1.AddItem "LEMARI 1", 0
Combo1.ListIndex = 0
        
End Sub

Private Sub Image2_Click()
SSave2 = "Select * From B001_HIS2"
Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
RSave2.AddNew
    RSave2("ID") = 0
    RSave2("NOPOL") = 0
    RSave2("TANGGAL_HIS") = Date
    RSave2("JAM_HIS") = Time
    RSave2("KETERANGAN_HIS") = "SOS LOCKER 1"
RSave2.Update
RSave2.Close
Set RSave2 = Nothing
    
If Image3.Visible = False Or Image6.Visible = False Then Exit Sub
    'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
    Image2.Visible = False
    Image1.Visible = True
    Timer1.Enabled = True
End Sub

Private Sub Image3_Click()
SSave2 = "Select * From B001_HIS2"
Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
RSave2.AddNew
    RSave2("ID") = 0
    RSave2("NOPOL") = 0
    RSave2("TANGGAL_HIS") = Date
    RSave2("JAM_HIS") = Time
    RSave2("KETERANGAN_HIS") = "SOS LOCKER 2"
RSave2.Update
RSave2.Close
Set RSave2 = Nothing

If Image2.Visible = False Or Image6.Visible = False Then Exit Sub
    'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
    Image3.Visible = False
    Image4.Visible = True
    Timer2.Enabled = True
End Sub

Private Sub Image6_Click()
SSave2 = "Select * From B001_HIS2"
Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
RSave2.AddNew
    RSave2("ID") = 0
    RSave2("NOPOL") = 0
    RSave2("TANGGAL_HIS") = Date
    RSave2("JAM_HIS") = Time
    RSave2("KETERANGAN_HIS") = "SOS LOCKER 3"
RSave2.Update
RSave2.Close
Set RSave2 = Nothing

If Image2.Visible = False Or Image3.Visible = False Then Exit Sub
    'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
    Image6.Visible = False
    Image5.Visible = True
    Timer3.Enabled = True
End Sub

Private Sub Timer1_Timer()
'Port_Out 888, 128
    A = A + 100
    If A = 100 Then
        Shape7.Visible = True
    ElseIf A = 200 Then
        Shape1.Visible = True
    ElseIf A = 300 Then
        Shape2.Visible = True
    ElseIf A = 400 Then
        Shape3.Visible = True
    ElseIf A = 500 Then
        Shape4.Visible = True
    ElseIf A = 600 Then
        Shape5.Visible = True
    ElseIf A = 700 Then
        Shape6.Visible = True
    ElseIf A = 800 Then
        'Port_Out 888, 0
        A = 0
        Timer1.Enabled = False
            Shape1.Visible = False
            Shape2.Visible = False
            Shape3.Visible = False
            Shape4.Visible = False
            Shape5.Visible = False
            Shape6.Visible = False
            Shape7.Visible = False
        sndPlaySound App.Path & "\LOCKER.wav", SND_ASYNC
            Image2.Visible = True
            Image1.Visible = False
            Timer1.Enabled = False
        'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
    End If
End Sub

Private Sub Timer2_Timer()
'Port_Out 888, 64
    A = A + 100
    If A = 100 Then
        Shape7.Visible = True
    ElseIf A = 200 Then
        Shape1.Visible = True
    ElseIf A = 300 Then
        Shape2.Visible = True
    ElseIf A = 400 Then
        Shape3.Visible = True
    ElseIf A = 500 Then
        Shape4.Visible = True
    ElseIf A = 600 Then
        Shape5.Visible = True
    ElseIf A = 700 Then
        Shape6.Visible = True
    ElseIf A = 800 Then
        'Port_Out 888, 0
        A = 0
        Timer2.Enabled = False
            Shape1.Visible = False
            Shape2.Visible = False
            Shape3.Visible = False
            Shape4.Visible = False
            Shape5.Visible = False
            Shape6.Visible = False
            Shape7.Visible = False
        sndPlaySound App.Path & "\LOCKER.wav", SND_ASYNC
            Image3.Visible = True
            Image4.Visible = False
            Timer2.Enabled = False
        'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
    End If
End Sub

Private Sub Timer3_Timer()
'Port_Out 888, 32
    A = A + 100
    If A = 100 Then
        Shape7.Visible = True
    ElseIf A = 200 Then
        Shape1.Visible = True
    ElseIf A = 300 Then
        Shape2.Visible = True
    ElseIf A = 400 Then
        Shape3.Visible = True
    ElseIf A = 500 Then
        Shape4.Visible = True
    ElseIf A = 600 Then
        Shape5.Visible = True
    ElseIf A = 700 Then
        Shape6.Visible = True
    ElseIf A = 800 Then
        'Port_Out 888, 0
        A = 0
        Timer3.Enabled = False
            Shape1.Visible = False
            Shape2.Visible = False
            Shape3.Visible = False
            Shape4.Visible = False
            Shape5.Visible = False
            Shape6.Visible = False
            Shape7.Visible = False
        sndPlaySound App.Path & "\LOCKER.wav", SND_ASYNC
            Image6.Visible = True
            Image5.Visible = False
            Timer3.Enabled = False
        'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
    End If
End Sub
