VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form LOGIN 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5940
      _Version        =   65536
      _ExtentX        =   10477
      _ExtentY        =   1058
      _StockProps     =   15
      Caption         =   "LOGIN USER"
      ForeColor       =   0
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   17.99
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Outline         =   -1  'True
      Font3D          =   4
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   840
      Left            =   0
      TabIndex        =   5
      Top             =   2145
      Width           =   5940
      _Version        =   65536
      _ExtentX        =   10477
      _ExtentY        =   1482
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      ShadowStyle     =   1
      Begin XPControls.XPButton XPButton1 
         Height          =   600
         Left            =   450
         TabIndex        =   2
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1058
         Caption         =   "MASUK"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SoundOnClick    =   1
         ColorScheme     =   1
         ColorBegin      =   16777215
         ColorEnd        =   8421504
         HoverEffect     =   -1  'True
      End
      Begin XPControls.XPButton XPButton2 
         Height          =   600
         Left            =   4155
         TabIndex        =   3
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1058
         Caption         =   "KELUAR"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SoundOnClick    =   1
         ColorScheme     =   1
         ColorBegin      =   16777215
         ColorEnd        =   8421504
         HoverEffect     =   -1  'True
      End
   End
   Begin Threed.SSPanel Pelanggan 
      Height          =   1395
      Left            =   90
      TabIndex        =   6
      Top             =   675
      Width           =   4260
      _Version        =   65536
      _ExtentX        =   7514
      _ExtentY        =   2461
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1650
         MaxLength       =   25
         TabIndex        =   0
         Top             =   300
         Width           =   2160
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1650
         MaxLength       =   25
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   705
         Width           =   2160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER"
         Height          =   195
         Left            =   450
         TabIndex        =   8
         Top             =   345
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         Height          =   195
         Left            =   450
         TabIndex        =   7
         Top             =   705
         Width           =   945
      End
   End
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   4410
      Picture         =   "LOGIN.frx":0000
      Stretch         =   -1  'True
      Top             =   675
      Width           =   1455
   End
End
Attribute VB_Name = "LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private SqlPass As String
Private tUser As rdoResultset
Private tMasuk As rdoResultset

Private RTgl, RHapus, RDEl, RSave2, RSave3, RSave4, RCari, RCari2, RSLNO, rscs3 As rdoResultset
Private STgl, SHapus, SDel, SSave2, SSave3, SSave4, SCari, SCari2, SqlNo, sqlcs3, Kode As String

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=LOCKER", rdDriverNoPrompt, False, CN)
Text1 = ""
Text2 = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub

Private Sub XPButton1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call XPButton1_Click
    End If
End Sub

Private Sub XPButton1_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
'Screen.MousePointer = vbHourglass
'FRM_PROSES.Show
SCari = "Select * From C013 where NAMA = '" + Text1 + "' and PASSWORD = '" + Text2 + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    SqlPass = "Select * from C013 where NAMA =  '" + Trim(Text1) + "' "
    Set tMasuk = RDCO.OpenResultset(SqlPass, rdOpenDynamic, rdConcurRowVer)
    If tMasuk.RowCount <> 0 Then
        Max_Locker = tMasuk("MAX_LKR")
        Max_Lemari = tMasuk("MAX_LMR")
        Max_Urutan = tMasuk("MAX_URT")
        Max_Total = tMasuk("STS_TOTAL")
        MAINMENU.Show
    End If
    tMasuk.Close
    Set tMasuk = Nothing
    Unload Me
Else
    LOGIN.Hide
    MsgBox "ANDA TIDAK BERHAK LOG IN KE SYSTEM", vbCritical, "KONFIRMASI"
    LOGIN.Show
    Text1 = ""
    Text2 = ""
    Text1.SetFocus
End If
RCari.Close
Set RCari = Nothing

Unload FRM_PROSES
Set FRM_PROSES = Nothing
'Screen.MousePointer = vbDefault
  
End Sub

Private Sub XPButton2_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
End
End Sub
