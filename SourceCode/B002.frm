VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form B002 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "B002.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":15F942
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   27
      Top             =   6840
      Width           =   420
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":166EB4
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   26
      Top             =   6210
      Width           =   420
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":16E426
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   25
      Top             =   5580
      Width           =   420
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":175998
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   24
      Top             =   4950
      Width           =   420
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":17CF0A
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   23
      Top             =   4320
      Width           =   420
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":18447C
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   22
      Top             =   3690
      Width           =   420
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   2400
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   45
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2505
      TabIndex        =   6
      Text            =   "7"
      Top             =   6840
      Width           =   2940
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2505
      TabIndex        =   5
      Text            =   "6"
      Top             =   6210
      Width           =   2940
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2505
      TabIndex        =   4
      Text            =   "5"
      Top             =   5580
      Width           =   2940
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2505
      TabIndex        =   3
      Text            =   "4"
      Top             =   4950
      Width           =   2940
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2505
      TabIndex        =   2
      Text            =   "3"
      Top             =   4320
      Width           =   2940
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2505
      TabIndex        =   1
      Text            =   "2"
      Top             =   3690
      Width           =   2940
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2505
      TabIndex        =   0
      Text            =   "1"
      Top             =   3060
      Width           =   2940
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   10350
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   28
      Top             =   4095
      Visible         =   0   'False
      Width           =   420
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   8580
      TabIndex        =   7
      ToolTipText     =   "Klik untuk edit"
      Top             =   3765
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   49152
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   16777088
      Format          =   16515073
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   420
      Left            =   8580
      TabIndex        =   8
      ToolTipText     =   "Klik untuk edit"
      Top             =   4395
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   49152
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   16777088
      Format          =   16515073
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin XPControls.XPButton SSCommand1 
      Height          =   480
      Left            =   6195
      TabIndex        =   10
      Top             =   6802
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   847
      Caption         =   "LOCKER"
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
   Begin XPControls.XPButton cmdKELUAR 
      Height          =   480
      Left            =   7245
      TabIndex        =   11
      Top             =   1665
      Width           =   2520
      _ExtentX        =   4445
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
   Begin XPControls.XPButton SSCommand2 
      Height          =   480
      Left            =   6195
      TabIndex        =   9
      Top             =   5542
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   847
      Caption         =   "CARI"
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
   Begin XPControls.XPButton XPButton1 
      Height          =   480
      Left            =   6195
      TabIndex        =   31
      Top             =   6172
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   847
      Caption         =   "TAMPILKAN PENCARIAN"
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
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   780
      Left            =   2505
      TabIndex        =   32
      ToolTipText     =   "Klik untuk memilih"
      Top             =   7470
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   1376
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16777088
      BackColorBkg    =   8421376
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":18B9EE
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   21
      Top             =   3060
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   270
      TabIndex        =   30
      Top             =   8415
      Width           =   1170
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informasi ______________________________________________"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   270
      TabIndex        =   29
      Top             =   8010
      Width           =   9765
   End
   Begin VB.Image Image8 
      Height          =   405
      Left            =   10350
      Stretch         =   -1  'True
      Top             =   4095
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image13 
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":192F62
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   420
   End
   Begin VB.Image Image11 
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":19A4D4
      Stretch         =   -1  'True
      Top             =   6210
      Width           =   420
   End
   Begin VB.Image Image9 
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":1A1A46
      Stretch         =   -1  'True
      Top             =   5580
      Width           =   420
   End
   Begin VB.Image Image7 
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":1A8FB8
      Stretch         =   -1  'True
      Top             =   4950
      Width           =   420
   End
   Begin VB.Image Image5 
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":1B052A
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":1B7A9C
      Stretch         =   -1  'True
      Top             =   3690
      Width           =   420
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TGL. BERLAKU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6690
      TabIndex        =   20
      Top             =   4155
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO. URUT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   570
      TabIndex        =   19
      Top             =   6900
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO. RACK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   570
      TabIndex        =   18
      Top             =   6270
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO. LEMARI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   570
      TabIndex        =   17
      Top             =   5640
      Width           =   1470
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KODE LOKASI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   570
      TabIndex        =   16
      Top             =   5010
      Width           =   1665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JENIS / MODEL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   570
      TabIndex        =   15
      Top             =   4380
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MERK / TYPE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   570
      TabIndex        =   14
      Top             =   3750
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMOR POLISI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   570
      TabIndex        =   13
      Top             =   3120
      Width           =   1845
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pencarian Data Arsip"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   36
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   825
      Left            =   5490
      TabIndex        =   12
      Top             =   585
      Width           =   6300
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   5490
      Picture         =   "B002.frx":1BF00E
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   420
   End
End
Attribute VB_Name = "B002"
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

Call SCAN_TOKET
Unload Me
MAINMENU.Show

Unload FRM_PROSES
Set FRM_PROSES = Nothing
'Screen.MousePointer = vbDefault
End Sub

Private Sub SCAN_TOKET()
SCari2 = "Select * From B001 where STS = '1'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
   RCari2.MoveFirst
   Do Until RCari2.EOF
        RCari2.Edit
            RCari2("STS") = "0"
        RCari2.Update
   RCari2.MoveNext
   Loop
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=LOCKER", rdDriverNoPrompt, False, CN)
ClearTextBoxes Me
Call ClearText

DTPicker1 = Date
DTPicker2 = Date

Call Mati
N_FORM2 = 0

Timer1.Enabled = False

Call SiapkanGrid
End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 4
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "NOPOL": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1500: .Text = "LEMARI": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1500: .Text = "RACK": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1500: .Text = "URUT": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
Call SiapkanGrid

SCari2 = "Select * From B001 where STS = '1' "
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
   RCari2.MoveFirst
   B = 1
   Do Until RCari2.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RCari2("NOPOL")
              .Col = 1: .Text = RCari2("NO_LMR")
              .Col = 2: .Text = RCari2("NO_LKR")
              .Col = 3: .Text = RCari2("NO_URT")
         End With
      B = B + 1
      RCari2.MoveNext
   Loop
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub Mati()
A = 0
Image1.Visible = False
Image3.Visible = False
Image5.Visible = False
Image7.Visible = False
Image9.Visible = False
Image11.Visible = False
Image13.Visible = False
Text1.Enabled = False
    Text1.BackColor = &H0&
Text2.Enabled = False
    Text2.BackColor = &H0&
Text2.Enabled = False
    Text2.BackColor = &H0&
Text3.Enabled = False
    Text3.BackColor = &H0&
Text4.Enabled = False
    Text4.BackColor = &H0&
Text5.Enabled = False
    Text5.BackColor = &H0&
Text6.Enabled = False
    Text6.BackColor = &H0&
Text7.Enabled = False
    Text7.BackColor = &H0&
DTPicker1.Enabled = False
DTPicker2.Enabled = False
Label10 = ""
XPButton1.Visible = False
SSCommand1.Visible = False
End Sub

Private Sub ClearText()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
End Sub

Private Sub Image1_Click()
SCari = "Select * from B001 where NOPOL = '" + Trim(Text1) + "'"
Call GANTI_TOKET2
Call JUMLAH_TOKET
Picture1.Visible = True
Image1.Visible = False
    Text1.Enabled = False
    Text1.BackColor = &H0&
    Text1 = ""
End Sub

Private Sub Picture1_Click()
Picture1.Visible = False
Image1.Visible = True
    Text1.Enabled = True
    Text1.BackColor = &HFFFFFF
    Text1.SetFocus
N_FORM = ""
N_FORM = "NOMOR POLISI"
B002_NOPOL.Show 1
Text1 = N_FORM2

If Text1 = "" Then
    Call Image1_Click
End If

End Sub

Private Sub Image3_Click()
SCari = "Select * from B001 where MERK like '%" + Trim(Text2) + "%'"
Call GANTI_TOKET2
Call JUMLAH_TOKET
Picture2.Visible = True
Image3.Visible = False
    Text2.Enabled = False
    Text2.BackColor = &H0&
    Text2 = ""
End Sub

Private Sub Picture2_Click()
Picture2.Visible = False
Image3.Visible = True
    Text2.Enabled = True
    Text2.BackColor = &HFFFFFF
    Text2.SetFocus
N_FORM = ""
N_FORM = "MERK"
B002_NOPOL.Show 1
Text2 = N_FORM2

If Text2 = "" Then
    Call Image3_Click
End If

End Sub

Private Sub Image5_Click()
SCari = "Select * from B001 where JENIS like '%" + Trim(Text3) + "%'"
Call GANTI_TOKET2
Call JUMLAH_TOKET
Picture3.Visible = True
Image5.Visible = False
    Text3.Enabled = False
    Text3.BackColor = &H0&
    Text3 = ""
End Sub

Private Sub Picture3_Click()
Picture3.Visible = False
Image5.Visible = True
    Text3.Enabled = True
    Text3.BackColor = &HFFFFFF
N_FORM = ""
N_FORM = "JENIS"
B002_NOPOL.Show 1
Text3 = N_FORM2

If Text3 = "" Then
    Call Image5_Click
End If

End Sub

Private Sub Image7_Click()
SCari = "Select * from B001 where KODE_LOKASI like '%" + Trim(Text4) + "%'"
Call GANTI_TOKET2
Call JUMLAH_TOKET
Picture4.Visible = True
Image7.Visible = False
    Text4.Enabled = False
    Text4.BackColor = &H0&
    Text4 = ""
End Sub

Private Sub Picture4_Click()
Picture4.Visible = False
Image7.Visible = True
    Text4.Enabled = True
    Text4.BackColor = &HFFFFFF
N_FORM = ""
N_FORM = "KODE LOKASI"
B002_NOPOL.Show 1
Text4 = N_FORM2

If Text4 = "" Then
    Call Image7_Click
End If

End Sub

Private Sub Image9_Click()
SCari = "Select * from B001 where NO_LMR like '%" + Trim(Text5) + "%'"
Call GANTI_TOKET2
Call JUMLAH_TOKET
Picture5.Visible = True
Image9.Visible = False
    Text5.Enabled = False
    Text5.BackColor = &H0&
    Text5 = ""
End Sub

Private Sub Picture5_Click()
Picture5.Visible = False
Image9.Visible = True
    Text5.Enabled = True
    Text5.BackColor = &HFFFFFF
N_FORM = ""
N_FORM = "LEMARI"
B002_NOPOL.Show 1
Text5 = N_FORM2

If Text5 = "" Then
    Call Image9_Click
End If

End Sub

Private Sub Image11_Click()
SCari = "Select * from B001 where NO_LKR like '%" + Trim(Text6) + "%'"
Call GANTI_TOKET2
Call JUMLAH_TOKET
Picture6.Visible = True
Image11.Visible = False
    Text6.Enabled = False
    Text6.BackColor = &H0&
    Text6 = ""
End Sub

Private Sub Picture6_Click()
Picture6.Visible = False
Image11.Visible = True
    Text6.Enabled = True
    Text6.BackColor = &HFFFFFF
N_FORM = ""
N_FORM = "LOKER"
B002_NOPOL.Show 1
Text6 = N_FORM2

If Text6 = "" Then
    Call Image11_Click
End If

End Sub

Private Sub Image13_Click()
SCari = "Select * from B001 where NO_URT like '%" + Trim(Text7) + "%'"
Call GANTI_TOKET2
Call JUMLAH_TOKET
Picture7.Visible = True
Image13.Visible = False
    Text7.Enabled = False
    Text7.BackColor = &H0&
    Text7 = ""
End Sub

Private Sub Picture7_Click()
Picture7.Visible = False
Image13.Visible = True
    Text7.Enabled = True
    Text7.BackColor = &HFFFFFF
N_FORM = ""
N_FORM = "BARIS"
B002_NOPOL.Show 1
Text7 = N_FORM2

If Text7 = "" Then
    Call Image13_Click
End If

End Sub

Private Sub Image8_Click()
Picture8.Visible = True
Image8.Visible = False
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
End Sub

Private Sub Picture8_Click()
Picture8.Visible = False
Image8.Visible = True
    DTPicker1.Enabled = True
    DTPicker2.Enabled = True
End Sub

Private Sub SSCommand1_Click()
SCari = "Select * From B001 where STS='1'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
    NO_PORT = RCari("PORT")
    
        SSave2 = "Select * From B001_HIS2"
        Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
        RSave2.AddNew
            RSave2("ID") = RCari("ID")
            RSave2("NOPOL") = RCari("NOPOL")
            RSave2("TANGGAL_HIS") = Date
            RSave2("JAM_HIS") = Time
            RSave2("KETERANGAN_HIS") = "OPEN"
        RSave2.Update
        RSave2.Close
        Set RSave2 = Nothing
    
RCari.Close
Set RCari = Nothing

'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
If N_FORM2 = 1 Then
    'Port_Out 888, NO_PORT
    sndPlaySound App.Path & "\LOCKER.wav", SND_ASYNC
    Timer1.Enabled = True
    'Call B001_HIS
ElseIf N_FORM2 > 1 Then
    MsgBox "HASIL PENCARIAN LEBIH DARI SATU ARSIP", vbCritical, "WARNING"
ElseIf N_FORM2 = 0 Then
    MsgBox "TIDAK ADA ARSIP YANG DIPILIH", vbCritical, "WARNING"
End If
End Sub

Private Sub Cari2()
'SCari = "Select * From B001P004 where NAMA like '%" + Trim(Text1) + "%' or ALAMAT like '%" + Trim(Text1) + "%' or KELUH like '%" + Trim(Text1) + "%' or PERIKSA like '%" + Trim(Text1) + "%'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
   Do Until RCari.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
            .Col = 0: .Text = RCari("NO_PASIEN"): .CellAlignment = 4
            .Col = 1: .Text = RCari("NAMA"): .CellAlignment = 4
            .Col = 2: .Text = RCari("ALAMAT"): .CellAlignment = 4
            .Col = 3: .Text = RCari("TGL_KUNJ"): .CellAlignment = 4
            .Col = 4: .Text = Format(RCari("BIAYA"), "##,###.00")
            .Col = 5: .Text = RCari("KELUH"): .CellAlignment = 4
            .Col = 6: .Text = RCari("PERIKSA"): .CellAlignment = 4
            .Col = 7: .Text = RCari("DIAG"): .CellAlignment = 4
            .Col = 8: .Text = RCari("TINDAKAN"): .CellAlignment = 4
            .Col = 9: .Text = RCari("OBAT"): .CellAlignment = 4
         End With
      B = B + 1
      RCari.MoveNext
   Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub SSCommand2_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
'Screen.MousePointer = vbHourglass
'FRM_PROSES.Show

If Text1 = "" And Text2 = "" And Text3 = "" And Text4 = "" And Text5 = "" And Text6 = "" And Text7 = "" Then
    MsgBox "OPTION TIDAK AKTIF", vbCritical, "KONFIRMASI"
    XPButton1.Visible = False
    SSCommand1.Visible = False
        Unload FRM_PROSES
        Set FRM_PROSES = Nothing
        'Screen.MousePointer = vbDefault
    Exit Sub
End If

If Picture1.Visible = False Then
    SCari = "Select * from B001 where NOPOL = '" + Trim(Text1) + "'"
    Call GANTI_TOKET
End If

    If Picture2.Visible = False Then
        SCari = "Select * from B001 where MERK like '%" + Trim(Text2) + "%'"
        Call GANTI_TOKET
    End If

If Picture3.Visible = False Then
    SCari = "Select * from B001 where JENIS like '%" + Trim(Text3) + "%'"
    Call GANTI_TOKET
End If

    If Picture4.Visible = False Then
        SCari = "Select * from B001 where KODE_LOKASI like '%" + Trim(Text4) + "%'"
        Call GANTI_TOKET
    End If

If Picture5.Visible = False Then
    SCari = "Select * from B001 where NO_LMR like '%" + Trim(Text5) + "%'"
    Call GANTI_TOKET
End If

    If Picture6.Visible = False Then
        SCari = "Select * from B001 where NO_LKR like '%" + Trim(Text6) + "%'"
        Call GANTI_TOKET
    End If

If Picture7.Visible = False Then
    SCari = "Select * from B001 where NO_URT like '%" + Trim(Text7) + "%'"
    Call GANTI_TOKET
End If

Call JUMLAH_TOKET
Call REMES_TOKET

Unload FRM_PROSES
Set FRM_PROSES = Nothing
'Screen.MousePointer = vbDefault

End Sub

Private Sub REMES_TOKET()
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.FontName = "Courier New"

Printer.FontSize = 8.5
Printer.Print Tab(5); ""
Printer.Print Tab(5); ""
Printer.Print Tab(5); ""

Printer.FontSize = 18
Printer.FontBold = True
Printer.Print Tab(1); "FILLING CABINET CONTROL"
Printer.FontBold = False

Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Tab(1); "ADI JAYASARANA"
Printer.FontBold = False

Printer.FontSize = 8.5
Printer.FontBold = True
Printer.Print Tab(2); "TGL. "; Now
Printer.FontBold = False

Printer.FontSize = 8.5
Printer.Print Tab(1); "|-------------------------------------|"
Printer.Print Tab(1); "|  NOPOL    = "; RKanan(grid.TextMatrix(1, 0)); "|"
Printer.Print Tab(1); "|-------------------------------------|"
Printer.Print Tab(1); "|  LEMARI   = "; RKanan(grid.TextMatrix(1, 1)); "|"
Printer.Print Tab(1); "|-------------------------------------|"
Printer.Print Tab(1); "|  RACK     = "; RKanan(grid.TextMatrix(1, 2)); "|"
Printer.Print Tab(1); "|-------------------------------------|"
Printer.Print Tab(1); "|  URUT     = "; RKanan(grid.TextMatrix(1, 3)); "|"
Printer.Print Tab(1); "|-------------------------------------|"
Printer.Print Tab(1); ""

Printer.FontSize = 14
Printer.FontBold = True
Printer.Print Tab(1); "     TERIMA KASIH     "
Printer.FontBold = False

Printer.FontSize = 18
Printer.Print Tab(1); ""
Printer.Print Tab(1); ""

Printer.EndDoc
End Sub

Private Sub GANTI_TOKET()
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Do Until RCari.EOF
        RCari.Edit
            RCari("STS") = "1"
        RCari.Update
        RCari.MoveNext
    Loop
    Call IsiGrid
ElseIf RCari.RowCount = 0 Then
    Call SCAN_TOKET
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub GANTI_TOKET2()
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Do Until RCari.EOF
        RCari.Edit
            RCari("STS") = "0"
        RCari.Update
        RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub JUMLAH_TOKET()
N_FORM2 = 0
SCari = "SELECT Count(B001.NOPOL) AS CountOfNOPOL FROM B001 where STS = '1'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    N_FORM2 = RCari("CountOfNOPOL")
    Label10 = "DITEMUKAN " + N_FORM2 + " RECORD"
End If
RCari.Close
Set RCari = Nothing

If N_FORM2 = 0 Then
    XPButton1.Visible = False
    SSCommand1.Visible = False
Else
    SSCommand1.Visible = True
    XPButton1.Visible = True
End If

End Sub

Private Sub XPButton1_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
CRPT.ReportFileName = App.Path & "\Report\B001.rpt"
CRPT.SelectionFormula = "{B001.STS} = '1'"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = False
CRPT.WindowMinButton = False
CRPT.Action = 1
End Sub

Private Sub Timer1_Timer()
A = A + 1
If A = 5 Then
    Timer1.Enabled = False
    'Port_Out 888, 0
        'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
        
        Call SCAN_TOKET
        Unload Me
        B002.Show
End If
End Sub
