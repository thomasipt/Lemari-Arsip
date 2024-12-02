VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form B001 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "B001.frx":0000
   ScaleHeight     =   599
   ScaleMode       =   0  'User
   ScaleWidth      =   672.842
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10406
      MaxLength       =   3
      TabIndex        =   20
      Text            =   "Text20"
      Top             =   5602
      Width           =   1080
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   2261
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2452
      Width           =   2655
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
      Left            =   2261
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   3127
      Width           =   9540
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
      Left            =   2261
      MaxLength       =   255
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   3817
      Width           =   9540
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2486
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   4792
      Width           =   2160
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2486
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   5197
      Width           =   2160
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2486
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   5602
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2486
      MaxLength       =   4
      TabIndex        =   6
      Text            =   "Text7"
      Top             =   6007
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2486
      TabIndex        =   7
      Text            =   "Text8"
      Top             =   6412
      Width           =   2160
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2486
      TabIndex        =   9
      Text            =   "Text10"
      Top             =   7222
      Width           =   2160
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2486
      TabIndex        =   10
      Text            =   "Text11"
      Top             =   7627
      Width           =   2160
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2486
      TabIndex        =   11
      Text            =   "Text12"
      Top             =   8032
      Width           =   2160
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6896
      TabIndex        =   12
      Text            =   "Text13"
      Top             =   4792
      Width           =   2160
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6896
      TabIndex        =   13
      Text            =   "Text14"
      Top             =   5197
      Width           =   2160
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6896
      TabIndex        =   14
      Text            =   "Text15"
      Top             =   5602
      Width           =   2160
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6896
      TabIndex        =   15
      Text            =   "Text16"
      Top             =   6007
      Width           =   2160
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6896
      TabIndex        =   16
      Text            =   "Text17"
      Top             =   6412
      Width           =   2160
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2486
      TabIndex        =   8
      Text            =   "Text9"
      Top             =   6817
      Width           =   2160
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10406
      MaxLength       =   3
      TabIndex        =   18
      Text            =   "Text18"
      Top             =   4792
      Width           =   1080
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10406
      MaxLength       =   1
      TabIndex        =   19
      Text            =   "Text19"
      Top             =   5197
      Width           =   1080
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   6889
      TabIndex        =   17
      ToolTipText     =   "Klik untuk edit"
      Top             =   6742
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Format          =   60555265
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin XPControls.XPButton CmdMasuk 
      Height          =   480
      Left            =   6844
      TabIndex        =   21
      Top             =   1687
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   847
      Caption         =   "SIMPAN"
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
   Begin XPControls.XPButton CmdKeluar 
      Height          =   480
      Left            =   10084
      TabIndex        =   22
      Top             =   1687
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
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label25"
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
      Left            =   45
      TabIndex        =   47
      Top             =   8595
      Width           =   1155
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KONTROL OTOMATIS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4999
      TabIndex        =   46
      Top             =   7597
      Width           =   1620
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TGL BERLAKU"
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
      Left            =   4999
      TabIndex        =   45
      Top             =   6862
      Width           =   1110
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(yyyy)"
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
      Left            =   4069
      TabIndex        =   44
      Top             =   5842
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   6972
      Picture         =   "B001.frx":15F942
      Stretch         =   -1  'True
      Top             =   7267
      Width           =   1995
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO. URUT"
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
      Left            =   9319
      TabIndex        =   43
      Top             =   5647
      Width           =   780
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO. RACK"
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
      Left            =   9319
      TabIndex        =   42
      Top             =   5242
      Width           =   765
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO. LEMARI"
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
      Left            =   9319
      TabIndex        =   41
      Top             =   4837
      Width           =   960
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
      Left            =   184
      TabIndex        =   40
      Top             =   2512
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA PEMILIK"
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
      Left            =   184
      TabIndex        =   39
      Top             =   3187
      Width           =   1860
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT"
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
      Left            =   184
      TabIndex        =   38
      Top             =   3877
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MERK / TYPE"
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
      Left            =   724
      TabIndex        =   37
      Top             =   4837
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JENIS / MODEL"
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
      Left            =   724
      TabIndex        =   36
      Top             =   5242
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAHUN PEMBUATAN"
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
      Left            =   724
      TabIndex        =   35
      Top             =   5647
      Width           =   1635
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAHUN PERAKITAN"
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
      Left            =   724
      TabIndex        =   34
      Top             =   6052
      Width           =   1560
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISI SILINDER"
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
      Left            =   724
      TabIndex        =   33
      Top             =   6457
      Width           =   1080
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WARNA"
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
      Left            =   724
      TabIndex        =   32
      Top             =   6862
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO RANGKA / NK"
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
      Left            =   724
      TabIndex        =   31
      Top             =   7267
      Width           =   1350
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO MESIN"
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
      Left            =   724
      TabIndex        =   30
      Top             =   7672
      Width           =   795
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO BPKB"
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
      Left            =   724
      TabIndex        =   29
      Top             =   8077
      Width           =   690
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WARNA TNKB"
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
      Left            =   4999
      TabIndex        =   28
      Top             =   4837
      Width           =   1095
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BAHAN BAKAR"
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
      Left            =   4999
      TabIndex        =   27
      Top             =   5242
      Width           =   1185
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KODE LOKASI"
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
      Left            =   4999
      TabIndex        =   26
      Top             =   5647
      Width           =   1095
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH BERAT YANG DIPERBOLEHKAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4999
      TabIndex        =   25
      Top             =   5947
      Width           =   2190
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO. URUT PENDAFT"
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
      Left            =   4999
      TabIndex        =   24
      Top             =   6457
      Width           =   1560
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Data Arsip"
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
      Left            =   6844
      TabIndex        =   23
      Top             =   667
      Width           =   4950
   End
   Begin VB.Image Image2 
      Height          =   1050
      Left            =   6972
      Picture         =   "B001.frx":166EB6
      Stretch         =   -1  'True
      Top             =   7267
      Width           =   1995
   End
End
Attribute VB_Name = "B001"
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

Private Sub cmdMASUK_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
Dim Tanya

If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" _
        Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" _
        Or Text11 = "" Or Text12 = "" Or Text13 = "" Or Text14 = "" Or Text15 = "" _
        Or Text16 = "" Or Text17 = "" Or Text18 = "" Or Text19 = "" Or Text20 = "" _
    Then
        MsgBox "DATA TIDAK BOLEH KOSONG", vbCritical, "KONFIRMASI"
        Text1.SetFocus
        Exit Sub
End If

Tanya = MsgBox("SIMPAN DATA", vbOKCancel, "KONFIRMASI")
    If Tanya = vbOK Then
        
        'Screen.MousePointer = vbHourglass
        'FRM_PROSES.Show
    
        SSave = "Select * From B001"
        Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
        RSave.AddNew
            RSave("ID") = NO_ID
            RSave("NOPOL") = Trim(Text1)
            RSave("NAMA") = Trim(Text2)
            RSave("ALAMAT") = Trim(Text3)
            RSave("MERK") = Trim(Text4)
            RSave("JENIS") = Trim(Text5)
            RSave("TH_PMBT") = Text6
            RSave("TH_PRKT") = Text7
            RSave("HP") = Text8
            RSave("WARNA") = Trim(Text9)
            RSave("NO_RANGKA") = Trim(Text10)
            RSave("NO_MESIN") = Trim(Text11)
            RSave("NO_BPKB") = Trim(Text12)
            RSave("WARNA_TNKB") = Trim(Text13)
            RSave("BBM") = Trim(Text14)
            RSave("KODE_LOKASI") = Trim(Text15)
            RSave("JML_BRT") = Trim(Text16)
            RSave("NO_URUT") = Trim(Text17)
            RSave("TGL_BRLK") = DTPicker1
            RSave("NO_LMR") = Trim(Text18)
            RSave("NO_LKR") = Trim(Text19)
            RSave("NO_URT") = Trim(Text20)
            
            If Text19 = 1 Then
                RSave("PORT") = 128
            ElseIf Text19 = 2 Then
                RSave("PORT") = 64
            ElseIf Text19 = 3 Then
                RSave("PORT") = 32
            End If
            
        RSave.Update
        RSave.Close
        Set RSave = Nothing
        
            SSave2 = "Select * From B001_HIS"
            Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
            RSave2.AddNew
                RSave2("ID") = NO_ID
                RSave2("NOPOL") = Trim(Text1)
                RSave2("NAMA") = Trim(Text2)
                RSave2("ALAMAT") = Trim(Text3)
                RSave2("MERK") = Trim(Text4)
                RSave2("JENIS") = Trim(Text5)
                RSave2("TH_PMBT") = Text6
                RSave2("TH_PRKT") = Text7
                RSave2("HP") = Text8
                RSave2("WARNA") = Trim(Text9)
                RSave2("NO_RANGKA") = Trim(Text10)
                RSave2("NO_MESIN") = Trim(Text11)
                RSave2("NO_BPKB") = Trim(Text12)
                RSave2("WARNA_TNKB") = Trim(Text13)
                RSave2("BBM") = Trim(Text14)
                RSave2("KODE_LOKASI") = Trim(Text15)
                RSave2("JML_BRT") = Trim(Text16)
                RSave2("NO_URUT") = Trim(Text17)
                RSave2("TGL_BRLK") = DTPicker1
                RSave2("NO_LMR") = Trim(Text18)
                RSave2("NO_LKR") = Trim(Text19)
                RSave2("NO_URT") = Trim(Text20)
                
                RSave2("TANGGAL_HIS") = Date
                RSave2("JAM_HIS") = Time
                RSave2("KETERANGAN_HIS") = "ENTRI"
                
                    If Text19 = 1 Then
                        RSave2("PORT") = 128
                    ElseIf Text19 = 1 Then
                        RSave2("PORT") = 64
                    ElseIf Text19 = 1 Then
                        RSave2("PORT") = 32
                    End If
            
            RSave2.Update
            RSave2.Close
            Set RSave2 = Nothing
        
                SSave3 = "Select * From A001"
                Set RSave3 = RDCO.OpenResultset(SSave3, rdOpenDynamic, rdConcurRowVer)
                RSave3.Edit
                    RSave3("ID") = NO_ID
                RSave3.Update
                RSave3.Close
                Set RSave3 = Nothing
            
        If Text18 = Max_Lemari And Text19 = Max_Locker And Text20 = Max_Urutan Then
            Call Miyabi_selesai
        End If
        
        Unload Me
        B001.Show
    Else
        Exit Sub
    End If
    
Unload FRM_PROSES
Set FRM_PROSES = Nothing
'Screen.MousePointer = vbDefault

End Sub

Private Sub Miyabi_selesai()
SCari = "Select * From C013"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    RCari.Edit
        RCari("STS_LEMARI") = "1"
    RCari.Update
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=LOCKER", rdDriverNoPrompt, False, CN)
ClearTextBoxes Me
Image2.Visible = False

Call Cek_Lemari

Call Hot_Miyabi
    Text18.Enabled = False
    Text19.Enabled = False
    Text20.Enabled = False
    
DTPicker1 = Date

Call IDENTITAS

SCari = "SELECT Count(B001.NOPOL) AS CountOfNOPOL FROM B001"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Label25 = Max_Total - RCari("CountOfNOPOL")
End If
RCari.Close
Set RCari = Nothing

End Sub

Private Sub Cek_Lemari()
SCari = "Select * From C013"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
If RCari("STS_LEMARI") = 1 And RCari("STS_JML") = 1 Then
    MsgBox "PENOMORAN TELAH MENCAPAI BATAS MAKSIMAL", vbCritical, "WARNING"
    Exit Sub
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub IDENTITAS()
SCari = "Select * From A001 Order By NO Asc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    NO_ID = RCari("ID") + 1
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Hot_Miyabi()
Dim Lemari, Locker, Urutan As Double

SCari = "Select * From B001 Order By NO Asc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
If RCari.RowCount = 0 Then
    Text18 = 1
    Text19 = 1
    Text20 = 1
Else
    Lemari = RCari("NO_LMR")
    Locker = RCari("NO_LKR")
    Urutan = RCari("NO_URT")
    
    If Urutan < Max_Urutan Then
        Text20 = Urutan + 1
        Text19 = Locker
        Text18 = Lemari
    ElseIf Urutan = Max_Urutan Then
        Text20 = 1
        If Locker < Max_Locker Then
            Text19 = Locker + 1
            Text18 = Lemari
        ElseIf Locker = Max_Locker Then
            Text19 = 1
            If Lemari < Max_Lemari Then
                Text18 = Lemari + 1
            ElseIf Lemari = Max_Lemari Then
                Text18 = 1
            End If
        End If
    End If
   
End If
RCari.Close
Set RCari = Nothing
        
End Sub

Private Sub Image1_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
Image1.Visible = False
Image2.Visible = True
    Text18.Enabled = True
    Text19.Enabled = True
    Text20.Enabled = True
End Sub

Private Sub Image2_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
Image1.Visible = True
Image2.Visible = False
    Text18.Enabled = False
    Text19.Enabled = False
    Text20.Enabled = False
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

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
Text3 = Format(Text3, ">")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text4_LostFocus()
Text4 = Format(Text4, ">")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
Text5 = Format(Text5, ">")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text6_LostFocus()
If Text6 = "" Then Exit Sub
    If Len(Text6) < 4 Then
        MsgBox "BUKAN FORMAT TAHUN", vbSystemModal, "KONFIRMASI"
        Text6 = ""
        Text6.SetFocus
    End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text7_LostFocus()
If Text7 = "" Then Exit Sub
    If Len(Text7) < 4 Then
        MsgBox "BUKAN FORMAT TAHUN", vbSystemModal, "KONFIRMASI"
        Text7 = ""
        Text7.SetFocus
    End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text9_LostFocus()
Text9 = Format(Text9, ">")
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text10_LostFocus()
Text10 = Format(Text10, ">")
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text11_LostFocus()
Text11 = Format(Text11, ">")
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text12_LostFocus()
Text12 = Format(Text12, ">")
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text13_LostFocus()
Text13 = Format(Text13, ">")
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text14_LostFocus()
Text14 = Format(Text14, ">")
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text15_LostFocus()
Text15 = Format(Text15, ">")
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text16_LostFocus()
Text16 = Format(Text16, ">")
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text17_LostFocus()
Text17 = Format(Text17, ">")
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
