VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form MAINMENU 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8595
   ClientLeft      =   -3780
   ClientTop       =   150
   ClientWidth     =   11925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
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
      Left            =   9855
      TabIndex        =   40
      Text            =   "2"
      Top             =   4140
      Width           =   1980
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
      Left            =   10965
      MaxLength       =   3
      TabIndex        =   36
      Text            =   "19"
      Top             =   3150
      Width           =   855
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
      Left            =   10965
      MaxLength       =   3
      TabIndex        =   35
      Text            =   "18"
      Top             =   2655
      Width           =   855
   End
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
      Left            =   10965
      MaxLength       =   3
      TabIndex        =   34
      Text            =   "20"
      Top             =   3645
      Width           =   855
   End
   Begin VB.ComboBox cboSearch 
      BackColor       =   &H00FFFFC0&
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
      Left            =   5775
      TabIndex        =   0
      Top             =   8025
      Width           =   3510
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   25
      Left            =   90
      Max             =   255
      Min             =   100
      SmallChange     =   25
      TabIndex        =   18
      Top             =   8310
      Value           =   100
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   45
      Top             =   8820
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3195
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "MAINMENU.frx":0000
      Top             =   9270
      Width           =   6345
   End
   Begin VB.PictureBox TOOLS1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   -645
      Picture         =   "MAINMENU.frx":0006
      ScaleHeight     =   945
      ScaleWidth      =   3540
      TabIndex        =   10
      Top             =   3855
      Width           =   3540
   End
   Begin VB.PictureBox LAPORAN1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   -645
      Picture         =   "MAINMENU.frx":B40C
      ScaleHeight     =   945
      ScaleWidth      =   3540
      TabIndex        =   11
      Top             =   5145
      Width           =   3540
   End
   Begin VB.PictureBox EXIT1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   9750
      Picture         =   "MAINMENU.frx":16812
      ScaleHeight     =   960
      ScaleWidth      =   2670
      TabIndex        =   8
      Top             =   5130
      Width           =   2670
   End
   Begin VB.PictureBox DATA1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   -645
      Picture         =   "MAINMENU.frx":1F88C
      ScaleHeight     =   945
      ScaleWidth      =   3540
      TabIndex        =   9
      Top             =   2640
      Width           =   3540
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XPControls.XPButton XPButton1 
      Height          =   945
      Left            =   3330
      TabIndex        =   1
      Top             =   2640
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1667
      Caption         =   "ENTRI"
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
   Begin XPControls.XPButton XPButton2 
      Height          =   945
      Left            =   7815
      TabIndex        =   3
      Top             =   2640
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1667
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
   Begin XPControls.XPButton XPButton5 
      Height          =   945
      Left            =   3330
      TabIndex        =   4
      Top             =   3855
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1667
      Caption         =   "SOS"
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
      ColorEnd        =   16777088
   End
   Begin XPControls.XPButton XPButton9 
      Height          =   945
      Left            =   3330
      TabIndex        =   6
      Top             =   5145
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1667
      Caption         =   "ARSIP"
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
      ColorEnd        =   16777088
   End
   Begin XPControls.XPButton XPButton6 
      Height          =   945
      Left            =   5573
      TabIndex        =   5
      Top             =   3855
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1667
      Caption         =   "SISTEM"
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
      ColorEnd        =   16777088
   End
   Begin XPControls.XPButton XPButton10 
      Height          =   945
      Left            =   5573
      TabIndex        =   7
      Top             =   5145
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1667
      Caption         =   "AKTIFITAS"
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
      ColorEnd        =   16777088
   End
   Begin XPControls.XPButton XPButton3 
      Height          =   945
      Left            =   5573
      TabIndex        =   2
      Top             =   2640
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1667
      Caption         =   "EDIT"
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
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LEMARI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   9885
      TabIndex        =   39
      Top             =   2655
      Width           =   960
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RACK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   9885
      TabIndex        =   38
      Top             =   3150
      Width           =   690
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URUT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   9885
      TabIndex        =   37
      Top             =   3645
      Width           =   705
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3540
      TabIndex        =   33
      Top             =   8085
      Width           =   1845
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl (F7)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   5580
      TabIndex        =   32
      Top             =   6030
      Width           =   795
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl (F6)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   3330
      TabIndex        =   31
      Top             =   6030
      Width           =   795
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl (F5)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   5580
      TabIndex        =   30
      Top             =   4740
      Width           =   795
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl (F4)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   3330
      TabIndex        =   29
      Top             =   4740
      Width           =   795
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl (F3)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   7815
      TabIndex        =   28
      Top             =   3525
      Width           =   795
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl (F2)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   5580
      TabIndex        =   27
      Top             =   3525
      Width           =   795
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl (F1)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   3330
      TabIndex        =   26
      Top             =   3525
      Width           =   795
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   1860
      TabIndex        =   25
      Top             =   7395
      Width           =   630
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1695
      TabIndex        =   24
      Top             =   7950
      Width           =   975
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tersedia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   90
      TabIndex        =   23
      Top             =   7950
      Width           =   1065
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   1860
      TabIndex        =   22
      Top             =   7125
      Width           =   645
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record Maks"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   90
      TabIndex        =   21
      Top             =   7395
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Record"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   90
      TabIndex        =   20
      Top             =   7125
      Width           =   1410
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Utama"
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
      Left            =   8100
      TabIndex        =   19
      Top             =   495
      Width           =   3720
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   9675
      TabIndex        =   17
      Top             =   45
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   9225
      TabIndex        =   15
      Top             =   45
      Width           =   2220
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pencarian Cepat_____________________________"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3150
      TabIndex        =   14
      Top             =   7620
      Width           =   6360
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Adi Jaya Sarana"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   90
      TabIndex        =   13
      Top             =   6600
      Width           =   1860
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyrighted® EDP IPT 2008"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   90
      TabIndex        =   12
      Top             =   6285
      Width           =   3615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   8790
      Left            =   -45
      Picture         =   "MAINMENU.frx":2AC92
      Stretch         =   -1  'True
      Top             =   -165
      Width           =   12000
   End
End
Attribute VB_Name = "MAINMENU"
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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_FINDSTRING = &H14C
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private miSelStart As Integer
Private blnEdit As Boolean
Private blnGaAda As Boolean

Private strProvider As String

Private Sub DATA0_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
    'DATA0.Visible = False
    'DATA1.Visible = True
    'XPButton1.Visible = True
    'XPButton3.Visible = True
    ''XPButton4.Visible = True
    'XPButton2.Visible = True
    'XPButton1.SetFocus
    'TOOLS0.Visible = False
    'TOOLS1.Visible = False
    'LAPORAN0.Visible = False
    'LAPORAN1.Visible = False
    Text1 = " Pilih salah satu tombol untuk melakukan pendataan semua arsip pada rack cabinet "
End Sub

Private Sub DATA1_Click()
'DATA1.Visible = False
'DATA0.Visible = True
'TombolMati Me
    'TOOLS0.Visible = True
    'TOOLS1.Visible = True
    'LAPORAN0.Visible = True
    'LAPORAN1.Visible = True
    ClearTextBoxes Me
End Sub

Private Sub EXIT0_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
Dim Tanya

EXIT0.Visible = False
EXIT1.Visible = True

Tanya = MsgBox("KELUAR DARI SISTEM FILING CABINET CONTROL", vbQuestion + vbOKCancel, "KONFIRMASI")
If Tanya = vbCancel Then
    EXIT0.Visible = True
    EXIT1.Visible = False
Else
    End
End If
End Sub

Private Sub EXIT1_Click()
Dim Tanya

Tanya = MsgBox("KELUAR DARI SISTEM FILING CABINET CONTROL", vbQuestion + vbOKCancel, "KONFIRMASI")
If Tanya = vbOK Then
    End
End If

End Sub

Private Sub Form_Activate()
Call INFO_ARSIP
End Sub

Private Sub Form_Load()
'Port_Out 888, 0
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=LOCKER", rdDriverNoPrompt, False, CN)
'TombolMati Me
Call Ngumpet
ClearTextBoxes Me

Label1 = Date
Label2 = Time

HScroll1.Value = 255

Text18 = 0
Text19 = 0
Text20 = 0

Call INFO_ARSIP

Call CARI_NOPOL

End Sub

Private Sub CARI_NOPOL()
SCombo1 = "Select NOPOL from B001 Order By NOPOL Asc"
Set RCombo1 = RDCO.OpenResultset(SCombo1, rdOpenDynamic, rdConcurRowVer)
If RCombo1.RowCount <> 0 Then
    RCombo1.MoveFirst
    Do Until RCombo1.EOF
        cboSearch.AddItem RCombo1("NOPOL")
    RCombo1.MoveNext
    Loop
    RCombo1.Close
    Set RCombo1 = Nothing
    cboSearch = ""
End If
End Sub
Private Sub cboSearch_Change()
SCari = "Select * From B001 where NOPOL = '" + Trim(cboSearch) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
    Text18 = RCari("NO_LMR")
    Text19 = RCari("NO_LKR")
    Text20 = RCari("NO_URT")
    Text2 = RCari("PORT")
Else
    Text18 = 0
    Text19 = 0
    Text20 = 0
End If
RCari.Close
Set RCari = Nothing
    
   Dim i As Long, ii As Long
   Dim strBagian As String, strTotal As String
   
   If blnEdit Then
      blnEdit = False
      Exit Sub
   End If
   
   With cboSearch
      strBagian = .Text
      i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strBagian)
      
      If i <> -1 Then
         strTotal = .List(i)
         ii = Len(strTotal) - Len(strBagian)
         If ii <> 0 Then
             blnEdit = True
             .SelText = Right$(strTotal, ii)
             .SelStart = Len(strBagian)
             .SelLength = ii
         End If
         blnGaAda = False
      Else
         blnGaAda = True
      End If
   End With
End Sub
Private Sub cboSearch_KeyDown(KeyCode As Integer, Shift As Integer)
blnEdit = False
Select Case KeyCode
   Case vbKeyDelete
      blnEdit = True
   Case vbKeyBack
      blnEdit = True
End Select

Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyF1 Then
        If CtrlDown Then
            Call XPButton1_Click
        End If
    End If
        If KeyCode = vbKeyF2 Then
            If CtrlDown Then
                Call XPButton3_Click
            End If
        End If
    If KeyCode = vbKeyF3 Then
        If CtrlDown Then
            Call XPButton2_Click
        End If
    End If
        If KeyCode = vbKeyF4 Then
            If CtrlDown Then
                Call XPButton5_Click
            End If
        End If
    If KeyCode = vbKeyF5 Then
        If CtrlDown Then
            Call XPButton6_Click
        End If
    End If
        If KeyCode = vbKeyF6 Then
            If CtrlDown Then
                Call XPButton9_Click
            End If
        End If
    If KeyCode = vbKeyF7 Then
        If CtrlDown Then
            Call XPButton10_Click
        End If
    End If
End Sub
Private Sub cboSearch_KeyPress(KeyAscii As Integer)
Dim Tanya
If KeyAscii = 13 Then
    Tanya = MsgBox("BUKA LOCKER ?", vbQuestion + vbOKCancel, "KONFIRMASI")
    If Tanya = vbCancel Then
        Exit Sub
    Else
        NO_PORT = Text2
        'Port_Out 888, NO_PORT
    End If
End If
End Sub

Private Sub INFO_ARSIP()
SCari = "SELECT Count(B001.NOPOL) AS CountOfNOPOL FROM B001"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Label8 = RCari("CountOfNOPOL")
    Label9 = Max_Total
    Label10 = Label9 - Label8
    If Label10 = 0 Then
        SCari = "Select * From C013"
        Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
        If RCari.RowCount <> 0 Then
            RCari.Edit
                RCari("STS_JML") = "1"
            RCari.Update
        End If
        RCari.Close
        Set RCari = Nothing
    End If
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub HScroll1_Change()
AA = HScroll1.Value
MakeTransparent Me.hwnd, AA
End Sub

Private Sub XPButton10_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
CRPT.ReportFileName = App.Path & "\Report\AKTIFITAS.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = False
CRPT.WindowMinButton = False
CRPT.Action = 1
End Sub

Private Sub XPButton2_Click()
'Screen.MousePointer = vbHourglass
'FRM_PROSES.Show

Me.Hide
B002.Show

Unload FRM_PROSES
Set FRM_PROSES = Nothing
'Screen.MousePointer = vbDefault
End Sub

Private Sub Timer_Timer()
If Text.Left > -Text.Width Then
    Text.Left = Text.Left - 20
Else
    Text.Left = Picture1.ScaleWidth
End If
End Sub

Private Sub Ngumpet()
'DATA1.Visible = False
'TOOLS1.Visible = False
'LAPORAN1.Visible = False
End Sub

Private Sub XPButton1_Click()
'Screen.MousePointer = vbHourglass
'FRM_PROSES.Show

Me.Hide
B001.Show

Unload FRM_PROSES
Set FRM_PROSES = Nothing
'Screen.MousePointer = vbDefault
End Sub

Private Sub TOOLS0_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
    'TOOLS0.Visible = False
    'TOOLS1.Visible = True
    'XPButton5.Visible = True
    'XPButton6.Visible = True
    'XPButton5.SetFocus
    'DATA0.Visible = False
    'DATA1.Visible = False
    'LAPORAN0.Visible = False
    'LAPORAN1.Visible = False
    Text1 = " Merupakan menu untuk melakukan setting pada sistem "
End Sub

Private Sub TOOLS1_Click()
'TOOLS1.Visible = False
'TOOLS0.Visible = True
'TombolMati Me
    'DATA0.Visible = True
    'DATA1.Visible = True
    'LAPORAN0.Visible = True
    'LAPORAN1.Visible = True
    ClearTextBoxes Me
End Sub

'Private Sub LAPORAN0_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
    'LAPORAN0.Visible = False
    'LAPORAN1.Visible = True
    'XPButton9.Visible = True
    'XPButton10.Visible = True
    'XPButton9.SetFocus
    'DATA0.Visible = False
    'DATA1.Visible = False
    'TOOLS0.Visible = False
    'TOOLS1.Visible = False
'    Text1 = " Cetak Laporan seluruh arsip dan aktifitas filing cabinet "
'End Sub

'Private Sub LAPORAN1_Click()
'LAPORAN1.Visible = False
'LAPORAN0.Visible = True
'TombolMati Me
    'DATA0.Visible = True
    'DATA1.Visible = True
    'TOOLS0.Visible = True
    'TOOLS1.Visible = True
'    ClearTextBoxes Me
'End Sub

Private Sub XPButton3_Click()
'Screen.MousePointer = vbHourglass
'FRM_PROSES.Show

Me.Hide
B001_1.Show

Unload FRM_PROSES
Set FRM_PROSES = Nothing
'Screen.MousePointer = vbDefault
End Sub

Private Sub XPButton5_Click()
'Screen.MousePointer = vbHourglass
'FRM_PROSES.Show

Me.Hide
FRM_SOS.Show

Unload FRM_PROSES
Set FRM_PROSES = Nothing
'Screen.MousePointer = vbDefault
End Sub

Private Sub XPButton6_Click()
'Screen.MousePointer = vbHourglass
'FRM_PROSES.Show

Me.Hide
C013.Show

Unload FRM_PROSES
Set FRM_PROSES = Nothing
'Screen.MousePointer = vbDefault
End Sub

Private Sub XPButton9_Click()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
CRPT.ReportFileName = App.Path & "\Report\B001.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = False
CRPT.WindowMinButton = False
CRPT.Action = 1
End Sub

Private Sub XPButton1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyF1 Then
        If CtrlDown Then
            Call XPButton1_Click
        End If
    End If
        If KeyCode = vbKeyF2 Then
            If CtrlDown Then
                Call XPButton3_Click
            End If
        End If
    If KeyCode = vbKeyF3 Then
        If CtrlDown Then
            Call XPButton2_Click
        End If
    End If
        If KeyCode = vbKeyF4 Then
            If CtrlDown Then
                Call XPButton5_Click
            End If
        End If
    If KeyCode = vbKeyF5 Then
        If CtrlDown Then
            Call XPButton6_Click
        End If
    End If
        If KeyCode = vbKeyF6 Then
            If CtrlDown Then
                Call XPButton9_Click
            End If
        End If
    If KeyCode = vbKeyF7 Then
        If CtrlDown Then
            Call XPButton10_Click
        End If
    End If
End Sub

Private Sub XPButton3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyF1 Then
        If CtrlDown Then
            Call XPButton1_Click
        End If
    End If
        If KeyCode = vbKeyF2 Then
            If CtrlDown Then
                Call XPButton3_Click
            End If
        End If
    If KeyCode = vbKeyF3 Then
        If CtrlDown Then
            Call XPButton2_Click
        End If
    End If
        If KeyCode = vbKeyF4 Then
            If CtrlDown Then
                Call XPButton5_Click
            End If
        End If
    If KeyCode = vbKeyF5 Then
        If CtrlDown Then
            Call XPButton6_Click
        End If
    End If
        If KeyCode = vbKeyF6 Then
            If CtrlDown Then
                Call XPButton9_Click
            End If
        End If
    If KeyCode = vbKeyF7 Then
        If CtrlDown Then
            Call XPButton10_Click
        End If
    End If
End Sub

Private Sub XPButton2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyF1 Then
        If CtrlDown Then
            Call XPButton1_Click
        End If
    End If
        If KeyCode = vbKeyF2 Then
            If CtrlDown Then
                Call XPButton3_Click
            End If
        End If
    If KeyCode = vbKeyF3 Then
        If CtrlDown Then
            Call XPButton2_Click
        End If
    End If
        If KeyCode = vbKeyF4 Then
            If CtrlDown Then
                Call XPButton5_Click
            End If
        End If
    If KeyCode = vbKeyF5 Then
        If CtrlDown Then
            Call XPButton6_Click
        End If
    End If
        If KeyCode = vbKeyF6 Then
            If CtrlDown Then
                Call XPButton9_Click
            End If
        End If
    If KeyCode = vbKeyF7 Then
        If CtrlDown Then
            Call XPButton10_Click
        End If
    End If
End Sub

Private Sub XPButton5_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyF1 Then
        If CtrlDown Then
            Call XPButton1_Click
        End If
    End If
        If KeyCode = vbKeyF2 Then
            If CtrlDown Then
                Call XPButton3_Click
            End If
        End If
    If KeyCode = vbKeyF3 Then
        If CtrlDown Then
            Call XPButton2_Click
        End If
    End If
        If KeyCode = vbKeyF4 Then
            If CtrlDown Then
                Call XPButton5_Click
            End If
        End If
    If KeyCode = vbKeyF5 Then
        If CtrlDown Then
            Call XPButton6_Click
        End If
    End If
        If KeyCode = vbKeyF6 Then
            If CtrlDown Then
                Call XPButton9_Click
            End If
        End If
    If KeyCode = vbKeyF7 Then
        If CtrlDown Then
            Call XPButton10_Click
        End If
    End If
End Sub

Private Sub XPButton6_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyF1 Then
        If CtrlDown Then
            Call XPButton1_Click
        End If
    End If
        If KeyCode = vbKeyF2 Then
            If CtrlDown Then
                Call XPButton3_Click
            End If
        End If
    If KeyCode = vbKeyF3 Then
        If CtrlDown Then
            Call XPButton2_Click
        End If
    End If
        If KeyCode = vbKeyF4 Then
            If CtrlDown Then
                Call XPButton5_Click
            End If
        End If
    If KeyCode = vbKeyF5 Then
        If CtrlDown Then
            Call XPButton6_Click
        End If
    End If
        If KeyCode = vbKeyF6 Then
            If CtrlDown Then
                Call XPButton9_Click
            End If
        End If
    If KeyCode = vbKeyF7 Then
        If CtrlDown Then
            Call XPButton10_Click
        End If
    End If
End Sub

Private Sub XPButton9_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyF1 Then
        If CtrlDown Then
            Call XPButton1_Click
        End If
    End If
        If KeyCode = vbKeyF2 Then
            If CtrlDown Then
                Call XPButton3_Click
            End If
        End If
    If KeyCode = vbKeyF3 Then
        If CtrlDown Then
            Call XPButton2_Click
        End If
    End If
        If KeyCode = vbKeyF4 Then
            If CtrlDown Then
                Call XPButton5_Click
            End If
        End If
    If KeyCode = vbKeyF5 Then
        If CtrlDown Then
            Call XPButton6_Click
        End If
    End If
        If KeyCode = vbKeyF6 Then
            If CtrlDown Then
                Call XPButton9_Click
            End If
        End If
    If KeyCode = vbKeyF7 Then
        If CtrlDown Then
            Call XPButton10_Click
        End If
    End If
End Sub

Private Sub XPButton10_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyF1 Then
        If CtrlDown Then
            Call XPButton1_Click
        End If
    End If
        If KeyCode = vbKeyF2 Then
            If CtrlDown Then
                Call XPButton3_Click
            End If
        End If
    If KeyCode = vbKeyF3 Then
        If CtrlDown Then
            Call XPButton2_Click
        End If
    End If
        If KeyCode = vbKeyF4 Then
            If CtrlDown Then
                Call XPButton5_Click
            End If
        End If
    If KeyCode = vbKeyF5 Then
        If CtrlDown Then
            Call XPButton6_Click
        End If
    End If
        If KeyCode = vbKeyF6 Then
            If CtrlDown Then
                Call XPButton9_Click
            End If
        End If
    If KeyCode = vbKeyF7 Then
        If CtrlDown Then
            Call XPButton10_Click
        End If
    End If
End Sub
