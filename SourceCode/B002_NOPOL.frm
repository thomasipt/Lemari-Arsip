VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form B002_NOPOL 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BANTUAN"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   2280
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   195
      Width           =   3960
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3750
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Klik untuk memilih"
      Top             =   690
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6615
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16777088
      BackColorBkg    =   8421376
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KATA KUNCI"
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
      Left            =   195
      TabIndex        =   2
      Top             =   255
      Width           =   1560
   End
End
Attribute VB_Name = "B002_NOPOL"
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

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=LOCKER", rdDriverNoPrompt, False, CN)
ClearTextBoxes Me
Me.Left = Screen.Width / 2
Me.Top = (Screen.Height - Me.Height) / 2
Me.Caption = "BANTUAN " + N_FORM

Call SiapkanGrid

    If N_FORM = "NOMOR POLISI" Then
        SCari = "Select NOPOL From B001"
    ElseIf N_FORM = "MERK" Then
        SCari = "Select MERK From B001 GROUP BY B001.MERK"
    ElseIf N_FORM = "JENIS" Then
        SCari = "Select JENIS From B001 GROUP BY B001.JENIS"
    ElseIf N_FORM = "KODE LOKASI" Then
        SCari = "Select KODE_LOKASI From B001 GROUP BY B001.KODE_LOKASI"
    ElseIf N_FORM = "LEMARI" Then
        SCari = "Select NO_LMR From B001 GROUP BY B001.NO_LMR"
    ElseIf N_FORM = "LOKER" Then
        SCari = "Select NO_LKR From B001 GROUP BY B001.NO_LKR"
    ElseIf N_FORM = "BARIS" Then
        SCari = "Select NO_URT From B001 GROUP BY B001.NO_URT"
    End If
    
Call IsiGrid

End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 2
    .Row = 0
    .Col = 0: .ColWidth(0) = 1000: .text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 4500: .text = N_FORM: .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
Dim BARIS As String

Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
   RCari.MoveFirst
   B = 1
   BARIS = 0
   Do Until RCari.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .text = BARIS
              If N_FORM = "NOMOR POLISI" Then
                .Col = 1: .text = RCari("NOPOL")
              ElseIf N_FORM = "MERK" Then
                .Col = 1: .text = RCari("MERK")
              ElseIf N_FORM = "JENIS" Then
                .Col = 1: .text = RCari("JENIS")
              ElseIf N_FORM = "KODE LOKASI" Then
                .Col = 1: .text = RCari("KODE_LOKASI")
              ElseIf N_FORM = "LEMARI" Then
                .Col = 1: .text = RCari("NO_LMR")
              ElseIf N_FORM = "LOKER" Then
                .Col = 1: .text = RCari("NO_LKR")
              ElseIf N_FORM = "BARIS" Then
                .Col = 1: .text = RCari("NO_URT")
              End If
         End With
      B = B + 1
      RCari.MoveNext
      BARIS = BARIS + 1
   Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub grid_dblClick()
'sndPlaySound App.Path & "\CLICK.wav", SND_ASYNC
grid.Col = 1
N_FORM2 = ""
Clipboard.SetText (grid.text)
N_FORM2 = grid.text

If N_FORM2 = "" Then Exit Sub

Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If N_FORM = "NOMOR POLISI" Then
        SCari = "Select * From B001 where NOPOL like '%" + Trim(Text1) + "%'"
    ElseIf N_FORM = "MERK" Then
        SCari = "Select MERK From B001 where MERK like '%" + Trim(Text1) + "%' GROUP BY B001.MERK"
    ElseIf N_FORM = "JENIS" Then
        SCari = "Select JENIS From B001 where JENIS like '%" + Trim(Text1) + "%' GROUP BY B001.JENIS"
    ElseIf N_FORM = "KODE LOKASI" Then
        SCari = "Select KODE_LOKASI From B001 where KODE_LOKASI like '%" + Trim(Text1) + "%' GROUP BY B001.KODE_LOKASI"
    ElseIf N_FORM = "LEMARI" Then
        SCari = "Select NO_LMR From B001 where NO_LMR like '%" + Trim(Text1) + "%' GROUP BY B001.NO_LMR"
    ElseIf N_FORM = "LOKER" Then
        SCari = "Select NO_LKR From B001 where NO_LKR like '%" + Trim(Text1) + "%' GROUP BY B001.NO_LKR"
    ElseIf N_FORM = "BARIS" Then
        SCari = "Select NO_URT From B001 where NO_URT like '%" + Trim(Text1) + "%' GROUP BY B001.NO_URT"
    End If

Call IsiGrid

End Sub
