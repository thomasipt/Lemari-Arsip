VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form FRM_PROSES 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   315
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   315
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   900
      Top             =   900
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   315
      Top             =   900
   End
   Begin XPControls.ProgBarXP ProgBarXP1 
      Height          =   285
      Left            =   8
      TabIndex        =   0
      Top             =   15
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   503
      Max             =   10
      Style           =   2
      ScrollType      =   1
      BlockSize       =   50
      BarColor        =   16777088
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WaitBarTail     =   1
      WaitBarDelay    =   0
   End
End
Attribute VB_Name = "FRM_PROSES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = 0
Me.Left = (Screen.Width - Me.Width) / 2
'Screen.MousePointer = vbHourglass
A = 1
Timer1.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
A = A + 1
If A <= 10 Then
    ProgBarXP1.Value = A
Else
    Timer1.Enabled = False
    Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
A = A - 1
If A >= 0 Then
    ProgBarXP1.Value = A
Else
    Timer1.Enabled = True
    Timer2.Enabled = False
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Screen.MousePointer = vbDefault
End Sub

