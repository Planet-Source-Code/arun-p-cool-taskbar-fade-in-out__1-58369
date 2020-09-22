VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "__Cool TaskBar Transperency__"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   2640
      Width           =   375
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Min             =   1
      Max             =   75
      SelStart        =   12
      Value           =   12
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Min             =   1
      Max             =   75
      SelStart        =   6
      Value           =   12
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   767
      _Version        =   393216
      Max             =   255
      SelStart        =   128
      TickStyle       =   3
      Value           =   128
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4680
      Top             =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Fade Down Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Fade Up Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transparency Value"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Cursor At : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c) Arun P
'arun_pbk@rediffmail.com
'

Dim hWtask As Long
Dim ptXY As POINTAPI
Dim hBuf As Long

Dim blnUp As Boolean

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
     hWtask = FindWindow("Shell_traywnd", "") 'Get taskbar handle
     hBuf = 255
     intTransed = 0
End Sub

Private Sub Form_Terminate()
    SetWindowLong hWtask, GWL_EXSTYLE, WS_VISIBLE 'restore
End Sub

Private Sub Slider1_Click()
    If hBuf <> 255 Then
        SetLayeredWindowAttributes hWtask, 0, Slider1.Value, LWA_ALPHA
    End If
End Sub

Private Sub Timer1_Timer()
    GetCursorPos ptXY
    Label2.Caption = "Cursor At (" & ptXY.X & "," & ptXY.Y & ")"
    
    If ptXY.Y < Screen.Height * 10 / (11 * Screen.TwipsPerPixelY) Then
        'fade out
        If hBuf = 255 Then
            SetWindowLong hWtask, GWL_EXSTYLE, WS_EX_LAYERED
            For hBuf = 255 To Slider1.Value Step -Slider3.Value / 3
                SetLayeredWindowAttributes hWtask, 0, hBuf, LWA_ALPHA '--
                Sleep (1)
            Next hBuf
        End If
    Else
        'fade up
        For hBuf = Slider1.Value To 255 Step Slider2.Value / 3
            SetLayeredWindowAttributes hWtask, 0, hBuf, LWA_ALPHA '++
            Sleep (1)
        Next hBuf
        hBuf = 255
        SetWindowLong hWtask, GWL_EXSTYLE, WS_VISIBLE
    End If

End Sub

'''''''''''''''''''
