VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Stream MP3 and WMA from Resources"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer tmrPos 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4350
      Top             =   0
   End
   Begin MSComctlLib.Slider sldPos 
      Height          =   270
      Left            =   150
      TabIndex        =   3
      Top             =   1125
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   476
      _Version        =   393216
      TickStyle       =   3
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   465
      Left            =   3300
      TabIndex        =   2
      Top             =   300
      Width           =   1515
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   465
      Left            =   1725
      TabIndex        =   1
      Top             =   300
      Width           =   1515
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   465
      Left            =   150
      TabIndex        =   0
      Top             =   300
      Width           =   1515
   End
   Begin MSComctlLib.Slider sldVol 
      Height          =   270
      Left            =   1125
      TabIndex        =   6
      Top             =   1875
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   476
      _Version        =   393216
      Min             =   -10000
      Max             =   0
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldBal 
      Height          =   270
      Left            =   1125
      TabIndex        =   8
      Top             =   2175
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   476
      _Version        =   393216
      Min             =   -10000
      Max             =   10000
      TickStyle       =   3
   End
   Begin VB.Label Label1 
      Caption         =   "MP3 gets played without being written to disk!"
      Height          =   240
      Left            =   300
      TabIndex        =   9
      Top             =   2625
      Width           =   4515
   End
   Begin VB.Label lblBal 
      Caption         =   "Balance:"
      Height          =   240
      Left            =   300
      TabIndex        =   7
      Top             =   2175
      Width           =   690
   End
   Begin VB.Label lblVol 
      Caption         =   "Volume:"
      Height          =   240
      Left            =   300
      TabIndex        =   5
      Top             =   1875
      Width           =   690
   End
   Begin VB.Label lblPos 
      AutoSize        =   -1  'True
      Caption         =   "0:00/0:00"
      Height          =   195
      Left            =   3975
      TabIndex        =   4
      Top             =   1425
      Width           =   705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_clsAudioMem    As AudioMemoryPlayer
Attribute m_clsAudioMem.VB_VarHelpID = -1

Private m_blnDontMove               As Boolean

Private Sub cmdPause_Click()
    If Not m_clsAudioMem.Pause() Then
        MsgBox "Couldn't pause!"
    End If
End Sub

Private Sub cmdPlay_Click()
    If Not m_clsAudioMem.Play() Then
        MsgBox "Couldn't play!"
    End If
End Sub

Private Sub cmdStop_Click()
    If Not m_clsAudioMem.StopPlayback() Then
        MsgBox "Couldn't stop!"
    End If
End Sub

Private Sub Form_Load()
    Dim btAudio()   As Byte
    
    Set m_clsAudioMem = New AudioMemoryPlayer

    ' load the MP3 file from the resource into btAudio
    btAudio = LoadResData("MP3SOUND", "CUSTOM")
    
    ' OpenStream will copy btAudio into its own allocated memory,
    ' btAudio can fall out of scope
    If Not m_clsAudioMem.OpenStream(VarPtr(btAudio(0)), UBound(btAudio) + 1) Then
        MsgBox "Couldn't open the audio stream!", vbExclamation
        cmdPlay.Enabled = False
    Else
        sldPos.Max = m_clsAudioMem.Duration
        sldPos.value = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' free allocated memory
    m_clsAudioMem.CloseStream
End Sub

Private Sub m_clsAudioMem_EndOfStream()
    Debug.Print "End Of Stream!"
    m_clsAudioMem_StatusChanged PlaybackStopped
End Sub

Private Sub m_clsAudioMem_StatusChanged(ByVal stat As PlaybackStatus)
    Select Case True
        Case stat = PlaybackPausing
            cmdPlay.Enabled = True
            cmdPause.Enabled = False
            cmdStop.Enabled = True
            tmrPos.Enabled = False
        Case stat = PlaybackPlaying
            cmdPlay.Enabled = False
            cmdPause.Enabled = True
            cmdStop.Enabled = True
            tmrPos.Enabled = True
        Case stat = PlaybackStopped
            cmdPlay.Enabled = True
            cmdPause.Enabled = False
            cmdStop.Enabled = False
            tmrPos.Enabled = False
    End Select
End Sub

Private Sub sldBal_Scroll()
    m_clsAudioMem.Balance = sldBal.value
End Sub

Private Sub sldPos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_blnDontMove = True
End Sub

Private Sub sldPos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_clsAudioMem.Position = sldPos.value
    m_blnDontMove = False
End Sub

Private Sub sldVol_Scroll()
    m_clsAudioMem.Volume = sldVol.value
End Sub

Private Sub tmrPos_Timer()
    If Not m_clsAudioMem Is Nothing Then
        With m_clsAudioMem
            lblPos.Caption = FmtMs(.Position) & "/" & FmtMs(.Duration)
            
            If Not m_blnDontMove Then
                sldPos.value = .Position
            End If
        End With
    End If
End Sub

Private Function FmtMs(ByVal lngMS As Long) As String
    Dim mins As Long, secs As Long
    
    secs = lngMS / 1000
    mins = secs / 60
    secs = secs Mod 60
    
    FmtMs = mins & ":" & Format(secs, "00")
End Function
