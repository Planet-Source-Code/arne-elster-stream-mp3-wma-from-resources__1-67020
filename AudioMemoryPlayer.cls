VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioMemoryPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Buffer Length < 1000 can only be even hundreds
' (200, 400, 600, 800). I can't explain why,
' but for odd hundreds smaller streams can be
' played back wrong (missing or swapped buffers).
' The bigger the buffer, the better ;)
Private Const StreamBufferLength    As Long = 2000

Public Event EndOfStream()
Public Event StatusChanged(ByVal stat As PlaybackStatus)

Private WithEvents m_clsAudioOut    As DirectSoundStream
Attribute m_clsAudioOut.VB_VarHelpID = -1
Private m_clsDirectSound            As DirectSound
Private m_clsWaveStream             As WmStream

Public Property Get Volume() As Long
    Volume = m_clsAudioOut.Volume
End Property

Public Property Let Volume(ByVal value As Long)
    m_clsAudioOut.Volume = value
End Property

Public Property Get Balance() As Long
    Balance = m_clsAudioOut.Balance
End Property

Public Property Let Balance(ByVal value As Long)
    m_clsAudioOut.Balance = value
End Property

Public Property Get Duration() As Long
    Duration = m_clsWaveStream.Duration
End Property

Public Property Get Position() As Long
    Position = m_clsAudioOut.Elapsed
End Property

Public Property Let Position(ByVal value As Long)
    Dim i                   As Long
    Dim intData()           As Integer
    Dim lngDataSize         As Long
    Dim lngRead             As Long
    Dim lngStreamPosition   As Long
    
    m_clsWaveStream.StreamSeek value \ 1000, SND_SEEK_SECONDS
    lngStreamPosition = m_clsWaveStream.Position
    
    If m_clsAudioOut.PlaybackStatus = PlaybackPlaying Then
        ' clear buffer queue
        m_clsAudioOut.PlaybackStop
    
        ' buffer 2 seconds of audio data
        lngDataSize = m_clsAudioOut.BytesFromMs(200)
        ReDim intData(lngDataSize \ 2 - 1) As Integer
    
        For i = 1 To 10
            m_clsWaveStream.StreamRead VarPtr(intData(0)), lngDataSize, lngRead
            If lngRead > 0 Then
                m_clsAudioOut.AudioBufferAdd VarPtr(intData(0)), lngRead
            Else
                Exit For
            End If
        Next
    
        m_clsAudioOut.PlaybackStart
    End If
    
    m_clsAudioOut.Elapsed = lngStreamPosition
End Property

Public Function StopPlayback() As Boolean
    If m_clsAudioOut.PlaybackStop() Then
        m_clsWaveStream.StreamSeek 0, SND_SEEK_PERCENT
        RaiseEvent StatusChanged(PlaybackStopped)
        StopPlayback = True
    End If
End Function

Public Function Pause() As Boolean
    If m_clsAudioOut.PlaybackPause() Then
        RaiseEvent StatusChanged(PlaybackPausing)
        Pause = True
    End If
End Function

Public Function Play() As Boolean
    Dim i           As Long
    Dim intData()   As Integer
    Dim lngDataSize As Long
    Dim lngRead     As Long
    
    If m_clsAudioOut.PlaybackStatus = PlaybackStopped Then
        ' buffer 2 seconds of audio data
        lngDataSize = m_clsAudioOut.BytesFromMs(200)
        ReDim intData(lngDataSize \ 2 - 1) As Integer
        
        For i = 1 To 10
            m_clsWaveStream.StreamRead VarPtr(intData(0)), lngDataSize, lngRead
            If lngRead > 0 Then
                m_clsAudioOut.AudioBufferAdd VarPtr(intData(0)), lngRead
            Else
                Exit For
            End If
        Next
    End If
    
    If m_clsAudioOut.PlaybackStart() Then
        RaiseEvent StatusChanged(PlaybackPlaying)
        Play = True
    End If
End Function

Public Function OpenStream(ByVal pMem As Long, ByVal DataSize As Long) As Boolean
    Dim clsStream   As DirectSoundStream
    
    CloseStream
    
    If m_clsWaveStream.StreamOpenByPointer(pMem, DataSize) = SND_ERR_SUCCESS Then
        With m_clsWaveStream
            If m_clsDirectSound.CreateStream(.Sameplerate, .Channels, .BitsPerSample, StreamBufferLength, clsStream) Then
                Set m_clsAudioOut = clsStream
                OpenStream = True
            Else
                m_clsWaveStream.StreamClose
            End If
        End With
    End If
End Function

Public Function CloseStream() As Boolean
    CloseStream = m_clsWaveStream.StreamClose() = SND_ERR_SUCCESS
    Set m_clsAudioOut = Nothing
End Function

Private Sub Class_Initialize()
    Set m_clsDirectSound = New DirectSound
    Set m_clsWaveStream = New WmStream
    
    m_clsDirectSound.Initialize 1, 44100, 2, 16
End Sub

Private Sub m_clsAudioOut_BufferDone()
    Dim intData()   As Integer
    Dim lngDataSize As Long
    Dim lngRead     As Long
    
    If Not m_clsWaveStream.EndOfStream Then
        lngDataSize = m_clsAudioOut.BytesFromMs(200)
        ReDim intData(lngDataSize \ 2 - 1) As Integer
    
        m_clsWaveStream.StreamRead VarPtr(intData(0)), lngDataSize, lngRead
        
        If lngRead > 0 Then
            m_clsAudioOut.AudioBufferAdd VarPtr(intData(0)), lngRead
        End If
    End If
End Sub

Private Sub m_clsAudioOut_NoDataLeft()
    Dim intData()   As Integer
    Dim lngDataSize As Long
    Dim lngRead     As Long
    Dim i           As Long
    
    If m_clsWaveStream.EndOfStream Then
        m_clsAudioOut.PlaybackStop
        m_clsWaveStream.StreamSeek 0, SND_SEEK_PERCENT
        RaiseEvent EndOfStream
    Else
        ' buffer underrun, buffer 2 seconds of audio data
        lngDataSize = m_clsAudioOut.BytesFromMs(200)
        ReDim intData(lngDataSize \ 2 - 1) As Integer
        
        For i = 1 To 10
            m_clsWaveStream.StreamRead VarPtr(intData(0)), lngDataSize, lngRead
            If lngRead > 0 Then
                m_clsAudioOut.AudioBufferAdd VarPtr(intData(0)), lngRead
            Else
                Exit For
            End If
        Next
    End If
End Sub
