VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WmStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Windows Media Audio/MP3 Decoder
' (WMF SDK v9 Runtime or WMP 9/10 required)

Private Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal wFlags As Long, _
    ByVal dwBytes As Long _
) As Long

Private Declare Function GlobalFree Lib "kernel32" ( _
    ByVal hMem As Long _
) As Long

Private Declare Function GlobalLock Lib "kernel32" ( _
    ByVal hMem As Long _
) As Long

Private Declare Function GlobalUnlock Lib "kernel32" ( _
    ByVal hMem As Long _
) As Long

Private Declare Sub CpyMem Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal dwLen As Long _
)

Private Declare Function CreateStreamOnHGlobal Lib "ole32" ( _
    ByVal hGlobal As Long, _
    ByVal fDeleteOnRelease As Long, _
    ppStream As Any _
) As Long

Private Declare Function IsBadReadPtr Lib "kernel32" ( _
    ptr As Any, _
    ByVal ucb As Long _
) As Long

Private Declare Function IsBadWritePtr Lib "kernel32" ( _
    ptr As Any, _
    ByVal ucb As Long _
) As Long

Private Declare Function CallWindowProc Lib "user32" _
Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private Declare Function CreateFile Lib "kernel32.dll" _
Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long _
) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any _
) As Long

Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Any _
) As Long

Private Declare Function SetFilePointer Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lDistanceToMove As Long, _
    ByVal lpDistanceToMoveHigh As Long, _
    ByVal dwMoveMethod As Long _
) As Long

Private Declare Function GetFileSize Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpFileSizeHigh As Long _
) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long _
) As Long

Private Const INVALID_HANDLE    As Long = -1

Private Const GMEM_MOVEABLE     As Long = &H2
Private Const GMEM_NODISCARD    As Long = &H20

Private Type hFile
    handle      As Long
    path        As String
End Type

Private Enum FILE_OPEN_METHOD
    CREATE_NEW = 1
    CREATE_ALWAYS = 2
    OPEN_EXISTING = 3
    OPEN_ALWAYS = 4
End Enum

Private Enum FILE_SHARE_RIGHTS
    FILE_SHARE_READ = &H1
    FILE_SHARE_WRITE = &H2
End Enum

Private Enum FILE_ACCESS_RIGHTS
    GENERIC_READ = &H80000000
    GENERIC_WRITE = &H40000000
End Enum

Private Enum SEEK_METHOD
    FILE_BEGIN = 0
    FILE_CURRENT = 1
    FILE_END = 2
End Enum

Public Enum SND_RESULT
    SND_ERR_SUCCESS
    SND_ERR_INVALID_SOURCE
    SND_ERR_INVALID_OUTPUT
    SND_ERR_INTERNAL
    SND_ERR_OUT_OF_RANGE
    SND_ERR_END_OF_STREAM
    SND_ERR_INVALID_TAG
    SND_ERR_INVALID_PARAM
    SND_ERR_TOO_BIG
    SND_ERR_NEED_MORE
    SND_ERR_UNKNOWN
End Enum

Public Enum SND_SEEK_MODE
    SND_SEEK_PERCENT
    SND_SEEK_SECONDS
End Enum

' WMF SDK Header translation

Private Enum HRESULT
    S_OK = 0
End Enum

Private Enum WMT_VERSION
    WMT_VER_4_0 = &H40000
    WMT_VER_7_0 = &H70000
    WMT_VER_8_0 = &H80000
    WMT_VER_9_0 = &H90000
End Enum

Private Enum WMT_ATTR_DATATYPE
    WMT_TYPE_DWORD = 0
    WMT_TYPE_STRING = 1
    WMT_TYPE_BINARY = 2
    WMT_TYPE_BOOL = 3
    WMT_TYPE_QWORD = 4
    WMT_TYPE_WORD = 5
    WMT_TYPE_GUID = 6
End Enum

Private Enum WMT_ERRORS
    NS_E_NO_MORE_SAMPLES = &HC00D0BCF
End Enum

Private Enum WMT_RIGHTS
    WMT_RIGHT_PLAYBACK = &H1
    WMT_RIGHT_COPY_TO_NON_SDMI_DEVICE = &H2
    WMT_RIGHT_COPY_TO_CD = &H8
    WMT_RIGHT_COPY_TO_SDMI_DEVICE = &H10
    WMT_RIGHT_ONE_TIME = &H20
    WMT_RIGHT_SAVE_STREAM_PROTECTED = &H40
    WMT_RIGHT_COPY = &H80
    WMT_RIGHT_COLLABORATIVE_PLAY = &H100
    WMT_RIGHT_SDMI_TRIGGER = &H10000
    WMT_RIGHT_SDMI_NOMORECOPIES = &H20000
End Enum

Private Type MMWAVEFORMATEX
    wFormatTag                  As Integer
    nChannels                   As Integer
    nSamplesPerSec              As Long
    nAvgBytesPerSec             As Long
    nBlockAlign                 As Integer
    wBitsPerSample              As Integer
    cbSize                      As Integer
End Type

Private Type QWORD
    lo                          As Long
    hi                          As Long
End Type

Private Type GUID
    Data1                       As Long
    Data2                       As Integer
    Data3                       As Integer
    Data4(7)                    As Byte
End Type

Private Type IUnknown
    QueryInterface              As Long
    AddRef                      As Long
    Release                     As Long
End Type

Private Type IWMWriter
    Iunk                        As IUnknown
    SetProfileByID              As Long
    SetProfile                  As Long
    SetOutputFilename           As Long
    GetInputCount               As Long
    GetInputProps               As Long
    SetInputProps               As Long
    GetInputFormatCount         As Long
    GetInputFormat              As Long
    BeginWriting                As Long
    EndWriting                  As Long
    AllocateSample              As Long
    WriteSample                 As Long
    Flush                       As Long
End Type

Private Type INSSBuffer
    Iunk                        As IUnknown
    GetLength                   As Long
    SetLength                   As Long
    GetMaxLength                As Long
    GetBuffer                   As Long
    GetBufferAndLength          As Long
End Type

Private Type IWMHeaderInfo
    Iunk                        As IUnknown
    GetAttributeCount           As Long
    GetAttributeByIndex         As Long
    GetAttributeByName          As Long
    SetAttribute                As Long
    GetMarkerCount              As Long
    GetMarker                   As Long
    AddMarker                   As Long
    RemoveMarker                As Long
    GetScriptCount              As Long
    GetScript                   As Long
    AddScript                   As Long
    RemoveScript                As Long
End Type

Private Type IWMHeaderInfo2 ' : IWMHeaderInfo
    Iunk                        As IUnknown
    GetAttributeCount           As Long
    GetAttributeByIndex         As Long
    GetAttributeByName          As Long
    SetAttribute                As Long
    GetMarkerCount              As Long
    GetMarker                   As Long
    AddMarker                   As Long
    RemoveMarker                As Long
    GetScriptCount              As Long
    GetScript                   As Long
    AddScript                   As Long
    RemoveScript                As Long
    GetCodecInfoCount           As Long
    GetCodecInfo                As Long
End Type

Private Type IWMHeaderInfo3 ' : IWMHeaderInfo2
    Iunk                        As IUnknown
    GetAttributeCount           As Long
    GetAttributeByIndex         As Long
    GetAttributeByName          As Long
    SetAttribute                As Long
    GetMarkerCount              As Long
    GetMarker                   As Long
    AddMarker                   As Long
    RemoveMarker                As Long
    GetScriptCount              As Long
    GetScript                   As Long
    AddScript                   As Long
    RemoveScript                As Long
    GetCodecInfoCount           As Long
    GetCodecInfo                As Long
    GetAttributeCountEx         As Long
    GetAttributeIndices         As Long
    GetAttributeByIndexEx       As Long
    ModifyAttribute             As Long
    AddAttribute                As Long
    DeleteAttribute             As Long
    AddCodecInfo                As Long
End Type

Private Type IWMMediaProps
    Iunk                        As IUnknown
    GetType                     As Long
    GetMediaType                As Long
    SetMediaType                As Long
End Type

Private Type IWMInputMediaProps ' : IWMMediaProps
    Iunk                        As IUnknown
    GetType                     As Long
    GetMediaType                As Long
    SetMediaType                As Long
    GetConnectionName           As Long
    GetGroupName                As Long
End Type

Private Type IWMProfileManager
    Iunk                        As IUnknown
    CreateEmtpyProfile          As Long
    LoadProfileByID             As Long
    LoadProfileByData           As Long
    SaveProfile                 As Long
    GetSystemProfileCount       As Long
    LoadSystemProfile           As Long
End Type

Private Type IWMProfile
    Iunk                        As IUnknown
    GetVersion                  As Long
    GetName                     As Long
    SetName                     As Long
    GetDescription              As Long
    SetDescription              As Long
    GetStreamCount              As Long
    GetStream                   As Long
    GetStreamByNumber           As Long
    RemoveStream                As Long
    RemoveStreamByNumber        As Long
    AddStream                   As Long
    ReconfigStream              As Long
    CreateNewStream             As Long
    GetMutualExclusionCount     As Long
    AddMutualExclusion          As Long
    CreateNewMutualExclusion    As Long
End Type

Private Type IWMProfile2 ' : IWMProfile
    Iunk                        As IUnknown
    GetVersion                  As Long
    GetName                     As Long
    SetName                     As Long
    GetDescription              As Long
    SetDescription              As Long
    GetStreamCount              As Long
    GetStream                   As Long
    GetStreamByNumber           As Long
    RemoveStream                As Long
    RemoveStreamByNumber        As Long
    AddStream                   As Long
    ReconfigStream              As Long
    CreateNewStream             As Long
    GetMutualExclusionCount     As Long
    AddMutualExclusion          As Long
    CreateNewMutualExclusion    As Long
    GetProfileID                As Long
End Type

Private Type IWMProfile3
    Iunk                        As IUnknown
    GetVersion                  As Long
    GetName                     As Long
    SetName                     As Long
    GetDescription              As Long
    SetDescription              As Long
    GetStreamCount              As Long
    GetStream                   As Long
    GetStreamByNumber           As Long
    RemoveStream                As Long
    RemoveStreamByNumber        As Long
    AddStream                   As Long
    ReconfigStream              As Long
    CreateNewStream             As Long
    GetMutualExclusionCount     As Long
    AddMutualExclusion          As Long
    CreateNewMutualExclusion    As Long
    GetProfileID                As Long
    GetStorageFormat            As Long
    SetStorageFormat            As Long
    GetBandwidthSharingCount    As Long
    GetBandwidthSharing         As Long
    RemoveBandwidthSharing      As Long
    AddBandwidthSharing         As Long
    CreateNewBandwidthSharing   As Long
    GetStreamPrioritization     As Long
    SetStreamPrioritization     As Long
    RemoveStreamPrioritization  As Long
    CreateNewStreamPrioritization As Long
    GetExpectedPacketCount      As Long
End Type

Private Type IWMCodecInfo
    Iunk                        As IUnknown
    GetCodecInfoCount           As Long
    GetCodecFormatCount         As Long
    GetCodecFormat              As Long
End Type

Private Type IWMCodecInfo2
    Iunk                        As IUnknown
    GetCodecInfoCount           As Long
    GetCodecFormatCount         As Long
    GetCodecFormat              As Long
    GetCodecName                As Long
    GetCodecFormatDesc          As Long
End Type

Private Type IWMStreamConfig
    Iunk                        As IUnknown
    GetStreamType               As Long
    GetStreamNumber             As Long
    SetStreamNumber             As Long
    GetStreamName               As Long
    SetStreamName               As Long
    GetConnectionName           As Long
    SetConnectionName           As Long
    GetBitrate                  As Long
    SetBitrate                  As Long
    GetBufferWindow             As Long
    SetBufferWindow             As Long
End Type

Private Type IWMOutputMediaProps
    Iunk                        As IUnknown
    GetType                     As Long
    GetMediaType                As Long
    SetMediaType                As Long
    GetStreamGroupName          As Long
    GetConnectionName           As Long
End Type

Private Type IWMSyncReader
    Iunk                        As IUnknown
    Open                        As Long
    Close                       As Long
    SetRange                    As Long
    SetRangeByFrame             As Long
    GetNextSample               As Long
    SetStreamsSelected          As Long
    GetStreamSelected           As Long
    SetReadStreamSamples        As Long
    GetReadStreamSamples        As Long
    GetOutputSetting            As Long
    SetOutputSetting            As Long
    GetOutputCount              As Long
    GetOutputProps              As Long
    SetOutputProps              As Long
    GetOutputFormatCount        As Long
    GetOutputFormat             As Long
    GetOutputNumberForStream    As Long
    GetStreamNumberForOutput    As Long
    GetMaxOutputSampleSize      As Long
    GetMaxStreamSampleSize      As Long
    OpenStream                  As Long
End Type

Private Type IStream
    Iunk                        As IUnknown
End Type

Private Type WM_MEDIA_TYPE
    majortype                   As GUID
    subtype                     As GUID
    bFixedSizeSamples           As Long
    bTemporalCompression        As Long
    lSampleSize                 As Long
    formattype                  As GUID
    pUnk                        As Long
    cbFormat                    As Long
    pbFormat                    As Long
End Type


Private Const WMMEDIATYPE_Audio      As String _
    = "{73647561-0000-0010-8000-00AA00389B71}"

Private Const WMMEDIASUBTYPE_PCM     As String _
    = "{00000001-0000-0010-8000-00AA00389B71}"


Private Const WMFORMAT_WaveFormatEx  As String _
    = "{05589f81-c356-11ce-bf01-00aa0055595a}"


Private Const IID_IWMHeaderInfo      As String _
    = "{96406BDA-2B2B-11d3-B36B-00C04F6108FF}"

Private Const IID_IWMHeaderInfo2     As String _
    = "{15CF9781-454E-482e-B393-85FAE487A810}"

Private Const IID_IWMHeaderInfo3     As String _
    = "{15CC68E3-27CC-4ecd-B222-3F5D02D80BD5}"


Private Const IID_IWMProfileManager  As String _
    = "{d16679f2-6ca0-472d-8d31-2f5d55aee155}"

Private Const IID_IWMProfileManager2 As String _
    = "{7A924E51-73C1-494d-8019-23D37ED9B89A}"


Private Const IID_IWMCodecInfo       As String _
    = "{A970F41E-34DE-4a98-B3BA-E4B3CA7528F0}"

Private Const IID_IWMCodecInfo2     As String _
    = "{AA65E273-B686-4056-91EC-DD768D4DF710}"


Private Const IID_IWMMediaProps      As String _
    = "{96406BCE-2B2B-11d3-B36B-00C04F6108FF}"


Private Const WMMEDIASUBTYPE_DRM     As String _
    = "{00000009-0000-0010-8000-00AA00389B71}"



Private Const IID_IUnknown           As String _
    = "{00000000-0000-0000-C000-000000000046}"

Private Const IID_IWMOutputMediaProps As String _
    = "{96406BD7-2B2B-11d3-B36B-00C04F6108FF}"

Private Const IID_IWMSyncReader      As String _
    = "{9397F121-7705-4dc9-B049-98B698188414}"

Private Const IID_INSSBuffer         As String _
    = "{E1CD3524-03D7-11d2-9EED-006097D2D7CF}"


Private Const attr_WMDuration    As String = "Duration"
Private Const attr_WMBitrate     As String = "Bitrate"
Private Const attr_WMSeekable    As String = "Seekable"
Private Const attr_WMStridable   As String = "Stridable"
Private Const attr_WMBroadcast   As String = "Broadcast"
Private Const attr_WMProtected   As String = "Is_Protected"
Private Const attr_WMTrusted     As String = "Is_Trusted"
Private Const attr_WMSigName     As String = "Signature_Name"
Private Const attr_WMHasAudio    As String = "HasAudio"
Private Const attr_WMHasImage    As String = "HasImage"
Private Const attr_WMHasScript   As String = "HasScript"
Private Const attr_WMHasVideo    As String = "HasVideo"
Private Const attr_WMCurBitrate  As String = "CurrentBitrate"
Private Const attr_WMOptBitrate  As String = "OptimalBitrate"
Private Const attr_WMSkipBackw   As String = "Can_Skip_Backward"
Private Const attr_WMSkipForw    As String = "Can_Skip_Forward"
Private Const attr_WMNumFrames   As String = "NumberOfFrames"
Private Const attr_WMFileSize    As String = "FileSize"

Private Const attr_WMTitle       As String = "Title"
Private Const attr_WMAuthor      As String = "Author"
Private Const attr_WMDescript    As String = "Description"
Private Const attr_WMRating      As String = "Rating"
Private Const attr_WMCopyright   As String = "Copyright"

Private Const attr_WMAlbumTitle  As String = "WM/AlbumTitle"
Private Const attr_WMAlbumArtist As String = "WM/AlbumArtist"
Private Const attr_WMTrack       As String = "WM/Track"
Private Const attr_WMGenre       As String = "WM/Genre"
Private Const attr_WMYear        As String = "WM/Year"
Private Const attr_WMGenreID     As String = "WM/GenreID"
Private Const attr_WMMCDI        As String = "WM/MCDI"
Private Const attr_WMComposer    As String = "WM/Composer"
Private Const attr_WMLyrics      As String = "WM/Lyrics"
Private Const attr_WMTrackNumber As String = "WM/TrackNumber"
Private Const attr_WMIsVBR       As String = "IsVBR"

' Synchronous WMF Reader
Private IReader                 As IWMSyncReader
Private oReader                 As Long
Private pReaderVTbl             As Long

' Media Properties
Private IProps                  As IWMOutputMediaProps
Private oProps                  As Long
Private pPropsVTbl              As Long

' Media Header
Private IHeader                 As IWMHeaderInfo3
Private oHeader                 As Long
Private pHeaderVTbl             As Long

' IStream
Private IStream                 As IStream
Private oStream                 As Long
Private pStreamVTbl             As Long
Private udtStreamMem            As Memory

' waveOut WAVEFORMATEX
Private wfxout                  As MMWAVEFORMATEX

' WMA Attribute
Private curDuration             As Currency
Private curBitrate              As Currency
Private curPosition             As Currency
Private blnSeekable             As Boolean
Private blnEOS                  As Boolean

' PCM Buffer
Private btBuffer()              As Byte
Private lngBufferData           As Long
Private lngPosInBuffer          As Long

' WM Core DLL
Private clsWMCore               As DynamicLinkLibrary

Private blnFileOpened           As Boolean

'Private clsTags                 As StreamTags

''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''

Private Function FillBuffer( _
) As Boolean

    Dim lngRet          As Long
    Dim ISample         As INSSBuffer
    Dim oSample         As Long
    Dim pSampleVTbl     As Long
    Dim curSampleDur    As Currency
    Dim pdwFlags        As Long
    Dim pdwOutputNum    As Long
    Dim ppdwBuffer      As Long
    Dim pdwLength       As Long

    If oReader = 0 Then Exit Function

    ' get some more samples from the stream
    lngRet = CallPointer(IReader.GetNextSample, _
                         oReader, _
                         0, _
                         VarPtr(oSample), _
                         VarPtr(curPosition), _
                         VarPtr(curSampleDur), _
                         VarPtr(pdwFlags), _
                         VarPtr(pdwOutputNum), _
                         0)

    If lngRet <> 0 Then
        blnEOS = lngRet = NS_E_NO_MORE_SAMPLES
        Exit Function
    End If

    ' INSSBuffer's VTable
    CpyMem pSampleVTbl, ByVal oSample, 4
    CpyMem ISample, ByVal pSampleVTbl, Len(ISample)

    ' get a pointer to the samples and their length
    lngRet = CallPointer(ISample.GetBufferAndLength, _
                         oSample, _
                         VarPtr(ppdwBuffer), _
                         VarPtr(pdwLength))

    If lngRet <> 0 Then
        lngRet = CallPointer(ISample.Iunk.Release, oSample)
        Exit Function
    End If

    ' fill PCM buffer
    ReDim btBuffer(pdwLength - 1) As Byte
    CpyMem btBuffer(0), ByVal ppdwBuffer, pdwLength

    ' release INSSBuffer
    lngRet = CallPointer(ISample.Iunk.Release, oSample)

    ' reset buffer position and size
    lngPosInBuffer = 0
    lngBufferData = pdwLength

    FillBuffer = True
End Function

Private Function GUID2Str( _
    g As GUID _
) As String

    Dim nTemp   As String

    nTemp = "{"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data1), 8) & "-"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data2), 4) & "-"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data3), 4) & "-"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(0)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(1)), 2) & "-"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(2)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(3)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(4)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(5)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(6)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(7)), 2) & "}"

    GUID2Str = nTemp
End Function

Private Function GUIDEqual( _
    g1 As GUID, _
    g2 As GUID _
) As Boolean

    GUIDEqual = GUID2Str(g1) = GUID2Str(g2)
End Function

Private Function Str2GUID( _
    strGUID As String _
) As GUID

    Dim g           As GUID
    Dim strParts()  As String

    strParts = Split(Mid$(strGUID, 2, Len(strGUID) - 2), "-")

    g.Data1 = Val("&H" & strParts(0))
    g.Data2 = Val("&H" & strParts(1))
    g.Data3 = Val("&H" & strParts(2))
    g.Data4(0) = Val("&H" & Mid$(strParts(3), 1, 2))
    g.Data4(1) = Val("&H" & Mid$(strParts(3), 3, 2))
    g.Data4(2) = Val("&H" & Mid$(strParts(4), 1, 2))
    g.Data4(3) = Val("&H" & Mid$(strParts(4), 3, 2))
    g.Data4(4) = Val("&H" & Mid$(strParts(4), 5, 2))
    g.Data4(5) = Val("&H" & Mid$(strParts(4), 7, 2))
    g.Data4(6) = Val("&H" & Mid$(strParts(4), 9, 2))
    g.Data4(7) = Val("&H" & Mid$(strParts(4), 11, 2))

    Str2GUID = g
End Function

Private Function RemNullChars( _
    ByVal strString As String _
) As String

    If InStr(strString, Chr$(0)) > 0 Then
        RemNullChars = Left$(strString, InStr(strString, Chr$(0)) - 1)
    Else
        RemNullChars = strString
    End If
End Function

Private Sub CpyAttribute( _
    pDst As Long, _
    pSrc As Long, _
    dt As WMT_ATTR_DATATYPE _
)

    Dim nSize   As Long

    Select Case dt
        Case WMT_TYPE_BOOL: nSize = 2
        Case WMT_TYPE_DWORD: nSize = 4
        Case WMT_TYPE_GUID: nSize = 16
        Case WMT_TYPE_QWORD: nSize = 8
        Case WMT_TYPE_WORD: nSize = 2
    End Select

    CpyMem ByVal pDst, ByVal pSrc, nSize
End Sub

Private Function FmtStrLen( _
    ByVal B As String, _
    length As Integer _
) As String

    B = String(length - Len(B), "0") & B
    FmtStrLen = B
End Function

' http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=49987&lngWId=1
Private Function DecToHex( _
    ByVal nSource As Double _
) As String

    Const BASECHAR  As String = "0123456789ABCDEF"
    Dim n           As Double
    Dim nSrc        As Double

    On Error GoTo ErrorLine

    If (nSource = 0) Then
        DecToHex = "00"
        Exit Function
    End If

    If (nSource < 2147483648#) Then
        DecToHex = Hex(nSource)
    Else
        nSrc = nSource

        Do
            n = CDec(nSrc - (16 * Int(nSrc / 16)))
            DecToHex = Mid$(BASECHAR, n + 1, 1) & DecToHex
            nSrc = CDec(Int(nSrc / 16))
        Loop While (nSrc > 0)

    End If

    If (Len(DecToHex) Mod 2) Then
        DecToHex = "0" & DecToHex
    End If
    Exit Function

ErrorLine:
    DecToHex = ""
End Function

Private Function CallPointer( _
    ByVal fnc As Long, _
    ParamArray Params() _
) As Long

    Dim udtMem              As Memory
    Dim pASM                As Long
    Dim i                   As Integer

    If fnc = 0 Then
        Err.Raise 255, , "Null Pointer called!"
    End If

    udtMem = AllocMemory(1024, , PAGE_EXECUTE_READWRITE)
    pASM = udtMem.address

    AddByte pASM, &H58                  ' POP EAX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H50                  ' PUSH EAX

    If UBound(Params) = 0 Then
        If IsArray(Params(0)) Then
            For i = UBound(Params(0)) To 0 Step -1
                AddPush pASM, CLng(Params(0)(i))    ' PUSH dword
            Next
        Else
            For i = UBound(Params) To 0 Step -1
                AddPush pASM, CLng(Params(i))       ' PUSH dword
            Next
        End If
    Else
        For i = UBound(Params) To 0 Step -1
            AddPush pASM, CLng(Params(i))           ' PUSH dword
        Next
    End If

    AddCall pASM, fnc                   ' CALL rel addr
    AddByte pASM, &HC3                  ' RET

    CallPointer = CallWindowProc(udtMem.address, _
                                 0, 0, 0, 0)
                                 
    FreeMemory udtMem
End Function

Private Sub AddPush( _
    pASM As Long, _
    lng As Long _
)

    AddByte pASM, &H68
    AddLong pASM, lng
End Sub

Private Sub AddCall( _
    pASM As Long, _
    addr As Long _
)

    AddByte pASM, &HE8
    AddLong pASM, addr - pASM - 4
End Sub

Private Sub AddLong( _
    pASM As Long, _
    lng As Long _
)

    CpyMem ByVal pASM, lng, 4
    pASM = pASM + 4
End Sub

Private Sub AddByte( _
    pASM As Long, _
    Bt As Byte _
)

    CpyMem ByVal pASM, Bt, 1
    pASM = pASM + 1
End Sub

Private Sub Class_Initialize()
    Dim lngRet  As Long

    Set clsWMCore = New DynamicLinkLibrary
    'Set clsTags = New StreamTags

    If Not clsWMCore.LoadDLL("wmvcore.dll") Then
        Exit Sub
    End If

    ' create a IWMSyncReader instance for decoding
    lngRet = clsWMCore.CallFunc("WMCreateSyncReader", 0, WMT_RIGHT_PLAYBACK, VarPtr(oReader))
    If lngRet <> 0 Then
        oReader = 0
        Exit Sub
    End If

    ' IWMSyncReader`s VTable
    CpyMem pReaderVTbl, ByVal oReader, 4
    CpyMem IReader, ByVal pReaderVTbl, Len(IReader)
End Sub

Private Sub Class_Terminate()
    Dim lngRet  As Long

    StreamClose

    ' release IWMSyncReader
    ' (Set IWMSyncReader = Nothing)
    If oReader <> 0 Then
        CallPointer IReader.Iunk.Release, oReader
    End If

    clsWMCore.UnloadDLL
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get BitsPerSample( _
) As Integer

    BitsPerSample = wfxout.wBitsPerSample
End Property

Public Property Get BitsPerSecond( _
) As Long

    BitsPerSecond = curBitrate * 10000
End Property

Public Property Get Channels( _
) As Integer

    Channels = wfxout.nChannels
End Property

Public Property Get Duration( _
) As Long

    Duration = curDuration
End Property

Public Property Get Position( _
) As Long

    Position = curPosition
End Property

Public Property Get Sameplerate( _
) As Long

    Sameplerate = wfxout.nSamplesPerSec
End Property

'Private Property Get Tags( _
') As StreamTags
'
'    Set Tags = clsTags
'End Property

''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Description( _
) As String

    Description = "Windows Media Audio"
End Property

Public Property Get EndOfStream( _
) As Boolean

    EndOfStream = blnEOS
End Property

Public Function Extensions( _
) As String()

    Extensions = Split("WMA;MP3", ";")
End Function

Public Function StreamClose( _
) As SND_RESULT

    Dim lngRet  As Long

    If oReader = 0 Then
        StreamClose = SND_ERR_INTERNAL
        Exit Function
    End If

    If Not blnFileOpened Then
        StreamClose = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    lngRet = CallPointer(IReader.Close, oReader)
    
    If oStream <> 0 Then
        CallPointer IStream.Iunk.Release, oStream
        oStream = 0
    End If
    
    blnFileOpened = False

    blnEOS = False

    StreamClose = SND_ERR_SUCCESS
End Function

Public Function StreamOpenByPointer( _
    ByVal pMemory As Long, _
    ByVal DataSize As Long _
) As SND_RESULT

    Dim lngRet  As Long, pMem   As Long

    If oReader = 0 Then
        StreamOpenByPointer = SND_ERR_INTERNAL
        Exit Function
    End If
    
    StreamClose

    udtStreamMem.address = GlobalAlloc(GMEM_MOVEABLE Or GMEM_NODISCARD, DataSize)
    If udtStreamMem.address = 0 Then
        StreamOpenByPointer = SND_ERR_UNKNOWN
        Exit Function
    End If

    udtStreamMem.bytes = DataSize
    pMem = GlobalLock(udtStreamMem.address)
    CpyMem ByVal pMem, ByVal pMemory, DataSize
    GlobalUnlock udtStreamMem.address
    
    If CreateStreamOnHGlobal(udtStreamMem.address, 1, oStream) <> 0 Then
        GlobalFree udtStreamMem.address
        oStream = 0
        StreamOpenByPointer = SND_ERR_UNKNOWN
        Exit Function
    End If
    
    CpyMem pStreamVTbl, ByVal oStream, 4
    CpyMem IStream, ByVal pStreamVTbl, Len(IStream)
    
    ' open the stream we just allocated
    lngRet = CallPointer(IReader.OpenStream, _
                         oReader, _
                         oStream)

    If lngRet <> 0 Then
        CallPointer IStream.Iunk.Release, oStream
        oStream = 0
        StreamOpenByPointer = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    If Not GetOutputFormat() Then
        CallPointer IReader.Close, oReader
        CallPointer IStream.Iunk.Release, oStream
        oStream = 0
        StreamOpenByPointer = SND_ERR_UNKNOWN
        Exit Function
    End If

    lngRet = CallPointer(IReader.SetRangeByFrame, _
                         oReader, _
                         1, 0, 0, _
                         0, 0)

    If lngRet <> 0 Then
        lngRet = CallPointer(IReader.Close, oReader)
        CallPointer IStream.Iunk.Release, oStream
        oStream = 0
        StreamOpenByPointer = SND_ERR_UNKNOWN
        Exit Function
    End If

    If Not GetAttributes Then
        lngRet = CallPointer(IReader.Close, oReader)
        CallPointer IStream.Iunk.Release, oStream
        oStream = 0
        StreamOpenByPointer = SND_ERR_UNKNOWN
        Exit Function
    End If

    blnFileOpened = True

    StreamOpenByPointer = SND_ERR_SUCCESS
End Function

Public Function StreamOpen( _
    ByVal SourceFile As String _
) As SND_RESULT

    Dim lngRet  As Long

    ' we got an IWMSyncReader instance?
    If oReader = 0 Then
        StreamOpen = SND_ERR_INTERNAL
        Exit Function
    End If

    If Not IsValidFile(SourceFile) Then
        StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    StreamClose

    ' open the file
    lngRet = CallPointer(IReader.Open, _
                         oReader, _
                         StrPtr(SourceFile))

    If lngRet <> 0 Then
        StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    If Not GetOutputFormat() Then
        CallPointer IReader.Close, oReader
        StreamOpen = SND_ERR_UNKNOWN
        Exit Function
    End If

    ' range to decode (0 to EOF)
    lngRet = CallPointer(IReader.SetRangeByFrame, _
                         oReader, _
                         1, 0, 0, _
                         0, 0)

    If lngRet <> 0 Then
        lngRet = CallPointer(IReader.Close, oReader)
        StreamOpen = SND_ERR_UNKNOWN
        Exit Function
    End If

    If Not GetAttributes Then
        lngRet = CallPointer(IReader.Close, oReader)
        StreamOpen = SND_ERR_UNKNOWN
        Exit Function
    End If

    blnFileOpened = True

    StreamOpen = SND_ERR_SUCCESS
End Function

Private Function GetAttributes( _
) As Boolean

    Dim lngRet      As Long
    Dim cAttrs      As Long
    Dim attIndex    As Long
    Dim cchName     As Long
    Dim cbValue     As Long

    Dim pwszName    As String
    Dim strValue    As String

    Dim pbValue()   As Byte

    Dim attType     As WMT_ATTR_DATATYPE

    Dim iid_hdr     As GUID

    'clsTags.ClearTags

    ' get IWMheaderInfo3 from IWMSyncReader
    iid_hdr = Str2GUID(IID_IWMHeaderInfo3)
    lngRet = CallPointer(IReader.Iunk.QueryInterface, _
                         oReader, _
                         VarPtr(iid_hdr), _
                         VarPtr(oHeader))

    ' IWMHeaderInfo3`s VTable
    CpyMem pHeaderVTbl, ByVal oHeader, 4
    CpyMem IHeader, ByVal pHeaderVTbl, Len(IHeader)

    ' attributes count
    lngRet = CallPointer(IHeader.GetAttributeCountEx, _
                         oHeader, _
                         &HFFFF&, _
                         VarPtr(cAttrs))

    If lngRet <> 0 Then
        lngRet = CallPointer(IHeader.Iunk.Release, oHeader)
        oHeader = 0
        Exit Function
    End If

    For attIndex = 0 To cAttrs - 1

        ' size of the attribute's name and content
        lngRet = CallPointer(IHeader.GetAttributeByIndexEx, _
                             oHeader, _
                             &HFFFF&, _
                             attIndex, _
                             0, _
                             VarPtr(cchName), _
                             0, 0, 0, _
                             VarPtr(cbValue))

        If lngRet <> 0 Then GoTo NextAttribute

        ' buffer for name and content
        pwszName = Space(cchName)
        ReDim pbValue(cbValue - 1) As Byte

        ' get the attribute
        lngRet = CallPointer(IHeader.GetAttributeByIndexEx, _
                             oHeader, _
                             &HFFFF&, _
                             attIndex, _
                             StrPtr(pwszName), _
                             VarPtr(cchName), _
                             VarPtr(attType), _
                             0, _
                             VarPtr(pbValue(0)), _
                             VarPtr(cbValue))

        If lngRet <> 0 Then GoTo NextAttribute

        pwszName = RemNullChars(pwszName)

        Select Case attType
            Case WMT_TYPE_WORD                  ' integer
                Dim intVal  As Integer
                CpyMem intVal, pbValue(0), 2
                strValue = intVal
            Case WMT_TYPE_DWORD                 ' long
                Dim lngVal  As Long
                CpyMem lngVal, pbValue(0), 4
                strValue = lngVal
            Case WMT_TYPE_QWORD                 ' currency?
                Dim curVal  As Currency
                CpyMem curVal, pbValue(0), 8
                strValue = curVal
            Case WMT_TYPE_STRING                ' Unicode String
                strValue = RemNullChars(pbValue)
            Case WMT_TYPE_BINARY                ' binary...
                ' nicht lesbar...
                ' Debug.Print "<binary>"
            Case WMT_TYPE_BOOL                  ' boolean
                CpyMem lngVal, pbValue(0), cbValue
                strValue = CBool(lngVal)
            Case WMT_TYPE_GUID                  ' global unique identifier
                Dim g   As GUID
                CpyMem g, pbValue(0), Len(g)
                strValue = GUID2Str(g)
        End Select

        ' attributes we want to save
        Select Case pwszName

            Case attr_WMDuration
                CpyAttribute VarPtr(curDuration), _
                             VarPtr(pbValue(0)), _
                             attType

            Case attr_WMBitrate
                CpyAttribute VarPtr(curBitrate), _
                             VarPtr(pbValue(0)), _
                             attType

            Case attr_WMSeekable
                CpyAttribute VarPtr(blnSeekable), _
                             VarPtr(pbValue(0)), _
                             attType

'            Case attr_WMAuthor
'                clsTags.AddTag "artist", strValue
'
'            Case attr_WMAlbumArtist
'                clsTags.AddTag "artist", strValue
'
'            Case attr_WMAlbumTitle
'                clsTags.AddTag "album", strValue
'
'            Case attr_WMGenre
'                clsTags.AddTag "genre", strValue
'
'            Case attr_WMTitle
'                clsTags.AddTag "title", strValue

'            Case attr_WMTrackNumber
'                If lngVal < 100 And lngVal > -1 Then
'                    clsTags.AddTag "tracknumber", CStr(lngVal)
'                Else
'                    clsTags.AddTag "tracknumber", strValue
'                End If
'
'            Case attr_WMTrackNumber
'                If lngVal < 100 And lngVal > -1 Then
'                    clsTags.AddTag "tracknumber", CStr(lngVal)
'                Else
'                    clsTags.AddTag "tracknumber", strValue
'                End If
'
'            Case attr_WMDescript
'                clsTags.AddTag "comment", strValue
'
'            Case attr_WMYear
'                clsTags.AddTag "year", strValue

        End Select

NextAttribute:
    Next

    ' release IWMHeaderInfo3
    lngRet = CallPointer(IHeader.Iunk.Release, oHeader)
    oHeader = 0

    GetAttributes = True
End Function

Private Function GetOutputFormat( _
) As Boolean

    Dim lngRet      As Long
    Dim pcOutputs   As Long
    Dim pcFormats   As Long
    Dim pcbType     As Long
    Dim i           As Long
    Dim j           As Long
    Dim MediaType   As WM_MEDIA_TYPE

    Dim btMT()      As Byte

    ' output count
    lngRet = CallPointer(IReader.GetOutputCount, _
                         oReader, _
                         VarPtr(pcOutputs))

    If lngRet <> 0 Then
        lngRet = CallPointer(IReader.Close, oReader)
        Exit Function
    End If

    ' format of each output
    For i = 0 To pcOutputs - 1

        ' number of formats for the current output
        lngRet = CallPointer(IReader.GetOutputFormatCount, _
                             oReader, _
                             i, _
                             VarPtr(pcFormats))

        If lngRet <> 0 Then GoTo NextOutput

        ' details for the formats
        For j = 0 To pcFormats

            ' get a IWMOutputMediaProps instance for the current format
            lngRet = CallPointer(IReader.GetOutputFormat, _
                                 oReader, _
                                 j, i, _
                                 VarPtr(oProps))

            If lngRet <> 0 Then GoTo NextFormat

            ' IWMOutputMediaProps`s VTable
            CpyMem pPropsVTbl, ByVal oProps, 4
            CpyMem IProps, ByVal pPropsVTbl, Len(IProps)

            ' It's not allowed to get the WM_MEDIA_TYPE directly,
            ' you first have to query it for its size and allocate
            ' enough memory, then you can recieve it.

            lngRet = CallPointer(IProps.GetMediaType, _
                                 oProps, _
                                 0, _
                                 VarPtr(pcbType))

            If lngRet <> 0 Then
                lngRet = CallPointer(IProps.Iunk.Release, oProps)
                oProps = 0
                GoTo NextFormat
            Else
                ReDim btMT(pcbType - 1) As Byte
            End If

            lngRet = CallPointer(IProps.GetMediaType, _
                                 oProps, _
                                 VarPtr(btMT(0)), _
                                 VarPtr(pcbType))

            CpyMem MediaType, btMT(0), Len(MediaType)

            ' Audio?
            If GUIDEqual(MediaType.majortype, Str2GUID(WMMEDIATYPE_Audio)) Then
                ' PCM?
                If GUIDEqual(MediaType.subtype, Str2GUID(WMMEDIASUBTYPE_PCM)) Then
                    ' 16 Bit?
                    CpyMem wfxout, ByVal MediaType.pbFormat, Len(wfxout)
                    If wfxout.wBitsPerSample = 16 Then
                        lngRet = CallPointer(IProps.SetMediaType, oProps, VarPtr(btMT(0)))
                        GoTo ExitLoops
                    End If
                End If
            End If

            ' release IWMOutputMediaProps
            lngRet = CallPointer(IProps.Iunk.Release, oProps)
            oProps = 0

NextFormat:         ' Not beautiful,
        Next j

NextOutput:         ' but makes nested loops
    Next i
ExitLoops:          ' a breeze.

    ' No format found?
    If i = pcOutputs Then
        Exit Function
    End If

    ' set new output format
    lngRet = CallPointer(IReader.SetOutputProps, oReader, i, oProps)
    lngRet = CallPointer(IProps.Iunk.Release, oProps)
    oProps = 0

    GetOutputFormat = True
End Function

Public Function StreamRead( _
    ByVal buffer_ptr As Long, _
    ByVal buffer_len As Long, _
    buffer_read As Long _
) As SND_RESULT

    StreamRead = SND_ERR_SUCCESS

    buffer_read = 0

    ' decode while not all requested bytes have been read
    ' or the end of the stream has been reached
    Do While buffer_read < buffer_len

        ' PCM Buffer is empty
        If lngBufferData = 0 Then
            If Not FillBuffer Then
                StreamRead = SND_ERR_END_OF_STREAM
                Exit Function
            End If

        ' data in the PCM buffer isn' enough to
        ' return the amount of bytes left
        ElseIf (lngBufferData - lngPosInBuffer) < (buffer_len - buffer_read) Then
            If 0 < (lngBufferData - lngPosInBuffer) Then

                If 0 = IsBadReadPtr(ByVal VarPtr(btBuffer(0)) + lngPosInBuffer, lngBufferData - lngPosInBuffer) Then
                    If 0 = IsBadWritePtr(ByVal buffer_ptr + buffer_read, lngBufferData - lngPosInBuffer) Then
                        CpyMem ByVal buffer_ptr + buffer_read, ByVal VarPtr(btBuffer(0)) + lngPosInBuffer, lngBufferData - lngPosInBuffer
                    End If
                End If

                buffer_read = buffer_read + (lngBufferData - lngPosInBuffer)
            End If

            If Not FillBuffer Then
                StreamRead = SND_ERR_END_OF_STREAM
                Exit Function
            End If

        ' enough data in the PCM buffer
        Else
            If 0 = IsBadReadPtr(ByVal VarPtr(btBuffer(0)) + lngPosInBuffer, buffer_len - buffer_read) Then
                If 0 = IsBadWritePtr(ByVal buffer_ptr + buffer_read, buffer_len - buffer_read) Then
                    CpyMem ByVal buffer_ptr + buffer_read, ByVal VarPtr(btBuffer(0)) + lngPosInBuffer, buffer_len - buffer_read
                End If
            End If

            lngPosInBuffer = lngPosInBuffer + (buffer_len - buffer_read)
            buffer_read = buffer_read + (buffer_len - buffer_read)

        End If

    Loop
End Function

Public Function StreamSeek( _
    ByVal value As Long, _
    ByVal seek_mode As SND_SEEK_MODE _
) As SND_RESULT

    Dim curTime     As QWORD, curTime0  As QWORD
    Dim dblTime     As Double, strHex   As String
    Dim lngRet      As Long

    ' IWMSyncReader::SetRange wants QWords (64-bit integers).
    ' Since VB doesn't support these (actually there is Currency,
    ' but I lack a deeper understanding of that type), we need
    ' a way to calculate them on our own.
    ' Simple solution: Convert the integer part of a double
    ' to hexadecimals, cut that value into 2 DWords, which VB
    ' can handle. Thanks to Cristian Calugar for his Dbl2Hex function:
    ' http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=49987&lngWId=1

    Select Case seek_mode

        Case SND_SEEK_PERCENT
            If value < 0 Or value > 99 Then
                StreamSeek = SND_ERR_OUT_OF_RANGE
                Exit Function
            End If

            ' Position in 100-Nanoseconds
            dblTime = Fix(curDuration / CDbl(100) * CDbl(value) * 10000#)

        Case SND_SEEK_SECONDS
            If value < 0 Or value > (Duration / 1000) Then
                StreamSeek = SND_ERR_OUT_OF_RANGE
                Exit Function
            End If

            ' Position in 100-Nanoseconds
            dblTime = Fix(curDuration / CDbl(100) * CDbl(value / (Duration / 1000) * 100) * 10000#)

    End Select

    ' from Base 10 to Base 16
    strHex = FmtStrLen(DecToHex(dblTime), 16)

    ' lo DWORD
    curTime.lo = Val("&H" & Right$(strHex, 8) & "&")
    ' hi DWORD
    curTime.hi = Val("&H" & Left$(strHex, 8) & "&")

    ' set new range to decode
    lngRet = CallPointer(IReader.SetRange, oReader, _
                         curTime.lo, curTime.hi, _
                         curTime0.lo, curTime0.hi)

    If lngRet = 0 Then
        FillBuffer
        StreamSeek = SND_ERR_SUCCESS
    Else
        StreamSeek = SND_ERR_OUT_OF_RANGE
    End If
End Function

' WinAPI File Access

Private Function IsValidFile( _
    ByVal strFile As String _
) As Boolean

    Dim hInp    As hFile

    hInp = FileOpen(strFile, GENERIC_READ, FILE_SHARE_READ)
    IsValidFile = hInp.handle <> INVALID_HANDLE
    FileClose hInp
End Function

Private Function FileOpen( _
    ByVal strFile As String, _
    Optional access As FILE_ACCESS_RIGHTS = GENERIC_READ Or GENERIC_WRITE, _
    Optional share As FILE_SHARE_RIGHTS = FILE_SHARE_READ Or FILE_SHARE_WRITE, _
    Optional method As FILE_OPEN_METHOD = OPEN_EXISTING _
) As hFile

    FileOpen.handle = CreateFile(strFile, _
                                 access, _
                                 share, _
                                 ByVal 0&, _
                                 method, _
                                 0, 0)

    FileOpen.path = strFile
End Function

Private Sub FileClose( _
    filehandle As hFile _
)

    CloseHandle filehandle.handle
    filehandle.handle = INVALID_HANDLE
    filehandle.path = vbNullString
End Sub

Private Function FileRead( _
    filehandle As hFile, _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As Long

    Dim dwRead  As Long
    Dim lngRet  As Long

    If filehandle.handle = INVALID_HANDLE Then Exit Function

    lngRet = ReadFile(filehandle.handle, ByVal ptr, bytes, dwRead, 0&)

    If lngRet = 1 Then
        FileRead = dwRead
    Else
        FileRead = -1
    End If
End Function

Private Function FileWrite( _
    filehandle As hFile, _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As Long

    Dim dwWritten   As Long
    Dim lngRet      As Long

    If filehandle.handle = INVALID_HANDLE Then Exit Function

    lngRet = WriteFile(filehandle.handle, ByVal ptr, bytes, dwWritten, 0&)

    If lngRet = 1 Then
        FileWrite = dwWritten
    Else
        FileWrite = -1
    End If
End Function

Private Function FileSeek( _
    filehandle As hFile, _
    ByVal bytes As Long, _
    ByVal method As SEEK_METHOD _
) As Long

    FileSeek = SetFilePointer(filehandle.handle, bytes, 0, method)

End Function

Private Function FilePosition( _
    filehandle As hFile _
) As Long

    FilePosition = FileSeek(filehandle, 0, FILE_CURRENT)

End Function

Private Function FileLength( _
    filehandle As hFile _
) As Long

    FileLength = GetFileSize(filehandle.handle, 0)
End Function

Private Function FileEnd( _
    filehandle As hFile _
) As Boolean

    FileEnd = FilePosition(filehandle) >= FileLength(filehandle)
End Function
