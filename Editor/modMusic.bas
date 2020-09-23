Attribute VB_Name = "modMusic"

' DirectMusic variables for MIDI files!
Dim M_dx As New DirectX7
Dim MID_dmp As DirectMusicPerformance
Dim MID_Loader As DirectMusicLoader
Dim MID_Segment As DirectMusicSegment
Dim MID_SegmentSt As DirectPlaySessionData

Dim M_DS As DirectSound
Dim M_dsBuffer As DirectSoundBuffer

Sub PlayMIDI(wFile As String)
On Error Resume Next

    Set MID_Segment = MID_Loader.LoadSegment(wFile)
    
    If StrConv(Right(wFile, 4), vbLowerCase) = ".mid" Then MID_Segment.SetStandardMidiFile
    
    MID_dmp.SetMasterAutoDownload True
    MID_Segment.Download MID_dmp
     
    Set MID_SegmentSt = MID_dmp.PlaySegment(MID_Segment, 0, 0)
End Sub

Sub InitSound()
  Set MID_Loader = M_dx.DirectMusicLoaderCreate
  Set MID_dmp = M_dx.DirectMusicPerformanceCreate
    
  MID_dmp.Init Nothing, hwnd
  MID_dmp.SetPort -1, 1
  
  ' Supposed to be the DSOUND support, but not now....
  'Set M_DS = M_dx.DirectSoundCreate("")
  'M_DS.SetCooperativeLevel frmMain.hwnd, DSSCL_PRIORITY
End Sub

Sub StopMIDI()
  If MID_Segment Is Nothing Then Exit Sub

  MID_dmp.Stop MID_Segment, MID_SegmentSt, 0, 0
  MID_Segment.Unload MID_dmp
End Sub

Sub MIDIVolume(wVol As Long)
  MID_dmp.SetMasterVolume wVol
End Sub


