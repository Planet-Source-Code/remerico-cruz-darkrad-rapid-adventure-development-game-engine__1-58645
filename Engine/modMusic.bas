Attribute VB_Name = "modMusic"
' Music Engine
' This is how the music is played in the game...

' DirectMusic variables for MIDI files!
Dim MID_dx As New DirectX7
Dim MID_dmp As DirectMusicPerformance
Dim MID_Loader As DirectMusicLoader
Dim MID_Segment As DirectMusicSegment
Dim MID_SegmentSt As DirectMusicSegmentState

Sub PlayMIDI(wFile As String)
On Error Resume Next

    Set MID_Segment = MID_Loader.LoadSegment(wFile)
    
    If StrConv(Right(wFile, 4), vbLowerCase) = ".mid" Then MID_Segment.SetStandardMidiFile
    
    MID_dmp.SetMasterAutoDownload True
    MID_Segment.Download MID_dmp
    MID_Segment.SetRepeats 50
     
    Set MID_SegmentSt = MID_dmp.PlaySegment(MID_Segment, 0, 0)
End Sub

Sub InitMIDI()
  Set MID_Loader = MID_dx.DirectMusicLoaderCreate
  Set MID_dmp = MID_dx.DirectMusicPerformanceCreate
    
  MID_dmp.Init Nothing, hwnd
  MID_dmp.SetPort -1, 1
End Sub

Sub StopMIDI()
  If MID_Segment Is Nothing Then Exit Sub

  MID_dmp.Stop MID_Segment, MID_SegmentSt, 0, 0
  MID_Segment.Unload MID_dmp
End Sub

Sub MIDIVolume(wVol As Long)
  MID_dmp.SetMasterVolume wVol
End Sub

Sub StopMusic()
  Select Case LCase$(Right$(GMusic, 3))
    Case "mid"
      StopMIDI
  End Select
End Sub
