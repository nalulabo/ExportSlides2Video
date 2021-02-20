Attribute VB_Name = "ExportSlides2Video"
Option Explicit
'''**********************************************
''' ExportSlides2Video.vbs
''' Author: nalulabo
''' Copyright 2021- nalulabo
''' License: MIT
'''**********************************************

Private fso As Object

Public Sub ExportSlides2Video()
    Dim outputDir As String: outputDir = GetTempName
    Dim sl As Slide
    Dim mp4name As String
    
    mp4name = GetVideoName(ActivePresentation.FullName)
    
    If Dir(outputDir) = "" Then
        MkDir outputDir
    End If
    
    For Each sl In ActivePresentation.Slides
        CreateSpeakNote sl.Name, GetNoteText(sl), outputDir
        EmbedNoteVoices sl, outputDir
        sl.SlideShowTransition.AdvanceTime = 10
    Next
    ActivePresentation.CreateVideo FileName:=mp4name, VertResolution:=1080, Quality:=80
    
    DeleteTempFolder outputDir
    
    MsgBox "動画へのエクスポートが開始されました。" & vbCrLf & _
           "動画の書き出しが完了するまでお待ちください。", vbInformation, "ExportSlides2Video"
    
End Sub

Private Sub Initialize()

    Set fso = CreateObject("Scripting.FileSystemObject")

End Sub

Private Function GetNoteText(target As Slide)

    GetNoteText = target.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text

End Function

Private Function GetVideoName(target As String) As String
    
    Dim timestamp As String: timestamp = Format(Now(), "yyyy-MM-dd-hh-mm-ss")
    
    GetVideoName = fso.BuildPath(fso.GetParentFolderName(target), fso.GetBaseName(target) & "_" & timestamp & ".mp4")
    Set fso = Nothing
    
End Function

Private Function GetTempName() As String
    
    GetTempName = fso.GetTempName()
    Set fso = Nothing
    
End Function

Private Sub DeleteTempFolder(target As String)

    fso.DeleteFolder (target)
    Set fso = Nothing
    
End Sub

Private Function JoinPath(parent As String, child As String) As String

    JoinPath = fso.BuildPath(parent, child)
    Set fso = Nothing

End Function

Private Sub CreateSpeakNote(Name As String, text As String, output As String)
    
    Const SAFT48kHz16BitStereo = 39
    Const SSFMCreateForWrite = 3
    
    Dim sapi As Object: Set sapi = CreateObject("SAPI.SpVoice")
    Dim stream As Object: Set stream = CreateObject("SAPI.SpFileStream")
    Dim outfile As String: outfile = JoinPath(output, Name)
    
    Set sapi.Voice = sapi.GetVoices("Language=411; Gender=Female")(0)
    
    stream.Format.Type = SAFT48kHz16BitStereo
    stream.Open outfile, SSFMCreateForWrite
    
    Set sapi.AudioOutputStream = stream
    sapi.Speak text
    stream.Close
    
    Set stream = Nothing
    Set sapi = Nothing
    
End Sub


Private Sub EmbedNoteVoices(target As Slide, wavDir As String)
    Dim wavPath As String: wavPath = JoinPath(wavDir, target.Name)
    Dim sh As Shape
    
    For Each sh In target.Shapes
        If sh.Type = msoMedia Then
            If sh.MediaType = ppMediaTypeSound Then
                sh.Delete
            End If
        End If
    Next
    
    With target.Shapes.AddMediaObject2(wavPath, False, True, 10, 10).AnimationSettings.PlaySettings
        .PlayOnEntry = msoTrue
        .HideWhileNotPlaying = msoTrue
    End With
End Sub
