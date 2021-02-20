Option Explicit

'''**********************************************
''' ExportSlides2Video.vbs
''' Author: nalulabo
''' Copyright 2021- nalulabo
''' License: MIT
'''**********************************************
Dim pptx
Dim fso
Dim stdout
Const ppMediaTaskStatusDone = 3
Const ppMediaTypeSound = 2
Const msoTrue = -1
Const msoFalse = 0
Const msoMedia = 16
Const msoFileDialogFilePicker = 3
Const INFINITE = -1

Sub ExportSlides2Video(ppt)
    '''*****************************************
    ''' ExportSlides2Video
    ''' param: Presentation(Presentation-Object)
    '''*****************************************
    
    Dim outputDir: outputDir = fso.GetTempName()
    Dim sl
    Dim mp4name
    
    mp4name = GetVideoName(ppt.FullName)
    
    If Not fso.FileExists(outputDir) Then
        fso.CreateFolder(outputDir)
    End If
    
    For Each sl In ppt.Slides
        CreateSpeakNote sl.Name, GetNoteText(sl), outputDir
        EmbedNoteVoices sl, outputDir
        sl.SlideShowTransition.AdvanceTime = 10
    Next
    WScript.Echo "�ǂݏグ�ꎞ�t�@�C�����������Ă��܂�..."
    fso.DeleteFolder(outputDir)
    WScript.Echo "����̏����o�����J�n���܂�... ===> [ " & mp4name & " ]"
    ppt.CreateVideo mp4name, , , 1080, ,80
    
    do Until ppt.CreateVideoStatus = ppMediaTaskStatusDone
        WScript.Sleep 500
    Loop
    WScript.Echo "����̏����o�����������܂����B"
End Sub

Function GetNoteText(target)
    '''*****************************************
    ''' GetNoteText
    ''' param: Slide(Slide-Object)
    ''' return: note text(String)
    '''*****************************************

    GetNoteText = target.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text

End Function

Function GetVideoName(target)
    '''*****************************************
    ''' GetVideoName
    ''' param: pptx file path(String)
    ''' return: mp4 file path(String)
    ''' �t�@�C�������K��� "{powerpoint-file}_yyyy-MM-dd-hh-mm-ss.mp4"
    '''*****************************************
    
    Dim timestamp: timestamp = Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "-")
    GetVideoName = fso.BuildPath(fso.GetParentFolderName(target), fso.GetBaseName(target) & "_" & timestamp & ".mp4")
    
End Function

Function JoinPath(parent, child)
    '''*****************************************
    ''' JoinPath
    ''' param: parent folder path(String), file name(String)
    ''' return: file path(String)
    '''*****************************************

    JoinPath = fso.BuildPath(parent, child)

End Function

Sub CreateSpeakNote(Name, text, output)
    '''*****************************************
    ''' CreateSpeakNote
    ''' param: slide name(String), text to speak(String), file path(String)
    ''' requirement: SAPI
    ''' Voice: 411(Japanese)�A�����̐��݂̂ɑΉ�
    ''' todo: �ق��̉����i�j���̐��Ƃ��j�ɑΉ�����H
    '''*****************************************
    
    Const SAFT48kHz16BitStereo = 39
    Const SSFMCreateForWrite = 3
    
    Dim sapi: Set sapi = CreateObject("SAPI.SpVoice")
    Dim stream: Set stream = CreateObject("SAPI.SpFileStream")
    Dim outfile: outfile = JoinPath(output, Name)
    
    Set sapi.Voice = sapi.GetVoices("Language=411; Gender=Female")(0)
    
    stream.Format.Type = SAFT48kHz16BitStereo
    stream.Open outfile, SSFMCreateForWrite
    
    Set sapi.AudioOutputStream = stream
    sapi.Speak text
    WScript.Echo "�ǂݏグ�Ă��܂�..."
    sapi.WaitUntilDone(INFINITE)
    stream.Close
    
    Set stream = Nothing
    Set sapi = Nothing
    
End Sub


Sub EmbedNoteVoices(target, wavDir)
    '''*****************************************
    ''' EmbedNoteVoices
    ''' param: slide(Slide Object), file path(String)
    ''' remark: ���Ƃ���X���C�h�ɑ��݂��Ă��鉹�����f�B�A�͏�������܂�
    ''' todo: ���Ƃ��瑶�݂��Ă��鉹�����f�B�A���c���Ȃ���ǂݏグ���������ł��Ȃ���
    '''*****************************************
    Dim wavPath: wavPath = JoinPath(wavDir, target.Name)
    Dim sh
    
    For Each sh In target.Shapes
        If sh.Type = msoMedia Then
            If sh.MediaType = ppMediaTypeSound Then
                sh.Delete
            End If
        End If
    Next
    WScript.Echo "�ǂݏグ���ʂ̉����t�@�C���F[ " & wavPath & " ]"
    WScript.Sleep 1000
    WScript.Echo "�X���C�h�ɉ����𖄂ߍ���ł��܂�..."
    With target.Shapes.AddMediaObject2(wavPath, False, True, 10, 10).AnimationSettings.PlaySettings
        .PlayOnEntry = msoTrue
        .HideWhileNotPlaying = msoTrue
    End With
End Sub

Sub Main()
    Dim arg: Set arg = WScript.Arguments
    Dim target: target = ""
    Dim pp
    If arg.Count = 0 Then
        WScript.Echo "PowerPoint�h�L�������g�t�@�C�����w�肵�Ă��������B"
        Exit Sub
    End If
    target = arg.Item(0)
    WScript.Echo "�ϊ��t�@�C���F[ " & target & " ]"
    Set pp = pptx.Presentations.Open(target)
    ExportSlides2Video pp
    pp.Saved = True
    pp.Close
End Sub

Sub PreRequire()
    Set pptx = CreateObject("PowerPoint.Application")
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

Sub PostRequire()
    If pptx.Presentations.Count = 0 Then
        WScript.Echo "PowerPoint���I�����܂��B"
        pptx.Quit
    End If
    Set pptx = Nothing
End Sub

Function IsCscript()
    IsCscript = Instr(LCase(WScript.FullName), "cscript.exe") > 0
End Function

'''***
''' WScript.exe��PowerPoint�̃C���X�^���X�������ۂ����悤�Ȃ̂�
''' CScript.exe�݂̂Ɏ��s�𐧌����邱�Ƃɂ���
''' todo: �ق�Ƃ��ɋ��ۂ����̂��v�m�F
If IsCscript() Then
    Call PreRequire()
    Call Main()
    Call PostRequire()
Else
    WScript.Echo "�R�}���h���C��������s���Ă��������B" & VbCrLf & "cscript.exe " & WScript.ScriptName & " [PowerPoint�t�@�C��]"
End If
