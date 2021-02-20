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
    WScript.Echo "読み上げ一時ファイルを除去しています..."
    fso.DeleteFolder(outputDir)
    WScript.Echo "動画の書き出しを開始します... ===> [ " & mp4name & " ]"
    ppt.CreateVideo mp4name, , , 1080, ,80
    
    do Until ppt.CreateVideoStatus = ppMediaTaskStatusDone
        WScript.Sleep 500
    Loop
    WScript.Echo "動画の書き出しを完了しました。"
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
    ''' ファイル命名規約は "{powerpoint-file}_yyyy-MM-dd-hh-mm-ss.mp4"
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
    ''' Voice: 411(Japanese)、女性の声のみに対応
    ''' todo: ほかの音声（男性の声とか）に対応する？
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
    WScript.Echo "読み上げています..."
    sapi.WaitUntilDone(INFINITE)
    stream.Close
    
    Set stream = Nothing
    Set sapi = Nothing
    
End Sub


Sub EmbedNoteVoices(target, wavDir)
    '''*****************************************
    ''' EmbedNoteVoices
    ''' param: slide(Slide Object), file path(String)
    ''' remark: もとからスライドに存在している音声メディアは除去されます
    ''' todo: もとから存在している音声メディアを残しながら読み上げだけ除去できないか
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
    WScript.Echo "読み上げ結果の音声ファイル：[ " & wavPath & " ]"
    WScript.Sleep 1000
    WScript.Echo "スライドに音声を埋め込んでいます..."
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
        WScript.Echo "PowerPointドキュメントファイルを指定してください。"
        Exit Sub
    End If
    target = arg.Item(0)
    WScript.Echo "変換ファイル：[ " & target & " ]"
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
        WScript.Echo "PowerPointを終了します。"
        pptx.Quit
    End If
    Set pptx = Nothing
End Sub

Function IsCscript()
    IsCscript = Instr(LCase(WScript.FullName), "cscript.exe") > 0
End Function

'''***
''' WScript.exeでPowerPointのインスタンス化が拒否されるようなので
''' CScript.exeのみに実行を制限することにした
''' todo: ほんとうに拒否されるのか要確認
If IsCscript() Then
    Call PreRequire()
    Call Main()
    Call PostRequire()
Else
    WScript.Echo "コマンドラインから実行してください。" & VbCrLf & "cscript.exe " & WScript.ScriptName & " [PowerPointファイル]"
End If
