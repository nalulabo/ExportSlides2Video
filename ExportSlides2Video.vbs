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
Const msoFileDialogSaveAs = 2
Const INFINITE = -1
Const IDENTITY_NAME = "ExportSlides2Video"
Const WAV_EXT = ".wav"

Function GetWavLength(outName)
    Dim sh: Set sh = CreateObject("Shell.Application")
    Dim outdir: outdir = fso.GetParentFolderName(outName)
    Dim name: name = fso.GetFileName(outName)
    Dim ns: Set ns = sh.Namespace(outdir)
    Dim f: f = ns.ParseName(name)
    Dim timeLength: timeLength = ns.GetDetailsOf(f, 27)
    Dim times: times = Split(timeLength, ":")
    If Ubound(times) + 1 = 3 Then
        GetWavLength = CInt(times(0)) * 3600 + CInt(times(1)) * 60 + CInt(times(2))
    Else
        GetWavLength = 10
    End If
End Function

Sub TreatSlide(slide, outputDir)
    Dim note: note = GetNoteText(slide)
    If Trim(note) <> "" Then
        Dim outName: outName = CreateSpeakNote(slide.Name, note, outputDir)
        Dim timeLength: timeLength = GetWavLength(outName)
        With slide.SlideShowTransition
            .AdvanceOnTime = True
            .AdvanceTime = timeLength
        End With
        EmbedNoteVoices slide, outputDir
    End If
End Sub

Sub Export(ppt, exportName, noVideo)
    If noVideo Then
        WriteHost "動画は出力しません。"
        WScript.Echo "PowerPointファイルに音声を埋め込みました。"
    Else
        WriteHost "動画の書き出しを開始します... ===> [ " & exportName & " ]"
        ppt.CreateVideo exportName, True, , 1080, ,80
        
        do Until ppt.CreateVideoStatus = ppMediaTaskStatusDone
            WScript.Sleep 500
        Loop
        WScript.Echo "動画の書き出しを完了しました。"
    End If
End Sub

Function ExportSlides2Video(ppt, noVideo)
    '''*****************************************
    ''' ExportSlides2Video
    ''' param: Presentation(Presentation-Object)
    ''' param: dont export video (boolean)
    '''*****************************************
    
    Dim outputDir: outputDir = fso.GetAbsolutePathName(fso.GetTempName())
    Dim sl
    Dim exportName
    
    exportName = GetExportName(ppt.FullName, noVideo)
    
    If Not fso.FileExists(outputDir) Then
        fso.CreateFolder(outputDir)
    End If
    
    For Each sl In ppt.Slides
        TreatSlide sl, outputDir
    Next
    WriteHost "読み上げ一時ファイルを除去しています..."
    fso.DeleteFolder(outputDir)
    Export ppt, exportName, noVideo
    ExportSlides2Video = exportName
End Function

Function GetNoteText(target)
    '''*****************************************
    ''' GetNoteText
    ''' param: Slide(Slide-Object)
    ''' return: note text(String)
    '''*****************************************

    GetNoteText = target.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text

End Function

Function IIf(condition, trueCase, falseCase)
    If condition Then
        IIf = trueCase
    Else
        IIf = falseCase
    End If
End Function

Function GetExportName(target, isPptx)
    '''*****************************************
    ''' GetExportName
    ''' param: file path(String)
    ''' param: is pptx ?(boolean)
    ''' return: rule based file path(String)
    ''' ファイル命名規約は "{powerpoint-file}_yyyy-MM-dd-hh-mm-ss.{extention}"
    '''*****************************************
    
    Dim timestamp: timestamp = Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "-")
    Dim ext: ext = IIf(isPptx, ".pptx", ".mp4")
    GetExportName = fso.BuildPath(fso.GetParentFolderName(target), fso.GetBaseName(target) & "_" & timestamp & ext)
    
End Function

Function JoinPath(parent, child)
    '''*****************************************
    ''' JoinPath
    ''' param: parent folder path(String), file name(String)
    ''' return: file path(String)
    '''*****************************************

    JoinPath = fso.BuildPath(parent, child)

End Function

Function CreateSpeakNote(Name, text, output)
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
    Dim outfile: outfile = JoinPath(output, Name & WAV_EXT)
    
    Set sapi.Voice = sapi.GetVoices("Language=411; Gender=Female")(0)
    
    stream.Format.Type = SAFT48kHz16BitStereo
    stream.Open outfile, SSFMCreateForWrite
    
    Set sapi.AudioOutputStream = stream
    sapi.Speak text
    WriteHost "読み上げています..."
    sapi.WaitUntilDone(INFINITE)
    stream.Close
    
    Set stream = Nothing
    Set sapi = Nothing
    CreateSpeakNote = outfile    
End Function


Sub EmbedNoteVoices(target, wavDir)
    '''*****************************************
    ''' EmbedNoteVoices
    ''' param: slide(Slide Object), file path(String)
    ''' 指定された代替テキストの音声メディアオブジェクトを除去して埋め込みます
    '''*****************************************
    Dim wavPath: wavPath = JoinPath(wavDir, target.Name & WAV_EXT)
    Dim sh, wav
    
    For Each sh In target.Shapes
        If sh.Type = msoMedia Then
            If sh.MediaType = ppMediaTypeSound And sh.AlternativeText = IDENTITY_NAME Then
                sh.Delete
            End If
        End If
    Next
    WriteHost "読み上げ結果の音声ファイル：[ " & wavPath & " ]"
    WScript.Sleep 1000
    WriteHost "スライドに音声を埋め込んでいます..."
    Set wav = target.Shapes.AddMediaObject2(wavPath, False, True, 10, 10)
    With wav.AnimationSettings.PlaySettings
        .PlayOnEntry = True
        .HideWhileNotPlaying = True
    End With
    wav.AlternativeText = IDENTITY_NAME
End Sub

Sub Main()
    Dim arg: Set arg = WScript.Arguments
    Dim target: target = ""
    Dim pp
    Dim noVideo: noVideo = False
    Dim exportedName
    Const askVideo = "動画書き出しを行います。音声埋め込みまでにとどめたい場合は「いいえ」をクリックしてください。"
    If arg.Count = 0 Then
        WScript.Echo "PowerPointドキュメントファイルを指定してください。"
        pptx.Activate()
        With pptx.FileDialog(msoFileDialogFilePicker)
            .Filters.Add "*.pptx", "*.pptx"
            .ButtonName = "変換する"
            .InitialFileName = JoinPath(fso.GetParentFolderName(WScript.ScriptFullName), "presentation.pptx")
            .Title = "マークダウンからPowerPointを作成して動画に書き出します"
            If .Show Then
                target = .SelectedItems(1)
            Else
                Exit Sub
            End If
        End With
        If MsgBox(askVideo, vbYesNo, IDENTITY_NAME) = vbNo Then
            noVideo = True
        End If
    Else
        target = fso.GetAbsolutePathName(arg.Item(0))
    End If
    WriteHost "変換ファイル：[ " & target & " ]"
    Set pp = pptx.Presentations.Open(target)
    pp.AutoSaveOn = False
    exportedName = ExportSlides2Video(pp, noVideo)
    If noVideo Then
        With pptx.FileDialog(msoFileDialogSaveAs)
            .ButtonName = "保存する"
            .InitialFileName = JoinPath(fso.GetParentFolderName(WScript.ScriptFullName), exportedName)
            .Title = "保存先を指定してください"
            If .Show Then
                pp.SaveAs(.SelectedItems(1))
            Else
                pp.Saved = True
            End If
        End With
    Else
        pp.Saved = True
    End IF
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

Sub WriteHost(message)
    If isCui Then
        WScript.Echo message
    End If
End Sub

Function IsCscript()
    IsCscript = Instr(LCase(WScript.FullName), "cscript.exe") > 0
End Function

Dim isCui: isCui = IsCscript()
Call PreRequire()
Call Main()
Call PostRequire()
