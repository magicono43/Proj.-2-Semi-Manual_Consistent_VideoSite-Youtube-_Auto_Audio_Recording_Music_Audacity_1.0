/*
"Dependent Requirements Before Running"
1. https://autohotkey.com/board/topic/21984-vista-audio-control-functions/ V2.3 (Imported Library)
2. Made and designed with Windows 7 OS
3. Internet Explorer Browser (With Ad-blocking addon for optimal performance)
4. Youtube pause and resume video short-cut being "k"
5. Youtube volume being set to 100%
6. Audacity "Export as WAV" short-cut being "Shift+W"
7. Audacity "Close (File)" short-cut being "Ctrl+W"
8. Audacity "Exit (Window) short-cut being "Shift+L"
9. Whatever other files I link to this for next steps of file editing and such.....### (#1: Audacity_Auto_Silence_Removal_From_WAV_Files.ahk), 
*/
#SingleInstance,Force
#Include C:\Users\KirkO\Desktop\Coding Repository\AutoHotKey\AHK_Library_Files\Vista_Audio_Control_Functions_V2.3.ahk

/*
Invalid_File_Text := ["*","""","/","\","<",">",":","|","?"]
qt = /***??<?dust\|gr>?"av"el:21
MsgBox % qt
 if qt contains *,`"`,/,\,<,>,:,|,? ;-- Used to detect if title of song has an illegal character for a file-name, then removes those illegal characters and allows changed version to be used later.
    {
        for each, char in Invalid_File_Text
            qt := StrReplace(qt, char, "")
    }
MsgBox % qt
Exit
*/
/*      ;-- Only using for audio-level testing/trouble-shooting.
audioMeter := VA_GetAudioMeter()
Loop ;-- Too keep start of program from instantly saving recording due to moment of silence at start.
{
    peakLevel := Peakvalue_Monitor(audioMeter)
    ToolTip, %peakLevel%
        }
*/
/*      ;-- Only using for title testing/trouble shooting.
wb := ComObjCreate("InternetExplorer.Application")
wb.Visible := True
wb.Navigate("https://www.youtube.com/watch?v=hvXsl27-B3k&list=PLE16D748F63ACBB97&index=4") ;-- First video URL I want to start at for play-list.
while (wb.readystate != 4)
    sleep 10
Sleep, 500
qt := Video_YT_Pause(wb)
MsgBox % qt
Exit
*/
Run, "C:\Program Files (x86)\Audacity\audacity.exe"
Sleep, 6000
wb := ComObjCreate("InternetExplorer.Application")
global Record_Count := 0 ;- Meant to prevent duplicate names from messing up script running.
wb.Visible := True
wb.Navigate("https://www.youtube.com/watch?v=z70n7S_yKRw&list=PLE16D748F63ACBB97&index=1") ;-- First video URL I want to start at for play-list.
while (wb.readystate != 4)
    sleep 10
WinActivate, ahk_class wxWindowNR
ControlClick, wxWindowNR10, ahk_class wxWindowNR,,,, NA
Sleep, 50

SetControlDelay -1

audioMeter := VA_GetAudioMeter()

Loop ;-- Too keep start of program from instantly saving recording due to moment of silence at start.
{
    peakLevel := Peakvalue_Monitor(audioMeter)
    if (peakLevel >= 0.005) {
        break
        }
}

RestartPoint:
Loop
{   ;- Audio Monitoring Loop, Main loop. Note, don't run program without Audacity Open.
peakLevel := Peakvalue_Monitor(audioMeter)

if (peakLevel <= 0.003) {
        Loop {
            Sleep, 80
            peakLevel := Peakvalue_Monitor(audioMeter)
            if (A_Index >= 5)
                break
            if (peakLevel <= 0.003)
                continue
            else
                goto, RestartPoint
        }
        
        qt := Video_YT_Pause(wb)
        sleep, 300
        Func_Name(qt)
        Func_Name2()
        Sleep, 200
        Video_YT_Unpause()
        
        Loop {      ;- After video is unpaused, waits for audio to peak above a value again before restarting loop, otherwise few seconds on silence causes script to keep trying to export file during small silence.
            Sleep, 80
            peakLevel := Peakvalue_Monitor(audioMeter)
            if (peakLevel >= 0.005) 
                goto, RestartPoint
            if (A_Index >= 1500) ;--- If there is a long enough period of silence, End_of_Playlist function will run and either close programs and put PC to sleep, or run next scripts for new audio files.
                End_of_Playlist() ;- Takes Approx. 4 minutes of silence before this Ending function is called.
        }
        ;--- I may also have this "Ending" open another script that will take all the new WAV files and automatically reduce their audio levels and crop dead-silence and possibly even convert the files into the format I can use on my phone, that's another few small projects after this one though, polish, this should just be to end this program and when to continue looping and such.
    }
}




; +-------------------------------------------Functions---------------------------------------------------+

Peakvalue_Monitor(audioMeter) ;- Function Call for monitoring the audio level, returns peakLevel var which can be used in if-statements and such, this is audio level as a float.
{
    VA_IAudioMeterInformation_GetPeakValue(audioMeter, peakValue)

    ;ToolTip, %peakValue% ;- For Testing/Debug purposes only.
    Sleep, 50 ;- Will Adjust Later On, just like many other values, including ending current project if-statement peakValue threshold parameter.
    return peakValue
}


Video_YT_Pause(wb) ;-- Function that will pause Youtube video at time to allow Audacity and videos to stay in sync to prevent missed audio.
{
    WinActivate, ahk_class IEFrame
    sleep, 200
    title := wb.Document.getElementById("eow-title").getAttribute("title")
    SendInput {k} ;-- Apparently Youtube has the "k" key as a short-cut for pausing and resuming videos playing, that might make this a bit easier, potentially.
    Sleep, 200
    return title
}


Video_YT_Unpause()
{
    WinActivate, ahk_class IEFrame
    sleep, 400
    SendInput {k}
    Sleep, 100
}


Get_Video_Title(wb) ;- Function for getting Title of Youtube video to use as exported file name. NOT CURRENTLY IN USE.
{
    /*
    rq := comobjcreate("WinHttp.WinHttpRequest.5.1")
    doc := comobjcreate("HTMLfile")
    rq.open("GET", "https://www.youtube.com/watch?v=W7OGaZyjSY8&list=PLuQj3usCmdXhZGNiBu6nC8OfsfK32ODOk&index=5&t=0s")
    rq.send()
    doc.write(rq.responsetext)
    doc.close()
    while doc.readystate == "loading"
    sleep 10

    msgbox % doc.title              ;-- Got this from person on Discord, this gets the title of a web-page without opening any windows, basically webscrapping with just server requests and responses, no GUI.
    Exit
    */
}


End_of_Playlist()
{
    ;#########  Only use this if I want everything to close and PC to go into Hibernate mode after. Otherwise I will eventually have this function close this script and run next ones for audio file editing, etc.
    ;;;;
    WinActivate, ahk_class IEFrame
    Sleep, 2000
    SendInput {Alt down}            ;--- Closes Active IE Window
    SendInput {F4}+{Alt up}
    Sleep, 2000
    ;;;;
    ;;;;;;
    WinActivate, ahk_class wxWindowNR
    Sleep, 2000
    ControlClick, wxWindowNR7, ahk_class wxWindowNR,,,, NA
    Sleep, 1500
    SendInput {Control down}
    SendInput {w}+{Control up}
    Sleep, 1000
    SendInput {NumpadRight}
    Sleep, 200
    SendInput {Enter}
    
    ;Sleep, 400                      ;--- Closes Current Audacity Project (Empty One) and then Closes Audacity Itself.
    ;WinActivate, ahk_class wxWindowNR
    ;Sleep, 1000
    ;SendInput {Shift down}
    ;SendInput {l}+{Shift up} ; Closes Audacity Window Custom Short-Cut
    
    Sleep, 1000
    ;;;;;;
    ;;
    ;MsgBox "Putting Computer To Sleep Testing, Program Ending"
    Sleep, 2000                     ;--- Exits Script and then puts whole computer into Hibernate/Sleep mode.
    ;DllCall("PowrProf\SetSuspendState", "int", 1, "int", 0, "int", 0)       ;- Puts the PC into Hibernate Mode, keep only if I want PC to Hibernate after script is done, like if I run while sleeping.
    ;Exit ;-@@@@@@@---@@@@-- Try on larger play-list and if that is good as well, start making the scripts that would do the play-list clearing (not that important). Also, finalize all scripts into just normal AHK compilable scripts, and put them in respective file locations and back-up everything, also make notes and have documentation for later use in libraries and other reference materials, then finally start on my next project when this is all done.                --  7/7/2019, 3:10 PM  --     @@@@@@---@@@@@###@@@@---  START HERE NEXT SESSION!!!
    ;;
    ;##########
    Run, Audacity_Auto_Silence_Removal_From_WAV_Files.ahk ;- This will run the next "step" of the process being the next script in line and then exit this current app after.
    ExitApp
}


Func_Name(qt)
{
    global Record_Count ++ ;- Adds 1 to this variable everytime this function is ran, this will prevent duplicate names for files.
    Invalid_File_Text := ["*","""","/","\","<",">",":","|","?"] ;- Strange with single double-quotes, need to use 4 quotes to make it a string, bottom table, https://www.autohotkey.com/docs/commands/_EscapeChar.htm 
    WinActivate, ahk_class wxWindowNR
    ControlClick, wxWindowNR7, ahk_class wxWindowNR,,,, NA 
    Sleep, 100
    WinActivate, ahk_class wxWindowNR
    Sleep, 300
    SendInput {Shift down}+{w}+{Shift up} ; Export WAV Custom Short-Cut
    Sleep, 800 ; Will remove sleeps where possible later, keep if it increases reliability and reduces errors.
    if (qt = "") {   ;- This is simply there to give a "unique" name to files that don't get a title from the web-page or it comes out invalid and gives back a blank.
        Random, rand, 1, 1000
        rand += Record_Count
        qt = INVALID_TITLE_%rand%
        }
    if qt contains *,`"`,/,\,<,>,:,|,? ;-- Used to detect if title of song has an illegal character for a file-name, then removes those illegal characters and allows changed version to be used later.
    {
        for each, char in Invalid_File_Text
            qt := StrReplace(qt, char, "")
    }
    SendInput %Record_Count%+{NumpadDot}
    Sleep, 300
    SendInput %qt% ;-- Name of Youtube Video Title of current loaded page.
    Sleep, 200
    SendInput {NumpadDot}+{w}+{a}+{v} ;-- Writes in file format extension after video title is put in.
    Sleep, 500
    SendInput {Enter} ; Presses "Enter" key to confirm file-name and location for exported WAV file.
    Sleep, 1000
    SendInput {Enter}
    Sleep, 12000 ;- This line will be some sort of wait command or possibly just a simple "Sleep, 8000" or something, depends if I can get videos to pause correctly on Youtube and such.
    WinActivate, ahk_class wxWindowNR
    Sleep, 200
}


Func_Name2()
{
    WinActivate, ahk_class wxWindowNR
    Sleep, 50
    SendInput {Control down} ;-- Really, when I look at it, the only way this could really run while doing other things would be to have "WinActive command" every line, which is fair enough if I want to, Messy though.
    SendInput {w}+{Control up}
    Sleep, 700
    SendInput {NumpadRight}
    Sleep, 50
    SendInput {Enter}
    Sleep, 500
    WinActivate, ahk_class wxWindowNR
    Sleep, 300
    ControlClick, wxWindowNR10, ahk_class wxWindowNR,,,, NA ;--- "wxWindowNR10" is the Record button address, can't find with window-spy for some reason.
    Sleep, 300
}


;x::msgbox % wb.locationurl