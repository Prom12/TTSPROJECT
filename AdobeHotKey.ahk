; Start recording in Adobe Audition
^r::  ; Press Ctrl + R to start recording
{
    Run, "C:\Program Files\Adobe\Adobe Audition 2024\Adobe Audition.exe"
    WinWait, Adobe Audition  ; Wait for Adobe Audition window to appear
    ControlSend, , {Ctrl down}r{Ctrl up}, Adobe Audition  ; Send Ctrl + R to start recording
}

; Stop recording in Adobe Audition
^s::  ; Press Ctrl + S to stop recording
{
    ControlSend, , {Ctrl down}s{Ctrl up}, Adobe Audition  ; Send Ctrl + S to stop recording
    WinClose, Adobe Audition  ; Close Adobe Audition window
}