Attribute VB_Name = "log"
Option Explicit

'***** Set the APP constant to the name of the application it is included in
Private Const APP As String = "Leonard.exe"


Sub UL(msg As String)
    Dim intLogNum As Integer
    Dim LogPath As String
    
    LogPath = "P:\Downloads\morning logs\MC" & Format(Now(), "yymmdd") & ".log"
    intLogNum = FreeFile
    Open LogPath For Append Access Write As intLogNum
    Write #intLogNum, APP, Format(Now(), "general date"), msg
    Close intLogNum
    
End Sub
