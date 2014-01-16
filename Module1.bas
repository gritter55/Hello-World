Attribute VB_Name = "Module1"
Dim appExcel As Excel.Application

Sub Main()
    
    frmStatus.Show
    If IsNewWeek() Then
        BackUpSheet
        UL "backed up workbook for new week"
    End If
    StartExcel
    UL "opened Excel"
    If IsNewWeek() Then
        CleanupSheet
        UL "cleaned up workbook for new week"
    End If
    UpdateLeonard
    UL "updated Adjustments.xls"
    ShutdownExcel
    UL "shutdown Excel"
    UL " finished Leonard.exe"
    DoEvents
    Unload frmStatus

End Sub

Function IsNewWeek() As Boolean
    
    If Format(Now(), "ddd") = "Sun" Then
        IsNewWeek = True
    Else
        IsNewWeek = False
    End If

End Function

Sub BackUpSheet()
    
        
    FileCopy Source:="G:\Accounting\Adjustments.xls", _
        Destination:="G:\Accounting\Adjustments for week ending " & Format(Now() - 1, "yymmdd") & ".xls"
     
End Sub

Sub CleanupSheet()

    frmStatus.lblStatus.Caption = "cleaning up workbook..."
    DoEvents
    appExcel.Application.Run ("Adjustments.xls!Cleanup")


End Sub
Sub StartExcel()

    frmStatus.lblStatus.Caption = "creating an instance of Excel..."
    DoEvents
    Set appExcel = New Excel.Application
    appExcel.Application.Visible = False
    frmStatus.lblStatus.Caption = "opening Leonard.xls..."
    DoEvents
    appExcel.Workbooks.Open ("G:\accounting\Adjustments.xls")
   
End Sub


Sub UpdateLeonard()
    
    frmStatus.lblStatus.Caption = "updating workbook..."
    DoEvents
    appExcel.Application.Run ("Adjustments.xls!Main")

End Sub


Sub ShutdownExcel()

    frmStatus.lblStatus.Caption = "done!  Shutdown Excel..."
    DoEvents
    appExcel.Workbooks.Close
    appExcel.Quit
    Set appExcel = Nothing

End Sub

