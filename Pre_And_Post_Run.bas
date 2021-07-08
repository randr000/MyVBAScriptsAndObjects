Attribute VB_Name = "Pre_And_Post_Run"
Sub Pre_Run()
'
' Run before main macro runs to speed things up and hide alerts
'
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
End Sub

Sub Post_Run()
'
' Run after main macro finishes running to reverse whe Pre_Run did
'
    Application.DisplayAlerts = True
    Application.Calculation = x1CalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub
