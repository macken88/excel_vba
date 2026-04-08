Attribute VB_Name = "BookA_TestMacros"
Option Explicit

Public Sub WriteTodayToActiveCell()
    ActiveCell.Value = Date
    ActiveCell.NumberFormatLocal = "yyyy/mm/dd"
End Sub

Public Sub ShowBookAGreeting()
    MsgBox "BookA test macro is ready.", vbInformation
End Sub
