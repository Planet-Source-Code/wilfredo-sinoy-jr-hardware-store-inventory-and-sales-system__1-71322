Attribute VB_Name = "SubMain"
Public Const HiLyt = "{HOME}+{END}"

Option Explicit
'initializations are located at this sub procedure

Public Sub Main()
'detects if application is already open
If App.PrevInstance = True Then
    MsgBox "System is already open.", vbOKOnly + vbInformation, "Inventory System"
    End
End If

'Main_On = False
frmSplash.Show
End Sub
