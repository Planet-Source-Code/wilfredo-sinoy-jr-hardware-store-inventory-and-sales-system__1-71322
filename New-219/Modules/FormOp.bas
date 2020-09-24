Attribute VB_Name = "FormOp"
Option Explicit


'API declarations for dragging form
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

'variables for ADODB
Dim Connect As New ADODB.Connection


'Calling connection to Database
Public Sub ConDB()
    On Error Resume Next
    Connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False;Jet OLEDB:Database Password=noah "
End Sub


Public Function AppDir() As String
    If Right$(App.Path, 1) = "\" Then
        AppDir = App.Path
    Else
        AppDir = App.Path & "\"
    End If
End Function

'PROCEDURE TO CENTER A CHILD FORM ONTO A PARENT FORM
Public Sub CenterFrm(ByVal Parentfrm As MDIForm, ByVal Childfrm As Form) 'used for the frmInsignia

    Childfrm.Left = (Parentfrm.Width \ 2) - (Childfrm.Width \ 2)
    Childfrm.Top = (Parentfrm.ScaleHeight \ 2) - (Childfrm.Height \ 2)

End Sub

Public Sub ConnectToDb(adoObj As Adodc, AdoRec As String) 'for table Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdTable
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub ConnectToDb1(adoObj As Adodc, AdoRec As String) 'for table Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdTable
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub ConnectToDb2(adoObj As Adodc, AdoRec As String) 'for table Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdTable
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub ConnectToDb3(adoObj As Adodc, AdoRec As String) 'for table Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdTable
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub ConnectToDb4(adoObj As Adodc, AdoRec As String) 'for table Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdTable
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub ConnectToDb5(adoObj As Adodc, AdoRec As String) 'for table Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdTable
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub ConnectToDb6(adoObj As Adodc, AdoRec As String) 'for table Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdTable
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub SQLDB(adoObj As Adodc, AdoRec As String) 'for SQL Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdText
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub SQLDB1(adoObj As Adodc, AdoRec As String) 'for SQL Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdText
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub SQLDB2(adoObj As Adodc, AdoRec As String) 'for SQL Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdText
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub SQLDB3(adoObj As Adodc, AdoRec As String) 'for SQL Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdText
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub SQLDB4(adoObj As Adodc, AdoRec As String) 'for SQL Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdText
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub SQLDB5(adoObj As Adodc, AdoRec As String) 'for SQL Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdText
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub
Public Sub SQLDB6(adoObj As Adodc, AdoRec As String) 'for SQL Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\DB\DbaseHardware.mdb;Persist Security Info=False; Jet OLEDB:Database Password =noah"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdText
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub


