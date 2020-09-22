Attribute VB_Name = "modDatabase"
Option Explicit

Public DB_Path As String
Public DB_Connection As ADODB.Connection

Public Sub InitializeDatabase()
    On Error GoTo ErrInitializeDatabase

    'Set new database connection
    Set DB_Connection = New ADODB.Connection
    'Set database path
    DB_Path = App.Path & "\AIMServer.mdb"
    'Set are database driver
    DB_Connection.Provider = "Microsoft.Jet.OLEDB.4.0"
    'Set are cursor location
    DB_Connection.CursorLocation = adUseClient
    'Set database R/W mode
    DB_Connection.Mode = adModeReadWrite
    'Open the database connection
    DB_Connection.Open DB_Path, "Admin"

    On Error GoTo 0
    Exit Sub

ErrInitializeDatabase:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure InitializeDatabase of modDatabase"

End Sub

Public Sub TerminateDatabase()
    On Error GoTo ErrTerminateDatabase

    'Close database
    DB_Connection.Close
    'Kill the object
    Set DB_Connection = Nothing

    On Error GoTo 0
    Exit Sub

ErrTerminateDatabase:

    ErrMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure TerminateDatabase of modDatabase"

End Sub
