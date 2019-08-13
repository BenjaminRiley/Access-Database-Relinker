Option Explicit
Option Compare Database

'Automatic Database Relinker v1.0.0
'https://github.com/BenjaminRiley/Access-Database-Relinker

Public Enum BackendType
    NotLinked = 0
    DevBE
    LocalBE
    ServerBE
End Enum

Private pLinkedBackendType As BackendType
'Stores the type of database linked at runtime.
Public Property Get LinkedBackendType() As BackendType
    LinkedBackendType = pLinkedBackendType
End Property


Public Function RelinkTables(ReferenceTableName As String, Optional ServerBackendPath As String) As BackendType
'DESCRIPTION:
'   Function to relink tables on startup. The goal is to allow usage of a shared, split database away from the network
'       and development with a local backend file without having to manually relink the backend every time a change is
'       made.
'   It is assumed that the local copy of the backed will be stored in the same folder as the frontend with _be appended
'       to the filename.
'   Copies of the backend for development purposes should have _be_dev appeneded to the filename.
'   i.e. for a project MyDatabase.accdb:
'       - local backend would be MyDatabase_be.accdb
'       - dev backend would be MyDatabase_be_dev.accdb
'PARAMETERS:
'   ReferenceTableName - The name of a linked table to get the current backend path from.
'   ServerBackendPath - Path to backend stored on the server.
'RETURNS:
'   Sets LinkedBackend and returns it
    
    'Initialise return value
    pLinkedBackendType = NotLinked
    RelinkTables = pLinkedBackendType
    
    Dim BackendPath As String
    Dim CurrentBackendPath As String
    Dim FoundBackendType As BackendType

    CurrentBackendPath = CurrentDb.TableDefs(ReferenceTableName).connect
    FoundBackendType = FindBackend(ServerBackendPath, BackendPath)
    
    If FoundBackendType And Len(BackendPath) > 0 Then
        'Prepend start of connection string to found path
        BackendPath = ";DATABASE=" & BackendPath

        'Skip relink if backend path hasn't changed.
        If StrComp(CurrentBackendPath, BackendPath) <> 0 Then
            'Relink the tables
            Dim LinkedTable As TableDef
            For Each LinkedTable In CurrentDb.TableDefs
                If LinkedTable.connect = CurrentBackendPath Then
                    LinkedTable.connect = BackendPath
                    LinkedTable.RefreshLink
                End If
            Next
        End If
        
        'Set return value
        pLinkedBackendType = FoundBackendType
        RelinkTables = pLinkedBackendType
    End If
End Function

Private Function FindBackend(ServerBackendPath As String, ByRef BackendPath) As BackendType
'DESCRIPTION:
'   Determines which path to use.
'PARAMETERS:
'   ServerBackendPath - The path to the backend on the server.
'   BackendPath - Output variable for the located path.
'RETURNS:
'   BackendType enum corresponding to the location of the found backend.

    'Initialise this to a default here so we don't forget later
    FindBackend = NotLinked

    Dim DevBackendPath As String
    Dim LocalBackendPath As String
    'ServerBackendPath is a parameter
    Dim FrontendPath As String
    Dim FrontendExtension As String
    
    FrontendPath = Left(CurrentProject.FullName, InStrRev(CurrentProject.FullName, ".") - 1)
    FrontendExtension = Right(CurrentProject.FullName, Len(CurrentProject.FullName) - InStrRev(CurrentProject.FullName, "."))
    
    DevBackendPath = FrontendPath & "_be_dev." & FrontendExtension
    LocalBackendPath = FrontendPath & "_be." & FrontendExtension
    
    If Len(Dir(DevBackendPath)) > 0 Then
        BackendPath = DevBackendPath
        FindBackend = DevBE
    ElseIf Len(Dir(LocalBackendPath)) > 0 Then
        BackendPath = LocalBackendPath
        FindBackend = LocalBE
    ElseIf Len(ServerBackendPath) > 0 And Len(Dir(ServerBackendPath)) > 0 Then
        BackendPath = ServerBackendPath
        FindBackend = ServerBE
    End If
End Function