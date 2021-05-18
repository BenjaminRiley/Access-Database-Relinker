# Automatic Database Relinker v1.0.0

Code and example database for automatically reacquiring the backend for a split Access database.

For an Access frontend called `Database.accdb`, a backends will be acquired by searching for the following files, in order:

1. **A development backend copy** named `Database_be_dev.accdb`, in the same folder as the frontend.
2. **A local backend** named `Database_be.accdb`, in the same folder as the frontend.
3. **A server backend** with a filepath provided in code.

This allows for easier development or home usage of the database without having to manually relink the backend if it is moved or unavailable.

## Usage

From the VBA editor, right click in the Project Explorer window and select *Import file…* and import the `DBLinker.bas` module from this repository.

Modify your startup form to call the `RelinkTables` function in the form's On Load or On Open event, e.g.

```vbscript
Private Sub Form_Load()
    RelinkTables  "Linked Table", "\\Server\Folder\MyBackend.accdb"
End Sub
```

If no shared backend is needed (e.g. you are working locally and only want a development copy of the backend), you can omit the server path e.g.

```vbscript
Private Sub Form_Load()
    RelinkTables  "Linked Table"
End Sub
```

To prevent the function running repeatedly, it may be preferable to use a dedicated 'spash screen' form that opens the existing startup form after linking, or to check the `LinkedBackendType` property before attempting to link the table, e.g.

```vbscript
Private Sub Form_Load()
    If LinkedBackendType = NotLinked Then
        RelinkTables "Linked Table", "\\Server\Folder\MyBackend.accdb"
    End If
End Sub
```

The `RelinkTables` function returns a `BackendType` Enum (included in the module) and sets the `LinkedBackendType` Property. This could be used to display an error message if the backend was not found, or to place a notification on your main form that says which backend is being used.

```vbscript
'Shows an error message and then quits if backend missing using returned value
Private Sub Form_Load()
    If RelinkTables("Linked Table", "\\Server\Folder\MyBackend.accdb") = NotLinked Then
        MsgBox "Couldn't find the backend!", vbOKOnly + vbCritical
        Application.Quit
    End If
End Sub
```

```vbscript
'Changes the caption of the "Relink output" label on the form using the LinkedBackendType property
Private Sub Form_Load()
    RelinkTables "Linked Table", "\\Server\Folder\MyBackend.accdb"
    Select Case LinkedBackendType
        Case DevBE
            Relink_output.Caption = "Using development backend."
        Case LocalBE
            Relink_output.Caption = "Using local backend."
        Case ServerBE
            Relink_output.Caption = "Using server backend."
        Case Else
            Relink_output.Caption = "Backend missing :("
    End Select
End Sub
```

See the example database for a more complete example using a simple splash screen.

## `RelinkTables` Function

Definition:

```vbscript
Public Function RelinkTables(ReferenceTableName As String, Optional ServerBackendPath As String) As BackendType
```

### Parameters

- **`ReferenceTableName`** Name of a linked table in the database. This is used as a reference to determine which links the script should replace. Required.
- **`ServerBackendPath`** Path to a shared backend on a server. This is used if there is no dev or local copy of the backend. Optional.

### Returns

Sets the `LinkedBackend` Property and returns it.

## `LinkedBackendType` Property

Read-only property that stores the type of backend that was linked by `RelinkTables`.

Initially will be `NotLinked` (`0`).

This property is set every time `RelinkTables` is called.

## `BackendType` Enum

Used to distinguish between the different types of backend that could be linked.

The possible values are as follows:

| Value             | Description                                                                       |
|-------------------|-----------------------------------------------------------------------------------|
| `NotLinked` (`0`) | No backend linked (may be missing or `RelinkTables` not called yet)               |
| `DevBE`           | Using a development copy of the backend                                           |
| `LocalBE`         | Using a local copy of the backend                                                 |
| `ServerBE`        | Using a shared backend from the `ServerBackendPath` provided to `RelinkTables`    |

When used as a `Boolean` (e.g. in an `If … Then` statement), `NotLinked` should evaulate to `False` and all other values to `True`.

## Other notes

The example Access project makes use of [msaccess-vcs-integration](https://github.com/msaccess-vcs-integration/msaccess-vcs-integration) to aid with version control.
