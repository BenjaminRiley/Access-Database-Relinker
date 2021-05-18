Version =21
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4920
    DatasheetFontHeight =11
    ItemSuffix =16
    Right =21975
    Bottom =11385
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x7746d12fb254e540
    End
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =2664
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =360
                    Top =360
                    Width =4530
                    Height =555
                    BorderColor =8355711
                    Name ="Label0"
                    Caption ="Connecting to database..."
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =360
                    LayoutCachedWidth =4890
                    LayoutCachedHeight =915
                    LayoutGroup =1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =360
                    Top =975
                    Width =4530
                    Height =555
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label12"
                    Caption ="If a yellow security warning is displayed above you may need to enable active co"
                        "ntent."
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =975
                    LayoutCachedWidth =4890
                    LayoutCachedHeight =1530
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Database
'@Folder Forms

'Switch for editing the splash screen in LayoutView. Remember to set it to False before deploying the frontend or it
'   will 'break'.
Private Const IsSplashDisabled As Boolean = False

'Change this to point to the desired backend file on the server.
Private Const ServerPath As String = "C:\PretendThisIsANetworkShare\RelinkExample_be.accdb"

'Name of a linked table.
'The relink script uses the link information from this table to determine which tables it should relink.
Private Const LinkedTable As String = "RemoteTable"

'Name of the form to open after relinking the backend.
Private Const MainForm As String = "Main form"

'Message to display if the backend cannot be found.
Private Const ErrorMessage As String = "The program failed to locate the backend database."


Private Sub Form_Load()
    If Not IsSplashDisabled Then
        DoRelink
    End If
End Sub


Private Sub DoRelink()
    If RelinkTables(LinkedTable, ServerPath) Then
        OpenMainAndExitSplash
    Else
        'Show an Abort/Retry/Ignore error, defaulting to Retry
        Dim errorResponse As VbMsgBoxResult
        errorResponse = MsgBox(ErrorMessage, vbAbortRetryIgnore + vbCritical + vbDefaultButton2)
        Select Case errorResponse
            Case vbRetry
                DoRelink    'Re-call this subroutine.
                            'This will cause a stack overflow if repeated enough. Good luck!
            Case vbAbort
                Application.Quit
            Case Else 'Ignore or any other response opens the main form anyway
                OpenMainAndExitSplash
        End Select
    End If
End Sub


Private Sub OpenMainAndExitSplash()
    'Open Main form
    If Not CurrentProject.AllForms.Item(MainForm).IsLoaded Then
        DoCmd.OpenForm MainForm
    End If
    'Close splash form
    DoCmd.Close acForm, Me.Name
End Sub
