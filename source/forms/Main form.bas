Version =21
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4932
    DatasheetFontHeight =11
    ItemSuffix =11
    Right =21720
    Bottom =11415
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x1ba91f6bb054e540
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
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =1582
            Name ="Detail"
            AutoHeight =255
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
                    Width =4230
                    Height =555
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label0"
                    Caption ="Here's the main form"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =360
                    LayoutCachedWidth =4590
                    LayoutCachedHeight =915
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =360
                    Top =975
                    Width =4230
                    Height =285
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label2"
                    Caption ="It's ugly."
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =975
                    LayoutCachedWidth =4590
                    LayoutCachedHeight =1260
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =360
                    Top =1320
                    Width =4230
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Relink output"
                    Caption ="No relink code has run yet"
                    EventProcPrefix ="Relink_output"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =1320
                    LayoutCachedWidth =4590
                    LayoutCachedHeight =1560
                    RowStart =2
                    RowEnd =2
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

'Suppress Rubberduck code inspections as Rubberduck is unable to resolve form
'   elements as of v2.5.2.
'@IgnoreModule VariableNotAssigned, UnassignedVariableUsage, UndeclaredVariable


Private Sub Form_Load()
    Dim databaseMessageOutput As String
    Select Case LinkedBackendType
        Case DevBE
            databaseMessageOutput = "Using development backend."
        Case LocalBE
            databaseMessageOutput = "Using local backend."
        Case ServerBE
            databaseMessageOutput = "Using server backend."
        Case Else
            databaseMessageOutput = "Not using a backend :("
    End Select
    Relink_output.Caption = databaseMessageOutput
End Sub
