#tag DesktopWindow
Begin DesktopWindow wNewOrEditDefinition
   Backdrop        =   0
   BackgroundColor =   &cFFFFFF
   Composite       =   False
   DefaultLocation =   2
   FullScreen      =   False
   HasBackgroundColor=   False
   HasCloseButton  =   True
   HasFullScreenButton=   False
   HasMaximizeButton=   False
   HasMinimizeButton=   False
   HasTitleBar     =   True
   Height          =   451
   ImplicitInstance=   False
   MacProcID       =   0
   MaximumHeight   =   32000
   MaximumWidth    =   32000
   MenuBar         =   ""
   MenuBarVisible  =   False
   MinimumHeight   =   64
   MinimumWidth    =   64
   Resizeable      =   False
   Title           =   "New annual event definition"
   Type            =   1
   Visible         =   True
   Width           =   580
   BeginDesktopSegmentedButton DesktopSegmentedButton SegmentedButton1
      Enabled         =   True
      Height          =   24
      Index           =   -2147483648
      Left            =   10
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      Segments        =   "Fixed day\n\nTrue\rFrom Easter\n\nFalse\rFrom Orthodox Easter\n\nFalse\rRank WeekDay\n\nFalse"
      SelectionStyle  =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   False
      Tooltip         =   ""
      Top             =   10
      Transparent     =   False
      Visible         =   True
      Width           =   560
   End
   Begin DesktopLabel Label14
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      Italic          =   False
      Left            =   10
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   5
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "First year of the cycle"
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   331
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   113
   End
   Begin DesktopLabel Label15
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      Italic          =   False
      Left            =   225
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   7
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Cycle Duration (years)"
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   331
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   117
   End
   Begin DesktopTextField txtCycleDuration
      AllowAutoDeactivate=   True
      AllowFocusRing  =   True
      AllowSpellChecking=   False
      AllowTabs       =   False
      BackgroundColor =   &cFFFFFF
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Format          =   ""
      HasBorder       =   True
      Height          =   22
      Hint            =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   345
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MaximumCharactersAllowed=   0
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   8
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "1"
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   331
      Transparent     =   False
      Underline       =   False
      ValidationMask  =   ""
      Visible         =   True
      Width           =   51
   End
   Begin DesktopTextField txtFirstYearOfTheCycle
      AllowAutoDeactivate=   True
      AllowFocusRing  =   True
      AllowSpellChecking=   False
      AllowTabs       =   False
      BackgroundColor =   &cFFFFFF
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Format          =   ""
      HasBorder       =   True
      Height          =   22
      Hint            =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   127
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MaximumCharactersAllowed=   0
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   6
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "1900"
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   331
      Transparent     =   False
      Underline       =   False
      ValidationMask  =   ""
      Visible         =   True
      Width           =   51
   End
   Begin DesktopUpDownArrows UpDownArrows4
      Active          =   False
      AllowAutoDeactivate=   True
      AllowFocus      =   False
      AllowTabStop    =   True
      Enabled         =   True
      Height          =   23
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   401
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   9
      TabPanelIndex   =   0
      Tooltip         =   ""
      Top             =   331
      Transparent     =   False
      Visible         =   True
      Width           =   14
      _mIndex         =   0
      _mInitialParent =   ""
      _mName          =   ""
      _mPanelIndex    =   0
   End
   Begin DesktopDateTimePicker dtStart
      Active          =   False
      AllowAutoDeactivate=   True
      AllowFocusRing  =   False
      AllowTabStop    =   True
      DisplayMode     =   1
      DisplaySeconds  =   False
      Enabled         =   True
      GraphicalDisplay=   True
      Height          =   22
      HourMode        =   0
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   10
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   12
      TabPanelIndex   =   0
      TodayButtonCaption=   ""
      Tooltip         =   ""
      Top             =   385
      Transparent     =   False
      Visible         =   True
      Width           =   113
      _mIndex         =   0
      _mInitialParent =   ""
      _mName          =   ""
      _mPanelIndex    =   0
   End
   Begin DesktopLabel Label16
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   17
      Index           =   -2147483648
      Italic          =   False
      Left            =   10
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   10
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Start of validity"
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   363
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   113
   End
   Begin DesktopDateTimePicker dtEnd
      Active          =   False
      AllowAutoDeactivate=   True
      AllowFocusRing  =   False
      AllowTabStop    =   True
      DisplayMode     =   1
      DisplaySeconds  =   False
      Enabled         =   True
      GraphicalDisplay=   True
      Height          =   22
      HourMode        =   0
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   135
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   13
      TabPanelIndex   =   0
      TodayButtonCaption=   ""
      Tooltip         =   ""
      Top             =   385
      Transparent     =   False
      Visible         =   True
      Width           =   113
      _mIndex         =   0
      _mInitialParent =   ""
      _mName          =   ""
      _mPanelIndex    =   0
   End
   Begin DesktopLabel Label17
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   17
      Index           =   -2147483648
      Italic          =   False
      Left            =   135
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   11
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "End of validity"
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   363
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   113
   End
   Begin DesktopCheckBox chkDayOff
      AllowAutoDeactivate=   True
      Bold            =   False
      Caption         =   "Day off"
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      Italic          =   False
      Left            =   10
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   0
      TabIndex        =   4
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   299
      Transparent     =   False
      Underline       =   False
      Value           =   False
      Visible         =   True
      VisualState     =   1
      Width           =   100
   End
   Begin DesktopLabel Label5
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      Italic          =   False
      Left            =   10
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Caption"
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   46
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   43
   End
   Begin DesktopTextField txtCaption
      AllowAutoDeactivate=   True
      AllowFocusRing  =   True
      AllowSpellChecking=   False
      AllowTabs       =   False
      BackgroundColor =   &cFFFFFF
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Format          =   ""
      HasBorder       =   True
      Height          =   22
      Hint            =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   10
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MaximumCharactersAllowed=   0
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "New day"
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   67
      Transparent     =   False
      Underline       =   False
      ValidationMask  =   ""
      Visible         =   True
      Width           =   560
   End
   Begin DesktopPagePanel PagePanel1
      AllowAutoDeactivate=   True
      Enabled         =   True
      Height          =   186
      Index           =   -2147483648
      Left            =   10
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      PanelCount      =   4
      Panels          =   ""
      Scope           =   0
      SelectedPanelIndex=   0
      TabIndex        =   3
      TabPanelIndex   =   0
      TabStop         =   False
      Tooltip         =   ""
      Top             =   101
      Transparent     =   False
      Value           =   0
      Visible         =   True
      Width           =   560
      Begin DesktopLabel Label1
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   10
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   0
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "Month"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   101
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   43
         Begin DesktopLabel Label9
            AllowAutoDeactivate=   True
            Bold            =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   20
            Index           =   -2147483648
            InitialParent   =   "Label1"
            Italic          =   False
            Left            =   10
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   0
            TabPanelIndex   =   4
            TabStop         =   True
            Text            =   "Month"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   101
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   43
         End
      End
      Begin DesktopPopupMenu popFDMonth
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         InitialValue    =   ""
         Italic          =   False
         Left            =   10
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedRowIndex=   -1
         TabIndex        =   2
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   122
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   165
      End
      Begin DesktopPopupMenu popFDDay
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         InitialValue    =   ""
         Italic          =   False
         Left            =   190
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedRowIndex=   -1
         TabIndex        =   3
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   122
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   165
      End
      Begin DesktopLabel Label3
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   10
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   7
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "Shift to the                            day of the week"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   216
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   61
      End
      Begin DesktopPopupMenu popFDWeekDay
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         InitialValue    =   ""
         Italic          =   False
         Left            =   154
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedRowIndex=   -1
         TabIndex        =   9
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   215
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   165
      End
      Begin DesktopLabel Label4
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   329
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   10
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "or adding days"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   215
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   85
      End
      Begin DesktopTextField txtFDAdding
         AllowAutoDeactivate=   True
         AllowFocusRing  =   True
         AllowSpellChecking=   False
         AllowTabs       =   False
         BackgroundColor =   &cFFFFFF
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Format          =   ""
         HasBorder       =   True
         Height          =   22
         Hint            =   ""
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   413
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MaximumCharactersAllowed=   0
         Password        =   False
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   11
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "0"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   215
         Transparent     =   False
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   51
      End
      Begin DesktopUpDownArrows UpDownArrows1
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         Enabled         =   True
         Height          =   23
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Left            =   469
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         PanelIndex      =   0
         Scope           =   0
         TabIndex        =   12
         TabPanelIndex   =   1
         Tooltip         =   ""
         Top             =   215
         Transparent     =   False
         Visible         =   True
         Width           =   14
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopLabel Label6
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   12
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   0
         TabPanelIndex   =   2
         TabStop         =   True
         Text            =   "Difference in days with Easter"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   101
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   158
      End
      Begin DesktopTextField txtEDDiff
         AllowAutoDeactivate=   True
         AllowFocusRing  =   True
         AllowSpellChecking=   False
         AllowTabs       =   False
         BackgroundColor =   &cFFFFFF
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Format          =   ""
         HasBorder       =   True
         Height          =   22
         Hint            =   ""
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   174
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MaximumCharactersAllowed=   0
         Password        =   False
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   1
         TabPanelIndex   =   2
         TabStop         =   True
         Text            =   "0"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   101
         Transparent     =   False
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   51
      End
      Begin DesktopButton ButtonEaster
         AllowAutoDeactivate=   True
         Bold            =   True
         Cancel          =   False
         Caption         =   "Easter"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   147
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   18
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   262
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster1
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Good Friday"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   147
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   12
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   198
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster2
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Whit Monday"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   284
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   13
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   198
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster3
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Ascension Thursday"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   284
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   16
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   230
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster4
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Corpus Christi"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   421
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   6
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   133
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster5
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Sunday after Corpus Christi"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   10.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   421
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   10
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   166
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster6
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Trinity Sunday"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   284
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   19
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   262
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster7
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Mardi gras"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   10
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   7
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   166
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster8
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Palm Sunday"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   147
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   4
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   133
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster9
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Ash Wednesday"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   10
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   11
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   198
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster16
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Mothering Sunday"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   10
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   14
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   230
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster18
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Maundy Thursday"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   147
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   8
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   166
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster19
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Holy Saturday"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   147
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   15
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   230
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster20
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Second Sunday of Easter"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   10.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   284
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   5
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   133
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster21
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Pentecost"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   284
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   9
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   166
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopButton ButtonEaster17
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Passion Sunday"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   10
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   17
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   262
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopLabel Label10
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   198
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   0
         TabPanelIndex   =   4
         TabStop         =   True
         Text            =   "Day of week"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   101
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   73
      End
      Begin DesktopPopupMenu popWDWeekDay
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         InitialValue    =   ""
         Italic          =   False
         Left            =   198
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedRowIndex=   -1
         TabIndex        =   3
         TabPanelIndex   =   4
         TabStop         =   True
         Tooltip         =   ""
         Top             =   122
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   165
      End
      Begin DesktopPopupMenu popWDRank
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         InitialValue    =   ""
         Italic          =   False
         Left            =   375
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedRowIndex=   -1
         TabIndex        =   4
         TabPanelIndex   =   4
         TabStop         =   True
         Tooltip         =   ""
         Top             =   122
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   165
      End
      Begin DesktopLabel Label12
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   10
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   5
         TabPanelIndex   =   4
         TabStop         =   True
         Text            =   "Shift to the                            day of the week"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   170
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   59
      End
      Begin DesktopPopupMenu popWDToWeekDay
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         InitialValue    =   ""
         Italic          =   False
         Left            =   161
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedRowIndex=   -1
         TabIndex        =   7
         TabPanelIndex   =   4
         TabStop         =   True
         Tooltip         =   ""
         Top             =   170
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   165
      End
      Begin DesktopLabel Label13
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   338
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   8
         TabPanelIndex   =   4
         TabStop         =   True
         Text            =   "or adding days"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   170
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   85
      End
      Begin DesktopTextField txtWDAdding
         AllowAutoDeactivate=   True
         AllowFocusRing  =   True
         AllowSpellChecking=   False
         AllowTabs       =   False
         BackgroundColor =   &cFFFFFF
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Format          =   ""
         HasBorder       =   True
         Height          =   22
         Hint            =   ""
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   427
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MaximumCharactersAllowed=   0
         Password        =   False
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   9
         TabPanelIndex   =   4
         TabStop         =   True
         Text            =   "0"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   170
         Transparent     =   False
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   51
      End
      Begin DesktopUpDownArrows UpDownArrows3
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         Enabled         =   True
         Height          =   23
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Left            =   483
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         PanelIndex      =   0
         Scope           =   0
         TabIndex        =   10
         TabPanelIndex   =   4
         Tooltip         =   ""
         Top             =   170
         Transparent     =   False
         Visible         =   True
         Width           =   14
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopPopupMenu popWDShiftMode
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         InitialValue    =   ""
         Italic          =   False
         Left            =   75
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedRowIndex=   -1
         TabIndex        =   6
         TabPanelIndex   =   4
         TabStop         =   True
         Tooltip         =   ""
         Top             =   170
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   74
      End
      Begin DesktopButton ButtonEaster15
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Clean Monday"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   10
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   3
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   133
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   125
      End
      Begin DesktopPopupMenu popFDShiftMode
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         InitialValue    =   ""
         Italic          =   False
         Left            =   75
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedRowIndex=   -1
         TabIndex        =   8
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   215
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   74
      End
      Begin DesktopLabel Label11
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   375
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   1
         TabPanelIndex   =   4
         TabStop         =   True
         Text            =   "Rank"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   101
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   43
      End
      Begin DesktopLabel Label2
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   190
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   1
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "Day"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   101
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   43
      End
      Begin DesktopCheckBox cbFDFridayIfSaturday
         AllowAutoDeactivate=   True
         Bold            =   False
         Caption         =   "Previous friday if saturday"
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   167
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         TabIndex        =   5
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   167
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         VisualState     =   0
         Width           =   158
      End
      Begin DesktopCheckBox cbFDMondayIfSunday
         AllowAutoDeactivate=   True
         Bold            =   False
         Caption         =   "Next monday if sunday"
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   10
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         TabIndex        =   4
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   167
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         VisualState     =   0
         Width           =   142
      End
      Begin DesktopLabel Label7
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   12
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   0
         TabPanelIndex   =   3
         TabStop         =   True
         Text            =   "Difference in days with Easter"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   101
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   158
      End
      Begin DesktopTextField txtEDDiff1
         AllowAutoDeactivate=   True
         AllowFocusRing  =   True
         AllowSpellChecking=   False
         AllowTabs       =   False
         BackgroundColor =   &cFFFFFF
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Format          =   ""
         HasBorder       =   True
         Height          =   22
         Hint            =   ""
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   174
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MaximumCharactersAllowed=   0
         Password        =   False
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   1
         TabPanelIndex   =   3
         TabStop         =   True
         Text            =   "0"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   101
         Transparent     =   False
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   51
      End
      Begin DesktopPopupMenu popWDMonth
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         InitialValue    =   ""
         Italic          =   False
         Left            =   10
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedRowIndex=   -1
         TabIndex        =   2
         TabPanelIndex   =   4
         TabStop         =   True
         Tooltip         =   ""
         Top             =   121
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   165
      End
      Begin DesktopUpDownArrows UpDownArrows2
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         Enabled         =   True
         Height          =   23
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Left            =   230
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         PanelIndex      =   0
         Scope           =   0
         TabIndex        =   20
         TabPanelIndex   =   2
         Tooltip         =   ""
         Top             =   101
         Transparent     =   False
         Visible         =   True
         Width           =   14
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopUpDownArrows UpDownArrows5
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         Enabled         =   True
         Height          =   23
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Left            =   230
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         PanelIndex      =   0
         Scope           =   0
         TabIndex        =   2
         TabPanelIndex   =   3
         Tooltip         =   ""
         Top             =   101
         Transparent     =   False
         Visible         =   True
         Width           =   14
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopCheckBox chkDFAlways
         AllowAutoDeactivate=   True
         Bold            =   False
         Caption         =   "Always"
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "PagePanel1"
         Italic          =   False
         Left            =   154
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         TabIndex        =   13
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   248
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         VisualState     =   0
         Width           =   100
      End
   End
   Begin DesktopButton btnOK
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "&Add"
      Default         =   True
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   22
      Index           =   -2147483648
      Italic          =   False
      Left            =   398
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      TabIndex        =   15
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   419
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
   Begin DesktopButton btnCancel
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   True
      Caption         =   "&Cancel"
      Default         =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   22
      Index           =   -2147483648
      Italic          =   False
      Left            =   490
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      TabIndex        =   16
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   419
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
   Begin DesktopButton btnCopy
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "&Copy code"
      Default         =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   22
      Index           =   -2147483648
      Italic          =   False
      Left            =   10
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      TabIndex        =   14
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   419
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   113
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Method, Flags = &h0
		Function Charge(d as AnnualEvent = Nil) As AnnualEvent
		  If d <> Nil Then
		    Me.Title = "Edit annual event definition"
		    btnOK.Caption = "&OK"
		    LoadEvent(d)
		  end
		  
		  Me.ShowModal
		  
		  Return AnnualEventReturned
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadEvent(d as AnnualEvent)
		  txtCaption.Text= d.Caption
		  chkDayOff.Value = d.DayOff
		  dtStart.SelectedDate = d.StartOfValidity
		  dtEnd.SelectedDate = d.EndOfValidity
		  txtFirstYearOfTheCycle.Text = d.CycleFirstYear.ToString
		  txtCycleDuration.Text = d.CycleYearDuration.ToString
		  
		  Select Case d
		    
		  Case IsA AnnualEventFix 
		    
		    Var df As AnnualEventFix = d.DefinitionObject
		    
		    popFDMonth.SelectedRowIndex = df.Month - 1
		    popFDDay.SelectedRowIndex = df.day - 1
		    cbFDMondayIfSunday.Value = df.MondayIfSunday
		    cbFDFridayIfSaturday.Value = df.FridayIfSaturday
		    
		    If df.PreviousWeekDay > 0 Then
		      popFDShiftMode.SelectedRowIndex = 0 // Previous
		      popFDWeekDay.SelectedRowIndex = df.PreviousWeekDay
		    ElseIf df.NextWeekDay > 0 Then
		      popFDShiftMode.SelectedRowIndex = 1 // Next
		      popFDWeekDay.SelectedRowIndex = df.NextWeekDay
		    ElseIf df.AddDays <> 0 Then
		      txtFDAdding.Text = df.AddDays.ToString
		    End If
		    
		    
		    
		  Case IsA AnnualEventEaster
		    
		    Var de As AnnualEventEaster = d.DefinitionObject
		    txtEDDiff.Text = de.DeltaEaster.ToString
		    
		    SegmentedButton1.SelectedSegmentIndex = 1
		    PagePanel1.SelectedPanelIndex = 1
		    
		  Case IsA AnnualEventOrthodoxEaster
		    
		    Var doe As AnnualEventOrthodoxEaster = d.DefinitionObject
		    txtEDDiff1.Text = doe.DeltaEaster.ToString
		    
		    SegmentedButton1.SelectedSegmentIndex = 2
		    PagePanel1.SelectedPanelIndex = 2
		    
		  Case IsA AnnualEventWeekDay
		    
		    Var wd As AnnualEventWeekDay = d.DefinitionObject
		    
		    popWDMonth.SelectedRowIndex =  wd.Month - 1
		    popWDWeekDay.SelectedRowIndex = wd.WeekDay - 1
		    popWDRank.SelectedRowIndex = wd.Rank - 1
		    
		    If wd.PreviousWeekDay > 0 Then
		      popWDShiftMode.SelectedRowIndex = 0 // Previous
		      popWDToWeekDay.SelectedRowIndex = wd.PreviousWeekDay
		    ElseIf wd.NextWeekDay > 0 Then
		      popWDShiftMode.SelectedRowIndex = 1 // Next
		      popWDToWeekDay.SelectedRowIndex = wd.NextWeekDay
		    ElseIf wd.AddDays <> 0 Then
		      txtWDAdding.Text = wd.AddDays.ToString
		    End If
		    
		    SegmentedButton1.SelectedSegmentIndex = 3
		    PagePanel1.SelectedPanelIndex = 3
		    
		  end
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private AnnualEventReturned As AnnualEvent
	#tag EndProperty


#tag EndWindowCode

#tag Events SegmentedButton1
	#tag Event
		Sub Pressed(segmentIndex As Integer)
		  PagePanel1.SelectedPanelIndex = segmentIndex
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events txtCycleDuration
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if key.Asc < 32 then Return false// Specials char 
		  if key.asc = 127 then Return false // backspace
		  
		  If key = "-" Then
		    If Me.Text.IndexOf("-") > 0 Then Return True
		    If Me.SelectionStart > 0 Then Return True
		    Return False
		  end
		  
		  If Not (key>="0" And key<="9") Then Return True
		  
		  // ok
		  Return False
		End Function
	#tag EndEvent
	#tag Event
		Sub FocusLost()
		  
		  #Pragma BreakOnExceptions False
		  Try
		    Me.Text = Me.Text.ToInteger.ToString
		  Catch
		    me.Text = "0"
		  End
		  #Pragma BreakOnExceptions True
		  
		  Select Case SegmentedButton1.SelectedSegmentIndex
		    
		  Case 0
		    
		    
		    
		  Case 1
		    
		    
		    
		  Case 2
		    
		    
		    
		  end
		End Sub
	#tag EndEvent
	#tag Event
		Sub TextChanged()
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events txtFirstYearOfTheCycle
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if key.Asc < 32 then Return false// Specials char 
		  if key.asc = 127 then Return false // backspace
		  
		  If key = "-" Then
		    If Me.Text.IndexOf("-") > 0 Then Return True
		    If Me.SelectionStart > 0 Then Return True
		    Return False
		  end
		  
		  If Not (key>="0" And key<="9") Then Return True
		  
		  // ok
		  Return False
		End Function
	#tag EndEvent
	#tag Event
		Sub FocusLost()
		  
		  #Pragma BreakOnExceptions False
		  Try
		    Me.Text = Me.Text.ToInteger.ToString
		  Catch
		    me.Text = "0"
		  End
		  #Pragma BreakOnExceptions True
		  
		  if me.Text.ToInteger < 1 then me.text = "1"
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events UpDownArrows4
	#tag Event
		Sub DownPressed()
		  Var v As Integer
		  
		  
		  #Pragma BreakOnExceptions False
		  Try
		    v= txtCycleDuration.Text.ToInteger
		  Catch
		    v = 0
		  End
		  #Pragma BreakOnExceptions True
		  
		  v = v - 1
		  
		  if txtCycleDuration.Text.ToInteger < 0 then txtCycleDuration.Text = "0"
		  
		  txtCycleDuration.Text = v.ToString
		  
		  
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub UpPressed()
		  Var v As Integer
		  
		  
		  #Pragma BreakOnExceptions False
		  Try
		    v= txtCycleDuration.Text.ToInteger
		  Catch
		    v = 0
		  End
		  #Pragma BreakOnExceptions True
		  
		  v = v + 1
		  
		  txtCycleDuration.Text = v.ToString
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events dtStart
	#tag Event
		Sub Opening()
		  me.SelectedDate = new DateTime(1,1,1)
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events dtEnd
	#tag Event
		Sub Opening()
		  me.SelectedDate = new DateTime(3999,12,31)
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events popFDMonth
	#tag Event
		Sub Opening()
		  Var d As DateTime
		  For i As Integer = 1 To 12
		    d = New DateTime(2025, i, 1)
		    Me.AddRow d.ToString("MMMM", Locale.Current)
		  Next 
		  
		  me.SelectedRowIndex = 0
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events popFDDay
	#tag Event
		Sub Opening()
		  For i As Integer = 1 To 31
		    me.AddRow i.ToString
		  Next 
		  
		  me.SelectedRowIndex = 0
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events popFDWeekDay
	#tag Event
		Sub Opening()
		  Me.AddRow "0 - <Disabled>"
		  
		  Var d As New DateTime(2025, 8, 30)
		  For i As Integer = 1 To 7
		    d = d.AddInterval(0,0,1)
		    Me.AddRow d.DayOfWeek.ToString + " - " + d.ToString("EEEE", Locale.Current)
		  Next 
		  
		  
		  
		  me.SelectedRowIndex = 0
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events txtFDAdding
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if key.Asc < 32 then Return false// Specials char 
		  if key.asc = 127 then Return false // backspace
		  
		  If key = "-" Then
		    If Me.Text.IndexOf("-") > 0 Then Return True
		    If Me.SelectionStart > 0 Then Return True
		    Return False
		  end
		  
		  If Not (key>="0" And key<="9") Then Return True
		  
		  // ok
		  Return False
		End Function
	#tag EndEvent
	#tag Event
		Sub FocusLost()
		  
		  #Pragma BreakOnExceptions False
		  Try
		    Me.Text = Me.Text.ToInteger.ToString
		  Catch
		    me.Text = "0"
		  End
		  #Pragma BreakOnExceptions True
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events UpDownArrows1
	#tag Event
		Sub DownPressed()
		  Var v As Integer
		  
		  
		  #Pragma BreakOnExceptions False
		  Try
		    v= txtFDAdding.Text.ToInteger
		  Catch
		    v = 0
		  End
		  #Pragma BreakOnExceptions True
		  
		  v = v - 1
		  
		  txtFDAdding.Text = v.ToString
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub UpPressed()
		  Var v As Integer
		  
		  
		  #Pragma BreakOnExceptions False
		  Try
		    v= txtFDAdding.Text.ToInteger
		  Catch
		    v = 0
		  End
		  #Pragma BreakOnExceptions True
		  
		  v = v + 1
		  
		  txtFDAdding.Text = v.ToString
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events txtEDDiff
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if key.Asc < 32 then Return false// Specials char 
		  if key.asc = 127 then Return false // backspace
		  
		  If key = "-" Then
		    If Me.Text.IndexOf("-") > 0 Then Return True
		    If Me.SelectionStart > 0 Then Return True
		    Return False
		  end
		  
		  If Not (key>="0" And key<="9") Then Return True
		  
		  // ok
		  Return False
		End Function
	#tag EndEvent
	#tag Event
		Sub FocusLost()
		  
		  #Pragma BreakOnExceptions False
		  Try
		    Me.Text = Me.Text.ToInteger.ToString
		  Catch
		    me.Text = "0"
		  End
		  #Pragma BreakOnExceptions True
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "0"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster1
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "-2"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster2
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "50"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster3
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "40"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster4
	#tag Event
		Sub Pressed()
		  txtCaption.Text = Me.Caption
		  txtEDDiff.Text = "60"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster5
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "63"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster6
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "56"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster7
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = AnnualEventEaster.MardiGras.ToString // or txtEDDiff.Text = "-47"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster8
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "-7"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster9
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = AnnualEventEaster.AshWednesday.ToString // // or txtEDDiff.Text = "-46"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster16
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "-21" // Or txtEDDiff.Text = AnnualEventEaster.MotheringSunday.ToString
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster18
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "-3"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster19
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "-1"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster20
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "7"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster21
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "49"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster17
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "-14"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events popWDWeekDay
	#tag Event
		Sub Opening()
		  Var d As New DateTime(2025, 8, 30)
		  For i As Integer = 1 To 7
		    d = d.AddInterval(0,0,1)
		    Me.AddRow d.DayOfWeek.ToString + " - " + d.ToString("EEEE", Locale.Current)
		  Next 
		  
		  
		  
		  me.SelectedRowIndex = 0
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events popWDRank
	#tag Event
		Sub Opening()
		  Me.AddRow "1st"
		  Me.AddRow "2nd"
		  Me.AddRow "3rd"
		  Me.AddRow "4th"
		  Me.AddRow "5th"
		  me.AddRow "Last"
		  
		  me.SelectedRowIndex = 0
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events popWDToWeekDay
	#tag Event
		Sub Opening()
		  Me.AddRow "0 - <Disabled>"
		  
		  Var d As New DateTime(2025, 8, 30)
		  For i As Integer = 1 To 7
		    d = d.AddInterval(0,0,1)
		    Me.AddRow d.DayOfWeek.ToString + " - " + d.ToString("EEEE", Locale.Current)
		  Next 
		  
		  
		  
		  me.SelectedRowIndex = 0
		End Sub
	#tag EndEvent
	#tag Event
		Sub SelectionChanged(item As DesktopMenuItem)
		  #Pragma Unused item
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events txtWDAdding
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if key.Asc < 32 then Return false// Specials char 
		  if key.asc = 127 then Return false // backspace
		  
		  If key = "-" Then
		    If Me.Text.IndexOf("-") > 0 Then Return True
		    If Me.SelectionStart > 0 Then Return True
		    Return False
		  end
		  
		  If Not (key>="0" And key<="9") Then Return True
		  
		  // ok
		  Return False
		End Function
	#tag EndEvent
	#tag Event
		Sub FocusLost()
		  
		  #Pragma BreakOnExceptions False
		  Try
		    Me.Text = Me.Text.ToInteger.ToString
		  Catch
		    me.Text = "0"
		  End
		  #Pragma BreakOnExceptions True
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events UpDownArrows3
	#tag Event
		Sub DownPressed()
		  Var v As Integer
		  
		  
		  #Pragma BreakOnExceptions False
		  Try
		    v= txtWDAdding.Text.ToInteger
		  Catch
		    v = 0
		  End
		  #Pragma BreakOnExceptions True
		  
		  v = v - 1
		  
		  txtWDAdding.Text = v.ToString
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub UpPressed()
		  Var v As Integer
		  
		  
		  #Pragma BreakOnExceptions False
		  Try
		    v= txtWDAdding.Text.ToInteger
		  Catch
		    v = 0
		  End
		  #Pragma BreakOnExceptions True
		  
		  v = v + 1
		  
		  txtWDAdding.Text = v.ToString
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events popWDShiftMode
	#tag Event
		Sub Opening()
		  Me.AddRow "previous"
		  Me.AddRow "next"
		  me.SelectedRowIndex = 0
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster15
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = AnnualEventEaster.CleanMonday.ToString // or txtEDDiff.Text = "-48"
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events popFDShiftMode
	#tag Event
		Sub Opening()
		  Me.AddRow "previous"
		  Me.AddRow "next"
		  me.SelectedRowIndex = 0
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events txtEDDiff1
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if key.Asc < 32 then Return false// Specials char 
		  if key.asc = 127 then Return false // backspace
		  
		  If key = "-" Then
		    If Me.Text.IndexOf("-") > 0 Then Return True
		    If Me.SelectionStart > 0 Then Return True
		    Return False
		  end
		  
		  If Not (key>="0" And key<="9") Then Return True
		  
		  // ok
		  Return False
		End Function
	#tag EndEvent
	#tag Event
		Sub FocusLost()
		  
		  #Pragma BreakOnExceptions False
		  Try
		    Me.Text = Me.Text.ToInteger.ToString
		  Catch
		    me.Text = "0"
		  End
		  #Pragma BreakOnExceptions True
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub TextChanged()
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events popWDMonth
	#tag Event
		Sub Opening()
		  Var d As DateTime
		  For i As Integer = 1 To 12
		    d = New DateTime(2025, i, 1)
		    Me.AddRow d.ToString("MMMM", Locale.Current)
		  Next 
		  
		  me.SelectedRowIndex = 0
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events UpDownArrows2
	#tag Event
		Sub DownPressed()
		  Var v As Integer
		  
		  
		  #Pragma BreakOnExceptions False
		  Try
		    v= txtEDDiff.Text.ToInteger
		  Catch
		    v = 0
		  End
		  #Pragma BreakOnExceptions True
		  
		  v = v - 1
		  
		  txtEDDiff.Text = v.ToString
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub UpPressed()
		  Var v As Integer
		  
		  
		  #Pragma BreakOnExceptions False
		  Try
		    v= txtEDDiff.Text.ToInteger
		  Catch
		    v = 0
		  End
		  #Pragma BreakOnExceptions True
		  
		  v = v + 1
		  
		  txtEDDiff.Text = v.ToString
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events UpDownArrows5
	#tag Event
		Sub DownPressed()
		  Var v As Integer
		  
		  
		  #Pragma BreakOnExceptions False
		  Try
		    v= txtEDDiff1.Text.ToInteger
		  Catch
		    v = 0
		  End
		  #Pragma BreakOnExceptions True
		  
		  v = v - 1
		  
		  txtEDDiff1.Text = v.ToString
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub UpPressed()
		  Var v As Integer
		  
		  
		  #Pragma BreakOnExceptions False
		  Try
		    v= txtEDDiff1.Text.ToInteger
		  Catch
		    v = 0
		  End
		  #Pragma BreakOnExceptions True
		  
		  v = v + 1
		  
		  txtEDDiff1.Text = v.ToString
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnOK
	#tag Event
		Sub Pressed()
		  If txtCaption.Text.Trim = "" Then
		    MessageBox "Enter a caption"
		    Exit Sub
		  end
		  
		  If txtCaption.Text.Trim.Lowercase = "new day" Then
		    
		    Var m as new MessageDialog
		    m.Title = self.Title
		    m.Message = "Do you want to keep ""New day"" as the caption?"
		    m.IconType = MessageDialog.IconTypes.Question
		    m.ActionButton.Caption = "Keep"
		    m.CancelButton.Caption = "Cancel"
		    m.CancelButton.Visible = True
		    m.CancelButton.Default = True
		    m.CancelButton.Cancel = True
		    Var b As MessageDialogButton  = m.ShowModal
		    if b = m.CancelButton then exit sub
		    
		  end
		  
		  Select Case PagePanel1.SelectedPanelIndex
		    
		  Case 0 // Fixed day
		    
		    Var month as Integer = popFDMonth.SelectedRowIndex + 1
		    Var day As Integer = popFDDay.SelectedRowIndex + 1
		    
		    If (month = 4 Or month = 6 Or month = 9 Or month = 11) and day = 31 Then
		      MessageBox "The day was set to 30"
		      day = 30
		    elseif month = 2 and day > 29 then
		      MessageBox "The day was set to 29"
		      day = 29
		    end
		    
		    Var fd As New AnnualEventFix(txtCaption.Text.trim, month, day )
		    
		    If cbFDFridayIfSaturday.Value Then fd.FridayIfSaturday = True
		    if cbFDMondayIfSunday.Value then fd.MondayIfSunday = True
		    
		    
		    If popFDWeekDay.SelectedRowIndex > 0 Then
		      If popFDShiftMode.SelectedRowIndex = 0 Then
		        if chkDFAlways.value then fd.AlwaysPreviousWeekDay = popFDWeekDay.SelectedRowIndex else fd.PreviousWeekDay =  popFDWeekDay.SelectedRowIndex
		      Else
		        if chkDFAlways.value then fd.AlwaysNextWeekDay =  popFDWeekDay.SelectedRowIndex else fd.NextWeekDay=  popFDWeekDay.SelectedRowIndex
		      End
		      
		    ElseIf txtFDAdding.Text.ToInteger <> 0 Then
		      fd.AddDays =  txtFDAdding.Text.ToInteger
		    End
		    
		    
		    fd.DayOff = chkDayOff.Value
		    fd.CycleFirstYear =  txtFirstYearOfTheCycle.Text.ToInteger
		    fd.CycleYearDuration = txtCycleDuration.Text.ToInteger
		    fd.StartOfValidity = dtStart.SelectedDate
		    fd.EndOfValidity = dtEnd.SelectedDate
		    
		    self.AnnualEventReturned = fd
		    
		  case 1 // Easter
		    
		    Var delta as integer = txtEDDiff.Text.ToInteger
		    Var ed as new AnnualEventEaster(txtCaption.Text, delta)
		    
		    ed.DayOff = chkDayOff.Value
		    ed.CycleFirstYear =  txtFirstYearOfTheCycle.Text.ToInteger
		    ed.CycleYearDuration = txtCycleDuration.Text.ToInteger
		    ed.StartOfValidity = dtStart.SelectedDate
		    ed.EndOfValidity = dtEnd.SelectedDate
		    
		    self.AnnualEventReturned = ed
		    
		  Case 2 // Orthodox aster
		    
		    Var delta As Integer = txtEDDiff1.Text.ToInteger
		    Var oed As New AnnualEventOrthodoxEaster(txtCaption.Text, delta)
		    
		    oed.DayOff = chkDayOff.Value
		    oed.CycleFirstYear =  txtFirstYearOfTheCycle.Text.ToInteger
		    oed.CycleYearDuration = txtCycleDuration.Text.ToInteger
		    oed.StartOfValidity = dtStart.SelectedDate
		    oed.EndOfValidity = dtEnd.SelectedDate
		    
		    self.AnnualEventReturned = oed
		    
		  Case 3
		    
		    Var month As Integer = popWDMonth.SelectedRowIndex + 1
		    Var weekday As Integer = popWDWeekDay.SelectedRowIndex + 1
		    var rank as Integer = popWDRank.SelectedRowIndex + 1
		    
		    Var wdd as new AnnualEventWeekDay(txtCaption.Text, month, weekday, rank)
		    
		    
		    If popWDToWeekDay.SelectedRowIndex > 0 Then
		      
		      If popWDShiftMode.SelectedRowIndex = 1 Then
		        wdd.NextWeekDay =  popWDToWeekDay.SelectedRowIndex
		      Else
		        wdd.PreviousWeekDay = popWDToWeekDay.SelectedRowIndex
		      End
		      
		    ElseIf txtWDAdding.Text.ToInteger <> 0 Then
		      
		      wdd.AddDays =  txtWDAdding.Text.ToInteger
		      
		    End
		    
		    wdd.DayOff = chkDayOff.Value
		    wdd.CycleFirstYear =  txtFirstYearOfTheCycle.Text.ToInteger
		    wdd.CycleYearDuration = txtCycleDuration.Text.ToInteger
		    wdd.StartOfValidity = new DateTime(dtStart.SelectedDate.Year, dtStart.SelectedDate.Month, dtStart.SelectedDate.Day)
		    wdd.EndOfValidity = new DateTime(dtEnd.SelectedDate.Year, dtEnd.SelectedDate.Month, dtEnd.SelectedDate.Day, 23, 59, 59)
		    
		    self.AnnualEventReturned = wdd
		    
		  End Select
		  
		  self.Close
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnCancel
	#tag Event
		Sub Pressed()
		  Self.Close
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnCopy
	#tag Event
		Sub Pressed()
		  If txtCaption.Text.Trim = "" Then
		    MessageBox "Enter a caption"
		    Exit Sub
		  end
		  
		  Var strCode as string
		  
		  Select Case PagePanel1.SelectedPanelIndex
		    
		  Case 0 // Fixed day
		    
		    Var month as Integer = popFDMonth.SelectedRowIndex + 1
		    Var day As Integer = popFDDay.SelectedRowIndex + 1
		    
		    If (month = 4 Or month = 6 Or month = 9 Or month = 11) and day = 31 Then
		      day = 30
		    elseif month = 2 and day > 29 then
		      day = 29
		    end
		    
		    strCode = "Var NewEvent As New AnnualEventFix(""" + txtCaption.Text.ReplaceAll("""","""""") + """," + month.ToString + "," + day.ToString + ")"
		    
		    If cbFDFridayIfSaturday.Value Then 
		      strCode = strCode + EndOfLine
		      strCode = strCode + "NewEvent.FridayIfSaturday = True"
		    end
		    
		    If cbFDMondayIfSunday.Value Then 
		      strCode = strCode + EndOfLine
		      strCode = strCode + "NewEvent.MondayIfSunday = True"
		    End If
		    
		    
		    If popFDWeekDay.SelectedRowIndex > 0 Then
		      If popFDShiftMode.SelectedRowIndex = 0 Then
		        strCode = strCode + EndOfLine
		        If chkDFAlways.value Then _
		        strCode = strCode + "NewEvent.AlwaysPreviousWeekDay = " + popFDWeekDay.SelectedRowIndex.ToString _
		        Else _
		        strCode = strCode + "NewEvent.PreviousWeekDay = " + popFDWeekDay.SelectedRowIndex.ToString
		      Else
		        strCode = strCode + EndOfLine
		        If chkDFAlways.value Then _
		        strCode = strCode + "NewEvent.AlwaysNextWeekDay =  " + popFDWeekDay.SelectedRowIndex.ToString _
		        Else _
		        strCode = strCode + "NewEvent.NextWeekDay =  " + popFDWeekDay.SelectedRowIndex.ToString 
		      End
		      
		    ElseIf txtFDAdding.Text.ToInteger <> 0 Then
		      strCode = strCode + EndOfLine
		      strCode = strCode + "NewEvent.AddDays = " +  txtFDAdding.Text
		    End
		    
		    
		  case 1 // Easter
		    
		    Var delta as integer = txtEDDiff.Text.ToInteger
		    strCode = strCode + "Var NewEvent as new AnnualEventEaster(""" + txtCaption.Text.ReplaceAll("""","""""") + """," + delta.ToString + ")"
		    
		    
		  Case 2 // Orthodox aster
		    
		    Var delta As Integer = txtEDDiff1.Text.ToInteger
		    strCode = strCode + "Var NewEvent as new AnnualEventOrthodoxEaster(""" + txtCaption.Text.ReplaceAll("""","""""") + """," + delta.ToString + ")"
		    
		  Case 3
		    
		    Var month As Integer = popWDMonth.SelectedRowIndex + 1
		    Var weekday As Integer = popWDWeekDay.SelectedRowIndex + 1
		    var rank as Integer = popWDRank.SelectedRowIndex + 1
		    
		    strCode = strCode + "Var NewEvent as new (""" + txtCaption.Text.ReplaceAll("""","""""") + """," + month.ToString + "," + weekday.ToString + ","  + rank.ToString + ")"
		    
		    
		    If popWDToWeekDay.SelectedRowIndex > 0 Then
		      
		      If popWDShiftMode.SelectedRowIndex = 0 Then
		        strCode = strCode + EndOfLine
		        strCode = strCode + "NewEvent.PreviousWeekDay = " + popWDWeekDay.SelectedRowIndex.ToString
		      Else
		        strCode = strCode + EndOfLine
		        strCode = strCode + "NewEvent.NextWeekDay =  " + popWDWeekDay.SelectedRowIndex.ToString
		      End
		      
		    ElseIf txtWDAdding.Text.ToInteger <> 0 Then
		      
		      strCode = strCode + EndOfLine
		      strCode = strCode + "NewEvent.AddDays = " +  txtWDAdding.Text
		      
		    End
		    
		  End Select
		  
		  
		  If Not chkDayOff.Value Then
		    strCode = strCode + EndOfLine
		    strCode = strCode + "NewEvent.DayOff = False"
		  End
		  
		  If txtCycleDuration.Text.ToInteger > 1 Then
		    strCode = strCode + EndOfLine
		    strCode = strCode + "NewEvent.CycleFirstYear = " + txtFirstYearOfTheCycle.Text + EndOfLine
		    strCode = strCode + "NewEvent.CycleYearDuration = " +txtCycleDuration.Text + EndOfLine
		  End
		  
		  Var dCompare As new DateTime(1,1,1)
		  
		  
		  If Not (dtStart.SelectedDate.day = dCompare.day _
		    And dtStart.SelectedDate.Month = dCompare.Month _
		    And dtStart.SelectedDate.Year = dCompare.Year) Then
		    
		    strCode = strCode + EndOfLine
		    strCode = strCode + "NewEvent.StartOfValidity = New DateTime(" _ 
		    + dtStart.SelectedDate.Year.ToString _
		    + "," + dtStart.SelectedDate.Month.ToString _
		    + "," + dtStart.SelectedDate.day.ToString + ")"
		    
		  End
		  
		  dCompare = new DateTime(3999,12,31)
		  
		  if Not (dtEnd.SelectedDate.day = dCompare.day _
		    And dtEnd.SelectedDate.Month = dCompare.Month _
		    And dtEnd.SelectedDate.Year = dCompare.Year) Then
		    
		    strCode = strCode + EndOfLine
		    strCode = strCode + "NewEvent.EndValidity = New DateTime(" _ 
		    + dtEnd.SelectedDate.Year.ToString _
		    + "," + dtEnd.SelectedDate.Month.ToString _
		    + "," + dtEnd.SelectedDate.day.ToString + ", 23, 59, 59)"
		    
		  end
		  
		  Var cb As New Clipboard
		  cb.Text = strCode
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag ViewBehavior
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Interfaces"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Size"
		InitialValue="600"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Size"
		InitialValue="400"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumWidth"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumHeight"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumWidth"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumHeight"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Type"
		Visible=true
		Group="Frame"
		InitialValue="0"
		Type="Types"
		EditorType="Enum"
		#tag EnumValues
			"0 - Document"
			"1 - Movable Modal"
			"2 - Modal Dialog"
			"3 - Floating Window"
			"4 - Plain Box"
			"5 - Shadowed Box"
			"6 - Rounded Window"
			"7 - Global Floating Window"
			"8 - Sheet Window"
			"9 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Frame"
		InitialValue="Untitled"
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasCloseButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMaximizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMinimizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasFullScreenButton"
		Visible=true
		Group="Frame"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasTitleBar"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Visible=false
		Group="OS X (Carbon)"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Visible=false
		Group="OS X (Carbon)"
		InitialValue="0"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Visible=true
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="DefaultLocation"
		Visible=true
		Group="Behavior"
		InitialValue="2"
		Type="Locations"
		EditorType="Enum"
		#tag EnumValues
			"0 - Default"
			"1 - Parent Window"
			"2 - Main Screen"
			"3 - Parent Window Screen"
			"4 - Stagger"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Window Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="BackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="&cFFFFFF"
		Type="ColorGroup"
		EditorType="ColorGroup"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Background"
		InitialValue=""
		Type="Picture"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Menus"
		InitialValue=""
		Type="DesktopMenuBar"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Visible=true
		Group="Deprecated"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
#tag EndViewBehavior
