#tag DesktopWindow
Begin DesktopWindow wDemo
   Backdrop        =   0
   BackgroundColor =   &cFFFFFF
   Composite       =   False
   DefaultLocation =   2
   FullScreen      =   False
   HasBackgroundColor=   False
   HasCloseButton  =   True
   HasFullScreenButton=   False
   HasMaximizeButton=   False
   HasMinimizeButton=   True
   HasTitleBar     =   True
   Height          =   742
   ImplicitInstance=   True
   MacProcID       =   0
   MaximumHeight   =   32000
   MaximumWidth    =   32000
   MenuBar         =   776042495
   MenuBarVisible  =   False
   MinimumHeight   =   64
   MinimumWidth    =   64
   Resizeable      =   False
   Title           =   "Demo"
   Type            =   1
   Visible         =   True
   Width           =   886
   Begin DesktopTabPanel TabPanel1
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   712
      Index           =   -2147483648
      Italic          =   False
      Left            =   10
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Panels          =   ""
      Scope           =   0
      SmallTabs       =   False
      TabDefinition   =   "Code demo for a definition\rDates Calculator"
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   10
      Transparent     =   False
      Underline       =   False
      Value           =   0
      Visible         =   True
      Width           =   863
      BeginDesktopSegmentedButton DesktopSegmentedButton SegmentedButton1
         Enabled         =   True
         Height          =   24
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Left            =   16
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         Segments        =   "Fixed day\n\nTrue\rFrom Easter\n\nFalse\rRank WeekDay\n\nFalse"
         SelectionStyle  =   0
         TabIndex        =   0
         TabPanelIndex   =   1
         TabStop         =   False
         Tooltip         =   ""
         Top             =   41
         Transparent     =   False
         Visible         =   True
         Width           =   409
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
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   21
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
         Text            =   "Caption"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   68
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
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   18
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
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "Test Definition"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   91
         Transparent     =   False
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   451
      End
      Begin DesktopTextArea txtCode
         AllowAutoDeactivate=   True
         AllowFocusRing  =   True
         AllowSpellChecking=   False
         AllowStyledText =   False
         AllowTabs       =   False
         BackgroundColor =   &cFFFFFF
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Format          =   ""
         HasBorder       =   True
         HasHorizontalScrollbar=   False
         HasVerticalScrollbar=   True
         Height          =   327
         HideSelection   =   True
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   21
         LineHeight      =   0.0
         LineSpacing     =   1.0
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MaximumCharactersAllowed=   0
         Multiline       =   True
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   11
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   ""
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   352
         Transparent     =   False
         Underline       =   False
         UnicodeMode     =   1
         ValidationMask  =   ""
         Visible         =   True
         Width           =   435
      End
      Begin DesktopListBox ListDatesCode
         AllowAutoDeactivate=   True
         AllowAutoHideScrollbars=   True
         AllowExpandableRows=   False
         AllowFocusRing  =   True
         AllowResizableColumns=   False
         AllowRowDragging=   False
         AllowRowReordering=   False
         Bold            =   False
         ColumnCount     =   2
         ColumnWidths    =   "75%,25%"
         DefaultRowHeight=   35
         DropIndicatorVisible=   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         GridLineStyle   =   1
         HasBorder       =   True
         HasHeader       =   True
         HasHorizontalScrollbar=   False
         HasVerticalScrollbar=   True
         HeadingIndex    =   -1
         Height          =   327
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         InitialValue    =   "Dates	Reverse test"
         Italic          =   False
         Left            =   468
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         RequiresSelection=   False
         RowSelectionType=   0
         Scope           =   0
         TabIndex        =   13
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   352
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   405
         _ScrollOffset   =   0
         _ScrollWidth    =   -1
      End
      Begin DesktopRadioGroup RadioGroup1
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   18
         Horizontal      =   True
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         InitialValue    =   "List of dates\r\n5 years days (WWWW-WWWW)"
         Italic          =   False
         Left            =   468
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedIndex   =   0
         TabIndex        =   10
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   329
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   398
      End
      Begin DesktopListBox ListDates5Years
         AllowAutoDeactivate=   True
         AllowAutoHideScrollbars=   True
         AllowExpandableRows=   False
         AllowFocusRing  =   True
         AllowResizableColumns=   False
         AllowRowDragging=   False
         AllowRowReordering=   False
         Bold            =   False
         ColumnCount     =   2
         ColumnWidths    =   "75%,25%"
         DefaultRowHeight=   35
         DropIndicatorVisible=   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         GridLineStyle   =   1
         HasBorder       =   True
         HasHeader       =   True
         HasHorizontalScrollbar=   False
         HasVerticalScrollbar=   True
         HeadingIndex    =   -1
         Height          =   327
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         InitialValue    =   "Dates	Match ?"
         Italic          =   False
         Left            =   292
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         RequiresSelection=   False
         RowSelectionType=   0
         Scope           =   0
         TabIndex        =   12
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   352
         Transparent     =   False
         Underline       =   False
         Visible         =   False
         Width           =   400
         _ScrollOffset   =   0
         _ScrollWidth    =   -1
      End
      Begin DesktopLabel Label8
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   21
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   9
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "Example code for these parameters"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   329
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   232
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
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   135
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MaximumCharactersAllowed=   0
         Password        =   False
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   5
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "1900"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   301
         Transparent     =   False
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   51
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
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   18
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   4
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "First year of the cycle"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   301
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   113
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
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   353
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MaximumCharactersAllowed=   0
         Password        =   False
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   7
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "0"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   301
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
         InitialParent   =   "TabPanel1"
         Left            =   409
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         PanelIndex      =   0
         Scope           =   0
         TabIndex        =   8
         TabPanelIndex   =   1
         Tooltip         =   ""
         Top             =   301
         Transparent     =   False
         Visible         =   True
         Width           =   14
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
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
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   233
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   6
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "Cycle Duration (years)"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   301
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   117
      End
      Begin DesktopButton btnAddItem
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Adding ///"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   443
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
         Top             =   407
         Transparent     =   False
         Underline       =   False
         Visible         =   False
         Width           =   174
      End
      Begin DesktopLabel lblDBstatuts
         AllowAutoDeactivate=   True
         Bold            =   True
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   27
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
         Text            =   "days.db is connected"
         TextAlignment   =   0
         TextColor       =   &c00800000
         Tooltip         =   ""
         Top             =   31
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   197
      End
      Begin DesktopButton btnLoadingRegionsFromDB
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Loading regions from database"
         Default         =   False
         Enabled         =   False
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   20
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   1
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   51
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   197
      End
      Begin DesktopListBox lstRegions
         AllowAutoDeactivate=   True
         AllowAutoHideScrollbars=   True
         AllowExpandableRows=   False
         AllowFocusRing  =   True
         AllowResizableColumns=   False
         AllowRowDragging=   False
         AllowRowReordering=   False
         Bold            =   False
         ColumnCount     =   2
         ColumnWidths    =   ",60"
         DefaultRowHeight=   -1
         DropIndicatorVisible=   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         GridLineStyle   =   0
         HasBorder       =   True
         HasHeader       =   True
         HasHorizontalScrollbar=   False
         HasVerticalScrollbar=   True
         HeadingIndex    =   -1
         Height          =   347
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         InitialValue    =   "Region	Enabled"
         Italic          =   False
         Left            =   18
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         RequiresSelection=   False
         RowSelectionType=   0
         Scope           =   2
         TabIndex        =   2
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   82
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   405
         _ScrollOffset   =   0
         _ScrollWidth    =   -1
      End
      Begin DesktopButton btnCopyDates
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "&Copy in clipboard"
         Default         =   True
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   743
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   2
         TabIndex        =   14
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   691
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   130
      End
      Begin DesktopTextField txtNewRegion
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
         Hint            =   "New Region"
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   22
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MaximumCharactersAllowed=   0
         Password        =   False
         ReadOnly        =   False
         Scope           =   2
         TabIndex        =   7
         TabPanelIndex   =   2
         TabStop         =   True
         Text            =   ""
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   437
         Transparent     =   False
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   282
      End
      Begin DesktopButton btnNewRegion
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Add this region"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   316
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   2
         TabIndex        =   8
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   437
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   111
      End
      Begin DesktopLabel lblRegionIdentifier
         AllowAutoDeactivate=   True
         Bold            =   True
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   443
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   2
         Selectable      =   False
         TabIndex        =   3
         TabPanelIndex   =   2
         TabStop         =   True
         Text            =   "No region selected"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   79
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   258
      End
      Begin DesktopButton btnEditItem
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Edit ///"
         Default         =   False
         Enabled         =   False
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   629
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
         Top             =   407
         Transparent     =   False
         Underline       =   False
         Visible         =   False
         Width           =   174
      End
      Begin DesktopButton btnNewRegion1
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Add USA (a few states)"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   20
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   2
         TabIndex        =   10
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   462
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   146
      End
      Begin DesktopButton btnRestoreRegionToFactory
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Restore regions in database to factory settings"
         Default         =   False
         Enabled         =   False
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   580
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
         Top             =   48
         Transparent     =   False
         Underline       =   False
         Visible         =   False
         Width           =   276
      End
      Begin DesktopPagePanel PagePanel1
         AllowAutoDeactivate=   True
         Enabled         =   True
         Height          =   164
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Left            =   16
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         PanelCount      =   3
         Panels          =   ""
         Scope           =   0
         SelectedPanelIndex=   0
         TabIndex        =   18
         TabPanelIndex   =   1
         TabStop         =   False
         Tooltip         =   ""
         Top             =   125
         Transparent     =   False
         Value           =   1
         Visible         =   True
         Width           =   849
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
            Left            =   17
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
            Top             =   145
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   165
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
            Left            =   16
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
            Text            =   "Day"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   180
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   43
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
            Left            =   16
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
            Top             =   200
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
            Left            =   16
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   11
            TabPanelIndex   =   1
            TabStop         =   True
            Text            =   "Shift to the                            day of the week"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   235
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
            Left            =   168
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Scope           =   0
            SelectedRowIndex=   -1
            TabIndex        =   13
            TabPanelIndex   =   1
            TabStop         =   True
            Tooltip         =   ""
            Top             =   234
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
            Left            =   345
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   14
            TabPanelIndex   =   1
            TabStop         =   True
            Text            =   "or adding days"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   234
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
            Left            =   429
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MaximumCharactersAllowed=   0
            Password        =   False
            ReadOnly        =   False
            Scope           =   0
            TabIndex        =   15
            TabPanelIndex   =   1
            TabStop         =   True
            Text            =   "0"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   234
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
            Left            =   485
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            PanelIndex      =   0
            Scope           =   0
            TabIndex        =   16
            TabPanelIndex   =   1
            Tooltip         =   ""
            Top             =   234
            Transparent     =   False
            Visible         =   True
            Width           =   14
            _mIndex         =   0
            _mInitialParent =   ""
            _mName          =   ""
            _mPanelIndex    =   0
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
            Left            =   258
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Scope           =   0
            TabIndex        =   3
            TabPanelIndex   =   1
            TabStop         =   True
            Tooltip         =   ""
            Top             =   145
            Transparent     =   False
            Underline       =   False
            Value           =   False
            Visible         =   True
            VisualState     =   0
            Width           =   142
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
            Left            =   415
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
            Top             =   145
            Transparent     =   False
            Underline       =   False
            Value           =   False
            Visible         =   True
            VisualState     =   0
            Width           =   158
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
            Left            =   18
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
            Top             =   139
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
            Left            =   180
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
            Top             =   139
            Transparent     =   False
            Underline       =   False
            ValidationMask  =   ""
            Visible         =   True
            Width           =   51
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
            Left            =   236
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            PanelIndex      =   0
            Scope           =   0
            TabIndex        =   2
            TabPanelIndex   =   2
            Tooltip         =   ""
            Top             =   139
            Transparent     =   False
            Visible         =   True
            Width           =   14
            _mIndex         =   0
            _mInitialParent =   ""
            _mName          =   ""
            _mPanelIndex    =   0
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
            Left            =   432
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
            Top             =   267
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
            Left            =   432
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
            Top             =   203
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
            Left            =   585
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
            Top             =   203
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
            Left            =   585
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
            Top             =   235
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
            Left            =   738
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
            Top             =   139
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
            Left            =   738
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
            Top             =   171
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
            Left            =   585
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   20
            TabPanelIndex   =   2
            TabStop         =   True
            Tooltip         =   ""
            Top             =   267
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
            Left            =   278
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
            Top             =   171
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
            Left            =   432
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
            Top             =   139
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
            Left            =   278
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
            Top             =   203
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
            Left            =   278
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
            Top             =   235
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
            Left            =   432
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
            Top             =   171
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
            Left            =   432
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
            Top             =   235
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
            Left            =   585
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
            Top             =   139
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
            Left            =   585
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
            Top             =   171
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
            Left            =   278
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
            Top             =   267
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   125
         End
         Begin DesktopLabel Label7
            AllowAutoDeactivate=   True
            Bold            =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   122
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   18
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   True
            Scope           =   2
            Selectable      =   False
            TabIndex        =   11
            TabPanelIndex   =   2
            TabStop         =   True
            Text            =   "You also have a AnnualEventOrthodoxEaster class.\r\nThe date of Easter is the Orthodox date, but it's actually the date in the Gregorian calendar that's given (because Xojo works with this calendar!)"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   168
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   223
         End
         Begin DesktopLabel Label9
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
            Left            =   16
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
            Text            =   "Month"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   125
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   43
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
            Left            =   16
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Scope           =   0
            SelectedRowIndex=   -1
            TabIndex        =   5
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   146
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   165
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
            Left            =   16
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   10
            TabPanelIndex   =   3
            TabStop         =   True
            Text            =   "Day of week"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   180
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
            Left            =   16
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Scope           =   0
            SelectedRowIndex=   -1
            TabIndex        =   14
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   202
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   165
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
            Left            =   16
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   17
            TabPanelIndex   =   3
            TabStop         =   True
            Text            =   "Rank"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   236
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   43
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
            Left            =   16
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Scope           =   0
            SelectedRowIndex=   -1
            TabIndex        =   18
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   258
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
            Left            =   209
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   19
            TabPanelIndex   =   3
            TabStop         =   True
            Text            =   "Shift to the                            day of the week"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   258
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
            Left            =   360
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Scope           =   0
            SelectedRowIndex=   -1
            TabIndex        =   21
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   258
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
            Left            =   537
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   22
            TabPanelIndex   =   3
            TabStop         =   True
            Text            =   "or adding days"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   258
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
            Left            =   626
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MaximumCharactersAllowed=   0
            Password        =   False
            ReadOnly        =   False
            Scope           =   0
            TabIndex        =   23
            TabPanelIndex   =   3
            TabStop         =   True
            Text            =   "0"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   258
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
            Left            =   682
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            PanelIndex      =   0
            Scope           =   0
            TabIndex        =   24
            TabPanelIndex   =   3
            Tooltip         =   ""
            Top             =   258
            Transparent     =   False
            Visible         =   True
            Width           =   14
            _mIndex         =   0
            _mInitialParent =   ""
            _mName          =   ""
            _mPanelIndex    =   0
         End
         Begin DesktopButton ButtonWD1
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Presidential Election"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   212
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   15
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   212
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   125
         End
         Begin DesktopLabel Label16
            AllowAutoDeactivate=   True
            Bold            =   True
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   20
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   212
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   1
            TabPanelIndex   =   3
            TabStop         =   True
            Text            =   "U.S.A."
            TextAlignment   =   2
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   126
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   125
         End
         Begin DesktopButton ButtonWD2
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Labor Day"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   212
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   6
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   146
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   125
         End
         Begin DesktopButton ButtonWD3
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Thanksgiving Day "
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   212
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   11
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   178
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   125
         End
         Begin DesktopButton ButtonWD4
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Constitución de 1917"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   360
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   7
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   "Proclamación de la Constitución de 1917"
            Top             =   146
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   125
         End
         Begin DesktopLabel Label17
            AllowAutoDeactivate=   True
            Bold            =   True
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   20
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   360
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   2
            TabPanelIndex   =   3
            TabStop         =   True
            Text            =   "Mexico"
            TextAlignment   =   2
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   126
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   125
         End
         Begin DesktopButton ButtonWD5
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Aniv. Benito Juárez"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   360
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   12
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   180
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   125
         End
         Begin DesktopButton ButtonWD6
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Día de la Revolución"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   360
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   16
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   212
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   125
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
            Left            =   274
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Scope           =   0
            SelectedRowIndex=   -1
            TabIndex        =   20
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   258
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   74
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
            Left            =   82
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Scope           =   0
            SelectedRowIndex=   -1
            TabIndex        =   12
            TabPanelIndex   =   1
            TabStop         =   True
            Tooltip         =   ""
            Top             =   234
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   74
         End
         Begin DesktopLabel Label18
            AllowAutoDeactivate=   True
            Bold            =   True
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   20
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   506
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   3
            TabPanelIndex   =   3
            TabStop         =   True
            Text            =   "Canada"
            TextAlignment   =   2
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   126
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   170
         End
         Begin DesktopButton ButtonWD7
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Labour Day (Fête du travail)"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   506
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   8
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   146
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   170
         End
         Begin DesktopButton ButtonWD8
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Thanksgiving (Action de grâce)"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   11.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   506
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   13
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   180
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   170
         End
         Begin DesktopButton ButtonWD9
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Fête de la Wallonie"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   688
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   9
            TabPanelIndex   =   3
            TabStop         =   True
            Tooltip         =   ""
            Top             =   145
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   170
         End
         Begin DesktopLabel Label19
            AllowAutoDeactivate=   True
            Bold            =   True
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   20
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   688
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   4
            TabPanelIndex   =   3
            TabStop         =   True
            Text            =   "Belgium"
            TextAlignment   =   2
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   126
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   170
         End
         Begin DesktopButton ButtonWD10
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Victoria Day ou Fête des patriotes"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   655
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   8
            TabPanelIndex   =   1
            TabStop         =   True
            Tooltip         =   ""
            Top             =   195
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   195
         End
         Begin DesktopButton ButtonWD11
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Independence day"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   655
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   1
            TabPanelIndex   =   1
            TabStop         =   True
            Tooltip         =   ""
            Top             =   139
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   195
         End
         Begin DesktopButton ButtonWD12
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "29 of february"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   655
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   17
            TabPanelIndex   =   1
            TabStop         =   True
            Tooltip         =   ""
            Top             =   251
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   195
         End
         Begin DesktopButton ButtonWD13
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Día de la Independencia"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   655
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   5
            TabPanelIndex   =   1
            TabStop         =   True
            Tooltip         =   ""
            Top             =   167
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   195
         End
         Begin DesktopButton ButtonWD14
            AllowAutoDeactivate=   True
            Bold            =   False
            Cancel          =   False
            Caption         =   "Fête nationale Belge"
            Default         =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   22
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   655
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            MacButtonStyle  =   0
            Scope           =   0
            TabIndex        =   10
            TabPanelIndex   =   1
            TabStop         =   True
            Tooltip         =   ""
            Top             =   223
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   195
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
            Left            =   278
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
            Top             =   139
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   125
         End
         Begin DesktopLabel Label1
            AllowAutoDeactivate=   True
            Bold            =   False
            Enabled         =   True
            FontName        =   "System"
            FontSize        =   0.0
            FontUnit        =   0
            Height          =   17
            Index           =   -2147483648
            InitialParent   =   "PagePanel1"
            Italic          =   False
            Left            =   17
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   18
            TabPanelIndex   =   1
            TabStop         =   True
            Text            =   "Month"
            TextAlignment   =   0
            TextColor       =   &c000000
            Tooltip         =   ""
            Top             =   125
            Transparent     =   False
            Underline       =   False
            Visible         =   True
            Width           =   43
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
            Left            =   168
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Scope           =   0
            TabIndex        =   19
            TabPanelIndex   =   1
            TabStop         =   True
            Tooltip         =   ""
            Top             =   264
            Transparent     =   False
            Underline       =   False
            Value           =   False
            Visible         =   True
            VisualState     =   0
            Width           =   100
         End
      End
      Begin DesktopListBox lstRegionItems
         AllowAutoDeactivate=   True
         AllowAutoHideScrollbars=   True
         AllowExpandableRows=   False
         AllowFocusRing  =   True
         AllowResizableColumns=   False
         AllowRowDragging=   False
         AllowRowReordering=   False
         Bold            =   False
         ColumnCount     =   3
         ColumnWidths    =   ",25%,0"
         DefaultRowHeight=   -1
         DropIndicatorVisible=   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         GridLineStyle   =   0
         HasBorder       =   True
         HasHeader       =   True
         HasHorizontalScrollbar=   False
         HasVerticalScrollbar=   True
         HeadingIndex    =   -1
         Height          =   258
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         InitialValue    =   "Week Day	Type"
         Italic          =   False
         Left            =   443
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         RequiresSelection=   False
         RowSelectionType=   0
         Scope           =   2
         TabIndex        =   12
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   140
         Transparent     =   False
         Underline       =   False
         Visible         =   False
         Width           =   413
         _ScrollOffset   =   0
         _ScrollWidth    =   -1
      End
      Begin DesktopDateTimePicker dtLast
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
         InitialParent   =   "TabPanel1"
         Left            =   143
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         PanelIndex      =   0
         Scope           =   0
         TabIndex        =   13
         TabPanelIndex   =   2
         TodayButtonCaption=   ""
         Tooltip         =   ""
         Top             =   543
         Transparent     =   False
         Visible         =   True
         Width           =   113
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopDateTimePicker dtFirst
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
         InitialParent   =   "TabPanel1"
         Left            =   18
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         PanelIndex      =   0
         Scope           =   0
         TabIndex        =   12
         TabPanelIndex   =   2
         TodayButtonCaption=   ""
         Tooltip         =   ""
         Top             =   543
         Transparent     =   False
         Visible         =   True
         Width           =   113
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopLabel Label21
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   17
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   143
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   14
         TabPanelIndex   =   2
         TabStop         =   True
         Text            =   "Last day"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   524
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   113
      End
      Begin DesktopLabel Label20
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   17
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   18
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   15
         TabPanelIndex   =   2
         TabStop         =   True
         Text            =   "First day"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   524
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   113
      End
      Begin DesktopListBox lstResults
         AllowAutoDeactivate=   True
         AllowAutoHideScrollbars=   True
         AllowExpandableRows=   False
         AllowFocusRing  =   True
         AllowResizableColumns=   False
         AllowRowDragging=   False
         AllowRowReordering=   False
         Bold            =   False
         ColumnCount     =   1
         ColumnWidths    =   ""
         DefaultRowHeight=   -1
         DropIndicatorVisible=   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         GridLineStyle   =   0
         HasBorder       =   True
         HasHeader       =   True
         HasHorizontalScrollbar=   False
         HasVerticalScrollbar=   True
         HeadingIndex    =   -1
         Height          =   165
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         InitialValue    =   "Results"
         Italic          =   False
         Left            =   488
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         RequiresSelection=   False
         RowSelectionType=   0
         Scope           =   0
         TabIndex        =   16
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   496
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   300
         _ScrollOffset   =   0
         _ScrollWidth    =   -1
      End
      Begin DesktopRadioGroup rgDaysFunctions
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   26
         Horizontal      =   True
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         InitialValue    =   "For all regions marked as enabled\r\nFor the selected region"
         Italic          =   False
         Left            =   18
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedIndex   =   0
         TabIndex        =   17
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   496
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   358
      End
      Begin DesktopButton Button1
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Working days"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   268
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
         Top             =   543
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   90
      End
      Begin DesktopButton Button2
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Days off"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   363
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
         Top             =   543
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   90
      End
      Begin DesktopDateTimePicker dtThisDate
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
         InitialParent   =   "TabPanel1"
         Left            =   22
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         PanelIndex      =   0
         Scope           =   0
         TabIndex        =   20
         TabPanelIndex   =   2
         TodayButtonCaption=   ""
         Tooltip         =   ""
         Top             =   626
         Transparent     =   False
         Visible         =   True
         Width           =   113
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopLabel Label22
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   17
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   22
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   21
         TabPanelIndex   =   2
         TabStop         =   True
         Text            =   "This date"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   607
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   113
      End
      Begin DesktopLabel lblEvent
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   147
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   22
         TabPanelIndex   =   2
         TabStop         =   True
         Text            =   "No event"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   649
         Transparent     =   False
         Underline       =   False
         Visible         =   False
         Width           =   121
      End
      Begin DesktopCheckBox chkIsDayOff
         AllowAutoDeactivate=   True
         Bold            =   False
         Caption         =   "Is Day off"
         Enabled         =   False
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   147
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         TabIndex        =   23
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   626
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         VisualState     =   0
         Width           =   72
      End
      Begin DesktopCheckBox chkWorkingDay
         AllowAutoDeactivate=   True
         Bold            =   False
         Caption         =   "Is Working day"
         Enabled         =   False
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   231
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         TabIndex        =   24
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   626
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         VisualState     =   0
         Width           =   100
      End
      Begin DesktopCheckBox chkIsEvent
         AllowAutoDeactivate=   True
         Bold            =   False
         Caption         =   "Is Event"
         Enabled         =   False
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   343
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         TabIndex        =   25
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   626
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         VisualState     =   0
         Width           =   100
      End
      Begin DesktopTextField txtFollowingBD
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
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   147
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MaximumCharactersAllowed=   0
         Password        =   False
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   26
         TabPanelIndex   =   2
         TabStop         =   True
         Text            =   "1"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   690
         Transparent     =   False
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   29
      End
      Begin DesktopUpDownArrows UpDownArrows5
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         Enabled         =   True
         Height          =   23
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Left            =   178
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         PanelIndex      =   0
         Scope           =   0
         TabIndex        =   27
         TabPanelIndex   =   2
         Tooltip         =   ""
         Top             =   690
         Transparent     =   False
         Visible         =   True
         Width           =   14
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopLabel lblNextBusinessDay
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   202
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   28
         TabPanelIndex   =   2
         TabStop         =   True
         Text            =   "following business day is"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   690
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   294
      End
      Begin DesktopButton btnConnectDB
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "Connect to the db"
         Default         =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   22
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   231
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   33
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   38
         Transparent     =   False
         Underline       =   False
         Visible         =   False
         Width           =   115
      End
      Begin DesktopRadioGroup rgItem
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   18
         Horizontal      =   True
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         InitialValue    =   "Working week day\r\nClosure periods\r\nAnnual events"
         Italic          =   False
         Left            =   443
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedIndex   =   0
         TabIndex        =   34
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   110
         Transparent     =   False
         Underline       =   False
         Visible         =   False
         Width           =   418
      End
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Event
		Sub Opening()
		  
		  
		  Var d as DateTime = DateTime.Now
		  
		  d = new DateTime(d.Year - 3, 1, 1)
		  Var max as integer = d.Year + 5
		  
		  RadioGroup1.ItemAt(1).Caption = "5 years days (" + d.Year.ToString + "-" + max.ToString + ")"
		  
		  
		  do until d.Year = max
		    ListDates5Years.AddRow d.ToString("EEEE", Locale.Current) + " " + d.ToString(DateTime.FormatStyles.Long, DateTime.FormatStyles.None)
		    ListDates5Years.RowTagAt(ListDates5Years.LastAddedRowIndex) = d
		    
		    d = d.AddInterval(0,0,1)
		  loop
		  
		  ListDates5Years.top = ListDatesCode.Top
		  ListDates5Years.Left = ListDatesCode.Left
		  
		  FixedDayCode
		  
		  If app.db = Nil Then
		    lblDBstatuts.Text = "days.db is NOT connected"
		    lblDBstatuts.TextColor = &c80000000
		    btnConnectDB.Visible = True
		  Else
		    btnLoadingRegionsFromDB.Enabled = True
		    btnRestoreRegionToFactory.Enabled = True
		  end
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub CalculateThisDate()
		  app.MouseCursor = System.Cursors.Wait
		  
		  If rgDaysFunctions.SelectedIndex = 1 And lstRegions.SelectedRowIndex <> DesktopListBox.NoSelection Then
		    
		    chkWorkingDay.Value = Self.Regions(lstRegions.SelectedRowIndex).IsWorkingDay(dtThisDate.SelectedDate)
		    chkIsDayOff.value = Not chkWorkingDay.Value
		    
		    Var dc As DateAndCaption = Self.Regions(lstRegions.SelectedRowIndex).IsAnnualEventDateAndCaption(dtThisDate.SelectedDate)
		    If dc <> Nil Then 
		      chkIsEvent.Value = True
		      lblEvent.Text = dc.Caption
		    Else
		      chkIsEvent.Value = False
		      lblEvent.Text = "No event"
		    End
		    
		    Var nbd As DateTime = Self.regions(lstRegions.SelectedRowIndex).NextBusinessDay(dtThisDate.SelectedDate, txtFollowingBD.Text.ToInteger)
		    lblNextBusinessDay.Text = "following business day is " + nbd.ToString(DateTime.FormatStyles.Long, DateTime.FormatStyles.None)
		    
		  elseif rgDaysFunctions.SelectedIndex = 1 And lstRegions.SelectedRowIndex = DesktopListBox.NoSelection Then
		    
		  ElseIf rgDaysFunctions.SelectedIndex = 0 And Self.Regions.count > 0 Then
		    
		    // Creating a muli-region object
		    // We could also create it for the window and work on its Regions array
		    
		    Var mr As New MultiRegionDatesWorked
		    mr.Regions = Self.Regions
		    
		    chkWorkingDay.Value = mr.IsWorkingDay(dtThisDate.SelectedDate)
		    chkIsDayOff.Value = Not chkWorkingDay.Value
		    chkIsEvent.Value = mr.IsAnnualEvent(dtThisDate.SelectedDate)
		    
		    Var nbd As DateTime = mr.NextBusinessDay(dtThisDate.SelectedDate, txtFollowingBD.Text.ToInteger)
		    lblNextBusinessDay.Text = "following business day is " + nbd.ToString(DateTime.FormatStyles.Long, DateTime.FormatStyles.None)
		    
		    
		    
		  End
		  
		  
		  
		  
		  app.MouseCursor = nil
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub FixedDayCode()
		  Var s As String
		  s = "Var dFixedDay As New AnnualEventFix("
		  s = s + """" + txtCaption.text.ReplaceAll("""", """""") + ""","
		  s = s + CStr(popFDMonth.SelectedRowIndex + 1) + "," +CStr(popFDDay.SelectedRowIndex + 1) 
		  s = s + ")"
		  
		  If cbFDFridayIfSaturday.Value Then
		    s = s + EndOfLine
		    s = s + "dFixedDay.FridayIfSaturday = True" 
		  end
		  
		  if cbFDMondayIfSunday.Value then
		    s = s + EndOfLine
		    s = s + "dFixedDay.MondayIfSunday = True" 
		  end
		  
		  if popFDWeekDay.SelectedRowIndex > 0 then
		    if popFDShiftMode.SelectedRowIndex = 0 then
		      s = s + EndOfLine
		      If chkDFAlways.value Then 
		        s = s + "dFixedDay.AlwaysPreviousWeekDay = " + popFDWeekDay.SelectedRowIndex.ToString
		      else
		        s = s + "dFixedDay.PreviousWeekDay = " + popFDWeekDay.SelectedRowIndex.ToString
		      end
		    else
		      s = s + EndOfLine
		      If chkDFAlways.value Then 
		        s = s + "dFixedDay.AlwaysNextWeekDay = " + popFDWeekDay.SelectedRowIndex.ToString
		      else
		        s = s + "dFixedDay.NextWeekDay = " + popFDWeekDay.SelectedRowIndex.ToString
		      end
		    end
		    
		  elseif txtFDAdding.Text.ToInteger <> 0 then
		    s = s + EndOfLine
		    s = s + "dFixedDay.AddDays = " + txtFDAdding.Text.ToInteger.ToString
		  end
		  
		  if me.Freeze then exit sub
		  
		  
		  If txtCycleDuration.Text.ToInteger > 1 Then
		    s = s + EndOfLine
		    s = s + "dFixedDay.CycleFirstYear = " + txtFirstYearOfTheCycle.Text.ToInteger.ToString + "  // The event is not every year" + EndOfLine
		    s = s + "dFixedDay.CycleYearDuration = " + txtCycleDuration.Text.ToInteger.ToString
		  end
		  
		  s = s  + EndOfLine + EndOfLine + "// You can also set the DayOff, StartOfValidity and EndOfValidity properties"
		  
		  Var d as DateTime = DateTime.Now
		  
		  s = s + EndOfLine + EndOfLine
		  s = s + " // Determine the date for the year " + d.Year.ToString + EndOfLine
		  s = s + "Var Value" + d.Year.ToString +" as Datetime = dFixFixedDay.DateValue(" + d.Year.ToString + ")"
		  
		  s = s + EndOfLine + EndOfLine
		  s = s + "// Adding the definition to a RegionDatesWorked"
		  s = s + EndOfLine + "// Can use the RegionDatesWorked for calculate bussiness days or a list of holydays for a region"
		  s = s + EndOfLine + "// If you want to process data for multiple regions, use the MultiRegionDatesWorked class."
		  s = s + EndOfLine + "// If the event is not a holiday, set the DayOff property to False"
		  
		  s = s + EndOfLine
		  s = s + EndOfLine + "Var DatesEngine as new RegionDatesWorked"
		  s = s + EndOfLine + "DatesEngine.AnnualEvents.add dFixedDay"
		  
		  
		  txtCode.Text = s  
		  
		  ListDatesCode.RemoveAllRows
		  
		  Var dFixedDay As New AnnualEventFix(txtCaption.Text, popFDMonth.SelectedRowIndex + 1, popFDDay.SelectedRowIndex + 1)
		  If cbFDFridayIfSaturday.Value Then
		    dFixedDay.FridayIfSaturday = True
		  end
		  
		  If cbFDMondayIfSunday.Value Then
		    dFixedDay.MondayIfSunday = True
		  end
		  
		  if popFDWeekDay.SelectedRowIndex > 0 then
		    if popFDShiftMode.SelectedRowIndex = 0 then
		      If chkDFAlways.value Then dFixedDay.AlwaysPreviousWeekDay =  popFDWeekDay.SelectedRowIndex else dFixedDay.PreviousWeekDay =  popFDWeekDay.SelectedRowIndex
		    else
		      If chkDFAlways.value Then dFixedDay.AlwaysNextWeekDay =  popFDWeekDay.SelectedRowIndex else dFixedDay.NextWeekDay =  popFDWeekDay.SelectedRowIndex
		    end
		    
		  elseif txtFDAdding.Text.ToInteger <> 0 Then
		    dFixedDay.AddDays =  txtFDAdding.Text.ToInteger
		  End
		  
		  If txtCycleDuration.Text.ToInteger > 1 Then
		    dFixedDay.CycleFirstYear =  txtFirstYearOfTheCycle.Text.ToInteger
		    dFixedDay.CycleYearDuration = txtCycleDuration.Text.ToInteger
		  end
		  
		  
		  var dt as DateTime
		  
		  app.MouseCursor = System.Cursors.Wait
		  
		  if RadioGroup1.SelectedIndex = 0 then
		    
		    For i As Integer = 1995 To 2035
		      dt = dFixedDay.DateValue(i) 
		      If dt = Nil Then Continue
		      ListDatesCode.AddRow dt.ToString("EEEE", Locale.Current) + " " + dt.ToString(DateTime.FormatStyles.Long, DateTime.FormatStyles.None)
		      
		      If dFixedDay.TestDate(dt) Then
		        ListDatesCode.CellTextAt(ListDatesCode.LastAddedRowIndex, 1) = "OK !"
		      else
		        ListDatesCode.CellTextAt(ListDatesCode.LastAddedRowIndex, 1) = "ERROR !"
		      end
		      
		      ListDatesCode.RowTagAt(ListDatesCode.LastAddedRowIndex) = dt
		      
		    Next
		    
		    
		  Else
		    
		    for row as Integer = 0 to ListDates5Years.LastRowIndex
		      
		      
		      If dFixedDay.TestDate(ListDates5Years.RowTagAt(row).DateTimeValue) then
		        ListDates5Years.CellTextAt(row, 1) = "Match"
		      Else
		        ListDates5Years.CellTextAt(row, 1) = "No Match"
		      End
		      
		    next 
		    
		  End
		  
		  app.MouseCursor = Nil
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub FromEasterCode()
		  Var s As String
		  s = "Var dEasterDay As New AnnualEventEaster("
		  s = s + """" + txtCaption.text.ReplaceAll("""", """""") + ""","
		  s = s + CStr(txtEDDiff.Text) + ")"
		  
		  if me.Freeze then exit sub
		  
		  If txtCycleDuration.Text.ToInteger > 1 Then
		    s = s + EndOfLine
		    s = s + "dEasterDay.CycleFirstYear = " + txtFirstYearOfTheCycle.Text.ToInteger.ToString + "  // The event is not every year" + EndOfLine
		    s = s + "dEasterDay.CycleYearDuration = " + txtCycleDuration.Text.ToInteger.ToString
		  End
		  
		  s = s  + EndOfLine + EndOfLine + "// You can also set the DayOff, StartOfValidity and EndOfValidity properties"
		  
		  Var d as DateTime = DateTime.Now
		  
		  s = s + EndOfLine + EndOfLine
		  s = s + " // Determine the date for the year " + d.Year.ToString + EndOfLine
		  s = s + "Var Value" + d.Year.ToString +" as Datetime = dEasterDay.DateValue(" + d.Year.ToString + ")"
		  
		  s = s + EndOfLine + EndOfLine
		  s = s + "// Adding the definition to a RegionDatesWorked"
		  s = s + EndOfLine + "// Can use the RegionDatesWorked for calculate bussiness days or a list of holydays for a region"
		  s = s + EndOfLine + "// If you want to process data for multiple regions, use the MultiRegionDatesWorked class."
		  s = s + EndOfLine + "// If the event is not a holiday, set the DayOff property to False"
		  s = s + EndOfLine
		  s = s + EndOfLine + "Var DatesEngine as new RegionDatesWorked"
		  s = s + EndOfLine + "DatesEngine.AnnualEvents.add dEasterDay"
		  
		  
		  txtCode.Text = s  
		  
		  ListDatesCode.RemoveAllRows
		  
		  Var dEasterDay As New AnnualEventEaster(txtCaption.Text, txtEDDiff.Text.ToInteger)
		  
		  If txtCycleDuration.Text.ToInteger > 1 Then
		    dEasterDay.CycleFirstYear =  txtFirstYearOfTheCycle.Text.ToInteger
		    dEasterDay.CycleYearDuration = txtCycleDuration.Text.ToInteger
		  End
		  
		  
		  var dt as DateTime
		  
		  app.MouseCursor = System.Cursors.Wait
		  
		  if RadioGroup1.SelectedIndex = 0 then
		    
		    For i As Integer = 1995 To 2035
		      dt = dEasterDay.DateValue(i) 
		      If dt = Nil Then Continue
		      ListDatesCode.AddRow dt.ToString("EEEE", Locale.Current) + " " + dt.ToString(DateTime.FormatStyles.Long, DateTime.FormatStyles.None)
		      
		      If dEasterDay.TestDate(dt) Then
		        ListDatesCode.CellTextAt(ListDatesCode.LastAddedRowIndex, 1) = "OK !"
		      else
		        ListDatesCode.CellTextAt(ListDatesCode.LastAddedRowIndex, 1) = "ERROR !"
		      end
		      
		      ListDatesCode.RowTagAt(ListDatesCode.LastAddedRowIndex) = dt
		      
		    Next
		    
		    
		  Else
		    
		    for row as Integer = 0 to ListDates5Years.LastRowIndex
		      
		      
		      If dEasterDay.TestDate(ListDates5Years.RowTagAt(row).DateTimeValue) then
		        ListDates5Years.CellTextAt(row, 1) = "Match"
		      Else
		        ListDates5Years.CellTextAt(row, 1) = "No Match"
		      End
		      
		    next 
		    
		  End
		  
		  app.MouseCursor = Nil
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadingRegions()
		  If app.db = Nil Then Exit Sub
		  if not app.db.IsConnected then exit sub
		  
		  Var rs As RowSET
		  
		  Try
		    rs = app.db.SelectSQL("Select * FROM tregions ORDER BY identifier")
		  Catch err As DataBaseException
		    MessageBox "(" + Err.ErrorNumber.ToString + ") " + err.Message
		    Break
		    Exit Sub
		  End Try
		  
		  Me.Regions.RemoveAll
		  Me.Regions = MultiRegionDatesWorked.LoadRegions(rs)
		  
		  if me.Regions.Count = 0 then exit sub
		  
		  // Loading events for all regions
		  Try
		    rs = app.db.SelectSQL("Select * FROM tannualevents ORDER BY region")
		  Catch err As DataBaseException
		    MessageBox "(" + Err.ErrorNumber.ToString + ") " + err.Message
		    Break
		    Exit Sub
		  End Try
		  
		  MultiRegionDatesWorked.LoadAnnualEventsFromRowSet(me.Regions, rs)
		  
		  
		  // Loading Closure periofs for all regions
		  
		  Try
		    rs = app.db.SelectSQL("Select * FROM tclosureperiods ORDER BY region")
		  Catch err As DataBaseException
		    MessageBox "(" + Err.ErrorNumber.ToString + ") " + err.Message
		    Break
		    Exit Sub
		  End Try
		  
		  MultiRegionDatesWorked.LoadClosurePeriodsFromRowSet(me.Regions, rs)
		  
		  PopulateLstRegions
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub PopulateLstRegionItems()
		  Select Case rgItem.SelectedIndex
		    
		  Case 0
		    
		    lstRegionItems.RemoveAllRows
		    If lstRegions.SelectedRowIndex = DesktopListBox.NoSelection Then Exit Sub
		    
		    Var d as new DateTime(2025,5,3)
		    
		    for i as integer = 1 to 7
		      d = d.AddInterval(0, 0, 1)
		      lstRegionItems.AddRow d.ToString("EEEE")
		      lstRegionItems.CellCheckBoxValueAt(lstRegionItems.LastAddedRowIndex, 1) = me.Regions(lstRegions.SelectedRowIndex).WorkingWeekDays.WorkingDay(i)
		    next
		    
		    
		  Case 1
		    
		    lstRegionItems.RemoveAllRows
		    If lstRegionItems.SelectedRowIndex = DesktopListBox.NoSelection Then Exit Sub
		    
		    Var r As RegionDatesWorked = Me.Regions(lstRegions.SelectedRowIndex)
		    
		    For p As Integer = 0 To r.ClosurePeriods.LastIndex
		      
		      lstRegionItems.AddRow r.ClosurePeriods(p).Caption, _
		      r.ClosurePeriods(p).FirstDay.ToString(DateTime.FormatStyles.Short, DateTime.FormatStyles.None), _
		      r.ClosurePeriods(p).LastDay.ToString(DateTime.FormatStyles.Short, DateTime.FormatStyles.None)
		      
		      lstRegionItems.CellTagAt(lstRegionItems.LastAddedRowIndex, 1) = r.ClosurePeriods(p).FirstDay
		      lstRegionItems.CellTagAt(lstRegionItems.LastAddedRowIndex, 2) = r.ClosurePeriods(p).LastDay
		      
		    next p
		    
		    
		  Case 2
		    
		    
		    lstRegionItems.RemoveAllRows
		    If lstRegions.SelectedRowIndex = DesktopListBox.NoSelection Then Exit Sub
		    
		    Var r As RegionDatesWorked = Me.Regions(lstRegions.SelectedRowIndex)
		    
		    If r.AnnualEvents.Count = 0 Then 
		      Break
		      Exit Sub
		    End
		    
		    For e As Integer = 0 To r.AnnualEvents.LastIndex
		      
		      lstRegionItems.AddRow r.AnnualEvents(e).Caption
		      Select Case r.AnnualEvents(e) 
		        
		      Case IsA AnnualEventFix 
		        lstRegionItems.CellTextAt(lstRegionItems.LastAddedRowIndex, 1) = "Fix"
		      Case IsA AnnualEventEaster
		        lstRegionItems.CellTextAt(lstRegionItems.LastAddedRowIndex, 1) = "Easter"
		      Case IsA AnnualEventOrthodoxEaster
		        lstRegionItems.CellTextAt(lstRegionItems.LastAddedRowIndex, 1) = "O.Easter"
		      Case IsA AnnualEventWeekDay
		        lstRegionItems.CellTextAt(lstRegionItems.LastAddedRowIndex, 1) = "WeekDay"
		      End
		      
		      
		    Next
		    
		    
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub PopulateLstRegions()
		  lstRegions.RemoveAllRows
		  
		  If Me.Regions.Count > 0 Then
		    
		    For i As Integer = 0 To Me.Regions.LastIndex
		      If Me.Regions(i).Identifier.StringValue.Trim <>"" Then
		        lstRegions.AddRow Me.Regions(i).Identifier.StringValue.Lowercase
		        lstRegions.CellCheckBoxValueAt(lstRegions.LastAddedRowIndex, 1) = True
		        'popRegions.AddRow me.Regions(i).Identifier.StringValue.Lowercase
		      Else
		        lstRegions.AddRow "RegionDatesWorked without identifier"
		        lstRegions.CellCheckBoxValueAt(lstRegions.LastAddedRowIndex, 1) = True
		        'popRegions.AddRow "RegionDatesWorked without identifier"
		      End
		    Next
		    
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub WeekDayCode()
		  if me.Freeze then exit sub
		  
		  
		  Var s As String
		  s = "Var dWeekDay As New AnnualEventWeekDay("
		  s = s + """" + txtCaption.Text.ReplaceAll("""", """""") + ""","
		  s = s + CStr(popWDMonth.SelectedRowIndex + 1) + "," + CStr(popWDWeekDay.SelectedRowIndex + 1) + "," + CStr(popWDRank.SelectedRowIndex + 1 ) + ")"
		  
		  
		  
		  If popWDToWeekDay.SelectedRowIndex > 0 Then
		    s = s + EndOfLine
		    if popWDShiftMode.SelectedRowIndex = 1 then
		      s = s + "dWeekDay.NextWeekDay = " + popWDToWeekDay.SelectedRowIndex.ToString
		    Else
		      s = s + "dWeekDay.PreviousWeekDay = " + popWDToWeekDay.SelectedRowIndex.ToString
		    end
		  End
		  
		  If txtWDAdding.Text.ToInteger <> 0 Then
		    s = s + EndOfLine
		    s = s + "dWeekDay.AddDays = " + txtWDAdding.Text.ToInteger.ToString
		  End
		  
		  If txtCycleDuration.Text.ToInteger > 1 Then
		    s = s + EndOfLine
		    s = s + "dWeekDay.CycleFirstYear = " + txtFirstYearOfTheCycle.Text.ToInteger.ToString + "  // The event is not every year" + EndOfLine
		    s = s + "dWeekDay.CycleYearDuration = " + txtCycleDuration.Text.ToInteger.ToString
		  End
		  
		  s = s  + EndOfLine + EndOfLine + "// You can also set the DayOff, StartOfValidity and EndOfValidity properties"
		  
		  Var d As DateTime = DateTime.Now
		  
		  s = s + EndOfLine + EndOfLine
		  s = s + " // Determine the date for the year " + d.Year.ToString + EndOfLine
		  s = s + "Var Value" + d.Year.ToString +" as Datetime = dWeekDay.DateValue(" + d.Year.ToString + ")"
		  
		  s = s + EndOfLine + EndOfLine
		  s = s + "// Adding the definition to a RegionDatesWorked"
		  s = s + EndOfLine + "// Can use the RegionDatesWorked for calculate bussiness days or a list of holydays for a region"
		  s = s + EndOfLine + "// If you want to process data for multiple regions, use the MultiRegionDatesWorked class."
		  s = s + EndOfLine + "// If the event is not a holiday, set the DayOff property to False"
		  s = s + EndOfLine
		  s = s + EndOfLine + "Var DatesEngine as new RegionDatesWorked"
		  s = s + EndOfLine + "DatesEngine.AnnualEvents.add dWeekDay"
		  
		  
		  txtCode.Text = s  
		  
		  ListDatesCode.RemoveAllRows
		  
		  Var dWeekDay As New AnnualEventWeekDay(txtCaption.Text, popWDMonth.SelectedRowIndex + 1, popWDWeekDay.SelectedRowIndex + 1, popWDRank.SelectedRowIndex + 1)
		  
		  If popWDToWeekDay.SelectedRowIndex > 0 Then
		    
		    if popWDShiftMode.SelectedRowIndex = 1 then
		      dWeekDay.NextWeekDay =  popWDToWeekDay.SelectedRowIndex
		    Else
		      dWeekDay.PreviousWeekDay = popWDToWeekDay.SelectedRowIndex
		    end
		    
		  elseif txtWDAdding.Text.ToInteger <> 0 Then
		    
		    dWeekDay.AddDays =  txtWDAdding.Text.ToInteger
		    
		  End
		  
		  If txtCycleDuration.Text.ToInteger > 1 Then
		    dWeekDay.CycleFirstYear =  txtFirstYearOfTheCycle.Text.ToInteger
		    dWeekDay.CycleYearDuration = txtCycleDuration.Text.ToInteger
		  End
		  
		  Var dt As DateTime
		  
		  app.MouseCursor = System.Cursors.Wait
		  
		  If RadioGroup1.SelectedIndex = 0 Then
		    
		    For i As Integer = 1995 To 2035
		      dt = dWeekDay.DateValue(i) 
		      If dt = Nil Then Continue
		      ListDatesCode.AddRow dt.ToString("EEEE", Locale.Current) + " " + dt.ToString(DateTime.FormatStyles.Long, DateTime.FormatStyles.None)
		      
		      If dWeekDay.TestDate(dt) Then
		        ListDatesCode.CellTextAt(ListDatesCode.LastAddedRowIndex, 1) = "OK !"
		      Else
		        ListDatesCode.CellTextAt(ListDatesCode.LastAddedRowIndex, 1) = "ERROR !"
		      End
		      
		      ListDatesCode.RowTagAt(ListDatesCode.LastAddedRowIndex) = dt
		      
		    Next
		    
		    
		  Else
		    
		    For row As Integer = 0 To ListDates5Years.LastRowIndex
		      
		      
		      If dWeekDay.TestDate(ListDates5Years.RowTagAt(row).DateTimeValue) Then
		        ListDates5Years.CellTextAt(row, 1) = "Match"
		      Else
		        ListDates5Years.CellTextAt(row, 1) = "No Match"
		      End
		      
		    Next 
		    
		  End
		  
		  app.MouseCursor = Nil
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		Freeze As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		Regions() As RegionDatesWorked
	#tag EndProperty


#tag EndWindowCode

#tag Events SegmentedButton1
	#tag Event
		Sub Pressed(segmentIndex As Integer)
		  
		  PagePanel1.SelectedPanelIndex = segmentIndex
		  
		  Select Case segmentIndex
		    
		  Case 0
		    
		    FixedDayCode
		    
		  Case 1
		    
		    FromEasterCode
		    
		  Case 2
		    
		    WeekDayCode
		    
		  end
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events txtCaption
	#tag Event
		Sub TextChanged()
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub FocusLost()
		  Select Case SegmentedButton1.SelectedSegmentIndex
		    
		  Case 0
		    
		    FixedDayCode
		    
		  end
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ListDatesCode
	#tag Event
		Function RowComparison(row1 as Integer, row2 as Integer, column as Integer, ByRef result as Integer) As Boolean
		  Select Case column
		    
		    
		  Case 0
		    
		    
		    If Me.RowTagAt(row1).DateTimeValue > Me.RowTagAt(row2).DateTimeValue Then
		      result = 1
		    ElseIf Me.RowTagAt(row1).DateTimeValue < Me.RowTagAt(row2).DateTimeValue Then
		      result = -1
		    Else
		      result = 0
		    End
		    
		    Return True
		    
		  Case 1 
		    
		    Return False // Let Xojo do it
		    
		  end
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events RadioGroup1
	#tag Event
		Sub SelectionChanged(button As DesktopRadioButton)
		  #Pragma Unused button
		  
		  if me.SelectedIndex = 0 then
		    btnCopyDates.Enabled = True
		    ListDatesCode.Visible = True
		    ListDates5Years.Visible = False
		  Else
		    btnCopyDates.Enabled = False
		    ListDatesCode.Visible = False 
		    ListDates5Years.Visible = True
		  end
		  
		  select case SegmentedButton1.SelectedSegmentIndex
		    
		  case 0
		    
		    FixedDayCode
		    
		  Case 1
		    
		    FromEasterCode
		    
		  Case 2
		    
		    WeekDayCode
		    
		  end
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ListDates5Years
	#tag Event
		Function RowComparison(row1 as Integer, row2 as Integer, column as Integer, ByRef result as Integer) As Boolean
		  
		  Select case column
		    
		    
		  case 0
		    
		    
		    If Me.RowTagAt(row1).DateTimeValue > Me.RowTagAt(row2).DateTimeValue Then
		      result = 1
		    ElseIf Me.RowTagAt(row1).DateTimeValue < Me.RowTagAt(row2).DateTimeValue Then
		      result = -1
		    Else
		      result = 0
		    end
		    
		    Return True
		    
		  Case 1 
		    
		    Return False // Let Xojo do it
		    
		  end
		End Function
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
		  
		  Select Case SegmentedButton1.SelectedSegmentIndex
		    
		  Case 0
		    
		    FixedDayCode
		    
		  Case 1
		    
		    FromEasterCode
		    
		  Case 2
		    
		    WeekDayCode
		    
		  end
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
		    
		    FixedDayCode
		    
		  Case 1
		    
		    FromEasterCode
		    
		  Case 2
		    
		    WeekDayCode
		    
		  end
		End Sub
	#tag EndEvent
	#tag Event
		Sub TextChanged()
		  
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
		  
		  
		  Select Case SegmentedButton1.SelectedSegmentIndex
		    
		  Case 0
		    
		    FixedDayCode
		    
		  Case 1
		    
		    FromEasterCode
		    
		  Case 2
		    
		    WeekDayCode
		    
		  end
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
		  
		  Select Case SegmentedButton1.SelectedSegmentIndex
		    
		  Case 0
		    
		    FixedDayCode
		    
		  Case 1
		    
		    FromEasterCode
		    
		  Case 2
		    
		    WeekDayCode
		    
		  end
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnAddItem
	#tag Event
		Sub Pressed()
		  Select case rgItem.SelectedIndex
		    
		  case 1
		    
		    var w as wNewOrEditPeriod
		    Var p As ClosurePeriod = w.Charge
		    
		    if p = nil then exit sub
		    
		    Self.Regions(lstRegions.SelectedRowIndex).ClosurePeriods.Add p
		    
		    PopulateLstRegionItems
		    
		  case 2
		    
		    Var w As New wNewOrEditAnnualEvent
		    Var d As AnnualEvent = w.Charge
		    
		    If d = Nil Then Exit Sub
		    
		    Self.Regions(lstRegions.SelectedRowIndex).AddAnnualEvents(d, True) // Add without duplicate (caption or parameters)
		    
		    PopulateLstRegionItems
		    
		  end
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnLoadingRegionsFromDB
	#tag Event
		Sub Pressed()
		  Self.LoadingRegions
		  
		  Select Case Self.Regions.Count
		    
		  Case 0
		    
		    MessageBox "No region loaded"
		    
		  Case 1
		    
		    MessageBox "One region (and the events) loaded"
		    
		  Case Else
		    
		    MessageBox Self.Regions.Count.ToString + " regions (and the events) loaded"
		    
		  end
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events lstRegions
	#tag Event
		Sub Opening()
		  me.ColumnTypeAt(1) = DesktopListBox.CellTypes.CheckBox
		End Sub
	#tag EndEvent
	#tag Event
		Sub SelectionChanged()
		  If Me.SelectedRowIndex = DesktopListBox.NoSelection Then
		    lblRegionIdentifier.Text = "No region selected"
		  else
		    lblRegionIdentifier.Text = me.CellTextAt(me.SelectedRowIndex,0)
		  end
		  
		  rgItem.Visible = not (Me.SelectedRowIndex = DesktopListBox.NoSelection)
		  lstRegionItems.Visible = not (Me.SelectedRowIndex = DesktopListBox.NoSelection)
		  btnAddItem.Visible = not (Me.SelectedRowIndex = DesktopListBox.NoSelection)
		  btnEditItem.Visible = not (Me.SelectedRowIndex = DesktopListBox.NoSelection)
		  
		  If  Me.SelectedRowIndex <> DesktopListBox.NoSelection then PopulateLstRegionItems
		End Sub
	#tag EndEvent
	#tag Event
		Function KeyDown(key As String) As Boolean
		  If key.Asc = 127 And Me.SelectedRowIndex <> DesktopListBox.NoSelection Then
		    Var index as integer = me.SelectedRowIndex
		    self.Regions.RemoveAt(index)
		    me.RemoveRowAt(index)
		    
		    If Me.RowCount = 0 Then
		      
		      // Nothing
		      
		    ElseIf index >= Me.RowCount  then
		      
		      me.SelectedRowIndex = me.LastRowIndex
		      
		    else
		      
		      me.SelectedRowIndex = index
		      
		    end
		    
		    Return True
		    
		  End If
		  
		  
		  
		  
		  
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events btnCopyDates
	#tag Event
		Sub Pressed()
		  if ListDatesCode.RowCount = 0 then exit sub
		  
		  Var strDates as string
		  for row as Integer = 0 to ListDatesCode.LastRowIndex
		    if row > 0 then strDates = strDates + EndOfLine
		    strDates = strDates + ListDatesCode.RowTagAt(row).DateTimeValue.ToString(DateTime.FormatStyles.Long, DateTime.FormatStyles.None)
		  Next row
		  
		  Var cb As New Clipboard
		  cb.Text = strDates
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnNewRegion
	#tag Event
		Sub Pressed()
		  If txtNewRegion.Text.Trim = "" Then Exit Sub
		  Self.Regions.Add New RegionDatesWorked(txtNewRegion.Text.Trim.Lowercase)
		  PopulateLstRegions
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnEditItem
	#tag Event
		Sub Pressed()
		  
		  select case rgItem.SelectedIndex
		    
		  Case 1
		    
		    Var w As New wNewOrEditPeriod
		    Var p As ClosurePeriod = w.Charge
		    
		    If p = Nil Then Exit Sub
		    
		    Var row As Integer = lstRegionItems.SelectedRowIndex
		    
		    Self.Regions(lstRegions.SelectedRowIndex).ClosurePeriods(row) = Nil // Clean the old object
		    Self.Regions(lstRegions.SelectedRowIndex).ClosurePeriods(row)  = p // Affects the new object
		    
		    lstRegionItems.CellTextAt(row, 0)=  p.Caption
		    
		    lstRegionItems.CellTextAt(row, 1) = p.FirstDay.ToString(DateTime.FormatStyles.Short, DateTime.FormatStyles.None)
		    lstRegionItems.CellTextat(row, 2) = p.LastDay.ToString(DateTime.FormatStyles.Short, DateTime.FormatStyles.None)
		    
		    lstRegionItems.CellTagAt(row, 1) = p.FirstDay
		    lstRegionItems.CellTagAt(row, 2) = p.LastDay   
		    
		  case 2
		    
		    Var w As New wNewOrEditAnnualEvent
		    Var d As AnnualEvent = w.Charge(Self.Regions(lstRegions.SelectedRowIndex).AnnualEvents(lstRegionItems.SelectedRowIndex))
		    If d = Nil Then Exit Sub
		    
		    Self.Regions(lstRegions.SelectedRowIndex).AnnualEvents(lstRegionItems.SelectedRowIndex) = Nil // Clean the old object
		    Self.Regions(lstRegions.SelectedRowIndex).AnnualEvents(lstRegionItems.SelectedRowIndex) = d // Affects the new object
		    
		    lstRegionItems.CellTextAt(lstRegionItems.SelectedRowIndex, 0) = d.Caption
		    
		    Select Case d
		      
		    Case IsA AnnualEventFix 
		      
		      lstRegionItems.CellTextAt(lstRegionItems.SelectedRowIndex, 1) = "Fix"
		      
		    Case IsA AnnualEventEaster
		      
		      lstRegionItems.CellTextAt(lstRegionItems.SelectedRowIndex, 1) = "Easter"
		      
		    Case IsA AnnualEventOrthodoxEaster
		      
		      lstRegionItems.CellTextAt(lstRegionItems.SelectedRowIndex, 1) = "O.Easter"
		      
		    Case IsA AnnualEventWeekDay
		      
		      lstRegionItems.CellTextAt(lstRegionItems.SelectedRowIndex, 1) = "WeekDay"
		      
		    End
		    
		  end
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnNewRegion1
	#tag Event
		Sub Pressed()
		  Self.Regions.RemoveAll
		  
		  Self.Regions.Add New RegionDatesWorked("usa-federal", "Federal")
		  Self.Regions(Regions.LastIndex).AddAnnualEvents RegionDatesWorked.AnnualEventsUSA(RegionDatesWorked.USA.Federal)
		  
		  Self.Regions.Add New RegionDatesWorked("usa-alabama", "Alabama")
		  Self.Regions(Regions.LastIndex).AddAnnualEvents RegionDatesWorked.AnnualEventsUSA(RegionDatesWorked.USA.Alabama)
		  
		  Self.Regions.Add New RegionDatesWorked("usa-alaska", "Alaska")
		  Self.Regions(Regions.LastIndex).AddAnnualEvents RegionDatesWorked.AnnualEventsUSA(RegionDatesWorked.USA.Alaska)
		  
		  Self.Regions.Add New RegionDatesWorked("usa-arkansas", "Arkansas")
		  Self.Regions(Regions.LastIndex).AddAnnualEvents RegionDatesWorked.AnnualEventsUSA(RegionDatesWorked.USA.Arkansas)
		  
		  Self.Regions.Add New RegionDatesWorked("usa-arizona", "Arizona")
		  Self.Regions(Regions.LastIndex).AddAnnualEvents RegionDatesWorked.AnnualEventsUSA(RegionDatesWorked.USA.Arizona)
		  
		  Self.Regions.Add New RegionDatesWorked("usa-california", "California")
		  Self.Regions(Regions.LastIndex).AddAnnualEvents RegionDatesWorked.AnnualEventsUSA(RegionDatesWorked.USA.California)
		  
		  Self.Regions.Add New RegionDatesWorked("usa-colorado", "Colorado")
		  Self.Regions(Regions.LastIndex).AddAnnualEvents RegionDatesWorked.AnnualEventsUSA(RegionDatesWorked.USA.Colorado)
		  
		  Self.Regions.Add New RegionDatesWorked("usa-connecticut", "Connecticut")
		  Self.Regions(Regions.LastIndex).AddAnnualEvents RegionDatesWorked.AnnualEventsUSA(RegionDatesWorked.USA.Connecticut)
		  
		  Self.Regions.Add New RegionDatesWorked("usa-delaware", "Delaware")
		  Self.Regions(Regions.LastIndex).AddAnnualEvents RegionDatesWorked.AnnualEventsUSA(RegionDatesWorked.USA.Delaware)
		  
		  
		  Self.Regions.Add New RegionDatesWorked("usa-dc", "District of Columbia")
		  Self.Regions(Regions.LastIndex).AddAnnualEvents RegionDatesWorked.AnnualEventsUSA(RegionDatesWorked.USA.DistrictOfColumbia)
		  
		  
		  // No more states implanted, because it's a bit time-consuming to find the official texts for each state. Not to mention the counties...
		  
		  PopulateLstRegions
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnRestoreRegionToFactory
	#tag Event
		Sub Pressed()
		  
		  Var m As New MessageDialog
		  m.Title = Self.Title
		  m.Message = "Are you sure ?"
		  m.IconType = MessageDialog.IconTypes.Question
		  m.ActionButton.Caption = "Yes"
		  m.CancelButton.Caption = "No"
		  m.CancelButton.Visible = True
		  m.CancelButton.Default = True
		  m.CancelButton.Cancel = True
		  Var b As MessageDialogButton  = m.ShowModal
		  if b = m.CancelButton then exit sub
		  
		  
		  // Remove all rows
		  
		  Try
		    app.db.ExecuteSQL "DELETE FROM tdays;"
		  Catch err As DatabaseException
		    MessageBox "(" + Err.ErrorNumber.ToString + ") " + err.Message
		    Break
		    Exit Sub
		  End
		  
		  Try
		    app.db.ExecuteSQL "DELETE FROM tregions_weekdays"
		  Catch err As DatabaseException
		    MessageBox "(" + Err.ErrorNumber.ToString + ") " + err.Message
		    Break
		    Exit Sub
		  End
		  
		  // Copying from the table tdays_factorysettings
		  
		  Var strSQL As String
		  
		  strSQL = strSQL + " Insert into tdays(" + EndOfLine
		  strSQL = strSQL + " start," + EndOfLine
		  strSQL = strSQL + " ""end""," + EndOfLine
		  strSQL = strSQL + " caption," + EndOfLine
		  strSQL = strSQL + " region," + EndOfLine
		  strSQL = strSQL + " definitiontype," + EndOfLine
		  strSQL = strSQL + " dayoff," + EndOfLine
		  strSQL = strSQL + " cycleyearduration," + EndOfLine
		  strSQL = strSQL + " cyclefirstyear," + EndOfLine
		  strSQL = strSQL + " month," + EndOfLine
		  strSQL = strSQL + " day," + EndOfLine
		  strSQL = strSQL + " weekday," + EndOfLine
		  strSQL = strSQL + " rank," + EndOfLine
		  strSQL = strSQL + " adddays," + EndOfLine
		  strSQL = strSQL + " nextweekday," + EndOfLine
		  strSQL = strSQL + " previousweekday," + EndOfLine
		  strSQL = strSQL + " saturdaytofriday," + EndOfLine
		  strSQL = strSQL + " sundaytomonday," + EndOfLine
		  strSQL = strSQL + " saturdaytomonday," + EndOfLine
		  strSQL = strSQL + " fingerprint)" + EndOfLine
		  
		  strSQL = strSQL + " Select" + EndOfLine
		  strSQL = strSQL + " start," + EndOfLine
		  strSQL = strSQL + " ""end""," + EndOfLine
		  strSQL = strSQL + " caption," + EndOfLine
		  strSQL = strSQL + " region," + EndOfLine
		  strSQL = strSQL + " definitiontype," + EndOfLine
		  strSQL = strSQL + " dayoff," + EndOfLine
		  strSQL = strSQL + " cycleyearduration," + EndOfLine
		  strSQL = strSQL + " cyclefirstyear," + EndOfLine
		  strSQL = strSQL + " month," + EndOfLine
		  strSQL = strSQL + " day," + EndOfLine
		  strSQL = strSQL + " weekday," + EndOfLine
		  strSQL = strSQL + " rank," + EndOfLine
		  strSQL = strSQL + " adddays," + EndOfLine
		  strSQL = strSQL + " nextweekday," + EndOfLine
		  strSQL = strSQL + " previousweekday," + EndOfLine
		  strSQL = strSQL + " saturdaytofriday," + EndOfLine
		  strSQL = strSQL + " sundaytomonday," + EndOfLine
		  strSQL = strSQL + " saturdaytomonday," + EndOfLine
		  strSQL = strSQL + " fingerprint" + EndOfLine
		  strSQL = strSQL + " From" + EndOfLine
		  strSQL = strSQL + " tdays_factorysettings" + EndOfLine
		  strSQL = strSQL + " Order by id_date;" + EndOfLine
		  
		  Try
		    app.db.ExecuteSQL strSQL
		  Catch err As DatabaseException
		    MessageBox "(" + Err.ErrorNumber.ToString + ") " + err.Message
		    Break
		    Exit Sub
		  End
		  
		  
		  strSQL = ""
		  strSQL = strSQL + " INSERT INTO ""tregions_weekdays_factorysettings""" + EndOfLine
		  strSQL = strSQL + " SELECT * FROM ""tregions_weekdays"";" + EndOfLine
		  
		  Try
		    app.db.ExecuteSQL strSQL
		  Catch err As DatabaseException
		    MessageBox "(" + Err.ErrorNumber.ToString + ") " + err.Message
		    Break
		    Exit Sub
		  End
		  
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
	#tag Event
		Sub SelectionChanged(item As DesktopMenuItem)
		  #Pragma Unused item
		  FixedDayCode
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
	#tag Event
		Sub SelectionChanged(item As DesktopMenuItem)
		  #Pragma Unused item
		  FixedDayCode
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
	#tag Event
		Sub SelectionChanged(item As DesktopMenuItem)
		  #Pragma Unused item
		  FixedDayCode
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
		  
		  FixedDayCode
		End Sub
	#tag EndEvent
	#tag Event
		Sub TextChanged()
		  
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
		  FixedDayCode
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
		  FixedDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events cbFDMondayIfSunday
	#tag Event
		Sub ValueChanged()
		  If Me.value Then
		    cbFDFridayIfSaturday.Value = False
		  End
		  
		  FixedDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events cbFDFridayIfSaturday
	#tag Event
		Sub ValueChanged()
		  FixedDayCode
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
		  
		  FixedDayCode
		End Sub
	#tag EndEvent
	#tag Event
		Sub TextChanged()
		  
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
		  FixedDayCode
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
		  FixedDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption 
		  txtEDDiff.Text = "0"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster1
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "-2"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster2
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "50"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster3
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "40"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster4
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "60"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster5
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "63"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster6
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "56"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster7
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = AnnualEventEaster.MardiGras.ToString // or txtEDDiff.Text = "-47"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster8
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "-7"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster9
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = AnnualEventEaster.AshWednesday.ToString // // or txtEDDiff.Text = "-46"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster16
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "-21" // Or txtEDDiff.Text = AnnualEventEaster.MotheringSunday.ToString
		  FromEasterCode
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster18
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "-3"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster19
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "-1"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster20
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "7"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster21
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "49"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster17
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = "-14"
		  FromEasterCode
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
	#tag Event
		Sub SelectionChanged(item As DesktopMenuItem)
		  #Pragma Unused item
		  WeekDayCode
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
	#tag Event
		Sub SelectionChanged(item As DesktopMenuItem)
		  #Pragma Unused item
		  WeekDayCode
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
	#tag Event
		Sub SelectionChanged(item As DesktopMenuItem)
		  #Pragma Unused item
		  WeekDayCode
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
		  WeekDayCode
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
		  
		  WeekDayCode
		End Sub
	#tag EndEvent
	#tag Event
		Sub TextChanged()
		  
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
		  WeekDayCode
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
		  WeekDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD1
	#tag Event
		Sub Pressed()
		  // https://en.wikipedia.org/wiki/Election_Day_(United_States)
		  
		  Self.Freeze = True
		  txtCaption.Text = Me.Caption
		  
		  popWDMonth.SelectedRowIndex = 10 // November
		  popWDWeekDay.SelectedRowIndex = 1 // Monday
		  popWDRank.SelectedRowIndex = 0 // The first
		  // But the tuesday after the 1st monday
		  popWDShiftMode.SelectedRowIndex = 1 // Next
		  popWDToWeekDay.SelectedRowIndex = 3 // To the next tuesday
		  
		  txtCycleDuration.Text = "4"
		  txtFirstYearOfTheCycle.Text = "1848"
		  
		  Self.Freeze = False
		  
		  self.WeekDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD2
	#tag Event
		Sub Pressed()
		  // https://en.wikipedia.org/wiki/Public_holidays_in_the_United_States
		  
		  self.Freeze = True
		  txtCaption.Text = me.Caption
		  
		  popWDMonth.SelectedRowIndex = 8 // September
		  popWDWeekDay.SelectedRowIndex = 1 // Monday
		  popWDRank.SelectedRowIndex = 0 // The first
		  
		  popWDToWeekDay.SelectedRowIndex = 0
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  Self.Freeze = False
		  
		  self.WeekDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD3
	#tag Event
		Sub Pressed()
		  // https://en.wikipedia.org/wiki/Public_holidays_in_the_United_States
		  
		  self.Freeze = True
		  txtCaption.Text = me.Caption
		  
		  popWDMonth.SelectedRowIndex = 10 // November
		  popWDWeekDay.SelectedRowIndex = 4 // Thursday
		  popWDRank.SelectedRowIndex = 3 // The first
		  
		  popWDToWeekDay.SelectedRowIndex = 0
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  Self.Freeze = False
		  
		  self.WeekDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD4
	#tag Event
		Sub Pressed()
		  // https://es.wikipedia.org/wiki/Anexo:D%C3%ADas_festivos_en_M%C3%A9xico
		  
		  Self.Freeze = True
		  txtCaption.Text = "Proclamación de la Constitución de 1917"
		  
		  popWDMonth.SelectedRowIndex = 1 // February
		  popWDWeekDay.SelectedRowIndex = 1 // Monday
		  popWDRank.SelectedRowIndex = 0 // The first
		  
		  popWDToWeekDay.SelectedRowIndex = 0
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  Self.Freeze = False
		  
		  self.WeekDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD5
	#tag Event
		Sub Pressed()
		  // https://es.wikipedia.org/wiki/Anexo:D%C3%ADas_festivos_en_M%C3%A9xico
		  
		  self.Freeze = True
		  txtCaption.Text = "Aniversario del nacimiento de Benito Juárez"
		  
		  popWDMonth.SelectedRowIndex = 2 // March
		  popWDWeekDay.SelectedRowIndex = 1 // Monday
		  popWDRank.SelectedRowIndex = 2 // The third
		  
		  popWDToWeekDay.SelectedRowIndex = 0
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  Self.Freeze = False
		  
		  self.WeekDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD6
	#tag Event
		Sub Pressed()
		  // https://es.wikipedia.org/wiki/Anexo:D%C3%ADas_festivos_en_M%C3%A9xico
		  
		  self.Freeze = True
		  txtCaption.Text = me.Caption
		  
		  popWDMonth.SelectedRowIndex = 10 // November
		  popWDWeekDay.SelectedRowIndex = 1 // Monday
		  popWDRank.SelectedRowIndex = 2 // The third
		  
		  popWDToWeekDay.SelectedRowIndex = 0
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  Self.Freeze = False
		  
		  self.WeekDayCode
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
	#tag Event
		Sub SelectionChanged(item As DesktopMenuItem)
		  #Pragma Unused item
		  WeekDayCode
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
	#tag Event
		Sub SelectionChanged(item As DesktopMenuItem)
		  #Pragma Unused item
		  FixedDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD7
	#tag Event
		Sub Pressed()
		  Self.Freeze = True
		  txtCaption.Text = me.Caption
		  
		  popWDMonth.SelectedRowIndex = 8 // September
		  popWDWeekDay.SelectedRowIndex = 1 // Monday
		  popWDRank.SelectedRowIndex = 0 // The first
		  
		  popWDToWeekDay.SelectedRowIndex = 0
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  Self.Freeze = False
		  
		  self.WeekDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD8
	#tag Event
		Sub Pressed()
		  Self.Freeze = True
		  txtCaption.Text = Me.Caption
		  
		  popWDMonth.SelectedRowIndex = 9 // October
		  popWDWeekDay.SelectedRowIndex = 1 // Monday
		  popWDRank.SelectedRowIndex = 1 // The first
		  
		  popWDToWeekDay.SelectedRowIndex = 0
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  Self.Freeze = False
		  
		  self.WeekDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD9
	#tag Event
		Sub Pressed()
		  // https://fr.wikipedia.org/wiki/F%C3%AAte_de_Wallonie
		  
		  Self.Freeze = True
		  txtCaption.Text = me.Caption
		  
		  popWDMonth.SelectedRowIndex = 8 // September
		  popWDWeekDay.SelectedRowIndex = 0 // Sunday
		  popWDRank.SelectedRowIndex = 3 // The third
		  
		  popWDToWeekDay.SelectedRowIndex = 0
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  Self.Freeze = False
		  
		  self.WeekDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD10
	#tag Event
		Sub Pressed()
		  Self.Freeze = True
		  txtCaption.Text = me.Caption
		  
		  popFDMonth.SelectedRowIndex = 4 // Mai
		  popFDDay.SelectedRowIndex = 24 // 25
		  
		  popFDWeekDay.SelectedRowIndex = 2
		  popFDShiftMode.SelectedRowIndex = 0
		  chkDFAlways.Value = True
		  
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  
		  
		  Self.Freeze = False
		  
		  self.FixedDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD11
	#tag Event
		Sub Pressed()
		  Self.Freeze = True
		  txtCaption.Text = me.Caption
		  
		  popFDMonth.SelectedRowIndex = 6 // Jully
		  popFDDay.SelectedRowIndex = 3 // 4th
		  
		  popFDWeekDay.SelectedRowIndex = 0 // Off
		  popFDShiftMode.SelectedRowIndex = 0
		  chkDFAlways.Value = False
		  
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  
		  
		  Self.Freeze = False
		  
		  self.FixedDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD12
	#tag Event
		Sub Pressed()
		  Self.Freeze = True
		  txtCaption.Text = me.Caption
		  
		  popFDMonth.SelectedRowIndex = 1 // February
		  popFDDay.SelectedRowIndex = 28 // 29rd
		  
		  popFDWeekDay.SelectedRowIndex = 0 // Off
		  popFDShiftMode.SelectedRowIndex = 0
		  chkDFAlways.Value = False
		  
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  Self.Freeze = False
		  
		  self.FixedDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD13
	#tag Event
		Sub Pressed()
		  Self.Freeze = True
		  txtCaption.Text = me.Caption
		  
		  popFDMonth.SelectedRowIndex = 8 // September
		  popFDDay.SelectedRowIndex = 15 // 16th
		  
		  popFDWeekDay.SelectedRowIndex = 0 // Off
		  popFDShiftMode.SelectedRowIndex = 0
		  chkDFAlways.Value = False
		  
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  Self.Freeze = False
		  
		  self.FixedDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonWD14
	#tag Event
		Sub Pressed()
		  Self.Freeze = True
		  txtCaption.Text = me.Caption
		  
		  popFDMonth.SelectedRowIndex = 6 // Jully
		  popFDDay.SelectedRowIndex = 20 //
		  
		  popFDWeekDay.SelectedRowIndex = 2
		  popFDShiftMode.SelectedRowIndex = 0
		  chkDFAlways.Value = False
		  
		  txtCycleDuration.Text = "1"
		  txtFirstYearOfTheCycle.Text = "1900"
		  
		  Self.Freeze = False
		  
		  self.FixedDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonEaster15
	#tag Event
		Sub Pressed()
		  txtCaption.Text = me.Caption
		  txtEDDiff.Text = AnnualEventEaster.CleanMonday.ToString // or txtEDDiff.Text = "-48"
		  FromEasterCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events chkDFAlways
	#tag Event
		Sub ValueChanged()
		  FixedDayCode
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events lstRegionItems
	#tag Event
		Sub SelectionChanged()
		  btnEditItem.Enabled = not (Me.SelectedRowIndex = DesktopListBox.NoSelection)
		End Sub
	#tag EndEvent
	#tag Event
		Sub DoublePressed()
		  if rgItem.SelectedIndex = 0 then exit sub
		  if me.SelectedRowIndex = DesktopListBox.NoSelection then exit sub
		  
		  
		  select case rgItem.SelectedIndex
		    
		  case 1 // Closure Periods
		    
		    Var w As New wNewOrEditPeriod
		    Var p As ClosurePeriod = w.Charge
		    
		    If p = Nil Then Exit Sub
		    
		    Self.Regions(lstRegions.SelectedRowIndex).ClosurePeriods(Me.SelectedRowIndex) = Nil // Clean the old object
		    Self.Regions(lstRegions.SelectedRowIndex).ClosurePeriods(Me.SelectedRowIndex)  = p // Affects the new object
		    
		  case 2 // Annual events
		    
		    Var w As New wNewOrEditAnnualEvent
		    Var d As AnnualEvent = w.Charge(Self.Regions(lstRegions.SelectedRowIndex).AnnualEvents(Me.SelectedRowIndex))
		    If d = Nil Then Exit Sub
		    
		    Self.Regions(lstRegions.SelectedRowIndex).AnnualEvents(Me.SelectedRowIndex) = Nil // Clean the old object
		    Self.Regions(lstRegions.SelectedRowIndex).AnnualEvents(Me.SelectedRowIndex) = d // Affects the new object
		    
		    Me.CellTextAt(Me.SelectedRowIndex, 0) = d.Caption
		    
		    Select Case d
		      
		    Case IsA AnnualEventFix 
		      
		      Me.CellTextAt(Me.SelectedRowIndex, 1) = "Fix"
		      
		    Case IsA AnnualEventEaster
		      
		      Me.CellTextAt(Me.SelectedRowIndex, 1) = "Easter"
		      
		    Case IsA AnnualEventOrthodoxEaster
		      
		      Me.CellTextAt(Me.SelectedRowIndex, 1) = "O.Easter"
		      
		    Case IsA AnnualEventWeekDay
		      
		      Me.CellTextAt(Me.SelectedRowIndex, 1) = "WeekDay"
		      
		    End
		    
		  end Select
		End Sub
	#tag EndEvent
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if rgItem.SelectedIndex = 0 then Return False
		  
		  If key.Asc = 127 And Me.SelectedRowIndex <> DesktopListBox.NoSelection Then
		    Var index As Integer = Me.SelectedRowIndex
		    
		    select case rgItem.SelectedIndex
		      
		      
		    case 1
		      self.Regions(lstRegions.SelectedRowIndex).ClosurePeriods.RemoveAt index
		    case 2
		      Self.Regions(lstRegions.SelectedRowIndex).AnnualEvents.RemoveAt index
		    end
		    
		    Me.RemoveRowAt(index)
		    
		    If Me.RowCount = 0 Then
		      
		      // Nothing
		      
		    ElseIf index >= Me.RowCount  Then
		      
		      Me.SelectedRowIndex = Me.LastRowIndex
		      
		    Else
		      
		      Me.SelectedRowIndex = index
		      
		    End
		    
		    Return True
		    
		  End If
		End Function
	#tag EndEvent
	#tag Event
		Sub CellAction(row As Integer, column As Integer)
		  #Pragma Unused column
		  
		  Select Case rgItem.SelectedIndex
		    
		  Case 0
		    
		    self.Regions(lstRegions.SelectedRowIndex).WorkingWeekDays.WorkingDay(row + 1) = lstRegionItems.CellCheckBoxValueAt(row, 1)
		    
		  Case 1
		    
		    // Nothing
		    
		    
		  Case 2
		    
		    // Nothing
		    
		  end
		End Sub
	#tag EndEvent
	#tag Event
		Sub Opening()
		  Me.ColumnTypeAt(1) = DesktopListBox.CellTypes.CheckBox
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events dtLast
	#tag Event
		Sub Opening()
		  Var d As DateTime = DateTime.Now
		  d = d.AddInterval (0,3)
		  d = d.SubtractInterval(0,0,1)
		  me.SelectedDate = new DateTime(d.Year,d.Month,d.Day)
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events dtFirst
	#tag Event
		Sub Opening()
		  me.SelectedDate = new DateTime(DateTime.Now.Year,DateTime.Now.Month,DateTime.Now.Day)
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events rgDaysFunctions
	#tag Event
		Sub SelectionChanged(button As DesktopRadioButton)
		  #Pragma Unused button
		  lblEvent.Visible =  (me.SelectedIndex = 1)
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events Button1
	#tag Event
		Sub Pressed()
		  If rgDaysFunctions.SelectedIndex = 1 And lstRegions.SelectedRowIndex = DesktopListBox.NoSelection Then
		    MessageBox "No region selected."
		    Exit Sub
		  ElseIf rgDaysFunctions.SelectedIndex = 0 And Self.Regions.Count = 0 Then
		    MessageBox " No region in the list."
		    Exit Sub
		  End
		  
		  app.MouseCursor = System.Cursors.Wait
		  
		  Var d() as DateTime
		  
		  If rgDaysFunctions.SelectedIndex = 1 Then
		    
		    
		    d = self.Regions(lstRegions.SelectedRowIndex).BusinessDays(dtFirst.SelectedDate, dtLast.SelectedDate)
		    
		    
		  Else
		    
		    // Creating a muli-region object
		    // We could also create it for the window and work on its Regions array
		    
		    Var mr As new MultiRegionDatesWorked
		    mr.Regions = self.Regions
		    
		    // Only days that are worked for ALL regions at the same time.
		    d = mr.BusinessDays(dtFirst.SelectedDate, dtLast.SelectedDate)
		    
		  end
		  
		  
		  lstResults.RemoveAllRows
		  For i As Integer = 0 To d.LastIndex
		    
		    lstResults.AddRow d(i).ToString(DateTime.FormatStyles.Full, DateTime.FormatStyles.None)
		    lstResults.RowTagAt(lstResults.LastAddedRowIndex) = d(i)
		    
		  next i
		  
		  app.MouseCursor = nil
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events Button2
	#tag Event
		Sub Pressed()
		  If rgDaysFunctions.SelectedIndex = 1 And lstRegions.SelectedRowIndex = DesktopListBox.NoSelection Then
		    MessageBox "No region selected."
		    Exit Sub
		  ElseIf rgDaysFunctions.SelectedIndex = 0 And Self.Regions.Count = 0 Then
		    MessageBox " No region in the list."
		    Exit Sub
		  End
		  
		  app.MouseCursor = System.Cursors.Wait
		  
		  Var d() as DateTime
		  
		  If rgDaysFunctions.SelectedIndex = 1 Then
		    
		    
		    d = self.Regions(lstRegions.SelectedRowIndex).NonWorkingDays(dtFirst.SelectedDate, dtLast.SelectedDate)
		    
		    
		  Else
		    
		    // Creating a muli-region object
		    // We could also create it for the window and work on its Regions array
		    
		    Var mr As new MultiRegionDatesWorked
		    mr.Regions = self.Regions
		    
		    // Only days that are worked for ALL regions at the same time.
		    d = mr.NonWorkingDays(dtFirst.SelectedDate, dtLast.SelectedDate)
		    
		  end
		  
		  
		  lstResults.RemoveAllRows
		  
		  For i As Integer = 0 To d.LastIndex
		    
		    lstResults.AddRow d(i).ToString(DateTime.FormatStyles.Full, DateTime.FormatStyles.None)
		    lstResults.RowTagAt(lstResults.LastAddedRowIndex) = d(i)
		    
		  next i
		  
		  app.MouseCursor = Nil
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events dtThisDate
	#tag Event
		Sub Opening()
		  Me.SelectedDate = DateTime.Now
		  
		  
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub DateChanged(value as DateTime)
		  #Pragma Unused value
		  CalculateThisDate
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events txtFollowingBD
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
		  
		  CalculateThisDate
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events UpDownArrows5
	#tag Event
		Sub DownPressed()
		  Var v As Integer
		  
		  
		  #Pragma BreakOnExceptions False
		  Try
		    v= txtFollowingBD.Text.ToInteger
		  Catch
		    v = 0
		  End
		  #Pragma BreakOnExceptions True
		  
		  v = v - 1
		  
		  txtFollowingBD.Text = v.ToString
		  
		  CalculateThisDate
		End Sub
	#tag EndEvent
	#tag Event
		Sub UpPressed()
		  Var v As Integer
		  
		  
		  #Pragma BreakOnExceptions False
		  Try
		    v= txtFollowingBD.Text.ToInteger
		  Catch
		    v = 0
		  End
		  #Pragma BreakOnExceptions True
		  
		  v = v + 1
		  
		  txtFollowingBD.Text = v.ToString
		  
		  CalculateThisDate
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnConnectDB
	#tag Event
		Sub Pressed()
		  Var m As New OpenFileDialog
		  m.ActionButtonCaption = "Connect"
		  m.Title = "Connect to days.db"
		  m.SuggestedFileName = "days.db"
		  m.InitialFolder = new FolderItem
		  m.Filter ="DbType"
		  Var f as FolderItem = m.ShowModal 
		  
		  If f = Nil Then Exit Sub
		  if not f.Exists then exit sub
		  
		  Try
		    app.db = New SQLiteDatabase(f)
		    app.db.Connect
		  Catch err1 As DatabaseException
		    MessageBox "(" + Err1.ErrorNumber.ToString + ") " + err1.Message
		    Break
		    Exit Sub
		  Catch err2 As IOException
		    MessageBox "(" + Err2.ErrorNumber.ToString + ") " + err2.Message
		    Break
		    Exit Sub
		  End
		  
		  btnLoadingRegionsFromDB.Enabled = True
		  btnRestoreRegionToFactory.Enabled = True
		  lblDBstatuts.Text = "days.db is connected"
		  lblDBstatuts.TextColor = Color.RGB(0,128,0)
		  me.Visible = False
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events rgItem
	#tag Event
		Sub SelectionChanged(button As DesktopRadioButton)
		  #Pragma Unused button
		  
		  Select Case Me.SelectedIndex
		    
		  Case 0 // Working day
		    
		    btnAddItem.Visible = False
		    btnEditItem.Visible = False
		    
		    lstRegionItems.ColumnWidths = ",25%,0"
		    lstRegionItems.HeaderAt(0) = "Week day"
		    lstRegionItems.HeaderAt(1) = "Working ?"
		    lstRegionItems.HeaderAt(2) = " "
		    
		    lstRegionItems.ColumnTypeAt(1) = DesktopListBox.CellTypes.CheckBox
		    
		  Case 1
		    btnAddItem.Visible = True
		    btnEditItem.Visible = True
		    btnAddItem.Caption = "Adding a Closure period"
		    btnEditItem.Caption = "Edit a Closure period"
		    lstRegionItems.ColumnWidths = ",25%,25%"
		    lstRegionItems.HeaderAt(0) = "Period caption"
		    lstRegionItems.HeaderAt(1) = "First day"
		    lstRegionItems.HeaderAt(2) = "Last day"
		    
		    lstRegionItems.ColumnTypeAt(1) = DesktopListBox.CellTypes.Normal
		    
		  Case 2
		    btnAddItem.Visible = True
		    btnEditItem.Visible = True
		    btnAddItem.Caption = "Adding a annual event"
		    btnEditItem.Caption = "Edit a annual event"
		    
		    lstRegionItems.ColumnWidths = ",25%,0"
		    lstRegionItems.HeaderAt(0) = "Annual event"
		    lstRegionItems.HeaderAt(1) = "Type"
		    lstRegionItems.HeaderAt(2) = " "
		    
		    lstRegionItems.ColumnTypeAt(1) = DesktopListBox.CellTypes.Normal
		    
		  End
		  
		  PopulateLstRegionItems
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
	#tag ViewProperty
		Name="Freeze"
		Visible=false
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
#tag EndViewBehavior
