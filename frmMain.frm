VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Woodpecker"
   ClientHeight    =   7875
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9630
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraNoFiles 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   5640
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Label lblNoFiles 
         BackStyle       =   0  'Transparent
         Caption         =   "No Valid Icons Could Be Found In Any Files In The Selected Folder"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2535
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7560
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picIconHolder 
      Height          =   5535
      Left            =   4440
      ScaleHeight     =   5475
      ScaleWidth      =   4275
      TabIndex        =   5
      Top             =   0
      Width           =   4335
      Begin VB.Frame fraIconHolder 
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4215
         Begin VB.VScrollBar VS 
            Height          =   1575
            Left            =   3840
            Max             =   0
            TabIndex        =   9
            Top             =   120
            Width           =   255
         End
         Begin VB.Frame fraIcons 
            BorderStyle     =   0  'None
            Height          =   4215
            Left            =   360
            TabIndex        =   7
            Top             =   360
            Width           =   3495
            Begin VB.PictureBox picIcon 
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               Height          =   495
               Index           =   0
               Left            =   240
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   8
               Top             =   240
               Width           =   495
            End
            Begin VB.Line lnFrame 
               BorderColor     =   &H0000FF00&
               Index           =   0
               Visible         =   0   'False
               X1              =   0
               X2              =   0
               Y1              =   0
               Y2              =   720
            End
            Begin VB.Line lnFrame 
               BorderColor     =   &H0000FF00&
               Index           =   1
               Visible         =   0   'False
               X1              =   0
               X2              =   0
               Y1              =   1080
               Y2              =   1800
            End
            Begin VB.Line lnFrame 
               BorderColor     =   &H0000FF00&
               Index           =   2
               Visible         =   0   'False
               X1              =   0
               X2              =   0
               Y1              =   1920
               Y2              =   2640
            End
            Begin VB.Line lnFrame 
               BorderColor     =   &H0000FF00&
               Index           =   3
               Visible         =   0   'False
               X1              =   0
               X2              =   0
               Y1              =   2880
               Y2              =   3600
            End
         End
      End
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9570
      TabIndex        =   2
      Top             =   7500
      Width           =   9630
      Begin MSComctlLib.ProgressBar progbar 
         Height          =   135
         Left            =   2040
         TabIndex        =   4
         Top             =   120
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin MSComctlLib.ListView File1 
      Height          =   3975
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename (0)"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "0"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Menu mnuTop 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "&Quit"
         Index           =   0
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTop 
      Caption         =   "&Icons"
      Index           =   1
      Begin VB.Menu mnuIcons 
         Caption         =   "&Copy to Clipboard"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuIcons 
         Caption         =   "&Save to File"
         Index           =   1
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuTop 
      Caption         =   "&Directories"
      Index           =   2
      Begin VB.Menu mnuDirectories 
         Caption         =   "Show &all Icons"
         Index           =   0
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This program was written by Michael Drotar
'   Visual Basic Version:       6.0
'   Tab Size:                   4 spaces
'   Screen Resolution:          1024 x 768
'   Processor:                  Pentium 500mHz
'   Operating System:           Windows XP
    
'All code and comments are written to fit within the above version of VB at the
'   above screen resolution with room for toolbar and other docked windows.

'Note: At certain times, I'll toggle the .Visible property of the fraIcons control.
'   This is because the program runs significantly faster when it's not visible and
'   certain operations tend to take a very long time otherwise.

'During testing, I was able to process 2039 files and load the 303 of them that had icons
'   into the file list along with their number of icons in 10 seconds.
'   I then loaded the 2135 icons in those files in 20 seconds and positioned them in
'   roughly 55 seconds.

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" _
            (lpPoint As POINTAPI) _
                                                                            As Long

Private Declare Function SetCursorPos Lib "user32" _
            (ByVal X As Long, ByVal Y As Long) _
                                                                            As Long

Private Declare Function DrawIcon Lib "user32.dll" _
            (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
                ByVal hIcon As Long) _
                                                                            As Long

Private Declare Function DestroyIcon Lib "user32.dll" _
            (ByVal hIcon As Long) _
                                                                            As Long
                                                                            
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" _
            (ByVal hInst As Long, ByVal lpszExeFileName As String, _
                ByVal nIconIndex As Long) _
                                                                            As Long

Const DEF_FOLDER = "C:\Windows\System32"    'Default folder to load (if it exists)
                                            '   I chose this because \System32 has the
                                            '   most Windows icons
                                                                            
Dim lSelected As Long                       'The index of the selected icon
Dim lCount As Long                          'The total number of icons visible
Dim lIconCount As Long                      'The total number of icons in the directory
Dim bStopLoading As Boolean                 'In case it reaches the limit number of icons

Dim lIconWidth As Long                      'The icon width (in twips)
Dim lIconHeight As Long                     'The icon height (in twips)

'Allows me to position a Line control in one line of code
Private Sub MoveLine(ln As Line, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    With ln
        .x1 = x1
        .y1 = y1
        .x2 = x2
        .y2 = y2
    End With
End Sub

'Toggle the frame's visibility (the box drawn by the 4 Line controls [lnFrame(0-3)])
Private Sub ShowFrame(ByVal bOn As Boolean)
    Dim i As Integer
    For i = 0 To 3
        lnFrame(i).Visible = bOn
    Next i
End Sub

'Draws the frame around picIcon(Index)
Private Sub DrawFrame(ByVal Index As Long)
    Dim i As Integer
    i = 18              'A value needed to subtract from the left and top of the frame
                        '   in order for the lines to show. I'm not sure why, I just played around
                        '   with it until it looked right
    With picIcon(Index)
        'Top
        MoveLine lnFrame(0), .Left - i, .Top - i, .Left + .Width + i, .Top - i
        'Left
        MoveLine lnFrame(1), .Left - i, .Top - i, .Left - i, .Top + .Height + i
        'Right
        MoveLine lnFrame(2), .Left + .Width, .Top - i, .Left + .Width, .Top + .Height + i
        'Bottom
        MoveLine lnFrame(3), .Left - i, .Top + .Height, .Left + .Width + i, .Top + .Height
    End With
End Sub

'Sets the progress bar.  If lMax = 0 then turn it off, if lMax <> -1 then turn it on
Private Sub SetProgress(ByVal lValue As Long, Optional ByVal lMax As Long = -1)

    If lMax > 1 Then
        progbar.Min = 0
        progbar.Max = lMax - 1
        progbar.Value = 0
        progbar.Visible = True
    ElseIf lMax = 0 Then
        progbar.Visible = False
    End If

    If lValue <= progbar.Max Then progbar.Value = lValue
    DoEvents                'Ensures that the progress bar is visibly updated
End Sub

'Sets the status message next to the progress bar
Private Sub SetStatus(ByVal sMsg As String)
    lblStatus.Caption = sMsg
    With progbar
        .Left = lblStatus.Left + lblStatus.Width + 50
        .Width = Me.ScaleWidth - .Left
    End With
End Sub

'Clear all the icons and selection box
Private Sub ClearIcons()
    Dim i As Integer
    For i = 0 To picIcon.Count - 1          'Clear all the picIcon controls
        picIcon(i).Cls                      '   (we don't want to show icons from
    Next i                                  '    previous files)
    fraIcons.Top = 0                        'And we won't be needing to scroll
    VS.Min = 0                              '   anything so clear that stuff up too
    VS.Max = 0
    VS.Value = 0
    
    lCount = 0                              'There are no icons visible
    lSelected = -1                          '   and so nothing is selected
    ShowFrame False                         '   and no frame is needed
End Sub

'Show all the icons in the selected file
Private Sub LoadIcons(ByVal Index As Integer, Optional ByVal bAutoShowHide As Boolean = True)
    On Error Resume Next    'Occasionally, an error occurs with the icon, so just
                            '   move on and get the next one
                            
    If bStopLoading = True Then Exit Sub
    If bAutoShowHide Then fraIcons.Visible = False
                            
    'hIcon is the handle to the extracted icon
    'i keeps track of the icon in the file
    'lIconIndex keeps track of what picIcon to use
    'lNewCount is how many total icons there are (lNewCount - lCount is number of new icons)
    'l is for looping through the icons
    Dim hIcon As Long, i As Integer, lIconIndex As Long, lNewCount As Long, l As Long
    Dim sFile As String
    
    sFile = Dir1.Path & "\" & File1.ListItems(Index).Text
    lNewCount = lCount + File1.ListItems(Index).SubItems(1) 'The number of icons
    
    If lNewCount > lCount Then              'If icons were in the file...
        lIconIndex = picIcon.Count          'Add more picIcon controls, if needed
        While lIconIndex < lNewCount
            Load picIcon(lIconIndex)
            lIconIndex = lIconIndex + 1
        Wend
        
        lIconIndex = picIcon.Count          'Remove any excess picIcon controls
        While lIconIndex > lNewCount
            Unload picIcon(lIconIndex - 1)
            lIconIndex = lIconIndex - 1
        Wend
        
        If lNewCount > picIcon.Count Then   'This will only be true if the limit to the number of
            GoTo LimitReached               '   picIcon controls that can be loaded has been reached
        End If                              '   so the loading would need to cease
        
        'Inform the user that icons are loading from the selected file
        SetStatus "Loading from " & File1.ListItems(Index).Text
        SetProgress 0, lNewCount - lCount
        
        i = 0
        lIconIndex = lCount
        For l = lCount To lNewCount - 1
            picIcon(lIconIndex).Cls                         'Clear the current icon
            hIcon = ExtractIcon(App.hInstance, sFile, i)    'Extract the new icon
            If hIcon Then                                   'If it's a valid icon then..
                DrawIcon picIcon(lIconIndex).hdc, 0, 0, hIcon   'Draw it
                DestroyIcon hIcon                               'Destroy the icon handle
                picIcon(lIconIndex).Visible = True              'Show it
                lIconIndex = lIconIndex + 1
            End If
            
            SetProgress i                                   'Update the progress
            i = i + 1
        Next l
        
        SetStatus ""                            'All done getting the icons from
        SetProgress 0, 0                        '   this file
        
        lSelected = 0                           'Whenever loading new icons, select the first
        DrawFrame lSelected                     'Show that the first icon is selected
        ShowFrame True                          '   and ensure that it's visible
        
        lCount = lIconIndex
        If bAutoShowHide Then fraIcons.Visible = True
    End If
    Exit Sub
        
LimitReached:
    lCount = picIcon.Count                      'Update the total number of icons
    bStopLoading = True
    If bAutoShowHide Then fraIcons.Visible = True
End Sub

'Positions all the picIcon controls by using the given amount of room width-wise and as
'   much height room as needed to fit them all (scrolled with the vertical scrollbar [VS])
Private Sub PositionIcons()
    On Error Resume Next
    
    Dim i As Integer, Y As Long, X As Long
    Dim iRow As Integer, iCol As Integer
    Dim iPad As Integer
    Dim yDist As Long, xDist As Long
    
    If lCount <= 0 Then GoTo Done
    
    fraIcons.Visible = False
    SetStatus "Positioning Icons"                       'Inform the user of what is
    SetProgress 0, lCount                               '   happening
    
    iPad = 200      'Amount of room to be used to seperate the icons
    iRow = 0        'Starting row
    iCol = 0        '   and column
    
    yDist = lIconHeight + iPad      'These values never change so a speed enhancement
    xDist = lIconWidth + iPad       '   is gained by putting them to 2 variables
                                    '   rather than looking up both variables each time
                                    '   their sum is needed
    
    X = iPad        'Starting position for first icon
    Y = iPad
    
    For i = 0 To picIcon.Count - 1
Calc:
        If X + lIconWidth + 50 > fraIcons.Width Then    'If the icon is going to show
            iCol = 0                                    '   outside designated area
            iRow = iRow + 1                             '   then reset to next row
            X = iPad
            Y = (yDist * iRow) + iPad
        End If
        
        picIcon(i).Move X, Y                            'Position the icon
        iCol = iCol + 1                                 'Update for next column
        X = (xDist * iCol) + iPad
        
        SetProgress i                                   'Update progress display
    Next i
    
    If Y > picIconHolder.Height Then                    'If the last icon is below
        VS.Min = 0                                      '   the area of visible icons
        VS.Value = 0                                    '   then you'll need to
        VS.Max = Y - picIconHolder.Height               '   setup the scrollbar (VS)
        VS.SmallChange = lIconHeight + 120              '   to be able to scroll
        VS.LargeChange = (lIconHeight + 120) * 5        '   the icons
        fraIcons.Height = Y
    Else
        VS.Min = 0                                      'Otherwise, you don't need
        VS.Max = 0                                      '   the scrollbar
        VS.Value = 0
    End If
    
Done:
    If picIcon.Count > 0 Then
        DrawFrame lSelected
    End If
    
    SetProgress 0, 0
    SetStatus ""
    fraIcons.Visible = True
End Sub

'Loads all the files in the selected directory that have icons
Private Sub LoadFiles()
    Dim sFile As String, sFileList() As String, sExt As String, sDir As String
    Dim lExCount As Long, lFileCount As Long, i As Integer
    Dim newItem As ListItem
    
    Dim fso As New FileSystemObject
    
    fraNoFiles.Visible = False                                      'Hide this message while
                                                                    '   the files are loaded
                                                                    
    ClearIcons                                                      'Clear any visible icons
    lSelected = -1                                                  'Nothing selected
    ShowFrame False
    
    i = 0
    sDir = Dir1.Path & IIf(Right$(Dir1.Path, 1) <> "\", "\", "")    'Get the directory
    sFile = Dir(sDir)                                               'Get the first file
    
    File1.ListItems.Clear                                           'Clear the list
    ReDim sFileList(0)
    
    SetStatus "Loading File List"                                   'Set the status
    SetProgress 0, fso.GetFolder(sDir).Files.Count                  '   and progress bar by
    lFileCount = 0                                                  '   using fso for
                                                                    '   number of files
                                                                    
    lIconCount = 0                                                  'No icons yet
    
    While Len(sFile)                                                'While there are files
        
        If InStr(sFile, ".") Then
            sExt = UCase$(Mid$(sFile, InStr(sFile, ".") + 1))       'Get the extension

            If sExt = "EXE" Or sExt = "DLL" Or sExt = "ICO" Or sExt = "CUR" Then

                lExCount = ExtractIcon(App.hInstance, sDir & sFile, -1) 'Get the number
                                                                        '   of icons
                                                                        
                lIconCount = lIconCount + lExCount                  'Update the total
                
                If lExCount > 0 Then
                    Set newItem = File1.ListItems.Add(, , sFile)    'Add filename and
                    newItem.SubItems(1) = lExCount                  '   icon count to the
                End If                                              '   list of files
                
            End If
            
        End If

        lFileCount = lFileCount + 1                                 'Update the file count
        SetProgress lFileCount                                      'Set the progress
        
        If sDir <> Dir1.Path & IIf(Right$(Dir1.Path, 1) <> "\", "\", "") Then
            Exit Sub    'Exit if directory changed during DoEvents of SetProgress
        End If
        
        sFile = Dir()                                               'Get the next file
    Wend
    
    SetStatus ""                                                    'Clear the status
    SetProgress 0, 0                                                '   and progrss bar
    
    File1.ColumnHeaders(1).Text = "Filename (" & File1.ListItems.Count & ")"
    'File1.ColumnHeaders(2).Text = "Icons (" & lIconCount & ")"
    File1.ColumnHeaders(2).Text = lIconCount
    
    'If no files with icons were found in the directory, then display the message
    fraNoFiles.Visible = IIf(File1.ListItems.Count = 0, True, False)
End Sub

'When the directory changes, change the file list to match
Private Sub Dir1_Change()
    LoadFiles
End Sub

Private Sub Dir1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 0 And Button = vbRightButton Then
        PopupMenu mnuTop(2)             'Popup the Directory menu
    End If
End Sub

'When the drive changes, change the directory list to match
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

'When the user clicks on a column header...
Private Sub File1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If File1.SortKey = ColumnHeader.Position - 1 Then   'If already selected, reverse order
        File1.SortOrder = IIf(File1.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        File1.SortOrder = lvwAscending              'Otherwise, ascend the order
        File1.SortKey = ColumnHeader.Position - 1   '   and set the new sort key
    End If
End Sub

'When a file is selected...
Private Sub File1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ClearIcons                                                  'Clear current icons
    LoadIcons Item.Index                                        'Show new icons
    PositionIcons                                               'Position the icons
    bStopLoading = False    'More can be loaded after a file is clicked
End Sub

'When the form becomes visible (after loading)...
Private Sub Form_Activate()
    Dim fso As New FileSystemObject
    If fso.FolderExists(DEF_FOLDER) Then            'If the Default Folder exists
        Drive1.Drive = Left$(DEF_FOLDER, 3)         '   then set the drive and
        Dir1.Path = DEF_FOLDER                      '   directory to it
    End If
End Sub

'While the form is loading...
Private Sub Form_Load()

    'Set some default locations for various controls.. these locations do not change
    '   and hence only need to be set once.  This cuts down on the amount of time
    '   needed during the Form_Resize() event
    With picIconHolder
        .Left = Drive1.Left + Drive1.Width + 120
        .Top = 120
    End With
        With fraIconHolder
            .Left = 0
            .Top = 0
        End With
            With fraIcons
                .Top = 120
            End With
            With VS
                .Top = 0
            End With
    
    lIconWidth = ScaleX(32, vbPixels, ScaleMode)    'Set the icon width and height from
    lIconHeight = ScaleY(32, vbPixels, ScaleMode)   '   32 in pixels to whatever scalemode
                                                    '   is being used
    
    picIcon(0).Width = lIconWidth                   'Set the original picIcon control to
    picIcon(0).Height = lIconHeight                 '   match (all subsequent picIcon
                                                    '   controls that are loaded will
                                                    '   match this control)
                                                    
    Me.Caption = App.Title & " v" & App.Major & "." & App.Minor
    
    SetStatus ""                'Clear the status and progress bar since nothing has
    SetProgress 0, 0            '   been done yet
    
    lSelected = -1              'No icon is selected
    ShowFrame False
End Sub

'Whenever the form is resized, move the controls to fill the space and position correctly
Private Sub Form_Resize()
    'Errors result when there isn't enough room specified for the control, so just move on
    '   to the next
    On Error Resume Next
    
    'If the form is minimized, then there's no need to do any of this
    If Me.WindowState = vbMinimized Then Exit Sub
    
    
    'Special Note: Although the picStatus control is set to be aligned to the
    '   bottom of the contrl, it doesn't actually move until after this event is executed.
    '   Because of this, I set the lHeight to .ScaleHeight - picStatus.Height and not
    '   just to picStatus.Top (its .Height never changes, so it works).
    '   Also, I force the picStatus's top to change before Positioning the Icons
    '   so that the progress bar and status message don't appear in the wrong place.
    '   (If you comment out the line that adjusts it's top, then run the app and
    '    maximize it, you'll see what I mean.)
    
    Dim CursPos As POINTAPI
    GetCursorPos CursPos
    
    'Set minimum values for the width and height of the form and restrict them to those
    '   to ensure that all content will always be visible and not srunched up too much
    '   Also, adjust the cursor's position so that those annoying horizontal and vertical
    '       lines don't show up as the cursor moves over the form while it's being reset
    '       to a larger size
    
    If Me.Width < 8535 And Me.Height < 6500 Then
        Me.Width = 8535
        Me.Height = 6500
        SetCursorPos ScaleX(Me.Left + Me.Width, Me.ScaleMode, vbPixels), _
                    ScaleY(Me.Top + Me.Height, Me.ScaleMode, vbPixels)
                    
    ElseIf Me.Width < 8535 Then
        Me.Width = 8535
        SetCursorPos ScaleX(Me.Left + Me.Width, Me.ScaleMode, vbPixels), CursPos.Y
    
    ElseIf Me.Height < 6500 Then
        Me.Height = 6500
        SetCursorPos CursPos.X, ScaleY(Me.Top + Me.Height, Me.ScaleMode, vbPixels)
    
    End If
    
    Dim lMid As Long, lWidth As Long, lHeight As Long
    lWidth = Me.ScaleWidth                          'It's smaller, faster, and easier to
    lHeight = Me.ScaleHeight - picStatus.Height     '   set these values to variables
    
    lMid = lHeight - Drive1.Top - Drive1.Height     'Find the mid point for the
    lMid = lMid / 2                                 '   directory and file lists
    
    Dir1.Height = lMid - 100
    File1.Top = lMid + 400
    File1.Height = lHeight - File1.Top - 120
    
    With fraNoFiles
        .Left = File1.Left + ((File1.Width - .Width) / 2)
        .Top = File1.Top + ((File1.Height - .Height) / 2)
    End With
    
    With picIconHolder
        .Width = lWidth - .Left - 120
        .Height = lHeight - .Top - 120
    End With
    
        With fraIconHolder
            .Width = picIconHolder.Width
            .Height = picIconHolder.Height
        End With
        
            With VS
                .Left = picIconHolder.ScaleWidth - .Width
                .Height = picIconHolder.ScaleHeight
            End With
        
            With fraIcons
                .Left = 50
                .Width = fraIconHolder.Width - VS.Width - 50
                .Height = fraIconHolder.Height - 300
            End With
    
    With progbar
        .Left = lblStatus.Left + lblStatus.Width + 120
        .Width = picStatus.ScaleWidth - .Left - 120
    End With
    
    With picStatus
        .Top = lHeight - .Height
    End With

    PositionIcons       'Since the dimensions of the form has changed, we'll need to
                        '   reposition the icons to fit within it
End Sub

'When the form unloads..
Private Sub Form_Unload(Cancel As Integer)
    Dim lIcon As Long
    If picIcon.Count > 1 Then
        For lIcon = picIcon.Count - 1 To 1      'Ensure to unload each excess picIcon
            Unload picIcon(lIcon)               '   control from memory
        Next lIcon
    End If
    End             'Ensure all processes terminate and the program ends
End Sub

Private Sub fraIcons_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 0 And Button = vbRightButton And lSelected > -1 Then
        PopupMenu mnuTop(1)             'Popup the Icon menu
    End If
End Sub

Private Sub mnuDirectories_Click(Index As Integer)
    Select Case Index
        Case 0          'Show all Icons in Directory
            Dim i As Integer, bGoAhead As Boolean
            
            'Make sure the users know what they're in for...
            If Int(File1.ColumnHeaders(2).Text) > 1000 Then
                If MsgBox("Are you sure you wish to load all icons in the selected" & _
                        " directory?" & vbCrLf & _
                        "This may be a very time-and-resource consuming process.", _
                        vbYesNo + vbCritical) = vbYes Then
                    bGoAhead = True
                End If
            Else
                bGoAhead = True
            End If
            
            If bGoAhead Then
                fraIcons.Visible = False
                
                lCount = 0                          'Set the count back to 0 so it removes
                                                    '   all previous icons first
                
                For i = 1 To File1.ListItems.Count  'Enumerate through the file list
                    LoadIcons i, False              'Show the icons from each file
                    DoEvents
                    If bStopLoading Then GoTo Done  'If the limit is reached, position
                Next i                              '   whatever icons were found and exit
                GoTo Done
            End If
            Exit Sub
            
Done:
            PositionIcons           'Position the icons
            bStopLoading = False    'More icons can be loaded afterwards
            
    End Select
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Select Case Index
        Case 0          'Quit
            End             'End the entire program
    End Select
End Sub

Private Sub mnuIcons_Click(Index As Integer)
    Select Case Index
    
        Case 0          'Copy to Clipboard
            If lSelected > -1 Then                          'Ensure that an icon is selected
                Clipboard.Clear                             'Clear the clipboard and set the
                Clipboard.SetData picIcon(lSelected).Image  '   selected icon into it
            End If
            
        Case 1          'Save to File
            On Error GoTo CancelError
    
            If lSelected > -1 Then                  'Ensure that an icon is selected
                With CD
                    If Len(.FileName) = 0 Then      'If there is no .Filename (the user
                        .FileName = App.Path        '   hasn't saved anything yet) then set
                    End If                          '   the default path to the app's path
                    
                    .Filter = "Bitmap Files|*.bmp"  'It can only be saved as a bitmap
                    .ShowSave                       'If the user cancels here, an error
                                                    '   occurs; otherwise, move on to save
                                                    '   the file
                                                            
                    On Error GoTo SavingError   'In case an error occurs saving the file
                    SavePicture picIcon(lSelected).Image, .FileName
                End With
            End If
            
            Exit Sub    'Exit the sub.. there were no errors and it's done
            
SavingError:
            MsgBox "An error occured saving your file."
            Exit Sub
            
CancelError:            'If the user hit cancel, then it's done
            
            
    End Select
End Sub

'Whenever a picIcon is clicked...
Private Sub picIcon_Click(Index As Integer)
    lSelected = Index           'It becomes the selected icon
    DrawFrame Index             '   and we draw it's frame (should already be visible
                                '   from when we loaded the icons)
End Sub

Private Sub picIcon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 0 And Button = vbRightButton And lSelected > -1 Then
        PopupMenu mnuTop(1)         'Popup the Icon menu
    End If
End Sub

'When the vertical scrollbar's value changes, adjust the position of the frame with
'   the icons to match (so we can scroll the icons within the picture box)
Private Sub vs_Change()
    fraIcons.Top = (-1 * VS.Value)
End Sub

'The scrollbar does an annoying flashing thing when it has focus...
Private Sub VS_GotFocus()
    picIconHolder.SetFocus  'so we send the focus to the picIcon area instead
                            '   I chose this object so that the mouse wheel still works
End Sub

'As the user scrolls the bar up and down, call the change event
Private Sub VS_Scroll()
    vs_Change
End Sub
