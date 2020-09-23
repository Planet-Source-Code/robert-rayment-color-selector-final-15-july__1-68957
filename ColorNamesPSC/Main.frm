VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   Color selector"
   ClientHeight    =   3690
   ClientLeft      =   150
   ClientTop       =   0
   ClientWidth     =   5445
   ForeColor       =   &H00000000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   246
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSysColors 
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   195
      Picture         =   "Main.frx":0ABA
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   " System colors "
      Top             =   495
      Width           =   375
   End
   Begin VB.HScrollBar scrDarkBright 
      Height          =   180
      LargeChange     =   4
      Left            =   630
      Max             =   255
      Min             =   -255
      TabIndex        =   36
      Top             =   1350
      Width           =   3825
   End
   Begin VB.OptionButton optFile 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Spectral "
      Height          =   285
      Index           =   1
      Left            =   4020
      TabIndex        =   34
      Top             =   3300
      Width           =   1110
   End
   Begin VB.OptionButton optFile 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alphabetical "
      Height          =   285
      Index           =   0
      Left            =   195
      TabIndex        =   33
      Top             =   3300
      Value           =   -1  'True
      Width           =   1230
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   825
      Top             =   2685
   End
   Begin VB.CheckBox chkPicker 
      DownPicture     =   "Main.frx":0C84
      Height          =   375
      Left            =   4935
      Picture         =   "Main.frx":12D6
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   165
      Width           =   390
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Selected color"
      Height          =   1530
      Left            =   1350
      TabIndex        =   19
      Top             =   1635
      Width           =   4035
      Begin VB.PictureBox picSEL 
         BackColor       =   &H00000000&
         Height          =   240
         Left            =   135
         ScaleHeight     =   180
         ScaleWidth      =   1725
         TabIndex        =   21
         Top             =   240
         Width           =   1785
      End
      Begin VB.CommandButton cmdToClipB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "To Clipboard"
         Height          =   360
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label LabWebCul 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3330
         TabIndex        =   42
         ToolTipText     =   " Closest Web Safe Color, Click to select "
         Top             =   855
         Width           =   570
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Web"
         Height          =   225
         Index           =   2
         Left            =   2130
         TabIndex        =   41
         Top             =   870
         Width           =   375
      End
      Begin VB.Label LabHexWeb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2565
         TabIndex        =   40
         ToolTipText     =   " Closest Web Safe Color "
         Top             =   855
         Width           =   690
      End
      Begin VB.Label LabHexLng 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   2565
         TabIndex        =   29
         Top             =   540
         Width           =   690
      End
      Begin VB.Label LabHexLng 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   2565
         TabIndex        =   28
         Top             =   1170
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hex"
         Height          =   225
         Index           =   0
         Left            =   2130
         TabIndex        =   27
         Top             =   555
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Long"
         Height          =   225
         Index           =   1
         Left            =   2130
         TabIndex        =   26
         Top             =   1185
         Width           =   405
      End
      Begin VB.Label LabSelName 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   45
         TabIndex        =   25
         Top             =   540
         Width           =   1935
      End
      Begin VB.Label LabSelRGB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   2145
         TabIndex        =   24
         Top             =   255
         Width           =   420
      End
      Begin VB.Label LabSelRGB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   2655
         TabIndex        =   23
         Top             =   255
         Width           =   420
      End
      Begin VB.Label LabSelRGB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   3150
         TabIndex        =   22
         Top             =   255
         Width           =   420
      End
   End
   Begin VB.PictureBox picC 
      BackColor       =   &H00E0E0E0&
      Height          =   4905
      Left            =   195
      ScaleHeight     =   323
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   342
      TabIndex        =   11
      Top             =   3765
      Width           =   5190
      Begin VB.VScrollBar scrPicName 
         Height          =   4830
         LargeChange     =   22
         Left            =   4680
         Max             =   1230
         Min             =   4
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   15
         Value           =   4
         Width           =   420
      End
      Begin VB.PictureBox picNames 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         Height          =   23250
         Left            =   -15
         ScaleHeight     =   1546
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   332
         TabIndex        =   12
         Top             =   -30
         Width           =   5040
         Begin VB.Label LabName 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LabName"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   375
            TabIndex        =   16
            Top             =   90
            Width           =   1965
         End
         Begin VB.Label LabCul 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LabCul"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   45
            TabIndex        =   15
            Top             =   90
            Width           =   300
         End
         Begin VB.Label LabName 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LabName"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2685
            TabIndex        =   14
            Top             =   90
            Width           =   1965
         End
         Begin VB.Label LabCul 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LabCul"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   2355
            TabIndex        =   13
            Top             =   90
            Width           =   300
         End
      End
   End
   Begin VB.PictureBox picDISP 
      BackColor       =   &H00000000&
      Height          =   1050
      Left            =   180
      ScaleHeight     =   990
      ScaleWidth      =   480
      TabIndex        =   7
      ToolTipText     =   " Click to select "
      Top             =   1740
      Width           =   540
   End
   Begin VB.HScrollBar scrRB 
      Height          =   225
      Index           =   0
      LargeChange     =   5
      Left            =   630
      Max             =   255
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   990
      Width           =   3825
   End
   Begin VB.HScrollBar scrRB 
      Height          =   225
      Index           =   2
      LargeChange     =   5
      Left            =   630
      Max             =   255
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   3825
   End
   Begin VB.PictureBox picGRAD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   615
      ScaleHeight     =   45
      ScaleMode       =   0  'User
      ScaleWidth      =   256
      TabIndex        =   0
      ToolTipText     =   " Click to select "
      Top             =   420
      Width           =   3840
   End
   Begin VB.VScrollBar scrGreen 
      Height          =   1050
      LargeChange     =   5
      Left            =   4500
      Max             =   0
      Min             =   255
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   165
      Width           =   240
   End
   Begin VB.Line Line1 
      X1              =   168
      X2              =   168
      Y1              =   84
      Y2              =   108
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Brighter"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   4485
      TabIndex        =   38
      ToolTipText     =   " Brighter "
      Top             =   1335
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Darker"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   37
      ToolTipText     =   " Darker "
      Top             =   1335
      Width           =   600
   End
   Begin VB.Label LabDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "12 Jul 2007"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   4365
      TabIndex        =   35
      Top             =   8700
      Width           =   900
   End
   Begin VB.Label LabS 
      BackColor       =   &H00E0E0E0&
      Caption         =   "'S'"
      Height          =   180
      Left            =   5040
      TabIndex        =   32
      Top             =   765
      Width           =   210
   End
   Begin VB.Label LabPick 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Picker "
      Height          =   210
      Left            =   4920
      TabIndex        =   31
      Top             =   540
      Width           =   480
   End
   Begin VB.Label LabToggler 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOGGLE COLOR NAMES"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1560
      TabIndex        =   18
      Top             =   3300
      Width           =   2280
   End
   Begin VB.Label LabG 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Green"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   4665
      TabIndex        =   6
      Top             =   990
      Width           =   495
   End
   Begin VB.Label LabRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   825
      TabIndex        =   10
      Top             =   2490
      Width           =   420
   End
   Begin VB.Label LabRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "G"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   825
      TabIndex        =   9
      Top             =   2160
      Width           =   420
   End
   Begin VB.Label LabRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   825
      TabIndex        =   8
      Top             =   1830
      Width           =   420
   End
   Begin VB.Label LabRB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Blue"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   195
      TabIndex        =   5
      Top             =   990
      Width           =   480
   End
   Begin VB.Label LabRB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Red"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   195
      TabIndex        =   4
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Color Selector  by  Robert Rayment  7th July 2007

' Color names from :-
' www.learningwebdesign.com/colornames.html

' Files needed in App Folder :-
' AlphaNames.txt
' SpectraNames.txt
' NB Fuschia & Magenta are the same in the refs.

' File made :-
' FormLoc.ini

' Update  15/7/07

'1. Added system color dialog
'2. Added closest web-safe color


'Update 13/7/07

'1. Better Darker/Brighter action
'2. Select also on picDISP (mainly for Dark/Bright)
'3. Added key 'P' to toggle picker

'Update 12/7/07

'1. Add choice of alphabetical or spectral ordering of colors.
'   Spectral colors sorted by eye.
'2. Widen color labels to allow 'medium spring green'.
'3. Made hex 6 characters.


'Update 10/7/07

'1. Added simpler screen color picker,
'   ie not including layered windows.
'   Added balloon tooltip to picker button.
'   Move cursor over screen to show color.
'   Key 'S' to select when app in focus.
'2. Include Minimize on right-click title bar.

' Update 9/7/07

'1.  Added gradient to white as well as black
'    Larger font for color names
'    Color name also saved to clipboard
'    Tooltip on picGRAD 'Click to select'.

Option Explicit

' Window on top
Private Declare Function SetWindowPos Lib "USER32" (ByVal hwnd As Long, _
   ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
   ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1

' For DISPLY
Private Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type
Private BMIH As BITMAPINFOHEADER

Private Declare Function SetDIBitsToDevice Lib "gdi32" _
   (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, _
   ByVal SrcX As Long, ByVal SrcY As Long, _
   ByVal Scan As Long, ByVal NumScans As Long, _
   Bits As Any, BitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long) As Long

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Const COLORONCOLOR = 3

' For picker
Private Type POINTAPI
  kx As Long
  ky As Long
End Type
Private Declare Function GetDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "USER32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" _
   (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetAsyncKeyState Lib "USER32" _
   (ByVal vKey As KeyCodeConstants) As Long
Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long
Private PrevCurPos As POINTAPI  ' Previous mouse screen position
Private CurPos As POINTAPI      ' Mouse screen position


Private CulARR() As Byte
Private prevIndex As Integer
Private ColorsText$
Private PathSpec$
Private IniSpec$
Private FormTop As Long
Private FormLeft As Long

' Original selected values - used by Brighter/Darker
Private RORG As Byte, GORG As Byte, BORG As Byte

Private aPick As Boolean
Private PickCul As Long
Dim cc As clsToolTip


Private Sub chkPicker_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   cc.ShowToolTip chkPicker.hwnd, "Keep app in focus" & vbCrLf & "Key 'S' to select color", "Pick color from screen"
   picGRAD.SetFocus
End Sub

Private Sub chkPicker_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   aPick = -chkPicker.Value
   If aPick Then
      Timer1.Enabled = True
      LabS.Visible = True
   Else
      Timer1.Enabled = False
      LabS.Visible = False
   End If
End Sub

Private Sub cmdSysColors_Click()
Dim CF As ColorDialog
Dim R As Byte, G As Byte, B As Byte
Dim Cul As Long
   Set CF = New ColorDialog
   If CF.VBChooseColor(Cul, , , , Me.hwnd) Then
      LngToRGB Cul, R, G, B
      scrRB(2).Value = R
      scrRB(0).Value = B
      scrGreen.Value = G
      picSEL.BackColor = Cul
      ShowColorInfo Cul
      ShowColorInfo2 Cul
      LabSelName = ""
   End If
   Set CF = Nothing
   picGRAD.SetFocus
End Sub

Private Sub cmdToClipB_Click()
Dim R As Byte, G As Byte, B As Byte
Dim Cul As Long
Dim a$
Dim h$
   Cul = picSEL.BackColor
   LngToRGB Cul, R, G, B
   If LabSelName.Caption <> "" Then
      a$ = "' Color Name =  " & LabSelName.Caption & vbCrLf
   Else
      a$ = ""
   End If
   h$ = Hex$(Cul)
   If Len(h$) < 6 Then
      h$ = String$(6 - Len(h$), "0") & h$
   End If
   a$ = a$ & "' RGB(" & Str$(R) & ", " & Str$(G) & ", " & Str$(B) & ")" & vbCrLf
   a$ = a$ & "' &H" & h$ & vbCrLf
   
   a$ = a$ & "' Nearest WebSafe = " & LabHexWeb.Caption & vbCrLf
   
   
   a$ = a$ & "' Long =  " & Str$(Cul)
   Clipboard.Clear
   Clipboard.SetText a$

   picDISP.SetFocus
End Sub

Private Sub Form_Initialize()
   ' Window on top
   Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub

Private Sub Form_Load()
Dim Ret$
Dim j As Long, k As Long

' As on Clipboard:

' Color Name =  gold
' RGB( 255,  215,  0)
' &H00D7FF
' Nearest WebSafe = FFCC00
' Long =   55295


' eg
' Form1.BackColor = 16777184
   
   Set cc = New clsToolTip

   On Error GoTo FileError

   LabDate = "15/Jul/2007"
   LabS.Visible = False
   
   ' For picker
   aPick = False
   Timer1.Enabled = False
   KeyPreview = True
   PrevCurPos.kx = 0
   PrevCurPos.ky = 0
   
   FormTop = 600
   FormLeft = 600
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
      
   IniSpec$ = PathSpec$ & "FormLoc.ini"
   If FileExists(IniSpec$) Then
      If GetINI("FormLoc", "FormTop", Ret$, IniSpec$) Then
         FormTop = Val(Ret$)
         GetINI "FormLoc", "FormLeft", Ret$, IniSpec$
         FormLeft = Val(Ret$)
      End If
   End If
   
   'Form1.Left = FormLeft
   'Form1.Top = FormTop
   'Form1.Height = 3780
   
   Form1.Move FormLeft, FormTop, 5535, 4080
   
   ColorsText$ = PathSpec$ & "SpectraNames.txt"
   If Not FileExists(ColorsText$) Then
      MsgBox "Color names file, SpectraNames.txt, missing !  ", vbCritical, "Color Selector"
      Unload Me
      Exit Sub
   End If
   
   ColorsText$ = PathSpec$ & "AlphaNames.txt"
   If Not FileExists(ColorsText$) Then
      MsgBox "Color names file, AlphaNames.txt, missing !  ", vbCritical, "Color Selector"
      Unload Me
      Exit Sub
   End If
   
   
   
   ' Size gradient display (short for Large Fonts but useable)
   ' (Could locate scrollbars to picColrs)
   picGRAD.Width = 256
   picGRAD.Height = 36
   
   ' Size to hold labels for 140 colors on 70 rows
   picNames.Height = 1554   ' 140 \ 2 * 22 + 14  (22 = label vertical separation)
   scrPicName.Min = 4
   scrPicName.Max = 1227 + 4 ' 1554 - 327 (ie picC.height) + 4
   
   LabSelName.Height = 20 * Screen.TwipsPerPixelY
   
   LoadLabels
   ' To switch label border styles
   prevIndex = 0
   
   ' For gradient display
   ReDim CulARR(0 To 2, 0 To 255, 0 To 35)
   LabRB(0) = 0   ' Blue
   LabRB(2) = 0   ' Red
   LabG = 0       ' Green
   ShowColorInfo 0
   ShowColorInfo2 0
   
   
   ' Fill labels from AlphaNames.txt in app folder
   ColorsText$ = PathSpec$ & "AlphaNames.txt"
   FillColors ColorsText$
   
   LabSelName = "Black"
   
   scrPicName.Value = scrPicName.Min
   
   RORG = 0
   GORG = 0
   BORG = 0
   
   Show
   
   scrDarkBright.Value = 1
   scrDarkBright.Value = 0
   
   On Error GoTo 0
   Exit Sub
'=========
FileError:
   MsgBox "Color names file, error !  ", vbCritical, "Color Selector"
   Unload Me
End Sub

Private Sub FillColors(ColorFileName$)
Dim Fnum As Long
Dim a$
Dim R As Byte, G As Byte, B As Byte
Dim k As Long
   Fnum = FreeFile
   Open ColorsText$ For Input As #Fnum
   Do Until EOF(Fnum)
      Input #Fnum, a$, R, G, B
      LabName(k) = a$
      LabName(k).BorderStyle = 0
      LabCul(k).BackColor = RGB(R, G, B)
      k = k + 1
      If k > 140 Then Exit Do
   Loop
   Close #Fnum
End Sub
   

Private Sub Form_KeyPress(KeyAscii As Integer)
' For selecting color from picker
Dim R As Byte, G As Byte, B As Byte
   
   ' Key p or P to toggle picker
   ' KeyAscii 112 = p, 80 = P
   If KeyAscii = 112 Or KeyAscii = 80 Then
      aPick = Not aPick
      If aPick Then
         chkPicker.Value = Checked
         Timer1.Enabled = True
         LabS.Visible = True
      Else
         chkPicker.Value = Unchecked
         Timer1.Enabled = False
         LabS.Visible = False
      End If
   End If
   
   If Not aPick Then Exit Sub
   
   ' Key s or S to select color
   ' KeyAscii 115 = s, 83 = S
   If KeyAscii = 115 Or KeyAscii = 83 Then
      If PickCul < 0 Then PickCul = 0
      LngToRGB PickCul, R, G, B
      scrRB(2).Value = R
      scrRB(0).Value = B
      scrGreen.Value = G
      LabSelName = ""
      picSEL.BackColor = PickCul
      ShowColorInfo PickCul
      ShowColorInfo2 PickCul
   End If
End Sub

Private Sub LoadLabels()
' (22 = label vertical separation)
Dim k As Long
   LabCul(0) = ""
   LabCul(1) = ""
   LabName(0).Height = 20
   LabName(1).Height = 20
   LabName(0).Width = 131 '118
   LabName(1).Width = 131 '118
   LabName(0).BorderStyle = 0
   LabName(1).BorderStyle = 0
   For k = 2 To 138 Step 2
      Load LabName(k)
      LabName(k).Left = LabName(0).Left
      LabName(k).Width = LabName(0).Width
      LabName(k).Top = LabName(0).Top + 22 * (k \ 2)
      LabName(k) = Str$(k)
      LabName(k).BorderStyle = 0
      LabName(k).Visible = True
      Load LabCul(k)
      LabCul(k).Left = LabCul(0).Left
      LabCul(k).Top = LabCul(0).Top + 22 * (k \ 2)
      LabCul(k) = ""
      LabCul(k).Visible = True
   Next k
   For k = 3 To 139 Step 2
      Load LabName(k)
      LabName(k).Left = LabName(1).Left
      LabName(k).Width = LabName(1).Width
      LabName(k).Top = LabName(1).Top + 22 * ((k - 1) \ 2)
      LabName(k) = Str$(k)
      LabName(k).BorderStyle = 0
      LabName(k).Visible = True
      Load LabCul(k)
      LabCul(k).Left = LabCul(1).Left
      LabCul(k).Top = LabCul(1).Top + 22 * ((k - 1) \ 2)
      LabCul(k) = ""
      LabCul(k).Visible = True
   Next k

   Line1.X1 = scrDarkBright.Left + 0.5 * (scrDarkBright.Width)
   Line1.X2 = Line1.X1

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Timer1.Enabled = False
   
   FormTop = Form1.Top
   FormLeft = Form1.Left

   If FileExists(IniSpec$) Then Kill IniSpec$

   WriteINI "FormLoc", "FormTop", Trim$(Str$(FormTop)), IniSpec$
   WriteINI "FormLoc", "FormLeft", Trim$(Str$(FormLeft)), IniSpec$

   Set cc = Nothing
   Set Form1 = Nothing
End Sub

Private Sub LabName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' Select from named colors
   Call LabCul_MouseUp(Index, 1, 0, 0, 0)
End Sub

Private Sub LabCul_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' Select from named colors
Dim R As Byte, G As Byte, B As Byte
Dim Cul As Long
   Cul = LabCul(Index).BackColor
   LngToRGB Cul, R, G, B
   scrRB(2).Value = R
   scrRB(0).Value = B
   scrGreen.Value = G
   LabName(prevIndex).BorderStyle = 0
   LabName(Index).BorderStyle = 1
   prevIndex = Index
   LabSelName = LabName(Index)

   picSEL.BackColor = Cul
   ShowColorInfo Cul
   ShowColorInfo2 Cul
   
   scrDarkBright.Value = 0
   LabSelName = LabName(Index)
End Sub


Private Sub LabToggler_Click()
   If Form1.Height = 4080 Then
      Form1.Height = 9345
   Else
      Form1.Height = 4080
   End If
End Sub

Private Sub LabWebCul_Click()
' Select from LabWebCul display
Dim R As Byte, G As Byte, B As Byte
Dim Cul As Long
   Cul = LabWebCul.BackColor
   If Cul < 0 Then Cul = 0
   picSEL.BackColor = Cul
   ShowColorInfo Cul
   ShowColorInfo2 Cul
   LabSelName = ""
End Sub

Private Sub optFile_Click(Index As Integer)
   If Index = 0 Then
      ColorsText$ = PathSpec$ & "AlphaNames.txt"
   Else
      ColorsText$ = PathSpec$ & "SpectraNames.txt"
   End If
   FillColors ColorsText$
End Sub

Private Sub picGRAD_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Select from picGRAD display
Dim R As Byte, G As Byte, B As Byte
Dim Cul As Long
   Cul = picGRAD.Point(x, y)
   If Cul < 0 Then Cul = 0
   picSEL.BackColor = Cul
   ShowColorInfo Cul
   ShowColorInfo2 Cul
   LabSelName = ""
End Sub

Private Sub picGRAD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim R As Byte, G As Byte, B As Byte
Dim Cul As Long
   Cul = picGRAD.Point(x, y)
   If Cul < 0 Then Cul = 0
   picDISP.BackColor = Cul
   ShowColorInfo Cul
End Sub

Private Sub picDISP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' Select from picDISP display
Dim R As Byte, G As Byte, B As Byte
Dim Cul As Long
   Cul = picDISP.Point(x, y)
   If Cul < 0 Then Cul = 0
   picSEL.BackColor = Cul
   ShowColorInfo Cul
   ShowColorInfo2 Cul
   LabSelName = ""
End Sub

Private Sub scrDarkBright_Change()
   Call scrDarkBright_Scroll
End Sub

Private Sub scrDarkBright_Scroll()
Dim Cul As Long
Dim LR As Long, LG As Long, LB As Long
Dim addto As Long
   
   addto = scrDarkBright.Value
   LR = RORG + addto
   LG = GORG + addto
   LB = BORG + addto
   If LR < 0 Then LR = 0
   If LG < 0 Then LG = 0
   If LB < 0 Then LB = 0
   If LR > 255 Then LR = 255
   If LG > 255 Then LG = 255
   If LB > 255 Then LB = 255
   scrRB(2).Value = LR
   scrGreen.Value = LG
   scrRB(0).Value = LB
   Cul = RGB(scrRB(2).Value, scrGreen.Value, scrRB(0).Value)
   ShowColorInfo Cul
   picGRAD.SetFocus
End Sub

Private Sub scrPicName_Scroll()
   Call scrPicName_Change
End Sub

Private Sub scrPicName_Change()
   picNames.Top = -scrPicName.Value
End Sub

Private Sub scrRB_Change(Index As Integer)
   Call scrRB_Scroll(Index)
End Sub
Private Sub scrRB_Scroll(Index As Integer)
' Index = 0 BLUE, Index = 2 RED
Dim Cul As Long
Dim R As Byte, G As Byte, B As Byte
Dim culstep As Single
Dim k As Long, j As Long
   LabRB(Index) = scrRB(Index).Value
   LabRB(Index).Refresh

   culstep = scrRB(Index).Value / 255
   For j = 0 To 18
   For k = 0 To 255
      CulARR(Index, k, j) = (k * culstep)
   Next k
   Next j
   
   culstep = (1 - scrRB(Index).Value / 255)
   For j = 18 To 35
   For k = 0 To 255
      Cul = scrRB(Index).Value + (k * culstep)
      If Cul > 255 Then Cul = 255
      CulARR(Index, 255 - k, j) = Cul
   Next k
   Next j
   
   DISPLAY
   Cul = RGB(scrRB(2).Value, scrGreen.Value, scrRB(0).Value)
   ShowColorInfo Cul
End Sub

Private Sub scrGreen_Change()
   Call scrGreen_Scroll
End Sub
Private Sub scrGreen_Scroll()
Dim Cul As Long
Dim culstep As Single
Dim R As Byte, G As Byte, B As Byte
Dim k As Long, j As Long
   LabG = scrGreen.Value
   LabG.Refresh

   culstep = scrGreen.Value / 255
   For j = 0 To 18
   For k = 0 To 255
      CulARR(1, k, j) = (k * culstep)
   Next k
   Next j
   
   culstep = (1 - scrGreen.Value / 255)
   For j = 18 To 35
   For k = 0 To 255
      Cul = scrGreen.Value + (k * culstep)
      If Cul > 255 Then Cul = 255
      CulARR(1, 255 - k, j) = Cul
   Next k
   Next j
   
   DISPLAY
   Cul = RGB(scrRB(2).Value, scrGreen.Value, scrRB(0).Value)
   ShowColorInfo Cul
End Sub

Private Sub ShowColorInfo(ByVal Col As Long)
Dim k As Long
Dim R As Byte, G As Byte, B As Byte
   LngToRGB Col, R, G, B
   LabRGB(0) = R
   LabRGB(1) = G
   LabRGB(2) = B
   picDISP.BackColor = Col
End Sub

Private Sub ShowColorInfo2(ByVal Col As Long)
Dim R As Byte, G As Byte, B As Byte
Dim h$
   LngToRGB Col, R, G, B
   LabSelRGB(0) = R
   LabSelRGB(1) = G
   LabSelRGB(2) = B
   RORG = R
   GORG = G
   BORG = B
   h$ = Hex$(Col)
   If Len(h$) < 6 Then
      h$ = String$(6 - Len(h$), "0") & h$
   End If
   LabHexLng(0) = h$
   LabHexLng(1) = Col
   Select Case B
   Case Is <= 25: B = 0
   Case Is <= 75: B = 51
   Case Is <= 125: B = 102
   Case Is <= 175: B = 153
   Case Is <= 225: B = 204
   Case Else: B = 255
   End Select
   Select Case G
   Case Is <= 25: G = 0
   Case Is <= 75: G = 51
   Case Is <= 125: G = 102
   Case Is <= 175: G = 153
   Case Is <= 225: G = 204
   Case Else: G = 255
   End Select
   Select Case R
   Case Is <= 25: R = 0
   Case Is <= 75: R = 51
   Case Is <= 125: R = 102
   Case Is <= 175: R = 153
   Case Is <= 225: R = 204
   Case Else: R = 255
   End Select
   LabWebCul.BackColor = RGB(R, G, B)
   Col = RGB(B, G, R) And &HFFFFFFFF
   h$ = Hex$(Col)
   If Len(h$) < 6 Then
      h$ = String$(6 - Len(h$), "0") & h$
   End If
   LabHexWeb = h$
   scrDarkBright.Value = 0
End Sub

Private Sub DISPLAY()
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biheight = 36
      .biwidth = 256
      .biBitCount = 24
   End With
   
   picGRAD.Picture = LoadPicture
   SetStretchBltMode picGRAD.hdc, COLORONCOLOR
   
   If SetDIBitsToDevice(picGRAD.hdc, 0, 0, picGRAD.Width, picGRAD.Height, _
      0, 0, 0, picGRAD.Height, CulARR(0, 0, 0), BMIH, 0) = 0 Then
      MsgBox " DISPLAY ERROR", vbCritical, "Color Selector"
      Unload Me
      End
   End If
   picGRAD.Picture = picGRAD.Image
End Sub

Private Sub LngToRGB(LCul As Long, R As Byte, G As Byte, B As Byte)
   R = LCul And &HFF&
   G = (LCul And &HFF00&) \ &H100&
   B = (LCul And &HFF0000) \ &H10000
End Sub

Private Function FileExists(FSpec$) As Boolean
On Error Resume Next   ' Needed if CD, Zip etc. disk removed
   'If Dir(FSpec$) <> "" Then FileExists = True
   FileExists = LenB(Dir$(FSpec$))
End Function

Private Sub Timer1_Timer()
' Simple color picker
' Private PickCul As Long
Dim rdc As Long
   If Not aPick Then Exit Sub
   GetCursorPos CurPos
   If CurPos.kx <> PrevCurPos.kx Or _
      CurPos.ky <> PrevCurPos.ky Then
         PrevCurPos.kx = CurPos.kx
         PrevCurPos.ky = CurPos.ky
         rdc = GetDC(0&)   ' Get Device Context to whole screen
         PickCul = GetPixel(rdc, CurPos.kx, CurPos.ky)
         If PickCul < 0 Then PickCul = 0
         picDISP.BackColor = PickCul
         ShowColorInfo PickCul
         ReleaseDC 0&, rdc
   End If
End Sub
